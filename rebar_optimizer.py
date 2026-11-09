"""Bounded cutting-stock optimization for reinforcing steel bars.

Small instances use full pattern enumeration and CBC. Large instances use a
restricted set of useful patterns with a solver time limit, followed by a
feasible greedy cutting plan when necessary. This prevents a long-running
optimization from blocking Excel export indefinitely.
"""

from __future__ import annotations

import os
import sys
from typing import Any, Callable

import pulp

# Keeping the integer-programming model small is essential. Enumeration grows
# exponentially in the number of different cut lengths.
MAX_EXACT_PATTERNS = 2400
MAX_EXACT_PIECE_TYPES = 12
MAX_SOLVER_PIECE_TYPES = 120
MAX_RESTRICTED_PATTERNS = 2400
DEFAULT_SOLVER_TIME_LIMIT = 8  # seconds per diameter


def mm(x_m: float) -> int:
    return int(x_m * 1000 + 0.5)


def m(x_mm: int) -> float:
    return x_mm / 1000.0


def get_solver_path():
    """Locate CBC both in source installs and packaged Windows builds."""
    if getattr(sys, 'frozen', False):
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, 'cbc.exe')
    import pulp.apis
    return pulp.apis.PULP_CBC_CMD().path


class PatternLimitReached(Exception):
    """Stop exhaustive enumeration before its runtime/memory explodes."""


def enumerate_patterns(stock_len_mm: int, piece_lengths_mm: list[int],
                       max_counts: list[int], max_patterns: int = MAX_EXACT_PATTERNS) -> list:
    """Enumerate feasible nonempty patterns, raising at the hard pattern cap."""
    if any(length <= 0 for length in piece_lengths_mm):
        raise ValueError('Cut lengths must be positive.')
    ub = [min(max_counts[i], stock_len_mm // length)
          for i, length in enumerate(piece_lengths_mm)]
    patterns: list[tuple[tuple[int, ...], int]] = []
    counts = [0] * len(piece_lengths_mm)

    def rec(i: int, used: int, occupied: bool) -> None:
        if i == len(counts):
            if occupied:
                if len(patterns) >= max_patterns:
                    raise PatternLimitReached()
                patterns.append((tuple(counts), used))
            return
        length = piece_lengths_mm[i]
        for count in range(min(ub[i], (stock_len_mm - used) // length) + 1):
            counts[i] = count
            rec(i + 1, used + count * length, occupied or count > 0)
        counts[i] = 0

    rec(0, 0, False)
    return patterns


def _greedy_pattern(stock_mm: int, piece_mm: list[int], quantities: list[int],
                    kerf_mm: int, order: list[int], seed: int | None = None):
    """Fill a stock bar with integer cut counts, respecting per-cut demand."""
    counts = [0] * len(piece_mm)
    used = 0
    number_of_pieces = 0
    if seed is not None:
        if quantities[seed] == 0 or piece_mm[seed] > stock_mm:
            return None
        counts[seed] = 1
        number_of_pieces = 1
        used = piece_mm[seed]

    for i in order:
        available = quantities[i] - counts[i]
        if available <= 0:
            continue
        piece = piece_mm[i]
        remaining = stock_mm - used
        if number_of_pieces:
            add = min(available, remaining // (piece + kerf_mm))
            used += add * (piece + kerf_mm)
        elif remaining >= piece:
            add = min(available, 1 + (remaining - piece) // (piece + kerf_mm))
            used += add * piece + max(0, add - 1) * kerf_mm
        else:
            add = 0
        counts[i] += add
        number_of_pieces += add
    if number_of_pieces == 0:
        return None
    return tuple(counts), used


def _orders(piece_mm: list[int], quantities: list[int]) -> list[list[int]]:
    indices = list(range(len(piece_mm)))
    return [
        sorted(indices, key=lambda i: (-piece_mm[i], i)),
        sorted(indices, key=lambda i: (piece_mm[i], i)),
        sorted(indices, key=lambda i: (-quantities[i], -piece_mm[i], i)),
        sorted(indices, key=lambda i: (-quantities[i] * piece_mm[i], -piece_mm[i], i)),
    ]


def _restricted_patterns(stock_mm: list[int], piece_mm: list[int],
                         quantities: list[int], kerf_mm: int) -> dict:
    """Build a bounded and deterministic pattern pool without enumerating all combinations."""
    per_stock_limit = max(1, MAX_RESTRICTED_PATTERNS // len(stock_mm))
    orders = _orders(piece_mm, quantities)
    result = {}
    for stock in stock_mm:
        patterns = {}
        # Pure patterns guarantee that each individually feasible cut has a way
        # to be produced; the other candidates create efficient mixed patterns.
        for i, length in enumerate(piece_mm):
            if length > stock or quantities[i] <= 0:
                continue
            pattern = _greedy_pattern(stock, piece_mm, quantities, kerf_mm, [i], i)
            if pattern:
                patterns[pattern[0]] = pattern[1]
        for order in orders:
            if len(patterns) >= per_stock_limit:
                break
            pattern = _greedy_pattern(stock, piece_mm, quantities, kerf_mm, order)
            if pattern:
                patterns[pattern[0]] = pattern[1]
            for seed in order:
                if len(patterns) >= per_stock_limit:
                    break
                pattern = _greedy_pattern(stock, piece_mm, quantities, kerf_mm, order, seed)
                if pattern:
                    patterns[pattern[0]] = pattern[1]
        if patterns:
            result[stock] = list(patterns.items())
    return result


def build_patterns_all_stocks(piece_lengths: list[float], piece_qty: list[int],
                              stock_lengths: list[float], kerf: float = 0.0,
                              return_exact: bool = False):
    """Create cutting patterns with a strict cap and scalable fallback."""
    piece_mm = [mm(x) for x in piece_lengths]
    stock_mm = sorted(set(mm(x) for x in stock_lengths))
    kerf_mm = mm(kerf)
    if not piece_mm or not stock_mm:
        return ({}, True) if return_exact else {}
    if kerf_mm < 0 or any(length <= 0 for length in piece_mm):
        raise ValueError('Cut lengths must be positive and kerf cannot be negative.')
    if any(q <= 0 or int(q) != q for q in piece_qty):
        raise ValueError('Rebar quantities must be positive whole numbers.')

    exact = len(piece_mm) <= MAX_EXACT_PIECE_TYPES
    patterns_by_stock: dict[int, list] = {}
    total_patterns = 0
    if exact:
        try:
            for stock in stock_mm:
                patterns = enumerate_patterns(stock, piece_mm, piece_qty,
                                              MAX_EXACT_PATTERNS - total_patterns)
                patterns_with_kerf = []
                for counts, used in patterns:
                    adjusted = used + max(0, sum(counts) - 1) * kerf_mm
                    if adjusted <= stock:
                        patterns_with_kerf.append((counts, adjusted))
                if patterns_with_kerf:
                    patterns_by_stock[stock] = patterns_with_kerf
                total_patterns += len(patterns)  # count even the kerf-infeasible ones
        except PatternLimitReached:
            exact = False
            patterns_by_stock = {}
    if not exact:
        patterns_by_stock = _restricted_patterns(stock_mm, piece_mm, piece_qty, kerf_mm)
    return (patterns_by_stock, exact) if return_exact else patterns_by_stock


def _pack_solution(patterns: list[dict], quantities: list[int],
                   chosen: list[int], status: str) -> dict:
    purchases = []
    produced = [0] * len(quantities)
    purchased_mm = used_mm = 0
    for pat, count in zip(patterns, chosen):
        if count <= 0:
            continue
        stock = pat['stock_mm']
        used = pat['used_mm']
        purchases.append({
            'stock_length_m': m(stock), 'pattern_counts': pat['counts'],
            'used_length_m': m(used), 'waste_m': m(stock - used),
            'quantity': count,
        })
        purchased_mm += stock * count
        used_mm += used * count
        for i, pattern_qty in enumerate(pat['counts']):
            produced[i] += pattern_qty * count
    if any(produced[i] < demand for i, demand in enumerate(quantities)):
        raise RuntimeError('Cutting plan did not cover all requested quantities.')
    return {
        'status': status, 'total_purchased_m': m(purchased_mm),
        'total_used_m': m(used_mm), 'total_waste_m': m(purchased_mm - used_mm),
        'purchases': purchases, 'produced_counts': produced,
        'demand_counts': quantities,
    }


def _greedy_solution(piece_mm: list[int], piece_qty: list[int],
                     stock_mm: list[int], kerf_mm: int) -> dict:
    """Always-terminating feasible fallback if CBC fails or hits its time limit."""
    remaining = list(piece_qty)
    orders = _orders(piece_mm, remaining)
    counts_by_pattern: dict[tuple[int, tuple[int, ...]], int] = {}
    max_bars = sum(remaining)
    iterations = 0
    while any(remaining):
        # Prefer leftover-free cuts, followed by stock utilization. Restrict
        # seeds on huge instances so candidate evaluation remains bounded.
        seeds = sorted((i for i, q in enumerate(remaining) if q),
                       key=lambda i: (-piece_mm[i], i))[:24]
        best = None
        for stock in stock_mm:
            for order in orders:
                for seed in [None] + seeds:
                    item = _greedy_pattern(stock, piece_mm, remaining, kerf_mm, order, seed)
                    if not item:
                        continue
                    counts, used = item
                    waste = stock - used
                    key = (waste, waste / stock, -used, stock)
                    if best is None or key < best[0]:
                        best = (key, stock, counts, used)
        if best is None:
            raise ValueError('At least one cut is longer than every selected market length.')
        _, stock, counts, _ = best
        # Reuse identical patterns in bulk instead of iterating per stock bar.
        multiple = min(remaining[i] // count for i, count in enumerate(counts) if count)
        if multiple == 0:
            multiple = 1
        key = (stock, counts)
        counts_by_pattern[key] = counts_by_pattern.get(key, 0) + multiple
        for i, count in enumerate(counts):
            remaining[i] = max(0, remaining[i] - count * multiple)
        iterations += 1
        if iterations > max_bars:
            raise RuntimeError('Failed to construct a complete cutting plan.')
    patterns = []
    chosen = []
    for (stock, counts), count in counts_by_pattern.items():
        used = sum(p * q for p, q in zip(piece_mm, counts))
        used += max(0, sum(counts) - 1) * kerf_mm
        patterns.append({'stock_mm': stock, 'counts': counts, 'used_mm': used})
        chosen.append(count)
    return _pack_solution(patterns, piece_qty, chosen, 'Feasible')


def solve_with_pulp(piece_lengths: list[float], piece_qty: list[int],
                    stock_lengths: list[float], kerf: float = 0.0,
                    verbose: bool = False) -> dict[str, Any]:
    """Solve exactly for small jobs; use bounded MILP/feasible fallback for large jobs."""
    if not piece_lengths:
        return {'status': 'NoPatterns', 'message': 'No cut lengths were supplied.'}
    piece_mm = [mm(x) for x in piece_lengths]
    stock_mm = sorted(set(mm(x) for x in stock_lengths))
    kerf_mm = mm(kerf)
    if any(x <= 0 for x in piece_mm) or kerf_mm < 0:
        raise ValueError('Cut lengths must be positive and kerf cannot be negative.')
    if len(piece_qty) != len(piece_mm) or any(q <= 0 or int(q) != q for q in piece_qty):
        raise ValueError('Provide a positive whole-number quantity for every cut length.')
    if not stock_mm or any(stock <= 0 for stock in stock_mm):
        return {'status': 'NoPatterns', 'message': 'No market lengths were selected.'}
    if any(length > max(stock_mm) for length in piece_mm):
        return {'status': 'NoPatterns', 'message': 'A required cut exceeds every selected market length.'}

    if len(piece_mm) > MAX_SOLVER_PIECE_TYPES:
        return _greedy_solution(piece_mm, piece_qty, stock_mm, kerf_mm)

    patterns_by_stock, exact = build_patterns_all_stocks(
        piece_lengths, piece_qty, stock_lengths, kerf, return_exact=True)
    patterns = [
        {'stock_mm': stock, 'counts': counts, 'used_mm': used}
        for stock, choices in patterns_by_stock.items() for counts, used in choices
    ]
    if not patterns:
        return {'status': 'NoPatterns', 'message': 'No feasible cutting patterns for the selected lengths.'}

    prob = pulp.LpProblem('rebar_cutting_stock', pulp.LpMinimize)
    variables = pulp.LpVariable.dicts('pattern', range(len(patterns)), lowBound=0, cat='Integer')
    for i in range(len(piece_mm)):
        prob += (pulp.lpSum(variables[j] * pattern['counts'][i]
                            for j, pattern in enumerate(patterns)) >= piece_qty[i], f'demand_{i}')
    purchased = pulp.lpSum(variables[j] * p['stock_mm'] for j, p in enumerate(patterns))
    waste = pulp.lpSum(variables[j] * (p['stock_mm'] - p['used_mm'])
                       for j, p in enumerate(patterns))
    prob.setObjective((max(stock_mm) + 1) * purchased + waste)

    try:
        seconds = max(1, min(60, int(os.environ.get('REBAR_SOLVER_TIME_LIMIT', DEFAULT_SOLVER_TIME_LIMIT))))
    except (ValueError, TypeError):
        seconds = DEFAULT_SOLVER_TIME_LIMIT
    try:
        solver = pulp.COIN_CMD(path=get_solver_path(), msg=verbose, timeLimit=seconds)
        prob.solve(solver)
        values = [pulp.value(variables[j]) for j in range(len(patterns))]
        if all(value is not None and value >= -1e-5 and abs(value - round(value)) < 1e-4
               for value in values):
            chosen = [max(0, int(round(value))) for value in values]
            if all(sum(chosen[j] * pattern['counts'][i]
                       for j, pattern in enumerate(patterns)) >= piece_qty[i]
                   for i in range(len(piece_mm))):
                # PuLP can mark a CBC time-limited incumbent as LpStatus
                # 'Optimal'; sol_status distinguishes a proved optimum.
                optimal = (exact and pulp.LpStatus[prob.status] == 'Optimal'
                           and getattr(prob, 'sol_status', None) == pulp.LpSolutionOptimal)
                return _pack_solution(patterns, piece_qty, chosen,
                                      'Optimal' if optimal else 'Feasible')
    except (pulp.PulpSolverError, OSError) as exc:
        if verbose:
            print(f'CBC unavailable; switching to feasible plan: {exc}')
    # CBC timed out without an integer incumbent, or failed to execute.
    return _greedy_solution(piece_mm, piece_qty, stock_mm, kerf_mm)


def find_optimized_cutting_plan(demands: dict[str, list[tuple]],
                                stocks: dict[str, list[float]], kerf: float = 0.0,
                                verbose: bool = False, notes_out: list | None = None,
                                progress: Callable[[str], None] | None = None):
    """Return purchase and cutting rows; report heuristic cases via notes_out."""
    cutting_plan = []
    purchase_list = []
    all_available = sorted({length for choices in stocks.values() for length in choices})
    for n, (size, pieces) in enumerate(demands.items(), 1):
        if progress:
            progress(f'Optimizing {size} ({n}/{len(demands)})...')
        if size not in stocks or not stocks[size]:
            raise ValueError(f'{size}: Select at least one market length for this diameter.')
        piece_qty = [q for q, _ in pieces]
        piece_lengths = [length for _, length in pieces]
        result = solve_with_pulp(piece_lengths, piece_qty, stocks[size], kerf, verbose)
        if result['status'] not in ('Optimal', 'Feasible'):
            raise ValueError(f'{size}: {result.get("message", "No cutting plan could be created.")}')
        if result['status'] == 'Feasible' and notes_out is not None:
            notes_out.append((size,
                              'Feasible cutting plan (not proven globally optimal); '
                              'pattern enumeration and solver runtime are bounded.'))
        row = {'Diameter': size}
        row.update({f'{length:.1f}m': 0 for length in all_available})
        for purchase in result['purchases']:
            key = f'{purchase["stock_length_m"]:.1f}m'
            if key in row:
                row[key] += purchase['quantity']
            cuts = [(quantity, length) for quantity, length in
                    zip(purchase['pattern_counts'], piece_lengths) if quantity > 0]
            descriptions = [f'{q}x{mm(length) / 1000:.3f}m' for q, length in cuts]
            cutting_plan.append({
                'Diameter': size, 'Quantity': purchase['quantity'],
                'Length': purchase['stock_length_m'], 'Cut Per RSB': descriptions,
            })
        purchase_list.append(row)
    purchase_list.sort(key=lambda item: item['Diameter'])
    cutting_plan.sort(key=lambda item: (item['Diameter'], item['Length'], item['Quantity']))
    return purchase_list, cutting_plan


if __name__ == '__main__':
    sample_demands = {'#10': [(16, 2.095), (12, 1.695)]}
    sample_stocks = {'#10': [6.0, 9.0, 12.0]}
    purchases, cuts = find_optimized_cutting_plan(sample_demands, sample_stocks, verbose=True)
    print('Purchases:', purchases)
    print('Cutting plan:', cuts)
