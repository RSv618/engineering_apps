import pulp
from typing import Any
from utils import resource_path


def mm(x_m: float) -> int:
    """
    Converts a length from meters to integer millimeters to avoid floating point issues.

    Args:
        x_m: Length in meters.

    Returns:
        Length in millimeters, rounded to the nearest integer.
    """
    return int(x_m * 1000 + 0.5)

def m(x_mm: int) -> float:
    """
    Converts a length from integer millimeters back to meters.

    Args:
        x_mm: Length in millimeters.

    Returns:
        Length in meters.
    """
    return x_mm / 1000.0

def get_solver_path():
    """Gets the correct path to the CBC solver executable."""
    # The default solver is included with pulp
    # We can get its path relative to the pulp library
    import pulp.apis
    solver_path = pulp.apis.PULP_CBC_CMD().path
    # resource_path will handle the temporary folder for PyInstaller
    return resource_path(solver_path)

def enumerate_patterns(stock_len_mm: int, piece_lengths_mm: list[int], max_counts: list[int]) -> list:
    """
    Enumerate integer patterns for one stock length.
    Each pattern is a tuple (counts_tuple, used_length_mm_without_kerf).
    """
    n = len(piece_lengths_mm)
    # Upper bound for the count of each piece in a single stock
    ub = [min(max_counts[i], stock_len_mm // piece_lengths_mm[i]) for i in range(n)]
    patterns = []

    def rec(i, cur_counts, cur_used):
        if i == n:
            if any(c > 0 for c in cur_counts):
                patterns.append((tuple(cur_counts), cur_used))
            return

        pl = piece_lengths_mm[i]
        for cnt in range(ub[i] + 1):
            new_used = cur_used + cnt * pl
            if new_used <= stock_len_mm:
                cur_counts.append(cnt)
                rec(i + 1, cur_counts, new_used)
                cur_counts.pop()
            else:
                # Since piece_lengths are sorted, no further counts will fit
                break

    rec(0, [], 0)
    return patterns

def build_patterns_all_stocks(piece_lengths: list[float], piece_qty: list[int],
                              stock_lengths: list[float], kerf: float = 0.0) -> dict:
    """
    Returns a dict stock_mm -> list of (counts_tuple, used_with_kerf_mm)
    """
    piece_mm = [mm(x) for x in piece_lengths]
    stock_mm = sorted([mm(x) for x in stock_lengths])  # Sorting can help pattern generation
    kerf_mm = mm(kerf)

    patterns_by_stock = {}
    for s in stock_mm:
        # Pass total quantities as max_counts for enumeration
        patt = enumerate_patterns(s, piece_mm, piece_qty)
        patt2 = []
        for counts, used in patt:
            k_pieces = sum(counts)
            # Kerf is only applied if there is more than one piece
            used_with_kerf = used + max(0, k_pieces - 1) * kerf_mm
            if used_with_kerf <= s:
                patt2.append((counts, used_with_kerf))
        if patt2:
            patterns_by_stock[s] = patt2

    return patterns_by_stock

def solve_with_pulp(piece_lengths: list[float], piece_qty: list[int],
                    stock_lengths: list[float], kerf: float = 0.0, verbose: bool = False) -> dict[str, Any]:
    """
    Build and solve the cutting stock ILP via PuLP using a single-stage optimization.
    Returns a structured solution dict.
    """
    piece_mm = [mm(x) for x in piece_lengths]
    stock_mm = [mm(x) for x in stock_lengths]
    patterns_by_stock = build_patterns_all_stocks(piece_lengths, piece_qty, stock_lengths, kerf)

    # Flatten patterns for easy indexing
    pattern_index = []
    for s, patt_list in patterns_by_stock.items():
        for idx, (counts, used) in enumerate(patt_list):
            pattern_index.append({'stock_mm': s, 'pattern_idx': idx, 'counts': counts, 'used_mm': used})

    if not pattern_index:
        return {'status': 'NoPatterns',
                'message': 'No feasible patterns found. Check piece sizes, stock lengths, or kerf.'}

    # Create the ILP problem
    prob = pulp.LpProblem('rebar_cutting_stock', pulp.LpMinimize)

    # Decision variables: How many times to use each pattern
    y = pulp.LpVariable.dicts('pattern', range(len(pattern_index)), lowBound=0, cat='Integer')

    # Demand constraints: Ensure we produce at least the required quantity of each piece
    n_pieces = len(piece_mm)
    for i in range(n_pieces):
        prob += (
            pulp.lpSum(y[j] * pattern_index[j]['counts'][i] for j in range(len(pattern_index))) >= piece_qty[i],
            f'demand_{i}'
        )

    # --- Combined Objective Function ---
    # Primary objective: Minimize total purchased length (cost)
    # Secondary objective: Minimize total waste
    # We combine them using a large weight M for the primary objective.

    total_purchased_mm = pulp.lpSum(y[j] * p['stock_mm'] for j, p in enumerate(pattern_index))
    total_used_mm = pulp.lpSum(y[j] * p['used_mm'] for j, p in enumerate(pattern_index))
    total_waste_mm = total_purchased_mm - total_used_mm

    # M must be larger than the maximum possible total waste to prioritize cost saving.
    # A safe value is 1 + the maximum possible stock length in mm.
    # M = max(stock_mm) + 1

    prob.setObjective((max(stock_mm) + 1) * total_purchased_mm + total_waste_mm)

    # Solve the problem
    # solver = pulp.PULP_CBC_CMD(path=get_solver_path(), msg=verbose)
    solver = pulp.COIN_CMD(path=get_solver_path(), msg=verbose)  # Resolves pulp error on a different windows
    prob.solve(solver)

    if pulp.LpStatus[prob.status] != 'Optimal':
        return {'status': pulp.LpStatus[prob.status], 'message': 'Optimal solution not found.'}

    # --- Collect Results ---
    purchases = []
    final_total_purchased_mm = 0
    final_total_used_mm = 0
    produced = [0] * n_pieces

    for j, p in enumerate(pattern_index):
        quantity = int(round(pulp.value(y[j])))
        if quantity > 0:
            stock = p['stock_mm']
            used = p['used_mm']
            purchases.append({
                'stock_length_m': m(stock),
                'pattern_counts': p['counts'],
                'used_length_m': m(used),
                'waste_m': m(stock - used),
                'quantity': quantity
            })
            final_total_purchased_mm += stock * quantity
            final_total_used_mm += used * quantity
            for i in range(n_pieces):
                produced[i] += p['counts'][i] * quantity

    solution = {
        'status': 'Optimal',
        'total_purchased_m': m(final_total_purchased_mm),
        'total_used_m': m(final_total_used_mm),
        'total_waste_m': m(final_total_purchased_mm - final_total_used_mm),
        'purchases': purchases,
        'produced_counts': produced,
        'demand_counts': piece_qty,
    }
    return solution

def find_optimized_cutting_plan(demands: dict[str, list[tuple]], stocks: dict[str, list[float]], kerf: float = 0.0,
                                verbose: bool = False):
    cutting_plan = []
    purchase_list = []
    all_available_stocks = sorted(list(set(l for stock_list in stocks.values() for l in stock_list)))

    for size in demands:
        piecelist = demands[size]
        piece_qty = [q for (q, l) in piecelist]
        piece_lengths = [l for (q, l) in piecelist]

        result = solve_with_pulp(piece_lengths, piece_qty, stocks[size], kerf=kerf, verbose=verbose)

        if result['status'] != 'Optimal':
            cutting_plan.append({'Error': result['message'], 'Diameter': size, 'Length': None, 'Quantity': None})
            print(f'Could not find optimal solution for diameter {size}: {result['message']}')
            continue

        row: dict[str, int | float | str] = {'Diameter': size}
        row.update({f'{l:.1f}m': 0 for l in all_available_stocks})

        for res in result['purchases']:
            # Use a robust way to format the stock length key
            market_length_key = f"{res['stock_length_m']:.1f}m"
            quantity = res['quantity']
            if market_length_key in row:
                row[market_length_key] += quantity

            cut = [(q, l) for q, l in zip(res['pattern_counts'], piece_lengths) if q > 0]
            cut_per_rsb = [f'{q}x{int(l * 1000 + 0.5) / 1000:0.3f}m' for q, l in cut]
            cutting_plan.append({'Diameter': size, 'Quantity': quantity, 'Length': res['stock_length_m'],
                                 'Cut Per RSB': cut_per_rsb})
        purchase_list.append(row)

    purchase_list = sorted(purchase_list, key=lambda item: item['Diameter'])
    cutting_plan = sorted(cutting_plan, key=lambda x: (x['Diameter'], x['Length'], x['Quantity']))
    return purchase_list, cutting_plan


if __name__ == '__main__':
    # --- Example Test Case ---
    print("--- Running Rebar Optimizer Test ---")

    # Define a sample problem
    sample_demands = {
        '#10': [(16, 2.095), (12, 1.695)]  # 16 pcs of 2.095m, 12 pcs of 1.695m
    }
    sample_stocks = {
        '#10': [6.0, 9.0, 12.0]
    }

    # Run the main function
    purchase_list_result, cutting_plan_result = find_optimized_cutting_plan(sample_demands, sample_stocks, verbose=True)

    # Print the results
    print("\n--- PURCHASE LIST ---")
    for this in purchase_list_result:
        print(this)

    print("\n--- CUTTING PLAN ---")
    for this in cutting_plan_result:
        print(this)