from utils import get_bar_dia
from typing import Literal
from math import ceil

def get_bend_diameter(bar_diameter: float, structure: Literal['stirrup', 'tensile']) -> float:
    """Calculates the minimum internal bend diameter based on ACI 318."""
    # ACI 318-25 Table 25.3.2 Minimum inside bend diameters and standard hook geometry for stirrups, ties, and hoops
    if structure == 'stirrup':
        if get_bar_dia(3, system='imperial') <= bar_diameter <= get_bar_dia(5, system='imperial'):
            return 4.0 * bar_diameter
        elif get_bar_dia(6, system='imperial') <= bar_diameter <= get_bar_dia(8, system='imperial'):
            return 6.0 * bar_diameter
        raise ValueError(f'Bar diameter {bar_diameter} not implemented.')

    # ACI 318-25 Table 25.3.1 Standard hook geometry for development of deformed bars in tension
    elif structure == 'tensile':
        if get_bar_dia(3, system='imperial') <= bar_diameter <= get_bar_dia(8, system='imperial'):
            return 6.0 * bar_diameter
        elif get_bar_dia(9, system='imperial') <= bar_diameter <= get_bar_dia(11, system='imperial'):
            return 8.0 * bar_diameter
        elif get_bar_dia(14, system='imperial') <= bar_diameter <= get_bar_dia(18, system='imperial'):
            return 10.0 * bar_diameter
        raise ValueError(f'Bar diameter {bar_diameter} not implemented.')

def get_bend_deduction(bend_angle: Literal[45, 90, 135, 180], bar_diameter: float) -> float:
    """Calculates the bend deduction value for a given angle."""
    if bend_angle == 45:
        return 1.0 * bar_diameter
    if bend_angle == 90:
        return 2.0 * bar_diameter
    elif bend_angle == 135:
        return 3.0 * bar_diameter
    elif bend_angle == 180:
        return 4.0 * bar_diameter
    raise ValueError(f'Bend angle {bend_angle} not implemented.')

def get_hook_ext(hook_angle: float | None, bar_diameter: float, structure: Literal['stirrup', 'tensile']) -> float:
    """Calculates the standard straight extension length of a hook per ACI 318."""
    if hook_angle is None:
        return 0.0

    # ACI 318-25 Table 25.3.2 Minimum inside bend diameters and standard hook geometry for stirrups, ties, and hoops
    if structure == 'stirrup':
        if hook_angle == 90:
            if get_bar_dia(3, system='imperial') <= bar_diameter <= get_bar_dia(5, system='imperial'):
                return max(6.0 * bar_diameter, 3 * 25.4)
            elif get_bar_dia(6, system='imperial') <= bar_diameter <= get_bar_dia(8, system='imperial'):
                return 12.0 * bar_diameter
            raise ValueError(f'Bar diameter {bar_diameter} not applicable for stirrup 90deg hook.')

        elif hook_angle == 135:
            if get_bar_dia(3, system='imperial') <= bar_diameter <= get_bar_dia(8, system='imperial'):
                return max(6.0 * bar_diameter, 3 * 25.4)
            raise ValueError(f'Bar diameter {bar_diameter} not applicable for stirrup 90deg hook.')

        elif hook_angle == 180:
            if get_bar_dia(3, system='imperial') <= bar_diameter <= get_bar_dia(8, system='imperial'):
                return max(4.0 * bar_diameter, 2.5 * 25.4)
            raise ValueError(f'Bar diameter {bar_diameter} not applicable for stirrup 90deg hook.')

        raise ValueError(f'Hook angle {hook_angle} not implemented.')

    # ACI 318-25 Table 25.3.1 Standard hook geometry for development of deformed bars in tension
    elif structure == 'tensile':
        if hook_angle == 90:
            return 12 * bar_diameter
        elif hook_angle == 180:
            return max(4.0 * bar_diameter, 2.5 * 25.4)
        raise ValueError(f'Hook angle {hook_angle} not implemented.')

def get_hook_length(hook_angle: float | None, bar_diameter: float, structure: Literal['stirrup', 'tensile']) -> float:
    if hook_angle is None:
        return 0.0
    bend_radius = get_bend_diameter(bar_diameter, structure) / 2
    return get_hook_ext(hook_angle, bar_diameter, structure) + bend_radius + bar_diameter

def perimeter_bar_calculation(bar_diameter: float, layers: int, pad_width_x: float, pad_width_y: float,
                               concrete_cover: float) -> dict:
    """
    Calculates rebar for the perimeter.
    """
    length_x = pad_width_x - concrete_cover * 2
    length_y = pad_width_y - concrete_cover * 2

    shape_dim_x = {'A': length_x}
    shape_dim_y = {'A': length_y}

    cut_length_x = sum(shape_dim_x.values())
    cut_length_y = sum(shape_dim_y.values())

    qty_x_dir = layers * 2
    qty_y_dir = layers * 2

    return {
        'bar_in_x_direction': {
            'shape': 'straight',
            'diameter': bar_diameter,
            'shape_dimensions': {k: round(v, 1) for k, v in shape_dim_x.items()},
            'total_cut_length_mm': round(cut_length_x, 1),
            'quantity': int(qty_x_dir)
        },
        'bar_in_y_direction': {
            'shape': 'straight',
            'diameter': bar_diameter,
            'shape_dimensions': {k: round(v, 1) for k, v in shape_dim_y.items()},
            'total_cut_length_mm': round(cut_length_y, 1),
            'quantity': int(qty_y_dir)
        }
    }


def top_bottom_bar_calculation(bar_diameter: float,
                               pad_width_x: float, pad_width_y: float, pad_thickness: float,
                               concrete_cover: float, spacing_x: float | None = None, spacing_y: float | None = None,
                               quantity_x: float | None = None, quantity_y: float | None = None) -> dict:
    """
    Calculates rebar for the top and bottom mat of a pad footing.
    """
    if (spacing_x is None or (spacing_y is None)) and (quantity_x is None or (quantity_y is None)):
        raise ValueError("Either 'spacing' or 'quantity' must be provided.")

    vertical_length = pad_thickness - concrete_cover * 2
    main_span_x = pad_width_x - concrete_cover * 2
    main_span_y = pad_width_y - concrete_cover * 2

    shape_dim_x = {'A': vertical_length, 'B': main_span_x, 'C': vertical_length}
    shape_dim_y = {'A': vertical_length, 'B': main_span_y, 'C': vertical_length}

    total_deduction = 2 * get_bend_deduction(90, bar_diameter)

    cut_length_x = sum(shape_dim_x.values()) - total_deduction
    cut_length_y = sum(shape_dim_y.values()) - total_deduction

    qty_x_dir = quantity_x if quantity_x is not None else (ceil(main_span_y / spacing_x) + 1)
    qty_y_dir = quantity_y if quantity_y is not None else (ceil(main_span_x / spacing_y) + 1)

    return {
        'bar_in_x_direction': {
            'shape': 'U',
            'diameter': bar_diameter,
            'shape_dimensions': {k: round(v, 1) for k, v in shape_dim_x.items()},
            'total_cut_length_mm': round(cut_length_x, 1),
            'quantity': int(qty_x_dir)
        },
        'bar_in_y_direction': {
            'shape': 'U',
            'diameter': bar_diameter,
            'shape_dimensions': {k: round(v, 1) for k, v in shape_dim_y.items()},
            'total_cut_length_mm': round(cut_length_y, 1),
            'quantity': int(qty_y_dir)
        }
    }


def vertical_bar_calculation(bar_diameter: float,
                             qty: int,
                             ped_height: float,
                             pad_thickness: float,
                             concrete_cover: float,
                             bot_bar_diameter: float,
                             hook_length: float | None = None) -> dict:
    if hook_length is None:
        hook_length = get_hook_length(90, bar_diameter, 'tensile')
    vertical_length = ped_height + pad_thickness - concrete_cover * 2 - 2 * bot_bar_diameter
    shape_dim = {'A': vertical_length, 'B': hook_length}

    total_deduction = get_bend_deduction(90, bar_diameter)
    cut_length = sum(shape_dim.values()) - total_deduction

    return {
        'shape': 'L',
        'diameter': bar_diameter,
        'shape_dimensions': {k: round(v, 1) for k, v in shape_dim.items()},
        'total_cut_length_mm': round(cut_length, 1),
        'quantity': int(qty)
    }


def stirrups_calculation(bar_diameter: float,
                         qty: int,
                         ped_width_x: float,
                         ped_width_y: float,
                         concrete_cover: float,
                         config: Literal['outer', 'tall', 'wide', 'diamond', 'octagon', 'vertical', 'horizontal'] = 'outer',
                         a: float | None = None) -> dict:
    if config in ['tall', 'wide', 'octagon'] and a is None:
        raise ValueError(f"Parameter 'a' must be provided for '{config}' stirrups.")

    hook_length = get_hook_length(135, bar_diameter, 'stirrup')
    if config == 'outer':
        shape = 'rectangular'

        stirrup_width_x = ped_width_x - concrete_cover * 2
        stirrup_width_y = ped_width_y - concrete_cover * 2

        shape_dim = {'A': hook_length, 'B': stirrup_width_x, 'C': stirrup_width_y,
                 'D': stirrup_width_x, 'E': stirrup_width_y, 'F': hook_length}

        total_deduction = 3 * get_bend_deduction(90, bar_diameter) + 2 * get_bend_deduction(135, bar_diameter)
        cut_length = sum(shape_dim.values()) - total_deduction

    elif config == 'tall':
        shape = 'rectangular (tall)'

        tall = ped_width_y - concrete_cover * 2
        shape_dim = {'A': hook_length, 'B': a, 'C': tall,
                 'D': a, 'E': tall, 'F': hook_length}

        total_deduction = 3 * get_bend_deduction(90, bar_diameter) + 2 * get_bend_deduction(135, bar_diameter)
        cut_length = sum(shape_dim.values()) - total_deduction

    elif config == 'wide':
        shape = 'rectangular (wide)'

        wide = ped_width_x - concrete_cover * 2
        shape_dim = {'A': hook_length, 'B': wide, 'C': a,
                 'D': wide, 'E': a, 'F': hook_length}

        total_deduction = 3 * get_bend_deduction(90, bar_diameter) + 2 * get_bend_deduction(135, bar_diameter)
        cut_length = sum(shape_dim.values()) - total_deduction

    elif config == 'diamond':
        shape = 'rectangular (diamond)'

        x = (ped_width_x - concrete_cover * 2)/2.0
        y = (ped_width_y - concrete_cover * 2)/2.0
        hypotenuse = (x**2 + y**2)**0.5
        shape_dim = {'A': hook_length, 'B': hypotenuse, 'C': hypotenuse,
                 'D': hypotenuse, 'E': hypotenuse, 'F': hook_length}

        total_deduction = 3 * get_bend_deduction(90, bar_diameter) + 2 * get_bend_deduction(135, bar_diameter)
        cut_length = sum(shape_dim.values()) - total_deduction

    elif config == 'octagon':
        shape = 'octagonal'

        x = (ped_width_x - concrete_cover * 2 - a)/2.0
        y = (ped_width_y - concrete_cover * 2 - a)/2.0
        hypotenuse = (x**2 + y**2)**0.5
        shape_dim = {'A': hook_length, 'B': a, 'C': hypotenuse, 'D': a, 'E': hypotenuse,
                 'F': a, 'G': hypotenuse, 'H': a, 'I': hypotenuse, 'J': hook_length}

        total_deduction = 7 * get_bend_deduction(45, bar_diameter) + 2 * get_bend_deduction(135, bar_diameter)
        cut_length = sum(shape_dim.values()) - total_deduction

    elif config == 'vertical':
        shape = 'flat (tall)'

        tall = ped_width_y - concrete_cover * 2
        shape_dim = {'A': hook_length, 'B': tall, 'C': hook_length}

        total_deduction = 2 * get_bend_deduction(180, bar_diameter)
        cut_length = sum(shape_dim.values()) - total_deduction

    elif config == 'horizontal':
        shape = 'flat (wide)'

        wide = ped_width_x - concrete_cover * 2
        shape_dim = {'A': hook_length, 'B': wide, 'C': hook_length}

        total_deduction = 2 * get_bend_deduction(180, bar_diameter)
        cut_length = sum(shape_dim.values()) - total_deduction

    else:
        raise ValueError(f"Stirrup Type '{config}' not supported.")

    return {
        'shape': shape,
        'diameter': bar_diameter,
        'shape_dimensions': {k: round(v, 1) for k, v in shape_dim.items()},
        'total_cut_length_mm': round(cut_length, 1),
        'quantity': qty
    }


def compile_rebar(data: dict) -> dict:
    n_ped = data['n_ped']
    n_footing = data['n_footing']
    cc = data['cc']
    bx = data['bx']
    by = data['by']
    h = data['h']
    Bx = data['Bx']
    By = data['By']
    t = data['t']

    result = {}

    # Top and Bottom Bar
    def top_bottom_bar_helper(title):
        bar_detail = data[title]
        qty_or_spacing = bar_detail['Input Type']
        value_along_x = bar_detail['Value Along X']
        value_along_y = bar_detail['Value Along Y']
        if 'Spacing' in qty_or_spacing:
            bar_spacing_value_x = value_along_x
            bar_spacing_value_y = value_along_y
            bar_qty_x = bar_qty_y = None
        else:
            bar_spacing_value_x = bar_spacing_value_y = None
            bar_qty_x = value_along_x
            bar_qty_y = value_along_y
        return top_bottom_bar_calculation(get_bar_dia(bar_detail['Diameter']), Bx, By, t, cc,
                                          bar_spacing_value_x,
                                          bar_spacing_value_y, bar_qty_x, bar_qty_y)

    if data['Top Bar']['Enabled']:
        result['Top Bar'] = top_bottom_bar_helper('Top Bar')
        result['Top Bar']['bar_in_x_direction']['quantity'] *= n_footing
        result['Top Bar']['bar_in_y_direction']['quantity'] *= n_footing
    result['Bottom Bar'] = top_bottom_bar_helper('Bottom Bar')
    result['Bottom Bar']['bar_in_x_direction']['quantity'] *= n_footing
    result['Bottom Bar']['bar_in_y_direction']['quantity'] *= n_footing

    # Perimeter Bar
    perim_bar = data['Perimeter Bar']
    if perim_bar['Enabled']:
        layers = perim_bar['Layers']
        if layers > 0:
            dia = get_bar_dia(perim_bar['Diameter'])
            result['Perimeter Bar'] = perimeter_bar_calculation(dia, layers, Bx, By, cc)
            result['Perimeter Bar']['bar_in_x_direction']['quantity'] *= n_footing
            result['Perimeter Bar']['bar_in_y_direction']['quantity'] *= n_footing

    # Vertical Bar
    vert_bar = data['Vertical Bar']
    if vert_bar['Enabled']:
        dia = get_bar_dia(vert_bar['Diameter'])
        hook_calc = vert_bar['Hook Calculation']
        bot_bar_dia = get_bar_dia(data['Bottom Bar']['Diameter'])
        qty = vert_bar['Quantity']
        if 'Manual' in hook_calc:
            hook_len = vert_bar['Hook Length']
        else:
            hook_len = None
        result['Vertical Bar'] = vertical_bar_calculation(dia, qty, h, t, cc, bot_bar_dia, hook_len)
        result['Vertical Bar']['quantity'] *= n_ped * n_footing

    # Stirrups
    stirrup = data['Stirrups']
    if stirrup['Enabled']:
        qty = stirrup['Quantity']
        if qty > 0:
            stirrups_cutting_list = []
            for row in stirrup['Types']:
                dia = get_bar_dia(row['Diameter'])
                a = row['a_input']
                if a == '':
                    a = None
                stirrup_result = stirrups_calculation(dia, qty, bx, by, cc, row['Type'].lower(), a)
                stirrup_result['quantity'] *= n_footing
                stirrups_cutting_list.append(stirrup_result)
            result['Stirrups'] = stirrups_cutting_list

    return result
