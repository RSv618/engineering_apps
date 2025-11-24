from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from typing import Any
import collections
from utils import get_dia_code, resource_path
from rebar_optimizer import find_optimized_cutting_plan

def excel_col_width_to_px(width: float | None) -> int:
    """
    Approximates the conversion of an openpyxl column width to pixels.

    Args:
        width: The column width in openpyxl units.

    Returns:
        The approximate width in pixels.
    """
    return int((width + 0.71) * 7) if width is not None else int((8.43 + 0.71) * 7)


def excel_row_height_to_px(height: float | None) -> float:
    """
    Converts an openpyxl row height (in points) to pixels.

    Args:
        height: The row height in points.

    Returns:
        The height in pixels.
    """
    return (height if height is not None else 15) * 96 / 72


def center_img(img: Image, cell_ref: str, ws: Worksheet) -> Image:
    """
    Calculates the anchor position to center an image within a worksheet cell.

    Args:
        img: The openpyxl Image object.
        cell_ref: The cell reference string (e.g., 'A3').
        ws: The openpyxl Worksheet object.

    Returns:
        The Image object with its anchor property set.
    """
    cell = ws[cell_ref]
    row, col = cell.row, cell.column

    # --- cumulative Y offset ---
    row_heights = [
        ws.row_dimensions[r].height or ws.sheet_format.defaultRowHeight
        for r in range(1, row)
    ]
    anchor_y = excel_row_height_to_px(sum(row_heights))

    # --- current row height ---
    row_h = excel_row_height_to_px(ws.row_dimensions[row].height or ws.sheet_format.defaultRowHeight)

    # --- cumulative X offset ---
    col_widths = [
        ws.column_dimensions[get_column_letter(c)].width or ws.sheet_format.defaultColWidth
        for c in range(1, col)
    ]
    anchor_x = excel_col_width_to_px(sum(col_widths))

    # --- current column width ---
    col_w = excel_col_width_to_px(ws.column_dimensions[get_column_letter(col)].width or ws.sheet_format.defaultColWidth)

    # --- offsets to center ---
    y_offset = max((row_h - img.height) / 2, 0)
    x_offset = max((col_w - img.width) / 2, 0)

    # --- build anchor ---
    pos = XDRPoint2D(pixels_to_EMU(anchor_x + x_offset), pixels_to_EMU(anchor_y + y_offset))
    size = XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height))
    img.anchor = AbsoluteAnchor(pos=pos, ext=size)
    return img

def get_canonical_representation(bar: dict[str, Any]) -> tuple[str, tuple]:
    """
    Creates a unique, normalized key for a bar based on its shape and dimensions
    to group identical or mirrored shapes.
    """
    shape = bar['shape']
    dims = bar['shape_dimensions']

    # Normalize shape names for grouping (e.g., tall/wide rectangles are the same shape)
    if 'rectangular' in shape and shape != 'rectangular (diamond)':
        canonical_shape = 'rectangular'
        key = (
            dims.get('A', 0),
            min(dims.get('B', 0), dims.get('C', 0)),
            max(dims.get('B', 0), dims.get('C', 0))
        )
        return canonical_shape, key
    elif shape == 'rectangular (diamond)':
        key = (dims.get('A', 0), dims.get('B', 0))
        return shape, key
    elif shape == 'U':
        key = (
            min(dims.get('A', 0), dims.get('C', 0)),
            dims.get('B', 0),
            max(dims.get('A', 0), dims.get('C', 0))
        )
        return shape, key
    elif shape == 'L':
        key = (dims.get('A', 0), dims.get('B', 0))
        return shape, key
    elif shape == 'flat':
        key = (dims.get('A', 0), dims.get('B', 0))
        return shape, key
    else:
        key = tuple(dims.get(k, 0) for k in sorted(dims.keys()))
        return shape, key

def process_rebar_input(rebar_config: dict[str, Any] | list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Flattens the input dictionary and groups identical bars by shape, dimensions, AND diameter.
    """
    flat_list = []
    if isinstance(rebar_config, dict):
        for bar_type, data in rebar_config.items():
            if bar_type in ['Top Bar', 'Bottom Bar', 'Perimeter Bar']:
                if 'bar_in_x_direction' in data:
                    flat_list.append({'bar_type': bar_type, **data['bar_in_x_direction']})
                if 'bar_in_y_direction' in data:
                    flat_list.append({'bar_type': bar_type, **data['bar_in_y_direction']})
            elif bar_type == 'Vertical Bar':
                flat_list.append({'bar_type': bar_type, **data})
            elif bar_type == 'Stirrups':
                for stirrup in data:
                    flat_list.append({'bar_type': bar_type, **stirrup})
    elif isinstance(rebar_config, list):
        for rebar_config_instance in rebar_config:
            for bar_type, data in rebar_config_instance.items():
                if bar_type in ['Top Bar', 'Bottom Bar', 'Perimeter Bar']:
                    if 'bar_in_x_direction' in data:
                        flat_list.append({'bar_type': bar_type, **data['bar_in_x_direction']})
                    if 'bar_in_y_direction' in data:
                        flat_list.append({'bar_type': bar_type, **data['bar_in_y_direction']})
                elif bar_type == 'Vertical Bar':
                    flat_list.append({'bar_type': bar_type, **data})
                elif bar_type == 'Stirrups':
                    for stirrup in data:
                        flat_list.append({'bar_type': bar_type, **stirrup})
    else:
        raise TypeError(f'Invalid rebar_config type. Expected dict or list, got {type(rebar_config)}')

    grouped_bars = collections.OrderedDict()
    for bar in flat_list:
        canonical_shape, dim_key = get_canonical_representation(bar)
        group_key = (canonical_shape, bar['diameter'], dim_key)

        if group_key not in grouped_bars:
            grouped_bars[group_key] = {
                'original_shape': bar['shape'],
                'bar_types': {bar['bar_type']},
                'quantity': 0,
                'diameter': bar['diameter'],
                'cut_length': bar['total_cut_length_mm'],
                'shape_dimensions': bar['shape_dimensions']
            }

        grouped_bars[group_key]['quantity'] += bar['quantity']
        grouped_bars[group_key]['bar_types'].add(bar['bar_type'])

    processed_list = []
    for data in grouped_bars.values():
        processed_list.append({
            'shape': data['original_shape'],
            'bar_type': ',\n'.join(sorted(list(data['bar_types']))),
            'quantity': data['quantity'],
            'diameter': data['diameter'],
            'cut_length': data['cut_length'],
            'shape_dimensions': data['shape_dimensions']
        })

    return processed_list

def create_excel_cutting_list(rebar_config: dict[str, Any],
                              cuts_by_diameter: dict,
                              market_lengths: dict[str, list],
                              output_filename: str = 'rebar_cutting_schedule.xlsx'):
    """
    Generates a formatted Excel rebar cutting list from a rebar configuration dictionary.
    """
    processed_data = process_rebar_input(rebar_config)
    purchase_list, cutting_plan = find_optimized_cutting_plan(cuts_by_diameter, market_lengths)

    proceed_cutting_plan = True
    for plan in cutting_plan:
        if 'Error' in plan:
            proceed_cutting_plan = False

    max_legs = 0
    if processed_data:
        max_legs = max(len(bar['shape_dimensions']) for bar in processed_data)

    # This list comprehension is creating the dictionary keys 'A', 'B', etc., which is fine.
    dimension_headers = [chr(ord('A') + i) for i in range(max_legs)]

    wb = Workbook()
    ws = wb.active
    ws.title = 'Rebar Cutting Schedule'

    # --- Styles ---
    white_side = Side(style='thin', color='FFFFFF')
    black_side = Side(style='thin', color='404040')
    thick_black_side = Side(style='thick', color='404040')
    title_font = Font(name='Calibri', size=16, bold=True)
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='404040', end_color='404040', fill_type='solid')
    alter_row_fill = PatternFill(start_color='F3F3F3', end_color='404040', fill_type='solid')
    cell_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border = Border(left=black_side, right=black_side, top=black_side, bottom=black_side)
    header_border = Border(left=white_side, right=white_side, top=black_side, bottom=black_side)

    # --- Static and Dynamic Headers ---
    static_headers = ['Illustration', 'Bar Type', 'Diameter', 'Quantity', 'Cut Length']
    all_headers = static_headers + dimension_headers

    # --- Title ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(static_headers) + max_legs)
    title_cell = ws['A1']
    title_cell.value = 'Rebar Cutting and Bending Schedule'
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    if not proceed_cutting_plan:
        for plan in cutting_plan:
            if 'Error' in plan:
                title_cell.comment = Comment(plan['Error'], '✨rs_uy')
    ws.row_dimensions[1].height = 30

    # --- Headers ---
    for col_num, header_text in enumerate(all_headers, 1):
        cell = ws.cell(row=2, column=col_num, value=header_text)
        cell.font = header_font
        cell.alignment = cell_alignment
        cell.fill = header_fill
        cell.border = header_border

    # Apply black border to left and right outer edges
    cell = ws.cell(row=2, column=1)
    cell.border = Border(left=black_side, top=black_side, right=white_side, bottom=black_side)
    cell = ws.cell(row=2, column=len(all_headers))
    cell.border = Border(left=white_side, top=black_side, right=black_side, bottom=black_side)

    # --- Column Widths ---
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 18
    ws.row_dimensions[2].height = 25

    num_static_cols = len(static_headers)
    for i in range(max_legs):
        # Calculate column index: static columns + 1 (for 1-based index) + loop index
        col_idx = num_static_cols + 1 + i
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 12

    shape_to_image_map = {
        'straight': 'straight.png',
        'U': 'u.png',
        'L': 'l.png',
        'rectangular': 'rectangular_outer.png',
        'rectangular (tall)': 'rectangular_tall.png',
        'rectangular (wide)': 'rectangular_wide.png',
        'rectangular (diamond)': 'rectangular_diamond.png',
        'octagonal': 'octagon.png'
    }

    # --- Data Rows ---
    current_row = 3
    for bar in processed_data:
        ws.row_dimensions[current_row].height = 75

        # Use the map to get the correct image filename
        image_filename = shape_to_image_map.get(bar['shape'])
        if image_filename:
            try:
                img_path = f'images/{image_filename}'

                with PILImage.open(img_path) as pil_img:
                    original_width, original_height = pil_img.size

                aspect_ratio = original_height / original_width
                target_width = 90
                target_height = int(target_width * aspect_ratio + 0.5)
                if target_height > 90:
                    target_height = 90
                    target_width = int(90 / aspect_ratio + 0.5)

                img = Image(img_path)
                img.width = target_width
                img.height = target_height

                ws.add_image(center_img(img, f'A{current_row}', ws))

            except FileNotFoundError:
                ws.cell(row=current_row, column=1, value='No Image')
        else:
            ws.cell(row=current_row, column=1, value='No Image')

        try:
            val = bar['diameter']
            diameter_str = f'{val:.1f} mm\n({get_dia_code(val)})'
        except KeyError:
            # Fallback if the diameter code is not found
            diameter_str = f'{bar['diameter']:.1f} mm'

        data_to_write = [
            None,
            bar['bar_type'],
            diameter_str,  # Use the newly formatted string
            bar['quantity'],
            round(bar['cut_length'], 1),
        ]

        for letter in dimension_headers:
            data_to_write.append(bar['shape_dimensions'].get(letter, '-'))

        for col_num, value in enumerate(data_to_write, 1):
            if value is None:
                cell = ws.cell(row=current_row, column=col_num)
            else:
                cell = ws.cell(row=current_row, column=col_num, value=value)

            if col_num == 5:  # Cutlength
                max_length = max(market_lengths[get_dia_code(bar['diameter'])])
                if value > max_length * 1000:
                    proceed_cutting_plan = False
                    cell.font = Font(color='FF0000')
                    cell.comment = Comment(f'Splicing required.\nCutting length exceeds available market length of {max_length:}m.\n'
                                    f'Cannot proceed with purchase order analysis.', '✨rs_uy', height=150, width=200)

            # Alternating BG Color Fill
            if current_row % 2 == 0:
                cell.fill = alter_row_fill
            cell.alignment = cell_alignment

            # Borders
            cell.border = border
            if col_num == 6:  #Shape Dimensions
                cell.border = Border(left=thick_black_side, bottom=black_side, top=black_side, right=black_side)

            # Number Format
            if col_num >= 5:
                cell.number_format = '#,##0" mm"'

        current_row += 1

    if proceed_cutting_plan:
        add_sheet_purchase_plan(wb, purchase_list)
        add_sheet_cutting_plan(wb, cutting_plan)
    wb.save(output_filename)
    print(f'Excel sheet {output_filename} has been created successfully.')

def add_sheet_purchase_plan(wb, purchase_list) -> Workbook:
    ws = wb.create_sheet('Rebar Purchase')

    # --- Styles ---
    white_side = Side(style='thin', color='FFFFFF')
    black_side = Side(style='thin', color='404040')
    title_font = Font(name='Calibri', size=16, bold=True)
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='404040', end_color='404040', fill_type='solid')
    alter_row_fill = PatternFill(start_color='F3F3F3', end_color='404040', fill_type='solid')
    cell_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border = Border(left=black_side, right=black_side, top=black_side, bottom=black_side)
    header_border = Border(left=white_side, right=white_side, top=black_side, bottom=black_side)

    # --- Static and Dynamic Headers ---
    headers = purchase_list[0].keys()

    # --- Title ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell = ws['A1']
    title_cell.value = 'Purchase Qty by Length & Diameter'
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # --- Headers ---
    for col_num, header_text in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_num, value=header_text)
        cell.font = header_font
        cell.alignment = cell_alignment
        cell.fill = header_fill
        cell.border = header_border

    # Apply black border to left and right outer edges
    cell = ws.cell(row=2, column=1)
    cell.border = Border(left=black_side, top=black_side, right=white_side, bottom=black_side)
    cell = ws.cell(row=2, column=len(headers))
    cell.border = Border(left=white_side, top=black_side, right=black_side, bottom=black_side)

    # --- Column Widths ---
    ws.column_dimensions['A'].width = 20
    ws.row_dimensions[2].height = 25

    for item in purchase_list:
        ws.append(list(item.values()))

    for current_row in range(3, 3 + len(purchase_list)):
        ws.row_dimensions[current_row].height = 25

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_num)

            if cell.value == 0:
                cell.value = ''

            # Alternating BG Color Fill
            if current_row % 2 == 0:
                cell.fill = alter_row_fill
            cell.alignment = cell_alignment

            # Borders
            cell.border = border

            # Number Format
            if col_num >= 2:
                cell.number_format = '#,##0'
    return wb

def add_sheet_cutting_plan(wb, cutting_plan) -> Workbook:
    ws = wb.create_sheet('Cutting Plan')

    # --- Styles ---
    white_side = Side(style='thin', color='FFFFFF')
    black_side = Side(style='thin', color='404040')
    title_font = Font(name='Calibri', size=16, bold=True)
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='404040', end_color='404040', fill_type='solid')
    alter_row_fill = PatternFill(start_color='F3F3F3', end_color='404040', fill_type='solid')
    cell_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border = Border(left=black_side, right=black_side, top=black_side, bottom=black_side)
    header_border = Border(left=white_side, right=white_side, top=black_side, bottom=black_side)

    # --- Static and Dynamic Headers ---
    headers = ['Diameter', 'Quantity', 'Length', 'Cuts', 'Detailed Instructions']

    # --- Title ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell = ws['A1']
    title_cell.value = 'Cutting Plan'
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # --- Headers ---
    for col_num, header_text in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_num, value=header_text)
        cell.font = header_font
        cell.alignment = cell_alignment
        cell.fill = header_fill
        cell.border = header_border

    # Apply black border to left and right outer edges
    cell = ws.cell(row=2, column=1)
    cell.border = Border(left=black_side, top=black_side, right=white_side, bottom=black_side)
    cell = ws.cell(row=2, column=len(headers))
    cell.border = Border(left=white_side, top=black_side, right=black_side, bottom=black_side)

    # --- Column Widths ---
    ws.column_dimensions['E'].width = 70
    ws.column_dimensions['D'].width = 15
    ws.row_dimensions[2].height = 25
    col_map = {i:header for i, header in enumerate(headers, 1)}
    for current_row in range(3, 3 + len(cutting_plan)):
        ws.row_dimensions[current_row].height = 50

        for col_num in range(1, len(headers) + 1):
            data = cutting_plan[current_row - 3]
            if col_num < 3:
                cell = ws.cell(row=current_row, column=col_num, value=data[col_map[col_num]])
            elif col_map[col_num] == 'Length':
                cell = ws.cell(row=current_row, column=col_num, value=f'{data[col_map[col_num]]}m')
            elif col_map[col_num] == 'Cuts':
                # get cuts
                cuts = '\n'.join(data['Cut Per RSB'])
                cell = ws.cell(row=current_row, column=col_num, value=cuts)
            else: # col_num == 'Detailed Instructions':
                # get detailed instructions
                qty = data['Quantity']
                length = data['Length']
                dia = data['Diameter']
                cuts = [cut.replace('x', '×') for cut in data['Cut Per RSB']]
                # Cut each of the 4 pcs of 13.5m Ø10 bars into 4×2.095m and 3×1.695m lengths.
                if len(cuts) > 2:
                    cuts = cuts[:-1] + ['and ' + cuts[-1]]
                    cuts = ', '.join(cuts)
                elif len(cuts) == 2:
                    cuts = ' and '.join(cuts)
                else:
                    cuts = cuts[0]

                if qty > 1:
                    value = f'  Cut each of the {qty}pcs of {length}m RSB ({dia}) into {cuts} lengths.'
                else:
                    value = f'  Cut 1pc of {length}m RSB ({dia}) into {cuts} lengths.'
                cell = ws.cell(row=current_row, column=col_num, value=value)

            # Alternating BG Color Fill
            if current_row % 2 == 0:
                cell.fill = alter_row_fill
            if col_num == 5:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            else:
                cell.alignment = cell_alignment

            # Borders
            cell.border = border
    return wb

def add_sheet_cutting_list(title: str, rebar_config: list[dict[str, Any]],
                              market_lengths: dict[str, list], wb: Workbook) -> tuple[Workbook, bool]:
    """
    Generates a formatted Excel rebar cutting list from a rebar configuration dictionary.
    """
    ws = wb.create_sheet(f'{title} Cutting List')
    proceed_purchase_plan = True

    max_legs = 0
    if rebar_config:
        max_legs = max(len(bar['shape_dimensions']) for bar in rebar_config)

    # This list comprehension is creating the dictionary keys 'A', 'B', etc., which is fine.
    dimension_headers = [chr(ord('A') + i) for i in range(max_legs)]

    # --- Styles ---
    white_side = Side(style='thin', color='FFFFFF')
    black_side = Side(style='thin', color='404040')
    thick_black_side = Side(style='thick', color='404040')
    title_font = Font(name='Calibri', size=16, bold=True)
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='404040', end_color='404040', fill_type='solid')
    alter_row_fill = PatternFill(start_color='F3F3F3', end_color='404040', fill_type='solid')
    cell_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border = Border(left=black_side, right=black_side, top=black_side, bottom=black_side)
    header_border = Border(left=white_side, right=white_side, top=black_side, bottom=black_side)

    # --- Static and Dynamic Headers ---
    static_headers = ['Illustration', 'Bar Type', 'Diameter', 'Quantity', 'Cut Length']
    all_headers = static_headers + dimension_headers

    # --- Title ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(static_headers) + max_legs)
    title_cell = ws['A1']
    title_cell.value = f'{title} Rebar Cutting and Bending Schedule'
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # --- Headers ---
    for col_num, header_text in enumerate(all_headers, 1):
        cell = ws.cell(row=2, column=col_num, value=header_text)
        cell.font = header_font
        cell.alignment = cell_alignment
        cell.fill = header_fill
        cell.border = header_border

    # Apply black border to left and right outer edges
    cell = ws.cell(row=2, column=1)
    cell.border = Border(left=black_side, top=black_side, right=white_side, bottom=black_side)
    cell = ws.cell(row=2, column=len(all_headers))
    cell.border = Border(left=white_side, top=black_side, right=black_side, bottom=black_side)

    # --- Column Widths ---
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 18
    ws.row_dimensions[2].height = 25

    num_static_cols = len(static_headers)
    for i in range(max_legs):
        # Calculate column index: static columns + 1 (for 1-based index) + loop index
        col_idx = num_static_cols + 1 + i
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 12

    shape_to_image_map = {
        'straight': 'straight.png',
        'U': 'u.png',
        'L': 'l.png',
        'rectangular': 'rectangular_outer.png',
        'rectangular (tall)': 'rectangular_tall.png',
        'rectangular (wide)': 'rectangular_wide.png',
        'rectangular (diamond)': 'rectangular_diamond.png',
        'octagonal': 'octagon.png',
        'flat (tall)': 'flat_tall.png',
        'flat (wide)': 'flat_wide.png',
    }

    # --- Data Rows ---
    current_row = 3
    for bar in rebar_config:
        ws.row_dimensions[current_row].height = 75

        # Use the map to get the correct image filename
        image_filename = shape_to_image_map.get(bar['shape'])
        if image_filename:
            try:
                img_path = resource_path(f'images/{image_filename}')

                with PILImage.open(img_path) as pil_img:
                    original_width, original_height = pil_img.size

                aspect_ratio = original_height / original_width
                target_width = 90
                target_height = int(target_width * aspect_ratio + 0.5)
                if target_height > 90:
                    target_height = 90
                    target_width = int(90 / aspect_ratio + 0.5)

                img = Image(img_path)
                img.width = target_width
                img.height = target_height

                ws.add_image(center_img(img, f'A{current_row}', ws))

            except FileNotFoundError:
                ws.cell(row=current_row, column=1, value='No Image')
        else:
            ws.cell(row=current_row, column=1, value='No Image')

        try:
            val = bar['diameter']
            diameter_str = f'{val:.1f} mm\n({get_dia_code(val)})'
        except KeyError:
            # Fallback if the diameter code is not found
            diameter_str = f'{bar['diameter']:.1f} mm'

        data_to_write = [
            None,
            bar['bar_type'],
            diameter_str,  # Use the newly formatted string
            bar['quantity'],
            round(bar['cut_length'], 1),
        ]

        for letter in dimension_headers:
            data_to_write.append(bar['shape_dimensions'].get(letter, '-'))

        for col_num, value in enumerate(data_to_write, 1):
            if value is None:
                cell = ws.cell(row=current_row, column=col_num)
            else:
                cell = ws.cell(row=current_row, column=col_num, value=value)

            if col_num == 5:  # Cutlength
                dia_code = get_dia_code(bar['diameter'])
                # Check if market lengths are defined for this diameter
                if dia_code in market_lengths and market_lengths[dia_code]:
                    max_length = max(market_lengths[dia_code])
                    if value > max_length * 1000:
                        proceed_purchase_plan = False
                        cell.font = Font(color='FF0000')
                        cell.comment = Comment(
                            f'Splicing required.\nCutting length exceeds available market length of {max_length:}m.',
                            '✨rs_uy', height=150, width=200)
                else:
                    # This case handles when validation is bypassed or no lengths are available.
                    proceed_purchase_plan = False
                    cell.font = Font(color='FF0000')
                    cell.comment = Comment(
                        f'No market length selected for this diameter ({dia_code}).\nCannot proceed with purchase plan analysis.',
                        '✨rs_uy', height=150, width=200)

            # Alternating BG Color Fill
            if current_row % 2 == 0:
                cell.fill = alter_row_fill
            cell.alignment = cell_alignment

            # Borders
            cell.border = border
            if col_num == 6:  #Shape Dimensions
                cell.border = Border(left=thick_black_side, bottom=black_side, top=black_side, right=black_side)

            # Number Format
            if col_num >= 5:
                cell.number_format = '#,##0" mm"'

        current_row += 1

    return wb, proceed_purchase_plan

def add_concrete_plan_to_workbook(wb: Workbook, breakdown: list):
    """
    Adds a sheet with a concrete volume breakdown and an EDITABLE mix design section.
    Calculations for materials are done via Excel Formulas so the user can tweak the mix.
    """
    ws = wb.create_sheet('Concrete Purchase')

    # --- Styles ---
    white_side = Side(style='thin', color='FFFFFF')
    black_side = Side(style='thin', color='404040')
    title_font = Font(name='Calibri', size=16, bold=True)
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='404040', end_color='404040', fill_type='solid')
    alter_row_fill = PatternFill(start_color='F3F3F3', end_color='404040', fill_type='solid')
    header_border = Border(left=white_side, right=white_side, top=black_side, bottom=black_side)

    # Style for input cells to indicate they are editable (Light Yellow)
    input_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')

    center_align = Alignment(horizontal='center', vertical='center')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # =========================================================
    # SECTION 1: VOLUME BREAKDOWN
    # =========================================================
    ws.merge_cells('A1:D1')
    cell = ws['A1']
    cell.value = 'Concrete Volume Breakdown'
    cell.font = title_font
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    ws['A2'] = 'Foundation Type'
    ws['B2'] = 'Vol. Per Footing'
    ws['C2'] = 'Quantity'
    ws['D2'] = 'Vol. Per Type'
    ws.row_dimensions[2].height = 25
    for col in ['A', 'B', 'C', 'D']:
        c = ws[f'{col}2']
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = header_border

    # Apply black border to left and right outer edges
    cell = ws.cell(row=2, column=1)
    cell.border = Border(left=black_side, top=black_side, right=white_side, bottom=black_side)
    cell = ws.cell(row=2, column=4)
    cell.border = Border(left=white_side, top=black_side, right=black_side, bottom=black_side)

    row_idx = 3
    for name, vol, n_footing in breakdown:
        vol_per_footing = vol / n_footing
        ws[f'A{row_idx}'] = name
        ws[f'B{row_idx}'] = vol_per_footing
        ws[f'C{row_idx}'] = n_footing
        ws[f'D{row_idx}'] = f'=B{row_idx}*C{row_idx}'

        ws[f'B{row_idx}'].number_format = f'#,##0.00" m³"'
        ws[f'D{row_idx}'].number_format = f'#,##0.00" m³"'
        ws[f'A{row_idx}'].border = thin_border
        ws[f'B{row_idx}'].border = thin_border
        ws[f'C{row_idx}'].border = thin_border
        ws[f'D{row_idx}'].border = thin_border
        if row_idx %2==0:
            ws[f'A{row_idx}'].fill = alter_row_fill
            ws[f'B{row_idx}'].fill = alter_row_fill
            ws[f'C{row_idx}'].fill = alter_row_fill
            ws[f'D{row_idx}'].fill = alter_row_fill
        row_idx += 1

    # Total Row
    total_vol_row = row_idx
    ws.merge_cells(f'A{row_idx}:C{row_idx}')
    ws[f'A{total_vol_row}'] = 'Total Volume'
    ws[f'D{total_vol_row}'].value = f'=sum(D3:D{total_vol_row-1})'
    ws[f'A{total_vol_row}'].font = Font(bold=True)
    ws[f'D{total_vol_row}'].font = Font(bold=True)
    ws[f'A{total_vol_row}'].border = thin_border
    ws[f'B{total_vol_row}'].border = thin_border
    ws[f'C{total_vol_row}'].border = thin_border
    ws[f'D{total_vol_row}'].border = thin_border
    ws[f'A{total_vol_row}'].alignment = Alignment(horizontal='right')
    ws[f'D{total_vol_row}'].number_format = f'#,##0.00" m³"'

    # Store reference to total volume cell (e.g., 'B10')
    ref_total_vol = f'D{total_vol_row}'

    # =========================================================
    # SECTION 2: EDITABLE MIX PARAMETERS
    # =========================================================
    param_start_row = total_vol_row + 2

    ws.merge_cells(f'A{param_start_row}:B{param_start_row}')
    cell = ws[f'A{param_start_row}']
    cell.value = 'Mix Design Parameters'
    cell.font = title_font
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[param_start_row].height = 30

    # Parameter Labels and Default Values
    # (Label, Default Value, Key for reference)
    params = [
        ('Cement Ratio', 1, 'ratio_c'),
        ('Sand Ratio', 2, 'ratio_s'),
        ('Gravel Ratio', 3, 'ratio_g'),
        ('Water-Cement Ratio', 0.50, 'wc'),
        ('Bag Weight', 40, 'bag_wt'),
        ('Cement Density', 1440, 'cem_dens'),  # Usually hidden or static, but making it explicit is safer
        ('Dry Volume Factor', 1.54, 'factor'),
        ('Wastage Multiplier', 1.05, 'waste'),
    ]

    # Dictionary to store cell addresses for formulas
    refs = {}

    current_row = param_start_row + 1
    for label, val, key in params:
        # Label
        ws[f'A{current_row}'] = label
        ws[f'A{current_row}'].border = thin_border

        # Value (Input)
        c = ws[f'B{current_row}']
        c.value = val
        c.alignment = center_align
        c.border = thin_border
        c.fill = input_fill  # Highlight as editable
        if label == 'Bag Weight':
            c.number_format = '#,##0" kg"'
        elif label == 'Cement Density':
            c.number_format = f'#,##0" kg/m³"'
        # Save address (e.g., 'B15')
        refs[key] = f'B{current_row}'

        current_row += 1

    # =========================================================
    # SECTION 3: MATERIAL ESTIMATION (FORMULAS)
    # =========================================================
    est_row = current_row + 1
    ws.merge_cells(f'A{est_row}:B{est_row}')
    cell = ws[f'A{est_row}']
    cell.value = 'Materials for Purchase'
    cell.font = title_font
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[est_row].height = 30

    headers = ['Material', 'Quantity']
    ws.row_dimensions[est_row + 1].height = 25
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=est_row + 1, column=i)
        c.value = h
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = header_border

    # Apply black border to left and right outer edges
    cell = ws.cell(row=est_row + 1, column=1)
    cell.border = Border(left=black_side, top=black_side, right=white_side, bottom=black_side)
    cell = ws.cell(row=est_row + 1, column=2)
    cell.border = Border(left=white_side, top=black_side, right=black_side, bottom=black_side)

    # --- FORMULA CONSTRUCTION ---
    # 1. Total Ratio Sum = (Rc + Rs + Rg)
    sum_ratio = f'({refs['ratio_c']} + {refs['ratio_s']} + {refs['ratio_g']})'

    # 2. Total Dry Volume = TotalWet * Factor
    dry_vol = f'({ref_total_vol} * {refs['factor']})'

    # 3. Cement Calculation
    # Vol_Cement_Part = DryVol * (Rc / Sum)
    # Weight_Cement = Vol_Cement_Part * Density
    # Bags = Weight / BagWeight
    # Total = Bags * Wastage
    formula_cement = f'=(({dry_vol} * ({refs['ratio_c']} / {sum_ratio})) * {refs['cem_dens']} / {refs['bag_wt']}) * {refs['waste']}'

    # 4. Sand Calculation
    # Vol_Sand = DryVol * (Rs / Sum) * Wastage
    formula_sand = f'=({dry_vol} * ({refs['ratio_s']} / {sum_ratio})) * {refs['waste']}'

    # 5. Gravel Calculation
    # Vol_Gravel = DryVol * (Rg / Sum) * Wastage
    formula_gravel = f'=({dry_vol} * ({refs['ratio_g']} / {sum_ratio})) * {refs['waste']}'

    # 6. Water Calculation
    # Weight of Cement (Total) = (Bags * BagWeight) ... or derived from formula 3 without bag division
    # Let's derive it cleanly: Weight_Cement = Vol_Cement_Part * Density * Wastage
    # Water Weight = Weight_Cement * WC_Ratio
    # Water Vol (L) = Water Weight (1kg = 1L)
    # Formula:
    cement_weight_kg = f'((({dry_vol} * ({refs['ratio_c']} / {sum_ratio})) * {refs['cem_dens']}) * {refs['waste']})'
    formula_water = f'={cement_weight_kg} * {refs['wc']}'

    materials = [
        ('Cement', formula_cement, 'Bags', 'Based on bag weight & ratio'),
        ('Sand', formula_sand, 'm³', 'Loose Volume'),
        ('Gravel', formula_gravel, 'm³', 'Loose Volume'),
        ('Water', formula_water, 'Liters', 'Based on w/c ratio')
    ]

    r = est_row + 2
    for mat, formula, unit, note in materials:
        # Name
        ws.cell(row=r, column=1, value=mat).border = thin_border

        # Formula Cell
        c = ws.cell(row=r, column=2)
        c.value = formula  # Write the Excel formula string
        c.number_format = f'#,##0.00" {unit}"'
        c.alignment = center_align
        c.border = thin_border

        # Note
        # ws.cell(row=r, column=3, value=note).border = thin_border

        if r%2==0:
            for i in range(len(headers)):
                c = ws.cell(row=r, column=i + 1)
                c.fill = alter_row_fill
        r += 1

    # Add Footer Note
    cell = ws.cell(row=r, column=1, value='Note: Modify values in yellow cells to update quantities automatically.')
    cell.font = Font(name='Calibri', size=10, color='5D5D5D')

    # Column widths
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20

def delete_blank_worksheets(wb: Workbook) -> Workbook:
    """
    Deletes all blank worksheets from an Excel workbook.

    Args:
        wb: the Excel workbook.
    """
    sheets_to_delete = []

    for sheet in wb.worksheets:
        # Check if the sheet contains any data
        # max_row and max_column will be 1 if the sheet is empty
        if sheet.max_row == 1 and sheet.max_column == 1 and sheet.cell(row=1, column=1).value is None:
            sheets_to_delete.append(sheet)

    for sheet in sheets_to_delete:
        wb.remove(sheet)

    return wb