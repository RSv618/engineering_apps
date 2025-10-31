import sys
import os
import subprocess
from typing import Any
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QLabel, QLineEdit, QComboBox,
    QGroupBox, QGridLayout, QTextEdit, QFrame, QSizePolicy,
    QFormLayout, QCheckBox, QScrollArea, QMessageBox, QFileDialog,
    QSpinBox, QDoubleSpinBox, QPushButton
)
from PyQt6.QtGui import QPainter, QPen, QColor, QPaintEvent, QIcon
from PyQt6.QtCore import Qt, QPointF, QTimer, QEvent
from utils import (
    load_stylesheet, get_img, update_image,
    toggle_obj_visibility, parse_spacing_string, get_bar_dia,
    parse_nested_dict, get_dia_code, global_exception_hook, InfoPopup, HoverLabel,
    BlankSpinBox, resource_path, HoverButton, MemoryGroupBox, style_invalid_input
)
from rebar_calculations import (
    top_bottom_bar_calculation, perimeter_bar_calculation,
    vertical_bar_calculation, stirrups_calculation
)
from excel_writer import process_rebar_input, create_excel_cutting_list
from constants import (BAR_DIAMETERS, BAR_DIAMETERS_FOR_STIRRUPS,
                       MARKET_LENGTHS, FOOTING_IMAGE_WIDTH, STIRRUP_ROW_IMAGE_WIDTH,
                       RSB_IMAGE_WIDTH, DEBUG_MODE)
from functools import partial

"""
TO BUILD:
pyinstaller --name "CuttingList" --onefile --windowed --icon="images/logo.png" --add-data "images:images" --add-data "style.qss:." app_cutting_list.py
"""
class DrawStirrup(QWidget):
    def __init__(self, width: int, parent: QWidget | None = None) -> None:
        """
        Initializes the DrawStirrup widget for visualizing stirrup layouts.

        Args:
            width: The base width of the canvas.
            parent: The parent widget, if any.
        """
        super().__init__(parent)
        width += 20
        self.setFixedWidth(width)
        self.setMaximumHeight(int(1.6 * width))

        # Initialize with default values
        self.ped_h = 1000
        self.ped_bx = 700
        self.pad_t = 300
        self.cc = 75
        self.extent = 'From Face of Pad'
        self.spacing = []
        self.bot_bar_diameter = 10
        self.vert_bar_diameter = 16
        self.stirrup_qty = 0

    def update_dimensions(self, footing_details, extent, spacing, bot_bar_diameter, vert_bar_diameter):
        """Updates the drawing dimensions from the input widgets and triggers a repaint."""
        self.ped_h = footing_details['Pedestal Height'].value()
        self.ped_bx = footing_details['Pedestal Width (Along X)'].value()
        self.pad_t = footing_details['Pad Thickness'].value()
        self.cc = footing_details['Concrete Cover'].value()
        bot_bar_diameter_str = bot_bar_diameter.currentText()
        vert_bar_diameter_str = vert_bar_diameter.currentText()
        self.bot_bar_diameter = get_bar_dia(bot_bar_diameter_str.strip(), 'ph')
        self.vert_bar_diameter = get_bar_dia(vert_bar_diameter_str.strip(), 'ph')
        self.extent = extent.currentText()

        if len(spacing.toPlainText()) == 0:
            self.spacing = []
        else:
            try:
                self.spacing = parse_spacing_string(spacing.toPlainText())
            except (TypeError, ValueError):
                self.spacing = []

        self.update()  # Crucial: schedules a repaint which calls paintEvent

    def paintEvent(self, event: QPaintEvent) -> None:
        """
        Handles the repaint event to draw the footing and stirrups on the widget.

        Args:
            event: The paint event.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)  # Makes the lines smooth

        # Get the dimensions of the widget
        padding = 10
        c_width = self.width()
        c_height = self.height()

        real_h = self.ped_h
        real_bx = self.ped_bx
        real_t = self.pad_t
        real_cc = self.cc

        real_height = real_h + real_t
        if real_height == 0 or real_bx == 0:
            painter.end()  # Ensure the painter is properly closed
            return  # Stop drawing
        px_height = c_height - 2 * padding
        px_bx = (c_width - 2 * padding) / 2
        scale = px_height / real_height
        scale_x = px_bx / real_bx
        px_t = real_t * scale
        px_h = real_h * scale
        px_cc = real_cc * scale

        # Define the pen for drawing the lines
        light_dark_pen = QPen(QColor('#666666'), 0.5)
        top_bottom_bar_pen = QPen(QColor('#9F9F9F9F'), 1.5)
        vert_bar_pen = QPen(QColor('#999999'), 2)
        stirrups_pen = QPen(QColor('#FF3333'), 2)
        painter.setPen(light_dark_pen)

        # Define the three points of the triangle
        x1 = padding
        x2 = px_bx/2 + x1
        x3 = px_bx + x2
        x4 = px_bx/2 + x3
        y1 = (c_height - padding)
        y2 = y1 - px_t
        y3 = y2 - px_h

        p1 = QPointF(x1, y1)
        p2 = QPointF(x1, y2)
        p3 = QPointF(x2, y2)
        p4 = QPointF(x2, y3)
        p5 = QPointF(x3, y3)
        p6 = QPointF(x3, y2)
        p7 = QPointF(x4, y2)
        p8 = QPointF(x4, y1)
        cc_y = QPointF(0, px_cc)

        # Draw the lines connecting the points
        # painter.drawLine(p1, p2)
        painter.drawLine(p2, p3)
        painter.drawLine(p3, p4)
        painter.drawLine(p4, p5)
        painter.drawLine(p5, p6)
        painter.drawLine(p6, p7)
        # painter.drawLine(p7, p8)
        painter.drawLine(p8, p1)

        # Draw Top Bottom Bar
        painter.setPen(top_bottom_bar_pen)
        painter.drawLine(p1 - cc_y, p8 - cc_y)
        painter.drawLine(p2 + cc_y, p7 + cc_y)

        # Draw Vertical Bar
        vbar_x1 = x1 + real_cc * scale_x
        vbar_x2 = x2 + real_cc * scale_x
        vbar_x3 = x3 - real_cc * scale_x
        vbar_x4 = x4 - real_cc * scale_x
        vbar_y1 = y1 - px_cc - 2.5
        vbar_y2 = y3 + px_cc

        painter.setPen(vert_bar_pen)
        painter.drawLine(QPointF(vbar_x1, vbar_y1), QPointF(vbar_x2, vbar_y1))
        painter.drawLine(QPointF(vbar_x2, vbar_y1), QPointF(vbar_x2, vbar_y2))
        painter.drawLine(QPointF(vbar_x3, vbar_y1), QPointF(vbar_x4, vbar_y1))
        painter.drawLine(QPointF(vbar_x3, vbar_y1), QPointF(vbar_x3, vbar_y2))

        painter.setPen(stirrups_pen)

        if self.extent == 'From Face of Pad':
            start_y = y2
            target_y = vbar_y2
        elif self.extent == 'From Bottom Bar':
            start_y = (y1 - px_cc - scale * (2*self.bot_bar_diameter + self.vert_bar_diameter))
            target_y = vbar_y2
        else:  # From Top
            start_y = vbar_y2
            target_y = y2
        actual_count = 0

        if self.extent in ['From Face of Pad', 'From Bottom Bar']:
            lines, count, last_y = self.loop_stirrup(self.spacing, start_y=start_y, target_y=target_y,
                                                     left_x=vbar_x2, right_x=vbar_x3, scale=scale)
            for p1, p2 in lines:
                painter.drawLine(p1, p2)
            actual_count += count

            # Add Topmost Stirrup if remaining gap >= concrete cover
            if (actual_count > 0) and (last_y - vbar_y2 >= px_cc):
                painter.drawLine(QPointF(vbar_x2, vbar_y2), QPointF(vbar_x3, vbar_y2))
                actual_count += 1

        else:  # From Top
            lines, count, last_y = self.loop_stirrup(self.spacing, start_y=start_y, target_y=target_y,
                                                     left_x=vbar_x2, right_x=vbar_x3, scale=scale)
            for p1, p2 in lines:
                painter.drawLine(p1, p2)
            actual_count += count

        self.stirrup_qty = actual_count
        painter.end()

    @staticmethod
    def loop_stirrup(spacing_list: list[tuple[int | str, float]], start_y: float, target_y: float, left_x: float,
                     right_x: float, scale: float) -> tuple[list[tuple[QPointF, QPointF]], int, float]:
        """
        Calculates the line coordinates for stirrups based on a spacing list.

        Args:
            spacing_list: A list of (quantity, spacing) tuples.
            start_y: The starting vertical coordinate for drawing.
            target_y: The ending vertical coordinate.
            left_x: The starting horizontal coordinate.
            right_x: The ending horizontal coordinate.
            scale: The drawing scale factor.

        Returns:
            A tuple containing the list of lines to draw, the total count of stirrups,
            and the last y-coordinate drawn.
        """
        count = 0
        current_y = start_y
        lines = []
        for qty, spacing in spacing_list:
            spacing = spacing * scale
            if isinstance(qty, int):
                for _ in range(qty):
                    if target_y < start_y:
                        current_y -= spacing
                        if current_y < target_y:
                            break
                    else:
                        current_y += spacing
                        if current_y > target_y:
                            break
                    lines.append((QPointF(left_x, current_y), QPointF(right_x, current_y)))
                    count += 1
            elif qty == 'rest':
                if target_y < start_y:
                    while current_y - spacing >= target_y:
                        current_y -= spacing
                        lines.append((QPointF(left_x, current_y), QPointF(right_x, current_y)))
                        count += 1
                else:
                    while current_y + spacing <= target_y:
                        current_y += spacing
                        lines.append((QPointF(left_x, current_y), QPointF(right_x, current_y)))
                        count += 1
        return lines, count, current_y

    def get_qty(self) -> int:
        """
        Returns the calculated quantity of stirrups from the last paint event.

        Returns:
            The total number of stirrups.
        """
        return self.stirrup_qty


class MultiPageApp(QMainWindow):
    def __init__(self) -> None:
        """Initializes the main application window and its components."""
        super().__init__()

        self.setWindowTitle('Cutting List')
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
        self.setGeometry(50, 50, 980, 720)
        self.setMinimumWidth(980)
        self.setMinimumHeight(600)

        self.debounce_timer = QTimer(self)
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.setInterval(500)  # 500 ms delay

        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        self.footing_details = {}
        self.rsb_details = {}
        self.market_lengths_checkboxes = {}
        self.summary_labels = {}
        self.remove_stirrup_button = None
        self.stirrup_rows_layout = None
        self.stirrup_canvas = None
        self.group_box = {}
        self.footing_details_values = {}
        self.rsb_details_values = {}
        self.info_popup = InfoPopup(self)

        # Mousewheel protection
        QApplication.instance().installEventFilter(self)

        self.create_footing_page()
        self.create_rsb_page()
        self.create_market_lengths_page()
        self.create_summary_page()
        self.setup_connections()

        if DEBUG_MODE:
            self.prefill_for_debug()
        self.stacked_widget.setCurrentIndex(0)

    def create_footing_page(self) -> None:
        """Builds the UI for the first page (Footing Dimensions)."""
        page = QWidget()
        page.setProperty('class', 'page')
        page_layout = QVBoxLayout(page)
        content_layout = QHBoxLayout()

        # Left side: Image
        footing_img = get_img(resource_path('images/label_1ped.png'), FOOTING_IMAGE_WIDTH, FOOTING_IMAGE_WIDTH)
        content_layout.addWidget(footing_img)

        # Right side: Form
        rows = {
            # Description  Variable  Widget
            'Footing Type': (None, QComboBox()),
            'Total Number of Footing': (None, BlankSpinBox(1, 9_999, 1)),
            'Concrete Cover': (None, BlankSpinBox(1, 999, 75,suffix=' mm')),
            'Pedestal Width (Along X)': ('bx', BlankSpinBox(0, 99_999, suffix=' mm')),
            'Pedestal Width (Along Y)': ('by', BlankSpinBox(0, 99_999, suffix=' mm')),
            'Pedestal Height': ('h', BlankSpinBox(0, 999_999, suffix=' mm')),
            'Pad Width (Along X)': ('Bx', BlankSpinBox(0, 999_999, suffix=' mm')),
            'Pad Width (Along Y)': ('By', BlankSpinBox(0, 999_999, suffix=' mm')),
            'Pad Thickness': ('t', BlankSpinBox(0, 99_999, suffix=' mm')),
        }

        # --- Create the form widget and the QGridLayout ---
        form_widget = QWidget()
        form_layout = QGridLayout(form_widget)
        form_layout.setColumnMinimumWidth(1, 50)

        # Rebuild the inputs_page1 dict from the new rows structure
        self.footing_details = {label: widget for label, (_, widget) in rows.items()}

        # Set defaults and connect signals
        n_ped = {'Isolated': 1, 'Mat (2 Ped)': 2, 'Mat (3 Ped)': 3, 'Mat (4 Ped)': 4}
        image_map = {
            'Isolated': resource_path('images/label_1ped.png'),
            'Mat (2 Ped)': resource_path('images/label_2ped.png'),
            'Mat (3 Ped)': resource_path('images/label_3ped.png'),
            'Mat (4 Ped)': resource_path('images/label_4ped.png')
        }
        self.footing_details['Footing Type'].addItems(n_ped.keys())
        self.footing_details['n_ped'] = n_ped['Isolated']
        self.footing_details['Footing Type'].currentTextChanged.connect(
            lambda selected_text: update_image(selected_text, image_map, footing_img,
                                               fallback=resource_path('images/label_0ped.png')))
        self.footing_details['Footing Type'].currentTextChanged.connect(
            lambda selected_text: self.footing_details.update({'n_ped': n_ped[selected_text]}))

        # --- Build the form row by row using the grid layout ---
        for row_index, (label_text, (variable, widget)) in enumerate(rows.items()):

            # Column 0: Description Label
            description_label = QLabel(label_text)

            # Column 1: Red Variable Label (if it exists)
            if variable is not None:
                form_layout.addWidget(description_label, row_index, 0)
                label = QLabel(variable)
                label.setProperty('class', 'footing-variable')
                form_layout.addWidget(label, row_index, 1)
            else:
                form_layout.addWidget(description_label, row_index, 0, 1, 2)

            if 'Along Y' in label_text:
                # Add the 'Along Y' spinbox first
                form_layout.addWidget(widget, row_index, 2)

                # Create and configure the checkbox
                lock_ratio_checkbox = QCheckBox()
                lock_ratio_checkbox.setProperty('class', 'lock-ratio')
                lock_ratio_checkbox.setChecked(True)
                lock_ratio_checkbox.setCursor(Qt.CursorShape.PointingHandCursor)
                lock_ratio_checkbox.setToolTip('Square [Lock Ratio]')
                form_layout.addWidget(lock_ratio_checkbox, row_index, 3)

                # 1. Identify the 'Along Y' spinbox (the current 'widget') and its
                #    corresponding 'Along X' spinbox from our dictionary.
                spinbox_y = widget
                x_label_text = label_text.replace('Along Y', 'Along X')
                spinbox_x = self.footing_details[x_label_text]

                # 2. Connect the checkbox's 'toggled' signal to a lambda function.
                #    This function will:
                #    a) Disable the 'Y' spinbox when the box is checked.
                #    b) If checked, immediately copy the 'X' value to the 'Y' spinbox.
                #    NOTE: We use default arguments (sb_y=spinbox_y) to "capture" the
                #    correct spinbox for this loop iteration.
                # noinspection PyUnresolvedReferences
                lock_ratio_checkbox.toggled.connect(
                    lambda checked, sb_y=spinbox_y, sb_x=spinbox_x: (
                        sb_y.setEnabled(not checked),
                        sb_y.setValue(sb_x.value()) if checked else None
                    )
                )

                # 3. Connect the 'X' spinbox's 'valueChanged' signal to another lambda.
                #    This function will update the 'Y' spinbox's value in real-time,
                #    but only if the lock checkbox is currently checked.
                spinbox_x.valueChanged.connect(
                    lambda value, sb_y=spinbox_y, chk=lock_ratio_checkbox: (
                        sb_y.setValue(value) if chk.isChecked() else None
                    )
                )

                # 4. Set the initial state when the UI is first created.
                #    Since setChecked(True) is the default, we disable the 'Y' spinbox
                #    and sync its value with the 'X' spinbox's starting value.
                spinbox_y.setEnabled(False)
                spinbox_y.setValue(spinbox_x.value())

            else:
                form_layout.addWidget(widget, row_index, 2, 1, 2)

        # Allow the input column (2) to stretch, keeping other columns fixed
        form_layout.setColumnStretch(2, 1)

        # Container to vertically center the form
        right_side_widget = QWidget()
        right_side_layout = QVBoxLayout(right_side_widget)
        right_side_layout.addStretch(1)
        label = QLabel('Footing Dimensions')
        label.setProperty('class', 'header-0')
        right_side_layout.addWidget(label)
        right_side_layout.addWidget(form_widget)
        right_side_layout.addStretch(1)

        content_layout.addWidget(right_side_widget)
        page_layout.addLayout(content_layout)

        # Bottom: Navigation buttons (Your existing code is perfect)
        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        next_button = HoverButton('Next')
        next_button.setProperty('class', 'green-button')
        next_button.clicked.connect(self.go_to_rsb_page)
        button_layout.addWidget(next_button)
        page_layout.addLayout(button_layout)
        self.stacked_widget.addWidget(self.make_scrollable(page))

    def create_rsb_page(self) -> None:
        """Builds the UI for the second page (Reinforcement Details)."""
        page = QWidget()
        page.setProperty('class', 'page')

        page_layout = QVBoxLayout(page)  # Change this from QGridLayout to QVBoxLayout

        # Create a container widget for the scrollable content
        scroll_content = QWidget()
        scroll_content.setProperty('class', 'scroll-area')
        grid_layout = QGridLayout(scroll_content)  # This will now hold your GroupBoxes
        grid_layout.setColumnStretch(0, 1)
        grid_layout.setColumnStretch(1, 1)

        # --- A helper function to create each group box to avoid repeating code ---
        def create_top_bot_bar_section(title, image_path, image_width):
            if 'Top Bar' in title:
                group_box = MemoryGroupBox(title)
            else:
                group_box = QGroupBox(title)
            section_layout = QHBoxLayout(group_box)

            section_layout.addWidget(get_img(image_path, image_width, image_width))

            grid_top_bottom = QGridLayout()
            grid_top_bottom.setColumnStretch(0, 1)
            grid_top_bottom.setColumnStretch(1, 1)
            grid_top_bottom.setColumnStretch(2, 1)
            grid_top_bottom.setColumnStretch(3, 1)

            # Row 0: Diameter
            label = QLabel('Diameter:')
            label.setProperty('class', 'rsb-forms-label')
            grid_top_bottom.addWidget(label, 0, 0)
            bar_size = QComboBox()
            bar_size.addItems(BAR_DIAMETERS)
            grid_top_bottom.addWidget(bar_size, 0, 1, 1, 3)

            # Row 1: ComboBox
            input_type = QComboBox()
            input_type.addItems(['Quantity', 'Spacing'])
            grid_top_bottom.addWidget(input_type, 1, 0, 1, 4)

            # Row 2: Inputs
            value_along_x = BlankSpinBox(0, 99_999)
            value_along_x.setMinimumWidth(100)
            grid_top_bottom.addWidget(value_along_x, 2, 0, 1, 2)
            value_along_y = BlankSpinBox(0, 99_999)
            value_along_y.setMinimumWidth(100)
            grid_top_bottom.addWidget(value_along_y, 2, 2, 1, 2)

            # Row 3: Along X / Y
            along_x_label = QLabel('Along X')
            along_x_label.setProperty('class', 'rsb-forms-label along')
            grid_top_bottom.addWidget(along_x_label, 3, 0, 1, 2)
            along_y_label = QLabel('Along Y')
            along_y_label.setProperty('class', 'rsb-forms-label along')
            grid_top_bottom.addWidget(along_y_label, 3, 2, 1, 1)
            h_layout = QHBoxLayout()
            same_for_both = QCheckBox()
            same_for_both.setToolTip('Same for both directions')
            same_for_both.setProperty('class', 'lock-ratio')
            same_for_both.setCursor(Qt.CursorShape.PointingHandCursor)
            h_layout.addStretch()
            h_layout.addWidget(same_for_both)
            grid_top_bottom.addLayout(h_layout, 3, 3)


            v_layout = QVBoxLayout()
            v_layout.addLayout(grid_top_bottom)
            v_layout.addStretch()
            section_layout.addLayout(v_layout)

            # --- Store controls for later data retrieval and manipulation ---
            self.rsb_details[title] = {
                'Diameter': bar_size,
                'Input Type': input_type,
                'Value Along X': value_along_x,
                'Value Along Y': value_along_y,
                'connection': None
            }

            # --- Connections for dynamic UI changes ---
            # 1. Change spinbox suffix based on input type selection
            # noinspection PyUnresolvedReferences
            input_type.currentTextChanged.connect(
                partial(self.update_spinbox_suffixes, title)
            )
            # 2. Handle the "Same for both" checkbox state
            # noinspection PyUnresolvedReferences
            same_for_both.stateChanged.connect(
                partial(self.toggle_same_for_both, title)
            )

            # --- Set initial UI state correctly ---
            self.update_spinbox_suffixes(title, input_type.currentText())
            self.toggle_same_for_both(title, same_for_both.checkState())
            same_for_both.setChecked(True)

            self.group_box[title] = group_box
            return group_box

        def create_vert_bar_section(image_width):
            title = 'Vertical Bar'
            group_box = QGroupBox(title)
            section_layout = QHBoxLayout(group_box)

            # --- Image (Left side) ---
            section_layout.addWidget(get_img(resource_path('images/vert_bar.png'), image_width, image_width))

            # --- Container for the right side controls ---
            form_layout = QFormLayout()

            # Row 0: Diameter
            bar_size = QComboBox()
            bar_size.addItems(BAR_DIAMETERS)
            size_policy = bar_size.sizePolicy()
            size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
            bar_size.setSizePolicy(size_policy)
            label = QLabel('Diameter:')
            label.setProperty('class', 'rsb-forms-label')
            form_layout.addRow(label, bar_size)

            # Row 1: Qty
            qty = BlankSpinBox(0, 99_999, suffix=' pcs')
            label = QLabel('Qty:')
            label.setProperty('class', 'rsb-forms-label')
            form_layout.addRow(label, qty)

            # Row 2: Hook Calculation
            calculation = QComboBox()
            calculation.addItems(['Automatic', 'Manual'])
            label = HoverLabel('Hook Calculation:')
            label.setProperty('class', 'rsb-forms-label')
            label.mouseEntered.connect(self.show_hook_info)
            label.mouseLeft.connect(self.info_popup.hide)
            form_layout.addRow(label, calculation)

            # Row 3: Hook Length (Label)
            hook_length_label = QLabel('Hook Length:')
            hook_length_label.setProperty('class', 'rsb-forms-label')
            hook_length = BlankSpinBox(0, 99_999, suffix=' mm')
            form_layout.addRow(hook_length_label, hook_length)

            # Connect the combo box's signal to our new function
            hook_length_label.setVisible(False)
            hook_length.setVisible(False)
            # noinspection PyUnresolvedReferences
            calculation.currentTextChanged.connect(
                lambda selected_text: toggle_obj_visibility(selected_text, 'Manual', [hook_length_label, hook_length])
            )

            # Add the right side to the main layout
            section_layout.addLayout(form_layout)

            # Store the controls for later data retrieval
            self.rsb_details[title] = {
                'Diameter': bar_size,
                'Quantity': qty,
                'Hook Calculation': calculation,
                'Hook Length': hook_length
            }

            self.group_box[title] = group_box
            return group_box

        def create_perim_bar_section(image_width):
            title = 'Perimeter Bar'
            group_box = MemoryGroupBox(title)
            section_layout = QHBoxLayout(group_box)

            # --- Image (Left side) ---
            image_map = {#'None': resource_path('images/perim_bar_0.png'),
                         '1': resource_path('images/perim_bar_1.png'),
                         '2': resource_path('images/perim_bar_2.png'),
                         '3': resource_path('images/perim_bar_3.png')}
            perim_bar_img = get_img(image_map['1'], image_width, image_width)
            section_layout.addWidget(perim_bar_img)

            # --- Container for the right side controls ---
            form_layout = QFormLayout()

            # Row 0: Layers
            layers = QComboBox()
            layers.addItems(['1', '2', '3'])  # Add None if needed
            size_policy = layers.sizePolicy()
            size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
            layers.setSizePolicy(size_policy)
            layers_label = QLabel('Layers:')
            layers_label.setProperty('class', 'rsb-forms-label rsb-layers-label')
            form_layout.addRow(layers_label, layers)
            bar_size = QComboBox()
            bar_size.addItems(BAR_DIAMETERS)
            diameter_label = QLabel('Diameter:')
            diameter_label.setProperty('class', 'rsb-forms-label')
            form_layout.addRow(diameter_label, bar_size)

            # --- Add the right side to the main layout ---
            section_layout.addLayout(form_layout)
            self.rsb_details[title] = {'Diameter': bar_size,
                                       'Layers': layers}

            # diameter_label.setVisible(False)
            # bar_size.setVisible(False)
            # noinspection PyUnresolvedReferences
            layers.currentTextChanged.connect(
                lambda selected_text: update_image(selected_text, image_map, perim_bar_img, image_width,
                                                   fallback=resource_path('images/perim_bar_0.png')))

            # if none is in combobox
            # layers.currentTextChanged.connect(
            #     lambda selected_text: toggle_obj_visibility(selected_text, 'None',
            #                                                 [diameter_label, bar_size], True)
            # )

            group_box.setChecked(False)
            self.group_box[title] = group_box
            return group_box

        def create_stirrup_group_box(image_width):
            group_box = MemoryGroupBox('Stirrups')
            main_layout = QHBoxLayout(group_box)
            left_section = QVBoxLayout()
            left_section.setContentsMargins(0, 0, 10, 0)

            # --- Button Layout for adding/removing rows ---
            add_remove_layout = QHBoxLayout()
            label = HoverLabel('Bundle of Shapes')
            label.setProperty('class', 'stirrup-header')
            label.mouseEntered.connect(self.show_bundle_info)
            label.mouseLeft.connect(self.info_popup.hide)
            add_button = HoverButton('+')
            add_button.setProperty('class', 'green-button add-row-button')
            self.remove_stirrup_button = HoverButton('-')
            self.remove_stirrup_button.setProperty('class', 'red-button remove-row-button')

            add_button.clicked.connect(self.add_stirrup_row)
            self.remove_stirrup_button.clicked.connect(self.remove_stirrup_row)

            add_remove_layout.addWidget(label)
            add_remove_layout.addStretch()
            add_remove_layout.addWidget(add_button)
            add_remove_layout.addWidget(self.remove_stirrup_button)
            left_section.addLayout(add_remove_layout)

            # --- Container for dynamic stirrup rows ---
            stirrup_rows_container = QWidget()
            self.stirrup_rows_layout = QVBoxLayout(stirrup_rows_container)
            self.stirrup_rows_layout.setContentsMargins(0, 0, 0, 0)
            self.stirrup_rows_layout.setSpacing(5)

            left_section.addWidget(stirrup_rows_container)

            left_section.addStretch(1)  # Pushes rows to the top

            # --- Add the first, initial row ---
            self.rsb_details['Stirrups'] = {'Types': []}
            self.add_stirrup_row()


            # -----------------------------------------RIGHT SECTION
            right_section = QHBoxLayout()

            # --- Image (Left side) ---
            canvas_container = QVBoxLayout()
            canvas_container.setContentsMargins(10,0,0,0)
            label = HoverLabel('Spacing Per Bundle')
            label.setProperty('class', 'stirrup-header')
            label.mouseEntered.connect(self.show_spacing_header_info)
            label.mouseLeft.connect(self.info_popup.hide)
            canvas_container.addWidget(label)
            self.stirrup_canvas = DrawStirrup(image_width)
            self.stirrup_canvas.setProperty('class', 'stirrup-canvas')
            canvas_container.addWidget(self.stirrup_canvas)
            canvas_container.addStretch()
            right_section.addLayout(canvas_container)

            # --- Container for the right side controls ---
            form_layout = QFormLayout()

            # Row 0: Extent
            extent = QComboBox()
            extent.addItems(['From Face of Pad', 'From Bottom Bar', 'From Top'])
            extent_label = HoverLabel('Start From:')
            extent_label.setProperty('class','rsb-forms-label')
            extent_label.mouseEntered.connect(self.show_spacing_extent_info)
            extent_label.mouseLeft.connect(self.info_popup.hide)
            form_layout.addRow(extent_label, extent)

            # Row 1: Spacing
            spacing = QTextEdit()
            spacing.setProperty('class', 'rsb-spacing-text-edit')
            spacing.setPlaceholderText('Example: 1@50, 5@80, rest@100')
            spacing_label = HoverLabel('Spacing:') # Use HoverLabel
            spacing_label.setProperty('class','rsb-forms-label')

            # Connect its hover signals
            spacing_label.mouseEntered.connect(self.show_spacing_info)
            spacing_label.mouseLeft.connect(self.info_popup.hide)

            form_layout.addRow(spacing_label, spacing)

            vert_layout = QVBoxLayout()
            vert_layout.addLayout(form_layout)
            vert_layout.addStretch(1)
            right_section.addLayout(vert_layout)

            # Store
            self.rsb_details['Stirrups']['Extent'] = extent
            self.rsb_details['Stirrups']['Spacing'] = spacing


            # -----------------------------------------COMBINE SECTION
            main_layout.addLayout(left_section, 1)
            separator = QFrame()
            separator.setFrameShape(QFrame.Shape.VLine)
            separator.setProperty('class', 'separator')
            # separator.setFrameShadow(QFrame.Shadow.Sunken)  # Optional: adds a 3D effect
            main_layout.addWidget(separator)
            main_layout.addLayout(right_section, 1)

            self.group_box['Stirrups'] = group_box
            return group_box

        # --- Create and add the group boxes ---
        width = RSB_IMAGE_WIDTH
        top_bar_box = create_top_bot_bar_section('Top Bar', resource_path('images/top_bar.png'), width)
        bot_bar_box = create_top_bot_bar_section('Bottom Bar', resource_path('images/bot_bar.png'), width)
        vert_bar_box = create_vert_bar_section(width)
        perim_bar_box = create_perim_bar_section(width)
        stirrup_group_box = create_stirrup_group_box(width)


        grid_layout.addWidget(top_bar_box, 0, 0)
        grid_layout.addWidget(bot_bar_box, 0, 1)
        grid_layout.addWidget(vert_bar_box, 1, 0)
        grid_layout.addWidget(perim_bar_box, 1, 1)
        grid_layout.addWidget(stirrup_group_box, 2, 0, 1, 2)
        grid_layout.setRowStretch(2, 1)

        # --- Navigation Buttons ---
        button_layout = QHBoxLayout()
        back_button = HoverButton('Back')
        back_button.setProperty('class', 'red-button')
        back_button.clicked.connect(self.go_to_footing_page)
        next_button = HoverButton('Next')
        next_button.setProperty('class', 'green-button')
        next_button.clicked.connect(self.go_to_market_lengths_page)

        button_layout.addWidget(back_button)
        button_layout.addStretch()
        button_layout.addWidget(next_button)

        scroll_area = self.make_scrollable(scroll_content, True)
        scroll_area.setProperty('class', 'scroll-bar-area')
        page_layout.addWidget(scroll_area)  # Add the scrollable part
        page_layout.addLayout(button_layout)  # Add the fixed buttons at the bottom

        self.stacked_widget.addWidget(page)
        # self.stacked_widget.addWidget(self.make_scrollable(page, True))

    def create_market_lengths_page(self) -> None:
        """Builds the UI for the third page (Rebar Market Lengths)."""
        page = QWidget()
        page.setProperty('class', 'page')
        main_layout = QVBoxLayout(page)

        container = QFrame()
        grid = QGridLayout(container)
        grid.setSpacing(0)

        # --- Helper function to create a styled cell ---
        def create_cell(widget, is_header=False, is_alternate=False):
            cell = QWidget()
            cell.setAutoFillBackground(True)

            cell_layout = QHBoxLayout(cell)
            cell_layout.setContentsMargins(0, 0, 0, 0)
            cell_layout.setSpacing(0)
            if isinstance(widget, QPushButton):
                cell_layout.addWidget(widget)
            else:
                cell_layout.addStretch(1)
                cell_layout.addWidget(widget)
                cell_layout.addStretch(1)


            style_class = 'grid-cell'
            if is_header:
                style_class += ' header-cell'
            if is_alternate:
                style_class += ' alternate-row-cell'
            cell.setProperty('class', style_class)
            return cell

        # --- Top-Left Header ('Diameter') ---
        dia_header = QLabel('Diameter')
        dia_header.setProperty('class', 'market-header-label')
        grid.addWidget(create_cell(dia_header, is_header=True), 0, 0)

        # --- Column Headers (Lengths) ---
        for col, length in enumerate(MARKET_LENGTHS):
            btn = HoverButton(length)
            btn.setProperty('class', 'clickable-header clickable-column-header')
            btn.clicked.connect(lambda checked, l=length: self.toggle_market_column(l))
            # No width passed, so it uses the default of 65
            grid.addWidget(create_cell(btn, is_header=True), 0, col + 1)

        # --- Rows (Diameters and Checkboxes) ---
        self.market_lengths_checkboxes = {}
        for row, dia in enumerate(BAR_DIAMETERS):
            is_alternate_row = row % 2 == 1
            self.market_lengths_checkboxes[dia] = {}

            # Row Header (Diameter)
            btn = HoverButton(dia)
            btn.setProperty('class', 'clickable-header clickable-row-header')
            btn.clicked.connect(lambda checked, d=dia: self.toggle_market_row(d))
            grid.addWidget(create_cell(btn, is_header=True, is_alternate=is_alternate_row), row + 1, 0)

            # Checkboxes for each length
            for col, length in enumerate(MARKET_LENGTHS):
                cb = QCheckBox()

                cb.setChecked(True)
                self.market_lengths_checkboxes[dia][length] = cb
                # No width passed here, uses the default 65 for the cell
                grid.addWidget(create_cell(cb, is_alternate=is_alternate_row), row + 1, col + 1)

        # Horizontal layout to center the grid and keep it from stretching
        h_layout = QHBoxLayout()
        h_layout.addStretch(1)
        h_layout.addWidget(container)
        h_layout.addStretch(1)

        # Add title and the centering layout to the main page layout
        main_layout.addStretch(1)
        label = QLabel('Rebar Market Lengths')
        label.setProperty('class', 'header-0')
        main_layout.addWidget(label)
        main_layout.addLayout(h_layout)
        main_layout.addStretch(1)

        # --- Navigation Buttons ---
        button_layout = QHBoxLayout()
        back_button = HoverButton('Back')
        back_button.setAutoDefault(True)
        back_button.setProperty('class', 'red-button')
        back_button.clicked.connect(self.go_to_rsb_page)

        next_button = HoverButton('Next')
        next_button.setAutoDefault(True)
        next_button.setProperty('class', 'green-button')
        next_button.clicked.connect(self.go_to_summary_page)

        button_layout.addWidget(back_button)
        button_layout.addStretch()
        button_layout.addWidget(next_button)

        main_layout.addLayout(button_layout)
        self.stacked_widget.addWidget(page)

    def create_summary_page(self) -> None:
        """Builds the UI for the fourth page (Summary and Generation)."""
        page = QWidget()
        page.setProperty('class', 'page')
        main_layout = QVBoxLayout(page)

        # --- Helper to create a styled section ---
        def create_summary_section(title):
            group_box = QGroupBox(title)
            layout = QFormLayout(group_box)
            layout.setRowWrapPolicy(QFormLayout.RowWrapPolicy.DontWrapRows)  # Prevents ugly wrapping
            return group_box, layout

        # Create the main sections
        footing_box, footing_layout = create_summary_section('Footing Dimensions')
        reinf_box, reinf_layout = create_summary_section('Reinforcement Details')
        market_box, market_layout = create_summary_section('Market Lengths')

        # --- Populate sections with labels and store the value labels for updating ---
        # Footing Section
        for field in self.footing_details.keys():
            if field == 'n_ped':
                continue
            value_label = QLabel('...')
            value_label.setProperty('class', 'summary-value')
            self.summary_labels[field] = value_label
            label = QLabel(f'{field}:')
            label.setProperty('class', 'summary-label')
            footing_layout.addRow(label, value_label)

        # Reinforcement Section
        for field in self.rsb_details.keys():
            if field == 'Stirrups':
                label = QLabel('Stirrup:')
                label.setProperty('class', 'summary-label')
                reinf_layout.addRow(label)
                extent = QLabel('...')
                extent.setProperty('class', 'summary-value')
                label = QLabel('Extent:')
                label.setProperty('class', 'summary-label-2')
                reinf_layout.addRow(label, extent)
                spacing = QLabel('...')
                spacing.setProperty('class', 'summary-value')
                label = QLabel('Spacing:')
                label.setProperty('class', 'summary-label-2')
                reinf_layout.addRow(label, spacing)
                types = QLabel('...')
                types.setProperty('class', 'summary-value')
                label = QLabel('Bundled Types:')
                label.setProperty('class', 'summary-label-2')
                reinf_layout.addRow(label, types)
                self.summary_labels[field] = {'Extent': extent, 'Spacing': spacing, 'Types': types}
                continue
            value_label = QLabel('...')
            value_label.setProperty('class', 'summary-value')
            self.summary_labels[field] = value_label

            label = QLabel(f'{field}:')
            label.setProperty('class', 'summary-value')
            reinf_layout.addRow(label, value_label)

        # Market Lengths Section - Using a single label for rich text
        market_label = QLabel('...')
        market_label.setProperty('class', 'summary-value')
        market_label.setWordWrap(True)
        self.summary_labels['market_lengths'] = market_label
        market_layout.addRow(market_label)  # AddRow without a label makes it span both columns

        # --- Create two columns ---
        label = QLabel('Summary')
        label.setProperty('class', 'header-0')
        main_layout.addWidget(label)
        grid_layout = QGridLayout()
        grid_layout.addWidget(footing_box, 0, 0)
        grid_layout.addWidget(reinf_box, 0, 1, 2, 1)
        grid_layout.addWidget(market_box, 1, 0)
        main_layout.addLayout(grid_layout)
        main_layout.addStretch(1)

        # --- Navigation Buttons ---
        button_layout = QHBoxLayout()
        back_button = HoverButton('Back')
        back_button.setProperty('class', 'red-button')
        back_button.clicked.connect(self.go_back_to_market_lengths_page)

        generate_button = HoverButton('Generate Excel')
        generate_button.setProperty('class', 'green-button')
        generate_button.clicked.connect(self.generate_cutting_list)

        button_layout.addWidget(back_button)
        button_layout.addStretch()
        button_layout.addWidget(generate_button)

        main_layout.addLayout(button_layout)

        self.stacked_widget.addWidget(page)

    def populate_summary_page(self) -> None:
        """Gathers data from all input pages and populates the summary page."""
        # -- Clean the data ---
        self.footing_details_values = parse_nested_dict(self.footing_details)
        self.rsb_details_values = parse_nested_dict(self.rsb_details)
        self.rsb_details_values['Stirrups']['Spacing'] = parse_spacing_string(self.rsb_details_values['Stirrups']['Spacing'])

        # --- Footing Details ---
        for name, value in self.footing_details_values.items():
            if name in self.summary_labels:
                if value == '':
                    continue
                elif 'Number of Footing' in name:
                    text = f'{value}'
                else:
                    text = f'{value} mm'
                # Special case for 'n_ped' which is not a widget
                if name == 'Footing Type':
                    n_ped = self.footing_details_values['n_ped']
                    if n_ped == 1:
                        text += f' ({n_ped} Pedestal)'
                    else:
                        text += f' ({n_ped} Pedestals)'
                self.summary_labels[name].setText(text)

        # --- Reinforcement Details ---
        top_bar_enabled = self.group_box['Top Bar'].isChecked()
        perimeter_bar_enabled = self.group_box['Perimeter Bar'].isChecked()
        stirrups_enabled = self.group_box['Stirrups'].isChecked()

        # Top & Bottom Bars
        for bar_type in ['Top Bar', 'Bottom Bar']:
            details = self.rsb_details_values[bar_type]
            if (bar_type == 'Top Bar' and top_bar_enabled) or (bar_type == 'Bottom Bar'):
                dia = details['Diameter']
                qty_or_spacing = details['Input Type']
                qty_or_spacing = qty_or_spacing.strip(':')
                x_value = details['Value Along X']
                y_value = details['Value Along Y']
                unit = 'mm' if qty_or_spacing=='Spacing' else ('pcs' if x_value > 1 else 'pc')
                text = (f'{dia} | {qty_or_spacing} along X: {x_value} {unit}'
                        f' | {qty_or_spacing} along Y: {y_value} {unit}')
            else:
                text = 'None'
            self.summary_labels[bar_type].setText(text)

        # Vertical Bars
        details: dict
        details = self.rsb_details_values['Vertical Bar']
        dia = details['Diameter']
        qty = details['Quantity']
        hook_calc = details['Hook Calculation']
        hook_len = details['Hook Length']
        if hook_calc == 'Manual':
            text = f'{dia} | Qty: {qty} pcs/pedestal | Hook: {hook_len} mm'
        else:
            text = f'{dia} | Qty: {qty} pcs/pedestal | Hook: {hook_calc}'
        self.summary_labels['Vertical Bar'].setText(text)

        # Perimeter Bars
        details = self.rsb_details_values['Perimeter Bar']
        dia = details['Diameter']
        layers = details['Layers']
        if not perimeter_bar_enabled:
            self.summary_labels['Perimeter Bar'].setText('None')
        elif layers == '1':
            self.summary_labels['Perimeter Bar'].setText(f'{dia} | {layers} layer')
        else:
            self.summary_labels['Perimeter Bar'].setText(f'{dia} | {layers} layers')

        # Stirrup Detail
        details = self.rsb_details_values['Stirrups']
        if stirrups_enabled:
            extent_text = details['Extent']
        else:
            extent_text = 'N/A'
        self.summary_labels['Stirrups']['Extent'].setText(extent_text)
        if stirrups_enabled:
            rebuilt_listed_spacing = []
            for qty, spacing in details['Spacing']:
                rebuilt_listed_spacing.append(f'{qty}@{spacing}')
            self.summary_labels['Stirrups']['Spacing'].setText(', '.join(rebuilt_listed_spacing))
        else:
            self.summary_labels['Stirrups']['Spacing'].setText('N/A')

        if stirrups_enabled:
            stirrup_types_text = []
            for row in details['Types']:
                s_type = row['Type']
                s_dia = row['Diameter']
                s_a = row['a_input']
                type_str = f'{s_type} ({s_dia})'
                if s_type in ['Tall', 'Wide', 'Octagon']:
                    type_str += f' | a = {s_a} mm'
                stirrup_types_text.append(type_str)
            self.summary_labels['Stirrups']['Types'].setText('\n'.join(stirrup_types_text))
        else:
            self.summary_labels['Stirrups']['Types'].setText('None')

        # --- Market Lengths ---
        market_text_lines = []
        for dia, lengths in self.market_lengths_checkboxes.items():
            available = [l for l, cb in lengths.items() if cb.isChecked()]
            if available:
                market_text_lines.append(f'<b>{dia}:</b> {', '.join(available)}')
        self.summary_labels['market_lengths'].setText('<br>'.join(market_text_lines))

    def setup_connections(self) -> None:
        """Connects signals from input widgets to their corresponding slots."""
        # Widgets from the footing page that affect the stirrup drawing
        self.footing_details['Pedestal Height'].textChanged.connect(self.update_stirrup_drawing)
        self.footing_details['Pedestal Width (Along X)'].textChanged.connect(self.update_stirrup_drawing)
        self.footing_details['Pad Thickness'].textChanged.connect(self.update_stirrup_drawing)
        self.footing_details['Concrete Cover'].textChanged.connect(self.update_stirrup_drawing)

        # Widgets from the rsb page that affect the stirrup drawing
        self.rsb_details['Bottom Bar']['Diameter'].currentTextChanged.connect(self.update_stirrup_drawing)
        self.rsb_details['Vertical Bar']['Diameter'].currentTextChanged.connect(self.update_stirrup_drawing)
        self.rsb_details['Stirrups']['Extent'].currentTextChanged.connect(self.update_stirrup_drawing)
        self.rsb_details['Stirrups']['Spacing'].textChanged.connect(self.on_stirrup_spacing_changed)
        # noinspection PyUnresolvedReferences
        self.debounce_timer.timeout.connect(self.update_stirrup_drawing)

        self.group_box['Stirrups'].toggled.connect(self.update_stirrup_drawing)
        # noinspection PyUnresolvedReferences
        self.stacked_widget.currentChanged.connect(self.on_page_changed)

    def on_page_changed(self, index: int) -> None:
        """
        Slot that triggers when the stacked widget's current page changes.

        Args:
            index: The index of the newly visible page.
        """
        # Check if the new page is the summary page (index 3)
        if index == 3:
            self.populate_summary_page()

    def update_spinbox_suffixes(self, title: str, selection: str) -> None:
        """Updates the suffix for the value spin boxes."""
        details = self.rsb_details[title]
        suffix = ' pcs' if 'Quantity' in selection else ' mm'
        details['Value Along X'].setSuffix(suffix)
        details['Value Along Y'].setSuffix(suffix)

    def toggle_same_for_both(self, title: str, state: int) -> None:
        """Enables/disables the Y-value spinbox and connects/disconnects signals."""
        details = self.rsb_details[title]
        value_x_spinbox = details['Value Along X']
        value_y_spinbox = details['Value Along Y']

        # --- THIS IS THE CORRECTED LINE ---
        if state == Qt.CheckState.Checked.value:
            value_y_spinbox.setValue(value_x_spinbox.value())
            value_y_spinbox.setEnabled(False)
            # Store the connection within the details dictionary
            details['connection'] = value_x_spinbox.valueChanged.connect(
                value_y_spinbox.setValue
            )
        else:
            value_y_spinbox.setEnabled(True)
            # Disconnect the signal if the connection exists
            if details['connection']:
                try:
                    value_x_spinbox.valueChanged.disconnect(details['connection'])
                except TypeError:
                    pass  # Connection might have already been broken
                details['connection'] = None

    def generate_cutting_list(self) -> None:
        """
        Performs all calculations, generates the Excel file, and handles post-generation actions.
        """
        results = {}
        footing_details = self.footing_details_values
        rsb_details = self.rsb_details_values

        n_ped = footing_details['n_ped']
        n_footing = footing_details['Total Number of Footing']
        cc = footing_details['Concrete Cover']
        bx = footing_details['Pedestal Width (Along X)']
        by = footing_details['Pedestal Width (Along Y)']
        h = footing_details['Pedestal Height']
        Bx = footing_details['Pad Width (Along X)']
        By = footing_details['Pad Width (Along Y)']
        t = footing_details['Pad Thickness']

        # Top and Bottom Bar
        def top_bottom_bar_helper(title):
            bar_detail = rsb_details[title]
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
            return top_bottom_bar_calculation(get_bar_dia(bar_detail['Diameter']), Bx, By, t, cc, bar_spacing_value_x,
                                              bar_spacing_value_y, bar_qty_x, bar_qty_y)
        if self.group_box['Top Bar'].isChecked():
            results['Top Bar'] = top_bottom_bar_helper('Top Bar')
            results['Top Bar']['bar_in_x_direction']['quantity'] *= n_footing
            results['Top Bar']['bar_in_y_direction']['quantity'] *= n_footing
        results['Bottom Bar'] = top_bottom_bar_helper('Bottom Bar')
        results['Bottom Bar']['bar_in_x_direction']['quantity'] *= n_footing
        results['Bottom Bar']['bar_in_y_direction']['quantity'] *= n_footing

        # Perimeter Bar
        if self.group_box['Perimeter Bar'].isChecked():
            perim_bar = rsb_details['Perimeter Bar']
            layers = perim_bar['Layers']
            if isinstance(layers, int):
                dia = get_bar_dia(perim_bar['Diameter'])
                results['Perimeter Bar'] = perimeter_bar_calculation(dia, layers, Bx, By, cc)
                results['Perimeter Bar']['bar_in_x_direction']['quantity'] *= n_footing
                results['Perimeter Bar']['bar_in_y_direction']['quantity'] *= n_footing

        # Vertical Bar
        vert_bar = rsb_details['Vertical Bar']
        dia = get_bar_dia(vert_bar['Diameter'])
        hook_calc = vert_bar['Hook Calculation']
        bot_bar_dia = get_bar_dia(rsb_details['Bottom Bar']['Diameter'])
        qty = vert_bar['Quantity']
        if 'Manual' in hook_calc:
            hook_len = vert_bar['Hook Length']
        else:
            hook_len = None
        results['Vertical Bar'] = vertical_bar_calculation(dia, qty, h, t, cc, bot_bar_dia, hook_len)
        results['Vertical Bar']['quantity'] *= n_ped * n_footing

        # Stirrups
        if self.group_box['Stirrups'].isChecked():
            stirrup = rsb_details['Stirrups']
            qty = self.stirrup_canvas.get_qty() * n_ped
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
                results['Stirrups'] = stirrups_cutting_list

        # Market Length
        processed_for_optimization = process_rebar_input(results)
        cuts_by_diameter = {}
        for bar in processed_for_optimization:
            dia = get_dia_code(bar['diameter'])
            if dia not in cuts_by_diameter:
                cuts_by_diameter[dia] = {}

            length = bar['cut_length']
            quantity = bar['quantity']

            if length in cuts_by_diameter[dia]:
                cuts_by_diameter[dia][length] += quantity
            else:
                cuts_by_diameter[dia][length] = quantity
        for key, value in cuts_by_diameter.items():
            cuts_by_diameter[key] = [(q, l/1000) for l, q in value.items()]

        market_lengths = {}
        for dia_code, lengths in self.market_lengths_checkboxes.items():
            if dia_code in cuts_by_diameter:
                # Market lengths are already whole numbers, so they are fine.
                available_lengths = [float(l.replace('m', '')) for l, cb in lengths.items() if cb.isChecked()]
                if not available_lengths:
                    continue
                market_lengths[dia_code] = available_lengths

        # --- Prompt user for save location ---
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            'Save Cutting List As',
            'rebar_cutting_schedule.xlsx',
            'Excel Files (*.xlsx);;All Files (*)'
        )

        if not save_path:
            print('File save cancelled by user.')
            return

        # --- Try to save the file and handle potential errors ---
        try:
            # Pass optimization_results to the Excel function
            create_excel_cutting_list(results, cuts_by_diameter, market_lengths, output_filename=save_path)
        except PermissionError:
            QMessageBox.warning(
                self,
                'Save Error',
                f"Could not save the file to '{os.path.basename(save_path)}'.\n\n"
                'Please ensure the file is not already open in another program and that you have permission to write to this location.'
            )
            return  # Stop execution if save fails

        # --- Open the saved file ---
        try:
            if sys.platform == 'win32':
                os.startfile(save_path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', save_path])
            else:
                subprocess.call(['xdg-open', save_path])
        except Exception as e:
            print(f'Could not open file automatically: {e}')

        # --- Ask the user what to do next, with reliable button styling ---
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle('Generation Complete')
        msg_box.setText('The cutting list has been generated and saved.')
        msg_box.setInformativeText('What would you like to do next?')
        msg_box.setIcon(QMessageBox.Icon.Question)

        msg_box.setStandardButtons(QMessageBox.StandardButton.Reset | QMessageBox.StandardButton.Close)

        start_over_btn = msg_box.button(QMessageBox.StandardButton.Reset)
        start_over_btn.setText('Start Over')
        start_over_btn.setStyleSheet("""background-color: #3498db; 
        color: white; 
        border: 1px solid #2980b9;
        min-width: 90px; 
        font-weight: bold; 
        padding: 8px 16px; 
        border-radius: 5px;""")

        close_btn = msg_box.button(QMessageBox.StandardButton.Close)
        close_btn.setText('Close Program')
        close_btn.setStyleSheet("""background-color: #E1E1E1; 
        color: #2c3e50; 
        border: 1px solid #ADADAD;
        min-width: 90px; 
        font-weight: bold; 
        padding: 8px 16px; 
        border-radius: 5px;""")

        msg_box.setDefaultButton(start_over_btn)
        reply = msg_box.exec()

        if reply == QMessageBox.StandardButton.Reset:
            self.reset_application()
        else:
            self.close()

    def prefill_for_debug(self) -> None:
        """Pre-fills all input fields with sample data for faster testing."""
        print("--- DEBUG MODE: Pre-filling forms with sample data. ---")

        # --- Page 1: Footing Details ---
        self.footing_details['Footing Type'].setCurrentText('Isolated')
        self.footing_details['Total Number of Footing'].setValue(2)
        self.footing_details['Concrete Cover'].setValue(75)
        self.footing_details['Pedestal Width (Along X)'].setValue(600)
        self.footing_details['Pedestal Width (Along Y)'].setValue(600)
        self.footing_details['Pedestal Height'].setValue(1200)
        self.footing_details['Pad Width (Along X)'].setValue(1800)
        self.footing_details['Pad Width (Along Y)'].setValue(1800)
        self.footing_details['Pad Thickness'].setValue(400)

        # --- Page 2: Reinforcement Details ---
        # Top Bar
        self.rsb_details['Top Bar']['Diameter'].setCurrentText('#20')
        self.rsb_details['Top Bar']['Input Type'].setCurrentText('Spacing:')
        self.rsb_details['Top Bar']['Value Along X'].setValue(150)
        self.rsb_details['Top Bar']['Value Along Y'].setValue(150)

        # Bottom Bar
        self.rsb_details['Bottom Bar']['Diameter'].setCurrentText('#20')
        self.rsb_details['Bottom Bar']['Input Type'].setCurrentText('Spacing:')
        self.rsb_details['Bottom Bar']['Value Along X'].setValue(170)
        self.rsb_details['Bottom Bar']['Value Along Y'].setValue(170)

        # Vertical Bar
        self.rsb_details['Vertical Bar']['Diameter'].setCurrentText('#20')
        self.rsb_details['Vertical Bar']['Quantity'].setValue(8)
        self.rsb_details['Vertical Bar']['Hook Calculation'].setCurrentText('Manual')
        self.rsb_details['Vertical Bar']['Hook Length'].setValue(300)

        # Perimeter Bar
        self.rsb_details['Perimeter Bar']['Layers'].setCurrentText('2')
        self.rsb_details['Perimeter Bar']['Diameter'].setCurrentText('#12')

        # Stirrups
        self.rsb_details['Stirrups']['Extent'].setCurrentText('From Face of Pad')
        self.rsb_details['Stirrups']['Spacing'].setText('1 @ 50, 5 @ 100, rest @ 150')
        # Add a second stirrup type for more complex testing
        self.add_stirrup_row()
        stirrup_row_1 = self.rsb_details['Stirrups']['Types'][0]
        stirrup_row_1['Type'].setCurrentText('Outer')
        stirrup_row_1['Diameter'].setCurrentText('#10')

        stirrup_row_2 = self.rsb_details['Stirrups']['Types'][1]
        stirrup_row_2['Type'].setCurrentText('Tall')
        stirrup_row_2['Diameter'].setCurrentText('#10')
        stirrup_row_2['a_input'].setValue(250)

        # --- Page 3: Market Lengths ---
        # Uncheck a couple of options for testing
        self.market_lengths_checkboxes['#10']['6m'].setChecked(False)
        self.market_lengths_checkboxes['#25']['12m'].setChecked(False)

    def reset_application(self) -> None:
        """Resets all input fields to their default states and returns to the first page."""
        # --- Reset Footing Page ---
        for name, widget in self.footing_details.items():
            if name == 'Footing Type':
                widget.setCurrentIndex(0)
            elif name == 'Total Number of Footing':
                widget.setValue(1)
            elif name == 'Concrete Cover':
                widget.setValue(75)
            elif isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                widget.setValue(widget.minimum())

        # --- Reset RSB Page ---
        for section, details in self.rsb_details.items():
            for widget in details.values():
                if isinstance(widget, QComboBox):
                    widget.setCurrentIndex(0)
                elif isinstance(widget, (QLineEdit, QTextEdit)):
                    widget.clear()
                elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                    widget.setValue(widget.minimum())

        # Reset stirrup rows to a single, default row
        while len(self.rsb_details['Stirrups']['Types']) > 1:
            self.remove_stirrup_row()

        if self.rsb_details['Stirrups']['Types']:
            first_row = self.rsb_details['Stirrups']['Types'][0]
            first_row['Type'].setCurrentIndex(0)
            first_row['Diameter'].setCurrentIndex(0)
            first_row['a_input'].setValue(first_row['a_input'].minimum())

        # --- Reset Market Lengths Page ---
        for dia_lengths in self.market_lengths_checkboxes.values():
            for checkbox in dia_lengths.values():
                checkbox.setChecked(True)

        # --- Go back to the first page ---
        self.stacked_widget.setCurrentIndex(0)

    def validate_and_style_stirrup_spacing(self) -> bool:
        """
        Validates the stirrup spacing QTextEdit, applies styling, and returns the validity.

        Returns:
            True if the spacing format is valid, False otherwise.
        """
        widget = self.rsb_details['Stirrups']['Spacing']
        text = widget.toPlainText()
        is_valid = True

        # An empty string is considered valid (no stirrups)
        if text.strip():
            try:
                parse_spacing_string(text)
            except (ValueError, TypeError):
                is_valid = False

        style_invalid_input(widget, is_valid)
        return is_valid

    def show_error_message(self, title: str, message: str) -> None:
        """
        Displays a standardized error message box.

        Args:
            title: The title for the message box window.
            message: The informative text to display.
        """
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Warning)
        msg_box.setWindowTitle(title)
        msg_box.setText('Please correct the following errors before proceeding:')
        msg_box.setInformativeText(message)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.exec()

    def validate_footing_page(self) -> list[str]:
        """
        Validates all inputs on the footing dimensions page with enhanced logic.

        Returns:
            A list of error messages. An empty list indicates success.
        """
        errors = []
        details = self.footing_details
        validity_map = {widget: True for widget in details.values() if isinstance(widget, QWidget)}

        # --- Get key values for cross-validation ---
        cc_widget = details['Concrete Cover']
        cc = cc_widget.value()
        pad_t_widget = details['Pad Thickness']
        ped_h_widget = details['Pedestal Height']
        pad_bx_widget = details['Pad Width (Along X)']
        ped_bx_widget = details['Pedestal Width (Along X)']
        pad_by_widget = details['Pad Width (Along Y)']
        ped_by_widget = details['Pedestal Width (Along Y)']

        # --- Rule 1: All dimension and quantity fields must be > 0 ---
        required_fields = [
            'Total Number of Footing', 'Concrete Cover', 'Pedestal Width (Along X)',
            'Pedestal Width (Along Y)', 'Pedestal Height', 'Pad Width (Along X)',
            'Pad Width (Along Y)', 'Pad Thickness'
        ]
        for name in required_fields:
            widget = details[name]
            if widget.value() <= 0:
                validity_map[widget] = False
                errors.append(f"- '{name}' must be greater than 0.")

        # --- Rule 2: Pedestal must be smaller than the Pad ---
        if ped_bx_widget.value() >= pad_bx_widget.value():
            validity_map[ped_bx_widget] = validity_map[pad_bx_widget] = False
            errors.append("- 'Pedestal Width (X)' must be smaller than 'Pad Width (X)'.")

        if ped_by_widget.value() >= pad_by_widget.value():
            validity_map[ped_by_widget] = validity_map[pad_by_widget] = False
            errors.append("- 'Pedestal Width (Y)' must be smaller than 'Pad Width (Y)'.")

        # --- NEW: Rule 3: Dimensions must be greater than twice the concrete cover ---
        if cc > 0:
            min_dim = 2 * cc
            if pad_t_widget.value() <= min_dim:
                validity_map[pad_t_widget] = validity_map[cc_widget] = False
                errors.append(f"- 'Pad Thickness' must be > 2 * Concrete Cover (> {min_dim} mm).")

            if ped_bx_widget.value() <= min_dim:
                validity_map[ped_bx_widget] = validity_map[cc_widget] = False
                errors.append(f"- 'Pedestal Width (X)' must be > 2 * Concrete Cover (> {min_dim} mm).")

            if ped_by_widget.value() <= min_dim:
                validity_map[ped_by_widget] = validity_map[cc_widget] = False
                errors.append(f"- 'Pedestal Width (Y)' must be > 2 * Concrete Cover (> {min_dim} mm).")

            if ped_h_widget.value() <= min_dim:
                validity_map[ped_h_widget] = validity_map[cc_widget] = False
                errors.append(f"- 'Pedestal Height' must be > 2 * Concrete Cover (> {min_dim} mm).")

        # --- Apply styles based on final validity ---
        for widget, is_valid in validity_map.items():
            style_invalid_input(widget, is_valid)

        return sorted(list(set(errors)))

    def validate_rsb_page(self) -> list[str]:
        """
        Validates all inputs on the reinforcement (RSB) page with enhanced, cross-dependent logic.

        Returns:
            A list of error messages. An empty list indicates success.
        """
        errors = []
        validity_map = {}

        # --- Get necessary values from the Footing page for cross-validation ---
        footing = self.footing_details
        cc_widget = footing['Concrete Cover']
        cc = cc_widget.value()
        pad_bx_widget = footing['Pad Width (Along X)']
        pad_bx = pad_bx_widget.value()
        pad_by_widget = footing['Pad Width (Along Y)']
        pad_by = pad_by_widget.value()
        ped_bx_widget = footing['Pedestal Width (Along X)']
        ped_bx = ped_bx_widget.value()
        ped_by_widget = footing['Pedestal Width (Along Y)']
        ped_by = ped_by_widget.value()

        # --- Rule 1: Validate Top and Bottom Bars ---
        for bar_type in ['Top Bar', 'Bottom Bar']:
            details = self.rsb_details[bar_type]
            # Map directions to their corresponding pad dimensions
            dir_map = {
                'Value Along X': (pad_bx_widget, pad_bx, 'X'),
                'Value Along Y': (pad_by_widget, pad_by, 'Y')
            }

            for direction, (pad_widget, pad_dim, axis) in dir_map.items():
                widget = details[direction]
                validity_map.setdefault(widget, True)
                value = widget.value()
                input_type = details['Input Type'].currentText()

                if value <= 0:
                    validity_map[widget] = False
                    errors.append(f"- {bar_type} ({axis}): {input_type.strip(':')} must be > 0.")
                    continue  # Skip further checks if the base rule fails

                if 'Spacing' in input_type:
                    limit = pad_dim / 2
                    if pad_dim > 0 and value > limit:
                        validity_map[widget] = validity_map[pad_widget] = False
                        errors.append(
                            f"- {bar_type} ({axis}): Spacing cannot be > half the Pad Width (> {limit:.0f} mm).")
                else:  # Quantity
                    if value < 3:
                        validity_map[widget] = False
                        errors.append(f"- {bar_type} ({axis}): Quantity should not be less than 3.")

                    limit = pad_dim - 2 * cc
                    if pad_dim > 0 and cc > 0 and value > limit:
                        validity_map[widget] = validity_map[pad_widget] = validity_map[cc_widget] = False
                        errors.append(
                            f"- {bar_type} ({axis}): Quantity seems too high for the Pad Width (limit is approx. {limit:.0f}).")

        # --- Rule 2: Validate Vertical Bar ---
        vert_details = self.rsb_details['Vertical Bar']
        qty_widget = vert_details['Quantity']
        validity_map.setdefault(qty_widget, True)
        if qty_widget.value() <= 0:
            validity_map[qty_widget] = False
            errors.append("- Vertical Bar: 'Quantity' must be > 0.")
        else:
            # NEW: Check quantity against a calculated perimeter (as a loose upper bound)
            if ped_bx > 0 and ped_by > 0:
                perimeter = 2 * (ped_bx + ped_by)
                if qty_widget.value() > perimeter:
                    validity_map[qty_widget] = validity_map[ped_bx_widget] = validity_map[ped_by_widget] = False
                    errors.append(
                        f"- Vertical Bar: Quantity ({qty_widget.value()}) is unusually high for the pedestal perimeter ({perimeter:.0f} mm).")

        if vert_details['Hook Calculation'].currentText() == 'Manual':
            hook_widget = vert_details['Hook Length']
            validity_map.setdefault(hook_widget, True)
            if hook_widget.value() <= 0:
                validity_map[hook_widget] = False
                errors.append("- Vertical Bar: Manual 'Hook Length' must be > 0.")
            # NEW: Check hook length against pad dimensions
            elif pad_bx > 0 and pad_by > 0 and cc > 0:
                limit = (min(pad_bx, pad_by) / 2) - cc
                if hook_widget.value() > limit:
                    validity_map[hook_widget] = validity_map[pad_bx_widget] = validity_map[pad_by_widget] = \
                    validity_map[cc_widget] = False
                    errors.append(
                        f"- Vertical Bar: 'Hook Length' cannot extend beyond the center of the pad (> {limit:.0f} mm).")

        # --- Rule 3: Validate Stirrup Spacing Format ---
        if not self.validate_and_style_stirrup_spacing():
            errors.append("- Stirrup 'Spacing': Invalid format. Example: 1@50, rest@100")

        # --- Rule 4: Validate Stirrup Bundle 'a' inputs ---
        if ped_bx > 0 and cc > 0:
            limit = min(ped_bx, ped_by) - 2 * cc
            for i, stirrup_row in enumerate(self.rsb_details['Stirrups']['Types']):
                stirrup_type = stirrup_row['Type'].currentText()
                a_input_widget = stirrup_row['a_input']
                validity_map.setdefault(a_input_widget, True)

                if stirrup_type in ['Tall', 'Wide', 'Octagon']:
                    if a_input_widget.value() <= 0:
                        validity_map[a_input_widget] = False
                        errors.append(
                            f"- Stirrup (Row {i + 1}): The 'a' value for a '{stirrup_type}' type must be > 0.")
                    # NEW: Check 'a' value against pedestal dimensions
                    elif a_input_widget.value() > limit:
                        validity_map[a_input_widget] = validity_map[ped_bx_widget] = validity_map[ped_by_widget] = \
                        validity_map[cc_widget] = False
                        errors.append(
                            f"- Stirrup (Row {i + 1}): 'a' value cannot be larger than the internal pedestal dimension (> {limit:.0f} mm).")

        # --- Apply styles based on final validity ---
        for widget, is_valid in validity_map.items():
            if widget:  # Ensure widget is not None
                style_invalid_input(widget, is_valid)

        return sorted(list(set(errors)))

    def update_remove_button_state(self) -> None:
        """Enables or disables the 'remove stirrup row' button based on the row count."""
        self.remove_stirrup_button.setEnabled(len(self.rsb_details['Stirrups']['Types']) > 1)

    def add_stirrup_row(self) -> None:
        """Creates and adds a new UI row for defining a stirrup type."""
        # --- Main container for the row ---
        row_widget = QWidget()
        row_widget.setProperty('class', 'stirrup-row')
        row_layout = QHBoxLayout(row_widget)

        # --- Image (Left) ---
        image_label = get_img(resource_path('images/stirrup_outer.png'), STIRRUP_ROW_IMAGE_WIDTH, STIRRUP_ROW_IMAGE_WIDTH)
        row_layout.addWidget(image_label)

        image_map = {
            'Outer': resource_path('images/stirrup_outer.png'),
            'Diamond': resource_path('images/stirrup_diamond.png'),
            'Tall': resource_path('images/stirrup_tall.png'),
            'Wide': resource_path('images/stirrup_wide.png'),
            'Octagon': resource_path('images/stirrup_octagon.png')
        }

        # --- Form (Right) ---
        form_layout = QFormLayout()
        type_combo = QComboBox()
        type_combo.addItems(image_map.keys())
        size_policy = type_combo.sizePolicy()
        size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
        type_combo.setSizePolicy(size_policy)

        dia_combo = QComboBox()
        dia_combo.addItems(BAR_DIAMETERS_FOR_STIRRUPS)

        a_label = QLabel('a:')
        a_label.setProperty('class', 'rsb-forms-label')
        a_input = BlankSpinBox(0, 99_999, suffix=' mm')

        label = QLabel('Type:')
        label.setProperty('class', 'rsb-forms-label')
        form_layout.addRow(label, type_combo)
        label = QLabel('Diameter:')
        label.setProperty('class', 'rsb-forms-label')
        form_layout.addRow(label, dia_combo)
        form_layout.addRow(a_label, a_input)
        row_layout.addLayout(form_layout)

        # --- Store widgets for later access ---
        row_widgets = {
            'Row': row_widget,
            'Image': image_label,
            'Type': type_combo,
            'Diameter': dia_combo,
            'a_label': a_label,
            'a_input': a_input
        }
        self.rsb_details['Stirrups']['Types'].append(row_widgets)

        # --- Connections ---
        # noinspection PyUnresolvedReferences
        type_combo.currentTextChanged.connect(
            lambda text, widgets=row_widgets: self.update_stirrup_row_visibility(text, widgets, image_map)
        )

        # --- Set initial state ---
        self.update_stirrup_row_visibility(type_combo.currentText(), row_widgets, image_map)

        # --- Add to the main container ---
        self.stirrup_rows_layout.addWidget(row_widget)
        self.update_remove_button_state()

    def remove_stirrup_row(self) -> None:
        """Removes the last stirrup definition row from the UI."""
        if len(self.rsb_details['Stirrups']['Types']) > 1:  # Keep at least one row
            widgets_to_remove = self.rsb_details['Stirrups']['Types'].pop()
            widgets_to_remove['Row'].deleteLater()  # Safely delete the widget

        self.update_remove_button_state()

    @staticmethod
    def update_stirrup_row_visibility(selected_text: str, widgets: dict[str, Any],
                                      image_map: dict[str, str]) -> None:
        """
        Updates a stirrup row's image and the visibility of its 'a' input field.

        Args:
            selected_text: The selected stirrup type from the combo box.
            widgets: A dictionary of the widgets in that specific row.
            image_map: A dictionary mapping stirrup types to image paths.
        """
        update_image(selected_text, image_map, widgets['Image'], STIRRUP_ROW_IMAGE_WIDTH,
                     fallback=resource_path('images/stirrup_outer.png'))

        # Update visibility of 'a' input
        is_visible = selected_text in ['Tall', 'Wide', 'Octagon']
        widgets['a_label'].setVisible(is_visible)
        widgets['a_input'].setVisible(is_visible)

    def eventFilter(self, obj, event):
        """
        Filters out mouse wheel events on all QComboBoxes to prevent
        accidental value changes while scrolling the page.
        """
        # Check if the event is a wheel event and the object is a QComboBox.
        if event.type() == QEvent.Type.Wheel and isinstance(obj, (QComboBox, QSpinBox, QDoubleSpinBox)):
            # Return True to indicate the event has been handled and should be ignored.
            return True

        # For all other events, pass them to the default implementation.
        return super().eventFilter(obj, event)

    def show_hook_info(self) -> None:
        """Displays an informational popup for the hook calculation method."""
        # NOTE: You should update the text to reflect the actual standard you are using.
        # For example, ACI 318 or a local building code.
        info_text = (
            "<b>Hook Calculation Method</b>"
            "<ul><li><b>Automatic:</b> Calculates the required hook length "
            "based on ACI 318-25 Table 25.3.1 Standard 90° hook geometry for "
            "development of deformed bars in tension (<i>12d</i><sub>b</sub>) </li>"
            "<li><b>Manual:</b> Allows you to enter a custom, pre-calculated "
            "length for the hook.</li></ul>"
        )
        self.info_popup.set_info_text(info_text)

        # Position and show the popup
        cursor_pos = self.cursor().pos()
        self.info_popup.move(cursor_pos.x() + 15, cursor_pos.y() + 15)
        self.info_popup.show()

    def show_spacing_header_info(self) -> None:
        """Displays an informational popup for the stirrup spacing section."""
        info_text = (
            """<b>Stirrup Placement Guide</b><br><br>
This section controls the vertical position and distribution of the stirrup bundles along the pedestal.
<br><br>
It's a two-step process:
<ol>
    <li><b>Start From:</b> First, select your 'zero' reference point from which all measurements will begin.</li>
    <li><b>Spacing:</b> Next, enter the series of spacing values. The first value positions the first stirrup relative to your chosen start point.</li>
</ol>
Use the diagram on the left to visually confirm that the stirrup placement matches your input."""
        )
        self.info_popup.set_info_text(info_text)

        # Position and show the popup
        cursor_pos = self.cursor().pos()
        self.info_popup.move(cursor_pos.x() + 15, cursor_pos.y() + 15)
        self.info_popup.show()

    def show_spacing_extent_info(self) -> None:
        """Displays an informational popup for the stirrup extent."""
        info_text = (
            """<b>Spacing Start Point</b><br><br>
This sets the <b>'zero' reference point</b> for the first measurement in the 'Spacing' field.
<hr>
<ul>
    <li><b>From Face of Pad:</b> 'Zero' is the top face of the footing (pad).</li>
    <li><b>From Bottom Bar:</b> 'Zero' is at the elevation of the bottom rebar.</li>
    <li><b>From Top (to Face of Pad):</b> 'Zero' is the top of the pedestal, measuring downwards. Spacing will only be applied within the concrete pedestal.</li>
</ul>"""
        )
        self.info_popup.set_info_text(info_text)

        # Position and show the popup
        cursor_pos = self.cursor().pos()
        self.info_popup.move(cursor_pos.x() + 15, cursor_pos.y() + 15)
        self.info_popup.show()

    def show_bundle_info(self) -> None:
        """Displays an informational popup for the stirrup bundle."""
        info_text = (
            """<b>Understanding the Stirrup Bundle</b><br><br>
This section defines the combination of stirrups that will be installed together as a single unit.
<ul>
    <li><b>Forms One Set:</b> All stirrup shapes you add (e.g., an Outer, a Tall) are considered one complete set.</li>
    <li><b>Installed as a Group:</b> At each specified height, all stirrups in the set are installed as a single, tightly packed group.</li>
    <li><b>Spacing Applies to the Group:</b> The spacing you define (e.g., <code>5@100</code>) dictates the vertical distance from the center of one group to the center of the next.</li>
</ul>
Think of it as designing a "kit" of stirrups that gets repeated along the height of the pedestal."""
        )
        self.info_popup.set_info_text(info_text)

        # Position and show the popup
        cursor_pos = self.cursor().pos()
        self.info_popup.move(cursor_pos.x() + 15, cursor_pos.y() + 15)
        self.info_popup.show()

    def show_spacing_info(self) -> None:
        """Displays an informational popup for the stirrup spacing format."""
        info_text = (
            """<b>Stirrup Spacing Guide</b><br><br>
Defines stirrup locations relative to your chosen 'Start From' point.

<hr>

<b>Key Principle:</b>
<p>The <u>first spacing value</u> in your list always positions the <u>first stirrup</u>.</p>

<b>Example: </b> <b><code>5@100, rest@150</code></b></p>
<ul>
    <li>The <b>first</b> of the 5 stirrups is placed <b>100mm</b> from the start point.</li>
    <li>The next 4 are also 100mm apart. The remaining are 150mm apart.</li>"""
        )
        self.info_popup.set_info_text(info_text)

        # Position and show the popup
        cursor_pos = self.cursor().pos()
        self.info_popup.move(cursor_pos.x() + 15, cursor_pos.y() + 15)
        self.info_popup.show()

    def on_stirrup_spacing_changed(self) -> None:
        """Slot called on every keystroke in the stirrup spacing editor to start the debounce timer."""
        self.validate_and_style_stirrup_spacing()
        self.debounce_timer.start()  # This restarts the timer for the drawing

    def update_stirrup_drawing(self) -> None:
        """Triggers a repaint of the stirrup drawing canvas with current input values."""
        if hasattr(self, 'stirrup_canvas'):  # Check if canvas exists
            self.stirrup_canvas.update_dimensions(
                self.footing_details,
                self.rsb_details['Stirrups']['Extent'],
                self.rsb_details['Stirrups']['Spacing'],
                self.rsb_details['Bottom Bar']['Diameter'],
                self.rsb_details['Vertical Bar']['Diameter']
            )

    def toggle_market_row(self, dia: str) -> None:
        """
        Toggles all checkboxes in a given market length row.

        Args:
            dia: The diameter string (e.g., '#10') identifying the row.
        """
        row_cbs = self.market_lengths_checkboxes[dia]
        if not row_cbs: return

        # Determine target state based on the opposite of the first checkbox
        first_len = MARKET_LENGTHS[0]
        new_state = not row_cbs[first_len].isChecked()

        for cb in row_cbs.values():
            cb.setChecked(new_state)

    def toggle_market_column(self, length: str) -> None:
        """
        Toggles all checkboxes in a given market length column.

        Args:
            length: The market length string (e.g., '6m') identifying the column.
        """
        if not BAR_DIAMETERS: return

        # Determine target state based on the opposite of the first row's checkbox in this column
        first_dia = BAR_DIAMETERS[0]
        new_state = not self.market_lengths_checkboxes[first_dia][length].isChecked()

        for dia in BAR_DIAMETERS:
            self.market_lengths_checkboxes[dia][length].setChecked(new_state)

    def go_to_footing_page(self) -> None:
        """Navigates to the footing dimensions page (index 0)."""
        self.stacked_widget.setCurrentIndex(0)

    def go_to_rsb_page(self) -> None:
        """Validates the footing page and navigates to the reinforcement page (index 1)."""
        if not DEBUG_MODE:
            errors = self.validate_footing_page()
            if errors:
                self.show_error_message('Footing Page Errors', '\n'.join(errors))
                return  # Stop navigation

        self.update_stirrup_drawing()
        self.stacked_widget.setCurrentIndex(1)

    def go_to_market_lengths_page(self) -> None:
        """Validates the reinforcement page and navigates to the market lengths page (index 2)."""
        # --- Pre-navigation Validation ---
        if not DEBUG_MODE:
            errors = self.validate_rsb_page()
            if errors:
                self.show_error_message('Reinforcement Page Errors', '\n'.join(errors))
                return  # Stop navigation

        # --- Confirmation for Disabled GroupBoxes ---
        disabled_groups = []
        # We check the essential, commonly used bar types.
        # Perimeter are often optional, so we can omit them from this warning.
        if not self.group_box['Top Bar'].isChecked():
            disabled_groups.append("'Top Bar'")
        # if not self.group_box['Perimeter Bar'].isChecked():
        #     disabled_groups.append("'Perimeter Bar'")
        if not self.group_box['Stirrups'].isChecked():
            disabled_groups.append("'Stirrups'")

        if disabled_groups:
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setWindowTitle('Confirm Omissions')
            msg_box.setText(
                "You have disabled the following reinforcement sections:\n\n"
                f"- {', '.join(disabled_groups)}\n\n"
                "This means they will be excluded from the cutting list calculation. "
                "Do you want to proceed?"
            )
            msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            yes_btn = msg_box.button(QMessageBox.StandardButton.Yes)
            yes_btn.setText('Yes')
            yes_btn.setStyleSheet("""background-color: #3498db; 
                    color: white; 
                    border: 1px solid #2980b9;
                    min-width: 90px; 
                    font-weight: bold; 
                    padding: 8px 16px; 
                    border-radius: 5px;""")

            no_btn = msg_box.button(QMessageBox.StandardButton.No)
            no_btn.setText('No')
            no_btn.setStyleSheet("""background-color: #E1E1E1; 
                    color: #2c3e50; 
                    border: 1px solid #ADADAD;
                    min-width: 90px; 
                    font-weight: bold; 
                    padding: 8px 16px; 
                    border-radius: 5px;""")

            msg_box.setDefaultButton(no_btn)
            reply = msg_box.exec()

            if reply == no_btn:
                return # Stop navigation if user clicks No

        self.stacked_widget.setCurrentIndex(2)

    def go_back_to_market_lengths_page(self) -> None:
        """Navigates back to the market lengths page (index 2) without validation."""
        self.stacked_widget.setCurrentIndex(2)

    def go_to_summary_page(self) -> None:
        """Navigates to the summary page (index 3)."""
        self.stacked_widget.setCurrentIndex(3)

    @staticmethod
    def make_scrollable(widget: QWidget, always_on: bool = False) -> QScrollArea:
        """
        Wraps a widget in a QScrollArea.

        Args:
            widget: The widget to make scrollable.
            always_on: If True, the vertical scrollbar is always visible.

        Returns:
            The configured QScrollArea containing the widget.
        """
        scroll = QScrollArea()
        scroll.setWidget(widget)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        if always_on:
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        else:
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        return scroll

if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    app.setStyleSheet(load_stylesheet('style.qss'))
    window = MultiPageApp()
    window.show()
    sys.exit(app.exec())
