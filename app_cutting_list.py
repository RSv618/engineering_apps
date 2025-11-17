import os
import subprocess
import sys
from typing import Any

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QStackedWidget, QWidget, QVBoxLayout,
    QHBoxLayout, QLabel, QPushButton, QLineEdit, QDialog,
    QFormLayout, QSpinBox, QComboBox, QGridLayout, QCheckBox, QTextEdit, QFrame,
    QSizePolicy, QGroupBox, QStyle, QStyleOption, QMessageBox, QFileDialog, QInputDialog
)
from PyQt6.QtGui import QIcon, QColor, QPen, QPainter, QPaintEvent
from PyQt6.QtCore import (Qt, pyqtSignal as Signal, QEvent, QPointF,
                          QTimer)

from constants import (FOOTING_IMAGE_WIDTH, RSB_IMAGE_WIDTH,
                       BAR_DIAMETERS, STIRRUP_ROW_IMAGE_WIDTH,
                       BAR_DIAMETERS_FOR_STIRRUPS, MARKET_LENGTHS,
                       DEBUG_MODE)
from excel_writer import (process_rebar_input, add_sheet_cutting_list,
                          add_shet_purchase_plan, add_sheet_cutting_plan,
                          delete_blank_worksheets)
from rebar_calculations import compile_rebar
from rebar_optimizer import find_optimized_cutting_plan
from utils import (HoverButton, HoverLabel, resource_path,
                   global_exception_hook, load_stylesheet, get_img,
                   BlankSpinBox, update_image, MemoryGroupBox, InfoPopup,
                   parse_spacing_string, get_bar_dia, make_scrollable,
                   LinkSpinboxes, toggle_obj_visibility,
                   GlobalWheelEventFilter, is_widget_empty,
                   style_invalid_input, get_dia_code)
from openpyxl import Workbook

r"""
TO BUILD:
pyinstaller --name 'CuttingList' --onefile --windowed --icon='images/logo.png' --add-data 'images:images' --add-data 'style.qss:.' --add-binary 'C:\Users\rober\PycharmProjects\engineering_apps\.venv\Lib\site-packages\pulp\solverdir\cbc\win\i64\cbc.exe;.' app_cutting_list.py
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
        width += 0
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
        self._recalculate_quantity() # Recalculate the quantity immediately whenever dimensions change.
        self.update()  # Crucial: schedules a repaint which calls paintEvent

    def _recalculate_quantity(self):
        """
        Calculates the stirrup quantity without performing any drawing.
        This allows the value to be updated even if the widget is not visible.
        """
        # This logic is a subset of paintEvent, focused only on calculation.
        real_h = self.ped_h
        real_bx = self.ped_bx
        real_t = self.pad_t

        if real_h + real_t == 0 or real_bx == 0:
            self.stirrup_qty = 0
            return

        # Simplified scale calculation (only what's needed for Y-axis)
        scale = (self.height() - 4) / (real_h + real_t)
        px_cc = self.cc * scale
        px_t = real_t * scale
        px_h = real_h * scale

        # Simplified Y-coordinate setup
        y1 = self.height() - 2
        y2 = y1 - px_t
        y3 = y2 - px_h
        vbar_y2 = y3 + px_cc

        if self.extent == 'From Face of Pad':
            start_y = y2
            target_y = vbar_y2
        elif self.extent == 'From Bottom Bar':
            start_y = (y1 - px_cc - scale * (2 * self.bot_bar_diameter + self.vert_bar_diameter))
            target_y = vbar_y2
        else:  # From Top
            start_y = vbar_y2
            target_y = y2

        actual_count = 0
        if self.extent in ['From Face of Pad', 'From Bottom Bar']:
            _, count, last_y = self.loop_stirrup(self.spacing, start_y=start_y, target_y=target_y,
                                                 left_x=0, right_x=0, scale=scale)  # xs don't matter for count
            actual_count += count
            if (actual_count > 0) and (last_y - vbar_y2 >= px_cc):
                actual_count += 1
        else:  # From Top
            _, count, _ = self.loop_stirrup(self.spacing, start_y=start_y, target_y=target_y,
                                            left_x=0, right_x=0, scale=scale)
            actual_count += count

        self.stirrup_qty = actual_count

    def paintEvent(self, event: QPaintEvent) -> None:
        """
        Handles the repaint event to draw the footing and stirrups on the widget.

        Args:
            event: The paint event.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)  # Makes the lines smooth

        # Get the dimensions of the widget
        padding = 2
        c_width = self.width()  # Make the final drawing thinner
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

class FoundationDetailsDialog(QDialog):
    """
    A modal dialog with multiple pages to enter or edit details for a foundation type.
    """
    def __init__(self, existing_details: dict = None, parent=None,
                 existing_names: list[str] = None):
        """Initializes the multi-page dialog."""
        super().__init__(parent)
        self.setWindowTitle('Foundation Details')
        self.setModal(True)
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
        self.setGeometry(100, 100, 500, 650)
        self.setMinimumWidth(500)
        self.setMinimumHeight(650)

        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Stacked widget for pages
        self.stacked_widget = QStackedWidget()
        main_layout.addWidget(self.stacked_widget)

        # Data
        self.widgets = {}
        self.group_box = {}
        self.stirrup_canvas = None
        self.stirrup_rows_layout = None
        self.remove_stirrup_button = None
        self.info_popup = InfoPopup(self)
        self.existing_names = existing_names if existing_names is not None else []

        # Redraw debounce
        self.debounce_timer = QTimer(self)
        self.debounce_timer.setInterval(300)  # 300ms delay
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.timeout.connect(self.update_stirrup_drawing)

        # Create pages and add them to the stacked widget
        self.create_footing_page()
        self.create_rsb_page()

        # Connect signals after all widgets have been created
        self.connect_stirrup_redraw_signals()

        # Set initial state
        self.stacked_widget.setCurrentIndex(0)

        # Pre-fill fields if editing existing data
        if existing_details:
            # If editing, remove its own name from the list of names to check against
            current_name = existing_details.get('name', '')
            if current_name in self.existing_names:
                self.existing_names.remove(current_name)
            self.populate_data(existing_details)
            self.update_stirrup_drawing()
        else:
            self.add_stirrup_row()

    def create_footing_page(self):
        """Creates the first page of the form."""
        page = QWidget()
        page.setObjectName('footingPage')
        page.setProperty('class', 'page')
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)

        content_frame = QFrame()
        content_frame.setObjectName('footingPageContent')
        content_layout = QVBoxLayout(content_frame)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        # Name
        label = QLabel('Foundation Type')
        label.setProperty('class', 'subtitle')
        name = QLineEdit()
        name.setProperty('class', 'header-1')
        content_layout.addWidget(name)
        content_layout.addWidget(label)
        self.widgets['name'] = name

        # Image
        footing_img = get_img(resource_path('images/label_1ped.png'), FOOTING_IMAGE_WIDTH, FOOTING_IMAGE_WIDTH)
        content_layout.addWidget(footing_img)
        content_layout.addSpacing(10)

        # --- Create the form widget and the QGridLayout ---
        form_layout = QGridLayout()
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(0)
        form_layout.setColumnStretch(2, 1)  # Allow the column (2) to stretch, keeping other columns fixed

        # Pedestal Per Footing
        ped_per_footing = QSpinBox()
        ped_per_footing.setProperty('class', 'form-value')
        ped_per_footing.setRange(1, 4)
        image_map = {
            '1': resource_path('images/label_1ped.png'),
            '2': resource_path('images/label_2ped.png'),
            '3': resource_path('images/label_3ped.png'),
            '4': resource_path('images/label_4ped.png')
        }
        label = QLabel('Pedestal Per Footing:')
        label.setProperty('class', 'form-label')
        form_layout.addWidget(label, 1, 0, 1, 2)
        form_layout.addWidget(ped_per_footing, 1, 2, 1, 2)
        ped_per_footing.valueChanged.connect(
            lambda value: update_image(str(value), image_map, footing_img,
                                               fallback=resource_path('images/label_0ped.png')))
        self.widgets['n_ped'] = ped_per_footing

        # Total Number of Footing
        n_footing = BlankSpinBox(1, 9_999, 1)
        n_footing.setProperty('class', 'form-value')
        label = QLabel('Total Number of Footing:')
        label.setProperty('class', 'form-label')
        form_layout.addWidget(label, 2, 0, 1, 2)
        form_layout.addWidget(n_footing, 2, 2, 1, 2)
        self.widgets['n_footing'] = n_footing

        # Concrete Cover
        cc = BlankSpinBox(1, 999, 75, suffix=' mm')
        cc.setProperty('class', 'form-value')
        label = QLabel('Concrete Cover:')
        label.setProperty('class', 'form-label')
        form_layout.addWidget(label, 3, 0, 1, 2)
        form_layout.addWidget(cc, 3, 2, 1, 2)
        self.widgets['cc'] = cc

        # Pedestal Width
        ped_width_x = BlankSpinBox(0, 99_999, suffix=' mm')
        ped_width_y = BlankSpinBox(0, 99_999, suffix=' mm')
        ped_width_x.setProperty('class', 'form-value')
        ped_width_y.setProperty('class', 'form-value')
        ped_link_checkbox = LinkSpinboxes(ped_width_x, ped_width_y, 'Keep square')
        label = QLabel('Pedestal Width (Along X)')
        label.setProperty('class', 'form-label')
        form_layout.addWidget(label, 4, 0)
        variable = QLabel('bx')
        variable.setProperty('class', 'variable-label')
        form_layout.addWidget(label, 4, 0)
        form_layout.addWidget(variable, 4, 1)
        form_layout.addWidget(ped_width_x, 4, 2, 1, 2)
        label = QLabel('Pedestal Width (Along Y)')
        label.setProperty('class', 'form-label')
        form_layout.addWidget(label, 5, 0)
        variable = QLabel('by')
        variable.setProperty('class', 'variable-label')
        form_layout.addWidget(label, 5, 0)
        form_layout.addWidget(variable, 5, 1)
        form_layout.addWidget(ped_width_y, 5, 2)
        form_layout.addWidget(ped_link_checkbox, 5, 3)
        self.widgets['bx'] = ped_width_x
        self.widgets['by'] = ped_width_y

        # Pedestal Height
        ped_height = BlankSpinBox(0, 999_999, suffix=' mm')
        ped_height.setProperty('class', 'form-value')
        label = QLabel('Pedestal Height')
        label.setProperty('class', 'form-label')
        form_layout.addWidget(label, 6, 0)
        variable = QLabel('h')
        variable.setProperty('class', 'variable-label')
        form_layout.addWidget(label, 6, 0)
        form_layout.addWidget(variable, 6, 1)
        form_layout.addWidget(ped_height, 6, 2, 1, 2)
        self.widgets['h'] = ped_height

        # Pad Width
        pad_width_x = BlankSpinBox(0, 999_999, suffix=' mm')
        pad_width_y = BlankSpinBox(0, 999_999, suffix=' mm')
        pad_width_x.setProperty('class', 'form-value')
        pad_width_y.setProperty('class', 'form-value')
        pad_link_checkbox = LinkSpinboxes(pad_width_x, pad_width_y, 'Keep square')
        label = QLabel('Pad Width (Along X)')
        label.setProperty('class', 'form-label')
        form_layout.addWidget(label, 7, 0)
        variable = QLabel('Bx')
        variable.setProperty('class', 'variable-label')
        form_layout.addWidget(label, 7, 0)
        form_layout.addWidget(variable, 7, 1)
        form_layout.addWidget(pad_width_x, 7, 2, 1, 2)
        label = QLabel('Pad Width (Along Y)')
        label.setProperty('class', 'form-label')
        form_layout.addWidget(label, 8, 0)
        variable = QLabel('By')
        variable.setProperty('class', 'variable-label')
        form_layout.addWidget(label, 8, 0)
        form_layout.addWidget(variable, 8, 1)
        form_layout.addWidget(pad_width_y, 8, 2)
        form_layout.addWidget(pad_link_checkbox, 8, 3)
        self.widgets['Bx'] = pad_width_x
        self.widgets['By'] = pad_width_y

        # Pad thickness
        pad_thickness = BlankSpinBox(0, 99_999, suffix=' mm')
        pad_thickness.setProperty('class', 'form-value')
        label = QLabel('Pad Thickness')
        label.setProperty('class', 'form-label')
        form_layout.addWidget(label, 9, 0)
        variable = QLabel('t')
        variable.setProperty('class', 'variable-label')
        form_layout.addWidget(label, 9, 0)
        form_layout.addWidget(variable, 9, 1)
        form_layout.addWidget(pad_thickness, 9, 2, 1, 2)
        self.widgets['t'] = pad_thickness

        # Add layouts and widgets to page
        content_layout.addLayout(form_layout)
        content_layout.addStretch()
        horizontal_layout = QHBoxLayout()
        horizontal_layout.setContentsMargins(0, 0, 0, 0)
        horizontal_layout.setSpacing(0)
        horizontal_layout.addWidget(content_frame)
        page_layout.addLayout(horizontal_layout)

        # --- Bottom Navigation ---
        bottom_nav = QFrame()
        bottom_nav.setProperty('class', 'bottom-nav')
        button_layout = QHBoxLayout(bottom_nav)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(0)
        button_layout.addStretch()
        next_button = HoverButton('Next')
        next_button.setProperty('class', 'green-button next-button')
        next_button.clicked.connect(self.go_to_rsb_page)
        button_layout.addWidget(next_button)
        page_layout.addWidget(bottom_nav)

        self.stacked_widget.addWidget(page)

    def create_rsb_page(self) -> None:
        """Builds the UI for the second page (Reinforcement Details)."""
        page = QWidget()
        page.setObjectName('rsbPage')
        page.setProperty('class', 'page')
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)

        # --- A helper function to create each group box to avoid repeating code ---
        def create_top_bot_bar_section(title, image_path, image_width):
            if 'Top Bar' in title:
                group_box = MemoryGroupBox(title)
            else:
                group_box = QGroupBox(title)
            section_layout = QHBoxLayout(group_box)
            section_layout.setContentsMargins(0,0,0,0)
            section_layout.setSpacing(0)

            # Left Image
            image_map = {'True': image_path, 'False': resource_path('images/no_top_bar.png')}
            image_label = get_img(image_map['True'], image_width, image_width)
            section_layout.addWidget(image_label)

            form_layout = QVBoxLayout()
            form_layout.setContentsMargins(0,0,0,0)
            form_layout.setSpacing(3)

            # Row 0: Diameter
            row_1 = QHBoxLayout()
            row_1.setContentsMargins(0, 0, 0, 0)
            row_1.setSpacing(0)
            label = QLabel('Diameter:')
            label.setProperty('class', 'form-label')
            bar_size = QComboBox()
            bar_size.addItems(BAR_DIAMETERS)
            bar_size.setProperty('class', 'form-value')
            size_policy = bar_size.sizePolicy()
            size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
            bar_size.setSizePolicy(size_policy)
            row_1.addWidget(label)
            row_1.addWidget(bar_size, 1)
            form_layout.addLayout(row_1)

            # Row 1: ComboBox
            input_type = QComboBox()
            input_type.addItems(['Quantity', 'Spacing'])
            input_type.setProperty('class', 'form-label')
            form_layout.addWidget(input_type)

            # Row 2: Inputs
            row_3 = QHBoxLayout()
            row_3.setContentsMargins(0, 0, 0, 0)
            row_3.setSpacing(3)
            value_along_x = BlankSpinBox(0, 99_999, suffix=' pcs')
            value_along_y = BlankSpinBox(0, 99_999, suffix=' pcs')
            value_along_x.setProperty('class', 'form-value')
            value_along_y.setProperty('class', 'form-value')
            row_3.addWidget(value_along_x)
            row_3.addWidget(value_along_y)
            form_layout.addLayout(row_3)

            # Row 3: Along X / Y
            row_4 = QHBoxLayout()
            row_4.setContentsMargins(0, 0, 0, 0)
            row_4.setSpacing(0)
            along_x_label = QLabel('Along X')
            along_x_label.setProperty('class', 'subtitle')
            left_along_layout = QHBoxLayout()
            left_along_layout.setContentsMargins(0,0,0,0)
            along_y_label = QLabel('Along Y')
            along_y_label.setProperty('class', 'subtitle')
            right_along_layout = QHBoxLayout()
            right_along_layout.setContentsMargins(0,0,0,0)
            right_along_layout.setSpacing(0)
            link_spinbox = LinkSpinboxes(value_along_x, value_along_y, 'Same for both directions')

            left_along_layout.addWidget(along_x_label)
            left_along_layout.addStretch()
            right_along_layout.addWidget(along_y_label)
            right_along_layout.addStretch()
            right_along_layout.addWidget(link_spinbox)
            row_4.addLayout(left_along_layout)
            row_4.addLayout(right_along_layout)
            form_layout.addLayout(row_4)

            # Add to main layout
            section_layout.addLayout(form_layout)

            # --- Store controls for later data retrieval and manipulation ---
            self.widgets[title] = {
                'Diameter': bar_size,
                'Input Type': input_type,
                'Value Along X': value_along_x,
                'Value Along Y': value_along_y,
            }

            # --- Connections for dynamic UI changes ---
            def update_spinbox_suffix():
                if 'Quantity' in input_type.currentText():
                    value_along_x.setSuffix(' pcs')
                    value_along_y.setSuffix(' pcs')
                else:
                    value_along_x.setSuffix(' mm')
                    value_along_y.setSuffix(' mm')
            input_type.currentTextChanged.connect(update_spinbox_suffix)
            group_box.toggled.connect(lambda checked: update_image(str(checked), image_map, image_label, image_width,
                                                                   fallback=resource_path('images/no_top_bar.png')))
            self.group_box[title] = group_box
            return group_box

        def create_vert_bar_section(image_width):
            title = 'Vertical Bar'
            group_box = QGroupBox(title)
            section_layout = QHBoxLayout(group_box)
            section_layout.setContentsMargins(0,0,0,0)
            section_layout.setSpacing(0)

            # --- Image (Left side) ---
            section_layout.addWidget(get_img(resource_path('images/vert_bar.png'), image_width, image_width))

            # --- Container for the right side controls ---
            form_layout = QFormLayout()
            form_layout.setContentsMargins(0,0,0,0)
            form_layout.setSpacing(3)

            # Row 0: Diameter
            bar_size = QComboBox()
            bar_size.addItems(BAR_DIAMETERS)
            bar_size.setProperty('class', 'form-value')
            size_policy = bar_size.sizePolicy()
            size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
            bar_size.setSizePolicy(size_policy)
            label = QLabel('Diameter:')
            label.setProperty('class', 'form-label')
            form_layout.addRow(label, bar_size)

            # Row 1: Qty
            qty = BlankSpinBox(0, 99_999, suffix=' pcs')
            qty.setProperty('class', 'form-value')
            label = QLabel('Quantity:')
            label.setProperty('class', 'form-label')
            form_layout.addRow(label, qty)

            # Row 2: Hook Calculation
            calculation = QComboBox()
            calculation.addItems(['Automatic', 'Manual'])
            calculation.setProperty('class', 'form-value')
            label = HoverLabel('Hook Calculation:')
            label.setProperty('class', 'form-label')
            label.mouseEntered.connect(self.show_hook_info)
            label.mouseLeft.connect(self.info_popup.hide)
            form_layout.addRow(label, calculation)

            # Row 3: Hook Length (Label)
            hook_length_label = QLabel('Hook Length:')
            hook_length_label.setProperty('class', 'form-label')
            hook_length = BlankSpinBox(0, 99_999, suffix=' mm')
            hook_length.setProperty('class', 'form-value')
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
            self.widgets[title] = {
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
            section_layout.setContentsMargins(0,0,0,0)
            section_layout.setSpacing(0)

            # --- Image (Left side) ---
            image_map = {  # 'None': resource_path('images/perim_bar_0.png'),
                '1': resource_path('images/perim_bar_1.png'),
                '2': resource_path('images/perim_bar_2.png'),
                '3': resource_path('images/perim_bar_3.png'),
                '4': resource_path('images/perim_bar_4.png'),
                '5': resource_path('images/perim_bar_5.png'),
            }
            perim_bar_img = get_img(image_map['1'], image_width, image_width)
            section_layout.addWidget(perim_bar_img)

            # --- Container for the right side controls ---
            form_layout = QFormLayout()
            form_layout.setContentsMargins(0,0,0,0)
            form_layout.setSpacing(3)

            # Row 0: Diameter
            bar_size = QComboBox()
            bar_size.addItems(BAR_DIAMETERS)
            bar_size.setProperty('class', 'form-value')
            diameter_label = QLabel('Diameter:')
            diameter_label.setProperty('class', 'form-label')
            form_layout.addRow(diameter_label, bar_size)

            # Row 1: Layers
            layers = QComboBox()
            layers.addItems(['1', '2', '3', '4', '5'])  # Add None if needed
            layers.setProperty('class', 'form-value')
            size_policy = layers.sizePolicy()
            size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
            layers.setSizePolicy(size_policy)
            layers_label = QLabel('Layers:')
            layers_label.setProperty('class', 'form-label')
            form_layout.addRow(layers_label, layers)

            # --- Add the right side to the main layout ---
            section_layout.addLayout(form_layout)
            self.widgets[title] = {'Diameter': bar_size,
                                   'Layers': layers}

            # noinspection PyUnresolvedReferences
            layers.currentTextChanged.connect(
                lambda selected_text: update_image(selected_text, image_map, perim_bar_img, image_width,
                                                   fallback=resource_path('images/perim_bar_0.png')))
            group_box.setChecked(False)
            self.group_box[title] = group_box
            return group_box

        def create_stirrup_group_box(image_width):
            title = 'Stirrups'
            group_box = MemoryGroupBox(title)
            main_layout = QVBoxLayout(group_box)
            main_layout.setContentsMargins(0, 0, 0, 0)
            main_layout.setSpacing(0)

            # Initialize
            self.widgets[title] = {'Types': []}

            # Spacing section
            spacing_layout = QHBoxLayout()
            spacing_layout.setContentsMargins(0, 0, 0, 0)
            spacing_layout.setSpacing(0)

            # --- Image (Left side) ---
            canvas_layout = QVBoxLayout()
            canvas_layout.setContentsMargins(0, 0, 0, 0)
            canvas_layout.setSpacing(0)
            label = HoverLabel('Spacing Per Bundle')
            label.setObjectName('rsbPageSpacingHeader')
            label.mouseEntered.connect(self.show_spacing_header_info)
            label.mouseLeft.connect(self.info_popup.hide)
            canvas_layout.addWidget(label)
            self.stirrup_canvas = DrawStirrup(image_width)
            self.stirrup_canvas.setProperty('class', 'drawing-canvas')
            canvas_layout.addWidget(self.stirrup_canvas)
            spacing_layout.addLayout(canvas_layout)

            # --- Container for the right side controls ---
            form_layout = QFormLayout()
            form_layout.setContentsMargins(0, 0, 0, 0)
            form_layout.setSpacing(3)

            # Row 0: Start From
            extent_label = HoverLabel('Start From:')
            extent_label.setProperty('class', 'form-label')
            start_from = QComboBox()
            start_from.addItems(['From Face of Pad', 'From Bottom Bar', 'From Top'])
            start_from.setProperty('class', 'form-value')
            extent_label.mouseEntered.connect(self.show_spacing_extent_info)
            extent_label.mouseLeft.connect(self.info_popup.hide)
            form_layout.addRow(extent_label, start_from)

            # Row 1: Spacing
            spacing_label = HoverLabel('Spacing:')
            spacing_label.setProperty('class', 'form-label')
            spacing = QTextEdit()
            spacing.setProperty('class', 'form-value')
            spacing.setPlaceholderText('Example: 1@50, 5@80, rest@100')
            spacing_label.mouseEntered.connect(self.show_spacing_info)
            spacing_label.mouseLeft.connect(self.info_popup.hide)
            spacing.textChanged.connect(self.debounce_timer.start)
            form_layout.addRow(spacing_label, spacing)
            spacing_layout.addLayout(form_layout)

            # Store
            self.widgets[title]['Extent'] = start_from
            self.widgets[title]['Spacing'] = spacing

            # Spacing section
            bundle_layout = QVBoxLayout()
            bundle_layout.setContentsMargins(0, 0, 0, 0)
            bundle_layout.setSpacing(0)

            # --- Button Layout for adding/removing rows ---
            title_row_container = QFrame()
            title_row_container.setProperty('class', 'title-row')
            title_row_layout = QHBoxLayout(title_row_container)
            title_row_layout.setContentsMargins(0, 0, 0, 0)
            title_row_layout.setSpacing(3)
            label = HoverLabel('Stirrup Bundle')
            label.setObjectName('rsbPageBundleHeader')
            label.mouseEntered.connect(self.show_bundle_info)
            label.mouseLeft.connect(self.info_popup.hide)
            add_button = HoverButton('+')
            add_button.setProperty('class', 'green-button add-button')
            self.remove_stirrup_button = HoverButton('-')
            self.remove_stirrup_button.setProperty('class', 'red-button remove-button')
            add_button.clicked.connect(self.add_stirrup_row)
            self.remove_stirrup_button.clicked.connect(self.remove_stirrup_row)
            title_row_layout.addWidget(label)
            title_row_layout.addStretch()
            title_row_layout.addWidget(add_button)
            title_row_layout.addWidget(self.remove_stirrup_button)

            # --- Container for dynamic stirrup rows ---
            self.stirrup_rows_layout = QVBoxLayout()
            self.stirrup_rows_layout.setContentsMargins(0, 0, 0, 0)
            self.stirrup_rows_layout.setSpacing(3)

            stirrup_rows_container = QFrame()
            stirrup_rows_container.setProperty('class', 'list-container')
            container = QVBoxLayout(stirrup_rows_container)
            container.setContentsMargins(0, 0, 0, 0)
            container.setSpacing(0)
            container.addWidget(title_row_container)
            container.addLayout(self.stirrup_rows_layout)

            bundle_layout.addWidget(stirrup_rows_container)
            bundle_layout.addStretch()  # Pushes rows to the top

            # --- Add the first, initial row ---
            # self.add_stirrup_row()

            # --- COMBINE SECTION ---
            main_layout.addLayout(spacing_layout)
            main_layout.addSpacing(20)
            main_layout.addLayout(bundle_layout)

            self.group_box[title] = group_box
            return group_box

        # --- Create and add the group boxes ---
        top_bar_box = create_top_bot_bar_section('Top Bar', resource_path('images/top_bar.png'), RSB_IMAGE_WIDTH)
        bot_bar_box = create_top_bot_bar_section('Bottom Bar', resource_path('images/bot_bar.png'), RSB_IMAGE_WIDTH)
        vert_bar_box = create_vert_bar_section(RSB_IMAGE_WIDTH)
        perim_bar_box = create_perim_bar_section(RSB_IMAGE_WIDTH)
        stirrup_group_box = create_stirrup_group_box(RSB_IMAGE_WIDTH)

        def create_hline():
            separator = QFrame()
            separator.setProperty('class', 'separator')
            separator.setFrameShape(QFrame.Shape.HLine)
            # separator.setFrameShadow(QFrame.Shadow.Sunken)
            return separator

        separator_space = 30
        scroll_content = QFrame()
        scroll_content.setObjectName('rsbScrollContent')
        v_layout = QVBoxLayout(scroll_content)
        v_layout.setContentsMargins(0, 0, 0, 0)
        v_layout.setSpacing(0)
        v_layout.addWidget(top_bar_box)
        v_layout.addWidget(create_hline())
        v_layout.addSpacing(separator_space)
        v_layout.addWidget(bot_bar_box)
        v_layout.addWidget(create_hline())
        v_layout.addSpacing(separator_space)
        v_layout.addWidget(vert_bar_box)
        v_layout.addWidget(create_hline())
        v_layout.addSpacing(separator_space)
        v_layout.addWidget(perim_bar_box)
        v_layout.addWidget(create_hline())
        v_layout.addSpacing(separator_space)
        v_layout.addWidget(stirrup_group_box)

        # center_layout = QHBoxLayout(scroll_content)
        # center_layout.setContentsMargins(0, 0, 0, 0)
        # center_layout.setSpacing(0)
        # center_layout.addStretch()
        # center_layout.addLayout(v_layout)
        # center_layout.addStretch()

        # Connection to redraw
        self.connect_stirrup_redraw_signals()

        scroll_area = make_scrollable(scroll_content, True)
        scroll_area.setProperty('class', 'scroll-bar')
        scroll_area.setObjectName('scrollBar')
        scroll_area.setAlignment(Qt.AlignmentFlag.AlignCenter)
        page_layout.addWidget(scroll_area)  # Add the scrollable part

        # --- Bottom Navigation ---
        bottom_nav = QFrame()
        bottom_nav.setProperty('class', 'bottom-nav')
        button_layout = QHBoxLayout(bottom_nav)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(0)
        back_button = HoverButton('Back')
        back_button.setProperty('class', 'red-button back-button')
        back_button.clicked.connect(self.go_to_footing_page)
        button_layout.addWidget(back_button)
        button_layout.addStretch()
        next_button = HoverButton('Save')
        next_button.setProperty('class', 'green-button next-button')
        next_button.clicked.connect(self.save_and_accept)
        button_layout.addWidget(next_button)
        page_layout.addWidget(bottom_nav)

        self.stacked_widget.addWidget(page)

    def go_to_footing_page(self):
        """Switches the stacked widget to the previous page."""
        self.stacked_widget.setCurrentIndex(0)
        self.setFocus()  # Remove focus

    def go_to_rsb_page(self):
        """Switches the stacked widget to the next page."""
        if self.validate_footing_page():
            self.stacked_widget.setCurrentIndex(1)
        self.setFocus()  # Remove focus

    def validate_footing_page(self) -> bool:
        """
        Validates all inputs on the footing page with stricter rules.
        Returns True if valid.
        """
        if DEBUG_MODE:
            return True

        errors = []
        widgets_with_errors = []

        # --- Get all widgets ---
        name_widget = self.widgets['name']
        n_footing_widget = self.widgets['n_footing']
        cc_widget = self.widgets['cc']
        bx_widget = self.widgets['bx']
        by_widget = self.widgets['by']
        h_widget = self.widgets['h']
        Bx_widget = self.widgets['Bx']
        By_widget = self.widgets['By']
        t_widget = self.widgets['t']

        # --- List of widgets to check for emptiness and reset styles ---
        all_widgets = [
            name_widget, n_footing_widget, cc_widget, bx_widget, by_widget,
            h_widget, Bx_widget, By_widget, t_widget
        ]
        for widget in all_widgets:
            style_invalid_input(widget, True)

        # 1. All visible input fields must be filled.
        is_empty_found = False
        for widget in all_widgets:
            if is_widget_empty(widget):
                widgets_with_errors.append(widget)
                is_empty_found = True
        if is_empty_found:
            errors.append('• All input fields must be filled.')

        # --- Get numeric values for comparison checks ---
        name = name_widget.text().strip()
        cc = cc_widget.value()
        bx = bx_widget.value()
        by = by_widget.value()
        h = h_widget.value()
        Bx = Bx_widget.value()
        By = By_widget.value()
        t = t_widget.value()

        # 2. Name should be unique.
        if name in self.existing_names:
            errors.append(f'• Foundation Type name {name} already exists.')
            widgets_with_errors.append(name_widget)

        # 3. Pedestal Width and pad thickness > 2 * concrete cover.
        if not is_widget_empty(bx_widget) and not is_widget_empty(cc_widget) and bx <= 2 * cc:
            errors.append(f'• Pedestal Width X ({bx}mm) must be > 2 × Concrete Cover ({2 * cc}mm).')
            widgets_with_errors.extend([bx_widget, cc_widget])
        if not is_widget_empty(by_widget) and not is_widget_empty(cc_widget) and by <= 2 * cc:
            errors.append(f'• Pedestal Width Y ({by}mm) must be > 2 × Concrete Cover ({2 * cc}mm).')
            widgets_with_errors.extend([by_widget, cc_widget])
        if not is_widget_empty(t_widget) and not is_widget_empty(cc_widget) and t <= 2 * cc:
            errors.append(f'• Pad Thickness ({t}mm) must be > 2 × Concrete Cover ({2 * cc}mm).')
            widgets_with_errors.extend([t_widget, cc_widget])

        # 4. Pedestal Height > concrete cover.
        if not is_widget_empty(h_widget) and not is_widget_empty(cc_widget) and h <= cc:
            errors.append(f'• Pedestal Height ({h}mm) must be > Concrete Cover ({cc}mm).')
            widgets_with_errors.extend([h_widget, cc_widget])

        # 5. Pad Width > (concrete_cover * 2 + ped_width)
        if not is_widget_empty(Bx_widget) and not is_widget_empty(bx_widget) and not is_widget_empty(cc_widget) and Bx <= (bx + 2 * cc):
            errors.append(f'• Pad Width X ({Bx}mm) must be > Pedestal Width X + 2 × Cover ({bx + 2 * cc}mm).')
            widgets_with_errors.extend([Bx_widget, bx_widget, cc_widget])
        if not is_widget_empty(By_widget) and not is_widget_empty(by_widget) and not is_widget_empty(cc_widget) and By <= (by + 2 * cc):
            errors.append(f'• Pad Width Y ({By}mm) must be > Pedestal Width Y + 2 × Cover ({by + 2 * cc}mm).')
            widgets_with_errors.extend([By_widget, by_widget, cc_widget])

        if errors:
            for widget in set(widgets_with_errors):
                style_invalid_input(widget, False)
            msg_box = QMessageBox(self)
            msg_box.setObjectName('warningMessageBox')
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setWindowTitle('Invalid Input on Footing Page')
            msg_box.setText('Please correct the following errors:')
            msg_box.setInformativeText('\n'.join(errors))
            msg_box.exec()
            return False

        return True

    def validate_rsb_page(self) -> bool:
        """Validates all visible inputs on the RSB page with stricter rules. Returns True if valid."""
        if DEBUG_MODE:
            return True

        errors = []
        widgets_with_errors = []

        # --- Get values from footing page for context ---
        cc = self.widgets['cc'].value()
        bx = self.widgets['bx'].value()
        by = self.widgets['by'].value()
        Bx = self.widgets['Bx'].value()
        By = self.widgets['By'].value()

        # --- Reset all styles first ---
        for section_name in self.group_box:
            section_widgets = self.widgets.get(section_name, {})
            if isinstance(section_widgets, dict):
                for widget in section_widgets.values():
                    if isinstance(widget, QWidget) and widget.isWidgetType():
                        style_invalid_input(widget, True)
            elif isinstance(section_widgets, list): # For stirrup types
                 for row_dict in section_widgets:
                     for widget in row_dict.values():
                         if isinstance(widget, QWidget) and widget.isWidgetType():
                             style_invalid_input(widget, True)


        # --- Iterate through group boxes for validation ---
        for section_name, group_box in self.group_box.items():
            if group_box.isCheckable() and not group_box.isChecked():
                continue  # Skip disabled sections

            # --- A. Check for general emptiness in visible fields ---
            is_empty_found_in_section = False
            widgets_in_section = self.widgets.get(section_name, {})
            # This handles both flat dicts (Vertical Bar) and nested dicts (Top Bar)
            widgets_to_check = list(widgets_in_section.values()) if isinstance(widgets_in_section, dict) else []

            for widget in widgets_to_check:
                if isinstance(widget, QWidget) and widget.isWidgetType() and widget.isVisible():
                    if is_widget_empty(widget):
                        widgets_with_errors.append(widget)
                        is_empty_found_in_section = True
            if is_empty_found_in_section:
                 errors.append(f'• A required field in the {section_name} section is empty.')


            # --- B. Specific validation rules ---
            # 1. Top/Bottom Bar Validation
            if section_name in ['Top Bar', 'Bottom Bar']:
                bar_widgets = self.widgets[section_name]
                val_x_widget = bar_widgets['Value Along X']
                val_y_widget = bar_widgets['Value Along Y']
                clear_width_x = Bx - 2 * cc
                clear_width_y = By - 2 * cc
                # Bars along Y are spaced across Bx
                if not is_widget_empty(val_x_widget) and val_x_widget.value() >= clear_width_x:
                    errors.append(f'• {section_name} (Along Y): Spacing ({val_x_widget.value()}mm) must be less than clear pad width X ({clear_width_x}mm).')
                    widgets_with_errors.extend([val_x_widget, self.widgets['Bx'], self.widgets['cc']])
                # Bars along X are spaced across By
                if not is_widget_empty(val_y_widget) and val_y_widget.value() >= clear_width_y:
                    errors.append(f'• {section_name} (Along X): Spacing ({val_y_widget.value()}mm) must be less than clear pad width Y ({clear_width_y}mm).')
                    widgets_with_errors.extend([val_y_widget, self.widgets['By'], self.widgets['cc']])

            # 2. Vertical Bar Validation
            elif section_name == 'Vertical Bar':
                vert_widgets = self.widgets[section_name]
                qty_widget = vert_widgets['Quantity']
                hook_calc = vert_widgets['Hook Calculation'].currentText()
                hook_len_widget = vert_widgets['Hook Length']

                pedestal_core_perimeter = (bx - 2 * cc) * 2 + (by - 2 * cc) * 2
                if not is_widget_empty(qty_widget) and qty_widget.value() >= pedestal_core_perimeter:
                    errors.append(f'• Vertical Bar quantity ({qty_widget.value()}) is high; it must be less than the pedestal core perimeter ({pedestal_core_perimeter:.0f}mm).')
                    widgets_with_errors.extend([qty_widget, self.widgets['bx'], self.widgets['by'], self.widgets['cc']])

                if hook_calc == 'Manual':
                    clear_pad_width_min = min(Bx, By) - 2 * cc
                    if not is_widget_empty(hook_len_widget) and hook_len_widget.value() >= clear_pad_width_min / 2:
                        errors.append(f'• Manual Hook Length ({hook_len_widget.value()}mm) must be less than half the clear pad width ({clear_pad_width_min / 2:.0f}mm).')
                        widgets_with_errors.extend([hook_len_widget, self.widgets['Bx'], self.widgets['By'], self.widgets['cc']])

            # 3. Stirrups Validation
            elif section_name == 'Stirrups':
                spacing_widget = self.widgets['Stirrups']['Spacing']
                spacing_text = spacing_widget.toPlainText().strip()
                if not spacing_text:
                     errors.append('• Stirrup Spacing field cannot be empty.')
                     widgets_with_errors.append(spacing_widget)
                else:
                    try:
                        parse_spacing_string(spacing_text)
                    except (ValueError, TypeError) as e:
                        errors.append(f'• Invalid Stirrup Spacing format: {e}')
                        widgets_with_errors.append(spacing_widget)

        if errors:
            for widget in set(widgets_with_errors):
                style_invalid_input(widget, False)
            msg_box = QMessageBox(self)
            msg_box.setObjectName('warningMessageBox')
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setWindowTitle('Invalid Input on Reinforcement Page')
            msg_box.setText('Please correct the following errors:')
            msg_box.setInformativeText('\n'.join(errors))
            msg_box.exec()
            return False

        return True

    def save_and_accept(self):
        """Runs validation before accepting the dialog."""
        if self.validate_footing_page() and self.validate_rsb_page():
            self.accept()

    def update_stirrup_drawing(self) -> None:
        """Triggers a repaint of the stirrup drawing canvas with current input values."""
        if hasattr(self, 'stirrup_canvas'):  # Check if canvas exists
            # Create a dictionary with the required footing dimension widgets
            footing_details = {
                'Pedestal Height': self.widgets['h'],
                'Pedestal Width (Along X)': self.widgets['bx'],
                'Pad Thickness': self.widgets['t'],
                'Concrete Cover': self.widgets['cc']
            }
            self.stirrup_canvas.update_dimensions(
                footing_details,
                self.widgets['Stirrups']['Extent'],
                self.widgets['Stirrups']['Spacing'],
                self.widgets['Bottom Bar']['Diameter'],
                self.widgets['Vertical Bar']['Diameter']
            )

    def connect_stirrup_redraw_signals(self):
        """Connects all widgets that affect the stirrup drawing to the redraw logic."""
        # Widgets that affect dimensions
        dimension_widgets = [
            self.widgets['h'],
            self.widgets['bx'],
            self.widgets['t'],
            self.widgets['cc'],
        ]
        # Widgets that affect rebar sizes
        rebar_widgets = [
            self.widgets['Bottom Bar']['Diameter'],
            self.widgets['Vertical Bar']['Diameter'],
            self.widgets['Stirrups']['Extent']
        ]

        for widget in dimension_widgets:
            widget.valueChanged.connect(self.update_stirrup_drawing)

        for widget in rebar_widgets:
            widget.currentTextChanged.connect(self.update_stirrup_drawing)

    def disconnect_stirrup_redraw_signals(self):
        """Disconnects signals that trigger stirrup redraws to prevent signal storms."""
        dimension_widgets = [
            self.widgets['h'],
            self.widgets['bx'],
            self.widgets['t'],
            self.widgets['cc'],
        ]
        rebar_widgets = [
            self.widgets['Bottom Bar']['Diameter'],
            self.widgets['Vertical Bar']['Diameter'],
            self.widgets['Stirrups']['Extent']
        ]

        for widget in dimension_widgets:
            try:
                widget.valueChanged.disconnect(self.update_stirrup_drawing)
            except TypeError:
                pass  # Signal was not connected, so we can ignore the error

        for widget in rebar_widgets:
            try:
                widget.currentTextChanged.disconnect(self.update_stirrup_drawing)
            except TypeError:
                pass  # Signal was not connected, so we can ignore the error

    def show_hook_info(self) -> None:
        """Displays an informational popup for the hook calculation method."""
        # NOTE: You should update the text to reflect the actual standard you are using.
        # For example, ACI 318 or a local building code.
        info_text = (
            '<b>Hook Calculation Method</b>'
            '<ul><li><b>Automatic:</b> Calculates the required hook length '
            'based on ACI 318-25 Table 25.3.1 Standard 90° hook geometry for '
            'development of deformed bars in tension (<i>12d</i><sub>b</sub>) </li>'
            '<li><b>Manual:</b> Allows you to enter a custom, pre-calculated '
            'length for the hook.</li></ul>'
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
Think of it as designing a 'kit' of stirrups that gets repeated along the height of the pedestal."""
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

    def populate_data(self, details: dict):
        """Fills the form fields with existing data for editing."""
        self.disconnect_stirrup_redraw_signals()

        # Page 1 (Footing Dimensions)
        self.widgets['name'].setText(details.get('name', ''))
        self.widgets['n_footing'].setValue(details.get('n_footing', 1))
        self.widgets['n_ped'].setValue(details.get('n_ped', 1))
        self.widgets['cc'].setValue(details.get('cc', 75))
        self.widgets['bx'].setValue(details.get('bx', 0))
        self.widgets['by'].setValue(details.get('by', 0))
        self.widgets['h'].setValue(details.get('h', 0))
        self.widgets['Bx'].setValue(details.get('Bx', 0))
        self.widgets['By'].setValue(details.get('By', 0))
        self.widgets['t'].setValue(details.get('t', 0))

        # Page 2 (Reinforcement Details)
        sections = ['Top Bar', 'Bottom Bar', 'Vertical Bar', 'Perimeter Bar', 'Stirrups']
        for section_name in sections:
            section_data = details.get(section_name, {})
            if not section_data:
                continue

            # Handle GroupBox check state
            if self.group_box[section_name].isCheckable():
                is_enabled = section_data.get('Enabled', False)
                self.group_box[section_name].setChecked(is_enabled)
                # If the section is not enabled, skip populating its widgets
                if not is_enabled:
                    continue

            if section_name in ['Top Bar', 'Bottom Bar']:
                widget = self.widgets[section_name]
                widget['Diameter'].setCurrentText(section_data.get('Diameter', ''))
                widget['Input Type'].setCurrentText(section_data.get('Input Type', 'Quantity'))
                widget['Value Along X'].setValue(section_data.get('Value Along X', 0))
                widget['Value Along Y'].setValue(section_data.get('Value Along Y', 0))

            elif section_name == 'Vertical Bar':
                widget = self.widgets[section_name]
                widget['Diameter'].setCurrentText(section_data.get('Diameter', ''))
                widget['Quantity'].setValue(section_data.get('Quantity', 0))
                widget['Hook Calculation'].setCurrentText(section_data.get('Hook Calculation', 'Automatic'))
                widget['Hook Length'].setValue(section_data.get('Hook Length', 0))

            elif section_name == 'Perimeter Bar':
                widget = self.widgets[section_name]
                widget['Diameter'].setCurrentText(section_data.get('Diameter', ''))
                widget['Layers'].setCurrentText(str(section_data.get('Layers', '1')))

            elif section_name == 'Stirrups':
                self.widgets['Stirrups']['Extent'].setCurrentText(section_data.get('Extent', 'From Face of Pad'))
                self.widgets['Stirrups']['Spacing'].setPlainText(section_data.get('Spacing', ''))

                # Add and populate new rows from saved data
                saved_stirrup_types = section_data.get('Types', [])
                for stirrup_type_data in saved_stirrup_types:
                    self.add_stirrup_row()
                    row_widgets = self.widgets['Stirrups']['Types'][-1]
                    row_widgets['Type'].setCurrentText(stirrup_type_data.get('Type', 'Outer'))
                    row_widgets['Diameter'].setCurrentText(stirrup_type_data.get('Diameter', ''))
                    row_widgets['a_input'].setValue(stirrup_type_data.get('a_input', 0))

                # After populating, ensure at least one row exists.
                # This handles cases where a user saved an item with no stirrups.
                if not self.widgets['Stirrups']['Types']:
                    self.add_stirrup_row()

        # Reconnect signals
        self.connect_stirrup_redraw_signals()

    def get_data(self) -> dict:
        """Returns all entered data from both pages as a dictionary."""
        data = {
            # Page 1
            'name': self.widgets['name'].text(),
            'n_footing': self.widgets['n_footing'].value(),
            'n_ped': self.widgets['n_ped'].value(),
            'cc': self.widgets['cc'].value(),
            'bx': self.widgets['bx'].value(),
            'by': self.widgets['by'].value(),
            'h': self.widgets['h'].value(),
            'Bx': self.widgets['Bx'].value(),
            'By': self.widgets['By'].value(),
            't': self.widgets['t'].value(),

            # Page 2
            'Top Bar': {
                'Enabled': self.group_box['Top Bar'].isChecked(),
                'Diameter': self.widgets['Top Bar']['Diameter'].currentText(),
                'Input Type': self.widgets['Top Bar']['Input Type'].currentText(),
                'Value Along X': self.widgets['Top Bar']['Value Along X'].value(),
                'Value Along Y': self.widgets['Top Bar']['Value Along Y'].value(),
            },
            'Bottom Bar': {
                'Enabled': True,
                'Diameter': self.widgets['Bottom Bar']['Diameter'].currentText(),
                'Input Type': self.widgets['Bottom Bar']['Input Type'].currentText(),
                'Value Along X': self.widgets['Bottom Bar']['Value Along X'].value(),
                'Value Along Y': self.widgets['Bottom Bar']['Value Along Y'].value(),
            },
            'Vertical Bar': {
                'Enabled': True,
                'Diameter': self.widgets['Vertical Bar']['Diameter'].currentText(),
                'Quantity': self.widgets['Vertical Bar']['Quantity'].value(),
                'Hook Calculation': self.widgets['Vertical Bar']['Hook Calculation'].currentText(),
                'Hook Length': self.widgets['Vertical Bar']['Hook Length'].value(),
            },
            'Perimeter Bar': {
                'Enabled': self.group_box['Perimeter Bar'].isChecked(),
                'Diameter': self.widgets['Perimeter Bar']['Diameter'].currentText(),
                'Layers': self.widgets['Perimeter Bar']['Layers'].currentText(),
            },
            'Stirrups': {
                'Enabled': self.group_box['Stirrups'].isChecked(),
                'Extent': self.widgets['Stirrups']['Extent'].currentText(),
                'Spacing': self.widgets['Stirrups']['Spacing'].toPlainText(),
                'Quantity': self.stirrup_canvas.get_qty(),
                'Types': [
                    {
                        'Type': row['Type'].currentText(),
                        'Diameter': row['Diameter'].currentText(),
                        'a_input': row['a_input'].value()
                    }
                    for row in self.widgets['Stirrups']['Types']
                ]
            }
        }
        if data['Perimeter Bar']['Enabled']:
            data['Perimeter Bar']['Layers'] = int(data['Perimeter Bar']['Layers'])
        return data


    def add_stirrup_row(self) -> None:
        """Creates and adds a new UI row for defining a stirrup type."""
        # --- Main container for the row ---
        row_widget = QFrame()
        row_widget.setProperty('class', 'stirrup-row')
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(0)

        # --- Image (Left) ---
        image_label = get_img(resource_path('images/stirrup_outer.png'), STIRRUP_ROW_IMAGE_WIDTH, STIRRUP_ROW_IMAGE_WIDTH)
        row_layout.addWidget(image_label)

        image_map = {
            'Outer': resource_path('images/stirrup_outer.png'),
            'Diamond': resource_path('images/stirrup_diamond.png'),
            'Tall': resource_path('images/stirrup_tall.png'),
            'Wide': resource_path('images/stirrup_wide.png'),
            'Octagon': resource_path('images/stirrup_octagon.png'),
            'Vertical': resource_path('images/stirrup_flat_tall.png'),
            'Horizontal': resource_path('images/stirrup_flat_wide.png'),
        }

        # --- Form (Right) ---
        form_layout = QFormLayout()
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(3)
        type_combo = QComboBox()
        type_combo.addItems(image_map.keys())
        type_combo.setProperty('class', 'form-value')
        size_policy = type_combo.sizePolicy()
        size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
        type_combo.setSizePolicy(size_policy)

        dia_combo = QComboBox()
        dia_combo.addItems(BAR_DIAMETERS_FOR_STIRRUPS)
        dia_combo.setProperty('class', 'form-value')

        a_label = QLabel('a:')
        a_label.setProperty('class', 'form-label')
        a_input = BlankSpinBox(0, 99_999, suffix=' mm')
        a_input.setProperty('class', 'form-value')

        label = QLabel('Type:')
        label.setProperty('class', 'form-label')
        form_layout.addRow(label, type_combo)
        label = QLabel('Diameter:')
        label.setProperty('class', 'form-label')
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
        self.widgets['Stirrups']['Types'].append(row_widgets)

        # --- Connections ---
        def update_stirrup_row_visibility(selected_text: str, widgets: dict[str, Any],
                                          stirrup_type_image_map: dict[str, str]) -> None:
            """
            Updates a stirrup row's image and the visibility of its 'a' input field.

            Args:
                selected_text: The selected stirrup type from the combo box.
                widgets: A dictionary of the widgets in that specific row.
                stirrup_type_image_map: A dictionary mapping stirrup types to image paths.
            """
            update_image(selected_text, stirrup_type_image_map, widgets['Image'], STIRRUP_ROW_IMAGE_WIDTH,
                         fallback=resource_path('images/stirrup_none.png'))

            # Update visibility of 'a' input
            is_visible = selected_text in ['Tall', 'Wide', 'Octagon']
            widgets['a_label'].setVisible(is_visible)
            widgets['a_input'].setVisible(is_visible)

        # noinspection PyUnresolvedReferences
        type_combo.currentTextChanged.connect(
            lambda text: update_stirrup_row_visibility(text, row_widgets, image_map)
        )

        # --- Set initial state ---
        update_stirrup_row_visibility(type_combo.currentText(), row_widgets, image_map)

        # --- Add to the main container ---
        self.stirrup_rows_layout.addWidget(row_widget)
        self.update_remove_button_state()

    def update_remove_button_state(self) -> None:
        """Enables or disables the 'remove stirrup row' button based on the row count."""
        self.remove_stirrup_button.setEnabled(len(self.widgets['Stirrups']['Types']) > 1)

    def remove_stirrup_row(self) -> None:
        """Removes the last stirrup definition row from the UI."""
        if len(self.widgets['Stirrups']['Types']) > 1:  # Keep at least one row
            widgets_to_remove = self.widgets['Stirrups']['Types'].pop()
            widgets_to_remove['Row'].deleteLater()  # Safely delete the widget

        self.update_remove_button_state()

    def keyPressEvent(self, event):
        """
        Handles key press events for the main window.
        """
        # If the Escape key is pressed, set focus to the main window.
        if event.key() == Qt.Key.Key_Escape:
            self.setFocus()
        else:
            # Otherwise, let the default event handling proceed
            super().keyPressEvent(event)

class FoundationItem(QFrame):
    """A custom widget representing a single item in the foundation list."""
    edit_requested = Signal(object)
    remove_requested = Signal(object)
    selected = Signal(object)

    def __init__(self, data: dict, parent=None) -> None:
        """Initializes the foundation item widget."""
        super().__init__(parent)
        self.data = data
        self.setProperty('class', 'list-item')
        self._is_selected = False
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.label = QLabel(self.data.get('name', 'Unnamed'))
        layout.addWidget(self.label)
        layout.addStretch(1)

        # --- Edit Button (Icon) ---
        self.edit_button = HoverButton('')
        self.edit_button.setProperty('class', 'edit-button') # Use your yellow class
        self.edit_button.setToolTip('Edit Foundation')
        self.edit_button.clicked.connect(lambda: self.edit_requested.emit(self))
        layout.addWidget(self.edit_button)

        self.remove_button = HoverButton('')
        self.remove_button.setProperty('class', 'trash-button')
        self.remove_button.clicked.connect(lambda: self.remove_requested.emit(self))
        layout.addWidget(self.remove_button)

        # --- Hide buttons initially ---
        self.edit_button.hide()
        self.remove_button.hide()

    def paintEvent(self, event: QPaintEvent) -> None:
        """
        This is the magic method that allows the widget to be styled
        using QSS for background-color, border, etc.
        """
        opt = QStyleOption()
        opt.initFrom(self)
        painter = QPainter(self)
        self.style().drawPrimitive(QStyle.PrimitiveElement.PE_Widget, opt, painter, self)

    def enterEvent(self, event: QEvent) -> None:
        """Show buttons when the mouse enters the widget."""
        self.edit_button.show()
        self.remove_button.show()
        super().enterEvent(event)

    def leaveEvent(self, event: QEvent) -> None:
        """Hide buttons when the mouse leaves the widget."""
        self.edit_button.hide()
        self.remove_button.hide()
        super().leaveEvent(event)

    def mousePressEvent(self, event) -> None:
        """Emit a signal when the item is clicked."""
        if not self._is_selected:
            self.selected.emit(self)
        super().mousePressEvent(event)

    def select(self):
        """Sets the visual state to selected."""
        self._is_selected = True
        self.setProperty('class', 'list-item list-item-selected')
        self.style().polish(self)

    def deselect(self):
        """Sets the visual state to de-selected."""
        self._is_selected = False
        self.setProperty('class', 'list-item')
        self.style().polish(self)

    def update_details(self, new_details: dict):
        """Updates the item's data and refreshes the label."""
        self.data = new_details
        self.label.setText(self.data.get('name', 'Unnamed'))

class MultiPageApp(QMainWindow):
    def __init__(self) -> None:
        """Initializes the main application window and its components."""
        super().__init__()

        self.setWindowTitle('Cutting List')
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
        self.setGeometry(50, 50, 750, 550)
        self.setMinimumWidth(750)
        self.setMinimumHeight(500)

        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        # Initialize
        self.scroll_layout = None
        self.detail_area_stack = None
        self.detail_widgets = {}
        self.current_item = None
        self.detail_stirrup_types_layout = None
        self.market_lengths_checkboxes = None
        self.market_lengths_grid = None
        self.current_market_lengths = list(MARKET_LENGTHS)

        self.create_foundation_entry_page()
        self.create_market_lengths_page()

        if DEBUG_MODE:
            self.prefill_debug_data()

        self.setFocus()  # Remove focus on first open

    def create_foundation_entry_page(self) -> None:
        """Builds the UI with a master-detail layout."""
        page = QWidget()
        page.setObjectName('foundationPage')
        page.setProperty('class', 'page')
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)

        # --- Main Horizontal Layout ---
        main_horizontal_layout = QHBoxLayout()
        main_horizontal_layout.setContentsMargins(0, 0, 0, 0)
        main_horizontal_layout.setSpacing(0)

        # --- Left Panel (Master View - The List) ---
        panel = QFrame()
        panel.setProperty('class', 'panel')
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(0, 0, 0, 0)
        panel_layout.setSpacing(0)

        title_row_container = QFrame()
        title_row_container.setProperty('class', 'title-row')
        title_row_layout = QHBoxLayout(title_row_container)
        title_row_layout.setContentsMargins(0, 0, 0, 0)
        title_row_layout.setSpacing(0)
        label = QLabel('Foundation Types')
        label.setProperty('class', 'header-2')
        title_row_layout.addWidget(label)
        title_row_layout.addStretch()
        add_button = HoverButton('+')
        add_button.setProperty('class', 'green-button add-button')
        add_button.clicked.connect(self.add_foundation_item)
        title_row_layout.addWidget(add_button)
        panel_layout.addWidget(title_row_container)

        scroll_content = QWidget()
        scroll_content.setProperty('class', 'scroll-content')
        self.scroll_layout = QVBoxLayout(scroll_content)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll_layout.setSpacing(0)
        self.scroll_layout.addStretch(1)
        scroll_area = make_scrollable(scroll_content)
        panel_layout.addWidget(scroll_area)

        # --- Right Panel (Detail View - The Summary) ---
        self.detail_area_stack = QStackedWidget()  # Use a stack to show/hide content

        # Page 0: Placeholder when nothing is selected
        placeholder = QLabel('Add a foundation type by clicking the plus button on the upper left corner.')
        placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Page 1: The actual detail view
        detail_view_widget = self.create_detail_panel()

        self.detail_area_stack.addWidget(placeholder)
        self.detail_area_stack.addWidget(detail_view_widget)

        # --- Add panels to main layout ---
        main_horizontal_layout.addWidget(panel, 2)  # 1 stretch factor
        main_horizontal_layout.addWidget(self.detail_area_stack, 6)  # 2 stretch factor (wider)
        page_layout.addLayout(main_horizontal_layout)

        # --- Bottom Navigation ---
        bottom_nav = QFrame()
        bottom_nav.setProperty('class', 'bottom-nav')
        button_layout = QHBoxLayout(bottom_nav)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(0)
        button_layout.addStretch()
        next_button = HoverButton('Next')
        next_button.setProperty('class', 'green-button next-button')
        next_button.clicked.connect(self.go_to_market_length_page)
        button_layout.addWidget(next_button)
        page_layout.addWidget(bottom_nav)

        self.stacked_widget.addWidget(page)

    def create_market_lengths_page(self) -> None:
        """Builds the UI for the third page (Rebar Market Lengths) with improved layout."""
        page = QWidget()
        page.setObjectName('marketLengthsPage')
        page.setProperty('class', 'page')
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)

        # --- This will be the main container for the title and the grid ---
        content_container = QFrame()
        content_container.setProperty('class', 'market-lengths-container')
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        # --- 1. Create the Title and Buttons Row ---
        title_row_container = QFrame()
        title_row_container.setProperty('class', 'title-row')
        title_row_layout = QHBoxLayout(title_row_container)
        title_row_layout.setContentsMargins(0, 0, 0, 0)
        title_row_layout.setSpacing(3)
        title_label = QLabel('Rebar Market Lengths')
        title_label.setProperty('class', 'header-1')
        add_button = HoverButton('+')
        add_button.setProperty('class', 'add-button green-button')
        add_button.clicked.connect(self.add_market_length)
        remove_button = HoverButton('-')
        remove_button.setProperty('class', 'remove-button red-button')
        remove_button.clicked.connect(self.remove_market_length)
        title_row_layout.addWidget(title_label)
        title_row_layout.addStretch()
        title_row_layout.addWidget(add_button)
        title_row_layout.addWidget(remove_button)

        # --- 2. Create the Grid Container ---
        grid_frame = QFrame()
        self.market_lengths_grid = QGridLayout(grid_frame)
        self.market_lengths_grid.setContentsMargins(0, 0, 0, 0)
        self.market_lengths_grid.setSpacing(0)
        # Initial drawing of the grid with a default empty state
        self.redraw_market_lengths_grid({})

        # --- 3. Add Title Row and Grid to the Content Layout ---
        content_layout.addWidget(title_row_container)
        content_layout.addWidget(grid_frame)

        # --- 4. Center the entire content block on the page ---
        centering_layout = QHBoxLayout()
        centering_layout.setContentsMargins(0, 0, 0, 0)
        centering_layout.setSpacing(0)
        centering_layout.addStretch()
        centering_layout.addWidget(content_container)
        centering_layout.addStretch()

        page_layout.addStretch()
        page_layout.addLayout(centering_layout)
        page_layout.addStretch()

        # --- 5. Navigation Buttons (at the bottom of the page) ---
        bottom_nav = QFrame()
        bottom_nav.setProperty('class', 'bottom-nav')
        button_layout = QHBoxLayout(bottom_nav)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(0)
        back_button = HoverButton('Back')
        back_button.setProperty('class', 'red-button back-button')
        back_button.clicked.connect(self.go_to_foundation_page)
        next_button = HoverButton('Generate Excel')
        next_button.setProperty('class', 'green-button next-button')
        next_button.clicked.connect(self.generate_excel)
        button_layout.addWidget(back_button)
        button_layout.addStretch(0)
        button_layout.addWidget(next_button)
        page_layout.addWidget(bottom_nav)
        self.stacked_widget.addWidget(page)

    def add_foundation_item(self) -> None:
        """Opens a dialog to add a new foundation item."""
        existing_names = [item.data['name'] for item in self.findChildren(FoundationItem)]
        dialog = FoundationDetailsDialog(parent=self, existing_names=existing_names)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            if data['name'].strip():
                new_item = FoundationItem(data)
                new_item.edit_requested.connect(self.edit_foundation_item)
                new_item.remove_requested.connect(self.remove_foundation_item)
                new_item.selected.connect(self.update_detail_view)
                self.scroll_layout.insertWidget(self.scroll_layout.count() - 1, new_item)
                self.update_detail_view(new_item)

    def create_detail_panel(self) -> QWidget:
        """Creates a comprehensive, scrollable widget to display all foundation details."""
        # The main container for the entire right panel's content
        scroll_content = QFrame()
        scroll_content.setProperty('class', 'foundation-page-detail-content')
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.detail_widgets = {'name_header': QLabel('Foundation Details')}  # Reset the dictionary

        # --- Main Title ---
        self.detail_widgets['name_header'].setProperty('class', 'header-1')  # Big, bold title
        layout.addWidget(self.detail_widgets['name_header'])

        layout.addSpacing(16)
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setProperty('class', 'separator')
        layout.addWidget(separator)
        layout.addSpacing(24)

        # --- General & Dimensions Section ---
        gen_info_label = QLabel('General Information')
        gen_info_label.setProperty('class', 'header-3')
        layout.addWidget(gen_info_label, 0, Qt.AlignmentFlag.AlignLeft)
        layout.addSpacing(14)

        form_layout_general = QFormLayout()
        form_layout_general.setSpacing(0)
        form_layout_general.setContentsMargins(0, 0, 0, 0)
        self.detail_widgets['n_footing'] = QLabel()
        self.detail_widgets['n_ped'] = QLabel()
        self.detail_widgets['cc'] = QLabel()
        self.detail_widgets['pad_dims'] = QLabel()
        self.detail_widgets['pedestal_dims'] = QLabel()

        def add_row(form_layout: Any, label: str, widget_name: str):
            form_label = QLabel(label)
            form_label.setProperty('class', 'form-label')
            form_value = self.detail_widgets[widget_name]
            form_value.setProperty('class', 'form-value')
            form_layout.addRow(form_label, form_value)

        add_row(form_layout_general, 'Total Number of Footings:', 'n_footing')
        add_row(form_layout_general, 'Pedestals per Footing:', 'n_ped')
        add_row(form_layout_general, 'Concrete Cover:', 'cc')
        add_row(form_layout_general, 'Pad Dimensions (Bx, By, t):', 'pad_dims')
        add_row(form_layout_general, 'Pedestal Dims (bx, by, h):', 'pedestal_dims')
        layout.addLayout(form_layout_general)

        layout.addSpacing(21)  # Add some vertical space between sections

        # --- Reinforcement Section ---
        reinf_detail_label = QLabel('Steel Reinforcements')
        reinf_detail_label.setProperty('class', 'header-3')
        layout.addWidget(reinf_detail_label, 0, Qt.AlignmentFlag.AlignLeft)
        layout.addSpacing(14)

        form_layout_rebar = QFormLayout()
        form_layout_rebar.setSpacing(0)
        form_layout_rebar.setContentsMargins(0, 0, 0, 0)
        self.detail_widgets['top_bar'] = QLabel('None')
        self.detail_widgets['bottom_bar'] = QLabel()
        self.detail_widgets['vertical_bar'] = QLabel()
        self.detail_widgets['perimeter_bar'] = QLabel('None')
        self.detail_widgets['stirrups_summary'] = QLabel('None')

        add_row(form_layout_rebar, 'Top Bar:', 'top_bar')
        add_row(form_layout_rebar, 'Bottom Bar:', 'bottom_bar')
        add_row(form_layout_rebar, 'Vertical Bar:', 'vertical_bar')
        add_row(form_layout_rebar, 'Perimeter Bar:', 'perimeter_bar')
        add_row(form_layout_rebar, 'Stirrups:', 'stirrups_summary')
        layout.addLayout(form_layout_rebar)

        # --- Dynamic Layout for Stirrup Types ---
        # This special layout will hold the list of individual stirrup shapes
        self.detail_stirrup_types_layout = QVBoxLayout()
        self.detail_stirrup_types_layout.setContentsMargins(0, 0, 0, 0)  # Indent the list
        self.detail_stirrup_types_layout.setSpacing(0)
        layout.addLayout(self.detail_stirrup_types_layout)

        layout.addStretch()

        # --- Set properties for all created labels ---
        for widget in self.detail_widgets.values():
            widget.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            widget.setWordWrap(True)

        # --- Make the entire panel scrollable ---
        scroll_area = make_scrollable(scroll_content)
        return scroll_area

    def get_current_checkbox_states(self) -> dict:
        """Captures the checked state of all checkboxes into a simple dictionary."""
        states = {}
        if not self.market_lengths_checkboxes:
            return {}
        for dia, lengths_dict in self.market_lengths_checkboxes.items():
            states[dia] = {}
            for length, cb_widget in lengths_dict.items():
                states[dia][length] = cb_widget.isChecked()
        return states

    def redraw_market_lengths_grid(self, previous_states: dict):
        """
        Clears and redraws the grid, applying states from the provided dictionary.

        Args:
            previous_states: A dict of {'dia': {'length': is_checked}} to restore.
        """
        if self.market_lengths_grid is None:
            return

        # Clear all existing widgets from the grid
        while self.market_lengths_grid.count():
            item = self.market_lengths_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        self.market_lengths_checkboxes = {}

        # Helper to create styled cells (this is unchanged)
        def create_cell(widget, is_header=False, is_alternate=False, x=0, y=0):
            cell = QFrame()
            cell.setAutoFillBackground(True)
            cell_layout = QHBoxLayout(cell)
            cell_layout.setContentsMargins(0, 0, 0, 0)
            cell_layout.setSpacing(0)
            if isinstance(widget, QPushButton):
                cell_layout.addWidget(widget)
            else:
                cell_layout.addStretch()
                cell_layout.addWidget(widget)
                cell_layout.addStretch()
            style_class = 'grid-cell'
            if is_header: style_class += ' header-cell'
            if is_alternate: style_class += ' alternate-row-cell'
            if x==0 and y > 0:
                style_class += ' header-column-cell'
            elif y==0 and x > 0:
                style_class += ' header-row-cell'
            elif x==0 and y==0:
                style_class += ' header-corner-cell'
            else:
                style_class += ' non-header-cell'
            cell.setProperty('class', style_class)
            return cell

        # Re-create Top-Left Header as a 'Toggle All' button
        toggle_all_btn = HoverButton('Diameter')
        toggle_all_btn.setToolTip('Toggle All Checkboxes')  # Helpful tooltip
        toggle_all_btn.setProperty('class', 'clickable-header')
        toggle_all_btn.clicked.connect(self.toggle_all_market_checkboxes)
        self.market_lengths_grid.addWidget(create_cell(toggle_all_btn, is_header=True, x=0, y=0), 0, 0)

        # Re-create Column Headers
        for col, length in enumerate(self.current_market_lengths):
            btn = HoverButton(length)
            btn.setProperty('class', 'clickable-header clickable-column-header')
            btn.clicked.connect(lambda checked, l=length: self.toggle_market_column(l))
            self.market_lengths_grid.addWidget(create_cell(btn, is_header=True, x=0, y=col + 1), 0, col + 1)

        # Re-create Rows
        for row, dia in enumerate(BAR_DIAMETERS):
            is_alternate_row = row % 2 == 1
            self.market_lengths_checkboxes[dia] = {}

            # Row Header
            btn = HoverButton(dia)
            btn.setProperty('class', 'clickable-header clickable-row-header')
            btn.clicked.connect(lambda checked, d=dia: self.toggle_market_row(d))
            self.market_lengths_grid.addWidget(create_cell(btn, is_header=True, is_alternate=is_alternate_row, x=row + 1, y=0),
                                               row + 1,
                                               0)

            # Checkboxes for each length
            for col, length in enumerate(self.current_market_lengths):
                cb = QCheckBox()
                cb.setProperty('class', 'check-box')

                # Restore the state if it exists, otherwise default to True for new lengths
                is_checked = previous_states.get(dia, {}).get(length, False)
                cb.setChecked(is_checked)
                # -----------------------------

                self.market_lengths_checkboxes[dia][length] = cb
                self.market_lengths_grid.addWidget(create_cell(cb, is_alternate=is_alternate_row, x=row+1, y=col+1), row + 1, col + 1)

    def toggle_all_market_checkboxes(self):
        """Toggles the state of every checkbox in the market lengths grid."""
        # Do nothing if the grid is empty
        if not self.market_lengths_checkboxes or not BAR_DIAMETERS or not self.current_market_lengths:
            return

        # Determine the new state by checking the first checkbox
        try:
            first_dia = BAR_DIAMETERS[0]
            first_len = self.current_market_lengths[0]
            first_checkbox = self.market_lengths_checkboxes[first_dia][first_len]
            new_state = not first_checkbox.isChecked()
        except (IndexError, KeyError):
            # If the grid is somehow malformed, default to checking all
            new_state = True

        # Apply the new state to all checkboxes
        for dia_dict in self.market_lengths_checkboxes.values():
            for checkbox in dia_dict.values():
                checkbox.setChecked(new_state)

    def add_market_length(self):
        """Prompts the user for a new market length and redraws the grid."""
        # --- Create an instance of the dialog ---
        dialog = QInputDialog(self)

        # --- Set an objectName for QSS styling ---
        dialog.setObjectName('marketLengthInputDialog')

        # --- Configure the dialog's properties ---
        dialog.setWindowTitle('Add Market Length')
        dialog.setLabelText('Enter new length (in meters):')
        dialog.setInputMode(QInputDialog.InputMode.DoubleInput)
        dialog.setDoubleRange(1.0, 50.0)
        dialog.setDoubleDecimals(1)
        dialog.setDoubleValue(1.0)

        # --- Execute the dialog and check the result ---
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_length = dialog.doubleValue()
            # The rest of your logic remains the same
            if new_length > 0:
                new_length_str = f'{new_length:.0f}m' if int(new_length) == new_length else f'{new_length:.1f}m'

                if new_length_str not in self.current_market_lengths:
                    saved_states = self.get_current_checkbox_states()
                    self.current_market_lengths.append(new_length_str)
                    self.current_market_lengths.sort(key=lambda s: float(s.replace('m', '')))
                    self.redraw_market_lengths_grid(saved_states)
                else:
                    # You can apply the same principle to QMessageBox
                    msg_box = QMessageBox(self)
                    msg_box.setObjectName('warningMessageBox')  # Style this in QSS
                    msg_box.setIcon(QMessageBox.Icon.Warning)
                    msg_box.setWindowTitle('Duplicate Length')
                    msg_box.setText('That market length already exists.')
                    msg_box.exec()

    def remove_market_length(self):
        """Prompts the user to select a market length to remove and redraws the grid."""
        if not self.current_market_lengths:
            # You can style this info box as well
            msg_box = QMessageBox(self)
            msg_box.setObjectName('infoMessageBox')
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setWindowTitle('No Lengths')
            msg_box.setText('There are no market lengths to remove.')
            msg_box.exec()
            return

        # --- Instantiate the dialog ---
        dialog = QInputDialog(self)
        dialog.setObjectName('marketLengthRemoveDialog')  # For QSS styling
        dialog.setWindowTitle('Remove Market Length')
        dialog.setLabelText('Select a length to remove:')

        # --- Configure for item selection ---
        dialog.setInputMode(QInputDialog.InputMode.TextInput)  # Necessary for combo box mode
        dialog.setComboBoxItems(self.current_market_lengths)
        dialog.setComboBoxEditable(False)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            length_to_remove = dialog.textValue()
            if length_to_remove:
                # --- SAVE STATE BEFORE REDRAWING ---
                saved_states = self.get_current_checkbox_states()
                self.current_market_lengths.remove(length_to_remove)
                # --- PASS SAVED STATE TO REDRAW METHOD ---
                self.redraw_market_lengths_grid(saved_states)

    def toggle_market_row(self, dia: str) -> None:
        """Toggles all checkboxes in a given market length row."""
        row_cbs = self.market_lengths_checkboxes[dia]
        if not row_cbs: return

        if not self.current_market_lengths: return
        first_len = self.current_market_lengths[0]
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

    def get_all_foundation_data(self) -> list[dict]:
        """
        Iterates through the layout and collects the .data dictionary from
        every FoundationItem widget.

        Returns:
            A list of dictionaries, where each dictionary contains all the
            details for one foundation type.
        """
        all_data = []
        # A QLayout's items are accessed by index.
        for i in range(self.scroll_layout.count()):
            # itemAt() returns a QLayoutItem, which is a wrapper.
            layout_item = self.scroll_layout.itemAt(i)

            # We need to get the actual widget from the layout item.
            widget = layout_item.widget()

            # IMPORTANT: Check if the widget is a FoundationItem.
            # This safely ignores the stretch/spacer at the end of the layout,
            # which has no .data attribute and would otherwise cause a crash.
            if isinstance(widget, FoundationItem):
                all_data.append(widget.data)

        return all_data

    def update_detail_view(self, item: FoundationItem):
        """Updates the right panel with ALL details of the selected item."""
        if self.current_item:
            self.current_item.deselect()

        self.current_item = item
        self.current_item.select()
        data = item.data

        # --- Populate General & Dimensions ---
        self.detail_widgets['name_header'].setText(data.get('name', 'N/A'))
        self.detail_widgets['n_footing'].setText(f'{data.get('n_footing', 0):,.0f}')
        self.detail_widgets['n_ped'].setText(f'{data.get('n_ped', 0):,.0f}')
        self.detail_widgets['cc'].setText(f'{data.get('cc', 0):,.0f} mm')
        pad_dims_text = f'{data.get('Bx', 0):,.0f} x {data.get('By', 0):,.0f} x {data.get('t', 0):,.0f} mm'
        self.detail_widgets['pad_dims'].setText(pad_dims_text)
        ped_dims_text = f'{data.get('bx', 0):,.0f} x {data.get('by', 0):,.0f} x {data.get('h', 0):,.0f} mm'
        self.detail_widgets['pedestal_dims'].setText(ped_dims_text)

        # --- Helper function for styling disabled text ---
        def format_disabled(text):
            return f"<i><font color='#A7A6A6'>{text}</font></i>"

        # --- Populate Top Bar ---
        top_bar_data = data['Top Bar']
        if top_bar_data['Enabled']:
            if top_bar_data['Input Type'] == 'Quantity':
                details = f'{top_bar_data['Value Along X']} pcs (Along X), {top_bar_data['Value Along Y']} pcs (Along Y)'
            else:  # Spacing
                details = f'@{top_bar_data['Value Along X']} mm (Along X), @{top_bar_data['Value Along Y']} mm (Along Y)'
            self.detail_widgets['top_bar'].setText(f'{top_bar_data['Diameter']} | {details}')
        else:
            self.detail_widgets['top_bar'].setText(format_disabled('Not Used'))

        # --- Populate Bottom Bar ---
        bot_bar_data = data['Bottom Bar']
        if bot_bar_data['Enabled']:
            if bot_bar_data['Input Type'] == 'Quantity':
                details = f'{bot_bar_data['Value Along X']} pcs (Along X), {bot_bar_data['Value Along Y']} pcs (Along Y)'
            else:  # Spacing
                details = f'@{bot_bar_data['Value Along X']} mm (Along X), @{bot_bar_data['Value Along Y']} mm (Along Y)'
            self.detail_widgets['bottom_bar'].setText(f'{bot_bar_data['Diameter']} | {details}')
        else:
            self.detail_widgets['bottom_bar'].setText(format_disabled('Not Used'))

        # --- Populate Vertical Bar ---
        vert_bar_data = data['Vertical Bar']
        if vert_bar_data['Enabled']:
            hook_details = f'({vert_bar_data['Hook Calculation']} Hook Length'
            if vert_bar_data['Hook Calculation'] == 'Manual':
                hook_details += f': {vert_bar_data['Hook Length']} mm'
            hook_details += ')'
            details = f'{vert_bar_data['Quantity']} pcs | {vert_bar_data['Diameter']} {hook_details}'
            self.detail_widgets['vertical_bar'].setText(details)
        else:
            self.detail_widgets['vertical_bar'].setText(format_disabled('Not Used'))

        # --- Populate Perimeter Bar ---
        perim_bar_data = data['Perimeter Bar']
        if perim_bar_data['Enabled']:
            layers = perim_bar_data.get('Layers', '1')
            self.detail_widgets['perimeter_bar'].setText(f'{layers} Layer(s) | {perim_bar_data['Diameter']}')
        else:
            self.detail_widgets['perimeter_bar'].setText(format_disabled('Not Used'))

        # --- Populate Stirrups ---
        stirrup_data = data['Stirrups']
        # Clear previous stirrup type labels
        while self.detail_stirrup_types_layout.count():
            child = self.detail_stirrup_types_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        if stirrup_data['Enabled']:
            summary = f'{stirrup_data['Quantity']} total sets, starting from <b>{stirrup_data['Extent']}</b>'
            summary += f'<br>Spacing: <code>{stirrup_data['Spacing']}</code>'
            self.detail_widgets['stirrups_summary'].setText(summary)

            # Dynamically add a label for each stirrup type in the bundle
            for stirrup_type in stirrup_data.get('Types', []):
                type_text = f'• {stirrup_type['Type']}: {stirrup_type['Diameter']}'
                if stirrup_type['Type'] in ['Tall', 'Wide', 'Octagon']:
                    type_text += f' (a: {stirrup_type['a_input']} mm)'

                type_label = QLabel(type_text)
                type_label.setProperty('class', 'form-value')
                type_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
                self.detail_stirrup_types_layout.addWidget(type_label)
        else:
            self.detail_widgets['stirrups_summary'].setText(format_disabled('Not Used'))

        # Switch the stacked widget to show the details
        self.detail_area_stack.setCurrentIndex(1)

    def edit_foundation_item(self, item: FoundationItem) -> None:
        """Opens a dialog to edit an existing foundation item."""
        existing_names = [i.data['name'] for i in self.findChildren(FoundationItem)]
        dialog = FoundationDetailsDialog(existing_details=item.data, parent=self, existing_names=existing_names)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_data = dialog.get_data()
            if new_data['name'].strip():
                item.update_details(new_data)
                self.update_detail_view(item)

    def remove_foundation_item(self, item: FoundationItem) -> None:
        """
        Removes a foundation item from the list and selects the closest
        remaining item.
        """
        was_selected = (item == self.current_item)
        index = self.scroll_layout.indexOf(item)

        # Remove the widget from the layout and schedule it for deletion
        self.scroll_layout.removeWidget(item)
        item.deleteLater()

        # If the deleted item was the selected one, decide what to select next
        if was_selected:
            # The count of actual FoundationItem widgets (excluding the final stretch)
            remaining_items_count = self.scroll_layout.count() - 1

            if remaining_items_count > 0:
                # Determine the index of the new item to select.
                # Default to the item before the deleted one, but don't go below zero.
                new_index = max(0, index - 1)

                # Ensure the index is not out of bounds if the last item was removed
                new_index = min(new_index, remaining_items_count - 1)

                new_item_to_select = self.scroll_layout.itemAt(new_index).widget()
                if isinstance(new_item_to_select, FoundationItem):
                    self.update_detail_view(new_item_to_select)
            else:
                # No items left, so show the placeholder
                self.current_item = None
                self.detail_area_stack.setCurrentIndex(0)

    @staticmethod
    def _get_used_diameters(all_foundation_data: list[dict]) -> set[str]:
        """
        Parses all foundation data and returns a set of unique diameter codes (#10, #12, etc.) that are enabled and used.
        """
        used_diameters = set()
        if not all_foundation_data:
            return used_diameters

        def process_section(section_data):
            # A section is considered enabled if its 'Enabled' key is True,
            # or if it doesn't have an 'Enabled' key (like Bottom Bar, which is always enabled).
            is_enabled = section_data.get('Enabled', True)
            if is_enabled:
                diameter_str = section_data.get('Diameter', '')
                if diameter_str:
                    used_diameters.add(diameter_str)

        for data in all_foundation_data:
            process_section(data.get('Top Bar', {}))
            process_section(data.get('Bottom Bar', {}))
            process_section(data.get('Vertical Bar', {}))
            process_section(data.get('Perimeter Bar', {}))

            stirrups_data = data.get('Stirrups', {})
            if stirrups_data.get('Enabled', True):
                for stirrup_type in stirrups_data.get('Types', []):
                    dia_text = stirrup_type.get('Diameter', '')
                    if dia_text:
                        used_diameters.add(dia_text)

        return used_diameters

    def prefill_debug_data(self):
        """Creates and adds sample foundation data if DEBUG_MODE is True."""
        print('--- DEBUG MODE: Prefilling sample data ---')
        debug_data_1 = {
            'name': 'F1 (Debug)', 'n_footing': 10, 'n_ped': 1, 'cc': 75,
            'bx': 700, 'by': 700, 'h': 1200, 'Bx': 2500, 'By': 2500, 't': 400,
            'Top Bar': {
                'Enabled': True, 'Diameter': '#16', 'Input Type': 'Spacing',
                'Value Along X': 150, 'Value Along Y': 150
            },
            'Bottom Bar': {
                'Enabled': True, 'Diameter': '#20', 'Input Type': 'Quantity',
                'Value Along X': 12, 'Value Along Y': 12
            },
            'Vertical Bar': {
                'Enabled': True, 'Diameter': '#16', 'Quantity': 8,
                'Hook Calculation': 'Automatic', 'Hook Length': 0
            },
            'Perimeter Bar': {'Enabled': False, 'Diameter': '#12', 'Layers': 1},
            'Stirrups': {
                'Enabled': True, 'Extent': 'From Face of Pad',
                'Spacing': '1@50, 5@100, rest@150', 'Quantity': 0,
                'Types': [
                    {'Type': 'Outer', 'Diameter': '#10', 'a_input': 0},
                    {'Type': 'Tall', 'Diameter': '#10', 'a_input': 150}
                ]
            }
        }

        debug_data_2 = {
            'name': 'F2 (Debug)', 'n_footing': 5, 'n_ped': 2, 'cc': 75,
            'bx': 600, 'by': 800, 'h': 1500, 'Bx': 3000, 'By': 3200, 't': 500,
            'Top Bar': {'Enabled': False, 'Diameter': '#16', 'Input Type': 'Spacing', 'Value Along X': 200,
                        'Value Along Y': 200},
            'Bottom Bar': {'Enabled': True, 'Diameter': '#25', 'Input Type': 'Quantity', 'Value Along X': 15,
                           'Value Along Y': 16},
            'Vertical Bar': {'Enabled': True, 'Diameter': '#20', 'Quantity': 12, 'Hook Calculation': 'Manual',
                             'Hook Length': 300},
            'Perimeter Bar': {'Enabled': True, 'Diameter': '#12', 'Layers': 2},
            'Stirrups': {'Enabled': True, 'Extent': 'From Bottom Bar', 'Spacing': '1@50, rest@200', 'Quantity': 0,
                         'Types': [{'Type': 'Outer', 'Diameter': '#10', 'a_input': 0}]}
        }

        all_debug_data = [debug_data_1, debug_data_2]
        first_item = None

        for data in all_debug_data:
            # 1. Create a temporary dialog in memory to run the calculations
            temp_dialog = FoundationDetailsDialog(existing_details=data, parent=self)

            # 2. The dialog's __init__ automatically populates and updates the drawing,
            #    which calculates the correct stirrup quantity.

            # 3. Get the complete, CORRECTED data (including the new quantity) from the dialog.
            corrected_data = temp_dialog.get_data()

            # 4. Create the FoundationItem with the corrected data.
            new_item = FoundationItem(corrected_data)

            # 5. Clean up the temporary dialog (optional, but good practice)
            temp_dialog.deleteLater()

            # --- Original logic continues below ---
            new_item.edit_requested.connect(self.edit_foundation_item)
            new_item.remove_requested.connect(self.remove_foundation_item)
            new_item.selected.connect(self.update_detail_view)
            self.scroll_layout.insertWidget(self.scroll_layout.count() - 1, new_item)
            if not first_item:
                first_item = new_item

        # Auto-select the first item to show details on startup
        if first_item:
            self.update_detail_view(first_item)

    def reset_application(self):
        """Resets the application to its initial state."""
        # 1. Clear all foundation items from the list
        # We loop until only the stretch item (count = 1) is left.
        while self.scroll_layout.count() > 1:
            layout_item = self.scroll_layout.takeAt(0)
            if layout_item.widget():
                # Remove the widget from the layout and schedule it for deletion
                layout_item.widget().deleteLater()

        # 2. Reset the detail view to the placeholder
        self.current_item = None
        self.detail_area_stack.setCurrentIndex(0)

        # 3. Reset all checkboxes on the market lengths page
        if self.market_lengths_checkboxes:
            for dia_dict in self.market_lengths_checkboxes.values():
                for checkbox in dia_dict.values():
                    checkbox.setChecked(False)

        # 4. Switch back to the first page
        self.stacked_widget.setCurrentIndex(0)

    def keyPressEvent(self, event):
        """
        Handles key press events for the main window.
        """
        # If the Escape key is pressed, set focus to the main window.
        if event.key() == Qt.Key.Key_Escape:
            self.setFocus()
        else:
            # Otherwise, let the default event handling proceed
            super().keyPressEvent(event)

    def go_to_foundation_page(self):
        self.stacked_widget.setCurrentIndex(0)
        self.setFocus()  # Remove focus

    def go_to_market_length_page(self):
        self.stacked_widget.setCurrentIndex(1)
        self.setFocus()  # Remove focus

    def generate_excel(self):
        all_data = self.get_all_foundation_data()
        if not all_data:
            msg_box = QMessageBox(self)
            msg_box.setObjectName('infoMessageBox')
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setWindowTitle('No Data')
            msg_box.setText('Please add at least one foundation type before generating the Excel file.')
            msg_box.exec()
            return

        required_diameters = self._get_used_diameters(all_data)
        market_lengths = {}
        for dia_code, lengths in self.market_lengths_checkboxes.items():
            available_lengths = [float(l.replace('m', '')) for l, cb in lengths.items() if cb.isChecked()]
            if available_lengths:
                market_lengths[dia_code] = available_lengths

        missing_market_lengths = sorted([dia for dia in required_diameters if dia not in market_lengths])
        proceed_with_optimization = True

        if missing_market_lengths:
            missing_list_str = '\n'.join([f'•  {d}' for d in missing_market_lengths])
            msg_box = QMessageBox(self)
            # Give this a specific name for more detailed styling
            msg_box.setObjectName('warningMessageBoxWithChoices')
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setWindowTitle('Missing Market Lengths')
            msg_box.setText('The following required rebar diameters have no market lengths selected:')
            msg_box.setInformativeText(
                f'{missing_list_str}\n\n'
                'This will prevent the generation of the Purchase and Cutting Plan sheets.\n\n'
                'Do you want to proceed with generating only the Cutting Lists?'
            )
            msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg_box.setDefaultButton(QMessageBox.StandardButton.No)

            if msg_box.exec() == QMessageBox.StandardButton.No:
                return
            proceed_with_optimization = False

        all_results, splicing_ok, wb = [], True, Workbook()
        for data in all_data:
            rebars_per_fdn_type = compile_rebar(data)
            all_results.append(rebars_per_fdn_type)
            grouped_rebars_per_fdn_type = process_rebar_input(rebars_per_fdn_type)
            wb, proceed = add_sheet_cutting_list(data['name'], grouped_rebars_per_fdn_type, market_lengths, wb)
            if not proceed:
                splicing_ok = False

        if proceed_with_optimization and splicing_ok:
            cuts_by_diameter = {}
            for bar in process_rebar_input(all_results):
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
                cuts_by_diameter[key] = [(q, l / 1000) for l, q in value.items()]

            purchase_list, cutting_plan = find_optimized_cutting_plan(cuts_by_diameter, market_lengths)
            wb = add_shet_purchase_plan(wb, purchase_list)
            wb = add_sheet_cutting_plan(wb, cutting_plan)

        # --- 4. Save and Open the Excel File ---
        wb = delete_blank_worksheets(wb)
        save_path, _ = QFileDialog.getSaveFileName(
            self, 'Save Cutting List As', 'rebar_cutting_schedule.xlsx',
            'Excel Files (*.xlsx);;All Files (*)'
        )
        if not save_path:
            return

        try:
            wb.save(save_path)
        except PermissionError:
            err_box = QMessageBox(self)
            err_box.setObjectName('criticalMessageBox')  # Style for critical errors
            err_box.setIcon(QMessageBox.Icon.Critical)
            err_box.setWindowTitle('Save Error')
            err_box.setText(f'Could not save the file to {os.path.basename(save_path)}.')
            err_box.setInformativeText('Please ensure the file is not already open in another program.')
            err_box.exec()
            return

        try:
            if sys.platform == 'win32':
                os.startfile(save_path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', save_path])
            else:
                subprocess.call(['xdg-open', save_path])
        except Exception as e:
            print(f'Could not open file automatically: {e}')

            # Refactor the final prompt
        msg_box = QMessageBox(self)
        msg_box.setObjectName('questionMessageBox')  # Style for questions
        msg_box.setWindowTitle('Generation Complete')
        msg_box.setText('The cutting list has been generated and saved.')
        msg_box.setInformativeText('What would you like to do next?')
        msg_box.setIcon(QMessageBox.Icon.Question)

        # Keep the existing button setup, we will style them via QSS
        start_over_btn = msg_box.addButton('Start Over', QMessageBox.ButtonRole.ResetRole)
        msg_box.addButton('Close Program', QMessageBox.ButtonRole.RejectRole)
        msg_box.setDefaultButton(start_over_btn)

        msg_box.exec()

        if msg_box.clickedButton() == start_over_btn:
            self.reset_application()
        else:
            self.close()

if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    wheel_event_filter = GlobalWheelEventFilter()
    app.installEventFilter(wheel_event_filter)
    app.setStyleSheet(load_stylesheet('style.qss'))
    window = MultiPageApp()
    window.show()
    sys.exit(app.exec())