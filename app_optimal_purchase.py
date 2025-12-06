import sys
import os
import subprocess
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QLabel, QComboBox, QGridLayout, QFrame,
    QCheckBox, QScrollArea, QMessageBox, QFileDialog, QInputDialog, QPushButton, QDialog, QDialogButtonBox
)
from PyQt6.QtGui import QCursor, QIcon
from PyQt6.QtCore import Qt, QPoint
from openpyxl import Workbook
from utils import (load_stylesheet, parse_nested_dict, global_exception_hook,
                   InfoPopup, HoverLabel, BlankSpinBox, HoverButton, resource_path,
                   style_invalid_input, GlobalWheelEventFilter, BlankDoubleSpinBox)
from rebar_optimizer import find_optimized_cutting_plan
from constants import BAR_DIAMETERS, MARKET_LENGTHS, DEBUG_MODE, LOGO_MAP
from excel_writer import add_sheet_purchase_plan, add_sheet_cutting_plan, delete_blank_worksheets

class OptimalPurchaseWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('RSB Purchase and Cutting Plan')
        self.setWindowIcon(QIcon(resource_path(LOGO_MAP['app_optimal_purchase'])))
        self.setGeometry(50, 50, 600, 600)
        self.setMinimumWidth(600)
        self.setMinimumHeight(500)

        # --- Initialize class members ---
        self.market_lengths_checkboxes = {}
        self.cutting_lengths = {'Diameter': [], 'Cutting Length': [], 'Quantity': [], 'Rows': []}
        self.current_market_lengths = list(MARKET_LENGTHS)
        self.cutting_rows_layout = None
        self.remove_cutting_button = None
        self.summary_labels = {}
        self.summary_cutting_list_layout = None
        self.market_lengths_grid = None
        self.parsed_cutting_lengths = {}
        self.active_diameters = set(BAR_DIAMETERS) # Default to all

        self.info_popup = InfoPopup(self)

        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        self.create_cutting_length_page()
        self.create_market_lengths_page()

        if DEBUG_MODE:
            self.prefill_for_debug()
        self.stacked_widget.setCurrentIndex(0)
        self.setFocus()

    def create_cutting_length_page(self) -> None:
        """Builds the UI for the first page (Cutting Lengths input)."""
        page = QWidget()
        page.setProperty('class', 'page')
        page.setObjectName('cuttingLengthPage')
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)

        # Title
        title_row = QFrame()
        title_row.setProperty('class', 'title-row')
        title_row_layout = QHBoxLayout(title_row)
        title_row_layout.setContentsMargins(0, 0, 0, 0)
        title_row_layout.setSpacing(3)
        title = QLabel('Required Rebars')
        title.setProperty('class', 'header-1')
        add_button = HoverButton('+')
        add_button.setProperty('class', 'green-button add-button')
        add_button.clicked.connect(self.add_cutting_row)
        add_button.setAutoRepeat(True)
        add_button.setAutoRepeatDelay(500)  # Wait 500ms (0.5s) before starting to repeat
        add_button.setAutoRepeatInterval(50)  # Then add a row every 50ms (fast speed)
        self.remove_cutting_button = HoverButton('-')
        self.remove_cutting_button.setProperty('class', 'red-button remove-button')
        self.remove_cutting_button.clicked.connect(self.remove_cutting_row)
        self.remove_cutting_button.setAutoRepeat(True)
        self.remove_cutting_button.setAutoRepeatDelay(500)  # Wait 500ms (0.5s) before starting to repeat
        self.remove_cutting_button.setAutoRepeatInterval(50)  # Then add a row every 50ms (fast speed)
        title_row_layout.addWidget(title)
        title_row_layout.addStretch()
        title_row_layout.addWidget(add_button)
        title_row_layout.addWidget(self.remove_cutting_button)
        page_layout.addWidget(title_row)

        # --- Header Row for Inputs ---
        header_row = QFrame()
        header_row.setProperty('class', 'header-row')
        header_row_layout = QHBoxLayout(header_row)
        header_row_layout.setContentsMargins(0, 0, 0, 0)
        header_row_layout.setSpacing(0)
        dia_header = QLabel('Diameter')
        dia_header.setProperty('class', 'header-4')
        dia_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl_header = HoverLabel('Cutting Length')  # Use the new HoverLabel
        cl_header.setProperty('class', 'header-4')
        cl_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl_header.mouseEntered.connect(self.show_cutting_length_info)
        cl_header.mouseLeft.connect(self.info_popup.hide)
        qty_header = QLabel('Quantity')
        qty_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        qty_header.setProperty('class', 'header-4')
        header_row_layout.addWidget(dia_header, 1)
        header_row_layout.addWidget(cl_header, 2)
        header_row_layout.addWidget(qty_header, 1)
        page_layout.addWidget(header_row)

        # --- Scroll Area for Dynamic Rows ---
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setProperty('class', 'scroll-bar')

        # Container widget and layout for the rows inside the scroll area
        rows_container = QWidget()
        rows_container.setProperty('class', 'list-container')
        self.cutting_rows_layout = QVBoxLayout(rows_container)
        self.cutting_rows_layout.setContentsMargins(0, 0, 0, 0)
        self.cutting_rows_layout.setSpacing(0)
        self.cutting_rows_layout.addStretch()
        self.add_cutting_row()  # Add the initial row
        scroll_area.setWidget(rows_container)
        page_layout.addWidget(scroll_area) # Add the scroll area to the main layout

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
        back_button.clicked.connect(self.go_to_cutting_length_page)
        next_button = HoverButton('Generate Excel')
        next_button.setProperty('class', 'green-button next-button')
        next_button.clicked.connect(self.generate_excel)
        button_layout.addWidget(back_button)
        button_layout.addStretch(0)
        button_layout.addWidget(next_button)
        page_layout.addWidget(bottom_nav)
        self.stacked_widget.addWidget(page)

    def go_to_cutting_length_page(self) -> None:
        """Navigates to the Cutting Lengths page (index 0)."""
        self.stacked_widget.setCurrentIndex(0)
        self.setFocus()

    def go_to_market_length_page(self):
        """Navigates to the Market Lengths page (index 1) and updates grid."""
        if not DEBUG_MODE:
            errors = self.validate_cutting_length_page()
            if errors:
                self.show_error_message('Cutting Length Page Errors', '\n'.join(errors))
                return

        # 1. Parse current inputs to find used diameters
        self.parsed_cutting_lengths = parse_nested_dict(self.cutting_lengths)
        self.active_diameters = self.get_used_diameters()

        # 2. Redraw grid
        saved_states = self.get_current_checkbox_states()
        self.redraw_market_lengths_grid(saved_states)

        self.stacked_widget.setCurrentIndex(1)
        self.setFocus()

    def show_cutting_length_info(self) -> None:
        """Displays an informational popup explaining 'Cutting Length' near the cursor."""
        info_text = (
            "<b>What is Cutting Length?</b><br>"
            "<i>Straight rebar length <u>before</u> bending.</i><br>"
            "Visible segments + Hooks − Bend deductions<br><br>"
            "<b>Ex: Square stirrup (200mm sides, 150mm hooks)</b><br>"
            "4×200 + 2×150 − (3×2<i>d</i><sub>b</sub> + 2×3<i>d</i><sub>b</sub>) = 1100 − 12<i>d</i><sub>b</sub>"
        )
        self.info_popup.set_info_text(info_text)

        # Position the popup near the cursor, offset slightly
        cursor_pos = QCursor.pos()
        self.info_popup.move(cursor_pos + QPoint(15, 15))

        self.info_popup.show()

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
        Clears and redraws the grid, showing only ACTIVE diameters.
        """
        if self.market_lengths_grid is None:
            return

        # Clear all existing widgets from the grid
        while self.market_lengths_grid.count():
            item = self.market_lengths_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        self.market_lengths_checkboxes = {}

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
            if x == 0 and y > 0:
                style_class += ' header-column-cell'
            elif y == 0 and x > 0:
                style_class += ' header-row-cell'
            elif x == 0 and y == 0:
                style_class += ' header-corner-cell'
            else:
                style_class += ' non-header-cell'
            cell.setProperty('class', style_class)
            return cell

        # Re-create Top-Left Header
        toggle_all_btn = HoverButton('Diameter')
        toggle_all_btn.setToolTip('Toggle All Checkboxes')
        toggle_all_btn.setProperty('class', 'clickable-header')
        toggle_all_btn.clicked.connect(self.toggle_all_market_checkboxes)
        self.market_lengths_grid.addWidget(create_cell(toggle_all_btn, is_header=True, x=0, y=0), 0, 0)

        # Re-create Column Headers
        for col, length in enumerate(self.current_market_lengths):
            btn = HoverButton(length)
            btn.setProperty('class', 'clickable-header clickable-column-header')
            btn.clicked.connect(lambda checked, l=length: self.toggle_market_column(l))
            self.market_lengths_grid.addWidget(create_cell(btn, is_header=True, x=0, y=col + 1), 0, col + 1)

        # Re-create Rows (FILTERED)
        visual_row_index = 0
        for dia in BAR_DIAMETERS:
            # FILTER LOGIC: Skip if not in active set
            if dia not in self.active_diameters:
                continue

            visual_row_index += 1
            is_alternate_row = visual_row_index % 2 == 1
            self.market_lengths_checkboxes[dia] = {}

            # Row Header
            btn = HoverButton(dia)
            btn.setProperty('class', 'clickable-header clickable-row-header')
            btn.clicked.connect(lambda checked, d=dia: self.toggle_market_row(d))
            self.market_lengths_grid.addWidget(
                create_cell(btn, is_header=True, is_alternate=is_alternate_row, x=visual_row_index, y=0),
                visual_row_index, 0)

            # Checkboxes
            for col, length in enumerate(self.current_market_lengths):
                cb = QCheckBox()
                cb.setProperty('class', 'check-box')
                is_checked = previous_states.get(dia, {}).get(length, False)
                cb.setChecked(is_checked)
                self.market_lengths_checkboxes[dia][length] = cb
                self.market_lengths_grid.addWidget(
                    create_cell(cb, is_alternate=is_alternate_row, x=visual_row_index, y=col + 1),
                    visual_row_index, col + 1)

        # If no diameters are active, show a placeholder message in grid
        if visual_row_index == 0:
            lbl = QLabel("No diameters required based on current inputs.")
            self.market_lengths_grid.addWidget(lbl, 1, 0, 1, len(self.current_market_lengths) + 1)

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
        # 1. Create a custom dialog
        dialog = QDialog(self)
        dialog.setObjectName('marketLengthInputDialog')  # Kept for your QSS styling
        dialog.setWindowTitle('Add Market Length')

        # 2. Setup Layout
        layout = QVBoxLayout(dialog)

        # 3. Add Label
        label = QLabel('Enter new length (in meters):')
        layout.addWidget(label)

        # 4. Add your Custom Spinbox (No buttons, custom behavior)
        # Using 0.01 as min to prevent zero input
        spinbox = BlankDoubleSpinBox(0.01, 999.9, decimals=1, initial=6.0, parent=dialog)
        layout.addWidget(spinbox)

        # 5. Add Standard OK/Cancel Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        # noinspection PyUnresolvedReferences
        button_box.accepted.connect(dialog.accept)
        # noinspection PyUnresolvedReferences
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        # 6. Execute and Process
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_length = spinbox.value()

            # Use specific check because BlankSpinBox can return a special value if empty,
            # though here it is initialized to 6.0
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
                    msg_box.setIcon(QMessageBox.Icon.Warning)
                    msg_box.setWindowTitle('Duplicate Length')
                    msg_box.setText('That market length already exists.')
                    msg_box.exec()

    def remove_market_length(self):
        """Prompts the user to select a market length to remove and redraws the grid."""
        if not self.current_market_lengths:
            # You can style this info box as well
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setWindowTitle('No Lengths')
            msg_box.setText('There are no market lengths to remove.')
            msg_box.exec()
            return

        # --- Instantiate the dialog ---
        dialog = QInputDialog(self)
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
        """Toggles active checkboxes in a column."""
        if not self.active_diameters: return
        # Find the first visible diameter to determine state
        first_visible = next((d for d in BAR_DIAMETERS if d in self.active_diameters), None)
        if not first_visible: return

        new_state = not self.market_lengths_checkboxes[first_visible][length].isChecked()

        for dia in self.active_diameters:
            if dia in self.market_lengths_checkboxes:
                self.market_lengths_checkboxes[dia][length].setChecked(new_state)

    # --- Row Management Methods ---
    def add_cutting_row(self) -> None:
        """Adds a new UI row for inputting a rebar diameter, length, and quantity."""
        container = QFrame()
        container.setProperty('class', 'cutting-length-page-cutting-row')
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)  # Add some minimal vertical margin
        row.setSpacing(3)  # Add some minimal vertical margin

        dia_input = QComboBox()
        dia_input.addItems(BAR_DIAMETERS)

        cutting_length_input = BlankSpinBox(0, 999_999, suffix=' mm')

        qty_input = BlankSpinBox(0, 999_999, suffix=' pcs')

        # Set stretch factors to align columns. Must match the header's factors.
        row.addWidget(dia_input, 1)
        row.addWidget(cutting_length_input, 2)
        row.addWidget(qty_input, 1)

        # Store widgets for later data retrieval
        self.cutting_lengths['Diameter'].append(dia_input)
        self.cutting_lengths['Cutting Length'].append(cutting_length_input)
        self.cutting_lengths['Quantity'].append(qty_input)
        self.cutting_lengths['Rows'].append(container)

        index = self.cutting_rows_layout.count() - 1
        self.cutting_rows_layout.insertWidget(index, container)
        self.update_remove_button_state()

    def remove_cutting_row(self) -> None:
        """Removes the last cutting length input row from the UI."""
        if len(self.cutting_lengths['Rows']) > 1:
            # Pop widgets from the data dictionary
            self.cutting_lengths['Diameter'].pop()
            self.cutting_lengths['Cutting Length'].pop()
            self.cutting_lengths['Quantity'].pop()
            row_to_remove = self.cutting_lengths['Rows'].pop()

            # Remove the widget from the layout and schedule it for deletion
            self.cutting_rows_layout.removeWidget(row_to_remove)
            row_to_remove.deleteLater()

            self.update_remove_button_state()

    def update_remove_button_state(self) -> None:
        """Enables or disables the 'remove row' button based on the row count."""
        is_enabled = len(self.cutting_lengths['Rows']) > 1
        if self.remove_cutting_button:
            self.remove_cutting_button.setEnabled(is_enabled)

    def get_used_diameters(self) -> set[str]:
        """
        Returns a set of unique diameter codes (#10, #12, etc.) that are enabled and used.
        """
        return set(self.parsed_cutting_lengths['Diameter'])

    def validate_market_length_page(self):
        required_diameters = self.get_used_diameters()
        market_lengths = {}
        for dia_code, lengths in self.market_lengths_checkboxes.items():
            available_lengths = [float(l.replace('m', '')) for l, cb in lengths.items() if cb.isChecked()]
            if available_lengths:
                market_lengths[dia_code] = available_lengths

        missing_market_lengths = sorted([dia for dia in required_diameters if dia not in market_lengths])
        if missing_market_lengths:
            missing_list_str = '\n'.join([f'•  {d}' for d in missing_market_lengths])
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setWindowTitle('Missing Market Lengths')
            msg_box.setText('Please select available market lengths for the following required diameters:')
            msg_box.setInformativeText(missing_list_str)
            msg_box.exec()
            return False

        # Check splicing
        diameters = self.parsed_cutting_lengths['Diameter']
        cut_lengths = self.parsed_cutting_lengths['Cutting Length']
        splicing_list = []
        for dia, c_len in zip(diameters, cut_lengths):
            max_available = max(market_lengths[dia])
            if c_len / 1000 > max_available:
                splicing_list.append((dia, c_len / 1000, max_available))
        if splicing_list:
            missing_list_str = '\n'.join(
                [f'•  {dia}: Required cut of {c_len:.1f}m exceeds the maximum available length of {max_available:.1f}m.'
                 for dia, c_len, max_available in splicing_list])
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setWindowTitle('Do not splice')
            msg_box.setText('The following rebars require splicing, which should be avoided.')
            msg_box.setInformativeText(
                f'Add or select longer rebar market lengths to accommodate longer cuts:\n\n{missing_list_str}')
            msg_box.exec()
            return False
        return True

    def prefill_for_debug(self):
        """Pre-fills all input fields with sample data for faster testing."""
        print("--- DEBUG MODE: Pre-filling forms with sample data. ---")

        # --- Page 1: Cutting Lengths ---
        # Define some sample data
        sample_cuts = [
            {'dia': '#20', 'len': 1824, 'qty': 96},
            {'dia': '#20', 'len': 1729, 'qty': 32},
            {'dia': '#12', 'len': 2476, 'qty': 20},
            {'dia': '#12', 'len': 727, 'qty': 40}
        ]

        # Clear any existing rows beyond the first one
        while len(self.cutting_lengths['Rows']) > 1:
            self.remove_cutting_row()

        # Populate the rows
        for i, cut in enumerate(sample_cuts):
            # Add a new row if needed
            if i > 0:
                self.add_cutting_row()

            # Get the widgets for the current row
            dia_input = self.cutting_lengths['Diameter'][i]
            len_input = self.cutting_lengths['Cutting Length'][i]
            qty_input = self.cutting_lengths['Quantity'][i]

            # Set the values
            dia_input.setCurrentText(cut['dia'])
            len_input.setValue(cut['len'])
            qty_input.setValue(cut['qty'])

        # --- Page 2: Market Lengths ---
        # Uncheck a couple of options for testing
        if '#10' in self.market_lengths_checkboxes:
            self.market_lengths_checkboxes['#10']['6m'].setChecked(False)
        if '#25' in self.market_lengths_checkboxes:
            self.market_lengths_checkboxes['#25']['12m'].setChecked(False)

    def reset_application(self):
        """Resets all input fields and returns to the first page."""
        # --- Reset Cutting Lengths Page ---
        # 1. Remove all but the first row
        while len(self.cutting_lengths['Rows']) > 1:
            self.remove_cutting_row()

        # 2. Clear the inputs in the remaining first row
        if self.cutting_lengths['Rows']: # Check if at least one row exists
            self.cutting_lengths['Diameter'][0].setCurrentIndex(0)
            self.cutting_lengths['Cutting Length'][0].clear()
            self.cutting_lengths['Quantity'][0].clear()

        # --- Reset Market Lengths Page ---
        for dia_lengths in self.market_lengths_checkboxes.values():
            for checkbox in dia_lengths.values():
                checkbox.setChecked(True)

        # --- Go back to the first page ---
        self.stacked_widget.setCurrentIndex(0)

    def show_error_message(self, title, message):
        """Displays a standardized error message box."""
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Warning)
        msg_box.setWindowTitle(title)
        msg_box.setText('Please correct the following errors before proceeding:')
        msg_box.setInformativeText(message)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.exec()

    def validate_cutting_length_page(self) -> list[str]:
        """
        Validates all inputs on the cutting length page, styles invalid widgets,
        and returns a list of errors.
        """
        errors = []
        # Use a validity map to track the state of each widget.
        validity_map = {}

        # Iterate through all input rows.
        all_widgets = self.cutting_lengths['Cutting Length'] + self.cutting_lengths['Quantity']
        for widget in all_widgets:
            validity_map[widget] = True  # Assume all are valid initially

        for i, (len_widget, qty_widget) in enumerate(zip(
                self.cutting_lengths['Cutting Length'],
                self.cutting_lengths['Quantity']
        )):
            row_num = i + 1
            # Rule 1: Cutting Length must be greater than 0
            if len_widget.value() <= 0:
                errors.append(f"- Row {row_num}: 'Cutting Length' must be greater than 0.")
                validity_map[len_widget] = False

            # Rule 2: Quantity must be greater than 0
            if qty_widget.value() <= 0:
                errors.append(f"- Row {row_num}: 'Quantity' must be greater than 0.")
                validity_map[qty_widget] = False

        # --- Apply styles based on final validity ---
        for widget, is_valid in validity_map.items():
            style_invalid_input(widget, is_valid)

        return sorted(list(set(errors)))

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

    def generate_excel(self):
        self.parsed_cutting_lengths = parse_nested_dict(self.cutting_lengths)
        if not DEBUG_MODE and (not self.validate_market_length_page()):
            return

        market_lengths = {}
        for dia_code, lengths in self.market_lengths_checkboxes.items():
            available_lengths = [float(l.replace('m', '')) for l, cb in lengths.items() if cb.isChecked()]
            if available_lengths:
                market_lengths[dia_code] = available_lengths

        wb = Workbook()
        cuts_by_diameter = {}
        for dia, length, quantity in zip(self.parsed_cutting_lengths['Diameter'],
                                         self.parsed_cutting_lengths['Cutting Length'],
                                         self.parsed_cutting_lengths['Quantity']):
            if dia not in cuts_by_diameter:
                cuts_by_diameter[dia] = {length: quantity}
            elif length in cuts_by_diameter[dia]:
                cuts_by_diameter[dia][length] += quantity
            else:
                cuts_by_diameter[dia][length] = quantity
        for key, value in cuts_by_diameter.items():
            cuts_by_diameter[key] = [(q, l / 1000) for l, q in value.items()]

        purchase_list, cutting_plan = find_optimized_cutting_plan(cuts_by_diameter, market_lengths)
        wb = add_sheet_purchase_plan(wb, purchase_list)
        wb = add_sheet_cutting_plan(wb, cutting_plan)

        # --- 4. Save and Open the Excel File ---
        wb = delete_blank_worksheets(wb)
        save_path, _ = QFileDialog.getSaveFileName(
            self, 'Save Cutting List As', 'rebar_purchase_plan.xlsx',
            'Excel Files (*.xlsx);;All Files (*)'
        )
        if not save_path:
            return

        try:
            wb.save(save_path)
        except PermissionError:
            err_box = QMessageBox(self)
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
        msg_box.setWindowTitle('Generation Complete')
        msg_box.setText('The purchase plan has been generated and saved.')
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
        return

if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    wheel_event_filter = GlobalWheelEventFilter()
    app.installEventFilter(wheel_event_filter)
    app.setStyleSheet(load_stylesheet('style.qss'))
    window = OptimalPurchaseWindow()
    window.show()
    sys.exit(app.exec())
