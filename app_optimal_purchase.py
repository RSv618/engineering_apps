import sys
import os
import subprocess
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QLabel, QComboBox,
    QGroupBox, QGridLayout, QFrame, QSizePolicy,
    QCheckBox, QScrollArea, QMessageBox, QFileDialog, QSpinBox, QDoubleSpinBox, QInputDialog, QPushButton, QDialog
)
from PyQt6.QtGui import QCursor, QIcon
from PyQt6.QtCore import Qt, QEvent, QPoint
from openpyxl import Workbook
from utils import (load_stylesheet, parse_nested_dict, global_exception_hook,
                   InfoPopup, HoverLabel, BlankSpinBox, HoverButton, resource_path,
                   style_invalid_input)
from rebar_optimizer import find_optimized_cutting_plan
from constants import BAR_DIAMETERS, MARKET_LENGTHS, DEBUG_MODE
from excel_writer import add_shet_purchase_plan, add_sheet_cutting_plan

class MultiPageApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('RSB Purchase and Cutting Plan')
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
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
        self.parsed_cutting_lengths = {}

        self.info_popup = InfoPopup(self)

        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        # Mousewheel protection
        QApplication.instance().installEventFilter(self)

        self.create_cutting_length_page()
        self.create_market_lengths_page()
        self.create_summary_page()

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
        self.remove_cutting_button = HoverButton('-')
        self.remove_cutting_button.setProperty('class', 'red-button remove-button')
        self.remove_cutting_button.clicked.connect(self.remove_cutting_row)
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
        next_button = HoverButton('Next')
        next_button.setProperty('class', 'green-button next-button')
        next_button.clicked.connect(self.go_to_summary_page)
        button_layout.addWidget(back_button)
        button_layout.addStretch(0)
        button_layout.addWidget(next_button)
        page_layout.addWidget(bottom_nav)
        self.stacked_widget.addWidget(page)

    def create_summary_page(self):
        """Creates the final page to summarize all user inputs with improved formatting."""
        page = QWidget()
        page.setProperty('class', 'page')
        main_layout = QVBoxLayout(page)

        # --- Helper to create a styled section ---
        def create_summary_section(title):
            group_box = QGroupBox(title)
            layout = QVBoxLayout(group_box)
            return group_box, layout

        # Create the main sections
        cutting_list_box, cutting_list_layout = create_summary_section('Rebars Cuts')
        market_box, market_layout = create_summary_section('Available Market Lengths')

        # --- Populate Cutting List Section ---
        # Header Row
        header_layout = QHBoxLayout()
        dia_header = QLabel('Diameter')
        dia_header.setProperty('class', 'summary-label')
        cl_header = QLabel('Cutting Length')
        cl_header.setProperty('class', 'summary-label')
        qty_header = QLabel('Quantity')
        qty_header.setProperty('class', 'summary-label')
        dia_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        qty_header.setAlignment(Qt.AlignmentFlag.AlignCenter)

        header_layout.addWidget(dia_header)
        header_layout.addWidget(cl_header)
        header_layout.addWidget(qty_header)
        cutting_list_layout.addLayout(header_layout)

        # Scroll Area for the dynamic list of cuts
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_container = QWidget()
        self.summary_cutting_list_layout = QVBoxLayout(scroll_container)
        self.summary_cutting_list_layout.addStretch()  # Pushes items to the top
        scroll_area.setWidget(scroll_container)
        cutting_list_layout.addWidget(scroll_area)
        cutting_list_layout.setContentsMargins(0, 0, 0, 0)

        # --- Populate Market Lengths Section ---
        market_label = QLabel('...')
        market_label.setProperty('class', 'summary-value')
        market_label.setWordWrap(True)
        self.summary_labels['market_lengths'] = market_label
        market_layout.addWidget(market_label)

        # --- Arrange sections ---
        label = QLabel('Summary')
        label.setProperty('class', 'header-0')
        main_layout.addWidget(label)

        # 1. Set the size policy of the top box to be vertically expanding
        size_policy = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        cutting_list_box.setSizePolicy(size_policy)

        # 2. Add widgets directly to the main layout with stretch factors
        #    This tells the layout how to distribute any extra vertical space.
        #    The cutting_list_box will get 2/3 of the extra space, and market_box will get 1/3.
        main_layout.addWidget(cutting_list_box, 2)
        main_layout.addWidget(market_box, 1)

        # --- Navigation Buttons ---
        button_layout = QHBoxLayout()
        back_button = HoverButton('Back')
        back_button.setProperty('class', 'red-button')
        back_button.clicked.connect(self.go_to_market_length_page)

        generate_button = HoverButton('Generate Excel')
        generate_button.setProperty('class', 'green-button')
        generate_button.clicked.connect(self.generate_purchase_list)  # Connects to your existing method

        button_layout.addWidget(back_button)
        button_layout.addStretch()
        button_layout.addWidget(generate_button)
        main_layout.addLayout(button_layout)

        self.stacked_widget.addWidget(page)

    def go_to_cutting_length_page(self) -> None:
        """Navigates to the Cutting Lengths page (index 0)."""
        self.stacked_widget.setCurrentIndex(0)
        self.setFocus()

    def go_to_market_length_page(self):
        """Navigates to the Market Lengths page (index 1)."""
        if not DEBUG_MODE:
            errors = self.validate_cutting_length_page()
            if errors:
                self.show_error_message('Cutting Length Page Errors', '\n'.join(errors))
                return  # Stop navigation if errors are found

        self.stacked_widget.setCurrentIndex(1)
        self.setFocus()

    def go_to_summary_page(self):
        """Navigates to the Summary page (index 2) after populating it with data."""
        self.populate_summary_page() # Call the population method here
        self.stacked_widget.setCurrentIndex(2)

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
            self.market_lengths_grid.addWidget(
                create_cell(btn, is_header=True, is_alternate=is_alternate_row, x=row + 1, y=0),
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
                self.market_lengths_grid.addWidget(create_cell(cb, is_alternate=is_alternate_row, x=row + 1, y=col + 1),
                                                   row + 1, col + 1)

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

    def populate_summary_page(self):
        """Reads all input data and populates the labels on the summary page."""
        # --- Clear previous summary cutting list widgets ---
        # This is important for when the user goes back and forth
        while self.summary_cutting_list_layout.count() > 1:  # Keep the stretch
            item = self.summary_cutting_list_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        # --- Populate Cutting List ---
        parsed_cutting_lengths = parse_nested_dict(self.cutting_lengths)
        self.parsed_cutting_lengths = parsed_cutting_lengths
        num_rows = len(parsed_cutting_lengths['Rows'])
        for i in range(num_rows):
            dia = parsed_cutting_lengths['Diameter'][i]
            c_len = parsed_cutting_lengths['Cutting Length'][i]
            qty = parsed_cutting_lengths['Quantity'][i]

            # Create a new row widget for the summary display
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0,0,0,0)

            dia_lbl = QLabel(dia)
            dia_lbl.setProperty('class', 'summary-value')
            c_len_lbl = QLabel(f'{c_len:,.1f} mm')
            c_len_lbl.setProperty('class', 'summary-value')
            qty_lbl = QLabel(f'{qty:,} pc{"s" if qty>1 else ""}')
            qty_lbl.setProperty('class', 'summary-value')
            dia_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            c_len_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            qty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

            row_layout.addWidget(dia_lbl)
            row_layout.addWidget(c_len_lbl)
            row_layout.addWidget(qty_lbl)

            # Insert the new row before the stretch
            index = self.summary_cutting_list_layout.count() - 1
            self.summary_cutting_list_layout.insertWidget(index, row_widget)

        # --- Populate Market Lengths ---
        market_text_lines = []
        for dia, lengths in self.market_lengths_checkboxes.items():
            available = [l for l, cb in lengths.items() if cb.isChecked()]
            if available:
                market_text_lines.append(f'<b>{dia}:</b> {', '.join(available)}')

        if market_text_lines:
            self.summary_labels['market_lengths'].setText('<br>'.join(market_text_lines))
        else:
            self.summary_labels['market_lengths'].setText('No market lengths selected.')

    def generate_purchase_list(self) -> None:
        """
        Gathers all input, runs the optimization, generates the Excel file, and
        handles post-generation actions.
        """
        # Market Length
        cuts_by_diameter = {}
        diameters = self.parsed_cutting_lengths['Diameter']
        quantities = self.parsed_cutting_lengths['Quantity']
        lengths = self.parsed_cutting_lengths['Cutting Length']
        for dia, quantity, length in zip(diameters, quantities, lengths):
            if dia not in cuts_by_diameter:
                cuts_by_diameter[dia] = {}

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
            'Save Purchase/Cutting Plan As',
            'rebar_purchase_cutting_plan.xlsx',
            'Excel Files (*.xlsx);;All Files (*)'
        )

        if not save_path:
            print('File save cancelled by user.')
            return

        # --- Try to save the file and handle potential errors ---
        try:
            # Pass optimization_results to the Excel function
            create_excel_cutting_plan(cuts_by_diameter, market_lengths, output_filename=save_path)
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


    def toggle_market_row(self, dia):
        """Toggles all checkboxes in a row based on the state of the first one."""
        row_cbs = self.market_lengths_checkboxes[dia]
        if not row_cbs: return

        # Determine target state based on the opposite of the first checkbox
        first_len = MARKET_LENGTHS[0]
        new_state = not row_cbs[first_len].isChecked()

        for cb in row_cbs.values():
            cb.setChecked(new_state)

    def toggle_market_column(self, length):
        """Toggles all checkboxes in a column based on the state of the first one."""
        if not BAR_DIAMETERS: return

        # Determine target state based on the opposite of the first row's checkbox in this column
        first_dia = BAR_DIAMETERS[0]
        new_state = not self.market_lengths_checkboxes[first_dia][length].isChecked()

        for dia in BAR_DIAMETERS:
            self.market_lengths_checkboxes[dia][length].setChecked(new_state)

def create_excel_cutting_plan(cuts_by_diameter: dict[str, list[tuple]],
                              market_lengths: dict[str, list],
                              output_filename: str = 'rebar_cutting_schedule.xlsx') -> None:
    """
    Generates a formatted Excel file with purchase and cutting plan sheets.

    Args:
        cuts_by_diameter: A dictionary of required cuts, keyed by diameter.
        market_lengths: A dictionary of available stock lengths, keyed by diameter.
        output_filename: The path to save the output .xlsx file.
    """
    purchase_list, cutting_plan = find_optimized_cutting_plan(cuts_by_diameter, market_lengths)

    proceed_cutting_plan = True
    for plan in cutting_plan:
        if 'Error' in plan:
            raise ValueError('Cannot proceed. Double check if a length exceed the available market length.')
    wb = Workbook()
    if proceed_cutting_plan:
        add_shet_purchase_plan(wb, purchase_list)
        add_sheet_cutting_plan(wb, cutting_plan)
    wb.remove(wb.active)
    wb.save(output_filename)
    print(f"Excel sheet '{output_filename}' has been created successfully.")

if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    app.setStyleSheet(load_stylesheet('style.qss'))
    window = MultiPageApp()
    window.show()
    sys.exit(app.exec())
