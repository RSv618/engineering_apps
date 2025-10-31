import sys
import os
import subprocess
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QLabel, QComboBox,
    QGroupBox, QGridLayout, QFrame, QSizePolicy,
    QCheckBox, QScrollArea, QMessageBox, QFileDialog, QSpinBox, QDoubleSpinBox
)
from PyQt6.QtGui import QCursor, QIcon
from PyQt6.QtCore import Qt, QEvent, QPoint
from openpyxl import Workbook
from utils import (load_stylesheet, parse_nested_dict, global_exception_hook,
                   InfoPopup, HoverLabel, BlankSpinBox, HoverButton, resource_path,
                   style_invalid_input)
from rebar_optimizer import find_optimized_cutting_plan
from constants import BAR_DIAMETERS, MARKET_LENGTHS, DEBUG_MODE
from excel_writer import create_purchase_sheet, create_cutting_plan_sheet

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

    def create_cutting_length_page(self) -> None:
        """Builds the UI for the first page (Cutting Lengths input)."""
        page = QWidget()
        page.setProperty('class', 'page')
        main_layout = QVBoxLayout(page)

        header_layout = QHBoxLayout()

        # --- Main Header ---
        header = QLabel('Required Rebars')
        header.setProperty('class', 'header-0')

        # --- Add/Remove Buttons ---
        add_button = HoverButton('+')
        add_button.setProperty('class', 'green-button add-row-button')
        add_button.clicked.connect(self.add_cutting_row)

        self.remove_cutting_button = HoverButton('-')
        self.remove_cutting_button.setProperty('class', 'red-button remove-row-button')
        self.remove_cutting_button.clicked.connect(self.remove_cutting_row)

        header_layout.addWidget(header)
        header_layout.addStretch()
        header_layout.addWidget(add_button)
        header_layout.addWidget(self.remove_cutting_button)
        header_layout.setContentsMargins(10, 0, 10, 0)
        main_layout.addLayout(header_layout)

        # --- Header Row for Inputs ---
        header_row_layout = QHBoxLayout()

        # Create labels and set their alignment
        dia_header = QLabel('Diameter')
        dia_header.setProperty('class', 'market-header-label')
        dia_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl_header = HoverLabel('Cutting Length')  # Use the new HoverLabel
        cl_header.setProperty('class', 'market-header-label')
        cl_header.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Connect the hover signals to handler methods
        cl_header.mouseEntered.connect(self.show_cutting_length_info)
        cl_header.mouseLeft.connect(self.info_popup.hide)

        qty_header = QLabel('Quantity')
        qty_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        qty_header.setProperty('class', 'market-header-label')

        # Add headers with stretch factors to define column widths
        header_row_layout.addWidget(dia_header, 1)
        header_row_layout.addWidget(cl_header, 2)
        header_row_layout.addWidget(qty_header, 1)
        main_layout.addLayout(header_row_layout)

        # --- Scroll Area for Dynamic Rows ---
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setProperty('class', 'scroll-area')

        # Container widget and layout for the rows inside the scroll area
        rows_container = QWidget()
        self.cutting_rows_layout = QVBoxLayout(rows_container)
        scroll_area.setWidget(rows_container)
        main_layout.addWidget(scroll_area) # Add the scroll area to the main layout
        self.cutting_rows_layout.addStretch()

        # --- Navigation ---
        button_layout = QHBoxLayout()
        next_button = HoverButton('Next')
        next_button.setProperty('class', 'green-button')
        next_button.clicked.connect(self.go_to_market_lengths_page)
        button_layout.addStretch()
        button_layout.addWidget(next_button)
        main_layout.addLayout(button_layout)

        self.stacked_widget.addWidget(page)

        # --- Add the initial row ---
        self.add_cutting_row()

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
            if isinstance(widget, HoverButton):
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
        back_button.clicked.connect(self.go_to_cutting_length_page)

        next_button = HoverButton('Next')
        next_button.setAutoDefault(True)
        next_button.setProperty('class', 'green-button')
        next_button.clicked.connect(self.go_to_summary_page)

        button_layout.addWidget(back_button)
        button_layout.addStretch()
        button_layout.addWidget(next_button)

        main_layout.addLayout(button_layout)
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
        back_button.clicked.connect(self.go_to_market_lengths_page)

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

    def go_to_market_lengths_page(self):
        """Navigates to the Market Lengths page (index 1)."""
        if not DEBUG_MODE:
            errors = self.validate_cutting_length_page()
            if errors:
                self.show_error_message('Cutting Length Page Errors', '\n'.join(errors))
                return  # Stop navigation if errors are found

        self.stacked_widget.setCurrentIndex(1)

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

    # --- Row Management Methods ---
    def add_cutting_row(self) -> None:
        """Adds a new UI row for inputting a rebar diameter, length, and quantity."""
        container_widget = QWidget()
        row = QHBoxLayout(container_widget)
        row.setContentsMargins(0, 0, 0, 0)  # Add some minimal vertical margin

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
        self.cutting_lengths['Rows'].append(container_widget)

        index = self.cutting_rows_layout.count() - 1
        self.cutting_rows_layout.insertWidget(index, container_widget)
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
            {'dia': '#10', 'len': 2095, 'qty': 16},
            {'dia': '#10', 'len': 1695, 'qty': 12},
            {'dia': '#16', 'len': 5500, 'qty': 8},
            {'dia': '#25', 'len': 8950, 'qty': 20}
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
        create_purchase_sheet(wb, purchase_list)
        create_cutting_plan_sheet(wb, cutting_plan)
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
