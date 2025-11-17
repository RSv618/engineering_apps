import sys
import os
import subprocess
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QLabel, QComboBox, QGridLayout, QFrame,
    QCheckBox, QScrollArea, QMessageBox, QFileDialog, QInputDialog, QPushButton, QDialog
)
from PyQt6.QtGui import QCursor, QIcon
from PyQt6.QtCore import Qt, QPoint
from openpyxl import Workbook
from utils import (load_stylesheet, parse_nested_dict, global_exception_hook,
                   InfoPopup, HoverLabel, BlankSpinBox, HoverButton, resource_path,
                   style_invalid_input, GlobalWheelEventFilter)
from rebar_optimizer import find_optimized_cutting_plan
from constants import BAR_DIAMETERS, MARKET_LENGTHS, DEBUG_MODE
from excel_writer import add_shet_purchase_plan, add_sheet_cutting_plan, delete_blank_worksheets

r"""
TO BUILD:
pyinstaller --name 'ConcreteMix' --onefile --windowed --icon='images/logo.png' --add-data 'images:images' --add-data 'style.qss:.' app_concrete_mix.py
"""

class MultiPageApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Concrete Mix Design')
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
        self.setGeometry(50, 50, 600, 600)
        self.setMinimumWidth(600)
        self.setMinimumHeight(500)

        # --- Initialize class members ---
        # self.some_variable = None

        self.info_popup = InfoPopup(self)

        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        self.create_cutting_length_page()

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


        # --- Bottom Navigation ---
        bottom_nav = QFrame()
        bottom_nav.setProperty('class', 'bottom-nav')
        button_layout = QHBoxLayout(bottom_nav)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(0)
        button_layout.addStretch()
        next_button = HoverButton('Next')
        next_button.setProperty('class', 'green-button next-button')
        # next_button.clicked.connect(self.go_to_next)
        button_layout.addWidget(next_button)
        page_layout.addWidget(bottom_nav)

        self.stacked_widget.addWidget(page)

    def prefill_for_debug(self):
        """Pre-fills all input fields with sample data for faster testing."""
        print("--- DEBUG MODE: Pre-filling forms with sample data. ---")
        ...

    def reset_application(self):
        """Resets all input fields and returns to the first page."""
        ...

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

if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    wheel_event_filter = GlobalWheelEventFilter()
    app.installEventFilter(wheel_event_filter)
    app.setStyleSheet(load_stylesheet('style.qss'))
    window = MultiPageApp()
    window.show()
    sys.exit(app.exec())
