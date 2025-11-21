import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QGridLayout, QLabel, QFrame)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QPixmap

from app_concrete_mix import ConcreteMixWindow
from app_cutting_list import CuttingListWindow
from app_optimal_purchase import OptimalPurchaseWindow
from utils import load_stylesheet, resource_path, GlobalWheelEventFilter, HoverButton


class AppCard(QFrame):
    """A clickable card representing an application."""

    def __init__(self, title, description, icon_path, callback, parent=None):
        super().__init__(parent)
        self.callback = callback
        self.setProperty('class', 'app-card')
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # Icon
        self.icon_label = QLabel()
        pixmap = QPixmap(resource_path(icon_path))
        if not pixmap.isNull():
            self.icon_label.setPixmap(
                pixmap.scaled(64, 64, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label)

        # Title
        self.title_label = QLabel(title)
        self.title_label.setProperty('class', 'card-title')
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.title_label)

        # Description
        self.desc_label = QLabel(description)
        self.desc_label.setProperty('class', 'card-desc')
        self.desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.desc_label.setWordWrap(True)
        layout.addWidget(self.desc_label)

        layout.addStretch()

        # Button (Visual cue)
        self.btn = HoverButton("Launch")
        self.btn.setProperty('class', 'green-button next-button')
        self.btn.clicked.connect(self.on_click)
        layout.addWidget(self.btn)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.on_click()

    def on_click(self):
        self.callback()


class LauncherWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.window = None
        self.setWindowTitle("Engineering Apps Suite")
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
        self.setGeometry(100, 100, 800, 500)

        # Central Widget
        central = QWidget()
        central.setProperty('class', 'page')
        self.setCentralWidget(central)

        # Main Layout
        layout = QVBoxLayout(central)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(30)

        # Header
        header = QLabel("Select an Application")
        header.setProperty('class', 'header-1')
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header)

        # Grid for Cards
        grid = QGridLayout()
        grid.setSpacing(20)
        layout.addLayout(grid)

        # --- Define Apps ---
        # App 1: Cutting List
        card1 = AppCard(
            "Foundation Cutting List",
            "Calculate rebar cutting lists for pad footings, pedestals, and columns. Generates detailed Excel schedules.",
            "images/logo.png",  # Use specific icon if available
            self.launch_cutting_list
        )
        grid.addWidget(card1, 0, 0)

        # App 2: Optimal Purchase
        card2 = AppCard(
            "Rebar Optimal Purchase",
            "Input raw rebar requirements to optimize stock lengths, minimize waste (1D Cutting Stock), and generate purchase orders.",
            "images/logo.png",
            self.launch_optimal_purchase
        )
        grid.addWidget(card2, 0, 1)

        # App 3: Optimal Purchase
        card3 = AppCard(
            "Concrete Mix Design",
            "Implements the ACI 211.11 concrete mix design standard to find mix proportions. Estimates concrete strength based on given proportions.",
            "images/logo.png",
            self.launch_concrete_mix_design
        )
        grid.addWidget(card3, 0, 2)

        layout.addStretch()

        # Footer
        footer = QLabel("v1.0.0 | Engineering Suite")
        footer.setProperty('class', 'subtitle')
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(footer)

        # Remove focus
        self.setFocus()
    # def do_nothing(self):
    #     ...

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.setFocus()
        else:
            super().keyPressEvent(event)

    def launch_cutting_list(self):
        self.window = CuttingListWindow()
        self.window.show()
        self.close()

    def launch_optimal_purchase(self):
        self.window = OptimalPurchaseWindow()
        self.window.show()
        self.close()

    def launch_concrete_mix_design(self):
        self.window = ConcreteMixWindow()
        self.window.show()
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Apply Global Filters and Styles
    wheel_event_filter = GlobalWheelEventFilter()
    app.installEventFilter(wheel_event_filter)
    app.setStyleSheet(load_stylesheet('style.qss'))

    launcher = LauncherWindow()
    launcher.show()
    sys.exit(app.exec())