import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QGridLayout, QLabel, QFrame, QDialog, QScrollArea, QPushButton,
                             QSizePolicy, QSpacerItem)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QIcon, QPixmap, QDesktopServices, QAction
from PyQt6.QtCore import QUrl

# Assuming these imports exist based on your snippet
from app_concrete_mix import ConcreteMixWindow
from app_cutting_list import CuttingListWindow
from app_optimal_purchase import OptimalPurchaseWindow
from app_timeline import TimelineWindow
from constants import LOGO_MAP, VERSION
from utils import load_stylesheet, resource_path, GlobalWheelEventFilter, HoverButton, make_scrollable


class AboutDialog(QDialog):
    """A dialog to show app info, version, and contact details."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('launcherAboutPage')
        self.setWindowTitle("About")
        self.setFixedSize(400, 450)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(50, 30, 50, 30)
        layout.setSpacing(0)

        # 1. Logo
        logo_label = QLabel()
        # Using the general logo.png provided in your main window code
        pixmap = QPixmap(resource_path('images/logo.png'))
        if not pixmap.isNull():
            logo_label.setPixmap(
                pixmap.scaled(200, 200, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(logo_label)
        layout.addSpacing(25)

        # 2. Title & Version
        title = QLabel("Engineering Apps Suite")
        title.setProperty('class', 'header-1')
        layout.addWidget(title)

        version = QLabel(f"Version {VERSION}")
        version.setProperty('class', 'subtitle')
        layout.addWidget(version)
        layout.addSpacing(10)

        # 3. Description
        desc = QLabel("A collection of engineering tools for concrete mix design, rebar optimization, and cutting schedules.")
        desc.setProperty('class', 'form-value')
        desc.setWordWrap(True)
        layout.addWidget(desc)
        layout.addSpacing(10)

        # 4. Links (Email & Github)
        # We use HTML for clickable links

        # Update these details
        github_url = "https://github.com/RSv618/engineering_apps.git"
        email_address = "robertsimonuy@gmail.com"

        contact_label = QLabel()
        contact_label.setOpenExternalLinks(True)  # Crucial for opening browser
        contact_label.setText(f"""
            <p style='line-height: 120%'>
                <b>Contact:</b> <a href='mailto:{email_address}' style='color: #009580; text-decoration: none;'>{email_address}</a><br>
                <b>Source:</b> <a href='{github_url}' style='color: #009580; text-decoration: none;'>GitHub Repository</a>
            </p>
        """)
        contact_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        contact_label.setProperty('class', 'form-value')
        layout.addWidget(contact_label)

        layout.addSpacing(25)
        layout.addStretch()

        # Close Button
        close_btn = HoverButton("Close")
        close_btn.setProperty('class', 'transparent-button next-button')
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        self.setFocus()


class AppCard(QFrame):
    """A clickable card representing an application."""

    def __init__(self, title, description, icon_path, callback, parent=None):
        super().__init__(parent)
        self.callback = callback
        self.setProperty('class', 'app-card')
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setSpacing(20)
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
        self.title_label.setProperty('class', 'header-3')
        self.title_label.setWordWrap(True)
        layout.addWidget(self.title_label)

        # Description
        self.desc_label = QLabel(description)
        self.desc_label.setProperty('class', 'card-desc')
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


# Assuming make_scrollable is defined as you provided earlier
# and AppCard / related imports are available

class LauncherWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Engineering Apps Suite")
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
        self.resize(650, 600)
        self.setMinimumSize(650, 500)
        self.window = None

        # --- 1. THE MAIN WINDOW FRAME ---
        # This holds the 3 main sections: Header, Scroll Area, Footer
        main_widget = QFrame()
        main_widget.setObjectName('launcherPage')
        main_widget.setProperty('class', 'page')
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        self.setCentralWidget(main_widget)

        # --- 2. HEADER (FIXED at Top) ---
        header_widget = QFrame()
        header_widget.setProperty('class', 'title-row')
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(0)
        header_layout.addSpacing(95)

        header_title = QLabel("Select an Application")
        header_title.setProperty('class', 'header-1')
        header_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(header_title)

        self.about_btn = HoverButton("About")  # Simplified for example
        self.about_btn.setObjectName('launcherAboutButton')
        self.about_btn.setProperty('class', 'transparent-button next-button')
        self.about_btn.clicked.connect(self.show_about_dialog)
        header_layout.addWidget(self.about_btn)

        # Add Header to Main Layout immediately
        main_layout.addWidget(header_widget)

        # --- 3. THE CARD LIST (SCROLLABLE Middle) ---

        # A. Create the container that holds the cards
        self.cards_container = QFrame()
        self.cards_container.setProperty('class', 'app-card-container')

        # B. Use QHBoxLayout to stack cards
        cards_layout = QHBoxLayout(self.cards_container)
        cards_layout.setContentsMargins(0, 0, 0, 0)  # Add breathing room around cards
        cards_layout.setSpacing(0)  # Space between cards

        # C. Add your App Cards
        desc_cutting = (
            "Automate the generation of rebar cutting lists for reinforced concrete footings. "
            "Input geometry and reinforcement details to generate a fully visualized Excel schedule, "
            "optimized purchase plan, and step-by-step cutting instructions."
        )
        card1 = AppCard("Foundation Cutting List", desc_cutting, LOGO_MAP['app_cutting_list'], self.launch_cutting_list)
        cards_layout.addWidget(card1)

        desc_purchase = (
            "Minimize waste and reduce material costs using advanced linear programming. "
            "Enter your required rebar cut lengths, and the algorithm calculates the exact "
            "combination of market-length bars to purchase, complete with a waste-minimized cutting guide."
        )
        card2 = AppCard("Rebar Optimal Purchase", desc_purchase, LOGO_MAP['app_optimal_purchase'], self.launch_optimal_purchase)
        cards_layout.addWidget(card2)

        desc_mix = (
            "Calculate precise concrete mix proportions based on ACI 211.1 standards. "
            "Features field moisture adjustments, detailed aggregate property inputs, and a "
            "compressive strength maturity estimator (Dreux-Gorisse/GL2000) to predict performance."
        )
        card3 = AppCard("Concrete Mix Design", desc_mix, LOGO_MAP['app_concrete_mix'], self.launch_concrete_mix_design)
        cards_layout.addWidget(card3)

        desc_timeline = (
            "Generate professional Project Timelines and S-Curves in Excel. "
            "Track Original, Revised, and Actual schedules, assign weights to activities, "
            "and visualize progress with automatically generated charts."
        )
        card4 = AppCard("Timeline & S-Curve", desc_timeline, LOGO_MAP.get('app_timeline', 'images/logo.png'),
                        self.launch_timeline)
        cards_layout.addWidget(card4)

        # Add a stretch at the end so cards stick to the top if there are only a few
        cards_layout.addStretch()

        # D. Wrap the container in the Scroll Area
        self.scroll_area = make_scrollable(self.cards_container)

        # E. Add the SCROLL AREA to the Main Layout
        # This is the most important line. It ensures the scroll area fills the middle space.
        main_layout.addWidget(self.scroll_area)

        # --- 4. FOOTER (FIXED at Bottom) ---
        footer_widget = QFrame()
        footer_layout = QVBoxLayout(footer_widget)
        footer_layout.setContentsMargins(10, 10, 10, 10)

        footer_text = QLabel(f"v{VERSION} | Engineering Suite")
        footer_text.setProperty('class', 'subtitle')
        footer_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_layout.addWidget(footer_text)

        # Add Footer to Main Layout
        main_layout.addWidget(footer_widget)

        # Remove focus from scroll area on startup
        self.cards_container.setFocus()


    def show_about_dialog(self):
        dlg = AboutDialog(self)
        dlg.exec()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.setFocus()
        else:
            super().keyPressEvent(event)

    def _launch_app(self, app_window_class):
        """
        Helper to hide launcher, open app, and re-show launcher on close.
        """
        self.hide()  # Hide the launcher
        self.window = app_window_class()

        # We need to hook into the sub-window's close event.
        # We save the original closeEvent to ensure we don't break
        # any validation/saving logic inside the specific apps.
        original_close_event = self.window.closeEvent

        def on_close_wrapper(event):
            # Run the app's standard close logic (e.g., "Are you sure?")
            original_close_event(event)

            # If the event was accepted (window actually closed), show launcher
            if event.isAccepted():
                self.show()

        # Override the instance's closeEvent
        self.window.closeEvent = on_close_wrapper
        self.window.show()

    def launch_cutting_list(self):
        self._launch_app(CuttingListWindow)

    def launch_optimal_purchase(self):
        self._launch_app(OptimalPurchaseWindow)

    def launch_concrete_mix_design(self):
        self._launch_app(ConcreteMixWindow)

    def launch_timeline(self):
        self._launch_app(TimelineWindow)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Apply Global Filters and Styles
    wheel_event_filter = GlobalWheelEventFilter()
    app.installEventFilter(wheel_event_filter)
    app.setStyleSheet(load_stylesheet('style.qss'))
    launcher = LauncherWindow()
    launcher.show()
    sys.exit(app.exec())