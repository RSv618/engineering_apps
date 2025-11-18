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
from excel_writer import add_sheet_purchase_plan, add_sheet_cutting_plan, delete_blank_worksheets
import math

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

class SimpleConcreteEstimator:
    """
    Estimate 7,14,21,28-day compressive strength (MPa)
    Inputs:
      - w_c: water/cement ratio (mass ratio, dimensionless)
      - mix_ratio: tuple/list of (cement: sand: gravel) by volume, e.g. (1,2,3)
      - sand_quality: 'poor'|'fair'|'good'|'excellent' (grading/cleanliness)
      - gravel_quality: 'poor'|'fair'|'good'|'excellent' (hardness/shape)
      - vibration: 'poor'|'fair'|'good'|'excellent' (consolidation quality)
      - curing: float 0..1 where 1 = ideal continuous wet curing
    Output: dict with keys '7day','14day','21day','28day' (values in MPa)
    """

    # ---- base model coefficients (tune to local data) ----
    K = 95.0         # base constant (MPa scale)
    n = 1.5          # w/c exponent (Abrams-like)
    m = 0.25         # cement-content exponent
    alpha_air = 0.06 # strength loss per % air (approx)

    # Reference cement content for normalization (kg/m3)
    C_REF = 300.0

    def _estimate_cement_content_from_volume_mix(self, mix_ratio):
        """
        Heuristic conversion from common volume mixes to cement kg/m3.
        These are approximate typical values; replace with site-specific values when available.
        """
        r = tuple(mix_ratio)
        # common presets for frequently used mixes (volume ratios -> kg/m3)
        presets = {
            (1,2,3): 320.0,
            (1,1.5,3): 360.0,
            (1,3,6): 240.0,
            (1,2,2): 380.0,   # richer mix
            (1,1,2): 420.0
        }
        key = tuple(int(x) if float(x).is_integer() else x for x in r)
        if key in presets:
            return presets[key]
        # fallback: use proportion-driven approximation
        # assume cement part fraction * typical solids density factor
        total_parts = sum(r)
        # rough baseline: for mixes where cement fraction is f, estimate cement kg/m3 ~ 300*(f / (1/6))
        frac = r[0] / total_parts
        # scale such that when frac = 1/6 (for 1:2:3), we get ~320
        return max(180.0, 320.0 * frac / (1/6))

    def _quality_to_factor(self, quality):
        """
        Map qualitative rating to numeric multiplier or numeric index.
        quality in {'poor','fair','good','excellent'}
        """
        q = quality.lower()
        if q == 'poor':
            return 0.88
        if q == 'fair':
            return 0.96
        if q == 'good':
            return 1.03
        if q == 'excellent':
            return 1.08
        # default neutral
        return 1.0

    def _vibration_to_air_modifier(self, vibration):
        v = vibration.lower()
        if v == 'poor':
            return 2.0   # adds ~2% air-equivalent (honeycombing/voids)
        if v == 'fair':
            return 1.0
        if v == 'good':
            return 0.2
        if v == 'excellent':
            return -0.3
        return 0.8

    def _compute_agg_factor(self, sand_q, gravel_q):
        """
        Combine sand and gravel quality into a single aggregate factor.
        Weighted average (sand often affects workability more -> weight 0.6)
        Returns a multiplier around 0.85..1.15
        """
        sand_f = self._quality_to_factor(sand_q)
        gravel_f = self._quality_to_factor(gravel_q)
        return 0.6 * sand_f + 0.4 * gravel_f

    def _estimate_air_percent(self, w_c, vibration, entrained=False):
        """
        Baseline air:
          - entrained air (intentional) if entrained=True -> assume 4.0%
          - otherwise baseline 1.2% for well-made concrete
        Apply vibration modifier and low w/c honeycomb risk.
        """
        base = 4.0 if entrained else 1.2
        base += self._vibration_to_air_modifier(vibration)
        # honeycomb risk for very low w/c (<= 0.40) if poor vibration
        if w_c <= 0.40:
            if vibration.lower() == 'poor':
                base += 1.5
            elif vibration.lower() == 'fair':
                base += 0.7
        # clamp sensible range
        return max(0.3, min(base, 8.0))

    def _maturity_factor(self, days):
        k = 6.0
        return days / (k + days)

    def estimate(self, w_c, mix_ratio,
                 sand_quality='good',
                 gravel_quality='good',
                 vibration='good',
                 curing=1.0,
                 entrained_air=False):
        # validate curing range
        curing = float(curing)
        if curing < 0.0:
            curing = 0.0
        if curing > 1.0:
            curing = 1.0  # keep into 0..1 for this function

        cement_content = self._estimate_cement_content_from_volume_mix(mix_ratio)  # kg/m3
        agg_factor = self._compute_agg_factor(sand_quality, gravel_quality)
        air_percent = self._estimate_air_percent(w_c, vibration, entrained=entrained_air)

        # base 28-day reference strength (MPa)
        f28_ref = (
            self.K * (w_c ** -self.n) *
            (cement_content / self.C_REF) ** self.m *
            agg_factor *
            math.exp(-self.alpha_air * air_percent) *
            curing
        )

        strengths = {}
        m_28 = self._maturity_factor(28)
        for t in (7, 14, 21, 28):
            m_t = self._maturity_factor(t)
            strengths[f"{t}day"] = max(1.0, f28_ref * (m_t / m_28))

        # return supplemental info too
        return {
            'strengths_MPa': strengths,
            'cement_content_kg_m3': cement_content,
            'agg_factor': agg_factor,
            'air_percent': round(air_percent, 2),
            'curing_factor': curing
        }


# ---------------- Example usage ----------------
if __name__ == "__main__":
    model = SimpleConcreteEstimator()
    out = model.estimate(
        w_c=0.50,
        mix_ratio=(1,6,0),
        sand_quality='good',
        gravel_quality='good',
        vibration='fair',
        curing=0.95,
        entrained_air=False
    )
    print("Results:")
    for k,v in out['strengths_MPa'].items():
        print(f"  {k}: {v:.1f} MPa")
    print("Cement content (kg/m3):", out['cement_content_kg_m3'])
    print("Aggregate factor:", out['agg_factor'])
    print("Air %:", out['air_percent'])
    print("Curing factor:", out['curing_factor'])


# if __name__ == '__main__':
    # sys.excepthook = global_exception_hook
    # app = QApplication(sys.argv)
    # wheel_event_filter = GlobalWheelEventFilter()
    # app.installEventFilter(wheel_event_filter)
    # app.setStyleSheet(load_stylesheet('style.qss'))
    # window = MultiPageApp()
    # window.show()
    # sys.exit(app.exec())

