import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QFrame, QCheckBox, QGridLayout, QGroupBox,
    QFormLayout
)
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt

# Import utils and constants
from utils import (load_stylesheet, global_exception_hook,
                   InfoPopup, HoverLabel, BlankSpinBox, BlankDoubleSpinBox,
                   resource_path,
                   GlobalWheelEventFilter, make_scrollable)
from constants import DEBUG_MODE

# Import the logic engine directly
from concrete_aci import ACIMixDesign

# --- UNIT CONVERSION CONSTANTS ---
PSI_TO_MPA = 0.00689476
MPA_TO_PSI = 145.038
INCH_TO_MM = 25.4
MM_TO_INCH = 0.0393701
KG_M3_TO_LB_FT3 = 0.062428
LB_FT3_TO_KG_M3 = 16.0185

# Density Conversion ACI Standard (lb/yd3 -> kg/m3)
LB_YD3_TO_KG_M3 = 0.593276

# Volume Conversions
FT3_TO_M3 = 0.0283168
YD3_TO_M3 = 0.764555
M3_TO_YD3 = 1.30795
L_TO_GAL = 0.264172


class ConcreteMixDesign(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Concrete Mix Design (ACI 211.1)')
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
        self.setGeometry(50, 50, 1100, 750)
        self.setMinimumWidth(1000)
        self.setMinimumHeight(600)

        self.info_popup = InfoPopup(self)

        # Widgets Storage
        self.input_widgets = {}
        self.result_widgets = {}  # Changed name for clarity
        self.exposure_combos = {}

        # Main Layout
        main_widget = QWidget()
        main_widget.setProperty('class', 'page')
        self.setCentralWidget(main_widget)

        # Splitter Layout (Left: Inputs, Right: Results)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # --- LEFT PANEL (Inputs) ---
        left_panel = self.create_input_panel()
        main_layout.addWidget(left_panel, stretch=5)

        # --- RIGHT PANEL (Results) ---
        right_panel = self.create_result_panel()
        main_layout.addWidget(right_panel, stretch=5)

        # Connect signals for live calculation
        self.connect_signals()

        # Initial Calc
        if DEBUG_MODE:
            self.prefill_for_debug()
        else:
            self._set_defaults()

        self.update_calculations()

    def create_input_panel(self) -> QWidget:
        """Creates the scrollable input area using QFormLayout."""
        scroll_content = QWidget()
        scroll_content.setProperty('class', 'scroll-content')

        main_v_layout = QVBoxLayout(scroll_content)
        main_v_layout.setContentsMargins(20, 20, 20, 20)
        main_v_layout.setSpacing(20)

        # === 1. Project Parameters ===
        project_group = QGroupBox("Project Parameters")
        form_layout = QFormLayout(project_group)
        form_layout.setSpacing(10)
        form_layout.setContentsMargins(10, 15, 10, 15)

        # Total Volume (Unit depends on Output Selection)
        self.lbl_vol_input = HoverLabel("Total Volume to Pour:")
        self.lbl_vol_input.setToolTip("Total volume of concrete required.")
        self.lbl_vol_input.setProperty('class', 'form-label')
        self.input_widgets['total_vol'] = BlankDoubleSpinBox(0.1, 10000.0, decimals=2, suffix=' m³')
        form_layout.addRow(self.lbl_vol_input, self.input_widgets['total_vol'])

        # f'c
        lbl_fc = HoverLabel("Specified Strength (f'c):")
        lbl_fc.setToolTip("Specified compressive strength in MPa")
        lbl_fc.setProperty('class', 'form-label')
        self.input_widgets['fc'] = BlankSpinBox(10, 60, suffix=' MPa')

        fc_layout = QHBoxLayout()
        fc_layout.addWidget(self.input_widgets['fc'])
        self.lbl_fc_eq = QLabel("(~0 psi)")
        self.lbl_fc_eq.setStyleSheet("color: #009580; font-size: 9pt;")
        fc_layout.addWidget(self.lbl_fc_eq)
        form_layout.addRow(lbl_fc, fc_layout)

        # Slump
        lbl_slump = QLabel("Target Slump:")
        lbl_slump.setProperty('class', 'form-label')
        self.input_widgets['slump'] = BlankSpinBox(10, 200, suffix=' mm')
        slump_layout = QHBoxLayout()
        slump_layout.addWidget(self.input_widgets['slump'])
        self.lbl_slump_eq = QLabel("(~0 in)")
        self.lbl_slump_eq.setStyleSheet("color: #009580; font-size: 9pt;")
        slump_layout.addWidget(self.lbl_slump_eq)
        form_layout.addRow(lbl_slump, slump_layout)

        # NMAS
        lbl_nmas = QLabel("Max Aggregate Size:")
        lbl_nmas.setProperty('class', 'form-label')
        self.input_widgets['nmas'] = QComboBox()
        self.input_widgets['nmas'].addItems([
            "9.5 mm (3/8\")", "12.5 mm (1/2\")", "19.0 mm (3/4\")",
            "25.0 mm (1\")", "37.5 mm (1-1/2\")", "50.0 mm (2\")", "75.0 mm (3\")"
        ])
        self.input_widgets['nmas'].setCurrentIndex(2)
        form_layout.addRow(lbl_nmas, self.input_widgets['nmas'])

        # Std Dev
        self.input_widgets['chk_std_dev'] = QCheckBox("Use Historical Data (Std Dev)")
        self.input_widgets['chk_std_dev'].setProperty('class', 'form-label')
        self.input_widgets['std_dev'] = BlankDoubleSpinBox(0, 10, suffix=' MPa')
        self.input_widgets['std_dev'].setEnabled(False)

        sd_layout = QHBoxLayout()
        sd_layout.addWidget(self.input_widgets['chk_std_dev'])
        sd_layout.addWidget(self.input_widgets['std_dev'])
        form_layout.addRow("Standard Deviation:", sd_layout)

        # Air Entrained
        self.input_widgets['air'] = QCheckBox("Air Entrained Concrete")
        self.input_widgets['air'].setProperty('class', 'form-label')
        form_layout.addRow("", self.input_widgets['air'])

        main_v_layout.addWidget(project_group)

        cement_group = QGroupBox("Cement Properties")
        cement_form = QFormLayout(cement_group)
        cement_form.setSpacing(10)

        self.input_widgets['cement_type'] = QComboBox()
        self.input_widgets['cement_type'].addItems([
            "Type I (Ordinary Portland)",
            "Type IP (Blended / Portland-Pozzolan)",
            "Type IS (Portland-Slag)",
            "Custom"
        ])

        # Specific Gravity Input
        self.input_widgets['cement_sg'] = BlankDoubleSpinBox(2.0, 4.0, decimals=2)
        self.input_widgets['cement_sg'].setValue(3.15)  # Default Type I
        self.input_widgets['cement_sg'].setToolTip("Specific Gravity (Relative Density). Check cement bag/data sheet.")

        cement_form.addRow("Cement Type:", self.input_widgets['cement_type'])
        cement_form.addRow("Specific Gravity:", self.input_widgets['cement_sg'])

        # Connect combo change to update SG
        self.input_widgets['cement_type'].currentIndexChanged.connect(self.update_cement_sg)

        main_v_layout.addWidget(cement_group)

        # === 2. Exposure Classes ===
        exposure_group = QGroupBox("Exposure Classes")
        exp_form = QFormLayout(exposure_group)
        exp_form.setSpacing(8)

        def add_exposure_combo(label_text, code_prefix, options):
            combo = QComboBox()
            combo.addItems(options)
            exp_form.addRow(QLabel(label_text), combo)
            self.exposure_combos[code_prefix] = combo

        add_exposure_combo("Freezing (F):", 'F',
                           ["F0 - Not exposed", "F1 - Occasional", "F2 - Frequent", "F3 - Deicing chems"])
        add_exposure_combo("Sulfate (S):", 'S',
                           ["S0 - <0.10%", "S1 - 0.10-0.20%", "S2 - 0.20-2.00%", "S3 - >2.00%"])
        add_exposure_combo("Water (W):", 'W',
                           ["W0 - Dry", "W1 - Wet, no low perm", "W2 - Watertight"])
        add_exposure_combo("Corrosion (C):", 'C',
                           ["C0 - Dry", "C1 - Moisture", "C2 - External chlorides"])

        main_v_layout.addWidget(exposure_group)

        # === 3. Coarse Aggregate ===
        ca_group = QGroupBox("Coarse Aggregate")
        ca_form = QFormLayout(ca_group)
        ca_form.setSpacing(8)

        self.input_widgets['ca_sg'] = BlankDoubleSpinBox(1.0, 4.0, decimals=2)
        ca_form.addRow("Specific Gravity (SSD):", self.input_widgets['ca_sg'])

        self.input_widgets['ca_abs'] = BlankDoubleSpinBox(0.0, 10.0, decimals=2, suffix='%')
        ca_form.addRow("Absorption (%):", self.input_widgets['ca_abs'])

        self.input_widgets['ca_druw'] = BlankDoubleSpinBox(800, 2500, decimals=0, suffix=' kg/m³')
        ca_form.addRow("Dry Rodded Unit Wt:", self.input_widgets['ca_druw'])

        self.input_widgets['ca_moist'] = BlankDoubleSpinBox(0.0, 20.0, decimals=2, suffix='%')
        ca_form.addRow("Moisture Content (%):", self.input_widgets['ca_moist'])

        self.input_widgets['ca_shape'] = QComboBox()
        self.input_widgets['ca_shape'].addItems(["Angular (Crushed)", "Rounded (River Gravel)"])
        ca_form.addRow("Shape:", self.input_widgets['ca_shape'])

        main_v_layout.addWidget(ca_group)

        # === 4. Fine Aggregate ===
        fa_group = QGroupBox("Fine Aggregate")
        fa_form = QFormLayout(fa_group)
        fa_form.setSpacing(8)

        self.input_widgets['fa_sg'] = BlankDoubleSpinBox(1.0, 4.0, decimals=2)
        fa_form.addRow("Specific Gravity (SSD):", self.input_widgets['fa_sg'])

        self.input_widgets['fa_abs'] = BlankDoubleSpinBox(0.0, 10.0, decimals=2, suffix='%')
        fa_form.addRow("Absorption (%):", self.input_widgets['fa_abs'])

        self.input_widgets['fa_fm'] = BlankDoubleSpinBox(2.0, 3.5, decimals=2)
        fa_form.addRow("Fineness Modulus:", self.input_widgets['fa_fm'])

        self.input_widgets['fa_moist'] = BlankDoubleSpinBox(0.0, 20.0, decimals=2, suffix='%')
        fa_form.addRow("Moisture Content (%):", self.input_widgets['fa_moist'])

        main_v_layout.addWidget(fa_group)
        main_v_layout.addStretch()

        return make_scrollable(scroll_content)

    def create_result_panel(self) -> QWidget:
        """Creates the right-side results panel with configuration options."""
        container = QFrame()
        container.setStyleSheet("background-color: #f8f8f7; border-left: 1px solid #e0e0e0;")

        layout = QVBoxLayout(container)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        # --- 0. Result Controls ---
        controls_layout = QHBoxLayout()

        # Unit Selection
        self.combo_units = QComboBox()
        self.combo_units.addItems(["Metric (m³, kg)", "Imperial (yd³, lb)"])
        self.combo_units.setProperty('class', 'form-value')
        controls_layout.addWidget(QLabel("Units:"))
        controls_layout.addWidget(self.combo_units)

        controls_layout.addSpacing(20)

        # Ratio Type Selection
        self.combo_ratio = QComboBox()
        self.combo_ratio.addItems(["By Volume", "By Weight"])
        self.combo_ratio.setProperty('class', 'form-value')
        controls_layout.addWidget(QLabel("Ratio:"))
        controls_layout.addWidget(self.combo_ratio)

        controls_layout.addStretch()
        layout.addLayout(controls_layout)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("color: #ccc;")
        layout.addWidget(line)

        # --- 1. BIG PROPORTION DISPLAY ---
        self.lbl_mix_title = QLabel("Mix Proportions (by Volume)")
        self.lbl_mix_title.setProperty('class', 'header-1')
        self.lbl_mix_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_mix_title)

        self.lbl_mix_sub = QLabel("Cement : Sand : Gravel")
        self.lbl_mix_sub.setStyleSheet("color: #7f8c8d; font-size: 10pt; font-weight: bold;")
        self.lbl_mix_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_mix_sub)

        self.res_ratio_lbl = QLabel("1 : - : -")
        self.res_ratio_lbl.setStyleSheet("font-size: 36pt; font-weight: bold; color: #009580;")
        self.res_ratio_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.res_ratio_lbl)

        self.res_water_ratio = QLabel("+ Water (-)")
        self.res_water_ratio.setStyleSheet("font-size: 14pt; color: #5d5d5d;")
        self.res_water_ratio.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.res_water_ratio)

        layout.addSpacing(20)

        # --- 2. Total Requirements ---
        req_group = QGroupBox("Total Material Requirements")
        req_layout = QGridLayout(req_group)
        req_layout.setVerticalSpacing(15)

        def add_req_row(row, name, key_val, key_unit):
            l_name = QLabel(name)
            l_name.setProperty('class', 'header-2')

            l_val = QLabel("0")
            l_val.setProperty('class', 'header-2')
            l_val.setAlignment(Qt.AlignmentFlag.AlignRight)

            l_unit = QLabel("-")
            l_unit.setStyleSheet("color: #7f8c8d;")

            req_layout.addWidget(l_name, row, 0)
            req_layout.addWidget(l_val, row, 1)
            req_layout.addWidget(l_unit, row, 2)
            self.result_widgets[key_val] = l_val
            self.result_widgets[key_unit] = l_unit

        add_req_row(0, "Portland Cement", 'req_cement', 'unit_cement')
        add_req_row(1, "Sand (Fine Agg)", 'req_sand', 'unit_sand')
        add_req_row(2, "Gravel (Coarse Agg)", 'req_gravel', 'unit_gravel')
        add_req_row(3, "Water", 'req_water', 'unit_water')

        layout.addWidget(req_group)

        # --- 3. Detailed Metrics ---
        details_group = QGroupBox("Design Details")
        details_layout = QGridLayout(details_group)

        def add_det_row(row, name, key):
            lbl = QLabel(name)
            val = QLabel("-")
            val.setAlignment(Qt.AlignmentFlag.AlignRight)
            val.setStyleSheet("font-weight: bold; color: #333;")
            details_layout.addWidget(lbl, row, 0)
            details_layout.addWidget(val, row, 1)
            self.result_widgets[key] = val

        add_det_row(0, "Required f'cr:", 'det_fcr')
        add_det_row(1, "w/cm Ratio:", 'det_wcm')
        add_det_row(2, "Concrete Density:", 'det_density')
        add_det_row(3, "Air Content:", 'det_air')

        layout.addWidget(details_group)
        layout.addStretch()

        # Error Message Label
        self.err_label = QLabel("")
        self.err_label.setStyleSheet("color: #ff003c; font-size: 9pt;")
        self.err_label.setWordWrap(True)
        self.err_label.hide()
        layout.addWidget(self.err_label)

        return container

    def update_cement_sg(self):
        """Updates SG based on Cement Type selection."""
        idx = self.input_widgets['cement_type'].currentIndex()
        if idx == 0:  # Type I
            self.input_widgets['cement_sg'].setValue(3.15)
        elif idx == 1:  # Type IP (Common PH brand average)
            self.input_widgets['cement_sg'].setValue(3.05)
        elif idx == 2:  # Type IS
            self.input_widgets['cement_sg'].setValue(3.00)
        # Custom (idx 3) does not change the value, leaves it to user

        self.update_calculations()

    def connect_signals(self):
        spinboxes = [
            'total_vol', 'fc', 'slump', 'std_dev', 'cement_sg',
            'ca_sg', 'ca_abs', 'ca_druw', 'ca_moist',
            'fa_sg', 'fa_abs', 'fa_fm', 'fa_moist'
        ]
        for key in spinboxes:
            self.input_widgets[key].valueChanged.connect(self.update_calculations)

        self.input_widgets['nmas'].currentIndexChanged.connect(self.update_calculations)
        self.input_widgets['ca_shape'].currentIndexChanged.connect(self.update_calculations)
        for combo in self.exposure_combos.values():
            combo.currentIndexChanged.connect(self.update_calculations)

        self.input_widgets['air'].toggled.connect(self.update_calculations)
        self.input_widgets['chk_std_dev'].toggled.connect(self.toggle_std_dev)

        # Connect Result Configs
        self.combo_units.currentIndexChanged.connect(self.update_ui_units)
        self.combo_ratio.currentIndexChanged.connect(self.update_calculations)

    def toggle_std_dev(self, checked):
        self.input_widgets['std_dev'].setEnabled(checked)
        self.update_calculations()

    def _set_defaults(self):
        """Set safe defaults."""
        self.input_widgets['total_vol'].setValue(1.0)
        self.input_widgets['fc'].setValue(21)
        self.input_widgets['slump'].setValue(100)
        self.input_widgets['ca_sg'].setValue(2.65)
        self.input_widgets['ca_abs'].setValue(1.0)
        self.input_widgets['ca_druw'].setValue(1550)
        self.input_widgets['fa_sg'].setValue(2.60)
        self.input_widgets['fa_abs'].setValue(1.5)
        self.input_widgets['fa_fm'].setValue(2.8)
        self.input_widgets['fa_moist'].setValue(4.0)

    def update_ui_units(self):
        """Updates input labels and suffixes when Unit System changes."""
        is_metric = self.combo_units.currentIndex() == 0

        if is_metric:
            self.input_widgets['total_vol'].setSuffix(" m³")
            # If converting logic desired, do it here. For now, just update suffix.
        else:
            self.input_widgets['total_vol'].setSuffix(" yd³")

        self.update_calculations()

    def update_calculations(self):
        try:
            self.err_label.hide()

            # Check Configuration
            is_metric = self.combo_units.currentIndex() == 0
            is_vol_ratio = self.combo_ratio.currentIndex() == 0

            # --- Inputs ---
            vol_input_val = self.input_widgets['total_vol'].value()
            fc_val = self.input_widgets['fc'].value()
            slump_val = self.input_widgets['slump'].value()

            # Update conversion labels
            self.lbl_fc_eq.setText(f"(~{fc_val * MPA_TO_PSI:,.0f} psi)")
            self.lbl_slump_eq.setText(f"(~{slump_val * MM_TO_INCH:.2f} in)")

            # Basic validation
            if fc_val <= 0 or slump_val <= 0:
                return

            # --- Prepare Engine ---
            engine = ACIMixDesign()
            engine.fc = fc_val * MPA_TO_PSI
            engine.slump_target = slump_val * MM_TO_INCH
            engine.cement_sg = self.input_widgets['cement_sg'].value()

            nmas_map = {0: 0.375, 1: 0.5, 2: 0.75, 3: 1.0, 4: 1.5, 5: 2.0, 6: 3.0}
            engine.nmas = nmas_map[self.input_widgets['nmas'].currentIndex()]

            if self.input_widgets['chk_std_dev'].isChecked():
                engine.standard_deviation = self.input_widgets['std_dev'].value() * MPA_TO_PSI
            else:
                engine.standard_deviation = None

            engine.is_air_entrained = self.input_widgets['air'].isChecked()
            engine.exposure_classes = [c.currentText()[:2] for c in self.exposure_combos.values()]

            # Aggregates (Inputs are Metric, convert to Imperial for calculation)
            engine.ca_sg_ssd = self.input_widgets['ca_sg'].value()
            engine.ca_absorption = self.input_widgets['ca_abs'].value()
            engine.ca_druw = self.input_widgets['ca_druw'].value() * KG_M3_TO_LB_FT3
            engine.ca_moisture = self.input_widgets['ca_moist'].value()
            engine.ca_shape = "Angular" if self.input_widgets['ca_shape'].currentIndex() == 0 else "Rounded"

            engine.fa_sg_ssd = self.input_widgets['fa_sg'].value()
            engine.fa_absorption = self.input_widgets['fa_abs'].value()
            engine.fa_fineness_modulus = self.input_widgets['fa_fm'].value()
            engine.fa_moisture = self.input_widgets['fa_moist'].value()

            # --- Run Calculation (Returns values per 1 cubic yard in Imperial) ---
            results = engine.calculate_mix()

            # --- 1. Process Ratios (Volume vs Weight) ---
            vols_ft3 = results['volumes_ft3']
            weights_lb = results['weights_lb']

            if is_vol_ratio:
                # Ratios based on Absolute Volume (ft3)
                self.lbl_mix_title.setText("Mix Proportions (by Volume)")
                base = vols_ft3['cement']
                r_sand = vols_ft3['fa'] / base
                r_grav = vols_ft3['ca'] / base
                r_water = vols_ft3['water'] / base
            else:
                # Ratios based on Weight (lbs) - usually SSD weight for design mix
                self.lbl_mix_title.setText("Mix Proportions (by Weight)")
                base = weights_lb['cement']
                r_sand = weights_lb['fa_wet'] / base  # Using wet weight for field application
                r_grav = weights_lb['ca_wet'] / base
                r_water = weights_lb['water_net'] / base

            self.res_ratio_lbl.setText(f"1 : {r_sand:.2f} : {r_grav:.2f}")
            self.res_water_ratio.setText(f"+ Water ({r_water:.2f})")

            # --- 2. Process Total Requirements ---

            # Standard Bag Weights
            BAG_WT_METRIC = 40.0  # kg
            BAG_WT_IMPERIAL = 94.0  # lb

            if is_metric:
                # Convert inputs/outputs to Metric
                # Input volume is already m3
                total_vol_m3 = vol_input_val

                # Convert density from lb/yd3 to kg/m3
                cem_kg_m3 = weights_lb['cement'] * LB_YD3_TO_KG_M3
                sand_kg_m3 = weights_lb['fa_wet'] * LB_YD3_TO_KG_M3
                grav_kg_m3 = weights_lb['ca_wet'] * LB_YD3_TO_KG_M3
                wat_kg_m3 = weights_lb['water_net'] * LB_YD3_TO_KG_M3

                total_cem_kg = cem_kg_m3 * total_vol_m3
                total_bags = total_cem_kg / BAG_WT_METRIC

                # For Aggregates, convert Volume Fraction to m3
                # 1 yd3 = 0.764555 m3.
                # Volume in m3 = (Vol in ft3) * 0.0283168
                # But results['volumes_ft3'] sums to 27 ft3 (1 yd3).

                # Proportion of 1 m3:
                prop_sand = vols_ft3['fa'] / 27.0
                prop_grav = vols_ft3['ca'] / 27.0

                req_sand = prop_sand * total_vol_m3
                req_grav = prop_grav * total_vol_m3
                req_water = wat_kg_m3 * total_vol_m3  # 1 kg = 1 L approx

                # Update Widgets
                self.result_widgets['req_cement'].setText(f"{total_bags:,.1f}")
                self.result_widgets['unit_cement'].setText(f"bags ({BAG_WT_METRIC:.0f}kg)")

                self.result_widgets['req_sand'].setText(f"{req_sand:,.2f}")
                self.result_widgets['unit_sand'].setText("m³")

                self.result_widgets['req_gravel'].setText(f"{req_grav:,.2f}")
                self.result_widgets['unit_gravel'].setText("m³")

                self.result_widgets['req_water'].setText(f"{req_water:,.0f}")
                self.result_widgets['unit_water'].setText("Liters")

                # Details
                density = weights_lb['total'] * LB_YD3_TO_KG_M3
                self.result_widgets['det_density'].setText(f"{density:.0f} kg/m³")

            else:  # Imperial
                # Input volume is yd3
                total_vol_yd3 = vol_input_val

                # Weights are already lb/yd3
                total_cem_lb = weights_lb['cement'] * total_vol_yd3
                total_bags = total_cem_lb / BAG_WT_IMPERIAL

                # Volumes
                prop_sand = vols_ft3['fa'] / 27.0
                prop_grav = vols_ft3['ca'] / 27.0

                req_sand = prop_sand * total_vol_yd3
                req_grav = prop_grav * total_vol_yd3

                # Water: 1 lb water ~= 0.1198 gallons
                # or convert Lb to Gallons: 1 gal = 8.34 lbs
                total_water_lb = weights_lb['water_net'] * total_vol_yd3
                req_water_gal = total_water_lb / 8.34

                # Update Widgets
                self.result_widgets['req_cement'].setText(f"{total_bags:,.1f}")
                self.result_widgets['unit_cement'].setText(f"bags ({BAG_WT_IMPERIAL:.0f}lb)")

                self.result_widgets['req_sand'].setText(f"{req_sand:,.2f}")
                self.result_widgets['unit_sand'].setText("yd³")

                self.result_widgets['req_gravel'].setText(f"{req_grav:,.2f}")
                self.result_widgets['unit_gravel'].setText("yd³")

                self.result_widgets['req_water'].setText(f"{req_water_gal:,.1f}")
                self.result_widgets['unit_water'].setText("Gallons")

                # Details
                self.result_widgets['det_density'].setText(f"{weights_lb['total']:.0f} lb/yd³")

            # Common Details
            self.result_widgets['det_fcr'].setText(f"{results['f_cr'] * PSI_TO_MPA:.1f} MPa")
            self.result_widgets['det_wcm'].setText(f"{results['wcm']:.2f}")
            self.result_widgets['det_air'].setText(f"{results['air_percent']:.1f}%")

        except Exception as e:
            self.err_label.setText(f"Error: {str(e)}")
            self.err_label.show()
            self.res_ratio_lbl.setText("- : - : -")

    def prefill_for_debug(self):
        self._set_defaults()


if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    wheel_event_filter = GlobalWheelEventFilter()
    app.installEventFilter(wheel_event_filter)
    app.setStyleSheet(load_stylesheet('style.qss'))
    window = ConcreteMixDesign()
    window.show()
    sys.exit(app.exec())