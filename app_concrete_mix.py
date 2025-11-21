import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QFrame, QCheckBox, QFormLayout, QGridLayout,
    QTabWidget, QSizePolicy, QSpinBox, QDoubleSpinBox
)
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt, QTimer

from concrete_aci import ACIMixDesign
from utils import (
    load_stylesheet, global_exception_hook,
    BlankDoubleSpinBox, make_scrollable,
    resource_path, GlobalWheelEventFilter
)
from constants import DEBUG_MODE

# --- Constants for Conversion ---
PSI_TO_MPA = 0.00689476
MPA_TO_PSI = 145.038
KG_M3_TO_LB_FT3 = 0.062428
LB_FT3_TO_KG_M3 = 16.0185
MM_TO_INCH = 0.0393701
INCH_TO_MM = 25.4
M3_TO_FT3 = 35.3147
FT3_TO_M3 = 0.0283168
KG_TO_LB = 2.20462
LB_TO_KG = 0.453592
LB_YD3_TO_KG_M3 = 0.593276
YD3_TO_M3 = 0.764555
M3_TO_YD3 = 1.307951


class ConcreteMixWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Concrete Mix Design (ACI 211.1) - Metric')
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
        self.setGeometry(50, 50, 900, 600)
        self.setMinimumWidth(900)
        self.setMinimumHeight(600)

        # Main Container
        main_widget = QWidget()
        main_widget.setObjectName('concreteMixMainWidget')
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setObjectName('mainTabs')

        # Page 1: Design
        self.design_page = ConcreteDesignPage()
        self.tabs.addTab(self.design_page, 'ACI Mix Design')

        # Page 2: Estimator (Placeholder)
        self.estimator_placeholder = QLabel('Strength Estimator Feature Coming Soon')
        self.estimator_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tabs.addTab(self.estimator_placeholder, 'Strength Estimator (Beta)')

        main_layout.addWidget(self.tabs)

        # Debounce for expensive ACI calculations
        self.calc_timer = QTimer()
        self.calc_timer.setSingleShot(True)
        self.calc_timer.setInterval(200)
        self.calc_timer.timeout.connect(self.design_page.run_design_calculation)

        self.connect_inputs()
        self.design_page.run_design_calculation()

    def connect_inputs(self):
        inputs = self.design_page.get_calculation_trigger_widgets()
        for widget in inputs:
            if isinstance(widget, (QComboBox, QCheckBox)):
                widget.currentIndexChanged.connect(self.start_debounce) if isinstance(widget,
                                                                                      QComboBox) else widget.toggled.connect(
                    self.start_debounce)
            elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                widget.valueChanged.connect(self.start_debounce)

    def start_debounce(self):
        self.calc_timer.start()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.setFocus()
        else:
            super().keyPressEvent(event)


class ConcreteDesignPage(QFrame):
    def __init__(self):
        super().__init__()
        self.setObjectName('concreteMixDesignPage')
        self.setProperty('class', 'page')

        # Store the raw results from ACI logic (always in Imperial Base)
        self.base_results = None
        self.nmas_map = None
        self.cement_map = None

        # Layouts
        page_layout = QHBoxLayout(self)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)

        # --- LEFT PANEL (Inputs) ---
        left_panel = QFrame()
        left_panel.setObjectName('concreteMixDesignLeftPanel')
        left_panel.setProperty('class', 'panel')
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)

        # Scrollable Inputs
        scroll_content = QWidget()
        scroll_content.setObjectName('concreteMixDesignScrollContent')
        scroll_content.setProperty('class', 'scroll-content')
        self.form_layout = QVBoxLayout(scroll_content)
        self.form_layout.setContentsMargins(0, 0, 0, 0)
        self.form_layout.setSpacing(0)

        self.inputs = {}
        self.equiv_labels = {}  # Store labels for dual units
        self.create_general_inputs()
        self.form_layout.addSpacing(35)
        self.create_fine_agg_inputs()
        self.form_layout.addSpacing(35)
        self.create_coarse_agg_inputs()

        self.form_layout.addStretch()
        scroll_area = make_scrollable(scroll_content)
        left_layout.addWidget(scroll_area)

        # --- RIGHT PANEL (Outputs) ---
        right_panel = QFrame()
        right_panel.setObjectName('concreteMixDesignRightPanel')
        right_panel.setProperty('class', 'page')
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        # 1. Output Controls (Top Bar)
        controls_layout = QHBoxLayout()
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(0)

        # Display Mode Dropdown (Replaces Radio Buttons)
        self.combo_display_mode = QComboBox()
        self.combo_display_mode.addItems(['By Volume', 'By Weight'])
        self.combo_display_mode.setCurrentIndex(0)
        self.combo_display_mode.currentIndexChanged.connect(self.update_output_display)

        controls_layout.addStretch()
        controls_layout.addWidget(self.combo_display_mode)

        right_layout.addLayout(controls_layout)

        # 2. Ratio Display
        mix_proportions_widget = QFrame()
        mix_proportions_widget.setObjectName('concreteMixDesignProportions')
        mix_proportions_layout = QVBoxLayout(mix_proportions_widget)
        mix_proportions_layout.setContentsMargins(0, 0, 0, 0)
        mix_proportions_layout.setSpacing(0)
        self.lbl_ratio_title = QLabel('Mix Proportions by Volume')
        self.lbl_ratio_title.setProperty('class', 'subtitle')
        self.lbl_ratio_value = QLabel('- : - : -')
        self.lbl_ratio_value.setProperty('class', 'header-1')
        mix_proportions_layout.addWidget(self.lbl_ratio_title)
        mix_proportions_layout.addWidget(self.lbl_ratio_value)
        right_layout.addWidget(mix_proportions_widget)

        # 3. Scaler Section (Volume & Bags)
        scaler_layout = QGridLayout()
        scaler_layout.setContentsMargins(0, 0, 0, 0)
        scaler_layout.setSpacing(3)

        # Total Volume Input & Imperial Label
        lbl_vol = QLabel('Total Volume:')
        self.spin_total_vol = BlankDoubleSpinBox(1, 999_999.99, decimals=2, initial=1, suffix=' m³')
        self.spin_total_vol.valueChanged.connect(self.update_output_display)
        size_policy = self.spin_total_vol.sizePolicy()
        size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
        self.spin_total_vol.setSizePolicy(size_policy)

        self.lbl_vol_imperial = QLabel('(- yd³)')
        self.lbl_vol_imperial.setProperty('class', 'unit-convert')

        # Bag Size Input
        lbl_bag = QLabel('Cement Bag:')
        self.spin_bag_size = BlankDoubleSpinBox(1, 999_999.99, decimals=2, initial=40, suffix=' kg')
        self.spin_bag_size.valueChanged.connect(self.update_output_display)
        self.spin_bag_imperial = QLabel('(- lb)')
        self.spin_bag_imperial.setProperty('class', 'unit-convert')

        # Layout for Scaler (Grid)
        scaler_layout.addWidget(lbl_vol, 0, 0)
        scaler_layout.addWidget(self.spin_total_vol, 0, 1)
        scaler_layout.addWidget(self.lbl_vol_imperial, 0, 2)
        scaler_layout.addWidget(lbl_bag, 1, 0)
        scaler_layout.addWidget(self.spin_bag_size, 1, 1)
        scaler_layout.addWidget(self.spin_bag_imperial, 1, 2)

        right_layout.addLayout(scaler_layout)

        # 4. Output Grid
        grid_widget = QFrame()
        grid_widget.setObjectName('concreteMixDesignGrid')
        self.output_grid = QGridLayout(grid_widget)
        self.output_grid.setContentsMargins(0, 0, 0, 0)
        self.output_grid.setSpacing(5)

        headers = ['Material', 'Weight', 'Volume', 'Bags']
        for c, h in enumerate(headers):
            lbl = QLabel(h)
            lbl.setProperty('class', 'header-4')
            self.output_grid.addWidget(lbl, 0, c)

        self.out_labels = {}
        materials = ['Cement', 'Sand', 'Gravel', 'Water']
        for r, mat in enumerate(materials, 1):
            lbl_mat = QLabel(mat)
            lbl_mat.setProperty('class', 'form-value')
            self.output_grid.addWidget(lbl_mat, r, 0)

            # Columns: Weight, Volume, Bags
            for c, key in enumerate(['weight', 'vol', 'bags'], 1):
                lbl_val = QLabel('0.0')
                lbl_val.setProperty('class', 'form-value')
                lbl_val.setAlignment(Qt.AlignmentFlag.AlignRight)
                self.output_grid.addWidget(lbl_val, r, c)
                self.out_labels[f'{mat}_{key}'] = lbl_val

        right_layout.addWidget(grid_widget)
        right_layout.addStretch()

        page_layout.addWidget(left_panel, 3)
        page_layout.addWidget(right_panel, 2)

        # Initial State
        if DEBUG_MODE:
            self.prefill_defaults()
        self.update_equiv_labels()
        self.update_output_display()

    # --- UI Creation Helpers ---
    def create_general_inputs(self):
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(0, 0, 0, 0)
        section_layout.setSpacing(0)

        section_title = QLabel('Concrete Material')
        section_title.setProperty('class', 'header-4')
        section_layout.addWidget(section_title)

        form_layout = QFormLayout()
        form_layout.setContentsMargins(3, 0, 0, 0)
        form_layout.setSpacing(3)

        # --- 1. CEMENT SECTION ---
        self.inputs['cement_type'] = QComboBox()
        self.cement_map = {
            'Portland (Type I, II, III, V)': 3.15,
            'Blended (Type IS, IP, IT)': 2.95,
            'Custom': 3.15
        }
        self.inputs['cement_type'].addItems(self.cement_map.keys())
        self.inputs['cement_type'].currentTextChanged.connect(self.update_cement_sg)

        self.inputs['cement_sg'] = BlankDoubleSpinBox(1.0, 5.0, initial=3.15, decimals=2)
        self.inputs['cement_sg'].setEnabled(False)  # Disabled by default

        form_layout.addRow('Cement Type:', self.inputs['cement_type'])
        form_layout.addRow('Cement S.G.:', self.inputs['cement_sg'])

        # --- 2. STRENGTH SECTION ---
        # Strength (MPa)
        self.inputs['fc'] = BlankDoubleSpinBox(0.01, 999.99, initial=20.68, decimals=2)
        self.inputs['fc'].setSuffix(' MPa')
        self.equiv_labels['fc'] = QLabel('(- psi)')
        self.equiv_labels['fc'].setProperty('class', 'unit-convert')

        fc_layout = QHBoxLayout()
        fc_layout.setContentsMargins(0, 0, 0, 0)
        fc_layout.setSpacing(3)
        fc_layout.addWidget(self.inputs['fc'])
        fc_layout.addWidget(self.equiv_labels['fc'])

        form_layout.addRow('Target Strength:', fc_layout)

        # --- 3. STANDARD DEVIATION SECTION ---
        self.inputs['use_std_dev'] = QCheckBox('Standard Deviation')
        self.inputs['use_std_dev'].setProperty('class', 'check-box')
        self.inputs['use_std_dev'].toggled.connect(self.toggle_std_dev_input)

        # Std Dev Input (Hidden/Disabled initially)
        self.inputs['std_dev'] = BlankDoubleSpinBox(0.00, 50.00, initial=2.00, decimals=2)
        self.inputs['std_dev'].setSuffix(' MPa')
        self.inputs['std_dev'].setEnabled(False)  # Locked until checkbox is ticked

        self.equiv_labels['std_dev'] = QLabel('(- psi)')
        self.equiv_labels['std_dev'].setProperty('class', 'unit-convert')

        sd_layout = QHBoxLayout()
        sd_layout.setContentsMargins(0, 0, 0, 0)
        sd_layout.setSpacing(3)
        sd_layout.addWidget(self.inputs['std_dev'])
        sd_layout.addWidget(self.equiv_labels['std_dev'])

        # Add logic row for Std Dev
        form_layout.addRow(self.inputs['use_std_dev'], sd_layout)

        # --- 4. SLUMP SECTION ---
        self.inputs['slump'] = BlankDoubleSpinBox(0.1, 999, initial=127, decimals=1)
        self.inputs['slump'].setSuffix(' mm')
        self.equiv_labels['slump'] = QLabel('(- in)')
        self.equiv_labels['slump'].setProperty('class', 'unit-convert')

        slump_layout = QHBoxLayout()
        slump_layout.setContentsMargins(0, 0, 0, 0)
        slump_layout.setSpacing(3)
        slump_layout.addWidget(self.inputs['slump'])
        slump_layout.addWidget(self.equiv_labels['slump'])

        form_layout.addRow('Target Slump:', slump_layout)

        section_layout.addLayout(form_layout)
        section_layout.addSpacing(5)

        # Air Checkbox
        self.inputs['air'] = QCheckBox('Air Entrained Concrete')
        self.inputs['air'].setProperty('class', 'check-box')
        self.inputs['air'].setChecked(False)
        section_layout.addWidget(self.inputs['air'])

        self.form_layout.addLayout(section_layout)

        # Connect labels for updates
        self.inputs['fc'].valueChanged.connect(self.update_equiv_labels)
        self.inputs['slump'].valueChanged.connect(self.update_equiv_labels)
        self.inputs['std_dev'].valueChanged.connect(self.update_equiv_labels)

    def create_coarse_agg_inputs(self):
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(0, 0, 0, 0)
        section_layout.setSpacing(0)

        section_title = QLabel('Gravel')
        section_title.setProperty('class', 'header-4')
        section_layout.addWidget(section_title)

        layout = QFormLayout()
        layout.setContentsMargins(3, 0, 0, 0)
        layout.setSpacing(3)

        # NMAS
        self.inputs['nmas'] = QComboBox()
        # Map displayed text to Imperial inch values for backend
        self.nmas_map = {
            '9.5 mm': 0.375,
            '12.5 mm': 0.5,
            '19.0 mm': 0.75,
            '25.0 mm': 1.0,
            '37.5 mm': 1.5,
            '50.0 mm': 2.0
        }

        self.inputs['nmas'].addItems(self.nmas_map.keys())
        self.inputs['nmas'].setCurrentIndex(4)
        self.inputs['ca_sg'] = BlankDoubleSpinBox(0, 10, initial=2.75, decimals=2)
        self.inputs['ca_abs'] = BlankDoubleSpinBox(0, 10, initial=1.49, decimals=2, suffix='%')
        nmas_layout = QHBoxLayout()
        nmas_layout.setContentsMargins(0, 0, 0, 0)
        nmas_layout.setSpacing(3)
        nmas_layout.addWidget(self.inputs['nmas'])
        self.equiv_labels['nmas'] = QLabel('(- inch)')
        self.equiv_labels['nmas'].setProperty('class', 'unit-convert')
        nmas_layout.addWidget(self.equiv_labels['nmas'])

        # DRUW (kg/m³)
        self.inputs['ca_druw'] = BlankDoubleSpinBox(0, 3000, initial=1588, decimals=1)
        self.inputs['ca_druw'].setSuffix(' kg/m³')
        self.equiv_labels['druw'] = QLabel('(- lb/ft³)')
        self.equiv_labels['druw'].setProperty('class', 'unit-convert')

        druw_layout = QHBoxLayout()
        druw_layout.setContentsMargins(0, 0, 0, 0)
        druw_layout.setSpacing(3)
        druw_layout.addWidget(self.inputs['ca_druw'])
        druw_layout.addWidget(self.equiv_labels['druw'])

        self.inputs['ca_mc'] = BlankDoubleSpinBox(0, 20, initial=5.00, decimals=2, suffix='%')
        self.inputs['ca_shape'] = QComboBox()
        self.inputs['ca_shape'].addItems(['Angular (Crushed)', 'Rounded (River Run)'])

        layout.addRow('Max Aggregate Size:', nmas_layout)
        layout.addRow('Particle Shape:', self.inputs['ca_shape'])
        layout.addRow('Specific Gravity (SSD):', self.inputs['ca_sg'])
        layout.addRow('Absorption:', self.inputs['ca_abs'])
        layout.addRow('Moisture Content:', self.inputs['ca_mc'])
        layout.addRow('Dry Rodded Unit Wt:', druw_layout)

        section_layout.addLayout(layout)
        self.form_layout.addLayout(section_layout)
        self.inputs['ca_druw'].valueChanged.connect(self.update_equiv_labels)
        self.inputs['nmas'].currentTextChanged.connect(self.update_equiv_labels)

    def create_fine_agg_inputs(self):
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(0, 0, 0, 0)
        section_layout.setSpacing(0)

        section_title = QLabel('Sand')
        section_title.setProperty('class', 'header-4')
        section_layout.addWidget(section_title)

        layout = QFormLayout()
        layout.setContentsMargins(3, 0, 0, 0)
        layout.setSpacing(3)

        self.inputs['fa_sg'] = BlankDoubleSpinBox(0, 10, initial=2.70, decimals=2)
        self.inputs['fa_abs'] = BlankDoubleSpinBox(0, 10, initial=1.78, decimals=2, suffix='%')
        self.inputs['fa_fm'] = BlankDoubleSpinBox(0, 10, initial=2.60, decimals=2)
        self.inputs['fa_mc'] = BlankDoubleSpinBox(0, 20, initial=6.00, decimals=2, suffix='%')

        layout.addRow('Specific Gravity (SSD):', self.inputs['fa_sg'])
        layout.addRow('Absorption:', self.inputs['fa_abs'])
        layout.addRow('Fineness Modulus:', self.inputs['fa_fm'])
        layout.addRow('Moisture Content:', self.inputs['fa_mc'])

        section_layout.addLayout(layout)
        self.form_layout.addLayout(section_layout)

    def get_calculation_trigger_widgets(self):
        return list(self.inputs.values())

    def toggle_std_dev_input(self, checked):
        self.inputs['std_dev'].setEnabled(checked)
        if checked:
            self.inputs['std_dev'].setFocus()
        self.run_design_calculation()

    def update_cement_sg(self, text):
        """
        Updates the SG spinbox based on selection.
        Disables input for presets, Enables input for Custom.
        """
        if text == 'Custom':
            self.inputs['cement_sg'].setEnabled(True)
            # We don't change the value automatically here;
            # we let the user keep whatever was there or edit it.
        else:
            self.inputs['cement_sg'].setEnabled(False)
            if text in self.cement_map:
                self.inputs['cement_sg'].setValue(self.cement_map[text])

    def update_equiv_labels(self):
        """Updates the gray secondary unit labels based on Metric inputs."""

        # Strength: MPa -> psi
        val_fc = self.inputs['fc'].value()
        equiv_psi = val_fc * MPA_TO_PSI
        self.equiv_labels['fc'].setText(f'({equiv_psi:,.0f} psi)')

        # Std Dev: MPa -> psi (NEW)
        val_sd = self.inputs['std_dev'].value()
        equiv_sd_psi = val_sd * MPA_TO_PSI
        self.equiv_labels['std_dev'].setText(f'({equiv_sd_psi:,.0f} psi)')

        # Slump: mm -> inch
        val_slump = self.inputs['slump'].value()
        equiv_inch = val_slump * MM_TO_INCH
        self.equiv_labels['slump'].setText(f'({equiv_inch:.1f} in)')

        # DRUW: kg/m³ -> lb/ft³
        val_druw = self.inputs['ca_druw'].value()
        equiv_lb_ft3 = val_druw * KG_M3_TO_LB_FT3
        self.equiv_labels['druw'].setText(f'({equiv_lb_ft3:.1f} lb/ft³)')

        # NMAS: mm -> inch
        nmas_mm = self.inputs['nmas'].currentText()
        nmas_to_inch_map = {
            '9.5 mm': '3/8\'',
            '12.5 mm': '1/2\'',
            '19.0 mm': '3/4\'',
            '25.0 mm': '1\'',
            '37.5 mm': '1.5\'',
            '50.0 mm': '2\''
        }
        self.equiv_labels['nmas'].setText(f'({nmas_to_inch_map[nmas_mm]})')

        # Update the Imperial Label (m3 -> yd3)
        batch_vol_m3 = self.spin_total_vol.value()
        batch_vol_yd3 = batch_vol_m3 * M3_TO_YD3
        self.lbl_vol_imperial.setText(f'({batch_vol_yd3:,.2f} yd³)')
        self.spin_bag_imperial.setText(f'({self.spin_bag_size.value() * KG_TO_LB:,.2f} lb)')

    def prefill_defaults(self):
        self.inputs['cement_type'].setCurrentIndex(0)  # Portland
        self.inputs['cement_sg'].setValue(3.15)

        self.inputs['fc'].setValue(20.68)  # 3000 psi
        self.inputs['slump'].setValue(127.0)  # 5 in
        self.inputs['nmas'].setCurrentIndex(4)  # 1.5 in

        self.inputs['ca_sg'].setValue(2.68)
        self.inputs['ca_abs'].setValue(0.5)
        self.inputs['ca_druw'].setValue(1600)  # 100 lb/ft3
        self.inputs['ca_mc'].setValue(2.0)

        self.inputs['fa_sg'].setValue(2.64)
        self.inputs['fa_abs'].setValue(0.7)
        self.inputs['fa_fm'].setValue(2.8)
        self.inputs['fa_mc'].setValue(6.0)

        self.update_equiv_labels()
        self.run_design_calculation()

    # --- LOGIC SECTION 1: INPUT & CALCULATION ---
    def run_design_calculation(self):
        """
        Reads Metric inputs, converts them to Imperial for the ACI Backend,
        and stores the raw Imperial results.
        """
        try:
            # 1. Read Metric Inputs & Convert to Imperial for Logic
            fc_psi = self.inputs['fc'].value() * MPA_TO_PSI
            slump_inch = self.inputs['slump'].value() * MM_TO_INCH
            ca_druw_lb_ft3 = self.inputs['ca_druw'].value() * KG_M3_TO_LB_FT3

            if fc_psi <= 0: return

            # 2. Configure ACI Object
            aci = ACIMixDesign()
            aci.fc = fc_psi

            # --- STANDARD DEVIATION LOGIC ---
            if self.inputs['use_std_dev'].isChecked():
                # Pass the value in PSI
                aci.standard_deviation = self.inputs['std_dev'].value() * MPA_TO_PSI
            else:
                # Pass None to trigger ACI 'No Data' default logic
                aci.standard_deviation = None
            # --------------------------------------

            aci.slump_target = slump_inch
            aci.cement_sg = self.inputs['cement_sg'].value()  # From previous step
            aci.nmas = self.nmas_map[self.inputs['nmas'].currentText()]
            aci.is_air_entrained = self.inputs['air'].isChecked()

            # Coarse Agg
            aci.ca_sg_ssd = self.inputs['ca_sg'].value()
            aci.ca_absorption = self.inputs['ca_abs'].value()
            aci.ca_druw = ca_druw_lb_ft3
            aci.ca_moisture = self.inputs['ca_mc'].value()
            aci.ca_shape = 'Angular' if self.inputs['ca_shape'].currentIndex() == 0 else 'Rounded'

            # Fine Agg
            aci.fa_sg_ssd = self.inputs['fa_sg'].value()
            aci.fa_absorption = self.inputs['fa_abs'].value()
            aci.fa_fineness_modulus = self.inputs['fa_fm'].value()
            aci.fa_moisture = self.inputs['fa_mc'].value()

            # 3. Run Calculation (Returns Imperial Units per 1 yd3)
            self.base_results = aci.calculate_mix()

            # 4. Update Display
            self.update_output_display()

        except Exception as e:
            if DEBUG_MODE: print(f'Calc Error: {e}')
            pass

    # --- LOGIC SECTION 2: DISPLAY ---
    def update_output_display(self):
        """
        Converts Imperial ACI results to Metric for display using
        the total volume scaler and bag size scaler.
        """
        if not self.base_results: return

        batch_vol_m3 = self.spin_total_vol.value()

        bag_size_kg = self.spin_bag_size.value()
        show_by_volume = (self.combo_display_mode.currentIndex() == 0)

        if bag_size_kg <= 0: bag_size_kg = 40.0  # Prevent divide by zero

        # Unpack Base Results (Imperial per 1 yd3)
        w_lb = self.base_results['weights_lb']
        v_ft3 = self.base_results['volumes_ft3']

        # 1. Get Raw Quantities per 1 yd3
        # Weights (Wet/Field)
        c_wet_lb = w_lb['cement']
        ca_wet_lb = w_lb['ca_wet']
        fa_wet_lb = w_lb['fa_wet']
        w_net_lb = w_lb['water_net']

        # Volumes (Absolute)
        v_c_ft3 = v_ft3['cement']
        v_ca_ft3 = v_ft3['ca']
        v_fa_ft3 = v_ft3['fa']
        v_w_ft3 = v_ft3['water']

        # 2. Convert Densities to Metric (kg/m³)
        # Conversion Factor: lb/yd³ -> kg/m³ = 0.593276
        wf = LB_YD3_TO_KG_M3

        # Calculate Metric Density (kg/m³)
        d_c_kgm3 = c_wet_lb * wf
        d_ca_kgm3 = ca_wet_lb * wf
        d_fa_kgm3 = fa_wet_lb * wf
        d_w_kgm3 = w_net_lb * wf

        # 3. Calculate Total Batch Weights (kg) based on user volume (m³)
        # Weight = Density (kg/m³) * Volume (m³)
        batch_w_c = d_c_kgm3 * batch_vol_m3
        batch_w_ca = d_ca_kgm3 * batch_vol_m3
        batch_w_fa = d_fa_kgm3 * batch_vol_m3
        batch_w_w = d_w_kgm3 * batch_vol_m3

        # 4. Calculate Total Batch Volumes (m³)
        # Volume Fraction per m³ is same as Volume Fraction per yd³
        # 1 yd³ = 27 ft³. Fraction = v_ft3 / 27
        vf = 1.0 / 27.0
        batch_v_c = (v_c_ft3 * vf) * batch_vol_m3
        batch_v_ca = (v_ca_ft3 * vf) * batch_vol_m3
        batch_v_fa = (v_fa_ft3 * vf) * batch_vol_m3
        batch_v_w = (v_w_ft3 * vf) * batch_vol_m3

        # 5. Calculate Bags (Total Weight / Bag Size)
        bags_c = batch_w_c / bag_size_kg
        bags_ca = batch_w_ca / bag_size_kg
        bags_fa = batch_w_fa / bag_size_kg

        # 6. Update Ratio Display (Dimensionless)
        if show_by_volume:
            self.lbl_ratio_title.setText('Mix Proportions by Volume')
            if v_c_ft3 > 0:
                r_sand = v_fa_ft3 / v_c_ft3
                r_gravel = v_ca_ft3 / v_c_ft3
                self.lbl_ratio_value.setText(f'1 : {r_sand:.2f} : {r_gravel:.2f}')
        else:
            self.lbl_ratio_title.setText('Mix Proportions by Weight')
            # Use SSD weights for design ratio (Base Imperial Results)
            c_ssd_lb = w_lb['cement']
            # SSD Weight = Vol * SG * 62.4
            ca_ssd_lb = v_ft3['ca'] * self.inputs['ca_sg'].value() * 62.4
            fa_ssd_lb = v_ft3['fa'] * self.inputs['fa_sg'].value() * 62.4

            if c_ssd_lb > 0:
                r_sand = fa_ssd_lb / c_ssd_lb
                r_gravel = ca_ssd_lb / c_ssd_lb
                self.lbl_ratio_value.setText(f'1 : {r_sand:.2f} : {r_gravel:.2f}')

        # 7. Update Grid
        # Order: Cement, Sand, Gravel, Water
        # Data tuple: (Weight, Volume, Bags)
        rows = [
            ('Cement', batch_w_c, batch_v_c, bags_c),
            ('Sand', batch_w_fa, batch_v_fa, bags_fa),
            ('Gravel', batch_w_ca, batch_v_ca, bags_ca),
            ('Water', batch_w_w, batch_v_w, 0.0)
        ]

        for r_idx, (mat_name, w_val, v_val, b_val) in enumerate(rows):
            self.out_labels[f'{mat_name}_weight'].setText(f'{w_val:,.1f} kg')
            self.out_labels[f'{mat_name}_vol'].setText(f'{v_val:,.3f} m³')

            if mat_name == 'Water':
                self.out_labels[f'{mat_name}_bags'].setText('-')
            else:
                self.out_labels[f'{mat_name}_bags'].setText(f'{b_val:,.1f}')


if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    wheel_event_filter = GlobalWheelEventFilter()
    app.installEventFilter(wheel_event_filter)

    app.setStyleSheet(load_stylesheet('style.qss'))

    window = ConcreteMixWindow()
    window.show()
    sys.exit(app.exec())