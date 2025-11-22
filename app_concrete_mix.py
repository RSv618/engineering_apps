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
from constants import DEBUG_MODE, MPA_TO_PSI, MM_TO_INCH, KG_M3_TO_LB_FT3, M3_TO_YD3, KG_TO_LB, LB_YD3_TO_KG_M3
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np

r"""
TO BUILD:
pyinstaller --name 'ConcreteMix' --onefile --windowed --icon='images/logo.png' --add-data 'images:images' --add-data 'style.qss:.' app_concrete_mix.py
"""

class ConcreteMixWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Concrete Mix Design (ACI 211.1) - Metric')
        self.setWindowIcon(QIcon(resource_path('images/logo.png')))
        self.setGeometry(50, 50, 900, 650)
        self.setMinimumWidth(900)
        self.setMinimumHeight(450)

        # Main Container
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setObjectName('mainTabs')

        # Page 1: Design
        self.design_page = ConcreteDesignPage()
        self.tabs.addTab(self.design_page, 'ACI Mix Design')

        # Page 2: Estimator
        self.estimator_page = ConcreteEstimatorPage()
        self.tabs.addTab(self.estimator_page, 'Strength Estimator')

        main_layout.addWidget(self.tabs)

        # Debounce for expensive ACI calculations
        self.calc_timer = QTimer()
        self.calc_timer.setSingleShot(True)
        self.calc_timer.setInterval(200)
        # noinspection PyUnresolvedReferences
        self.calc_timer.timeout.connect(self.design_page.run_design_calculation)

        self.connect_inputs()
        self.design_page.run_design_calculation()

    def connect_inputs(self):
        inputs = self.design_page.get_calculation_trigger_widgets()
        for widget in inputs:
            if isinstance(widget, (QComboBox, QCheckBox)):
                # noinspection PyUnresolvedReferences
                widget.currentIndexChanged.connect(self.start_debounce) if isinstance(widget,
                                                                                      QComboBox) else widget.toggled.connect(
                    self.start_debounce)
            elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                # noinspection PyUnresolvedReferences
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
        left_panel.setProperty('class', 'panel')
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)

        # Scrollable Inputs
        scroll_content = QFrame()
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
        # noinspection PyUnresolvedReferences
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
        lbl_vol.setProperty('class', 'form-label')
        self.spin_total_vol = BlankDoubleSpinBox(1, 999_999.99, decimals=2, initial=1, suffix=' m³')
        self.spin_total_vol.valueChanged.connect(self.update_output_display)
        size_policy = self.spin_total_vol.sizePolicy()
        size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
        self.spin_total_vol.setSizePolicy(size_policy)

        self.lbl_vol_imperial = QLabel('(- yd³)')
        self.lbl_vol_imperial.setProperty('class', 'unit-convert')

        # Bag Size Input
        lbl_bag = QLabel('Cement Bag:')
        lbl_bag.setProperty('class', 'form-label')
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

        self.inputs['cement_sg'] = BlankDoubleSpinBox(1.0, 5.0, initial=3.15, increment=0.1, decimals=2)
        self.inputs['cement_sg'].setEnabled(False)  # Disabled by default

        form_layout.addRow('Cement Type:', self.inputs['cement_type'])
        form_layout.addRow('Specific Gravity (Cement):', self.inputs['cement_sg'])

        # --- 2. STRENGTH SECTION ---
        # Strength (MPa)
        self.inputs['fc'] = BlankDoubleSpinBox(0.01, 99.99, initial=20.68, increment=0.5, decimals=2)
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
        self.inputs['std_dev'] = BlankDoubleSpinBox(0.01, 50.00, initial=2.00, increment=0.5, decimals=2)
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
        self.inputs['ca_sg'] = BlankDoubleSpinBox(0, 10, initial=2.75, increment=0.1, decimals=2)
        self.inputs['ca_abs'] = BlankDoubleSpinBox(0, 10, initial=1.49, decimals=2, increment=0.1, suffix='%')
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

        self.inputs['ca_mc'] = BlankDoubleSpinBox(0, 20, initial=5.00, decimals=2, increment=0.1, suffix='%')
        self.inputs['ca_shape'] = QComboBox()
        self.inputs['ca_shape'].addItems(['Angular (Crushed)', 'Rounded (River Run)'])

        layout.addRow('Max Gravel Size:', nmas_layout)
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

        self.inputs['fa_sg'] = BlankDoubleSpinBox(0, 10, initial=2.70, increment=0.1, decimals=2)
        self.inputs['fa_abs'] = BlankDoubleSpinBox(0, 10, initial=1.78, increment=0.1, decimals=2, suffix='%')
        self.inputs['fa_fm'] = BlankDoubleSpinBox(0, 10, initial=2.60, increment=0.1, decimals=2)
        self.inputs['fa_mc'] = BlankDoubleSpinBox(0, 20, initial=6.00, increment=0.1, decimals=2, suffix='%')

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


class ConcreteEstimatorPage(QFrame):
    def __init__(self):
        super().__init__()
        self.setObjectName('concreteEstimatorPage')
        self.setProperty('class', 'page')

        # Initialize storage
        self.inputs = {}

        # --- DATA MAPPING ---
        # Map ASTM Cement Types to approximate ISO Strength Classes (MPa) for the formula
        self.cement_map = {
            "Type I (General Purpose)": 42.5,
            "Type II (Moderate Sulfate)": 42.5,
            "Type III (High Early Strength)": 52.5,
            "Type IV (Low Heat)": 32.5,
            "Type V (High Sulfate)": 42.5
        }

        self.cement_map_S_CONSTANT_GL2000 = {
            "Type I (General Purpose)": 0.335,
            "Type II (Moderate Sulfate)": 0.4,
            "Type III (High Early Strength)": 0.13,
            "Type IV (Low Heat)": 0.335,  # Assumed
            "Type V (High Sulfate)": 0.335,  # Assumed
        }

        self.agg_map = {
            "Excellent (Clean)": 0.60,
            "Average (Standard)": 0.50,
            "Poor (Dirty)": 0.40
        }

        self.gravel_map = {
            "Small (< 20mm)": -0.05,
            "Medium (20mm - 40mm)": 0.00,
            "Large (> 40mm)": 0.05
        }

        # Layouts
        page_layout = QHBoxLayout(self)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)

        # --- LEFT PANEL (Inputs) ---
        left_panel = QFrame()
        left_panel.setProperty('class', 'panel')
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)

        # Scrollable Inputs
        scroll_content = QFrame()
        scroll_content.setObjectName('concreteEstimatorScrollContent')
        scroll_content.setProperty('class', 'scroll-content')
        self.form_layout = QVBoxLayout(scroll_content)
        self.form_layout.setContentsMargins(0, 0, 0, 0)
        self.form_layout.setSpacing(0)

        self.create_inputs()

        scroll_area = make_scrollable(scroll_content)
        left_layout.addWidget(scroll_area)

        # --- RIGHT PANEL (Graph) ---
        right_panel = QFrame()
        right_panel.setObjectName('concreteEstimatorRightPanel')
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        # Matplotlib Canvas
        self.figure = Figure(figsize=(3, 3), dpi=100, facecolor='#ffffff')
        self.canvas = FigureCanvas(self.figure)
        self.ax = self.figure.add_subplot(111)

        # Store plot data for hover interactivity
        self.sc_plot = None
        self.annot = None

        # Connect Hover Event
        self.canvas.mpl_connect("motion_notify_event", self.on_hover)

        right_layout.addStretch()
        right_layout.addWidget(self.canvas)
        right_layout.addStretch()

        # Add panels
        page_layout.addWidget(left_panel, 2)  # Smaller width for inputs
        page_layout.addWidget(right_panel, 3)  # Larger width for graph

        # Initial Calculation
        self.calculate_strength()

    def create_inputs(self):
        # --- SECTION 1: MIX PROPORTIONS ---
        lbl_mix = QLabel("Cement")
        lbl_mix.setProperty('class', 'header-4')
        self.form_layout.addWidget(lbl_mix)

        form_cement = QFormLayout()
        form_cement.setContentsMargins(3, 0, 0, 0)
        form_cement.setSpacing(3)

        self.inputs['bags'] = BlankDoubleSpinBox(1, 1000, initial=10, decimals=1, suffix=" bags")
        self.inputs['bag_weight'] = BlankDoubleSpinBox(1, 100, initial=40, decimals=1, suffix=" kg")
        self.inputs['water'] = BlankDoubleSpinBox(1, 1000, initial=200, decimals=1, suffix=" kg (L)")
        self.inputs['cement_type'] = QComboBox()
        self.inputs['cement_type'].addItems(self.cement_map.keys())
        self.inputs['cement_type'].setCurrentIndex(0)  # Type I default

        form_cement.addRow("Cement Type:", self.inputs['cement_type'])
        form_cement.addRow("Bag Weight:", self.inputs['bag_weight'])
        form_cement.addRow("Cement Qty:", self.inputs['bags'])

        self.form_layout.addLayout(form_cement)
        self.form_layout.addSpacing(35)

        # --- SECTION 2: Water ---
        form_wat = QFormLayout()
        form_wat.setContentsMargins(3, 0, 0, 0)
        form_wat.setSpacing(3)
        lbl_fac = QLabel("Water")
        lbl_fac.setProperty('class', 'header-4')
        self.form_layout.addWidget(lbl_fac)
        form_wat.addRow("Total Water:", self.inputs['water'])
        self.form_layout.addLayout(form_wat)
        self.form_layout.addSpacing(35)

        # --- SECTION 3: Aggregates ---
        lbl_fac = QLabel("Aggregates")
        lbl_fac.setProperty('class', 'header-4')
        self.form_layout.addWidget(lbl_fac)

        form_fac = QFormLayout()
        form_fac.setContentsMargins(3, 0, 0, 0)
        form_fac.setSpacing(3)

        # Agg Quality
        self.inputs['agg_quality'] = QComboBox()
        self.inputs['agg_quality'].addItems(self.agg_map.keys())
        self.inputs['agg_quality'].setCurrentIndex(1)

        # Gravel Size
        self.inputs['gravel_size'] = QComboBox()
        self.inputs['gravel_size'].addItems(self.gravel_map.keys())
        self.inputs['gravel_size'].setCurrentIndex(1)

        form_fac.addRow("Gravel Size:", self.inputs['gravel_size'])
        form_fac.addRow("Agg. Quality:", self.inputs['agg_quality'])

        self.form_layout.addLayout(form_fac)
        self.form_layout.addStretch()

        # Reference
        lbl_ref = QLabel("Dreux-Gorisse Formula for 28th-day strength:\nFc₂₈ = G × Rc × (C/W - 0.5)\n\n"
                         "GL2000 ACI.209R Formula for maturity:\nB = exp((s/2)*(1-sqrt(28/t))\nFc_t = Fc₂₈*B²")
        lbl_ref.setProperty('class', 'formula')
        lbl_ref.setWordWrap(True)
        lbl_ref.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.form_layout.addWidget(lbl_ref)

        # --- CONNECT SIGNALS ---
        for widget in self.inputs.values():
            if isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                widget.valueChanged.connect(self.calculate_strength)
            elif isinstance(widget, QComboBox):
                widget.currentIndexChanged.connect(self.calculate_strength)

    def calculate_strength(self):
        # Guard clause for safety
        if not all(k in self.inputs for k in ['bags', 'bag_weight', 'water']): return

        try:
            # 1. Ratios
            bags = self.inputs['bags'].value()
            bag_wt = self.inputs['bag_weight'].value()
            water = self.inputs['water'].value()
            if water <= 0: return

            cw_ratio = (bags * bag_wt) / water

            # 2. Coefficients
            c_type = self.inputs['cement_type'].currentText()
            rc = self.cement_map.get(c_type, 42.5)

            agg_q = self.inputs['agg_quality'].currentText()
            g_base = self.agg_map.get(agg_q, 0.48)

            grav_s = self.inputs['gravel_size'].currentText()
            g_adj = self.gravel_map.get(grav_s, 0.00)

            G = g_base + g_adj

            # 3. 28-Day Strength
            fc_28 = max(0, G * rc * (cw_ratio - 0.5))

            # 4. Maturity Curve
            # USES GL2000 AS DESCRIBED IN ACI-209.2R-08
            # Points to plot: 3, 7, 14, 21, 28 days
            days = np.array([3, 7, 14, 21, 28])
            s = self.cement_map_S_CONSTANT_GL2000[c_type]

            # GL2000 formula
            beta_cc = np.exp(s/2 * (1 - np.sqrt(28 / days)))
            strengths = fc_28 * beta_cc**2

            # Insert 0,0 for the graph visuals
            plot_days = np.insert(days, 0, 0)
            plot_strengths = np.insert(strengths, 0, 0)

            self.update_plot(plot_days, plot_strengths)

        except Exception as e:
            print(f"Calc Error: {e}")

    def update_plot(self, x_data, y_data):
        self.ax.clear()

        # --- STYLE CONFIGURATION ---
        text_color = '#555555'  # Softer dark gray for text
        border_color = '#CCCCCC'  # Light gray for the box borders

        # --- PRIMARY AXIS (MPa) ---
        self.ax.set_xlabel("Age (Days)", color=text_color)
        self.ax.set_ylabel("Strength (MPa)", color=text_color)

        # Style the grid
        self.ax.grid(True, which='major', linestyle='--', alpha=0.5, color='#e0e0e0')

        # Style the ticks (numbers)
        self.ax.tick_params(axis='both', colors=text_color, which='both')

        # Style the Spines (The box around the chart)
        for spine in self.ax.spines.values():
            spine.set_color(border_color)

        # --- PLOT DATA ---
        self.ax.plot(x_data, y_data, color='#009580', linewidth=2, label='Maturity Curve')
        self.ax.fill_between(x_data, y_data, color='#009580', alpha=0.1)

        points_x = x_data[1:]
        points_y = y_data[1:]
        self.sc_plot = self.ax.scatter(points_x, points_y, color='white', edgecolor='#009580', s=50, zorder=5)

        # Limits
        max_y = max(y_data) * 1.2 if len(y_data) > 0 else 10
        self.ax.set_xlim(0, 30)
        self.ax.set_ylim(0, max_y)

        # --- SECONDARY AXIS (PSI) ---
        def mpa_to_psi(x):
            return x * 145.038

        def psi_to_mpa(x):
            return x / 145.038

        secax = self.ax.secondary_yaxis('right', functions=(mpa_to_psi, psi_to_mpa))
        secax.set_ylabel("Strength (psi)", color=text_color)

        # Style Secondary Axis Ticks
        secax.tick_params(axis='y', colors=text_color)

        # Style Secondary Axis Spine (The right vertical line)
        secax.spines['right'].set_color(border_color)

        # --- REFERENCE LINES ---
        ref_lines = [(3000, 20.684, '#ffc600'), (4000, 27.579, '#ff003c')]
        for psi_val, mpa_val, color in ref_lines:
            if mpa_val < max_y:
                self.ax.axhline(y=mpa_val, color=color, linestyle=':', linewidth=1.5, alpha=0.8)
                self.ax.text(0.5, mpa_val + (max_y * 0.01), f'{psi_val} psi',
                             color=color, fontsize=8, fontweight='bold')

        # --- ANNOTATION (Bottom Fixed) ---
        self.annot = self.ax.annotate(
            "",
            xy=(0, 0),
            xytext=(0, -15),
            textcoords="offset points",
            ha='center', va='top',
            # Added styling to the tooltip box as well
            bbox=dict(boxstyle="round", fc="white", ec="#cccccc", alpha=0.95),
            arrowprops=dict(arrowstyle="->", color=text_color)
        )
        self.annot.set_visible(False)

        # Layout fix
        self.figure.tight_layout()
        self.canvas.draw()

    def on_hover(self, event):
        vis = self.annot.get_visible()

        if event.inaxes == self.ax and self.sc_plot is not None:
            cont, ind = self.sc_plot.contains(event)
            if cont:
                # 1. Move anchor to the point
                pos = self.sc_plot.get_offsets()[ind["ind"][0]]
                self.annot.xy = pos

                # 2. Get Data
                day = int(pos[0])
                mpa = pos[1]
                psi = mpa * 145.038

                # 3. Force "Bottom Center" alignment
                # We reset this every time just in case the previous dynamic code changed it
                self.annot.set_position((0, -15))
                self.annot.set_ha('center')
                self.annot.set_va('top')

                # 4. Update Text
                text = f"Day: {day}\n{mpa:.1f} MPa\n{psi:,.0f} psi"
                self.annot.set_text(text)
                self.annot.get_bbox_patch().set_alpha(0.9)
                self.annot.set_visible(True)
                self.canvas.draw_idle()
                return

        if vis:
            self.annot.set_visible(False)
            self.canvas.draw_idle()

if __name__ == '__main__':
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    wheel_event_filter = GlobalWheelEventFilter()
    app.installEventFilter(wheel_event_filter)

    app.setStyleSheet(load_stylesheet('style.qss'))

    window = ConcreteMixWindow()
    window.show()
    sys.exit(app.exec())