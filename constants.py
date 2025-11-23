"""
This module serves as the single source of truth for all static
application data, such as rebar sizes, market lengths, and UI dimensions.
"""

# --- Rebar Domain Data ---
BAR_DIAMETERS = ['#10', '#12', '#16', '#20', '#25', '#28', '#32', '#36', '#40', '#50']
BAR_DIAMETERS_FOR_STIRRUPS = ['#10', '#12', '#16', '#20', '#25']
MARKET_LENGTHS = ['6m', '7.5m', '9m', '10.5m', '12m', '13.5m', '15m']

# --- UI Layout Constants ---
# It's also good practice to centralize these so UI elements are consistent.
FOOTING_IMAGE_WIDTH = 350
RSB_IMAGE_WIDTH = 165
STIRRUP_ROW_IMAGE_WIDTH = 80

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

# APP Logo
LOGO_MAP = {
    'app_cutting_list': 'images/logo_blue.png',
    'app_optimal_purchase': 'images/logo_purple.png',
    'app_concrete_mix': 'images/logo_red.png'
}
VERSION = '1.0.0'

# --- Application Configuration ---
# This is the perfect place for the debug flag.
DEBUG_MODE = False