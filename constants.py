"""
This module serves as the single source of truth for all static
application data, such as rebar sizes, market lengths, and UI dimensions.
"""

# --- Rebar Domain Data ---
BAR_DIAMETERS = ['#10', '#12', '#16', '#20', '#25', '#28', '#32', '#36', '#40', '#50']
BAR_DIAMETERS_FOR_STIRRUPS = ['#10', '#12', '#16', '#20', '#25']
STIRRUP_SHAPES = ['Outer', 'Tall', 'Wide', 'Octagon', 'Diamond']
MARKET_LENGTHS = ['6m', '7.5m', '9m', '10.5m', '12m', '13.5m', '15m']

# --- UI Layout Constants ---
# It's also good practice to centralize these so UI elements are consistent.
FOOTING_IMAGE_WIDTH = 450
RSB_IMAGE_WIDTH = 150
STIRRUP_ROW_IMAGE_WIDTH = 80

# --- Application Configuration ---
# This is the perfect place for the debug flag.
DEBUG_MODE = True