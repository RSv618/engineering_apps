# -*- mode: python ; coding: utf-8 -*-

import sys
import os
import pulp

# -----------------------------------------------------------------------------
# DYNAMIC PATH CONFIGURATION
# -----------------------------------------------------------------------------

# 1. Locate the PuLP CBC executable dynamically.
# This prevents hardcoding user paths (C:\Users\...) and makes the spec file portable.
pulp_dir = os.path.dirname(pulp.__file__)

# Common path for CBC in Windows installations of PuLP (64-bit)
cbc_path = os.path.join(pulp_dir, 'solverdir', 'cbc', 'win', 'i64', 'cbc.exe')

# Fallback check for 32-bit if 64-bit doesn't exist
if not os.path.exists(cbc_path):
    cbc_path = os.path.join(pulp_dir, 'solverdir', 'cbc', 'win', '32', 'cbc.exe')

if not os.path.exists(cbc_path):
    raise FileNotFoundError(f"Could not find cbc.exe at {cbc_path}. Please check your PuLP installation.")

print(f"INFO: Including solver executable from: {cbc_path}")

# 2. Define Binaries
# Format: (Source Path, Destination Path in Bundle)
# We place cbc.exe in '.' (root) so your get_solver_path() finds it at sys._MEIPASS/cbc.exe
my_binaries = [
    (cbc_path, '.')
]

# 3. Define Data
# Format: (Source Path, Destination Path in Bundle)
my_datas = [
    ('images', 'images'),   # Copy local 'images' folder to 'images' in bundle
    ('style.qss', '.'),     # Copy style.qss to root
]

# -----------------------------------------------------------------------------
# PYINSTALLER ANALYSIS
# -----------------------------------------------------------------------------

block_cipher = None

a = Analysis(
    ['app_launcher.py'],         # Your main entry point
    pathex=[],
    binaries=my_binaries,
    datas=my_datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# -----------------------------------------------------------------------------
# EXE BUILD
# -----------------------------------------------------------------------------

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='EngineerApps',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,               # False = Windowed mode (no black box)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='images/logo.png',      # The .exe file icon
    version='version_info.txt'   # Includes your version info file
)