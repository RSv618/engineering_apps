import os
import sys
import pulp
import subprocess

# -----------------------------------------------------------------------------
# 1. DYNAMIC PATH CONFIGURATION (Same as your .spec file)
# -----------------------------------------------------------------------------
pulp_dir = os.path.dirname(pulp.__file__)

# Check 64-bit
cbc_path = os.path.join(pulp_dir, 'solverdir', 'cbc', 'win', 'i64', 'cbc.exe')

# Fallback check for 32-bit
if not os.path.exists(cbc_path):
    cbc_path = os.path.join(pulp_dir, 'solverdir', 'cbc', 'win', '32', 'cbc.exe')

if not os.path.exists(cbc_path):
    raise FileNotFoundError(f"Could not find cbc.exe at {cbc_path}.")

print(f"INFO: Found CBC solver at: {cbc_path}")

# -----------------------------------------------------------------------------
# 2. CONSTRUCT NUITKA COMMAND
# -----------------------------------------------------------------------------
# We map cbc_path (source) to "cbc.exe" (destination in the root of the build)
# Syntax: --include-data-file=SOURCE=DEST

nuitka_cmd = [
    sys.executable, "-m", "nuitka",
    "--onefile",  # vs standalone  |  Build a folder (prevents virus flags)
    "--plugin-enable=pyqt6",  # Enable PyQt6 support
    "--windows-disable-console",  # Hide the black terminal window
    "--enable-plugin=anti-bloat",  # Reduce size

    # OUTPUT DIRECTORY
    "--output-dir=build_nuitka",
    "--remove-output",  # Clean up intermediate build files

    # ICON (Nuitka prefers .ico, but newer versions handle png sometimes.
    # If this fails, convert logo.png to logo.ico)
    "--windows-icon-from-ico=images/logo.png",

    # DATA FILES
    "--include-data-dir=images=images",
    "--include-data-file=style.qss=style.qss",
    "--include-data-file=version_info.txt=version_info.txt",

    # THE SOLVER (Dynamically Found)
    f"--include-data-file={cbc_path}=cbc.exe",

    # MAIN SCRIPT
    "app_launcher.py"
]

# -----------------------------------------------------------------------------
# 3. RUN BUILD
# -----------------------------------------------------------------------------
print("Starting Nuitka build...")
print(" ".join(nuitka_cmd))
subprocess.run(nuitka_cmd, check=True)
print("Build Complete. Check the 'build_nuitka/app_launcher.dist' folder.")