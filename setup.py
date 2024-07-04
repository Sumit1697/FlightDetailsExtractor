import sys
from cx_Freeze import setup, Executable

# Additional files to include
files = ['adblock.crx']

# Dependencies are automatically detected, but it might need fine-tuning.
build_exe_options = {
    "packages": ["os", "selenium"],
    "include_files": files,
}

# No base or set to 'Console' to ensure the console window opens
base = None
if sys.platform == "win32":
    base = "Console"

setup(
    name="FlightDetailsExtractor_v2",
    version="0.2",
    description="Extract the flight details using selenium",
    options={"build_exe": build_exe_options},
    executables=[Executable("FlightDetailsExtractor.py", base=base)],
)