import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable("main.py", base=base, icon="1.ico")
]

setup(
    name="EXCEL2PDF4WB",
    version="0.1",
    description="Excel to PDF for Wildberries Pallet",
    executables=executables
)
