import sys
from cx_Freeze import setup, Executable


# base="Win32GUI" should be used only for Windows GUI app
# base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="IOdata",
    version="1.0",
    description="IOdata",
    executables=[Executable("IOdata.py", icon="favicon.ico", base="Win32GUI")],
)