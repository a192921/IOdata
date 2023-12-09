import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "packages": ["tkinter", "tkcalendar", "datetime","pyppeteer"],
    "includes": ["tkinter", "numpy","os","asyncio"],
    "include_msvcr": True
    }

# base="Win32GUI" should be used only for Windows GUI app
base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="IOdata",
    version="1.0",
    description="IOdata",
    options={"build_exe": build_exe_options},
    executables=[Executable("IOdata.py", icon="favicon.ico", base=base)],
    # includes=["tkinter", "pandas", "sys", "tkcalendar", "numpy", "datetime","os", "asyncio", "pyppeteer"]
)