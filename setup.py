# -*- coding: shift_jis -*-

import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["sys", "os", "csv"]}


base = None
if sys.platform == "win32":
    base = "Console"
#    base = "Win32GUI"

setup(  name = "proc_csv",
        version = "0.1",
        description = "説明",
        options = {"build_exe": build_exe_options},
        executables = [Executable("proc_csv.py", base=base)])
