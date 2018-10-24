from cx_Freeze import setup, Executable  
import sys
import os

build_exe_options = {"packages": ["idna","tkinter"], "include_files": ["exe/tcl86t.dll", "exe/tk86t.dll", "client_secret.json"]}  

base = None  
if sys.platform == "win32":  
    base = "Win32GUI"  

os.chdir("..")

setup(
    name="pdf2calendar",  
    version="1.0",  
    description="Description",  
    options={"build_exe": build_exe_options},  
    executables=[Executable("main.py", base=base)],    
    )  

