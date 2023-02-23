# Import cx_Freeze
import cx_Freeze
import setuptools
import os
import sys

os.environ['TCL_LIBRARY'] = r'C:\Users\matth\anaconda3\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\matth\anaconda3\tcl\tk8.6'
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = r'C:\Users\matth\anaconda3\Library\plugins\platforms'
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = r'C:\Users\matth\anaconda3\Library\plugins'

options = {
    "build_exe": {
        "copy_dependent_files": True,
        "compressor": "upx"
    }
}

# Set up the executables
executables = [cx_Freeze.Executable("JN_OP_Tool.py")]

# Set up the build options
build_options = {
    "packages": ["warnings", "pandas", "matplotlib", "sys", "os", "tkinter", 
    "openpyxl", "jdcal", "et_xmlfile", "PyQt5.QtCore", "PyQt5.QtGui", "PyQt5.QtWidgets", 
    "PyQt5.QtSvg", "PyQt5.QtPrintSupport", "atexit", "xlsxwriter", "glob"],
    
    "include_files": [
        "QUERY_FILE_GOES_HERE",
        r"C:\Users\matth\anaconda3\Library\plugins\platforms\qwindows.dll",
        r"C:\Users\matth\anaconda3\Library\bin\mkl_intel_thread.1.dll",
        r"C:\Users\matth\anaconda3\Library\bin\mkl_def.1.dll",
        r"C:\Users\matth\anaconda3\pkgs\mkl-2021.4.0-haa95532_640\Library\bin\mkl_sequential.1.dll"
    ]
}

sys.path.append(r'C:\Users\matth\anaconda3\Library\plugins')

# Run the build process
cx_Freeze.setup(
    name = "JN_OP_Tool",
    options = {"build_exe": build_options},
    executables = executables
)
