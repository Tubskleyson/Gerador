from cx_Freeze import setup, Executable

base = "Win32GUI"

includes      = []
include_files = [r"C:\Program Files (x86)\Python36-32\DLLs\tcl86t.dll", \
                 r"C:\Program Files (x86)\Python36-32\DLLs\tk86t.dll"]

setup(
    name = "Certificados",
    version = "0.1",
    options = {"build_exe": {"includes": includes, "include_files": include_files}},
    executables = [Executable("interface.pyw", base=base,icon='assets/certi.ico')]
)
