import sys
from cx_Freeze import setup, Executable

# Name of your main script
target_script = 'laybyreport.py'

# Build options
build_options = {
    'packages': ['os', 'tkinter', 'pandas', 'pyodbc', 'configparser'],
    'excludes': [],
    'include_files': ['config.ini'],  # Include the config.ini file
}

# Executable options
executables = [
    Executable(
        script=target_script,  # The name of your main script
        base="Win32GUI",  # Use "Win32GUI" to create a GUI executable without the console window
        targetName='Layby Info Report.exe',  # The name of the executable
    )
]

# Create the setup
setup(
    name='Layby Info Report',  # Name of the application
    version='1.0',
    description='Layby Info Report',
    options={'build_exe': build_options},
    executables=executables,
)
