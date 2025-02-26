# Converting Python Scripts to Standalone Windows Executables

This guide explains how to convert Python scripts like the Excel Character Scanner into standalone Windows executables that can run without requiring Python installation.

## Method 1: Using PyInstaller

PyInstaller is one of the most popular tools for converting Python scripts into standalone executables.

### Installing PyInstaller

1. **Open a command prompt** with administrator privileges
2. **Install PyInstaller** using pip:
   ```
   pip install pyinstaller
   ```

### Creating an Executable

#### Simple Approach
Navigate to the directory containing your script and run:

```bash
pyinstaller --onefile --noconsole excel-character-scanner.py
```

- `--onefile`: Creates a single executable file (instead of a directory with many files)
- `--noconsole`: Hides the console window (useful for GUI applications)

#### Advanced Approach (with Icon and Version Info)

For a more polished application:

```bash
pyinstaller --onefile --noconsole --icon=app_icon.ico --version-file=version_info.txt excel-character-scanner.py
```

- `--icon`: Sets a custom icon for your executable
- `--version-file`: Adds version information to the executable properties

### Finding Your Executable

After PyInstaller completes:
1. Look in the newly created `dist` folder
2. Your executable will be named `excel-character-scanner.exe`

## Method 2: Using Auto-Py-To-Exe (GUI for PyInstaller)

If you prefer a graphical interface, Auto-Py-To-Exe provides a user-friendly way to use PyInstaller.

### Installing Auto-Py-To-Exe

```bash
pip install auto-py-to-exe
```

### Using the GUI

1. Run the application:
   ```bash
   auto-py-to-exe
   ```

2. In the GUI:
   - Select your script file
   - Choose "One File" option
   - Select "Window Based" (no console)
   - Add any additional files (icons, data files)
   - Click "Convert"

## Method 3: Using cx_Freeze

cx_Freeze is another popular option for creating executables.

### Installing cx_Freeze

```bash
pip install cx_Freeze
```

### Creating a Setup Script

Create a file named `setup.py` in the same directory as your script:

```python
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but might need adjustments
build_exe_options = {
    "packages": ["pandas", "openpyxl", "xlrd", "tkinter", "re", "unicodedata"],
    "excludes": [],
    "include_files": []
}

setup(
    name="Excel Character Scanner",
    version="1.0",
    description="Tool to scan Excel files for problematic characters",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "excel-character-scanner.py",
            base="Win32GUI",  # Use "console" for command line apps
            icon="app_icon.ico",  # Optional
            target_name="ExcelCharScanner.exe"
        )
    ]
)
```

### Building the Executable

```bash
python setup.py build
```

Look for the executable in the `build` directory.

## Troubleshooting Common Issues

### Missing Dependencies

If your executable fails with "missing module" errors:

1. **For PyInstaller**: Use the `--hidden-import` option:
   ```bash
   pyinstaller --onefile --noconsole --hidden-import=openpyxl.cell._writer excel-character-scanner.py
   ```

2. **For cx_Freeze**: Add the module to the "packages" list in `setup.py`

### File Access Issues

If your application can't find data files:

1. **For PyInstaller**: Use the `--add-data` option:
   ```bash
   pyinstaller --onefile --noconsole --add-data "config.ini;." excel-character-scanner.py
   ```

2. **For cx_Freeze**: Add files to the "include_files" list in `setup.py`

### Large File Size

Executables created with these tools can be large (50-100MB) because they include Python and all dependencies.

To reduce size:
- Use UPX compression: `pyinstaller --onefile --noconsole --upx-dir=/path/to/upx excel-character-scanner.py`
- Use virtual environments with only required packages installed

## Specific Instructions for Excel Character Scanner

For the Excel Character Scanner with tkinter and pandas dependencies:

```bash
# Step 1: Create a virtual environment (optional but recommended)
python -m venv scanner_env
scanner_env\Scripts\activate

# Step 2: Install required packages
pip install pandas openpyxl xlrd

# Step 3: Install PyInstaller
pip install pyinstaller

# Step 4: Create the executable
pyinstaller --onefile --noconsole --hidden-import=openpyxl.cell._writer --hidden-import=tkinter excel-character-scanner.py

# Step 5: Test the executable
dist\excel-character-scanner.exe
```

## Distribution

Once you've created your executable:

1. Test it thoroughly on a system without Python installed
2. Package it with any required files (documentation, examples)
3. Consider using an installer creator like NSIS or Inno Setup for professional distribution

## Additional Resources

- [PyInstaller Documentation](https://pyinstaller.readthedocs.io/)
- [cx_Freeze Documentation](https://cx-freeze.readthedocs.io/)
- [Auto-Py-To-Exe GitHub](https://github.com/brentvollebregt/auto-py-to-exe)
