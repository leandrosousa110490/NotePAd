# Building Modern Notepad as a Standalone Application

This guide will help you create a standalone executable of Modern Notepad that can run on any Windows computer without requiring Python to be installed.

## Prerequisites

1. **Python 3.6 or higher** installed on your system
2. **pip** (Python package installer)
3. All project dependencies installed

## Step 1: Install Dependencies

First, make sure all required packages are installed:

```bash
pip install -r requirements.txt
```

## Step 2: Install PyInstaller

Install PyInstaller using pip:

```bash
pip install pyinstaller
```

## Step 3: Build the Executable

### Option A: Simple Build (Recommended)

Run this command in the project directory:

```bash
pyinstaller --onefile --windowed --name "ModernNotepad" notepad.py
```

### Option B: Advanced Build with Icon (if you have an icon file)

If you have an icon file (`.ico` format), use:

```bash
pyinstaller --onefile --windowed --name "ModernNotepad" --icon=icon.ico notepad.py
```

### Option C: Build with Spec File (For Advanced Users)

Create a spec file for more control:

```bash
pyinstaller --onefile --windowed --name "ModernNotepad" notepad.py --specpath=.
```

Then edit the generated `ModernNotepad.spec` file if needed and rebuild:

```bash
pyinstaller ModernNotepad.spec
```

## Command Line Options Explained

- `--onefile`: Creates a single executable file instead of a folder with multiple files
- `--windowed`: Prevents a console window from appearing (important for GUI apps)
- `--name "ModernNotepad"`: Sets the name of the output executable
- `--icon=icon.ico`: Adds a custom icon to the executable (optional)

## Step 4: Locate Your Executable

After the build completes successfully:

1. Look for a `dist` folder in your project directory
2. Inside `dist`, you'll find `ModernNotepad.exe`
3. This is your standalone application!

## Step 5: Test the Executable

1. Copy `ModernNotepad.exe` to a different location or computer
2. Double-click to run it
3. Verify all features work correctly:
   - File operations (New, Open, Save)
   - Mode switching (Normal, Code, Spreadsheet)
   - Image pasting
   - Syntax highlighting in code mode
   - Spreadsheet functionality

## Distribution

Your `ModernNotepad.exe` file is now ready for distribution! Users can:

- Run it directly without installing Python
- Copy it to any Windows computer
- Place it in their Programs folder or desktop
- Create shortcuts to it

## Troubleshooting

### Common Issues and Solutions

**Issue**: "Failed to execute script" error
- **Solution**: Try building without `--onefile` flag first to see detailed error messages

**Issue**: Missing modules error
- **Solution**: Make sure all dependencies are installed with `pip install -r requirements.txt`

**Issue**: Large file size
- **Solution**: This is normal for PyInstaller builds. The executable includes Python and all dependencies

**Issue**: Antivirus flags the executable
- **Solution**: This is common with PyInstaller. Add an exception or use `--exclude-module` for unused modules

### Advanced Optimization

To reduce file size, you can exclude unused modules:

```bash
pyinstaller --onefile --windowed --name "ModernNotepad" --exclude-module matplotlib --exclude-module numpy notepad.py
```

## File Structure After Build

```
project/
├── notepad.py
├── requirements.txt
├── README.md
├── BUILD_GUIDE.md
├── build/                 # Temporary build files (can be deleted)
├── dist/                  # Contains your executable
│   └── ModernNotepad.exe  # Your standalone application
└── ModernNotepad.spec     # Build specification (if created)
```

## Notes

- The first run might be slower as Windows scans the new executable
- The executable size will be larger (typically 20-50MB) as it includes Python runtime
- Consider creating an installer using tools like NSIS or Inno Setup for professional distribution
- Test on different Windows versions to ensure compatibility

## Creating a Desktop Shortcut

After building, you can create a desktop shortcut:

1. Right-click on `ModernNotepad.exe`
2. Select "Create shortcut"
3. Move the shortcut to your desktop
4. Rename it to "Modern Notepad"

Your Modern Notepad application is now ready for use by anyone, even without Python installed!