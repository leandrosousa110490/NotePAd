# Modern Notepad

A powerful, multi-mode text editor built with Python and Tkinter that combines the simplicity of a notepad with advanced features like code editing and spreadsheet functionality.

## Quick Start

### Running from Source
1. Install Python 3.6 or higher
2. Install dependencies: `pip install -r requirements.txt`
3. Run the application: `python notepad.py`

### Building Standalone Executable
Want to create an executable that runs without Python? See our detailed [BUILD_GUIDE.md](BUILD_GUIDE.md) for step-by-step instructions using PyInstaller.

## Features

### Core Features
- **Modern UI**: Dark theme with clean, minimalist design
- **Multi-Mode Editor**: Switch between Normal, Code, and Spreadsheet modes
- **Stay on Top**: Toggle button to keep the notepad window always on top
- **File Operations**: New, Open, Save, and Save As with smart file type detection
- **Right-Click Context Menu**: Copy, paste, cut, and select all operations
- **Image Support**: Paste images directly from clipboard into the text area
- **Status Bar**: Shows cursor position, total line count, and current mode

### Code Editor Mode
- **Syntax Highlighting**: Support for 20+ programming languages
- **Line Numbers**: Automatic line numbering with scroll synchronization
- **Language Detection**: Smart file extension mapping
- **Supported Languages**: Python, JavaScript, HTML, CSS, Java, C++, C#, PHP, Ruby, Go, Swift, TypeScript, SQL, Rust, Kotlin, Bash, PowerShell, XML, JSON, YAML

### Spreadsheet Mode
- **Excel-like Grid**: Interactive spreadsheet with resizable columns
- **Cell Navigation**: Arrow key navigation and click-to-select
- **Data Export**: Save as Excel (.xlsx) or CSV format
- **Column Headers**: A-Z column labeling with row numbers
- **Clipboard Support**: Copy and paste between cells

### Keyboard Shortcuts
- **File Operations**:
  - Ctrl+N: New file
  - Ctrl+O: Open file
  - Ctrl+S: Save file
  - Ctrl+Shift+S: Save as
- **Edit Operations**:
  - Ctrl+C: Copy
  - Ctrl+V: Paste
  - Ctrl+X: Cut
  - Ctrl+A: Select all
- **Spreadsheet Navigation**:
  - Arrow Keys: Move between cells
  - Tab: Move to next cell
  - Shift+Tab: Move to previous cell
  - Enter: Move to cell below

## Requirements

- Python 3.6 or higher
- Tkinter (included in standard Python installation)
- Pillow (PIL) for image handling

## How to Run

1. Make sure you have Python installed on your system
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

3. Navigate to the project directory
4. Run the application:

```bash
python notepad.py
```

## Usage

### File Operations
- **Creating a New File**: File menu → New or press Ctrl+N
- **Opening a File**: File menu → Open or press Ctrl+O
- **Saving a File**: File menu → Save or press Ctrl+S
- **Saving as a New File**: File menu → Save As or press Ctrl+Shift+S

### Edit Operations
- **Copy Text**: Select text and right-click → Copy or press Ctrl+C
- **Paste Content**: Right-click → Paste or press Ctrl+V (supports text and images)
- **Cut Text**: Select text and right-click → Cut or press Ctrl+X
- **Select All**: Right-click → Select All or press Ctrl+A

### Special Features
- **Toggle Stay on Top**: Click the "📌 Stay on Top" button in the top-right corner
- **Image Pasting**: Copy an image to clipboard and paste it directly into the text area
- **Context Menu**: Right-click anywhere in the text area for quick access to edit operations

## Customization

You can customize the appearance by modifying the style settings in the `setup_styles()` method in the `notepad.py` file.