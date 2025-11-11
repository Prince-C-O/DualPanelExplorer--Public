# Dual-Pane File Explorer

A modern, feature-rich dual-pane file manager built with Python and Tkinter with **MTP Device Support** for managing files on smartphones, tablets, and portable devices.

## Features

### Core Functionality
- **Dual-Panel Interface**: Side-by-side file browsing for easy file management
- **File Operations**: Copy, move, delete files and folders
- **Smart Navigation**: History-based back/forward navigation
- **Search**: Fast file search within directories
- **Archive Support**: Create and extract ZIP archives
- **🆕 MTP Device Support**: Browse and manage files on connected portable devices (phones, tablets, cameras)

### MTP Features (New!)
- **Device Detection**: Automatically detects connected MTP devices
- **Seamless Integration**: MTP devices appear in drive selector with 📱 icon
- **File Operations**: Copy, move, delete files between PC and MTP devices
- **Cross-Device**: Transfer files between multiple MTP devices
- **Open Files**: View files from MTP devices in their default applications


### User Interface
- **Modern Design**: Clean, intuitive interface
- **Themes**: Dark and light theme support
- **Icons**: File type-specific icons
- **Keyboard Shortcuts**: Extensive keyboard support for power users
- **Drag & Drop**: Drag files between panels
- **Preview Pane**: Preview text files, images, and view directory summaries (Ctrl+P to toggle)

### Advanced Features
- **Folder Size Calculation**: Async calculation without UI blocking
- **File Properties**: Detailed file/folder information
- **Progress Tracking**: Visual progress for long operations
- **Error Handling**: Graceful error handling with user feedback
- **Settings Persistence**: Remember your preferences
- **Drive Information**: Real-time drive space monitoring
- **Context Menus**: Right-click for quick actions (copy, move, delete, new file/folder)

## Installation

### From Source

1. Clone or download the repository
2. Ensure Python 3.9+ is installed
3. Install MTP support (Windows only): `pip install pywin32`
4. Run: `python dual_pane_explorer.py`

### Requirements
- Python 3.9 or later
- pywin32 (for MTP device support on Windows)
- Standard library (tkinter, pathlib, etc.)

### Building Executable

1. Install PyInstaller: `pip install pyinstaller`
2. Run build script: `python build.py`
3. Find executable in `dist/` folder

## Usage

### Keyboard Shortcuts

**Navigation**:
- `Tab` - Switch between panels
- `Backspace` - Go to parent directory
- `Enter` - Open folder/file
- `Ctrl+A` - Select all items

**File Operations**:
- `Ctrl+C` - Copy selected items
- `Ctrl+X` - Move selected items
- `Delete` - Delete selected items
- `Ctrl+N` - Create new folder
- `Ctrl+Z` - Create archive

**View & Search**:
- `F5` - Refresh panels
- `Ctrl+F` - Search in current folder
- `Ctrl+T` - Toggle theme
- `Alt+Enter` - Show properties

**Application**:
- `Ctrl+,` - Settings
- `Ctrl+Q` - Quit

### Mouse Operations

- **Double-click** folder to navigate
- **Double-click** file to open with default app
- **Right-click** for context menu
- **Drag** files between panels to copy
- **Click** in path bar to edit location

## Configuration

Settings are stored in:
- **Windows**: `%USERPROFILE%\.dualpane-explorer\config.json`
- **macOS**: `~/.dualpane-explorer/config.json`
- **Linux**: `~/.dualpane-explorer/config.json`

## Building Executables

### Windows
```bash
python build.py
```
Produces: `dist/DualPaneExplorer.exe`

### macOS
```bash
python build.py
```
Produces: `dist/DualPaneExplorer.app`

### Linux
```bash
python build.py
```
Produces: `dist/DualPaneExplorer`

## System Requirements

- Python 3.8 or higher
- Tkinter (included with Python)
- 50 MB disk space
- 100 MB RAM minimum

## Platform Support

- ✅ Windows 10/11
- ✅ macOS 10.14+
- ✅ Linux (Ubuntu 20.04+, Fedora 35+, etc.)

## License

MIT License - Feel free to use, modify, and distribute.

## Credits

Built by: Chukwuemeka Onyebuenyi with Python and Tkinter
Icons: Unicode Emoji

## Helpful? 
Feel free to support me. You can reachout via www.linkedin.com/in/chukwuemeka-onyebuenyi-967304201

