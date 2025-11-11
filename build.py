"""
PyInstaller build script for Dual-Pane Explorer
Run: python build.py
Built by Chukwuemeka Onyebuenyi
"""

import PyInstaller.__main__
import sys
import os

# Application info
APP_NAME = "DualPaneExplorer"
VERSION = "1.0.0"
SCRIPT_NAME = "dual_pane_explorer.py"

# Build options
options = [
    SCRIPT_NAME,
    '--name=' + APP_NAME,
    '--windowed',  # No console window
    '--onefile',   # Single executable
    '--clean',     # Clean build
]

# Platform-specific options
if sys.platform == 'win32':
    # Only add icon if it exists
    if os.path.exists('icon.ico'):
        options.append('--icon=icon.ico')
    # Only add version file if it exists
    if os.path.exists('version.txt'):
        options.append('--version-file=version.txt')
elif sys.platform == 'darwin':
    if os.path.exists('icon.icns'):
        options.append('--icon=icon.icns')
    options.append('--osx-bundle-identifier=com.dualpane.explorer')

# Run PyInstaller
PyInstaller.__main__.run(options)

print(f"\n{'='*60}")
print(f"Build complete!")
print(f"Executable location: dist/{APP_NAME}")
print(f"{'='*60}")
