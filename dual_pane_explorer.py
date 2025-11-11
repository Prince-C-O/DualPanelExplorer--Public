"""
Dual-Pane File Explorer
Built by Chukwuemeka Onyebuenyi
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from pathlib import Path
import shutil
import os
import threading
import time
from datetime import datetime
from typing import List, Optional, Tuple, Callable, Dict
from dataclasses import dataclass
import queue
import json
import zipfile
import mimetypes
import platform
import re


# MTP Device Handler
# ============================================================================

class MTPDevice:
    """Handler for MTP (portable) device operations using Windows Shell COM APIs."""

    def __init__(self, device_name: str):
        """Initialize MTP device handler."""
        self.device_name = device_name
        self.shell = None
        self.device_folder = None

        if platform.system() == 'Windows':
            try:
                import win32com.client
                self.shell = win32com.client.Dispatch("Shell.Application")
                self._find_device()
            except Exception as e:
                print(f"Error initializing MTP handler: {e}")

    def _find_device(self):
        """Find the portable device in Windows Shell namespace."""
        try:
            computer = self.shell.Namespace(17)  # ssfDRIVES = 17
            if computer:
                for item in computer.Items():
                    if self.device_name.lower() in item.Name.lower():
                        self.device_folder = item
                        return True
        except Exception as e:
            print(f"Error finding device: {e}")
        return False

    def is_available(self) -> bool:
        """Check if device is available and accessible."""
        return self.device_folder is not None

    def get_storage_info(self, folder_path: str = ""):
        """Get storage information for MTP device or storage partition."""
        try:
            # Navigate to the specified folder (or device root)
            target_folder = self._navigate_to_folder(folder_path)
            if not target_folder:
                target_folder = self.device_folder

            if not target_folder:
                return None

            folder_obj = target_folder.GetFolder

            # Try to get storage info through extended properties
            try:
                # Try getting capacity and free space
                total_bytes = folder_obj.GetDetailsOf(None, -1)  # Capacity
                free_bytes = folder_obj.GetDetailsOf(None, -2)   # Free space

                if total_bytes and free_bytes:
                    return {
                        'total': self._parse_size_string(str(total_bytes)),
                        'free': self._parse_size_string(str(free_bytes))
                    }
            except:
                pass

            # Use Alternative method to get storage info from the folder item itself
            try:
                # try to get size info for storage partitions,
                total_str = target_folder.ExtendedProperty("System.Capacity")
                free_str = target_folder.ExtendedProperty("System.FreeSpace")

                if total_str and free_str:
                    return {
                        'total': int(total_str),
                        'free': int(free_str)
                    }
            except:
                pass

            # Return None to indicate when storage info unavailable
            return None

        except Exception as e:
            print(f"Error getting MTP storage info: {e}")
            return None

    def list_contents(self, folder_path: str = "") -> List[Dict[str, any]]:
        """List contents of a folder on the MTP device."""
        items = []
        try:
            current_folder = self._navigate_to_folder(folder_path)
            if not current_folder:
                return items

            for item in current_folder.GetFolder.Items():
                internal_path = self._build_path(folder_path, item.Name)
                full_path = f"mtp:::{self.device_name}/{internal_path}"

                # Get file size
                file_size = 0
                if not item.IsFolder:
                    try:
                        # Check direct Size property first
                        file_size = item.Size
                    except:
                        pass

                    # try ExtendedProperty if size is 0 or failed
                    if file_size == 0:
                        try:
                            # Try to get size through extended properties
                            size_str = item.ExtendedProperty("System.Size")
                            if size_str:
                                file_size = int(size_str)
                        except:
                            pass

                    # try GetDetailsOf if it is still 0
                    if file_size == 0:
                        try:
                            folder = current_folder.GetFolder
                            size_str = folder.GetDetailsOf(
                                item, 2)  # Column 2 is usually size
                            if size_str:
                                # Parse size string (might be formatted like "1.5 MB")
                                file_size = self._parse_size_string(size_str)
                        except:
                            pass

                # Try to get modified date
                modified_time = None
                try:
                    modified_time = item.ModifyDate
                    if modified_time:
                        # Convert from COM date to timestamp
                        import pywintypes
                        modified_time = int(pywintypes.Time(
                            modified_time).timestamp())
                except:
                    pass

                item_info = {
                    'name': item.Name,
                    'path': full_path,
                    'is_dir': item.IsFolder,
                    'size': file_size,
                    'modified': modified_time,
                    'item_object': item
                }
                items.append(item_info)

        except Exception as e:
            print(f"Error listing MTP contents: {e}")

        return items

    def _navigate_to_folder(self, folder_path: str):
        """Navigate to a specific folder within the device."""
        if not self.device_folder:
            return None

        if not folder_path or folder_path == "":
            return self.device_folder

        try:
            current = self.device_folder
            parts = folder_path.split('/')

            for part in parts:
                if not part:
                    continue

                found = False
                folder_obj = current.GetFolder

                for item in folder_obj.Items():
                    if item.Name == part and item.IsFolder:
                        current = item
                        found = True
                        break

                if not found:
                    return None

            return current

        except Exception as e:
            print(f"Error navigating to folder: {e}")
            return None

    def _build_path(self, base_path: str, name: str) -> str:
        """Build a path string."""
        if not base_path:
            return name
        return f"{base_path}/{name}"

    def _parse_size_string(self, size_str: str) -> int:
        """
        Parse a formatted size string (e.g., "1.5 MB", "500 KB") to bytes.

        Args:
            size_str: Formatted size string

        Returns:
            Size in bytes
        """
        try:
            # Remove commas and whitespace
            size_str = size_str.replace(',', '').strip()

            # Check if it's already just a number
            if size_str.isdigit():
                return int(size_str)

            # Parse formatted size (e.g., "1.5 MB")
            import re
            match = re.match(r'([\d.]+)\s*([KMGT]?B)', size_str, re.IGNORECASE)
            if match:
                value = float(match.group(1))
                unit = match.group(2).upper()

                multipliers = {
                    'B': 1,
                    'KB': 1024,
                    'MB': 1024 * 1024,
                    'GB': 1024 * 1024 * 1024,
                    'TB': 1024 * 1024 * 1024 * 1024
                }

                return int(value * multipliers.get(unit, 1))
        except:
            pass

        return 0

    def copy_from_device(self, source_path: str, dest_path: str) -> bool:
        """Copy file from MTP device to local PC."""
        try:
            parts = source_path.split('/')
            if not parts:
                print(f"Invalid source path: {source_path}")
                return False

            filename = parts[-1]
            folder_path = '/'.join(parts[:-1])

            print(f"Copying from MTP: {filename} from folder: {folder_path}")

            source_folder = self._navigate_to_folder(folder_path)
            if not source_folder:
                print(f"Could not navigate to folder: {folder_path}")
                return False

            source_item = None
            folder_obj = source_folder.GetFolder

            for item in folder_obj.Items():
                if item.Name == filename:
                    source_item = item
                    break

            if not source_item:
                print(f"File not found in device: {filename}")
                return False

            dest_dir = os.path.dirname(dest_path)
            if not dest_dir:
                dest_dir = dest_path

            print(f"Copying to: {dest_dir}")
            dest_folder = self.shell.Namespace(dest_dir)

            if not dest_folder:
                print(
                    f"Could not create destination folder object: {dest_dir}")
                return False

            # Check file size for large file handling
            file_size = 0
            try:
                file_size = source_item.Size
                if file_size > 100 * 1024 * 1024:  # Files larger than 100MB
                    print(
                        f"Copying large file ({file_size / (1024*1024):.1f} MB): {filename}")
            except:
                pass

            # CopyHere will copy with original filename
            # Flags: 4=no UI, 16=yes to all
            dest_folder.CopyHere(source_item, 4 | 16)

            # Wait longer for large files to complete
            wait_time = 0.5
            if file_size > 100 * 1024 * 1024:
                # Scale with size
                wait_time = min(5.0, 0.5 + (file_size / (100 * 1024 * 1024)))
            time.sleep(wait_time)

            # Verify the file was copied
            expected_file = os.path.join(dest_dir, filename)
            if os.path.exists(expected_file):
                print(f"Successfully copied to: {expected_file}")
                return True
            else:
                print(f"File not found after copy: {expected_file}")
                return False

        except Exception as e:
            print(f"Error copying from device: {e}")
            import traceback
            traceback.print_exc()
            return False

    def copy_to_device(self, source_path: str, dest_folder_path: str = "") -> bool:
        """Copy file from local PC to MTP device."""
        try:
            if not os.path.exists(source_path):
                return False

            dest_folder = self._navigate_to_folder(dest_folder_path)
            if not dest_folder:
                return False

            folder_obj = dest_folder.GetFolder

            # Use flags: 4=no UI, 16=yes to all, 1024=don't display progress
            # For large files, we want to show progress, so we use 4|16 (no UI dialogs but allow progress), also adding a small delay for large files to ensure proper handling
            file_size = os.path.getsize(source_path)
            if file_size > 100 * 1024 * 1024:  # Files larger than 100MB
                print(
                    f"Copying large file ({file_size / (1024*1024):.1f} MB): {os.path.basename(source_path)}")

            folder_obj.CopyHere(source_path, 4 | 16)

            # For large files, give Windows time to complete the operation
            if file_size > 100 * 1024 * 1024:
                time.sleep(0.5)

            return True

        except Exception as e:
            print(f"Error copying to device: {e}")
            return False

    def delete_from_device(self, file_path: str) -> bool:
        """Delete file from MTP device."""
        try:
            print(f"DEBUG: Attempting to delete: {file_path}")

            parts = file_path.split('/')
            if not parts:
                return False

            filename = parts[-1]
            folder_path = '/'.join(parts[:-1])

            print(f"DEBUG: Filename: {filename}, Folder path: {folder_path}")

            # Find the file
            source_folder = self._navigate_to_folder(folder_path)
            if not source_folder:
                print(f"Could not navigate to folder: {folder_path}")
                return False

            folder_obj = source_folder.GetFolder

            # Find the item to delete
            for item in folder_obj.Items():
                if item.Name == filename:
                    print(f"DEBUG: Found item: {item.Name}")

                    # Try multiple delete methods (matching mtp_handler.py)
                    try:
                        # Method 1: InvokeVerb on the item itself
                        item.InvokeVerb("delete")
                        print(f"Successfully deleted via Method 1: {filename}")
                        return True
                    except Exception as e1:
                        print(f"Method 1 (InvokeVerb) failed: {e1}")

                    try:
                        # Method 2: Use folder's ParseName and InvokeVerb
                        file_item = folder_obj.ParseName(filename)
                        if file_item:
                            file_item.InvokeVerb("delete")
                            print(
                                f"Successfully deleted via Method 2: {filename}")
                            return True
                    except Exception as e2:
                        print(f"Method 2 (ParseName) failed: {e2}")

                    try:
                        # Method 3: Use namespace to delete
                        if self.shell:
                            namespace = self.shell.Namespace(
                                source_folder.GetFolder.Self.Path)
                            if namespace:
                                item_to_delete = namespace.ParseName(filename)
                                if item_to_delete:
                                    item_to_delete.InvokeVerb("delete")
                                    print(
                                        f"Successfully deleted via Method 3: {filename}")
                                    return True
                    except Exception as e3:
                        print(f"Method 3 (Namespace) failed: {e3}")

                    print(f"All delete methods failed for: {filename}")
                    return False

            print(f"File not found: {filename}")
            return False

        except Exception as e:
            print(f"Error deleting from device: {e}")
            return False


def get_mtp_devices() -> List[Dict[str, str]]:
    """Get list of connected MTP devices."""
    devices = []

    if platform.system() != 'Windows':
        return devices

    try:
        import win32com.client
        shell = win32com.client.Dispatch("Shell.Application")
        computer = shell.Namespace(17)

        if computer:
            for item in computer.Items():
                try:
                    # MTP devices have paths that start with ::{ or contain usb# or are empty while regular drives have paths like "C:\" which end with backslash
                    path = item.Path
                    is_regular_drive = path and (
                        path.endswith(':\\') or path.endswith(':/'))

                    if not is_regular_drive:
                        # This is likely an MTP/portable device
                        devices.append({
                            'name': item.Name,
                            'path': f"mtp:::{item.Name}",
                            'type': 'Portable'
                        })
                except:
                    pass

    except Exception as e:
        print(f"Error getting MTP devices: {e}")

    return devices


def is_mtp_path(path: str) -> bool:
    """Check if path is an MTP device path."""
    return str(path).startswith('mtp:::')


def parse_mtp_path(path: str) -> Tuple[Optional[str], Optional[str]]:
    """Parse MTP path into device name and internal path."""
    if not is_mtp_path(path):
        return None, None

    path = str(path).replace('mtp:::', '', 1)

    if '/' in path:
        parts = path.split('/', 1)
        return parts[0], parts[1]
    else:
        return path, ""


# Configuration & Settings
# ============================================================================

class AppConfig:
    """Application configuration and settings"""

    CONFIG_DIR = Path.home() / ".dualpane-explorer"
    CONFIG_FILE = CONFIG_DIR / "config.json"

    DEFAULT_CONFIG = {
        "theme": "dark",
        "show_hidden": False,
        "confirm_delete": True,
        "confirm_overwrite": True,
        "window_geometry": "1400x800",
        "last_left_path": str(Path.home()),
        "last_right_path": str(Path.home()),
        "font_size": 10
    }

    def __init__(self):
        self.config = self.DEFAULT_CONFIG.copy()
        self.load()

    def load(self):
        """Load configuration from file"""
        try:
            if self.CONFIG_FILE.exists():
                with open(self.CONFIG_FILE, 'r') as f:
                    loaded = json.load(f)
                    self.config.update(loaded)
        except Exception as e:
            print(f"Error loading config: {e}")

    def save(self):
        """Save configuration to file"""
        try:
            self.CONFIG_DIR.mkdir(exist_ok=True)
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")

    def get(self, key: str, default=None):
        """Get configuration value"""
        return self.config.get(key, default)

    def set(self, key: str, value):
        """Set configuration value"""
        self.config[key] = value


# Data Models
# ============================================================================

@dataclass
class FileInfo:
    """File information structure"""
    path: Path
    name: str
    is_directory: bool
    size: int = 0
    modified: Optional[datetime] = None
    item_count: int = 0
    extension: str = ""


@dataclass
class OperationResult:
    """Result of a file operation"""
    success: bool
    items_processed: int
    errors: List[str]


# Utility Functions
# ============================================================================

def format_size(size: int) -> str:
    """Format file size in human-readable format"""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size < 1024.0:
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} PB"


def format_date(timestamp: float) -> str:
    """Format timestamp to readable date"""
    dt = datetime.fromtimestamp(timestamp)
    return dt.strftime("%Y-%m-%d %H:%M")


def get_folder_size(path: Path) -> Tuple[int, int]:
    """
    Calculate folder size and item count.
    Returns (total_size, item_count)
    """
    total_size = 0
    item_count = 0

    try:
        for entry in os.scandir(path):
            item_count += 1
            if entry.is_file(follow_symlinks=False):
                total_size += entry.stat().st_size
            elif entry.is_dir(follow_symlinks=False):
                sub_size, sub_count = get_folder_size(Path(entry.path))
                total_size += sub_size
                item_count += sub_count
    except (PermissionError, OSError):
        pass

    return total_size, item_count


def get_file_icon(path: Path) -> str:
    """Get appropriate icon for file type"""
    if path.is_dir():
        return "📁"

    ext = path.suffix.lower()
    icon_map = {
        '.txt': '📄', '.doc': '📄', '.docx': '📄', '.pdf': '📕',
        '.jpg': '🖼️', '.jpeg': '🖼️', '.png': '🖼️', '.gif': '🖼️', '.bmp': '🖼️',
        '.mp3': '🎵', '.wav': '🎵', '.flac': '🎵', '.ogg': '🎵',
        '.mp4': '🎬', '.avi': '🎬', '.mkv': '🎬', '.mov': '🎬',
        '.zip': '📦', '.rar': '📦', '.7z': '📦', '.tar': '📦', '.gz': '📦',
        '.py': '🐍', '.js': '📜', '.html': '🌐', '.css': '🎨',
        '.exe': '⚙️', '.app': '⚙️', '.deb': '⚙️', '.dmg': '⚙️'
    }

    return icon_map.get(ext, '📄')


def search_files(root_path, pattern: str, show_hidden: bool = False) -> List:
    """Search for files matching pattern (supports both filesystem and MTP)"""
    results = []
    pattern = pattern.lower()

    try:
        root_path_str = str(root_path)

        # Check if it's an MTP path
        if is_mtp_path(root_path_str):
            # MTP device search
            device_name, internal_path = parse_mtp_path(root_path_str)
            if device_name:
                mtp_device = MTPDevice(device_name)
                if mtp_device.is_available():
                    results.extend(_search_mtp_recursive(
                        mtp_device, internal_path or "", pattern, show_hidden))
        else:
            # Regular filesystem search
            for entry in os.scandir(root_path):
                try:
                    # Skip hidden files if needed
                    if not show_hidden and entry.name.startswith('.'):
                        continue

                    # Check if name matches pattern
                    if pattern in entry.name.lower():
                        results.append(Path(entry.path))

                    # Recursively search subdirectories
                    if entry.is_dir():
                        results.extend(search_files(
                            Path(entry.path), pattern, show_hidden))
                except (PermissionError, OSError):
                    continue
    except (PermissionError, OSError):
        pass

    return results


def _search_mtp_recursive(mtp_device: MTPDevice, folder_path: str, pattern: str, show_hidden: bool) -> List[str]:
    """Recursively search MTP device for files matching pattern"""
    results = []

    try:
        items = mtp_device.list_contents(folder_path)

        for item in items:
            # Skip hidden files if needed
            if not show_hidden and item['name'].startswith('.'):
                continue

            # Check if name matches pattern
            if pattern in item['name'].lower():
                results.append(item['path'])  # Full MTP path

            # Recursively search subdirectories
            if item['is_dir']:
                # Extract internal path from full MTP path
                _, internal_path = parse_mtp_path(item['path'])
                if internal_path:
                    results.extend(_search_mtp_recursive(
                        mtp_device, internal_path, pattern, show_hidden))
    except Exception as e:
        print(f"Error searching MTP folder {folder_path}: {e}")

    return results


# File Operations
# ============================================================================

class FileOperations:
    """Handles file operations"""

    @staticmethod
    def _copy_file_chunked(src: Path, dst: Path, buffer_size: int = 1024 * 1024) -> None:
        """Copy a file in chunks for better performance with large files"""
        with open(src, 'rb') as fsrc:
            with open(dst, 'wb') as fdst:
                while True:
                    chunk = fsrc.read(buffer_size)
                    if not chunk:
                        break
                    fdst.write(chunk)

        # Copy metadata (timestamps, permissions)
        shutil.copystat(src, dst)

    @staticmethod
    def _copy_function_smart(src: str, dst: str) -> str:
        """Smart copy function that uses chunked copy for large files"""
        src_path = Path(src)
        dst_path = Path(dst)

        try:
            file_size = src_path.stat().st_size
            if file_size > 10 * 1024 * 1024:  # 10MB threshold
                FileOperations._copy_file_chunked(src_path, dst_path)
            else:
                shutil.copy2(src, dst)
        except:
            # Fallback to standard copy
            shutil.copy2(src, dst)

        return dst

    @staticmethod
    def copy_files(sources: List[Path], destination: Path,
                   progress_callback: Optional[Callable] = None) -> OperationResult:
        """Copy files to destination (supports MTP)"""
        errors = []
        processed = 0

        dest_str = str(destination)
        dest_is_mtp = is_mtp_path(dest_str)

        for source in sources:
            try:
                source_str = str(source)
                source_is_mtp = is_mtp_path(source_str)

                # Determine destination path
                if dest_is_mtp:
                    dest_device, dest_internal = parse_mtp_path(dest_str)
                    file_name = source_str.split(
                        '/')[-1] if source_is_mtp else source.name
                    dest_path_str = f"mtp:::{dest_device}/{dest_internal}/{file_name}" if dest_internal else f"mtp:::{dest_device}/{file_name}"
                else:
                    dest_path = destination / \
                        (source_str.split('/')[-1]
                         if source_is_mtp else source.name)

                # Handle MTP operations
                if source_is_mtp and dest_is_mtp:
                    # MTP to MTP - copy through temp file
                    source_device, source_internal = parse_mtp_path(source_str)
                    dest_device, dest_internal = parse_mtp_path(dest_str)
                    mtp_source = MTPDevice(source_device)
                    mtp_dest = MTPDevice(dest_device)

                    temp_dir = Path(os.environ.get('TEMP', '/tmp'))
                    temp_file = temp_dir / source_str.split('/')[-1]

                    # Use internal paths, not full mtp::: paths
                    mtp_source.copy_from_device(
                        source_internal or "", str(temp_file))
                    mtp_dest.copy_to_device(
                        str(temp_file), dest_internal or "")
                    temp_file.unlink()

                elif source_is_mtp and not dest_is_mtp:
                    # MTP to local
                    source_device, source_internal = parse_mtp_path(source_str)
                    mtp_device = MTPDevice(source_device)
                    # Pass only the internal path, not the full mtp::: path
                    mtp_device.copy_from_device(
                        source_internal or "", str(dest_path))

                elif not source_is_mtp and dest_is_mtp:
                    # Local to MTP
                    dest_device, dest_internal = parse_mtp_path(dest_str)
                    mtp_device = MTPDevice(dest_device)
                    # Pass only the internal path for destination
                    mtp_device.copy_to_device(str(source), dest_internal or "")

                else:
                    # Local to local (improved logic with chunked copy for large files)
                    # Handle conflicts
                    if dest_path.exists():
                        counter = 1
                        stem = source.stem
                        suffix = source.suffix
                        while dest_path.exists():
                            dest_path = destination / \
                                f"{stem}_{counter}{suffix}"
                            counter += 1

                    if source.is_file():
                        # Use chunked copy for files larger than 10MB
                        file_size = source.stat().st_size
                        if file_size > 10 * 1024 * 1024:  # 10MB threshold
                            print(
                                f"Copying large file ({file_size / (1024*1024):.1f} MB): {source.name}")
                            FileOperations._copy_file_chunked(
                                source, dest_path)
                        else:
                            shutil.copy2(source, dest_path)
                    elif source.is_dir():
                        # Use smart copy function for directories with large files
                        shutil.copytree(
                            source, dest_path, copy_function=FileOperations._copy_function_smart)

                processed += 1
                if progress_callback:
                    progress_callback(processed, len(sources))

            except Exception as e:
                file_name = source_str.split(
                    '/')[-1] if isinstance(source, str) or is_mtp_path(str(source)) else source.name
                errors.append(f"{file_name}: {str(e)}")

        return OperationResult(
            success=len(errors) == 0,
            items_processed=processed,
            errors=errors
        )

    @staticmethod
    def move_files(sources: List[Path], destination: Path,
                   progress_callback: Optional[Callable] = None) -> OperationResult:
        """Move files to destination (supports MTP)"""
        errors = []
        processed = 0

        dest_str = str(destination)
        dest_is_mtp = is_mtp_path(dest_str)

        for source in sources:
            try:
                source_str = str(source)
                source_is_mtp = is_mtp_path(source_str)

                # For MTP, move = copy + delete
                if source_is_mtp or dest_is_mtp:
                    # Use copy logic
                    result = FileOperations.copy_files(
                        [source], destination, None)
                    if result.success:
                        # Delete source after successful copy
                        if source_is_mtp:
                            source_device, source_internal = parse_mtp_path(
                                source_str)
                            if source_device:
                                mtp_device = MTPDevice(source_device)
                                # Pass only the internal path, not the full mtp::: path
                                mtp_device.delete_from_device(
                                    source_internal or "")
                        else:
                            if source.is_file():
                                source.unlink()
                            elif source.is_dir():
                                shutil.rmtree(source)
                    else:
                        errors.extend(result.errors)
                        continue
                else:
                    # Local to local move
                    dest_path = destination / source.name

                    # Handle conflicts
                    if dest_path.exists():
                        counter = 1
                        stem = source.stem
                        suffix = source.suffix
                        while dest_path.exists():
                            dest_path = destination / \
                                f"{stem}_{counter}{suffix}"
                            counter += 1

                    shutil.move(str(source), str(dest_path))

                processed += 1
                if progress_callback:
                    progress_callback(processed, len(sources))

            except Exception as e:
                file_name = source_str.split(
                    '/')[-1] if isinstance(source, str) or is_mtp_path(str(source)) else source.name
                errors.append(f"{file_name}: {str(e)}")

        return OperationResult(
            success=len(errors) == 0,
            items_processed=processed,
            errors=errors
        )

    @staticmethod
    def delete_files(paths: List[Path],
                     progress_callback: Optional[Callable] = None) -> OperationResult:
        """Delete files (supports MTP)"""
        errors = []
        processed = 0

        for path in paths:
            try:
                path_str = str(path)

                if is_mtp_path(path_str):
                    # MTP device file - extract internal path
                    device_name, internal_path = parse_mtp_path(path_str)
                    mtp_device = MTPDevice(device_name)
                    if mtp_device.is_available():
                        # Pass only the internal path, not the full mtp::: path
                        mtp_device.delete_from_device(internal_path or "")
                    else:
                        raise Exception(
                            f"Device '{device_name}' not available")
                else:
                    # Local filesystem
                    if path.is_file():
                        path.unlink()
                    elif path.is_dir():
                        shutil.rmtree(path)

                processed += 1
                if progress_callback:
                    progress_callback(processed, len(paths))

            except Exception as e:
                file_name = path_str.split(
                    '/')[-1] if isinstance(path, str) or is_mtp_path(str(path)) else path.name
                errors.append(f"{file_name}: {str(e)}")

        return OperationResult(
            success=len(errors) == 0,
            items_processed=processed,
            errors=errors
        )

    @staticmethod
    def create_archive(sources: List[Path], archive_path: Path,
                       progress_callback: Optional[Callable] = None) -> OperationResult:
        """Create ZIP archive from files"""
        errors = []

        try:
            with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                total_files = len(sources)
                for idx, source in enumerate(sources):
                    try:
                        if source.is_file():
                            zipf.write(source, source.name)
                        elif source.is_dir():
                            for file in source.rglob('*'):
                                if file.is_file():
                                    zipf.write(
                                        file, file.relative_to(source.parent))

                        if progress_callback:
                            progress_callback(idx + 1, total_files)
                    except Exception as e:
                        errors.append(f"{source.name}: {str(e)}")

            return OperationResult(
                success=len(errors) == 0,
                items_processed=len(sources),
                errors=errors
            )
        except Exception as e:
            return OperationResult(
                success=False,
                items_processed=0,
                errors=[str(e)]
            )

    @staticmethod
    def extract_archive(archive_path: Path, destination: Path,
                        progress_callback: Optional[Callable] = None) -> OperationResult:
        """Extract ZIP archive"""
        try:
            with zipfile.ZipFile(archive_path, 'r') as zipf:
                members = zipf.namelist()
                total = len(members)

                for idx, member in enumerate(members):
                    zipf.extract(member, destination)
                    if progress_callback:
                        progress_callback(idx + 1, total)

            return OperationResult(
                success=True,
                items_processed=len(members),
                errors=[]
            )
        except Exception as e:
            return OperationResult(
                success=False,
                items_processed=0,
                errors=[str(e)]
            )


# Theme Manager
# ============================================================================

class ThemeManager:
    """Manages application themes"""

    THEMES = {
        "dark": {
            "bg_primary": "#1E1E1E",
            "bg_secondary": "#2D2D2D",
            "bg_tertiary": "#3E3E3E",
            "fg_primary": "#FFFFFF",
            "fg_secondary": "#B0B0B0",
            "accent": "#FF6B35",
            "accent_hover": "#FF8555",
            "border": "#4A4A4A",
            "selection": "#FF6B35"
        },
        "light": {
            "bg_primary": "#FFFFFF",
            "bg_secondary": "#F5F5F5",
            "bg_tertiary": "#EEEEEE",
            "fg_primary": "#000000",
            "fg_secondary": "#666666",
            "accent": "#2196F3",
            "accent_hover": "#42A5F5",
            "border": "#DDDDDD",
            "selection": "#2196F3"
        }
    }

    def __init__(self, theme_name: str = "dark"):
        self.current_theme_name = theme_name
        self.current_theme = self.THEMES[theme_name]

    def apply_theme(self, root: tk.Tk, style: ttk.Style):
        """Apply theme to application"""
        theme = self.current_theme

        # Configure root window
        root.configure(bg=theme["bg_primary"])

        # Configure ttk styles
        style.configure('TFrame', background=theme["bg_secondary"])
        style.configure('TLabel', background=theme["bg_secondary"],
                        foreground=theme["fg_primary"])
        style.configure('TButton', background=theme["bg_tertiary"],
                        foreground=theme["fg_primary"])
        style.map('TButton',
                  background=[('active', theme["accent_hover"])],
                  foreground=[('active', theme["fg_primary"])])

        style.configure('TEntry', fieldbackground=theme["bg_tertiary"],
                        foreground=theme["fg_primary"])

        style.configure('Treeview',
                        background=theme["bg_secondary"],
                        foreground=theme["fg_primary"],
                        fieldbackground=theme["bg_secondary"],
                        borderwidth=1,
                        relief="solid")

        style.map('Treeview',
                  background=[('selected', theme["selection"])],
                  foreground=[('selected', '#FFFFFF')])

        style.configure('Treeview.Heading',
                        background=theme["bg_tertiary"],
                        foreground=theme["fg_primary"])

        style.configure('TLabelframe', background=theme["bg_secondary"],
                        foreground=theme["fg_primary"])
        style.configure('TLabelframe.Label', background=theme["bg_secondary"],
                        foreground=theme["fg_primary"])

        style.configure('Horizontal.TProgressbar',
                        background=theme["accent"])

    def switch_theme(self):
        """Switch between light and dark themes"""
        self.current_theme_name = "light" if self.current_theme_name == "dark" else "dark"
        self.current_theme = self.THEMES[self.current_theme_name]
        return self.current_theme_name


# Centered Input Dialog
# ============================================================================

class CenteredInputDialog(tk.Toplevel):
    """Custom input dialog that opens centered"""

    def __init__(self, parent, title: str, prompt: str, initialvalue: str = ""):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)

        self.result = None

        # Main frame
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Prompt label
        ttk.Label(main_frame, text=prompt).pack(anchor=tk.W, pady=(0, 10))

        # Entry
        self.entry_var = tk.StringVar(value=initialvalue)
        self.entry = ttk.Entry(
            main_frame, textvariable=self.entry_var, width=40)
        self.entry.pack(fill=tk.X, pady=(0, 15))
        self.entry.focus()
        self.entry.select_range(0, tk.END)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="OK", command=self._on_ok,
                   width=10).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel",
                   command=self._on_cancel, width=10).pack(side=tk.RIGHT)

        # Bind Enter and Escape
        self.entry.bind('<Return>', lambda e: self._on_ok())
        self.bind('<Escape>', lambda e: self._on_cancel())

        # Center the dialog
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = parent.winfo_x() + (parent.winfo_width() - width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - height) // 2
        self.geometry(f"+{x}+{y}")

        # Make modal
        self.grab_set()

    def _on_ok(self):
        self.result = self.entry_var.get()
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()

    @staticmethod
    def ask_string(parent, title: str, prompt: str, initialvalue: str = ""):
        """Show centered input dialog and return the entered string"""
        dialog = CenteredInputDialog(parent, title, prompt, initialvalue)
        parent.wait_window(dialog)
        return dialog.result


# Delete Confirmation Dialog
# ============================================================================

class DeleteConfirmDialog(tk.Toplevel):
    """Custom delete confirmation dialog with Yes/Yes to All/No options"""

    def __init__(self, parent, title: str, message: str, file_list: str):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)

        self.result = None  # Will be 'yes', 'yes_all', or None

        # Main frame
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Warning icon and message
        msg_frame = ttk.Frame(main_frame)
        msg_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(msg_frame, text="⚠️", font=('', 32)
                  ).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(msg_frame, text=message, wraplength=350).pack(
            side=tk.LEFT, fill=tk.X, expand=True)

        # File list
        if file_list:
            list_frame = ttk.Frame(main_frame)
            list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

            # Scrollable text widget
            text_widget = tk.Text(list_frame, height=8, width=50, wrap=tk.WORD)
            scrollbar = ttk.Scrollbar(
                list_frame, orient=tk.VERTICAL, command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)

            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            text_widget.insert('1.0', file_list)
            text_widget.configure(state='disabled')

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="No", command=self._on_no,
                   width=12).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Yes to All", command=self._on_yes_all,
                   width=12).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Yes", command=self._on_yes,
                   width=12).pack(side=tk.RIGHT, padx=(5, 0))

        # Bind Escape
        self.bind('<Escape>', lambda e: self._on_no())
        self.bind('<Return>', lambda e: self._on_yes())

        # Center the dialog
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = parent.winfo_x() + (parent.winfo_width() - width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - height) // 2
        self.geometry(f"+{x}+{y}")

        # Make modal
        self.grab_set()

    def _on_yes(self):
        self.result = 'yes'
        self.destroy()

    def _on_yes_all(self):
        self.result = 'yes_all'
        self.destroy()

    def _on_no(self):
        self.result = None
        self.destroy()

    @staticmethod
    def ask_delete(parent, title: str, message: str, file_list: str = ""):
        """Show centered delete confirmation dialog"""
        dialog = DeleteConfirmDialog(parent, title, message, file_list)
        parent.wait_window(dialog)
        return dialog.result


# Search Dialog
# ============================================================================

class SearchDialog(tk.Toplevel):
    """Dialog for searching files"""

    def __init__(self, parent, root_path, show_hidden: bool):
        super().__init__(parent)
        self.title("Search Files")
        self.geometry("500x400")
        self.transient(parent)

        # Center the dialog
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

        self.root_path = root_path
        self.show_hidden = show_hidden
        self.is_mtp = is_mtp_path(str(root_path))
        self.results = []  # Can be List[Path] or List[str] for MTP
        self.selected_result = None

        self._create_widgets()

    def _create_widgets(self):
        """Create search dialog widgets"""
        # Search frame
        search_frame = ttk.Frame(self, padding=10)
        search_frame.pack(fill=tk.X)

        ttk.Label(search_frame, text="Search in:").pack(anchor=tk.W)
        ttk.Label(search_frame, text=str(self.root_path),
                  font=('', 9, 'italic')).pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(search_frame, text="Pattern:").pack(anchor=tk.W)
        self.pattern_var = tk.StringVar()
        pattern_entry = ttk.Entry(search_frame, textvariable=self.pattern_var)
        pattern_entry.pack(fill=tk.X, pady=(0, 10))
        pattern_entry.focus()
        pattern_entry.bind('<Return>', lambda e: self.search())

        ttk.Button(search_frame, text="🔍 Search",
                   command=self.search).pack(fill=tk.X)

        # Results frame
        results_frame = ttk.LabelFrame(self, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Results listbox with scrollbar
        scrollbar = ttk.Scrollbar(results_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.results_listbox = tk.Listbox(results_frame,
                                          yscrollcommand=scrollbar.set)
        self.results_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.results_listbox.yview)

        self.results_listbox.bind(
            '<Double-Button-1>', lambda e: self.open_selected())

        # Button frame
        button_frame = ttk.Frame(self, padding=10)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="Open Location",
                   command=self.open_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close",
                   command=self.destroy).pack(side=tk.RIGHT, padx=5)

        self.status_label = ttk.Label(
            self, text="Enter search pattern and press Search")
        self.status_label.pack(fill=tk.X, padx=10, pady=(0, 10))

    def search(self):
        """Perform search"""
        pattern = self.pattern_var.get().strip()
        if not pattern:
            return

        self.status_label.config(text="Searching...")
        self.results_listbox.delete(0, tk.END)
        self.update()

        # Search in background
        def worker():
            results = search_files(self.root_path, pattern, self.show_hidden)
            self.after(0, lambda: self.display_results(results))

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def display_results(self, results):
        """Display search results (handles both Path objects and MTP strings)"""
        self.results = results
        self.results_listbox.delete(0, tk.END)

        for result in results:
            if isinstance(result, Path):
                # Regular filesystem path
                icon = get_file_icon(result)
                relative = result.relative_to(self.root_path) if result.is_relative_to(
                    self.root_path) else result
                self.results_listbox.insert(tk.END, f"{icon} {relative}")
            else:
                # MTP path string
                result_str = str(result)
                # Extract filename and path
                if '/' in result_str:
                    # Get relative path from root
                    root_str = str(self.root_path)
                    if result_str.startswith(root_str):
                        relative_path = result_str[len(root_str):].lstrip('/')
                    else:
                        relative_path = result_str.split(
                            '/', 3)[-1] if result_str.count('/') >= 3 else result_str

                    # Get file icon based on extension
                    filename = result_str.split('/')[-1]
                    icon = get_file_icon(Path(filename))
                    self.results_listbox.insert(
                        tk.END, f"{icon} {relative_path}")
                else:
                    self.results_listbox.insert(tk.END, f"📄 {result_str}")

        self.status_label.config(text=f"Found {len(results)} items")

    def open_selected(self):
        """Open selected result location"""
        selection = self.results_listbox.curselection()
        if selection:
            self.selected_result = self.results[selection[0]]
            self.destroy()


# Properties Dialog
# ============================================================================

class PropertiesDialog(tk.Toplevel):
    """Show file/folder properties"""

    def __init__(self, parent, path, file_info=None):
        super().__init__(parent)
        self.path = path
        self.file_info = file_info
        self.is_mtp = is_mtp_path(str(path))

        # Get name
        if self.is_mtp:
            name = str(path).split('/')[-1] if '/' in str(path) else str(path)
        else:
            name = path.name

        self.title(f"Properties - {name}")
        self.geometry("400x300")
        self.transient(parent)
        self.resizable(False, False)

        # Center the dialog
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

        self._create_widgets()

    def _create_widgets(self):
        """Create properties widgets"""
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Get name
        if self.is_mtp:
            name = str(self.path).split(
                '/')[-1] if '/' in str(self.path) else str(self.path)
        else:
            name = self.path.name

        # Icon and name
        icon = get_file_icon(self.path)
        ttk.Label(main_frame, text=icon, font=('', 48)).pack(pady=10)
        ttk.Label(main_frame, text=name,
                  font=('', 12, 'bold')).pack(pady=5)

        # Properties
        props_frame = ttk.Frame(main_frame)
        props_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        if self.is_mtp:
            # MTP device properties
            if self.file_info:
                self._add_property(props_frame, "Type:",
                                   "Folder" if self.file_info.is_directory else "File")
                self._add_property(props_frame, "Location:", str(
                    self.path).rsplit('/', 1)[0] if '/' in str(self.path) else "")

                if not self.file_info.is_directory:
                    self._add_property(props_frame, "Size:",
                                       format_size(self.file_info.size))
                else:
                    self._add_property(props_frame, "Size:", "---")

                if self.file_info.modified:
                    self._add_property(props_frame, "Modified:",
                                       self.file_info.modified.strftime("%Y-%m-%d %H:%M"))
                else:
                    self._add_property(props_frame, "Modified:", "---")
            else:
                self._add_property(props_frame, "Type:", "MTP Item")
                self._add_property(props_frame, "Location:", str(self.path))
        else:
            # Local file properties
            stat = self.path.stat()

            self._add_property(props_frame, "Type:",
                               "Folder" if self.path.is_dir() else "File")
            self._add_property(props_frame, "Location:", str(self.path.parent))

            if self.path.is_file():
                self._add_property(props_frame, "Size:",
                                   format_size(stat.st_size))
            else:
                size, count = get_folder_size(self.path)
                self._add_property(props_frame, "Size:",
                                   f"{format_size(size)} ({count} items)")

            self._add_property(props_frame, "Modified:",
                               format_date(stat.st_mtime))
            self._add_property(props_frame, "Created:",
                               format_date(stat.st_ctime))

        # Close button
        ttk.Button(main_frame, text="Close",
                   command=self.destroy).pack(pady=10)

    def _add_property(self, parent, label: str, value: str):
        """Add property row"""
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=2)

        ttk.Label(row, text=label, font=('', 9, 'bold'),
                  width=12).pack(side=tk.LEFT)
        ttk.Label(row, text=value, font=('', 9)).pack(side=tk.LEFT, fill=tk.X)


# File Panel Component
# ============================================================================

class FilePanel(ttk.Frame):
    """Individual file panel for navigation and display"""

    def __init__(self, parent, panel_id: str, config: AppConfig, theme_manager: ThemeManager, main_window=None):
        super().__init__(parent)
        self.panel_id = panel_id
        self.config = config
        self.theme_manager = theme_manager
        self.main_window = main_window  # Reference to main window for context menu actions
        self.current_path = Path.home()
        self.selected_items: List[Path] = []
        self.size_cache = {}
        self.history: List[Path] = []
        self.history_index = -1
        self.is_active = False

        self._create_widgets()

    def _create_widgets(self):
        """Create panel UI components"""
        # Drive selector
        drive_frame = ttk.Frame(self)
        drive_frame.pack(fill=tk.X, padx=5, pady=(5, 2))

        ttk.Label(drive_frame, text="Drive:").pack(side=tk.LEFT, padx=(0, 5))

        self.drive_var = tk.StringVar()
        self.drive_combo = ttk.Combobox(drive_frame, textvariable=self.drive_var,
                                        state="readonly", width=25)
        self.drive_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.drive_combo.bind('<<ComboboxSelected>>', self._on_drive_selected)
        self._populate_drives()

        # Refresh drives button
        ttk.Button(drive_frame, text="🔄", width=3,
                   command=self._refresh_drives).pack(side=tk.LEFT, padx=(5, 0))

        # Drive info label
        self.drive_info_label = ttk.Label(self, text="", font=('', 8))
        self.drive_info_label.pack(fill=tk.X, padx=5, pady=(0, 2))

        # Path and navigation bar
        path_frame = ttk.Frame(self)
        path_frame.pack(fill=tk.X, padx=5, pady=(2, 5))

        # Navigation buttons
        nav_frame = ttk.Frame(path_frame)
        nav_frame.pack(side=tk.LEFT, padx=(0, 5))

        self.back_btn = ttk.Button(nav_frame, text="◀", width=3,
                                   command=self.go_back, state='disabled')
        self.back_btn.pack(side=tk.LEFT, padx=1)

        self.forward_btn = ttk.Button(nav_frame, text="▶", width=3,
                                      command=self.go_forward, state='disabled')
        self.forward_btn.pack(side=tk.LEFT, padx=1)

        ttk.Button(nav_frame, text="⬆️", width=3,
                   command=self.go_up).pack(side=tk.LEFT, padx=1)

        # Path entry
        ttk.Label(path_frame, text="📂").pack(side=tk.LEFT, padx=(5, 2))

        self.path_var = tk.StringVar(value=str(self.current_path))
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.path_entry.bind('<Return>', self._on_path_enter)

        # Action buttons
        action_frame = ttk.Frame(path_frame)
        action_frame.pack(side=tk.LEFT)

        ttk.Button(action_frame, text="🔍", width=3,
                   command=self.search_files).pack(side=tk.LEFT, padx=1)
        ttk.Button(action_frame, text="🔄", width=3,
                   command=self.refresh).pack(side=tk.LEFT, padx=1)

        # Tree view with scrollbar
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Treeview
        self.tree = ttk.Treeview(
            tree_frame,
            columns=('size', 'modified', 'type'),
            show='tree headings',
            selectmode='extended',
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)

        # Column configuration
        self.tree.heading('#0', text='Name',
                          command=lambda: self.sort_by('name'))
        self.tree.heading('size', text='Size',
                          command=lambda: self.sort_by('size'))
        self.tree.heading('modified', text='Modified',
                          command=lambda: self.sort_by('modified'))
        self.tree.heading('type', text='Type',
                          command=lambda: self.sort_by('type'))

        self.tree.column('#0', width=300, minwidth=150)
        self.tree.column('size', width=100, minwidth=80)
        self.tree.column('modified', width=150, minwidth=120)
        self.tree.column('type', width=80, minwidth=60)

        # Bindings
        self.tree.bind('<Double-Button-1>', self._on_double_click)
        self.tree.bind('<<TreeviewSelect>>', self._on_select)
        self.tree.bind('<Button-3>', self._on_right_click)
        self.tree.bind('<Control-a>', lambda e: self.select_all())

        # Drag and drop setup
        self.tree.bind('<ButtonPress-1>', self._on_drag_start)
        self.tree.bind('<B1-Motion>', self._on_drag_motion)
        self.tree.bind('<ButtonRelease-1>', self._on_drag_end)
        self.dragging = False
        self.drag_data = None

        # Status bar
        self.status_label = ttk.Label(
            self, text="", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(fill=tk.X, padx=5, pady=(0, 5))

        # Sort state
        self.sort_column = 'name'
        self.sort_reverse = False

    def navigate_to(self, path, add_to_history: bool = True):
        """Navigate to specified directory (handles both Path and MTP paths)"""
        # Convert to string for MTP path handling
        path_str = str(path)

        # Check if it's an MTP path
        if is_mtp_path(path_str):
            # MTP device - no validation needed, will be checked in refresh()
            pass
        else:
            # Regular filesystem path
            path = Path(path_str) if isinstance(path, str) else path
            if not path.exists() or not path.is_dir():
                messagebox.showerror("Error", f"Invalid path: {path}")
                return

        if add_to_history:
            # Add current path to history
            if self.history_index < len(self.history) - 1:
                self.history = self.history[:self.history_index + 1]
            self.history.append(self.current_path)
            self.history_index = len(self.history) - 1
            self._update_nav_buttons()

        self.current_path = path_str if is_mtp_path(path_str) else path
        self.path_var.set(path_str)
        self.refresh()
        self._update_drive_info()

        # Update drive selector
        if is_mtp_path(path_str):
            device_name, _ = parse_mtp_path(path_str)
            if device_name:
                for label in self.drive_paths:
                    if device_name in label:
                        self.drive_var.set(label)
                        break
        else:
            for label, drive_path in self.drive_paths.items():
                if not is_mtp_path(drive_path) and str(path).startswith(drive_path):
                    self.drive_var.set(label)
                    break

    def go_back(self):
        """Navigate back in history"""
        if self.history_index > 0:
            self.history_index -= 1
            self.current_path = self.history[self.history_index]
            self.path_var.set(str(self.current_path))
            self.refresh()
            self._update_nav_buttons()

    def go_forward(self):
        """Navigate forward in history"""
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            self.current_path = self.history[self.history_index]
            self.path_var.set(str(self.current_path))
            self.refresh()
            self._update_nav_buttons()

    def _update_nav_buttons(self):
        """Update navigation button states"""
        self.back_btn.config(
            state='normal' if self.history_index > 0 else 'disabled')
        self.forward_btn.config(state='normal' if self.history_index < len(
            self.history) - 1 else 'disabled')

    def refresh(self):
        """Refresh current directory view (handles both filesystem and MTP)"""
        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.size_cache.clear()

        try:
            items = []
            show_hidden = self.config.get('show_hidden', False)

            # Check if it's an MTP path
            current_path_str = str(self.current_path)
            if is_mtp_path(current_path_str):
                # Handle MTP device
                device_name, internal_path = parse_mtp_path(current_path_str)
                if device_name:
                    mtp_device = MTPDevice(device_name)
                    if mtp_device.is_available():
                        mtp_items = mtp_device.list_contents(
                            internal_path or "")

                        for mtp_item in mtp_items:
                            # Handle modified time safely
                            modified_dt = None
                            if mtp_item['modified']:
                                try:
                                    modified_dt = datetime.fromtimestamp(
                                        mtp_item['modified'])
                                except (OSError, ValueError, OverflowError):
                                    # Invalid timestamp, use None
                                    modified_dt = None

                            items.append(FileInfo(
                                # Already has mtp::: prefix
                                path=mtp_item['path'],
                                name=mtp_item['name'],
                                is_directory=mtp_item['is_dir'],
                                size=mtp_item['size'],
                                modified=modified_dt,
                                extension=Path(mtp_item['name']).suffix.lower(
                                ) if not mtp_item['is_dir'] else ""
                            ))
                    else:
                        messagebox.showerror(
                            "MTP Error", f"Device '{device_name}' not available")
                        return
            else:
                # Regular filesystem path
                for entry in os.scandir(self.current_path):
                    try:
                        # Skip hidden files if configured
                        if not show_hidden and entry.name.startswith('.'):
                            continue

                        stat = entry.stat()
                        items.append(FileInfo(
                            path=Path(entry.path),
                            name=entry.name,
                            is_directory=entry.is_dir(),
                            size=stat.st_size if entry.is_file() else 0,
                            modified=datetime.fromtimestamp(stat.st_mtime),
                            extension=Path(entry.name).suffix.lower()
                        ))
                    except (PermissionError, OSError):
                        continue

            # Sort items
            self._sort_items(items)

            # Store items for later use (e.g., properties dialog)
            self._current_items = items

            # Insert into tree
            for item in items:
                icon = get_file_icon(item.path if not isinstance(
                    item.path, str) else Path(item.name))
                name = f"{icon} {item.name}"

                if item.is_directory:
                    size_text = "---"
                    type_text = "Folder"
                    # Calculate size in background
                    path_str = str(item.path)
                    if is_mtp_path(path_str):
                        # Only calculate MTP folder sizes if enabled in settings
                        if self.config.get('calculate_mtp_folder_sizes', True):
                            self._calculate_folder_size_async(item.path)
                    else:
                        # Only calculate Windows folder sizes if enabled in settings
                        if self.config.get('calculate_windows_folder_sizes', True):
                            self._calculate_folder_size_async(item.path)
                else:
                    size_text = format_size(item.size)
                    type_text = item.extension[1:].upper(
                    ) if item.extension else "File"

                modified_text = item.modified.strftime(
                    "%Y-%m-%d %H:%M") if item.modified else ""

                self.tree.insert('', 'end', str(item.path),
                                 text=name,
                                 values=(size_text, modified_text, type_text),
                                 tags=('directory' if item.is_directory else 'file',))

            # Update status
            dir_count = sum(1 for i in items if i.is_directory)
            file_count = len(items) - dir_count
            self.status_label.config(
                text=f"{len(items)} items ({dir_count} folders, {file_count} files) | {str(self.current_path)}"
            )

        except PermissionError:
            messagebox.showerror(
                "Error", f"Access denied: {self.current_path}")

    def _sort_items(self, items: List[FileInfo]):
        """Sort items list"""
        if self.sort_column == 'name':
            items.sort(key=lambda x: (not x.is_directory,
                       x.name.lower()), reverse=self.sort_reverse)
        elif self.sort_column == 'size':
            items.sort(key=lambda x: (not x.is_directory, x.size),
                       reverse=self.sort_reverse)
        elif self.sort_column == 'modified':
            items.sort(key=lambda x: (not x.is_directory,
                       x.modified or datetime.min), reverse=self.sort_reverse)
        elif self.sort_column == 'type':
            items.sort(key=lambda x: (not x.is_directory,
                       x.extension), reverse=self.sort_reverse)

    def sort_by(self, column: str):
        """Sort tree by column"""
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        self.refresh()

    def _calculate_folder_size_async(self, path):
        """Calculate folder size in background thread (supports both filesystem and MTP)"""
        def worker():
            try:
                path_str = str(path)
                if is_mtp_path(path_str):
                    # MTP folder
                    device_name, internal_path = parse_mtp_path(path_str)
                    if device_name:
                        mtp_device = MTPDevice(device_name)
                        if mtp_device.is_available():
                            total_size, item_count = self._calculate_mtp_folder_size(
                                mtp_device, internal_path or "")
                            self.size_cache[path_str] = (
                                total_size, item_count)
                            try:
                                self.after(0, lambda: self._update_folder_size(
                                    path_str, total_size, item_count))
                            except Exception:
                                pass  # Ignore thread safety errors when updating UI
                else:
                    # Regular filesystem folder
                    total_size, item_count = get_folder_size(path)
                    self.size_cache[path] = (total_size, item_count)
                    try:
                        self.after(0, lambda: self._update_folder_size(
                            str(path), total_size, item_count))
                    except Exception:
                        pass  # Ignore thread safety errors when updating UI
            except Exception as e:
                # Only print errors that aren't related to UI updates
                if "main loop" not in str(e):
                    print(f"Error calculating folder size: {e}")

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def _calculate_mtp_folder_size(self, mtp_device: MTPDevice, folder_path: str) -> tuple:
        """Recursively calculate total size of MTP folder"""
        total_size = 0
        item_count = 0

        try:
            items = mtp_device.list_contents(folder_path)
            for item in items:
                if item['is_dir']:
                    # Recursively calculate subfolder size
                    _, internal_path = parse_mtp_path(item['path'])
                    if internal_path:
                        sub_size, sub_count = self._calculate_mtp_folder_size(
                            mtp_device, internal_path)
                        total_size += sub_size
                        item_count += sub_count
                else:
                    # Add file size
                    total_size += item['size']
                    item_count += 1
        except Exception as e:
            print(f"Error calculating MTP folder size for {folder_path}: {e}")

        return total_size, item_count

    def _update_folder_size(self, path_str: str, size: int, count: int):
        """Update folder size in tree view"""
        try:
            if self.tree.exists(path_str):
                current_values = list(self.tree.item(path_str)['values'])
                current_values[0] = f"{format_size(size)} ({count})"
                self.tree.item(path_str, values=current_values)
        except:
            pass

    def _populate_drives(self):
        """Populate drive selector with available drives and MTP devices"""
        drives = []
        drive_paths = {}

        if os.name == 'nt':  # Windows
            import string
            for letter in string.ascii_uppercase:
                drive = f"{letter}:\\"
                if os.path.exists(drive):
                    try:
                        # Get volume name
                        import ctypes
                        kernel32 = ctypes.windll.kernel32
                        volumeNameBuffer = ctypes.create_unicode_buffer(1024)
                        kernel32.GetVolumeInformationW(
                            ctypes.c_wchar_p(drive),
                            volumeNameBuffer,
                            ctypes.sizeof(volumeNameBuffer),
                            None, None, None, None, 0
                        )
                        volume_name = volumeNameBuffer.value
                        if volume_name:
                            label = f"{letter}: ({volume_name})"
                        else:
                            label = f"{letter}:"
                    except:
                        label = f"{letter}:"
                    drives.append(label)
                    drive_paths[label] = drive

            # Add MTP devices
            mtp_devices = get_mtp_devices()
            for device in mtp_devices:
                label = f"📱 {device['name']}"
                drives.append(label)
                drive_paths[label] = device['path']

        else:  # Unix-like
            drives.append("/ (Root)")
            drive_paths["/ (Root)"] = "/"
            # Add common mount points
            for mount in ["/mnt", "/media", "/Volumes"]:
                if os.path.exists(mount):
                    try:
                        for entry in os.listdir(mount):
                            full_path = os.path.join(mount, entry)
                            if os.path.isdir(full_path):
                                drives.append(f"{entry} ({mount})")
                                drive_paths[f"{entry} ({mount})"] = full_path
                    except:
                        pass

        self.drive_combo['values'] = drives
        self.drive_paths = drive_paths

        # Set current drive
        current_drive = str(self.current_path)
        if is_mtp_path(current_drive):
            device_name, _ = parse_mtp_path(current_drive)
            if device_name:
                for label in drive_paths:
                    if device_name in label:
                        self.drive_var.set(label)
                        break
        else:
            for label, path in drive_paths.items():
                if not is_mtp_path(path) and current_drive.startswith(path):
                    self.drive_var.set(label)
                    break

    def _refresh_drives(self):
        """Refresh the drive list (useful when drives are connected/disconnected)"""
        # Save current drive selection
        current_selection = self.drive_var.get()

        # Re-populate drives
        self._populate_drives()

        # Try to restore selection if it still exists
        if current_selection in self.drive_combo['values']:
            self.drive_var.set(current_selection)

        # Show feedback
        if hasattr(self, 'main_window') and self.main_window:
            self.main_window.status_bar.config(text="✓ Drive list refreshed")
        elif hasattr(self, 'status_label'):
            self.status_label.config(text="✓ Drive list refreshed")

    def _on_drive_selected(self, event):
        """Handle drive selection"""
        selected = self.drive_var.get()
        if selected in self.drive_paths:
            path_value = self.drive_paths[selected]
            # Don't wrap MTP paths in Path() - keep as string
            if is_mtp_path(path_value):
                self.navigate_to(path_value)
            else:
                self.navigate_to(Path(path_value))

    def _update_drive_info(self):
        """Update drive information display (supports both filesystem and MTP)"""
        try:
            current_path_str = str(self.current_path)

            # Check if it's an MTP path
            if is_mtp_path(current_path_str):
                # MTP device
                device_name, internal_path = parse_mtp_path(current_path_str)
                if device_name:
                    mtp_device = MTPDevice(device_name)
                    if mtp_device.is_available():
                        storage_info = mtp_device.get_storage_info(
                            internal_path or "")
                        if storage_info:
                            total = storage_info['total']
                            free = storage_info['free']
                            used = total - free
                            percent = (used / total * 100) if total > 0 else 0
                            self.drive_info_label.config(
                                text=f"Drive: {format_size(free)} free of {format_size(total)} ({percent:.1f}% used)"
                            )
                        else:
                            # Storage info not available for this MTP device
                            self.drive_info_label.config(
                                text="Drive: Storage info unavailable")
                    else:
                        self.drive_info_label.config(text="")
                else:
                    self.drive_info_label.config(text="")
            elif os.name == 'nt':
                # Windows filesystem
                import ctypes
                drive = os.path.splitdrive(str(self.current_path))[0] + "\\"
                free_bytes = ctypes.c_ulonglong(0)
                total_bytes = ctypes.c_ulonglong(0)
                ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                    ctypes.c_wchar_p(drive),
                    None,
                    ctypes.pointer(total_bytes),
                    ctypes.pointer(free_bytes)
                )
                total = total_bytes.value
                free = free_bytes.value
                used = total - free
                percent = (used / total * 100) if total > 0 else 0
                self.drive_info_label.config(
                    text=f"Drive: {format_size(free)} free of {format_size(total)} ({percent:.1f}% used)"
                )
            else:
                # Non-Windows filesystem
                import shutil
                stat = shutil.disk_usage(str(self.current_path))
                percent = (stat.used / stat.total *
                           100) if stat.total > 0 else 0
                self.drive_info_label.config(
                    text=f"Drive: {format_size(stat.free)} free of {format_size(stat.total)} ({percent:.1f}% used)"
                )
        except Exception as e:
            print(f"Error updating drive info: {e}")
            self.drive_info_label.config(text="")

    def go_up(self):
        """Navigate to parent directory (handles both filesystem and MTP)"""
        current_path_str = str(self.current_path)

        if is_mtp_path(current_path_str):
            # MTP path navigation
            device_name, internal_path = parse_mtp_path(current_path_str)
            if internal_path and '/' in internal_path:
                # Go up one level in MTP device
                parent_path = '/'.join(internal_path.split('/')[:-1])
                new_mtp_path = f"mtp:::{device_name}/{parent_path}" if parent_path else f"mtp:::{device_name}"
                self.navigate_to(new_mtp_path)
            # If at root of device, do nothing (can't go up from device root)
        else:
            # Regular filesystem path
            parent = self.current_path.parent
            if parent != self.current_path:
                self.navigate_to(parent)

    def get_selected_items(self) -> List:
        """Get list of selected file paths (can be Path or MTP string)"""
        selection = self.tree.selection()
        # Return as-is: Path objects for local files, strings for MTP paths
        result = []
        for item in selection:
            if is_mtp_path(item):
                result.append(item)  # Keep as string
            else:
                result.append(Path(item))  # Convert to Path
        return result

    def select_all(self):
        """Select all items"""
        self.tree.selection_set(self.tree.get_children())

    def search_files(self):
        """Open search dialog"""
        dialog = SearchDialog(self, self.current_path,
                              self.config.get('show_hidden', False))
        self.wait_window(dialog)

        if dialog.selected_result:
            # Navigate to parent and select the result
            self.navigate_to(dialog.selected_result.parent)

    def _on_double_click(self, event):
        """Handle double-click on item (handles both filesystem and MTP)"""
        selection = self.tree.selection()
        if not selection:
            return

        path_str = selection[0]

        if is_mtp_path(path_str):
            # MTP device path
            device_name, internal_path = parse_mtp_path(path_str)
            if device_name:
                mtp_device = MTPDevice(device_name)
                if mtp_device.is_available():
                    # Check if it's a directory by listing its parent and finding this item
                    items = mtp_device.list_contents(internal_path or "")
                    # If list_contents succeeds, it's likely a folder (folders can be listed), assume if path ends without extension, it's a folder

                    if not Path(path_str.split('/')[-1]).suffix:
                        # Likely a folder
                        self.navigate_to(path_str)
                    else:
                        # It's a file - copy to temp and open
                        try:
                            temp_dir = Path(os.environ.get('TEMP', '/tmp'))
                            temp_file = temp_dir / \
                                Path(path_str.split('/')[-1]).name

                            # Copy from device to temp
                            mtp_device.copy_from_device(
                                path_str, str(temp_file))

                            # Open the temp file
                            if os.name == 'nt':
                                os.startfile(str(temp_file))
                            elif os.name == 'posix':
                                import subprocess
                                subprocess.call(('xdg-open', str(temp_file)))
                            else:
                                os.system(f'open "{temp_file}"')
                        except Exception as e:
                            messagebox.showerror(
                                "Error", f"Cannot open MTP file: {e}")
        else:
            # Regular filesystem path
            path = Path(path_str)
            if path.is_dir():
                self.navigate_to(path)
            else:
                # Open file with default application
                try:
                    if os.name == 'nt':
                        os.startfile(str(path))
                    elif os.name == 'posix':
                        import subprocess
                        subprocess.call(('xdg-open', str(path)))
                    else:
                        os.system(f'open "{path}"')
                except Exception as e:
                    messagebox.showerror("Error", f"Cannot open file: {e}")

    def _on_select(self, event):
        """Handle selection change"""
        self.selected_items = self.get_selected_items()

    def _on_path_enter(self, event):
        """Handle Enter key in path entry"""
        path = Path(self.path_var.get())
        self.navigate_to(path)

    def _on_right_click(self, event):
        """Handle right-click context menu"""
        # Set this panel as active when right-clicking
        if self.main_window:
            self.main_window._set_active_panel(self)

        # Select item under cursor
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self._show_context_menu(event)
        else:
            # Right-click on empty space
            self._show_empty_context_menu(event)

    def _show_context_menu(self, event):
        """Show context menu for selected items"""
        menu = tk.Menu(self, tearoff=0)

        selected = self.get_selected_items()
        if len(selected) == 1:
            path = selected[0]

            menu.add_command(
                label="Open", command=lambda: self._on_double_click(event))
            menu.add_separator()

            # Check if it's a file with archive extension (handle both Path and string)
            path_str = str(path)
            if isinstance(path, str):
                path_obj = Path(path_str.split(
                    '/')[-1]) if is_mtp_path(path_str) else Path(path_str)
            else:
                path_obj = path

            if path_obj.suffix.lower() in ['.zip', '.tar', '.gz', '.7z', '.rar']:
                menu.add_command(label="Extract Here",
                                 command=lambda: self.extract_archive(path))
                menu.add_separator()

            menu.add_command(
                label="Rename", command=lambda: self.rename_item(path))
            menu.add_command(label="Properties",
                             command=lambda: self.show_properties(path))
            menu.add_separator()

        if selected:
            if len(selected) > 1:
                menu.add_command(
                    label=f"Selected {len(selected)} items", state='disabled')
                menu.add_separator()

            # Add copy, move, delete options (icons on the right)
            menu.add_command(label="Copy\t📋",
                             command=lambda: self._copy_to_other_panel())
            menu.add_command(label="Move\t✂️",
                             command=lambda: self._move_to_other_panel())
            menu.add_separator()
            menu.add_command(label="Create Archive\t🗜️",
                             command=lambda: self.create_archive_dialog(selected))
            menu.add_separator()
            menu.add_command(
                label="Delete\t🗑️", command=lambda: self._delete_selected())

        menu.post(event.x_root, event.y_root)

    def _show_empty_context_menu(self, event):
        """Show context menu for empty space (no item selected)"""
        menu = tk.Menu(self, tearoff=0)

        menu.add_command(label="New Folder\t📁",
                         command=lambda: self._create_new_folder())
        menu.add_command(label="New File\t📄",
                         command=lambda: self._create_new_file())
        menu.add_separator()
        menu.add_command(label="Refresh\t🔄",
                         command=self.refresh)
        menu.add_separator()
        menu.add_command(label="Paste\t📋",
                         command=lambda: self._paste_here(),
                         state='disabled')  # To be enabled when clipboard has data

        menu.post(event.x_root, event.y_root)

    def _copy_to_other_panel(self):
        """Copy selected items to the other panel"""
        # This will be handled by the main window
        if self.main_window and hasattr(self.main_window, 'copy_files'):
            self.main_window.copy_files()

    def _move_to_other_panel(self):
        """Move selected items to the other panel"""
        # This will be handled by the main window
        if self.main_window and hasattr(self.main_window, 'move_files'):
            self.main_window.move_files()

    def _create_new_folder(self):
        """Create a new folder in current directory"""
        # This will be handled by the main window
        if self.main_window and hasattr(self.main_window, 'new_folder'):
            self.main_window.new_folder()

    def _create_new_file(self):
        """Create a new file in current directory"""
        # This will be handled by the main window
        if self.main_window and hasattr(self.main_window, 'new_file'):
            self.main_window.new_file()

    def _paste_here(self):
        """Paste clipboard contents here"""
        # Placeholder for future clipboard implementation
        pass

    def rename_item(self, path: Path):
        """Rename file or folder"""
        new_name = CenteredInputDialog.ask_string(
            self, "Rename", "Enter new name:", initialvalue=path.name)
        if new_name and new_name != path.name:
            try:
                new_path = path.parent / new_name
                path.rename(new_path)
                self.refresh()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to rename: {e}")

    def show_properties(self, path: Path):
        """Show properties dialog"""
        # Try to get file info from the items list
        file_info = None
        if hasattr(self, '_current_items'):
            for item in self._current_items:
                if str(item.path) == str(path):
                    file_info = item
                    break

        PropertiesDialog(self, path, file_info)

    def create_archive_dialog(self, sources: List[Path]):
        """Create archive from selected files"""
        archive_name = CenteredInputDialog.ask_string(
            self, "Create Archive", "Enter archive name:", initialvalue="archive.zip")
        if archive_name:
            if not archive_name.endswith('.zip'):
                archive_name += '.zip'

            archive_path = self.current_path / archive_name

            # Show progress
            progress = ProgressDialog(self, "Creating Archive", len(sources))
            result = None

            def worker():
                nonlocal result
                result = FileOperations.create_archive(
                    sources, archive_path,
                    lambda c, t: self.after(
                        0, lambda: progress.update_progress(c, t))
                )
                self.after(0, progress.destroy)

            thread = threading.Thread(target=worker, daemon=True)
            thread.start()
            self.wait_window(progress)

            if result and result.success:
                messagebox.showinfo(
                    "Success", f"Archive created: {archive_name}")
                self.refresh()
            elif result:
                messagebox.showerror(
                    "Error", f"Failed to create archive:\n" + "\n".join(result.errors[:5]))

    def extract_archive(self, archive_path: Path):
        """Extract archive"""
        dest_name = archive_path.stem
        dest_path = self.current_path / dest_name

        if dest_path.exists():
            if not messagebox.askyesno("Folder Exists",
                                       f"Folder '{dest_name}' already exists. Extract anyway?"):
                return
        else:
            dest_path.mkdir()

        progress = ProgressDialog(self, "Extracting Archive", 100)
        result = None

        def worker():
            nonlocal result
            result = FileOperations.extract_archive(
                archive_path, dest_path,
                lambda c, t: self.after(
                    0, lambda: progress.update_progress(c, t))
            )
            self.after(0, progress.destroy)

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
        self.wait_window(progress)

        if result and result.success:
            messagebox.showinfo(
                "Success", f"Archive extracted to: {dest_name}")
            self.refresh()
        elif result:
            messagebox.showerror(
                "Error", f"Failed to extract archive:\n" + "\n".join(result.errors[:5]))

    def _delete_selected(self):
        """Delete selected items (internal method)"""
        # This will be called from the main window's delete method
        if self.main_window and hasattr(self.main_window, 'delete_files'):
            self.main_window.delete_files()

    # Drag and drop support
    def _on_drag_start(self, event):
        """Start drag operation"""
        item = self.tree.identify_row(event.y)
        if item:
            self.drag_data = {'source_panel': self,
                              'items': self.get_selected_items()}

    def _on_drag_motion(self, event):
        """Handle drag motion"""
        if self.drag_data:
            self.dragging = True

    def _on_drag_end(self, event):
        """End drag operation"""
        if self.dragging and self.drag_data:
            # Find target widget
            widget = self.winfo_containing(event.x_root, event.y_root)
            if widget and hasattr(widget, 'master'):
                # Check if dropped on another FilePanel
                target = widget
                while target:
                    if isinstance(target, FilePanel) and target != self:
                        # Perform copy operation
                        self._perform_drag_drop(target)
                        break
                    target = target.master if hasattr(
                        target, 'master') else None

        self.dragging = False
        self.drag_data = None

    def _perform_drag_drop(self, target_panel):
        """Perform drag and drop copy"""
        if self.drag_data and 'items' in self.drag_data:
            # This will be handled by the main window
            pass


# Progress Dialog
# ============================================================================

class ProgressDialog(tk.Toplevel):
    """Dialog showing operation progress"""

    def __init__(self, parent, title: str, total_items: int):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x240")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # Center the dialog
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

        self.total_items = total_items
        self.cancelled = False

        # Main frame
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Status label
        self.status_label = ttk.Label(
            main_frame, text="Preparing...", font=('', 10))
        self.status_label.pack(pady=(0, 10))

        # Progress bar
        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=total_items,
            mode='determinate',
            length=450
        )
        self.progress_bar.pack(fill=tk.X, pady=10)

        # Progress label
        self.progress_label = ttk.Label(main_frame, text="0 / 0", font=('', 9))
        self.progress_label.pack(pady=(0, 10))

        # Speed and ETA frame
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 15))

        self.speed_label = ttk.Label(info_frame, text="", font=('', 8))
        self.speed_label.pack(side=tk.LEFT)

        self.eta_label = ttk.Label(info_frame, text="", font=('', 8))
        self.eta_label.pack(side=tk.RIGHT)

        # Cancel button
        cancel_btn = ttk.Button(main_frame, text="Cancel", command=self.cancel)
        cancel_btn.pack(pady=(15, 5), ipady=8)
        cancel_btn.configure(width=20)

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.start_time = time.time()

    def update_progress(self, current: int, total: int):
        """Update progress display"""
        self.progress_var.set(current)
        self.progress_label.config(text=f"{current} / {total}")
        self.status_label.config(
            text=f"Processing item {current} of {total}...")

        # Calculate speed and ETA
        elapsed = time.time() - self.start_time
        if elapsed > 0 and current > 0:
            speed = current / elapsed
            remaining = total - current
            eta = remaining / speed if speed > 0 else 0

            self.speed_label.config(text=f"Speed: {speed:.1f} items/sec")
            self.eta_label.config(text=f"ETA: {int(eta)}s")

        self.update()

    def cancel(self):
        """Cancel operation"""
        self.cancelled = True
        self.destroy()

    def is_cancelled(self) -> bool:
        """Check if operation was cancelled"""
        return self.cancelled


# Settings Dialog
# ============================================================================

class SettingsDialog(tk.Toplevel):
    """Application settings dialog"""

    def __init__(self, parent, config: AppConfig):
        super().__init__(parent)
        self.title("Settings")
        self.geometry("400x350")
        self.transient(parent)
        self.resizable(False, False)

        # Center the dialog
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

        self.config = config
        self._create_widgets()

    def _create_widgets(self):
        """Create settings widgets"""
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Show hidden files
        self.show_hidden_var = tk.BooleanVar(
            value=self.config.get('show_hidden', False))
        ttk.Checkbutton(main_frame, text="Show hidden files",
                        variable=self.show_hidden_var).pack(anchor=tk.W, pady=5)

        # Confirm delete
        self.confirm_delete_var = tk.BooleanVar(
            value=self.config.get('confirm_delete', True))
        ttk.Checkbutton(main_frame, text="Confirm before deleting",
                        variable=self.confirm_delete_var).pack(anchor=tk.W, pady=5)

        # Confirm overwrite
        self.confirm_overwrite_var = tk.BooleanVar(
            value=self.config.get('confirm_overwrite', True))
        ttk.Checkbutton(main_frame, text="Confirm before overwriting",
                        variable=self.confirm_overwrite_var).pack(anchor=tk.W, pady=5)

        # Calculate folder sizes for Windows
        self.calculate_windows_sizes_var = tk.BooleanVar(
            value=self.config.get('calculate_windows_folder_sizes', True))
        ttk.Checkbutton(main_frame, text="Calculate folder sizes for Windows folders",
                        variable=self.calculate_windows_sizes_var).pack(anchor=tk.W, pady=5)

        # Calculate MTP folder sizes
        self.calculate_mtp_sizes_var = tk.BooleanVar(
            value=self.config.get('calculate_mtp_folder_sizes', True))
        ttk.Checkbutton(main_frame, text="Calculate folder sizes on MTP devices (slower)",
                        variable=self.calculate_mtp_sizes_var).pack(anchor=tk.W, pady=5)

        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(
            fill=tk.X, pady=20)

        # Font size
        font_frame = ttk.Frame(main_frame)
        font_frame.pack(fill=tk.X, pady=5)

        ttk.Label(font_frame, text="Font Size:").pack(
            side=tk.LEFT, padx=(0, 10))
        self.font_size_var = tk.IntVar(
            value=self.config.get('font_size', 10))
        ttk.Spinbox(font_frame, from_=8, to=16, width=10,
                    textvariable=self.font_size_var).pack(side=tk.LEFT)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))

        ttk.Button(button_frame, text="Save",
                   command=self.save).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel",
                   command=self.destroy).pack(side=tk.RIGHT)

    def save(self):
        """Save settings"""
        self.config.set('show_hidden', self.show_hidden_var.get())
        self.config.set('confirm_delete', self.confirm_delete_var.get())
        self.config.set('confirm_overwrite',
                        self.confirm_overwrite_var.get())
        self.config.set('calculate_windows_folder_sizes',
                        self.calculate_windows_sizes_var.get())
        self.config.set('calculate_mtp_folder_sizes',
                        self.calculate_mtp_sizes_var.get())
        self.config.set('font_size', self.font_size_var.get())
        self.config.save()

        messagebox.showinfo(
            "Settings", "Settings saved! Restart application for all changes to take effect.")
        self.destroy()


# Main Application Window
# ============================================================================

class DualPaneExplorer(tk.Tk):
    """Main application window"""

    def __init__(self):
        super().__init__()

        self.title("Dual-Pane File Explorer")

        # Load configuration
        self.app_config = AppConfig()

        # Apply saved geometry
        geometry = self.app_config.get('window_geometry', '1400x800')
        self.geometry(geometry)

        # Theme manager
        self.theme_manager = ThemeManager(self.app_config.get('theme', 'dark'))

        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.theme_manager.apply_theme(self, self.style)

        self._create_widgets()
        self._create_menu()
        self._setup_keybindings()

        # Restore last paths
        left_path = Path(self.app_config.get(
            'last_left_path', str(Path.home())))
        right_path = Path(self.app_config.get(
            'last_right_path', str(Path.home())))

        if left_path.exists():
            self.left_panel.navigate_to(left_path, add_to_history=False)
        if right_path.exists():
            self.right_panel.navigate_to(right_path, add_to_history=False)

        # Center window
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

        # Save state on close
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _create_widgets(self):
        """Create main UI components"""
        # Toolbar
        toolbar = ttk.Frame(self, padding=5)
        toolbar.pack(fill=tk.X)

        ttk.Button(toolbar, text="📋 Copy",
                   command=self.copy_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="✂️ Move",
                   command=self.move_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🗑️ Delete",
                   command=self.delete_files).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(toolbar, text="📁+ New Folder",
                   command=self.new_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📦 Archive",
                   command=self.create_archive).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🔄 Refresh All",
                   command=self.refresh_all).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(toolbar, text="🎨 Theme",
                   command=self.toggle_theme).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="⚙️ Settings",
                   command=self.show_settings).pack(side=tk.LEFT, padx=2)

        # Panels container
        panels_frame = ttk.Frame(self)
        panels_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create paned window for resizable panels
        self.paned = ttk.PanedWindow(panels_frame, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        # Left panel
        self.left_container = ttk.LabelFrame(
            self.paned, text="📂 Left Panel", padding=5)
        self.left_panel = FilePanel(
            self.left_container, "left", self.app_config, self.theme_manager, self)
        self.left_panel.pack(fill=tk.BOTH, expand=True)
        self.paned.add(self.left_container, weight=1)

        # Right panel
        self.right_container = ttk.LabelFrame(
            self.paned, text="📂 Right Panel", padding=5)
        self.right_panel = FilePanel(
            self.right_container, "right", self.app_config, self.theme_manager, self)
        self.right_panel.pack(fill=tk.BOTH, expand=True)
        self.paned.add(self.right_container, weight=1)

        # Preview pane
        self.preview_container = ttk.LabelFrame(
            self.paned, text="👁️ File Preview", padding=5)
        self._create_preview_pane()

        # Only add preview pane if enabled in settings (default: True)
        if self.app_config.get('show_preview', True):
            self.paned.add(self.preview_container, weight=1)

        # Bind click events for active panel tracking
        self.left_panel.tree.bind(
            '<Button-1>', lambda e: self._set_active_panel(self.left_panel))
        self.right_panel.tree.bind(
            '<Button-1>', lambda e: self._set_active_panel(self.right_panel))

        # Bind selection events for preview
        self.left_panel.tree.bind(
            '<<TreeviewSelect>>', lambda e: self._update_preview(self.left_panel))
        self.right_panel.tree.bind(
            '<<TreeviewSelect>>', lambda e: self._update_preview(self.right_panel))

        # Set initial active panel and highlight
        self._set_active_panel(self.left_panel)

        # Status bar
        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_bar = ttk.Label(status_frame, text="Ready | Dual-Pane File Explorer v1.0 | Helpful? Feel free to support me. You can reachout via www.linkedin.com/in/chukwuemeka-onyebuenyi-967304201",
                                    relief=tk.SUNKEN, anchor=tk.W, cursor="hand2")
        self.status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Bind click event to copy wallet address
        self.status_bar.bind('<Button-1>', self._copy_number_to_clipboard)

        # Version label
        ttk.Label(status_frame, text="v1.0.0",
                  relief=tk.SUNKEN).pack(side=tk.RIGHT)

        # Track active panel
        self.active_panel = self.left_panel
        self.left_panel.tree.bind(
            '<FocusIn>', lambda e: self._set_active_panel(self.left_panel))
        self.right_panel.tree.bind(
            '<FocusIn>', lambda e: self._set_active_panel(self.right_panel))

    def _create_menu(self):
        """Create menu bar"""
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Folder",
                              command=self.new_folder, accelerator="Ctrl+N")
        file_menu.add_command(label="New File", command=self.new_file)
        file_menu.add_separator()
        file_menu.add_command(
            label="Settings", command=self.show_settings, accelerator="Ctrl+,")
        file_menu.add_separator()
        file_menu.add_command(
            label="Exit", command=self.on_closing, accelerator="Ctrl+Q")

        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(
            label="Copy", command=self.copy_files, accelerator="Ctrl+C")
        edit_menu.add_command(
            label="Move", command=self.move_files, accelerator="Ctrl+X")
        edit_menu.add_command(
            label="Delete", command=self.delete_files, accelerator="Del")
        edit_menu.add_separator()
        edit_menu.add_command(label="Select All", command=lambda: self.active_panel.select_all(),
                              accelerator="Ctrl+A")

        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(
            label="Refresh", command=self.refresh_all, accelerator="F5")
        view_menu.add_command(label="Toggle Theme",
                              command=self.toggle_theme, accelerator="Ctrl+T")
        view_menu.add_separator()
        view_menu.add_checkbutton(label="Show Hidden Files",
                                  command=self.toggle_hidden_files)
        view_menu.add_checkbutton(label="Show Preview Pane",
                                  command=self.toggle_preview_pane,
                                  accelerator="Ctrl+P")

        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Search", command=lambda: self.active_panel.search_files(),
                               accelerator="Ctrl+F")
        tools_menu.add_command(label="Create Archive", command=self.create_archive,
                               accelerator="Ctrl+Z")
        tools_menu.add_separator()
        tools_menu.add_command(label="Properties", command=self.show_properties,
                               accelerator="Alt+Enter")

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Keyboard Shortcuts",
                              command=self.show_shortcuts)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)

    def _setup_keybindings(self):
        """Setup keyboard shortcuts"""
        self.bind('<Control-c>', lambda e: self.copy_files())
        self.bind('<Control-x>', lambda e: self.move_files())
        self.bind('<Delete>', lambda e: self.delete_files())
        self.bind('<F5>', lambda e: self.refresh_all())
        self.bind('<Control-n>', lambda e: self.new_folder())
        self.bind('<Control-f>', lambda e: self.active_panel.search_files())
        self.bind('<Control-z>', lambda e: self.create_archive())
        self.bind('<Control-t>', lambda e: self.toggle_theme())
        self.bind('<Control-comma>', lambda e: self.show_settings())
        self.bind('<Control-q>', lambda e: self.on_closing())
        self.bind('<Tab>', lambda e: self._switch_panel())
        self.bind('<Alt-Return>', lambda e: self.show_properties())
        self.bind('<Control-p>', lambda e: self.toggle_preview_pane())

    def _create_preview_pane(self):
        """Create file preview pane"""
        # Preview text widget with scrollbar
        preview_frame = ttk.Frame(self.preview_container)
        preview_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(preview_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Text widget for preview
        self.preview_text = tk.Text(
            preview_frame,
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            font=('Consolas', 10),
            state='disabled'
        )
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.preview_text.yview)

        # Info label at bottom
        self.preview_info = ttk.Label(
            self.preview_container,
            text="No file selected",
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.preview_info.pack(fill=tk.X, pady=(5, 0))

    def _update_preview(self, panel: FilePanel):
        """Update preview pane with selected file content"""
        selected = panel.get_selected_items()

        # Clear preview
        self.preview_text.config(state='normal')
        self.preview_text.delete('1.0', tk.END)

        if not selected or len(selected) != 1:
            self.preview_text.config(state='disabled')
            self.preview_info.config(
                text="No file selected" if not selected else "Multiple items selected")
            return

        path = selected[0]

        # Handle MTP paths - skip preview for now
        if isinstance(path, str) and is_mtp_path(path):
            self.preview_text.insert(
                '1.0', "Preview not available for MTP devices")
            self.preview_info.config(text="MTP file - preview not supported")
            self.preview_text.config(state='disabled')
            return

        # Convert to Path if string
        if isinstance(path, str):
            path = Path(path)

        # Check if it's a directory
        if path.is_dir():
            try:
                # Show directory contents summary
                items = list(path.iterdir())
                dirs = sum(1 for item in items if item.is_dir())
                files = len(items) - dirs
                total_size = sum(
                    item.stat().st_size for item in items if item.is_file())

                info = f"Directory: {path.name}\n\n"
                info += f"Total items: {len(items)}\n"
                info += f"Folders: {dirs}\n"
                info += f"Files: {files}\n"
                info += f"Total size: {format_size(total_size)}\n\n"
                info += "Contents:\n" + "-" * 40 + "\n"

                for item in sorted(items)[:50]:  # Limit to first 50 items
                    icon = "📁" if item.is_dir() else "📄"
                    info += f"{icon} {item.name}\n"

                if len(items) > 50:
                    info += f"\n... and {len(items) - 50} more items"

                self.preview_text.insert('1.0', info)
                self.preview_info.config(text=f"Directory: {len(items)} items")
            except (PermissionError, OSError) as e:
                self.preview_text.insert(
                    '1.0', f"Cannot access directory:\n{str(e)}")
                self.preview_info.config(text="Access denied")

            self.preview_text.config(state='disabled')
            return

        # Handle files
        try:
            file_size = path.stat().st_size

            # Check file size (limit preview to 1MB)
            if file_size > 1024 * 1024:
                self.preview_text.insert('1.0',
                                         f"File too large to preview\n\n"
                                         f"File: {path.name}\n"
                                         f"Size: {format_size(file_size)}\n"
                                         f"Type: {path.suffix[1:].upper() if path.suffix else 'Unknown'}")
                self.preview_info.config(
                    text=f"File too large: {format_size(file_size)}")
                self.preview_text.config(state='disabled')
                return

            # Try to read as text
            text_extensions = {
                '.txt', '.py', '.js', '.java', '.cpp', '.c', '.h', '.hpp',
                '.cs', '.html', '.css', '.xml', '.json', '.yaml', '.yml',
                '.md', '.rst', '.log', '.ini', '.cfg', '.conf', '.sh',
                '.bat', '.ps1', '.sql', '.r', '.rb', '.php', '.go', '.rs'
            }

            if path.suffix.lower() in text_extensions or file_size < 1024:
                try:
                    # Try different encodings
                    content = None
                    for encoding in ['utf-8', 'latin-1', 'cp1252']:
                        try:
                            with open(path, 'r', encoding=encoding) as f:
                                content = f.read()
                            break
                        except UnicodeDecodeError:
                            continue

                    if content:
                        # Limit lines for very long files
                        lines = content.split('\n')
                        if len(lines) > 1000:
                            content = '\n'.join(
                                lines[:1000]) + f"\n\n... ({len(lines) - 1000} more lines)"

                        self.preview_text.insert('1.0', content)
                        self.preview_info.config(
                            text=f"{path.name} - {format_size(file_size)} - {len(lines)} lines")
                    else:
                        raise UnicodeDecodeError(
                            'all', b'', 0, 0, 'All encodings failed')

                except UnicodeDecodeError:
                    self.preview_text.insert('1.0',
                                             f"Cannot preview binary file\n\n"
                                             f"File: {path.name}\n"
                                             f"Size: {format_size(file_size)}\n"
                                             f"Type: {path.suffix[1:].upper() if path.suffix else 'Binary'}")
                    self.preview_info.config(
                        text=f"Binary file: {format_size(file_size)}")
            else:
                # Binary file or unsupported type
                self.preview_text.insert('1.0',
                                         f"Preview not available\n\n"
                                         f"File: {path.name}\n"
                                         f"Size: {format_size(file_size)}\n"
                                         f"Type: {path.suffix[1:].upper() if path.suffix else 'Unknown'}\n\n"
                                         f"Supported text formats:\n"
                                         f"Source code, scripts, config files, logs, etc.")
                self.preview_info.config(
                    text=f"{path.suffix[1:].upper()} file - {format_size(file_size)}")

        except (PermissionError, OSError) as e:
            self.preview_text.insert('1.0', f"Cannot read file:\n{str(e)}")
            self.preview_info.config(text="Access denied")

        self.preview_text.config(state='disabled')

    def _set_active_panel(self, panel: FilePanel):
        """Set active panel with visual highlighting"""
        # Update previous active panel
        if hasattr(self, 'active_panel') and self.active_panel:
            self.active_panel.is_active = False
            # Remove highlight from inactive panel container
            if self.active_panel == self.left_panel:
                self.left_container.configure(relief='flat', borderwidth=1)
            else:
                self.right_container.configure(relief='flat', borderwidth=1)

        # Set new active panel
        self.active_panel = panel
        panel.is_active = True

        # Add highlight to active panel container
        if panel == self.left_panel:
            self.left_container.configure(relief='solid', borderwidth=2)
        else:
            self.right_container.configure(relief='solid', borderwidth=2)

    def _switch_panel(self):
        """Switch focus between panels"""
        if self.active_panel == self.left_panel:
            self.right_panel.tree.focus_set()
        else:
            self.left_panel.tree.focus_set()

    def _copy_number_to_clipboard(self, event=None):
        """Copy Linkedin address to clipboard"""
        wallet_address = "www.linkedin.com/in/chukwuemeka-onyebuenyi-967304201"
        self.clipboard_clear()
        self.clipboard_append(wallet_address)
        self.status_bar.config(text="✓ Linkedin URL copied to clipboard!")
        # Reset status after 3 seconds
        self.after(3000, lambda: self.status_bar.config(
            text="Ready | Dual-Pane File Explorer v1.0 | Helpful? Feel free to support me. You can reachout via www.linkedin.com/in/chukwuemeka-onyebuenyi-967304201"))

    def _get_inactive_panel(self) -> FilePanel:
        """Get the inactive panel"""
        return self.right_panel if self.active_panel == self.left_panel else self.left_panel

    def copy_files(self):
        """Copy selected files to other panel"""
        selected = self.active_panel.get_selected_items()
        if not selected:
            messagebox.showwarning(
                "No Selection", "Please select files to copy")
            return

        destination = self._get_inactive_panel().current_path

        if not messagebox.askyesno("Confirm Copy",
                                   f"Copy {len(selected)} item(s) to:\n{destination}?"):
            return

        self._execute_operation(
            "Copying files...", selected, destination, "copy")

    def move_files(self):
        """Move selected files to other panel"""
        selected = self.active_panel.get_selected_items()
        if not selected:
            messagebox.showwarning(
                "No Selection", "Please select files to move")
            return

        destination = self._get_inactive_panel().current_path

        if not messagebox.askyesno("Confirm Move",
                                   f"Move {len(selected)} item(s) to:\n{destination}?"):
            return

        self._execute_operation(
            "Moving files...", selected, destination, "move")

    def delete_files(self):
        """Delete selected files"""
        selected = self.active_panel.get_selected_items()
        if not selected:
            messagebox.showwarning(
                "No Selection", "Please select files to delete")
            return

        if self.app_config.get('confirm_delete', True):
            # Handle both Path objects and MTP path strings
            file_list = "\n".join([
                p.name if isinstance(p, Path) else str(p).split('/')[-1]
                for p in selected
            ])

            # Use custom dialog with Yes/Yes to All/No options
            result = DeleteConfirmDialog.ask_delete(
                self,
                "Confirm Delete",
                f"Permanently delete {len(selected)} item(s)?",
                file_list
            )

            if not result:  # User clicked No or closed dialog
                return

        self._execute_operation("Deleting files...", selected, None, "delete")

    def new_folder(self):
        """Create new folder in active panel"""
        name = CenteredInputDialog.ask_string(
            self, "New Folder", "Enter folder name:")
        if not name:
            return

        new_path = self.active_panel.current_path / name

        try:
            new_path.mkdir(exist_ok=False)
            self.active_panel.refresh()
            self.status_bar.config(text=f"✓ Created folder: {name}")
        except FileExistsError:
            messagebox.showerror("Error", f"Folder already exists: {name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create folder: {e}")

    def new_file(self):
        """Create new file in active panel"""
        name = CenteredInputDialog.ask_string(
            self, "New File", "Enter file name:")
        if not name:
            return

        new_path = self.active_panel.current_path / name

        try:
            new_path.touch(exist_ok=False)
            self.active_panel.refresh()
            self.status_bar.config(text=f"✓ Created file: {name}")
        except FileExistsError:
            messagebox.showerror("Error", f"File already exists: {name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create file: {e}")

    def create_archive(self):
        """Create archive from selected files"""
        selected = self.active_panel.get_selected_items()
        if not selected:
            messagebox.showwarning(
                "No Selection", "Please select files to archive")
            return

        self.active_panel.create_archive_dialog(selected)

    def refresh_all(self):
        """Refresh both panels"""
        self.left_panel.refresh()
        self.right_panel.refresh()
        self.status_bar.config(text="✓ Panels refreshed")

    def toggle_theme(self):
        """Toggle between light and dark themes"""
        new_theme = self.theme_manager.switch_theme()
        self.theme_manager.apply_theme(self, self.style)
        self.app_config.set('theme', new_theme)
        self.app_config.save()
        self.status_bar.config(text=f"✓ Switched to {new_theme} theme")

        # Refresh panels to apply theme
        self.left_panel.refresh()
        self.right_panel.refresh()

    def toggle_hidden_files(self):
        """Toggle showing hidden files"""
        current = self.app_config.get('show_hidden', False)
        self.app_config.set('show_hidden', not current)
        self.app_config.save()
        self.refresh_all()
        status = "showing" if not current else "hiding"
        self.status_bar.config(text=f"✓ Now {status} hidden files")

    def toggle_preview_pane(self):
        """Toggle preview pane visibility"""
        current = self.app_config.get('show_preview', True)
        new_state = not current
        self.app_config.set('show_preview', new_state)
        self.app_config.save()

        if new_state:
            # Show preview pane
            self.paned.add(self.preview_container, weight=1)
            self.status_bar.config(text="✓ Preview pane shown")
        else:
            # Hide preview pane
            self.paned.remove(self.preview_container)
            self.status_bar.config(text="✓ Preview pane hidden")

    def show_settings(self):
        """Show settings dialog"""
        SettingsDialog(self, self.app_config)

    def show_properties(self):
        """Show properties of selected item"""
        selected = self.active_panel.get_selected_items()
        if len(selected) == 1:
            # Get file info from the active panel
            file_info = None
            if hasattr(self.active_panel, '_current_items'):
                for item in self.active_panel._current_items:
                    if str(item.path) == str(selected[0]):
                        file_info = item
                        break

            PropertiesDialog(self, selected[0], file_info)
        elif len(selected) > 1:
            messagebox.showinfo("Multiple Selection",
                                f"{len(selected)} items selected\n"
                                f"Properties dialog works with single selection only")

    def show_shortcuts(self):
        """Show keyboard shortcuts"""
        shortcuts = """
Keyboard Shortcuts:

Navigation:
  Tab             - Switch between panels
  Backspace       - Go to parent directory
  Enter           - Open folder/file
  Ctrl+A          - Select all items

File Operations:
  Ctrl+C          - Copy selected items
  Ctrl+X          - Move selected items
  Delete          - Delete selected items
  Ctrl+N          - Create new folder
  Ctrl+Z          - Create archive
  F2              - Rename (in context menu)

View & Search:
  F5              - Refresh panels
  Ctrl+F          - Search in current folder
  Ctrl+T          - Toggle theme
  Ctrl+P          - Toggle preview pane
  Alt+Enter       - Show properties

Application:
  Ctrl+,          - Settings
  Ctrl+Q          - Quit
        """

        dialog = tk.Toplevel(self)
        dialog.title("Keyboard Shortcuts")
        dialog.geometry("450x500")
        dialog.transient(self)
        dialog.resizable(False, False)

        # Center the dialog
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        text = tk.Text(frame, wrap=tk.WORD, font=('Courier', 9))
        text.pack(fill=tk.BOTH, expand=True)
        text.insert('1.0', shortcuts)
        text.config(state='disabled')

        ttk.Button(frame, text="Close",
                   command=dialog.destroy).pack(pady=(10, 0))

    def show_about(self):
        """Show about dialog"""
        about_text = """
Dual-Pane File Explorer
Version 1.0.0
Built by Chukwuemeka Onyebuenyi

If this was helpful to you, Feel free to support 
me with donations to my USDT TRC20 wallet: 
TPdPTVZm4nSYU24mes6g8nrfp2rMDAEJMj

A modern, feature-rich dual-pane file manager
built with Python and Tkinter.

Features:
• Dual-panel interface for easy file management
• Copy, move, and delete operations
• Archive creation and extraction
• Search functionality
• Drag-and-drop support
• Dark and light themes
• Keyboard shortcuts
• File properties viewer
• Cross-platform support


© 2025 - Open Source Project
        """

        dialog = tk.Toplevel(self)
        dialog.title("About")
        dialog.geometry("400x450")
        dialog.transient(self)
        dialog.resizable(False, False)

        # Center the dialog
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Icon
        ttk.Label(frame, text="📂🔄📁", font=('', 48)).pack(pady=10)

        # Text
        text = tk.Text(frame, wrap=tk.WORD, font=('', 10), height=15)
        text.pack(fill=tk.BOTH, expand=True, pady=10)
        text.insert('1.0', about_text)
        text.config(state='disabled')

        ttk.Button(frame, text="Close", command=dialog.destroy).pack()

    def _execute_operation(self, title: str, sources: List[Path],
                           destination: Optional[Path], operation: str):
        """Execute file operation in background thread"""
        progress_dialog = ProgressDialog(self, title, len(sources))
        result = None

        def worker():
            nonlocal result

            def progress_callback(current, total):
                if not progress_dialog.is_cancelled():
                    self.after(
                        0, lambda: progress_dialog.update_progress(current, total))

            try:
                if operation == "copy":
                    result = FileOperations.copy_files(
                        sources, destination, progress_callback)
                elif operation == "move":
                    result = FileOperations.move_files(
                        sources, destination, progress_callback)
                elif operation == "delete":
                    result = FileOperations.delete_files(
                        sources, progress_callback)
            except Exception as e:
                result = OperationResult(
                    success=False, items_processed=0, errors=[str(e)])

            # Close dialog
            self.after(0, progress_dialog.destroy)

        # Start operation thread
        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

        # Wait for dialog to close
        self.wait_window(progress_dialog)

        # Wait for thread to complete
        thread.join(timeout=1.0)

        # Show result
        if result:
            if result.success:
                self.status_bar.config(
                    text=f"✓ {operation.capitalize()} completed: {result.items_processed} items")
                messagebox.showinfo("Success",
                                    f"Operation completed successfully!\n"
                                    f"Processed {result.items_processed} item(s)")
            else:
                error_msg = "\n".join(result.errors[:10])
                if len(result.errors) > 10:
                    error_msg += f"\n... and {len(result.errors) - 10} more errors"
                messagebox.showerror("Errors",
                                     f"Operation completed with errors:\n\n{error_msg}")
                self.status_bar.config(
                    text=f"✗ {operation.capitalize()} completed with errors")

            # Refresh both panels
            self.refresh_all()

    def on_closing(self):
        """Handle application closing"""
        # Save window geometry
        self.app_config.set('window_geometry', self.geometry())

        # Save last paths
        self.app_config.set('last_left_path', str(
            self.left_panel.current_path))
        self.app_config.set('last_right_path', str(
            self.right_panel.current_path))

        # Save configuration
        self.app_config.save()

        # Close application
        self.destroy()


# Application Main Entry Point
# ============================================================================

def main():
    """Run the application"""
    app = DualPaneExplorer()
    app.mainloop()


if __name__ == "__main__":
    main()
