# Window Utils - Windows API Utilities

## Overview

`window_utils.py` is a pure Win32 API utility module that provides low-level window management capabilities for InfraCommand. It contains no Qt dependencies and focuses on core Windows API operations for window positioning, process management, and display handling.

## Features

### 🖥️ Display Management
- **DPI Awareness**: Automatic DPI scaling support for high-resolution displays
- **Monitor Detection**: Primary monitor workarea detection
- **Quadrant Calculations**: Smart window positioning in screen quadrants

### 🪟 Window Operations
- **UAC-Aware Movement**: Safe window positioning that handles elevated processes
- **Window Enumeration**: Find windows by process ID, class names, titles, or executables
- **Process Integration**: Extract process information and image names

### ⚡ Performance Optimized
- **Pure Win32 API**: Direct Windows API calls for maximum performance
- **No External Dependencies**: Only uses `pywin32` and `ctypes`
- **Memory Efficient**: Minimal overhead and resource usage

## Core Functions

### DPI Awareness
```python
enable_dpi_awareness()
```
Enables DPI awareness for the current process, supporting:
- Per-Monitor V2 (Windows 10 1703+)
- Per-Monitor V1 (Windows 8.1+)
- System DPI (Fallback)

### Monitor Geometry
```python
primary_workarea_physical() -> Tuple[int, int, int, int]
```
Returns primary monitor workarea in physical pixels (Left, Top, Right, Bottom).

```python
get_quadrant_physical(quad: str, fill_ratio: float = 0.995, edge_margin_ratio: float = 0.01) -> Tuple[int, int, int, int]
```
Calculates window position and size for screen quadrants:
- `quad`: "TL", "TR", "BL", "BR" (Top-Left, Top-Right, Bottom-Left, Bottom-Right)
- `fill_ratio`: Percentage of screen to use (default 99.5%)
- `edge_margin_ratio`: Margin from screen edges (default 1%)

### Window Movement
```python
move_window(hwnd: int, x: int, y: int, w: int, h: int) -> bool
```
Safely moves a window to specified position and size. Handles UAC elevation boundaries.

```python
safe_set_window_pos_hwnd(hwnd: int, x: int, y: int, w: int, h: int) -> bool
```
Low-level window positioning with UAC boundary detection.

### Process Management
```python
get_image_name_basename(pid: int) -> str
```
Extracts the executable name (basename) for a given process ID.

### Window Discovery
```python
enum_visible_toplevel_windows() -> List[int]
```
Enumerates all visible top-level windows, returning a list of window handles.

```python
find_windows_by_criteria(
    pid: int = None,
    class_names: List[str] = None,
    title_contains: List[str] = None,
    image_names: List[str] = None,
    session_id: int = None
) -> List[int]
```
Advanced window search with multiple criteria:
- **pid**: Match specific process ID
- **class_names**: Match window class names
- **title_contains**: Match window titles (case-insensitive)
- **image_names**: Match executable names
- **session_id**: Match RDSH session ID for remote desktop scenarios

## Technical Details

### UAC Handling
The module includes sophisticated UAC (User Account Control) boundary detection:
- Tracks windows that refuse movement due to elevation differences
- Gracefully handles access denied errors
- Maintains a blacklist of unmanageable windows

### Error Handling
- All functions include comprehensive exception handling
- Graceful degradation when Windows API calls fail
- No crashes on invalid window handles or process IDs

### Performance Considerations
- Direct Win32 API calls for maximum speed
- Minimal memory allocations
- Efficient window enumeration with early termination
- Cached results where appropriate

## Usage in InfraCommand

This module is used by InfraCommand for:
- **Smart Window Placement**: Positioning launched tools in screen quadrants
- **Process Monitoring**: Tracking launched applications
- **UAC Management**: Handling elevated processes safely
- **Multi-Monitor Support**: Working across different display configurations

## Dependencies

- `ctypes` - Windows API access
- `pywin32` - Windows API bindings
- `typing` - Type hints

## Compatibility

- **Windows 10/11**: Full support
- **Windows 8.1**: Limited DPI awareness
- **Python 3.7+**: Type hints support required

## Example Usage

```python
import window_utils

# Enable DPI awareness
window_utils.enable_dpi_awareness()

# Get primary monitor info
workarea = window_utils.primary_workarea_physical()
print(f"Workarea: {workarea}")

# Calculate top-right quadrant
x, y, w, h = window_utils.get_quadrant_physical("TR")
print(f"Top-right quadrant: {x}, {y}, {w}, {h}")

# Find PowerShell windows
ps_windows = window_utils.find_windows_by_criteria(
    image_names=["powershell.exe", "pwsh.exe"]
)

# Move a window
if ps_windows:
    success = window_utils.move_window(ps_windows[0], x, y, w, h)
    print(f"Window moved: {success}")
```

## Architecture Notes

This module is designed to be:
- **Stateless**: No global state or persistent data
- **Pure Functions**: No side effects beyond Windows API calls
- **Reusable**: Can be used independently of InfraCommand
- **Testable**: Easy to unit test individual functions

The module serves as the foundation for InfraCommand's window management capabilities, providing a clean abstraction over the complex Windows API.
