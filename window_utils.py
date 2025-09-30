# window_utils.py
# Pure Win32 window utilities - no Qt dependencies, no state dependencies

import ctypes
import os
from ctypes import wintypes
from typing import Tuple, List
import win32gui
import win32process
import win32api

# ---------- Constants ----------
MONITORINFOF_PRIMARY = 1
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

# ---------- DPI Awareness ----------
def enable_dpi_awareness():
    """Enable DPI awareness for the process."""
    try:
        # Per-Monitor V2
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except Exception:
        pass
    try:
        # Per-Monitor V1
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        # System DPI
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ---------- Monitor / Geometry ----------
def primary_workarea_physical() -> Tuple[int, int, int, int]:
    """Get primary monitor workarea in physical pixels (L, T, R, B)."""
    mons = win32api.EnumDisplayMonitors()
    for hmon, hdc, r in mons:
        inf = win32api.GetMonitorInfo(hmon)
        if inf.get('Flags', 0) == MONITORINFOF_PRIMARY:
            L, T, R, B = inf['Work']
            return L, T, R, B
    # Fallback to 1920x1080
    return 0, 0, 1920, 1080

def get_quadrant_physical(quad: str, fill_ratio: float = 0.995, edge_margin_ratio: float = 0.01) -> Tuple[int, int, int, int]:
    """
    Calculate quadrant geometry in physical pixels.
    
    Args:
        quad: One of "TL", "TR", "BL", "BR", "FULL"
        fill_ratio: How much of the quadrant to fill (0.0-1.0)
        edge_margin_ratio: Margin from screen edges (0.0-1.0)
    
    Returns:
        Tuple of (x, y, width, height) in pixels
    """
    L, T, R, B = primary_workarea_physical()
    W, H = R - L, B - T
    
    # Apply edge margin
    outer = max(8, int(edge_margin_ratio * min(W, H)))
    L += outer
    T += outer
    R -= outer
    B -= outer
    W, H = R - L, B - T
    
    # Calculate quadrant
    half_w, half_h = W // 2, H // 2
    
    quadrant_map = {
        "TL": (L, T, half_w, half_h),
        "TR": (L + half_w, T, half_w, half_h),
        "BL": (L, T + half_h, half_w, half_h),
        "BR": (L + half_w, T + half_h, half_w, half_h),
        "FULL": (L, T, W, H)
    }
    
    x, y, w, h = quadrant_map.get(quad, quadrant_map["FULL"])
    
    # Apply fill ratio
    target_w = int(w * fill_ratio)
    target_h = int(h * fill_ratio)
    x += (w - target_w) // 2
    y += (h - target_h) // 2
    
    return x, y, target_w, target_h

def clamp_rect_to_primary(x: int, y: int, w: int, h: int) -> Tuple[int, int, int, int]:
    """Clamp rectangle to primary monitor workarea."""
    L, T, R, B = primary_workarea_physical()
    
    if x < L:
        x = L
    if y < T:
        y = T
    if x + w > R:
        x = R - w
    if y + h > B:
        y = B - h
        
    return x, y, w, h

# ---------- Process Info ----------
def _get_process_session_id(pid: int) -> int:
    """Get the session ID for a process (important for RDSH scenarios)."""
    try:
        import win32ts
        return win32ts.ProcessIdToSessionId(pid)
    except Exception:
        # Fallback: return current session if we can't determine
        try:
            return win32ts.WTSGetCurrentSessionId()
        except Exception:
            return 0

def get_image_name_basename(pid: int) -> str:
    """Get the executable name for a process ID."""
    try:
        kernel32 = ctypes.windll.kernel32
        h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not h:
            return ""
        
        buf = ctypes.create_unicode_buffer(32768)
        size = wintypes.DWORD(len(buf))
        
        if kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            kernel32.CloseHandle(h)
            return os.path.basename(buf.value).lower()
        
        kernel32.CloseHandle(h)
    except Exception:
        pass
    return ""

# ---------- Window Movement (UAC-aware) ----------
_user32 = ctypes.WinDLL("user32", use_last_error=True)
_SetWindowPos = _user32.SetWindowPos
_SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
_SetWindowPos.restype = wintypes.BOOL
_HWND_TOP = wintypes.HWND(0)
_SWP_FLAGS = 0x0004 | 0x0040  # NOZORDER | SHOWWINDOW

# Track HWNDs that refused moves (likely elevated / UAC boundary)
_UNMANAGEABLE_HWND = set()

def safe_set_window_pos_hwnd(hwnd: int, x: int, y: int, w: int, h: int) -> bool:
    """
    Safely move a window, handling UAC elevation boundaries.
    
    Returns:
        True if successful, False if access denied or other error
    """
    if hwnd in _UNMANAGEABLE_HWND:
        return False
    
    ok = _SetWindowPos(wintypes.HWND(hwnd), _HWND_TOP, int(x), int(y), int(w), int(h), _SWP_FLAGS)
    if ok:
        return True
    
    err = ctypes.get_last_error()
    if err == 5:  # ERROR_ACCESS_DENIED
        _UNMANAGEABLE_HWND.add(hwnd)
        return False
    
    return False

def move_window(hwnd: int, x: int, y: int, w: int, h: int) -> bool:
    """
    Move and resize a window to the specified coordinates.
    
    Args:
        hwnd: Window handle
        x, y: Top-left position
        w, h: Width and height
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # Restore if minimized
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
    except Exception:
        pass
    
    # Clamp to screen bounds
    x, y, w, h = clamp_rect_to_primary(int(x), int(y), int(w), int(h))
    
    # Use UAC-aware positioning
    return safe_set_window_pos_hwnd(hwnd, x, y, w, h)

# ---------- Window Enumeration ----------
def get_window_area(hwnd: int) -> int:
    """Calculate the visible area of a window."""
    try:
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        return max(0, r - l) * max(0, b - t)
    except Exception:
        return 0

def enum_visible_toplevel_windows() -> List[int]:
    """
    Enumerate all visible top-level windows.
    
    Returns:
        List of window handles (HWNDs)
    """
    windows = []
    
    def enum_callback(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetParent(hwnd):
                return True
            windows.append(hwnd)
        except Exception:
            pass
        return True
    
    try:
        win32gui.EnumWindows(enum_callback, None)
    except Exception:
        pass
    
    return windows

def find_windows_by_criteria(
    pid: int = None,
    class_names: List[str] = None,
    title_contains: List[str] = None,
    image_names: List[str] = None,
    session_id: int = None
) -> List[int]:
    """
    Find windows matching the given criteria.
    
    Args:
        pid: Process ID to match
        class_names: List of window class names to match
        title_contains: List of strings that should be in the window title
        image_names: List of executable names to match
        session_id: RDSH session ID to match (for remote desktop scenarios)
    
    Returns:
        List of matching window handles
    """
    matching_windows = []
    
    # Normalize inputs
    class_set = set(class_names or [])
    title_searches = [s.lower() for s in (title_contains or []) if s]
    image_set = set((name or "").lower() for name in (image_names or []))
    
    def enum_callback(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetParent(hwnd):
                return True
            
            # Check class name
            if class_names:
                window_class = win32gui.GetClassName(hwnd)
                if window_class not in class_set:
                    # Special case for ApplicationFrameWindow
                    if window_class != "ApplicationFrameWindow":
                        return True
            
            # Check PID
            if pid is not None:
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if int(window_pid) != int(pid):
                    return True
            
            # Check title
            if title_searches:
                window_title = (win32gui.GetWindowText(hwnd) or "").lower()
                if not any(search in window_title for search in title_searches):
                    # Allow ApplicationFrameWindow with matching title even if class doesn't match
                    if not class_names or win32gui.GetClassName(hwnd) != "ApplicationFrameWindow":
                        return True
            
            # Check image name
            if image_set:
                try:
                    _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                    image_name = get_image_name_basename(int(window_pid))
                    if image_name not in image_set:
                        return True
                except Exception:
                    return True
            
            # Check RDSH session ID (for remote desktop scenarios)
            if session_id is not None:
                try:
                    _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                    window_session_id = _get_process_session_id(window_pid)
                    if window_session_id != session_id:
                        return True
                except Exception:
                    return True
            
            # Check minimum size (filter out tiny windows)
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            if (r - l) < 200 or (b - t) < 120:
                return True
            
            matching_windows.append(hwnd)
            
        except Exception:
            pass
        return True
    
    try:
        win32gui.EnumWindows(enum_callback, None)
    except Exception:
        pass
    
    return matching_windows

# ---------- Explorer-specific utilities ----------
def get_explorer_windows() -> List[int]:
    """
    Get all File Explorer top-level windows.
    
    Returns:
        List of Explorer window handles
    """
    return find_windows_by_criteria(
        class_names=["CabinetWClass", "ExploreWClass", "ApplicationFrameWindow"],
        image_names=["explorer.exe"]
    )