# -*- coding: utf-8 -*-
"""
🚀 InfraCommand Support Cockpit v.0.3.81
==========================================

A powerful Windows system administration tool with modern UI and advanced automation capabilities.

## ✨ Highlights & Special Features

### 🎯 Core Features
- **Multi-Menu System**: Dynamic menu switching with JSON configuration
- **Smart Button Grid**: Auto-adapting button layout based on window width
- **UAC Management**: Intelligent elevation handling for admin tools
- **Process Monitoring**: Real-time process and window tracking
- **Baseline Automation**: Automated system baseline collection

### 🎨 Modern UI/UX
- **Neon Glass Effect**: Beautiful translucent windows with neon styling
- **Frameless Design**: Clean, modern interface without traditional borders
- **Always on Top**: Stays visible while working with other applications
- **Responsive Layout**: Adapts to different screen sizes and configurations
- **Dark Theme**: Professional dark interface optimized for long sessions

### ⚡ Advanced Capabilities
- **JSON Configuration**: Flexible, version-controlled configuration system
- **Multi-Threading**: Non-blocking operations with background processing
- **Error Handling**: Comprehensive error logging and user feedback
- **Tool Integration**: Seamless integration with Windows admin tools
- **Console Logging**: Detailed operation logging with timestamps

### 🛠️ Technical Excellence
- **Clean Architecture**: Modular, maintainable codebase
- **Type Safety**: Proper type handling and validation
- **Memory Efficient**: Optimized resource usage
- **Cross-Platform Ready**: Built with Qt for future platform expansion
- **GitHub Ready**: Professional documentation and version control

## 🚀 Quick Start
1. Configure your tools in `grid.config.json`
2. Run `python infracommand.py`
3. Select your menu and start working!

## 📋 System Requirements
- Windows 10/11
- Python 3.7+
- Qt5/PyQt5
- Administrator privileges (for some tools)

---
*Built with ❤️ for Windows system administrators*
"""

import sys, os, time, json, subprocess, ctypes, shlex, shutil, html, threading, tempfile, math
from queue import Queue, Empty
from ctypes import wintypes
from datetime import datetime, timezone
from dataclasses import dataclass, replace, asdict
from typing import Dict, Tuple, List, Optional, Any
import importlib.util

# Import our window utilities - KREATIVER ANSATZ!
def get_window_utils():
    """Dynamischer Import mit mehreren Fallback-Strategien"""
    
    # Strategie 1: Direkter Import
    try:
        import window_utils
        return window_utils
    except ImportError:
        pass
    
    # Strategie 2: Import aus dem aktuellen Verzeichnis
    try:
        import sys
        import os
        current_dir = os.path.dirname(os.path.abspath(__file__))
        if current_dir not in sys.path:
            sys.path.insert(0, current_dir)
        import window_utils
        return window_utils
    except ImportError:
        pass
    
    # Strategie 3: Datei-basierter Import (für Entwicklung)
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        module_path = os.path.join(script_dir, "window_utils.py")
        
        if os.path.exists(module_path):
            spec = importlib.util.spec_from_file_location("window_utils", module_path)
            if spec and spec.loader:
                wu = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(wu)
                return wu
    except Exception:
        pass
    
    # Strategie 4: Notfall - erstelle ein Dummy-Modul
    print("WARNING: window_utils not found, using fallback functions")
    class DummyWindowUtils:
        def enable_dpi_awareness(self): pass
        def get_quadrant_physical(self, *args): return 1
        def move_window(self, *args): pass
        def get_image_name_basename(self, *args): return "unknown"
        def find_windows_by_criteria(self, *args): return []
    
    return DummyWindowUtils()

# Verwende die Funktion
wu = get_window_utils()
enable_dpi_awareness = wu.enable_dpi_awareness
get_quadrant_physical = wu.get_quadrant_physical
move_window = wu.move_window
get_image_name_basename = wu.get_image_name_basename
find_windows_by_criteria = wu.find_windows_by_criteria
get_explorer_windows = wu.get_explorer_windows
get_window_area = wu.get_window_area
clamp_rect_to_primary = wu.clamp_rect_to_primary

from PySide6.QtWidgets import (
    QApplication, QWidget, QGridLayout, QVBoxLayout,
    QToolButton, QLabel, QTextEdit, QSplitter, QHBoxLayout, QMessageBox,
    QComboBox, QGraphicsDropShadowEffect, QGraphicsBlurEffect
)
from PySide6.QtCore import Qt, QTimer, QRect, QPoint
from PySide6.QtGui import QGuiApplication, QTextCursor, QFont, QIcon, QPixmap, QColor
import win32gui, win32process, win32con, win32api

@dataclass
class Config:
    CONFIG_FILENAME: str = "grid.config.json"
    TOOLS_DIR: str = r"C:\Support\Tools"
    SCRIPTS_DIR: str = "./scripts"
    LOCAL_LOG_DIR: str = r"C:\Support\Logs\Cockpit"
    CENTRAL_LOG_DIR: str = r"C:\Support\Logs_CENTRAL\Cockpit"
    RR_ORDER: Tuple[str, str, str] = ("BR", "TL", "BL")
    FILL_RATIO: float = 0.995
    EDGE_MARGIN_RATIO: float = 0.01
    FALLBACK_DELAY_MS: int = 1800
    MAX_TOOL_BUTTONS: int = 16
    BUTTON_FEEDBACK_MS: int = 1500
    BUTTON_COLUMNS: int = 0  # 0 = dynamic, >0 = fixed columns
    GUI_START_QUADRANT: str = "TR"  # TR, TL, BL, BR
    BASELINE_PATH: Optional[str] = None
    BASELINE_DELAY_MS: int = 1200
    BASELINE_ARGS: Tuple[str, ...] = ()
    BASELINE_UAC: bool = False
    DEFAULT_CLUSTER: Optional[str] = None
    CONSOLE_LOG_LEVEL: str = "INFO"  # DEBUG, INFO, WARNING, ERROR

def _expand_env(p: str) -> str:
    try: return os.path.expandvars(p)
    except Exception: return p

def _is_abs(p: str) -> bool:
    try:
        return os.path.isabs(p) or (len(p) > 1 and p[1] == ":" and (p[2:3] == "\\" or p[2:3] == "/"))
    except Exception:
        return False

def _discover_config_path(argv: List[str], base_dir: str) -> str:
    # In PyInstaller EXE, use the directory where the EXE is located
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller EXE
        base_dir = os.path.dirname(sys.executable)
    
    for i,a in enumerate(argv):
        if a.lower() in ("--config","-configpath"):
            if i+1 < len(argv):
                c = os.path.abspath(argv[i+1])
                if os.path.isdir(c):
                    for n in ("grid.config.json","config.supportcockpit.json","config.json"):
                        p=os.path.join(c,n)
                        if os.path.isfile(p): return p
                    return ""
                return c if os.path.isfile(c) else ""
    for n in ("grid.config.json","config.supportcockpit.json","config.json"):
        p=os.path.join(base_dir,n)
        if os.path.isfile(p): return p
    return ""

def _load_json(path: str) -> dict:  # type: ignore
    with open(path, "r", encoding="utf-8") as f: return json.load(f)

def apply_json_overrides(cfg: Config, j: dict) -> Config:
    c = cfg
    for key, attr in [
        ("TOOLS_DIR","TOOLS_DIR"), ("SCRIPTS_DIR","SCRIPTS_DIR"),
        ("LOCAL_LOG_DIR","LOCAL_LOG_DIR"), ("CENTRAL_LOG_DIR","CENTRAL_LOG_DIR")
    ]:
        v = j.get(key)
        if isinstance(v, str) and v:
            expanded = _expand_env(v)
            # Für SCRIPTS_DIR: Auflösung zu absolutem Pfad
            if attr == "SCRIPTS_DIR" and expanded.startswith("./"):
                if getattr(sys, 'frozen', False):
                    # Running as PyInstaller EXE
                    script_dir = os.path.dirname(sys.executable)
                else:
                    # Running as script
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                expanded = os.path.join(script_dir, expanded[2:])
                expanded = os.path.abspath(expanded)
            c = replace(c, **{attr: expanded})  # type: ignore

    if isinstance(j.get("RR_ORDER"), list) and len(j["RR_ORDER"]) >= 3:
        rr = tuple(str(j["RR_ORDER"][i]) for i in range(3))
        c = replace(c, RR_ORDER=rr)  # type: ignore
    if isinstance(j.get("FILL_RATIO"), (int,float)):
        c = replace(c, FILL_RATIO=float(j["FILL_RATIO"]))
    if isinstance(j.get("EDGE_MARGIN_RATIO"), (int,float)):
        c = replace(c, EDGE_MARGIN_RATIO=float(j["EDGE_MARGIN_RATIO"]))
    if isinstance(j.get("MAX_TOOL_BUTTONS"), int):
        c = replace(c, MAX_TOOL_BUTTONS=int(j["MAX_TOOL_BUTTONS"]))
    if isinstance(j.get("BUTTON_COLUMNS"), int):
        c = replace(c, BUTTON_COLUMNS=int(j["BUTTON_COLUMNS"]))
    if isinstance(j.get("GUI_START_QUADRANT"), str):
        c = replace(c, GUI_START_QUADRANT=str(j["GUI_START_QUADRANT"]))

    bp = j.get("BASELINE_PATH")
    if isinstance(bp, str) and bp:
        c = replace(c, BASELINE_PATH=resolve_path(c, bp, "ps1"))
    else:
        c = replace(c, BASELINE_PATH=resolve_path(c, "Baseline.ps1", "ps1"))
    if isinstance(j.get("BASELINE_DELAY_MS"), int):
        c = replace(c, BASELINE_DELAY_MS=int(j["BASELINE_DELAY_MS"]))

    dc = j.get("DEFAULT_CLUSTER")
    if isinstance(dc, str) and dc:
        c = replace(c, DEFAULT_CLUSTER=dc.strip())

    return c

def resolve_path(cfg: Config, path: str, path_type: str = "auto") -> str:
    """
    Zentrale Pfadauflösungsfunktion für das gesamte System.
    
    Args:
        cfg: Konfigurationsobjekt
        path: Der aufzulösende Pfad (kann relativ oder absolut sein)
        path_type: "ps1", "cmd", "exe", "url", "auto" (erkennt automatisch an Dateiendung)
    
    Returns:
        Aufgelöster absoluter Pfad
        
    Beispiele:
        resolve_path(cfg, "CMD.ps1", "ps1") -> "C:\\Support\\v81\\scripts\\CMD.ps1"
        resolve_path(cfg, "./scripts/Baseline.ps1", "ps1") -> "C:\\Support\\v81\\scripts\\Baseline.ps1"
        resolve_path(cfg, "C:\\Windows\\explorer.exe", "exe") -> "C:\\Windows\\explorer.exe"
    """
    if not path:
        return path
    
    # Environment-Variablen expandieren
    expanded_path = _expand_env(path)
    
    # URL-Pfade nicht auflösen
    if path_type == "url" or not expanded_path:
        return expanded_path
    
    # Bereits absolute Pfade direkt zurückgeben
    if _is_abs(expanded_path):
        return os.path.abspath(expanded_path)
    
    # Automatische Erkennung des Pfadtyps
    if path_type == "auto":
        if expanded_path.lower().endswith((".ps1", ".cmd")):
            path_type = "ps1" if expanded_path.lower().endswith(".ps1") else "cmd"
        elif expanded_path.lower().endswith((".exe", ".msc")):
            path_type = "exe"
        else:
            path_type = "exe"  # Default für unbekannte Typen
    
    # Basisverzeichnis bestimmen
    if path_type in ("ps1", "cmd"):
        base_dir = cfg.SCRIPTS_DIR
    else:
        base_dir = cfg.TOOLS_DIR
    
    # UNC-kompatible Pfadauflösung für relative Basisverzeichnisse
    if base_dir.startswith("./"):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        base_dir = os.path.join(script_dir, base_dir[2:])
    
    # Entferne "./" vom Anfang des Pfads, falls vorhanden
    if expanded_path.startswith("./"):
        expanded_path = expanded_path[2:]
    
    # Pfade kombinieren und absolut machen
    resolved_path = os.path.join(base_dir, expanded_path)
    return os.path.abspath(resolved_path)

def _resolve_tool_path(cfg: Config, ttype: str, path: str) -> str:
    p = _expand_env(path or "")
    if ttype == "url" or not p:
        return p
    if _is_abs(p):
        return p
    # Für ps1/cmd: Normale Auflösung mit SCRIPTS_DIR
    if ttype in ("ps1","cmd"):
        base = cfg.SCRIPTS_DIR
        candidate = os.path.join(base, p)
        return os.path.abspath(candidate) if os.path.exists(candidate) else p
    # Für andere Dateien: Normale Auflösung
    base = cfg.TOOLS_DIR
    candidate = os.path.join(base, p)
    return candidate if os.path.exists(candidate) else p

def _split_args(argval: Any) -> List[str]:
    if argval is None: return []
    if isinstance(argval, list):
        return [str(x) for x in argval]
    if isinstance(argval, str):
        s = argval.strip()
        if not s: return []
        try:
            return shlex.split(s)
        except Exception:
            return [s]
    return [str(argval)]

def extract_menus(j: dict, cfg: Config) -> Dict[str, List[Dict[str,Any]]]:
    menus: Dict[str, List[Dict[str,Any]]] = {}

    if isinstance(j.get("TOOLS"), dict):
        for mname, items in j["TOOLS"].items():
            if not isinstance(items, list): continue
            out: List[Dict[str,Any]] = []
            for it in items:
                if not isinstance(it, dict): continue
                lab = it.get("label") or it.get("name")
                ttype = (it.get("type") or "").strip().lower()
                raw_path = it.get("path") or it.get("Path") or ""
                if not lab or not raw_path: continue
                p = _resolve_tool_path(cfg, ttype, raw_path)
                entry = {"label": lab, "path": p, "type": ttype}
                if "elevate" in it: entry["uac"] = bool(it.get("elevate", False))
                if isinstance(it.get("browser"), str):
                    entry["browser"] = it["browser"].strip().lower()
                entry["args"] = _split_args(it.get("args"))
                out.append(entry)
            menus[mname] = out
    return menus

def _ts(): return time.strftime("%H:%M:%S")

# Log levels for console filtering
class LogLevel:
    DEBUG = 0
    INFO = 1
    WARNING = 2
    ERROR = 3
    
    @staticmethod
    def from_string(level_str: str) -> int:
        return {
            "DEBUG": LogLevel.DEBUG,
            "INFO": LogLevel.INFO,
            "WARNING": LogLevel.WARNING,
            "ERROR": LogLevel.ERROR
        }.get(level_str.upper(), LogLevel.INFO)

class Logger:
    def __init__(self, local_dir: str, central_dir: Optional[str]):
        os.makedirs(local_dir, exist_ok=True)
        self.user=os.environ.get("USERNAME","user"); self.host=os.environ.get("COMPUTERNAME","host")
        ts=datetime.now().strftime("%Y%m%d_%H%M%S")
        base=f"cockpit_{self.user}_{self.host}_{ts}"
        self.txt=os.path.join(local_dir, base+".log")

    def text(self, msg: str):
        line=f"[{_ts()}] {msg}"
        try:
            with open(self.txt,"a",encoding="utf-8") as f: f.write(line+"\n")
        except Exception: pass

def primary_rect_logical() -> QRect:
    return QGuiApplication.primaryScreen().availableGeometry()

def get_quadrant_for_config(quad: str, cfg: Config) -> Tuple[int,int,int,int]:  # type: ignore
    return get_quadrant_physical(quad, cfg.FILL_RATIO, cfg.EDGE_MARGIN_RATIO)

# Start helpers
ERROR_ELEVATION_REQUIRED = 740
SEE_MASK_NOCLOSEPROCESS = 0x00000040
CREATE_NEW_CONSOLE = subprocess.CREATE_NEW_CONSOLE
CREATE_NO_WINDOW = 0x08000000

class SHELLEXECUTEINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD), ("fMask", wintypes.ULONG), ("hwnd", wintypes.HWND),
        ("lpVerb", wintypes.LPCWSTR), ("lpFile", wintypes.LPCWSTR), ("lpParameters", wintypes.LPCWSTR),
        ("lpDirectory", wintypes.LPCWSTR), ("nShow", ctypes.c_int), ("hInstApp", wintypes.HINSTANCE),
        ("lpIDList", wintypes.LPVOID), ("lpClass", wintypes.LPCWSTR), ("hkeyClass", wintypes.HKEY),
        ("dwHotKey", wintypes.DWORD), ("hIcon", wintypes.HANDLE), ("hProcess", wintypes.HANDLE),
    ]

def _shell_run(verb:str, file:str, params:str, cwd:Optional[str]=None):
    sei=SHELLEXECUTEINFO(); sei.cbSize=ctypes.sizeof(SHELLEXECUTEINFO)
    sei.fMask=SEE_MASK_NOCLOSEPROCESS; sei.hwnd=None; sei.lpVerb=verb
    sei.lpFile=file; sei.lpParameters=params; sei.lpDirectory=cwd or None
    sei.nShow=win32con.SW_SHOWNORMAL
    ok=ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(sei))
    if not ok: return False, None, f"ShellExecuteExW failed ({verb})"
    pid=None
    if sei.hProcess:
        pid=int(ctypes.windll.kernel32.GetProcessId(sei.hProcess)) or None
        ctypes.windll.kernel32.CloseHandle(sei.hProcess)
    return True, pid, None

def try_start_exe(path: str, args: List[str]|None=None, cwd: Optional[str]=None):
    args = args or []
    try:
        p = subprocess.Popen([path]+args, cwd=cwd)
        return True, p.pid, False, None
    except OSError as e:
        if getattr(e,'winerror',None)==ERROR_ELEVATION_REQUIRED:
            ok,pid,err=_shell_run("runas", path, " ".join(args))
            return (ok, pid, True, err if not ok else None)
        return False, None, False, str(e)
    except Exception as e:
        return False, None, False, str(e)

def try_start_exe_uac(path:str, args:List[str]|None=None, cwd:Optional[str]=None):
    args=args or []
    ok,pid,err=_shell_run("runas", path, " ".join(args), cwd)
    return (ok, pid, True, err if not ok else None)

def which(*names: str) -> str:
    for n in names:
        p = shutil.which(n)
        if p: return p
    return ""

def build_cmd_command(title:str, run_after:Optional[str]=None, keep_open:bool=True) -> List[str]:
    args = ["cmd.exe", "/K" if keep_open else "/C", f"title {title}"]
    if run_after: args[-1] = args[-1] + f" & {run_after}"
    return args

def build_ps_command(script_path: Optional[str], title: str,
                     extra_args: Optional[List[str]] = None,
                     keep_open: bool = True) -> List[str]:
    extra_args = extra_args or []
    ps = which("pwsh.exe", "powershell.exe") or r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    if script_path:
        cmd = f"[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; [Console]::InputEncoding=[System.Text.Encoding]::UTF8; chcp 65001 > $null; $Host.UI.RawUI.WindowTitle='{title}'; & '{script_path}' " + " ".join(map(shlex.quote, extra_args))
    else:
        cmd = f"[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; [Console]::InputEncoding=[System.Text.Encoding]::UTF8; chcp 65001 > $null; $Host.UI.RawUI.WindowTitle='{title}'"
    args = [ps, "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass"]
    args += (["-NoExit"] if keep_open else [])
    args += ["-Command", cmd]
    return args

def spawn_console(args: List[str], cwd: Optional[str]=None) -> subprocess.Popen:
    return subprocess.Popen(args, creationflags=CREATE_NEW_CONSOLE, cwd=cwd)

class AppState:
    def __init__(self, cfg: Config, menus: Dict[str, List[Dict[str,Any]]], start_menu: str):
        self.cfg = cfg
        os.makedirs(cfg.LOCAL_LOG_DIR, exist_ok=True)
        self.logger = Logger(cfg.LOCAL_LOG_DIR, cfg.CENTRAL_LOG_DIR)
        self.menus = menus
        self.menu_name = start_menu
        self.rr_index=0
    
    def next_quad(self) -> str:
        q = self.cfg.RR_ORDER[self.rr_index]
        self.rr_index = (self.rr_index+1) % len(self.cfg.RR_ORDER)
        return q
    
    def tools_for_current(self) -> List[Dict[str,Any]]:
        return list(self.menus.get(self.menu_name, []))

def _btn_style() -> str:
    return (
        "QToolButton{background:#2b3036;color:#dbe6ef;border:1px solid #3a4046;"
        "border-radius:6px;padding:6px 10px;}"
        "QToolButton:hover{background:#343a41;}"
    )

class NeonButton(QToolButton):
    def __init__(self, text: str, square: Optional[int] = None):
        super().__init__(text=text)
        self.setCursor(Qt.PointingHandCursor)
        s = square or 0
        if s > 0:
            self.setFixedSize(s, s)
        else:
            self.setMinimumSize(132, 40)
        self._install_futuristic_effects()
        self.setStyleSheet("""
        QToolButton {
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:1,
                stop:0 #1a1a2e, stop:0.3 #16213e, stop:0.7 #0f3460, stop:1 #1a1a2e
            );
            color: #e8f4fd;
            border: 1px solid qlineargradient(
                x1:0, y1:0, x2:1, y2:0,
                stop:0 rgba(74, 144, 226, 120), stop:0.5 rgba(123, 179, 240, 100), stop:1 rgba(74, 144, 226, 120)
            );
            border-radius: 12px;
            padding: 8px 12px;
            font-weight: 600;
        }
        QToolButton:hover {
            border: 1px solid qlineargradient(
                x1:0, y1:0, x2:1, y2:0,
                stop:0 rgba(123, 179, 240, 150), stop:0.5 rgba(168, 200, 248, 130), stop:1 rgba(123, 179, 240, 150)
            );
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:1,
                stop:0 #2a2a3e, stop:0.3 #26314e, stop:0.7 #1f4470, stop:1 #2a2a3e
            );
            color: #f0f8ff;
        }
        QToolButton:pressed {
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:1,
                stop:0 #0f1419, stop:0.5 #1a1a2e, stop:1 #0f1419
            );
            border: 1px solid rgba(74, 144, 226, 100);
            color: #c8e6f5;
        }
        """)
        
        # Force style refresh
        self.style().unpolish(self)
        self.style().polish(self)
        self.update()

    def _install_futuristic_effects(self):
        """Install holographic glow and shadow effects with animation support."""
        # Primary glow effect (animatable)
        self.glow = QGraphicsDropShadowEffect(self)
        self.glow.setBlurRadius(20)
        self.glow.setOffset(0, 0)
        self.glow.setColor(QColor(74, 144, 226, 100))  # Chrome blue with transparency
        self.setGraphicsEffect(self.glow)
        
        # Secondary outer glow for depth (animatable)
        self.outer_glow = QGraphicsDropShadowEffect(self)
        self.outer_glow.setBlurRadius(35)
        self.outer_glow.setOffset(0, 0)
        self.outer_glow.setColor(QColor(123, 179, 240, 50))  # Light blue outer glow
        self.setGraphicsEffect(self.outer_glow)
        
        # Store base values for animation
        self._base_glow_intensity = 100
        self._base_outer_intensity = 50

    def _update_glow(self, pulse_intensity):
        """Update glow effects with pulsing animation - CPU efficient."""
        try:
            # Calculate animated intensities
            glow_alpha = int(self._base_glow_intensity * (0.8 + 0.4 * pulse_intensity))
            outer_alpha = int(self._base_outer_intensity * (0.6 + 0.8 * pulse_intensity))
            
            # Update glow colors
            self.glow.setColor(QColor(74, 144, 226, min(255, glow_alpha)))
            self.outer_glow.setColor(QColor(123, 179, 240, min(255, outer_alpha)))
            
        except Exception:
            # Silent fail for VDI compatibility
            pass

    def _init_hotkeys(self):
        """Initialize hotkey system for window management."""
        # Track last active window for round-robin placement
        self._last_active_window = None
        self._last_active_pid = None
        
        # Enable key events
        self.setFocusPolicy(Qt.StrongFocus)
        
    def keyPressEvent(self, event):
        """Handle hotkey presses for window management."""
        try:
            # Check for Ctrl+Number combinations
            if event.modifiers() == Qt.ControlModifier:
                if event.key() == Qt.Key_0:
                    self._center_gui()
                    event.accept()
                    return
                elif event.key() == Qt.Key_1:
                    self._place_last_window("TL")
                    event.accept()
                    return
                elif event.key() == Qt.Key_2:
                    self._place_last_window("TR")
                    event.accept()
                    return
                elif event.key() == Qt.Key_3:
                    self._place_last_window("BL")
                    event.accept()
                    return
                    
        except Exception as e:
            self._log(f"Hotkey error: {e}")
            
        # Call parent keyPressEvent for other keys
        super().keyPressEvent(event)

    def _center_gui(self):
        """Center the GUI window in its configured start quadrant."""
        try:
            # Get primary screen geometry
            screen = QApplication.primaryScreen()
            screen_geometry = screen.availableGeometry()
            
            # Get configured start quadrant (default: TR)
            quadrant = getattr(self.cfg, 'GUI_START_QUADRANT', 'TR')
            
            # Calculate quadrant dimensions with DPI scaling
            margin = int(screen_geometry.width() * self.cfg.EDGE_MARGIN_RATIO)
            
            # Get DPI scaling factor
            dpi_scale = screen.devicePixelRatio()
            logical_dpi = screen.logicalDotsPerInch()
            scale_factor = logical_dpi / 96.0  # 96 DPI is standard
            
            # Calculate scaled dimensions based on screen size
            if screen_geometry.width() <= 1920:  # 1080p or smaller
                w = int(screen_geometry.width() * 0.6) - margin  # 60% for smaller screens
                h = int(screen_geometry.height() * 0.7) - margin  # 70% height for better visibility
            else:  # 4K or larger
                w = int(screen_geometry.width() * 0.4) - margin  # 40% for larger screens
                h = int(screen_geometry.height() * 0.5) - margin  # 50% height
                
            # Apply minimum and maximum size constraints
            min_width = int(600 * scale_factor)
            min_height = int(400 * scale_factor)
            max_width = int(screen_geometry.width() * 0.8)
            max_height = int(screen_geometry.height() * 0.8)
            
            w = max(min_width, min(w, max_width))
            h = max(min_height, min(h, max_height))
            
            # Calculate position based on quadrant
            if quadrant == "TL":  # Top-Left
                x = margin
                y = margin
            elif quadrant == "TR":  # Top-Right
                x = int(screen_geometry.width() * 0.5) + margin
                y = margin
            elif quadrant == "BL":  # Bottom-Left
                x = margin
                y = int(screen_geometry.height() * 0.5) + margin
            elif quadrant == "BR":  # Bottom-Right
                x = int(screen_geometry.width() * 0.5) + margin
                y = int(screen_geometry.height() * 0.5) + margin
            else:  # Fallback to TR
                x = int(screen_geometry.width() * 0.5) + margin
                y = margin
                quadrant = "TR"
            
            # Center within the quadrant
            center_x = x + (w - self.width()) // 2
            center_y = y + (h - self.height()) // 2
            
            # Move to quadrant center
            self.move(center_x, center_y)
            self._log(f"GUI centered in {quadrant} quadrant at ({center_x}, {center_y})")
            
        except Exception as e:
            self._log(f"Center GUI error: {e}")

    def _place_last_window(self, quadrant):
        """Place the last active window in the specified quadrant."""
        try:
            if not self._last_active_window:
                self.log("No last active window to place")
                return
                
            # Get window handle and PID
            hwnd = self._last_active_window
            pid = self._last_active_pid
            
            if not hwnd or not win32gui.IsWindow(hwnd):
                self.log("Last active window no longer valid")
                self._last_active_window = None
                self._last_active_pid = None
                return
                
            # Place window using existing logic
            if quadrant == "TL":
                self._place_window_tl(hwnd, pid)
            elif quadrant == "TR":
                self._place_window_tr(hwnd, pid)
            elif quadrant == "BL":
                self._place_window_bl(hwnd, pid)
            elif quadrant == "BR":
                self._place_window_br(hwnd, pid)
                
            self.log(f"Last active window placed in {quadrant}")
            
        except Exception as e:
            self.log(f"Place last window error: {e}")

    def _update_last_active_window(self, hwnd, pid):
        """Update the last active window for hotkey placement."""
        try:
            if hwnd and win32gui.IsWindow(hwnd):
                self._last_active_window = hwnd
                self._last_active_pid = pid
                self.log(f"Last active window updated: hwnd={hwnd}, pid={pid}")
        except Exception:
            pass

    def _place_window_tl(self, hwnd, pid):
        """Place window in Top-Left quadrant."""
        try:
            # Get screen dimensions
            screen_width = QApplication.primaryScreen().availableGeometry().width()
            screen_height = QApplication.primaryScreen().availableGeometry().height()
            
            # Calculate TL quadrant dimensions
            x = int(screen_width * 0.01)  # Edge margin
            y = int(screen_height * 0.01)  # Edge margin
            w = int(screen_width * 0.48)   # Half width minus margin
            h = int(screen_height * 0.48)  # Half height minus margin
            
            # Place window
            win32gui.SetWindowPos(hwnd, 0, x, y, w, h, win32con.SWP_SHOWWINDOW)
            self.log(f"Window placed TL: ({x}, {y}, {w}, {h})")
            
        except Exception as e:
            self.log(f"TL placement error: {e}")

    def _place_window_tr(self, hwnd, pid):
        """Place window in Top-Right quadrant."""
        try:
            # Get screen dimensions
            screen_width = QApplication.primaryScreen().availableGeometry().width()
            screen_height = QApplication.primaryScreen().availableGeometry().height()
            
            # Calculate TR quadrant dimensions
            x = int(screen_width * 0.51)  # Right half
            y = int(screen_height * 0.01)  # Edge margin
            w = int(screen_width * 0.48)   # Half width minus margin
            h = int(screen_height * 0.48)  # Half height minus margin
            
            # Place window
            win32gui.SetWindowPos(hwnd, 0, x, y, w, h, win32con.SWP_SHOWWINDOW)
            self.log(f"Window placed TR: ({x}, {y}, {w}, {h})")
            
        except Exception as e:
            self.log(f"TR placement error: {e}")

    def _place_window_bl(self, hwnd, pid):
        """Place window in Bottom-Left quadrant."""
        try:
            # Get screen dimensions
            screen_width = QApplication.primaryScreen().availableGeometry().width()
            screen_height = QApplication.primaryScreen().availableGeometry().height()
            
            # Calculate BL quadrant dimensions
            x = int(screen_width * 0.01)  # Edge margin
            y = int(screen_height * 0.51)  # Bottom half
            w = int(screen_width * 0.48)   # Half width minus margin
            h = int(screen_height * 0.48)  # Half height minus margin
            
            # Place window
            win32gui.SetWindowPos(hwnd, 0, x, y, w, h, win32con.SWP_SHOWWINDOW)
            self.log(f"Window placed BL: ({x}, {y}, {w}, {h})")
            
        except Exception as e:
            self.log(f"BL placement error: {e}")

    def _place_window_br(self, hwnd, pid):
        """Place window in Bottom-Right quadrant."""
        try:
            # Get screen dimensions
            screen_width = QApplication.primaryScreen().availableGeometry().width()
            screen_height = QApplication.primaryScreen().availableGeometry().height()
            
            # Calculate BR quadrant dimensions
            x = int(screen_width * 0.51)  # Right half
            y = int(screen_height * 0.51)  # Bottom half
            w = int(screen_width * 0.48)   # Half width minus margin
            h = int(screen_height * 0.48)  # Half height minus margin
            
            # Place window
            win32gui.SetWindowPos(hwnd, 0, x, y, w, h, win32con.SWP_SHOWWINDOW)
            self.log(f"Window placed BR: ({x}, {y}, {w}, {h})")
            
        except Exception as e:
            self.log(f"BR placement error: {e}")

class CommandCenter(QWidget):
    def __init__(self, state: AppState, parent=None):
        QWidget.__init__(self, parent)

        self.state = state
        self.cfg   = state.cfg

        self.setWindowTitle(f"Support Cockpit — {self.state.menu_name}")
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Window | Qt.WindowMaximizeButtonHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setWindowOpacity(0.75)
        
        # Set minimum and maximum size constraints
        self.setMinimumSize(600, 400)
        self.setMaximumSize(1600, 1200)

        self._dragging = False
        self._drag_off = QPoint()

        # Remember original geometry for maximize toggle
        self._original_geometry = None
        self._is_maximized = False
        self._q_out: Optional[Queue] = None
        self._q_err: Optional[Queue] = None
        self._baseline_timer: Optional[QTimer] = None

        # Layout
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(10)

        # Chrome container
        self.chrome = QWidget(self)
        self.chrome.setObjectName("chrome")
        self._apply_chrome_style(self.chrome, border_boost=0.0)
        chrome_lay = QVBoxLayout(self.chrome)
        chrome_lay.setContentsMargins(14, 14, 14, 14)
        chrome_lay.setSpacing(10)
        lay.addWidget(self.chrome)
        
        # Force style refresh
        self.style().unpolish(self)
        self.style().polish(self)
        self.update()

        # Add outer shadow to chrome
        outer_shadow = QGraphicsDropShadowEffect(self.chrome)
        outer_shadow.setBlurRadius(24)
        outer_shadow.setOffset(0, 6)
        outer_shadow.setColor(Qt.black)
        self.chrome.setGraphicsEffect(outer_shadow)

        # Initialize animation system
        self._init_animations()
        
        # Initialize hotkey system (track last active window)
        self._last_active_window = None
        self._last_active_pid = None
        self.setFocusPolicy(Qt.StrongFocus)

        # Header row
        head = QHBoxLayout()
        head.setSpacing(8)
        chrome_lay.addLayout(head)

        # Optional logo
        logo_lbl = self._maybe_logo_label()
        if logo_lbl:
            head.addWidget(logo_lbl)

        # Title
        self.header = QLabel(f"Support Cockpit — {self.state.menu_name}")
        self.header.setObjectName("hdr")
        self.header.setCursor(Qt.SizeAllCursor)
        self._apply_header_style(self.header)

        # Menu combo
        self.menu_combo = QComboBox()
        self.menu_combo.setObjectName("menucombo")
        self._apply_combo_style(self.menu_combo)
        self._rebuild_menu_combo()
        self.menu_combo.currentTextChanged.connect(self._switch_menu)

        # Window action buttons (compact square)
        squ = 28
        btn_copy = NeonButton("Copy", square=squ)  # copy console
        btn_min  = NeonButton("➖", square=squ)
        self.btn_max = NeonButton("⬜", square=squ)  # toggles to "⬛"
        btn_close = NeonButton("❌", square=squ)

        btn_copy.setToolTip("Copy console to clipboard")
        btn_min.setToolTip("Minimize")
        self.btn_max.setToolTip("Maximize / Restore")
        btn_close.setToolTip("Close")

        btn_copy.clicked.connect(self._copy_console)
        btn_min.clicked.connect(self.showMinimized)
        self.btn_max.clicked.connect(self._toggle_maximize)
        btn_close.clicked.connect(self.close)

        # Assemble header
        head.addWidget(self.header)
        head.addSpacing(6)
        lbl = QLabel("Menu:"); lbl.setObjectName("hdrsub"); head.addWidget(lbl)
        head.addWidget(self.menu_combo, 0)
        head.addStretch(1)
        head.addWidget(btn_copy)
        head.addWidget(btn_min)
        head.addWidget(self.btn_max)
        head.addWidget(btn_close)
        
        # Force style refresh for header elements
        self.header.style().unpolish(self.header)
        self.header.style().polish(self.header)
        self.header.update()
        lbl.style().unpolish(lbl)
        lbl.style().polish(lbl)
        lbl.update()
        self.menu_combo.style().unpolish(self.menu_combo)
        self.menu_combo.style().polish(self.menu_combo)
        self.menu_combo.update()

        # Accent line - removed for cleaner look

        # Splitter
        self.splitter = QSplitter(Qt.Vertical)
        self.splitter.setObjectName("split")
        self.splitter.setHandleWidth(1)  # Make handle very thin
        chrome_lay.addWidget(self.splitter, 1)

        self._build_log()
        self._build_buttons()
        self.place_command_center()

        self.timer = QTimer(self)
        self.timer.timeout.connect(lambda: None)
        self.timer.start(200)

        self._log("Grid60.03 gestartet (clean final)", "INFO")

        baseline_path = getattr(self.cfg, "BASELINE_PATH", None)
        baseline_delay = int(getattr(self.cfg, "BASELINE_DELAY_MS", 0) or 0)

        if baseline_path and os.path.isfile(baseline_path):
            delay = max(0, baseline_delay)
            QTimer.singleShot(delay, self._start_baseline)
            self._log(f"BASELINE geplant in {delay} ms → {baseline_path}", "INFO")

    # ===== NEW FIXED BUTTON SYSTEM =====
    def _build_buttons(self):
        """Build button grid with dynamic column count based on window width."""
        w = QWidget()
        g = QGridLayout(w)
        g.setSpacing(8)
        self.btns = {}
        
        items = self.state.tools_for_current()[:max(0, self.cfg.MAX_TOOL_BUTTONS)]
        self._log(f"★ BUILDING {len(items)} buttons", "DEBUG")
        
        # Calculate dynamic columns based on window width
        # Default to 4 columns, but adjust based on available width
        window_width = self.width() if hasattr(self, 'width') else 640
        button_width = 140  # Approximate button width including spacing
        max_cols = max(2, min(6, window_width // button_width))  # 2-6 columns
        
        # Use config override if available, otherwise use dynamic calculation
        cols = getattr(self.cfg, 'BUTTON_COLUMNS', max_cols)
        if cols <= 0:  # If config is 0 or negative, use dynamic
            cols = max_cols
        
        self._log(f"★ BUTTON GRID: {cols} columns (window_width={window_width})", "DEBUG")
        
        r = c = 0
        for i, tool in enumerate(items):
            label = tool["label"]
            path = tool["path"] 
            tool_type = tool.get("type", "")
            uac = bool(tool.get("uac", False))
            args = tool.get("args", [])
            
            self._log(f"★ BTN[{i}]: '{label}' type='{tool_type}' path='{path}' uac={uac}", "DEBUG")
            
            # Add UAC symbol if elevated
            if uac:
                label += " 🛡️"
            b = NeonButton(text=label)
            key = f"{r}:{c}"
            
            # Create proper closure
            b.clicked.connect(self._make_button_handler(tool))
            
            g.addWidget(b, r, c)
            self.btns[key] = b
            c += 1
            if c >= cols: 
                c = 0
                r += 1
        
        self._log(f"★ BUTTONS CREATED: {len(self.btns)} in {cols} columns", "DEBUG")
        
        # Force style refresh for buttons
        for btn in self.btns.values():
            btn.style().unpolish(btn)
            btn.style().polish(btn)
            btn.update()
        
        if self.splitter.count() > 0:
            self.splitter.insertWidget(0, w)
            old = self.splitter.widget(1)
            if isinstance(old, QWidget) and old is not self.console:
                old.setParent(None)
        else:
            self.splitter.addWidget(w)

    def _make_button_handler(self, tool):
        """Create button click handler with proper closure."""
        def handler():
            self._handle_tool_launch(tool)
        return handler

    def _handle_tool_launch(self, tool):
        """Handle tool launch based on JSON type with RDSH support."""
        label = tool["label"]
        path = tool["path"]
        tool_type = tool.get("type", "")
        uac = bool(tool.get("uac", False))
        args = tool.get("args", [])
        quad = self.state.next_quad()
        
        # RDSH-specific parameters
        is_rds_session = bool(tool.get("rds_session", False))
        session_detection = bool(tool.get("session_detection", False))
        window_placement = tool.get("window_placement", {})
        
        self._log(f"★ LAUNCH: '{label}' type='{tool_type}' path='{path}' uac={uac} rds={is_rds_session} → {quad}", "INFO")
        
        try:
            if tool_type == "ps1":
                self._launch_powershell(path, label, args, uac, quad)
            elif tool_type == "cmd":
                self._launch_cmd(label, uac, quad)
            elif tool_type == "url":
                self._launch_url_isolated(path, quad)
            elif tool_type == "exe" or "explorer" in path.lower():
                if "explorer" in path.lower():
                    self._launch_explorer_fixed(label, path, args, quad)
                elif is_rds_session:
                    self._launch_rds_application(path, label, args, uac, quad, window_placement)
                else:
                    self._launch_executable(path, label, args, uac, quad)
            else:
                # Fallback detection
                if path.lower().endswith(".ps1"):
                    self._launch_powershell(path, label, args, uac, quad)
                elif path.lower() == "cmd:blank":
                    self._launch_cmd(label, uac, quad)
                elif path.lower() == "ps:blank":
                    self._launch_powershell_blank(label, uac, quad)
                elif path.lower().startswith("http"):
                    self._launch_url_isolated(path, quad)
                else:
                    self._launch_executable(path, label, args, uac, quad)
        except Exception as e:
            self._log(f"★ LAUNCH ERROR: {e}", "ERROR")

    def _launch_powershell(self, script_path: str, label: str, args: List[str], uac: bool, quad: str):
        """Launch PowerShell script with RDSH-aware placement."""
        if script_path == "ps:blank":
            self._launch_powershell_blank(label, uac, quad)
            return
            
        # Der Pfad wurde bereits von _resolve_tool_path aufgelöst
        # Stelle nur sicher, dass der Pfad absolut ist
        script_path = os.path.abspath(script_path)
        
        if not os.path.isfile(script_path):
            self._log(f"PS1 script not found: {script_path}")
            return
            
        # RDSH-aware title with session info
        session_id = self._get_current_session_id()
        title = f"{label} – {os.environ.get('USERNAME','User')}@{os.environ.get('COMPUTERNAME','Host')} [S{session_id}]"
        
        if uac:
            # Elevated PowerShell
            ps = which("pwsh.exe", "powershell.exe") or r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
            script_args = " ".join(map(shlex.quote, args)) if args else ""
            cmd = (
                f'-NoLogo -NoProfile -ExecutionPolicy Bypass -NoExit -Command '
                f'"$Host.UI.RawUI.WindowTitle=\'{title}\'; & \'{script_path}\' {script_args}"'
            )
            ok, pid, err = _shell_run("runas", ps, cmd, cwd=os.path.dirname(script_path))
            if ok:
                self._log(f"PS script elevated started: '{title}' → {quad} (Session {session_id})")
                # Enhanced PowerShell placement for elevated
                self._schedule_powershell_placement(pid, quad, label, title)
            else:
                self._log(f"PS script elevated failed: {err}")
        else:
            # Normal PowerShell
            ps_args = build_ps_command(script_path, title, extra_args=args, keep_open=True)
            p = spawn_console(ps_args, cwd=os.path.dirname(script_path))
            self._log(f"PS script started: '{title}' → {quad} (Session {session_id})")
            # Enhanced PowerShell placement
            self._schedule_powershell_placement(p.pid, quad, label, title)

    def _launch_powershell_blank(self, label: str, uac: bool, quad: str):
        """Launch blank PowerShell console with basic placement."""
        if uac:
            ps = which("pwsh.exe", "powershell.exe") or r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
            cmd = f'-NoLogo -NoProfile -ExecutionPolicy Bypass -NoExit -Command "$Host.UI.RawUI.WindowTitle=\'{label}\'"'
            ok, pid, err = _shell_run("runas", ps, cmd)
            if ok:
                self._log(f"PS elevated started: '{label}' → {quad}")
                # Enhanced PowerShell placement for elevated
                self._schedule_powershell_placement(pid, quad, label, label)
            else:
                self._log(f"PS elevated failed: {err}")
        else:
            args = build_ps_command(None, label, keep_open=True)
            p = spawn_console(args)
            self._log(f"PS started: '{label}' → {quad}")
            # Only try placement for non-elevated
            self._schedule_window_placement_by_pid(p.pid, quad, label)

    def _launch_cmd(self, label: str, uac: bool, quad: str):
        """Launch CMD console with basic placement."""
        if uac:
            ok, pid, err = _shell_run("runas", "cmd.exe", f'/K "title {label}"')
            if ok:
                self._log(f"CMD elevated started: '{label}' → {quad}")
            else:
                self._log(f"CMD elevated failed: {err}")
        else:
            args = build_cmd_command(label, keep_open=True)
            p = spawn_console(args)
            self._log(f"CMD started: '{label}' → {quad}")
            # Only try placement for non-elevated
            self._schedule_window_placement_by_pid(p.pid, quad, label)

    def _launch_executable(self, path: str, label: str, args: List[str], uac: bool, quad: str):
        """Launch executable with window placement."""
        # Special case: Settings URI
        if path.startswith("ms-settings:"):
            if os.name == "nt":
                os.startfile(path)
                self._log(f"Settings URI launched: {path}")
            return
            
        if not os.path.exists(path) and not shutil.which(path):
            self._log(f"Executable not found: {path}")
            return
        
        # Special case: Browser with snapshot-based placement
        exe_name = os.path.basename(path).lower()
        if exe_name in ("msedge.exe", "chrome.exe", "firefox.exe") and any(arg for arg in args if "user-data-dir" in str(arg)):
            self._launch_browser_with_snapshot(path, label, args, uac, quad)
            return
        
        # Special case: MMC-based tools with snapshot (to handle single-instance behavior)
        if exe_name in ("eventvwr.exe", "mmc.exe"):
            self._launch_mmc_with_snapshot(path, label, args, uac, quad)
            return
            
        try:
            if uac:
                ok, pid, elevated, err = try_start_exe_uac(path, args=args)
            else:
                ok, pid, elevated, err = try_start_exe(path, args=args)

            if ok and pid:
                self._log(f"{'EXE elevated' if uac else 'EXE'} started (PID={pid}) → {quad}")
                # Use simple PID-based placement for non-browsers
                self._schedule_window_placement_by_pid(pid, quad, label, is_browser=False)
            elif ok:
                self._log(f"{'EXE elevated' if uac else 'EXE'} started (no PID) → {quad}")
                self._schedule_window_placement_generic(path, quad, label)
            else:
                self._log(f"EXE start failed: {label} | {err}")
        except Exception as e:
            self._log(f"EXE launch exception: {e}")

    def _launch_browser_with_snapshot(self, path: str, label: str, args: List[str], uac: bool, quad: str):
        """Launch isolated browser with snapshot-based window placement."""
        exe_name = os.path.basename(path).lower()
        
        # Take snapshot of current browser windows BEFORE launch
        before_windows = find_windows_by_criteria(
            image_names=[exe_name],
            class_names=["Chrome_WidgetWin_1", "MozillaWindowClass", "ApplicationFrameWindow"]
        )
        
        try:
            if uac:
                ok, pid, elevated, err = try_start_exe_uac(path, args=args)
            else:
                ok, pid, elevated, err = try_start_exe(path, args=args)

            if ok:
                self._log(f"{'Browser elevated' if uac else 'Browser'} started → {quad}")
                # Schedule snapshot-based placement
                self._schedule_browser_placement_by_snapshot(before_windows, exe_name, quad, label)
            else:
                self._log(f"Browser start failed: {label} | {err}")
        except Exception as e:
            self._log(f"Browser launch exception: {e}")

    def _schedule_browser_placement_by_snapshot(self, before_windows: List[int], exe_name: str, quad: str, label: str):
        """Place browser window using before/after snapshot comparison."""
        def place_by_snapshot(attempt=0):
            attempt += 1
            
            # Get current browser windows
            current_windows = find_windows_by_criteria(
                image_names=[exe_name],
                class_names=["Chrome_WidgetWin_1", "MozillaWindowClass", "ApplicationFrameWindow"]
            )
            
            # Find new windows (not in before snapshot)
            new_windows = [h for h in current_windows if h not in before_windows]
            
            if new_windows:
                # Take the largest new window
                target = max(new_windows, key=get_window_area)
                x, y, w, h = get_quadrant_for_config(quad, self.cfg)
                ok = move_window(target, x, y, w, h)
                if ok:
                    self._log(f"Browser placed: {label} → {quad} (hwnd={target}, new window)")
                    
                    # Update last active window for hotkey system
                    self._update_last_active_window(target, None)
                    
                    QTimer.singleShot(500, lambda: move_window(target, x, y, w, h))
                else:
                    self._log(f"Browser placement failed for {label} (hwnd={target})")
                return
            
            if attempt < 20:  # ~10 seconds for browsers
                QTimer.singleShot(500, lambda: place_by_snapshot(attempt))
            else:
                self._log(f"Browser placement timeout for {label} after {attempt} attempts")
        
    def _launch_mmc_with_snapshot(self, path: str, label: str, args: List[str], uac: bool, quad: str):
        """Launch MMC-based tool with snapshot to handle single-instance behavior."""
        # Take snapshot of current MMC windows BEFORE launch
        before_windows = find_windows_by_criteria(
            image_names=["mmc.exe"],
            class_names=["MMCMainFrame"]
        )
        
        try:
            if uac:
                ok, pid, elevated, err = try_start_exe_uac(path, args=args)
            else:
                ok, pid, elevated, err = try_start_exe(path, args=args)

            if ok:
                self._log(f"{'MMC elevated' if uac else 'MMC'} started → {quad}")
                # Schedule snapshot-based placement
                self._schedule_mmc_placement_by_snapshot(before_windows, quad, label)
            else:
                self._log(f"MMC start failed: {label} | {err}")
        except Exception as e:
            self._log(f"MMC launch exception: {e}")

    def _schedule_mmc_placement_by_snapshot(self, before_windows: List[int], quad: str, label: str):
        """Place MMC window using before/after snapshot comparison."""
        def place_by_snapshot(attempt=0):
            attempt += 1
            
            # Get current MMC windows
            current_windows = find_windows_by_criteria(
                image_names=["mmc.exe"],
                class_names=["MMCMainFrame"]
            )
            
            # Find new windows (not in before snapshot)
            new_windows = [h for h in current_windows if h not in before_windows]
            
            if new_windows:
                # Take the largest new window
                target = max(new_windows, key=get_window_area)
                x, y, w, h = get_quadrant_for_config(quad, self.cfg)
                self._log(f"MMC trying to place hwnd={target} at ({x},{y},{w},{h})")
                
                # Try standard move first
                ok = move_window(target, x, y, w, h)
                if ok:
                    self._log(f"MMC placed: {label} → {quad} (hwnd={target}, new window)")
                    
                    # Update last active window for hotkey system
                    self._update_last_active_window(target, None)
                    
                    QTimer.singleShot(500, lambda: move_window(target, x, y, w, h))
                    return
                
                # If standard move fails, try alternative WinAPI flags
                self._log(f"Standard move failed, trying alternative WinAPI flags...")
                ok_alt = self._try_alternative_window_move(target, x, y, w, h)
                if ok_alt:
                    self._log(f"MMC placed (alternative): {label} → {quad} (hwnd={target})")
                    
                    # Update last active window for hotkey system
                    self._update_last_active_window(target, None)
                    
                    QTimer.singleShot(500, lambda: self._try_alternative_window_move(target, x, y, w, h))
                    return
                
                # Last resort: just detect and log
                try:
                    import win32gui
                    rect = win32gui.GetWindowRect(target)
                    self._log(f"MMC placement failed for {label} (hwnd={target}) - window rect: {rect}")
                    self._log(f"MMC detected but unmoveable: {label} → {quad} (hwnd={target})")
                except Exception as e:
                    self._log(f"MMC placement failed for {label} (hwnd={target}) - error: {e}")
                return
            
            if attempt < 18:  # MMC can be slow to start
                QTimer.singleShot(600, lambda: place_by_snapshot(attempt))
            else:
                self._log(f"MMC placement timeout for {label} after {attempt} attempts")
        
        # MMC needs time to initialize
        QTimer.singleShot(1200, place_by_snapshot)

    def _load_logo(self):
        """Load logo.png if it exists, otherwise show placeholder."""
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
            if os.path.exists(logo_path):
                from PySide6.QtGui import QPixmap
                pixmap = QPixmap(logo_path)
                if not pixmap.isNull():
                    scaled_pixmap = pixmap.scaled(30, 30, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    self.logo_label.setPixmap(scaled_pixmap)
                    self.logo_label.setAlignment(Qt.AlignCenter)
                    return
            
            # Fallback: show text placeholder
            self.logo_label.setText("🚀")
            self.logo_label.setAlignment(Qt.AlignCenter)
            self.logo_label.setStyleSheet(
                "background:#2b3036;border:1px solid #3a4046;border-radius:6px;"
                "color:#7fd0ff;font-size:16px;font-weight:bold;"
            )
        except Exception:
            # Safe fallback
            self.logo_label.setText("SC")
            self.logo_label.setAlignment(Qt.AlignCenter)
            self.logo_label.setStyleSheet(
                "background:#2b3036;border:1px solid #3a4046;border-radius:6px;"
                "color:#7fd0ff;font-size:12px;font-weight:bold;"
            )

    def _maybe_logo_label(self) -> Optional[QLabel]:
        """Load logo.png if it exists, otherwise return None."""
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
            if os.path.exists(logo_path):
                pm = QPixmap(logo_path)
                if not pm.isNull():
                    h = 18
                    pm = pm.scaledToHeight(h, Qt.SmoothTransformation)
                    lbl = QLabel()
                    lbl.setPixmap(pm)
                    lbl.setContentsMargins(0, 0, 6, 0)
                    return lbl
        except Exception:
            pass
        return None

    def _apply_chrome_style(self, w: QWidget, border_boost: float = 0.0):
        """Apply futuristic chrome container styling with glowing effects."""
        # Enhanced glow intensity based on boost
        glow_intensity = int(60 + (40 * max(0.0, min(1.0, border_boost))))
        glow_alpha = min(255, glow_intensity)
        
        w.setStyleSheet(f"""
        QWidget#chrome {{
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:1,
                stop:0 rgba(26, 26, 46, 180), stop:0.2 rgba(22, 33, 62, 160), stop:0.5 rgba(15, 52, 96, 140), stop:0.8 rgba(22, 33, 62, 160), stop:1 rgba(26, 26, 46, 180)
            );
            border: 1px solid qlineargradient(
                x1:0, y1:0, x2:1, y2:0,
                stop:0 rgba(74, 144, 226, {glow_alpha}), 
                stop:0.3 rgba(123, 179, 240, {glow_alpha}), 
                stop:0.7 rgba(123, 179, 240, {glow_alpha}), 
                stop:1 rgba(74, 144, 226, {glow_alpha})
            );
            border-radius: 16px;
        }}
        QSplitter#split::handle {{
            background: transparent;
            margin: 0px;
            height: 1px;
        }}
        """)
        
        # Force style refresh
        w.style().unpolish(w)
        w.style().polish(w)
        w.update()
        
        # Add glowing shadow effect
        self._add_glow_effect(w, intensity=0.4)

    def _add_glow_effect(self, widget, intensity=0.3, color=None):
        """Add futuristic glowing effect to widget."""
        if color is None:
            color = QColor(74, 144, 226)  # Chrome blue
        
        glow = QGraphicsDropShadowEffect()
        glow.setBlurRadius(25)
        glow.setOffset(0, 0)
        glow.setColor(QColor(color.red(), color.green(), color.blue(), int(255 * intensity)))
        widget.setGraphicsEffect(glow)

    def _apply_header_style(self, lbl: QLabel):
        """Apply futuristic header styling with glow effects."""
        lbl.setStyleSheet("""
        QLabel#hdr { 
            color: #e8f4fd; 
            font-weight: 700; 
            font-size: 14px;
            /* text-shadow: 0 0 10px #e8f4fd; */  /* Qt doesn't support text-shadow */
        }
        QLabel#hdrsub { 
            color: #7bb3f0; 
            font-weight: 500; 
            font-size: 12px;
            /* text-shadow: 0 0 5px #7bb3f0; */  /* Qt doesn't support text-shadow */
        }
        """)
        
        # Force style refresh
        lbl.style().unpolish(lbl)
        lbl.style().polish(lbl)
        lbl.update()
        
        # Add subtle glow effect to header
        self._add_glow_effect(lbl, intensity=0.3, color=QColor(123, 179, 240))

    def _apply_combo_style(self, combo: QComboBox):
        """Apply futuristic combo box styling."""
        combo.setStyleSheet("""
        QComboBox#menucombo {
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:1,
                stop:0 #1a1a2e, stop:0.3 #16213e, stop:0.7 #0f3460, stop:1 #1a1a2e
            );
            color: #e8f4fd;
            border: 1px solid qlineargradient(
                x1:0, y1:0, x2:1, y2:0,
                stop:0 rgba(74, 144, 226, 120), stop:0.5 rgba(123, 179, 240, 100), stop:1 rgba(74, 144, 226, 120)
            );
            border-radius: 10px;
            padding: 6px 10px;
            min-width: 160px;
            font-weight: 500;
        }
        QComboBox#menucombo::drop-down { 
            width: 0px; 
            border: none;
            background: transparent;
        }
        QComboBox#menucombo::drop-down:hover {
            background: transparent;
        }
        QComboBox#menucombo:hover { 
            border: 1px solid rgba(123, 179, 240, 150);
            color: #f0f8ff;
        }
        QComboBox QAbstractItemView {
            background: #1a1a2e;
            color: #e8f4fd;
            selection-background-color: #4a90e2;
            selection-color: #ffffff;
            border: 1px solid rgba(123, 179, 240, 120);
            border-radius: 8px;
            padding: 4px;
        }
        QComboBox QAbstractItemView::item {
            padding: 6px 10px;
            border-radius: 4px;
        }
        QComboBox QAbstractItemView::item:hover {
            background: #1a1a1a;
            color: #00ffff;
        }
        """)
        
        # Force style refresh
        combo.style().unpolish(combo)
        combo.style().polish(combo)
        combo.update()
        
        # Add glow effect to combo box
        self._add_glow_effect(combo, intensity=0.2, color=QColor(74, 144, 226))

    def keyPressEvent(self, event):
        """Handle hotkey presses for window management."""
        try:
            # Track user activity for animation control
            self._resume_animations()
            
            # Check for Ctrl+Number combinations
            if event.modifiers() == Qt.ControlModifier:
                if event.key() == Qt.Key_0:
                    # Ctrl+0: Center GUI to its start quadrant
                    self._center_gui()
                    event.accept()
                    return
                elif event.key() == Qt.Key_1:
                    # Ctrl+1: Place last window in first RR_ORDER quadrant
                    self._place_last_window_rr(0)
                    event.accept()
                    return
                elif event.key() == Qt.Key_2:
                    # Ctrl+2: Place last window in second RR_ORDER quadrant
                    self._place_last_window_rr(1)
                    event.accept()
                    return
                elif event.key() == Qt.Key_3:
                    # Ctrl+3: Place last window in third RR_ORDER quadrant
                    self._place_last_window_rr(2)
                    event.accept()
                    return
                    
        except Exception as e:
            self._log(f"Hotkey error: {e}")
            
        # Call parent keyPressEvent for other keys
        super().keyPressEvent(event)

    def _center_gui(self):
        """Center the GUI window in its configured start quadrant."""
        try:
            # Get primary screen geometry
            screen = QApplication.primaryScreen()
            screen_geometry = screen.availableGeometry()
            
            # Get configured start quadrant (default: TR)
            quadrant = getattr(self.cfg, 'GUI_START_QUADRANT', 'TR')
            
            # Calculate quadrant dimensions with DPI scaling
            margin = int(screen_geometry.width() * self.cfg.EDGE_MARGIN_RATIO)
            
            # Get DPI scaling factor
            dpi_scale = screen.devicePixelRatio()
            logical_dpi = screen.logicalDotsPerInch()
            scale_factor = logical_dpi / 96.0  # 96 DPI is standard
            
            # Calculate scaled dimensions based on screen size
            if screen_geometry.width() <= 1920:  # 1080p or smaller
                w = int(screen_geometry.width() * 0.6) - margin  # 60% for smaller screens
                h = int(screen_geometry.height() * 0.7) - margin  # 70% height for better visibility
            else:  # 4K or larger
                w = int(screen_geometry.width() * 0.4) - margin  # 40% for larger screens
                h = int(screen_geometry.height() * 0.5) - margin  # 50% height
                
            # Apply minimum and maximum size constraints
            min_width = int(600 * scale_factor)
            min_height = int(400 * scale_factor)
            max_width = int(screen_geometry.width() * 0.8)
            max_height = int(screen_geometry.height() * 0.8)
            
            w = max(min_width, min(w, max_width))
            h = max(min_height, min(h, max_height))
            
            # Calculate position based on quadrant
            if quadrant == "TL":  # Top-Left
                x = margin
                y = margin
            elif quadrant == "TR":  # Top-Right
                x = int(screen_geometry.width() * 0.5) + margin
                y = margin
            elif quadrant == "BL":  # Bottom-Left
                x = margin
                y = int(screen_geometry.height() * 0.5) + margin
            elif quadrant == "BR":  # Bottom-Right
                x = int(screen_geometry.width() * 0.5) + margin
                y = int(screen_geometry.height() * 0.5) + margin
            else:  # Fallback to TR
                x = int(screen_geometry.width() * 0.5) + margin
                y = margin
                quadrant = "TR"
            
            # Center within the quadrant
            center_x = x + (w - self.width()) // 2
            center_y = y + (h - self.height()) // 2
            
            # Move to quadrant center
            self.move(center_x, center_y)
            self._log(f"GUI centered in {quadrant} quadrant at ({center_x}, {center_y})")
            
        except Exception as e:
            self._log(f"Center GUI error: {e}")

    def _place_last_window(self, quadrant):
        """Place the last active window in the specified quadrant."""
        try:
            if not self._last_active_window:
                self.log("No last active window to place")
                return
                
            # Get window handle and PID
            hwnd = self._last_active_window
            pid = self._last_active_pid
            
            if not hwnd or not win32gui.IsWindow(hwnd):
                self.log("Last active window no longer valid")
                self._last_active_window = None
                self._last_active_pid = None
                return
                
            # Place window using existing logic
            if quadrant == "TL":
                self._place_window_tl(hwnd, pid)
            elif quadrant == "TR":
                self._place_window_tr(hwnd, pid)
            elif quadrant == "BL":
                self._place_window_bl(hwnd, pid)
            elif quadrant == "BR":
                self._place_window_br(hwnd, pid)
                
            self._log(f"Last active window placed in {quadrant}")
            
        except Exception as e:
            self._log(f"Place last window error: {e}")

    def _place_last_window_rr(self, rr_index: int):
        """Place the last active window in the RR_ORDER quadrant at the specified index."""
        try:
            if not self._last_active_window:
                self._log("No last active window to place")
                return
                
            # Get RR_ORDER quadrants from config
            rr_quadrants = self.cfg.RR_ORDER
            if rr_index >= len(rr_quadrants):
                self._log(f"RR_ORDER index {rr_index} out of range (max: {len(rr_quadrants)-1})")
                return
                
            quad = rr_quadrants[rr_index]
            self._log(f"Placing last window in RR_ORDER[{rr_index}] = {quad}")
            
            # Get window handle and PID
            hwnd = self._last_active_window
            pid = self._last_active_pid
            
            if not hwnd or not win32gui.IsWindow(hwnd):
                self._log("Last active window no longer valid")
                self._last_active_window = None
                self._last_active_pid = None
                return
                
            # Place window using existing logic
            if quad == "TL":
                self._place_window_tl(hwnd, pid)
            elif quad == "TR":
                self._place_window_tr(hwnd, pid)
            elif quad == "BL":
                self._place_window_bl(hwnd, pid)
            elif quad == "BR":
                self._place_window_br(hwnd, pid)
                
            self._log(f"Last active window placed in {quad} (RR_ORDER[{rr_index}])")
            
        except Exception as e:
            self._log(f"Error placing last active window: {e}")

    def _schedule_powershell_placement(self, pid: int, quad: str, label: str, title: str):
        """Enhanced PowerShell window placement with multiple detection methods."""
        max_attempts = 25  # More attempts for PowerShell
        
        def place_powershell(attempt=0):
            attempt += 1
            
            # Try multiple detection methods for PowerShell
            windows = []
            
            # Method 1: PID-based detection
            windows = find_windows_by_criteria(pid=pid)
            
            # Method 2: Title-based detection
            if not windows:
                windows = find_windows_by_criteria(title_contains=[title])
            
            # Method 3: PowerShell-specific class detection
            if not windows:
                windows = find_windows_by_criteria(
                    class_names=["ConsoleWindowClass", "VirtualConsoleClass", "CASCADIA_HOSTING_WINDOW_CLASS"],
                    image_names=["powershell.exe", "pwsh.exe", "WindowsTerminal.exe"]
                )
            
            # Method 4: Console window detection
            if not windows:
                windows = self._find_console_windows_by_pid(pid)
            
            if windows:
                target = max(windows, key=get_window_area)
                x, y, w, h = get_quadrant_for_config(quad, self.cfg)
                
                # Debug: Log window details
                try:
                    win_title = win32gui.GetWindowText(target)
                    win_class = win32gui.GetClassName(target)
                    self._log(f"PowerShell window found: hwnd={target}, class={win_class}, title='{win_title}'")
                except:
                    pass
                
                # Enhanced PowerShell placement
                success = self._place_console_window(target, x, y, w, h, label)
                
                if success:
                    self._log(f"PowerShell placed: {label} → {quad} (hwnd={target}, PID={pid})")
                    
                    # Update last active window for hotkey system
                    self._update_last_active_window(target, pid)
                    
                    # Double-check placement
                    QTimer.singleShot(500, lambda: self._place_console_window(target, x, y, w, h, label))
                else:
                    self._log(f"PowerShell placement failed for {label} (hwnd={target}) - trying alternative methods")
                    # Try alternative placement methods
                    self._try_powershell_alternative_placement(target, x, y, w, h, label)
                return
                
            if attempt < max_attempts:
                # Longer delay for PowerShell
                delay = 800 if attempt < 10 else 1200
                QTimer.singleShot(delay, lambda: place_powershell(attempt))
            else:
                self._log(f"PowerShell placement timeout for {label} after {attempt} attempts")
        
        QTimer.singleShot(1000, place_powershell)

    def _try_powershell_alternative_placement(self, hwnd: int, x: int, y: int, w: int, h: int, label: str):
        """Try alternative placement methods for PowerShell windows."""
        try:
            # Method 1: Direct WinAPI with different flags
            flags = win32con.SWP_SHOWWINDOW | win32con.SWP_NOZORDER
            result = win32gui.SetWindowPos(hwnd, 0, x, y, w, h, flags)
            if result:
                self._log(f"PowerShell placed (WinAPI): {label} at ({x}, {y}, {w}, {h})")
                return True
        except Exception as e:
            self._log(f"WinAPI placement failed: {e}")
        
        try:
            # Method 2: MoveWindow
            result = win32gui.MoveWindow(hwnd, x, y, w, h, True)
            if result:
                self._log(f"PowerShell placed (MoveWindow): {label} at ({x}, {y}, {w}, {h})")
                return True
        except Exception as e:
            self._log(f"MoveWindow placement failed: {e}")
        
        try:
            # Method 3: SetWindowPos with different flags
            flags = win32con.SWP_SHOWWINDOW | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE
            result = win32gui.SetWindowPos(hwnd, 0, x, y, w, h, flags)
            if result:
                self._log(f"PowerShell placed (SetWindowPos): {label} at ({x}, {y}, {w}, {h})")
                return True
        except Exception as e:
            self._log(f"SetWindowPos placement failed: {e}")
        
        self._log(f"All alternative placement methods failed for {label}")
        return False

    def _update_last_active_window(self, hwnd, pid):
        """Update the last active window for hotkey placement."""
        try:
            if hwnd and win32gui.IsWindow(hwnd):
                self._last_active_window = hwnd
                self._last_active_pid = pid
                self._log(f"Last active window updated: hwnd={hwnd}, pid={pid}")
        except Exception:
            pass

    def _init_animations(self):
        """Initialize CPU-efficient animation system for VDI environments."""
        # Animation state - MUCH more conservative
        self._animation_phase = 0.0
        self._animation_timer = QTimer(self)
        self._animation_timer.timeout.connect(self._update_animations)
        self._animation_timer.start(200)  # 5 FPS - VDI optimized (was 20 FPS)
        
        # Pulsing state for different elements
        self._chrome_pulse = 0.0
        self._button_pulse = 0.0
        self._console_pulse = 0.0
        
        # Animation control - only animate when needed
        self._animations_enabled = True
        self._last_activity = time.time()

    def _update_animations(self):
        """Update all animations - CPU efficient for VDI."""
        try:
            # Smart animation control - pause when idle
            current_time = time.time()
            if current_time - self._last_activity > 10:  # Pause after 10s idle
                if self._animations_enabled:
                    self._animations_enabled = False
                    self._pause_animations()
                return
            elif not self._animations_enabled:
                self._animations_enabled = True
            
            # Increment animation phase - much slower
            self._animation_phase += 0.05  # Was 0.1 - half speed
            if self._animation_phase > 6.28:  # 2*PI
                self._animation_phase = 0.0
            
            # Calculate pulse values - much more subtle
            self._chrome_pulse = 0.2 + 0.05 * math.sin(self._animation_phase * 0.3)  # Was 0.3 + 0.1
            self._button_pulse = 0.1 + 0.03 * math.sin(self._animation_phase * 0.4)  # Was 0.2 + 0.05
            self._console_pulse = 0.05 + 0.02 * math.sin(self._animation_phase * 0.2)  # Was 0.1 + 0.05
            
            # Update chrome glow - only if significant change
            if abs(self._chrome_pulse - getattr(self, '_last_chrome_pulse', 0)) > 0.01:
                self._apply_chrome_style(self.chrome, border_boost=self._chrome_pulse)
                self._last_chrome_pulse = self._chrome_pulse
            
            # Update button glows - only if buttons exist and significant change
            if hasattr(self, 'btns') and abs(self._button_pulse - getattr(self, '_last_button_pulse', 0)) > 0.01:
                self._update_button_animations()
                self._last_button_pulse = self._button_pulse
                
        except Exception as e:
            # Silent fail for VDI compatibility
            pass

    def _pause_animations(self):
        """Pause animations to save CPU when idle."""
        try:
            # Reset to base state
            self._chrome_pulse = 0.2
            self._button_pulse = 0.1
            self._console_pulse = 0.05
            
            # Apply static state
            self._apply_chrome_style(self.chrome, border_boost=self._chrome_pulse)
            if hasattr(self, 'btns'):
                self._update_button_animations()
        except Exception:
            pass

    def _resume_animations(self):
        """Resume animations when user becomes active."""
        self._last_activity = time.time()
        self._animations_enabled = True

    def _update_button_animations(self):
        """Update button glow animations - CPU efficient."""
        try:
            for btn in self.btns.values():
                if hasattr(btn, '_update_glow'):
                    btn._update_glow(self._button_pulse)
        except Exception:
            pass

    def _apply_console_style(self, console: QTextEdit):
        """Apply futuristic Matrix-style console styling."""
        console.setStyleSheet("""
        QTextEdit {
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:1,
                stop:0 #0a0a0a, stop:0.3 #1a1a1a, stop:0.7 #0f0f0f, stop:1 #0a0a0a
            );
            color: #00ff41;
            border: 1px solid qlineargradient(
                x1:0, y1:0, x2:1, y2:0,
                stop:0 rgba(74, 144, 226, 120), stop:0.5 rgba(123, 179, 240, 100), stop:1 rgba(74, 144, 226, 120)
            );
            border-radius: 12px;
            font-family: 'Consolas', 'Cascadia Mono', 'Courier New', monospace;
            font-size: 11px;
            padding: 12px;
            selection-background-color: #00ff41;
            selection-color: #000000;
        }
        QTextEdit::scrollbar:vertical {
            background: #1a1a1a;
            width: 14px;
            border-radius: 7px;
            border: 1px solid #00ff41;
        }
        QTextEdit::scrollbar::handle:vertical {
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:0,
                stop:0 #4a90e2, stop:0.5 #7bb3f0, stop:1 #4a90e2
            );
            border-radius: 6px;
            min-height: 20px;
        }
        QTextEdit::scrollbar::handle:vertical:hover {
            background: #00ffff;
        }
        QTextEdit::scrollbar::add-line:vertical,
        QTextEdit::scrollbar::sub-line:vertical {
            height: 0px;
        }
        """)
        
        # Set monospace font
        f = QFont()
        f.setStyleHint(QFont.Monospace)
        f.setFamily("Consolas")
        f.setPointSize(11)
        console.setFont(f)
        
        # Add subtle glow effect to console (keep Python green)
        self._add_glow_effect(console, intensity=0.2, color=QColor(0, 255, 65))

    def _toggle_maximize(self):
        """Toggle between maximized and normal window state."""
        if self.isMaximized():
            self.showNormal()
            self.btn_max.setText("⬜")
            self._log("Window restored to normal size")
        else:
            self.showMaximized()
            self.btn_max.setText("⬛")
            self._log("Window maximized")

    def _copy_console(self):
        """Copy console content to clipboard."""
        try:
            console_text = self.console.toPlainText()
            if console_text:
                from PySide6.QtGui import QGuiApplication
                clipboard = QGuiApplication.clipboard()
                clipboard.setText(console_text)
                
                # Brief visual feedback
                original_text = self.console.toHtml()
                self.console.setHtml(original_text + "<br/><span style='color:#06d6a0;font-weight:bold;'>Console content copied to clipboard!</span>")
                
                # Remove feedback after 2 seconds
                QTimer.singleShot(2000, lambda: self.console.setHtml(original_text))
                
                self._log("Console content copied to clipboard")
            else:
                self._log("Console is empty - nothing to copy")
        except Exception as e:
            self._log(f"Copy to clipboard failed: {e}")

    def _try_alternative_window_move(self, hwnd: int, x: int, y: int, w: int, h: int) -> bool:
        """Try alternative WinAPI methods for stubborn windows like MMC."""
        try:
            import win32gui
            import win32con
            
            # Method 1: Try with different SetWindowPos flags
            try:
                # SWP_NOACTIVATE | SWP_NOZORDER (don't activate or change Z-order)
                win32gui.SetWindowPos(hwnd, None, x, y, w, h, 
                                     win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER)
                self._log(f"Alternative move method 1 succeeded for hwnd={hwnd}")
                return True
            except Exception:
                pass
            
            # Method 2: Try MoveWindow API instead of SetWindowPos
            try:
                win32gui.MoveWindow(hwnd, x, y, w, h, True)  # True = repaint
                self._log(f"Alternative move method 2 (MoveWindow) succeeded for hwnd={hwnd}")
                return True
            except Exception:
                pass
            
            # Method 3: Try ShowWindow + SetWindowPos combination
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # Ensure not minimized
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, w, h, 
                                     win32con.SWP_NOACTIVATE)
                self._log(f"Alternative move method 3 (ShowWindow+SetWindowPos) succeeded for hwnd={hwnd}")
                return True
            except Exception:
                pass
            
            self._log(f"All alternative move methods failed for hwnd={hwnd}")
            return False
            
        except Exception as e:
            self._log(f"Alternative move method exception: {e}")
            return False

    def _schedule_window_placement_by_pid(self, pid: int, quad: str, label: str, is_browser: bool = False):
        """Schedule window placement based on PID with enhanced CMD/PS1 support."""        
        max_attempts = 20  # Increased for CMD/PS1 which can be slower
        
        def place_by_pid(attempt=0):
            attempt += 1
            
            # Enhanced window detection for CMD/PS1
            windows = find_windows_by_criteria(pid=pid)
            
            # Special handling for console applications (CMD/PS1)
            if not windows and label.lower() in ['cmd', 'powershell', 'ps']:
                # Try alternative detection methods for console windows
                windows = self._find_console_windows_by_pid(pid)
            
            if windows:
                target = max(windows, key=get_window_area)
                x, y, w, h = get_quadrant_for_config(quad, self.cfg)
                
                # Enhanced placement for console applications
                if label.lower() in ['cmd', 'powershell', 'ps']:
                    success = self._place_console_window(target, x, y, w, h, label)
                else:
                    success = move_window(target, x, y, w, h)
                
                if success:
                    self._log(f"Window placed: {label} → {quad} (hwnd={target}, PID={pid})")
                    
                    # Update last active window for hotkey system
                    self._update_last_active_window(target, pid)
                    
                    # Double-check placement for console apps
                    if label.lower() in ['cmd', 'powershell', 'ps']:
                        QTimer.singleShot(300, lambda: self._place_console_window(target, x, y, w, h, label))
                    else:
                        QTimer.singleShot(500, lambda: move_window(target, x, y, w, h))
                else:
                    self._log(f"Window placement failed for {label} (hwnd={target})")
                return
                
            if attempt < max_attempts:
                # Longer delay for console applications
                delay = 600 if label.lower() in ['cmd', 'powershell', 'ps'] else 500
                QTimer.singleShot(delay, lambda: place_by_pid(attempt))
            else:
                self._log(f"Window placement timeout for {label} (PID={pid}) after {attempt} attempts")
        
        # Longer initial delay for console applications
        initial_delay = 1200 if label.lower() in ['cmd', 'powershell', 'ps'] else 800
        QTimer.singleShot(initial_delay, place_by_pid)

    def _find_console_windows_by_pid(self, pid: int) -> List[int]:
        """Enhanced console window detection for CMD/PS1 applications."""
        try:
            import win32gui
            import win32process
            
            # Look for console-specific window classes
            console_classes = ["ConsoleWindowClass", "VirtualConsoleClass"]
            windows = []
            
            def enum_callback(hwnd, _):
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return True
                    if win32gui.GetParent(hwnd):
                        return True
                    
                    # Check if it's a console window
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name not in console_classes:
                        return True
                    
                    # Check PID match
                    _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                    if int(window_pid) != int(pid):
                        return True
                    
                    # Check minimum size
                    l, t, r, b = win32gui.GetWindowRect(hwnd)
                    if (r - l) < 200 or (b - t) < 120:
                        return True
                    
                    windows.append(hwnd)
                except Exception:
                    pass
                return True
            
            win32gui.EnumWindows(enum_callback, None)
            return windows
            
        except Exception as e:
            self._log(f"Console window detection failed: {e}")
            return []

    def _place_console_window(self, hwnd: int, x: int, y: int, w: int, h: int, label: str) -> bool:
        """Enhanced console window placement with multiple fallback methods."""
        try:
            import win32gui
            import win32con
            
            # Method 1: Standard SetWindowPos
            try:
                if move_window(hwnd, x, y, w, h):
                    return True
            except Exception:
                pass
            
            # Method 2: Force console window restoration and placement
            try:
                # Ensure window is not minimized
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                
                # Try SetWindowPos with different flags
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, w, h, 
                                    win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE)
                return True
            except Exception:
                pass
            
            # Method 3: MoveWindow API (often works better for console windows)
            try:
                win32gui.MoveWindow(hwnd, x, y, w, h, True)
                return True
            except Exception:
                pass
            
            # Method 4: Alternative SetWindowPos flags
            try:
                win32gui.SetWindowPos(hwnd, None, x, y, w, h, 
                                    win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
                return True
            except Exception:
                pass
            
            return False
            
        except Exception as e:
            self._log(f"Console placement failed for {label}: {e}")
            return False

    def _get_current_session_id(self) -> int:
        """Get current RDSH session ID for proper window targeting."""
        try:
            import win32ts
            return win32ts.WTSGetCurrentSessionId()
        except Exception:
            return 0

    def _launch_rds_application(self, path: str, label: str, args: List[str], uac: bool, quad: str, window_placement: dict):
        """Launch RDSH application with enhanced session-aware window placement."""
        try:
            # Get current session ID for logging
            current_session = self._get_current_session_id()
            self._log(f"RDSH Launch: {label} in session {current_session}")
            
            # Launch the application
            if uac:
                ok, pid, elevated, err = try_start_exe_uac(path, args=args)
            else:
                ok, pid, elevated, err = try_start_exe(path, args=args)
            
            if ok and pid:
                self._log(f"RDSH App started: {label} (PID={pid}, Session={current_session}) → {quad}")
                
                # Enhanced RDSH window placement
                self._schedule_rds_window_placement(pid, quad, label, window_placement, current_session)
            elif ok:
                self._log(f"RDSH App started (no PID): {label} → {quad}")
                # Fallback to generic placement
                self._schedule_window_placement_generic(path, quad, label)
            else:
                self._log(f"RDSH App start failed: {label} | {err}")
                
        except Exception as e:
            self._log(f"RDSH Launch exception: {e}")

    def _schedule_rds_window_placement(self, pid: int, quad: str, label: str, window_placement: dict, session_id: int):
        """Schedule RDSH-aware window placement with enhanced detection."""
        max_attempts = 25  # Longer timeout for RDSH applications
        method = window_placement.get("method", "pid")
        title_contains = window_placement.get("title_contains", [])
        class_names = window_placement.get("class_names", [])
        fallback_timeout = window_placement.get("fallback_timeout", 5000)
        
        def place_rds_window(attempt=0):
            attempt += 1
            
            # Try different detection methods based on configuration
            windows = []
            
            if method == "pid":
                # Standard PID-based detection
                windows = find_windows_by_criteria(pid=pid, session_id=session_id)
            elif method == "title_match":
                # Title-based detection for RDSH applications
                windows = find_windows_by_criteria(
                    title_contains=title_contains,
                    class_names=class_names,
                    session_id=session_id
                )
            elif method == "hybrid":
                # Try both PID and title matching
                windows = find_windows_by_criteria(pid=pid, session_id=session_id)
                if not windows:
                    windows = find_windows_by_criteria(
                        title_contains=title_contains,
                        class_names=class_names,
                        session_id=session_id
                    )
            
            if windows:
                target = max(windows, key=get_window_area)
                x, y, w, h = get_quadrant_for_config(quad, self.cfg)
                
                # Enhanced placement for RDSH windows
                success = self._place_rds_window(target, x, y, w, h, label)
                
                if success:
                    self._log(f"RDSH Window placed: {label} → {quad} (hwnd={target}, Session={session_id})")
                    
                    # Update last active window for hotkey system
                    self._update_last_active_window(target, pid)
                    
                    # Double-check placement
                    QTimer.singleShot(500, lambda: self._place_rds_window(target, x, y, w, h, label))
                else:
                    self._log(f"RDSH Window placement failed for {label} (hwnd={target})")
                return
            
            if attempt < max_attempts:
                # Longer delay for RDSH applications
                delay = 800 if attempt < 10 else 1000
                QTimer.singleShot(delay, lambda: place_rds_window(attempt))
            else:
                self._log(f"RDSH Window placement timeout for {label} after {attempt} attempts")
        
        # Initial delay for RDSH applications
        QTimer.singleShot(1500, place_rds_window)

    def _place_rds_window(self, hwnd: int, x: int, y: int, w: int, h: int, label: str) -> bool:
        """Enhanced RDSH window placement with multiple fallback methods."""
        try:
            import win32gui
            import win32con
            
            # Method 1: Standard placement
            if move_window(hwnd, x, y, w, h):
                return True
            
            # Method 2: RDSH-specific placement (often needs different flags)
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, w, h, 
                                    win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE)
                return True
            except Exception:
                pass
            
            # Method 3: Alternative flags for RDSH windows
            try:
                win32gui.SetWindowPos(hwnd, None, x, y, w, h, 
                                    win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_SHOWWINDOW)
                return True
            except Exception:
                pass
            
            # Method 4: MoveWindow as last resort
            try:
                win32gui.MoveWindow(hwnd, x, y, w, h, True)
                return True
            except Exception:
                pass
            
            return False
            
        except Exception as e:
            self._log(f"RDSH placement failed for {label}: {e}")
            return False

    def _schedule_window_placement_by_title(self, title: str, quad: str, label: str):
        """Schedule window placement based on window title."""
        def place_by_title(attempt=0):
            attempt += 1
            windows = find_windows_by_criteria(title_contains=[title])
            if windows:
                target = max(windows, key=get_window_area)
                x, y, w, h = get_quadrant_for_config(quad, self.cfg)
                ok = move_window(target, x, y, w, h)
                if ok:
                    self._log(f"Window placed: {label} → {quad} (hwnd={target}, title='{title}')")
                    
                    # Update last active window for hotkey system
                    self._update_last_active_window(target, None)
                    
                    QTimer.singleShot(500, lambda: move_window(target, x, y, w, h))
                else:
                    self._log(f"Window placement failed for {label} (hwnd={target})")
                return
            if attempt < 15:
                QTimer.singleShot(500, lambda: place_by_title(attempt))
            else:
                self._log(f"Window placement timeout for {label} (title='{title}')")
        
        QTimer.singleShot(1200, place_by_title)  # Longer delay for elevated processes

    def _schedule_window_placement_generic(self, exe_path: str, quad: str, label: str):
        """Schedule window placement based on executable name."""
        exe_name = os.path.basename(exe_path).lower()
        
        # Heuristics for different program types
        image_names = [exe_name]
        class_names = []
        title_hints = []
        
        if exe_name in ("taskmgr.exe", "taskmanager.exe"):
            class_names = ["TaskManagerWindow"]
            title_hints = ["task manager", "taskmanager"]
        elif exe_name in ("msedge.exe", "chrome.exe", "firefox.exe"):
            image_names = ["msedge.exe", "chrome.exe", "firefox.exe"]
            class_names = ["Chrome_WidgetWin_1", "MozillaWindowClass", "ApplicationFrameWindow"]
        elif exe_name in ("eventvwr.exe", "mmc.exe"):
            class_names = ["MMCMainFrame"]
            title_hints = ["event viewer", "ereignisanzeige"]
        
        def place_generic(attempt=0):
            attempt += 1
            windows = find_windows_by_criteria(
                image_names=image_names,
                class_names=class_names if class_names else None,
                title_contains=title_hints if title_hints else None
            )
            if windows:
                target = max(windows, key=get_window_area)
                x, y, w, h = get_quadrant_for_config(quad, self.cfg)
                ok = move_window(target, x, y, w, h)
                if ok:
                    self._log(f"Window placed: {label} → {quad} (hwnd={target}, exe={exe_name})")
                    
                    # Update last active window for hotkey system
                    self._update_last_active_window(target, None)
                    
                    QTimer.singleShot(500, lambda: move_window(target, x, y, w, h))
                else:
                    self._log(f"Window placement failed for {label} (hwnd={target})")
                return
            if attempt < 12:
                QTimer.singleShot(600, lambda: place_generic(attempt))
            else:
                self._log(f"Window placement timeout for {label} (exe={exe_name})")
        
        QTimer.singleShot(1000, place_generic)

    def _launch_url_isolated(self, url: str, quad: str):
        """Launch URL in isolated browser."""
        browsers = {
            "edge": [which("msedge.exe"), r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"],
            "chrome": [which("chrome.exe"), r"C:\Program Files\Google\Chrome\Application\chrome.exe"],
            "firefox": [which("firefox.exe"), r"C:\Program Files\Mozilla Firefox\firefox.exe"]
        }
        
        browser_exe = None
        browser_name = ""
        
        for name, paths in browsers.items():
            for path in paths:
                if path and os.path.exists(path):
                    browser_exe = path
                    browser_name = name
                    break
            if browser_exe:
                break
        
        if not browser_exe:
            # Fallback to default browser
            if os.name == "nt":
                os.startfile(url)
            self._log(f"URL launched (default browser): {url}")
            return
        
        # Create temp profile
        import random
        temp_profile = os.path.join(tempfile.gettempdir(), f"CockpitBrowser_{browser_name}_{random.randint(1000,9999)}")
        os.makedirs(temp_profile, exist_ok=True)
        
        if browser_name in ("edge", "chrome"):
            args = [browser_exe, "--new-window", f"--user-data-dir={temp_profile}", "--no-first-run", url]
        elif browser_name == "firefox":
            args = [browser_exe, "-no-remote", "-new-instance", "-profile", temp_profile, url]
        
        subprocess.Popen(args)
        self._log(f"URL launched (isolated {browser_name}): {url}")

    def _launch_explorer_fixed(self, label: str, path: str, args, quad: str):
        """Launch Explorer with placement."""
        try:
            explorer_path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "explorer.exe")
            
            if args:
                # Use provided args as-is (they should already include /n, /e,)
                full_args = [explorer_path] + list(args)
            elif path and os.path.basename(path).lower() not in ("explorer.exe", "explorer"):
                full_args = [explorer_path, "/n,", "/e,", path]
            else:
                # Default: This PC
                full_args = [explorer_path, "/n,", "/e,", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"]
            
            before = get_explorer_windows()
            subprocess.Popen(full_args)
            self._log(f"Explorer launched: {' '.join(full_args)}")
            
            def poll_explorer(attempt=0):
                attempt += 1
                current = get_explorer_windows()
                new_hwnds = [h for h in current if h not in before]
                if new_hwnds:
                    target = max(new_hwnds, key=get_window_area)
                    x, y, w, h = get_quadrant_for_config(quad, self.cfg)
                    ok = move_window(target, x, y, w, h)
                    if ok:
                        self._log(f"Explorer placed: {label} → {quad} (hwnd={target})")
                        
                        # Update last active window for hotkey system
                        self._update_last_active_window(target, None)
                        
                    else:
                        self._log(f"Explorer placement failed for {target}")
                    return
                if attempt < 20:
                    QTimer.singleShot(500, lambda: poll_explorer(attempt))
                else:
                    self._log(f"Explorer timeout for {label}")
            
            QTimer.singleShot(300, poll_explorer)
            
        except Exception as e:
            self._log(f"Explorer launch failed: {e}")

    # Drag handling
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            if getattr(e, "position", None):
                y = e.position().y()
                pos_to_point = e.globalPosition().toPoint()
            else:
                y = e.pos().y()
                pos_to_point = e.globalPos()
            if y <= 40 or self.header.underMouse():
                self._dragging = True
                self._drag_off = pos_to_point - self.frameGeometry().topLeft()
                e.accept()
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if self._dragging and (e.buttons() & Qt.LeftButton):
            if getattr(e, "globalPosition", None):
                pt = e.globalPosition().toPoint() - self._drag_off
            else:
                pt = e.globalPos() - self._drag_off
            x, y, w, h = clamp_rect_to_primary(pt.x(), pt.y(), self.width(), self.height())
            self.setGeometry(x, y, w, h)
            e.accept()
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        self._dragging = False
        super().mouseReleaseEvent(e)

    def resizeEvent(self, e):
        """Handle window resize to dynamically adjust button grid."""
        super().resizeEvent(e)
        # Rebuild buttons with new column count after a short delay
        # to avoid rebuilding too frequently during resize
        if hasattr(self, '_resize_timer'):
            self._resize_timer.stop()
        self._resize_timer = QTimer(self)  # type: ignore
        self._resize_timer.timeout.connect(self._rebuild_buttons_on_resize)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.start(300)  # 300ms delay

    def _rebuild_buttons_on_resize(self):
        """Rebuild buttons after window resize."""
        if hasattr(self, 'splitter') and self.splitter.count() > 0:
            self._build_buttons()

    # Baseline
    def _start_baseline(self):
        ps = which("pwsh.exe","powershell.exe") or r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        script = self.cfg.BASELINE_PATH  # Wurde bereits von apply_json_overrides aufgelöst
        args = list(self.cfg.BASELINE_ARGS)

        env = os.environ.copy()
        env["COCKPIT_NONINTERACTIVE"] = "1"

        ps_cmd = [
            ps, "-NoLogo","-NoProfile","-ExecutionPolicy","Bypass","-NonInteractive",
            "-WindowStyle", "Hidden", "-Command",
            "[Console]::OutputEncoding=[Text.UTF8Encoding]::new(); "
            "$OutputEncoding=[Text.UTF8Encoding]::new(); "
            "$env:COCKPIT_NONINTERACTIVE='1'; "
            "$InformationPreference='Continue'; "
            f"& '{script}' " + " ".join(map(shlex.quote, args)) + " *>&1"
        ]

        try:
            self._baseline_p = subprocess.Popen(
                ps_cmd,
                cwd=os.path.dirname(script) or None,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                stdin=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW,
                env=env
            )
        except Exception as e:
            self._log(f"[Baseline] Start failed: {e}")
            return

        self._log("[Baseline] started")
        self._q_out, self._q_err = Queue(), Queue()

        def _reader(stream, q):
            for raw in iter(stream.readline, b""):
                try:
                    q.put(raw.decode("utf-8", errors="ignore"))
                except Exception:
                    pass
            try: stream.close()
            except Exception: pass

        threading.Thread(target=_reader, args=(self._baseline_p.stdout, self._q_out), daemon=True).start()
        threading.Thread(target=_reader, args=(self._baseline_p.stderr, self._q_err), daemon=True).start()

        self._baseline_timer = QTimer(self)
        self._baseline_timer.timeout.connect(self._baseline_pump)
        self._baseline_timer.start(120)

    def _baseline_pump(self):
        try:
            while True:
                line = self._q_out.get_nowait()
                self._append_html(self._stylize(f"[Baseline] {line.rstrip()}"))
        except Empty:
            pass
        try:
            while True:
                line = self._q_err.get_nowait()
                self._append_html(self._fmt_line(f"[Baseline][ERR] {line.rstrip()}", "#ff6b6b"))
        except Empty:
            pass
        if self._baseline_p and (self._baseline_p.poll() is not None):
            code = self._baseline_p.returncode
            self._append_html(self._fmt_line(f"[Baseline] finished (ExitCode={code})", "#9fa7ad"))
            if self._baseline_timer: self._baseline_timer.stop()
            self._baseline_p = None

    # Console helpers
    def _append_html(self, html_line: str):
        self.console.moveCursor(QTextCursor.End)
        self.console.insertHtml(html_line)
        self.console.insertHtml("<br/>")
        self.console.moveCursor(QTextCursor.End)

    def _fmt_line(self, text:str, col:str="#bfe0bf", emoji:str=""):
        esc = html.escape(text)
        prefix = f"<span style='color:{col}'>{emoji}</span> " if emoji else ""
        return f"<span style='font-family:Consolas; color:{col}'>{prefix}{esc}</span>"

    def _stylize(self, raw:str) -> str:
        low = raw.lower()
        if any(k in low for k in ("error","err","fail","failed")):
            return self._fmt_line(raw, "#ff6b6b")
        if any(k in low for k in ("warn","warning")):
            return self._fmt_line(raw, "#ffd166")
        if any(k in low for k in ("success","ok","passed","done")):
            return self._fmt_line(raw, "#06d6a0")
        return self._fmt_line(raw, "#bfe0bf")

    def _build_log(self):
        self.console=QTextEdit()
        self.console.setReadOnly(True)
        self._apply_console_style(self.console)
        self.splitter.addWidget(self.console)
        self.splitter.setStretchFactor(0,1); self.splitter.setStretchFactor(1,2)
        
        # Force style refresh for console
        self.console.style().unpolish(self.console)
        self.console.style().polish(self.console)
        self.console.update()

    def _rebuild_menu_combo(self):
        names = list(self.state.menus.keys())
        self.menu_combo.blockSignals(True)
        self.menu_combo.clear()
        self.menu_combo.addItems(names)
        target = self.state.cfg.DEFAULT_CLUSTER if (self.state.cfg.DEFAULT_CLUSTER and self.state.cfg.DEFAULT_CLUSTER in self.state.menus) else self.state.menu_name
        idx = self.menu_combo.findText(target)
        if idx < 0: idx = self.menu_combo.findText(self.state.menu_name)
        if idx < 0: idx = 0
        self.menu_combo.setCurrentIndex(idx)
        self.menu_combo.blockSignals(False)

    def _switch_menu(self, name: str):
        if not name or name not in self.state.menus:
            return
        if name == self.state.menu_name:
            return
        self.state.menu_name = name
        self.header.setText(f"Support Cockpit — {self.state.menu_name}")
        self.setWindowTitle(f"Support Cockpit — {self.state.menu_name}")
        self._build_buttons()
        self._log(f"MENU: Switched to '{name}'")

    def place_command_center(self):
        r=primary_rect_logical()
        w=min(640,max(440,int(r.width()*0.33))); h=min(520,max(380,int(r.height()*0.45)))
        m=max(24,int(0.02*r.width())); x=r.right()-w-m; y=r.top()+m
        x,y,w,h = clamp_rect_to_primary(x,y,w,h)
        self.setGeometry(x,y,w,h); self.showNormal(); self.raise_(); self.activateWindow()
        QTimer.singleShot(150, lambda:(self.raise_(), self.activateWindow()))
        self._append_html(self._fmt_line(f"Cockpit @ {w}x{h} ({x},{y})", "#9fa7ad"))

    def _log(self, s: str, level: str = "INFO", **ev):
        # Always write to file (unchanged logging behavior)
        self.state.logger.text(s)
        
        # Check if we should show this in console based on log level
        console_level = LogLevel.from_string(getattr(self.cfg, "CONSOLE_LOG_LEVEL", "INFO"))
        msg_level = LogLevel.from_string(level)
        
        if msg_level >= console_level:
            line = f"[{_ts()}] {s}"
            color = {
                "DEBUG": "#6c757d",    # Gray
                "INFO": "#bfe0bf",     # Green (original)
                "WARNING": "#ffc107",  # Yellow
                "ERROR": "#dc3545"     # Red
            }.get(level.upper(), "#bfe0bf")
            self._append_html(self._fmt_line(line, color))

def _fatal(msg: str):
    try:
        app = QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(None, "Support Cockpit – Error", msg)
    except Exception:
        pass
    print(f"[{_ts()}] FATAL: {msg}")
    sys.exit(2)

def main():
    enable_dpi_awareness()
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    cfg_path = _discover_config_path(sys.argv[1:], base_dir)
    if not cfg_path:
        _fatal("No configuration found.")

    try:
        j=_load_json(cfg_path)
    except Exception as e:
        _fatal(f"Configuration could not be read:\n{e}")

    cfg = apply_json_overrides(Config(), j)
    menus = extract_menus(j, cfg)
    if not menus:
        _fatal("Configuration contains no TOOLS.")

    start_menu = cfg.DEFAULT_CLUSTER if (cfg.DEFAULT_CLUSTER and cfg.DEFAULT_CLUSTER in menus) else list(menus.keys())[0]
    app=QApplication(sys.argv)
    state=AppState(cfg, menus, start_menu)
    cc=CommandCenter(state); cc.show()
    sys.exit(app.exec())

if __name__=="__main__":
    main()