#!/usr/bin/env python3
"""
Admin-Z — System administration dashboard for Windows 10/11.

Requests administrator privileges on startup (UAC). If elevation is denied,
the app shows an error box and exits.

Dependencies:
    pip install PyQt6 psutil pywin32
Optional:
    pip install PyYAML     (nicer YAML export; a built-in fallback is used otherwise)

Run:
    python admin-z.py
"""

import ctypes
import json
import os
import platform
import re
import shutil
import socket
import subprocess
import sys
import time
from collections import deque
from datetime import datetime

APP_NAME = "Admin-Z"
APP_VERSION = "1.0.0"
GITHUB_URL = "https://github.com/admin-z/admin-z"
WINDOW_W, WINDOW_H = 1280, 960
GRAPH_POINTS = 60  # points kept per line graph

IS_WINDOWS = sys.platform == "win32"


def _fatal_box(text):
    """Native error box that works before Qt is available."""
    if IS_WINDOWS:
        ctypes.windll.user32.MessageBoxW(None, text, APP_NAME, 0x10)  # MB_ICONERROR
    else:
        print(text, file=sys.stderr)
    sys.exit(1)


if not IS_WINDOWS:
    _fatal_box("Admin-Z only runs on Windows 10/11.")

# ---------------------------------------------------------------- elevation

def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def ensure_admin():
    """Relaunch elevated via UAC; error box + exit if the user declines."""
    if is_admin():
        return
    if "--elevated" in sys.argv:
        # We already tried to elevate and still are not admin.
        _fatal_box("Admin-Z requires administrative privileges to run.")
    if getattr(sys, "frozen", False):
        exe = sys.executable
        args = " ".join(f'"{a}"' for a in sys.argv[1:] + ["--elevated"])
    else:
        exe = sys.executable
        script = os.path.abspath(sys.argv[0])
        args = " ".join(f'"{a}"' for a in [script] + sys.argv[1:] + ["--elevated"])
    rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, args, None, 1)
    if rc <= 32:  # UAC declined or elevation failed
        _fatal_box("Admin-Z requires administrative privileges to run.")
    sys.exit(0)  # elevated instance has taken over


# ---------------------------------------------------------------- imports that need install

try:
    from PyQt6.QtCore import (Qt, QTimer, QThread, pyqtSignal, QUrl, QSize,
                              QRectF, QPointF)
    from PyQt6.QtGui import (QPainter, QPen, QBrush, QColor, QFont, QPixmap,
                             QIcon, QIntValidator, QDesktopServices,
                             QPainterPath, QLinearGradient)
    from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QTabWidget,
                                 QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
                                 QPushButton, QComboBox, QLineEdit, QCheckBox,
                                 QListWidget, QListWidgetItem, QStackedWidget,
                                 QTableWidget, QTableWidgetItem, QHeaderView,
                                 QFrame, QScrollArea, QFileDialog, QMessageBox,
                                 QSizePolicy, QAbstractItemView, QStyle)
    from PyQt6.QtSvg import QSvgRenderer
except ImportError:
    _fatal_box("PyQt6 is not installed.\n\nInstall dependencies with:\n"
               "pip install PyQt6 psutil pywin32")

try:
    import psutil
except ImportError:
    _fatal_box("psutil is not installed.\n\nInstall dependencies with:\n"
               "pip install PyQt6 psutil pywin32")

HAS_PDH = HAS_EVT = HAS_WMI = False
try:
    import win32pdh
    HAS_PDH = True
except ImportError:
    pass
try:
    import win32evtlog
    HAS_EVT = True
except ImportError:
    pass
try:
    import win32com.client
    import pythoncom
    HAS_WMI = True
except ImportError:
    pass

try:
    import yaml as _pyyaml
except ImportError:
    _pyyaml = None

# ---------------------------------------------------------------- settings

def app_dir():
    """Directory of the EXE (frozen) or of this script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


SETTINGS_PATH = os.path.join(app_dir(), "settings.json")

DEFAULT_SETTINGS = {
    "theme": "light",            # "light" | "dark"
    "follow_system": True,
    "refresh_ms": 1000,
    "export_dir": os.path.join(os.path.expanduser("~"), "Desktop"),
    "always_on_top": False,
}


def load_settings():
    s = dict(DEFAULT_SETTINGS)
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            for k in DEFAULT_SETTINGS:
                if k in data:
                    s[k] = data[k]
            if "refresh_ms" not in data and "refresh_seconds" in data:
                try:  # migrate pre-1.0 settings expressed in seconds
                    s["refresh_ms"] = int(data["refresh_seconds"]) * 1000
                except Exception:
                    pass
    except Exception:
        pass
    if s["theme"] not in ("light", "dark"):
        s["theme"] = "light"
    try:
        s["refresh_ms"] = max(100, int(s["refresh_ms"]))
    except Exception:
        s["refresh_ms"] = 1000
    if not isinstance(s["export_dir"], str) or not s["export_dir"]:
        s["export_dir"] = DEFAULT_SETTINGS["export_dir"]
    s["follow_system"] = bool(s["follow_system"])
    s["always_on_top"] = bool(s["always_on_top"])
    return s


def save_settings(s):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(s, f, indent=2)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------- themes

LIGHT = {
    "bg": "#f4f5f7", "card": "#ffffff", "border": "#e2e4e8",
    "text": "#1b1e23", "dim": "#6b7280", "accent": "#0078d4",
    "accent_hover": "#106ebe", "accent_press": "#005a9e",
    "input": "#ffffff", "hover": "#eceef1", "sel": "#dbeafe",
    "graph_line": "#0078d4", "graph_fill": "#0078d4", "grid": "#e7e9ed",
    "ok": "#16a34a", "warn": "#d97706", "err": "#dc2626", "crit": "#7f1d1d",
    "logo": "#0078d4", "header": "#f0f1f4",
}
DARK = {
    "bg": "#16181d", "card": "#1f2229", "border": "#2c3038",
    "text": "#e8eaed", "dim": "#9aa0a8", "accent": "#4cc2ff",
    "accent_hover": "#6fceff", "accent_press": "#2fb4f8",
    "input": "#262a32", "hover": "#262a32", "sel": "#1e3a5f",
    "graph_line": "#4cc2ff", "graph_fill": "#4cc2ff", "grid": "#2c3038",
    "ok": "#4ade80", "warn": "#fbbf24", "err": "#f87171", "crit": "#ef4444",
    "logo": "#4cc2ff", "header": "#262a32",
}

PAL = dict(LIGHT)  # current palette, mutated in place by apply_theme()


_QSS_TEMPLATE = """
* { font-family: 'Segoe UI Variable Text', 'Segoe UI', sans-serif; font-size: 10pt; }
QMainWindow, QWidget { background: @bg; color: @text; }
QTabWidget::pane { border: none; background: @bg; }
QTabBar { background: @bg; }
QTabBar::tab {
    background: transparent; color: @dim; border: none;
    padding: 9px 22px; margin: 4px 2px 0 2px;
    border-bottom: 2px solid transparent; font-size: 10.5pt;
}
QTabBar::tab:hover { color: @text; }
QTabBar::tab:selected { color: @accent; border-bottom: 2px solid @accent; font-weight: 600; }
QFrame#card { background: @card; border: 1px solid @border; border-radius: 10px; }
QLabel { background: transparent; }
QLabel#cardTitle { font-size: 11pt; font-weight: 700; color: @text; }
QLabel#dim { color: @dim; }
QLabel#big { font-size: 14pt; font-weight: 700; }
QLabel#accent { color: @accent; font-weight: 700; }
QPushButton {
    background: @accent; color: white; border: none; border-radius: 6px;
    padding: 7px 18px; font-weight: 600;
}
QPushButton:hover { background: @accent_hover; }
QPushButton:pressed { background: @accent_press; }
QPushButton:disabled { background: @border; color: @dim; }
QPushButton#flat { background: transparent; color: @text; padding: 4px; }
QPushButton#flat:hover { background: @hover; border-radius: 6px; }
QComboBox, QLineEdit {
    background: @input; color: @text; border: 1px solid @border;
    border-radius: 6px; padding: 5px 10px; min-height: 20px;
}
QComboBox:hover, QLineEdit:hover { border-color: @accent; }
QComboBox:disabled, QLineEdit:disabled { color: @dim; background: @bg; }
QComboBox::drop-down { border: none; width: 24px; }
QComboBox::down-arrow {
    image: none; border-left: 4px solid transparent; border-right: 4px solid transparent;
    border-top: 5px solid @dim; margin-right: 8px;
}
QComboBox QAbstractItemView {
    background: @card; color: @text; border: 1px solid @border;
    selection-background-color: @sel; selection-color: @text; outline: none;
}
QCheckBox { spacing: 8px; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid @border;
    border-radius: 4px; background: @input; }
QCheckBox::indicator:hover { border-color: @accent; }
QCheckBox::indicator:checked { background: @accent; border-color: @accent; }
QListWidget {
    background: @card; border: 1px solid @border; border-radius: 10px;
    outline: none; padding: 4px;
}
QListWidget::item { padding: 9px 10px; border-radius: 6px; color: @text; }
QListWidget::item:hover { background: @hover; }
QListWidget::item:selected { background: @sel; color: @text; }
QTableWidget {
    background: @card; border: 1px solid @border; border-radius: 10px;
    gridline-color: @border; outline: none;
    selection-background-color: @sel; selection-color: @text;
    alternate-background-color: @bg;
}
QTableWidget::item { padding: 2px 6px; }
QHeaderView::section {
    background: @header; color: @text; border: none;
    border-bottom: 1px solid @border; border-right: 1px solid @border;
    padding: 6px 8px; font-weight: 600;
}
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: @border; border-radius: 5px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: @dim; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 2px; }
QScrollBar::handle:horizontal { background: @border; border-radius: 5px; min-width: 30px; }
QScrollBar::add-line, QScrollBar::sub-line { height: 0; width: 0; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }
QMessageBox { background: @card; }
QToolTip { background: @card; color: @text; border: 1px solid @border; }
"""


def build_qss(p):
    qss = _QSS_TEMPLATE
    for key in sorted(p, key=len, reverse=True):
        qss = qss.replace("@" + key, p[key])
    return qss


# ---------------------------------------------------------------- formatting helpers

def fmt_bytes(n, dec=1):
    """1024-based size with GB/TB style units."""
    if n is None:
        return "—"
    n = float(n)
    for unit in ("B", "KB", "MB", "GB", "TB", "PB"):
        if n < 1024.0 or unit == "PB":
            return f"{n:.{0 if unit == 'B' else dec}f} {unit}"
        n /= 1024.0


def fmt_bps(bits):
    """Bits/sec -> kbps / mbps / gbps (auto)."""
    if bits is None:
        return "—"
    bits = max(0.0, float(bits))
    if bits >= 1e9:
        return f"{bits / 1e9:.2f} gbps"
    if bits >= 1e6:
        return f"{bits / 1e6:.2f} mbps"
    if bits >= 1e3:
        return f"{bits / 1e3:.1f} kbps"
    return f"{bits:.0f} bps"


def fmt_used_total(used, total):
    if used is None or total is None:
        return "—"
    return f"{fmt_bytes(used)} / {fmt_bytes(total)}"


def fmt_uptime(seconds):
    seconds = int(seconds)
    d, rem = divmod(seconds, 86400)
    h, rem = divmod(rem, 3600)
    m, _ = divmod(rem, 60)
    parts = []
    if d:
        parts.append(f"{d} day{'s' if d != 1 else ''}")
    if h:
        parts.append(f"{h} hour{'s' if h != 1 else ''}")
    parts.append(f"{m} min{'s' if m != 1 else ''}")
    return ", ".join(parts)


# ---------------------------------------------------------------- export serializers

def _yaml_fallback(data, indent=0):
    """Minimal YAML emitter used when PyYAML is unavailable."""
    pad = "  " * indent
    lines = []

    def scalar(v):
        if v is None:
            return "null"
        if isinstance(v, bool):
            return "true" if v else "false"
        if isinstance(v, (int, float)):
            return str(v)
        s = str(v)
        if s == "" or re.search(r"[:#\-\[\]{}&*!|>'\"%@`,\n]", s) or s != s.strip():
            return json.dumps(s)
        return s

    if isinstance(data, dict):
        if not data:
            return [pad + "{}"]
        for k, v in data.items():
            if isinstance(v, (dict, list)) and v:
                lines.append(f"{pad}{k}:")
                lines.extend(_yaml_fallback(v, indent + 1))
            elif isinstance(v, dict):
                lines.append(f"{pad}{k}: {{}}")
            elif isinstance(v, list):
                lines.append(f"{pad}{k}: []")
            else:
                lines.append(f"{pad}{k}: {scalar(v)}")
    elif isinstance(data, list):
        if not data:
            return [pad + "[]"]
        for v in data:
            if isinstance(v, (dict, list)) and v:
                sub = _yaml_fallback(v, indent + 1)
                lines.append(f"{pad}- {sub[0].lstrip()}")
                lines.extend(sub[1:])
            else:
                lines.append(f"{pad}- {scalar(v)}")
    else:
        lines.append(pad + scalar(data))
    return lines


def to_yaml(data):
    if _pyyaml is not None:
        return _pyyaml.safe_dump(data, sort_keys=False, allow_unicode=True)
    return "\n".join(_yaml_fallback(data)) + "\n"


def _xml_key(k):
    k = re.sub(r"[^A-Za-z0-9_.-]", "_", str(k))
    if not k or not (k[0].isalpha() or k[0] == "_"):
        k = "_" + k
    return k


def to_xml(data, root_tag):
    import xml.etree.ElementTree as ET

    def build(parent, obj, singular="item"):
        if isinstance(obj, dict):
            for k, v in obj.items():
                child = ET.SubElement(parent, _xml_key(k))
                build(child, v, singular=_xml_key(k).rstrip("s") or "item")
        elif isinstance(obj, list):
            for v in obj:
                child = ET.SubElement(parent, singular)
                build(child, v)
        else:
            parent.text = "" if obj is None else str(obj)

    root = ET.Element(_xml_key(root_tag))
    build(root, data)
    ET.indent(root, space="  ")
    return ('<?xml version="1.0" encoding="utf-8"?>\n'
            + ET.tostring(root, encoding="unicode") + "\n")


def write_export(data, root_tag, fmt, directory, basename):
    """Serialize `data` as JSON/XML/YAML into directory. Returns filename."""
    fmt = fmt.lower()
    ext = {"json": "json", "xml": "xml", "yaml": "yaml"}[fmt]
    filename = f"{basename}.{ext}"
    path = os.path.join(directory, filename)
    if fmt == "json":
        text = json.dumps(data, indent=2, ensure_ascii=False)
    elif fmt == "xml":
        text = to_xml(data, root_tag)
    else:
        text = to_yaml(data)
    os.makedirs(directory, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return filename


# ================================================================ data collection

def _wmi_connect(namespace="root\\cimv2"):
    moniker = "winmgmts:{impersonationLevel=impersonate}!\\\\.\\" + namespace
    return win32com.client.GetObject(moniker)


class PdhSampler:
    """Thin wrapper over a persistent PDH query (locale-independent paths)."""

    def __init__(self):
        self.query = win32pdh.OpenQuery()
        self.counters = {}

    def add(self, key, path):
        try:
            try:
                h = win32pdh.AddEnglishCounter(self.query, path)
            except AttributeError:
                h = win32pdh.AddCounter(self.query, path)
            self.counters[key] = h
        except Exception:
            pass

    def collect(self):
        try:
            win32pdh.CollectQueryData(self.query)
        except Exception:
            pass

    def value(self, key):
        h = self.counters.get(key)
        if h is None:
            return None
        try:
            _, v = win32pdh.GetFormattedCounterValue(h, win32pdh.PDH_FMT_DOUBLE)
            return v
        except Exception:
            return None

    def close(self):
        try:
            win32pdh.CloseQuery(self.query)
        except Exception:
            pass


def _gpu_mem_instances():
    """Instance names of the 'GPU Adapter Memory' perf object (one per adapter)."""
    try:
        _, insts = win32pdh.EnumObjectItems(None, None, "GPU Adapter Memory",
                                            win32pdh.PERF_DETAIL_WIZARD)
        return sorted(set(insts))
    except Exception:
        return []


FORM_FACTORS = {0: "Unknown", 1: "Other", 2: "SIP", 3: "DIP", 4: "ZIP", 5: "SOJ",
                6: "Proprietary", 7: "SIMM", 8: "DIMM", 9: "TSOP", 10: "PGA",
                11: "RIMM", 12: "SODIMM", 13: "SRIMM", 14: "SMD", 15: "SSMP",
                16: "QFP", 17: "TQFP", 18: "SOIC", 19: "LCC", 20: "PLCC",
                21: "BGA", 22: "FPBGA", 23: "LGA"}
MEM_TYPES = {20: "DDR", 21: "DDR2", 24: "DDR3", 26: "DDR4", 34: "DDR5"}
MEDIA_TYPES = {3: "HDD", 4: "SSD", 5: "SCM"}
BUS_TYPES = {1: "SCSI", 2: "ATAPI", 3: "ATA", 4: "IEEE 1394", 5: "SSA",
             6: "Fibre Channel", 7: "USB", 8: "RAID", 9: "iSCSI", 10: "SAS",
             11: "SATA", 12: "SD", 13: "MMC", 16: "Spaces", 17: "NVMe"}


def _read_os_info():
    info = {}
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                             r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")

        def rv(name):
            try:
                return winreg.QueryValueEx(key, name)[0]
            except OSError:
                return None

        product = rv("ProductName") or "Windows"
        build = int(rv("CurrentBuildNumber") or 0)
        ubr = rv("UBR")
        if build >= 22000:  # registry still reports "Windows 10" on Win11
            product = product.replace("Windows 10", "Windows 11")
        info["name"] = product
        info["version"] = rv("DisplayVersion") or rv("ReleaseId") or ""
        info["build"] = f"{build}.{ubr}" if ubr is not None else str(build)
        winreg.CloseKey(key)
    except Exception:
        info["name"] = f"{platform.system()} {platform.release()}"
        info["version"] = platform.version()
        info["build"] = platform.version()
    info["kernel"] = f"Windows NT {platform.version()}"
    info["arch"] = platform.machine()
    info["hostname"] = socket.gethostname()
    info["user"] = os.environ.get("USERNAME", "user")
    return info


def _dedicated_vram_from_registry():
    """GPU name -> dedicated VRAM bytes (HardwareInformation.qwMemorySize)."""
    out = {}
    try:
        import winreg
        base = r"SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}"
        root = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base)
        i = 0
        while True:
            try:
                sub = winreg.EnumKey(root, i)
            except OSError:
                break
            i += 1
            if not re.fullmatch(r"\d{4}", sub):
                continue
            try:
                k = winreg.OpenKey(root, sub)
                desc = winreg.QueryValueEx(k, "DriverDesc")[0]
                size = winreg.QueryValueEx(k, "HardwareInformation.qwMemorySize")[0]
                if isinstance(size, int) and size > 0:
                    out[str(desc)] = size
                winreg.CloseKey(k)
            except OSError:
                continue
        winreg.CloseKey(root)
    except Exception:
        pass
    return out


class SpecsThread(QThread):
    """One-shot WMI/registry sweep for slow-changing hardware specifications."""
    ready = pyqtSignal(dict)

    def run(self):
        specs = {"os": _read_os_info(), "board": None, "cpu": {}, "ram": {},
                 "gpus": [], "disks": [], "disk_map": {}, "net": {}}
        svc = None
        if HAS_WMI:
            try:
                pythoncom.CoInitialize()
                svc = _wmi_connect()
            except Exception:
                svc = None
        if svc is not None:
            try:  # motherboard
                for b in svc.ExecQuery("SELECT Manufacturer, Product FROM Win32_BaseBoard"):
                    specs["board"] = f"{b.Manufacturer} {b.Product}".strip()
                    break
            except Exception:
                pass
            try:  # CPU
                for c in svc.ExecQuery("SELECT * FROM Win32_Processor"):
                    specs["cpu"] = {
                        "name": (c.Name or "").strip(),
                        "base_mhz": int(c.MaxClockSpeed or 0),
                        "cores": int(c.NumberOfCores or 0),
                        "logical": int(c.NumberOfLogicalProcessors or 0),
                        "socket": c.SocketDesignation,
                    }
                    break
            except Exception:
                pass
            try:  # RAM modules
                mods, speed, ff, mtype = [], None, None, None
                for m in svc.ExecQuery("SELECT * FROM Win32_PhysicalMemory"):
                    cap = int(m.Capacity or 0)
                    spd = None
                    try:
                        spd = int(m.ConfiguredClockSpeed or 0) or int(m.Speed or 0)
                    except Exception:
                        pass
                    f = FORM_FACTORS.get(int(m.FormFactor or 0), "Unknown")
                    try:
                        mtype = mtype or MEM_TYPES.get(int(m.SMBIOSMemoryType or 0))
                    except Exception:
                        pass
                    speed = speed or spd
                    ff = ff or f
                    mods.append({"capacity": cap, "speed_mhz": spd, "form_factor": f,
                                 "manufacturer": (m.Manufacturer or "").strip(),
                                 "part_number": (m.PartNumber or "").strip()})
                specs["ram"] = {"modules": mods, "speed_mhz": speed,
                                "form_factor": ff, "type": mtype,
                                "slots_used": len(mods)}
            except Exception:
                pass
            try:  # GPUs
                vram_reg = _dedicated_vram_from_registry()
                gpus = []
                for g in svc.ExecQuery("SELECT * FROM Win32_VideoController"):
                    name = (g.Name or "").strip()
                    if not name:
                        continue
                    dedicated = vram_reg.get(name)
                    if dedicated is None:
                        try:
                            dedicated = int(g.AdapterRAM or 0)
                        except Exception:
                            dedicated = 0
                    gpus.append({"name": name,
                                 "dedicated_bytes": max(0, dedicated),
                                 "driver": g.DriverVersion,
                                 "mode": g.VideoModeDescription})
                # drop the fallback software adapter if a real one exists
                real = [g for g in gpus if "microsoft basic display" not in g["name"].lower()]
                specs["gpus"] = real or gpus
            except Exception:
                pass
            try:  # physical disks + logical-drive mapping
                media, bus = {}, {}
                try:
                    stor = _wmi_connect(r"root\Microsoft\Windows\Storage")
                    for pd in stor.ExecQuery("SELECT DeviceId, MediaType, BusType FROM MSFT_PhysicalDisk"):
                        idx = int(pd.DeviceId)
                        media[idx] = MEDIA_TYPES.get(int(pd.MediaType or 0))
                        bus[idx] = BUS_TYPES.get(int(pd.BusType or 0))
                except Exception:
                    pass
                disks = []
                for d in svc.ExecQuery("SELECT * FROM Win32_DiskDrive"):
                    idx = int(d.Index)
                    dtype = media.get(idx)
                    if not dtype:
                        mt = (d.MediaType or "")
                        dtype = "HDD" if "fixed" in mt.lower() else (mt or "Disk")
                    if bus.get(idx):
                        dtype = f"{dtype} ({bus[idx]})"
                    disks.append({"index": idx, "model": (d.Model or "").strip(),
                                  "type": dtype, "size_bytes": int(d.Size or 0),
                                  "interface": d.InterfaceType,
                                  "serial": (d.SerialNumber or "").strip(),
                                  "partitions": int(d.Partitions or 0)})
                    mounts = []
                    try:
                        dev = d.DeviceID.replace("\\", "\\\\")
                        q1 = ("ASSOCIATORS OF {Win32_DiskDrive.DeviceID='%s'} "
                              "WHERE AssocClass=Win32_DiskDriveToDiskPartition" % dev)
                        for part in svc.ExecQuery(q1):
                            q2 = ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='%s'} "
                                  "WHERE AssocClass=Win32_LogicalDiskToPartition" % part.DeviceID)
                            for ld in svc.ExecQuery(q2):
                                mounts.append(ld.DeviceID + "\\")
                    except Exception:
                        pass
                    specs["disk_map"][idx] = mounts
                specs["disks"] = sorted(disks, key=lambda x: x["index"])
            except Exception:
                pass
            try:  # network adapter details keyed by connection name
                for n in svc.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID IS NOT NULL"):
                    specs["net"][str(n.NetConnectionID)] = {
                        "description": (n.Name or "").strip(),
                        "mac": n.MACAddress,
                        "adapter_type": n.AdapterType,
                        "manufacturer": n.Manufacturer,
                    }
            except Exception:
                pass
        if not specs["cpu"]:
            specs["cpu"] = {"name": platform.processor() or "CPU",
                            "base_mhz": int(getattr(psutil.cpu_freq(), "max", 0) or 0),
                            "cores": psutil.cpu_count(logical=False) or 0,
                            "logical": psutil.cpu_count() or 0, "socket": None}
        self.ready.emit(specs)


class MetricsThread(QThread):
    """Background sampler for everything that changes per refresh tick."""
    sample = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.interval = 1000        # milliseconds
        self._stop = False
        self.disk_map = {}          # physical index -> [mountpoints]
        self._cpu_temp_ok = True
        self._gpu_temp_ok = shutil.which("nvidia-smi") is not None
        self._last_temps = (None, [])
        self._last_temp_t = 0.0
        self._wmi_thermal = None
        self._pnames = {}           # pid -> process name cache

    def stop(self):
        self._stop = True

    def set_disk_map(self, m):
        self.disk_map = dict(m)

    # ---- temperature (sampled every 5th tick; disabled after first failure)
    def _cpu_temp(self):
        if not (HAS_WMI and self._cpu_temp_ok):
            return None
        try:
            if self._wmi_thermal is None:
                self._wmi_thermal = _wmi_connect(r"root\wmi")
            temps = []
            for t in self._wmi_thermal.ExecQuery(
                    "SELECT CurrentTemperature FROM MSAcpi_ThermalZoneTemperature"):
                temps.append(int(t.CurrentTemperature) / 10.0 - 273.15)
            temps = [t for t in temps if -20 < t < 150]
            if not temps:
                raise ValueError
            return max(temps)
        except Exception:
            self._cpu_temp_ok = False
            return None

    def _gpu_temps(self):
        if not self._gpu_temp_ok:
            return []
        try:
            out = subprocess.run(
                ["nvidia-smi", "--query-gpu=temperature.gpu",
                 "--format=csv,noheader,nounits"],
                capture_output=True, text=True, timeout=3,
                creationflags=0x08000000)  # CREATE_NO_WINDOW
            if out.returncode != 0:
                raise RuntimeError
            return [int(x) for x in out.stdout.split() if x.strip().isdigit()]
        except Exception:
            self._gpu_temp_ok = False
            return []

    def _connections(self):
        """TCP snapshot: established connections and listening ports."""
        est, lst = [], []
        try:
            conns = psutil.net_connections(kind="inet")
        except Exception:
            return est, lst
        if len(self._pnames) > 4096:
            self._pnames.clear()

        def pname(pid):
            if pid in (None, 0):
                return "—"
            if pid == 4:
                return "System"
            if pid not in self._pnames:
                try:
                    self._pnames[pid] = psutil.Process(pid).name()
                except Exception:
                    self._pnames[pid] = f"PID {pid}"
            return self._pnames[pid]

        for c in conns:
            try:
                if c.status == psutil.CONN_ESTABLISHED and c.raddr:
                    est.append({"ip": c.laddr.ip, "port": c.laddr.port,
                                "rip": c.raddr.ip, "rport": c.raddr.port,
                                "pid": c.pid, "pname": pname(c.pid)})
                elif c.status == psutil.CONN_LISTEN:
                    lst.append({"ip": c.laddr.ip, "port": c.laddr.port,
                                "pid": c.pid, "pname": pname(c.pid)})
            except Exception:
                continue
        est.sort(key=lambda x: (x["rip"], x["rport"]))
        lst.sort(key=lambda x: x["port"])
        return est, lst

    def run(self):
        if HAS_WMI:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
        pdh = None
        gpu_insts = []
        if HAS_PDH:
            try:
                pdh = PdhSampler()
                pdh.add("procs", r"\System\Processes")
                pdh.add("threads", r"\System\Threads")
                pdh.add("handles", r"\Process(_Total)\Handle Count")
                pdh.add("wset", r"\Process(_Total)\Working Set")
                pdh.add("perf", r"\Processor Information(_Total)\% Processor Performance")
                gpu_insts = _gpu_mem_instances()
                for i, inst in enumerate(gpu_insts):
                    pdh.add(f"g{i}d", rf"\GPU Adapter Memory({inst})\Dedicated Usage")
                    pdh.add(f"g{i}s", rf"\GPU Adapter Memory({inst})\Shared Usage")
                pdh.collect()
                time.sleep(0.1)
                pdh.collect()
            except Exception:
                pdh = None
        psutil.cpu_percent(None)  # prime
        base_mhz = 0.0
        try:
            f = psutil.cpu_freq()
            base_mhz = float(f.max or f.current or 0)
        except Exception:
            pass
        prev_disk = psutil.disk_io_counters(perdisk=True)
        prev_net = psutil.net_io_counters(pernic=True)
        prev_t = time.time()

        while not self._stop:
            target = max(100, int(self.interval)) / 1000.0
            slept = 0.0
            while slept < target and not self._stop:
                time.sleep(0.05)
                slept += 0.05
            if self._stop:
                break
            now = time.time()
            dt = max(0.001, now - prev_t)
            if pdh:
                pdh.collect()
            d = {"cpu": {}, "mem": {}, "disks": {}, "gpus": [], "nets": {}}

            # ---- CPU
            cpu = d["cpu"]
            cpu["percent"] = psutil.cpu_percent(None)
            ghz = None
            if pdh:
                perf = pdh.value("perf")
                if perf and base_mhz:
                    ghz = base_mhz * perf / 100.0 / 1000.0
            if ghz is None:
                try:
                    ghz = (psutil.cpu_freq().current or 0) / 1000.0
                except Exception:
                    ghz = 0.0
            cpu["ghz"] = ghz
            cpu["processes"] = int(pdh.value("procs") or 0) if pdh else len(psutil.pids())
            cpu["threads"] = int(pdh.value("threads")) if pdh and pdh.value("threads") else None
            cpu["handles"] = int(pdh.value("handles")) if pdh and pdh.value("handles") else None

            # ---- memory
            vm = psutil.virtual_memory()
            sw = psutil.swap_memory()
            d["mem"] = {"used": vm.used, "total": vm.total, "percent": vm.percent,
                        "page_used": sw.used, "page_total": sw.total,
                        "working_set": int(pdh.value("wset")) if pdh and pdh.value("wset") else None}

            # ---- disks (throughput + capacity per physical disk)
            cur_disk = psutil.disk_io_counters(perdisk=True)
            for key, io in cur_disk.items():
                m = re.search(r"(\d+)$", key)
                if not m:
                    continue
                idx = int(m.group(1))
                p = prev_disk.get(key)
                rd = (io.read_bytes - p.read_bytes) / dt if p else 0.0
                wr = (io.write_bytes - p.write_bytes) / dt if p else 0.0
                used = total = None
                mounts = self.disk_map.get(idx) or []
                if mounts:
                    used = total = 0
                    for mp in mounts:
                        try:
                            u = psutil.disk_usage(mp)
                            used += u.used
                            total += u.total
                        except Exception:
                            pass
                d["disks"][idx] = {"read_bps": max(0.0, rd), "write_bps": max(0.0, wr),
                                   "used": used, "total": total}
            prev_disk = cur_disk

            # ---- GPU adapter memory usage
            for i in range(len(gpu_insts)):
                ded = pdh.value(f"g{i}d") if pdh else None
                sha = pdh.value(f"g{i}s") if pdh else None
                d["gpus"].append({"dedicated_used": int(ded) if ded is not None else None,
                                  "shared_used": int(sha) if sha is not None else None})

            # ---- network
            cur_net = psutil.net_io_counters(pernic=True)
            try:
                stats = psutil.net_if_stats()
                addrs = psutil.net_if_addrs()
            except Exception:
                stats, addrs = {}, {}
            for name, st in stats.items():
                if not st.isup:
                    continue
                low = name.lower()
                if "loopback" in low or low.startswith("lo"):
                    continue
                ip4 = ip6 = None
                for a in addrs.get(name, []):
                    if a.family == socket.AF_INET and not ip4:
                        ip4 = a.address
                    elif a.family == socket.AF_INET6 and not ip6:
                        ip6 = a.address.split("%")[0]
                io, p = cur_net.get(name), prev_net.get(name)
                up = (io.bytes_sent - p.bytes_sent) * 8 / dt if io and p else 0.0
                dn = (io.bytes_recv - p.bytes_recv) * 8 / dt if io and p else 0.0
                d["nets"][name] = {"ip4": ip4, "ip6": ip6,
                                   "up_bps": max(0.0, up), "down_bps": max(0.0, dn),
                                   "speed_mbps": st.speed}
            prev_net = cur_net
            prev_t = now

            # ---- TCP connections (established + listening, per refresh)
            est, lst = self._connections()
            d["conns"] = {"est": est, "lst": lst}

            # ---- temperatures (slow path, at most every 5 s)
            if now - self._last_temp_t >= 5.0:
                self._last_temp_t = now
                self._last_temps = (self._cpu_temp(), self._gpu_temps())
            d["cpu_temp"], d["gpu_temps"] = self._last_temps

            d["uptime"] = time.time() - psutil.boot_time()
            self.sample.emit(d)
        if pdh:
            pdh.close()


class LogsThread(QThread):
    """Loads the newest 1000 events from Application / Setup / System."""
    loaded = pyqtSignal(list)
    LEVELS = {1: "Critical", 2: "Error", 3: "Warning", 4: "Information",
              5: "Verbose", 0: "Information"}

    def run(self):
        events = []
        if HAS_EVT:
            import xml.etree.ElementTree as ET
            ns = {"e": "http://schemas.microsoft.com/win/2004/08/events/event"}
            for channel in ("Application", "Setup", "System"):
                try:
                    q = win32evtlog.EvtQuery(
                        channel,
                        win32evtlog.EvtQueryChannelPath | win32evtlog.EvtQueryReverseDirection)
                except Exception:
                    continue
                got = 0
                while got < 1000:
                    try:
                        batch = win32evtlog.EvtNext(q, min(100, 1000 - got))
                    except Exception:
                        break
                    if not batch:
                        break
                    for h in batch:
                        got += 1
                        try:
                            xml_text = win32evtlog.EvtRender(
                                h, win32evtlog.EvtRenderEventXml)
                            root = ET.fromstring(xml_text)  # OS-generated XML
                            sysn = root.find("e:System", ns)
                            tc = sysn.find("e:TimeCreated", ns)
                            ts = (tc.get("SystemTime") if tc is not None else "") or ""
                            ts = re.sub(r"\.(\d{1,6})\d*", r".\1", ts).replace("Z", "+00:00")
                            try:
                                dt = datetime.fromisoformat(ts).astimezone()
                            except ValueError:
                                dt = datetime.now().astimezone()
                            lvl_el = sysn.find("e:Level", ns)
                            lvl = int(lvl_el.text) if lvl_el is not None and lvl_el.text else 0
                            prov = sysn.find("e:Provider", ns)
                            source = (prov.get("Name") if prov is not None else "") or "?"
                            eid_el = sysn.find("e:EventID", ns)
                            eid = int(eid_el.text) if eid_el is not None and eid_el.text else 0
                            task_el = sysn.find("e:Task", ns)
                            task = task_el.text if task_el is not None else None
                            task = "None" if task in (None, "0") else str(task)
                            events.append({
                                "channel": channel,
                                "level_num": lvl,
                                "level": self.LEVELS.get(lvl, str(lvl)),
                                "dt": dt,
                                "time_str": dt.strftime("%Y-%m-%d %H:%M:%S"),
                                "source": source.replace("Microsoft-Windows-", ""),
                                "event_id": eid,
                                "task": task,
                            })
                        except Exception:
                            continue
        events.sort(key=lambda e: e["dt"], reverse=True)
        self.loaded.emit(events[:1000])


# ================================================================ UI primitives

def _nice_ceiling(v):
    """Smallest 1/2/5 * 10^k that is >= v (for graph autoscaling)."""
    if v <= 0:
        return 1.0
    import math
    exp = math.floor(math.log10(v))
    for mult in (1, 2, 5, 10):
        cand = mult * (10 ** exp)
        if cand >= v:
            return float(cand)
    return float(10 ** (exp + 1))


class LineGraph(QWidget):
    """Minimal sliding line chart (Task-Manager style, newest at right)."""

    def __init__(self, fixed_max=100.0, formatter=None, parent=None):
        super().__init__(parent)
        self.fixed_max = fixed_max          # None -> autoscale
        self.formatter = formatter or (lambda v: f"{v:.0f}%")
        self.points = deque(maxlen=GRAPH_POINTS)
        self.setMinimumHeight(84)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def add_point(self, v):
        self.points.append(max(0.0, float(v or 0.0)))
        self.update()

    def paintEvent(self, ev):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        path = QPainterPath()
        path.addRoundedRect(r, 6.0, 6.0)
        p.setClipPath(path)
        p.fillRect(r, QColor(PAL["bg"]))
        grid = QColor(PAL["grid"])
        p.setPen(QPen(grid, 1))
        for frac in (0.25, 0.5, 0.75):
            y = r.top() + r.height() * frac
            p.drawLine(QPointF(r.left(), y), QPointF(r.right(), y))

        if self.points:
            maxv = self.fixed_max or _nice_ceiling(max(max(self.points), 1e-9) * 1.15)
            n = len(self.points)
            step = r.width() / max(1, GRAPH_POINTS - 1)
            pts = []
            for i, v in enumerate(self.points):
                x = r.right() - (n - 1 - i) * step
                y = r.bottom() - min(1.0, v / maxv) * (r.height() - 6) - 2
                pts.append(QPointF(x, y))
            line_c = QColor(PAL["graph_line"])
            fill = QPainterPath()
            fill.moveTo(pts[0].x(), r.bottom())
            for pt in pts:
                fill.lineTo(pt)
            fill.lineTo(pts[-1].x(), r.bottom())
            fill.closeSubpath()
            fc = QColor(PAL["graph_fill"])
            fc.setAlpha(42)
            p.fillPath(fill, fc)
            p.setPen(QPen(line_c, 1.6))
            for i in range(1, len(pts)):
                p.drawLine(pts[i - 1], pts[i])
            # current value label
            p.setPen(QColor(PAL["dim"]))
            f = p.font()
            f.setPointSizeF(8.5)
            p.setFont(f)
            p.drawText(QRectF(r.left() + 8, r.top() + 4, r.width() - 16, 16),
                       Qt.AlignmentFlag.AlignRight,
                       self.formatter(self.points[-1]))
        p.setClipping(False)
        p.setPen(QPen(QColor(PAL["border"]), 1))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRoundedRect(r, 6.0, 6.0)
        p.end()


class Card(QFrame):
    """Rounded section box with an optional bold title."""

    def __init__(self, title=None, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.v = QVBoxLayout(self)
        self.v.setContentsMargins(16, 13, 16, 14)
        self.v.setSpacing(8)
        if title:
            t = QLabel(title)
            t.setObjectName("cardTitle")
            self.v.addWidget(t)


class InfoGrid(QWidget):
    """Two-column 'Label     value' rows used inside cards."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.g = QGridLayout(self)
        self.g.setContentsMargins(0, 0, 0, 0)
        self.g.setHorizontalSpacing(16)
        self.g.setVerticalSpacing(5)
        self.g.setColumnStretch(1, 1)
        self._rows = {}

    def add(self, label, value="—"):
        r = self.g.rowCount()
        k = QLabel(label)
        k.setObjectName("dim")
        v = QLabel(value)
        v.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        v.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        v.setWordWrap(True)
        self.g.addWidget(k, r, 0)
        self.g.addWidget(v, r, 1)
        self._rows[label] = (k, v)
        return v

    def set(self, label, text):
        if label in self._rows:
            self._rows[label][1].setText(text)

    def set_visible(self, label, vis):
        if label in self._rows:
            for w in self._rows[label]:
                w.setVisible(vis)


class WindowsLogoWidget(QWidget):
    """Flat four-pane Windows logo, colored from the active palette."""

    def __init__(self, size=230, parent=None):
        super().__init__(parent)
        self.setFixedSize(size, size)

    def paintEvent(self, ev):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        w = self.width()
        gap = max(6, int(w * 0.045))
        pane = (w - gap) / 2
        c = QColor(PAL["logo"])
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QBrush(c))
        r = max(3.0, pane * 0.045)
        for (x, y) in ((0, 0), (pane + gap, 0), (0, pane + gap),
                       (pane + gap, pane + gap)):
            p.drawRoundedRect(QRectF(x, y, pane, pane), r, r)
        p.end()


class ColorStrip(QWidget):
    """Neofetch-style palette blocks."""

    KEYS = ("accent", "ok", "warn", "err", "logo", "dim", "sel", "border")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(20)
        self.setFixedWidth((18 + 6) * len(self.KEYS))

    def paintEvent(self, ev):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        x = 0
        for k in self.KEYS:
            p.setBrush(QColor(PAL[k]))
            p.drawRoundedRect(QRectF(x, 1, 18, 18), 4, 4)
            x += 24
        p.end()


GITHUB_PATH = ("M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17"
               ".55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-"
               ".23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1"
               ".23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-"
               "3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2"
               ".12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27s1.36.09 2 .27c1."
               "53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27."
               "82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07"
               "-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4."
               "42-3.58-8-8-8z")


class GitHubButton(QPushButton):
    """Flat GitHub-mark button that opens the project repository."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("flat")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip(GITHUB_URL)
        self.setFixedSize(44, 44)
        self.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(GITHUB_URL)))
        self.refresh_icon()

    def refresh_icon(self):
        svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">'
               f'<path fill="{PAL["text"]}" d="{GITHUB_PATH}"/></svg>')
        renderer = QSvgRenderer(svg.encode("utf-8"))
        pix = QPixmap(56, 56)
        pix.fill(Qt.GlobalColor.transparent)
        p = QPainter(pix)
        renderer.render(p)
        p.end()
        self.setIcon(QIcon(pix))
        self.setIconSize(QSize(28, 28))


def launch_windows_tool(parent, *cmd):
    """Start a Windows built-in tool (resolved from System32 when possible)."""
    try:
        sys32 = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32")
        exe = os.path.join(sys32, cmd[0])
        if not os.path.exists(exe):
            exe = cmd[0]
        subprocess.Popen([exe, *cmd[1:]])
    except Exception as e:
        QMessageBox.critical(parent, APP_NAME, f"Could not open {cmd[0]}:\n{e}")


class ExportBar(QWidget):
    """'Format [combo]  [Export]' toolbar row for spec exports."""

    def __init__(self, what, basename, root_tag, get_data, get_dir,
                 extra=None, parent=None):
        super().__init__(parent)
        self.what, self.basename, self.root_tag = what, basename, root_tag
        self.get_data, self.get_dir = get_data, get_dir
        h = QHBoxLayout(self)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(10)
        h.addStretch(1)
        lab = QLabel("Export format")
        lab.setObjectName("dim")
        h.addWidget(lab)
        self.combo = QComboBox()
        self.combo.addItems(["JSON", "XML", "YAML"])
        self.combo.setFixedWidth(90)
        h.addWidget(self.combo)
        self.btn = QPushButton(f"Export {what.lower()} information")
        self.btn.clicked.connect(self._export)
        h.addWidget(self.btn)
        if extra:
            text, callback = extra
            extra_btn = QPushButton(text)
            extra_btn.clicked.connect(callback)
            h.addWidget(extra_btn)

    def _export(self):
        data = self.get_data()
        if not data:
            QMessageBox.warning(self, APP_NAME,
                                f"{self.what} information is still loading. Try again shortly.")
            return
        try:
            fn = write_export(data, self.root_tag, self.combo.currentText(),
                              self.get_dir(), self.basename)
        except Exception as e:
            QMessageBox.critical(self, APP_NAME, f"Export failed:\n{e}")
            return
        QMessageBox.information(
            self, "Export successful",
            f"{self.what} information was exported successfully as “{fn}”.")


# ================================================================ System tab

class SystemTab(QWidget):
    """Neofetch-style overview: OS logo on the left, details on the right."""

    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QHBoxLayout(self)
        outer.setContentsMargins(48, 40, 48, 40)
        outer.setSpacing(48)

        logo_col = QVBoxLayout()
        logo_col.addStretch(1)
        self.logo = WindowsLogoWidget()
        logo_col.addWidget(self.logo, 0, Qt.AlignmentFlag.AlignHCenter)
        logo_col.addStretch(1)
        outer.addLayout(logo_col, 0)

        details = QVBoxLayout()
        details.setSpacing(4)
        details.addStretch(1)          # center the list against the logo height
        self.header = QLabel("…")
        self.header.setObjectName("accent")
        f = self.header.font()
        f.setPointSizeF(13)
        f.setBold(True)
        self.header.setFont(f)
        details.addWidget(self.header)
        self.sep = QFrame()
        self.sep.setFixedHeight(1)
        self.sep.setStyleSheet(f"background: {PAL['border']}; border: none;")
        details.addWidget(self.sep)
        details.addSpacing(6)

        self.grid = QGridLayout()
        self.grid.setHorizontalSpacing(18)
        self.grid.setVerticalSpacing(6)
        self.grid.setColumnStretch(1, 1)
        details.addLayout(self.grid)
        details.addSpacing(14)
        self.strip = ColorStrip()
        details.addWidget(self.strip)
        details.addStretch(1)
        outer.addLayout(details, 1)

        self._values = {}
        self._dyn = {}

    def _add_row(self, key, value):
        r = self.grid.rowCount()
        k = QLabel(key)
        k.setObjectName("accent")
        v = QLabel(value)
        v.setWordWrap(True)
        v.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.grid.addWidget(k, r, 0, Qt.AlignmentFlag.AlignTop)
        self.grid.addWidget(v, r, 1)
        self._values[key] = v
        return v

    def set_specs(self, specs):
        osd = specs.get("os", {})
        self.header.setText(f"{osd.get('user', '?')}@{osd.get('hostname', '?')}")
        rows = [
            ("OS", f"{osd.get('name', 'Windows')} {osd.get('arch', '')}".strip()),
            ("Version", str(osd.get("version") or "—")),
            ("Build", str(osd.get("build") or "—")),
            ("Kernel", str(osd.get("kernel") or "—")),
            ("Uptime", "—"),
            ("Host", str(specs.get("board") or "—")),
            ("CPU", (specs.get("cpu") or {}).get("name") or "—"),
        ]
        for i, g in enumerate(specs.get("gpus") or []):
            rows.append((f"GPU{'' if i == 0 else ' ' + str(i + 1)}", g["name"]))
        rows += [
            ("Memory", "—"),
            ("Disk", "—"),
            ("Resolution", self._resolution()),
            ("Local IP", "—"),
        ]
        for k, v in rows:
            if k in self._values:
                self._values[k].setText(v)
            else:
                self._add_row(k, v)
        disks = specs.get("disks") or []
        if disks:
            self._dyn["disk_model"] = disks[0]["model"]

    @staticmethod
    def _resolution():
        try:
            scr = QApplication.primaryScreen()
            size = scr.size()
            dpr = scr.devicePixelRatio()
            return f"{int(size.width() * dpr)}×{int(size.height() * dpr)}"
        except Exception:
            return "—"

    def update_metrics(self, d):
        if "Uptime" in self._values:
            self._values["Uptime"].setText(fmt_uptime(d.get("uptime", 0)))
        mem = d.get("mem") or {}
        if "Memory" in self._values and mem:
            self._values["Memory"].setText(
                f"{fmt_bytes(mem['used'])} / {fmt_bytes(mem['total'])} ({mem['percent']:.0f}%)")
        disks = d.get("disks") or {}
        if "Disk" in self._values and 0 in disks and disks[0]["total"]:
            dd = disks[0]
            pct = dd["used"] / dd["total"] * 100 if dd["total"] else 0
            self._values["Disk"].setText(
                f"{fmt_bytes(dd['used'])} / {fmt_bytes(dd['total'])} ({pct:.0f}%)")
        nets = d.get("nets") or {}
        if "Local IP" in self._values:
            ip = next((n["ip4"] for n in nets.values() if n.get("ip4")), None)
            self._values["Local IP"].setText(ip or "—")

    def retheme(self):
        self.sep.setStyleSheet(f"background: {PAL['border']}; border: none;")
        self.logo.update()
        self.strip.update()


# ================================================================ Hardware tab

class HardwareTab(QWidget):
    def __init__(self, get_dir, parent=None):
        super().__init__(parent)
        self.specs = None
        v = QVBoxLayout(self)
        v.setContentsMargins(18, 14, 18, 18)
        v.setSpacing(12)
        v.addWidget(ExportBar("Hardware", "hardware_specs", "hardware",
                              self._export_data, get_dir,
                              extra=("Open Resource Monitor",
                                     lambda: launch_windows_tool(self, "resmon.exe"))))
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.content = QWidget()
        self.grid = QGridLayout(self.content)
        self.grid.setContentsMargins(0, 0, 6, 0)
        self.grid.setSpacing(12)
        self.grid.setColumnStretch(0, 1)
        self.grid.setColumnStretch(1, 1)
        self.scroll.setWidget(self.content)
        v.addWidget(self.scroll, 1)
        self.loading = QLabel("Gathering hardware information…")
        self.loading.setObjectName("dim")
        self.loading.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.grid.addWidget(self.loading, 0, 0, 1, 2)

        self.cpu = {}
        self.gpu_cards = []      # [{grid, graph, ded_total, shared_total}]
        self.ram = {}
        self.disk_cards = {}     # index -> {grid, graph, total}
        self.shared_total = psutil.virtual_memory().total // 2

    # ---- card builders -------------------------------------------------
    def _cpu_card(self, spec):
        card = Card("CPU")
        g = InfoGrid()
        g.add("Name", spec.get("name") or "—")
        g.add("Speed")
        g.add("Utilization")
        g.add("Processes")
        g.add("Threads")
        g.add("Handles")
        g.add("Temperature")
        g.set_visible("Temperature", False)   # omitted until a sensor reports
        card.v.addWidget(g)
        card.v.addStretch(1)
        graph = LineGraph(100.0, lambda v: f"{v:.0f}%")
        card.v.addWidget(graph)
        self.cpu = {"grid": g, "graph": graph}
        return card

    def _gpu_card(self, i, spec):
        card = Card(f"GPU {i}")
        g = InfoGrid()
        g.add("Name", spec["name"])
        g.add("Total GPU memory")
        g.add("Dedicated GPU memory")
        g.add("Shared GPU memory")
        g.add("Temperature")
        g.set_visible("Temperature", False)
        card.v.addWidget(g)
        card.v.addStretch(1)
        graph = LineGraph(100.0, lambda v: f"{v:.0f}%")
        card.v.addWidget(graph)
        ded_total = spec.get("dedicated_bytes") or 0
        total = ded_total + self.shared_total
        g.set("Total GPU memory", fmt_bytes(total))
        self.gpu_cards.append({"grid": g, "graph": graph,
                               "ded_total": ded_total,
                               "shared_total": self.shared_total})
        return card

    def _ram_card(self, spec):
        card = Card("RAM")
        g = InfoGrid()
        speed = spec.get("speed_mhz")
        mtype = spec.get("type")
        g.add("Speed", (f"{mtype} • " if mtype else "") + (f"{speed} MHz" if speed else "—"))
        g.add("Form factor", spec.get("form_factor") or "—")
        g.add("Capacity")
        g.add("Page file")
        g.add("Working set")
        card.v.addWidget(g)
        card.v.addStretch(1)
        graph = LineGraph(100.0, lambda v: f"{v:.0f}%")
        card.v.addWidget(graph)
        self.ram = {"grid": g, "graph": graph}
        return card

    def _disk_card(self, spec):
        idx = spec["index"]
        title = f"Disk {idx}"
        card = Card(title)
        g = InfoGrid()
        g.add("Model", spec.get("model") or "—")
        g.add("Type", spec.get("type") or "—")
        g.add("Capacity")
        g.add("Read speed")
        g.add("Write speed")
        card.v.addWidget(g)
        card.v.addStretch(1)
        graph = LineGraph(100.0, lambda v: f"{v:.0f}%")
        card.v.addWidget(graph)
        self.disk_cards[idx] = {"grid": g, "graph": graph}
        return card

    # ---- wiring ---------------------------------------------------------
    def set_specs(self, specs):
        self.specs = specs
        self.loading.hide()
        gpus = specs.get("gpus") or []
        disks = specs.get("disks") or []
        self.grid.addWidget(self._cpu_card(specs.get("cpu") or {}), 0, 0)
        # systems without any GPU: hide the GPU section entirely
        right_row = 0
        if gpus:
            self.grid.addWidget(self._gpu_card(0, gpus[0]), 0, 1)
            right_row = 1
        self.grid.addWidget(self._ram_card(specs.get("ram") or {}), 1, 0)
        left_row = 2
        if disks:
            self.grid.addWidget(self._disk_card(disks[0]), max(1, right_row), 1)
            right_row = max(1, right_row) + 1
        for i, g in enumerate(gpus[1:], start=1):     # extra GPUs, ordered
            self.grid.addWidget(self._gpu_card(i, g), left_row, 0)
            left_row += 1
        for dspec in disks[1:]:                        # extra disks, ordered
            self.grid.addWidget(self._disk_card(dspec), right_row, 1)
            right_row += 1
        self.grid.setRowStretch(max(left_row, right_row), 1)

    def update_metrics(self, d):
        if not self.specs:
            return
        cpu = d.get("cpu") or {}
        if self.cpu:
            g = self.cpu["grid"]
            g.set("Speed", f"{cpu.get('ghz', 0):.2f} GHz")
            g.set("Utilization", f"{cpu.get('percent', 0):.0f}%")
            g.set("Processes", str(cpu.get("processes") or "—"))
            g.set("Threads", str(cpu["threads"]) if cpu.get("threads") else "—")
            g.set("Handles", str(cpu["handles"]) if cpu.get("handles") else "—")
            if d.get("cpu_temp") is not None:
                g.set_visible("Temperature", True)
                g.set("Temperature", f"{d['cpu_temp']:.0f} °C")
            self.cpu["graph"].add_point(cpu.get("percent", 0))
        mem = d.get("mem") or {}
        if self.ram and mem:
            g = self.ram["grid"]
            g.set("Capacity", fmt_used_total(mem["used"], mem["total"]))
            g.set("Page file", fmt_used_total(mem["page_used"], mem["page_total"]))
            g.set("Working set", fmt_bytes(mem.get("working_set")))
            self.ram["graph"].add_point(mem.get("percent", 0))
        gpu_samples = d.get("gpus") or []
        gpu_temps = d.get("gpu_temps") or []
        for i, gc in enumerate(self.gpu_cards):
            s = gpu_samples[i] if i < len(gpu_samples) else {}
            ded = s.get("dedicated_used")
            sha = s.get("shared_used")
            g = gc["grid"]
            g.set("Dedicated GPU memory", fmt_used_total(ded, gc["ded_total"]))
            g.set("Shared GPU memory", fmt_used_total(sha, gc["shared_total"]))
            total = gc["ded_total"] + gc["shared_total"]
            if total and ded is not None:
                pct = ((ded or 0) + (sha or 0)) / total * 100.0
                gc["graph"].add_point(pct)
            if i < len(gpu_temps):
                g.set_visible("Temperature", True)
                g.set("Temperature", f"{gpu_temps[i]} °C")
        for idx, dc in self.disk_cards.items():
            s = (d.get("disks") or {}).get(idx)
            if not s:
                continue
            g = dc["grid"]
            g.set("Capacity", fmt_used_total(s["used"], s["total"]))
            g.set("Read speed", f"{fmt_bytes(s['read_bps'])}/s")
            g.set("Write speed", f"{fmt_bytes(s['write_bps'])}/s")
            pct = (s["used"] / s["total"] * 100.0) if s.get("total") else 0.0
            dc["graph"].add_point(pct)

    def _export_data(self):
        """Static hardware specifications only (no live measurements)."""
        if not self.specs:
            return None
        s = self.specs
        cpu = s.get("cpu") or {}
        ram = s.get("ram") or {}
        return {
            "cpu": {
                "name": cpu.get("name"),
                "base_clock_mhz": cpu.get("base_mhz"),
                "cores": cpu.get("cores"),
                "logical_processors": cpu.get("logical"),
                "socket": cpu.get("socket"),
            },
            "gpus": [{
                "number": i,
                "name": g["name"],
                "dedicated_memory": fmt_bytes(g.get("dedicated_bytes")),
                "dedicated_memory_bytes": g.get("dedicated_bytes"),
                "driver_version": g.get("driver"),
            } for i, g in enumerate(s.get("gpus") or [])],
            "ram": {
                "type": ram.get("type"),
                "speed_mhz": ram.get("speed_mhz"),
                "form_factor": ram.get("form_factor"),
                "total": fmt_bytes(psutil.virtual_memory().total),
                "total_bytes": psutil.virtual_memory().total,
                "modules": ram.get("modules") or [],
            },
            "disks": [{
                "number": dd["index"],
                "model": dd["model"],
                "type": dd["type"],
                "size": fmt_bytes(dd["size_bytes"]),
                "size_bytes": dd["size_bytes"],
                "interface": dd.get("interface"),
                "serial_number": dd.get("serial"),
                "partitions": dd.get("partitions"),
            } for dd in (s.get("disks") or [])],
        }


# ================================================================ Network tab

class NetworkTab(QWidget):
    def __init__(self, get_dir, parent=None):
        super().__init__(parent)
        self.net_specs = {}
        self._last_nets = {}
        v = QVBoxLayout(self)
        v.setContentsMargins(18, 14, 18, 18)
        v.setSpacing(12)
        v.addWidget(ExportBar("Network", "network_specs", "network",
                              self._export_data, get_dir,
                              extra=("Open network adapter settings",
                                     lambda: launch_windows_tool(self, "control.exe",
                                                                 "ncpa.cpl"))))
        body = QHBoxLayout()
        body.setSpacing(12)
        self.list = QListWidget()
        self.list.setFixedWidth(220)
        self.list.currentRowChanged.connect(self._on_select)
        body.addWidget(self.list)
        self.stack = QStackedWidget()
        body.addWidget(self.stack, 1)
        v.addLayout(body, 1)
        self.pages = {}   # adapter name -> {"widget", "grid", "up_graph", "down_graph"}
        self._placeholder = QLabel("No active network adapters.")
        self._placeholder.setObjectName("dim")
        self._placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.stack.addWidget(self._placeholder)

    def set_net_specs(self, net_specs):
        self.net_specs = net_specs or {}
        for name, page in self.pages.items():
            page["grid"].set("Connection type", self._conn_type(name))

    def _conn_type(self, name):
        low = name.lower()
        spec = self.net_specs.get(name) or {}
        desc = (spec.get("description") or "").lower()
        if "wi-fi" in low or "wireless" in low or "wi-fi" in desc or "wireless" in desc or "802.11" in desc:
            return "Wi-Fi (802.11)"
        if "bluetooth" in low or "bluetooth" in desc:
            return "Bluetooth PAN"
        if "vethernet" in low or "vpn" in low or "virtual" in desc or "tap" in desc:
            return "Virtual adapter"
        return spec.get("adapter_type") or "Ethernet (802.3)"

    @staticmethod
    def _conn_table(headers):
        t = QTableWidget(0, len(headers))
        t.setHorizontalHeaderLabels(headers)
        t.verticalHeader().setVisible(False)
        t.verticalHeader().setDefaultSectionSize(24)
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        t.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        t.setAlternatingRowColors(True)
        t.setShowGrid(False)
        t.setWordWrap(False)
        t.setFixedHeight(190)
        hdr = t.horizontalHeader()
        for i in range(len(headers)):
            hdr.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        return t

    @staticmethod
    def _fill_table(t, rows):
        t.setRowCount(len(rows))
        for r, cols in enumerate(rows):
            for c, text in enumerate(cols):
                it = QTableWidgetItem(text)
                it.setFlags(Qt.ItemFlag.ItemIsEnabled)
                t.setItem(r, c, it)

    def _build_page(self, name):
        wrap = QScrollArea()
        wrap.setWidgetResizable(True)
        container = QWidget()
        col = QVBoxLayout(container)
        col.setContentsMargins(0, 0, 0, 0)
        col.setSpacing(12)

        card = Card(name)
        g = InfoGrid()
        g.add("Adapter name", name)
        g.add("Connection type", self._conn_type(name))
        g.add("IPv4 address")
        g.add("IPv6 address")
        g.add("Send throughput")
        card.v.addWidget(g)
        up_graph = LineGraph(None, fmt_bps)
        card.v.addWidget(up_graph)
        recv_grid = InfoGrid()
        recv_grid.add("Receive throughput")
        card.v.addWidget(recv_grid)
        down_graph = LineGraph(None, fmt_bps)
        card.v.addWidget(down_graph)
        col.addWidget(card)

        est_card = Card("Active connections")
        est_table = self._conn_table(["Local address", "Remote address", "Process"])
        est_card.v.addWidget(est_table)
        col.addWidget(est_card)

        lst_card = Card("Listening ports")
        lst_table = self._conn_table(["Port", "Local address", "Process"])
        lst_card.v.addWidget(lst_table)
        col.addWidget(lst_card)
        col.addStretch(1)

        wrap.setWidget(container)
        self.stack.addWidget(wrap)
        self.pages[name] = {"widget": wrap, "grid": g, "recv_grid": recv_grid,
                            "up_graph": up_graph, "down_graph": down_graph,
                            "est_table": est_table, "lst_table": lst_table}

    def _on_select(self, row):
        if row < 0:
            self.stack.setCurrentWidget(self._placeholder)
            return
        name = self.list.item(row).text()
        page = self.pages.get(name)
        if page:
            self.stack.setCurrentWidget(page["widget"])

    def update_metrics(self, d):
        nets = d.get("nets") or {}
        self._last_nets = nets
        current = [self.list.item(i).text() for i in range(self.list.count())]
        wanted = sorted(nets.keys())
        if current != wanted:
            selected = self.list.currentItem().text() if self.list.currentItem() else None
            # remove tabs for adapters that went inactive
            for name in list(self.pages):
                if name not in nets:
                    page = self.pages.pop(name)
                    self.stack.removeWidget(page["widget"])
                    page["widget"].deleteLater()
            self.list.blockSignals(True)
            self.list.clear()
            for name in wanted:
                if name not in self.pages:
                    self._build_page(name)   # adapter became active
                self.list.addItem(QListWidgetItem(name))
            self.list.blockSignals(False)
            row = wanted.index(selected) if selected in wanted else (0 if wanted else -1)
            self.list.setCurrentRow(row)
            if row < 0:
                self.stack.setCurrentWidget(self._placeholder)
        conns = d.get("conns") or {}
        for name, s in nets.items():
            page = self.pages.get(name)
            if not page:
                continue
            g = page["grid"]
            g.set("IPv4 address", s.get("ip4") or "—")
            g.set("IPv6 address", s.get("ip6") or "—")
            g.set("Send throughput", fmt_bps(s["up_bps"]))
            page["recv_grid"].set("Receive throughput", fmt_bps(s["down_bps"]))
            page["up_graph"].add_point(s["up_bps"])
            page["down_graph"].add_point(s["down_bps"])
            if self.stack.currentWidget() is page["widget"]:
                self._update_conn_tables(page, s, conns)

    def _update_conn_tables(self, page, s, conns):
        """Connections bound to this adapter's addresses (wildcard listeners
        included, since they accept on every adapter)."""
        ips = {x for x in (s.get("ip4"), s.get("ip6")) if x}
        est_rows = [(f"{c['ip']}:{c['port']}",
                     f"{c['rip']}:{c['rport']}",
                     f"{c['pname']} ({c['pid']})" if c.get("pid") else c["pname"])
                    for c in (conns.get("est") or []) if c["ip"] in ips]
        self._fill_table(page["est_table"], est_rows[:150])
        seen = set()
        lst_rows = []
        for c in (conns.get("lst") or []):
            if c["ip"] in ips or c["ip"] in ("0.0.0.0", "::"):
                key = (c["port"], c["ip"], c.get("pid"))
                if key in seen:
                    continue
                seen.add(key)
                lst_rows.append((str(c["port"]), c["ip"],
                                 f"{c['pname']} ({c['pid']})" if c.get("pid") else c["pname"]))
        self._fill_table(page["lst_table"], lst_rows[:150])

    def _export_data(self):
        """Per-adapter specifications only (no throughput measurements)."""
        if not self._last_nets:
            return None
        adapters = []
        for name in sorted(self._last_nets):
            s = self._last_nets[name]
            spec = self.net_specs.get(name) or {}
            adapters.append({
                "name": name,
                "description": spec.get("description"),
                "connection_type": self._conn_type(name),
                "mac_address": spec.get("mac"),
                "manufacturer": spec.get("manufacturer"),
                "link_speed_mbps": s.get("speed_mbps"),
                "ipv4_address": s.get("ip4"),
                "ipv6_address": s.get("ip6"),
            })
        return {"adapters": adapters}


# ================================================================ Logs tab

class SortItem(QTableWidgetItem):
    """Table item that sorts by a typed key stored in UserRole."""

    def __lt__(self, other):
        a = self.data(Qt.ItemDataRole.UserRole)
        b = other.data(Qt.ItemDataRole.UserRole)
        if a is not None and b is not None:
            try:
                return a < b
            except TypeError:
                return str(a) < str(b)
        return super().__lt__(other)


class CheckHeader(QHeaderView):
    """Horizontal header with a select-all checkbox over the first column."""
    toggled = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(Qt.Orientation.Horizontal, parent)
        self._checked = False

    def set_checked_silent(self, val):
        val = bool(val)
        if self._checked != val:
            self._checked = val
            self.viewport().update()

    def paintSection(self, painter, rect, logicalIndex):
        painter.save()
        super().paintSection(painter, rect, logicalIndex)
        painter.restore()
        if logicalIndex != 0:
            return
        size = 14.0
        x = rect.x() + (rect.width() - size) / 2
        y = rect.y() + (rect.height() - size) / 2
        box = QRectF(x, y, size, size)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        if self._checked:
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(PAL["accent"]))
            painter.drawRoundedRect(box, 4, 4)
            painter.setPen(QPen(QColor("#ffffff"), 1.6))
            painter.drawLine(QPointF(x + 3.2, y + 7.4), QPointF(x + 6.0, y + 10.2))
            painter.drawLine(QPointF(x + 6.0, y + 10.2), QPointF(x + 11.0, y + 4.2))
        else:
            painter.setPen(QPen(QColor(PAL["border"]), 1))
            painter.setBrush(QBrush(QColor(PAL["input"])))
            painter.drawRoundedRect(box, 4, 4)

    def mousePressEvent(self, e):
        # first section toggles select-all instead of sorting
        if self.logicalIndexAt(e.position().toPoint()) == 0:
            self._checked = not self._checked
            self.toggled.emit(self._checked)
            self.viewport().update()
            return
        super().mousePressEvent(e)


class LogsTab(QWidget):
    HEADERS = ["", "Level", "Date and Time", "Source", "Event ID",
               "Task Category", "Log"]

    def __init__(self, get_dir, parent=None):
        super().__init__(parent)
        self.get_dir = get_dir
        self._anchor_row = None
        v = QVBoxLayout(self)
        v.setContentsMargins(18, 14, 18, 18)
        v.setSpacing(12)

        top = QHBoxLayout()
        self.status = QLabel("Loading logs from Windows Event Viewer…")
        self.status.setObjectName("dim")
        top.addWidget(self.status)
        top.addStretch(1)
        self.export_btn = QPushButton("Export selected logs")
        self.export_btn.clicked.connect(self._export)
        top.addWidget(self.export_btn)
        evt_btn = QPushButton("Open Windows Event Manager")
        evt_btn.clicked.connect(lambda: launch_windows_tool(self, "eventvwr.exe"))
        top.addWidget(evt_btn)
        v.addLayout(top)

        t = QTableWidget(0, len(self.HEADERS))
        self.table = t
        t.setHorizontalHeaderLabels(self.HEADERS)
        self.check_header = CheckHeader(t)
        t.setHorizontalHeader(self.check_header)
        self.check_header.toggled.connect(self._select_all)
        t.verticalHeader().setVisible(False)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setAlternatingRowColors(True)
        t.setShowGrid(False)
        t.setWordWrap(False)
        hdr = t.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        t.setColumnWidth(0, 34)
        t.setColumnWidth(1, 105)
        t.setColumnWidth(2, 160)
        t.setColumnWidth(4, 80)
        t.setColumnWidth(5, 130)
        t.setColumnWidth(6, 100)
        t.itemClicked.connect(self._on_item_clicked)
        t.itemChanged.connect(self._on_item_changed)
        v.addWidget(t, 1)

    def set_logs(self, events):
        t = self.table
        t.setSortingEnabled(False)
        t.blockSignals(True)
        t.setRowCount(len(events))
        level_colors = {1: PAL["crit"], 2: PAL["err"], 3: PAL["warn"]}
        for r, e in enumerate(events):
            chk = SortItem("")
            chk.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
                         | Qt.ItemFlag.ItemIsUserCheckable)
            chk.setCheckState(Qt.CheckState.Unchecked)
            chk.setData(Qt.ItemDataRole.UserRole, 0)
            chk.setData(Qt.ItemDataRole.UserRole + 1, e)   # full event record
            t.setItem(r, 0, chk)

            lvl = SortItem(("● " if e["level_num"] in level_colors else "") + e["level"])
            lvl.setData(Qt.ItemDataRole.UserRole, e["level_num"])
            if e["level_num"] in level_colors:
                lvl.setForeground(QColor(level_colors[e["level_num"]]))
            t.setItem(r, 1, lvl)

            dt = SortItem(e["time_str"])
            dt.setData(Qt.ItemDataRole.UserRole, e["dt"].timestamp())
            t.setItem(r, 2, dt)

            src = SortItem(e["source"])
            src.setData(Qt.ItemDataRole.UserRole, e["source"].lower())
            t.setItem(r, 3, src)

            eid = SortItem(str(e["event_id"]))
            eid.setData(Qt.ItemDataRole.UserRole, e["event_id"])
            t.setItem(r, 4, eid)

            task = SortItem(e["task"])
            task.setData(Qt.ItemDataRole.UserRole, e["task"])
            t.setItem(r, 5, task)

            ch = SortItem(e["channel"])
            ch.setData(Qt.ItemDataRole.UserRole, e["channel"])
            t.setItem(r, 6, ch)
        t.blockSignals(False)
        t.setSortingEnabled(True)
        t.sortItems(2, Qt.SortOrder.DescendingOrder)  # newest first, like Event Viewer
        self._update_status()
        if not HAS_EVT:
            self.status.setText("pywin32 is required to read the Windows Event Log "
                                "(pip install pywin32).")

    # ---- selection ------------------------------------------------------
    def _on_item_clicked(self, item):
        if item.column() != 0:
            return
        row = item.row()
        mods = QApplication.keyboardModifiers()
        if mods & Qt.KeyboardModifier.ShiftModifier and self._anchor_row is not None:
            state = item.checkState()   # click already toggled this one
            lo, hi = sorted((self._anchor_row, row))
            self.table.blockSignals(True)
            for r in range(lo, hi + 1):
                it = self.table.item(r, 0)
                it.setCheckState(state)
                it.setData(Qt.ItemDataRole.UserRole,
                           1 if state == Qt.CheckState.Checked else 0)
            self.table.blockSignals(False)
            self._update_status()
        else:
            self._anchor_row = row

    def _on_item_changed(self, item):
        if item.column() != 0:
            return
        item.setData(Qt.ItemDataRole.UserRole,
                     1 if item.checkState() == Qt.CheckState.Checked else 0)
        self._update_status()

    def _select_all(self, checked):
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        t = self.table
        t.blockSignals(True)
        for r in range(t.rowCount()):
            it = t.item(r, 0)
            if it:
                it.setCheckState(state)
                it.setData(Qt.ItemDataRole.UserRole, 1 if checked else 0)
        t.blockSignals(False)
        self._anchor_row = None
        self._update_status()

    def _checked_events(self):
        out = []
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it and it.checkState() == Qt.CheckState.Checked:
                out.append(it.data(Qt.ItemDataRole.UserRole + 1))
        return out

    def _update_status(self):
        n = self.table.rowCount()
        sel = len(self._checked_events())
        self.status.setText(f"{n} logs loaded (Application, Setup, System) — {sel} selected")
        self.check_header.set_checked_silent(n > 0 and sel == n)

    # ---- export ---------------------------------------------------------
    def _export(self):
        events = self._checked_events()
        if not events:
            QMessageBox.warning(self, APP_NAME,
                                "Select at least one log entry to export (use the "
                                "checkboxes; Shift+click selects a range).")
            return
        newest = max(e["dt"] for e in events)   # from the "Date and Time" column
        filename = "logs_" + newest.strftime("%Y-%m-%d_%H-%M-%S") + ".txt"
        headers = ["Level", "Date and Time", "Source", "Event ID",
                   "Task Category", "Log"]
        lines = ["\t".join(headers)]
        for e in events:
            lines.append("\t".join([e["level"], e["time_str"], e["source"],
                                    str(e["event_id"]), e["task"], e["channel"]]))
        try:
            directory = self.get_dir()
            os.makedirs(directory, exist_ok=True)
            with open(os.path.join(directory, filename), "w", encoding="utf-8") as f:
                f.write("\n".join(lines) + "\n")
        except Exception as ex:
            QMessageBox.critical(self, APP_NAME, f"Export failed:\n{ex}")
            return
        QMessageBox.information(
            self, "Export successful",
            f"{len(events)} log entr{'y was' if len(events) == 1 else 'ies were'} "
            f"exported successfully as “{filename}”.")


# ================================================================ Settings tab

ABOUT_TEXT = (f"{APP_NAME} {APP_VERSION} is a lightweight system administration "
              "dashboard for Windows 10 and 11. It brings live hardware telemetry, "
              "network adapter monitoring, and the Windows Event Log together in a "
              "single, minimal window — with one-click exports to JSON, XML, and "
              "YAML for reporting and support workflows. Built with Python, PyQt6, "
              "and psutil.")


class SettingsTab(QWidget):
    """All user preferences; live-applies and persists to settings.json."""
    changed = pyqtSignal()          # theme / on-top / refresh changed

    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.s = settings
        v = QVBoxLayout(self)
        v.setContentsMargins(120, 24, 120, 24)
        v.setSpacing(14)

        # ---- appearance
        ap = Card("Appearance")
        row = QHBoxLayout()
        lab = QLabel("Theme")
        lab.setObjectName("dim")
        row.addWidget(lab)
        row.addStretch(1)
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Light", "Dark"])
        self.theme_combo.setCurrentIndex(1 if self.s["theme"] == "dark" else 0)
        self.theme_combo.setFixedWidth(140)
        self.theme_combo.currentIndexChanged.connect(self._theme_changed)
        row.addWidget(self.theme_combo)
        ap.v.addLayout(row)
        self.follow_chk = QCheckBox("Follow system theme for light/dark mode")
        self.follow_chk.setChecked(self.s["follow_system"])
        self.follow_chk.toggled.connect(self._follow_changed)
        ap.v.addWidget(self.follow_chk)
        self.theme_combo.setDisabled(self.s["follow_system"])
        v.addWidget(ap)

        # ---- data refresh
        dc = Card("Data")
        row = QHBoxLayout()
        lab = QLabel("Data refresh rate (milliseconds) — applies to Hardware and "
                     "Network graphs and to the network connection lists")
        lab.setObjectName("dim")
        lab.setWordWrap(True)
        row.addWidget(lab, 1)
        row.addStretch(1)
        self.refresh_edit = QLineEdit(str(self.s["refresh_ms"]))
        self.refresh_edit.setValidator(QIntValidator(1, 3600000, self))
        self.refresh_edit.setFixedWidth(90)
        self.refresh_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.refresh_edit.editingFinished.connect(self._refresh_changed)
        row.addWidget(self.refresh_edit)
        dc.v.addLayout(row)
        v.addWidget(dc)

        # ---- export location
        ec = Card("Exports")
        row = QHBoxLayout()
        lab = QLabel("Export folder")
        lab.setObjectName("dim")
        row.addWidget(lab)
        self.dir_edit = QLineEdit(self.s["export_dir"])
        self.dir_edit.editingFinished.connect(self._dir_changed)
        row.addWidget(self.dir_edit, 1)
        browse = QPushButton("…")
        browse.setFixedWidth(44)
        browse.clicked.connect(self._browse)
        row.addWidget(browse)
        ec.v.addLayout(row)
        v.addWidget(ec)

        # ---- window
        wc = Card("Window")
        self.ontop_chk = QCheckBox("Keep the application window always on top")
        self.ontop_chk.setChecked(self.s["always_on_top"])
        self.ontop_chk.toggled.connect(self._ontop_changed)
        wc.v.addWidget(self.ontop_chk)
        v.addWidget(wc)

        # ---- save button (above the About box)
        save_row = QHBoxLayout()
        save_row.addStretch(1)
        save_btn = QPushButton("Save settings")
        save_btn.clicked.connect(self._save_clicked)
        save_row.addWidget(save_btn)
        v.addLayout(save_row)

        # ---- about
        about_lab = QLabel("About")
        about_lab.setObjectName("cardTitle")
        about_lab.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        v.addWidget(about_lab)
        ab = Card()
        text = QLabel(ABOUT_TEXT)
        text.setWordWrap(True)
        text.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        ab.v.addWidget(text)
        self.github = GitHubButton()
        ab.v.addWidget(self.github, 0, Qt.AlignmentFlag.AlignHCenter)
        v.addWidget(ab)
        v.addStretch(1)

    # ---- handlers (live-apply; persisted by Save or on app close) -------
    def _theme_changed(self, idx):
        self.s["theme"] = "dark" if idx == 1 else "light"
        self.changed.emit()

    def _follow_changed(self, on):
        self.s["follow_system"] = bool(on)
        self.theme_combo.setDisabled(on)
        self.changed.emit()

    def _refresh_changed(self):
        try:  # floor of 100 ms keeps the sampler from pegging a core
            self.s["refresh_ms"] = max(100, int(self.refresh_edit.text()))
        except ValueError:
            pass
        self.refresh_edit.setText(str(self.s["refresh_ms"]))
        self.changed.emit()

    def _dir_changed(self):
        path = self.dir_edit.text().strip()
        if path:
            self.s["export_dir"] = path

    def _browse(self):
        path = QFileDialog.getExistingDirectory(self, "Choose export folder",
                                                self.s["export_dir"])
        if path:
            self.s["export_dir"] = os.path.normpath(path)
            self.dir_edit.setText(self.s["export_dir"])

    def _ontop_changed(self, on):
        self.s["always_on_top"] = bool(on)
        self.changed.emit()

    def sync(self):
        """Pull any uncommitted textbox edits into the settings dict."""
        self._dir_changed()
        try:
            self.s["refresh_ms"] = max(100, int(self.refresh_edit.text()))
        except ValueError:
            pass

    def _save_clicked(self):
        self.sync()
        if save_settings(self.s):
            QMessageBox.information(
                self, "Settings saved",
                "User settings were saved successfully inside the configuration "
                "file (settings.json).")
        else:
            QMessageBox.critical(self, APP_NAME, "Could not write settings.json.")


# ================================================================ main window

def system_prefers_dark():
    try:
        scheme = QApplication.styleHints().colorScheme()
        if scheme == Qt.ColorScheme.Dark:
            return True
        if scheme == Qt.ColorScheme.Light:
            return False
    except Exception:
        pass
    try:
        import winreg
        k = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        val = winreg.QueryValueEx(k, "AppsUseLightTheme")[0]
        winreg.CloseKey(k)
        return val == 0
    except Exception:
        return False


class MainWindow(QMainWindow):
    def __init__(self, settings):
        super().__init__()
        self.s = settings
        self.setWindowTitle(APP_NAME)
        self.setFixedSize(WINDOW_W, WINDOW_H)      # fixed 1280x960
        flags = (Qt.WindowType.Window
                 | Qt.WindowType.WindowTitleHint
                 | Qt.WindowType.WindowSystemMenuHint
                 | Qt.WindowType.WindowMinimizeButtonHint
                 | Qt.WindowType.WindowCloseButtonHint)   # no maximize
        if self.s["always_on_top"]:
            flags |= Qt.WindowType.WindowStaysOnTopHint
        self.setWindowFlags(flags)

        get_dir = lambda: self.s["export_dir"]
        self.tabs = QTabWidget()
        self.system_tab = SystemTab()
        self.hardware_tab = HardwareTab(get_dir)
        self.network_tab = NetworkTab(get_dir)
        self.logs_tab = LogsTab(get_dir)
        self.settings_tab = SettingsTab(self.s)
        for w, name in ((self.system_tab, "System"),
                        (self.hardware_tab, "Hardware"),
                        (self.network_tab, "Network"),
                        (self.logs_tab, "Logs"),
                        (self.settings_tab, "Settings")):
            self.tabs.addTab(w, name)
        self.setCentralWidget(self.tabs)

        self.settings_tab.changed.connect(self._apply_live_settings)

        # ---- background workers
        self.metrics = MetricsThread(self)
        self.metrics.interval = self.s["refresh_ms"]
        self.metrics.sample.connect(self._on_sample)
        self.specs_thread = SpecsThread(self)
        self.specs_thread.ready.connect(self._on_specs)
        self.logs_thread = LogsThread(self)
        self.logs_thread.loaded.connect(self.logs_tab.set_logs)
        self.specs_thread.start()
        self.logs_thread.start()
        self.metrics.start()

        try:  # live-follow OS light/dark switches (Qt >= 6.5)
            QApplication.styleHints().colorSchemeChanged.connect(self._on_scheme_changed)
        except Exception:
            pass
        self.apply_theme()

    # ---- data plumbing --------------------------------------------------
    def _on_specs(self, specs):
        self.system_tab.set_specs(specs)
        self.hardware_tab.set_specs(specs)
        self.network_tab.set_net_specs(specs.get("net") or {})
        self.metrics.set_disk_map(specs.get("disk_map") or {})

    def _on_sample(self, d):
        self.system_tab.update_metrics(d)
        self.hardware_tab.update_metrics(d)
        self.network_tab.update_metrics(d)

    # ---- settings / theme ------------------------------------------------
    def _on_scheme_changed(self, *_):
        if self.s["follow_system"]:
            self.apply_theme()

    def _apply_live_settings(self):
        self.metrics.interval = self.s["refresh_ms"]
        want_top = self.s["always_on_top"]
        have_top = bool(self.windowFlags() & Qt.WindowType.WindowStaysOnTopHint)
        if want_top != have_top:
            self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, want_top)
            self.show()
        self.apply_theme()

    def apply_theme(self):
        if self.s["follow_system"]:
            dark = system_prefers_dark()
        else:
            dark = self.s["theme"] == "dark"
        PAL.clear()
        PAL.update(DARK if dark else LIGHT)
        QApplication.instance().setStyleSheet(build_qss(PAL))
        self.system_tab.retheme()
        self.settings_tab.github.refresh_icon()
        self.setWindowIcon(self._make_icon())

    def _make_icon(self):
        pix = QPixmap(64, 64)
        pix.fill(Qt.GlobalColor.transparent)
        p = QPainter(pix)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(PAL["logo"]))
        for (x, y) in ((4, 4), (34, 4), (4, 34), (34, 34)):
            p.drawRoundedRect(QRectF(x, y, 26, 26), 4, 4)
        p.end()
        return QIcon(pix)

    # ---- shutdown ---------------------------------------------------------
    def closeEvent(self, ev):
        self.settings_tab.sync()
        save_settings(self.s)          # auto-save changed settings on close
        self.metrics.stop()
        self.metrics.wait(3000)
        for th in (self.specs_thread, self.logs_thread):
            if th.isRunning():
                th.wait(1000)
        super().closeEvent(ev)


def main():
    ensure_admin()
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    settings = load_settings()
    win = MainWindow(settings)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
