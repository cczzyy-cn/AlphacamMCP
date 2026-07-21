"""
Configuration helpers — detect AlphaCAM install dir, read .env, manage prog ID.
"""

from __future__ import annotations

import os
import logging
from pathlib import Path

log = logging.getLogger("alphacam-bridge.config")

# ---------------------------------------------------------------------------
# .env support
# ---------------------------------------------------------------------------

def _load_dotenv():
    """Load .env file from project root (if present)."""
    # Look for .env next to this file, or in CWD
    candidates = [
        Path(__file__).resolve().parent.parent / ".env",
        Path.cwd() / ".env",
    ]
    for env_path in candidates:
        if env_path.is_file():
            log.info(f"Loading .env from {env_path}")
            for line in env_path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip().strip("\"'")
                if key not in os.environ:
                    os.environ[key] = val


_load_dotenv()

# ---------------------------------------------------------------------------
# Public configuration getters
# ---------------------------------------------------------------------------

def get_prog_id() -> str:
    """Return the COM ProgID (env override or default)."""
    return os.environ.get("ALPHACAM_PROG_ID", "aroutaps.Application")


def get_visible() -> bool:
    """Return whether AlphaCAM window should be visible."""
    val = os.environ.get("ALPHACAM_VISIBLE", "1")
    return val.lower() in ("1", "true", "yes")


def get_acPort() -> int | None:
    """Return optional SSE port (env override)."""
    val = os.environ.get("ALPHACAM_PORT", "")
    return int(val) if val else None


def get_sse_token() -> str | None:
    """Return optional Bearer token for SSE mode."""
    val = os.environ.get("ALPHACAM_SSE_TOKEN", "")
    return val if val else None


# ---------------------------------------------------------------------------
# AlphaCAM installation directory detection
# ---------------------------------------------------------------------------

_COMMON_ACAM_DIRS = [
    r"C:\Program Files (x86)\Vero Software\Alphacam 2016 R1",
    r"C:\Program Files\Vero Software\Alphacam 2016 R1",
    r"C:\Program Files (x86)\Vero Software\Alphacam 2019",
    r"C:\Program Files\Vero Software\Alphacam 2019",
    r"C:\Program Files (x86)\Vero Software\Alphacam 2021",
    r"C:\Program Files\Vero Software\Alphacam 2021",
    r"C:\Program Files (x86)\Vero Software\Alphacam 2024",
    r"C:\Program Files\Vero Software\Alphacam 2024",
]


def detect_alphacam_dir() -> str | None:
    """Detect the AlphaCAM installation directory (best-effort)."""
    # Method 1: env override
    env_dir = os.environ.get("ALPHACAM_DIR", "")
    if env_dir and os.path.isdir(env_dir):
        return env_dir

    # Method 2: COM application path (if already connected)
    try:
        # Late import to avoid circular dependency
        from .handler import _get_acam_instance
        inst = _get_acam_instance()
        if inst is not None:
            info = inst.get_info()
            exe_path = info.get("path", "")
            if exe_path and os.path.isdir(exe_path):
                return exe_path
    except Exception:
        pass

    # Method 3: read registry
    try:
        import winreg
        for reg_key in [
            r"SOFTWARE\Vero Software\Alphacam 2016 R1",
            r"SOFTWARE\WOW6432Node\Vero Software\Alphacam 2016 R1",
            r"SOFTWARE\Licom\Alphacam\2016 R1",
        ]:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_key) as key:
                    val, _ = winreg.QueryValueEx(key, "InstallPath")
                    if val and os.path.isdir(val):
                        return val
            except Exception:
                pass
    except Exception:
        pass

    # Method 4: check common paths
    for d in _COMMON_ACAM_DIRS:
        if os.path.isdir(d):
            return d

    return None


# ---------------------------------------------------------------------------
# CHM documentation index
# ---------------------------------------------------------------------------

CHM_DOCS = {
    "ACAMAPI": {"file": "ACAMAPI.chm", "desc": "API reference (VBA object model, methods, properties, events)"},
    "ACAM3": {"file": "ACAM3.chm", "desc": "3D module user manual"},
    "ACAM4": {"file": "ACAM4.chm", "desc": "4-axis module user manual"},
    "AEDIT3": {"file": "AEDIT3.chm", "desc": "3D editor help"},
    "AEDITAPI": {"file": "AEDITAPI.chm", "desc": "Editor API reference"},
    "AcamReports": {"file": "AcamReports.chm", "desc": "Reports system help"},
    "C2C": {"file": "C2C.chm", "desc": "CAD to CAM conversion help"},
    "ConstraintsAPI": {"file": "ConstraintsAPI.chm", "desc": "Constraints API reference"},
    "Feature": {"file": "Feature.chm", "desc": "Feature help"},
    "ModuleWorks": {"file": "ModuleWorks_-_Documentation.chm", "desc": "ModuleWorks machining engine docs (5-axis)"},
    "Primitives": {"file": "Primitives.chm", "desc": "Primitives help"},
    "R2V": {"file": "R2V.chm", "desc": "Raster to Vector help"},
    "simulate": {"file": "simulate.chm", "desc": "Simulator help"},
    # --- 以下为本次扫描补充的文档条目 ---
    "CDM": {"file": "StartUp/CDM/CDM.chm", "desc": "CDM (Cabinet Design Module) user manual"},
    "APM": {"file": "StartUp/APM/APM.chm", "desc": "APM advanced machining module help"},
    "AcamMultidrill": {"file": "StartUp/Machine/AcamMultidrill/AcamMultidrill.chm", "desc": "Multi-drill machining help"},
    "AutoCutout": {"file": "StartUp/Machine/AutoCutout/AutoCutout.chm", "desc": "Auto cutout help"},
    "AutoGeometryBridge": {"file": "StartUp/Utils/AutoGeometryBridge/AUTOGEOMETRYBRIDGE.chm", "desc": "Auto geometry bridge help"},
    "BoringAlong2DLine": {"file": "StartUp/Machine/BoringAlong2DLine/BoringAlong2DLine.chm", "desc": "Boring along 2D line help"},
    "BXF": {"file": "StartUp/File/BXF/BXF.chm", "desc": "BXF file format help"},
    "ColorToCAM": {"file": "StartUp/Machine/ColorToCAM/ColorToCAM.chm", "desc": "Color to CAM machining help"},
    "CornerDrill": {"file": "StartUp/Machine/CornerDrill/CornerDrill.chm", "desc": "Corner drilling help"},
    "Define3DWorkPlane": {"file": "StartUp/3D/3DWorkPlane/Define3DWorkPlane.chm", "desc": "3D workplane definition help"},
    "DefineAggregate": {"file": "StartUp/Machine/SelectAggregate/DefineAggregate.chm", "desc": "Aggregate/spindle definition help"},
    "HolesForMultidrill": {"file": "StartUp/Geometry/HolesforMultidrill/HOLESFORMULTIDRILL.chm", "desc": "Holes for multi-drill help"},
    "ImportBMP": {"file": "StartUp/File/ImportBMP/ImportBMP.chm", "desc": "BMP import help"},
    "InvoluteGear": {"file": "StartUp/Geometry/InvoluteGear/InvoluteGear.chm", "desc": "Involute gear help"},
    "MultiDrill": {"file": "StartUp/Machine/MultiDrill/MULTIDRILL.chm", "desc": "Multi-drill help"},
    "NcOutputManager": {"file": "StartUp/File/NcOutputManager/NcOutputManager.chm", "desc": "NC output manager help"},
    "OpenPocket": {"file": "StartUp/Machine/MillingCycles/OPENPOCKET.chm", "desc": "Open pocket milling help"},
    "OpListJP": {"file": "Add-Ins/OpListJP/OpListJP.chm", "desc": "OpListJP add-in help (Japanese)"},
    "ReliefCut": {"file": "StartUp/Machine/ReliefCut/ReliefCut.chm", "desc": "Relief cutting help"},
    "WineRack": {"file": "StartUp/Geometry/WineRack/WineRack.chm", "desc": "Wine rack geometry help"},
}
