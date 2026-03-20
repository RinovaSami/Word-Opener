# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Word Opener
Produces a single self-contained WordOpener.exe
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ── data files ────────────────────────────────────────────────────────────────
datas = []
datas += collect_data_files("mammoth")
datas += collect_data_files("flask")
datas += collect_data_files("jinja2")
datas += collect_data_files("markupsafe")

# ── hidden imports ────────────────────────────────────────────────────────────
hiddenimports = []
hiddenimports += collect_submodules("mammoth")
hiddenimports += [
    "flask",
    "flask.templating",
    "jinja2",
    "jinja2.ext",
    "werkzeug",
    "werkzeug.serving",
    "werkzeug.routing",
    "werkzeug.wrappers",
    "werkzeug.utils",
    "werkzeug.datastructures",
    "markupsafe",
    "click",
]

# ── analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ["word_opener.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "unittest", "test", "distutils"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="WordOpener",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,          # keep console so user can see URL + Ctrl-C to stop
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
