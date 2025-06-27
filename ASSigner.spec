# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules
import os

hidden_nacl = collect_submodules('nacl.bindings') + ['_cffi_backend']

# whatever Playwright needs
hidden_playwright = ['playwright.__main__']

# combine them
all_hidden = hidden_playwright + hidden_nacl

a = Analysis(
    ['ASSigner.py'],
    pathex=[os.path.abspath('.')],
    binaries=[],
    datas=[('client_config.py', '.')],
    hiddenimports=all_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ASSigner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
