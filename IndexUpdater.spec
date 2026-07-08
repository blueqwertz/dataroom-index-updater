# -*- mode: python ; coding: utf-8 -*-
import sys

from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

is_mac = sys.platform == 'darwin'

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='IndexUpdater',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=not is_mac,          # UPX is unreliable/unavailable on macOS
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,        # build for the runner's native arch
    codesign_identity=None,
    entitlements_file=None,
)

if is_mac:
    app = BUNDLE(
        exe,
        name='IndexUpdater.app',
        icon=None,
        bundle_identifier='com.blueqwertz.dataroom-index-updater',
        info_plist={
            'CFBundleName': 'Dataroom Index Updater',
            'CFBundleDisplayName': 'Dataroom Index Updater',
            'NSHighResolutionCapable': True,
            'LSMinimumSystemVersion': '11.0',
        },
    )
