# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[('resources', 'resources')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Playwright's Chromium lives in a separate cache, not the wheel
        'playwright._driver.chromium',
        # Qt modules the app never imports (verified: no WebEngine/Quick/Multimedia refs).
        # Qt6WebEngineCore.dll alone is ~193 MB.
        'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineQuick',
        'PySide6.QtQuick', 'PySide6.QtQuick3D', 'PySide6.QtQml', 'PySide6.QtQuickWidgets',
        'PySide6.QtMultimedia', 'PySide6.QtMultimediaWidgets', 'PySide6.QtCharts',
        'PySide6.Qt3DCore', 'PySide6.QtDesigner',
        # Other Qt bindings present in the venv - prevent accidental double-bundling
        'PyQt5', 'PyQt6', 'tkinter',
        # ML / scientific stacks installed in the venv but NOT imported by the app
        # (verified: zero references in *.py). Confirmed in TOC: ~450 MB uncompressed.
        'torch', 'torchvision', 'torchaudio', 'scipy', 'sympy',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

splash = Splash(
    'resources\\images\\logo.jpg',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
    text_size=12,
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    splash,
    splash.binaries,
    [],
    name='start',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # UPX packing often trips Windows Defender and slows launch; excludes save far more
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['logo.ico'],
)
