# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for RPA Tool
# Build with: pyinstaller build.spec
# Output: dist/rpa_tool.exe
#
# Run once to install PyInstaller:
#   pip install pyinstaller

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('config.toml', '.'),          # bundle default config
    ],
    hiddenimports=[
        # pywinauto backend modules not auto-detected
        'pywinauto.backends',
        'pywinauto.backends.uia',
        'pywinauto.backends.win32',
        'pywinauto.base_wrapper',
        'pywinauto.controls.uia_controls',
        'pywinauto.controls.win32_controls',
        # pynput platform backend
        'pynput.mouse._win32',
        'pynput.keyboard._win32',
        # pystray platform backend
        'pystray._win32',
        # PIL imaging plugins
        'PIL._imaging',
        'PIL.PngImagePlugin',
        'PIL.JpegImagePlugin',
        # OpenCV
        'cv2',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'pandas', 'IPython',
        'jupyter', 'notebook', 'setuptools',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='rpa_tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no console window in production
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # add a .ico path here for a custom tray icon
    uac_admin=True,         # request elevation on launch
)
