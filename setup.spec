# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['converter.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='File Converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='File Converter',
)

app = BUNDLE(
    coll,
    name='File Converter.app',
    icon='aqvmq-rle18.icns',
    bundle_identifier='com.fileconverter.app',
    info_plist={
        'NSHighResolutionCapable': 'True',
        'LSApplicationCategoryType': 'public.app-category.utilities',
        'CFBundleShortVersionString': '1.0.0',
    },
)