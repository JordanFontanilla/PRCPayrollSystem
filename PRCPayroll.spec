# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['PRCPayrollSystem\\Main\\Main.py'],
    pathex=['.'],
    binaries=[],
    datas=[('PRCPayrollSystem/Components', 'Components'), ('PRCPayrollSystem/settingsAndFields', 'settingsAndFields'), ('PRCPayrollSystem/pastLoadedHistory', 'pastLoadedHistory'), ('PRCPayrollSystem/pastPayslips', 'pastPayslips')],
    hiddenimports=[],
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
    name='PRCPayroll',
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
    icon=['PRCPayrollSystem\\Components\\PRClogo.png'],
)
