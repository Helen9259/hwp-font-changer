# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files [cite: 1]

# pywin32 관련 하위 모듈과 데이터 파일들을 유실 없이 수집합니다. 
hidden_imports = ['win32com', 'win32com.client', 'win32api', 'win32con', 'pywintypes'] + collect_submodules('win32com') [cite: 1]
datas_collected = collect_data_files('win32com') [cite: 1]

a = Analysis(
    ['main.py'], [cite: 1]
    pathex=[], [cite: 1]
    binaries=[], [cite: 1]
    datas=datas_collected,          # 강제 수집 데이터 주입 
    hiddenimports=hidden_imports,  # 강제 수집 히든 임포트 주입 
    hookspath=[], [cite: 1]
    hooksconfig={}, [cite: 1]
    runtime_hooks=[], [cite: 1]
    excludes=[], [cite: 1]
    noarchive=False, [cite: 1]
)

pyz = PYZ(a.pure) [cite: 1]

exe = EXE(
    pyz, [cite: 1]
    a.scripts, [cite: 1]
    a.binaries, [cite: 1]
    a.datas, [cite: 1]
    [], [cite: 1]
    name='HWP_Font_Changer', [cite: 1]
    debug=False, [cite: 1]
    bootloader_ignore_signals=False, [cite: 1]
    strip=False, [cite: 1]
    upx=True, [cite: 1]
    upx_exclude=[], [cite: 2]
    runtime_tmpdir=None, [cite: 1]
    console=False,                  # GUI 전용 프로그램이므로 False 유지 
    disable_windowed_traceback=False, [cite: 1]
    argv_emulation=False, [cite: 1]
    target_arch=None, [cite: 1]
    codesign_identity=None, [cite: 1]
    entitlements_file=None, [cite: 1]
    icon=None, [cite: 1]
)
