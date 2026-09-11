# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置
# 用法：pyinstaller 发货数据填入工具.spec
# 或带版本号：pyinstaller --name "发货数据填入工具_v1.10" 发货数据填入工具.spec

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('app.ico', '.')],   # 应用图标随包分发（模板不打包，由用户自行选择）
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
    name='发货数据填入工具',
    icon='app.ico',            # exe 自身的图标（资源管理器 / 任务栏）
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
