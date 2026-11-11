# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：就业质量报告生成器

用法（在仓库根目录执行）：
    pyinstaller "就业质量报告生成器.spec" --noconfirm

产物：dist/就业质量报告生成器.exe（单文件，脱离 Python 环境即可运行）
"""

import os

# 以 spec 文件所在目录为基准，克隆到任意路径都能直接打包
_HERE = globals().get("SPECPATH", os.getcwd())

a = Analysis(
    [os.path.join(_HERE, "report_generator.py")],
    pathex=[],
    binaries=[],
    datas=[],
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
    name="就业质量报告生成器",
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
