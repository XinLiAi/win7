# -*- mode: python ; coding: utf-8 -*-
"""
设备密码管理系统 - PyInstaller 打包配置
用法（Windows）：
    python -m PyInstaller --noconfirm --clean device_manager.spec
产物：dist\DeviceManager.exe （单文件、无控制台窗口）
"""
import os

block_cipher = None

# 收集 openpyxl（可选依赖；已安装则一并打入，未安装也可正常打包）
hidden = []
try:
    import openpyxl  # noqa
    hidden.append("openpyxl")
except ImportError:
    pass

a = Analysis(
    ["device_manager.py"],
    pathex=[SPECPATH],
    binaries=[],
    datas=[],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["numpy", "pandas", "matplotlib", "scipy", "PIL.ImageQt", "PIL._tkinter_finder"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 图标文件存在才使用，避免因缺图标导致打包失败
_icon = os.path.join(SPECPATH, "device_manager.ico")
_icon = _icon if os.path.exists(_icon) else None

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="DeviceManager",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,               # GUI 程序，不弹黑色控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=_icon,                  # 程序图标（不存在则自动跳过）
)
