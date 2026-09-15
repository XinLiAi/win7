@echo off
chcp 65001 >nul
rem ============================================================
rem  设备密码管理系统 - Windows 一键打包脚本（EXE + 安装包）
rem  双击本文件即可，产物：
rem    1. dist\DeviceManager.exe          （单文件绿色版，拷走就能用）
rem    2. Output\设备密码管理系统_Setup_v3.1.exe （一键安装包，带快捷方式/卸载）
rem  要求：Python 3.8+（勾选 Add to PATH）
rem        可选：Inno Setup 6（用于生成安装包，未装则只生成 EXE）
rem ============================================================
setlocal
cd /d "%~dp0"

echo.
echo  ============================================
echo   设备密码管理系统  Windows 一键打包
echo  ============================================
echo.

rem ---- 1. 检查 Python ----
where python >nul 2>nul
if errorlevel 1 (
    echo  [错误] 未找到 Python，请先安装 Python 3.8+ 并勾选 Add to PATH
    echo         下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo  [1/3] 安装打包依赖（固定版本，确保 Win7 兼容）...
python -m pip install --upgrade "pip<24"
python -m pip install pyinstaller==5.13.2 openpyxl==3.1.2
if errorlevel 1 (
    echo  [错误] 依赖安装失败，请检查网络后重试
    pause
    exit /b 1
)

echo.
echo  [2/3] 打包 EXE（约 1~3 分钟）...
python -m PyInstaller --noconfirm --clean device_manager.spec
if errorlevel 1 (
    echo  [错误] EXE 打包失败，请查看上方报错信息
    pause
    exit /b 1
)

echo.
echo  [3/3] 生成一键安装包...
rem 尝试查找 Inno Setup 编译器 ISCC.exe
set "ISCC="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
where ISCC >nul 2>nul && for /f "delims=" %%i in ('where ISCC') do set "ISCC=%%i"

if defined ISCC (
    echo  找到 Inno Setup: %ISCC%
    "%ISCC%" device_manager_installer.iss
    if errorlevel 1 (
        echo  [警告] 安装包编译失败，但 EXE 已生成
    ) else (
        echo.
        echo  ============================================
        echo   打包全部完成！
        echo   绿色版:   dist\DeviceManager.exe
        echo   安装包:   Output\设备密码管理系统_Setup_v3.1.exe
        echo   数据目录: %%USERPROFILE%%\.device_manager\
        echo  ============================================
    )
) else (
    echo.
    echo  [提示] 未检测到 Inno Setup，已生成绿色版 EXE。
    echo  如需一键安装包，请安装 Inno Setup 6 后重新运行本脚本。
    echo  下载地址: https://jrsoftware.org/isdl.php
    echo.
    echo  ============================================
    echo   绿色版 EXE: dist\DeviceManager.exe
    echo   双击即可运行，无需安装 Python
    echo  ============================================
)

echo.
pause
endlocal
