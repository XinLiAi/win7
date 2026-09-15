; ============================================================
;  设备密码管理系统 - Inno Setup 安装包脚本
;  作用：把 PyInstaller 产出的 DeviceManager.exe 打包成一键安装程序
;  产物：Output\设备密码管理系统_Setup.exe
;  使用：安装 Inno Setup 后，双击本文件或用 ISCC.exe 编译
;  下载 Inno Setup：https://jrsoftware.org/isdl.php
; ============================================================

#define MyAppName "设备密码管理系统"
#define MyAppVersion "3.1"
#define MyAppPublisher "DeviceManager"
#define MyAppExeName "DeviceManager.exe"

[Setup]
AppId={{8F3A2B1D-7C4E-4F6A-9B8D-2E5F1A3C7D90}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\DeviceManager
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=设备密码管理系统_Setup_v{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\{#MyAppExeName}
AppPublisherURL=
AppSupportURL=
AppUpdatesURL=
; 安装包图标（如果有）
; SetupIconFile=device_manager.ico

; 使用项目自带的中文语言包（ChineseSimplified.isl），不依赖 Inno Setup 安装目录
[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加图标:"; Flags: unchecked

[Files]
; 主程序（PyInstaller 单文件 EXE）
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; 如果有图标文件也一并带上
Source: "device_manager.ico"; DestDir: "{app}"; Flags: ignoreversion; Check: FileExists('device_manager.ico')

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即运行 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 卸载时可选删除用户数据（默认不删，保留数据）
; Type: filesandordirs; Name: "{userappdata}\.device_manager"
