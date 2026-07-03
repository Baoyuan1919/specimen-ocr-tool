; Inno Setup 安装脚本
; 用 Inno Setup 6 编译：ISCC.exe install_script.iss

#define MyAppName "标本OCR填表工具"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "标本助手"
#define MyAppExeName "标本OCR填表工具.exe"
#define MyAppAssocName "标本数据文件"

[Setup]
AppId={{F3A7C1B8-9D5E-4A1B-8C3F-6E2D7A0B9C1E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
OutputDir=.\installer
OutputBaseFilename=标本OCR填表工具_安装包
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=app.ico
; 自动创建桌面快捷方式
DisableWelcomePage=no
CloseApplications=yes

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即运行"; Flags: postinstall nowait skipifsilent shellexec
