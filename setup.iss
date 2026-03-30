; Inno Setup Script for 협력업체 관리 프로그램

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
AppId={{F22A892D-A469-4A39-8C2E-389589548239}}
AppName=협력업체 관리 프로그램
AppVersion=2.0
DefaultDirName={autopf}\협력업체 관리 프로그램
DefaultGroupName=협력업체 관리 프로그램
UninstallDisplayIcon={app}\icon.ico
SetupIconFile=icon.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
OutputBaseFilename=협력업체_관리_프로그램_v2.0_설치
OutputDir=installer

[Languages]
Name: "korean"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}";

[Files]
Source: "icon.ico"; DestDir: "{app}"
Source: "build\main\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\협력업체 관리 프로그램"; Filename: "{app}\협력업체 관리 프로그램.exe"
Name: "{autodesktop}\협력업체 관리 프로그램"; Filename: "{app}\협력업체 관리 프로그램.exe"; Tasks: desktopicon
