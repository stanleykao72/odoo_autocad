; Inno Setup Script for Odoo AutoCAD Integration (C# Version)
; Version: 6.0.0 (.NET 8 / WPF / Self-Contained)

#define MyAppName "Odoo AutoCAD Integration"
#define MyAppVersion "6.0.0"
#define MyAppPublisher "承暉精品股份有限公司"
#define MyAppURL "https://odoo-esmith.odoo.com/"
#define MyAppExeName "OdooAutoCAD.exe"
#define MyAppDescription "Odoo and AutoCAD Integration Tool (C# Edition)"
#define PublishDir "C:\odoo\autocad_source\csharp\publish"

[Setup]
; Unique AppId for C# version (different from Python version)
AppId={{7F3A9B2C-D4E5-6789-ABCD-EF0123456789}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName=c:\odoo\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=yes
; Version info
VersionInfoVersion={#MyAppVersion}
VersionInfoTextVersion={#MyAppVersion}
VersionInfoDescription={#MyAppDescription}
VersionInfoCopyright=Copyright (C) 2026 承暉精品股份有限公司
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
; Architecture
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Output
OutputDir=C:\odoo\autocad_source\csharp\installer
OutputBaseFilename=odoo-autocad-integration-6.0.0-setup
SetupIconFile=C:\odoo\autocad_source\icon\odoo_autocad.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
; Compression - use lzma2/normal to avoid out-of-memory with large .NET publish
Compression=lzma2/normal
SolidCompression=yes
; UI
WizardStyle=modern
DisableProgramGroupPage=yes
; Privileges
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; All published files (self-contained .NET 8 deployment)
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Config directory from Python project (shared YAML configs)
Source: "C:\odoo\autocad_source\config\*.yaml"; DestDir: "{app}\config"; Flags: ignoreversion onlyifdoesntexist
; Icon files
Source: "C:\odoo\autocad_source\icon\odoo_autocad.ico"; DestDir: "{app}\icon"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon\odoo_autocad.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon\odoo_autocad.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\logs"
