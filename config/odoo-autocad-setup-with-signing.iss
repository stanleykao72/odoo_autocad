; Inno Setup Script with Code Signing Configuration
; 包含程式碼簽章配置的 Inno Setup 腳本

#define MyAppName "Odoo and AutoCAD Integration"
#define MyAppVersion "3.0"
#define MyAppPublisher "承暉精品股份有限公司"
#define MyAppURL "https://odoo-esmith.odoo.com/"
#define MyAppExeName "odoo-autocad-integration.exe"
#define MyAppAssocName MyAppName + " File"
#define MyAppAssocExt ".myp"
#define MyAppAssocKey StringChange(MyAppAssocName, " ", "") + MyAppAssocExt

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
AppId={{BB4DEBAA-C42E-4501-BD37-747E47153D45}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName=c:\odoo\{#MyAppName}
DisableDirPage=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
ChangesAssociations=yes
DisableProgramGroupPage=yes
OutputDir=C:\odoo\autocad_source\installer
OutputBaseFilename=odoo-autocad-integration-3.0-setup
SetupIconFile=C:\odoo\autocad_source\icon\odoo_autocad.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; 版本資訊
VersionInfoVersion={#MyAppVersion}
VersionInfoTextVersion={#MyAppVersion}
VersionInfoDescription=Odoo AutoCAD Integration Tool
VersionInfoCopyright=Copyright (C) 2025 承暉精品股份有限公司

; === 程式碼簽章配置 ===
; 請根據您的憑證類型選擇其中一種方法，並移除前面的分號

; 方法1: 使用 PFX 檔案 (需要密碼)
; SignTool=signtool /f "C:\certs\company.pfx" /p "your_password" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 $f

; 方法2: 使用環境變數保護密碼
; 先設定: set CERT_PASSWORD=your_password
; SignTool=signtool /f "C:\certs\company.pfx" /p "%CERT_PASSWORD%" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 $f

; 方法3: 使用憑證存放區中的憑證 (依憑證主體名稱)
; SignTool=signtool /n "承暉精品股份有限公司" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 $f

; 方法4: 使用憑證指紋 (最安全)
; SignTool=signtool /sha1 "certificate_thumbprint_here" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 $f

; 方法5: 簡化版本 (使用舊格式時間戳記)
; SignTool=signtool /f "C:\certs\company.pfx" /p "your_password" /t "http://timestamp.digicert.com" $f

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\odoo\autocad_source\output\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\odoo\autocad_source\output\*.md"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "C:\odoo\autocad_source\output\*.json"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[Registry]
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocExt}\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppAssocKey}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}"; ValueType: string; ValueName: ""; ValueData: "{#MyAppAssocName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".myp"; ValueData: ""

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// 自訂安裝邏輯
function InitializeSetup(): Boolean;
begin
  Result := True;
  // 檢查 .NET Framework 或其他相依性
end;