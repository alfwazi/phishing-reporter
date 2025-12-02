; Inno Setup Script for Phishing Reporter
; Creates standalone installer - no Visual Studio required

#define AppName "Phishing Reporter"
#define AppVersion "1.1.0"
#define AppPublisher "Geidea"
#define AppURL "https://github.com/alfwazi/phishing-reporter"
#define AppExeName "PhishingReporter.dll"

[Setup]
AppId={{A5AE1C29-18E0-4E3E-A3E1-AC8EE23FF9A6}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
DefaultDirName={pf}\Geidea\PhishingReporter
DefaultGroupName=Phishing Reporter
AllowNoIcons=yes
;LicenseFile=..\LICENSE
OutputDir=Release
OutputBaseFilename=PhishingReporter-Setup
Compression=zip
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
;SetupIconFile=..\phishing.ico
;WizardImageFile=..\splash.jpg
;WizardSmallImageFile=..\phishing.png

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\PhishingReporter\bin\Release\PhishingReporter.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\PhishingReporter\bin\Release\PhishingReporter.dll.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\packages\HtmlAgilityPack.1.12.2\lib\Net45\HtmlAgilityPack.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"

[Run]
; Optional: Add post-install tasks here
; Filename: "{app}\register.vbs"; Description: "Register with Outlook"; Flags: runhidden

[Code]
function InitializeSetup(): Boolean;
var
  OutlookPath: String;
begin
  Result := True;
  // Check if Outlook is installed
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE', '', OutlookPath) then
  begin
    if MsgBox('Microsoft Outlook was not detected. The add-in requires Outlook to function. Continue anyway?', mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

