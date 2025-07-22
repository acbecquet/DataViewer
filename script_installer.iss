; TestingGUI Professional Installer Script
; Generated for production deployment

[Setup]
AppName=Standardized Testing GUI
AppVersion=3.0.0
AppPublisher=Charlie Becquet
AppPublisherURL=https://your-website.com
AppSupportURL=https://your-website.com/support
AppUpdatesURL=https://your-website.com/updates
DefaultDirName={autopf}\TestingGUI
DefaultGroupName=Testing GUI
AllowNoIcons=yes
LicenseFile=LICENSE.txt
InfoBeforeFile=README.txt
OutputDir=installer_output
OutputBaseFilename=TestingGUI_Setup_v3.0.0
SetupIconFile=resources\icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
MinVersion=6.1sp1
PrivilegesRequired=admin
UninstallDisplayIcon={app}\TestingGUI.exe

; License key validation
[Code]
var
  LicenseKeyPage: TInputQueryWizardPage;
  
function IsValidLicenseKey(Key: String): Boolean;
begin
  // Simple validation - replace with your actual validation logic
  // This is a placeholder - implement your real license validation
  Result := (Length(Key) = 16) and (Pos('-', Key) > 0);
  
  // Advanced validation example:
  // Result := (Key = 'DEMO-1234-5678-ABCD') or 
  //           (Key = 'FULL-9876-5432-WXYZ') or
  //           ValidateKeyWithServer(Key);
end;

procedure InitializeWizard;
begin
  LicenseKeyPage := CreateInputQueryPage(wpLicense,
    'License Key Required', 'Please enter your license key',
    'A valid license key is required to install this software. ' +
    'Contact support@your-website.com if you need assistance.');
  LicenseKeyPage.Add('License Key:', False);
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  
  if CurPageID = LicenseKeyPage.ID then
  begin
    if not IsValidLicenseKey(LicenseKeyPage.Values[0]) then
    begin
      MsgBox('Invalid license key. Please check your key and try again.', 
             mbError, MB_OK);
      Result := False;
    end;
  end;
end;

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1

[Files]
Source: "dist\TestingGUI.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\README.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "resources\*"; DestDir: "{app}\resources"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Testing GUI"; Filename: "{app}\TestingGUI.exe"
Name: "{group}\{cm:ProgramOnTheWeb,Testing GUI}"; Filename: "https://your-website.com"
Name: "{group}\{cm:UninstallProgram,Testing GUI}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Testing GUI"; Filename: "{app}\TestingGUI.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\Testing GUI"; Filename: "{app}\TestingGUI.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\TestingGUI.exe"; Description: "{cm:LaunchProgram,Testing GUI}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"