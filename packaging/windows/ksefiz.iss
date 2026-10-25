; Inno Setup script: per-user installer for the KSeF IZ Tool (no admin rights).
;   iscc /DAppVersion=0.2.0 packaging\windows\ksefiz.iss
; Expects dist\ksefiz\ (PyInstaller, see ksefiz.spec) and, if present,
; excel\dist\KSeF_IZ_Tool.xlsm (scripts\build_xlsm.ps1). Paths match src/ksefiz/userconfig.py.

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#define AppName "KSeF IZ Tool"
#define Root AddBackslash(SourcePath) + "..\.."
#define Workbook Root + "\excel\dist\KSeF_IZ_Tool.xlsm"
#define TrustKey "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\KSeFIZTool"

[Setup]
AppId={{6F0B2E8B-6C4E-4F51-9C2D-3B7A1E5D9A41}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=KSeF IZ Tool (proof of concept)
DefaultDirName={localappdata}\Programs\{#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=no
PrivilegesRequired=lowest
OutputDir={#Root}\dist
OutputBaseFilename=KSeF-IZ-Tool-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\ksefiz.exe
SetupLogging=yes

[Messages]
WelcomeLabel2=This installs the KSeF IZ Tool for your Windows user account (no administrator rights needed):%n%n- the ksefiz helper that connects Excel to KSeF,%n- the KSeF IZ Tool workbook in %USERPROFILE%\{#AppName},%n- an Excel Trusted Location for that folder, so the workbook's macros run.%n%nA setup assistant starts at the end to connect your company's KSeF token.

[Files]
Source: "{#Root}\dist\ksefiz\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#Root}\docs\USER_GUIDE.md"; DestDir: "{app}"; Flags: ignoreversion
#if FileExists(Workbook)
; never removed by the uninstaller: it may contain the user's work
Source: "{#Workbook}"; DestDir: "{%USERPROFILE}\{#AppName}"; Flags: ignoreversion uninsneveruninstall
#endif

[Dirs]
Name: "{%USERPROFILE}\{#AppName}"; Flags: uninsneveruninstall

[Registry]
; Excel Trusted Location for the workbook folder only (no subfolders), current user only
Root: HKCU; Subkey: "{#TrustKey}"; ValueType: string; ValueName: "Path"; ValueData: "{%USERPROFILE}\{#AppName}\"; Flags: uninsdeletekey
Root: HKCU; Subkey: "{#TrustKey}"; ValueType: dword; ValueName: "AllowSubfolders"; ValueData: 0
Root: HKCU; Subkey: "{#TrustKey}"; ValueType: string; ValueName: "Description"; ValueData: "{#AppName} workbook (added by the installer)"

[Icons]
Name: "{userprograms}\{#AppName}\{#AppName}"; Filename: "{%USERPROFILE}\{#AppName}\KSeF_IZ_Tool.xlsm"
Name: "{userprograms}\{#AppName}\{#AppName} setup"; Filename: "{app}\KSeF IZ Setup.exe"
Name: "{userprograms}\{#AppName}\User guide"; Filename: "{app}\USER_GUIDE.md"
Name: "{userprograms}\{#AppName}\Uninstall {#AppName}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\KSeF IZ Setup.exe"; Description: "Start the setup assistant (connect your KSeF token)"; Flags: postinstall nowait skipifsilent

[Code]
// An existing workbook may hold the user's work: keep it as a dated backup before replacing.
procedure CurStepChanged(CurStep: TSetupStep);
var
  WbPath, Backup: String;
begin
  if CurStep = ssInstall then
  begin
    WbPath := ExpandConstant('{%USERPROFILE}\{#AppName}\KSeF_IZ_Tool.xlsm');
    if FileExists(WbPath) then
    begin
      Backup := ExpandConstant('{%USERPROFILE}\{#AppName}\KSeF_IZ_Tool.backup-') +
                GetDateTimeString('yyyymmdd-hhnnss', #0, #0) + '.xlsm';
      if not RenameFile(WbPath, Backup) then
        MsgBox('The existing workbook could not be backed up (is it open in Excel?). ' +
               'Close Excel and run the installer again.', mbError, MB_OK)
      else
        Log('existing workbook backed up to ' + Backup);
    end;
  end;
end;
