; Inno Setup script for tANk.
; Installs for the current user only, so Windows never asks for an administrator password.
; Everything it does: copy manifest.xml into the user's AppData, add one registry value that
; tells Office where to find it, and clear Office's add-in cache so the button shows up.

#define AppName "tANk"
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#define Publisher "tANk"
#define AddinId "7f3c1a54-9e2b-4d61-8a77-2c5b9e0f41d3"
; The build passes the repo URL in: ISCC /DSupportUrl="https://github.com/owner/repo"
#ifndef SupportUrl
  #define SupportUrl "https://github.com"
#endif

[Setup]
AppId={{9C4A1E30-7F2B-4E56-9E0D-1B7A6C2F8D41}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#Publisher}
AppSupportURL={#SupportUrl}
DefaultDirName={localappdata}\{#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\release
OutputBaseFilename=tANk-Setup
SetupIconFile=..\public\assets\icon-128.ico
UninstallDisplayIcon={app}\icon.ico
UninstallDisplayName={#AppName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\dist\manifest.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\public\assets\icon-128.ico"; DestDir: "{app}"; DestName: "icon.ico"; Flags: ignoreversion

[Registry]
; This is the whole installation as far as Office is concerned.
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\WEF\Developer"; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; \
  ValueName: "{#AddinId}"; ValueData: "{app}\manifest.xml"; Flags: uninsdeletevalue

[Code]
function OfficeIsRunning(): Boolean;
var
  ResultCode: Integer;
begin
  { tasklist returns 0 when it finds the process }
  Exec(ExpandConstant('{cmd}'),
    '/C tasklist /FI "IMAGENAME eq EXCEL.EXE" /FI "IMAGENAME eq WINWORD.EXE" | find /I "EXE" > nul',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := ResultCode = 0;
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if OfficeIsRunning() then
  begin
    if MsgBox('Excel or Word is still open.' + #13#10 + #13#10 +
              'tANk will not appear until they are restarted. Close them now, then click Yes to carry on.',
              mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

procedure ClearRibbonCache();
var
  Base: String;
begin
  { AppCommands and AddinInfo hold the cached ribbon buttons and Office rebuilds them on the
    next start, so clearing them is harmless. Wef\webview2 is deliberately left alone:
    every add-in on this machine keeps its saved data in there. }
  Base := ExpandConstant('{localappdata}\Microsoft\Office\16.0\Wef');
  if DirExists(Base + '\AppCommands') then
    DelTree(Base + '\AppCommands\*', False, True, True);
  if DirExists(Base + '\AddinInfo') then
    DelTree(Base + '\AddinInfo\*', False, True, True);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    ClearRibbonCache();
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  if OfficeIsRunning() then
  begin
    MsgBox('Close Excel, Word and PowerPoint first, then run the uninstaller again.',
           mbError, MB_OK);
    Result := False;
    Exit;
  end;

  { The pane stores the API keys itself, in Office's own web storage, which no installer
    can pick apart without wiping every other add-in's data too. }
  if MsgBox('Remove tANk?' + #13#10 + #13#10 +
            'The button, the add-in registration and its files all go.' + #13#10 + #13#10 +
            'Your saved API keys live inside the pane, not in these files. To clear those as' + #13#10 +
            'well, choose No, open tANk in Excel, go to Settings and press "Erase everything",' + #13#10 +
            'then run this uninstaller again.',
            mbConfirmation, MB_YESNO) = IDNO then
    Result := False;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
    ClearRibbonCache();
end;

[Run]
Filename: "{#SupportUrl}/blob/main/INSTALL.md"; \
  Description: "Show me how to add a free API key"; Flags: postinstall shellexec nowait unchecked

[Messages]
FinishedLabel=tANk is installed.%n%nOpen Excel, Word or PowerPoint and click tANk on the Home tab. The first time, open Settings in the pane and paste a free API key.
