; PADLE Installer Script for Inno Setup
; Panopto Audio Descriptions List Editor
;
; Prerequisites before compiling:
;   1. Build PADLE with PyInstaller (run build.bat)
;   2. Ensure dist\PADLE folder exists with PADLE.exe
;   3. Ensure icon.ico exists in the project root
;
; To compile:
;   1. Open this file in Inno Setup Compiler
;   2. Press Ctrl+F9 or Build > Compile
;   3. Output: Output\PADLE_Setup.exe

#define MyAppName "PADLE"
#define MyAppFullName "Panopto Audio Descriptions List Editor"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "University of Pittsburgh Center for Teaching and Learning"
#define MyAppURL "https://www.teaching.pitt.edu/"
#define MyAppExeName "PADLE.exe"

[Setup]
; Unique identifier for this application (generate new GUID for your own apps)
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Installation directory
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}

; Allow user to choose install location
DisableProgramGroupPage=yes
AllowNoIcons=yes

; Output settings
OutputDir=Output
OutputBaseFilename=PADLE_Setup_{#MyAppVersion}
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

; Compression (lzma2 gives best compression)
Compression=lzma2
SolidCompression=yes

; Require admin rights to install to Program Files
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

; Windows version requirements (Windows 10+)
MinVersion=10.0

; Installer appearance
WizardStyle=modern
WizardSizePercent=110

; License and info pages (optional - uncomment if you have these files)
; LicenseFile=LICENSE.txt
; InfoBeforeFile=README.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1

[Files]
; Main application files from PyInstaller output
Source: "dist\PADLE\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Voice download script (for users to download TTS voices)
Source: "download_voices.py"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu shortcuts
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Comment: "{#MyAppFullName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

; Desktop shortcut (if selected)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; Comment: "{#MyAppFullName}"

; Quick Launch shortcut (legacy, for older Windows)
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
; Option to launch app after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// Pascal script for dependency checking

var
  DependencyPage: TWizardPage;
  VLCCheck, TesseractCheck, FFmpegCheck, MoondreamCheck: TNewCheckBox;
  VLCStatus, TesseractStatus, FFmpegStatus, MoondreamStatus: TNewStaticText;

// Check if VLC is installed
function IsVLCInstalled: Boolean;
begin
  Result := FileExists('C:\Program Files\VideoLAN\VLC\vlc.exe') or
            FileExists('C:\Program Files (x86)\VideoLAN\VLC\vlc.exe');
end;

// Check if Tesseract is installed
function IsTesseractInstalled: Boolean;
begin
  Result := FileExists('C:\Program Files\Tesseract-OCR\tesseract.exe');
end;

// Check if FFmpeg is installed
function IsFFmpegInstalled: Boolean;
var
  Path: String;
begin
  // Check common locations
  Result := FileExists('C:\ffmpeg\bin\ffmpeg.exe') or
            FileExists('C:\Program Files\ffmpeg\bin\ffmpeg.exe');
  
  // Also check if it's in PATH
  if not Result then
  begin
    Path := GetEnv('PATH');
    Result := Pos('ffmpeg', LowerCase(Path)) > 0;
  end;
end;

// Check if Python/pip is available for Moondream
function IsPythonInstalled: Boolean;
var
  ResultCode: Integer;
begin
  Result := Exec('cmd.exe', '/c python --version', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := Result and (ResultCode = 0);
end;

// Update status labels with colored indicators
procedure UpdateStatusLabel(StatusLabel: TNewStaticText; IsInstalled: Boolean);
begin
  if IsInstalled then
  begin
    StatusLabel.Caption := 'Installed';
    StatusLabel.Font.Color := clGreen;
  end
  else
  begin
    StatusLabel.Caption := 'Not Found';
    StatusLabel.Font.Color := clRed;
  end;
end;

// Create the dependency check page
procedure CreateDependencyPage;
var
  TopPos: Integer;
  LabelWidth: Integer;
begin
  DependencyPage := CreateCustomPage(wpSelectDir,
    'Check Dependencies',
    'PADLE requires the following software to be installed.');

  TopPos := 8;
  LabelWidth := 300;

  // Introduction text
  with TNewStaticText.Create(DependencyPage) do
  begin
    Parent := DependencyPage.Surface;
    Caption := 'The following dependencies are required for full functionality.' + #13#10 +
               'Check any items below to open their download pages after installation.';
    Left := 0;
    Top := TopPos;
    Width := DependencyPage.SurfaceWidth;
    Height := 40;
    WordWrap := True;
  end;
  TopPos := TopPos + 50;

  // VLC
  VLCCheck := TNewCheckBox.Create(DependencyPage);
  with VLCCheck do
  begin
    Parent := DependencyPage.Surface;
    Caption := 'VLC Media Player (required for audio playback)';
    Left := 0;
    Top := TopPos;
    Width := LabelWidth;
    Checked := not IsVLCInstalled;
    Enabled := not IsVLCInstalled;
  end;
  
  VLCStatus := TNewStaticText.Create(DependencyPage);
  with VLCStatus do
  begin
    Parent := DependencyPage.Surface;
    Left := LabelWidth + 20;
    Top := TopPos + 2;
    Width := 100;
    Font.Style := [fsBold];
  end;
  UpdateStatusLabel(VLCStatus, IsVLCInstalled);
  TopPos := TopPos + 28;

  // Tesseract
  TesseractCheck := TNewCheckBox.Create(DependencyPage);
  with TesseractCheck do
  begin
    Parent := DependencyPage.Surface;
    Caption := 'Tesseract OCR (required for Slide+OCR mode)';
    Left := 0;
    Top := TopPos;
    Width := LabelWidth;
    Checked := not IsTesseractInstalled;
    Enabled := not IsTesseractInstalled;
  end;
  
  TesseractStatus := TNewStaticText.Create(DependencyPage);
  with TesseractStatus do
  begin
    Parent := DependencyPage.Surface;
    Left := LabelWidth + 20;
    Top := TopPos + 2;
    Width := 100;
    Font.Style := [fsBold];
  end;
  UpdateStatusLabel(TesseractStatus, IsTesseractInstalled);
  TopPos := TopPos + 28;

  // FFmpeg
  FFmpegCheck := TNewCheckBox.Create(DependencyPage);
  with FFmpegCheck do
  begin
    Parent := DependencyPage.Surface;
    Caption := 'FFmpeg (required for MP3 audio export)';
    Left := 0;
    Top := TopPos;
    Width := LabelWidth;
    Checked := not IsFFmpegInstalled;
    Enabled := not IsFFmpegInstalled;
  end;
  
  FFmpegStatus := TNewStaticText.Create(DependencyPage);
  with FFmpegStatus do
  begin
    Parent := DependencyPage.Surface;
    Left := LabelWidth + 20;
    Top := TopPos + 2;
    Width := 100;
    Font.Style := [fsBold];
  end;
  UpdateStatusLabel(FFmpegStatus, IsFFmpegInstalled);
  TopPos := TopPos + 28;

  // Moondream (Python)
  MoondreamCheck := TNewCheckBox.Create(DependencyPage);
  with MoondreamCheck do
  begin
    Parent := DependencyPage.Surface;
    Caption := 'Python + Moondream AI (required for AI descriptions)';
    Left := 0;
    Top := TopPos;
    Width := LabelWidth;
    Checked := not IsPythonInstalled;
    Enabled := True; // Always allow checking - complex install
  end;
  
  MoondreamStatus := TNewStaticText.Create(DependencyPage);
  with MoondreamStatus do
  begin
    Parent := DependencyPage.Surface;
    Left := LabelWidth + 20;
    Top := TopPos + 2;
    Width := 100;
    Font.Style := [fsBold];
  end;
  if IsPythonInstalled then
  begin
    MoondreamStatus.Caption := 'Python OK';
    MoondreamStatus.Font.Color := clGreen;
  end
  else
  begin
    MoondreamStatus.Caption := 'Python Missing';
    MoondreamStatus.Font.Color := clRed;
  end;
  TopPos := TopPos + 40;

  // Note about Moondream
  with TNewStaticText.Create(DependencyPage) do
  begin
    Parent := DependencyPage.Surface;
    Caption := 'Note: Moondream AI must be running separately before using PADLE.' + #13#10 +
               'See the documentation for Moondream setup instructions.';
    Left := 0;
    Top := TopPos;
    Width := DependencyPage.SurfaceWidth;
    Height := 40;
    Font.Color := clGray;
    WordWrap := True;
  end;
end;

// Open download pages for selected dependencies
procedure OpenSelectedDownloads;
var
  ErrorCode: Integer;
begin
  if VLCCheck.Checked and VLCCheck.Enabled then
    ShellExec('open', 'https://www.videolan.org/vlc/', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
  
  if TesseractCheck.Checked and TesseractCheck.Enabled then
    ShellExec('open', 'https://github.com/UB-Mannheim/tesseract/wiki', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
  
  if FFmpegCheck.Checked and FFmpegCheck.Enabled then
    ShellExec('open', 'https://www.gyan.dev/ffmpeg/builds/', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
  
  if MoondreamCheck.Checked then
    ShellExec('open', 'https://www.python.org/downloads/', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
end;

// Initialize wizard
procedure InitializeWizard;
begin
  CreateDependencyPage;
end;

// Handle finish - open download pages
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    OpenSelectedDownloads;
  end;
end;

// Uninstall - clean up app data (optional)
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDataPath: String;
  VoicesPath: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Ask if user wants to remove app data
    AppDataPath := ExpandConstant('{localappdata}\padle');
    VoicesPath := ExpandConstant('{localappdata}\piper-voices');
    
    if DirExists(AppDataPath) then
    begin
      if MsgBox('Do you want to remove PADLE application data?', mbConfirmation, MB_YESNO) = IDYES then
      begin
        DelTree(AppDataPath, True, True, True);
      end;
    end;
    
    // Note: We don't automatically remove piper-voices since they might be used by other apps
  end;
end;
