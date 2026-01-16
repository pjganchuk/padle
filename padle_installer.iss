; PADLE Installer Script for Inno Setup
; Panopto Audio Descriptions List Editor
;
; Requirements:
; - Inno Setup 6.x (https://jrsoftware.org/isinfo.php)
; - Built PADLE from PyInstaller (dist/PADLE folder)
; - Tesseract OCR installer or portable version
;
; Build with: ISCC padle_installer.iss

#define MyAppName "PADLE"
#define MyAppFullName "Panopto Audio Descriptions List Editor"
#define MyAppVersion "1.1.0"
#define MyAppPublisher "University of Pittsburgh Center for Teaching and Learning"
#define MyAppURL "https://github.com/pitt-ctl/padle"
#define MyAppExeName "PADLE.exe"

[Setup]
; Unique identifier for this application
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Installation paths
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes

; Output settings
OutputDir=installer_output
OutputBaseFilename=PADLE_Setup_{#MyAppVersion}
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

; Compression
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes

; Modern installer appearance
WizardStyle=modern
WizardSizePercent=120

; Privileges - per-user install by default (no admin required)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; Minimum Windows version (Windows 10)
MinVersion=10.0

; Misc
DisableProgramGroupPage=yes
DisableWelcomePage=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
WelcomeLabel2=This will install [name/ver] on your computer.%n%n{#MyAppFullName} creates audio descriptions for recorded lecture videos, compatible with Panopto's WebVTT format.%n%nThe AI model (~3.7GB) will be downloaded on first launch.

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main application (from PyInstaller dist folder)
Source: "dist\PADLE\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Tesseract OCR (portable version)
Source: "tesseract\*"; DestDir: "{app}\tesseract"; Flags: ignoreversion recursesubdirs createallsubdirs

; Piper TTS voices directory placeholder
; Users will download voices on first use

[Dirs]
; Create voices directory
Name: "{localappdata}\piper-voices"; Flags: uninsneveruninstall

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Comment: "{#MyAppFullName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Comment: "{#MyAppFullName}"; Tasks: desktopicon

[Registry]
; Add Tesseract to app's local path
Root: HKCU; Subkey: "Environment"; ValueType: string; ValueName: "PADLE_TESSERACT"; ValueData: "{app}\tesseract\tesseract.exe"; Flags: uninsdeletevalue

[Run]
; Option to launch after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up autosave files
Type: filesandordirs; Name: "{userappdata}\padle"

[Code]
var
  VLCInstalled: Boolean;
  FFmpegInstalled: Boolean;

// Check if VLC is installed
function IsVLCInstalled(): Boolean;
begin
  Result := FileExists(ExpandConstant('{pf}\VideoLAN\VLC\vlc.exe')) or
            FileExists(ExpandConstant('{pf32}\VideoLAN\VLC\vlc.exe'));
end;

// Check if FFmpeg is in PATH or common locations
function IsFFmpegInstalled(): Boolean;
begin
  Result := False;
  
  // Check common locations
  if FileExists('C:\ffmpeg\bin\ffmpeg.exe') or
     FileExists(ExpandConstant('{localappdata}\ffmpeg\bin\ffmpeg.exe')) or
     FileExists(ExpandConstant('{pf}\ffmpeg\bin\ffmpeg.exe')) then
  begin
    Result := True;
  end;
end;

// After installation, show info about missing dependencies
procedure CurStepChanged(CurStep: TSetupStep);
var
  Msg: String;
begin
  if CurStep = ssPostInstall then
  begin
    VLCInstalled := IsVLCInstalled();
    FFmpegInstalled := IsFFmpegInstalled();
    
    if not VLCInstalled or not FFmpegInstalled then
    begin
      Msg := 'PADLE installed successfully!' + #13#10 + #13#10;
      Msg := Msg + 'OPTIONAL DEPENDENCIES' + #13#10;
      Msg := Msg + '================================' + #13#10 + #13#10;
      
      if not VLCInstalled then
      begin
        Msg := Msg + 'VLC Media Player (for video audio playback)' + #13#10;
        Msg := Msg + '  Download: https://www.videolan.org/vlc/' + #13#10;
        Msg := Msg + '  Winget:   winget install VideoLAN.VLC' + #13#10 + #13#10;
      end;
      
      if not FFmpegInstalled then
      begin
        Msg := Msg + 'FFmpeg (for MP3 audio export)' + #13#10;
        Msg := Msg + '  Download: https://ffmpeg.org/download.html' + #13#10;
        Msg := Msg + '  Winget:   winget install Gyan.FFmpeg' + #13#10 + #13#10;
      end;
      
      Msg := Msg + '================================' + #13#10;
      Msg := Msg + 'Note: The AI model (~3.7GB) will download on first launch.';
      
      MsgBox(Msg, mbInformation, MB_OK);
    end;
  end;
end;

// Uninstall: Ask about removing model cache
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  CachePath: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    CachePath := ExpandConstant('{%USERPROFILE}\.cache\huggingface\hub\models--vikhyatk--moondream2');
    
    if DirExists(CachePath) then
    begin
      if MsgBox('Do you want to remove the downloaded AI model cache?' + #13#10 +
                '(This will free up ~3.7GB of disk space)' + #13#10 + #13#10 +
                'Location: ' + CachePath,
                mbConfirmation, MB_YESNO) = IDYES then
      begin
        DelTree(CachePath, True, True, True);
      end;
    end;
  end;
end;
