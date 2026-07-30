#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif

#define MyAppName "大統進貨助手"
#define MyAppPublisher "ontaiko"
#define MyAppExeName "大統進貨助手.exe"
#define MyAppId "{{6E3B5C4E-6FAF-4B7B-A6C0-410EA98F3828}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir=..\dist-installer
OutputBaseFilename=Datong-Invoice-Assistant-Setup-v{#MyAppVersion}
SetupLogging=yes
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} 一鍵安裝程式
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}

[Tasks]
Name: "desktopicon"; Description: "建立桌面捷徑"; GroupDescription: "其他選項："; Flags: checkedonce

[Files]
Source: "..\大統進貨助手.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\package-manifest.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\app_settings.example.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\scripts\*"; DestDir: "{app}\scripts"; Excludes: "__pycache__\*,*.pyc"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\reference_data\*"; DestDir: "{app}\reference_data"; Flags: onlyifdoesntexist uninsneveruninstall recursesubdirs createallsubdirs
Source: "..\installer\setup-runtime.ps1"; DestDir: "{app}\installer"; Flags: ignoreversion
Source: "..\installer\verify-installation.ps1"; DestDir: "{app}\installer"; Flags: ignoreversion
Source: "..\..\engine\requirements-ocr.txt"; DestDir: "{app}\engine"; Flags: ignoreversion
Source: "..\..\engine\model-manifest.json"; DestDir: "{app}\engine"; Flags: ignoreversion
Source: "..\..\engine\official_models\*"; DestDir: "{app}\engine\official_models"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\驗證安裝"; Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\verify-installation.ps1"" -InstallDir ""{app}"""; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[UninstallDelete]
Type: filesandordirs; Name: "{app}\engine\.venv"
Type: files; Name: "{app}\install-runtime.log"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "啟動 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
procedure ExitProcess(ExitCode: Integer);
  external 'ExitProcess@kernel32.dll stdcall';

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  PowerShellPath: String;
  Parameters: String;
begin
  if CurStep = ssPostInstall then
  begin
    WizardForm.StatusLabel.Caption := '正在安裝本機 OCR 引擎，第一次安裝可能需要數分鐘...';
    PowerShellPath := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
    Parameters :=
      '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' +
      ExpandConstant('{app}\installer\setup-runtime.ps1') +
      '" -InstallDir "' + ExpandConstant('{app}') + '"';
    if (not Exec(PowerShellPath, Parameters, ExpandConstant('{app}'), SW_HIDE,
      ewWaitUntilTerminated, ResultCode)) or (ResultCode <> 0) then
    begin
      MsgBox(
        'OCR 引擎安裝失敗。請查看安裝記錄：' + #13#10 +
        ExpandConstant('{app}\install-runtime.log'),
        mbError, MB_OK
      );
      ExitProcess(1);
    end;
  end;
end;
