; Inno Setup script for HFin WPF Application
#define MyAppName "HFin"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Hdevs"
#define MyAppExeName "HFin.exe"

[Setup]
AppId={{D8F8F982-FFDE-4E11-9A33-BC7C1D18AF10}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={code:GetDefaultDirName}
DefaultGroupName={#MyAppPublisher} office
DisableProgramGroupPage=yes
OutputDir=..\InstallerOutput
OutputBaseFilename=HFinSetup
SetupIconFile=logo.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "bin\Release\net10.0-windows\publish\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
function GetDefaultDirName(Param: String): String;
begin
  if DirExists('D:\') then
    Result := 'D:\Hdevs office'
  else
    Result := 'C:\Hdevs office';
end;
