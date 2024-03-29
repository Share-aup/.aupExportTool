; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName ".aupExportTool"
#define MyAppVersion "1.2.0"
#define MyAppPublisher "Share.aup"
#define MyAppURL "https://share-aup.com/download/aup_exp_tool"
#define MyAppExeName "aupExportTool.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{E82AA7BE-3B7E-43D8-AC65-73D30B7C9ABE}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=C:\Users\yashi\source\repos\aupExportTool\installer_license.txt
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
OutputDir=C:\Users\yashi\Desktop
OutputBaseFilename=win-aupExportTool.1-2-0
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\yashi\source\repos\aupExportTool\aup_data\bin\Release\net6.0-windows\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\yashi\source\repos\aupExportTool\aup_data\bin\Release\net6.0-windows\aupExportTool.deps.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\yashi\source\repos\aupExportTool\aup_data\bin\Release\net6.0-windows\aupExportTool.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\yashi\source\repos\aupExportTool\aup_data\bin\Release\net6.0-windows\aupExportTool.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\yashi\source\repos\aupExportTool\aup_data\bin\Release\net6.0-windows\aupExportTool.runtimeconfig.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\yashi\source\repos\aupExportTool\aup_data\bin\Release\net6.0-windows\Karoterra.AupDotNet.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\yashi\source\repos\aupExportTool\aup_data\bin\Release\net6.0-windows\Karoterra.AupDotNet.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\yashi\source\repos\aupExportTool\aup_data\bin\Release\net6.0-windows\Karoterra.AupDotNet.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\yashi\source\repos\aupExportTool\aup_data\bin\Release\net6.0-windows\MetroSet UI.dll"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

