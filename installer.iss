[Setup]
AppName=연락처 자동 삭제 및 마스킹 도구 V2 Pro
AppVersion=2.0
DefaultDirName={pf}\ContactRemoverV2
DefaultGroupName=연락처 자동 삭제 및 마스킹 도구 V2
OutputDir=Output
OutputBaseFilename=ContactRemoverSetup_V2Pro
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64

[Files]
Source: "dist\contact_remover\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
; xlsx
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\RemoveContact"; ValueType: string; ValueName: ""; ValueData: "연락처 자동 삭제 및 마스킹 도구 실행"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\RemoveContact"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\RemoveContact\command"; ValueType: string; ValueName: ""; ValueData: """{app}\contact_remover.exe"" ""%1"""; Flags: uninsdeletekey

; xls
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\RemoveContact"; ValueType: string; ValueName: ""; ValueData: "연락처 자동 삭제 및 마스킹 도구 실행"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\RemoveContact"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\RemoveContact\command"; ValueType: string; ValueName: ""; ValueData: """{app}\contact_remover.exe"" ""%1"""; Flags: uninsdeletekey

; directories
Root: HKCR; Subkey: "Directory\shell\RemoveContact"; ValueType: string; ValueName: ""; ValueData: "연락처 자동 삭제 및 마스킹 도구 실행"; Flags: uninsdeletekey
Root: HKCR; Subkey: "Directory\shell\RemoveContact"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "Directory\shell\RemoveContact\command"; ValueType: string; ValueName: ""; ValueData: """{app}\contact_remover.exe"" ""%1"""; Flags: uninsdeletekey

[Icons]
Name: "{sendto}\연락처 자동 삭제 및 마스킹 도구 실행"; Filename: "{app}\contact_remover.exe"
Name: "{group}\연락처 자동 삭제 및 마스킹 도구 V2 Pro"; Filename: "{app}\contact_remover.exe"
