[Setup]
AppName=연락처 열 삭제 도구
AppVersion=1.0
DefaultDirName={pf}\ContactRemover
DefaultGroupName=연락처 열 삭제 도구
OutputDir=Output
OutputBaseFilename=ContactRemoverSetup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64

[Files]
Source: "dist\contact_remover.exe"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; xlsx
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\RemoveContact"; ValueType: string; ValueName: ""; ValueData: "연락처 자동 삭제 실행"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\RemoveContact"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\RemoveContact\command"; ValueType: string; ValueName: ""; ValueData: """{app}\contact_remover.exe"" ""%1"""; Flags: uninsdeletekey

; xls
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\RemoveContact"; ValueType: string; ValueName: ""; ValueData: "연락처 자동 삭제 실행"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\RemoveContact"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\RemoveContact\command"; ValueType: string; ValueName: ""; ValueData: """{app}\contact_remover.exe"" ""%1"""; Flags: uninsdeletekey

; directories
Root: HKCR; Subkey: "Directory\shell\RemoveContact"; ValueType: string; ValueName: ""; ValueData: "연락처 자동 삭제 실행"; Flags: uninsdeletekey
Root: HKCR; Subkey: "Directory\shell\RemoveContact"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "Directory\shell\RemoveContact\command"; ValueType: string; ValueName: ""; ValueData: """{app}\contact_remover.exe"" ""%1"""; Flags: uninsdeletekey

[Icons]
Name: "{sendto}\연락처 자동 삭제 실행"; Filename: "{app}\contact_remover.exe"
Name: "{group}\연락처 자동 삭제 도구"; Filename: "{app}\contact_remover.exe"


