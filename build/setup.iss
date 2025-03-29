[Setup]
; 基本信息
AppName=CADtoExcel
AppVersion=1.0
AppPublisher=华阳通机电有限公司
AppPublisherURL=http://www.yourcompany.com
AppSupportURL=http://www.yourcompany.com/support
AppUpdatesURL=http://www.yourcompany.com/updates
DefaultDirName={commonpf}\CADtoExcel
DefaultGroupName=CADtoExcel
OutputDir=output
OutputBaseFilename=CADtoExcel_Trial_Setup
SetupIconFile=..\web\static\img\icon.ico
Compression=lzma
SolidCompression=yes
UninstallDisplayIcon={app}\CADtoExcel.exe
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

; 许可协议
LicenseFile=..\LICENSE.txt

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面图标"; GroupDescription: "附加图标:"; Flags: checkedonce
Name: "quicklaunchicon"; Description: "创建快速启动栏图标"; GroupDescription: "附加图标:"; Flags: checkedonce

[Files]
; 主程序
Source: "..\dist\app.exe"; DestDir: "{app}"; DestName: "CADtoExcel.exe"; Flags: ignoreversion

; 资源文件
Source: "..\maps\*"; DestDir: "{app}\maps"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\fonts\*"; DestDir: "{app}\fonts"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\web\static\*"; DestDir: "{app}\web\static"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\web\templates\*"; DestDir: "{app}\web\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

; 文档
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\CADtoExcel"; Filename: "{app}\CADtoExcel.exe"
Name: "{group}\帮助文档"; Filename: "{app}\README.md"
Name: "{group}\{cm:UninstallProgram,CADtoExcel}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\CADtoExcel"; Filename: "{app}\CADtoExcel.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\CADtoExcel"; Filename: "{app}\CADtoExcel.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\CADtoExcel.exe"; Description: "立即启动程序"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\CADtoExcel"

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // 安装后可执行的操作
  end;
end; 