[Setup]
AppName=Columns Report Addin
AppVersion=1.0.0
AppPublisher=KAITECH
DefaultDirName={commonpf}\ColumnsReportAddin
DefaultGroupName=Columns Report Addin
OutputBaseFilename=ColumnsReportAddin_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
OutputDir=Output

[Files]
Source: "bin\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Code]
procedure GenerateManifest(Year: String);
var
  Lines: TStringList;
  Dir, FilePath: String;
begin
  Dir := ExpandConstant('{commonappdata}\Autodesk\Revit\Addins\') + Year;
  FilePath := Dir + '\ColumnsReportAddin.addin';
  if not DirExists(Dir) then
    Exit;

  Lines := TStringList.Create;
  try
    Lines.Add('<?xml version="1.0" encoding="utf-8"?>');
    Lines.Add('<RevitAddIns>');
    Lines.Add('  <AddIn Type="Application">');
    Lines.Add('    <Name>Columns Report</Name>');
    Lines.Add('    <Assembly>' + ExpandConstant('{app}') + '\ColumnsReportAddin.dll</Assembly>');
    Lines.Add('    <AddInId>E8F3A2D1-7B4C-4E9A-B6D5-1C3F8A0E2D4B</AddInId>');
    Lines.Add('    <FullClassName>ColumnsReportAddin.App</FullClassName>');
    Lines.Add('    <VendorId>KAITECH</VendorId>');
    Lines.Add('    <VendorDescription>KAITECH</VendorDescription>');
    Lines.Add('  </AddIn>');
    Lines.Add('</RevitAddIns>');
    Lines.SaveToFile(FilePath);
  finally
    Lines.Free;
  end;
end;

procedure DeleteManifest(Year: String);
var
  FilePath: String;
begin
  FilePath := ExpandConstant('{commonappdata}\Autodesk\Revit\Addins\') + Year + '\ColumnsReportAddin.addin';
  if FileExists(FilePath) then
    DeleteFile(FilePath);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    GenerateManifest('2020');
    GenerateManifest('2021');
    GenerateManifest('2022');
    GenerateManifest('2023');
    GenerateManifest('2024');
    GenerateManifest('2025');
    GenerateManifest('2026');
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    DeleteManifest('2020');
    DeleteManifest('2021');
    DeleteManifest('2022');
    DeleteManifest('2023');
    DeleteManifest('2024');
    DeleteManifest('2025');
    DeleteManifest('2026');
  end;
end;
