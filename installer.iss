; Inno Setup script - Invoice Printer
; Build with build_installer.bat (passes /DAppVersion=x.y.z)

#ifndef AppVersion
  #define AppVersion "2.1.0"
#endif

[Setup]
; Keep this GUID forever - it lets new versions upgrade the old install in place.
AppId={{8B3F6A52-4C1D-4E7A-9A55-2D7E1F0C9B31}
AppName=Invoice Printer
AppVersion={#AppVersion}
AppVerName=Invoice Printer {#AppVersion}
DefaultDirName={localappdata}\Programs\InvoicePrinter
DefaultGroupName=Invoice Printer
; Per-user install: no admin rights needed, and the app can write its config/log in its own folder.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=installer_output
OutputBaseFilename=InvoicePrinter_Setup_{#AppVersion}
SetupIconFile=app.ico
UninstallDisplayIcon={app}\InvoicePrinter.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
; Close a running copy during upgrade instead of failing on locked files
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Shortcuts:"

[Files]
; Whole PyInstaller folder. NO config file is shipped, so each PC starts blank and the
; first-run prompt asks for its own invoice folder. User settings/logs are never overwritten on upgrade.
Source: "dist\InvoicePrinter\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion
; Optional: PDF mode prints best with SumatraPDF. Put the portable SumatraPDF.exe in the "tools" folder before
; building the installer and it is installed next to the app (SumatraPDF is free, GPLv3).
Source: "tools\SumatraPDF.exe"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\Invoice Printer"; Filename: "{app}\InvoicePrinter.exe"
Name: "{autodesktop}\Invoice Printer"; Filename: "{app}\InvoicePrinter.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\InvoicePrinter.exe"; Description: "Launch Invoice Printer"; Flags: nowait postinstall skipifsilent

[Code]
// Friendly heads-up if desktop Excel is missing (install still continues).
function ExcelInstalled: Boolean;
begin
  Result := RegKeyExists(HKEY_CLASSES_ROOT, 'Excel.Application');
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if (CurPageID = wpFinished) and (not ExcelInstalled) then
    MsgBox('Microsoft Excel (desktop) was not found on this PC.' + #13#10 +
           'Invoice Printer needs it to print. Please install Excel before using the app.',
           mbInformation, MB_OK);
end;
