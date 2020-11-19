; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=ReNamer
AppVerName=ReNamer 7.2
AppPublisher=Ferla Daniele
AppPublisherURL=http://www.desdinova.it
AppSupportURL=http://www.desdinova.it
AppUpdatesURL=http://www.desdinova.it
DefaultDirName={pf}\ReNamer
DefaultGroupName=ReNamer 7.2
OutputBaseFilename=ReNamer 7.2
Compression=lzma
SolidCompression=true
AppCopyright=2008 by Ferla Daniele
WizardImageFile=C:\Programmi\Inno Setup 5\WizModernImage-IS.bmp
WizardSmallImageFile=C:\Programmi\Inno Setup 5\WizModernSmallImage-IS.bmp
AppVersion=7.2
UninstallDisplayIcon={app}\ReNamer.exe
UninstallDisplayName=ReNamer
LicenseFile=License.txt

[Languages]
Name: english; MessagesFile: compiler:Default.isl

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}

[Files]
Source: Install\Support\asycfilt.dll; DestDir: {app}
Source: Install\Support\COMCAT.DLL; DestDir: {app}
Source: Install\Support\DAO350.DLL; DestDir: {app}
Source: Install\Support\expsrv.dll; DestDir: {app}
Source: Install\Support\FLXGDIT.DLL; DestDir: {app}
Source: Install\Support\MSFLXGRD.OCX; DestDir: {app}
Source: Install\Support\MSJET35.DLL; DestDir: {app}
Source: Install\Support\MSJINT35.DLL; DestDir: {app}
Source: Install\Support\MSJTER35.DLL; DestDir: {app}
Source: Install\Support\MSRD2X35.DLL; DestDir: {app}
Source: Install\Support\MSREPL35.DLL; DestDir: {app}
Source: Install\Support\msvbvm60.dll; DestDir: {app}
Source: Install\Support\MSVCRT40.DLL; DestDir: {app}
Source: Install\Support\oleaut32.dll; DestDir: {app}
Source: Install\Support\olepro32.dll; DestDir: {app}
Source: Install\Support\ReNamer.DDF; DestDir: {app}
Source: Install\Support\ReNamer.exe; DestDir: {app}
Source: Install\Support\stdole2.tlb; DestDir: {app}
Source: Install\Support\TABCTIT.DLL; DestDir: {app}
Source: Install\Support\TABCTL32.OCX; DestDir: {app}
Source: Install\Support\VB5DB.DLL; DestDir: {app}
Source: Install\Support\VB6IT.DLL; DestDir: {app}
Source: Install\Support\VB6STKIT.DLL; DestDir: {app}
Source: Install\Support\vbajet32.dll; DestDir: {app}

[Icons]
Name: {group}\ReNamer 7.2; Filename: {app}\ReNamer.exe
Name: {group}\{cm:UninstallProgram,ReNamer}; Filename: {uninstallexe}
Name: {commondesktop}\ReNamer 7.2; Filename: {app}\ReNamer.exe; Tasks: desktopicon

[Run]
Filename: {app}\ReNamer.exe; Description: {cm:LaunchProgram,ReNamer}; Flags: nowait postinstall skipifsilent
