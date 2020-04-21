[Setup]
AppName=viewIT
AppVerName=viewIT v2.5
PrivilegesRequired=admin
DefaultDirName={pf}\viewIT
DefaultGroupName=viewIT
AppCopyright=© 2005-2009 JPElectron.com
AllowCancelDuringInstall=false
AllowUNCPath=false
ShowLanguageDialog=no
AppPublisher=JPElectron.com
AppPublisherURL=http://www.jpelectron.com
AppVersion=2.5
AppReadmeFile=http://www.jpelectron.com/readme/viewit.asp
UninstallDisplayIcon={app}\viewit.exe
VersionInfoVersion=2.5
VersionInfoCompany=JPElectron.com
VersionInfoDescription=viewIT Setup
VersionInfoCopyright=(c) 2005-2009 JPElectron.com
EnableDirDoesntExistWarning=false
DirExistsWarning=no
AppendDefaultGroupName=false
[Files]
Source: stdole2.tlb; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: msvbvm60.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: oleaut32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: olepro32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: asycfilt.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
Source: comcat.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: msinet.ocx; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: viewit.exe; DestDir: {app}; Flags: promptifolder
Source: viewit.ini; DestDir: {app}; Flags: promptifolder confirmoverwrite
Source: Readme.url; DestDir: {app}
Source: Readme.url; DestDir: {group}
[Icons]
Name: {group}\viewIT; Filename: {app}\viewit.exe; WorkingDir: {app}
[UninstallRun]
Filename: http://www.jpelectron.com/uf/frm.asp?type=new&app=viewit; Flags: shellexec
[Messages]
FinishedLabel=Setup has finished installing viewIT on your computer.
FinishedHeadingLabel=viewIT Setup Complete
