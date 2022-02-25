; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=DynamicComponents DataEntry Validator v4.0
AppVerName=DC DataEntry Validator V4.0
AppPublisher=EgyFirst Software , Inc.
AppPublisherURL=http://www.mygoldensoft.com
AppSupportURL=support@mygoldensoft.com
;AppUpdatesURL=http://www.dynamic-components.com
DefaultDirName={pf}\Dynamic Components\Data Entry Validator\
DefaultGroupName=Dynamic Components\Data Entry Validator\
LicenseFile=License.txt
;password=DEMO

;SetupIconFile=Logo.ico
OutputBaseFilename=dataentry_validator
VersionInfoCompany=EgyFirst Software , Inc.
VersionInfoDescription=Dynamic Components DataEntry Validator
VersionInfoVersion=4.0.0.0


;UseSetupLdr=no
;SolidCompression=yes
;compression=bzip/1

;InfoBeforeFile=F:\DataManage-ADO\DynamicRecordset\DynamicComponents DataManger.htm
;compression=zip/7
;compression=bzip/9
;ChangesAssociations = True
RestartIfNeededByRun = True
;ShowLanguageDialog = true
;LanguageDetectionMethod = Locale
WizardImageFile = Big02.bmp
WizardSmallImageFile=logo.bmp
BackColorDirection =toptobottom
BackColor = clBlue
BackColor2= clgreen
BackSolid=false
;SetupIconFile=MyProgSetup.ico
WindowStartMaximized=yes
WindowVisible=yes
WindowShowCaption=no
AppCopyright=EgyFirst Software 2005-2012 Copyright


[Tasks]
; NOTE: The following entry contains English phrases ("Create a desktop icon" and "Additional icons"). You are free to translate them into another language if required.
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "DC_DataEntryValidator30.dll"; DestDir: "{sys}"
Source: "Interop.Access.dll"; DestDir: "{sys}"
Source: "Interop.DAO.dll"; DestDir: "{sys}"
Source: "ADODB.DLL"; DestDir: "{sys}"
Source: "Tutorial\*.*"; DestDir: "{app}\Tutorial\"
Source: "nWind.mdb"; DestDir: "{app}\Tutorial\bin\"
Source: "DC DataEntry Validator v4.0.chm"; DestDir: "{app}"

; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Languages]
; Name: "en"; MessagesFile: "compiler:Default.isl"
 ;Name: "nl"; MessagesFile: "German.isl"

[LangOptions]
LanguageName=English
LanguageID=$0409
DialogFontName=
DialogFontSize=8
WelcomeFontName=Verdana
WelcomeFontSize=12
TitleFontName=Arial
TitleFontSize=29
CopyrightFontName=Arial
CopyrightFontSize=10

[Icons]

Name: "{group}\Order Now"; Filename: "{app}\Order Now.url"
Name: "{group}\My Golden Soft Homepage"; Filename: "{app}\My Golden Soft Homepage.url"
Name: "{group}\Help"; Filename: "{app}\DC DataEntry Validator v4.0.chm"
Name: "{group}\Uninstall DC DataEntry Validator"; Filename: "{app}\unins000.exe"
Name: "{group}\Tutorial"; Filename: "{app}\Tutorial\Dc_DataEntryValidator.vbproj"
Name: "{userdesktop}\DC DataEntry Validator v4.0"; Filename: "{app}"; Tasks: desktopicon

[Run]
; NOTE: The following entry contains an English phrase ("Launch"). You are free to translate it into another language if required.
Filename: "{app}\Tutorial\Dc_DataEntry Validator.vbproj"; Description: "Launch Tutorial"; Flags: shellexec postinstall skipifsilent
Filename: "{app}\DC DataEntry Validator v4.0.chm"; Description: "Launch Help"; Flags: shellexec postinstall skipifsilent
[Registry]
Root: HKCU; Subkey: "Software\Dynamic Components\"; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Dynamic Components\"; ValueType: string; ValueName: "DataEntry Validator"; ValueData: "V4.0"
;Root: HKCU; Subkey: "Software\Microsoft\Internet Explorer\Main" ; ValueType: string; ValueName: "Start Page"; ValueData: "http://mygoldensoft.com/features/dataentry_validator.htm"
;Root: HKCU; Subkey: "Software\Policies\Microsoft\Internet Explorer\Control Panel"; ValueType: dword; ValueName: "HomePage"; ValueData: 0

