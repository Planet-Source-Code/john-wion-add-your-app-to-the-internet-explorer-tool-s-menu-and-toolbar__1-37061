Attribute VB_Name = "IETOOLS"
Option Explicit
' Shlwapi.dll (MSIE Version Info) (All required...)
Type DllVersionInfo
cbSize As Long
dwMajorVersion As Long '...But the only one we need
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type

Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DllVersionInfo) As Long

Dim IEMV As DllVersionInfo
Dim CheckReg As String
Dim GetIEMajor As String
Dim Hico As String
Dim Ico As String
Dim Prog As String

Public Function DetectIE()
'See Remarks in Private Sub Form_Load()
CheckReg = REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE", "")
IEMV.cbSize = Len(IEMV)
Call DllGetVersion(IEMV)
GetIEMajor = IEMV.dwMajorVersion
If Dir(CheckReg) = "" Or GetIEMajor < 5 Then
Form1.mnuAddMSIE.Visible = False
Form1.mnuDeleteMSIE.Visible = False
Form1.mnuSep1.Visible = False
Else
CheckReg = REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "CLSID")
If CheckReg = "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}" Then
Form1.mnuAddMSIE.Enabled = False
Form1.mnuDeleteMSIE.Enabled = True
Else
Form1.mnuAddMSIE.Enabled = True
Form1.mnuDeleteMSIE.Enabled = False
End If
End If
End Function

Public Function mnuAddIE()
' Path of yor App and HotIcon
Hico = App.Path & "\" & "hoticon.ico"
' Path of yor App and Icon
Ico = App.Path & "\" & "icon.ico"
' Path of yor App and Apps *.exe name
Prog = App.Path & "\" & App.EXEName
' Adds your App to MSIE's Tools Menu and adds an Icon on the Toolbar
' {ECC5777A-6E88-BFCE-13CE-81F134789E7B} Any GUID
' Your App (Tools Menu Button Text)
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "ButtonText", "Your App"
' {1FBA04EE-3024-11D2-8F1F-0000F87ABD16} MUST BE THIS GUID
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}"
' Show Icon if IE Toolbar is  reset
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "Default Visible", "Yes"
' Your APP Path and Name (Prog)
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "Exec", Prog
' Colered icon (Hico)
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "HotIcon", Hico
' Default icon (Ico)
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "Icon", Ico
'Statusbar text for your App
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "MenuStatusBar", "Your App"
'Tools Menu text for your APP
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "MenuText", "&Your App"
Form1.mnuAddMSIE.Enabled = False
Form1.mnuDeleteMSIE.Enabled = True
End Function

Public Function mnuDeleteIE()
' Deletes your App in MSIE's Tools Menu and the Icon on the Toolbar
REGDeleteSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}"
Form1.mnuDeleteMSIE.Enabled = False
Form1.mnuAddMSIE.Enabled = True
End Function
