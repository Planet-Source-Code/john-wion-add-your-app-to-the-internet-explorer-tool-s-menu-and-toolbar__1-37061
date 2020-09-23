Attribute VB_Name = "CRegister"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKEY As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal HKEY As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
'Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKEY As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long

Enum HKEYS
vHKEY_CLASSES_ROOT = &H80000000
vHKEY_CURRENT_USER = &H80000001
vHKEY_LOCAL_MACHINE = &H80000002
vHKEY_USERS = &H80000003
vHKEY_PERFORMcANCE_DATA = &H80000004
vHKEY_CURRENT_CONFIG = &H80000005
vHKEY_DYN_DATA = &H80000006
End Enum

Private Const HKEY_CURRENT_USER         As Long = &H80000001
Private Const REG_OPTION_NON_VOLATILE   As Long = 0    ' Key is preserved when system is rebooted
Private Const SYNCHRONIZE               As Long = &H100000
Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const KEY_ALL_ACCESS            As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS             As Long = 0&
Private Const REG_SZ                    As Long = 1                   ' Unicode nul terminated string
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
'KEY_ALL_ACCESS 'Combination of KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, KEY_CREATE_SUB_KEY, KEY_CREATE_LINK, and KEY_SET_VALUE access.
'KEY_CREATE_LINK 'Permission to create a symbolic link.
'KEY_CREATE_SUB_KEY 'Permission to create subkeys.
'KEY_ENUMERATE_SUB_KEYS 'Permission to enumerate subkeys.
'KEY_EXECUTE 'Permission for read access.
'KEY_NOTIFY 'Permission for change notification.
'KEY_QUERY_VALUE 'Permission to query subkey data.
'KEY_READ 'Combination of KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, and KEY_NOTIFY access.
'KEY_SET_VALUE 'Permission to set subkey data.
'KEY_WRITE
'********************************************
'        Begin Registry Function
'********************************************
'============================================
' REGDeleteSetting
' Delete Section/Key from Registry
'--------------------------------------------
'
'  REGDeleteSetting vHKEY_USERS,"Section"
'    Deletes "HKEY_USER\Section\"
'    from the registry and all Key and Values under the section
'
'  REGDeleteSetting vHKEY_USERS,"Section1\Section2"
'    idem but deletes all from "HKEY_USERS\Section1\Section2"
'
'  REGDeleteSetting vHKEY_USERS,"Section",Key"
'    Deletes "HKEY_USER\Section\Key"
'    from the registry and Values under the key
'============================================

Public Function REGDeleteSetting(ByVal regHKEY As HKEYS, ByVal sSection As String, Optional ByVal sKey As String) As Boolean
Dim lReturn As Long
Dim HKEY As Long
If Len(sKey) Then
lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_ALL_ACCESS, HKEY)
If lReturn = ERROR_SUCCESS Then
If sKey = "*" Then sKey = vbNullString
lReturn = RegDeleteValue(HKEY, sKey)
End If
Else
lReturn = RegOpenKeyEx(regHKEY, REGSubKey(), 0&, KEY_ALL_ACCESS, HKEY)
If lReturn = ERROR_SUCCESS Then
lReturn = RegDeleteKey(HKEY, sSection)
End If
End If
REGDeleteSetting = (lReturn = ERROR_SUCCESS)
End Function
'============================================
' REGGetSetting
' Gets Values from the registry
'--------------------------------------------
'
' REGGetSetting vHKEY_CURRENT_USER,"Section","Key","DefaultStringWhenEmpty"
'   Gets Value from "HKEY_CURRENT_USER\Section\Key"
'   When empty it returns the omitted default("DefaultStringWhenEmpty")
'   or an empty string when no default is specified
'
' REGGetSetting vHKEY_CURRENT_USER,"Section1\Section2","Key","DefaultStringWhenEmpty"
'   idem but gets value from "HKEY_CURRENT_USER\Section1\Section2\Key"
'============================================

Public Function REGGetSetting(ByVal regHKEY As HKEYS, ByVal sSection As String, ByVal sKey As String, Optional ByVal sDefault As String) As String
Dim lReturn As Long
Dim HKEY As Long
Dim lType As Long
Dim lBytes As Long
Dim sBuffer As String
REGGetSetting = sDefault
'Original
lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_ALL_ACCESS, HKEY)
If lReturn = 5 Then  'We hebben geen rechten om hem te openen met KEY_ALL_ACCESS, dus we gaan hem read only openen
lReturn = RegOpenKeyEx(regHKEY, REGSubKey(sSection), 0&, KEY_EXECUTE, HKEY)
End If
If lReturn = ERROR_SUCCESS Then
If sKey = "*" Then
sKey = vbNullString
End If
lReturn = RegQueryValueEx(HKEY, sKey, 0&, lType, ByVal sBuffer, lBytes)
If lReturn = ERROR_SUCCESS Then
If lBytes > 0 Then
sBuffer = Space$(lBytes)
lReturn = RegQueryValueEx(HKEY, sKey, 0&, lType, ByVal sBuffer, Len(sBuffer))
If lReturn = ERROR_SUCCESS Then
REGGetSetting = Left$(sBuffer, lBytes - 1)
End If
End If
End If
End If
End Function
'============================================
' REGSaveSetting
' Save Value to the registry
'--------------------------------------------
'
' REGSaveSetting vHKEY_CURRENT_USER, "Section", "Key", "Test"
'   Saves the value "Test" to "HKEY_CURRENT_USER\Section\Key"
'   And will create the The Sections if they do not exist
'
' REGSaveSetting vHKEY_CURRENT_USER, "Section1\Section2", "Key", "Test"
'   idem but save "to HKEY_CURRENT_USER\Section1\Section2\Key"
'============================================

Public Function REGSaveSetting(ByVal regHKEY As HKEYS, ByVal sSection As String, ByVal sKey As String, ByVal sValue As String) As Boolean
Dim lRet As Long
Dim HKEY As Long
Dim lResult As Long
lRet = RegCreateKeyEx(regHKEY, REGSubKey(sSection), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, HKEY, lResult)
If lRet = ERROR_SUCCESS Then
If sKey = "*" Then sKey = vbNullString
lRet = RegSetValueEx(HKEY, sKey, 0&, REG_SZ, ByVal sValue, Len(sValue))
Call RegCloseKey(HKEY)
End If
REGSaveSetting = (lRet = ERROR_SUCCESS)
End Function
'Deletes "\" after section if there

Private Function REGSubKey(Optional ByVal sSection As String) As String
If Left$(sSection, 1) = "\" Then
sSection = Mid$(sSection, 2)
End If
If Right$(sSection, 1) = "\" Then
sSection = Mid$(sSection, 1, Len(sSection) - 1)
End If
REGSubKey = sSection
End Function
'********************************************
'          End Registry Function
'********************************************
