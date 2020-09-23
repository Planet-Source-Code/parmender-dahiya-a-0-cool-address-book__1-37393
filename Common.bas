Attribute VB_Name = "Common"
Option Explicit

Public OpenError As Boolean
Public code As String
Public db As Database
Public rs As Recordset
Public ws As Workspace
Public max As Long
Public i As Long
Public filests As String
Public GetErrorMsg As String

Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'API for launching the HTML help file
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
    'Constants used by HtmlHelp
    Const HH_DISPLAY_TOPIC = &H0
    Const HH_SET_WIN_TYPE = &H4
    Const HH_GET_WIN_TYPE = &H5
    Const HH_GET_WIN_HANDLE = &H6
    Const HH_DISPLAY_TEXT_POPUP = &HE 'Display String resource ID or text in a pop-up window.
    Const HH_HELP_CONTEXT = &HF 'Display mapped numeric value In dwData.
    Const HH_TP_HELP_CONTEXTMENU = &H10 'Text pop-up help, similar To WinHelp's HELP_CONTEXTMENU.
    Const HH_TP_HELP_WM_HELP = &H11 'Text pop-up help, similar To WinHelp 's HELP_WM_HELP.
'API for puuting on top of all other applications


Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

'Const REG_NONE = 0&
Const REG_SZ = 1&
'Const REG_EXPAND_SZ = 2&
'Const REG_BINARY = 3&
'Const REG_DWORD = 4&
'Const REG_DWORD_LITTLE_ENDIAN = 4&
'Const REG_DWORD_BIG_ENDIAN = 5&
'Const REG_LINK = 6&
'Const REG_MULTI_SZ = 7&
'Const REG_RESOURCE_LIST = 8&
'Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
'Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
'Const KEY_EXECUTE = KEY_READ

Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte
Const DisplayErrorMsg = True

Public Function initdb()
On Error GoTo error
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\Addresses.mdb")
Set rs = db.OpenRecordset("Contacts", dbOpenTable)
rs.Index = "First"
max = rs.RecordCount
Exit Function
error:
MsgBox "Either database with name 'Addresses' and with table name 'Contacts' not found in the current folder." _
& vbCrLf & "Or there is some unexpected error while opening the database.", _
vbCritical + vbOKOnly, "Address Book - Database Error"
MsgBox "Address Book can not be started.", vbCritical + vbOKOnly, "Address Book - Opening Error"
End
End Function

Function SetStringValue(SubKey As String, Entry As String, Value As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            ErrorMsg (rtn)       'display the error
            MsgBox GetErrorMsg, vbCritical + vbOKOnly, "Address Book - Error"
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         ErrorMsg (rtn)              'display the error
         MsgBox GetErrorMsg, vbCritical + vbOKOnly, "Address Book - Error"
      End If
   End If
End If

End Function

Private Sub ParseKey(KeyName As String, Keyhandle As Long)
    
rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname

If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName 'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(KeyName)
   KeyName = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1)) 'seperate the Keyname
   KeyName = Right(KeyName, Len(KeyName) - rtn)
End If

End Sub
Function ErrorMsg(lErrorCode As Long) As String
    
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
       Case 2, 1010
            GetErrorMsg = "Bad Key Name"
       Case 1011
            GetErrorMsg = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg = "Can't Read Key"
       Case 5
            GetErrorMsg = "Access is denied"
       Case 1013
            GetErrorMsg = "Can't Write Key"
       Case 8, 14
            GetErrorMsg = "Out of memory"
       Case 87
            GetErrorMsg = "Invalid Parameter"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select

End Function
Function GetMainKeyHandle(MainKeyName As String) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function


