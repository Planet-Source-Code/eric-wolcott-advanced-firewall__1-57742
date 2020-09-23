Attribute VB_Name = "modReg"
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_CREATE_LINK As Long = &H20
Public Const SYNCHRONIZE As Long = &H100000
Public Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const ERROR_NO_MORE_ITEMS As Long = 259&
Public Const REG_SZ As Long = 1
Public pwdVar As String
Public sldLev As Integer
Public pwTemp As Integer
Public sldtemp As Integer
Public tempof As Integer
Public tempof2 As Integer
Public chkTemp As Integer
Public chkTemp2 As Integer

Public Function loadPWD()
    pwdVar = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\AdvFirewall\", "pwd")
End Function

Public Function getSecLevel() As Integer
    If frmFirewall.chkPw.Value = 0 Then
        getSecLevel = 0
    Else
        getSecLevel = frmFirewall.sld.Value
    End If
End Function

Public Function isAllowed(secLev As Integer) As Boolean
Dim temp As String
    If getSecLevel > secLev Then
        temp = InputBox("Enter Password", "Password Proctected Function")
        If pwdVar <> temp Then
            isAllowed = False
        Else
            chkTemp = 1
            isAllowed = True
        End If
    Else
        chkTemp = 0
        isAllowed = True
    End If
End Function

Public Function loadCheck() As Boolean
    If Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\AdvFirewall\", "load") = "yes" Then
        loadCheck = True
    Else
        loadCheck = False
    End If
End Function

Public Function loadPath() As String
    loadPath = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\AdvFirewall\", "path")
End Function
