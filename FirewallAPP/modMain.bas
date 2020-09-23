Attribute VB_Name = "modMain"
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nameRules As Integer
Public ipRules As Integer
Public rportRules As Integer
Public lportRules As Integer
Public sysPID(1 To 5) As Integer
Public expPID As Integer
Public servPID As Integer
Public notFirewall As Integer
Public promptBlock As Integer
Public noShow As Integer
Public blockedLife As Integer
Public blockAlert As Integer
Public holdLoop As Integer
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205
Public TrayI As NOTIFYICONDATA
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Sub writeAccess()
On Error Resume Next
Dim fcgName As String
    fcgName = frmFirewall.txtSaveAccess.Text
    If fcgName <> "" Then
        Open fcgName For Append As #1
        Close #1
        Open fcgName For Output As #1
        Print #1, "[Access Log. For Use With Advanced Firewall]"
        Print #1, "[ProcessName][ProcessID][LocalAddress][LocalPort][RemoteAddress][RemotePort][Attempts][Time]"
        For i = 1 To lstvwAccess.ListItems.Count - 1
            Print #1, lstvwAccess.ListItems(i).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(1).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(2).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(3).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(4).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(5).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(6).Text & ";" & lstvwAccess.ListItems(i).ListSubItems(7).Text & ";" & Date & " " & Time
        Next i
        Close #1
    End If
End Sub

Public Sub writeBlock()
On Error Resume Next
Dim fcgName As String
    fcgName = frmFirewall.txtSaveBlock.Text
    If fcgName <> "" Then
        Open fcgName For Append As #1
        Close #1
        Open fcgName For Output As #1
        Print #1, "[Block Log. For Use With Advanced Firewall]"
        Print #1, "[ProcessName][ProcessID][LocalAddress][LocalPort][RemoteAddress][RemotePort][Attempts][Time]"
        For i = 1 To lstvwBlock.ListItems.Count
            Print #1, lstvwBlock.ListItems(i).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(1).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(2).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(3).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(4).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(5).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(6).Text & ";" & lstvwBlock.ListItems(i).ListSubItems(7).Text & ";" & Date & " " & Time
        Next i
        Close #1
    End If
End Sub
