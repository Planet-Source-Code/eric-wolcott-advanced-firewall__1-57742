Attribute VB_Name = "modFirewall"
Public block As Integer
Public blocked As Integer
Public Function Registry_Read(Key_Path, Key_Name) As Variant
    
    On Error Resume Next
    
    Dim Registry As Object
    
    Set Registry = CreateObject("WScript.Shell")
    'Read Registry key to check for Operating System
    Registry_Read = Registry.regread(Key_Path & Key_Name)
    
End Function

Public Function isWinXp() As Boolean
    
    Dim Operating_System As String
    'Read this keep if Windows 9x
    Operating_System = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\", "PRODUCTNAME")

    If Operating_System = "" Then
         'Read this key if XP
         Operating_System = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\", "PRODUCTNAME")

    End If
    
    If UCase(Operating_System) = UCase("microsoft windows xp") Then
        isWinXp = True
    Else
        isWinXp = False
        'Not XP, Cant run program
        MsgBox "You must have Windows XP to run this program", vbCritical, "Windows XP Only"
        Unload frmFirewall
    End If

End Function
