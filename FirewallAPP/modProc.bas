Attribute VB_Name = "modProc"
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Const PROCESS_TERMINATE As Long = (&H1)
Public Const MAX_PATH As Integer = 260
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public infoProcInfo As PROCESSENTRY32

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Sub enumProc()
Dim procType As String
    procType = ""
    servProc = 0
    uknProc = 0
    sysProc = 0
    tempName = ""
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    r = Process32Next(hSnapShot, uProcess)
    Do While r
        processname = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
        If UCase(processname) = UCase("services.exe") Then
                servPID = uProcess.th32ProcessID
            ElseIf UCase(processname) = UCase("explorer.exe") Then
                expPID = uProcess.th32ProcessID
            ElseIf UCase(processname) = UCase("system") Then
                sysPID(1) = uProcess.th32ProcessID
            ElseIf UCase(processname) = UCase("smss.exe") Then
                sysPID(2) = uProcess.th32ProcessID
            ElseIf UCase(processname) = UCase("winlogon.exe") Then
                sysPID(3) = uProcess.th32ProcessID
            ElseIf UCase(processname) = UCase("csrss.exe") Then
                sysPID(4) = uProcess.th32ProcessID
            ElseIf UCase(processname) = UCase("lsass.exe") Then
                sysPID(5) = uProcess.th32ProcessID
        End If
        If tempPID = uProcess.th32ProcessID Then
            tempName = processname
            foundName = 1
        End If
        r = Process32Next(hSnapShot, uProcess)
    Loop
    CloseHandle hSnapShot
End Sub

