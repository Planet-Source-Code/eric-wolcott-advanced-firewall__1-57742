Attribute VB_Name = "modProcPort"
'Const MIB_TCP_STATE_CLOSE_WAIT As Long = 8
'Const MIB_TCP_STATE_CLOSED As Long = 1
'Const MIB_TCP_STATE_CLOSING As Long = 9
'Const MIB_TCP_STATE_DELETE_TCB As Long = 12
'Const MIB_TCP_STATE_ESTAB As Long = 5
'Const MIB_TCP_STATE_FIN_WAIT1 As Long = 6
'Const MIB_TCP_STATE_FIN_WAIT2 As Long = 7
'Const MIB_TCP_STATE_LAST_ACK As Long = 10
'Const MIB_TCP_STATE_LISTEN As Long = 2
'Const MIB_TCP_STATE_SYN_RCVD As Long = 4
'Const MIB_TCP_STATE_SYN_SENT As Long = 3
'Const MIB_TCP_STATE_TIME_WAIT As Long = 11
Public colHead As ColumnHeader
Public lstItem As ListItem
Public refreshPort As Integer
Public tempPID As Long
Public tempName As String
Public foundName As Integer
Public checkforID As Integer
Public tempProcName As Long
Public procNum

Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Private Declare Function GetProcessHeap Lib "kernel32" () As Long


Private Declare Function htons Lib "ws2_32.dll" (ByVal dwLong As Long) As Long


Private Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi.dll" (pTcpTableEx As Any, ByVal bOrder As Long, ByVal heap As Long, ByVal zero As Long, ByVal flags As Long) As Long


Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (pTcpTableEx As MIB_TCPROW) As Long


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private pTablePtr As Long
    Private pDataRef As Long
    Private nRows As Long
    Private nCurrentRow As Long
    Private udtRow As MIB_TCPROW
    Private nState As Long
    Private nLocalAddr As Long
    Private nLocalPort As Long
    Private nRemoteAddr As Long
    Private nRemotePort As Long
    Private nProcId As Long


Public Function GetIPAddress(dwAddr As Long) As String
    Dim arrIpParts(3) As Byte
    CopyMemory arrIpParts(0), dwAddr, 4
    GetIPAddress = CStr(arrIpParts(0)) & "." & _
    CStr(arrIpParts(1)) & "." & _
    CStr(arrIpParts(2)) & "." & _
    CStr(arrIpParts(3))
End Function


Public Function GetPort(ByVal dwPort As Long) As Long
    GetPort = htons(dwPort)
End Function


Public Function RefreshStack() As Boolean
    Dim nRet As Long
    pDataRef = 0
    nRet = AllocateAndGetTcpExTableFromStack(pTablePtr, 0, GetProcessHeap, 0, 2)

    If nRet = 0 Then
        CopyMemory nRows, ByVal pTablePtr, 4
        RefreshStack = True
    Else
        RefreshStack = False
    End If
End Function


Public Function GetEntryCount() As Long
    GetEntryCount = nRows - 2 '// The last entry is always an EOF of sorts
End Function


Public Function EnumEntries() As Boolean
    procNum = 0
    EnumEntries = True
    If nRows = 0 Or pTablePtr = 0 Then
        EnumEntries = False
        Exit Function
    End If


    For i = 0 To nRows '// read 24 bytes at a time
        procNum = procNum + 1
        CopyMemory nState, ByVal pTablePtr + (pDataRef + 4), 4
        CopyMemory nLocalAddr, ByVal pTablePtr + (pDataRef + 8), 4
        CopyMemory nLocalPort, ByVal pTablePtr + (pDataRef + 12), 4
        CopyMemory nRemoteAddr, ByVal pTablePtr + (pDataRef + 16), 4
        CopyMemory nRemotePort, ByVal pTablePtr + (pDataRef + 20), 4
        CopyMemory nProcId, ByVal pTablePtr + (pDataRef + 24), 4
    
        DoEvents
        procNum = procNum + 1
        If nRemoteAddr <> 0 Or nRemotePort <> 0 Or nLocalPort <> 0 Then
            tempPID = nProcId
            foundName = 0
            Call enumProc
            If foundName = 0 Then
            
                tempName = "Unknown"
            End If
            foundName = 0
        End If
        If nProcId < 70000 And nProcId > 0 And nState > 0 And nState < 13 Then
            If notFirewall = 1 Then
                Set lstItem = frmFirewall.lstvwProc.ListItems.Add(, , tempName)
                lstItem.SubItems(1) = nProcId
                lstItem.SubItems(2) = GetIPAddress(nLocalAddr)
                lstItem.SubItems(3) = GetPort(nLocalPort)
                lstItem.SubItems(4) = GetIPAddress(nRemoteAddr)
                lstItem.SubItems(5) = GetPort(nRemotePort)
                lstItem.SubItems(6) = getState(nState)
            Else
                If block = 1 Then
                    For t = 0 To frmFirewall.lstBlock.ListCount - 1
                        If UCase(frmFirewall.lstBlock.List(t)) = UCase(tempName) Then
                            foundmain = 1
                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                            blocked = blocked + 1
                            Call updateLog
                            Exit For
                        Else
                            foundmain = 0
                        End If
                    Next t
                    For t = 0 To frmFirewall.lstAllow.ListCount - 1
                        If UCase(frmFirewall.lstAllow.List(t)) = UCase(tempName) Then
                            foundmaina = 1
                            Exit For
                        Else
                            foundmaina = 0
                        End If
                    Next t
                    If foundmain <> 1 And foundmaina <> 1 Then
                        For z = 0 To frmFirewall.lstRules(0).ListCount - 1
                            For t = 0 To frmFirewall.lstAllow.ListCount - 1
                                If UCase(frmFirewall.lstAllow.List(t)) = UCase(tempName) Then
                                    foundHA = 1
                                    Exit For
                                Else
                                    foundHA = 0
                                End If
                            Next t
                            For t = 0 To frmFirewall.lstBlock.ListCount - 1
                                If UCase(frmFirewall.lstBlock.List(t)) = UCase(tempName) Then
                                    foundBL = 1
                                    Exit For
                                Else
                                    foundBL = 0
                                End If
                            Next t
                            If foundBL <> 1 Then
                                If foundHA <> 1 Then
                                    If UCase(tempName) = UCase(frmFirewall.lstRules(0).List(z)) Then
                                        If promptBlock = 1 Then
                                            SuspendThreads nProcId
                                            If frmFirewall.chkDetail.Value = 1 Then
                                                getResponse tempName, nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                                Do While holdLoop = 0
                                                    DoEvents
                                                Loop
                                                If blockAlert = 1 Then
                                                    ResumeThreads nProcId
                                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                                    blocked = blocked + 1
                                                    Call updateLog
                                                Else
                                                    ResumeThreads nProcId
                                                    Call updateAccess
                                                End If
                                            Else
                                                getrespond = MsgBox("Would you like you block " & tempName & " from connection to the internet?", vbYesNo, "Block Process")
                                                If getrespond = vbYes Then
                                                    getrespond = MsgBox("Add to block list?", vbYesNo, "Block List")
                                                    If getrespond = vbYes Then
                                                        frmFirewall.lstBlock.AddItem tempName
                                                        DoEvents
                                                    End If
                                                    ResumeThreads nProcId
                                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                                    blocked = blocked + 1
                                                    Call updateLog
                                                ElseIf getrespond = vbNo Then
                                                    ResumeThreads nProcId
                                                    getrespond = MsgBox("Add to allow list?", vbYesNo, "Allow List")
                                                    If getrespond = vbYes Then
                                                        frmFirewall.lstAllow.AddItem tempName
                                                        DoEvents
                                                    End If
                                                    Call updateAccess
                                                End If
                                            End If
                                        Else
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        End If
                                    End If
                                End If
                            ElseIf foundBL = 1 Then
                                TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                blocked = blocked + 1
                                Call updateLog
                            End If
                        Next z
                        For z = 0 To frmFirewall.lstRules(1).ListCount - 1
                            If nRemoteAddr = frmFirewall.lstRules(1).List(z) Then
                                If promptBlock = 1 Then
                                    SuspendThreads nProcId
                                     If frmFirewall.chkDetail.Value = 1 Then
                                        getResponse tempName, nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        Do While holdLoop = 0
                                            DoEvents
                                        Loop
                                        If blockAlert = 1 Then
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        Else
                                            ResumeThreads nProcId
                                            Call updateAccess
                                        End If
                                    Else
                                        getrespond = MsgBox("Would you like you block " & tempName & " from connection to the internet?", vbYesNo, "Block Process")
                                        If getrespond = vbYes Then
                                            getrespond = MsgBox("Add to block list?", vbYesNo, "Block List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstBlock.AddItem tempName
                                            End If
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        ElseIf getrespond = vbNo Then
                                            ResumeThreads nProcId
                                            getrespond = MsgBox("Add to allow list?", vbYesNo, "Allow List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstAllow.AddItem tempName
                                            End If
                                            Call updateAccess
                                        End If
                                    End If
                                Else
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    blocked = blocked + 1
                                    Call updateLog
                                End If
                            End If
                        Next z
                        For z = 0 To frmFirewall.lstRules(2).ListCount - 1
                            If nRemotePort = frmFirewall.lstRules(2).List(z) Then
                                If promptBlock = 1 Then
                                    SuspendThreads nProcId
                                     If frmFirewall.chkDetail.Value = 1 Then
                                        getResponse tempName, nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        Do While holdLoop = 0
                                            DoEvents
                                        Loop
                                        If blockAlert = 1 Then
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        Else
                                            ResumeThreads nProcId
                                            Call updateAccess
                                        End If
                                    Else
                                        getrespond = MsgBox("Would you like you block " & tempName & " from connection to the internet?", vbYesNo, "Block Process")
                                        If getrespond = vbYes Then
                                            getrespond = MsgBox("Add to block list?", vbYesNo, "Block List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstBlock.AddItem tempName
                                            End If
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        ElseIf getrespond = vbNo Then
                                            ResumeThreads nProcId
                                            getrespond = MsgBox("Add to allow list?", vbYesNo, "Allow List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstAllow.AddItem tempName
                                            End If
                                            Call updateAccess
                                        End If
                                    End If
                                Else
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    blocked = blocked + 1
                                    Call updateLog
                                End If
                            End If
                        Next z
                        For z = 0 To frmFirewall.lstRules(3).ListCount - 1
                            If nLocalPort = frmFirewall.lstRules(3).List(z) Then
                               If promptBlock = 1 Then
                                    SuspendThreads nProcId
                                     If frmFirewall.chkDetail.Value = 1 Then
                                        getResponse tempName, nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        Do While holdLoop = 0
                                            DoEvents
                                        Loop
                                        If blockAlert = 1 Then
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        Else
                                            ResumeThreads nProcId
                                            Call updateAccess
                                        End If
                                    Else
                                        getrespond = MsgBox("Would you like you block " & tempName & " from connection to the internet?", vbYesNo, "Block Process")
                                        If getrespond = vbYes Then
                                            getrespond = MsgBox("Add to block list?", vbYesNo, "Block List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstBlock.AddItem tempName
                                            End If
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        ElseIf getrespond = vbNo Then
                                            ResumeThreads nProcId
                                            getrespond = MsgBox("Add to allow list?", vbYesNo, "Allow List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstAllow.AddItem tempName
                                            End If
                                            Call updateAccess
                                        End If
                                    End If
                                Else
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    blocked = blocked + 1
                                    Call updateLog
                                End If
                            End If
                        Next z
                    ElseIf block = 0 Then
                        For z = 0 To frmFirewall.lstRules(0).ListCount - 1
                            If UCase(tempName) <> UCase(frmFirewall.lstRules(0).List(z)) Then
                                If promptBlock = 1 Then
                                    SuspendThreads nProcId
                                     If frmFirewall.chkDetail.Value = 1 Then
                                        getResponse tempName, nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        Do While holdLoop = 0
                                            DoEvents
                                        Loop
                                        If blockAlert = 1 Then
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        Else
                                            ResumeThreads nProcId
                                            Call updateAccess
                                        End If
                                    Else
                                        getrespond = MsgBox("Would you like you block " & tempName & " from connection to the internet?", vbYesNo, "Block Process")
                                        If getrespond = vbYes Then
                                            getrespond = MsgBox("Add to block list?", vbYesNo, "Block List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstBlock.AddItem tempName
                                            End If
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        ElseIf getrespond = vbNo Then
                                            ResumeThreads nProcId
                                            getrespond = MsgBox("Add to allow list?", vbYesNo, "Allow List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstAllow.AddItem tempName
                                            End If
                                            Call updateAccess
                                        End If
                                    End If
                                Else
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    blocked = blocked + 1
                                    Call updateLog
                                End If
                            End If
                        Next z
                        For z = 0 To frmFirewall.lstRules(1).ListCount - 1
                            If nRemoteAddr <> frmFirewall.lstRules(1).List(z) Then
                                If promptBlock = 1 Then
                                    SuspendThreads nProcId
                                     If frmFirewall.chkDetail.Value = 1 Then
                                        getResponse tempName, nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        Do While holdLoop = 0
                                            DoEvents
                                        Loop
                                        If blockAlert = 1 Then
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        Else
                                            ResumeThreads nProcId
                                            Call updateAccess
                                        End If
                                    Else
                                        getrespond = MsgBox("Would you like you block " & tempName & " from connection to the internet?", vbYesNo, "Block Process")
                                        If getrespond = vbYes Then
                                            getrespond = MsgBox("Add to block list?", vbYesNo, "Block List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstBlock.AddItem tempName
                                            End If
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        ElseIf getrespond = vbNo Then
                                            ResumeThreads nProcId
                                            getrespond = MsgBox("Add to allow list?", vbYesNo, "Allow List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstAllow.AddItem tempName
                                            End If
                                            Call updateAccess
                                        End If
                                    End If
                                Else
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    blocked = blocked + 1
                                    Call updateLog
                                End If
                            End If
                        Next z
                        For z = 0 To frmFirewall.lstRules(2).ListCount - 1
                            If nRemotePort <> frmFirewall.lstRules(3).List(z) Then
                                If promptBlock = 1 Then
                                    SuspendThreads nProcId
                                     If frmFirewall.chkDetail.Value = 1 Then
                                        getResponse tempName, nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        Do While holdLoop = 0
                                            DoEvents
                                        Loop
                                        If blockAlert = 1 Then
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        Else
                                            ResumeThreads nProcId
                                            Call updateAccess
                                        End If
                                    Else
                                        getrespond = MsgBox("Would you like you block " & tempName & " from connection to the internet?", vbYesNo, "Block Process")
                                        If getrespond = vbYes Then
                                            getrespond = MsgBox("Add to block list?", vbYesNo, "Block List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstBlock.AddItem tempName
                                            End If
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        ElseIf getrespond = vbNo Then
                                            ResumeThreads nProcId
                                            getrespond = MsgBox("Add to allow list?", vbYesNo, "Allow List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstAllow.AddItem tempName
                                            End If
                                            Call updateAccess
                                        End If
                                    End If
                                Else
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    blocked = blocked + 1
                                    Call updateLog
                                End If
                            End If
                        Next z
                        For z = 0 To frmFirewall.lstRules(3).ListCount - 1
                            If nLocalPort <> frmFirewall.lstRules(3).List(z) Then
                                If promptBlock = 1 Then
                                    SuspendThreads nProcId
                                     If frmFirewall.chkDetail.Value = 1 Then
                                        getResponse tempName, nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        Do While holdLoop = 0
                                            DoEvents
                                        Loop
                                        If blockAlert = 1 Then
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        Else
                                            ResumeThreads nProcId
                                            Call updateAccess
                                        End If
                                    Else
                                        getrespond = MsgBox("Would you like you block " & tempName & " from connection to the internet?", vbYesNo, "Block Process")
                                        If getrespond = vbYes Then
                                            getrespond = MsgBox("Add to block list?", vbYesNo, "Block List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstBlock.AddItem tempName
                                            End If
                                            ResumeThreads nProcId
                                            TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                            blocked = blocked + 1
                                            Call updateLog
                                        ElseIf getrespond = vbNo Then
                                            ResumeThreads nProcId
                                            getrespond = MsgBox("Add to allow list?", vbYesNo, "Allow List")
                                            If getrespond = vbYes Then
                                                frmFirewall.lstAllow.AddItem tempName
                                            End If
                                            Call updateAccess
                                        End If
                                    End If
                                Else
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    blocked = blocked + 1
                                    Call updateLog
                                End If
                            End If
                        Next z
                    End If
                End If
            End If
        End If
        pDataRef = pDataRef + 24
        DoEvents
    Next i
    Call updateProcNum
End Function

Public Sub TerminateThisConnection(xLocalAddr As Long, xLocalPort As Long, xRemoteAddr As Long, xRemotePort As Long)
    udtRow.dwLocalAddr = xLocalAddr
    udtRow.dwLocalPort = xLocalPort
    udtRow.dwRemoteAddr = xRemoteAddr
    udtRow.dwRemotePort = xRemotePort
    udtRow.dwState = 12
    SetTcpEntry udtRow
End Sub

Public Function getState(stateOf As Long) As String

    Select Case stateOf
    
        Case 1
            getState = "Closed"
        Case 2
            getState = "Listening"
        Case 3
            getState = "SYN Sent"
        Case 4
            getState = "SYN Recieved"
        Case 5
            getState = "Established"
        Case 6
            getState = "FIN Wait 1"
        Case 7
            getState = "FIN Wait 2"
        Case 8
            getState = "Close Wait"
        Case 9
            getState = "Closing"
        Case 10
            getState = "Last ACK"
        Case 11
            getState = "Time Wait"
        Case 12
            getState = "Delete TCB"
    End Select

End Function

Public Sub updateLog()
    If frmFirewall.lstvwBlock.ListItems.Count = 30 Then
        frmFirewall.lstvwBlock.ListItems.Remove 1
    End If
    
    If frmFirewall.lstvwBlock.ListItems.Count <> 0 Then
        If UCase(frmFirewall.lstvwBlock.ListItems(frmFirewall.lstvwBlock.ListItems.Count).Text) = UCase(tempName) And frmFirewall.lstvwBlock.ListItems(frmFirewall.lstvwBlock.ListItems.Count).ListSubItems(1).Text = nProcId And frmFirewall.lstvwBlock.ListItems(frmFirewall.lstvwBlock.ListItems.Count).ListSubItems(2).Text = GetIPAddress(nLocalAddr) And frmFirewall.lstvwBlock.ListItems(frmFirewall.lstvwBlock.ListItems.Count).ListSubItems(3).Text = GetPort(nLocalPort) And frmFirewall.lstvwBlock.ListItems(frmFirewall.lstvwBlock.ListItems.Count).ListSubItems(4).Text = GetIPAddress(nRemoteAddr) And frmFirewall.lstvwBlock.ListItems(frmFirewall.lstvwBlock.ListItems.Count).ListSubItems(5).Text = GetPort(nRemotePort) Then
            frmFirewall.lstvwBlock.ListItems(frmFirewall.lstvwBlock.ListItems.Count).ListSubItems(6) = frmFirewall.lstvwBlock.ListItems(frmFirewall.lstvwBlock.ListItems.Count).ListSubItems(6) + 1
            frmFirewall.lstvwBlock.ListItems(frmFirewall.lstvwBlock.ListItems.Count).ListSubItems(7) = Date
        Else
            Set lstItem = frmFirewall.lstvwBlock.ListItems.Add(, , tempName)
            lstItem.SubItems(1) = nProcId
            lstItem.SubItems(2) = GetIPAddress(nLocalAddr)
            lstItem.SubItems(3) = GetPort(nLocalPort)
            lstItem.SubItems(4) = GetIPAddress(nRemoteAddr)
            lstItem.SubItems(5) = GetPort(nRemotePort)
            lstItem.SubItems(6) = 0
            lstItem.SubItems(7) = Date & Now
        End If
     Else
         Set lstItem = frmFirewall.lstvwBlock.ListItems.Add(, , tempName)
        lstItem.SubItems(1) = nProcId
        lstItem.SubItems(2) = GetIPAddress(nLocalAddr)
        lstItem.SubItems(3) = GetPort(nLocalPort)
        lstItem.SubItems(4) = GetIPAddress(nRemoteAddr)
        lstItem.SubItems(5) = GetPort(nRemotePort)
        lstItem.SubItems(6) = 0
        lstItem.SubItems(7) = Now
     End If
End Sub

Public Sub updateAccess()
    If frmFirewall.lstvwAccess.ListItems.Count = 30 Then
        frmFirewall.lstvwAccess.ListItems.Remove 1
    End If
    
    If frmFirewall.lstvwAccess.ListItems.Count <> 0 Then
        If UCase(frmFirewall.lstvwAccess.ListItems(frmFirewall.lstvwAccess.ListItems.Count).Text) = UCase(tempName) And frmFirewall.lstvwAccess.ListItems(frmFirewall.lstvwAccess.ListItems.Count).ListSubItems(1).Text = nProcId And frmFirewall.lstvwAccess.ListItems(frmFirewall.lstvwAccess.ListItems.Count).ListSubItems(2).Text = GetIPAddress(nLocalAddr) And frmFirewall.lstvwAccess.ListItems(frmFirewall.lstvwAccess.ListItems.Count).ListSubItems(3).Text = GetPort(nLocalPort) And frmFirewall.lstvwAccess.ListItems(frmFirewall.lstvwAccess.ListItems.Count).ListSubItems(4).Text = GetIPAddress(nRemoteAddr) And frmFirewall.lstvwAccess.ListItems(frmFirewall.lstvwAccess.ListItems.Count).ListSubItems(5).Text = GetPort(nRemotePort) Then
            frmFirewall.lstvwAccess.ListItems(frmFirewall.lstvwAccess.ListItems.Count).ListSubItems(6) = frmFirewall.lstvwAccess.ListItems(frmFirewall.lstvwAccess.ListItems.Count).ListSubItems(6) + 1
            frmFirewall.lstvwAccess.ListItems(frmFirewall.lstvwAccess.ListItems.Count).ListSubItems(7) = Date
        Else
            Set lstItem = frmFirewall.lstvwAccess.ListItems.Add(, , tempName)
            lstItem.SubItems(1) = nProcId
            lstItem.SubItems(2) = GetIPAddress(nLocalAddr)
            lstItem.SubItems(3) = GetPort(nLocalPort)
            lstItem.SubItems(4) = GetIPAddress(nRemoteAddr)
            lstItem.SubItems(5) = GetPort(nRemotePort)
            lstItem.SubItems(6) = 0
            lstItem.SubItems(7) = Date & Now
        End If
     Else
         Set lstItem = frmFirewall.lstvwAccess.ListItems.Add(, , tempName)
        lstItem.SubItems(1) = nProcId
        lstItem.SubItems(2) = GetIPAddress(nLocalAddr)
        lstItem.SubItems(3) = GetPort(nLocalPort)
        lstItem.SubItems(4) = GetIPAddress(nRemoteAddr)
        lstItem.SubItems(5) = GetPort(nRemotePort)
        lstItem.SubItems(6) = 0
        lstItem.SubItems(7) = Date & Now
     End If
End Sub

Public Sub updateProcNum()
    frmFirewall.lblProcNum.Caption = procNum
    frmFirewall.lblBlocked.Caption = blocked
    frmFirewall.lblLife.Caption = blockedLife + blocked
End Sub

Public Sub getResponse(xProcName As String, xLocalAddr As Long, xLocalPort As Long, xRemoteAddr As Long, xRemotePort As Long)
    blockAlert = 0
    holdLoop = 0
    Load frmAlert
    frmAlert.Show
    frmAlert.lblFname = xProcName
    frmAlert.lblLIP = GetIPAddress(xLocalAddr)
    frmAlert.lblRIP = GetIPAddress(xRemoteAddr)
    frmAlert.lblLport = GetPort(xLocalPort)
    frmAlert.lblRport = GetPort(xRemotePort)
End Sub
