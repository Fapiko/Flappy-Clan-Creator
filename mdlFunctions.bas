Attribute VB_Name = "mdlFunctions"
Option Explicit
Public Sub AddC(Index As Byte, ParamArray saElements() As Variant)
    Dim Data As String
    Dim newText() As String
    Dim strTimeStamp As String
    
    strTimeStamp = "[" & Format(Time, "c") & "] "
    With frmMain.rtbStatus(Index)
        Do Until UBound(Split(.Text, vbNewLine)) <= 200
            newText = Split(.Text, vbNewLine, 6)
            .SelStart = 0
            .SelLength = Len(newText(0)) + Len(newText(1)) + Len(newText(2)) + Len(newText(3)) + Len(newText(4)) + 10
            .SelText = ""
        Loop
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelColor = vbBlack
        .SelText = strTimeStamp
        .SelStart = Len(.Text)
        Data = strTimeStamp
        Dim i As Byte
        For i = LBound(saElements) To UBound(saElements) Step 2
            .SelStart = Len(.Text)
            .SelLength = 0
            .SelColor = saElements(i)
            .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
            .SelStart = Len(.Text)
            Data = Data & saElements(i + 1)
        Next i
    End With
End Sub
Public Sub Connect(Index As Byte)
    On Error GoTo ErrorEncountered
    Dim Proxy() As String
    Dim blnConnect As Boolean
    Dim strConnectionType As String
    Dim strIPBots As String
    
    With frmMain
        .mnuConnect_Current.Enabled = False
        Wait (Index)
        strConnectionType = ReadINI("Connection", "ConnectionType")
        strIPBots = ReadINI(LoadResString(102), LoadResString(126))
        If strConnectionType = 0 Then
            Proxy = Split(GetProxy(Index), ":"): blnConnect = True
            blnConnectionProxied(Index) = True
        Else
            blnConnectionProxied(Index) = False
            If strConnectionType = 1 Then
                If Index < strIPBots Then blnConnect = True
            Else
                If Index >= strIPBots Then
                    Proxy = Split(GetProxy(Index), ":")
                    blnConnectionProxied(Index) = True
                Else
                    blnConnect = True
                End If
            End If
        End If
        If blnConnectionProxied(Index) = True Then
            .Winsock(Index).Connect Proxy(0), Proxy(1)
            Timeout(Index) = 0
            TimeoutEnabled(Index) = True
            TimeoutP(Index) = LoadResString(129)
            AddC Val#(Index), vbYellow, LoadResString(128)
        Else
            If blnConnect = True Then .Winsock(Index).Connect Settings.Server, 6112
            Timeout(Index) = 0
            TimeoutEnabled(Index) = True
            TimeoutP(Index) = LoadResString(130)
            AddC Val#(Index), vbYellow, LoadResString(109)
        End If
    End With
    Connecting(Index) = True
    frmSettings.cmbProfiles.Locked = True
    
    Exit Sub
ErrorEncountered:
    Select Case Err.Number
        Case 9: MsgBox LoadResString(131), vbCritical, LoadResString(116)
        Case Else: MsgBox Err.Number & ": " & Err.Description, vbCritical, LoadResString(116)
    End Select
End Sub
Public Sub errHandle(Location As String)
    Close #6
    DoEvents
    Open App.Path & "\Logs\Errors.log" For Append As #6
    Print #6, Location
    Print #6, Err.Number & ": " & Err.Description
    Print #6, Err.Source
    Print #6, ""
    Close #6
    MsgBox Err.Number & ": " & Err.Description & vbCrLf & "Please report this error to: Fapiko@Fapiko.Com", vbCritical
End Sub
Public Function GetProxy(Index As Byte) As String
    On Error GoTo ErrorEncountered
    Dim BPP As Byte
    Dim i As Byte
    Dim s() As String
    
    For i = 0 To UBound(uProxies())
        BPP = 3
        If LenB(ReadINI(LoadResString(102), LoadResString(182))) > 0 Then BPP = ReadINI(LoadResString(102), LoadResString(182))
        If uProxies(i) < 3 Then
            GetProxy = Proxies(i)
            uProxies(i) = uProxies(i) + 1
            cProxy(Index) = i
            Exit Function
        End If
    Next i
    
    Exit Function
ErrorEncountered:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, LoadResString(116) & Index
End Function
Public Sub formDrag(theForm As Form)
    ReleaseCapture
    Call SendMessage(theForm.hWnd, &HA1, 2, 0&)
End Sub
Public Sub LoadProxies()
    Dim i As Byte
    Dim Proxy As String
    ReDim Preserve Proxies(0)
    
    Open App.Path & LoadResString(132) For Input As #1
    Do Until EOF(1)
        Input #1, Proxy
        Proxies(UBound(Proxies)) = Proxy
        ReDim Preserve Proxies(UBound(Proxies) + 1)
    Loop
    Close #1
    
    ReDim Preserve Proxies(UBound(Proxies) - 1)
    ReDim uProxies(UBound(Proxies))
    For i = 0 To UBound(Proxies())
        uProxies(i) = 0
    Next i
End Sub
Public Sub Queue(Index As Byte, Message As String)
    frmMain.lstQueue(Index).AddItem Message
End Sub
Public Function ReadINI(riSection As String, riKey As String) As String
    Dim riFile As String
    Dim sRiBuffer As String
    Dim sRiValue As String
    Dim sRiLong As String
    
    riFile = (App.Path & LoadResString(137))
    If Dir(riFile) <> "" Then
        sRiBuffer = String(255, vbNull)
        sRiLong = GetPrivateProfileString(riSection, riKey, Chr(1), sRiBuffer, 255, riFile)
        If Left(sRiBuffer, 1) <> Chr(1) Then
            sRiValue = Left(sRiBuffer, sRiLong)
            ReadINI = sRiValue
        End If
    Else
        ReadINI = ""
    End If
End Function
Public Sub Reconnect(Index As Byte)
    Wait (Index)
    RotateProxies
    Waiting(Index) = 1
End Sub
Public Sub RemoveKey(Index As Byte)
    On Error GoTo HandleError
    Dim i As Integer
    
    Open App.Path & LoadResString(119) & Settings.Profile & LoadResString(115) For Output As #4
    With frmMain.lbCDKeys
        For i = 0 To .ListCount - 1
            If .List(i) = Key(Index) Then
                .RemoveItem (i)
                Open App.Path & LoadResString(119) & Settings.Profile & LoadResString(217) & LoadResString(115) For Append As #5
                Print #5, Key(Index)
                Close #5
                Key(Index) = vbNullChar
            Else
                Print #4, .List(i)
            End If
        Next i
    End With
    Close #4
HandleError:
End Sub
Public Sub RemoveProxy(Index As Byte)
    Dim i As Integer
    For i = Index To UBound(Proxies()) - 1
        Proxies(i) = Proxies(i + 1)
        uProxies(i) = uProxies(i + 1)
    Next i
    If UBound(Proxies) = 0 Then GoTo NoProxies
    ReDim Preserve Proxies(UBound(Proxies) - 1)
    ReDim Preserve uProxies(UBound(uProxies) - 1)
    
    Exit Sub
NoProxies:
    MsgBox LoadResString(138), vbCritical, LoadResString(116)
    Wait (Index)
    Timeout(Index) = 0
    TimeoutEnabled(Index) = False
End Sub
Public Sub RotateProxies()
    Dim i As Integer
    Dim fProxy As String
    Dim fUProxy As String
    
    fProxy = Proxies(0)
    fUProxy = uProxies(0)
    For i = 0 To UBound(Proxies()) - 1
        Proxies(i) = Proxies(i + 1)
        uProxies(i) = uProxies(i + 1)
    Next i
    Proxies(UBound(Proxies)) = fProxy
    uProxies(UBound(Proxies)) = fUProxy
End Sub
Public Sub SendChat(Index As Byte, Message As String)
    If Connected(Index) = False Then Exit Sub
    If LenB(intTime(Index)) = 0 Then intTime(Index) = 0
    intTime(Index) = Runs(Index) + 2
    If Len(Message) > 50 Then intTime(Index) = Runs(Index) + 4
    If Len(Message) > 100 Then intTime(Index) = Runs(Index) + 6
    If Len(Message) > 150 Then intTime(Index) = Runs(Index) + 9
    If Len(Message) > 200 Then intTime(Index) = Runs(Index) + 12
    
    With PBuffer
        .InsertNTString Left$(Message, 224)
        .SendPacket SID_CHATCOMMAND, Index
    End With
End Sub
Public Sub TimedOut(Index As Byte)
    Dim Proxy() As String
    
    With frmMain.Winsock(Index)
        Wait (Index)
        If iSocksified(Index) = True Then
            .Connect
        Else
            AddC Val#(Index), vbRed, LoadResString(139)
            AddC Val#(Index), vbRed, LoadResString(113) & Proxies(cProxy(Index))
            RotateProxies
            Connect (Index)
        End If
    End With
End Sub
Public Sub Wait(Index As Byte)
    Connected(Index) = False
    Connecting(Index) = False
    If Index = frmMain.cmbProfiles.ListIndex Then
        With frmMain
            .mnuConnect_Current.Enabled = True
            .mnuBot_Current.Enabled = False
            .mnuDisconnect_Current.Enabled = False
            .mnuReconnect = False
        End With
    End If
    
    If lngNLS(Index) <> 0 Then nls_free (lngNLS(Index)): lngNLS(Index) = 0
    If iSocksified(Index) = True Then uProxies(cProxy(Index)) = uProxies(cProxy(Index)) - 1
    iSocksified(Index) = False
    Waiting(Index) = 0
    frmMain.Winsock(Index).Close
    Do Until frmMain.Winsock(Index).State = sckClosed
        DoEvents
    Loop
End Sub
Public Sub WriteINI(wiSection As String, wiKey As String, wiValue As String)
    Dim wiFile As String
    wiFile = (App.Path & LoadResString(137))
    WritePrivateProfileString wiSection, wiKey, wiValue, wiFile
End Sub
