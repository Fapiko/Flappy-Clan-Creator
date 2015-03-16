Attribute VB_Name = "rPackets"
Option Explicit
Public Sub ParseData(Data As String, Index As Byte)
    If LenB(Data) = 0 Then Exit Sub
    Dim PacketID As Integer
    
    If Asc(Left(Data, 1)) <> 255 Then
        bData = bData & Data
        If Len(bData) < lData Then Exit Sub
    ElseIf PBuffer.GetWORD(Mid(Data, 3, 2)) > Len(Data) Then
        bData = Data
        lData = PBuffer.GetWORD(Mid(Data, 3, 2))
        Exit Sub
    Else
        bData = Data
    End If
    
    PacketID = Asc(Mid(bData, 2, 1))

    Select Case PacketID
        Case SID_NULL: PBuffer.SendPacket SID_NULL, Index
        Case SID_ENTERCHAT: Parse0x0A bData, Index
        Case SID_CHATEVENT: Parse0x0F bData, Index
        Case SID_PING: Parse0x25 bData, Index
        Case SID_AUTH_INFO: Parse0x50 bData, Index
        Case SID_AUTH_CHECK: Parse0x51 bData, Index
        Case SID_AUTH_ACCOUNTCREATE: Parse0x52 bData, Index
        Case SID_AUTH_ACCOUNTLOGON: Parse0x53 bData, Index
        Case SID_AUTH_ACCOUNTLOGONPROOF: Parse0x54 bData, Index
        Case SID_FRIENDSLIST: Parse0x65 bData, Index
        Case SID_FRIENDSUPDATE:
        Case SID_FRIENDSADD: Parse0x67 bData, Index
        Case SID_CLANFINDCANDIDATES: Parse0x70 bData, Index
        Case SID_CLANINVITEMULTIPLE: Parse0x71 bData, Index
        Case SID_CLANSENDINVITE: Parse0x72 bData, Index
        Case SID_CLANINFO: Parse0x75 bData, Index
        Case SID_CLANQUITNOTIFY:
        Case SID_CLANREMOVEMEMBER: Parse0x78 bData, Index
        Case Else: ParseUnknown bData, Index
    End Select
End Sub
Public Sub Parse0x0A(Data As String, Index As Byte)
    Dim StatString As String
    Dim Username As String
    
    With PBuffer
        Username = .KillNull(Mid(Data, 5))
        StatString = .KillNull(Mid(Data, 5 + Len(Username)))
    End With
    Connecting(Index) = False
    
    AddC Index, vbYellow, LoadResString(140), vbGreen, Username, vbYellow, "."
    If InStr(1, Username, "#") Then AddC Index, vbRed, LoadResString(141) & Username & LoadResString(142)
End Sub
Public Sub Parse0x0F(Data As String, Index As Byte)
    Dim EID As Byte
    Dim Flags As Byte
    Dim Ping As Long
    Dim Client As String
    Dim Message As String
    Dim Username As String
    
    Username = PBuffer.KillNull(Mid(Data, 29))
    Message = PBuffer.KillNull(Mid$(Data, Len(Username) + 30))
    Ping = PBuffer.MakeLONG(Mid$(Data, 13, 4))
    Client = StrReverse(PBuffer.KillNull(Mid$(Data, Len(Username) + 30, 4)))
    Flags = PBuffer.MakeLONG(Mid$(Data, 9, 4))
    EID = PBuffer.MakeLONG(Mid$(Data, 5, 4))
    
    Select Case EID
        Case EID_WHISPER: AddC Index, vbOrange, "<From ", vbYellow, Username, vbOrange, "> ", vbGrey, Message
        Case EID_TALK:
            Message = Replace$(Message, "\rtf", "/rtf")
            Select Case Flags
                Case 2, 18: AddC Index, vbTeal, "<", vbWhite, Username, vbTeal, "> ", vbWhite, Message
                Case 17: AddC Index, vbRBlue, "<", vbYellow, Username, vbRBlue, "> " & Message
                Case Else: AddC Index, vbTeal, "<", vbYellow, Username, vbTeal, "> ", vbWhite, Message
            End Select
        Case EID_CHANNEL:
            AddC Index, vbTurquoise, LoadResString(143), vbYellow, Message, vbTurquoise, LoadResString(120)
            If Index <> 0 Then Connected(Index) = True
            If Message = LoadResString(144) Then
                If Index <> 0 Then
                    PBuffer.SendPacket SID_FRIENDSLIST, Index
                Else
                    AddC 0, vbRed, LoadResString(145)
                    Wait (0)
                    Waiting(0) = 1
                End If
            Else
                If Index = 0 Then Connected(Index) = True
            End If
        Case EID_WHISPERSENT: AddC Index, vbOrange, LoadResString(146), vbYellow, Username, vbOrange, "> ", vbGrey, Message
        Case EID_CHANNELFULL: AddC Index, vbRed, LoadResString(147)
        Case EID_CHANNELDOESNOTEXIST: AddC Index, vbRed, LoadResString(148)
        Case EID_CHANNELRESTRICTED: AddC Index, vbRed, LoadResString(149)
        Case EID_INFO: AddC Index, vbTeal, Message
        Case EID_ERROR:
            AddC Index, vbRed, Message
            If Message = LoadResString(150) Then
                RemoveFriend(Index) = True
                PBuffer.SendPacket SID_FRIENDSLIST, Index
            End If
        Case EID_EMOTE: AddC Index, vbYellow, "<", vbWhite, Username & " ", vbYellow, Message & ">"
    End Select
End Sub
Public Sub Parse0x25(Data As String, Index As Byte)
    Dim Ping As Long
    Ping = PBuffer.GetDWORD(Mid(Data, 5, 4))
    
    With PBuffer
        .InsertDWORD Ping
        .SendPacket SID_PING, Index
    End With
End Sub
Public Sub Parse0x50(Data As String, Index As Byte)
    ServerToken = Val("&h" & PBuffer.StrToHex(StrReverse(Mid(Data, 9, 4))))
    HashCommand = PBuffer.KillNull(Mid$(Data, 38))
    lngMPQNumber = Mid$(Mid$(Data, InStr(1, Data, LoadResString(152)), Len(Data)), 8, 1)
    
    AddC Index, vbGreen, LoadResString(151)
    Send0x51 Index
End Sub

Public Sub Parse0x51(Data As String, Index As Byte)
    Select Case PBuffer.GetDWORD(Mid(Data, 5, 4))
        Case &H0:
            AddC Index, vbGreen, LoadResString(153)
            If lngNLS(Index) <> 0 Then nls_free (lngNLS(Index))
            lngNLS(Index) = 0
            Timeout(Index) = 0
            TimeoutEnabled(Index) = False
            TimeoutP(Index) = vbNullChar
            Send0x53 Accounts.U(Index), Index
        Case &H100: AddC Index, vbRed, LoadResString(154) & PBuffer.KillNull(Mid(Data, 9))
        Case &H101: AddC Index, vbRed, LoadResString(155)
        Case &H102: AddC Index, vbRed, LoadResString(156) & PBuffer.KillNull(Mid(Data, 9))
        Case &H200:
            AddC Index, vbRed, LoadResString(157)
            RemoveKey (Index)
            Reconnect (Index)
        Case &H201:
            AddC Index, vbRed, LoadResString(158) & PBuffer.KillNull(Mid(Data, 9))
            Reconnect (Index)
        Case &H202:
            AddC Index, vbRed, LoadResString(159)
            RemoveKey (Index)
            Reconnect (Index)
        Case &H203:
            AddC Index, vbRed, LoadResString(160)
            RemoveKey (Index)
            Reconnect (Index)
        Case Else: AddC Index, vbRed, LoadResString(161)
    End Select
End Sub
Public Sub Parse0x52(Data As String, Index As Byte)
    Select Case PBuffer.GetDWORD(Mid(Data, 5, 4))
      Case &H0:
        AddC Index, vbGreen, LoadResString(162)
        Send0x53 Accounts.U(Index), Index
      Case &H6: AddC Index, vbRed, LoadResString(163): Wait (Index)
      Case &H7: AddC Index, vbRed, LoadResString(164): Wait (Index)
      Case &H8: AddC Index, vbRed, LoadResString(165): Wait (Index)
      Case &H9: AddC Index, vbRed, LoadResString(166): Wait (Index)
      Case &HA: AddC Index, vbRed, LoadResString(167): Wait (Index)
      Case &HB: AddC Index, vbRed, LoadResString(168): Wait (Index)
      Case &HC: AddC Index, vbRed, LoadResString(169): Wait (Index)
      Case Else: AddC Index, vbRed, LoadResString(170): Wait (Index)
    End Select
End Sub
Public Sub Parse0x53(Data As String, Index As Byte)
    Select Case PBuffer.GetDWORD(Mid(Data, 5, 4))
      Case &H0: Send0x54 Mid(Data, 9, 32), Mid(Data, 41, 32), Index
      Case &H1:
        AddC Index, vbRed, LoadResString(171)
        Send0x52 Index
      Case &H5: AddC Index, vbRed, LoadResString(172): Wait (Index)
      Case Else: AddC Index, vbRed, LoadResString(170): Wait (Index)
    End Select
End Sub
Public Sub Parse0x54(Data As String, Index As Byte)
    Select Case PBuffer.GetDWORD(Mid(Data, 5, 4))
      Case &H0:
        TimeoutEnabled(Index) = False
        Timeout(Index) = 0
        With PBuffer
          .InsertNTString Accounts.U(Index)
          .InsertBYTE &H0
          .SendPacket SID_ENTERCHAT, Index
          .InsertDWORD &H2
          If LenB(Settings.HomeChannel) > 0 Then
            .InsertNTString Settings.HomeChannel
          Else
            .InsertNTString LoadResString(173)
          End If
          .SendPacket SID_JOINCHANNEL, Index
        End With
      Case &H2
        AddC Index, vbRed, LoadResString(174) & Accounts.U(Index)
        AddC Index, vbRed, Accounts.U(Index) & " - " & Accounts.P(Index)
      Case &HE
        AddC Index, vbYellow, LoadResString(175)
        Send0x59 Index
        With PBuffer
            .InsertNTString Accounts.U(Index)
            .InsertBYTE &H0
            .SendPacket SID_ENTERCHAT, Index
            .InsertDWORD &H2
            If LenB(Settings.HomeChannel) > 0 Then
                .InsertNTString Settings.HomeChannel
            Else
                .InsertNTString LoadResString(173)
            End If
            .SendPacket SID_JOINCHANNEL, Index
        End With
      Case Else: AddC Index, vbRed, LoadResString(170): Wait (Index)
    End Select
End Sub
Public Sub Parse0x65(Data As String, Index As Byte)
    Dim Friends As Byte
    Dim i As Byte
    Dim j As Byte
    Dim Location As Byte
    Dim Status As Byte
    Dim iUser As Integer
    Dim Channel As String
    Dim Username As String
    
    Friends = PBuffer.GetByte(Mid(Data, 5, 1))
    For i = 1 To Friends
        Username = PBuffer.KillNull(Mid(Data, iUser + 6))
        Status = PBuffer.GetByte(Mid(Data, iUser + 7 + Len(Username), 1))
        Channel = PBuffer.KillNull(Mid(Data, iUser + 13 + Len(Username)))
        iUser = iUser + 8 + Len(Username) + Len(Channel)
        
        If RemoveFriend(Index) = True Then
            RemoveFriend(Index) = False
            With PBuffer
                If Index <> 0 Then
                    Queue Index, LoadResString(176) & Username '/f r
                    Queue Index, LoadResString(177) & Accounts.U(Index) '/f a
                Else
                    For j = 1 To 9
                        If AddChief(j) = True Then
                            Queue 0, LoadResString(176) & Username '/f r
                            Queue 0, LoadResString(177) & Accounts.U(j) '/f a
                            AddChief(j) = False
                            j = 9
                        End If
                    Next j
                End If
            End With
            Exit Sub
        End If
        
        If Index <> 0 And Username = Accounts.U(0) Then
            Select Case Status
                Case 1, 3, 5, 7: Exit Sub
                Case Else:
                    AddChief(Index) = True
                    Queue 0, LoadResString(177) & Accounts.U(Index) '/f a
                Exit Sub
            End Select
        End If
    Next i
    Queue 0, LoadResString(177) & Accounts.U(Index) '/f a
    AddChief(Index) = True
    Queue Index, LoadResString(177) & Accounts.U(0) '/f a
End Sub
Public Sub Parse0x67(Data As String, Index As Byte)
    AddC Index, vbGreen, LoadResString(178), vbYellow, PBuffer.KillNull(Mid$(Data, 5)), vbGreen, LoadResString(179)
End Sub
Public Sub Parse0x70(Data As String, Index As Byte)
    Dim Candidates As Byte
    Dim i As Byte
    Dim Status As Byte
    Dim cLocation As Integer
    Dim Cookie As String
    Dim cUser As String
    
    Cookie = PBuffer.GetDWORD(Mid$(Data, 5, 4))
    Status = PBuffer.GetByte(Mid(Data, 9, 1))
    Candidates = PBuffer.GetByte(Mid(Data, 10, 1))
    
    Select Case Status
        Case 0:
            AddC Index, vbGreen, LoadResString(180), vbYellow, Candidates, vbGreen, LoadResString(181)
            If Cookie = 3 Then
                With PBuffer
                    .InsertDWORD .MakeLONG(LoadResString(105))
                    .InsertNTString Settings.ClanName
                    Select Case Len(Settings.ClanTag)
                        Case 2: .InsertBYTE &H0: .InsertBYTE &H0
                        Case 3: .InsertBYTE &H0
                    End Select
                    .InsertNonNTString StrReverse$(Settings.ClanTag)
                    .InsertBYTE Val#(Candidates)
                    cLocation = 11
                    For i = 1 To Candidates
                        cUser = PBuffer.KillNull(Mid$(Data, cLocation))
                        cLocation = cLocation + Len(cUser) + 1
                        .InsertNTString cUser
                    Next i
                    .SendPacket SID_CLANINVITEMULTIPLE, 0
                End With
            End If
        Case 1: AddC Index, vbRed, LoadResString(184), vbWhite, Settings.ClanTag, vbRed, LoadResString(185): Exit Sub
        Case 2:
            AddC Index, vbRed, LoadResString(186)
            MsgBox LoadResString(188), vbInformation, Accounts.U(Index)
            Exit Sub
        Case 8: AddC Index, vbRed, LoadResString(187): Exit Sub
        Case &HA: AddC Index, vbRed, LoadResString(184), vbWhite, Settings.ClanTag, vbRed, LoadResString(189): Exit Sub
        Case Else: AddC Index, vbRed, LoadResString(190) & Status: Exit Sub
    End Select
    
    If Index = 0 And Candidates >= 9 Then frmMain.lbCreate.Enabled = True
    If Index = 0 And Candidates < 9 Then frmMain.lbCreate.Enabled = False
End Sub
Public Sub Parse0x71(Data As String, Index As Byte)
    Dim Status As Integer
    Status = PBuffer.GetDWORD(Mid(Data, 9, 4))
    
    If Status = 0 Then
        ClanCreated = True
        AddC Index, vbGreen, LoadResString(191), vbYellow, "Clan " & Settings.ClanTag, vbGreen, "!"
    End If
End Sub
Public Sub Parse0x72(Data As String, Index As Byte)
Dim Token As String

Token = Mid(Data, 5, 4)
With PBuffer
    .InsertNonNTString Token
    Select Case Len(Settings.ClanTag)
        Case 2: .InsertBYTE &H0: .InsertBYTE &H0
        Case 3: .InsertBYTE &H0
    End Select
    .InsertNonNTString StrReverse$(Settings.ClanTag)
    .InsertNTString Accounts.U(0)
    .InsertBYTE 6
    .SendPacket SID_CLANSENDINVITE, Index
End With
End Sub
Public Sub Parse0x75(Data As String, Index As Byte)
    Dim Tag As String
    Dim Rank As Byte
    Dim Remove As Byte
    
    Tag = PBuffer.KillNull(StrReverse(Mid(Data, 6, 4)))
    Rank = PBuffer.GetByte(Mid(Data, 10, 1))
    
    If ClanCreated = False Then
        Select Case Rank
            Case 0: MsgBox LoadResString(192) & Tag & LoadResString(193), vbInformation, Accounts.U(Index)
            Case 1, 2, 3: Remove = MsgBox(LoadResString(192) & Tag & LoadResString(194), vbYesNo, Accounts.U(Index))
            Case 4: MsgBox LoadResString(192) & Tag & LoadResString(195), vbInformation, Accounts.U(Index)
        End Select
    End If
    
    If Remove = 6 Then
        With PBuffer
            .InsertDWORD 1
            .InsertNTString Accounts.U(Index)
            .SendPacket SID_CLANREMOVEMEMBER, Index
        End With
    End If
End Sub
Public Sub Parse0x78(Data As String, Index As Byte)
    Dim Status As Byte
    
    Status = PBuffer.GetByte(Mid(Data, 9, 1))
    If Status = 0 Then
        AddC Index, vbGreen, LoadResString(196)
    Else
        AddC Index, vbRed, LoadResString(197)
    End If
End Sub
Public Sub ParseUnknown(Data As String, Index As Byte)
Dim r() As String
r = Split(PBuffer.StrToHex(Data))
AddC Index, vbRed, LoadResString(198) & r(2)
AddC Index, vbRed, PBuffer.DebugOutput(Data)
End Sub
