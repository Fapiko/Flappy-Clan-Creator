Attribute VB_Name = "sPackets"
Option Explicit
Public Sub Send0x50(Index As Byte)
    With PBuffer
       .InsertDWORD &H0
       .InsertNonNTString LoadResString(199)
       .InsertDWORD "&H" & ReadINI(LoadResString(102), LoadResString(200))
       .InsertDWORD &H0
       .InsertDWORD &H0
       .InsertDWORD &H0
       .InsertDWORD &H0
       .InsertDWORD &H0
       .InsertNTString LoadResString(201) 'USA
       .InsertNTString LoadResString(202) 'United States
       .SendPacket SID_AUTH_INFO, Index
    End With
End Sub
Public Sub Send0x51(Index As Byte)
    Dim Checksum As Long
    Dim CRevision As Long
    Dim ClientToken As Long
    Dim lngDecoder As Long
    Dim HashLength As Long
    Dim Version As Long
    Dim CDKey As String
    Dim EXEInfo As String
    Dim KeyHash As String
    
    EXEInfo = Space(256)
    CRevision = checkRevision(HashCommand, App.Path & LoadResString(203), App.Path & LoadResString(204), App.Path & LoadResString(205), lngMPQNumber, Checksum)
    If CRevision = 0 Then
        AddC Index, vbRed, LoadResString(206)
        If blnConnectionProxied(Index) = True Then uProxies(cProxy(Index)) = uProxies(cProxy(Index)) - 1
        Wait (Index)
        Exit Sub
    End If
    
    AddC Index, vbYellow, LoadResString(207)
    getExeInfo App.Path & LoadResString(203), EXEInfo
    EXEInfo = PBuffer.KillNull(EXEInfo)
    ClientToken = GetTickCount()
    
    CDKey = frmMain.lbCDKeys.List(0)
    lngDecoder = kd_create(CDKey, Len(CDKey))
    If lngDecoder = -1 Then AddC Index, vbRed, LoadResString(208), vbBlack, CDKey, vbRed, LoadResString(209): Wait (Index): Exit Sub
    HashLength = kd_calculateHash(lngDecoder, ClientToken, ServerToken)
    If HashLength = 0 Then AddC Index, vbRed, LoadResString(208), vbBlack, CDKey, vbRed, LoadResString(210): Wait (Index): Exit Sub
    KeyHash = String$(HashLength, vbNullChar)
    Call kd_getHash(lngDecoder, KeyHash)
    
    With PBuffer
        .InsertDWORD ClientToken
        .InsertDWORD Version
        .InsertDWORD Checksum
        .InsertDWORD &H1
        .InsertDWORD &H0
        .InsertDWORD &H1A
        .InsertDWORD kd_product(lngDecoder)
        .InsertDWORD kd_val1(lngDecoder)
        .InsertDWORD &H0
        .InsertNonNTString KeyHash
        .InsertNTString EXEInfo
        .InsertNTString Accounts.U(Index) & LoadResString(210)
        .SendPacket SID_AUTH_CHECK, Index
    End With
    
    Call kd_free(lngDecoder)
    
    With frmMain.lbCDKeys
        Key(Index) = .List(0)
        .RemoveItem (0)
        .AddItem Key(Index)
    End With
End Sub
Public Sub Send0x52(Index As Byte)
    Dim Buffer As String, BufLen As Long
    
    BufLen = 65 + Len(Accounts.U(Index))
    Buffer = String$(BufLen, vbNullChar)
    If (nls_account_create(lngNLS(Index), Buffer, BufLen) = 0) Then
        Wait (Index)
        AddC Index, vbRed, LoadResString(212)
        Exit Sub
    End If
    
    With PBuffer
        .InsertNonNTString Buffer
        .SendPacket SID_AUTH_ACCOUNTCREATE, Index
    End With
End Sub
Public Sub Send0x53(Username As String, Index As Byte)
    Dim strReturn As String
    strReturn = String(32, vbNullChar)
    
    If lngNLS(Index) = 0 Then lngNLS(Index) = nls_init(Accounts.U(Index), Accounts.P(Index))
    If lngNLS(Index) = 0 Then AddC vbRed, LoadResString(213)
    Call nls_get_A(lngNLS(Index), strReturn)
    With PBuffer
        .InsertNonNTString strReturn
        .InsertNTString Username
        .SendPacket SID_AUTH_ACCOUNTLOGON, Index
    End With
End Sub
Public Sub Send0x54(strSalt As String, strPubKey As String, Index As Byte)
    Dim strReturn As String
    
    strReturn = String(20, vbNullChar)
    Call nls_get_M1(lngNLS(Index), strReturn, strPubKey, strSalt)
    
    With PBuffer
        .InsertNonNTString strReturn
        .SendPacket SID_AUTH_ACCOUNTLOGONPROOF, Index
    End With
End Sub
Public Sub Send0x59(Index As Byte)
    With PBuffer
        If Settings.RegEmail <> "" Then
            .InsertNTString Settings.RegEmail
            AddC Index, vbGreen, LoadResString(214) & Settings.RegEmail
        Else: .InsertNTString LoadResString(215)
        End If
        .SendPacket SID_SETEMAIL, Index
    End With
End Sub
