VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   3615
   ClientLeft      =   4185
   ClientTop       =   3510
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5535
   Begin VB.ListBox lbCDKeys 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrSecond 
      Interval        =   1000
      Left            =   360
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   360
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.FileListBox flbProfiles 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame frFlappy 
      Caption         =   "Flappy Clan Creator"
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
      Begin VB.Timer tmrUpdate 
         Interval        =   100
         Left            =   240
         Top             =   1680
      End
      Begin VB.Timer tmrDelay 
         Interval        =   100
         Left            =   240
         Top             =   1320
      End
      Begin VB.ListBox lstQueue 
         Appearance      =   0  'Flat
         Height          =   1785
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cmbProfiles 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
      Begin RichTextLib.RichTextBox rtbStatus 
         Height          =   2655
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4683
         _Version        =   393217
         BackColor       =   14737632
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":0ECA
      End
      Begin VB.Label lbCreate 
         BackColor       =   &H80000004&
         Caption         =   "Create Clan"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Menu mnuBot 
      Caption         =   "&Bot"
      Begin VB.Menu mnuBot_SelectProfile 
         Caption         =   "Load &Settings"
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBot_All 
         Caption         =   "Check &All"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBot_Current 
         Caption         =   "Check &Current"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBot_Quit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connect"
      Begin VB.Menu mnuConnect_Current 
         Caption         =   "C&urrent Account"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConnect_All 
         Caption         =   "&All"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuDisconnect 
      Caption         =   "&Disconnect"
      Begin VB.Menu mnuDisconnect_Current 
         Caption         =   "&Current Account"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDisconnect_All 
         Caption         =   "&All"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuReconnect 
      Caption         =   "&Reconnect"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuMinimize 
      Caption         =   "&Minimize"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbProfiles_Click()
    Dim i As Byte
    Dim Index As Byte
    
    Index = cmbProfiles.ListIndex
    If Winsock(Index).State <> sckConnected Then mnuConnect_Current.Enabled = True
    
    If Winsock(Index).State = sckConnected Then
        mnuDisconnect_Current.Enabled = True
        mnuReconnect.Enabled = True
    End If
    
    If Connected(Index) = True Then mnuBot_Current.Enabled = True
    
    For i = 0 To 9
        If i <> cmbProfiles.ListIndex Then rtbStatus(i).Visible = False
    Next i
    rtbStatus(cmbProfiles.ListIndex).Visible = True
End Sub

Private Sub Form_Load()
    Dim i As Byte
    
    flbProfiles.Path = App.Path & LoadResString(104)
    For i = 1 To 9
        Load rtbStatus(i)
        Load Winsock(i)
        Load lstQueue(i)
    Next i
    LoadProxies
    Call kd_init
    WriteINI "Auto Updater", "Version", App.Major & "." & App.Minor & "." & App.Revision
    frFlappy.Caption = LoadResString(101) & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    formDrag Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub frFlappy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    formDrag Me
End Sub

Private Sub lbCreate_Click()
    With PBuffer
        .InsertDWORD 3
        Select Case Len(Settings.ClanTag)
            Case 2: .InsertBYTE &H0: .InsertBYTE &H0
            Case 3: .InsertBYTE &H0
        End Select
        .InsertNonNTString StrReverse$(Settings.ClanTag)
        .SendPacket SID_CLANFINDCANDIDATES, cmbProfiles.ListIndex
    End With
End Sub

Private Sub mnuBot_All_Click()
    Dim i As Byte
    
    For i = 0 To 9
        If Connected(i) = True Then
            With PBuffer
                .InsertDWORD 1
                Select Case Len(Settings.ClanTag)
                    Case 2: .InsertBYTE &H0: .InsertBYTE &H0
                    Case 3: .InsertBYTE &H0
                End Select
                .InsertNonNTString StrReverse$(Settings.ClanTag)
                .SendPacket SID_CLANFINDCANDIDATES, i
            End With
        End If
    Next i
End Sub

Private Sub mnuBot_Current_Click()
    If Connected(cmbProfiles.ListIndex) = True Then
        With PBuffer
            .InsertDWORD 1
            Select Case Len(Settings.ClanTag)
                Case 2: .InsertBYTE &H0: .InsertBYTE &H0
                Case 3: .InsertBYTE &H0
            End Select
            .InsertNonNTString StrReverse$(Settings.ClanTag)
            .SendPacket SID_CLANFINDCANDIDATES, cmbProfiles.ListIndex
        End With
    End If
End Sub

Private Sub mnuBot_Quit_Click()
    End
End Sub

Private Sub mnuBot_SelectProfile_Click()
    frmSettings.Visible = True
End Sub

Private Sub mnuConnect_All_Click()
    Dim i As Byte
    
    For i = 0 To 9
        If Winsock(i).State <> sckConnected Then Connect (i)
    Next i
End Sub

Private Sub mnuConnect_Current_Click()
    Dim Index As Byte
    
    Index = cmbProfiles.ListIndex
    Connect (Index)
End Sub

Private Sub mnuDisconnect_All_Click()
    Dim i As Byte
    
    For i = 0 To 9
        If Winsock(i).State <> sckClosed Then
            AddC i, vbRed, LoadResString(106)
            Wait (i)
        End If
    Next i
End Sub

Private Sub mnuDisconnect_Current_Click()
    AddC cmbProfiles.ListIndex, vbRed, LoadResString(106)
    Wait (cmbProfiles.ListIndex)
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuMinimize_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub mnuReconnect_Click()
    Wait (cmbProfiles.ListIndex)
    Connect (cmbProfiles.ListIndex)
End Sub

Private Sub tmrDelay_Timer()
    On Error GoTo HandleError
    Dim i As Byte
    
    For i = 0 To 9
        If Connected(i) = True Then
            If LenB(intTime(i)) > 0 Then
                If intTime(i) > 0 Then
                    If intTime(i) <= 1 Then
                        With lstQueue(i)
                            If .ListCount > 0 Then
                                SendChat i, .List(0)
                                .RemoveItem (0)
                                If Runs(i) <= 6 Then Runs(i) = Runs(i) + 1
                            End If
                        End With
                    End If
                Else
                    With lstQueue(i)
                    If .ListCount > 0 Then
                        SendChat i, .List(0)
                        .RemoveItem (0)
                    End If
                    End With
                    Runs(i) = 1
                End If
            Else
                With lstQueue(i)
                    If .ListCount > 0 Then
                        SendChat i, .List(0)
                        .RemoveItem (0)
                    End If
                End With
                Runs(i) = 1
            End If
        End If
    Next i
    
    Exit Sub
HandleError:
    errHandle ("Private Sub tmrDelay_Timer()")
End Sub

Private Sub tmrSecond_Timer()
    Dim i As Byte
    Dim iTimeout As Byte
    Dim Check As Boolean
    Dim Disconnect As Boolean
    
    For i = 0 To 9
        If Winsock(Val#(i)).State <> sckClosed Then
            mnuDisconnect_All.Enabled = True
            Disconnect = True
            i = 9
        End If
    Next i
    If Disconnect = False Then mnuDisconnect_All.Enabled = False
    
    For i = 0 To 9
        If Connected(i) = True Then
            mnuBot_All.Enabled = True
            Check = True
            i = 9
        End If
    Next i
    If Check = False Then mnuBot_All.Enabled = False
    iTimeout = 15
    If Settings.ProxyTimeout <> 0 Then iTimeout = Settings.ProxyTimeout
    
    For i = 0 To 9
        If TimeoutEnabled(i) = True Then Timeout(i) = Timeout(i) + 1
        If Timeout(i) > iTimeout = True Then
            Timeout(i) = 0
            TimeoutEnabled(i) = False
            TimedOut (i)
        End If
        If Connected(i) = True And cmbProfiles.ListIndex = i Then
            mnuBot_Current.Enabled = True
            mnuDisconnect_Current.Enabled = True
        End If
        If Waiting(i) >= 5 Then
            Connect (i)
            Waiting(i) = 0
        End If
        If Waiting(i) > 0 Then Waiting(i) = Waiting(i) + 1
    Next i
End Sub

Private Sub tmrUpdate_Timer()
    Dim i As Byte
    For i = 0 To 9
        If LenB(intTime(i)) > 0 Then
            If intTime(i) > 0 Then
                intTime(i) = intTime(i) - 1
            End If
            If intTime(i) < 4 Then
                tmrUpdate.Interval = 500
            ElseIf intTime(i) >= 4 Then
                tmrUpdate.Interval = 400
            ElseIf intTime(i) > 6 Then
                tmrUpdate.Interval = 200
            ElseIf intTime(i) > 7 Then
                tmrUpdate.Interval = 100
            End If
        End If
    Next i
End Sub

Private Sub Winsock_Close(Index As Integer)
    Wait (Index)
    AddC Val#(Index), vbRed, LoadResString(106)
    RotateProxies
    If TimeoutP(Index) = LoadResString(130) Then AddC Val#(Index), vbRed, LoadResString(183): Exit Sub
    Waiting(Index) = 1
End Sub

Private Sub Winsock_Connect(Index As Integer)
    Dim i As Byte
    Dim s As String
    Dim splt() As String
    Dim Str As String
    
    If blnConnectionProxied(Index) = True Then
        s = "63.240.202.131"
        If LenB(Settings.Server) > 0 Then s = Settings.Server
        splt = Split(s, ".")
        For i = 0 To UBound(splt)
            Str = Str & Chr$(CStr(splt(i)))
        Next i
        Winsock(Index).SendData Chr$(&H4) & Chr$(&H1) & Chr$(&H17) & Chr$(&HE0) & Str & LoadResString(107) & Chr$(&H0)
    Else
        Winsock(Index).SendData Chr(1)
        Send0x50 (Index)
    End If
End Sub
Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim i As Byte
    Dim lngLen As Long
    Dim strBuffer As String
    Dim strTemp As String
    
    Winsock(Index).GetData strTemp, vbString
    If iSocksified(Index) = False And blnConnectionProxied(Index) = True Then
        i = Asc(Mid$(strTemp, 2, 1))
        If Asc(Mid$(strTemp, 1, 1)) = 0 And Left$(i, 1) = 9 Then
            If Right$(i, 1) = 0 Then
                iSocksified(Index) = True
                AddC Val#(Index), vbGreen, LoadResString(108)
                AddC Val#(Index), vbYellow, LoadResString(109)
                Timeout(Index) = 0
                TimeoutEnabled(Index) = True
                Winsock(Index).SendData Chr$(1)
                Send0x50 (Index)
             Else
                AddC Val#(Index), vbRed, LoadResString(110)
                Wait (Index)
             End If
          End If
       Exit Sub
    End If
    
    strBuffer = strBuffer & strTemp
    While Len(strBuffer) > 4
        lngLen = PBuffer.GetWORD(Mid(strBuffer, 3, 2))
        ParseData (Left(strBuffer, lngLen)), Val#(Index)
        strBuffer = Mid(strBuffer, lngLen + 1)
    Wend
End Sub
Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Select Case Number
        Case 10061:
            AddC Val#(Index), vbRed, LoadResString(111)
            AddC Val#(Index), vbRed, LoadResString(112)
            AddC Val#(Index), vbRed, LoadResString(113) & Proxies(cProxy(Index))
            RemoveProxy (cProxy(Index))
            Waiting(Index) = 1
        Case Else: AddC Val#(Index), vbRed, LoadResString(114) & Number & " " & Description
    End Select
    Wait (Index)
End Sub
