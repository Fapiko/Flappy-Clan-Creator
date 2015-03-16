VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   0  'None
   ClientHeight    =   2895
   ClientLeft      =   4380
   ClientTop       =   4260
   ClientWidth     =   4815
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Flappy Clan Creator > Settings"
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtTimeout 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "10"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Text            =   "63.240.202.139"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtRegEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtHome 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Text            =   "Clan Cell"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtTag 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmbProfiles 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label cmdOK 
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label cmdApply 
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Proxy Timeout"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   1730
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Server"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1730
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Email"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Home"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Clan Tag"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Profile"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Clan Name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   750
         Width           =   855
      End
   End
   Begin VB.Menu mnuGui 
      Caption         =   "-> GUI by Fleet- <-"
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim errEncountered As Boolean
Private Sub cmdApply_Click()
    SaveSettings
End Sub

Private Sub cmdCancel_Click()
    frmSettings.Visible = False
End Sub

Private Sub cmdOK_Click()
    SaveSettings
    If errEncountered <> True Then frmSettings.Visible = False
    errEncountered = False
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    With frmMain.flbProfiles
        For i = 0 To .ListCount - 1
            cmbProfiles.AddItem Replace$(.List(i), LoadResString(115), "")
        Next i
    End With
    cmbProfiles.ListIndex = 0
End Sub
Private Sub strError(Message As String)
    MsgBox Message, vbCritical, LoadResString(116)
    errEncountered = True
End Sub
Private Sub SaveSettings()
    Dim i As Byte
    Dim Account As String
    Dim Key As String
    Dim s() As String
    On Error GoTo errOccured
    
    If Len(txtName) = 0 Then strError (LoadResString(117)): Exit Sub
    If Len(txtTag) <= 1 Then strError (LoadResString(118)): Exit Sub
    
    With Settings
        .ClanName = txtName.Text
        .ClanTag = txtTag.Text
        .HomeChannel = txtHome.Text
        .Profile = cmbProfiles.Text
        .ProxyTimeout = txtTimeout.Text
        .RegEmail = txtRegEmail.Text
        .Server = txtServer.Text
        .Setup = True
    End With
    
    Open App.Path & LoadResString(104) & Settings.Profile & LoadResString(115) For Input As #1
    Open App.Path & LoadResString(119) & Settings.Profile & LoadResString(115) For Input As #2
    
    frmMain.cmbProfiles.Clear
    For i = 0 To 9
        Input #1, Account
        s = Split(Account, " ", 2)
        Accounts.U(i) = s(0)
        Accounts.P(i) = s(1)
        With frmMain.cmbProfiles
            .AddItem Accounts.U(i)
            .ListIndex = 0
        End With
        
        AddC i, vbTeal, LoadResString(120) & Accounts.U(i) & LoadResString(121)
    Next i
    Do Until EOF(2) = True
        Input #2, Key
        If Len(Key) = 26 Then frmMain.lbCDKeys.AddItem Key
    Loop
    
    
    Close #1
    Close #2
    
    With frmMain
        .mnuConnect_Current.Enabled = True
        .mnuConnect_All.Enabled = True
    End With
    
    Exit Sub
errOccured:
    Select Case Err.Number
        
        Case 9: MsgBox LoadResString(122), vbCritical, LoadResString(116)
        Case 53: MsgBox LoadResString(123), vbCritical, LoadResString(116)
        Case 62: MsgBox LoadResString(124), vbCritical, LoadResString(116)
        Case 76: MsgBox LoadResString(125), vbCritical, LoadResString(116)
        Case Else: MsgBox Err.Number & ": " & Err.Description, vbCritical, LoadResString(116)
    End Select
    Close #1
    Close #2
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    formDrag Me
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    formDrag Me
End Sub
