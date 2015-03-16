VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDebug 
   Caption         =   "Fg Clan Creator Debug"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfDebug 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6800
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmDebug.frx":0ECA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
With rtfDebug
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = 16777215
        .SelText = LoadResString(101) & App.Major & "." & App.Minor & "." & App.Revision & LoadResString(103) & vbNewLine
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        If Me.Width < 8070 Then Me.Width = 8070
        If Me.Height < 2000 Then Me.Height = 2000
        rtfDebug.Width = Me.Width - 120
        rtfDebug.Height = Me.Height - 405
     End If
     Debug.Print Me.Width
     Debug.Print Me.Height
End Sub
