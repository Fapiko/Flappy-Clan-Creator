VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Buffer As String
Public Function Clear()
    Buffer = ""
End Function
Public Function DebugOutput(ByVal sIn As String) As String
     Dim X1 As Long, Y1 As Long
     Dim iLen As Long, iPos As Long
     Dim sB As String, ST As String
     Dim sOut As String
     Dim Offset As Long, sOffset As String

     iLen = Len(sIn)
     If iLen = 0 Then Exit Function
     sOut = ""
     Offset = 0
     For X1 = 0 To ((iLen - 1) \ 16)
         sOffset = Right$("0000" & Hex(Offset), 4)
         sB = String(48, " ")
         ST = LoadResString(216)
         For Y1 = 1 To 16
             iPos = 16 * X1 + Y1
             If iPos > iLen Then Exit For
             Mid(sB, 3 * (Y1 - 1) + 1, 2) = Right("00" & Hex(Asc(Mid(sIn, iPos, 1))), 2) & " "
             Select Case Asc(Mid(sIn, iPos, 1))
             Case 0, 9, 10, 11, 13
                Mid(ST, Y1, 1) = "."
             Case Else
                 Mid(ST, Y1, 1) = Mid(sIn, iPos, 1)
             End Select
         Next Y1
         If LenB(sOut) > 0 Then
            sOut = sOut & vbNewLine
        End If
         sOut = sOut & sOffset & ":  "
         sOut = sOut & sB & "  " & ST
         Offset = Offset + 16
     Next X1
     DebugOutput = sOut
 End Function
Public Function GetByte(Data As String) As Byte
Dim byt As Byte
byt = 0

Call CopyMemory(byt, ByVal Data, 1)
GetByte = byt
End Function
Public Function GetDWORD(Data As String) As Long
Dim lReturn As Long
    Call CopyMemory(lReturn, ByVal Data, 4)
    GetDWORD = lReturn
End Function
Public Function GetWORD(Data As String) As Long
Dim lReturn As Long
    Call CopyMemory(lReturn, ByVal Data, 2)
    GetWORD = lReturn
End Function
Public Function InsertBYTE(Data As Integer)
    Buffer = Buffer & Chr(Data)
End Function
Public Function InsertDWORD(Data As Long)
    Buffer = Buffer & MakeDWORD(Data)
End Function
Public Function InsertNTString(Data As String)
    Buffer = Buffer & Data & Chr(0)
End Function
Public Function InsertNonNTString(Data As String)
    Buffer = Buffer & Data
End Function
Public Function InsertWORD(Data As Integer)
    Buffer = Buffer & MakeWORD(Data)
End Function
Public Function KillNull(ByVal Text As String) As String
    Dim i As Integer
    i = InStr(1, Text, Chr(0))
    If i = 0 Then
        KillNull = Text
        Exit Function
    End If
    KillNull = Left(Text, i - 1)
End Function
Public Function MakeLONG(X As String) As Long
    If Len(X) < 4 Then
        Exit Function
    End If
    CopyMemory MakeLONG, ByVal X, 4
End Function
Function MakeDWORD(Value As Long) As String
Dim Result As String * 4
    CopyMemory ByVal Result, Value, 4
    MakeDWORD = Result
End Function
Function MakeWORD(Value As Integer) As String
Dim Result As String * 2
    CopyMemory ByVal Result, Value, 2
    MakeWORD = Result
End Function
Public Sub SendPacket(PacketID As Byte, Index As Byte)
    If frmMain.Winsock(Index).State <> sckConnected Then Exit Sub
    frmMain.Winsock(Index).SendData Chr(&HFF) & Chr(PacketID) & MakeWORD(Len(Buffer) + 4) & Buffer
    Clear
End Sub
Public Function SendData(Index As Byte)
    If frmMain.Winsock(Index).State <> sckConnected Then Exit Function
    frmMain.Winsock(Index).SendData Buffer
    Clear
End Function
Public Function StrToHex(ByVal String1 As String) As String
    Dim strTemp As String, strReturn As String, i As Long
    For i = 1 To Len(String1)
        strTemp = Hex(Asc(Mid(String1, i, 1)))
    If Len(strTemp) = 1 Then strTemp = "0" & strTemp
    strReturn = strReturn & " " & strTemp
    Next i
    StrToHex = strReturn
End Function
