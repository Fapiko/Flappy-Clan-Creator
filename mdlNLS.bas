Attribute VB_Name = "mdlNLS"
Public Declare Sub DecodeWar3CDkey Lib "NLS.dll" Alias "#8" (ByVal strCDKey As String, ByVal strPrivate As String, ByRef ptrValue1 As Long, ByRef ptrValue2 As Long)

Private Declare Sub SHA1Init Lib "StandardSHA1.dll" (ByVal ptrContext As Long)
Private Declare Sub SHA1Update Lib "StandardSHA1.dll" (ByVal ptrContext As Long, ByVal ptrData As Long, ByVal lngLen As Long)
Private Declare Sub SHA1Final Lib "StandardSHA1.dll" (ByVal ptrDigest As Long, ByVal ptrContext As Long)

Private Type SHA1_CTX
    lngState(4) As Long
    lngCount(1) As Long
    bytBuffer(63) As Byte
End Type

Private Const SHA_DIGESTSIZE As Byte = 20
Public Sub DecodeHashCDKey(ByVal strCDKey As String, ByVal lngClientKey As Long, ByVal lngServerKey As Long, ByRef lngProdID As Long, ByRef lngValue1 As Long, ByRef strOutBuf As String)
Dim u1 As Long, u2 As String, u3 As Long
Dim bytHashBuffer(25) As Byte
Dim Context As SHA1_CTX
Dim bytHashedData() As Byte
Dim i As Byte
    
    ' u3 = ProdID '
    ' u2 = 10 Byte Buffer '
    ' u1 = Long
    
    u2 = String(10, vbNullChar)
    
    DecodeWar3CDkey UCase$(strCDKey), u2, u3, u1

    RtlMoveMemory ByVal VarPtr(bytHashBuffer(0)), lngClientKey, 4
    RtlMoveMemory ByVal VarPtr(bytHashBuffer(4)), lngServerKey, 4
    RtlMoveMemory ByVal VarPtr(bytHashBuffer(8)), u3, 4
    RtlMoveMemory ByVal VarPtr(bytHashBuffer(12)), u1, 4

    For i = 0 To 9
        RtlMoveMemory ByVal VarPtr(bytHashBuffer(16 + i)), CByte(Asc(Mid$(u2, i + 1, 1))), 1
    Next i
    
    bytHashedData() = SHA1Add(VarPtr(Context), VarPtr(bytHashBuffer(0)), UBound(bytHashBuffer) + 1)
    
    lngProdID = u3
    lngValue1 = u1
    strOutBuf = StrConv(bytHashedData, vbUnicode)
    
End Sub
Public Function SHA1Add(ByVal ptrContext As Long, ByVal ptrData As Long, ByVal lngLen As Long) As Byte()
Dim bytDigest(SHA_DIGESTSIZE - 1) As Byte
    Call SHA1Init(ptrContext)
    Call SHA1Update(ptrContext, ptrData, lngLen)
    Call SHA1Final(VarPtr(bytDigest(0)), ptrContext)
    SHA1Add = bytDigest()
End Function
