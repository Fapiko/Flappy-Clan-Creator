Attribute VB_Name = "mdlBNCSU"
Option Explicit
Public Declare Function bncsutil_getVersion Lib "bncsutil.dll" () As Long
Public Declare Function bncsutil_getVersionString_Raw Lib "bncsutil.dll" Alias "bncsutil_getVersionString" (ByVal outBuf As String) As Long

Public Declare Function checkRevision_Raw Lib "bncsutil.dll" Alias "checkRevisionFlat" (ByVal ValueString As String, ByVal File1 As String, ByVal File2 As String, ByVal File3 As String, ByVal mpqNumber As Long, ByRef Checksum As Long) As Long
Public Declare Function getExeInfo_Raw Lib "bncsutil.dll" Alias "getExeInfo" (ByVal Filename As String, ByVal exeInfoString As String, ByVal infoBufferSize As Long, Version As Long, ByVal Platform As Long) As Long

Public Declare Sub doubleHashPassword_Raw Lib "bncsutil.dll" Alias "doubleHashPassword" (ByVal Password As String, ByVal ClientToken As Long, ByVal ServerToken As Long, ByVal outBuffer As String)
Public Declare Sub hashPassword_Raw Lib "bncsutil.dll" Alias "hashPassword" (ByVal Password As String, ByVal outBuffer As String)

Public Declare Sub calcHashBuf Lib "bncsutil.dll" (ByVal Data As String, ByVal Length As Long, ByVal Hash As String)

Public Declare Function kd_quick Lib "bncsutil.dll" (ByVal CDKey As String, ByVal ClientToken As Long, ByVal ServerToken As Long, PublicValue As Long, Product As Long, ByVal HashBuffer As String, ByVal BufferLen As Long) As Long
Public Declare Function kd_init Lib "bncsutil.dll" () As Long
Public Declare Function kd_create Lib "bncsutil.dll" (ByVal CDKey As String, ByVal keyLength As Long) As Long
Public Declare Function kd_free Lib "bncsutil.dll" (ByVal decoder As Long) As Long
Public Declare Function kd_val2Length Lib "bncsutil.dll" (ByVal decoder As Long) As Long
Public Declare Function kd_product Lib "bncsutil.dll" (ByVal decoder As Long) As Long
Public Declare Function kd_val1 Lib "bncsutil.dll" (ByVal decoder As Long) As Long
Public Declare Function kd_val2 Lib "bncsutil.dll" (ByVal decoder As Long) As Long
Public Declare Function kd_longVal2 Lib "bncsutil.dll" (ByVal decoder As Long, ByVal Out As String) As Long
Public Declare Function kd_calculateHash Lib "bncsutil.dll" (ByVal decoder As Long, ByVal ClientToken As Long, ByVal ServerToken As Long) As Long
Public Declare Function kd_getHash Lib "bncsutil.dll" (ByVal decoder As Long, ByVal Out As String) As Long
Public Declare Function kd_isValid Lib "bncsutil.dll" (ByVal decoder As Long) As Long

Public Declare Function nls_init Lib "bncsutil.dll" (ByVal Username As String, ByVal Password As String) As Long
Public Declare Function nls_init_l Lib "bncsutil.dll" (ByVal Username As String, ByVal Username_Length As Long, ByVal Password As String, ByVal Password_Length As Long) As Long
Public Declare Function nls_reinit Lib "bncsutil.dll" (ByVal NLS As Long, ByVal Username As String, ByVal Password As String) As Long
Public Declare Function nls_reinit_l Lib "bncsutil.dll" (ByVal NLS As Long, ByVal Username As String, ByVal Username_Length As Long, ByVal Password As String, ByVal Password_Length As Long) As Long
Public Declare Sub nls_free Lib "bncsutil.dll" (ByVal NLS As Long)
Public Declare Function nls_account_create Lib "bncsutil.dll" (ByVal NLS As Long, ByVal Buffer As String, ByVal BufLen As Long) As Long
Public Declare Function nls_account_logon Lib "bncsutil.dll" (ByVal NLS As Long, ByVal Buffer As String, ByVal BufLen As Long) As Long
Public Declare Sub nls_get_A Lib "bncsutil.dll" (ByVal NLS As Long, ByVal Out As String)
Public Declare Sub nls_get_M1 Lib "bncsutil.dll" (ByVal NLS As Long, ByVal Out As String, ByVal B As String, ByVal Salt As String)
Public Declare Function nls_check_M2 Lib "bncsutil.dll" (ByVal NLS As Long, ByVal M2 As String, ByVal B As String, ByVal Salt As String) As Long
Public Declare Function nls_check_signature Lib "bncsutil.dll" (ByVal Address As Long, ByVal Signature As String) As Long
Public Declare Function nls_account_change_proof Lib "bncsutil.dll" (ByVal NLS As Long, ByVal Buffer As String, ByVal NewPassword As String, ByVal B As String, ByVal Salt As String) As Long 'returns a new NLS pointer for the new password
Public Declare Sub nls_get_S Lib "bncsutil.dll" (ByVal NLS As Long, ByVal Out As String, ByVal B As String, ByVal Salt As String)
Public Declare Sub nls_get_K Lib "bncsutil.dll" (ByVal NLS As Long, ByVal Out As String, ByVal s As String)
    
'  Constants
'---------------------------
Public Const BNCSUTIL_PLATFORM_X86& = &H1
Public Const BNCSUTIL_PLATFORM_WINDOWS& = &H1
Public Const BNCSUTIL_PLATFORM_WIN& = &H1

Public Const BNCSUTIL_PLATFORM_PPC& = &H2
Public Const BNCSUTIL_PLATFORM_MAC& = &H2

Public Const BNCSUTIL_PLATFORM_OSX& = &H3



'  VB-Specifc Functions
'---------------------------

' RequiredVersion must be a version as a.b.c
' Returns True if the current BNCSutil version is sufficent, False if not.
Public Function bncsutil_checkVersion(ByVal RequiredVersion As String) As Boolean
    Dim i&, j&
    Dim Frag() As String
    Dim Req As Long, Check As Long
    bncsutil_checkVersion = False
    Frag = Split(RequiredVersion, ".")
    j = 0
    For i = UBound(Frag) To 0 Step -1
        Check = Check + (CLng(Val(Frag(i))) * (100 ^ j))
        j = j + 1
    Next i
    Check = bncsutil_getVersion()
    If (Check >= Req) Then
        bncsutil_checkVersion = True
    End If
End Function

Public Function bncsutil_getVersionString() As String
    Dim Str As String
    Str = String$(10, vbNullChar)
    Call bncsutil_getVersionString_Raw(Str)
    bncsutil_getVersionString = Str
End Function

'CheckRevision
Public Function checkRevision(ValueString As String, File1$, File2$, File3$, mpqNumber As Long, Checksum As Long) As Boolean
    checkRevision = (checkRevision_Raw(ValueString, File1, File2, File3, mpqNumber, Checksum) > 0)
End Function

Public Function checkRevisionA(ValueString As String, Files() As String, mpqNumber As Long, Checksum As Long) As Boolean
    checkRevisionA = (checkRevision_Raw(ValueString, Files(0), Files(1), Files(2), mpqNumber, Checksum) > 0)
End Function

Public Function getExeInfo(EXEFile As String, InfoString As String, Optional ByVal Platform As Long = BNCSUTIL_PLATFORM_WINDOWS) As Long
    Dim Version As Long, InfoSize As Long, Result As Long
    Dim i&
    InfoSize = 256
    InfoString = String$(256, vbNullChar)
    Result = getExeInfo_Raw(EXEFile, InfoString, InfoSize, Version, Platform)
    If Result = 0 Then
        getExeInfo = 0
        Exit Function
    End If
    While Result > InfoSize
        If InfoSize > 1024 Then
            getExeInfo = 0
            Exit Function
        End If
        InfoSize = InfoSize + 256
        InfoString = String$(InfoSize, vbNullChar)
        Result = getExeInfo_Raw(EXEFile, InfoString, InfoSize, Version, Platform)
    Wend
    getExeInfo = Version
    i = InStr(InfoString, vbNullChar)
    If i = 0 Then Exit Function
    InfoString = Left$(InfoString, i - 1)
End Function

'OLS Password Hashing
Public Function doubleHashPassword(Password As String, ByVal ClientToken&, ByVal ServerToken&) As String
    Dim Hash As String * 20
    doubleHashPassword_Raw Password, ClientToken, ServerToken, Hash
    doubleHashPassword = Hash
End Function

Public Function hashPassword(Password As String) As String
    Dim Hash As String * 20
    hashPassword_Raw Password, Hash
    hashPassword = Hash
End Function
