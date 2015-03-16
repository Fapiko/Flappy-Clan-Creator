Attribute VB_Name = "mdlVariables"
Option Explicit
Public AddChief(9) As Boolean
Public blnConnectionProxied(9) As Boolean
Public ClanCreated As Boolean
Public Connected(9) As Boolean
Public Connecting(9) As Boolean
Public TimeoutEnabled(9) As Boolean
Public iSocksified(9) As Boolean
Public ProfileRequest As Boolean
Public RemoveFriend(9) As Boolean
Public intTime(9) As Byte
Public Runs(9) As Byte
Public Waiting(9) As Byte
Public Timeout(9) As Integer
Public lData As Long
Public lngMPQNumber As Long
Public lngNLS(9) As Long
Public bData As String
Public cProxy(9) As String
Public HashCommand As String
Public Key(9) As String
Public mpqName As String
Public Proxies() As String
Public ServerToken As String
Public strCurrentUsername As String
Public TimeoutP(9) As String
Public uProxies() As String

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long

Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

Public PBuffer As New clsPBuffer

Public Accounts As Accounts
Public Settings As Settings

Public MouseDownForm
Public MouseDownFormX
Public MouseDownFormY

Public Const EID_SHOWUSER = &H1
Public Const EID_JOIN = &H2
Public Const EID_LEAVE = &H3
Public Const EID_WHISPER = &H4
Public Const EID_TALK = &H5
Public Const EID_BROADCAST = &H6
Public Const EID_CHANNEL = &H7
Public Const EID_USERFLAGS = &H9
Public Const EID_WHISPERSENT = &HA
Public Const EID_CHANNELFULL = &HD
Public Const EID_CHANNELDOESNOTEXIST = &HE
Public Const EID_CHANNELRESTRICTED = &HF
Public Const EID_INFO = &H12
Public Const EID_ERROR = &H13
Public Const EID_EMOTE = &H17

Public Const SID_NULL = &H0
Public Const SID_ENTERCHAT = &HA
Public Const SID_JOINCHANNEL = &HC
Public Const SID_CHATCOMMAND = &HE
Public Const SID_CHATEVENT = &HF
Public Const SID_PING = &H25
Public Const SID_AUTH_INFO = &H50
Public Const SID_AUTH_CHECK = &H51
Public Const SID_AUTH_ACCOUNTCREATE = &H52
Public Const SID_AUTH_ACCOUNTLOGON = &H53
Public Const SID_AUTH_ACCOUNTLOGONPROOF = &H54
Public Const SID_SETEMAIL = &H59
Public Const SID_FRIENDSLIST = &H65
Public Const SID_FRIENDSUPDATE = &H66
Public Const SID_FRIENDSADD = &H67
Public Const SID_CLANFINDCANDIDATES = &H70
Public Const SID_CLANINVITEMULTIPLE = &H71
Public Const SID_CLANSENDINVITE = &H72
Public Const SID_CLANINFO = &H75
Public Const SID_CLANQUITNOTIFY = &H76
Public Const SID_CLANREMOVEMEMBER = &H78

Public Const vbGreen = 7824686
Public Const vbGrey = 8093055
Public Const vbOrange = 4168661
Public Const vbRBlue = &HFF6961
Public Const vbRed = 3487146
Public Const vbYellow = 4428437
Public Const vbTeal = 8421440
Public Const vbTurquoise = 8747834
Public Const vbWhite = vbBlack

Public Type Accounts
    U(9) As String
    P(9) As String
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type Settings
    ClanName As String
    ClanTag As String
    HomeChannel As String
    Profile As String
    ProxyTimeout As Integer
    RegEmail As String
    Server As String
    Setup As Boolean
End Type
