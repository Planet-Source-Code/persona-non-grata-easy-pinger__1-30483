Attribute VB_Name = "Module1"
'Option Explicit
Public ok As Boolean
Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Public Const IP_NO_RESOURCES = (11000 + 6)
Public Const IP_BAD_OPTION = (11000 + 7)
Public Const IP_HW_ERROR = (11000 + 8)
Public Const IP_PACKET_TOO_BIG = (11000 + 9)
Public Const IP_REQ_TIMED_OUT = (11000 + 10)
Public Const IP_BAD_REQ = (11000 + 11)
Public Const IP_BAD_ROUTE = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Public Const IP_PARAM_PROBLEM = (11000 + 15)
Public Const IP_SOURCE_QUENCH = (11000 + 16)
Public Const IP_OPTION_TOO_BIG = (11000 + 17)
Public Const IP_BAD_DESTINATION = (11000 + 18)
Public Const IP_ADDR_DELETED = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Public Const IP_MTU_CHANGE = (11000 + 21)
Public Const IP_UNLOAD = (11000 + 22)
Public Const IP_ADDR_ADDED = (11000 + 23)
Public Const IP_GENERAL_FAILURE = (11000 + 50)
Public Const MAX_IP_STATUS = 11000 + 50
Public Const IP_PENDING = (11000 + 255)
Public Const PING_TIMEOUT = 255
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Estado_host As String
Public Time_rate As Currency
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

Public Type ICMP_OPTIONS
Ttl As Byte
Tos As Byte
Flags As Byte
OptionsSize As Byte
OptionsData As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS

Public Type ICMP_ECHO_REPLY
Address As Long
status As Long
RoundTripTime As Long
DataSize As Integer
Reserved As Integer
DataPointer As Long
Options As ICMP_OPTIONS
Data As String * 250
End Type


Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
    End Type
Public Type WSADATA
wversion As Integer
wHighVersion As Integer
szDescription(0 To MAX_WSADescription) As Byte
szSystemStatus(0 To MAX_WSASYSStatus) As Byte
wMaxSockets As Integer
wMaxUDPDG As Integer
dwVendorInfo As Long
End Type

Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
(ByVal wVersionRequired As Long, _
lpWSADATA As WSADATA) As Long

Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
(ByVal szHost As String, _
ByVal dwHostLen As Long) As Long

Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
(ByVal szHost As String) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" _
(hpvDest As Any, _
ByVal hpvSource As Long, _
ByVal cbCopy As Long)

Public Declare Function IcmpSendEcho Lib "icmp.dll" _
(ByVal IcmpHandle As Long, _
ByVal DestinationAddress As Long, _
ByVal RequestData As String, _
ByVal RequestSize As Integer, _
ByVal RequestOptions As Long, _
ReplyBuffer As ICMP_ECHO_REPLY, _
ByVal ReplySize As Long, _
ByVal Timeout As Long) As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" _
(ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128
Private Declare Function gethostbyaddr Lib "wsock32" (addr As Long, addrLen As Long, _
    addrType As Long) As Long
Public Function IsIP(ByVal strIP As String) As Boolean
    On Error Resume Next
    Dim t As String: Dim s As String: Dim i As Integer
    s = strIP
    While InStr(s, ".") <> 0
        t = Left(s, InStr(s, ".") - 1)
        If IsNumeric(t) And Val(t) >= 0 And Val(t) <= 255 Then s = Mid(s, InStr(s, ".") + 1) _
    Else Exit Function
        i = i + 1
    Wend
    t = s
    If IsNumeric(t) And InStr(t, ".") = 0 And Len(t) = Len(Trim(Str(Val(t)))) And _
    Val(t) >= 0 And Val(t) <= 255 And strIP <> "255.255.255.255" And i = 3 Then IsIP = True
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function
Public Function MakeIP(strIP As String) As Long
    On Error Resume Next
    Dim lIP As Long
    lIP = Left(strIP, InStr(strIP, ".") - 1)
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256 * 256
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    If strIP < 128 Then
        lIP = lIP + strIP * 256 * 256 * 256
    Else
        lIP = lIP + (strIP - 256) * 256 * 256 * 256
    End If
    MakeIP = lIP
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function
Public Function NameByAddr(strAddr As String) As String
    On Error Resume Next
    Dim nRet As Long
    Dim lIP As Long
    Dim strHost As String * 255: Dim strTemp As String
    Dim hst As HOSTENT
    If IsIP(strAddr) Then
        lIP = MakeIP(strAddr)
        nRet = gethostbyaddr(lIP, 4, 2)
        If nRet <> 0 Then
            RtlMoveMemory hst, nRet, Len(hst)
            RtlMoveMemory ByVal strHost, hst.hName, 255
            strTemp = strHost
            If InStr(strTemp, Chr(10)) <> 0 Then strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
            strTemp = Trim(strTemp)
            NameByAddr = strTemp
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function
Public Function AddrByName(ByVal strHost As String)
    On Error Resume Next
    Dim hostent_addr As Long
    Dim hst As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    If IsIP(strHost) Then
        AddrByName = strHost
        Exit Function
    End If
    hostent_addr = gethostbyname(strHost)
    If hostent_addr = 0 Then
        Exit Function
    End If
    RtlMoveMemory hst, hostent_addr, LenB(hst)
    RtlMoveMemory hostip_addr, hst.hAddrList, 4
    ReDim temp_ip_address(1 To hst.hLength)
    RtlMoveMemory temp_ip_address(1), hostip_addr, hst.hLength
    For i = 1 To hst.hLength
        ip_address = ip_address & temp_ip_address(i) & "."
    Next
    ip_address = Mid(ip_address, 1, Len(ip_address) - 1)
    AddrByName = ip_address
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function
Public Sub IP_Initialize()
    Dim udtWSAData As WSADATA
    If WSAStartup(257, udtWSAData) Then
    End If
End Sub
Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Public Sub Main()
AlwaysOnTop Form1, True
End Sub
Public Function GetStatusCode(status As Long) As String

Dim msg As String

Select Case status
Case IP_SUCCESS: msg = "Online"
Case IP_BUF_TOO_SMALL: msg = "Host buffer is too small"
Case IP_DEST_NET_UNREACHABLE: msg = "Host NET unreachable"
Case IP_DEST_HOST_UNREACHABLE: msg = "Host unreachable"
Case IP_DEST_PROT_UNREACHABLE: msg = "Protocol unreachable"
Case IP_DEST_PORT_UNREACHABLE: msg = "Port unreachable"
Case IP_NO_RESOURCES: msg = "Host with no resources"
Case IP_BAD_OPTION: msg = "Bad option"
Case IP_HW_ERROR: msg = "Hardware error"
Case IP_PACKET_TOO_BIG: msg = "Pachage too big for this host"
Case IP_REQ_TIMED_OUT: msg = "Host didn't answer the request"
Case IP_BAD_REQ: msg = "Bad requirement"
Case IP_BAD_ROUTE: msg = "Bad route"
Case IP_TTL_EXPIRED_TRANSIT: msg = "TTL expired"
Case IP_TTL_EXPIRED_REASSEM: msg = "TTL expired with no reason"
Case IP_PARAM_PROBLEM: msg = "Host with parameters problem"
Case IP_SOURCE_QUENCH: msg = "Master host with trouble"
Case IP_OPTION_TOO_BIG: msg = "Host option is too big"
Case IP_BAD_DESTINATION: msg = "Bad destination"
Case IP_ADDR_DELETED: msg = "Address is deleted"
Case IP_SPEC_MTU_CHANGE: msg = "Specific MTU change in IP"
Case IP_MTU_CHANGE: msg = "General IP MTU change"
Case IP_UNLOAD: msg = "IP not loaded"
Case IP_ADDR_ADDED: msg = "Address added"
Case IP_GENERAL_FAILURE: msg = "General failure"
Case IP_PENDING: msg = "IP pendente"
Case PING_TIMEOUT: msg = "Ping timed out"
Case Else: msg = "Unexpected error"
End Select

GetStatusCode = CStr(status) & " [ " & msg & " ]"
Estado_host = msg
If status = 0 Then
   ok = True
Else
   ok = False
End If
End Function


Public Function HiByte(ByVal wParam As Integer)

HiByte = wParam \ &H100 And &HFF&

End Function


Public Function LoByte(ByVal wParam As Integer)

LoByte = wParam And &HFF&

End Function


Public Function Ping(szAddress As String, echo As ICMP_ECHO_REPLY) As Long

Dim hPort As Long
Dim dwAddress As Long
Dim sDataToSend As String
Dim iOpt As Long

sDataToSend = "Echo This"
dwAddress = AddressStringToLong(szAddress)

Call SocketsInitialize
hPort = IcmpCreateFile()

If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, echo, Len(echo), _
PING_TIMEOUT) Then

Ping = echo.RoundTripTime
Time_rate = Ping / 1000
Else: Ping = echo.status * -1
End If
Dim tst As String
tst = GetStatusCode(echo.status)
Call IcmpCloseHandle(hPort)
Call SocketsCleanup

End Function


Function AddressStringToLong(ByVal tmp As String) As Long

On Error Resume Next

Dim i As Integer
Dim parts(1 To 4) As String

i = 0



While InStr(tmp, ".") > 0
i = i + 1
parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
tmp = Mid(tmp, InStr(tmp, ".") + 1)
Wend

i = i + 1
parts(i) = tmp

If i <> 4 Then
AddressStringToLong = 0
Exit Function
End If

AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
Right("00" & Hex(parts(3)), 2) & _
Right("00" & Hex(parts(2)), 2) & _
Right("00" & Hex(parts(1)), 2))

End Function


Public Function SocketsCleanup() As Boolean

Dim x As Long

x = WSACleanup()

If x <> 0 Then
SocketsCleanup = False
Else
SocketsCleanup = True
End If

End Function


Public Function SocketsInitialize() As Boolean

Dim WSAD As WSADATA
Dim x As Integer
Dim szLoByte As String, szHiByte As String, szBuf As String

x = WSAStartup(WS_VERSION_REQD, WSAD)

If x <> 0 Then
SocketsInitialize = False
Exit Function
End If

If LoByte(WSAD.wversion) < WS_VERSION_MAJOR Or _
(LoByte(WSAD.wversion) = WS_VERSION_MAJOR And _
HiByte(WSAD.wversion) < WS_VERSION_MINOR) Then

szHiByte = Trim$(Str$(HiByte(WSAD.wversion)))
szLoByte = Trim$(Str$(LoByte(WSAD.wversion)))
SocketsInitialize = False
Exit Function

End If

If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
szBuf = "This app requires at least" & _
Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
SocketsInitialize = False
Exit Function
End If

SocketsInitialize = True

End Function



