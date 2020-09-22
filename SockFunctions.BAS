Attribute VB_Name = "mod5"
Option Explicit
Private Const WM_MOUSEMOVE                As Long = &H200
Public Const FD_SETSIZE = 64
'<:-) :WARNING: Untyped Variable will behave as Variant
Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer
End Type
Type timeval
    tv_sec As Long
    tv_usec As Long
End Type
Public Const hostent_size = 16
'<:-) :WARNING: Untyped Variable will behave as Variant
Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Const servent_size = 14
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type
Public Const protoent_size = 10

Public Type sockaddr
    sin_family                              As Integer
    sin_port                                As Integer
    sin_addr                                As Long
    sin_zero                                As String * 8
End Type

'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Type protoent
    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type
Public Const IPPROTO_TCP = 6
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const IPPROTO_UDP = 17
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const INADDR_NONE = &HFFFFFFFF
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const INADDR_ANY = &H0
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const sockaddr_size = 16
Public saZero                            As sockaddr
'<:-) :WARNING: Untyped Variable will behave as Variant

Public Const WSA_DESCRIPTIONLEN = 256
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSA_SYS_STATUS_LEN = 128
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1
'<:-) :WARNING: Untyped Variable will behave as Variant
Type WSADataType
    wversion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Public Const INVALID_SOCKET = -1
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const SOCKET_ERROR = -1
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const SOCK_STREAM = 1
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const SOCK_DGRAM = 2
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const MAXGETHOSTSTRUCT = 1024
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const AF_INET = 2
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const PF_INET = 2
'<:-) :WARNING: Untyped Variable will behave as Variant
Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type
Public Const WSAEINTR = 10004
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEBADF = 10009
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEACCES = 10013
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEFAULT = 10014
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEINVAL = 10022
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEMFILE = 10024
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEWOULDBLOCK = 10035
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEINPROGRESS = 10036
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEALREADY = 10037
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAENOTSOCK = 10038
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEDESTADDRREQ = 10039
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEMSGSIZE = 10040
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEPROTOTYPE = 10041
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAENOPROTOOPT = 10042
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEPROTONOSUPPORT = 10043
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAESOCKTNOSUPPORT = 10044
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEOPNOTSUPP = 10045
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEPFNOSUPPORT = 10046
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEAFNOSUPPORT = 10047
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEADDRINUSE = 10048
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEADDRNOTAVAIL = 10049
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAENETDOWN = 10050
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAENETUNREACH = 10051
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAENETRESET = 10052
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAECONNABORTED = 10053
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAECONNRESET = 10054
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAENOBUFS = 10055
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEISCONN = 10056
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAENOTCONN = 10057
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAESHUTDOWN = 10058
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAETOOMANYREFS = 10059
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAETIMEDOUT = 10060
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAECONNREFUSED = 10061
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAELOOP = 10062
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAENAMETOOLONG = 10063
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEHOSTDOWN = 10064
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEHOSTUNREACH = 10065
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAENOTEMPTY = 10066
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEPROCLIM = 10067
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEUSERS = 10068
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEDQUOT = 10069
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAESTALE = 10070
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAEREMOTE = 10071
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSASYSNOTREADY = 10091
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAVERNOTSUPPORTED = 10092
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSANOTINITIALISED = 10093
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSAHOST_NOT_FOUND = 11001
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSATRY_AGAIN = 11002
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSANO_RECOVERY = 11003
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSANO_DATA = 11004
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const WSANO_ADDRESS = 11004
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const FIONREAD = &H8004667F
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const FIONBIO = &H8004667E
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const FIOASYNC = &H8004667D
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Const SOL_SOCKET = &HFFFF&
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const SO_LINGER = &H80&
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const FD_READ = &H1&
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const FD_WRITE = &H2&
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const FD_OOB = &H4&
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const FD_ACCEPT = &H8&
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const FD_CONNECT = &H10&
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const FD_CLOSE = &H20&
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Declare Function accept Lib "WSOCK32.DLL" (ByVal s As Long, addr As sockaddr, addrLen As Long) As Long
Public Declare Function bind Lib "WSOCK32.DLL" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function closesocket Lib "WSOCK32.DLL" (ByVal s As Long) As Long
Public Declare Function connect Lib "WSOCK32.DLL" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function ioctlsocket Lib "WSOCK32.DLL" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Public Declare Function getpeername Lib "WSOCK32.DLL" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
Public Declare Function getsockname Lib "WSOCK32.DLL" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
Public Declare Function getsockopt Lib "WSOCK32.DLL" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Public Declare Function htonl Lib "WSOCK32.DLL" (ByVal hostlong As Long) As Long
Public Declare Function htons Lib "WSOCK32.DLL" (ByVal hostshort As Long) As Integer
Public Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal inn As Long) As Long
Public Declare Function listen Lib "WSOCK32.DLL" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function ntohl Lib "WSOCK32.DLL" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "WSOCK32.DLL" (ByVal netshort As Long) As Integer
Public Declare Function recv Lib "WSOCK32.DLL" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function recvfrom Lib "WSOCK32.DLL" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long, from As sockaddr, fromlen As Long) As Long
Public Declare Function ws_select Lib "WSOCK32.DLL" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, TimeOut As timeval) As Long
Public Declare Function send Lib "WSOCK32.DLL" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function sendto Lib "WSOCK32.DLL" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
Public Declare Function setsockopt Lib "WSOCK32.DLL" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function ShutDown Lib "WSOCK32.DLL" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function socket Lib "WSOCK32.DLL" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function gethostbyaddr Lib "WSOCK32.DLL" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal host_name As String) As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function getservbyport Lib "WSOCK32.DLL" (ByVal port As Long, ByVal proto As String) As Long
Public Declare Function getservbyname Lib "WSOCK32.DLL" (ByVal serv_name As String, ByVal proto As String) As Long
Public Declare Function getprotobynumber Lib "WSOCK32.DLL" (ByVal proto As Long) As Long
Public Declare Function getprotobyname Lib "WSOCK32.DLL" (ByVal proto_name As String) As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Sub WSASetLastError Lib "WSOCK32.DLL" (ByVal iError As Long)
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAIsBlocking Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAUnhookBlockingHook Lib "WSOCK32.DLL" () As Long
Public Declare Function WSASetBlockingHook Lib "WSOCK32.DLL" (ByVal lpBlockFunc As Long) As Long
Public Declare Function WSACancelBlockingCall Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAAsyncGetServByName Lib "WSOCK32.DLL" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetServByPort Lib "WSOCK32.DLL" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetProtoByName Lib "WSOCK32.DLL" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetProtoByNumber Lib "WSOCK32.DLL" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal number As Long, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetHostByName Lib "WSOCK32.DLL" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetHostByAddr Lib "WSOCK32.DLL" (ByVal hwnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSACancelAsyncRequest Lib "WSOCK32.DLL" (ByVal hAsyncTaskHandle As Long) As Long
Public Declare Function WSAAsyncSelect Lib "WSOCK32.DLL" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSARecvEx Lib "WSOCK32.DLL" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

'Public Declare Function accept Lib "WSOCK32.DLL" (ByVal s As Long, addr As sockaddr, addrLen As Long) As Long
'Public Declare Function bind Lib "WSOCK32.DLL" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
'Public Declare Function closesocket Lib "WSOCK32.DLL" (ByVal s As Long) As Long
'Public Declare Function connect Lib "WSOCK32.DLL" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
'Public Declare Function ioctlsocket Lib "WSOCK32.DLL" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
'Public Declare Function getpeername Lib "WSOCK32.DLL" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long

'Public MySocket As Integer
'<:-) :SUGGESTION: Scope should be changed to Public
Public SockReadBuffer As String
Public Const WSA_NoName = "Unknown"
'<:-) :WARNING: Untyped Variable will behave as Variant
Public WSAStartedUp As Boolean
Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type
Public Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type
Public Const PING_TIMEOUT = 200
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Declare Function IcmpCreateFile Lib "Icmp.dll" () As Long
'<:-) :SUGGESTION: Scope should be changed to Public
Public Declare Function IcmpSendEcho Lib "Icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
'<:-) :SUGGESTION: Scope should be changed to Public
Public Declare Function IcmpCloseHandle Lib "Icmp.dll" (ByVal IcmpHandle As Long) As Long
'<:-) :SUGGESTION: Scope should be changed to Public
Public Type IP_ADDRESS_STRING
    IpAddressString(4 * 4 - 1) As Byte
End Type
Public Type IP_MASK_STRING
    IpMaskString(4 * 4 - 1) As Byte
End Type
Public Type IP_ADDR_STRING
Next      As Long
IPAddress As IP_ADDRESS_STRING
IpMask    As IP_MASK_STRING
Context   As Long
End Type
Public Const MAX_HOSTNAME_LEN = 128
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const MAX_DOMAIN_NAME_LEN = 128
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Const MAX_SCOPE_ID_LEN = 256
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Type FIXED_INFO
    hostname(MAX_HOSTNAME_LEN + 4 - 1) As Byte
    DomainName(MAX_DOMAIN_NAME_LEN + 4 - 1) As Byte
    CurrentDnsServer As Long
    DnsServerList    As IP_ADDR_STRING
    NodeType         As Long
    ScopeId(MAX_SCOPE_ID_LEN + 4 - 1) As Byte
    EnableRouting    As Long
    EnableProxy      As Long
    EnableDns        As Long
End Type
Public w_connect As Long, w_listen As Long
Public Const ERROR_BUFFER_OVERFLOW = 111
'<:-) :SUGGESTION: Scope should be changed to Public
'<:-) :WARNING: Untyped Variable will behave as Variant
Public Declare Function GetNetworkParams Lib "iphlpapi.dll" (pFixedInfo As Any, pOutBufLen As Long) As Long
'<:-) :SUGGESTION: Scope should be changed to Public
Declare Function timeGetTime Lib "winmm.dll" () As Long
'<:-) :SUGGESTION: Scope should be changed to Public Declare
Public Function AddrToIP(ByVal AddrOrIP As String) As String
    On Error Resume Next
    AddrToIP = GetAscIP(GetHostByNameAlias(AddrOrIP))
    If err Then
        AddrToIP = "255.255.255.255"
    End If
    On Error GoTo 0
End Function
Public Sub CloseSock(Optional Sockid As Long)
    'Listening = False
    If Sockid = 0 Then
        closesocket w_connect
        closesocket w_listen
        'Sstate = 0
    Else
        closesocket Sockid
    End If
End Sub
Public Function ConnectSock(ByVal host As String, ByVal IntPort As Long, ByVal retIpPort As String, ByVal HWndToMsg, ByVal Async As Long) As Long
Dim s         As Long
Dim SelectOps As Long
Dim sockin    As sockaddr
    SockReadBuffer = ""
    sockin = saZero
    With sockin
        .sin_family = AF_INET
        .sin_port = htons(IntPort)
        If .sin_port = INVALID_SOCKET Then
            ConnectSock = INVALID_SOCKET
            Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
        End If
    End With
    sockin.sin_addr = GetHostByNameAlias(host)
    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    'SRemoteHost = GetAscIP(sockin.sin_addr)
    'SRemotePort = ntohs(sockin.sin_port)
    s = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    If SetSockLinger(s, 1, 1) = SOCKET_ERROR Then
        If s > 0 Then
            closesocket s
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    If Not Async Then
        If connect(s, sockin, sockaddr_size) <> 0 Then
            If s > 0 Then
                closesocket s
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
        End If
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(s, HWndToMsg, ByVal WM_MOUSEMOVE, ByVal SelectOps) Then
            If s > 0 Then
                closesocket s
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
        End If
    Else
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(s, HWndToMsg, ByVal WM_MOUSEMOVE, ByVal SelectOps) Then
            If s > 0 Then
                closesocket s
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
        End If
        If connect(s, sockin, sockaddr_size) <> -1 Then
            If s > 0 Then
                closesocket s
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
        End If
    End If
    ConnectSock = s
End Function
Public Function GetAscIP(ByVal inn As Long) As String
Dim lpStr     As Long
Dim nStr      As Long
Dim retString As String
    On Error Resume Next
    retString = String$(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        GetAscIP = "255.255.255.255"
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then
        nStr = 32
    End If
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left$(retString, nStr)
    GetAscIP = retString
    If err Then
        GetAscIP = "255.255.255.255"
    End If
    On Error GoTo 0
End Function
Public Function SendData(ByVal s, vMessage As Variant) As Long
Dim TheMsg() As Byte
Dim sTemp    As String
    'Debug.Print vbCrLf & vMessage & vbCrLf
    TheMsg = ""
    Select Case VarType(vMessage)
    Case 8209
        sTemp = vMessage
        TheMsg = sTemp
    Case 8
        sTemp = StrConv(vMessage, vbFromUnicode)
    Case Else
        sTemp = CStr(vMessage)
        sTemp = StrConv(vMessage, vbFromUnicode)
    End Select
    TheMsg = sTemp
    If UBound(TheMsg) > -1 Then
        SendData = send(s, TheMsg(0), (UBound(TheMsg) - LBound(TheMsg) + 1), 0)
    End If
End Function

Public Function GetHostByAddress(ByVal addr As Long) As String
Dim phe        As Long
Dim heDestHost As HOSTENT
Dim hostname   As String
    On Error Resume Next
    phe = gethostbyaddr(addr, 4, PF_INET)
    If phe <> 0 Then
        MemCopy heDestHost, ByVal phe, hostent_size
        hostname = String$(256, 0)
        MemCopy ByVal hostname, ByVal heDestHost.h_name, 256
        GetHostByAddress = Left$(hostname, InStr(hostname, vbNullChar) - 1)
    Else
        GetHostByAddress = WSA_NoName
    End If
    If err Then
        GetHostByAddress = WSA_NoName
    End If
    On Error GoTo 0
End Function
Public Function GetHostByNameAlias(ByVal hostname As String) As Long
Dim phe        As Long
Dim heDestHost As HOSTENT
Dim addrList   As Long
Dim retIP      As Long
    On Error Resume Next
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            MemCopy heDestHost, ByVal phe, hostent_size
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If err Then
        GetHostByNameAlias = INADDR_NONE
    End If
    On Error GoTo 0
End Function
Public Function GetLocalHostName() As String
Dim dummy     As Long
Dim LocalName As String
Dim s         As String
    On Error Resume Next
    LocalName = String$(256, 0)
    LocalName = WSA_NoName
    dummy = 1
    s = String$(256, 0)
    dummy = gethostname(s, 256)
    If dummy = 0 Then
        s = Left$(s, InStr(s, vbNullChar) - 1)
        If Len(s) > 0 Then
            LocalName = s
        End If
    End If
    GetLocalHostName = LocalName
    If err Then
        GetLocalHostName = WSA_NoName
    End If
    On Error GoTo 0
End Function
Public Function IpToAddr(ByVal AddrOrIP As String) As String
    On Error Resume Next
    IpToAddr = GetHostByAddress(GetHostByNameAlias(AddrOrIP))
    If err Then
        IpToAddr = WSA_NoName
    End If
    On Error GoTo 0
End Function
Public Function ListenForConnect(ByVal IntPort, ByVal HWndToMsg) As Long
Dim s         As Long
Dim SelectOps As Long
Dim sockin    As sockaddr
    sockin = saZero
    With sockin
        .sin_family = AF_INET
        .sin_port = htons(IntPort)
        If .sin_port = INVALID_SOCKET Then
            ListenForConnect = INVALID_SOCKET
'Actives = False
            Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
        End If
    End With
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
'Actives = False
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    s = socket(PF_INET, SOCK_STREAM, 0)
    If s < 0 Then
        ListenForConnect = INVALID_SOCKET
'Actives = False
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    If bind(s, sockin, sockaddr_size) Then
        If s > 0 Then
            closesocket s
        End If
        ListenForConnect = INVALID_SOCKET
'Actives = False
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, HWndToMsg, ByVal WM_MOUSEMOVE, ByVal SelectOps) Then
        If s > 0 Then
            closesocket s
        End If
        ListenForConnect = SOCKET_ERROR
'Actives = False
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    If listen(s, 1) Then
        If s > 0 Then
            closesocket s
        End If
        ListenForConnect = INVALID_SOCKET
'Actives = False
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    ListenForConnect = s
    If Not s = 0 Then
'Actives = True
    End If
End Function
Public Function WSAGetAsyncError(ByVal lParam As Long) As Integer
    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
End Function

Public Sub EndWinsock()
Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub
Public Function GetDNSInfo() As String
Dim sFinalBuff            As String
Dim lngFixedInfoNeeded    As Long
Dim bytFixedInfoBuffer()  As Byte
Dim udtFixedInfo          As FIXED_INFO
Dim lngIpAddrStringPtr    As Long
Dim udtIpAddrString       As IP_ADDR_STRING
Dim strDnsIpAddress       As String
Dim lngWin32apiResultCode As Long
Dim lngDNSPing            As Long
Dim lngDNSPing2           As Long
Dim arrDNS                As Variant
Dim Member                As Variant
    lngWin32apiResultCode = GetNetworkParams(ByVal vbNullString, lngFixedInfoNeeded)
    If lngWin32apiResultCode = ERROR_BUFFER_OVERFLOW Then
        ReDim bytFixedInfoBuffer(lngFixedInfoNeeded)
    Else
        GoTo TerminateGetNetworkParams
    End If
    lngWin32apiResultCode = GetNetworkParams(bytFixedInfoBuffer(0), lngFixedInfoNeeded)
    MemCopy udtFixedInfo, bytFixedInfoBuffer(0), Len(udtFixedInfo)
    With udtFixedInfo
        lngIpAddrStringPtr = VarPtr(.DnsServerList)
        Do While lngIpAddrStringPtr
            MemCopy udtIpAddrString, ByVal lngIpAddrStringPtr, Len(udtIpAddrString)
            With udtIpAddrString
                strDnsIpAddress = StrConv(.IPAddress.IpAddressString, vbUnicode)
                If sFinalBuff = vbNullString Then
                    sFinalBuff = Left$(strDnsIpAddress, InStr(strDnsIpAddress, vbNullChar) - 1) & ","
                Else
                    If InStr(1, sFinalBuff, Left$(strDnsIpAddress, InStr(strDnsIpAddress, vbNullChar) - 1) & ",") = 0 Then
                        sFinalBuff = sFinalBuff & Left$(strDnsIpAddress, InStr(strDnsIpAddress, vbNullChar) - 1) & ","
                    End If
                End If
                lngIpAddrStringPtr = .Next
            End With
        Loop
    End With
    If Right$(sFinalBuff, 1) = "," Then sFinalBuff = Left$(sFinalBuff, Len(sFinalBuff) - 1)
    arrDNS = Split(sFinalBuff, ",")
    lngDNSPing = Ping(arrDNS(0), , True)
    sFinalBuff = arrDNS(0)
    For Each Member In arrDNS
        lngDNSPing2 = Ping(Member, , True)
        If (lngDNSPing2 < lngDNSPing And lngDNSPing2 > -1) Or lngDNSPing = -1 Then sFinalBuff = Member: lngDNSPing = lngDNSPing2
    Next Member
    GetDNSInfo = sFinalBuff
TerminateGetNetworkParams:
End Function
Public Function GetPeerAddress(ByVal s&) As String
Dim addrLen&
Dim sa As sockaddr
    addrLen = sockaddr_size
    If getpeername(s, sa, addrLen) Then
        GetPeerAddress = ""
    Else
        GetPeerAddress = SockAddressToString(sa)
    End If
End Function
Public Function GetPortFromString(ByVal PortStr$) As Long
    If Val(PortStr$) > 32767 Then
        GetPortFromString = CInt(Val(PortStr$) - &H10000)
    Else
        GetPortFromString = Val(PortStr$)
    End If
    If err Then GetPortFromString = 0
End Function
Public Function GetProtocolByName(ByVal protocol$) As Long
Dim tmpShort&
Dim ppe&
Dim peDestProt As protoent
    ppe = getprotobyname(protocol)
    If ppe Then
        MemCopy peDestProt, ByVal ppe, protoent_size
        GetProtocolByName = peDestProt.p_proto
    Else
        tmpShort = Val(protocol)
        If tmpShort Then
            GetProtocolByName = htons(tmpShort)
        Else
            GetProtocolByName = SOCKET_ERROR
        End If
    End If
End Function
Public Function GetServiceByName(ByVal service$, ByVal protocol$) As Long
Dim serv&
Dim pse&
Dim seDestServ As servent
    pse = getservbyname(service, protocol)
    If pse Then
        MemCopy seDestServ, ByVal pse, servent_size
        GetServiceByName = seDestServ.s_port
    Else
        serv = Val(service)
        If serv Then
            GetServiceByName = htons(serv)
        Else
            GetServiceByName = INVALID_SOCKET
        End If
    End If
End Function
Public Function GetSockAddress(ByVal s&) As String
Dim addrLen&
Dim sa As sockaddr
Dim szRet$
    szRet = String$(32, 0)
    addrLen = sockaddr_size
    If getsockname(s, sa, addrLen) Then
        GetSockAddress = ""
    Else
        GetSockAddress = SockAddressToString(sa)
    End If
End Function
Public Function GetWSAErrorString(ByVal errnum&) As String
''on error Resume Next
    Select Case errnum
    Case 10004
        GetWSAErrorString = "Interrupted system call."
    Case 10009
        GetWSAErrorString = "Bad file number."
    Case 10013
        GetWSAErrorString = "Permission Denied."
    Case 10014
        GetWSAErrorString = "Bad Address."
    Case 10022
        GetWSAErrorString = "Invalid Argument."
    Case 10024
        GetWSAErrorString = "Too many open files."
    Case 10035
        GetWSAErrorString = "Operation would block."
    Case 10036
        GetWSAErrorString = "Operation now in progress."
    Case 10037
        GetWSAErrorString = "Operation already in progress."
    Case 10038
        GetWSAErrorString = "Socket operation on nonsocket."
    Case 10039
        GetWSAErrorString = "Destination address required."
    Case 10040
        GetWSAErrorString = "Message too long."
    Case 10041
        GetWSAErrorString = "Protocol wrong type for socket."
    Case 10042
        GetWSAErrorString = "Protocol not available."
    Case 10043
        GetWSAErrorString = "Protocol not supported."
    Case 10044
        GetWSAErrorString = "Socket type not supported."
    Case 10045
        GetWSAErrorString = "Operation not supported on socket."
    Case 10046
        GetWSAErrorString = "Protocol family not supported."
    Case 10047
        GetWSAErrorString = "Address family not supported by protocol family."
    Case 10048
        GetWSAErrorString = "Address already in use."
    Case 10049
        GetWSAErrorString = "Can"
    Case 10050
        GetWSAErrorString = "Network is down."
    Case 10051
        GetWSAErrorString = "Network is unreachable."
    Case 10052
        GetWSAErrorString = "Network dropped connection."
    Case 10053
        GetWSAErrorString = "Software caused connection abort."
    Case 10054
        GetWSAErrorString = "Connection reset by peer."
    Case 10055
        GetWSAErrorString = "No buffer space available."
    Case 10056
        GetWSAErrorString = "Socket is already connected."
    Case 10057
        GetWSAErrorString = "Socket is not connected."
    Case 10058
        GetWSAErrorString = "Can"
    Case 10059
        GetWSAErrorString = "Too many references: can"
    Case 10060
        GetWSAErrorString = "Connection timed out."
    Case 10061
        GetWSAErrorString = "Connection refused."
    Case 10062
        GetWSAErrorString = "Too many levels of symbolic links."
    Case 10063
        GetWSAErrorString = "File name too long."
    Case 10064
        GetWSAErrorString = "Host is down."
    Case 10065
        GetWSAErrorString = "No route to host."
    Case 10066
        GetWSAErrorString = "Directory not empty."
    Case 10067
        GetWSAErrorString = "Too many processes."
    Case 10068
        GetWSAErrorString = "Too many users."
    Case 10069
        GetWSAErrorString = "Disk quota exceeded."
    Case 10070
        GetWSAErrorString = "Stale NFS file handle."
    Case 10071
        GetWSAErrorString = "Too many levels of remote in path."
    Case 10091
        GetWSAErrorString = "Network subsystem is unusable."
    Case 10092
        GetWSAErrorString = "Winsock DLL cannot support this application."
    Case 10093
        GetWSAErrorString = "Winsock not initialized."
    Case 10101
        GetWSAErrorString = "Disconnect."
    Case 11001
        GetWSAErrorString = "Host not found."
    Case 11002
        GetWSAErrorString = "Nonauthoritative host not found."
    Case 11003
        GetWSAErrorString = "Nonrecoverable error."
    Case 11004
        GetWSAErrorString = "Valid name, no data record of requested type."
    Case Else
    End Select
End Function
Public Function HTTPRequest(ByVal StrUrl As String, strMethod As String, HWndToMsg As Long, Optional strProxy As String, Optional strPostCommand As String) As Long
''on error GoTo Problems
Dim Sock    As Integer
Dim strhost As String
Dim IntPort As Integer
Dim intsep  As Integer
Dim StrMsg  As String
    If StrComp(Left$(StrUrl, 4), "http", vbTextCompare) = 0 And strProxy = "" Then
        StrUrl = Mid$(StrUrl, InStr(5, StrUrl, "//") + 2)
    Else
        If strProxy <> "" Then
            If StrComp(Left$(StrUrl, 4), "http", vbTextCompare) <> 0 Then StrUrl = "http://" & StrUrl
        End If
    End If
    If strProxy <> "" Then
        strhost = Mid$(strProxy, 1, InStr(1, strProxy, ":") - 1)
        IntPort = Mid$(strProxy, InStr(1, strProxy, ":") + 1)
        If StrComp(Left$(StrUrl, 4), "http", vbTextCompare) = 0 Then
            If InStr(9, StrUrl, "/") = 0 Then StrUrl = StrUrl & "/"
        End If
    Else
'intDot = InStr(1, StrUrl, ":")
'If intDot > 0 Then
'intSep = InStr(intDot + 1, StrUrl, "/")
'If intSep = 0 Then intSep = Len(StrUrl) + 1
'strhost = Mid$(StrUrl, 1, intDot - 1)
'IntPort = Mid$(StrUrl, intDot + 1, intSep - intDot - 1)
'Else
        intsep = InStr(1, StrUrl, "/")
        If intsep = 0 Then intsep = Len(StrUrl) + 1
        strhost = Mid$(StrUrl, 1, intsep - 1)
        IntPort = 80
    End If
    StrUrl = Mid$(StrUrl, intsep)
    If StrUrl = "" Then StrUrl = "/"
'End If
    Sock = ConnectSock(strhost, IntPort, 0, HWndToMsg, 1)
    If Sock = SOCKET_ERROR Then
        HTTPRequest = SOCKET_ERROR
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
    End If
    StrMsg = strMethod & " " & StrUrl & " HTTP/1.0" & vbCrLf
    StrMsg = StrMsg & "Accept: */*" & vbCrLf
    StrMsg = StrMsg & "User-Agent: " & "MSIE6.0" & vbCrLf
    StrMsg = StrMsg & "Host: " & strhost & vbCrLf
    If strMethod = "POST" Then
        StrMsg = StrMsg & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
        StrMsg = StrMsg & "Content-Length: " & Len(strPostCommand) & vbCrLf
        StrMsg = StrMsg & vbCrLf & strPostCommand
    Else
        StrMsg = StrMsg & vbCrLf
    End If
    SendData Sock, StrMsg
    HTTPRequest = Sock
Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
Problems:
    HTTPRequest = SOCKET_ERROR
    Call EndWinsock
End Function
Public Function IrcGetAscIp(ByVal IPL$) As String
''on error GoTo IrcGetAscIPError
Dim lpStr&
Dim nStr&
Dim retString$
Dim inn&
    If Val(IPL) > 2147483647 Then
        inn = Val(IPL) - 4294967296#
    Else
        inn = Val(IPL)
    End If
    inn = ntohl(inn)
    retString = String$(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        IrcGetAscIp = "0.0.0.0"
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left$(retString, nStr)
    IrcGetAscIp = retString
Exit Function
IrcGetAscIPError:
    IrcGetAscIp = "0.0.0.0"
End Function
Public Function IrcGetLongIp(ByVal AscIp$) As String
''on error GoTo IrcGetLongIpError
Dim inn&
    inn = inet_addr(AscIp)
    inn = htonl(inn)
    If inn < 0 Then
        IrcGetLongIp = CVar(inn + 4294967296#)
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 3)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/Else/Rest_of_Code/End If' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    Else
        IrcGetLongIp = CVar(inn)
        Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    End If
Exit Function
IrcGetLongIpError:
    IrcGetLongIp = "0"
End Function
Public Function IsConnected() As Boolean
    IsConnected = InternetGetConnectedState(0&, 0&)
End Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow. (Fix ID 2)
'<:-) Convert 'If..Then/Code(with Explicit Exit)/End If/Rest_of_Code' to
'<:-) 'If..Then/Exit Code(without Explicit Exit)/Else/ Rest_Of_Code/End If
'<:-) OR if Exit Code block is only the Exit Command
'<:-) 'If Not ..Then/ Rest_Of_Code/End If
    Public Function Ping(ByVal hostnameOrIpaddress As String, Optional timeOutmSec As Long = PING_TIMEOUT, Optional bolReturnTime As Boolean)
Dim echoValues  As ICMP_ECHO_REPLY
Dim hPort       As Long
Dim dwAddress   As Long
Dim sDataToSend As String
''on error GoTo e_Trap
    If Trim$(hostnameOrIpaddress) = "" Then GoTo e_Trap
    Call StartWinsock
    sDataToSend = "Echo This"
    hostnameOrIpaddress = AddrToIP(hostnameOrIpaddress)
    dwAddress = GetHostByNameAlias(hostnameOrIpaddress)
    If dwAddress = -1 Then GoTo e_Trap
    hPort = IcmpCreateFile()
    Call IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, echoValues, Len(echoValues), timeOutmSec)
    Call IcmpCloseHandle(hPort)
    If echoValues.Status = 0 And sDataToSend = Left$(echoValues.Data, Len(sDataToSend)) Then
        If bolReturnTime Then
            Ping = echoValues.RoundTripTime
        Else
            If echoValues.RoundTripTime <= timeOutmSec Then Ping = True
        End If
    Else
        If bolReturnTime Then
            Ping = SOCKET_ERROR
        Else
            Ping = False
        End If
    End If
    Call EndWinsock
Exit Function
'<:-) :SUGGESTION: (EXPERIMENTAL follow advice with care )
'<:-) Explict 'Exit ProcedureType' can make code flow harder to follow.(Fix ID 11)
'<:-) No recommended action but consider coding around it.
e_Trap:
    If bolReturnTime Then
        Ping = SOCKET_ERROR
    Else
        Ping = False
    End If
    Call EndWinsock
End Function
Public Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long
Dim Linger As LingerType
    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime
    If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
        SetSockLinger = SOCKET_ERROR
    Else
        If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
            SetSockLinger = SOCKET_ERROR
        Else
        End If
    End If
End Function
Public Function SockAddressToString(sa As sockaddr) As String
    SockAddressToString = GetAscIP(sa.sin_addr) & ":" & ntohs(sa.sin_port)
End Function
Public Function StartWinsock(Optional sDescription As String) As Boolean
Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function
Public Function WSAGetAsyncBufLen(ByVal lParam As Long) As Long
    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetAsyncBufLen = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetAsyncBufLen = lParam And &HFFFF&
    End If
End Function
Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetSelectEvent = lParam And &HFFFF&
    End If
End Function
Public Function WSAMakeSelectReply(TheEvent%, TheError%) As Long
    WSAMakeSelectReply = (TheError * &H10000) + (TheEvent And &HFFFF&)
End Function
':)Code Fixer V3.0.9 (11/15/2006 12:12:19 PM) 392 + 844 = 1236 Lines Thanks Ulli for inspiration and lots of code.
