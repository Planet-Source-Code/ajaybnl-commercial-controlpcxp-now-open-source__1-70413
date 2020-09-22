VERSION 5.00
Begin VB.UserControl Socket 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   ClipControls    =   0   'False
   Enabled         =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   MaskColor       =   &H00C0C0C0&
   MaskPicture     =   "Socket.ctx":0000
   PaletteMode     =   4  'None
   Picture         =   "Socket.ctx":0242
   ScaleHeight     =   405
   ScaleWidth      =   420
   ToolboxBitmap   =   "Socket.ctx":0484
   Begin VB.PictureBox pcAjf 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Warning : This Program Control and its required files are Open Source to the world thus the Author is not responsible For Any Misuse.
Option Explicit
Private errnum                            As Long
Private ErrDescription                    As String
Private WSock                             As Long
Private WSock1                            As Long
Private GotIP As String
Public Event Error(ByVal number As Long, Description As String)
Public Event ConnectionRequest()
Public Event DataArrival(ByVal Data As String)
Public Event Connected()
Public Event Closed()
Private SPort                             As Long
Private SState                            As Long
Private SLocalPort                        As String
Private SRemoteHost                       As String
Private SRemotePort                       As Long
Private Const FD_SETSIZE                  As Long = 64
Private Type fd_set
    fd_count                                As Integer
    fd_array(FD_SETSIZE)                    As Integer
End Type
Private Type timeval
    tv_sec                                  As Long
    tv_usec                                 As Long
End Type
Private Type HOSTENT
    h_name                                  As Long
    h_aliases                               As Long
    h_addrtype                              As Integer
    h_length                                As Integer
    h_addr_list                             As Long
End Type
Private Const hostent_size                As Long = 16
Private Type protoent
    p_name                                  As Long
    p_aliases                               As Long
    p_proto                                 As Integer
End Type
Private Const IPPROTO_TCP                 As Long = 6
Private Const INADDR_NONE                 As Long = &HFFFF
Private Const INADDR_ANY                  As Long = &H0
Private Type sockaddr
    sin_family                              As Integer
    sin_port                                As Integer
    sin_addr                                As Long
    sin_zero                                As String * 8
End Type
Private Const sockaddr_size               As Long = 16
Private saZero                            As sockaddr
Private Const WSA_DESCRIPTIONLEN          As Long = 256
Private Const WSA_DescriptionSize         As Double = WSA_DESCRIPTIONLEN + 1
Private Const WSA_SYS_STATUS_LEN          As Long = 128
Private Const WSA_SysStatusSize           As Double = WSA_SYS_STATUS_LEN + 1
Private Type WSADataType
    wversion                                As Integer
    wHighVersion                            As Integer
    szDescription                           As String * WSA_DescriptionSize
    szSystemStatus                          As String * WSA_SysStatusSize
    iMaxSockets                             As Integer
    iMaxUdpDg                               As Integer
    lpVendorInfo                            As Long
End Type
Private Const INVALID_SOCKET              As Long = -1
Private Const SOCKET_ERROR                As Long = -1
Private Const SOCK_STREAM                 As Long = 1
Private Const AF_INET                     As Long = 2
Private Const PF_INET                     As Long = 2
Private Const FD_READ                     As Long = &H1
Private Const FD_WRITE                    As Long = &H2
Private Const FD_ACCEPT                   As Long = &H8
Private Const FD_CONNECT                  As Long = &H10
Private Const FD_CLOSE                    As Long = &H20
Private Const WM_MOUSEMOVE                As Long = &H200
Private SockReadBuffer                    As String
Private Const WSA_NoName                  As String = "Unknown"
Private WSAStartedUp                      As Boolean
Private Timein                            As Boolean
Private Connin                            As Boolean
Private ActiveS                           As Boolean
Private IdeCheck                             As Boolean
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, _
                                                                  Src As Any, _
                                                                  ByVal cb As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long
Private Declare Function accept Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                   addr As sockaddr, _
                                                   addrLen As Long) As Long
Private Declare Function bind Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                 addr As sockaddr, _
                                                 ByVal namelen As Long) As Long
Private Declare Function closesocket Lib "WSOCK32.DLL" (ByVal s As Long) As Long
Private Declare Function connect Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                    addr As sockaddr, _
                                                    ByVal namelen As Long) As Long
Private Declare Function ioctlsocket Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                        ByVal cmd As Long, _
                                                        argp As Long) As Long
Private Declare Function getpeername Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                        sName As sockaddr, _
                                                        namelen As Long) As Long
Private Declare Function getsockname Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                        sName As sockaddr, _
                                                        namelen As Long) As Long
Private Declare Function getsockopt Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                       ByVal level As Long, _
                                                       ByVal optname As Long, _
                                                       optval As Any, _
                                                       optlen As Long) As Long
Private Declare Function htonl Lib "WSOCK32.DLL" (ByVal hostlong As Long) As Long
Private Declare Function htons Lib "WSOCK32.DLL" (ByVal hostshort As Long) As Integer
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long
Private Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal inn As Long) As Long
Private Declare Function listen Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                   ByVal backlog As Long) As Long
Private Declare Function ntohl Lib "WSOCK32.DLL" (ByVal netlong As Long) As Long
Private Declare Function ntohs Lib "WSOCK32.DLL" (ByVal netshort As Long) As Integer
Private Declare Function recv Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                 buf As Any, _
                                                 ByVal buflen As Long, _
                                                 ByVal Flags As Long) As Long
Private Declare Function recvfrom Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                     buf As Any, _
                                                     ByVal buflen As Long, _
                                                     ByVal Flags As Long, _
                                                     from As sockaddr, _
                                                     fromlen As Long) As Long
Private Declare Function ws_select Lib "WSOCK32.DLL" Alias "select" (ByVal nfds As Long, _
                                                                     readfds As fd_set, _
                                                                     writefds As fd_set, _
                                                                     exceptfds As fd_set, _
                                                                     TimeOut As timeval) As Long
Private Declare Function send Lib "WSOCK32.DLL" (ByVal s As Long, buf As _
                                                              Any, ByVal buflen _
                                                              As Long, ByVal Flags _
                                                              As Long) As Long
Private Declare Function sendto Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                   buf As Any, _
                                                   ByVal buflen As Long, _
                                                   ByVal Flags As Long, _
                                                   to_addr As sockaddr, _
                                                   ByVal tolen As Long) As Long
Private Declare Function setsockopt Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                       ByVal level As Long, _
                                                       ByVal optname As Long, _
                                                       optval As Any, _
                                                       ByVal optlen As Long) As Long
Private Declare Function ShutDown Lib "WSOCK32.DLL" Alias "shutdown" (ByVal s As Long, _
                                                                      ByVal how As Long) As Long
Private Declare Function socket Lib "WSOCK32.DLL" (ByVal af As Long, _
                                                   ByVal s_type As Long, _
                                                   ByVal protocol As Long) As Long
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" (addr As Long, _
                                                          ByVal addr_len As Long, _
                                                          ByVal addr_type As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal host_name As String) As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal host_name As String, _
                                                        ByVal namelen As Long) As Long
Private Declare Function getservbyport Lib "WSOCK32.DLL" (ByVal port As Long, _
                                                          ByVal proto As String) As Long
Private Declare Function getservbyname Lib "WSOCK32.DLL" (ByVal serv_name As String, _
                                                          ByVal proto As String) As Long
Private Declare Function getprotobynumber Lib "WSOCK32.DLL" (ByVal proto As Long) As Long
Private Declare Function getprotobyname Lib "WSOCK32.DLL" (ByVal proto_name As String) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVR As Long, _
                                                       lpWSAD As WSADataType) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Sub WSASetLastError Lib "WSOCK32.DLL" (ByVal iError As Long)
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAIsBlocking Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAUnhookBlockingHook Lib "WSOCK32.DLL" () As Long
Private Declare Function WSASetBlockingHook Lib "WSOCK32.DLL" (ByVal lpBlockFunc As Long) As Long
Private Declare Function WSACancelBlockingCall Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAAsyncGetServByName Lib "WSOCK32.DLL" (ByVal hwnd As Long, _
                                                                  ByVal wMsg As Long, _
                                                                  ByVal serv_name As String, _
                                                                  ByVal proto As String, _
                                                                  buf As Any, _
                                                                  ByVal buflen As Long) As Long
Private Declare Function WSAAsyncGetServByPort Lib "WSOCK32.DLL" (ByVal hwnd As Long, _
                                                                  ByVal wMsg As Long, _
                                                                  ByVal port As Long, _
                                                                  ByVal proto As String, _
                                                                  buf As Any, _
                                                                  ByVal buflen As Long) As Long
Private Declare Function WSAAsyncGetProtoByName Lib "WSOCK32.DLL" (ByVal hwnd As Long, _
                                                                   ByVal wMsg As Long, _
                                                                   ByVal proto_name As String, _
                                                                   buf As Any, _
                                                                   ByVal buflen As Long) As Long
Private Declare Function WSAAsyncGetProtoByNumber Lib "WSOCK32.DLL" (ByVal hwnd As Long, _
                                                                     ByVal wMsg As Long, _
                                                                     ByVal number As Long, _
                                                                     buf As Any, _
                                                                     ByVal buflen As Long) As Long
Private Declare Function WSAAsyncGetHostByName Lib "WSOCK32.DLL" (ByVal hwnd As Long, _
                                                                  ByVal wMsg As Long, _
                                                                  ByVal host_name As String, _
                                                                  buf As Any, _
                                                                  ByVal buflen As Long) As Long
Private Declare Function WSAAsyncGetHostByAddr Lib "WSOCK32.DLL" (ByVal hwnd As Long, _
                                                                  ByVal wMsg As Long, _
                                                                  addr As Long, _
                                                                  ByVal addr_len As Long, _
                                                                  ByVal addr_type As Long, _
                                                                  buf As Any, _
                                                                  ByVal buflen As Long) As Long
Private Declare Function WSACancelAsyncRequest Lib "WSOCK32.DLL" (ByVal hAsyncTaskHandle As Long) As Long
Private Declare Function WSAAsyncSelect Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                           ByVal hwnd As Long, _
                                                           ByVal wMsg As Long, _
                                                           ByVal lEvent As Long) As Long
Private Declare Function WSARecvEx Lib "WSOCK32.DLL" (ByVal s As Long, _
                                                      buf As Any, _
                                                      ByVal buflen As Long, _
                                                      ByVal Flags As Long) As Long
                                                      Private Listening As Boolean
Public Property Get isListening()
isListening = Listening
End Property

Private Function AddrToIP(ByVal AddrOrIP As String) As String


    On Error Resume Next
    AddrToIP = GetAscIP(GetHostByNameAlias(AddrOrIP))
    If err Then
        AddrToIP = "255.255.255.255"
    End If
    On Error GoTo 0

End Function

Public Sub CloseSock(Optional Sockid As Long)
Listening = False

    If Sockid = 0 Then
        closesocket WSock1
        closesocket WSock
        SState = 0
     Else 'NOT SOCKID...
        closesocket Sockid
    End If

End Sub

Public Sub ConnectionClose()


closesocket WSock1
RaiseEvent Closed
End Sub

Private Function ConnectSock(ByVal host As String, _
                             ByVal IntPort As Long, _
                             ByVal retIpPort As String, _
                             ByVal HWndToMsg, _
                             ByVal Async As Long) As Long

  
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
        End If
    End With 'SOCKIN
    sockin.sin_addr = GetHostByNameAlias(host)
    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    SRemoteHost = GetAscIP(sockin.sin_addr)
    SRemotePort = ntohs(sockin.sin_port)
    s = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If SetSockLinger(s, 1, 1) = SOCKET_ERROR Then
        If s > 0 Then
            closesocket s
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If Not Async Then
        If connect(s, sockin, sockaddr_size) <> 0 Then
            If s > 0 Then
                closesocket s
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(s, HWndToMsg, ByVal WM_MOUSEMOVE, ByVal SelectOps) Then
            If s > 0 Then
                closesocket s
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
     Else 'NOT NOT...
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(s, HWndToMsg, ByVal WM_MOUSEMOVE, ByVal SelectOps) Then
            If s > 0 Then
                closesocket s
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        If connect(s, sockin, sockaddr_size) <> -1 Then
            If s > 0 Then
                closesocket s
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
    End If
    ConnectSock = s

End Function

Public Sub ConnectTo(RemotePort As Long, _
                     RemoteHost As String)


    WSock1 = ConnectSock(RemoteHost, RemotePort, SLocalPort, pcAjf.hwnd, True)
    If WSock1 > 0 Then
        SState = 2
     Else 'NOT WSOCK1...
        RaiseEvent Error$(0, "Error ( Socket is Busy ,Error Connecting )")
        SState = 1
    End If

End Sub

Private Function GetAscIP(ByVal inn As Long) As String

  
  Dim lpStr     As Long
  Dim nStr      As Long
  Dim retString As String

    On Error Resume Next
    retString = String$(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        GetAscIP = "255.255.255.255"
        Exit Function
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

Private Function GetHostByAddress(ByVal addr As Long) As String

  
  Dim phe        As Long
  Dim heDestHost As HOSTENT
  Dim hostname   As String

    On Error Resume Next
    phe = gethostbyaddr(addr, 4, PF_INET)
    If phe <> 0 Then
        MemCopy heDestHost, ByVal phe, hostent_size
        hostname = String$(256, 0)
        MemCopy ByVal hostname, ByVal heDestHost.h_name, 256
        GetHostByAddress = Left$(hostname, InStr(hostname, Chr$(0)) - 1)
     Else 'NOT PHE...
        GetHostByAddress = WSA_NoName
    End If
    If err Then
        GetHostByAddress = WSA_NoName
    End If
    On Error GoTo 0

End Function

Private Function GetHostByNameAlias(ByVal hostname As String) As Long

  
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
         Else 'NOT PHE...
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
        s = Left$(s, InStr(s, Chr$(0)) - 1)
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


Private Function IpToAddr(ByVal AddrOrIP As String) As String


    On Error Resume Next
    IpToAddr = GetHostByAddress(GetHostByNameAlias(AddrOrIP))
    If err Then
        IpToAddr = WSA_NoName
    End If
    On Error GoTo 0

End Function

Private Function ListenForConnect(ByVal IntPort, _
                                  ByVal HWndToMsg) As Long

  
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
        End If
    End With 'SOCKIN
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        'Actives = False
        Exit Function
    End If
    s = socket(PF_INET, SOCK_STREAM, 0)
    If s < 0 Then
        ListenForConnect = INVALID_SOCKET
        'Actives = False
        Exit Function
    End If
    If bind(s, sockin, sockaddr_size) Then
        If s > 0 Then
            closesocket s
        End If
        ListenForConnect = INVALID_SOCKET
        'Actives = False
        Exit Function
    End If
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, HWndToMsg, ByVal WM_MOUSEMOVE, ByVal SelectOps) Then
        If s > 0 Then
            closesocket s
        End If
        ListenForConnect = SOCKET_ERROR
        'Actives = False
        Exit Function
    End If
    If listen(s, 1) Then
        If s > 0 Then
            closesocket s
        End If
        ListenForConnect = INVALID_SOCKET
        'Actives = False
        Exit Function
    End If
    ListenForConnect = s
    If Not s = 0 Then
        'Actives = True
    End If

End Function

Public Sub ListenTo()


    WSock = ListenForConnect(SPort, pcAjf.hwnd)
    If WSock > 0 Then
        SState = 2
        Listening = True
     Else 'NOT WSOCK...
        Listening = False
        Debug.Print err.Description
        RaiseEvent Error$(4, "Port Error ( Socket is Busy Try Another Port )")
        SState = 1
        
    End If

End Sub
Public Function LocalIP() As String
On Error GoTo err
Dim T As Long
GotIP = ""
AsyncRead "http://members.lycos.co.uk/uptomoon/ip.php", vbAsyncTypeFile, "IP", vbAsyncReadGetFromCacheIfNetFail
T = timeGetTime + 30000
re:
Sleep 50
DoEvents
If Not T < timeGetTime And GotIP = "" Then GoTo re
LocalIP = GotIP
Exit Function
err:
err.Clear
LocalIP = LocalIPOwn
End Function

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error GoTo err
If AsyncProp.AsyncType = vbAsyncTypeFile Then
If Dir(AsyncProp.Value) <> "" Then
Dim A As String
Open AsyncProp.Value For Input As #1
A = input$(LOF(1), 1)
Close #1
GotIP = Split(A, "[~|~]")(1)
Else
GotIP = LocalIPOwn
End If
Else
GotIP = LocalIPOwn
End If
Exit Sub
err:
GotIP = LocalIPOwn
err.Clear
End Sub

Public Function LocalHost() As String


    LocalHost = IpToAddr(LocalIPOwn)

End Function

Public Function LocalIPOwn() As String


    LocalIPOwn = AddrToIP(GetLocalHostName)

End Function

Public Property Let LocalPort(ByVal lngPort As Long)


    SLocalPort = lngPort

End Property

Public Property Get NetConnected() As Boolean

  
  Dim ip_address        As String
  Dim hostent_addr      As Long
  Dim host              As HOSTENT
  Dim hostip_addr       As Long
  Dim temp_ip_address() As Byte
  Dim i                 As Long

    hostent_addr = gethostbyname("www.yahoo.com")
    If hostent_addr = 0 Then
        NetConnected = False
        Exit Property
    End If
    MemCopy host, hostent_addr, LenB(host)
    With host
    If .h_length = 0 Then .h_length = 1
        MemCopy hostip_addr, .h_addr_list, 4
        ReDim temp_ip_address(1 To .h_length)
        MemCopy temp_ip_address(1), hostip_addr, .h_length
        For i = 1 To .h_length
            DoEvents
            ip_address = ip_address & temp_ip_address(i) & "."
        Next i
    End With 'HOST
    ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
    If IsNumeric(Mid$(ip_address, 1, 1)) Then
        NetConnected = True
     Else 'NOT ISNUMERIC(MID$(IP_ADDRESS,...
        NetConnected = False
    End If

End Property

Private Sub pcAjf_MouseMove(Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

  
  Dim udtRemoteAddr    As sockaddr
  Dim ReadBuffer(1000) As Byte
  Dim DataLen          As Long
  Dim Data             As String

    errnum = WSAGetAsyncError(X)
    If errnum <> 0 Then
        
        RaiseEvent Error$(errnum, "")
        
    End If
    Select Case X / Screen.TwipsPerPixelX
     Case FD_READ
        'Debug.Print "Data"
        DataLen = recv(WSock1, ReadBuffer(0), 1000, 0)
        If DataLen > 0 Then
            Data = StrConv(ReadBuffer, vbUnicode)
            RaiseEvent DataArrival(Data)
         Else 'NOT DATALEN...
            Timein = True
        End If
     Case FD_ACCEPT
        'Debug.Print vbCrLf & "Accepted"
        WSock1 = accept(WSock, udtRemoteAddr, LenB(udtRemoteAddr))
        If WSock1 > 0 Then
            RaiseEvent ConnectionRequest
            SState = 3
        End If
     Case FD_CONNECT
        'Debug.Print vbCrLf & "Connected"
        RaiseEvent Connected
        Connin = True
     Case FD_CLOSE
        'Debug.Print vbCrLf & "Closed"
        ConnectionClose
        Connin = False
    End Select

End Sub
Public Property Get Connected() As Boolean
Connected = Connin
End Property

Public Property Get port() As Long


    port = SPort

End Property

Public Property Let port(ByVal P As Long)


    SPort = P

End Property

Private Function SendData(ByVal s, _
                          vMessage As Variant) As Long

  
  Dim TheMsg() As Byte
  Dim sTemp    As String

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

Public Sub SendDataTo(ByVal Data As String)


    SendData WSock1, Data

End Sub

Public Sub Show(ByVal s As Boolean)


    IdeCheck = s

End Sub

Public Sub SockClose()
Listening = False

    closesocket WSock1
    closesocket WSock

End Sub



Private Sub UserControl_Initialize()

  
  Dim s As String

    Randomize: SLocalPort = Int(Rnd(32000) * 32000)
    If StartWinsock(s) = False Then
        RaiseEvent Error$(0, "Error Starting Winsock")
        SState = 0
     Else 'NOT STARTWINSOCK(S)...
        SState = 1
    End If

End Sub

Private Sub UserControl_Resize()


    If IdeCheck = False Then
        Exit Sub
    End If
    UserControl.Width = 435
    UserControl.Height = 435

End Sub

Private Function WSAGetAsyncError(ByVal lParam As Long) As Integer


    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000

End Function


