VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinSock2_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WinSock 2.2 Client/Server Class                                '
'Description:                                                   '
'   A native implementation of a Winsock 2.2 Client/Server      '
'   Program. It is designed to allow Standard TCP/UDP Server and'
'   Client functions within the VBE Environment.                '
'Functions:                                                     '
'   StartServer                                                 '
'   AcceptSocket                                                '
'   CloseAcceptedSocket                                         '
'   CloseServer                                                 '
'   StartClient                                                 '
'   CloseClient                                                 '
'   Receive                                                     '
'   ReceiveBytes                                                '
'   ReceiveFrom                                                 '
'   ReceiveBytesFrom                                            '
'   Send                                                        '
'   SendBytes                                                   '
'   SendTo                                                      '
'   SendBytesTo                                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'C Data Types                                                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type WORD1
    Lo As Byte
    Hi As Byte
End Type
Private Type WORD2
    Value As Integer
End Type
Private Type DWORD1
    Lo As WORD1
    Hi As WORD1
End Type
Private Type DWORD2
    Lo As Integer
    Hi As Integer
End Type
Private Type DWORD3
    Value As Long
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Structs                                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'Description:
'   The IN6_ADDR structure specifies an IPv6 transport address.
'Definition:
'   Typedef struct in6_addr {
'       union {
'           UCHAR  Byte[16];
'           USHORT Word[8];
'       } u;
'   } IN6_ADDR, *PIN6_ADDR, *LPIN6_ADDR;
Private Type IN6_ADDR1
    Value(0 To 15) As Byte
End Type
Private Type IN6_ADDR2
    Value(0 To 7) As Integer
End Type
#If VBA7 Then
Private Type PIN6_ADDR
    Value As LongPtr
End Type
#Else
Private Type PIN6_ADDR
    Value As Long
End Type
#End If
''
'Description:
'   The in_addr structure represents an IPv4 address.
'Definition:
'   Typedef struct in_addr {
'       union {
'           struct {
'               UCHAR s_b1;
'               UCHAR s_b2;
'               UCHAR s_b3;
'               UCHAR s_b4;
'           } S_un_b;
'           struct {
'               USHORT s_w1;
'               USHORT s_w2;
'           } S_un_w;
'           ULONG S_addr;
'       } S_un;
'   } IN_ADDR, *PIN_ADDR, *LPIN_ADDR;
Private Type IN_ADDR1
    Value(0 To 3) As Byte
End Type
Private Type IN_ADDR2
    Value(0 To 1) As Integer
End Type
Private Type IN_ADDR3
    Value As Long
End Type
#If VBA7 Then
Private Type PIN_ADDR
    Value As LongPtr
End Type
#Else
Private Type PIN_ADDR
    Value As Long
End Type
#End If
''
'Description:
'   The WSADATA structure contains information about the Windows Sockets implementation.
'Definition:
'    Typedef struct WSAData {
'            WORD                    wVersion;
'            WORD                    wHighVersion;
'            char                    szDescription[WSADESCRIPTION_LEN+1];
'            char                    szSystemStatus[WSASYS_STATUS_LEN+1];
'            unsigned short          iMaxSockets;
'            unsigned short          iMaxUdpDg;
'            char FAR *              lpVendorInfo;
'    } WSADATA, FAR * LPWSADATA;
#If VBA7 Then
Private Type WSADATA
    wVersion As WORD1
    wHighVersion As WORD1
    szDescription As LongPtr
    szSystemStatus As LongPtr
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As LongPtr
End Type
#Else
Private Type WSADATA
    wVersion As WORD1
    wHighVersion As WORD1
    szDescription As Long
    szSystemStatus As Long
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
#End If
''
'Description:
'   Socket Address Structure
'Definition:
'   struct sockaddr {
'        ushort  sa_family;
'        char    sa_data[14];
'   };
'
'   struct sockaddr_in {
'        short   sin_family;
'        u_short sin_port;
'        struct  in_addr sin_addr;
'        char    sin_zero[8];
'   };
'
'   struct sockaddr_in6 {
'        short   sin6_family;
'        u_short sin6_port;
'        u_long  sin6_flowinfo;
'        struct  in6_addr sin6_addr;
'        u_long  sin6_scope_id;
'   };
Private Type SOCKADDR
    sa_family As Integer
    sa_data(0 To 13) As Byte
End Type
Private Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As IN_ADDR3
    sin_zero(0 To 7) As Byte
End Type
Private Type sockaddr_in6
    sin6_family As Integer
    sin6_port As Integer
    sin6_flowinfo As Long
    sin6_addr As IN6_ADDR2
    sin6_scape_id As Long
End Type
''
'Description:
'   The addrinfo structure is used by the getaddrinfo function to hold host address information.
'Definition:
'   Typedef struct addrinfo {
'       int             ai_flags;
'       int             ai_family;
'       int             ai_sockType;
'       int             ai_protocol;
'       size_t          ai_addrlen;
'       char            *ai_canonname;
'       struct sockaddr *ai_addr;
'       struct addrinfo *ai_next;
'   } ADDRINFOA, *PADDRINFOA;
#If VBA7 Then
Private Type ADDRINFO
    ai_flags As Long
    ai_family As Long
    ai_sockType As Long
    ai_protocol As Long
    ai_addrlen As Long
    ai_canonName As LongPtr
    ai_addr As LongPtr
    ai_next As LongPtr
End Type
#Else
Private Type ADDRINFO
    ai_flags As Long
    ai_family As Long
    ai_sockType As Long
    ai_protocol As Long
    ai_addrlen As Long
    ai_canonName As Long
    ai_addr As Long
    ai_next As Long
End Type
#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Constants                                                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If VBA7 Then
Private Const NULL_PTR As LongPtr = 0&
#Else
Private Const NULL_PTR As Long = 0
#End If
Private Const CFALSE = 0&
Private Const CTRUE = 1&
Private Const INADDR_ANY As Long = 0
Private Const INADDR_NONE = -1&
Private Const INVALID_SOCKET = NULL_PTR
Private Const SOCKET_ERROR = -1&
Private Const SOMAXCONN = 10
Private Const MAX_PACKET = 65507
Private Const SO_MAX_MSG_SIZE = &H2003&
Private Const SOL_SOCKET = &HFFFF&
Private Const DEFAULT_LNG = 0&
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100&
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000&
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800&
Private Const FORMAT_MESSAGE_FROM_STRING = &H400&
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200&
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF&
Private Const NI_MAXSERV = 32
Private Const NI_MAXHOST = 1025
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Enums                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Enum AF
    AF_UNSPEC = 0&
    AF_INET = 2&
    AF_IPX = 6&
    AF_APPLETALK = 16&
    AF_NETBIOS = 17&
    AF_INET6 = 23&
    AF_IRDA = 26&
    AF_BTH = 32&
End Enum
Private Enum sock
    SOCK_STREAM = 1
    SOCK_DGRAM = 2
    SOCK_RAW = 3
    SOCK_RDM = 4
    SOCK_SEQPACKET = 5
End Enum
Private Enum IPPROTO
    IPPROTO_ICMP = 1
    IPPROTO_IGMP = 2
    BTHPROTO_RFCOMM = 3
    IPPROTO_TCP = 6
    IPPROTO_UDP = 17
    IPPROTO_ICMPV6 = 58
    IPPROTO_RM = 113
End Enum
Private Enum ai_flags
    AI_PASSIVE = &H1&
    AI_CONONNAME = &H2&
    AI_NUMERICHOST = &H4&
    AI_ALL = &H100&
    AI_ADDRCONFIG = &H400&
    AI_V4MAPPED = &H800&
    AI_NON_AUTHORITATIVE = &H4000&
    AI_SECURE = &H8000&
    AI_RETURN_PREFERRED_NAMES = &H10000
    AI_FQDN = &H20000
    AI_FILESERVER = &H40000
End Enum
Private Enum SD_FLAGs
    SD_RECEIVE = 0
    SD_SEND = 1
    SD_BOTH = 2
End Enum
Private Enum MSG_FLAGs
    Default = 0
    MSG_PEEK = &H2&
    MSG_OOB = &H1&
    MSG_WAITALL = &H8&
End Enum
Private Enum NameInfo_Flags
    NI_NOFQDN = &H1&
    NI_NUMERICHOST = &H2&
    NI_NAMEREQD = &H4&
    NI_NUMERICSERV = &H8&
    NI_DGRAM = &H10&
End Enum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'API Delecrations                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'Description:
'   Copies a block of memory from one location to another.
'Definition:
'   void CopyMemory(
'       _In_       PVOID  Destination,
'       _In_ const VOID   *Source,
'       _In_       SIZE_T Length
'   );
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "Kernel32.dll" Alias "RtlCopyMemory" _
( _
    ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, _
    ByVal length As Long _
)
#Else
Private Declare Sub CopyMemory Lib "Kernel32.Dll" Alias "RtlMoveMemory" _
( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal length As Long _
)
#End If
''
'Description:
'   Moves a block of memory from one location to another.
'Definition:
'   void MoveMemory(
'       _In_       PVOID  Destination,
'       _In_ const VOID   *Source,
'       _In_       SIZE_T Length
'   );
#If VBA7 Then
Private Declare PtrSafe Sub MoveMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" _
( _
  ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, _
    ByVal length As Long _
)
#Else
Private Declare Sub MoveMemory Lib "Kernel32.Dll" Alias "RtlMoveMemory" _
( _
  ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal length As Long _
)
#End If
''
'Description:
'   Returns the error status for the last Windows Sockets operation that failed.
'Definition:
'   int WSAGetLastError();
#If VBA7 Then
Private Declare PtrSafe Function WSAGetLastError Lib "Ws2_32.dll" () As Long
#Else
Private Declare Function WSAGetLastError Lib "Ws2_32.dll" () As Long
#End If
''
'Description:
'   The WSAStartup function initiates use of the Winsock DLL by a process.
'Definition:
'   int WSAStartup(
'             WORD      wVersionRequired,
'       [out] LPWSADATA lpWSAData
'   );
#If VBA7 Then
Private Declare PtrSafe Function WSAStartup Lib "Ws2_32.dll" _
( _
    wVersionRequired As WORD1, _
    ByVal lpWSAData As LongPtr _
) As Long
#Else
Private Declare Function WSAStartup Lib "Ws2_32.dll" _
( _
    wVersionRequired As WORD1, _
    ByVal lpWSAData As Long _
) As Long
#End If
''
'Description:
'   The WSACleanup function terminates use of the Winsock 2 DLL (Ws2_32.dll).
'Definition:
'   int WSACleanup();
#If VBA7 Then
Private Declare PtrSafe Function WSACleanup Lib "Ws2_32.dll" () As Long
#Else
Private Declare Function WSACleanup Lib "Ws2_32.dll" () As Long
#End If
''
'Descripiton:
'   The socket function creates a socket that is bound to a specific transport service provider.
'Definition:
'   SOCKET WSAAPI socket(
'       [in] int af,
'       [in] int Type,
'       [in] int protocol
'   );
#If VBA7 Then
Private Declare PtrSafe Function CreateSocket Lib "Ws2_32.dll" Alias "socket" _
( _
    ByVal AddressFamily As AF, _
    ByVal SockType As sock, _
    ByVal Protocal As IPPROTO _
) As LongPtr
#Else
Private Declare Function CreateSocket Lib "Ws2_32.dll" Alias "socket" _
( _
    ByVal AddressFamily As AF, _
    ByVal SockType As sock, _
    ByVal Protocal As IPPROTO _
) As Long
#End If
''
'Description:
'   The htonl function converts a u_long from host to TCP/IP network byte order (which is big-endian).
'Definition:
'   u_long WSAAPI htonl(
'       [in] u_long hostlong
'   );
#If VBA7 Then
Private Declare PtrSafe Function htonl Lib "Ws2_32.dll" _
( _
    ByVal hostlong As Long _
) As Long
#Else
Private Declare Function htonl Lib "Ws2_32.dll" _
( _
    ByVal hostlong As Long _
) As Long
#End If
''
'Description:
'   The htons function converts a u_short from host to TCP/IP network byte order (which is big-endian).
'Definiton:
'   u_short WSAAPI htons(
'       [in] u_short hostlong
'   );
#If VBA7 Then
Private Declare PtrSafe Function htons Lib "Ws2_32.dll" _
( _
    ByVal hostlong As Integer _
) As Integer
#Else
Private Declare Function htons Lib "Ws2_32.dll" _
( _
    ByVal hostlong As Integer _
) As Integer
#End If
''
'Description:
'   The ntohs function converts a u_short from TCP/IP network byte order to host byte order (which is little-endian on Intel processors).
'Definition:
'   u_short WSAAPI ntohs(
'       [in] u_short netshort
'   );
#If VBA7 Then
Private Declare PtrSafe Function ntohs Lib "Ws2_32.dll" _
( _
    ByVal netshort As Integer _
) As Integer
#Else
Private Declare Function ntohs Lib "Ws2_32.dll" _
( _
    ByVal netshort As Integer _
) As Integer
#End If
''
'Description:
'   The bind function associates a local address with a socket.
'Definition:
'   int WSAAPI bind(
'       [in] SOCKET         s,
'       [in] const sockaddr *name,
'       [in] int            namelen
'   );
#If VBA7 Then
Private Declare PtrSafe Function bind Lib "Ws2_32.dll" _
( _
    ByVal s As LongPtr, _
    ByVal name As LongPtr, _
    ByVal namelen As Long _
) As Long
#Else
Private Declare Function bind Lib "Ws2_32.dll" _
( _
    ByVal s As Long, _
    ByVal name As Long, _
    ByVal namelen As Long _
) As Long
#End If
''
'Description:
'   The closesocket function closes an existing socket.
'Definition:
'   int closesocket(
'       [in] SOCKET s
'   );
#If VBA7 Then
Private Declare PtrSafe Function CloseSocket Lib "Ws2_32.dll" Alias "closesocket" _
( _
    ByVal s As LongPtr _
) As Long
#Else
Private Declare Function CloseSocket Lib "Ws2_32.dll" Alias "closesocket" _
( _
    ByVal s As Long _
) As Long
#End If
''
'Description:
'   The inet_addr function converts a string containing an IPv4 dotted-decimal address into a proper address for the IN_ADDR structure.
'Definition:
'   unsigned long WSAAPI inet_addr(
'       const char *cp
'   );
#If VBA7 Then
Private Declare PtrSafe Function inet_addr Lib "Ws2_32.dll" _
( _
    ByVal cp As LongPtr _
) As Long
#Else
Private Declare Function inet_addr Lib "Ws2_32.dll" _
( _
    ByVal cp As Long _
) As Long
#End If
''
'Description:
'   The getaddrinfo function provides protocol-independent translation from an ANSI host name to an address.
'Definition:
'   INT WSAAPI getaddrinfo(
'       [in, optional] PCSTR           pNodeName,
'       [in, optional] PCSTR           pServiceName,
'       [in, optional] const ADDRINFOA *pHints,
'       [out] PADDRINFOA * ppResult
'   );
#If VBA7 Then
Private Declare PtrSafe Function GetAddrInfo Lib "Ws2_32.dll" Alias "getaddrinfo" _
( _
    ByVal pNodeName As String, _
    ByVal pServiceName As String, _
    ByVal pHints As LongPtr, _
    ppResult As LongPtr _
) As Long
#Else
Private Declare Function GetAddrInfo Lib "Ws2_32.dll" Alias "getaddrinfo" _
( _
    ByVal pNodeName As String, _
    ByVal pServiceName As String, _
    ByVal pHints As Long, _
    ppResult As Long _
) As Long
#End If
''
'Description:
'   The freeaddrinfo function frees address information that the getaddrinfo function dynamically allocates in addrinfo structures.
'Definition:
'   VOID WSAAPI freeaddrinfo(
'       [in] PADDRINFOA pAddrInfo
'   );
#If VBA7 Then
Private Declare PtrSafe Sub FreeAddrInfo Lib "Ws2_32.dll" Alias "freeaddrinfo" _
( _
    ByVal pAddrInfo As LongPtr _
)
#Else
Private Declare Sub FreeAddrInfo Lib "Ws2_32.dll" Alias "freeaddrinfo" _
( _
    ByVal pAddrInfo As Long _
)
#End If
''
'Description:
'   The listen function places a socket in a state in which it is listening for an incoming connection.
'Definition:
'   int WSAAPI listen(
'       [in] SOCKET s,
'       [in] int    backlog
'   );
#If VBA7 Then
Private Declare PtrSafe Function Listen Lib "Ws2_32.dll" Alias "listen" _
( _
    ByVal s As LongPtr, _
    ByVal backlog As Long _
) As Long
#Else
Private Declare Function Listen Lib "Ws2_32.dll" Alias "listen" _
( _
    ByVal s As Long, _
    ByVal backlog As Long _
) As Long
#End If
''
'Description:
'   The accept function permits an incoming connection attempt on a socket.
'Definition:
'   SOCKET WSAAPI accept(
'       [in]      SOCKET   s,
'       [out]     sockaddr *addr,
'       [in, out] int      *addrlen
'   );
#If VBA7 Then
Private Declare PtrSafe Function Accept Lib "Ws2_32.dll" Alias "accept" _
( _
    ByVal s As LongPtr, _
    ByVal addr As LongPtr, _
    ByVal addrLen As LongPtr _
) As LongPtr
#Else
Private Declare Function Accept Lib "Ws2_32.dll" Alias "accept" _
( _
    ByVal s As Long, _
    ByVal addr As Long, _
    ByVal addrLen As Long _
) As Long
#End If
''
'Description:
'   The shutdown function disables sends or receives on a socket.
'Definition:
'   int WSAAPI shutdown(
'       [in] SOCKET s,
'       [in] int    how
'   );
#If VBA7 Then
Private Declare PtrSafe Function Shutdown Lib "Ws2_32.dll" Alias "shutdown" _
( _
    ByVal s As LongPtr, _
    ByVal how As Long _
) As Long
#Else
Private Declare Function Shutdown Lib "Ws2_32.dll" Alias "shutdown" _
( _
    ByVal s As Long, _
    ByVal how As Long _
) As Long
#End If
''
'Description:
'   The recv function receives data from a connected socket or a bound connectionless socket.
'Definition:
'   int WSAAPI recv(
'       [in]  SOCKET s,
'       [out] char   *buf,
'       [in]  int    len,
'       [in]  int    flags
'   );
#If VBA7 Then
Private Declare PtrSafe Function Recv Lib "Ws2_32.dll" Alias "recv" _
( _
    ByVal s As LongPtr, _
    ByVal buf As LongPtr, _
    ByVal rLen As Long, _
    ByVal Flags As Long _
) As Long
#Else
Private Declare Function Recv Lib "Ws2_32.dll" Alias "recv" _
( _
    ByVal s As Long, _
    ByVal buf As Long, _
    ByVal rLen As Long, _
    ByVal Flags As Long _
) As Long
#End If
''
'Description:
'   The send function sends data on a connected socket.
'Definition:
'   int WSAAPI send(
'       [in] SOCKET     s,
'       [in] const char *buf,
'       [in] int        len,
'       [in] int        flags
'   );
#If VBA7 Then
Private Declare PtrSafe Function SocketSend Lib "Ws2_32.dll" Alias "send" _
( _
    ByVal s As LongPtr, _
    ByVal buf As LongPtr, _
    ByVal sLen As Long, _
    ByVal Flags As Long _
) As Long
#Else
Private Declare Function SocketSend Lib "Ws2_32.dll" Alias "send" _
( _
    ByVal s As Long, _
    ByVal buf As Long, _
    ByVal sLen As Long, _
    ByVal Flags As Long _
) As Long
#End If
''
'Description:
'   The recvfrom function receives a datagram, and stores the source address.
'Definition:
'   int WSAAPI recvfrom(
'       [in]                SOCKET   s,
'       [out]               char     *buf,
'       [in]                int      len,
'       [in]                int      flags,
'       [out]               sockaddr *from,
'       [in, out, optional] int      *fromlen
'   );
#If VBA7 Then
Private Declare PtrSafe Function RecvFrom Lib "Ws2_32.dll" Alias "recvfrom" _
( _
    ByVal s As LongPtr, _
    ByVal buf As LongPtr, _
    ByVal bLen As Long, _
    ByVal Flags As Long, _
    ByVal from As LongPtr, _
    ByVal fromlen As LongPtr _
) As Long
#Else
Private Declare Function RecvFrom Lib "Ws2_32.dll" Alias "recvfrom" _
( _
    ByVal s As Long, _
    ByVal buf As Long, _
    ByVal bLen As Long, _
    ByVal Flags As Long, _
    ByVal from As Long, _
    ByVal fromlen As Long _
) As Long
#End If
''
'Description:
'   The sendto function sends data to a specific destination.
'Definition:
'   int WSAAPI sendto(
'      [in] SOCKET         s,
'      [in] const char     *buf,
'      [in] int            len,
'      [in] int            flags,
'      [in] const sockaddr *to,
'      [in] int            tolen
'    );
#If VBA7 Then
Private Declare PtrSafe Function SocketSendTo Lib "Ws2_32.dll" Alias "sendto" _
( _
    ByVal s As LongPtr, _
    ByVal buf As LongPtr, _
    ByVal bLen As Long, _
    ByVal Flags As Long, _
    ByVal sTo As LongPtr, _
    ByVal tolen As Long _
) As Long
#Else
Private Declare Function SocketSendTo Lib "Ws2_32.dll" Alias "sendto" _
( _
    ByVal s As Long, _
    ByVal buf As Long, _
    ByVal bLen As Long, _
    ByVal Flags As Long, _
    ByVal sTo As Long, _
    ByVal tolen As Long _
) As Long
#End If
''
'Description:
'   The getsockopt function retrieves a socket option.
'Definition:
'   int getsockopt(
'       [in]      SOCKET s,
'       [in]      int    level,
'       [in]      int    optname,
'       [out]     char   *optval,
'       [in, out] int    *optlen
'   );
#If VBA7 Then
Private Declare PtrSafe Function GetSockOpt Lib "Ws2_32.dll" Alias "getsockopt" _
( _
    ByVal s As LongPtr, _
    ByVal level As Long, _
    ByVal optname As Long, _
    ByVal optval As LongPtr, _
    ByVal optlen As LongPtr _
) As Long
#Else
Private Declare Function GetSockOpt Lib "Ws2_32.dll" Alias "getsockopt" _
( _
    ByVal s As Long, _
    ByVal level As Long, _
    ByVal optname As Long, _
    ByVal optval As Long, _
    ByVal optlen As Long _
) As Long
#End If
''
'Description:
'   The connect function establishes a connection to a specified socket.
'Definition:
'   int WSAAPI connect(
'       [in] SOCKET         s,
'       [in] const sockaddr *name,
'       [in] int            namelen
'   );
#If VBA7 Then
Private Declare PtrSafe Function Connect Lib "Ws2_32.dll" Alias "connect" _
( _
    ByVal socket As LongPtr, _
    ByVal sAddr As LongPtr, _
    ByVal nLen As Long _
) As Long
#Else
Private Declare Function Connect Lib "Ws2_32.dll" Alias "connect" _
( _
    ByVal socket As Long, _
    ByVal sAddr As Long, _
    ByVal nLen As Long _
) As Long
#End If
''
'Description:
'   Formats a message string. The function requires a message definition as input.
'   The message definition can come from a buffer passed into the function. It can come
'   from a message table resource in an already-loaded module. Or the caller can ask the
'   function to search the system's message table resource(s) for the message definition.
'   The function finds the message definition in a message table resource based on a message
'   identifier and a language identifier. The function copies the formatted message text to
'   an output buffer, processing any embedded insert sequences if requested.
'Definition:
'   DWORD FormatMessage(
'       [in]           DWORD   dwFlags,
'       [in, optional] LPCVOID lpSource,
'       [in]           DWORD   dwMessageId,
'       [in]           DWORD   dwLanguageId,
'       [out]          LPTSTR  lpBuffer,
'       [in]           DWORD   nSize,
'       [in, optional] va_list * Arguments
'   );
#If VBA7 Then
Private Declare PtrSafe Function FormatMessage Lib "Kernel32.dll" Alias "FormatMessageA" _
( _
    ByVal dwFlags As Long, _
    ByVal lpSource As LongPtr, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByVal Args As LongPtr _
) As Long
#Else
Private Declare Function FormatMessage Lib "Kernel32.Dll" Alias "FormatMessageA" _
( _
    ByVal dwFlags As Long, _
    ByVal lpSource As Long, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByVal Args As Long _
) As Long
#End If
''
'Description:
'   The getnameinfo function provides protocol-independent name resolution from an address to an ANSI host name and from a port number to the ANSI service name.
'Definition:
'   INT WSAAPI getnameinfo(
'       [in]  const SOCKADDR *pSockaddr,
'       [in]  socklen_t      SockaddrLength,
'       [out] PCHAR          pNodeBuffer,
'       [in]  DWORD          NodeBufferSize,
'       [out] PCHAR          pServiceBuffer,
'       [in]  DWORD          ServiceBufferSize,
'       [in]  INT            Flags
'   );
#If VBA7 Then
Private Declare PtrSafe Function GetNameInfo Lib "Ws2_32.dll" Alias "getnameinfo" _
( _
    ByVal pSockaddr As LongPtr, _
    ByVal SockaddrLength As Long, _
    ByVal pNodeBuffer As LongPtr, _
    ByVal NodeBufferSize As Long, _
    ByVal pServiceBuffer As LongPtr, _
    ByVal ServiceBufferSize As Long, _
    ByVal Flags As Long _
) As Long
#Else
Private Declare Function GetNameInfo Lib "Ws2_32.dll" Alias "getnameinfo" _
( _
    ByVal pSockaddr As Long, _
    ByVal SockaddrLength As Long, _
    ByVal pNodeBuffer As Long, _
    ByVal NodeBufferSize As Long, _
    ByVal pServiceBuffer As Long, _
    ByVal ServiceBufferSize As Long, _
    ByVal Flags As Long _
) As Long
#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Global Variables                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Initialized As Boolean
#If VBA7 Then
Private currSock As LongPtr
Private currClient As LongPtr
#Else
Private currSock As Long
Private currClient As Long
#End If
Private currType As Integer
Private isServer As Boolean
Private data As WSADATA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Enums                                                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Enum PROTOCOL
    UDP
    TCP
End Enum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Enums                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Enum CUSTOMERROR
    PROTO_UDP
    PROTO_TCP
    INVAILD_PROTO
    IS_SERVER
    IS_CLIENT
    ACCEPTED
    NOT_ACCEPTED
    NOT_STARTED
End Enum
Private Enum ERR_TYPE
    winsock
    custom
End Enum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal Helper Functions                                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ST()
    If Not Initialized Then
        Dim wVersion As WORD1
        Dim wErr As Long
        
        wVersion.Lo = 2
        wVersion.Hi = 2
        wErr = WSAStartup(wVersion, VarPtr(data))
        If wErr <> 0 Then
            ''  Tell the user that we could not find a usable Winsock DLL
            err.Raise wErr, "WinSock2", "Could Not Start Winsock API"
        End If
    End If
End Sub
Private Sub EN()
    Initialized = False
    isServer = False
    currType = -1
    If currSock <> NULL_PTR Then CloseSocket currSock
    If currClient <> NULL_PTR Then CloseSocket currClient
    currSock = NULL_PTR
    currClient = NULL_PTR
    
    WSACleanup
End Sub
Private Function StringToBuffer(Value As String, Optional WCHAR As Boolean = True) As Byte()
    Dim out() As Byte
    Dim i As Long
    
    If WCHAR Then
        Dim ii As Long
        
        ReDim out(0 To Len(Value) * 2 - 1)
        For i = 0 To UBound(out) Step 2
            ii = ii + 1
            out(i) = VBA.Asc(VBA.Mid$(Value, ii, 1))
        Next i
    Else
        ReDim out(0 To Len(Value) - 1)
        For i = 0 To UBound(out)
            out(i) = VBA.Asc(VBA.Mid$(Value, i + 1, 1))
        Next i
    End If
    
    StringToBuffer = out
End Function
Private Function BufferToString(Value() As Byte, Optional WCHAR As Boolean = True) As String
    Dim out As String
    Dim i As Long
    
    If WCHAR Then
        out = Value
    Else
        For i = LBound(Value) To UBound(Value)
            out = out & VBA.Chr$(Value(i))
        Next i
    End If
    
    BufferToString = VBA.Replace(out, VBA.Chr$(0), "")
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Struct Memory Functions                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'User Functions                                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'Description:
'   Starts a WinSock Server at the Specified Port
'Params:
'   Proto As PROTOCOL  -  This the Protocal used for the socket. Valid Values
'                         are UPD or TCP.
'   Port  As String    -  This is the Port Number that the socket will listen on.
'                         Valid Values:
'                             Service Name such as "http"
'                             Port Number such as  "80"
Public Function StartServer(Proto As PROTOCOL, Port As String)
    ST
    
    Dim sErr As Long
    Dim Source As String: Source = "StartServer"
    Dim errType As ERR_TYPE
    Dim Hints As ADDRINFO: Hints.ai_family = AF.AF_UNSPEC
    Dim Result As ADDRINFO
#If VBA7 Then
    Dim pResult As LongPtr: pResult = NULL_PTR
#Else
    Dim pResult As Long: pResult = NULL_PTR
#End If
    
    If Proto = UDP Then
        Hints.ai_sockType = sock.SOCK_DGRAM
        Hints.ai_protocol = IPPROTO.IPPROTO_UDP
    ElseIf Proto = TCP Then
        Hints.ai_sockType = sock.SOCK_STREAM
        Hints.ai_protocol = IPPROTO.IPPROTO_TCP
    Else
        sErr = CUSTOMERROR.INVAILD_PROTO
        errType = custom
        GoTo ErrHnd
    End If
    
    If currSock <> NULL_PTR Then
        If isServer Then
            sErr = CUSTOMERROR.IS_SERVER
        Else
            sErr = CUSTOMERROR.IS_CLIENT
        End If
        errType = custom
        GoTo ErrHnd
    End If
    
    sErr = GetAddrInfo("localhost", Port, VarPtr(Hints), pResult)
    If sErr <> 0 Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    Else
        CopyMemory VarPtr(Result), pResult, LenB(Result)
    End If
    
    currSock = CreateSocket(Result.ai_family, Result.ai_sockType, Result.ai_protocol)
    If currSock = INVALID_SOCKET Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
    
    sErr = bind(currSock, Result.ai_addr, Result.ai_addrlen)
    If sErr = SOCKET_ERROR Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
    
    If Proto = TCP Then
        sErr = Listen(currSock, SOMAXCONN)
        If sErr = SOCKET_ERROR Then
            sErr = WSAGetLastError
            errType = winsock
            GoTo ErrHnd
        End If
    End If
    
    FreeAddrInfo pResult
    currType = Proto
    isServer = True
    Exit Function
ErrHnd:
    EN
    If pResult <> NULL_PTR Then FreeAddrInfo pResult
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
''
'Description:
'   Accepts an incoming TCP Client Connection. This function is blocking
'   and will wait until a client is accepted.
'Params:
'   Optional ByRef Host As String  -  The String to fill with the Sender's Host Name.
'   Optional ByRef Port As String  -  The String to fill with the Sender's Port Number.
'Return:
'   AcceptSocket As Boolean  -  Returns True if a Connection is accepted.
Public Function AcceptSocket(Optional ByRef Host As String, Optional ByRef Port As String) As Boolean
    Dim sErr As Long
    Dim Source As String: Source = "AcceptSocket"
    Dim errType As ERR_TYPE
    Dim addr As sockaddr_in6
#If VBA7 Then
    Dim addrPtr As LongPtr: addrPtr = VarPtr(addr)
#Else
    Dim addrPtr As Long: addptr = VarPtr(addr)
#End If
    Dim addrLen As Long: addrLen = LenB(addr)
    Dim rHost(NI_MAXHOST) As Byte
    Dim rPort(NI_MAXSERV) As Byte
    
    If currSock = NULL_PTR Then
        sErr = CUSTOMERROR.NOT_STARTED
        errType = custom
        GoTo ErrHnd
    End If
    If currType = PROTOCOL.UDP Then
        sErr = CUSTOMERROR.PROTO_TCP
        errType = custom
        GoTo ErrHnd
    End If
    If currClient <> NULL_PTR Then
        sErr = CUSTOMERROR.ACCEPTED
        errType = custom
        GoTo ErrHnd
    End If
    
    currClient = Accept(currSock, addrPtr, VarPtr(addrLen))
    If currClient = INVALID_SOCKET Or currClient = -1 Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
    
    sErr = GetNameInfo(addrPtr, addrLen, VarPtr(rHost(0)), NI_MAXHOST, VarPtr(rPort(0)), NI_MAXSERV, NameInfo_Flags.NI_NUMERICSERV)
    If sErr <> 0 Then
        errType = winsock
        GoTo ErrHnd
    End If
    Host = BufferToString(rHost, False)
    Port = BufferToString(rPort, False)
    
    AcceptSocket = True
    Exit Function
ErrHnd:
    EN
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
''
'Description:
'   Closes a previously Accepted Socket. This function MUST be called if a Socket
'   was previously Accepted using the "AcceptSocket" function before accepting a
'   NEW Socket using "AcceptSocket"
Public Function CloseAcceptedSocket()
    Dim sErr As Long
    Dim errType As ERR_TYPE
    Dim Source As String: Source = "CloseAcceptedSocket"
    
    If currSock = NULL_PTR Then
        sErr = CUSTOMERROR.NOT_STARTED
        errType = custom
        GoTo ErrHnd
    End If
    If currType = PROTOCOL.UDP Then
        sErr = CUSTOMERROR.PROTO_TCP
        errType = custom
        GoTo ErrHnd
    End If
    
    If currType = PROTOCOL.TCP And currClient <> NULL_PTR Then
        sErr = Shutdown(currClient, SD_FLAGs.SD_BOTH)
        If sErr = SOCKET_ERROR Then
            sErr = WSAGetLastError
            errType = winsock
            GoTo ErrHnd
        End If
        
        sErr = CloseSocket(currClient)
        If sErr <> 0 Then
            sErr = WSAGetLastError
            errType = winsock
            GoTo ErrHnd
        End If
    End If
    
    currClient = NULL_PTR
    Exit Function
ErrHnd:
    EN
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
''
'Description:
'   Closes the Current Server
Public Function CloseServer()
    Dim sErr As Long
    Dim errType As ERR_TYPE
    Dim Source As String: Source = "CloseServer"
    
    If currSock = NULL_PTR Then
        sErr = CUSTOMERROR.NOT_STARTED
        errType = custom
        GoTo ErrHnd
    End If
        
    If Not isServer Then
        sErr = CUSTOMERROR.IS_CLIENT
        errType = custom
        GoTo ErrHnd
    End If
    
    sErr = CloseSocket(currSock)
    If sErr <> 0 Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
    
    If currType = PROTOCOL.TCP And currClient <> NULL_PTR Then
        sErr = Shutdown(currClient, SD_FLAGs.SD_BOTH)
        If sErr = SOCKET_ERROR Then
            sErr = WSAGetLastError
            errType = winsock
            GoTo ErrHnd
        End If
        
        sErr = CloseSocket(currClient)
        If sErr <> 0 Then
            sErr = WSAGetLastError
            errType = winsock
            GoTo ErrHnd
        End If
    End If
    
    EN
    Exit Function
ErrHnd:
    EN
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
''
'Description:
'   Starts a WinSock Client connected to the specified Host and Port
'Params:
'   Proto As PROTOCOL  -  This the Protocal used for the socket. Valid Values
'                         are UPD or TCP.
'   Host  As String    -  This is the Host the Socket will Connect to. Valid
'                         Values:
'                             IPV6 Address such as "2001:0db8:85a3:0000:0000:8a2e:0370:7334"
'                             IPV4 Address such as "192.0. 2.146"
'                             Host/PC Name such as "google.com"
'   Port  As String    -  This is the Port the Socket will Connect to. Valid
'                         Values:
'                             Service Name such as "http"
'                             Port Number such as  "80"
Public Function StartClient(Proto As PROTOCOL, Optional Host As String = vbNullString, Optional Port As String = "0")
    ST
    
    Dim sErr As Long
    Dim errType As ERR_TYPE
    Dim Source As String: Source = "StartClient"
    Dim Hints As ADDRINFO: Hints.ai_family = AF.AF_UNSPEC
    Dim Result As ADDRINFO
#If VBA7 Then
    Dim pResult As LongPtr
#Else
    Dim pResult As Long
#End If
    If currSock <> NULL_PTR Then
        If isServer Then
            sErr = CUSTOMERROR.IS_SERVER
        Else
            sErr = CUSTOMERROR.IS_CLIENT
        End If
        errType = custom
        GoTo ErrHnd
    End If
    
    If Proto = TCP Then
        Hints.ai_sockType = sock.SOCK_STREAM
        Hints.ai_protocol = IPPROTO.IPPROTO_TCP
        
        sErr = GetAddrInfo(Host, Port, VarPtr(Hints), pResult)
        If sErr <> 0 Then
            sErr = WSAGetLastError
            errType = winsock
            GoTo ErrHnd
        Else
            CopyMemory VarPtr(Result), pResult, LenB(Result)
        End If
    ElseIf Proto = UDP Then
        Result.ai_family = AF.AF_UNSPEC
        Result.ai_sockType = sock.SOCK_DGRAM
        Result.ai_protocol = IPPROTO.IPPROTO_UDP
    Else
        sErr = CUSTOMERROR.INVAILD_PROTO
        errType = custom
        GoTo ErrHnd
    End If
    
    currSock = CreateSocket(Result.ai_family, Result.ai_sockType, Result.ai_protocol)
    If currSock = INVALID_SOCKET Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
    
    If Proto = TCP Then
        sErr = Connect(currSock, Result.ai_addr, Result.ai_addrlen)
        If sErr = SOCKET_ERROR Then
            sErr = WSAGetLastError
            errType = winsock
            GoTo ErrHnd
        End If
    End If
    
    FreeAddrInfo pResult
    currType = Proto
    isServer = False
    Exit Function
ErrHnd:
    EN
    If pResult <> NULL_PTR Then FreeAddrInfo pResult
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
''
'Description:
'   Closes the Current Client.
Public Function CloseClient()
    Dim sErr As Long
    Dim errType As ERR_TYPE
    Dim Source As String: Source = "CloseClient"
    
    If currSock = NULL_PTR Then
        sErr = CUSTOMERROR.NOT_STARTED
        errType = custom
        GoTo ErrHnd
    End If
    If isServer Then
        sErr = CUSTOMERROR.IS_SERVER
        errType = custom
        GoTo ErrHnd
    End If
    
    sErr = Shutdown(currSock, SD_FLAGs.SD_BOTH)
    If sErr = SOCKET_ERROR Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
        
    sErr = CloseSocket(currSock)
    If sErr <> 0 Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
    
    EN
    Exit Function
ErrHnd:
    EN
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
''
'Description:
'   This function Blocks until the Current Client recieves Data.
'Params:
'   Optional WCHAR As Boolean  -  Whether the String Supplied should be WCHAR or not
'Return:
'   The String that was recieved from the Connected Client.
Public Function Receive(Optional WCHAR As Boolean = False) As String
    Receive = BufferToString(ReceiveBytes, WCHAR)
End Function
''
'Description:
'   This function Blocks until the Current Client recieves Data.
'Return:
'   The Bytes that was recieved from the Connected Client.
Public Function ReceiveBytes() As Byte()
    Dim iResult As Long
    Dim sErr As Long
    Dim errType As ERR_TYPE
    Dim Source As String: Source = "Receive"
    Dim rBuff(MAX_PACKET) As Byte
    Dim rBuff2() As Byte
    Dim out() As Byte
    Dim i As Long
    
    If currSock = NULL_PTR Then
        sErr = CUSTOMERROR.NOT_STARTED
        errType = custom
        GoTo ErrHnd
    End If
    If currType = PROTOCOL.UDP Then
        sErr = CUSTOMERROR.PROTO_TCP
        errType = custom
        GoTo ErrHnd
    End If
    
    If isServer Then
        If currClient = NULL_PTR Then
            sErr = CUSTOMERROR.NOT_ACCEPTED
            errType = custom
            GoTo ErrHnd
        End If
        iResult = Recv(currClient, VarPtr(rBuff(0)), MAX_PACKET - 1, MSG_FLAGs.Default)
    Else
        iResult = Recv(currSock, VarPtr(rBuff(0)), MAX_PACKET - 1, MSG_FLAGs.Default)
    End If
    
    If iResult = SOCKET_ERROR Then
        sErr = WSAGetLastError
        If sErr = 10040 Then
            rBuff2 = ReceiveBytes()
        Else
            errType = winsock
            GoTo ErrHnd
        End If
    End If
    
    
    If iResult <> 0 And iResult <> SOCKET_ERROR Then
        ReDim out(0 To iResult - 1)
        For i = 0 To iResult - 1
            out(i) = rBuff(i)
        Next i
    ElseIf sErr = 10040 Then
        ReDim out(0 To MAX_PACKET - 1)
        For i = 0 To MAX_PACKET - 1
            out(i) = rBuff(i)
        Next i
    End If
    
    If (Not Not rBuff2) <> 0 Then
        ReDim Preserve out(0 To UBound(out) + UBound(rBuff2))
        For i = 0 To UBound(rBuff2)
            out(i + UBound(rBuff2)) = rBuff2(i)
        Next i
    End If
    
    ReceiveBytes = out
    Exit Function
ErrHnd:
    EN
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
''
'Description:
'   This function Blocks until the Current UDP Client/Server recieves Data.
'Params:
'   Optional ByRef Host As String    -  The String to fill with the Sender's Host Name.
'   Optional ByRef Port As String    -  The String to fill with the Sender's Port Number.
'   Optional       WCHAR As Boolean  -  Whether the String Supplied should be WCHAR or not
'Return:
'   ReceiveFrom As String  -  The String that was recieved from the UPD Client/Server.
Public Function ReceiveFrom(Optional ByRef Host As String, Optional ByRef Port As String, Optional WCHAR As Boolean = False) As String
    Dim rHost As String
    Dim rPort As String
    
    ReceiveFrom = BufferToString(ReceiveBytesFrom(rHost, rPort), WCHAR)
    Host = rHost
    Port = rPort
End Function
''
'Description:
'   This function Blocks until the Current UDP Client/Server recieves Data.
'Params:
'   Optional ByRef Host As String  -  The String to fill with the Sender's Host Name.
'   Optional ByRef Port As String  -  The String to fill with the Sender's Port Number.
'Return:
'   ReceiveBytesFrom As Byte()     -  The Bytes that was recieved from the UPD Client/Server.
Public Function ReceiveBytesFrom(Optional ByRef Host As String, Optional ByRef Port As String) As Byte()
    Dim iResult As Long
    Dim sErr As Long
    Dim errType As ERR_TYPE
    Dim Source As String: Source = "ReceiveFrom"
    Dim rBuff(MAX_PACKET) As Byte
    Dim out() As Byte
    Dim addr As sockaddr_in6
#If VBA7 Then
    Dim addrPtr As LongPtr: addrPtr = VarPtr(addr)
#Else
    Dim addrPtr As Long: addrPtr = VarPtr(addr)
#End If
    Dim addrLen As Long: addrLen = LenB(addr)
    Dim rHost(NI_MAXHOST) As Byte
    Dim rPort(NI_MAXSERV) As Byte
    Dim i As Long
    
    If currSock = NULL_PTR Then
        sErr = CUSTOMERROR.NOT_STARTED
        errType = custom
        GoTo ErrHnd
    End If
    If currType = PROTOCOL.TCP Then
        sErr = CUSTOMERROR.PROTO_UDP
        errType = custom
        GoTo ErrHnd
    End If
    
    iResult = RecvFrom(currSock, VarPtr(rBuff(0)), MAX_PACKET, MSG_FLAGs.Default, addrPtr, VarPtr(addrLen))
    If iResult = SOCKET_ERROR Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
    
    sErr = GetNameInfo(addrPtr, addrLen, VarPtr(rHost(0)), NI_MAXHOST, VarPtr(rPort(0)), NI_MAXSERV, NameInfo_Flags.NI_NUMERICSERV)
    If sErr <> 0 Then
        errType = winsock
        GoTo ErrHnd
    End If
    Host = BufferToString(rHost, False)
    Port = BufferToString(rPort, False)
    
    ReDim out(0 To iResult - 1)
    For i = 0 To UBound(out)
        out(i) = rBuff(i)
    Next i
    
    ReceiveBytesFrom = out
    Exit Function
ErrHnd:
    EN
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
''
'Description:
'   Sends the Supplied String to a Connected Client.
'Params:
'            Value As String   -  The String to Send
'   Optional WCHAR As Boolean  -  Whether the String Supplied should be WCHAR or not
'Returns:
'   Send as Long  -  The Number of Bytes Transmitted
Public Function Send(Value As String, Optional WCHAR As Boolean = False) As Long
    Dim rBuff() As Byte
    
    rBuff = StringToBuffer(Value, WCHAR)
    
    Send = SendBytes(rBuff)
End Function
''
'Description:
'   Sends the Supplied Byte Array to a Connected Client.
'Params:
'   rBuff() As Byte  -  The Byte Array to be Transmitted.
'Returns:
'   Send as Long  -  The Number of Bytes Transmitted
Public Function SendBytes(rBuff() As Byte) As Long
    Dim iResult As Long
    Dim sErr As Long
    Dim errType As ERR_TYPE
    Dim Source As String: Source = "Send"
    Dim bLen As Long
    
    bLen = UBound(rBuff) + 1
    
    If currSock = NULL_PTR Then
        sErr = CUSTOMERROR.NOT_STARTED
        errType = custom
        GoTo ErrHnd
    End If
    If currType = PROTOCOL.UDP Then
        sErr = CUSTOMERROR.PROTO_TCP
        errType = custom
        GoTo ErrHnd
    End If
    
    If isServer Then
        iResult = SocketSend(currClient, VarPtr(rBuff(0)), bLen, MSG_FLAGs.Default)
    Else
        iResult = SocketSend(currSock, VarPtr(rBuff(0)), bLen, MSG_FLAGs.Default)
    End If
    
    If iResult = SOCKET_ERROR Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
    
    SendBytes = iResult
    Exit Function
ErrHnd:
    EN
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
''
'Description:
'   Sends the Supplied String to a Specified UDP Client/Server.
'Params:
'            Value As String   -  The String to Send
'            Host  As String   -  This is the Host the UDP Client/Server will Send to. Valid
'                                 Values:
'                                     IPV6 Address such as "2001:0db8:85a3:0000:0000:8a2e:0370:7334"
'                                     IPV4 Address such as "192.0. 2.146"
'                                     Host/PC Name such as "google.com"
'            Port  As String   -  This is the Port the UDP Client/Server will Send to. Valid
'                                 Values:
'                                     Service Name such as "http"
'                                     Port Number such as  "80"
'   Optional WCHAR As Boolean  -  Whether the String Supplied should be WCHAR or not.
'Returns:
'   Send as Long  -  The Number of Bytes Transmitted.
Public Function SendTo(Value As String, Host As String, Port As String, Optional WCHAR As Boolean = False) As Long
    Dim rBuff() As Byte
    
    rBuff = StringToBuffer(Value, WCHAR)
    
    SendTo = SendBytesTo(rBuff, Host, Port)
End Function
''
'Description:
'   Sends the Supplied Byte Array to a Specified UDP Client/Server.
'Params:
'   rBuff() As Byte   -  The String to Send
'   Host  As String   -  This is the Host the UDP Client/Server will Send to. Valid
'                        Values:
'                            IPV6 Address such as "2001:0db8:85a3:0000:0000:8a2e:0370:7334"
'                            IPV4 Address such as "192.0. 2.146"
'                            Host/PC Name such as "google.com"
'   Port  As String   -  This is the Port the UDP Client/Server will Send to. Valid
'                        Values:
'                            Service Name such as "http"
'                            Port Number such as  "80"
'Returns:
'   Send as Long  -  The Number of Bytes Transmitted.
Public Function SendBytesTo(rBuff() As Byte, Host As String, Port As String) As Long
    Dim iResult As Long
    Dim sErr As Long
    Dim errType As ERR_TYPE
    Dim Source As String: Source = "SendTo"
    Dim bLen As Long
    Dim Hints As ADDRINFO
        Hints.ai_family = AF.AF_UNSPEC
        Hints.ai_sockType = sock.SOCK_DGRAM
        Hints.ai_protocol = IPPROTO.IPPROTO_UDP
    Dim Result As ADDRINFO
#If VBA7 Then
    Dim pResult As LongPtr
#Else
    Dim pResult As Long
#End If
    
    bLen = UBound(rBuff) + 1
    
    If currSock = NULL_PTR Then
        sErr = CUSTOMERROR.NOT_STARTED
        errType = custom
        GoTo ErrHnd
    End If
    If currType = PROTOCOL.TCP Then
        sErr = CUSTOMERROR.PROTO_UDP
        errType = custom
        GoTo ErrHnd
    End If
    
    sErr = GetAddrInfo(Host, Port, VarPtr(Hints), pResult)
    If sErr <> 0 Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    Else
        CopyMemory VarPtr(Result), pResult, LenB(Result)
    End If
    
    iResult = SocketSendTo(currSock, VarPtr(rBuff(0)), bLen, MSG_FLAGs.Default, Result.ai_addr, Result.ai_addrlen)
    If iResult = SOCKET_ERROR Then
        sErr = WSAGetLastError
        errType = winsock
        GoTo ErrHnd
    End If
    
    FreeAddrInfo pResult
    SendBytesTo = iResult
    Exit Function
ErrHnd:
    EN
    If pResult <> NULL_PTR Then FreeAddrInfo pResult
    If errType = winsock Then WinSockErrorMsg sErr, Source
    If errType = custom Then CustomErrorMsg Source, sErr
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Error Handler                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function WinSockErrorMsg(ErrNum As Long, Source As String)
    Dim iResult As Long
    Dim errStr As String * 1085
    Dim str As String
    
    iResult = FormatMessage( _
                FORMAT_MESSAGE_MAX_WIDTH_MASK Or FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                NULL_PTR, _
                ErrNum, _
                DEFAULT_LNG, _
                errStr, _
                1085, _
                NULL_PTR _
              )
              
    str = VBA.Left(errStr, iResult)
    ErrNum = vbObjectError + ErrNum
    
    err.Raise ErrNum, "WinSock2_2." & Source, str
End Function
Private Function CustomErrorMsg(Source As String, ErrNum As CUSTOMERROR)
    Dim Msg As String
    
    Select Case ErrNum
        Case CUSTOMERROR.PROTO_TCP
            Msg = "This Function is Only valid for TCP Connections."
        Case CUSTOMERROR.PROTO_UDP
            Msg = "This Function is Only valid for UDP Connections."
        Case CUSTOMERROR.INVAILD_PROTO
            Msg = "Invalid Protocal, Options are UDP or TCP."
        Case CUSTOMERROR.IS_CLIENT
            Msg = "Client Has Been Started!! Please run only one Server/Client at a time!!"
        Case CUSTOMERROR.IS_SERVER
            Msg = "Server Has Been Started!! Please run only one Server/Client at a time!!"
        Case CUSTOMERROR.ACCEPTED
            Msg = "A Socketed Connection is already active!! You MUST close the current connection before accepting a new one."
        Case CUSTOMERROR.NOT_ACCEPTED
            Msg = "Server Has not Accepted a Connection."
        Case CUSTOMERROR.NOT_STARTED
            Msg = "Server/Client Has not been Started."
    End Select
    
    err.Raise vbObjectError + ErrNum, "WinSock2_2." & Source, Msg
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Class Functions                                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    currSock = NULL_PTR
    currClient = NULL_PTR
End Sub

Private Sub Class_Terminate()
    EN
End Sub
