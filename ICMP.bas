Attribute VB_Name = "ICMP"
' WSock32 UDTs

Type Inet_address
    Byte4 As String * 1
    Byte3 As String * 1
    Byte2 As String * 1
    Byte1 As String * 1
End Type

Public IPLong As Inet_address

Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Type IP_OPTION_INFORMATION
    TTL As Byte                   ' Time to Live (used for traceroute)
    Tos As Byte                   ' Type of Service (usually 0)
    Flags As Byte                 ' IP header Flags (usually 0)
    OptionsSize As Long           ' Size of Options data (usually 0, max 40)
    OptionsData As String * 128   ' Options data buffer
End Type

Public pIPo As IP_OPTION_INFORMATION

Type IP_ECHO_REPLY
    Address(0 To 3) As Byte           ' Replying Address
    Status As Long                    ' Reply Status
    RoundTripTime As Long             ' Round Trip Time in milliseconds
    DataSize As Integer               ' reply data size
    Reserved As Integer               ' for system use
    data As Long                      ' pointer to echo data
    Options As IP_OPTION_INFORMATION  ' Reply Options
End Type

Public pIPe As IP_ECHO_REPLY

' WSock32 Subroutines and Functions

Declare Function gethostname Lib "wsock32.dll" (ByVal hostname$, HostLen&) As Long
Declare Function gethostbyname& Lib "wsock32.dll" (ByVal hostname$)
Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAData As WSAdata) As Long
Declare Function WSACleanup Lib "wsock32.dll" () As Long

' Kernel32 Subroutines and Functions

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

' ICMP Subroutines and Functions

    ' IcmpCreateFile will return a file handle
    
Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
     
     ' Pass the handle value from IcmpCreateFile to the IcmpCloseHandle.  It will return
     ' a boolean value indicating whether or not it closed successfully.
     
Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean

    ' IcmpHandle returned from IcmpCreateFile
    ' DestAddress is a pointer to the first entry in the hostent.h_addr_list
    ' RequestData is a null-terminated 64-byte string filled with ASCII 170 characters
    ' RequestSize is 64-bytes
    ' RequestOptions is a NULL at this time
    ' ReplyBuffer
    ' ReplySize
    ' Timeout is the timeout in milliseconds

Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, _
    ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, _
     ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

