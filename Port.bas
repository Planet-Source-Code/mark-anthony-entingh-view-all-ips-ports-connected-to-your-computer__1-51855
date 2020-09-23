Attribute VB_Name = "TCPIP"
Public arrTCPBuffer() As Byte
Public TCPBuffer() As MIB_TCPROW
Public arrUDPBuffer() As Byte
Public UDPBuffer() As MIB_UDPROW


Public Type MIB_TCPSTATS
    dwRtoAlgorithm  As Long '// timeout algorithm
    dwRtoMin        As Long '// minimum timeout
    dwRtoMax        As Long '// maximum timeout
    dwMaxConn       As Long '// maximum connections
    dwActiveOpens   As Long '// active opens
    dwPassiveOpens  As Long '// passive opens
    dwAttemptFails  As Long '// failed attempts
    dwEstabResets   As Long '// establised connections reset
    dwCurrEstab     As Long '// established connections
    dwInSegs        As Long '// segments received
    dwOutSegs       As Long '// segment sent
    dwRetransSegs   As Long '// segments retransmitted
    dwInErrs        As Long '// incoming errors
    dwOutRsts       As Long '// outgoing resets
    dwNumConns      As Long '// cumulative connections                                                                                   '// cumulative connections
End Type

Public Type MIB_IPSTATS
    dwForwarding    As Long     '// IP forwarding enabled or disabled
    dwDefaultTTL    As Long     '// default time-to-live
    dwInReceives    As Long     '// datagrams received
    dwInHdrErrors   As Long     '// received header errors
    dwInAddrErrors  As Long     '// received address errors
    dwForwDatagrams As Long     '// datagrams forwarded
    dwInUnknownProtos As Long   '// datagrams with unknown protocol
    dwInDiscards    As Long     '// received datagrams discarded
    dwInDelivers    As Long     '// received datagrams delivered
    dwOutRequests   As Long     '//
    dwRoutingDiscards As Long   '//
    dwOutDiscards   As Long     '// sent datagrams discarded
    dwOutNoRoutes   As Long     '// datagrams for which no route exists
    dwReasmTimeout  As Long     '// datagrams for which all frags did not arrive
    dwReasmReqds    As Long     '// datagrams requiring reassembly
    dwReasmOks      As Long     '// successful reassemblies
    dwReasmFails    As Long     '// failed reassemblies
    dwFragOks       As Long     '// successful fragmentations
    dwFragFails     As Long     '// failed fragmentations
    dwFragCreates   As Long     '// datagrams fragmented
    dwNumIf         As Long     '// number of interfaces on computer
    dwNumAddr       As Long     '// number of IP address on computer
    dwNumRoutes     As Long     '// number of routes in routing table
End Type

Public Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Public TcpTableRow As MIB_TCPROW

Public Type MIB_UDPROW
    dwLocalAddr As Long
    dwLocalPort As Long
End Type

Public UdpTableRow As MIB_UDPROW

Public Const ERROR_BUFFER_OVERFLOW = 111&
Public Const ERROR_INVALID_PARAMETER = 87
Public Const ERROR_NO_DATA = 232&
Public Const ERROR_NOT_SUPPORTED = 50&
'
Public Const MIB_TCP_STATE_CLOSED = 1
Public Const MIB_TCP_STATE_LISTEN = 2
Public Const MIB_TCP_STATE_SYN_SENT = 3
Public Const MIB_TCP_STATE_SYN_RCVD = 4
Public Const MIB_TCP_STATE_ESTAB = 5
Public Const MIB_TCP_STATE_FIN_WAIT1 = 6
Public Const MIB_TCP_STATE_FIN_WAIT2 = 7
Public Const MIB_TCP_STATE_CLOSE_WAIT = 8
Public Const MIB_TCP_STATE_CLOSING = 9
Public Const MIB_TCP_STATE_LAST_ACK = 10
Public Const MIB_TCP_STATE_TIME_WAIT = 11
Public Const MIB_TCP_STATE_DELETE_TCB = 12

Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
Public Declare Function GetTcpStatistics Lib "iphlpapi.dll" (pStats As MIB_TCPSTATS) As Long
Public Declare Function GetIpStatistics Lib "iphlpapi.dll" (pStats As MIB_IPSTATS) As Long
'Public Declare Function GetIcmpStatistics Lib "iphlpapi.dll" (pStats As MIB_TCPSTATS) As Long
'Public Declare Function GetUdpStatistics Lib "iphlpapi.dll" (pStats As MIB_TCPSTATS) As Long
Public Declare Function GetTcpTable Lib "iphlpapi.dll" _
(ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function SetTcpEntry Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPROW) As Long
Public Declare Function GetUdpTable Lib "iphlpapi.dll" _
(ByRef pUdpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long

Public Function GetIpFromLong(lngIPAddress As Long) As String
    Dim arrIpParts(3) As Byte
    CopyMem arrIpParts(0), lngIPAddress, 4
    GetIpFromLong = CStr(arrIpParts(0)) & "." & CStr(arrIpParts(1)) & "." & CStr(arrIpParts(2)) & "." & CStr(arrIpParts(3))
End Function

Public Function GetState(lngState As Long) As String
    Select Case lngState
        Case MIB_TCP_STATE_CLOSED: GetState = "CLOSED"
        Case MIB_TCP_STATE_LISTEN: GetState = "LISTEN"
        Case MIB_TCP_STATE_SYN_SENT: GetState = "SYN_SENT"
        Case MIB_TCP_STATE_SYN_RCVD: GetState = "SYN_RCVD"
        Case MIB_TCP_STATE_ESTAB: GetState = "ESTAB"
        Case MIB_TCP_STATE_FIN_WAIT1: GetState = "FIN_WAIT1"
        Case MIB_TCP_STATE_FIN_WAIT2: GetState = "FIN_WAIT2"
        Case MIB_TCP_STATE_CLOSE_WAIT: GetState = "CLOSE_WAIT"
        Case MIB_TCP_STATE_CLOSING: GetState = "CLOSING"
        Case MIB_TCP_STATE_LAST_ACK: GetState = "LAST_ACK"
        Case MIB_TCP_STATE_TIME_WAIT: GetState = "TIME_WAIT"
        Case MIB_TCP_STATE_DELETE_TCB: GetState = "DELETE_TCB"
    End Select
End Function

Public Function GetTcpPortNumber(DWord As Long) As Long
    GetTcpPortNumber = DWord / 256 + (DWord Mod 256) * 256
End Function

Public Function GetUdpPortNumber(DWord As Long) As Long
    GetUdpPortNumber = DWord / 256 + (DWord Mod 256) * 256
End Function

Public Sub KillPort(index)
    Dim TcpTableRow As MIB_TCPROW
    Dim lngRetValue As Long
    '
        '
        TcpTableRow = TCPBuffer(index)
        '
        TcpTableRow.dwState = MIB_TCP_STATE_DELETE_TCB
        '
        lngRetValue = SetTcpEntry(TcpTableRow)
        '
        If lngRetValue = 0 Then
            'item killed
        Else
            'item NOT killed
        End If
End Sub
