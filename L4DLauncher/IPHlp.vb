Option Strict On
Imports System.Net
Imports System.Net.Sockets
Imports System.Runtime.InteropServices

Public Class IPHlp

    Private Const ICMP_SUCCESS As Long = 0
    Private Const ICMP_STATUS_BUFFER_TO_SMALL As Int32 = 11001
    'Buffer Too Small
    Private Const ICMP_STATUS_DESTINATION_NET_UNREACH As Int32 = 11002
    'Destination Net Unreachable
    Private Const ICMP_STATUS_DESTINATION_HOST_UNREACH As Int32 = 11003 ' Destination Host Unreachable
    Private Const ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH As Int32 = 11004 'Destination Protocol Unreachable
    Private Const ICMP_STATUS_DESTINATION_PORT_UNREACH As Int32 = 11005 'Destination Port Unreachable
    Private Const ICMP_STATUS_NO_RESOURCE As Int32 = 11006 'No Resources
    Private Const ICMP_STATUS_BAD_OPTION As Int32 = 11007 'Bad Option
    Private Const ICMP_STATUS_HARDWARE_ERROR As Int32 = 11008
    'Hardware Error
    Private Const ICMP_STATUS_LARGE_PACKET As Int32 = 11009 'Packet Too Big
    Private Const ICMP_STATUS_REQUEST_TIMED_OUT As Int32 = 11010
    'Request Timed Out
    Private Const ICMP_STATUS_BAD_REQUEST As Int32 = 11011 'Bad Request
    Private Const ICMP_STATUS_BAD_ROUTE As Int32 = 11012 'Bad Route
    Private Const ICMP_STATUS_TTL_EXPIRED_TRANSIT As Int32 = 11013
    'TimeToLive Expired Transit
    Private Const ICMP_STATUS_TTL_EXPIRED_REASSEMBLY As Int32 = 11014
    'TimeToLive Expired Reassembly
    Private Const ICMP_STATUS_PARAMETER As Int32 = 11015 'Parameter Problem
    Private Const ICMP_STATUS_SOURCE_QUENCH As Int32 = 11016 'Source Quench
    Private Const ICMP_STATUS_OPTION_TOO_BIG As Int32 = 11017 'Option Too Big
    Private Const ICMP_STATUS_BAD_DESTINATION As Int32 = 11018 'Bad Destination
    Private Const ICMP_STATUS_NEGOTIATING_IPSEC As Int32 = 11032
    'Negotiating IPSEC
    Private Const ICMP_STATUS_GENERAL_FAILURE As Int32 = 11050
    'General Failure

    Public Shared Function ReturnMachines(ByVal sIPBase As String, ByVal nStart As Integer, ByVal nEnd As Integer, ByVal nTimeout As Integer, ByVal nPort As Integer) As PingObject
        Dim oPingObject As New PingObject
        oPingObject.nStart = nStart
        oPingObject.nEnd = nEnd
        Dim sIP As String
        For i = nStart To nEnd
            If PortConnect.bCanceled Then
                Exit For
            End If
            sIP = sIPBase & "." & CStr(i)
            'IPHlp.Ping(oPingObject, sIP, Nothing, reply, nTimeout)
            Try
                Dim IP As IPAddress = Dns.GetHostByName(sIP).AddressList(0)
                Dim connector = New PortConnect
                For p = nPort To nPort ' + 30
                    If connector.Connect(oPingObject, IP, p, nTimeout) Then
                        Exit For
                    End If
                Next p
            Catch
                Continue For
            End Try
        Next i
        Return oPingObject
    End Function

    Public Shared Function Localhost(Optional ByVal nBlocks As Integer = 4) As String
        Dim LocalHostName As String
        Dim i As Integer
        LocalHostName = Dns.GetHostName()
        Dim ipEnter As IPHostEntry = Dns.GetHostByName(LocalHostName)
        Dim IpAdd() As IPAddress = ipEnter.AddressList
        Localhost = ""
        Dim sSplit As String() = IpAdd.GetValue(0).ToString.Split(CChar("."))
        For i = 0 To Math.Min(UBound(sSplit), nBlocks - 1)
            If i > 0 Then Localhost &= "."
            Localhost &= sSplit(i)
        Next
    End Function

    Public Shared Function Ping(ByRef oPingObject As PingObject, ByVal sIP As String, Optional ByRef IP As IPAddress = Nothing, Optional ByRef Reply As String = "", Optional ByRef RoundTrip As Int32 = 0) As Boolean
        IP = Dns.GetHostByName(sIP).AddressList(0)
        Return Ping(oPingObject, IP, Reply, RoundTrip)
    End Function

    Public Shared Function Ping(ByRef oPingObject As PingObject, ByVal IP As IPAddress, Optional ByRef Reply As String = "", Optional ByRef RoundTrip As Int32 = 0) As Boolean
        Dim ICMPHandle As IntPtr
        Dim iIP As Int32
        Dim sData As String
        Dim oICMPOptions As New IPHlp.ICMP_OPTIONS
        Dim ICMPReply As ICMP_ECHO_REPLY
        Dim iReplies As Int32

        Ping = False

        ICMPHandle = IcmpCreateFile

        iIP = System.BitConverter.ToInt32(IP.GetAddressBytes, 0)
        sData = "x" 'Placeholder?
        oICMPOptions.Ttl = 255

        iReplies = IcmpSendEcho(ICMPHandle, iIP, sData, sData.Length, oICMPOptions, ICMPReply, Marshal.SizeOf(ICMPReply), 30)

        Reply = EvaluatePingResponse(oPingObject, IP, ICMPReply.Status)
        RoundTrip = ICMPReply.RoundTripTime

        If ICMPReply.Status = 0 Then
            Ping = True
        End If

        ICMPHandle = Nothing
        IP = Nothing
        oICMPOptions = Nothing
        ICMPReply = Nothing

    End Function

    Public Shared Function EvaluatePingResponse(ByRef oPingObject As PingObject, ByVal IP As IPAddress, ByVal PingResponse As Long) As String
        Select Case PingResponse
            'success
            Case ICMP_SUCCESS
                EvaluatePingResponse = "Success!"
                MsgBox("Success!")
                'AddMachine(oPingObject, IP)
            Case ICMP_STATUS_BUFFER_TO_SMALL
                EvaluatePingResponse = "Buffer Too Small"
            Case ICMP_STATUS_DESTINATION_NET_UNREACH
                EvaluatePingResponse = "Destination Net Unreachable"
            Case ICMP_STATUS_DESTINATION_HOST_UNREACH
                EvaluatePingResponse = "Destination Host Unreachable"
            Case ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH
                EvaluatePingResponse = "Destination Protocol"
            Case ICMP_STATUS_DESTINATION_PORT_UNREACH
                EvaluatePingResponse = "Destination Port Unreachable"
            Case ICMP_STATUS_NO_RESOURCE
                EvaluatePingResponse = "No Resources"
            Case ICMP_STATUS_BAD_OPTION
                EvaluatePingResponse = "Bad Option"
            Case ICMP_STATUS_HARDWARE_ERROR
                EvaluatePingResponse = "Hardware Error"
            Case ICMP_STATUS_LARGE_PACKET
                EvaluatePingResponse = "Packet Too Big"
            Case ICMP_STATUS_REQUEST_TIMED_OUT
                EvaluatePingResponse = "Request Timed Out"
            Case ICMP_STATUS_BAD_REQUEST
                EvaluatePingResponse = "Bad Request"
            Case ICMP_STATUS_BAD_ROUTE
                EvaluatePingResponse = "Bad Route"
            Case ICMP_STATUS_TTL_EXPIRED_TRANSIT
                EvaluatePingResponse = "TimeToLive Expired Transit"
            Case ICMP_STATUS_TTL_EXPIRED_REASSEMBLY
                EvaluatePingResponse = "TimeToLive Expired Reassembly"
            Case ICMP_STATUS_PARAMETER
                EvaluatePingResponse = "Parameter Problem"
            Case ICMP_STATUS_SOURCE_QUENCH
                EvaluatePingResponse = "Source Quench"
            Case ICMP_STATUS_OPTION_TOO_BIG
                EvaluatePingResponse = "Option Too Big"
            Case ICMP_STATUS_BAD_DESTINATION
                EvaluatePingResponse = "Bad Destination"
            Case ICMP_STATUS_NEGOTIATING_IPSEC
                EvaluatePingResponse = "Negotiating IPSEC"
            Case ICMP_STATUS_GENERAL_FAILURE
                EvaluatePingResponse = "General Failure"
                'unknown error occurred
            Case Else
                EvaluatePingResponse = "Unknown Response"
        End Select

    End Function

    Declare Auto Function IcmpCreateFile Lib "ICMP.DLL" () As IntPtr

    Declare Auto Function IcmpSendEcho Lib "ICMP.DLL" _
    (ByVal IcmpHandle As IntPtr, _
    ByVal DestinationAddress As Integer, _
    ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    ByRef RequestOptions As ICMP_OPTIONS, _
    ByRef ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Integer, _
    ByVal Timeout As Integer) As Integer

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> Public Structure ICMP_OPTIONS
        Public Ttl As Byte
        Public Tos As Byte
        Public Flags As Byte
        Public OptionsSize As Byte
        Public OptionsData As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> Public Structure ICMP_ECHO_REPLY
        'Public Address As Integer
        Public Address As Integer
        Public Status As Integer
        Public RoundTripTime As Integer
        Public DataSize As Short
        Public Reserved As Short
        Public DataPtr As IntPtr
        Public Options As ICMP_OPTIONS
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=250)> _
        Public Data As String
    End Structure

End Class
