Imports System.Net.Sockets
Imports System.Net
Imports System.Text
Imports System.Collections.Specialized
Imports L4D_launcher.Launcher

Public Class PortConnect

    Private Shared nTotalIPs As Integer
    Private Shared nMaxIPs As Integer
    Private Shared nFirstIP As Integer = 0
    Private Shared nLastIP As Integer = 255
    Private Shared nStepIP As Integer = 1

    Private Shared _window As Launcher

    Delegate Function GetPingIPs(ByVal sIPBase As String, ByVal nStart As Integer, ByVal nEnd As Integer, ByVal nTimeout As Integer, ByVal nPort As Integer) As PingObject
    Private Shared ThreadPing As GetPingIPs = AddressOf IPHlp.ReturnMachines
    Private Shared ThreadPingCallback As AsyncCallback = AddressOf ThreadPingComplete

    'Private Shared connectionMessage As String = Chr(255) & Chr(255) & Chr(255) & Chr(255) & "TSource Engine Query"
    Private Shared connectionMessage As String = "ÿÿÿÿTSource Engine Query" & Chr(0)

    Private connection As Socket

    Private _readBuffer As Byte(), _responseString As String
    Private _offset As Integer
    Private _params As StringDictionary

    Private bTimedout As Boolean
    Public Shared bCanceled As Boolean

    WithEvents Timeout As System.Timers.Timer

    Public Sub New()
        _params = New StringDictionary()
    End Sub

    Public Shared Sub Test(ByVal sIP As String)
        Dim connector As New PortConnect
        Dim oPingObject As New PingObject
        connector.Connect(oPingObject, IPAddress.Parse(sIP), 27015, 1000)
        System.Threading.Thread.Sleep(1000)
        End
    End Sub

    Public Shared Sub Connect(ByVal window As Launcher, ByVal bBroadcast As Boolean, ByVal nPort As Integer, ByVal nTimeout As Integer)
        bCanceled = False
        _window = window
        nTotalIPs = 0
        If bBroadcast Then
            nMaxIPs = 1
            ThreadPing.BeginInvoke("255.255.255", 255, 255, nTimeout, nPort, ThreadPingCallback, Nothing)
        Else
            Dim sLocalhost As String = IPHlp.Localhost(3)
            nMaxIPs = nLastIP - nFirstIP
            For i = nFirstIP To nLastIP Step nStepIP
                If bCanceled Then
                    Exit For
                End If
                Dim nThisStep As Integer = Math.Min(i + nStepIP - 1, nLastIP)
                ThreadPing.BeginInvoke(sLocalhost, i, nThisStep, nTimeout, nPort, ThreadPingCallback, Nothing)
            Next i
        End If
    End Sub

    Public Shared Sub Cancel()
        bCanceled = True
        [Delegate].RemoveAll(_listDelegate, _listDelegate)
    End Sub

    Private Shared Sub ThreadPingComplete(ByVal iAsyncResult As IAsyncResult)
Retry:
        Try
            SyncLock _window
                Dim oPingObject As PingObject
                oPingObject = ThreadPing.EndInvoke(iAsyncResult)
                nTotalIPs += oPingObject.nEnd - oPingObject.nStart + 1
                Dim handler1 As New UpdateMachinesHandler(AddressOf _window.UpdateMachines)
                Dim args1() As Object = {oPingObject.cMachines, nTotalIPs >= nMaxIPs}
                _window.BeginInvoke(handler1, args1)
                Dim handler2 As New UpdateProgressHandler(AddressOf _window.UpdateProgress)
                Dim args2() As Object = {Math.Min(nTotalIPs, nMaxIPs), nMaxIPs}
                _window.BeginInvoke(handler2, args2)
            End SyncLock
        Catch ex As InvalidOperationException
            Threading.Thread.Sleep(500)
            GoTo Retry
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    Public Function Connect(ByRef oPingObject As PingObject, ByVal IP As IPAddress, ByVal nPort As Integer, ByVal nTimeout As Integer) As Boolean
        Dim bSuccess As Boolean
        Dim remoteIpEndPoint As New IPEndPoint(IP, nPort)
        Dim remoteEndPoint As EndPoint = DirectCast(remoteIpEndPoint, EndPoint)

        bTimedout = False
        connection = New Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)
        connection.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, nTimeout)
        If IP.ToString = "255.255.255.255" Then
            connection.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Broadcast, 1)
        End If
        connection.SendTo(System.Text.Encoding.[Default].GetBytes(connectionMessage), remoteIpEndPoint)
        _readBuffer = New Byte(100 * 1024 - 1) {}
        Do
            Try
                connection.ReceiveFrom(_readBuffer, remoteIpEndPoint)
                _responseString = System.Text.Encoding.[Default].GetString(_readBuffer)
                _params("ip") = remoteIpEndPoint.Address.ToString
                ParseResponse()
                AddMachine(oPingObject)
                bSuccess = True
            Catch ex As SocketException
                'Console.WriteLine("Failed to receive response on " & IP.ToString & ":" & nPort & ":" & vbCrLf & ex.Message)
                bTimedout = True
            Finally
                connection.Disconnect(False)
            End Try
        Loop Until bTimedout
        Return bSuccess
    End Function

    Private Sub ParseResponse()
        _params("protocolver") = Response(5).ToString()
        If Response(5) = 0 Then
            Return
        End If
        _params("hostname") = ReadNextParam(6)
        _params("mapname") = ReadNextParam()
        _params("mod") = ReadNextParam()
        _params("modname") = ReadNextParam()

        ' The field that denotes the number of players on the server is not necessarily always on this index (Response.Length - 7), therefor the variable i is not accurate.
        ' The next field in the response now is the AppID field (2 byte long), which is a unique ID for all Steam applications.
        ' I'm not sure you want to include it or not, but for now I will.
        'Dim i As Integer = Response.Length - 7
        _params("appid") = (Response(Offset) Or (CShort(Response(System.Threading.Interlocked.Increment(Offset)) << 8))).ToString() ' Perform binary OR to get the short value from the two bytes.
        _params("players") = Response(System.Threading.Interlocked.Increment(Offset)).ToString()
        _params("maxplayers") = Response(System.Threading.Interlocked.Increment(Offset)).ToString()
        _params("botcount") = Response(System.Threading.Interlocked.Increment(Offset)).ToString()
        If Response(5) <> 15 Then ' Protocol version 15 seems to end here. So lets not go any further if thats the protocol of this reply.
            _params("servertype") = ChrW(Response(System.Threading.Interlocked.Increment(Offset))).ToString()
            _params("serveros") = ChrW(Response(System.Threading.Interlocked.Increment(Offset))).ToString()
            _params("passworded") = Response(System.Threading.Interlocked.Increment(Offset)).ToString()
            _params("secured") = Response(System.Threading.Interlocked.Increment(Offset)).ToString()

            Offset += 1 'Increment the offset to take the last read byte into account.

            _params("gameversion") = ReadNextParam()
            ' In most cases, the reply will end here. But some servers include an Extra Data Flag along
            ' with some extra data, so if we'll read that too, if available.

            If Offset < Response.Length Then
                Dim flag As Byte = Response(Offset)
                Offset += 1
                If (flag And &H80) > 0 Then '  	The server's game port # is included
                    _params("serverport") = CStr(BitConverter.ToInt16(Response, Offset))
                    Offset += 2
                End If


                If (flag And &H40) > 0 Then ' The spectator port # and then the spectator server name are included
                    _params("spectatorport") = CStr(BitConverter.ToInt16(Response, Offset))
                    Offset += 2
                    _params("spectatorname") = ReadNextParam()
                End If

                If (flag And &H20) > 0 Then ' The game tag data string for the server is included [future use]
                    _params("gametagdata") = ReadNextParam()
                End If
            End If
        End If
    End Sub

    Private Sub AddMachine(ByRef oPingObject As PingObject)
        Dim obj As New NetworkObject()
        Dim sIP As String = _params("ip")
        Dim sHostName As String = System.Net.Dns.GetHostEntry(sIP).HostName
        If sHostName = sIP Then
            obj.Name = sHostName
        Else
            If sIP = IPHlp.Localhost Then
                sIP = "localhost"
            End If
            obj.Name = sHostName & " (" & sIP & ")"
        End If
        'For Each key In _params.Keys
        '    Console.WriteLine(key & " = " & _params(key))
        'Next
        obj.IP = sIP
        obj.Params = _params
        oPingObject.cMachines.Add(obj)
    End Sub

    Private Sub TimeoutHandler(ByVal source As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timeout.Elapsed
        Timeout.Stop()
        Timeout.Close()
        bTimedout = True
        Try
            If Not connection Is Nothing Then
                If connection.ProtocolType = ProtocolType.Tcp Then
                    connection.LingerState.Enabled = False
                End If
                connection.Close()
            End If
        Catch ex As Exception
            Console.WriteLine("Error forcably disconnecting: " & ex.Message)
        End Try
    End Sub

    Protected ReadOnly Property Response() As Byte()
        Get
            Return _readBuffer
        End Get
    End Property

    Protected ReadOnly Property ResponseString() As String
        Get
            Return _responseString
        End Get
    End Property

    Private Property Offset() As Integer
        Get
            Return _offset
        End Get
        Set(ByVal value As Integer)
            _offset = value
        End Set
    End Property

    Private Function ReadNextParam(ByVal offset As Integer) As String
        If offset > _readBuffer.Length Then
            Throw New IndexOutOfRangeException()
        End If
        _offset = offset
        Return ReadNextParam()
    End Function

    Private Function ReadNextParam() As String
        Dim temp As String = ""
        While _offset < _readBuffer.Length
            If _readBuffer(_offset) = 0 Then
                _offset += 1
                Exit While
            End If
            temp += ChrW(_readBuffer(_offset))
            _offset += 1
        End While
        Return temp
    End Function

End Class
