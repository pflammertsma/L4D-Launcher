Imports System.Net.NetworkInformation
Imports System.DirectoryServices
Imports System.Net
Imports System.Management
Imports System.Collections.Specialized

Class NetworkProbe

    Private Shared machines As New Collection()

    Shared Function Probe() As Collection
        'ShowNetworkInterfaces()
        ShowNetworkMachines_NT()
        'ShowNetworkMachines_MO()
        Return machines
    End Function

    Public Shared Sub ShowNetworkInterfaces()
        Dim computerProperties As IPGlobalProperties = IPGlobalProperties.GetIPGlobalProperties()
        Dim nics As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
        If nics Is Nothing OrElse nics.Length < 1 Then
            Console.WriteLine("  No network interfaces found.")
            Exit Sub
        End If
        For Each adapter As NetworkInterface In nics
            Console.WriteLine(adapter.Name)
        Next
    End Sub

    Public Shared Sub ShowNetworkMachines_NT2()
        Dim root As New DirectoryEntry("WinNT:")
        Dim parent As DirectoryServices.DirectoryEntries
        parent = root.Children
        Dim d As DirectoryEntries = parent
        For Each complist As DirectoryEntry In parent
            For Each c As DirectoryEntry In complist.Children
                If (c.Name <> "Schema") Then
                    Console.WriteLine(c.Name)
                End If
            Next
        Next
    End Sub

    Public Shared Sub ShowNetworkMachines_NT()
        Dim childEntry As DirectoryEntry
        Dim ParentEntry As New DirectoryEntry
        Try
            ParentEntry.Path = "WinNT://WORKGROUP"
            For Each childEntry In ParentEntry.Children
                Console.WriteLine(childEntry.SchemaClassName)
                Select Case childEntry.SchemaClassName
                    Case "Domain"
                        Console.WriteLine(childEntry.Name)
                        Dim SubChildEntry As DirectoryEntry
                        Dim SubParentEntry As New DirectoryEntry
                        SubParentEntry.Path = "WinNT://" & childEntry.Name
                        Continue For
                        Dim index As Integer = 0
                        For Each SubChildEntry In SubParentEntry.Children
                            'Dim handler As New Launcher.UpdateProgressHandler(AddressOf Launcher.UpdateProgress)
                            'Dim args() As Object = {index, total}
                            'Launcher.BeginInvoke(handler, args)
                            Try
                                Select Case SubChildEntry.SchemaClassName
                                    Case "Computer"
                                        Console.WriteLine("-> " & SubChildEntry.SchemaClassName & ":")
                                        Console.WriteLine("   " & SubChildEntry.Name)
                                        Dim address As String = Nothing
                                        Try
                                            For Each ip In Dns.GetHostEntry(SubChildEntry.Name).AddressList
                                                address = ip.ToString
                                                Console.WriteLine("   " & ip.ToString)
                                                Exit For
                                            Next
                                        Catch Excep As Exception
                                            Console.WriteLine("   (unknown IP)")
                                        End Try
                                        Dim obj As New NetworkObject()
                                        obj.Name = SubChildEntry.Name
                                        obj.IP = address
                                        machines.Add(obj)
                                End Select
                            Catch Excep As Exception
                                Console.WriteLine("Could not identify schema.")
                            End Try
                            index += 1
                        Next
                End Select
            Next
        Catch Excep As Exception
            Console.WriteLine("Could not populate the network.")
            MsgBox("Could not populate network.", MsgBoxStyle.Exclamation)
            End
        Finally
            ParentEntry = Nothing
        End Try
    End Sub

    Private Shared m_objNetworkInfo As New NetworkInfo()

    Private Shared Function ShowNetworkMachines_MO() As Boolean
        Try
            If False Then
                'show computers for this domain
                ShowComputerList(m_objNetworkInfo.LocalDomainName)
            Else
                Dim collection As ManagementObjectCollection
                collection = m_objNetworkInfo.GetDomainsInfoCollection()
                Dim management_object As ManagementObject
                For Each management_object In collection
                    'add computers for this domain
                    ShowComputerList(management_object("Domain"))
                Next management_object
            End If
        Catch err As Exception
            MsgBox("Error loading domain list:" & vbCrLf & err.ToString(), _
                        vbExclamation, _
                        Application.ProductName)
        End Try

    End Function

    Private Shared Function ShowComputerList(ByVal domain_node As String) As Boolean
        Try
            Dim collection As DirectoryEntry
            collection = m_objNetworkInfo.GetComputersInfoCollection(domain_node)

            Dim management_object As DirectoryEntry

            For Each management_object In collection.Children
                If management_object.Name = m_objNetworkInfo.LocalComputerName Then
                    Console.WriteLine("localhost: " & management_object.Name)
                Else
                    Console.WriteLine(management_object.Name & " (" & management_object.Path & ")")
                End If

                Application.DoEvents()

                management_object.Dispose()
            Next management_object

            collection.Dispose()

        Catch err As Exception
            MsgBox("Error loading user list:" & vbCrLf & err.ToString(), _
                        vbExclamation, _
                        Application.ProductName)
        Finally

        End Try
    End Function

End Class

Class NetworkObject

    Public Name As String
    Public IP As String
    Public Params As StringDictionary

End Class
