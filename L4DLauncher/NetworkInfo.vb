Imports System
Imports System.Management
Imports System.DirectoryServices

Public Class NetworkInfo
    Private m_local_computer_name As String
    Public ReadOnly Property LocalComputerName() As String
        Get
            LocalComputerName = m_local_computer_name
        End Get
    End Property

    Private m_local_domain_name As String
    Public ReadOnly Property LocalDomainName() As String
        Get
            LocalDomainName = m_local_domain_name
        End Get
    End Property

    Public Sub New(Optional ByVal LoadLocalInfo As Boolean = True)
        'as we create object - populate with local info:
        'domain
        'local computer name
        'IP address...etc
        If LoadLocalInfo Then
            GetLocalComputerInfo()
        End If

    End Sub

    Public Function GetLocalComputerInfo() As Boolean

        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection
        Dim query_command As String = "SELECT * FROM Win32_ComputerSystem"

        Dim msc As ManagementScope = New ManagementScope("root\cimv2")

        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Dim management_object As ManagementObject

        For Each management_object In queryCollection
            m_local_domain_name = management_object("Domain")
            m_local_computer_name = management_object("Name")
        Next management_object

        Return True
    End Function

    Public Function GetUsersInfoCollection(ByVal domain As String) As ManagementObjectCollection
        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim msc As ManagementScope = New ManagementScope("root\cimv2")
        Dim query_command As String = "SELECT * FROM Win32_UserAccount  WHERE Domain=" & Chr(34).ToString() & domain & Chr(34).ToString()

        'Win32_UserAccount:
        'see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_useraccount.asp
        'class Win32_UserAccount : Win32_Account
        '{
        '  uint32 AccountType;
        '  string Caption;
        '  string Description;
        '  boolean Disabled;
        '  string Domain;
        '  string FullName;
        '  datetime InstallDate;
        '  boolean LocalAccount;
        '  boolean Lockout;
        '  string Name;
        '  boolean PasswordChangeable;
        '  boolean PasswordExpires;
        '  boolean PasswordRequired;
        '  string SID;
        '  uint8 SIDType;
        '  string Status;
        '};
        '
        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Return queryCollection
    End Function

    Public Function GetServicesInfoCollection(ByVal computer_name As String) As ManagementObjectCollection
        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim msc As ManagementScope = New ManagementScope("\\" & computer_name & "\root\cimv2")
        Dim query_command As String = "SELECT * FROM Win32_Service"
        'Win32_UserAccount:
        'see http://msdn.microsoft.com/library/en-us/wmisdk/wmi/win32_service.asp?frame=true
        '        class Win32_Service : Win32_BaseService
        '{
        '  boolean AcceptPause;
        '  boolean AcceptStop;
        '  string Caption;
        '  uint32 CheckPoint;
        '  string CreationClassName;
        '  string Description;
        '  boolean DesktopInteract;
        '  string DisplayName;
        '  string ErrorControl;
        '  uint32 ExitCode;
        '  datetime InstallDate;
        '  string Name;
        '  string PathName;
        '  uint32 ProcessId;
        '  uint32 ServiceSpecificExitCode;
        '  string ServiceType;
        '  boolean Started;
        '  string StartMode;
        '  string StartName;
        '  string State;
        '  string Status;
        '  string SystemCreationClassName;
        '  string SystemName;
        '  uint32 TagId;
        '  uint32 WaitHint;
        '};

        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Return queryCollection
    End Function

    Public Function GetDomainsInfoCollection() As ManagementObjectCollection
        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim msc As ManagementScope = New ManagementScope("root\cimv2")
        Dim query_command As String = "SELECT * FROM Win32_ComputerSystem"
        'see 
        'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_ntdomain.asp
        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Return queryCollection
    End Function

    Public Function GetComputersInfoCollection(ByVal domain As String) As DirectoryEntry
        'Dim query As ManagementObjectSearcher
        'Dim queryCollection As ManagementObjectCollection

        'Dim msc As ManagementScope = New ManagementScope( "<domain name>\root\directory\LDAP")
        'Dim query_command As String = "SELECT * FROM DS_computer"

        'Dim select_query As SelectQuery = New SelectQuery(query_command)

        'query = New ManagementObjectSearcher(msc, select_query)
        'queryCollection = query.Get()

        'Return queryCollection
        Dim domainEntry As DirectoryEntry = New DirectoryEntry("WinNT://" + domain)
        domainEntry.Children.SchemaFilter.Add("computer")
        Return domainEntry

    End Function

    Public Function GetComputerInfo(ByVal computer_name As String) As ManagementObjectCollection
        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim query_command As String = "SELECT * FROM Win32_ComputerSystem"
        Dim msc As ManagementScope = New ManagementScope("\\" & computer_name & "\root\cimv2")
        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()
        Return queryCollection
    End Function

    Public Function GetComputerOSInfo(ByVal computer_name As String) As ManagementObjectCollection

        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection
        Dim query_command As String = "SELECT * FROM Win32_OperatingSystem"  ' WHERE Name=" & Chr(34).ToString() & computer_name & Chr(34).ToString()"
        Dim msc As ManagementScope = New ManagementScope("\\" & computer_name & "\root\cimv2")
        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()
        Return queryCollection
    End Function



End Class
