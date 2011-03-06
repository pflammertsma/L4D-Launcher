Imports Microsoft.Win32
Imports System.IO
Imports System.Collections.Specialized
Imports System.Resources
Imports System.Reflection
Imports System.Threading

' -game left4dead -console -novid +sv_lan 1 +sv_allow_lobby_connect_only 0 +z_difficulty normal +map l4d_hospital01_apartment

' HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Left 4 Dead
' "C:\Windows\Left 4 Dead\uninstall.exe" "/U:C:\Program Files\Left 4 Dead\Uninstall\uninstall.xml"

Public Class Launcher
    Private sAppDir As String = ""

    Public Const UpdateURL As String = "http://paul.luminos.nl/download/software/"

    Private gL4D As New L4DGame()
    Private gL4D2 As New L4D2Game()

    Private curGame As Game

    Private sSteamPath As String
    Private Const STEAM_APP As String = "Steam App "

    Delegate Sub ListDelegate(ByVal count As Integer)

    Public Shared _listDelegate As ListDelegate

    Private ThreadWinNT As GetWinNTList = AddressOf NetworkProbe.Probe
    Private ThreadWinNTCallback As AsyncCallback = AddressOf ThreadWinNTComplete

    '*** a delegate for executing handler methods
    Delegate Function GetWinNTList() As Collection
    Delegate Sub UpdateMachinesHandler(ByVal Machines As Collection, ByVal bLast As Boolean)
    Delegate Sub UpdateProgressHandler(ByVal Value As Integer, ByVal Max As Integer)

    Private nDefaultPort As Integer = 27015
    Private nDefaultTimeout As Integer = 2000
    Dim nPort As Integer = nDefaultPort
    Dim nTimeout As Integer = nDefaultTimeout

    Private cControls As New Collection()
    Private settingIniFile As String = "prefs.ini"
    Private nSelectedMap As Integer = -1

    Private cMaps As New Collection
    Private cAddons As New Collection

    Public cSurvivalMaps As New ArrayList

    Private RefreshingMaps As Boolean
    Private RefreshingNetwork As Boolean
    Private ListEmpty As Boolean
    Private cMachines As New Collection

    Private dLastCheck As Date

    Private Shared _window As Launcher
    Delegate Function CheckForUpdatesT(ByVal window As Launcher, ByVal bUpdatePrompt As Boolean) As Boolean
    Delegate Sub CheckForUpdatesCompleteT()
    Private Shared UpdateCheck As CheckForUpdatesT = AddressOf CheckForUpdates
    Private Shared UpdateCheckCallback As AsyncCallback = AddressOf CheckForUpdatesComplete

    Private Sub Launcher_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Hide()
        [Delegate].RemoveAll(_listDelegate, _listDelegate)
        End
    End Sub

    Private Sub SetPath(ByRef game As Game, ByVal sPath As String)
        If Mid(sPath, Len(sPath) - 1) <> "\" Then
            sPath = sPath & "\"
        End If
        Dim iPath As New DirectoryInfo(sPath)
        If iPath.Exists() Then
            If game.sGameDir <> "" And LCase(game.sGameDir) <> LCase(sPath) Then
                PathChoose.btn1.Text = game.sGameDir
                PathChoose.btn2.Text = sPath
                If PathChoose.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
                    End
                End If
                sPath = PathChoose.Path
            End If
            game.sGameDir = sPath
            game.bNeedFix = False
        End If
    End Sub

    Private Function GetSteam(ByRef Hive As RegistryHive, ByVal sKey As String, ByRef sRegPath As String) As Boolean
        Dim sErr As String = ""
        Dim sPath As String = GetRegValue(Hive, "SOFTWARE\Valve\Steam\", sKey, sErr)
        If sErr = "" Then
            sRegPath = Replace(sPath, "/", "\")
            Dim sDir As String = sRegPath & "\steamapps\common\"
            SetPath(gL4D, sDir & gL4D.sFolderName)
            SetPath(gL4D2, sDir & gL4D2.sFolderName)
            Return True
        End If
        Return False
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim iBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
            lblVersion.Text = "v" & iBuildInfo.ProductVersion
            lblVersion.BackColor = Color.Transparent
        Catch
            lblVersion.Visible = False
        End Try

        sAppDir = AppDomain.CurrentDomain.BaseDirectory
        If sAppDir.Substring(sAppDir.Length - 1) <> "\" Then
            sAppDir &= "\"
        End If

        Try
            Dim sErr As String = ""
            GetSteam(RegistryHive.LocalMachine, "InstallPath", sSteamPath)
            GetSteam(RegistryHive.CurrentUser, "InstallPath", sSteamPath)
            GetSteam(RegistryHive.LocalMachine, "SteamPath", sSteamPath)
            GetSteam(RegistryHive.CurrentUser, "SteamPath", sSteamPath)
            If sSteamPath <> "" Then
                Dim keys As String()
                keys = GetRegKeys(RegistryHive.LocalMachine, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", sErr)
                If sErr <> "" And False Then
                    MsgBox("Something went wrong while populating the Steam application paths from the registry." & vbCrLf & vbCrLf & sErr, MsgBoxStyle.Critical)
                End If
                ' Steam app ids from:
                ' http://developer.valvesoftware.com/wiki/Steam_Application_IDs
                For Each key As String In keys
                    If key.StartsWith(STEAM_APP) Then
                        Dim sPath As String = GetRegValue(RegistryHive.LocalMachine, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & key, "InstallLocation", sErr)
                        If sErr = "" Then
                            Select Case key.Substring(Len(STEAM_APP))
                                Case "500"
                                    SetPath(gL4D, sPath)
                                Case "550"
                                    SetPath(gL4D2, sPath)
                            End Select
                        End If
                    End If
                Next
            Else
                ' MsgBox("No Steam installation found!" & vbCrLf & vbCrLf & sErr, MsgBoxStyle.Critical)
            End If

            With gL4D
                If .bNeedFix Then
                    Dim pos1 As Integer, pos2 As Integer
                    Dim sRegPath As String = GetRegValue(RegistryHive.LocalMachine, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Left 4 Dead", "UninstallString", sErr)
                    If sErr <> "" And False Then
                        MsgBox("Something went wrong while reading the install path from the registry." & vbCrLf & vbCrLf & sErr, MsgBoxStyle.Critical)
                    End If
                    pos1 = InStr(sRegPath, "/U:")
                    If pos1 > -1 Then
                        pos1 += 3
                        pos2 = InStrRev(sRegPath, "\Uninstall\uninstall.xml")
                        If pos2 > pos1 Then
                            SetPath(gL4D, Mid(sRegPath, pos1, pos2 - pos1 + 1))
                        End If
                    End If
                End If
            End With
        Catch
            MsgBox(Err.Description)
        End Try

        curGame = gL4D
        If gL4D2.sGameDir <> "" Then
            curGame = gL4D2
        End If

        If gL4D.sGameDir <> "" And gL4D2.sGameDir <> "" Then
            GameSelect.Text = Application.ProductName
            If GameSelect.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                End
            End If
            Select Case GameSelect.Selected
                Case gL4D.sId
                    curGame = gL4D
                Case gL4D2.sId
                    curGame = gL4D2
            End Select
        End If
        If curGame.Equals(gL4D) Then
            picBanner.Image = GameSelect.btn1.Image
        Else
            picBanner.Image = GameSelect.btn2.Image
        End If

        GetControls(Me)
        For Each sPlayer As String In curGame.GetPlayers()
            cboPlayer.Items.Add(sPlayer)
        Next

        LoadSettings()
        btnFix.Visible = curGame.bNeedFix
        If cboPlayer.Items.Count > 0 Then
            cboPlayer.SelectedIndex = 0
        Else
            chkPlayer.Enabled = False
        End If
        cboGameType.SelectedIndex = 0
        cboScan.SelectedIndex = 0
        txtPort.Text = nDefaultPort
        txtTimeout.Text = nDefaultTimeout

        Call chkName_CheckedChanged(Nothing, Nothing)
        Call chkPlayer_CheckedChanged(Nothing, Nothing)
        Call chkCustomPeer_CheckedChanged(Nothing, Nothing)
        lblPort.Text = "Default: " & nDefaultPort
        lblTimeout.Text = "Recommended: " & nDefaultTimeout

        Me.Show()

        For Each arg As String In Environment.GetCommandLineArgs()
            Dim parts As String() = arg.Split("=")
            If parts.Length = 2 Then
                Select Case parts(0)
                    Case "gamedir"
                        Dim sDir As String = StripQuotes(parts(1))
                        If CheckGameDir(sDir, False, True) Then
                            curGame.sGameDir = sDir
                        End If
                End Select
            End If
        Next arg

        CheckGameDir(curGame.sGameDir, False, False)

        btnRefresh.Enabled = False
        RefreshMaps()
        RefreshNetwork()

    End Sub

    Sub RefreshMaps()
        RefreshingMaps = True
        tabs_SelectedIndexChanged(Nothing, Nothing)

        lstMaps.Items.Clear()
        grpVPK.Text = "Scanning directories..."
        grpVPK.Visible = True
        pgbVPK.Value = 0
        pgbNetwork.Style = ProgressBarStyle.Marquee
        cboMaps.Enabled = False

        cboMaps.Items.Clear()
        cboMaps.Items.Add("All original maps")
        cboMaps.Items.Add("Original coop maps")
        cboMaps.Items.Add("Original versus maps")
        cboMaps.Items.Add("Original survival maps")

        cMaps.Clear()
        Dim pos As Integer
        Dim map As String = Dir(curGame.sGameDir & curGame.sDataFolder & "\maps\*.bsp")
        Do While map <> ""
            Application.DoEvents()
            Dim oMap As New Map(map)
            cMaps.Add(oMap)
            map = Dir()
        Loop

        cAddons.Clear()
        Dim addon As String = Dir(curGame.sGameDir & curGame.sDataFolder & "\addons\*.vpk")
        Do While addon <> ""
            Application.DoEvents()
            pos = InStrRev(addon, ".vpk")
            If pos > 0 Then
                cAddons.Add(addon)
                'Console.WriteLine(addon)
                Dim parser As New VPKParser()
                parser.Parse(curGame.sGameDir & curGame.sDataFolder & "\addons\", addon, cMaps, pgbVPK)
                addon = Mid(addon, 1, pos - 1)
                cboMaps.Items.Add("Addon: " & addon)
            End If
            addon = Dir()
        Loop

        grpVPK.Visible = False
        If nSelectedMap >= 0 And nSelectedMap < cboMaps.Items.Count Then
            cboMaps.SelectedIndex = nSelectedMap
        Else
            cboMaps.SelectedIndex = 0
        End If
        cboMaps.Enabled = True
        cboMaps_SelectedIndexChanged(Nothing, Nothing)
        RefreshingMaps = False
        tabs_SelectedIndexChanged(Nothing, Nothing)
    End Sub

    Public Function StripQuotes(ByVal input As String) As String
        If Mid(input, 1, 1) = """" Then
            input = Mid(input, 2)
        End If
        If Mid(input, Len(input)) = """" Then
            input = Mid(input, 1, Len(input) - 1)
        End If
        Return input
    End Function

    Public Function GetRegKey(ByVal Hive As RegistryHive)
        Select Case Hive
            Case RegistryHive.ClassesRoot
                Return Registry.ClassesRoot
            Case RegistryHive.CurrentConfig
                Return Registry.CurrentConfig
            Case RegistryHive.CurrentUser
                Return Registry.CurrentUser
            Case RegistryHive.DynData
                Return Registry.DynData
            Case RegistryHive.LocalMachine
                Return Registry.LocalMachine
            Case RegistryHive.PerformanceData
                Return Registry.PerformanceData
            Case RegistryHive.Users
                Return Registry.Users
            Case Else
                Return Nothing
        End Select
    End Function

    Public Function GetRegValue(ByVal Hive As RegistryHive, ByVal Key As String, ByVal ValueName As String, Optional ByRef ErrInfo As String = "") As String
        Dim objParent As RegistryKey = GetRegKey(Hive)
        If objParent Is Nothing Then Return Nothing

        Dim objSubkey As RegistryKey
        Dim sAns As String = ""

        Try
            objSubkey = objParent.OpenSubKey(Key)
            'if can't be found, object is not initialized
            If Not objSubkey Is Nothing Then
                sAns = (objSubkey.GetValue(ValueName))
            End If

        Catch ex As Exception

            ErrInfo = ex.Message
        Finally

            'if no error but value is empty, populate errinfo
            If ErrInfo = "" And sAns = "" Then
                ErrInfo = _
                   "No value found for requested registry key"
            End If
        End Try
        Return sAns
    End Function

    Public Function GetRegKeys(ByVal Hive As RegistryHive, ByVal Key As String, Optional ByRef ErrInfo As String = "") As String()
        Dim objParent As RegistryKey = GetRegKey(Hive)
        If objParent Is Nothing Then Return Nothing

        Dim objSubkey As RegistryKey
        Dim subkeys As String()

        Try
            objSubkey = objParent.OpenSubKey(Key)
            'if can't be found, object is not initialized
            If Not objSubkey Is Nothing Then
                subkeys = objSubkey.GetSubKeyNames()
            End If

        Catch ex As Exception

            ErrInfo = ex.Message
        Finally

            'if no error but value is empty, populate errinfo
            If ErrInfo = "" And subkeys Is Nothing Then
                ErrInfo = _
                   "No value found for requested registry key"
            End If
        End Try
        Return subkeys
    End Function

    Public Function SetRegValue(ByVal Hive As RegistryHive, ByVal Key As String, ByVal ValueName As String, ByVal Value As String, Optional ByRef ErrInfo As String = "") As Boolean
        Dim success As Boolean = False

        Try
            Dim objSubkey As RegistryKey = CreateRegKeys(Hive, Key, "", ErrInfo)
            'if can't be found, object is not initialized
            If Not objSubkey Is Nothing Then
                objSubkey.SetValue(ValueName, Value)
                success = True
                objSubkey.Close()
            ElseIf ErrInfo = "" Then
                ErrInfo = "Failed opening registry key: """ & Key & """"
            End If
        Catch ex As Exception
            ErrInfo = ex.Message
        End Try
        Return success
    End Function

    Public Function CreateRegKeys(ByVal Hive As RegistryHive, ByVal Key As String, ByVal SubKey As String, ByRef ErrInfo As String) As RegistryKey
        If SubKey = "" Then
            Console.WriteLine("Navigating to registry subkey: """ & SubKey & """ in """ & Key & """")
        Else
            Console.WriteLine("Navigating to registry key """ & Key & """")
        End If
        Try
            Dim objParent As RegistryKey = GetRegKey(Hive)
            If objParent Is Nothing Then
                ErrInfo = "Registry hive does not exist."
                Return Nothing
            End If

            Dim objSubkey As RegistryKey = objParent.OpenSubKey(Key, True)
            If objSubkey Is Nothing Then
                Dim nPos As Integer = Key.LastIndexOf("\")
                If nPos > 0 Then
                    objSubkey = CreateRegKeys(Hive, Mid(Key, 1, nPos), Mid(Key, nPos + 2), ErrInfo)
                    If objSubkey Is Nothing Then
                        Return Nothing
                    End If
                Else
                    ErrInfo = "Reached registry root; something went wrong with creating the registry key"
                    Return Nothing
                End If
            ElseIf SubKey <> "" Then
                Console.WriteLine("Creating registry subkey: """ & SubKey & """ in """ & Key & """")
                objSubkey.CreateSubKey(SubKey).Close()
            End If
            objSubkey = objParent.OpenSubKey(Key, True)
            If objSubkey Is Nothing Then
                ErrInfo = "The registry key couldn't be created:" & vbCrLf & Key
            End If
            Return objSubkey
        Catch ex As Exception
            ErrInfo = ex.Message
        End Try
        Return Nothing
    End Function

    Function CheckGameDir(ByRef sDir As String, ByVal showError As Boolean, ByVal fixReg As Boolean) As Boolean
        If sDir.Length > 1 Then
            If sDir.Substring(sDir.Length - 1) <> "\" Then
                sDir &= "\"
            End If
        End If
        sDir = sDir.Replace("\\", "\")
        Dim exeFile As New FileInfo(sDir & curGame.sEXE)
        If exeFile.Exists() Then
            If fixReg Then
                If VistaSecurity.IsAdmin Then
                    Dim str As String = """C:\Windows\Left 4 Dead\uninstall.exe"" ""/U:" & sDir & "\Uninstall\uninstall.xml"""
                    Dim sErr As String = ""
                    If SetRegValue(RegistryHive.LocalMachine, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Left 4 Dead", "UninstallString", str, sErr) Then
                        curGame.bNeedFix = False
                    Else
                        If sErr <> "" Then
                            MsgBox("Something went wrong while storing the registry information. You might want to try again running this application in administrator mode." & vbCrLf & vbCrLf & sErr, MsgBoxStyle.Critical)
                        Else
                            MsgBox("Something went wrong while storing the registry information. You might want to try again running this application in administrator mode.", MsgBoxStyle.Critical)
                        End If
                    End If
                Else
                    VistaSecurity.RestartElevated("gamedir=""" & sDir & """")
                End If
            End If
            btnFix.Visible = curGame.bNeedFix
            Return True
        End If
        btnFix.Visible = curGame.bNeedFix
        If sDir = "" Then
            sDir = sSteamPath & "\Left 4 Dead\"
        End If
        If showError Then
            MsgBox("The specified path does not contain """ & curGame.sEXE & """.", MsgBoxStyle.Exclamation)
        End If
        NotFound.SetPath(sDir)
        NotFound.NoExit = False
        Dim res As DialogResult = NotFound.ShowDialog(Me)
        If res <> Windows.Forms.DialogResult.Cancel Then
            sDir = NotFound.GetPath()
            If res = Windows.Forms.DialogResult.Retry Then
                CheckGameDir(sDir, True, True)
            Else
                CheckGameDir(sDir, True, False)
            End If
        Else
            Me.Close()
        End If
    End Function

    Function IsLocalhost(ByVal sIP As String) As Boolean
        If sIP = "localhost" Or sIP = IPHlp.Localhost Then
            Return True
        End If
        Return False
    End Function

    Sub Launch(ByVal mode As Integer)
        Dim sCommandLine As String
        sCommandLine = " -port 27015 +sv_pure 0"
        If chkConsole.Checked Then
            sCommandLine &= " -console"
        End If
        If Not chkVideo.Checked Then
            sCommandLine &= " -novid"
        End If
        If txtCommandLine.TextLength > 0 Then
            sCommandLine &= " " & txtCommandLine.Text
        End If
        If txtName.Enabled Then
            sCommandLine &= " +name """ & txtName.Text & """"
            If curGame.Equals(gL4D) Then
                Dim revIniFile As String = "rev.ini"
                Dim iniFile As New FileInfo(curGame.sGameDir & revIniFile)
                If iniFile.Exists() Then
                    Dim revIni As New IniFile(curGame.sGameDir & revIniFile)
                    Dim sName As String = revIni.GetString("steamclient", "PlayerName", "")
                    If sName <> txtName.Text Then
                        Dim iniBakFile As New FileInfo(curGame.sGameDir & revIniFile & ".bak")
                        If Not iniBakFile.Exists() Then
                            FileCopy(curGame.sGameDir & revIniFile, curGame.sGameDir & revIniFile & ".bak")
                        End If
                        revIni.WriteString("steamclient", "PlayerName", txtName.Text)
                    End If
                Else
                    If MsgBox("The INI file required to set the player name is missing:" & vbCrLf & vbCrLf & "    " & curGame.sGameDir & revIniFile & vbCrLf & vbCrLf & "Continue?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        If cboPlayer.Enabled Then
            sCommandLine &= " +team_desired ""Survivor " & cboPlayer.Text & """"
        End If
        If mode = 2 Then
            If chkCustomPeer.Checked Then
                If txtPeer.Text = "" Then
                    MsgBox("You didn't specify a custom peer to connect to. Please enter a machine name.", MsgBoxStyle.Exclamation)
                    Exit Sub
                End If
                If IsLocalhost(txtPeer.Text) Then
                    If MsgBox("Are you sure you want to connect to your own machine?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
                sCommandLine &= " +connect " & txtPeer.Text
            ElseIf Not lstNetwork.Enabled Or ListEmpty Then
                MsgBox("The list of network peers is empty. Refresh it and select one, or specify a custom machine name.", MsgBoxStyle.Exclamation)
                Exit Sub
            ElseIf lstNetwork.SelectedItems.Count = 0 Then
                MsgBox("Please make a selection from the list of network peers or specify a custom machine name.", MsgBoxStyle.Exclamation)
                Exit Sub
            Else
                Dim index As Integer = lstNetwork.SelectedItems(0).Index
                If IsLocalhost(cMachines.Item(index + 1)) Then
                    If MsgBox("Are you sure you want to connect to your own machine?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
                sCommandLine &= " +connect " & cMachines.Item(index + 1)
            End If
        Else
            If lstMaps.SelectedItems.Count = 0 Then
                MsgBox("Please make a selection from the list of maps.", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            Dim map As String = lstMaps.SelectedItem.ToString
            If cboGameType.SelectedIndex = 1 Then
                If map.IndexOf("_vs_") = 0 Then
                    If MsgBox("You seem to have selected a non-versus map. Are you sure you would like to use the versus game mode?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
                sCommandLine &= " +sv_gametypes versus +mp_gamemode versus"
            ElseIf cboGameType.SelectedIndex = 2 Then
                If MsgBox("You need to have the Left 4 Dead Survival Pack to play survival mode. Also, be sure that the selected map supports survival mode." & vbCrLf & vbCrLf & "Are you sure you want to use the survival game mode?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If
                sCommandLine &= " +sv_gametypes survival +mp_gamemode survival"
            Else
                If map.IndexOf("_vs_") > 0 Then
                    If MsgBox("You seem to have selected a versus map. Are you sure you would like to use the cooperative game mode?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
                sCommandLine &= " +sv_gametypes coop +mp_gamemode coop"
            End If
            If optEasy.Checked Then
                sCommandLine &= " +z_difficulty easy"
            End If
            If optNormal.Checked Then
                sCommandLine &= " +z_difficulty normal"
            End If
            If optHard.Checked Then
                sCommandLine &= " +z_difficulty hard"
            End If
            If optImpossible.Checked Then
                sCommandLine &= " +z_difficulty impossible"
            End If
            If chkMultiplayer.Checked Then
                sCommandLine &= " +sv_lan 1 +sv_allow_lobby_connect_only 0 +hostport " & nPort
            End If
            If lstMaps.SelectedItems.Count > 0 Then
                sCommandLine &= " +map " & map
            End If
        End If
        Dim sExec As String = curGame.sGameDir & curGame.sEXE & " -game " & curGame.sDataFolder & sCommandLine
        If chkClipboard.Checked Then
            Clipboard.SetText(sExec)
            MsgBox("The execution command has been copied to the clipboard:" & vbCrLf & vbCrLf & sExec, MsgBoxStyle.Information)
        Else
            Shell(sExec, AppWinStyle.NormalFocus)
        End If
    End Sub

    Private Sub btnLaunch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLaunch.Click
        If tabs.SelectedIndex = 0 Then
            Launch(1)
        ElseIf tabs.SelectedIndex = 1 Then
            Launch(2)
        Else
            MsgBox("Please specify whether you want to host or join a game by selecting the appropriate tab.", MsgBoxStyle.Information)
        End If
        SaveSettings()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Visible = False
        SaveSettings()
        Me.Close()
    End Sub

    Sub RefreshNetwork()
        RefreshingNetwork = True
        tabs_SelectedIndexChanged(Nothing, Nothing)
        ListEmpty = True

        lstNetwork.Items.Clear()
        UpdateNetwork("Populating servers...", , ProgressBarStyle.Marquee)

        nTimeout = CInt(txtTimeout.Text)
        nPort = CInt(txtPort.Text)
        Dim bBroadcast As Boolean = (cboScan.SelectedIndex = 0)
        PortConnect.Connect(Me, bBroadcast, nPort, nTimeout)
    End Sub

    Sub UpdateNetwork(ByVal sTitle As String, Optional ByVal sText As String = Nothing, Optional ByVal pStyle As ProgressBarStyle = ProgressBarStyle.Blocks)
        If sTitle Is Nothing Then
            grpNetwork.Visible = False
            Exit Sub
        End If
        grpNetwork.Text = sTitle
        grpNetwork.Visible = True
        pgbNetwork.Value = 0
        If sText Is Nothing Then
            lblNetwork.Visible = False
            pgbNetwork.Visible = True
            btnNetworkCancel.Visible = True
        Else
            pgbNetwork.Visible = False
            btnNetworkCancel.Visible = False
            lblNetwork.Text = sText
            lblNetwork.Visible = True
        End If
        pgbNetwork.Style = pStyle
    End Sub

    Sub ThreadWinNTComplete(ByVal iAsyncResult As IAsyncResult)
        Try
            Dim Machines As Collection
            Machines = ThreadWinNT.EndInvoke(iAsyncResult)
            If Me.InvokeRequired Then
                Dim handler As New UpdateMachinesHandler(AddressOf UpdateMachines)
                Dim args() As Object = {Machines, True}
                Me.BeginInvoke(handler, args)
            Else
                UpdateMachines(Machines, True)
            End If
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    Sub ResetUpdates()
        SyncLock Me
            btnUpdate.Enabled = True
        End SyncLock
    End Sub

    Sub UpdateMachines(ByVal Machines As Collection, ByVal bLast As Boolean)
        SyncLock Me
            If ListEmpty And Machines.Count > 0 Then
                lstNetwork.Items.Clear()
                cMachines.Clear()
                ListEmpty = False
            End If
            Dim obj As NetworkObject
            For Each obj In Machines
                If obj.Params.ContainsKey("mod") Then
                    Debug.Print("Server on " & obj.IP & " has mod " & obj.Params("mod"))
                    If curGame.sServerModName.Length > 0 And obj.Params("mod") <> curGame.sServerModName Then
                        Continue For
                    End If
                End If
                Dim items(4) As String
                items(0) = obj.Name
                If obj.Params.ContainsKey("hostname") Then
                    items(1) = obj.Params("hostname")
                End If
                If obj.Params.ContainsKey("players") Then
                    If obj.Params.ContainsKey("maxplayers") Then
                        items(2) = obj.Params("players") & " of " & obj.Params("maxplayers")
                    Else
                        items(2) = obj.Params("players")
                    End If
                End If
                If obj.Params.ContainsKey("mapname") Then
                    items(3) = obj.Params("mapname")
                End If
                Dim item As New ListViewItem(items)
                item.ToolTipText = obj.IP
                lstNetwork.Items.Add(item)
                cMachines.Add(obj.IP)
            Next
            If bLast Then
                If cMachines.Count = 0 Then
                    If curGame.sName.Length > 0 Then
                        UpdateNetwork("", "No " & curGame.sName & " servers found on the network.")
                    Else
                        UpdateNetwork("", "No servers found on the network.")
                    End If
                Else
                    UpdateNetwork(Nothing)
                End If
                btnRefresh.Enabled = True
                RefreshingNetwork = False
                tabs_SelectedIndexChanged(Nothing, Nothing)
            End If
        End SyncLock
    End Sub

    Sub UpdateProgress(ByVal Value As Integer, ByVal Max As Integer)
        pgbNetwork.Maximum = Max
        pgbNetwork.Value = Value
        pgbNetwork.Style = ProgressBarStyle.Blocks
    End Sub

    Sub GetControls(ByRef cParent As Control)
        For Each cControl As Control In cParent.Controls
            cControls.Add(cControl)
            If cControl.HasChildren Then
                GetControls(cControl)
            End If
        Next
    End Sub

    Sub LoadSettings()
        Dim settingIni As New IniFile(sAppDir & settingIniFile)
        LoadSettingsGame(settingIni, gL4D)
        LoadSettingsGame(settingIni, gL4D2)
        With gL4D
            Dim sID As String = curGame.sId
            ' If [controls].cboPlayer is set, we need to grab the options from this section
            If settingIni.GetInteger("controls", "cboPlayer", -1) <> -1 Then
                ' Backwards compatibility for old versions
                sID = "controls"
            End If
        End With
        SetControls(curGame, cControls)
        Dim sVersion As String = settingIni.GetString("launcher", "version", "1.4.1")
        Dim sLastCheck As String = settingIni.GetString("launcher", "last_update_check", New Date(2010, 1, 1).ToString)
        dLastCheck = Date.Parse(sLastCheck)
        ' Use sVersion to add backwards compatibility
        cSurvivalMaps.Add("l4d_hospital02_subway")
        cSurvivalMaps.Add("l4d_hospital03_sewers")
        cSurvivalMaps.Add("l4d_hospital04_interior")
        cSurvivalMaps.Add("l4d_smalltown02_drainage")
        cSurvivalMaps.Add("l4d_smalltown03_ranchhouse")
        cSurvivalMaps.Add("l4d_smalltown04_mainstreet")
        cSurvivalMaps.Add("l4d_smalltown05_houseboat")
        cSurvivalMaps.Add("l4d_airport02_offices")
        cSurvivalMaps.Add("l4d_airport03_garage")
        cSurvivalMaps.Add("l4d_airport04_terminal")
        cSurvivalMaps.Add("l4d_farm02_traintunnel")
        cSurvivalMaps.Add("l4d_farm03_bridge")
        Dim sMap As String
        Dim count As Integer
        Do
            count += 1
            sMap = settingIni.GetString("survivor_maps", CStr(count), Nothing)
            If Not cSurvivalMaps.Contains(sMap) Then
                cSurvivalMaps.Add(sMap)
            End If
        Loop Until sMap Is Nothing
        If DateDiff(DateInterval.Day, dLastCheck, Date.Today) > 14 Then
            StartUpdateCheck(False)
        End If
    End Sub

    Private Sub LoadSettingsGame(ByVal settingIni As IniFile, ByVal game As Game)
        With game
            Dim sDir As String = settingIni.GetString(.sId, "gamedir", Nothing)
            If Not sDir Is Nothing And Directory.Exists(sDir) Then
                .sGameDir = sDir
            End If
            For Each cControl In cControls
                LoadSettingIni(settingIni, gL4D, curGame.sId, cControl)
            Next
        End With
    End Sub

    Public Sub StartUpdateCheck(ByVal bUpdatePrompt As Boolean)
        btnUpdate.Enabled = False
        UpdateCheck.BeginInvoke(Me, bUpdatePrompt, UpdateCheckCallback, Nothing)
    End Sub

    Public Shared Function CheckForUpdates(ByVal window As Launcher, ByVal bUpdatePrompt As Boolean) As Boolean
        _window = window
        Dim nResult As Integer = Updater.Update("", UpdateURL, bUpdatePrompt)
        If nResult <> Updater.UpdateError Then
            _window.dLastCheck = Date.Now
            If bUpdatePrompt And nResult = Updater.UpdateNotAvailable Then
                MsgBox("No updates are available at the moment.", MsgBoxStyle.Information)
            End If
        End If
    End Function

    Public Shared Sub CheckForUpdatesComplete(ByVal iAsyncResult As IAsyncResult)
Retry:
        Try
            SyncLock _window
                Dim handler1 As New CheckForUpdatesCompleteT(AddressOf _window.ResetUpdates)
                Dim args1() As Object = {}
                _window.BeginInvoke(handler1, args1)
            End SyncLock
        Catch ex As InvalidOperationException
            Threading.Thread.Sleep(500)
            GoTo Retry
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    Sub SaveSettings()
        Try
            File.Delete(sAppDir & settingIniFile)
            Dim settingIni As New IniFile(sAppDir & settingIniFile)
            settingIni.WriteString("launcher", "version", Application.ProductVersion)
            settingIni.WriteString("launcher", "last_update_check", dLastCheck.ToString)
            With gL4D
                settingIni.WriteString(.sId, "gamedir", .sGameDir)
            End With
            With gL4D2
                settingIni.WriteString(.sId, "gamedir", .sGameDir)
            End With
            GetControls(curGame, cControls)
            For Each cControl In cControls
                SaveSettingIni(settingIni, gL4D)
            Next
            For Each cControl In cControls
                SaveSettingIni(settingIni, gL4D2)
            Next
            Dim count As Integer
            For Each sMap In cSurvivalMaps
                count += 1
                settingIni.WriteString("survivor_maps", CStr(count), sMap)
            Next
        Catch
            MsgBox("The settings could not be saved.", MsgBoxStyle.Critical)
        End Try
    End Sub

    Sub SetControls(ByRef game As Game, ByRef cControls As Collection)
        For Each cControl As Control In cControls
            Try
                If cControl.Name = "cboMaps" Then
                    nSelectedMap = Integer.Parse(game.GetControl(cControl.Name))
                Else
                    If TypeOf cControl Is CheckBox Then
                        Dim cCheck As CheckBox = cControl
                        cCheck.Checked = Integer.Parse(game.GetControl(cControl.Name))
                    ElseIf TypeOf cControl Is TextBox Then
                        Dim cText As TextBox = cControl
                        cText.Text = game.GetControl(cControl.Name)
                    ElseIf TypeOf cControl Is ComboBox Then
                        Dim cCombo As ComboBox = cControl
                        Dim index As Integer = Integer.Parse(game.GetControl(cControl.Name))
                        If cCombo.Items.Count > index Then
                            cCombo.SelectedIndex = index
                        End If
                    End If
                End If
            Catch ex As Exception
            End Try
        Next
    End Sub

    Sub GetControls(ByRef game As Game, ByRef cControls As Collection)
        For Each cControl As Control In cControls
            If TypeOf cControl Is CheckBox Then
                Dim cCheck As CheckBox = cControl
                game.SetControl(cControl.Name, IIf(cCheck.Checked, 1, 0))
            ElseIf TypeOf cControl Is TextBox Then
                Dim cText As TextBox = cControl
                game.SetControl(cControl.Name, cText.Text)
            ElseIf TypeOf cControl Is ComboBox Then
                Dim cCombo As ComboBox = cControl
                game.SetControl(cControl.Name, cCombo.SelectedIndex)
            End If
        Next
    End Sub

    Sub LoadSettingIni(ByRef settingIni As IniFile, ByRef game As Game, ByRef sId As String, ByRef cControl As Control)
        Dim value As Object = ""
        If cControl.Name = "cboMaps" Then
            value = settingIni.GetInteger(sId, cControl.Name, cboMaps.SelectedIndex)
        Else
            If TypeOf cControl Is CheckBox Then
                Dim cCheck As CheckBox = cControl
                value = settingIni.GetInteger(sId, cControl.Name, cCheck.Checked)
            ElseIf TypeOf cControl Is TextBox Then
                Dim cText As TextBox = cControl
                value = settingIni.GetString(sId, cControl.Name, cText.Text)
            ElseIf TypeOf cControl Is ComboBox Then
                Dim cCombo As ComboBox = cControl
                value = settingIni.GetInteger(sId, cControl.Name, cCombo.SelectedIndex)
            Else
                Exit Sub
            End If
        End If
        If Not value Is Nothing Then
            game.SetControl(cControl.Name, value)
        End If
    End Sub

    Sub SaveSettingIni(ByRef settingIni As IniFile, ByRef game As Game)
        For Each sKey As String In game.GetControlKeys()
            Dim cControl As Object = game.GetControl(sKey)
            If Not cControl Is Nothing Then
                If TypeOf cControl Is Integer Then
                    settingIni.WriteInteger(game.sId, sKey, cControl)
                Else
                    settingIni.WriteString(game.sId, sKey, cControl)
                End If
            End If
        Next
    End Sub

    Private Sub chkPlayer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPlayer.CheckedChanged
        cboPlayer.Enabled = chkPlayer.Checked
    End Sub

    Private Sub chkName_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkName.CheckedChanged
        txtName.Enabled = chkName.Checked
        lblName.Visible = chkName.Checked
    End Sub

    Private Sub btnFix_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFix.Click
        NotFound.SetPath(curGame.sGameDir)
        NotFound.NoExit = True
        Dim res As DialogResult = NotFound.ShowDialog(Me)
        If res <> Windows.Forms.DialogResult.Cancel Then
            curGame.sGameDir = NotFound.GetPath()
            If res = Windows.Forms.DialogResult.Retry Then
                CheckGameDir(curGame.sGameDir, True, True)
            Else
                CheckGameDir(curGame.sGameDir, True, False)
            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Sub cboMaps_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMaps.SelectedIndexChanged
        nSelectedMap = cboMaps.SelectedIndex
        lstMaps.Items.Clear()
        Dim items As New ArrayList
        For Each mMap As Map In cMaps
            Select Case cboMaps.SelectedIndex
                Case -1
                    Exit For
                Case 0
                    If Not mMap.AddOn Is Nothing Then Continue For
                Case 1
                    If Not mMap.AddOn Is Nothing Then Continue For
                    If Not mMap.IsCoop Then Continue For
                Case 2
                    If Not mMap.AddOn Is Nothing Then Continue For
                    If Not mMap.IsVersus Then Continue For
                Case 3
                    If Not mMap.AddOn Is Nothing Then Continue For
                    If Not mMap.IsSurival Then Continue For
                Case Else
                    Dim index As Integer = cboMaps.SelectedIndex - 3
                    If index <= 0 Or index > cAddons.Count Then
                        Exit For
                    End If
                    Dim sAddon As String = cAddons(index)
                    If mMap.AddOn <> sAddon Then Continue For
            End Select
            items.Add(mMap.Name)
        Next
        items.Sort()
        lstMaps.Items.AddRange(items.ToArray)
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        If tabs.SelectedIndex = 0 Then
            btnRefresh.Enabled = False
            RefreshMaps()
        ElseIf tabs.SelectedIndex = 1 Then
            btnRefresh.Enabled = False
            RefreshNetwork()
        End If
    End Sub

    Private Sub chkCustomPeer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCustomPeer.CheckedChanged
        txtPeer.Enabled = chkCustomPeer.Checked
        lstNetwork.Enabled = Not txtPeer.Enabled
        If lstNetwork.Enabled Then
            grpNetwork.BackColor = SystemColors.Window
            lstNetwork.ForeColor = SystemColors.WindowText
        Else
            grpNetwork.BackColor = SystemColors.Control
            lstNetwork.ForeColor = SystemColors.ControlDark
        End If
    End Sub

    Private Sub tabs_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabs.SelectedIndexChanged
        Select Case tabs.SelectedIndex
            Case 0
                btnLaunch.Enabled = True
                btnRefresh.Enabled = Not RefreshingMaps
            Case 1
                btnLaunch.Enabled = True
                btnRefresh.Enabled = Not RefreshingNetwork
            Case Else
                btnLaunch.Enabled = False
                btnRefresh.Enabled = False
        End Select
    End Sub

    Private Sub grpGeneral_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnNetworkCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNetworkCancel.Click
        PortConnect.Cancel()
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        StartUpdateCheck(True)
    End Sub
End Class
