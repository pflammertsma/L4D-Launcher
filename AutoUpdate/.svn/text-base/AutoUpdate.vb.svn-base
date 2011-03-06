Imports System.IO
Imports System.Net
Imports System.Security.Permissions
Imports System.Security

Module AutoUpdate

    Public Sub Main()
        Dim ExeFile As String ' the program that called the auto update
        Dim RemoteUri As String ' the web location of the files
        Dim Files() As String ' the list of files to be updated
        Dim Key As String ' the key used by the program when called back 
        ' to know that the program was launched by the 
        ' Auto Update program
        Dim CommandLine As String ' the command line passed to the original 
        Try
            ' Get the parameters sent by the application
            Dim param() As String = Split(Microsoft.VisualBasic.Command(), "|")
            ExeFile = param(0)
            RemoteUri = param(1)
            ' the files to be updated should be separeted by "?"
            Files = Split(param(2), "?")
            Key = param(3)
            CommandLine = param(4)
        Catch ex As Exception
            ' if the parameters wasn't right just terminate the program
            ' this will happen if the program wasn't called by the system 
            ' to be updated
            MsgBox("Automatic updates are performed through the host application. Please check for updates from the interface." & vbCrLf & vbCrLf & "If you are having difficulties updating, please check http://paul.luminos.nl for the latest version of the software.", MsgBoxStyle.Information)
            Exit Sub
        End Try
        Threading.Thread.Sleep(2000)
        Dim nSuccess As Integer = Updater.UpdateFiles(Files, RemoteUri, ExeFile, CommandLine, Key)
        ' Return number of files that couldn't be moved
        Environment.ExitCode = Files.Length - nSuccess
    End Sub

End Module

Class Updater

    Public Const UpdateNotAvailable As Integer = 0
    Public Const UpdateAvailable As Integer = 1
    Public Const UpdateError As Integer = 2

    Public Shared Function Update(ByRef CommandLine As String, ByVal RemotePath As String, Optional ByVal bShowError As Boolean = True) As Integer
        Dim Key As String = "%(luminos!)" ' any unique sequence of characters
        ' the file with the update information
        Dim sUpdateFile As String = "update.dat"
        Dim sEXE As String = IO.Path.GetFileName(Application.ExecutablePath)
        Dim sUpdateEXE As String = "AutoUpdate.exe"
        ' the Assembly name 
        Dim AssemblyName As String = _
                System.Reflection.Assembly.GetEntryAssembly.GetName.Name
        ' where are the files for a specific system
        Dim RemoteUri As String = RemotePath & AssemblyName & "/"
        ' clean up the command line getting rid of the key
        CommandLine = Replace(Microsoft.VisualBasic.Command(), Key, "")
        ' Verify if was called by the autoupdate
        If InStr(Microsoft.VisualBasic.Command(), Key) > 0 Then
            Try
                Dim sFile As String = Application.StartupPath & "\" & sUpdateEXE
                ' try to delete the AutoUpdate program, 
                ' since it is not needed anymore
                Dim oFp As FileIOPermission = New FileIOPermission(FileIOPermissionAccess.Write, sFile)
                If (SecurityManager.IsGranted(oFp)) Then
                    System.IO.File.Delete(sFile)
                End If
            Catch ex As Exception
            End Try
            ' return false means that no update is needed
            Return UpdateNotAvailable
        Else
            ' was called by the user
            Try
                Dim myWebClient As New System.Net.WebClient 'the webclient
                ' Download the update info file to the memory, 
                ' read and close the stream
                Dim file As New System.IO.StreamReader( _
                    myWebClient.OpenRead(RemoteUri & sUpdateFile))
                Dim sVersion As String = file.ReadLine()
                Dim sFiles As String = ""
                ' if something was read
                While sVersion > Application.ProductVersion And Not file.EndOfStream
                    Dim sFile As String = file.ReadLine()
                    ' the first parameter is the version. if it's 
                    ' greater then the current version starts the 
                    ' update process
                    If LCase(sFile) = LCase(sUpdateEXE) Then
                        ' Download autoupdate.exe to the application 
                        ' path, so you always have the last version runing
                        myWebClient.DownloadFile(RemotePath & sFile, _
                            Application.StartupPath & "\" & sFile)
                    Else
                        If Len(sFiles) > 0 Then
                            sFiles &= "?"
                        End If
                        sFiles &= sFile
                    End If
                End While
                file.Close()
                If Len(sFiles) > 0 Then
                    If MsgBox(Application.ProductName & " version " & sVersion & " is available!" & vbCrLf & vbCrLf & _
                                "Would you like to restart " & Application.ProductName & " to apply it?", _
                                MsgBoxStyle.Question + MsgBoxStyle.YesNo, Application.ProductName & " v" & Application.ProductVersion) = MsgBoxResult.No Then
                        Return UpdateAvailable
                    End If
                    ' assembly the parameter to be passed to the auto 
                    ' update program
                    ' sFiles is the files that need to be 
                    ' updated separated by "?"
                    Dim arg As String = Application.ExecutablePath & "|" & _
                                RemoteUri & "|" & sFiles & "|" & Key & "|" & _
                                Microsoft.VisualBasic.Command()
                    ' Call the auto update program with all the parameters
                    Dim proc As Process = System.Diagnostics.Process.Start( _
                            Application.StartupPath & "\" & sUpdateEXE, arg)
                    If proc.Id > 0 Then
                        End
                    End If
                End If
            Catch ex As Exception
                ' Something went wrong
                If bShowError Then
                    MsgBox("There was a problem automatically updating. Please visit http://paul.luminos.nl to check for a software update." & _
                           vbCrLf & vbCrLf & "Error: " & ex.Message & vbCrLf & vbCrLf & "Update file: " & RemoteUri & sUpdateFile, MsgBoxStyle.Critical)
                End If
                Return UpdateError
            End Try
            Return UpdateNotAvailable
        End If
    End Function

    Public Shared Function UpdateFiles(ByRef Files() As String, ByVal RemoteUri As String, ByVal ExeFile As String, ByVal CommandLine As String, ByVal Key As String) As Integer
        Dim nSuccess As Integer = 0
        Dim sErr As String = ""
        Try
            ' program if is the case
            Dim myWebClient As New WebClient ' the web client
            Dim cToDo As New Collection
            ' Process each file 
            For i As Integer = 0 To Files.Length - 1
                MsgBox(RemoteUri & Files(i))
                ' download the new version
                Dim sFile As String = Application.StartupPath & "\" & Files(i)
                myWebClient.DownloadFile(RemoteUri & Files(i), sFile & ".updated")
                Dim oFp As FileIOPermission = New FileIOPermission(FileIOPermissionAccess.Write, sFile)
                Dim bFailed As Boolean = True
                If (SecurityManager.IsGranted(oFp)) Then
                    Try
                        ' try to rename the current file before download the new one
                        ' this is a good procedure since the file can be in use
                        File.Delete(sFile & ".old")
                        File.Move(sFile, sFile & ".old")
                        File.Move(sFile & ".updated", sFile)
                        nSuccess += 1
                        bFailed = False
                    Catch ex As Exception
                    End Try
                End If
                If bFailed Then
                    cToDo.Add(Files(i))
                End If
            Next
            Threading.Thread.Sleep(2000)
            ' do some clean up -  delete all .old files (if possible) 
            ' in the current directory
            ' if some file stays it will be cleaned next time
            For Each sFile As String In cToDo
                Try
                    File.Delete(sFile & ".old")
                    File.Move(sFile, sFile & ".old")
                    File.Move(sFile & ".updated", sFile)
                    nSuccess += 1
                Catch ex As Exception
                    MsgBox("Failed updating file:" & vbCrLf & vbCrLf & "   " & sFile, MsgBoxStyle.Critical)
                    File.Delete(Application.StartupPath & "\" & sFile & ".updated")
                End Try
            Next
            ' Call back the system with the original command line 
            ' with the key at the end
            System.Diagnostics.Process.Start(ExeFile, CommandLine & Key)
        Catch ex As Exception
            sErr = "Error: " & ex.Message
        End Try
        If nSuccess < Files.Length Then
            ' something went wrong... 
            MsgBox("There was a problem automatically updating. Please visit http://paul.luminos.nl to check for a software update." & _
                   vbCrLf & vbCrLf & IIf(sErr.Length > 0, sErr, "Only " & nSuccess & " of " & Files.Length & " were updated."), MsgBoxStyle.Critical)
        End If
        Return nSuccess
    End Function

End Class