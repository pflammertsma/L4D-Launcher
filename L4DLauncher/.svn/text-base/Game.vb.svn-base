Public Class Game
    Public sId As String = ""
    Public sName As String = ""
    Public sFolderName As String = ""
    Public sEXE As String = ""
    Public sDataFolder As String = ""
    Public sGameDir As String = ""
    Public sServerModName As String = ""
    Public bNeedFix As Boolean = True

    Private cControls As New Hashtable

    Dim players(0) As String

    Public Overridable Function GetPlayers() As String()
        Return players
    End Function

    Public Sub SetControl(ByVal key As String, ByVal value As Object)
        If cControls.ContainsKey(key) Then
            cControls.Item(key) = value
        Else
            cControls.Add(key, value)
        End If
    End Sub

    Public Function GetControl(ByVal key As String) As String
        Return cControls.Item(key)
    End Function

    Public Function GetControlKeys() As ICollection
        Return cControls.Keys
    End Function

End Class

Public Class L4DGame
    Inherits Game

    Dim players() As String = {"Bill", "Francis", "Louis", "Zoey"}

    Public Sub New()
        sId = "l4d"
        sName = "Left 4 Dead"
        sFolderName = "left 4 dead"
        sEXE = "left4dead.exe"
        sDataFolder = "left4dead"
        sServerModName = "left4dead"
    End Sub

    Public Overrides Function GetPlayers() As String()
        Return players
    End Function

End Class

Public Class L4D2Game
    Inherits Game

    Dim players() As String = {"Nick", "Ellis", "Rochelle", "Coach"}

    Public Sub New()
        sId = "l4d2"
        sName = "Left 4 Dead 2"
        sFolderName = "left 4 dead 2"
        sEXE = "left4dead2.exe"
        sDataFolder = "left4dead2"
        sServerModName = "left4dead2"
    End Sub

    Public Overrides Function GetPlayers() As String()
        Return players
    End Function

End Class
