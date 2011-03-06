Public Class Map
    Public Filename As String
    Public Name As String
    Public AddOn As String = Nothing
    Public IsCoop As Boolean
    Public IsVersus As Boolean
    Public IsSurival As Boolean

    Public Sub New(ByVal Filename As String, Optional ByVal AddOn As String = Nothing)
        Me.Filename = Filename
        Me.AddOn = AddOn
        Dim pos As Integer = InStrRev(Filename, ".bsp")
        If pos > 0 Then
            Me.Name = Mid(Filename, 1, pos - 1)
        Else
            Me.Name = Filename
        End If
        Me.IsCoop = (Me.Name.IndexOf("_vs_") < 0)
        Me.IsVersus = (Me.Name.IndexOf("_vs_") > 0)
        Me.IsSurival = (Me.Name.IndexOf("_sv_") > 0) Or Launcher.cSurvivalMaps.Contains(Me.Name)
    End Sub

End Class
