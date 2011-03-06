Imports System.Windows.Forms

Public Class NotFound

    Public NoExit As Boolean

    Private Sub Dialog1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblInfo.Text = "Left 4 Dead could not be located via the Windows Uninstall list. This means that not only can you not use this Launcher, but you are also unlikely to be able to uninstall the game. This can automatically be fixed by entering the path to the game below."
        If Not VistaSecurity.IsAdmin Then
            VistaSecurity.AddShieldToButton(btnFix)
        End If
        btnExit.Text = IIf(NoExit, "C&ancel", "&Exit")
        txtPath.Focus()
        txtPath.SelectAll()
    End Sub

    Public Sub SetPath(ByVal path As String)
        txtPath.Text = path
    End Sub

    Public Function GetPath() As String
        Return txtPath.Text
    End Function

    Private Sub btnFix_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFix.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Retry
        Me.Close()
    End Sub

    Private Sub btnContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnContinue.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
