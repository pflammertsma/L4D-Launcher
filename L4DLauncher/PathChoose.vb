Imports System.Windows.Forms

Public Class PathChoose

    Public Path As String

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1.Click, btn2.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Path = DirectCast(sender, Button).Text
        Me.Close()
    End Sub

End Class
