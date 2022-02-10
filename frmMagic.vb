Public Class frmMagic
    Public Property CanBeClosed As Boolean = False

    Private Sub frmMagic_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If CanBeClosed = False Then
            e.Cancel = True
            Me.Hide()
        End If
    End Sub
End Class