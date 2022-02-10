Public Class frmMap
    Public Property CanBeClosed As Boolean = False

    Private Sub frmMap_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If CanBeClosed = False Then
            e.Cancel = True
            Me.Hide()
        End If
    End Sub
    Private Sub frmMap_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class