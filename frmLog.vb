Public Class frmLog

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        lstLog.Items.Clear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If lstLog.SelectedIndex = -1 Then Return
        Dim i As Integer = lstLog.SelectedIndex
        lstLog.Items(i) = lstLog.Items(i).ToString + "----------------"
    End Sub
End Class