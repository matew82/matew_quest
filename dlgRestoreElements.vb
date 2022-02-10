Imports System.Windows.Forms

Public Class dlgRestoreElements
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click, lstElements.DoubleClick
        Dim selId As Integer = lstElements.SelectedIndex
        If selId = -1 Then
            MessageBox.Show("Не выбран элемент для восстановления!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        removedObjects.RetreiveItem(selId)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgRestoreElements_Load(sender As Object, e As EventArgs) Handles Me.Load
        SplitContainer1.FixedPanel = FixedPanel.Panel2
    End Sub


End Class
