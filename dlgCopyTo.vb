Imports System.Windows.Forms

Public Class dlgCopyTo

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click, treeCopy.DoubleClick
        Dim n As TreeNode = treeCopy.SelectedNode
        If IsNothing(n) OrElse n.Tag = "GROUP" Then Return

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgCopyTo_Load(sender As Object, e As EventArgs) Handles Me.Load
        splitMain.FixedPanel = FixedPanel.Panel2
        treeCopy.ImageList = frmMainEditor.imgLstGroupIcons
    End Sub
End Class
