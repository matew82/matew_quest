Imports System.Windows.Forms

Public Class dlgSetChildrenValue

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub SplitContainer1_Panel2_Paint(sender As Object, e As PaintEventArgs) Handles splitMain.Panel2.Paint

    End Sub

    Private Sub dlgSetChildrenValue_Load(sender As Object, e As EventArgs) Handles Me.Load
        treeChildren.ImageList = frmMainEditor.imgLstGroupIcons
        splitMain.FixedPanel = FixedPanel.Panel2
        splitInner.FixedPanel = FixedPanel.Panel1
    End Sub

    Private Sub treeChildren_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles treeChildren.AfterCheck
        Dim n As TreeNode = e.Node
        If n.Nodes.Count = 0 Then Return
        For i As Integer = 0 To n.Nodes.Count - 1
            n.Nodes(i).Checked = n.Checked
        Next
    End Sub

    Private Sub treeChildren_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles treeChildren.AfterSelect

    End Sub

    Private Sub btnCheckAll_Click(sender As Object, e As EventArgs) Handles btnCheckAll.Click
        For i As Integer = 0 To treeChildren.Nodes.Count - 1
            treeChildren.Nodes(i).Checked = True
        Next
    End Sub

    Private Sub btnUnselectAll_Click(sender As Object, e As EventArgs) Handles btnUnselectAll.Click
        For i As Integer = 0 To treeChildren.Nodes.Count - 1
            treeChildren.Nodes(i).Checked = False
        Next
    End Sub
End Class
