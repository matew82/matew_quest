Imports System.Windows.Forms

Public Class dlgChars
    Private WithEvents hDocument As HtmlDocument
    Public Property SelectedText As String = ""

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click, hDocument.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgChars_Load(sender As Object, e As EventArgs) Handles Me.Load
        splitMain.FixedPanel = FixedPanel.Panel2
        wbChars.Navigate(FileIO.FileSystem.CombinePath(Application.StartupPath, "src\symbols.html"))
    End Sub

    Private Sub wbChars_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbChars.DocumentCompleted
        hDocument = wbChars.Document
    End Sub

    Private Sub hDocumet_Click(sender As Object, e As HtmlElementEventArgs) Handles hDocument.Click
        If wbChars.ReadyState <> WebBrowserReadyState.Complete Then Return
        Dim hEl As HtmlElement = hDocument.GetElementFromPoint(e.ClientMousePosition)
        If IsNothing(hEl) Then Return

        Do
            If IsNothing(hEl) OrElse hEl.TagName = "BODY" Then Return
            If hEl.TagName = "TR" Then
                SelectedText = hEl.Children(2).InnerText
                Return
            End If
            hEl = hEl.Parent
        Loop
    End Sub

End Class
