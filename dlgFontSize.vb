Imports System.Windows.Forms

Public Class dlgFontSize
    Dim hDocument As HtmlDocument
    Public selectedSize As String = "16px"
    Public strTest As String = "Пример для тестирования шрифта"

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgFontSize_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        splitMain.FixedPanel = FixedPanel.Panel2
        wbFont.Navigate(FileIO.FileSystem.CombinePath(Application.StartupPath, "src\editor.html"))
        cmbPoints.Text = "px"
    End Sub

    Private Sub wbFont_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbFont.DocumentCompleted
        hDocument = wbFont.Document
        If IsNothing(hDocument) Then Return
        'hDocument.Body.InnerHtml = "<table width=100% height=100%><tr><td><span id='fontTest'>Пример для тестирования шрифта</span></td></tr></table>"
        Dim hTable As HtmlElement = hDocument.CreateElement("TABLE")
        hTable.Style = "width:100%;height:100%"
        hDocument.Body.AppendChild(hTable)
        Dim hTBody As HtmlElement = hDocument.CreateElement("TBODY")
        hTable.AppendChild(hTBody)
        Dim TR As HtmlElement = hDocument.CreateElement("TR")
        hTBody.AppendChild(TR)
        Dim TD As HtmlElement = hDocument.CreateElement("TD")
        TD.Style = "text-align: center;vertical-align: middle"
        TR.AppendChild(TD)
        Dim hSpan As HtmlElement = hDocument.CreateElement("SPAN")
        hSpan.Id = "fontTest"
        hSpan.InnerHtml = WrapTextToFont(strTest)
        TD.AppendChild(hSpan)

    End Sub

    Public Function WrapTextToFont(ByVal strText As String) As String
        Dim txt As String = "<span style=" + Chr(34) + "font-size: " + cmbSize.Text + cmbPoints.Text
        If cmbFontName.Text.Length > 0 Then
            txt += ";font-family: " + cmbFontName.Text
        End If
        If chkBold.Checked Then
            txt += ";font-weight: bold"
        End If
        If chkItalic.Checked Then
            txt += ";font-style: italic"
        End If
        If chkUnderline.Checked Then
            txt += ";text-decoration:underline"
        End If
        txt += Chr(34) + ">" + strText + "</span>"
        Return txt
    End Function

    Private Sub cmbSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSize.SelectedIndexChanged, cmbSize.TextChanged, cmbPoints.SelectedIndexChanged, _
        cmbFontName.SelectedIndexChanged, cmbFontName.TextChanged, chkBold.CheckedChanged, chkItalic.CheckedChanged, chkUnderline.CheckedChanged
        If IsNothing(hDocument) Then Return
        Dim hSpan As HtmlElement = hDocument.GetElementById("fontTest")
        If IsNothing(hSpan) Then Return
        hSpan.InnerHtml = WrapTextToFont(strTest)
    End Sub
End Class
