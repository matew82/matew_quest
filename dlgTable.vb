Imports System.Windows.Forms

Public Class dlgTable
    Private hDocument As HtmlDocument
    Dim prevRows As Integer, prevCols As Integer

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgTable_Load(sender As Object, e As EventArgs) Handles Me.Load
        splitMain.FixedPanel = FixedPanel.Panel2
        wbTable.Navigate(FileIO.FileSystem.CombinePath(Application.StartupPath, "src\editor.html"))
        prevRows = numRows.Value
        prevCols = numColumns.Value
    End Sub

    Private Sub numRows_ValueChanged(sender As Object, e As EventArgs) Handles numRows.ValueChanged
        If IsNothing(hDocument) Then Return

        Dim diff As Integer = numRows.Value - prevRows 'разница между предыдущим количеством рядов и текущим
        prevRows = numRows.Value 'сохраняем предыдущее значение
        If diff < 0 Then
            'разница отрицательная - удаляем ряды
            Dim hDomDocument As mshtml.HTMLDocument = hDocument.DomDocument
            Dim hTBody As mshtml.IHTMLDOMNode = hDomDocument.getElementById("tbodyValues")
            If IsNothing(hTBody) Then Return
            For i As Integer = 1 To diff * -1
                'удаляем количество рядов = diff
                'Dim tr As mshtml.IHTMLDOMNode = hTBody.lastChild 'children(hTBody.children.length - 1)
                'tr.removeNode(True)
                hTBody.lastChild.removeNode(True)
            Next
        Else
            'разница положительная - добавляем ряды
            Dim hTBody As HtmlElement = hDocument.GetElementById("tbodyValues")
            If IsNothing(hTBody) Then Return
            For i As Integer = 1 To diff
                'добавляем количество рядов = diff
                Dim TR As HtmlElement = hDocument.CreateElement("TR")
                hTBody.AppendChild(TR)
                For c As Integer = 1 To numColumns.Value
                    Dim TD As HtmlElement = hDocument.CreateElement("TD")
                    TR.AppendChild(TD)
                    Dim hText As HtmlElement = hDocument.CreateElement("INPUT")
                    hText.SetAttribute("type", "text")
                    TD.AppendChild(hText)
                    AddHandler hText.KeyUp, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                txtResult.Text = CreateTableString()
                                            End Sub
                    AddHandler hText.GotFocus, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                   hText.SetAttribute("ClassName", "focused")
                                                   TD.SetAttribute("ClassName", "focused")
                                               End Sub
                    AddHandler hText.LostFocus, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                    hText.SetAttribute("ClassName", "")
                                                    TD.SetAttribute("ClassName", "")
                                                End Sub
                Next c
            Next i
        End If
        txtResult.Text = CreateTableString()
    End Sub

    Private Sub numColumns_ValueChanged(sender As Object, e As EventArgs) Handles numColumns.ValueChanged
        If IsNothing(hDocument) Then Return

        Dim diff As Integer = numColumns.Value - prevCols  'разница между предыдущим количеством колонок и текущим
        prevCols = numColumns.Value 'сохраняем предыдущее значение

        If diff < 0 Then
            'разница отрицательная - удаляем колонки
            Dim hDomDocument As mshtml.HTMLDocument = hDocument.DomDocument
            Dim hTBody As mshtml.IHTMLDOMNode = hDomDocument.getElementById("tbodyValues")
            If IsNothing(hTBody) Then Return
            For r As Integer = 0 To hTBody.childNodes.length - 1
                'проходим все ряды, в каждом удаляем количество колонок = diff
                For i As Integer = 1 To diff * -1
                    hTBody.childNodes(r).lastChild.removeNode(True)
                Next
            Next r

        Else
            'разница положительная - добавляем колонки
            Dim hTBody As HtmlElement = hDocument.GetElementById("tbodyValues")
            If IsNothing(hTBody) Then Return
            For r As Integer = 0 To hTBody.Children.Count - 1
                'проходим все ряды, в каждом добавляем количество колонок = diff
                Dim TR As HtmlElement = hTBody.Children(r)
                For i As Integer = 1 To diff
                    Dim TD As HtmlElement = hDocument.CreateElement("TD")
                    TR.AppendChild(TD)
                    Dim hText As HtmlElement = hDocument.CreateElement("INPUT")
                    hText.SetAttribute("type", "text")
                    TD.AppendChild(hText)
                    AddHandler hText.KeyUp, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                txtResult.Text = CreateTableString()
                                            End Sub
                    AddHandler hText.GotFocus, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                   hText.SetAttribute("ClassName", "focused")
                                                   TD.SetAttribute("ClassName", "focused")
                                               End Sub
                    AddHandler hText.LostFocus, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                    hText.SetAttribute("ClassName", "")
                                                    TD.SetAttribute("ClassName", "")
                                                End Sub
                Next i
            Next r
        End If
        txtResult.Text = CreateTableString()
    End Sub

    Private Sub wbTable_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbTable.DocumentCompleted
        hDocument = wbTable.Document

        Dim rows As Integer = numRows.Value
        Dim cols As Integer = numColumns.Value

        UpdateTable()
        txtResult.Text = CreateTableString()
    End Sub

    Private Function CreateTableString() As String
        If IsNothing(hDocument) Then Return ""
        Dim hTBody As HtmlElement = hDocument.GetElementById("tbodyValues")
        If IsNothing(hTBody) Then
            UpdateTable()
            hTBody = hDocument.GetElementById("tbodyValues")
            If IsNothing(hTBody) Then Return ""
        End If

        Dim res As New System.Text.StringBuilder
        res.Append("<TABLE CELLPADDING=0 CELLSPACING=0")
        If txtWidth.TextLength > 0 Then res.Append(" WIDTH = " + Chr(34) + txtWidth.Text + Chr(34))
        res.AppendLine(">")
        If txtCaption.TextLength > 0 Then
            res.AppendLine("<CAPTION>" + txtCaption.Text + "</CAPTION>")
        End If

        For r As Integer = 0 To hTBody.Children.Count - 1
            res.AppendLine("  <TR>")
            Dim tr As HtmlElement = hTBody.Children(r)
            For c As Integer = 0 To tr.Children.Count - 1
                res.Append("    <TD>")
                Dim hText As HtmlElement = tr.Children(c).Children(0)
                Dim strText As String = hText.GetAttribute("Value")
                If String.IsNullOrEmpty(strText) = False Then res.Append(strText)
                res.AppendLine("</TD>")
            Next c
            res.AppendLine("  </TR>")
        Next r

        res.Append("</TABLE>")
        Return res.ToString
    End Function

    Private Sub UpdateTable()
        If IsNothing(hDocument) Then Return
        Dim rows As Integer = numRows.Value
        Dim cols As Integer = numColumns.Value

        Dim hTable As HtmlElement = hDocument.GetElementById("tableValues")
        If IsNothing(hTable) Then
            Dim hCenter As HtmlElement = hDocument.CreateElement("CENTER")
            hDocument.Body.AppendChild(hCenter)
            hTable = hDocument.CreateElement("TABLE")
            hTable.Id = "tableValues"
            If txtWidth.TextLength > 0 Then hTable.SetAttribute("Width", txtWidth.Text)
            hTable.SetAttribute("cellpadding", "0")
            hTable.SetAttribute("cellspacing", "0")
            hCenter.AppendChild(hTable)

            Dim hCaption As HtmlElement = hDocument.CreateElement("CAPTION")
            hCaption.Id = "caption"
            hCaption.InnerHtml = txtCaption.Text
            hTable.AppendChild(hCaption)

            Dim hTBody As HtmlElement = hDocument.CreateElement("TBODY")
            hTBody.Id = "tbodyValues"
            hTable.AppendChild(hTBody)

            For r As Integer = 1 To rows
                Dim TR As HtmlElement = hDocument.CreateElement("TR")
                hTBody.AppendChild(TR)
                For c As Integer = 1 To cols
                    Dim TD As HtmlElement = hDocument.CreateElement("TD")
                    TR.AppendChild(TD)
                    Dim hText As HtmlElement = hDocument.CreateElement("INPUT")
                    hText.SetAttribute("type", "text")
                    TD.AppendChild(hText)
                    AddHandler hText.KeyUp, Sub(sender As Object, e As HtmlElementEventArgs)
                                                txtResult.Text = CreateTableString()
                                            End Sub
                    AddHandler hText.GotFocus, Sub(sender As Object, e As HtmlElementEventArgs)
                                                   hText.SetAttribute("ClassName", "focused")
                                                   TD.SetAttribute("ClassName", "focused")
                                               End Sub
                    AddHandler hText.LostFocus, Sub(sender As Object, e As HtmlElementEventArgs)
                                                    hText.SetAttribute("ClassName", "")
                                                    TD.SetAttribute("ClassName", "")
                                                End Sub
                Next c
            Next r
        End If
    End Sub

    Private Sub txtCaption_TextChanged(sender As Object, e As EventArgs) Handles txtCaption.TextChanged
        If IsNothing(hDocument) Then Return
        Dim hCaption As HtmlElement = hDocument.GetElementById("caption")
        If IsNothing(hCaption) Then
            UpdateTable()
            txtResult.Text = CreateTableString()
            Return
        End If
        hCaption.InnerHtml = txtCaption.Text
        txtResult.Text = CreateTableString()
    End Sub

    Private Sub txtWidth_TextChanged(sender As Object, e As EventArgs) Handles txtWidth.TextChanged
        If IsNothing(hDocument) Then Return
        Dim hTable As HtmlElement = hDocument.GetElementById("tableValues")
        If IsNothing(hTable) Then
            UpdateTable()
            txtResult.Text = CreateTableString()
            Return
        End If
        hTable.SetAttribute("Width", txtWidth.Text)
        txtResult.Text = CreateTableString()
    End Sub
End Class
