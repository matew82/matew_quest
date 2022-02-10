Imports System.Windows.Forms

Public Class dlgListHtml
    Private hDocument As HtmlDocument
    Private selectedType As String = "disc"

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgListHtml_Load(sender As Object, e As EventArgs) Handles Me.Load
        splitMain.FixedPanel = FixedPanel.Panel2
        wbList.Navigate(FileIO.FileSystem.CombinePath(Application.StartupPath, "src\editor.html"))
        UpdateListOfPictures()
    End Sub

    Private Sub wbList_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbList.DocumentCompleted
        hDocument = wbList.Document
        If IsNothing(hDocument) Then Return

        '1. Создае еаблицу с вариантами списоков
        Dim hTable As HtmlElement = hDocument.CreateElement("TABLE")
        hTable.Id = "tableList"
        hDocument.Body.AppendChild(hTable)
        Dim hTBody As HtmlElement = hDocument.CreateElement("TBODY")
        hTable.AppendChild(hTBody)

        Dim lTypes() As String 'варианты css-свойства list-style-type
        Dim lOrdered() As Boolean 'массив одинакового размера с lTypes, в котором хранится инфо о каждом соответствующем типе list-style-type: для сориьованного он списка или для несортированного
        For i As Integer = 0 To 1 'таблица в 2 ряда
            If i = 0 Then
                'типы списоков в первом ряду таблицы
                lTypes = {"none", "disc", "circle", "square", "lower-alpha", "upper-alpha"}
                lOrdered = {False, False, False, False, True, True}
            Else
                'типы списоков во втором ряду таблицы
                lTypes = {"decimal", "decimal-leading-zero", "lower-roman", "upper-roman", "lower-greek", "list-style-image:url(" + questEnvironment.QuestPath.Replace("\", "/") + "/" + cmbImage.Text.Replace("\", "/") + ")"}
                lOrdered = {True, True, True, True, True, False}
            End If
            Dim TR As HtmlElement = hDocument.CreateElement("TR")
            hTBody.AppendChild(TR)
            For j As Integer = 0 To 5'в каждом ряде 6 колонок
                Dim TD As HtmlElement = hDocument.CreateElement("TD")
                Dim strTag As String = IIf(lOrdered(j), "<ol", "<ul")

                Dim strHTML As String 'создаем содержимое каждой клетки таблицы - внешний вид списка
                If i = 1 AndAlso j = 5 Then
                    strHTML = strTag & " style=" & Chr(34) & lTypes(j) & Chr(34) & ">"
                    TD.Id = "listSampleImage"
                    TD.SetAttribute("rVal", "image")
                Else
                    strHTML = strTag & " style=" & Chr(34) & "list-style-type:" & lTypes(j) & Chr(34) & ">"
                    TD.SetAttribute("rVal", lTypes(j))
                End If
                strHTML &= "<li>__________</li><li>__________</li><li>__________</li>"
                If lOrdered(j) Then
                    strHTML &= "</ol>"
                Else
                    strHTML &= "</ul>"
                End If
                TD.InnerHtml = strHTML
                If lTypes(j) = "disc" Then TD.SetAttribute("ClassName", "selected") 'тип disc выбираем по умолчанию
                TR.AppendChild(TD)
                AddHandler TD.Click, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                         For z As Integer = 0 To hTBody.Children.Count - 1
                                             'убираем класс selected со всех клеток таблицы (отменяем выделение)
                                             Dim hTr As HtmlElement = hTBody.Children(z)
                                             For q As Integer = 0 To hTr.Children.Count - 1
                                                 Dim hTd As HtmlElement = hTr.Children(q)
                                                 hTd.SetAttribute("ClassName", "")
                                             Next q
                                         Next z
                                         'выделяем выбранный список
                                         selectedType = sender2.GetAttribute("rVal")
                                         sender2.SetAttribute("ClassName", "selected")
                                     End Sub
            Next j
        Next i

        'отделяем верхнюю таблицу от остальной части документа
        Dim hHR As HtmlElement = hDocument.CreateElement("HR")
        hDocument.Body.AppendChild(hHR)

        '2. Создаем текстбокс для ввода содержимого списков (ввода первого пункта списка, остальные будут автоматом добавляться)
        Dim hDiv As HtmlElement = hDocument.CreateElement("DIV") 'контейнер для текстбоксов
        hDiv.Id = "listInputContainer"
        hDocument.Body.AppendChild(hDiv)
        'сам текстбокс
        Dim hText As HtmlElement = hDocument.CreateElement("INPUT")
        hText.SetAttribute("Type", "text")
        hDiv.AppendChild(hText)
        hText.AttachEventHandler("oninput", Sub() del_text_changed(hText, EventArgs.Empty)) 'событие, срабатывающее на изменение текста
    End Sub

    ''' <summary>
    ''' Делегат события изменения текста текстбоксов
    ''' </summary>
    ''' <param name="sender">ссылка на текстбокс</param>
    ''' <param name="e">пусто</param>
    Private Sub del_text_changed(sender As HtmlElement, e As EventArgs)
        Dim txt As String = sender.GetAttribute("value")
        Dim par As HtmlElement = sender.Parent 'контейнер текстбоксов
        Dim lastChild As HtmlElement = par.Children(par.Children.Count - 1) 'последний текстбокс в контейнере
        If String.IsNullOrEmpty(txt) Then
            If par.Children.Count > 1 Then
                'Текущий текстбокс пуст, при этом текстбоксов как минимум 2
                Dim preLastChild As HtmlElement = par.Children(par.Children.Count - 2)
                If Object.Equals(preLastChild, sender) AndAlso String.IsNullOrEmpty(lastChild.GetAttribute("value")) Then
                    'редактируется предпоследний текстбокс, при этом он сейчас стал пустым. Последний текстбокс таже пуст. Удаляем последний
                    Dim hDomDocument As mshtml.HTMLDocument = hDocument.DomDocument
                    Dim hDiv As mshtml.IHTMLDOMNode = hDomDocument.getElementById("listInputContainer")
                    If IsNothing(hDiv) = False Then hDiv.lastChild.removeNode(True)
                End If
            End If
        Else
            If Object.Equals(sender, lastChild) Then
                'редактируется последнй текстбокс, при этом он не пустой. Добавляем после него еще один текстбокс
                Dim hText As HtmlElement = hDocument.CreateElement("INPUT")
                hText.SetAttribute("Type", "text")
                par.AppendChild(hText)
                hText.AttachEventHandler("oninput", Sub() del_text_changed(hText, EventArgs.Empty))
            End If
        End If
    End Sub

    ''' <summary>
    ''' Обновляет список изображений для списка с картинкой вместо значка в соответствующей комбобоксе
    ''' </summary>
    Private Sub UpdateListOfPictures()
        allowedChanging = False 'запрет на изменение картинки в сэмпле списка с картинкой
        Dim txt As String = cmbImage.Text 'сохраняем текущий текст в комбо
        cmbImage.Items.Clear()
        Dim lst() As String = cListManager.GetFilesList(cListManagerClass.fileTypesEnum.PICTURES, False)
        cmbImage.Items.AddRange(lst) 'заполняем список путями к картинкам относительно папки квеста
        allowedChanging = True
        If txt.Length = 0 Then
            cmbImage.Text = "img/Default/listItem.png" 'если комбо был пуст, то ставим картинку для списков по умолчанию
        Else
            cmbImage.Text = txt 'восстанавливаем тект в комбо до редактирования
        End If
    End Sub

    Dim allowedChanging As Boolean = True 'разрешение / запрет на изменение картинки в сэмпле списка с картинкой
    Private Sub cmbImage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbImage.SelectedIndexChanged, cmbImage.TextChanged
        If allowedChanging = False OrElse IsNothing(hDocument) OrElse wbList.ReadyState <> WebBrowserReadyState.Complete Then Return
        'изменяем изображение сэмпла с картинками перед пунктами списков
        Dim hEL As HtmlElement = hDocument.GetElementById("listSampleImage")
        If IsNothing(hEL) Then Return
        Dim strHTML As String = "<ul style=" & Chr(34) & "list-style-image:url(" + questEnvironment.QuestPath.Replace("\", "/") + "/" + cmbImage.Text.Replace("\", "/") + ")" & Chr(34) & ">"
        strHTML &= "<li>__________</li><li>__________</li><li>__________</li></ul>"
        hEL.InnerHtml = strHTML
    End Sub

    ''' <summary>
    ''' Возвращает строку со списокм для вставки в код
    ''' </summary>
    Public Function GetListString() As String
        Dim res As New System.Text.StringBuilder
        Dim closeTag As String = "" 'для хранения закрывающего тэга списка - </ul> или </ol>
        Select Case selectedType
            'в зависимости от типа выбранного списка selectedType начинаем создание сортированного OL или несортированного UL списка
            Case "none", "disc", "circle", "square"
                res.AppendLine("<UL style=" & Chr(34) & selectedType & Chr(34) & ">")
                closeTag = "</UL>"
            Case "lower-alpha", "upper-alpha", "decimal", "decimal-leading-zero", "lower-roman", "upper-roman", "lower-greek"
                Dim start As Integer = 1
                Integer.TryParse(txtStart.Text, start)
                If start = 1 Then
                    res.AppendLine("<OL style=" & Chr(34) & "list-style-type: " & selectedType & Chr(34) & ">")
                Else
                    res.AppendLine("<OL style=" & Chr(34) & "list-style-type: " & selectedType & Chr(34) & " start=" & start.ToString & ">")
                End If
                closeTag = "</OL>"
            Case "image"
                res.AppendLine("<UL style=" & Chr(34) & "list-style-image:url(" + cmbImage.Text.Replace("\", "/") + ")" & Chr(34) & ">")
                closeTag = "</UL>"
        End Select

        Dim hDiv As HtmlElement = hDocument.GetElementById("listInputContainer") 'получаем контейнер текстбоксов
        If IsNothing(hDiv) Then Return res.ToString & closeTag
        'получаем количество пунктов списка
        Dim UB As Integer = hDiv.Children.Count - 1 'количество текстбоксов
        Dim txt As String = hDiv.Children(hDiv.Children.Count - 1).GetAttribute("Value")
        If String.IsNullOrEmpty(txt) AndAlso UB > 0 Then UB -= 1 'если послдений текстбокс пуст - то его не учитываем (по идее он всегда пуст)

        'вставляем в возвращаемый текст пункты списка
        For i As Integer = 0 To UB
            txt = hDiv.Children(i).GetAttribute("value")
            res.AppendLine("  <LI>" & txt & "</LI>")
        Next i
        res.AppendLine(closeTag) 'и в конце - закрывающий тэг

        Return res.ToString
    End Function

    Private Sub txtStart_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtStart.KeyPress
        'разрешен ввод только цифр, а также Del и Backspace
        If IsNumeric(e.KeyChar) = False AndAlso Asc(e.KeyChar) <> Keys.Delete AndAlso Asc(e.KeyChar) <> Keys.Back Then e.KeyChar = ""
    End Sub

End Class
