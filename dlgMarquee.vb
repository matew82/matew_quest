Imports System.Windows.Forms

Public Class dlgMarquee
    Private WithEvents hDocument As HtmlDocument
    Public sampleText As String ' = "Пример бегущей строки"
    Public selectedMarquee As String

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK        
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgMarquee_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        splitMain.FixedPanel = FixedPanel.Panel2
        wbMarquee.Navigate(FileIO.FileSystem.CombinePath(Application.StartupPath, "src\marquee.html"))
    End Sub

    ''' <summary>
    ''' Обновляет строку, выведенную в качестве примера
    ''' </summary>
    Public Sub UpdateMarqueSample()
        If IsNothing(hDocument) Then Return
        Dim paramType As String = "", paramDirection As String = "", paramLoop As String = "", paramDelay As String = ""
        ReadMarqueeSettings(paramType, paramDirection, paramLoop, paramDelay)

        Dim hCont As HtmlElement = hDocument.GetElementById("marqueeContainer")
        If IsNothing(hCont) Then Return
        selectedMarquee = String.Format("<MARQUEE DIRECTION={0} BEHAVIOR={1} LOOP={2} SCROLLDELAY={3}>", paramDirection, paramType, paramLoop, paramDelay)
        hCont.InnerHtml = selectedMarquee & sampleText & "</MARQUEE>"
    End Sub

    ''' <summary>
    ''' Получает настройки Marquee, выбранные Писателем
    ''' </summary>
    ''' <param name="paramType">Ссылка для получения типа</param>
    ''' <param name="paramDirection">Ссылка для получения направления</param>
    ''' <param name="paramLoop">Ссылка для получения циклов</param>
    ''' <param name="paramDelay">Ссылка для получения задержки</param>
    Private Sub ReadMarqueeSettings(ByRef paramType As String, ByRef paramDirection As String, ByRef paramLoop As String, ByRef paramDelay As String)
        'получаем тип
        paramType = "SCROLL"
        Dim hTable As HtmlElement = hDocument.GetElementById("paramType")
        If IsNothing(hTable) = False Then
            For i As Integer = 0 To hTable.Children(0).Children.Count - 1
                Dim hTD As HtmlElement = hTable.Children(0).Children(i).Children(0)
                If hTD.GetAttribute("ClassName") = "selected" Then
                    paramType = hTD.GetAttribute("rVal")
                    Exit For
                End If
            Next
        End If

        'получаем направление
        paramDirection = "LEFT"
        hTable = hDocument.GetElementById("paramDirection")
        If IsNothing(hTable) = False Then
            For i As Integer = 0 To hTable.Children(0).Children.Count - 1
                Dim hTD As HtmlElement = hTable.Children(0).Children(i).Children(0)
                If hTD.GetAttribute("ClassName") = "selected" Then
                    paramDirection = hTD.GetAttribute("rVal")
                    Exit For
                End If
            Next
        End If

        'получаем кол-во циклов
        paramLoop = "0"
        hTable = hDocument.GetElementById("paramLoop")
        If IsNothing(hTable) = False Then
            For i As Integer = 0 To hTable.Children(0).Children.Count - 1
                Dim hTD As HtmlElement = hTable.Children(0).Children(i).Children(0)
                If hTD.GetAttribute("ClassName") = "selected" Then
                    paramLoop = hTD.GetAttribute("rVal")
                    Exit For
                End If
            Next
        End If

        'получаем задержку
        paramDelay = "85"
        hTable = hDocument.GetElementById("paramDelay")
        If IsNothing(hTable) = False Then
            For i As Integer = 0 To hTable.Children(0).Children.Count - 1
                Dim hTD As HtmlElement = hTable.Children(0).Children(i).Children(0)
                If hTD.GetAttribute("ClassName") = "selected" Then
                    paramDelay = hTD.GetAttribute("rVal")
                    Exit For
                End If
            Next
        End If

    End Sub

    Private Sub hDocument_Click(sender As Object, e As HtmlElementEventArgs) Handles hDocument.Click
        If wbMarquee.ReadyState <> WebBrowserReadyState.Complete Then Return
        Dim hEl As HtmlElement = hDocument.GetElementFromPoint(e.ClientMousePosition)

        Do While Not hEl.TagName = "BODY"
            Dim rVal As String = hEl.GetAttribute("rVal")
            If String.IsNullOrEmpty(rVal) = False AndAlso hEl.TagName = "TD" Then
                'это TD с аттрибутом rVal
                'получаем тип по id таблицы (TD-TR-TBODY-TABLE)
                Dim hTBody As HtmlElement = hEl.Parent.Parent
                For i As Integer = 0 To hTBody.Children.Count - 1
                    'убираем класс selected со всех позиций данного типа
                    Dim hTD As HtmlElement = hTBody.Children(i).Children(0) 'TD
                    hTD.SetAttribute("ClassName", "")
                Next i
                hEl.SetAttribute("ClassName", "selected") 'выделяем выбранную
                UpdateMarqueSample() 'обновляем
                Return
            End If

            hEl = hEl.Parent
        Loop
    End Sub

    Private Sub wbMarquee_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbMarquee.DocumentCompleted
        hDocument = wbMarquee.Document
        If IsNothing(hDocument) Then Return
        UpdateMarqueSample()
    End Sub
End Class
