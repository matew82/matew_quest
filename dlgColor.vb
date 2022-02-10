Imports System.Windows.Forms

Public Class dlgColor
    Private WithEvents hDocument As HtmlDocument
    Public Property SelectedColor As String = ""

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If IsNothing(hDocument) = False Then
            Dim hSample As HtmlElement = hDocument.GetElementById("colorSample")
            If IsNothing(hSample) = False Then
                'Добавляем цвет в список выбранных
                Dim strColor As String = hSample.Style
                If strColor.StartsWith("background-color: rgb") = False AndAlso strColor.StartsWith("background-color: #") = False Then Return
                strColor = strColor.Substring("background-color: ".Length).Trim
                If strColor.EndsWith(";") Then strColor = strColor.Substring(0, strColor.Length - 1).Trim
                If questEnvironment.lstSelectedColors.IndexOf(strColor) = -1 Then questEnvironment.lstSelectedColors.Add(strColor)
            End If
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgColor_Load(sender As Object, e As EventArgs) Handles Me.Load
        splitMain.FixedPanel = FixedPanel.Panel2
        wbColor.Navigate(FileIO.FileSystem.CombinePath(Application.StartupPath, "src\color_selection.html"))
    End Sub

    Private Sub wbColor_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbColor.DocumentCompleted
        hDocument = wbColor.Document
        Dim hDataContainer As HtmlElement = hDocument.GetElementById("dataContainer")
        If IsNothing(hDataContainer) Then Return
        Dim mapPath As String = Application.StartupPath + "\src\img\colorMap.bmp"
        If My.Computer.FileSystem.FileExists(mapPath) Then
            'получаем элемент с картой цвета
            Dim hMap As HtmlElement = hDocument.GetElementById("colorMap")
            If IsNothing(hMap) Then Return
            hMap.SetAttribute("src", mapPath)
            'получаем элемент для вывода сэмпла цвета
            Dim hSample As HtmlElement = hDocument.GetElementById("colorSample")
            If IsNothing(hSample) Then Return
            'выводим в него текущий выбранный цвет
            SelectedColor = "rgba(0,0,0,1)"

            hSample.Style = "background-color:" + SelectedColor
            'получаем элемент-текстбокс для ввода прозрачности
            Dim hTransp As HtmlElement = hDocument.GetElementById("colorTransparency")
            If IsNothing(hTransp) = False Then
                AddHandler hMap.Click, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                           'устанавливаем цвет пикселя, по которому кликнули
                                           Dim curCol As Color = frmMainEditor.imgSelectColorMap.GetPixel(e2.OffsetMousePosition.X, e2.OffsetMousePosition.Y)
                                           Dim transp As Double = 1
                                           If IsNothing(transp) = False Then
                                               'получаем прозрачность
                                               Double.TryParse(hTransp.GetAttribute("Value").Replace(",", "."), System.Globalization.NumberStyles.Float, provider_points, transp)
                                           End If
                                           'Dim strRGBA As String = "rgba(" + curCol.R.ToString + "," + curCol.G.ToString + "," + curCol.B.ToString + ", " + Convert.ToString(transp, provider_points) + ")"
                                           SelectedColor = colorToString(curCol.R, curCol.G, curCol.B, transp)
                                           hSample.Style = "background-color: " + SelectedColor
                                           hTransp.SetAttribute("value", "1")
                                       End Sub
                AddHandler hTransp.LosingFocus, AddressOf del_hTransp

                AddHandler hTransp.KeyDown, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                If e2.KeyPressedCode = Keys.Return Then del_hTransp(sender2, e2)
                                            End Sub
            End If
            'выводим выбранные цвета
            Dim hColorsContainer As HtmlElement = hDocument.GetElementById("selectedColorsContainer")
            If IsNothing(hColorsContainer) Then Return
            For i As Integer = 0 To questEnvironment.lstSelectedColors.Count - 1
                Dim hColor As HtmlElement = hDocument.CreateElement("SPAN")
                hColor.SetAttribute("className", "selectedColor")
                hColor.SetAttribute("Title", questEnvironment.lstSelectedColors(i))
                hColor.Style = "background-color: " + questEnvironment.lstSelectedColors(i)
                hColorsContainer.AppendChild(hColor)
                AddHandler hColor.Click, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                             SelectedColor = sender2.GetAttribute("Title")
                                             hSample.Style = "background-color:" + SelectedColor
                                             If SelectedColor.StartsWith("#") OrElse SelectedColor.StartsWith("rgb(") Then
                                                 hTransp.SetAttribute("value", "1")
                                             ElseIf SelectedColor.StartsWith("rgba(") Then
                                                 Dim val As String = SelectedColor
                                                 Dim pos As Integer = val.LastIndexOf(")"c)
                                                 If pos = -1 Then Return
                                                 val = val.Substring(0, pos)
                                                 pos = val.LastIndexOf(","c)
                                                 If pos = -1 Then Return
                                                 val = val.Substring(pos + 1).Trim
                                                 hTransp.SetAttribute("value", val)
                                             End If
                                         End Sub
            Next

            Dim fileContents As String
            fileContents = My.Computer.FileSystem.ReadAllText(Application.StartupPath + "\src\colors.html")


            Dim hFrame As HtmlElement = hDocument.CreateElement("DIV")
            hFrame.InnerHtml = fileContents
            hDataContainer.AppendChild(hFrame)
            hFrame.Id = "colorsFrame"
        End If

    End Sub

    Private Sub del_hTransp(sender As Object, e As HtmlElementEventArgs)
        If IsNothing(hDocument) Then Return
        Dim hSample As HtmlElement = hDocument.GetElementById("colorSample")
        If IsNothing(hSample) Then Return
        Dim strColor As String = hSample.Style
        If strColor.StartsWith("background-color: rgb") = False AndAlso strColor.StartsWith("background-color: #") = False Then Return
        If strColor.StartsWith("background-color: rgba(") Then
            strColor = strColor.Substring("background-color: rgba(".Length)
        ElseIf strColor.StartsWith("background-color: rgb(") Then
            strColor = strColor.Substring("background-color: rgb(".Length)
        ElseIf strColor.StartsWith("background-color: #") Then
            strColor = strColor.Substring("background-color: ".Length)
        End If
        If strColor.EndsWith(";") Then strColor = strColor.Substring(0, strColor.Length - 1)
        If strColor.EndsWith(")") Then strColor = strColor.Substring(0, strColor.Length - 1)
        Dim arrColors() As String = strColor.Split(","c)
        If arrColors.Count < 3 Then Return
        Dim r As Byte, g As Byte, b As Byte, alpha As Double
        Byte.TryParse(arrColors(0).Trim, r)
        Byte.TryParse(arrColors(1).Trim, g)
        Byte.TryParse(arrColors(2).Trim, b)

        Double.TryParse(sender.GetAttribute("Value").Replace(",", "."), System.Globalization.NumberStyles.Float, provider_points, alpha)
        'Dim strRGBA As String = "rgba(" + r.ToString + "," + g.ToString + "," + b.ToString + ", " + Convert.ToString(alpha, provider_points) + ")"
        SelectedColor = colorToString(r, g, b, alpha)
        hSample.Style = "background-color:" + SelectedColor
    End Sub

    ''' <summary>
    ''' Возвращает цвет в формате #rrggbb (если цвет непрозрачный) или в rgba
    ''' </summary>
    ''' <param name="r">красный</param>
    ''' <param name="g">зеленый</param>
    ''' <param name="b">синий</param>
    ''' <param name="alpha">прозрачность (от 0 до 1)</param>
    Private Function colorToString(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal alpha As Double) As String
        If alpha = 1 Then
            Dim strR As String = Format(r, "x")
            If strR.Length = 1 Then strR = "0" + strR
            Dim strG As String = Format(g, "x")
            If strG.Length = 1 Then strG = "0" + strG
            Dim strB As String = Format(b, "x")
            If strB.Length = 1 Then strB = "0" + strB

            Return "#" + strR + strG + strB
        Else
            Return "rgba(" + r.ToString + "," + g.ToString + "," + b.ToString + ", " + Convert.ToString(alpha, provider_points) + ")"
        End If
    End Function

    Private Sub hDocument_Click(sender As Object, e As HtmlElementEventArgs) Handles hDocument.Click
        If wbColor.ReadyState <> WebBrowserReadyState.Complete Then Return
        Dim hEl As HtmlElement = hDocument.GetElementFromPoint(e.ClientMousePosition)

        Do While Not hEl.TagName = "BODY"
            Dim res As String = hEl.GetAttribute("rVal")
            If res.Length > 0 AndAlso res.ToLower = "[title]" Then res = hEl.GetAttribute("Title")

            If hEl.GetAttribute("bgcolor").Length > 0 AndAlso res.Length = 7 AndAlso res.StartsWith("#") Then
                'надо отобразить выбранный цвет в colorSample
                Dim hSample As HtmlElement = hDocument.GetElementById("colorSample")
                Dim hTransp As HtmlElement = hDocument.GetElementById("colorTransparency")
                If IsNothing(hSample) = False Then
                    Dim r As Byte, g As Byte, b As Byte
                    Byte.TryParse(res.Substring(1, 2), System.Globalization.NumberStyles.HexNumber, provider_points, r)
                    Byte.TryParse(res.Substring(3, 2), System.Globalization.NumberStyles.HexNumber, provider_points, g)
                    Byte.TryParse(res.Substring(5, 2), System.Globalization.NumberStyles.HexNumber, provider_points, b)
                    'Dim strARGB As String = "rgba(" + r.ToString + "," + g.ToString + "," + b.ToString + ",1)"
                    SelectedColor = colorToString(r, g, b, 1)
                    hSample.Style = "background-color:" + SelectedColor
                    If IsNothing(hTransp) = False Then hTransp.SetAttribute("value", "1")
                End If
            End If

            hEl = hEl.Parent
        Loop


    End Sub
End Class
