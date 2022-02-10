''' <summary>
''' Класс построения у работы с картами
''' </summary>
Public Class cMapManager
    Public Enum DirectionEnum As Byte
        FORWARD = 0
        BACKWARD = 1
        RIGHT = 2
        LEFT = 3
    End Enum

    Public Enum BuildStyleEnum As Byte
        AllCells = 0
        OnlyVisible = 1
        OnlyVisited = 2
    End Enum


#Region "Editor part"
    ''' <summary>
    ''' Строит карту для добавления новых клеток / перемещения текущей клетки на новое место
    ''' </summary>
    ''' <param name="curCellName">Имя текущей клетки. Если не указано, то карта откроется в режиме создания новых клеток</param>
    Public Sub BuildMapForCellsEdit(Optional curCellName As String = "")
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim mapId As Integer

        If currentParentName.Length > 0 Then
            'режим редактора
            mapId = GetSecondChildIdByName(currentParentName, mScript.mainClass(classMap).ChildProperties)
        Else
            'режим создания клеток
            mapId = cPanelManager.ActivePanel.GetChild2Id
        End If
        If mapId < 0 Then Return

        'открываем докумет, внутри которого будет построена карта
        Try
            frmMainEditor.WBhelp.Navigate(Application.StartupPath + "\src\editor.html")
        Catch ex As Exception
            MessageBox.Show("Ошибка при открытии файла src\editor.html!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
        Do While Not frmMainEditor.WBhelp.ReadyState = WebBrowserReadyState.Complete
            Application.DoEvents()
        Loop
        Dim hDoc As HtmlDocument = frmMainEditor.WBhelp.Document
        Dim hBoby As HtmlElement = hDoc.Body

        'размеры карты
        Dim cellsByX As Integer = GetPropertyValueAsInteger(classMap, "CellsByX", mapId, -1, 10)
        Dim cellsByY As Integer = GetPropertyValueAsInteger(classMap, "CellsByY", mapId, -1, 10)

        'Пишем заголовок
        Dim hEl As HtmlElement = hDoc.CreateElement("H1")
        If curCellName.Length > 0 Then
            hEl.InnerHtml = "Редактирование клетки карты " + mScript.PrepareStringToPrint(curCellName, Nothing, False)
        Else
            hEl.InnerHtml = "Создание клеток карты"
        End If
        hBoby.AppendChild(hEl)

        'создаем элемент ЦЕНТР, внутри которого создаем табдицу - карту
        Dim hCenter As HtmlElement = hDoc.CreateElement("CENTER")
        hBoby.AppendChild(hCenter)

        'Dim hSpan As HtmlElement = hDoc.CreateElement("SPAN")
        'hSpan.Style = "float:right"
        'hSpan.Id = "sInfo"
        'hEl.AppendChild(hSpan)
        Dim hChkPreview As HtmlElement = hDoc.CreateElement("INPUT")
        hChkPreview.SetAttribute("Type", "checkbox")
        hChkPreview.Id = "chkPreview"
        hChkPreview.Style = "display: none"
        hCenter.AppendChild(hChkPreview)
        Dim hLblPreview As HtmlElement = hDoc.CreateElement("LABEL")
        hLblPreview.SetAttribute("for", "chkPreview")
        hLblPreview.Id = "lblPreview"
        hLblPreview.SetAttribute("ClassName", "lblPreview")
        hLblPreview.InnerHtml = "Предварительный просмотр"
        hCenter.AppendChild(hLblPreview)
        Dim hBr As HtmlElement = hDoc.CreateElement("DIV")
        hBr.Style = "padding:20px"
        hCenter.AppendChild(hBr)

        Dim hTable As HtmlElement = hDoc.CreateElement("TABLE")
        If curCellName.Length > 0 Then
            'режим редактирования
            hTable.Id = "MapGoToCell"
        Else
            'режим создания
            hTable.Id = "MapAddCells"
        End If
        hCenter.AppendChild(hTable)
        Dim hTBody As HtmlElement = hDoc.CreateElement("TBODY")
        hTable.AppendChild(hTBody)

        'Получаем Id текущей клетки (если такая есть)
        Dim curCellId As Integer = -1
        If curCellName.Length > 0 Then curCellId = GetThirdChildIdByName(curCellName, mapId, mScript.mainClass(classMap).ChildProperties)
        Dim tr As HtmlElement, td As HtmlElement
        For y As Integer = 1 To cellsByY
            'строим горизонтальные ряды
            tr = hDoc.CreateElement("TR")
            hTBody.AppendChild(tr)
            tr.Style = "height:" + GetRowHeight(mapId, y)
            For x As Integer = 1 To cellsByX
                'строим вертикальные колонки
                Dim cellId As Integer = GetCellIdByXYEditor(mapId, x, y) 'Id клетки по текущим координатам или -1, если там пусто
                td = hDoc.CreateElement("TD")
                td.Style = "width:" + GetColumnWidth(mapId, x)
                td.SetAttribute("x", x.ToString)
                td.SetAttribute("y", y.ToString)
                td.SetAttribute("mapId", mapId.ToString)
                td.SetAttribute("cellId", cellId.ToString)
                If cellId = -1 Then
                    'По данным координатам клетки не существует
                    If curCellId = -1 Then
                        'режим создания - при нажатии на пустую клетку будет создана новая клетка
                        AddHandler td.Click, AddressOf del_cells_addNew
                    Else
                        'режим редактора - при нажатии на пустую клетку текущая клетка переместится на это место
                        AddHandler td.Click, AddressOf del_cells_ChangePos
                    End If
                Else
                    Dim cellLoc As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("Location").ThirdLevelProperties(cellId), Nothing, False)
                    Dim cellName As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("Name").ThirdLevelProperties(cellId), Nothing, False)
                    'Клетка уже существует
                    If cellId = curCellId Then
                        'клетка - текущая
                        If cellLoc.Length > 0 Then
                            td.SetAttribute("ClassName", "currentCellWithLocation")
                        Else
                            td.SetAttribute("ClassName", "currentCell")
                        End If
                        hCurrentCell = td
                    Else
                        If cellLoc.Length > 0 Then
                            td.SetAttribute("ClassName", "existedCellWithLocation")
                        Else
                            td.SetAttribute("ClassName", "existedCell")
                        End If
                    End If
                    If cellLoc.Length = 0 Then cellLoc = "локация не выбрана"
                    td.SetAttribute("Title", cellName + " (" + cellLoc + ")")

                    'при нажатии на заполенную клетку будет совершен переход к ней
                    AddHandler td.Click, AddressOf del_cells_GoTo
                End If
                tr.AppendChild(td)
            Next x
        Next y

        'Создаем фрейм для отображения реальной карты
        Dim classLW As Integer = mScript.mainClassHash("LW")
        Dim mWidth As Integer = GetPropertyValueAsInteger(classLW, "Width", -1, -1, frmMainEditor.WBhelp.ClientSize.Width - 30)
        Dim mHeight As Integer = GetPropertyValueAsInteger(classLW, "Height", -1, -1, 800)
        Dim hFrame As HtmlElement = hDoc.CreateElement("IFRAME")
        hFrame.SetAttribute("Width", mWidth.ToString + "px")
        hFrame.SetAttribute("Height", mHeight.ToString + "px")
        'hFrame.Style = "display:none"
        hFrame.Id = "frmMap"
        hFrame.Style = "display:none"
        hCenter.AppendChild(hFrame)

        hEl = hDoc.CreateElement("P")
        If curCellName.Length > 0 Then
            hEl.InnerHtml = "Для изменения расположения клетки на карте выберите другую свободную клетку."
        Else
            hEl.InnerHtml = "Для создания новой клетки карты достаточно выбрать соответствующую ей позицию в таблице выше. Предварительно рекомендуется настроить все характеристики в свойствах данной карты, а также в настройках по умолчанию."
        End If
        hBoby.AppendChild(hEl)

        AddHandler hLblPreview.Click, Sub(sender As HtmlElement, hE As HtmlElementEventArgs)
                                          Dim chk As String = hChkPreview.GetAttribute("Checked")
                                          If chk = "True" Then
                                              hChkPreview.SetAttribute("Checked", "")
                                              hFrame.Style = "display:none"
                                              hTable.Style = ""
                                              sender.SetAttribute("ClassName", "lblPreview")
                                          Else
                                              hChkPreview.SetAttribute("Checked", "True")
                                              hFrame.Style = ""
                                              hTable.Style = "display:none"
                                              BuildMapInEditor(hDoc.Window.Frames("frmMap"), mapId, curCellId)
                                              sender.SetAttribute("ClassName", "lblPreviewChecked")
                                          End If
                                      End Sub

        If frmMainEditor.codeBoxPanel.Visible Then frmMainEditor.codeBoxPanel.Hide()
        frmMainEditor.WBhelp.Show()
        frmMainEditor.treeProperties.Hide()
        frmMainEditor.abortNavigationEvent = True
    End Sub

    Public Sub BuildMapInEditor(ByRef frameWnd As HtmlWindow, ByVal mapId As Integer, ByVal curCellId As Integer)
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim mWidth As Integer = GetPropertyValueAsInteger(classMap, "Width", -1, -1, frmMainEditor.WBhelp.ClientSize.Width - 30)
        Dim mHeight As Integer = GetPropertyValueAsInteger(classMap, "Height", -1, -1, 800)
        Dim hFrame As HtmlElement = frmMainEditor.WBhelp.Document.GetElementById("frmMap")
        hFrame.SetAttribute("Width", mWidth.ToString + "px")
        hFrame.SetAttribute("Height", mHeight.ToString + "px")

        frmMainEditor.abortNavigationEvent = True
        frameWnd.Navigate(questEnvironment.QuestPath + "\Map.html")
        Dim cnt As Integer = 0
        Dim mapConvas As HtmlElement = Nothing
        Do While IsNothing(mapConvas)
            Application.DoEvents()
            mapConvas = frameWnd.Document.GetElementById("MapConvas")
            cnt += 1
            If cnt > 100000 Then Exit Do
        Loop
        Do While frmMainEditor.WBhelp.IsBusy
            Application.DoEvents()
        Loop
        mapConvas = frameWnd.Document.GetElementById("MapConvas")

        If IsNothing(mapConvas) Then
            MessageBox.Show("Ошибка в структуре файла " + questEnvironment.QuestPath + "\Map.html. Файл в обязательном порядке должен иметь html-элемент с id = 'MapConvas'.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        mapConvas.InnerHtml = ""

        'устанавливаем css
        Dim strCSS As String = mScript.mainClass(classMap).ChildProperties(mapId)("CSS").Value
        Dim retType As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(strCSS)
        If retType = MatewScript.ContainsCodeEnum.NOT_CODE Then
            strCSS = UnWrapString(strCSS)
            HtmlChangeFirstCSSLink(frameWnd.Document, strCSS)
        End If

        'BackColor
        Dim strStyle As New System.Text.StringBuilder
        Dim propValue As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("BackColor").Value, Nothing)
        If String.IsNullOrEmpty(propValue) = False Then strStyle.Append(" background-color:" + propValue + ";")
        'BackPicture
        propValue = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("BackPicture").Value, Nothing)
        If propValue.Length > 0 Then
            strStyle.AppendFormat("background-image: url({0});", "'" + propValue.Replace("\", "/") + "'")
            propValue = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("BackPicStyle").Value, Nothing)
            '0 простая загрузка, 1 - заполнить, 2 - масштабировать, 3 - размножить, 4 - размножить по Х, 5 - размножить по Y 
            strStyle.Append("background-repeat:")
            Dim propInt As Integer = Val(propValue)
            Select Case propInt
                Case 0 '0 простая загрузка
                    strStyle.Append("no-repeat;")
                Case 1 '1 растянуть пропорционально
                    strStyle.Append("no-repeat;background-size:cover;")
                Case 2 '2 заполнить
                    strStyle.Append("no-repeat;background-size:contain;")
                Case 3 '3 масштабировать
                    strStyle.Append("repeat;")
                Case 4 '4 размножить по Х
                    strStyle.Append("repeat-x;")
                Case 5 '5 размножить по Y
                    strStyle.Append("repeat-y;")
            End Select

            If propInt = 0 Then
                'BackPicPos
                propValue = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("BackPicPos").Value, Nothing)
                '0 в левом верхнем углу, 1 слева по центру, 2 в левом нижнем углу, 3 сверху по центру, 4 в центре, 5 снизу по центру, 6 в правом верхнем углу, 7 справа по центру, 8 в правом нижнем углу
                strStyle.Append("background-position:")
                propInt = Val(propValue)
                Select Case propInt
                    Case 0 '0 в левом верхнем углу
                        strStyle.Append("left top;")
                    Case 1 '1 слева по центру
                        strStyle.Append("left center;")
                    Case 2 '2 в левом нижнем углу
                        strStyle.Append("left bottom;")
                    Case 3 '3 сверху по центру
                        strStyle.Append("center top;")
                    Case 4 '4 в центре
                        strStyle.Append("center center;")
                    Case 5 '5 снизу по центру
                        strStyle.Append("center bottom;")
                    Case 6 '6 в правом верхнем углу
                        strStyle.Append("right top;")
                    Case 7 '7 справа по центру
                        strStyle.Append("right center;")
                    Case 8 '8 в правом нижнем углу
                        strStyle.Append("right bottom;")
                End Select
            End If
        End If
        mapConvas.Style = strStyle.ToString


        '''''''Создание карты
        Dim hDoc As HtmlDocument = frameWnd.Document
        'размеры карты
        Dim cellsByX As Integer = GetPropertyValueAsInteger(classMap, "CellsByX", mapId, -1, 10)
        Dim cellsByY As Integer = GetPropertyValueAsInteger(classMap, "CellsByY", mapId, -1, 10)

        'Пишем заголовок
        Dim hEl As HtmlElement = Nothing
        propValue = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("Caption").Value, Nothing)
        If propValue.Length > 0 Then
            hEl = hDoc.CreateElement("H1")
            hEl.Id = "MapCaption"
            hEl.InnerHtml = propValue
            mapConvas.AppendChild(hEl)
        End If

        'создаем элемент ЦЕНТР, внутри которого создаем табдицу - карту
        Dim hCenter As HtmlElement = hDoc.CreateElement("CENTER")
        mapConvas.AppendChild(hCenter)

        Dim hTable As HtmlElement = hDoc.CreateElement("TABLE")
        hTable.Id = "MapMain"
        hTable.SetAttribute("cellpadding", "0")
        hTable.SetAttribute("cellspacing", "0")

        hCenter.AppendChild(hTable)
        Dim hTBody As HtmlElement = hDoc.CreateElement("TBODY")
        hTable.AppendChild(hTBody)

        'Картинка карты MapPicture
        propValue = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("MapPicture").Value, Nothing)
        If propValue.Length > 0 Then
            strStyle.Clear()
            strStyle.AppendFormat("background-image: url({0});", "'" + propValue.Replace("\", "/") + "'")
            propValue = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("MapPicStyle").Value, Nothing)
            '0 простая загрузка, 1 - заполнить, 2 - масштабировать, 3 - размножить, 4 - размножить по Х, 5 - размножить по Y 
            strStyle.Append("background-repeat:")
            Dim propInt As Integer = Val(propValue)
            Select Case propInt
                Case 0 '0 простая загрузка
                    strStyle.Append("no-repeat;")
                Case 1 '1 растянуть пропорционально
                    strStyle.Append("no-repeat;background-size:cover;")
                Case 2 '2 заполнить
                    strStyle.Append("no-repeat;background-size:contain;")
                Case 3 '3 масштабировать
                    strStyle.Append("repeat;")
                Case 4 '4 размножить по Х
                    strStyle.Append("repeat-x;")
                Case 5 '5 размножить по Y
                    strStyle.Append("repeat-y;")
            End Select
            hTable.Style = strStyle.ToString
        End If

        'расположение карты
        Dim mapOffsetX As Integer = 0, mapOffsetY As Integer = 0
        mapOffsetX = Val(mScript.mainClass(classMap).ChildProperties(mapId)("MapOffsetX").Value)
        mapOffsetY = Val(mScript.mainClass(classMap).ChildProperties(mapId)("MapOffsetY").Value)
        Dim mStyle As String = hTable.Style
        strStyle.Clear()
        strStyle.Append(mStyle)
        Dim mapPositionWasSet As Boolean = False
        If mapOffsetX <> 0 Then
            strStyle.Append("left:" & mapOffsetX.ToString & "px;")
            mapPositionWasSet = True
        End If
        If mapOffsetY <> 0 Then
            strStyle.Append("top:" & mapOffsetY.ToString & "px;")
            mapPositionWasSet = True
        End If
        If mapPositionWasSet Then
            strStyle.Append("position:absolute")
            hTable.Style = strStyle.ToString
        End If

        Dim cellClassEmpty As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("CellEmptyClass").Value, Nothing)

        'Получаем Id текущей клетки (если такая есть)
        Dim tr As HtmlElement, td As HtmlElement
        For y As Integer = 1 To cellsByY
            'строим горизонтальные ряды
            tr = hDoc.CreateElement("TR")
            hTBody.AppendChild(tr)
            tr.Style = "height:" + GetRowHeight(mapId, y)
            For x As Integer = 1 To cellsByX
                'строим вертикальные колонки
                Dim cellId As Integer = GetCellIdByXYEditor(mapId, x, y) 'Id клетки по текущим координатам или -1, если там пусто
                td = hDoc.CreateElement("TD")
                td.Style = "width:" + GetColumnWidth(mapId, x)
                td.SetAttribute("x", x.ToString)
                td.SetAttribute("y", y.ToString)
                td.SetAttribute("mapId", mapId.ToString)
                td.SetAttribute("cellId", cellId.ToString)
                tr.AppendChild(td)
                If cellId = -1 Then
                    'По данным координатам клетки не существует
                    td.SetAttribute("ClassName", cellClassEmpty)
                Else
                    'Dim strContent As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("Content").ThirdLevelProperties(cellId), Nothing)
                    Dim strContent As String = ""
                    ReadProperty(classMap, "Content", mapId, cellId, strContent, Nothing)
                    strContent = UnWrapString(strContent)
                    mScript.LAST_ERROR = ""
                    If strContent.Length > 0 Then td.InnerHtml = strContent
                    'Клетка уже существует
                    Dim cellClassUsual As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("CellUsualClass").ThirdLevelProperties(cellId), Nothing)
                    td.SetAttribute("ClassName", cellClassUsual)

                    If cellId = curCellId Then
                        'клетка - текущая
                        Dim hCur As HtmlElement = hDoc.CreateElement("DIV")
                        Dim hPos As Point = GetHTMLelementCoordinates(td)
                        Dim hSize As Size = td.OffsetRectangle.Size
                        Dim cellClassCurrent As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("CellCurrentClass").Value, Nothing)
                        hCur.Style = "position:absolute;left:" & hPos.X.ToString & "px;top:" & hPos.Y.ToString & "px;width:" & hSize.Width.ToString & "px;height:" & hSize.Height.ToString & "px;overflow:hidden;"
                        hCur.SetAttribute("ClassName", cellClassCurrent)
                        hDoc.Body.AppendChild(hCur)
                    End If
                End If
            Next x
        Next y

        'устанавливаем изображения клеток
        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMap).ChildProperties(mapId)("CellPicture")
        If IsNothing(ch.ThirdLevelProperties) = False Then
            For cellId As Integer = ch.ThirdLevelProperties.Count - 1 To 0 Step -1
                'получаем картинку
                Dim imgPath As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("CellPicture").ThirdLevelProperties(cellId), Nothing)
                If String.IsNullOrEmpty(imgPath) Then Continue For
                Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, imgPath)
                If FileIO.FileSystem.FileExists(fPath) = False Then Continue For

                'создаем элемент картинки и задаем его стиль и класс
                Dim hImg As HtmlElement = hDoc.CreateElement("IMG")
                hImg.SetAttribute("ClassName", "CellImage")
                hImg.SetAttribute("src", imgPath)

                Dim pWidth As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("CellPictureWidth").ThirdLevelProperties(cellId), Nothing)
                Dim pHeight As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("CellPictureHeight").ThirdLevelProperties(cellId), Nothing)
                If IsNumeric(pWidth) Then pWidth &= "px"
                If IsNumeric(pHeight) Then pHeight &= "px"

                strStyle.Clear()
                If String.IsNullOrEmpty(pWidth) = False Then strStyle.Append("width:" & pWidth & ";")
                If String.IsNullOrEmpty(pHeight) = False Then strStyle.Append("height:" & pHeight & ";")

                Dim hCell As HtmlElement = GetCellHtmlElement(cellId, hDoc)
                hDoc.Body.AppendChild(hImg)

                'устанавливаем положение картинки
                Dim offsetX As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("CellPictureOffsetX").ThirdLevelProperties(cellId), Nothing)
                Dim offsetY As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("CellPictureOffsetY").ThirdLevelProperties(cellId), Nothing)
                Dim locCell As Point = GetHTMLelementCoordinates(hCell)
                Dim locImg As Point = GetHTMLelementCoordinates(hImg)
                'offsetX = offsetX - locImg.X + locCell.X
                'offsetY = offsetY - locImg.Y + locCell.Y
                offsetX = offsetX + locCell.X - 10
                offsetY = offsetY + locCell.Y

                strStyle.Append("left:" & offsetX.ToString & "px;top:" & offsetY.ToString & "px")
                hImg.Style = strStyle.ToString
            Next cellId
        End If
    End Sub

    ''' <summary>
    ''' Делегат html-события при клике на заполенную клетку. Совершает переход на выбранную клетку.
    ''' </summary>
    ''' <param name="sender">html-элемент-клетка</param>
    ''' <param name="e"></param>
    Private Sub del_cells_GoTo(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim mapId As Integer = sender.GetAttribute("mapId")
        Dim cellId As Integer = sender.GetAttribute("cellId")
        If mapId < 0 OrElse cellId < 0 Then Return

        If currentParentName.Length > 0 Then
            'сейчас работа с клетками
            cPanelManager.FindAndOpen(classMap, mapId, cellId, -1, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
        Else
            'сейчас просмотр клеток из поддерева
            Dim tree As TreeView = cPanelManager.dictDefContainers(classMap).subTree
            Dim n As TreeNode = frmMainEditor.FindItemNodeByText(tree, mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("Name").ThirdLevelProperties(cellId), Nothing, False))
            If IsNothing(n) = False Then tree.SelectedNode = n
        End If
    End Sub

    Private hCurrentCell As HtmlElement = Nothing 'текущая клетка таблицы (соответствующая текущей клетке)
    ''' <summary>
    ''' Делегат html-события при клике на пустую клетку в режиме редактора. Перемещает координаты выбранной клетки на новое место.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub del_cells_ChangePos(sender As Object, e As HtmlElementEventArgs)
        Dim hNewCell As HtmlElement = sender
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim mapId As Integer = CInt(sender.GetAttribute("mapId"))
        Dim cellId As Integer = -1
        Dim oldCellFound As Boolean = True
        Try
            cellId = CInt(hCurrentCell.GetAttribute("cellId"))
        Catch ex As Exception
            'Старая клетка не найдена. Такое возможно, если ее координаты выходят за пределы карты
            'Получаем cellId исходя из глобальных данных
            If IsNothing(cPanelManager.ActivePanel) = False AndAlso currentClassName = "Map" Then
                oldCellFound = False
                If currentParentName.Length > 0 Then
                    cellId = cPanelManager.ActivePanel.GetChild3Id
                Else
                    Dim tree As TreeView = cPanelManager.dictDefContainers(classMap).subTree
                    If IsNothing(tree.SelectedNode) Then Return
                    Dim ni As frmMainEditor.NodeInfo = frmMainEditor.GetNodeInfo(tree.SelectedNode)
                    cellId = ni.GetChild3Id(mapId)
                End If
            Else
                Return
            End If
        End Try

        'Делаем текущую клетку пустой
        If IsNothing(hCurrentCell) = False AndAlso oldCellFound Then
            AddHandler hCurrentCell.Click, AddressOf del_cells_ChangePos
            hCurrentCell.SetAttribute("ClassName", "")
            hCurrentCell.SetAttribute("cellId", "-1")
        End If

        'Устанавливаем новые параметры на новую выбранную клетку
        Dim newX As Integer = GetPropertyValueAsInteger(classMap, "CellX", mapId, cellId, -1)
        Dim newY As Integer = GetPropertyValueAsInteger(classMap, "CellY", mapId, cellId, -1)
        RemoveHandler hNewCell.Click, AddressOf del_cells_ChangePos
        Dim cellLoc As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("Location").ThirdLevelProperties(cellId), Nothing, False)
        If cellLoc.Length > 0 Then
            hNewCell.SetAttribute("ClassName", "currentCellWithLocation")
        Else
            hNewCell.SetAttribute("ClassName", "currentCell")
        End If
        hNewCell.SetAttribute("cellId", cellId.ToString)
        mScript.mainClass(classMap).ChildProperties(mapId)("CellX").ThirdLevelProperties(cellId) = hNewCell.GetAttribute("x")
        mScript.mainClass(classMap).ChildProperties(mapId)("CellY").ThirdLevelProperties(cellId) = hNewCell.GetAttribute("y")
        hCurrentCell = hNewCell
    End Sub

    ''' <summary>
    ''' Делегат html-события при клике на пустую клетку в режиме создания. Создает новую клетку на указанном месте
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub del_cells_addNew(sender As Object, e As HtmlElementEventArgs)
        Dim hCell As HtmlElement = sender
        Dim className As String = hCell.GetAttribute("ClassName")
        If className = "existedCell" OrElse className = "existedCellWithLocation" Then Return
        Dim tree As TreeView

        'Получаем Id карты и клетки
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim mapId As Integer
        Dim cellId As Integer

        If currentParentName.Length > 0 Then
            mapId = currentParentId
            tree = currentTreeView
        Else
            mapId = cPanelManager.ActivePanel.GetChild2Id
            tree = cPanelManager.dictDefContainers(classMap).subTree
        End If

        'Создает клетку карты
        Dim cellName As String = frmMainEditor.AddElement("Map", tree, Nothing, "", False)
        cellId = GetThirdChildIdByName(cellName, mapId, mScript.mainClass(classMap).ChildProperties)
        If cellName.Length = 0 Then Return

        If mapId < 0 OrElse cellId < 0 Then Return
        'устанавливаем координаты новой клетки
        mScript.mainClass(classMap).ChildProperties(mapId)("CellX").ThirdLevelProperties(cellId) = hCell.GetAttribute("x")
        mScript.mainClass(classMap).ChildProperties(mapId)("CellY").ThirdLevelProperties(cellId) = hCell.GetAttribute("y")
        Dim cellLoc As String = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)("Location").ThirdLevelProperties(cellId), Nothing, False)
        If cellLoc.Length > 0 Then
            hCell.SetAttribute("ClassName", "existedCellWithLocation")
        Else
            hCell.SetAttribute("ClassName", "existedCell")
        End If
        If cellLoc.Length = 0 Then cellLoc = "локация не выбрана"
        hCell.SetAttribute("mapId", mapId.ToString)
        hCell.SetAttribute("cellId", cellId.ToString)
        hCell.SetAttribute("Title", mScript.PrepareStringToPrint(cellName, Nothing, False) + " (" + cellLoc + ")")
        AddHandler hCell.Click, AddressOf del_cells_GoTo
    End Sub

    ''' <summary>
    ''' Возвращает ширину колонки карты для в виде "ХХpx"
    ''' </summary>
    ''' <param name="mapId">Id карты</param>
    ''' <param name="colNumber">Номер колонки, начиная от 1</param>
    Private Function GetColumnWidth(ByVal mapId As Integer, ByVal colNumber As Integer) As String
        Dim propName As String = "CellWidth"
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim res As String = ""

        Dim prop As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMap).ChildProperties(mapId)(propName)
        If IsNothing(prop.ThirdLevelProperties) = False Then
            'Возвращаем ширину первой клетки из нашей колонки, значение CellWidth которой не пустое
            For cellId As Integer = 0 To prop.ThirdLevelProperties.Count - 1
                Dim cellX As Integer = GetPropertyValueAsInteger(classMap, "CellX", mapId, cellId, -1)
                If cellX <> colNumber Then Continue For 'клетка не из нашей колонки - продолжаем
                res = mScript.PrepareStringToPrint(prop.ThirdLevelProperties(cellId), Nothing)
                If res.Length > 0 Then Return res + "px"
            Next cellId
        End If
        'Нужной клетки не найдено. Возвращаем ширину клетки данной карты по умолчанию
        res = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)(propName).Value, Nothing)
        If res.Length > 0 Then Return res + "px"
        'Ширина клеток данной карты не установлена. Получаем ширину клеток карт по умолчанию
        res = mScript.PrepareStringToPrint(mScript.mainClass(classMap).Properties(propName).Value, Nothing)
        If res.Length > 0 Then Return res + "px"
        'И это тоже н установлено. Возвращаем стандартное 40 пикселей
        Return "40px"
    End Function

    ''' <summary>
    ''' Возвращает высоту ряда карты для в виде "ХХpx"
    ''' </summary>
    ''' <param name="mapId">Id карты</param>
    ''' <param name="rowNumber">Номер ряда, начиная от 1</param>
    Private Function GetRowHeight(ByVal mapId As Integer, ByVal rowNumber As Integer) As String
        Dim propName As String = "CellHeight"
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim res As String = ""

        Dim prop As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMap).ChildProperties(mapId)(propName)
        If IsNothing(prop.ThirdLevelProperties) = False Then
            'Возвращаем высоту первой клетки из нашего ряда, значение CellHeight которой не пустое
            For cellId As Integer = 0 To prop.ThirdLevelProperties.Count - 1
                Dim cellY As Integer = GetPropertyValueAsInteger(classMap, "CellY", mapId, cellId, -1)
                If cellY <> rowNumber Then Continue For 'клетка не из нашего ряда - продолжаем
                res = mScript.PrepareStringToPrint(prop.ThirdLevelProperties(cellId), Nothing)
                If res.Length > 0 Then Return res + "px"
            Next cellId
        End If
        'Нужной клетки не найдено. Возвращаем высоту клетки данной карты по умолчанию
        res = mScript.PrepareStringToPrint(mScript.mainClass(classMap).ChildProperties(mapId)(propName).Value, Nothing)
        If res.Length > 0 Then Return res + "px"
        'Ширина клеток данной карты не установлена. Получаем высоту клеток карт по умолчанию
        res = mScript.PrepareStringToPrint(mScript.mainClass(classMap).Properties(propName).Value, Nothing)
        If res.Length > 0 Then Return res + "px"
        'И это тоже н установлено. Возвращаем стандартное 40 пикселей
        Return "40px"
    End Function

    ''' <summary>
    ''' Возвращает Id клетки по ее координатам
    ''' </summary>
    ''' <param name="mapId">Id карты</param>
    ''' <param name="x">Координата Х клетки</param>
    ''' <param name="y">Координата Y клетки</param>
    ''' <returns>Id клметки или -1 если не найдена</returns>
    Public Function GetCellIdByXYEditor(ByVal mapId As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMap).ChildProperties(mapId)("Name")
        If IsNothing(ch.ThirdLevelProperties) OrElse ch.ThirdLevelProperties.Count = 0 Then Return -1

        Dim cX As Integer = -1, cY As Integer = -1
        For i As Integer = 0 To ch.ThirdLevelProperties.Count - 1
            cX = GetPropertyValueAsInteger(classMap, "CellX", mapId, i, -1)
            cY = GetPropertyValueAsInteger(classMap, "CellY", mapId, i, -1)
            If cX = x AndAlso cY = y Then Return i
        Next
        Return -1
    End Function
#End Region

#Region "Game part"
    Public Function BuildMap(ByVal mapId As Integer, ByVal curCellId As Integer, ByRef arrParams() As String, Optional ByVal prevCellId As Integer = -1) As String
        Dim classMap As Integer = mScript.mainClassHash("Map")
        GVARS.G_CURMAP = mapId
        GVARS.G_CURMAPCELL = curCellId
        GVARS.G_PREVMAPCELL = prevCellId

        Dim isNewWindow As Boolean = False
        Dim slot As Integer = 0
        If ReadPropertyInt(classMap, "Slot", -1, -1, slot, arrParams) = False Then Return "#Error"
        If slot = 5 Then isNewWindow = True 'в отдельном окне

        If isNewWindow Then
            'устанавливаем ширину окна
            Dim wndWidth As String = "", wndHeight As String = ""
            If ReadProperty(classMap, "Width", -1, -1, wndWidth, arrParams) = False Then Return "#Error"
            If ReadProperty(classMap, "Height", -1, -1, wndHeight, arrParams) = False Then Return "#Error"
            wndWidth = UnWrapString(wndWidth)
            wndHeight = UnWrapString(wndHeight)
            Dim wSize As New Size
            'width
            If IsNumeric(wndWidth) Then
                wSize.Width = Val(wndWidth) 'px
            ElseIf wndWidth.EndsWith("%") Then
                wSize.Width = Math.Round(My.Computer.Screen.WorkingArea.Width * Val(wndWidth) / 100) '%
            Else
                Return _ERROR("Нераспознан формат ширины окна карт (свойство Map.Width).", "Map Builder")
            End If
            'height
            If IsNumeric(wndHeight) Then
                wSize.Height = Val(wndHeight) 'px
            ElseIf wndHeight.EndsWith("%") Then
                wSize.Height = Math.Round(My.Computer.Screen.WorkingArea.Height * Val(wndHeight) / 100) '%
            Else
                Return _ERROR("Нераспознан формат высоты окна карт (свойство Map.Height).", "Map Builder")
            End If
            frmMap.Size = wSize
        End If

        'получаем конвас карты
        Dim hDoc As HtmlDocument = questEnvironment.wbMap.Document
        If IsNothing(hDoc) Then Return _ERROR("Html-документ карт не загружен.", "Map Builder")
        Dim mapConvas As HtmlElement = Nothing
        mapConvas = hDoc.GetElementById("MapConvas")
        If IsNothing(mapConvas) Then Return _ERROR("Ошибка в структуре файла " + questEnvironment.QuestPath + "\Map.html. Файл в обязательном порядке должен иметь html-элемент с id = 'MapConvas'.", "Map Builder")
        mapConvas.InnerHtml = ""

        'устанавливаем css
        Dim strCSS As String = ""
        If ReadProperty(classMap, "CSS", mapId, -1, strCSS, arrParams) = False Then Return "#Error"
        strCSS = UnWrapString(strCSS)
        HtmlChangeFirstCSSLink(hDoc, strCSS)

        'удаляем старые картинки клеток
        ClearPreviousBitmap()
        For i As Integer = hDoc.Images.Count - 1 To 0 Step -1
            Dim hImg As HtmlElement = hDoc.Images(i)
            Dim strCell As String = hImg.GetAttribute("cellId")
            If String.IsNullOrEmpty(strCell) Then Continue For
            Dim msImg As mshtml.IHTMLDOMNode = hImg.DomElement
            msImg.removeNode(True)
        Next i

        'BackColor
        Dim strStyle As New System.Text.StringBuilder
        Dim bkColor As String = ""
        If ReadProperty(classMap, "BackColor", mapId, -1, bkColor, arrParams) = False Then Return "#Error"
        bkColor = UnWrapString(bkColor)

        If String.IsNullOrEmpty(bkColor) = False Then strStyle.Append(" background-color:" + bkColor + ";")
        'BackPicture
        Dim bkPicture As String = ""
        If ReadProperty(classMap, "BackPicture", mapId, -1, bkPicture, arrParams) = False Then Return "#Error"
        bkPicture = UnWrapString(bkPicture)
        If bkPicture.Length > 0 Then
            strStyle.AppendFormat("background-image: url({0});", "'" + bkPicture.Replace("\", "/") + "'")
            Dim bkPicStyle As Integer = 0
            If ReadPropertyInt(classMap, "BackPicStyle", mapId, -1, bkPicStyle, arrParams) = False Then Return "#Error"
            '0 простая загрузка, 1 - заполнить, 2 - масштабировать, 3 - размножить, 4 - размножить по Х, 5 - размножить по Y 
            strStyle.Append("background-repeat:")
            Select Case bkPicStyle
                Case 0 '0 простая загрузка
                    strStyle.Append("no-repeat;")
                Case 1 '1 растянуть пропорционально
                    strStyle.Append("no-repeat;background-size:cover;")
                Case 2 '2 заполнить
                    strStyle.Append("no-repeat;background-size:contain;")
                Case 3 '3 масштабировать
                    strStyle.Append("repeat;")
                Case 4 '4 размножить по Х
                    strStyle.Append("repeat-x;")
                Case 5 '5 размножить по Y
                    strStyle.Append("repeat-y;")
            End Select

            If bkPicStyle = 0 Then
                'BackPicPos
                Dim bkPicPos As Integer = 0
                If ReadPropertyInt(classMap, "BackPicPos", mapId, -1, bkPicPos, arrParams) = False Then Return "#Error"
                '0 в левом верхнем углу, 1 слева по центру, 2 в левом нижнем углу, 3 сверху по центру, 4 в центре, 5 снизу по центру, 6 в правом верхнем углу, 7 справа по центру, 8 в правом нижнем углу
                strStyle.Append("background-position:")
                Select Case bkPicPos
                    Case 0 '0 в левом верхнем углу
                        strStyle.Append("left top;")
                    Case 1 '1 слева по центру
                        strStyle.Append("left center;")
                    Case 2 '2 в левом нижнем углу
                        strStyle.Append("left bottom;")
                    Case 3 '3 сверху по центру
                        strStyle.Append("center top;")
                    Case 4 '4 в центре
                        strStyle.Append("center center;")
                    Case 5 '5 снизу по центру
                        strStyle.Append("center bottom;")
                    Case 6 '6 в правом верхнем углу
                        strStyle.Append("right top;")
                    Case 7 '7 справа по центру
                        strStyle.Append("right center;")
                    Case 8 '8 в правом нижнем углу
                        strStyle.Append("right bottom;")
                End Select
            End If
        End If
        mapConvas.Style = strStyle.ToString


        '''''''Создание карты
        'размеры карты
        Dim cellsByX As Integer = 10, cellsByY As Integer = 10 'GetPropertyValueAsInteger(classMap, "CellsByX", mapId, -1, 10)
        If ReadPropertyInt(classMap, "CellsByX", mapId, -1, cellsByX, arrParams) = False Then Return "#Error"
        If ReadPropertyInt(classMap, "CellsByY", mapId, -1, cellsByY, arrParams) = False Then Return "#Error"

        'Пишем заголовок
        Dim hEl As HtmlElement = Nothing
        Dim mCaption As String = ""
        If ReadProperty(classMap, "Caption", mapId, -1, mCaption, arrParams) = False Then Return "#Error"
        mCaption = UnWrapString(mCaption)

        Dim hClose As HtmlElement = hDoc.CreateElement("IMG")
        hClose.SetAttribute("src", "img/Default/close.png")
        hClose.Id = "MapClose"
        AddHandler hClose.Click, Sub(sender As Object, e As HtmlElementEventArgs)
                                     ClearPreviousBitmap()
                                     sender.SetAttribute("src", "img/Default/close.png")

                                     isNewWindow = False
                                     slot = 0
                                     If ReadPropertyInt(classMap, "Slot", -1, -1, slot, Nothing) = False Then Return
                                     If slot = 5 Then isNewWindow = True 'в отдельном окне

                                     If isNewWindow Then
                                         frmMap.Hide()
                                     Else
                                         questEnvironment.wbMap.Hide()
                                     End If
                                 End Sub
        AddHandler hClose.MouseOver, Sub(sender As Object, e As HtmlElementEventArgs)
                                         sender.SetAttribute("src", "img/Default/closeHover.png")
                                     End Sub
        AddHandler hClose.MouseLeave, Sub(sender As Object, e As HtmlElementEventArgs)
                                          sender.SetAttribute("src", "img/Default/close.png")
                                      End Sub

        If mCaption.Length > 0 Then
            hEl = hDoc.CreateElement("H1")
            hEl.Id = "MapCaption"
            hEl.InnerHtml = mCaption
            mapConvas.AppendChild(hEl)
        Else
            hEl = hDoc.CreateElement("DIV")
            mapConvas.AppendChild(hEl)
        End If
        hEl.AppendChild(hClose)

        'создаем элемент ЦЕНТР, внутри которого создаем табдицу - карту
        Dim hCenter As HtmlElement = hDoc.CreateElement("CENTER")
        mapConvas.AppendChild(hCenter)

        Dim hTable As HtmlElement = hDoc.CreateElement("TABLE")
        hTable.Id = "MapMain"
        hTable.SetAttribute("cellpadding", "0")
        hTable.SetAttribute("cellspacing", "0")

        hCenter.AppendChild(hTable)
        Dim hTBody As HtmlElement = hDoc.CreateElement("TBODY")
        hTable.AppendChild(hTBody)

        'Картинка карты MapPicture
        Dim mPicture As String = ""
        If ReadProperty(classMap, "MapPicture", mapId, -1, mPicture, arrParams) = False Then Return "#Error"
        mPicture = UnWrapString(mPicture)
        If mPicture.Length > 0 Then
            strStyle.Clear()
            strStyle.AppendFormat("background-image: url({0});", "'" + mPicture.Replace("\", "/") + "'")
            Dim MapPicStyle As Integer = 0
            If ReadPropertyInt(classMap, "MapPicStyle", mapId, -1, MapPicStyle, arrParams) = False Then Return "#Error"
            '0 простая загрузка, 1 - заполнить, 2 - масштабировать, 3 - размножить, 4 - размножить по Х, 5 - размножить по Y 
            strStyle.Append("background-repeat:")
            Select Case MapPicStyle
                Case 0 '0 простая загрузка
                    strStyle.Append("no-repeat;")
                Case 1 '1 растянуть пропорционально
                    strStyle.Append("no-repeat;background-size:cover;")
                Case 2 '2 заполнить
                    strStyle.Append("no-repeat;background-size:contain;")
                Case 3 '3 масштабировать
                    strStyle.Append("repeat;")
                Case 4 '4 размножить по Х
                    strStyle.Append("repeat-x;")
                Case 5 '5 размножить по Y
                    strStyle.Append("repeat-y;")
            End Select
            hTable.Style = strStyle.ToString
        End If

        Dim cellClassEmpty As String = "", cellClassFog As String = "", fogStyle As Integer = 0
        If ReadPropertyInt(classMap, "FogStyle", mapId, -1, fogStyle, arrParams) = False Then Return "#Error"
        If ReadProperty(classMap, "CellEmptyClass", mapId, -1, cellClassEmpty, arrParams) = False Then Return "#Error"
        If ReadProperty(classMap, "CellInFogClass", mapId, -1, cellClassFog, arrParams) = False Then Return "#Error"
        cellClassEmpty = UnWrapString(cellClassEmpty)
        cellClassFog = UnWrapString(cellClassFog)
        'cellClassCurrent = UnWrapString(cellClassCurrent)


        'Получаем Id текущей клетки (если такая есть)
        Dim tr As HtmlElement, td As HtmlElement
        For y As Integer = 1 To cellsByY
            'строим горизонтальные ряды
            tr = hDoc.CreateElement("TR")
            hTBody.AppendChild(tr)
            tr.Style = "height:" + GetRowHeight(mapId, y)
            For x As Integer = 1 To cellsByX
                'строим вертикальные колонки
                Dim cellId As Integer = GetCellIdByXY(mapId, x, y) 'Id клетки по текущим координатам или -1, если там пусто
                td = hDoc.CreateElement("TD")
                td.Style = "width:" + GetColumnWidth(mapId, x)
                td.SetAttribute("x", x.ToString)
                td.SetAttribute("y", y.ToString)
                td.SetAttribute("mapId", mapId.ToString)
                td.SetAttribute("cellId", cellId.ToString)
                If cellId = -1 Then
                    'По данным координатам клетки не существует
                    If IsEmptyCellInFog(x, y, arrParams) Then
                        td.SetAttribute("ClassName", cellClassFog)
                    Else
                        td.SetAttribute("ClassName", cellClassEmpty)
                    End If
                Else
                    'Клетка существует

                    'в тумане?
                    Dim fog As Boolean = False, showStyle As Integer = 0
                    If ReadPropertyInt(classMap, "ShowStyle", mapId, cellId, showStyle, arrParams) = False Then Return "#Error"
                    If showStyle = 1 Then
                        'клетка видна всегда
                        fog = False
                    ElseIf showStyle = 2 Then
                        'клетка всегда спрятана
                        fog = True
                    Else
                        'клетка не видна пока в тумане
                        If fogStyle > 0 Then
                            If ReadPropertyBool(classMap, "Fog", mapId, cellId, fog, arrParams) = False Then Return "#Error"
                        End If
                    End If

                    If fog Then
                        td.SetAttribute("ClassName", cellClassFog)
                    Else
                        'вставляем в нее Content
                        Dim cContent As String = ""
                        If ReadProperty(classMap, "Content", mapId, cellId, cContent, arrParams) = False Then Return "#Error"
                        cContent = UnWrapString(cContent)
                        If cContent.Length > 0 Then td.InnerHtml = cContent

                        Dim cellClassUsual As String = ""
                        If ReadProperty(classMap, "CellUsualClass", mapId, cellId, cellClassUsual, arrParams) = False Then Return "#Error"
                        cellClassUsual = UnWrapString(cellClassUsual)
                        td.SetAttribute("ClassName", cellClassUsual)
                    End If

                    Dim cLoc As String = UnWrapString(mScript.mainClass(classMap).ChildProperties(mapId)("Location").ThirdLevelProperties(cellId))
                    If String.IsNullOrEmpty(cLoc) Then td.Style = "cursor:default"
                End If
                tr.AppendChild(td)
            Next x
        Next y

        'расположение карты
        Dim mapOffsetX As Integer = 0, mapOffsetY As Integer = 0
        If ReadPropertyInt(classMap, "MapOffsetX", mapId, -1, mapOffsetX, arrParams) = False Then Return "#Error"
        If ReadPropertyInt(classMap, "MapOffsetY", mapId, -1, mapOffsetY, arrParams) = False Then Return "#Error"
        Dim mStyle As String = hTable.Style
        strStyle.Clear()
        strStyle.Append(mStyle)
        Dim mapPositionWasSet As Boolean = False
        If mapOffsetX <> 0 Then
            strStyle.Append("left:" & mapOffsetX.ToString & "px;")
            mapPositionWasSet = True
        End If
        If mapOffsetY <> 0 Then
            strStyle.Append("top:" & mapOffsetY.ToString & "px;")
            mapPositionWasSet = True
        End If
        If mapPositionWasSet Then
            strStyle.Append("position:absolute")
            hTable.Style = strStyle.ToString
        End If

        'создаем div для вывода подписей
        Dim hCaption As HtmlElement = hDoc.CreateElement("DIV")
        Dim infoClass As String = ""
        ReadProperty(classMap, "CellInfoClass", mapId, -1, infoClass, arrParams)
        infoClass = UnWrapString(infoClass)

        hCaption.Id = "CellInfo"
        hCaption.SetAttribute("ClassName", infoClass)
        hCaption.InnerHtml = "&nbsp;"
        mapConvas.AppendChild(hCaption)
        If mapPositionWasSet Then
            strStyle.Clear()
            Dim posTable As Point = GetHTMLelementCoordinates(hTable)
            Dim sizeTable As Size = hTable.OffsetRectangle.Size
            hCaption.Style = "position:absolute;left:" & posTable.X.ToString & "px;top:" & (posTable.Y + sizeTable.Height).ToString & "px;width:" & sizeTable.Width.ToString & "px;"

            Dim posCaption As Point = GetHTMLelementCoordinates(hCaption)
            Dim convasPos As Point = GetHTMLelementCoordinates(mapConvas)
            Dim minHeight As Integer = posCaption.Y - convasPos.Y + hCaption.OffsetRectangle.Height + 10
            mStyle = mapConvas.Style
            If String.IsNullOrEmpty(mapConvas.Style) = False AndAlso mapConvas.Style.EndsWith(";") = False Then mStyle &= ";"
            mStyle &= "min-height:" & minHeight & "px;"
            mapConvas.Style = mStyle
        Else
            'hCaption.Style = "display:none"
        End If

        'устанавливаем изображения клеток
        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMap).ChildProperties(mapId)("CellPicture")
        If IsNothing(ch.ThirdLevelProperties) = False Then
            For cellId As Integer = ch.ThirdLevelProperties.Count - 1 To 0 Step -1
                'получаем картинку
                Dim imgPath As String = ""
                Dim wasFound As Boolean = False

                If ReadProperty(classMap, "CellPicture", mapId, cellId, imgPath, arrParams) Then
                    imgPath = UnWrapString(imgPath)
                    If String.IsNullOrEmpty(imgPath) = False Then
                        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, imgPath)
                        If FileIO.FileSystem.FileExists(fPath) Then wasFound = True
                    End If
                End If

                Dim hCell As HtmlElement = GetCellHtmlElement(cellId, hDoc)
                If IsNothing(hCell) Then Continue For

                If Not wasFound Then
                    AddHandler hCell.MouseOver, AddressOf del_CellMouseOver
                    AddHandler hCell.MouseLeave, AddressOf del_CellMouseLeave
                    AddHandler hCell.Click, AddressOf del_CellClick
                    Continue For
                End If

                'создаем элемент картинки и задаем его стиль и класс
                Dim hImg As HtmlElement = hDoc.CreateElement("IMG")
                hImg.SetAttribute("ClassName", "CellImage")
                hImg.SetAttribute("src", imgPath)
                hImg.SetAttribute("mapId", mapId.ToString)
                hImg.SetAttribute("cellId", cellId.ToString)

                Dim pWidth As String = "", pHeight As String = ""
                If ReadProperty(classMap, "CellPictureWidth", mapId, cellId, pWidth, arrParams) = False Then Continue For
                If ReadProperty(classMap, "CellPictureHeight", mapId, cellId, pHeight, arrParams) = False Then Continue For
                pWidth = UnWrapString(pWidth)
                pHeight = UnWrapString(pHeight)
                If IsNumeric(pWidth) Then pWidth &= "px"
                If IsNumeric(pHeight) Then pHeight &= "px"

                strStyle.Clear()
                If String.IsNullOrEmpty(pWidth) = False Then strStyle.Append("width:" & pWidth & ";")
                If String.IsNullOrEmpty(pHeight) = False Then strStyle.Append("height:" & pHeight & ";")

                hDoc.Body.AppendChild(hImg)

                'устанавливаем положение картинки
                Dim offsetX As Integer = 0, offsetY As Integer = 0
                If ReadPropertyInt(classMap, "CellPictureOffsetX", mapId, cellId, offsetX, arrParams) = False Then Continue For
                If ReadPropertyInt(classMap, "CellPictureOffsetY", mapId, cellId, offsetY, arrParams) = False Then Continue For
                Dim locCell As Point = GetHTMLelementCoordinates(hCell)
                offsetX = offsetX + locCell.X
                offsetY = offsetY + locCell.Y

                strStyle.Append("left:" & offsetX.ToString & "px;top:" & offsetY.ToString & "px;")
                If IsCellVisible(mapId, cellId) = False Then strStyle.Append("display: none;")
                hImg.Style = strStyle.ToString

                AddHandler hImg.MouseMove, AddressOf del_CellImgMouseMove
                AddHandler hImg.MouseLeave, AddressOf del_CellImgMouseLeave
                AddHandler hImg.Click, AddressOf del_CellClick
                AddHandler hImg.MouseDown, AddressOf del_CellImgMouseDown
            Next cellId
        End If

        ClearNearbyFog(arrParams)
        ShowArrowInCurrentCell()
        Return ""
    End Function

    ''' <summary>
    ''' Возвращает направление движения по карте
    ''' </summary>
    Public Function GetDirection(Optional cellFrom As Integer = -1, Optional cellTo As Integer = -1) As DirectionEnum
        If GVARS.G_CURMAP < 0 Then Return DirectionEnum.FORWARD
        If cellFrom < 0 Then cellFrom = GVARS.G_PREVMAPCELL
        If cellTo < 0 Then cellTo = GVARS.G_CURMAPCELL
        If cellTo < 0 Then Return DirectionEnum.FORWARD

        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim defDirection As DirectionEnum = DirectionEnum.FORWARD
        If cellFrom < 0 Then
            ReadPropertyInt(classId, "DefaultDirection", GVARS.G_CURMAP, -1, defDirection, Nothing)
            Return defDirection
        End If

        Dim x As Integer = 0, y As Integer = 0
        ReadPropertyInt(classId, "CellX", GVARS.G_CURMAP, cellTo, x, Nothing)
        ReadPropertyInt(classId, "CellY", GVARS.G_CURMAP, cellTo, y, Nothing)
        Dim locCur As New Point(x, y)

        ReadPropertyInt(classId, "CellX", GVARS.G_CURMAP, cellFrom, x, Nothing)
        ReadPropertyInt(classId, "CellY", GVARS.G_CURMAP, cellFrom, y, Nothing)
        Dim locPrev As New Point(x, y)

        Dim dirs As New List(Of DirectionEnum)

        If locCur.X > locPrev.X Then
            dirs.Add(DirectionEnum.RIGHT)
        ElseIf locCur.X < locPrev.X Then
            dirs.Add(DirectionEnum.LEFT)
        End If

        If locCur.Y > locPrev.Y Then
            dirs.Add(DirectionEnum.BACKWARD)
        ElseIf locCur.Y < locPrev.Y Then
            dirs.Add(DirectionEnum.FORWARD)
        End If

        If dirs.Count = 0 Then
            ReadPropertyInt(classId, "DefaultDirection", GVARS.G_CURMAP, cellTo, defDirection, Nothing)
            Return defDirection
        ElseIf dirs.Count = 1 Then
            Return dirs(0)
        Else
            '2 направления
            If dirs(0) = DirectionEnum.RIGHT Then
                If GetCellIdByXY(GVARS.G_CURMAP, locCur.X - 1, locCur.Y) > 0 Then
                    Return dirs(0)
                Else
                    Return dirs(1)
                End If
            Else
                If GetCellIdByXY(GVARS.G_CURMAP, locCur.X + 1, locCur.Y) > 0 Then
                    Return dirs(0)
                Else
                    Return dirs(1)
                End If
            End If
        End If
    End Function

    Private Function MoveArrowToNewPlace(ByRef hDoc As HtmlDocument, ByRef hCur As HtmlElement, hArrow As HtmlElement, ByRef newCellSize As Size, ByRef newCellLoc As Point, _
                                         ByVal transDuration As Integer, ByRef arrPath() As Integer) As Boolean

        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim pathPoints As New List(Of Point), pathDirections As New List(Of DirectionEnum)
        Dim hCell As HtmlElement = GetCellHtmlElement(arrPath(0), hDoc)
        If IsNothing(hCell) Then
            _ERROR("Сбой при анимации перемещения позиции на карте. Клетка не найдена.", "")
            Return False
        End If
        Dim pt As Point = GetHTMLelementCoordinates(hCell)
        pathPoints.Add(pt)

        'создаем массивы с контрольными точками перемещения и с направлениями
        For i As Integer = 1 To arrPath.Count - 1
            Dim dir As DirectionEnum = GetDirection(arrPath(i - 1), arrPath(i))
            If pathDirections.Count = 0 OrElse dir <> pathDirections.Last Then
                'новое направление
                hCell = GetCellHtmlElement(arrPath(i), hDoc)
                If IsNothing(hCell) Then
                    _ERROR("Сбой при анимации перемещения позиции на карте. Клетка не найдена.", "")
                    Return False
                End If
                pt = GetHTMLelementCoordinates(hCell)
                pathPoints.Add(pt)
                pathDirections.Add(dir)
            Else
                'направление не изменилось
                hCell = GetCellHtmlElement(arrPath(i), hDoc)
                If IsNothing(hCell) Then
                    _ERROR("Сбой при анимации перемещения позиции на карте. Клетка не найдена.", "")
                    Return False
                End If
                pt = GetHTMLelementCoordinates(hCell)
                pathPoints(pathPoints.Count - 1) = pt
            End If
        Next

        Dim sWatch As New Stopwatch
        For i As Integer = 1 To pathPoints.Count - 1

            'меняем стиль, указав новое расположение
            Dim strStyle As String = hCur.Style
            Dim pos As Integer = strStyle.IndexOf("left:")
            If pos = -1 Then
                strStyle &= "left:" & pathPoints(i).X.ToString & "px;"
            Else
                Dim pos2 As Integer = strStyle.IndexOf(";"c, pos + 1)
                If pos2 = -1 Then
                    Mid(strStyle, pos + 1) = "left:" & pathPoints(i).X.ToString & "px;"
                Else
                    Mid(strStyle, pos + 1, pos2 - pos + 2) = "left:" & pathPoints(i).X.ToString & "px;"
                End If
            End If

            pos = strStyle.IndexOf("top:")
            If pos = -1 Then
                strStyle &= "top:" & pathPoints(i).Y.ToString & "px;"
            Else
                Dim pos2 As Integer = strStyle.IndexOf(";"c, pos + 1)
                If pos2 = -1 Then
                    Mid(strStyle, pos + 1) = "top:" & pathPoints(i).Y.ToString & "px;"
                Else
                    Mid(strStyle, pos + 1, pos2 - pos + 2) = "top:" & pathPoints(i).Y.ToString & "px;"
                End If
            End If
            hCur.Style = strStyle

            If IsNothing(hArrow) = False Then
                Dim picPath As String = ""
                Select Case pathDirections(i - 1)
                    Case DirectionEnum.RIGHT
                        ReadProperty(classId, "ArrowRight", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                    Case DirectionEnum.LEFT
                        ReadProperty(classId, "ArrowLeft", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                    Case DirectionEnum.BACKWARD
                        ReadProperty(classId, "ArrowBack", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                    Case Else
                        ReadProperty(classId, "ArrowFront", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                End Select
                picPath = UnWrapString(picPath)
                If String.IsNullOrEmpty(picPath) Then
                    hArrow = Nothing
                Else
                    hArrow.SetAttribute("src", picPath)
                End If
            End If

            sWatch.Start()
            Do While sWatch.ElapsedMilliseconds < transDuration
                Application.DoEvents()
            Loop
            sWatch.Stop()
            sWatch.Reset()
        Next

        Return True
    End Function

    ''' <summary>
    ''' Отображает стрелку в текущей клетке, позывающую направление движения, и убирает стрелку из предыдущего места. Также отображает текущую клетку как активную и меняет 
    ''' картинку текущей клетки CellPicture на CellPictureCurrent
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ShowArrowInCurrentCell(Optional ByRef arrPath() As Integer = Nothing)
        If GVARS.G_CURMAP < 0 OrElse GVARS.G_CURMAPCELL < 0 Then Return
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim arrParams() As String = {GVARS.G_CURMAP.ToString, GVARS.G_PREVMAPCELL.ToString}
        Dim transDuration As Integer = 0
        'получаем прродолжительность эффекта перехода
        If IsNothing(arrPath) = False AndAlso GVARS.G_PREVMAPCELL > -1 Then
            ReadPropertyInt(classId, "TransitionDuration", GVARS.G_CURMAP, -1, transDuration, arrParams)
            If transDuration <= 0 Then
                transDuration = 0
            End If
        End If

        Dim hDoc As HtmlDocument = questEnvironment.wbMap.Document
        If IsNothing(hDoc) Then Return

        Dim picPath As String = "", fPath As String = ""
        'меняем картинку предыдущей клетки на CellPicture
        If GVARS.G_PREVMAPCELL > -1 Then
            ReadProperty(classId, "CellPicture", GVARS.G_CURMAP, GVARS.G_PREVMAPCELL, picPath, arrParams)
            picPath = UnWrapString(picPath)
            If String.IsNullOrEmpty(picPath) = False Then
                fPath = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picPath)
                If FileIO.FileSystem.FileExists(fPath) Then
                    Dim hImg As HtmlElement = GetCellImageHTMLElement(GVARS.G_PREVMAPCELL, hDoc)
                    If IsNothing(hImg) = False Then
                        hImg.SetAttribute("src", picPath)
                    End If
                End If
            End If
        End If

        'меняем картинку текущей клетки на CellPictureCurrent
        Dim isCellImg As Boolean = False
        ReadProperty(classId, "CellPicture", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, arrParams)
        picPath = UnWrapString(picPath)
        If String.IsNullOrEmpty(picPath) = False Then
            isCellImg = True
            ReadProperty(classId, "CellPictureCurrent", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, arrParams)
            picPath = UnWrapString(picPath)
            If String.IsNullOrEmpty(picPath) = False Then
                fPath = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picPath)
                If FileIO.FileSystem.FileExists(fPath) Then
                    Dim hImg As HtmlElement = GetCellImageHTMLElement(GVARS.G_CURMAPCELL, hDoc)
                    If IsNothing(hImg) = False Then
                        hImg.SetAttribute("src", picPath)
                    End If
                End If
            End If
        End If

        Dim direct As DirectionEnum = GetDirection()

        Dim hCell As HtmlElement = GetCellHtmlElement(GVARS.G_CURMAPCELL, hDoc)
        If IsNothing(hCell) Then Return
        Dim locCell As Point = GetHTMLelementCoordinates(hCell)
        Dim sizeCell As Size = hCell.OffsetRectangle.Size

        Dim useArrows As Boolean = True
        ReadPropertyBool(classId, "UseArrows", GVARS.G_CURMAP, -1, useArrows, Nothing)

        'получаем стрелку и текущее выделение с предыдущего места
        Dim hArrow As HtmlElement = hDoc.GetElementById("Arrow")
        Dim hCur As HtmlElement = hDoc.GetElementById("Current")
        Dim arrowNew As Boolean = False, curNew As Boolean = False

        'Получение / создание контура
        If IsNothing(hCur) = False Then
            'hCur уже существует
            If transDuration = 0 Then
                'перемещение мгновенное
                If isCellImg Then
                    Dim msCur As mshtml.IHTMLDOMNode = hCur.DomElement
                    msCur.removeNode(True)
                    hCur = Nothing
                Else
                    hCur.Style = "position:absolute;left:" & locCell.X.ToString & "px;top:" & locCell.Y.ToString & "px;width:" & sizeCell.Width.ToString & "px;height:" & sizeCell.Height.ToString & _
                        "px;overflow:hidden;transition-property: all;transition-duration: " & transDuration.ToString & "ms;"
                End If
            Else
                'перемещение надо строить
                HTMLAddCSSstyle(hCur, "transition-duration", transDuration.ToString & "ms")
            End If
        ElseIf (isCellImg AndAlso transDuration = 0) = False Then
            'hCur еще не существует
            'создание выделения текущей клетки
            curNew = True
            hCur = hDoc.CreateElement("DIV")
            hCur.Id = "Current"

            Dim curClassName As String = ""
            ReadProperty(classId, "CellCurrentClass", GVARS.G_CURMAP, -1, curClassName, {GVARS.G_CURMAP.ToString, GVARS.G_CURMAPCELL.ToString})
            curClassName = UnWrapString(curClassName)
            hCur.SetAttribute("ClassName", curClassName)

            If transDuration = 0 Then
                'перемещение мгновенное - создаем сразу в новой клетке
                hCur.Style = "position:absolute;left:" & locCell.X.ToString & "px;top:" & locCell.Y.ToString & "px;width:" & sizeCell.Width.ToString & "px;height:" & sizeCell.Height.ToString & _
                    "px;overflow:hidden;transition-property: all;transition-duration: " & transDuration.ToString & "ms;"
            Else
                'перемещение надо строить
                'Создаем сначала в старом месте
                Dim hPrevCell As HtmlElement = GetCellHtmlElement(GVARS.G_PREVMAPCELL, hDoc)
                If IsNothing(hPrevCell) Then
                    'не нашли предыдущую клетку - мгновенное перемещение
                    transDuration = 0
                    hCur.Style = "position:absolute;left:" & locCell.X.ToString & "px;top:" & locCell.Y.ToString & "px;width:" & sizeCell.Width.ToString & "px;height:" & sizeCell.Height.ToString & _
                        "px;overflow:hidden;transition-property: all;transition-duration: " & transDuration.ToString & "ms;"
                Else
                    Dim locPrevCell As Point = GetHTMLelementCoordinates(hPrevCell)
                    Dim sizePrevCell As Size = hPrevCell.OffsetRectangle.Size
                    hCur.Style = "position:absolute;left:" & locPrevCell.X.ToString & "px;top:" & locPrevCell.Y.ToString & "px;width:" & sizePrevCell.Width.ToString & "px;height:" & sizePrevCell.Height.ToString & _
                        "px;overflow:hidden;transition-property: all;transition-duration: " & transDuration.ToString & "ms;"
                End If
            End If

            hDoc.Body.AppendChild(hCur)
        End If

        'Получение / создание стрелки
        If IsNothing(hArrow) = False Then
            'hArrow уже существует
            If useArrows = False Then
                'стрелки не используются
                Dim maArrow As mshtml.IHTMLDOMNode = hArrow.DomElement
                maArrow.removeNode(True)
                hArrow = Nothing
            ElseIf transDuration = 0 Then
                'перемещение мгновенное
                If isCellImg Then
                    Dim maArrow As mshtml.IHTMLDOMNode = hArrow.DomElement
                    maArrow.removeNode(True)
                    hArrow = Nothing
                Else
                    Dim sizeArrow As Size = hArrow.OffsetRectangle.Size
                    Dim locArrow As New Point
                    locArrow.X = (sizeCell.Width - sizeArrow.Width) / 2
                    locArrow.Y = (sizeCell.Height - sizeArrow.Height) / 2
                    hArrow.Style = "position:absolute;left:" & locArrow.X.ToString & "px;top:" & locArrow.Y.ToString & "px"

                    Select Case direct
                        Case DirectionEnum.RIGHT
                            ReadProperty(classId, "ArrowRight", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                        Case DirectionEnum.LEFT
                            ReadProperty(classId, "ArrowLeft", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                        Case DirectionEnum.BACKWARD
                            ReadProperty(classId, "ArrowBack", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                        Case Else
                            ReadProperty(classId, "ArrowFront", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                    End Select
                    picPath = UnWrapString(picPath)
                    If String.IsNullOrEmpty(picPath) Then
                        hArrow = Nothing
                    Else
                        hArrow.SetAttribute("src", picPath)
                    End If
                End If
            Else
                ''перемещение плавное
                'Select Case direct
                '    Case DirectionEnum.RIGHT
                '        ReadProperty(classId, "ArrowRight", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                '    Case DirectionEnum.LEFT
                '        ReadProperty(classId, "ArrowLeft", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                '    Case DirectionEnum.BACKWARD
                '        ReadProperty(classId, "ArrowBack", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                '    Case Else
                '        ReadProperty(classId, "ArrowFront", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                'End Select
                'picPath = UnWrapString(picPath)
                'If String.IsNullOrEmpty(picPath) Then
                '    hArrow = Nothing
                'Else
                '    hArrow.SetAttribute("src", picPath)
                'End If
            End If
        ElseIf (isCellImg AndAlso transDuration = 0) = False AndAlso useArrows Then
            'hArrow еще не существует
            'создание стрелки

            'стрелка с направлением
            arrowNew = True
            hArrow = hDoc.CreateElement("IMG")
            hArrow.Id = "Arrow"
            hCur.AppendChild(hArrow)

            Select Case direct
                Case DirectionEnum.RIGHT
                    ReadProperty(classId, "ArrowRight", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                Case DirectionEnum.LEFT
                    ReadProperty(classId, "ArrowLeft", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                Case DirectionEnum.BACKWARD
                    ReadProperty(classId, "ArrowBack", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
                Case Else
                    ReadProperty(classId, "ArrowFront", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, picPath, Nothing)
            End Select
            picPath = UnWrapString(picPath)
            If String.IsNullOrEmpty(picPath) Then
                hArrow = Nothing
            Else
                hArrow.SetAttribute("src", picPath)
            End If


            If transDuration = 0 Then
                'перемещение мгновенное - создаем сразу в новой клетке
                Dim sizeArrow As Size = hArrow.OffsetRectangle.Size
                Dim locArrow As New Point
                locArrow.X = (sizeCell.Width - sizeArrow.Width) / 2
                locArrow.Y = (sizeCell.Height - sizeArrow.Height) / 2
                hArrow.Style = "position:absolute;left:" & locArrow.X.ToString & "px;top:" & locArrow.Y.ToString & "px"
            Else
                'перемещение надо строить
                'Создаем сначала в старом месте
                Dim hPrevCell As HtmlElement = GetCellHtmlElement(GVARS.G_PREVMAPCELL, hDoc)
                If IsNothing(hPrevCell) Then
                    'не нашли предыдущую клетку - мгновенное перемещение
                    transDuration = 0
                    Dim sizeArrow As Size = hArrow.OffsetRectangle.Size
                    Dim locArrow As New Point
                    locArrow.X = (sizeCell.Width - sizeArrow.Width) / 2
                    locArrow.Y = (sizeCell.Height - sizeArrow.Height) / 2
                    hArrow.Style = "position:absolute;left:" & locArrow.X.ToString & "px;top:" & locArrow.Y.ToString & "px"
                Else
                    Dim locPrevCell As Point = GetHTMLelementCoordinates(hPrevCell)
                    Dim sizePrevCell As Size = hPrevCell.OffsetRectangle.Size
                    Dim sizeArrow As Size = hArrow.OffsetRectangle.Size
                    Dim locArrow As New Point
                    locArrow.X = (sizePrevCell.Width - sizeArrow.Width) / 2
                    locArrow.Y = (sizePrevCell.Height - sizeArrow.Height) / 2

                    hArrow.Style = "position:absolute;left:" & locArrow.X.ToString & "px;top:" & locArrow.Y.ToString & "px"
                End If
            End If

        End If


        If isCellImg AndAlso transDuration = 0 Then Return 'если клетка-картинка и перемещение мгновенное, то стрелку и контур не отображаем


        'If useArrows Then
        '    'стрелка с направлением
        '    Dim sizeArrow As Size = hArrow.OffsetRectangle.Size
        '    Dim locArrow As New Point
        '    locArrow.X = (sizeCell.Width - sizeArrow.Width) / 2
        '    locArrow.Y = (sizeCell.Height - sizeArrow.Height) / 2
        '    hArrow.Style = "position:absolute;left:" & locArrow.X.ToString & "px;top:" & locArrow.Y.ToString & "px;"


        'End If


        'визаульный эффект перемещения контура и стрелки
        If transDuration > 0 Then
            MoveArrowToNewPlace(hDoc, hCur, hArrow, sizeCell, locCell, transDuration, arrPath)

            If isCellImg Then
                'если пункт назначения - картинка, то убираем стрелку и выделение
                If IsNothing(hCur) = False Then
                    Dim msCur As mshtml.IHTMLDOMNode = hCur.DomElement
                    msCur.removeNode(True)
                    hCur = Nothing
                End If

                If IsNothing(hArrow) = False Then
                    Dim msArrow As mshtml.IHTMLDOMNode = hArrow.DomElement
                    msArrow.removeNode(True)
                    hArrow = Nothing
                End If
            End If
        End If

        'события стрелки
        If arrowNew AndAlso IsNothing(hArrow) = False Then
            AddHandler hArrow.MouseOver, Sub(sender As Object, e As HtmlElementEventArgs)
                                             Dim El As HtmlElement = GetCellHtmlElement(GVARS.G_CURMAPCELL, hDoc)
                                             If IsNothing(El) = False Then Call del_CellMouseOver(El, e)
                                         End Sub

            AddHandler hArrow.MouseLeave, Sub(sender As Object, e As HtmlElementEventArgs)
                                              Dim El As HtmlElement = GetCellHtmlElement(GVARS.G_CURMAPCELL, hDoc)
                                              If IsNothing(El) = False Then Call del_CellMouseLeave(El, e)
                                          End Sub
        End If

        'события контура
        If curNew AndAlso IsNothing(hCur) = False Then
            AddHandler hCur.MouseOver, Sub(sender As Object, e As HtmlElementEventArgs)
                                           Dim El As HtmlElement = GetCellHtmlElement(GVARS.G_CURMAPCELL, hDoc)
                                           If IsNothing(El) = False Then Call del_CellMouseOver(El, e)
                                       End Sub

            AddHandler hCur.MouseLeave, Sub(sender As Object, e As HtmlElementEventArgs)
                                            Dim El As HtmlElement = GetCellHtmlElement(GVARS.G_CURMAPCELL, hDoc)
                                            If IsNothing(El) = False Then Call del_CellMouseLeave(El, e)
                                        End Sub
        End If
    End Sub

    ''' <summary>
    ''' Строит кратчайший путь на карте и возвращает массив с Id клеток, входящих в путь, или массив с -1, если путь не существует
    ''' </summary>
    ''' <param name="cellStart">Id начальной клетки</param>
    ''' <param name="cellFinish">Id конечной клетки</param>
    ''' <param name="useDiagonals">Использовать диагонали</param>
    ''' <param name="buildStyle">Стиль построения - из каких клеток строить</param>
    Public Function BuildPath(ByVal cellStart As Integer, ByVal cellFinish As Integer, ByVal useDiagonals As Boolean, ByVal buildStyle As BuildStyleEnum, ByVal doEvents As Boolean) As String()
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim CellsByX As Integer = 0, CellsByY As Integer = 0, xStart As Integer = 0, yStart As Integer = 0, xFinal As Integer = 0, yFinal As Integer = 0
        Dim curMapStr As String = GVARS.G_CURMAP.ToString, cellStartStr As String = cellStart.ToString, cellFinishStr As String = cellFinish.ToString
        Dim arrs() As String = {curMapStr, GVARS.G_CURMAPCELL.ToString}
        If cellStart < 0 OrElse cellFinish < 0 Then Return {"-1"}
        If cellStart = cellFinish Then Return {cellStart.ToString}

        If ReadPropertyInt(classId, "CellX", GVARS.G_CURMAP, cellStart, xStart, arrs) = False Then Return {"#Error"}
        If ReadPropertyInt(classId, "CellY", GVARS.G_CURMAP, cellStart, yStart, arrs) = False Then Return {"#Error"}
        If ReadPropertyInt(classId, "CellX", GVARS.G_CURMAP, cellFinish, xFinal, arrs) = False Then Return {"#Error"}
        If ReadPropertyInt(classId, "CellY", GVARS.G_CURMAP, cellFinish, yFinal, arrs) = False Then Return {"#Error"}
        If ReadPropertyInt(classId, "CellsByX", GVARS.G_CURMAP, -1, CellsByX, arrs) = False Then Return {"#Error"}
        If ReadPropertyInt(classId, "CellsByY", GVARS.G_CURMAP, -1, CellsByY, arrs) = False Then Return {"#Error"}

        Dim cellsMap() As Byte  'карта для поиска пути, размером CellsByX * CellsByY (начиная от 0) 
        Dim lstPosToCellId As New SortedList(Of Integer, Integer) 'ключ - положение в cellsMap, значение - соответствующее cellId
        ReDim cellsMap(CellsByX * CellsByY - 1)
        Dim ch As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)

        'наносим на карту cellsMap клетки, доступные для посещения
        Dim locStr As String = "", locId As Integer = -1, visits As Integer = -1
        Dim classL As Integer = mScript.mainClassHash("L")
        For i As Integer = 0 To ch("Name").ThirdLevelProperties.Count - 1
            arrs = {curMapStr, i.ToString}
            If ReadProperty(classId, "Location", GVARS.G_CURMAP, i, locStr, arrs) = False Then Return {"#Error"}
            locId = GetSecondChildIdByName(locStr, mScript.mainClass(classL).ChildProperties)
            If locId < 0 Then Continue For
            Dim addCell As Boolean = False

            If buildStyle = BuildStyleEnum.AllCells Then
                addCell = True
            ElseIf IsCellVisible(GVARS.G_CURMAP, i) Then
                If buildStyle = BuildStyleEnum.OnlyVisited Then
                    If ReadPropertyInt(classId, "Visits", GVARS.G_CURMAP, i, visits, {GVARS.G_CURMAP.ToString, i.ToString}) = False Then Return {"#Error"}
                    If visits > 0 Then addCell = True
                Else
                    addCell = True
                End If
            End If

            If addCell Then
                Dim iX As Integer = 0, iY As Integer = 0
                If ReadProperty(classId, "CellX", GVARS.G_CURMAP, i, iX, arrs) = False Then Return {"#Error"}
                If ReadProperty(classId, "CellY", GVARS.G_CURMAP, i, iY, arrs) = False Then Return {"#Error"}
                'Dim mPos As Integer = (iX - 1) * CellsByX + iY - 1
                Dim mPos As Integer = (iY - 1) * CellsByX + iX
                cellsMap(mPos) = 2 '2 - клетка существует и доступна для посещения 
                lstPosToCellId.Add(mPos, i)
            End If
        Next i
        'стартовая клетка
        'Dim posStart As Integer = (xStart - 1) * CellsByX + yStart - 1
        Dim posStart As Integer = (yStart - 1) * CellsByX + xStart
        cellsMap(posStart) = 1 '1 - клетка уже пройдена
        'финишная клетка
        'Dim posFinish As Integer = (xFinal - 1) * CellsByX + yFinal - 1
        Dim posFinish As Integer = (yFinal - 1) * CellsByX + xFinal
        If cellsMap(posFinish) = 0 Then Return {"-1"} 'конечная клетка не была посещена при buildStyle = BuildStyleEnum.OnlyVisited 
        cellsMap(posFinish) = 3 '3 - финиш


        'карта построена
        'события CellTransitEvent
        Dim eventId1 As Integer = 0, eventId2 As Integer = 0, eventId3 As Integer = 0
        If doEvents Then
            eventId1 = mScript.mainClass(classId).Properties("CellBuildingPathEvent").eventId
            eventId2 = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("CellBuildingPathEvent").eventId
        End If

        ' 0 - клетки нет; 1 - клетка уже пройдена; 2 - клетка существует и доступна для посещения; 3 - финиш
        Dim curX As Integer = xStart, curY As Integer = yStart, pos As Integer = 0, result As String = "", transitCellId As Integer, pbArrs() As String
        Dim lstPath As New List(Of List(Of Integer)) 'пути, содержащие pos в cellsMap
        Dim lstCur As New List(Of Integer)
        lstCur.Add(posStart)
        lstPath.Add(lstCur)
        Do
            Dim lstUB As Integer = lstPath.Count - 1
            For lstId As Integer = lstUB To 0 Step -1
                'перебираем все построенные пути
                lstCur = lstPath(lstId)
                Dim curPos As Integer = lstCur.Last
                'curX = Math.Floor(curPos / CellsByX) + 1
                curY = Math.Floor((curPos - 1) / CellsByX) + 1 ' 'curPos - (curX - 1) * CellsByX + 1
                curX = curPos - (curY - 1) * CellsByX
                Dim directions As Integer = 0

                If useDiagonals Then
                    'диагональные направления
                    pos = BuildPath_PosCellLeftTop(curX, curY, CellsByX) 'left top
                    If pos > -1 AndAlso cellsMap(pos) > 1 Then
                        '''''''''''''''События'''''''''''''''''''''''
                        If doEvents Then
                            transitCellId = lstPosToCellId(pos)
                            pbArrs = {curMapStr, transitCellId.ToString, cellStartStr, cellFinishStr}
                            result = ""
                            'событие CellBuildingPathEvent глобальное
                            If eventId1 > 0 Then
                                result = mScript.eventRouter.RunEvent(eventId1, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If

                            'событие CellBuildingPathEvent карты
                            If eventId2 > 0 AndAlso result <> "False" Then
                                result = mScript.eventRouter.RunEvent(eventId2, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If

                            'событие CellBuildingPathEvent клетки
                            If result <> "False" Then
                                Dim eventId As Integer = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("CellBuildingPathEvent").ThirdLevelEventId(transitCellId)
                                If eventId > 0 Then
                                    result = mScript.eventRouter.RunEvent(eventId, pbArrs, "CellBuildingPathEvent", False)
                                    If result = "#Error" Then Return {"#Error"}
                                End If
                            End If
                        End If

                        If result <> "False" Then
                            If directions > 0 Then
                                'новое ветвление пути
                                Dim lstNew As New List(Of Integer)
                                lstNew.AddRange(lstCur)
                                lstNew(lstNew.Count - 1) = pos
                                lstPath.Add(lstNew)
                            Else
                                'продолжение пути
                                lstCur.Add(pos)
                            End If
                            If cellsMap(pos) = 3 Then
                                If directions > 0 Then lstCur = lstPath.Last
                                Exit Do 'путь построен
                            End If
                            directions += 1
                            cellsMap(pos) = 1 'клетка посещена
                        End If
                    End If

                    pos = BuildPath_PosCellRightTop(curX, curY, CellsByX) 'right top
                    If pos > -1 AndAlso cellsMap(pos) > 1 Then
                        '''''''''''''''События'''''''''''''''''''''''
                        If doEvents Then
                            transitCellId = lstPosToCellId(pos)
                            pbArrs = {curMapStr, transitCellId.ToString, cellStartStr, cellFinishStr}
                            result = ""
                            'событие CellBuildingPathEvent глобальное
                            If eventId1 > 0 Then
                                result = mScript.eventRouter.RunEvent(eventId1, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If

                            'событие CellBuildingPathEvent карты
                            If eventId2 > 0 AndAlso result <> "False" Then
                                result = mScript.eventRouter.RunEvent(eventId2, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If

                            'событие CellBuildingPathEvent клетки
                            If result <> "False" Then
                                Dim eventId As Integer = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("CellBuildingPathEvent").ThirdLevelEventId(transitCellId)
                                If eventId > 0 Then
                                    result = mScript.eventRouter.RunEvent(eventId, pbArrs, "CellBuildingPathEvent", False)
                                    If result = "#Error" Then Return {"#Error"}
                                End If
                            End If
                        End If

                        If result <> "False" Then
                            If directions > 0 Then
                                'новое ветвление пути
                                Dim lstNew As New List(Of Integer)
                                lstNew.AddRange(lstCur)
                                lstNew(lstNew.Count - 1) = pos
                                lstPath.Add(lstNew)
                            Else
                                'продолжение пути
                                lstCur.Add(pos)
                            End If
                            If cellsMap(pos) = 3 Then
                                If directions > 0 Then lstCur = lstPath.Last
                                Exit Do 'путь построен
                            End If
                            directions += 1
                            cellsMap(pos) = 1 'клетка посещена
                        End If
                    End If

                    pos = BuildPath_PosCellLeftBottom(curX, curY, CellsByX, CellsByY) 'left bottom
                    If pos > -1 AndAlso cellsMap(pos) > 1 Then
                        '''''''''''''''События'''''''''''''''''''''''
                        If doEvents Then
                            transitCellId = lstPosToCellId(pos)
                            pbArrs = {curMapStr, transitCellId.ToString, cellStartStr, cellFinishStr}
                            result = ""
                            'событие CellBuildingPathEvent глобальное
                            If eventId1 > 0 Then
                                result = mScript.eventRouter.RunEvent(eventId1, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If

                            'событие CellBuildingPathEvent карты
                            If eventId2 > 0 AndAlso result <> "False" Then
                                result = mScript.eventRouter.RunEvent(eventId2, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If

                            'событие CellBuildingPathEvent клетки
                            If result <> "False" Then
                                Dim eventId As Integer = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("CellBuildingPathEvent").ThirdLevelEventId(transitCellId)
                                If eventId > 0 Then
                                    result = mScript.eventRouter.RunEvent(eventId, pbArrs, "CellBuildingPathEvent", False)
                                    If result = "#Error" Then Return {"#Error"}
                                End If
                            End If
                        End If

                        If result <> "False" Then
                            If directions > 0 Then
                                'новое ветвление пути
                                Dim lstNew As New List(Of Integer)
                                lstNew.AddRange(lstCur)
                                lstNew(lstNew.Count - 1) = pos
                                lstPath.Add(lstNew)
                            Else
                                'продолжение пути
                                lstCur.Add(pos)
                            End If
                            If cellsMap(pos) = 3 Then
                                If directions > 0 Then lstCur = lstPath.Last
                                Exit Do 'путь построен
                            End If
                            directions += 1
                            cellsMap(pos) = 1 'клетка посещена
                        End If
                    End If

                    pos = BuildPath_PosCellRightBottom(curX, curY, CellsByX, CellsByY) 'right bottom
                    If pos > -1 AndAlso cellsMap(pos) > 1 Then
                        '''''''''''''''События'''''''''''''''''''''''
                        If doEvents Then
                            transitCellId = lstPosToCellId(pos)
                            pbArrs = {curMapStr, transitCellId.ToString, cellStartStr, cellFinishStr}
                            result = ""
                            'событие CellBuildingPathEvent глобальное
                            If eventId1 > 0 Then
                                result = mScript.eventRouter.RunEvent(eventId1, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If

                            'событие CellBuildingPathEvent карты
                            If eventId2 > 0 AndAlso result <> "False" Then
                                result = mScript.eventRouter.RunEvent(eventId2, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If

                            'событие CellBuildingPathEvent клетки
                            If result <> "False" Then
                                Dim eventId As Integer = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("CellBuildingPathEvent").ThirdLevelEventId(transitCellId)
                                If eventId > 0 Then
                                    result = mScript.eventRouter.RunEvent(eventId, pbArrs, "CellBuildingPathEvent", False)
                                    If result = "#Error" Then Return {"#Error"}
                                End If
                            End If
                        End If

                        If result <> "False" Then
                            If directions > 0 Then
                                'новое ветвление пути
                                Dim lstNew As New List(Of Integer)
                                lstNew.AddRange(lstCur)
                                lstNew(lstNew.Count - 1) = pos
                                lstPath.Add(lstNew)
                            Else
                                'продолжение пути
                                lstCur.Add(pos)
                            End If
                            If cellsMap(pos) = 3 Then
                                If directions > 0 Then lstCur = lstPath.Last
                                Exit Do 'путь построен
                            End If
                            directions += 1
                            cellsMap(pos) = 1 'клетка посещена
                        End If
                    End If
                End If

                pos = BuildPath_PosCellLeft(curX, curY, CellsByX) 'left
                If pos > -1 AndAlso cellsMap(pos) > 1 Then
                    '''''''''''''''События'''''''''''''''''''''''
                    If doEvents Then
                        transitCellId = lstPosToCellId(pos)
                        pbArrs = {curMapStr, transitCellId.ToString, cellStartStr, cellFinishStr}
                        result = ""
                        'событие CellBuildingPathEvent глобальное
                        If eventId1 > 0 Then
                            result = mScript.eventRouter.RunEvent(eventId1, pbArrs, "CellBuildingPathEvent", False)
                            If result = "#Error" Then Return {"#Error"}
                        End If

                        'событие CellBuildingPathEvent карты
                        If eventId2 > 0 AndAlso result <> "False" Then
                            result = mScript.eventRouter.RunEvent(eventId2, pbArrs, "CellBuildingPathEvent", False)
                            If result = "#Error" Then Return {"#Error"}
                        End If

                        'событие CellBuildingPathEvent клетки
                        If result <> "False" Then
                            Dim eventId As Integer = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("CellBuildingPathEvent").ThirdLevelEventId(transitCellId)
                            If eventId > 0 Then
                                result = mScript.eventRouter.RunEvent(eventId, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If
                        End If
                    End If

                    If result <> "False" Then
                        If directions > 0 Then
                            'новое ветвление пути
                            Dim lstNew As New List(Of Integer)
                            lstNew.AddRange(lstCur)
                            lstNew(lstNew.Count - 1) = pos
                            lstPath.Add(lstNew)
                        Else
                            'продолжение пути
                            lstCur.Add(pos)
                        End If
                        If cellsMap(pos) = 3 Then
                            If directions > 0 Then lstCur = lstPath.Last
                            Exit Do 'путь построен
                        End If
                        directions += 1
                        cellsMap(pos) = 1 'клетка посещена
                    End If
                End If

                pos = BuildPath_PosCellTop(curX, curY, CellsByX) 'top
                If pos > -1 AndAlso cellsMap(pos) > 1 Then
                    '''''''''''''''События'''''''''''''''''''''''
                    If doEvents Then
                        transitCellId = lstPosToCellId(pos)
                        pbArrs = {curMapStr, transitCellId.ToString, cellStartStr, cellFinishStr}
                        result = ""
                        'событие CellBuildingPathEvent глобальное
                        If eventId1 > 0 Then
                            result = mScript.eventRouter.RunEvent(eventId1, pbArrs, "CellBuildingPathEvent", False)
                            If result = "#Error" Then Return {"#Error"}
                        End If

                        'событие CellBuildingPathEvent карты
                        If eventId2 > 0 AndAlso result <> "False" Then
                            result = mScript.eventRouter.RunEvent(eventId2, pbArrs, "CellBuildingPathEvent", False)
                            If result = "#Error" Then Return {"#Error"}
                        End If

                        'событие CellBuildingPathEvent клетки
                        If result <> "False" Then
                            Dim eventId As Integer = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("CellBuildingPathEvent").ThirdLevelEventId(transitCellId)
                            If eventId > 0 Then
                                result = mScript.eventRouter.RunEvent(eventId, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If
                        End If
                    End If

                    If result <> "False" Then
                        If directions > 0 Then
                            'новое ветвление пути
                            Dim lstNew As New List(Of Integer)
                            lstNew.AddRange(lstCur)
                            lstNew(lstNew.Count - 1) = pos
                            lstPath.Add(lstNew)
                        Else
                            'продолжение пути
                            lstCur.Add(pos)
                        End If
                        If cellsMap(pos) = 3 Then
                            If directions > 0 Then lstCur = lstPath.Last
                            Exit Do 'путь построен
                        End If
                        directions += 1
                        cellsMap(pos) = 1 'клетка посещена
                    End If
                End If

                pos = BuildPath_PosCellRight(curX, curY, CellsByX) 'right
                If pos > -1 AndAlso cellsMap(pos) > 1 Then
                    '''''''''''''''События'''''''''''''''''''''''
                    If doEvents Then
                        transitCellId = lstPosToCellId(pos)
                        pbArrs = {curMapStr, transitCellId.ToString, cellStartStr, cellFinishStr}
                        result = ""
                        'событие CellBuildingPathEvent глобальное
                        If eventId1 > 0 Then
                            result = mScript.eventRouter.RunEvent(eventId1, pbArrs, "CellBuildingPathEvent", False)
                            If result = "#Error" Then Return {"#Error"}
                        End If

                        'событие CellBuildingPathEvent карты
                        If eventId2 > 0 AndAlso result <> "False" Then
                            result = mScript.eventRouter.RunEvent(eventId2, pbArrs, "CellBuildingPathEvent", False)
                            If result = "#Error" Then Return {"#Error"}
                        End If

                        'событие CellBuildingPathEvent клетки
                        If result <> "False" Then
                            Dim eventId As Integer = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("CellBuildingPathEvent").ThirdLevelEventId(transitCellId)
                            If eventId > 0 Then
                                result = mScript.eventRouter.RunEvent(eventId, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If
                        End If
                    End If

                    If result <> "False" Then
                        If directions > 0 Then
                            'новое ветвление пути
                            Dim lstNew As New List(Of Integer)
                            lstNew.AddRange(lstCur)
                            lstNew(lstNew.Count - 1) = pos
                            lstPath.Add(lstNew)
                        Else
                            'продолжение пути
                            lstCur.Add(pos)
                        End If
                        If cellsMap(pos) = 3 Then
                            If directions > 0 Then lstCur = lstPath.Last
                            Exit Do 'путь построен
                        End If
                        directions += 1
                        cellsMap(pos) = 1 'клетка посещена
                    End If
                End If

                pos = BuildPath_PosCellBottom(curX, curY, CellsByX, CellsByY) 'bottom
                If pos > -1 AndAlso cellsMap(pos) > 1 Then
                    '''''''''''''''События'''''''''''''''''''''''
                    If doEvents Then
                        transitCellId = lstPosToCellId(pos)
                        pbArrs = {curMapStr, transitCellId.ToString, cellStartStr, cellFinishStr}
                        result = ""
                        'событие CellBuildingPathEvent глобальное
                        If eventId1 > 0 Then
                            result = mScript.eventRouter.RunEvent(eventId1, pbArrs, "CellBuildingPathEvent", False)
                            If result = "#Error" Then Return {"#Error"}
                        End If

                        'событие CellBuildingPathEvent карты
                        If eventId2 > 0 AndAlso result <> "False" Then
                            result = mScript.eventRouter.RunEvent(eventId2, pbArrs, "CellBuildingPathEvent", False)
                            If result = "#Error" Then Return {"#Error"}
                        End If

                        'событие CellBuildingPathEvent клетки
                        If result <> "False" Then
                            Dim eventId As Integer = mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("CellBuildingPathEvent").ThirdLevelEventId(transitCellId)
                            If eventId > 0 Then
                                result = mScript.eventRouter.RunEvent(eventId, pbArrs, "CellBuildingPathEvent", False)
                                If result = "#Error" Then Return {"#Error"}
                            End If
                        End If
                    End If

                    If result <> "False" Then
                        If directions > 0 Then
                            'новое ветвление пути
                            Dim lstNew As New List(Of Integer)
                            lstNew.AddRange(lstCur)
                            lstNew(lstNew.Count - 1) = pos
                            lstPath.Add(lstNew)
                        Else
                            'продолжение пути
                            lstCur.Add(pos)
                        End If
                        If cellsMap(pos) = 3 Then
                            If directions > 0 Then lstCur = lstPath.Last
                            Exit Do 'путь построен
                        End If
                        directions += 1
                        cellsMap(pos) = 1 'клетка посещена
                    End If
                End If

                If directions = 0 Then
                    'тупиковая ветвь - удаляем
                    lstPath.RemoveAt(lstId)
                    If lstPath.Count = 0 Then
                        lstPath.Clear()
                        lstCur.Clear()
                        Return {"-1"} 'путь не найден
                    End If
                End If
            Next lstId
        Loop

        'возвращаем массив из Id клеток пути
        Dim arrPath() As String
        ReDim arrPath(lstCur.Count - 1)
        For i As Integer = 0 To lstCur.Count - 1
            'curX = Math.Floor(lstCur(i) / CellsByX) + 1
            'curY = lstCur(i) - (curX - 1) * CellsByX + 1
            'Dim cellId As Integer = GetCellIdByXY(GVARS.G_CURMAP, curX, curY)
            Dim cellId As Integer = lstPosToCellId(lstCur(i))
            arrPath(i) = cellId.ToString
        Next

        lstPath.Clear()
        lstCur.Clear()
        Return arrPath
    End Function

    Private Function BuildPath_PosCellLeft(ByVal curX As Integer, ByVal curY As Integer, ByVal cellsByX As Integer) As Integer
        If curX < 2 Then Return -1
        'Return (curX - 2) * cellsByX + curX - 1
        Return (curY - 1) * cellsByX + curX - 1
    End Function

    Private Function BuildPath_PosCellRight(ByVal curX As Integer, ByVal curY As Integer, ByVal cellsByX As Integer) As Integer
        If curX >= cellsByX Then Return -1
        'Return curX * cellsByX + curY - 1
        Return (curY - 1) * cellsByX + curX + 1
    End Function

    Private Function BuildPath_PosCellTop(ByVal curX As Integer, ByVal curY As Integer, ByVal cellsByX As Integer) As Integer
        If curY < 2 Then Return -1
        'Return (curX - 1) * cellsByX + curY - 2
        Return (curY - 2) * cellsByX + curX
    End Function

    Private Function BuildPath_PosCellBottom(ByVal curX As Integer, ByVal curY As Integer, ByVal cellsByX As Integer, ByVal cellsByY As Integer) As Integer
        If curY >= cellsByY Then Return -1
        'Return (curX - 1) * cellsByX + curY
        Return curY * cellsByX + curX
    End Function

    Private Function BuildPath_PosCellLeftTop(ByVal curX As Integer, ByVal curY As Integer, ByVal cellsByX As Integer) As Integer
        If curY < 2 OrElse curX < 2 Then Return -1
        'Return (curX - 2) * cellsByX + curY - 2
        Return (curY - 2) * cellsByX + curX - 1
    End Function

    Private Function BuildPath_PosCellRightTop(ByVal curX As Integer, ByVal curY As Integer, ByVal cellsByX As Integer) As Integer
        If curY < 2 OrElse curX >= cellsByX Then Return -1
        'Return curX * cellsByX + curY - 2
        Return (curY - 2) * cellsByX + curX + 1
    End Function

    Private Function BuildPath_PosCellLeftBottom(ByVal curX As Integer, ByVal curY As Integer, ByVal cellsByX As Integer, ByVal cellsByY As Integer) As Integer
        If curY >= cellsByY OrElse curX < 2 Then Return -1
        'Return (curX - 2) * cellsByX + curY
        Return curY * cellsByX + curX - 1
    End Function

    Private Function BuildPath_PosCellRightBottom(ByVal curX As Integer, ByVal curY As Integer, ByVal cellsByX As Integer, ByVal cellsByY As Integer) As Integer
        If curY >= cellsByY OrElse curX >= cellsByX Then Return -1
        'Return curX * cellsByX + curY
        Return curY * cellsByX + curX + 1
    End Function


    ''' <summary>Возвращает html-элемент TD данной клетки</summary>
    ''' <param name="cellId"></param>
    Public Function GetCellHtmlElement(ByVal cellId As Integer, ByRef hDoc As HtmlDocument) As HtmlElement
        If IsNothing(hDoc) Then
            hDoc = questEnvironment.wbMap.Document
            If IsNothing(hDoc) Then Return Nothing
        End If

        Dim hTable As HtmlElement = hDoc.GetElementById("MapMain")
        If IsNothing(hTable) Then Return Nothing
        hTable = hTable.Children(0) 'TBody

        For i As Integer = 0 To hTable.Children.Count - 1
            'для каждого TR
            Dim TR As HtmlElement = hTable.Children(i)
            For j As Integer = 0 To TR.Children.Count - 1
                'для каждого TD
                Dim TD As HtmlElement = TR.Children(j)
                Dim strCell As String = TD.GetAttribute("cellId")
                If String.IsNullOrEmpty(strCell) = False Then
                    If cellId = Val(strCell) Then Return TD
                End If
            Next j
        Next i

        Return Nothing
    End Function

    ''' <summary>Возвращает html-элемент IMG данной клетки</summary>
    Public Function GetCellImageHTMLElement(ByVal cellId As Integer, ByRef hDoc As HtmlDocument) As HtmlElement
        If IsNothing(hDoc) Then
            hDoc = questEnvironment.wbMap.Document
            If IsNothing(hDoc) Then Return Nothing
        End If

        For i As Integer = 0 To hDoc.Images.Count - 1
            Dim hImg As HtmlElement = hDoc.Images(i)
            Dim strCell As String = hImg.GetAttribute("cellId")
            If String.IsNullOrEmpty(strCell) Then Continue For
            If CInt(strCell) = cellId Then Return hImg
        Next i

        Return Nothing
    End Function

    ''' <summary>
    ''' Убирает туман с текущей клетки и, если надо, то с соседних
    ''' </summary>
    ''' <param name="arrParams"></param>
    Public Sub ClearNearbyFog(ByRef arrParams() As String)
        If GVARS.G_CURMAP = -1 OrElse GVARS.G_CURMAPCELL = -1 Then Return
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim fogStyle As Integer = 0
        ReadPropertyInt(classId, "FogStyle", GVARS.G_CURMAP, -1, fogStyle, arrParams)

        Dim hCurCell As HtmlElement = GetCellHtmlElement(GVARS.G_CURMAPCELL, Nothing)
        If IsNothing(hCurCell) Then Return
        If mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("Fog").ThirdLevelProperties(GVARS.G_CURMAPCELL) <> "False" Then
            If PropertiesRouter(classId, "Fog", {GVARS.G_CURMAP.ToString, GVARS.G_CURMAPCELL.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "#Error" Then Return
        End If

        If fogStyle = 1 Then Return 'только текущая клетка
        'далее если 0 - тумана нет, то просто меняются свойста Fog соседних клеток на False; если же 2 - 8 клеток, то будет изменено визуальное представление

        Dim x As Integer = Val(hCurCell.GetAttribute("x")), y As Integer = Val(hCurCell.GetAttribute("y"))
        Dim xTotal As Integer = 10, yTotal As Integer = 10
        ReadPropertyInt(classId, "CellsByX", GVARS.G_CURMAP, -1, xTotal, arrParams)
        ReadPropertyInt(classId, "CellsByY", GVARS.G_CURMAP, -1, yTotal, arrParams)
        Dim cellClassEmpty As String = ""
        If ReadProperty(classId, "CellEmptyClass", GVARS.G_CURMAP, -1, cellClassEmpty, arrParams) = False Then Return
        cellClassEmpty = UnWrapString(cellClassEmpty)

        Dim hCell As HtmlElement = Nothing, cellId As Integer = -1

        If y > 1 Then
            'top
            hCell = hCurCell.Parent.Parent.Children(y - 2).Children(x - 1)
            cellId = hCell.GetAttribute("cellId")
            If cellId > -1 Then
                If mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("Fog").ThirdLevelProperties(cellId) <> "False" Then
                    If PropertiesRouter(classId, "Fog", {GVARS.G_CURMAP.ToString, cellId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "#Error" Then Return
                End If
            Else
                hCell.SetAttribute("ClassName", cellClassEmpty)
            End If
        End If

        If x > 1 Then
            'left
            hCell = hCurCell.Parent.Children(x - 2)
            cellId = hCell.GetAttribute("cellId")
            If cellId > -1 Then
                If mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("Fog").ThirdLevelProperties(cellId) <> "False" Then
                    If PropertiesRouter(classId, "Fog", {GVARS.G_CURMAP.ToString, cellId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "#Error" Then Return
                End If
            Else
                hCell.SetAttribute("ClassName", cellClassEmpty)
            End If

            'left top
            If y > 1 Then
                hCell = hCurCell.Parent.Parent.Children(y - 2).Children(x - 2)
                cellId = hCell.GetAttribute("cellId")
                If cellId > -1 Then
                    If mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("Fog").ThirdLevelProperties(cellId) <> "False" Then
                        If PropertiesRouter(classId, "Fog", {GVARS.G_CURMAP.ToString, cellId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "#Error" Then Return
                    End If
                Else
                    hCell.SetAttribute("ClassName", cellClassEmpty)
                End If
            End If

            'left bottom
            If y < yTotal Then
                hCell = hCurCell.Parent.Parent.Children(y).Children(x - 2)
                cellId = hCell.GetAttribute("cellId")
                If cellId > -1 Then
                    If mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("Fog").ThirdLevelProperties(cellId) <> "False" Then
                        If PropertiesRouter(classId, "Fog", {GVARS.G_CURMAP.ToString, cellId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "#Error" Then Return
                    End If
                Else
                    hCell.SetAttribute("ClassName", cellClassEmpty)
                End If
            End If
        End If

        If y < yTotal Then
            'bottom
            hCell = hCurCell.Parent.Parent.Children(y).Children(x - 1)
            cellId = hCell.GetAttribute("cellId")
            If cellId > -1 Then
                If mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("Fog").ThirdLevelProperties(cellId) <> "False" Then
                    If PropertiesRouter(classId, "Fog", {GVARS.G_CURMAP.ToString, cellId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "#Error" Then Return
                End If
            Else
                hCell.SetAttribute("ClassName", cellClassEmpty)
            End If
        End If

        If x < xTotal Then
            'right
            hCell = hCurCell.Parent.Children(x)
            cellId = hCell.GetAttribute("cellId")
            If cellId > -1 Then
                If mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("Fog").ThirdLevelProperties(cellId) <> "False" Then
                    If PropertiesRouter(classId, "Fog", {GVARS.G_CURMAP.ToString, cellId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "#Error" Then Return
                End If
            Else
                hCell.SetAttribute("ClassName", cellClassEmpty)
            End If

            'right top
            If y > 1 Then
                hCell = hCurCell.Parent.Parent.Children(y - 2).Children(x)
                cellId = hCell.GetAttribute("cellId")
                If cellId > -1 Then
                    If mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("Fog").ThirdLevelProperties(cellId) <> "False" Then
                        If PropertiesRouter(classId, "Fog", {GVARS.G_CURMAP.ToString, cellId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "#Error" Then Return
                    End If
                Else
                    hCell.SetAttribute("ClassName", cellClassEmpty)
                End If
            End If

            'right bottom
            If y < yTotal Then
                hCell = hCurCell.Parent.Parent.Children(y).Children(x)
                cellId = hCell.GetAttribute("cellId")
                If cellId > -1 Then
                    If mScript.mainClass(classId).ChildProperties(GVARS.G_CURMAP)("Fog").ThirdLevelProperties(cellId) <> "False" Then
                        If PropertiesRouter(classId, "Fog", {GVARS.G_CURMAP.ToString, cellId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "#Error" Then Return
                    End If
                Else
                    hCell.SetAttribute("ClassName", cellClassEmpty)
                End If
            End If
        End If
    End Sub

    ''' <summary>Убирает туман с указанной клетки</summary>
    Public Sub CellChangeFog(ByVal fogRemove As Boolean, ByVal cellId As Integer, ByRef arrParams() As String, Optional hCell As HtmlElement = Nothing)
        If GVARS.G_CURMAP = -1 Then Return
        If IsNothing(hCell) Then hCell = GetCellHtmlElement(cellId, Nothing)
        If IsNothing(hCell) Then Return

        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim imgPath As String = ""
        ReadProperty(classId, "CellPicture", GVARS.G_CURMAP, cellId, imgPath, arrParams)
        imgPath = UnWrapString(imgPath)


        If fogRemove Then
            'убираем туман

            'получаем нужный класс клетки
            Dim cellClass As String = ""
            Dim showStyle As Integer = 0
            ReadPropertyInt(classId, "ShowStyle", GVARS.G_CURMAP, cellId, showStyle, arrParams)
            If showStyle = 2 Then Return 'клетка всегда спрятана
            ReadProperty(classId, "CellUsualClass", GVARS.G_CURMAP, cellId, cellClass, arrParams)

            'If cellId = GVARS.G_CURMAPCELL Then
            '    ReadProperty(classId, "CellCurrentClass", GVARS.G_CURMAP, -1, cellClass, arrParams)
            'Else
            '    Dim showStyle As Integer = 0
            '    ReadPropertyInt(classId, "ShowStyle", GVARS.G_CURMAP, cellId, showStyle, arrParams)
            '    If showStyle = 2 Then Return 'клетка всегда спрятана
            '    ReadProperty(classId, "CellUsualClass", GVARS.G_CURMAP, cellId, cellClass, arrParams)
            'End If
            cellClass = UnWrapString(cellClass)

            'убираем класс тумана
            hCell.SetAttribute("ClassName", cellClass)

            'вставляем контент
            Dim strContent As String = ""
            ReadProperty(classId, "Content", GVARS.G_CURMAP, cellId, strContent, arrParams)
            strContent = UnWrapString(strContent)
            If String.IsNullOrEmpty(strContent) = False Then hCell.InnerHtml = strContent

            'делаем картинку клетки видимой
            If String.IsNullOrEmpty(imgPath) = False Then
                Dim hImg As HtmlElement = GetCellImageHTMLElement(cellId, hCell.Document)
                If IsNothing(hImg) Then Return
                Dim strStyle As String = hImg.Style
                If String.IsNullOrEmpty(strStyle) OrElse strStyle.EndsWith("display: none;") = False Then Return
                strStyle = strStyle.Substring(0, strStyle.Length - 15)
                hImg.Style = strStyle
            End If
        Else
            'добавляем туман
            'получаем класс тумана
            Dim CellInFogClass As String = ""
            ReadProperty(classId, "CellInFogClass", GVARS.G_CURMAP, -1, CellInFogClass, arrParams)
            CellInFogClass = UnWrapString(CellInFogClass)

            'добавляем класс тумана
            hCell.SetAttribute("ClassName", CellInFogClass)
            hCell.InnerHtml = "" 'убираем контент
            'делаем картинку клетки невидимой
            If String.IsNullOrEmpty(imgPath) = False Then
                Dim hImg As HtmlElement = GetCellImageHTMLElement(cellId, hCell.Document)
                If IsNothing(hImg) Then Return
                Dim strStyle As String = hImg.Style
                If String.IsNullOrEmpty(strStyle) OrElse strStyle.EndsWith("display: none;") = False Then
                    strStyle &= "display: none;"
                End If
                hImg.Style = strStyle
            End If
        End If

    End Sub

    ''' <summary>
    ''' Возвращает находится ли в тумане текущая клетка
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="arrParams"></param>
    Private Function IsEmptyCellInFog(ByVal x As Integer, ByVal y As Integer, ByRef arrParams() As String) As Boolean
        If GVARS.G_CURMAP = -1 Then Return True
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim fogStyle As Integer = 0
        ReadPropertyInt(classId, "FogStyle", GVARS.G_CURMAP, -1, fogStyle, arrParams)
        If fogStyle = 0 Then
            Return False 'туман отсутствует
        ElseIf fogStyle = 1 Then
            Return True 'в тумане все, кроме посещенных клеток - значит пустая тоже в тумане
        End If

        'при посещении клетки открываются 8 соседних - ищем открыты ли соседние
        Dim nCellId As Integer = -1, fog As Boolean = False
        If y > 1 Then
            'top
            nCellId = GetCellIdByXY(GVARS.G_CURMAP, x, y - 1)
            If nCellId > -1 Then
                ReadPropertyBool(classId, "Fog", GVARS.G_CURMAP, nCellId, fog, arrParams)
                If fog = False Then Return False
            End If
        End If

        If x > 1 Then
            'left
            nCellId = GetCellIdByXY(GVARS.G_CURMAP, x - 1, y)
            If nCellId > -1 Then
                ReadPropertyBool(classId, "Fog", GVARS.G_CURMAP, nCellId, fog, arrParams)
                If fog = False Then Return False
            End If

            If y > 1 Then
                'left top
                nCellId = GetCellIdByXY(GVARS.G_CURMAP, x - 1, y - 1)
                If nCellId > -1 Then
                    ReadPropertyBool(classId, "Fog", GVARS.G_CURMAP, nCellId, fog, arrParams)
                    If fog = False Then Return False
                End If
            End If

            'left bottom
            nCellId = GetCellIdByXY(GVARS.G_CURMAP, x - 1, y + 1)
            If nCellId > -1 Then
                ReadPropertyBool(classId, "Fog", GVARS.G_CURMAP, nCellId, fog, arrParams)
                If fog = False Then Return False
            End If
        End If

        'bottom
        nCellId = GetCellIdByXY(GVARS.G_CURMAP, x, y + 1)
        If nCellId > -1 Then
            ReadPropertyBool(classId, "Fog", GVARS.G_CURMAP, nCellId, fog, arrParams)
            If fog = False Then Return False
        End If

        'right
        nCellId = GetCellIdByXY(GVARS.G_CURMAP, x + 1, y)
        If nCellId > -1 Then
            ReadPropertyBool(classId, "Fog", GVARS.G_CURMAP, nCellId, fog, arrParams)
            If fog = False Then Return False
        End If

        'right top
        nCellId = GetCellIdByXY(GVARS.G_CURMAP, x + 1, y - 1)
        If nCellId > -1 Then
            ReadPropertyBool(classId, "Fog", GVARS.G_CURMAP, nCellId, fog, arrParams)
            If fog = False Then Return False
        End If

        'right bottom
        nCellId = GetCellIdByXY(GVARS.G_CURMAP, x + 1, y + 1)
        If nCellId > -1 Then
            ReadPropertyBool(classId, "Fog", GVARS.G_CURMAP, nCellId, fog, arrParams)
            If fog = False Then Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' Возвращает Id клетки по ее координатам
    ''' </summary>
    ''' <param name="mapId">Id карты</param>
    ''' <param name="x">Координата Х клетки</param>
    ''' <param name="y">Координата Y клетки</param>
    ''' <returns>Id клметки или -1 если не найдена</returns>
    Public Function GetCellIdByXY(ByVal mapId As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMap).ChildProperties(mapId)("Name")
        If IsNothing(ch.ThirdLevelProperties) OrElse ch.ThirdLevelProperties.Count = 0 Then Return -1

        Dim cX As Integer = -1, cY As Integer = -1
        For i As Integer = 0 To ch.ThirdLevelProperties.Count - 1
            If ReadPropertyInt(classMap, "CellX", mapId, i, cX, Nothing) = False Then Return -1
            If ReadPropertyInt(classMap, "CellY", mapId, i, cY, Nothing) = False Then Return -1
            If cX = x AndAlso cY = y Then Return i
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Возвращает Id клетки по ее локации
    ''' </summary>
    ''' <param name="mapId">Id карты</param>
    ''' <param name="locId">Id локации</param>
    Public Function GetCellByLocation(ByVal mapId As Integer, ByVal locId As Integer) As Integer
        Dim classMap As Integer = mScript.mainClassHash("Map")
        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMap).ChildProperties(mapId)("Name")
        If IsNothing(ch.ThirdLevelProperties) OrElse ch.ThirdLevelProperties.Count = 0 Then Return -1

        Dim classL As Integer = mScript.mainClassHash("L")
        Dim locName As String = mScript.mainClass(classL).ChildProperties(locId)("Name").Value
        Dim locIdStr As String = locId.ToString

        Dim curLoc As String = ""
        For i As Integer = 0 To ch.ThirdLevelProperties.Count - 1
            If ReadProperty(classMap, "Location", mapId, i, curLoc, Nothing) = False Then Return -1
            If curLoc = locIdStr OrElse String.Compare(curLoc, locName, True) = 0 Then Return i
        Next
        Return -1
    End Function

    ''' <summary>Возвращает координаты html-элемента относительно BODY</summary>
    ''' <param name="hEl">html-элемент</param>
    Public Function GetHTMLelementCoordinates(ByVal hEl As HtmlElement) As Point
        Dim loc As New Point(0, 0)

        Do While hEl.TagName <> "BODY"
            loc.X += hEl.OffsetRectangle.Left
            loc.Y += hEl.OffsetRectangle.Top
            hEl = hEl.OffsetParent
        Loop

        Return loc
    End Function

    ''' <summary>
    ''' Возвращает видна ли клетка фактически
    ''' </summary>
    ''' <param name="mapId"></param>
    ''' <param name="cellId"></param>
    Public Function IsCellVisible(ByVal mapId As Integer, cellId As Integer) As Boolean
        If GVARS.G_CURMAP < 0 OrElse GVARS.G_CURMAP <> mapId Then Return False
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim fogStyle As Integer = 0, showStyle As Integer = 0
        ReadPropertyInt(classId, "FogStyle", mapId, -1, fogStyle, Nothing)
        ReadPropertyInt(classId, "ShowStyle", mapId, cellId, showStyle, Nothing)
        If showStyle = 2 Then
            Return False 'всегда спрятана
        ElseIf showStyle = 1 OrElse fogStyle = 0 Then
            Return True 'всегда видна / туман отсутствует
        End If

        'showStyle = 0 - клетка не видна пока в тумане, туман есть
        Dim fog As Boolean = False
        ReadPropertyBool(classId, "Fog", mapId, cellId, fog, Nothing)
        Return Not fog
    End Function

    ''' <summary>Событие при попадании курсора в область клетки, если на ней нет картинки. Выводится инфо о клетке</summary>
    Private Sub del_CellMouseOver(sender As HtmlElement, e As HtmlElementEventArgs)
        'sender - TD
        Dim cellId As Integer = Val(sender.GetAttribute("CellId"))
        Dim mapId As Integer = Val(sender.GetAttribute("mapId"))
        If mapId < 0 OrElse cellId < 0 Then Return

        'получаем локацию
        Dim arrs() As String = {mapId.ToString, cellId.ToString}
        Dim classId As Integer = mScript.mainClassHash("Map")
        'Dim classL As Integer = mScript.mainClassHash("L")
        'Dim loc As String = ""
        'If ReadProperty(classId, "Location", mapId, cellId, loc, arrs) = False Then Return
        'Dim locId As Integer = GetSecondChildIdByName(loc, mScript.mainClass(classL).ChildProperties)
        'If locId < 0 Then Return

        'Вывод описания
        Dim result As String = "", retFormat As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL

        'читаем DescriptionTemplate
        If mapId >= 0 Then
            If ReadProperty(classId, "DescriptionTemplate", mapId, -1, result, arrs, retFormat) = False Then Return
            If retFormat = MatewScript.ReturnFormatEnum.TO_STRING Then result = UnWrapString(result)
        End If

        Dim hCapt As HtmlElement = sender.Document.GetElementById("CellInfo")
        If IsNothing(hCapt) Then Return


        If result.Length = 0 OrElse IsCellVisible(mapId, cellId) = False Then
            'hCapt.Style = "display:none"
            hCapt.InnerHtml = "&nbsp;"
        Else
            'hCapt.Style = ""
            hCapt.InnerHtml = result
        End If
    End Sub

    Private Sub del_CellMouseLeave(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim hCapt As HtmlElement = sender.Document.GetElementById("CellInfo")
        If IsNothing(hCapt) Then Return
        hCapt.InnerHtml = "&nbsp;"
    End Sub

    Public Sub del_CellClick(sender As HtmlElement, e As HtmlElementEventArgs)
        'sender - TD / IMG
        Dim cellId As Integer = Val(sender.GetAttribute("CellId"))
        Dim mapId As Integer = Val(sender.GetAttribute("mapId"))
        If mapId < 0 OrElse cellId < 0 Then Return
        Static movingStarted As Boolean = False
        If movingStarted Then Return

        'видна ли клетка?
        'Dim cellVisible As Boolean = IsCellVisible(mapId, cellId
        If IsCellVisible(mapId, cellId) = False Then Return
        'получаем локацию
        Dim arrs() As String = {mapId.ToString, cellId.ToString}
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim classL As Integer = mScript.mainClassHash("L")
        Dim loc As String = ""
        If ReadProperty(classId, "Location", mapId, cellId, loc, arrs) = False Then Return
        Dim locId As Integer = GetSecondChildIdByName(loc, mScript.mainClass(classL).ChildProperties)
        If locId < 0 Then Return

        If sender.TagName = "IMG" AndAlso EventGeneratedFromScript = False Then
            If prevTransparent OrElse IsNothing(prevCellBitmap) OrElse Object.Equals(prevHimg, sender) = False Then Return 'кликнули по прозрачному

            'восстанавливаем исходную картинку
            Dim bkPicture As String = ""
            If ReadProperty(classId, "CellPicture", mapId, cellId, bkPicture, arrs) = False Then Return
            bkPicture = UnWrapString(bkPicture)
            If String.IsNullOrEmpty(bkPicture) = False Then
                Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, bkPicture)
                If FileIO.FileSystem.FileExists(fPath) Then
                    sender.SetAttribute("src", bkPicture)
                End If
            End If

            ClearPreviousBitmap()
        End If

        Dim walking As Integer = 0
        ReadPropertyInt(classId, "Walking", mapId, -1, walking, arrs)
        Dim shouldGo As Boolean = False
        Dim arrPath() As String = Nothing 'путь к клетке
        If walking = 2 Then
            shouldGo = True 'разрешено перемещаться как угодно
        ElseIf walking = 0 Then
            shouldGo = False
        ElseIf walking = 1 Then
            'walking = 1 - разрешено перемещаться на посещенные клетки
            Dim visits As Integer = 0
            ReadPropertyInt(classL, "Visits", locId, -1, visits, {locId.ToString})
            If visits > 0 OrElse EventGeneratedFromScript Then
                shouldGo = True
            Else
                shouldGo = False
            End If
        ElseIf walking = 3 Then
            'walking = 3 - если открыт путь
            arrPath = BuildPath(GVARS.G_CURMAPCELL, cellId, False, BuildStyleEnum.OnlyVisible, True)
            If arrPath(0) = "-1" Then
                shouldGo = False
            Else
                shouldGo = True
            End If
        Else
            'walking = 4 - если клетки по пути посещены
            arrPath = BuildPath(GVARS.G_CURMAPCELL, cellId, False, BuildStyleEnum.OnlyVisited, True)
            If arrPath(0) = "-1" Then
                shouldGo = False
            Else
                shouldGo = True
            End If
        End If

        'Событие выбора клетки карты CellChangedEvent
        Dim res As String = ""
        Dim wasEvent As Boolean = False
        Dim returnedFalse As Boolean = False, returnedTrue As Boolean = False
        arrs = {mapId.ToString, cellId.ToString, shouldGo.ToString}
        'глобальное
        Dim eventId As Integer = mScript.mainClass(classId).Properties("CellChangedEvent").eventId
        If eventId > 0 Then
            wasEvent = True
            res = mScript.eventRouter.RunEvent(eventId, arrs, "CellChangedEvent", False)
            If res = "#Error" Then
                Return
            ElseIf res = "True" Then
                returnedTrue = True
            ElseIf res = "False" Then
                returnedFalse = True
            End If
        End If

        'данной карты
        If returnedFalse = False AndAlso returnedTrue = False Then
            eventId = mScript.mainClass(classId).ChildProperties(mapId)("CellChangedEvent").eventId
            If eventId > 0 Then
                wasEvent = True
                res = mScript.eventRouter.RunEvent(eventId, arrs, "CellChangedEvent", False)
                If res = "#Error" Then
                    Return
                ElseIf res = "True" Then
                    returnedTrue = True
                ElseIf res = "False" Then
                    returnedFalse = True
                End If
            End If
        End If

        'данной клетки
        If returnedFalse = False AndAlso returnedTrue = False Then
            eventId = mScript.mainClass(classId).ChildProperties(mapId)("CellChangedEvent").ThirdLevelEventId(cellId)
            If eventId > 0 Then
                wasEvent = True
                res = mScript.eventRouter.RunEvent(eventId, arrs, "CellChangedEvent", False)
                If res = "#Error" Then
                    Return
                ElseIf res = "True" Then
                    returnedTrue = True
                ElseIf res = "False" Then
                    returnedFalse = True
                End If
            End If
        End If

        If returnedFalse Then
            shouldGo = False
        ElseIf returnedTrue Then
            shouldGo = True
        End If

        If shouldGo Then
            'LocationExitEvent выполняем ДО транзита
            Dim resLocExit As String = ""
            If GVARS.G_CURLOC > -1 Then
                Dim exArrs = {GVARS.G_CURLOC.ToString, locId.ToString, "True"}
                'событие свойств по умолчанию
                eventId = mScript.mainClass(classL).Properties("LocationExitEvent").eventId
                If eventId > 0 Then
                    resLocExit = mScript.eventRouter.RunEvent(eventId, exArrs, "LocationExitEvent", False)
                    If resLocExit = "#Error" Then Return
                End If
                'событие выхода из текущей локации
                If resLocExit <> "False" Then
                    eventId = mScript.mainClass(classL).ChildProperties(GVARS.G_CURLOC)("LocationExitEvent").eventId
                    If eventId > 0 Then
                        resLocExit = mScript.eventRouter.RunEvent(eventId, exArrs, "LocationExitEvent", False)
                        If resLocExit = "#Error" Then Return
                    End If
                End If
            End If

            Dim goParams() As String = {locId.ToString, "True", "True"}

            If IsNothing(arrPath) Then
                arrPath = mapManager.BuildPath(GVARS.G_CURMAPCELL, cellId, False, BuildStyleEnum.OnlyVisible, True)
            End If
            If arrPath.Length > 0 AndAlso arrPath(0) = "-1" Then
                Dim transType As Integer = 0
                ReadPropertyInt(classId, "TransitionType", GVARS.G_CURMAP, -1, transType, arrs)
                If transType = 1 AndAlso walking <> 2 Then
                    'от начала к концу                    
                    resLocExit = "False"
                    shouldGo = False
                End If
            End If

            If resLocExit <> "False" AndAlso arrPath.Length > 2 Then
                'построен путь. Проходим по нему, вызывая события CellTransitEvent в каждой клетке
                Dim eventId1 As Integer = mScript.mainClass(classId).Properties("CellTransitEvent").eventId
                Dim eventId2 As Integer = mScript.mainClass(classId).ChildProperties(mapId)("CellTransitEvent").eventId
                Dim mapIdStr As String = mapId.ToString
                Dim cellFinalStr As String = cellId.ToString
                Dim cellStartStr As String = GVARS.G_CURMAPCELL.ToString

                res = ""
                For i As Integer = 1 To arrPath.Length - 2 'исключаем начальную и конечную клетку
                    'CellTransitEvent глобальный
                    Dim arrTransit() As String = {mapIdStr, arrPath(i), cellStartStr, cellFinalStr}
                    If eventId1 > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId1, {}, "CellTransitEvent", False)
                        If res = "#Error" Then
                            Return
                        ElseIf String.IsNullOrEmpty(res) = False Then
                            cellId = CInt(arrPath(i))
                            If ReadProperty(classId, "Location", mapId, cellId, loc, arrs) = False Then Return
                            locId = GetSecondChildIdByName(loc, mScript.mainClass(classL).ChildProperties)
                            If locId < 0 Then
                                mScript.LAST_ERROR = "Не найдена локация " & loc & " транзитной клетки."
                                Return
                            End If
                            goParams = {locId.ToString, "True", "True", res}
                            ReDim Preserve arrPath(i)
                            Exit For
                        End If
                    End If

                    'CellTransitEvent карты
                    If eventId2 > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId2, {}, "CellTransitEvent", False)
                        If res = "#Error" Then
                            Return
                        ElseIf String.IsNullOrEmpty(res) = False Then
                            cellId = CInt(arrPath(i))
                            If ReadProperty(classId, "Location", mapId, cellId, loc, arrs) = False Then Return
                            locId = GetSecondChildIdByName(loc, mScript.mainClass(classL).ChildProperties)
                            If locId < 0 Then
                                mScript.LAST_ERROR = "Не найдена локация " & loc & " транзитной клетки."
                                Return
                            End If
                            goParams = {locId.ToString, "True", "True", res}
                            ReDim Preserve arrPath(i)
                            Exit For
                        End If
                    End If

                    'CellTransitEvent клетки
                    eventId = mScript.mainClass(classId).ChildProperties(mapId)("CellTransitEvent").ThirdLevelEventId(CInt(arrPath(i)))
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, {}, "CellTransitEvent", False)
                        If res = "#Error" Then
                            Return
                        ElseIf String.IsNullOrEmpty(res) = False Then
                            cellId = CInt(arrPath(i))
                            If ReadProperty(classId, "Location", mapId, cellId, loc, arrs) = False Then Return
                            locId = GetSecondChildIdByName(loc, mScript.mainClass(classL).ChildProperties)
                            If locId < 0 Then
                                mScript.LAST_ERROR = "Не найдена локация " & loc & " транзитной клетки."
                                Return
                            End If
                            goParams = {locId.ToString, "True", "True", res}
                            ReDim Preserve arrPath(i)
                            Exit For
                        End If
                    End If

                    'События способностей AbilityOnCellTransitEvent
                    Dim abList As SortedList(Of Integer, Integer) = HeroGetAbilitySetsIdList() 'Key - heroId, Value - AbSetId
                    If abList.Count > 0 Then
                        Dim classAb As Integer = mScript.mainClassHash("Ab")
                        Dim globalEventId As Integer = mScript.mainClass(classAb).Properties("AbilityOnCellTransitEvent").eventId
                        Dim abArrs() As String = {"", "", "", GVARS.G_CURMAP.ToString, cellId.ToString}
                        For j As Integer = 0 To abList.Count - 1
                            Dim hId As Integer = abList.ElementAt(j).Key
                            Dim abSetId As Integer = abList.ElementAt(j).Value
                            abArrs(0) = abSetId.ToString
                            abArrs(2) = hId.ToString
                            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityOnCellTransitEvent")
                            Dim abSetEventId As Integer = ch.eventId

                            For abId As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                'запускаем событие AbilityOnChangeLocEvent для кажодй способности из всех доступных наборов, прикрепленных к персонажам
                                abArrs(1) = abId.ToString

                                Dim isEnabled As Boolean = True
                                If ReadPropertyBool(classAb, "Enabled", abSetId, abId, isEnabled, abArrs) = False Then Return
                                If Not isEnabled Then Continue For 'способность недоступна - пропускаем

                                If globalEventId > 0 Then
                                    'глобальное событие
                                    If mScript.eventRouter.RunEvent(globalEventId, abArrs, "AbilityOnCellTransitEvent", False) = "#Error" Then Return
                                End If

                                If abSetEventId > 0 Then
                                    'событие данного набора
                                    If mScript.eventRouter.RunEvent(abSetEventId, abArrs, "AbilityOnCellTransitEvent", False) = "#Error" Then Return
                                End If

                                eventId = ch.ThirdLevelEventId(abId)
                                If eventId > 0 Then
                                    'событие данной способности
                                    If mScript.eventRouter.RunEvent(eventId, abArrs, "AbilityOnCellTransitEvent", False) = "#Error" Then Return
                                End If
                            Next abId
                        Next j

                    End If
                Next i
            End If

            If resLocExit <> "False" Then
                GVARS.G_PREVMAPCELL = GVARS.G_CURMAPCELL
                GVARS.G_CURMAPCELL = cellId
                mapManager.ClearNearbyFog(Nothing)

                'создаем массив для эффекта перехода от клетки к клетке
                Dim transType As Integer = 0, aPath() As Integer = Nothing
                ReadPropertyInt(classId, "TransitionType", GVARS.G_CURMAP, -1, transType, arrs)
                If transType = 1 Then
                    '1 от начала к концу
                    ReDim aPath(1)
                    aPath(0) = GVARS.G_PREVMAPCELL
                    aPath(1) = GVARS.G_CURMAPCELL
                ElseIf transType = 2 Then
                    '2 пройти по пути
                    If IsNothing(arrPath) = False AndAlso arrPath(0) <> "-1" Then
                        'путь строился и он существует
                        ReDim aPath(arrPath.Count - 1)
                        For i As Integer = 0 To arrPath.Length - 1
                            aPath(i) = CInt(arrPath(i))
                        Next
                    ElseIf IsNothing(arrPath) Then
                        'путь не строился
                        arrPath = mapManager.BuildPath(GVARS.G_PREVMAPCELL, GVARS.G_CURMAPCELL, False, BuildStyleEnum.OnlyVisible, True)
                        If arrPath(0) = "-1" Then
                            'пути нет
                            ReDim aPath(1)
                            aPath(0) = GVARS.G_PREVMAPCELL
                            aPath(1) = GVARS.G_CURMAPCELL
                        Else
                            ReDim aPath(arrPath.Count - 1)
                            For i As Integer = 0 To arrPath.Length - 1
                                aPath(i) = CInt(arrPath(i))
                            Next
                        End If
                    Else 'arrPath(0) = "-1"
                        'путь строился, но пути нет
                        ReDim aPath(1)
                        aPath(0) = GVARS.G_PREVMAPCELL
                        aPath(1) = GVARS.G_CURMAPCELL
                    End If
                End If

                'перемещаем текущую позицию в новую клетку
                movingStarted = True
                mapManager.ShowArrowInCurrentCell(aPath)
                movingStarted = False
                If Go(goParams, False) = "#Error" Then Return
            End If
            wasEvent = True 'для генерации RunScriptFinishedEvent
        End If

        If shouldGo Then
            'прячем карту если был совершен переход. Иначе оставляем
            Dim autoHidding As Boolean = False
            ReadPropertyBool(classId, "MapAutoClosing", GVARS.G_CURMAP, -1, autoHidding, arrs)
            If autoHidding Then
                'Прячем карту
                Dim slot As Integer = 0, isNewWindow As Boolean = False
                ReadPropertyInt(classId, "Slot", -1, -1, slot, arrs)
                If slot = 5 Then isNewWindow = True 'в отдельном окне

                If isNewWindow Then
                    frmMap.Hide()
                Else
                    questEnvironment.wbMap.Hide()
                End If
            End If
        End If

        'событие ScriptFinishedEvent
        If EventGeneratedFromScript = False AndAlso wasEvent Then
            If mScript.eventRouter.RunScriptFinishedEvent(Nothing) = "#Error" Then Return
        End If
        EventGeneratedFromScript = False
    End Sub

    Dim prevCellBitmap As Bitmap = Nothing, prevHimg As HtmlElement, prevTransparent As Boolean = True
    ''' <summary>Событие при попадании курсора в область картинки. Выводится инфо о клетке</summary>
    Private Sub del_CellImgMouseMove(sender As HtmlElement, e As HtmlElementEventArgs)
        'sender - IMG
        Dim cellId As Integer = Val(sender.GetAttribute("CellId"))
        Dim mapId As Integer = Val(sender.GetAttribute("mapId"))
        If mapId < 0 OrElse cellId < 0 Then Return
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim arrs() As String = {mapId.ToString, cellId.ToString}
        'получаем локацию
        Dim classL As Integer = mScript.mainClassHash("L")
        Dim loc As String = ""
        If ReadProperty(classId, "Location", mapId, cellId, loc, arrs) = False Then Return
        Dim locId As Integer = GetSecondChildIdByName(loc, mScript.mainClass(classL).ChildProperties)

        Dim picPath As String = "" ', changeHImg As Boolean = False
        If IsNothing(prevHimg) Then
            'первое вхождение - создаем prevCellBitmap , устанавливаем текущую prevHimg

            If locId > -1 Then
                HTMLRemoveCSSstyle(sender, "cursor")
            Else
                HTMLAddCSSstyle(sender, "cursor", "default")
            End If

            If locId > -1 Then
                'клетка с картинкой связана с локацией

                'картинка Hover
                ReadProperty(classId, "CellPictureHover", mapId, cellId, picPath, arrs)
                picPath = UnWrapString(picPath)
                Dim fullPath As String = ""
                If String.IsNullOrEmpty(picPath) = False Then
                    Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picPath)
                    If FileIO.FileSystem.FileExists(fPath) Then
                        'changeHImg = True
                        fullPath = fPath
                    End If
                End If

                If IsNothing(prevCellBitmap) Then
                    If String.IsNullOrEmpty(fullPath) Then
                        Dim pPath As String = ""
                        If GVARS.G_CURMAPCELL = cellId Then
                            ReadProperty(classId, "CellPictureCurrent", mapId, cellId, pPath, arrs)
                            pPath = UnWrapString(pPath)
                            If String.IsNullOrEmpty(pPath) Then
                                ReadProperty(classId, "CellPicture", mapId, cellId, pPath, arrs)
                                pPath = UnWrapString(pPath)
                            End If
                        Else
                            ReadProperty(classId, "CellPicture", mapId, cellId, pPath, arrs)
                            pPath = UnWrapString(pPath)
                        End If

                        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, pPath)
                        If FileIO.FileSystem.FileExists(fPath) Then
                            fullPath = fPath
                        End If
                    End If

                    If String.IsNullOrEmpty(fullPath) = False Then
                        CreateCellImageBitmap(mapId, cellId, fullPath, sender)
                    End If
                End If
            Else
                'клетка с картинкой не связана с локацией
                'картинка Hover
                ReadProperty(classId, "CellPicture", mapId, cellId, picPath, arrs)
                picPath = UnWrapString(picPath)
                Dim fullPath As String = ""
                If String.IsNullOrEmpty(picPath) = False Then
                    Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picPath)
                    If FileIO.FileSystem.FileExists(fPath) Then
                        'changeHImg = True
                        fullPath = fPath
                    End If
                End If

                If IsNothing(prevCellBitmap) AndAlso String.IsNullOrEmpty(fullPath) = False Then
                    CreateCellImageBitmap(mapId, cellId, fullPath, sender)
                End If
            End If

            prevHimg = sender
        End If


        If IsNothing(prevCellBitmap) = False Then
            'проверяем пиксел под мышью на предмет прозрачности
            Dim mousePos As Point = e.ClientMousePosition
            Dim elPos As Point = GetHTMLelementCoordinates(sender)
            Dim hHTML As HtmlElement = sender.Document.GetElementsByTagName("HTML")(0)

            mousePos.X -= elPos.X - hHTML.ScrollLeft
            mousePos.Y -= elPos.Y - hHTML.ScrollTop


            Dim pColor As Color = prevCellBitmap.GetPixel(mousePos.X, mousePos.Y)
            If pColor.A <= 10 Then
                'пиксел прозрачный - выход
                If prevTransparent Then Return 'выход уже был

                If locId > -1 Then
                    'картинка обычная
                    If GVARS.G_CURMAPCELL = cellId AndAlso locId > -1 Then
                        ReadProperty(classId, "CellPictureCurrent", mapId, cellId, picPath, arrs)
                        picPath = UnWrapString(picPath)
                        If String.IsNullOrEmpty(picPath) Then
                            ReadProperty(classId, "CellPicture", mapId, cellId, picPath, arrs)
                            picPath = UnWrapString(picPath)
                        End If
                    Else
                        ReadProperty(classId, "CellPicture", mapId, cellId, picPath, arrs)
                        picPath = UnWrapString(picPath)
                    End If

                    If String.IsNullOrEmpty(picPath) = False Then
                        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picPath)
                        If FileIO.FileSystem.FileExists(fPath) Then
                            sender.SetAttribute("src", picPath)
                        End If
                    End If
                End If

                Dim hCapt As HtmlElement = sender.Document.GetElementById("CellInfo")
                If IsNothing(hCapt) = False Then hCapt.InnerHtml = "&nbsp;"
                prevTransparent = True
                HTMLRemoveClass(sender, "CellImageHover")
            Else
                'пиксел непрозрачный - вход
                If Not prevTransparent Then Return 'вход уже был

                If locId > -1 Then
                    'картинка Hover
                    ReadProperty(classId, "CellPictureHover", mapId, cellId, picPath, arrs)
                    picPath = UnWrapString(picPath)
                    If String.IsNullOrEmpty(picPath) = False Then
                        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picPath)
                        If FileIO.FileSystem.FileExists(fPath) Then
                            sender.SetAttribute("src", picPath)
                        End If
                    End If
                End If

                'Вывод описания
                Dim result As String = "", retFormat As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL

                'читаем DescriptionTemplate
                If mapId >= 0 Then
                    If ReadProperty(classId, "DescriptionTemplate", mapId, -1, result, arrs, retFormat) = False Then Return
                    If retFormat = MatewScript.ReturnFormatEnum.TO_STRING Then result = UnWrapString(result)
                End If

                Dim hCapt As HtmlElement = sender.Document.GetElementById("CellInfo")
                If IsNothing(hCapt) Then Return
                If result.Length = 0 Then
                    'hCapt.Style = "display:none"
                    hCapt.InnerHtml = "&nbsp;"
                Else
                    'hCapt.Style = ""
                    hCapt.InnerHtml = result
                End If
                prevTransparent = False
                HTMLAddClass(sender, "CellImageHover")
            End If
        End If

    End Sub

    Private Sub del_CellImgMouseLeave(sender As HtmlElement, e As HtmlElementEventArgs)
        'sender - IMG
        ClearPreviousBitmap()

        Dim cellId As Integer = Val(sender.GetAttribute("CellId"))
        Dim mapId As Integer = Val(sender.GetAttribute("mapId"))
        If mapId < 0 OrElse cellId < 0 Then Return
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim arrs() As String = {mapId.ToString, cellId.ToString}

        'картинка обычная
        Dim picPath As String = ""
        If GVARS.G_CURMAPCELL = cellId Then
            ReadProperty(classId, "CellPictureCurrent", mapId, cellId, picPath, arrs)
            picPath = UnWrapString(picPath)
            If String.IsNullOrEmpty(picPath) Then
                ReadProperty(classId, "CellPicture", mapId, cellId, picPath, arrs)
                picPath = UnWrapString(picPath)
            End If
        Else
            ReadProperty(classId, "CellPicture", mapId, cellId, picPath, arrs)
            picPath = UnWrapString(picPath)
        End If

        If String.IsNullOrEmpty(picPath) = False Then
            Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picPath)
            If FileIO.FileSystem.FileExists(fPath) Then
                sender.SetAttribute("src", picPath)
            End If
        End If

        Dim hCapt As HtmlElement = sender.Document.GetElementById("CellInfo")
        If IsNothing(hCapt) = False Then hCapt.InnerHtml = "&nbsp;"
        HTMLRemoveClass(sender, "CellImageHover")
    End Sub

    Private Sub del_CellImgMouseDown(sender As HtmlElement, e As HtmlElementEventArgs)
        'sender IMG
        If prevTransparent OrElse IsNothing(prevCellBitmap) OrElse Object.Equals(prevHimg, sender) = False Then Return

        Dim cellId As Integer = Val(sender.GetAttribute("CellId"))
        Dim mapId As Integer = Val(sender.GetAttribute("mapId"))
        If mapId < 0 OrElse cellId < 0 Then Return
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim arrs() As String = {mapId.ToString, cellId.ToString}

        'получаем локацию
        Dim classL As Integer = mScript.mainClassHash("L")
        Dim loc As String = ""
        If ReadProperty(classId, "Location", mapId, cellId, loc, arrs) = False Then Return
        Dim locId As Integer = GetSecondChildIdByName(loc, mScript.mainClass(classL).ChildProperties)
        If locId < 0 Then Return

        'картинка Active
        Dim picPath As String = ""
        ReadProperty(classId, "CellPictureActive", mapId, cellId, picPath, arrs)
        picPath = UnWrapString(picPath)
        If String.IsNullOrEmpty(picPath) = False Then
            Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picPath)
            If FileIO.FileSystem.FileExists(fPath) Then
                sender.SetAttribute("src", picPath)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Создает prevBitmap - структуру для хранения последней картинки, на которую был наведен курсор. Надо для получения прозрачности/непрозрачности пискела
    ''' </summary>
    ''' <param name="mapId">Id карты</param>
    ''' <param name="cellId">Id клетки с картинкой</param>
    ''' <param name="picPath">Полный путь к файлу картинки</param>
    ''' <param name="hImg">html-элемент картинки IMG</param>
    Private Sub CreateCellImageBitmap(ByVal mapId As Integer, ByVal cellId As Integer, ByVal picPath As String, ByRef hImg As HtmlElement)
        If IsNothing(prevCellBitmap) = False Then prevCellBitmap.Dispose()
        Dim classId As Integer = mScript.mainClassHash("Map")
        Dim strWidth As String = "", strHeight As String = ""
        ReadProperty(classId, "CellPictureWidth", mapId, cellId, strWidth, Nothing)
        ReadProperty(classId, "CellPictureHeight", mapId, cellId, strHeight, Nothing)
        strWidth = UnWrapString(strWidth)
        strHeight = UnWrapString(strHeight)
        If String.IsNullOrEmpty(strWidth) = False OrElse String.IsNullOrEmpty(strHeight) = False Then
            'размеры картинки изменены - создаем графику и масштабируем картинку
            Dim pSize As Size = hImg.OffsetRectangle.Size
            Dim hCopy As Bitmap = Image.FromFile(picPath) 'получаем файл картинки
            prevCellBitmap = New Bitmap(pSize.Width, pSize.Height) 'получаем размеры
            Dim g As Graphics = Graphics.FromImage(prevCellBitmap)
            g.DrawImage(hCopy, 0, 0, pSize.Width, pSize.Height) 'копируем картинку в новых размерах
            hCopy.Dispose()
        Else
            'размеры исходные - просто получаем изображения
            prevCellBitmap = Image.FromFile(picPath)
        End If
    End Sub

    Public Sub ClearPreviousBitmap()
        prevHimg = Nothing
        prevTransparent = True
        If IsNothing(prevCellBitmap) = False Then
            prevCellBitmap.Dispose()
            prevCellBitmap = Nothing
        End If
    End Sub
#End Region

End Class
