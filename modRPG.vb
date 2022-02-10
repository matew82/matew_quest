Module modRPG
#Region "Hero"
    ''' <summary>
    ''' Восстанавливает характеристики героя/бойца на основании их парных свойств *Total
    ''' </summary>
    ''' <param name="hId">Id героя</param>
    ''' <returns>True если характеристики были восстановлены, False если они и так были полными</returns>
    Public Function HeroRestoreParams(ByVal hId As Integer) As String
        Dim classId As Integer = mScript.mainClassHash("H")
        Dim ch As SortedList(Of String, MatewScript.ChildPropertiesInfoType)
        If GVARS.G_ISBATTLE Then
            ch = mScript.Battle.Fighters(hId).heroProps
        Else
            ch = mScript.mainClass(classId).ChildProperties(hId)
        End If
        Dim arrs() As String = {hId.ToString}
        Dim wasChanged As Boolean = False
        Dim totalValue As String = "", usualValue As String = ""
        For pId As Integer = 0 To ch.Count - 1
            Dim pName As String = ch.ElementAt(pId).Key
            Dim pNameTotal As String = pName & "Total"
            Dim pTotal As MatewScript.ChildPropertiesInfoType = Nothing
            If ch.TryGetValue(pNameTotal, pTotal) = False Then Continue For

            If GVARS.G_ISBATTLE Then
                If mScript.Battle.ReadFighterProperty(pNameTotal, hId, totalValue, arrs) = False Then Return "#Error"
                If mScript.Battle.ReadFighterProperty(pName, hId, usualValue, arrs) = False Then Return "#Error"
            Else
                If ReadProperty(classId, pNameTotal, hId, -1, totalValue, arrs) = False Then Return "#Error"
                If ReadProperty(classId, pName, hId, -1, usualValue, arrs) = False Then Return "#Error"
            End If

            If totalValue = usualValue Then Continue For

            'событие отслеживания свойства
            If PropertiesRouter(classId, pName, {hId.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, totalValue) = "#Error" Then Return "#Error"
            'Dim trackResult As String = mScript.trackingProperties.RunBefore(classId, pName, arrs, totalValue)
            'If trackResult <> "False" AndAlso trackResult <> "#Error" Then
            '    SetPropertyValue(classId, pName, totalValue, hId)
            '    wasChanged = True
            'End If
        Next

        Return wasChanged.ToString
    End Function

    ''' <summary>
    ''' Устанавливаем параметры персонажа/бойца из переменной-массива
    ''' </summary>
    ''' <param name="hId">Id героя/бойца</param>
    ''' <param name="var">переменная</param>
    Public Function HeroSetParamsFromVariable(ByVal hId As Integer, ByRef var As cVariable.variableEditorInfoType) As String
        If IsNothing(var.lstSingatures) OrElse var.lstSingatures.Count = 0 Then Return ""
        Dim classId As Integer = mScript.mainClassHash("H")
        Dim arrs() As String = {hId.ToString}
        For i As Integer = 0 To var.lstSingatures.Count - 1
            Dim pName As String = var.lstSingatures.ElementAt(i).Key
            If mScript.mainClass(classId).Properties.ContainsKey(pName) = False Then Continue For
            'устанавливаем свойство из массива
            Dim newValue As String = var.arrValues(var.lstSingatures(pName))

            If pName = "Life" AndAlso Val(newValue) <= 0 Then Return _ERROR("Нельзя установить значение жизни меньшим или равным 0.", "SetParams")
            If PropertiesRouter(classId, pName, {hId.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, newValue) = "#Error" Then Return "#Error"

            ''событие отслеживания свойства
            'Dim trackResult As String = mScript.trackingProperties.RunBefore(classId, pName, arrs, newValue)
            'If trackResult <> "False" AndAlso trackResult <> "#Error" Then
            '    SetPropertyValue(classId, pName, newValue, hId)
            'End If
        Next

        Return ""
    End Function

    ''' <summary>
    ''' Возвращает список с ключом Id героя, значением Id его набора способностей. В списке оказываются все герои, имеющие доступные и непустые наборы способностей. Если нет ни одного - пустой список
    ''' </summary>
    Public Function HeroGetAbilitySetsIdList() As SortedList(Of Integer, Integer)
        Dim lst As New SortedList(Of Integer, Integer)

        Dim classH As Integer = mScript.mainClassHash("H"), classAb As Integer = mScript.mainClassHash("Ab"), abSet As String = ""
        If GVARS.G_ISBATTLE Then
            If mScript.Battle.Fighters.Count = 0 Then Return lst 'нет ни одного бойца
        Else
            If IsNothing(mScript.mainClass(classH).ChildProperties) OrElse mScript.mainClass(classH).ChildProperties.Count = 0 Then Return lst 'нет ни одного персонажа
        End If
        Dim isEnabled As Boolean = True
        If ReadPropertyBool(classAb, "Enabled", -1, -1, isEnabled, Nothing) = False Then Return lst

        If Not isEnabled Then Return lst 'все способности отключены
        Dim forEnd As Integer
        If GVARS.G_ISBATTLE Then
            forEnd = mScript.Battle.Fighters.Count - 1
        Else
            forEnd = mScript.mainClass(classH).ChildProperties.Count - 1
        End If
        For heroId As Integer = 0 To forEnd
            'получаем Id набора способностей персонажа
            If GVARS.G_ISBATTLE Then
                If mScript.Battle.ReadFighterProperty("AbilitiesSet", heroId, abSet, {heroId.ToString}) = False Then Return lst
            Else
                If ReadProperty(classH, "AbilitiesSet", heroId, -1, abSet, {heroId.ToString}) = False Then Return lst
            End If
            If abSet = "''" Then Continue For
            Dim abSetId As Integer = GetSecondChildIdByName(abSet, mScript.mainClass(classAb).ChildProperties)
            If abSetId <= 0 Then Continue For
            'не является ли набор пустым
            If IsNothing(mScript.mainClass(classAb).ChildProperties(abSetId)("Name").ThirdLevelProperties) OrElse mScript.mainClass(classAb).ChildProperties(abSetId)("Name").ThirdLevelProperties.Count = 0 Then Continue For
            If ReadPropertyBool(classAb, "Enabled", abSetId, -1, isEnabled, {abSet}) = False Then Continue For
            If Not isEnabled Then Continue For 'набор недоступен

            'набор есть и он доступен и непустой
            lst.Add(heroId, abSetId)
        Next heroId

        Return lst
    End Function
#End Region

#Region "Magics"
    ''' <summary>
    ''' Создает книгу заклинаний
    ''' </summary>
    ''' <param name="heroId">Id персонажа/бойца-хозяина книги</param>
    ''' <param name="mBookId">Id его книги магий</param>
    ''' <param name="arrParams">параметы скрипта</param>
    ''' <param name="curMagicId">Id магии, которая должна быть видна (страницу с которой надо открыть)</param>
    Public Function CreateMagicBook(ByVal heroId As Integer, ByVal mBookId As Integer, ByRef arrParams() As String, Optional ByVal curMagicId As Integer = -1) As String
        Dim classMg As Integer = mScript.mainClassHash("Mg")
        Dim classMW As Integer = mScript.mainClassHash("MgW")
        Dim classH As Integer = mScript.mainClassHash("H")

        Dim isNewWindow As Boolean = False
        Dim slot As Integer = 0
        If ReadPropertyInt(classMW, "Slot", -1, -1, slot, arrParams) = False Then Return "#Error"
        If slot = 5 Then isNewWindow = True 'в отдельном окне

        If isNewWindow Then
            'устанавливаем ширину окна
            Dim wndWidth As String = "", wndHeight As String = ""
            If ReadProperty(classMW, "Width", -1, -1, wndWidth, arrParams) = False Then Return "#Error"
            If ReadProperty(classMW, "Height", -1, -1, wndHeight, arrParams) = False Then Return "#Error"
            wndWidth = UnWrapString(wndWidth)
            wndHeight = UnWrapString(wndHeight)
            Dim wSize As New Size
            'width
            If IsNumeric(wndWidth) Then
                wSize.Width = Val(wndWidth) 'px
            ElseIf wndWidth.EndsWith("%") Then
                wSize.Width = Math.Round(My.Computer.Screen.WorkingArea.Width * Val(wndWidth) / 100) '%
            Else
                Return _ERROR("Нераспознан формат ширины окна магий (свойство MW.Width).", "MagicBook Builder")
            End If
            'height
            If IsNumeric(wndHeight) Then
                wSize.Height = Val(wndHeight) 'px
            ElseIf wndHeight.EndsWith("%") Then
                wSize.Height = Math.Round(My.Computer.Screen.WorkingArea.Height * Val(wndHeight) / 100) '%
            Else
                Return _ERROR("Нераспознан формат высоты окна магий (свойство MW.Height).", "MagicBook Builder")
            End If
            frmMagic.Size = wSize
        End If

        'получаем конвас окна магий и html-элемент magicBook
        Dim hDoc As HtmlDocument = questEnvironment.wbMagic.Document
        If IsNothing(hDoc) Then Return _ERROR("Html-документ книги магий не загружен.", "MagicBook Builder")
        Dim mgConvas As HtmlElement = Nothing
        mgConvas = hDoc.GetElementById("MagicConvas")
        If IsNothing(mgConvas) Then Return _ERROR("Ошибка в структуре файла " + questEnvironment.QuestPath + "\Magic.html. Файл в обязательном порядке должен иметь html-элемент с id = 'MapConvas'.", "MagicBook Builder")
        hDoc.InvokeScript("removeMagicBook", Nothing)

        Dim hMagicBook As HtmlElement = hDoc.GetElementById("magicBook")
        If IsNothing(hMagicBook) Then Return _ERROR("Ошибка в структуре файла " + questEnvironment.QuestPath + "\Magic.html. Файл в обязательном порядке должен иметь html-элемент с id = 'magicBook'.", "MagicBook Builder")

        ''устанавливаем css
        'Dim strCSS As String = ""
        'If ReadProperty(classMap, "CSS", mapId, -1, strCSS, arrParams) = False Then Return "#Error"
        'strCSS = UnWrapString(strCSS)
        'HtmlChangeFirstCSSLink(hDoc, strCSS)

        'BackColor
        Dim strStyle As New System.Text.StringBuilder
        Dim bkColor As String = ""
        If ReadProperty(classMg, "BackColor", mBookId, -1, bkColor, arrParams) = False Then Return "#Error"
        bkColor = UnWrapString(bkColor)

        If String.IsNullOrEmpty(bkColor) = False Then strStyle.Append(" background-color:" + bkColor + ";")
        'BackPicture
        Dim bkPicture As String = ""
        If ReadProperty(classMg, "BackPicture", mBookId, -1, bkPicture, arrParams) = False Then Return "#Error"
        bkPicture = UnWrapString(bkPicture)
        If bkPicture.Length > 0 Then
            strStyle.AppendFormat("background-image: url({0});", "'" + bkPicture.Replace("\", "/") + "'")
            Dim bkPicStyle As Integer = 0
            If ReadPropertyInt(classMg, "BackPicStyle", mBookId, -1, bkPicStyle, arrParams) = False Then Return "#Error"
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
                If ReadPropertyInt(classMg, "BackPicPos", mBookId, -1, bkPicPos, arrParams) = False Then Return "#Error"
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
        If strStyle.Length > 0 Then
            'устанавливаем высоту конваса для нормального отображения заднего фона
            strStyle.Append("min-height:" & questEnvironment.wbMagic.ClientSize.Height.ToString & "px;")
        End If
        mgConvas.Style = strStyle.ToString
        'hDoc.Body.Style = strStyle.ToString

        'получаем размеры книги заклинаний
        Dim mbWidth As Integer = 0, mbHeight As Integer = 0
        If ReadPropertyInt(classMg, "MagicBookWidth", mBookId, -1, mbWidth, arrParams) = False Then Return "#Error"

        '''''Загрузка изображения книги магий
        Dim hBookConvas As HtmlElement = hDoc.GetElementById("MagicBookConvas")
        If IsNothing(hBookConvas) Then Return _ERROR("Ошибка в структуре файла " + questEnvironment.QuestPath + "\Magic.html. Файл в обязательном порядке должен иметь html-элемент с id = 'MagicBookConvas'.", "MagicBook Builder")
        If ReadProperty(classMg, "MagicBookPicture", mBookId, -1, bkPicture, arrParams) = False Then Return "#Error"
        bkPicture = UnWrapString(bkPicture)
        If String.IsNullOrEmpty(bkPicture) Then Return _ERROR("Не указано изображение книги магий.", "MagicBook Builder")
        bkPicture = bkPicture.Replace("\", "/")
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, bkPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then Return _ERROR("Не найден файл изображения книги магий" & fPath & ".", "MagicBook Builder")
        Dim mbCustomSize As Boolean = True
        'получаем реальные размеры изображения 
        Dim bm As Bitmap = Image.FromFile(fPath)
        If mbWidth < 1 Then
            'реальный размер
            mbCustomSize = False
            mbWidth = bm.Width
            mbHeight = bm.Height
        Else
            'размер масштабирован - вычисляем высоту
            mbHeight = Math.Floor(bm.Height * (mbWidth / bm.Width))
        End If
        bm.Dispose()
        hBookConvas.Style = String.Format("width:{0}px;height:{1}px;background-repeat:no-repeat;background-size:contain;background-image: url({2});", mbWidth.ToString, mbHeight.ToString, "'" + bkPicture + "'")


        'Устанавливаем расположение книги магий
        Dim offsetX As String = "", offsetY As String = ""
        If ReadProperty(classMg, "OffsetByX", mBookId, -1, offsetX, arrParams) = False Then Return "#Error"
        If ReadProperty(classMg, "OffsetByY", mBookId, -1, offsetY, arrParams) = False Then Return "#Error"
        offsetX = UnWrapString(offsetX)
        offsetY = UnWrapString(offsetY)
        If String.IsNullOrEmpty(offsetX) = False Then
            If IsNumeric(offsetX) Then offsetX &= "px"
        End If
        If String.IsNullOrEmpty(offsetY) = False Then
            If IsNumeric(offsetY) Then offsetY &= "px"
        End If

        If String.IsNullOrEmpty(offsetX) Then
            offsetX = (Math.Floor((questEnvironment.wbMagic.ClientSize.Width - mbWidth) / 2)).ToString & "px"
        End If
        If String.IsNullOrEmpty(offsetY) Then
            offsetY = (Math.Floor((questEnvironment.wbMagic.ClientSize.Height - mbHeight) / 2)).ToString & "px"
        End If
        hBookConvas.Style &= "left:" & offsetX & ";top:" & offsetY & ";"

        'получаем изображения страниц книги магий
        Dim picLeftPage As String = "", picRightPage As String = ""
        'левой страницы
        If ReadProperty(classMg, "MagicBookLeftPage", mBookId, -1, picLeftPage, arrParams) = False Then Return "#Error"
        picLeftPage = UnWrapString(picLeftPage)
        If String.IsNullOrEmpty(picLeftPage) Then Return _ERROR("Не указано изображение левой страницы книги магий.", "MagicBook Builder")
        picLeftPage = picLeftPage.Replace("\", "/")
        fPath = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picLeftPage)
        If FileIO.FileSystem.FileExists(fPath) = False Then Return _ERROR("Не найден файл изображения левой страницы книги магий" & fPath & ".", "MagicBook Builder")
        'If mbCustomSize Then
        '    picLeftPage = String.Format("width:{0}px;height:{1}px;background-repeat:no-repeat;background-size:contain;background-image: url({2});", Math.Floor(mbWidth / 2).ToString, Math.Floor(mbHeight / 2).ToString, "'" + picLeftPage + "'")
        'Else
        '    picLeftPage = String.Format("background-image: url({0});", "'" + picLeftPage + "'")
        'End If
        picLeftPage = String.Format("width:{0}px;height:{1}px;background-repeat:no-repeat;background-size:contain;background-image: url({2});", Math.Floor(mbWidth / 2).ToString, Math.Floor(mbHeight / 2).ToString, "'" + picLeftPage + "'")
        'правой страницы
        If ReadProperty(classMg, "MagicBookRightPage", mBookId, -1, picRightPage, arrParams) = False Then Return "#Error"
        picRightPage = UnWrapString(picRightPage)
        If String.IsNullOrEmpty(picRightPage) Then Return _ERROR("Не указано изображение правой страницы книги магий.", "MagicBook Builder")
        picRightPage = picRightPage.Replace("\", "/")
        fPath = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, picRightPage)
        If FileIO.FileSystem.FileExists(fPath) = False Then Return _ERROR("Не найден файл изображения правой страницы книги магий" & fPath & ".", "MagicBook Builder")
        'If mbCustomSize Then
        '    picRightPage = String.Format("width:{0}px;height:{1}px;background-repeat:no-repeat;background-size:contain;background-image: url({2});", Math.Floor(mbWidth / 2).ToString, Math.Floor(mbHeight / 2).ToString, "'" + picRightPage + "'")
        'Else
        '    picRightPage = String.Format("background-image: url({0});", "'" + picRightPage + "'")
        'End If
        picRightPage = String.Format("width:{0}px;height:{1}px;background-repeat:no-repeat;background-size:contain;background-image: url({2});", Math.Floor(mbWidth / 2).ToString, Math.Floor(mbHeight / 2).ToString, "'" + picRightPage + "'")

        'кнопка скрытия книги
        Dim showCloseButton As Boolean = True
        If ReadPropertyBool(classMg, "UseCancelButton", mBookId, -1, showCloseButton, arrParams) = False Then Return "#Error"

        If showCloseButton Then
            Dim hClose As HtmlElement = hDoc.CreateElement("IMG")
            hClose.Id = "MagicBookClose"
            hClose.SetAttribute("src", "img/Default/close.png")
            AddHandler hClose.Click, Sub(sender As Object, e As HtmlElementEventArgs)
                                         sender.SetAttribute("src", "img/Default/close.png")

                                         isNewWindow = False
                                         slot = 0
                                         If ReadPropertyInt(classMW, "Slot", -1, -1, slot, Nothing) = False Then Return
                                         If slot = 5 Then isNewWindow = True 'в отдельном окне

                                         If isNewWindow Then
                                             frmMagic.Hide()
                                         Else
                                             questEnvironment.wbMagic.Hide()
                                         End If
                                     End Sub
            AddHandler hClose.MouseOver, Sub(sender As Object, e As HtmlElementEventArgs)
                                             sender.SetAttribute("src", "img/Default/closeHover.png")
                                         End Sub
            AddHandler hClose.MouseLeave, Sub(sender As Object, e As HtmlElementEventArgs)
                                              sender.SetAttribute("src", "img/Default/close.png")
                                          End Sub
            mgConvas.AppendChild(hClose)
        End If

        '''''''Создание книги заклинаний
        Dim pageNum As Integer = -1, page As Integer = 1
        Dim hPage As HtmlElement = Nothing
        Dim hMPage As HtmlElement = Nothing
        Dim hTable As HtmlElement = Nothing
        Dim maxPageHeight As Integer = 0
        If ReadPropertyInt(classMg, "VerticalPaddings", mBookId, -1, maxPageHeight, arrParams) = False Then Return "#Error"
        maxPageHeight = mbHeight - maxPageHeight 'максимальная высота содержимого страницы

        If mBookId > -1 AndAlso IsNothing(mScript.mainClass(classMg).ChildProperties(mBookId)("Name").ThirdLevelProperties) = False Then
            'есть как минимум одна магия в книге заклинаний
            Dim hTBody As HtmlElement = Nothing
            Dim arrs() As String, mBookStr As String = mBookId.ToString
            Dim isFirstMagicAtThisPage = True
            For mgId As Integer = 0 To mScript.mainClass(classMg).ChildProperties(mBookId)("Name").ThirdLevelProperties.Count - 1
                'перебираем все магии указанной книги по очереди
                If pageNum = -1 Then
                    'создаем первую страницу
                    pageNum += 1
                    hPage = hDoc.CreateElement("DIV")
                    hPage.SetAttribute("ignore", "1") 'первая страница должна идти с этим атрибутом
                    hMagicBook.AppendChild(hPage)

                    hMPage = hDoc.CreateElement("DIV")
                    'hMPage.SetAttribute("ClassName", "magicsPageFirst")
                    hPage.AppendChild(hMPage)

                    'hPage.SetAttribute("ClassName", "SheetLeft")
                    'hMPage.SetAttribute("ClassName", "magicsPageLeft")
                    hPage.SetAttribute("ClassName", "")
                    hMPage.SetAttribute("ClassName", "magicsPageFirst")

                    hTable = hDoc.CreateElement("TABLE")
                    hTable.SetAttribute("ClassName", "magicInfo")
                    hTable.Style = "height:" & maxPageHeight.ToString & "px;"
                    hTBody = hDoc.CreateElement("TBODY")
                    hTable.AppendChild(hTBody)
                    hMPage.AppendChild(hTable)
                End If

                arrs = {mBookStr, mgId.ToString}
                'получаем видимость
                Dim isVisible As Boolean = True
                If ReadPropertyBool(classMg, "Visible", mBookId, mgId, isVisible, arrs) = False Then Return "#Error"
                If Not isVisible Then Continue For 'магия не видна - пропускаем
                'получаем тип - 0 походная, 1 боевая, 2 обоюдная
                Dim spellType As Integer = 0
                If ReadPropertyInt(classMg, "SpellType", mBookId, mgId, spellType, arrs) = False Then Return "#Error"
                If spellType = 0 Then
                    'походная магия
                    If GVARS.G_ISBATTLE Then Continue For 'пропускаем эту магию если идет бой
                ElseIf spellType = 1 Then
                    'боевая магия
                    If Not GVARS.G_ISBATTLE Then Continue For 'пропускаем эту магию если бой не идет
                End If

                'получаем изображение магии и его размеры
                Dim imgPath As String = "", picWidth As String = "", picHeight As String = ""
                If ReadProperty(classMg, "Picture", mBookId, mgId, imgPath, arrs) = False Then Return "#Error"
                If ReadProperty(classMg, "picWidth", mBookId, mgId, picWidth, arrs) = False Then Return "#Error"
                If ReadProperty(classMg, "picHeight", mBookId, mgId, picHeight, arrs) = False Then Return "#Error"
                imgPath = UnWrapString(imgPath)
                picWidth = UnWrapString(picWidth)
                picHeight = UnWrapString(picHeight)
                If String.IsNullOrEmpty(picWidth) = False AndAlso IsNumeric(picWidth) Then picWidth &= "px"
                If String.IsNullOrEmpty(picHeight) = False AndAlso IsNumeric(picHeight) Then picHeight &= "px"


                Dim isChapter As Boolean = False
                If ReadPropertyBool(classMg, "Chapter", mBookId, mgId, isChapter, arrs) = False Then Return "#Error"

                Dim TR As HtmlElement = hDoc.CreateElement("TR") 'ряд с информацией о данной магии
                TR.SetAttribute("magicId", mgId.ToString) '!!!!!
                TR.SetAttribute("magicBookId", mBookId.ToString)
                TR.SetAttribute("heroId", heroId.ToString)
                Dim TD As HtmlElement = hDoc.CreateElement("TD")
                Dim IMG As HtmlElement = hDoc.CreateElement("IMG")
                If String.IsNullOrEmpty(imgPath) = False Then
                    IMG.SetAttribute("src", imgPath.Replace("\", "/"))
                    strStyle.Clear()
                    If String.IsNullOrEmpty(picWidth) = False Then strStyle.Append("width:" & picWidth & ";")
                    If String.IsNullOrEmpty(picHeight) = False Then strStyle.Append("height:" & picHeight & ";")
                    If strStyle.Length > 0 Then IMG.Style = strStyle.ToString
                End If
                TD.AppendChild(IMG)
                TR.AppendChild(TD)

                'получаем доступность
                Dim isEnabled As Boolean = True
                If ReadPropertyBool(classMg, "Enabled", -1, -1, isEnabled, arrs) = False Then Return "#Error"
                If isEnabled Then
                    If ReadPropertyBool(classMg, "Enabled", mBookId, -1, isEnabled, arrs) = False Then Return "#Error"
                    If isEnabled Then
                        If ReadPropertyBool(classMg, "Enabled", mBookId, mgId, isEnabled, arrs) = False Then Return "#Error"
                    End If
                End If

                'проверка по MaxSpellPerBattle SpelledThisBattle
                If GVARS.G_ISBATTLE Then
                    Dim spellMax As Integer = -1
                    If ReadPropertyInt(classMg, "MaxSpellPerBattle", mBookId, mgId, spellMax, arrs) = False Then Return "#Error"
                    If spellMax > 0 Then
                        Dim spelled As Integer = 0
                        If ReadPropertyInt(classMg, "SpelledThisBattle", mBookId, mgId, spelled, arrs) = False Then Return "#Error"
                        If spelled >= spellMax Then isEnabled = False
                    ElseIf spellMax = 0 Then
                        isEnabled = False
                    End If
                End If

                Dim eventId As Integer = 0
                If isEnabled Then
                    'MagicEnabledQuery
                    Dim eqArrs() As String = {heroId.ToString, mBookStr, mgId.ToString}
                    Dim res As String = ""
                    'глобальное
                    eventId = mScript.mainClass(classMg).Properties("MagicEnabledQuery").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, eqArrs, "MagicEnabledQuery", False)
                        If res = "#Error" Then
                            Return res
                        ElseIf res = "False" Then
                            isEnabled = False
                        End If
                    End If

                    'данной книги магий
                    If isEnabled Then
                        eventId = mScript.mainClass(classMg).ChildProperties(mBookId)("MagicEnabledQuery").eventId
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, eqArrs, "MagicEnabledQuery", False)
                            If res = "#Error" Then
                                Return res
                            ElseIf res = "False" Then
                                isEnabled = False
                            End If
                        End If

                        'данной магии
                        If isEnabled Then
                            eventId = mScript.mainClass(classMg).ChildProperties(mBookId)("MagicEnabledQuery").ThirdLevelEventId(mgId)
                            If eventId > 0 Then
                                res = mScript.eventRouter.RunEvent(eventId, eqArrs, "MagicEnabledQuery", False)
                                If res = "#Error" Then
                                    Return res
                                ElseIf res = "False" Then
                                    isEnabled = False
                                End If
                            End If
                        End If
                    End If
                End If

                TD = hDoc.CreateElement("TD")
                Dim hP As HtmlElement = hDoc.CreateElement("P")
                Dim strCaption As String = ""
                If ReadProperty(classMg, "Caption", mBookId, mgId, strCaption, arrs) = False Then Return "#Error"
                strCaption = UnWrapString(strCaption)
                hP.InnerHtml = strCaption
                TD.AppendChild(hP)

                Dim strDescription As String = ""
                If ReadProperty(classMg, "Description", mBookId, mgId, strDescription, arrs) = False Then Return "#Error"
                strDescription = UnWrapString(strDescription)

                If isChapter Then
                    'данный элемент - раздел магий
                    hP.SetAttribute("ClassName", "magicNameChapter")
                    TR.SetAttribute("ClassName", "magicChapter")
                    hP.SetAttribute("Title", strDescription)
                Else
                    'данный элемент - обычная магия
                    hP.SetAttribute("ClassName", "magicName")

                    Dim strCost As String = "", cost As Double = 0
                    If ReadPropertyDbl(classMg, "Cost", mBookId, mgId, cost, arrs) = False Then Return "#Error"
                    If ReadProperty(classMg, "CostString", mBookId, mgId, strCost, arrs) = False Then Return "#Error"
                    strCost = UnWrapString(strCost)
                    If String.IsNullOrEmpty(strCost) Then
                        If cost = 0 Then
                            strCost = "Бесплатно"
                        Else
                            strCost = "Стоит " & cost.ToString & " единиц маны"
                        End If
                    End If

                    If isEnabled Then
                        'доступно ли заклинание исходя из цены
                        If cost > 0 Then
                            'указана стоимость в мане
                            'получаем количество маны у персонажа
                            Dim mana As Double = 0
                            If GVARS.G_ISBATTLE Then
                                If mScript.Battle.ReadFighterPropertyDbl("Mana", heroId, mana, {heroId.ToString}) = False Then Return "#Error"
                            Else
                                If ReadPropertyDbl(classH, "Mana", heroId, -1, mana, {heroId.ToString}) = False Then Return "#Error"
                            End If
                            If cost > mana Then isEnabled = False 'маны не хватает
                        End If
                    End If

                    'получаем класс магии
                    If isEnabled Then
                        Dim mgClass As String = ""
                        If ReadProperty(classMg, "MagicClass", mBookId, mgId, mgClass, arrs) = False Then Return "#Error"
                        mgClass = UnWrapString(mgClass)
                        TR.SetAttribute("ClassName", mgClass)
                    End If

                    hP = hDoc.CreateElement("P")
                    hP.SetAttribute("ClassName", "magicDescription")
                    hP.InnerHtml = strDescription
                    TD.AppendChild(hP)

                    hP = hDoc.CreateElement("P")
                    hP.SetAttribute("ClassName", "magicCost")
                    hP.InnerHtml = strCost
                    TD.AppendChild(hP)
                End If

                If Not isEnabled Then
                    TR.SetAttribute("disabled", "True")
                End If

                TR.AppendChild(TD)
                hTBody.AppendChild(TR)
                'Dim msImg As mshtml.IHTMLImgElement = IMG.DomElement
                'Do While Not msImg.complete
                '    Application.DoEvents()
                'Loop

                Dim pageHeight As Integer = hTable.OffsetRectangle.Height
                Dim fromNewPage As Boolean = False
                If isFirstMagicAtThisPage = False Then
                    If ReadProperty(classMg, "FromNewPage", mBookId, mgId, fromNewPage, arrs) = False Then Return "#Error"
                End If

                If pageHeight > maxPageHeight OrElse fromNewPage Then
                    'If prevPageNum = 0 Then
                    '    'подправляем первую страницу (так надо, чтобы нормально считалась ее высота)
                    '    hPage.SetAttribute("ClassName", "")
                    '    hMPage.SetAttribute("ClassName", "magicsPageFirst")
                    'End If

                    'создание новой страницы
                    Dim isRightPage As Boolean = False
                    pageNum += 1
                    If Math.Floor(pageNum / 2) <> Math.Ceiling(pageNum / 2) Then isRightPage = True

                    hPage = hDoc.CreateElement("DIV")
                    hMagicBook.AppendChild(hPage)

                    hMPage = hDoc.CreateElement("DIV")
                    hPage.AppendChild(hMPage)

                    If isRightPage Then
                        hPage.SetAttribute("ClassName", "SheetRight")
                        hPage.Style = picRightPage ' String.Format("background-image: url({0});", "'" + picRightPage + "'")
                        hMPage.SetAttribute("ClassName", "magicsPageRight")
                    Else
                        hPage.SetAttribute("ClassName", "SheetLeft")
                        hPage.Style = picLeftPage ' String.Format("background-image: url({0});", "'" + picLeftPage + "'")
                        hMPage.SetAttribute("ClassName", "magicsPageLeft")
                    End If

                    hTable = hDoc.CreateElement("TABLE")
                    hTable.SetAttribute("ClassName", "magicInfo")
                    hTBody = hDoc.CreateElement("TBODY")
                    hTable.AppendChild(hTBody)
                    hMPage.AppendChild(hTable)
                    hTable.Style = "height:" & maxPageHeight.ToString & "px;"

                    'удаляем из старого места
                    Dim msTR As mshtml.IHTMLDOMNode = TR.DomElement
                    msTR.removeNode(True)

                    'добавляем в новое место
                    hTBody.AppendChild(TR)
                End If
                If curMagicId > -1 AndAlso mgId = curMagicId Then page = pageNum
                isFirstMagicAtThisPage = False

                'события каждой магии
                If isEnabled Then
                    AddHandler TR.MouseOver, AddressOf del_magic_MouseOver
                    AddHandler TR.MouseLeave, AddressOf del_magic_MouseLeave
                    AddHandler TR.Click, AddressOf del_magic_Click
                    AddHandler TR.MouseDown, AddressOf del_magic_MouseDown
                    AddHandler TR.MouseUp, AddressOf del_magic_MouseUp
                End If
            Next mgId
        End If
        If page = 0 Then pageNum = 1

        If pageNum < 1 Then
            'If prevPageNum = 0 Then
            '    'подправляем первую страницу (так надо, чтобы нормально считалась ее высота)
            '    hPage.SetAttribute("ClassName", "")
            '    hMPage.SetAttribute("ClassName", "magicsPageFirst")
            'End If

            'создаем хоть одну (пустую) страницу, если ее не существует
            hPage = hDoc.CreateElement("DIV")
            hPage.SetAttribute("ClassName", "SheetRight")
            hPage.Style = picRightPage ' String.Format("background-image: url({0});", "'" + picRightPage + "'")
            hMagicBook.AppendChild(hPage)

            hMPage = hDoc.CreateElement("DIV")
            hMPage.SetAttribute("ClassName", "magicsPageRight")
            hPage.AppendChild(hMPage)
        End If

        'Dim hh As HtmlElement = hMagicBook.Children(1).Children(0).Children(0)
        'MsgBox(hh.ClientRectangle.Height.ToString)
        'For i = 0 To 1000
        '    Application.DoEvents()
        'Next
        'MsgBox(hh.ClientRectangle.Height.ToString)
        Dim ObjArr(2) As Object
        ObjArr(0) = CObj(mbWidth)
        ObjArr(1) = CObj(mbHeight)
        ObjArr(2) = CObj(page)
        'ObjArr(1) = CObj(New String(mbHeight.ToString))
        hDoc.InvokeScript("createMagicBook", ObjArr)

        Return ""
    End Function

    ''' <summary>
    ''' Проверяет есть ли у героя/бойца хоть одна магия, которую он может использовать в данный момент
    ''' </summary>
    ''' <param name="heroId">Id героя</param>
    Public Function HasEnabledMagics(ByVal heroId As Integer) As Boolean
        If GVARS.G_ISBATTLE AndAlso GVARS.G_CANUSEMAGIC = False Then Return False
        'есть ли у героя книга магий?
        Dim classH As Integer = mScript.mainClassHash("H"), classMg As Integer = mScript.mainClassHash("Mg")
        Dim mBook As String = "", mBookId As Integer = -1
        If GVARS.G_ISBATTLE Then
            If mScript.Battle.ReadFighterProperty("MagicBook", heroId, mBook, {heroId.ToString}) = False Then Return False
        Else
            If ReadProperty(classH, "MagicBook", heroId, -1, mBook, {heroId.ToString}) = False Then Return False
        End If
        mBookId = GetSecondChildIdByName(mBook, mScript.mainClass(classMg).ChildProperties)
        If mBookId < 0 Then Return False

        'есть ли в ней хоть одна магия?
        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classMg).ChildProperties(mBookId)("Name")
        If IsNothing(ch.ThirdLevelProperties) OrElse ch.ThirdLevelProperties.Count = 0 Then Return False

        'проверяем доступность кажной магии
        For mgId As Integer = 0 To ch.ThirdLevelProperties.Count - 1
            If CanUseMagic(heroId, mBookId, mgId) Then Return True
        Next mgId

        Return False
    End Function

    ''' <summary>
    ''' Возвращает может ли указанный персонаж использовать указанную магию
    ''' </summary>
    ''' <param name="heroId">Id персонажа/бойца</param>
    ''' <param name="mBookId">Id его книги заклинаний</param>
    ''' <param name="mgId">Id магии</param>
    Public Function CanUseMagic(ByVal heroId As Integer, ByVal mBookId As Integer, mgId As Integer) As Boolean
        'G_CANUSEMAGIC
        If GVARS.G_ISBATTLE AndAlso GVARS.G_CANUSEMAGIC = False Then Return False

        '1. Visible
        Dim arrs() As String = {mBookId.ToString, mgId.ToString}
        Dim classMg As Integer = mScript.mainClassHash("Mg")
        'получаем видимость
        Dim isVisible As Boolean = True
        If ReadPropertyBool(classMg, "Visible", mBookId, mgId, isVisible, arrs) = False Then Return "#Error"
        If Not isVisible Then Return False 'магия не видна

        '2. Enabled
        'получаем доступность
        Dim isEnabled As Boolean = True
        If ReadPropertyBool(classMg, "Enabled", -1, -1, isEnabled, arrs) = False Then Return "#Error"
        If isEnabled Then
            If ReadPropertyBool(classMg, "Enabled", mBookId, -1, isEnabled, arrs) = False Then Return "#Error"
            If isEnabled Then
                If ReadPropertyBool(classMg, "Enabled", mBookId, mgId, isEnabled, arrs) = False Then Return "#Error"
            End If
        End If
        If isEnabled = False Then Return False

        '3. MaxSpellPerBattle SpelledThisBattle
        'проверка по MaxSpellPerBattle SpelledThisBattle
        If GVARS.G_ISBATTLE Then
            Dim spellMax As Integer = -1
            If ReadPropertyInt(classMg, "MaxSpellPerBattle", mBookId, mgId, spellMax, arrs) = False Then Return "#Error"
            If spellMax > 0 Then
                Dim spelled As Integer = 0
                If ReadPropertyInt(classMg, "SpelledThisBattle", mBookId, mgId, spelled, arrs) = False Then Return "#Error"
                If spelled >= spellMax Then Return False
            ElseIf spellMax = 0 Then
                Return False
            End If
        End If

        '4. MagicEnabledQuery
        Dim eventId As Integer = 0
        'MagicEnabledQuery
        Dim eqArrs() As String = {heroId.ToString, mBookId.ToString, mgId.ToString}
        Dim res As String = ""
        'глобальное
        eventId = mScript.mainClass(classMg).Properties("MagicEnabledQuery").eventId
        If eventId > 0 Then
            res = mScript.eventRouter.RunEvent(eventId, eqArrs, "MagicEnabledQuery", False)
            If res = "False" OrElse res = "#Error" Then Return False
        End If

        'данной книги магий
        eventId = mScript.mainClass(classMg).ChildProperties(mBookId)("MagicEnabledQuery").eventId
        If eventId > 0 Then
            res = mScript.eventRouter.RunEvent(eventId, eqArrs, "MagicEnabledQuery", False)
            If res = "False" OrElse res = "#Error" Then Return False
        End If

        'данной магии
        eventId = mScript.mainClass(classMg).ChildProperties(mBookId)("MagicEnabledQuery").ThirdLevelEventId(mgId)
        If eventId > 0 Then
            res = mScript.eventRouter.RunEvent(eventId, eqArrs, "MagicEnabledQuery", False)
            If res = "False" OrElse res = "#Error" Then Return False
        End If

        '5. SpellType
        'получаем тип - 0 походная, 1 боевая, 2 обоюдная
        Dim spellType As Integer = 0
        If ReadPropertyInt(classMg, "SpellType", mBookId, mgId, spellType, arrs) = False Then Return "#Error"
        If spellType = 0 Then
            'походная магия
            If GVARS.G_ISBATTLE Then Return False 'пропускаем эту магию если идет бой
        ElseIf spellType = 1 Then
            'боевая магия
            If Not GVARS.G_ISBATTLE Then Return False 'пропускаем эту магию если бой не идет
        End If

        'Mana Cost
        'доступно ли заклинание исходя из цены
        Dim cost As Double = 0
        If ReadPropertyDbl(classMg, "Cost", mBookId, mgId, cost, arrs) = False Then Return "#Error"
        If cost > 0 Then
            'указана стоимость в мане
            'получаем количество маны у персонажа
            Dim mana As Double = 0, classH As Integer = mScript.mainClassHash("H")
            If GVARS.G_ISBATTLE Then
                If mScript.Battle.ReadFighterPropertyDbl("Mana", heroId, mana, {heroId.ToString}) = False Then Return "#Error"
            Else
                If ReadPropertyDbl(classH, "Mana", heroId, -1, mana, {heroId.ToString}) = False Then Return "#Error"
            End If
            If cost > mana Then Return False 'маны не хватает
        End If

        Return True
    End Function

    Private Sub del_magic_MouseOver(sender As HtmlElement, e As HtmlElementEventArgs)
        'sender - TR с атрибутами magicId, magicBookId, heroId
        'Выводим картинку магии при наведении PictureHover
        Dim classMg As Integer = mScript.mainClassHash("Mg")
        Dim mBookId As Integer = Val(sender.GetAttribute("magicBookId"))
        Dim mgId As Integer = Val(sender.GetAttribute("magicId"))

        Dim arrs() As String = {mBookId.ToString, mgId.ToString}
        'Dim enabled As Boolean = True
        'ReadPropertyBool(classMg, "Enabled", mBookId, mgId, enabled, arrs)
        'If enabled = False Then Return

        Dim res As String = ""
        Dim wasEvent As Boolean = False

        'получаем PictureHover. Если пусто, то выход
        Dim mgPictureHover As String = ""
        If ReadProperty(classMg, "PictureHover", mBookId, mgId, mgPictureHover, arrs) = False Then Return
        mgPictureHover = UnWrapString(mgPictureHover)
        If String.IsNullOrEmpty(mgPictureHover) = False Then
            Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, mgPictureHover)
            If FileIO.FileSystem.FileExists(fPath) = False Then
                MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            'получаем Picture. PictureHover и Picture не должны быть равны
            Dim mgPicture As String = ""
            If ReadProperty(classMg, "Picture", mBookId, mgId, mgPicture, arrs) = False Then Return
            mgPicture = UnWrapString(mgPicture)

            'убеждаемся в правильном формате пути к картинкам
            mgPictureHover = mgPictureHover.Replace("\", "/")
            If String.IsNullOrEmpty(mgPicture) = False Then mgPicture = mgPicture.Replace("\", "/")
            If String.Compare(mgPicture, mgPictureHover, True) <> 0 Then
                'собственно замена картинки Picture на PictureHover
                Dim hEl As HtmlElement = sender.Children(0)
                If hEl.Children.Count > 0 AndAlso hEl.Children(0).TagName = "IMG" Then
                    hEl = hEl.Children(0) 'изображение магии в первой ячейке (левая страница)
                ElseIf sender.Children.Count > 1 Then
                    hEl = hEl.Children(1) 'изображение магии во второй ячейке (правая страница)
                    If hEl.Children.Count > 0 AndAlso hEl.Children(0).TagName = "IMG" Then
                        hEl = hEl.Children(0)
                    Else
                        Return
                    End If
                Else
                    Return
                End If

                hEl.SetAttribute("Src", mgPictureHover)
            End If
        End If
    End Sub

    Private Sub del_magic_MouseLeave(sender As HtmlElement, e As HtmlElementEventArgs)
        'sender - TR с атрибутами magicId, magicBookId, heroId
        'возвращаем картинку магии Picture после того, как она была зменена на PictureHover
        Dim classMg As Integer = mScript.mainClassHash("Mg")
        Dim mBookId As Integer = Val(sender.GetAttribute("magicBookId"))
        Dim mgId As Integer = Val(sender.GetAttribute("magicId"))

        Dim arrs() As String = {mBookId.ToString, mgId.ToString}
        'Dim enabled As Boolean = True
        'ReadPropertyBool(classMg, "Enabled", mBookId, mgId, enabled, arrs)
        'If enabled = False Then Return

        'получаем PictureHover. Если пусто, то выход
        Dim mgPictureHover As String = ""
        If ReadProperty(classMg, "PictureHover", mBookId, mgId, mgPictureHover, arrs) = False Then Return
        mgPictureHover = UnWrapString(mgPictureHover)
        If String.IsNullOrEmpty(mgPictureHover) = False Then
            Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, mgPictureHover)
            If FileIO.FileSystem.FileExists(fPath) = False Then
                Return
            End If

            'получаем Picture. PictureHover и Picture не должны быть равны
            Dim mgPicture As String = ""
            If ReadProperty(classMg, "Picture", mBookId, mgId, mgPicture, arrs) = False Then Return
            mgPicture = UnWrapString(mgPicture)

            'убеждаемся в правильном формате пути к картинкам
            mgPictureHover = mgPictureHover.Replace("\", "/")
            If String.IsNullOrEmpty(mgPicture) = False Then mgPicture = mgPicture.Replace("\", "/")
            If String.Compare(mgPicture, mgPictureHover, True) <> 0 Then
                'собственно замена картинки Picture на PictureHover
                '0 - обычный, 1 - контейнер
                'возможное расположение картинки: 
                '1) IMG 2) SPAN-NOBR-IMG 3) TABLE-TBODY-TR-TD
                Dim hEl As HtmlElement = sender.Children(0)
                If hEl.Children.Count > 0 AndAlso hEl.Children(0).TagName = "IMG" Then
                    hEl = hEl.Children(0) 'изображение магии в первой ячейке (левая страница)
                ElseIf sender.Children.Count > 1 Then
                    hEl = hEl.Children(1) 'изображение магии во второй ячейке (правая страница)
                    If hEl.Children.Count > 0 AndAlso hEl.Children(0).TagName = "IMG" Then
                        hEl = hEl.Children(0)
                    Else
                        Return
                    End If
                Else
                    Return
                End If
                hEl.SetAttribute("Src", mgPicture)
            End If
        End If
    End Sub

    Friend Sub del_magic_Click(sender As HtmlElement, e As HtmlElementEventArgs)
        'sender - TR с атрибутами magicId, magicBookId, heroId
        Dim mBookId As Integer = Val(sender.GetAttribute("magicBookId"))
        Dim mgId As Integer = Val(sender.GetAttribute("magicId"))
        Dim heroId As Integer = Val(sender.GetAttribute("heroId"))
        CastMagic(heroId, mBookId, mgId, True)
    End Sub

    ''' <summary>
    ''' Запускает магию. Если бой не идет, то событие MagicCastEvent запускается прямо здесь, если идет, то происходит выбор цели
    ''' </summary>
    ''' <param name="heroId">Id персонажа/бойца</param>
    ''' <param name="mBookId">Id его книги заклинаний</param>
    ''' <param name="mgId">Id магии</param>
    ''' <param name="selectEvent">Выполять ли MagicSelectEvent</param>
    ''' <returns>Было ли запущено MagicCastEvent (или дальнейшие события, связанные с выбором цели) или MagicSelectEvent вернуло False</returns>
    Public Function CastMagic(ByVal heroId As Integer, ByVal mBookId As Integer, ByVal mgId As Integer, ByVal selectEvent As Boolean) As Boolean
        Dim classMg As Integer = mScript.mainClassHash("Mg")
        Dim arrs() As String = {mBookId.ToString, mgId.ToString}

        'событие MagicSelectEvent
        Dim res As String = "", eventId As Integer
        Dim eArrs() As String = {heroId.ToString, mBookId.ToString, mgId.ToString}
        If selectEvent Then
            'глобальное
            eventId = mScript.mainClass(classMg).Properties("MagicSelectEvent").eventId
            If eventId > 0 Then
                res = mScript.eventRouter.RunEvent(eventId, eArrs, "MagicSelectEvent", False)
                If res = "#Error" Then Return False
            End If

            'данной книги магий
            If res <> "False" Then
                eventId = mScript.mainClass(classMg).ChildProperties(mBookId)("MagicSelectEvent").eventId
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, eArrs, "MagicSelectEvent", False)
                    If res = "#Error" Then Return False
                End If
            End If

            'данной магии
            If res <> "False" Then
                eventId = mScript.mainClass(classMg).ChildProperties(mBookId)("MagicSelectEvent").ThirdLevelEventId(mgId)
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, eArrs, "MagicSelectEvent", False)
                    If res = "#Error" Then Return False
                End If
            End If

            If res = "False" Then
                'событие ScriptFinishedEvent
                If EventGeneratedFromScript = False Then
                    If mScript.eventRouter.RunScriptFinishedEvent(Nothing) = "#Error" Then Return False
                End If
                EventGeneratedFromScript = False
                Return False
            End If
        End If

        'прячем книгу заклинаний
        Dim slot As Integer = 0, isNewWindow As Boolean = False
        Dim classMW As Integer = mScript.mainClassHash("MgW")
        If ReadPropertyInt(classMW, "Slot", -1, -1, slot, arrs) = False Then Return False
        If slot = 5 Then isNewWindow = True 'в отдельном окне

        If isNewWindow Then
            questEnvironment.wbMagic.Hide()
            frmMagic.Hide()
        Else
            questEnvironment.wbMagic.Hide()
        End If

        'Вычитаем стоимость заклинания из количества маны героя
        Dim classH As Integer = mScript.mainClassHash("H")
        Dim mana As Double = 0, cost As Double = 0
        If GVARS.G_ISBATTLE Then
            If mScript.Battle.ReadFighterProperty("Mana", heroId, mana, {heroId.ToString}) = False Then Return False
        Else
            If ReadPropertyDbl(classH, "Mana", heroId, -1, mana, {heroId.ToString}) = False Then Return False
        End If
        If ReadPropertyDbl(classMg, "Cost", mBookId, mgId, cost, arrs) = False Then Return False
        mana -= cost

        'событие отслеживания свойства Mana
        Dim manaStr As String = mana.ToString(provider_points)
        'Dim trackResult As String = mScript.trackingProperties.RunBefore(classH, "Mana", {heroId.ToString}, manaStr)
        'If trackResult <> "False" AndAlso trackResult <> "#Error" Then
        '    SetPropertyValue(classH, "Mana", manaStr, heroId)
        'End If
        If PropertiesRouter(classH, "Mana", {heroId.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, manaStr) = "#Error" Then Return False

        'SpelledThisBattle +1
        If GVARS.G_ISBATTLE Then
            Dim spelled As Integer = 0
            If ReadPropertyDbl(classMg, "SpelledThisBattle", mBookId, mgId, spelled, arrs) = False Then Return False
            spelled += 1

            'событие отслеживания свойства SpelledThisBattle
            Dim trackResult As String = mScript.trackingProperties.RunBefore(classMg, "SpelledThisBattle", {mBookId.ToString, mgId.ToString}, spelled.ToString)
            If trackResult <> "False" AndAlso trackResult <> "#Error" Then
                SetPropertyValue(classMg, "SpelledThisBattle", spelled.ToString, mBookId, mgId)
            End If
        End If

        If GVARS.G_ISBATTLE = False Then
            'выполняем магию если сейчас нет боя - MagicCastEvent
            eArrs = {heroId.ToString, "-1", "1", "0", mBookId.ToString, mgId.ToString}

            'сначала - звук
            Dim snd As String = ""
            If ReadProperty(classMg, "Sound", mBookId, mgId, snd, arrs) = False Then Return False
            snd = UnWrapString(snd)
            If String.IsNullOrEmpty(snd) = False Then
                If Say(snd, 100, False) = "#Error" Then Return False
            End If

            'глобальное
            eventId = mScript.mainClass(classMg).Properties("MagicCastEvent").eventId
            If eventId > 0 Then
                res = mScript.eventRouter.RunEvent(eventId, eArrs, "MagicCastEvent", False)
                If res = "#Error" Then Return False
            End If

            'данной книги магий
            eventId = mScript.mainClass(classMg).ChildProperties(mBookId)("MagicCastEvent").eventId
            If eventId > 0 Then
                res = mScript.eventRouter.RunEvent(eventId, eArrs, "MagicCastEvent", False)
                If res = "#Error" Then Return False
            End If

            'данной магии
            eventId = mScript.mainClass(classMg).ChildProperties(mBookId)("MagicCastEvent").ThirdLevelEventId(mgId)
            If eventId > 0 Then
                res = mScript.eventRouter.RunEvent(eventId, eArrs, "MagicCastEvent", False)
                If res = "#Error" Then Return False
            End If
        Else
            'если бой идет, то обработчик MagicCastEvent будет запущен из Battle.Hit
            If GVARS.G_CANUSEMAGIC Then
                GVARS.G_STRIKETYPE = 1
                GVARS.G_CURMAGIC = mgId
                If mScript.Battle.Hit(GVARS.G_CURFIGHTER, GVARS.G_CURVICTIM, GVARS.G_STRIKETYPE) = "#Error" Then Return False
            End If
        End If

        'событие ScriptFinishedEvent
        If EventGeneratedFromScript = False Then
            If mScript.eventRouter.RunScriptFinishedEvent(Nothing) = "#Error" Then Return False
        End If
        EventGeneratedFromScript = False
        Return True
    End Function

    Private Sub del_magic_MouseDown(sender As Object, e As HtmlElementEventArgs)
        'sender - TR с атрибутами magicId, magicBookId, heroId
        Dim classMg As Integer = mScript.mainClassHash("Mg")
        Dim mBookId As Integer = Val(sender.GetAttribute("magicBookId"))
        Dim mgId As Integer = Val(sender.GetAttribute("magicId"))

        Dim arrs() As String = {mBookId.ToString, mgId.ToString}
        'Dim enabled As Boolean = True
        'ReadPropertyBool(classMg, "Enabled", mBookId, mgId, enabled, arrs)
        'If enabled = False Then Return

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim mgPictureActive As String = ""
        If ReadProperty(classMg, "PictureActive", mBookId, mgId, mgPictureActive, arrs) = False Then Return
        mgPictureActive = UnWrapString(mgPictureActive)
        If String.IsNullOrEmpty(mgPictureActive) = False Then
            Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, mgPictureActive)
            If FileIO.FileSystem.FileExists(fPath) = False Then
                MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            'убеждаемся в правильном формате пути к картинкам
            mgPictureActive = mgPictureActive.Replace("\", "/")

            'собственно замена картинки PictureActive
            Dim hEl As HtmlElement = sender.Children(0)
            If hEl.Children.Count > 0 AndAlso hEl.Children(0).TagName = "IMG" Then
                hEl = hEl.Children(0) 'изображение магии в первой ячейке (левая страница)
            ElseIf sender.Children.Count > 1 Then
                hEl = hEl.Children(1) 'изображение магии во второй ячейке (правая страница)
                If hEl.Children.Count > 0 AndAlso hEl.Children(0).TagName = "IMG" Then
                    hEl = hEl.Children(0)
                Else
                    Return
                End If
            Else
                Return
            End If
            hEl.SetAttribute("Src", mgPictureActive)
        End If
    End Sub

    Private Sub del_magic_MouseUp(sender As Object, e As HtmlElementEventArgs)
        'sender - TR с атрибутами magicId, magicBookId, heroId
        Dim classMg As Integer = mScript.mainClassHash("Mg")
        Dim mBookId As Integer = Val(sender.GetAttribute("magicBookId"))
        Dim mgId As Integer = Val(sender.GetAttribute("magicId"))

        Dim arrs() As String = {mBookId.ToString, mgId.ToString}
        'Dim enabled As Boolean = True
        'ReadPropertyBool(classMg, "Enabled", mBookId, mgId, enabled, arrs)
        'If enabled = False Then Return

        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim mgPictureActive As String = ""
        If ReadProperty(classMg, "PictureActive", mBookId, mgId, mgPictureActive, arrs) = False Then Return
        mgPictureActive = UnWrapString(mgPictureActive)
        If String.IsNullOrEmpty(mgPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim mgPicture As String = ""
        If ReadProperty(classMg, "PictureHover", mBookId, mgId, mgPicture, arrs) = False Then Return
        mgPicture = UnWrapString(mgPicture)
        If String.IsNullOrEmpty(mgPicture) Then
            If ReadProperty(classMg, "Picture", mBookId, mgId, mgPicture, arrs) = False Then Return
            mgPicture = UnWrapString(mgPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, mgPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        mgPicture = mgPicture.Replace("\", "/")

        'собственно замена картинки PictureActive
        Dim hEl As HtmlElement = sender.Children(0)
        If hEl.Children.Count > 0 AndAlso hEl.Children(0).TagName = "IMG" Then
            hEl = hEl.Children(0) 'изображение магии в первой ячейке (левая страница)
        ElseIf sender.Children.Count > 1 Then
            hEl = hEl.Children(1) 'изображение магии во второй ячейке (правая страница)
            If hEl.Children.Count > 0 AndAlso hEl.Children(0).TagName = "IMG" Then
                hEl = hEl.Children(0)
            Else
                Return
            End If
        Else
            Return
        End If
        hEl.SetAttribute("Src", mgPicture)
    End Sub

#End Region

#Region "Abilities"
    ''' <summary>
    ''' Запускает события AbilityOnTimeEllapsedEvent
    ''' </summary>
    Public Function Abilitities_RaiseAbilityOnTimeEllapsedEvents() As String
        Dim abList As SortedList(Of Integer, Integer) = HeroGetAbilitySetsIdList() 'Key - heroId/fighterId, Value - AbSetId
        If abList.Count = 0 Then Return ""
        'событие AbilityOnTimeEllapsedEvent
        Dim classAb As Integer = mScript.mainClassHash("Ab")
        Dim globalEventId As Integer = mScript.mainClass(classAb).Properties("AbilityOnTimeEllapsedEvent").eventId

        For i As Integer = 0 To abList.Count - 1
            Dim heroId As Integer = abList.ElementAt(i).Key
            Dim abSetId As Integer = abList.ElementAt(i).Value
            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityOnTimeEllapsedEvent")
            Dim abSetEventId As Integer = ch.eventId
            For abId As Integer = ch.ThirdLevelProperties.Count - 1 To 0 Step -1
                'запускаем событие AbilityOnTimeEllapsedEvent для каждой способности из всех доступных наборов, прикрепленных к персонажам
                Dim arrs() As String = {abSetId.ToString, abId.ToString}
                Dim isEnabled As Boolean = True
                If ReadPropertyBool(classAb, "Enabled", abSetId, abId, isEnabled, arrs) = False Then Return "#Error"
                If Not isEnabled Then Continue For 'способность недоступна - пропускаем

                Dim appliedTime As Double = 0, lastRaisingTime As Double = 0, ellapsedType As Integer = 0, interval As Long = 0
                If ReadPropertyDbl(classAb, "TimeWhenApplied", abSetId, abId, appliedTime, arrs) = False Then Return "#Error"
                If ReadPropertyDbl(classAb, "TimeLastRaising", abSetId, abId, lastRaisingTime, arrs) = False Then Return "#Error"
                If lastRaisingTime = 0 Then lastRaisingTime = appliedTime
                If ReadPropertyDbl(classAb, "TimeEllapsedType", abSetId, abId, ellapsedType, arrs) = False Then Return "#Error"
                Dim timeActive As Long = 0
                If ReadPropertyDbl(classAb, "TimeActive", abSetId, abId, timeActive, arrs) = False Then Return "#Error"

                Dim param3 As String = "", tDiff As Long = (GVARS.PLAYER_TIME - CLng(lastRaisingTime))
                If ellapsedType = 0 Then
                    'TimeEllapsedType = 0 - через целые интервалы
                    If ReadPropertyDbl(classAb, "Interval", abSetId, abId, interval, arrs) = False Then Return "#Error"
                    If interval <= 0 Then Continue For
                    If timeActive > 0 Then
                        Dim intervalsSparedTotal As Long = Math.Floor((GVARS.PLAYER_TIME - CLng(appliedTime)) / interval)
                        If intervalsSparedTotal >= timeActive Then
                            'удаляем способность если время ее действия прошло
                            If FunctionRouter(classAb, "Remove", arrs, Nothing) = "#Error" Then Return "#Error"
                            Continue For
                        End If
                    End If
                    Dim fullIntervals As Integer = Math.Floor(tDiff / interval)
                    If fullIntervals = 0 Then Continue For 'времени, равного или более полному интервалу не прошло - событие не запускается

                    Dim raiseTime As Long = lastRaisingTime + fullIntervals * interval

                    'установка нового значения свойства TimeLastRaising и событие отслеживания свойства TimeLastRaising
                    Dim trackNewValue As String = raiseTime.ToString
                    Dim trackResult As String = mScript.trackingProperties.RunBefore(classAb, "TimeLastRaising", arrs, trackNewValue)
                    If trackResult <> "False" AndAlso trackResult <> "#Error" Then
                        SetPropertyValue(classAb, "TimeLastRaising", trackNewValue, abSetId, abId)
                    End If

                    param3 = fullIntervals.ToString
                Else
                    'TimeEllapsedType = 1 - при любом изменении времени
                    If timeActive > 0 Then
                        Dim timeSparedTotal As Long = Math.Floor(GVARS.PLAYER_TIME - CLng(appliedTime))
                        If timeSparedTotal >= timeActive Then
                            'удаляем способность если время ее действия прошло
                            If FunctionRouter(classAb, "Remove", arrs, Nothing) = "#Error" Then Return "#Error"
                            Continue For
                        End If
                    End If

                    param3 = tDiff.ToString

                    'установка ного значения свойства TimeLastRaising и событие отслеживания свойства TimeLastRaising
                    Dim trackNewValue As String = GVARS.PLAYER_TIME.ToString
                    Dim trackResult As String = mScript.trackingProperties.RunBefore(classAb, "TimeLastRaising", arrs, trackNewValue)
                    If trackResult <> "False" AndAlso trackResult <> "#Error" Then
                        SetPropertyValue(classAb, "TimeLastRaising", trackNewValue, abSetId, abId)
                    End If

                End If
                Dim abArrs() As String = {abSetId.ToString, abId.ToString, heroId.ToString, param3, (GVARS.PLAYER_TIME - CLng(appliedTime)).ToString} 'параметры для AbilityOnTimeEllapsedEvent


                If globalEventId > 0 Then
                    'глобальное событие
                    If mScript.eventRouter.RunEvent(globalEventId, abArrs, "AbilityOnTimeEllapsedEvent", False) = "#Error" Then Return "#Error"
                End If

                If abSetEventId > 0 Then
                    'событие данного набора
                    If mScript.eventRouter.RunEvent(abSetEventId, abArrs, "AbilityOnTimeEllapsedEvent", False) = "#Error" Then Return "#Error"
                End If

                Dim eventId As Integer = ch.ThirdLevelEventId(abId)
                If eventId > 0 Then
                    'событие данной способности
                    If mScript.eventRouter.RunEvent(eventId, abArrs, "AbilityOnTimeEllapsedEvent", False) = "#Error" Then Return "#Error"
                End If

                'звук при активанции SoundOnActivate
                Dim snd As String = ""
                If ReadProperty(classAb, "SoundOnActivate", abSetId, abId, snd, arrs) = False Then Return "#Error"
                snd = UnWrapString(snd)
                If String.IsNullOrEmpty(snd) = False Then
                    'убеждаемся в правильном формате пути к звуку
                    snd = snd.Replace("\", "/")
                    Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, snd)
                    If FileIO.FileSystem.FileExists(fPath) = False Then
                        MessageBox.Show("Аудиофайл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return "#Error"
                    End If

                    'проигрывем звук
                    Say(snd, 100, False)
                End If
            Next abId
        Next i

        Return ""
    End Function

    ''' <summary>
    ''' Возвращает список героев/бойцов по значению свойства H.AbilitiesSet
    ''' </summary>
    ''' <param name="abSetId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetHeroesListByAbilitySet(ByVal abSetId As Integer) As List(Of Integer)
        Dim lst As New List(Of Integer)
        Dim classAb As Integer = mScript.mainClassHash("Ab"), classH As Integer = mScript.mainClassHash("H")
        If IsNothing(mScript.mainClass(classAb).ChildProperties) OrElse mScript.mainClass(classAb).ChildProperties.Count = 0 Then Return lst
        If GVARS.G_ISBATTLE Then
            If mScript.Battle.Fighters.Count = 0 Then Return lst
        Else
            If IsNothing(mScript.mainClass(classH).ChildProperties) OrElse mScript.mainClass(classH).ChildProperties.Count = 0 Then Return lst
        End If
        Dim abSetName As String = mScript.mainClass(classAb).ChildProperties(abSetId)("Name").Value
        Dim abSetIdStr As String = abSetId.ToString

        Dim forEnd As Integer
        If GVARS.G_ISBATTLE Then
            forEnd = mScript.Battle.Fighters.Count - 1
        Else
            forEnd = mScript.mainClass(classH).ChildProperties.Count - 1
        End If
        For hId As Integer = 0 To forEnd
            Dim hSet As String = ""
            If GVARS.G_ISBATTLE Then
                If mScript.Battle.ReadFighterProperty("AbilitiesSet", hId, hSet, {hId.ToString}) = False Then Continue For
            Else
                If ReadProperty(classH, "AbilitiesSet", hId, -1, hSet, {hId.ToString}) = False Then Continue For
            End If
            If hSet = abSetIdStr OrElse String.Compare(hSet, abSetName, True) = 0 Then lst.Add(hId)
        Next hId

        Return lst
    End Function
#End Region
End Module
