Module modPlayer
    Public Class cGlobalVariables
        Public G_CURLOC As Integer
        Public G_PREVLOC As Integer
        Public G_CURMENU As Integer = -1
        Public G_CURMAGIC As Integer = -1
        Public G_CURLIST As Integer
        Public G_CURAUDIO As Integer
        'Public G_SOUND As Boolean
        'Public G_USERSEND As String
        Public G_CURMAP As Integer
        Public G_CURMAPCELL As Integer
        Public G_PREVMAPCELL As Integer
        'Public G_CURHERO As Integer
        Public G_ISBATTLE As Boolean
        Public G_CURFIGHTER As Integer
        Public G_CURVICTIM As Integer
        Public G_STRIKETYPE As Integer
        'Public G_BATTLEDURATION As Integer
        ''' <summary>Только для боя. Можно ли использовать магию прямо сейчас.</summary>
        Public G_CANUSEMAGIC As Boolean = True
        'Public G_IGNOREERRORS As Boolean
        'Public G_LASTERROR As String
        'Public Q_PATH As String
        'Public Q_SHORTNAME As String
        'Public Q_FULLNAME As String
        'Public Q_DESC As String
        'Public Q_STARTLOC As String
        Public TIME_STARTING As Long
        Public TIME_SAVING As Long
        Public TIME_IN_GAME As Long
        Public TIME_IN_THIS_LOCATION As Long
        Public TIME_IN_BATTLE As Long
        Public HOLD_UP_TIMERS As Boolean = False
        ''' <summary>время, определенное свойствами и функциями Date.[InGame] - количество секунд</summary>
        Public PLAYER_TIME As Long
        'Public SAVES_COUNT As Integer
        'Public OPEN_PARAMS As String
        'Public EDIT_MODE As Boolean
        'Public LAST_GAME As String
        'Public Q_PASSWORD As String
    End Class
    ''' <summary>Глобальные переменные игры</summary>
    Public GVARS As New cGlobalVariables
    ''' <summary>Было ли сгенерировано последнее событие со скрипта</summary>
    Public EventGeneratedFromScript As Boolean = False

    Public Const CONST_WBHEIGHT_CORRECTION As Integer = 20 'при получении высоты WebBrowser надо отнять данную величину чтобы получить правильный результат (ошибка vb.net???)
    ''' <summary>Запрет на вывод кнопок действий на экран во время загрузки локации</summary>
    Public ActionsInputProhibited As Boolean = False

#Region "StartGame"
    ''' <summary>
    ''' Класс для хранения ширины и высоты панелей, а также их видимости
    ''' </summary>
    Private Class WindowSizeClass
        ''' <summary>Ширина окна</summary>
        Public Width As Double
        ''' <summary>Высота окна</summary>
        Public Height As Double
        ''' <summary>Ширина абсолютная (в пикселах) или в %</summary>
        Public IsWidthAbsolute As Boolean
        ''' <summary>Высота абсолютная (в пикселах) или в %</summary>
        Public isHeightAbsolute
        ''' <summary>Видимость окна</summary>
        Public Visible As Boolean

        ''' <summary>Является ли окно растянутым на весь экран</summary>
        Public Function IsMaximized()
            If Width >= 100 AndAlso Height >= 100 AndAlso IsWidthAbsolute = False AndAlso isHeightAbsolute = False Then
                Return True
            ElseIf Width >= Screen.PrimaryScreen.WorkingArea.Width AndAlso Height >= Screen.PrimaryScreen.WorkingArea.Height Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Sub New(ByRef wSize As WindowSizeClass)
            With wSize
                Width = .Width
                Height = .Height
                IsWidthAbsolute = .IsWidthAbsolute
                isHeightAbsolute = .isHeightAbsolute
                Visible = .Visible
            End With
        End Sub

        Public Sub New(ByVal classId As Integer)
            GetSizes(classId)
        End Sub

        'Получает размеры окнон из свойств Width, Height и Visible соответствующих классов
        Private Sub GetSizes(ByVal classId As Integer)
            'получаем ширину
            Dim aStr As String = mScript.mainClass(classId).Properties("Width").Value
            If String.IsNullOrWhiteSpace(aStr) OrElse aStr = "''" Then
                '100%
                Width = 100
                IsWidthAbsolute = False
            ElseIf mScript.Param_GetType(aStr) = MatewScript.ReturnFormatEnum.TO_NUMBER Then
                'в пикселах
                aStr = aStr.Replace(".", ",")
                Width = Val(aStr)
                IsWidthAbsolute = True
            Else
                'в %
                aStr = mScript.PrepareStringToPrint(aStr, Nothing).Replace(".", ",")
                Width = Val(aStr)
                IsWidthAbsolute = False
            End If

            'получаем высоту
            aStr = mScript.mainClass(classId).Properties("Height").Value
            If String.IsNullOrWhiteSpace(aStr) OrElse aStr = "''" Then
                '100%
                Height = 100
                isHeightAbsolute = False
            ElseIf mScript.Param_GetType(aStr) = MatewScript.ReturnFormatEnum.TO_NUMBER Then
                'в пикселах
                aStr = aStr.Replace(".", ",")
                Height = Val(aStr)
                isHeightAbsolute = True
            Else
                'в %
                aStr = mScript.PrepareStringToPrint(aStr, Nothing).Replace(".", ",")
                Height = Val(aStr)
                isHeightAbsolute = False
            End If

            'получаем видимость
            Dim pVis As MatewScript.PropertiesInfoType = Nothing
            If mScript.mainClass(classId).Properties.TryGetValue("Visible", pVis) = False Then
                Visible = True
            Else
                Visible = CBool(pVis.Value)
            End If
        End Sub
    End Class

    ''' <summary>
    ''' Устанавливает расположение окнов в слотах, ставит размеры
    ''' </summary>
    Public Sub SetWindowsContainers()
        If questEnvironment.wbMap.IsDisposed Then questEnvironment.wbMap = New WebBrowser With {.IsWebBrowserContextMenuEnabled = False, .AllowNavigation = False, .Dock = DockStyle.Fill}
        If questEnvironment.wbMagic.IsDisposed Then questEnvironment.wbMagic = New WebBrowser With {.IsWebBrowserContextMenuEnabled = False, .AllowNavigation = False, .Dock = DockStyle.Fill}
        AddHandler questEnvironment.wbMap.PreviewKeyDown, AddressOf frmPlayer.wbMain_PreviewKeyDown
        AddHandler questEnvironment.wbMagic.PreviewKeyDown, AddressOf frmPlayer.wbMain_PreviewKeyDown
        'получаем классы окон
        Dim classLW As Integer = mScript.mainClassHash("LW")
        Dim classOW As Integer = mScript.mainClassHash("OW")
        Dim classAW As Integer = mScript.mainClassHash("AW")
        Dim classDW As Integer = mScript.mainClassHash("DW")
        Dim classCmd As Integer = mScript.mainClassHash("Cmd")
        Dim classMgW As Integer = mScript.mainClassHash("MgW")
        Dim classMap As Integer = mScript.mainClassHash("Map")
        'получаем слоты всех окон
        Dim slotLW As Integer = Val(mScript.PrepareStringToPrint(mScript.mainClass(classLW).Properties("Slot").Value, Nothing))
        Dim slotOW As Integer = Val(mScript.PrepareStringToPrint(mScript.mainClass(classOW).Properties("Slot").Value, Nothing))
        Dim slotAW As Integer = Val(mScript.PrepareStringToPrint(mScript.mainClass(classAW).Properties("Slot").Value, Nothing))
        Dim slotDW As Integer = Val(mScript.PrepareStringToPrint(mScript.mainClass(classDW).Properties("Slot").Value, Nothing))
        Dim slotCmd As Integer = Val(mScript.PrepareStringToPrint(mScript.mainClass(classCmd).Properties("Slot").Value, Nothing))
        Dim slotMgW As Integer = Val(mScript.PrepareStringToPrint(mScript.mainClass(classMgW).Properties("Slot").Value, Nothing))
        Dim slotMap As Integer = Val(mScript.PrepareStringToPrint(mScript.mainClass(classMap).Properties("Slot").Value, Nothing))
        'размещаем окна по слотам
        With frmPlayer
            PutWindowToSlot(.wbMain, slotLW)
            PutWindowToSlot(.wbObjects, slotOW)
            PutWindowToSlot(.wbActions, slotAW)
            PutWindowToSlot(.wbDescription, slotDW)
            PutWindowToSlot(.wbCommand, slotCmd)
            PutWindowToSlot(questEnvironment.wbMap, slotMap)
            PutWindowToSlot(questEnvironment.wbMagic, slotMgW)
        End With
        'получаем размеры и видимость окон
        Dim sizeLW As New WindowSizeClass(classLW)
        Dim sizeOW As New WindowSizeClass(classOW)
        Dim sizeAW As New WindowSizeClass(classAW)
        Dim sizeDW As New WindowSizeClass(classDW)
        Dim sizeCmd As New WindowSizeClass(classCmd)

        Dim sizeMgW As New WindowSizeClass(classMgW)
        Dim sizeMap As New WindowSizeClass(classMap)
        'создаем список sizes, где ключ - номер слота, а value - структура WindowSizeClass первой из размещенных в этом слоте панелей
        Dim sizes As New SortedList(Of Integer, WindowSizeClass)
        Dim visibles As New SortedList(Of Integer, Boolean)
        sizes.Add(slotLW, sizeLW)
        If sizes.ContainsKey(slotOW) = False Then
            sizes.Add(slotOW, sizeOW)
        End If
        If sizes.ContainsKey(slotAW) = False Then
            sizes.Add(slotAW, sizeAW)
        End If
        If sizes.ContainsKey(slotDW) = False Then
            sizes.Add(slotDW, sizeDW)
        End If
        If sizes.ContainsKey(slotCmd) = False Then
            sizes.Add(slotCmd, sizeCmd)
        End If
        If sizes.ContainsKey(slotMgW) = False Then
            sizes.Add(slotMgW, sizeMgW)
        End If
        If sizes.ContainsKey(slotMap) = False Then
            sizes.Add(slotMap, sizeMap)
        End If

        Dim leftSize As WindowSizeClass = Nothing 'панель в 1 слоте
        sizes.TryGetValue(0, leftSize)
        Dim rightSize As WindowSizeClass = Nothing 'панель во 2 слоте
        sizes.TryGetValue(1, rightSize)

        Dim totalWidthCalculated As Boolean = False 'рассчитана ли уже ширина игрового окна
        Dim wMaximized As Boolean = False 'игровое окно на весь экран?

        'устанавливаем ширину если в 1 или 2 слоте есть хоть одна панель
        With frmPlayer
            If IsNothing(leftSize) AndAlso IsNothing(rightSize) Then
                'в 1 и 2 слоте пусто - коллапсируем оба слота
                .splitHorizontal.Panel1Collapsed = True
            ElseIf IsNothing(rightSize) Then
                'слот предметов(2) пуст, а первый занят
                'устанавливаем ширину окна
                .split12.Panel1Collapsed = False
                .split12.Panel2Collapsed = True
                If leftSize.IsMaximized Then
                    wMaximized = True
                    .Width = Screen.PrimaryScreen.WorkingArea.Width
                ElseIf leftSize.IsWidthAbsolute Then
                    'absolute
                    .Width = leftSize.Width
                Else
                    '%
                    .Width = CInt(Math.Round(100 * Screen.PrimaryScreen.WorkingArea.Width / leftSize.Width, 0))
                End If
                totalWidthCalculated = True
            ElseIf IsNothing(leftSize) Then
                'главный слот пуст, а 2 занят
                'устанавливаем ширину окна
                .split12.Panel1Collapsed = True
                .split12.Panel2Collapsed = False
                If rightSize.IsMaximized Then
                    wMaximized = True
                    .Width = Screen.PrimaryScreen.WorkingArea.Width
                ElseIf rightSize.IsWidthAbsolute Then
                    'absolute
                    .Width = rightSize.Width
                Else
                    '%
                    .Width = CInt(Math.Round(0.01 * Screen.PrimaryScreen.WorkingArea.Width * rightSize.Width, 0))
                End If
                totalWidthCalculated = True
            Else
                'оба слота заполнены
                'устанавливаем ширину окна, а также ширину обоих слотов
                .split12.Panel1Collapsed = False
                .split12.Panel2Collapsed = False
                If leftSize.IsMaximized AndAlso rightSize.IsMaximized Then
                    wMaximized = True
                    .Width = Screen.PrimaryScreen.WorkingArea.Width
                    .split12.SplitterDistance = .Width * 0.75
                    totalWidthCalculated = True
                ElseIf leftSize.IsWidthAbsolute AndAlso rightSize.IsWidthAbsolute Then
                    'размеры обеих слотов указаны в абсолютном значении
                    .Width = leftSize.Width + rightSize.Width
                    .split12.SplitterDistance = leftSize.Width
                    totalWidthCalculated = True
                ElseIf leftSize.IsWidthAbsolute = False AndAlso rightSize.IsWidthAbsolute = False Then
                    'размеры обеих слотов указаны в процентах
                    Dim w As Integer
                    If leftSize.Width >= 100 AndAlso rightSize.Width >= 100 Then
                        'в 1 и 2 размеры 100%+. На весь экран, ширина 1 слота = 25% первого
                        .Width = Screen.PrimaryScreen.WorkingArea.Width
                        .split12.SplitterDistance = .Width * 0.75
                        wMaximized = True
                    ElseIf leftSize.Width >= 100 Then
                        'ширина 1-го 100%+, а 2-го меньше 100%. На весь экран, ширину 2 в сохраняем, остальная ширина экрана - 1 слот
                        w = 0.01 * Screen.PrimaryScreen.WorkingArea.Width * (100 - rightSize.Width)
                        .Width = Screen.PrimaryScreen.WorkingArea.Width
                        .split12.SplitterDistance = w
                        wMaximized = True
                    ElseIf rightSize.Width >= 100 OrElse (leftSize.Width + rightSize.Width >= 100) Then
                        'ширина 2-го 100%+, а 1-го меньше 100%. На весь экран, ширину 1 в сохраняем, остальная ширина экрана - 2 слот
                        w = 0.01 * Screen.PrimaryScreen.WorkingArea.Width * leftSize.Width
                        .Width = Screen.PrimaryScreen.WorkingArea.Width
                        .split12.SplitterDistance = w
                        wMaximized = True
                    Else
                        'суммарная ширина меньше 100%. На Расчитываем ширину, беря за 100% ширину экрана
                        w = 0.01 * Screen.PrimaryScreen.WorkingArea.Width * leftSize.Width
                        Dim w2 As Integer = 0.01 * Screen.PrimaryScreen.WorkingArea.Width * rightSize.Width
                        .Width = w + w2
                        .split12.SplitterDistance = w
                    End If
                    totalWidthCalculated = True
                End If
            End If

            'устанавливаем высоту
            Dim bottomSize As WindowSizeClass = Nothing 'нижний слот, учитывая который определяем высоту
            'получаем нижний слот, с которым будем сравнивать верхние для получения высоты
            If sizes.ContainsKey(3) Then
                'берем 4 слот
                bottomSize = sizes(3)
            ElseIf sizes.ContainsKey(2) Then
                'слот 4 пуст, берем третий (действий)
                If sizes.ContainsKey(4) Then
                    'есть и actW, и cmd
                    bottomSize = New WindowSizeClass(sizes(2))
                    If bottomSize.isHeightAbsolute Then
                        If sizes(4).isHeightAbsolute Then
                            'обе высоты абсолютны - получаем суммарную высоту 3 и 5 слота
                            bottomSize.Height += sizes(4).Height
                        Else
                            'высота cmd в % - получаем суммарную
                            bottomSize.Height = bottomSize.Height * (1 + sizes(4).Height * 100)
                        End If
                    Else
                        If sizes(4).isHeightAbsolute Then
                            'высота cmd абсолютна, act % - получаем суммарную высоту
                            bottomSize.Height = sizes(4).Height * (1 + bottomSize.Height / 100)
                        Else
                            'обе высоты в % (при этом в DescW пусто). Берем высоту = 40% от ширины экрана
                            bottomSize.Height = Screen.PrimaryScreen.WorkingArea.Height * 0.4 '40% от LW
                        End If
                    End If
                Else
                    'только actW. cmd (5 слот) пустой - берем высоту 3 слота
                    bottomSize = sizes(2)
                End If
            ElseIf sizes.ContainsKey(4) Then
                '3 и слоты пустые.Внизу только 5 слот - берем его высоту
                bottomSize = sizes(4)
            End If

            'коллапсируем нижние слоты которые пустые
            If sizes.ContainsKey(2) Then
                If sizes.ContainsKey(4) Then
                    'обе есть
                    .split35.Panel1Collapsed = False
                    .split35.Panel2Collapsed = False
                Else
                    'act есть, cmd нет
                    .split35.Panel1Collapsed = False
                    .split35.Panel2Collapsed = True
                End If
            Else
                If sizes.ContainsKey(4) Then
                    'act пусто, cmd есть
                    .split35.Panel1Collapsed = True
                    .split35.Panel2Collapsed = False
                Else
                    'обе пусты
                    .split35.Panel1Collapsed = False
                    .split35.Panel2Collapsed = False
                End If
            End If

            If IsNothing(leftSize) Then leftSize = rightSize 'право/лево уже не важно, далее главное - высота
            'теперь для вычисления высоты игрового окна сравниваем leftSize (хранит высоту 1 или 2 панели) и bottomSize (хранит информацию о высоте нижних слотов)
            If IsNothing(bottomSize) Then
                'внизу нет ни одного окна
                'IsNothing(leftSize) быть не может - окна должны быть хоть где-то
                .splitHorizontal.Panel1Collapsed = False
                .splitHorizontal.Panel2Collapsed = True
                If Not wMaximized Then
                    If leftSize.isHeightAbsolute Then
                        'высота абсолютна
                        .Height = leftSize.Height
                    Else
                        'высота в %
                        If leftSize.Height >= 100 Then
                            .Height = Screen.PrimaryScreen.WorkingArea.Height
                        Else
                            .Height = 0.01 * Screen.PrimaryScreen.WorkingArea.Height * leftSize.Height
                        End If
                    End If
                End If
            ElseIf IsNothing(leftSize) Then
                'внизу окна есть, сверху - нет
                .splitHorizontal.Panel1Collapsed = True
                .splitHorizontal.Panel2Collapsed = False
                If bottomSize.IsMaximized Then
                    wMaximized = True
                    .Width = Screen.PrimaryScreen.WorkingArea.Width
                ElseIf bottomSize.isHeightAbsolute Then
                    .Height = bottomSize.Height
                Else
                    If bottomSize.Height >= 100 Then
                        .Height = Screen.PrimaryScreen.WorkingArea.Height
                        .Width = Screen.PrimaryScreen.WorkingArea.Width
                        wMaximized = True
                    Else
                        .Height = 0.01 * Screen.PrimaryScreen.WorkingArea.Height * bottomSize.Height
                    End If
                End If
            Else
                'окна есть и сверху, и снизу
                .splitHorizontal.Panel1Collapsed = False
                .splitHorizontal.Panel2Collapsed = False
                If leftSize.isHeightAbsolute AndAlso bottomSize.isHeightAbsolute Then
                    .Height = leftSize.Height + bottomSize.Height
                    .splitHorizontal.SplitterDistance = leftSize.Height
                ElseIf leftSize.isHeightAbsolute AndAlso bottomSize.isHeightAbsolute = False Then
                    .Height = leftSize.Height * (1 + bottomSize.Height / 100)
                    .splitHorizontal.SplitterDistance = leftSize.Height
                ElseIf leftSize.isHeightAbsolute = False AndAlso bottomSize.isHeightAbsolute Then
                    .Height = bottomSize.Height * (1 + leftSize.Height / 100)
                    .splitHorizontal.SplitterDistance = .Height - bottomSize.Height
                Else
                    'оба в процентах
                    Dim h1 As Integer = 0.01 * Screen.PrimaryScreen.WorkingArea.Height * leftSize.Height
                    Dim h2 As Integer = 0.01 * Screen.PrimaryScreen.WorkingArea.Height * bottomSize.Height
                    .Height = h1 + h2
                    .splitHorizontal.SplitterDistance = h1
                End If
            End If

            'устанавливаем ширину 3/5 и 4 слотов
            If IsNothing(bottomSize) = False Then
                'снизу есть хоть одна панель
                If sizes.ContainsKey(3) = False Then
                    'DW  пусто
                    .split345.Panel1Collapsed = False
                    .split345.Panel2Collapsed = True
                ElseIf sizes.ContainsKey(2) = False AndAlso sizes.ContainsKey(4) = False Then
                    'AW и Cmd пусто
                    .split345.Panel1Collapsed = True
                    .split345.Panel2Collapsed = False
                Else
                    'внизу есть что-то и справа, и слева
                    .split345.Panel1Collapsed = False
                    .split345.Panel2Collapsed = False
                    If sizes.ContainsKey(2) Then
                        leftSize = sizes(2)
                    Else
                        leftSize = sizes(4)
                    End If
                    rightSize = sizes(3)

                    If leftSize.IsWidthAbsolute AndAlso rightSize.IsWidthAbsolute Then
                        'справа и слева в AW/Cmd и DW  указаны абсолютные размеры
                        If Not totalWidthCalculated Then
                            .Width = leftSize.Width + rightSize.Width
                        End If
                        .split345.SplitterDistance = leftSize.Width
                    ElseIf leftSize.IsWidthAbsolute AndAlso rightSize.isHeightAbsolute = False Then
                        'слева в AW/Cmd абсолютные, справа %
                        If Not totalWidthCalculated Then
                            .Width = leftSize.Width * (1 + rightSize.Width / 100)
                        End If
                        .split345.SplitterDistance = leftSize.Width
                    ElseIf leftSize.IsWidthAbsolute = False AndAlso rightSize.IsWidthAbsolute Then
                        'слева %, справа - px
                        If Not totalWidthCalculated Then
                            .Width = rightSize.Width * (1 + leftSize.Width / 100)
                        End If
                        .split345.SplitterDistance = .Width - leftSize.Width
                    Else
                        'справа и слева в %
                        If Not totalWidthCalculated Then
                            wMaximized = True
                            .Width = Screen.PrimaryScreen.WorkingArea.Width
                            .split345.SplitterDistance = 0.01 * Screen.PrimaryScreen.WorkingArea.Width * leftSize.Width
                        Else
                            .split345.SplitterDistance = 0.01 * .Width * leftSize.Width
                        End If
                    End If
                End If
            End If

            If wMaximized Then
                .WindowState = FormWindowState.Maximized
            Else
                .WindowState = FormWindowState.Normal
            End If

            'коллапсируем панели если там невидимый контент
            Dim vis As Boolean = False

            For i = 0 To 4
                'идем дальше если слот уже коллапсирован
                Select Case i
                    Case 0
                        If .split12.Panel1Collapsed Then Continue For
                    Case 1
                        If .split12.Panel2Collapsed Then Continue For
                    Case 2
                        If .split345.Panel1Collapsed OrElse .split35.Panel1Collapsed Then Continue For
                    Case 3
                        If .split345.Panel2Collapsed Then Continue For
                    Case 4
                        If .split345.Panel1Collapsed OrElse .split35.Panel2Collapsed Then Continue For
                End Select
                'на основании того какое окно в слоте и видимо ли оно получаем видимость слота
                vis = False
                If slotLW = i Then
                    vis = True
                ElseIf slotOW = i AndAlso sizeOW.Visible Then
                    vis = True
                ElseIf slotAW = i AndAlso sizeAW.Visible Then
                    vis = True
                ElseIf slotDW = i AndAlso sizeDW.Visible Then
                    vis = True
                ElseIf slotCmd = i AndAlso sizeCmd.Visible Then
                    vis = True
                End If

                If Not vis Then
                    'слот надо коллапсировать
                    Select Case i
                        Case 0
                            .split12.Panel1Collapsed = True
                        Case 1
                            .split12.Panel2Collapsed = True
                        Case 2
                            .split35.Panel1Collapsed = True
                        Case 3
                            .split345.Panel2Collapsed = True
                        Case 4
                            If .split35.Panel1Collapsed Then
                                .split345.Panel1Collapsed = True
                            Else
                                .split35.Panel2Collapsed = True
                            End If
                    End Select
                End If
            Next
            If .split345.Panel1Collapsed AndAlso .split345.Panel1Collapsed Then
                'слот 3, 4 и 5 коллапсированы. Коллапсируем их общую родительскую панель
                .splitHorizontal.Panel2Collapsed = True
            ElseIf .split12.Panel1Collapsed AndAlso .split12.Panel2Collapsed Then
                'слот 1 и 2 коллапсированы. Коллапсируем их общую родительскую панель
                .splitHorizontal.Panel1Collapsed = True
            End If
        End With

        If slotMap = 5 Then
            'карты - в отдельном окне. Устанавливаем размеры этого окна
            If sizeMap.IsMaximized Then
                frmMap.WindowState = FormWindowState.Maximized
            Else
                frmMap.WindowState = FormWindowState.Normal
                frmMap.FormBorderStyle = FormBorderStyle.FixedSingle
                If sizeMap.IsWidthAbsolute = False Then sizeMap.Width = 0.01 * Screen.PrimaryScreen.WorkingArea.Width * Math.Min(sizeMap.Width, 100)
                If sizeMap.isHeightAbsolute = False Then sizeMap.Height = 0.01 * Screen.PrimaryScreen.WorkingArea.Height * Math.Min(sizeMap.Height, 100)
                frmMap.Size = New Size(sizeMap.Width, sizeMap.Height)
            End If
        End If

        If slotMgW = 5 Then
            'магии - в отдельном окне. Устанавливаем размеры этого окна
            If sizeMgW.IsMaximized Then
                frmMagic.WindowState = FormWindowState.Maximized
            Else
                frmMagic.WindowState = FormWindowState.Normal
                If sizeMgW.IsWidthAbsolute = False Then sizeMgW.Width = 0.01 * Screen.PrimaryScreen.WorkingArea.Width * Math.Min(sizeMgW.Width, 100)
                If sizeMgW.isHeightAbsolute = False Then sizeMgW.Height = 0.01 * Screen.PrimaryScreen.WorkingArea.Height * Math.Min(sizeMgW.Height, 100)
                frmMagic.Size = New Size(sizeMgW.Width, sizeMgW.Height)
            End If
        End If

    End Sub

    ''' <summary>Размещает панель WB в указанном слоте</summary>
    ''' <param name="wb">ссылка на wb-панель</param>
    ''' <param name="slot">номер слота от 0</param>
    Private Sub PutWindowToSlot(ByRef wb As WebBrowser, ByVal slot As Integer)
        With frmPlayer
            Select Case slot
                Case 0 'Main
                    If Object.Equals(wb.Parent, .split12.Panel1) Then Return
                    .split12.Panel1.Controls.Add(wb)
                Case 1 'Obj
                    If Object.Equals(wb.Parent, .split12.Panel2) Then Return
                    .split12.Panel2.Controls.Add(wb)
                Case 2 'Act
                    If Object.Equals(wb.Parent, .split35.Panel1) Then Return
                    .split35.Panel1.Controls.Add(wb)
                Case 3 'Desc
                    If Object.Equals(wb.Parent, .split345.Panel2) Then Return
                    .split345.Panel1.Controls.Add(wb)
                Case 4 'Cmd
                    If Object.Equals(wb.Parent, .split35.Panel2) Then Return
                    .split35.Panel2.Controls.Add(wb)
                Case 5 'Other Wnd
                    If Object.Equals(wb, questEnvironment.wbMap) Then
                        If Object.Equals(wb.Parent, frmMap) Then Return
                        frmMap.Controls.Add(wb)
                        wb.Dock = DockStyle.Fill
                    Else
                        If Object.Equals(wb.Parent, frmMagic) Then Return
                        frmMagic.Controls.Add(wb)
                        wb.Dock = DockStyle.Fill
                    End If
            End Select
        End With
    End Sub

    ''' <summary>Загружает документы из папки квеста</summary>
    Public Function LoadPlayerDocuments() As Boolean
        Dim fNames() = {"Location.html", "Objects.html", "Actions.html", "Description.html", "Command.html", "Map.html", "Magic.html"}
        For i = 0 To fNames.Count - 1
            Dim f As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, fNames(i))
            If FileIO.FileSystem.FileExists(f) = False Then
                MessageBox.Show("В папке квеста не найден файл " & f & ", необходимый для запуска игры. Попробуйте скопировать его вручную из " & _
                                FileIO.FileSystem.CombinePath(Application.StartupPath, "src\defaultFiles") & " .", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
        Next

        With frmPlayer
            .wbMain.Navigate(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Location.html"))
            .wbObjects.Navigate(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Objects.html"))
            .wbActions.Navigate(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Actions.html"))
            .wbDescription.Navigate(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Description.html"))
            .wbCommand.Navigate(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Command.html"))
            questEnvironment.wbMap.Navigate(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Map.html"))
            questEnvironment.wbMagic.Navigate(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Magic.html"))

            Do Until .wbMain.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            Do Until .wbObjects.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            Do Until .wbActions.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            Do Until .wbDescription.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            Do Until .wbCommand.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            Do Until questEnvironment.wbMap.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            Do Until questEnvironment.wbMagic.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop

            'добавляем css действий в главное окно (для корректного вывода действий)
            Dim actCSS As String = ""
            ReadProperty(mScript.mainClassHash("AW"), "CSS", -1, -1, actCSS, Nothing)
            actCSS = UnWrapString(actCSS)
            If String.IsNullOrEmpty(actCSS) = False Then HtmlAppendCSS(.wbMain.Document, actCSS)

            'добавляем css меню в главное окно (для корректного вывода меню)
            ReadProperty(mScript.mainClassHash("DW"), "CSS", -1, -1, actCSS, Nothing)
            actCSS = UnWrapString(actCSS)
            If String.IsNullOrEmpty(actCSS) = False Then HtmlAppendCSS(.wbMain.Document, actCSS)

            Dim hText As HtmlElement = .wbCommand.Document.GetElementById("cmdText")
            If IsNothing(hText) = False Then
                AddHandler hText.KeyPress, Sub(sender As HtmlElement, e As HtmlElementEventArgs)
                                               If e.KeyPressedCode = Keys.Enter Then CommandClick(hText)
                                           End Sub
            End If
            Dim hBtn As HtmlElement = .wbCommand.Document.GetElementById("cmdButton")
            If IsNothing(hBtn) = False Then AddHandler hBtn.Click, Sub(sender As Object, e As HtmlElementEventArgs) CommandClick(hText)
        End With

        If PrepareVisualization() = "#Error" Then Return False
        questEnvironment.wbMagic.ObjectForScripting = frmPlayer
        Return True
    End Function

    ''' <summary>
    ''' Подготавливает внешний вид окон
    ''' </summary>
    ''' <remarks></remarks>
    Private Function PrepareVisualization() As String
        For i As Integer = 1 To 3
            Dim classAW As Integer = Choose(i, mScript.mainClassHash("AW"), mScript.mainClassHash("DW"), mScript.mainClassHash("Cmd"))
            Dim hDoc As HtmlDocument = Choose(i, frmPlayer.wbActions.Document, frmPlayer.wbDescription.Document, frmPlayer.wbCommand.Document)
            If IsNothing(hDoc) Then Continue For
            Dim bkColor As String = "", bkPicture As String = "", strStyle As New System.Text.StringBuilder
            If ReadProperty(classAW, "BackColor", -1, -1, bkColor, Nothing) = False Then Return "#Error"
            bkColor = UnWrapString(bkColor)

            If ReadProperty(classAW, "BackPicture", -1, -1, bkPicture, Nothing) = False Then Return "#Error"
            bkPicture = UnWrapString(bkPicture)
            If bkPicture.Length > 0 Then
                strStyle.AppendFormat("background-color: {0};", bkColor)
                strStyle.AppendFormat("background-image: url({0});", "'" + bkPicture.Replace("\", "/") + "'")
                Dim bkPicStyle As Integer = 0
                If ReadPropertyInt(classAW, "BackPicStyle", -1, -1, bkPicStyle, Nothing) = False Then Return "#Error"
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
                    If ReadPropertyInt(classAW, "BackPicPos", -1, -1, bkPicPos, Nothing) = False Then Return "#Error"
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
            Else
                strStyle.AppendFormat("background: {0};", bkColor)
            End If

            hDoc.Body.Style = strStyle.ToString
            strStyle.Clear()
        Next i

        Return ""
    End Function

    Private Sub del_LocationHtmlEvent(hText As Object, e As EventArgs)
        Dim strText As String = hText.GetAttribute("Value")
        If String.IsNullOrEmpty(strText) Then Return

        'событие LocationHtmlEvent
        Dim classL As Integer = mScript.mainClassHash("L")
        Dim eventId As Integer = mScript.mainClass(classL).Properties("LocationHtmlEvent").eventId
        If eventId > 0 Then
            'глобальное событие
            If mScript.eventRouter.RunEvent(eventId, {WrapString(strText)}, "LocationHtmlEvent", True) = "#Error" Then Return
        End If

        If GVARS.G_CURLOC > -1 Then
            eventId = mScript.mainClass(classL).ChildProperties(GVARS.G_CURLOC)("LocationHtmlEvent").eventId
            If eventId > 0 Then
                'событие текущей локации
                mScript.eventRouter.RunEvent(eventId, {WrapString(strText)}, "LocationHtmlEvent", True)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Событие командной строки
    ''' </summary>
    ''' <param name="hText">ссылка на html-элемент текстбокса окна команд</param>
    ''' <remarks></remarks>
    Private Sub CommandClick(ByRef hText As HtmlElement)
        Dim strCommand As String = hText.GetAttribute("Value").Trim
        If String.IsNullOrEmpty(strCommand) Then Return

        If strCommand.StartsWith("?") Then
            'введена команда - на получение значения
            With questEnvironment.codeBoxShadowed
                .Tag = Nothing
                .codeBox.IsTextBlockByDefault = False
                .Text = strCommand
                Dim result As String = mScript.ExecuteCode(mScript.PrepareBlock(.codeBox.CodeData), Nothing, True)
                MessageBox.Show(result, "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End With
            Return
        ElseIf strCommand.StartsWith("#") Then
            'тоже команда - на выполнение
            With questEnvironment.codeBoxShadowed
                .Tag = Nothing
                .codeBox.IsTextBlockByDefault = False
                .Text = strCommand.Substring(1)
                mScript.ExecuteCode(mScript.PrepareBlock(.codeBox.CodeData), Nothing, True)
            End With

            Return
        End If

        'событие LocationCommandEvent
        Dim classL As Integer = mScript.mainClassHash("L")
        Dim eventId As Integer = mScript.mainClass(classL).Properties("LocationCommandEvent").eventId
        Dim eventRunned As Boolean = False
        If eventId > 0 Then
            'глобальное событие
            If mScript.eventRouter.RunEvent(eventId, {WrapString(strCommand)}, "LocationCommandEvent", False) = "#Error" Then Return
            eventRunned = True
        End If

        If GVARS.G_CURLOC > -1 Then
            eventId = mScript.mainClass(classL).ChildProperties(GVARS.G_CURLOC)("LocationCommandEvent").eventId
            If eventId > 0 Then
                'событие текущей локации
                If mScript.eventRouter.RunEvent(eventId, {WrapString(strCommand)}, "LocationCommandEvent", False) = "#Error" Then Return
                eventRunned = True
            End If
        End If

        If eventRunned Then mScript.eventRouter.RunScriptFinishedEvent({WrapString(strCommand)})
    End Sub
#End Region

#Region "Locations And Actions"
    ''' <summary>
    ''' Выполняет функцию L.Go
    ''' </summary>
    ''' <param name="arrParams">0 - id локации, 1 - выполнять обработчики, 2 - переход через карту, 3... - еще параметры</param>
    ''' <returns>Пустую строку или #Error или False, если переход был отменен</returns>
    Public Function Go(ByRef arrParams() As String, Optional doLocationExitEvent As Boolean = True) As String
        'If questEnvironment.EDIT_MODE Then Return ""
        'Загружаем Battle.html
        If GVARS.G_ISBATTLE = False AndAlso mScript.Battle.JustAfterBattle Then
            mScript.Battle.JustAfterBattle = False
            Dim f As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Location.html")
            If FileIO.FileSystem.FileExists(f) = False Then
                Return _ERROR("В папке квеста не найден файл " & f & ", необходимый для запуска игры. Попробуйте скопировать его вручную из " & _
                                FileIO.FileSystem.CombinePath(Application.StartupPath, "src\defaultFiles") & " .")
            End If
            frmPlayer.wbMain.AllowNavigation = True
            frmPlayer.wbMain.Navigate(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Location.html"))
            Do Until frmPlayer.wbMain.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
            If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть html-документ Location.html!")
            frmPlayer.wbMain.AllowNavigation = False
        End If

        Dim result As String = "" 'для получения результатов событий
        Dim arrs() As String
        Dim classL As Integer = 0
        If mScript.mainClassHash.TryGetValue("L", classL) = False Then
            mScript.LAST_ERROR = "Не загружен или поврежден класс локаций."
            Return "#Error"
        End If
        Dim locId As Integer = GetSecondChildIdByName(arrParams(0), mScript.mainClass(classL).ChildProperties)
        If locId < 0 OrElse locId > mScript.mainClass(classL).ChildProperties.Count - 1 Then
            mScript.LAST_ERROR = "Локации с Id = " & locId.ToString & " не существует."
            Return "#Error"
        End If
        Dim locName As String = mScript.mainClass(classL).ChildProperties(locId)("Name").Value

        Dim doHandler As Boolean = True 'выполнять ли обработчик LocationEnterEvent
        Dim mapUsed As Boolean = False 'использована ли при переходе карта

        If arrParams.Count > 1 Then
            If String.Compare(arrParams(1), "False", True) = 0 Then doHandler = False
        End If
        If arrParams.Count > 2 Then
            If String.Compare(arrParams(2), "True", True) = 0 Then mapUsed = True
        End If

        '1. LocationExitEvent
        Dim eventId As Integer
        If GVARS.G_CURLOC > -1 AndAlso doLocationExitEvent Then
            arrs = {GVARS.G_CURLOC.ToString, locId.ToString, mapUsed.ToString}
            'событие свойств по умолчанию
            eventId = mScript.mainClass(classL).Properties("LocationExitEvent").eventId
            If eventId > 0 Then
                result = mScript.eventRouter.RunEvent(eventId, arrs, "LocationExitEvent", False)
                If result = "#Error" OrElse result = "False" Then Return result
            End If
            'событие предыдущей локации
            eventId = mScript.mainClass(classL).ChildProperties(GVARS.G_CURLOC)("LocationExitEvent").eventId
            If eventId > 0 Then
                result = mScript.eventRouter.RunEvent(eventId, arrs, "LocationExitEvent", False)
                If result = "#Error" OrElse result = "False" Then Return result
            End If
        End If

        'очищаем от содержимого предыдущей локации
        Dim hDocument As HtmlDocument = frmPlayer.wbMain.Document
        Dim hOld As HtmlElement = hDocument.GetElementById("MainConvasOld")
        If IsNothing(hOld) = False Then
            Dim hDomOld As mshtml.HTMLDivElement = hOld.DomElement
            hDomOld.removeNode(True)
        End If
        Dim hConvasOld As HtmlElement = hDocument.GetElementById("MainConvas")
        hConvasOld.Id = "MainConvasOld"
        hConvasOld.SetAttribute("ClassName", "pt-page")
        Application.DoEvents()
        Dim hConvas As HtmlElement = hDocument.CreateElement("DIV")
        hConvas.SetAttribute("ClassName", "pt-page")
        hConvas.Id = "MainConvas"
        hConvasOld.Parent.AppendChild(hConvas)
        hConvasOld.InsertAdjacentElement(HtmlElementInsertionOrientation.BeforeBegin, hConvas)
        'Application.DoEvents()


        'hConvas.Style = "background-color:red"
        hDocument.InvokeScript("initAnimation", {1})

        hConvas.InnerHtml = ""
        Dim prevLoc As Integer = GVARS.G_PREVLOC
        GVARS.G_PREVLOC = GVARS.G_CURLOC
        GVARS.G_CURLOC = locId

        'Dim hDocument As HtmlDocument = frmPlayer.wbMain.Document
        'hDocument.InvokeScript("asss", {"www"})

        If GVARS.G_CURMAP > -1 Then
            'действия с картой: получаем текущую, предыдущую клетку и убираем туман с соседних (если надо)
            Dim changeCell As Boolean = True, cellNew As Integer = -1
            Dim classMap As Integer = mScript.mainClassHash("Map")
            Dim arrMap() As String = {GVARS.G_CURMAP.ToString, GVARS.G_CURMAPCELL.ToString}

            If GVARS.G_CURMAPCELL > -1 Then
                Dim cellStr As String = ""
                ReadProperty(classMap, "Location", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, cellStr, arrMap)
                cellNew = GetSecondChildIdByName(cellStr, mScript.mainClass(classL).ChildProperties)
                If cellNew = locId Then
                    changeCell = False 'у текущей клетки карты та же самая локация - значит клетку не меняем (на несколько клеток может быть одна локация)
                    cellNew = GVARS.G_CURMAPCELL
                End If
            End If

            If changeCell Then
                cellNew = mapManager.GetCellByLocation(GVARS.G_CURMAP, GVARS.G_CURLOC)
                If cellNew > -1 Then
                    Dim hCell As HtmlElement = Nothing
                    Dim curCellClass As String = ""

                    GVARS.G_PREVMAPCELL = GVARS.G_CURMAPCELL
                    GVARS.G_CURMAPCELL = cellNew
                    mapManager.ClearNearbyFog(Nothing)
                    mapManager.ShowArrowInCurrentCell()
                End If
            End If
        End If


        'собираем параметры для передачи событию
        arrs = {locId.ToString, mapUsed.ToString}
        Dim additionalParamsCount As Integer = arrParams.Count - 3
        If additionalParamsCount > 0 Then
            ReDim Preserve arrs(1 + additionalParamsCount)
            For i As Integer = 2 To arrs.Count - 1
                arrs(i) = arrParams(i + 1)
            Next i
        End If

        'меняем цвет фона и текста главного окна, прилегание
        Dim strStyle As New System.Text.StringBuilder
        Dim bgColor As String = "", fColor As String = "", tAlign As String = "", classLW As Integer = mScript.mainClassHash("LW")
        ReadProperty(classL, "BackColor", GVARS.G_CURLOC, -1, bgColor, Nothing)
        bgColor = UnWrapString(bgColor)
        ReadProperty(classLW, "TextColor", -1, -1, fColor, Nothing)
        fColor = UnWrapString(fColor)
        ReadProperty(classLW, "Align", -1, -1, tAlign, Nothing)
        tAlign = UnWrapString(tAlign)
        If String.IsNullOrEmpty(bgColor) = False Then strStyle.Append("background:" & bgColor & ";")
        If String.IsNullOrEmpty(fColor) = False Then strStyle.Append("color:" & fColor & ";")
        If String.IsNullOrEmpty(tAlign) = False Then strStyle.Append("text-align:" & tAlign & ";")

        'Меняем фоновую картинку
        Dim bkPicture As String = ""
        ReadProperty(classL, "BackPicture", locId, -1, bkPicture, arrs)
        bkPicture = UnWrapString(bkPicture)
        Dim bkPicPos As Integer = 0, bkPicStyle As Integer = 0
        If String.IsNullOrEmpty(bkPicture) = False Then
            bkPicture = bkPicture.Replace("\"c, "/"c)
            ReadPropertyInt(classL, "BackPicPos", locId, -1, bkPicPos, arrs)
            ReadPropertyInt(classL, "BackPicStyle", locId, -1, bkPicStyle, arrs)
        End If

        If String.IsNullOrEmpty(bkPicture) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, bkPicture)) Then
            'файл картинки существует
            strStyle.AppendFormat("background-image: url({0});", "'" + bkPicture + "'")
            '0 простая загрузка, 1 - заполнить, 2 - масштабировать, 3 - размножить, 4 - размножить по Х, 5 - размножить по Y 
            strStyle.Append("background-repeat:")
            Select Case bkPicStyle
                Case 0 '0 простая загрузка
                    strStyle.Append("no-repeat;")
                Case 1 '1 растянуть пропорционально
                    strStyle.Append("no-repeat;background-size:cover;")
                Case 2 '2 заполнить
                    strStyle.Append("no-repeat;background-size:contain;")
                    'обязательно указываем высоту окна, иначе отображается некорректно
                    strStyle.AppendFormat("height:{0}px;", frmPlayer.wbMain.ClientSize.Height - CONST_WBHEIGHT_CORRECTION)
                Case 3 '3 масштабировать
                    strStyle.Append("repeat;")
                Case 4 '4 размножить по Х
                    strStyle.Append("repeat-x;")
                Case 5 '5 размножить по Y
                    strStyle.Append("repeat-y;")
            End Select
            If bkPicStyle = 0 Then
                'BackPicPos
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
            hConvas.Document.Body.Style = strStyle.ToString + "height:" & (frmPlayer.wbMain.ClientSize.Height - 20).ToString & "px;"
            strStyle.Clear()
        End If

        'сохраняем измененные действия если RetreiveInitialState = False
        If GVARS.G_PREVLOC > -1 Then
            Dim retreiveInitial As Boolean = True 'если True, то сохранять изменения в действиях не нужно
            ReadPropertyBool(classL, "RetreiveInitialState", GVARS.G_PREVLOC, -1, retreiveInitial, arrs)
            If Not retreiveInitial Then actionsRouter.SaveActions() '(mScript.mainClass(classL).ChildProperties(GVARS.G_PREVLOC)("Name").Value)
        End If

        'Очищаем действия
        Dim actConvas As HtmlElement = frmPlayer.wbActions.Document.GetElementById("ActionsConvas")
        If IsNothing(actConvas) = False Then actConvas.InnerHtml = ""

        'загружаем действия чтобы игрок смог к ним обращаться
        actionsRouter.RetreiveActions(locName)

        'Command.AutoVisible processing
        If GVARS.G_CURLOC > -1 Then
            Dim classCmd As Integer = mScript.mainClassHash("Cm")
            Dim cmAutoVis As Boolean = False
            If ReadPropertyBool(classCmd, "AutoVisible", -1, -1, cmAutoVis, Nothing) = False Then Return "#Error"
            If cmAutoVis Then
                'Command string visibility shoud be changed
                eventId = mScript.mainClass(classL).ChildProperties(GVARS.G_CURLOC)("LocationCommandEvent").eventId
                'новое значение видимости кмд
                Dim newVisible As Boolean = (eventId >= 1)

                'было ли видно окно раньше
                Dim oldVisible As Boolean = False
                'FindActionHTMLelementDOM ReadPropertyBool(classCmd,"Visible",-1,-1
                Dim wbCmd As WebBrowser = frmPlayer.wbCommand
                If wbCmd.Parent.GetType.Name = "SplitterPanel" Then
                    Dim pnl As SplitterPanel = wbCmd.Parent
                    Dim splitCont As SplitContainer = wbCmd.Parent.Parent
                    If splitCont.Visible Then
                        If Object.Equals(pnl, splitCont.Panel1) Then
                            oldVisible = Not splitCont.Panel1Collapsed
                        ElseIf Object.Equals(pnl, splitCont.Panel2) Then
                            oldVisible = Not splitCont.Panel2Collapsed
                        End If
                    End If
                End If

                'если видимость изменилась - прячем
                If newVisible <> oldVisible Then
                    If PropertiesRouter(classCmd, "Visible", Nothing, arrParams, PropertiesOperationEnum.PROPERTY_SET, newVisible.ToString) = "#Error" Then Return "#Error"
                End If
            End If
        End If

        Dim firstlyEvent As Boolean = True
        ReadPropertyBool(classL, "EventInitially", locId, -1, firstlyEvent, arrs)

        ActionsInputProhibited = True
        If firstlyEvent Then
            ''load actions to html-element before LocationEnterEvent
            'If IsNothing(actConvas) = False Then
            '    actConvas.InnerHtml = ""
            '    'Создаем действия
            '    Dim classA As Integer = mScript.mainClassHash("A")
            '    If IsNothing(mScript.mainClass(classA).ChildProperties) = False AndAlso mScript.mainClass(classA).ChildProperties.Count > 0 Then
            '        'действия на данной локации существуют. Создаем их
            '        For actId As Integer = 0 To mScript.mainClass(classA).ChildProperties.Count - 1
            '            'перебираем все действия
            '            ActionCreateHtmlElement(actId)
            '        Next actId
            '    End If
            'End If

            '2. LocationEnterEvent
            If doHandler Then
                'событие свойств по умолчанию
                eventId = mScript.mainClass(classL).Properties("LocationEnterEvent").eventId
                If eventId > 0 Then
                    result = mScript.eventRouter.RunEvent(eventId, arrs, "LocationEnterEvent", False)
                    If result = "#Error" Then ActionsInputProhibited = False : Return result
                End If
                'событие предыдущей локации
                eventId = mScript.mainClass(classL).ChildProperties(GVARS.G_CURLOC)("LocationEnterEvent").eventId
                If eventId > 0 Then
                    result = mScript.eventRouter.RunEvent(eventId, arrs, "LocationEnterEvent", False)
                    If result = "#Error" Then ActionsInputProhibited = False : Return result
                End If
            End If
        End If

        '3. Вывод описания локации
        If GVARS.G_CURLOC = locId Then 'если локация при исполнении LocationEnterEvent не изменилась, то
            'вывод описания локации
            result = ShowDescription(hConvas, classL, "Description", locId, -1, arrs)
        End If

        '4. Header (верхний колонтитул)
        Dim hHeader As HtmlElement = hConvas.Document.GetElementById("header")
        If IsNothing(hHeader) = False Then
            Dim fInner As String = hHeader.InnerHtml
            If IsNothing(fInner) Then fInner = ""
            eventId = mScript.mainClass(classL).Properties("Header").eventId
            If eventId > 0 Then
                'колонтитул задан
                ReadProperty(classL, "Header", -1, -1, result, arrs)
                If String.IsNullOrEmpty(result) = False AndAlso result <> "#Error" AndAlso result <> fInner Then
                    hHeader.InnerHtml = result
                End If
            End If
            'If String.IsNullOrEmpty(hHeader.InnerHtml) Then
            '    hHeader.Style = "display:none"
            'Else
            '    hHeader.Style = ""
            'End If
        End If

        '5. Footer (нижний колонтитул)
        Dim hFooter As HtmlElement = hConvas.Document.GetElementById("footer")
        If IsNothing(hFooter) = False Then
            Dim fInner As String = hFooter.InnerHtml
            If IsNothing(fInner) Then fInner = ""
            eventId = mScript.mainClass(classL).Properties("Footer").eventId
            If eventId > 0 Then
                'колонтитул задан
                ReadProperty(classL, "Footer", -1, -1, result, arrs)
                If String.IsNullOrEmpty(result) = False AndAlso result <> "#Error" AndAlso result <> fInner Then
                    hFooter.InnerHtml = result
                End If
            Else
                fInner = ""
                hFooter.InnerHtml = ""
            End If
            'If String.IsNullOrEmpty(hFooter.InnerHtml) Then
            '    hFooter.Style = "display:none"
            'Else
            '    hFooter.Style = ""
            'End If
        End If

        If Not firstlyEvent Then
            '2. LocationEnterEvent
            If doHandler Then
                'событие свойств по умолчанию
                eventId = mScript.mainClass(classL).Properties("LocationEnterEvent").eventId
                If eventId > 0 Then
                    result = mScript.eventRouter.RunEvent(eventId, arrs, "LocationEnterEvent", False)
                    If result = "#Error" Then ActionsInputProhibited = False : Return result
                End If
                'событие предыдущей локации
                eventId = mScript.mainClass(classL).ChildProperties(GVARS.G_CURLOC)("LocationEnterEvent").eventId
                If eventId > 0 Then
                    result = mScript.eventRouter.RunEvent(eventId, arrs, "LocationEnterEvent", False)
                    If result = "#Error" Then ActionsInputProhibited = False : Return result
                End If
            End If
        End If

        ActionsInputProhibited = False
        'load actions to html-element before LocationEnterEvent
        If IsNothing(actConvas) = False Then
            actConvas.InnerHtml = ""
            'Создаем действия
            Dim classA As Integer = mScript.mainClassHash("A")
            If IsNothing(mScript.mainClass(classA).ChildProperties) = False AndAlso mScript.mainClass(classA).ChildProperties.Count > 0 Then
                'действия на данной локации существуют. Создаем их
                For actId As Integer = 0 To mScript.mainClass(classA).ChildProperties.Count - 1
                    'перебираем все действия
                    ActionCreateHtmlElement(actId)
                Next actId
            End If
        End If

        '6. Abilities
        Dim abList As SortedList(Of Integer, Integer) = HeroGetAbilitySetsIdList() 'Key - heroId, Value - AbSetId
        If abList.Count > 0 Then
            'событие AbilityOnChangeLocEvent
            Dim classAb As Integer = mScript.mainClassHash("Ab")
            Dim globalEventId As Integer = mScript.mainClass(classAb).Properties("AbilityOnChangeLocEvent").eventId
            Dim abArrs() As String = {"", "", "", locId.ToString, mapUsed.ToString} 'параметры для AbilityOnChangeLocEvent
            ReDim Preserve abArrs(arrParams.Count + 1)
            For i = 5 To arrParams.Count + 1
                abArrs(i) = arrParams(i - 2) 'дополнительные параметры, переданные функции Go
            Next

            For i As Integer = 0 To abList.Count - 1
                Dim heroId As Integer = abList.ElementAt(i).Key
                Dim abSetId As Integer = abList.ElementAt(i).Value
                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classAb).ChildProperties(abSetId)("AbilityOnChangeLocEvent")
                Dim abSetEventId As Integer = ch.eventId
                abArrs(0) = abSetId.ToString
                abArrs(2) = heroId.ToString
                For abId As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                    'запускаем событие AbilityOnChangeLocEvent для кажодй способности из всех доступных наборов, прикрепленных к персонажам
                    abArrs(1) = abId.ToString

                    Dim isEnabled As Boolean = True
                    If ReadPropertyBool(classAb, "Enabled", abSetId, abId, isEnabled, abArrs) = False Then Return "#Error"
                    If Not isEnabled Then Continue For 'способность недоступна - пропускаем

                    If globalEventId > 0 Then
                        'глобальное событие
                        If mScript.eventRouter.RunEvent(globalEventId, abArrs, "AbilityOnChangeLocEvent", False) = "#Error" Then Return "#Error"
                    End If

                    If abSetEventId > 0 Then
                        'событие данного набора
                        If mScript.eventRouter.RunEvent(abSetEventId, abArrs, "AbilityOnChangeLocEvent", False) = "#Error" Then Return "#Error"
                    End If

                    eventId = ch.ThirdLevelEventId(abId)
                    If eventId > 0 Then
                        'событие данной способности
                        If mScript.eventRouter.RunEvent(eventId, abArrs, "AbilityOnChangeLocEvent", False) = "#Error" Then Return "#Error"
                    End If

                    'звук при активации SoundOnActivate
                    Dim snd As String = ""
                    If ReadProperty(classAb, "SoundOnActivate", abSetId, abId, snd, {abArrs(0), abArrs(1)}) = False Then Return "#Error"
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
        End If

        'увеличиваем Visits локации
        Dim visits As Integer = 0
        ReadPropertyInt(classL, "Visits", locId, -1, visits, arrs)

        'событие отслеживания свойства
        Dim visitsRes As String = (visits + 1).ToString
        Dim trackResult As String = mScript.trackingProperties.RunBefore(classL, "Visits", {locId.ToString}, visitsRes)
        If trackResult <> "False" AndAlso trackResult <> "#Error" Then
            SetPropertyValue(classL, "Visits", visitsRes, locId)
        End If

        'увеличиваем Visits клетки карты
        If GVARS.G_CURMAP > -1 AndAlso GVARS.G_CURMAPCELL > -1 Then
            Dim classMap As Integer = mScript.mainClassHash("Map")
            visits = 0
            ReadPropertyInt(classMap, "Visits", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, visits, arrs)

            'событие отслеживания свойства
            visitsRes = (visits + 1).ToString
            trackResult = mScript.trackingProperties.RunBefore(classMap, "Visits", {GVARS.G_CURMAP.ToString, GVARS.G_CURMAPCELL.ToString}, visitsRes)
            If trackResult <> "False" AndAlso trackResult <> "#Error" Then
                SetPropertyValue(classMap, "Visits", visitsRes, GVARS.G_CURMAP, GVARS.G_CURMAPCELL)
            End If
        End If

        'прячем / отображает footer / header
        If IsNothing(hHeader) = False Then
            If String.IsNullOrEmpty(hHeader.InnerHtml) Then
                hHeader.Style = "display:none"
            Else
                hHeader.Style = ""
            End If
        End If
        If IsNothing(hFooter) = False Then
            If String.IsNullOrEmpty(hFooter.InnerHtml) Then
                hFooter.Style = "display:none"
            Else
                hFooter.Style = ""
            End If
        End If

        Application.DoEvents()
        hDocument.InvokeScript("NextPage", {1})
        Application.DoEvents()

        'начинаем отсчет времени на данной локации
        GVARS.TIME_IN_THIS_LOCATION = Date.Now.Ticks

        Return result
    End Function

    ''' <summary>Проверяет условие отображения всех действий</summary>
    Public Sub ActionCheckShowQueries()
        Dim classA As Integer = mScript.mainClassHash("A")
        If IsNothing(mScript.mainClass(classA).ChildProperties) OrElse mScript.mainClass(classA).ChildProperties.Count = 0 Then Return
        Dim result As String = ""

        Dim mainEventId As Integer = mScript.mainClass(classA).Properties("ShowQuery").eventId
        For actId As Integer = 0 To mScript.mainClass(classA).ChildProperties.Count - 1
            'получаем видимость
            Dim actVisible As Boolean = True, actArrs() As String = {actId.ToString, "-1"}
            ReadPropertyBool(classA, "Visible", actId, -1, actVisible, actArrs)

            'Выполняем результат A.ShowQuery всех действий
            Dim showQAll As Boolean = True
            If actVisible Then
                If mainEventId > 0 Then
                    result = mScript.eventRouter.RunEvent(mainEventId, actArrs, "Action.ShowQuery", False)
                    Boolean.TryParse(result, showQAll)
                End If
            End If
            '... текущего действия
            Dim showQ As Boolean = True
            If actVisible AndAlso showQAll Then
                Dim eventId As Integer = mScript.mainClass(classA).ChildProperties(actId)("ShowQuery").eventId
                If eventId > 0 Then
                    result = mScript.eventRouter.RunEvent(eventId, actArrs, "Action.ShowQuery", False)
                    Boolean.TryParse(result, showQ)
                End If
            End If
            'получаем надо ли фактически отображать действие
            Dim isShown As Boolean = (actVisible = True AndAlso showQ = True AndAlso showQAll = True)

            Dim hEl As HtmlElement = FindActionHTMLelement(classA, actId, actArrs)
            If IsNothing(hEl) Then Continue For

            If isShown Then
                'действие открыто
                hEl.Style = ""
            Else
                'действие скрыто
                hEl.Style = "display:none"
            End If

        Next
    End Sub

    ''' <summary>
    ''' Создает html-элемент действия
    ''' </summary>
    ''' <param name="actId">Id действия</param>
    ''' <param name="elementToPlace">html-элемент, перед котрым надо вставить действие</param>
    Public Function ActionCreateHtmlElement(ByVal actId As Integer, Optional ByRef elementToPlace As HtmlElement = Nothing) As HtmlElement
        If ActionsInputProhibited Then Return Nothing
        Dim actArrs() As String = {actId.ToString}
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim eventId As Integer, result As String
        'получаем html-котейнер для размещения свойства
        Dim actContainer As String = ""
        ReadProperty(classA, "HTMLContainerId", actId, -1, actContainer, actArrs)
        actContainer = UnWrapString(actContainer)
        Dim htmlContainer As HtmlElement
        If String.IsNullOrEmpty(actContainer) Then
            htmlContainer = frmPlayer.wbActions.Document.GetElementById("ActionsConvas")
        Else
            htmlContainer = frmPlayer.wbMain.Document.GetElementById(actContainer)
            If IsNothing(htmlContainer) Then htmlContainer = frmPlayer.wbMain.Document.GetElementById(actContainer)
            If IsNothing(htmlContainer) Then
                MessageBox.Show("HTML element с Id " & actContainer & " не найден. Действие будет вставлено в главное окно.", "Matew quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                htmlContainer = frmPlayer.wbMain.Document.GetElementById("MainConvas")
            End If
        End If
        'получаем Visible
        Dim actVisible As Boolean = True
        ReadPropertyBool(classA, "Visible", actId, -1, actVisible, actArrs)
        'Выполняем результат A.ShowQuery всех действий
        Dim showQAll As Boolean = True
        If actVisible Then
            eventId = mScript.mainClass(classA).Properties("ShowQuery").eventId
            If eventId > 0 Then
                result = mScript.eventRouter.RunEvent(eventId, actArrs, "Action.ShowQuery", False)
                Boolean.TryParse(result, showQAll)
            End If
        End If
        '... текущего действия
        Dim showQ As Boolean = True
        If actVisible AndAlso showQAll Then
            eventId = mScript.mainClass(classA).ChildProperties(actId)("ShowQuery").eventId
            If eventId > 0 Then
                result = mScript.eventRouter.RunEvent(eventId, actArrs, "Action.ShowQuery", False)
                Boolean.TryParse(result, showQ)
            End If
        End If
        'получаем надо ли фактически отображать действие
        Dim isShown As Boolean = (actVisible = True AndAlso showQ = True AndAlso showQAll = True)
        'получаем Caption
        Dim actCaption As String = ""
        ReadProperty(classA, "Caption", actId, -1, actCaption, actArrs)
        actCaption = UnWrapString(actCaption)
        'получаем Enabled
        Dim actEnabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, actEnabled, actArrs)
        'получаем Picture
        Dim actPicture As String = ""
        ReadProperty(classA, "Picture", actId, -1, actPicture, actArrs)
        actPicture = UnWrapString(actPicture).Replace("\", "/")
        'получаем PictureFloat
        Dim actPictureFloat As Integer = 0 '0 - без прилегания, 1 - лево, 2 - право
        If String.IsNullOrEmpty(actPicture) = False Then ReadPropertyInt(classA, "PictureFloat", actId, -1, actPictureFloat, actArrs)
        'получаем имя
        Dim actName As String = mScript.mainClass(classA).ChildProperties(actId)("Name").Value
        'получаем тип свойства
        '0 - кнопка, 1 - изображение, 2 - ссылка, 3 - разделитель, 4 - текст
        Dim actType As Integer = 0
        ReadPropertyInt(classA, "Type", actId, -1, actType, actArrs)
        Dim hEl As HtmlElement
        Select Case actType
            Case 1 'img
                hEl = htmlContainer.Document.CreateElement("IMG")
                hEl.SetAttribute("Src", actPicture)
                hEl.SetAttribute("Title", actCaption)
                hEl.SetAttribute("ClassName", "ActionImage")
            Case 2 'anchor
                hEl = htmlContainer.Document.CreateElement("A")
                hEl.SetAttribute("ClassName", "ActionAnchor")
                hEl.SetAttribute("href", "#")
            Case 3 'hr
                hEl = htmlContainer.Document.CreateElement("HR")
                hEl.SetAttribute("ClassName", "ActionHR")
            Case 4 'text
                hEl = htmlContainer.Document.CreateElement("DIV")
                hEl.SetAttribute("ClassName", "ActionText")
            Case Else '0 - button
                hEl = htmlContainer.Document.CreateElement("DIV")
                hEl.SetAttribute("ClassName", "ActionButton")
        End Select

        If Not isShown Then
            'действие скрыто
            hEl.Style = "display:none"
        End If

        'enabled
        If Not actEnabled Then
            hEl.SetAttribute("disabled", "True")
        End If

        'получаем текст действия
        ActionSetInnerHTML(hEl, actType, actCaption, actPicture, actPictureFloat)
        'hEl.SetAttribute("Name", actName)
        hEl.Name = actName

        'создаем событие клика
        If actType = 0 Then 'Button
            AddHandler hEl.MouseDown, AddressOf del_action_MouseDown_Button
            AddHandler hEl.MouseUp, AddressOf del_action_MouseUp_Button
        ElseIf actType = 1 Then 'Image
            AddHandler hEl.MouseDown, AddressOf del_action_MouseDown_Image
            AddHandler hEl.MouseUp, AddressOf del_action_MouseUp_Image
        ElseIf actType = 2 Then 'Anchor
            AddHandler hEl.MouseDown, AddressOf del_action_MouseDown_Anchor
            AddHandler hEl.MouseUp, AddressOf del_action_MouseUp_Anchor
        End If

        If actType <= 2 Then 'Button, Image, Anchor
            AddHandler hEl.Click, AddressOf del_action_Click
            AddHandler hEl.MouseOver, AddressOf del_action_MouseOver
            AddHandler hEl.MouseLeave, AddressOf del_action_MouseLeave
        End If

        'добавляем действие
        If IsNothing(elementToPlace) Then
            htmlContainer.AppendChild(hEl)
        Else
            elementToPlace.InsertAdjacentElement(HtmlElementInsertionOrientation.BeforeBegin, hEl)
        End If
        Return hEl
    End Function

    ''' <summary>
    ''' Создает надпись для отображения действий типа button, anchor, text
    ''' </summary>
    ''' <param name="hAction">ссылка на html-элемент действия</param>
    ''' <param name="actType">тия действия</param>
    ''' <param name="actCaption">подпись без кавычек</param>
    ''' <param name="actPicture">путь к картинке относительно папки квеста, без кавычек и в правильном формате</param>
    ''' <param name="actPictureFloat">прилегание картинки</param>
    Public Sub ActionSetInnerHTML(ByRef hAction As HtmlElement, ByVal actType As Integer, ByVal actCaption As String, ByVal actPicture As String, ByVal actPictureFloat As Integer)
        'получаем текст действия
        Dim hInner As New System.Text.StringBuilder
        If actType <> 3 AndAlso actType <> 1 Then
            'вставляем картинку если тип button, anchor или text
            If String.IsNullOrEmpty(actPicture) = False Then
                hInner.Append("<img src='" & actPicture)
                If actPictureFloat = 1 Then
                    hInner.Append("' style='float:left'/>")
                ElseIf actPictureFloat = 2 Then
                    hInner.Append("' style='float:right'/>")
                Else
                    hInner.Append("'/>")
                End If
            End If
            If String.IsNullOrEmpty(actCaption) = False Then actCaption = "<span>" & actCaption & "</span>"
            hInner.Append(actCaption) 'вставляем сам текст
            hAction.InnerHtml = hInner.ToString
            hInner.Clear()
        End If
    End Sub

    Public Sub del_action_Click(sender As HtmlElement, e As HtmlElementEventArgs)
        'событие выбора действия
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim actId As Integer = GetSecondChildIdByName(sender.Name, mScript.mainClass(classA).ChildProperties)
        Dim arrs() As String = {actId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, enabled, arrs)
        If enabled = False Then Return

        'Visits + 1
        Dim visits As Integer = 0
        ReadPropertyInt(classA, "Visits", actId, -1, visits, arrs)
        visits += 1
        Dim visitsRes As String = visits.ToString
        Dim trackingResult As String = mScript.trackingProperties.RunBefore(classA, "Visits", arrs, visitsRes)
        If trackingResult <> "False" AndAlso trackingResult <> "#Error" Then
            mScript.mainClass(classA).ChildProperties(actId)("Visits").Value = visitsRes
        End If

        'изменяем Visits в сохраненных действиях
        Dim classL As Integer = mScript.mainClassHash("L")
        Dim curLocName As String = mScript.mainClass(classL).ChildProperties(GVARS.G_CURLOC)("Name").Value
        If actionsRouter.lstActions.ContainsKey(curLocName) AndAlso IsNothing(actionsRouter.lstActions(curLocName)) = False AndAlso actId <= actionsRouter.lstActions(curLocName).Count - 1 Then
            actionsRouter.lstActions(curLocName)(actId)("Visits").Value = visits.ToString
        End If

        'Прячем / делаем недоступным если надо
        Dim afterVisiting As Integer = 0
        ReadPropertyInt(classA, "AfterVisiting", actId, -1, afterVisiting, arrs)
        If afterVisiting = 1 Then
            'скрыть
            PropertiesRouter(classA, "Visible", arrs, arrs, PropertiesOperationEnum.PROPERTY_SET, "False")
        ElseIf afterVisiting = 2 Then
            'сделать недоступным
            PropertiesRouter(classA, "Enabled", arrs, arrs, PropertiesOperationEnum.PROPERTY_SET, "False")
        End If

        Dim initLocation As Integer = GVARS.G_CURLOC
        'событие выбора действия (глобальное)
        Dim wasEvent As Boolean = False
        Dim eventId As Integer = mScript.mainClass(classA).Properties("ActionSelectEvent").eventId
        If eventId > 0 Then
            mScript.eventRouter.RunEvent(eventId, arrs, "ActionSelectEvent", False)
            If mScript.LAST_ERROR.Length > 0 Then Return
            wasEvent = True
        End If

        'событие выбора действия
        eventId = mScript.mainClass(classA).ChildProperties(actId)("ActionSelectEvent").eventId
        If eventId > 0 Then
            mScript.eventRouter.RunEvent(eventId, arrs, "ActionSelectEvent", False)
            If mScript.LAST_ERROR.Length > 0 Then Return
            wasEvent = True
        End If

        'переход на новую локацию посредством свойства GoTo
        If GVARS.G_CURLOC = initLocation Then
            'если в скрипте произошла смена локации, то A.GoTo уже не выполняем
            Dim newLoc As String = ""
            ReadProperty(classA, "GoTo", actId, -1, newLoc, arrs)
            If String.IsNullOrEmpty(newLoc) = False Then
                Dim locId As Integer = GetSecondChildIdByName(newLoc, mScript.mainClass(classL).ChildProperties)
                If locId >= 0 Then
                    Go({locId.ToString, "True", "False"})
                    wasEvent = True
                End If
            End If
        End If

        If wasEvent AndAlso EventGeneratedFromScript = False Then mScript.eventRouter.RunScriptFinishedEvent(arrs)
        EventGeneratedFromScript = False
    End Sub

    Private Sub del_action_MouseOver(sender As HtmlElement, e As HtmlElementEventArgs)
        'Выводим картинку действия при наведении PictureHover
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim actId As Integer = GetSecondChildIdByName(sender.Name, mScript.mainClass(classA).ChildProperties)
        Dim arrs() As String = {actId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, enabled, arrs)
        If enabled = False Then Return

        'получаем PictureHover. Если пусто, то выход
        Dim actPictureHover As String = ""
        If ReadProperty(classA, "PictureHover", actId, -1, actPictureHover, arrs) = False Then Return
        actPictureHover = UnWrapString(actPictureHover)
        If String.IsNullOrEmpty(actPictureHover) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, actPictureHover)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'получаем Picture. PictureHover и Picture не должны быть равны
        Dim actPicture As String = ""
        If ReadProperty(classA, "Picture", actId, -1, actPicture, arrs) = False Then Return
        actPicture = UnWrapString(actPicture)

        'получаем тип. Подходят только типы кнопка, картинка и ссылка
        Dim actType As Integer = 0
        ReadPropertyInt(classA, "Type", actId, -1, actType, arrs)
        If actType > 2 Then Return 'для типов сеператор и простой текст не актуально

        'убеждаемсяв правильном формате пути к картинкам
        actPictureHover = actPictureHover.Replace("\", "/")
        If String.IsNullOrEmpty(actPicture) = False Then actPicture = actPicture.Replace("\", "/")
        If String.Compare(actPicture, actPictureHover, True) = 0 Then Return

        'собственно замена картинки Picture на PictureHover
        If actType = 1 Then
            'type = картинка
            sender.SetAttribute("src", actPictureHover)
        ElseIf sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'type = кнопка или ссылка (есть IMG)
            sender.Children(0).SetAttribute("src", actPictureHover)
        Else
            'type = кнопка или ссылка (нет IMG)
            Dim actPictureFloat As Integer = 0
            ReadPropertyInt(classA, "PictureFloat", actId, -1, actPictureFloat, arrs)
            Dim actCaption As String = ""
            ReadProperty(classA, "Caption", actId, -1, actCaption, arrs)
            ActionSetInnerHTML(sender, actType, UnWrapString(actCaption), actPictureHover, actPictureFloat)
        End If

    End Sub

    Private Sub del_action_MouseLeave(sender As HtmlElement, e As HtmlElementEventArgs)
        'возвращаем картинку действия Picture после того, как она была зменена на PictureHover
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim actId As Integer = GetSecondChildIdByName(sender.Name, mScript.mainClass(classA).ChildProperties)
        Dim arrs() As String = {actId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, enabled, arrs)
        If enabled = False Then Return

        'получаем PictureHover. Если пусто, то выход
        Dim actPictureHover As String = ""
        If ReadProperty(classA, "PictureHover", actId, -1, actPictureHover, arrs) = False Then Return
        actPictureHover = UnWrapString(actPictureHover)
        If String.IsNullOrEmpty(actPictureHover) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, actPictureHover)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            'MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'получаем Picture. PictureHover и Picture не должны быть равны
        Dim actPicture As String = ""
        If ReadProperty(classA, "Picture", actId, -1, actPicture, arrs) = False Then Return
        actPicture = UnWrapString(actPicture)

        'получаем тип. Подходят только типы кнопка, картинка и ссылка
        Dim actType As Integer = 0
        ReadPropertyInt(classA, "Type", actId, -1, actType, arrs)
        If actType > 2 Then Return 'для типов сеператор и простой текст не актуально

        'убеждаемся в правильном формате пути к картинкам
        actPictureHover = actPictureHover.Replace("\", "/")
        If String.IsNullOrEmpty(actPicture) = False Then actPicture = actPicture.Replace("\", "/")
        If String.Compare(actPicture, actPictureHover, True) = 0 Then Return

        'собственно замена картинки PictureHover на Picture
        If actType = 1 Then
            'type = картинка
            sender.SetAttribute("src", actPicture)
        ElseIf String.IsNullOrEmpty(actPicture) = True AndAlso sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'type = кнопка или ссылка (есть IMG)
            sender.Children(0).SetAttribute("src", actPicture)
        Else
            'type = кнопка или ссылка (нет IMG)
            Dim actPictureFloat As Integer = 0
            ReadPropertyInt(classA, "PictureFloat", actId, -1, actPictureFloat, arrs)
            Dim actCaption As String = ""
            ReadProperty(classA, "Caption", actId, -1, actCaption, arrs)
            ActionSetInnerHTML(sender, actType, UnWrapString(actCaption), actPicture, actPictureFloat)
        End If

    End Sub

    Private Sub del_action_MouseDown_Button(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim actId As Integer = GetSecondChildIdByName(sender.Name, mScript.mainClass(classA).ChildProperties)
        Dim arrs() As String = {actId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, enabled, arrs)
        If enabled = False Then Return

        HTMLReplaceClass(sender, "ActionButton", "ActionButtonMouseDown")

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim actPictureActive As String = ""
        If ReadProperty(classA, "PictureActive", actId, -1, actPictureActive, arrs) = False Then Return
        actPictureActive = UnWrapString(actPictureActive)
        If String.IsNullOrEmpty(actPictureActive) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, actPictureActive)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        actPictureActive = actPictureActive.Replace("\", "/")

        'собственно замена картинки PictureActive
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", actPictureActive)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim actType As Integer = 0
            ReadPropertyInt(classA, "Type", actId, -1, actType, arrs)
            Dim actPictureFloat As Integer = 0
            ReadPropertyInt(classA, "PictureFloat", actId, -1, actPictureFloat, arrs)
            Dim actCaption As String = ""
            ReadProperty(classA, "Caption", actId, -1, actCaption, arrs)
            ActionSetInnerHTML(sender, actType, UnWrapString(actCaption), actPictureActive, actPictureFloat)
        End If

    End Sub

    Private Sub del_action_MouseUp_Button(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim actId As Integer = GetSecondChildIdByName(sender.Name, mScript.mainClass(classA).ChildProperties)
        Dim arrs() As String = {actId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, enabled, arrs)
        If enabled = False Then Return

        HTMLReplaceClass(sender, "ActionButtonMouseDown", "ActionButton")

        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim actPictureActive As String = ""
        If ReadProperty(classA, "PictureActive", actId, -1, actPictureActive, arrs) = False Then Return
        actPictureActive = UnWrapString(actPictureActive)
        If String.IsNullOrEmpty(actPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim actPicture As String = ""
        If ReadProperty(classA, "PictureHover", actId, -1, actPicture, arrs) = False Then Return
        actPicture = UnWrapString(actPicture)
        If String.IsNullOrEmpty(actPicture) Then
            If ReadProperty(classA, "Picture", actId, -1, actPicture, arrs) = False Then Return
            actPicture = UnWrapString(actPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, actPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        actPicture = actPicture.Replace("\", "/")

        'собственно замена картинки PictureActive на PictureHover/Picture
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", actPicture)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim actType As Integer = 0
            ReadPropertyInt(classA, "Type", actId, -1, actType, arrs)
            Dim actPictureFloat As Integer = 0
            ReadPropertyInt(classA, "PictureFloat", actId, -1, actPictureFloat, arrs)
            Dim actCaption As String = ""
            ReadProperty(classA, "Caption", actId, -1, actCaption, arrs)
            ActionSetInnerHTML(sender, actType, UnWrapString(actCaption), actPicture, actPictureFloat)
        End If
    End Sub

    Private Sub del_action_MouseDown_Image(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim actId As Integer = GetSecondChildIdByName(sender.Name, mScript.mainClass(classA).ChildProperties)
        Dim arrs() As String = {actId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, enabled, arrs)
        If enabled = False Then Return

        HTMLReplaceClass(sender, "ActionImage", "ActionImageMouseDown")

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim actPictureActive As String = ""
        If ReadProperty(classA, "PictureActive", actId, -1, actPictureActive, arrs) = False Then Return
        actPictureActive = UnWrapString(actPictureActive)
        If String.IsNullOrEmpty(actPictureActive) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, actPictureActive)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        actPictureActive = actPictureActive.Replace("\", "/")

        'собственно замена картинки PictureActive
        sender.SetAttribute("src", actPictureActive)
    End Sub

    Private Sub del_action_MouseUp_Image(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim actId As Integer = GetSecondChildIdByName(sender.Name, mScript.mainClass(classA).ChildProperties)
        Dim arrs() As String = {actId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, enabled, arrs)
        If enabled = False Then Return

        HTMLReplaceClass(sender, "ActionImageMouseDown", "ActionImage")
        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim actPictureActive As String = ""
        If ReadProperty(classA, "PictureActive", actId, -1, actPictureActive, arrs) = False Then Return
        actPictureActive = UnWrapString(actPictureActive)
        If String.IsNullOrEmpty(actPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim actPicture As String = ""
        If ReadProperty(classA, "PictureHover", actId, -1, actPicture, arrs) = False Then Return
        actPicture = UnWrapString(actPicture)
        If String.IsNullOrEmpty(actPicture) Then
            If ReadProperty(classA, "Picture", actId, -1, actPicture, arrs) = False Then Return
            actPicture = UnWrapString(actPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, actPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        actPicture = actPicture.Replace("\", "/")

        'собственно замена картинки PictureActive
        sender.SetAttribute("src", actPicture)
    End Sub

    Private Sub del_action_MouseDown_Anchor(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim actId As Integer = GetSecondChildIdByName(sender.Name, mScript.mainClass(classA).ChildProperties)
        Dim arrs() As String = {actId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, enabled, arrs)
        If enabled = False Then Return

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim actPictureActive As String = ""
        If ReadProperty(classA, "PictureActive", actId, -1, actPictureActive, arrs) = False Then Return
        actPictureActive = UnWrapString(actPictureActive)
        If String.IsNullOrEmpty(actPictureActive) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, actPictureActive)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        actPictureActive = actPictureActive.Replace("\", "/")

        'собственно замена картинки PictureActive
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", actPictureActive)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim actType As Integer = 0
            ReadPropertyInt(classA, "Type", actId, -1, actType, arrs)
            Dim actPictureFloat As Integer = 0
            ReadPropertyInt(classA, "PictureFloat", actId, -1, actPictureFloat, arrs)
            Dim actCaption As String = ""
            ReadProperty(classA, "Caption", actId, -1, actCaption, arrs)
            ActionSetInnerHTML(sender, actType, UnWrapString(actCaption), actPictureActive, actPictureFloat)
        End If

    End Sub

    Private Sub del_action_MouseUp_Anchor(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim actId As Integer = GetSecondChildIdByName(sender.Name, mScript.mainClass(classA).ChildProperties)
        Dim arrs() As String = {actId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classA, "Enabled", actId, -1, enabled, arrs)
        If enabled = False Then Return

        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim actPictureActive As String = ""
        If ReadProperty(classA, "PictureActive", actId, -1, actPictureActive, arrs) = False Then Return
        actPictureActive = UnWrapString(actPictureActive)
        If String.IsNullOrEmpty(actPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim actPicture As String = ""
        If ReadProperty(classA, "PictureHover", actId, -1, actPicture, arrs) = False Then Return
        actPicture = UnWrapString(actPicture)
        If String.IsNullOrEmpty(actPicture) Then
            If ReadProperty(classA, "Picture", actId, -1, actPicture, arrs) = False Then Return
            actPicture = UnWrapString(actPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, actPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        actPicture = actPicture.Replace("\", "/")

        'собственно замена картинки PictureActive на PictureHover/Picture
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", actPicture)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim actType As Integer = 0
            ReadPropertyInt(classA, "Type", actId, -1, actType, arrs)
            Dim actPictureFloat As Integer = 0
            ReadPropertyInt(classA, "PictureFloat", actId, -1, actPictureFloat, arrs)
            Dim actCaption As String = ""
            ReadProperty(classA, "Caption", actId, -1, actCaption, arrs)
            ActionSetInnerHTML(sender, actType, UnWrapString(actCaption), actPicture, actPictureFloat)
        End If
    End Sub

    ''' <summary>
    ''' Выводит длинный текст по типу описания локации с учетом результата события выбора селектора.
    ''' </summary>
    ''' <param name="hConvas">Html-элемент, в котором размещать текст описания</param>
    ''' <param name="classId">Id класса, хранящего данное свойство-селектор</param>
    ''' <param name="propertyName">Имя свойства с длинным текстом</param>
    ''' <param name="child2Id">id элемента 2 уровня (если вовдится текст не по умолчанию), иначе -1</param>
    ''' <param name="child3Id">id элемента 3 уровня или -1, если выводится текст элемента 2 уровня или по умолчанию</param>
    ''' <param name="arrParams">список параметром, переданных функции, вызывающей данную процедуру (напр., функции L.Go)</param>
    ''' <param name="elementIdForDescription">Id элемента, в котором расположить текст описания. Если не указано, то будет создан новый с Id DescriptionText</param>
    ''' <returns>Пустую строку или #Error при ошибке</returns>
    Public Function ShowDescription(ByRef hConvas As HtmlElement, ByVal classId As Integer, ByVal propertyName As String, ByVal child2Id As Integer, ByVal child3Id As Integer, _
                                    ByRef arrParams() As String, Optional ByVal selector As Integer = -1, Optional elementIdForDescription As String = "") As String
        Dim result As String = "", retFormat As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL
        Dim strText As New System.Text.StringBuilder

        'читаем по умолчанию
        If ReadProperty(classId, "Description", -1, -1, result, arrParams, retFormat, selector) = False Then Return "#Error"
        If retFormat = MatewScript.ReturnFormatEnum.TO_STRING Then result = UnWrapString(result)
        If String.IsNullOrEmpty(result) = False Then strText.Append(result)

        'читаем второй уровень
        If child2Id >= 0 Then
            If ReadProperty(classId, "Description", child2Id, -1, result, arrParams, retFormat, selector) = False Then Return "#Error"
            If retFormat = MatewScript.ReturnFormatEnum.TO_STRING Then result = UnWrapString(result)
            If String.IsNullOrEmpty(result) = False Then strText.Append(result)
        End If

        'читаем третий уровень
        If child3Id >= 0 Then
            If ReadProperty(classId, "Description", child2Id, child3Id, result, arrParams, retFormat, selector) = False Then Return "#Error"
            If retFormat = MatewScript.ReturnFormatEnum.TO_STRING Then result = UnWrapString(result)
            If String.IsNullOrEmpty(result) = False Then strText.Append(result)
        End If

        'создаем html-элемент div, в котором будем размещать текст
        Dim newEl As HtmlElement = Nothing ' 
        If String.IsNullOrEmpty(elementIdForDescription) Then
            newEl = hConvas.Document.CreateElement("DIV")
            newEl.Id = "DescriptionText"
        Else
            newEl = hConvas.Document.GetElementById(elementIdForDescription)
            If IsNothing(newEl) Then
                newEl = hConvas.Document.CreateElement("DIV")
                newEl.Id = elementIdForDescription
            End If
        End If
        newEl.InnerHtml = strText.ToString
        hConvas.AppendChild(newEl)
        strText.Clear()

        'позиционируем и устанавливаем класс блока описания
        Dim dLeft As String = "", dTop As String = "", dWidth As String = "", dHeight As String = "", dClass As String = "", dBackColor As String = "", dForeColor As String = ""
        Dim strStyle As New System.Text.StringBuilder
        Dim wasError As Boolean = False
        If ReadProperty(classId, "OffsetByX", child2Id, child3Id, dLeft, arrParams) = False Then wasError = True
        If ReadProperty(classId, "OffsetByY", child2Id, child3Id, dTop, arrParams) = False Then wasError = True
        If ReadProperty(classId, "DescriptionWidth", child2Id, child3Id, dWidth, arrParams) = False Then wasError = True
        If ReadProperty(classId, "DescriptionHeight", child2Id, child3Id, dHeight, arrParams) = False Then wasError = True
        If ReadProperty(classId, "DescriptionCssClass", child2Id, child3Id, dClass, arrParams) = False Then wasError = True
        If ReadProperty(classId, "DescriptionBackColor", child2Id, child3Id, dBackColor, arrParams) = False Then wasError = True
        If ReadProperty(classId, "DescriptionTextColor", child2Id, child3Id, dForeColor, arrParams) = False Then wasError = True
        If wasError Then mScript.LAST_ERROR = ""

        dLeft = UnWrapString(dLeft)
        dTop = UnWrapString(dTop)
        dWidth = UnWrapString(dWidth)
        dHeight = UnWrapString(dHeight)
        dBackColor = UnWrapString(dBackColor)
        dForeColor = UnWrapString(dForeColor)
        dClass = UnWrapString(dClass)
        If String.IsNullOrEmpty(dClass) = False Then newEl.SetAttribute("ClassName", dClass)

        If String.IsNullOrEmpty(dLeft) = False Then
            If IsNumeric(dLeft) Then dLeft &= "px"
            strStyle.Append("left:" & dLeft & ";")
        End If
        If String.IsNullOrEmpty(dTop) = False Then
            If IsNumeric(dTop) Then dTop &= "px"
            strStyle.Append("top:" & dTop & ";")
        End If
        If String.IsNullOrEmpty(dWidth) = False Then
            If IsNumeric(dWidth) Then dWidth &= "px"
            strStyle.Append("width:" & dWidth & "!important;")
        ElseIf String.IsNullOrEmpty(dLeft) = False Then
            'автопределение ширины
            strStyle.Append("width:calc(100% - " & dLeft & " - 20px)!important;")
        End If
        If String.IsNullOrEmpty(dHeight) = False Then
            If IsNumeric(dHeight) Then dHeight &= "px"
            strStyle.Append("height:" & dHeight & ";overflow:auto;")
        End If
        If String.IsNullOrEmpty(dBackColor) = False Then
            strStyle.Append("background:" & dBackColor & ";")
        End If
        If String.IsNullOrEmpty(dForeColor) = False Then
            strStyle.Append("color:" & dForeColor & ";")
        End If
        If strStyle.Length > 0 Then
            strStyle.Append("position:relative")
            newEl.Style = strStyle.ToString
            strStyle.Clear()
        End If

        'If child2Id = -1 Then
        '    newEl.SetAttribute("ClassName", propertyName & "Level1")
        'ElseIf child3Id = -1 Then
        '    newEl.SetAttribute("ClassName", propertyName & "Level2")
        'Else
        '    newEl.SetAttribute("ClassName", propertyName & "Level3")
        'End If
        'вводим текст в созданный div
        'newEl.InnerHtml = result

        Return ""
er:
        Return "#Error"
    End Function

    ''' <summary>
    ''' Возвращает количество селекторов в блоке LONG_TEXT.
    ''' </summary>
    ''' <param name="exCode">Ссылка на структуру ExecuteDataType c LONG_TEXT - описанием локации</param>
    Public Function GetSelectorsCount(ByRef exCode As List(Of MatewScript.ExecuteDataType)) As Integer
        If IsNothing(exCode) OrElse exCode.Count = 0 Then Return 0

        Dim cnt As Integer = 0
        For lineId As Integer = 0 To exCode.Count - 1
            Dim curCode() As CodeTextBox.EditWordType = exCode(lineId).Code
            If IsNothing(curCode) OrElse curCode.Length = 0 Then Continue For
            If curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA AndAlso curCode(0).Word.StartsWith("#") AndAlso mScript.IsItSelector(curCode(0).Word, 0) > 0 Then
                cnt += 1
            End If
        Next lineId
        Return cnt
    End Function
#End Region

#Region "ReadProperty"
    ''' <summary>
    ''' Функция получаем итоговое значение свойства. Если оно содержит скрипт или командную строку, то возвращается их результат, если длинный текст, то возвращает текст, отобранный селекторм.
    ''' Иначе возвращает простое значение свойства. 'Кавычки' со строк не убираются (а результат кода дается без кавычек)
    ''' </summary>
    ''' <param name="classId">Id класса свойства</param>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="child2Id">Id элемента 2 уровня</param>
    ''' <param name="child3Id">Id элемента 3 уровня</param>
    ''' <param name="result">ссылка для получения значения свойства</param>
    ''' <param name="arrParams">массив параметров на случай, если его надо будет передать скрипту</param>
    ''' <param name="retFormat">для получения формата свойства</param>
    ''' <param name="selector">номер селектора для свойств Description</param>
    ''' <param name="ignoreBattle">Если True, то поиск будет производиться в классе Hero даже если идет битва</param>
    ''' <returns>False если ошибка</returns>
    Public Function ReadProperty(ByVal classId As Integer, ByVal propName As String, ByVal child2Id As Integer, ByVal child3Id As Integer, ByRef result As String, ByRef arrParams() As String, _
                                 Optional ByRef retFormat As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL, Optional selector As Integer = -1, _
                                 Optional ByVal ignoreBattle As Boolean = False) As Boolean
        On Error GoTo er
        retFormat = MatewScript.ReturnFormatEnum.ORIGINAL

        If GVARS.G_ISBATTLE AndAlso ignoreBattle = False AndAlso mScript.mainClass(classId).Names(0) = "H" AndAlso child2Id > -1 Then
            Return mScript.Battle.ReadFighterProperty(propName, child2Id, result, arrParams, retFormat, selector)
        End If

        Dim prop As MatewScript.PropertiesInfoType = Nothing
        If mScript.mainClass(classId).Properties.TryGetValue(propName, prop) = False Then
            mScript.LAST_ERROR = String.Format("Свойство {0} в классе {1} не найдено.", propName, mScript.mainClass(classId).Names.Last)
            Return False
        End If

        Dim value As String, eventId As Integer
        If child2Id < 0 Then
            eventId = prop.eventId
            value = prop.Value 'если eventId > 0, то свойство типа скрипт/длинный текст
        ElseIf child3Id < 0 Then
            If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse child2Id > mScript.mainClass(classId).ChildProperties.Count - 1 Then
                mScript.LAST_ERROR = String.Format("Элемента с Id = {0} в классе {1}  не существует.", child2Id.ToString, mScript.mainClass(classId).Names.Last)
                Return False
            End If
            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(propName)
            eventId = ch.eventId
            value = ch.Value
        Else
            If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse child2Id > mScript.mainClass(classId).ChildProperties.Count - 1 Then
                mScript.LAST_ERROR = String.Format("Элемента с Id = {0} в классе {1}  не существует.", child2Id.ToString, mScript.mainClass(classId).Names.Last)
                Return False
            End If
            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(propName)
            If IsNothing(ch.ThirdLevelProperties) OrElse child3Id > ch.ThirdLevelProperties.Count - 1 Then
                mScript.LAST_ERROR = String.Format("Элемента третьего порядка с Id = {0} у родителя с Id {1} в классе {2} не существует.", child3Id.ToString, child2Id.ToString, mScript.mainClass(classId).Names.Last)
                Return False
            End If
            eventId = ch.ThirdLevelEventId(child3Id)
            value = ch.ThirdLevelProperties(child3Id)
        End If

        If eventId > 0 Then
            Dim codeType As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(value)
            If codeType = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                'тип - длинный текст

                If selector = -1 Then
                    'селектор не указан
                    'получаем событие селектора
                    Dim propSelector As String = propName & "Selector"
                    'получаем eventId события выбора селектора данного длинного текста (если такое свойство есть)
                    Dim selectorEventId As Integer = 0
                    If child2Id < 0 Then
                        If mScript.mainClass(classId).Properties.ContainsKey(propSelector) Then
                            selectorEventId = mScript.mainClass(classId).Properties(propSelector).eventId
                        End If
                    ElseIf child3Id < 0 Then
                        If mScript.mainClass(classId).ChildProperties(child2Id).ContainsKey(propSelector) Then
                            selectorEventId = mScript.mainClass(classId).ChildProperties(child2Id)(propSelector).eventId
                        End If
                    Else
                        If mScript.mainClass(classId).ChildProperties(child2Id).ContainsKey(propSelector) Then
                            selectorEventId = mScript.mainClass(classId).ChildProperties(child2Id)(propSelector).ThirdLevelEventId(child3Id)
                        End If
                    End If

                    If selectorEventId > 0 Then
                        'событие выбора селектора описано
                        selector = 0
                        'Dim cnt As Integer = GetSelectorsCount(mScript.eventRouter.lstEvents(eventId)) 'получаем количестово селекторов в тексте свойства
                        Dim cnt As Integer = GetSelectorsCount(mScript.eventRouter.GetExDataByEventId(eventId)) 'получаем количестово селекторов в тексте свойства
                        If cnt = 0 Then
                            'нет ни одного селектора
                            selector = 0
                        Else
                            Dim selArrs() As String 'параметры для передачи событию выбора селектора
                            If IsNothing(arrParams) OrElse arrParams.Count = 0 Then
                                ReDim selArrs(0)
                                selArrs(0) = cnt.ToString
                            Else
                                ReDim selArrs(arrParams.Count)
                                Array.ConstrainedCopy(arrParams, 0, selArrs, 1, arrParams.Count)
                                selArrs(0) = cnt.ToString
                            End If
                            Dim res As String = mScript.eventRouter.RunEvent(selectorEventId, selArrs, mScript.mainClass(classId).Names.Last & "." & propName & "Selector", False) 'запускаем событие вроде DescriptionSelector, которое возвращает номер селектора для вывода текста
                            If res = "#Error" Then
                                selector = 1 'если произошла ошибка при выборе - берем первое описание (Писатель все-равно увидит сообщение об ошибке и EXIT_CODE все-равно уже True)
                            Else
                                If Integer.TryParse(res, selector) = False Then selector = 1
                            End If
                        End If
                        'result = mScript.MakeExec(mScript.eventRouter.lstEvents(eventId), 0, -1, arrParams, False, selector)
                        result = mScript.MakeExec(mScript.eventRouter.GetExDataByEventId(eventId), 0, -1, arrParams, False, selector)
                    Else
                        'селектора нет или он незаполнен
                        result = mScript.MakeExec(mScript.eventRouter.GetExDataByEventId(eventId), 0, -1, arrParams, False, 0)
                    End If
                Else
                    'указан нужный селектор - получаем описание из нужного селектора
                    result = mScript.MakeExec(mScript.eventRouter.GetExDataByEventId(eventId), 0, -1, arrParams, False, selector)
                End If
                retFormat = MatewScript.ReturnFormatEnum.TO_LONG_TEXT
            Else
                'тип - скрипт
                value = mScript.eventRouter.RunEvent(eventId, arrParams, "Свойство " & mScript.mainClass(classId).Names.Last & "." & propName, False)
                If value = "#Error" Then Return False
                result = value
                'retFormat = MatewScript.ReturnFormatEnum.TO_CODE
                retFormat = mScript.Param_GetType(value)
            End If
        Else
            'простое значение
            If value.StartsWith("'?") Then
                'executable string
                result = mScript.ExecuteString({value}, arrParams)
                retFormat = MatewScript.ReturnFormatEnum.TO_EXECUTABLE_STRING
            Else
                'теперь точно простое значение
                result = value
                retFormat = mScript.Param_GetType(result)
            End If
        End If

        Return True
er:
        Return False
    End Function

    ''' <summary>
    ''' Функция получает итоговое значение свойства в формате целого числа. Если оно содержит скрипт или командную строку, то возвращается их результат, если длинный текст, то возвращает текст, отобранный селекторм.
    ''' Иначе возвращает простое значение свойства. 'Кавычки' со строк не убираются
    ''' </summary>
    ''' <param name="classId">Id класса свойства</param>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="child2Id">Id элемента 2 уровня</param>
    ''' <param name="child3Id">Id элемента 3 уровня</param>
    ''' <param name="result">ссылка для получения значения свойства</param>
    ''' <param name="arrParams">массив параметров на случай, если его надо будет передать скрипту</param>
    ''' <returns>False если ошибка</returns>
    Public Function ReadPropertyInt(ByVal classId As Integer, ByVal propName As String, ByVal child2Id As Integer, ByVal child3Id As Integer, ByRef result As Integer, ByRef arrParams() As String) As Boolean
        Dim strResult As String = ""
        If ReadProperty(classId, propName, child2Id, child3Id, strResult, arrParams) = False Then Return False

        If strResult.StartsWith("'"c) Then strResult = UnWrapString(strResult) 'если строка, то получаем ее без кавычек. Это полезно для получения численных Enum
        If IsNumeric(strResult) Then
            result = CInt(strResult)
        Else
            result = Val(strResult)
        End If

        Return True
    End Function

    ''' <summary>
    ''' Функция получает итоговое значение свойства в формате дробного числа. Если оно содержит скрипт или командную строку, то возвращается их результат, если длинный текст, то возвращает текст, отобранный селекторм.
    ''' Иначе возвращает простое значение свойства. 'Кавычки' со строк не убираются
    ''' </summary>
    ''' <param name="classId">Id класса свойства</param>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="child2Id">Id элемента 2 уровня</param>
    ''' <param name="child3Id">Id элемента 3 уровня</param>
    ''' <param name="result">ссылка для получения значения свойства</param>
    ''' <param name="arrParams">массив параметров на случай, если его надо будет передать скрипту</param>
    ''' <returns>False если ошибка</returns>
    Public Function ReadPropertyDbl(ByVal classId As Integer, ByVal propName As String, ByVal child2Id As Integer, ByVal child3Id As Integer, ByRef result As Double, ByRef arrParams() As String) As Boolean
        Dim strResult As String = ""
        If ReadProperty(classId, propName, child2Id, child3Id, strResult, arrParams) = False Then Return False

        If strResult.StartsWith("'"c) Then strResult = UnWrapString(strResult) 'если строка, то получаем ее без кавычек. Это полезно для получения численных Enum
        If Double.TryParse(strResult, System.Globalization.NumberStyles.Any, provider_points, result) = False Then
            result = Val(strResult)
        End If

        Return True
    End Function


    ''' <summary>
    ''' Функция получает итоговое значение свойства в формате Bool. Если оно содержит скрипт или командную строку, то возвращается их результат, если длинный текст, то возвращает текст, отобранный селекторм.
    ''' Иначе возвращает простое значение свойства. 'Кавычки' со строк не убираются
    ''' </summary>
    ''' <param name="classId">Id класса свойства</param>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="child2Id">Id элемента 2 уровня</param>
    ''' <param name="child3Id">Id элемента 3 уровня</param>
    ''' <param name="result">ссылка для получения значения свойства</param>
    ''' <param name="arrParams">массив параметров на случай, если его надо будет передать скрипту</param>
    ''' <returns>False если ошибка</returns>
    Public Function ReadPropertyBool(ByVal classId As Integer, ByVal propName As String, ByVal child2Id As Integer, ByVal child3Id As Integer, ByRef result As Boolean, ByRef arrParams() As String) As Boolean
        Dim strResult As String = ""
        If ReadProperty(classId, propName, child2Id, child3Id, strResult, arrParams) = False Then Return False

        Boolean.TryParse(strResult, result)
        Return True
    End Function
#End Region

#Region "HTML functions"
    ''' <summary>
    ''' Заменяет указанный класс html-элемента на другой. Если старого класса нет, то просто добавляется новый
    ''' </summary>
    ''' <param name="hElement">html-элемент</param>
    ''' <param name="oldClassName">старый класс</param>
    ''' <param name="newClassName">новый класс</param>
    Public Function HTMLReplaceClass(ByRef hElement As HtmlElement, ByVal oldClassName As String, ByVal newClassName As String) As Boolean
        Dim strCurrent As String = hElement.GetAttribute("ClassName")
        If String.IsNullOrEmpty(strCurrent) OrElse String.Compare(strCurrent, oldClassName, True) = 0 Then
            hElement.SetAttribute("ClassName", newClassName)
            Return Not String.IsNullOrEmpty(strCurrent)
        ElseIf String.Compare(strCurrent, newClassName, True) = 0 Then
            Return False
        End If

        Dim lst As New List(Of String)
        lst = Split(strCurrent, " ").ToList
        Dim blnRemoved As Boolean = False
        If lst.Contains(oldClassName, StringComparer.CurrentCultureIgnoreCase) Then
            lst.Remove(oldClassName)
            blnRemoved = True
        End If

        If lst.Contains(newClassName, StringComparer.CurrentCultureIgnoreCase) Then Return True

        lst.Add(newClassName)
        hElement.SetAttribute("ClassName", Join(lst.ToArray, " "))
        Return True
    End Function

    ''' <summary>
    ''' Добавляет указанный класс html-элементу
    ''' </summary>
    ''' <param name="hElement">html-элемент</param>
    ''' <param name="className">добавляемый класс</param>
    ''' <returns>True если класс был добавлен, False - если уже был</returns>
    Public Function HTMLAddClass(ByRef hElement As HtmlElement, ByVal className As String) As Boolean
        Dim strCurrent As String = hElement.GetAttribute("ClassName")
        If String.IsNullOrEmpty(strCurrent) Then
            hElement.SetAttribute("ClassName", className)
            Return True
        End If
        Dim arr() As String = Split(strCurrent, " ")
        For i As Integer = 0 To arr.Count - 1
            If String.Compare(arr(i), className, True) = 0 Then Return False
        Next

        hElement.SetAttribute("ClassName", strCurrent & " " & className)
        Return True
    End Function

    ''' <summary>
    ''' Добавляет указанный класс если его не было, и удаляет, если был
    ''' </summary>
    ''' <param name="hElement">html-элемент</param>
    ''' <param name="className">добавляемый/убираемый класс</param>
    Public Function HTMLSwitchClass(ByRef hElement As HtmlElement, ByVal className As String) As Boolean
        Dim strCurrent As String = hElement.GetAttribute("ClassName")
        If String.IsNullOrEmpty(strCurrent) Then
            hElement.SetAttribute("ClassName", className)
            Return True
        End If

        Dim lst As New List(Of String)
        lst = Split(strCurrent, " ").ToList
        If lst.Contains(className, StringComparer.CurrentCultureIgnoreCase) Then
            lst.Remove(className)
            hElement.SetAttribute("ClassName", Join(lst.ToArray, " "))
            Return False
        Else
            lst.Add(className)
            hElement.SetAttribute("ClassName", Join(lst.ToArray, " "))
            Return True
        End If
    End Function

    ''' <summary>
    ''' Проверяет есть ли указанный класс у данного элемента
    ''' </summary>
    ''' <param name="hElement">html-элемент</param>
    ''' <param name="className">проверяемый класс</param>
    Public Function HTMLHasClass(ByRef hElement As HtmlElement, ByVal className As String) As Boolean
        Dim strCurrent As String = hElement.GetAttribute("ClassName")
        If String.IsNullOrEmpty(strCurrent) Then
            hElement.SetAttribute("ClassName", className)
            Return True
        End If

        Dim lst As New List(Of String)
        lst = Split(strCurrent, " ").ToList
        If lst.Contains(className, StringComparer.CurrentCultureIgnoreCase) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Удаляет указанный класс у html-элемента
    ''' </summary>
    ''' <param name="hElement">html-элемент</param>
    ''' <param name="className">удаляемый класс</param>
    Public Function HTMLRemoveClass(ByRef hElement As HtmlElement, ByVal className As String) As Boolean
        Dim strCurrent As String = hElement.GetAttribute("ClassName")
        If String.IsNullOrEmpty(strCurrent) Then Return False
        If String.Compare(strCurrent, className, True) = 0 Then
            hElement.SetAttribute("ClassName", "")
            Return True
        End If
        Dim lst As New List(Of String)
        lst = Split(strCurrent, " ").ToList
        Dim isExists As Boolean = False
        If lst.Contains(className, StringComparer.CurrentCultureIgnoreCase) Then
            lst.Remove(className)
            isExists = True
        End If
        hElement.SetAttribute("ClassName", Join(lst.ToArray, " "))
        Return isExists
    End Function

    ''' <summary>
    ''' Добавляет CSS-файл html-документу
    ''' </summary>
    ''' <param name="hDoc">html-документ</param>
    ''' <param name="strCSS">CSS-файл</param>
    Public Sub HtmlAppendCSS(ByRef hDoc As HtmlDocument, ByVal strCSS As String)
        If IsNothing(hDoc) Then Return
        Dim hDom As mshtml.HTMLDocument = hDoc.DomDocument
        hDom.createStyleSheet(strCSS)
    End Sub

    ''' <summary>
    ''' Заменяем путь к css в первой ссылке link на указанный файл. Сам html-файл при этом не меняется
    ''' </summary>
    ''' <param name="hDoc">html документ</param>
    ''' <param name="strCSS">ссылка на новый css-файл</param>
    Public Function HtmlChangeFirstCSSLink(ByRef hDoc As HtmlDocument, ByVal strCSS As String) As Boolean
        If IsNothing(hDoc) Then Return False
        strCSS = strCSS.Replace("\", "/")
        For i As Integer = 0 To hDoc.All.Count - 1
            Dim itm As HtmlElement = hDoc.All(i)
            If itm.TagName <> "LINK" AndAlso itm.GetAttribute("type") <> "text/css" Then Continue For
            Dim prevCSS As String = itm.GetAttribute("href")
            If String.IsNullOrEmpty(prevCSS) = False AndAlso String.Compare(strCSS, prevCSS, True) = 0 Then Return True
            itm.SetAttribute("href", strCSS)
            Return True
        Next
        Return False
    End Function

    ''' <summary>
    ''' Удаляет ссылку на css-файл. Сам html-файл при этом не меняется
    ''' </summary>
    ''' <param name="hDoc">html документ</param>
    ''' <param name="strCSS">путь к файлу css, ссылку на который надо убрать</param>
    Public Function HtmlRemoveCSS(ByRef hDoc As HtmlDocument, ByVal strCSS As String) As Boolean
        If IsNothing(hDoc) Then Return False
        strCSS = strCSS.Replace("\", "/")
        For i As Integer = 0 To hDoc.All.Count - 1
            Dim itm As HtmlElement = hDoc.All(i)
            If itm.TagName <> "LINK" AndAlso itm.GetAttribute("type") <> "text/css" Then Continue For
            Dim path As String = itm.GetAttribute("href")
            If String.IsNullOrEmpty(path) OrElse String.Compare(path, strCSS, True) = 0 Then
                Dim msNode As mshtml.IHTMLDOMNode = itm.DomElement
                msNode.removeNode(True)
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Добавляет новый css-стиль к html-элементу
    ''' </summary>
    ''' <param name="hEl">html-элемент</param>
    ''' <param name="styleName">название стиля, например "color"</param>
    ''' <param name="styleValue">значение, нарпимер "#000000"</param>
    ''' <returns>True если стиль добавлен, False если уже был</returns>
    Public Function HTMLAddCSSstyle(ByRef hEl As HtmlElement, ByVal styleName As String, ByVal styleValue As String) As Boolean
        Dim strStyle As String = hEl.Style
        If String.IsNullOrEmpty(strStyle) Then
            hEl.Style = styleName & ":" & styleValue
            Return True
        End If

        Dim arrStyles() As String = Split(strStyle, ";")
        Dim sCount As Integer = arrStyles.Count
        If String.IsNullOrEmpty(arrStyles(sCount - 1).Trim) Then sCount -= 1

        Dim blnFound As Boolean = False
        For i As Integer = 0 To sCount - 1
            Dim curStyle As String = arrStyles(i).Trim
            Dim pos As Integer = curStyle.IndexOf(":")
            If pos <= 0 Then Continue For
            Dim curName As String = curStyle.Substring(0, pos).Trim
            Dim curValue As String = curStyle.Substring(pos + 1).Trim
            If String.Compare(styleName, curName, True) = 0 Then
                'стиль найден
                If curValue = styleValue Then Return False 'значение не изменилось
                arrStyles(i) = curName & ":" & styleValue
                blnFound = True
                Exit For
            End If
        Next i

        If Not blnFound Then
            'такого стиля не было
            If strStyle.EndsWith(";") = False Then strStyle &= ";"
            strStyle &= styleName & ":" & styleValue
        Else
            'такой стиль уже был, мы изменили значение. Собираем строку заново
            strStyle = Join(arrStyles, ";")
        End If

        hEl.Style = strStyle
        Return True
    End Function

    ''' <summary>
    ''' Удаляет css-стиль из html-элемента
    ''' </summary>
    ''' <param name="hEl">html-элемент</param>
    ''' <param name="styleName">название стиля, например "color"</param>
    ''' <returns>True если стиль удален, False если его и не было</returns>
    Public Function HTMLRemoveCSSstyle(ByRef hEl As HtmlElement, ByVal styleName As String) As Boolean
        Dim strStyle As String = hEl.Style
        If String.IsNullOrEmpty(strStyle) Then Return False

        Dim arrStyles() As String = Split(strStyle, ";")
        Dim sCount As Integer = arrStyles.Count
        If String.IsNullOrEmpty(arrStyles(sCount - 1).Trim) Then sCount -= 1

        Dim posId As Integer = -1
        For i As Integer = 0 To sCount - 1
            Dim curStyle As String = arrStyles(i).Trim
            Dim pos As Integer = curStyle.IndexOf(":")
            If pos <= 0 Then Continue For
            Dim curName As String = curStyle.Substring(0, pos).Trim
            'Dim curValue As String = curStyle.Substring(pos + 1).Trim
            If String.Compare(styleName, curName, True) = 0 Then
                'стиль найден
                posId = i
                Exit For
            End If
        Next i

        If posId = -1 Then
            'такого стиля не было
            Return False
        Else
            'такой стиль был. Удаляем
            Dim lstStyles As List(Of String) = arrStyles.ToList
            lstStyles.RemoveAt(posId)
            strStyle = Join(lstStyles.ToArray, ";")
        End If

        hEl.Style = strStyle
        Return True
    End Function

    ''' <summary>
    ''' Получает значение указанного css-стиля из html-элемента
    ''' </summary>
    ''' <param name="hEl">html-элемент</param>
    ''' <param name="styleName">название стиля, например "color"</param>
    ''' <returns>значение или пустую строку, если не найден</returns>
    Public Function HTMLGetCSSstyleValue(ByRef hEl As HtmlElement, ByVal styleName As String) As String
        Dim strStyle As String = hEl.Style
        If String.IsNullOrEmpty(strStyle) Then Return ""

        Dim arrStyles() As String = Split(strStyle, ";")
        Dim sCount As Integer = arrStyles.Count
        If String.IsNullOrEmpty(arrStyles(sCount - 1).Trim) Then sCount -= 1

        Dim posId As Integer = -1
        For i As Integer = 0 To sCount - 1
            Dim curStyle As String = arrStyles(i).Trim
            Dim pos As Integer = curStyle.IndexOf(":")
            If pos <= 0 Then Continue For
            Dim curName As String = curStyle.Substring(0, pos).Trim
            Dim curValue As String = curStyle.Substring(pos + 1).Trim
            If String.Compare(styleName, curName, True) = 0 Then
                'стиль найден
                Return curValue
                Exit For
            End If
        Next i

        Return ""
    End Function

#End Region

    ''' <summary>
    ''' Убирает 'кавычки' со строки. Указанная строка гарантировано должна не быть исполняемой и не должна содержать скрипт/длинный текст. В случае, если в переданой строке число или bool, то оно же и возвращается
    ''' </summary>
    ''' <param name="strText">Страка</param>
    Public Function UnWrapString(ByVal strText As String) As String
        If String.IsNullOrEmpty(strText) Then Return strText
        If strText.Chars(0) = "'"c AndAlso strText.Last = "'"c Then
            strText = strText.Substring(1, strText.Length - 2).Replace("/'", "'")
        End If
        Return strText
    End Function

    ''' <summary>
    ''' Функция сверяет массив переданных данных с указанной строкой и возвращает похожи они или нет
    ''' </summary>
    ''' <param name="strToCheck">строка для проверки</param>
    ''' <param name="start">начальный индекс в массиве arrVariants</param>
    ''' <param name="final">конечный индекс в массиве arrVariants</param>
    ''' <param name="arrVariants">массив с варианами для сравнения</param>
    ''' <param name="arrParams">массив arrParams</param>
    ''' <param name="maxErrors">маскимальное допустимое количество ошибок</param>
    ''' <returns>индекс элемента в arrVariants с совпавшим значением</returns>
    Public Function IndexOfAnyWithErrors(ByVal strToCheck As String, ByVal start As Integer, ByVal final As Integer, ByRef arrVariants() As String, ByRef arrParams() As String, _
                                         Optional ByVal maxErrors As Integer = -1) As Integer
        If IsNothing(arrVariants) OrElse arrVariants.Count = 0 Then Return -1
        If final = -1 OrElse final > arrVariants.Count - 1 Then final = arrVariants.Count - 1
        strToCheck = strToCheck.Trim.ToLower
        If String.IsNullOrEmpty(strToCheck) Then Return -1

        strToCheck = UnWrapString(strToCheck).Trim
        If maxErrors = -1 Then
            'Dim strUnWrapped As String = UnWrapString(strToCheck)
            Select Case strToCheck.Length
                Case Is <= 5
                    maxErrors = 1
                Case 6 To 10
                    maxErrors = 2
                Case Else
                    maxErrors = 2
            End Select
        End If

        If maxErrors = 0 Then
            'поиск точного совпадения без учетов регистра и пробелов
            For i As Integer = start To final
                If String.Compare(strToCheck, arrVariants(i).Trim, True) = 0 Then Return i
            Next i
            Return -1
        End If

        Dim arrCheck() As Char = strToCheck.ToCharArray
        Dim arrCheckUBound As Integer = arrCheck.Count - 1
        Dim cStep As Integer = 1
        If start > final Then cStep = -1
        For wrd As Integer = start To final Step cStep
            'перебираем все слова для поиска
            Dim curWord As String = mScript.PrepareStringToPrint(arrVariants(wrd), arrParams).Trim.ToLower
            If curWord.Length = 0 Then Continue For
            Dim errCount As Integer = 0

            Dim arrVar() As Char = curWord.ToCharArray
            Dim arrVarUBound As Integer = arrVar.Count - 1
            If Math.Abs(arrCheckUBound - arrVarUBound) > maxErrors Then Continue For

            Dim posInVariant As Integer = 0
            Dim chId As Integer = 0
            For chId = 0 To arrCheck.Count - 1
                If posInVariant > arrVarUBound Then
                    errCount += 1
                    Continue For
                End If
                If errCount > maxErrors Then Exit For

                If arrCheck(chId) = arrVar(posInVariant) Then
                    posInVariant += 1
                    Continue For
                End If

                'arrCheck пИВо
                'arrVar   пВИо
                If chId + 1 <= arrCheckUBound AndAlso posInVariant + 1 <= arrVarUBound Then
                    If arrCheck(chId) = arrVar(posInVariant + 1) AndAlso arrCheck(chId + 1) = arrVar(posInVariant) Then
                        chId += 1
                        posInVariant += 2
                        errCount += 1
                        If chId + 1 > arrCheckUBound AndAlso posInVariant <= arrVarUBound Then
                            'после то, как перепутали две буквы, поставили еще одну лишнюю
                            errCount += 1
                        End If
                        Continue For
                    End If
                End If

                'arrCheck пИво
                'arrVar   пво
                If chId + 1 <= arrCheckUBound Then
                    If arrCheck(chId + 1) = arrVar(posInVariant) Then
                        errCount += 1
                        Continue For
                    End If
                End If

                'arrCheck пво
                'arrVar   пИво
                If posInVariant + 1 <= arrVarUBound Then
                    If arrCheck(chId) = arrVar(posInVariant + 1) Then
                        posInVariant += 2
                        errCount += 1
                        Continue For
                    End If
                End If

                posInVariant += 1
                errCount += 1

            Next chId

            If chId = arrCheckUBound + 1 AndAlso posInVariant <= arrVarUBound Then
                errCount += arrVarUBound - posInVariant + 1
            End If

            If errCount <= maxErrors Then Return wrd
        Next wrd

        Return -1
    End Function

#Region "InGame Date and Time"
    ''' <summary>получаем размерность минуты, часа, суток </summary>
    Private Sub PlayerDate_GetTimeIntervals(ByRef hoursPerDay As Integer, ByRef minutesPerHour As Integer, ByRef secondsPerMinute As Integer, ByRef arrParams() As String)
        Dim classD As Integer = mScript.mainClassHash("Date")
        hoursPerDay = 24
        minutesPerHour = 60
        secondsPerMinute = 60
        ReadPropertyInt(classD, "PlayerDate_HoursPerDay", -1, -1, hoursPerDay, arrParams)
        ReadPropertyInt(classD, "PlayerDate_MinutesPerHour", -1, -1, minutesPerHour, arrParams)
        ReadPropertyInt(classD, "PlayerDate_SecondsPerMinutes", -1, -1, secondsPerMinute, arrParams)
        If hoursPerDay <= 0 Then hoursPerDay = 24
        If minutesPerHour <= 0 Then minutesPerHour = 60
        If secondsPerMinute <= 0 Then secondsPerMinute = 60

    End Sub

    ''' <summary>получаем игровые месяцы </summary>
    Private Function PlayerDate_Get_Monthes(ByRef arrParams() As String) As Dictionary(Of String, Integer)
        Dim classD As Integer = mScript.mainClassHash("Date")
        Dim lstMonthes As New Dictionary(Of String, Integer)
        Dim varMonthes As String = ""
        ReadProperty(classD, "PlayerDate_Monthes", -1, -1, varMonthes, arrParams)
        varMonthes = UnWrapString(varMonthes)
        Dim var As cVariable.variableEditorInfoType = Nothing

        If String.IsNullOrEmpty(varMonthes) = False AndAlso mScript.csPublicVariables.lstVariables.TryGetValue(varMonthes, var) Then
            If IsNothing(var.lstSingatures) = False AndAlso var.lstSingatures.Count = var.arrValues.Count Then
                For i As Integer = 0 To var.arrValues.Count - 1
                    Dim daysCount As Integer = var.arrValues(i)
                    Dim signId As Integer = var.lstSingatures.IndexOfValue(i)
                    If signId = -1 Then
                        MessageBox.Show("Установка игровой даты : неправильный формат массива месяцев.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        daysCount = 30
                    End If
                    If daysCount <= 0 Then
                        MessageBox.Show("Установка игровой даты, " & var.lstSingatures.ElementAt(i).Key & ": количество дней вне диапазона возможных значений.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        daysCount = 30
                    End If
                    lstMonthes.Add(var.lstSingatures.ElementAt(signId).Key, daysCount)
                Next i
            End If
        End If

        If lstMonthes.Count = 0 Then
            For i As Integer = 1 To 12
                Dim aDate As Date = Date.ParseExact("01." & i.ToString & ".2000", "dd.M.yyyy", System.Globalization.CultureInfo.InvariantCulture)
                Dim monthName As String = Format(aDate, "dd.MMMM")

                monthName = monthName.Substring(3)

                lstMonthes.Add(monthName, Date.DaysInMonth(2000, i))
            Next
        End If

        Return lstMonthes
    End Function


    ''' <summary>получаем игровые дни недели </summary>
    Private Function PlayerDate_Get_Weekdays(ByRef arrParams() As String) As List(Of String)
        'получаем игровые дни недели и начальный день недели
        Dim classD As Integer = mScript.mainClassHash("Date")
        Dim weekDay As Integer = 0
        ReadPropertyInt(classD, "PlayerDate_InitialWeekday", -1, -1, weekDay, arrParams)

        Dim lstWeekDays As New List(Of String)
        Dim varWeek As String = ""
        ReadProperty(classD, "PlayerDate_Weekdays", -1, -1, varWeek, arrParams)
        varWeek = UnWrapString(varWeek)
        Dim var As cVariable.variableEditorInfoType = Nothing

        If String.IsNullOrEmpty(varWeek) = False AndAlso mScript.csPublicVariables.lstVariables.TryGetValue(varWeek, var) Then
            If (var.arrValues.Count = 0 AndAlso var.arrValues(0) = "") = False Then
                lstWeekDays.AddRange(var.arrValues)
                For i As Integer = 0 To lstWeekDays.Count - 1
                    lstWeekDays(i) = UnWrapString(lstWeekDays(i))
                Next i
            End If
        End If

        If lstWeekDays.Count = 0 Then
            For i As Integer = 1 To 7
                lstWeekDays.Add(DateAndTime.WeekdayName(i, False, FirstDayOfWeek.Sunday))
            Next
        End If

        Return lstWeekDays
    End Function


    Public Sub PlayerDate_SetInitialTime(ByRef arrParams() As String)
        Dim classD As Integer = mScript.mainClassHash("Date")

        'получаем размерность минуты, часа, суток
        Dim hoursPerDay As Integer = 0, minutesPerHour As Integer = 0, secondsPerMinute As Integer = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        'получаем исходный год, месяц, день
        Dim initYear As Integer = 0, initMonth As Integer = 0, initDay As Integer = 0
        ReadPropertyInt(classD, "PlayerDate_InitialYear", -1, -1, initYear, arrParams)
        ReadPropertyInt(classD, "PlayerDate_InitialMonth", -1, -1, initMonth, arrParams)
        ReadPropertyInt(classD, "PlayerDate_InitialDay", -1, -1, initDay, arrParams)

        If initMonth < 1 OrElse initMonth > lstMonthes.Count Then
            MessageBox.Show("Установка начальной игровой даты. Номер месяца вне диапазона возможных значений.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            initMonth = 0
        Else
            initMonth -= 1
        End If
        If initDay > lstMonthes.ElementAt(initMonth).Value Then
            MessageBox.Show("Установка начальной игровой даты. Количество дней (PlayerDate_InitialDay) в начальном месяце (PlayerDate_InitialMonth) вне диапазона возможных значений.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            initDay = 0
        Else
            initDay -= 1
        End If

        'трансформируем исходную дату в Long (количество секунд)
        GVARS.PLAYER_TIME = PlayerDate_Calculate_Date(secondsPerMinute, minutesPerHour, hoursPerDay, initDay, initMonth, initYear, 0, 0, 0, lstMonthes)
    End Sub

    ''' <summary>
    ''' Возвращает количество игровых секунд в дате
    ''' </summary>
    ''' <param name="secondsPerMinute"></param>
    ''' <param name="minutesPerHour"></param>
    ''' <param name="hoursPerDay"></param>
    ''' <param name="days"></param>
    ''' <param name="monthes"></param>
    ''' <param name="years"></param>
    ''' <param name="hours"></param>
    ''' <param name="minutes"></param>
    ''' <param name="seconds"></param>
    ''' <param name="lstMonthes"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PlayerDate_Calculate_Date(ByVal secondsPerMinute As Integer, ByVal minutesPerHour As Integer, ByVal hoursPerDay As Integer, _
                                          ByVal days As Integer, ByVal monthes As Integer, ByVal years As Integer, _
                                          ByVal hours As Integer, ByVal minutes As Integer, ByVal seconds As Integer, ByRef lstMonthes As Dictionary(Of String, Integer)) As Long
        'трансформируем исходную дату в Long (количество секунд)
        Dim secondsPerDay As Long = secondsPerMinute * minutesPerHour * hoursPerDay
        Dim result As Long = days * secondsPerDay  'посчитали дни
        'добавляем месяцы
        For i As Integer = 0 To monthes - 1
            result += lstMonthes.ElementAt(i).Value * secondsPerDay
        Next
        'получаем количество секунд в году
        Dim secondsPerYear As Long = lstMonthes.Values.Sum * secondsPerDay
        'добавляем годы
        result += years * secondsPerYear

        'добавляем время
        result += hours * secondsPerMinute * minutesPerHour + minutes * secondsPerMinute + seconds

        Return result
    End Function

    ''' <summary>
    ''' Устанавливает указанную игровую дату
    ''' </summary>
    ''' <param name="strDate">дата в формате dd.MM.yyyy HH:mm:ss</param>
    Public Function PlayerDate_SetDate(ByVal strDate As String, ByRef arrParams() As String) As String
        If IsNumeric(strDate) Then
            GVARS.PLAYER_TIME = Val(strDate)
            Return ""
        End If

        strDate = UnWrapString(strDate).Trim
        If String.IsNullOrEmpty(strDate) Then Return _ERROR("Формат даты неправильный.", "PlayerDate_Set")
        Dim years As Integer = 0, monthes As Integer = 0, days As Integer = 0, hours As Integer = 0, minutes As Integer = 0, seconds As Integer = 0
        Dim arr1() As String = strDate.Split(" "c)
        Dim dateOnly As Boolean = False, timeOnly As Boolean = False

        Dim arr2() As String = Split(arr1(0), "."c)

        If arr2.Count <> 3 Then
            arr2 = Split(arr1(0), ":"c)
            If arr2.Count <> 3 Then Return _ERROR("Формат даты неправильный.", "PlayerDate_Set")
            timeOnly = True
            hours = Val(arr2(0))
            minutes = Val(arr2(1))
            seconds = Val(arr2(2))
        Else
            If arr1.Count = 1 Then
                dateOnly = True
                days = Val(arr2(0)) - 1
                monthes = Val(arr2(1)) - 1
                years = Val(arr2(2))
            ElseIf arr1.Count <> 2 Then
                Return _ERROR("Формат даты неправильный.", "PlayerDate_Set")
            Else
                days = Val(arr2(0)) - 1
                monthes = Val(arr2(1)) - 1
                years = Val(arr2(2))

                arr2 = Split(arr1(1), ":"c)
                hours = Val(arr2(0))
                minutes = Val(arr2(1))
                seconds = Val(arr2(2))
            End If
        End If

        'получаем размерность минуты, часа, суток
        Dim hoursPerDay As Integer = 0, minutesPerHour As Integer = 0, secondsPerMinute As Integer = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        GVARS.PLAYER_TIME = PlayerDate_Calculate_Date(secondsPerMinute, minutesPerHour, hoursPerDay, days, monthes, years, hours, minutes, seconds, lstMonthes)
        Return ""
    End Function

    ''' <summary>
    ''' Возвращает количество игровых секунд в указанной дате
    ''' </summary>
    ''' <param name="strDate">дата в формате dd.MM.yyyy HH:mm:ss / dd.MM.yyyy / HH:mm:ss</param>
    ''' <param name="arrParams"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PlayerDate_DateToSeconds(ByVal strDate As String, ByRef arrParams() As String) As Long
        If IsNumeric(strDate) Then
            Return Val(strDate)
        End If

        strDate = UnWrapString(strDate).Trim
        If String.IsNullOrEmpty(strDate) Then Return _ERROR("Формат даты неправильный.", "PlayerDate_DateToSeconds")
        Dim years As Integer = 0, monthes As Integer = 0, days As Integer = 0, hours As Integer = 0, minutes As Integer = 0, seconds As Integer = 0
        Dim arr1() As String = strDate.Split(" "c)
        Dim dateOnly As Boolean = False, timeOnly As Boolean = False

        Dim arr2() As String = Split(arr1(0), "."c)

        If arr2.Count <> 3 Then
            arr2 = Split(arr1(0), ":"c)
            If arr2.Count <> 3 Then Return _ERROR("Формат даты неправильный.", "PlayerDate_DateToSeconds")
            timeOnly = True
            hours = Val(arr2(0))
            minutes = Val(arr2(1))
            seconds = Val(arr2(2))
        Else
            If arr1.Count = 1 Then
                dateOnly = True
                days = Val(arr2(0)) - 1
                monthes = Val(arr2(1)) - 1
                years = Val(arr2(2))
            ElseIf arr1.Count <> 2 Then
                Return _ERROR("Формат даты неправильный.", "PlayerDate_DateToSeconds")
            Else
                days = Val(arr2(0)) - 1
                monthes = Val(arr2(1)) - 1
                years = Val(arr2(2))

                arr2 = Split(arr1(1), ":"c)
                hours = Val(arr2(0))
                minutes = Val(arr2(1))
                seconds = Val(arr2(2))
            End If
        End If

        'получаем размерность минуты, часа, суток
        Dim hoursPerDay As Integer = 0, minutesPerHour As Integer = 0, secondsPerMinute As Integer = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        Return PlayerDate_Calculate_Date(secondsPerMinute, minutesPerHour, hoursPerDay, days, monthes, years, hours, minutes, seconds, lstMonthes)
    End Function

    ''' <summary>
    ''' Получает игровую дату в формате dd.MM.yyyy HH:mm:ss
    ''' </summary>
    ''' <param name="arrParams"></param>
    ''' <param name="dFormat">Формат, в котором нужно вернуть дату</param>
    ''' <remarks></remarks>
    Public Function PlayerDate_Get(ByVal dFormat As String, ByRef arrParams() As String, Optional aDate As String = "") As String
        'получаем размерность минуты, часа, суток
        Dim curDate As Long = GVARS.PLAYER_TIME
        If String.IsNullOrEmpty(aDate) = False Then
            curDate = PlayerDate_DateToSeconds(aDate, arrParams)
        End If
        Dim hoursPerDay As Long = 0, minutesPerHour As Integer = 0, secondsPerMinute As Integer = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        'получаем количество секунд в дне
        Dim secondsPerDay As Long = secondsPerMinute * minutesPerHour * hoursPerDay
        'получаем количество секунд в году
        Dim secondsPerYear As Long = lstMonthes.Values.Sum * secondsPerDay

        'получили годы
        Dim years As Integer = Math.Floor(curDate / secondsPerYear)

        Dim timeLeft As Long = curDate - years * secondsPerYear

        'получили месяцы
        Dim monthes As Integer = 0
        For i As Integer = 0 To lstMonthes.Count - 1
            Dim secondsInMonth As Integer = lstMonthes.ElementAt(i).Value * secondsPerDay
            If timeLeft > secondsInMonth Then
                timeLeft -= secondsInMonth
            Else
                monthes = i ' + 1
                Exit For
            End If
        Next i

        'получили дни
        Dim days As Integer = Math.Floor(timeLeft / secondsPerDay) '+ 1
        timeLeft -= days * secondsPerDay

        'получили часы
        Dim secondsPerHour = secondsPerMinute * minutesPerHour
        Dim hours As Integer = Math.Floor(timeLeft / secondsPerHour)
        timeLeft -= hours * secondsPerHour

        'получили минуты и секунды
        Dim minutes As Integer = Math.Floor(timeLeft / secondsPerMinute)
        Dim seconds As Integer = timeLeft - minutes * secondsPerMinute

        'dd.MM.yyyy HH:mm:ss

        Dim strDate As String = PlayerDate_FormatDate(dFormat, years, monthes + 1, days + 1, hours, minutes, seconds, lstMonthes, PlayerDate_Get_Weekdays(arrParams), hoursPerDay, arrParams)
        Return strDate
    End Function

    ''' <summary>
    ''' Получает игровую дату в виде массива: год, номер месяца, число, часы, минуты, секунды, название месяца, день недели
    ''' </summary>
    ''' <param name="arrParams"></param>
    Public Function PlayerDate_DateToArray(ByRef arrParams() As String, Optional aDate As String = "") As cVariable.variableEditorInfoType
        'получаем размерность минуты, часа, суток
        Dim curDate As Long = GVARS.PLAYER_TIME
        If String.IsNullOrEmpty(aDate) = False Then
            curDate = PlayerDate_DateToSeconds(aDate, arrParams)
        End If
        Dim hoursPerDay As Long = 0, minutesPerHour As Integer = 0, secondsPerMinute As Integer = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        'получаем количество секунд в дне
        Dim secondsPerDay As Long = secondsPerMinute * minutesPerHour * hoursPerDay
        'получаем количество секунд в году
        Dim secondsPerYear As Long = lstMonthes.Values.Sum * secondsPerDay

        'получили годы
        Dim years As Integer = Math.Floor(curDate / secondsPerYear)

        Dim timeLeft As Long = curDate - years * secondsPerYear

        'получили месяцы
        Dim monthes As Integer = 0
        For i As Integer = 0 To lstMonthes.Count - 1
            Dim secondsInMonth As Integer = lstMonthes.ElementAt(i).Value * secondsPerDay
            If timeLeft > secondsInMonth Then
                timeLeft -= secondsInMonth
            Else
                monthes = i '+ 1
                Exit For
            End If
        Next i

        'получили дни
        Dim days As Integer = Math.Floor(timeLeft / secondsPerDay)
        timeLeft -= days * secondsPerDay

        'получили часы
        Dim secondsPerHour = secondsPerMinute * minutesPerHour
        Dim hours As Integer = Math.Floor(timeLeft / secondsPerHour)
        timeLeft -= hours * secondsPerHour

        'получили минуты и секунды
        Dim minutes As Integer = Math.Floor(timeLeft / secondsPerMinute)
        Dim seconds As Integer = timeLeft - minutes * secondsPerMinute

        'название месяца
        Dim monthName As String = lstMonthes.ElementAt(monthes).Key

        'день недели
        Dim weekNum As Integer = PlayerDate_GetPeriod(DateInterval.Weekday, arrParams)
        Dim wDay As String = PlayerDate_Get_Weekdays(arrParams)(weekNum - 1)

        'возвращаем массив с датой
        'год, номер месяца, число, часы, минуты, секунды, название месяца, день недели
        Dim retArray() As String = {years.ToString, (monthes + 1).ToString, (days + 1).ToString, hours.ToString, minutes.ToString, seconds.ToString, WrapString(monthName), WrapString(wDay)}
        Dim var As New cVariable.variableEditorInfoType
        var.arrValues = retArray

        var.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
        Dim arrSignatures() = {"year", "month", "day", "hour", "minute", "second", "monthName", "weekday"}
        For i As Integer = 0 To arrSignatures.Count - 1
            var.lstSingatures.Add(arrSignatures(i), i)
        Next

        Return var
    End Function

    ''' <summary>
    ''' Получает игровую дату в виде массива: год, номер месяца, число, часы, минуты, секунды, название месяца, день недели
    ''' </summary>
    ''' <param name="arrParams"></param>
    Public Function PlayerDate_IntervalToArray(ByRef arrParams() As String, ByVal strTime As String, Optional ByVal initMonth As Integer = -1) As cVariable.variableEditorInfoType
        'получаем размерность минуты, часа, суток
        Dim Interval As Long = 0
        If String.IsNullOrEmpty(strTime) = False Then
            Interval = PlayerDate_DateToSeconds(strTime, arrParams)
        End If
        Dim hoursPerDay As Long = 0, minutesPerHour As Integer = 0, secondsPerMinute As Integer = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        'получаем количество секунд в дне
        Dim secondsPerDay As Long = secondsPerMinute * minutesPerHour * hoursPerDay
        'получаем количество секунд в году
        Dim secondsPerYear As Long = lstMonthes.Values.Sum * secondsPerDay

        'получили годы
        Dim years As Integer = Math.Floor(Interval / secondsPerYear)

        Dim timeLeft As Long = Interval - years * secondsPerYear

        'получили месяцы
        If initMonth = -1 Then
            initMonth = PlayerDate_GetPeriod(DateInterval.Month, arrParams) - 1
        End If
        If initMonth > lstMonthes.Count - 1 OrElse initMonth < 0 Then initMonth = 0
        Dim monthes As Integer = 0
        Dim i As Integer = initMonth
        Do
            Dim secondsInMonth As Integer = lstMonthes.ElementAt(i).Value * secondsPerDay
            If timeLeft > secondsInMonth Then
                timeLeft -= secondsInMonth
            Else
                monthes = i - initMonth
                Exit Do
            End If

            i += 1
            If i > lstMonthes.Count - 1 Then i = 0
        Loop

        'получили дни
        Dim days As Integer = Math.Floor(timeLeft / secondsPerDay)
        timeLeft -= days * secondsPerDay

        'получили часы
        Dim secondsPerHour = secondsPerMinute * minutesPerHour
        Dim hours As Integer = Math.Floor(timeLeft / secondsPerHour)
        timeLeft -= hours * secondsPerHour

        'получили минуты и секунды
        Dim minutes As Integer = Math.Floor(timeLeft / secondsPerMinute)
        Dim seconds As Integer = timeLeft - minutes * secondsPerMinute

        'возвращаем массив интервалом времени
        'годы, месяцы, дни, часы, минуты, секунды
        Dim retArray() As String = {years.ToString, monthes.ToString, days.ToString, hours.ToString, minutes.ToString, seconds.ToString}
        Dim var As New cVariable.variableEditorInfoType
        var.arrValues = retArray

        var.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
        Dim arrSignatures() = {"year", "month", "day", "hour", "minute", "second"}
        For i = 0 To arrSignatures.Count - 1
            var.lstSingatures.Add(arrSignatures(i), i)
        Next

        Return var
    End Function


    ''' <summary>
    ''' Возвращает указанную составляющую игровой даты (год, месяц, день, час, минуты, секунды, номер дня недели от 1)
    ''' </summary>
    ''' <param name="dInterval">какую часть текущей игровой даты надо вернуть</param>
    ''' <param name="arrParams"></param>
    Public Function PlayerDate_GetPeriod(ByVal dInterval As DateInterval, ByRef arrParams() As String, Optional aDate As String = "") As Integer
        Dim curDate As Long = GVARS.PLAYER_TIME
        If String.IsNullOrEmpty(aDate) = False Then
            curDate = PlayerDate_DateToSeconds(aDate, arrParams)
        End If
        'получаем размерность минуты, часа, суток
        Dim hoursPerDay As Long = 0, minutesPerHour As Long = 0, secondsPerMinute As Long = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        'получаем количество секунд в дне
        Dim secondsPerDay As Long = secondsPerMinute * minutesPerHour * hoursPerDay
        'получаем количество секунд в году
        Dim secondsPerYear As Long = lstMonthes.Values.Sum * secondsPerDay

        'получили годы
        Dim years As Integer = Math.Floor(curDate / secondsPerYear)
        If dInterval = DateInterval.Year Then Return years

        Dim timeLeft As Long = curDate - years * secondsPerYear

        'получили месяцы
        Dim monthes As Integer = 0
        For i As Integer = 0 To lstMonthes.Count - 1
            Dim secondsInMonth As Integer = lstMonthes.ElementAt(i).Value * secondsPerDay
            If timeLeft > secondsInMonth Then
                timeLeft -= secondsInMonth
            Else
                monthes = i ' + 1
                Exit For
            End If
        Next i
        If dInterval = DateInterval.Month Then Return monthes + 1

        'получили дни
        Dim days As Integer = Math.Floor(timeLeft / secondsPerDay)
        timeLeft -= days * secondsPerDay
        If dInterval = DateInterval.Day Then Return days + 1

        'получили часы
        Dim secondsPerHour = secondsPerMinute * minutesPerHour
        Dim hours As Integer = Math.Floor(timeLeft / secondsPerHour)
        If dInterval = DateInterval.Hour Then Return hours
        timeLeft -= hours * secondsPerHour

        'получили минуты и секунды
        Dim minutes As Integer = Math.Floor(timeLeft / secondsPerMinute)
        If dInterval = DateInterval.Minute Then Return minutes
        Dim seconds As Integer = timeLeft - minutes * secondsPerMinute

        If dInterval = DateInterval.Second Then Return seconds

        'считаем день недели

        'получаем игровые дни недели и начальный день недели
        Dim classD As Integer = mScript.mainClassHash("Date")
        Dim initWeekDay As Integer = 0
        ReadPropertyInt(classD, "PlayerDate_InitialWeekday", -1, -1, initWeekDay, arrParams)

        Dim lstWeekDays As List(Of String) = PlayerDate_Get_Weekdays(arrParams)

        If initWeekDay <= 0 OrElse initWeekDay > lstWeekDays.Count Then
            MessageBox.Show("PlayerDate_GetWeekday. День недели (PlayerDate_Weekdays) вне диапазона возможных значений.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            initWeekDay = 1
        End If

        'считаем день текущий день недели
        Dim timeCopy As Long = GVARS.PLAYER_TIME
        PlayerDate_SetInitialTime(Nothing)
        Dim initialSeconds As Long = GVARS.PLAYER_TIME
        GVARS.PLAYER_TIME = timeCopy

        Dim secondsPerWeek As Long = lstWeekDays.Count * secondsPerDay
        timeLeft = curDate - initialSeconds - secondsPerWeek * Math.Floor((curDate - initialSeconds) / secondsPerWeek)
        Dim dWeek As Integer = Math.Floor(timeLeft / secondsPerDay) 'сколько дней недели надо добавить к initWeekday
        dWeek += initWeekDay
        If initWeekDay > lstWeekDays.Count Then dWeek -= lstWeekDays.Count

        Return dWeek
    End Function

    Public Function PlayerDate_Add(ByVal dInterval As DateInterval, ByVal addCount As Integer, ByRef arrParams() As String) As String
        If dInterval = DateInterval.Second Then
            GVARS.PLAYER_TIME += addCount
            Return PlayerDate_Get("dd.MM.yyyy HH:mm:ss", arrParams)
        End If

        'получаем размерность минуты, часа, суток
        Dim hoursPerDay As Integer = 0, minutesPerHour As Integer = 0, secondsPerMinute As Integer = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        Select Case dInterval
            Case DateInterval.Minute
                GVARS.PLAYER_TIME += addCount * secondsPerMinute
                Return PlayerDate_Get("dd.MM.yyyy HH:mm:ss", arrParams)
            Case DateInterval.Hour
                GVARS.PLAYER_TIME += addCount * secondsPerMinute * minutesPerHour
                Return PlayerDate_Get("dd.MM.yyyy HH:mm:ss", arrParams)
            Case DateInterval.Day, DateInterval.DayOfYear
                GVARS.PLAYER_TIME += addCount * secondsPerMinute * minutesPerHour * hoursPerDay
                Return PlayerDate_Get("dd.MM.yyyy HH:mm:ss", arrParams)
            Case DateInterval.Weekday, DateInterval.WeekOfYear
                Dim lstWeekDays As List(Of String) = PlayerDate_Get_Weekdays(arrParams)

                addCount *= lstWeekDays.Count
                GVARS.PLAYER_TIME += addCount * secondsPerMinute * minutesPerHour * hoursPerDay
                Return PlayerDate_Get("dd.MM.yyyy HH:mm:ss", arrParams)
        End Select

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        Select Case dInterval
            Case DateInterval.Month, DateInterval.Quarter
                If dInterval = DateInterval.Quarter Then addCount *= 3
                Dim curMonth As Integer = PlayerDate_GetPeriod(DateInterval.Month, arrParams)
                Dim secondsPerDay As Long = secondsPerMinute * minutesPerHour * hoursPerDay
                curMonth -= 1

                Do
                    Dim nextMonth As Integer = curMonth + 1
                    If nextMonth > lstMonthes.Count - 1 Then nextMonth = 0

                    GVARS.PLAYER_TIME += lstMonthes.ElementAt(nextMonth).Value * secondsPerDay
                    curMonth += 1
                    If curMonth > lstMonthes.Count - 1 Then curMonth = 0
                    addCount -= 1
                    If addCount <= 0 Then Exit Do
                Loop

            Case DateInterval.Year
                GVARS.PLAYER_TIME += addCount * secondsPerMinute * minutesPerHour * hoursPerDay * lstMonthes.Values.Sum
                Return PlayerDate_Get("dd.MM.yyyy HH:mm:ss", arrParams)
        End Select

        Return ""
    End Function

    ''' <summary>
    ''' Возвращает разницу между игровыми датами strDate2 - strDate1 в указанных игровых интервалах
    ''' </summary>
    ''' <param name="dInterval"></param>
    ''' <param name="strDate1"></param>
    ''' <param name="strDate2"></param>
    ''' <param name="arrParams"></param>
    ''' <returns></returns>
    Public Function PlayerDate_DateDiff(ByVal dInterval As DateInterval, ByVal strDate1 As String, ByVal strDate2 As String, ByRef arrParams() As String) As Long
        On Error GoTo er
        Dim aDate1 As Long, aDate2 As Long, interv As Long, rs As Long
        aDate1 = PlayerDate_DateToSeconds(strDate1, arrParams)
        aDate2 = PlayerDate_DateToSeconds(strDate2, arrParams)
        interv = aDate2 - aDate1 'разница в секундах

        If dInterval = DateInterval.Second Then
            Return interv
        End If

        'получаем размерность минуты, часа, суток
        Dim hoursPerDay As Long = 0, minutesPerHour As Long = 0, secondsPerMinute As Long = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        Select Case dInterval
            Case DateInterval.Minute
                Return Math.Ceiling(interv / secondsPerMinute)
            Case DateInterval.Hour
                Return Math.Ceiling(interv / secondsPerMinute / minutesPerHour)
            Case DateInterval.Day, DateInterval.DayOfYear
                rs = CLng(Math.Ceiling(interv / secondsPerMinute / minutesPerHour / hoursPerDay))
                Return rs
            Case DateInterval.Weekday, DateInterval.WeekOfYear
                Dim lstWeekDays As List(Of String) = PlayerDate_Get_Weekdays(arrParams)
                Return Math.Ceiling(interv / secondsPerMinute / minutesPerHour / hoursPerDay / lstWeekDays.Count)
        End Select

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        Select Case dInterval
            Case DateInterval.Month, DateInterval.Quarter
                Dim mCount As Integer = 0, k As Integer = 1
                If interv < 0 Then
                    interv *= -1
                    k = -1
                End If

                Dim secondsPerDay As Long = secondsPerMinute * minutesPerHour * hoursPerDay
                Dim secondsPerYear As Long = secondsPerDay * lstMonthes.Values.Sum
                mCount = Math.Ceiling(interv / secondsPerYear) * lstMonthes.Count  'количество годов * количество месяцев в году
                interv -= secondsPerYear * mCount / lstMonthes.Count 'теперь остаются только месяцы этого года

                Dim month2 As Integer
                'month1 = PlayerDate_GetPeriod(DateInterval.Month, arrParams, aDate1.ToString) - 1
                month2 = PlayerDate_GetPeriod(DateInterval.Month, arrParams, aDate2.ToString) - 1


                Dim nextMonth As Integer = month2
                Do
                    interv -= lstMonthes.ElementAt(nextMonth).Value * secondsPerDay
                    If interv < 0 Then Exit Do
                    mCount += 1
                    nextMonth -= 1
                    If nextMonth < 0 Then nextMonth = lstMonthes.Count - 1
                Loop
                mCount = mCount * k
                If dInterval = DateInterval.Quarter Then mCount = Math.Ceiling(mCount / 3)
                Return mCount

            Case DateInterval.Year
                Return Math.Ceiling(interv / secondsPerMinute / minutesPerHour / hoursPerDay / lstMonthes.Values.Sum)
        End Select

        Return 0
er:
        MsgBox(Err.Description)
    End Function

    ''' <summary> Добавляет к текущей игровой дате необходимое количество часов, минут и секунд, что бы время равнялось указанному</summary>
    ''' <param name="strTimeTo">время в формате "HH:mm:ss", в которое надо "проснуться" (без кавычек)</param>
    ''' <param name="arrParams"></param>
    ''' <returns>массив из 4 элементов. Первый элемент (с индексом 0) хранит добавленные часы, второй - минуты, третий - секунды, а четвертый - значение True или False, 
    ''' показывающий наступил ли следующий ден</returns>
    Public Function PlayerDate_SleepTill(ByVal strTimeTo As String, ByRef arrParams() As String) As cVariable.variableEditorInfoType
        'получаем новое время
        Dim var As New cVariable.variableEditorInfoType
        ReDim var.arrValues(0)
        var.arrValues(0) = ""

        Dim arr1() As String = strTimeTo.Split(":"c)
        If arr1.Count <> 3 Then
            _ERROR("Неправильный формат даты", "PlayerDate_SpleepTill")
            Return var
        End If

        Dim newHours As Integer = Val(arr1(0))
        Dim newMinutes As Integer = Val(arr1(1))
        Dim newSeconds As Integer = Val(arr1(2))

        'получаем размерность минуты, часа, суток
        Dim hoursPerDay As Integer = 0, minutesPerHour As Integer = 0, secondsPerMinute As Integer = 0
        PlayerDate_GetTimeIntervals(hoursPerDay, minutesPerHour, secondsPerMinute, arrParams)

        'получаем игровые месяцы
        Dim lstMonthes As Dictionary(Of String, Integer) = PlayerDate_Get_Monthes(arrParams)

        'получаем количество секунд в дне
        Dim secondsPerDay As Long = secondsPerMinute * minutesPerHour * hoursPerDay
        'получаем количество секунд в году
        Dim secondsPerYear As Long = lstMonthes.Values.Sum * secondsPerDay

        'получили годы
        Dim years As Integer = Math.Floor(GVARS.PLAYER_TIME / secondsPerYear)

        Dim timeLeft As Long = GVARS.PLAYER_TIME - years * secondsPerYear

        'получили месяцы
        Dim monthes As Integer = 0
        For i As Integer = 0 To lstMonthes.Count - 1
            Dim secondsInMonth As Integer = lstMonthes.ElementAt(i).Value * secondsPerDay
            If timeLeft > secondsInMonth Then
                timeLeft -= secondsInMonth
            Else
                monthes = i '+ 1
                Exit For
            End If
        Next i

        'получили дни
        Dim days As Integer = Math.Floor(timeLeft / secondsPerDay)
        timeLeft -= days * secondsPerDay

        'получили часы
        Dim secondsPerHour = secondsPerMinute * minutesPerHour
        Dim hours As Integer = Math.Floor(timeLeft / secondsPerHour)
        timeLeft -= hours * secondsPerHour

        'получили минуты и секунды
        Dim minutes As Integer = Math.Floor(timeLeft / secondsPerMinute)
        Dim seconds As Integer = timeLeft - minutes * secondsPerMinute

        'считаем старое и новое время в секундах
        Dim oldTime As Long = seconds + minutes * secondsPerMinute + hours * secondsPerHour
        Dim newTime As Long = newSeconds + newMinutes * secondsPerMinute + newHours * secondsPerHour
        Dim isNewDay As Boolean = False

        If newTime <= oldTime Then
            'наступил новый день
            newTime += secondsPerDay
            isNewDay = True
        End If

        'получаем разницу и устанавливаем новое время
        Dim diff As Long = newTime - oldTime
        GVARS.PLAYER_TIME += diff

        'возвращаем количество добавленных часов, минут, секунд и наступил ли новый день
        hours = Math.Floor(diff / secondsPerHour)
        diff -= hours * secondsPerHour

        minutes = Math.Floor(diff / secondsPerMinute)
        seconds = diff - minutes * secondsPerMinute

        var.arrValues = {hours.ToString, minutes.ToString, seconds.ToString, isNewDay.ToString}
        var.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
        Dim arrSign() As String = {"hour", "minute", "second", "newDay"}
        For i As Integer = 0 To arrSign.Count - 1
            var.lstSingatures.Add(arrSign(i), i)
        Next
        Return var
    End Function

    ''' <summary> Возвращает текущую дату в нужном формате</summary>
    Private Function PlayerDate_FormatDate(ByVal dFormat As String, ByVal year As Integer, month As Integer, day As Integer, hour As Integer, minute As Integer, second As Integer, _
                                           ByRef lstMonthes As Dictionary(Of String, Integer), ByRef lstWeekDays As List(Of String), ByVal hoursPerDay As Integer, ByRef arrParams() As String) As String
        'допустимуе форматы:
        '% - неформатный символ (например %h - это не часы, а просто "h"
        'd dd ddd dddd - 1 01 Sun Sunday
        'M MM MMM MMM - 1 01 Jan January
        'h hh H HH часы 12 / часы 24
        'm mm - минуты 
        's ss - секунды
        't tt A/P AM/PM
        'y yy yyy yyyy - формат года: 1 01 2001 2001
        dFormat = UnWrapString(dFormat)
        If String.IsNullOrEmpty(dFormat) Then Return ""
        Dim fLength As Integer = dFormat.Length

        Dim pos As Integer = 0
        Dim result As New System.Text.StringBuilder
        Do While pos < fLength
            Dim isLast As Boolean = (pos = fLength - 1)
            Dim ch As Char = dFormat.Chars(pos)
            Select Case ch
                Case "%"c
                    If isLast Then
                        result.Append(ch)
                    Else
                        pos += 1
                        result.Append(dFormat.Chars(pos))
                    End If
                Case "d"c
                    Dim symbols As Integer = 1
                    Do
                        If isLast = False AndAlso dFormat.Chars(pos + 1) = ch Then
                            symbols += 1
                            pos += 1
                            isLast = (pos = fLength - 1)
                            If symbols >= 4 Then Exit Do
                        Else
                            Exit Do
                        End If
                    Loop

                    Select Case symbols
                        Case 1
                            result.Append(day.ToString)
                        Case 2
                            result.Append(day.ToString.PadLeft(2, "0"c))
                        Case 3
                            Dim weekNum As Integer = PlayerDate_GetPeriod(DateInterval.Weekday, arrParams)
                            Dim wDay As String = lstWeekDays(weekNum - 1)
                            If wDay.Length > 3 Then wDay = wDay.Substring(0, 3)
                            result.Append(wDay)
                        Case 4
                            Dim weekNum As Integer = PlayerDate_GetPeriod(DateInterval.Weekday, arrParams)
                            result.Append(lstWeekDays(weekNum - 1))
                    End Select
                Case "M"c 'month
                    Dim symbols As Integer = 1
                    Do
                        If isLast = False AndAlso dFormat.Chars(pos + 1) = ch Then
                            symbols += 1
                            pos += 1
                            isLast = (pos = fLength - 1)
                            If symbols >= 4 Then Exit Do
                        Else
                            Exit Do
                        End If
                    Loop

                    Select Case symbols
                        Case 1
                            result.Append(month.ToString)
                        Case 2
                            result.Append(month.ToString.PadLeft(2, "0"c))
                        Case 3
                            Dim monthName As String = lstMonthes.ElementAt(month - 1).Key
                            If monthName.Length > 3 Then monthName = monthName.Substring(0, 3)
                            result.Append(monthName)
                        Case 4
                            result.Append(lstMonthes.ElementAt(month - 1).Key)
                    End Select
                Case "y"c, "Y"c
                    Dim symbols As Integer = 1
                    Do
                        If isLast = False AndAlso dFormat.Chars(pos + 1) = ch Then
                            symbols += 1
                            pos += 1
                            isLast = (pos = fLength - 1)
                            If symbols >= 4 Then Exit Do
                        Else
                            Exit Do
                        End If
                    Loop

                    Select Case symbols
                        Case 1
                            result.Append(year.ToString.Last)
                        Case 2
                            Dim yr As String = year.ToString
                            If yr.Length > 2 Then yr = Right(yr, 2)
                            result.Append(yr)
                        Case 3, 4
                            result.Append(year.ToString)
                    End Select
                Case "H"c
                    Dim symbols As Integer = 1
                    If isLast = False AndAlso dFormat.Chars(pos + 1) = ch Then
                        symbols += 1
                        pos += 1
                        isLast = (pos = fLength - 1)
                    End If

                    Select Case symbols
                        Case 1
                            result.Append(hour.ToString)
                        Case 2
                            result.Append(hour.ToString.PadLeft(2, "0"c))
                    End Select
                Case "h"c
                    Dim half As Integer = Math.Round(hoursPerDay / 2)
                    Dim hour12 As String = ""
                    If hour > half Then
                        hour12 = (hour - half).ToString
                    Else
                        hour12 = hour.ToString
                    End If

                    Dim symbols As Integer = 1
                    If isLast = False AndAlso dFormat.Chars(pos + 1) = ch Then
                        symbols += 1
                        pos += 1
                        isLast = (pos = fLength - 1)
                    End If

                    Select Case symbols
                        Case 1
                            result.Append(hour12.ToString)
                        Case 2
                            result.Append(hour12.ToString.PadLeft(2, "0"c))
                    End Select
                Case "m"c
                    Dim symbols As Integer = 1
                    If isLast = False AndAlso dFormat.Chars(pos + 1) = ch Then
                        symbols += 1
                        pos += 1
                        isLast = (pos = fLength - 1)
                    End If

                    Select Case symbols
                        Case 1
                            result.Append(minute.ToString)
                        Case 2
                            result.Append(minute.ToString.PadLeft(2, "0"c))
                    End Select
                Case "s"c
                    Dim symbols As Integer = 1
                    If isLast = False AndAlso dFormat.Chars(pos + 1) = ch Then
                        symbols += 1
                        pos += 1
                        isLast = (pos = fLength - 1)
                    End If

                    Select Case symbols
                        Case 1
                            result.Append(second.ToString)
                        Case 2
                            result.Append(second.ToString.PadLeft(2, "0"c))
                    End Select
                Case "t"c
                    Dim aStr As String = "A"
                    Dim half As Integer = Math.Round(hoursPerDay / 2)
                    If hour > half Then aStr = "P"

                    Dim symbols As Integer = 1
                    If isLast = False AndAlso dFormat.Chars(pos + 1) = ch Then
                        symbols += 1
                        pos += 1
                        isLast = (pos = fLength - 1)
                    End If

                    Select Case symbols
                        Case 1
                            result.Append(aStr)
                        Case 2
                            result.Append(aStr & "M")
                    End Select
                Case Else
                    result.Append(ch)
            End Select

            pos += 1
        Loop

        Return result.ToString
    End Function
#End Region

#Region "Objects"
    ''' <summary>objType As Integer = 0 '0 - обычный, 1 - контейнер, 2 - описание, 3 - разделитель</summary>
    Friend Enum ObjectTypeEnum As Byte
        USUAL = 0
        CONTAINER = 1
        DESCRIPTION = 2
        SEPARATOR = 3
    End Enum

    ''' <summary>showStyle 0 - только название предмета, 1 - название и картинка, 2 - название, картинка и число, 3 - картинка и число</summary>
    Friend Enum ObjectWindowShowStyle As Integer
        NONE = -1
        CAPTION_ONLY = 0
        CAPTION_IMAGE = 1
        CAPTION_IMAGE_COUNT = 2
        IMAGE_COUNT = 3
    End Enum

    ''' <summary>ContainedStyle: 0 - как и вне хранилища, 1 - только картинки, 2 - только картинки с количеством</summary>
    Friend Enum ObjectContainedStyleEnum As Byte
        USUAL_AS_OUTSIDE = 0
        IMAGE_ONLY = 1
        IMAGE_AND_COUNT = 2
    End Enum

    ''' <summary>
    ''' Подготавливает начальную структуру для вставки предметов в окно предметов
    ''' </summary>
    ''' <param name="arrParams"></param>
    Public Sub ObjectWindowPrepareStruct(ByRef arrParams() As String)
        Dim hDoc As HtmlDocument = frmPlayer.wbObjects.Document
        If IsNothing(hDoc) Then Return
        Dim hConvas As HtmlElement = hDoc.GetElementById("ObjectsConvas")
        If IsNothing(hConvas) Then Return
        hConvas.InnerHtml = ""

        'добавляем дивы для каждого предмета
        Dim classO As Integer = mScript.mainClassHash("O")
        Dim objCount As Integer = 0
        If IsNothing(mScript.mainClass(classO).ChildProperties) = False Then objCount = mScript.mainClass(classO).ChildProperties.Count

        For oId As Integer = 0 To objCount - 1
            'наружный контейнер предмета
            Dim outerDiv As HtmlElement = hDoc.CreateElement("DIV")
            outerDiv.SetAttribute("ClassName", "OuterContainer")
            outerDiv.Id = "object" & oId.ToString
            outerDiv.Style = "display:none"
            hConvas.AppendChild(outerDiv)

            'див для хранения содержимого предмета если он - контейнер
            Dim objType As Integer = 0
            ReadPropertyInt(classO, "Type", oId, -1, objType, arrParams) '1 - контейнер
            If objType = 1 Then
                Dim containerDiv As HtmlElement = hDoc.CreateElement("DIV")
                containerDiv.SetAttribute("ClassName", "InnerContainer")
                outerDiv.AppendChild(containerDiv)
            End If
        Next oId
    End Sub

    Public Sub ObjectsAppendAll(ByRef arrParams() As String)
        Dim hDoc As HtmlDocument = frmPlayer.wbObjects.Document
        If IsNothing(hDoc) Then Return
        Dim hConvas As HtmlElement = hDoc.GetElementById("ObjectsConvas")
        If IsNothing(hConvas) Then Return

        Dim classOW As Integer = mScript.mainClassHash("OW")
        Dim showStyle As Integer = 0 '0 - только название предмета, 1 - название и картинка, 2 - название, картинка и число, 3 - только картинка и количество
        ReadPropertyInt(classOW, "Style", -1, -1, showStyle, arrParams)

        Dim classO As Integer = mScript.mainClassHash("O")
        Dim objCount As Integer = 0
        If IsNothing(mScript.mainClass(classO).ChildProperties) = False Then objCount = mScript.mainClass(classO).ChildProperties.Count

        For oId As Integer = 0 To objCount - 1
            ObjectSetAppearance(oId, arrParams, showStyle, hDoc, hConvas)
        Next oId

    End Sub

    Public Function ObjectSetAppearance(ByVal oId As Integer, ByRef arrParams() As String, Optional ByVal showStyle As ObjectWindowShowStyle = ObjectWindowShowStyle.NONE, _
                                        Optional ByRef hDoc As HtmlDocument = Nothing, _
                                        Optional ByRef hConvas As HtmlElement = Nothing) As String
        If IsNothing(hDoc) Then hDoc = frmPlayer.wbObjects.Document
        If IsNothing(hDoc) Then Return _ERROR("Не удалось получить доступ к окну предметов.")
        If IsNothing(hConvas) Then hConvas = hDoc.GetElementById("ObjectsConvas")
        If IsNothing(hConvas) Then Return _ERROR("Не удалось получить доступ к окну предметов.")

        Dim outerDiv As HtmlElement = hDoc.GetElementById("object" & oId.ToString)
        If IsNothing(outerDiv) Then Return _ERROR("Нарушена структура внутри окна предметов. Html-контейнер предмета не найден.")

        'очищаем хранилище предмета
        Dim hContainerClone As mshtml.IHTMLDOMNode = Nothing
        If outerDiv.Children.Count > 0 Then
            Dim lastChild As HtmlElement = outerDiv.Children(outerDiv.Children.Count - 1)
            If lastChild.GetAttribute("ClassName") = "ContainerTab" Then lastChild = outerDiv.Children(outerDiv.Children.Count - 2)
            If HTMLHasClass(lastChild, "InnerContainer") Then
                'в предмете есть контейнер, внутри которого могут находится другие предметы. Их трогать нельзя
                Dim hCont As mshtml.IHTMLDOMNode = lastChild.DomElement
                hContainerClone = hCont.cloneNode(True)
            Else
                'очищаем все
                'outerDiv.InnerHtml = ""
            End If
            'очищаем все
            outerDiv.InnerHtml = ""
        End If

        Dim classO As Integer = mScript.mainClassHash("O")
        Dim classOW As Integer = mScript.mainClassHash("OW")

        'стиль отображения в окне предметов
        If showStyle = ObjectWindowShowStyle.NONE Then
            ReadPropertyInt(classOW, "Style", -1, -1, showStyle, arrParams) '0 - только название предмета, 1 - название и картинка, 2 - название, картинка и число, 3 - картинка и число
        End If

        'тип предмета
        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL  '0 - обычный, 1 - контейнер, 2 - описание, 3 - разделитель
        ReadPropertyInt(classO, "Type", oId, -1, objType, arrParams)

        'получаем контейнер, в который поместить предмет
        Dim newContainer As String = ""
        ReadProperty(classO, "Container", oId, -1, newContainer, arrParams)
        If String.IsNullOrEmpty(newContainer) = False Then
            Dim containerId As Integer = GetSecondChildIdByName(newContainer, mScript.mainClass(classO).ChildProperties)
            If containerId >= 0 Then
                newContainer = "object" & containerId.ToString
            Else
                newContainer = ""
            End If
        End If

        'получаем тип размещения в контейнере containedStyle
        'ContainedStyle: 0 - как и вне хранилища, 1 - только картинки, 2 - только картинки с количеством
        Dim containedStyle As ObjectContainedStyleEnum = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE  'стиль отображения предметов в контейнере
        If newContainer.StartsWith("object") Then
            Dim contId As Integer = Val(newContainer.Substring(6))
            ReadPropertyInt(classO, "ContainedStyle", contId, -1, containedStyle, arrParams)
        End If

        'перемещаем предмет в родительский контейнер
        Dim oldContainer As String = outerDiv.Parent.Id
        If String.IsNullOrEmpty(oldContainer) Then oldContainer = outerDiv.Parent.Parent.Id
        If String.Compare(oldContainer, newContainer, True) <> 0 AndAlso (newContainer.Length = 0 AndAlso oldContainer = "ObjectsConvas") = False AndAlso _
            (oldContainer.Length = 0 AndAlso newContainer = "ObjectsConvas") = False Then
            'получаем стиль отображения предметов в контейнере

            'меняем контейнер элемента
            'удаляем из старого места
            Dim hOuterDiv As mshtml.IHTMLDOMNode = outerDiv.DomElement

            If String.IsNullOrEmpty(oldContainer) OrElse oldContainer = "ObjectsConvas" Then
                'вставляем метку для возврата на место
                Dim hParent As mshtml.IHTMLDOMNode = hOuterDiv.parentNode
                Dim hMark As HtmlElement = hDoc.CreateElement("SPAN")
                hMark.Style = "display:none"
                hMark.SetAttribute("placeHolder", mScript.mainClass(classO).ChildProperties(oId)("Name").Value)
                hParent.insertBefore(hMark.DomElement, hOuterDiv)
            End If

            hOuterDiv.removeNode(True)
            'вставляем в новое

            If String.IsNullOrEmpty(newContainer) Then
                'возвращаем предмет в окно вне контейнера
                'ищем метку
                Dim hMark As HtmlElement = Nothing
                Dim oName As String = mScript.mainClass(classO).ChildProperties(oId)("Name").Value
                For i As Integer = 0 To hConvas.Children.Count - 1
                    Dim hEl As HtmlElement = hConvas.Children(i)
                    If hEl.GetAttribute("placeHolder") = oName Then
                        hMark = hEl
                        Exit For
                    End If
                Next i
                If IsNothing(hMark) Then Return _ERROR("Нарушена структура внутри окна предметов. Html-маркер положения предмета не найден.")

                'создаем новый наружный контейнер предмета
                outerDiv = hDoc.CreateElement("DIV")
                outerDiv.SetAttribute("ClassName", "OuterContainer")
                outerDiv.Id = "object" & oId.ToString
                outerDiv.Style = "display:none"

                'див для хранения содержимого предмета если он - контейнер
                If objType = ObjectTypeEnum.CONTAINER AndAlso IsNothing(hContainerClone) Then
                    Dim containerDiv As HtmlElement = hDoc.CreateElement("DIV")
                    containerDiv.SetAttribute("ClassName", "InnerContainer")
                    outerDiv.AppendChild(containerDiv)
                End If

                Dim hConvasDOM As mshtml.IHTMLDOMNode = hConvas.DomElement
                Dim hMarkDOM As mshtml.IHTMLDOMNode = hMark.DomElement
                hConvasDOM.insertBefore(outerDiv.DomElement, hMarkDOM)
                hMarkDOM.removeNode(True)
                'hConvas.AppendChild(outerDiv)
            Else
                'вставляем в новый контейнер
                Dim parentContainer As HtmlElement = hDoc.GetElementById(newContainer)
                If IsNothing(parentContainer) Then Return _ERROR("Нарушена структура внутри окна предметов. Html-контейнер предмета не найден. Возможно, была попытка циклической вставки предметов (вставка в контейнер, являющийся дочерним контейнером).")
                If parentContainer.Children.Count < 2 Then
                    If HTMLHasClass(parentContainer.Children(0), "InnerContainer") Then
                        parentContainer = parentContainer.Children(0) 'контейнер предметов
                    Else
                        Return _ERROR("Не удалось вставить предмет в контейнер. Вероятно, предмет не являлся контейнером.")
                    End If
                Else
                    parentContainer = parentContainer.Children(1) 'контейнер предметов
                End If

                'ставляем новый элемент в контейнер перед тем, который имеет больший Id (чтобы порядок отображения элементов был таким же как и в редакторе)
                If parentContainer.Children.Count = 0 Then
                    'в контейнер пока пусто - просто вставляем новый предмет
                    parentContainer.AppendChild(outerDiv)
                Else
                    Dim elementBefore As mshtml.IHTMLDOMNode = Nothing
                    For i As Integer = 0 To parentContainer.Children.Count - 1
                        Dim El As HtmlElement = parentContainer.Children(i)
                        Dim elName As String = El.Id
                        If String.IsNullOrEmpty(elName) OrElse elName.StartsWith("object") = False Then Continue For
                        Dim elId As Integer = Val(elName.Substring(6))
                        If elId > oId Then
                            'получен элемент с большим Id
                            elementBefore = El.DomElement
                            Exit For
                        End If
                    Next i

                    If IsNothing(elementBefore) Then
                        'предмета с большим чем у текущего предмета Id в контейнере нет
                        parentContainer.AppendChild(outerDiv)
                    Else
                        'предмет с большим Id найден - вставляем перед ним
                        Dim hPar As mshtml.IHTMLDOMNode = parentContainer.DomElement
                        hPar.insertBefore(outerDiv.DomElement, elementBefore)
                    End If
                End If
            End If

        End If

        If containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
            outerDiv.SetAttribute("ClassName", "OuterContainer")
        Else
            outerDiv.SetAttribute("ClassName", "OuterContainerInStorage")
        End If

        'добавлен ли персонажу
        Dim belongsToPlayer As Boolean = False
        ReadPropertyBool(classO, "BelongsToPlayer", oId, -1, belongsToPlayer, arrParams)

        'visible
        Dim visible As Boolean = False
        If belongsToPlayer Then ReadPropertyBool(classO, "Visible", oId, -1, visible, arrParams)

        If visible Then
            outerDiv.Style = "" '"display:block"
        Else
            outerDiv.Style = "display:none"
        End If

        'enabled
        Dim objEnabled As Boolean = True
        ReadPropertyBool(classO, "Enabled", oId, -1, objEnabled, arrParams)
        If Not objEnabled Then
            outerDiv.SetAttribute("disabled", "disabled")
        ElseIf outerDiv.GetAttribute("disabled") <> "False" Then
            Dim hOuter As mshtml.IHTMLElement = outerDiv.DomElement
            hOuter.removeAttribute("disabled")
        End If

        'на стиль отображения данного предмета влияют три характеристики
        'showStyle - глобальный стиль отображения предметов (строка, текст и картинка...)
        'objType - тип данного предмета (обычный, описание..)
        'containedStyle - тип отображения в контейнере
        'а также исчисляемость, наличие картинки и т. д.

        Dim objCaption As String = "", objPicture As String = "", objPicWidth As String = "", objPicHeight As String = "", objCount As String = "", isCountable As Boolean = False
        ReadProperty(classO, "Caption", oId, -1, objCaption, arrParams)
        objCaption = UnWrapString(objCaption)
        ReadProperty(classO, "Picture", oId, -1, objPicture, arrParams)
        objPicture = UnWrapString(objPicture)
        If String.IsNullOrEmpty(objPicture) = False Then
            objPicture = objPicture.Replace("\"c, "/"c)
            If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, objPicture)) Then
                ReadProperty(classO, "PicWidth", oId, -1, objPicWidth, arrParams)
                objPicWidth = UnWrapString(objPicWidth)
                If IsNumeric(objPicWidth) Then objPicWidth &= "px"

                ReadProperty(classO, "PicHeight", oId, -1, objPicHeight, arrParams)
                objPicHeight = UnWrapString(objPicHeight)
                If IsNumeric(objPicHeight) Then objPicHeight &= "px"
            Else
                objPicture = ""
            End If
        End If
        ReadPropertyBool(classO, "Countable", oId, -1, isCountable, arrParams)
        If isCountable Then
            'получаем количество
            ReadProperty(classO, "Count", oId, -1, objCount, arrParams)

            'получаем формат вывода количества
            Dim countFormat As String = ""
            ReadProperty(classO, "CountFormat", oId, -1, countFormat, arrParams)
            countFormat = UnWrapString(countFormat)

            If String.IsNullOrEmpty(countFormat) = False Then
                Dim dblCnt As Double = 0
                If Double.TryParse(objCount, System.Globalization.NumberStyles.Float, provider_points, dblCnt) Then
                    objCount = Format(dblCnt, countFormat)
                End If
            End If
        End If

        Dim parentForObjInfo As HtmlElement 'поледний родительский элемент, в котором непосредственно содержится инфо о предмете
        Dim hObjText As HtmlElement = Nothing, hObjImg As HtmlElement = Nothing, hObjCount As HtmlElement = Nothing
        Dim hTab As HtmlElement = Nothing 'элемент для закладки

        'подготавливаем структуру для отображения
        If objType = ObjectTypeEnum.SEPARATOR Then
            'разделитель - создаем для него HR / SPAN
            If containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
                'обычный вид
                parentForObjInfo = hDoc.CreateElement("HR")
            Else
                'вертикальный сепаратор внутри контейнера
                parentForObjInfo = hDoc.CreateElement("SPAN")
                If showStyle = ObjectWindowShowStyle.CAPTION_ONLY Then
                    parentForObjInfo.SetAttribute("ClassName", "SeparatorVerticalTextOnly")
                Else
                    parentForObjInfo.SetAttribute("ClassName", "SeparatorVertical")
                End If
            End If
            outerDiv.AppendChild(parentForObjInfo)
        ElseIf objType = ObjectTypeEnum.DESCRIPTION Then
            'описание. Создаем для него див/спан
            If containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
                'располагается как обычно 
                parentForObjInfo = hDoc.CreateElement("DIV")
                parentForObjInfo.SetAttribute("ClassName", "ObjectInfo")
            Else
                'расположение внутри контейнера
                parentForObjInfo = hDoc.CreateElement("SPAN")
                parentForObjInfo.SetAttribute("ClassName", "ObjectAtStorageInfo")
            End If
            hObjText = parentForObjInfo
            outerDiv.AppendChild(parentForObjInfo)
        ElseIf containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
            '0 - обычный, 1 - контейнер. Тип размещения - обычный
            'ShowStyle: 0 - только название предмета, 1 - название и картинка, 2 - название, картинка и число, 3 - картинка и число
            If showStyle = ObjectWindowShowStyle.CAPTION_ONLY Then
                'контейнер для всей инфы - див
                parentForObjInfo = hDoc.CreateElement("DIV")
                hObjText = parentForObjInfo
                outerDiv.AppendChild(parentForObjInfo)
            Else
                'получаем размер места под картинки
                Dim imgSpace As String = ""
                ReadProperty(classOW, "SpaceForImages", -1, -1, imgSpace, arrParams)
                imgSpace = UnWrapString(imgSpace).Trim
                If IsNumeric(imgSpace) Then imgSpace &= "px"

                'контейнер для всей инфы о предмете - таблица
                parentForObjInfo = hDoc.CreateElement("TABLE")
                outerDiv.AppendChild(parentForObjInfo)
                Dim hBody As HtmlElement = hDoc.CreateElement("TBODY")
                parentForObjInfo.AppendChild(hBody)

                Dim TR As HtmlElement = hDoc.CreateElement("TR")
                hBody.AppendChild(TR)
                Dim TD As HtmlElement = hDoc.CreateElement("TD")
                TR.AppendChild(TD)

                hObjImg = hDoc.CreateElement("IMG")
                TD.AppendChild(hObjImg)
                If String.IsNullOrEmpty(imgSpace) = False Then TD.Style = "width:" & imgSpace

                TD = hDoc.CreateElement("TD")
                TR.AppendChild(TD)
                If showStyle = ObjectWindowShowStyle.IMAGE_COUNT Then
                    hObjCount = TD
                Else
                    hObjText = TD
                End If

                If showStyle = ObjectWindowShowStyle.CAPTION_IMAGE_COUNT Then
                    TD = hDoc.CreateElement("TD")
                    TR.AppendChild(TD)
                    hObjCount = TD
                End If
            End If

            If objType = ObjectTypeEnum.CONTAINER Then
                parentForObjInfo.SetAttribute("ClassName", "ObjectInfoForContainers")
                'контейнер. Создаем хранилище для других предметов
                If IsNothing(hContainerClone) = False Then
                    'возвращаем контейнер с хранящимися предметами на место
                    Dim hOuter As mshtml.IHTMLDOMNode = outerDiv.DomElement
                    hOuter.appendChild(hContainerClone)
                    'делаем контейнер видимым
                    Dim containerDiv As HtmlElement = outerDiv.Children(outerDiv.Children.Count - 1)
                    If IsNothing(containerDiv) = False Then
                        If objEnabled Then
                            containerDiv.Style = ""
                        Else
                            containerDiv.Style = "display:none"
                        End If
                    End If

                Else
                    Dim containerDiv As HtmlElement = hDoc.CreateElement("DIV")
                    containerDiv.SetAttribute("ClassName", "InnerContainer")
                    outerDiv.AppendChild(containerDiv)
                    If IsNothing(containerDiv) = False Then
                        If objEnabled Then
                            containerDiv.Style = ""
                        Else
                            containerDiv.Style = "display:none"
                        End If
                    End If
                End If


                'создаем закладку
                If containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
                    hTab = hDoc.CreateElement("DIV")
                    hTab.SetAttribute("ClassName", "ContainerTab")
                    hTab.InnerHtml = "<hr/>"

                    outerDiv.AppendChild(hTab)
                    If Not objEnabled Then hTab.Style = "display:none"
                End If
            Else
                parentForObjInfo.SetAttribute("ClassName", "ObjectInfo")
            End If
        ElseIf containedStyle = ObjectContainedStyleEnum.IMAGE_ONLY Then
            '0 - обычный, 1 - контейнер. Тип размещения - только картинки. Создаем IMG/SPAN
            'ShowStyle: 0 - только название предмета, 1 - название и картинка, 2 - название, картинка и число, 3 - картинка и число
            If showStyle = ObjectWindowShowStyle.CAPTION_ONLY Then
                parentForObjInfo = hDoc.CreateElement("SPAN")
                hObjText = parentForObjInfo
            Else
                parentForObjInfo = hDoc.CreateElement("IMG")
                hObjImg = parentForObjInfo
            End If
            parentForObjInfo.SetAttribute("ClassName", "ObjectAtStorageInfo")
            outerDiv.AppendChild(parentForObjInfo)

            If objType = ObjectTypeEnum.CONTAINER Then
                'контейнер. Создаем хранилище для других предметов
                If IsNothing(hContainerClone) = False Then
                    'возвращаем контейнер с хранящимися предметами на место
                    Dim hOuter As mshtml.IHTMLDOMNode = outerDiv.DomElement
                    hOuter.appendChild(hContainerClone)
                    'делаем контейнер невидимым
                    Dim containerDiv As HtmlElement = outerDiv.Children(outerDiv.Children.Count - 1)
                    containerDiv.Style = "display:none"
                Else
                    'контейнер. Создаем хранилище для других предметов но делаем его невидимым
                    Dim containerDiv As HtmlElement = hDoc.CreateElement("DIV")
                    containerDiv.SetAttribute("ClassName", "InnerContainer")
                    containerDiv.Style = "display:none"
                    outerDiv.AppendChild(containerDiv)
                End If

                If containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
                    hTab = hDoc.CreateElement("DIV")
                    hTab.SetAttribute("ClassName", "ContainerTab")
                    hTab.InnerHtml = "<hr/>"

                    outerDiv.AppendChild(hTab)
                    If Not objEnabled Then hTab.Style = "display:none"
                End If
            End If
        Else
            'containedStyle=ObjectContainedStyleEnum.IMAGE_AND_COUNT 
            '0 - обычный, 1 - контейнер. Тип размещения - только картинки с количеством. Созадем картинку и спаны
            If showStyle = ObjectWindowShowStyle.CAPTION_ONLY Then
                parentForObjInfo = hDoc.CreateElement("SPAN")
                hObjText = parentForObjInfo
            Else
                parentForObjInfo = hDoc.CreateElement("SPAN")
                Dim noBr As HtmlElement = hDoc.CreateElement("NOBR")
                parentForObjInfo.AppendChild(noBr)

                hObjImg = hDoc.CreateElement("IMG")
                noBr.AppendChild(hObjImg)
                hObjCount = hDoc.CreateElement("SPAN")
                noBr.AppendChild(hObjCount)
            End If
            parentForObjInfo.SetAttribute("ClassName", "ObjectAtStorageInfo")
            outerDiv.AppendChild(parentForObjInfo)

            If objType = ObjectTypeEnum.CONTAINER Then
                'контейнер. Создаем хранилище для других предметов
                If IsNothing(hContainerClone) = False Then
                    'возвращаем контейнер с хранящимися предметами на место
                    Dim hOuter As mshtml.IHTMLDOMNode = outerDiv.DomElement
                    hOuter.appendChild(hContainerClone)
                    'делаем контейнер невидимым
                    Dim containerDiv As HtmlElement = outerDiv.Children(outerDiv.Children.Count - 1)
                    containerDiv.Style = "display:none"
                Else
                    'контейнер. Создаем хранилище для других предметов но делаем его невидимым
                    Dim containerDiv As HtmlElement = hDoc.CreateElement("DIV")
                    containerDiv.SetAttribute("ClassName", "InnerContainer")
                    containerDiv.Style = "display:none"
                    outerDiv.AppendChild(containerDiv)
                End If

                If containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
                    hTab = hDoc.CreateElement("DIV")
                    hTab.SetAttribute("ClassName", "ContainerTab")
                    hTab.InnerHtml = "<hr/>"

                    outerDiv.AppendChild(hTab)
                    If Not objEnabled Then hTab.Style = "display:none"
                End If
            End If
        End If

        'устанавливаем классы
        If IsNothing(hObjImg) = False Then hObjImg.SetAttribute("ClassName", "objectImage")
        If IsNothing(hObjCount) = False Then hObjCount.SetAttribute("ClassName", "objectCount")
        If IsNothing(hObjText) = False Then
            If objType = ObjectTypeEnum.DESCRIPTION Then
                hObjText.SetAttribute("ClassName", "objectDescription")
            Else
                hObjText.SetAttribute("ClassName", "objectText")
            End If
        End If

        If IsNothing(hObjText) = False Then
            If objType = ObjectTypeEnum.DESCRIPTION Then
                Dim res As String = ""
                ReadProperty(classO, "Description", oId, -1, res, arrParams)
                res = UnWrapString(res)
                hObjText.InnerHtml = res
            Else
                hObjText.InnerHtml = objCaption
            End If

            If Not objEnabled Then hObjText.SetAttribute("disabled", "disabled")
        End If

        If IsNothing(hObjImg) = False Then
            'вывод на экран картинки
            If String.IsNullOrEmpty(objPicture) Then
                Dim imgStyle As String = "visibility:hidden;"
                If objPicWidth.Length > 0 Then imgStyle &= "width:" & objPicWidth & ";"
                If objPicHeight.Length > 0 Then imgStyle &= "height:" & objPicHeight
                hObjImg.Style = imgStyle
            Else
                hObjImg.Style = ""
                hObjImg.SetAttribute("src", objPicture)
                Dim imgStyle As String = ""
                If objPicWidth.Length > 0 Then imgStyle = "width:" & objPicWidth & ";"
                If objPicHeight.Length > 0 Then imgStyle &= "height:" & objPicHeight
                hObjImg.Style = imgStyle

                If IsNothing(hObjText) Then hObjImg.SetAttribute("Title", objCaption)
            End If
        End If

        If IsNothing(hObjCount) = False AndAlso isCountable Then
            hObjCount.InnerHtml = objCount
        End If

        If objType = ObjectTypeEnum.SEPARATOR Then Return ""

        'устанавливаем класс экипировки
        Dim isEquipment As Boolean = False
        ReadPropertyBool(classO, "IsEquipment", oId, -1, isEquipment, arrParams)
        If isEquipment Then
            'это экипировка
            Dim equipped As Boolean = False
            ReadPropertyBool(classO, "Equipped", oId, -1, equipped, arrParams)
            EquipmentSetClass(oId, arrParams, outerDiv.Children(0), equipped)
            'Dim equipmentClass As String = ""
            'ReadProperty(classO, "EquipmentCSSclass", oId, -1, equipmentClass, arrParams)
            'If String.IsNullOrEmpty(equipmentClass) = False Then
            '    'имеется css-класс для спец. отображения экипировки
            '    If equipped Then
            '        'экипировка одета
            '        Dim equippedClass As String = "equipped" & UnWrapString(equipmentClass).Substring(9)
            '        HTMLAddClass(outerDiv.Children(0), equippedClass)
            '    Else
            '        'экипировка не одета
            '        HTMLAddClass(outerDiv.Children(0), UnWrapString(equipmentClass))
            '    End If
            'End If
        End If

        'События данного предмета
        'RemoveHandler outerDiv.Children(0).MouseOver, AddressOf del_object_MouseOver
        'RemoveHandler outerDiv.Children(0).MouseLeave, AddressOf del_object_MouseLeave
        'RemoveHandler outerDiv.Children(0).MouseDown, AddressOf del_object_MouseDown
        'RemoveHandler outerDiv.Children(0).MouseUp, AddressOf del_object_MouseUp
        'RemoveHandler outerDiv.Children(0).Click, AddressOf del_object_Click

        AddHandler outerDiv.Children(0).MouseOver, AddressOf del_object_MouseOver
        AddHandler outerDiv.Children(0).MouseLeave, AddressOf del_object_MouseLeave
        AddHandler outerDiv.Children(0).MouseDown, AddressOf del_object_MouseDown
        AddHandler outerDiv.Children(0).MouseUp, AddressOf del_object_MouseUp
        AddHandler outerDiv.Children(0).Click, AddressOf del_object_Click

        If IsNothing(hTab) = False Then
            'событие нажатия на вкладку
            hTab.Id = "objTab" & oId.ToString
            AddHandler hTab.Click, AddressOf del_object_Click
        End If

        If IsNothing(hContainerClone) = False AndAlso outerDiv.Children.Count > 1 Then
            'События предметов внутри контейнера данного элемента
            Dim hInner As HtmlElement = outerDiv.Children(1)
            For i As Integer = 0 To outerDiv.Children(1).Children.Count - 1
                Dim hEl As HtmlElement = outerDiv.Children(1).Children(i).Children(0)
                'RemoveHandler hEl.MouseOver, AddressOf del_object_MouseOver
                'RemoveHandler hEl.MouseLeave, AddressOf del_object_MouseLeave
                'RemoveHandler hEl.MouseDown, AddressOf del_object_MouseDown
                'RemoveHandler hEl.MouseUp, AddressOf del_object_MouseUp
                'RemoveHandler hEl.Click, AddressOf del_object_Click

                AddHandler hEl.MouseOver, AddressOf del_object_MouseOver
                AddHandler hEl.MouseLeave, AddressOf del_object_MouseLeave
                AddHandler hEl.MouseDown, AddressOf del_object_MouseDown
                AddHandler hEl.MouseUp, AddressOf del_object_MouseUp
                AddHandler hEl.Click, AddressOf del_object_Click
            Next
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Устанавливает css-класс на предмете экипировки. Проверка явлется ли предмет экипировкой не выполняется
    ''' </summary>
    ''' <param name="oId">Id предмета</param>
    ''' <param name="arrParams"></param>
    ''' <param name="hEl">Id html-элемента предмета outerDiv.Children(0)</param>
    ''' <param name="equipped">одеть или снять</param>
    Public Function EquipmentSetClass(ByVal oId As Integer, ByRef arrParams() As String, ByRef hEl As HtmlElement, ByVal equipped As Boolean) As Boolean
        Dim classO As Integer = mScript.mainClassHash("O")
        Dim equipmentClass As String = ""
        If ReadProperty(classO, "EquipmentCSSclass", oId, -1, equipmentClass, arrParams) = False Then Return False
        If String.IsNullOrEmpty(equipmentClass) = False Then
            'имеется css-класс для спец. отображения экипировки
            equipmentClass = UnWrapString(equipmentClass)
            Dim equippedClass As String = "equipped" & equipmentClass.Substring(9)
            If equipped Then
                'экипировка одета
                HTMLReplaceClass(hEl, equipmentClass, equippedClass)
            Else
                'экипировка не одета
                HTMLReplaceClass(hEl, equippedClass, equipmentClass)
            End If
        End If
        Return True
    End Function

    ''' <summary>
    ''' Изменяет структуру окна предметов при удалении предмета
    ''' </summary>
    ''' <param name="oId">Id удаляемого предмета</param>
    ''' <param name="arrParams"></param>
    ''' <returns></returns>
    Public Function ObjectRemoveElement(ByVal oId As Integer, ByRef arrParams() As String) As String
        Dim classO As Integer = mScript.mainClassHash("O")
        Dim wbDoc As HtmlDocument = frmPlayer.wbObjects.Document
        If IsNothing(wbDoc) Then Return _ERROR("Не удалось получить доступ к окну предметов.")

        'удаляем предмет
        Dim hDoc As mshtml.HTMLDocument = wbDoc.DomDocument
        Dim remEl As mshtml.IHTMLDOMNode = hDoc.getElementById("object" & oId.ToString)
        If IsNothing(remEl) Then Return _ERROR("Нарушена структура html-документа. Удаляемый предмет не найден.", "Object.Remove")
        remEl.removeNode(True)

        'изменяем id предметов выше
        For i As Integer = oId + 1 To mScript.mainClass(classO).ChildProperties.Count - 1
            Dim hEl As mshtml.IHTMLElement = hDoc.getElementById("object" & i.ToString)
            If IsNothing(hEl) Then Return "" 'предметы закончились
            hEl.id = "object" & (i - 1).ToString
            'изменяем id вкладки контейнера
            hEl = hDoc.getElementById("objTab" & i.ToString)
            If IsNothing(hEl) = False Then hEl.id = "objTab" & (i - 1).ToString
        Next i
        Return ""
    End Function

    ''' <summary>
    ''' Возвращает конвас для вывода описаний в DescWindow
    ''' </summary>
    Private Function GetDescriptionHtmlElement() As HtmlElement
        Dim hDoc As HtmlDocument = frmPlayer.wbDescription.Document
        If IsNothing(hDoc) Then Return Nothing
        Dim hConvas As HtmlElement = hDoc.GetElementById("DescriptionConvas")
        Return hConvas
    End Function


    Private Sub del_object_MouseOver(sender As Object, e As HtmlElementEventArgs)
        'sender - ObjectInfo
        'Выводим картинку действия при наведении PictureHover
        Dim classO As Integer = mScript.mainClassHash("O")
        Dim objId As Integer = Val(sender.Parent.Id.Substring(6))
        Dim arrs() As String = {objId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classO, "Enabled", objId, -1, enabled, arrs)
        If enabled = False Then Return

        'получаем тип. '0 - обычный, 1 - контернер, 2 - описание, 3 - разделитель
        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
        ReadPropertyInt(classO, "Type", objId, -1, objType, arrs)
        If objType = ObjectTypeEnum.SEPARATOR Then Return 'для типа сеператор не подходит

        'событие ObjectMouseOver
        Dim res As String = ""
        Dim wasEvent As Boolean = False
        'глобальное
        Dim eventId As Integer = mScript.mainClass(classO).Properties("ObjectMouseOver").eventId
        If eventId > 0 Then
            wasEvent = True
            res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectMouseOver", False)
            If res = "#Error" Then Return
        End If

        'данного предмета
        If res <> "False" Then
            eventId = mScript.mainClass(classO).ChildProperties(objId)("ObjectMouseOver").eventId
            If eventId > 0 Then
                wasEvent = True
                res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectMouseOver", False)
                If res = "#Error" Then Return
            End If
        End If

        'событие ScriptFinishedEvent
        If wasEvent Then
            If mScript.eventRouter.RunScriptFinishedEvent(Nothing) = "#Error" Then Return
        End If
        If res = "False" Then Return


        'получаем PictureHover. Если пусто, то выход
        Dim objPictureHover As String = ""
        If ReadProperty(classO, "PictureHover", objId, -1, objPictureHover, arrs) = False Then Return
        objPictureHover = UnWrapString(objPictureHover)
        If String.IsNullOrEmpty(objPictureHover) = False Then
            Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, objPictureHover)
            If FileIO.FileSystem.FileExists(fPath) = False Then
                MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            'получаем Picture. PictureHover и Picture не должны быть равны
            Dim actPicture As String = ""
            If ReadProperty(classO, "Picture", objId, -1, actPicture, arrs) = False Then Return
            actPicture = UnWrapString(actPicture)

            'убеждаемся в правильном формате пути к картинкам
            objPictureHover = objPictureHover.Replace("\", "/")
            If String.IsNullOrEmpty(actPicture) = False Then actPicture = actPicture.Replace("\", "/")
            If String.Compare(actPicture, objPictureHover, True) <> 0 Then
                'собственно замена картинки Picture на PictureHover
                If objType = ObjectTypeEnum.DESCRIPTION Then
                    'type = описание. Находим первую же картинку и работаем с ней
                    'description - DIV/SPAN-...(IMG)...
                    Dim hEl As HtmlElement = sender
                    Dim hImgCol As HtmlElementCollection = hEl.GetElementsByTagName("IMG")
                    If hImgCol.Count > 0 Then
                        hImgCol(0).SetAttribute("src", objPictureHover)
                    End If
                Else
                    '0 - обычный, 1 - контейнер
                    'возможное расположение картинки: 
                    '1) IMG 2) SPAN-NOBR-IMG 3) TABLE-TBODY-TR-TD
                    Dim hEl As HtmlElement = sender
                    If hEl.TagName <> "IMG" Then
                        Do
                            If hEl.Children.Count = 0 Then
                                hEl = Nothing
                                Exit Do
                            End If
                            hEl = hEl.Children(0)
                            If hEl.TagName = "IMG" Then Exit Do
                        Loop

                        If IsNothing(hEl) = False Then
                            hEl.SetAttribute("Src", objPictureHover)
                        End If
                    End If
                End If
            End If
        End If

        'Вывод описания
        If objType <> ObjectTypeEnum.DESCRIPTION Then
            'получаем html-элемент для вывода описания
            Dim hDoc As HtmlDocument = frmPlayer.wbDescription.Document
            If IsNothing(hDoc) Then Return
            Dim hDesc As HtmlElement = hDoc.GetElementById("DescriptionConvas")
            If IsNothing(hDesc) Then Return

            'получение и вывод описания
            Dim strDesc As String = ""
            If ReadProperty(classO, "Description", objId, -1, strDesc, {objId.ToString}) = False Then Return
            strDesc = UnWrapString(strDesc)
            hDesc.InnerHtml = strDesc
        End If
    End Sub

    Private Sub del_object_MouseLeave(sender As Object, e As HtmlElementEventArgs)
        'sender - ObjectInfo
        'возвращаем картинку действия Picture после того, как она была зменена на PictureHover
        Dim classO As Integer = mScript.mainClassHash("O")
        Dim objId As Integer = Val(sender.Parent.Id.Substring(6))
        Dim arrs() As String = {objId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classO, "Enabled", objId, -1, enabled, arrs)
        If enabled = False Then Return

        'получаем тип. '0 - обычный, 1 - контернер, 2 - описание, 3 - разделитель
        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
        ReadPropertyInt(classO, "Type", objId, -1, objType, arrs)
        If objType = ObjectTypeEnum.SEPARATOR Then Return 'для типа сеператор не подходит

        'событие ObjectMouseLeave
        Dim res As String = ""
        Dim wasEvent As Boolean = False
        'глобальное
        Dim eventId As Integer = mScript.mainClass(classO).Properties("ObjectMouseLeave").eventId
        If eventId > 0 Then
            wasEvent = True
            res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectMouseLeave", False)
            If res = "#Error" Then Return
        End If

        'данного предмета
        If res <> "False" Then
            eventId = mScript.mainClass(classO).ChildProperties(objId)("ObjectMouseLeave").eventId
            If eventId > 0 Then
                wasEvent = True
                res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectMouseLeave", False)
                If res = "#Error" Then Return
            End If
        End If

        'событие ScriptFinishedEvent
        If wasEvent Then
            If mScript.eventRouter.RunScriptFinishedEvent(Nothing) = "#Error" Then Return
        End If

        'получаем PictureHover. Если пусто, то выход
        Dim objPictureHover As String = ""
        If ReadProperty(classO, "PictureHover", objId, -1, objPictureHover, arrs) = False Then Return
        objPictureHover = UnWrapString(objPictureHover)
        If String.IsNullOrEmpty(objPictureHover) = False Then
            Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, objPictureHover)
            If FileIO.FileSystem.FileExists(fPath) = False Then
                Return
            End If

            'получаем Picture. PictureHover и Picture не должны быть равны
            Dim objPicture As String = ""
            If ReadProperty(classO, "Picture", objId, -1, objPicture, arrs) = False Then Return
            objPicture = UnWrapString(objPicture)

            'убеждаемся в правильном формате пути к картинкам
            objPictureHover = objPictureHover.Replace("\", "/")
            If String.IsNullOrEmpty(objPicture) = False Then objPicture = objPicture.Replace("\", "/")
            If String.Compare(objPicture, objPictureHover, True) <> 0 Then
                'собственно замена картинки Picture на PictureHover
                If objType = ObjectTypeEnum.DESCRIPTION Then
                    'type = описание. Находим первую же картинку и работаем с ней
                    'description - DIV/SPAN-...(IMG)...
                    Dim hEl As HtmlElement = sender
                    Dim hImgCol As HtmlElementCollection = hEl.GetElementsByTagName("IMG")
                    If hImgCol.Count > 0 Then
                        hImgCol(0).SetAttribute("src", objPicture)
                    End If
                Else
                    '0 - обычный, 1 - контейнер
                    'возможное расположение картинки: 
                    '1) IMG 2) SPAN-NOBR-IMG 3) TABLE-TBODY-TR-TD
                    Dim hEl As HtmlElement = sender
                    If hEl.TagName <> "IMG" Then
                        Do
                            If hEl.Children.Count = 0 Then
                                hEl = Nothing
                                Exit Do
                            End If
                            hEl = hEl.Children(0)
                            If hEl.TagName = "IMG" Then Exit Do
                        Loop

                        If IsNothing(hEl) = False Then
                            hEl.SetAttribute("Src", objPicture)
                        End If
                    End If
                End If
            End If
        End If

        'очищаем описание
        If objType <> ObjectTypeEnum.DESCRIPTION Then
            Dim hDoc As HtmlDocument = frmPlayer.wbDescription.Document
            If IsNothing(hDoc) Then Return
            Dim hDesc As HtmlElement = hDoc.GetElementById("DescriptionConvas")
            If IsNothing(hDesc) = False Then hDesc.InnerHtml = ""
        End If
    End Sub

    Friend Sub del_object_Click(sender As Object, e As HtmlElementEventArgs)
        'sender - ObjectInfo
        Dim classO As Integer = mScript.mainClassHash("O")
        Dim objId As Integer = Val(sender.Parent.Id.Substring(6))
        Dim arrs() As String = {objId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classO, "Enabled", objId, -1, enabled, arrs)
        If enabled = False Then Return

        If GVARS.G_ISBATTLE Then
            Dim enabledInBattle As Boolean = False
            ReadProperty(classO, "EnabledInBattle", objId, -1, enabledInBattle, arrs)
            If Not enabled Then GoTo 1
        End If

        'получаем тип. '0 - обычный, 1 - контернер, 2 - описание, 3 - разделитель
        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
        ReadPropertyInt(classO, "Type", objId, -1, objType, arrs)
        If objType = ObjectTypeEnum.SEPARATOR Then Return 'для типа сеператор не подходит

        'событие ObjectSelectEvent
        Dim res As String = ""
        Dim wasEvent As Boolean = False
        'глобальное
        Dim eventId As Integer = mScript.mainClass(classO).Properties("ObjectSelectEvent").eventId
        If eventId > 0 Then
            wasEvent = True
            res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectSelectEvent", False)
            If res = "#Error" Then Return
        End If

        'данного предмета
        If res <> "False" Then
            eventId = mScript.mainClass(classO).ChildProperties(objId)("ObjectSelectEvent").eventId
            If eventId > 0 Then
                wasEvent = True
                res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectSelectEvent", False)
                If res = "#Error" Then Return
            End If
        End If

        'событие ScriptFinishedEvent
        If wasEvent AndAlso EventGeneratedFromScript = False Then
            If mScript.eventRouter.RunScriptFinishedEvent(Nothing) = "#Error" Then Return
        End If
        EventGeneratedFromScript = False

        If res = "False" Then Return

        'отображение меню
        Dim objMnu As String = ""
        If ReadProperty(classO, "Menu", objId, -1, objMnu, arrs) = False Then Return
        If String.IsNullOrEmpty(objMnu) = False AndAlso objMnu <> "''" Then
            Dim classM As Integer = mScript.mainClassHash("M")
            Dim mnuId As Integer = GetSecondChildIdByName(objMnu, mScript.mainClass(classM).ChildProperties)
            If mnuId >= 0 Then
                If ShowMenu(mnuId, arrs, objId) = "#Error" Then Return
            End If
        End If


1:
        'отображение контейнера
        If objType = ObjectTypeEnum.CONTAINER Then
            Dim outerDiv As HtmlElement = sender.Parent
            If outerDiv.Children.Count > 1 Then
                Dim hCont As HtmlElement = outerDiv.Children(1)

                HTMLSwitchClass(hCont, "Collapsed")
            End If
        End If
    End Sub

    Private Sub del_object_MouseDown(sender As Object, e As HtmlElementEventArgs)
        'sender - ObjectInfo
        Dim classO As Integer = mScript.mainClassHash("O")
        Dim objId As Integer = Val(sender.Parent.Id.Substring(6))
        Dim arrs() As String = {objId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classO, "Enabled", objId, -1, enabled, arrs)
        If enabled = False Then Return

        'получаем тип. '0 - обычный, 1 - контернер, 2 - описание, 3 - разделитель
        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
        ReadPropertyInt(classO, "Type", objId, -1, objType, arrs)
        If objType = ObjectTypeEnum.SEPARATOR Then Return 'для типа сеператор не подходит

        If objType = ObjectTypeEnum.CONTAINER Then
            Dim parCont As String = ""
            ReadProperty(classO, "Container", objId, -1, parCont, arrs)
            Dim parId As Integer = GetSecondChildIdByName(parCont, mScript.mainClass(classO).ChildProperties)
            Dim containedStyle As ObjectContainedStyleEnum = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE
            If parId >= 0 Then
                ReadPropertyInt(classO, "ContainedStyle", parId, -1, containedStyle, arrs)
            End If
            If containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
                HTMLAddClass(sender, "ObjectMouseDownForContainers")
            Else
                HTMLAddClass(sender, "ObjectMouseDown")
            End If
        Else
            HTMLAddClass(sender, "ObjectMouseDown")
        End If

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim objPictureActive As String = ""
        If ReadProperty(classO, "PictureActive", objId, -1, objPictureActive, arrs) = False Then Return
        objPictureActive = UnWrapString(objPictureActive)
        If String.IsNullOrEmpty(objPictureActive) = False Then
            Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, objPictureActive)
            If FileIO.FileSystem.FileExists(fPath) = False Then
                MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            'убеждаемся в правильном формате пути к картинкам
            objPictureActive = objPictureActive.Replace("\", "/")

            'собственно замена картинки PictureActive
            If objType = ObjectTypeEnum.DESCRIPTION Then
                'type = описание. Находим первую же картинку и работаем с ней
                'description - DIV/SPAN-...(IMG)...
                Dim hEl As HtmlElement = sender
                Dim hImgCol As HtmlElementCollection = hEl.GetElementsByTagName("IMG")
                If hImgCol.Count > 0 Then
                    hImgCol(0).SetAttribute("src", objPictureActive)
                End If
            Else
                '0 - обычный, 1 - контейнер
                'возможное расположение картинки: 
                '1) IMG 2) SPAN-NOBR-IMG 3) TABLE-TBODY-TR-TD
                Dim hEl As HtmlElement = sender
                If hEl.TagName <> "IMG" Then
                    Do
                        If hEl.Children.Count = 0 Then
                            hEl = Nothing
                            Exit Do
                        End If
                        hEl = hEl.Children(0)
                        If hEl.TagName = "IMG" Then Exit Do
                    Loop

                    If IsNothing(hEl) = False Then
                        hEl.SetAttribute("Src", objPictureActive)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub del_object_MouseUp(sender As Object, e As HtmlElementEventArgs)
        Dim classO As Integer = mScript.mainClassHash("O")
        Dim objId As Integer = Val(sender.Parent.Id.Substring(6))
        Dim arrs() As String = {objId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classO, "Enabled", objId, -1, enabled, arrs)
        If enabled = False Then Return

        'получаем тип. '0 - обычный, 1 - контернер, 2 - описание, 3 - разделитель
        Dim objType As Integer = 0
        ReadPropertyInt(classO, "Type", objId, -1, objType, arrs)
        If objType = 3 Then Return 'для типа сеператор не подходит

        If objType = ObjectTypeEnum.CONTAINER Then
            Dim parCont As String = ""
            ReadProperty(classO, "Container", objId, -1, parCont, arrs)
            Dim parId As Integer = GetSecondChildIdByName(parCont, mScript.mainClass(classO).ChildProperties)
            Dim containedStyle As ObjectContainedStyleEnum = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE
            If parId >= 0 Then
                ReadPropertyInt(classO, "ContainedStyle", parId, -1, containedStyle, arrs)
            End If

            If containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
                HTMLRemoveClass(sender, "ObjectMouseDownForContainers")
            Else
                HTMLRemoveClass(sender, "ObjectMouseDown")
            End If
        Else
            HTMLRemoveClass(sender, "ObjectMouseDown")
        End If


        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim objPictureActive As String = ""
        If ReadProperty(classO, "PictureActive", objId, -1, objPictureActive, arrs) = False Then Return
        objPictureActive = UnWrapString(objPictureActive)
        If String.IsNullOrEmpty(objPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim objPicture As String = ""
        If ReadProperty(classO, "PictureHover", objId, -1, objPicture, arrs) = False Then Return
        objPicture = UnWrapString(objPicture)
        If String.IsNullOrEmpty(objPicture) Then
            If ReadProperty(classO, "Picture", objId, -1, objPicture, arrs) = False Then Return
            objPicture = UnWrapString(objPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, objPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        objPicture = objPicture.Replace("\", "/")

        'собственно замена картинки PictureActive
        If objType = ObjectTypeEnum.DESCRIPTION Then
            'type = описание. Находим первую же картинку и работаем с ней
            'description - DIV/SPAN-...(IMG)...
            Dim hEl As HtmlElement = sender
            Dim hImgCol As HtmlElementCollection = hEl.GetElementsByTagName("IMG")
            If hImgCol.Count > 0 Then
                hImgCol(0).SetAttribute("src", objPictureActive)
            End If
        Else
            '0 - обычный, 1 - контейнер
            'возможное расположение картинки: 
            '1) IMG 2) SPAN-NOBR-IMG 3) TABLE-TBODY-TR-TD
            Dim hEl As HtmlElement = sender
            If hEl.TagName <> "IMG" Then
                Do
                    If hEl.Children.Count = 0 Then
                        hEl = Nothing
                        Exit Do
                    End If
                    hEl = hEl.Children(0)
                    If hEl.TagName = "IMG" Then Exit Do
                Loop

                If IsNothing(hEl) = False Then
                    hEl.SetAttribute("Src", objPictureActive)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Подготавливает внешний вид окна действий
    ''' </summary>
    ''' <param name="arrParams"></param>
    Public Sub ObjectWindowPrepareFace(ByRef arrParams() As String)
        Dim hDoc As HtmlDocument = frmPlayer.wbObjects.Document
        If IsNothing(hDoc) Then Return
        Dim hConvas As HtmlElement = hDoc.GetElementById("ObjectsConvas")
        If IsNothing(hConvas) Then Return
        Dim classOW As Integer = mScript.mainClassHash("OW")

        'Меняем фон
        Dim strStyle As New System.Text.StringBuilder
        Dim bkColor As String = ""
        ReadProperty(classOW, "BackColor", -1, -1, bkColor, arrParams)
        bkColor = UnWrapString(bkColor)
        If String.IsNullOrEmpty(bkColor) = False Then
            strStyle.Append("background-color:" & bkColor & ";")
        End If

        'Меняем фоновую картинку
        Dim bkPicture As String = ""
        ReadProperty(classOW, "BackPicture", -1, -1, bkPicture, arrParams)
        bkPicture = UnWrapString(bkPicture)
        Dim bkPicPos As Integer = 0, bkPicStyle As Integer = 0
        If String.IsNullOrEmpty(bkPicture) = False Then
            bkPicture = bkPicture.Replace("\"c, "/"c)
            ReadPropertyInt(classOW, "BackPicPos", -1, -1, bkPicPos, arrParams)
            ReadPropertyInt(classOW, "BackPicStyle", -1, -1, bkPicStyle, arrParams)
        End If

        If String.IsNullOrEmpty(bkPicture) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, bkPicture)) Then
            'файл картинки существует
            strStyle.AppendFormat("background-image: url({0});", "'" + bkPicture + "'")
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

        hConvas.Document.Body.Style = strStyle.ToString
        strStyle.Clear()
    End Sub

#End Region

#Region "Menu"
    ''' <summary>
    ''' Отображает указанный блок меню
    ''' </summary>
    ''' <param name="mId">Id блока меню</param>
    ''' <param name="objId">Id премета, вызвавшего блок меню</param>
    ''' <param name="arrParams"></param>
    Public Function ShowMenu(ByVal mId As Integer, ByRef arrParams() As String, Optional ByVal objId As Integer = -1) As String
        Dim classId As Integer = mScript.mainClassHash("M")
        Dim hDoc As HtmlDocument, hConvas As HtmlElement
        If IsNothing(mScript.mainClass(classId).ChildProperties(mId)("Name").ThirdLevelProperties) OrElse _
            mScript.mainClass(classId).ChildProperties(mId)("Name").ThirdLevelProperties.Count = 0 Then Return "" 'меню пустое

        'получаем контейнер для размещения меню
        Dim strContainer As String = ""
        If ReadProperty(classId, "HTMLContainerId", mId, -1, strContainer, arrParams) = False Then Return "#Error"
        strContainer = UnWrapString(strContainer)
        If String.IsNullOrEmpty(strContainer) Then
            'меню в обычном месте
            hDoc = frmPlayer.wbDescription.Document
            If IsNothing(hDoc) Then Return _ERROR("Не удалось получить html-документ окна описаний.")
            hConvas = hDoc.GetElementById("DescriptionConvas")
            If IsNothing(hConvas) = False Then hConvas.Style = "display:none"
            hConvas = hDoc.GetElementById("MenuConvas")
            If IsNothing(hConvas) Then _ERROR("Нарушеня структура документа description.html. Не найден элемент MenuConvas.")
            hConvas.InnerHtml = ""
        Else
            'меню в главном окне
            hDoc = frmPlayer.wbMain.Document
            If IsNothing(hDoc) Then Return _ERROR("Не удалось получить html-документ окна локаций.")
            hConvas = hDoc.GetElementById(strContainer)
            If IsNothing(hConvas) Then _ERROR("Html-контейнер " & strContainer & " меню не найден!")
            hConvas.InnerHtml = ""
        End If
        GVARS.G_CURMENU = -1

        'События MenuSelectEvent глобальное и блока меню
        Dim res As String = "", arrs() As String = {mId.ToString, objId.ToString}
        Dim eventId As Integer = mScript.mainClass(classId).Properties("MenuSelectEvent").eventId
        If eventId > 0 Then
            'глобальное
            res = mScript.eventRouter.RunEvent(eventId, arrs, "MenuSelectEvent", False)
            If res = "#Error" Then Return res
        End If

        If res <> "False" Then
            eventId = mScript.mainClass(classId).ChildProperties(mId)("MenuSelectEvent").eventId
            If eventId > 0 Then
                res = mScript.eventRouter.RunEvent(eventId, arrs, "MenuSelectEvent", False)
                If res = "#Error" Then Return res
            End If
        End If

        If res = "False" Then Return "" 'событие отменило отображение блока

        'показываем подпись
        Dim mLabel As String = ""
        If ReadProperty(classId, "Caption", mId, -1, mLabel, arrs) = False Then Return "#Error"
        mLabel = UnWrapString(mLabel)
        Dim hTitle As HtmlElement = hDoc.CreateElement("DIV")
        hTitle.SetAttribute("ClassName", "menuTitle")
        hTitle.InnerHtml = mLabel
        hConvas.AppendChild(hTitle)
        If String.IsNullOrEmpty(mLabel) Then
            hTitle.Style = "display:none"
        End If

        'показываем меню
        Dim lstCh As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classId).ChildProperties(mId)
        For itmId As Integer = 0 To lstCh("Name").ThirdLevelProperties.Count - 1
            MenuCreateHtmlElement(mId, itmId, objId, hConvas)
        Next itmId

        'показывает кнопку отмены
        Dim showCancel As Integer = 0
        If ReadPropertyInt(classId, "AutoCancelButton", mId, -1, showCancel, arrParams) = False Then Return "#Error"
        If showCancel <> 0 Then
            MenuCreateHtmlCancelElement(mId, showCancel = 2, hConvas, arrParams, objId)
        End If
        GVARS.G_CURMENU = mId

        Return ""
    End Function

    ''' <summary>
    ''' Создает html-элемент пункта меню
    ''' </summary>
    ''' <param name="mId">Id блока меню</param>
    ''' <param name="itemId">Id пункта меню</param>
    ''' <param name="objId">Id предмета, вызвавшего блок меню</param>
    ''' <param name="htmlContainer">html-контейнер блока меню</param>
    Public Function MenuCreateHtmlElement(ByVal mId As Integer, ByVal itemId As Integer, ByVal objId As Integer, ByRef htmlContainer As HtmlElement, _
                                          Optional ByRef elementToPlace As HtmlElement = Nothing) As HtmlElement
        Dim menuArrs() As String = {mId.ToString, itemId.ToString, objId.ToString}
        Dim classId As Integer = mScript.mainClassHash("M")
        Dim eventId As Integer, result As String
        'получаем Visible
        Dim itmVisible As Boolean = True '150-170
        ReadPropertyBool(classId, "Visible", mId, itemId, itmVisible, menuArrs)
        'Выполняем результат M.ShowQuery текущего пункта
        Dim showQ As Boolean = True
        If itmVisible Then
            eventId = mScript.mainClass(classId).ChildProperties(mId)("ShowQuery").ThirdLevelEventId(itemId)
            If eventId > 0 Then
                result = mScript.eventRouter.RunEvent(eventId, menuArrs, "Menu.ShowQuery", False)
                Boolean.TryParse(result, showQ)
            End If
        End If
        'получаем надо ли фактически отображать пункт
        Dim isShown As Boolean = (itmVisible = True AndAlso showQ = True)
        'получаем Caption
        Dim itemCaption As String = ""
        ReadProperty(classId, "Caption", mId, itemId, itemCaption, menuArrs)
        itemCaption = UnWrapString(itemCaption)
        'получаем Enabled
        Dim itemEnabled As Boolean = True
        ReadPropertyBool(classId, "Enabled", mId, itemId, itemEnabled, menuArrs)
        'получаем Picture
        Dim itemPicture As String = ""
        ReadProperty(classId, "Picture", mId, itemId, itemPicture, menuArrs)
        itemPicture = UnWrapString(itemPicture).Replace("\", "/")
        'получаем PictureFloat
        Dim itemPictureFloat As Integer = 0 '0 - без прилегания, 1 - лево, 2 - право
        If String.IsNullOrEmpty(itemPicture) = False Then ReadPropertyInt(classId, "PictureFloat", mId, itemId, itemPictureFloat, menuArrs)
        'получаем имя
        Dim itemName As String = mScript.mainClass(classId).ChildProperties(mId)("Name").ThirdLevelProperties(itemId)
        'получаем тип свойства
        '0 - кнопка, 1 - изображение, 2 - ссылка, 3 - разделитель, 4 - текст
        Dim itemType As Integer = 0
        ReadPropertyInt(classId, "Type", mId, itemId, itemType, menuArrs)
        Dim hEl As HtmlElement
        Select Case itemType
            Case 1 'img
                hEl = htmlContainer.Document.CreateElement("IMG")
                hEl.SetAttribute("Src", itemPicture)
                hEl.SetAttribute("Title", itemCaption)
                hEl.SetAttribute("ClassName", "MenuItemImage")
            Case 2 'anchor
                hEl = htmlContainer.Document.CreateElement("A")
                hEl.SetAttribute("ClassName", "MenuItemAnchor")
                hEl.SetAttribute("href", "#")
            Case 3 'hr
                hEl = htmlContainer.Document.CreateElement("HR")
                hEl.SetAttribute("ClassName", "MenuItemHR")
            Case 4 'text
                hEl = htmlContainer.Document.CreateElement("DIV")
                hEl.SetAttribute("ClassName", "MenuItemText")
            Case Else '0 - button
                hEl = htmlContainer.Document.CreateElement("DIV")
                hEl.SetAttribute("ClassName", "MenuItemButton")
        End Select

        If Not isShown Then
            'действие скрыто
            hEl.Style = "display:none"
        End If

        'enabled
        If Not itemEnabled Then
            hEl.SetAttribute("disabled", "True")
        End If

        'получаем текст действия
        MenuSetInnerHTML(hEl, itemType, itemCaption, itemPicture, itemPictureFloat)
        hEl.Name = itemName
        hEl.SetAttribute("menu", mId.ToString)
        hEl.SetAttribute("objId", objId.ToString)

        'создаем событие клика
        If itemType = 0 Then 'Button
            AddHandler hEl.MouseDown, AddressOf del_menu_MouseDown_Button
            AddHandler hEl.MouseUp, AddressOf del_menu_MouseUp_Button
        ElseIf itemType = 1 Then 'Image
            AddHandler hEl.MouseDown, AddressOf del_menu_MouseDown_Image
            AddHandler hEl.MouseUp, AddressOf del_menu_MouseUp_Image
        ElseIf itemType = 2 Then 'Anchor
            AddHandler hEl.MouseDown, AddressOf del_menu_MouseDown_Anchor
            AddHandler hEl.MouseUp, AddressOf del_menu_MouseUp_Anchor
        End If

        If itemType <= 2 Then 'Button, Image, Anchor
            AddHandler hEl.Click, AddressOf del_menu_Click
            AddHandler hEl.MouseOver, AddressOf del_menu_MouseOver
            AddHandler hEl.MouseLeave, AddressOf del_menu_MouseLeave
        End If

        'добавляем пункт меню
        If IsNothing(elementToPlace) Then
            htmlContainer.AppendChild(hEl)
        Else
            elementToPlace.InsertAdjacentElement(HtmlElementInsertionOrientation.BeforeBegin, hEl)
        End If
        Return hEl
    End Function

    ''' <summary>
    ''' Создает html-элемент пункта меню
    ''' </summary>
    ''' <param name="mId">Id блока меню</param>
    ''' <param name="showSeparator">Предварять ли пункт отмены сепаратором</param>
    ''' <param name="htmlContainer">html-контейнер блока меню</param>
    Public Function MenuCreateHtmlCancelElement(ByVal mId As Integer, ByVal showSeparator As Boolean, ByRef htmlContainer As HtmlElement, ByRef arrParams() As String, ByVal objId As Integer) As HtmlElement
        Dim classId As Integer = mScript.mainClassHash("M")
        'получаем Caption
        Dim itemCaption As String = ""
        ReadProperty(classId, "CancelButtonText", mId, -1, itemCaption, arrParams)
        itemCaption = UnWrapString(itemCaption)
        'получаем Picture
        Dim itemPicture As String = ""
        ReadProperty(classId, "CancelButtonPicture", mId, -1, itemPicture, arrParams)
        itemPicture = UnWrapString(itemPicture).Replace("\", "/")
        'получаем PictureFloat
        Dim itemPictureFloat As Integer = 0 '0 - без прилегания, 1 - лево, 2 - право
        If String.IsNullOrEmpty(itemPicture) = False Then ReadPropertyInt(classId, "CancelButtonPictureFloat", mId, -1, itemPictureFloat, arrParams)

        'получаем тип свойства
        '0 - кнопка, 1 - изображение, 2 - ссылка, 3 - разделитель, 4 - текст
        Dim itemType As Integer = 0
        ReadPropertyInt(classId, "CancelButtonType", mId, -1, itemType, arrParams)
        Dim hEl As HtmlElement

        If showSeparator Then
            'выводим разделитель
            hEl = htmlContainer.Document.CreateElement("HR")
            hEl.SetAttribute("ClassName", "MenuItemHR")
            htmlContainer.AppendChild(hEl)
        End If

        Select Case itemType
            Case 1 'img
                hEl = htmlContainer.Document.CreateElement("IMG")
                hEl.SetAttribute("Src", itemPicture)
                hEl.SetAttribute("Title", itemCaption)
                hEl.SetAttribute("ClassName", "MenuCancelImage")
            Case 2 'anchor
                hEl = htmlContainer.Document.CreateElement("A")
                hEl.SetAttribute("ClassName", "MenuCancelAnchor")
                hEl.SetAttribute("href", "#")
            Case Else '0 - button
                hEl = htmlContainer.Document.CreateElement("DIV")
                hEl.SetAttribute("ClassName", "MenuCancelButton")
        End Select

        'получаем текст действия
        MenuSetInnerHTML(hEl, itemType, itemCaption, itemPicture, itemPictureFloat)
        hEl.Name = "sys_AutoCancelButton"
        hEl.SetAttribute("menu", mId.ToString)
        hEl.SetAttribute("objId", objId.ToString)

        'создаем событие клика
        If itemType = 0 Then 'Button
            AddHandler hEl.MouseDown, AddressOf del_menu_cancel_MouseDown_Button
            AddHandler hEl.MouseUp, AddressOf del_menu_cancel_MouseUp_Button
        ElseIf itemType = 1 Then 'Image
            AddHandler hEl.MouseDown, AddressOf del_menu_cancel_MouseDown_Image
            AddHandler hEl.MouseUp, AddressOf del_menu_cancel_MouseUp_Image
        ElseIf itemType = 2 Then 'Anchor
            AddHandler hEl.MouseDown, AddressOf del_menu_cancel_MouseDown_Anchor
            AddHandler hEl.MouseUp, AddressOf del_menu_cancel_MouseUp_Anchor
        End If

        If itemType <= 2 Then 'Button, Image, Anchor
            AddHandler hEl.Click, AddressOf del_menu_cancel_Click
            AddHandler hEl.MouseOver, AddressOf del_menu_cancel_MouseOver
            AddHandler hEl.MouseLeave, AddressOf del_menu_cancel_MouseLeave
        End If

        'добавляем действие
        htmlContainer.AppendChild(hEl)
        Return hEl
    End Function


    ''' <summary>
    ''' Создает надпись для отображения пунктов меню типа button, anchor, text
    ''' </summary>
    ''' <param name="hItem">ссылка на html-элемент пункта меню</param>
    ''' <param name="mType">тип пункта меню</param>
    ''' <param name="itemCaption">подпись без кавычек</param>
    ''' <param name="itemPicture">путь к картинке относительно папки квеста, без кавычек и в правильном формате</param>
    ''' <param name="itemPictureFloat">прилегание картинки</param>
    Public Sub MenuSetInnerHTML(ByRef hItem As HtmlElement, ByVal mType As Integer, ByVal itemCaption As String, ByVal itemPicture As String, ByVal itemPictureFloat As Integer)
        'получаем текст пункта меню
        Dim hInner As New System.Text.StringBuilder
        If mType <> 3 AndAlso mType <> 1 Then
            'вставляем картинку если тип button, anchor или text
            If String.IsNullOrEmpty(itemPicture) = False Then
                hInner.Append("<img src='" & itemPicture)
                If itemPictureFloat = 1 Then
                    hInner.Append("' style='float:left'/>")
                ElseIf itemPictureFloat = 2 Then
                    hInner.Append("' style='float:right'/>")
                Else
                    hInner.Append("'/>")
                End If
            End If
            If String.IsNullOrEmpty(itemCaption) = False Then itemCaption = "<span>" & itemCaption & "</span>"
            hInner.Append(itemCaption) 'вставляем сам текст
            hItem.InnerHtml = hInner.ToString
            hInner.Clear()
        End If
    End Sub

    ''' <summary>
    ''' Прячет указанный блок меню
    ''' </summary>
    ''' <param name="mId">Id блока меню</param>
    ''' <param name="arrParams"></param>
    Public Function HideMenu(ByVal mId As Integer, ByRef arrParams() As String) As String
        If mId = -1 Then mId = GVARS.G_CURMENU
        If mId = -1 Then Return ""

        Dim classId As Integer = mScript.mainClassHash("M")
        Dim hDoc As HtmlDocument, hConvas As HtmlElement

        Dim strContainer As String = ""
        If ReadProperty(classId, "HTMLContainerId", mId, -1, strContainer, arrParams) = False Then Return "#Error"
        strContainer = UnWrapString(strContainer)
        If String.IsNullOrEmpty(strContainer) Then
            'меню в обычном месте
            hDoc = frmPlayer.wbDescription.Document
            If IsNothing(hDoc) Then Return _ERROR("Не удалось получить html-документ окна описаний.")
            hConvas = hDoc.GetElementById("MenuConvas")
            If IsNothing(hConvas) = False Then hConvas.InnerHtml = ""
            hConvas = hDoc.GetElementById("DescriptionConvas")
            If IsNothing(hConvas) = False Then hConvas.Style = ""
        Else
            'меню в главном окне
            hDoc = frmPlayer.wbMain.Document
            If IsNothing(hDoc) Then Return _ERROR("Не удалось получить html-документ окна локаций.")
            hConvas = hDoc.GetElementById(strContainer)
            If IsNothing(hConvas) Then _ERROR("Html-контейнер " & strContainer & " меню не найден!")
            hConvas.InnerHtml = ""
        End If
        Return ""
    End Function


    Private Sub del_menu_MouseOver(sender As HtmlElement, e As HtmlElementEventArgs)
        'Выводим картинку действия при наведении PictureHover
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim itemId As Integer = GetThirdChildIdByName(sender.Name, mId, mScript.mainClass(classM).ChildProperties)
        Dim arrs() As String = {mId.ToString, itemId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classM, "Enabled", mId, itemId, enabled, arrs)
        If enabled = False Then Return

        'получаем PictureHover. Если пусто, то выход
        Dim itemPictureHover As String = ""
        If ReadProperty(classM, "PictureHover", mId, itemId, itemPictureHover, arrs) = False Then Return
        itemPictureHover = UnWrapString(itemPictureHover)
        If String.IsNullOrEmpty(itemPictureHover) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureHover)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'получаем Picture. PictureHover и Picture не должны быть равны
        Dim itemPicture As String = ""
        If ReadProperty(classM, "Picture", mId, itemId, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)

        'получаем тип. Подходят только типы кнопка, картинка и ссылка
        Dim itemType As Integer = 0
        ReadPropertyInt(classM, "Type", mId, itemId, itemType, arrs)
        If itemType > 2 Then Return 'для типов сеператор и простой текст не актуально

        'убеждаемся в правильном формате пути к картинкам
        itemPictureHover = itemPictureHover.Replace("\", "/")
        If String.IsNullOrEmpty(itemPicture) = False Then itemPicture = itemPicture.Replace("\", "/")
        If String.Compare(itemPicture, itemPictureHover, True) = 0 Then Return

        'собственно замена картинки Picture на PictureHover
        If itemType = 1 Then
            'type = картинка
            sender.SetAttribute("src", itemPictureHover)
        ElseIf sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'type = кнопка или ссылка (есть IMG)
            sender.Children(0).SetAttribute("src", itemPictureHover)
        Else
            'type = кнопка или ссылка (нет IMG)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "PictureFloat", mId, itemId, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "Caption", mId, itemId, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPictureHover, itemPictureFloat)
        End If
    End Sub

    Private Sub del_menu_MouseLeave(sender As HtmlElement, e As HtmlElementEventArgs)
        'возвращаем картинку пункта меню Picture после того, как она была зменена на PictureHover
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim itemId As Integer = GetThirdChildIdByName(sender.Name, mId, mScript.mainClass(classM).ChildProperties)
        Dim arrs() As String = {mId.ToString, itemId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classM, "Enabled", mId, itemId, enabled, arrs)
        If enabled = False Then Return

        'получаем PictureHover. Если пусто, то выход
        Dim itemPictureHover As String = ""
        If ReadProperty(classM, "PictureHover", mId, itemId, itemPictureHover, arrs) = False Then Return
        itemPictureHover = UnWrapString(itemPictureHover)
        If String.IsNullOrEmpty(itemPictureHover) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureHover)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            'MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'получаем Picture. PictureHover и Picture не должны быть равны
        Dim itemPicture As String = ""
        If ReadProperty(classM, "Picture", mId, itemId, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)

        'получаем тип. Подходят только типы кнопка, картинка и ссылка
        Dim itemType As Integer = 0
        ReadPropertyInt(classM, "Type", mId, itemId, itemType, arrs)
        If itemType > 2 Then Return 'для типов сеператор и простой текст не актуально

        'убеждаемся в правильном формате пути к картинкам
        itemPictureHover = itemPictureHover.Replace("\", "/")
        If String.IsNullOrEmpty(itemPicture) = False Then itemPicture = itemPicture.Replace("\", "/")
        If String.Compare(itemPicture, itemPictureHover, True) = 0 Then Return

        'собственно замена картинки PictureHover на Picture
        If itemType = 1 Then
            'type = картинка
            sender.SetAttribute("src", itemPicture)
        ElseIf String.IsNullOrEmpty(itemPicture) = True AndAlso sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'type = кнопка или ссылка (есть IMG)
            sender.Children(0).SetAttribute("src", itemPicture)
        Else
            'type = кнопка или ссылка (нет IMG)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "PictureFloat", mId, itemId, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "Caption", mId, itemId, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPicture, itemPictureFloat)
        End If

    End Sub

    Friend Sub del_menu_Click(sender As HtmlElement, e As HtmlElementEventArgs)
        'событие выбора пункта меню
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim itemId As Integer = GetThirdChildIdByName(sender.Name, mId, mScript.mainClass(classM).ChildProperties)
        Dim arrs() As String = {mId.ToString, itemId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classM, "Enabled", mId, itemId, enabled, arrs)
        If enabled = False Then Return

        'Visits + 1
        Dim visits As Integer = 0
        ReadPropertyInt(classM, "Visits", mId, itemId, visits, arrs)
        visits += 1
        Dim visitsRes As String = visits.ToString
        Dim trackingResult As String = mScript.trackingProperties.RunBefore(classM, "Visits", arrs, visitsRes)
        If trackingResult <> "False" AndAlso trackingResult <> "#Error" Then
            mScript.mainClass(classM).ChildProperties(mId)("Visits").ThirdLevelProperties(itemId) = visitsRes
        End If

        'Прячем / делаем недоступным если надо
        Dim afterVisiting As Integer = 0
        ReadPropertyInt(classM, "AfterVisiting", mId, itemId, afterVisiting, arrs)
        If afterVisiting = 1 Then
            'скрыть
            PropertiesRouter(classM, "Visible", arrs, arrs, PropertiesOperationEnum.PROPERTY_SET, "False")
        ElseIf afterVisiting = 2 Then
            'сделать недоступным
            PropertiesRouter(classM, "Enabled", arrs, arrs, PropertiesOperationEnum.PROPERTY_SET, "False")
        End If

        'событие выбора пункта меню
        Dim objId As Integer = Val(sender.GetAttribute("objId"))
        ReDim Preserve arrs(2)
        arrs(2) = objId
        Dim eventId As Integer = mScript.mainClass(classM).ChildProperties(mId)("MenuSelectEvent").ThirdLevelEventId(itemId)
        If eventId > 0 Then
            mScript.eventRouter.RunEvent(eventId, arrs, "MenuSelectEvent", False)
            If mScript.LAST_ERROR.Length > 0 Then Return
            If EventGeneratedFromScript = False Then mScript.eventRouter.RunScriptFinishedEvent(arrs)
        End If
        EventGeneratedFromScript = False

        Dim res As String = PropertiesRouter(classM, "Visible", {mId.ToString}, arrs, PropertiesOperationEnum.PROPERTY_SET, "False")
        If res <> "False" AndAlso res <> "#Error" AndAlso mId = GVARS.G_CURMENU Then
            GVARS.G_CURMENU = -1
        End If

    End Sub

    Private Sub del_menu_MouseDown_Image(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim itemId As Integer = GetThirdChildIdByName(sender.Name, mId, mScript.mainClass(classM).ChildProperties)
        Dim arrs() As String = {mId.ToString, itemId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classM, "Enabled", mId, itemId, enabled, arrs)
        If enabled = False Then Return

        HTMLReplaceClass(sender, "MenuItemImage", "MenuItemImageMouseDown")

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "PictureActive", mId, itemId, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureActive)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPictureActive = itemPictureActive.Replace("\", "/")

        'собственно замена картинки PictureActive
        sender.SetAttribute("src", itemPictureActive)
    End Sub

    Private Sub del_menu_MouseUp_Image(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim itemId As Integer = GetThirdChildIdByName(sender.Name, mId, mScript.mainClass(classM).ChildProperties)
        Dim arrs() As String = {mId.ToString, itemId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classM, "Enabled", mId, itemId, enabled, arrs)
        If enabled = False Then Return

        HTMLReplaceClass(sender, "MenuItemImageMouseDown", "MenuItemImage")
        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "PictureActive", mId, itemId, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim itemPicture As String = ""
        If ReadProperty(classM, "PictureHover", mId, itemId, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)
        If String.IsNullOrEmpty(itemPicture) Then
            If ReadProperty(classM, "Picture", mId, itemId, itemPicture, arrs) = False Then Return
            itemPicture = UnWrapString(itemPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPicture = itemPicture.Replace("\", "/")

        'собственно замена картинки PictureActive
        sender.SetAttribute("src", itemPicture)
    End Sub

    Private Sub del_menu_MouseDown_Anchor(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim itemId As Integer = GetThirdChildIdByName(sender.Name, mId, mScript.mainClass(classM).ChildProperties)
        Dim arrs() As String = {mId.ToString, itemId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classM, "Enabled", mId, itemId, enabled, arrs)
        If enabled = False Then Return

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "PictureActive", mId, itemId, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureActive)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPictureActive = itemPictureActive.Replace("\", "/")

        'собственно замена картинки PictureActive
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", itemPictureActive)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim itemType As Integer = 0
            ReadPropertyInt(classM, "Type", mId, itemId, itemType, arrs)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "PictureFloat", mId, itemId, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "Caption", mId, itemId, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPictureActive, itemPictureFloat)
        End If

    End Sub

    Private Sub del_menu_MouseUp_Anchor(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim itemId As Integer = GetThirdChildIdByName(sender.Name, mId, mScript.mainClass(classM).ChildProperties)
        Dim arrs() As String = {mId.ToString, itemId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classM, "Enabled", mId, itemId, enabled, arrs)
        If enabled = False Then Return

        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "PictureActive", mId, itemId, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim itemPicture As String = ""
        If ReadProperty(classM, "PictureHover", mId, itemId, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)
        If String.IsNullOrEmpty(itemPicture) Then
            If ReadProperty(classM, "Picture", mId, itemId, itemPicture, arrs) = False Then Return
            itemPicture = UnWrapString(itemPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPicture = itemPicture.Replace("\", "/")

        'собственно замена картинки PictureActive на PictureHover/Picture
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", itemPicture)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim itemType As Integer = 0
            ReadPropertyInt(classM, "Type", mId, itemId, itemType, arrs)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "PictureFloat", mId, itemId, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "Caption", mId, itemId, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPicture, itemPictureFloat)
        End If
    End Sub

    Private Sub del_menu_MouseDown_Button(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim itemId As Integer = GetThirdChildIdByName(sender.Name, mId, mScript.mainClass(classM).ChildProperties)
        Dim arrs() As String = {mId.ToString, itemId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classM, "Enabled", mId, itemId, enabled, arrs)
        If enabled = False Then Return

        HTMLReplaceClass(sender, "MenuItemButton", "MenuItemButtonMouseDown")

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "PictureActive", mId, itemId, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureActive)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPictureActive = itemPictureActive.Replace("\", "/")

        'собственно замена картинки PictureActive
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", itemPictureActive)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim itemType As Integer = 0
            ReadPropertyInt(classM, "Type", mId, itemId, itemType, arrs)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "PictureFloat", mId, itemId, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "Caption", mId, itemId, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPictureActive, itemPictureFloat)
        End If

    End Sub

    Private Sub del_menu_MouseUp_Button(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim itemId As Integer = GetThirdChildIdByName(sender.Name, mId, mScript.mainClass(classM).ChildProperties)
        Dim arrs() As String = {mId.ToString, itemId.ToString}
        Dim enabled As Boolean = True
        ReadPropertyBool(classM, "Enabled", mId, itemId, enabled, arrs)
        If enabled = False Then Return

        HTMLReplaceClass(sender, "MenuItemButtonMouseDown", "MenuItemButton")

        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "PictureActive", mId, itemId, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim itemPicture As String = ""
        If ReadProperty(classM, "PictureHover", mId, itemId, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)
        If String.IsNullOrEmpty(itemPicture) Then
            If ReadProperty(classM, "Picture", mId, itemId, itemPicture, arrs) = False Then Return
            itemPicture = UnWrapString(itemPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPicture = itemPicture.Replace("\", "/")

        'собственно замена картинки PictureActive на PictureHover/Picture
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", itemPicture)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim itemType As Integer = 0
            ReadPropertyInt(classM, "Type", mId, itemId, itemType, arrs)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "PictureFloat", mId, itemId, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "Caption", mId, itemId, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPicture, itemPictureFloat)
        End If
    End Sub

    Private Sub del_menu_cancel_MouseOver(sender As HtmlElement, e As HtmlElementEventArgs)
        'Выводим картинку действия при наведении PictureHover
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim arrs() As String = {mId.ToString}

        'получаем PictureHover. Если пусто, то выход
        Dim itemPictureHover As String = ""
        If ReadProperty(classM, "CancelButtonPictureHover", mId, -1, itemPictureHover, arrs) = False Then Return
        itemPictureHover = UnWrapString(itemPictureHover)
        If String.IsNullOrEmpty(itemPictureHover) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureHover)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'получаем Picture. PictureHover и Picture не должны быть равны
        Dim itemPicture As String = ""
        If ReadProperty(classM, "CancelButtonPicture", mId, -1, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)

        'получаем тип. Подходят только типы кнопка, картинка и ссылка
        Dim itemType As Integer = 0
        ReadPropertyInt(classM, "CancelButtonType", mId, -1, itemType, arrs)
        If itemType > 2 Then Return 'для типов сеператор и простой текст не актуально

        'убеждаемся в правильном формате пути к картинкам
        itemPictureHover = itemPictureHover.Replace("\", "/")
        If String.IsNullOrEmpty(itemPicture) = False Then itemPicture = itemPicture.Replace("\", "/")
        If String.Compare(itemPicture, itemPictureHover, True) = 0 Then Return

        'собственно замена картинки Picture на PictureHover
        If itemType = 1 Then
            'type = картинка
            sender.SetAttribute("src", itemPictureHover)
        ElseIf sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'type = кнопка или ссылка (есть IMG)
            sender.Children(0).SetAttribute("src", itemPictureHover)
        Else
            'type = кнопка или ссылка (нет IMG)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "CancelButtonPictureFloat", mId, -1, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "CancelButtonText", mId, -1, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPictureHover, itemPictureFloat)
        End If
    End Sub

    Private Sub del_menu_cancel_MouseLeave(sender As HtmlElement, e As HtmlElementEventArgs)
        'возвращаем картинку пункта меню Picture после того, как она была зменена на PictureHover
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim arrs() As String = {mId.ToString}

        'получаем PictureHover. Если пусто, то выход
        Dim itemPictureHover As String = ""
        If ReadProperty(classM, "CancelButtonPictureHover", mId, -1, itemPictureHover, arrs) = False Then Return
        itemPictureHover = UnWrapString(itemPictureHover)
        If String.IsNullOrEmpty(itemPictureHover) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureHover)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            'MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'получаем Picture. PictureHover и Picture не должны быть равны
        Dim itemPicture As String = ""
        If ReadProperty(classM, "CancelButtonPicture", mId, -1, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)

        'получаем тип. Подходят только типы кнопка, картинка и ссылка
        Dim itemType As Integer = 0
        ReadPropertyInt(classM, "CancelButtonType", mId, -1, itemType, arrs)
        If itemType > 2 Then Return 'для типов сеператор и простой текст не актуально

        'убеждаемся в правильном формате пути к картинкам
        itemPictureHover = itemPictureHover.Replace("\", "/")
        If String.IsNullOrEmpty(itemPicture) = False Then itemPicture = itemPicture.Replace("\", "/")
        If String.Compare(itemPicture, itemPictureHover, True) = 0 Then Return

        'собственно замена картинки PictureHover на Picture
        If itemType = 1 Then
            'type = картинка
            sender.SetAttribute("src", itemPicture)
        ElseIf String.IsNullOrEmpty(itemPicture) = True AndAlso sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'type = кнопка или ссылка (есть IMG)
            sender.Children(0).SetAttribute("src", itemPicture)
        Else
            'type = кнопка или ссылка (нет IMG)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "CancelButtonPictureFloat", mId, -1, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "CancelButtonText", mId, -1, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPicture, itemPictureFloat)
        End If

    End Sub

    Private Sub del_menu_cancel_MouseDown_Button(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim arrs() As String = {mId.ToString}

        HTMLReplaceClass(sender, "MenuCancelButton", "MenuCancelButtonMouseDown")

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "CancelButtonPictureActive", mId, -1, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureActive)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPictureActive = itemPictureActive.Replace("\", "/")

        'собственно замена картинки PictureActive
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", itemPictureActive)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim itemType As Integer = 0
            ReadPropertyInt(classM, "CancelButtonType", mId, -1, itemType, arrs)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "CancelButtonPictureFloat", mId, -1, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "CancelButtonText", mId, -1, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPictureActive, itemPictureFloat)
        End If

    End Sub

    Private Sub del_menu_cancel_MouseUp_Button(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim arrs() As String = {mId.ToString}

        HTMLReplaceClass(sender, "MenuCancelButtonMouseDown", "MenuCancelButton")

        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "CancelButtonPictureActive", mId, -1, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim itemPicture As String = ""
        If ReadProperty(classM, "CancelButtonPictureHover", mId, -1, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)
        If String.IsNullOrEmpty(itemPicture) Then
            If ReadProperty(classM, "CancelButtonPicture", mId, -1, itemPicture, arrs) = False Then Return
            itemPicture = UnWrapString(itemPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPicture = itemPicture.Replace("\", "/")

        'собственно замена картинки PictureActive на PictureHover/Picture
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", itemPicture)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim itemType As Integer = 0
            ReadPropertyInt(classM, "CancelButtonType", mId, -1, itemType, arrs)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "CancelButtonPictureFloat", mId, -1, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "CancelButtonText", mId, -1, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPicture, itemPictureFloat)
        End If
    End Sub

    Private Sub del_menu_cancel_MouseDown_Image(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim arrs() As String = {mId.ToString}

        HTMLReplaceClass(sender, "MenuCancelImage", "MenuCancelImageMouseDown")

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "CancelButtonPictureActive", mId, -1, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureActive)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPictureActive = itemPictureActive.Replace("\", "/")

        'собственно замена картинки PictureActive
        sender.SetAttribute("src", itemPictureActive)
    End Sub

    Private Sub del_menu_cancel_MouseUp_Image(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim arrs() As String = {mId.ToString}

        HTMLReplaceClass(sender, "MenuCancelImageMouseDown", "MenuCancelImage")
        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "CancelButtonPictureActive", mId, -1, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim itemPicture As String = ""
        If ReadProperty(classM, "CancelButtonPictureHover", mId, -1, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)
        If String.IsNullOrEmpty(itemPicture) Then
            If ReadProperty(classM, "CancelButtonPicture", mId, -1, itemPicture, arrs) = False Then Return
            itemPicture = UnWrapString(itemPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPicture = itemPicture.Replace("\", "/")

        'собственно замена картинки PictureActive
        sender.SetAttribute("src", itemPicture)
    End Sub

    Private Sub del_menu_cancel_MouseDown_Anchor(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim arrs() As String = {mId.ToString}

        'меняем картинку на PictureActive
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "CancelButtonPictureActive", mId, -1, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPictureActive)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPictureActive = itemPictureActive.Replace("\", "/")

        'собственно замена картинки PictureActive
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", itemPictureActive)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim itemType As Integer = 0
            ReadPropertyInt(classM, "CancelButtonType", mId, -1, itemType, arrs)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "CancelButtonPictureFloat", mId, -1, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "CancelButtonText", mId, -1, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPictureActive, itemPictureFloat)
        End If

    End Sub

    Private Sub del_menu_cancel_MouseUp_Anchor(sender As HtmlElement, e As HtmlElementEventArgs)
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim arrs() As String = {mId.ToString}

        'меняем обратно картинку с PictureActive на Picture/PictureHover
        'получаем PictureActive. Если пусто, то выход
        Dim itemPictureActive As String = ""
        If ReadProperty(classM, "CancelButtonPictureActive", mId, -1, itemPictureActive, arrs) = False Then Return
        itemPictureActive = UnWrapString(itemPictureActive)
        If String.IsNullOrEmpty(itemPictureActive) Then Return

        'получаем PictureHover, а если пусто - то Picture
        Dim itemPicture As String = ""
        If ReadProperty(classM, "CancelButtonPictureHover", mId, -1, itemPicture, arrs) = False Then Return
        itemPicture = UnWrapString(itemPicture)
        If String.IsNullOrEmpty(itemPicture) Then
            If ReadProperty(classM, "CancelButtonPicture", mId, -1, itemPicture, arrs) = False Then Return
            itemPicture = UnWrapString(itemPicture)
        End If

        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, itemPicture)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'убеждаемся в правильном формате пути к картинкам
        itemPicture = itemPicture.Replace("\", "/")

        'собственно замена картинки PictureActive на PictureHover/Picture
        If sender.Children.Count > 1 AndAlso sender.Children(0).TagName = "IMG" Then
            'есть IMG
            sender.Children(0).SetAttribute("src", itemPicture)
        Else
            'нет IMG
            'получаем тип. Подходят только типы кнопка и ссылка
            Dim itemType As Integer = 0
            ReadPropertyInt(classM, "CancelButtonType", mId, -1, itemType, arrs)
            Dim itemPictureFloat As Integer = 0
            ReadPropertyInt(classM, "CancelButtonPictureFloat", mId, -1, itemPictureFloat, arrs)
            Dim itemCaption As String = ""
            ReadProperty(classM, "CancelButtonText", mId, -1, itemCaption, arrs)
            MenuSetInnerHTML(sender, itemType, UnWrapString(itemCaption), itemPicture, itemPictureFloat)
        End If
    End Sub

    Public Sub del_menu_cancel_Click(sender As HtmlElement, e As HtmlElementEventArgs)
        'событие выбора действия
        Dim classM As Integer = mScript.mainClassHash("M")
        Dim mId As Integer = Val(sender.GetAttribute("menu"))
        Dim arrs() As String = {mId.ToString}
        Dim res As String = PropertiesRouter(classM, "Visible", {mId.ToString}, arrs, PropertiesOperationEnum.PROPERTY_SET, "False")
        If res <> "False" AndAlso res <> "#Error" AndAlso mId = GVARS.G_CURMENU Then
            GVARS.G_CURMENU = -1
        End If
    End Sub

#End Region

#Region "Timer"
    Public lstTimers As New SortedList(Of String, Timer)(StringComparer.CurrentCultureIgnoreCase)

    Public Sub InitTimers()
        Dim classT As Integer = mScript.mainClassHash("T")
        If IsNothing(mScript.mainClass(classT).ChildProperties) OrElse mScript.mainClass(classT).ChildProperties.Count = 0 Then Return

        For tId As Integer = 0 To mScript.mainClass(classT).ChildProperties.Count - 1
            Dim interval As Integer = 1000, enabled As Boolean = True
            Dim arr() As String = {tId.ToString}
            ReadPropertyInt(classT, "Interval", tId, -1, interval, arr)
            ReadPropertyBool(classT, "IsWorking", tId, -1, enabled, arr)
            TimerCreate(UnWrapString(mScript.mainClass(classT).ChildProperties(tId)("Name").Value), interval, enabled)
        Next
    End Sub

    'Создает новый таймер и его событие
    Public Sub TimerCreate(ByVal tName As String, ByVal interval As Integer, ByVal enabled As Boolean)
        Dim tim As New Timer With {.Enabled = enabled, .Interval = interval}
        lstTimers.Add(tName, tim)

        AddHandler tim.Tick, AddressOf del_tim_tick
    End Sub

    ''' <summary>
    ''' Удаляет таймер или все таймеры
    ''' </summary>
    ''' <param name="tId"></param>
    ''' <remarks></remarks>
    Public Sub TimerRemove(ByVal tId As Integer)
        Dim classT As Integer = mScript.mainClassHash("T")

        If tId < 0 Then
            'Удаляем все таймеры
            If IsNothing(mScript.mainClass(classT).ChildProperties) OrElse mScript.mainClass(classT).ChildProperties.Count = 0 Then Return
            For tId = 0 To mScript.mainClass(classT).ChildProperties.Count - 1
                'Удаляем таймер из lstTimers, удаляем событие и делаем Dispose
                Dim tName As String = UnWrapString(mScript.mainClass(classT).ChildProperties(tId)("Name").Value)
                Dim tim As Timer = lstTimers(tName)
                RemoveHandler tim.Tick, AddressOf del_tim_tick
                tim.Dispose()
            Next tId
            lstTimers.Clear()
        Else
            'Удаляем таймер из lstTimers, удаляем событие и делаем Dispose
            Dim tName As String = UnWrapString(mScript.mainClass(classT).ChildProperties(tId)("Name").Value)
            Dim tim As Timer = lstTimers(tName)
            RemoveHandler tim.Tick, AddressOf del_tim_tick
            lstTimers.Remove(tName)
            tim.Dispose()
        End If
    End Sub

    Private Sub del_tim_tick(ByVal sender As Object, e As EventArgs)
        If GVARS.HOLD_UP_TIMERS Then Return

        Dim classT As Integer = mScript.mainClassHash("T")
        Dim tName As String = lstTimers.ElementAt(lstTimers.IndexOfValue(sender)).Key
        Dim tId As Integer = GetSecondChildIdByName(WrapString(tName), mScript.mainClass(classT).ChildProperties)
        Dim arrs() As String = {tId.ToString}

        'Visits + 1
        Dim visits As Integer = 0
        ReadPropertyInt(classT, "Visits", tId, -1, visits, arrs)
        visits += 1
        Dim visitsRes As String = visits.ToString
        Dim trackingResult As String = mScript.trackingProperties.RunBefore(classT, "Visits", arrs, visitsRes)
        If trackingResult <> "False" AndAlso trackingResult <> "#Error" Then
            mScript.mainClass(classT).ChildProperties(tId)("Visits").Value = visitsRes
        End If

        'событие TimerEvent
        Dim res As String = ""
        Dim wasEvent As Boolean = False
        'глобальное
        Dim eventId As Integer = mScript.mainClass(classT).Properties("TimerEvent").eventId
        If eventId > 0 Then
            wasEvent = True
            res = mScript.eventRouter.RunEvent(eventId, arrs, "TimerEvent", False)
            If res = "#Error" Then Return
        End If

        'данного таймера
        If res <> "False" Then
            eventId = mScript.mainClass(classT).ChildProperties(tId)("TimerEvent").eventId
            If eventId > 0 Then
                wasEvent = True
                res = mScript.eventRouter.RunEvent(eventId, arrs, "TimerEvent", False)
                If res = "#Error" Then Return
            End If
        End If

        'событие ScriptFinishedEvent
        If wasEvent Then
            If mScript.eventRouter.RunScriptFinishedEvent(Nothing) = "#Error" Then Return
        End If

    End Sub
#End Region

#Region "Media"
    ''' <summary> Главный медиаплеер</summary>
    Public mPlayer As New MediaPlayer.MediaPlayer
    Public WithEvents timAudio As New Timer With {.Enabled = False, .Interval = 250}
    'Public lstSay As New List(Of MediaPlayer.MediaPlayer)

    ''' <summary>
    ''' Проигравыет произвольный аудиофайл, накладывая звук на фоновый
    ''' </summary>
    ''' <param name="fileName">Путь относительно папки квеста (без кавычек)</param>
    ''' <param name="volume">громкость 0 - 100</param>
    ''' <param name="shouldWait">True если ожидать окончания воспроизведения, иначе False</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Say(ByVal fileName As String, ByVal volume As Integer, ByVal shouldWait As Boolean) As String
        'убеждаемся в наличии файла
        Dim fullPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, fileName)
        If FileIO.FileSystem.FileExists(fullPath) = False Then
            Return _ERROR("Файл " & fileName & " не найден.", "Say")
        End If

        'Открываем файл
        Dim mp As New MediaPlayer.MediaPlayer
        mp.Open(fullPath)
        If volume < 0 Then
            volume = 0
        ElseIf volume > 100 Then
            volume = 100
        End If
        If volume <> 100 Then
            mp.Volume = (100 - volume) * -41
        End If

        'проигрываем
        mp.Play()

        'ждем загрузки
        Do
            If mp.PlayState = MediaPlayer.MPPlayStateConstants.mpWaiting Then
                Application.DoEvents()
            Else
                Exit Do
            End If
        Loop

        If shouldWait Then
            'ждем завершения
            Do
                If mp.PlayState <> MediaPlayer.MPPlayStateConstants.mpPlaying Then
                    mp = Nothing
                    Return ""
                End If
                Application.DoEvents()
            Loop
        Else
            'не ждем звершения - запускаем таймер, который потом удалит ненужный mp
            Dim t As New Timer With {.Enabled = True, .Interval = 1}
            AddHandler t.Tick, Sub(sender As Object, e As EventArgs)
                                   If mp.PlayState <> MediaPlayer.MPPlayStateConstants.mpPlaying Then
                                       t.Enabled = False
                                       mp = Nothing
                                       t.Dispose()
                                   End If
                               End Sub
            'Dim curMilliseconds As Long = New TimeSpan(Date.Now.Ticks).TotalMilliseconds
            'Dim finalMilliseconds As Long = curMilliseconds + Math.Round(mp.Duration * 1000) + 1

            'Dim bgWorker As New System.ComponentModel.BackgroundWorker With {.WorkerSupportsCancellation = True}

            'AddHandler bgWorker.DoWork, Sub(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
            '                                Do While curMilliseconds < finalMilliseconds
            '                                    If bgWorker.CancellationPending Then
            '                                        'сигнал к прекращению процесса
            '                                        sender.Dispose()
            '                                        Return
            '                                    End If
            '                                    Threading.Thread.Sleep(10)
            '                                    curMilliseconds = New TimeSpan(Date.Now.Ticks).TotalMilliseconds
            '                                Loop
            '                                sender.Dispose()
            '                            End Sub

            ''таймер для связи фонового процесса bgWorker и mp
            'Dim t As New Timer With {.Enabled = True, .Interval = 1}
            'AddHandler t.Tick, Sub(sender As Object, e As EventArgs)
            '                       If bgWorker.IsBusy = False Then
            '                           bgWorker.CancelAsync()
            '                           sender.Enabled = False
            '                           sender.Dispose()
            '                           mp = Nothing
            '                       End If
            '                   End Sub
            'bgWorker.RunWorkerAsync()
        End If

        Return ""
    End Function

    ''' <summary>Получаем аудиофайл, который будет проигран следующим</summary>
    Public Function AudioGetNext(ByRef arrParams() As String, ByVal prevAudio As Integer) As Integer
        Dim classId As Integer = mScript.mainClassHash("Med")
        Dim aCnt As Integer = 0
        Dim nextAudio As Integer = -1
        If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso IsNothing(mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("Name").ThirdLevelProperties) = False Then _
            aCnt = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("Name").ThirdLevelEventId.Count

        If aCnt = 1 Then
            'только 1 аудиофайл - проигрываем его постоянно
            nextAudio = 0
        ElseIf aCnt = 2 AndAlso prevAudio > -1 Then
            'только 2 аудиофайла, при этом один из них только что проиграл - выбираем второй
            If prevAudio = 0 Then
                nextAudio = 1
            Else
                nextAudio = 0
            End If
        Else
            Dim isAssorted As Boolean = False
            If ReadPropertyBool(classId, "PlayAssorted", GVARS.G_CURLIST, -1, isAssorted, arrParams) = False Then Return "#Error"

            If isAssorted Then
                'выбираем аудиофайл случайным образом
                Do While nextAudio = -1 OrElse nextAudio = prevAudio
                    nextAudio = Math.Round(Rnd() * (aCnt - 1))
                Loop
            Else
                'выбираем следующий
                If prevAudio < aCnt - 1 Then
                    nextAudio = prevAudio + 1
                Else
                    nextAudio = 0
                End If
            End If
        End If

        Return nextAudio
    End Function

    ''' <summary> Запускает проигрывание следующего аудиофайла в списке воспроизведения</summary>
    Public Function AudioPlayFromList(ByVal nextAudio As Integer, ByRef arrParams() As String) As String
        If GVARS.G_CURLIST < 0 Then
            Return _ERROR("Попытка воспроизведения аудио файла до выбора списка воспроизведения. Воспользуйтесь функцией Med.SelectList.")
        End If
        If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPlaying OrElse mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPaused Then AudioStop()

        Dim classId As Integer = mScript.mainClassHash("Med")
        Dim aCnt As Integer = 0
        If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso IsNothing(mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("Name").ThirdLevelProperties) = False Then _
            aCnt = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("Name").ThirdLevelEventId.Count
        If aCnt = 0 Then
            Return _ERROR("Попытка воспроизведения аудио файла из пустого списка воспроизведения.")
        End If

        Dim prevAudio As Integer = GVARS.G_CURAUDIO
        If nextAudio < 0 Then nextAudio = AudioGetNext(arrParams, prevAudio)

        GVARS.G_CURAUDIO = -1
        timAudioPrevTime = -1
        If nextAudio = -1 Then Return "-1"

        'событие TrackStarted
        Dim arrs() As String = {GVARS.G_CURLIST.ToString, nextAudio.ToString}
        Dim res As String = ""
        Dim wasEvent As Boolean = False
        'глобальное
        Dim eventId As Integer = mScript.mainClass(classId).Properties("TrackStarted").eventId
        If eventId > 0 Then
            wasEvent = True
            res = mScript.eventRouter.RunEvent(eventId, arrs, "TrackStarted", False)
            If res = "#Error" Then
                Return res
            ElseIf IsNumeric(res) Then
                Integer.TryParse(res, nextAudio)
            ElseIf res = "False" Then
                nextAudio = -1
            End If
        End If
        If nextAudio < 0 OrElse nextAudio > aCnt - 1 Then nextAudio = -1
        arrs(1) = nextAudio

        'данного списка воспроизведения
        If nextAudio >= 0 Then
            eventId = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("TrackStarted").eventId
            If eventId > 0 Then
                wasEvent = True
                res = mScript.eventRouter.RunEvent(eventId, arrs, "TrackStarted", False)
                If res = "#Error" Then
                    Return res
                ElseIf IsNumeric(res) Then
                    Integer.TryParse(res, nextAudio)
                ElseIf res = "False" Then
                    nextAudio = -1
                End If
            End If
        End If
        If nextAudio < 0 OrElse nextAudio > aCnt - 1 Then nextAudio = -1
        arrs(1) = nextAudio

        'данного аудиофайла
        If nextAudio >= 0 Then
            eventId = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("TrackStarted").ThirdLevelEventId(nextAudio)
            If eventId > 0 Then
                wasEvent = True
                res = mScript.eventRouter.RunEvent(eventId, arrs, "TrackStarted", False)
                If res = "#Error" Then
                    Return res
                ElseIf IsNumeric(res) Then
                    Integer.TryParse(res, nextAudio)
                ElseIf res = "False" Then
                    nextAudio = -1
                End If
            End If
        End If
        If nextAudio < 0 OrElse nextAudio > aCnt - 1 Then nextAudio = -1
        arrs(1) = nextAudio
        If nextAudio < 0 Then Return "-1"

        'получаем аудиофайл и необходимые для его запуска свойства
        Dim fName As String = "", volume As Integer = 100, InitialTime As Double = 0, VolumeIncreasing As Integer = 0
        If ReadProperty(classId, "FilePath", GVARS.G_CURLIST, nextAudio, fName, arrs) = False Then Return "#Error"
        fName = UnWrapString(fName)
        If String.IsNullOrEmpty(fName) Then Return _ERROR("Аудиофайл в списке " & GVARS.G_CURLIST.ToString & ", аудио с Id " & nextAudio.ToString & " не установлен")
        fName = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, fName)
        If FileIO.FileSystem.FileExists(fName) = False Then
            Return _ERROR("Аудиофайл " & fName & " не найден.")
        End If

        If ReadPropertyInt(classId, "Volume", GVARS.G_CURLIST, nextAudio, volume, arrs) = False Then Return "#Error"
        If volume < 0 Then
            volume = 0
        ElseIf volume > 100 Then
            volume = 100
        End If

        If ReadPropertyDbl(classId, "InitialTime", GVARS.G_CURLIST, nextAudio, InitialTime, arrs) = False Then Return "#Error"
        If InitialTime < 0 Then
            InitialTime = 0
        End If

        If ReadPropertyInt(classId, "VolumeIncreasing", GVARS.G_CURLIST, nextAudio, VolumeIncreasing, arrs) = False Then Return "#Error"
        If VolumeIncreasing < 0 Then
            VolumeIncreasing = 0
        End If

        'начало воспроизведения
        mPlayer.Open(fName)
        'If InitialTime > mPlayer.Duration Then InitialTime = CInt(mPlayer.Duration)
        'If VolumeIncreasing > mPlayer.Duration - InitialTime Then VolumeIncreasing = CInt(mPlayer.Duration - InitialTime)
        mPlayer.CurrentPosition = InitialTime

        If VolumeIncreasing = 0 Then
            mPlayer.Volume = (100 - volume) * -41
        Else
            mPlayer.Volume = -4100
        End If

        Try
            mPlayer.Play()
        Catch ex As Exception
            Return _ERROR("Не удалось воспроизвести " & fName & ".")
        End Try
        GVARS.G_CURAUDIO = nextAudio

        If VolumeIncreasing > 0 Then
            'плавное поднятие громкости
            ShiftVolumeAsynch(volume, VolumeIncreasing)
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Плавно изменяет громкость в новом потоке
    ''' </summary>
    ''' <param name="finalVolume">Конечная громкость</param>
    ''' <param name="sMilliseconds">Время изменения громкости в мс</param>
    Public Sub ShiftVolumeAsynch(ByVal finalVolume As Integer, ByVal sMilliseconds As Long)
        If GVARS.G_CURAUDIO < 0 Then Return

        If sMilliseconds <= 0 Then sMilliseconds = 1000

        Dim initMilliseconds As Long = New TimeSpan(Date.Now.Ticks).TotalMilliseconds
        Dim initVolume As Integer = 100 - Math.Round(-mPlayer.Volume / 41) '-4100 - 0
        Dim curMilliseconds As Long = New TimeSpan(Date.Now.Ticks).TotalMilliseconds - initMilliseconds
        'V_cur=((V_f-V_init ) T_cur)/T_f +V_init
        Dim initAudio As Integer = GVARS.G_CURAUDIO
        Dim curVolume As Integer = initVolume

        'создаем фоновый процесс для плавного повышения звука
        Dim bgWorker As New System.ComponentModel.BackgroundWorker With {.WorkerSupportsCancellation = True}

        AddHandler bgWorker.DoWork, Sub(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
                                        Do While curMilliseconds < sMilliseconds
                                            If bgWorker.CancellationPending Then
                                                'сигнал к прекращению процесса
                                                sender.Dispose()
                                                Return
                                            End If
                                            'устанавливаем 
                                            curVolume = (100 - Math.Round(((finalVolume - initVolume) * curMilliseconds) / sMilliseconds + initVolume)) * -41
                                            Threading.Thread.Sleep(10)
                                            curMilliseconds = New TimeSpan(Date.Now.Ticks).TotalMilliseconds - initMilliseconds
                                        Loop
                                        sender.Dispose()
                                    End Sub

        'таймер для связи фонового процесса bgWorker и mPlayer
        Dim t As New Timer With {.Enabled = True, .Interval = 1}
        AddHandler t.Tick, Sub(sender As Object, e As EventArgs)
                               If GVARS.G_CURAUDIO <> initAudio OrElse mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpStopped OrElse curVolume >= finalVolume OrElse bgWorker.IsBusy = False Then
                                   mPlayer.Volume = (100 - finalVolume) * -41
                                   bgWorker.CancelAsync()
                                   sender.Enabled = False
                                   sender.Dispose()
                               Else
                                   mPlayer.Volume = curVolume
                               End If
                           End Sub
        bgWorker.RunWorkerAsync()
    End Sub

    ''' <summary>
    ''' Плавно изменяет громкость в основном потоке
    ''' </summary>
    ''' <param name="finalVolume">Конечная громкость</param>
    ''' <param name="sMilliseconds">Время изменения громкости в мс</param>
    Public Sub ShiftVolumeSynch(ByVal finalVolume As Integer, ByVal sMilliseconds As Long)
        If GVARS.G_CURAUDIO < 0 Then Return

        If sMilliseconds <= 0 Then sMilliseconds = 1000

        Dim initMilliseconds As Long = New TimeSpan(Date.Now.Ticks).TotalMilliseconds
        Dim initVolume As Integer = 100 - Math.Round(-mPlayer.Volume / 41) '-4100 - 0
        Dim curMilliseconds As Long = New TimeSpan(Date.Now.Ticks).TotalMilliseconds - initMilliseconds
        'V_cur=((V_f-V_init ) T_cur)/T_f +V_init
        Dim initAudio As Integer = GVARS.G_CURAUDIO

        Do While curMilliseconds < sMilliseconds
            If GVARS.G_CURAUDIO <> initAudio OrElse mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpStopped Then
                mPlayer.Volume = (100 - finalVolume) * -41
                Return
            End If
            Dim curVolume As Integer = (100 - Math.Round(((finalVolume - initVolume) * curMilliseconds) / sMilliseconds + initVolume)) * -41
            mPlayer.Volume = curVolume
            curMilliseconds = New TimeSpan(Date.Now.Ticks).TotalMilliseconds - initMilliseconds
            Application.DoEvents()
        Loop
    End Sub

    ''' <summary>Останавливает воспроизведение текущего аудио</summary>
    Public Function AudioStop() As Boolean
        If GVARS.G_CURAUDIO = -1 Then Return True
        Dim classId As Integer = mScript.mainClassHash("Med")
        If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpStopped Then Return True
        mPlayer.Stop()
        Dim prevAudio As Integer = GVARS.G_CURAUDIO
        GVARS.G_CURAUDIO = -1
        timAudio.Enabled = False
        timAudioPrevTime = -1

        'событие TrackFinished
        Dim aCnt As Integer = 0
        If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso IsNothing(mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("Name").ThirdLevelProperties) = False Then _
            aCnt = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("Name").ThirdLevelEventId.Count
        If aCnt = 0 Then Return True

        Dim arrs() As String = {GVARS.G_CURLIST.ToString, prevAudio.ToString, "-1"}
        Dim res As String = ""
        'глобальное
        Dim eventId As Integer = mScript.mainClass(classId).Properties("TrackFinished").eventId
        If eventId > 0 Then
            res = mScript.eventRouter.RunEvent(eventId, arrs, "TrackFinished", False)
            If res = "#Error" Then
                Return False
            End If
        End If

        'данного списка воспроизведения
        If res <> "False" Then
            eventId = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("TrackFinished").eventId
            If eventId > 0 Then
                res = mScript.eventRouter.RunEvent(eventId, arrs, "TrackFinished", False)
                If res = "#Error" Then
                    Return False
                End If
            End If
        End If

        'данного аудиофайла
        If res <> "False" Then
            eventId = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("TrackFinished").ThirdLevelEventId(prevAudio)
            If eventId > 0 Then
                res = mScript.eventRouter.RunEvent(eventId, arrs, "TrackFinished", False)
                If res = "#Error" Then
                    Return False
                End If
            End If
        End If

        Return True
    End Function


    ''' <summary>Время от начала трека во время предыдущего тика</summary>
    Private timAudioPrevTime As Double = -1
    ''' <summary>Таймер проверяет не закончился ли аудиофайл. Если закончился, то вызывает событие TrackFinished и запускает следующее аудио</summary>
    Private Sub timAudio_Tick(sender As Object, e As EventArgs) Handles timAudio.Tick
        If GVARS.G_CURAUDIO = -1 Then
            timAudioPrevTime = -1
            Return
        End If
        Dim classId As Integer = mScript.mainClassHash("Med")
        Dim stopped As Boolean = False
        If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpStopped Then
            stopped = True
        Else
            Dim finalPos As Double = 0
            ReadPropertyDbl(classId, "FinalTime", GVARS.G_CURLIST, GVARS.G_CURAUDIO, finalPos, Nothing)
            If finalPos > 0 AndAlso mPlayer.CurrentPosition >= finalPos Then
                stopped = True
                mPlayer.Stop()
            End If
        End If

        If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpStopped Then
            'файл закончился
            timAudioPrevTime = -1
            'событие TrackFinished
            Dim nextAudio As Integer = AudioGetNext(Nothing, GVARS.G_CURAUDIO)
            Dim aCnt As Integer = 0
            If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso IsNothing(mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("Name").ThirdLevelProperties) = False Then _
                aCnt = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("Name").ThirdLevelEventId.Count
            If aCnt = 0 Then Return

            Dim arrs() As String = {GVARS.G_CURLIST.ToString, GVARS.G_CURAUDIO.ToString, nextAudio.ToString}
            Dim res As String = ""
            Dim wasEvent As Boolean = False
            'глобальное
            Dim eventId As Integer = mScript.mainClass(classId).Properties("TrackFinished").eventId
            If eventId > 0 Then
                wasEvent = True
                res = mScript.eventRouter.RunEvent(eventId, arrs, "TrackFinished", False)
                If res = "#Error" Then
                    Return
                ElseIf IsNumeric(res) Then
                    Integer.TryParse(res, nextAudio)
                ElseIf res = "False" Then
                    nextAudio = -1
                End If
            End If
            If nextAudio < 0 OrElse nextAudio > aCnt - 1 Then nextAudio = -1
            arrs(2) = nextAudio

            'данного списка воспроизведения
            If nextAudio >= 0 Then
                eventId = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("TrackFinished").eventId
                If eventId > 0 Then
                    wasEvent = True
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "TrackFinished", False)
                    If res = "#Error" Then
                        Return
                    ElseIf IsNumeric(res) Then
                        Integer.TryParse(res, nextAudio)
                    ElseIf res = "False" Then
                        nextAudio = -1
                    End If
                End If
            End If
            If nextAudio < 0 OrElse nextAudio > aCnt - 1 Then nextAudio = -1
            arrs(2) = nextAudio

            'данного аудиофайла
            If nextAudio >= 0 Then
                eventId = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("TrackFinished").ThirdLevelEventId(GVARS.G_CURAUDIO)
                If eventId > 0 Then
                    wasEvent = True
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "TrackFinished", False)
                    If res = "#Error" Then
                        Return
                    ElseIf IsNumeric(res) Then
                        Integer.TryParse(res, nextAudio)
                    ElseIf res = "False" Then
                        nextAudio = -1
                    End If
                End If
            End If
            If nextAudio < 0 OrElse nextAudio > aCnt - 1 Then nextAudio = -1
            arrs(2) = nextAudio

            'событие ScriptFinishedEvent
            If wasEvent Then
                If mScript.eventRouter.RunScriptFinishedEvent(Nothing) = "#Error" Then Return
            End If
            If nextAudio < 0 Then Return

            'Запуск следующего
            AudioPlayFromList(nextAudio, arrs)
        Else
            Dim MarkPosition As Double = 0
            ReadPropertyDbl(classId, "MarkPosition", GVARS.G_CURLIST, GVARS.G_CURAUDIO, MarkPosition, Nothing)
            Dim curPos As Double = mPlayer.CurrentPosition

            If MarkPosition > 0 AndAlso MarkPosition <= curPos AndAlso MarkPosition > timAudioPrevTime Then
                'Событие MarkInitiated
                timAudioPrevTime = curPos

                Dim arrs() As String = {GVARS.G_CURLIST.ToString, GVARS.G_CURAUDIO.ToString}
                Dim res As String = ""
                Dim wasEvent As Boolean = False
                'глобальное
                Dim eventId As Integer = mScript.mainClass(classId).Properties("MarkInitiated").eventId
                If eventId > 0 Then
                    wasEvent = True
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "MarkInitiated", False)
                    If res = "#Error" Then
                        Return
                    End If
                End If

                'данного списка воспроизведения
                eventId = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("MarkInitiated").eventId
                If eventId > 0 Then
                    wasEvent = True
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "MarkInitiated", False)
                    If res = "#Error" Then
                        Return
                    End If
                End If

                'данного аудиофайла
                eventId = mScript.mainClass(classId).ChildProperties(GVARS.G_CURLIST)("MarkInitiated").ThirdLevelEventId(GVARS.G_CURAUDIO)
                If eventId > 0 Then
                    wasEvent = True
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "MarkInitiated", False)
                    If res = "#Error" Then
                        Return
                    End If
                End If

                'событие ScriptFinishedEvent
                If wasEvent Then
                    If mScript.eventRouter.RunScriptFinishedEvent(Nothing) = "#Error" Then Return
                End If
            Else
                timAudioPrevTime = curPos
            End If

        End If
    End Sub

#End Region

#Region "LocWindow"
    Public Enum PrintInsertionEnum As Byte
        REPLACE = 0
        APPEND = 1
        REPLACE_NEW_BLOCK = 2
        APPEND_NEW_BLOCK = 3
        NEW_BLOCK = 4
    End Enum

    Public Enum PrintDataEnum As Byte
        TEXT = 0
        PICTURE = 1
    End Enum

    ''' <summary>
    ''' Печатает текст в главное окно
    ''' </summary>
    ''' <param name="newId">Id html-элемента или порядковый номер, начиная с 1, или отрицательный номер, если надо вести отсчет от последнего добавленного элемента</param>
    ''' <param name="strText">текст для печати (без кавычек)</param>
    ''' <param name="appendType"></param>
    ''' <param name="arrParams"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PrintTextToMainWindow(ByVal destId As String, ByVal strText As String, ByVal appendType As PrintInsertionEnum, ByVal newId As String, ByVal printData As PrintDataEnum, ByRef arrParams() As String, _
                                          Optional ByVal additionalStyles As String = "", Optional ByVal usePresets As Boolean = True, Optional ByVal picWidth As String = "", _
                                          Optional ByVal picHeight As String = "", Optional ByVal picAlign As String = "") As HtmlElement
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then
            _ERROR("Не открыт документ главного окна.", "Печать текста")
            Return Nothing
        End If

        If printData = PrintDataEnum.PICTURE Then
            'проверка существования картинки
            Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, strText)
            If FileIO.FileSystem.FileExists(fPath) = False Then
                _ERROR("Файла изображения не существует.", "Печать текста")
                Return Nothing
            End If
        End If

        Dim hConvas As HtmlElement, hPrint As HtmlElement = Nothing, toMainConvas As Boolean = True
        destId = UnWrapString(destId)
        additionalStyles = UnWrapString(additionalStyles)

        If String.IsNullOrEmpty(destId) OrElse appendType = PrintInsertionEnum.NEW_BLOCK Then
            'печатаем новой строкой
            hConvas = hDoc.GetElementById("MainConvas")
            If IsNothing(hConvas) Then
                _ERROR("Структура html-документа главного окна нарушена. Контейнер для вывода текста не найден.", "Печать текста")
                Return Nothing
            End If

            hPrint = hDoc.CreateElement("DIV")
            hConvas.AppendChild(hPrint)
            appendType = PrintInsertionEnum.NEW_BLOCK
        ElseIf IsNumeric(destId) Then
            'печатаем в элемент с указанным номером
            Dim pos As Integer = Val(destId)
            hConvas = hDoc.GetElementById("MainConvas")
            If IsNothing(hConvas) Then
                _ERROR("Структура html-документа главного окна нарушена. Контейнер для вывода текста не найден.", "Печать текста")
                Return Nothing
            End If
            If pos <= 0 Then
                pos = hConvas.Children.Count - 1 + pos
            Else
                pos -= 1
            End If
            If hConvas.Children.Count = 0 OrElse pos < 0 OrElse pos > hConvas.Children.Count - 1 Then
                hPrint = hDoc.CreateElement("DIV")
                hConvas.AppendChild(hPrint)
                appendType = PrintInsertionEnum.NEW_BLOCK
            Else
                hPrint = hConvas.Children(pos)
                If appendType = PrintInsertionEnum.REPLACE_NEW_BLOCK OrElse appendType = PrintInsertionEnum.REPLACE Then hPrint.InnerHtml = ""
                If appendType = PrintInsertionEnum.APPEND_NEW_BLOCK OrElse appendType = PrintInsertionEnum.REPLACE_NEW_BLOCK Then
                    Dim newEl As HtmlElement = Nothing
                    newEl = hDoc.CreateElement("DIV")

                    If hPrint.TagName = "TABLE" Then
                        'если вставляем новый фрагмент в таблицу после PrintInfo, то на самом деле вставляем в ее ячейку (туда, где основной текст)
                        Dim toLeft As Boolean = True
                        If hPrint.GetAttribute("Direction") = "right" Then toLeft = False
                        hPrint = hPrint.Children(0) 'TBODY
                        If hPrint.Children.Count > 0 Then
                            hPrint = hPrint.Children(hPrint.Children.Count - 1) 'TR
                            If toLeft Then
                                hPrint = hPrint.Children(hPrint.Children.Count - 1) 'last td
                            Else
                                hPrint = hPrint.Children(0) 'first TD
                            End If
                            hPrint.AppendChild(newEl)
                        Else
                            hPrint.AppendChild(newEl)
                        End If
                    Else
                        hPrint.AppendChild(newEl)
                    End If

                    hPrint = newEl
                End If
                toMainConvas = False
            End If
        Else
            'печатаем в элемент с указанным Id
            hPrint = hDoc.GetElementById(destId)
            If IsNothing(hPrint) Then
                hConvas = hDoc.GetElementById("MainConvas")
                If IsNothing(hConvas) Then
                    _ERROR("Структура html-документа главного окна нарушена. Контейнер для вывода текста не найден.", "Печать текста")
                    Return Nothing
                End If
                hPrint = hDoc.CreateElement("DIV")
                hConvas.AppendChild(hPrint)
                appendType = PrintInsertionEnum.NEW_BLOCK
            Else
                If appendType = PrintInsertionEnum.REPLACE_NEW_BLOCK OrElse appendType = PrintInsertionEnum.REPLACE Then hPrint.InnerHtml = ""
                If appendType = PrintInsertionEnum.APPEND_NEW_BLOCK OrElse appendType = PrintInsertionEnum.REPLACE_NEW_BLOCK Then
                    Dim newEl As HtmlElement = Nothing
                    newEl = hDoc.CreateElement("DIV")

                    If hPrint.TagName = "TABLE" Then
                        'если вставляем новый фрагмент в таблицу после PrintInfo, то на самом деле вставляем в ее ячейку (туда, где основной текст)
                        Dim toLeft As Boolean = True
                        If hPrint.GetAttribute("Direction") = "right" Then toLeft = False
                        hPrint = hPrint.Children(0) 'TBODY
                        If hPrint.Children.Count > 0 Then
                            hPrint = hPrint.Children(hPrint.Children.Count - 1) 'TR
                            If toLeft Then
                                hPrint = hPrint.Children(hPrint.Children.Count - 1) 'last td
                            Else
                                hPrint = hPrint.Children(0) 'first TD
                            End If
                            hPrint.AppendChild(newEl)
                        Else
                            hPrint.AppendChild(newEl)
                        End If
                    Else
                        hPrint.AppendChild(newEl)
                    End If

                    hPrint = newEl
                ElseIf appendType = PrintInsertionEnum.APPEND AndAlso hPrint.TagName = "TABLE" Then
                    'если вставляем новый фрагмент в таблицу после PrintInfo, то на самом деле вставляем в ее ячейку (туда, где основной текст)
                    Dim toLeft As Boolean = True
                    If hPrint.GetAttribute("Direction") = "right" Then toLeft = False
                    hPrint = hPrint.Children(0) 'TBODY
                    If hPrint.Children.Count > 0 Then
                        hPrint = hPrint.Children(hPrint.Children.Count - 1) 'TR
                        If toLeft Then
                            hPrint = hPrint.Children(hPrint.Children.Count - 1) 'last td
                        Else
                            hPrint = hPrint.Children(0) 'first TD
                        End If
                    End If
                End If
                toMainConvas = False
            End If
        End If
        If String.IsNullOrEmpty(newId) = False Then hPrint.Id = newId

        If toMainConvas = False AndAlso String.IsNullOrEmpty(hPrint.Id) = False AndAlso hPrint.Id = "MainConvas" Then toMainConvas = True

        'получаем преднастройки печати
        'позиционируем и устанавливаем класс блока текста
        Dim classId As Integer = mScript.mainClassHash("LW")
        Dim dLeft As String = "", dTop As String = "", dWidth As String = "", dHeight As String = "", dClass As String = "", dBackColor As String = "", dForeColor As String = ""

        If appendType <> PrintInsertionEnum.APPEND AndAlso appendType <> PrintInsertionEnum.REPLACE AndAlso usePresets Then
            Dim strStyle As New System.Text.StringBuilder
            If ReadProperty(classId, "PrintPreset_OffsetByX", -1, -1, dLeft, arrParams) = False Then Return Nothing
            If ReadProperty(classId, "PrintPreset_OffsetByY", -1, -1, dTop, arrParams) = False Then Return Nothing
            If ReadProperty(classId, "PrintPreset_Width", -1, -1, dWidth, arrParams) = False Then Return Nothing
            If ReadProperty(classId, "PrintPreset_Height", -1, -1, dHeight, arrParams) = False Then Return Nothing
            If ReadProperty(classId, "PrintPreset_CssClass", -1, -1, dClass, arrParams) = False Then Return Nothing
            If ReadProperty(classId, "PrintPreset_BackColor", -1, -1, dBackColor, arrParams) = False Then Return Nothing
            If ReadProperty(classId, "PrintPreset_TextColor", -1, -1, dForeColor, arrParams) = False Then Return Nothing

            dLeft = UnWrapString(dLeft)
            dTop = UnWrapString(dTop)
            dWidth = UnWrapString(dWidth)
            dHeight = UnWrapString(dHeight)
            dBackColor = UnWrapString(dBackColor)
            dForeColor = UnWrapString(dForeColor)
            dClass = UnWrapString(dClass)
            If String.IsNullOrEmpty(dClass) = False Then hPrint.SetAttribute("ClassName", dClass)

            If String.IsNullOrEmpty(dLeft) = False Then
                If IsNumeric(dLeft) Then dLeft &= "px"
                strStyle.Append("left:" & dLeft & ";")
            End If
            If String.IsNullOrEmpty(dTop) = False Then
                If IsNumeric(dTop) Then dTop &= "px"
                strStyle.Append("top:" & dTop & ";")
            ElseIf GVARS.G_CURLOC > -1 AndAlso toMainConvas Then
                'если PrintPreset_Height пусто и печатаем в MainConvas, то устанавливаем высоту из OffsetByY текущей локации (чтобы печатаемый текст не оказался под блоком описания)
                Dim lTop As String = ""
                ReadProperty(mScript.mainClassHash("L"), "OffsetByY", GVARS.G_CURLOC, -1, lTop, arrParams)
                lTop = UnWrapString(lTop)
                If String.IsNullOrEmpty(lTop) = False Then
                    If IsNumeric(lTop) Then lTop &= "px"
                    strStyle.Append("top:" & lTop & ";")
                End If
            End If
            If String.IsNullOrEmpty(dWidth) = False Then
                If IsNumeric(dWidth) Then dWidth &= "px"
                strStyle.Append("width:" & dWidth & "!important;")
            ElseIf String.IsNullOrEmpty(dLeft) = False Then
                'автопределение ширины
                strStyle.Append("width:calc(100% - " & dLeft & " - 20px)!important;")
            End If
            If String.IsNullOrEmpty(dHeight) = False Then
                If IsNumeric(dHeight) Then dHeight &= "px"
                strStyle.Append("height:" & dHeight & ";overflow:scroll;")
            End If
            If String.IsNullOrEmpty(dBackColor) = False Then
                strStyle.Append("background:" & dBackColor & ";")
            End If
            If String.IsNullOrEmpty(dForeColor) = False Then
                strStyle.Append("color:" & dForeColor & ";")
            End If
            If strStyle.Length > 0 Then strStyle.Append("position:relative;")

            If printData = PrintDataEnum.TEXT Then
                'дополнительные стили
                If String.IsNullOrEmpty(additionalStyles) = False Then
                    strStyle.Append(additionalStyles)
                End If

                If strStyle.Length > 0 Then
                    hPrint.Style = strStyle.ToString
                    strStyle.Clear()
                End If
            End If
        Else
            Dim strStyles As String = hPrint.Style
            If String.IsNullOrEmpty(strStyles) = False AndAlso strStyles.EndsWith(";") = False Then strStyles &= ";"
            If toMainConvas Then
                'добаляем высоту разположения описания, чтобы текст не оказался под ним
                Dim lTop As String = ""
                ReadProperty(mScript.mainClassHash("L"), "OffsetByY", GVARS.G_CURLOC, -1, lTop, arrParams)
                lTop = UnWrapString(lTop)
                If String.IsNullOrEmpty(lTop) = False Then
                    If IsNumeric(lTop) Then lTop &= "px"
                    strStyles &= "top:" & lTop & ";position:relative;"
                End If
                hPrint.Style = strStyles
                strStyles = ""
            End If

            If printData = PrintDataEnum.TEXT Then
                'дополнительные стили
                If String.IsNullOrEmpty(additionalStyles) = False Then
                    strStyles &= additionalStyles
                End If
            End If
            If String.IsNullOrEmpty(strStyles) = False Then hPrint.Style = strStyles
        End If

        Dim imgSizeStyle As String = "", hImg As HtmlElement = Nothing
        If printData = PrintDataEnum.PICTURE Then
            If IsNumeric(picWidth) Then picWidth &= "px"
            If picWidth.Length > 0 Then imgSizeStyle &= "width:" & picWidth & ";"
            If IsNumeric(picHeight) Then picHeight &= "px"
            If picHeight.Length > 0 Then imgSizeStyle &= "height:" & picHeight & ";"
            'дополнительные стили
            If String.IsNullOrEmpty(additionalStyles) = False Then
                imgSizeStyle &= additionalStyles
            End If

            hImg = hDoc.CreateElement("IMG")
            hImg.SetAttribute("src", strText)
            hImg.Style = imgSizeStyle
        End If

        If appendType = PrintInsertionEnum.APPEND OrElse appendType = PrintInsertionEnum.REPLACE Then
            If printData = PrintDataEnum.TEXT Then
                hPrint.InnerHtml &= strText
            ElseIf printData = PrintDataEnum.PICTURE Then
                hPrint.AppendChild(hImg)
            End If
        Else
            If printData = PrintDataEnum.TEXT Then
                hPrint.InnerHtml = strText
            ElseIf printData = PrintDataEnum.PICTURE Then
                If picAlign.Length > 0 Then
                    Dim strStyles As String = hPrint.Style
                    If String.IsNullOrEmpty(strStyles) = False AndAlso strStyles.EndsWith(";") = False Then strStyles &= ";"
                    If picAlign.Length > 0 Then strStyles &= "text-align:" & picAlign & " !important"
                    If strStyles.Length > 0 Then hPrint.Style = strStyles
                End If
                hPrint.AppendChild(hImg)
            End If
        End If

        If GVARS.G_ISBATTLE Then
            Dim classB As Integer = mScript.mainClassHash("B"), keepHistory As Boolean = False
            If ReadPropertyBool(classB, "KeepHistory", -1, -1, keepHistory, Nothing) = False Then Return Nothing
            If keepHistory Then mScript.Battle.lstHistory.Add(hPrint.OuterHtml)
        End If
        Return hPrint
    End Function

    Public Function PrintInfoToMainWindow(ByVal elId As String, ByVal strText As String, ByVal imgPath As String, ByVal strCaption As String, ByVal printStyle As String, _
                                          ByVal appendType As PrintInsertionEnum, ByVal newId As String, ByVal imgFromLeft As Boolean, ByVal imgWidth As String, ByVal imgHeight As String, _
                                          ByVal bgText As String, ByVal bgImg As String, ByVal bgCaption As String, ByVal fgText As String, ByVal fgCaption As String, ByRef arrParams() As String) As HtmlElement
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then
            _ERROR("Не открыт документ главного окна.", "Печать текста")
            Return Nothing
        End If

        'проверка существования картинки
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, imgPath)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            _ERROR("Файла изображения не существует.", "Печать текста")
            Return Nothing
        End If

        Dim hConvas As HtmlElement, hPrint As HtmlElement = Nothing, toMainConvas As Boolean = True
        elId = UnWrapString(elId)

        If String.IsNullOrEmpty(elId) OrElse appendType = PrintInsertionEnum.NEW_BLOCK Then
            'печатаем новой строкой
            hConvas = hDoc.GetElementById("MainConvas")
            If IsNothing(hConvas) Then
                _ERROR("Структура html-документа главного окна нарушена. Контейнер для вывода текста не найден.", "Печать текста")
                Return Nothing
            End If
            hPrint = hDoc.CreateElement("TABLE")
            hConvas.AppendChild(hPrint)
            appendType = PrintInsertionEnum.NEW_BLOCK
        ElseIf IsNumeric(elId) Then
            'печатаем в элемент с указанным номером
            Dim pos As Integer = Val(elId)
            hConvas = hDoc.GetElementById("MainConvas")
            If IsNothing(hConvas) Then
                _ERROR("Структура html-документа главного окна нарушена. Контейнер для вывода текста не найден.", "Печать текста")
                Return Nothing
            End If
            If pos <= 0 Then
                pos = hConvas.Children.Count - 1 + pos
            Else
                pos -= 1
            End If
            If hConvas.Children.Count = 0 OrElse pos < 0 OrElse pos > hConvas.Children.Count - 1 Then
                hPrint = hDoc.CreateElement("TABLE")
                hConvas.AppendChild(hPrint)
                appendType = PrintInsertionEnum.NEW_BLOCK
            Else
                hPrint = hConvas.Children(pos)
                If appendType = PrintInsertionEnum.REPLACE_NEW_BLOCK OrElse appendType = PrintInsertionEnum.REPLACE Then hPrint.InnerHtml = ""
                If hPrint.TagName <> "TABLE" Then
                    If appendType = PrintInsertionEnum.APPEND Then
                        appendType = PrintInsertionEnum.APPEND_NEW_BLOCK
                    ElseIf appendType = PrintInsertionEnum.REPLACE Then
                        appendType = PrintInsertionEnum.REPLACE_NEW_BLOCK
                    End If
                End If

                If appendType = PrintInsertionEnum.APPEND_NEW_BLOCK OrElse appendType = PrintInsertionEnum.REPLACE_NEW_BLOCK Then
                    Dim newEl As HtmlElement = Nothing
                    newEl = hDoc.CreateElement("TABLE")

                    If hPrint.TagName = "TABLE" Then
                        Dim toLeft As Boolean = True
                        If hPrint.GetAttribute("Direction") = "right" Then toLeft = False
                        hPrint = hPrint.Children(0) 'TBODY
                        If hPrint.Children.Count > 0 Then
                            hPrint = hPrint.Children(hPrint.Children.Count - 1) 'TR
                            If toLeft Then
                                hPrint = hPrint.Children(hPrint.Children.Count - 1) 'last td
                            Else
                                hPrint = hPrint.Children(0) 'first TD
                            End If
                            hPrint.AppendChild(newEl)
                        Else
                            hPrint.AppendChild(newEl)
                        End If
                    Else
                        hPrint.AppendChild(newEl)
                    End If

                    hPrint = newEl
                ElseIf hPrint.TagName = "TABLE" Then
                    'Append / Replace to TABLE
                    Dim TR As HtmlElement = hDoc.CreateElement("TR")
                    hPrint = hPrint.Children(0) 'TBODY
                    hPrint.AppendChild(TR)
                    hPrint = TR
                End If
                toMainConvas = False
            End If
        Else
            'печатаем в элемент с указанным Id
            hPrint = hDoc.GetElementById(elId)
            If IsNothing(hPrint) Then
                hConvas = hDoc.GetElementById("MainConvas")
                If IsNothing(hConvas) Then
                    _ERROR("Структура html-документа главного окна нарушена. Контейнер для вывода текста не найден.", "Печать текста")
                    Return Nothing
                End If
                hPrint = hDoc.CreateElement("TABLE")
                hConvas.AppendChild(hPrint)
                appendType = PrintInsertionEnum.NEW_BLOCK
            Else
                If appendType = PrintInsertionEnum.REPLACE_NEW_BLOCK OrElse appendType = PrintInsertionEnum.REPLACE Then hPrint.InnerHtml = ""
                If appendType = PrintInsertionEnum.APPEND_NEW_BLOCK OrElse appendType = PrintInsertionEnum.REPLACE_NEW_BLOCK Then
                    Dim newEl As HtmlElement = Nothing
                    newEl = hDoc.CreateElement("TABLE")

                    If hPrint.TagName = "TABLE" Then
                        Dim toLeft As Boolean = True
                        If hPrint.GetAttribute("Direction") = "right" Then toLeft = False
                        hPrint = hPrint.Children(0) 'TBODY
                        If hPrint.Children.Count > 0 Then
                            hPrint = hPrint.Children(hPrint.Children.Count - 1) 'TR
                            If toLeft Then
                                hPrint = hPrint.Children(hPrint.Children.Count - 1) 'last td
                            Else
                                hPrint = hPrint.Children(0) 'first TD
                            End If
                            hPrint.AppendChild(newEl)
                        Else
                            hPrint.AppendChild(newEl)
                        End If
                    Else
                        hPrint.AppendChild(newEl)
                    End If
                    hPrint = newEl
                ElseIf hPrint.TagName = "TABLE" Then
                    'Append / Replace to TABLE
                    Dim TR As HtmlElement = hDoc.CreateElement("TR")
                    hPrint = hPrint.Children(0) 'TBODY
                    hPrint.AppendChild(TR)
                    hPrint = TR
                Else
                    'Append / Replace to DIV
                    Dim newEl As HtmlElement = Nothing
                    newEl = hDoc.CreateElement("TABLE")
                    hPrint.AppendChild(newEl)
                    hPrint = newEl
                End If
                toMainConvas = False
            End If
        End If
        If toMainConvas = False AndAlso String.IsNullOrEmpty(hPrint.Id) = False AndAlso hPrint.Id = "MainConvas" Then toMainConvas = True
        If String.IsNullOrEmpty(newId) = False Then hPrint.Id = newId

        'позиционируем и устанавливаем класс блока информации
        'Dim classId As Integer = mScript.mainClassHash("LW")
        Dim TDimg As HtmlElement = hDoc.CreateElement("TD")
        Dim TDtxt As HtmlElement = hDoc.CreateElement("TD")

        If hPrint.TagName = "TABLE" Then
            Dim TBODY As HtmlElement = hDoc.CreateElement("TBODY")
            hPrint.AppendChild(TBODY)
            Dim TR As HtmlElement = hDoc.CreateElement("TR")
            TBODY.AppendChild(TR)

            If imgFromLeft Then
                TR.AppendChild(TDimg)
                TR.AppendChild(TDtxt)
                hPrint.SetAttribute("Direction", "left")
            Else
                TR.AppendChild(TDtxt)
                TR.AppendChild(TDimg)
                hPrint.SetAttribute("Direction", "right")
            End If

            'set class
            Dim className As String = "info_" & printStyle
            If imgFromLeft Then
                className &= "_left"
            Else
                className &= "_right"
            End If
            hPrint.SetAttribute("ClassName", className)

            'set top
            If toMainConvas AndAlso GVARS.G_CURLOC > -1 Then
                Dim classL As Integer = mScript.mainClassHash("L")
                Dim strTop As String = ""
                ReadProperty(classL, "OffsetByY", GVARS.G_CURLOC, -1, strTop, arrParams)
                strTop = UnWrapString(strTop)
                If String.IsNullOrEmpty(strTop) = False Then
                    If IsNumeric(strTop) Then strTop &= "px"
                    Dim tStyle As String = hPrint.Style
                    If String.IsNullOrEmpty(tStyle) = False AndAlso tStyle.EndsWith(";") = False Then tStyle &= ";"
                    tStyle &= "position:relative;top:" & strTop
                    hPrint.Style = tStyle
                End If
            End If
        Else 'TR
            If imgFromLeft Then
                hPrint.AppendChild(TDimg)
                hPrint.AppendChild(TDtxt)
            Else
                hPrint.AppendChild(TDtxt)
                hPrint.AppendChild(TDimg)
            End If
        End If

        Dim strStyle As New System.Text.StringBuilder
        'image
        Dim IMG As HtmlElement = hDoc.CreateElement("IMG")
        IMG.SetAttribute("src", imgPath)
        If imgWidth.Length > 0 Then
            If IsNumeric(imgWidth) Then imgWidth &= "px"
            strStyle.Append("width:" & imgWidth & ";")
        End If
        If imgHeight.Length > 0 Then
            If IsNumeric(imgHeight) Then imgHeight &= "px"
            strStyle.Append("height:" & imgHeight & ";")
        End If
        If bgImg.Length > 0 Then
            strStyle.Append("background:" & bgImg & ";")
        End If
        If strStyle.Length > 0 Then IMG.Style = strStyle.ToString
        TDimg.AppendChild(IMG)
        strStyle.Clear()

        'caption
        If String.IsNullOrEmpty(strCaption) = False Then
            Dim CAPT As HtmlElement = hDoc.CreateElement("DIV")
            CAPT.SetAttribute("ClassName", "Caption")
            CAPT.InnerHtml = strCaption
            If fgCaption.Length > 0 Then
                strStyle.Append("color:" & fgCaption & ";")
            End If
            If bgCaption.Length > 0 Then
                strStyle.Append("background:" & bgCaption & ";")
            End If
            If strStyle.Length > 0 Then CAPT.Style = strStyle.ToString
            TDimg.AppendChild(CAPT)
            strStyle.Clear()
        End If

        'text
        If fgText.Length > 0 Then
            strStyle.Append("color:" & fgText & ";")
        End If
        If bgText.Length > 0 Then
            strStyle.Append("background:" & bgText & ";")
        End If
        If strStyle.Length > 0 Then TDtxt.Style = strStyle.ToString
        TDtxt.InnerHtml = strText

        If GVARS.G_ISBATTLE Then
            Dim classB As Integer = mScript.mainClassHash("B"), keepHistory As Boolean = False
            If ReadPropertyBool(classB, "KeepHistory", -1, -1, keepHistory, Nothing) = False Then Return Nothing
            If keepHistory Then mScript.Battle.lstHistory.Add(hPrint.OuterHtml)
        End If
        Return hPrint
    End Function

    Public Function RemoveText(ByVal elId As String, ByRef arrParams() As String) As Boolean
        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDoc) Then Return False
        elId = UnWrapString(elId)

        If IsNumeric(elId) Then
            Dim hConvas As HtmlElement = hDoc.GetElementById("MainConvas")
            If IsNothing(hConvas) OrElse hConvas.Children.Count = 0 Then Return False
            Dim pos As Integer = Val(elId)
            If pos <= 0 Then
                pos = hConvas.Children.Count - 1 + pos
                If pos < 0 Then Return False
            Else
                pos -= 1
            End If

            If pos > hConvas.Children.Count - 1 Then Return False
            Dim hNode As mshtml.IHTMLDOMNode = hConvas.Children(pos).DomElement
            hNode.removeNode(True)
        Else
            Dim hEL As HtmlElement = hDoc.GetElementById(elId)
            If IsNothing(hEL) Then Return False
            Dim hNode As mshtml.IHTMLDOMNode = hEL.DomElement
            hNode.removeNode(True)
        End If

        Return True
    End Function
#End Region
End Module
