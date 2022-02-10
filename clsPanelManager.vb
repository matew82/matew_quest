Public Class clsPanelManager
    Public Class PanelEx
        Inherits Panel
        'Public line1Top As New clsPanelManager.cPropertyControlsTop
        'Public line2Top As New clsPanelManager.cPropertyControlsTop
        'Public line1 As Microsoft.VisualBasic.PowerPacks.LineShape
        'Public line2 As Microsoft.VisualBasic.PowerPacks.LineShape
        ''' <summary>Чекбокс указывающий отображать ли кнопки настроек</summary>
        Public chkConfig As CheckBox
        ''' <summary>Кнопка перехода родитель/дочь</summary>
        Public NavigationButton As Control
        ''' <summary>Хранит правый край (lbl.Right) самой длинной надписи перед контролом свойства</summary>
        Public maxRight As Integer
        ''' <summary>Загружены ли базовые функции</summary>
        Public showBasicFunctions As Boolean = False
        ''' <summary>Ссылка на дерево с элементами 3 уровня/действиями для локаций</summary>
        Public subTree As TreeView
        ''' <summary>Ширина данной панели</summary>
        Public CurrentWidth As Integer = 0
        ''' <summary>Высота панели, где вводится описания</summary>
        Public CurrentInfoPanelHeight = questEnvironment.DefaultInfoPanelHeight
    End Class

    ''' <summary>Положение по У для контролов свойств на панелях управления свойствами по умолчанию, 2 и 3 уровня</summary>
    Public Class cPropertyControlsTop
        Public TopLevel1 As Integer = 0
        Public TopLevel2 As Integer = 0
        Public TopLevel3 As Integer = 0

        ''' <summary>
        ''' Рассчитывает положение по У для всех 3 случаев с учетом указанного смещения
        ''' </summary>
        ''' <param name="originTop">Исходное положение по У</param>
        ''' <param name="offset">Смещение для 3 случаев</param>
        Public Sub CalculatePosition(ByVal originTop As Integer, ByRef offset As cPropertyControlsTop)
            TopLevel1 = originTop - offset.TopLevel1
            TopLevel2 = originTop - offset.TopLevel2
            TopLevel3 = originTop - offset.TopLevel3
        End Sub

        ''' <summary>
        ''' Возвращает положение по Y исходя из уровня элемента (по умолчанию, 2 или 3)
        ''' </summary>
        ''' <param name="child2Id">Id элемента 2 порядка</param>
        ''' <param name="child3Id">Id элемента 3 порядка</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTop(ByVal child2Id As Integer, ByVal child3Id As Integer) As Integer
            If child2Id < 0 Then Return TopLevel1
            If child3Id < 0 Then Return TopLevel2
            Return TopLevel3
        End Function
    End Class

    ''' <summary>Список всех открытых окон-панелей (= экземпляров классов clsChildPanel)</summary>
    Public lstPanels As New List(Of clsChildPanel)
    ''' <summary>Родительский контейнер, в который надо поместить панель</summary>
    Public Property parent As SplitContainer
    ''' <summary>Ссылка на браузер для отображения нужной информации (он один на всех)</summary>
    Public Property WBhelp As WebBrowser
    ''' <summary>Тулбар для вывода вкладок открытых панелей</summary>
    Public Property panelToolStrip As ToolStrip
    ''' <summary>Хранит ссылку на кнопку для выбора настроек по умолчанию</summary>
    Public Property btnDefSettings As Button
    ''' <summary>Открытая в данный момент панель</summary>
    Public Property ActivePanel As clsChildPanel = Nothing

    ''' <summary>Содержит ссылки на контейнеры со свойствами каждого класса, ключом является classId</summary>
    Public dictDefContainers As New Dictionary(Of Integer, PanelEx)
    ''' <summary>Содержит ссылки на последние открытые вкладки в каждом классе</summary>
    Public dictLastPanel As New Dictionary(Of Integer, clsChildPanel)
    ''' <summary>Используется для отмены некоторых нежелательных событий</summary>
    Public ignoreNextEvent As Boolean = False

    ''' <summary>
    ''' Сохраняет текущую панель как последнюю выбранную в данном классе и возвращает панель, которая была выбрана в этом классе раньше. Испльзуется при смене класса
    ''' </summary>
    ''' <param name="className">Имя класса, которому устанавливается последняя панель</param>
    Public Function SetLastPanel(ByVal className As String) As clsChildPanel
        If IsNothing(ActivePanel) Then Return Nothing
        Dim classId As Integer
        If className = "Variable" Then
            classId = -2
        ElseIf className = "Function" Then
            classId = -3
        Else
            classId = mScript.mainClassHash(className)
        End If

        If dictLastPanel.ContainsKey(classId) Then
            Dim prevPanel As clsChildPanel = dictLastPanel(classId)
            dictLastPanel(classId) = ActivePanel
            Return prevPanel
        Else
            dictLastPanel.Add(classId, ActivePanel)
            Return Nothing
        End If
    End Function

#Region "Делегаты"
    ''' <summary>Делегат для событий ButtonClick кнопок помощи событий панели свойств</summary>
    Private Sub del_propertyButtonHelpMouseClick(sender As Object, e As EventArgs)
        'Отображение файла помощи о событии в браузере + код события в html
        Dim propName As String = sender.Name
        propName = propName.Substring(0, propName.Length - 4) 'Имя свойства (напр, было DescriptionHelp, а стало просто Description)
        Dim ch As clsChildPanel = sender.childPanel
        'If IsNothing(sender.Parent) Then Return

        If IsNothing(sender.Parent) Then
            sender.Dispose()
            frmMainEditor.codeBoxPanel.Hide()
            Return
        End If

        Dim c As Control = sender.Parent.Controls(propName)
        ch.ActiveControl = c

        Dim defProp As MatewScript.PropertiesInfoType
        If sender.IsFunctionButton Then
            defProp = mScript.mainClass(ch.classId).Functions(propName)
        Else
            defProp = mScript.mainClass(ch.classId).Properties(propName)
        End If
        Dim hFile As String = GetHelpPath(defProp.helpFile)
        WBhelp.Tag = Nothing
        If String.IsNullOrEmpty(hFile) Then
            WBhelp.Visible = False
        Else
            WBhelp.Navigate(hFile)
            frmMainEditor.HtmlShowEventText(sender)
            WBhelp.Visible = True
        End If
    End Sub

    ''' <summary> Показываем help-файл для обычных свойств (не событий) </summary>
    ''' <param name="ch">Панель-родитель свойства</param>
    ''' <param name="propName">Имя свойства</param>
    Private Sub ShowHelpFileForEventProperty(ByRef ch As clsChildPanel, ByVal propName As String)
        Dim hFile As String = GetHelpPath(mScript.mainClass(ch.classId).Properties(propName).helpFile)
        WBhelp.Tag = Nothing
        If String.IsNullOrEmpty(hFile) Then
            WBhelp.Visible = False
        Else
            WBhelp.Navigate(hFile)
            WBhelp.Visible = True
        End If
    End Sub

    ''' <summary>Делегат для событий ButtonClick кнопок событий панели свойств</summary>
    Public Sub del_propertyButtonMouseClick(sender As Object, e As EventArgs)
        If IsNothing(sender.childPanel) Then Return
        Dim ch As clsChildPanel = sender.childPanel
        Dim classId As Integer = ch.classId
        Dim child2Id As Integer = ch.GetChild2Id
        Dim child3Id As Integer = ch.GetChild3Id(child2Id)
        Dim propName As String = sender.Name
        Dim cb As CodeTextBox = frmMainEditor.codeBox
        Dim cbPanel As Panel = frmMainEditor.codeBoxPanel
        cbPanel.Hide()
        cb.Tag = sender
        'код или описание
        If sender.IsFunctionButton = False AndAlso mScript.mainClass(classId).Properties(propName).returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION Then
            cb.codeBox.IsTextBlockByDefault = True 'описания (например, L.Description)
        Else
            cb.codeBox.IsTextBlockByDefault = False 'Обработчики
        End If

        If child2Id = -1 Then
            If sender.IsFunctionButton Then
                cb.codeBox.LoadCodeFromProperty(mScript.mainClass(classId).Functions(propName).Value)
            Else
                cb.codeBox.LoadCodeFromProperty(mScript.mainClass(classId).Properties(propName).Value)
            End If
        ElseIf child3Id = -1 Then
            cb.codeBox.LoadCodeFromProperty(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value)
        Else
            cb.codeBox.LoadCodeFromProperty(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id))
        End If
        ActivePanel.ActiveControl = sender
        cbPanel.Show()
        WBhelp.Hide()
    End Sub

    ''' <summary>Делегат для событий MouseDown текст/комбобоксов событий панели свойств</summary>
    Public Sub del_TextPropertyMouseDown(sender As Object, e As MouseEventArgs)
        'If IsNothing(sender.Parent) Then Return 'элемента нет на форме
        If sender.Visible = False Then Return
        If mScript.LAST_ERROR.Length > 0 Then Return

        Dim ch As clsChildPanel = sender.childPanel

        If sender.GetType.Name <> "ComboBoxEx" Then
            If sender.SelectionLength < sender.TextLength Then sender.SelectAll()
        End If
        If Object.Equals(sender, ch.ActiveControl) Then Return
        'If IsNothing(ch.ActiveControl) OrElse Object.Equals(sender, ch.ActiveControl) Then Return
        'Выделяем цветом новое свойство и снимаем выделение со старого
        'ActivePanel.ActiveControl = sender
        If sender.Visible Then ActivePanel.ActiveControl = sender

        Dim propValue As String
        Dim propName As String = sender.Name
        Dim cRes As MatewScript.ContainsCodeEnum
        Dim child2Id As Integer = ch.GetChild2Id
        Dim child3Id As Integer = ch.GetChild3Id(child2Id)
        If child2Id < 0 Then
            propValue = mScript.mainClass(ch.classId).Properties(propName).Value
        ElseIf child3Id < 0 Then
            propValue = mScript.mainClass(ch.classId).ChildProperties(child2Id)(propName).Value
        Else
            propValue = mScript.mainClass(ch.classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id)
        End If
        cRes = mScript.IsPropertyContainsCode(propValue)

        If cRes = MatewScript.ContainsCodeEnum.NOT_CODE Then
            frmMainEditor.codeBoxPanel.Hide()
            'Показываем help-файл
            WBhelp.Tag = Nothing
            ShowHelpFileForUsualProperty(ch, sender.Name)
        Else
            frmMainEditor.WBhelp.Hide()
            With frmMainEditor.codeBox
                .Tag = Nothing
                .Text = ""
                .codeBox.IsTextBlockByDefault = (cRes = MatewScript.ContainsCodeEnum.LONG_TEXT)
                If cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                    .Text = mScript.PrepareStringToPrint(propValue, Nothing, False)
                Else
                    .codeBox.LoadCodeFromProperty(propValue)
                End If
                .Tag = sender
                .codeBox.Refresh()
            End With
            ignoreNextEvent = True 'отмена Validating текстбокса при смене фокуса
            frmMainEditor.trakingEventState = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT
            frmMainEditor.codeBoxPanel.Show()
            frmMainEditor.codeBox.codeBox.Focus()
        End If

    End Sub

    ''' <summary> Показываем help-файл для обычных свойств (не событий) </summary>
    ''' <param name="ch">Панель-родитель свойства</param> 
    ''' <param name="propName">Имя свойства</param>
    Private Overloads Sub ShowHelpFileForUsualProperty(ByRef ch As clsChildPanel, ByVal propName As String)
        Dim hFile As String = ""
        Dim p As MatewScript.PropertiesInfoType = Nothing
        If mScript.mainClass(ch.classId).Properties.TryGetValue(propName, p) = False Then
            If mScript.mainClass(ch.classId).Functions.TryGetValue(propName, p) = False Then Return
        End If
        hFile = GetHelpPath(p.helpFile)
        If String.IsNullOrEmpty(hFile) Then
            WBhelp.Visible = False
        Else
            WBhelp.Navigate(hFile)
            WBhelp.Visible = True
        End If
    End Sub

    ''' <summary> Показываем help-файл для обычных свойств (не событий) </summary>
    ''' <param name="classId">Id класса</param> 
    ''' <param name="propName">Имя свойства</param>
    Private Overloads Sub ShowHelpFileForUsualProperty(ByVal classId As Integer, ByVal propName As String)
        Dim hFile As String = ""
        Dim p As MatewScript.PropertiesInfoType = Nothing
        If mScript.mainClass(classId).Properties.TryGetValue(propName, p) = False Then
            If mScript.mainClass(classId).Functions.TryGetValue(propName, p) = False Then Return
        End If
        hFile = GetHelpPath(p.helpFile)
        If String.IsNullOrEmpty(hFile) Then
            WBhelp.Visible = False
        Else
            WBhelp.Navigate(hFile)
            WBhelp.Visible = True
        End If
    End Sub

    ''' <summary>Делегат для событий TextChanged панели свойств</summary>
    Private Sub del_propertyTextChanged(sender As Object, e As EventArgs)
        If Not sender.AllowEvents Then Return
        If IsNothing(sender.childPanel) Then Return
        Dim ch As clsChildPanel = sender.childPanel
        Dim classId As Integer = ch.classId
        Dim child2Id As Integer = ch.GetChild2Id
        Dim child3Id As Integer = ch.GetChild3Id(child2Id)
        Dim propName As String = sender.Name
        'Получаем отформатированный в нужном для движка виде текст
        Dim sText As String = sender.Text
        If sText = My.Resources.script OrElse sText = My.Resources.longText Then
            'это свойство заменено скриптом или длинным текстом - ничего менять сейчас не надо (иначе текст скрипта заменится словом '[скрипт]')
            Return
        End If
        If (sText.Length = 0 OrElse sText = "True" OrElse sText = "False" OrElse IsNumeric(sText.Replace(".", ","))) = False Then
            If (sText.Length >= 2 AndAlso sText.Chars(0) = "'"c AndAlso sText.Chars(sText.Length - 1) = "'"c) = False Then
                sText = "'" + sText + "'"
            End If
        End If
        'Сохраняем результат в mScript.mainClass
        SetPropertyValue(classId, propName, sText, child2Id, child3Id)
    End Sub

    ''' <summary>Делегат для событий Validating панели свойств</summary>
    Private Sub del_propertyValidating(sender As Object, e As System.ComponentModel.CancelEventArgs)
        If Not sender.AllowEvents Then Return
        If ignoreNextEvent Then
            ignoreNextEvent = False
            Return
        End If
        If IsNothing(sender.childPanel) Then Return
        If sender.Text = My.Resources.script Then Return
        If sender.Text = My.Resources.longText Then Return

        Dim ch As clsChildPanel = sender.childPanel
        Dim classId As Integer = ch.classId
        Dim child2Id As Integer = ch.GetChild2Id
        Dim child3Id As Integer = ch.GetChild3Id(child2Id)
        Dim propName As String = sender.Name
        'Получаем отформатированный в нужном для движка виде текст
        Dim propValue As String = sender.Text
        If (propValue.Length = 0 OrElse propValue = "True" OrElse propValue = "False" OrElse IsNumeric(propValue.Replace(".", ","))) = False Then
            propValue = WrapString(propValue)
        End If
        If sender.ForeColor = Color.Red Then sender.ForeColor = System.Drawing.Color.FromKnownColor(KnownColor.WindowText)
        'Сохраняем результат в mScript.mainClass
        mScript.LAST_ERROR = ""
        SetPropertyValue(classId, propName, propValue, child2Id, child3Id)
        If propName.EndsWith("Total") AndAlso mScript.mainClass(classId).Properties.ContainsKey(Left(propName, propName.Length - 5)) Then
            'Существует пара свойств xxx & xxxTotal
            Dim propDefName As String = Left(propName, propName.Length - 5)
            Dim mPanel As Control = sender.Parent
            Dim arrCont() As Control = mPanel.Controls.Find(propDefName, True)
            If IsNothing(arrCont) = False AndAlso arrCont.Count > 0 Then
                If child2Id = -1 Then
                    arrCont(0).Text = mScript.PrepareStringToPrint(mScript.mainClass(classId).Properties(propDefName).Value, Nothing, False)
                ElseIf child3Id = -1 Then
                    arrCont(0).Text = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propDefName).Value, Nothing, False)
                Else
                    arrCont(0).Text = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propDefName).ThirdLevelProperties(child3Id), Nothing, False)
                End If
            End If
        End If

        If mScript.IsPropertyContainsCode(propValue) = MatewScript.ContainsCodeEnum.NOT_CODE Then Return
        'здесь только если в свойстве код (или исполняемая строка). При этом, если код с ошибкой то в SetPropertyValue установилась LAST_ERROR
        If mScript.LAST_ERROR.Length > 0 Then
            e.Cancel = True
            sender.ForeColor = Color.Red
            mScript.LAST_ERROR = ""
            Return
        End If
        'меняем текст в контроле, ставя правильные пробелы и регистры (например "?1+1" превращается в "?1 + 1")
        If child2Id = -1 Then
            sender.Text = mScript.PrepareStringToPrint(mScript.mainClass(classId).Properties(propName).Value, Nothing, False)
        ElseIf child3Id = -1 Then
            sender.Text = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False)
        Else
            sender.Text = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False)
        End If
    End Sub

    ''' <summary>Процедура-делегат, которая запускается при изменении внутреннего размера Panel1 контейнера splitProperies (sender = splitProperties.Panel1)</summary>
    Public Sub del_splitProperies_ClientSizeChanged(sender As Object, e As EventArgs)
        If questEnvironment.DisableEventsDuringLoading = False Then Return
        If questEnvironment.EnabledEvents = False OrElse String.IsNullOrEmpty(currentClassName) Then Return
        Dim classId As Integer
        If currentClassName = "Variable" Then
            classId = -2
        ElseIf currentClassName = "Function" Then
            classId = -3
        Else
            classId = mScript.mainClassHash(currentClassName)
        End If

        If dictDefContainers.ContainsKey(classId) = False Then Return
        Dim chkPropConfig As CheckBox = dictDefContainers(classId).chkConfig  '= sender.Controls("chkPropConfig")
        chkPropConfig.Left = sender.ClientSize.Width - chkPropConfig.Width - questEnvironment.defPaddingLeft

        For i As Integer = 0 To sender.Controls.Count - 1
            Dim c As Object = sender.Controls(i)
            If c.GetType.Name = "TextBoxEx" OrElse c.GetType.Name = "ComboBoxEx" Then
                If chkPropConfig.Checked Then
                    c.Width = sender.ClientSize.Width - c.Left - questEnvironment.defPaddingLeft * 2 - questEnvironment.btnConfigSize.Width
                Else
                    c.Width = sender.ClientSize.Width - c.Left - questEnvironment.defPaddingLeft
                End If
            ElseIf c.GetType.Name = "ButtonEx" Then
                If c.Name.EndsWith("Help") Then Continue For
                If c.IsConfigButton Then
                    c.Left = sender.ClientSize.Width - c.Width - questEnvironment.defPaddingLeft
                Else
                    If IsNothing(c.ButtonConfig) = False AndAlso chkPropConfig.Checked Then
                        c.Width = sender.ClientSize.Width - c.Left - questEnvironment.defPaddingLeft * 2 - questEnvironment.btnConfigSize.Width
                    Else
                        c.Width = sender.ClientSize.Width - c.Left - questEnvironment.defPaddingLeft
                    End If
                End If
            ElseIf c.GetType.Name = "ShapeContainer" Then
                Dim shCont As Microsoft.VisualBasic.PowerPacks.ShapeContainer = c
                For j As Integer = 0 To shCont.Shapes.Count - 1
                    shCont.Shapes(j).x2 = sender.ClientSize.Width - shCont.Shapes(j).x1 - questEnvironment.defPaddingLeft
                Next j
            ElseIf c.GetType.Name = "TreeView" Then
                c.Width = sender.ClientSize.Width - c.Left - questEnvironment.defPaddingLeft
                c.Height = sender.Parent.SplitterDistance - c.Top - questEnvironment.defPaddingLeft

            End If
        Next

        'ElementName
        Dim lbl() As Control = sender.Controls.Find("LabelSubInfo", True)
        If IsNothing(lbl) = False AndAlso lbl.Count > 0 Then lbl(0).MaximumSize = New Size(sender.ClientSize.Width - lbl(0).Left - questEnvironment.defPaddingLeft, lbl(0).Height)

        sender.Parent.Parent.CurrentWidth = sender.Width ' sender.ClientSize.Width
        sender.Parent.Parent.CurrentInfoPanelHeight = sender.Parent.ClientSize.Height - sender.Parent.SplitterDistance
    End Sub

    Sub del_chkPropConfig_CheckedChanged(sender As Object, e As EventArgs)
        Dim chk As Boolean = sender.Checked
        Dim parent As Control = sender.Parent
        For i As Integer = 0 To parent.Controls.Count - 1
            Dim tName As String = parent.Controls(i).GetType.Name
            If tName = "ButtonEx" OrElse tName = "TextBoxEx" OrElse tName = "ComboBoxEx" Then
                Dim b As Object = parent.Controls(i)
                If b.hasConfigButton AndAlso IsNothing(b.ButtonConfig) Then CreateConfigButtons()
                If b.Visible = False OrElse IsNothing(b.ButtonConfig) Then Continue For
                b.ButtonConfig.Visible = chk
            End If
        Next
        del_splitProperies_ClientSizeChanged(parent, New EventArgs)
    End Sub

    Private Sub del_PropertyMouseEnter(sender As Object, e As EventArgs)
        Dim ch As clsPanelManager.clsChildPanel = sender.childPanel
        Dim sName As String = sender.Name
        If sender.GetType.Name = "ButtonEx" AndAlso sender.Name.ToString.EndsWith("Help") Then
            sName = Left(sender.Name, sender.Name.ToString.Length - 4)
        End If
        Dim pValue As MatewScript.PropertiesInfoType = mScript.mainClass(ch.classId).Properties(sName)

        'If mScript.mainClass(ch.classId).Properties.TryGetValue(sName, pValue) = False Then
        '    Dim a As Integer = 0
        '    Return
        'End If
        Dim desc As String = GetPropertyDescription(pValue.Description)
        '
        Dim c As Control = dictDefContainers(ch.classId).Controls.Find("lblPropertyDescription", True)(0)
        If String.IsNullOrWhiteSpace(desc) Then
            If pValue.UserAdded Then
                c.Text = "(Свойство определено Писателем)"
            Else
                c.Text = ""
            End If
        Else
            c.Text = desc
        End If
    End Sub

    Private Sub del_FunctionMouseEnter(sender As Object, e As EventArgs)
        Dim ch As clsPanelManager.clsChildPanel = sender.childPanel
        Dim sName As String = sender.Name
        If sender.GetType.Name = "ButtonEx" AndAlso sender.Name.ToString.EndsWith("Help") Then
            sName = Left(sender.Name, sender.Name.ToString.Length - 4)
        End If
        Dim pValue As MatewScript.PropertiesInfoType = mScript.mainClass(ch.classId).Functions(sName)
        Dim desc As String = GetFunctionDescription(pValue.Description)
        '
        Dim c As Control = dictDefContainers(ch.classId).Controls.Find("lblPropertyDescription", True)(0)
        If String.IsNullOrWhiteSpace(desc) Then
            If pValue.UserAdded Then
                c.Text = "(Функция определена Писателем)"
            Else
                c.Text = ""
            End If
        Else
            c.Text = desc
        End If
    End Sub

    Private Sub del_BtnAddProperty_Click(sender As Object, e As EventArgs)
        'Создание нового свойства
        Dim classId As Integer = mScript.mainClassHash(currentClassName)
        dlgProperty.PrepareData(classId, True)
        Dim dResult As DialogResult = dlgProperty.ShowDialog(frmMainEditor)
        If dResult = DialogResult.Cancel Then Return
        Dim res As String = AddUserPropertyByPropData(classId, dlgProperty.newPropertyName, dlgProperty.newPropertyData)
        If res = "#Error" Then Return
        If dlgProperty.newPropertyData.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
         dlgProperty.newPropertyData.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL Then
            If currentClassName = "A" Then actionsRouter.AddProperty(dlgProperty.newPropertyName)
            removedObjects.AddProperty(classId, dlgProperty.newPropertyName)
            Return
        End If
        'Деактивируем и удаляем текущую панель (сейчас активна панель свойств по умолчанию)
        ActivePanel.IsCodeBoxVisible = False
        ActivePanel.IsWbVisible = False
        Dim pnl As clsChildPanel = ActivePanel
        ActivePanel = Nothing
        Dim showAllFunctions As Boolean = dictDefContainers(classId).showBasicFunctions
        If frmMainEditor.Controls.Contains(iconMenuElements) = False Then frmMainEditor.Controls.Add(iconMenuElements)
        If frmMainEditor.Controls.Contains(iconMenuGroups) = False Then frmMainEditor.Controls.Add(iconMenuGroups)
        dictDefContainers(classId).Dispose()
        dictDefContainers.Remove(classId)
        'Создаем новую панель
        If currentClassName = "A" Then actionsRouter.AddProperty(dlgProperty.newPropertyName)
        removedObjects.AddProperty(classId, dlgProperty.newPropertyName)
        CreatePropertiesControl(classId, showAllFunctions)
        pnl.lcActiveControl = dictDefContainers(classId).Controls.Find(dlgProperty.newPropertyName, True)(0)
        If mScript.mainClass(classId).LevelsCount > 0 Then
            frmMainEditor.btnShowSettings_Click(btnDefSettings, New EventArgs)
        Else
            OpenPanel(Nothing, classId, "")
            If dictDefContainers(classId).Controls(0).GetType.Name = "SplitContainer" Then
                Dim sp As SplitContainer = dictDefContainers(classId).Controls(0)
                del_splitProperies_ClientSizeChanged(sp.Panel1, New EventArgs)
            End If
        End If
        mScript.FillFuncAndPropHash()
        pnl.lcActiveControl.Focus()
    End Sub

    Private Sub del_BtnAddFunction_Click(sender As Object, e As EventArgs)
        'Создание новой функции
        Dim classId As Integer = mScript.mainClassHash(currentClassName)
        dlgFunction.PrepareData(classId, True)
        Dim dResult As DialogResult = dlgFunction.ShowDialog(frmMainEditor)
        If dResult = DialogResult.Cancel Then Return
        'Запускаем соответствующую функцию движка
        Dim res As String = AddUserFunctionByPropData(classId, dlgFunction.newFunctionName, dlgFunction.newFunctionData)
        If res = "#Error" Then Return
        If dlgFunction.newFunctionData.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
            dlgFunction.newFunctionData.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL Then Return

        'Деактивируем и удаляем текущую панель (сейчас активна панель свойств по умолчанию)
        ActivePanel.IsCodeBoxVisible = False
        ActivePanel.IsWbVisible = False
        Dim pnl As clsChildPanel = ActivePanel
        ActivePanel = Nothing
        Dim showAllFunctions As Boolean = dictDefContainers(classId).showBasicFunctions
        If frmMainEditor.Controls.Contains(iconMenuElements) = False Then frmMainEditor.Controls.Add(iconMenuElements)
        If frmMainEditor.Controls.Contains(iconMenuGroups) = False Then frmMainEditor.Controls.Add(iconMenuGroups)
        dictDefContainers(classId).Dispose()
        dictDefContainers.Remove(classId)
        'Создаем новую панель
        CreatePropertiesControl(classId, showAllFunctions)
        pnl.lcActiveControl = dictDefContainers(classId).Controls.Find(dlgFunction.newFunctionName, True)(0)
        If mScript.mainClass(classId).LevelsCount > 0 Then
            frmMainEditor.btnShowSettings_Click(btnDefSettings, New EventArgs)
        Else
            OpenPanel(Nothing, classId, "")
            If dictDefContainers(classId).Controls(0).GetType.Name = "SplitContainer" Then
                Dim sp As SplitContainer = dictDefContainers(classId).Controls(0)
                del_splitProperies_ClientSizeChanged(sp.Panel1, New EventArgs)
            End If
        End If
        mScript.FillFuncAndPropHash()
        pnl.lcActiveControl.Focus()
    End Sub

    Private Sub del_EditUserClass_Click(sender As Object, e As EventArgs)
        Dim strNames As String
        Dim levels As Integer = 3
        Dim img As Image
        Dim classId As Integer = mScript.mainClassHash(currentClassName)
        Dim oldClassName As String = currentClassName
        Dim helpFile As String
        Dim defProperty As String
        With dlgNewClass
            .PrepareForEditClass(classId)
            Dim res As DialogResult = dlgNewClass.ShowDialog(frmMainEditor)
            If res = Windows.Forms.DialogResult.Cancel Then Return
            If res = DialogResult.Abort Then Return 'Удаление класса уже выполнено в диалоге
            strNames = .txtNewClassName.Text.Trim
            levels = .nudClassLevels.Value
            img = .classIcon
            helpFile = .txtClassHelpFile.Text
            defProperty = .cmbClassDefProperty.Text
        End With

        If levels - 1 < mScript.mainClass(classId).LevelsCount Then
            Dim res As DialogResult = MessageBox.Show("Класс из " + (mScript.mainClass(classId).LevelsCount + 1).ToString + "-уровневого станет " + levels.ToString + "-уровневым, что может привести к необратимой потере элементов класса. Продолжить?", "Matew Quest", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If res = DialogResult.No Then Return
        End If
        'Редактируем класс
        'Заполняем arrNames новыми именами класса
        Dim arrNames() As String = Split(strNames, ",")
        For i As Integer = 0 To arrNames.GetUpperBound(0)
            arrNames(i) = arrNames(i).Trim
            If arrNames(i).Length = 0 Then
                MsgBox("Ошибка при заполнении имен класса. Вероятно, лишняя запятая.", MsgBoxStyle.Exclamation)
                Return
            End If
            'If mScript.mainClassHash.ContainsKey(arrNames(i)) Then
            '    MsgBox("Имя " + arrNames(i) + " уже существует.", MsgBoxStyle.Exclamation)
            '    Exit Sub
            'End If
        Next
        'Переходим из изменяемого класса в другое место
        frmMainEditor.ChangeCurrentClass("Q")
        'Заменяем старые имена класса на новые в коде
        GlobalSeeker.ReplaceClassNameInStruct(classId, arrNames)
        'Удаляем старый файл картинки
        Dim aPath As String = questEnvironment.QuestPath + "\img\classMenu\" + mScript.mainClass(classId).Names(0) + ".png"
        If My.Computer.FileSystem.FileExists(aPath) Then
            Try
                My.Computer.FileSystem.DeleteFile(aPath)
            Catch ex As Exception
                MessageBox.Show("Не удалось удалить старый файл иконки класса " + aPath + ". Попытайтесь это сделать вручную.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Try
        End If
        'Сохраняем новый файл картинки в папке квеста
        If IsNothing(img) = False Then
            aPath = questEnvironment.QuestPath + "\img\classMenu"
            If Not My.Computer.FileSystem.DirectoryExists(aPath) Then
                My.Computer.FileSystem.CreateDirectory(aPath)
            End If
            Try
                img.Save(aPath + "\" + arrNames(0) + ".png", System.Drawing.Imaging.ImageFormat.Png)
            Catch ex As Exception
                MessageBox.Show("Не удалось изменить файл иконки класса " + aPath + "\" + arrNames(0) + ".png" + ". Попытайтесь это сделать вручную.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Try
        End If
        'изменяем иконку кнопки
        If IsNothing(img) Then img = Image.FromFile(Application.StartupPath + "\src\img\classMenu\U.png")
        For i As Integer = 0 To frmMainEditor.ToolStripMain.Items.Count - 1
            If frmMainEditor.ToolStripMain.Items(i).Tag = oldClassName Then
                With frmMainEditor.ToolStripMain.Items(i)
                    .Tag = arrNames(0)
                    .Text = arrNames.Last
                    .Image = img
                    Exit For
                End With
            End If
        Next

        If arrNames(0) <> oldClassName Then
            'Переименовываем класс в группах
            Dim g As List(Of clsGroups.clsGroupInfo) = cGroups.dictGroups(oldClassName)
            cGroups.dictGroups.Remove(oldClassName)
            cGroups.dictGroups.Add(arrNames(0), g)
            If cGroups.dictRemoved.ContainsKey(oldClassName) Then
                g = cGroups.dictRemoved(oldClassName)
                cGroups.dictRemoved.Remove(oldClassName)
                cGroups.dictRemoved.Add(arrNames(0), g)
            End If
            'переименовываем в событиях изменения свойств
            mScript.trackingProperties.RenameClass(oldClassName, arrNames(0))            
        End If
        mScript.UpdateBasicFunctionsParamsWhichIsElements(classId, oldClassName)

        'сохраняем имена класса в структуре mainClass
        'ReDim Preserve mScript.mainClass(classId)
        mScript.mainClass(classId).Names = arrNames
        mScript.mainClass(classId).HelpFile = helpFile
        mScript.mainClass(classId).DefaultProperty = defProperty
        mScript.mainClass(classId).UserAdded = True

        'Выводим меню иконок из дерева или поддерева, чтобы они случайно не удалились
        If frmMainEditor.Controls.Contains(iconMenuElements) = False Then frmMainEditor.Controls.Add(iconMenuElements)
        If frmMainEditor.Controls.Contains(iconMenuGroups) = False Then frmMainEditor.Controls.Add(iconMenuGroups)
        'Удаляем контейнер
        If dictDefContainers.ContainsKey(classId) Then
            dictDefContainers(classId).Dispose()
            dictDefContainers.Remove(classId)
        End If

        'заполняем данные о новом классе. Классам 2 и 3 уровня добавляем набор базовых функций
        Dim updateTree As Boolean = False
        If mScript.mainClass(classId).LevelsCount <> levels - 1 Then
            'уровень класса изменился
            If mScript.mainClass(classId).LevelsCount = 0 Then
                'Уровень класса с первого стал 2 или 3
                'Создаем дерево
                mScript.mainClass(classId).LevelsCount = levels - 1
                frmMainEditor.AppendTree(classId)
                'заполняем класса бызовыми свойствами и функциями
                If levels = 2 Then
                    'уровень стал 2
                    For i As Integer = 0 To mScript.basicFunctionsHashLevel2.Count - 1
                        Dim func As MatewScript.PropertiesInfoType = mScript.basicFunctionsHashLevel2.ElementAt(i).Value.Clone
                        Dim funcName As String = mScript.basicFunctionsHashLevel2.ElementAt(i).Key
                        If mScript.mainClass(classId).Functions.ContainsKey(funcName) = False Then
                            func.editorIndex = mScript.mainClass(classId).Functions.Count
                            mScript.mainClass(classId).Functions.Add(funcName, func)
                        End If
                    Next
                    If mScript.mainClass(classId).Properties.ContainsKey("Name") = False Then
                        mScript.mainClass(classId).Properties.Add("Name", New MatewScript.PropertiesInfoType With {.Value = "'default'", .Description = "Имя элемента", _
                                                                  .editorIndex = mScript.mainClass(classId).Properties.Count, .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR})
                    End If
                Else
                    'уровень стал 3
                    For i As Integer = 0 To mScript.basicFunctionsHashLevel3.Count - 1
                        Dim func As MatewScript.PropertiesInfoType = mScript.basicFunctionsHashLevel3.ElementAt(i).Value.Clone
                        Dim funcName As String = mScript.basicFunctionsHashLevel3.ElementAt(i).Key
                        If mScript.mainClass(classId).Functions.ContainsKey(funcName) = False Then
                            func.editorIndex = mScript.mainClass(classId).Functions.Count
                            mScript.mainClass(classId).Functions.Add(funcName, func)
                        End If
                    Next
                    If mScript.mainClass(classId).Properties.ContainsKey("Name") = False Then
                        mScript.mainClass(classId).Properties.Add("Name", New MatewScript.PropertiesInfoType With {.Value = "'default'", .Description = "Имя элемента", _
                                                                  .editorIndex = mScript.mainClass(classId).Properties.Count, .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR})
                    End If
                End If
            ElseIf levels = 1 Then
                'уровень класса с 2 или 3 стал первым
                If frmMainEditor.dictTrees.ContainsKey(classId) Then
                    'Удаляем дерево - первому классу оно не нужно
                    frmMainEditor.dictTrees(classId).Dispose()
                    frmMainEditor.dictTrees.Remove(classId)
                End If
                'Удаляем все группы
                Dim lstGroup As List(Of clsGroups.clsGroupInfo) = cGroups.dictGroups(arrNames(0))
                If IsNothing(lstGroup) = False Then lstGroup.Clear()
                If cGroups.dictRemoved.TryGetValue(arrNames(0), lstGroup) Then
                    If IsNothing(lstGroup) = False Then lstGroup.Clear()
                End If
                'Удаляем все вкладки
                If IsNothing(lstPanels) = False Then
                    For i As Integer = cPanelManager.lstPanels.Count - 1 To 0 Step -1
                        Dim pnl As clsChildPanel = lstPanels(i)
                        If pnl.classId = classId Then
                            'lstPanels.Remove(pnl)
                            RemovePanel(pnl, False)
                        End If
                    Next i
                End If
                'очищаем mainClass
                RemoveObject(classId, Nothing)
                'Очищаем lastPanel
                If dictLastPanel.ContainsKey(classId) Then dictLastPanel.Remove(classId)
                'удаляем функции и свойства, не нужные классу 1 уровня
                mScript.mainClass(classId).LevelsCount = levels - 1
                For Each func As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In IIf(levels = 2, mScript.basicFunctionsHashLevel2, mScript.basicFunctionsHashLevel3)
                    If mScript.mainClass(classId).Functions.ContainsKey(func.Key) Then
                        mScript.mainClass(classId).Functions.Remove(func.Key)
                    End If
                Next
                UpdateFunctionsEditorIndexes(classId)
                If mScript.mainClass(classId).Properties.ContainsKey("Name") Then
                    Dim edIndex As Integer = mScript.mainClass(classId).Properties("Name").editorIndex
                    mScript.mainClass(classId).Properties.Remove("Name")
                    If IsNothing(mScript.mainClass(classId).Properties) = False Then
                        For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                            Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(i).Value
                            If p.editorIndex > edIndex Then
                                p.editorIndex -= 1
                            End If
                        Next i
                    End If
                End If
                'Добавляем функции 1 класса
                For Each func As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.basicFunctionsHashLevel1
                    If mScript.mainClass(classId).Functions.ContainsKey(func.Key) = False Then
                        mScript.mainClass(classId).Functions.Add(func.Key, func.Value.Clone)
                    End If
                Next
            ElseIf levels = 2 Then
                'Уровень класса с третьего стал вторым
                updateTree = True
                'Удаляем все группы 3-го уровня
                Dim lstGroup As List(Of clsGroups.clsGroupInfo) = cGroups.dictGroups(arrNames(0))
                If IsNothing(lstGroup) = False Then
                    For i As Integer = lstGroup.Count - 1 To 0 Step -1
                        If lstGroup(i).isThirdLevelGroup Then lstGroup.RemoveAt(i)
                    Next i
                End If
                If cGroups.dictRemoved.TryGetValue(arrNames(0), lstGroup) Then
                    For i As Integer = lstGroup.Count - 1 To 0 Step -1
                        If lstGroup(i).isThirdLevelGroup Then lstGroup.RemoveAt(i)
                    Next i
                End If
                'Удаляем все вкладки 3-го уровня. В остальных переделываем надписи
                If IsNothing(lstPanels) = False Then
                    For i As Integer = cPanelManager.lstPanels.Count - 1 To 0 Step -1
                        Dim pnl As clsChildPanel = lstPanels(i)
                        If pnl.classId = classId Then
                            If pnl.child3Name.Length > 0 Then
                                RemovePanel(pnl, False)
                            Else
                                pnl.toolButton.Text = MakeToolButtonText(pnl)
                            End If
                        End If
                    Next i
                End If
                'Удаляемм элементы 3 уровня из mainClass
                If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso IsNothing(mScript.mainClass(classId).ChildProperties(0)("Name").ThirdLevelProperties) = False Then
                    For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        For j As Integer = mScript.mainClass(classId).ChildProperties(i).Count - 1 To 0 Step -1
                            RemoveObject(classId, {i.ToString, j.ToString})
                        Next j
                    Next i
                End If
                'заменяем базовыем функции третьего класса на аналогичные им второго
                mScript.mainClass(classId).LevelsCount = levels - 1
                For Each func As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.basicFunctionsHashLevel2
                    If mScript.mainClass(classId).Functions.ContainsKey(func.Key) Then
                        mScript.mainClass(classId).Functions(func.Key) = func.Value.Clone
                    End If
                Next
            Else
                'Уровень класса со второго стал третьим
                'заменяем базовые функции второго класса на аналогичные им третьего
                mScript.mainClass(classId).LevelsCount = levels - 1
                For Each func As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.basicFunctionsHashLevel3
                    If mScript.mainClass(classId).Functions.ContainsKey(func.Key) Then
                        mScript.mainClass(classId).Functions(func.Key) = func.Value.Clone
                    End If
                Next
            End If
        End If

        mScript.MakeMainClassHash() 'обновляем хэш с именами классов
        mScript.FillFuncAndPropHash()

        frmMainEditor.ChangeCurrentClass(arrNames(0))
        If updateTree Then frmMainEditor.FillTree(arrNames(0), currentTreeView, frmMainEditor.chkShowHidden.Checked, Nothing, "")
    End Sub

    ''' <summary>
    ''' Делегат для события AfterSelect второстепенного дерева
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Public Sub del_subTree_AfterSelect(sender As Object, e As TreeViewEventArgs)
        Dim tree As TreeView = sender
        Dim btnRemove As Button = tree.Parent.Controls("RemoveSubElement")
        Dim n As TreeNode = tree.SelectedNode
        Dim cmi As ToolStripItemCollection = tree.ContextMenuStrip.Items

        If IsNothing(ActivePanel) = False Then
            ActivePanel.ActiveControl = Nothing
        End If
        Dim cbPanel As Panel = frmMainEditor.codeBoxPanel
        If IsNothing(n) Then
            btnRemove.Enabled = False
            cmi(2).Enabled = False
            cmi(3).Enabled = False
            cmi(4).Enabled = False
            cbPanel.Hide()
            Return
        End If
        cmi(4).Enabled = True

        If n.Nodes.Count > 0 Then
            btnRemove.Enabled = False
            cmi(2).Enabled = False
        Else
            btnRemove.Enabled = True
            cmi(2).Enabled = True
        End If
        If n.Tag = "GROUP" Then
            cmi(3).Enabled = False
            If frmMainEditor.codeBoxPanel.Visible Then
                frmMainEditor.codeBoxChangeOwner(Nothing)
                frmMainEditor.codeBoxPanel.Hide()
            End If
            Return
        Else
            cmi(3).Enabled = True
        End If
        'показываем кодбокс со стандартным обработчиком для данного элемента (функция GetDefaultProperty)
        Dim ni As frmMainEditor.NodeInfo = frmMainEditor.GetNodeInfo(n)
        If ni.nodeChild2Name.Length = 0 Then Return
        Dim propName As String = mScript.mainClass(ni.classId).DefaultProperty
        If String.IsNullOrEmpty(propName) Then Return
        cbPanel.Hide()

        Dim cb As CodeTextBox = frmMainEditor.codeBox
        If ni.classId = mScript.mainClassHash("Map") AndAlso Object.Equals(n, cb.Tag) = False Then
            'Отображаем редактор клетки карты
            cb.Tag = Nothing
            cb.Text = ""
            cb.Tag = n
            mapManager.BuildMapForCellsEdit(WrapString(n.Text))
            Return
        End If
        cb.Tag = Nothing
        cb.Text = ""

        If mScript.mainClass(ni.classId).Properties(propName).returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT OrElse mScript.mainClass(ni.classId).Properties(propName).returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION Then
            'Свойство по умолчанию - событие / длинный текст. Отображаем кодбокс
            cb.codeBox.IsTextBlockByDefault = (mScript.mainClass(ni.classId).Properties(propName).returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION)

            If ni.ThirdLevelNode Then
                Dim child2Id As Integer = ni.GetChild2Id
                cb.codeBox.LoadCodeFromProperty(mScript.mainClass(ni.classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(ni.GetChild3Id(child2Id)))
            Else
                cb.codeBox.LoadCodeFromProperty(mScript.mainClass(ni.classId).ChildProperties(ni.GetChild2Id)(propName).Value)
            End If
            cb.Tag = n
            cbPanel.Show()

            frmMainEditor.WBhelp.Hide()
        Else
            'Свойство - не событие и не длинный текст - отображаем файл помощи
            frmMainEditor.WBhelp.Tag = n
            ShowHelpFileForUsualProperty(ni.classId, propName)
            'frmMainEditor.WBhelp.Show()
        End If
    End Sub
#End Region

    ''' <summary>
    ''' Заполняет панель управления свойствами класса данными, соответствующие определенному объекту из mScript.mainClass
    ''' </summary>
    ''' <param name="childPanel">Панель для заполнения</param>
    Public Function FillProperties(ByRef childPanel As clsChildPanel) As Integer
        Dim chCopy As clsChildPanel = childPanel 'для предачи лямбда-членам нужна копия
        Dim classId As Integer = childPanel.classId
        If dictDefContainers.ContainsKey(classId) = False Then CreatePropertiesControl(classId)
        Dim mainPanel As PanelEx = dictDefContainers(classId)
        If Not dictDefContainers.TryGetValue(classId, mainPanel) Then
            MessageBox.Show("Ошибка при заполнении данными. Не найдена панель управления свойствами.")
            Return -1
        End If

        Dim splitCont As SplitContainer = mainPanel.Controls(0)
        splitCont.Panel1.VerticalScroll.Value = 0
        splitCont.Panel1.VerticalScroll.Value = 0
        splitCont.Panel1.HorizontalScroll.Value = 0
        splitCont.Panel1.HorizontalScroll.Value = 0

        'Пишем текст в верхней части панели (название текущего элемента)
        Dim arrC() As Control
        Dim lblElName As Label = Nothing
        arrC = mainPanel.Controls.Find("ElementName", True)
        If IsNothing(arrC) = False AndAlso arrC.Count > 0 Then
            lblElName = arrC(0)
            Dim resText As String
            If childPanel.child2Name.Length = 0 Then
                If mScript.mainClass(classId).Properties.ContainsKey("Name") Then
                    resText = mScript.mainClass(classId).Properties("Name").Value
                Else
                    resText = "'Класс " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.Count - 1) + "'"
                End If
            ElseIf childPanel.child3Name.Length = 0 Then
                If IsNothing(mScript.mainClass(classId).ChildProperties) Then
                    MessageBox.Show("Ошибка при заполнении данными. Элемент не найден (удален).")
                    Return -1
                End If
                resText = childPanel.child2Name
            Else
                resText = childPanel.child3Name
            End If
            lblElName.Text = mScript.PrepareStringToPrint(resText, Nothing, False)
        End If
        CodeTextBox.SendMessage(splitCont.Handle, CodeTextBox.WM_SetRedraw, 0, 0)
        'If mScript.mainClass(classId).Names(0) = "L" AndAlso childPanel.child2Id > -1 Then actionsRouter.RetreiveActions(childPanel.child2Id)
        Dim c As Object
        Dim controlsShifted As Boolean = False
        For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
            Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(i).Key
            Dim pValue As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(i).Value

            'перебираем все свойства класса и заполняем ими соответствующие текс/комбо боксы
            If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR Then Continue For
            arrC = mainPanel.Controls.Find(pName, True) 'ищем контрол, соответствующий свойству
            If IsNothing(arrC) OrElse arrC.Count = 0 Then Continue For

            c = arrC(0)
            Dim tName As String = c.GetType.Name
            With c
                If Array.IndexOf({"TextBoxEx", "ComboBoxEx", "ButtonEx"}, tName) = -1 Then Continue For
                .childPanel = childPanel
                If .Label.Left <> questEnvironment.defPaddingLeft Then
                    'произошло смещение всех контролов (похоже, вещь неконтролируемая, какой-то сбой самого Васика). Возвращаем все на место
                    controlsShifted = True
                    If IsNothing(.Label) = False Then
                        .Label.Left = questEnvironment.defPaddingLeft
                    End If
                    If IsNothing(.ButtonHelp) Then
                        .Left = mainPanel.maxRight
                    Else
                        .ButtonHelp.Left = mainPanel.maxRight
                        .Left = .ButtonHelp.Right + questEnvironment.defPaddingLeft
                    End If
                    If IsNothing(.ButtonConfig) = False Then
                        .ButtonConfig.Left = .Right + questEnvironment.defPaddingLeft
                    End If
                End If
                'кнопка помощи рядом с кнопкой события
                If IsNothing(.ButtonHelp) = False Then .ButtonHelp.childPanel = childPanel
                'если элемент на данной панели невидимый - пропускаем
                If .ChangeTopAndVisible(childPanel, pValue.Hidden) = False Then Continue For
            End With

            Dim child2Id As Integer = childPanel.GetChild2Id
            Dim child3Id As Integer = childPanel.GetChild3Id(child2Id)
            Dim propValue As String
            If child2Id < 0 Then
                propValue = mScript.mainClass(classId).Properties(pName).Value
            ElseIf child3Id < 0 Then
                propValue = mScript.mainClass(classId).ChildProperties(child2Id)(pName).Value
            Else
                propValue = mScript.mainClass(classId).ChildProperties(child2Id)(pName).ThirdLevelProperties(child3Id)
            End If

            If IsNothing(propValue) Then propValue = ""
            If c.GetType.Name = "ButtonEx" Then
                Dim cc As ButtonEx = c
                If String.IsNullOrEmpty(propValue) = False Then
                    cc.Text = "(заполнено)"
                    cc.ForeColor = DEFAULT_COLORS.EventButtonFilled
                Else
                    cc.Text = "(пусто)"
                    cc.ForeColor = DEFAULT_COLORS.EventButtonEmpty
                End If
            Else
                'TextBoxEx, ComboBoxEx
                c.AllowEvents = False

                Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(propValue)
                Dim propContent As String = UnWrapString(propValue)
                'If propContent = My.Resources.script Then
                '    cRes = MatewScript.ContainsCodeEnum.CODE
                'ElseIf propContent = My.Resources.longText Then
                '    cRes = MatewScript.ContainsCodeEnum.LONG_TEXT
                'End If

                If tName = "ComboBoxEx" AndAlso cRes <> MatewScript.ContainsCodeEnum.NOT_CODE Then
                    'надо превратить комбо в текстбокс
                    c = ReplaceComboWithTextBox(c)
                    tName = "TextBoxEx"
                    c.ReadOnly = True
                    If cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                        c.Text = mScript.PrepareStringToPrint(propValue, Nothing, False)
                    ElseIf cRes = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                        c.Text = My.Resources.longText
                    Else
                        c.Text = My.Resources.script
                    End If
                ElseIf tName = "TextBoxEx" AndAlso cRes = MatewScript.ContainsCodeEnum.NOT_CODE Then
                    'возможно, надо восстановить комбобокс после того, как он был превращен в текстбокс для скрипта
                    If mScript.mainClass(classId).Properties(pName).returnType = MatewScript.ReturnFunctionEnum.RETURN_ENUM Then
                        c = RestoreDefaultControl(c, False)
                    End If
                    tName = c.GetType.Name
                    If tName = "TextBoxEx" Then c.ReadOnly = False
                    c.Text = mScript.PrepareStringToPrint(propValue, Nothing, False)
                ElseIf tName = "TextBoxEx" Then
                    'в свойстве скрипт, у нас текстбокс. Надо только поставить ReadOnly
                    If Not c.ReadOnly Then c.ReadOnly = True
                    If cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                        c.Text = mScript.PrepareStringToPrint(propValue, Nothing, False)
                    ElseIf cRes = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                        c.Text = My.Resources.longText
                    Else
                        c.Text = My.Resources.script
                    End If
                Else
                    'обычное свойство, без всяких скриптов
                    c.Text = mScript.PrepareStringToPrint(propValue, Nothing, False)
                End If
                If c.GetType.Name = "ComboBoxEx" Then
                    cListManager.FillListByChildPanel(pName, childPanel, c)
                    c.SelectionLength = 0
                End If
                c.AllowEvents = True

                End If
        Next i

        ''смещаем линии
        'Dim lineTop As Integer = mainPanel.line1Top.GetTop(childPanel.child2Id, childPanel.child3Id)
        'mainPanel.line1.Y1 = lineTop
        'mainPanel.line1.Y2 = lineTop

        'lineTop = mainPanel.line2Top.GetTop(childPanel.child2Id, childPanel.child3Id)
        'mainPanel.line2.Y1 = lineTop
        'mainPanel.line2.Y2 = lineTop

        'если сейчас настройки по умолчанию, то делаем видимыми кнопки AddProperty и AddFunction
        Dim blnDefSettings As Boolean = (childPanel.child2Name.Length = 0) 'если True, то мы сейчас открываем панель настроек по умолчанию
        Dim bProp As Button = mainPanel.Controls.Find("AddProperty", True)(0)
        bProp.Visible = (childPanel.child2Name.Length = 0)
        Dim bFunc As Button = mainPanel.Controls.Find("AddFunction", True)(0)
        bFunc.Visible = (childPanel.child2Name.Length = 0)
        Dim bAllFunc As CheckBox = mainPanel.Controls.Find("chkShowAllFunctions", True)(0)
        bAllFunc.Visible = (childPanel.child2Name.Length = 0)
        If mScript.mainClass(classId).UserAdded Then
            Dim btnClassEdit As Button = mainPanel.Controls.Find("btnClassEdit", True)(0)
            btnClassEdit.Visible = (childPanel.child2Name.Length = 0)
        End If

        If controlsShifted Then
            If IsNothing(lblElName) = False Then lblElName.Left = questEnvironment.defPaddingLeft
            'mainPanel.line1.X1 = questEnvironment.defPaddingLeft
            'mainPanel.line2.X1 = questEnvironment.defPaddingLeft
            'Dim shCont As Microsoft.VisualBasic.PowerPacks.ShapeContainer = mainPanel.line1.Parent
            'shCont.Left = 0
            bProp.Left = questEnvironment.defPaddingLeft
            bFunc.Left = bProp.Right + questEnvironment.defPaddingLeft
        End If


        For Each f As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(classId).Functions
            If mainPanel.showBasicFunctions = False AndAlso f.Value.UserAdded = False Then Continue For
            arrC = mainPanel.Controls.Find(f.Key, True)
            If IsNothing(arrC) OrElse arrC.Count = 0 Then Continue For
            arrC(0).Visible = blnDefSettings
        Next

        If blnDefSettings Then
            'обработка кнопок с функциями пользователя (только если настройки по умолчанию)
            For i As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(i).Key
                Dim pValue As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(i).Value
                'перебираем все пользовательские функции класса и заполняем ими соответствующие кнопки, назначаем обработчики
                If mainPanel.showBasicFunctions = False AndAlso pValue.UserAdded = False Then Continue For
                If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR Then Continue For
                arrC = mainPanel.Controls.Find(pName, True) 'ищем контрол, соответствующий функции
                If IsNothing(arrC) OrElse arrC.Count = 0 Then Continue For

                c = arrC(0)
                With c
                    If controlsShifted Then
                        If IsNothing(.Label) = False Then
                            .Label.Left = questEnvironment.defPaddingLeft
                        End If
                        If IsNothing(.ButtonHelp) Then
                            .Left = mainPanel.maxRight
                        Else
                            .ButtonHelp.Left = mainPanel.maxRight
                            .Left = .ButtonHelp.Right + questEnvironment.defPaddingLeft
                        End If
                        If IsNothing(.ButtonConfig) = False Then
                            .ButtonConfig.Left = .Right + questEnvironment.defPaddingLeft
                        End If
                    End If

                    If .GetType.Name <> "ButtonEx" Then Continue For
                    .childPanel = childPanel
                    If .ChangeTopAndVisible(childPanel, pValue.Hidden) = False Then Continue For

                    'кнопка помощи рядом с кнопкой функции
                    If IsNothing(.ButtonHelp) = False Then .ButtonHelp.childPanel = childPanel

                    If IsNothing(pValue.Value) = False AndAlso pValue.Value.Length > 0 Then
                        .Text = "(заполнено)"
                        .ForeColor = DEFAULT_COLORS.EventButtonFilled
                    Else
                        .Text = "(пусто)"
                        .ForeColor = DEFAULT_COLORS.EventButtonEmpty
                    End If
                End With
            Next i
        End If

        'прячем LabelSubInfo в настройках по умолчанию
        If childPanel.child2Name.Length = 0 Then
            Dim arrL() As Control = mainPanel.Controls.Find("LabelSubInfo", True)
            If IsNothing(arrL) = False AndAlso arrL.Count > 0 Then
                arrL(0).Visible = False
                If controlsShifted Then arrL(0).Left = questEnvironment.defPaddingLeft
            End If
        End If

        'добавляем кнопку перехода к действиям
        Dim arrShouldBeButtonNavigate As List(Of String) = {"L", "A", "M", "Map", "Med", "Mg", "Ab", "Army"}.ToList
        If arrShouldBeButtonNavigate.IndexOf(mScript.mainClass(classId).Names(0)) > -1 OrElse (mScript.mainClass(classId).UserAdded AndAlso mScript.mainClass(classId).LevelsCount = 2) Then
            Dim btnSubClass As ButtonEx = mainPanel.NavigationButton  'кнопка перехода к дочерний элементам / родительскому элементу

            If IsNothing(btnSubClass) = False Then
                btnSubClass.Visible = (childPanel.child2Name.Length > 0)

                Dim subTree As TreeView = Nothing
                arrC = mainPanel.Controls.Find("subTree", True) 'ищем treeViev с дочерними элементами
                If IsNothing(arrC) = False AndAlso arrC.Count > 0 Then
                    subTree = arrC(0)
                    subTree.Visible = True
                End If

                Dim lblSubInfo As Label = Nothing
                arrC = mainPanel.Controls.Find("LabelSubInfo", True) 'ищем надпись вида Локация1 - Действия
                If IsNothing(arrC) = False AndAlso arrC.Count > 0 Then
                    lblSubInfo = arrC(0)
                    lblSubInfo.Visible = True
                End If

                If controlsShifted Then
                    btnSubClass.Left = questEnvironment.defPaddingLeft
                    If IsNothing(subTree) = False Then subTree.Left = questEnvironment.defPaddingLeft
                    If IsNothing(lblSubInfo) = False Then lblSubInfo.Left = questEnvironment.defPaddingLeft
                    arrC = mainPanel.Controls.Find("AddSubElement", True)
                    If IsNothing(arrC) AndAlso arrC.Count > 0 Then
                        Dim btnAdd As Button = arrC(0)
                        btnAdd.Left = questEnvironment.defPaddingLeft
                        arrC = mainPanel.Controls.Find("RemoveSubElement", True)
                        If IsNothing(arrC) AndAlso arrC.Count > 0 Then
                            arrC(0).Left = btnAdd.Right + questEnvironment.defPaddingLeft
                        End If
                    End If
                End If

                If childPanel.child2Name.Length > 0 Then
                    Select Case mScript.mainClass(classId).Names(0)
                        Case "L"
                            btnSubClass.NavigateToParentName = "" 'действия - не родители локаций
                            'If childPanel.child2Id > -1 Then RetreiveLocationActions(childPanel.child2Id)
                        Case "A"
                            btnSubClass.NavigateToParentName = currentParentName  ' ActivePanel.child2Id 'к этому моменту ActivePanel еще панель родительской локации
                        Case "M"
                            If currentParentName.Length > 0 Then
                                'из пунктов меню в блоки
                                lblSubInfo.Visible = False

                                With btnSubClass
                                    .Top = .PropertyControlsTop.TopLevel3
                                    '.Top = lblSubInfo.Top
                                    .NavigateToThirdLevel = False
                                    .Text = "Перейти к блокам меню"
                                    .Tag = "Перейти к блоку данного пункта меню"
                                    .Image = My.Resources.mmnu64
                                End With
                            Else
                                'из пунктов блоков меню в пункты
                                lblSubInfo.Visible = True
                                With btnSubClass
                                    .Top = lblSubInfo.Bottom + questEnvironment.defPaddingTop
                                    .NavigateToThirdLevel = True
                                    .Text = "Перейти к пунктам меню"
                                    .Tag = "Перейти к пунктам данного меню"
                                    .Image = My.Resources.menu_items64
                                End With
                            End If
                        Case "Map"
                            If currentParentName.Length > 0 Then
                                'из клеток к карте
                                lblSubInfo.Visible = False
                                With btnSubClass
                                    .Top = .PropertyControlsTop.TopLevel3 ' lblSubInfo.Top
                                    .NavigateToThirdLevel = False
                                    .Text = "Перейти к карте"
                                    .Tag = "Перейти к карте, которой принажлежат данные клетки"
                                    .Image = My.Resources.Map64
                                End With
                            Else
                                'из карт к клеткам
                                lblSubInfo.Visible = True
                                With btnSubClass
                                    .Top = lblSubInfo.Bottom + questEnvironment.defPaddingTop
                                    .NavigateToThirdLevel = True
                                    .Text = "Перейти к клеткам карты"
                                    .Tag = "Перейти к клеткам на данной карте"
                                    .Image = My.Resources.MapCells64
                                End With
                            End If
                        Case "Med"
                            If currentParentName.Length > 0 Then
                                'из аудиофайлов к спискам
                                lblSubInfo.Visible = False
                                With btnSubClass
                                    .Top = .PropertyControlsTop.TopLevel3
                                    '.Top = lblSubInfo.Top
                                    .NavigateToThirdLevel = False
                                    .Text = "Перейти к спискам аудио"
                                    .Tag = "Перейти к списку воспроизведения, которым принажлежат данные аудиофайлы"
                                    .Image = My.Resources.MedList64
                                End With
                            Else
                                'из списков к аудиофайлам
                                lblSubInfo.Visible = True
                                With btnSubClass
                                    .Top = lblSubInfo.Bottom + questEnvironment.defPaddingTop
                                    .NavigateToThirdLevel = True
                                    .Text = "Перейти к аудиофайлам"
                                    .Tag = "Перейти к аудиофайлам данного списка воспроизведения"
                                    .Image = My.Resources.MedAudio64
                                End With
                            End If
                        Case "Mg"
                            If currentParentName.Length > 0 Then
                                'из магий к книге
                                lblSubInfo.Visible = False
                                With btnSubClass
                                    .Top = .PropertyControlsTop.TopLevel3
                                    '.Top = lblSubInfo.Top
                                    .NavigateToThirdLevel = False
                                    .Text = "Перейти к книгам магий"
                                    .Tag = "Перейти к книге заклинаний, которой принажлежат данные магии"
                                    .Image = My.Resources.MagicBook64
                                End With
                            Else
                                'из книги к магиям
                                lblSubInfo.Visible = True
                                With btnSubClass
                                    .Top = lblSubInfo.Bottom + questEnvironment.defPaddingTop
                                    .NavigateToThirdLevel = True
                                    .Text = "Перейти к магиям"
                                    .Tag = "Перейти к магиям из данной книги заклинаний"
                                    .Image = My.Resources.Magic64
                                End With
                            End If
                        Case "Ab"
                            If currentParentName.Length > 0 Then
                                'из способностей к набору
                                lblSubInfo.Visible = False
                                With btnSubClass
                                    .Top = .PropertyControlsTop.TopLevel3
                                    '.Top = lblSubInfo.Top
                                    .NavigateToThirdLevel = False
                                    .Text = "Перейти к наборам способностей"
                                    .Tag = "Перейти к набору способностей, которому принадлежат данные способности"
                                    .Image = My.Resources.AbSet64
                                End With
                            Else
                                'из набора к способностям
                                lblSubInfo.Visible = True
                                With btnSubClass
                                    .Top = lblSubInfo.Bottom + questEnvironment.defPaddingTop
                                    .NavigateToThirdLevel = True
                                    .Text = "Перейти к способностям"
                                    .Tag = "Перейти к способностям из данного набора способностей"
                                    .Image = My.Resources.Ab64
                                End With
                            End If
                        Case "Army"
                            If currentParentName.Length > 0 Then
                                'из бойцов к армии
                                lblSubInfo.Visible = False
                                With btnSubClass
                                    .Top = .PropertyControlsTop.TopLevel3
                                    '.Top = lblSubInfo.Top
                                    .NavigateToThirdLevel = False
                                    .Text = "Перейти к армиям"
                                    .Tag = "Перейти к армии, которому принадлежат данные бойцы"
                                    .Image = My.Resources.Army
                                End With
                            Else
                                'из армии к бойцам
                                lblSubInfo.Visible = True
                                With btnSubClass
                                    .Top = lblSubInfo.Bottom + questEnvironment.defPaddingTop
                                    .NavigateToThirdLevel = True
                                    .Text = "Перейти к бойцам"
                                    .Tag = "Перейти к бойцам данной армии"
                                    .Image = My.Resources.Unit
                                End With
                            End If
                        Case Else
                            'Пользовательский класс 3 уровня
                            If currentParentName.Length > 0 Then
                                'из 3 уровня во 2
                                lblSubInfo.Visible = False

                                With btnSubClass
                                    .Top = .PropertyControlsTop.TopLevel3
                                    '.Top = lblSubInfo.Top
                                    .NavigateToThirdLevel = False
                                    .Text = "Перейти к родителю"
                                    .Tag = "Перейти к родителю данного элемента 3 уровня"
                                    .Image = My.Resources.userClass64
                                End With
                            Else
                                'из 2 в 3
                                lblSubInfo.Visible = True
                                With btnSubClass
                                    .Top = lblSubInfo.Bottom + questEnvironment.defPaddingTop
                                    .NavigateToThirdLevel = True
                                    .Text = "Перейти к дочерним"
                                    .Tag = "Перейти к дочерним элементам"
                                    .Image = My.Resources.userClass64
                                End With
                            End If
                    End Select
                    If IsNothing(subTree) = False Then
                        If childPanel.child3Name.Length = 0 Then
                            subTree.Visible = True
                            If btnSubClass.NavigateToParentName.Length = 0 AndAlso IsNothing(subTree) = False Then
                                frmMainEditor.FillTree(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), subTree, True, Nothing, childPanel.child2Name)
                            End If
                        Else
                            subTree.Visible = False
                        End If
                    End If
                Else
                    If IsNothing(subTree) = False Then subTree.Visible = False
                End If
            End If
        End If


        If IsNothing(mainPanel.Parent) Then
            Dim prevWidth As Integer = mainPanel.CurrentWidth
            Dim prevHeight As Integer = mainPanel.CurrentInfoPanelHeight
            parent.Panel1.Controls.Add(mainPanel)
            mainPanel.CurrentWidth = prevWidth
            mainPanel.CurrentInfoPanelHeight = prevHeight
        End If
        'Восстанавливаем размеры
        Dim splitProperties As SplitContainer = mainPanel.Controls(0)
        splitProperties.SplitterDistance = mainPanel.ClientSize.Height - mainPanel.CurrentInfoPanelHeight
        frmMainEditor.SplitInner.SplitterDistance = mainPanel.CurrentWidth

        CodeTextBox.SendMessage(splitCont.Handle, CodeTextBox.WM_SetRedraw, 1, 0)
        splitCont.Refresh()
        Return 0
    End Function

    ''' <summary>
    ''' Заполняет панель переенной данными, соответствующие определенному объекту из mScript.csPublicVariable
    ''' </summary>
    ''' <param name="childPanel">Панель для заполнения</param>
    Public Function FillPropertiesVariables(ByRef childPanel As clsChildPanel) As Integer
        frmMainEditor.pnlVariables.Show()
        'Заполняем DataGridView данными о переменной
        Dim dg As DataGridView = frmMainEditor.dgwVariables
        dg.Tag = Nothing
        dg.Rows.Clear()
        frmMainEditor.txtVariableDescription.Tag = Nothing
        If childPanel.child2Name.Length = 0 Then
            frmMainEditor.txtVariableDescription.Text = ""
            Return -1
        End If

        Dim vValue As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables(childPanel.child2Name)
        'Вставляем описание
        frmMainEditor.txtVariableDescription.Text = vValue.Description
        frmMainEditor.txtVariableDescription.Tag = childPanel
        'Вставляем значения переменной
        If IsNothing(vValue.arrValues) OrElse vValue.arrValues.Count = 0 Then
            dg.Tag = childPanel
            Return 0
        End If
        dg.Rows.Add(vValue.arrValues.Count)
        For i As Integer = 0 To vValue.arrValues.Count - 1
            Dim res As String = vValue.arrValues(i)
            If IsNothing(res) Then res = ""
            dg.Item(1, i).Value = mScript.PrepareStringToPrint(res, Nothing, False)
        Next
        'вставляем сигнатуры
        If IsNothing(vValue.lstSingatures) = False AndAlso vValue.lstSingatures.Count > 0 Then
            For i As Integer = 0 To vValue.lstSingatures.Count - 1
                Dim signName As String = vValue.lstSingatures.ElementAt(i).Key
                Dim signId As Integer = vValue.lstSingatures.ElementAt(i).Value
                dg.Item(0, signId).Value = signName
            Next
        End If
        dg.Tag = childPanel
        Return 0
    End Function

    ''' <summary>Создает элементы управления всеми классами (окна свойств)</summary>
    Private Sub CreatePropertiesControls()
        'убираем старые если есть
        If IsNothing(dictDefContainers) = False AndAlso dictDefContainers.Count > 0 Then
            For i As Integer = 0 To mScript.mainClass.Count - 1
                If dictDefContainers.ContainsKey(i) Then dictDefContainers(i).Dispose()
            Next
            dictDefContainers = New Dictionary(Of Integer, PanelEx)
        End If

        'CreateConfigButtonMenu()
        Dim arrSkip() As String = {"Math", "S", "D", "File", "Script", "Arr"}
        'Dim b As New System.ComponentModel.BackgroundWorker
        'AddHandler b.DoWork, Sub(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        '                         CreatePropertiesControl(0)
        '                     End Sub
        'b.RunWorkerAsync()

        For classId As Integer = 0 To mScript.mainClass.Count - 1
            Dim className As String = mScript.mainClass(classId).Names(0)
            If Array.IndexOf(arrSkip, className) > -1 Then Continue For
            CreatePropertiesControl(classId)
        Next classId
        parent.SplitterDistance = 400

    End Sub

    ''' <summary>
    ''' Создает кнопки настроек возле контролов свойств
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateConfigButtons()
        Dim mPanel As PanelEx = dictDefContainers(mScript.mainClassHash(currentClassName))
        Dim parent As Object = mPanel.Controls(0) 'splitProperties
        Dim pnl As Panel = parent.Panel1
        For i As Integer = 0 To pnl.Controls.Count - 1
            Dim c As Object = pnl.Controls(i)
            Dim tName As String = c.GetType.Name
            If tName = "TextBoxEx" OrElse tName = "ComboBoxEx" OrElse tName = "ButtonEx" Then
                If c.hasConfigButton Then
                    Dim btnConf As New ButtonEx With {.Left = .Right + questEnvironment.defPaddingLeft, .Size = questEnvironment.btnConfigSize, .Image = My.Resources.config16, .Text = "", _
                                                      .Visible = False, .IsConfigButton = True, .Name = c.Name + "Config"}
                    btnConf.Top = c.Top + (c.Height - btnConf.Height) / 2
                    AddHandler btnConf.Click, Sub(sender As Object, e As EventArgs)
                                                  questEnvironment.propertiesConfigMenu.OwnerControl = sender.Parent.Controls(sender.Name.Substring(0, sender.Name.Length - 6))
                                                  questEnvironment.propertiesConfigMenu.Show(sender, sender.Width, 0)
                                              End Sub
                    pnl.Controls.Add(btnConf)
                    c.ButtonConfig = btnConf
                End If
            End If

        Next
    End Sub

    ''' <summary>Создает элементы управления всеми классами (окна свойств)</summary>
    ''' <param name="classId">Id класса</param>
    Public Sub CreatePropertiesControl(ByVal classId As Integer, Optional showBasicFunctions As Boolean = False)
        'создаем главную панель-контейнер всех остальных элементов
        Dim mainPanel As New PanelEx With {.Dock = DockStyle.Fill}
        mainPanel.showBasicFunctions = showBasicFunctions
        '!!!A. главная панель splitProperties, надпись ElementName, чекбокс chkPropConfig
        dictDefContainers.Add(classId, mainPanel) 'сохраняем панель в словаре ключом - Id соответствующего класса

        Dim splitProperties As New SplitContainer With {.Orientation = Orientation.Horizontal, .Visible = False, .Dock = DockStyle.Fill}
        splitProperties.Panel1.AutoScroll = True
        mainPanel.Controls.Add(splitProperties)


        Dim ldesc As New Label With {.Text = "", .Dock = DockStyle.Fill, .AutoSize = False, .BackColor = Color.White}
        ldesc.Name = "lblPropertyDescription"
        splitProperties.Panel2.Controls.Add(ldesc)
        AddHandler splitProperties.Panel1.ClientSizeChanged, AddressOf del_splitProperies_ClientSizeChanged

        Dim maxRight As Integer = 0
        Dim lastBotom As Integer = questEnvironment.defPaddingTop

        'Надпись вида "Локация 1"
        Dim lblInfo As New Label With {.Text = "ElementName", .AutoSize = True, .Left = 5, .Top = 5, .BackColor = Color.Transparent, .TextAlign = ContentAlignment.MiddleLeft, .ForeColor = Color.IndianRed, _
                                   .Font = New Font(frmMainEditor.Font.FontFamily, 20, FontStyle.Italic)}
        lblInfo.Name = "ElementName"
        splitProperties.Panel1.Controls.Add(lblInfo)

        'кнопка насторойки свойства
        Dim lInfH As Integer = lblInfo.Height
        Dim chkPropConfig As New CheckBox With {.Image = My.Resources.chkPropConfig32, .Width = lInfH, .Height = lInfH, .Text = "", .Checked = False, _
                                                .Appearance = Appearance.Button, .Name = "chkPropConfig"}
        With chkPropConfig
            .Location = New Point(splitProperties.SplitterDistance - .Width - questEnvironment.defPaddingLeft, lblInfo.Top + (lInfH - .Height) / 2)
            AddHandler .CheckedChanged, AddressOf del_chkPropConfig_CheckedChanged
            AddHandler .MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = "Показать / скрыть кнопки управления содержимым свойств"
        End With
        mainPanel.chkConfig = chkPropConfig
        'frmMainEditor.ttMain.SetToolTip(chkPropConfig, "Показать/скрыть кнопки настройки свойств")
        splitProperties.Panel1.Controls.Add(chkPropConfig)

        lastBotom = lblInfo.Bottom

        '!!!B. Надписи свойств, линии
        Dim classProperties As SortedList(Of String, MatewScript.PropertiesInfoType) = mScript.mainClass(classId).Properties
        Dim propOrder() As Integer = mScript.CreatePropertiesOrderArray(classProperties)
        Dim pValue As MatewScript.PropertiesInfoType
        Dim pName As String
        For i As Integer = 0 To classProperties.Count - 1
            pValue = classProperties.ElementAt(propOrder(i)).Value
            pName = classProperties.ElementAt(propOrder(i)).Key
            'печатаем надписи для событий
            If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                (pValue.returnType <> MatewScript.ReturnFunctionEnum.RETURN_EVENT AndAlso pValue.returnType <> MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION) Then Continue For

            Dim lbl As New Label With {.Text = pValue.EditorCaption, .Left = questEnvironment.defPaddingLeft, .Top = lastBotom + questEnvironment.defPaddingTop, .AutoSize = False, .TextAlign = ContentAlignment.MiddleLeft, _
                           .BackColor = Color.Transparent, .Font = frmMainEditor.Font}
            With lbl
                .Name = "Label" + pName
                'AddHandler .MouseEnter, AddressOf del_PropertyMouseEnter 'Sub(sender As Object, e As EventArgs) ldesc.Text = pValue.Description
                splitProperties.Panel1.Controls.Add(lbl)
                .Size = New Size(.GetPreferredSize(Size.Empty).Width, questEnvironment.defPropHeight)
                maxRight = Math.Max(.Width, maxRight)
                lastBotom = .Bottom
            End With
        Next i

        ''печатаем линию-разделитель
        'Dim shContainer As New Microsoft.VisualBasic.PowerPacks.ShapeContainer
        'splitProperties.Panel1.Controls.Add(shContainer)
        'Dim lin As New Microsoft.VisualBasic.PowerPacks.LineShape(questEnvironment.defPaddingLeft, lastBotom + questEnvironment.defPaddingTop, splitProperties.Panel1.Width - questEnvironment.defPaddingLeft * 2, lastBotom + questEnvironment.defPaddingTop)
        'mainPanel.line1 = lin
        'lin.Parent = shContainer
        lastBotom += questEnvironment.defPaddingTop

        For i As Integer = 0 To classProperties.Count - 1
            pValue = classProperties.ElementAt(propOrder(i)).Value
            pName = classProperties.ElementAt(propOrder(i)).Key
            'печатаем надписи для остальных встроенных свойств
            If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                pValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT OrElse pValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION Then Continue For

            Dim lbl As New Label With {.Text = pValue.EditorCaption, .Left = questEnvironment.defPaddingLeft, .Top = lastBotom + questEnvironment.defPaddingTop + 3, .AutoSize = False, .TextAlign = ContentAlignment.TopLeft,
                                       .BackColor = Color.Transparent, .Font = frmMainEditor.Font, .Name = "Label" + pName}
            With lbl
                'AddHandler .MouseEnter, AddressOf del_PropertyMouseEnter 'Sub(sender As Object, e As EventArgs) ldesc.Text = pValue.Description
                splitProperties.Panel1.Controls.Add(lbl)
                .Size = New Size(lbl.GetPreferredSize(Size.Empty).Width, questEnvironment.defPropHeight - 3)
                maxRight = Math.Max(.Width, maxRight)
                lastBotom = .Bottom
            End With
        Next i

        ''печатаем линию-разделитель
        'Dim lin2 As New Microsoft.VisualBasic.PowerPacks.LineShape(questEnvironment.defPaddingLeft, lastBotom + questEnvironment.defPaddingTop, splitProperties.Panel1.Width - questEnvironment.defPaddingLeft * 2, lastBotom + questEnvironment.defPaddingTop)
        'lin2.Parent = shContainer
        'mainPanel.line2 = lin2
        lastBotom += questEnvironment.defPaddingTop

        Dim classFunctions As SortedList(Of String, MatewScript.PropertiesInfoType) = mScript.mainClass(classId).Functions
        Dim funcOrder() As Integer = mScript.CreatePropertiesOrderArray(classFunctions)
        For i As Integer = 0 To classFunctions.Count - 1
            pValue = classFunctions.ElementAt(funcOrder(i)).Value
            pName = classFunctions.ElementAt(funcOrder(i)).Key
            If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR Then Continue For
            If showBasicFunctions = False AndAlso pValue.UserAdded = False Then Continue For
            'печатаем надписи для функций пользователя
            Dim lblText As String = pValue.EditorCaption
            If String.IsNullOrEmpty(lblText) Then lblText = pName
            Dim lbl As New Label With {.Text = lblText, .Left = questEnvironment.defPaddingLeft, .Top = lastBotom + questEnvironment.defPaddingTop + 3, .AutoSize = False, .TextAlign = ContentAlignment.TopLeft, _
                                       .Font = frmMainEditor.Font, .BackColor = Color.Transparent, .Name = "Label" + pName}
            splitProperties.Panel1.Controls.Add(lbl)
            With lbl
                'AddHandler .MouseEnter, AddressOf del_FunctionMouseEnter
                .Size = New Size(lbl.GetPreferredSize(Size.Empty).Width, questEnvironment.defPropHeight - 3)
                maxRight = Math.Max(.Width, maxRight)
                lastBotom = lbl.Bottom
            End With
        Next i
        lastBotom = lblInfo.Bottom
        maxRight += questEnvironment.defPaddingLeft
        mainPanel.maxRight = maxRight

        Dim offetByTop As New cPropertyControlsTop

        '!!!С. Кнопки, текст- и комбобоксы событий
        For i As Integer = 0 To classProperties.Count - 1
            pValue = classProperties.ElementAt(propOrder(i)).Value
            pName = classProperties.ElementAt(propOrder(i)).Key
            'вставляем кнопки для событий
            If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                (pValue.returnType <> MatewScript.ReturnFunctionEnum.RETURN_EVENT AndAlso pValue.returnType <> MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION) Then Continue For

            Dim btnHelp As New ButtonEx With {.Image = My.Resources.help26, .Left = maxRight, .Top = lastBotom + questEnvironment.defPaddingTop, .Size = New Size(questEnvironment.defPropHeight, questEnvironment.defPropHeight), _
                              .Name = pName + "Help", .ImageAlign = ContentAlignment.MiddleCenter, .TextImageRelation = TextImageRelation.ImageBeforeText, .TextAlign = ContentAlignment.MiddleRight}
            AddHandler btnHelp.MouseEnter, AddressOf del_PropertyMouseEnter 'Sub(sender As Object, e As EventArgs) ldesc.Text = pValue.Description
            AddHandler btnHelp.Click, AddressOf del_propertyButtonHelpMouseClick
            splitProperties.Panel1.Controls.Add(btnHelp)

            Dim btnCode As New ButtonEx With {.Image = My.Resources.edit26, .Left = btnHelp.Right + questEnvironment.defPaddingLeft, .Top = lastBotom + questEnvironment.defPaddingTop, .Size = New Size(36, questEnvironment.defPropHeight), _
                                             .Name = pName, .ImageAlign = ContentAlignment.TopLeft, .TextImageRelation = TextImageRelation.ImageBeforeText, .TextAlign = ContentAlignment.MiddleRight}
            With btnCode
                AddHandler .MouseEnter, AddressOf del_PropertyMouseEnter 'Sub(sender As Object, e As EventArgs) ldesc.Text = pValue.Description
                AddHandler .MouseClick, AddressOf del_propertyButtonMouseClick
                If String.IsNullOrEmpty(pValue.Value) = False Then
                    .Text = "(заполнено)"
                    .ForeColor = DEFAULT_COLORS.EventButtonFilled
                Else
                    .Text = "(пусто)"
                    .ForeColor = DEFAULT_COLORS.EventButtonEmpty
                End If
                splitProperties.Panel1.Controls.Add(btnCode)
                .PropertyControlsTop.CalculatePosition(.Top, offetByTop)
                Select Case pValue.Hidden
                    Case MatewScript.PropertyHiddenEnum.LEVEL1_ONLY
                        offetByTop.TopLevel2 += .Height + questEnvironment.defPaddingTop
                        offetByTop.TopLevel3 += .Height + questEnvironment.defPaddingTop
                    Case MatewScript.PropertyHiddenEnum.LEVEL2_ONLY
                        offetByTop.TopLevel1 += .Height + questEnvironment.defPaddingTop
                        offetByTop.TopLevel3 += .Height + questEnvironment.defPaddingTop
                    Case MatewScript.PropertyHiddenEnum.LEVEL3_ONLY
                        offetByTop.TopLevel1 += .Height + questEnvironment.defPaddingTop
                        offetByTop.TopLevel2 += .Height + questEnvironment.defPaddingTop
                    Case MatewScript.PropertyHiddenEnum.LEVEL12_ONLY
                        offetByTop.TopLevel3 += .Height + questEnvironment.defPaddingTop
                    Case MatewScript.PropertyHiddenEnum.LEVEL13_ONLY
                        offetByTop.TopLevel2 += .Height + questEnvironment.defPaddingTop
                    Case MatewScript.PropertyHiddenEnum.LEVEL23_ONLY
                        offetByTop.TopLevel1 += .Height + questEnvironment.defPaddingTop
                End Select
                .ButtonHelp = btnHelp
                .Label = .Parent.Controls("Label" + pName)
            End With

            'кнопка настройки свойств
            If pValue.UserAdded Then
                btnCode.hasConfigButton = True
            End If

            lastBotom = btnCode.Bottom
        Next i

        'mainPanel.line1Top.CalculatePosition(lin.Y1, offetByTop)
        lastBotom += questEnvironment.defPaddingTop
        For i As Integer = 0 To classProperties.Count - 1
            pValue = classProperties.ElementAt(propOrder(i)).Value
            pName = classProperties.ElementAt(propOrder(i)).Key
            'вставляем текстбоксы для свойств
            If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                pValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT OrElse pValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION Then Continue For

            Dim c As Object
            If pValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_USUAL OrElse pValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_COLOR Then
                c = New TextBoxEx With {.Text = "", .Left = maxRight, .Top = lastBotom + questEnvironment.defPaddingTop, .Size = New Size(splitProperties.Panel1.Width - maxRight - questEnvironment.defPaddingLeft * 2 - 26, questEnvironment.defPropHeight), _
                                        .Name = pName}
                Dim txtBox As TextBoxEx = c
                AddHandler txtBox.MouseEnter, AddressOf del_PropertyMouseEnter
                AddHandler txtBox.MouseDown, AddressOf del_TextPropertyMouseDown
                AddHandler txtBox.TextChanged, AddressOf del_propertyTextChanged '????
                AddHandler txtBox.Validating, AddressOf del_propertyValidating
            Else
                c = New ComboBoxEx With {.Text = "", .Left = maxRight, .Top = lastBotom + questEnvironment.defPaddingTop, .Size = _
                    New Size(splitProperties.Panel1.Width - maxRight - questEnvironment.defPaddingLeft * 2 - 26, questEnvironment.defPropHeight), .Name = pName}
                Dim cmbBox As ComboBoxEx = c
                AddHandler cmbBox.MouseEnter, AddressOf del_PropertyMouseEnter
                AddHandler cmbBox.MouseDown, AddressOf del_TextPropertyMouseDown
                AddHandler cmbBox.SelectedIndexChanged, AddressOf del_propertyTextChanged
                AddHandler cmbBox.Validating, AddressOf del_propertyValidating

                Select Case pValue.returnType
                    Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                        cmbBox.DropDownStyle = ComboBoxStyle.DropDownList
                        cmbBox.Items.AddRange({"True", "False"})
                    Case MatewScript.ReturnFunctionEnum.RETURN_ENUM
                        Dim newArr(pValue.returnArray.Count - 1) As String
                        For j As Integer = 0 To newArr.Count - 1
                            newArr(j) = mScript.PrepareStringToPrint(pValue.returnArray(j), Nothing, False)
                        Next j
                        cmbBox.Items.AddRange(newArr)
                    Case Else
                        cmbBox.DropDownStyle = ComboBoxStyle.DropDown
                End Select
            End If
            c.PropertyControlsTop.CalculatePosition(c.Top, offetByTop)
            splitProperties.Panel1.Controls.Add(c)

            'кнопка настройки свойств
            c.hasConfigButton = True
            lastBotom += questEnvironment.defPaddingTop + questEnvironment.defPropHeight
            c.ButtonHelp = Nothing
            'c.ButtonConfig = btnConf
            c.Label = c.Parent.Controls("Label" + pName)

            Select Case pValue.Hidden
                Case MatewScript.PropertyHiddenEnum.LEVEL1_ONLY
                    offetByTop.TopLevel2 += questEnvironment.defPropHeight + questEnvironment.defPaddingTop
                    offetByTop.TopLevel3 += questEnvironment.defPropHeight + questEnvironment.defPaddingTop
                Case MatewScript.PropertyHiddenEnum.LEVEL2_ONLY
                    offetByTop.TopLevel1 += questEnvironment.defPropHeight + questEnvironment.defPaddingTop
                    offetByTop.TopLevel3 += questEnvironment.defPropHeight + questEnvironment.defPaddingTop
                Case MatewScript.PropertyHiddenEnum.LEVEL3_ONLY
                    offetByTop.TopLevel1 += questEnvironment.defPropHeight + questEnvironment.defPaddingTop
                    offetByTop.TopLevel2 += questEnvironment.defPropHeight + questEnvironment.defPaddingTop
                Case MatewScript.PropertyHiddenEnum.LEVEL12_ONLY
                    offetByTop.TopLevel3 += questEnvironment.defPropHeight + questEnvironment.defPaddingTop
                Case MatewScript.PropertyHiddenEnum.LEVEL13_ONLY
                    offetByTop.TopLevel2 += questEnvironment.defPropHeight + questEnvironment.defPaddingTop
                Case MatewScript.PropertyHiddenEnum.LEVEL23_ONLY
                    offetByTop.TopLevel1 += questEnvironment.defPropHeight + questEnvironment.defPaddingTop
            End Select
        Next

        'mainPanel.line2Top.CalculatePosition(lin2.Y1, offetByTop)
        Dim lastBottomBeforeUserControls As Integer = lastBotom 'для восстановления положения по высоте перед кнопками добавить свойство/функцию и кнопками функций пользователя
        lastBotom += questEnvironment.defPaddingTop
        Dim isFuncButtons As Boolean = False

        'кнопки функций пользователя
        For i As Integer = 0 To classFunctions.Count - 1
            pValue = classFunctions.ElementAt(funcOrder(i)).Value
            pName = classFunctions.ElementAt(funcOrder(i)).Key
            If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR Then Continue For
            If showBasicFunctions = False AndAlso pValue.UserAdded = False Then Continue For

            isFuncButtons = True
            Dim btnHelp As New ButtonEx With {.Image = My.Resources.help26, .Left = maxRight, .Top = lastBotom + questEnvironment.defPaddingTop, .Size = New Size(36, questEnvironment.defPropHeight), _
                              .IsFunctionButton = True, .Name = pName + "Help", .ImageAlign = ContentAlignment.MiddleCenter, .TextImageRelation = TextImageRelation.ImageBeforeText, .TextAlign = ContentAlignment.MiddleRight}
            AddHandler btnHelp.MouseEnter, AddressOf del_FunctionMouseEnter  'Sub(sender As Object, e As EventArgs) ldesc.Text = pValue.Description
            AddHandler btnHelp.Click, AddressOf del_propertyButtonHelpMouseClick
            splitProperties.Panel1.Controls.Add(btnHelp)

            Dim btnCode As New ButtonEx With {.Image = My.Resources.edit26, .Left = btnHelp.Right + questEnvironment.defPaddingLeft, .Top = lastBotom + questEnvironment.defPaddingTop, .Size = New Size(36, questEnvironment.defPropHeight), _
                                             .IsFunctionButton = True, .Name = pName, .ImageAlign = ContentAlignment.TopLeft, .TextImageRelation = TextImageRelation.ImageBeforeText, .TextAlign = ContentAlignment.MiddleRight}

            With btnCode
                AddHandler .MouseEnter, AddressOf del_FunctionMouseEnter  'Sub(sender As Object, e As EventArgs) ldesc.Text = pValue.Description
                AddHandler .MouseClick, AddressOf del_propertyButtonMouseClick

                If IsNothing(pValue.Value) = False AndAlso pValue.Value.Length > 0 Then
                    .Text = "(заполнено)"
                    .ForeColor = DEFAULT_COLORS.EventButtonFilled
                Else
                    .Text = "(пусто)"
                    .ForeColor = DEFAULT_COLORS.EventButtonEmpty
                End If
                splitProperties.Panel1.Controls.Add(btnCode)
                .PropertyControlsTop.CalculatePosition(.Top, offetByTop)
            End With

            'кнопка настройки свойств
            If pValue.UserAdded Then btnCode.hasConfigButton = True
            lastBotom = btnCode.Bottom
            btnCode.ButtonHelp = btnHelp
            'btnCode.ButtonConfig = btnConf
            btnCode.Label = btnCode.Parent.Controls("Label" + pName)
            btnCode.Label.Top = btnCode.Top + 3 + questEnvironment.defPaddingTop
        Next i

        '!!!D. Контролы перехода: надпись lblSubInfo, кнопка Navigation, дополнительное дерево с кнопками add & remove
        'создание кнопок AddProperty и AddFunction
        Dim btnAddProperty As New Button With {.Text = "", .Image = My.Resources.addProp32, .ImageAlign = ContentAlignment.MiddleCenter, .Size = New Size(My.Resources.addProp32.Width + 6, _
                                              My.Resources.addProp32.Height + 6), .Name = "AddProperty", .Left = questEnvironment.defPaddingLeft, _
                                               .Top = lastBotom + questEnvironment.defPaddingTop - offetByTop.TopLevel1}
        AddHandler btnAddProperty.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = "Создать новое свойство"
        AddHandler btnAddProperty.Click, AddressOf del_BtnAddProperty_Click
        frmMainEditor.ttMain.SetToolTip(btnAddProperty, "Создать новое свойство")
        splitProperties.Panel1.Controls.Add(btnAddProperty)
        Dim btnAddFunction As New Button With {.Text = "", .Image = My.Resources.addFunc32, .ImageAlign = ContentAlignment.MiddleCenter, .Size = New Size(My.Resources.addFunc32.Width + 6, _
                                              My.Resources.addFunc32.Height + 6), .Name = "AddFunction", .Left = btnAddProperty.Right + questEnvironment.defPaddingLeft, _
                                               .Top = lastBotom + questEnvironment.defPaddingTop - offetByTop.TopLevel1}
        AddHandler btnAddFunction.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = "Создать новую функцию"
        AddHandler btnAddFunction.Click, AddressOf del_BtnAddFunction_Click
        frmMainEditor.ttMain.SetToolTip(btnAddFunction, "Создать новую функцию")
        splitProperties.Panel1.Controls.Add(btnAddFunction)
        lastBotom = lastBottomBeforeUserControls

        'Кнопка Редактировать пользовательский класс
        Dim nxtButtonLeft As Integer = btnAddFunction.Right
        If mScript.mainClass(classId).UserAdded Then
            Dim btnClassEdit As New Button With {.Size = btnAddFunction.Size, .Image = My.Resources.userClass32, .ImageAlign = ContentAlignment.MiddleCenter, _
                                                  .Left = btnAddFunction.Right + questEnvironment.defPaddingLeft, .Top = btnAddFunction.Top, .Name = "btnClassEdit"}
            nxtButtonLeft = btnClassEdit.Right
            splitProperties.Panel1.Controls.Add(btnClassEdit)
            AddHandler btnClassEdit.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = "Редактировать класс..."
            AddHandler btnClassEdit.Click, AddressOf del_EditUserClass_Click
        End If

        'Кнопка Загрузить / Скрыть базовые функции
        Dim chkShowAllFunctions As CheckBox = New CheckBox With {.Size = btnAddFunction.Size, .Image = My.Resources.basic_functions, .ImageAlign = ContentAlignment.MiddleCenter, _
                                                                 .Left = nxtButtonLeft + questEnvironment.defPaddingLeft, .Top = btnAddFunction.Top, .Name = "chkShowAllFunctions", _
                                                                 .Appearance = Appearance.Button}
        splitProperties.Panel1.Controls.Add(chkShowAllFunctions)
        If showBasicFunctions Then
            chkShowAllFunctions.Checked = True
            AddHandler chkShowAllFunctions.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = "Скрыть базовые функции"
        Else
            chkShowAllFunctions.Checked = False
            AddHandler chkShowAllFunctions.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = "Загрузить базовые функции"
        End If
        AddHandler chkShowAllFunctions.Click, Sub(sender As Object, e As EventArgs)
                                                  ActivePanel.IsCodeBoxVisible = False
                                                  ActivePanel.IsWbVisible = False
                                                  Dim pnl As clsChildPanel = ActivePanel
                                                  ActivePanel = Nothing
                                                  dictDefContainers(classId).Dispose()
                                                  dictDefContainers.Remove(classId)
                                                  'Создаем новую панель
                                                  CreatePropertiesControl(classId, Not showBasicFunctions)
                                                  frmMainEditor.btnShowSettings_Click(btnDefSettings, New EventArgs)
                                              End Sub

        'добавляем контролы для перехода к действиям/пунктам меню и другим дочерним элементам
        Dim arrShouldBeButtonNavigate As List(Of String) = {"L", "A", "M", "Map", "Med", "Mg", "Ab", "Army"}.ToList 'у этих классов должны быть кнопки перехода родитель/дочь
        If arrShouldBeButtonNavigate.IndexOf(mScript.mainClass(classId).Names(0)) > -1 OrElse (mScript.mainClass(classId).UserAdded AndAlso mScript.mainClass(classId).LevelsCount = 2) Then
            'Создаем надпись вида Локация1 -> Действия
            Dim lblSubInfo As Label = Nothing
            If mScript.mainClass(classId).Names(0) <> "A" Then
                lblSubInfo = New Label With {.Text = "ElementName", .AutoSize = True, .Left = questEnvironment.defPaddingLeft, .BackColor = Color.Transparent, .TextAlign = ContentAlignment.MiddleLeft, _
                                             .ForeColor = Color.IndianRed, .Font = New Font(frmMainEditor.Font.FontFamily, 18, FontStyle.Italic), _
                                             .Top = lastBotom + questEnvironment.defPaddingTop - offetByTop.TopLevel2, .AutoEllipsis = True}
                lblSubInfo.Name = "LabelSubInfo"
                splitProperties.Panel1.Controls.Add(lblSubInfo)
                lastBotom = lblSubInfo.Bottom
            End If
            'Создаем кнопку перехода родитель/дочь
            Dim btnSubClass As New ButtonEx With {.Top = lastBotom + questEnvironment.defPaddingTop, .ImageAlign = ContentAlignment.MiddleLeft, .Left = questEnvironment.defPaddingLeft, _
                                                .TextAlign = ContentAlignment.MiddleRight, .Font = New Font(frmMainEditor.Font.FontFamily, 16, FontStyle.Italic), .Name = "Navigation", _
                                                  .ForeColor = DEFAULT_COLORS.LabelHighLighted}
            'Сохраняем положение кнопки перехода между 2/3 уровнем для 3 уровня
            btnSubClass.PropertyControlsTop = New cPropertyControlsTop With {.TopLevel3 = lastBottomBeforeUserControls - offetByTop.TopLevel3 + questEnvironment.defPaddingTop}

            mainPanel.NavigationButton = btnSubClass

            'Treeview для отображения дочерних элементов должен быть не всегда, а именно только если мы в родителе
            Dim subTree As TreeView = Nothing
            Dim btnAdd As Button = Nothing, btnRemove As Button = Nothing
            'Спецнастройка отображения контролов под каждый случай + создание TreeView где надо
            Select Case mScript.mainClass(classId).Names(0)
                Case "L"
                    'Локации - должна быть кнопка перехода к действиям и дерево
                    btnSubClass.Text = "Перейти к действиям"
                    btnSubClass.Image = My.Resources.crossroads64
                    btnSubClass.NavigateToClassId = mScript.mainClassHash("A")
                    btnSubClass.NavigateToThirdLevel = False
                    AddHandler btnSubClass.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = "Перейти к действиям - вариантам поведения Игрока на этой локации"
                    AddHandler btnSubClass.Click, Sub(sender As Object, e As EventArgs)
                                                      Dim childName As String = ""
                                                      Dim n As TreeNode = subTree.SelectedNode
                                                      If IsNothing(n) = False AndAlso n.Tag = "ITEM" Then childName = WrapString(n.Text)
                                                      frmMainEditor.ChangeCurrentClass("A", False, ActivePanel.child2Name, childName)
                                                  End Sub

                    subTree = New TreeView With {.AllowDrop = True, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, .HotTracking = True, _
                         .ShowPlusMinus = True, .ShowRootLines = True, .Name = "subTree", .Left = questEnvironment.defPaddingLeft, .MinimumSize = New Size(50, 200), .ImageList = frmMainEditor.imgLstGroupIcons}
                    splitProperties.Panel1.Controls.Add(subTree)

                    btnAdd = New Button With {.Image = My.Resources.add26, .Text = "", .Width = 32, .Height = 32, .Left = questEnvironment.defPaddingLeft} ', .Name = "AddSubElement"}
                    btnRemove = New Button With {.Image = My.Resources.delete26, .Text = "", .Width = 32, .Height = 32, .Left = btnAdd.Right + questEnvironment.defPaddingLeft, .Enabled = False, .Name = "RemoveSubElement"}
                    AddHandler btnAdd.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.AddElement(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), subTree)
                    AddHandler btnRemove.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.RemoveElement(subTree.SelectedNode)
                    splitProperties.Panel1.Controls.Add(btnAdd)
                    splitProperties.Panel1.Controls.Add(btnRemove)

                    AddHandler lblInfo.TextChanged, Sub(sender As Object, e As EventArgs) lblSubInfo.Text = lblInfo.Text & " → Действия"
                Case "A"
                    'Действия - должна быть кнопка перехода к родительской локации
                    btnSubClass.Text = "Вернуться к локации"
                    btnSubClass.Image = My.Resources.Location64
                    btnSubClass.NavigateToClassId = mScript.mainClassHash("L")
                    btnSubClass.NavigateToThirdLevel = False
                    AddHandler btnSubClass.MouseEnter, Sub(sender As Object, e As EventArgs)
                                                           If currentParentName.Length = 0 Then
                                                               ldesc.Text = "Вернуться назад к локации"
                                                               Return
                                                           End If
                                                           ldesc.Text = "Вернуться назад к локации " + mScript.PrepareStringToPrint(currentParentName, Nothing, False)
                                                       End Sub

                    AddHandler btnSubClass.Click, Sub(sender As Object, e As EventArgs)
                                                      frmMainEditor.ChangeCurrentClass("L", False, "", currentParentName)
                                                  End Sub
                Case "M"
                    'Меню - должна быть кнопка перехода к пунктам меню и дерево
                    btnSubClass.Text = "Перейти к пунктам меню"
                    btnSubClass.Tag = "Перейти к пунктам данного меню"
                    btnSubClass.Image = My.Resources.menu_items64
                    btnSubClass.NavigateToClassId = mScript.mainClassHash("M")
                    btnSubClass.NavigateToThirdLevel = True
                    AddHandler btnSubClass.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = sender.Tag.ToString

                    subTree = New TreeView With {.AllowDrop = True, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, .HotTracking = True, _
                         .ShowPlusMinus = True, .ShowRootLines = True, .Name = "subTree", .Left = questEnvironment.defPaddingLeft, .MinimumSize = New Size(50, 200), .ImageList = frmMainEditor.imgLstGroupIcons}
                    splitProperties.Panel1.Controls.Add(subTree)

                    btnAdd = New Button With {.Image = My.Resources.add26, .Text = "", .Width = 32, .Height = 32, .Left = questEnvironment.defPaddingLeft, .Name = "AddSubElement"}
                    btnRemove = New Button With {.Image = My.Resources.delete26, .Text = "", .Width = 32, .Height = 32, .Left = btnAdd.Right + questEnvironment.defPaddingLeft, .Enabled = False, _
                                                 .Name = "RemoveSubElement"}
                    AddHandler btnAdd.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.AddElement(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), subTree)
                    AddHandler btnRemove.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.RemoveElement(subTree.SelectedNode)
                    splitProperties.Panel1.Controls.Add(btnAdd)
                    splitProperties.Panel1.Controls.Add(btnRemove)

                    AddHandler btnSubClass.Click, Sub(sender As Object, e As EventArgs)
                                                      If currentParentName.Length > 0 Then
                                                          'переход от пунктов меню к блокам
                                                          Dim parentName As String = ActivePanel.child2Name
                                                          frmMainEditor.ChangeCurrentClass("M", False, "", parentName)
                                                      Else
                                                          'переход от блоков меню к пунктам
                                                          Dim childName As String = ""
                                                          Dim n As TreeNode = subTree.SelectedNode
                                                          If IsNothing(n) = False Then childName = WrapString(n.Text)
                                                          frmMainEditor.ChangeCurrentClass("M", True, ActivePanel.child2Name, childName)
                                                      End If
                                                  End Sub
                    AddHandler lblInfo.TextChanged, Sub(sender As Object, e As EventArgs) lblSubInfo.Text = lblInfo.Text & " → Пункты меню"
                Case "Map"
                    'Карта - должна быть кнопка перехода и дерево
                    btnSubClass.Text = "Перейти к клеткам карты"
                    btnSubClass.Tag = "Перейти к клеткам на данной карте"
                    btnSubClass.NavigateToClassId = mScript.mainClassHash("Map")
                    btnSubClass.NavigateToThirdLevel = True
                    AddHandler btnSubClass.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = sender.Tag.ToString

                    subTree = New TreeView With {.AllowDrop = True, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, .HotTracking = True, _
                         .ShowPlusMinus = True, .ShowRootLines = True, .Name = "subTree", .Left = questEnvironment.defPaddingLeft, .MinimumSize = New Size(50, 200), .ImageList = frmMainEditor.imgLstGroupIcons}
                    splitProperties.Panel1.Controls.Add(subTree)

                    btnAdd = New Button With {.Image = My.Resources.add26, .Text = "", .Width = 32, .Height = 32, .Left = questEnvironment.defPaddingLeft, .Name = "AddSubElement"}
                    btnRemove = New Button With {.Image = My.Resources.delete26, .Text = "", .Width = 32, .Height = 32, .Left = btnAdd.Right + questEnvironment.defPaddingLeft, .Enabled = False, _
                                                 .Name = "RemoveSubElement"}
                    AddHandler btnAdd.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.AddElement(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), subTree)
                    AddHandler btnRemove.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.RemoveElement(subTree.SelectedNode)
                    splitProperties.Panel1.Controls.Add(btnAdd)
                    splitProperties.Panel1.Controls.Add(btnRemove)

                    AddHandler btnSubClass.Click, Sub(sender As Object, e As EventArgs)
                                                      If currentParentName.Length > 0 Then
                                                          'переход от клеток к карте
                                                          Dim parentName As String = ActivePanel.child2Name
                                                          frmMainEditor.ChangeCurrentClass("Map", False, "", parentName)
                                                      Else
                                                          'переход от карты к клеткам
                                                          Dim childName As String = ""
                                                          Dim n As TreeNode = subTree.SelectedNode
                                                          If IsNothing(n) = False Then childName = WrapString(n.Text)
                                                          frmMainEditor.ChangeCurrentClass("Map", True, ActivePanel.child2Name, childName)
                                                      End If
                                                  End Sub
                    AddHandler lblInfo.TextChanged, Sub(sender As Object, e As EventArgs) lblSubInfo.Text = lblInfo.Text & " → Клетки карты"
                Case "Med"
                    'Медиа - должна быть кнопка перехода и дерево
                    btnSubClass.Text = "Перейти к аудиофайлам"
                    btnSubClass.Tag = "Перейти к аудиофайлам данного списка воспроизведения"
                    btnSubClass.NavigateToClassId = mScript.mainClassHash("Med")
                    btnSubClass.NavigateToThirdLevel = True
                    AddHandler btnSubClass.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = sender.Tag.ToString

                    subTree = New TreeView With {.AllowDrop = True, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, .HotTracking = True, _
                         .ShowPlusMinus = True, .ShowRootLines = True, .Name = "subTree", .Left = questEnvironment.defPaddingLeft, .MinimumSize = New Size(50, 200), .ImageList = frmMainEditor.imgLstGroupIcons}
                    splitProperties.Panel1.Controls.Add(subTree)

                    btnAdd = New Button With {.Image = My.Resources.add26, .Text = "", .Width = 32, .Height = 32, .Left = questEnvironment.defPaddingLeft, .Name = "AddSubElement"}
                    btnRemove = New Button With {.Image = My.Resources.delete26, .Text = "", .Width = 32, .Height = 32, .Left = btnAdd.Right + questEnvironment.defPaddingLeft, .Enabled = False, _
                                                 .Name = "RemoveSubElement"}
                    AddHandler btnAdd.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.AddElement(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), subTree)
                    AddHandler btnRemove.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.RemoveElement(subTree.SelectedNode)
                    splitProperties.Panel1.Controls.Add(btnAdd)
                    splitProperties.Panel1.Controls.Add(btnRemove)

                    AddHandler btnSubClass.Click, Sub(sender As Object, e As EventArgs)
                                                      If currentParentName.Length > 0 Then
                                                          'переход от клеток к карте
                                                          Dim parentName As String = ActivePanel.child2Name
                                                          frmMainEditor.ChangeCurrentClass("Med", False, "", parentName)
                                                      Else
                                                          'переход от карты к клеткам
                                                          Dim childName As String = ""
                                                          Dim n As TreeNode = subTree.SelectedNode
                                                          If IsNothing(n) = False Then childName = WrapString(n.Text)
                                                          frmMainEditor.ChangeCurrentClass("Med", True, ActivePanel.child2Name, childName)
                                                      End If
                                                  End Sub
                    AddHandler lblInfo.TextChanged, Sub(sender As Object, e As EventArgs) lblSubInfo.Text = lblInfo.Text & " → Аудиофайлы"
                Case "Mg"
                    'Магии - должна быть кнопка перехода и дерево
                    btnSubClass.Text = "Перейти к магиям"
                    btnSubClass.Tag = "Перейти к магиям из данной книги заклинаний"
                    btnSubClass.NavigateToClassId = mScript.mainClassHash("Mg")
                    btnSubClass.NavigateToThirdLevel = True
                    AddHandler btnSubClass.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = sender.Tag.ToString

                    subTree = New TreeView With {.AllowDrop = True, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, .HotTracking = True, _
                         .ShowPlusMinus = True, .ShowRootLines = True, .Name = "subTree", .Left = questEnvironment.defPaddingLeft, .MinimumSize = New Size(50, 200), .ImageList = frmMainEditor.imgLstGroupIcons}
                    splitProperties.Panel1.Controls.Add(subTree)

                    btnAdd = New Button With {.Image = My.Resources.add26, .Text = "", .Width = 32, .Height = 32, .Left = questEnvironment.defPaddingLeft, .Name = "AddSubElement"}
                    btnRemove = New Button With {.Image = My.Resources.delete26, .Text = "", .Width = 32, .Height = 32, .Left = btnAdd.Right + questEnvironment.defPaddingLeft, .Enabled = False, _
                                                 .Name = "RemoveSubElement"}
                    AddHandler btnAdd.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.AddElement(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), subTree)
                    AddHandler btnRemove.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.RemoveElement(subTree.SelectedNode)
                    splitProperties.Panel1.Controls.Add(btnAdd)
                    splitProperties.Panel1.Controls.Add(btnRemove)

                    AddHandler btnSubClass.Click, Sub(sender As Object, e As EventArgs)
                                                      If currentParentName.Length > 0 Then
                                                          'переход от клеток к карте
                                                          Dim parentName As String = ActivePanel.child2Name
                                                          frmMainEditor.ChangeCurrentClass("Mg", False, "", parentName)
                                                      Else
                                                          'переход от карты к клеткам
                                                          Dim childName As String = ""
                                                          Dim n As TreeNode = subTree.SelectedNode
                                                          If IsNothing(n) = False Then childName = WrapString(n.Text)
                                                          frmMainEditor.ChangeCurrentClass("Mg", True, ActivePanel.child2Name, childName)
                                                      End If
                                                  End Sub
                    AddHandler lblInfo.TextChanged, Sub(sender As Object, e As EventArgs) lblSubInfo.Text = lblInfo.Text & " → Магии"
                Case "Ab"
                    'Способности - должна быть кнопка перехода и дерево
                    btnSubClass.Text = "Перейти к способностям"
                    btnSubClass.Tag = "Перейти к способностям из данного набора способностей"
                    btnSubClass.NavigateToClassId = mScript.mainClassHash("Ab")
                    btnSubClass.NavigateToThirdLevel = True
                    AddHandler btnSubClass.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = sender.Tag.ToString

                    subTree = New TreeView With {.AllowDrop = True, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, .HotTracking = True, _
                         .ShowPlusMinus = True, .ShowRootLines = True, .Name = "subTree", .Left = questEnvironment.defPaddingLeft, .MinimumSize = New Size(50, 200), .ImageList = frmMainEditor.imgLstGroupIcons}
                    splitProperties.Panel1.Controls.Add(subTree)

                    btnAdd = New Button With {.Image = My.Resources.add26, .Text = "", .Width = 32, .Height = 32, .Left = questEnvironment.defPaddingLeft, .Name = "AddSubElement"}
                    btnRemove = New Button With {.Image = My.Resources.delete26, .Text = "", .Width = 32, .Height = 32, .Left = btnAdd.Right + questEnvironment.defPaddingLeft, .Enabled = False, _
                                                 .Name = "RemoveSubElement"}
                    AddHandler btnAdd.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.AddElement(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), subTree)
                    AddHandler btnRemove.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.RemoveElement(subTree.SelectedNode)
                    splitProperties.Panel1.Controls.Add(btnAdd)
                    splitProperties.Panel1.Controls.Add(btnRemove)

                    AddHandler btnSubClass.Click, Sub(sender As Object, e As EventArgs)
                                                      If currentParentName.Length > 0 Then
                                                          'переход от клеток к карте
                                                          Dim parentName As String = ActivePanel.child2Name
                                                          frmMainEditor.ChangeCurrentClass("Ab", False, "", parentName)
                                                      Else
                                                          'переход от карты к клеткам
                                                          Dim childName As String = ""
                                                          Dim n As TreeNode = subTree.SelectedNode
                                                          If IsNothing(n) = False Then childName = WrapString(n.Text)
                                                          frmMainEditor.ChangeCurrentClass("Ab", True, ActivePanel.child2Name, childName)
                                                      End If
                                                  End Sub
                    AddHandler lblInfo.TextChanged, Sub(sender As Object, e As EventArgs) lblSubInfo.Text = lblInfo.Text & " → Способности"
                Case "Army"
                    'Способности - должна быть кнопка перехода и дерево
                    btnSubClass.Text = "Перейти к бойцам"
                    btnSubClass.Tag = "Перейти к бойцам данной армии"
                    btnSubClass.Image = My.Resources.Unit
                    btnSubClass.NavigateToClassId = mScript.mainClassHash("Army")
                    btnSubClass.NavigateToThirdLevel = True
                    AddHandler btnSubClass.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = sender.Tag.ToString

                    subTree = New TreeView With {.AllowDrop = True, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, .HotTracking = True, _
                         .ShowPlusMinus = True, .ShowRootLines = True, .Name = "subTree", .Left = questEnvironment.defPaddingLeft, .MinimumSize = New Size(50, 200), .ImageList = frmMainEditor.imgLstGroupIcons}
                    splitProperties.Panel1.Controls.Add(subTree)

                    btnAdd = New Button With {.Image = My.Resources.add26, .Text = "", .Width = 32, .Height = 32, .Left = questEnvironment.defPaddingLeft, .Name = "AddSubElement"}
                    btnRemove = New Button With {.Image = My.Resources.delete26, .Text = "", .Width = 32, .Height = 32, .Left = btnAdd.Right + questEnvironment.defPaddingLeft, .Enabled = False, _
                                                 .Name = "RemoveSubElement"}
                    AddHandler btnAdd.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.AddElement(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), subTree)
                    AddHandler btnRemove.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.RemoveElement(subTree.SelectedNode)
                    splitProperties.Panel1.Controls.Add(btnAdd)
                    splitProperties.Panel1.Controls.Add(btnRemove)

                    AddHandler btnSubClass.Click, Sub(sender As Object, e As EventArgs)
                                                      If currentParentName.Length > 0 Then
                                                          'переход от клеток к карте
                                                          Dim parentName As String = ActivePanel.child2Name
                                                          frmMainEditor.ChangeCurrentClass("Army", False, "", parentName)
                                                      Else
                                                          'переход от карты к клеткам
                                                          Dim childName As String = ""
                                                          Dim n As TreeNode = subTree.SelectedNode
                                                          If IsNothing(n) = False Then childName = WrapString(n.Text)
                                                          frmMainEditor.ChangeCurrentClass("Army", True, ActivePanel.child2Name, childName)
                                                      End If
                                                  End Sub
                    AddHandler lblInfo.TextChanged, Sub(sender As Object, e As EventArgs) lblSubInfo.Text = lblInfo.Text & " → Армии"
                Case Else
                    'Пользовательский 3-уровневый класс
                    btnSubClass.Text = "Перейти к дочерним элементам"
                    btnSubClass.Tag = "Перейти к дочерним элементам третьего уровня"
                    btnSubClass.Image = My.Resources.userClass64
                    btnSubClass.NavigateToClassId = classId
                    btnSubClass.NavigateToThirdLevel = True
                    AddHandler btnSubClass.MouseEnter, Sub(sender As Object, e As EventArgs) ldesc.Text = sender.Tag.ToString

                    subTree = New TreeView With {.AllowDrop = True, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, .HotTracking = True, _
                         .ShowPlusMinus = True, .ShowRootLines = True, .Name = "subTree", .Left = questEnvironment.defPaddingLeft, .MinimumSize = New Size(50, 200), .ImageList = frmMainEditor.imgLstGroupIcons}
                    splitProperties.Panel1.Controls.Add(subTree)

                    btnAdd = New Button With {.Image = My.Resources.add26, .Text = "", .Width = 32, .Height = 32, .Left = questEnvironment.defPaddingLeft, .Name = "AddSubElement"}
                    btnRemove = New Button With {.Image = My.Resources.delete26, .Text = "", .Width = 32, .Height = 32, .Left = btnAdd.Right + questEnvironment.defPaddingLeft, .Enabled = False, _
                                                 .Name = "RemoveSubElement"}
                    AddHandler btnAdd.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.AddElement(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), subTree)
                    AddHandler btnRemove.Click, Sub(sender As Object, e As EventArgs) Call frmMainEditor.RemoveElement(subTree.SelectedNode)
                    splitProperties.Panel1.Controls.Add(btnAdd)
                    splitProperties.Panel1.Controls.Add(btnRemove)

                    AddHandler btnSubClass.Click, Sub(sender As Object, e As EventArgs)
                                                      If currentParentName.Length > 0 Then
                                                          'переход от 3 уровня ко 2-му
                                                          Dim parentName As String = ActivePanel.child2Name
                                                          frmMainEditor.ChangeCurrentClass(mScript.mainClass(sender.NavigateToClassId).Names(0), False, "", parentName)
                                                      Else
                                                          'переход от 2 уроня к 3-му
                                                          Dim childName As String = ""
                                                          Dim n As TreeNode = subTree.SelectedNode
                                                          If IsNothing(n) = False Then childName = WrapString(n.Text)
                                                          frmMainEditor.ChangeCurrentClass(mScript.mainClass(sender.NavigateToClassId).Names(0), True, ActivePanel.child2Name, childName)
                                                      End If
                                                  End Sub
                    AddHandler lblInfo.TextChanged, Sub(sender As Object, e As EventArgs) lblSubInfo.Text = lblInfo.Text & " → дочерние элементы"
            End Select

            'подгонка размеров новой кнопки и treeView (если есть)
            btnSubClass.Size = New Size(splitProperties.Width - 10, My.Resources.crossroads64.Height + 6)
            splitProperties.Panel1.Controls.Add(btnSubClass)
            lastBotom = btnSubClass.Bottom
            If IsNothing(subTree) = False Then
                mainPanel.subTree = subTree
                btnAdd.Top = lastBotom + questEnvironment.defPaddingTop
                btnRemove.Top = lastBotom + questEnvironment.defPaddingTop
                lastBotom = btnAdd.Bottom
                subTree.Top = lastBotom + questEnvironment.defPaddingTop
                'События для treeView подклассов
                AddHandler subTree.AfterLabelEdit, Sub(sender As Object, e As NodeLabelEditEventArgs)
                                                       frmMainEditor.TreeAfterLabelEdit(sender, e, mScript.mainClass(btnSubClass.NavigateToClassId).Names(0))
                                                   End Sub
                AddHandler subTree.MouseDoubleClick, AddressOf frmMainEditor.tree_DoubleClick
                AddHandler subTree.BeforeLabelEdit, Sub(sender As Object, e As NodeLabelEditEventArgs) Call frmMainEditor.ShowIconMenu(subTree)
                AddHandler subTree.DragDrop, AddressOf frmMainEditor.tree_DragDrop
                AddHandler subTree.DragOver, AddressOf frmMainEditor.tree_DragOver
                AddHandler subTree.MouseDown, AddressOf frmMainEditor.tree_MouseDown
                AddHandler subTree.MouseMove, AddressOf frmMainEditor.tree_MouseMove
                AddHandler subTree.MouseUp, AddressOf frmMainEditor.tree_MouseUp
                AddHandler subTree.QueryContinueDrag, AddressOf frmMainEditor.tree_QueryContinueDrag
                AddHandler subTree.AfterSelect, AddressOf frmMainEditor.tree_AfterSelect
                AddHandler subTree.AfterSelect, AddressOf del_subTree_AfterSelect
                AddHandler subTree.VisibleChanged, Sub(sender As Object, e As EventArgs)
                                                       btnAdd.Visible = sender.Visible
                                                       btnRemove.Visible = sender.Visible
                                                       If IsNothing(lblSubInfo) = False Then lblSubInfo.Visible = sender.Visible
                                                   End Sub
                'Создание контекстного меню для второстепенного TreeView
                subTree.ContextMenuStrip = questEnvironment.subTreeMenu
                AddHandler subTree.MouseDown, Sub(sender As Object, e As MouseEventArgs) If subTree.Nodes.Count = 0 Then questEnvironment.subTreeMenu.Items(2).Enabled = False
            End If
        End If

        'последнии штрихи для панели splitProperties - размеры, фон, отображение
        questEnvironment.EnabledEvents = False
        'splitProperties.SplitterDistance = mainPanel.CurrentInfoPanelHeight
        'If lastBotom + questEnvironment.defPaddingLeft > splitProperties.SplitterDistance Then splitProperties.SplitterDistance = lastBotom + questEnvironment.defPaddingLeft
        splitProperties.Panel1.BackgroundImage = My.Resources.bg01
        splitProperties.Visible = True
        questEnvironment.EnabledEvents = True
        frmMainEditor.SplitInner.SplitterDistance = maxRight + 200
        mainPanel.CurrentWidth = maxRight + 200
        mainPanel.CurrentInfoPanelHeight = questEnvironment.DefaultInfoPanelHeight
    End Sub

    Private Sub CreateSubTreeMenu()
        'Создание контекстного меню для второстепенного TreeView
        Dim cm As New ContextMenuStrip
        Dim cmi(4) As ToolStripMenuItem

        For i As Integer = 0 To cmi.Count - 1
            cmi(i) = New ToolStripMenuItem
            cmi(i).Text = Choose(i + 1, "Добавить группу", "Добавить действие", "Удалить", "Дублировать", "Вывести из группы")
            cmi(i).Image = Choose(i + 1, My.Resources.add26, My.Resources.add26, My.Resources.delete26, My.Resources.duplicate26, My.Resources.ungroup26)
        Next
        AddHandler cm.Opened, Sub(sender As Object, e As EventArgs)
                                  Dim mp As PanelEx = dictDefContainers(mScript.mainClassHash(currentClassName))
                                  Dim btnSubClass As ButtonEx = mp.NavigationButton
                                  cmi(1).Text = frmMainEditor.GetTranslationAddNew(btnSubClass.NavigateToClassId, btnSubClass.NavigateToThirdLevel)
                              End Sub
        AddHandler cmi(0).Click, Sub(sender As Object, e As EventArgs)
                                     Dim mp As PanelEx = dictDefContainers(mScript.mainClassHash(currentClassName))
                                     Dim btnSubClass As ButtonEx = mp.NavigationButton
                                     Call frmMainEditor.AddGroup(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), mp.subTree)
                                 End Sub
        AddHandler cmi(1).Click, Sub(sender As Object, e As EventArgs)
                                     Dim mp As PanelEx = dictDefContainers(mScript.mainClassHash(currentClassName))
                                     Dim btnSubClass As ButtonEx = mp.NavigationButton
                                     Call frmMainEditor.AddElement(mScript.mainClass(btnSubClass.NavigateToClassId).Names(0), mp.subTree)
                                 End Sub
        AddHandler cmi(2).Click, Sub(sender As Object, e As EventArgs)
                                     Dim mp As PanelEx = dictDefContainers(mScript.mainClassHash(currentClassName))
                                     Call frmMainEditor.RemoveElement(mp.subTree.SelectedNode)
                                     If mp.subTree.Nodes.Count = 0 Then cmi(2).Enabled = False
                                 End Sub
        AddHandler cmi(3).Click, Sub(sender As Object, e As EventArgs)
                                     Dim mp As PanelEx = dictDefContainers(mScript.mainClassHash(currentClassName))
                                     Call frmMainEditor.Duplicate(mp.subTree.SelectedNode)
                                 End Sub
        AddHandler cmi(4).Click, Sub(sender As Object, e As EventArgs)
                                     Dim mp As PanelEx = dictDefContainers(mScript.mainClassHash(currentClassName))
                                     Call frmMainEditor.ExcludeFromGroup(mp.subTree)
                                 End Sub
        cmi(2).Enabled = False : cmi(3).Enabled = False : cmi(4).Enabled = False
        cm.Items.AddRange(cmi)
        questEnvironment.subTreeMenu = cm
    End Sub

#Region "Config Menu"
    ''' <summary>Создает контекстное меню для настройки свойств</summary>
    Private Sub CreateConfigButtonMenu()
        Dim cm As New ContextMenuStripEx
        'код/список/простое значение/длинный текст/значение по умолчанию/показать справку. Для пользовательских свойств/функций - удалить, редактор. 
        Dim arrText() As String = {"Сделать простым свойством", "Превратить в скрипт", "Превратить в длинный текст/html", "Показать справку", _
                                   "Удалить свойство/функцию", "Установить значение всем", "Установить значение дочерним", "Переименовать", "Редактировать", "Настроить список", _
                                   "Событие перед изменением свойства", "Событие после изменения свойства"}
        Dim arrImg() As Image = {My.Resources.return_usual, My.Resources.convetToCode, My.Resources.longText26, My.Resources.help26, My.Resources.delete26, _
                                 My.Resources.SetForAll26, My.Resources.SetForAll26, My.Resources.rename26, My.Resources.editFunc26, My.Resources.ConfigList, My.Resources.editFunc26, My.Resources.editFunc26}
        Dim cmi() As ToolStripMenuItem
        ReDim cmi(arrText.Count - 1)

        For i As Integer = 0 To arrText.Count - 1
            cmi(i) = New ToolStripMenuItem With {.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText, .Text = arrText(i), .Image = arrImg(i)}
            cm.Items.Add(cmi(i))
        Next

        ''''''''
        'Сделать простым свойством
        AddHandler cmi(0).Click, Sub(sender As Object, e As EventArgs)
                                     Dim c As Object = sender.Owner.OwnerControl
                                     Dim ch As clsChildPanel = c.childPanel
                                     Dim cb As CodeTextBox = frmMainEditor.codeBox
                                     ch.ActiveControl = c
                                     Dim child2Id As Integer = ch.GetChild2Id
                                     Dim child3Id As Integer = ch.GetChild3Id(child2Id)
                                     mScript.eventRouter.RemoveEvent(mScript.eventRouter.GetEventId(ch.classId, c.Name, child2Id, child3Id))
                                     cb.Tag = Nothing
                                     Dim newC As Object = RestoreDefaultControl(c, True)
                                     If newC.GetType.Name = "TextBoxEx" Then newC.ReadOnly = False
                                     frmMainEditor.codeBoxPanel.Hide()
                                 End Sub
        'Исполняемое свойство
        AddHandler cmi(1).Click, Sub(sender As Object, e As EventArgs)
                                     ConfigMenu_ExecutableProperty(sender.Owner.OwnerControl, False)
                                 End Sub
        'Свойство - длинный текст/html
        AddHandler cmi(2).Click, Sub(sender As Object, e As EventArgs)
                                     ConfigMenu_ExecutableProperty(sender.Owner.OwnerControl, True)
                                 End Sub
        'Показать справку
        AddHandler cmi(3).Click, Sub(sender As Object, e As EventArgs)
                                     'Отображение файла помощи о событии в браузере + код события в html
                                     Dim relatedControl As Object = sender.Owner.OwnerControl
                                     Dim propName As String = relatedControl.Name 'Имя свойства 
                                     Dim ch As clsChildPanel = relatedControl.childPanel
                                     ch.ActiveControl = Nothing
                                     Dim hFileShort As String
                                     If relatedControl.IsFunctionButton Then
                                         hFileShort = mScript.mainClass(ch.classId).Functions(propName).helpFile
                                     Else
                                         hFileShort = mScript.mainClass(ch.classId).Properties(propName).helpFile
                                     End If
                                     Dim hFile As String = GetHelpPath(hFileShort)
                                     If String.IsNullOrEmpty(hFile) Then
                                         WBhelp.Visible = False
                                         MessageBox.Show("Файл помощи для данного свойства/функции не установлен!")
                                     Else
                                         WBhelp.Navigate(hFile)
                                         frmMainEditor.HtmlShowEventText(relatedControl)
                                         WBhelp.Visible = True
                                     End If
                                 End Sub

        'Удалить свойство/функцию
        AddHandler cmi(4).Click, Sub(sender As Object, e As EventArgs)
                                     'Удаление свойства/функции пользователя
                                     Dim relatedControl As Object = sender.Owner.OwnerControl
                                     ConfigMenu_RemoveUserElement(relatedControl)
                                 End Sub

        'Установить значение всем
        AddHandler cmi(5).Click, Sub(sender As Object, e As EventArgs)
                                     'Установить текущее значение свойства всем элементам
                                     Dim relatedControl As Object = sender.Owner.OwnerControl
                                     ConfigMenu_SetValueToAll(relatedControl)
                                 End Sub

        'Установить значение дочерним
        AddHandler cmi(6).Click, Sub(sender As Object, e As EventArgs)
                                     'Установить текущее значение свойства всем элементам
                                     Dim relatedControl As Object = sender.Owner.OwnerControl
                                     ConfigMenu_SetValueToChildren(relatedControl)
                                 End Sub

        'Переименовать
        AddHandler cmi(7).Click, Sub(sender As Object, e As EventArgs)
                                     'Переименовать свойство / функцию
                                     Dim relatedControl As Object = sender.Owner.OwnerControl
                                     ConfigMenu_RenameFuncOrProp(relatedControl)
                                 End Sub
        'Редактировать
        AddHandler cmi(8).Click, Sub(sender As Object, e As EventArgs)
                                     'Переименовать свойство / функцию
                                     Dim relatedControl As Object = sender.Owner.OwnerControl
                                     If relatedControl.IsFunctionButton Then
                                         ConfigMenu_EditFunction(relatedControl)
                                     Else
                                         ConfigMenu_EditProperty(relatedControl)
                                     End If
                                 End Sub
        'Настроить список
        AddHandler cmi(9).Click, Sub(sender As Object, e As EventArgs)
                                     'Настроить список
                                     Dim relatedControl As Object = sender.Owner.OwnerControl
                                     ConfigMenu_ConfigList(relatedControl)
                                 End Sub
        'Событие перед изменением свойства
        AddHandler cmi(10).Click, Sub(sender As Object, e As EventArgs)
                                      Dim relatedControl As Object = sender.Owner.OwnerControl
                                      mScript.trackingProperties.LoadEventBeforeToCodeBox(relatedControl)
                                  End Sub
        'Событие после изменения свойства
        AddHandler cmi(11).Click, Sub(sender As Object, e As EventArgs)
                                      Dim relatedControl As Object = sender.Owner.OwnerControl
                                      mScript.trackingProperties.LoadEventAfterToCodeBox(relatedControl)
                                  End Sub

        questEnvironment.propertiesConfigMenu = cm

        AddHandler cm.VisibleChanged, Sub(sender As Object, e As EventArgs)
                                          'основная задача - сделать доступными/едоступными пункты меню в зависимости от описываемого свойства
                                          If Not sender.Visible Then Return

                                          If IsNothing(sender.OwnerControl) Then
                                              cm.Hide()
                                              Return
                                          End If

                                          Dim pnl As clsChildPanel = sender.OwnerControl.childPanel
                                          If IsNothing(pnl) Then
                                              cm.Hide()
                                              Return
                                          End If

                                          Dim propName As String = sender.OwnerControl.Name
                                          Dim isUserFunction As Boolean = sender.OwnerControl.IsFunctionButton
                                          Dim pDef As MatewScript.PropertiesInfoType
                                          If isUserFunction Then
                                              pDef = mScript.mainClass(pnl.classId).Functions(propName)
                                          Else
                                              pDef = mScript.mainClass(pnl.classId).Properties(propName)
                                          End If

                                          Dim userAdded As Boolean = pDef.UserAdded
                                          Dim strValue As String = ""
                                          Dim child2Id As Integer = pnl.GetChild2Id
                                          Dim child3Id As Integer = pnl.GetChild3Id(child2Id)
                                          If child2Id < 0 Then
                                              strValue = pDef.Value
                                          ElseIf child3Id < 0 Then
                                              Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(pnl.classId).ChildProperties(child2Id)(propName)
                                              strValue = p.Value
                                          Else
                                              Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(pnl.classId).ChildProperties(child2Id)(propName)
                                              strValue = p.ThirdLevelProperties(child3Id)
                                          End If
                                          Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(strValue)
                                          If cRes = MatewScript.ContainsCodeEnum.CODE Then
                                              cmi(0).Enabled = Not isUserFunction 'Сделать простым свойством
                                              cmi(1).Enabled = False 'Исполняемое свойство
                                              cmi(2).Enabled = Not isUserFunction  'Свойство - длинный текст/html
                                              cmi(3).Enabled = True 'Показать справку
                                          ElseIf cRes = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                                              cmi(0).Enabled = True 'Сделать простым свойством
                                              cmi(1).Enabled = False 'Исполняемое свойство
                                              cmi(2).Enabled = False 'Свойство - длинный текст/html
                                              cmi(3).Enabled = True 'Показать справку
                                          ElseIf cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                                              cmi(0).Enabled = True   'Сделать простым свойством
                                              cmi(1).Enabled = False 'Исполняемое свойство
                                              cmi(2).Enabled = True  'Свойство - длинный текст/html
                                              cmi(3).Enabled = True 'Показать справку
                                          Else
                                              cmi(0).Enabled = False  'Сделать простым свойством
                                              cmi(1).Enabled = Not isUserFunction  'Исполняемое свойство
                                              cmi(2).Enabled = Not isUserFunction 'Свойство - длинный текст/html
                                              cmi(3).Enabled = False  'Показать справку
                                          End If
                                          If pDef.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT OrElse pDef.returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION Then
                                              'пользовательское свойство - событие
                                              cmi(0).Enabled = False 'Сделать простым свойством
                                              cmi(1).Enabled = False 'Исполняемое свойство
                                              cmi(2).Enabled = False  'Свойство - длинный текст/html
                                              cmi(3).Enabled = True   'Показать справку
                                          End If
                                          If userAdded Then
                                              cmi(3).Enabled = True   'Показать справку
                                              cmi(4).Enabled = True    'Удалить свойство/функцию
                                              If isUserFunction Then
                                                  cmi(4).Text = "Удалить функцию"
                                                  cmi(8).Text = "Редактировать функцию"
                                                  cmi(8).Image = My.Resources.editFunc26
                                              Else
                                                  cmi(4).Text = "Удалить свойство"
                                                  cmi(8).Text = "Редактировать свойство"
                                                  cmi(8).Image = My.Resources.editProp26
                                              End If
                                              cmi(5).Enabled = (isUserFunction = False) AndAlso (mScript.mainClass(pnl.classId).LevelsCount > 0) 'Установить значение всем
                                              cmi(7).Enabled = True  'Переименовать
                                              cmi(8).Enabled = True  'Редактировать свойство/функцию
                                              cmi(9).Enabled = Not isUserFunction 'Настроить список
                                          Else
                                              cmi(4).Enabled = False  'Удалить свойство/функцию
                                              cmi(5).Enabled = (mScript.mainClass(pnl.classId).LevelsCount > 0)   'Установить значение всем
                                              cmi(7).Enabled = False  'Переименовать
                                              cmi(8).Enabled = False  'Редактировать свойство/функцию
                                              cmi(9).Enabled = True   'Настроить список
                                          End If
                                          If child2Id >= 0 AndAlso child3Id < 0 And mScript.mainClass(pnl.classId).LevelsCount = 2 Then
                                              If mScript.mainClass(pnl.classId).Properties(propName).Hidden <> MatewScript.PropertyHiddenEnum.LEVEL12_ONLY AndAlso _
                                                  mScript.mainClass(pnl.classId).Properties(propName).Hidden <> MatewScript.PropertyHiddenEnum.LEVEL2_ONLY Then
                                                  cmi(6).Enabled = True
                                              Else
                                                  cmi(6).Enabled = False
                                              End If
                                          Else
                                              cmi(6).Enabled = False
                                          End If
                                          If isUserFunction Then
                                              cmi(10).Enabled = False
                                              cmi(10).Checked = False
                                              cmi(11).Enabled = False
                                              cmi(11).Checked = False
                                          Else
                                              cmi(10).Enabled = True
                                              cmi(10).Checked = mScript.trackingProperties.ContainsPropertyBefore(pnl.classId, propName)
                                              cmi(11).Enabled = True
                                              cmi(11).Checked = mScript.trackingProperties.ContainsPropertyAfter(pnl.classId, propName)
                                          End If
                                      End Sub
    End Sub

    Private Sub ConfigMenu_ConfigList(ByRef relatedControl As Object)
        Dim propName As String = relatedControl.Name
        Dim ch As clsChildPanel = relatedControl.childPanel
        'Редактирование свойства
        dlgList.PrepareData(ch.classId, propName)
        Dim dResult As DialogResult = dlgList.ShowDialog(frmMainEditor)
        If dResult = DialogResult.Cancel Then Return
        'список в mainClass уже восстановлен в dlgList
        Dim tName As String = relatedControl.GetType.Name
        Dim newList() As String = dlgList.newList
        Dim isList As Boolean = (IsNothing(newList) = False AndAlso newList.Count > 0)
        If tName = "ComboBoxEx" AndAlso isList Then
            'вставляем новый список
            Dim combo As ComboBoxEx = relatedControl
            combo.Items.Clear()
            For i As Integer = 0 To newList.Count - 1
                combo.Items.Add(mScript.PrepareStringToPrint(newList(i), Nothing, False))
            Next
            combo.SelectedIndex = 0
        ElseIf tName = "TextBoxEx" AndAlso isList Then
            'заменяем текстбокс на комбо
            Dim combo As ComboBoxEx = ReplaceTextBoxWithCombo(relatedControl)
        ElseIf tName = "ComboBoxEx" AndAlso isList = False Then
            'заменяем комбо на текстбокс - список удален
            Dim retType As MatewScript.ReturnFunctionEnum = mScript.mainClass(ch.classId).Properties(relatedControl.Name).returnType
            If retType <> MatewScript.ReturnFunctionEnum.RETURN_USUAL Then
                Dim combo As ComboBoxEx = relatedControl
                'FillComboWithData(combo)
                cListManager.FillListByChildPanel(combo.Name, ch, combo)
                If combo.Items.Count > 0 Then combo.SelectedIndex = 0
            Else
                ReplaceComboWithTextBox(relatedControl)
            End If
        ElseIf tName = "TextBoxEx" AndAlso isList = False Then
            'списка не было и нет - ничего не делаем
        End If
    End Sub

    Private Sub ConfigMenu_EditFunction(ByRef relatedControl As Object)
        Dim oldName As String = relatedControl.Name
        Dim ch As clsChildPanel = relatedControl.childPanel
        'Редактирование функции
        dlgFunction.PrepareData(ch.classId, False, oldName)
        Dim dResult As DialogResult = dlgFunction.ShowDialog(frmMainEditor)
        If dResult = DialogResult.Cancel Then Return
        Dim newName As String = dlgFunction.newFunctionName
        If String.IsNullOrWhiteSpace(newName) Then Return
        Dim newData As MatewScript.PropertiesInfoType = dlgFunction.newFunctionData

        Dim p As MatewScript.PropertiesInfoType
        'переименовываем пользователькую функцию
        'вносим изменения в mainClass
        p = mScript.mainClass(ch.classId).Functions(oldName)
        Dim pValue As String = p.Value
        Dim eventId As String = p.eventId
        mScript.mainClass(ch.classId).Functions.Remove(oldName)
        newData.editorIndex = p.editorIndex
        p = newData
        p.Value = pValue
        p.eventId = eventId
        mScript.mainClass(ch.classId).Functions.Add(newName, p.Clone)

        'Деактивируем и удаляем текущую панель
        ActivePanel.IsCodeBoxVisible = False
        ActivePanel.IsWbVisible = False
        Dim pnl As clsChildPanel = ActivePanel
        ActivePanel = Nothing
        Dim showAllFunctions As Boolean = dictDefContainers(ch.classId).showBasicFunctions
        dictDefContainers(ch.classId).Dispose()
        dictDefContainers.Remove(ch.classId)
        'Вносим изменения в скрипты
        'mScript.FillFuncAndPropHash()
        If oldName <> newName Then GlobalSeeker.ReplaceElementNameInStruct(ch.classId, oldName, newName, CodeTextBox.EditWordTypeEnum.W_FUNCTION)
        'Создаем новую панель
        CreatePropertiesControl(ch.classId, showAllFunctions)
        pnl.lcActiveControl = dictDefContainers(ch.classId).Controls.Find(newName, True)(0)
        If mScript.mainClass(ch.classId).LevelsCount > 0 Then
            frmMainEditor.btnShowSettings_Click(btnDefSettings, New EventArgs)
        Else
            OpenPanel(Nothing, ch.classId, "")
            If dictDefContainers(ch.classId).Controls(0).GetType.Name = "SplitContainer" Then
                Dim sp As SplitContainer = dictDefContainers(ch.classId).Controls(0)
                del_splitProperies_ClientSizeChanged(sp.Panel1, New EventArgs)
            End If
        End If
        pnl.lcActiveControl.Focus()
    End Sub

    Private Sub ConfigMenu_EditProperty(ByRef relatedControl As Object)
        Dim oldName As String = relatedControl.Name
        Dim ch As clsChildPanel = relatedControl.childPanel
        'Редактирование свойства
        dlgProperty.PrepareData(ch.classId, False, oldName)
        Dim dResult As DialogResult = dlgProperty.ShowDialog(frmMainEditor)
        If dResult = DialogResult.Cancel Then Return
        Dim newName As String = dlgProperty.newPropertyName
        If String.IsNullOrWhiteSpace(newName) Then Return
        Dim newData As MatewScript.PropertiesInfoType = dlgProperty.newPropertyData

        Dim p As MatewScript.PropertiesInfoType

        'переименовываем пользователькое свойство
        'события изменения свойства
        mScript.trackingProperties.RenameProperty(ch.classId, oldName, newName)

        'вносим изменения в mainClass
        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(newData.Value)
        If cRes = MatewScript.ContainsCodeEnum.NOT_CODE Then
            If newData.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                cRes = MatewScript.ContainsCodeEnum.CODE
            ElseIf newData.returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION Then
                cRes = MatewScript.ContainsCodeEnum.LONG_TEXT
            End If
        End If
        'в свойства по умолчанию
        p = mScript.mainClass(ch.classId).Properties(oldName)
        Dim eventId As Integer = p.eventId
        mScript.mainClass(ch.classId).Properties.Remove(oldName)
        If IsNothing(newData.Value) Then newData.Value = ""
        newData.editorIndex = p.editorIndex
        p = newData
        mScript.mainClass(ch.classId).Properties.Add(newName, p.Clone)
        If cRes = MatewScript.ContainsCodeEnum.NOT_CODE Then
            mScript.eventRouter.RemoveEvent(eventId)
            eventId = -1
        ElseIf eventId > -1 Then
            mScript.eventRouter.SetEventId(eventId, newName, p.Value, cRes, -1, -1)
        End If

        If IsNothing(mScript.mainClass(ch.classId).ChildProperties) = False AndAlso mScript.mainClass(ch.classId).ChildProperties.Count > 0 Then
            'дочерним элементам 2 порядка
            For i As Integer = 0 To mScript.mainClass(ch.classId).ChildProperties.Count - 1
                Dim chP As MatewScript.ChildPropertiesInfoType = mScript.mainClass(ch.classId).ChildProperties(i)(oldName)
                mScript.mainClass(ch.classId).ChildProperties(i).Remove(oldName)
                chP.Value = newData.Value
                If eventId = -1 AndAlso chP.eventId > 0 Then
                    mScript.eventRouter.RemoveEvent(chP.eventId)
                    chP.eventId = -1
                ElseIf eventId > 0 Then
                    chP.eventId = mScript.eventRouter.DuplicateEvent(eventId, chP.eventId)
                End If
                mScript.mainClass(ch.classId).ChildProperties(i).Add(newName, chP)

                If IsNothing(chP.ThirdLevelProperties) = False AndAlso chP.ThirdLevelProperties.Count > 0 Then
                    'дочерним элементам 3 порядка
                    Dim newProp As MatewScript.ChildPropertiesInfoType = mScript.mainClass(ch.classId).ChildProperties(i)(newName)
                    For j As Integer = 0 To newProp.ThirdLevelProperties.GetUpperBound(0)
                        newProp.ThirdLevelProperties(j) = newData.Value
                        If eventId = -1 AndAlso chP.ThirdLevelEventId(j) > 0 Then
                            mScript.eventRouter.RemoveEvent(chP.ThirdLevelEventId(j))
                            chP.ThirdLevelEventId(j) = -1
                        ElseIf eventId > 0 Then
                            newProp.ThirdLevelEventId(j) = mScript.eventRouter.DuplicateEvent(eventId, chP.ThirdLevelEventId(j))
                        End If
                    Next j
                End If
            Next i
        End If

        'Если редактируемый контрол является текущим еще где-нибудь - убираем его из списка текущих
        For i As Integer = 0 To lstPanels.Count - 1
            With lstPanels(i)
                If .classId <> ch.classId Then Continue For
                If IsNothing(.lcActiveControl) Then Continue For
                If .lcActiveControl.Name = oldName Then
                    .lcActiveControl = Nothing
                    .IsWbVisible = False
                    .IsCodeBoxVisible = False
                End If
            End With
        Next
        If currentClassName = "A" Then actionsRouter.RenameProperty(oldName, newName)
        removedObjects.RenameProperty(ch.classId, oldName, newName)
        'Деактивируем и удаляем текущую панель
        ActivePanel.IsCodeBoxVisible = False
        ActivePanel.IsWbVisible = False
        Dim pnl As clsChildPanel = ActivePanel
        ActivePanel = Nothing
        Dim showAllFunctions As Boolean = dictDefContainers(ch.classId).showBasicFunctions
        If frmMainEditor.Controls.Contains(iconMenuElements) = False Then frmMainEditor.Controls.Add(iconMenuElements)
        If frmMainEditor.Controls.Contains(iconMenuGroups) = False Then frmMainEditor.Controls.Add(iconMenuGroups)
        dictDefContainers(ch.classId).Dispose()
        dictDefContainers.Remove(ch.classId)
        'Вносим изменения в скрипты
        'mScript.FillFuncAndPropHash()
        If oldName <> newName Then GlobalSeeker.ReplaceElementNameInStruct(ch.classId, oldName, newName, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
        'Создаем новую панель
        CreatePropertiesControl(ch.classId, showAllFunctions)
        Dim arrC() As Control = dictDefContainers(ch.classId).Controls.Find(newName, True)
        If arrC.Count > 0 Then
            pnl.lcActiveControl = arrC(0)
        Else
            MsgBox("Ошибка при вставки отредактированного свойства. Свойство не найдено.", vbExclamation)
            Return
        End If
        'изменяем текст узла
        If IsNothing(currentTreeView) = False Then
            Dim n As TreeNode = currentTreeView.SelectedNode
            If ch.child2Name.Length > 0 AndAlso IsNothing(n) = False Then
                frmMainEditor.tree_AfterSelect(currentTreeView, New TreeViewEventArgs(n))
            Else
                frmMainEditor.btnShowSettings_Click(btnDefSettings, New EventArgs)
            End If
        ElseIf mScript.mainClass(ch.classId).LevelsCount = 0 Then
            OpenPanel(Nothing, ch.classId, "")
            If dictDefContainers(ch.classId).Controls(0).GetType.Name = "SplitContainer" Then
                Dim sp As SplitContainer = dictDefContainers(ch.classId).Controls(0)
                del_splitProperies_ClientSizeChanged(sp.Panel1, New EventArgs)
            End If
        End If

        pnl.lcActiveControl.Focus()
    End Sub

    ''' <summary>
    ''' Переименовывает пользовательское свойство или функцию
    ''' </summary>
    ''' <param name="relatedControl">Контрол, ассоциированный с данным свойством/функцией</param>
    Private Sub ConfigMenu_RenameFuncOrProp(ByRef relatedControl As Object)
        Dim isUserFunction As Boolean = relatedControl.IsFunctionButton
        Dim oldName As String = relatedControl.Name
        Dim ch As clsChildPanel = relatedControl.childPanel
        Dim newName As String = InputBox("Введите новое имя " + IIf(isUserFunction, "функции ", "свойства ") + oldName + ":")
        If String.IsNullOrWhiteSpace(newName) Then Return
        newName = newName.Trim

        Dim p As MatewScript.PropertiesInfoType
        If isUserFunction Then
            'переименовываем пользователькую функцию
            'проверка существования одноименной функции
            If mScript.mainClass(ch.classId).Functions.ContainsKey(newName) Then
                MessageBox.Show("В классе " + mScript.mainClass(ch.classId).Names(mScript.mainClass(ch.classId).Names.GetUpperBound(0)) + " функция с именем " + newName + " уже существует.")
                Return
            End If
            'вносим изменения в mainClass
            p = mScript.mainClass(ch.classId).Functions(oldName)
            mScript.mainClass(ch.classId).Functions.Remove(oldName)
            If p.EditorCaption = oldName Then p.EditorCaption = newName
            mScript.mainClass(ch.classId).Functions.Add(newName, p)
            'переименовываем контролы
            With relatedControl
                relatedControl.Name = newName
                If IsNothing(.Label) = False Then
                    .Label.Name = "Label" + newName
                    .Label.Text = p.EditorCaption
                    .Label.Size = .Label.PreferredSize
                End If
                If IsNothing(.ButtonHelp) = False Then .ButtonHelp.Name = newName + "Help"
                If IsNothing(.ButtonConfig) = False Then .ButtonConfig.Name = newName + "Config"
            End With
            GlobalSeeker.ReplaceElementNameInStruct(ch.classId, oldName, newName, CodeTextBox.EditWordTypeEnum.W_FUNCTION)
            Return
        End If

        'переименовываем пользователькое свойство
        'проверка существования одноименного свойства
        If mScript.mainClass(ch.classId).Properties.ContainsKey(newName) Then
            MessageBox.Show("В классе " + mScript.mainClass(ch.classId).Names(mScript.mainClass(ch.classId).Names.GetUpperBound(0)) + " свойство с именем " + newName + " уже существует.")
            Return
        End If
        'изменения в событиях изменения свойств
        mScript.trackingProperties.RenameProperty(ch.classId, oldName, newName)
        'вносим изменения в mainClass
        If mScript.mainClass(ch.classId).Names(0) = "A" Then
            actionsRouter.RenameProperty(oldName, newName)
        End If
        removedObjects.RenameProperty(ch.classId, oldName, newName)
        'в свойства по умолчанию
        p = mScript.mainClass(ch.classId).Properties(oldName)
        mScript.mainClass(ch.classId).Properties.Remove(oldName)
        If p.EditorCaption = oldName Then p.EditorCaption = newName
        mScript.mainClass(ch.classId).Properties.Add(newName, p)
        'дочерним элементам
        If IsNothing(mScript.mainClass(ch.classId).ChildProperties) = False AndAlso mScript.mainClass(ch.classId).ChildProperties.Count > 0 Then
            For i As Integer = 0 To mScript.mainClass(ch.classId).ChildProperties.Count - 1
                Dim chP As MatewScript.ChildPropertiesInfoType = mScript.mainClass(ch.classId).ChildProperties(i)(oldName)
                mScript.mainClass(ch.classId).ChildProperties(i).Remove(oldName)
                mScript.mainClass(ch.classId).ChildProperties(i).Add(newName, chP)
            Next
        End If
        'переименовываем контролы
        With relatedControl
            .Name = newName
            If IsNothing(.Label) = False Then
                .Label.Name = "Label" + newName
                .Label.Text = p.EditorCaption
                .Label.Size = .Label.PreferredSize
            End If
            If IsNothing(.ButtonHelp) = False Then .ButtonHelp.Name = newName + "Help"
            If IsNothing(.ButtonConfig) = False Then .ButtonConfig.Name = newName + "Config"
        End With
        'возникает проблема с активными контролами. Их тоже переименовываем
        For i As Integer = 0 To lstPanels.Count - 1
            With lstPanels(i)
                If .classId <> ch.classId Then Continue For
                If IsNothing(.lcActiveControl) Then Continue For
                If .lcActiveControl.Name = oldName Then
                    .lcActiveControl = relatedControl
                End If
            End With
        Next
        'Вносим изменения в скрипты
        GlobalSeeker.ReplaceElementNameInStruct(ch.classId, oldName, newName, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
        Return
    End Sub


    ''' <summary>
    ''' Устанавливает всем элментам данного класса значение указанного свойства таким же, как в указанном контроле
    ''' </summary>
    ''' <param name="relatedControl">Контрол, ассоциированный с данным свойством</param>
    Private Sub ConfigMenu_SetValueToAll(ByRef relatedControl As Object)
        Dim propName As String = relatedControl.Name
        Dim ch As clsChildPanel = relatedControl.childPanel
        'получаем новое значение
        Dim newValue As String
        Dim child2Id As Integer = ch.GetChild2Id
        Dim child3Id As Integer = ch.GetChild3Id(child2Id)
        If child2Id < 0 Then
            newValue = mScript.mainClass(ch.classId).Properties(propName).Value
        ElseIf child3Id < 0 Then
            newValue = mScript.mainClass(ch.classId).ChildProperties(child2Id)(propName).Value
        Else
            newValue = mScript.mainClass(ch.classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id)
        End If
        'выводим подтверждение операции
        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(newValue)
        Dim strToPrint As String
        If cRes = MatewScript.ContainsCodeEnum.CODE Then
            strToPrint = My.Resources.script
        ElseIf cRes = MatewScript.ContainsCodeEnum.LONG_TEXT Then
            strToPrint = My.Resources.longText
        Else
            strToPrint = mScript.PrepareStringToPrint(newValue, Nothing, False)
        End If
        Dim res As DialogResult = MessageBox.Show("Установить абсолютно всем элементам этого класса свойство " + propName + " равным " + strToPrint + "?", "Matew Quest", MessageBoxButtons.YesNo)
        If res = DialogResult.No Then Return

        'Устанавливаем значение по умолчанию
        If child2Id >= 0 Then SetPropertyValue(ch.classId, propName, newValue, -1, -1)

        If IsNothing(mScript.mainClass(ch.classId).ChildProperties) = False AndAlso mScript.mainClass(ch.classId).ChildProperties.Count > 0 Then
            For i As Integer = 0 To mScript.mainClass(ch.classId).ChildProperties.Count - 1
                'Устанавливаем значение элементам 2 уровня
                SetPropertyValue(ch.classId, propName, newValue, i, -1)

                If IsNothing(mScript.mainClass(ch.classId).ChildProperties(i)(propName).ThirdLevelProperties) = False AndAlso _
                    mScript.mainClass(ch.classId).ChildProperties(i)(propName).ThirdLevelProperties.Count > 0 Then
                    'Устанавливаем значение элементам 3 уровня
                    For j As Integer = 0 To mScript.mainClass(ch.classId).ChildProperties(i)(propName).ThirdLevelProperties.Count - 1
                        SetPropertyValue(ch.classId, propName, newValue, i, j)
                    Next j
                End If
            Next i
        End If
        If mScript.mainClass(ch.classId).Names(0) = "A" Then
            actionsRouter.SetPropertyValue(propName, newValue)
        End If
        'на всякий случай убираем контрол, значение в котором меняется, из списка активных на всех панелях (пригодится если меняется простое свойство на скрипт или обратно)
        For i As Integer = 0 To lstPanels.Count - 1
            Dim curPanel As clsChildPanel = lstPanels(i)
            With curPanel
                If .classId <> ch.classId Then Continue For
                Dim c As Control = .lcActiveControl
                If Object.Equals(c, relatedControl) Then
                    .lcActiveControl = Nothing
                    .IsWbVisible = False
                    .IsCodeBoxVisible = False
                End If
            End With
        Next
    End Sub

    ''' <summary>
    ''' Устанавливает всем дочерним элементам данного родителя значение указанного свойства таким же, как в указанном контроле
    ''' </summary>
    ''' <param name="relatedControl">Контрол, ассоциированный с данным свойством</param>
    Private Sub ConfigMenu_SetValueToChildren(ByRef relatedControl As Object)
        Dim propName As String = relatedControl.Name
        Dim ch As clsChildPanel = relatedControl.childPanel
        Dim child2Id As Integer = ch.GetChild2Id
        Dim child3Id As Integer = ch.GetChild3Id(child2Id)
        If child2Id < 0 OrElse child3Id > 0 OrElse mScript.mainClass(ch.classId).LevelsCount < 2 Then Return


        'получаем новое значение
        Dim newValue As String
        newValue = mScript.mainClass(ch.classId).ChildProperties(child2Id)(propName).Value

        ''выводим подтверждение операции
        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(newValue)
        Dim strToPrint As String
        If cRes = MatewScript.ContainsCodeEnum.CODE Then
            strToPrint = My.Resources.script
        ElseIf cRes = MatewScript.ContainsCodeEnum.LONG_TEXT Then
            strToPrint = My.Resources.longText
        Else
            strToPrint = mScript.PrepareStringToPrint(newValue, Nothing, False)
        End If
        'Dim res As DialogResult = MessageBox.Show("Установить всем дочерним по отношению к " + mScript.PrepareStringToPrint(ch.child2Name, Nothing, False) + " элементам свойство " + propName + " равным " + strToPrint + "?", "Matew Quest", MessageBoxButtons.YesNo)
        'If res = DialogResult.No Then Return

        If IsNothing(mScript.mainClass(ch.classId).ChildProperties(child2Id)(propName).ThirdLevelProperties) OrElse _
            mScript.mainClass(ch.classId).ChildProperties(child2Id)(propName).ThirdLevelProperties.Count = 0 Then
            MessageBox.Show("У данного элемента 2 порядка не найдено ни одного дочернего элемента.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim tree As TreeView = dlgSetChildrenValue.treeChildren
        tree.ImageList = frmMainEditor.imgLstGroupIcons
        frmMainEditor.FillTree(currentClassName, tree, frmMainEditor.chkShowHidden.Checked, Nothing, ch.child2Name)
        If tree.Nodes.Count = 0 Then
            MessageBox.Show("У данного элемента 2 порядка не найдено ни одного нескрытого дочернего элемента.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        'выделяем все узлы
        For i As Integer = 0 To tree.Nodes.Count - 1
            tree.Nodes(i).Checked = True
        Next
        dlgSetChildrenValue.lblValueToSet.Text = "Свойство " & propName & ", новое значение: " & strToPrint
        Dim res As DialogResult = dlgSetChildrenValue.ShowDialog(frmMainEditor)

        ''Устанавливаем значение элементам 3 уровня
        'For i As Integer = 0 To mScript.mainClass(ch.classId).ChildProperties(i)(propName).ThirdLevelProperties.Count - 1
        '    SetPropertyValue(ch.classId, propName, newValue, child2Id, i)
        'Next i
        If res = DialogResult.Cancel Then Return
        'Устанавливаем значение элементам 3 уровня
        For i As Integer = 0 To tree.Nodes.Count - 1
            Dim n As TreeNode = tree.Nodes(i)
            If n.Checked = False Then Continue For
            If n.Tag = "GROUP" Then
                If n.Nodes.Count > 0 Then
                    For j As Integer = 0 To n.Nodes.Count - 1
                        child3Id = GetThirdChildIdByName(WrapString(n.Nodes(j).Text), child2Id, mScript.mainClass(ch.classId).ChildProperties)
                        If child3Id > -1 Then SetPropertyValue(ch.classId, propName, newValue, child2Id, child3Id)
                    Next j
                End If
            Else
                child3Id = GetThirdChildIdByName(WrapString(n.Text), child2Id, mScript.mainClass(ch.classId).ChildProperties)
                If child3Id > -1 Then SetPropertyValue(ch.classId, propName, newValue, child2Id, child3Id)
            End If
        Next i

        'на всякий случай убираем контрол, значение в котором меняется, из списка активных на всех панелях (пригодится если меняется простое свойство на скрипт или обратно)
        For i As Integer = 0 To lstPanels.Count - 1
            Dim curPanel As clsChildPanel = lstPanels(i)
            With curPanel
                If .classId <> ch.classId Then Continue For
                Dim c As Control = .lcActiveControl
                If Object.Equals(c, relatedControl) Then
                    .lcActiveControl = Nothing
                    .IsWbVisible = False
                    .IsCodeBoxVisible = False
                End If
            End With
        Next
    End Sub

    ''' <summary>
    ''' Удаляет пользовательское свойство или функцию
    ''' </summary>
    ''' <param name="relatedControl">Контрол, ассоциированный с данным свойством/функцией</param>
    Private Sub ConfigMenu_RemoveUserElement(ByRef relatedControl As Object)
        Dim removeFunction As Boolean = relatedControl.IsFunctionButton 'если True, то удаляет функцию; иначе - свойство

        'Удаление свойства/функции пользователя
        Dim elName As String = relatedControl.Name 'Имя свойства 
        Dim ch As clsChildPanel = relatedControl.childPanel
        'убираем удаляемый контрол из активных контролов всех панелей
        If IsNothing(ActivePanel.lcActiveControl) = False Then
            For i As Integer = 0 To lstPanels.Count - 1
                If lstPanels(i).classId <> ch.classId Then Continue For
                If IsNothing(lstPanels(i).ActiveControl) Then Continue For
                If IsNothing(ActivePanel.lcActiveControl) = False AndAlso lstPanels(i).ActiveControl.Name = ActivePanel.lcActiveControl.Name Then lstPanels(i).lcActiveControl = Nothing
            Next
        End If
        ch.ActiveControl = Nothing
        'Запускаем соответствующую функцию движка
        Dim strToExecute As String
        If removeFunction Then
            strToExecute = "Script.RemoveFunction('" + mScript.mainClass(ch.classId).Names(0) + "', '" + elName + "')"
        Else
            If currentClassName = "A" Then actionsRouter.RemoveProperty(elName)
            removedObjects.RemoveProperty(ch.classId, elName) 'удаляеем свойство в удаленных элементах
            'strToExecute = mScript.mainClass(ch.classId).Names(0) + ".RemoveProperty('" + elName + "')"
            strToExecute = "Script.RemoveProperty('" + mScript.mainClass(ch.classId).Names(0) + "', '" + elName + "')"
            'удаляем из событий изменения свойств
            mScript.trackingProperties.RemoveProperty(ch.classId, elName)
        End If
        Dim res As String = mScript.ExecuteString({strToExecute}, Nothing)
        If res = "#Error" Then Return
        'Деактивируем и удаляем текущую панель 
        ActivePanel.IsCodeBoxVisible = False
        ActivePanel.IsWbVisible = False
        Dim pnl As clsChildPanel = ActivePanel
        ActivePanel = Nothing
        Dim showAllFunctions As Boolean = dictDefContainers(ch.classId).showBasicFunctions
        dictDefContainers(ch.classId).Dispose()
        dictDefContainers.Remove(ch.classId)
        'Ищем удаленный элемент в скрипах
        GlobalSeeker.CheckElementNameInStruct(ch.classId, elName, IIf(removeFunction, CodeTextBox.EditWordTypeEnum.W_FUNCTION, CodeTextBox.EditWordTypeEnum.W_PROPERTY))
        'mScript.FillFuncAndPropHash()
        'Создаем новую панель
        CreatePropertiesControl(ch.classId, showAllFunctions)
        If ch.child2Name.Length = 0 Then
            If mScript.mainClass(ch.classId).LevelsCount > 0 Then
                frmMainEditor.btnShowSettings_Click(btnDefSettings, New EventArgs)
            Else
                OpenPanel(Nothing, ch.classId, "")
                If dictDefContainers(ch.classId).Controls(0).GetType.Name = "SplitContainer" Then
                    Dim sp As SplitContainer = dictDefContainers(ch.classId).Controls(0)
                    del_splitProperies_ClientSizeChanged(sp.Panel1, New EventArgs)
                End If
            End If
        Else
            Dim n As TreeNode = currentTreeView.SelectedNode
            If IsNothing(n) Then Return
            frmMainEditor.tree_AfterSelect(currentTreeView, New TreeViewEventArgs(n))
        End If

    End Sub

    ''' <summary>
    ''' Превращает свойство в скрипт или в длинный текст
    ''' </summary>
    ''' <param name="relatedControl">Контрол, ассоциированный с данным свойством</param>
    ''' <param name="toLongText">Если True, то превращать в длинный текст; иначе - в скрипт</param>
    Private Sub ConfigMenu_ExecutableProperty(ByRef relatedControl As Object, ByVal toLongText As Boolean)
        Dim ch As clsChildPanel = relatedControl.childPanel
        ch.ActiveControl = relatedControl

        Dim cbPanel As Panel = frmMainEditor.codeBoxPanel
        Dim cb As CodeTextBox = frmMainEditor.codeBox
        Dim propValue As String = ""
        Dim propName As String = relatedControl.Name

        frmMainEditor.WBhelp.Hide()
        cb.Tag = Nothing
        cb.codeBox.IsTextBlockByDefault = False
        Dim child2Id As Integer = ch.GetChild2Id
        Dim child3Id As Integer = ch.GetChild3Id(child2Id)
        If child2Id < 0 Then
            propValue = mScript.mainClass(ch.classId).Properties(propName).Value
        ElseIf child3Id < 0 Then
            propValue = mScript.mainClass(ch.classId).ChildProperties(child2Id)(propName).Value
        Else
            propValue = mScript.mainClass(ch.classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id)
        End If
        cb.codeBox.IsTextBlockByDefault = toLongText
        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(propValue)
        If cRes = MatewScript.ContainsCodeEnum.CODE OrElse cRes = MatewScript.ContainsCodeEnum.LONG_TEXT Then
            cb.codeBox.LoadCodeFromProperty(propValue)
        Else
            cb.Text = mScript.PrepareStringToPrint(propValue, Nothing, False)
        End If
        'изменяем вид контролов, связанных со свойством, которое делается кодом
        Dim tb As TextBoxEx
        If relatedControl.GetType.Name = "ComboBoxEx" Then
            tb = ReplaceComboWithTextBox(relatedControl)
        Else
            tb = relatedControl
        End If
        tb.ReadOnly = True
        If toLongText Then
            tb.Text = My.Resources.longText
        Else
            tb.Text = My.Resources.script
        End If

        cb.Tag = tb
        cbPanel.Show()
    End Sub

    ''' <summary>
    ''' Замещает комбобокс на текстбокс при изменении отображения свойства из списка в скрипт
    ''' </summary>
    ''' <param name="tb">Текстбокс для замены на комбобокс</param>
    ''' <returns></returns>
    Private Function ReplaceTextBoxWithCombo(ByRef tb As TextBoxEx) As ComboBoxEx
        'сохраняем все необходимые для переноса свойства
        Dim ch As clsChildPanel = tb.childPanel
        Dim cSize As New Size(tb.Size)
        Dim cPos As New Point(tb.Location)
        Dim cParent As Control = tb.Parent
        Dim cName As String = tb.Name
        Dim fColor As Color = tb.ForeColor
        Dim bColor As Color = tb.BackColor
        Dim lbl As Label = tb.Label
        Dim btnHelp As ButtonEx = tb.ButtonHelp
        Dim btnConf As ButtonEx = tb.ButtonConfig
        Dim hasConf As Boolean = tb.hasConfigButton
        Dim pTop As cPropertyControlsTop = tb.PropertyControlsTop

        'удаляем текстбокс и создаем на его месте комбо
        Dim combo As New ComboBoxEx With {.Size = cSize, .Location = cPos, .childPanel = ch, .Name = cName, .ForeColor = fColor, .BackColor = bColor, .Label = lbl, .ButtonConfig = btnConf, _
                                      .ButtonHelp = btnHelp, .PropertyControlsTop = pTop, .hasConfigButton = hasConf}
        cParent.Controls.Add(combo)
        If Not IsNothing(lbl) Then cParent.Controls.Add(lbl)
        If Not IsNothing(btnConf) Then cParent.Controls.Add(btnConf)
        If Not IsNothing(btnHelp) Then cParent.Controls.Add(btnHelp)

        If IsNothing(lstPanels) = False AndAlso lstPanels.Count > 0 Then
            'заменяем во всех панелях старый контрол на новый
            For i As Integer = 0 To lstPanels.Count - 1
                Dim curCh As clsChildPanel = lstPanels(i)
                If ch.classId <> curCh.classId Then Continue For
                Dim ac As Control = curCh.ActiveControl
                If IsNothing(ac) Then Continue For
                If Object.Equals(ac, tb) Then lstPanels(i).lcActiveControl = combo
            Next
        End If
        'вставляем список
        Dim lst() As String = mScript.mainClass(ch.classId).Properties(cName).returnArray
        If IsNothing(lst) = False AndAlso lst.Count > 0 Then
            For i As Integer = 0 To lst.Count - 1
                combo.Items.Add(mScript.PrepareStringToPrint(lst(i), Nothing, False))
            Next
        End If
        'удаляем текстбокс
        tb.Label = Nothing
        tb.ButtonConfig = Nothing
        tb.ButtonHelp = Nothing
        tb.PropertyControlsTop = Nothing
        tb.Dispose()

        'назначаем обработчики
        AddHandler combo.SelectedIndexChanged, AddressOf del_propertyTextChanged
        AddHandler combo.Validating, AddressOf del_propertyValidating
        AddHandler combo.MouseEnter, Sub(sender As Object, e As EventArgs) dictDefContainers(ch.classId).Controls.Find("lblPropertyDescription", True)(0).Text = mScript.mainClass(ch.classId).Properties(cName).Description
        AddHandler combo.MouseDown, AddressOf del_TextPropertyMouseDown

        Return combo
    End Function

    ''' <summary>
    ''' Замещает комбобокс на текстбокс при изменении отображения свойства из списка в скрипт
    ''' </summary>
    ''' <param name="comboContorl">Комбобокс для замены на текстбокс</param>
    ''' <returns>TextBoxEx</returns>
    Public Function ReplaceComboWithTextBox(ByRef comboContorl As Control) As Control
        Dim combo As ComboBoxEx = comboContorl
        'сохраняем все необходимые для переноса свойства
        Dim ch As clsChildPanel = combo.childPanel
        Dim cSize As New Size(combo.Size)
        Dim cPos As New Point(combo.Location)
        Dim cParent As Control = combo.Parent
        Dim cName As String = combo.Name
        Dim fColor As Color = combo.ForeColor
        Dim bColor As Color = combo.BackColor
        Dim lbl As Label = combo.Label
        Dim btnHelp As ButtonEx = combo.ButtonHelp
        Dim btnConf As ButtonEx = combo.ButtonConfig
        Dim hasConf As Boolean = combo.hasConfigButton
        Dim pTop As cPropertyControlsTop = combo.PropertyControlsTop

        'удаляем комбо и создаем на его месте текстбокс
        Dim tb As New TextBoxEx With {.Size = cSize, .Location = cPos, .childPanel = ch, .Name = cName, .ForeColor = fColor, .BackColor = bColor, .Label = lbl, .ButtonConfig = btnConf, _
                                      .ButtonHelp = btnHelp, .PropertyControlsTop = pTop, .hasConfigButton = hasConf}
        cParent.Controls.Add(tb)
        If Not IsNothing(lbl) Then cParent.Controls.Add(lbl)
        If Not IsNothing(btnConf) Then cParent.Controls.Add(btnConf)
        If Not IsNothing(btnHelp) Then cParent.Controls.Add(btnHelp)

        If IsNothing(lstPanels) = False AndAlso lstPanels.Count > 0 Then
            'заменяем во всех панелях старый контрол на новый
            For i As Integer = 0 To lstPanels.Count - 1
                Dim curCh As clsChildPanel = lstPanels(i)
                If ch.classId <> curCh.classId Then Continue For
                Dim ac As Control = curCh.ActiveControl
                If IsNothing(ac) Then Continue For
                If Object.Equals(ac, combo) Then lstPanels(i).lcActiveControl = tb
            Next
        End If
        combo.Label = Nothing
        combo.ButtonConfig = Nothing
        combo.ButtonHelp = Nothing
        combo.PropertyControlsTop = Nothing
        combo.Dispose()

        'назначаем обработчики
        AddHandler tb.Validating, AddressOf del_propertyValidating
        AddHandler tb.MouseEnter, Sub(sender As Object, e As EventArgs) dictDefContainers(ch.classId).Controls.Find("lblPropertyDescription", True)(0).Text = mScript.mainClass(ch.classId).Properties(cName).Description
        AddHandler tb.MouseDown, AddressOf del_TextPropertyMouseDown
        AddHandler tb.TextChanged, AddressOf del_propertyTextChanged

        Return tb
    End Function

    ''' <summary>
    ''' Востанавливает контрол но умолчанию если он был изменен на текстбокс всвязи с тем, что использован скрипт
    ''' </summary>
    ''' <param name="c">контрол-текстбокс</param>
    ''' <param name="restoreDefaultValue">Восстанавливать ли значение по умолчанию</param>
    ''' <returns>новый контрол или тот же, если он не изменился</returns>
    Public Function RestoreDefaultControl(ByRef c As Control, Optional restoreDefaultValue As Boolean = False) As Control
        If c.GetType.Name <> "TextBoxEx" Then Return c 'если строка исполняема, то комбобокс заменяется на текстбокс. Если у нас не тестбокс, то заменять контро не надо

        'получаем контрол и сохраняем его характеристики
        Dim tb As TextBoxEx = c
        Dim ch As clsChildPanel = tb.childPanel
        Dim cSize As New Size(tb.Size)
        Dim cPos As New Point(tb.Location)
        Dim cParent As Control = tb.Parent
        Dim cName As String = tb.Name
        Dim fColor As Color = tb.ForeColor
        Dim bColor As Color = tb.BackColor
        Dim lbl As Label = tb.Label
        Dim btnHelp As ButtonEx = tb.ButtonHelp
        Dim btnConf As ButtonEx = tb.ButtonConfig
        Dim pTop As cPropertyControlsTop = tb.PropertyControlsTop

        Dim prop As MatewScript.PropertiesInfoType = mScript.mainClass(ch.classId).Properties(cName)

        Dim propCopy As MatewScript.PropertiesInfoType = Nothing
        Dim blnRestore As Boolean = False

        If mScript.mainClassCopy(ch.classId).Properties.TryGetValue(cName, propCopy) Then
            If IsPropertyContainsEnum(propCopy) OrElse IsPropertyContainsEnum(prop) OrElse propCopy.returnType = MatewScript.ReturnFunctionEnum.RETURN_BOOl Then
                'надо восстановить массив
                If Not IsPropertyContainsEnum(prop) Then
                    RestoreDefaultArray(ch.classId, cName)
                End If
                blnRestore = True
            End If
        Else
            If IsPropertyContainsEnum(prop) OrElse prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_BOOl Then
                'надо восстановить массив
                blnRestore = True
            End If
        End If

        Dim defValue As String
        Dim newC As Object
        If blnRestore Then
            'создаем новый комбо и ставим его взамен текстокса
            Dim combo As New ComboBoxEx With {.Size = cSize, .Location = cPos, .Name = cName, .ForeColor = fColor, .BackColor = bColor, .childPanel = ch, .Label = lbl, .ButtonConfig = btnConf, _
                                      .ButtonHelp = btnHelp, .PropertyControlsTop = pTop}
            cParent.Controls.Add(combo)

            If IsNothing(lstPanels) = False AndAlso lstPanels.Count > 0 Then
                'заменяем во всех панелях старый контрол на новый
                For i As Integer = 0 To lstPanels.Count - 1
                    Dim curCh As clsChildPanel = lstPanels(i)
                    If ch.classId <> curCh.classId Then Continue For
                    Dim ac As Control = curCh.ActiveControl
                    If IsNothing(ac) Then Continue For
                    If Object.Equals(ac, tb) Then lstPanels(i).lcActiveControl = combo
                Next
            End If

            'удаляем старый контрол через таймер, поскольку юзер мог успеть клацнуть по нему и в итоге возникает всякая хрень
            tb.Hide()
            tb.Name = ""
            Dim t As New Timer With {.Interval = 1}
            AddHandler t.Tick, Sub(sender As Object, e As EventArgs)
                                   sender.Stop()
                                   tb.Label = Nothing
                                   tb.ButtonConfig = Nothing
                                   tb.ButtonHelp = Nothing
                                   tb.PropertyControlsTop = Nothing
                                   tb.Dispose()
                                   If Not IsNothing(lbl) Then cParent.Controls.Add(lbl) : lbl.Visible = True
                                   If Not IsNothing(btnConf) Then cParent.Controls.Add(btnConf) : btnConf.Visible = True
                                   If Not IsNothing(btnHelp) Then cParent.Controls.Add(btnHelp) : btnHelp.Visible = True
                               End Sub
            t.Start()

            cListManager.FillListByChildPanel(cName, ch, combo, False)
            If prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_BOOl Then combo.DropDownStyle = ComboBoxStyle.DropDownList
            'If prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_BOOl Then
            '    combo.Items.AddRange({"True", "False"})
            '    combo.DropDownStyle = ComboBoxStyle.DropDownList
            'ElseIf prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_ENUM Then
            '    'восстанавливаем исходный список
            '    If IsNothing(prop.returnArray) = False AndAlso prop.returnArray.Count > 0 Then
            '        For i As Integer = 0 To prop.returnArray.Count - 1
            '            combo.Items.Add(mScript.PrepareStringToPrint(prop.returnArray(i), Nothing, False))
            '        Next
            '    End If
            'ElseIf prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then
            'End If

            AddHandler combo.SelectedIndexChanged, AddressOf del_propertyTextChanged
            AddHandler combo.Validating, AddressOf del_propertyValidating
            AddHandler combo.MouseEnter, Sub(sender As Object, e As EventArgs) dictDefContainers(ch.classId).Controls.Find("lblPropertyDescription", True)(0).Text = _
                                             mScript.mainClass(ch.classId).Properties(cName).Description
            AddHandler combo.MouseDown, AddressOf del_TextPropertyMouseDown
            newC = combo
        Else
            newC = tb
        End If

        If restoreDefaultValue Then
            'устанавливаем значение по умолчанию
            defValue = mScript.mainClass(ch.classId).Properties(cName).Value
            If ch.child2Name.Length = 0 Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClassCopy(ch.classId).Properties.TryGetValue(cName, p) Then
                    defValue = p.Value
                Else
                    defValue = "''"
                End If
            End If
            Dim child2Id As Integer = ch.GetChild2Id
            Dim child3Id As Integer = ch.GetChild3Id(child2Id)
            SetPropertyValue(ch.classId, cName, defValue, child2Id, child3Id)
            Select Case mScript.IsPropertyContainsCode(defValue)
                Case MatewScript.ContainsCodeEnum.CODE
                    defValue = My.Resources.script
                Case MatewScript.ContainsCodeEnum.LONG_TEXT
                    defValue = My.Resources.longText
                Case Else
                    defValue = mScript.PrepareStringToPrint(defValue, Nothing, False)
            End Select

            newC.Text = defValue
        End If

        Return newC
    End Function
#End Region

    ''' <summary>
    ''' Удаляет панель удаляемого элемента и панели дочерних элементов.
    ''' </summary>
    ''' <param name="removingNode"></param>
    ''' <remarks></remarks>
    Public Sub UpdateAfterNodeRemoving(ByRef removingNode As TreeNode)
        AAA()
        If IsNothing(lstPanels) OrElse lstPanels.Count = 0 Then Return
        Dim classId As Integer = -1
        Select Case currentClassName
            Case "Variable"
                classId = -2
            Case "Function"
                classId = -3
            Case Else
                classId = mScript.mainClassHash(currentClassName)
        End Select

        'Удаляем панели дочерних элементов
        Dim pnlToRemove As clsChildPanel = Nothing
        Dim remName As String = WrapString(removingNode.Text)
        If classId > -1 AndAlso mScript.mainClass(classId).LevelsCount = 2 Then
            For pId As Integer = lstPanels.Count - 1 To 0 Step -1
                If lstPanels(pId).classId <> classId Then Continue For 'не тот класс
                If lstPanels(pId).child2Name <> remName Then Continue For 'родитель не тот
                If lstPanels(pId).child3Name = "" Then
                    'элемент 2 уровня (эта же панель, но ее удалим позже)
                    pnlToRemove = lstPanels(pId)
                Else
                    'элемент 3 уровня даного класса - удаляем
                    Dim p As clsChildPanel = lstPanels(pId)
                    RemovePanel(p, False)
                End If
            Next pId
        ElseIf classId = mScript.mainClassHash("L") Then
            Dim classAct As Integer = mScript.mainClassHash("A")
            For pId As Integer = lstPanels.Count - 1 To 0 Step -1
                Dim p As clsChildPanel = lstPanels(pId)
                If p.classId <> classAct Then Continue For 'не тот класс
                If p.supraElementName <> remName Then Continue For 'родитель не тот
                If p.child2Name.Length = 0 Then Continue For 'действия по умолчанию
                'действия, родитель - удаляемая локация. Удаляем
                RemovePanel(p, False)
            Next pId
            pnlToRemove = GetPanelByNode(removingNode)
        Else
            pnlToRemove = GetPanelByNode(removingNode)
        End If

        'Удаляем наш элемент
        If IsNothing(pnlToRemove) = False Then
            RemovePanel(pnlToRemove)
        End If
        AAA()
    End Sub

    Public Sub ReassociateAllNodes()
        If IsNothing(lstPanels) OrElse lstPanels.Count = 0 Then Return
        Log.PrintToLog("ReassociateAllNodes")

        Dim tree As TreeView, n As TreeNode, txt As String
        For i As Integer = lstPanels.Count - 1 To 0 Step -1
            Dim pnl As clsChildPanel = lstPanels(i)
            'If pnl.supraElementName.Length > 0 Then Continue For
            If pnl.child3Name.Length > 0 Then
                'pnl.treeNode = New TreeNode(mScript.PrepareStringToPrint(pnl.child3Name, Nothing, False))
                Continue For
            End If
            If pnl.child2Name.Length = 0 Then Continue For
            If pnl.classId = -2 Then
                tree = frmMainEditor.treeVariables
            ElseIf pnl.classId = -3 Then
                tree = frmMainEditor.treeFunctions
            Else
                If mScript.mainClass(pnl.classId).LevelsCount <> 1 Then Continue For
                tree = Nothing
                frmMainEditor.dictTrees.TryGetValue(pnl.classId, tree)
                If IsNothing(tree) Then Continue For
            End If
            txt = pnl.child2Name
            If txt.StartsWith("'") Then txt = mScript.PrepareStringToPrint(txt, Nothing, False)
            n = frmMainEditor.FindItemNodeByText(tree, txt)
            If IsNothing(n) Then
                'RemovePanel(pnl)
            Else
                pnl.treeNode = n
            End If
        Next

    End Sub
    ''' <summary>
    ''' Пересопоставляет узлы в каждой панели если есть вероятность того, что связь была потеряна (например, при перестройке дерева)
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    ''' <param name="tree">дерево, с узлами котрого происходит ресопоставление</param>
    Public Sub ReassociateNodes(ByVal classId As Integer, ByRef tree As TreeView, ByVal supraName As String)
        If IsNothing(lstPanels) OrElse lstPanels.Count = 0 Then Return
        Log.PrintToLog("ReassociateNodes: " + mScript.mainClass(classId).Names(0) + ", supraName: " + IIf(supraName.Length > 0, supraName, "-"))
        Dim thirdLevel As Boolean = frmMainEditor.IsThirdLevelTree(tree)

        Dim txt As String
        For i As Integer = lstPanels.Count - 1 To 0 Step -1
            Dim pnl As clsChildPanel = lstPanels(i)
            If pnl.classId <> classId Then Continue For
            If pnl.supraElementName <> supraName Then Continue For
            If pnl.child3Name.Length > 0 AndAlso thirdLevel = False Then Continue For
            If pnl.classId < 0 AndAlso thirdLevel = True Then Continue For
            If pnl.child2Name.Length = 0 Then Continue For
            Dim n As TreeNode = pnl.treeNode
            'If IsNothing(n) Then Continue For
            If pnl.child3Name.Length > 0 Then
                txt = mScript.PrepareStringToPrint(pnl.child3Name, Nothing, False)
            Else
                txt = mScript.PrepareStringToPrint(pnl.child2Name, Nothing, False)
            End If
            n = frmMainEditor.FindItemNodeByText(tree, txt) ' n.Text)
            If IsNothing(n) Then
                RemovePanel(pnl)
            Else
                pnl.treeNode = n
            End If
        Next
    End Sub

    ''' <summary>
    ''' Пересопоставляет узлы в каждой панели переменных если есть вероятность того, что связь была потеряна (например, при перестройке дерева)
    ''' </summary>
    Public Sub ReassociateVariablesNodes()
        Dim tree As TreeView = frmMainEditor.treeVariables
        Dim classId As Integer = -2
        If IsNothing(lstPanels) OrElse lstPanels.Count = 0 Then Return
        Log.PrintToLog("ReassociateNodes: Variables")

        For i As Integer = lstPanels.Count - 1 To 0 Step -1
            Dim pnl As clsChildPanel = lstPanels(i)
            If pnl.classId <> classId Then Continue For
            If pnl.child2Name.Length = 0 Then Continue For
            Dim n As TreeNode = pnl.treeNode
            n = frmMainEditor.FindItemNodeByText(tree, pnl.child2Name) ' n.Text)
            If IsNothing(n) Then
                RemovePanel(pnl)
            Else
                pnl.treeNode = n
            End If
        Next
    End Sub

    ''' <summary>
    ''' Пересопоставляет узлы в каждой панели функций если есть вероятность того, что связь была потеряна (например, при перестройке дерева)
    ''' </summary>
    Public Sub ReassociateFunctionsNodes()
        Dim tree As TreeView = frmMainEditor.treeFunctions
        Dim classId As Integer = -3
        If IsNothing(lstPanels) OrElse lstPanels.Count = 0 Then Return
        Log.PrintToLog("ReassociateNodes: Functions")

        For i As Integer = lstPanels.Count - 1 To 0 Step -1
            Dim pnl As clsChildPanel = lstPanels(i)
            If pnl.classId <> classId Then Continue For
            If pnl.child2Name.Length = 0 Then Continue For
            Dim n As TreeNode = pnl.treeNode
            n = frmMainEditor.FindItemNodeByText(tree, pnl.child2Name) 'n.Text)
            If IsNothing(n) Then
                RemovePanel(pnl)
            Else
                pnl.treeNode = n
            End If
        Next
    End Sub

    ''' <summary>
    ''' Изменяет имя элемента в соответствующей ему вкладке после переименования. Также изменяет имена во вкладках дачерних элементов
    ''' </summary>
    ''' <param name="ch">Вкладка элемента, который был переименован</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя элемента</param>
    Public Sub ChangePanelChildsName(ByRef ch As clsChildPanel, ByVal oldName As String, ByVal newName As String)
        AAA()
        Dim isThirdLevelChild As Boolean = False
        'Изменяем имя в текущей панели
        If ch.child3Name.Length > 0 Then
            ch.child3Name = newName
            isThirdLevelChild = True
        Else
            ch.child2Name = newName
        End If
        ch.toolButton.Text = MakeToolButtonText(ch)

        If isThirdLevelChild OrElse ch.classId < 0 Then Return
        If mScript.mainClass(ch.classId).LevelsCount = 2 Then
            'Вкладка 2 уровня. Изменяем родителей вкладкам 3 уровня того же класса
            For pId As Integer = 0 To lstPanels.Count - 1
                Dim p As clsChildPanel = lstPanels(pId)
                If p.classId <> ch.classId Then Continue For
                If p.child2Name <> oldName Then Continue For
                p.child2Name = newName
                If p.supraElementName.Length > 0 Then p.supraElementName = newName
                p.toolButton.Text = MakeToolButtonText(p)
            Next pId
            'Переименовываем в удаленных элементах
            removedObjects.RenameParent(ch.classId, oldName, newName)
        ElseIf ch.classId = mScript.mainClassHash("L") Then
            'Вкладка локации. Изменяем вкладки действий
            Dim classAct As Integer = mScript.mainClassHash("A")
            For pId As Integer = 0 To lstPanels.Count - 1
                Dim p As clsChildPanel = lstPanels(pId)
                If p.classId <> classAct Then Continue For
                If p.supraElementName <> oldName Then Continue For
                p.supraElementName = newName
                p.toolButton.Text = MakeToolButtonText(p)
            Next pId
            'Переименовываем в удаленных элементах
            removedObjects.RenameParent(ch.classId, oldName, newName)
        End If
        AAA()
    End Sub

    Public Sub New(ByRef parent As SplitContainer, ByRef WBhelp As WebBrowser, ByRef panelToolStrip As ToolStrip, ByRef btnDefSettings As Button)
        Me.parent = parent
        Me.WBhelp = WBhelp
        Me.panelToolStrip = panelToolStrip
        Me.btnDefSettings = btnDefSettings
        CreateSubTreeMenu()
        CreateConfigButtonMenu()

        'CreatePropertiesControls()
    End Sub

    ''' <summary>
    ''' Создает подпись для вкладки панели
    ''' </summary>
    ''' <param name="childPanel">панель, хозяйка вкладки</param>
    Public Function MakeToolButtonText(ByRef childPanel As clsChildPanel) As String
        Dim pName As String
        If childPanel.child2Name.Length = 0 AndAlso childPanel.classId >= 0 Then
            If mScript.mainClass(childPanel.classId).LevelsCount = 0 Then
                pName = "''"
            Else
                pName = "'по умолчанию'"
            End If
        ElseIf childPanel.child3Name.Length = 0 Then
            pName = childPanel.child2Name
            If childPanel.classId = -2 Then
                Return pName + "  "
            ElseIf childPanel.classId = -3 Then
                Return pName + "  "
            End If
        Else
            pName = childPanel.child3Name
        End If
        'Dim parName As String = ""
        'If childPanel.supraElementName.Length > 0 Then
        '    parName = " (" + mScript.PrepareStringToPrint(childPanel.supraElementName, Nothing, False) + ")"
        'End If
        pName = mScript.PrepareStringToPrint(pName, Nothing, False)
        'If pName.Length > 0 Then pName = " - " + pName
        If pName.Length = 0 Then ' AndAlso parName.Length = 0 Then
            If childPanel.classId = -1 Then
                Return "Variable  "
            ElseIf childPanel.classId = -2 Then
                Return "Function  "
            Else
                Return "Класс " + mScript.mainClass(childPanel.classId).Names(0) + "  "
            End If
        Else
            Return pName + "  " ' parName + "  "
        End If
    End Function

    Dim isCrossUnderMouse As Boolean = False 'переменная для определения находится ли указатель мыши над крестиком закрытия вкладки (одна переменная для всех вкладок)
    ''' <summary>
    ''' Создает вкладку для панели
    ''' </summary>
    ''' <param name="childPanel">Панель, для которой надо создать вкладку</param>
    ''' <remarks></remarks>
    Public Sub CreateToolButton(ByRef childPanel As clsChildPanel)
        'Создаем подпись для вкладки
        Dim strText As String = MakeToolButtonText(childPanel) 'mScript.mainClass(childPanel.classId).Names(0) + " - " + mScript.PrepareStringToPrint(pName, Nothing, False)
        'Создание самой вкладки
        Dim btn As ToolStripButton = New ToolStripButton With {.AutoSize = True, .Checked = True, .CheckOnClick = False, .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText, _
                                                                .ImageAlign = ContentAlignment.MiddleLeft, .TextAlign = ContentAlignment.MiddleLeft, _
                                                               .TextImageRelation = TextImageRelation.ImageBeforeText, .Text = strText, .ImageScaling = ToolStripItemImageScaling.SizeToFit, .DoubleClickEnabled = True}

        SetToobButtonImage(btn, childPanel) 'set tool button picture
        childPanel.toolButton = btn
        'Обработчики вкладки
        Dim chCopy As clsChildPanel = childPanel
        AddHandler btn.Paint, AddressOf del_ToolButton_Paint
        AddHandler btn.DoubleClick, Sub(sender As Object, e As EventArgs)
                                        If IsNothing(chCopy) OrElse Object.Equals(chCopy, ActivePanel) = False Then Return
                                        For i As Integer = Me.lstPanels.Count - 1 To 0 Step -1
                                            Dim actCh As clsChildPanel = lstPanels(i)
                                            If Object.Equals(actCh, chCopy) Then Continue For
                                            Me.RemovePanel(actCh, False)
                                        Next
                                    End Sub

        AddHandler btn.MouseDown, Sub(sender As Object, e As MouseEventArgs)
                                      'Выбор вкладки
                                      'При выборе вкладки не срабатывает события Validating, поэтому надо вызвать принудительно
                                      If frmMainEditor.codeBoxPanel.Visible Then
                                          'Validating кодбокса
                                          frmMainEditor.codeBox.codeBox.CheckTextForSyntaxErrors(frmMainEditor.codeBox.codeBox)
                                          Call frmMainEditor.codeBox_Validating(frmMainEditor.codeBox, New System.ComponentModel.CancelEventArgs)
                                          If mScript.LAST_ERROR.Length > 0 Then Return
                                          frmMainEditor.codeBoxChangeOwner(Nothing)
                                      ElseIf IsNothing(ActivePanel) = False AndAlso IsNothing(ActivePanel.ActiveControl) = False AndAlso ActivePanel.ActiveControl.Focused Then
                                          'Validating свойства
                                          Dim tName As String
                                          tName = ActivePanel.ActiveControl.GetType.Name
                                          If tName <> "ButtonEx" Then
                                              Dim ee As New System.ComponentModel.CancelEventArgs
                                              Call del_propertyValidating(ActivePanel.ActiveControl, ee)
                                              If ee.Cancel Then Return
                                          End If
                                      End If
                                      Dim imgSize As Integer = 12 ' btn.Image.Width
                                      Dim r As New Rectangle(btn.Bounds.Width - imgSize, 0, imgSize, imgSize)
                                      'If currentClassName = "L" AndAlso IsNothing(ActivePanel) = False AndAlso ActivePanel.child2Id > -1 Then actionsRouter.SaveActions(ActivePanel.child2Id)
                                      If r.Contains(e.X, e.Y) Then
                                          isCrossUnderMouse = False
                                          RemovePanel(chCopy) 'нажатие а крестик - удаление вкладки
                                      Else
                                          actionsRouter.UpdateActions(chCopy.classId, IIf(chCopy.supraElementName.Length = 0, chCopy.child2Name, chCopy.supraElementName))
                                          If btn.Checked Then Return
                                          'нажатие на вкладку (мимо крестика) - показываем вкладку
                                          If chCopy.classId = -2 Then
                                              ShowPanelVariables(chCopy)
                                          ElseIf chCopy.classId = -3 Then
                                              ShowPanelFunctions(chCopy)
                                          Else
                                              ShowPanel(chCopy)
                                          End If
                                      End If
                                      UpdateToolButtonsColors(ActivePanel)
                                  End Sub

        AddHandler btn.MouseLeave, Sub(sender As Object, e As EventArgs)
                                       If isCrossUnderMouse Then
                                           'btn.Image = My.Resources.cross_grey 'меняем картинку крестика на невыделенную
                                           isCrossUnderMouse = False
                                           'del_ToolButton_Paint(btn, New PaintEventArgs(btn.cre)
                                           btn.Invalidate()
                                       End If
                                   End Sub

        AddHandler btn.MouseMove, Sub(sender As Object, e As MouseEventArgs)
                                      Dim imgSize As Integer = 12 'btn.Image.Width
                                      Dim r As New Rectangle(btn.Bounds.Width - imgSize, 0, imgSize, imgSize)
                                      If r.Contains(e.X, e.Y) Then
                                          If Not isCrossUnderMouse Then
                                              'btn.Image = My.Resources.cross_red
                                              isCrossUnderMouse = True 'мышь над крестиком
                                              btn.Invalidate()
                                          End If
                                      Else
                                          If isCrossUnderMouse Then
                                              'btn.Image = My.Resources.cross_grey 'мышь не над крестиком
                                              isCrossUnderMouse = False
                                              btn.Invalidate()
                                          End If
                                      End If
                                  End Sub
        panelToolStrip.Items.Add(btn)
        UpdateToolButtonsColors(childPanel)
    End Sub

    Public Sub del_ToolButton_Paint(sender As Object, e As PaintEventArgs)
        Dim g As Graphics = e.Graphics
        g.DrawImage(IIf(isCrossUnderMouse, My.Resources.cross_red, My.Resources.cross_grey), sender.Width - My.Resources.cross_grey.Width - 2, 2, 8, 8)
    End Sub

    ''' <summary>
    ''' Sets the picture of ToolButton depending on current ClassName (L - location picture, O - Object...)
    ''' </summary>
    ''' <param name="btn">ToolButton to set the picture</param>
    ''' <param name="childPanel">Child panel related to this tool button</param>
    ''' <remarks></remarks>
    Private Sub SetToobButtonImage(ByRef btn As ToolStripButton, ByRef childPanel As clsChildPanel)
        Dim classId As Integer = childPanel.classId
        Dim picName As String = ""
        If classId = -2 Then
            'variable
            picName = "V.png"
        ElseIf classId = -3 Then
            'function
            picName = "F.png"
        ElseIf classId = -1 Then
            'error
            'set the default image
            btn.Image = My.Resources.DefToolButton
        Else
            If String.IsNullOrEmpty(childPanel.child3Name) Then
                'second level
                picName = mScript.mainClass(classId).Names(0) + ".png"
            Else
                'third level
                picName = mScript.mainClass(classId).Names(0) + "_sub.png"
            End If
        End If

        'check picture in the Quest directory
        Dim picPath As String = My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "img\ToolButtons\" & picName)
        If My.Computer.FileSystem.FileExists(picPath) Then
            'set the picture from the quest dir
            btn.Image = Image.FromFile(picPath)
        Else
            'check picture in the program default directory
            picPath = My.Computer.FileSystem.CombinePath(ProgramPath, "src\img\ToolButtons\" & picName)
            If My.Computer.FileSystem.FileExists(picPath) Then
                'set the picture from the program dir
                btn.Image = Image.FromFile(picPath)
            Else
                'set the default image
                btn.Image = My.Resources.DefToolButton
            End If
        End If
    End Sub

    ''' <summary>
    ''' Находит панель с указанными характеристиками или создает новую, если таковой не существовало; выделяет контрол, ассоциированный с указанным элементом; 
    ''' если в элементе содержится содержится скрипт - выделяем указанное слово (или всю строку)
    ''' </summary>
    ''' <param name="classId">Класс открываемой панели</param>
    ''' <param name="child2Id">Элемент 2 порядка открываемой панели</param>
    ''' <param name="child3Id">Элемент 3 порядка открываемой панели</param>
    ''' <param name="supraId">Родительский элемент для элемента 2 порядка</param>
    ''' <param name="elementType">Тип - свойство/функция (ПОКА НЕ ИСПОЛЬЗУЕТСЯ)</param>
    ''' <param name="propName">Имя свойства/функции</param>
    ''' <param name="lineId">Номер линии для выделения, от 0</param>
    ''' <param name="word">Выделяемое слово</param>
    ''' <param name="wordStart">Начало выделяемого слова относительно начала строки</param>
    ''' <returns></returns>
    Public Function FindAndOpen(ByVal classId As Integer, ByVal child2Id As Integer, ByVal child3Id As Integer, ByVal supraId As Integer, _
                                ByVal elementType As CodeTextBox.EditWordTypeEnum, Optional ByVal propName As String = "", Optional ByVal lineId As Integer = -1, _
                                Optional ByVal word As String = "", Optional ByVal wordStart As Integer = -1, Optional tracking As frmMainEditor.trackingcodeEnum = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT) As clsChildPanel
        Dim child2Name As String = ""
        Dim child3Name As String = ""
        Dim supraName As String = ""
        If classId = -2 Then
            'Переменные
            child2Name = mScript.csPublicVariables.lstVariables.ElementAt(child2Id).Key
        ElseIf classId = -3 Then
            'Функции
            child2Name = mScript.functionsHash.ElementAt(child2Id).Key
        Else
            'Элементы
            If child2Id > -1 Then
                If child3Id > -1 Then
                    'элементы 3 уровня
                    child3Name = mScript.mainClass(classId).ChildProperties(child2Id)("Name").ThirdLevelProperties(child3Id)
                    child2Name = mScript.mainClass(classId).ChildProperties(child2Id)("Name").Value
                Else
                    'элементы 2 уровня
                    If classId = mScript.mainClassHash("A") Then
                        If supraId > -1 Then supraName = mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(supraId)("Name").Value
                        If actionsRouter.GetActiveLocationId <> supraId Then
                            Dim lstAct As List(Of String) = actionsRouter.GetActionsNames(supraId)
                            If child2Id > lstAct.Count - 1 Then Return Nothing
                            child2Name = lstAct(child2Id)
                        Else
                            child2Name = mScript.mainClass(classId).ChildProperties(child2Id)("Name").Value
                        End If
                    Else
                        child2Name = mScript.mainClass(classId).ChildProperties(child2Id)("Name").Value
                    End If
                End If
            End If
        End If

        Dim panelId As Integer = GetPanelId(classId, child2Name, child3Name, supraName)
        Dim ch As clsChildPanel = Nothing

        If panelId > -1 Then
            'панель уже создана. Открываем ее
            ch = OpenPanel(lstPanels(panelId))
        Else
            'панель еще не создана. Создаем ее
            Dim tree As TreeView = Nothing
            Dim nod As TreeNode = Nothing
            'Ищем узел дерева, связанный с данным элементом
            If classId = -2 Then
                'переменные
                tree = frmMainEditor.treeVariables
                Dim varName As String = ""
                Try
                    varName = mScript.csPublicVariables.lstVariables.ElementAt(child2Id).Key
                Catch ex As Exception
                    Return Nothing
                End Try
                nod = frmMainEditor.FindItemNodeByText(tree, varName)
                If IsNothing(nod) Then Return Nothing
            ElseIf classId = -3 Then
                'функции
                tree = frmMainEditor.treeFunctions
                Dim fName As String = ""
                Try
                    fName = mScript.functionsHash.ElementAt(child2Id).Key
                Catch ex As Exception
                    Return Nothing
                End Try
                nod = frmMainEditor.FindItemNodeByText(tree, fName)
                If IsNothing(nod) Then Return Nothing
            ElseIf child2Id < 0 Then
                'свойства по умолчанию элементов. Тут же все классы 1-го уровня
                nod = Nothing
            ElseIf child2Id >= 0 AndAlso child3Id < 0 Then
                'элементы 2 уровня 
                Try
                    tree = frmMainEditor.dictTrees(classId)
                    If supraId >= 0 Then frmMainEditor.ChangeCurrentClass(mScript.mainClass(classId).Names(0), False, supraName)
                    Dim chName As String = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("Name").Value, Nothing, False)
                    nod = frmMainEditor.FindItemNodeByText(tree, chName)
                    If IsNothing(nod) Then Return Nothing
                Catch ex As Exception
                    Return Nothing
                End Try
            Else
                'Элементы 3 порядка
                Try
                    tree = frmMainEditor.dictTrees(classId)
                    Dim chName As String = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("Name").ThirdLevelProperties(child3Id), Nothing, False)
                    frmMainEditor.ChangeCurrentClass(mScript.mainClass(classId).Names(0), True, child2Name, chName)
                    nod = frmMainEditor.FindItemNodeByText(tree, chName)
                    If IsNothing(nod) Then Return Nothing
                Catch ex As Exception
                    Return Nothing
                End Try
            End If

            'Создаем панель
            ch = OpenPanel(nod, classId, child2Name, child3Name, supraName)
        End If
        If IsNothing(ch) Then Return Nothing
        If String.IsNullOrEmpty(propName) Then Return ch
        'Переходим к указанному элементу, выделяем указанную строку в кодбоксе
        If classId = -2 Then
            'Переменные
            If child3Id >= 0 AndAlso frmMainEditor.dgwVariables.Visible Then
                frmMainEditor.dgwVariables.Item(1, child3Id).Selected = True
            End If
        ElseIf classId = -3 Then
            'Функции
            If lineId >= 0 AndAlso frmMainEditor.codeBox.Visible Then
                With frmMainEditor.codeBox.codeBox
                    .Focus()
                    If lineId > .Lines.Length - 1 Then Return ch
                    Dim sStart As Integer = .GetFirstCharIndexFromLine(lineId)
                    Dim sLength As Integer = 0
                    If wordStart >= 0 AndAlso String.Compare(.Text.Substring(sStart + wordStart, word.Length), word, True) = 0 Then
                        sStart += wordStart
                        sLength = word.Length
                    Else
                        sLength = .Lines(lineId).Length
                    End If
                    .Select(sStart, sLength)
                End With
            End If
        Else
            'Элементы
            Dim pnl As PanelEx = dictDefContainers(classId)
            Dim arrC() As Control = pnl.Controls.Find(propName, True)
            If arrC.Count = 0 Then Return ch
            Dim tName As String = arrC(0).GetType.Name
            If tName = "ButtonEx" Then
                del_propertyButtonMouseClick(arrC(0), New EventArgs)
            ElseIf tName = "TextBoxEx" OrElse tName = "ComboBoxEx" Then
                If tracking = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT Then
                    del_TextPropertyMouseDown(arrC(0), New MouseEventArgs(MouseButtons.Left, 1, 1, 1, 0))
                    arrC(0).Focus()
                ElseIf tracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE Then
                    mScript.trackingProperties.LoadEventBeforeToCodeBox(arrC(0))
                ElseIf tracking = frmMainEditor.trackingcodeEnum.EVENT_AFTER Then
                    mScript.trackingProperties.LoadEventAfterToCodeBox(arrC(0))
                End If
            End If
            If lineId >= 0 AndAlso frmMainEditor.codeBox.Visible Then
                With frmMainEditor.codeBox.codeBox
                    .Focus()
                    If lineId > .Lines.Length - 1 Then Return ch
                    Dim sLength As Integer = 0
                    Dim sStart As Integer = .GetFirstCharIndexFromLine(lineId)
                    Dim strText As String = .Text
                    If wordStart >= 0 AndAlso String.Compare(strText.Substring(sStart + wordStart, word.Length), word, True) = 0 Then
                        sStart += wordStart
                        sLength = word.Length
                    Else
                        sLength = .Lines(lineId).Length
                    End If
                    .Select(sStart, sLength)
                End With
            End If
        End If
        Return ch
    End Function

    ''' <summary>
    ''' Открывает окно-панель для описания какого-либо элемента или создает новую, если такой еще не было
    ''' </summary>
    ''' <param name="treeNode">Узел дерава, ассоциированный с открываемой панелью</param>
    ''' <param name="classId">Id класса в MainClassType</param>
    ''' <param name="child2Name">Имя элемента 2-го порядка в списке ChildProperties() (может быть Id.ToString)</param>
    ''' <param name="child3Name">Имя элемента 3-го порядка в списке ChildProperties().ThirdLevelProperties (может быть Id.ToString)</param>
    ''' <param name="supraName">Имя элемента, к которому относится данный элемент (например, Id локации, если в child2Id - Id его действия. При этом child3Id должен быть равен -1)</param>
    ''' <param name="shouldFillTree">Должно ли обновляться дерево при смене класса</param>
    ''' <returns>Ссылку на открываемую панель</returns>
    Public Overridable Function OpenPanel(ByRef treeNode As TreeNode, ByVal classId As Integer, ByVal child2Name As String, Optional ByVal child3Name As String = "", _
                              Optional ByVal supraName As String = "", Optional shouldFillTree As Boolean = True) As clsChildPanel
        AAA()
        If questEnvironment.EnabledEvents = False Then Return ActivePanel
        If classId = mScript.mainClassHash("A") Then actionsRouter.UpdateActions(classId, supraName)
        If classId <> mScript.mainClassHash("A") Then actionsRouter.UpdateActions(classId, IIf(supraName.Length = 0, child2Name, supraName))

        Dim panelId As Integer = GetPanelId(classId, child2Name, child3Name, supraName)
        Dim chld As clsChildPanel
        If panelId = -1 Then
            'панель еще не создана - создаем новую
            chld = New clsChildPanel(classId, child2Name, child3Name, treeNode, supraName)
            lstPanels.Add(chld)
            CreateToolButton(chld) 'создаем закладку для перехода/закрытия окна
        Else
            'панель создана - открываем ее            
            chld = lstPanels(panelId)
            If Object.Equals(chld, ActivePanel) AndAlso ActivePanel.toolButton.Checked Then Return chld
        End If

        Log.PrintToLog("OpenPanel: " + Log.GetChildInfo(chld))
        If classId = -2 Then
            ShowPanelVariables(chld)
        ElseIf classId = -3 Then
            ShowPanelFunctions(chld)
        Else
            ShowPanel(chld, shouldFillTree)
        End If
        UpdateToolButtonsColors(chld)
        AAA()
        Return chld
    End Function

    ''' <summary>
    ''' Открывает окно-панель для описания какого-либо элемента (новую не создает)
    ''' </summary>
    ''' <param name="chld">открываемая панель</param>
    ''' <param name="shouldFillTree">Должно ли обновляться дерево при смене класса</param>
    ''' <returns>Ссылку на открываемую панель</returns>
    Public Overridable Function OpenPanel(ByRef chld As clsChildPanel, Optional shouldFillTree As Boolean = True) As clsChildPanel
        If questEnvironment.EnabledEvents = False Then Return ActivePanel
        'панель создана - открываем ее            
        actionsRouter.UpdateActions(chld.classId, IIf(chld.supraElementName.Length = 0, chld.child2Name, chld.supraElementName))
        If Object.Equals(chld, ActivePanel) AndAlso ActivePanel.toolButton.Checked Then Return chld
        Log.PrintToLog("OpenPanel: " + Log.GetChildInfo(chld))
        If chld.classId = -2 Then
            ShowPanelVariables(chld)
        ElseIf chld.classId = -3 Then
            ShowPanelFunctions(chld)
        Else
            ShowPanel(chld, shouldFillTree)
        End If
        UpdateToolButtonsColors(chld)
        Return chld
    End Function

    ''' <summary>
    ''' В случае если предыдущая панель и новая из разных классов, то выполняет все необходимые изменения в структуре и отображении на форме.
    ''' Строго говоря, не всегда когда преобразования нужны происходит именно смена класса. Это может быть, например, переход от одного подменю к другому, если они от разных меню-родителей.
    ''' </summary>
    ''' <param name="childPanel">Выбранная вкладка, на которую совершается переход</param>
    ''' <param name="shouldFillTree">Должно ли обновляться дерево при смене класса</param>
    Private Sub ShowWithClassChanging(ByRef childPanel As clsChildPanel, Optional shouldFillTree As Boolean = True)
        AAA()
        'Собираем все характеристки предыдущего состояния
        Dim oldClassId As Integer
        If currentClassName = "Variable" Then
            oldClassId = -2
        ElseIf currentClassName = "Function" Then
            oldClassId = -3
        Else
            If String.IsNullOrEmpty(currentClassName) Then
                oldClassId = -1
            Else
                oldClassId = mScript.mainClassHash(currentClassName)
            End If
        End If
        Dim oldClassName As String = currentClassName
        Dim oldParentName As String = currentParentName
        Dim oldDefPanel As Boolean = False
        Dim oldChild2Name As String = ""
        Dim oldThirdLevel As Boolean = False
        If IsNothing(ActivePanel) = False Then
            If ActivePanel.child3Name.Length > 0 Then oldThirdLevel = True
            oldChild2Name = ActivePanel.child2Name
            If oldChild2Name.Length = 0 Then oldDefPanel = True
        ElseIf oldClassId >= 0 AndAlso frmMainEditor.IsThirdLevelTree(currentTreeView) Then
            oldThirdLevel = True
            oldChild2Name = currentParentName
            If oldChild2Name.Length = 0 Then oldDefPanel = True
        End If
        'Собираем все характеристики нового состояния
        Dim newClassId As Integer = childPanel.classId
        Dim newClassName As String = ""
        If newClassId = -2 Then
            newClassName = "Variable"
        ElseIf newClassId = -3 Then
            newClassName = "Function"
        Else
            newClassName = mScript.mainClass(childPanel.classId).Names(0)
        End If
        Dim newParentName As String = childPanel.supraElementName
        Dim newDefPanel As Boolean = (childPanel.child2Name.Length = 0)
        Dim newThirdLevel As Boolean = (childPanel.child3Name.Length > 0)
        Dim newChild2Name As String = childPanel.child2Name

        'Блок проверки на необходимость преобразований. Если это простой переход без смены класса и родителей, то выход
        If oldClassId = newClassId Then
            'Переходы внутри одного класса.
            If oldDefPanel AndAlso newDefPanel = False Then
                'переход от свойств по умолчанию (СПУ) к обычному
                If newParentName <> oldParentName Then
                    'fromDefPanel = True 'Смена должна произойти, так как были отображены действия от другой локации
                ElseIf newThirdLevel AndAlso oldChild2Name <> newChild2Name Then
                    'fromDefPanel = True 'Смена должна произойти, так как были отображены элементы 3 порядка от другого родителя
                Else
                    'переход от СПУ к обычному элементу - смена не происходит
                    Return
                End If
            ElseIf oldDefPanel = False AndAlso newDefPanel Then
                'переход от обычного элемента с СПУ - смена не происходит
                Return
            ElseIf oldDefPanel = False AndAlso newDefPanel = False Then
                'переход без вовлечения СПУ
                If newParentName <> oldParentName AndAlso (newThirdLevel = False AndAlso oldThirdLevel = False) Then
                    'переход от одного действия к другому, но от другой локации
                    'SameFromDifSupra = True
                ElseIf newThirdLevel AndAlso oldThirdLevel Then
                    'переход от элемента 3 порядка к другому такому же
                    If oldChild2Name <> newChild2Name Then
                        'SameFromDifParent = True 'элементы 3 порядка от разных родителей
                    Else
                        'элементы 3 порядка от одного родителя - смены не происходит
                        Return
                    End If
                ElseIf oldThirdLevel AndAlso newThirdLevel = False Then
                    'переход от элемента 3 порядка к элементу 2 порядка того же класса
                    'fromThirdLevel = True
                ElseIf oldThirdLevel = False AndAlso newThirdLevel Then
                    'переход от элемента 2 порядка к элементу 3 порядка того же класса
                    'toThirdLevel = True
                Else
                    Return
                End If
            ElseIf oldDefPanel AndAlso newDefPanel Then
                'oldDefPanel = newDefPanel - такого быть не должно (переход от панели по умолчанию к себе самой)
                MsgBox("Переход от панели по умолчанию к себе самой.", vbExclamation)
                Return
            End If
        ElseIf oldClassId <> newClassId Then
            'переход между элементами из разных классов
        End If

        frmMainEditor.CheckClassButton(oldClassName, newClassId)

        'Смена класса/родителей. Преобразования необходимы
        questEnvironment.EnabledEvents = False
        'сохраняем последний выбранный элемент в старом классе (для возможности возврата именно к нему, если надо будет перейти к старому классу без указания конкретного элемента)
        'для элементов 3, 1 уровня и действий необходимости в этом нет
        Dim prevPanel As clsPanelManager.clsChildPanel = Nothing
        If oldClassId <= -1 OrElse (oldThirdLevel = False AndAlso oldParentName.Length = 0 AndAlso mScript.mainClass(oldClassId).LevelsCount > 0) Then
            prevPanel = cPanelManager.SetLastPanel(oldClassName)
        End If

        'Получаем Name родительского элемента, если он не был получен явно
        If newChild2Name.Length = 0 Then
            'если совершен переход на СПУ:
            If oldClassId <> -1 AndAlso GetParentClassId(newClassName) = oldClassId Then
                'совершен переход от локации к СПУ действий - currentParentForThirdLevelId = Id старой локации
                newParentName = oldChild2Name
                If oldChild2Name.Length = 0 Then
                    'совершен переход от СПУ локаций к СПУ действий. Действия которой локации отображать - неизвестно
                    'пробуем получить список из предыдущей панели
                    If IsNothing(prevPanel) = False AndAlso prevPanel.child2Name.Length > 0 Then
                        newParentName = prevPanel.child2Name
                    Else
                        'предыдущая панель тоже была СПУ. Тогда просто берем список из первой локации
                        newParentName = mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(0)("Name").Value
                    End If
                    actionsRouter.RetreiveActions(newParentName)
                Else
                    actionsRouter.RetreiveActions(newParentName)
                End If
            ElseIf newDefPanel AndAlso newClassId = mScript.mainClassHash("A") Then
                'совершен переход НЕ от локации к СПУ действий - пробуем получить последнюю панель локаций
                Dim classL As Integer = mScript.mainClassHash("L")
                If dictLastPanel.ContainsKey(classL) AndAlso IsNothing(dictLastPanel(classL)) = False Then
                    newParentName = dictLastPanel(classL).child2Name
                End If
                If newParentName.Length = 0 Then
                    newParentName = actionsRouter.locationOfCurrentActions
                End If
                If newParentName.Length = 0 Then
                    newParentName = mScript.mainClass(classL).ChildProperties(0)("Name").Value
                End If
                actionsRouter.RetreiveActions(newParentName)
            End If
        End If
        Log.PrintToLog("ChangeClassAfterPanelSelection (old): " + Log.GetChildInfo(oldClassId, oldChild2Name, oldParentName, -1))
        Log.PrintToLog("ChangeClassAfterPanelSelection (new): " + Log.GetChildInfo(newClassId, newChild2Name, newParentName, -1))

        'If oldClassName = "A" Then actionsRouter.SaveActions(oldParentId)

        'установка глобальных переменных состояния
        currentClassName = newClassName
        currentParentName = newParentName
        'настройка отображения дерева и разных панелей
        Dim classLevels As Integer = 1
        Dim tree As TreeView = Nothing
        'If oldClassId < 0 OrElse mScript.mainClass(oldClassId).LevelsCount > 0 Then currentTreeView.Hide()
        If (oldClassId < 0 OrElse frmMainEditor.dictTrees.ContainsKey(oldClassId)) AndAlso IsNothing(currentTreeView) = False Then currentTreeView.Hide()

        If newClassId >= 0 Then
            frmMainEditor.btnShowSettings.Enabled = True
            If newThirdLevel OrElse newParentName.Length > 0 Then
                frmMainEditor.btnMakeHidden.Visible = False
                frmMainEditor.chkShowHidden.Visible = False
                frmMainEditor.btnCopyTo.Visible = True
            Else
                frmMainEditor.btnMakeHidden.Visible = True
                frmMainEditor.chkShowHidden.Visible = True
                frmMainEditor.btnCopyTo.Visible = False
            End If
            classLevels = mScript.mainClass(newClassId).LevelsCount
            If classLevels > 0 Then
                'tree = frmMainEditor.dictTrees(newClassId)
                frmMainEditor.dictTrees.TryGetValue(newClassId, tree)
                If IsNothing(tree) = False Then
                    currentTreeView = tree
                    frmMainEditor.splitTreeContainer.Show()
                    tree.Show()
                    frmMainEditor.SplitOuter.Panel1Collapsed = False
                End If
            Else
                currentTreeView = Nothing
                frmMainEditor.splitTreeContainer.Hide()
                frmMainEditor.SplitOuter.Panel1Collapsed = True
            End If
            frmMainEditor.SplitInner.Panel1Collapsed = False
        ElseIf newClassId = -2 Then
            frmMainEditor.btnShowSettings.Enabled = False
            frmMainEditor.btnMakeHidden.Visible = True
            frmMainEditor.chkShowHidden.Visible = True
            frmMainEditor.btnCopyTo.Visible = False
            tree = frmMainEditor.treeVariables
            currentTreeView = tree
            frmMainEditor.splitTreeContainer.Show()
            frmMainEditor.SplitOuter.Panel1Collapsed = False
            frmMainEditor.SplitInner.Panel1Collapsed = True
            tree.Show()
        ElseIf newClassId = -3 Then
            frmMainEditor.btnShowSettings.Enabled = False
            frmMainEditor.btnMakeHidden.Visible = True
            frmMainEditor.chkShowHidden.Visible = True
            frmMainEditor.btnCopyTo.Visible = False
            tree = frmMainEditor.treeFunctions
            currentTreeView = tree
            frmMainEditor.splitTreeContainer.Show()
            frmMainEditor.SplitInner.Panel1Collapsed = True
            frmMainEditor.SplitOuter.Panel1Collapsed = False
            tree.Show()
        End If

        'Восстанавливаем действия если новый класс - действия
        If shouldFillTree Then
            'заполняем дерево при необходмости (при смене родителя)
            If newThirdLevel Then
                frmMainEditor.FillTree(newClassName, tree, frmMainEditor.chkShowHidden.Checked, Nothing, newChild2Name)
            ElseIf newParentName.Length > 0 Then
                frmMainEditor.FillTree(newClassName, tree, frmMainEditor.chkShowHidden.Checked, Nothing, newParentName)
            ElseIf newParentName.Length = 0 AndAlso mScript.mainClass(newClassId).LevelsCount = 2 Then
                'перестраиваем дерево всем 3-уровневым классам
                frmMainEditor.FillTree(newClassName, tree, frmMainEditor.chkShowHidden.Checked, Nothing, "")
            End If
        End If

        questEnvironment.EnabledEvents = True
        If newChild2Name.Length > 0 Then
            'Выделяем нужный узел дерева
            Dim selectedChildName As String = ""
            If newClassId >= 0 Then
                selectedChildName = mScript.PrepareStringToPrint(newChild2Name, Nothing, False)
            ElseIf newClassId < 0 Then
                selectedChildName = newChild2Name
            End If
            Dim n As TreeNode = frmMainEditor.FindItemNodeByText(tree, selectedChildName)
            If IsNothing(n) = False Then
                If IsNothing(childPanel.treeNode) Then childPanel.treeNode = n
                questEnvironment.EnabledEvents = False
                If Object.Equals(tree.SelectedNode, n) Then
                    If newClassId = -2 Then
                        Call frmMainEditor.tree_AfterSelectVariables(tree, New TreeViewEventArgs(n))
                    ElseIf newClassId = -3 Then
                        Call frmMainEditor.tree_AfterSelectFunctions(tree, New TreeViewEventArgs(n))
                    Else
                        Call frmMainEditor.tree_AfterSelect(tree, New TreeViewEventArgs(n))
                    End If
                Else
                    tree.SelectedNode = n
                End If
                questEnvironment.EnabledEvents = True
            End If
        End If
        AAA()
    End Sub

    ''' <summary>
    ''' Выделяет панели данного класса
    ''' </summary>
    Public Sub UpdateToolButtonsColors(ByRef childPanel As clsChildPanel)
        If IsNothing(lstPanels) Or String.IsNullOrEmpty(currentClassName) Then Return
        Dim curClassId As Integer = -1
        If currentClassName = "Variable" Then
            curClassId = -2
        ElseIf currentClassName = "Function" Then
            curClassId = -3
        Else
            curClassId = mScript.mainClassHash(currentClassName)
        End If
        For i As Integer = 0 To lstPanels.Count - 1
            Dim ch As clsChildPanel = lstPanels(i)
            If IsNothing(ch.toolButton) Then Continue For
            'If IsNothing(childPanel) = False AndAlso Object.Equals(ch, childPanel) Then
            '    ch.toolButton.BackColor = DEFAULT_COLORS.NodeSelBackColor
            '    Continue For
            'End If

            Dim blnEligible As Boolean = False

            If ch.classId = curClassId Then
                If String.IsNullOrEmpty(ch.child3Name) Then
                    'i - second level class
                    If String.IsNullOrEmpty(currentParentName) Then
                        blnEligible = True
                    Else
                        blnEligible = False
                        If currentClassName = "A" Then
                            If ch.GetParentId = currentParentId Then
                                blnEligible = True
                            End If
                        End If
                    End If
                Else
                    'i - third level class
                    If String.IsNullOrEmpty(currentParentName) Then
                        blnEligible = False
                    Else
                        If ch.GetParentId = currentParentId Then
                            blnEligible = True
                        Else
                            blnEligible = False
                        End If
                    End If
                End If
            End If

            If blnEligible Then
                ch.toolButton.BackColor = DEFAULT_COLORS.SameClassToolButton
            Else
                ch.toolButton.BackColor = DEFAULT_COLORS.ControlTransparentBackground
            End If
        Next

        If currentClassName = "A" Then
            'highlight parent Location
            Dim classL As Integer = mScript.mainClassHash("L")
            For i = 0 To lstPanels.Count - 1
                Dim ch As clsChildPanel = lstPanels(i)
                If ch.classId = classL AndAlso String.IsNullOrEmpty(ch.child3Name) AndAlso currentParentName = ch.child2Name Then
                    ch.toolButton.BackColor = DEFAULT_COLORS.RelatedToolButton
                    Exit For
                End If
            Next
        ElseIf currentClassName = "L" Then
            'highlight children actions
            Dim classA As Integer = mScript.mainClassHash("A")
            If IsNothing(childPanel) Then Return
            Dim curChild2Id As Integer = childPanel.GetChild2Id
            For i = 0 To lstPanels.Count - 1
                Dim ch As clsChildPanel = lstPanels(i)
                If ch.classId = classA AndAlso ch.GetParentId = curChild2Id Then
                    ch.toolButton.BackColor = DEFAULT_COLORS.RelatedToolButton
                End If
            Next
        ElseIf currentParentId > -1 Then
            'Highlight parent element
            For i = 0 To lstPanels.Count - 1
                Dim ch As clsChildPanel = lstPanels(i)
                If ch.classId = curClassId AndAlso String.IsNullOrEmpty(ch.child3Name) AndAlso currentParentName = ch.child2Name Then
                    ch.toolButton.BackColor = DEFAULT_COLORS.RelatedToolButton
                    Exit For
                End If
            Next
        ElseIf currentParentId < 0 AndAlso curClassId >= 0 AndAlso mScript.mainClass(curClassId).LevelsCount = 2 Then
            'Highlight children elements
            If IsNothing(childPanel) Then Return
            Dim curChild2Name As String = childPanel.child2Name
            For i = 0 To lstPanels.Count - 1
                Dim ch As clsChildPanel = lstPanels(i)
                If ch.classId = curClassId AndAlso String.IsNullOrEmpty(ch.child3Name) = False AndAlso curChild2Name = ch.child2Name Then
                    ch.toolButton.BackColor = DEFAULT_COLORS.RelatedToolButton
                End If
            Next
        End If
    End Sub

    ''' <summary>Показывает панель (делает ее активной)</summary>
    ''' <param name="childPanel">Панель, которую надо открыть</param>
    ''' <param name="shouldFillTree">Должно ли обновляться дерево при смене класса</param>
    Private Sub ShowPanel(ByRef childPanel As clsChildPanel, Optional shouldFillTree As Boolean = True)
        AAA()
        If questEnvironment.EnabledEvents = False Then Return
        If IsNothing(childPanel) Then Return
        If Object.Equals(ActivePanel, childPanel) Then
            dictDefContainers(childPanel.classId).Visible = True
            childPanel.toolButton.Checked = True
            Return 'выход если панель уже текущая
        End If

        'убираем выделение со старой вкладки и ставим выделение на текущую
        Dim prevActiveControl As Control = Nothing
        If IsNothing(ActivePanel) = False Then
            prevActiveControl = ActivePanel.ActiveControl
            ActivePanel.toolButton.Checked = False
            If ActivePanel.classId <> childPanel.classId Then HidePanel(ActivePanel) 'прячем старую панель
        End If
        Log.PrintToLog("ShowPanel: " + Log.GetChildInfo(childPanel))

        ShowWithClassChanging(childPanel, shouldFillTree) 'при необходимости выполняет смену класса

        ActivePanel = childPanel   'делает панель активной
        'отображаем панель
        If FillProperties(childPanel) = -1 Then
            ActivePanel = Nothing
            AAA()
            Return
        End If
        dictDefContainers(childPanel.classId).Visible = True
        If IsNothing(childPanel.toolButton) Then CreateToolButton(childPanel)
        childPanel.toolButton.Checked = True

        clsChildPanel.DeselectControl(prevActiveControl) 'Убираем выделение с текущего контрола

        'Выделяем новый текущий узел дерева / кнопку настроек по умолчанию (если он не скрыт)
        Dim selNode As TreeNode = childPanel.treeNode
        If btnDefSettings.BackColor <> DEFAULT_COLORS.ControlTransparentBackground Then btnDefSettings.BackColor = DEFAULT_COLORS.ControlTransparentBackground
        If IsNothing(selNode) = False AndAlso IsNothing(selNode.TreeView) = False Then
            questEnvironment.EnabledEvents = False
            If Object.Equals(selNode.TreeView.SelectedNode, selNode) = False Then
                selNode.TreeView.SelectedNode = selNode
            Else
                Call frmMainEditor.tree_AfterSelect(selNode.TreeView, New TreeViewEventArgs(selNode))
            End If
            questEnvironment.EnabledEvents = True
        ElseIf childPanel.child2Name.Length = 0 AndAlso mScript.mainClass(childPanel.classId).LevelsCount > 0 Then
            questEnvironment.EnabledEvents = False
            Call frmMainEditor.btnShowSettings_Click(btnDefSettings, New EventArgs)
            questEnvironment.EnabledEvents = True
        End If

        'восстанавливаем внешний вид панели (какое свойство выбрано и что было открыто)
        'обновляем текущий контрол (выбранное свойство)
        prevActiveControl = ActivePanel.ActiveControl
        ActivePanel.lcActiveControl = Nothing
        ActivePanel.ActiveControl = prevActiveControl

        If IsNothing(prevActiveControl) Then
            WBhelp.Hide()
            frmMainEditor.codeBoxPanel.Hide()
            AAA()
            If currentClassName = "Map" AndAlso currentParentName.Length > 0 Then
                mapManager.BuildMapForCellsEdit(ActivePanel.child3Name)
            End If
            Return
        End If

        AAA()
        If IsNothing(ActivePanel.ActiveControl) = False AndAlso ActivePanel.ActiveControl.Name.Length = 0 Then
            If currentClassName = "Map" AndAlso currentParentName.Length > 0 Then
                mapManager.BuildMapForCellsEdit(ActivePanel.child3Name)
            End If
            Return
        End If

        Dim defProp As MatewScript.PropertiesInfoType
        Dim actControl As Object = ActivePanel.ActiveControl
        Dim isUserFunction As Boolean = actControl.IsFunctionButton
        Try
            If isUserFunction Then
                defProp = mScript.mainClass(ActivePanel.classId).Functions(actControl.Name)
            Else
                defProp = mScript.mainClass(ActivePanel.classId).Properties(actControl.Name)
            End If
        Catch ex As Exception
            ActivePanel.lcActiveControl.Name = ""
            ActivePanel.lcActiveControl = Nothing
            Return
        End Try

        Dim retType As MatewScript.ReturnFunctionEnum = defProp.returnType
        Dim helpFile As String = defProp.helpFile

        If currentClassName = "Map" AndAlso currentParentName.Length > 0 Then
            'клетки карты
            mapManager.BuildMapForCellsEdit(ActivePanel.child3Name)
        ElseIf ActivePanel.IsWbVisible Then
            'был открыт веб браузер
            frmMainEditor.codeBoxChangeOwner(Nothing) 'убираем старого хозяина с кодбокса дабы не возникло ошибок
            frmMainEditor.codeBoxPanel.Hide()
            If retType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION OrElse retType = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                'Свойство - событие/описание
                Dim btnHelp As ButtonEx = actControl.ButtonHelp
                If IsNothing(btnHelp) = False Then
                    del_propertyButtonHelpMouseClick(btnHelp, New EventArgs)
                Else
                    MessageBox.Show("Не удалось открыть помощь для свойства!")
                End If
            Else
                'Свойство - обычное
                WBhelp.Tag = Nothing
                ShowHelpFileForUsualProperty(ActivePanel, ActivePanel.ActiveControl.Name)
            End If
        Else
            'веб браузер был закрыт
            WBhelp.Hide()

            If retType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION OrElse retType = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                'Свойство - событие/описание
                frmMainEditor.codeBoxChangeOwner(ActivePanel.ActiveControl)
                frmMainEditor.codeBoxPanel.Show()
            ElseIf ActivePanel.IsCodeBoxVisible Then
                'Свойство со скриптом
                Dim cb As CodeTextBox = frmMainEditor.codeBox
                Dim cbPanel As Panel = frmMainEditor.codeBoxPanel
                cb.Tag = Nothing

                Dim child2Id As Integer = ActivePanel.GetChild2Id
                Dim child3Id As Integer = ActivePanel.GetChild3Id(child2Id)
                Dim propValue As String
                If child2Id < 0 Then
                    propValue = defProp.Value
                ElseIf child3Id < 0 Then
                    propValue = mScript.mainClass(ActivePanel.classId).ChildProperties(child2Id)(ActivePanel.ActiveControl.Name).Value
                Else
                    propValue = mScript.mainClass(ActivePanel.classId).ChildProperties(child2Id)(ActivePanel.ActiveControl.Name).ThirdLevelProperties(child3Id)
                End If
                Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(propValue)
                If cRes = MatewScript.ContainsCodeEnum.NOT_CODE Then
                    'Свойство - обычное
                    frmMainEditor.codeBoxChangeOwner(Nothing)
                    cbPanel.Hide()
                    Return
                ElseIf cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                    'исполняемая строка
                    cb.codeBox.IsTextBlockByDefault = False
                    cb.Text = mScript.PrepareStringToPrint(propValue, Nothing, False)
                Else
                    'код / длинный текст
                    cb.codeBox.IsTextBlockByDefault = (cRes = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    cb.codeBox.LoadCodeFromProperty(propValue)
                End If
                cb.Tag = ActivePanel.ActiveControl
                cbPanel.Show()
                cb.Refresh()
                cb.Focus()
            Else
                'Свойство - обычное
                frmMainEditor.codeBoxChangeOwner(Nothing)
                frmMainEditor.codeBoxPanel.Hide()
            End If
        End If
    End Sub


    ''' <summary>Показывает панель переменных(делает ее активной)</summary>
    ''' <param name="childPanel">Панель, которую надо открыть</param>
    Private Sub ShowPanelVariables(ByRef childPanel As clsChildPanel)
        If questEnvironment.EnabledEvents = False Then Return
        If IsNothing(childPanel) Then Return
        If Object.Equals(ActivePanel, childPanel) Then
            frmMainEditor.pnlVariables.Show()
            childPanel.toolButton.Checked = True
            Return 'выход если панель уже текущая
        End If

        'убираем выделение со старой вкладки и ставим выделение на текущую
        Dim prevActiveControl As Control = Nothing
        If IsNothing(ActivePanel) = False Then
            prevActiveControl = ActivePanel.ActiveControl
            ActivePanel.toolButton.Checked = False
            If ActivePanel.classId <> childPanel.classId Then HidePanel(ActivePanel) 'прячем старую панель
        End If
        Log.PrintToLog("ShowPanel: " + Log.GetChildInfo(childPanel))

        ShowWithClassChanging(childPanel, False) 'при необходимости выполняет смену класса
        ActivePanel = childPanel   'делает панель активной
        'отображаем панель
        If FillPropertiesVariables(childPanel) = -1 Then
            ActivePanel = Nothing
            Return
        End If
        frmMainEditor.pnlVariables.Show()
        childPanel.toolButton.Checked = True

        clsChildPanel.DeselectControl(prevActiveControl) 'Убираем выделение с текущего контрола

        'Выделяем новый текущий узел дерева / кнопку настроек по умолчанию (если он не скрыт)
        Dim selNode As TreeNode = childPanel.treeNode
        If btnDefSettings.BackColor <> DEFAULT_COLORS.ControlTransparentBackground Then btnDefSettings.BackColor = DEFAULT_COLORS.ControlTransparentBackground
        If IsNothing(selNode) = False AndAlso IsNothing(selNode.TreeView) = False Then
            questEnvironment.EnabledEvents = False
            If Object.Equals(selNode.TreeView.SelectedNode, selNode) = False Then
                selNode.TreeView.SelectedNode = selNode
            Else
                Call frmMainEditor.tree_AfterSelect(selNode.TreeView, New TreeViewEventArgs(selNode))
            End If
            questEnvironment.EnabledEvents = True
        End If

        'восстанавливаем внешний вид панели (какое свойство выбрано и что было открыто)
        'обновляем текущий контрол (выбранное свойство)
        WBhelp.Hide()
        frmMainEditor.codeBoxPanel.Hide()
    End Sub

    ''' <summary>Показывает панель функций (делает ее активной)</summary>
    ''' <param name="childPanel">Панель, которую надо открыть</param>
    Private Sub ShowPanelFunctions(ByRef childPanel As clsChildPanel)
        If questEnvironment.EnabledEvents = False Then Return
        If IsNothing(childPanel) Then Return
        If Object.Equals(ActivePanel, childPanel) Then
            frmMainEditor.trakingEventState = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT
            frmMainEditor.codeBoxPanel.Show()
            childPanel.toolButton.Checked = True
            Return 'выход если панель уже текущая
        End If

        'убираем выделение со старой вкладки и ставим выделение на текущую
        Dim prevActiveControl As Control = Nothing
        If IsNothing(ActivePanel) = False Then
            prevActiveControl = ActivePanel.ActiveControl
            ActivePanel.toolButton.Checked = False
            If ActivePanel.classId <> childPanel.classId Then HidePanel(ActivePanel) 'прячем старую панель
        End If
        Log.PrintToLog("ShowPanel: " + Log.GetChildInfo(childPanel))

        ShowWithClassChanging(childPanel, False) 'при необходимости выполняет смену класса
        ActivePanel = childPanel   'делает панель активной
        childPanel.toolButton.Checked = True

        clsChildPanel.DeselectControl(prevActiveControl) 'Убираем выделение с текущего контрола

        'Выделяем новый текущий узел дерева / кнопку настроек по умолчанию (если он не скрыт)
        Dim selNode As TreeNode = childPanel.treeNode
        'отображаем панель
        With frmMainEditor.codeBox.codeBox
            .Tag = Nothing
            .IsTextBlockByDefault = False
            .LoadCodeFromCodeData(mScript.functionsHash(childPanel.child2Name).ValueDt)
        End With
        frmMainEditor.codeBox.Tag = childPanel
        frmMainEditor.trakingEventState = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT
        frmMainEditor.codeBoxPanel.Show()

        If btnDefSettings.BackColor <> DEFAULT_COLORS.ControlTransparentBackground Then btnDefSettings.BackColor = DEFAULT_COLORS.ControlTransparentBackground
        If IsNothing(selNode) = False AndAlso IsNothing(selNode.TreeView) = False Then
            questEnvironment.EnabledEvents = False
            If Object.Equals(selNode.TreeView.SelectedNode, selNode) = False Then
                selNode.TreeView.SelectedNode = selNode
            Else
                Call frmMainEditor.tree_AfterSelect(selNode.TreeView, New TreeViewEventArgs(selNode))
            End If
            questEnvironment.EnabledEvents = True
        End If

        'восстанавливаем внешний вид панели (какое свойство выбрано и что было открыто)
        'обновляем текущий контрол (выбранное свойство)
        WBhelp.Hide()
    End Sub

    ''' <summary>
    ''' Возвращает панель по ассоциированному с ней узлу дерева
    ''' </summary>
    ''' <param name="n">узел для поиска</param>
    Public Function GetPanelByNode(ByRef n As TreeNode) As clsChildPanel
        If IsNothing(lstPanels) Then Return Nothing
        For i As Integer = 0 To lstPanels.Count - 1
            If Object.Equals(lstPanels(i).treeNode, n) Then Return lstPanels(i)
        Next
        Return Nothing
    End Function

    ''' <summary>Возвращает панель по ее данным</summary>
    ''' <param name="classId">Класс панели</param>
    ''' <param name="child2Name">Имя элемента второго порядка (или -1 если панель для свойств по умолчанию)</param>
    ''' <param name="child3Name">Имя элемента третьего порядка (или -1 если панель для свойств элемента 2-го порядка или свойств по умолчанию)</param>
    ''' <param name="supraName">Имя элемента, к которому относится данный элемент (например, Id локации, если в child2Name - имя его действия. При этом child3Id должен быть равен -1)</param>
    Public Function GetPanelByChildInfo(ByVal classId As Integer, ByVal child2Name As String, ByVal child3Name As String, ByVal supraName As String) As clsChildPanel
        Dim chId As Integer = GetPanelId(classId, child2Name, child3Name, supraName)
        If chId > -1 Then Return lstPanels(chId)
        Return Nothing
    End Function

    ''' <summary>Убирает панель (делает ее неактивной)</summary>
    ''' <param name="childPanel">Панель, которую надо убрать</param>
    Public Sub HidePanel(ByRef childPanel As clsChildPanel)
        'childPanel.parent.Panel1.Controls.Remove(childPanel.mainPanel)
        AAA()
        If IsNothing(childPanel) Then Return
        If childPanel.classId = -2 Then
            'variables
            frmMainEditor.pnlVariables.Hide()
            frmMainEditor.WBhelp.Hide()
        ElseIf childPanel.classId = -3 Then
            'functions
            With frmMainEditor.codeBox
                .Tag = Nothing
                .Text = ""
            End With
            frmMainEditor.codeBoxPanel.Hide()
            frmMainEditor.WBhelp.Hide()
        Else
            Try
                dictDefContainers(childPanel.classId).Visible = False
                Dim prevWBvisible As Boolean = childPanel.IsWbVisible
                Dim prevCbVisible As Boolean = childPanel.IsCodeBoxVisible
                If frmMainEditor.codeBoxPanel.Visible Then frmMainEditor.codeBoxPanel.Hide()
                If frmMainEditor.WBhelp.Visible Then frmMainEditor.WBhelp.Hide()
                childPanel.IsCodeBoxVisible = prevCbVisible
                childPanel.IsWbVisible = prevWBvisible
            Catch ex As Exception
                AAA()
                Return
            End Try
        End If
        AAA()
        childPanel.toolButton.Checked = False
    End Sub

    ''' <summary>Удаляет панель-окно</summary>
    ''' <param name="childPanel">Панель, которую надо удалить</param>
    ''' <param name="tryFindPrevPanel">Попытаться ли найти и открыть предыдущую панель этого же класса или не открывать ничего</param>
    Public Overloads Sub RemovePanel(ByRef childPanel As clsChildPanel, Optional ByVal tryFindPrevPanel As Boolean = True)
        AAA()
        If IsNothing(childPanel) Then Return
        Log.PrintToLog("RemovePanel: " + Log.GetChildInfo(childPanel))
        'Получаем панель, которую следует открыть вместо текущей (закрываемой) того же класса и того же родителя (если имеется), как можно ближе к текущей
        Dim currentPanelID As Integer = lstPanels.IndexOf(childPanel)
        Dim panelToOpen As clsChildPanel = Nothing
        If tryFindPrevPanel Then
            Dim offset As Integer = 1
            Do
                Dim pnl As clsChildPanel = Nothing
                For i As Integer = -1 To 1 Step 2
                    If i < 0 Then
                        If currentPanelID - offset >= 0 Then
                            pnl = lstPanels.ElementAt(currentPanelID - offset)
                        End If
                    Else
                        If currentPanelID + offset < lstPanels.Count Then
                            pnl = lstPanels.ElementAt(currentPanelID + offset)
                        End If
                    End If

                    If IsNothing(pnl) = False AndAlso pnl.classId = childPanel.classId Then
                        If pnl.classId < 0 Then
                            panelToOpen = pnl
                            Exit Do
                        ElseIf mScript.mainClass(pnl.classId).LevelsCount = 2 Then
                            If pnl.child2Name = childPanel.child2Name Then
                                panelToOpen = pnl
                                Exit Do
                            End If
                        Else
                            If pnl.supraElementName = childPanel.supraElementName Then
                                panelToOpen = pnl
                                Exit Do
                            End If
                        End If
                    End If
                Next i
                If IsNothing(pnl) Then Exit Do

                offset = offset + 1
            Loop
        End If
        childPanel.toolButton.Dispose()
        lstPanels.Remove(childPanel)

        If Object.Equals(ActivePanel, childPanel) Then
            HidePanel(childPanel)
            'Закрываемая вкладка - текущая
            'Убираем выделение с узла TreeView / кннопки отображения настроек по умолчанию
            If IsNothing(childPanel.treeNode) Then
                If childPanel.child2Name.Length = 0 Then
                    btnDefSettings.BackColor = System.Drawing.Color.FromName([Enum].GetName(GetType(KnownColor), KnownColor.Control))
                End If
            Else
                Dim tree As TreeView = childPanel.treeNode.TreeView
                If IsNothing(tree) = False AndAlso IsNothing(tree.SelectedNode) = False Then
                    If tree.SelectedNode.BackColor <> DEFAULT_COLORS.ControlTransparentBackground Then tree.SelectedNode.BackColor = DEFAULT_COLORS.ControlTransparentBackground
                    If tree.SelectedNode.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then tree.SelectedNode.ForeColor = DEFAULT_COLORS.NodeForeColor
                    tree.SelectedNode = Nothing
                End If
            End If
            'Сохраняем/восстанавливаем действия
            If currentClassName = "L" OrElse currentClassName = "A" Then actionsRouter.SaveActions()

            If IsNothing(panelToOpen) = False Then
                If panelToOpen.classId = mScript.mainClassHash("L") AndAlso panelToOpen.child2Name.Length > 0 Then
                    actionsRouter.RetreiveActions(panelToOpen.child2Name)
                ElseIf panelToOpen.classId = mScript.mainClassHash("A") AndAlso panelToOpen.supraElementName.Length > 0 Then
                    actionsRouter.RetreiveActions(panelToOpen.supraElementName)
                End If
            End If

            ActivePanel = Nothing
            'If IsNothing(childPanel.treeNode) = False AndAlso childPanel.classId >= 0 AndAlso rearrangePanelsIfNeed Then RearrangePanels(childPanel.child2Id, childPanel.treeNode.TreeView)
            If childPanel.classId = -2 Then
                ShowPanelVariables(panelToOpen)
            ElseIf childPanel.classId = -3 Then
                ShowPanelFunctions(panelToOpen)
            Else
                ShowPanel(panelToOpen)
            End If
        End If
        childPanel = Nothing
        AAA()
    End Sub

    ''' <summary>Удаляет панель-окно</summary>
    ''' <param name="childTreeNode">Узел дерева, с которым ассоциировано, которое надо удалить</param>
    ''' <remarks></remarks>
    Public Overloads Sub RemovePanel(ByRef childTreeNode As TreeNode)
        If IsNothing(childTreeNode) OrElse IsNothing(lstPanels) Then Return
        For i As Integer = 0 To lstPanels.Count - 1
            If Object.Equals(lstPanels(i).treeNode, childTreeNode) Then
                Dim pnl As clsChildPanel = lstPanels(i)
                RemovePanel(pnl)
                Return
            End If
        Next
    End Sub

    ''' <summary>
    ''' Возвращает Id панели в списке lstChildPanels или -1 если она не создана
    ''' </summary>
    ''' <param name="classId">Класс панели</param>
    ''' <param name="child2Name">Имя элемента второго порядка (или -1 если панель для свойств по умолчанию)</param>
    ''' <param name="child3Name">Имя элемента третьего порядка (или -1 если панель для свойств элемента 2-го порядка или свойств по умолчанию)</param>
    ''' <param name="supraName">Имя элемента, к которому относится данный элемент (например, Id локации, если в child2Id - Id его действия. При этом child3Id должен быть равен -1)</param>
    Private Function GetPanelId(ByVal classId As Integer, ByVal child2Name As String, ByVal child3Name As String, ByVal supraName As String) As Integer
        If IsNothing(lstPanels) OrElse lstPanels.Count = 0 Then Return -1
        For i As Integer = 0 To lstPanels.Count - 1
            Dim ch As clsChildPanel = lstPanels(i)
            If IsNothing(ch) Then Continue For
            If ch.classId = classId AndAlso ch.child2Name = child2Name AndAlso ch.child3Name = child3Name Then
                If ch.child2Name.Length = 0 Then Return i
                If ch.child3Name.Length = 0 AndAlso ch.supraElementName <> supraName Then Continue For
                Return i
            End If
        Next
        Return -1
    End Function

    Public Class clsChildPanel
        ''' <summary>Id класса, которому принадлежит панель</summary>
        Public Property classId As Integer = -1
        ''' <summary>Имя элемента второго порядка, которому принадлежит панель</summary>
        Public Property child2Name As String = ""
        ''' <summary>Имя элемента третьего порядка, которому принадлежит панель</summary>
        Public Property child3Name As String = ""
        ''' <summary>Имя элемента, к которому относится данный элемент (например, имя локации, если в child2Id - Id его действия. При этом child3Id должен быть равен -1)</summary>
        Public Property supraElementName As String = ""
        ''' <summary>Надпись для отображения на панели навигации</summary>
        Public Property Text As String = ""
        ''' <summary>Ссылка на вкладку для перехода/закрытия панели свойств</summary>
        Public Property toolButton As ToolStripButton
        ''' <summary>Ссылка на узел дерева, ассоциированный с данной панелью</summary>
        Public Property treeNode As TreeNode
        ''' <summary>Открыто ли окно браузера</summary>
        Public Property IsWbVisible As Boolean = False
        ''' <summary>Открыт ли кодбокс</summary>
        Public Property IsCodeBoxVisible As Boolean = False
        ''' <summary>Ссылка на текстбокс/комбобокс/кнопку последнего выбранного свойства</summary>
        Public lcActiveControl As Control

        Public Function GetParentId() As Integer
            If classId < 0 OrElse String.IsNullOrEmpty(supraElementName) Then Return -1
            Dim cId As Integer = classId
            If classId = mScript.mainClassHash("A") Then cId = mScript.mainClassHash("L")
            Dim res As Integer = GetSecondChildIdByName(supraElementName, mScript.mainClass(cId).ChildProperties)
            If res < 0 Then
                Return -1
            Else
                Return res
            End If
        End Function

        Public Function GetChild3Id(Optional ByVal child2Id As Integer = -1) As Integer
            If classId < 0 OrElse String.IsNullOrEmpty(child3Name) Then Return -1
            If child2Id < 0 Then child2Id = GetChild2Id()
            If child2Id < 0 Then Return -1
            Dim res As Integer = GetThirdChildIdByName(child3Name, child2Id, mScript.mainClass(classId).ChildProperties)
            If res < 0 Then
                Return -1
            Else
                Return res
            End If
        End Function

        Public Function GetChild2Id() As Integer
            If String.IsNullOrEmpty(child2Name) Then Return -1
            If classId = -2 Then
                Try
                    Return mScript.csPublicVariables.lstVariables.IndexOfKey(child2Name)
                Catch ex As Exception
                    Return -1
                End Try
            ElseIf classId = -3 Then
                Try
                    Return mScript.functionsHash.IndexOfKey(child2Name)
                Catch ex As Exception
                    Return -1
                End Try
            Else
                Dim res As Integer = GetSecondChildIdByName(child2Name, mScript.mainClass(classId).ChildProperties)
                If res < 0 Then
                    Return -1
                Else
                    Return res
                End If
            End If
        End Function
        ''' <summary>Ссылка на текстбокс/комбобокс/кнопку последнего выбранного свойства. Также устанавливает выделение нового контрола и снимает выделение со старого</summary>
        Public Property ActiveControl() As Control
            Get
                Return lcActiveControl
            End Get
            Set(value As Control)
                If Object.Equals(lcActiveControl, value) Then Return
                DeselectControl(lcActiveControl) 'Снимаем выделение со старого контрола
                lcActiveControl = value
                If IsNothing(lcActiveControl) Then Return
                Dim objActiveControl As Object = lcActiveControl
                With objActiveControl
                    .BackColor = DEFAULT_COLORS.TextBoxHighlighted 'Выделяем новый
                    If .Name.Length > 0 AndAlso IsNothing(.Parent) = False Then
                        If IsNothing(.Label) = False Then .Label.ForeColor = DEFAULT_COLORS.LabelHighLighted
                    End If
                End With
            End Set
        End Property

        ''' <summary>Снимает выделение с указанного контрола</summary>
        ''' <param name="c">ссылка на контрол, с которого надо снять выделение</param>
        Public Shared Sub DeselectControl(ByRef c As Object)
            If IsNothing(c) Then Return
            If c.GetType.Name = "ButtonEx" Then
                c.BackColor = DEFAULT_COLORS.ControlTransparentBackground
            Else
                c.BackColor = DEFAULT_COLORS.TextBoxBackColor
            End If
            If IsNothing(c.Label) = False Then c.Label.ForeColor = DEFAULT_COLORS.LabelForeColor
        End Sub

        Public Sub New(ByVal classId As Integer, ByVal child2Name As String, ByVal child3Name As String, ByRef treeNode As TreeNode, ByVal supraName As String)
            Me.classId = classId
            Me.treeNode = treeNode
            Me.child2Name = child2Name
            Me.child3Name = child3Name
            If child2Name.Length > 0 Then Me.supraElementName = supraName
        End Sub

    End Class
End Class