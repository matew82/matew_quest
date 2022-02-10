Public Class dlgEntrancies
    ''' <summary>
    ''' Класс для поиска вхождений чего-либо, содержащий всю инфрормацию, необходимую для перехода
    ''' </summary>
    Public Class cEntranciesClass
        Public classId As Integer = -1
        Public child2Id As Integer = -1
        Public child3Id As Integer = -1
        Public parentId As Integer = -1
        Public elementName As String = ""
        Public seekLine As Integer = -1
        Public word As String = ""
        Public tracking As frmMainEditor.trackingcodeEnum = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT
        ''' <summary>
        ''' Функция или свойство
        ''' </summary>
        Public elementType As CodeTextBox.EditWordTypeEnum
        ''' <summary>
        ''' Начало слова в относительно начала строки
        ''' </summary>
        Public wordStart As Integer = -1

        Public Sub New(ByVal classId As Integer, child2Id As Integer, ByVal child3Id As Integer, elementName As String, ByVal elementType As CodeTextBox.EditWordTypeEnum, _
                       Optional ByVal parentId As Integer = -1, Optional ByVal seekLine As Integer = -1, Optional word As String = "", Optional wordStart As Integer = 0, _
                       Optional tracking As frmMainEditor.trackingcodeEnum = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT)
            Me.classId = classId
            Me.child2Id = child2Id
            Me.child3Id = child3Id
            Me.elementName = elementName
            Me.elementType = elementType
            Me.parentId = parentId
            Me.seekLine = seekLine
            Me.word = word
            Me.wordStart = wordStart
            Me.tracking = tracking
        End Sub
    End Class

    Public Enum EntranciesStyleEnum As Byte
        Simple = 0
        Extended = 1
    End Enum

    ''' <summary>список входждений чего-либо для возможности навигации по этим вхождениям</summary>
    Private lstEntrancies As New List(Of cEntranciesClass)
    ''' <summary>Преднастройки - значения по умолчанию для новых вхождений</summary>
    Private entrancePreset As New cEntranciesClass(-1, -1, -1, "", CodeTextBox.EditWordTypeEnum.W_PROPERTY)
    ''' <summary>В случае, если выводится список при удалении действия, Id действия. Иначе -1</summary>
    Private actionRemovedId As Integer = -1
    Private currentShowStyle As EntranciesStyleEnum = EntranciesStyleEnum.Simple

    ''' <summary>
    ''' Очищает список для начала ввода
    ''' </summary>
    ''' <param name="strInfo">Строка информации</param>
    ''' <param name="actionRemovingId">В случае, если выводится список при удалении действия, Id действия. Иначе -1</param>
    Public Sub BeginNewEntrancies(ByVal strInfo As String, ByVal style As EntranciesStyleEnum, Optional actionRemovingId As Integer = -1)
        lstBoxEntrancies.Items.Clear()
        lstEntrancies.Clear()
        lblInfo.Text = strInfo
        actionRemovedId = actionRemovingId

        If style = EntranciesStyleEnum.Simple Then
            lstBoxEntranciesChecked.Hide()
            lstBoxEntrancies.Show()
            chkReplace.Visible = False
            txtReplace.Visible = False
            btnReplace.Visible = False
        Else
            lstBoxEntranciesChecked.Show()
            lstBoxEntrancies.Hide()
            chkReplace.Visible = True
            txtReplace.Visible = True
            btnReplace.Visible = True
        End If
        currentShowStyle = style
        Me.Hide()
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click, lstBoxEntrancies.DoubleClick, lstBoxEntranciesChecked.DoubleClick
        Dim id As Integer = -1
        If currentShowStyle = EntranciesStyleEnum.Simple Then
            id = lstBoxEntrancies.SelectedIndex
        Else
            id = lstBoxEntranciesChecked.SelectedIndex
        End If

        If id = -1 Then Return
        If id > lstEntrancies.Count - 1 Then Return
        Dim ent As cEntranciesClass = lstEntrancies(id)
        cPanelManager.FindAndOpen(ent.classId, ent.child2Id, ent.child3Id, ent.parentId, ent.elementType, ent.elementName, ent.seekLine, ent.word, ent.wordStart, ent.tracking)
    End Sub

    Private Sub dlgEntrancies_Load(sender As Object, e As EventArgs) Handles Me.Load
        splitMain.FixedPanel = FixedPanel.Panel2
    End Sub

    ''' <summary>
    ''' Создает объект преднастройки entrancePreset. При создании новых вхождений недостающие данные будут браться отсюда
    ''' </summary>
    ''' <param name="classId">Класс</param>
    ''' <param name="child2Id">Элемент 2 порядка</param>
    ''' <param name="child3Id">Элемент 3 порядка</param>
    ''' <param name="elementName">Имя Свойства/функции/переменной</param>
    ''' <param name="elementType">Тип элемента</param>
    ''' <param name="parentId">Id локации для действий</param>
    ''' <param name="seekLine">Номер строки, начиная от 0</param>
    ''' <remarks></remarks>
    Public Sub SetEntranceDefault(ByVal classId As Integer, child2Id As Integer, ByVal child3Id As Integer, elementName As String, ByVal elementType As CodeTextBox.EditWordTypeEnum, _
                                  Optional ByVal parentId As Integer = -2, Optional ByVal seekLine As Integer = -1, Optional ByVal word As String = "", Optional ByVal wordStart As Integer = -1, _
                                  Optional ByVal tracking As frmMainEditor.trackingcodeEnum = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT)
        With entrancePreset
            .classId = classId
            .child2Id = child2Id
            .child3Id = child3Id
            .elementName = elementName
            .elementType = elementType
            .tracking = tracking
            If parentId > -2 Then .parentId = parentId
            .seekLine = seekLine
            .word = word
            .wordStart = wordStart
        End With
    End Sub

    ' ''' <summary>
    ' ''' Устанавливает значение по умолчанию parentId для новых вхождений
    ' ''' </summary>
    ' ''' <param name="locationId"></param>
    ' ''' <remarks></remarks>
    'Public Sub SetSeekLocationId(ByVal locationId As Integer)
    '    entrancePreset.parentId = locationId
    'End Sub

    ''' <summary>
    ''' В объект преднастроек добавляет данные о линии
    ''' </summary>
    ''' <param name="seekLine">Номер линии от 0</param>
    ''' <param name="word">Найденное слово</param>
    ''' <param name="wordStart">Начало слова относительно начала строки</param>
    Public Sub SetSeekPosInfo(ByVal seekLine As Integer, Optional ByVal word As String = "", Optional ByVal wordStart As Integer = -1)
        entrancePreset.seekLine = seekLine
        If wordStart > -1 Then
            entrancePreset.word = word
            entrancePreset.wordStart = wordStart
        End If
    End Sub

    ''' <summary>
    ''' Возвращает имеется ли хоть одно вхождение
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function hasEntrancies() As Boolean
        If IsNothing(lstEntrancies) OrElse lstEntrancies.Count = 0 Then Return False
        Return True
    End Function

    ''' <summary>
    ''' Создает новое вхождение
    ''' </summary>
    ''' <param name="classId">Имя класса</param>
    ''' <param name="child2Id">Элемент 2 порядка</param>
    ''' <param name="child3Id">Элемент 3 порядка</param>
    ''' <param name="elementName">Имя Свойства/функции/переменной</param>
    ''' <param name="elementType">Тип элемента</param>
    ''' <param name="parentId">Id локации для действий</param>
    ''' <param name="seekLine">Номер строки, начиная от 0</param>
    Public Sub NewEntrance(Optional ByVal classId As Integer = -1, Optional ByVal child2Id As Integer = -2, Optional ByVal child3Id As Integer = -2, Optional ByVal elementName As String = "", _
                           Optional ByVal elementType As CodeTextBox.EditWordTypeEnum = CodeTextBox.EditWordTypeEnum.W_ERROR, Optional ByVal parentId As Integer = -2, _
                           Optional ByVal seekLine As Integer = -2, Optional ByVal word As String = "", Optional ByVal wordStart As Integer = -2, Optional ByVal tracking As frmMainEditor.trackingcodeEnum = 255)
        'Берем недостающие данные из entrancePreset
        With entrancePreset
            If classId = -1 Then classId = .classId
            If child2Id = -2 Then child2Id = .child2Id
            If child3Id = -2 Then child3Id = .child3Id
            If String.IsNullOrEmpty(elementName) Then elementName = .elementName
            If elementType = CodeTextBox.EditWordTypeEnum.W_ERROR Then elementType = .elementType
            If parentId = -2 Then parentId = .parentId
            If seekLine = -2 Then seekLine = .seekLine
            If tracking = 255 Then tracking = .tracking
            If wordStart = -2 Then
                word = .word
                wordStart = .wordStart
            End If
        End With
        If classId = mScript.mainClassHash("A") AndAlso child2Id >= 0 AndAlso parentId < 0 Then Return 'неполная инфо о действии - выход (дублируется другой строкой, с полной информацией, поэтому эта - лишняя)

        'Создаем строку
        Dim strSeek As String = ""
        If seekLine >= 0 Then
            strSeek = ", строка " + (seekLine + 1).ToString
        End If
        Dim sb As New System.Text.StringBuilder
        If elementType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
            Dim p As MatewScript.PropertiesInfoType = Nothing ' mScript.mainClass(classId).Properties(elementName)
            If mScript.mainClass(classId).Properties.TryGetValue(elementName, p) = False Then Return
            'Проверка является ли свойство скрытым
            If p.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse p.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR Then Return
            If child2Id = -1 Then
                '1 уровень
                If p.Hidden = MatewScript.PropertyHiddenEnum.LEVEL2_ONLY OrElse p.Hidden = MatewScript.PropertyHiddenEnum.LEVEL23_ONLY OrElse p.Hidden = MatewScript.PropertyHiddenEnum.LEVEL3_ONLY Then Return
            ElseIf child3Id < 0 Then
                '2 уровень
                If p.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY OrElse p.Hidden = MatewScript.PropertyHiddenEnum.LEVEL13_ONLY OrElse p.Hidden = MatewScript.PropertyHiddenEnum.LEVEL3_ONLY Then Return
            Else
                '3 уровень
                If p.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY OrElse p.Hidden = MatewScript.PropertyHiddenEnum.LEVEL12_ONLY OrElse p.Hidden = MatewScript.PropertyHiddenEnum.LEVEL2_ONLY Then Return
            End If
            sb.Append("Свойство " + elementName)
            If String.IsNullOrEmpty(p.EditorCaption) = False AndAlso p.EditorCaption <> elementName Then
                sb.Append(" (" + p.EditorCaption + ")")
            End If
            sb.Append(" класса " + mScript.mainClass(classId).Names.Last)

            If tracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE Then
                sb.Append(" (событие перед изменением свойства)")
            ElseIf tracking = frmMainEditor.trackingcodeEnum.EVENT_AFTER Then
                sb.Append(" (событие после изменения свойства)")
            ElseIf child2Id < 0 AndAlso mScript.mainClass(classId).LevelsCount > 0 Then
                sb.Append(" (значение по умолчанию)")
            End If

            If child2Id >= 0 Then
                If parentId < 0 Then
                    sb.Append(", " + mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("Name").Value, Nothing, False))
                Else
                    Dim lstA As List(Of String) = actionsRouter.GetActionsNames(parentId) ' 'GetActionNamesFromXml(mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(parentId)("Actions").Value)

                    If actionRemovedId > -1 AndAlso actionRemovedId < lstA.Count Then lstA.RemoveAt(actionRemovedId)
                    If child2Id < lstA.Count Then sb.Append(", " + lstA(child2Id))
                    sb.Append(" (локация " + mScript.PrepareStringToPrint(mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(parentId)("Name").Value, Nothing, False) + ")")
                End If
                If child3Id >= 0 Then
                    sb.Append(" -> " + mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("Name").ThirdLevelProperties(child3Id), Nothing, False))
                End If
            End If
        ElseIf elementType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
            sb.Append("Функция " + elementName + " класса " + mScript.mainClass(classId).Names.Last)
        ElseIf elementType = CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION Then
            sb.Append("Функция Писателя " + elementName)
        ElseIf elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
            If child3Id < 0 Then
                sb.Append("Глобальная переменная " + elementName)
            Else
                sb.Append("Глобальная переменная " + elementName + "[" + child3Id.ToString + "]")
            End If
        End If
        sb.Append(strSeek)
        lstEntrancies.Add(New cEntranciesClass(classId, child2Id, child3Id, elementName, elementType, parentId, seekLine, word, wordStart, tracking))
        If currentShowStyle = EntranciesStyleEnum.Simple Then
            lstBoxEntrancies.Items.Add(sb.ToString)
        Else
            lstBoxEntranciesChecked.Items.Add(sb.ToString, True)
        End If

        sb.Clear()
    End Sub

    Private Sub chkReplace_CheckedChanged(sender As Object, e As EventArgs) Handles chkReplace.CheckedChanged
        Dim chk As Boolean = sender.Checked
        If lstBoxEntranciesChecked.Items.Count = 0 Then chk = False
        txtReplace.Enabled = chk
        btnReplace.Enabled = chk
    End Sub

    Private Sub btnReplace_Click(sender As Object, e As EventArgs) Handles btnReplace.Click
        If hasEntrancies() = False OrElse lstBoxEntranciesChecked.Items.Count = 0 Then
            MessageBox.Show("Не найдено ни одного совпадения!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim strReplace As String = txtReplace.Text
        Dim cntErr As Integer = 0
        Dim cnt As Integer = 0
        For i As Integer = lstBoxEntranciesChecked.Items.Count - 1 To 0 Step -1
            If Not lstBoxEntranciesChecked.GetItemChecked(i) Then Continue For
            If GlobalSeeker.ReplaceString(lstEntrancies(i), strReplace) Then
                lstBoxEntranciesChecked.SetItemChecked(i, False)
                cnt += 1
            Else
                cntErr += 1
                MessageBox.Show("Ошибка при попытке замены. Строка была изменена ранее либо после замены в скрипте возникает синтаксическая ошибка. Пункт: " + vbNewLine + lstBoxEntranciesChecked.Items(i).ToString, _
                                "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Next i

        If cnt > 0 AndAlso cntErr > 0 Then
            MessageBox.Show("Успешно выполнено " + cnt.ToString + " замен(ы). Обнаружено " + cntErr.ToString + " Ошибка(и).", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf cnt = 0 AndAlso cntErr > 0 Then
            MessageBox.Show("Замены не произведены. Обнаружено " + cntErr.ToString + " Ошибка(и).", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf cnt > 0 AndAlso cntErr = 0 Then
            MessageBox.Show("Успешно выполнено " + cnt.ToString + " замен(ы). Ошибок не обнаружено.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Замены не произведены. Необходимо поставить галочки напротив найденных совпадений, которые надо заменить.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class