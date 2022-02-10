''' <summary>
''' Класс управления действиями вне текущей локации
''' </summary>
''' <remarks></remarks>
Public Class cActionsRouter
    ''' <summary>Список действий. Ключ - имя локации в кавычках, значение - список действий (Id действия)(имя свойства)</summary>
    Public lstActions As New SortedList(Of String, SortedList(Of String, MatewScript.ChildPropertiesInfoType)())
    ''' <summary>Имя локации последних восстановленных действий</summary>
    Public locationOfCurrentActions As String = ""

    ''' <summary>Список действий, сохраненный с помощью функции А.Save. Ключ - имя сохранения в кавычках, значение - список действий (Id действия)(имя свойства)</summary>
    Public lstSavedActions As New SortedList(Of String, SortedList(Of String, MatewScript.ChildPropertiesInfoType)())


    ''' <summary>Возвращает имеются ли в структуре класса сохраненные действия</summary>
    Public Function hasSavedActions() As Boolean
        If IsNothing(lstActions) OrElse lstActions.Count = 0 Then Return False
        Return True
    End Function

    ''' <summary>
    ''' Сохраняет действия локации (функция L.SaveActionsState)
    ''' </summary>
    Public Sub GameSaveActionsState()
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing
        Dim locId As Integer = GetSecondChildIdByName(locationOfCurrentActions, mScript.mainClass(mScript.mainClassHash("L")).ChildProperties)
        Dim locName As String = mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(locId)("Name").Value
        ClearOldActionsFromLstActions(locName)

        If IsNothing(mScript.mainClass(classA).ChildProperties) = False AndAlso mScript.mainClass(classA).ChildProperties.Count > 0 Then
            ReDim arrProp(mScript.mainClass(classA).ChildProperties.Count - 1)
            For actId As Integer = 0 To mScript.mainClass(classA).ChildProperties.Count - 1
                arrProp(actId) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
                For pId As Integer = 0 To mScript.mainClass(classA).ChildProperties(actId).Count - 1
                    Dim pName As String = mScript.mainClass(classA).ChildProperties(actId).ElementAt(pId).Key
                    arrProp(actId).Add(pName, mScript.mainClass(classA).ChildProperties(actId).ElementAt(pId).Value.Clone)
                Next pId
            Next actId
        End If

        If IsNothing(lstActions(locName)) OrElse lstActions.ContainsKey(locName) Then
            lstActions(locName) = arrProp
        Else
            lstActions.Add(locName, arrProp)
        End If
    End Sub

    ''' <summary>
    ''' Сохраняет действия текущей локации в структуре класса и удаляет их из mainClass
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SaveActions()
        If String.IsNullOrEmpty(locationOfCurrentActions) Then Return 'нельзя сохранить действия - действия с локации не были загружены ранее
        Dim locId As Integer = GetSecondChildIdByName(locationOfCurrentActions, mScript.mainClass(mScript.mainClassHash("L")).ChildProperties)
        If locId < 0 Then
            MsgBox("Ошибка при сохранении действий. Локации " + locationOfCurrentActions + " не найдено.", MsgBoxStyle.Exclamation)
            Return
        End If
        Dim locName As String = mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(locId)("Name").Value
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing
        arrProp = mScript.mainClass(classA).ChildProperties

        If lstActions.ContainsKey(locName) Then
            If questEnvironment.EDIT_MODE = False Then ClearOldActionsFromLstActions(locName)
            lstActions(locName) = arrProp
        Else
            lstActions.Add(locName, arrProp)
        End If
        mScript.mainClass(classA).ChildProperties = Nothing
        locationOfCurrentActions = ""
    End Sub

    ''' <summary>
    ''' Восстанавливает действия указанной локации в mainClass; в EDIT_MODE также удаляет их из структуры класса lstActions 
    ''' </summary>
    ''' <param name="locName">Имя локации</param>
    Public Sub RetreiveActions(ByVal locName As String)
        If String.IsNullOrEmpty(locName) Then Return
        If locName = locationOfCurrentActions Then Return 'нельзя дважды восстановить действия от одной локации - в структуре класса уже пусто
        Dim classA As Integer = mScript.mainClassHash("A")
        If lstActions.ContainsKey(locName) Then
            If questEnvironment.EDIT_MODE Then
                mScript.mainClass(classA).ChildProperties = lstActions(locName)
                lstActions(locName) = Nothing
            Else
                'playing quest - make full copy of actions
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing

                'make full copy of actions into arrProp array for ChildProperties
                If IsNothing(lstActions(locName)) = False Then
                    ReDim arrProp(lstActions(locName).Count - 1)
                    For actId As Integer = 0 To lstActions(locName).Count - 1
                        arrProp(actId) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
                        For pId As Integer = 0 To lstActions(locName)(actId).Count - 1
                            Dim pName As String = lstActions(locName)(actId).ElementAt(pId).Key
                            arrProp(actId).Add(pName, lstActions(locName)(actId).ElementAt(pId).Value.Clone)
                        Next pId
                    Next actId
                End If

                'remove all events from previous data in mScript.mainClass(classA).ChildProperties
                If IsNothing(mScript.mainClass(classA).ChildProperties) = False Then
                    For aId As Integer = mScript.mainClass(classA).ChildProperties.Count - 1 To 0 Step -1
                        For pId As Integer = mScript.mainClass(classA).ChildProperties(aId).Count - 1 To 0 Step -1
                            Dim eventId As Integer = mScript.mainClass(classA).ChildProperties(aId).ElementAt(pId).Value.eventId
                            If eventId > 0 Then mScript.eventRouter.RemoveEvent(eventId)
                        Next pId
                    Next aId
                    Erase mScript.mainClass(classA).ChildProperties 'clear childproperties before insertion new data
                End If

                mScript.mainClass(classA).ChildProperties = arrProp
            End If
        Else
            mScript.mainClass(classA).ChildProperties = Nothing
        End If

        locationOfCurrentActions = locName
        cListManager.UpdateList("A")
    End Sub

    ''' <summary>
    ''' Дублирует действия из текущей локации в другую
    ''' </summary>
    ''' <param name="newLocationName">Имя локации, в которую копируются действия</param>
    Public Sub DuplicateActions(ByVal newLocationName As String, Optional blnReplaceExisted As Boolean = True)
        If String.IsNullOrEmpty(locationOfCurrentActions) Then Return

        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing
        Dim srcProps() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classAct).ChildProperties
        If IsNothing(srcProps) OrElse srcProps.Count = 0 Then Return

        'Клонируем каждое свойство каждого действия локации
        ReDim arrProp(srcProps.Count - 1)
        For actId As Integer = 0 To srcProps.Count - 1
            arrProp(actId) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
            For pId As Integer = 0 To srcProps(actId).Count - 1
                Dim propName As String = srcProps(actId).ElementAt(pId).Key
                arrProp(actId).Add(propName, srcProps(actId).ElementAt(pId).Value.Clone)
                ''Дублируем событие
                'Dim eventId As Integer = arrProp(actId)(propName).eventId
                'If eventId > 0 Then arrProp(actId)(propName).eventId = mScript.eventRouter.DuplicateEvent(eventId)
            Next pId
        Next actId
        'Сохраняем в структуре класса
        If lstActions.ContainsKey(newLocationName) = False Then
            lstActions.Add(newLocationName, arrProp)
        Else
            If blnReplaceExisted Then
                'Копирование с заменой
                If IsNothing(lstActions(newLocationName)) = False AndAlso lstActions(newLocationName).Count > 0 Then
                    'в локации назначения уже существуют действия. Надо их убрать
                    For actId As Integer = lstActions(newLocationName).Count - 1 To 0 Step -1
                        'очищаем события
                        Dim arrCh As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = lstActions(newLocationName)(actId)
                        For propId = 0 To arrCh.Count - 1
                            Dim ch As MatewScript.ChildPropertiesInfoType = arrCh.ElementAt(propId).Value
                            If ch.eventId > 0 Then mScript.eventRouter.RemoveEvent(ch.eventId)
                        Next propId
                        arrCh.Clear()
                    Next actId
                End If
                lstActions(newLocationName) = arrProp
            Else
                'Копирование с добавлением к новым
                If IsNothing(lstActions(newLocationName)) OrElse lstActions(newLocationName).Count = 0 Then
                    lstActions(newLocationName) = arrProp
                    Return
                End If

                Dim cnt As Integer = arrProp.Count
                Dim initCnt As Integer = 0
                initCnt = lstActions(newLocationName).Count
                ReDim Preserve lstActions(newLocationName)(initCnt - 1 + cnt)
                Array.Copy(arrProp, 0, lstActions(newLocationName), initCnt, cnt)
                'проверяем имена
                For actId As Integer = initCnt To initCnt + cnt - 1
                    Dim aName As String = lstActions(newLocationName)(actId)("Name").Value
                    If actId > 0 Then
                        For i As Integer = actId - 1 To 0 Step -1
                            Dim curName As String = lstActions(newLocationName)(i)("Name").Value
                            If String.Compare(aName, curName, True) = 0 Then
                                'такое имя уже существует. Переименовываем
                                lstActions(newLocationName)(actId)("Name").Value = GetNewActionName(lstActions(newLocationName))
                                Exit For
                            End If
                        Next i
                    End If
                Next actId
            End If
        End If

    End Sub

    ''' <summary>
    ''' Возвращает новое имя для произвольной сохраненной локации
    ''' </summary>
    ''' <param name="lstActions">действия произвольной сохраненной локации</param>
    Private Function GetNewActionName(ByRef lstActions() As SortedList(Of String, MatewScript.ChildPropertiesInfoType)) As String
        Dim defElementName = "'" + "Действие "
        Dim i As Integer = 1
        Do
            Dim newName As String = defElementName + i.ToString + "'"
            Dim id As Integer = GetSecondChildIdByName(newName, lstActions)
            If id < 0 Then Return newName
            i += 1
        Loop
    End Function

    ''' <summary>
    ''' Обновляет действия: сохраняет действия текущей локации (если надо) и восстанавливает новой, открываемой (если надо)
    ''' </summary>
    ''' <param name="newClassId">Класс открываемого элемента</param>
    ''' <param name="locationName">Имя элемента, который предположительно является локацией</param>
    Public Sub UpdateActions(ByVal newClassId As Integer, ByVal locationName As String)
        If currentClassName = "A" OrElse currentClassName = "L" Then
            SaveActions()
        End If

        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If newClassId = classAct OrElse newClassId = classLoc AndAlso locationName.Length > 0 Then
            RetreiveActions(locationName)
        End If
    End Sub

    ''' <summary>
    ''' Возвращает Id открытой локации или -1 если сейчас открыты не действия и не локация
    ''' </summary>
    Public Function GetActiveLocationId() As Integer
        If String.IsNullOrEmpty(locationOfCurrentActions) Then Return -1
        Dim loc As Integer = GetSecondChildIdByName(locationOfCurrentActions, mScript.mainClass(mScript.mainClassHash("L")).ChildProperties)
        If loc < -1 Then loc = -1
        Return loc
    End Function

    ''' <summary>
    ''' Переименовывает локацию в структуре класса
    ''' </summary>
    ''' <param name="oldName">Старое имя локации</param>
    ''' <param name="newName">Новое имя локации</param>
    ''' <remarks></remarks>
    Public Sub RenameLocation(ByVal oldName As String, ByVal newName As String)
        If locationOfCurrentActions = oldName Then locationOfCurrentActions = newName
        Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing
        If lstActions.TryGetValue(oldName, arrProp) = False Then Return
        lstActions.Remove(oldName)
        lstActions.Add(newName, arrProp)
    End Sub

    ''' <summary>
    ''' Удаляет локацию из структуры класса, очищает действия в mainClass
    ''' </summary>
    ''' <param name="locName"></param>
    ''' <remarks></remarks>
    Public Sub RemoveLocation(ByVal locName As String)
        If locName = locationOfCurrentActions Then
            Dim classA As Integer = mScript.mainClassHash("A")
            mScript.mainClass(classA).ChildProperties = Nothing
        End If
        locationOfCurrentActions = ""
        If lstActions.ContainsKey(locName) Then lstActions.Remove(locName)
    End Sub

    ''' <summary>
    ''' Переименовывает свойство действий во всех сохраненных действиях
    ''' </summary>
    ''' <param name="oldName">Старое имя свойства</param>
    ''' <param name="newName">Новое имя свойства</param>
    ''' <remarks></remarks>
    Public Sub RenameProperty(ByVal oldName As String, ByVal newName As String)
        For i As Integer = 0 To lstActions.Count - 1
            Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = lstActions.ElementAt(i).Value
            Dim locName As String = lstActions.ElementAt(i).Key
            If IsNothing(arrProp) Then Continue For
            For aId As Integer = 0 To arrProp.Count - 1
                Dim p As MatewScript.ChildPropertiesInfoType = Nothing
                If arrProp(aId).TryGetValue(oldName, p) Then
                    arrProp(aId).Remove(oldName)
                    arrProp(aId).Add(newName, p)
                End If
            Next aId
        Next i
    End Sub

    ''' <summary>
    ''' Добавляет сохраненным действиям новое свойство
    ''' </summary>
    ''' <param name="propName">Имя свойства</param>
    ''' <remarks></remarks>
    Public Sub AddProperty(ByVal propName As String)
        For i As Integer = 0 To lstActions.Count - 1
            Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = lstActions.ElementAt(i).Value
            Dim locName As String = lstActions.ElementAt(i).Key
            If IsNothing(arrProp) Then Continue For
            For aId As Integer = 0 To arrProp.Count - 1
                If arrProp(aId).ContainsKey(propName) = False Then
                    Dim p As New MatewScript.ChildPropertiesInfoType
                    arrProp(aId).Add(propName, p)
                End If
            Next aId
        Next i
    End Sub

    ''' <summary>
    ''' Удаляет у сохраненных действий свойство
    ''' </summary>
    ''' <param name="remName">Имя свойства</param>
    ''' <remarks></remarks>
    Public Sub RemoveProperty(ByVal remName As String)
        For i As Integer = 0 To lstActions.Count - 1
            Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = lstActions.ElementAt(i).Value
            Dim locName As String = lstActions.ElementAt(i).Key
            If IsNothing(arrProp) Then Continue For
            For aId As Integer = 0 To arrProp.Count - 1
                If arrProp(aId).ContainsKey(remName) Then
                    arrProp(aId).Remove(remName)
                End If
            Next aId
        Next i
    End Sub

    ''' <summary>
    ''' Устанавливает значение указанного свойства всем сохраненным действиям. Если свойтсва нет - создает его.
    ''' </summary>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="newValue">Устанавливаемое значение</param>
    Public Sub SetPropertyValue(ByVal propName As String, ByVal newValue As String)
        If IsNothing(lstActions) OrElse lstActions.Count = 0 Then Return

        Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(newValue)
        For i As Integer = 0 To lstActions.Count - 1
            Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = lstActions.ElementAt(i).Value
            If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
            Dim locName As String = lstActions.ElementAt(i).Key
            For aId As Integer = 0 To arrProp.Count - 1
                Dim p As MatewScript.ChildPropertiesInfoType = Nothing
                If arrProp(aId).TryGetValue(propName, p) = False Then
                    p = New MatewScript.ChildPropertiesInfoType
                    arrProp(aId).Add(propName, p)
                End If

                If p.eventId > 0 Then mScript.eventRouter.RemoveEvent(p.eventId)

                If ret = MatewScript.ContainsCodeEnum.CODE OrElse ret = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                    Dim cd() As CodeTextBox.CodeDataType
                    With questEnvironment.codeBoxShadowed.codeBox
                        .Text = ""
                        .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                        .LoadCodeFromProperty(newValue)
                        cd = CopyCodeDataArray(.CodeData)
                    End With
                    p.eventId = mScript.eventRouter.SetEventId(p.eventId, cd)
                Else
                    p.eventId = -1
                End If
                p.Value = newValue
                arrProp(aId)(propName) = p
            Next aId
        Next i
    End Sub

    ''' <summary>
    ''' Создает список действий указанной локации. Если действие из текущей локации - список создается из mainClass
    ''' </summary>
    ''' <param name="locId">Id локации</param>
    ''' <returns></returns>
    Public Function GetActionsNames(ByVal locId As Integer) As List(Of String)
        Dim arrNames As New List(Of String)
        Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing
        Dim curLocId As Integer = -1
        If String.IsNullOrEmpty(locationOfCurrentActions) = False Then curLocId = GetSecondChildIdByName(locationOfCurrentActions, mScript.mainClass(mScript.mainClassHash("L")).ChildProperties)

        If curLocId = locId Then
            'Получаем список из текущих действий из mainClass
            arrProp = mScript.mainClass(mScript.mainClassHash("A")).ChildProperties
            If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Return arrNames
        Else
            'Получаем список из сохраненных в этом классе действий
            Dim locName As String = mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(locId)("Name").Value
            If lstActions.TryGetValue(locName, arrProp) = False Then Return arrNames
            If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Return arrNames
        End If

        For i As Integer = 0 To arrProp.Count - 1
            arrNames.Add(arrProp(i)("Name").Value)
        Next
        Return arrNames
    End Function

#Region "Game Functions"
    ''' <summary>
    ''' Очищает старые действия данной локации перед вставкой новых
    ''' </summary>
    ''' <param name="locName">Имя локации</param>
    Public Sub ClearOldActionsFromLstActions(ByVal locName As String)
        If lstActions.ContainsKey(locName) = False OrElse IsNothing(lstActions(locName)) OrElse lstActions(locName).Count = 0 Then
            'действий на локации нет
            Return
        End If

        Dim classA As Integer = mScript.mainClassHash("A")
        'удаляем события
        Dim curActions As SortedList(Of String, MatewScript.ChildPropertiesInfoType)() = lstActions(locName)
        For actId As Integer = 0 To curActions.Count - 1
            'перебираем все действия локации
            For pId As Integer = 0 To curActions(actId).Count - 1
                Dim eventId As Integer = curActions(actId).ElementAt(pId).Value.eventId
                If eventId > 0 Then mScript.eventRouter.RemoveEvent(eventId) 'удаляем события
            Next pId
        Next actId
        'удаляем действия
        Erase curActions
    End Sub

    ''' <summary>
    ''' Загружает копию действий указанной локации в игру
    ''' </summary>
    ''' <param name="locName">Имя локации</param>
    Public Sub GameLoadActions(ByVal locName As String)
        Dim classA As Integer = mScript.mainClassHash("A")

        If lstActions.ContainsKey(locName) = False OrElse IsNothing(lstActions(locName)) OrElse lstActions(locName).Count = 0 Then
            'действий на локации нет
            mScript.mainClass(classA).ChildProperties = Nothing
            Return
        End If

        'действия на локации есть
        Dim curActions As SortedList(Of String, MatewScript.ChildPropertiesInfoType)() = lstActions(locName)
        'клонируем действия (вместе с событиями)
        ReDim mScript.mainClass(classA).ChildProperties(curActions.Count - 1)
        For actId As Integer = 0 To curActions.Count - 1
            'перебираем все действия локации
            mScript.mainClass(classA).ChildProperties(actId) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
            Dim destProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classA).ChildProperties(actId)
            For pId As Integer = 0 To curActions(actId).Count - 1
                'клонируем все свойства действия
                destProps.Add(curActions(actId).ElementAt(pId).Key, curActions(actId).ElementAt(pId).Value.Clone)
            Next pId
        Next actId
    End Sub

    ''' <summary>
    ''' Сохраняет текущие действия для дальнейшего восстановления с помощью функции A.Load
    ''' </summary>
    ''' <param name="saveName">Имя сохранения</param>
    Public Sub GameSaveActions(Optional ByVal saveName As String = "temp")
        Dim classA As Integer = mScript.mainClassHash("A")
        Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing

        If IsNothing(mScript.mainClass(classA).ChildProperties) = False AndAlso mScript.mainClass(classA).ChildProperties.Count > 0 Then
            ReDim arrProp(mScript.mainClass(classA).ChildProperties.Count - 1)
            For actId As Integer = 0 To mScript.mainClass(classA).ChildProperties.Count - 1
                arrProp(actId) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
                For pId As Integer = 0 To mScript.mainClass(classA).ChildProperties(actId).Count - 1
                    Dim pName As String = mScript.mainClass(classA).ChildProperties(actId).ElementAt(pId).Key
                    arrProp(actId).Add(pName, mScript.mainClass(classA).ChildProperties(actId).ElementAt(pId).Value.Clone)
                Next pId
            Next actId
        End If

        If IsNothing(lstSavedActions) = False AndAlso lstSavedActions.ContainsKey(saveName) Then
            lstSavedActions(saveName) = arrProp
        Else
            lstSavedActions.Add(saveName, arrProp)
        End If
    End Sub

    ''' <summary>
    ''' Загружает копию действий, сохраненную с помощью функции A.Save
    ''' </summary>
    ''' <param name="saveName">Имя сохранения</param>
    Public Function GameLoadSavedActions(Optional ByVal saveName As String = "temp") As String
        Dim classA As Integer = mScript.mainClassHash("A")

        Dim curActions As SortedList(Of String, MatewScript.ChildPropertiesInfoType)() = mScript.mainClass(classA).ChildProperties
        'удаляем текущие действия
        If IsNothing(curActions) = False AndAlso curActions.Count > 0 Then
            'удаляем события
            For actId As Integer = 0 To curActions.Count - 1
                Dim actContainer As String = ""
                If ReadProperty(classA, "HTMLContainerId", actId, -1, actContainer, {actId.ToString}) Then
                    actContainer = UnWrapString(actContainer)
                    If String.IsNullOrEmpty(actContainer) = False Then
                        'указан контейнер. Удаляем действие из html
                        Dim hAct As mshtml.IHTMLDOMNode = FindActionHTMLelementDOM(classA, actId, {actId.ToString})
                        If IsNothing(hAct) = False Then hAct.removeNode(True)
                    End If
                End If
                'перебираем все действия локации
                For pId As Integer = 0 To curActions(actId).Count - 1
                    Dim eventId As Integer = curActions(actId).ElementAt(pId).Value.eventId
                    If eventId > 0 Then mScript.eventRouter.RemoveEvent(eventId) 'удаляем события
                Next pId
            Next actId
            'удаляем действия
            Erase curActions
            'очистка окна действий
            Dim actConvas As HtmlElement = frmPlayer.wbActions.Document.GetElementById("ActionsConvas")
            If IsNothing(actConvas) = False Then actConvas.InnerHtml = ""
        End If


        If lstSavedActions.ContainsKey(saveName) = False Then ' OrElse IsNothing(lstSavedActions(locName)) OrElse lstSavedActions(locName).Count = 0 Then
            'сохраненных действий с указанным именем нет
            Return _ERROR(String.Format("Набора сохраненных действий {0}не найдено.", IIf(saveName = "temp", "", "с именем " & saveName & " ")), "Action.Load")
        End If

        'сохраненные действия есть
        curActions = lstSavedActions(saveName)
        'клонируем действия (вместе с событиями)
        ReDim mScript.mainClass(classA).ChildProperties(curActions.Count - 1)
        For actId As Integer = 0 To curActions.Count - 1
            'перебираем все действия сохранения
            mScript.mainClass(classA).ChildProperties(actId) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
            Dim destProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classA).ChildProperties(actId)
            For pId As Integer = 0 To curActions(actId).Count - 1
                'клонируем все свойства действия
                destProps.Add(curActions(actId).ElementAt(pId).Key, curActions(actId).ElementAt(pId).Value.Clone)
            Next pId
        Next actId
        Return ""
    End Function

    ''' <summary>
    ''' Удаляет сохраненные действия под указанным именем или (если имя не указано) удаляет все сохраненные действия
    ''' </summary>
    ''' <param name="delName">имя удаляемой группы сохраненных действий</param>
    Public Sub GameRemoveSavedActions(ByVal delName As String)
        If IsNothing(lstSavedActions) OrElse lstSavedActions.Count = 0 Then Return
        Dim curActions As SortedList(Of String, MatewScript.ChildPropertiesInfoType)() = Nothing

        If String.IsNullOrEmpty(delName) Then
            'очищаем все сохраненные действия
            'очистка событий
            For i As Integer = 0 To lstSavedActions.Count - 1
                curActions = lstSavedActions.ElementAt(i).Value
                If IsNothing(curActions) OrElse curActions.Count = 0 Then Continue For
                For actId As Integer = 0 To curActions.Count - 1
                    Dim props As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = curActions(actId)
                    For pId As Integer = 0 To props.Count - 1
                        Dim eventId As Integer = props.ElementAt(pId).Value.eventId
                        If eventId > 0 Then mScript.eventRouter.RemoveEvent(eventId)
                    Next pId
                Next actId
            Next i
            lstSavedActions.Clear()
            Return
        End If

        'очистка действий под указанным именем
        If lstSavedActions.TryGetValue(delName, curActions) = True Then Return

        If IsNothing(curActions) = False AndAlso curActions.Count > 0 Then
            'очистка событий
            For actId As Integer = 0 To curActions.Count - 1
                Dim props As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = curActions(actId)
                For pId As Integer = 0 To props.Count - 1
                    Dim eventId As Integer = props.ElementAt(pId).Value.eventId
                    If eventId > 0 Then mScript.eventRouter.RemoveEvent(eventId)
                Next pId
            Next actId
        End If
        lstSavedActions.Remove(delName)
    End Sub

#End Region
End Class
