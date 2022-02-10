''' <summary>Класс для работы , создания и восстановления удаленных элементов</summary>
Public Class cRemovedObjects
    ''' <summary>Класс для хранения данных об удаленном элементе</summary>
    Public Class cRemovedItem
        Public element As Object
        Public elementName As String
        Public parentName As String
        Public classId As Integer = -1

        ''' <summary>Создает строку с информацией об элементе</summary>
        Public Function GetString() As String
            Dim aStr As String = ""
            Dim aName As String

            If classId < 0 Then
                aName = elementName
            Else
                aName = mScript.PrepareStringToPrint(elementName, Nothing, False)
            End If

            Select Case classId
                Case mScript.mainClassHash("L")
                    Dim tName As String = element.GetType.Name
                    aStr = frmMainEditor.GetTranslatedName(classId, False).ToUpper + ": " + aName
                Case Else
                    Dim tName As String = element.GetType.ToString
                    If tName.EndsWith("+ChildPropertiesInfoType]") Then
                        'Элемент 2 уровня
                        aStr = frmMainEditor.GetTranslatedName(classId, False).ToUpper + ": " + aName
                    ElseIf tName.EndsWith("+PropertiesInfoType]") Then
                        'Элемент 3 уровня
                        aStr = frmMainEditor.GetTranslatedName(classId, True).ToUpper + ": " + aName + " (" + mScript.PrepareStringToPrint(parentName, Nothing, False) + ")"
                    ElseIf tName.EndsWith("+FunctionInfoType") OrElse tName.EndsWith("+variableEditorInfoType") Then
                        'Функции / Перменные
                        aStr = frmMainEditor.GetTranslatedName(classId, False).ToUpper + ": " + aName
                    End If
            End Select

            Return aStr
        End Function
    End Class

    ''' <summary>Список удаленных элементов</summary>
    Public lstRemoved As New List(Of cRemovedItem)
    ''' <summary>Класс для хранения скриптов удаленных элементов</summary>
    Private eventContainer As New EventContainerClass

    ''' <summary>
    ''' Добавляет новый элемент для восстановления
    ''' </summary>
    ''' <param name="element">Объект, хранящий элемент - список ChildProperies, FunctionInfo, VariableInfo ...</param>
    ''' <param name="classId">Id класса</param>
    ''' <param name="elementName">Имя удаляемого элемента</param>
    ''' <param name="parentId">id его родителя или -1, если такового нет</param>
    Public Sub AddItem(ByRef element As Object, ByVal classId As Integer, ByVal elementName As String, Optional ByVal parentId As Integer = -1)
        AAA()
        Dim parentName As String = ""
        Dim classAct As Integer = mScript.mainClassHash("A")
        If parentId >= 0 Then
            If classId = classAct Then
                parentName = mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(parentId)("Name").Value
            Else
                parentName = mScript.mainClass(classId).ChildProperties(parentId)("Name").Value
            End If
        End If

        Dim itm As New cRemovedItem With {.classId = classId, .elementName = elementName, .parentName = parentName}
        Select Case classId
            Case -2, -3
                'Variable, Function
                itm.element = element
            Case mScript.mainClassHash("L")
                'Удаляется локация
                eventContainer.CutAllEvents(classId, elementName, -1, -1)

                Dim col As New List(Of Object)
                col.Add(element)
                If IsNothing(mScript.mainClass(classAct).ChildProperties) = False AndAlso mScript.mainClass(classAct).ChildProperties.Count > 0 Then
                    'Локация с действиями
                    If IsNothing(cPanelManager.ActivePanel) OrElse cPanelManager.ActivePanel.child2Name.Length = 0 OrElse cPanelManager.ActivePanel.classId <> mScript.mainClassHash("L") Then
                        MessageBox.Show("Не удалось найти текущую вкладку. Резервное сохранение удаляемой локации невозможно.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If

                    col.Add(mScript.mainClass(classAct).ChildProperties)
                    'Вырезаем все события действий удаленной локации
                    For i As Integer = 0 To mScript.mainClass(classAct).ChildProperties.Count - 1
                        eventContainer.CutAllEvents(classAct, mScript.mainClass(classAct).ChildProperties(i)("Name").Value, i, -1)
                    Next
                End If
                itm.element = col
            Case Else
                If String.IsNullOrEmpty(parentName) OrElse classId = classAct Then
                    'удаляем элемент 2 уровня
                    itm.element = element
                    eventContainer.CutAllEvents(classId, elementName, -1, -1)
                Else
                    'удаляем элемент 3 уровня
                    Dim lst As New SortedList(Of String, MatewScript.PropertiesInfoType) 'Name,Value
                    Dim child3Id As Integer = GetThirdChildIdByName(elementName, parentId, mScript.mainClass(classId).ChildProperties)
                    If child3Id < 0 Then
                        MessageBox.Show("Не удалось создать резервную копию. Элемент не найден.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                    For i As Integer = 0 To mScript.mainClass(classId).ChildProperties(parentId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(parentId).ElementAt(i).Value
                        Dim chName As String = mScript.mainClass(classId).ChildProperties(parentId).ElementAt(i).Key
                        lst.Add(chName, New MatewScript.PropertiesInfoType With {.Value = ch.ThirdLevelProperties(child3Id), .eventId = ch.ThirdLevelEventId(child3Id)})
                    Next i
                    eventContainer.CutAllEvents(classId, elementName, parentId, child3Id)
                    itm.element = lst
                End If
        End Select
        lstRemoved.Add(itm)
        AAA()
    End Sub

    ''' <summary>
    ''' Восстанавливет удаленный элемент
    ''' </summary>
    ''' <param name="itemId">Id элемента в структуре удаленных</param>
    ''' <returns>True если ошибки не произошло</returns>
    Public Function RetreiveItem(itemId As Integer) As Boolean
        AAA()
        If IsNothing(lstRemoved) OrElse itemId > lstRemoved.Count - 1 Then Return False
        Dim itm As cRemovedItem = lstRemoved(itemId)
        Dim newName As String = itm.elementName
        Dim newNameNoWrap As String = newName
        Dim isNewName As Boolean = False
        Select Case itm.classId
            Case -2
                'Восстановление переменной
                frmMainEditor.ChangeCurrentClassToVariables()
                If mScript.csPublicVariables.lstVariables.ContainsKey(newName) Then
                    isNewName = True
                    newName = frmMainEditor.GetNewDefName("Variable", -1)
                    newNameNoWrap = newName
                End If
                frmMainEditor.AddElement("Variable", currentTreeView, itm.element, newName)
            Case -3
                'Восстановление функции
                frmMainEditor.ChangeCurrentClassToFunctions()
                If mScript.functionsHash.ContainsKey(newName) Then
                    isNewName = True
                    newName = frmMainEditor.GetNewDefName("Function", -1)
                    newNameNoWrap = newName
                End If
                frmMainEditor.AddElement("Function", currentTreeView, itm.element, newName)
            Case mScript.mainClassHash("L")
                'Восстановление локации (с действиями)
                Dim className As String = mScript.mainClass(itm.classId).Names(0)
                Dim parentId As Integer = -1

                'Получаем новое имя при восстановлении
                If GetSecondChildIdByName(newName, mScript.mainClass(itm.classId).ChildProperties) > -1 Then
                    isNewName = True
                    newName = WrapString(frmMainEditor.GetNewDefName(className, -1))
                End If
                'восстанавливаем группы дочерних элементов
                Dim wasGroupsRetreived As Boolean = cGroups.RetreiveChildrenGroups(itm, newName)
                'Переход к классу локаций
                frmMainEditor.ChangeCurrentClass(className, False, "")

                newNameNoWrap = mScript.PrepareStringToPrint(newName, Nothing, False)
                frmMainEditor.AddElement(className, currentTreeView, itm.element(0), newName)
                'восстанавливаем все сообытия
                eventContainer.RetreiveAllEvents(itm, newName)

                If itm.element.Count > 1 Then
                    'Восстанавливаем действия данной локации
                    Dim locName As String = newName
                    className = "A"
                    Dim classAct As Integer = mScript.mainClassHash("A")
                    'parentId = GetSecondChildIdByName(locName, mScript.mainClass(mScript.mainClassHash("L")).ChildProperties)
                    frmMainEditor.ChangeCurrentClass(className, False, locName)

                    For actId As Integer = 0 To itm.element(1).Length - 1
                        Dim chProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element(1)(actId)
                        newName = chProps("Name").Value
                        newNameNoWrap = mScript.PrepareStringToPrint(newName, Nothing, False)
                        frmMainEditor.AddElement(className, currentTreeView, itm.element(1)(actId), newName)
                        eventContainer.RetreiveAllEvents(New cRemovedItem With {.classId = classAct, .element = itm.element(1)(actId), .elementName = newName, .parentName = locName}, newName)
                    Next
                    If wasGroupsRetreived Then frmMainEditor.FillTree("A", currentTreeView, frmMainEditor.chkShowHidden.CheckState, Nothing, locName)
                    newName = locName 'Для корректного сообщения о переименовании при восстановлении
                End If
            Case mScript.mainClassHash("A")
                'Восстановление действий
                Dim className As String = mScript.mainClass(itm.classId).Names(0)
                Dim classL As Integer = mScript.mainClassHash("L")
                Dim parentId As Integer = -1
                If itm.parentName.Length > 0 Then
                    parentId = GetSecondChildIdByName(itm.parentName, mScript.mainClass(classL).ChildProperties)
                End If
                If parentId < 0 Then
                    MessageBox.Show("Ошибка при восстановлении действия. Вероятно, его локация также была удалена.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return False
                End If
                frmMainEditor.ChangeCurrentClass(className, False, itm.parentName)

                If GetSecondChildIdByName(newName, mScript.mainClass(itm.classId).ChildProperties) > -1 Then
                    isNewName = True
                    newName = WrapString(frmMainEditor.GetNewDefName(className, parentId))
                End If

                newNameNoWrap = mScript.PrepareStringToPrint(newName, Nothing, False)
                frmMainEditor.AddElement(className, currentTreeView, itm.element, newName)
                eventContainer.RetreiveAllEvents(itm, newName)
            Case Else
                'Восстановление других элементов
                Dim className As String = mScript.mainClass(itm.classId).Names(0)
                Dim parentId As Integer = -1
                If itm.parentName.Length > 0 Then
                    parentId = GetSecondChildIdByName(itm.parentName, mScript.mainClass(itm.classId).ChildProperties)
                End If
                Dim thirdLevel As Boolean = False
                If String.IsNullOrEmpty(itm.parentName) Then
                    'Элемент 2 уровня
                    'Получаем новое имя при восстановлении
                    If GetSecondChildIdByName(newName, mScript.mainClass(itm.classId).ChildProperties) > -1 Then
                        isNewName = True
                        newName = WrapString(frmMainEditor.GetNewDefName(className, parentId))
                    End If
                    'восстанавливаем группы дочерних элементов
                    If mScript.mainClass(itm.classId).LevelsCount = 2 Then cGroups.RetreiveChildrenGroups(itm, newName)
                    'Переход к классу восстанавлимоего элемента
                    frmMainEditor.ChangeCurrentClass(className, thirdLevel, "")
                Else
                    'Элемент 3 уровня
                    Dim res = GetSecondChildIdByName(itm.parentName, mScript.mainClass(itm.classId).ChildProperties)
                    If res < 0 Then
                        MessageBox.Show("Ошибка при восстановлении элемента. Вероятно, его родительский элемент 2 уровня также был удален.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Return False
                    End If

                    'Получаем новое имя при восстановлении
                    If GetThirdChildIdByName(newName, parentId, mScript.mainClass(itm.classId).ChildProperties) > -1 Then
                        isNewName = True
                        newName = WrapString(frmMainEditor.GetNewDefName(className, parentId))
                    End If

                    'Переход к классу восстанавлимоего элемента
                    thirdLevel = True
                    frmMainEditor.ChangeCurrentClass(className, thirdLevel, itm.parentName)
                End If
                newNameNoWrap = mScript.PrepareStringToPrint(newName, Nothing, False)
                frmMainEditor.AddElement(className, currentTreeView, itm.element, newName)
                eventContainer.RetreiveAllEvents(itm, newName)
        End Select
        lstRemoved.Remove(itm)
        'Переход к узлу восстановленного элемента
        Dim n As TreeNode = frmMainEditor.FindItemNodeByText(currentTreeView, newNameNoWrap)
        If IsNothing(n) = False Then currentTreeView.SelectedNode = n

        If isNewName Then
            MessageBox.Show("Имя удаленного элемента уже занято, поэтому он восстановлен под новым именем " + Chr(34) + newName + Chr(34) + ".", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        AAA()
        Return True
    End Function

    ''' <summary>
    ''' Вносит изменения в структуру удаленных элементов после изменения имени связанного с ними родителя. Это могут быть 1) элементы 3 уровня; 2) действия
    ''' </summary>
    ''' <param name="classId">Класс элемента, который был переименован</param>
    ''' <param name="oldName">Старое имя</param>
    ''' <param name="newName">Новое имя</param>
    ''' <remarks></remarks>
    Public Sub RenameParent(ByVal classId As Integer, ByVal oldName As String, ByVal newName As String)
        If classId < 0 OrElse IsNothing(lstRemoved) OrElse lstRemoved.Count = 0 Then Return

        Dim classAct As Integer = mScript.mainClassHash("A")
        If mScript.mainClass(classId).LevelsCount = 2 Then
            'Переименование родителя в удаленных элементах 3 уровня
            For i As Integer = 0 To lstRemoved.Count - 1
                If lstRemoved(i).classId <> classId Then Continue For
                If lstRemoved(i).parentName = oldName Then lstRemoved(i).parentName = newName
            Next
        ElseIf classId = mScript.mainClassHash("L") Then
            'Переименование локации-родителя в удаленных действиях
            For i As Integer = 0 To lstRemoved.Count - 1
                If lstRemoved(i).classId <> classAct Then Continue For
                If lstRemoved(i).parentName = oldName Then lstRemoved(i).parentName = newName
            Next
        End If
    End Sub

    ''' <summary>
    ''' Переименовывает свойство в структуре удаленных элементов
    ''' </summary>
    ''' <param name="classId">Класс элементов, в котором переименовывается свойство</param>
    ''' <param name="oldName">Старое имя свойства</param>
    ''' <param name="newName">Новое имя свойства</param>
    Public Sub RenameProperty(ByVal classId As Integer, ByVal oldName As String, ByVal newName As String)
        If classId < 0 OrElse IsNothing(lstRemoved) OrElse lstRemoved.Count = 0 Then Return

        Dim classL As Integer = mScript.mainClassHash("L")
        Dim classAct As Integer = mScript.mainClassHash("A")
        'Переименование свойства элемента
        For i As Integer = 0 To lstRemoved.Count - 1
            Dim itm As cRemovedItem = lstRemoved(i)

            If classId = classAct AndAlso itm.classId = classL Then
                'Переименование свойства в действиях локации
                If itm.element.Count = 1 Then Continue For 'действий нет
                Dim chProps() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element(1)
                For j As Integer = 0 To chProps.Length - 1
                    Dim act As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = chProps(j)
                    Dim p As MatewScript.ChildPropertiesInfoType = Nothing
                    If act.TryGetValue(oldName, p) Then
                        act.Remove(oldName)
                        act.Add(newName, p)
                    End If
                Next j
                Continue For
            End If

            If itm.classId <> classId Then Continue For 'не наш класс - дальше
            If classId = classL Then
                'Переименование свойства в локации
                Dim chProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element(0)
                Dim p As MatewScript.ChildPropertiesInfoType = Nothing
                If chProps.TryGetValue(oldName, p) Then
                    chProps.Remove(oldName)
                    chProps.Add(newName, p)
                End If
            Else
                If itm.parentName.Length = 0 OrElse itm.classId = classAct Then
                    'Переименование свойства в элементе 2 уровня
                    Dim chProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element
                    Dim p As MatewScript.ChildPropertiesInfoType = Nothing
                    If chProps.TryGetValue(oldName, p) Then
                        chProps.Remove(oldName)
                        chProps.Add(newName, p)
                    End If
                Else
                    'Переименование свойства в элементе 3 уровня
                    Dim chProps As SortedList(Of String, MatewScript.PropertiesInfoType) = itm.element
                    Dim p As MatewScript.PropertiesInfoType = Nothing
                    If chProps.TryGetValue(oldName, p) Then
                        chProps.Remove(oldName)
                        chProps.Add(newName, p)
                    End If
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Удаляет свойство из структуры удаленных элементов
    ''' </summary>
    ''' <param name="classId">Класс элементов, в котором удаляется свойство</param>
    ''' <param name="propName">Имя удаляемого свойства</param>
    Public Sub RemoveProperty(ByVal classId As Integer, ByVal propName As String)
        If classId < 0 OrElse IsNothing(lstRemoved) OrElse lstRemoved.Count = 0 Then Return

        Dim classL As Integer = mScript.mainClassHash("L")
        Dim classAct As Integer = mScript.mainClassHash("A")
        'Удаление свойства у элемента
        For i As Integer = 0 To lstRemoved.Count - 1
            Dim itm As cRemovedItem = lstRemoved(i)

            If classId = classAct AndAlso itm.classId = classL Then
                'удаление свойства в действиях локации
                If itm.element.Count = 1 Then Continue For 'действий нет
                Dim chProps() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element(1)
                For j As Integer = 0 To chProps.Length - 1
                    Dim act As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = chProps(j)
                    If act.ContainsKey(propName) Then act.Remove(propName)
                Next j
                Continue For
            End If

            If itm.classId <> classId Then Continue For 'не наш класс - дальше
            If classId = classL Then
                'удаление свойства в локации
                Dim chProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element(0)
                If chProps.ContainsKey(propName) Then chProps.Remove(propName)
            Else
                If itm.parentName.Length = 0 OrElse itm.classId = classAct Then
                    'удаление свойства в элементе 2 уровня
                    Dim chProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element
                    If chProps.ContainsKey(propName) Then chProps.Remove(propName)
                Else
                    'удаление свойства в элементе 3 уровня
                    Dim chProps As SortedList(Of String, MatewScript.PropertiesInfoType) = itm.element
                    If chProps.ContainsKey(propName) Then chProps.Remove(propName)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Добавляет свойство в структуру удаленных элементов
    ''' </summary>
    ''' <param name="classId">Класс элементов, в которое добавляется свойство</param>
    ''' <param name="propName">Имя добавляемого свойства</param>
    Public Sub AddProperty(ByVal classId As Integer, ByVal propName As String)
        If classId < 0 OrElse IsNothing(lstRemoved) OrElse lstRemoved.Count = 0 Then Return

        Dim classL As Integer = mScript.mainClassHash("L")
        Dim classAct As Integer = mScript.mainClassHash("A")
        'Добавление свойства элементу
        For i As Integer = 0 To lstRemoved.Count - 1
            Dim itm As cRemovedItem = lstRemoved(i)

            If classId = classAct AndAlso itm.classId = classL Then
                'Добавление свойства действиям локации
                If itm.element.Count = 1 Then Continue For 'действий нет
                Dim chProps() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element(1)
                For j As Integer = 0 To chProps.Length - 1
                    Dim act As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = chProps(j)
                    If act.ContainsKey(propName) = False Then act.Add(propName, New MatewScript.ChildPropertiesInfoType)
                Next j
                Continue For
            End If

            If itm.classId <> classId Then Continue For 'не наш класс - дальше
            If classId = classL Then
                'Добавление свойства в локации
                Dim chProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element(0)
                If chProps.ContainsKey(propName) = False Then chProps.Add(propName, New MatewScript.ChildPropertiesInfoType)
            Else
                If itm.parentName.Length = 0 OrElse itm.classId = classAct Then
                    'Добавление свойства в элемент 2 уровня
                    Dim chProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = itm.element
                    If chProps.ContainsKey(propName) = False Then chProps.Add(propName, New MatewScript.ChildPropertiesInfoType)
                Else
                    'Добавление свойства в элемент 3 уровня
                    Dim chProps As SortedList(Of String, MatewScript.PropertiesInfoType) = itm.element
                    If chProps.ContainsKey(propName) = False Then chProps.Add(propName, New MatewScript.PropertiesInfoType)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Вносит изменение в структуру удаленных после удаления одного из классов
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    Public Sub RemoveClass(ByVal classId As Integer)
        If IsNothing(lstRemoved) OrElse lstRemoved.Count = 0 Then Return

        For i As Integer = lstRemoved.Count - 1 To 0 Step -1
            If lstRemoved(i).classId > classId Then
                lstRemoved(i).classId -= 1
            ElseIf lstRemoved(i).classId = classId Then
                lstRemoved.RemoveAt(i)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Заполняет листбокс строками, описывающими удаленные элементы
    ''' </summary>
    ''' <param name="lst">Листбокс для заполнения</param>
    ''' <returns>False если список пуст</returns>
    Public Function FillListBoxWithItems(ByRef lst As ListBox) As Boolean
        lst.Items.Clear()
        If IsNothing(lstRemoved) OrElse lstRemoved.Count = 0 Then Return False

        For i As Integer = 0 To lstRemoved.Count - 1
            lst.Items.Add(lstRemoved(i).GetString)
        Next
        Return True
    End Function

    ''' <summary>Класс для хранения скриптов удаленных элементов</summary>
    Private Class EventContainerClass
        Public lstEvents As New SortedList(Of Integer, List(Of MatewScript.ExecuteDataType))

        ''' <summary>
        ''' Вставляет в структуру класса скрипт из указанного события
        ''' </summary>
        ''' <param name="eventId">Id, под которым записать событие</param>
        ''' <param name="exData">Содержимое скрипта для вставки</param>
        Private Sub Add(ByVal eventId As Integer, ByVal exData As List(Of MatewScript.ExecuteDataType))
            If lstEvents.ContainsKey(eventId) Then
                MessageBox.Show("Не удалось создать резервную копию события. Id события уже занято.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            lstEvents.Add(eventId, exData)
        End Sub

        ''' <summary>
        ''' Вырезает все события из EventRouter и перемещает их в данный класс на хранение
        ''' </summary>
        ''' <param name="classId">Id класса удаляемого элемента</param>
        ''' <param name="elementName">Имя удаляемого элемента</param>
        ''' <param name="child2Id">Id элемента 2 уровня. Если в elementName - имя элемента также 2 уровня, то может быть -1 (определится внутри процедуры)</param>
        ''' <param name="child3Id">Id элемента 3 уровня или -1, если удаляется элемент 2 уровня</param>
        Public Sub CutAllEvents(ByVal classId As Integer, ByVal elementName As String, ByVal child2Id As Integer, ByVal child3Id As Integer)
            If child3Id < 0 AndAlso child2Id < 0 Then
                child2Id = GetSecondChildIdByName(elementName, mScript.mainClass(classId).ChildProperties)
            End If
            If child2Id < 0 Then
                MessageBox.Show("Не удалось создать резервную копию события. Элемент 3 уровня не найден.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            Dim childProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classId).ChildProperties(child2Id)
            '13
            If IsNothing(childProps) OrElse childProps.Count = 0 Then Return
            For i As Integer = 0 To childProps.Count - 1
                'Перебираем все свойства удаляемого элемента
                Dim ch As MatewScript.ChildPropertiesInfoType = childProps.ElementAt(i).Value
                'Получаем eventId свойства
                Dim eventId As Integer = -1
                If child3Id > -1 Then
                    eventId = ch.ThirdLevelEventId(child3Id)
                Else
                    eventId = ch.eventId
                End If

                If eventId > 0 Then
                    'Перемещаем скрипт в данный класс на хранение
                    Add(eventId, mScript.eventRouter.GetExDataByEventId(eventId)) 'mScript.eventRouter.lstEvents(eventId))
                    mScript.eventRouter.RemoveEvent(eventId)
                End If

                'Удаляем события всех дочерних элементов 3 уровня, если удаляется их родитель
                If child3Id < 0 AndAlso IsNothing(ch.ThirdLevelEventId) = False Then
                    For j As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                        If ch.ThirdLevelEventId(j) > 0 Then
                            eventId = ch.ThirdLevelEventId(j)
                            Add(eventId, mScript.eventRouter.GetExDataByEventId(eventId)) ' mScript.eventRouter.lstEvents(eventId))
                            mScript.eventRouter.RemoveEvent(eventId)
                        End If
                    Next j
                End If
            Next
        End Sub

        ''' <summary>
        ''' Восстанавливает все события восстанавливаемого элемента
        ''' </summary>
        ''' <param name="itm">класс восстанавливаемого элемента</param>
        ''' <param name="newName">имя при восстановлении</param>
        Public Sub RetreiveAllEvents(ByRef itm As cRemovedItem, ByVal newName As String)
            Dim classAct As Integer = mScript.mainClassHash("A")
            If String.IsNullOrEmpty(itm.parentName) OrElse itm.classId = classAct Then
                'Восстановление событий
                Dim child2Id As Integer = GetSecondChildIdByName(newName, mScript.mainClass(itm.classId).ChildProperties)
                If child2Id < 0 Then
                    MessageBox.Show("Ошибка при восстановлении. Не удалось восстановить события элемента.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                For pId As Integer = 0 To mScript.mainClass(itm.classId).ChildProperties(child2Id).Count - 1
                    Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(itm.classId).ChildProperties(child2Id).ElementAt(pId).Value
                    Dim eventId As Integer = ch.eventId
                    If eventId > 0 AndAlso lstEvents.ContainsKey(eventId) Then
                        mScript.eventRouter.SetEventId(eventId, lstEvents(eventId))
                        lstEvents.Remove(eventId)
                    End If

                    'Восстановление событий всех дочерних элементов 3 уровня (если есть)
                    If IsNothing(ch.ThirdLevelEventId) = False Then
                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            eventId = ch.ThirdLevelEventId(child3Id)
                            If eventId > 0 Then
                                mScript.eventRouter.SetEventId(eventId, lstEvents(eventId))
                                lstEvents.Remove(eventId)
                            End If
                        Next child3Id
                    End If
                Next pId
            Else
                'Восстановление событий 3 уровня
                Dim child2Id As Integer = GetSecondChildIdByName(itm.parentName, mScript.mainClass(itm.classId).ChildProperties)
                If child2Id < 0 Then
                    MessageBox.Show("Ошибка при восстановлении. Не удалось восстановить события элемента.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim child3Id As Integer = GetThirdChildIdByName(newName, child2Id, mScript.mainClass(itm.classId).ChildProperties)
                If child2Id < 0 Then
                    MessageBox.Show("Ошибка при восстановлении. Не удалось восстановить события элемента.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                For pId As Integer = 0 To mScript.mainClass(itm.classId).ChildProperties(child2Id).Count - 1
                    Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(itm.classId).ChildProperties(child2Id).ElementAt(pId).Value
                    Dim eventId As Integer = ch.ThirdLevelEventId(child3Id)
                    If eventId > 0 AndAlso lstEvents.ContainsKey(eventId) Then
                        mScript.eventRouter.SetEventId(eventId, lstEvents(eventId))
                        lstEvents.Remove(eventId)
                    End If
                Next pId
            End If
        End Sub
    End Class
End Class
