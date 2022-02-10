''' <summary>Класс, хранящий все группы для всех классов</summary>
Public Class clsGroups
    ''' <summary>
    ''' Информация о группе
    ''' </summary>
    Public Class clsGroupInfo
        ''' <summary>Название группы</summary>
        Public Property Name As String = ""
        ''' <summary>Скрыта группа или нет</summary>
        Public Property Hidden As Boolean = False
        ''' <summary>Название иконки группы</summary>
        Public Property iconName As String
        ''' <summary>Относится ли группа к дочерним элементам 3 уровня (для 3-уровневых классов)</summary>
        Public Property isThirdLevelGroup As Boolean = False
        ''' <summary>Имя родителя: элемента 2 уровня для элементов 3-го и Id локации для действий</summary>
        Public Property parentName As String
    End Class
    ''' <summary>Ключ - имя класса, Value - список групп данного класса</summary>
    Public dictGroups As New Dictionary(Of String, List(Of clsGroupInfo))(StringComparer.CurrentCultureIgnoreCase)
    ''' <summary>Ключ - имя класса, Value - список групп данного класса</summary>
    Public dictRemoved As New Dictionary(Of String, List(Of clsGroupInfo))(StringComparer.CurrentCultureIgnoreCase) 'для хранения удаленных групп с возможностью восстановления

    ''' <summary>
    ''' Инициализирует класс
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PrepareGroups()
        dictGroups = New Dictionary(Of String, List(Of clsGroupInfo))
        dictRemoved = New Dictionary(Of String, List(Of clsGroupInfo))
        If IsNothing(mScript.mainClass) OrElse mScript.mainClass.Count = 0 Then Return

        For i As Integer = 0 To mScript.mainClass.Count - 1
            dictGroups.Add(mScript.mainClass(i).Names(0), New List(Of clsGroupInfo))
        Next
        dictGroups.Add("Variable", New List(Of clsGroupInfo))
        dictGroups.Add("Function", New List(Of clsGroupInfo))
    End Sub

    ''' <summary>Определяет существует ли группа у данного класса </summary>
    ''' <param name="strClass">Имя класса, для которого проверяется наличие группы</param>
    ''' <param name="strName">Строка с имененм групы, наличие которой проверяется</param>
    ''' <param name="thirdLevel">Происходит ли поиск среди групп для элементов 3 уровня</param>
    ''' <param name="parentName">Name родительского элемента (например, локации для действия)</param>
    Public Function IsGroupExist(ByVal strClass As String, ByVal strName As String, ByVal thirdLevel As Boolean, Optional ByVal parentName As String = "") As Boolean
        If String.IsNullOrWhiteSpace(strClass) OrElse String.IsNullOrWhiteSpace(strName) Then Return False
        If IsNothing(dictGroups) OrElse dictGroups.Count = 0 Then Return False
        If dictGroups.ContainsKey(strClass) = False Then Return False

        Dim lstGroups As List(Of clsGroupInfo) = dictGroups(strClass)
        If IsNothing(lstGroups) OrElse lstGroups.Count = 0 Then Return False
        If GetGroupIdByName(strClass, strName, thirdLevel, parentName) >= 0 Then Return True
        Return False
    End Function

    ''' <summary>
    ''' Создает новую группу
    ''' </summary>
    ''' <param name="strClass">Имя класса с данной группой</param>
    ''' <param name="strName">Строка с именем групы</param>
    ''' <param name="thirdLevel">Добавляем ли группу элементам 3 уровня</param>
    ''' <param name="hidden">Скрыть граппу или нет</param>
    ''' <param name="parentName">Name родительского элемента (например, локации для действия)</param>
    ''' <returns>-1 в случае ошибки. Иначе Id группы.</returns>
    Public Function AddGroup(ByVal strClass As String, ByVal strName As String, ByVal thirdLevel As Boolean, ByVal hidden As Boolean, Optional ByVal parentName As String = "") As Integer
        If String.IsNullOrWhiteSpace(strClass) OrElse String.IsNullOrWhiteSpace(strName) Then Return -1

        Dim lstGroups As List(Of clsGroupInfo) = Nothing
        If dictGroups.TryGetValue(strClass, lstGroups) = False Then Return -1
        Dim gId As Integer = GetGroupIdByName(strClass, strName, thirdLevel, parentName)
        If gId >= 0 Then Return gId
        lstGroups.Add(New clsGroupInfo With {.Name = strName, .isThirdLevelGroup = thirdLevel, .Hidden = hidden, .iconName = "groupDefault.png", .parentName = parentName})
        Return lstGroups.Count - 1
    End Function

    ''' <summary>
    ''' Удаляет группу
    ''' </summary>
    ''' <param name="strClass">Имя класса с данной группой</param>
    ''' <param name="strName">Строка с именем групы</param>
    ''' <param name="thirdLevel">Удаляем ли группу из элементов 3 уровня</param>
    ''' <param name="parentName">Name родительского элемента (например, локации для действия)</param>
    Public Function RemoveGroup(ByVal strClass As String, ByVal strName As String, ByVal thirdLevel As Boolean, Optional ByVal parentName As String = "") As Integer
        If String.IsNullOrWhiteSpace(strClass) OrElse String.IsNullOrWhiteSpace(strName) Then Return -1
        'получаем индекс удаляемого элемента
        Dim gId As Integer = GetGroupIdByName(strClass, strName, thirdLevel, parentName)
        If gId = -1 Then Return -1
        'получаем список элементов нужного класса
        Dim lstGroups As List(Of clsGroupInfo) = Nothing
        If dictGroups.TryGetValue(strClass, lstGroups) = False Then Return -1
        'собственно удаление
        lstGroups.RemoveAt(gId)
        Return 0
    End Function

    ''' <summary>
    ''' Изменяет имя группы
    ''' </summary>
    ''' <param name="strClass">Имя класса, для которого проверяется наличие группы</param>
    ''' <param name="oldName">Старое имя группы</param>
    ''' <param name="newName">Новое имя группы</param>
    ''' <param name="thirdLevel">Происходит ли смена имени среди групп для элементов 3 уровня</param>
    ''' <param name="parentName">Name родительского элемента (например, локации для действия)</param>
    ''' <returns>-1 в случае ошибки. Иначе 0.</returns>
    Public Function ChangeGroupName(ByVal strClass As String, ByVal oldName As String, ByVal newName As String, ByVal thirdLevel As Boolean, Optional ByVal parentName As String = "") As Integer
        If String.IsNullOrWhiteSpace(strClass) OrElse String.IsNullOrWhiteSpace(oldName) OrElse String.IsNullOrWhiteSpace(newName) Then Return -1
        Dim lstGroups As List(Of clsGroupInfo) = Nothing
        If dictGroups.TryGetValue(strClass, lstGroups) = False Then Return -1
        If IsNothing(lstGroups) OrElse lstGroups.Count = 0 Then Return -1
        If thirdLevel AndAlso String.IsNullOrEmpty(parentName) Then Return -1

        Dim itmPos As Integer = GetGroupIdByName(strClass, oldName, thirdLevel, parentName)
        If itmPos < 0 Then Return -1

        lstGroups.Item(itmPos).Name = newName

        'переименовываем группу в дочерних элементах 2-го / 3-го порядка
        If strClass = "Variable" Then
            mScript.csPublicVariables.RenameVariablesGroup(oldName, newName)
            Return 0
        ElseIf strClass = "Function" Then
            If IsNothing(mScript.functionsHash) OrElse mScript.functionsHash.Count = 0 Then Return 0
            For i As Integer = 0 To mScript.functionsHash.Count - 1
                Dim f As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(i).Value
                If String.Compare(f.Group, oldName, True) = 0 Then
                    f.Group = newName
                End If
            Next
            Return 0
        End If

        Dim classId As Integer = mScript.mainClassHash(strClass)
        If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Return 0

        oldName = "'" + oldName + "'"
        newName = "'" + newName + "'"

        If Not thirdLevel Then
            'переименовываем группу в элементах 2-го порядка
            For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(i)("Group")
                If p.Value = oldName Then
                    p.Value = newName
                End If
            Next
        Else
            Dim parentId As Integer = GetSecondChildIdByName(parentName, mScript.mainClass(classId).ChildProperties)
            If parentId < 0 Then Return 0
            If IsNothing(mScript.mainClass(classId).ChildProperties(parentId)("Group").ThirdLevelProperties) OrElse _
               mScript.mainClass(classId).ChildProperties(parentId)("Group").ThirdLevelProperties.Count = 0 Then Return 0

            'переименовываем группу в элементах 3-го порядка
            For i As Integer = 0 To mScript.mainClass(classId).ChildProperties(parentId)("Group").ThirdLevelProperties.Count - 1
                Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(parentId)("Group")
                If p.ThirdLevelProperties(i) = oldName Then
                    p.ThirdLevelProperties(i) = newName
                End If
            Next
        End If

        Return 0
    End Function

    ''' <summary>
    ''' Изменяет иконку группы
    ''' </summary>
    ''' <param name="strClass">Имя класса, для которого проверяется наличие группы</param>
    ''' <param name="groupName">Имя группы</param>
    ''' <param name="iconName">Название новой иконки</param>
    ''' <param name="thirdLevel">Происходит ли смена иконки среди групп для элементов 3 уровня</param>
    ''' <param name="parentName">Name родительского элемента (например, локации для действия)</param>
    ''' <returns>-1 в случае ошибки. Иначе 0.</returns>
    Public Function ChangeGroupIcon(ByVal strClass As String, ByVal groupName As String, ByVal iconName As String, thirdLevel As Boolean, Optional ByVal parentName As String = "") As Integer
        If String.IsNullOrWhiteSpace(strClass) OrElse String.IsNullOrWhiteSpace(groupName) OrElse String.IsNullOrWhiteSpace(iconName) Then Return -1
        Dim lstGroups As List(Of clsGroupInfo) = Nothing
        If dictGroups.TryGetValue(strClass, lstGroups) = False Then Return -1
        If IsNothing(lstGroups) OrElse lstGroups.Count = 0 Then Return -1
        'получаем группу по ее имени
        Dim groupId As Integer = GetGroupIdByName(strClass, groupName, thirdLevel, parentName)
        If groupId < 0 Then Return -1
        lstGroups.Item(groupId).iconName = iconName 'устанавливаем иконку
        Return 0
    End Function

    ''' <summary>
    ''' Возвращает порядковый номер (от 0) группы в массиве lstGroups
    ''' </summary>
    ''' <param name="strClass">Имя класса с данной группой</param>
    ''' <param name="strName">Строка с именем групы</param>
    ''' <param name="thirdLevel">Искать ли группу из элементов 3 уровня</param>
    ''' <param name="parentName">Имя  родительского элемента (например, локации для действия)</param>
    Public Function GetGroupIdByName(ByVal strClass As String, ByVal strName As String, ByVal thirdLevel As Boolean, Optional ByVal parentName As String = "") As Integer
        If String.IsNullOrWhiteSpace(strClass) OrElse String.IsNullOrWhiteSpace(strName) Then Return -1

        Dim lstGroups As List(Of clsGroupInfo) = Nothing
        If dictGroups.TryGetValue(strClass, lstGroups) = False Then Return -1
        If IsNothing(lstGroups) OrElse lstGroups.Count = 0 Then Return -1

        For i As Integer = 0 To lstGroups.Count - 1
            Dim g As clsGroupInfo = lstGroups(i)
            If g.Name = strName AndAlso g.isThirdLevelGroup = thirdLevel Then
                If String.IsNullOrEmpty(parentName) Then Return i
                If parentName = g.parentName Then Return i
            End If
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Пермещает группу на новое место (в пределах того же класса)
    ''' </summary>
    ''' <param name="strClass">Имя класса с данной группой</param>
    ''' <param name="strName">Строка с именем групы</param>
    ''' <param name="elementToPlaceAfter">Группа, после которой на до разместить нашу группу</param>
    ''' <param name="thirdLevel">Перемещать ли группу среди элементов 3 уровня</param>
    ''' <param name="parentName">Name родительского элемента (например, Id локации для действия)</param>
    Public Sub MoveGroup(ByVal strClass As String, ByVal strName As String, ByVal elementToPlaceAfter As String, ByVal thirdLevel As Boolean, Optional ByVal parentName As String = "")
        'получаем id перемещаемой группы
        Dim gId As Integer = GetGroupIdByName(strClass, strName, thirdLevel, parentName)
        'получаем id группы, на место которой станет перемещаемая группа (это же новый id перемещаемой группы)
        Dim elementToPlaceId As Integer = GetGroupIdByName(strClass, elementToPlaceAfter, thirdLevel, parentName)
        If gId = -1 OrElse elementToPlaceId = -1 Then Return
        'создаем группу, идентичную перемещаемой, на новом месте
        Dim lstGroups As List(Of clsGroupInfo) = dictGroups(strClass)
        If elementToPlaceId >= lstGroups.Count Then
            lstGroups.Add(New clsGroupInfo With {.Name = strName, .Hidden = lstGroups(gId).Hidden, .isThirdLevelGroup = thirdLevel, .iconName = lstGroups(gId).iconName, .parentName = parentName})
        Else
            lstGroups.Insert(elementToPlaceId, New clsGroupInfo With {.Name = strName, .Hidden = lstGroups(gId).Hidden, .isThirdLevelGroup = thirdLevel, .iconName = lstGroups(gId).iconName, .parentName = parentName})
        End If
        If gId > elementToPlaceId Then gId += 1
        lstGroups.RemoveAt(gId) 'убираем группу со старого места
    End Sub

    ''' <summary>
    ''' Удаляет группы дочерних элементов 3 уровня данного класса при удалении элемента 2 уровня (или группы действий при удалении локации)
    ''' </summary>
    ''' <param name="className">Имя класса</param>
    ''' <param name="removedName">Имя удалеямого элемента 2 уровня</param>
    ''' <remarks></remarks>
    Public Sub RemoveChildrenGroups(ByVal className As String, ByVal removedName As String)
        'Получаем список групп нужного класса
        If className = "L" Then className = "A"
        Dim lstGroups As List(Of clsGroupInfo) = dictGroups(className)
        If IsNothing(lstGroups) OrElse lstGroups.Count = 0 Then Return
        'Получаем / создаем и получаем ссылку на список групп для хранения удаленных
        Dim grRem As List(Of clsGroupInfo) = Nothing
        If dictRemoved.TryGetValue(className, grRem) = False Then
            grRem = New List(Of clsGroupInfo)
            dictRemoved.Add(className, grRem)
        End If
        'Вносим изменения
        For gId As Integer = lstGroups.Count - 1 To 0 Step -1
            If lstGroups(gId).parentName = removedName Then
                'Удаляем группу и вносим ее в список удаленных
                grRem.Add(lstGroups(gId))
                lstGroups.RemoveAt(gId)
            End If
        Next gId
    End Sub

    ''' <summary>
    ''' Создаем дубликаты групп при дублировании дочерних элементов из одного родителя к другому
    ''' </summary>
    ''' <param name="className">Имя класса, внутри которого происходит копирование</param>
    ''' <param name="srcChild2Name">Имя родителя-донора</param>
    ''' <param name="destChild2Name">Имя родителя-реципиента</param>
    ''' <param name="blnReplace">Копировать с заменой или добавлять к существующим</param>
    Public Sub GroupsCopyTo(ByVal className As String, ByVal srcChild2Name As String, ByVal destChild2Name As String, ByVal blnReplace As Boolean)
        If className = "L" Then className = "A"
        Dim blnThirdLevel As Boolean = Not className = "A"

        Dim lstGroups As List(Of clsGroupInfo) = dictGroups(className)
        If IsNothing(lstGroups) OrElse lstGroups.Count = 0 Then Return

        If blnReplace Then
            'если копирование с заменой, то удаляем группы реципиента
            For gId As Integer = lstGroups.Count - 1 To 0 Step -1
                If lstGroups(gId).parentName = destChild2Name Then
                    'Удаляем группу 
                    lstGroups.RemoveAt(gId)
                End If
            Next gId
        End If

        If lstGroups.Count = 0 Then Return
        Dim destGid As Integer = -1 'Id группы родителя-реципиента, которую заменяем
        For gId As Integer = 0 To lstGroups.Count - 1
            'Просматриваем все группы класса в поисках групп донора
            Dim g As clsGroupInfo = lstGroups(gId)
            If g.parentName = srcChild2Name Then
                'группа донора найдена
                If Not blnReplace Then destGid = GetGroupIdByName(className, g.Name, blnThirdLevel, destChild2Name)
                If destGid >= 0 Then
                    'также у реципиента имеется одноименная группа. Не создаем новую, а просто изменяем характеристики уже существующей группы реципиента
                    With lstGroups(destGid)
                        .Hidden = g.Hidden
                        .iconName = g.iconName
                    End With
                Else
                    'Создаем реципиенту группу, аналогичную группе донора
                    Dim newGId As Integer = AddGroup(className, g.Name, blnThirdLevel, g.Hidden, destChild2Name)
                    lstGroups(newGId).iconName = g.iconName
                End If
            End If
        Next gId
    End Sub

    ''' <summary>
    ''' Переименовывает родителя в группах элементов 3 уровня + действий
    ''' </summary>
    ''' <param name="className">Класс, где происходит переименовка</param>
    ''' <param name="oldName">Старое имя родителя</param>
    ''' <param name="newName">Новое имя родителя</param>
    ''' <remarks></remarks>
    Public Sub RenameGroupInChilren(ByVal className As String, ByVal oldName As String, ByVal newName As String)
        'Получаем список групп нужного класса
        If className = "L" Then className = "A"
        Dim lstGroups As List(Of clsGroupInfo) = dictGroups(className)
        If IsNothing(lstGroups) OrElse lstGroups.Count = 0 Then Return

        For gId As Integer = 0 To lstGroups.Count - 1
            If lstGroups(gId).parentName = oldName Then lstGroups(gId).parentName = newName
        Next
    End Sub

    ''' <summary>
    ''' Восстанавливает действия дочерних элементов при восстановлении родителя
    ''' </summary>
    ''' <param name="itm">Класс восстанавливаемого элемента</param>
    ''' <param name="newName">Имя при восстановлении</param>
    ''' <returns>True усли была восстановлена хоть одна группа</returns>
    Public Function RetreiveChildrenGroups(ByRef itm As cRemovedObjects.cRemovedItem, ByVal newName As String) As Boolean
        Dim className As String = mScript.mainClass(itm.classId).Names(0)
        If className = "L" Then className = "A"
        Dim remGroups As List(Of clsGroupInfo) = Nothing
        If dictRemoved.TryGetValue(className, remGroups) = False Then Return False
        Dim lstGroups As List(Of clsGroupInfo) = dictGroups(className)

        Dim wasFound As Boolean = False
        For i As Integer = remGroups.Count - 1 To 0 Step -1
            If remGroups(i).parentName = itm.elementName Then
                'Группа для восстановления найдена. Изменяем имя ее родителя на новое при восстановлении, удаляем из списка удаленных и восстанавливаем в обычный список
                remGroups(i).parentName = newName
                lstGroups.Add(remGroups(i))
                remGroups.RemoveAt(i)
                wasFound = True
            End If
        Next i
        Return wasFound
    End Function
End Class
