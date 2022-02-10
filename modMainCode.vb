Module modMainCode
    Public mScript As New MatewScript

    Public Const APP_HELP_PATH As String = "D:\Projects\MatewQuest2\MatewQuest2\bin\Debug\Help\"
    Public ProgramPath As String = "D:\Projects\MatewQuest2\MatewQuest2\bin\Debug\" 'Windows.Forms.Application.StartupPath
    Public ProgramPathHTML As String = "file:///D:/Projects/MatewQuest2/MatewQuest2/bin/Debug"

    Public Enum PropertiesOperationEnum
        PROPERTY_GET = 0
        PROPERTY_SET = 1
    End Enum

#Region "Routers"

    Public Function FunctionRouter(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        functionName = functionName.Trim
        'исполняем исполняемые строки в параметрах, если таковые есть
        If ExecuteParams(funcParams, arrParams) = False Then Return "#Error"
        'расширение функций
        Dim f As MatewScript.PropertiesInfoType = Nothing
        If mScript.mainClass(classId).Functions.TryGetValue(functionName, f) = False Then
            Return _ERROR("Функция " & functionName & "в классе " & mScript.mainClass(classId).Names.Last & "не найдена.", functionName)
        End If
        If f.UserAdded = False AndAlso f.eventId > 0 Then
            'расширение встроеной функции сушествует
            Dim res As String = mScript.eventRouter.RunEvent(f.eventId, funcParams, "Расширение функции " & functionName, False)
            'расширение что-то вернуло - значит саму функию не запускаем
            If String.IsNullOrEmpty(res) = False Then Return res
        End If

        Select Case mScript.mainClass(classId).Names(0)
            Case "L"
                Return classL_functions(classId, functionName, funcParams, arrParams)
            Case "O"
                Return classO_functions(classId, functionName, funcParams, arrParams)
            Case "A"
                Return classA_functions(classId, functionName, funcParams, arrParams)
            Case "M"
                Return classM_functions(classId, functionName, funcParams, arrParams)
            Case "Code"
                Return classScript_functions(classId, functionName, funcParams, arrParams)
            Case "Math"
                Return classMath_functions(classId, functionName, funcParams, arrParams)
            Case "S"
                Return classS_functions(classId, functionName, funcParams, arrParams)
            Case "D"
                Return classD_functions(classId, functionName, funcParams, arrParams)
            Case "Q"
                Return classQ_functions(classId, functionName, funcParams, arrParams)
            Case "File"
                Return classFile_functions(classId, functionName, funcParams, arrParams)
            Case "Arr"
                Return classArr_functions(classId, functionName, funcParams, arrParams)
            Case "T"
                Return classT_functions(classId, functionName, funcParams, arrParams)
            Case "Med"
                Return classMed_functions(classId, functionName, funcParams, arrParams)
            Case "H"
                Return classH_functions(classId, functionName, funcParams, arrParams)
            Case "Mg"
                Return classMg_functions(classId, functionName, funcParams, arrParams)
            Case "Ab"
                Return classAb_functions(classId, functionName, funcParams, arrParams)
            Case "Map"
                Return classMap_functions(classId, functionName, funcParams, arrParams)
            Case "B"
                Return classB_functions(classId, functionName, funcParams, arrParams)
            Case "LW"
                Return classLW_functions(classId, functionName, funcParams, arrParams)
            Case "AW"
            Case "OW"
            Case "DW"
            Case "MgW"
                Return classMW_functions(classId, functionName, funcParams, arrParams)
            Case "Cm"
                Return classCm_functions(classId, functionName, funcParams, arrParams)
            Case "Army"
                Return classArmy_functions(classId, functionName, funcParams, arrParams)
            Case Else
                'Пользовательский класс. Вызываем обработчик функции
                Return classUser_functions(classId, functionName, funcParams, arrParams)
        End Select

        Return "#Error"
    End Function

    ''' <summary>Для PropertiesRouter PROPERTY SET. Сохраняет предыдущее значение свойства, до его изменения</summary>
    Private PREV_VALUE As String = ""
    Public Function PropertiesRouter(ByVal classId As Integer, ByVal propertyName As String, ByRef funcParams() As String, ByRef arrParams() As String, _
                                     Optional ByVal PropertyOperation As PropertiesOperationEnum = PropertiesOperationEnum.PROPERTY_GET, Optional ByVal newValue As String = "", _
                                     Optional ByVal ignoreBattle As Boolean = False) As String
        propertyName = propertyName.Trim

        If mScript.mainClass(classId).Properties.ContainsKey(propertyName) = False Then
            mScript.LAST_ERROR = String.Format("Свойства {0} у класса {1} не существует!", propertyName, mScript.mainClass(classId).Names.Last)
            Return "#Error"
        End If

        If PropertyOperation = PropertiesOperationEnum.PROPERTY_GET Then
            'получаем child2Id и child3Id
            Dim child2Id As Integer = -1, child3Id As Integer = -1
            If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 AndAlso funcParams(0) <> "-1" Then
                If GVARS.G_ISBATTLE AndAlso mScript.mainClass(classId).Names(0) = "H" Then
                    child2Id = mScript.Battle.GetFighterByName(funcParams(0))
                    If child2Id < 0 Then Return _ERROR("Бойца " & funcParams(0) & " не существует!")
                Else
                    Dim child2 = ObjectId(classId, funcParams, False)
                    If child2 = "#Error" Then Return child2
                    child2Id = CInt(child2)
                    If funcParams.Count > 1 Then
                        Dim child3 As String = ObjectId(classId, funcParams, True)
                        If child3 = "#Error" Then Return child3
                        child3Id = CInt(child3)
                    End If
                End If
            End If

            Dim result As String = ""
            Dim reType As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL
            If ReadProperty(classId, propertyName, child2Id, child3Id, result, funcParams, reType) = False Then Return "#Error"

            If child2Id > -1 AndAlso String.IsNullOrEmpty(result) OrElse result = "''" Then
                If propertyName = "DescriptionTemplate" Then
                    'Если DescriptionTemplate способности /бойца пусто, то возвращаем значение по умолчанию
                    If mScript.mainClass(classId).Names(0) = "Ab" OrElse mScript.mainClass(classId).Names(0) = "H" Then
                        If ReadProperty(classId, propertyName, -1, -1, result, funcParams, reType) = False Then Return "#Error"
                    End If
                End If
            End If

            If reType = MatewScript.ReturnFormatEnum.TO_LONG_TEXT OrElse reType = MatewScript.ReturnFormatEnum.TO_CODE Then
                reType = mScript.Param_GetType(result)
                If reType = MatewScript.ReturnFormatEnum.ORIGINAL Then result = WrapString(result)
            End If

            Dim className As String = mScript.mainClass(classId).Names(0)
            If className = "B" Then
                result = ClassB_PropertiesGet(classId, propertyName, result, funcParams, arrParams)
            End If
            Return result

            'Return class_common_properties_Get(classId, propertyName, funcParams, arrParams)


            'Select Case mScript.mainClass(classId).Names(0)
            '    Case "L", "A", "M", "Q"
            '        Return class_common_properties_Get(classId, propertyName, funcParams, arrParams)
            '    Case Else
            '        'классы пользователя
            '        If IsNothing(funcParams) OrElse funcParams.GetUpperBound(0) = -1 OrElse funcParams(0) = "-1" Then
            '            Return mScript.mainClass(classId).Properties(propertyName).Value 'получаем свойство по-умолчанию (1 уровня)
            '        ElseIf funcParams.GetUpperBound(0) = 0 Then
            '            'Свойства объектов 2 уровня
            '            If IsNothing(mScript.mainClass(classId).ChildProperties) Then
            '                mScript.LAST_ERROR = "Нет ни одного объекта в классе " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.GetUpperBound(0)) + ", к которому можно обратиться."
            '                Return "#Error"
            '            End If
            '            Dim child2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            '            If child2Id = -1 Then
            '                mScript.LAST_ERROR = "Нет обекта с Id = " + funcParams(0) + "."
            '                Return "#Error"
            '            ElseIf child2Id = -2 Then
            '                mScript.LAST_ERROR = "Обекта с именем " + funcParams(0) + " не существует."
            '                Return "#Error"
            '            End If
            '            Return mScript.mainClass(classId).ChildProperties(child2Id)(propertyName).Value
            '        Else
            '            'свойства объктов 3 уровня
            '            Dim level2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            '            If level2Id = -1 Then
            '                mScript.LAST_ERROR = "Нет объекта с Id = " + funcParams(0) + "."
            '                Return "#Error"
            '            ElseIf level2Id = -2 Then
            '                mScript.LAST_ERROR = "Объекта с именем " + funcParams(0) + " не существует."
            '                Return "#Error"
            '            End If
            '            Dim level3Id As Integer = GetThirdChildIdByName(funcParams(1), level2Id, mScript.mainClass(classId).ChildProperties)
            '            If level3Id = -1 Then
            '                mScript.LAST_ERROR = "Нет объекта с Id = [" + funcParams(0) + ", " + funcParams(1) + "]."
            '                Return "#Error"
            '            ElseIf level2Id = -2 Then
            '                mScript.LAST_ERROR = "Объекта с именем [" + funcParams(0) + ", " + funcParams(1) + "] не существует."
            '                Return "#Error"
            '            End If
            '            Return mScript.mainClass(classId).ChildProperties(level2Id)(propertyName).ThirdLevelProperties(level3Id)
            '        End If
            'End Select
        Else
            'Property SET
            If questEnvironment.EDIT_MODE = False AndAlso propertyName = "Name" Then
                Return _ERROR("Нельзя изменить имя элемента во время игры.")
            End If

            Dim isEvent As Boolean = False
            If (mScript.mainClass(classId).Properties(propertyName).returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT OrElse _
                mScript.mainClass(classId).Properties(propertyName).returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION) Then
                'Свойство - обработчик
                isEvent = True
                Dim res As String = mScript.eventRouter.SetPropertyWithEvent(classId, propertyName, newValue, funcParams, mScript.IsPropertyContainsCode(newValue), False, ignoreBattle)
                If res = "#Error" Then Return res
            End If

            If Not isEvent Then
                'событие отслеживания свойства
                Dim trackResult As String = mScript.trackingProperties.RunBefore(classId, propertyName, funcParams, newValue)
                If trackResult = "False" Then
                    Return "False"
                ElseIf trackResult = "#Error" Then
                    Return trackResult
                End If

                If IsNothing(funcParams) OrElse funcParams.Count = 0 Then
                    PREV_VALUE = mScript.mainClass(classId).Properties(propertyName).Value
                    SetPropertyValue(classId, propertyName, newValue, -1) 'Устанавливаются свойства по умолчанию
                ElseIf funcParams.Count = 1 Then
                    'Устанавливаются свойства элемента 2 уровня
                    Dim child2Id As Integer
                    If GVARS.G_ISBATTLE AndAlso ignoreBattle = False AndAlso mScript.mainClass(classId).Names(0) = "H" Then
                        child2Id = mScript.Battle.GetFighterByName(funcParams(0))
                        If child2Id < 0 Then Return _ERROR("Боец " & funcParams(0) & " не найден.")
                        PREV_VALUE = mScript.Battle.Fighters(child2Id).heroProps(propertyName).Value
                    Else
                        Dim elementId As String = ObjectId(classId, funcParams)
                        If elementId = "#Error" Then Return elementId
                        child2Id = Val(elementId)
                        PREV_VALUE = mScript.mainClass(classId).ChildProperties(child2Id)(propertyName).Value
                    End If
                    SetPropertyValue(classId, propertyName, newValue, child2Id, -1, ignoreBattle)
                Else
                    Dim res As String = ObjectId(classId, funcParams, False)
                    If res = "#Error" Then Return res
                    Dim elementId2 As Integer = CInt(res)
                    res = ObjectId(classId, funcParams, True)
                    If res = "#Error" Then Return res
                    Dim elementId3 As Integer = CInt(res)
                    PREV_VALUE = mScript.mainClass(classId).ChildProperties(elementId2)(propertyName).ThirdLevelProperties(elementId3)
                    SetPropertyValue(classId, propertyName, newValue, elementId2, elementId3)
                End If
            End If

            Dim className As String = mScript.mainClass(classId).Names(0)
            Return PropertiesRouterRunSpecific(classId, propertyName, newValue, funcParams, arrParams, False, ignoreBattle)
        End If
        Return "#Error"
    End Function

    Public Function PropertiesRouterRunSpecific(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String, _
                                                ByVal NewValueIsScriptOrLongText As Boolean, Optional ignoreBattle As Boolean = False) As String
        Select Case mScript.mainClass(classId).Names(0)
            Case "A"
                Return ClassA_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "L"
                Return ClassL_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams, NewValueIsScriptOrLongText)
            Case "O"
                Return ClassO_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "M"
                Return ClassM_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "T"
                Return ClassT_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "Med"
                Return ClassMed_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "Map"
                Return ClassMap_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "H"
                If GVARS.G_ISBATTLE AndAlso ignoreBattle = False AndAlso mScript.mainClass(classId).Names(0) = "H" Then
                    Return ClassFighter_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
                Else
                    Return ClassH_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
                End If
            Case "B"
                Return ClassB_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "Ab"
                Return ClassAb_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "Army"
                Return ClassAmy_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "AW", "OW", "DW"
                Return ClassAW_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case "Cm"
                Return ClassCm_PropertiesSet(classId, propertyName, newValue, funcParams, arrParams)
            Case Else
                Return ""
        End Select

    End Function


    'Private Function class_common_properties_Get(ByVal classId As Integer, ByVal propertyName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String

    '    If IsNothing(funcParams) OrElse funcParams.GetUpperBound(0) = -1 OrElse funcParams(0) = "-1" Then
    '        'получаем значение свойства по умолчанию / глобального обработчика
    '        Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties(propertyName)
    '        Dim propValue As String = p.Value  'возвращаем свойство 1-го порядка (по умолчанию)
    '        'если внутри не код - возвращаем значение свойства
    '        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(propValue)
    '        If cRes = MatewScript.ContainsCodeEnum.NOT_CODE Then Return propValue
    '        'Если внутри код/исполняемая строка - выполняем код (если в события код свойства еще не добавлен, то перед исполнением добавляем)
    '        If Not mScript.eventRouter.IsEventIdExists(p.eventId) Then p.eventId = mScript.eventRouter.SetEventId(classId, propertyName, propValue, cRes, -1, -1)
    '        If p.eventId > 0 Then
    '            propValue = mScript.eventRouter.RunEvent(p.eventId, arrParams)
    '        End If
    '        Return WrapString(propValue)
    '    Else
    '        If mScript.mainClass(classId).LevelsCount = 0 Then
    '            mScript.LAST_ERROR = "Класс " + mScript.mainClass(classId).Names.Last + " не содержит элементов 2-го и третьего порядков."
    '            Return "#Error"
    '        End If
    '        'получаем Id элемента 2-го порядка
    '        If IsNothing(mScript.mainClass(classId).ChildProperties) Then
    '            mScript.LAST_ERROR = "Нет ни одного элемента, к которому можно обратиться."
    '            Return "#Error"
    '        End If
    '        Dim child2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
    '        If child2Id = -1 Then
    '            mScript.LAST_ERROR = "Нет элемента с Id = " + funcParams(0) + "."
    '            Return "#Error"
    '        ElseIf child2Id = -2 Then
    '            mScript.LAST_ERROR = "Элемента с именем " + funcParams(0) + " не существует."
    '            Return "#Error"
    '        End If
    '        Dim child3Id As Integer = -1

    '        'получаем значение свойства
    '        Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(propertyName)
    '        Dim propValue As String   'возвращаем свойство 1-го порядка (по умолчанию)
    '        If funcParams.Count = 1 Then
    '            propValue = p.Value 'возвращаем свойство 2-го порядка
    '        Else
    '            If mScript.mainClass(classId).LevelsCount = 0 Then
    '                mScript.LAST_ERROR = "Класс " + mScript.mainClass(classId).Names.Last + " не содержит элементов третьего порядка."
    '                Return "#Error"
    '            End If
    '            'получаем Id элемента 3-го порядка
    '            child3Id = GetThirdChildIdByName(funcParams(0), child2Id, mScript.mainClass(classId).ChildProperties)
    '            If child3Id = -1 Then
    '                mScript.LAST_ERROR = "Нет элемента с Id = " + funcParams(0) + "."
    '                Return "#Error"
    '            ElseIf child3Id = -2 Then
    '                mScript.LAST_ERROR = "Элемента с именем " + funcParams(0) + " не существует."
    '                Return "#Error"
    '            End If
    '            propValue = p.ThirdLevelProperties(child3Id) 'возвращаем свойство 3-го порядка
    '        End If

    '        'если внутри не код - возвращаем значение свойства
    '        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(propValue)
    '        If cRes = MatewScript.ContainsCodeEnum.NOT_CODE Then Return propValue
    '        'Если внутри код/исполняемая строка - выполняем код (если в события код свойства еще не добавлен, то перед исполнением добавляем)
    '        If Not mScript.eventRouter.IsEventIdExists(p.eventId) Then p.eventId = mScript.eventRouter.SetEventId(classId, propertyName, propValue, cRes, child2Id, child3Id)
    '        If p.eventId > 0 Then
    '            propValue = mScript.eventRouter.RunEvent(p.eventId, arrParams)
    '        End If
    '        Return WrapString(propValue)

    '    End If

    '    Return "#Error"
    'End Function

#End Region

#Region "Functions"

#Region "Actions"
    Private Function classA_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Add"
                'Создает в указанном классе объект втого порядка
                '0 - Name/id; 1 - position; 2 - after visiting; 3 - видимость
                Dim res As String = CreateNewObject(classId, funcParams, True)
                If res = "#Error" Then Return res
                Dim newId As Integer = CInt(res)
                'funcParams(1) - position
                Dim paramsCount As Integer = funcParams.Count
                Dim actionPos As Integer = -1
                If paramsCount > 1 Then
                    actionPos = CInt(funcParams(1))
                    If actionPos > -1 AndAlso actionPos < mScript.mainClass(classId).ChildProperties.Count Then
                        'действие надо поместить не в конец
                        Dim lst As List(Of SortedList(Of String, MatewScript.ChildPropertiesInfoType)) = mScript.mainClass(classId).ChildProperties.ToList
                        Dim itm As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = lst(newId)
                        lst.Remove(itm)
                        lst.Insert(actionPos, itm)
                        mScript.mainClass(classId).ChildProperties = lst.ToArray
                        newId = actionPos
                    End If
                    '2 - после посещения
                    If paramsCount >= 3 Then
                        Dim afterVis As Integer = GetEnumIdByName(classId, functionName, 2, funcParams(2))
                        If afterVis < 0 Then afterVis = 0
                        SetPropertyValue(classId, "AfterVisiting", afterVis.ToString, newId)
                        '3 - Visible
                        If paramsCount >= 4 Then
                            Dim vis As Boolean = funcParams(3)
                            SetPropertyValue(classId, "Visible", afterVis.ToString, newId)
                        End If
                    End If
                End If

                If questEnvironment.EDIT_MODE = False Then
                    'создаем html-элемент
                    Dim hPrev As HtmlElement = Nothing
                    If actionPos > -1 AndAlso actionPos < ObjectsCount(classId, funcParams) Then
                        hPrev = FindActionHTMLelement(classId, actionPos, funcParams)
                    End If
                    ActionCreateHtmlElement(newId, hPrev)
                End If

                Return CStr(newId)
            Case "AddGoTo"
                'Создает действие с указанием свойства GoTo
                '0 - Name; 1 - position; 2 - location
                Dim res As String = CreateNewObject(classId, funcParams, True)
                If res = "#Error" Then Return res
                Dim newId As Integer = CInt(res)
                'funcParams(1) - position
                Dim paramsCount As Integer = funcParams.Count
                Dim actionPos As Integer = -1
                If paramsCount >= 2 Then
                    '2 - location
                    Dim gotoLoc As String = funcParams(1)
                    SetPropertyValue(classId, "GoTo", gotoLoc, newId)

                    If paramsCount >= 3 Then
                        actionPos = CInt(funcParams(2))
                        If actionPos > -1 AndAlso actionPos < ObjectsCount(classId, funcParams) Then
                            'действие надо поместить не в конец
                            Dim lst As List(Of SortedList(Of String, MatewScript.ChildPropertiesInfoType)) = mScript.mainClass(classId).ChildProperties.ToList
                            Dim itm As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = lst(newId)
                            lst.Remove(itm)
                            lst.Insert(actionPos, itm)
                            mScript.mainClass(classId).ChildProperties = lst.ToArray
                        End If
                    End If
                End If

                If questEnvironment.EDIT_MODE = False Then
                    'создаем html-элемент
                    Dim hPrev As HtmlElement = Nothing
                    If actionPos > -1 AndAlso actionPos < ObjectsCount(classId, funcParams) Then
                        hPrev = FindActionHTMLelement(classId, actionPos, funcParams)
                    End If
                    ActionCreateHtmlElement(newId, hPrev)
                End If

                Return CStr(newId)

            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "Count"
                Return ObjectsCount(classId, funcParams, False)
            Case "CheckShowQueries"
                ActionCheckShowQueries()
                Return ""
            Case "IsExist"
                'Определяет есть ли действие с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Remove"
                If questEnvironment.EDIT_MODE = False Then
                    'удаляем html-элемент(ы)
                    If funcParams.Count = 0 OrElse funcParams(0) = "-1" Then
                        'удаляем все действия
                        If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso mScript.mainClass(classId).ChildProperties.Count > 0 Then
                            For actId As Integer = mScript.mainClass(classId).ChildProperties.Count - 1 To 0 Step -1
                                Dim hAct As HtmlElement = FindActionHTMLelement(classId, actId, funcParams)
                                If IsNothing(actId) = False Then
                                    Dim hNod As mshtml.IHTMLDOMNode = Nothing  ' = hAct.DomElement
                                    If IsNothing(hAct) = False Then
                                        hNod = hAct.DomElement
                                        hNod.removeNode(True)
                                    End If
                                End If
                            Next
                        End If
                    Else
                        'удаляем указанное действие
                        Dim actId As Integer = ObjectId(classId, funcParams)
                        Dim hAct As HtmlElement = FindActionHTMLelement(classId, actId, funcParams)
                        If IsNothing(actId) = False Then
                            Dim hNod As mshtml.IHTMLDOMNode = hAct.DomElement
                            hNod.removeNode(True)
                        End If
                    End If
                End If
                'удаляем из структуры
                Return RemoveObject(classId, funcParams)
            Case "Load" 'Загружает сохраненные действия
                Dim saveName As String = "temp"
                If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then saveName = funcParams(0)
                If actionsRouter.GameLoadSavedActions(saveName) = "#Error" Then Return "#Error"
                'создаем все элементы
                If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Return "0"
                If questEnvironment.EDIT_MODE = False Then
                    For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        ActionCreateHtmlElement(i)
                    Next
                End If
                Return mScript.mainClass(classId).ChildProperties.Count.ToString
            Case "Save" 'Сохраняет текущий набор действий
                Dim saveName As String = "temp"
                If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then saveName = funcParams(0)
                actionsRouter.GameSaveActions(saveName)
                Return "1"
            Case "RemoveSaved" 'Удаляет сохраненные действие
                Dim delName As String = ""
                If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then delName = funcParams(0)
                actionsRouter.GameRemoveSavedActions(delName)
                Return "1"
            Case "Select" 'Эмулирует выбор действия
                If questEnvironment.EDIT_MODE Then Return ""
                Dim actId As Integer = CInt(ObjectId(classId, funcParams, False))
                If actId < 0 Then Return _ERROR("Не найдено действие " & funcParams(0), "Select")
                If funcParams.Count > 1 AndAlso funcParams(1) = "False" Then
                    'если второй параметр False, то проверяем на доступность/видимость
                    Dim enabled As Boolean = True
                    ReadPropertyBool(classId, "Enabled", actId, -1, enabled, funcParams)
                    If enabled = False Then Return "-1"
                    Dim visible As Boolean = True
                    ReadPropertyBool(classId, "Visible", actId, -1, visible, funcParams)
                    If visible = False Then Return "-1"
                End If

                Dim hAct As HtmlElement = FindActionHTMLelement(classId, actId, funcParams)
                If IsNothing(hAct) Then
                    If ActionsInputProhibited Then Return _ERROR("Нельзя эмулировать выбор действия до окончательной загрузки локации." & funcParams(0), "Select")
                    Return _ERROR("На игровом поле не найдено действие " & funcParams(0), "Select")
                End If

                EventGeneratedFromScript = True
                del_action_Click(hAct, Nothing)
                Return "1"
            Case "LoadFromLocation"
                Dim classL As Integer = mScript.mainClassHash("L")
                Dim locId As Integer = ObjectId(classL, funcParams, False)
                If IsNothing(mScript.mainClass(classL).ChildProperties) OrElse locId > mScript.mainClass(classL).ChildProperties.Count - 1 Then Return _ERROR("Локации с Id " & locId.ToString & " не существует!", "LoadFromLocation")
                'удаляем старые
                If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso mScript.mainClass(classId).ChildProperties.Count > 0 Then
                    For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        Dim actConvas As String = ""
                        If ReadProperty(classId, "HTMLContainerId", i, -1, actConvas, {i.ToString}) Then
                            actConvas = UnWrapString(actConvas)
                            If String.IsNullOrEmpty(actConvas) = False Then
                                'удаляем элемент
                                Dim hAct As mshtml.IHTMLDOMNode = FindActionHTMLelementDOM(classId, i, {i.ToString})
                                If IsNothing(hAct) = False Then hAct.removeNode(True)
                            End If
                        End If
                    Next
                    Dim hConvas As HtmlElement = frmPlayer.wbActions.Document.GetElementById("ActionsConvas")
                    If IsNothing(hConvas) = False Then hConvas.InnerHtml = ""
                End If
                'загружаем из структуры
                actionsRouter.GameLoadActions(mScript.mainClass(classL).ChildProperties(locId)("Name").Value)
                'создаем все элементы                
                If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Return "0"
                If questEnvironment.EDIT_MODE = False Then
                    'создаем новые
                    For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        ActionCreateHtmlElement(i)
                    Next
                End If
                Return mScript.mainClass(classId).ChildProperties.Count.ToString
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function ClassAW_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        Select Case propertyName
            Case "Visible", "Width", "Height"
                SetWindowsContainers()
        End Select
        Return ""
    End Function


    Private Function ClassA_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        Dim actId As Integer = -1
        If funcParams.Count > 0 Then
            If funcParams(0) = "-1" Then Return ""
            actId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If actId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует имени и Id действия {0}!", funcParams(0))
                Return "#Error"
            End If
        Else
            Return ""
        End If
        'получаем имя действия
        Dim actName As String = mScript.mainClass(classId).ChildProperties(actId)("Name").Value

        Select Case propertyName
            Case "Caption", "Picture", "PictureFloat"
                'изменяем имя действия - Name html-элемента
                Dim hAct As HtmlElement = FindActionHTMLelement(classId, actId, funcParams)
                If IsNothing(hAct) Then Return ""
                Dim actType As Integer, actCaption As String = "", actPicture As String = "", actPictureFloat As Integer
                If ReadPropertyInt(classId, "Type", actId, -1, actType, funcParams) = False Then Return "#Error"
                If ReadProperty(classId, "Caption", actId, -1, actCaption, funcParams) = False Then Return "#Error"
                If ReadProperty(classId, "Picture", actId, -1, actPicture, funcParams) = False Then Return "#Error"
                If ReadPropertyInt(classId, "PictureFloat", actId, -1, actPictureFloat, funcParams) = False Then Return "#Error"
                If actType = 1 Then
                    hAct.SetAttribute("Title", UnWrapString(actCaption))
                    hAct.SetAttribute("src", UnWrapString(actPicture).Replace("\", "/"))
                Else
                    ActionSetInnerHTML(hAct, actType, UnWrapString(actCaption), UnWrapString(actPicture).Replace("\", "/"), actPictureFloat)
                End If
                Return ""
            Case "HTMLContainerId"
                Dim hAct As HtmlElement = FindActionHTMLelement(classId, actId, funcParams, PREV_VALUE)
                If IsNothing(hAct) Then Return ""
                Dim actNod As mshtml.IHTMLDOMNode = hAct.DomElement
                actNod.removeNode(True)
                ActionCreateHtmlElement(actId)
            Case "Type"
                Dim hAct As HtmlElement = FindActionHTMLelement(classId, actId, funcParams)
                If IsNothing(hAct) Then Return ""
                ActionCreateHtmlElement(actId, hAct)
                Dim hdAct As mshtml.IHTMLDOMNode = hAct.DomElement
                hdAct.removeNode(True)
            Case "Enabled"
                'делаем действие недоступным
                If questEnvironment.EDIT_MODE Then Return ""
                Dim enabled As Boolean = True
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                Boolean.TryParse(nValue, enabled)
                If enabled Then
                    Dim hAct As mshtml.IHTMLElement = FindActionHTMLelementDOM(classId, actId, funcParams)
                    If IsNothing(hAct) Then Return ""
                    hAct.removeAttribute("disabled")
                Else
                    Dim hAct As HtmlElement = FindActionHTMLelement(classId, actId, funcParams)
                    If IsNothing(hAct) Then Return ""
                    hAct.SetAttribute("disabled", "True")
                    'hAct.Children(0).SetAttribute("disabled", "true")
                End If
                Return ""
            Case "Visible"
                Dim hAct As HtmlElement = FindActionHTMLelement(classId, actId, funcParams)
                If IsNothing(hAct) Then Return "#"
                'делаем действие видимым/невидимым
                Dim visible As Boolean = True
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                Boolean.TryParse(nValue, visible)
                If visible Then
                    hAct.Style = ""
                Else
                    hAct.Style = "display:none"
                End If
                Return ""
        End Select
        Return ""
    End Function

    ''' <summary>
    ''' Получает html-элемент, связанный с данным действием
    ''' </summary>
    ''' <param name="classId">класс Action</param>
    ''' <param name="actId">Id действия</param>
    ''' <param name="funcParams">Параметры, переданные функции</param>
    Public Function FindActionHTMLelementDOM(ByVal classId As Integer, ByVal actId As Integer, ByRef funcParams() As String) As mshtml.IHTMLElement
        'получаем Id контейнера
        If ActionsInputProhibited Then Return Nothing
        Dim actContainer As String = ""
        ReadProperty(classId, "HTMLContainerId", actId, -1, actContainer, funcParams)
        actContainer = UnWrapString(actContainer)
        Dim hContainer As mshtml.IHTMLElement
        Dim actName As String = mScript.mainClass(classId).ChildProperties(actId)("Name").Value
        If String.IsNullOrEmpty(actContainer) Then
            hContainer = frmPlayer.wbActions.Document.DomDocument.GetElementById("ActionsConvas")
        Else
            hContainer = frmPlayer.wbMain.Document.DomDocument.GetElementById(actContainer)
        End If
        If IsNothing(hContainer) Then
            mScript.LAST_ERROR = String.Format("Не удалось найти html-контейнер {0} для размещения в нем действия {1}.", actContainer, actName)
            Return Nothing
        End If
        'получаем html-элемент с нашим действием
        Dim hAct As mshtml.IHTMLElement = Nothing
        For i As Integer = 0 To hContainer.children.length - 1
            Dim hEl As mshtml.IHTMLElement = hContainer.children(i)
            If String.Compare(hEl.getAttribute("Name"), actName, True) = 0 Then
                'найден
                hAct = hEl
                Exit For
            End If
        Next
        If IsNothing(hAct) Then
            mScript.LAST_ERROR = String.Format("Не удалось найти действие {0} на экране.", actName)
            Return Nothing
        End If
        Return hAct
    End Function

    ''' <summary>
    ''' Получает html-элемент указанного действия
    ''' </summary>
    ''' <param name="classId">класса А</param>
    ''' <param name="actId">Id действия</param>
    ''' <param name="funcParams"></param>
    ''' <param name="actContainer">html-контейнер</param>
    Public Function FindActionHTMLelement(ByVal classId As Integer, ByVal actId As Integer, ByRef funcParams() As String, Optional ByVal actContainer As String = Nothing) As HtmlElement
        'получаем Id контейнера
        If questEnvironment.EDIT_MODE Then Return Nothing
        If IsNothing(actContainer) Then
            ReadProperty(classId, "HTMLContainerId", actId, -1, actContainer, funcParams)
            actContainer = UnWrapString(actContainer)
        End If
        Dim hContainer As HtmlElement
        If String.IsNullOrEmpty(actContainer) Then
            hContainer = frmPlayer.wbActions.Document.GetElementById("ActionsConvas")
        Else
            hContainer = frmPlayer.wbMain.Document.GetElementById(actContainer)
        End If
        Dim actName As String = mScript.mainClass(classId).ChildProperties(actId)("Name").Value
        If IsNothing(hContainer) Then
            If ActionsInputProhibited = False Then mScript.LAST_ERROR = String.Format("Не удалось найти html-контейнер {0} для размещения в нем действия {1}.", actContainer, actName)
            Return Nothing
        End If
        'получаем html-элемент с нашим действием
        Dim hAct As HtmlElement = Nothing
        For i As Integer = 0 To hContainer.Children.Count - 1
            Dim hEl As HtmlElement = hContainer.Children(i)
            If String.Compare(hEl.GetAttribute("Name"), actName, True) = 0 Then
                'найден
                hAct = hEl
                Exit For
            End If
        Next
        If IsNothing(hAct) Then
            mScript.LAST_ERROR = String.Format("Не удалось найти действие {0} на экране.", actName)
            Return Nothing
        End If
        Return hAct
    End Function

#End Region

#Region "Objects"
    Private Function classO_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго порядка
                Dim res As String = CreateNewObject(classId, funcParams)
                If res = "#Error" Then Return res

                If questEnvironment.EDIT_MODE = False Then
                    'изменения в стурктуре окна предметов
                    Dim hDoc As HtmlDocument = frmPlayer.wbObjects.Document
                    If IsNothing(hDoc) = False Then
                        Dim hConvas As HtmlElement = hDoc.GetElementById("ObjectsConvas")
                        If IsNothing(hConvas) = False Then
                            Dim oId As Integer = 0
                            If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then oId = mScript.mainClass(classId).ChildProperties.Count - 1
                            'наружный контейнер предмета
                            Dim outerDiv As HtmlElement = hDoc.CreateElement("DIV")
                            outerDiv.SetAttribute("ClassName", "OuterContainer")
                            outerDiv.Id = "object" & oId.ToString
                            outerDiv.Style = "display:none"
                            hConvas.AppendChild(outerDiv)

                            'див для хранения содержимого предмета если он - контейнер
                            Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                            ReadPropertyInt(classId, "Type", oId, -1, objType, arrParams) '1 - контейнер
                            If objType = ObjectTypeEnum.CONTAINER Then
                                Dim containerDiv As HtmlElement = hDoc.CreateElement("DIV")
                                containerDiv.SetAttribute("ClassName", "InnerContainer")
                                outerDiv.AppendChild(containerDiv)
                            End If
                        End If
                    End If
                End If

                Return res
            Case "CountTotal"
                'Dim strCount As String = ObjectsCount(classId, funcParams, False)
                If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Return "0"
                Dim cnt As Integer = mScript.mainClass(classId).ChildProperties.Count
                Dim heroOwner As Boolean = CBool(GetParam(funcParams, 0, "True"))
                If Not heroOwner Then Return cnt.ToString

                For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                    Dim BelongsToPlayer As Boolean = False
                    ReadPropertyBool(classId, "BelongsToPlayer", i, -1, BelongsToPlayer, funcParams)
                    If Not BelongsToPlayer Then cnt -= 1
                Next
                Return cnt.ToString
            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "IsExist"
                'Определяет есть ли предмет с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "IsPlayerOwner"
                Dim oId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return "False"
                Dim BelongsToPlayer As Boolean = False
                If ReadPropertyBool(classId, "BelongsToPlayer", oId, -1, BelongsToPlayer, funcParams) = False Then Return "#Error"
                Return BelongsToPlayer.ToString
            Case "Flash"
                Dim oId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return "False"

                Dim hObj As HtmlElement = GetObjectOuterElement(oId)
                If IsNothing(hObj) Then Return "False"

                If IsNothing(hObj) Then Return _ERROR("Предмата с указанным Id не существует.", "Flash")

                'flashing
                Dim effDuration As Integer = 1000, effName As String = ""
                If ReadPropertyInt(classId, "FlashEffDuration", -1, -1, effDuration, arrParams) = False Then Return "#Error"
                If ReadProperty(classId, "FlashEffect", -1, -1, effName, arrParams) = False Then Return "#Error"
                effName = UnWrapString(effName)
                hObj.Style &= "animation-name:" & effName & ";animation-duration: " & effDuration.ToString & "ms;"
                mScript.Battle.Wait(effDuration)
                HTMLRemoveCSSstyle(hObj, "animation-name")
                Return True
            Case "Remove"
                'удаление предмета
                If questEnvironment.EDIT_MODE = False Then
                    Dim obj As String = GetParam(funcParams, 0, "-1")
                    Dim oId As Integer = GetSecondChildIdByName(obj, mScript.mainClass(classId).ChildProperties)

                    If oId >= 0 Then
                        'удаляется один предмет
                        'если это контейнер, то выводим из него все предметы
                        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                        ReadPropertyInt(classId, "Type", oId, -1, objType, funcParams)
                        If objType = ObjectTypeEnum.CONTAINER Then
                            If ObjectsTakeOutFromContainer(oId, funcParams) = "#Error" Then Return "#Error"
                        End If

                        If ObjectRemoveElement(oId, funcParams) = "#Error" Then Return "#Error"
                    Else
                        'удаляются все предметы
                        Dim hDoc As HtmlDocument = frmPlayer.wbObjects.Document
                        If IsNothing(hDoc) = False Then
                            Dim hConvas As HtmlElement = hDoc.GetElementById("ObjectsConvas")
                            If IsNothing(hConvas) = False Then hConvas.InnerHtml = ""
                        End If
                    End If
                End If

                Return RemoveObject(classId, funcParams)
            Case "Add", "AddOnce"
                Dim oId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return _ERROR("Предмет " & funcParams(0) & " не найден!", functionName)

                If functionName = "AddOnce" Then
                    Dim WasObtained As Boolean = False
                    ReadPropertyBool(classId, "WasObtained", oId, -1, WasObtained, funcParams)
                    If WasObtained Then Return "-1"
                End If

                'получаем количество
                Dim isCountable As Boolean = False, cnt As Double = 1
                ReadPropertyBool(classId, "Countable", oId, -1, isCountable, funcParams)
                Dim strCnt As String = GetParam(funcParams, 1, "1")
                cnt = Val(strCnt)
                If cnt = 0 Then Return "-1"

                If Not isCountable Then
                    If cnt > 0 Then
                        cnt = 1
                    Else
                        cnt = -1
                    End If
                End If

                If cnt > 0 Then
                    'добавление предмета
                    Dim strContainer As String = "", contId As Integer = -1
                    'получение Id контейнера contId, в который будет произведена вставка
                    If funcParams.Count > 2 Then
                        strContainer = GetParam(funcParams, 2, "")
                        contId = GetSecondChildIdByName(strContainer, mScript.mainClass(classId).ChildProperties)
                        If contId < -1 Then contId = -1
                    Else
                        Dim sCont As String = ""
                        If ReadProperty(classId, "Container", oId, -1, sCont, funcParams) = False Then Return "#Error"
                        If sCont <> "''" Then contId = GetSecondChildIdByName(sCont, mScript.mainClass(classId).ChildProperties)
                    End If

                    'запускаем событие ObjectAddEvent
                    Dim arrs() As String = {oId.ToString, cnt.ToString(provider_points), contId.ToString}
                    'глобальное событие
                    Dim eventId As Integer = mScript.mainClass(classId).Properties("ObjectAddEvent").eventId
                    If eventId > 0 Then
                        Dim res As String = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectAddEvent", False)
                        If res = "#Error" Then
                            Return res
                        ElseIf res = "False" Then
                            Return "-1"
                        ElseIf IsNumeric(res.Replace(".", ",")) Then
                            cnt = Val(res) 'новое количество добавляемых предметов
                        End If
                    End If

                    'событие данного предмета
                    eventId = mScript.mainClass(classId).ChildProperties(oId)("ObjectAddEvent").eventId
                    If eventId > 0 Then
                        Dim res As String = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectAddEvent", False)
                        If res = "#Error" Then
                            Return res
                        ElseIf res = "False" Then
                            Return "-1"
                        ElseIf IsNumeric(res.Replace(".", ",")) Then
                            cnt = Val(res) 'новое количество добавляемых предметов
                        End If
                    End If

                    Dim wasExists As Boolean = False
                    ReadPropertyBool(classId, "BelongsToPlayer", oId, -1, wasExists, arrs)
                    If wasExists AndAlso isCountable Then
                        Dim oldCount As Integer = 0
                        ReadPropertyDbl(classId, "Count", oId, -1, oldCount, arrs)
                        cnt += oldCount
                    End If

                    SetPropertyValue(classId, "BelongsToPlayer", "True", oId)
                    SetPropertyValue(classId, "Count", cnt.ToString(provider_points), oId)
                    If wasExists = False Then
                        'отмечаем что предмет был получен
                        Dim res As String = PropertiesRouter(classId, "WasObtained", {oId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "True")
                        If res = "False" Then Return "-1"
                    End If

                    'устанавливаем контейнер
                    If funcParams.Count > 2 Then
                        SetPropertyValue(classId, "Container", strContainer, oId)
                    End If

                    If wasExists AndAlso funcParams.Count < 3 Then
                        ''0 - обычный, 1 - контейнер, 2 - описание, 3 - разделитель
                        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                        If ReadPropertyBool(classId, "Type", oId, -1, objType, funcParams) = False Then Return "#Error"
                        If objType = ObjectTypeEnum.DESCRIPTION Then
                            Dim res As String = ObjectUpdateDescription(oId, funcParams)
                            If res = "#Error" Then Return res
                        Else
                            If isCountable = False Then Return "-1" 'внешний вид предмета никак не изменяется - выход
                            'изменилось только количество
                            Dim res As String = ObjectChangeCount(oId, cnt.ToString, funcParams)
                            If res = "#Error" Then Return res
                        End If
                    Else
                        'добавляем на экран
                        If ObjectSetAppearance(oId, funcParams) = "#Error" Then Return "#Error"
                    End If

                    Return oId.ToString
                Else
                    'удаление предмета
                    cnt = cnt * -1
                    Dim arrs() As String = {oId.ToString, cnt.ToString(provider_points), "False"}
                    Dim wasExists As Boolean = False
                    ReadPropertyBool(classId, "BelongsToPlayer", oId, -1, wasExists, arrs)
                    If wasExists = False Then Return "-1"

                    'запускаем событие ObjectRemoveEvent
                    'глобальное событие
                    Dim eventId As Integer = mScript.mainClass(classId).Properties("ObjectRemoveEvent").eventId
                    If eventId > 0 Then
                        Dim res As String = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectRemoveEvent", False)
                        If res = "#Error" Then
                            Return res
                        ElseIf res = "False" Then
                            Return "-1"
                        ElseIf IsNumeric(res.Replace(".", ",")) Then
                            cnt = Val(res) 'новое количество удаляемых предметов
                        End If
                    End If

                    'событие данного предмета
                    eventId = mScript.mainClass(classId).ChildProperties(oId)("ObjectRemoveEvent").eventId
                    If eventId > 0 Then
                        Dim res As String = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectRemoveEvent", False)
                        If res = "#Error" Then
                            Return res
                        ElseIf res = "False" Then
                            Return "-1"
                        ElseIf IsNumeric(res.Replace(".", ",")) Then
                            cnt = Val(res) 'новое количество удаляемых предметов
                        End If
                    End If

                    'получение нового количества оставшихся преметов
                    Dim throwAway As Boolean = False
                    If isCountable Then
                        Dim oldCount As Integer = 0
                        ReadPropertyDbl(classId, "Count", oId, -1, oldCount, arrs)
                        cnt = oldCount - cnt

                        If cnt <= 0 Then
                            'удалять отрицательные?
                            Dim RemIfZero As Boolean = False
                            If ReadPropertyBool(classId, "RemIfZero", oId, -1, RemIfZero, arrs) = False Then Return "#Error"
                            If RemIfZero Then
                                'количество предметов стало отрицательным или равно 0, при этом отрицательное количество недопустимо
                                throwAway = True
                                cnt = 0
                            End If
                        End If
                    Else
                        throwAway = True
                        cnt = 0
                    End If

                    SetPropertyValue(classId, "Count", cnt.ToString(provider_points), oId)
                    If throwAway Then
                        'предмет удаляется

                        'если предмет - одетая аммуниция, то сначала снимаем
                        Dim isEquipment As Boolean = False, equipped As Boolean = False
                        ReadPropertyBool(classId, "IsEquipment", oId, -1, isEquipment, funcParams)
                        If isEquipment Then
                            ReadPropertyBool(classId, "Equipped", oId, -1, equipped, funcParams)
                            If equipped Then
                                If PropertiesRouter(classId, "Equipped", {oId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "False" Then Return "-1"
                            End If
                        End If

                        'если это контейнер, то выводим из него все предметы
                        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                        ReadPropertyInt(classId, "Type", oId, -1, objType, arrs)
                        If objType = ObjectTypeEnum.CONTAINER Then
                            If ObjectsTakeOutFromContainer(oId, arrs) = "#Error" Then Return "#Error"
                        End If
                        'предмет больше не принадлежит герою
                        SetPropertyValue(classId, "BelongsToPlayer", "False", oId)
                        'прячем предмет
                        Dim res As String = ObjectHide(oId, arrs)
                        If res = "#Error" Then Return res
                    Else
                        If funcParams.Count > 2 Then
                            'устанавливаем контейнер
                            Dim strContainer As String = GetParam(funcParams, 2, "")
                            SetPropertyValue(classId, "Container", strContainer, oId)
                            Dim res As String = ObjectSetAppearance(oId, funcParams)
                            If res = "#Error" Then Return res
                        Else
                            'изменяется количество предметов
                            Dim res As String = ObjectChangeCount(oId, cnt.ToString, funcParams)
                            If res = "#Error" Then Return res
                        End If
                    End If

                    Return oId.ToString
                End If
            Case "ThrowAway"
                Dim obj As String = GetParam(funcParams, 0, "-1")
                Dim oId As Integer = GetSecondChildIdByName(obj, mScript.mainClass(classId).ChildProperties)

                If oId >= 0 Then
                    'удаляем конкретный предмет
                    Dim cnt As Double = 0
                    If ReadPropertyDbl(classId, "Count", oId, -1, cnt, funcParams) = False Then Return "#Error"

                    Dim arrs() As String = {oId.ToString, cnt.ToString(provider_points), "False"}
                    Dim wasExists As Boolean = False
                    ReadPropertyBool(classId, "BelongsToPlayer", oId, -1, wasExists, arrs)
                    If wasExists = False Then Return "-1"

                    'запускаем событие ObjectRemoveEvent
                    'глобальное событие
                    Dim res As String = ""
                    Dim eventId As Integer = mScript.mainClass(classId).Properties("ObjectRemoveEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectRemoveEvent", False)
                        If res = "#Error" Then
                            Return res
                        ElseIf res = "False" Then
                            Return "-1"
                            'ElseIf IsNumeric(res.Replace(".", ",")) Then
                            '    cnt = Val(res) 'новое количество удаляемых предметов
                        End If
                    End If

                    'событие данного предмета
                    eventId = mScript.mainClass(classId).ChildProperties(oId)("ObjectRemoveEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectRemoveEvent", False)
                        If res = "#Error" Then
                            Return res
                        ElseIf res = "False" Then
                            Return "-1"
                        End If
                    End If

                    'если предмет - одетая аммуниция, то сначала снимаем
                    Dim isEquipment As Boolean = False, equipped As Boolean = False
                    ReadPropertyBool(classId, "IsEquipment", oId, -1, isEquipment, funcParams)
                    If isEquipment Then
                        ReadPropertyBool(classId, "Equipped", oId, -1, equipped, funcParams)
                        If equipped Then
                            If PropertiesRouter(classId, "Equipped", {oId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "False" Then Return "-1"
                        End If
                    End If

                    SetPropertyValue(classId, "Count", "0", oId)
                    'если это контейнер, то выводим из него все предметы
                    Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                    ReadPropertyInt(classId, "Type", oId, -1, objType, arrs)
                    If objType = ObjectTypeEnum.CONTAINER Then
                        If ObjectsTakeOutFromContainer(oId, arrs) = "#Error" Then Return "#Error"
                    End If
                    'предмет больше не принадлежит герою
                    SetPropertyValue(classId, "BelongsToPlayer", "False", oId)
                    'прячем предмет
                    res = ObjectHide(oId, arrs)
                    If res = "#Error" Then Return res

                    Return oId.ToString
                Else
                    'удаляем из рюкзака все предметы
                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Return "0"

                    Dim objLeft As Integer = 0
                    For i As Integer = mScript.mainClass(classId).ChildProperties.Count - 1 To 0 Step -1
                        Dim beToPlayer As Boolean = False
                        ReadPropertyBool(classId, "BelongsToPlayer", i, -1, beToPlayer, funcParams)
                        If Not beToPlayer Then Continue For

                        'получаем количество для параметра в ObjectRemoveEvent
                        Dim cnt As Double = 0
                        If ReadPropertyDbl(classId, "Count", i, -1, cnt, funcParams) = False Then Return "#Error"
                        Dim arrs() As String = {i.ToString, cnt.ToString(provider_points), "False"}

                        'запускаем событие ObjectRemoveEvent
                        'глобальное событие
                        Dim res As String = ""
                        Dim eventId As Integer = mScript.mainClass(classId).Properties("ObjectRemoveEvent").eventId
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectRemoveEvent", False)
                            If res = "#Error" Then
                                Return res
                            ElseIf res = "False" Then
                                objLeft += 1
                                Continue For
                            End If
                        End If

                        'событие данного предмета
                        eventId = mScript.mainClass(classId).ChildProperties(i)("ObjectRemoveEvent").eventId
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectRemoveEvent", False)
                            If res = "#Error" Then
                                Return res
                            ElseIf res = "False" Then
                                objLeft += 1
                                Continue For
                            End If
                        End If

                        'если предмет - одетая аммуниция, то сначала снимаем
                        Dim isEquipment As Boolean = False, equipped As Boolean = False
                        ReadPropertyBool(classId, "IsEquipment", i, -1, isEquipment, funcParams)
                        If isEquipment Then
                            ReadPropertyBool(classId, "Equipped", i, -1, equipped, funcParams)
                            If equipped Then
                                If PropertiesRouter(classId, "Equipped", {i.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "False" Then
                                    objLeft += 1
                                    Continue For
                                End If
                            End If
                        End If

                        'количество = 0
                        SetPropertyValue(classId, "Count", "0", i)
                        'предмет больше не принадлежит герою
                        SetPropertyValue(classId, "BelongsToPlayer", "False", i)
                        'прячем предмет
                        res = ObjectHide(i, arrs)
                        If res = "#Error" Then Return res
                    Next i
                    Return objLeft.ToString
                End If
            Case "CountInContainer"
                Dim oId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return _ERROR("Предмет " & funcParams(0) & " не найден.", functionName)

                Dim lst As List(Of Integer) = ObjectsGetContainerContent(oId, funcParams)
                Return lst.Count.ToString
                'mScript.lastArray = New cVariable.variableEditorInfoType
            Case "GetContainerContent"
                Dim oId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return _ERROR("Предмет " & funcParams(0) & " не найден.", functionName)

                Dim lst As List(Of Integer) = ObjectsGetContainerContent(oId, funcParams)
                mScript.lastArray = New cVariable.variableEditorInfoType
                If lst.Count = 0 Then
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = "-1"
                Else
                    Dim arr() As String
                    ReDim arr(lst.Count - 1)
                    For i As Integer = 0 To lst.Count - 1
                        arr(i) = lst(i).ToString
                    Next
                    mScript.lastArray.arrValues = arr
                End If

                Return "#ARRAY"
            Case "ReleaseContainer"
                Dim oId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return _ERROR("Предмет " & funcParams(0) & " не найден.", functionName)
                Dim destObj As String = GetParam(funcParams, 1, "''")
                Return ObjectsTakeOutFromContainer(oId, funcParams, destObj)
            Case "IsContainerOpened"
                Dim oId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return _ERROR("Предмет " & funcParams(0) & " не найден.", functionName)
                Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
                If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".")
                If hOuter.Children.Count < 2 Then Return "False"
                hOuter = hOuter.Children(1)
                If HTMLHasClass(hOuter, "InnerContainer") = False Then Return "False"
                If String.IsNullOrEmpty(hOuter.Style) = False AndAlso hOuter.Style.StartsWith("display: none") Then Return "False"
                If HTMLHasClass(hOuter, "Collapsed") Then Return "False"
                Return "True"
            Case "OpenContainer"
                Dim obj As String = GetParam(funcParams, 0, "-1")
                Dim oId As Integer = GetSecondChildIdByName(obj, mScript.mainClass(classId).ChildProperties)

                If oId < 0 Then
                    'открываем все контейнеры
                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Return "False"
                    Dim wasOpened As Boolean = False
                    For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                        ReadPropertyInt(classId, "Type", i, -1, objType, funcParams)
                        If objType <> ObjectTypeEnum.CONTAINER Then Continue For

                        'открываем контейнер
                        Dim hOuter As HtmlElement = GetObjectOuterElement(i)
                        If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".")
                        If hOuter.Children.Count < 2 Then Continue For
                        hOuter = hOuter.Children(1)
                        If HTMLHasClass(hOuter, "InnerContainer") = False Then Continue For
                        If String.IsNullOrEmpty(hOuter.Style) = False AndAlso hOuter.Style.StartsWith("display: none") Then Continue For

                        If HTMLRemoveClass(hOuter, "Collapsed") Then wasOpened = True
                    Next i
                    Return wasOpened.ToString
                Else
                    'открываем контейнер
                    Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
                    If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".")
                    If hOuter.Children.Count < 2 Then Return "False"
                    hOuter = hOuter.Children(1)
                    If HTMLHasClass(hOuter, "InnerContainer") = False Then Return "False"
                    If String.IsNullOrEmpty(hOuter.Style) = False AndAlso hOuter.Style.StartsWith("display: none") Then Return "False"
                    HTMLRemoveClass(hOuter, "Collapsed")
                    Return "True"
                End If
            Case "CloseContainer"
                Dim obj As String = GetParam(funcParams, 0, "-1")
                Dim oId As Integer = GetSecondChildIdByName(obj, mScript.mainClass(classId).ChildProperties)

                If oId < 0 Then
                    'закрываем все контейнеры
                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Return "False"
                    Dim wasClosed As Boolean = False
                    For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                        ReadPropertyInt(classId, "Type", i, -1, objType, funcParams)
                        If objType <> ObjectTypeEnum.CONTAINER Then Continue For

                        'закрываем контейнер
                        Dim hOuter As HtmlElement = GetObjectOuterElement(i)
                        If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".")
                        If hOuter.Children.Count < 2 Then Continue For
                        hOuter = hOuter.Children(1)
                        If HTMLHasClass(hOuter, "InnerContainer") = False Then Continue For
                        If String.IsNullOrEmpty(hOuter.Style) = False AndAlso hOuter.Style.StartsWith("display: none") Then Continue For
                        Dim bln As Boolean = HTMLAddClass(hOuter, "Collapsed")
                        If bln Then wasClosed = True
                    Next i
                    Return wasClosed.ToString
                Else
                    'закрываем контейнер
                    Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
                    If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".")
                    If hOuter.Children.Count < 2 Then Return "False"
                    hOuter = hOuter.Children(1)
                    If HTMLHasClass(hOuter, "InnerContainer") = False Then Return "False"
                    If String.IsNullOrEmpty(hOuter.Style) = False AndAlso hOuter.Style.StartsWith("display: none") Then Return "False"
                    Return HTMLAddClass(hOuter, "Collapsed").ToString
                End If
            Case "Select"
                Dim obj As String = GetParam(funcParams, 0, "-1")
                Dim oId As Integer = GetSecondChildIdByName(obj, mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return _ERROR("Предмет " & funcParams(0) & " не найден.", functionName)

                Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
                If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".")
                If hOuter.Children.Count = 0 Then Return ""

                Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                ReadPropertyInt(classId, "Type", oId, -1, objType, funcParams)
                If objType = ObjectTypeEnum.SEPARATOR Then Return ""

                EventGeneratedFromScript = True
                Call del_object_Click(hOuter.Children(0), Nothing)
                Return ""
            Case "GetEquipByType", "EquipCount"
                mScript.lastArray = New cVariable.variableEditorInfoType
                If IsNothing(mScript.mainClass(classId)) Then
                    If functionName = "EquipCount" Then Return "0"
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = "-1"
                    Return "#ARRAY"
                End If

                Dim seekType As Integer = Val(UnWrapString(funcParams(0)))
                Dim shouldBeEquipped As Boolean = CBool(GetParam(funcParams, 1, "False"))

                Dim lst As List(Of String) = EquipmentGetByType(seekType, shouldBeEquipped)
                If IsNothing(lst) Then Return "#Error"

                If functionName = "EquipCount" Then Return lst.Count
                If lst.Count = 0 Then
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = "-1"
                Else
                    mScript.lastArray.arrValues = lst.ToArray
                End If

                Return "#ARRAY"
            Case "PutOn"
                Dim oId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return _ERROR("Предмет " & funcParams(0) & " не найден.", functionName)
                Dim isEquipment As Boolean = False
                If ReadPropertyBool(classId, "IsEquipment", oId, -1, isEquipment, funcParams) = False Then Return "#Error"
                If Not isEquipment Then Return _ERROR("Предмет " & funcParams(0) & " не являет аммуницией.", functionName)
                If mScript.mainClass(classId).ChildProperties(oId)("Equipped").Value = "True" Then Return "True" 'предмет уже одет

                Dim maxEquipped As Integer = Val(GetParam(funcParams, 1, "-1"))
                Dim remObj As String = GetParam(funcParams, 2, "")
                Dim remId As Integer = -1
                If String.IsNullOrEmpty(remObj) = False Then
                    remId = GetSecondChildIdByName(remObj, mScript.mainClass(classId).ChildProperties)
                End If

                '1) Снимаем экипировку указанного типа (если надо), чтобы освободить место для новой
                Dim eventId As Integer = 0
                Dim res As String = ""
                Dim eqLeft As Integer = 0
                If remId < 0 AndAlso maxEquipped > -1 Then
                    'количество ограничено, замена не выбрана.
                    Dim seekType As Integer = 0
                    If ReadPropertyInt(classId, "EquipType", oId, -1, seekType, funcParams) = False Then Return "#Error"
                    Dim eqLst As List(Of String) = EquipmentGetByType(seekType, True)
                    If IsNothing(eqLst) Then Return "#Error"

                    eqLeft = eqLst.Count 'осталось одетых преметов данного вида
                    If eqLst.Count > 0 Then
                        'одет минимум один предмет указанного типа
                        Dim curId As Integer = eqLeft
                        Do While eqLeft >= maxEquipped
                            'выполняем до тех пор, пока количесто одетых предметов данного типа не станет меньше на 1 от максимума 
                            '(или не снимем все, если maxEquipped = 0; или не перепробуем снять все предметы, но так и не достигнем maxEquipped - 1)
                            curId -= 1
                            If curId < 0 Then Exit Do
                            'Dim curId As Integer = CInt(eqLst(eqLeft - 1))

                            res = PropertiesRouter(classId, "Equipped", {eqLst(curId)}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False")
                            If res = "#Error" Then
                                Return res
                            ElseIf res <> "False" Then
                                eqLeft -= 1
                            End If
                        Loop
                    End If
                ElseIf remId >= 0 Then
                    'указан предмет, который надо снять

                    res = PropertiesRouter(classId, "Equipped", {remId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False")
                    If res = "#Error" Then
                        Return res
                    ElseIf res = "False" Then
                        'снять предмет не удалось. Значит, и одеть новый не получится
                        Return "False"
                    End If
                    eqLeft -= 1
                    Dim hOuter As HtmlElement = GetObjectOuterElement(remId)
                    If IsNothing(hOuter) = False Then hOuter = hOuter.Children(0)
                    If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(2) & ".", functionName)
                    If EquipmentSetClass(remId, funcParams, hOuter, False) = False Then Return "#Error"
                End If

                If eqLeft >= maxEquipped AndAlso maxEquipped > -1 Then
                    'не удалось снять достаточно, чтобы одеть новый предмет. Одевания не происходит
                    Return "False"
                End If

                '2) Одеваем новую экипировку
                'Событие одевания ObjectPutOnEvent

                'Глобальное
                res = ""
                eventId = mScript.mainClass(classId).Properties("ObjectPutOnEvent").eventId
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, {oId.ToString}, "ObjectPutOnEvent", False)
                    If res = "#Error" Then Return res
                End If

                'данного предмета
                If res <> "False" Then
                    eventId = mScript.mainClass(classId).ChildProperties(oId)("ObjectPutOnEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, {oId.ToString}, "ObjectPutOnEvent", False)
                        If res = "#Error" Then Return res
                    End If
                End If

                If res = "False" Then Return "False"

                'события пройдены, экипировка должна быть надета
                Dim trackResult As String = mScript.trackingProperties.RunBefore(classId, "Equipped", {oId.ToString}, "True")
                If trackResult = "#Error" Then Return trackResult

                If trackResult <> "False" Then
                    SetPropertyValue(classId, "Equipped", "True", oId)

                    Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
                    If IsNothing(hOuter) = False Then hOuter = hOuter.Children(0)
                    If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".", functionName)
                    If EquipmentSetClass(oId, funcParams, hOuter, True) = False Then Return "#Error"
                    Return "True"
                Else
                    Return "False"
                End If
            Case "PutOff"
                Return PropertiesRouter(classId, "Equipped", funcParams, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False")
            Case "ShowDescription"
                Dim oId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If oId < 0 Then Return _ERROR("Предмет " & funcParams(0) & " не найден.", functionName)

                'получаем тип пердмета
                Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                If ReadPropertyInt(classId, "Type", oId, -1, objType, funcParams) = False Then Return "#Error"
                If objType = ObjectTypeEnum.SEPARATOR Then Return ""
                'получаем описание
                Dim selector As Integer, strDesc As String = "", retFormat As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL
                selector = Val(GetParam(funcParams, 1, "-1"))
                If ReadProperty(classId, "Description", oId, -1, strDesc, funcParams, retFormat, selector) = False Then Return "#Error"
                If retFormat = MatewScript.ReturnFormatEnum.TO_STRING Then strDesc = UnWrapString(strDesc)
                'выводим описание
                If objType = ObjectTypeEnum.DESCRIPTION Then
                    Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
                    If IsNothing(hOuter) = False Then
                        If hOuter.Children.Count > 0 Then
                            hOuter = hOuter.Children(0)
                        Else
                            Return ""
                        End If

                    End If
                    If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".", functionName)
                    'вывод описания
                    hOuter.InnerHtml = strDesc
                Else
                    'получаем html-элемент для вывода описания
                    Dim hDoc As HtmlDocument = frmPlayer.wbDescription.Document
                    If IsNothing(hDoc) Then Return _ERROR("Документ окна описаний не загружен.", functionName)
                    Dim hDesc As HtmlElement = hDoc.GetElementById("DescriptionConvas")
                    If IsNothing(hDesc) Then _ERROR("Нарушена структура документа description.html. Элемент DescriptionConvas не найден.", functionName)
                    'вывод описания
                    hDesc.InnerHtml = strDesc
                End If

                'запуск изменения отображение количества предметов
                Dim oCnt As Double = 1
                If ReadPropertyDbl(classId, "Count", oId, -1, oCnt, funcParams) = False Then Return "#Error"
                ObjectChangeCount(oId, oCnt.ToString.Replace(","c, "."), funcParams)

                Return ""
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    ''' <summary>
    ''' Получает список Id аммуниции указанного типа
    ''' </summary>
    ''' <param name="seekType">тип искомой аммуниции</param>
    ''' <param name="shouldBeEquipped">искать только среди одетых</param>
    Private Function EquipmentGetByType(ByVal seekType As Integer, ByVal shouldBeEquipped As Boolean) As List(Of String)
        Dim lst As New List(Of String)
        Dim classId As Integer = mScript.mainClassHash("O")
        For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
            Dim strI As String = i.ToString
            'является аммуницией?
            Dim isEquipment As Boolean = False
            If ReadPropertyBool(classId, "IsEquipment", i, -1, isEquipment, {strI}) = False Then Return Nothing
            If Not isEquipment Then Continue For

            'тип совпадает?
            Dim equipType As Integer = 0
            If ReadPropertyInt(classId, "EquipType", i, -1, equipType, {strI}) = False Then Return Nothing
            If equipType <> seekType Then Continue For

            'одето?
            If shouldBeEquipped Then
                Dim equipped As Boolean = False
                If ReadPropertyBool(classId, "Equipped", i, -1, equipped, {strI}) = False Then Return Nothing
                If Not equipped Then Continue For
            End If

            lst.Add(strI)
        Next

        Return lst
    End Function

    ''' <summary>
    ''' Прячет указанный предмет (например, при его удалении)
    ''' </summary>
    ''' <param name="oId">Id предмета</param>
    ''' <param name="arrParams"></param>
    ''' <returns></returns>
    Private Function ObjectHide(ByVal oId As Integer, ByRef arrParams() As String) As String
        Dim outerDiv As HtmlElement = GetObjectOuterElement(oId)
        If IsNothing(outerDiv) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет.")

        outerDiv.Style = "display:none"
        Return ""
    End Function

    ''' <summary>
    ''' Обновляет описание предмета внутри окна предметов (для типа "Описание")
    ''' </summary>
    ''' <param name="oId">Id премета</param>
    ''' <param name="arrParams"></param>
    Private Function ObjectUpdateDescription(ByVal oId As Integer, ByRef arrParams() As String) As String
        Dim classId As Integer = mScript.mainClassHash("O")

        '0 - обычный, 1 - контейнер, 2 - описание, 3 - разделитель
        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
        If ReadPropertyInt(classId, "Type", oId, -1, objType, arrParams) = False Then Return "#Error"
        If objType <> ObjectTypeEnum.DESCRIPTION Then Return ""

        Dim newDesc As String = ""
        ReadProperty(classId, "Description", oId, -1, newDesc, arrParams)
        newDesc = UnWrapString(newDesc)

        Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
        If IsNothing(hOuter) OrElse hOuter.Children.Count = 0 Then
            Dim objName As String = mScript.mainClass(classId).ChildProperties(oId)("Name").Value
            Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & objName & ".")
        End If
        hOuter = hOuter.Children(0)
        hOuter.InnerHtml = newDesc
        Return ""
    End Function

    ''' <summary>
    ''' Меняет подпись предмета в окне предметов
    ''' </summary>
    ''' <param name="oId">Id предмета</param>
    ''' <param name="newValue">новая подпись</param>
    ''' <param name="arrParams"></param>
    Private Function ObjectChangeCaption(ByVal oId As Integer, ByVal newValue As String, ByRef arrParams() As String) As String
        Dim classId As Integer = mScript.mainClassHash("O")
        newValue = UnWrapString(newValue)

        'получаем стиль отображения в окне
        Dim classOW As Integer = mScript.mainClassHash("OW")
        Dim showStyle As ObjectWindowShowStyle = ObjectWindowShowStyle.CAPTION_ONLY
        If ReadPropertyInt(classOW, "Style", -1, -1, showStyle, arrParams) = False Then Return "#Error"
        'If showStyle <> ObjectWindowShowStyle.CAPTION_IMAGE_COUNT AndAlso showStyle <> ObjectWindowShowStyle.IMAGE_COUNT Then Return "" 'только 2 - название, картинка и количество;3 - только картинка и количество

        '0 - обычный, 1 - контейнер, 2 - описание, 3 - разделитель
        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
        If ReadPropertyBool(classId, "Type", oId, -1, objType, arrParams) = False Then Return "#Error"
        If objType = ObjectTypeEnum.DESCRIPTION OrElse objType = ObjectTypeEnum.SEPARATOR Then Return ""

        'получаем ContainedStyle контейнера
        Dim containedStyle As ObjectContainedStyleEnum = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE
        Dim classO As Integer = mScript.mainClassHash("O")
        Dim objCont As String = ""
        ReadProperty(classO, "Container", oId, -1, objCont, arrParams)
        If String.IsNullOrEmpty(objCont) = False Then
            Dim parId As Integer = GetSecondChildIdByName(objCont, mScript.mainClass(classO).ChildProperties)
            If parId >= 0 Then
                ReadPropertyInt(classO, "ContainedStyle", parId, -1, containedStyle, arrParams)
            End If
        End If

        Dim outerDiv As HtmlElement = GetObjectOuterElement(oId)
        If IsNothing(outerDiv) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & mScript.mainClass(classO).ChildProperties(oId)("Name").Value & ".")

        Dim hObjText As HtmlElement = Nothing
        If containedStyle = ObjectContainedStyleEnum.USUAL_AS_OUTSIDE Then
            If showStyle = ObjectWindowShowStyle.CAPTION_ONLY Then
                hObjText = outerDiv.Children(0)
            Else
                If showStyle = ObjectWindowShowStyle.IMAGE_COUNT Then
                    hObjText = outerDiv.Children(0).Children(0).Children(0).Children(0).Children(0) 'TABLE-TBODY-TR-TD-IMG
                    hObjText.SetAttribute("Title", newValue)
                    Return ""
                End If

                hObjText = outerDiv.Children(0).Children(0).Children(0).Children(1) 'TABLE-TBODY-TR-TD(1)
            End If
        ElseIf containedStyle = ObjectContainedStyleEnum.IMAGE_ONLY Then
            If showStyle = ObjectWindowShowStyle.CAPTION_ONLY Then
                hObjText = outerDiv.Children(0)
            Else
                hObjText = outerDiv.Children(0) 'IMG
                hObjText.SetAttribute("Title", newValue)
                Return ""
            End If
        Else
            If showStyle = ObjectWindowShowStyle.CAPTION_ONLY Then
                hObjText = outerDiv.Children(0)
            Else
                hObjText = outerDiv.Children(0).Children(0).Children(0) 'SPAN-NOBR-IMG
                hObjText.SetAttribute("Title", newValue)
                Return ""
            End If
        End If

        If IsNothing(hObjText) = False Then
            hObjText.InnerHtml = newValue
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Изменяет количество предметов в окне предметов
    ''' </summary>
    ''' <param name="oId">Id предмета</param>
    ''' <param name="newValue">новое количество</param>
    ''' <param name="arrParams"></param>
    Private Function ObjectChangeCount(ByVal oId As Integer, ByVal newValue As String, ByRef arrParams() As String) As String
        Dim isCountable As Boolean = False
        Dim classId As Integer = mScript.mainClassHash("O")
        If ReadPropertyBool(classId, "Countable", oId, -1, isCountable, arrParams) = False Then Return "#Error"
        If isCountable = False Then Return ""

        'получаем имя предмета
        Dim objName As String = mScript.mainClass(classId).ChildProperties(oId)("Name").Value

        'получаем стиль отображения в окне
        Dim classOW As Integer = mScript.mainClassHash("OW")
        Dim showStyle As ObjectWindowShowStyle = ObjectWindowShowStyle.CAPTION_ONLY
        If ReadPropertyInt(classOW, "Style", -1, -1, showStyle, arrParams) = False Then Return "#Error"
        If showStyle <> ObjectWindowShowStyle.CAPTION_IMAGE_COUNT AndAlso showStyle <> ObjectWindowShowStyle.IMAGE_COUNT Then Return "" 'только 2 - название, картинка и количество;3 - только картинка и количество

        '0 - обычный, 1 - контейнер, 2 - описание, 3 - разделитель
        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
        If ReadPropertyBool(classId, "Type", oId, -1, objType, arrParams) = False Then Return "#Error"
        If objType <> ObjectTypeEnum.USUAL AndAlso objType <> ObjectTypeEnum.CONTAINER Then Return ""

        'получаем формат вывода количества
        Dim countFormat As String = ""
        ReadProperty(classId, "CountFormat", oId, -1, countFormat, arrParams)
        countFormat = UnWrapString(countFormat)

        If String.IsNullOrEmpty(countFormat) = False Then
            Dim dblCnt As Double = 0
            If Double.TryParse(newValue, System.Globalization.NumberStyles.Float, provider_points, dblCnt) Then
                newValue = Format(dblCnt, countFormat)
            End If
        End If

        Dim outerDiv As HtmlElement = GetObjectOuterElement(oId)
        If IsNothing(outerDiv) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & objName & ".")
        'TABLE-TBODY-TR-TD(1 or 2)
        'SPAN-NOBR-SPAN(1)


        Dim hEl As HtmlElement = outerDiv.Children(0).Children(0) 'TABLE-TBODY / SPAN-NOBR
        If hEl.TagName = "NOBR" Then
            hEl = hEl.Children(1)
            If hEl.GetAttribute("ClassName") <> "objectCount" Then hEl = Nothing
        Else
            hEl = hEl.Children(0) 'TR
            If hEl.Children.Count > 1 AndAlso hEl.Children(1).GetAttribute("ClassName") = "objectCount" Then
                hEl = hEl.Children(1)
            ElseIf hEl.Children.Count > 2 AndAlso hEl.Children(2).GetAttribute("ClassName") = "objectCount" Then
                hEl = hEl.Children(2)
            End If
        End If

        If IsNothing(hEl) = False Then hEl.InnerHtml = newValue
        Return ""
    End Function

    ''' <summary>
    ''' Создает список из Id элементов, входящий в данный контейнер
    ''' </summary>
    ''' <param name="oId">Id контейнера</param>
    ''' <param name="arrParams"></param>
    ''' <returns></returns>
    Private Function ObjectsGetContainerContent(ByVal oId As Integer, ByRef arrParams() As String) As List(Of Integer)
        Dim classId As Integer = mScript.mainClassHash("O")
        Dim lst As New List(Of Integer)

        'получаем имя и Id контейнера
        Dim seekName As String = mScript.mainClass(classId).ChildProperties(oId)("Name").Value
        'ReadProperty(classId, "Container", oId, -1, seekName, arrParams)
        'Dim seekId As Integer = GetSecondChildIdByName(seekName, mScript.mainClass(classId).ChildProperties)
        'If seekId < 0 Then Return lst

        Dim curName As String = ""
        For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
            'перебираем все предметы
            ReadProperty(classId, "Container", i, -1, curName, arrParams)
            If String.Compare(seekName, curName, True) = 0 OrElse (IsNumeric(curName) AndAlso Val(curName) = oId) Then
                'имя/Id контейнера совпадают
                lst.Add(i)
            End If
        Next i

        Return lst
    End Function

    ''' <summary>
    ''' Вынимает предметы из контейнера
    ''' </summary>
    ''' <param name="oId">Id контейнера</param>
    ''' <param name="arrParams"></param>
    Private Function ObjectsTakeOutFromContainer(ByVal oId As Integer, ByRef arrParams() As String, Optional ByVal destObj As String = "''") As String
        'получаем список предметов внутри контейнера
        Dim lst As List(Of Integer) = ObjectsGetContainerContent(oId, arrParams)
        If lst.Count = 0 Then Return "0"

        Dim classId As Integer = mScript.mainClassHash("O")
        For i As Integer = 0 To lst.Count - 1
            Dim curId As Integer = lst(i)
            Dim res As String = PropertiesRouter(classId, "Container", {curId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, destObj)
            If res = "#Error" Then Return res
        Next

        Return lst.Count.ToString
    End Function

    ''' <summary>
    ''' Изменяет стиль отображения предметов внутри контейнера в окне предметов
    ''' </summary>
    ''' <param name="oId">Id контейнера</param>
    ''' <param name="arrParams"></param>
    Private Function ObjectChangeInnerContainerAppearance(ByVal oId As Integer, ByRef arrParams() As String) As String
        'получаем список предметов внутри контейнера
        Dim lst As List(Of Integer) = ObjectsGetContainerContent(oId, arrParams)
        If lst.Count = 0 Then Return ""

        For i As Integer = 0 To lst.Count - 1
            'изменяем внешний вид всех предметов внутри контейнера
            Dim res As String = ObjectSetAppearance(lst(i), arrParams)
            If res = "#Error" Then Return res
        Next

        Return ""
    End Function

    Private Function ClassO_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        Dim oId As Integer = -1
        If funcParams.Count > 0 Then
            If funcParams(0) = "-1" Then Return ""
            oId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If oId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует имени и Id предмета {0}!", funcParams(0))
                Return "#Error"
            End If
        Else
            Return "" 'значения по умолчанию
        End If
        'получаем имя предмета
        Dim objName As String = mScript.mainClass(classId).ChildProperties(oId)("Name").Value

        Select Case propertyName
            Case "Count"
                'получаем количество
                Dim isCountable As Boolean = False, cnt As Double = 1
                ReadPropertyBool(classId, "Countable", oId, -1, isCountable, funcParams)
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                cnt = Val(nValue)
                Dim prevVal As Double = Val(mScript.PrepareStringToPrint(PREV_VALUE, funcParams))
                If cnt = prevVal Then Return ""

                If Not isCountable Then
                    If cnt > 0 Then
                        cnt = 1
                    Else
                        cnt = -1
                    End If
                Else
                    cnt = cnt - prevVal
                End If

                If Val(nValue) > prevVal Then
                    'добавление предмета
                    'запускаем событие ObjectAddEvent

                    Dim cont As String = "", contId As Integer = -1
                    If ReadProperty(classId, "Container", oId, -1, cont, {oId.ToString}) = False Then Return "#Error"
                    If cont <> "''" Then contId = GetSecondChildIdByName(cont, mScript.mainClass(classId).ChildProperties)
                    If contId < -1 Then contId = -1

                    Dim arrs() As String = {oId.ToString, cnt.ToString(provider_points), contId.ToString}
                    'глобальное событие
                    Dim eventId As Integer = mScript.mainClass(classId).Properties("ObjectAddEvent").eventId
                    If eventId > 0 Then
                        Dim res As String = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectAddEvent", False)
                        If res = "#Error" Then
                            SetPropertyValue(classId, "Count", PREV_VALUE, oId)
                            Return res
                        ElseIf res = "False" Then
                            SetPropertyValue(classId, "Count", PREV_VALUE, oId)
                            Return "-1"
                        ElseIf IsNumeric(res.Replace(".", ",")) Then
                            cnt = Val(res) 'новое количество добавляемых предметов
                        End If
                    End If

                    'событие данного предмета
                    eventId = mScript.mainClass(classId).ChildProperties(oId)("ObjectAddEvent").eventId
                    If eventId > 0 Then
                        Dim res As String = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectAddEvent", False)
                        If res = "#Error" Then
                            SetPropertyValue(classId, "Count", PREV_VALUE, oId)
                            Return res
                        ElseIf res = "False" Then
                            SetPropertyValue(classId, "Count", PREV_VALUE, oId)
                            Return "-1"
                        ElseIf IsNumeric(res.Replace(".", ",")) Then
                            cnt = Val(res) 'новое количество добавляемых предметов
                        End If
                    End If

                    Dim wasExists As Boolean = False
                    ReadPropertyBool(classId, "BelongsToPlayer", oId, -1, wasExists, arrs)
                    If wasExists AndAlso isCountable Then
                        cnt += prevVal
                    End If

                    SetPropertyValue(classId, "BelongsToPlayer", "True", oId)
                    If wasExists = False Then
                        'отмечаем что предмет был получен
                        Dim res As String = PropertiesRouter(classId, "WasObtained", {oId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "True")
                        If res = "False" Then Return "-1"
                    End If

                    If wasExists Then
                        ''0 - обычный, 1 - контейнер, 2 - описание, 3 - разделитель
                        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                        If ReadPropertyBool(classId, "Type", oId, -1, objType, funcParams) = False Then Return "#Error"
                        If isCountable = False Then Return "" 'внешний вид предмета никак не изменяется - выход 
                        'изменилось только количество
                        Dim res As String = ObjectChangeCount(oId, cnt.ToString, funcParams)
                        If res = "#Error" Then Return res
                    Else
                        ObjectSetAppearance(oId, funcParams)
                    End If
                Else
                    'удаление предмета
                    cnt = cnt * -1
                    Dim arrs() As String = {oId.ToString, cnt.ToString(provider_points), "False"}
                    Dim wasExists As Boolean = False
                    ReadPropertyBool(classId, "BelongsToPlayer", oId, -1, wasExists, arrs)
                    If wasExists = False Then
                        SetPropertyValue(classId, "Count", "0", oId, -1)
                        Return ""
                    End If

                    'запускаем событие ObjectRemoveEvent
                    'глобальное событие
                    Dim eventId As Integer = mScript.mainClass(classId).Properties("ObjectRemoveEvent").eventId
                    If eventId > 0 Then
                        Dim res As String = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectRemoveEvent", False)
                        If res = "#Error" Then
                            SetPropertyValue(classId, "Count", PREV_VALUE, oId)
                            Return res
                        ElseIf res = "False" Then
                            SetPropertyValue(classId, "Count", PREV_VALUE, oId)
                            Return ""
                        ElseIf IsNumeric(res.Replace(".", ",")) Then
                            cnt = Val(res) 'новое количество удаляемых предметов
                        End If
                    End If

                    'событие данного предмета
                    eventId = mScript.mainClass(classId).ChildProperties(oId)("ObjectRemoveEvent").eventId
                    If eventId > 0 Then
                        Dim res As String = mScript.eventRouter.RunEvent(eventId, arrs, "ObjectRemoveEvent", False)
                        If res = "#Error" Then
                            SetPropertyValue(classId, "Count", PREV_VALUE, oId)
                            Return res
                        ElseIf res = "False" Then
                            SetPropertyValue(classId, "Count", PREV_VALUE, oId)
                            Return "-1"
                        ElseIf IsNumeric(res.Replace(".", ",")) Then
                            cnt = Val(res) 'новое количество удаляемых предметов
                        End If
                    End If

                    'получение нового количества оставшихся преметов
                    Dim throwAway As Boolean = False
                    If isCountable Then
                        cnt = prevVal - cnt

                        If cnt <= 0 Then
                            'удалять отрицательные?
                            Dim RemIfZero As Boolean = False
                            If ReadPropertyBool(classId, "RemIfZero", oId, -1, RemIfZero, arrs) = False Then Return "#Error"
                            If RemIfZero Then
                                'количество предметов стало отрицательным или равно 0, при этом отрицательное количество недопустимо
                                throwAway = True
                                cnt = 0
                            End If
                        End If
                    Else
                        throwAway = True
                        cnt = 0
                    End If

                    If throwAway Then
                        'предмет удаляется

                        'если предмет - одетая аммуниция, то сначала снимаем
                        Dim isEquipment As Boolean = False, equipped As Boolean = False
                        ReadPropertyBool(classId, "IsEquipment", oId, -1, isEquipment, funcParams)
                        If isEquipment Then
                            ReadPropertyBool(classId, "Equipped", oId, -1, equipped, funcParams)
                            If equipped Then
                                If PropertiesRouter(classId, "Equipped", {oId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "False") = "False" Then Return "-1"
                            End If
                        End If

                        'если это контейнер, то выводим из него все предметы
                        Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                        ReadPropertyInt(classId, "Type", oId, -1, objType, arrs)
                        If objType = ObjectTypeEnum.CONTAINER Then
                            If ObjectsTakeOutFromContainer(oId, arrs) = "#Error" Then Return "#Error"
                        End If
                        'предмет больше не принадлежит герою
                        SetPropertyValue(classId, "BelongsToPlayer", "False", oId)
                        'прячем предмет
                        Dim res As String = ObjectHide(oId, arrs)
                        If res = "#Error" Then Return res
                    Else
                        'изменяется количество предметов 
                        Dim res As String = ObjectChangeCount(oId, cnt.ToString, funcParams)
                        If res = "#Error" Then Return res
                    End If
                End If
                Return ""

            Case "Container", "Countable", "CountFormat", "Enabled", "Visible", "PicWidth", "PicHeight", "BelongsToPlayer"
                Return ObjectSetAppearance(oId, funcParams)
            Case "ContainedStyle"
                Return ObjectChangeInnerContainerAppearance(oId, funcParams)
            Case "Caption"
                Return ObjectChangeCaption(oId, newValue, funcParams)
            Case "Description"
                Return ObjectUpdateDescription(oId, arrParams)
            Case "Type"
                Dim oldType As ObjectTypeEnum = Val(UnWrapString(PREV_VALUE))
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                Dim newType As ObjectTypeEnum = Val(nValue)
                If oldType = newType Then Return ""

                Dim res As String
                If oldType = ObjectTypeEnum.CONTAINER Then
                    res = ObjectsTakeOutFromContainer(oId, funcParams)
                    If res = "#Error" Then Return res
                End If
                Return ObjectSetAppearance(oId, funcParams)
            Case "Picture"
                'получаем тип. '0 - обычный, 1 - контернер, 2 - описание, 3 - разделитель
                Dim objType As ObjectTypeEnum = ObjectTypeEnum.USUAL
                ReadPropertyInt(classId, "Type", oId, -1, objType, funcParams)
                If objType = ObjectTypeEnum.SEPARATOR Then Return "" 'для типа сеператор не подходит

                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                Dim objPicture As String = nValue
                objPicture = objPicture.Replace("\", "/")

                Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, objPicture)
                If FileIO.FileSystem.FileExists(fPath) = False Then
                    Return _ERROR("Файл " & fPath & " не найден!")
                End If

                'собственно замена картинки на новую
                Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
                'If IsNothing(hOuter) = False Then hOuter = hOuter.Parent
                If IsNothing(hOuter) = False Then hOuter = hOuter.Children(0)
                If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & objName & ".")

                If objType = ObjectTypeEnum.DESCRIPTION Then
                    'type = описание. Находим первую же картинку и работаем с ней
                    'description - DIV/SPAN-...(IMG)...

                    Dim hImgCol As HtmlElementCollection = hOuter.GetElementsByTagName("IMG")
                    If hImgCol.Count > 0 Then
                        hImgCol(0).SetAttribute("src", objPicture)
                    End If
                Else
                    '0 - обычный, 1 - контейнер
                    'возможное расположение картинки: 
                    '1) IMG 2) SPAN-NOBR-IMG 3) TABLE-TBODY-TR-TD
                    If hOuter.TagName <> "IMG" Then
                        Do
                            If hOuter.Children.Count = 0 Then
                                hOuter = Nothing
                                Exit Do
                            End If
                            hOuter = hOuter.Children(0)
                            If hOuter.TagName = "IMG" Then Exit Do
                        Loop

                        If IsNothing(hOuter) = False Then
                            hOuter.SetAttribute("Src", objPicture)
                        End If
                    End If
                End If
                Return ""
            Case "Equipped"
                If newValue = PREV_VALUE Then Return "True"
                Dim isEquipment As Boolean = False
                ReadPropertyBool(classId, "IsEquipment", oId, -1, isEquipment, arrParams)
                If Not isEquipment Then Return _ERROR(funcParams(0) & " не является экипировкой.")

                Dim eventId As Integer = 0, res As String = ""
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                If nValue = "True" Then
                    'премет одевается
                    'вызываем события ObjectPutOnEvent
                    eventId = mScript.mainClass(classId).Properties("ObjectPutOnEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, {oId.ToString}, "ObjectPutOnEvent", False)
                        If res = "#Error" Then Return res
                    End If

                    If res <> "False" Then
                        eventId = mScript.mainClass(classId).ChildProperties(oId)("ObjectPutOnEvent").eventId
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, {oId.ToString}, "ObjectPutOnEvent", False)
                            If res = "#Error" Then Return res
                        End If
                    End If
                Else
                    'предмет снимается
                    'вызываем события ObjectPutOffEvent
                    eventId = mScript.mainClass(classId).Properties("ObjectPutOffEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, {oId.ToString, "-1"}, "ObjectPutOffEvent", False)
                        If res = "#Error" Then Return res
                    End If

                    If res <> "False" Then
                        eventId = mScript.mainClass(classId).ChildProperties(oId)("ObjectPutOffEvent").eventId
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, {oId.ToString, "-1"}, "ObjectPutOffEvent", False)
                            If res = "#Error" Then Return res
                        End If
                    End If
                End If

                If res = "False" Then
                    SetPropertyValue(classId, "Equipped", PREV_VALUE, oId)
                    Return "False"
                Else
                    Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
                    If IsNothing(hOuter) = False Then hOuter = hOuter.Children(0)
                    If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".")

                    EquipmentSetClass(oId, funcParams, hOuter, CBool(nValue))
                    Return "True"
                End If
            Case "IsEquipment"
                If newValue = PREV_VALUE Then Return ""
                Return ObjectSetAppearance(oId, funcParams)
            Case "EquipmentCSSclass"
                If newValue = PREV_VALUE Then Return ""
                Dim isEquipment As Boolean = False
                ReadPropertyBool(classId, "IsEquipment", oId, -1, isEquipment, arrParams)
                If Not isEquipment Then Return _ERROR(funcParams(0) & " не является экипировкой.")

                'получаем html-элемент
                Dim hOuter As HtmlElement = GetObjectOuterElement(oId)
                If IsNothing(hOuter) = False Then hOuter = hOuter.Children(0)
                If IsNothing(hOuter) Then Return _ERROR("Ошибка в структуре окна действий. Не найдено место под предмет " & funcParams(0) & ".")

                'получаем все классы и удаляем из них классы экипировки
                Dim strClasses As String = hOuter.GetAttribute("ClassName")
                Dim lstClasses As List(Of String) = strClasses.Split(" "c).ToList
                For i As Integer = lstClasses.Count - 1 To 0 Step -1
                    Dim curClass As String = lstClasses(i)
                    If curClass.StartsWith("equipment", StringComparison.CurrentCultureIgnoreCase) OrElse curClass.StartsWith("equipped", StringComparison.CurrentCultureIgnoreCase) Then
                        lstClasses.RemoveAt(i)
                    End If
                Next

                'добавляем новый класс
                Dim equipped As Boolean = False, equipmentClass As String = ""
                If ReadPropertyBool(classId, "Equipped", oId, -1, equipped, funcParams) = False Then Return "#Error"
                If ReadProperty(classId, "EquipmentCSSclass", oId, -1, equipmentClass, arrParams) = False Then Return "#Error"
                equipmentClass = UnWrapString(equipmentClass)
                If String.IsNullOrEmpty(equipmentClass) = False Then
                    If equipped Then equipmentClass = "equipped" & equipmentClass.Substring(9)
                    lstClasses.Add(equipmentClass)
                End If

                strClasses = Join(lstClasses.ToArray)
                hOuter.SetAttribute("ClassName", strClasses)
                Return ""
        End Select
        Return ""
    End Function

    Private Function GetObjectOuterElement(ByVal oId As Integer) As HtmlElement
        Dim hDoc As HtmlDocument = frmPlayer.wbObjects.Document
        Dim elName As String = "object" & oId.ToString
        Dim outerEl As HtmlElement = hDoc.GetElementById(elName)
        Return outerEl
    End Function
#End Region

#Region "Location"
    Private Function classL_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго порядка
                Return CreateNewObject(classId, funcParams)
            Case "Count"
                Return ObjectsCount(classId, funcParams, False)
            Case "IsExist"
                'Определяет есть ли локация с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "Remove"
                'удаление локации
                If questEnvironment.EDIT_MODE = False Then
                    Dim locId As Integer = CInt(funcParams(0))
                    If locId = GVARS.G_CURLOC Then
                        mScript.LAST_ERROR = "Нельзя удалить текущую локацию!"
                        Return "#Error"
                    ElseIf locId = GVARS.G_PREVLOC Then
                        GVARS.G_PREVLOC = -1
                    End If
                End If
                Return RemoveObject(classId, funcParams)
            Case "CurLoc"
                Return GVARS.G_CURLOC.ToString
            Case "PrevLoc"
                Return GVARS.G_PREVLOC.ToString
            Case "Go"
                Return Go(funcParams)
            Case "SaveActionsState"
                If questEnvironment.EDIT_MODE OrElse GVARS.G_CURLOC < 0 Then Return ""
                actionsRouter.GameSaveActionsState()                
                Return ""
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function ClassL_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String, _
                                          NewValueIsScriptOrLongText As Boolean) As String
        Dim locId As Integer = -1
        If questEnvironment.EDIT_MODE Then Return "" 'works only in game mode
        If PREV_VALUE = newValue AndAlso NewValueIsScriptOrLongText = False Then Return ""
        If funcParams.Count > 0 Then
            If funcParams(0) = "-1" Then Return ""
            locId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If locId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует имени и Id локации {0}!", funcParams(0))
                Return "#Error"
            End If
        Else
            locId = -1
        End If
        If locId <> GVARS.G_CURLOC AndAlso locId <> -1 Then Return "" 'for current location only

        'получаем имя локации
        Dim locName As String = ""
        If locId > -1 Then locName = mScript.mainClass(classId).ChildProperties(locId)("Name").Value
        Dim hDocument As HtmlDocument = frmPlayer.wbMain.Document
        If IsNothing(hDocument) Then Return ""

        Select Case propertyName
            Case "BackColor"
                If locId = -1 Then Return "" 'no screen changing for 1st level (default) value
                Dim bgColor As String = "background:" & UnWrapString(newValue) & ";"
                hDocument.Body.Style &= bgColor
                Return ""
            Case "BackPicture", "BackPicPos", "BackPicStyle"
                If locId = -1 Then Return "" 'no screen changing for 1st level (default) value
                Dim strStyle As New System.Text.StringBuilder, bkPicture As String = ""
                ReadProperty(classId, "BackPicture", locId, -1, bkPicture, funcParams)
                bkPicture = UnWrapString(bkPicture)
                Dim bkPicPos As Integer = 0, bkPicStyle As Integer = 0
                If String.IsNullOrEmpty(bkPicture) = False Then
                    bkPicture = bkPicture.Replace("\"c, "/"c)
                    ReadPropertyInt(classId, "BackPicPos", locId, -1, bkPicPos, funcParams)
                    ReadPropertyInt(classId, "BackPicStyle", locId, -1, bkPicStyle, funcParams)
                End If

                If String.IsNullOrEmpty(bkPicture) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, bkPicture)) Then
                    'файл картинки существует
                    strStyle.AppendFormat("background-image: url({0});", "'" + bkPicture + "'")
                    '0 простая загрузка, 1 - заполнить, 2 - масштабировать, 3 - размножить, 4 - размножить по Х, 5 - размножить по Y 
                    strStyle.Append("background-repeat:")
                    Select Case bkPicStyle
                        Case 0 '0 простая загрузка
                            strStyle.Append("no-repeat;background-size:auto;")
                            'HTMLRemoveCSSstyle(hDocument.Body, "background-size")
                        Case 1 '1 растянуть пропорционально
                            strStyle.Append("no-repeat;background-size:cover;")
                        Case 2 '2 заполнить
                            strStyle.Append("no-repeat;background-size:contain;")
                            'обязательно указываем высоту окна, иначе отображается некорректно
                            strStyle.AppendFormat("height:{0}px;", frmPlayer.wbMain.ClientSize.Height - CONST_WBHEIGHT_CORRECTION)
                        Case 3 '3 масштабировать
                            strStyle.Append("repeat;background-size:auto;")
                        Case 4 '4 размножить по Х
                            strStyle.Append("repeat-x;background-size:auto;")
                        Case 5 '5 размножить по Y
                            strStyle.Append("repeat-y;background-size:auto;")
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

                hDocument.Body.Style &= strStyle.ToString
                Return ""
            Case "Description", "DescriptionBackColor", "DescriptionTextColor", "DescriptionWidth", "DescriptionHeight", "DescriptionCssClass", "OffsetByX", "OffsetByY"
                If propertyName <> "Description" AndAlso locId = -1 Then Return "" 'no screen changing for 1st level (default) value

                Dim hConvas As HtmlElement = hDocument.GetElementById("MainConvas")
                If IsNothing(hConvas) Then Return ""
                Return ShowDescription(hConvas, classId, "Description", locId, -1, funcParams, -1, "DescriptionText")
            Case "Header"
                Dim hConvas As HtmlElement = hDocument.GetElementById("MainConvas")
                If IsNothing(hConvas) Then Return ""
                Dim hHeader As HtmlElement = hConvas.Document.GetElementById("header")
                If IsNothing(hHeader) Then Return ""

                Dim fInner As String = hHeader.InnerHtml
                If IsNothing(fInner) Then fInner = ""
                Dim eventId As Integer = mScript.mainClass(classId).Properties("Header").eventId
                If eventId > 0 Then
                    'колонтитул задан
                    Dim result As String = ""
                    ReadProperty(classId, "Header", -1, -1, result, funcParams)
                    If String.IsNullOrEmpty(result) = False AndAlso result <> "#Error" AndAlso result <> fInner Then
                        hHeader.InnerHtml = result
                    End If
                End If
                If String.IsNullOrEmpty(hHeader.InnerHtml) Then
                    hHeader.Style = "display:none"
                Else
                    hHeader.Style = ""
                End If
            Case "Footer"
                Dim hConvas As HtmlElement = hDocument.GetElementById("MainConvas")
                If IsNothing(hConvas) Then Return ""
                Dim hFooter As HtmlElement = hConvas.Document.GetElementById("footer")
                If IsNothing(hFooter) Then Return ""

                Dim fInner As String = hFooter.InnerHtml
                If IsNothing(fInner) Then fInner = ""
                Dim eventId As Integer = mScript.mainClass(classId).Properties("Footer").eventId
                If eventId > 0 Then
                    'колонтитул задан
                    Dim result As String = ""
                    ReadProperty(classId, "Footer", -1, -1, result, funcParams)
                    If String.IsNullOrEmpty(result) = False AndAlso result <> "#Error" AndAlso result <> fInner Then
                        hFooter.InnerHtml = result
                    End If
                Else
                    fInner = ""
                    hFooter.InnerHtml = ""
                End If
                If String.IsNullOrEmpty(hFooter.InnerHtml) Then
                    hFooter.Style = "display:none"
                Else
                    hFooter.Style = ""
                End If
                Return ""
        End Select
        Return ""
    End Function
#End Region

#Region "LocWindow"
    Private Function classLW_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "PR"
                Dim strText As String = GetParam(funcParams, 0, "")
                strText = UnWrapString(strText)
                If IsNothing(PrintTextToMainWindow("", strText, PrintInsertionEnum.NEW_BLOCK, "", PrintDataEnum.TEXT, funcParams)) Then Return "#Error"
                Return ""
            Case "PRH"
                Dim strText As String = GetParam(funcParams, 0, "")
                strText = UnWrapString(strText)
                If String.IsNullOrEmpty(strText) Then Return ""
                If IsNothing(PrintTextToMainWindow("0", strText, PrintInsertionEnum.APPEND, "", PrintDataEnum.TEXT, funcParams)) Then Return "#Error"
                Return ""
            Case "RemoveText"
                Return RemoveText(funcParams(0), funcParams).ToString
            Case "Print"
                Dim strText As String = UnWrapString(funcParams(0))
                Dim txtId As String = GetParam(funcParams, 1, "")
                Dim appendType As PrintInsertionEnum = Val(UnWrapString(GetParam(funcParams, 2, "0")))
                Dim newId As String = UnWrapString(GetParam(funcParams, 3, ""))
                Dim showTime As Integer = Val(GetParam(funcParams, 4, "0"))
                Dim usePresets As Boolean = CBool(GetParam(funcParams, 5, "True"))
                Dim aStyles As String = GetParam(funcParams, 6, "")
                Dim hEl As HtmlElement = PrintTextToMainWindow(txtId, strText, appendType, newId, PrintDataEnum.TEXT, funcParams, aStyles, usePresets)
                If IsNothing(hEl) Then Return "#Error"
                If showTime <= 0 Then Return ""

                Dim tim As New Timer With {.Enabled = True, .Interval = showTime}
                AddHandler tim.Tick, Sub(sender As Object, e As EventArgs)
                                         If IsNothing(hEl) = False Then
                                             Dim hNode As mshtml.IHTMLDOMNode = hEl.DomElement
                                             hNode.removeNode(True)
                                         End If
                                         tim.Enabled = False
                                         tim.Dispose()
                                     End Sub
                Return ""
            Case "PrintPicture"
                Dim strPath As String = UnWrapString(funcParams(0))
                Dim pWidth As String = UnWrapString(GetParam(funcParams, 1, ""))
                Dim pHeight As String = UnWrapString(GetParam(funcParams, 2, ""))
                Dim pAlign As String = UnWrapString(GetParam(funcParams, 3, ""))
                Dim txtId As String = GetParam(funcParams, 4, "")
                Dim appendType As PrintInsertionEnum = Val(UnWrapString(GetParam(funcParams, 5, "0")))
                Dim newId As String = UnWrapString(GetParam(funcParams, 6, ""))
                Dim showTime As Integer = Val(GetParam(funcParams, 7, "0"))
                Dim usePresets As Boolean = CBool(GetParam(funcParams, 8, "False"))
                Dim aStyles As String = GetParam(funcParams, 9, "")
                Dim hEl As HtmlElement = PrintTextToMainWindow(txtId, strPath, appendType, newId, PrintDataEnum.PICTURE, funcParams, aStyles, usePresets, pWidth, pHeight, pAlign)
                If IsNothing(hEl) Then Return "#Error"
                If showTime <= 0 Then Return ""

                Dim tim As New Timer With {.Enabled = True, .Interval = showTime}
                AddHandler tim.Tick, Sub(sender As Object, e As EventArgs)
                                         If IsNothing(hEl) = False Then
                                             Dim hNode As mshtml.IHTMLDOMNode = hEl.DomElement
                                             hNode.removeNode(True)
                                         End If
                                         tim.Enabled = False
                                         tim.Dispose()
                                     End Sub
                Return ""
            Case "PrintInfo"
                Dim strText As String = UnWrapString(funcParams(0))
                Dim imgPath As String = UnWrapString(funcParams(1))
                Dim strCaption As String = UnWrapString(GetParam(funcParams, 2, ""))
                Dim strStyle As String = UnWrapString(GetParam(funcParams, 3, "chat"))
                Dim imgPos As Integer = Val(UnWrapString(GetParam(funcParams, 4, "0")))
                Dim imgWidth As String = UnWrapString(GetParam(funcParams, 5, ""))
                Dim imgHeight As String = UnWrapString(GetParam(funcParams, 6, ""))
                Dim txtId As String = GetParam(funcParams, 7, "")
                Dim appendType As PrintInsertionEnum = Val(UnWrapString(GetParam(funcParams, 8, "0")))
                Dim newId As String = UnWrapString(GetParam(funcParams, 9, ""))
                Dim showTime As Integer = Val(GetParam(funcParams, 10, "0"))
                Dim bgText As String = UnWrapString(GetParam(funcParams, 11, ""))
                Dim fgText As String = UnWrapString(GetParam(funcParams, 12, ""))
                Dim bgImg As String = UnWrapString(GetParam(funcParams, 13, ""))
                Dim bgCaption As String = UnWrapString(GetParam(funcParams, 14, ""))
                Dim fgCaption As String = UnWrapString(GetParam(funcParams, 15, ""))

                Dim hEl As HtmlElement = PrintInfoToMainWindow(txtId, strText, imgPath, strCaption, strStyle, appendType, newId, imgPos = 0, imgWidth, imgHeight, bgText, bgImg, bgCaption, fgText, fgCaption, arrParams)
                If IsNothing(hEl) Then Return "#Error"
                If showTime <= 0 Then Return ""

                Dim tim As New Timer With {.Enabled = True, .Interval = showTime}
                AddHandler tim.Tick, Sub(sender As Object, e As EventArgs)
                                         If IsNothing(hEl) = False Then
                                             Dim hNode As mshtml.IHTMLDOMNode = hEl.DomElement
                                             hNode.removeNode(True)
                                         End If
                                         tim.Enabled = False
                                         tim.Dispose()
                                     End Sub
                Return ""
            Case "Clear"
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return ""
                Dim hConvas As HtmlElement = hDoc.GetElementById("MainConvas")
                If IsNothing(hConvas) Then Return ""
                hConvas.InnerHtml = ""
                Return ""
            Case "ClassAdd"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть документ главного окна.", functionName)

                Dim elId As String = UnWrapString(funcParams(0))
                Dim hEl As HtmlElement = hDoc.GetElementById(elId)
                If IsNothing(hEl) Then Return _ERROR("Html-элемента с Id " & elId & " не найдено.", functionName)
                Dim strClass As String = UnWrapString(funcParams(1))

                Return HTMLAddClass(hEl, strClass).ToString
            Case "ClassRemove"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть документ главного окна.", functionName)

                Dim elId As String = UnWrapString(funcParams(0))
                Dim hEl As HtmlElement = hDoc.GetElementById(elId)
                If IsNothing(hEl) Then Return _ERROR("Html-элемента с Id " & elId & " не найдено.", functionName)
                Dim strClass As String = UnWrapString(funcParams(1))

                Return HTMLRemoveClass(hEl, strClass).ToString
            Case "ClassSwitch"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть документ главного окна.", functionName)

                Dim elId As String = UnWrapString(funcParams(0))
                Dim hEl As HtmlElement = hDoc.GetElementById(elId)
                If IsNothing(hEl) Then Return _ERROR("Html-элемента с Id " & elId & " не найдено.", functionName)
                Dim strClass As String = UnWrapString(funcParams(1))

                Return HTMLSwitchClass(hEl, strClass).ToString
            Case "ClassReplace"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть документ главного окна.", functionName)

                Dim elId As String = UnWrapString(funcParams(0))
                Dim hEl As HtmlElement = hDoc.GetElementById(elId)
                If IsNothing(hEl) Then Return _ERROR("Html-элемента с Id " & elId & " не найдено.", functionName)
                Dim strClassOld As String = UnWrapString(funcParams(1))
                Dim strClassNew As String = UnWrapString(funcParams(2))

                Return HTMLReplaceClass(hEl, strClassOld, strClassNew).ToString
            Case "ClassHas"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть документ главного окна.", functionName)

                Dim elId As String = UnWrapString(funcParams(0))
                Dim hEl As HtmlElement = hDoc.GetElementById(elId)
                If IsNothing(hEl) Then Return _ERROR("Html-элемента с Id " & elId & " не найдено.", functionName)
                Dim strClass As String = UnWrapString(funcParams(1))

                Return HTMLHasClass(hEl, strClass).ToString
            Case "ClassesClear"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть документ главного окна.", functionName)

                Dim elId As String = UnWrapString(funcParams(0))
                Dim hEl As HtmlElement = hDoc.GetElementById(elId)
                If IsNothing(hEl) Then Return _ERROR("Html-элемента с Id " & elId & " не найдено.", functionName)
                Dim strClass As String = hEl.GetAttribute("ClassName")
                If String.IsNullOrEmpty(strClass) Then Return "False"

                hEl.SetAttribute("ClassName", "")
                Return "True"
            Case "AddRule"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim strSelector As String = UnWrapString(funcParams(0))
                Dim strStyle As String = UnWrapString(funcParams(1))

                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть документ главного окна.", functionName)
                Dim msDoc As mshtml.HTMLDocument = hDoc.DomDocument

                Dim hStyle As mshtml.IHTMLStyleSheet
                If msDoc.styleSheets.length = 0 Then
                    hStyle = msDoc.createStyleSheet()
                Else
                    hStyle = msDoc.styleSheets(0)
                End If
                hStyle.addRule(strSelector, strStyle)
                Return ""
            Case "SetAttribute"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть документ главного окна.", functionName)

                Dim elId As String = UnWrapString(funcParams(0))
                Dim hEl As HtmlElement = hDoc.GetElementById(elId)
                If IsNothing(hEl) Then Return _ERROR("Html-элемента с Id " & elId & " не найдено.", functionName)

                Dim attrName As String = UnWrapString(funcParams(1))
                Dim attValue As String = UnWrapString(funcParams(2))
                hEl.SetAttribute(attrName, attValue)
                Return ""
            Case "GetAttribute"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть документ главного окна.", functionName)

                Dim elId As String = UnWrapString(funcParams(0))
                Dim hEl As HtmlElement = hDoc.GetElementById(elId)
                If IsNothing(hEl) Then Return _ERROR("Html-элемента с Id " & elId & " не найдено.", functionName)

                Dim attrName As String = UnWrapString(funcParams(1))
                Return WrapString(hEl.GetAttribute(attrName))
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function classB_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Dim classH As Integer = mScript.mainClassHash("H"), classArmy As Integer = mScript.mainClassHash("Army")

        Select Case functionName
            Case "AddFighter"
                Dim hId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classH).ChildProperties)
                If hId < 0 Then Return _ERROR("Героя " & funcParams(0) & " не существует!")
                Dim fName As String = GetParam(funcParams, 1, "")
                Dim armyId As Integer = -1
                If funcParams.Count > 2 AndAlso funcParams(2) <> "''" AndAlso funcParams(2) <> "-1" Then
                    armyId = GetSecondChildIdByName(funcParams(2), mScript.mainClass(classArmy).ChildProperties)
                    If armyId < 0 Then Return _ERROR("Армии " & funcParams(2) & " не существует!")
                End If
                Dim owner As String = GetParam(funcParams, 3, ""), ownerId As Integer = -1
                If String.IsNullOrEmpty(owner) = False Then
                    ownerId = mScript.Battle.GetFighterByName(owner)                    
                End If

                Dim fId = mScript.Battle.AddFighter(hId, fName, armyId, -1, ownerId)
                If fId < -1 Then Return "#Error"
                If GVARS.G_ISBATTLE Then
                    If mScript.Battle.AppendFighterInBattle(fId) = "#Error" Then Return "#Error"

                    If armyId > -1 Then
                        'Событие ArmyOnCountChanged (0 - пришел на подмогу / 2 - был призван магией)
                        'получаем параметры
                        Dim unitsLeft As Integer = mScript.Battle.UnitsInArmyLeft(armyId)
                        If unitsLeft < 0 Then Return "#Error"
                        Dim reason As String = "0"
                        If ownerId > -1 Then reason = "2"
                        Dim arrs() As String = {armyId.ToString, fId.ToString, reason, "1", unitsLeft.ToString}

                        'глобальное событие
                        Dim eventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnCountChanged").eventId
                        Dim res As String = ""
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                            If res = "#Error" Then Return "#Error"
                        End If

                        'событие армии
                        If res <> "False" Then
                            eventId = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnCountChanged").eventId
                            If eventId > 0 Then
                                res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                                If res = "#Error" Then Return "#Error"
                            End If
                        End If
                    End If
                End If
                Return fId.ToString
            Case "AddArmy"
                Dim armyId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classArmy).ChildProperties)
                If armyId < 0 Then Return _ERROR("Армии " & funcParams(0) & " не существует!")
                Dim added As Integer = mScript.Battle.AddArmy(armyId)
                If added < 0 Then
                    Return "#Error"
                ElseIf added = 0 Then
                    Return "0"
                End If

                'Событие ArmyOnCountChanged (пришел на подмогу)
                'получаем параметры
                Dim unitsLeft As Integer = mScript.Battle.UnitsInArmyLeft(armyId)
                If unitsLeft < 0 Then Return "#Error"
                Dim arrs() As String = {armyId.ToString, (mScript.Battle.Fighters.Count - added).ToString, "0", added.ToString, unitsLeft.ToString}

                'глобальное событие
                Dim eventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnCountChanged").eventId
                Dim res As String = ""
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                    If res = "#Error" Then Return "#Error"
                End If

                'событие армии
                If res <> "False" Then
                    eventId = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnCountChanged").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                        If res = "#Error" Then Return "#Error"
                    End If
                End If

                Return added.ToString
            Case "AddArmyUnit"
                Dim armyId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classArmy).ChildProperties)
                If armyId < 0 Then Return _ERROR("Армии " & funcParams(0) & " не существует!")
                Dim unitId As Integer = GetThirdChildIdByName(funcParams(1), armyId, mScript.mainClass(classArmy).ChildProperties)
                If unitId < 0 Then Return _ERROR("Бойца " & funcParams(1) & " в армии " & funcParams(0) & " не существует!")

                Dim owner As String = GetParam(funcParams, 2, ""), ownerId As Integer = -1
                If String.IsNullOrEmpty(owner) = False Then
                    ownerId = mScript.Battle.GetFighterByName(owner)
                End If

                Dim added As Integer = mScript.Battle.AddArmyUnit(armyId, unitId, owner)
                If added < 0 Then
                    Return "#Error"
                ElseIf added = 0 Then
                    Return "0"
                End If

                'Событие ArmyOnCountChanged (пришел на подмогу)
                'получаем параметры
                Dim unitsLeft As Integer = mScript.Battle.UnitsInArmyLeft(armyId)
                If unitsLeft < 0 Then Return "#Error"
                Dim arrs() As String = {armyId.ToString, (mScript.Battle.Fighters.Count - added).ToString, "0", added.ToString, unitsLeft.ToString}

                'глобальное событие
                Dim eventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnCountChanged").eventId
                Dim res As String = ""
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                    If res = "#Error" Then Return "#Error"
                End If

                'событие армии
                If res <> "False" Then
                    eventId = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnCountChanged").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                        If res = "#Error" Then Return "#Error"
                    End If
                End If

                Return added.ToString
            Case "Begin"
                Return mScript.Battle.Begin
            Case "FightersCount"
                Dim aliveOnly As Boolean = CBool(GetParam(funcParams, 0, "True"))
                Dim fType As Integer = Val(GetParam(funcParams, 1, "0"))
                If fType < 0 OrElse fType > 7 Then Return _ERROR("Тип бойцов для подсчета указан неверно.", functionName)
                Return mScript.Battle.FightersCount(aliveOnly, fType).ToString
            Case "GetCurrentVictimsList"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                Dim var As New cVariable.variableEditorInfoType
                Dim lst As List(Of Integer) = mScript.Battle.lstVictims
                If lst.Count = 0 Then
                    ReDim var.arrValues(0)
                    var.arrValues(0) = "-1"
                Else
                    ReDim var.arrValues(lst.Count - 1)
                    For i As Integer = 0 To lst.Count - 1
                        var.arrValues(i) = lst(i).ToString
                    Next
                End If
                'восстанавливаем исходный список
                mScript.lastArray = var
                Return "#ARRAY"
            Case "GetVictimsList"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                Dim var As New cVariable.variableEditorInfoType
                Dim heroId As Integer = mScript.Battle.GetFighterByName(funcParams(0))
                If heroId < 0 Then Return _ERROR("Бойца " & funcParams(0) & " не существует.", functionName)
                Dim blnFriend As Boolean = CBool(GetParam(funcParams, 1, "False"))
                'создаем копию текущего списка
                Dim lstCopy As New List(Of Integer)
                If mScript.Battle.lstVictims.Count > 0 Then lstCopy.AddRange(mScript.Battle.lstVictims)
                If mScript.Battle.PrepareVictimsList(heroId, blnFriend) = "#Error" Then Return "#Error"
                Dim lst As List(Of Integer) = mScript.Battle.lstVictims
                If lst.Count = 0 Then
                    ReDim var.arrValues(0)
                    var.arrValues(0) = "-1"
                Else
                    ReDim var.arrValues(lst.Count - 1)
                    For i As Integer = 0 To lst.Count - 1
                        var.arrValues(i) = lst(i).ToString
                    Next
                End If
                'восстанавливаем исходный список
                mScript.Battle.lstVictims.Clear()
                mScript.Battle.lstVictims.AddRange(lstCopy)
                lstCopy.Clear()
                mScript.lastArray = var
                Return "#ARRAY"
            Case "GetNewName"
                Return mScript.Battle.GetNewName(funcParams(0))
            Case "GetNextFighter"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                Return mScript.Battle.GetNextFighter(GVARS.G_CURFIGHTER, False).ToString
            Case "GetSummonOwner"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                Dim fId As Integer = mScript.Battle.GetFighterByName(funcParams(0))
                If fId < 0 Then Return _ERROR("Боец " & funcParams(0) & " не найден!", functionName)
                Dim getMain As Boolean = CBool(GetParam(funcParams, 1, "False"))
                Return mScript.Battle.GetSummonOwner(fId, getMain).ToString
            Case "HitVictim"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                If GVARS.G_CURFIGHTER < 0 Then Return _ERROR("Еще не выбран атакующий боец.", functionName)
                If GVARS.G_STRIKETYPE < 0 Then Return _ERROR("Еще не выбран тип атаки.", functionName)
                Return mScript.Battle.Hit(GVARS.G_CURFIGHTER, GVARS.G_CURVICTIM, GVARS.G_STRIKETYPE)
            Case "IsBattle"
                Return GVARS.G_ISBATTLE.ToString
            Case "IsEnemy"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                Dim checkId As Integer = mScript.Battle.GetFighterByName(funcParams(0))
                If checkId < 0 Then Return _ERROR("Боец " & funcParams(0) & " не найден!", functionName)
                Dim fId As Integer = -1
                If funcParams.Count > 1 Then
                    fId = mScript.Battle.GetFighterByName(funcParams(1))
                    If fId < 0 Then Return _ERROR("Боец " & funcParams(1) & " не найден!", functionName)
                End If

                Dim res As cBattle.IsEnemyEnum = mScript.Battle.IsEnemy(fId, checkId)
                If res = cBattle.IsEnemyEnum.CheckError Then Return "#Error"
                Return GetFunctionReturnEnumTextById(classH, functionName, res)
            Case "IsFighterActive"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                Dim fId As Integer = mScript.Battle.GetFighterByName(funcParams(0))
                If fId < 0 Then Return _ERROR("Боец " & funcParams(0) & " не найден!", functionName)

                Dim res As Integer = 0 '0 - активен, 1 - погиб, 2 - сбежал
                Dim life As Double = 0, runAway As Boolean = False
                If mScript.Battle.ReadFighterPropertyDbl("Life", fId, life, {fId.ToString}) = False Then Return "#Error"
                If life <= 0 Then
                    res = 1
                Else
                    If mScript.Battle.ReadFighterPropertyBool("RunAway", fId, runAway, {fId.ToString}) = False Then Return "#Error"
                    If runAway Then res = 2
                End If

                Return GetFunctionReturnEnumTextById(classH, functionName, res)
            Case "KillSummoned"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                Dim hostId As Integer = GVARS.G_CURFIGHTER
                If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then
                    hostId = mScript.Battle.GetFighterByName(funcParams(0))
                    If hostId < 0 Then Return _ERROR("Боец " & funcParams(0) & " не найден!", functionName)
                ElseIf GVARS.G_CURFIGHTER < 0 Then
                    Return _ERROR("На данный момент текущий боец н выбран.", functionName)
                End If

                Return mScript.Battle.KillSummoned(hostId).ToString
            Case "NewTurn"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                mScript.Battle.NewTurn()
                If mScript.LAST_ERROR.Length = 0 Then
                    Return ""
                Else
                    Return "#Error"
                End If
            Case "NextTurn"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя.", functionName)
                Return mScript.Battle.NextTurn()
            Case "RandomizeDamage"
                Dim dam As Double = Val(funcParams(0))
                Dim maxValue As Double = Double.MaxValue
                If funcParams.Count > 1 Then maxValue = Val(funcParams(1))
                Dim digits As Integer = Val(GetParam(funcParams, 2, "0"))
                Return mScript.Battle.RandomizeDamage(dam, maxValue, digits).ToString(provider_points)
            Case "RemoveFighter"
                If GVARS.G_ISBATTLE Then
                    Return _ERROR("Нельзя удалить бойца во время битвы! Он может только умереть или сбежать (Life = 0 или RunAway = True)!")
                End If

                Dim fId As Integer = -1
                If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 AndAlso funcParams(0) <> "-1" Then
                    fId = mScript.Battle.GetFighterByName(funcParams(0))
                    If fId < 0 Then Return _ERROR("Боец " & funcParams(0) & " не найден!", functionName)
                End If

                Return mScript.Battle.RemoveFighter(fId)
            Case "SelectAttackType"
                If GVARS.G_CURFIGHTER < 0 Then Return _ERROR("Не выбран атакующий боец!", functionName)
                If FunctionRouter(mScript.mainClassHash("A"), "Remove", {"-1"}, Nothing) = "#Error" Then Return "#Error"
                'выбор типа удара - событие BattleAttackTypeSelect
                Dim isUnderControl As Boolean = mScript.Battle.IsFighterUnderControl(GVARS.G_CURFIGHTER)
                Dim eventId As Integer = mScript.mainClass(classId).Properties("BattleAttackTypeSelect").eventId
                Dim res As String = mScript.eventRouter.RunEvent(eventId, {GVARS.G_CURFIGHTER.ToString, GVARS.G_CURVICTIM.ToString, isUnderControl.ToString}, "BattleAttackTypeSelect", False)
                If res = "#Error" Then Return res

                If Not isUnderControl Then
                    'боец неподконтрольный
                    If res = "False" Then Return ""
                    If GVARS.G_STRIKETYPE < 0 Then GVARS.G_STRIKETYPE = mScript.Battle.ChooseStrikeType(GVARS.G_CURFIGHTER)
                    If GVARS.G_STRIKETYPE = -1 Then Return "#Error"
                End If
                Return ""
            Case "SelectVictim"
                If GVARS.G_CURFIGHTER < 0 Then Return _ERROR("Не выбран атакующий боец!", functionName)
                If GVARS.G_STRIKETYPE < 0 Then Return _ERROR("Не выбран тип атаки!", functionName)
                If FunctionRouter(mScript.mainClassHash("A"), "Remove", {"-1"}, Nothing) = "#Error" Then Return "#Error"
                'должна ли быть жертва?
                Dim ShoudBeVictim As Boolean = True
                If GVARS.G_STRIKETYPE = 1 Then
                    Dim classMg As Integer = mScript.mainClassHash("Mg")
                    Dim mType As cBattle.MagicTypesEnum = cBattle.MagicTypesEnum.Damage
                    Dim magicBookId As Integer = mScript.Battle.Fighters(GVARS.G_CURFIGHTER).magicBookId
                    If ReadPropertyInt(classMg, "MagicType", magicBookId, GVARS.G_CURMAGIC, mType, {magicBookId.ToString, GVARS.G_CURMAGIC.ToString}) = False Then Return "#Error"
                    If mType = cBattle.MagicTypesEnum.FriendsEnhancer OrElse mType = cBattle.MagicTypesEnum.OtherFriend Then
                        If mScript.Battle.PrepareVictimsList(GVARS.G_CURFIGHTER, True) = False Then Return "#Error"
                        If mScript.Battle.lstVictims.Count = 0 Then Return _ERROR("Использование магии " & mScript.mainClass(classMg).ChildProperties(magicBookId)("Name").ThirdLevelProperties(GVARS.G_CURMAGIC) & " не возможно! В битве нет ни одного вашего союзника!")
                        ShoudBeVictim = True
                    ElseIf mType = cBattle.MagicTypesEnum.Damage OrElse mType = cBattle.MagicTypesEnum.OtherWithvictim Then
                        ShoudBeVictim = True
                    Else
                        ShoudBeVictim = False
                    End If
                End If
                If Not ShoudBeVictim Then
                    GVARS.G_CURVICTIM = -1
                    Return ""
                End If

                'Содержит данный блок код выбора жертвы. В случае, если выбор предоставляется Игроку с помощью действий, то произойдет выход из функции.
                'Событие BattleOnVictimSelect
                Dim isUnderControl As Boolean = mScript.Battle.IsFighterUnderControl(GVARS.G_CURFIGHTER)
                If FunctionRouter(mScript.mainClassHash("A"), "Remove", {"-1"}, Nothing) = "#Error" Then Return "#Error"
                Dim eventId As Integer = mScript.mainClass(classId).Properties("BattleOnVictimSelect").eventId
                Dim res As String = ""
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, {GVARS.G_CURFIGHTER.ToString, mScript.Battle.lstVictims.Count, isUnderControl.ToString}, "BattleOnVictimSelect", False)
                    If res = "#Error" Then
                        Return res
                    ElseIf res = "False" Then
                        'Остановка скрипта BattleOnVictimSelect
                        Return ""
                    End If
                End If

                'выбор жертвы
                If isUnderControl Then
                    'для подконтрольных бойцов
                    If mScript.Battle.lstVictims.Count = 1 Then
                        GVARS.G_CURVICTIM = mScript.Battle.lstVictims(0)
                    Else
                        Return "" 'остановка так как не выбрана жертва (это должно произойти при выборе действия, описанного в BattleOnVictimSelect)
                    End If
                Else
                    'для НЕподконтрольных бойцов - автовыбор жертвы
                    GVARS.G_CURVICTIM = mScript.Battle.ChooseVictim(GVARS.G_CURFIGHTER)
                End If
                Return ""
            Case "Stop"
                Return mScript.Battle.StopBattle
            Case "ClearHistory"
                mScript.Battle.lstHistory.Clear()
                Return ""
            Case "SaveHistory"
                Dim var As New cVariable.variableEditorInfoType
                Dim lst As List(Of String) = mScript.Battle.lstHistory
                If lst.Count = 0 Then
                    ReDim var.arrValues(0)
                    var.arrValues(0) = ""
                Else
                    ReDim var.arrValues(lst.Count - 1)
                    For i As Integer = 0 To lst.Count - 1
                        var.arrValues(i) = WrapString(lst(i))
                    Next
                End If
                mScript.lastArray = var
                Return "#ARRAY"
            Case "LoadHistory"
                Dim var As cVariable.variableEditorInfoType = GetVariable(funcParams(0))
                If IsNothing(var) Then Return _ERROR("Переменной " & funcParams(0) & " не существует!", functionName)
                Dim lst As List(Of String) = mScript.Battle.lstHistory
                lst.Clear()
                If IsNothing(var.arrValues) OrElse var.arrValues.Count = 0 OrElse (var.arrValues.Count = 1 AndAlso var.arrValues(0) = "") Then Return ""
                For i As Integer = 0 To var.arrValues.Count - 1
                    lst.Add(UnWrapString(var.arrValues(i)))
                Next
                Return ""
                'prop KeepHistory
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function ClassB_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If questEnvironment.EDIT_MODE Then Return ""

        Dim Battle As cBattle = mScript.Battle
        Select Case propertyName
            Case "ArmiesOffset", "HistoryOffset"
                If GVARS.G_ISBATTLE Then
                    Return Battle.OffsetArmiesContainers(Nothing)
                End If
            Case "BackColor"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    hDoc.Body.Style &= "background:" & UnWrapString(newValue) & ";"
                End If
            Case "BackPicPos"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    Dim strStyle As New System.Text.StringBuilder
                    strStyle.Append("background-repeat:")
                    Select Case Val(UnWrapString(newValue))
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

                    hDoc.Body.Style &= strStyle.ToString
                End If
            Case "BackPicStyle"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    Dim strStyle As New System.Text.StringBuilder
                    strStyle.Append("background-position:")
                    Select Case Val(UnWrapString(newValue))
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

                    hDoc.Body.Style &= strStyle.ToString
                End If
            Case "BackPicture"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    Dim strStyle As New System.Text.StringBuilder
                    strStyle.AppendFormat("background-image: url({0});", "'" + UnWrapString(newValue) + "'")
                    hDoc.Body.Style &= strStyle.ToString
                End If
            Case "Caption"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    Dim hHeader As HtmlElement = hDoc.GetElementById("BattleHeader")
                    If IsNothing(hHeader) Then Return ""
                    newValue = UnWrapString(newValue)
                    If String.IsNullOrEmpty(newValue) Then
                        hHeader.InnerHtml = ""
                        hHeader.Style = "display:none;"
                    Else
                        hHeader.InnerHtml = newValue
                        hHeader.Style = ""
                    End If
                End If
            Case "CSS"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    HtmlChangeFirstCSSLink(hDoc, UnWrapString(newValue))
                End If
            Case "CurFighter"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Свойство дойтупно только во время боя.")
                Dim fId As Integer = Battle.GetFighterByName(newValue)
                If fId < 0 Then Return _ERROR("Бойца " & newValue & " не существует.")
                GVARS.G_CURFIGHTER = fId
                Return ""
            Case "CurVictim"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Свойство дойтупно только во время боя.")
                Dim fId As Integer = Battle.GetFighterByName(newValue)
                If fId < 0 Then Return _ERROR("Бойца " & newValue & " не существует.")
                GVARS.G_CURVICTIM = fId
                Return ""
            Case "FightersSpeed"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    For fId As Integer = 0 To Battle.Fighters.Count - 1
                        Dim hFighter As HtmlElement = Battle.FindFighterHTMLContainer(fId, hDoc)
                        If IsNothing(hFighter) Then Continue For
                        If Battle.PutFighterToBattleField(fId, hFighter.Parent, hFighter) = "#Error" Then Return "#Error"
                    Next fId

                    Dim hHistory As HtmlElement = hDoc.GetElementById("MainConvas")
                    If IsNothing(hHistory) Then Return ""
                    hHistory.Style &= "transition-duration:" & newValue & "ms;"
                End If
            Case "FooterVisible"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    Dim visible As Boolean = CBool(newValue)

                    Dim hFooter As HtmlElement = hDoc.GetElementById("footer")
                    If IsNothing(hFooter) Then Return ""
                    If visible Then
                        hFooter.Style = ""
                    Else
                        hFooter.Style = "display:none;"
                    End If
                End If
            Case "StrikeType"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Свойство дойтупно только во время боя.")
                GVARS.G_STRIKETYPE = Val(UnWrapString(newValue))
                If GVARS.G_CURFIGHTER > -1 Then Battle.SetAttackPicture(GVARS.G_CURFIGHTER)
            Case "PlayList"
                If GVARS.G_ISBATTLE Then
                    Dim classMed As Integer = mScript.mainClassHash("Med")
                    If newValue = "-1" OrElse newValue = "''" OrElse newValue = "" Then
                        If FunctionRouter(classMed, "SelectList", {"-1"}, Nothing) = "#Error" Then Return "#Error"
                    Else
                        If FunctionRouter(classMed, "SelectList", {newValue}, Nothing) = "#Error" Then Return "#Error"
                        If FunctionRouter(classMed, "StartPlay", Nothing, Nothing) = "#Error" Then Return "#Error"
                    End If
                End If
            Case "TextColor"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")

                    Dim hHistory As HtmlElement = hDoc.GetElementById("MainConvas")
                    If IsNothing(hHistory) Then Return ""
                    hHistory.Style &= "color:" & UnWrapString(newValue) & ";"
                End If
            Case "VSPictureWidth", "VSPictureHeight"
                If GVARS.G_ISBATTLE Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    Dim imgVS As HtmlElement = hDoc.GetElementById("vsPicture")
                    If IsNothing(imgVS) Then Return ""

                    Dim vsWidth As Integer = 0, vsHeight As Integer = 0
                    If ReadPropertyInt(classId, "VSPictureWidth", -1, -1, vsWidth, Nothing) = False Then Return "#Error"
                    If ReadPropertyInt(classId, "VSPictureHeight", -1, -1, vsHeight, Nothing) = False Then Return "#Error"
                    If vsWidth <= 0 Then vsWidth = 200
                    If vsHeight <= 0 Then vsHeight = 200
                    imgVS.Style &= "width:" & vsWidth.ToString & "px;height:" & vsHeight.ToString & "px;"
                End If
            Case Else
                If GVARS.G_ISBATTLE AndAlso propertyName.StartsWith("PropertyToDisplay") Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ главного окна!")
                    For fId As Integer = 0 To Battle.Fighters.Count - 1
                        Dim hFighter As HtmlElement = Battle.FindFighterHTMLContainer(fId, hDoc)
                        If IsNothing(hFighter) Then Continue For
                        If Battle.PutFighterToBattleField(fId, hFighter.Parent, hFighter) = "#Error" Then Return "#Error"
                    Next fId
                End If
        End Select
        Return ""
    End Function

    Private Function ClassB_PropertiesGet(ByVal classId As Integer, ByVal propertyName As String, ByVal defValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If questEnvironment.EDIT_MODE Then Return defValue

        Dim Battle As cBattle = mScript.Battle
        Select Case propertyName
            Case "CurFighter"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Свойство дойтупно только во время боя.")
                Return GVARS.G_CURFIGHTER.ToString
            Case "CurVictim"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Свойство дойтупно только во время боя.")
                Return GVARS.G_CURVICTIM.ToString
            Case "StrikeType"
                If GVARS.G_ISBATTLE = False Then Return _ERROR("Свойство дойтупно только во время боя.")
                Return GVARS.G_STRIKETYPE.ToString
        End Select
        Return defValue
    End Function

#End Region

#Region "Menu"
    Private Function classM_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"

        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго порядка
                Dim res As String = CreateNewObject(classId, funcParams)

                If questEnvironment.EDIT_MODE = False AndAlso funcParams.Count > 1 Then
                    Dim mId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                    If mId > -1 AndAlso GVARS.G_CURMENU = mId Then
                        'выводим на экран новый пункт меню
                        If ShowMenu(mId, funcParams, MenuGetObjId(classId, mId, funcParams)) = "#Error" Then Return "#Error"
                    End If
                End If
                Return res
            Case "Count"
                Return ObjectsCount(classId, funcParams)
            Case "IsExist"
                'Определяет есть ли элемент с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Remove"
                'удаление меню/пункта меню
                If IsNothing(funcParams) = False AndAlso funcParams.Count > 1 AndAlso funcParams(1) = "-1" Then
                    ReDim Preserve funcParams(0)
                End If
                Dim objId As Integer = -1, curMnuName As String = "", mId As Integer = -1
                If questEnvironment.EDIT_MODE = False Then
                    'убираем меню/рункт меню с экрана
                    If IsNothing(funcParams) = False OrElse funcParams.Count = 0 Then
                        If GVARS.G_CURMENU > -1 Then HideMenu(GVARS.G_CURMENU, funcParams)
                    ElseIf funcParams.Count = 1 Then
                        mId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                        If mId > -1 Then ' AndAlso GVARS.G_CURMENU = mId Then
                            HideMenu(mId, funcParams)
                            objId = MenuGetObjId(classId, mId, funcParams)
                            curMnuName = mScript.mainClass(classId).ChildProperties(mId)("Name").Value
                        End If
                    ElseIf funcParams.Count > 1 Then
                        mId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                        If mId > -1 AndAlso GVARS.G_CURMENU = mId Then
                            objId = MenuGetObjId(classId, mId, funcParams)
                            If ShowMenu(mId, funcParams, objId) = "#Error" Then Return "#Error"
                        End If
                    End If
                End If

                Return RemoveObject(classId, funcParams)

                If questEnvironment.EDIT_MODE = False AndAlso curMnuName.Length > 0 Then
                    mId = GetSecondChildIdByName(curMnuName, mScript.mainClass(classId).ChildProperties)
                    If mId > -1 Then ShowMenu(mId, {mId.ToString}, objId)
                End If
            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "CurMenu"
                Return GVARS.G_CURMENU.ToString
            Case "Show"
                Dim mId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If mId < 0 Then Return _ERROR("Меню " & funcParams(0) & " не найдено.", functionName)
                ShowMenu(mId, funcParams, -1)
                Return ""
            Case "Hide"
                Dim mId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If mId < 0 Then Return _ERROR("Меню " & funcParams(0) & " не найдено.", functionName)
                HideMenu(mId, funcParams)
            Case "Select"
                If questEnvironment.EDIT_MODE Then Return "-1"
                Dim mId As Integer = CInt(ObjectId(classId, funcParams, False))
                If mId < 0 Then Return _ERROR("Не найден блок меню " & funcParams(0), functionName)
                Dim itemId As Integer = CInt(ObjectId(classId, funcParams, True))
                If itemId < 0 Then Return _ERROR("Не найден пункт меню " & funcParams(1), functionName)
                Dim runAlways As Boolean = CBool(GetParam(funcParams, 2, "True"))

                If Not runAlways Then
                    'если третий параметр False, то проверяем на доступность/видимость
                    If GVARS.G_CURMENU <> mId Then Return "-1"
                    Dim enabled As Boolean = True
                    ReadPropertyBool(classId, "Enabled", mId, itemId, enabled, funcParams)
                    If enabled = False Then Return "-1"
                    Dim visible As Boolean = True
                    ReadPropertyBool(classId, "Visible", mId, itemId, visible, funcParams)
                    If visible = False Then Return "-1"
                End If

                Dim hItem As HtmlElement = FindMenuItemHTMLelement(classId, mId, itemId, funcParams)
                If IsNothing(hItem) Then
                    Dim hDoc As HtmlDocument = frmPlayer.wbDescription.Document
                    If IsNothing(hDoc) Then Return _ERROR("Не найден html-документ окна описаний.", functionName)
                    hItem = hDoc.CreateElement("SPAN")
                    hItem.Name = mScript.mainClass(classId).ChildProperties(mId)("Name").ThirdLevelProperties(itemId)
                    hItem.SetAttribute("menu", mId.ToString)
                    hItem.SetAttribute("objId", "-1")
                End If

                EventGeneratedFromScript = True
                del_menu_Click(hItem, Nothing)
                Return ""
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function ClassM_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If questEnvironment.EDIT_MODE Then Return ""
        Dim mId As Integer = -1
        If funcParams.Count > 0 Then
            If funcParams(0) = "-1" Then Return ""
            mId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If mId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует имени и Id блока меню {0}!", funcParams(0))
                Return "#Error"
            End If
        Else
            Return ""
        End If
        Dim itemId As Integer = -1
        If funcParams.Count > 1 Then
            itemId = GetThirdChildIdByName(funcParams(1), mId, mScript.mainClass(classId).ChildProperties)
            If itemId < -1 Then Return _ERROR("Не найден пунткт меню " & funcParams(1))
        End If
        'получаем имя действия
        Dim menuName As String = mScript.mainClass(classId).ChildProperties(mId)("Name").Value

        Select Case propertyName
            Case "Caption", "Picture", "PictureFloat"
                If GVARS.G_CURMENU <> mId Then Return ""
                If itemId >= 0 Then
                    'изменяем вид html-элемента пункта меню
                    Dim hItem As HtmlElement = FindMenuItemHTMLelement(classId, mId, itemId, funcParams)
                    If IsNothing(hItem) Then Return ""
                    Dim itemType As Integer, itemCaption As String = "", itemPicture As String = "", itemPictureFloat As Integer
                    If ReadPropertyInt(classId, "Type", mId, itemId, itemType, funcParams) = False Then Return "#Error"
                    If ReadProperty(classId, "Caption", mId, itemId, itemCaption, funcParams) = False Then Return "#Error"
                    If ReadProperty(classId, "Picture", mId, itemId, itemPicture, funcParams) = False Then Return "#Error"
                    If ReadPropertyInt(classId, "PictureFloat", mId, itemId, itemPictureFloat, funcParams) = False Then Return "#Error"
                    If itemType = 1 Then
                        hItem.SetAttribute("Title", UnWrapString(itemCaption))
                        hItem.SetAttribute("src", UnWrapString(itemPicture).Replace("\", "/"))
                    Else
                        MenuSetInnerHTML(hItem, itemType, UnWrapString(itemCaption), UnWrapString(itemPicture).Replace("\", "/"), itemPictureFloat)
                    End If
                ElseIf propertyName = "Caption" Then
                    Dim hCont As HtmlElement = MenuGetContainer(classId, mId, funcParams)
                    If IsNothing(hCont) OrElse hCont.Children.Count = 0 Then Return ""
                    Dim hTitle As HtmlElement = hCont.Children(0)
                    Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                    If String.IsNullOrEmpty(nValue) Then
                        hTitle.InnerHtml = ""
                        hTitle.Style = "display:none"
                    Else
                        hTitle.InnerHtml = nValue
                        hTitle.Style = ""
                    End If
                End If

                Return ""
            Case "HTMLContainerId"
                If itemId >= 0 Then Return ""
                If GVARS.G_CURMENU = -1 OrElse newValue = PREV_VALUE Then Return ""
                Dim objId As Integer = MenuGetObjId(classId, mId, funcParams)

                mScript.mainClass(classId).ChildProperties(mId)("HTMLContainerId").Value = PREV_VALUE
                HideMenu(mId, funcParams)
                mScript.mainClass(classId).ChildProperties(mId)("HTMLContainerId").Value = newValue

                Return ShowMenu(mId, arrParams, objId)
            Case "Type"
                If itemId < 0 OrElse GVARS.G_CURMENU <> mId Then Return ""

                Dim hItem As HtmlElement = FindMenuItemHTMLelement(classId, mId, itemId, funcParams)
                If IsNothing(hItem) Then Return ""
                Dim hPar As HtmlElement = hItem.Parent
                If IsNothing(hPar) Then Return ""
                Dim objId As Integer = Val(hItem.GetAttribute("objId"))
                MenuCreateHtmlElement(mId, itemId, objId, hPar, hItem)

                Dim msItem As mshtml.IHTMLDOMNode = hItem.DomElement
                msItem.removeNode(True)
                Return ""
            Case "Enabled"
                If itemId < 0 OrElse GVARS.G_CURMENU <> mId Then Return ""
                Dim enabled As Boolean = True
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                Boolean.TryParse(nValue, enabled)

                'делаем пункт меню недоступным
                Dim hItem As HtmlElement = FindMenuItemHTMLelement(classId, mId, itemId, funcParams)
                If IsNothing(hItem) Then Return "#Error"
                If enabled Then
                    hItem.DomElement.removeAttribute("disabled")
                Else
                    hItem.SetAttribute("disabled", "True")
                End If
                Return ""
            Case "Visible"
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                If itemId < 0 Then
                    If nValue = "True" Then
                        'отображаем блок меню
                        Return ShowMenu(mId, funcParams, -1)
                    Else
                        'прячем блок меню
                        Return HideMenu(mId, funcParams)
                    End If
                Else
                    If GVARS.G_CURMENU <> mId Then Return ""
                    'делаем пункт меню видимым/невидимым
                    Dim hItem As HtmlElement = FindMenuItemHTMLelement(classId, mId, itemId, funcParams)
                    If IsNothing(hItem) Then Return "#Error"
                    Dim visible As Boolean = True
                    Boolean.TryParse(nValue, visible)
                    If visible Then
                        hItem.Style = ""
                    Else
                        hItem.Style = "display:none"
                    End If
                End If
                Return ""
            Case "CancelButtonText", "CancelButtonPicture", "CancelButtonPictureFloat"
                If itemId >= 0 AndAlso GVARS.G_CURMENU <> mId Then Return ""

                'добавляется ли кнопка отмены автоматически
                Dim AutoCancelButton As Integer = 0
                If ReadPropertyInt(classId, "AutoCancelButton", mId, -1, AutoCancelButton, funcParams) = False Then Return "#Error"
                If AutoCancelButton = 0 Then Return "" 'если нет - не продолжаем

                'изменяем вид html-элемента пункта меню отмены
                Dim hItem As HtmlElement = FindMenuItemHTMLelement(classId, mId, -1, funcParams)
                If IsNothing(hItem) Then Return ""
                Dim itemType As Integer, itemCaption As String = "", itemPicture As String = "", itemPictureFloat As Integer
                If ReadPropertyInt(classId, "CancelButtonType", mId, -1, itemType, funcParams) = False Then Return "#Error"
                If ReadProperty(classId, "CancelButtonText", mId, -1, itemCaption, funcParams) = False Then Return "#Error"
                If ReadProperty(classId, "CancelButtonPicture", mId, -1, itemPicture, funcParams) = False Then Return "#Error"
                If ReadPropertyInt(classId, "CancelButtonPictureFloat", mId, -1, itemPictureFloat, funcParams) = False Then Return "#Error"
                If itemType = 1 Then
                    hItem.SetAttribute("Title", UnWrapString(itemCaption))
                    hItem.SetAttribute("src", UnWrapString(itemPicture).Replace("\", "/"))
                Else
                    MenuSetInnerHTML(hItem, itemType, UnWrapString(itemCaption), UnWrapString(itemPicture).Replace("\", "/"), itemPictureFloat)
                End If

                Return ""
            Case "CancelButtonType"
                If itemId >= 0 OrElse GVARS.G_CURMENU <> mId Then Return ""

                'добавляется ли кнопка отмены автоматически
                Dim AutoCancelButton As Integer = 0
                If ReadPropertyInt(classId, "AutoCancelButton", mId, -1, AutoCancelButton, funcParams) = False Then Return "#Error"
                If AutoCancelButton = 0 Then Return "" 'если нет - не продолжаем

                Dim hItem As HtmlElement = FindMenuItemHTMLelement(classId, mId, -1, funcParams)
                If IsNothing(hItem) Then Return ""
                Dim hPar As HtmlElement = hItem.Parent
                If IsNothing(hPar) Then Return ""
                Dim objId As Integer = Val(hItem.GetAttribute("objId"))


                Dim hNew As HtmlElement = MenuCreateHtmlCancelElement(mId, AutoCancelButton = 2, hItem.Parent, funcParams, objId)
                If IsNothing(hNew) Then Return "#Error"

                Dim msItem As mshtml.IHTMLDOMNode = hItem.DomElement
                msItem.removeNode(True)
                Return ""

        End Select
        Return ""
    End Function

    ''' <summary>
    ''' Возвращает Id предмета  отображенного блока меню
    ''' </summary>
    ''' <param name="classId">Id класса Menu</param>
    ''' <param name="mId">Id блока меню</param>
    ''' <param name="funcParams"></param>
    Private Function MenuGetObjId(ByVal classId As Integer, ByVal mId As Integer, ByRef funcParams() As String) As Integer
        If IsNothing(mScript.mainClass(classId).ChildProperties(mId)("Name").ThirdLevelProperties) OrElse mScript.mainClass(classId).ChildProperties(mId)("Name").ThirdLevelProperties.Count = 0 Then Return -1
        Dim hEl As HtmlElement = FindMenuItemHTMLelement(classId, mId, 0, funcParams)
        If IsNothing(hEl) Then Return -1
        Dim objId As String = hEl.GetAttribute("objId")
        If String.IsNullOrEmpty(objId) Then
            Return -1
        Else
            Return Val(objId)
        End If
    End Function
    ''' <summary>
    ''' Возвращает Id контейнера блока меню
    ''' </summary>
    ''' <param name="classId">Id класса Menu</param>
    ''' <param name="mId">Id блока меню</param>
    ''' <param name="funcParams"></param>
    Private Function MenuGetContainer(ByVal classId As Integer, ByVal mId As Integer, ByRef funcParams() As String) As HtmlElement
        Dim mContainer As String = ""
        ReadProperty(classId, "HTMLContainerId", mId, -1, mContainer, funcParams)
        mContainer = UnWrapString(mContainer)
        If String.IsNullOrEmpty(mContainer) Then
            Return frmPlayer.wbDescription.Document.GetElementById("MenuConvas")
        Else
            Return frmPlayer.wbMain.Document.GetElementById(mContainer)
        End If
    End Function

    ''' <summary>
    ''' Получает html-элемент указанного пункта меню
    ''' </summary>
    ''' <param name="classId">класса А</param>
    ''' <param name="mId">Id блока меню</param>
    ''' <param name="itemId">Id пункта меню. -1 если надо получить html-элемент кнопки отмены</param>
    ''' <param name="funcParams"></param>
    ''' <param name="mContainer">html-контейнер</param>
    Private Function FindMenuItemHTMLelement(ByVal classId As Integer, ByVal mId As Integer, ByVal itemId As Integer, ByRef funcParams() As String, Optional ByVal mContainer As String = Nothing) As HtmlElement
        'получаем Id контейнера
        If questEnvironment.EDIT_MODE Then Return Nothing
        If IsNothing(mContainer) Then
            ReadProperty(classId, "HTMLContainerId", mId, -1, mContainer, funcParams)
            mContainer = UnWrapString(mContainer)
        End If
        Dim hContainer As HtmlElement
        If String.IsNullOrEmpty(mContainer) Then
            hContainer = frmPlayer.wbDescription.Document.GetElementById("MenuConvas")
        Else
            hContainer = frmPlayer.wbMain.Document.GetElementById(mContainer)
        End If
        Dim mName As String = mScript.mainClass(classId).ChildProperties(mId)("Name").Value
        If IsNothing(hContainer) Then
            mScript.LAST_ERROR = String.Format("Не удалось найти html-контейнер {0} для размещения в нем блока меню {1}.", mContainer, mName)
            Return Nothing
        End If
        'получаем html-элемент с нашим пунктом меню
        Dim hItem As HtmlElement = Nothing

        If itemId = -1 Then
            'нужна кнопка отмены
            Dim hEl As HtmlElement = hContainer.Children(hContainer.Children.Count - 1)
            If hEl.Name <> "sys_AutoCancelButton" Then Return Nothing
            Return hEl
        End If

        For i As Integer = 0 To hContainer.Children.Count - 1
            Dim hEl As HtmlElement = hContainer.Children(i)
            Dim cmId As Integer = Val(hEl.GetAttribute("menu"))
            Dim citemId As Integer = GetThirdChildIdByName(hEl.Name, cmId, mScript.mainClass(classId).ChildProperties)

            If mId = cmId AndAlso itemId = citemId Then
                'найден
                hItem = hEl
                Exit For
            End If
        Next
        If IsNothing(hItem) Then
            mScript.LAST_ERROR = String.Format("Не удалось найти пункт меню {0} на экране.", mName)
            Return Nothing
        End If
        Return hItem
    End Function

#End Region

#Region "Media"
    Private Function classMed_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго порядка
                Return CreateNewObject(classId, funcParams)                
            Case "Count"
                Return ObjectsCount(classId, funcParams)
            Case "IsExist"
                'Определяет есть ли элемент с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "Remove"
                'удаление элемента
                If questEnvironment.EDIT_MODE = False AndAlso GVARS.G_CURLIST >= 0 Then
                    Dim aListName As String = GetParam(funcParams, 0, "-1")
                    Dim aListId As Integer = GetSecondChildIdByName(aListName, mScript.mainClass(classId).ChildProperties)
                    Dim aAudioId As Integer = -1
                    If aListId >= 0 Then
                        Dim aAudioName As String = GetParam(funcParams, 1, "-1")
                        aAudioId = GetThirdChildIdByName(aAudioName, aListId, mScript.mainClass(classId).ChildProperties)
                    End If

                    'останавливаем воспроизведение, если аудио  будет удалено
                    Dim curAudio As Integer = GVARS.G_CURAUDIO
                    If aListId < 0 OrElse GVARS.G_CURLIST = aListId Then
                        If curAudio > -1 AndAlso (aAudioId < 0 OrElse aAudioId = curAudio) Then
                            If AudioStop() = False Then Return "#Error"
                        End If
                    End If

                    If aListId >= 0 AndAlso aAudioId < 0 Then
                        'удаляется список. Изменяем если надо GVARS.G_CURLIST
                        If aListId > GVARS.G_CURLIST Then GVARS.G_CURLIST -= 1
                    ElseIf aListId >= 0 AndAlso aAudioId >= 0 Then
                        'удаляется аудио. Изменяем если надо GVARS.G_CURAUDIO
                        If aAudioId > curAudio Then GVARS.G_CURAUDIO -= 1
                    End If
                End If

                Return RemoveObject(classId, funcParams)
            Case "CurAudio"
                Return GVARS.G_CURAUDIO.ToString
            Case "CurList"
                Return GVARS.G_CURLIST.ToString
            Case "GetCurrentPosition"
                If GVARS.G_CURAUDIO < 0 Then Return "0"
                Dim curPos As Double = mPlayer.CurrentPosition
                Return curPos.ToString(provider_points)
            Case "SetCurrentPosition"
                Dim curPos As Double = Double.Parse(funcParams(0), Globalization.NumberStyles.Any, provider_points)
                If curPos < 0 Then curPos = 0
                mPlayer.CurrentPosition = curPos
                Return ""
            Case "GetDuration"
                If GVARS.G_CURAUDIO < 0 Then Return "0"
                Return mPlayer.Duration.ToString(provider_points)
            Case "PlayState"
                If GVARS.G_CURAUDIO < 0 Then Return mScript.mainClass(classId).Functions(functionName).returnArray(0)
                Dim enumId As Integer = 0
                If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPlaying Then
                    enumId = 1
                ElseIf mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPaused Then
                    enumId = 2
                End If
                Return mScript.mainClass(classId).Functions(functionName).returnArray(enumId)
            Case "Pause"
                If GVARS.G_CURAUDIO < 0 Then Return ""
                If mPlayer.PlayState <> MediaPlayer.MPPlayStateConstants.mpPaused AndAlso mPlayer.PlayState <> MediaPlayer.MPPlayStateConstants.mpPlaying Then Return ""
                Dim duration As Integer = Val(GetParam(funcParams, 0, "0"))
                If duration < 0 Then duration = 0
                Dim shouldWait As Boolean = CBool(GetParam(funcParams, 1, "False"))
                Dim volume As Integer = 0
                If ReadPropertyInt(classId, "Volume", GVARS.G_CURLIST, GVARS.G_CURAUDIO, volume, arrParams) = False Then Return "#Error"

                If duration = 0 Then
                    If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPlaying Then
                        mPlayer.Pause()
                    Else
                        mPlayer.Volume = (100 - volume) * -41
                        mPlayer.Play()
                    End If
                ElseIf shouldWait = True Then
                    If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPlaying Then
                        ShiftVolumeSynch(0, duration)
                        mPlayer.Pause()
                    Else
                        mPlayer.Volume = -4100
                        mPlayer.Play()
                        ShiftVolumeSynch(volume, duration)
                    End If
                Else
                    If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPlaying Then
                        ShiftVolumeAsynch(0, duration)
                        mPlayer.Pause()
                    Else
                        mPlayer.Volume = -4100
                        mPlayer.Play()
                        ShiftVolumeAsynch(volume, duration)
                    End If
                End If

                Return ""
            Case "Say"
                Dim fPath As String = UnWrapString(funcParams(0))
                Dim volume As Integer = Val(GetParam(funcParams, 1, "100"))
                Dim shouldWait As Boolean = CBool(GetParam(funcParams, 2, "False"))
                Return Say(fPath, volume, shouldWait)
            Case "SelectList"
                AudioStop()
                Dim lstId As Integer = -1
                If funcParams(0) <> "-1" Then
                    lstId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                    If lstId < 0 Then Return _ERROR("Список воспроизведения " & funcParams(0) & "Не найден.", functionName)
                End If

                GVARS.G_CURLIST = lstId
                Return ""
            Case "ShiftVolume"
                If GVARS.G_CURAUDIO < 0 Then Return ""
                Dim finalVolume As Integer = Val(funcParams(0))
                If finalVolume < 0 Then
                    finalVolume = 0
                ElseIf finalVolume > 100 Then
                    finalVolume = 100
                End If
                Dim duration As Integer = Val(GetParam(funcParams, 1, "1000"))
                If duration < 0 Then duration = 1000
                Dim shouldWait As Boolean = CBool(GetParam(funcParams, 2, "False"))

                If shouldWait Then
                    ShiftVolumeSynch(finalVolume, duration)
                    If PropertiesRouter(classId, "Volume", {GVARS.G_CURLIST.ToString, GVARS.G_CURAUDIO.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, funcParams(0)) = "#Error" Then Return "#Error"
                Else
                    ShiftVolumeAsynch(finalVolume, duration)
                    PropertiesRouter(classId, "Volume", {GVARS.G_CURLIST.ToString, GVARS.G_CURAUDIO.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, finalVolume.ToString)
                End If
                Return ""
            Case "StartPlay"
                If GVARS.G_CURLIST < 0 Then Return _ERROR("Список воспроизведения не установлен. Воспользуйтесь функцией Med.SelectList", functionName)
                If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPaused OrElse mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPlaying Then
                    If AudioStop() = False Then Return "#Error"
                End If
                Dim res As String = AudioPlayFromList(-1, Nothing)
                If res = "#Error" Then Return res
                timAudio.Enabled = True
                Return ""
            Case "StopPlay"
                If GVARS.G_CURAUDIO < 0 Then Return ""
                Dim duration As Integer = Val(GetParam(funcParams, 0, "0"))
                If duration < 0 Then duration = 0
                Dim shouldWait As Boolean = CBool(GetParam(funcParams, 1, "False"))

                If duration = 0 Then
                    If AudioStop() = False Then Return "#Error"
                ElseIf shouldWait = True Then
                    ShiftVolumeSynch(0, duration)
                    If AudioStop() = False Then Return "#Error"
                Else
                    ShiftVolumeAsynch(0, duration)
                    If AudioStop() = False Then Return "#Error"
                End If

                Return ""
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function ClassMed_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If questEnvironment.EDIT_MODE OrElse GVARS.G_CURAUDIO < 0 OrElse GVARS.G_CURLIST < 0 Then Return ""
        Dim listId As Integer = -1, audioId As Integer = -1
        If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then
            If funcParams(0) = "-1" Then Return ""
            listId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If listId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует списка воспроизведения {0}!", funcParams(0))
                Return "#Error"
            End If
            If listId <> GVARS.G_CURLIST Then Return ""

            If funcParams.Count > 1 Then
                audioId = GetThirdChildIdByName(funcParams(1), listId, mScript.mainClass(classId).ChildProperties)
                If audioId < 0 Then Return _ERROR(String.Format("Не существует аудиофайла {0} в списке воспроизведения {1}!", funcParams(1), funcParams(0)))
                If audioId <> GVARS.G_CURAUDIO Then Return ""
            Else
                Return ""
            End If
        Else
            Return ""
        End If
        'Здесь только если меняется свойство элемента 3 уровня (аудиофайла), который сейчас проигрывается

        ''получаем имя действия
        'Dim menuName As String = mScript.mainClass(classId).ChildProperties(listId)("Name").Value
        Select Case propertyName
            Case "Volume"
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                Dim volume As Integer = Val(nValue)
                If volume < 0 Then
                    volume = 0
                ElseIf volume > 100 Then
                    volume = 100
                End If
                mPlayer.Volume = (100 - volume) * -41
        End Select
        Return ""
    End Function

#End Region

#Region "Timer"
    Private Function classT_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Add"
                'Создает в указанном классе объект второго порядка
                Dim strId As String = CreateNewObject(classId, funcParams)
                If strId = "#Error" Then Return strId
                If questEnvironment.EDIT_MODE Then Return strId

                Dim tId As Integer = Val(strId)
                Dim interval As Integer = Val(GetParam(funcParams, 1, "1000"))
                Dim enabled As Boolean = CBool(GetParam(funcParams, 2, "True"))

                TimerCreate(UnWrapString(funcParams(0)), interval, enabled)
                Return strId
            Case "Count"
                Return ObjectsCount(classId, funcParams, False)
            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "IsExist"
                'Определяет есть ли элемент с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Remove"
                'удаление элемента
                If questEnvironment.EDIT_MODE = False Then
                    Dim tId As Integer = Val(GetParam(funcParams, 0, "-1"))
                    TimerRemove(tId)
                End If

                'TimerRemove 
                Return RemoveObject(classId, funcParams)
            Case "HoldUpAll"
                GVARS.HOLD_UP_TIMERS = True
                Return ""
            Case "ReleaseAll"
                GVARS.HOLD_UP_TIMERS = False
                Return ""
            Case "Begin"
                Return PropertiesRouter(classId, "IsWorking", funcParams, arrParams, PropertiesOperationEnum.PROPERTY_SET, "True")
            Case "Stop"
                Return PropertiesRouter(classId, "IsWorking", funcParams, arrParams, PropertiesOperationEnum.PROPERTY_SET, "True")
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function ClassT_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If questEnvironment.EDIT_MODE Then Return ""
        Dim tId As Integer = -1
        If funcParams.Count > 0 Then
            If funcParams(0) = "-1" Then Return ""
            tId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If tId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует имени и Id счетчика {0}!", funcParams(0))
                Return "#Error"
            End If
        Else
            Return ""
        End If
        'получаем имя счетчика
        Dim tName As String = mScript.mainClass(classId).ChildProperties(tId)("Name").Value

        Select Case propertyName
            Case "Interval"
                'изменяем интервал
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                Dim interval As Integer = Val(nValue)
                If interval <= 0 Then interval = 1
                lstTimers(UnWrapString(tName)).Interval = interval
                Return ""
            Case "IsWorking"
                'делаем счетчик доступным / недоступным
                Dim enabled As Boolean = True
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                Boolean.TryParse(nValue, enabled)
                lstTimers(UnWrapString(tName)).Enabled = enabled
                Return ""
        End Select
        Return ""
    End Function
#End Region

#Region "Map"
    Private Function classMap_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"

        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго порядка
                Dim res As String = CreateNewObject(classId, funcParams)
                Return res
            Case "Count"
                Return ObjectsCount(classId, funcParams)
            Case "IsExist"
                'Определяет есть ли элемент с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Remove"
                'удаление карты/клетки 
                Dim shouldRebuild As Boolean = False
                Dim mapId As Integer = -1, cellId As Integer = -1
                If questEnvironment.EDIT_MODE = False Then
                    If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 AndAlso funcParams(0) <> "-1" Then
                        mapId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                        If mapId < 0 Then Return _ERROR("Карта " & funcParams(0) & " не найдена.", functionName)

                        If funcParams.Count > 1 Then
                            cellId = GetThirdChildIdByName(funcParams(1), mapId, mScript.mainClass(classId).ChildProperties)
                            If cellId < 0 Then Return _ERROR("Клетки карты " & funcParams(1) & " у карты " & funcParams(0) & " не найдено.", functionName)
                        End If
                    End If

                    If mapId < 0 Then
                        'удаление всех карт
                        If GVARS.G_CURMAP > -1 Then
                            'прячем карту
                            Dim slot As Integer = 0, isNewWindow As Boolean = False
                            If ReadPropertyInt(classId, "Slot", -1, -1, slot, funcParams) = False Then Return "#Error"
                            If slot = 5 Then isNewWindow = True 'в отдельном окне

                            If isNewWindow Then
                                frmMap.Hide()
                            Else
                                questEnvironment.wbMap.Hide()
                            End If

                            'убираем текущую карту
                            GVARS.G_CURMAP = -1
                            GVARS.G_CURMAPCELL = -1
                            GVARS.G_PREVMAPCELL = -1
                            mapManager.ClearPreviousBitmap()
                        End If
                    ElseIf cellId < 0 Then
                        'удаление карты
                        If GVARS.G_CURMAP = mapId Then
                            'удаляется текущая карта
                            'прячем карту
                            Dim slot As Integer = 0, isNewWindow As Boolean = False
                            If ReadPropertyInt(classId, "Slot", -1, -1, slot, funcParams) = False Then Return "#Error"
                            If slot = 5 Then isNewWindow = True 'в отдельном окне

                            If isNewWindow Then
                                frmMap.Hide()
                            Else
                                questEnvironment.wbMap.Hide()
                            End If

                            'убираем текущую карту
                            GVARS.G_CURMAP = -1
                            GVARS.G_CURMAPCELL = -1
                            GVARS.G_PREVMAPCELL = -1
                            mapManager.ClearPreviousBitmap()
                        ElseIf mapId > GVARS.G_CURMAP Then
                            'смещаем индекс текущей карты на -1
                            GVARS.G_CURMAP -= 1
                        End If
                    Else
                        'удаление клетки
                        If GVARS.G_CURMAP = mapId Then
                            '... на текущей карте
                            If GVARS.G_PREVMAPCELL = cellId Then
                                GVARS.G_PREVMAPCELL = -1
                            ElseIf GVARS.G_CURMAPCELL = cellId Then
                                GVARS.G_CURMAPCELL = -1
                            End If
                        End If
                        shouldRebuild = True
                    End If
                End If

                Dim res As String = RemoveObject(classId, funcParams)
                If res = "#Error" Then Return res
                If shouldRebuild Then mapManager.BuildMap(mapId, cellId, funcParams, GVARS.G_PREVMAPCELL)
                Return res
            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "Show"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim mapId As Integer = GetSecondChildIdByName(GetParam(funcParams, 0, GVARS.G_CURMAP.ToString), mScript.mainClass(classId).ChildProperties)
                If mapId < 0 Then
                    Return _ERROR("Карты " & funcParams(0) & " не найдено.", functionName)
                End If

                If GVARS.G_CURMAP <> mapId Then
                    'смена текущей карты
                    GVARS.G_CURMAP = mapId
                    If GVARS.G_CURLOC > -1 Then
                        GVARS.G_CURMAPCELL = mapManager.GetCellByLocation(mapId, GVARS.G_CURLOC)
                    Else
                        GVARS.G_CURMAPCELL = -1
                    End If
                    GVARS.G_PREVMAPCELL = -1
                    If mapManager.BuildMap(mapId, GVARS.G_CURMAPCELL, funcParams) = "#Error" Then Return "#Error"
                End If

                Dim slot As Integer = 0, isNewWindow As Boolean = False
                If ReadPropertyInt(classId, "Slot", -1, -1, slot, funcParams) = False Then Return "#Error"
                If slot = 5 Then isNewWindow = True 'в отдельном окне

                If isNewWindow Then
                    questEnvironment.wbMap.Show()
                    frmMap.Show()
                Else
                    questEnvironment.wbMap.BringToFront()
                    questEnvironment.wbMap.Show()
                End If
                Return ""
            Case "Hide"
                Dim slot As Integer = 0, isNewWindow As Boolean = False
                If ReadPropertyInt(classId, "Slot", -1, -1, slot, funcParams) = False Then Return "#Error"
                If slot = 5 Then isNewWindow = True 'в отдельном окне

                If isNewWindow Then
                    frmMap.Hide()
                Else
                    questEnvironment.wbMap.Hide()
                End If
                mapManager.ClearPreviousBitmap()
            Case "SetCurrentMap"
                If funcParams(0) = "-1" Then
                    GVARS.G_CURMAP = -1
                    GVARS.G_CURMAPCELL = -1
                    GVARS.G_PREVMAPCELL = -1
                    Return ""
                End If
                mapManager.ClearPreviousBitmap()

                Dim mapId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If mapId < 0 Then Return _ERROR("Карта " & funcParams(0) & " не найдена.", functionName)
                GVARS.G_CURMAP = mapId
                If GVARS.G_CURLOC > -1 Then
                    GVARS.G_CURMAPCELL = mapManager.GetCellByLocation(mapId, GVARS.G_CURLOC)
                Else
                    GVARS.G_CURMAPCELL = -1
                End If
                GVARS.G_PREVMAPCELL = -1

                Return mapManager.BuildMap(mapId, GVARS.G_CURMAPCELL, funcParams)
            Case "SetCurrentCell"
                If funcParams(0) = "-1" Then
                    GVARS.G_CURMAP = -1
                    GVARS.G_CURMAPCELL = -1
                    GVARS.G_PREVMAPCELL = -1
                    Return ""
                End If
                Dim mapId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If mapId < 0 Then Return _ERROR("Карта " & funcParams(0) & " не найдена.", functionName)

                Dim cellId As Integer = GetThirdChildIdByName(funcParams(1), mapId, mScript.mainClass(classId).ChildProperties)
                If cellId < 0 Then Return _ERROR("Клетки карты " & funcParams(1) & " у карты " & funcParams(0) & " не найдено.", functionName)

                If mapId <> GVARS.G_CURMAP Then
                    mapManager.ClearPreviousBitmap()

                    GVARS.G_CURMAP = mapId                    
                    GVARS.G_PREVMAPCELL = -1
                    GVARS.G_CURMAPCELL = cellId
                    Return mapManager.BuildMap(mapId, GVARS.G_CURMAPCELL, funcParams)
                Else
                    GVARS.G_PREVMAPCELL = GVARS.G_CURMAPCELL
                    GVARS.G_CURMAPCELL = cellId
                    mapManager.ClearNearbyFog(funcParams)
                    mapManager.ShowArrowInCurrentCell()
                    Return ""
                End If
            Case "CurMap"
                Return GVARS.G_CURMAP.ToString
            Case "CurCell"
                Return GVARS.G_CURMAPCELL
            Case "PrevCell"
                Return GVARS.G_PREVMAPCELL
            Case "BackCell"
                If GVARS.G_CURMAP < 0 OrElse GVARS.G_CURMAPCELL < 0 Then Return "-1"
                Dim x As Integer = 0, y As Integer = 0
                If ReadPropertyInt(classId, "CellX", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, x, funcParams) = False Then Return "#Error"
                If ReadPropertyInt(classId, "CellY", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, y, funcParams) = False Then Return "#Error"
                Dim direct As cMapManager.DirectionEnum = mapManager.GetDirection

                Select Case direct
                    Case cMapManager.DirectionEnum.LEFT
                        '<--: -->
                        x += 1
                    Case cMapManager.DirectionEnum.RIGHT
                        x -= 1
                    Case cMapManager.DirectionEnum.FORWARD
                        y += 1
                    Case cMapManager.DirectionEnum.BACKWARD
                        y -= 1
                End Select
                Dim cellId As Integer = mapManager.GetCellIdByXY(GVARS.G_CURMAP, x, y)
                Return cellId.ToString
            Case "FrontCell"
                If GVARS.G_CURMAP < 0 OrElse GVARS.G_CURMAPCELL < 0 Then Return "-1"
                Dim x As Integer = 0, y As Integer = 0
                If ReadPropertyInt(classId, "CellX", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, x, funcParams) = False Then Return "#Error"
                If ReadPropertyInt(classId, "CellY", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, y, funcParams) = False Then Return "#Error"
                Dim direct As cMapManager.DirectionEnum = mapManager.GetDirection

                Select Case direct
                    Case cMapManager.DirectionEnum.LEFT
                        x -= 1
                    Case cMapManager.DirectionEnum.RIGHT
                        x += 1
                    Case cMapManager.DirectionEnum.FORWARD
                        y -= 1
                    Case cMapManager.DirectionEnum.BACKWARD
                        y += 1
                End Select
                Dim cellId As Integer = mapManager.GetCellIdByXY(GVARS.G_CURMAP, x, y)
                Return cellId.ToString
            Case "LeftCell"
                If GVARS.G_CURMAP < 0 OrElse GVARS.G_CURMAPCELL < 0 Then Return "-1"
                Dim x As Integer = 0, y As Integer = 0
                If ReadPropertyInt(classId, "CellX", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, x, funcParams) = False Then Return "#Error"
                If ReadPropertyInt(classId, "CellY", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, y, funcParams) = False Then Return "#Error"
                Dim direct As cMapManager.DirectionEnum = mapManager.GetDirection

                Select Case direct
                    Case cMapManager.DirectionEnum.LEFT
                        y += 1
                    Case cMapManager.DirectionEnum.RIGHT
                        y -= 1
                    Case cMapManager.DirectionEnum.FORWARD
                        x -= 1
                    Case cMapManager.DirectionEnum.BACKWARD
                        x += 1
                End Select
                Dim cellId As Integer = mapManager.GetCellIdByXY(GVARS.G_CURMAP, x, y)
                Return cellId.ToString
            Case "RightCell"
                If GVARS.G_CURMAP < 0 OrElse GVARS.G_CURMAPCELL < 0 Then Return "-1"
                Dim x As Integer = 0, y As Integer = 0
                If ReadPropertyInt(classId, "CellX", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, x, funcParams) = False Then Return "#Error"
                If ReadPropertyInt(classId, "CellY", GVARS.G_CURMAP, GVARS.G_CURMAPCELL, y, funcParams) = False Then Return "#Error"
                Dim direct As cMapManager.DirectionEnum = mapManager.GetDirection

                Select Case direct
                    Case cMapManager.DirectionEnum.LEFT
                        y -= 1
                    Case cMapManager.DirectionEnum.RIGHT
                        y += 1
                    Case cMapManager.DirectionEnum.FORWARD
                        x += 1
                    Case cMapManager.DirectionEnum.BACKWARD
                        x -= 1
                End Select
                Dim cellId As Integer = mapManager.GetCellIdByXY(GVARS.G_CURMAP, x, y)
                Return cellId.ToString
            Case "GetCellByLocation"
                Dim locId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(mScript.mainClassHash("L")).ChildProperties)
                If locId < 0 Then Return _ERROR("Локации " & funcParams(0) & " не найдено.", functionName)
                Dim mapId As Integer = GVARS.G_CURMAP
                If funcParams.Count > 1 Then
                    mapId = GetSecondChildIdByName(funcParams(1), mScript.mainClass(classId).ChildProperties)
                    If mapId < 0 Then Return _ERROR("Карты " & funcParams(1) & " не найдено.", functionName)
                ElseIf mapId < 0 Then
                    Return _ERROR("Текущая карта не выбрана.", functionName)
                End If
                Return mapManager.GetCellByLocation(mapId, locId).ToString
            Case "GetCellByXY"
                Dim x As Integer = 0, y As Integer = 0
                x = Val(funcParams(0))
                y = Val(funcParams(1))
                Dim mapId As Integer = GVARS.G_CURMAP
                If funcParams.Count > 1 Then
                    mapId = GetSecondChildIdByName(funcParams(2), mScript.mainClass(classId).ChildProperties)
                    If mapId < 0 Then Return _ERROR("Карты " & funcParams(2) & " не найдено.", functionName)
                ElseIf mapId < 0 Then
                    Return _ERROR("Текущая карта не выбрана.", functionName)
                End If
                Return mapManager.GetCellIdByXY(mapId, x, y).ToString
            Case "IsCellVisible"
                Dim mapId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If mapId < 0 Then Return _ERROR("Карты " & funcParams(0) & " не найдено.", functionName)
                Dim cellId As Integer = GetThirdChildIdByName(funcParams(1), mapId, mScript.mainClass(classId).ChildProperties)
                If cellId < 0 Then Return _ERROR("Клетки карты " & funcParams(1) & " у карты " & funcParams(0) & " не найдено.", functionName)
                Return mapManager.IsCellVisible(mapId, cellId)
            Case "GoCell"
                If questEnvironment.EDIT_MODE = True Then Return ""
                Dim mapId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If mapId < 0 Then Return _ERROR("Карты " & funcParams(0) & " не найдено.", functionName)
                Dim cellId As Integer = GetThirdChildIdByName(funcParams(1), mapId, mScript.mainClass(classId).ChildProperties)
                If cellId < 0 Then Return _ERROR("Клетки карты " & funcParams(1) & " у карты " & funcParams(0) & " не найдено.", functionName)
                Dim hCell As HtmlElement = mapManager.GetCellHtmlElement(cellId, Nothing)
                If IsNothing(hCell) Then Return _ERROR("Ошибка в структуре html-документа карты. Клетка " & funcParams(1) & " не найдена.", functionName)
                Dim prevId As Integer = -2
                If mapId <> GVARS.G_CURMAP Then
                    prevId = GVARS.G_CURMAP
                    classMap_functions(classId, "SetCurrentMap", {mapId.ToString}, arrParams)
                End If

                EventGeneratedFromScript = True
                mapManager.del_CellClick(hCell, Nothing)

                If prevId = -2 Then
                    classMap_functions(classId, "SetCurrentMap", {prevId.ToString}, arrParams)
                End If
                Return ""
            Case "StopWhileBeClosed"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim blnHoldTimers As Boolean = GVARS.HOLD_UP_TIMERS
                GVARS.HOLD_UP_TIMERS = True
                Do While questEnvironment.wbMap.Visible
                    Application.DoEvents()
                Loop
                GVARS.HOLD_UP_TIMERS = blnHoldTimers
                Return ""
            Case "BuildPath"
                If questEnvironment.EDIT_MODE = True Then Return ""
                Dim mapId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If mapId < 0 Then Return _ERROR("Карты " & funcParams(0) & " не найдено.", functionName)
                If mapId <> GVARS.G_CURMAP Then
                    classMap_functions(classId, "SetCurrentMap", {mapId.ToString}, arrParams)
                End If

                Dim cellStart As Integer = GetThirdChildIdByName(funcParams(1), mapId, mScript.mainClass(classId).ChildProperties)
                If cellStart < 0 Then Return _ERROR("Клетки карты " & funcParams(1) & " у карты " & funcParams(0) & " не найдено.", functionName)
                Dim cellFinal As Integer = GetThirdChildIdByName(funcParams(2), mapId, mScript.mainClass(classId).ChildProperties)
                If cellFinal < 0 Then Return _ERROR("Клетки карты " & funcParams(2) & " у карты " & funcParams(0) & " не найдено.", functionName)
                Dim useDiagonals As Boolean = CBool(GetParam(funcParams, 3, "False"))
                Dim buildType As cMapManager.BuildStyleEnum = Val(UnWrapString(GetParam(funcParams, 4, "1")))
                Dim doEvents As Boolean = CBool(GetParam(funcParams, 5, "True"))

                Dim arrPath() As String = mapManager.BuildPath(cellStart, cellFinal, useDiagonals, buildType, doEvents)

                Dim var As New cVariable.variableEditorInfoType
                var.arrValues = arrPath
                mScript.lastArray = var
                Return "#ARRAY"
            Case "Direction"
                Dim direct As cMapManager.DirectionEnum = mapManager.GetDirection()
                Dim arr() As String = mScript.mainClass(classId).Functions(functionName).returnArray
                If IsNothing(arr) OrElse direct > arr.Count - 1 Then Return direct.ToString
                Return arr(direct)
            Case "Flash"
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function ClassMap_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If questEnvironment.EDIT_MODE Then Return ""
        If newValue = PREV_VALUE Then Return ""

        Dim mapId As Integer = -1
        If funcParams.Count > 0 Then
            If funcParams(0) = "-1" Then Return ""
            mapId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If mapId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует имени и Id карты {0}!", funcParams(0))
                Return "#Error"
            End If
        End If
        Dim cellId As Integer = -1
        If funcParams.Count > 1 Then
            cellId = GetThirdChildIdByName(funcParams(1), mapId, mScript.mainClass(classId).ChildProperties)
            If cellId < -1 Then Return _ERROR("Не найдена клетка карты " & funcParams(1))
        End If
        If GVARS.G_CURMAP <> mapId Then
            If propertyName <> "Width" AndAlso propertyName = "Height" AndAlso propertyName = "Slot" Then Return ""
        End If

        Select Case propertyName
            Case "Fog"
                Dim fogStyle As Integer = 0
                If ReadPropertyInt(classId, "FogStyle", mapId, -1, fogStyle, funcParams) = False Then Return "#Error"
                If fogStyle = 0 Then Return ""
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                mapManager.CellChangeFog(Not CBool(nValue), cellId, funcParams)
                Return ""
            Case "CellUsualClass"
                If cellId < 0 Then Return ""
                Dim hCell As HtmlElement = mapManager.GetCellHtmlElement(cellId, Nothing)
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                hCell.SetAttribute("ClassName", nValue)
                Return ""
            Case "ArrowBack", "ArrowFront", "ArrowLeft", "ArrowRight"
                If cellId < 0 OrElse cellId <> GVARS.G_CURMAPCELL Then Return ""
                mapManager.ShowArrowInCurrentCell()
                Return ""
            Case "CellCurrentClass"
                mapManager.ShowArrowInCurrentCell()
                Return ""
            Case "CellPictureCurrent"
                If cellId < 0 OrElse cellId <> GVARS.G_CURMAPCELL Then Return ""
                mapManager.ClearPreviousBitmap()
                mapManager.ShowArrowInCurrentCell()
                Return ""
            Case "UseArrows", "CellCurrentClass"
                mapManager.ShowArrowInCurrentCell()
                Return ""
            Case "CellPictureHover"
                If cellId < 0 OrElse cellId <> GVARS.G_CURMAPCELL Then Return ""
                mapManager.ClearPreviousBitmap()
                Return ""
            Case "CellPicture", "BackPicPos", "BackPicStyle", "CellWidth", "CellHeight", "FogStyle", "MapPicStyle", "UseCancelButton", _
                "MapOffsetX", "MapOffsetY"
                Return mapManager.BuildMap(mapId, GVARS.G_CURMAPCELL, funcParams, GVARS.G_PREVMAPCELL)
            Case "CellX"
                If cellId < 0 Then Return ""
                Dim cellY As String = ""
                If ReadProperty(classId, "CellY", mapId, cellId, cellY, funcParams) = False Then Return "#Error"
                cellY = UnWrapString(cellY)
                If String.IsNullOrEmpty(cellY) OrElse cellY = "0" Then Return ""
                'для отображерния клетки должны быть указаны обе координаты
                Return mapManager.BuildMap(mapId, GVARS.G_CURMAPCELL, funcParams, GVARS.G_PREVMAPCELL)
            Case "CellY"
                If cellId < 0 Then Return ""
                Dim cellX As String = ""
                If ReadProperty(classId, "CellX", mapId, cellId, cellX, funcParams) = False Then Return "#Error"
                cellX = UnWrapString(cellX)
                If String.IsNullOrEmpty(cellX) OrElse cellX = "0" Then Return ""
                'для отображерния клетки должны быть указаны обе координаты
                Return mapManager.BuildMap(mapId, GVARS.G_CURMAPCELL, funcParams, GVARS.G_PREVMAPCELL)
            Case "Caption", "CellEmptyClass", "CellInFogClass", "CellsByX", "CellsByY"
                If cellId > -1 Then Return ""
                'только самой карты (не клеток)
                Return mapManager.BuildMap(mapId, GVARS.G_CURMAPCELL, funcParams, GVARS.G_PREVMAPCELL)
            Case "CellInfoClass"
                Dim hDoc As HtmlDocument = questEnvironment.wbMap.Document
                If IsNothing(hDoc) Then Return ""
                Dim hCapt As HtmlElement = hDoc.GetElementById("CellInfo")
                If IsNothing(hCapt) Then Return ""
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                hCapt.SetAttribute("ClassName", nValue)
            Case "CellPictureWidth", "CellPictureHeight"
                If cellId < 0 Then Return ""
                Dim pName As String
                If propertyName = "CellPictureWidth" Then
                    pName = "width"
                Else
                    pName = "height"
                End If
                Dim hImg As HtmlElement = mapManager.GetCellImageHTMLElement(cellId, Nothing)
                If IsNothing(hImg) Then Return ""
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                If IsNumeric(nValue) Then nValue &= "px"
                If String.IsNullOrEmpty(nValue) Then
                    HTMLRemoveCSSstyle(hImg, pName)
                Else
                    HTMLAddCSSstyle(hImg, pName, nValue)
                End If
                Return ""
            Case "CellPictureOffsetX"
                If cellId < 0 Then Return ""
                Dim hImg As HtmlElement = mapManager.GetCellImageHTMLElement(cellId, Nothing)
                If IsNothing(hImg) Then Return ""
                Dim hCell As HtmlElement = mapManager.GetCellHtmlElement(cellId, Nothing)
                If IsNothing(hCell) Then Return ""
                Dim locCell As Point = mapManager.GetHTMLelementCoordinates(hCell)
                Dim offset As Integer = Val(mScript.PrepareStringToPrint(newValue, funcParams)) + locCell.X

                HTMLAddCSSstyle(hImg, "left", offset.ToString & "px")
                Return ""
            Case "CellPictureOffsetY"
                If cellId < 0 Then Return ""
                Dim hImg As HtmlElement = mapManager.GetCellImageHTMLElement(cellId, Nothing)
                If IsNothing(hImg) Then Return ""
                Dim hCell As HtmlElement = mapManager.GetCellHtmlElement(cellId, Nothing)
                If IsNothing(hCell) Then Return ""
                Dim locCell As Point = mapManager.GetHTMLelementCoordinates(hCell)
                Dim offset As Integer = Val(mScript.PrepareStringToPrint(newValue, funcParams)) + locCell.Y

                HTMLAddCSSstyle(hImg, "top", offset.ToString & "px")
                Return ""
            Case "BackColor"
                If cellId > -1 Then Return ""
                Dim hDoc As HtmlDocument = questEnvironment.wbMap.Document
                If IsNothing(hDoc) Then Return ""
                Dim hEl As HtmlElement = hDoc.GetElementById("MapConvas")
                If IsNothing(hEl) Then Return ""

                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                If String.IsNullOrEmpty(nValue) Then
                    HTMLRemoveCSSstyle(hEl, "background-color")
                Else
                    HTMLAddCSSstyle(hEl, "background-color", nValue)
                End If

                Return ""
            Case "BackPicture", "MapPicture"
                If cellId > -1 Then Return ""
                Dim hDoc As HtmlDocument = questEnvironment.wbMap.Document
                If IsNothing(hDoc) Then Return ""
                Dim hEl As HtmlElement
                If propertyName = "BackPicture" Then
                    hEl = hDoc.GetElementById("MapConvas")
                Else
                    hEl = hDoc.GetElementById("MapMain")
                End If
                If IsNothing(hEl) Then Return ""

                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                nValue = nValue.Replace("\", "/")
                If String.IsNullOrEmpty(nValue) Then
                    HTMLRemoveCSSstyle(hEl, "background-image")
                Else
                    HTMLAddCSSstyle(hEl, "background-image", "url(" & nValue & ")")
                End If
                Return ""
            Case "Content"
                If cellId < 0 Then Return ""
                Dim hCell As HtmlElement = mapManager.GetCellHtmlElement(cellId, Nothing)
                If IsNothing(hCell) Then Return ""
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                hCell.InnerHtml = nValue
                Return ""
            Case "CSS"
                If cellId > -1 Then Return ""
                Dim hDoc As HtmlDocument = questEnvironment.wbMap.Document
                If IsNothing(hDoc) Then Return ""
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                HtmlChangeFirstCSSLink(hDoc, nValue)
                Return ""
            Case "Width", "Height", "Slot"
                SetWindowsContainers()
                frmMap.Hide()
            Case "Location"
                If cellId < 0 Then Return ""
                Dim hCell As HtmlElement = mapManager.GetCellHtmlElement(cellId, Nothing)
                Dim nValue As String = mScript.PrepareStringToPrint(newValue, funcParams)
                If String.IsNullOrEmpty(nValue) Then
                    hCell.Style = "cursor:pointer"
                Else
                    hCell.Style = ""
                End If
                Return ""
            Case "ShowStyle"
                If cellId < 0 Then Return ""
                Dim hCell As HtmlElement = mapManager.GetCellHtmlElement(cellId, Nothing)
                If IsNothing(hCell) Then Return ""
                If mapManager.IsCellVisible(GVARS.G_CURMAP, cellId) Then
                    mapManager.CellChangeFog(True, cellId, funcParams, hCell)
                Else
                    mapManager.CellChangeFog(False, cellId, funcParams, hCell)
                End If
                Return ""
        End Select
        Return ""
    End Function

#End Region

#Region "Hero"
    Private Function classH_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго порядка
                Return CreateNewObject(classId, funcParams)
            Case "Count"
                If GVARS.G_ISBATTLE Then
                    'Идет бой
                    Dim ret As Integer = mScript.Battle.FightersCount(False, 0)
                    If ret = -1 Then
                        Return "#Error"
                    Else
                        Return ret.ToString
                    End If
                Else
                    'вне боя
                    Return ObjectsCount(classId, funcParams, False)
                End If
            Case "HasAbility"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hId As Integer = -1, abSet As String = ""
                Dim classAb As Integer = mScript.mainClassHash("Ab")
                If GVARS.G_ISBATTLE Then
                    hId = mScript.Battle.GetFighterByName(funcParams(0))
                    If hId < 0 Then Return _ERROR("Бойца " & funcParams(0) & " не найдено.", functionName)
                    If mScript.Battle.ReadFighterProperty("AbilitiesSet", hId, abSet, {hId.ToString}) = False Then Return "#Error"
                Else
                    hId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                    If hId < 0 Then Return _ERROR("Перонажа " & funcParams(0) & " не найдено.", functionName)
                    If ReadProperty(classId, "AbilitiesSet", hId, -1, abSet, {hId.ToString}) = False Then Return "#Error"
                End If
                If abSet = "''" Then Return "False"
                Dim abSetId As Integer = GetSecondChildIdByName(abSet, mScript.mainClass(classAb).ChildProperties)
                If abSetId < 0 Then
                    'Return _ERROR("Набора способностей " & abSet & " не найдено.", functionName)
                    Return "False"
                End If
                Dim abId As Integer = -1
                abId = GetThirdChildIdByName(funcParams(1), abSetId, mScript.mainClass(classAb).ChildProperties)
                Return (abId >= 0).ToString
            Case "HasMagic"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hId As Integer = -1, mBook As String = ""
                Dim classMg As Integer = mScript.mainClassHash("Mg")
                If GVARS.G_ISBATTLE Then
                    hId = mScript.Battle.GetFighterByName(funcParams(0))
                    If hId < 0 Then Return _ERROR("Бойца " & funcParams(0) & " не найдено.", functionName)
                    If mScript.Battle.ReadFighterProperty("MagicBook", hId, mBook, {hId.ToString}) = False Then Return "#Error"
                Else
                    hId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                    If hId < 0 Then Return _ERROR("Перонажа " & funcParams(0) & " не найдено.", functionName)
                    If ReadProperty(classId, "MagicBook", hId, -1, mBook, {hId.ToString}) = False Then Return "#Error"
                End If
                If mBook = "''" Then Return "False"
                Dim mBookId As Integer = GetSecondChildIdByName(mBook, mScript.mainClass(classMg).ChildProperties)
                If mBookId < 0 Then
                    'Return _ERROR("Книги заклинаний " & mBook & " не найдено.", functionName)
                    Return "False"
                End If
                Dim mgId As Integer = -1
                mgId = GetThirdChildIdByName(funcParams(1), mBookId, mScript.mainClass(classMg).ChildProperties)
                Return (mgId >= 0).ToString
            Case "Id"
                If GVARS.G_ISBATTLE Then
                    Return mScript.Battle.GetFighterByName(funcParams(0)).ToString
                Else
                    Return ObjectId(classId, funcParams)
                End If
            Case "IsExist"
                'Определяет есть ли элемент с таким именем/Id
                If GVARS.G_ISBATTLE Then
                    Dim id As Integer = mScript.Battle.GetFighterByName(funcParams(0))
                    If id < 0 Then
                        Return "False"
                    Else
                        Return "True"
                    End If
                Else
                    Return ObjectIsExists(classId, funcParams)
                End If
            Case "Remove"
                'удаление элемента
                Return RemoveObject(classId, funcParams)
            Case "SetParams"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hId As Integer = -1
                If GVARS.G_ISBATTLE Then
                    hId = mScript.Battle.GetFighterByName(funcParams(0))
                Else
                    hId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                End If
                If hId < 0 Then Return _ERROR("Персонаж " & funcParams(0) & " не найден.", functionName)

                Dim var As cVariable.variableEditorInfoType = GetVariable(UnWrapString(funcParams(1)))
                If IsNothing(var) Then Return _ERROR("Переменной " & funcParams(1) & " не существует.", functionName)

                Return HeroSetParamsFromVariable(hId, var)
            Case "RestoreParams"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hId As Integer = -1
                If GVARS.G_ISBATTLE Then
                    hId = mScript.Battle.GetFighterByName(funcParams(0))
                Else
                    hId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                End If
                If hId < 0 Then Return _ERROR("Персонаж " & funcParams(0) & " не найден.", functionName)

                Return HeroRestoreParams(hId)
            Case "ShowMagicBook"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hId As Integer = -1, mBook As String = ""
                Dim classMg As Integer = mScript.mainClassHash("Mg")
                If GVARS.G_ISBATTLE Then
                    hId = mScript.Battle.GetFighterByName(funcParams(0))
                    If hId < 0 Then Return _ERROR("Бойца " & funcParams(0) & " не найдено.", functionName)
                    If mScript.Battle.ReadFighterProperty("MagicBook", hId, mBook, {hId.ToString}) = False Then Return "#Error"
                Else
                    hId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                    If hId < 0 Then Return _ERROR("Перонажа " & funcParams(0) & " не найдено.", functionName)
                    If ReadProperty(classId, "MagicBook", hId, -1, mBook, {hId.ToString}) = False Then Return "#Error"
                End If
                If mBook = "''" Then Return _ERROR("У персонажа " & funcParams(0) & " не указана книга заклинаний.", functionName)
                Dim mBookId As Integer = GetSecondChildIdByName(mBook, mScript.mainClass(classMg).ChildProperties)
                If mBookId < 0 Then
                    Return _ERROR("Книги заклинаний " & mBook & " не найдено.", functionName)
                End If
                Dim mgId As Integer = -1, mgStr As String = GetParam(funcParams, 1, "-1")
                If mgStr <> "-1" Then
                    mgId = GetThirdChildIdByName(mgStr, mBookId, mScript.mainClass(classMg).ChildProperties)
                    If mgId < 0 Then Return _ERROR("Магии " & funcParams(1) & "в книге заклинаний персонажа " & funcParams(0) & " не найдено.", functionName)
                End If

                'смена текущей книги магии
                If CreateMagicBook(hId, mBookId, funcParams, mgId) = "#Error" Then Return "#Error"

                Dim slot As Integer = 0, isNewWindow As Boolean = False
                Dim classMW As Integer = mScript.mainClassHash("MgW")
                If ReadPropertyInt(classMW, "Slot", -1, -1, slot, funcParams) = False Then Return "#Error"
                If slot = 5 Then isNewWindow = True 'в отдельном окне

                If isNewWindow Then
                    questEnvironment.wbMagic.Show()
                    frmMagic.Show()
                Else
                    questEnvironment.wbMagic.BringToFront()
                    questEnvironment.wbMagic.Show()
                End If

                Return ""
            Case "CanUseMagic"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hId As Integer = -1, mBook As String = ""
                Dim classMg As Integer = mScript.mainClassHash("Mg")
                If GVARS.G_ISBATTLE Then
                    hId = mScript.Battle.GetFighterByName(funcParams(0))
                    If hId < 0 Then Return _ERROR("Бойца " & funcParams(0) & " не найдено.", functionName)
                    If mScript.Battle.ReadFighterProperty("MagicBook", hId, mBook, {hId.ToString}) = False Then Return "#Error"
                Else
                    hId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                    If hId < 0 Then Return _ERROR("Перонажа " & funcParams(0) & " не найдено.", functionName)
                    If ReadProperty(classId, "MagicBook", hId, -1, mBook, {hId.ToString}) = False Then Return "#Error"
                End If
                If mBook = "''" OrElse mBook = "" OrElse mBook = "-1" Then Return "False" '_ERROR("У персонажа " & funcParams(0) & " не указана книга заклинаний.", functionName)
                Dim mBookId As Integer = GetSecondChildIdByName(mBook, mScript.mainClass(classMg).ChildProperties)
                If mBookId < 0 Then
                    Return _ERROR("Книги заклинаний " & mBook & " не найдено.", functionName)
                End If
                Dim mgId As Integer = -1, mgStr As String = GetParam(funcParams, 1, "-1")
                If mgStr <> "-1" Then
                    mgId = GetThirdChildIdByName(mgStr, mBookId, mScript.mainClass(classMg).ChildProperties)
                    If mgId < 0 Then Return _ERROR("Магии " & funcParams(1) & "в книге заклинаний персонажа " & funcParams(0) & " не найдено.", functionName)
                End If

                If mgId > -1 Then
                    Return CanUseMagic(hId, mBookId, mgId).ToString
                Else
                    'проверяем есть ли хоть одна доступная магия
                    If IsNothing(mScript.mainClass(classMg).ChildProperties(mBookId)("Name").ThirdLevelProperties) OrElse mScript.mainClass(classMg).ChildProperties(mBookId)("Name").ThirdLevelProperties.Count = 0 Then Return "False"
                    For i As Integer = 0 To mScript.mainClass(classMg).ChildProperties(mBookId)("Name").ThirdLevelProperties.Count - 1
                        If CanUseMagic(hId, mBookId, i) Then Return "True"
                    Next
                    Return "False"
                End If
                'Case "HasEnabledMagics"
                '    If questEnvironment.EDIT_MODE Then Return ""
                '    Dim hId As Integer = -1
                '    If GVARS.G_ISBATTLE Then
                '        hId = mScript.Battle.GetFighterByName(funcParams(0))
                '    Else
                '        hId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                '    End If
                '    If hId < 0 Then
                '        Return _ERROR("Перонажа " & funcParams(0) & " не найдено.", functionName)
                '    End If
                '    Return HasEnabledMagics(hId).ToString
            Case "CastMagic"
                If questEnvironment.EDIT_MODE Then Return ""
                Dim hId As Integer = -1, mBook As String = ""
                Dim classMg As Integer = mScript.mainClassHash("Mg")
                If GVARS.G_ISBATTLE Then
                    hId = mScript.Battle.GetFighterByName(funcParams(0))
                    If hId < 0 Then Return _ERROR("Бойца " & funcParams(0) & " не найдено.", functionName)
                    If mScript.Battle.ReadFighterProperty("MagicBook", hId, mBook, {hId.ToString}) = False Then Return "#Error"
                Else
                    hId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                    If hId < 0 Then Return _ERROR("Перонажа " & funcParams(0) & " не найдено.", functionName)
                    If ReadProperty(classId, "MagicBook", hId, -1, mBook, {hId.ToString}) = False Then Return "#Error"
                End If

                If mBook = "''" Then Return _ERROR("У персонажа " & funcParams(0) & " не указана книга заклинаний.", functionName)
                Dim mBookId As Integer = GetSecondChildIdByName(mBook, mScript.mainClass(classMg).ChildProperties)
                If mBookId < 0 Then
                    Return _ERROR("Книги заклинаний " & mBook & " не найдено.", functionName)
                End If
                Dim mgId As Integer = -1, mgStr As String = GetParam(funcParams, 1, "-1")
                If mgStr <> "-1" Then
                    mgId = GetThirdChildIdByName(mgStr, mBookId, mScript.mainClass(classMg).ChildProperties)
                    If mgId < 0 Then Return _ERROR("Магии " & funcParams(1) & "в книге заклинаний персонажа " & funcParams(0) & " не найдено.", functionName)
                End If
                'проверка доступности
                If CBool(GetParam(funcParams, 2, "True")) Then
                    If Not CanUseMagic(hId, mBookId, mgId) Then Return "False"
                End If

                EventGeneratedFromScript = True
                Return CastMagic(hId, mBookId, mgId, CBool(GetParam(funcParams, 3, "True"))).ToString
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function


    Private Function ClassH_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If questEnvironment.EDIT_MODE Then Return ""

        Dim hId As Integer = -1
        If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then
            If funcParams(0) = "-1" Then Return ""
            hId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If hId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует персонажа {0}!", funcParams(0))
                Return "#Error"
            End If
        Else
            Return ""
        End If
        'Здесь только если меняется свойство элемента 2 уровня

        Select Case propertyName
            Case "Life"
                Dim nValue As Double = 0
                If ReadPropertyDbl(classId, "Life", hId, -1, nValue, funcParams) = False Then Return "#Error"

                If nValue <= 0 Then
                    'событие смерти персонажа
                    'глобальное
                    Dim res As String = "", arrs() As String = {hId.ToString, "False"}
                    Dim eventId As Integer = mScript.mainClass(classId).Properties("HeroDeadEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "HeroDeadEvent", False)
                        If res = "False" OrElse res = "#Error" Then Return res
                    End If
                    'данного героя
                    eventId = mScript.mainClass(classId).ChildProperties(hId)("HeroDeadEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "HeroDeadEvent", False)
                        If res = "False" OrElse res = "#Error" Then Return res
                    End If
                End If
                Return ""
        End Select
        Return ""
    End Function

    Private Function ClassFighter_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String)
        If questEnvironment.EDIT_MODE Then Return ""
        If IsNothing(funcParams) OrElse funcParams.Count <> 1 Then Return ""

        Dim Battle As cBattle = mScript.Battle
        Dim hId As Integer = Battle.GetFighterByName(funcParams(0))
        If hId < -1 Then
            mScript.LAST_ERROR = String.Format("Не существует бойца {0}!", funcParams(0))
            Return "#Error"
        End If

        'Здесь только если меняется свойство элемента 2 уровня
        Select Case propertyName
            Case "AbilitiesSet"
                Dim classAb As Integer = mScript.mainClassHash("Ab")
                Dim abSetId As Integer = GetSecondChildIdByName(newValue, mScript.mainClass(classAb).ChildProperties)
                If abSetId < 0 Then Return _ERROR("Не найден набор способностей " & newValue)

                Battle.Fighters(hId).abilitySetId = abSetId
                Dim hFighter As HtmlElement = Battle.FindFighterHTMLContainer(hId, Nothing)
                If IsNothing(hFighter) Then Return ""
                'If Battle.PutFighterToBattleField(hId, hFighter.Parent, hFighter) = "#Error" Then Return "#Error"
                If Battle.ChangeAbilitiesIcons(hId) = "#Error" Then Return "#Error"
            Case "Army"
                If newValue = PREV_VALUE Then Return ""
                Dim classArmy As Integer = mScript.mainClassHash("Army")
                Dim armyId As Integer = GetSecondChildIdByName(newValue, mScript.mainClass(classArmy).ChildProperties)
                If armyId < 0 AndAlso (newValue <> "" AndAlso newValue <> "''" AndAlso newValue <> "-1") Then Return _ERROR("Не найдена армия " & newValue)
                If Battle.RemoveFighterFromBattlefield(hId) = "#Error" Then Return "#Error"
                Battle.Fighters(hId).armyId = armyId
                If armyId < 0 Then Battle.Fighters(hId).armyUnitId = -1

                If Battle.AppendFighterInBattle(hId) = "#Error" Then Return "#Error"

                Dim prevArmyId As Integer = GetSecondChildIdByName(PREV_VALUE, mScript.mainClass(classArmy).ChildProperties)
                If prevArmyId > -1 Then
                    'Событие ArmyOnCountChanged (5 - перешел в другую армию)
                    'получаем параметры
                    Dim unitsLeft As Integer = mScript.Battle.UnitsInArmyLeft(prevArmyId)
                    If unitsLeft < 0 Then Return "#Error"
                    Dim arrs() As String = {prevArmyId.ToString, hId.ToString, "5", "-1", unitsLeft.ToString}

                    'глобальное событие
                    Dim eventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnCountChanged").eventId
                    Dim res As String = ""
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                        If res = "#Error" Then Return "#Error"
                    End If

                    'событие армии
                    If res <> "False" Then
                        eventId = mScript.mainClass(classArmy).ChildProperties(prevArmyId)("ArmyOnCountChanged").eventId
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                            If res = "#Error" Then Return "#Error"
                        End If
                    End If
                End If

                If armyId > -1 Then
                    'Событие ArmyOnCountChanged (1 - переметнулся из другой армии)
                    'получаем параметры
                    Dim unitsLeft As Integer = mScript.Battle.UnitsInArmyLeft(armyId)
                    If unitsLeft < 0 Then Return "#Error"
                    Dim arrs() As String = {armyId.ToString, hId.ToString, "1", "1", unitsLeft.ToString}

                    'глобальное событие
                    Dim eventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnCountChanged").eventId
                    Dim res As String = ""
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                        If res = "#Error" Then Return "#Error"
                    End If

                    'событие армии
                    If res <> "False" Then
                        eventId = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnCountChanged").eventId
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                            If res = "#Error" Then Return "#Error"
                        End If
                    End If
                End If
            Case "Life"
                Dim nValue As Double = 0
                If Battle.ReadFighterPropertyDbl("Life", hId, nValue, funcParams) = False Then Return "#Error"
                If Battle.ChangeFighterParam(hId, "Life") = "#Error" Then Return "#Error"

                ''эффект удара
                ''...
                'Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                'Dim hFighter As HtmlElement = Battle.FindFighterHTMLContainer(hId, frmPlayer.wbMain.Document)
                'Dim hEff As HtmlElement = Nothing
                'Dim hImg As HtmlElement = Battle.FindFighterHTMLImage(hFighter, hId, hEff)
                'If IsNothing(hImg) = False AndAlso IsNothing(hEff) = False Then
                '    Dim effClass As String = ""
                '    If ReadProperty(classId, "EffectClass", hId, -1, effClass, {hId.ToString}) = False Then Return "#Error"
                '    effClass = UnWrapString(effClass)

                '    If String.IsNullOrEmpty(effClass) = False Then
                '        hEff.Style &= "display:block;"
                '        Dim EffectPicture As String = ""
                '        If Battle.ReadFighterProperty("EffectPicture", hId, EffectPicture, {hId.ToString}) = False Then Return "#Error"
                '        EffectPicture = UnWrapString(EffectPicture).Replace("\"c, "/"c)
                '        hEff.SetAttribute("src", EffectPicture)
                '        HTMLAddClass(hEff, effClass)
                '        hEff.SetAttribute("ClassName", effClass)
                '        Dim msEff As mshtml.HTMLImg = hEff.DomElement
                '        Do While Not msEff.complete
                '            Application.DoEvents()
                '        Loop
                '    End If
                'End If

                'эффект удара
                Dim hEff As HtmlElement = Nothing
                If Battle.SetDamageEffectToVictim(hId, hEff) = "#Error" Then Return "#Error"

                If nValue <= 0 Then
                    'событие смерти персонажа
                    'событие боя BattleFighterDeathEvent
                    Dim res As String = "", arrs() As String = {hId.ToString, Battle.killingSummoned.ToString}
                    Dim eventId As Integer = mScript.mainClass(mScript.mainClassHash("B")).Properties("BattleFighterDeathEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "BattleFighterDeathEvent", False)
                        If res = "#Error" Then Return res
                    End If
                    Dim killed As Integer = 0
                    If res <> "False" Then
                        killed = Battle.KillSummoned(hId)
                        If killed = -1 Then Return "#Error"
                    End If

                    'глобальное
                    eventId = mScript.mainClass(classId).Properties("HeroDeadEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "HeroDeadEvent", False)
                        If res = "False" OrElse res = "#Error" Then Return res
                    End If
                    'данного героя
                    eventId = Battle.Fighters(hId).heroProps("HeroDeadEvent").eventId
                    If eventId > 0 Then
                        res = mScript.eventRouter.RunEvent(eventId, arrs, "HeroDeadEvent", False)
                        If res = "False" OrElse res = "#Error" Then Return res
                    End If

                    'звук смерти
                    Dim sound As String = ""
                    If ReadProperty(classId, "SoundOnDeath", hId, -1, sound, {hId.ToString}) = False Then Return "#Error"
                    sound = UnWrapString(sound)
                    If String.IsNullOrEmpty(sound) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, sound)) Then
                        Say(sound, 100, False)
                        mScript.Battle.Wait(100)
                    End If

                    Battle.Wait()
                    If Battle.RemoveAbilitiesWithApplierDeath(hId) = "True" Then Battle.Wait() 'удаление способностей с RemoveWithApplierDeath = True

                    Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                    Dim hFighter As HtmlElement = Battle.FindFighterHTMLContainer(hId, frmPlayer.wbMain.Document)
                    If IsNothing(hFighter) = False Then hFighter.Style &= "opacity:0;"
                    Battle.Wait()
                    'удаляем с поля боя
                    If Battle.RemoveFighterFromBattlefield(hId) = "#Error" Then Return "#Error"

                    If Battle.killingSummoned = False AndAlso Battle.Fighters(hId).armyId > -1 Then
                        'Событие ArmyOnCountChanged (3 - погиб)
                        Dim armyId As Integer = Battle.Fighters(hId).armyId, classArmy As Integer = mScript.mainClassHash("Army")
                        killed = -1 * (killed + 1)
                        'получаем параметры
                        Dim unitsLeft As Integer = mScript.Battle.UnitsInArmyLeft(armyId)
                        If unitsLeft < 0 Then Return "#Error"
                        arrs = {armyId.ToString, hId.ToString, "3", killed.ToString, unitsLeft.ToString}

                        'глобальное событие
                        eventId = mScript.mainClass(classArmy).Properties("ArmyOnCountChanged").eventId
                        res = ""
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                            If res = "#Error" Then Return "#Error"
                        End If

                        'событие армии
                        If res <> "False" Then
                            eventId = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnCountChanged").eventId
                            If eventId > 0 Then
                                res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                                If res = "#Error" Then Return "#Error"
                            End If
                        End If
                    End If
                ElseIf IsNothing(hEff) = False Then
                    Battle.Wait()
                End If

                'убираем эффект
                If IsNothing(hEff) = False Then
                    hEff.Style &= "display:none;"
                    hEff.SetAttribute("ClassName", "EffectPicture")
                End If

                Return ""
            Case "MagicBook"
                Dim classMg As Integer = mScript.mainClassHash("Mg")
                Dim mBookId As Integer = GetSecondChildIdByName(newValue, mScript.mainClass(classMg).ChildProperties)
                If mBookId < 0 Then Return _ERROR("Не найдена книга заклинаний " & newValue)

                Battle.Fighters(hId).abilitySetId = mBookId
            Case "Picture", "PicWidth", "PicHeight"
                Dim hFighter As HtmlElement = Battle.FindFighterHTMLContainer(hId, Nothing)
                If IsNothing(hFighter) Then Return ""
                If Battle.PutFighterToBattleField(hId, hFighter.Parent, hFighter) = "#Error" Then Return "#Error"
            Case "Range"
                If Battle.RemoveFighterFromBattlefield(hId) = "#Error" Then Return "#Error"
                If Battle.AppendFighterInBattle(hId) = "#Error" Then Return "#Error"
            Case "RunAway"
                If newValue = PREV_VALUE Then Return ""

                If newValue = "True" Then
                    'сбегает
                    Battle.lstFightersToMoveOut.Add(hId)
                    Battle.timMoveFighters.Enabled = True

                    If Battle.Fighters(hId).armyId > -1 Then
                        'Событие ArmyOnCountChanged (4 - сбежал)
                        Dim armyId As Integer = Battle.Fighters(hId).armyId, classArmy As Integer = mScript.mainClassHash("Army")
                        'получаем параметры
                        Dim unitsLeft As Integer = mScript.Battle.UnitsInArmyLeft(armyId)
                        If unitsLeft < 0 Then Return "#Error"
                        Dim arrs() As String = {armyId.ToString, hId.ToString, "4", "-1", unitsLeft.ToString}

                        'глобальное событие
                        Dim eventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnCountChanged").eventId
                        Dim res As String = ""
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                            If res = "#Error" Then Return "#Error"
                        End If

                        'событие армии
                        If res <> "False" Then
                            eventId = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnCountChanged").eventId
                            If eventId > 0 Then
                                res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                                If res = "#Error" Then Return "#Error"
                            End If
                        End If
                    End If
                Else
                    'возвращается
                    If Battle.AppendFighterInBattle(hId) = "#Error" Then Return "#Error"

                    If Battle.Fighters(hId).armyId > -1 Then
                        'Событие ArmyOnCountChanged (6 - вернулся после того как сбежал)
                        Dim armyId As Integer = Battle.Fighters(hId).armyId, classArmy As Integer = mScript.mainClassHash("Army")
                        'получаем параметры
                        Dim unitsLeft As Integer = mScript.Battle.UnitsInArmyLeft(armyId)
                        If unitsLeft < 0 Then Return "#Error"
                        Dim arrs() As String = {armyId.ToString, hId.ToString, "6", "1", unitsLeft.ToString}

                        'глобальное событие
                        Dim eventId As Integer = mScript.mainClass(classArmy).Properties("ArmyOnCountChanged").eventId
                        Dim res As String = ""
                        If eventId > 0 Then
                            res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                            If res = "#Error" Then Return "#Error"
                        End If

                        'событие армии
                        If res <> "False" Then
                            eventId = mScript.mainClass(classArmy).ChildProperties(armyId)("ArmyOnCountChanged").eventId
                            If eventId > 0 Then
                                res = mScript.eventRouter.RunEvent(eventId, arrs, "ArmyOnCountChanged", False)
                                If res = "#Error" Then Return "#Error"
                            End If
                        End If
                    End If
                End If
            Case Else
                If propertyName.EndsWith("Total") OrElse mScript.mainClass(classId).Properties.ContainsKey(propertyName & "Total") Then Return Battle.ChangeFighterParam(hId, propertyName)
        End Select
        Return ""
    End Function

#End Region

#Region "Magic"
    Private Function classMg_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго порядка
                Return CreateNewObject(classId, funcParams)
            Case "Count"
                Return ObjectsCount(classId, funcParams)
            Case "IsExist"
                'Определяет есть ли элемент с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "Remove"
                'удаление элемента
                Return RemoveObject(classId, funcParams)
            Case "AddToHero"
                Dim srcBookId As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If srcBookId < 0 Then Return _ERROR("Книги магии " & funcParams(0) & " не найдено.", functionName)
                Dim mgId As Integer = GetThirdChildIdByName(funcParams(1), srcBookId, mScript.mainClass(classId).ChildProperties)
                If mgId < 0 Then Return _ERROR("Магии " & funcParams(1) & " в книге " & funcParams(0) & " не найдено.", functionName)

                Dim destBookStr As String = "", destBookId As Integer
                Dim hId As Integer = -1
                If GVARS.G_ISBATTLE Then
                    hId = mScript.Battle.GetFighterByName(funcParams(2))
                    If hId < 0 Then Return _ERROR("Бойца " & funcParams(2) & " не найдено.", functionName)
                    'получаем книгу героя
                    If mScript.Battle.ReadFighterProperty("MagicBook", hId, destBookStr, {hId.ToString}) = False Then Return "#Error"
                Else
                    Dim classH As Integer = mScript.mainClassHash("H")
                    hId = GetSecondChildIdByName(funcParams(2), mScript.mainClass(classH).ChildProperties)
                    If hId < 0 Then Return _ERROR("Персонажа " & funcParams(2) & " не найдено.", functionName)
                    'получаем книгу героя
                    If ReadProperty(classH, "MagicBook", hId, -1, destBookStr, {hId.ToString}) = False Then Return "#Error"
                End If

                If destBookStr = "''" OrElse destBookStr = "" OrElse destBookStr = "-1" Then
                    'Создаем книгу заклинаний если ее не было
                    Dim newName As String = "", i As Integer = 1
                    Do
                        newName = "'MagicBook" & i.ToString & "'"
                        If ObjectIsExists(classId, {newName.ToArray}) <> "True" Then Exit Do
                        i += 1
                    Loop

                    destBookStr = CreateNewObject(classId, {newName})
                    destBookId = Val(destBookStr)
                    If PropertiesRouter(mScript.mainClassHash("H"), "MagicBook", {hId.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, destBookId.ToString) = "#Error" Then Return "#Error"
                Else
                    destBookId = GetSecondChildIdByName(destBookStr, mScript.mainClass(classId).ChildProperties)
                End If
                If destBookId < 0 Then Return _ERROR("Книги магии " & destBookStr & " не найдено.", functionName)

                Dim mgName As String = mScript.mainClass(classId).ChildProperties(srcBookId)("Name").ThirdLevelProperties(mgId)
                If ObjectIsExists(classId, {destBookId.ToString, mgName}) = "True" Then Return "False" 'магия уже существует в книги магий героя

                'копируем магию mgId из srcBookId в destBookId
                Dim res As String = CreateNewObject(classId, {destBookId.ToString, mgName})
                If res = "#Error" Then Return res

                Dim newMgId As Integer = CInt(res)
                For i As Integer = 0 To mScript.mainClass(classId).ChildProperties(srcBookId).Count - 1
                    Dim pName As String = mScript.mainClass(classId).ChildProperties(srcBookId).ElementAt(i).Key
                    Dim pValue As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(srcBookId).ElementAt(i).Value

                    If pName <> "Name" Then
                        SetPropertyValue(classId, pName, pValue.ThirdLevelProperties(mgId), destBookId, newMgId) 'событие само продублируется при необходимости
                    End If
                Next
                Return "True"
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function classMW_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Show"
                If questEnvironment.EDIT_MODE Then Return ""
                If GVARS.G_ISBATTLE AndAlso GVARS.G_CANUSEMAGIC = False Then Return ""
                Dim mBookId As Integer = 0 'GetSecondChildIdByName(GetParam(funcParams, 0, GVARS.G_CURMAP.ToString), mScript.mainClass(classId).ChildProperties)
                If mBookId < 0 Then
                    Return _ERROR("Карты " & funcParams(0) & " не найдено.", functionName)
                End If

                'смена текущей книги магии
                If CreateMagicBook(0, mBookId, funcParams) = "#Error" Then Return "#Error"

                Dim slot As Integer = 0, isNewWindow As Boolean = False
                If ReadPropertyInt(classId, "Slot", -1, -1, slot, funcParams) = False Then Return "#Error"
                If slot = 5 Then isNewWindow = True 'в отдельном окне

                If isNewWindow Then
                    questEnvironment.wbMagic.Show()
                    frmMagic.Show()
                Else
                    questEnvironment.wbMagic.BringToFront()
                    questEnvironment.wbMagic.Show()
                End If

                'Dim Hel As HtmlElement = questEnvironment.wbMagic.Document.GetElementById("magicBook")
                'Hel = Hel.Children(0).Children(0).Children(0)
                'MsgBox(Hel.OffsetRectangle.Height.ToString)
                Return ""
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function
#End Region

#Region "Ability"
    Private Function classAb_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго порядка
                Return CreateNewObject(classId, funcParams)
            Case "Count"
                Return ObjectsCount(classId, funcParams)
            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "IsExist"
                'Определяет есть ли элемент с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Remove"
                'удаление элемента
                If questEnvironment.EDIT_MODE = False Then
                    'событие снятия способности
                    Dim abSet As String = GetParam(funcParams, 0, "-1"), ab As String = GetParam(funcParams, 1, "-1")
                    Dim abSetId As Integer = -1, abId As Integer = -1
                    If abSet <> "-1" Then abSetId = GetSecondChildIdByName(abSet, mScript.mainClass(classId).ChildProperties)
                    If abSetId >= 0 AndAlso ab <> "-1" Then abId = GetThirdChildIdByName(ab, abSetId, mScript.mainClass(classId).ChildProperties)
                    If abId > -1 Then
                        'при удалении целого набора или всех наборов способностей не запускаем скрипт
                        Dim res As String
                        Dim arrs() As String = {abSetId.ToString, abId.ToString, ""}
                        'удаление способности
                        Dim eventIdGlobal As Integer = mScript.mainClass(classId).Properties("AbilityOnReleaseEvent").eventId
                        Dim eventIdAbSet As Integer = mScript.mainClass(classId).ChildProperties(abSetId)("AbilityOnReleaseEvent").eventId
                        Dim eventId As Integer = mScript.mainClass(classId).ChildProperties(abSetId)("AbilityOnReleaseEvent").ThirdLevelEventId(abId)
                        Dim hList As List(Of Integer) = GetHeroesListByAbilitySet(abSetId)
                        Dim BattleIcon As String = ""
                        If GVARS.G_ISBATTLE Then
                            'если идет бой, то получаем иконку способности для размещения воле бойца
                            If ReadProperty(classId, "BattleIcon", abSetId, abId, BattleIcon, arrs) = False Then Return "#Error"
                            BattleIcon = UnWrapString(BattleIcon)
                            If String.IsNullOrEmpty(BattleIcon) OrElse My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, BattleIcon)) = False Then BattleIcon = ""
                            BattleIcon = BattleIcon.Replace("\"c, "/"c)
                        End If

                        For h As Integer = 0 To hList.Count - 1
                            'перебираем всех героев, у которых набор способностей равен abSetId, и выполняем события снятия способности для каждого из них. 
                            'Установить один набор нескольким героям можно, но не желательно: если для одного из них событие AbilityOnApplyEvent вернет False, то событие для остальных не запустится
                            arrs(2) = hList(h).ToString
                            'глобальное
                            If eventIdGlobal > 0 Then
                                res = mScript.eventRouter.RunEvent(eventIdGlobal, arrs, "AbilityOnReleaseEvent", False)
                                If res = "False" Then
                                    Return ""
                                End If
                            End If

                            'данного набора
                            If eventIdAbSet > 0 Then
                                res = mScript.eventRouter.RunEvent(eventIdAbSet, arrs, "AbilityOnReleaseEvent", False)
                                If res = "False" Then
                                    Return ""
                                End If
                            End If

                            'данной способности
                            If eventId > 0 Then
                                res = mScript.eventRouter.RunEvent(eventId, arrs, "AbilityOnReleaseEvent", False)
                                If res = "False" Then
                                    Return ""
                                End If
                            End If

                            If String.IsNullOrEmpty(BattleIcon) = False Then
                                'печатаем икону способности
                                If mScript.Battle.ChangeAbilitiesIcons(hList(h), abId) = "#Error" Then Return "#Error"
                            End If
                        Next h

                        'звук при снятии SoundOnRelease
                        Dim snd As String = ""
                        If ReadProperty(classId, "SoundOnRelease", abSetId, abId, snd, {abSetId.ToString, abId.ToString}) = False Then Return "#Error"
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

                    End If
                End If
                Return RemoveObject(classId, funcParams)
            Case "AddToHero"
                Dim srcAbSet As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                If srcAbSet < 0 Then Return _ERROR("Набора способностей " & funcParams(0) & " не найдено.", functionName)
                Dim AbId As Integer = GetThirdChildIdByName(funcParams(1), srcAbSet, mScript.mainClass(classId).ChildProperties)
                If AbId < 0 Then Return _ERROR("Способности " & funcParams(1) & " в наборе " & funcParams(0) & " не найдено.", functionName)

                'получаем набор способностей героя
                Dim destAbSetStr As String = "", destAbSetId As Integer, hId As Integer
                If GVARS.G_ISBATTLE Then
                    hId = mScript.Battle.GetFighterByName(funcParams(2))
                    If hId < 0 Then Return _ERROR("Бойца " & funcParams(2) & " не найдено.", functionName)
                    'получаем набор способностей героя
                    If mScript.Battle.ReadFighterProperty("AbilitiesSet", hId, destAbSetStr, {hId.ToString}) = False Then Return "#Error"
                Else
                    Dim classH As Integer = mScript.mainClassHash("H")
                    hId = GetSecondChildIdByName(funcParams(2), mScript.mainClass(classH).ChildProperties)
                    If hId < 0 Then Return _ERROR("Персонажа " & funcParams(2) & " не найдено.", functionName)
                    'получаем книгу героя
                    If ReadProperty(classH, "AbilitiesSet", hId, -1, destAbSetStr, {hId.ToString}) = False Then Return "#Error"
                End If

                If destAbSetStr = "''" OrElse destAbSetStr = "" OrElse destAbSetStr = "-1" Then
                    'Создаем набор способностей если не было
                    Dim newName As String = "", i As Integer = 1
                    Do
                        newName = "'AbilitiesSet" & i.ToString & "'"
                        If ObjectIsExists(classId, {newName.ToArray}) <> "True" Then Exit Do
                        i += 1
                    Loop

                    destAbSetStr = CreateNewObject(classId, {newName})
                    destAbSetId = Val(destAbSetStr)
                    If PropertiesRouter(mScript.mainClassHash("H"), "AbilitiesSet", {hId.ToString}, Nothing, PropertiesOperationEnum.PROPERTY_SET, destAbSetId.ToString) = "#Error" Then Return "#Error"
                Else
                    destAbSetId = GetSecondChildIdByName(destAbSetStr, mScript.mainClass(classId).ChildProperties)
                End If
                If destAbSetId < 0 Then Return _ERROR("Набора способностей " & destAbSetStr & " не найдено.", functionName)

                Dim abName As String = mScript.mainClass(classId).ChildProperties(srcAbSet)("Name").ThirdLevelProperties(AbId)
                Dim arrProp() As String = {destAbSetId.ToString, abName}
                Dim isAbAlreadyExists As Boolean = CBool(ObjectIsExists(classId, {destAbSetId.ToString, abName}))
                Dim destAbId As Integer = -1
                If isAbAlreadyExists Then destAbId = CInt(ObjectId(classId, {destAbSetId.ToString, abName}))

                Dim onRepeat As Integer = 0
                '0 - игнорировать. Повторное наложение способности не происходит (событие AbilityAppyingEvent не возникает).
                '1 - начать заново. Повторное наложение способности не происходит (событие AbilityAppyingEvent не возникает), однако свойство Ability.TurnsCount сбивается на 0.
                '2 - повторить. Наложение способности повторяется (выполняется событие AbilityAppyingEvent).
                '3 - суммировать и повторить. Наложение способности повторяется (выполняется событие AbilityAppyingEvent), при этом Ability.Power суммируется.
                If ReadPropertyInt(classId, "OnRepeat", srcAbSet, AbId, onRepeat, arrProp) = False Then Return "#Error"
                If isAbAlreadyExists Then
                    If onRepeat = 0 Then
                        Return "-1"
                    ElseIf onRepeat = 1 Then
                        If PropertiesRouter(classId, "TurnsCount", {destAbSetId.ToString, destAbId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "0") = "#Error" Then Return "#Error"
                        Return "-1"
                    End If
                End If

                Dim power As String = GetParam(funcParams, 3, "0")
                Dim mage As String = GetParam(funcParams, 4, "-1")

                'Событие AbilityOnApplyEvent
                ''''проверить IsNumeric(res) если дробное
                Dim res As String
                Dim arrs() As String = {srcAbSet.ToString, AbId.ToString, hId.ToString, power, mage, isAbAlreadyExists.ToString}
                Dim eventId As Integer = mScript.mainClass(classId).Properties("AbilityOnApplyEvent").eventId
                'глобальное
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "AbilityOnApplyEvent", False)
                    If res = "False" Then
                        Return "-1"
                    ElseIf IsNumeric(res) Then
                        power = res
                    End If
                End If

                'данного набора
                eventId = mScript.mainClass(classId).ChildProperties(srcAbSet)("AbilityOnApplyEvent").eventId
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "AbilityOnApplyEvent", False)
                    If res = "False" Then
                        Return "-1"
                    ElseIf IsNumeric(res) Then
                        power = res
                    End If
                End If

                'данной способности
                eventId = mScript.mainClass(classId).ChildProperties(srcAbSet)("AbilityOnApplyEvent").ThirdLevelEventId(AbId)
                If eventId > 0 Then
                    res = mScript.eventRouter.RunEvent(eventId, arrs, "AbilityOnApplyEvent", False)
                    If res = "False" Then
                        Return "-1"
                    ElseIf IsNumeric(res) Then
                        power = res
                    End If
                End If

                Dim newAbId As Integer
                If isAbAlreadyExists Then
                    If PropertiesRouter(classId, "TurnsCount", {destAbSetId.ToString, destAbId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, "0") = "#Error" Then Return "#Error"
                    If onRepeat = 3 Then 'суммировать силу и повторить
                        Dim prevPower As Double = 0
                        If ReadPropertyDbl(classId, "Power", destAbSetId, destAbId, prevPower, {destAbSetId.ToString, destAbId.ToString}) = False Then Return "#Error"
                        prevPower += Double.Parse(power, Globalization.NumberStyles.Any, provider_points)
                        power = prevPower.ToString(provider_points)
                        If PropertiesRouter(classId, "Power", {destAbSetId.ToString, destAbId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, power) = "#Error" Then Return "#Error"
                    End If
                    newAbId = destAbId
                Else
                    'копируем способность AbId из srcAbSetId в destAbSetId
                    res = CreateNewObject(classId, {destAbSetId.ToString, abName})
                    If res = "#Error" Then Return res

                    newAbId = CInt(res)
                    For i As Integer = 0 To mScript.mainClass(classId).ChildProperties(srcAbSet).Count - 1
                        Dim pName As String = mScript.mainClass(classId).ChildProperties(srcAbSet).ElementAt(i).Key
                        Dim pValue As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(srcAbSet).ElementAt(i).Value

                        If pName <> "Name" Then
                            If pName = "Power" Then
                                SetPropertyValue(classId, pName, power, destAbSetId, newAbId)
                            ElseIf pName = "Applier" Then
                                SetPropertyValue(classId, pName, mage, destAbSetId, newAbId)
                            Else
                                SetPropertyValue(classId, pName, pValue.ThirdLevelProperties(AbId), destAbSetId, newAbId) 'событие само продублируется при необходимости
                            End If
                        End If
                    Next
                End If
                If PropertiesRouter(classId, "TimeWhenApplied", {destAbSetId.ToString, newAbId.ToString}, arrParams, PropertiesOperationEnum.PROPERTY_SET, GVARS.PLAYER_TIME.ToString) = "#Error" Then Return "#Error"

                If GVARS.G_ISBATTLE Then
                    Dim BattleIcon As String = ""
                    'если идет бой, то получаем иконку способности для размещения возле бойца
                    If ReadProperty(classId, "BattleIcon", destAbSetId, newAbId, BattleIcon, {destAbSetId.ToString, newAbId.ToString}) = False Then Return "#Error"
                    BattleIcon = UnWrapString(BattleIcon)
                    If String.IsNullOrEmpty(BattleIcon) OrElse My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, BattleIcon)) = False Then BattleIcon = ""
                    BattleIcon = BattleIcon.Replace("\"c, "/"c)
                    If String.IsNullOrEmpty(BattleIcon) = False Then
                        'печатаем икону способности
                        If mScript.Battle.ChangeAbilitiesIcons(hId) = "#Error" Then Return "#Error"
                    End If
                End If

                'звук при наложении SoundOnApply
                Dim snd As String = ""
                If ReadProperty(classId, "SoundOnApply", destAbSetId, newAbId, snd, arrProp) = False Then Return "#Error"
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


                Return newAbId.ToString
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function ClassAb_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If questEnvironment.EDIT_MODE Then Return ""

        Dim AbSetId As Integer = -1, abId As Integer = -1
        If IsNothing(funcParams) = False AndAlso funcParams.Count > 1 Then
            If funcParams(0) = "-1" Then Return ""
            AbSetId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If AbSetId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует набора способностей {0}!", funcParams(0))
                Return "#Error"
            End If
            abId = GetThirdChildIdByName(funcParams(1), AbSetId, mScript.mainClass(classId).ChildProperties)
            If abId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует способности {0} в наборе способностей {1}!", funcParams(1), funcParams(0))
                Return "#Error"
            End If
        Else
            Return ""
        End If
        'Здесь только если меняется свойство элемента 2 уровня

        Select Case propertyName
            Case "BattleIcon", "Enabled", "Visible"
                If GVARS.G_ISBATTLE AndAlso newValue <> PREV_VALUE Then
                    Dim BattleIcon As String = ""
                    'если идет бой, то получаем иконку способности для размещения возле бойца
                    If ReadProperty(classId, "BattleIcon", AbSetId, abId, BattleIcon, {AbSetId.ToString, abId.ToString}) = False Then Return "#Error"
                    BattleIcon = UnWrapString(BattleIcon)
                    If String.IsNullOrEmpty(BattleIcon) OrElse My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, BattleIcon)) = False Then BattleIcon = ""
                    BattleIcon = BattleIcon.Replace("\"c, "/"c)
                    If String.IsNullOrEmpty(BattleIcon) = False Then
                        'печатаем икону способности
                        Dim lst As List(Of Integer) = GetHeroesListByAbilitySet(AbSetId)
                        If lst.Count = 0 Then Return ""
                        For i As Integer = 0 To lst.Count - 1
                            If mScript.Battle.ChangeAbilitiesIcons(lst(i)) = "#Error" Then Return "#Error"
                        Next i
                    End If
                End If
        End Select
        Return ""
    End Function
#End Region

#Region "Army"
    Private Function classArmy_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго порядка
                Return CreateNewObject(classId, funcParams)
            Case "Count"
                Return ObjectsCount(classId, funcParams)
            Case "IsExist"
                'Определяет есть ли элемент с таким именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Id"
                Return ObjectId(classId, funcParams)
            Case "Remove"
                'удаление элемента
                If GVARS.G_ISBATTLE Then Return _ERROR("Нельзя удалять армии во время боя.", functionName)
                Return RemoveObject(classId, funcParams)
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function ClassAmy_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If questEnvironment.EDIT_MODE Then Return ""
        Dim armyId As Integer = -1
        If GVARS.G_ISBATTLE AndAlso IsNothing(funcParams) = False AndAlso funcParams.Count = 1 Then
            If funcParams(0) = "-1" Then Return ""
            armyId = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If armyId < -1 Then
                mScript.LAST_ERROR = String.Format("Не существует армии {0}!", funcParams(0))
                Return "#Error"
            End If
        Else
            Return ""
        End If

        If newValue = PREV_VALUE Then Return ""
        'Здесь только если меняется свойство элемента 2 уровня
        Dim Battle As cBattle = mScript.Battle

        Select Case propertyName
            Case "BackgroundClass"
                Dim hArmy As HtmlElement = Battle.FindArmyHTMLContainer(armyId, Nothing, Nothing)
                If IsNothing(hArmy) Then Return ""
                hArmy.SetAttribute("ClassName", UnWrapString(newValue))
            Case "Caption"
                Dim hCaption As HtmlElement = Nothing
                Dim hArmy As HtmlElement = Battle.FindArmyHTMLContainer(armyId, hCaption, Nothing)
                If IsNothing(hArmy) Then Return ""

                If IsNothing(hCaption) Then
                    'Надписи небыло. Создаем
                    hCaption = frmPlayer.wbMain.Document.CreateElement("H1")
                    hCaption.InnerHtml = UnWrapString(newValue)
                    hCaption.SetAttribute("ClassName", "ArmyCaption")

                    hArmy.Children(0).InsertAdjacentElement(HtmlElementInsertionOrientation.BeforeBegin, hCaption)
                Else
                    'надпись уже была. Изменяем ее
                    hCaption.InnerHtml = UnWrapString(newValue)
                End If
            Case "CoatOfArms"
                Dim hCoatOfArms As HtmlElement = Nothing
                Dim hArmy As HtmlElement = Battle.FindArmyHTMLContainer(armyId, Nothing, hCoatOfArms)
                If IsNothing(hArmy) Then Return ""

                If IsNothing(hCoatOfArms) Then
                    'Герба небыло. Создаем
                    Dim aCoat As String = UnWrapString(newValue)
                    aCoat = aCoat.Replace("\"c, "/"c)
                    If String.IsNullOrEmpty(aCoat) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, aCoat)) Then
                        'Создаем герб
                        hCoatOfArms = frmPlayer.wbMain.Document.CreateElement("IMG")
                        hCoatOfArms.SetAttribute("src", aCoat)
                        hCoatOfArms.SetAttribute("ClassName", "ArmyCoatOfArms")

                        hArmy.Children(hArmy.Children.Count - 1).InsertAdjacentElement(HtmlElementInsertionOrientation.BeforeBegin, hCoatOfArms)
                    Else
                        'герба так и нет
                        Return ""
                    End If
                Else
                    'герб уже был. Изменяем его
                    Dim aCoat As String = UnWrapString(newValue)
                    aCoat = aCoat.Replace("\"c, "/"c)
                    If String.IsNullOrEmpty(aCoat) = False AndAlso My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, aCoat)) Then
                        'меняем картинку герба
                        hCoatOfArms.SetAttribute("src", aCoat)
                    Else
                        'удаляем герб
                        Dim msCoat As mshtml.IHTMLDOMNode = hCoatOfArms.DomElement
                        msCoat.removeNode(True)
                    End If
                End If
            Case "FightersPerColumn"
                Dim hArmy As HtmlElement = Battle.FindArmyHTMLContainer(armyId, Nothing, Nothing)
                mScript.Battle.PrepareFightersPosition()
                Dim armyPosId As Integer = -1
                For i As Integer = 0 To mScript.Battle.lstPositions.Count - 1
                    If mScript.Battle.lstPositions(i).armyId = armyId Then
                        armyPosId = i
                        Exit For
                    End If
                Next i
                If armyPosId < 0 Then Return _ERROR("Ошибка при построении расположения бойцов на поле боя.")
                Return mScript.Battle.PutArmyToBattleField(armyPosId, True, hArmy)
            Case "Position" 'справа или слева
                Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                If IsNothing(hDoc) Then Return _ERROR("Не удалось открыть html-документ Battle.html!")

                Dim hArmy As HtmlElement = Battle.FindArmyHTMLContainer(armyId, Nothing, Nothing)
                mScript.Battle.PrepareFightersPosition()
                Dim armyPosId As Integer = -1
                For i As Integer = 0 To mScript.Battle.lstPositions.Count - 1
                    If mScript.Battle.lstPositions(i).armyId = armyId Then
                        armyPosId = i
                        Exit For
                    End If
                Next i
                If armyPosId < 0 Then Return _ERROR("Ошибка при построении расположения бойцов на поле боя.")

                Dim msArmy As mshtml.IHTMLDOMNode = hArmy.DomElement
                msArmy.removeNode(True)

                Dim hParent As HtmlElement
                If mScript.Battle.lstPositions(armyPosId).fromLeft Then
                    hParent = hDoc.GetElementById("LeftArmies")
                Else
                    hParent = hDoc.GetElementById("RightArmies")
                End If
                If IsNothing(hParent) Then Return _ERROR("Ошибка в структуре html-документа!")
                hParent.AppendChild(hArmy)
        End Select
        Return ""
    End Function
#End Region

#Region "Command"
    Private Function classCm_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "Contains"
                Dim hText As HtmlElement = getCommandTextElement()
                If IsNothing(hText) Then Return _ERROR("Не удалось получить доступ к командной строке. Ошибка в структуре документа.", "Command.Contains")
                Dim txt As String = hText.GetAttribute("Value")
                txt = txt.Trim
                If String.IsNullOrEmpty(txt) Then Return "False"

                Dim errCount As Integer = CInt(funcParams(0))
                Dim res As Integer = IndexOfAnyWithErrors(txt, 1, -1, funcParams, arrParams, errCount)
                If res = -1 Then
                    Return "False"
                Else
                    Return "True"
                End If
            Case "Focus"
                Dim hText As HtmlElement = getCommandTextElement()
                If IsNothing(hText) Then Return _ERROR("Не удалось получить доступ к командной строке. Ошибка в структуре документа.", "Command.Contains")
                hText.Focus()
                Return "True"
            Case "Flash"
                Dim hText As HtmlElement = getCommandTextElement()
                If IsNothing(hText) Then Return _ERROR("Не удалось получить доступ к командной строке. Ошибка в структуре документа.", "Command.Contains")
                Dim effDuration As Integer = 1000, effName As String = ""
                If ReadPropertyInt(classId, "FlashEffDuration", -1, -1, effDuration, arrParams) = False Then Return "#Error"
                If ReadProperty(classId, "FlashEffect", -1, -1, effName, arrParams) = False Then Return "#Error"
                effName = UnWrapString(effName)
                hText.Style &= "animation-name:" & effName & ";animation-duration: " & effDuration.ToString & "ms;"
                HTMLReplaceClass(hText, "cmdText", "cmdTextFlash")
                mScript.Battle.Wait(effDuration)
                HTMLReplaceClass(hText, "cmdTextFlash", "cmdText")
                Return True
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function getCommandTextElement() As HtmlElement
        If questEnvironment.EDIT_MODE Then Return Nothing
        Dim hDoc As HtmlDocument = frmPlayer.wbCommand.Document
        If IsNothing(hDoc) Then Return Nothing
        Return hDoc.GetElementById("cmdText")
    End Function


    Private Function ClassCm_PropertiesSet(ByVal classId As Integer, ByVal propertyName As String, ByVal newValue As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        Select Case propertyName
            Case "Visible", "Width", "Height"
                SetWindowsContainers()
        End Select
        Return ""
    End Function
#End Region

#Region "Quest"
    Private Function classQ_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function
#End Region

#Region "Script"
    Private Function classScript_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "SetElementProperties"
                'устанавливает массив свойств
                'получаем класс, элементу которого надо устанавливать свойства
                Dim setClassId As Integer
                If mScript.Param_GetType(funcParams(0)) = MatewScript.ReturnFormatEnum.TO_STRING Then
                    Dim cName As String = mScript.PrepareStringToPrint(funcParams(0), arrParams, True)
                    If cName = "#Error" Then Return cName
                    If mScript.mainClassHash.ContainsKey(cName) Then
                        setClassId = mScript.mainClassHash(cName)
                    Else
                        mScript.LAST_ERROR = "Класс с именем " + cName + " не найден."
                        Return "#Error"
                    End If
                Else
                    setClassId = Val(funcParams(0))
                End If
                'получаем Id элементов 2 и 3 порядков
                Dim child2Id As Integer = GetSecondChildIdByName(funcParams(1), mScript.mainClass(setClassId).ChildProperties)
                Dim child3Id As Integer = GetThirdChildIdByName(funcParams(2), child2Id, mScript.mainClass(setClassId).ChildProperties)
                'считываем в цикле все свойства, переданные в формате "имя_свойства:значение", и устанавливаем их
                For i As Integer = 3 To funcParams.Count - 1
                    Dim f As String = mScript.PrepareStringToPrint(funcParams(i), arrParams, True) 'новое "имя_свойства:значение"
                    If f = "#Error" Then Return f
                    Dim val() As String = Split(f, ":", 2)
                    If IsNothing(val) OrElse val.Count <> 2 Then
                        mScript.LAST_ERROR = "Неверный формат переданных данных. Каждый элемент массива данных должен имять вид свойство:значение."
                        Return "#Error"
                    End If
                    If mScript.mainClass(setClassId).Properties.ContainsKey(val(0)) = False Then
                        'mScript.LAST_ERROR = "В указаннном классе не найдено свойство " + Chr(34) + val(0) + Chr(34) + "."
                        Continue For 'возможно после удаления свойства в классе локаций. Просто пропускаем
                    End If
                    If mScript.mainClass(setClassId).Properties(val(0)).returnType <> MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION AndAlso _
                        mScript.mainClass(setClassId).Properties(val(0)).returnType <> MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                        val(1) = WrapString(val(1))
                    End If
                    SetPropertyValue(setClassId, val(0), val(1), child2Id, child3Id)
                Next
                Return "1"
            Case "Format"
                Dim strFormat As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                If mScript.Param_GetType(funcParams(0)) = MatewScript.ReturnFormatEnum.TO_NUMBER Then
                    Dim val As Single = Convert.ToSingle(funcParams(0), provider_points)
                    Return "'" + Format(val, strFormat).Replace("'", "/'") + "'"
                Else
                    Dim strExpression = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                    If strExpression Like "##.##.####*" Then
                        Dim d As Date = GetDateByString(strExpression)
                        If mScript.LAST_ERROR.Length > 0 Then Return _ERROR("Формат даты не распознан.", functionName)
                        Return "'" + Format(d, strFormat) + "'"
                    Else
                        Return "'" + Format(strExpression, strFormat).Replace("'", "/'") + "'"
                    End If
                End If
            Case "Eval"
                Dim fParams() As String = {}
                Dim strCode As String = funcParams(0)
                If funcParams.Count > 1 Then
                    Array.Resize(fParams, funcParams.Count - 1)
                    Array.ConstrainedCopy(funcParams, 1, fParams, 0, fParams.Count)
                End If
                Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(strCode)
                If cRes = MatewScript.ContainsCodeEnum.CODE OrElse cRes = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                    Return mScript.ExecuteCode(mScript.PrepareBlock(mScript.DeserializeCodeData(strCode)), fParams, True)
                ElseIf cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                    If strCode.StartsWith("'?") Then Return mScript.PrepareStringToPrint(strCode, fParams) 'то есть, запускается со своими параметрами
                Else
                    Return mScript.ExecuteCode(mScript.PrepareBlock(Split(mScript.PrepareStringToPrint(strCode, arrParams), vbNewLine)), fParams, True)
                End If
            Case "Choose"
                Dim chs As Integer = Val(funcParams(0))
                If chs <= 0 Then Return _ERROR("Неправильный номер параметра. Номер не может быть меньше 1.", functionName)
                If chs > funcParams.Count - 1 Then Return _ERROR("Номер параметра больше их общего количества.", functionName)
                Return funcParams(chs)
            Case "IIf"
                Dim res As Boolean = False
                If funcParams(0) = "True" Then res = True
                Return IIf(res, funcParams(1), funcParams(2))
            Case "Call"
                'Функция вызывает функцию пользователя, сохраненную в массиве functionsHash 
                If IsNothing(funcParams) OrElse funcParams.GetUpperBound(0) = -1 Then
                    mScript.LAST_ERROR = "Функция Call вызвана без параметров."
                    Return "#Error"
                End If
                Dim funcName As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                If funcName = "#Error" Then Return funcName
                Dim arrVars As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                mScript.csLocalVariables.CopyVariables(arrVars)
                mScript.csLocalVariables.KillVars()
                Dim curLine As Integer = mScript.CURRENT_LINE
                Dim strResult As String
                If mScript.functionsHash.ContainsKey(funcName) Then
                    Dim fParams() As String
                    ReDim fParams(funcParams.GetUpperBound(0) - 1)
                    Array.ConstrainedCopy(funcParams, 1, fParams, 0, fParams.Length)
                    mScript.codeStack.Push("Функция " + funcName)

                    strResult = mScript.functionsHash(funcName).Run(fParams)
                    'Dim strResult As String = mScript.ExecuteCode(mScript.functionsHash(funcName).ValueExecuteDt, fParams, True)
                    mScript.codeStack.Pop()
                    If strResult = "#Error" Then mScript.LAST_ERROR = "#DON'T_SHOW#" + mScript.LAST_ERROR
                Else
                    mScript.LAST_ERROR = "Функция Писателя " + funcName + " не найдена."
                    Return "#Error"
                End If
                mScript.EXIT_CODE = False
                mScript.CURRENT_LINE = curLine
                mScript.csLocalVariables.RestoreVariables(arrVars)
                Return strResult
            Case "MsgBox"
                If IsNothing(funcParams) OrElse funcParams.GetUpperBound(0) = -1 Then
                    MessageBox.Show("")
                    Return "'Ok'"
                ElseIf funcParams.GetUpperBound(0) = 0 Then
                    MessageBox.Show(mScript.PrepareStringToPrint(funcParams(0), arrParams))
                    Return "'Ok'"
                Else
                    Dim retVal As DialogResult
                    Select Case funcParams(1)
                        Case "'Yes + No'"
                            retVal = MessageBox.Show(mScript.PrepareStringToPrint(funcParams(0), arrParams), "", MessageBoxButtons.YesNo)
                            If retVal = DialogResult.Yes Then
                                Return "'Yes'"
                            Else
                                Return "'No'"
                            End If
                        Case "'Ok + Cancel'"
                            retVal = MessageBox.Show(mScript.PrepareStringToPrint(funcParams(0), arrParams), "", MessageBoxButtons.OKCancel)
                            If retVal = DialogResult.OK Then
                                Return "'Ok'"
                            Else
                                Return "'Cancel'"
                            End If
                        Case "'Yes + No + Cancel'"
                            retVal = MessageBox.Show(mScript.PrepareStringToPrint(funcParams(0), arrParams), "", MessageBoxButtons.YesNoCancel)
                            If retVal = DialogResult.Yes Then
                                Return "'Yes'"
                            ElseIf retVal = DialogResult.No Then
                                Return "'No'"
                            Else
                                Return "'Cancel'"
                            End If
                        Case Else
                            retVal = MessageBox.Show(mScript.PrepareStringToPrint(funcParams(0), arrParams), "", MessageBoxButtons.OK)
                            Return "'Ok'"
                    End Select
                End If
            Case "AddFunction"
                Return AddUserFunction(funcParams, arrParams)
            Case "RemoveFunction"
                Return RemoveUserFunction(funcParams, arrParams)
            Case "AddProperty"
                Return AddUserProperty(funcParams, arrParams)
            Case "RemoveProperty"
                Return RemoveUserProperty(funcParams, arrParams)
            Case "IsFunctionExists"
                'funcParams(0) - класс функции
                Dim className As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                If className = "#Error" Then Return "#Error"
                Dim aClassId As Integer = -1
                If Not mScript.mainClassHash.TryGetValue(className, aClassId) Then
                    Return _ERROR("Не удалось получить класс в который надо добавить функцию.", "IsFunctionExists")
                End If
                Dim aName As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                If aName = "#Error" Then Return aName
                If mScript.mainClass(aClassId).Functions.ContainsKey(aName) Then
                    Return "True"
                Else
                    Return "False"
                End If
            Case "IsPropertyExists"
                'funcParams(0) - класс свойства
                Dim className As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                If className = "#Error" Then Return "#Error"
                Dim aClassId As Integer = -1
                If Not mScript.mainClassHash.TryGetValue(className, aClassId) Then
                    Return _ERROR("Не удалось получить класс в который надо добавить функцию.", "IsFunctionExists")
                End If
                Dim aName As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                If aName = "#Error" Then Return aName
                If mScript.mainClass(classId).Properties.ContainsKey(aName) Then
                    Return "True"
                Else
                    Return "False"
                End If
            Case "SetProperty"
                'funcParams(0) - класс свойства
                Dim className As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                If className = "#Error" Then Return "#Error"
                Dim aClassId As Integer = -1
                If Not mScript.mainClassHash.TryGetValue(className, aClassId) Then
                    Return _ERROR("Не удалось получить класс в который надо добавить функцию.", "IsFunctionExists")
                End If
                Dim propName As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                If propName = "#Error" Then Return propName
                If Not mScript.mainClass(aClassId).Properties.ContainsKey(propName) Then
                    Dim res As String = AddUserProperty(funcParams, arrParams)
                    If res = "#Error" Then Return res
                End If
                Dim setValue As String = ""
                If funcParams.Count > 2 Then setValue = funcParams(2)

                SetPropertyValue(aClassId, propName, setValue, -1)
                If mScript.mainClass(aClassId).LevelsCount > 0 AndAlso IsNothing(mScript.mainClass(aClassId).ChildProperties) = False Then
                    For i As Integer = 0 To mScript.mainClass(aClassId).ChildProperties.Count - 1
                        SetPropertyValue(aClassId, propName, setValue, i)
                        If mScript.mainClass(aClassId).LevelsCount = 2 Then
                            If IsNothing(mScript.mainClass(aClassId).ChildProperties(i)(propName).ThirdLevelProperties) = False Then
                                For j As Integer = 0 To mScript.mainClass(aClassId).ChildProperties(i)(propName).ThirdLevelProperties.Count - 1
                                    SetPropertyValue(aClassId, propName, setValue, i, j)
                                Next j
                            End If
                        End If
                    Next i
                End If
            Case "RemoveEvent"
                Dim fullName As String = funcParams(0)
                Dim eventClass As Integer = 0, propName As String = "", isTracking As frmMainEditor.trackingcodeEnum
                'получаем свойство из переданной строки. Если свойство не найдено - выход
                If GetClassIdAndElementNameByString(fullName, eventClass, propName, False, True, False, isTracking, funcParams) = False Then Return ""

                If isTracking <> frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT Then
                    'удаляется событие отслеживания
                    mScript.trackingProperties.RemoveProperty(eventClass, propName, isTracking)
                    Return ""
                End If

                'получаем Id элементов 2 и 3 порядков
                Dim child2Name As String = GetParam(funcParams, 1, "-1")
                Dim child2Id As Integer = GetSecondChildIdByName(child2Name, mScript.mainClass(eventClass).ChildProperties)
                Dim child3Name As String = GetParam(funcParams, 2, "-1")
                Dim child3Id As Integer = GetThirdChildIdByName(child3Name, child2Id, mScript.mainClass(eventClass).ChildProperties)
                Dim eventId As Integer

                'получаем eventId, удаляем из mainClass
                If child2Id < 0 Then
                    eventId = mScript.mainClass(eventClass).Properties(propName).eventId
                    mScript.mainClass(eventClass).Properties(propName).eventId = 0
                ElseIf child3Id < 0 Then
                    eventId = mScript.mainClass(eventClass).ChildProperties(child2Id)(propName).eventId
                    mScript.mainClass(eventClass).ChildProperties(child2Id)(propName).eventId = 0
                Else
                    eventId = mScript.mainClass(eventClass).ChildProperties(child2Id)(propName).ThirdLevelEventId(child3Id)
                    mScript.mainClass(eventClass).ChildProperties(child2Id)(propName).ThirdLevelEventId(child3Id) = 0
                End If

                'непосредственно удаляем событие
                mScript.eventRouter.RemoveEvent(eventId)
                Return ""
            Case "Wait"
                'Приостанавливает на время выполнение кода
                Dim interval As Integer = Val(GetParam(funcParams, 0, "1000"))
                If interval <= 0 Then Return ""
                Dim wType As Integer = Val(GetParam(funcParams, 1, "0"))
                Select Case wType
                    Case 2 'полный стоп
                        System.Threading.Thread.Sleep(interval)
                    Case Else  '0 - мягкий стоп; 1 - мягкий стоп с таймерами
                        Dim initMilliseconds As Long = New TimeSpan(Date.Now.Ticks).TotalMilliseconds
                        Dim curMilliseconds As Long = New TimeSpan(Date.Now.Ticks).TotalMilliseconds - initMilliseconds
                        Dim prevHoldValue As Boolean = GVARS.HOLD_UP_TIMERS
                        If wType = 1 Then GVARS.HOLD_UP_TIMERS = True

                        Do While curMilliseconds < interval
                            curMilliseconds = New TimeSpan(Date.Now.Ticks).TotalMilliseconds - initMilliseconds
                            Application.DoEvents()
                        Loop
                        If wType = 1 Then GVARS.HOLD_UP_TIMERS = prevHoldValue
                End Select
                Return ""
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

#End Region

#Region "String"

    Private Function classS_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "IsNum"
                Dim res As Boolean = True
                For i As Integer = 0 To funcParams.Count - 1
                    res = IsNumeric(UnWrapString(funcParams(i)).Replace(".", ",")).ToString
                    If res = False Then Return "False"
                Next i
                Return res.ToString
            Case "Empty"
                Return ""
            Case "CharByCode"
                Return WrapString(Chr(CInt(funcParams(0))))
            Case "CodebyChar"
                Return Asc(mScript.PrepareStringToPrint(funcParams(0), arrParams)).ToString
            Case "StartsWith"
                Dim strFull As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim strFragment As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                Dim iCase As Boolean = True
                If funcParams.Count > 2 Then Boolean.TryParse(funcParams(2), iCase)
                Dim res As Boolean = False
                If iCase Then
                    res = strFull.StartsWith(strFragment, StringComparison.CurrentCultureIgnoreCase)
                Else
                    res = strFull.StartsWith(strFragment)
                End If
                Return res.ToString

            Case "EndsWith"
                Dim strFull As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim strFragment As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                Dim iCase As Boolean = True
                If funcParams.Count > 2 Then Boolean.TryParse(funcParams(2), iCase)
                Dim res As Boolean = False
                If iCase Then
                    res = strFull.EndsWith(strFragment, StringComparison.CurrentCultureIgnoreCase)
                Else
                    res = strFull.EndsWith(strFragment)
                End If
                Return res.ToString
            Case "Format"
                Dim strFormat As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                If mScript.Param_GetType(funcParams(0)) = MatewScript.ReturnFormatEnum.TO_NUMBER Then
                    Dim val As Single = Convert.ToSingle(funcParams(0), provider_points)
                    Return "'" + Format(val, strFormat).Replace("'", "/'") + "'"
                Else
                    Dim strExpression = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                    If strExpression Like "##.##.####*" Then
                        Dim d As Date = GetDateByString(strExpression)
                        If mScript.LAST_ERROR.Length > 0 Then Return _ERROR("Формат даты не распознан.", functionName)
                        Return "'" + Format(d, strFormat) + "'"
                    Else
                        Return "'" + Format(strExpression, strFormat).Replace("'", "/'") + "'"
                    End If
                End If
            Case "Replace"
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim strSeek As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                Dim strReplace As String = mScript.PrepareStringToPrint(funcParams(2), arrParams)
                Dim start As Integer = 1
                If funcParams.Count > 3 Then start = Val(funcParams(3))
                Dim rCount As Integer = -1
                If funcParams.Count > 4 Then rCount = Val(funcParams(4))
                If rCount < -1 Then rCount = -1

                If strSeek.Length = 0 Then Return _ERROR("строка поиска не может быть пустой", functionName)
                If start < 0 Then Return _ERROR("первый символ поиска не может быть отрицательным или равен нулю.", functionName)
                If start + 1 > strWork.Length Then Return _ERROR("первый символ поиска находится за пределами длины строки", functionName)

                Dim strResult As String = ""
                strResult = Replace(strWork, strSeek, strReplace, start, rCount)
                If start > 1 Then strResult = Left(strWork, start - 1) + strResult
                Return "'" + strResult.Replace("'", "/'") + "'"
            Case "InStrRev"
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim start As Integer = strWork.Length - 1
                If funcParams.Count > 2 Then start = Val(funcParams(2))
                If start = -1 Then start = strWork.Length - 1

                If start < 0 Then Return _ERROR("первый символ поиска не может быть отрицательным или равен нулю.", functionName)
                If start + 1 > strWork.Length Then Return _ERROR("первый символ поиска находится за пределами длины строки", functionName)
                Dim strSeek As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                If strSeek.Length = 0 Then Return _ERROR("строка поиска не может быть пустой", functionName)
                Dim iCase As Boolean = True
                If funcParams.Count > 3 Then Boolean.TryParse(funcParams(3), iCase)
                If iCase Then
                    Return strWork.LastIndexOf(strSeek, start, StringComparison.CurrentCultureIgnoreCase)
                Else
                    Return strWork.LastIndexOf(strSeek, start).ToString
                End If
            Case "InStr"
                Dim start As Integer = 0
                If funcParams.Count > 2 Then start = Val(funcParams(2))
                If start < 0 Then Return _ERROR("первый символ поиска не может быть отрицательным или равен нулю.", functionName)
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                If start + 1 > strWork.Length Then Return _ERROR("первый символ поиска находится за пределами длины строки", functionName)
                Dim strSeek As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                If strSeek.Length = 0 Then Return _ERROR("строка поиска не может быть пустой", functionName)
                Dim iCase As Boolean = True
                If funcParams.Count > 3 Then Boolean.TryParse(funcParams(3), iCase)
                If iCase Then
                    Return strWork.IndexOf(strSeek, start, StringComparison.CurrentCultureIgnoreCase)
                Else
                    Return strWork.IndexOf(strSeek, start).ToString
                End If
            Case "InStrRevAny"
                '0 - исходная строка, 1 - массив, 2 - начало, 3 - конец, 4 - игнорировать регистр
                Dim start As Integer = -1
                If funcParams.Count > 2 Then start = Val(funcParams(2))
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim iCase As Boolean = True
                If funcParams.Count > 4 Then Boolean.TryParse(funcParams(4), iCase)
                Dim varName As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                If String.IsNullOrEmpty(varName) Then Return _ERROR("Не указан массив для поиска.", functionName)
                Dim cVar As cVariable.variableEditorInfoType = Nothing
                If mScript.csLocalVariables.lstVariables.TryGetValue(varName, cVar) = False Then
                    If mScript.csPublicVariables.lstVariables.TryGetValue(varName, cVar) = False Then
                        Return _ERROR("Переменная " & " не найдена.", functionName)
                    End If
                End If
                Dim arrValues() As String = cVar.arrValues
                If IsNothing(arrValues) OrElse arrValues.Count = 0 Then Return -1

                If start < 0 Then start = arrValues.Count - 1
                If start >= arrValues.Length Then Return _ERROR("первый элемент массива для поиска находится за диапазона.", functionName)
                Dim final As Integer = 0
                If funcParams.Count > 3 Then Integer.TryParse(funcParams(3), final)
                If final >= arrValues.Length Then Return _ERROR("последний элемент массива для поиска находится за диапазона.", functionName)
                If start < final Then Return _ERROR("Индекс первого элемента массива для поиска в обратном порядке должен быть больше последнего.", functionName)

                Dim compType As StringComparison = StringComparison.CurrentCultureIgnoreCase
                If Not iCase Then compType = StringComparison.CurrentCulture

                For i As Integer = start To final Step -1
                    Dim strSeek As String = mScript.PrepareStringToPrint(arrValues(i), arrParams)
                    If strSeek.Length = 0 Then Continue For
                    If strWork.IndexOf(strSeek, compType) >= 0 Then Return i.ToString
                Next
                Return "-1"
            Case "InStrAny"
                '0 - исходная строка, 1 - массив, 2 - начало, 3 - конец, 4 - игнорировать регистр
                Dim start As Integer = 0
                If funcParams.Count > 2 Then start = Val(funcParams(2))
                If start < 0 Then Return _ERROR("первый символ поиска не может быть отрицательным или равен нулю.", functionName)
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim iCase As Boolean = True
                If funcParams.Count > 4 Then Boolean.TryParse(funcParams(4), iCase)
                Dim varName As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                If String.IsNullOrEmpty(varName) Then Return _ERROR("Не указан массив для поиска.", functionName)
                Dim cVar As cVariable.variableEditorInfoType = Nothing
                If mScript.csLocalVariables.lstVariables.TryGetValue(varName, cVar) = False Then
                    If mScript.csPublicVariables.lstVariables.TryGetValue(varName, cVar) = False Then
                        Return _ERROR("Переменная " & " не найдена.", functionName)
                    End If
                End If
                Dim arrValues() As String = cVar.arrValues
                If start >= arrValues.Length Then Return _ERROR("первый элемент массива для поиска находится за диапазона.", functionName)
                If IsNothing(arrValues) OrElse arrValues.Count = 0 Then Return -1
                Dim final As Integer = -1
                If funcParams.Count > 3 Then Integer.TryParse(funcParams(3), final)
                If final = -1 OrElse final > arrValues.Count - 1 Then final = arrValues.Count - 1

                Dim compType As StringComparison = StringComparison.CurrentCultureIgnoreCase
                If Not iCase Then compType = StringComparison.CurrentCulture

                For i As Integer = start To final
                    Dim strSeek As String = mScript.PrepareStringToPrint(arrValues(i), arrParams)
                    If strSeek.Length = 0 Then Continue For
                    If strWork.IndexOf(strSeek, compType) >= 0 Then Return i.ToString
                Next
                Return "-1"
            Case "StrFromCount"
                Dim cnt As Single = Val(funcParams(0))
                If cnt <> Math.Round(cnt) Then Return funcParams(2) 'дробное число: 2,4 копейки
                If funcParams(0).EndsWith("1") Then
                    If funcParams(0).EndsWith("11") Then Return funcParams(3)
                    Return funcParams(1) '21 копейка
                End If
                Dim last As Byte = CByte(Right(cnt.ToString, 1))
                If last <= 4 AndAlso last > 0 Then Return funcParams(2) '3 копейки
                Return funcParams(3) '7 копеек
            Case "Like"
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim pattern As String = mScript.PrepareStringToPrint(funcParams(1), arrParams, False)
                Dim res As Boolean = strWork Like pattern
                If res Then Return "True"
                Return "False"
            Case "Left", "Right"
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim lPos As Integer = Val(funcParams(1))
                If strWork.Length = 0 Then Return _ERROR("указана пустая строка.", functionName)
                If lPos < 1 Then Return _ERROR("число символов указано неправильно (равно нулю или отрицательно).", functionName)
                If lPos > strWork.Length Then Return _ERROR("длина строки меньше, чем указанное число символов.", functionName)
                If functionName = "Left" Then
                    strWork = "'" + Left(strWork, lPos).Replace("'", "/'") + "'"
                Else
                    strWork = "'" + Right(strWork, lPos).Replace("'", "/'") + "'"
                End If
                Return strWork
            Case "Mid"
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim mPos As Integer = Val(funcParams(1))
                If strWork.Length = 0 Then Return _ERROR("указана пустая строка.", functionName)
                If mPos < 1 Then Return _ERROR("первый символ указан неправильно (равен нулю или отрицательный).", functionName)
                If funcParams.Count < 3 Then
                    strWork = "'" + Mid(strWork, mPos).Replace("'", "/'") + "'"
                Else
                    Dim mCount As Integer = Val(funcParams(2))
                    If mPos + mCount > strWork.Length Then Return _ERROR("длина строки меньше, чем указанное число символов.", functionName)
                    strWork = "'" + Mid(strWork, mPos, mCount).Replace("'", "/'") + "'"
                End If
                Return strWork
            Case "UCase"
                Return funcParams(0).ToUpper
            Case "LCase"
                Return funcParams(0).ToLower
            Case "LTrim", "RTrim", "Trim"
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                If functionName = "LTrim" Then
                    strWork = strWork.TrimStart
                ElseIf functionName = "RTrim" Then
                    strWork = strWork.TrimEnd
                Else
                    strWork = strWork.Trim
                End If
                Return "'" + strWork.Replace("'", "/'") + "'"
                'Case "ToNumber"
                '    Dim strVal As String = funcParams(0)
                '    strVal = mScript.PrepareStringToPrint(strVal, arrParams).Replace(".", ",")
                '    If IsNumeric(strVal) Then
                '        Return Convert.ToSingle(strVal).ToString.Replace(",", ".")
                '    Else
                '        Return _ERROR("нельзя преобразовать данную строку в число.", functionName)
                '    End If
            Case "ToBool"
                Dim strVal As String = funcParams(0)
                If strVal.Length = 0 OrElse strVal = "''" Then Return "False"
                Return "True"
            Case "Join"
                Dim varName = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim cArr As cVariable.variableEditorInfoType = Nothing
                If mScript.csLocalVariables.lstVariables.TryGetValue(varName, cArr) = False Then
                    If mScript.csPublicVariables.lstVariables.TryGetValue(varName, cArr) = False Then
                        Return _ERROR("ошибка при поиске массива переменной " + varName + ". Переменная не найдена.", functionName)
                    End If
                End If
                Dim arrValues() As String = cArr.arrValues
                If IsNothing(arrValues) OrElse arrValues.Count = 0 Then Return ""

                Dim strSeparator As String = " "
                If funcParams.Count > 1 Then strSeparator = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                Dim sBuilder As New System.Text.StringBuilder
                For i As Integer = 0 To arrValues.Count - 1
                    sBuilder.Append(mScript.PrepareStringToPrint(arrValues(i), arrParams))
                    If i < arrValues.Count - 1 Then sBuilder.Append(strSeparator)
                Next
                Return WrapString(sBuilder.ToString)
            Case "Split"
                Dim varName = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim strWork As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                Dim strSplitter As String = " "
                If funcParams.Count > 2 Then strSplitter = mScript.PrepareStringToPrint(funcParams(2), arrParams)
                Dim sCount As Integer = 0
                If funcParams.Count > 3 Then sCount = Val(funcParams(3))
                Dim arrValues() As String = {}
                If sCount = 0 Then
                    arrValues = Split(strWork, strSplitter)
                Else
                    arrValues = Split(strWork, strSplitter, sCount)
                End If
                If IsNothing(arrValues) = False AndAlso arrValues.Count > 0 Then
                    For i As Integer = 0 To arrValues.Count - 1
                        arrValues(i) = "'" + arrValues(i).Replace("'", "/'") + "'"
                    Next
                End If
                If mScript.csLocalVariables.SetVariableArray(varName, arrValues) = "#Error" Then
                    If mScript.csPublicVariables.SetVariableArray(varName, arrValues) = "#Error" Then
                        Return _ERROR("ошибка при установке массива переменной " + varName + ". Переменная не найдена.", functionName)
                    End If
                End If
            Case "Len"
                Return mScript.PrepareStringToPrint(funcParams(0), arrParams).Length.ToString
            Case "Substring"
                Dim strSrc As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim pos As Integer = CInt(funcParams(1))
                If funcParams.Length = 2 Then
                    Return WrapString(strSrc.Substring(pos))
                Else
                    Dim l As Integer = CInt(funcParams(2))
                    Return WrapString(strSrc.Substring(pos, l))
                End If
            Case "Compare"
                Dim str1 As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim str2 As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                Dim iSpaces As Boolean = True
                If funcParams.Length > 2 Then iSpaces = CBool(funcParams(2))
                Dim errCount As Integer = 0
                If funcParams.Count > 3 Then Integer.TryParse(funcParams(3), errCount)
                If iSpaces Then
                    str1 = str1.Trim
                    str2 = str2.Trim
                End If
                If errCount = 0 Then
                    Dim res As Integer = String.Compare(str1, str2, True)
                    If res = 0 Then Return "True"
                    Return "False"
                Else
                    Dim res As Integer = IndexOfAnyWithErrors(str1, 0, 0, {str2}, arrParams, errCount)
                    If res = -1 Then Return "False"
                    Return "True"
                End If
                Return "False"
            Case "Contains"
                Dim strSrc As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim strCompare As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
                Dim iCase As Boolean = True
                If funcParams.Length > 2 Then iCase = CBool(funcParams(2))
                If iCase Then
                    strSrc = strSrc.ToLower
                    strCompare = strCompare.ToLower
                End If
                Return strSrc.Contains(strCompare).ToString
            Case "Reverse"
                Dim strSrc As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim strDest As String = "" '= strSrc.Reverse(strSrc)
                'Dim arrChars() As String = strSrc.Split("").ToList.AsEnumerable().Reverse().ToArray
                'ReDim arrChars(strSrc.Length - 1)
                Dim l As Integer = strSrc.Length - 1
                For i As Integer = 0 To strSrc.Length - 1
                    strDest += strSrc.Chars(l - i)
                Next
                Return WrapString(strDest)
            Case "PadLeft"
                Dim strSrc As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim l As Integer = CInt(funcParams(1))
                strSrc = strSrc.PadLeft(l)
                Return WrapString(strSrc)
            Case "PadRight"
                Dim strSrc As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim l As Integer = CInt(funcParams(1))
                strSrc = strSrc.PadRight(l)
                Return WrapString(strSrc)
            Case "Insert"
                Dim strSrc As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim pos As Integer = CInt(funcParams(1))
                Dim strFr As String = mScript.PrepareStringToPrint(funcParams(2), arrParams)
                Return WrapString(strSrc.Insert(pos, strFr))
            Case "Remove"
                Dim strSrc As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim pos As Integer = CInt(funcParams(1))
                Dim l As Integer = CInt(funcParams(2))
                strSrc = strSrc.Remove(pos, l)
                Return WrapString(strSrc)
            Case "ToNumber"
                Dim strVal As String = UnWrapString(funcParams(0)).Replace(",", ".")
                If strVal = "True" Then
                    Return "1"
                ElseIf strVal = "False" Then
                    Return "0"
                Else
                    Return Val(strVal).ToString(provider_points)
                End If
            Case "IsEmpty"
                Dim res As String = funcParams(0)
                If String.IsNullOrEmpty(res) Then
                    Return "True"
                Else
                    Return "False"
                End If
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

#End Region

#Region "File"
    Private Function classFile_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "IsFileExist, LoadFile, SaveFile, FileLen, FileDate, FileCopy"
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function
#End Region

#End Region
    Private Function classUser_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"

        Select Case functionName
            Case "Create"
                'Создает в указанном классе объект второго или тетьего порядка
                Return CreateNewObject(classId, funcParams)
            Case "Remove"
                'Удаляет объект(-ы)
                Return RemoveObject(classId, funcParams)
            Case "Id"
                'Получаем Id объекта
                Return ObjectId(classId, funcParams)
            Case "IsExist"
                'получаем существует ли объект с заданным именем/Id
                Return ObjectIsExists(classId, funcParams)
            Case "Count"
                'возвращает количество объектов 2 или 3 порядка
                Return ObjectsCount(classId, funcParams)
            Case Else
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(classId).Functions.TryGetValue(functionName, f) = False Then
                    mScript.LAST_ERROR = String.Format("Функции {0} в классе {1} не существует!", functionName, mScript.mainClass(classId).Names.Last)
                    Return "#Error"
                End If
                If mScript.mainClass(classId).Functions(functionName).eventId > 0 Then
                    Return mScript.eventRouter.RunEvent(mScript.mainClass(classId).Functions(functionName).eventId, funcParams, functionName, False)
                Else
                    Return mScript.CallFunction(mScript.mainClass(classId).Functions(functionName).Value, funcParams)
                End If
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    Private Function classArr_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        'Функции не проверенные

        Select Case functionName
            Case "Add"
                '0 varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName, True)

                '1 varValue
                Dim varValue As String = funcParams(1)
                '2 varSign
                Dim varSign As String = UnWrapString(GetParam(funcParams, 2, ""))

                Dim id As Integer
                If var.arrValues.Count = 1 AndAlso var.arrValues(0) = "" Then
                    id = 0
                Else
                    id = var.arrValues.Length
                    ReDim Preserve var.arrValues(id)
                End If
                var.arrValues(id) = varValue

                If String.IsNullOrEmpty(varSign) = False Then
                    If IsNothing(var.lstSingatures) Then var.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                    var.lstSingatures.Add(varSign, id)
                End If

                Return id.ToString
            Case "Array"
                Dim var As New cVariable.variableEditorInfoType
                ReDim var.arrValues(funcParams.Count - 1)
                Array.ConstrainedCopy(funcParams, 0, var.arrValues, 0, funcParams.Length)
                mScript.lastArray = var
                Return "#ARRAY"
            Case "ChangeSignature"
                '0 varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '2 new sign
                Dim newSign As String = UnWrapString(funcParams(2))
                '1 element
                Dim el As String = funcParams(1)
                If IsNumeric(el) Then
                    'изменяем сигнатуру по Id
                    If IsNothing(var.lstSingatures) Then var.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                    Dim id As Integer = CInt(el)
                    Dim pos As Integer = -1
                    pos = var.lstSingatures.IndexOfValue(id)
                    If pos = -1 Then
                        If id > var.arrValues.Count - 1 Then Return _ERROR("Индекс за пределами диапазона.", functionName)
                        var.lstSingatures.Add(newSign, id)
                    Else
                        var.lstSingatures.RemoveAt(pos)
                        var.lstSingatures.Add(newSign, id)
                    End If
                Else
                    'изменяем сигнатуру по старой сигнатуре
                    Dim oldSign As String = UnWrapString(el)
                    If IsNothing(var.lstSingatures) OrElse var.lstSingatures.ContainsKey(oldSign) = False Then
                        Return _ERROR("Сигнатуры " & oldSign & " в массиве " & varName & " не существует.", functionName)
                    End If
                    Dim id As Integer = var.lstSingatures(oldSign)
                    var.lstSingatures.Remove(oldSign)
                    var.lstSingatures.Add(newSign, id)
                End If
                Return ""
            Case "Clear"
                '0 varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                If IsNothing(var.lstSingatures) = False Then var.lstSingatures.Clear()
                ReDim var.arrValues(0)
                var.arrValues(0) = ""
                Return ""
            Case "ClearSignatures"
                '0 varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                If IsNothing(var.lstSingatures) = False Then var.lstSingatures.Clear()
                Return "#ARRAY"
            Case "Copy"
                '0 varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim src As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(src) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                Dim start As Integer = 0, final As Integer = -1
                Dim El As String = GetParam(funcParams, 1, "0")
                start = cVariable.GetElementIndex(src, El)

                El = GetParam(funcParams, 2, "-1")
                If El = "-1" Then
                    final = src.arrValues.Count - 1
                Else
                    final = cVariable.GetElementIndex(src, El)
                End If

                'копируем значения
                Dim length As Integer = final - start + 1
                If length <= 0 Then Return _ERROR("Начальный индекс копирования больше или равен конечному.", functionName)
                mScript.lastArray = New cVariable.variableEditorInfoType
                ReDim mScript.lastArray.arrValues(length - 1)
                Array.ConstrainedCopy(src.arrValues, start, mScript.lastArray.arrValues, 0, length)
                'копируем сигнатуры
                If IsNothing(src.lstSingatures) = False AndAlso src.lstSingatures.Count > 0 Then
                    For i As Integer = 0 To src.lstSingatures.Count - 1
                        Dim id As Integer = src.lstSingatures.ElementAt(i).Value
                        If id < start OrElse id > final Then Continue For
                        If IsNothing(mScript.lastArray.lstSingatures) Then mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                        mScript.lastArray.lstSingatures.Add(src.lstSingatures.ElementAt(i).Key, id - start)
                    Next
                End If
                Return "#ARRAY"
            Case "CopySignatures"
                '0 varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim src As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(src) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                Dim start As Integer = 0, final As Integer = -1
                Dim El As String = GetParam(funcParams, 1, "0")
                start = cVariable.GetElementIndex(src, El)

                El = GetParam(funcParams, 2, "-1")
                If El = "-1" Then
                    final = src.arrValues.Count - 1
                Else
                    final = cVariable.GetElementIndex(src, El)
                End If

                'копируем сигнатуры
                mScript.lastArray = New cVariable.variableEditorInfoType
                If IsNothing(src.lstSingatures) OrElse src.lstSingatures.Count = 0 Then
                    'сигнатур нет
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = ""
                Else

                    mScript.lastArray.arrValues = src.lstSingatures.Keys.ToArray
                    'заворачиваем в ''
                    For i As Integer = 0 To mScript.lastArray.arrValues.Count - 1
                        mScript.lastArray.arrValues(i) = WrapString(mScript.lastArray.arrValues(i))
                    Next
                End If
                Return "#ARRAY"
            Case "CopyValues"
                '0 varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim src As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(src) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                Dim start As Integer = 0, final As Integer = -1
                Dim El As String = GetParam(funcParams, 1, "0")
                start = cVariable.GetElementIndex(src, El)

                El = GetParam(funcParams, 2, "-1")
                If El = "-1" Then
                    final = src.arrValues.Count - 1
                Else
                    final = cVariable.GetElementIndex(src, El)
                End If

                'копируем значения
                Dim length As Integer = final - start + 1
                If length <= 0 Then Return _ERROR("Начальный индекс копирования больше или равен конечному.", functionName)
                mScript.lastArray = New cVariable.variableEditorInfoType
                ReDim mScript.lastArray.arrValues(length - 1)
                Array.ConstrainedCopy(src.arrValues, start, mScript.lastArray.arrValues, 0, length)
                Return "#ARRAY"
            Case "Count", "CountOfSignatures"
                '0 varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim src As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(src) Then Return "0"
                If functionName = "Count" Then
                    If src.arrValues.Count = 1 AndAlso src.arrValues(0) = "" Then Return 0
                    Return src.arrValues.Count
                Else
                    If IsNothing(src.lstSingatures) Then Return "0"
                    Return src.lstSingatures.Count
                End If
            Case "ExcludeSignatures"
                '0 src varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim src As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(src) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '1 exclusion varName
                varName = UnWrapString(funcParams(1))
                Dim exc As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(exc) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                'reserve
                Dim blnReverse As Boolean = CBool(GetParam(funcParams, 2, "False"))

                mScript.lastArray = New cVariable.variableEditorInfoType
                If IsNothing(src.lstSingatures) OrElse src.lstSingatures.Count = 0 Then
                    'передан пустой массив
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = ""
                    Return "#ARRAY"
                ElseIf exc.arrValues.Count = 0 Then
                    'исключения пустые - возвращаем весь массив 
                    If blnReverse Then
                        ReDim mScript.lastArray.arrValues(0)
                        mScript.lastArray.arrValues(0) = ""
                    Else
                        mScript.lastArray = src
                    End If
                    Return "#ARRAY"
                End If

                'поиск исключений
                'Array.ConstrainedCopy(src.arrValues, 0, mScript.lastArray.arrValues, 0, src.arrValues.Length)
                Dim lstValues As List(Of String)
                mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                Dim lstSign As List(Of String) = src.lstSingatures.Keys.ToList()

                If blnReverse Then
                    lstValues = New List(Of String)
                    For i As Integer = 0 To exc.arrValues.Count - 1
                        Dim curSign As String = UnWrapString(exc.arrValues(i))
                        Dim pos As Integer = lstSign.IndexOf(curSign)
                        If pos > -1 Then
                            lstValues.Add(src.arrValues(src.lstSingatures.ElementAt(pos).Value))
                            mScript.lastArray.lstSingatures.Add(curSign, lstValues.Count - 1)
                        End If
                    Next i
                Else
                    lstValues = src.arrValues.ToList
                    For i As Integer = 0 To src.lstSingatures.Count - 1
                        mScript.lastArray.lstSingatures.Add(src.lstSingatures.ElementAt(i).Key, src.lstSingatures.ElementAt(i).Value)
                    Next

                    For i As Integer = 0 To exc.arrValues.Count - 1
                        Dim curSign As String = UnWrapString(exc.arrValues(i))
                        Dim pos As Integer = lstSign.IndexOf(curSign)
                        If pos > -1 Then
                            lstValues.RemoveAt(src.lstSingatures.ElementAt(pos).Value)
                            mScript.lastArray.lstSingatures.Remove(curSign)
                        End If
                    Next i
                End If
                mScript.lastArray.arrValues = lstValues.ToArray
                lstValues.Clear()
                lstSign.Clear()

                Return "#ARRAY"
            Case "ExcludeValues"
                '0 src varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim src As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(src) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '1 exclusion varName
                varName = UnWrapString(funcParams(1))
                Dim exc As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(exc) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                'reserve
                Dim blnReverse As Boolean = CBool(GetParam(funcParams, 2, "False"))

                mScript.lastArray = New cVariable.variableEditorInfoType
                If src.arrValues.Count = 0 Then
                    'передан пустой массив
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = ""
                    Return "#ARRAY"
                ElseIf exc.arrValues.Count = 0 Then
                    'исключения пустые - возвращаем весь массив
                    'Array.ConstrainedCopy(src.arrValues, 0, mScript.lastArray.arrValues, 0, src.arrValues.Length)
                    If blnReverse Then
                        ReDim mScript.lastArray.arrValues(0)
                        mScript.lastArray.arrValues(0) = ""
                    Else
                        mScript.lastArray = src
                    End If
                    Return "#ARRAY"
                End If

                'поиск исключений
                'ReDim mScript.lastArray.arrValues(src.arrValues.Count - 1)
                Dim lstDest As New List(Of String)
                For i As Integer = 0 To src.arrValues.Count - 1
                    Dim blnFound As Boolean = False
                    Dim curValue As String = src.arrValues(i)
                    For j As Integer = 0 To exc.arrValues.Count - 1
                        If String.Compare(curValue, exc.arrValues(j), True) = 0 Then
                            blnFound = True
                            Exit For
                        End If
                    Next j

                    If blnFound = blnReverse Then
                        lstDest.Add(curValue)
                        'добавляем сигнатуру
                        If IsNothing(src.lstSingatures) = False Then
                            Dim pos As Integer = src.lstSingatures.IndexOfValue(i)
                            If pos > -1 Then
                                If IsNothing(mScript.lastArray.lstSingatures) Then mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                                mScript.lastArray.lstSingatures.Add(src.lstSingatures.ElementAt(pos).Key, lstDest.Count - 1)
                            End If
                        End If
                    End If

                    'If blnFound Then
                    '    'найдено исключение
                    '    If blnReverse Then
                    '        lstDest.Add(curValue)
                    '        'добавляем сигнатуру
                    '        If IsNothing(src.lstSingatures) Then
                    '            Dim pos As Integer = src.lstSingatures.IndexOfValue(i)
                    '            If pos > -1 Then
                    '                If IsNothing(mScript.lastArray.lstSingatures) Then mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                    '                mScript.lastArray.lstSingatures.Add(src.lstSingatures.ElementAt(pos).Key, lstDest.Count - 1)
                    '            End If
                    '        End If
                    '    Else
                    '        'nothing
                    '    End If
                    'Else
                    '    'исключение не найдено
                    '    If blnReverse Then
                    '        'nothing
                    '    Else
                    '        lstDest.Add(curValue)
                    '        'добавляем сигнатуру
                    '        If IsNothing(src.lstSingatures) Then
                    '            Dim pos As Integer = src.lstSingatures.IndexOfValue(i)
                    '            If pos > -1 Then
                    '                If IsNothing(mScript.lastArray.lstSingatures) Then mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                    '                mScript.lastArray.lstSingatures.Add(src.lstSingatures.ElementAt(pos).Key, lstDest.Count - 1)
                    '            End If
                    '        End If
                    '    End If
                    'End If
                Next i

                mScript.lastArray.arrValues = lstDest.ToArray
                If mScript.lastArray.arrValues.Count = 0 Then
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = ""
                End If

                lstDest.Clear()
                Return "#ARRAY"
            Case "GetValue"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                If var.arrValues.Count = 0 Then Return "-1"

                '1 element
                Dim id As Integer = cVariable.GetElementIndex(var, funcParams(1))
                If id < 0 Then Return funcParams(2)
                Return var.arrValues(id)
            Case "FindFirst", "FindLast"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                If var.arrValues.Count = 0 Then Return "-1"

                '1 delegate
                Dim deleg As String = UnWrapString(funcParams(1))
                Dim f As MatewScript.FunctionInfoType = Nothing
                If mScript.functionsHash.TryGetValue(deleg, f) = False Then Return _ERROR("Функция " & deleg & " не найдена.", functionName)

                '3,4 получение start, final
                Dim start As Integer, final As Integer, fStep As Integer = 1
                If functionName = "FindFirst" Then
                    start = CInt(GetParam(funcParams, 2, 0))
                    final = CInt(GetParam(funcParams, 3, -1))
                    If final = -1 Then final = var.arrValues.Count - 1
                    If start >= final Then Return _ERROR("Начальная позиция поиска больше конечной.", functionName)
                Else
                    start = CInt(GetParam(funcParams, 2, -1))
                    final = CInt(GetParam(funcParams, 3, 0))
                    If start = -1 Then start = var.arrValues.Count - 1
                    If final >= start Then Return _ERROR("Начальная позиция поиска меньше конечной.", functionName)
                    fStep = -1
                End If

                'поиск по результату делегата
                mScript.PrepareForExternalFunction("Функция " & functionName & ", делегат " & deleg)
                For i As Integer = start To final Step fStep
                    Dim strSign As String = ""
                    If IsNothing(var.lstSingatures) = False Then
                        Dim pos As Integer = var.lstSingatures.IndexOfValue(i)
                        If pos > -1 Then strSign = WrapString(var.lstSingatures.ElementAt(pos).Key)
                    End If
                    Dim res As String = mScript.ExecuteCode(f.ValueExecuteDt, {i.ToString, var.arrValues(i), strSign}, True)
                    If res = "#Error" Then
                        mScript.RestoreAfterExternalFunction(True)
                        Return res
                    ElseIf res = "True" Then
                        mScript.RestoreAfterExternalFunction(False)
                        Return res
                    End If
                Next
                mScript.RestoreAfterExternalFunction(False)

                'ничего не найдено
                Return "-1"
            Case "ForEach"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                If var.arrValues.Count = 0 Then Return "-1"

                '1 delegate
                Dim deleg As String = UnWrapString(funcParams(1))
                Dim f As MatewScript.FunctionInfoType = Nothing
                If mScript.functionsHash.TryGetValue(deleg, f) = False Then Return _ERROR("Функция " & deleg & " не найдена.", functionName)

                '2,3 получение start, final
                Dim start As Integer, final As Integer
                start = CInt(GetParam(funcParams, 2, 0))
                final = CInt(GetParam(funcParams, 3, -1))
                If final = -1 Then final = var.arrValues.Count - 1
                If start >= final Then Return _ERROR("Начальная позиция поиска больше конечной.", functionName)

                'применение делегата ко всем элементам массива
                mScript.lastArray = New cVariable.variableEditorInfoType
                mScript.PrepareForExternalFunction("Функция " & functionName & ", делегат " & deleg)
                Dim lstDest As New List(Of String)
                For i As Integer = start To final
                    Dim strSign As String = ""
                    If IsNothing(var.lstSingatures) = False Then
                        Dim pos As Integer = var.lstSingatures.IndexOfValue(i)
                        If pos > -1 Then strSign = WrapString(var.lstSingatures.ElementAt(pos).Key)
                    End If
                    Dim res As String = mScript.ExecuteCode(f.ValueExecuteDt, {i.ToString, var.arrValues(i), strSign}, True)
                    If res = "#Error" Then mScript.RestoreAfterExternalFunction(True) : Return res
                    mScript.EXIT_CODE = False
                    If String.IsNullOrEmpty(res) = False Then
                        lstDest.Add(res)
                        If String.IsNullOrEmpty(strSign) = False Then 'добавляем сигнатуры
                            If IsNothing(mScript.lastArray.lstSingatures) Then mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                            mScript.lastArray.lstSingatures.Add(UnWrapString(strSign), lstDest.Count - 1)
                        End If
                    End If
                Next
                mScript.RestoreAfterExternalFunction(False)

                If lstDest.Count = 0 Then
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = ""
                Else
                    mScript.lastArray.arrValues = lstDest.ToArray
                End If
                lstDest.Clear()

                Return "#ARRAY"
            Case "TrueForAll"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                If var.arrValues.Count = 0 Then Return "-1"

                '1 delegate
                Dim deleg As String = UnWrapString(funcParams(1))
                Dim f As MatewScript.FunctionInfoType = Nothing
                If mScript.functionsHash.TryGetValue(deleg, f) = False Then Return _ERROR("Функция " & deleg & " не найдена.", functionName)

                '2,3 получение start, final
                Dim start As Integer, final As Integer
                start = CInt(GetParam(funcParams, 2, 0))
                final = CInt(GetParam(funcParams, 3, -1))
                If final = -1 Then final = var.arrValues.Count - 1
                If start > final Then Return _ERROR("Начальная позиция поиска больше конечной.", functionName)

                'применение делегата ко всем элементам массива
                mScript.PrepareForExternalFunction("Функция " & functionName & ", делегат " & deleg)
                Dim lstDest As New List(Of String)
                For i As Integer = start To final
                    Dim strSign As String = ""
                    If IsNothing(var.lstSingatures) = False Then
                        Dim pos As Integer = var.lstSingatures.IndexOfValue(i)
                        If pos > -1 Then strSign = WrapString(var.lstSingatures.ElementAt(pos).Key)
                    End If
                    Dim res As String = mScript.ExecuteCode(f.ValueExecuteDt, {i.ToString, var.arrValues(i), strSign}, True)
                    If res = "#Error" Then
                        mScript.RestoreAfterExternalFunction(True)
                        Return res
                    ElseIf res = "False" Then
                        mScript.RestoreAfterExternalFunction(False)
                        Return res
                    End If
                Next
                mScript.RestoreAfterExternalFunction(False)

                Return "True"
            Case "IndexBySignature"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                If var.arrValues.Count = 0 Then Return "-1"
                '1 signature
                Dim strSign As String = UnWrapString(funcParams(1))

                Dim id As Integer = -1
                If IsNothing(var.lstSingatures) OrElse var.lstSingatures.TryGetValue(strSign, id) = False Then Return "-1"
                Return id.ToString
            Case "SignatureByIndex"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '1 index
                Dim id As Integer = CInt(funcParams(1))

                If IsNothing(var.lstSingatures) OrElse var.lstSingatures.Count = 0 Then Return ""
                Dim pos As Integer = var.lstSingatures.IndexOfValue(id)
                If pos = -1 Then Return ""
                Return WrapString(var.lstSingatures.ElementAt(pos).Key)
            Case "IndexesOfAllSimilars"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '1 strSeek
                Dim strSeek As String = funcParams(1)

                '2,3,4 получение start, final, errors
                Dim start As Integer, final As Integer
                start = CInt(GetParam(funcParams, 2, 0))
                final = CInt(GetParam(funcParams, 3, -1))
                If final = -1 Then final = var.arrValues.Count - 1
                If start >= final Then Return _ERROR("Начальная позиция поиска больше конечной.", functionName)
                Dim errCount As Integer = CInt(GetParam(funcParams, 4, "-1"))

                'создаем массив со всеми вхождениями
                Dim lstDest As New List(Of String)
                Do
                    Dim id As Integer = IndexOfAnyWithErrors(strSeek, start, final, var.arrValues, arrParams, errCount)
                    If id < 0 Then Exit Do
                    lstDest.Add(var.arrValues(id))
                    start = id + 1
                    If start > final Then Exit Do
                Loop

                mScript.lastArray = New cVariable.variableEditorInfoType
                If lstDest.Count = 0 Then
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = ""
                Else
                    mScript.lastArray.arrValues = lstDest.ToArray
                    lstDest.Clear()
                End If

                Return "#ARRAY"
            Case "IndexOfSimilar", "LastIndexOfSimilar"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '1 strSeek
                Dim strSeek As String = funcParams(1)

                '2,3,4 получение start, final, errors
                Dim start As Integer, final As Integer
                If functionName = "IndexOfSimilar" Then
                    start = CInt(GetParam(funcParams, 2, 0))
                    final = CInt(GetParam(funcParams, 3, -1))
                    If final = -1 Then final = var.arrValues.Count - 1
                    If start > final Then Return _ERROR("Начальная позиция поиска больше конечной.", functionName)
                    'поиск точного совпадения
                    Dim pos As Integer = var.arrValues.ToList.IndexOf(strSeek, start, final - start + 1)
                    If pos > -1 Then Return pos.ToString
                Else
                    start = CInt(GetParam(funcParams, 2, -1))
                    final = CInt(GetParam(funcParams, 3, 0))
                    If start = -1 Then start = var.arrValues.Count - 1
                    If final > start Then Return _ERROR("Начальная позиция поиска меньше конечной.", functionName)
                    'поиск точного совпадения
                    Dim pos As Integer = var.arrValues.ToList.LastIndexOf(strSeek, start, start - final + 1)
                    If pos > -1 Then Return pos.ToString
                End If
                Dim errCount As Integer = CInt(GetParam(funcParams, 4, "-1"))

                'ищем первое совпадение
                Dim id As Integer
                id = IndexOfAnyWithErrors(strSeek, start, final, var.arrValues, arrParams, errCount)
                Return id.ToString
            Case "IndexesOfAllValues"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '1 strSeek
                Dim strSeek As String = funcParams(1)

                '2,3,4 получение start, final, iCase
                Dim start As Integer, final As Integer
                start = CInt(GetParam(funcParams, 2, 0))
                final = CInt(GetParam(funcParams, 3, -1))
                If final = -1 Then final = var.arrValues.Count - 1
                If start > final Then Return _ERROR("Начальная позиция поиска больше конечной.", functionName)
                Dim iCase As Boolean = CBool(GetParam(funcParams, 4, "True"))

                Dim lstDest As New List(Of String)
                For i As Integer = start To final
                    If String.Compare(strSeek, var.arrValues(i), iCase) = 0 Then lstDest.Add(i.ToString)
                Next

                mScript.lastArray = New cVariable.variableEditorInfoType
                If lstDest.Count = 0 Then
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = ""
                Else
                    mScript.lastArray.arrValues = lstDest.ToArray
                    lstDest.Clear()
                End If

                Return "#ARRAY"
            Case "FilterValues"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '1 strSeek
                Dim strSeek As String = funcParams(1)

                '2,3,4 получение start, final, iCase
                Dim start As Integer, final As Integer
                start = CInt(GetParam(funcParams, 2, 0))
                final = CInt(GetParam(funcParams, 3, -1))
                If final = -1 Then final = var.arrValues.Count - 1
                If start > final Then Return _ERROR("Начальная позиция поиска больше конечной.", functionName)
                Dim iCase As Boolean = CBool(GetParam(funcParams, 4, "True"))
                Dim invertRes As Boolean = CBool(GetParam(funcParams, 5, "False"))

                Dim arrSrc() As String = Nothing
                If start = 0 AndAlso final = var.arrValues.Count - 1 Then
                    arrSrc = var.arrValues
                Else
                    ReDim arrSrc(final - start)
                    Array.ConstrainedCopy(var.arrValues, start, arrSrc, 0, final - start + 1)
                End If

                Dim arrDest() As String = Filter(arrSrc, UnWrapString(strSeek), Not invertRes, IIf(iCase, CompareMethod.Text, CompareMethod.Binary))

                mScript.lastArray = New cVariable.variableEditorInfoType
                If arrDest.Count = 0 Then
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = ""
                Else
                    mScript.lastArray.arrValues = arrDest
                End If
                Erase arrSrc

                Return "#ARRAY"
            Case "FilterSignatures"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '1 strSeek
                Dim strSeek As String = funcParams(1)

                '2,3,4 получение start, final, iCase
                Dim start As Integer, final As Integer
                start = CInt(GetParam(funcParams, 2, 0))
                final = CInt(GetParam(funcParams, 3, -1))
                If final = -1 Then final = var.arrValues.Count - 1
                If start > final Then Return _ERROR("Начальная позиция поиска больше конечной.", functionName)
                Dim iCase As Boolean = CBool(GetParam(funcParams, 4, "True"))
                Dim invertRes As Boolean = CBool(GetParam(funcParams, 5, "False"))
                If IsNothing(var.lstSingatures) OrElse var.lstSingatures.Count = 0 Then Return _ERROR("Массив не содержит сигнатур.", functionName)

                Dim arrSrc() As String = Nothing
                If start = 0 AndAlso final = var.arrValues.Count - 1 Then
                    arrSrc = var.lstSingatures.Keys.ToArray
                Else
                    Dim srcUB As Integer = -1
                    For i As Integer = 0 To var.lstSingatures.Count - 1
                        Dim id As Integer = var.lstSingatures.ElementAt(i).Value
                        If id < start OrElse id > final Then Continue For
                        srcUB += 1
                        ReDim Preserve arrSrc(srcUB)
                        arrSrc(srcUB) = var.lstSingatures.ElementAt(i).Key
                    Next i
                End If

                Dim arrDest() As String = Filter(arrSrc, UnWrapString(strSeek), Not invertRes, IIf(iCase, CompareMethod.Text, CompareMethod.Binary))

                mScript.lastArray = New cVariable.variableEditorInfoType
                If arrDest.Count = 0 Then
                    ReDim mScript.lastArray.arrValues(0)
                    mScript.lastArray.arrValues(0) = ""
                Else
                    mScript.lastArray.arrValues = arrDest
                End If
                Erase arrSrc

                Return "#ARRAY"
            Case "IndexOfValue", "LastIndexOfValue"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                '1 strSeek
                Dim strSeek As String = funcParams(1)

                '2,3,4 получение start, final,iCase
                Dim iCase As Boolean = CBool(GetParam(funcParams, 4, True))
                Dim start As Integer, final As Integer, pos As Integer = -1
                If functionName = "IndexOfValue" Then
                    start = CInt(GetParam(funcParams, 2, 0))
                    final = CInt(GetParam(funcParams, 3, -1))
                    If final = -1 Then final = var.arrValues.Count - 1
                    If start > final Then Return _ERROR("Начальная позиция поиска больше конечной.", functionName)
                    'поиск точного совпадения
                    If iCase Then
                        For i As Integer = start To final
                            If String.Compare(var.arrValues(i), strSeek, True) = 0 Then Return i.ToString
                        Next
                    Else
                        pos = var.arrValues.ToList.IndexOf(strSeek, start, final - start + 1)
                    End If
                Else
                    start = CInt(GetParam(funcParams, 2, -1))
                    final = CInt(GetParam(funcParams, 3, 0))
                    If start = -1 Then start = var.arrValues.Count - 1
                    If final > start Then Return _ERROR("Начальная позиция поиска меньше конечной.", functionName)
                    'поиск точного совпадения
                    If iCase Then
                        For i As Integer = start To final Step -1
                            If String.Compare(var.arrValues(i), strSeek, True) = 0 Then Return i.ToString
                        Next
                    Else
                        pos = var.arrValues.ToList.LastIndexOf(strSeek, start, start - final + 1)
                    End If
                End If
                Return pos.ToString

            Case "Insert"
                Dim varName As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                Dim varClass As cVariable = GetVariableClass(varName)
                If IsNothing(varClass) Then Return _ERROR("Массив не найден.", functionName)
                Dim elementIndex As Integer = Val(funcParams(2))
                Dim varSign As String = UnWrapString(GetParam(funcParams, 3, ""))
                If varClass.Insert(varName, elementIndex, funcParams(1), varSign) = "#Error" Then Return _ERROR("Массив не найден.", functionName)
                Return ""
            Case "IsClear"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                If var.arrValues.Count = 1 AndAlso var.arrValues(0) = "" Then Return "True"
                Return "False"
            Case "Remove"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                Dim remId As Integer = cVariable.GetElementIndex(var, funcParams(1))
                If remId < 0 Then Return _ERROR("Массив " & varName & " не содержит элемент " & funcParams(1) & ".", functionName)

                Dim lstValues As List(Of String) = var.arrValues.ToList
                lstValues.RemoveAt(remId)
                var.arrValues = lstValues.ToArray
                lstValues.Clear()

                If IsNothing(var.lstSingatures) = False AndAlso var.lstSingatures.Count > 0 Then
                    'смещаем сигнатуры
                    For i As Integer = var.lstSingatures.Count - 1 To 0 Step -1
                        Dim valId As Integer = var.lstSingatures.ElementAt(i).Value
                        If valId = remId Then
                            var.lstSingatures.RemoveAt(i)
                        ElseIf valId > remId Then
                            Dim strSign As String = var.lstSingatures.ElementAt(i).Key
                            var.lstSingatures(strSign) -= 1
                        End If
                    Next
                End If

                If var.arrValues.Count = 0 Then
                    ReDim var.arrValues(0)
                    var.arrValues(0) = ""
                    Return 0
                End If

                Return var.arrValues.Count.ToString
            Case "RemoveRange"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                '1 start
                Dim start As Integer = cVariable.GetElementIndex(var, funcParams(1))
                If start = -1 Then Return _ERROR("Элемент " & funcParams(1) & "в массиве " & varName & " не найден.", functionName)

                '2 final
                Dim final As Integer, El As String = GetParam(funcParams, 2, "-1")
                If El = "-1" Then
                    final = var.arrValues.Count - 1
                Else
                    final = cVariable.GetElementIndex(var, El)
                End If

                If cVariable.RemoveRange(var, start, final) = "#Error" Then Return _ERROR("Массив не найден.", functionName)
                Return ""
            Case "RemoveSignature"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                Dim id As Integer = cVariable.GetElementIndex(var, funcParams(1))
                If IsNothing(var.lstSingatures) Then Return "False"
                Dim pos As Integer = var.lstSingatures.IndexOfValue(id)
                If pos = -1 Then Return "False"
                var.lstSingatures.RemoveAt(pos)
                Return "True"
            Case "Replace", "ReplaceRev"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                Dim strOld As String = UnWrapString(funcParams(1))
                Dim strNew As String = funcParams(2) 'НЕ обрабатываем функцией PrepareStringToPrint

                Dim start As Integer = CInt(GetParam(funcParams, 3, 0))
                Dim finish As Integer = CInt(GetParam(funcParams, 4, -1))
                If finish > 0 AndAlso finish < start Then Return _ERROR("Неверные входные параметры функции. Конец поиска раньше, чем начало", functionName)
                Dim beginFromEnd As Boolean = False
                If functionName = "InArrayRev" Then beginFromEnd = True

                mScript.lastArray = New cVariable.variableEditorInfoType
                Dim replaceCount As Integer = 0
                mScript.lastArray.arrValues = cVariable.Replace(var, strOld, strNew, replaceCount, arrParams, start, finish, beginFromEnd)
                If IsNothing(mScript.lastArray.arrValues) Then Return _ERROR("Ошибка при выполнении поиска и замены.", functionName)

                Return "#ARRAY"
            Case "Resize"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                Dim newSize As Integer = 0
                Integer.TryParse(funcParams(1), newSize)
                If newSize <= 0 Then
                    ReDim var.arrValues(0)
                    var.arrValues(0) = ""
                Else
                    ReDim Preserve var.arrValues(newSize - 1)
                End If
                Return ""
            Case "Reverse"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                Dim cnt As Integer = var.arrValues.Length
                mScript.lastArray = New cVariable.variableEditorInfoType
                ReDim mScript.lastArray.arrValues(cnt - 1)
                Array.ConstrainedCopy(var.arrValues, 0, mScript.lastArray.arrValues, 0, cnt)
                Array.Reverse(mScript.lastArray.arrValues)

                'меняем индексы в сигнатурах
                If IsNothing(var.lstSingatures) = False AndAlso var.lstSingatures.Count > 0 Then
                    mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                    For i As Integer = 0 To var.lstSingatures.Count - 1
                        mScript.lastArray.lstSingatures.Add(var.lstSingatures.ElementAt(i).Key, cnt - var.lstSingatures.ElementAt(i).Value - 1)
                    Next
                End If

                Return "#ARRAY"
            Case "Sort"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                Dim desc As Boolean = CBool(GetParam(funcParams, 1, "False"))

                Dim cnt As Integer = var.arrValues.Length
                mScript.lastArray = New cVariable.variableEditorInfoType
                ReDim mScript.lastArray.arrValues(cnt - 1)
                Array.ConstrainedCopy(var.arrValues, 0, mScript.lastArray.arrValues, 0, cnt)
                Array.Sort(mScript.lastArray.arrValues)
                If desc Then Array.Reverse(mScript.lastArray.arrValues)

                'меняем индексы в сигнатурах
                If IsNothing(var.lstSingatures) = False AndAlso var.lstSingatures.Count > 0 Then
                    mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                    Dim lstSrc As List(Of String) = var.arrValues.ToList 'несортированный исходный массив значений
                    Dim lstDest As List(Of String) = mScript.lastArray.arrValues.ToList 'отсортированный итоговый массив значений
                    For i As Integer = 0 To var.lstSingatures.Count - 1
                        'перебор всех сигнатур
                        Dim seekValue As String = lstSrc(var.lstSingatures.ElementAt(i).Value) 'значение, соответствующее текущей сигнатуре
                        Dim newId As Integer = lstDest.IndexOf(seekValue) 'индекс элемента в отсортированном списке, значение которого соответствует сигнатуре
                        If newId = -1 Then Return _ERROR("Ошибка при сортировке сигнатур.", functionName)

                        If IsNothing(mScript.lastArray.lstSingatures) = False Then
                            'если новому индексу уже соответствует сигнатура (есть одинаковые значения в массиве), то берем следующий элемент (список отсортирован)
                            Do
                                If mScript.lastArray.lstSingatures.ContainsValue(newId) Then
                                    newId += 1
                                    If newId > lstDest.Count - 1 Then Return _ERROR("Ошибка при сортировке сигнатур.", functionName)
                                Else
                                    Exit Do
                                End If
                            Loop
                        End If

                        mScript.lastArray.lstSingatures.Add(var.lstSingatures.ElementAt(i).Key, newId)
                    Next
                    lstDest.Clear()
                    lstSrc.Clear()
                End If

                Return "#ARRAY"
            Case "Split"
                Dim strSrc As String = UnWrapString(funcParams(0))
                Dim delimer As String = UnWrapString(GetParam(funcParams, 1, "' '"))
                Dim limit As Integer = CInt(GetParam(funcParams, 2, "-1"))

                mScript.lastArray = New cVariable.variableEditorInfoType
                mScript.lastArray.arrValues = Split(strSrc, delimer, limit)

                For i As Integer = 0 To mScript.lastArray.arrValues.Count - 1
                    mScript.lastArray.arrValues(i) = WrapString(mScript.lastArray.arrValues(i))
                Next i
                Return "#ARRAY"
            Case "SwapValuesToSignatures"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                If IsNothing(var.lstSingatures) OrElse var.lstSingatures.Count <> var.arrValues.Count Then Return _ERROR("Число сигнатур не совпадает с количеством значений.", functionName)

                Dim lstValues As New List(Of String)
                For i As Integer = 0 To var.arrValues.Count - 1
                    Dim val As String = var.arrValues(i)
                    If lstValues.Contains(val) Then Return _ERROR("Значения в массиве не должны повторяться.", functionName)
                    lstValues.Add(val)
                Next
                lstValues.Clear()

                'мы уже убедились, что все значения в массиве уникальны, количество сигнатур = количеству значений
                mScript.lastArray = New cVariable.variableEditorInfoType
                ReDim mScript.lastArray.arrValues(var.arrValues.Count - 1)
                mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)

                For i As Integer = 0 To var.arrValues.Count - 1
                    Dim id As Integer = var.lstSingatures.ElementAt(i).Value
                    Dim sign As String = var.lstSingatures.ElementAt(i).Key
                    Dim val As String = var.arrValues(id)

                    mScript.lastArray.arrValues(id) = sign
                    mScript.lastArray.lstSingatures.Add(val, id)
                Next

                Return "#ARRAY"
            Case "IsValuesUnique"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)

                Dim lstValues As New List(Of String)
                For i As Integer = 0 To var.arrValues.Count - 1
                    Dim val As String = var.arrValues(i)
                    If lstValues.Contains(val) Then Return "False"
                    lstValues.Add(val)
                Next

                Return "True"
            Case "Union", "UnionUnique"
                Dim unSign As Boolean = False
                Boolean.TryParse(funcParams(0), unSign)

                mScript.lastArray = New cVariable.variableEditorInfoType
                Dim unUB As Integer = -1
                For i As Integer = 1 To funcParams.Count - 1
                    Dim varName As String = UnWrapString(funcParams(i))
                    Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                    If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                    If var.arrValues.Count = 1 AndAlso var.arrValues(0) = "" Then Continue For

                    If IsNothing(mScript.lastArray.arrValues) Then
                        ReDim mScript.lastArray.arrValues(var.arrValues.Count - 1)
                    Else
                        ReDim Preserve mScript.lastArray.arrValues(mScript.lastArray.arrValues.Count + var.arrValues.Count - 1)
                    End If
                    'mScript.lastArray.arrValues.Union(var.arrValues)
                    Array.ConstrainedCopy(var.arrValues, 0, mScript.lastArray.arrValues, mScript.lastArray.arrValues.Count - var.arrValues.Count, var.arrValues.Count)

                    If unSign AndAlso IsNothing(var.lstSingatures) = False AndAlso var.lstSingatures.Count > 0 Then
                        'объединение сигнатур
                        If IsNothing(mScript.lastArray.lstSingatures) Then mScript.lastArray.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)

                        For j As Integer = 0 To var.lstSingatures.Count - 1
                            Dim sign As String = var.lstSingatures.ElementAt(j).Key
                            If mScript.lastArray.lstSingatures.ContainsKey(sign) Then Return _ERROR("Массивы не могут быть объеденены. Одинаковые сигнатуры недопустимы.", functionName)
                            Dim oldId As Integer = var.lstSingatures.ElementAt(j).Value
                            Dim newId As Integer = oldId + unUB + 1
                            mScript.lastArray.lstSingatures.Add(sign, newId)
                        Next j
                        unUB = mScript.lastArray.arrValues.Count - 1
                    End If
                Next i

                If functionName = "UnionUnique" Then
                    Dim curId As Integer = 0
                    Dim arrValUB As Integer = mScript.lastArray.arrValues.Count - 1
                    Dim lstSeek As List(Of String) = mScript.lastArray.arrValues.ToList

                    If unSign Then
                        If IsNothing(mScript.lastArray.lstSingatures) OrElse mScript.lastArray.lstSingatures.Count = 0 Then unSign = False
                    End If

                    Do While curId < arrValUB
                        Dim pos As Integer = lstSeek.IndexOf(lstSeek(curId), curId + 1)
                        If pos = -1 Then
                            curId += 1
                            If curId >= lstSeek.Count - 2 Then Exit Do
                            Continue Do
                        End If

                        'значение неуникально, найден повтор
                        lstSeek.RemoveAt(pos)
                        If unSign Then
                            'удаляем сигнатуру повтора
                            Dim sId As Integer = mScript.lastArray.lstSingatures.IndexOfValue(pos)
                            If sId > -1 Then mScript.lastArray.lstSingatures.RemoveAt(sId)

                            'уменьшаем индекс в сигнатурах, если он был больше pos
                            For i As Integer = 0 To mScript.lastArray.lstSingatures.Count - 1
                                If mScript.lastArray.lstSingatures.ElementAt(i).Value > pos Then
                                    Dim sKey As String = mScript.lastArray.lstSingatures.ElementAt(i).Key
                                    mScript.lastArray.lstSingatures(sKey) -= 1
                                End If
                            Next i
                        End If
                    Loop
                    mScript.lastArray.arrValues = lstSeek.ToArray
                End If

                Return "#ARRAY"
            Case "GetStructure"
                '0 var varName
                Dim varName As String = UnWrapString(funcParams(0))
                Dim var As cVariable.variableEditorInfoType = GetVariable(varName)
                If IsNothing(var) Then Return _ERROR("Массив " & varName & " не найден.", functionName)
                Return cVariable.GetStructure(var)
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function


    ''' <summary>
    ''' Возвращает класс переменной по ее имени (локальный csLocalVariables или csPublicVariables
    ''' </summary>
    ''' <param name="varName">имя переменной</param>
    ''' <returns>класс переменной или Nothing, если не найден</returns>
    Private Function GetVariableClass(ByVal varName As String) As cVariable
        If varName.StartsWith("V.", StringComparison.CurrentCultureIgnoreCase) Then varName = varName.Substring(2)
        If mScript.csLocalVariables.GetVariable(varName, 0) <> "#Error" Then Return mScript.csLocalVariables
        If mScript.csPublicVariables.GetVariable(varName, 0) <> "#Error" Then Return mScript.csPublicVariables
        Return Nothing
    End Function

    ''' <summary>
    ''' Получаем ссылку на переменную по ее имена
    ''' </summary>
    ''' <param name="varName">имя переменной без кавычек</param>
    ''' <param name="makeNewIfAbsend">создавать ли новую если не найдена</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetVariable(ByVal varName As String, Optional ByVal makeNewIfAbsend As Boolean = False) As cVariable.variableEditorInfoType
        If varName.StartsWith("V.", StringComparison.CurrentCultureIgnoreCase) Then varName = varName.Substring(2)
        Dim var As cVariable.variableEditorInfoType = Nothing
        If mScript.csLocalVariables.lstVariables.TryGetValue(varName, var) Then Return var
        If mScript.csPublicVariables.lstVariables.TryGetValue(varName, var) Then Return var
        If Not makeNewIfAbsend Then Return Nothing
        mScript.csLocalVariables.SetVariableInternal(varName, "", 0)
        Return mScript.csLocalVariables.lstVariables(varName)
    End Function

    Private Function classD_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        'не проверена
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Select Case functionName
            Case "TimeInGame", "TimeAfterStarting", "TimeAfterSaving", "TimeInBattle", "TimeInLocation"
                Dim ts As New TimeSpan(Now.Ticks)
                Dim initTicks As Long
                Select Case functionName
                    Case "TimeAfterStarting"
                        initTicks = GVARS.TIME_STARTING
                    Case "TimeAfterSaving"
                        initTicks = GVARS.TIME_SAVING
                    Case "TimeInBattle"
                        If GVARS.G_ISBATTLE = False Then Return _ERROR("Функция доступна только во время боя!", functionName)
                        initTicks = GVARS.TIME_IN_BATTLE
                    Case "TimeInLocation"
                        initTicks = GVARS.TIME_IN_THIS_LOCATION
                    Case Else
                        initTicks = GVARS.TIME_IN_GAME
                End Select

                Return Math.Round(ts.Subtract(New TimeSpan(initTicks)).TotalMilliseconds).ToString
            Case "Now"
                Dim dFormat As String = GetParam(funcParams, 0, "dd.MM.yyyy HH:mm:ss")
                Dim d As Date = Now
                Return "'" + Format(d, UnWrapString(dFormat)) + "'"
            Case "NowDate"
                Dim d As Date = Now
                Return "'" + Format(d, "dd.MM.yyyy") + "'"
            Case "NowTime"
                Dim d As Date = Now
                Return "'" + Format(d, "HH:mm:ss") + "'"
            Case "DateDiff"
                Dim dInterval As DateInterval = Val(mScript.PrepareStringToPrint(funcParams(0), arrParams))

                If dInterval < 0 OrElse dInterval > 9 Then
                    Return _ERROR("Неправильно выбран интервал. Введите одно из предлагаемых значений.", functionName)
                End If

                Dim d1 As Date = GetDateByString(mScript.PrepareStringToPrint(funcParams(1), arrParams))
                If mScript.LAST_ERROR.Length > 0 Then Return _ERROR("Неверный формат первой даты. Используйте формат dd.mm.yyyy или dd.mm.yyyy HH:mm:ss", functionName)
                Dim d2 As Date
                If funcParams.Count < 3 Then
                    d2 = Now
                Else
                    d2 = GetDateByString(mScript.PrepareStringToPrint(funcParams(2), arrParams))
                    If mScript.LAST_ERROR.Length > 0 Then Return _ERROR("Неверный формат второй даты. Используйте формат dd.mm.yyyy или dd.mm.yyyy HH:mm:ss", functionName)
                End If

                Return DateDiff(dInterval, d1, d2).ToString
            Case "DaysInMonth"
                Dim y As Integer = CInt(funcParams(0))
                Dim m As Integer = CInt(funcParams(1))
                If m < 1 OrElse m > 12 Then Return _ERROR("Номер месяца должен быть в диапазоне 1-12.", functionName)
                Return Date.DaysInMonth(y, m).ToString
            Case "IsLeapYear"
                Dim y As Integer = CInt(funcParams(0))
                Return Date.IsLeapYear(y).ToString
            Case "Ticks"
                Dim cl As New Microsoft.VisualBasic.Devices.Clock
                Return cl.TickCount
            Case "DateAdd"
                Dim dInterval As DateInterval = Val(mScript.PrepareStringToPrint(funcParams(0), arrParams))

                If dInterval < 0 OrElse dInterval > 9 Then
                    Return _ERROR("Неправильно выбран интервал. Введите одно из предлагаемых значений.", functionName)
                End If

                Dim num As Integer = CInt(funcParams(1))
                Dim strDate As String = UnWrapString(GetParam(funcParams, 2, Format(Date.Now, "dd.MM.yyyy HH:mm:ss")))
                Dim startDate As Date = Now
                If Date.TryParseExact(strDate, "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.AllowLeadingWhite, startDate) = False Then
                    Return _ERROR("Неправильно указан формат даты.", functionName)
                End If
                Dim nDate As Date = DateAdd(dInterval, num, startDate)
                strDate = Format(nDate, "dd.MM.yyyy HH:mm:ss")
                Return WrapString(strDate)
            Case "Seconds", "Minutes", "Hours", "Days", "Monthes", "Years", "Weekday"
                Dim strDate As String = UnWrapString(GetParam(funcParams, 0, Format(Date.Now, "dd.MM.yyyy HH:mm:ss")))
                Dim aDate As Date = Now
                If Date.TryParseExact(strDate, "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.AllowLeadingWhite, aDate) = False Then
                    Return _ERROR("Неправильно указан формат даты.", functionName)
                End If
                Select Case functionName
                    Case "Seconds"
                        Return aDate.Second.ToString
                    Case "Minutes"
                        Return aDate.Minute.ToString
                    Case "Hours"
                        Return aDate.Hour.ToString
                    Case "Days"
                        Return aDate.Day.ToString
                    Case "Monthes"
                        Return aDate.Month.ToString
                    Case "Years"
                        Return aDate.Year.ToString
                    Case "Weekday"
                        Return (aDate.DayOfWeek + 1).ToString
                End Select
            Case "PlayerDate_Add"
                Dim dInterval As DateInterval = Val(UnWrapString(funcParams(0)))
                Dim addCount As Integer = Val(funcParams(1))
                Dim res As String = PlayerDate_Add(dInterval, addCount, funcParams)
                If res = "#Error" Then Return res
                If Abilitities_RaiseAbilityOnTimeEllapsedEvents() = "#Error" Then Return "#Error"
                res = WrapString(res)
                Return res
            Case "PlayerDate_DateDiff"
                Dim dInterval As DateInterval = Val(UnWrapString(funcParams(0)))
                Dim aDate2 As String = GetParam(funcParams, 2, "")
                Return PlayerDate_DateDiff(dInterval, funcParams(1), aDate2, arrParams).ToString
            Case "PlayerDate_GetDay", "PlayerDate_GetHours", "PlayerDate_GetMinutes", "PlayerDate_GetSeconds", "PlayerDate_GetWeekday", "PlayerDate_GetYear", "PlayerDate_GetMonth"
                Dim aDate As String = GetParam(funcParams, 0, "")
                Dim dInterval As DateInterval = DateInterval.Year
                Select Case functionName
                    Case "PlayerDate_GetDay"
                        dInterval = DateInterval.Day
                    Case "PlayerDate_GetHours"
                        dInterval = DateInterval.Hour
                    Case "PlayerDate_GetMinutes"
                        dInterval = DateInterval.Minute
                    Case "PlayerDate_GetSeconds"
                        dInterval = DateInterval.Second
                    Case "PlayerDate_GetWeekday"
                        dInterval = DateInterval.Weekday
                    Case "PlayerDate_GetMonth"
                        dInterval = DateInterval.Month
                        'Case "PlayerDate_GetYear"
                End Select
                Return PlayerDate_GetPeriod(dInterval, funcParams, aDate).ToString
            Case "PlayerDate_Get"
                Dim dFormat As String = GetParam(funcParams, 0, "dd.MM.yyyy HH:mm:ss")
                Dim aDate As String = GetParam(funcParams, 1, "")
                Return WrapString(PlayerDate_Get(dFormat, funcParams))
            Case "PlayerDate_DateToArray"
                Dim aDate As String = GetParam(funcParams, 0, "")
                mScript.lastArray = PlayerDate_DateToArray(funcParams, aDate)

                Return "#ARRAY"
            Case "PlayerDate_IntervalToArray"
                Dim aDate As String = funcParams(0)
                Dim initMonth As Integer = Val(GetParam(funcParams, 1, "-1"))
                If initMonth <= 0 Then
                    initMonth = -1
                Else
                    initMonth += 1
                End If
                mScript.lastArray = PlayerDate_IntervalToArray(funcParams, aDate, initMonth)

                Return "#ARRAY"
            Case "PlayerDate_Set"
                Dim res As String = PlayerDate_SetDate(funcParams(0), funcParams)
                If Abilitities_RaiseAbilityOnTimeEllapsedEvents() = "#Error" Then Return "#Error"
                Return res
            Case "PlayerDate_SleepTill"
                Dim var As cVariable.variableEditorInfoType = PlayerDate_SleepTill(UnWrapString(funcParams(0)), funcParams)
                If String.IsNullOrEmpty(var.arrValues(0)) Then Return "#Error"
                If Abilitities_RaiseAbilityOnTimeEllapsedEvents() = "#Error" Then Return "#Error"
                mScript.lastArray = var
                Return "#ARRAY"
            Case "PlayerDate_ToSeconds"
                Dim strDate As String = GetParam(funcParams, 0, "")
                If String.IsNullOrEmpty(strDate) Then Return GVARS.PLAYER_TIME.ToString
                Return PlayerDate_DateToSeconds(strDate, arrParams).ToString
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    ''' <summary>
    ''' Возвращает дату из ее строкового представления
    ''' </summary>
    ''' <param name="strDate">дата в формате dd.mm.yyyy hh:mm:ss (или без времени)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetDateByString(ByVal strDate As String) As Date
        If strDate Like "##.##.####*" = False Then
            mScript.LAST_ERROR = "date error"
            Return Nothing
        End If
        Dim y, m, d As Integer
        d = CInt(Left(strDate, 2))
        m = CInt(Mid(strDate, 4, 2))
        y = CInt(Mid(strDate, 7, 4))
        If strDate Like "##.##.#### ##:##:##" Then
            Dim h, mn, s As Integer
            h = CInt(Mid(strDate, 12, 2))
            mn = CInt(Mid(strDate, 15, 2))
            s = CInt(Mid(strDate, 18, 2))
            Try
                Return New Date(y, m, d, h, mn, s)
            Catch ex As Exception
                mScript.LAST_ERROR = "date error"
                Return Nothing
            End Try
        Else
            Try
                Return New Date(y, m, d)
            Catch ex As Exception
                mScript.LAST_ERROR = "date eror"
                Return Nothing
            End Try
        End If
    End Function

    Private Function classMath_functions(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String, ByRef arrParams() As String) As String
        If CheckFunctionParams(classId, functionName, funcParams) = False Then Return "#Error"
        Dim i As Integer
        Select Case functionName
            Case "Abs"
                Return Math.Abs(Double.Parse(funcParams(0), provider_points)).ToString(provider_points)
            Case "ArcCos"
                Return Math.Acos(Double.Parse(funcParams(0), provider_points)).ToString(provider_points)
            Case "ArcSin"
                Return Math.Asin(Double.Parse(funcParams(0), provider_points)).ToString(provider_points)
            Case "ArcTan"
                Return Math.Atan(Double.Parse(funcParams(0), provider_points)).ToString(provider_points)
            Case "Ceiling"
                Return Math.Ceiling(Double.Parse(funcParams(0), provider_points)).ToString(provider_points)
            Case "Floor"
                Return Math.Floor(Double.Parse(funcParams(0), provider_points)).ToString(provider_points)
            Case "Sqr"
                Dim sqr As Single = 2
                If funcParams.Count > 1 Then sqr = Val(funcParams(1))
                Return (Val(funcParams(0)) ^ (1 / sqr)).ToString.Replace(",", ".")
            Case "DegreesToRadians"
                Dim deg As Double = Double.Parse(funcParams(0), provider_points)
                Dim rad As Double = deg * Math.PI / 180
                Return rad.ToString(provider_points)
            Case "RadiansToDegrees"
                Dim rad As Double = Double.Parse(funcParams(0), provider_points)
                Dim deg As Double = rad * 180 / Math.PI
                Return deg.ToString(provider_points)
            Case "IsNum"
                Dim strVal As String = funcParams(0)
                Dim valType As MatewScript.ReturnFormatEnum = mScript.Param_GetType(strVal)
                If valType = MatewScript.ReturnFormatEnum.TO_NUMBER Then Return "True"
                If valType = MatewScript.ReturnFormatEnum.TO_STRING Then

                    strVal = strVal.Substring(1, strVal.Length - 2).Replace("/'", "'").Replace(".", ",")
                    If IsNumeric(strVal) Then
                        Return "True"
                    Else
                        Return "False"
                    End If
                End If
                Return "False"
            Case "Distance"
                Dim pt1 As New PointF(Double.Parse(funcParams(0), provider_points), Double.Parse(funcParams(1), provider_points))
                Dim pt2 As New PointF(Double.Parse(funcParams(2), provider_points), Double.Parse(funcParams(3), provider_points))
                Dim dist As Double = Math.Sqrt((pt2.X - pt1.X) * (pt2.X - pt1.X) + (pt2.Y - pt1.Y) * (pt2.Y - pt1.Y))
                Return dist.ToString(provider_points)
            Case "PI"
                Return Math.PI.ToString(provider_points)
            Case "E"
                Return Math.E.ToString(provider_points)
            Case "IsInteger"
                Dim d As Double = Double.Parse(funcParams(0), provider_points)
                Return Integer.TryParse(d, Nothing).ToString
            Case "IsPositive"
                Dim d As Double = Double.Parse(funcParams(0), provider_points)
                Return (d >= 0).ToString
            Case "Log"
                Dim bas As Double
                If funcParams.Count > 1 Then
                    bas = Double.Parse(funcParams(1), provider_points)
                Else
                    bas = Math.E
                End If
                Return Math.Log(Double.Parse(funcParams(0), provider_points), bas).ToString(provider_points)
            Case "Random"
                Dim lower As Single = 0
                If funcParams.Count > 0 Then lower = Val(funcParams(0))
                Dim upper As Single = 1
                If funcParams.Count > 1 Then upper = Val(funcParams(1))
                Dim digits As Integer = 0
                If funcParams.Count > 2 Then digits = Val(funcParams(2))
                If upper < lower Then
                    mScript.LAST_ERROR = "Максимальное значение меньше или равно минимальному."
                    Return "#Error"
                ElseIf upper = lower Then
                    Return upper.ToString.Replace(",", ".")
                End If
                Return Math.Round((upper - lower) * VBMath.Rnd + lower, digits).ToString.Replace(",", ".")
            Case "Round"
                Dim digits As Integer = 0
                If funcParams.Count > 1 Then digits = Val(funcParams(1))
                Return Math.Round(Val(funcParams(0)), digits).ToString.Replace(",", ".")
            Case "ToString"
                Return mScript.ConvertElement(funcParams(0), MatewScript.ReturnFormatEnum.TO_STRING)
            Case "ToBool"
                Return mScript.ConvertElement(funcParams(0), MatewScript.ReturnFormatEnum.TO_BOOL)
            Case "Sin", "Cos", "Tan"
                Dim rad As Double

                If mScript.Param_GetType(funcParams(0)) = MatewScript.ReturnFormatEnum.TO_NUMBER Then
                    'угол в радианах
                    rad = Val(funcParams(0))
                ElseIf mScript.Param_GetType(funcParams(0)) = MatewScript.ReturnFormatEnum.TO_STRING Then
                    'угол в градусах
                    rad = Val(mScript.PrepareStringToPrint(funcParams(0), arrParams)) * Math.PI / 180 'переводим в радианы
                Else
                    mScript.LAST_ERROR = "Неверный формат данных."
                    Return "#Error"
                End If
                Select Case functionName
                    Case "Sin"
                        Return Math.Sin(rad).ToString.Replace(",", ".")
                    Case "Cos"
                        Return Math.Cos(rad).ToString.Replace(",", ".")
                    Case "Tan"
                        Return Math.Tan(rad).ToString.Replace(",", ".")
                End Select
            Case "Min", "Max", "Average", "Sum"
                'создаем массив для дальнейшей работы
                Dim maxArr() As Double = {}
                Array.Resize(maxArr, funcParams.Length)
                'перебираем в цикле все параметры
                For i = 0 To funcParams.Length - 1
                    If mScript.Param_GetType(funcParams(i)) = MatewScript.ReturnFormatEnum.TO_NUMBER Then
                        maxArr(i) = Val(funcParams(i)) 'число - просто сохраняем в maxArr()
                    ElseIf mScript.Param_GetType(funcParams(i)) = MatewScript.ReturnFormatEnum.TO_STRING Then
                        'строка - это имя переменной-массива. Надо получить массив и выбрать из него наибольшее/наименьшее/среднее/сумму
                        Dim fPar As String = mScript.PrepareStringToPrint(funcParams(i), arrParams) 'имя переменной
                        'поиск переменной в локальных, переменных исследования и глобальных
                        Dim varArray() As String = mScript.csLocalVariables.GetVariableArray(fPar)
                        If varArray(0) = "#Error" Then
                            varArray = mScript.csPublicVariables.GetVariableArray(fPar)
                            If varArray(0) = "#Error" Then
                                varArray = mScript.csPublicVariables.GetVariableArray(fPar)
                                mScript.LAST_ERROR = "Ошибка при получении массива переменной " + fPar + ". Переменная не найдена."
                                Return "#Error"
                            End If
                        End If
                        'массив значений получен
                        'создаем временный массив, в который сохраняем значения массива из переменной в виде чисел (т. к. они сейчас в виде строк)
                        Dim tempArr() As Double = {}
                        Array.Resize(tempArr, varArray.Length)
                        For j As Integer = 0 To varArray.Length - 1
                            Select Case mScript.Param_GetType(varArray(j))
                                Case MatewScript.ReturnFormatEnum.TO_NUMBER
                                    tempArr(j) = Val(varArray(j))
                                Case MatewScript.ReturnFormatEnum.TO_BOOL
                                    If varArray(j) = "True" Then
                                        tempArr(j) = 1
                                    Else
                                        tempArr(j) = 0
                                    End If
                                Case MatewScript.ReturnFormatEnum.TO_STRING
                                    mScript.LAST_ERROR = "В массиве " + fPar + " определяются строковые значения. Допускаются только числа или True/False."
                                    Return "#Error"
                            End Select
                        Next
                        'выполняем тебуемую функцию и сохраняем результат в maxArr()
                        Select Case functionName
                            Case "Max"
                                maxArr(i) = tempArr.Max
                            Case "Min"
                                maxArr(i) = tempArr.Min
                            Case "Average"
                                maxArr(i) = tempArr.Average
                            Case "Sum"
                                maxArr(i) = tempArr.Sum
                        End Select
                    ElseIf mScript.Param_GetType(funcParams(i)) = MatewScript.ReturnFormatEnum.TO_BOOL Then
                        If funcParams(i) = "True" Then
                            maxArr(i) = 1
                        Else
                            maxArr(i) = 0
                        End If
                    End If
                Next
                Select Case functionName
                    Case "Max"
                        Return Convert.ToString(maxArr.Max, provider_points)
                    Case "Min"
                        Return Convert.ToString(maxArr.Min, provider_points)
                    Case "Average"
                        Return Convert.ToString(maxArr.Average, provider_points)
                    Case "Sum"
                        Return Convert.ToString(maxArr.Sum, provider_points)
                End Select
        End Select

        Return _ERROR("Неизвестная функция класса " & mScript.mainClass(classId).Names.Last, functionName)
    End Function

    ''' <summary>
    ''' Процедура установаливает значение свойства элементу любого из 3 уровней
    ''' </summary>
    ''' <param name="classId">Класс свойства</param>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="newValue">Новое значение свойства</param>
    ''' <param name="child2Id">Id элемента 2 уровня (-1 если надо установить значение по умолчанию)</param>
    ''' <param name="child3Id">Id элемента 3 уровня(-1 если надо установить значение элементу 2 уровня или по умолчанию)</param>
    Public Sub SetPropertyValue(ByVal classId As Integer, ByVal propName As String, ByVal newValue As String, ByVal child2Id As Integer, Optional child3Id As Integer = -1, Optional ByVal ignoreBattle As Boolean = False)
        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(newValue)

        If propName = "Name" AndAlso mScript.mainClass(classId).Properties.ContainsKey("Caption") AndAlso child2Id >= 0 Then
            'Изменения Caption при изменении Name (если и раньше было Caption = Name)
            Dim propCaption As MatewScript.PropertiesInfoType = Nothing
            If mScript.mainClass(classId).Properties.TryGetValue("Caption", propCaption) Then
                'существует свойство Caption. Если его значение идентично с прежним значением имени - копируем в него новое значение
                Dim captValue As String = "", nameValue As String = ""
                If child3Id >= 0 Then
                    captValue = mScript.mainClass(classId).ChildProperties(child2Id)("Caption").ThirdLevelProperties(child3Id)
                    nameValue = mScript.mainClass(classId).ChildProperties(child2Id)("Name").ThirdLevelProperties(child3Id)
                Else
                    If GVARS.G_ISBATTLE AndAlso ignoreBattle = False AndAlso mScript.mainClass(classId).Names(0) = "H" Then
                        captValue = mScript.Battle.Fighters(child2Id).heroProps("Caption").Value
                        nameValue = mScript.Battle.Fighters(child2Id).heroProps("Name").Value
                    Else
                        captValue = mScript.mainClass(classId).ChildProperties(child2Id)("Caption").Value
                        nameValue = mScript.mainClass(classId).ChildProperties(child2Id)("Name").Value
                    End If
                End If
                If captValue = nameValue Then
                    SetPropertyValue(classId, "Caption", newValue, child2Id, child3Id)
                End If
            End If
            SetPropertyValue(classId, "Caption", newValue, child2Id, child3Id, ignoreBattle)
        ElseIf propName.EndsWith("Total") AndAlso mScript.mainClass(classId).Properties.ContainsKey(Left(propName, propName.Length - 5)) Then
            'Существует пара свойств xxx & xxxTotal
            Dim propDefName As String = Left(propName, propName.Length - 5)
            Dim prevPropValue As Double = 0, prevPropTotal As Double = 0
            If child2Id < 0 Then
                Double.TryParse(mScript.mainClass(classId).Properties(propDefName).Value.Replace(".", ","), prevPropValue)
                Double.TryParse(mScript.mainClass(classId).Properties(propName).Value.Replace(".", ","), prevPropTotal)
            ElseIf child3Id < 0 Then
                If GVARS.G_ISBATTLE AndAlso ignoreBattle = False AndAlso mScript.mainClass(classId).Names(0) = "H" Then
                    Double.TryParse(mScript.Battle.Fighters(child2Id).heroProps(propDefName).Value.Replace(".", ","), prevPropValue)
                    Double.TryParse(mScript.Battle.Fighters(child2Id).heroProps(propName).Value.Replace(".", ","), prevPropTotal)
                Else
                    Double.TryParse(mScript.mainClass(classId).ChildProperties(child2Id)(propDefName).Value.Replace(".", ","), prevPropValue)
                    Double.TryParse(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value.Replace(".", ","), prevPropTotal)
                End If
            Else
                Double.TryParse(mScript.mainClass(classId).ChildProperties(child2Id)(propDefName).ThirdLevelProperties(child3Id).Replace(".", ","), prevPropValue)
                Double.TryParse(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id).Replace(".", ","), prevPropTotal)
            End If
            Dim diff As Double = prevPropTotal - prevPropValue
            Dim dblNewTotal As Double = 0
            Double.TryParse(newValue.Replace(".", ","), dblNewTotal)
            Dim dblNewValue As Double = dblNewTotal - diff
            SetPropertyValue(classId, propDefName, dblNewValue.ToString, child2Id, child3Id, ignoreBattle)
        ElseIf mScript.mainClass(classId).Properties(propName).returnType = MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS Then
            cListManager.UpdateCSSInFile(classId, child2Id, child3Id, propName, newValue)
        End If

        If cRes <> MatewScript.ContainsCodeEnum.NOT_CODE Then
            mScript.eventRouter.SetPropertyWithEvent(classId, propName, newValue, child2Id, child3Id, cRes, False, ignoreBattle)
            Return
        End If
        If child2Id < 0 Then
            Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties(propName)
            p.Value = newValue
            If p.eventId > 0 Then
                mScript.eventRouter.RemoveEvent(p.eventId) 'сменилось значение свойства с исполняемого на неисполняемое - убираем событие
                p.eventId = -1
            End If
        ElseIf child3Id < 0 Then
            Dim p As MatewScript.ChildPropertiesInfoType
            If GVARS.G_ISBATTLE AndAlso ignoreBattle = False AndAlso mScript.mainClass(classId).Names(0) = "H" Then
                p = mScript.Battle.Fighters(child2Id).heroProps(propName)
            Else
                p = mScript.mainClass(classId).ChildProperties(child2Id)(propName)
            End If
            p.Value = newValue
            If p.eventId > 0 Then
                mScript.eventRouter.RemoveEvent(p.eventId) 'сменилось значение свойства с исполняемого на неисполняемое - убираем событие
                p.eventId = -1
            End If
        Else
            Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(propName)
            p.ThirdLevelProperties(child3Id) = newValue
            If p.ThirdLevelEventId(child3Id) > 0 Then
                mScript.eventRouter.RemoveEvent(p.ThirdLevelEventId(child3Id)) 'сменилось значение свойства с исполняемого на неисполняемое - убираем событие
                p.ThirdLevelEventId(child3Id) = -1
            End If
        End If
        If questEnvironment.EDIT_MODE = False Then mScript.trackingProperties.RunAfter(classId, propName, {child2Id.ToString, child3Id.ToString})
    End Sub

    ''' <summary>
    ''' Возвращает значение свойства если она является целым числом
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    ''' <param name="propName">Имя свойста</param>
    ''' <param name="child2Id">Id элемента 2 уровня</param>
    ''' <param name="child3Id">Id элемента 2 уровня</param>
    ''' <param name="returnOnError">Что вернуть в случае ошибки</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPropertyValueAsInteger(ByVal classId As Integer, ByVal propName As String, ByRef child2Id As Integer, ByVal child3Id As Integer, Optional ByVal returnOnError As Integer = -1) As Integer
        On Error GoTo err
        Dim res As Integer = -1
        If child2Id < 0 Then
            If Integer.TryParse(mScript.PrepareStringToPrint(mScript.mainClass(classId).Properties(propName).Value, Nothing, False), res) = False Then Return returnOnError
        ElseIf child3Id < 0 Then
            If Integer.TryParse(mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False), res) = False Then Return returnOnError
        Else
            If Integer.TryParse(mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False), res) = False Then Return returnOnError
        End If

        Return res
err:
        Return returnOnError
    End Function

    ''' <summary>
    ''' Возвращает Id текстового значения в массиве Functions(functionName).params(paramIndex).EnumValues()
    ''' </summary>
    ''' <param name="classId">Класс функции</param>
    ''' <param name="functionName">Имя функции</param>
    ''' <param name="paramIndex">Номер параметра в функции, начиная с 0</param>
    ''' <param name="enumName">Имя, индекс в EnumValues() которого надо получить</param>
    Private Function GetEnumIdByName(ByVal classId As Integer, ByVal functionName As String, ByVal paramIndex As Integer, ByVal enumName As String) As Integer
        Dim arrValues() As String = mScript.mainClass(classId).Functions(functionName).params(paramIndex).EnumValues
        If IsNothing(arrValues) OrElse arrValues.Count = 0 Then Return -1
        For i As Integer = 0 To arrValues.Count - 1
            If arrValues(i) = enumName Then Return i
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Функция возвращает текстовое значение массива из mScript.mainClass(classId).Functions(functionName).returnArray по индексу массива
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    ''' <param name="functionName">Имя функции</param>
    ''' <param name="enumId">Id  в массиве returnArray</param>
    ''' <param name="onErrorValue">Значение, которое следует вернуть если возникла ошибка</param>
    Private Function GetFunctionReturnEnumTextById(ByVal classId As Integer, ByVal functionName As String, ByVal enumId As Integer, Optional ByVal onErrorValue As String = "#Error") As String
        Dim arr() As String = mScript.mainClass(classId).Functions(functionName).returnArray
        If IsNothing(arr) OrElse enumId < 0 OrElse enumId >= arr.Count Then
            If onErrorValue = "#Error" Then
                Return _ERROR("Ошибка в структуре возвращаемых значений функции.", functionName)
            Else
                Return onErrorValue
            End If
        End If
        Return arr(enumId)
    End Function

    ''' <summary>
    ''' Устанавливает последнюю ошибку скрипта и возвращает "#Error"
    ''' </summary>
    ''' <param name="errorString">Текст ошибки</param>
    ''' <param name="functionName">Имя функции, в которой произошла ошибка</param>
    ''' <returns>Всегда #Error</returns>
    Public Function _ERROR(ByVal errorString As String, Optional functionName As String = "") As String
        If functionName.Length = 0 Then
            mScript.LAST_ERROR = errorString
        Else
            mScript.LAST_ERROR = "Функция " + Chr(34) + functionName + Chr(34) + ": " + errorString
        End If
        Return "#Error"
    End Function

    ''' <summary>
    ''' Выполняет исполняемые строки в параметрах funcParams
    ''' </summary>
    ''' <param name="funcParams">массив funcParams</param>
    ''' <param name="arrParams">параметры запуска исполняемой среды</param>
    Private Function ExecuteParams(ByRef funcParams() As String, ByRef arrParams() As String) As Boolean
        If IsNothing(funcParams) OrElse funcParams.Length = 0 Then Return True
        For i As Integer = 0 To funcParams.Count - 1
            Dim pType As MatewScript.ReturnFormatEnum = mScript.Param_GetType(funcParams(i))
            If pType = MatewScript.ReturnFormatEnum.TO_EXECUTABLE_STRING Then
                funcParams(i) = mScript.ExecuteString({funcParams(i)}, arrParams)
                If funcParams(i) = "#Error" Then Return False
            ElseIf pType = MatewScript.ReturnFormatEnum.TO_STRING AndAlso GVARS.G_CURLOC > -1 AndAlso funcParams(i).StartsWith("'#") AndAlso funcParams(i).EndsWith(":'") Then
                Dim selNumber As String = funcParams(i).Substring(2, funcParams(i).Length - 4)
                If IsNumeric(selNumber) Then
                    'это селектор. Получаем его содержимое
                    Dim res As String = ""
                    If ReadProperty(mScript.mainClassHash("L"), "Description", GVARS.G_CURLOC, -1, res, Nothing, MatewScript.ReturnFormatEnum.ORIGINAL, Val(selNumber)) Then
                        funcParams(i) = WrapString(res)
                    End If
                End If
            End If
        Next i
        Return True
    End Function

    ''' <summary>
    ''' Проверка правильности количества и формата параметров функции
    ''' </summary>
    ''' <param name="classId">Id класса, содержащего данную функцию</param>
    ''' <param name="functionName">Имя функции</param>
    ''' <param name="funcParams">Ссылка на массив параметров, переданной функции (уже обработанных до простых значений)</param>
    Private Function CheckFunctionParams(ByVal classId As Integer, ByVal functionName As String, ByRef funcParams() As String) As Boolean
        'Проверка правильности количества параметров функции
        'If mScript.mainClass(classId).Functions(functionName).UserAdded = True Then Return True

        Dim parCount As Integer
        If IsNothing(funcParams) Then
            parCount = 0
        Else
            parCount = funcParams.Length
        End If

        If parCount < mScript.mainClass(classId).Functions(functionName).paramsMin Then
            mScript.LAST_ERROR = "Указано недостаточное количество параметров в функции " + functionName + "."
            Return False
        End If
        If mScript.mainClass(classId).Functions(functionName).paramsMax <> -1 AndAlso parCount > mScript.mainClass(classId).Functions(functionName).paramsMax Then
            mScript.LAST_ERROR = "Слишком много параметров в функции " + functionName + "."
            Return False
        End If
        'Проверка правильности формата параметров функции
        If IsNothing(funcParams) OrElse funcParams.Count = 0 Then Return True
        Dim pt As MatewScript.paramsType.paramsTypeEnum
        For i As Integer = 0 To funcParams.Count - 1
            If i > mScript.mainClass(classId).Functions(functionName).params.Count - 1 Then
                pt = mScript.mainClass(classId).Functions(functionName).params(mScript.mainClass(classId).Functions(functionName).params.Count - 1).Type 'Param Array
            Else
                pt = mScript.mainClass(classId).Functions(functionName).params(i).Type
            End If
            If pt = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY Then Continue For
            Dim rf As MatewScript.ReturnFormatEnum = mScript.Param_GetType(funcParams(i))
            Select Case rf
                Case MatewScript.ReturnFormatEnum.TO_STRING
                    If pt = MatewScript.paramsType.paramsTypeEnum.PARAM_BOOL OrElse pt = MatewScript.paramsType.paramsTypeEnum.PARAM_INTEGER OrElse pt = MatewScript.paramsType.paramsTypeEnum.PARAM_SINGLE Then
                        mScript.LAST_ERROR = "Формат параметра №" + (i + 1).ToString + ", переданного функции " + functionName + ", не верен (не может быть строкой)."
                        Return False
                    End If
                Case MatewScript.ReturnFormatEnum.TO_NUMBER
                    If pt = MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM OrElse pt = MatewScript.paramsType.paramsTypeEnum.PARAM_STRING OrElse pt = MatewScript.paramsType.paramsTypeEnum.PARAM_USER_FUNCTION _
                        OrElse pt = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY OrElse pt = MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_AUDIO OrElse _
                        pt = MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_CSS OrElse pt = MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_JS OrElse _
                        pt = MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_PICTURE OrElse pt = MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_TEXT OrElse _
                        pt = MatewScript.paramsType.paramsTypeEnum.PARAM_EVENT OrElse pt = MatewScript.paramsType.paramsTypeEnum.PARAM_PROPERTY Then
                        mScript.LAST_ERROR = "Формат параметра №" + (i + 1).ToString + ", переданного функции " + functionName + ", не верен (не может быть числом)."
                        Return False
                    ElseIf pt = MatewScript.paramsType.paramsTypeEnum.PARAM_INTEGER AndAlso funcParams(i).IndexOf(".") > -1 Then
                        mScript.LAST_ERROR = "Формат параметра №" + (i + 1).ToString + ", переданного функции " + functionName + ", не верен (число должно быть целым)."
                        Return False
                    End If
                Case MatewScript.ReturnFormatEnum.TO_BOOL
                    If pt <> MatewScript.paramsType.paramsTypeEnum.PARAM_ANY AndAlso pt <> MatewScript.paramsType.paramsTypeEnum.PARAM_BOOL AndAlso pt <> MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY Then
                        mScript.LAST_ERROR = "Формат параметра №" + (i + 1).ToString + ", переданного функции " + functionName + ", не верен (не может быть Да/Нет)."
                        Return False
                    End If
                Case MatewScript.ReturnFormatEnum.TO_ARRAY
                    mScript.LAST_ERROR = "Формат параметра №" + (i + 1).ToString + ", переданного функции " + functionName + ", не верен (массив не может служить параметром функции)."
                    Return False
            End Select
        Next
        Return True
    End Function

    Public Function CreateNewObject(ByVal classId As Integer, ByRef funcParams() As String, Optional ByVal forceSecondLevel As Boolean = False) As String
        'Создает в указанном классе объект второго или тетьего порядка
        Dim newId As Integer 'Id нового объекта
        If IsNothing(funcParams) OrElse funcParams.Length = 0 Then
            mScript.LAST_ERROR = "Не указано имя создаваемого объекта."
            Return "#Error"
        End If
        If funcParams.GetUpperBound(0) = 0 OrElse forceSecondLevel Then
            'funcParams(0) - имя создаваемого объекта второго порядка
            If IsSecondChildNameExists(mScript.mainClass(classId).ChildProperties, funcParams(0)) Then
                mScript.LAST_ERROR = "Такое имя уже существует."
                Return "#Error"
            End If
            'Получаем Id нового объекта
            If IsNothing(mScript.mainClass(classId).ChildProperties) Then
                newId = 0
            Else
                newId = mScript.mainClass(classId).ChildProperties.Length
            End If
            'заполняем свойства объекта из свойств по-умолчанию
            ReDim Preserve mScript.mainClass(classId).ChildProperties(newId)
            mScript.mainClass(classId).ChildProperties(newId) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
            Dim curPropHash As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classId).ChildProperties(newId)
            For Each parentProp As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(classId).Properties
                Dim eventId As Integer = 0
                Dim newValue As String
                If parentProp.Value.returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION OrElse parentProp.Value.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT AndAlso parentProp.Value.Hidden <> MatewScript.PropertyHiddenEnum.LEVEL2_ONLY Then
                    newValue = ""
                Else
                    newValue = parentProp.Value.Value
                    If parentProp.Value.eventId > 0 Then eventId = mScript.eventRouter.DuplicateEvent(parentProp.Value.eventId)
                End If
                curPropHash.Add(parentProp.Key, New MatewScript.ChildPropertiesInfoType With {.Value = newValue, .eventId = eventId})
            Next
            'устанавливаем свойство Name
            Dim curProp As MatewScript.ChildPropertiesInfoType = curPropHash("Name")
            curProp.Value = funcParams(0)
            curPropHash("Name") = curProp
            If curPropHash.TryGetValue("Caption", curProp) Then
                curProp.Value = funcParams(0)
                curPropHash("Caption") = curProp
            End If
            curProp = curPropHash("Group")
            curProp.Value = "''"
        Else
            'создаем объект 3 порядка
            'funcParams(0) - имя существующего объекта 2 порядка
            'funcParams(1) - имя создаваемого объекта 3 порядка
            Dim level2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If level2Id = -1 Then
                mScript.LAST_ERROR = "Нет объекта с Id = " + funcParams(0) + "."
                Return "#Error"
            ElseIf level2Id = -2 Then
                mScript.LAST_ERROR = "Объекта с именем " + funcParams(0) + " не существует."
                Return "#Error"
            End If
            If IsThirdChildNameExists(mScript.mainClass(classId).ChildProperties, level2Id, funcParams(0)) Then
                mScript.LAST_ERROR = "Такое имя уже существует."
                Return "#Error"
            End If
            'Получаем Id нового объекта
            If IsNothing(mScript.mainClass(classId).ChildProperties(level2Id)("Name").ThirdLevelProperties) Then
                newId = 0
            Else
                newId = mScript.mainClass(classId).ChildProperties(level2Id)("Name").ThirdLevelProperties.Length
            End If
            'заполняем свойства нового объекта 3 порядка из свойств его родителя 2-го порядка
            Dim parentProp As MatewScript.ChildPropertiesInfoType
            Dim propCopy As New Hashtable(mScript.mainClass(classId).ChildProperties(level2Id))
            For Each prop As DictionaryEntry In propCopy
                parentProp = mScript.mainClass(classId).ChildProperties(level2Id)(prop.Key)
                Dim eventId As Integer = parentProp.eventId

                ReDim Preserve parentProp.ThirdLevelProperties(newId)
                ReDim Preserve parentProp.ThirdLevelEventId(newId)
                Dim newValue As String
                If mScript.mainClass(classId).Properties(prop.Key).returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION OrElse _
                    mScript.mainClass(classId).Properties(prop.Key).returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT AndAlso mScript.mainClass(classId).Properties(prop.Key).Hidden <> _
                    MatewScript.PropertyHiddenEnum.LEVEL3_ONLY AndAlso mScript.mainClass(classId).Properties(prop.Key).Hidden <> MatewScript.PropertyHiddenEnum.LEVEL23_ONLY Then
                    newValue = ""
                Else
                    newValue = prop.Value.Value
                    If eventId > 0 Then parentProp.ThirdLevelEventId(newId) = mScript.eventRouter.DuplicateEvent(eventId)
                End If
                parentProp.ThirdLevelProperties(newId) = newValue
            Next
            mScript.mainClass(classId).ChildProperties(level2Id)("Name").ThirdLevelProperties(newId) = funcParams(1)
            mScript.mainClass(classId).ChildProperties(level2Id)("Group").ThirdLevelProperties(newId) = "''"
            If mScript.mainClass(classId).ChildProperties(level2Id).ContainsKey("Caption") Then mScript.mainClass(classId).ChildProperties(level2Id)("Caption").ThirdLevelProperties(newId) = funcParams(1)
        End If
        Return newId.ToString
    End Function

    ''' <summary>
    ''' Удаляет в указанном классе объект второго или тетьего порядка
    ''' </summary>
    ''' <param name="classId">Класс удаляемого элемента</param>
    ''' <param name="funcParams">Список параметров, переданной функции удаления</param>
    Public Function RemoveObject(ByVal classId As Integer, ByRef funcParams() As String) As String
        'Удаляет в указанном классе объект второго или тетьего порядка
        If IsNothing(funcParams) OrElse funcParams.Length = 0 OrElse (funcParams.Count = 1 AndAlso funcParams(0) = "-1") Then
            'удаляет все дочерние элементы (и 2, и 3 порядка)
            If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso mScript.mainClass(classId).ChildProperties.Count > 0 Then
                'удаляем все события
                For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                    Dim p2 As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classId).ChildProperties(i)
                    For Each chld As KeyValuePair(Of String, MatewScript.ChildPropertiesInfoType) In p2 'для каждого свойства
                        If chld.Value.eventId > 0 Then mScript.eventRouter.RemoveEvent(chld.Value.eventId) 'удаляем события элемента 2 порядка
                        If IsNothing(chld.Value.ThirdLevelProperties) = False AndAlso chld.Value.ThirdLevelProperties.Count > 0 Then
                            For j As Integer = 0 To chld.Value.ThirdLevelEventId.Count - 1
                                If chld.Value.ThirdLevelEventId(j) > 0 Then mScript.eventRouter.RemoveEvent(chld.Value.ThirdLevelEventId(j)) 'удаляем события элемента 3 порядка
                            Next j
                        End If
                    Next
                Next i
            End If
            Erase mScript.mainClass(classId).ChildProperties
            Return "0"
        End If
        If funcParams.GetUpperBound(0) = 0 Then
            'funcParams(0) - имя удаляемого объекта второго порядка
            'Получаем Id удаляемого объекта
            Dim level2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If level2Id = -1 Then
                mScript.LAST_ERROR = "Нет объекта с Id = " + funcParams(0) + "."
                Return "#Error"
            ElseIf level2Id = -2 Then
                mScript.LAST_ERROR = "Объекта с именем " + funcParams(0) + " не существует."
                Return "#Error"
            End If
            'удаляем события
            Dim p2 As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classId).ChildProperties(level2Id)
            For Each chld As KeyValuePair(Of String, MatewScript.ChildPropertiesInfoType) In p2 'для каждого свойства
                If chld.Value.eventId > 0 Then mScript.eventRouter.RemoveEvent(chld.Value.eventId) 'удаляем события элемента 2 порядка
                If IsNothing(chld.Value.ThirdLevelProperties) = False AndAlso chld.Value.ThirdLevelProperties.Count > 0 Then
                    For j As Integer = 0 To chld.Value.ThirdLevelEventId.Count - 1
                        If chld.Value.ThirdLevelEventId(j) > 0 Then mScript.eventRouter.RemoveEvent(chld.Value.ThirdLevelEventId(j)) 'удаляем события элемента 3 порядка
                    Next j
                End If
            Next
            'удаляем объект
            For i As Integer = level2Id To mScript.mainClass(classId).ChildProperties.GetUpperBound(0) - 1
                mScript.mainClass(classId).ChildProperties(i) = mScript.mainClass(classId).ChildProperties(i + 1)
            Next
            If mScript.mainClass(classId).ChildProperties.GetUpperBound(0) = 0 Then
                Erase mScript.mainClass(classId).ChildProperties
                Return "0"
            Else
                ReDim Preserve mScript.mainClass(classId).ChildProperties(mScript.mainClass(classId).ChildProperties.GetUpperBound(0) - 1)
                Return mScript.mainClass(classId).ChildProperties.GetUpperBound(0).ToString
            End If
        Else
            'удаляем объект 3 порядка
            'funcParams(0) - имя объекта 2 порядка (родителя удаляемого)
            'funcParams(1) - имя объекта 3 порядка (самого удаляемого)
            Dim level2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If level2Id = -1 Then
                mScript.LAST_ERROR = "Нет объекта с Id = " + funcParams(0) + "."
                Return "#Error"
            ElseIf level2Id = -2 Then
                mScript.LAST_ERROR = "Объекта с именем " + funcParams(0) + " не существует."
                Return "#Error"
            End If
            Dim level3Id As Integer = GetThirdChildIdByName(funcParams(1), level2Id, mScript.mainClass(classId).ChildProperties)
            If level3Id = -1 Then
                mScript.LAST_ERROR = "Нет объекта с Id = [" + funcParams(0) + ", " + funcParams(1) + "]."
                Return "#Error"
            ElseIf level2Id = -2 Then
                mScript.LAST_ERROR = "Объекта с именем [" + funcParams(0) + ", " + funcParams(1) + "] не существует."
                Return "#Error"
            End If

            'удаляем события
            Dim p2 As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classId).ChildProperties(level2Id)
            For Each chld As KeyValuePair(Of String, MatewScript.ChildPropertiesInfoType) In p2 'для каждого свойства
                If chld.Value.ThirdLevelEventId(level3Id) > 0 Then mScript.eventRouter.RemoveEvent(chld.Value.ThirdLevelEventId(level3Id)) 'удаляем события элемента 3 порядка
            Next

            Dim curProp As MatewScript.ChildPropertiesInfoType
            Dim propCopy As New Hashtable(mScript.mainClass(classId).ChildProperties(level2Id))
            For Each prop As DictionaryEntry In propCopy
                curProp = mScript.mainClass(classId).ChildProperties(level2Id)(prop.Key)
                If IsNothing(curProp.ThirdLevelProperties) OrElse curProp.ThirdLevelProperties.GetUpperBound(0) = 0 Then
                    Erase curProp.ThirdLevelProperties
                    Erase curProp.ThirdLevelEventId
                Else
                    For i As Integer = level3Id To curProp.ThirdLevelProperties.GetUpperBound(0) - 1
                        curProp.ThirdLevelProperties(i) = curProp.ThirdLevelProperties(i + 1)
                        curProp.ThirdLevelEventId(i) = curProp.ThirdLevelEventId(i + 1)
                    Next
                    ReDim Preserve curProp.ThirdLevelProperties(curProp.ThirdLevelProperties.GetUpperBound(0) - 1)
                    ReDim Preserve curProp.ThirdLevelEventId(curProp.ThirdLevelEventId.GetUpperBound(0) - 1)
                End If
            Next
            If IsNothing(mScript.mainClass(classId).ChildProperties(level2Id)("Name").ThirdLevelProperties) Then Return "0"
            Return mScript.mainClass(classId).ChildProperties(level2Id)("Name").ThirdLevelProperties.GetUpperBound(0).ToString
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Получает Id объекта второго или тетьего порядка
    ''' </summary>
    ''' <param name="classId">Имя класса</param>
    ''' <param name="funcParams">Параметры, переданные движку при выполнении кода, содержащие ссылки на элемент(ы) 2-го (и 3-го) порядка</param>
    ''' <param name="getThirdLevelId">Получать ли при возможности Id элемента 3-го уровня или все же второго</param>
    Public Function ObjectId(ByVal classId As Integer, ByRef funcParams() As String, Optional getThirdLevelId As Boolean = True) As String
        'Получает Id объекта второго или тетьего порядка
        If IsNothing(funcParams) OrElse funcParams.Length = 0 Then
            mScript.LAST_ERROR = "Не указано имя объекта."
            Return "#Error"
        End If
        If funcParams.GetUpperBound(0) = 0 Then
            'funcParams(0) - имя объекта второго порядка
            'Получаем Id объекта
            Dim level2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If level2Id = -1 Then
                Return "-1"
                mScript.LAST_ERROR = "Нет объекта с Id = " + funcParams(0) + "."
                Return "#Error"
            ElseIf level2Id = -2 Then
                Return "-1"
                mScript.LAST_ERROR = "Объекта с именем " + funcParams(0) + " не существует."
                Return "#Error"
            End If
            'возвращаем Id
            Return level2Id.ToString
        Else
            'объект 3 порядка
            'funcParams(0) - имя объекта 2 порядка (родителя)
            'funcParams(1) - имя объекта 3 порядка (самого)
            Dim level2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If level2Id = -1 Then
                mScript.LAST_ERROR = "Нет объекта с Id = " + funcParams(0) + "."
                Return "-1"
                Return "#Error"
            ElseIf level2Id = -2 Then
                mScript.LAST_ERROR = "Объекта с именем " + funcParams(0) + " не существует."
                Return "-1"
                Return "#Error"
            End If
            If Not getThirdLevelId Then Return level2Id.ToString

            Dim level3Id As Integer = GetThirdChildIdByName(funcParams(1), level2Id, mScript.mainClass(classId).ChildProperties)
            If level3Id = -1 Then
                mScript.LAST_ERROR = "Нет объекта с Id = [" + funcParams(0) + ", " + funcParams(1) + "]."
                Return "-1"
                Return "#Error"
            ElseIf level2Id = -2 Then
                mScript.LAST_ERROR = "Объекта с именем [" + funcParams(0) + ", " + funcParams(1) + "] не существует."
                Return "-1"
                Return "#Error"
            End If
            'возвращаем Id
            Return level3Id.ToString
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Получает Id объекта второго или третьего порядка
    ''' </summary>
    ''' <param name="classId">Id класса, которому принадлежит элемент</param>
    ''' <param name="funcParams">Парметры, переданые функции (с id элементов)</param>
    Public Function ObjectIsExists(ByVal classId As Integer, ByRef funcParams() As String) As String
        If IsNothing(funcParams) OrElse funcParams.Length = 0 Then
            If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Length = 0 Then
                Return "False"
            Else
                Return "True"
            End If
        End If
        If funcParams.GetUpperBound(0) = 0 Then
            'funcParams(0) - имя объекта второго порядка
            'Получаем Id искомого объекта
            Dim level2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            Return IIf(level2Id < 0, "False", "True")
        Else
            'объект 3 порядка
            'funcParams(0) - имя объекта 2 порядка (родителя)
            'funcParams(1) - имя объекта 3 порядка (самого)
            Dim level2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If level2Id < 0 Then Return "False"
            Dim level3Id As Integer = GetThirdChildIdByName(funcParams(1), level2Id, mScript.mainClass(classId).ChildProperties)
            Return IIf(level3Id < 0, "False", "True")
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Получает кол-во объектов второго или тетьего порядка
    ''' </summary>
    ''' <param name="classId">Имя класса</param>
    ''' <param name="funcParams">Параметры, переданные движку при выполнении кода, содержащие ссылки на элемент(ы) 2-го (и 3-го) порядка</param>
    ''' <param name="getThirdLevelId">Считать ли при возможности Id элемента 3-го уровня или все же второго</param>
    Private Function ObjectsCount(ByVal classId As Integer, ByRef funcParams() As String, Optional getThirdLevelId As Boolean = True) As String
        'Получает кол-во объектов второго или тетьего порядка
        If IsNothing(funcParams) OrElse funcParams.Length = 0 OrElse (IsNothing(funcParams) = False AndAlso funcParams(0) = "-1") Then
            If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Length = 0 Then
                Return "0"
            Else
                Return mScript.mainClass(classId).ChildProperties.Length.ToString
            End If
        Else
            'funcParams(0) - имя объекта второго порядка
            'Получаем Id объекта
            Dim level2Id As Integer = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
            If level2Id < 0 Then Return "0"
            'Возвращаем кол-во элементов 2-го порядка
            If Not getThirdLevelId Then Return mScript.mainClass(classId).ChildProperties.Count
            'Если здесь, то надо посчитать кол-во элементов 3-го порядка
            Dim prop() As String = mScript.mainClass(classId).ChildProperties(level2Id)("Name").ThirdLevelProperties
            If IsNothing(prop) OrElse prop.Length = 0 Then
                Return "0"
            Else
                Return prop.Length
            End If
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Добавляет новую функцию классу с детальными настройками, указанными в propData
    ''' </summary>
    ''' <param name="classId">Класс, в который добавляется</param>
    ''' <param name="funcName">Имя новой функции</param>
    ''' <param name="funcData">Данные касательно функции</param>
    Public Function AddUserFunctionByPropData(ByVal classId As Integer, ByVal funcName As String, ByVal funcData As MatewScript.PropertiesInfoType) As String
        'добавляет функцию классу
        If mScript.mainClass(classId).Functions.ContainsKey(funcName) Then
            mScript.LAST_ERROR = "В классе " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.GetUpperBound(0)) + " функция с именем " + funcName + " уже существует."
            Return "#Error"
        End If
        If IsNothing(funcData.Value) Then funcData.Value = ""

        funcData.editorIndex = mScript.mainClass(classId).Functions.Count
        mScript.mainClass(classId).Functions.Add(funcName, funcData)
        Return ""
    End Function

    ''' <summary>
    ''' Добавляет новую функцию классу
    ''' </summary>
    ''' <param name="funcParams">Список параметров, переданных функции</param>
    ''' <param name="arrParams">Параметры, с которыми запущен код (если недоступны, то Nothing)</param>
    Public Function AddUserFunction(ByRef funcParams() As String, ByRef arrParams() As String) As String
        'добавляет функцию классу
        If IsNothing(funcParams) OrElse funcParams.Length < 2 Then
            mScript.LAST_ERROR = "Указано неверное количество параметров."
            Return "#Error"
        End If
        'funcParams(0) - класс новой функции
        Dim className As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
        If className = "#Error" Then Return "#Error"
        Dim classId As Integer = -1
        If Not mScript.mainClassHash.TryGetValue(className, classId) Then
            Return _ERROR("Не удалось получить класс в который надо добавить функцию.", "AddFunction")
        End If
        'funcParams(1) - имя новой функции
        Dim funcName As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
        If funcName = "#Error" Then Return "#Error"
        If mScript.mainClass(classId).Functions.ContainsKey(funcName) Then
            mScript.LAST_ERROR = "В классе " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.GetUpperBound(0)) + " функция с именем " + funcName + " уже существует."
            Return "#Error"
        End If
        Dim editorId As Integer = mScript.mainClass(classId).Functions.Count
        mScript.mainClass(classId).Functions.Add(funcName, New MatewScript.PropertiesInfoType With {.UserAdded = True, .EditorCaption = funcName, .Value = "", .editorIndex = editorId})
        Return ""
    End Function

    ''' <summary>
    ''' Удаляет функцию из класса
    ''' </summary>
    ''' <param name="funcParams">Список параметров, переданных функции</param>
    ''' <param name="arrParams">Параметры, с которыми запущен код (если недоступны, то Nothing)</param>
    Public Function RemoveUserFunction(ByRef funcParams() As String, ByRef arrParams() As String, Optional ByVal dontCheckForInternalProperties As Boolean = False) As String
        'удаляет функцию класса
        If IsNothing(funcParams) OrElse funcParams.Length < 2 Then
            mScript.LAST_ERROR = "Указано неверное количество параметров."
            Return "#Error"
        End If
        'funcParams(0) - класс удаляемой функции
        Dim className As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
        If className = "#Error" Then Return "#Error"
        Dim classId As Integer = -1
        If Not mScript.mainClassHash.TryGetValue(className, classId) Then
            Return _ERROR("Не удалось получить класс из которого надо удалить функцию.", "RemoveFunction")
        End If
        'funcParams(1) - имя удаляемой функции
        Dim funcName As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
        If funcName = "#Error" Then Return "#Error"
        If dontCheckForInternalProperties = False AndAlso mScript.mainClass(classId).Functions(funcName).UserAdded = False Then
            mScript.LAST_ERROR = "Функция " + funcName + " не может быть удалена, так как является встроенной."
            Return "#Error"
        End If
        If mScript.mainClass(classId).Functions.ContainsKey(funcName) = False Then
            mScript.LAST_ERROR = "Функции " + funcName + " в классе " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.GetUpperBound(0)) + " не найдено."
            Return "#Error"
        End If

        'editorId -1 в последующих функциях
        Dim editorId As Integer = mScript.mainClass(classId).Functions(funcName).editorIndex
        If editorId < mScript.mainClass(classId).Functions.Count - 1 Then
            For i As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                Dim pValue As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(i).Value
                If pValue.editorIndex > editorId Then
                    pValue.editorIndex -= 1
                End If
            Next i
        End If

        If mScript.mainClass(classId).Functions(funcName).eventId > 0 Then mScript.eventRouter.RemoveEvent(mScript.mainClass(classId).Functions(funcName).eventId)

        mScript.mainClass(classId).Functions.Remove(funcName)
        Return ""
    End Function

    ''' <summary>
    ''' Добавляет новое свойство классу
    ''' </summary>
    ''' <param name="funcParams">Список параметров, переданных функции</param>
    ''' <param name="arrParams">Параметры, с которыми запущен код (если недоступны, то Nothing)</param>
    Private Function AddUserProperty(ByRef funcParams() As String, ByRef arrParams() As String) As String
        'добавляет свойство классу
        If IsNothing(funcParams) OrElse funcParams.Length < 2 Then
            mScript.LAST_ERROR = "Указано неверное количество параметров."
            Return "#Error"
        End If
        'funcParams(0) - класс нового свойства
        Dim className As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
        If className = "#Error" Then Return "#Error"
        Dim classId As Integer = -1
        If Not mScript.mainClassHash.TryGetValue(className, classId) Then
            Return _ERROR("Не удалось получить класс в который надо добавить свойство.", "AddProperty")
        End If
        'funcParams(1) - имя нового свойства
        'funcParams(2) - имя его значения по-умолчанию
        Dim propName As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
        If propName = "#Error" Then Return "#Error"
        Dim propValue As String
        If funcParams.Length = 2 Then
            propValue = ""
        Else
            propValue = funcParams(1)
        End If

        If mScript.mainClass(classId).Properties.ContainsKey(propName) Then
            mScript.LAST_ERROR = "В классе " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.GetUpperBound(0)) + " свойство с именем " + propName + " уже существует."
            Return "#Error"
        End If
        Dim editorId As Integer = mScript.mainClass(classId).Properties.Count
        mScript.mainClass(classId).Properties.Add(propName, New MatewScript.PropertiesInfoType With {.Value = propValue, .UserAdded = True, .EditorCaption = propName, .editorIndex = editorId})
        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(propValue)
        Dim eventId As Integer = -1
        If cRes <> MatewScript.ContainsCodeEnum.NOT_CODE Then
            eventId = mScript.eventRouter.SetEventId(classId, propName, propValue, cRes, -1, -1)
        End If

        If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso mScript.mainClass(classId).ChildProperties.Length > 0 Then
            Dim newProp As MatewScript.ChildPropertiesInfoType
            For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.GetUpperBound(0)
                Dim p As New MatewScript.ChildPropertiesInfoType With {.Value = propValue}
                If eventId > 0 Then p.eventId = mScript.eventRouter.DuplicateEvent(eventId)
                mScript.mainClass(classId).ChildProperties(i).Add(propName, p)


                If IsNothing(mScript.mainClass(classId).ChildProperties(i)("Name").ThirdLevelProperties) = False AndAlso mScript.mainClass(classId).ChildProperties(i)("Name").ThirdLevelProperties.GetUpperBound(0) > 0 Then
                    newProp = mScript.mainClass(classId).ChildProperties(i)(propName)
                    ReDim newProp.ThirdLevelProperties(mScript.mainClass(classId).ChildProperties(i)("Name").ThirdLevelProperties.GetUpperBound(0))
                    ReDim newProp.ThirdLevelEventId(mScript.mainClass(classId).ChildProperties(i)("Name").ThirdLevelProperties.GetUpperBound(0))
                    For j As Integer = 0 To newProp.ThirdLevelProperties.GetUpperBound(0)
                        newProp.ThirdLevelProperties(j) = propValue
                        If eventId > 0 Then newProp.ThirdLevelEventId(j) = mScript.eventRouter.DuplicateEvent(eventId)
                    Next
                End If
            Next
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Добавляет новое свойство классу с детальными настройками, указанными в propData
    ''' </summary>
    ''' <param name="classId">Класс, в который добавляется</param>
    ''' <param name="propName">Имя нового свойства</param>
    ''' <param name="propData">Данные касательно свойства</param>
    Public Function AddUserPropertyByPropData(ByVal classId As Integer, ByVal propName As String, ByVal propData As MatewScript.PropertiesInfoType) As String
        'добавляет свойство классу

        If mScript.mainClass(classId).Properties.ContainsKey(propName) Then
            mScript.LAST_ERROR = "В классе " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.GetUpperBound(0)) + " свойство с именем " + propName + " уже существует."
            Return "#Error"
        End If
        If IsNothing(propData.Value) Then propData.Value = ""
        Dim editorId As Integer = mScript.mainClass(classId).Properties.Count
        propData.editorIndex = editorId
        mScript.mainClass(classId).Properties.Add(propName, propData)
        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(propData.Value)
        Dim eventId As Integer = -1
        If cRes <> MatewScript.ContainsCodeEnum.NOT_CODE Then
            eventId = mScript.eventRouter.SetEventId(classId, propName, propData.Value, cRes, -1, -1)
        End If

        If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso mScript.mainClass(classId).ChildProperties.Length > 0 Then
            For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.GetUpperBound(0)
                Dim p As New MatewScript.ChildPropertiesInfoType With {.Value = propData.Value}
                If eventId > 0 Then p.eventId = mScript.eventRouter.DuplicateEvent(eventId)
                If mScript.mainClass(classId).ChildProperties(i).ContainsKey(propName) = False Then
                    mScript.mainClass(classId).ChildProperties(i).Add(propName, p)
                Else
                    mScript.mainClass(classId).ChildProperties(i)(propName) = p
                End If

                If IsNothing(mScript.mainClass(classId).ChildProperties(i)("Name").ThirdLevelProperties) = False AndAlso mScript.mainClass(classId).ChildProperties(i)("Name").ThirdLevelProperties.Count > 0 Then
                    Dim newProp As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(i)(propName)
                    ReDim newProp.ThirdLevelProperties(mScript.mainClass(classId).ChildProperties(i)("Name").ThirdLevelProperties.GetUpperBound(0))
                    ReDim newProp.ThirdLevelEventId(mScript.mainClass(classId).ChildProperties(i)("Name").ThirdLevelProperties.GetUpperBound(0))
                    For j As Integer = 0 To newProp.ThirdLevelProperties.GetUpperBound(0)
                        newProp.ThirdLevelProperties(j) = propData.Value
                        If eventId > 0 Then newProp.ThirdLevelEventId(j) = mScript.eventRouter.DuplicateEvent(eventId)
                    Next j
                End If
            Next i
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Изменяем свойство класса по детальными настройкам, указанным в propData, сохраняя дочерние элементы
    ''' </summary>
    ''' <param name="classId">Класс изменяемого свойства</param>
    ''' <param name="propOldName">Старое имя свойства</param>
    ''' <param name="propNewName">Новое имя свойства</param>
    ''' <param name="propNewData">Данные касательно свойства</param>
    Public Function UpdateUserPropertyByPropData(ByVal classId As Integer, ByVal propOldName As String, ByVal propNewName As String, ByVal propNewData As MatewScript.PropertiesInfoType) As String
        'изменяет свойство класса

        If mScript.mainClass(classId).Properties.ContainsKey(propOldName) = False Then
            mScript.LAST_ERROR = "В классе " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.GetUpperBound(0)) + " свойство с именем " + propOldName + " не существует."
            Return "#Error"
        End If
        If IsNothing(propNewData.Value) Then propNewData.Value = ""
        Dim propOldData As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties(propOldName)

        With propOldData
            SetPropertyValue(classId, propOldName, propNewData.Value, -1)
            propNewData.eventId = .eventId
            propNewData.editorIndex = .editorIndex
        End With

        If propNewName = propOldName Then
            'имя не изменилось - дочерние свойства не трогаем, меняем только главное
            mScript.mainClass(classId).Properties(propNewName) = propNewData
            Return ""
        End If

        'имя изменилось

        'меняем главное свойство
        mScript.mainClass(classId).Properties.Remove(propOldName)
        mScript.mainClass(classId).Properties.Add(propNewName, propNewData)

        If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso mScript.mainClass(classId).ChildProperties.Length > 0 Then
            'меняем дочерние свойства
            For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(i)(propOldName)
                mScript.mainClass(classId).ChildProperties(i).Remove(propOldName)
                mScript.mainClass(classId).ChildProperties(i).Add(propNewName, ch)
            Next i
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Удаляет свойство из класса
    ''' </summary>
    ''' <param name="funcParams">Список параметров, переданных функции</param>
    ''' <param name="arrParams">Параметры, с которыми запущен код (если недоступны, то Nothing)</param>
    Public Function RemoveUserProperty(ByRef funcParams() As String, ByRef arrParams() As String, Optional ByVal dontCheckForInternalProperties As Boolean = False) As String
        'удаляет свойство из класса
        If IsNothing(funcParams) OrElse funcParams.Length < 2 Then
            mScript.LAST_ERROR = "Указано неверное количество параметров."
            Return "#Error"
        End If
        'funcParams(0) - класс удаляемого свойства
        Dim className As String = mScript.PrepareStringToPrint(funcParams(0), arrParams)
        If className = "#Error" Then Return "#Error"
        Dim classId As Integer = -1
        If Not mScript.mainClassHash.TryGetValue(className, classId) Then
            Return _ERROR("Не удалось получить класс из которого надо удалить свойство.", "RemoveProperty")
        End If
        'funcParams(1) - имя удаляемого свойства
        Dim propName As String = mScript.PrepareStringToPrint(funcParams(1), arrParams)
        If propName = "#Error" Then Return "#Error"
        If mScript.mainClass(classId).Properties(propName).UserAdded = False AndAlso dontCheckForInternalProperties = False Then
            mScript.LAST_ERROR = "Свойство " + propName + " не может быть удалено, так как является встроенным."
            Return "#Error"
        End If
        If mScript.mainClass(classId).Properties.ContainsKey(propName) = False Then
            mScript.LAST_ERROR = "В классе " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.GetUpperBound(0)) + " свойство " + propName + " не существует."
            Return "#Error"
        End If

        mScript.eventRouter.RemoveEvent(mScript.mainClass(classId).Properties(propName).eventId) 'удаляем событие 1 уровня
        'editorId -1 в последующих свойствах
        Dim editorId As Integer = mScript.mainClass(classId).Properties(propName).editorIndex
        If editorId < mScript.mainClass(classId).Properties.Count - 1 Then
            For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                Dim pValue As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(i).Value
                If pValue.editorIndex > editorId Then
                    pValue.editorIndex -= 1
                End If
            Next i
        End If

        mScript.mainClass(classId).Properties.Remove(propName) 'удаляем из свойств по-умолчанию (1 порядка)
        If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso mScript.mainClass(classId).ChildProperties.GetUpperBound(0) <> -1 AndAlso mScript.mainClass(classId).ChildProperties(0).Count > 0 Then
            'удаляем из объектов 2 и 3 порядка
            For i As Integer = mScript.mainClass(classId).ChildProperties.GetUpperBound(0) To 0 Step -1
                Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(i)(propName)
                'удаляем события 2 и 3 уровня
                If p.eventId > 0 Then mScript.eventRouter.RemoveEvent(p.eventId)
                If IsNothing(p.ThirdLevelEventId) = False AndAlso p.ThirdLevelProperties.Count > 0 Then
                    For j As Integer = 0 To p.ThirdLevelProperties.Count - 1
                        If p.ThirdLevelEventId(j) > 0 Then mScript.eventRouter.RemoveEvent(p.ThirdLevelEventId(j))
                    Next j
                End If
                'удаляем свойство 2 (и 3) уровня
                mScript.mainClass(classId).ChildProperties(i).Remove(propName)
            Next
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Функция обновляет индексы отображения функций на панели управления данного класса в редакторе после операций множественного удаления функций (и вставки)
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    Public Sub UpdateFunctionsEditorIndexes(ByVal classId As Integer)
        If IsNothing(mScript.mainClass(classId).Functions) OrElse mScript.mainClass(classId).Functions.Count = 0 Then Return
        Dim arrNames() As String = Nothing, arrPos As Integer = -1 'массив имен функций которым индекс уже был обновлен
        ReDim arrNames(mScript.mainClass(classId).Functions.Count - 1)

        For i As Integer = mScript.mainClass(classId).Functions.Count - 1 To 0 Step -1
            'editorId функции, у которой на данный момент этот показатель наибольший, становится равным i
            Dim mVal As Integer = Integer.MinValue
            Dim mName As String = ""
            For j As Integer = mScript.mainClass(classId).Functions.Count - 1 To 0 Step -1
                'поиск функции с максимальным текущим значением editorId
                If mScript.mainClass(classId).Functions.ElementAt(j).Value.editorIndex > mVal Then
                    If arrNames.Contains(mScript.mainClass(classId).Functions.ElementAt(j).Key) Then Continue For 'исключаем те функции, которым индекс уже обновили
                    mVal = mScript.mainClass(classId).Functions.ElementAt(j).Value.editorIndex
                    mName = mScript.mainClass(classId).Functions.ElementAt(j).Key
                End If
            Next j
            arrPos += 1
            arrNames(arrPos) = mName
            'непосредственное обновление индекса
            Dim f As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions(mName)
            f.editorIndex = i
        Next i
    End Sub

    ''' <summary>
    ''' Функция определяет, существует ли объект 2-го порядка с таким именем (для организации уникальности имен)
    ''' </summary>
    ''' <param name="childProperties">Ссылка на childProperties нужного класса mainClass(x)</param>
    ''' <param name="childName">Имя искомого элемента второго порядка</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsSecondChildNameExists(ByRef childProperties() As SortedList(Of String, MatewScript.ChildPropertiesInfoType), ByVal childName As String) As Boolean
        If IsNothing(childProperties) Then Return False

        For i As Integer = 0 To childProperties.GetUpperBound(0)
            If childProperties(i)("Name").Value = childName Then Return True
        Next

        Return False
    End Function

    Private Function IsThirdChildNameExists(ByRef childProperties() As SortedList(Of String, MatewScript.ChildPropertiesInfoType), ByVal secondLevelId As Integer, ByVal childName As String) As Boolean
        'Функция определяет, существует ли объект 3-го порядка с таким именем (для организации уникальности имен)
        If IsNothing(childProperties) Then Return False
        If secondLevelId < 0 Then Return False
        If IsNothing(childProperties(secondLevelId)) OrElse childProperties.GetUpperBound(0) < secondLevelId Then Return False
        If IsNothing(childProperties(secondLevelId)("Name").ThirdLevelProperties) Then Return False

        For i As Integer = 0 To childProperties(secondLevelId)("Name").ThirdLevelProperties.GetUpperBound(0)
            If childProperties(secondLevelId)("Name").ThirdLevelProperties(i) = childName Then Return True
        Next

        Return False
    End Function

    ''' <summary>Функция возвращает Id объекта 2-го порядка по его имени. Если такого имени нет - возвращает -2
    '''Также происходит проверка, не передан ли Id в strName. Если объекта с таким Id нет - возвращает -1</summary>
    ''' <param name="strName">Имя/Id искомого объекта</param>
    ''' <param name="childProperties">Ссылка на массив childProperties() нужного класса mainClass(x)</param>
    Public Function GetSecondChildIdByName(ByVal strName As String, ByRef childProperties() As SortedList(Of String, MatewScript.ChildPropertiesInfoType)) As Integer
        Dim childId As Integer
        If Integer.TryParse(strName, childId) Then
            If childId < 0 Then Return -1
            If IsNothing(childProperties) Then Return -1
            If childProperties.GetUpperBound(0) < childId Then Return -1
            Return childId
        End If

        If IsNothing(childProperties) Then Return -2
        For i = 0 To childProperties.GetUpperBound(0)
            If childProperties(i)("Name").Value = strName Then Return i
        Next
        Return -2
    End Function

    Public Function GetThirdChildIdByName(ByVal strName As String, ByVal secondLevelId As Integer, ByRef childProperties() As SortedList(Of String, MatewScript.ChildPropertiesInfoType)) As Integer
        'Функция возвращает Id объекта 3-го порядка по его имени. Если такого имени нет - возвращает -2
        'Также происходит проверка, не передан ли Id в strName. Если объекта с таким Id нет - возвращает -1
        'Проверка правильности secondLevelId не здесь происходит, так как она проводилась ранее
        Dim childId As Integer
        If Integer.TryParse(strName, childId) Then
            If childId < 0 Then Return -1
            If IsNothing(childProperties) Then Return -1
            If IsNothing(childProperties(secondLevelId)("Name").ThirdLevelProperties) Then Return -1
            If childProperties(secondLevelId)("Name").ThirdLevelProperties.GetUpperBound(0) < childId Then Return -1
            Return childId
        End If

        If IsNothing(childProperties) Then Return -1
        If IsNothing(childProperties(secondLevelId)("Name").ThirdLevelProperties) Then Return -1
        Dim prop() As String = childProperties(secondLevelId)("Name").ThirdLevelProperties
        childId = Array.IndexOf(prop, strName)
        If childId = -1 Then Return -2
        Return childId
    End Function

    ''' <summary>
    ''' Возвращает элемент массива с указанным индексом или, если его нет, defValue
    ''' </summary>
    ''' <param name="arrParams">массив параметров</param>
    ''' <param name="id">индекс нужного элемента массива</param>
    ''' <param name="defValue">значение, которое слдует получить при ошибке</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetParam(ByRef arrParams() As String, ByVal id As Integer, ByVal defValue As String) As String
        If IsNothing(arrParams) OrElse id >= arrParams.Count Then Return defValue
        Return arrParams(id)
    End Function


    ''' <summary>
    ''' Получаем класс и имя свойства или функции по строке вида '[ClassName].PropOrFunctionName'
    ''' </summary>
    ''' <param name="strWord">Строка в кавычках</param>
    ''' <param name="classId">ссылка для получения класса</param>
    ''' <param name="elementName">ссылка для получения имени элемента</param>
    ''' <param name="seekInFunctions">искать в функциях или свойствах</param>
    ''' <param name="seekIfClassNotSpecified">Искать если класс не указан явно - 'PropOrFunctionName'</param>
    ''' <param name="ignoreIfNotExist">Проводить или нет проверку на существование данной функции/свойства</param>
    ''' <returns>False при ошибке</returns>
    Public Function GetClassIdAndElementNameByString(ByVal strWord As String, ByRef classId As Integer, ByRef elementName As String, ByVal seekInFunctions As Boolean, _
                                                      Optional ByVal seekIfClassNotSpecified As Boolean = True, Optional ByVal ignoreIfNotExist As Boolean = False,
                                                      Optional ByRef isTracking As frmMainEditor.trackingcodeEnum = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT, Optional ByRef arrParams() As String = Nothing) As Boolean
        classId = -1
        If questEnvironment.EDIT_MODE Then
            strWord = mScript.PrepareStringToPrint(strWord, Nothing, False)
        Else
            strWord = mScript.PrepareStringToPrint(strWord, arrParams)
        End If
        isTracking = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT
        If strWord.EndsWith(":changed") Then
            strWord = strWord.Substring(0, strWord.Length - 8)
            isTracking = frmMainEditor.trackingcodeEnum.EVENT_AFTER
        ElseIf strWord.EndsWith(":changing") Then
            strWord = strWord.Substring(0, strWord.Length - 9)
            isTracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE
        End If

        Dim arrEvent() As String = strWord.Split("."c) '0 - ClassName, 1 - EventName/FunctionName
        If IsNothing(arrEvent) = False AndAlso arrEvent.Count = 2 Then
            'Event 'ClassName.EventName',... или Function 'ClassName.FunctionName',...
            If mScript.mainClassHash.TryGetValue(arrEvent(0), classId) Then
                elementName = arrEvent(1)
                If ignoreIfNotExist Then Return True
                Dim prop As MatewScript.PropertiesInfoType = Nothing
                If seekInFunctions Then
                    mScript.mainClass(classId).Functions.TryGetValue(elementName, prop) 'получено имя функции
                Else
                    mScript.mainClass(classId).Properties.TryGetValue(elementName, prop) 'получено имя события
                End If
                If IsNothing(prop) Then Return False
            End If
            Return (classId > -1)
        ElseIf IsNothing(arrEvent) = False AndAlso arrEvent.Count = 1 Then
            'Event 'EventName',... или Function 'FunctionName',...
            If Not seekIfClassNotSpecified Then Return False
            elementName = strWord.ToString
            If seekInFunctions Then
                'Ищем функцию среди всех функций всех классов
                For i As Integer = 0 To mScript.mainClass.Count - 1
                    If mScript.mainClass(i).Functions.ContainsKey(elementName) Then
                        classId = i
                        Exit For
                    End If
                Next
            Else
                'Ищем свойство среди всех свойств всех классов
                For i As Integer = 0 To mScript.mainClass.Count - 1
                    If mScript.mainClass(i).Properties.ContainsKey(elementName) Then
                        classId = i
                        Exit For
                    End If
                Next
            End If
            Return (classId > -1)
        End If
        Return False
    End Function


#Region "Class Rearranging"

    ''' <summary>
    ''' Проводит реорганизацию класса mScript.mainClass(classId).ChildProperties при изменении порядка следования его элементов
    ''' </summary>
    ''' <param name="sourceTree">Дерево с именами элементов</param>
    ''' <param name="classId">Id класса</param>
    ''' <param name="child2Id">Родительский Id. -1 если реорганизовываестя 2 уровень класса</param>
    ''' <param name="hidenShown">Показаны ли скрытые элементы в указанном дереве sourceTree</param>
    Public Sub RearrangeClass(ByRef sourceTree As TreeView, ByVal classId As Integer, ByVal child2Id As Integer, ByVal hidenShown As Boolean)
        If sourceTree.Nodes.Count = 0 Then Return

        Dim n As TreeNode = Nothing
        If child2Id = -1 Then
            'реорганизация 2 уровня
            Dim lstCopy As New List(Of SortedList(Of String, MatewScript.ChildPropertiesInfoType)) 'пустая стуруктура mainClass(classId).ChildProperties
            Do
                n = GetNextNode(sourceTree, n) 'получаем следующий узел-негруппу
                If IsNothing(n) Then Exit Do
                Dim childId As Integer = GetSecondChildIdByName("'" + n.Text + "'", mScript.mainClass(classId).ChildProperties) 'получаем старое Id в ChildProperties
                If childId < 0 Then Continue Do

                lstCopy.Add(mScript.mainClass(classId).ChildProperties(childId)) 'ставим элемент на новое место
            Loop
            If Not hidenShown Then InsertHidden(classId, mScript.mainClass(classId).ChildProperties, lstCopy)
            mScript.mainClass(classId).ChildProperties = lstCopy.ToArray
        Else
            'реорганизация 3 уровня
            Dim arrNames() As String = mScript.mainClass(classId).ChildProperties(child2Id)("Name").ThirdLevelProperties
            If IsNothing(arrNames) OrElse arrNames.Count = 0 Then Return
            Dim lstId As New List(Of Integer) 'список, в котором помещаются Id элементов в новой правильной последовательности
            Do
                n = GetNextNode(sourceTree, n) 'получаем следующий узел-негруппу
                If IsNothing(n) Then Exit Do
                Dim childId As Integer = GetThirdChildIdByName("'" + n.Text + "'", child2Id, mScript.mainClass(classId).ChildProperties)
                If childId < 0 Then Continue Do
                lstId.Add(childId) 'сохраняем id в списке
            Loop
            'далее в каждом свойстве элемента 2 уровня, (в котором хранятся соответстывующие свойства элементов 3 уровня) переставляем значения свойст 3 уровня в правильном порядке
            For propIndex As Integer = 0 To mScript.mainClass(classId).ChildProperties(child2Id).Count - 1
                Dim propName As String = mScript.mainClass(classId).ChildProperties(child2Id).ElementAt(propIndex).Key
                Dim propValue As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id).ElementAt(propIndex).Value

                Dim arrValues(lstId.Count - 1) As String
                Dim arrEvents(lstId.Count - 1) As Integer
                For el3Id As Integer = 0 To arrValues.Count - 1
                    arrValues(el3Id) = propValue.ThirdLevelProperties(lstId(el3Id)) 'получаем список свойства в правильном порядке
                    arrEvents(el3Id) = propValue.ThirdLevelEventId(lstId(el3Id)) 'получаем список событий в правильном порядке
                Next el3Id
                Dim chInf As MatewScript.ChildPropertiesInfoType = propValue
                chInf.ThirdLevelProperties = arrValues 'заменяем старый список значений на новый, в правильном порядке
                chInf.ThirdLevelEventId = arrEvents 'заменяем старый список событий на новый, в правильном порядке
            Next propIndex
        End If

    End Sub

    ''' <summary>
    ''' Функция, при реорганизации класса с помощью RearrangeClass, если скрытые элементы в дереве не отображены, пытается вставить все скрытые элементы на правильные места, недалеко
    ''' от того места, где они были раньше.
    ''' </summary>
    ''' <param name="classId">Id реорганизовываемого класса</param>
    ''' <param name="childProp">массив со структурами ChildPropertiesInfoType данного класса</param>
    ''' <param name="lstCopy">заполненная копия структур ChildPropertiesInfoType после реорганизации, без скрытых элементов</param>
    ''' <remarks>Используется только внутри функции RearrangeClass</remarks>
    Private Sub InsertHidden(ByVal classId As Integer, ByRef childProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType), _
                             ByRef lstCopy As List(Of SortedList(Of String, MatewScript.ChildPropertiesInfoType)))

        For i As Integer = 0 To childProp.Count - 1
            Dim hiddenChild As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = childProp(i) 'получаем  ChildPropertiesInfoType следующего элемента
            Dim hiddenName As String = hiddenChild("Name").Value 'имя элемента
            Dim blnFound As Boolean = False
            For j As Integer = 0 To lstCopy.Count - 1
                If lstCopy(j)("Name").Value = hiddenName Then
                    blnFound = True
                    Exit For
                End If
            Next
            If blnFound Then Continue For
            'здесь известно, что элемент точно скрытый (его нет в lstCopy)
            Dim hiddenGroup As String = hiddenChild("Group").Value 'группа скрытого эелмента

            '1. На первом этапе находим элементы справа и слева от исходного положения скрытого элемента в массиве ChildPropertiesInfoType, на расстоянии seekDistance от исходного положения
            '(сначала справа, потом - слева). При этом группа этого жлемента должна совпадать с группой скрытого. Если такой будет найден - вставляем скрытый элемент перед ним
            Dim seekDistance As Integer = -1
            Dim endFound As Boolean = False, startFound As Boolean = False 'найден ли конец массива права и слева
            Dim insertId As Integer = -1 'индекс в массиве lstCopy, на который надо вставить наш скрытый элемент. Пока этот индекс не найден, то -1
            Do While Not (endFound AndAlso startFound) = True 'выполняем пока не будут найдена оба конца массива
                seekDistance = seekDistance + 1 'увеличиваем "радиус обзора"
                If i + seekDistance > lstCopy.Count - 1 Then
                    endFound = True
                Else
                    If lstCopy(i + seekDistance)("Group").Value = hiddenGroup Then
                        insertId = i + seekDistance 'найден элемент справа от скрытого и из той же группы - вставляем скрытый на его место
                        Exit Do
                    End If
                End If

                If seekDistance = 0 Then Continue Do

                If i - seekDistance < 0 Then
                    startFound = True
                Else
                    If lstCopy(i - seekDistance)("Group").Value = hiddenGroup Then
                        insertId = i - seekDistance + 1 'найден элемент слева от скрытого и из той же группы - вставляем скрытый сразу за ним
                        Exit Do
                    End If
                End If
            Loop

            If insertId > -1 Then
                lstCopy.Insert(insertId, hiddenChild) 'insertId получено, вставляем скрытый и ищем дальше другие скрытые элементы
                Continue For
            End If

            '2.а В группе, в которую входит скрытый элемент, нет видимых элементов (или скрыта вся группа)
            'На этом этапе ищем группу, следующую за группой нашего скрытого элемента, у которой есть хоть один видимый элемент в lstCopy.
            'Если таковая находится, то вставляем скрытый перед первым элемемнтом следующей группы
            hiddenGroup = mScript.PrepareStringToPrint(hiddenGroup, Nothing, False)
            Dim parentName As String = ""
            If currentClassName = "A" Then parentName = currentParentName
            Dim className As String = mScript.mainClass(classId).Names(0)
            Dim groupId As Integer = cGroups.GetGroupIdByName(className, hiddenGroup, False, parentName) 'Id группы нашего элемента
            Dim lstGroups As List(Of clsGroups.clsGroupInfo) = cGroups.dictGroups(className)
            Dim groupForInsert As String
            For j As Integer = groupId + 1 To lstGroups.Count - 1
                If lstGroups(j).isThirdLevelGroup <> False Then Continue For
                groupForInsert = "'" + lstGroups(j).Name + "'"

                For u As Integer = 0 To lstCopy.Count - 1
                    If lstCopy(u)("Group").Value = groupForInsert Then
                        insertId = u 'в группе, идущей вслед за группой нашего элемента,в массиве lstCopy есть видимый елемент. Вставляем наш на его место
                        Exit For
                    End If
                Next
                If insertId > -1 Then Exit For
            Next

            If insertId > -1 Then
                'insertId получено, вставляем скрытый и ищем дальше другие скрытые элементы
                If insertId = lstCopy.Count Then
                    lstCopy.Add(hiddenChild)
                Else
                    lstCopy.Insert(insertId, hiddenChild)
                End If
                Continue For
            End If

            '2.б За группой нашего элемента нет ни одной группы с видимыми элементами
            'Теперь и группу, следующую перед группой нашего скрытого элемента, у которой есть хоть один видимый элемент в lstCopy.
            'Если таковая находится, то вставляем скрытый перед последним элементом следующей группы
            For j As Integer = groupId - 1 To 0 Step -1
                If lstGroups(j).isThirdLevelGroup <> False Then Continue For
                groupForInsert = "'" + lstGroups(j).Name + "'"

                For u As Integer = lstCopy.Count - 1 To 0 Step -1
                    If lstCopy(u)("Group").Value = groupForInsert Then
                        insertId = u + 1 'в группе, идущей перед группой нашего элемента, в массиве lstCopy есть видимый елемент. Вставляем наш прямо перед ним
                        Exit For
                    End If
                Next
                If insertId > -1 Then Exit For
            Next

            If insertId > -1 Then
                'insertId получено, вставляем скрытый и ищем дальше другие скрытые элементы
                If insertId = lstCopy.Count Then
                    lstCopy.Add(hiddenChild)
                Else
                    lstCopy.Insert(insertId, hiddenChild)
                End If
                Continue For
            End If

            'Элемент еще не вставлен. Также не найдена ни одна группа ни после, ни до указанной.
            'То есть, нет ни элементов с той же группой, как у нашего, ни каких-либо других групп с видимыми элементами. В таком случае просто оставляем скрытый элемент на старом месте.
            lstCopy.Insert(i, hiddenChild)
        Next i
    End Sub


    ''' <summary>
    ''' Возвращает следующий за указанным узел элемента (не группы).
    ''' </summary>
    ''' <param name="tree">Дерево для поиска</param>
    ''' <param name="n">Узел, за которым надо искать. Если Nothing, то возвращается первый узел-элемент</param>
    Private Function GetNextNode(ByRef tree As TreeView, ByRef n As TreeNode) As TreeNode
        If IsNothing(n) Then
            'получаем первый узел
            If IsNothing(tree.Nodes) OrElse tree.Nodes.Count = 0 Then Return Nothing
            n = tree.Nodes(0)
            If n.Tag = "ITEM" Then Return n
            'это узел-группа
            If n.Nodes.Count > 0 Then Return n.Nodes(0)
            Return GetNextNode(tree, n) 'это пустой узел-группа - ищем дальше, начиная с него
        End If

        Dim parNode As TreeNode = n.Parent
        Dim nxtNode As TreeNode
        If IsNothing(parNode) Then
            'узлы первого порядка
            If n.Index = tree.Nodes.Count - 1 Then Return Nothing
            nxtNode = tree.Nodes(n.Index + 1)
        Else
            'узлы второго порядка (точно не группы)
            If n.Index < parNode.Nodes.Count - 1 Then Return parNode.Nodes(n.Index + 1)
            If tree.Nodes.Count - 1 = parNode.Index Then Return Nothing
            nxtNode = tree.Nodes(parNode.Index + 1) 'подузлы закончились, переходим к следующему узлу 1 порядка
        End If

        If nxtNode.Tag = "ITEM" Then Return nxtNode
        If nxtNode.Nodes.Count > 0 Then Return nxtNode.Nodes(0) 'возвращаем первый узел группы
        Return GetNextNode(tree, nxtNode) 'пустая группа, ищем от нее дальше
    End Function
#End Region

    ''' <summary>
    ''' Восстанавливает список возможных значения свойства из настроек по умолчанию
    ''' </summary>
    ''' <param name="classId">Id класса свойства</param>
    ''' <param name="propName">Имя свойства</param>
    Public Sub RestoreDefaultArray(ByVal classId As Integer, ByVal propName As String)
        Dim arr() As String = mScript.mainClassCopy(classId).Properties(propName).returnArray
        Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties(propName)

        If IsNothing(arr) OrElse arr.Count = 0 Then
            p.returnArray = Nothing
            Return
        End If

        Dim arrCopy() As String
        ReDim arrCopy(arr.Count - 1)
        Array.Copy(arr, arrCopy, arrCopy.Length)
        p.returnArray = arrCopy
    End Sub

End Module
