Public Class cVariable
    Public Class variableEditorInfoType
        'Public Name As String 'Имя переменной
        Public Description As String 'Описание переменной (в режиме редактирования)
        Public Hidden As Boolean 'Скрыть/показать переменную (в режиме редактирования)
        Public Group As String
        Public Icon As String
        Public arrValues() As String
        ''' <summary>Key - сигнатура, Value - индекс в массиве arrValues</summary>
        Public lstSingatures As SortedList(Of String, Integer)
    End Class

    Public lstVariables As New SortedList(Of String, variableEditorInfoType)(StringComparer.CurrentCultureIgnoreCase)
    ' ''' <summary>Если True - запретить уничтожение переменных</summary>
    'Public preserveVariables As Boolean = False

    'Private variableEditorInfo As New List(Of variableEditorInfoType)

    ''' <summary>Удаляет переменную с указанным именем</summary>
    ''' <param name="varName">имя удаляемой переменной</param>
    Public Sub DeleteVariable(ByVal varName As String)
        If lstVariables.ContainsKey(varName) = False Then Return
        lstVariables.Remove(varName)
    End Sub

    ''' <summary> Изменяет название переменной </summary>
    ''' <param name="oldName">старое имя перменной</param>
    ''' <param name="newName">новое имя</param>
    Public Function ChangeVariableName(ByVal oldName As String, ByVal newName As String) As Boolean
        If [String].IsNullOrWhiteSpace(oldName) OrElse [String].IsNullOrWhiteSpace(newName) Then Return False
        If lstVariables.ContainsKey(oldName) = False Then
            MsgBox("Перменная " + oldName + " не найдена.")
            Return False
        End If
        If lstVariables.ContainsKey(newName) Then
            MsgBox("Перменная " + newName + " уже существует.")
            Return False
        End If

        Dim v As variableEditorInfoType = lstVariables(oldName)
        lstVariables.Remove(oldName)
        lstVariables.Add(newName, v)
        Return True
    End Function

    ''' <summary> Функция устанавливает переменную (вызывается из исходника, а не из кода пользователя) </summary>    
    ''' <param name="varName">имя создаваемой (изменяемой) переменной переменной без кавычек</param>
    ''' <param name="varValue">ее значение</param>
    ''' <param name="varIndex">индекс в массиве (если не указан, то 0); Если -1 - создается следующий</param>
    ''' <param name="returnFormat">Формат, в который надо преобразовать. Если не указано, то ORIGINAL - без преобразования</param>
    ''' <param name="varDescription">характеристики переменной, используемые в режиме редактирования</param>
    ''' <param name="varHidden">характеристики переменной, используемые в режиме редактирования</param>
    ''' <param name="varGroup">Группа переменной</param>
    ''' <param name="varIcon">Иконка переменной</param>
    ''' <returns>Возвращает 1 - успех или -1 - ошибка</returns>
    Public Overridable Function SetVariableInternal(ByVal varName As String, ByVal varValue As String, ByVal varIndex As Integer, _
                                        Optional ByVal returnFormat As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL, Optional ByVal varDescription As String = "", _
                                        Optional ByVal varHidden As Boolean = False, Optional varGroup As String = "", Optional varIcon As String = "") As Integer
        On Error GoTo er
        'Преобразование типа, если надо
        If returnFormat <> MatewScript.ReturnFormatEnum.ORIGINAL Then varValue = mScript.ConvertElement(varValue, returnFormat)

        Dim v As variableEditorInfoType = Nothing
        If IsNothing(lstVariables) OrElse lstVariables.TryGetValue(varName, v) = False Then
            Dim varArray() As String 'Для копирования массивов
            'Переменная с указанным именем не объявлена
            'Создаем массив для новой переменной и вносим в него значение varValue в элемент с номером varIndex
            If varIndex = -1 Then varIndex = 0
            ReDim varArray(varIndex)
            varArray(varIndex) = varValue
            'Вставляем массив в хэш с переменными
            If questEnvironment.EDIT_MODE Then
                v = New variableEditorInfoType With {.Description = varDescription, .Hidden = varHidden, .arrValues = varArray, .Group = varGroup, .Icon = varIcon}
            Else
                v = New variableEditorInfoType
                v.arrValues = varArray
            End If
            lstVariables.Add(varName, v)
        Else
            'Переменная уже существует
            'Копируем массив из хэша
            'v уже содержит инфо о переменной

            'Изменяем его размер при необходимости
            If IsNothing(v.arrValues) Then
                If varIndex = -1 Then varIndex = 0
                ReDim v.arrValues(varIndex)
            ElseIf varIndex >= v.arrValues.Length Then
                ReDim Preserve v.arrValues(varIndex)
            ElseIf varIndex = -1 Then
                varIndex = v.arrValues.Length
                ReDim Preserve v.arrValues(varIndex)
            End If
            'Вносим новое значение varValue в позицию varIndex
            v.arrValues(varIndex) = varValue
            'Вносим информацию для редактирования
            v.Description = varDescription
            v.Hidden = varHidden
            v.Group = varGroup
            v.Icon = varIcon
        End If

        Return 1
er:
        Return -1
    End Function

 
    ''' <summary> Функция устанавливает переменную (вызывается из исходника, а не из кода пользователя) </summary>    
    ''' <param name="varName">имя создаваемой (изменяемой) переменной без кавычек</param>
    ''' <param name="varValue">ее значение</param>
    ''' <param name="varSignature">сигнатура перменной без кавычек</param>
    ''' <param name="returnFormat">Формат, в который надо преобразовать. Если не указано, то ORIGINAL - без преобразования</param>
    ''' <param name="varDescription">характеристики переменной, используемые в режиме редактирования</param>
    ''' <param name="varHidden">характеристики переменной, используемые в режиме редактирования</param>
    ''' <returns>Возвращает 1 - успех или -1 - ошибка</returns>
    Public Overridable Function SetVariableInternal(ByVal varName As String, ByVal varValue As String, ByVal varSignature As String, _
                                        Optional ByVal returnFormat As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL, Optional ByVal varDescription As String = "", _
                                        Optional ByVal varHidden As Boolean = False) As Integer
        On Error GoTo er
        'Преобразование типа, если надо
        If returnFormat <> MatewScript.ReturnFormatEnum.ORIGINAL Then varValue = mScript.ConvertElement(varValue, returnFormat)

        Dim v As variableEditorInfoType = Nothing
        If IsNothing(lstVariables) OrElse lstVariables.TryGetValue(varName, v) = False Then
            Dim varArray() As String 'Для копирования массивов
            'Переменная с указанным именем не объявлена
            'Создаем массив для новой переменной и вносим в него значение varValue в элемент с номером varIndex
            Dim varIndex As Integer = 0
            ReDim varArray(varIndex)
            varArray(varIndex) = varValue
            'Вставляем массив в хэш с переменными
            v = New variableEditorInfoType With {.Description = varDescription, .Hidden = varHidden, .arrValues = varArray}
            v.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
            v.lstSingatures.Add(varSignature, varIndex)
            lstVariables.Add(varName, v)
        Else
            'Переменная уже существует
            'v уже содержит инфо о переменной
            Dim varIndex As Integer = -1
            If IsNothing(v.lstSingatures) Then
                v.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
            Else
                If v.lstSingatures.ContainsKey(varSignature) Then
                    varIndex = v.lstSingatures(varSignature)
                Else
                    varIndex = -1
                End If
            End If
            If varIndex = -1 Then
                'элемента с указанной сигнатурой не существует
                varIndex = v.arrValues.Count
                ReDim Preserve v.arrValues(varIndex)
                'If IsNothing(v.lstSingatures) Then v.lstSingatures = New SortedList(Of String, Integer)
                v.lstSingatures.Add(varSignature, varIndex)
            Else
                'элемент с указанной сигнатурой уже есть

            End If
            'Вносим новое значение varValue в позицию varIndex
            v.arrValues(varIndex) = varValue
            'Вносим информацию для редактирования
            v.Description = varDescription
            v.Hidden = varHidden
        End If

        Return 1
er:
        Return -1
    End Function

    ''' <summary>
    ''' Возвращает переменную varName (точнее, элемент массива varName с индексом varIndex)
    ''' </summary>
    ''' <param name="varName">Имя переменной</param>
    ''' <param name="varIndex">индекс переменной</param>
    ''' <returns>Значение переменной, в случае ошибки #Error</returns>
    ''' <remarks></remarks>
    Public Overridable Function GetVariable(ByVal varName As String, ByVal varIndex As Integer) As String
        'Функция возвращает переменную varName (точнее, элемент массива varName с индексом varIndex)
        Dim var As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName.Trim, var) = False Then Return "#Error"
        If IsNothing(var.arrValues) OrElse varIndex >= var.arrValues.Count OrElse varIndex < 0 Then Return "#Error"
        Return var.arrValues(varIndex)
    End Function

    ''' <summary>
    ''' Возвращает переменную varName (точнее, элемент массива varName с сигнатурой varSignature)
    ''' </summary>
    ''' <param name="varName">Имя переменной</param>
    ''' <param name="varSignature">сигнатура переменной</param>
    ''' <returns>Значение переменной, в случае ошибки #Error</returns>
    ''' <remarks></remarks>
    Public Overridable Function GetVariable(ByVal varName As String, ByVal varSignature As String) As String
        Dim var As variableEditorInfoType = Nothing
        varSignature = UnWrapString(varSignature)
        If lstVariables.TryGetValue(varName, var) = False Then Return "#Error"
        If IsNothing(var.lstSingatures) Then Return "#Error"
        Dim id As Integer = 0
        If var.lstSingatures.TryGetValue(varSignature, id) = False Then Return "#Error"
        If id >= var.arrValues.Count Then Return "#Error"
        Return var.arrValues(id)
    End Function

    ''' <summary>
    ''' Возвращает весь массив переменной varName
    ''' </summary>
    ''' <param name="varName">Имя переменной</param>
    ''' <returns>массив с переменными; в случае ошибки - массив с одним элементом, равным #Error</returns>
    Public Function GetVariableArray(ByVal varName As String) As String()
        Dim var As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, var) = False Then Return {"#Error"}
        Return var.arrValues
    End Function


    ''' <summary>
    ''' Устанавливает весь массив переменной varName
    ''' </summary>
    ''' <param name="varName">Имя переменной</param>
    ''' <param name="arrValues">массив с новыми значениями переменных</param>
    ''' <returns>в случае ошибки - #Error</returns>
    Public Function SetVariableArray(ByVal varName As String, ByRef arrValues() As String) As String
        Dim v As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, v) = False Then Return "#Error"

        If IsNothing(arrValues) OrElse arrValues.Length = 0 Then
            v.arrValues = {""}
        Else
            v.arrValues = arrValues
        End If
        If IsNothing(v.lstSingatures) = False Then v.lstSingatures.Clear()
        Return ""
    End Function

    Public Sub ClearVariableArray(ByVal varName As String, Optional emptyValue As String = "")
        Dim v As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, v) = False Then Return
        v.arrValues = {emptyValue}
        v.lstSingatures.Clear()
    End Sub

    ''' <summary>
    ''' Функция совершает поиск в массиве переменной и возвращает массив с индексами найденных вхождений
    ''' </summary>
    ''' <param name="varName">имя переменной-массива</param>
    ''' <param name="strSearch">строка для поиска</param>
    ''' <param name="arrParams">параметры запуска кода</param>
    ''' <param name="start">индекс начала поиска</param>
    ''' <param name="finish">индекс конца поиска</param>
    ''' <param name="beginFromEnd">искать с начала или с конца</param>
    ''' <returns>массив найденных вхождений; если ничего не найдено - то массив с один элементом, равным -1</returns>
    Public Function InArray(ByVal varName As String, ByVal strSearch As String, ByRef arrParams() As String, Optional start As Integer = 0, Optional finish As Integer = -1, _
                            Optional beginFromEnd As Boolean = False, Optional ByVal ignoreCase As Boolean = True) As String()
        Dim v As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, v) = False Then Return Nothing

        Dim varArray() As String = v.arrValues 'получаем массив переменной
        If finish = -1 OrElse finish > varArray.Length Then finish = varArray.Length - 1 'получаем конец поиска
        If start < 0 Then start = 0

        Dim arrResult() As String = Nothing
        Dim varStep As Integer = 1
        If beginFromEnd Then
            'если поиск идет с конца - меняем местами start / finish и устанавливаем step = -1
            Dim tmpVar As Integer = start
            start = finish
            finish = tmpVar
            varStep = -1
        End If

        'выполняем поиск и схраняем индексы найденных вхождений в массив arrResult
        For i As Integer = start To finish Step varStep
            If String.Compare(mScript.PrepareStringToPrint(varArray(i), arrParams), strSearch, ignoreCase) = 0 Then
                If IsNothing(arrResult) Then
                    ReDim arrResult(0)
                Else
                    ReDim Preserve arrResult(arrResult.Length)
                End If
                arrResult(arrResult.Length - 1) = i.ToString
            End If
        Next
        If IsNothing(arrResult) Then Return {"-1"} 'ничего не найдено - возвращаем массив с одним элементом, равным -1
        Return arrResult 'возвращаем массив найденных вхождений
    End Function

    ''' <summary>
    ''' Функция совершает замены в массиве переменной и возвращает массив с заменами. При этом исходный массив не меняется.
    ''' </summary>
    ''' <param name="var">переменная-массив</param>
    ''' <param name="strSearch">строка для поиска</param>
    ''' <param name="strReplace">строка, на которую заменять</param>
    ''' <param name="replaceCount">параметр для получения количества замен</param>
    ''' <param name="arrParams">параметры запуска кода</param>
    ''' <param name="start">индекс начала поиска</param>
    ''' <param name="finish">индекс конца поиска</param>
    ''' <param name="beginFromEnd">искать с начала или с конца</param>
    ''' <returns>массив со всеми заменами, а также количество проведенных замен в переменную replaceCount</returns>
    Public Shared Function Replace(ByRef var As variableEditorInfoType, ByVal strSearch As String, ByVal strReplace As String, ByRef replaceCount As Integer, ByRef arrParams() As String, Optional start As Integer = 0, _
                            Optional finish As Integer = -1, Optional beginFromEnd As Boolean = False) As String()
        replaceCount = 0
        Dim varArray() As String = var.arrValues  'получаем массив переменной
        If finish = -1 OrElse finish > varArray.Length Then finish = varArray.Length - 1 'получчаем конец поиска
        If start < 0 Then start = 0

        'создаем копию массива переменной
        Dim arrResult() As String = Nothing
        Array.Resize(arrResult, varArray.Length)
        Array.ConstrainedCopy(varArray, 0, arrResult, 0, varArray.Length)

        Dim varStep As Integer = 1
        If beginFromEnd Then
            'если поиск идет с конца - меняем местами start / finish и устанавливаем step = -1
            Dim tmpVar As Integer = start
            start = finish
            finish = tmpVar
            varStep = -1
        End If

        strSearch = strSearch.ToLower
        'выполняем поиск и и замену в массив arrResult
        For i As Integer = start To finish Step varStep
            If mScript.PrepareStringToPrint(varArray(i), arrParams).ToLower = strSearch Then
                arrResult(i) = strReplace
                replaceCount += 1
            End If
        Next
        Return arrResult 'возвращаем массив со всеми заменами
    End Function

    Public Function JoinArray(ByVal varName As String, ByVal strSplitter As String, ByRef arrParams() As String) As String
        Dim v As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, v) = False Then Return "#Error"

        Dim varArray() As String = v.arrValues
        Dim arrResult() As String = Nothing
        Array.Resize(arrResult, varArray.Length)

        For i As Integer = 0 To varArray.Length - 1
            arrResult(i) = mScript.PrepareStringToPrint(varArray(i), arrParams)
        Next

        Return "'" + Join(arrResult, strSplitter).Replace("'", "/'") + "'"
    End Function

    ''' <summary>
    ''' Удаляет часть элементов из массива
    ''' </summary>
    ''' <param name="var">переменная-массива</param>
    ''' <param name="start">первый элемент для удаления</param>
    ''' <param name="finish">последний элемент для удаления</param>
    ''' <returns></returns>
    Public Shared Function RemoveRange(ByVal var As variableEditorInfoType, ByVal start As Integer, ByVal finish As Integer) As String
        Dim varArray() As String = var.arrValues
        If start < 0 Then start = 0
        If finish < 0 Then finish = varArray.Count - 1
        If finish > varArray.Length - 1 Then finish = varArray.Length - 1

        If varArray.Length = 0 Then Return ""
        If varArray.Length = 1 Then
            'только 1 элемент массива - очистка
            ReDim var.arrValues(0)
            var.arrValues(0) = ""
            If IsNothing(var.lstSingatures) = False Then var.lstSingatures.Clear()
            Return ""
        End If

        If IsNothing(var.lstSingatures) = False AndAlso var.lstSingatures.Count > 0 Then
            'вносим соответствующие изменения в сигнатуры
            For i As Integer = var.lstSingatures.Count - 1 To 0 Step -1
                Dim varIndex As Integer = var.lstSingatures.ElementAt(i).Value
                If varIndex < start Then Continue For
                If varIndex < finish Then
                    var.lstSingatures.RemoveAt(i)
                Else
                    Dim varSign As String = var.lstSingatures.ElementAt(i).Key
                    var.lstSingatures(varSign) = varIndex - (finish - start + 1)
                End If
            Next
        End If

        Dim tmpList As List(Of String) = varArray.ToList
        tmpList.RemoveRange(start, finish - start + 1)
        If tmpList.Count = 0 Then
            var.arrValues = {""}
        Else
            var.arrValues = tmpList.ToArray
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Удаляет элемент массива с указанным индексом
    ''' </summary>
    ''' <param name="varName">имя массива</param>
    ''' <param name="elementId">индекс удаляемого элемента</param>
    Public Overridable Function RemoveArrayElement(ByVal varName As String, ByVal elementId As Integer) As String
        Dim v As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, v) = False Then Return "#Error"

        Dim varArray() As String = v.arrValues
        If elementId < 0 Then elementId = 0
        If varArray.Length = 0 Then Return ""
        If varArray.Length = 1 Then
            v.arrValues = {""}
            v.lstSingatures.Clear()
            Return ""
        End If
        If elementId > varArray.Length - 1 Then elementId = varArray.Length - 1

        If IsNothing(v.lstSingatures) = False AndAlso v.lstSingatures.Count > 0 Then
            'вносим соответствующие изменения в сигнатуры
            For i As Integer = v.lstSingatures.Count - 1 To 0 Step -1
                Dim varIndex As Integer = v.lstSingatures.ElementAt(i).Value
                If varIndex < elementId Then Continue For
                If varIndex = elementId Then
                    v.lstSingatures.RemoveAt(i)
                Else
                    Dim varSign As String = v.lstSingatures.ElementAt(i).Key
                    v.lstSingatures(varSign) = varIndex - 1
                End If
            Next
        End If

        Dim tmpList As List(Of String) = varArray.ToList
        tmpList.RemoveAt(elementId)
        If tmpList.Count = 0 Then
            v.arrValues = {""}
        Else
            v.arrValues = tmpList.ToArray
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Удаляет элемент массива с указанной сигнатурой
    ''' </summary>
    ''' <param name="varName">имя массива</param>
    ''' <param name="elementSignature">сигнатура удаляемого элемента</param>
    Public Overridable Function RemoveArrayElement(ByVal varName As String, ByVal elementSignature As String) As String
        Dim v As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, v) = False Then Return "#Error"
        Dim varIndex As Integer = 0
        If v.lstSingatures.TryGetValue(elementSignature, varIndex) = False Then Return ""
        Return RemoveArrayElement(varName, varIndex)
    End Function

    ''' <summary>
    ''' Вставлет элемент в массив в указанное место
    ''' </summary>
    ''' <param name="varName">имя переменной-массива</param>
    ''' <param name="elementIndex">индекс, куда вставлять</param>
    ''' <param name="varValue">значение переменной</param>
    ''' <param name="varSignature">сигнатура без кавычек</param>
    ''' <returns></returns>
    Public Function Insert(ByVal varName As String, ByVal elementIndex As Integer, ByVal varValue As String, Optional ByVal varSignature As String = "") As String
        Dim v As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, v) = False Then Return "#Error"

        Dim varArray() As String = v.arrValues
        If varArray.Count = 1 AndAlso varArray(0) = "" Then
            varArray(0) = varValue
            If String.IsNullOrEmpty(varSignature) = False Then
                v.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                v.lstSingatures.Add(varSignature, 0)
            End If
            Return ""
        End If

        If elementIndex < 0 Then elementIndex = 0
        If elementIndex > varArray.Count Then elementIndex = varArray.Count
        If IsNothing(v.lstSingatures) = False AndAlso v.lstSingatures.Count > 0 Then
            'вносим соответствующие изменения в сигнатуры
            For i As Integer = 0 To v.lstSingatures.Count - 1
                Dim varIndex As Integer = v.lstSingatures.ElementAt(i).Value
                If varIndex >= elementIndex Then
                    Dim varSign As String = v.lstSingatures.ElementAt(i).Key
                    v.lstSingatures(varSign) = varIndex + 1
                End If
            Next
        End If

        Dim tmpList As List(Of String) = varArray.ToList
        tmpList.Insert(elementIndex, varValue)
        v.arrValues = tmpList.ToArray

        If String.IsNullOrEmpty(varSignature) = False Then
            If IsNothing(v.lstSingatures) Then v.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
            v.lstSingatures.Add(varSignature, elementIndex)
        End If
        Return ""
    End Function

    Public Shared Function GetStructure(ByRef var As variableEditorInfoType) As String
        Dim varArray() As String = var.arrValues
        Dim strResult As New System.Text.StringBuilder
        For i As Integer = 0 To varArray.Count - 1
            If IsNothing(var.lstSingatures) = False AndAlso var.lstSingatures.Count > 0 Then
                Dim res As Integer = var.lstSingatures.IndexOfValue(i)
                If res > -1 Then
                    If i >= varArray.Count - 1 Then
                        strResult.Append(var.lstSingatures.ElementAt(res).Key + " => " + varArray(i))
                    Else
                        strResult.AppendLine(var.lstSingatures.ElementAt(res).Key + " => " + varArray(i))
                    End If
                    Continue For
                End If
            End If
            strResult.AppendLine(i.ToString + " => " + varArray(i))
        Next
        Return "'" + strResult.ToString.Replace("'", "/'") + "'"
    End Function

    Public Sub KillVars()
        lstVariables.Clear()
    End Sub

    ''' <summary>
    ''' Создает полную копию класса переменных
    ''' </summary>
    ''' <param name="arrVars">ссылка на хранилище для новой копии</param>
    ''' <param name="arrSrc">Источник, откуда копировать переменные. Если не указан, то из этого класса</param>
    ''' <remarks></remarks>
    Public Sub CopyVariables(ByRef arrVars As SortedList(Of String, variableEditorInfoType), Optional ByRef arrSrc As SortedList(Of String, variableEditorInfoType) = Nothing)
        'arrVars = lstVariables
        If IsNothing(arrSrc) Then arrSrc = lstVariables
        arrVars = New SortedList(Of String, variableEditorInfoType)(StringComparer.CurrentCultureIgnoreCase)
        If IsNothing(arrSrc) OrElse arrSrc.Count = 0 Then Return

        Dim dest As variableEditorInfoType
        For i As Integer = 0 To arrSrc.Count - 1
            Dim src As variableEditorInfoType = arrSrc.ElementAt(i).Value
            dest = New variableEditorInfoType With {.Description = src.Description, .Hidden = src.Hidden}
            ReDim dest.arrValues(src.arrValues.Count - 1)
            Array.ConstrainedCopy(src.arrValues, 0, dest.arrValues, 0, dest.arrValues.Length)
            If IsNothing(src.lstSingatures) = False AndAlso src.lstSingatures.Count > 0 Then
                Dim arrValues() As Integer = src.lstSingatures.Values.ToArray
                Dim arrKeys() As String = src.lstSingatures.Keys.ToArray
                dest.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                For j As Integer = 0 To src.lstSingatures.Count - 1
                    dest.lstSingatures.Add(arrKeys(j), arrValues(j))
                Next j
            End If
            arrVars.Add(arrSrc.ElementAt(i).Key, dest)
        Next i
    End Sub

    ''' <summary>
    ''' Восстанавливает перменные из указанного хранилища после предварительного сохранения методом CopyVariables
    ''' </summary>
    ''' <param name="arrVars">массив для восстановления, полученный функцией CopyVariables</param>
    Public Sub RestoreVariables(ByRef arrVars As SortedList(Of String, variableEditorInfoType), Optional ByVal fullCopy As Boolean = False)

        If fullCopy Then
            lstVariables.Clear()

            Dim dest As variableEditorInfoType
            For i As Integer = 0 To arrVars.Count - 1
                Dim src As variableEditorInfoType = arrVars.ElementAt(i).Value
                dest = New variableEditorInfoType With {.Description = src.Description, .Hidden = src.Hidden}
                ReDim dest.arrValues(src.arrValues.Count - 1)
                Array.ConstrainedCopy(src.arrValues, 0, dest.arrValues, 0, dest.arrValues.Length)
                If IsNothing(src.lstSingatures) = False AndAlso src.lstSingatures.Count > 0 Then
                    Dim arrValues() As Integer = src.lstSingatures.Values.ToArray
                    Dim arrKeys() As String = src.lstSingatures.Keys.ToArray
                    dest.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                    For j As Integer = 0 To src.lstSingatures.Count - 1
                        dest.lstSingatures.Add(arrKeys(j), arrValues(j))
                    Next j
                End If
                lstVariables.Add(arrVars.ElementAt(i).Key, dest)
            Next i
        Else
            lstVariables = arrVars
        End If
    End Sub

    ''' <summary>
    ''' Создает список переменных
    ''' </summary>
    ''' <returns>Список переменных или, если их нет, пустой список (не Nothing)</returns>
    Public Function GetVariablesList(Optional ByVal wrapInQuotes As Boolean = False) As List(Of String)
        If IsNothing(lstVariables) OrElse lstVariables.Count = 0 Then Return New List(Of String) '(StringComparer.CurrentCultureIgnoreCase)
        If wrapInQuotes Then
            Dim arrNames As List(Of String) = lstVariables.Keys.ToList
            For i As Integer = 0 To arrNames.Count - 1
                arrNames(i) = "'" + arrNames(i) + "'"
            Next
            Return arrNames
        Else
            Return lstVariables.Keys.ToList
        End If
    End Function

    ''' <summary>
    ''' Создает полный дубликат переменной
    ''' </summary>
    ''' <param name="varName">Имя дублируемой переменной</param>
    ''' <param name="newName">Имя новой переменной-дубликата</param>
    ''' <returns>-1 при ошибке</returns>
    Public Function DuplicateVariable(ByVal varName As String, ByVal newName As String) As Integer
        Dim varToCopy As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, varToCopy) = False Then Return -1
        If lstVariables.ContainsKey(newName) Then Return -1

        Dim varNew As New variableEditorInfoType
        varNew.Description = varToCopy.Description
        varNew.Group = varToCopy.Group
        varNew.Hidden = varToCopy.Hidden
        varNew.Icon = varToCopy.Icon
        If IsNothing(varToCopy.arrValues) = False AndAlso varToCopy.arrValues.Count > 0 Then
            ReDim varNew.arrValues(varToCopy.arrValues.Count - 1)
            Array.Copy(varToCopy.arrValues, varNew.arrValues, varToCopy.arrValues.Count)
        End If
        If IsNothing(varToCopy.lstSingatures) = False AndAlso varToCopy.lstSingatures.Count > 0 Then
            varNew.lstSingatures = New SortedList(Of String, Integer)
            For i As Integer = 0 To varToCopy.lstSingatures.Count - 1
                varNew.lstSingatures.Add(varToCopy.lstSingatures.ElementAt(i).Key, varToCopy.lstSingatures.ElementAt(i).Value)
            Next
        End If
        lstVariables.Add(newName, varNew)
        Return 1
    End Function

    ''' <summary>
    ''' Создает список из сигнатур переменной-массива
    ''' </summary>
    ''' <param name="varName">имя переменной</param>
    ''' <param name="wrapInQuotes">Заворачивать ли их в кавычки</param>
    ''' <returns>Список сигнатур. Если сигнатур нет, то пустой список. Если переменной нет - то Nothing</returns>
    Public Function CreateListOfVariableSignatures(ByVal varName As String, Optional ByVal wrapInQuotes As Boolean = False) As List(Of String)
        Dim myVar As variableEditorInfoType = Nothing
        If lstVariables.TryGetValue(varName, myVar) = False Then Return Nothing
        If IsNothing(myVar.lstSingatures) OrElse myVar.lstSingatures.Count = 0 Then Return New List(Of String)
        If Not wrapInQuotes Then Return myVar.lstSingatures.Keys.ToList


        Dim arrSign() As String = myVar.lstSingatures.Keys.ToArray
        For i As Integer = 0 To arrSign.Count - 1
            arrSign(i) = "'" + arrSign(i) + "'"
        Next
        Return arrSign.ToList
    End Function

    Public Sub RenameVariablesGroup(ByVal oldName As String, ByVal newName As String)
        If IsNothing(lstVariables) OrElse lstVariables.Count = 0 Then Return
        For i As Integer = 0 To lstVariables.Count - 1
            Dim v As variableEditorInfoType = lstVariables.ElementAt(i).Value
            If String.Compare(v.Group, oldName, True) = 0 Then
                lstVariables(i).Group = newName
            End If
        Next
    End Sub

    ''' <summary>
    ''' Возвращает индекс элемента массива по его сигнатуре (или индексу строкой)
    ''' </summary>
    ''' <param name="var">ссылка на переменную</param>
    ''' <param name="varElement">id строкой или сигнатура (в кавычках или без)</param>
    Public Shared Function GetElementIndex(ByRef var As variableEditorInfoType, ByVal varElement As String) As Integer
        If IsNumeric(varElement) Then
            Dim id As Integer = CInt(varElement)
            If id >= var.arrValues.Count Then Return -1
            Return id
        Else
            If IsNothing(var.lstSingatures) Then Return -1
            varElement = UnWrapString(varElement)
            Dim id As Integer = -1
            If var.lstSingatures.TryGetValue(varElement, id) = False Then Return -1
            Return id
        End If
    End Function
End Class
