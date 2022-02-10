Public Class cTrackingProperties
    Public Class TrackingPropertyData
        Public eventBeforeId As Integer = -1
        Public propBeforeContent() As CodeTextBox.CodeDataType
        Public eventAfterId As Integer = -1
        Public propAfterContent() As CodeTextBox.CodeDataType
    End Class

    ''' <summary>
    ''' Ключ - первое имя класса .Свойство (MyClass.PropName) без кавычек
    ''' </summary>
    Public lstTrackingProperties As New SortedList(Of String, TrackingPropertyData)(StringComparer.CurrentCultureIgnoreCase)

    Public Function ContainsPropertyBefore(ByVal classId As Integer, ByVal propName As String)
        If lstTrackingProperties.Count = 0 Then Return False
        Dim className As String = mScript.mainClass(classId).Names(0)
        Dim pName As String = className & "." & propName
        If lstTrackingProperties.ContainsKey(pName) = False Then Return False
        If lstTrackingProperties(pName).eventBeforeId > 0 Then Return True
        Return False
    End Function

    Public Function ContainsPropertyAfter(ByVal classId As Integer, ByVal propName As String)
        If lstTrackingProperties.Count = 0 Then Return False
        Dim className As String = mScript.mainClass(classId).Names(0)
        Dim pName As String = className & "." & propName
        If lstTrackingProperties.ContainsKey(pName) = False Then Return False
        If lstTrackingProperties(pName).eventAfterId > 0 Then Return True
        Return False
    End Function

    ''' <summary>
    ''' Запускает событие до изменения свойства, и, если оно есть, возвращает его результат. Иначе возвращает пустую строку.
    ''' </summary>
    ''' <param name="classId">Id класса свойства</param>
    ''' <param name="propName">имя свойства</param>
    ''' <param name="arrParams">параметры запуска кода (содержит только Id элементов)</param>
    ''' <param name="newValue"></param>
    Public Function RunBefore(ByVal classId As Integer, ByVal propName As String, ByRef arrParams() As String, ByRef newValue As String) As String
        If lstTrackingProperties.Count = 0 Then Return ""

        Dim pName As String = mScript.mainClass(classId).Names(0) & "." & propName
        Dim pos As Integer = lstTrackingProperties.IndexOfKey(pName)
        If pos = -1 Then Return ""

        Dim eventId = lstTrackingProperties.ElementAt(pos).Value.eventBeforeId
        If eventId < 1 Then Return ""
        Dim propParams() As String
        ReDim propParams(2)
        propParams(2) = newValue

        If IsNothing(arrParams) OrElse arrParams.Count = 0 Then
            propParams(0) = "-1"
            propParams(1) = "-1"
        ElseIf arrParams.Count = 1 Then
            propParams(0) = arrParams(0)
            propParams(1) = "-1"
        Else
            propParams(0) = arrParams(0)
            propParams(1) = arrParams(1)
        End If

        Dim res As String = mScript.eventRouter.RunEvent(eventId, propParams, "Событие до изменения свойства" & propName, False)

        If String.IsNullOrEmpty(res) = False Then newValue = res
        Return res
    End Function

    ''' <summary>
    ''' Запускает событие после изменения свойства, и, если оно есть, возвращает его результат. Иначе возвращает пустую строку.
    ''' </summary>
    ''' <param name="classId">Id класса свойства</param>
    ''' <param name="propName">имя свойства</param>
    ''' <param name="arrParams">параметры запуска кода (содержит только Id элементов)</param>
    Public Function RunAfter(ByVal classId As Integer, ByVal propName As String, ByRef arrParams() As String) As String
        If lstTrackingProperties.Count = 0 Then Return ""

        Dim pName As String = mScript.mainClass(classId).Names(0) & "." & propName
        Dim pos As Integer = lstTrackingProperties.IndexOfKey(pName)
        If pos = -1 Then Return ""

        Dim eventId = lstTrackingProperties.ElementAt(pos).Value.eventAfterId
        If eventId < 1 Then Return ""
        Dim propParams() As String
        ReDim propParams(1)

        If IsNothing(arrParams) OrElse arrParams.Count = 0 Then
            propParams(0) = "-1"
            propParams(1) = "-1"
        ElseIf arrParams.Count = 1 Then
            propParams(0) = arrParams(0)
            propParams(1) = "-1"
        Else
            propParams(0) = arrParams(0)
            propParams(1) = arrParams(1)
        End If

        Dim res As String = mScript.eventRouter.RunEvent(eventId, propParams, "Событие после изменения свойства" & propName, False)
        Return res
    End Function

    Dim sep2 As String = "|#" & Chr(2) & "|"
    Dim sep3 As String = "|#" & Chr(3) & "|"
    Public Function CreateStringForSaveFile() As String
        If lstTrackingProperties.Count = 0 Then Return ""

        Dim txt As New System.Text.StringBuilder
        For i As Integer = 0 To lstTrackingProperties.Count - 1
            Dim val As TrackingPropertyData = lstTrackingProperties.ElementAt(i).Value
            txt.Append(lstTrackingProperties.ElementAt(i).Key & sep3 & val.eventBeforeId.ToString & sep3 & questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(val.propBeforeContent) & _
                       sep3 & val.eventAfterId.ToString & sep3 & questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(val.propAfterContent))
            If i < lstTrackingProperties.Count - 1 Then txt.Append(sep2)
        Next

        Return txt.ToString
    End Function

    Public Sub LoadDataFromSaveString(ByVal strData As String)
        strData = strData.Trim
        If String.IsNullOrEmpty(strData) Then Return

        Dim arr() As String = Split(strData, sep2)
        '0 - key, 1 - eventId, 2 - codedata
        'For i As Integer = 0 To arr.Count - 1
        '    Dim arr2() As String = Split(arr(i), sep3)
        '    Dim pName As String = arr2(0)
        '    Dim eventId As Integer = CInt(arr2(1))
        '    Dim cd() As CodeTextBox.CodeDataType = questEnvironment.codeBoxShadowed.codeBox.DeserializeCodeData(arr2(2))
        '    lstTrackingProperties.Add(pName, New TrackingPropertyData With {.eventBeforeId = eventId, .propBeforeContent = cd})
        'Next i

        For i As Integer = 0 To arr.Count - 1
            Dim arr2() As String = Split(arr(i), sep3)
            Dim pName As String = arr2(0)
            Dim eventBeforeId As Integer = CInt(arr2(1))
            Dim cd1() As CodeTextBox.CodeDataType = questEnvironment.codeBoxShadowed.codeBox.DeserializeCodeData(arr2(2))
            Dim eventAfterId As Integer = CInt(arr2(3))
            Dim cd2() As CodeTextBox.CodeDataType = questEnvironment.codeBoxShadowed.codeBox.DeserializeCodeData(arr2(4))
            lstTrackingProperties.Add(pName, New TrackingPropertyData With {.eventBeforeId = eventBeforeId, .propBeforeContent = cd1, .eventAfterId = eventAfterId, .propAfterContent = cd2})
        Next i

    End Sub

    Public Sub Clear()
        lstTrackingProperties.Clear()
    End Sub

    ''' <summary>
    ''' Полностью подготавливает кодбокс с событием редактирования свойства
    ''' </summary>
    ''' <param name="relatedControl">Контрол, ассоциированный с данным свойством</param>
    Public Sub LoadEventBeforeToCodeBox(ByRef relatedControl As Object)
        Dim propName As String = relatedControl.Name 'Имя свойства 
        Dim ch As clsPanelManager.clsChildPanel = relatedControl.childPanel

        Dim className As String = mScript.mainClass(ch.classId).Names(0)
        Dim pName As String = className & "." & propName

        With frmMainEditor.codeBox
            .Tag = Nothing
            .Text = ""
            .codeBox.IsTextBlockByDefault = False
            frmMainEditor.trakingEventState = frmMainEditor.trackingcodeEnum.EVENT_BEFORE

            Dim itm As cTrackingProperties.TrackingPropertyData = Nothing
            If lstTrackingProperties.TryGetValue(pName, itm) Then .codeBox.LoadCodeFromCodeData(itm.propBeforeContent)
            .Tag = relatedControl
        End With

        frmMainEditor.codeBoxPanel.Show()
        frmMainEditor.WBhelp.Hide()
    End Sub

    Public Sub LoadEventAfterToCodeBox(ByRef relatedControl As Object)
        Dim propName As String = relatedControl.Name 'Имя свойства 
        Dim ch As clsPanelManager.clsChildPanel = relatedControl.childPanel

        Dim className As String = mScript.mainClass(ch.classId).Names(0)
        Dim pName As String = className & "." & propName

        With frmMainEditor.codeBox
            .Tag = Nothing
            .Text = ""
            .codeBox.IsTextBlockByDefault = False
            frmMainEditor.trakingEventState = frmMainEditor.trackingcodeEnum.EVENT_AFTER

            Dim itm As cTrackingProperties.TrackingPropertyData = Nothing
            If lstTrackingProperties.TryGetValue(pName, itm) Then .codeBox.LoadCodeFromCodeData(itm.propAfterContent)
            .Tag = relatedControl
        End With

        frmMainEditor.codeBoxPanel.Show()
        frmMainEditor.WBhelp.Hide()
    End Sub


    ''' <summary>
    ''' Добавляет новое свойство для отслеживания. Если свойство уже есть, то заменяет его новыми данными
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    ''' <param name="propName">Имя свойства без кавычек</param>
    ''' <param name="content">данные из кодбокса</param>
    Public Sub AddPropertyBefore(ByVal classId As Integer, ByVal propName As String, ByRef content() As CodeTextBox.CodeDataType)
        Dim className As String = mScript.mainClass(classId).Names(0)
        Dim isEmpty As Boolean = False
        If IsNothing(content) OrElse content.Count = 0 OrElse (content.Count = 1 AndAlso IsNothing(content(0).Code)) Then isEmpty = True

        'создаем новое или редактируем старое событие
        Dim pName As String = className & "." & propName
        Dim itm As TrackingPropertyData = Nothing
        If lstTrackingProperties.TryGetValue(pName, itm) = False Then
            'создаем
            If isEmpty Then Return
            itm = New TrackingPropertyData
            lstTrackingProperties.Add(pName, itm)
        End If

        'редактируем
        If isEmpty Then
            'кода нет - удаляем 
            mScript.eventRouter.RemoveEvent(itm.eventBeforeId)
            If lstTrackingProperties(pName).eventAfterId < 1 Then
                lstTrackingProperties.Remove(pName)
                itm = Nothing
            End If
        Else
            'код есть - редактируем
            itm.eventBeforeId = mScript.eventRouter.SetEventId(itm.eventBeforeId, content)
            itm.propBeforeContent = CopyCodeDataArray(content)
        End If
    End Sub

    ''' <summary>
    ''' Добавляет новое свойство для отслеживания. Если свойство уже есть, то заменяет его новыми данными
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    ''' <param name="propName">Имя свойства без кавычек</param>
    ''' <param name="content">данные из кодбокса</param>
    Public Sub AddPropertyAfter(ByVal classId As Integer, ByVal propName As String, ByRef content() As CodeTextBox.CodeDataType)
        Dim className As String = mScript.mainClass(classId).Names(0)
        Dim isEmpty As Boolean = False
        If IsNothing(content) OrElse content.Count = 0 OrElse (content.Count = 1 AndAlso IsNothing(content(0).Code)) Then isEmpty = True

        'создаем новое или редактируем старое событие
        Dim pName As String = className & "." & propName
        Dim itm As TrackingPropertyData = Nothing
        If lstTrackingProperties.TryGetValue(pName, itm) = False Then
            'создаем
            If isEmpty Then Return
            itm = New TrackingPropertyData
            lstTrackingProperties.Add(pName, itm)
        End If

        'редактируем
        If isEmpty Then
            'кода нет - удаляем 
            mScript.eventRouter.RemoveEvent(itm.eventAfterId)
            If lstTrackingProperties(pName).eventBeforeId < 1 Then
                lstTrackingProperties.Remove(pName)
                itm = Nothing
            End If
        Else
            'код есть - редактируем
            itm.eventAfterId = mScript.eventRouter.SetEventId(itm.eventAfterId, content)
            itm.propAfterContent = CopyCodeDataArray(content)
        End If
    End Sub

    Public Sub RenameProperty(ByVal classId As Integer, ByVal oldName As String, ByVal newName As String)
        If lstTrackingProperties.Count = 0 Then Return
        Dim className As String = mScript.mainClass(classId).Names(0)

        Dim pOldName As String = className & "." & oldName
        Dim pos As Integer = lstTrackingProperties.IndexOfKey(pOldName)
        If pos = -1 Then Return
        Dim itm As TrackingPropertyData = lstTrackingProperties.ElementAt(pos).Value

        Dim pNewName As String = className & "." & newName
        lstTrackingProperties.RemoveAt(pos)
        lstTrackingProperties.Add(pNewName, itm)
    End Sub

    ''' <summary>
    ''' Удаляет свойство из отслеживаемых
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    ''' <param name="propName">Имя свойства без кавычек</param>
    Public Sub RemoveProperty(ByVal classId As Integer, ByVal propName As String, Optional tracking As frmMainEditor.trackingcodeEnum = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT)
        If lstTrackingProperties.Count = 0 Then Return
        Dim className As String = mScript.mainClass(classId).Names(0)


        Dim pName As String = className & "." & propName
        Dim pos As Integer = lstTrackingProperties.IndexOfKey(pName)
        If pos > -1 Then
            If tracking = tracking OrElse (tracking = frmMainEditor.trackingcodeEnum.EVENT_AFTER AndAlso lstTrackingProperties(pName).eventBeforeId < 1) OrElse _
                 (tracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE AndAlso lstTrackingProperties(pName).eventAfterId < 1) Then
                lstTrackingProperties(pName) = Nothing
                lstTrackingProperties.RemoveAt(pos)
            ElseIf tracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE Then
                lstTrackingProperties(pName).eventBeforeId = -1
                lstTrackingProperties(pName).propBeforeContent = Nothing
            ElseIf tracking = frmMainEditor.trackingcodeEnum.EVENT_AFTER Then
                lstTrackingProperties(pName).eventAfterId = -1
                lstTrackingProperties(pName).propAfterContent = Nothing
            End If
        End If
    End Sub

    Public Sub RenameClass(ByVal oldClassName As String, ByVal newClassName As String)
        If lstTrackingProperties.Count = 0 Then Return
        oldClassName = FirstClassName(oldClassName)
        If oldClassName.Length = 0 Then
            MessageBox.Show("неправильное имя класса", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

again:
        For i As Integer = 0 To lstTrackingProperties.Count - 1
            Dim cKey As String = lstTrackingProperties.ElementAt(i).Key
            If cKey.StartsWith(oldClassName & ".") Then
                Dim itm As TrackingPropertyData = lstTrackingProperties.ElementAt(i).Value
                Dim pNewName As String = newClassName & "." & cKey.Substring(oldClassName.Length + 1)
                lstTrackingProperties.RemoveAt(i)
                lstTrackingProperties.Add(pNewName, itm)
                GoTo again 'список мог пересортироваться и теперь невозможно свести концы с концами - просто поиск новых совпадений для переименования класа начинаем с самого начала
            End If
        Next i
    End Sub

    Public Sub RemoveClass(ByVal className As String)
        If lstTrackingProperties.Count = 0 Then Return
        className = FirstClassName(className)
        If className.Length = 0 Then
            MessageBox.Show("неправильное имя класса", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        For i As Integer = lstTrackingProperties.Count - 1 To 0 Step -1
            Dim cKey As String = lstTrackingProperties.ElementAt(i).Key
            If cKey.StartsWith(className & ".") Then
                lstTrackingProperties(cKey) = Nothing
                lstTrackingProperties.RemoveAt(i)
            End If
        Next i
    End Sub

    Private Function FirstClassName(ByVal anyClassName As String) As String
        Dim classId As Integer = 0
        If mScript.mainClassHash.TryGetValue(anyClassName, classId) = False Then Return ""
        Return mScript.mainClass(classId).Names(0)
    End Function
End Class
