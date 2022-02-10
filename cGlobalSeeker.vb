Public Class cGlobalSeeker
#Region "Code Navigation And Info Functions"
    ''' <summary>
    ''' Процедура находит слово функции или свойства, в которой находится каретка, а также номер вводимого параметра
    ''' </summary>
    ''' <param name="currentLine">текущая линия</param>
    ''' <param name="currentWordId">Id текущего слова в линии</param>
    ''' <param name="paramNumber">для получения номера параметра, внутри которого каретка</param>
    ''' <param name="elementWord">для получения слова функции/свойства</param>
    ''' <param name="classId">для получения Id класса функции/свойства</param>
    ''' <param name="elementWordId">Id слова в строке, которая является функцией/свойством</param>
    ''' <returns></returns>
    Public Function GetCurrentElementWord(ByRef CodeData() As CodeTextBox.CodeDataType, ByVal currentLine As Integer, ByVal currentWordId As Integer, ByRef paramNumber As Integer, _
                                           ByRef elementWord As CodeTextBox.EditWordType, ByRef classId As Integer, Optional ByRef elementWordId As Integer = -1) As Boolean
        If IsNothing(mScript.mainClass) Then Return False
        'class[...,X,...].funcName...
        '...funcName(...,X,...

        If currentWordId < 0 Then Return False
        If IsNothing(CodeData(currentLine).Code) Then Return False
        If CodeData(currentLine).Code.GetUpperBound(0) < currentWordId Then Return False
        elementWord = New CodeTextBox.EditWordType 'очищаем слово
        Dim quadBracketBalance As Integer = 0 'баланс [ ]
        Dim ovalBracketBalance As Integer = 0 'баланс 
        Dim paramIfInQuad As Integer = 1 'номер параметра, если мы внутри []
        Dim paramIfInOval As Integer = 1 'номер параметра, если мы внутри () или скобок нет вообще
        classId = -1 'для получения класса элемента (т. е. функции или свойства)
        'Dim wordId As Integer = -1 'Id слова с элементом
        'сначала проходим от текущего слова до первого, определяя внутри каких скобок мы находимся - () или []
        For i As Integer = currentWordId To 0 Step -1
            Select Case CodeData(currentLine).Code(i).wordType
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN '[
                    quadBracketBalance += 1
                    If quadBracketBalance = 1 Then
                        'мы внутри [..X..]
                        If i = 0 Then Return False 'перед [ ничего не стоит - выход
                        classId = CodeData(currentLine).Code(i - 1).classId
                        paramNumber = paramIfInQuad 'получаем номер параметра, внутри которого мы находимся)
                        If classId = -1 Then Return False
                        Exit For
                    End If
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE ']
                    quadBracketBalance -= 1
                Case CodeTextBox.EditWordTypeEnum.W_FUNCTION
                    If CodeData(currentLine).Code.GetUpperBound(0) = currentWordId OrElse (CodeData(currentLine).Code(i + 1).wordType <> CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN AndAlso _
                                                                                           quadBracketBalance = 0) Then
                        'мы внутри параметров функции (за ее именем), у которой нет круглых скобок
                        If i > 0 Then
                            If CodeData(currentLine).Code(i - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_POINT Then
                                If CodeData(currentLine).Code(i - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(i - 1).Word.Trim <> "Then" Then
                                    Continue For
                                End If
                            ElseIf (CodeData(currentLine).Code(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER OrElse CodeData(currentLine).Code(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING) _
                                AndAlso i > 1 AndAlso CodeData(currentLine).Code(i - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_POINT Then

                                If CodeData(currentLine).Code(i - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(i - 1).Word.Trim <> "Then" Then
                                    Continue For
                                End If
                            ElseIf i > 1 AndAlso CodeData(currentLine).Code(i - 2).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                                'If i = 2 Then Continue For
                                'If CodeData(currentLine).Code(i - 3).wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(i - 3).Word.Trim <> "Then" Then
                                '    Continue For
                                'End If

                                If i > 2 Then
                                    If CodeData(currentLine).Code(i - 3).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(i - 3).Word.Trim <> "Then" Then
                                        Continue For
                                    End If
                                End If
                            ElseIf i > 1 AndAlso CodeData(currentLine).Code(i - 2).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                'Class[...].Name
                                Dim internalQbalance As Integer = -1
                                Dim classPos As Integer = -1
                                For j = i - 3 To 0 Step -1
                                    If CodeData(currentLine).Code(j).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                                        internalQbalance += 1
                                        If internalQbalance = 0 Then
                                            classPos = j - 1
                                        End If
                                    ElseIf CodeData(currentLine).Code(j).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                        internalQbalance -= 1
                                    End If
                                Next
                                If classPos > 0 AndAlso CodeData(currentLine).Code(classPos).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                                    If CodeData(currentLine).Code(classPos - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(classPos - 1).Word.Trim <> "Then" Then
                                        Continue For
                                    End If
                                End If
                            End If
                        End If

                        'мы за функцией, не имеющей круглых скобок
                        classId = CodeData(currentLine).Code(i).classId
                        paramNumber = paramIfInOval 'получаем номер параметра, без учета возможных параметров в квадратных скобках
                        If classId = -1 Then Return False 'это переменная - выход (на всякий случай, хоть это и синтаксическая ошибка)
                        If CodeData(currentLine).Code(i).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY OrElse CodeData(currentLine).Code(i).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                            elementWord = CodeData(currentLine).Code(i) 'получаем слово элемента и его Id
                            elementWordId = i
                        Else
                            Return False
                        End If
                        Exit For
                    End If
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN '(
                    ovalBracketBalance += 1
                    If ovalBracketBalance = 1 Then
                        'мы внутри (..Х..)
                        If i = 0 Then Return False
                        classId = CodeData(currentLine).Code(i - 1).classId
                        paramNumber = paramIfInOval 'получаем номер параметра, без учета возможных параметров в квадратных скобках
                        If classId = -1 Then Return False 'переменная - выход
                        If CodeData(currentLine).Code(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY OrElse CodeData(currentLine).Code(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                            elementWord = CodeData(currentLine).Code(i - 1) 'получаем слово элемента и его Id
                            elementWordId = i - 1
                        Else
                            Return False
                        End If
                        Exit For
                    End If
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE ')
                    ovalBracketBalance -= 1
                Case CodeTextBox.EditWordTypeEnum.W_COMMA
                    'пока мы не знаем, внутри каких скобок находимся, считаем и как для круглых, и как для квадратных
                    If quadBracketBalance = 0 Then paramIfInQuad += 1
                    If ovalBracketBalance = 0 Then paramIfInOval += 1
            End Select
        Next

        'Сейчас возможны 2 варианта:
        '1) Мы в [..X..], известен класс, известен номер параметра, не известно слово
        '2) Мы в (..X..), известен класс, не известен номер параметра, известно слово

        If elementWordId = -1 Then
            'вариант 1). Ищем закрывающую квадратную скобку функции, вслед за которой должно идти .funcName
            quadBracketBalance = 0
            For i As Integer = currentWordId To CodeData(currentLine).Code.GetUpperBound(0)
                Select Case CodeData(currentLine).Code(i).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        quadBracketBalance += 1
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        quadBracketBalance -= 1
                        If quadBracketBalance = -1 Then
                            'наша закрывающая ] найдена
                            If i + 1 < CodeData(currentLine).Code.GetUpperBound(0) Then
                                If CodeData(currentLine).Code(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY OrElse CodeData(currentLine).Code(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                                    'элемент найден. Получаем слово
                                    elementWord = CodeData(currentLine).Code(i + 2)
                                    elementWordId = i + 2
                                    Return True
                                Else
                                    Return False
                                End If
                            End If
                        End If
                End Select
            Next
        Else
            'вариант 2). Не известно кол-во параметров в квадратных скобках
            If elementWordId = 0 Then Return True 'слово первое, значит квадратных скобок нет - все известно
            If CodeData(currentLine).Code(elementWordId - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_POINT Then Return True 'перед именем функции нет точки - квадратных скобок нет - все известно
            If elementWordId < 3 Then Return True 'ошибка синтаксиса
            If CodeData(currentLine).Code(elementWordId - 2).wordType <> CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then Return True 'квадратных скобок нет - все известно
            quadBracketBalance = -1
            paramNumber += 1
            'проверяем все слова от слова перед закрывающей ] до начала строки в поисках открывающей [ функции
            For i = elementWordId - 3 To 0 Step -1
                Select Case CodeData(currentLine).Code(i).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        quadBracketBalance += 1
                        If quadBracketBalance = 0 Then Return True 'найдена нужная [ - выход
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        quadBracketBalance -= 1
                    Case CodeTextBox.EditWordTypeEnum.W_COMMA
                        If quadBracketBalance = -1 Then paramNumber += 1 'найдена запятая этой функции, номер параметра +1
                End Select
            Next
        End If
        Return False
    End Function

    ''' <summary>
    ''' Процедура находит слово функции или свойства, указанного на месте currentWordId, а также номер вводимого параметра
    ''' </summary>
    ''' <param name="currentLine">текущая линия</param>
    ''' <param name="currentWordId">Id текущего слова в линии</param>
    ''' <param name="paramNumber">для получения номера параметра, внутри которого каретка</param>
    ''' <param name="elementWord">для получения слова функции/свойства</param>
    ''' <param name="classId">для получения Id класса функции/свойства</param>
    ''' <param name="elementWordId">Id слова в строке, которая является функцией/свойством</param>
    ''' <returns></returns>
    Public Function GetCurrentElementWordInExData(ByRef CodeData As List(Of MatewScript.ExecuteDataType), ByVal currentLine As Integer, ByVal currentWordId As Integer, ByRef paramNumber As Integer, _
                                           ByRef elementWord As CodeTextBox.EditWordType, ByRef classId As Integer, Optional ByRef elementWordId As Integer = -1) As Boolean
        If IsNothing(mScript.mainClass) Then Return False
        'class[...,X,...].funcName...
        '...funcName(...,X,...

        If currentWordId < 0 Then Return False
        If IsNothing(CodeData(currentLine).Code) Then Return False
        If CodeData(currentLine).Code.GetUpperBound(0) < currentWordId Then Return False
        elementWord = New CodeTextBox.EditWordType 'очищаем слово
        Dim quadBracketBalance As Integer = 0 'баланс [ ]
        Dim ovalBracketBalance As Integer = 0 'баланс 
        Dim paramIfInQuad As Integer = 1 'номер параметра, если мы внутри []
        Dim paramIfInOval As Integer = 1 'номер параметра, если мы внутри () или скобок нет вообще
        classId = -1 'для получения класса элемента (т. е. функции или свойства)
        'Dim wordId As Integer = -1 'Id слова с элементом
        'сначала проходим от текущего слова до первого, определяя внутри каких скобок мы находимся - () или []
        For i As Integer = currentWordId To 0 Step -1
            Select Case CodeData(currentLine).Code(i).wordType
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN '[
                    quadBracketBalance += 1
                    If quadBracketBalance = 1 Then
                        'мы внутри [..X..]
                        If i = 0 Then Return False 'перед [ ничего не стоит - выход
                        classId = CodeData(currentLine).Code(i - 1).classId
                        paramNumber = paramIfInQuad 'получаем номер параметра, внутри которого мы находимся)
                        If classId = -1 Then Return False
                        Exit For
                    End If
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE ']
                    quadBracketBalance -= 1
                Case CodeTextBox.EditWordTypeEnum.W_FUNCTION
                    If CodeData(currentLine).Code.GetUpperBound(0) = currentWordId OrElse (CodeData(currentLine).Code(i + 1).wordType <> CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN AndAlso _
                                                                                           quadBracketBalance = 0) Then
                        'мы внутри параметров функции (за ее именем), у которой нет круглых скобок
                        If i > 0 Then
                            If CodeData(currentLine).Code(i - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_POINT Then
                                If CodeData(currentLine).Code(i - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(i - 1).Word.Trim <> "Then" Then
                                    Continue For
                                End If
                            ElseIf (CodeData(currentLine).Code(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER OrElse CodeData(currentLine).Code(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING) _
                                AndAlso i > 1 AndAlso CodeData(currentLine).Code(i - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_POINT Then

                                If CodeData(currentLine).Code(i - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(i - 1).Word.Trim <> "Then" Then
                                    Continue For
                                End If
                            ElseIf i > 1 AndAlso CodeData(currentLine).Code(i - 2).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                                'If i = 2 Then Continue For
                                'If CodeData(currentLine).Code(i - 3).wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(i - 3).Word.Trim <> "Then" Then
                                '    Continue For
                                'End If

                                If i > 2 Then
                                    If CodeData(currentLine).Code(i - 3).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(i - 3).Word.Trim <> "Then" Then
                                        Continue For
                                    End If
                                End If
                            ElseIf i > 1 AndAlso CodeData(currentLine).Code(i - 2).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                'Class[...].Name
                                Dim internalQbalance As Integer = -1
                                Dim classPos As Integer = -1
                                For j = i - 3 To 0 Step -1
                                    If CodeData(currentLine).Code(j).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                                        internalQbalance += 1
                                        If internalQbalance = 0 Then
                                            classPos = j - 1
                                        End If
                                    ElseIf CodeData(currentLine).Code(j).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                        internalQbalance -= 1
                                    End If
                                Next
                                If classPos > 0 AndAlso CodeData(currentLine).Code(classPos).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                                    If CodeData(currentLine).Code(classPos - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso CodeData(currentLine).Code(classPos - 1).Word.Trim <> "Then" Then
                                        Continue For
                                    End If
                                End If
                            End If
                        End If

                        'мы за функцией, не имеющей круглых скобок
                        classId = CodeData(currentLine).Code(i).classId
                        paramNumber = paramIfInOval 'получаем номер параметра, без учета возможных параметров в квадратных скобках
                        If classId = -1 Then Return False 'это переменная - выход (на всякий случай, хоть это и синтаксическая ошибка)
                        If CodeData(currentLine).Code(i).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY OrElse CodeData(currentLine).Code(i).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                            elementWord = CodeData(currentLine).Code(i) 'получаем слово элемента и его Id
                            elementWordId = i
                        Else
                            Return False
                        End If
                        Exit For
                    End If
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN '(
                    ovalBracketBalance += 1
                    If ovalBracketBalance = 1 Then
                        'мы внутри (..Х..)
                        If i = 0 Then Return False
                        classId = CodeData(currentLine).Code(i - 1).classId
                        paramNumber = paramIfInOval 'получаем номер параметра, без учета возможных параметров в квадратных скобках
                        If classId = -1 Then Return False 'переменная - выход
                        If CodeData(currentLine).Code(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY OrElse CodeData(currentLine).Code(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                            elementWord = CodeData(currentLine).Code(i - 1) 'получаем слово элемента и его Id
                            elementWordId = i - 1
                        Else
                            Return False
                        End If
                        Exit For
                    End If
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE ')
                    ovalBracketBalance -= 1
                Case CodeTextBox.EditWordTypeEnum.W_COMMA
                    'пока мы не знаем, внутри каких скобок находимся, считаем и как для круглых, и как для квадратных
                    If quadBracketBalance = 0 Then paramIfInQuad += 1
                    If ovalBracketBalance = 0 Then paramIfInOval += 1
            End Select
        Next

        'Сейчас возможны 2 варианта:
        '1) Мы в [..X..], известен класс, известен номер параметра, не известно слово
        '2) Мы в (..X..), известен класс, не известен номер параметра, известно слово

        If elementWordId = -1 Then
            'вариант 1). Ищем закрывающую квадратную скобку функции, вслед за которой должно идти .funcName
            quadBracketBalance = 0
            For i As Integer = currentWordId To CodeData(currentLine).Code.GetUpperBound(0)
                Select Case CodeData(currentLine).Code(i).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        quadBracketBalance += 1
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        quadBracketBalance -= 1
                        If quadBracketBalance = -1 Then
                            'наша закрывающая ] найдена
                            If i + 1 < CodeData(currentLine).Code.GetUpperBound(0) Then
                                If CodeData(currentLine).Code(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY OrElse CodeData(currentLine).Code(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                                    'элемент найден. Получаем слово
                                    elementWord = CodeData(currentLine).Code(i + 2)
                                    elementWordId = i + 2
                                    Return True
                                Else
                                    Return False
                                End If
                            End If
                        End If
                End Select
            Next
        Else
            'вариант 2). Не известно кол-во параметров в квадратных скобках
            If elementWordId = 0 Then Return True 'слово первое, значит квадратных скобок нет - все известно
            If CodeData(currentLine).Code(elementWordId - 1).wordType <> CodeTextBox.EditWordTypeEnum.W_POINT Then Return True 'перед именем функции нет точки - квадратных скобок нет - все известно
            If elementWordId < 3 Then Return True 'ошибка синтаксиса
            If CodeData(currentLine).Code(elementWordId - 2).wordType <> CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then Return True 'квадратных скобок нет - все известно
            quadBracketBalance = -1
            paramNumber += 1
            'проверяем все слова от слова перед закрывающей ] до начала строки в поисках открывающей [ функции
            For i = elementWordId - 3 To 0 Step -1
                Select Case CodeData(currentLine).Code(i).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        quadBracketBalance += 1
                        If quadBracketBalance = 0 Then Return True 'найдена нужная [ - выход
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        quadBracketBalance -= 1
                    Case CodeTextBox.EditWordTypeEnum.W_COMMA
                        If quadBracketBalance = -1 Then paramNumber += 1 'найдена запятая этой функции, номер параметра +1
                End Select
            Next
        End If
        Return False
    End Function

    ''' <summary>
    ''' Получает параметр функции, типом которого является PARAM_ELEMENT. При этом определяется его индекс в массиве params и возвращается его содержимое, но только в том случае, если
    ''' оно является простой строкой или числом (то есть, именем или индексом элемента)
    ''' </summary>
    ''' <param name="currentLine">Текущая линия в кодбоксе</param>
    ''' <param name="funcWordId">Id слова в codeBox.CodeData текущей строки, которая является именем функции, с параметрами которого работаем</param>
    ''' <param name="params">Массив параметров данной функции</param>
    ''' <param name="elementParam">Ссылка для получения индекса параметра типа PARAM_ELEMENT</param>
    ''' <returns>Содержимое параметра, но только если оно является простой строкой или числом (то есть, именем или индексом элемента). Иначе - пустая строка</returns>
    Private Function GetFunctionParamWhichTypeIsElement(ByRef cd() As CodeTextBox.CodeDataType, ByVal currentLine As Integer, ByVal funcWordId As Integer, ByRef params() As MatewScript.paramsType, _
                                               ByRef elementParam As MatewScript.paramsType) As String
        'получаем индекс параметра с типом PARAM_ELEMENT
        elementParam = Nothing
        Dim elementParamId As Integer = -1
        For i As Integer = 0 To params.Count - 1
            If params(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                elementParam = params(i)
                elementParamId = i
                Exit For
            End If
        Next
        If IsNothing(elementParam) Then Return ""

        Dim curCD As CodeTextBox.CodeDataType = cd(currentLine)
        Dim parCount As Integer = 0 'хранит количество полученных параметров
        Dim qbCount As Integer = 0, obCount As Integer = 0 'для хранения разности открытых/закрытых скобок
        Dim lstContent As New List(Of String) 'В этот массив получаем список всех параметров функции
        If funcWordId > 4 AndAlso curCD.Code(funcWordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_POINT AndAlso curCD.Code(funcWordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
            'Имеются начальные скобки []. Возможно Class[param1, ...].FuncName
            If curCD.Code(funcWordId - 3).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(funcWordId - 3).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                (funcWordId >= 4 AndAlso (curCD.Code(funcWordId - 4).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA OrElse curCD.Code(funcWordId - 4).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN)) Then
                'слово перед ] является простой строкой или числом, при этом перед этим словом стоит , или [
                lstContent.Add(curCD.Code(funcWordId - 3).Word.Trim)
            Else
                'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                lstContent.Add("")
            End If
            parCount += 1
            'цикл идет от ] в сторону начала строки и начинается внутри [] функции
            qbCount = 1 '(то есть, первый символ ] уже получен)
            For wId As Integer = funcWordId - 3 To 1 Step -1
                Select Case curCD.Code(wId).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount -= 1
                        If qbCount = 0 Then Exit For 'дошли до открывающей [
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount += 1
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        obCount += 1
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                        obCount -= 1
                    Case CodeTextBox.EditWordTypeEnum.W_COMMA
                        'найдена запятая
                        If obCount <> 0 OrElse qbCount <> 1 Then Continue For 'она стоит вне данной функции (является разделителем внутри вложенной функции или свойства)
                        'данная запятая - разделитель параметров этой функции. Получаем параметр сразу перед запятой
                        parCount += 1
                        If curCD.Code(wId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(wId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                            (wId >= 2 AndAlso (curCD.Code(wId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA OrElse curCD.Code(wId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN)) Then
                            'слово перед запятой является простой строкой или числом, при этом перед этим словом стоит , или [
                            lstContent.Insert(0, curCD.Code(wId - 1).Word.Trim) 'вставляем вновь полученный параметр на первое место в lstContent, так как идем от конца к началу
                        Else
                            'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                            lstContent.Insert(0, "")
                        End If
                End Select
            Next wId

            If elementParamId <= lstContent.Count - 1 Then
                'Если среди уже отобранных параметров есть параметр типа PARAM_ELEMENT, то возвращаем его (идти в круглые скобки функции/свойства уже не надо)
                Return lstContent(elementParamId)
            End If
        End If

        Dim codeLen As Integer = curCD.Code.Count
        If codeLen >= funcWordId + 3 AndAlso curCD.Code(funcWordId + 1).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN Then
            '...funcName(paramX, ...
            obCount = 1 'то есть, первый символ ( уже получен
            qbCount = 0
            If curCD.Code(funcWordId + 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(funcWordId + 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                (funcWordId <= codeLen - 4 AndAlso (curCD.Code(funcWordId + 3).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA OrElse curCD.Code(funcWordId + 3).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE)) Then
                'слово после ( является простой строкой или числом, при этом после этого слова стоит , или )
                lstContent.Add(curCD.Code(funcWordId + 2).Word.Trim)
            Else
                'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                lstContent.Add("")
            End If
            parCount += 1
            'цикл идет от ( в сторону конца строки и начинается внутри () функции
            For wId As Integer = funcWordId + 2 To codeLen - 1
                Select Case curCD.Code(wId).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount += 1
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount -= 1
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        obCount -= 1
                        If obCount = 0 Then Exit For 'дошли до закрывающей (
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                        obCount += 1
                    Case CodeTextBox.EditWordTypeEnum.W_COMMA
                        'найдена запятая
                        If obCount <> 0 OrElse qbCount <> 1 Then Continue For 'она стоит вне данной функции (является разделителем внутри вложенной функции или свойства)
                        'данная запятая - разделитель параметров этой функции. Получаем параметр сразу перед запятой

                        If wId = codeLen - 1 Then Exit For
                        parCount += 1
                        If curCD.Code(wId + 1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(wId + 1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                            (wId <= codeLen - 3 AndAlso (curCD.Code(wId + 2).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA OrElse curCD.Code(wId + 2).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE)) Then
                            'слово после запятой является простой строкой или числом, при этом после этого слова стоит , или )
                            lstContent.Add(curCD.Code(wId - 1).Word.Trim)
                        Else
                            'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                            lstContent.Add("")
                        End If
                End Select
            Next wId
            If elementParamId <= lstContent.Count - 1 Then
                'На этом этапе получены в списке lstContent содержатся все параметры фнкции/класса (в правильном порядке). Если содержимое параметра - простое число/строка, то оно попало 
                'в данный список lstContent. Если же содержимое иное (не просчитывается на этапе написания кода), то вместо содержимого - пустая строка.
                Return lstContent(elementParamId) 'возвращаем параметр типа PARAM_ELEMENT
            End If
        End If
        Return "" 'Сюда, по идее, функция дойти не должна
    End Function

    ''' <summary>
    ''' Получает параметр функции, типом которого является PARAM_ELEMENT. При этом определяется его индекс в массиве params и возвращается его содержимое, но только в том случае, если
    ''' оно является простой строкой или числом (то есть, именем или индексом элемента)
    ''' </summary>
    ''' <param name="currentLine">Текущая линия в кодбоксе</param>
    ''' <param name="funcWordId">Id слова в ExecuteDataType текущей строки, которая является именем функции, с параметрами которого работаем</param>
    ''' <param name="params">Массив параметров данной функции</param>
    ''' <param name="elementParam">Ссылка для получения индекса параметра типа PARAM_ELEMENT</param>
    ''' <returns>Содержимое параметра, но только если оно является простой строкой или числом (то есть, именем или индексом элемента). Иначе - пустая строка</returns>
    Private Function GetFunctionParamWhichTypeIsElementEx(ByRef exData As List(Of MatewScript.ExecuteDataType), ByVal currentLine As Integer, ByVal funcWordId As Integer, ByRef params() As MatewScript.paramsType, _
                                               ByRef elementParam As MatewScript.paramsType) As String
        'получаем индекс параметра с типом PARAM_ELEMENT
        elementParam = Nothing
        Dim elementParamId As Integer = -1
        For i As Integer = 0 To params.Count - 1
            If params(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                elementParam = params(i)
                elementParamId = i
                Exit For
            End If
        Next
        If IsNothing(elementParam) Then Return ""

        Dim curCD As MatewScript.ExecuteDataType = exData(currentLine)
        Dim parCount As Integer = 0 'хранит количество полученных параметров
        Dim qbCount As Integer = 0, obCount As Integer = 0 'для хранения разности открытых/закрытых скобок
        Dim lstContent As New List(Of String) 'В этот массив получаем список всех параметров функции
        If funcWordId > 4 AndAlso curCD.Code(funcWordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_POINT AndAlso curCD.Code(funcWordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
            'Имеются начальные скобки []. Возможно Class[param1, ...].FuncName
            If curCD.Code(funcWordId - 3).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(funcWordId - 3).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                (funcWordId >= 4 AndAlso (curCD.Code(funcWordId - 4).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA OrElse curCD.Code(funcWordId - 4).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN)) Then
                'слово перед ] является простой строкой или числом, при этом перед этим словом стоит , или [
                lstContent.Add(curCD.Code(funcWordId - 3).Word.Trim)
            Else
                'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                lstContent.Add("")
            End If
            parCount += 1
            'цикл идет от ] в сторону начала строки и начинается внутри [] функции
            qbCount = 1 '(то есть, первый символ ] уже получен)
            For wId As Integer = funcWordId - 3 To 1 Step -1
                Select Case curCD.Code(wId).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount -= 1
                        If qbCount = 0 Then Exit For 'дошли до открывающей [
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount += 1
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        obCount += 1
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                        obCount -= 1
                    Case CodeTextBox.EditWordTypeEnum.W_COMMA
                        'найдена запятая
                        If obCount <> 0 OrElse qbCount <> 1 Then Continue For 'она стоит вне данной функции (является разделителем внутри вложенной функции или свойства)
                        'данная запятая - разделитель параметров этой функции. Получаем параметр сразу перед запятой
                        parCount += 1
                        If curCD.Code(wId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(wId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                            (wId >= 2 AndAlso (curCD.Code(wId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA OrElse curCD.Code(wId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN)) Then
                            'слово перед запятой является простой строкой или числом, при этом перед этим словом стоит , или [
                            lstContent.Insert(0, curCD.Code(wId - 1).Word.Trim) 'вставляем вновь полученный параметр на первое место в lstContent, так как идем от конца к началу
                        Else
                            'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                            lstContent.Insert(0, "")
                        End If
                End Select
            Next wId

            If elementParamId <= lstContent.Count - 1 Then
                'Если среди уже отобранных параметров есть параметр типа PARAM_ELEMENT, то возвращаем его (идти в круглые скобки функции/свойства уже не надо)
                Return lstContent(elementParamId)
            End If
        End If

        Dim codeLen As Integer = curCD.Code.Count
        If codeLen >= funcWordId + 3 AndAlso curCD.Code(funcWordId + 1).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN Then
            '...funcName(paramX, ...
            obCount = 1 'то есть, первый символ ( уже получен
            qbCount = 0
            If curCD.Code(funcWordId + 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(funcWordId + 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                (funcWordId <= codeLen - 4 AndAlso (curCD.Code(funcWordId + 3).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA OrElse curCD.Code(funcWordId + 3).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE)) Then
                'слово после ( является простой строкой или числом, при этом после этого слова стоит , или )
                lstContent.Add(curCD.Code(funcWordId + 2).Word.Trim)
            Else
                'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                lstContent.Add("")
            End If
            parCount += 1
            'цикл идет от ( в сторону конца строки и начинается внутри () функции
            For wId As Integer = funcWordId + 2 To codeLen - 1
                Select Case curCD.Code(wId).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount += 1
                    Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount -= 1
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        obCount -= 1
                        If obCount = 0 Then Exit For 'дошли до закрывающей (
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                        obCount += 1
                    Case CodeTextBox.EditWordTypeEnum.W_COMMA
                        'найдена запятая
                        If obCount <> 0 OrElse qbCount <> 1 Then Continue For 'она стоит вне данной функции (является разделителем внутри вложенной функции или свойства)
                        'данная запятая - разделитель параметров этой функции. Получаем параметр сразу перед запятой

                        If wId = codeLen - 1 Then Exit For
                        parCount += 1
                        If curCD.Code(wId + 1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(wId + 1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                            (wId <= codeLen - 3 AndAlso (curCD.Code(wId + 2).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA OrElse curCD.Code(wId + 2).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE)) Then
                            'слово после запятой является простой строкой или числом, при этом после этого слова стоит , или )
                            lstContent.Add(curCD.Code(wId - 1).Word.Trim)
                        Else
                            'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                            lstContent.Add("")
                        End If
                End Select
            Next wId
            If elementParamId <= lstContent.Count - 1 Then
                'На этом этапе получены в списке lstContent содержатся все параметры фнкции/класса (в правильном порядке). Если содержимое параметра - простое число/строка, то оно попало 
                'в данный список lstContent. Если же содержимое иное (не просчитывается на этапе написания кода), то вместо содержимого - пустая строка.
                Return lstContent(elementParamId) 'возвращаем параметр типа PARAM_ELEMENT
            End If
        End If
        Return "" 'Сюда, по идее, функция дойти не должна
    End Function

    ''' <summary>
    ''' Производит поиск свойства/функции перед математическим оператором / оператором сравнения, и, если их returnType = RETURN_ELEMENT, возвращает Id класса возвращаемого элемента
    ''' </summary>
    ''' <param name="exData">Скрипт в формате MatewScript.ExecuteDataType</param>
    ''' <param name="operatorId">Id математического опретора / оператора сравнения, идущего после строки-имени, в текущей строке кода</param>
    ''' <param name="line">текущая линия</param>
    ''' <returns>Id класса или -1</returns>
    Private Function GetLastPropertyReturnElementClassIdEx(ByRef exData As List(Of MatewScript.ExecuteDataType), ByVal operatorId As Integer, ByVal line As Integer) As Integer
        Dim i As Integer = operatorId - 1
        Dim curLine As Integer = line

        Dim qCount As Integer = 0 'равзница открытых/закрытых  [ ]
        Dim oCount As Integer = 0 'равзница открытых/закрытых  ( )
        Do
            If i <= 0 Then
                'Текущая строка закончилась. Переход к строке выше, если она заканчивается на _
                curLine -= 1
                If curLine < 0 Then Return -1
                If IsNothing(exData(curLine).Code) OrElse exData(curLine).Code.Count <= 1 Then Return -1
                If exData(curLine).Code.Last.wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION Then Return -1
                i = exData(curLine).Code.Count - 2
            End If

            Select Case exData(curLine).Code(i).wordType
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                    oCount -= 1
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                    oCount += 1
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                    qCount -= 1
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                    qCount += 1
                Case CodeTextBox.EditWordTypeEnum.W_PROPERTY, CodeTextBox.EditWordTypeEnum.W_FUNCTION
                    If qCount <> 0 OrElse oCount <> 0 Then Continue Do
                    'Свойство/функция найдена
                    Dim p As MatewScript.PropertiesInfoType = Nothing
                    If exData(curLine).Code(i).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                        If mScript.mainClass(exData(curLine).Code(i).classId).Functions.TryGetValue(exData(curLine).Code(i).Word.Trim, p) = False Then Return -1
                    Else
                        If mScript.mainClass(exData(curLine).Code(i).classId).Properties.TryGetValue(exData(curLine).Code(i).Word.Trim, p) = False Then Return -1
                    End If
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_FUNCTION Then Return -3
                    If p.returnType <> MatewScript.ReturnFunctionEnum.RETURN_ELEMENT OrElse IsNothing(p.returnArray) OrElse p.returnArray.Count = 0 Then Return -1
                    'И ее тип = RETURN_ELEMENT
                    Dim retClass As Integer = -1
                    mScript.mainClassHash.TryGetValue(p.returnArray(0), retClass)
                    Return retClass
                Case CodeTextBox.EditWordTypeEnum.W_PARAM, CodeTextBox.EditWordTypeEnum.W_RETURN, CodeTextBox.EditWordTypeEnum.W_SWITCH, CodeTextBox.EditWordTypeEnum.W_VARIABLE
                    If qCount <> 0 OrElse oCount <> 0 Then Continue Do
                    Return -1
            End Select

            i -= 1
        Loop
    End Function

    ''' <summary>
    ''' Производит поиск свойства/функции перед математическим оператором / оператором сравнения, и, если их returnType = RETURN_ELEMENT, возвращает Id класса возвращаемого элемента
    ''' </summary>
    ''' <param name="cd">Скрипт в формате CodeTextBox.CodeDataType</param>
    ''' <param name="operatorId">Id математического опретора / оператора сравнения, идущего после строки-имени, в текущей строке кода</param>
    ''' <param name="line">текущая линия</param>
    ''' <returns>Id класса или -1</returns>
    Private Function GetLastPropertyReturnElementClassId(ByRef cd() As CodeTextBox.CodeDataType, ByVal operatorId As Integer, ByVal line As Integer) As Integer
        Dim i As Integer = operatorId - 1
        Dim curLine As Integer = line

        Dim qCount As Integer = 0 'равзница открытых/закрытых  [ ]
        Dim oCount As Integer = 0 'равзница открытых/закрытых  ( )
        Do
            If i <= 0 Then
                'Текущая строка закончилась. Переход к строке выше, если она заканчивается на _
                curLine -= 1
                If curLine < 0 Then Return -1
                If IsNothing(cd(curLine).Code) OrElse cd(curLine).Code.Count <= 1 Then Return -1
                If cd(curLine).Code.Last.wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION Then Return -1
                i = cd(curLine).Code.Count - 2
            End If

            Select Case cd(curLine).Code(i).wordType
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                    oCount -= 1
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                    oCount += 1
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                    qCount -= 1
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                    qCount += 1
                Case CodeTextBox.EditWordTypeEnum.W_PROPERTY, CodeTextBox.EditWordTypeEnum.W_FUNCTION
                    'Свойство/функция найдена
                    If qCount <> 0 OrElse oCount <> 0 Then Continue Do

                    Dim p As MatewScript.PropertiesInfoType = Nothing
                    If cd(curLine).Code(i).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                        If mScript.mainClass(cd(curLine).Code(i).classId).Functions.TryGetValue(cd(curLine).Code(i).Word.Trim, p) = False Then Return -1
                    Else
                        If mScript.mainClass(cd(curLine).Code(i).classId).Properties.TryGetValue(cd(curLine).Code(i).Word.Trim, p) = False Then Return -1
                    End If
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_FUNCTION Then Return -3
                    If p.returnType <> MatewScript.ReturnFunctionEnum.RETURN_ELEMENT OrElse IsNothing(p.returnArray) OrElse p.returnArray.Count = 0 Then Return -1
                    'И ее тип = RETURN_ELEMENT
                    Dim retClass As Integer = -1
                    mScript.mainClassHash.TryGetValue(p.returnArray(0), retClass)
                    Return retClass
                Case CodeTextBox.EditWordTypeEnum.W_PARAM, CodeTextBox.EditWordTypeEnum.W_RETURN, CodeTextBox.EditWordTypeEnum.W_SWITCH, CodeTextBox.EditWordTypeEnum.W_VARIABLE
                    If qCount <> 0 OrElse oCount <> 0 Then Continue Do
                    Return -1
            End Select

            i -= 1
        Loop
    End Function

    ''' <summary>
    ''' Рассчитывает положение начала текущего слова в указанной строке скрипта. Если слово в 'кавычках', то возвращает положение начала имеено слова, за начальной кавычкой
    ''' </summary>
    ''' <param name="cd">Строка скрипта</param>
    ''' <param name="wordId">Id слова в строке</param>
    ''' <returns></returns>
    Private Function GetWordPosInLine(ByRef cd As CodeTextBox.CodeDataType, ByVal wordId As Integer) As Integer
        Dim pos As Integer = cd.StartingSpaces.Length
        For i As Integer = 0 To cd.Code.Count - 1
            If i < wordId Then
                pos = pos + cd.Code(i).Word.Length
            ElseIf i = wordId Then
                Dim w As String = cd.Code(i).Word
                If String.IsNullOrEmpty(w) = False AndAlso w.Chars(0) = "'"c Then
                    pos += 1
                End If
                Exit For
            Else
                Exit For
            End If
        Next i
        Return pos
    End Function
#End Region

#Region "Find String"
    Private lstEndWord As List(Of String) = {" ", vbCr, vbLf, "<", ">", "(", ")", "[", "]", "!", "@", "'", Chr(34), "#", "№", "$", ";", "%", "^", ":", "&", "?", "*", "-", "+", "=", ",", "."}.ToList

    ''' <summary>
    ''' Ищет указанный текст во всех свойствах, функциях, переменных...
    ''' </summary>
    ''' <param name="strSearch">Искомая строка (без кавычек)</param>
    Public Function FindStringInStruct(ByVal strSearch As String, ByVal wholeWord As Boolean, ByVal caseSensitive As Boolean) As Boolean
        dlgEntrancies.BeginNewEntrancies("Поиск указанной строки дал следующие результаты:", dlgEntrancies.EntranciesStyleEnum.Extended)

        'Прочесывание структуры mainClass на предмет наличия свойств, содержащих указанную строку
        Dim compar As System.StringComparison
        If caseSensitive Then
            compar = StringComparison.CurrentCulture
        Else
            compar = StringComparison.CurrentCultureIgnoreCase
        End If

        Dim wordStart As Integer = -1
        Dim ret As MatewScript.ContainsCodeEnum
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'Поиск в свойствах
                Dim parId As Integer = -1 'Id локации, если активная панель - панель действий/локаций. Надо только для событий действий
                If mScript.mainClass(classId).Names(0) = "A" Then
                    parId = actionsRouter.GetActiveLocationId
                End If
                'проверка свойств каждого класса
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key

                    'Свойство по умолчанию
                    If String.IsNullOrEmpty(p.Value) Then p.Value = ""
                    ret = mScript.IsPropertyContainsCode(p.Value)
                    If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso wholeWord Then
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                        FindWholeWordInExecutableString(p.Value, strSearch, compar)
                    ElseIf ret = MatewScript.ContainsCodeEnum.NOT_CODE OrElse ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                        'Сравнение свойства с искомой строкой
                        wordStart = -1
                        If wholeWord Then
                            If String.Compare(mScript.PrepareStringToPrint(p.Value, Nothing, False), strSearch, Not caseSensitive) = 0 Then wordStart = 0
                        Else
                            wordStart = mScript.PrepareStringToPrint(p.Value, Nothing, False).IndexOf(strSearch, compar)
                        End If

                        If wordStart > -1 Then
                            dlgEntrancies.NewEntrance(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, strSearch, wordStart)
                        End If
                    Else
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                        FindStringInExPropery(p.Value, ret, strSearch, wholeWord, compar)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                        'Свойства 2 уровня
                        For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                            If String.IsNullOrEmpty(ch.Value) Then ch.Value = ""
                            ret = mScript.IsPropertyContainsCode(ch.Value)
                            wordStart = -1
                            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso wholeWord Then
                                dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, parId)
                                FindWholeWordInExecutableString(ch.Value, strSearch, compar)
                            ElseIf ret = MatewScript.ContainsCodeEnum.NOT_CODE OrElse ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                                If wholeWord Then
                                    If String.Compare(mScript.PrepareStringToPrint(ch.Value, Nothing, False), strSearch, Not caseSensitive) = 0 Then wordStart = 0
                                Else
                                    wordStart = mScript.PrepareStringToPrint(ch.Value, Nothing, False).IndexOf(strSearch, compar)
                                End If

                                If wordStart > -1 Then
                                    dlgEntrancies.NewEntrance(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, parId, 0, strSearch, wordStart)
                                End If
                            Else
                                dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, parId)
                                FindStringInExPropery(ch.Value, ret, strSearch, wholeWord, compar)
                            End If

                            If IsNothing(ch.ThirdLevelProperties) = False Then
                                'Свойства 3 уровня
                                For child3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                    Dim thrdValue As String = ""
                                    'ch.ThirdLevelProperties(child3Id)
                                    If String.IsNullOrEmpty(ch.ThirdLevelProperties(child3Id)) = False Then thrdValue = ch.ThirdLevelProperties(child3Id)

                                    ret = mScript.IsPropertyContainsCode(thrdValue)
                                    wordStart = -1
                                    If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso wholeWord Then
                                        dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                                        FindWholeWordInExecutableString(thrdValue, strSearch, compar)
                                    ElseIf ret = MatewScript.ContainsCodeEnum.NOT_CODE OrElse ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                                        If wholeWord Then
                                            If String.Compare(mScript.PrepareStringToPrint(thrdValue, Nothing, False), strSearch, Not caseSensitive) = 0 Then wordStart = 0
                                        Else
                                            wordStart = mScript.PrepareStringToPrint(thrdValue, Nothing, False).IndexOf(strSearch, compar)
                                        End If

                                        If wordStart > -1 Then
                                            dlgEntrancies.NewEntrance(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, strSearch, wordStart)
                                        End If
                                    Else
                                        dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                                        FindStringInExPropery(thrdValue, ret, strSearch, wholeWord, compar)
                                    End If
                                Next child3Id
                            End If
                        Next child2Id
                    End If
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'Поиск в функциях класса
                'проверка функций каждого класса
                For fId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    Dim f As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(fId).Value
                    Dim fName As String = mScript.mainClass(classId).Functions.ElementAt(fId).Key

                    If String.IsNullOrEmpty(f.Value) Then Continue For
                    ret = mScript.IsPropertyContainsCode(f.Value)
                    dlgEntrancies.SetEntranceDefault(classId, -1, -1, fName, CodeTextBox.EditWordTypeEnum.W_FUNCTION)
                    FindStringInExPropery(f.Value, MatewScript.ContainsCodeEnum.CODE, strSearch, wholeWord, compar)
                Next fId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locId As Integer = GetSecondChildIdByName(actionsRouter.lstActions.ElementAt(i).Key, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        ret = mScript.IsPropertyContainsCode(ch.Value)
                        If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso wholeWord Then
                            dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                            FindWholeWordInExecutableString(ch.Value, strSearch, compar)
                        ElseIf ret = MatewScript.ContainsCodeEnum.NOT_CODE OrElse ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                            'Сравнение свойства с искомой строкой
                            wordStart = -1
                            If wholeWord Then
                                If String.Compare(mScript.PrepareStringToPrint(ch.Value, Nothing, False), strSearch, Not caseSensitive) = 0 Then wordStart = 0
                            Else
                                If String.IsNullOrEmpty(ch.Value) = False Then wordStart = mScript.PrepareStringToPrint(ch.Value, Nothing, False).IndexOf(strSearch, compar)
                            End If

                            If wordStart > -1 Then
                                dlgEntrancies.NewEntrance(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId, 0, strSearch, wordStart)
                            End If
                        Else
                            dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                            FindStringInExPropery(ch.Value, ret, strSearch, wholeWord, compar)
                        End If

                    Next pId
                Next aId
            Next i
        End If

        'Поиск в функциях Писателя
        If IsNothing(mScript.functionsHash) = False Then
            For fId As Integer = 0 To mScript.functionsHash.Count - 1
                dlgEntrancies.SetEntranceDefault(-3, fId, -1, mScript.functionsHash.ElementAt(fId).Key, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION)
                FindStringInExProperyData(mScript.functionsHash.ElementAt(fId).Value.ValueDt, strSearch, wholeWord, compar)
            Next fId
        End If

        'Поиск в событиях изменения свойства
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) = False AndAlso mScript.trackingProperties.lstTrackingProperties.Count > 0 Then
            For tId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
                Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(tId).Value
                Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(tId).Key, classId As Integer, propName As String
                Dim pos As Integer = strKey.IndexOf("."c)
                If pos = -1 Then Continue For
                classId = mScript.mainClassHash(strKey.Substring(0, pos))
                propName = strKey.Substring(pos + 1)

                If tr.eventBeforeId > 0 Then
                    dlgEntrancies.SetEntranceDefault(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -2, -1, "", -1, frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                    FindStringInExProperyData(tr.propBeforeContent, strSearch, wholeWord, compar)
                End If

                If tr.eventAfterId > 0 Then
                    dlgEntrancies.SetEntranceDefault(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -2, -1, "", -1, frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                    FindStringInExProperyData(tr.propAfterContent, strSearch, wholeWord, compar)
                End If
            Next tId
        End If

        'Поиск в переменных
        If IsNothing(mScript.csPublicVariables.lstVariables) = False Then
            For vId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
                Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(vId).Value.arrValues
                Dim vName As String = mScript.csPublicVariables.lstVariables.ElementAt(vId).Key
                If IsNothing(arrValues) OrElse arrValues.Count = 0 Then Continue For
                For arrId = 0 To arrValues.Count - 1
                    ret = mScript.IsPropertyContainsCode(arrValues(arrId))
                    If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso wholeWord Then
                        dlgEntrancies.SetEntranceDefault(-2, vId, arrId, vName, CodeTextBox.EditWordTypeEnum.W_VARIABLE)
                        FindWholeWordInExecutableString(arrValues(arrId), strSearch, compar)
                    ElseIf ret = MatewScript.ContainsCodeEnum.NOT_CODE OrElse ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                        'Сравнение свойства с искомой строкой
                        wordStart = -1
                        If wholeWord Then
                            If String.Compare(mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False), strSearch, Not caseSensitive) = 0 Then wordStart = 0
                        Else
                            wordStart = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False).IndexOf(strSearch, compar)
                        End If

                        If wordStart > -1 Then
                            dlgEntrancies.NewEntrance(-2, vId, arrId, vName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, 0, strSearch, wordStart)
                        End If
                    Else
                        dlgEntrancies.SetEntranceDefault(-2, vId, arrId, vName, CodeTextBox.EditWordTypeEnum.W_VARIABLE)
                        FindStringInExPropery(arrValues(arrId), ret, strSearch, wholeWord, compar)
                    End If
                Next arrId
            Next vId
        End If


        If dlgEntrancies.hasEntrancies Then
            dlgEntrancies.Show()
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Поиск совпадения исполняемой строке EXECUTABLE_STRING при условии, что ищем слово целиком
    ''' </summary>
    ''' <param name="strScript">Текст скрипта в формате xml</param>
    ''' <param name="strSearch">Строка для поиска</param>
    ''' <param name="compar">Для с равнения с учетом/без учета регистра</param>
    Private Sub FindWholeWordInExecutableString(ByVal strScript As String, ByVal strSearch As String, ByRef compar As System.StringComparison) ', Optional locationId As Integer = -1)
        strScript = mScript.PrepareStringToPrint(strScript, Nothing, False)
        Dim sLength As Integer = strSearch.Length
        Dim pos As Integer = 0
        Do
            pos = strScript.IndexOf(strSearch, pos, compar)
            If pos = -1 Then Return
            'Это начало слова?
            Dim startOk As Boolean = False
            If pos = 0 Then
                startOk = True
            Else
                If lstEndWord.IndexOf(strScript.Chars(pos - 1)) > -1 Then startOk = True
            End If
            If startOk = False Then
                pos += 1
                Continue Do
            End If

            'Это конец слова?
            Dim endOk As Boolean = False
            If pos + sLength - 1 >= strScript.Length Then
                endOk = True
            Else
                If lstEndWord.IndexOf(strScript.Chars(pos + sLength)) > -1 Then endOk = True
            End If
            If endOk = False Then
                pos += 1
                Continue Do
            End If
            'Добавляем вхождения
            dlgEntrancies.SetSeekPosInfo(0, strSearch, pos)
            'If locationId > -1 Then dlgEntrancies.SetSeekLocationId(locationId)
            dlgEntrancies.NewEntrance()
            pos += strSearch.Length
        Loop
    End Sub

    ''' <summary>
    ''' Поиск совпадения в сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="strXml">Скрипт / длинный текст (но НЕ в формате EXECUTABLE_STRING)</param>
    ''' <param name="containsCodeFormat">Формат - скрипт или длинный текст</param>
    ''' <param name="strSearch">Строка для поиска</param>
    ''' <param name="wholeWord">Искать целое слово?</param>
    ''' <param name="compar">Учет регистра</param>
    ''' <returns></returns>
    Private Function FindStringInExPropery(ByVal strXml As String, ByVal containsCodeFormat As MatewScript.ContainsCodeEnum, ByVal strSearch As String, ByVal wholeWord As Boolean, _
                                           ByRef compar As System.StringComparison) As Boolean
        If String.IsNullOrEmpty(strXml) Then Return False

        'получаем структуру кода cd
        Dim strText As String
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            .IsTextBlockByDefault = (containsCodeFormat = MatewScript.ContainsCodeEnum.LONG_TEXT)
            .LoadCodeFromProperty(strXml)
            strText = .Text
        End With
        Dim sLength As Integer = strSearch.Length

        Dim wasFound As Boolean = False
        Dim pos As Integer = 0
        Dim line As Integer = -1
        Do
            pos = strText.IndexOf(strSearch, pos, compar)
            If pos = -1 Then Return wasFound
            If wholeWord Then
                'Если ищем слово целиком, то определяем начало и конец найденной строки
                Dim startOk As Boolean = False
                If pos = 0 Then
                    startOk = True
                Else
                    If lstEndWord.IndexOf(strText.Chars(pos - 1)) > -1 Then startOk = True
                End If

                If startOk = False Then
                    pos += 1
                    Continue Do
                End If

                Dim endOk As Boolean = False
                If pos + sLength - 1 >= strText.Length Then
                    endOk = True
                Else
                    If lstEndWord.IndexOf(strText.Chars(pos + sLength)) > -1 Then endOk = True
                End If
                If endOk = False Then
                    pos += 1
                    Continue Do
                End If
            End If

            line = questEnvironment.codeBoxShadowed.codeBox.GetLineFromCharIndex(pos)
            dlgEntrancies.SetSeekPosInfo(line, strSearch, pos - questEnvironment.codeBoxShadowed.codeBox.GetFirstCharIndexFromLine(line))
            dlgEntrancies.NewEntrance()
            pos += strSearch.Length
            wasFound = True
        Loop
        Return wasFound
    End Function

    ''' <summary>
    ''' Поиск совпадения в сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="dt">Скрипт / длинный текст (но НЕ в формате EXECUTABLE_STRING)</param>
    ''' <param name="strSearch">Строка для поиска</param>
    ''' <param name="wholeWord">Искать целое слово?</param>
    ''' <param name="compar">Учет регистра</param>
    ''' <returns></returns>
    Private Function FindStringInExProperyData(ByVal dt() As CodeTextBox.CodeDataType, ByVal strSearch As String, ByVal wholeWord As Boolean, _
                                           ByRef compar As System.StringComparison) As Boolean
        If IsNothing(dt) OrElse dt.Count = 0 Then Return False

        'получаем структуру кода cd
        Dim strText As String
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            .IsTextBlockByDefault = False
            .LoadCodeFromCodeData(dt)
            strText = .Text
        End With
        Dim sLength As Integer = strSearch.Length

        Dim wasFound As Boolean = False
        Dim pos As Integer = 0
        Dim line As Integer = -1
        Do
            pos = strText.IndexOf(strSearch, pos, compar)
            If pos = -1 Then Return wasFound
            If wholeWord Then
                'Если ищем слово целиком, то определяем начало и конец найденной строки
                Dim startOk As Boolean = False
                If pos = 0 Then
                    startOk = True
                Else
                    If lstEndWord.IndexOf(strText.Chars(pos - 1)) > -1 Then startOk = True
                End If

                If startOk = False Then
                    pos += 1
                    Continue Do
                End If

                Dim endOk As Boolean = False
                If pos + sLength - 1 >= strText.Length Then
                    endOk = True
                Else
                    If lstEndWord.IndexOf(strText.Chars(pos + sLength)) > -1 Then endOk = True
                End If
                If endOk = False Then
                    pos += 1
                    Continue Do
                End If
            End If

            line = questEnvironment.codeBoxShadowed.codeBox.GetLineFromCharIndex(pos)
            dlgEntrancies.SetSeekPosInfo(line, strSearch, pos - questEnvironment.codeBoxShadowed.codeBox.GetFirstCharIndexFromLine(line))
            dlgEntrancies.NewEntrance()
            pos += strSearch.Length
            wasFound = True
        Loop
        Return wasFound
    End Function
#End Region

#Region "Replace String"
    ''' <summary>
    ''' Производит замену строки текста из указанного вхождения
    ''' </summary>
    ''' <param name="ent">Класс вхождения, содержащий всю информацию о заменяемой строке</param>
    ''' <param name="newName">Новая строка</param>
    ''' <returns>False при ошибке</returns>
    Public Function ReplaceString(ByRef ent As dlgEntrancies.cEntranciesClass, ByVal newName As String) As Boolean
        Select Case ent.elementType
            Case CodeTextBox.EditWordTypeEnum.W_PROPERTY
                'Вхождение - в свойствах
                Dim propName As String = ent.elementName
                Dim val As String = ""
                Dim eventId As Integer
                Dim ret As MatewScript.ReturnFormatEnum
                Dim isInActions As Boolean = False
                Dim locName As String = ""
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing

                If ent.tracking <> frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT Then
                    'Вхождение - в событиях изменения свойства
                    Dim trName As String = mScript.mainClass(ent.classId).Names(0) & "." & ent.elementName
                    Dim f As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties(trName)

                    'Получаем десериализованный скрипт 
                    Dim strText As String = ""
                    With questEnvironment.codeBoxShadowed.codeBox
                        .Text = ""
                        .IsTextBlockByDefault = False
                        If ent.tracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE Then
                            .LoadCodeFromCodeData(f.propBeforeContent)
                        Else
                            .LoadCodeFromCodeData(f.propAfterContent)
                        End If
                        strText = .Text
                    End With

                    Try
                        'Выделяем заменяемое слово
                        Dim selStart As Integer = questEnvironment.codeBoxShadowed.codeBox.GetFirstCharIndexFromLine(ent.seekLine) + ent.wordStart
                        questEnvironment.codeBoxShadowed.codeBox.Select(selStart, ent.word.Length)
                        If String.Compare(questEnvironment.codeBoxShadowed.codeBox.SelectedText, ent.word, True) <> 0 Then Return False
                        questEnvironment.codeBoxShadowed.codeBox.SelectedText = newName 'производим замену
                        'Проверка не произошло ли ошибки
                        questEnvironment.codeBoxShadowed.codeBox.PrepareText(questEnvironment.codeBoxShadowed.codeBox, ent.seekLine, ent.seekLine, False)
                        If mScript.LAST_ERROR.Length > 0 OrElse questEnvironment.codeBoxShadowed.codeBox.CodeData(ent.seekLine).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_ERROR Then
                            mScript.LAST_ERROR = ""
                            Return False
                        End If
                        'Устанавливаем новое значение событию
                        If ent.tracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE Then
                            f.propBeforeContent = CopyCodeDataArray(questEnvironment.codeBoxShadowed.codeBox.CodeData)
                            mScript.eventRouter.SetEventId(f.eventBeforeId, mScript.PrepareBlock(f.propBeforeContent))
                        Else
                            f.propAfterContent = CopyCodeDataArray(questEnvironment.codeBoxShadowed.codeBox.CodeData)
                            mScript.eventRouter.SetEventId(f.eventAfterId, mScript.PrepareBlock(f.propAfterContent))
                        End If
                    Catch ex As Exception
                        Return False
                    End Try
                    Return True
                ElseIf ent.child2Id < 0 Then
                    'Вхождение - 1 уровень
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(ent.classId).Properties(propName)
                    eventId = p.eventId
                    ret = mScript.Param_GetType(p.Value)
                    If ret = MatewScript.ReturnFormatEnum.TO_STRING Then
                        val = mScript.PrepareStringToPrint(p.Value, Nothing, False)
                    Else
                        val = p.Value
                    End If
                ElseIf ent.child3Id < 0 Then
                    'Вхождение - 2 уровень
                    isInActions = (ent.classId = mScript.mainClassHash("A") AndAlso ent.child2Id > -1 AndAlso actionsRouter.GetActiveLocationId <> ent.parentId)
                    Dim p As MatewScript.ChildPropertiesInfoType = Nothing
                    If isInActions Then
                        'Вхождение - в сохраненных действиях
                        locName = mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(ent.parentId)("Name").Value
                        arrProp = Nothing
                        If actionsRouter.lstActions.TryGetValue(locName, arrProp) = False OrElse IsNothing(arrProp) OrElse ent.child2Id > arrProp.Count - 1 Then Return False
                        p = arrProp(ent.child2Id)(propName)
                    Else
                        'Вхождение - в других элементах (не в сохраненных действиях)
                        If ent.child2Id > mScript.mainClass(ent.classId).ChildProperties.Count - 1 Then Return False
                        p = mScript.mainClass(ent.classId).ChildProperties(ent.child2Id)(propName)
                    End If
                    eventId = p.eventId
                    ret = mScript.Param_GetType(p.Value)
                    If ret = MatewScript.ReturnFormatEnum.TO_STRING Then
                        val = mScript.PrepareStringToPrint(p.Value, Nothing, False)
                    Else
                        val = p.Value
                    End If
                Else
                    'Вхождение - 3 уровень
                    If ent.child2Id > mScript.mainClass(ent.classId).ChildProperties.Count - 1 Then Return False
                    If ent.child3Id > mScript.mainClass(ent.classId).ChildProperties(ent.child2Id)("Name").ThirdLevelProperties.Count - 1 Then Return False
                    Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(ent.classId).ChildProperties(ent.child2Id)(propName)
                    eventId = p.ThirdLevelEventId(ent.child3Id)
                    ret = mScript.Param_GetType(p.ThirdLevelProperties(ent.child3Id))
                    If ret = MatewScript.ReturnFormatEnum.TO_STRING Then
                        val = mScript.PrepareStringToPrint(p.ThirdLevelProperties(ent.child3Id), Nothing, False)
                    Else
                        val = p.ThirdLevelProperties(ent.child3Id)
                    End If
                End If

                Dim codeType As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(val)
                If eventId > 0 AndAlso codeType <> MatewScript.ContainsCodeEnum.NOT_CODE Then
                    'Свойство вхождения - скрипт
                    'Получаем десериализованный скрипт
                    Dim strText As String = ""
                    With questEnvironment.codeBoxShadowed.codeBox
                        .Text = ""
                        .IsTextBlockByDefault = (codeType = MatewScript.ContainsCodeEnum.LONG_TEXT)
                        If codeType = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                            .Text = mScript.PrepareStringToPrint(val, Nothing, False)
                        Else
                            .LoadCodeFromProperty(val)
                        End If
                        strText = .Text
                    End With
                    Try
                        'Выделяем заменяемое слово
                        Dim selStart As Integer = questEnvironment.codeBoxShadowed.codeBox.GetFirstCharIndexFromLine(ent.seekLine) + ent.wordStart
                        questEnvironment.codeBoxShadowed.codeBox.Select(selStart, ent.word.Length)
                        If String.Compare(questEnvironment.codeBoxShadowed.codeBox.SelectedText, ent.word, True) <> 0 Then Return False
                        questEnvironment.codeBoxShadowed.codeBox.SelectedText = newName 'производим замену
                        'Проверка не произошло ли ошибки
                        questEnvironment.codeBoxShadowed.codeBox.PrepareText(questEnvironment.codeBoxShadowed.codeBox, ent.seekLine, ent.seekLine, False)
                        If mScript.LAST_ERROR.Length > 0 OrElse questEnvironment.codeBoxShadowed.codeBox.CodeData(ent.seekLine).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_ERROR Then
                            mScript.LAST_ERROR = ""
                            Return False
                        End If

                        If isInActions Then
                            val = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(questEnvironment.codeBoxShadowed.codeBox.CodeData)
                            Dim p As MatewScript.ChildPropertiesInfoType = arrProp(ent.child2Id)(propName)
                            p.Value = val
                            p.eventId = mScript.eventRouter.SetEventId(eventId, questEnvironment.codeBoxShadowed.codeBox.CodeData)
                        Else
                            mScript.eventRouter.SetPropertyWithEvent(ent.classId, ent.elementName, questEnvironment.codeBoxShadowed.codeBox.CodeData, ent.child2Id.ToString, ent.child3Id.ToString, _
                                                                     questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(questEnvironment.codeBoxShadowed.codeBox.CodeData), _
                                                                     codeType = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING, False)
                        End If
                    Catch ex As Exception
                        Return False
                    End Try
                Else
                    'Свойство вхождения - простое свойство
                    Dim wLength As Integer = ent.word.Length
                    If String.IsNullOrEmpty(val) OrElse val.Length < ent.wordStart + wLength OrElse String.Compare(val.Substring(ent.wordStart, wLength), ent.word, True) <> 0 Then
                        Return False
                    End If

                    Dim newVal As String = ""
                    If ent.wordStart > 0 Then newVal = val.Substring(0, ent.wordStart)
                    newVal = newVal & newName
                    If val.Length > ent.wordStart + wLength Then
                        newVal = newVal & val.Substring(ent.wordStart + wLength)
                    End If

                    If ret = MatewScript.ReturnFormatEnum.TO_STRING Then newVal = WrapString(newVal)
                    If isInActions Then
                        Dim p As MatewScript.ChildPropertiesInfoType = arrProp(ent.child2Id)(propName)
                        p.Value = newVal
                        If eventId > 0 Then
                            With questEnvironment.codeBoxShadowed.codeBox
                                .Text = ""
                                .IsTextBlockByDefault = False
                                .Text = mScript.PrepareStringToPrint(newVal, Nothing, False)
                                mScript.eventRouter.SetEventId(eventId, .CodeData)
                            End With
                        End If
                    Else
                        'Устанавливаем новое значение свойства
                        SetPropertyValue(ent.classId, propName, newVal, ent.child2Id, ent.child3Id)
                    End If
                End If
            Case CodeTextBox.EditWordTypeEnum.W_FUNCTION
                'Вхождение - в функциях классов
                Dim f As MatewScript.PropertiesInfoType = mScript.mainClass(ent.classId).Functions(ent.elementName)

                'Получаем десериализованный скрипт
                Dim strText As String = ""
                With questEnvironment.codeBoxShadowed.codeBox
                    .Text = ""
                    .IsTextBlockByDefault = False
                    .LoadCodeFromProperty(f.Value)
                    strText = .Text
                End With

                Try
                    'Выделяем заменяемое слово
                    Dim selStart As Integer = questEnvironment.codeBoxShadowed.codeBox.GetFirstCharIndexFromLine(ent.seekLine) + ent.wordStart
                    questEnvironment.codeBoxShadowed.codeBox.Select(selStart, ent.word.Length)
                    If String.Compare(questEnvironment.codeBoxShadowed.codeBox.SelectedText, ent.word, True) <> 0 Then Return False
                    questEnvironment.codeBoxShadowed.codeBox.SelectedText = newName 'производим замену
                    'Проверка не произошло ли ошибки
                    questEnvironment.codeBoxShadowed.codeBox.PrepareText(questEnvironment.codeBoxShadowed.codeBox, ent.seekLine, ent.seekLine, False)
                    If mScript.LAST_ERROR.Length > 0 OrElse questEnvironment.codeBoxShadowed.codeBox.CodeData(ent.seekLine).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_ERROR Then
                        mScript.LAST_ERROR = ""
                        Return False
                    End If
                    'Устанавливаем новое значение функции
                    mScript.eventRouter.SetPropertyWithEvent(ent.classId, ent.elementName, questEnvironment.codeBoxShadowed.codeBox.CodeData, ent.child2Id.ToString, ent.child3Id.ToString, _
                                                         questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(questEnvironment.codeBoxShadowed.codeBox.CodeData), _
                                                         False, True)
                Catch ex As Exception
                    Return False
                End Try
            Case CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION
                'Вхождение - в функциях Писателя
                Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(ent.elementName)

                'Получаем десериализованный скрипт
                Dim strText As String = ""
                With questEnvironment.codeBoxShadowed.codeBox
                    .Text = ""
                    .IsTextBlockByDefault = False
                    .LoadCodeFromCodeData(f.ValueDt)
                    strText = .Text
                End With

                Try
                    'Выделяем заменяемое слово
                    Dim selStart As Integer = questEnvironment.codeBoxShadowed.codeBox.GetFirstCharIndexFromLine(ent.seekLine) + ent.wordStart
                    questEnvironment.codeBoxShadowed.codeBox.Select(selStart, ent.word.Length)
                    If String.Compare(questEnvironment.codeBoxShadowed.codeBox.SelectedText, ent.word, True) <> 0 Then Return False
                    questEnvironment.codeBoxShadowed.codeBox.SelectedText = newName 'производим замену
                    'Проверка не произошло ли ошибки
                    questEnvironment.codeBoxShadowed.codeBox.PrepareText(questEnvironment.codeBoxShadowed.codeBox, ent.seekLine, ent.seekLine, False)
                    If mScript.LAST_ERROR.Length > 0 OrElse questEnvironment.codeBoxShadowed.codeBox.CodeData(ent.seekLine).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_ERROR Then
                        mScript.LAST_ERROR = ""
                        Return False
                    End If
                    'Устанавливаем новое значение функции
                    f.ValueDt = CopyCodeDataArray(questEnvironment.codeBoxShadowed.codeBox.CodeData)
                    f.ValueExecuteDt = mScript.PrepareBlock(f.ValueDt)
                Catch ex As Exception
                    Return False
                End Try
            Case CodeTextBox.EditWordTypeEnum.W_VARIABLE
                'Вхождение - в переменных
                Dim v As cVariable.variableEditorInfoType = Nothing
                If mScript.csPublicVariables.lstVariables.TryGetValue(ent.elementName, v) = False Then Return False
                Dim arrId As Integer = ent.child3Id
                If arrId > v.arrValues.Count - 1 Then Return False

                Dim val As String = v.arrValues(arrId)
                Dim ret As MatewScript.ReturnFormatEnum = mScript.Param_GetType(val)
                If ret = MatewScript.ReturnFormatEnum.TO_STRING Then
                    val = mScript.PrepareStringToPrint(val, Nothing, False)
                End If

                Dim wLength As Integer = ent.word.Length
                If String.IsNullOrEmpty(val) OrElse val.Length < ent.wordStart + wLength OrElse String.Compare(val.Substring(ent.wordStart, wLength), ent.word, True) <> 0 Then
                    Return False
                End If

                Dim newVal As String = ""
                If ent.wordStart > 0 Then newVal = val.Substring(0, ent.wordStart)
                newVal = newVal & newName
                If val.Length > ent.wordStart + wLength Then
                    newVal = newVal & val.Substring(ent.wordStart + wLength)
                End If

                If ret = MatewScript.ReturnFormatEnum.TO_STRING Then newVal = WrapString(newVal)
                'Устанавливаем новое значение переменной
                mScript.csPublicVariables.SetVariableInternal(ent.elementName, newVal, arrId)
        End Select
        Return True
    End Function
#End Region

#Region "Rename Functions"
    ''' <summary>
    ''' Переименовывает функцию Писателя во всех свойствах, функциях, перменных со скриптами и т. д.
    ''' </summary>
    ''' <param name="oldName">Старое имя функции (в кавычках)</param>
    ''' <param name="newName">Новое имя функции (в кавычках)</param>
    Public Sub RenameFunctionsInStruct(oldName As String, newName As String)
        'Прочесывание структуры mainClass на предмет наличия свойств типа RETURN_FUNCTION. Если найдем - переименовываем
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств каждого класса
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_FUNCTION Then
                        'тип = функция Писателя
                        Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                        'Свойство по умолчанию
                        If String.Compare(p.Value, oldName, True) = 0 Then
                            p.Value = newName
                        End If
                        If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                            'Свойства 2 уровня
                            For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                If String.Compare(ch.Value, oldName, True) = 0 Then
                                    ch.Value = newName
                                End If
                                If IsNothing(ch.ThirdLevelProperties) = False Then
                                    'Свойства 3 уровня
                                    For child3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                        If String.Compare(ch.ThirdLevelProperties(child3Id), oldName, True) = 0 Then
                                            ch.ThirdLevelProperties(child3Id) = newName
                                        End If
                                    Next child3Id
                                End If
                            Next child2Id
                        End If
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For
                        If pp.returnType = MatewScript.ReturnFunctionEnum.RETURN_FUNCTION Then
                            'тип = функция Писателя
                            If String.Compare(ch.Value, oldName, True) = 0 Then
                                ch.Value = newName
                            End If
                        End If
                    Next pId
                Next aId
            Next i
        End If

        RenameFunctionsInScripts(oldName, newName)
    End Sub

    ''' <summary>
    ''' Переименовывает функцию Писателя во всех скриптах
    ''' </summary>
    ''' <param name="oldName">Старое имя функции (в кавычках)</param>
    ''' <param name="newName">Новое имя функции (в кавычках)</param>
    Private Sub RenameFunctionsInScripts(ByVal oldName As String, ByVal newName As String)
        'Изменения в коде
        dlgEntrancies.BeginNewEntrancies("Найдены строки, идентичные старому имени функции, значение которых вычислить не удалось. Провертьте их самостоятельно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    If p.eventId > 0 AndAlso RenameFunctionInCode(p.eventId, oldName, newName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        p.Value = RenameFunctionInExPropery(p.Value, oldName, newName)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RenameFunctionInCode(ch.eventId, oldName, newName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                            ch.Value = RenameFunctionInExPropery(ch.Value, oldName, newName)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso RenameFunctionInCode(ch.ThirdLevelEventId(child3Id), oldName, newName) Then
                                'Значение - скрипт
                                dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                                ch.ThirdLevelProperties(child3Id) = RenameFunctionInExPropery(ch.ThirdLevelProperties(child3Id), oldName, newName)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If RenameFunctionInCode(p.eventId, oldName, newName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        p.Value = RenameFunctionInExPropery(p.Value, oldName, newName)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        If ch.eventId > 0 AndAlso RenameFunctionInCode(ch.eventId, oldName, newName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                            ch.Value = RenameFunctionInExPropery(ch.Value, oldName, newName)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        RenameFunctionInFunctionHash(oldName, newName)
        'В глобальных переменных
        RenameFunctionInVariables(oldName, newName)
        'В событиях изменения свойства
        RenameFunctionInTracking(oldName, newName)

        If dlgEntrancies.hasEntrancies Then
            dlgEntrancies.Show()
        End If
    End Sub

    ''' <summary>
    ''' Переименовывает функцию Писателя в глобальных переменных
    ''' </summary>
    ''' <param name="oldName">Старое имя функции</param>
    ''' <param name="newName">Новое имя функции</param>
    Private Sub RenameFunctionInVariables(ByVal oldName As String, ByVal newName As String)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return
        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            Dim wasFoundGlobal As Boolean = False 'были ли изменения хоть в одном элементе массива переменной
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim cd() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    cd = .codeBox.CodeData
                End With
                If IsNothing(cd) Then Continue For
                'В элементе массива - исполняемый код
                Dim wasFound As Boolean = False 'была ли замена в данном элементе массива
                For line As Integer = 0 To cd.Count - 1
                    If IsNothing(cd(line).Code) Then Continue For

                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            Dim res As String = SeekFunctionInString(cd, line, wordId, newName)
                            If res = "Entrance" Then
                                dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line)
                            ElseIf res.Length > 0 Then
                                cd(line).Code(wordId).Word = res
                                wasFound = True
                            End If
                        End If
                    Next wordId
                Next line
                If wasFound Then
                    wasFoundGlobal = True
                    If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                        'Исполняемая строка. Собираем измененную строку
                        Dim w As String = cd(0).StartingSpaces
                        For i As Integer = 0 To cd(0).Code.Count - 1
                            w += cd(0).Code(i).Word
                        Next i
                        w += cd(0).Comments
                        arrValues(arrId) = WrapString(w)
                    Else
                        'Скрипт или длинный текст. Собираем с внесенными изменениями
                        arrValues(arrId) = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
                    End If
                End If
            Next arrId
            If wasFoundGlobal Then
                'Сохраняем результат
                Dim var As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value
                var.arrValues = arrValues
            End If
        Next varId
    End Sub

    ''' <summary>
    ''' Переименовывает функцию Писателя в функциях писателя functionsHash
    ''' </summary>
    ''' <param name="oldName">Старое имя функции</param>
    ''' <param name="newName">Новое имя функции</param>
    Private Sub RenameFunctionInFunctionHash(ByVal oldName As String, ByVal newName As String)
        If IsNothing(mScript.functionsHash) Then Return
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            'В функции хранится код в исполняемом виде - в ValueExecuteDt, и в виде для редактирования - в ValueDt (при этом второе может быть пустым, если функция объявлена из кода)
            'Первое быть пустым не должно
            Dim wasFound As Boolean = False
            If IsNothing(func.ValueExecuteDt) = False Then
                Dim edt As List(Of MatewScript.ExecuteDataType) = func.ValueExecuteDt
                For line As Integer = 0 To edt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(edt(line).Code) Then Continue For
                    For wordId As Integer = 1 To edt(line).Code.Count - 1 '1 словом быть не может
                        If edt(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(edt(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If edt(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(edt(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                Dim res As String = SeekFunctionInStringEx(edt, line, wordId, newName)
                                If res = "Entrance" Then
                                    dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line)
                                ElseIf res.Length > 0 Then
                                    edt(line).Code(wordId).Word = res
                                    wasFound = True
                                End If
                            End If
                        End If
                    Next wordId
                Next line
            End If

            If Not wasFound Then Continue For 'если не найдено в ValueExecuteDt, то здесь тоже не будет
            If IsNothing(func.ValueDt) = False Then
                Dim cd() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перебираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                Dim res As String = SeekFunctionInString(cd, line, wordId, newName)
                                If res = "Entrance" Then
                                    dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line)
                                ElseIf res.Length > 0 Then
                                    cd(line).Code(wordId).Word = res
                                    wasFound = True
                                End If
                            End If
                        End If
                    Next wordId
                Next line
            End If
        Next funcId
    End Sub

    ''' <summary>
    ''' Переименовывает функцию Писателя в событиях изменения свойства
    ''' </summary>
    ''' <param name="oldName">Старое имя функции</param>
    ''' <param name="newName">Новое имя функции</param>
    Private Sub RenameFunctionInTracking(ByVal oldName As String, ByVal newName As String)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return
        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            If IsNothing(tr.propBeforeContent) = False Then
                If RenameFunctionInCode(tr.eventBeforeId, oldName, newName) Then
                    Dim cd() As CodeTextBox.CodeDataType = tr.propBeforeContent
                    For line As Integer = 0 To cd.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(cd(line).Code) Then Continue For
                        For wordId As Integer = 0 To cd(line).Code.Count - 1
                            'перебираем слова в строках кода
                            If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                    Dim res As String = SeekFunctionInString(cd, line, wordId, newName)
                                    If res = "Entrance" Then
                                        dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                                    ElseIf res.Length > 0 Then
                                        cd(line).Code(wordId).Word = res
                                    End If
                                End If
                            End If
                        Next wordId
                    Next line
                End If
            End If

            If IsNothing(tr.propAfterContent) = False Then
                If RenameFunctionInCode(tr.eventAfterId, oldName, newName) Then
                    Dim cd() As CodeTextBox.CodeDataType = tr.propAfterContent
                    For line As Integer = 0 To cd.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(cd(line).Code) Then Continue For
                        For wordId As Integer = 0 To cd(line).Code.Count - 1
                            'перебираем слова в строках кода
                            If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                    Dim res As String = SeekFunctionInString(cd, line, wordId, newName)
                                    If res = "Entrance" Then
                                        dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                                    ElseIf res.Length > 0 Then
                                        cd(line).Code(wordId).Word = res
                                    End If
                                End If
                            End If
                        Next wordId
                    Next line
                End If
            End If

        Next trId
    End Sub

    ''' <summary>
    ''' Переименовывает функцию Писателя в исполняемой строке или сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="oldName">Старое имя функции</param>
    ''' <param name="newName">Новое имя функции</param>
    ''' <returns>Исполнемую строку с внесенными изменениями</returns>
    Private Function RenameFunctionInExPropery(ByVal strXml As String, ByVal oldName As String, ByVal newName As String) As String
        If String.IsNullOrEmpty(strXml) Then Return strXml

        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return strXml
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wordId As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    Dim res As String = SeekFunctionInString(cd, line, wordId, newName)
                    If res = "Entrance" Then
                        dlgEntrancies.SetSeekPosInfo(line)
                        dlgEntrancies.NewEntrance()
                    ElseIf res.Length > 0 Then
                        cd(line).Code(wordId).Word = res
                        wasFound = True
                    End If
                End If
            Next wordId
        Next line
        If wasFound Then
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                'собираем заново исполняемую строку
                strXml = cd(0).StartingSpaces
                For i As Integer = 0 To cd(0).Code.Count - 1
                    strXml += cd(0).Code(i).Word
                Next i
                strXml += cd(0).Comments
            Else
                'сериализаем код с внесенными изменениями
                strXml = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
            End If
        End If
        Return strXml
    End Function

    ''' <summary>
    ''' Переименовывает функцию Писателя в содержимом события
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="oldName">Старое имя функции</param>
    ''' <param name="newName">Новое имя функции</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function RenameFunctionInCode(ByVal evendId As Integer, ByVal oldName As String, ByVal newName As String) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 1 To exData(line).Code.Count - 1 '1 словом быть не может
                If exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    Dim res As String = SeekFunctionInStringEx(exData, line, wordId, newName)
                    If res = "Entrance" Then
                        wasFound = True
                    ElseIf res.Length > 0 Then
                        exData(line).Code(wordId).Word = res
                        wasFound = True
                    End If
                End If
            Next wordId
        Next line
        Return wasFound
    End Function

    ''' <summary>
    ''' Внутрення функция для функций группы Rename Functions. Определяет является ли слово скрипта формата CodeTextBox.CodeDataType названием функции Писателя.
    ''' Если является, то возвращает слово с измененным именем функции.
    ''' </summary>
    ''' <param name="cd">Скрипт</param>
    ''' <param name="line">Текущая линия в cd</param>
    ''' <param name="wordId">Id проверяемого слова в текущей линии</param>
    ''' <param name="newName">Новое имя функции</param>
    Private Function SeekFunctionInString(ByRef cd() As CodeTextBox.CodeDataType, ByVal line As Integer, ByVal wordId As Integer, ByVal newName As String) As String
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWord(cd, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_USER_FUNCTION Then
                        'Строка точно является именем функции Писателя, стоит параметром в функции
                        Dim w As String = cd(line).Code(wordId).Word
                        Return w.Replace(w.Trim, newName)
                    Else
                        Return "Entrance"
                    End If
                Else
                    Return "Entrance"
                End If
            Else
                Return "Entrance"
            End If
        ElseIf wordId > 0 AndAlso (cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                   cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                   cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) AndAlso _
                               GetLastPropertyReturnElementClassId(cd, wordId - 1, line) = -3 Then
            Dim w As String = cd(line).Code(wordId).Word
            Return w.Replace(w.Trim, newName)
        ElseIf wordId = 1 AndAlso cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION Then
            Dim w As String = cd(line).Code(wordId).Word
            Return w.Replace(w.Trim, newName)
            Return "Entrance"
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Внутрення функция для функций группы Rename Functions. Определяет является ли слово скрипта формата MatewScript.ExecuteDataType названием функции Писателя.
    ''' Если является, то возвращает слово с измененным именем функции.
    ''' </summary>
    ''' <param name="exData">Скрипт</param>
    ''' <param name="line">Текущая линия в cd</param>
    ''' <param name="wordId">Id проверяемого слова в текущей линии</param>
    ''' <param name="newName">Новое имя функции</param>
    Private Function SeekFunctionInStringEx(ByRef exData As List(Of MatewScript.ExecuteDataType), ByVal line As Integer, ByVal wordId As Integer, ByVal newName As String) As String
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWordInExData(exData, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_USER_FUNCTION Then
                        'Строка точно является именем функции Писателя, стоит параметром в функции
                        Dim w As String = exData(line).Code(wordId).Word
                        Return w.Replace(w.Trim, newName)
                    Else
                        Return "Entrance"
                    End If
                Else
                    Return "Entrance"
                End If
            Else
                Return "Entrance"
            End If
        ElseIf wordId > 0 AndAlso (exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                   exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                   exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) AndAlso _
                               GetLastPropertyReturnElementClassIdEx(exData, wordId - 1, line) = -3 Then
            Dim w As String = exData(line).Code(wordId).Word
            Return w.Replace(w.Trim, newName)
        ElseIf wordId = 1 AndAlso exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION Then
            Dim w As String = exData(line).Code(wordId).Word
            Return w.Replace(w.Trim, newName)
        Else
            Return "Entrance"
        End If
        Return ""
    End Function
#End Region

#Region "Check Functions"
    ''' <summary>
    ''' Ищет ссылки на функции Писателя во всех свойствах, функциях, переменных...
    ''' </summary>
    ''' <param name="remName">Имя элемента (в кавычках)</param>
    ''' <param name="afterRemoving">Запускается ли процедура после удаления элемента или же это просто поиск</param>
    Public Sub CheckFunctionInStruct(remName As String, Optional ByVal afterRemoving As Boolean = True)
        If afterRemoving Then
            dlgEntrancies.BeginNewEntrancies("Найдены ссылки на удаленные функции, а также идентичные ее имени строки, значение которых вычислить не удалось. Провертьте их самостоятельно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        Else
            dlgEntrancies.BeginNewEntrancies("Результаты поиска функции " + remName + ". Включены также те результаты, значение которых определить невозможно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        End If

        'Прочесывание структуры mainClass на предмет наличия свойств типа ELEMENT нашего класса и имеющим значение = remName.
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств каждого класса
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_FUNCTION Then
                        'тип = функция Писателя
                        Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                        'Свойство по умолчанию
                        If String.Compare(p.Value, remName, True) = 0 Then
                            dlgEntrancies.NewEntrance(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                        End If
                        If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                            'Свойства 2 уровня
                            For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                If String.Compare(ch.Value, remName, True) = 0 Then
                                    dlgEntrancies.NewEntrance(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                                End If
                                If IsNothing(ch.ThirdLevelProperties) = False Then
                                    'Свойства 3 уровня
                                    For child3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                        If String.Compare(ch.ThirdLevelProperties(child3Id), remName, True) = 0 Then
                                            dlgEntrancies.NewEntrance(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                                        End If
                                    Next child3Id
                                End If
                            Next child2Id
                        End If
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For
                        If pp.returnType = MatewScript.ReturnFunctionEnum.RETURN_FUNCTION Then
                            'тип = функция Писателя
                            If String.Compare(ch.Value, remName, True) = 0 Then
                                dlgEntrancies.NewEntrance(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                            End If
                        End If
                    Next pId
                Next aId
            Next i
        End If

        CheckFunctionInScripts(remName)
        If dlgEntrancies.hasEntrancies Then
            dlgEntrancies.Show()
        End If
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные функции Писателя во всех скриптах
    ''' </summary>
    ''' <param name="remName">Имя элемента (в кавычках)</param>
    Private Sub CheckFunctionInScripts(ByVal remName As String)
        'Изменения в коде
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    If p.eventId > 0 AndAlso RemoveFunctionInCode(p.eventId, remName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        RemoveFunctionInExPropery(p.Value, remName)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RemoveFunctionInCode(ch.eventId, remName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                            RemoveFunctionInExPropery(ch.Value, remName)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso RemoveFunctionInCode(ch.ThirdLevelEventId(child3Id), remName) Then
                                'Значение - скрипт
                                dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                                RemoveFunctionInExPropery(ch.ThirdLevelProperties(child3Id), remName)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If RemoveFunctionInCode(p.eventId, remName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        RemoveFunctionInExPropery(p.Value, remName)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        If ch.eventId > 0 AndAlso RemoveFunctionInCode(ch.eventId, remName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                            RemoveFunctionInExPropery(ch.Value, remName)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        RemoveFunctionInFunctionHash(remName)
        'В глобальных переменных
        RemoveFunctionInVariables(remName)
        'В событиях изменения свойства
        RemoveFunctionInTracking(remName)
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные функции Писателя в глобальных переменных
    ''' </summary>
    ''' <param name="remName">Имя элемента</param>
    Private Sub RemoveFunctionInVariables(ByVal remName As String)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return
        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim cd() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    cd = .codeBox.CodeData
                End With
                If IsNothing(cd) Then Continue For
                'В элементе массива - исполняемый код
                For line As Integer = 0 To cd.Count - 1
                    If IsNothing(cd(line).Code) Then Continue For

                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekFunctionInStringRemoving(cd, line, wordId) Then
                                dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), GetWordPosInLine(cd(line), wordId))
                            End If
                        End If
                    Next wordId
                Next line
            Next arrId
        Next varId
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные функции Писателя в функциях писателя functionsHash
    ''' </summary>
    ''' <param name="remName">Имя элемента</param>
    Private Sub RemoveFunctionInFunctionHash(ByVal remName As String)
        If IsNothing(mScript.functionsHash) Then Return
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            'В функции хранится код в исполняемом виде - в ValueExecuteDt, и в виде для редактирования - в ValueDt (при этом второе может быть пустым, если функция объявлена из кода)
            'Первое быть пустым не должно
            If IsNothing(func.ValueDt) = False Then
                Dim cd() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekFunctionInStringRemoving(cd, line, wordId) Then
                                dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), GetWordPosInLine(cd(line), wordId))
                            End If
                        End If
                    Next wordId
                Next line
            End If
        Next funcId
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные функции Писателя в функциях писателя functionsHash
    ''' </summary>
    ''' <param name="remName">Имя элемента</param>
    Private Sub RemoveFunctionInTracking(ByVal remName As String)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return
        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
            Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            If IsNothing(tr.propBeforeContent) = False Then
                Dim cd() As CodeTextBox.CodeDataType = tr.propBeforeContent
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If RemoveFunctionInCode(tr.eventBeforeId, remName) Then
                                If SeekFunctionInStringRemoving(cd, line, wordId) Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), _
                                                              GetWordPosInLine(cd(line), wordId), frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                                End If
                            End If
                        End If
                    Next wordId
                Next line
            End If

            If IsNothing(tr.propAfterContent) = False Then
                Dim cd() As CodeTextBox.CodeDataType = tr.propAfterContent
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If RemoveFunctionInCode(tr.eventAfterId, remName) Then
                                If SeekFunctionInStringRemoving(cd, line, wordId) Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), _
                                                              GetWordPosInLine(cd(line), wordId), frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                                End If
                            End If
                        End If
                    Next wordId
                Next line
            End If

        Next trId
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные функции Писателя в исполняемой строке или сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="remName">Старое имя элемента</param>
    Private Sub RemoveFunctionInExPropery(ByVal strXml As String, ByVal remName As String)
        If String.IsNullOrEmpty(strXml) Then Return

        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wordId As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    If SeekFunctionInStringRemoving(cd, line, wordId) Then
                        dlgEntrancies.SetSeekPosInfo(line, mScript.PrepareStringToPrint(remName, Nothing, False), GetWordPosInLine(cd(line), wordId))
                        dlgEntrancies.NewEntrance()
                    End If
                End If
            Next wordId
        Next line
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 2 уровня в содержимом события
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="remName">имя элемента</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function RemoveFunctionInCode(ByVal evendId As Integer, ByVal remName As String) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 1 To exData(line).Code.Count - 1 '1 словом быть не может
                If exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    If SeekFunctionInStringRemovingEx(exData, line, wordId) = True Then wasFound = True
                End If
            Next wordId
        Next line
        Return wasFound
    End Function

    ''' <summary>
    ''' Внутрення функция группы Remove Function. Определяет может ли слово с индексом wordId являться именем функции Писателя
    ''' </summary>
    ''' <param name="cd">Код в формате CodeTextBox.CodeDataType</param>
    ''' <param name="line">Линим в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>True если может</returns>
    Private Function SeekFunctionInStringRemoving(ByRef cd() As CodeTextBox.CodeDataType, ByVal line As Integer, wordId As Integer) As Boolean
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWord(cd, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) AndAlso IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count Then
                    If f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                        'Это не функция точно, можно пропустить
                        Return False
                    Else
                        Return True 'надо проверить что это
                    End If
                Else
                    Return True 'надо проверить что это
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    Return False 'Внутри свойства - не мжет быть функцией
                Else
                    Return True 'надо проверить что это
                End If
            Else
                Return True 'надо проверить что это
            End If
        ElseIf wordId > 0 AndAlso (cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                   cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                   cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) Then
            Dim res As Integer = GetLastPropertyReturnElementClassId(cd, wordId - 1, line)
            If res > -1 Then
                'Это другой класс
                Return False
            Else
                Return True
            End If
        Else
            Return True 'надо проверить что это
        End If
        Return False
    End Function

    ''' <summary>
    ''' Внутрення функция группы Remove Function. Определяет может ли слово с индексом wordId являться именем функции Писателя
    ''' </summary>
    ''' <param name="exData">Код в формате MatewScript.ExecuteDataType</param>
    ''' <param name="line">Линим в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>True если может</returns>
    Private Function SeekFunctionInStringRemovingEx(ByRef exData As List(Of MatewScript.ExecuteDataType), ByVal line As Integer, wordId As Integer) As Boolean
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWordInExData(exData, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) AndAlso IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count Then
                    If f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                        'Это не функция точно, можно пропустить
                        Return False
                    Else
                        Return True 'надо проверить что это
                    End If
                Else
                    Return True 'надо проверить что это
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    Return False 'Внутри свойства - не мжет быть функцией
                Else
                    Return True 'надо проверить что это
                End If
            Else
                Return True 'надо проверить что это
            End If
        ElseIf wordId > 0 AndAlso (exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                   exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                   exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) Then
            Dim res As Integer = GetLastPropertyReturnElementClassIdEx(exData, wordId - 1, line)
            If res > -1 Then
                'Это другой класс
                Return False
            Else
                Return True
            End If
        Else
            Return True 'надо проверить что это
        End If
        Return False
    End Function
#End Region

#Region "Rename Child2"
    ''' <summary>
    ''' Переименовывает элемент 2 уровня во всех свойствах, функциях, перменных со скриптами и т. д.
    ''' </summary>
    ''' <param name="editClassId">Id класса для замены</param>
    ''' <param name="oldName">Старое имя элемента (в кавычках)</param>
    ''' <param name="newName">Новое имя элемента (в кавычках)</param>
    Public Sub RenameChild2InStruct(ByVal editClassId As Integer, oldName As String, newName As String)
        'Dim newShort As String
        Dim className As String = mScript.mainClass(editClassId).Names.Last
        'Dim newLong As String

        'Прочесывание структуры mainClass на предмет наличия свойств типа ELEMENT нашего класса и имеющим значение = старое имя. Если найдем - переименовываем
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств каждого класса
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso p.returnArray(0) = className Then
                        'тип = элемент нужного класса
                        Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                        'Свойство по умолчанию
                        If String.Compare(p.Value, oldName, True) = 0 Then
                            p.Value = newName
                        End If
                        If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                            'Свойства 2 уровня
                            For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                If String.Compare(ch.Value, oldName, True) = 0 Then
                                    ch.Value = newName
                                End If
                                If IsNothing(ch.ThirdLevelProperties) = False Then
                                    'Свойства 3 уровня
                                    For child3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                        If String.Compare(ch.ThirdLevelProperties(child3Id), oldName, True) = 0 Then
                                            ch.ThirdLevelProperties(child3Id) = newName
                                        End If
                                    Next child3Id
                                End If
                            Next child2Id
                        End If
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For
                        If pp.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(pp.returnArray) = False AndAlso pp.returnArray.Count = 1 AndAlso pp.returnArray(0) = className Then
                            'тип = элемент нужного класса
                            If String.Compare(ch.Value, oldName, True) = 0 Then
                                ch.Value = newName
                            End If
                        End If
                    Next pId
                Next aId
            Next i
        End If

        RenameChild2InScripts(editClassId, oldName, newName)
    End Sub

    ''' <summary>
    ''' Переименовывает элемент 2 уровня во всех скриптах
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="oldName">Старое имя элемента (в кавычках)</param>
    ''' <param name="newName">Новое имя элемента (в кавычках)</param>
    Private Sub RenameChild2InScripts(ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String)
        'Изменения в коде
        dlgEntrancies.BeginNewEntrancies("Найдены строки, идентичные старому имени элемента, значение которых вычислить не удалось. Провертьте их самостоятельно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    If p.eventId > 0 AndAlso RenameChild2InCode(p.eventId, editClassId, oldName, newName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        p.Value = RenameChild2InExPropery(p.Value, editClassId, oldName, newName)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RenameChild2InCode(ch.eventId, editClassId, oldName, newName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                            ch.Value = RenameChild2InExPropery(ch.Value, editClassId, oldName, newName)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso RenameChild2InCode(ch.ThirdLevelEventId(child3Id), editClassId, oldName, newName) Then
                                'Значение - скрипт
                                dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                                ch.ThirdLevelProperties(child3Id) = RenameChild2InExPropery(ch.ThirdLevelProperties(child3Id), editClassId, oldName, newName)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If RenameChild2InCode(p.eventId, editClassId, oldName, newName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        p.Value = RenameChild2InExPropery(p.Value, editClassId, oldName, newName)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        If ch.eventId > 0 AndAlso RenameChild2InCode(ch.eventId, editClassId, oldName, newName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                            ch.Value = RenameChild2InExPropery(ch.Value, editClassId, oldName, newName)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        RenameChild2InFunctionHash(editClassId, oldName, newName)
        'В глобальных переменных
        RenameChild2InVariables(editClassId, oldName, newName)
        'В событиях изменения свойтсва
        RenameChild2InTracking(editClassId, oldName, newName)

        If dlgEntrancies.hasEntrancies Then
            dlgEntrancies.Show()
        End If
    End Sub

    ''' <summary>
    ''' Производит замены имени элемента 2 порядка в глобальных переменных
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    Private Sub RenameChild2InVariables(ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return
        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            Dim wasFoundGlobal As Boolean = False 'были ли изменения хоть в одном элементе массива переменной
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim cd() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    cd = .codeBox.CodeData
                End With
                If IsNothing(cd) Then Continue For
                'В элементе массива - исполняемый код
                Dim wasFound As Boolean = False 'была ли замена в данном элементе массива
                For line As Integer = 0 To cd.Count - 1
                    If IsNothing(cd(line).Code) Then Continue For

                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            Dim res As String = SeekChild2NameInString(cd, editClassId, newName, line, wordId)
                            If res = "Entrance" Then
                                dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line)
                                wasFound = True
                            ElseIf res.Length > 0 Then
                                cd(line).Code(wordId).Word = res
                                wasFound = True
                            End If
                        End If
                    Next wordId
                Next line
                If wasFound Then
                    wasFoundGlobal = True
                    If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                        'Исполняемая строка. Собираем измененную строку
                        Dim w As String = cd(0).StartingSpaces
                        For i As Integer = 0 To cd(0).Code.Count - 1
                            w += cd(0).Code(i).Word
                        Next i
                        w += cd(0).Comments
                        arrValues(arrId) = WrapString(w)
                    Else
                        'Скрипт или длинный текст. Собираем с внесенными изменениями
                        arrValues(arrId) = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
                    End If
                End If
            Next arrId
            If wasFoundGlobal Then
                'Сохраняем результат
                Dim var As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value
                var.arrValues = arrValues
            End If
        Next varId
    End Sub

    ''' <summary>
    ''' Производит замены имени элемента 2 порядка в функциях писателя functionsHash
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    Private Sub RenameChild2InFunctionHash(ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String)
        If IsNothing(mScript.functionsHash) Then Return
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            'В функции хранится код в исполняемом виде - в ValueExecuteDt, и в виде для редактирования - в ValueDt (при этом второе может быть пустым, если функция объявлена из кода)
            'Первое быть пустым не должно
            Dim wasFound As Boolean = False
            If IsNothing(func.ValueExecuteDt) = False Then
                Dim edt As List(Of MatewScript.ExecuteDataType) = func.ValueExecuteDt
                For line As Integer = 0 To edt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(edt(line).Code) Then Continue For
                    For wordId As Integer = 1 To edt(line).Code.Count - 1 '1 словом быть не может
                        If edt(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(edt(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            Dim res As String = SeekChild2NameInStringEx(edt, editClassId, newName, line, wordId)
                            If res = "Entrance" Then
                                wasFound = True
                            ElseIf res.Length > 0 Then
                                edt(line).Code(wordId).Word = res
                                wasFound = True
                            End If
                        End If
                    Next wordId
                Next line
            End If

            If Not wasFound Then Continue For 'если не найдено в ValueExecuteDt, то здесь тоже не будет
            If IsNothing(func.ValueDt) = False Then
                Dim cd() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            Dim res As String = SeekChild2NameInString(cd, editClassId, newName, line, wordId)
                            If res = "Entrance" Then
                                dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line)
                                wasFound = True
                            ElseIf res.Length > 0 Then
                                cd(line).Code(wordId).Word = res
                                wasFound = True
                            End If
                        End If
                    Next wordId
                Next line
            End If
        Next funcId
    End Sub

    Private Sub RenameChild2InTracking(ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return
        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
            Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            If IsNothing(tr.propBeforeContent) = False Then
                If RenameChild2InCode(tr.eventBeforeId, editClassId, oldName, newName) Then
                    Dim cd() As CodeTextBox.CodeDataType = tr.propBeforeContent
                    For line As Integer = 0 To cd.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(cd(line).Code) Then Continue For
                        For wordId As Integer = 0 To cd(line).Code.Count - 1
                            'перибираем слова в строках кода
                            If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                Dim res As String = SeekChild2NameInString(cd, editClassId, newName, line, wordId)
                                If res = "Entrance" Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                                ElseIf res.Length > 0 Then
                                    cd(line).Code(wordId).Word = res
                                End If
                            End If
                        Next wordId
                    Next line
                End If
            End If

            If IsNothing(tr.propAfterContent) = False Then
                If RenameChild2InCode(tr.eventAfterId, editClassId, oldName, newName) Then
                    Dim cd() As CodeTextBox.CodeDataType = tr.propAfterContent
                    For line As Integer = 0 To cd.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(cd(line).Code) Then Continue For
                        For wordId As Integer = 0 To cd(line).Code.Count - 1
                            'перибираем слова в строках кода
                            If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                Dim res As String = SeekChild2NameInString(cd, editClassId, newName, line, wordId)
                                If res = "Entrance" Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                                ElseIf res.Length > 0 Then
                                    cd(line).Code(wordId).Word = res
                                End If
                            End If
                        Next wordId
                    Next line
                End If
            End If
        Next trId
    End Sub

    ''' <summary>
    ''' Производит замены имени элемента 2 порядка в исполняемой строке или сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="editClassId">класс, в котором меняется имя</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    ''' <returns>Исполнемую строку с внесенными изменениями</returns>
    Private Function RenameChild2InExPropery(ByVal strXml As String, ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String) As String
        If String.IsNullOrEmpty(strXml) Then Return strXml

        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return strXml
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wordId As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    Dim res As String = SeekChild2NameInString(cd, editClassId, newName, line, wordId)
                    If res = "Entrance" Then
                        dlgEntrancies.SetSeekPosInfo(line)
                        dlgEntrancies.NewEntrance()
                        wasFound = True
                    ElseIf res.Length > 0 Then
                        cd(line).Code(wordId).Word = res
                        wasFound = True
                    End If
                End If
            Next wordId
        Next line
        If wasFound Then
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                'собираем заново исполняемую строку
                strXml = cd(0).StartingSpaces
                For i As Integer = 0 To cd(0).Code.Count - 1
                    strXml += cd(0).Code(i).Word
                Next i
                strXml += cd(0).Comments
            Else
                'сериализаем код с внесенными изменениями
                strXml = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
            End If
        End If
        Return strXml
    End Function

    ''' <summary>
    ''' Переименовывает элементы 2 уровня в содержимом события
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="editClassId">Id класса</param>
    ''' <param name="oldName">Старое имя</param>
    ''' <param name="newName">Новое имя</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function RenameChild2InCode(ByVal evendId As Integer, ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line)) OrElse IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 1 To exData(line).Code.Count - 1 '1 словом быть не может
                If exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    Dim res As String = SeekChild2NameInStringEx(exData, editClassId, newName, line, wordId)
                    If res = "Entrance" Then
                        wasFound = True
                    ElseIf res.Length > 0 Then
                        exData(line).Code(wordId).Word = res
                        wasFound = True
                    End If
                End If
            Next wordId
        Next line
        Return wasFound
    End Function

    ''' <summary>
    ''' Внутрення функция группы Rename Child2. Определяет может ли слово с индексом wordId являться именем элемента 2 уровня класса editClassId
    ''' </summary>
    ''' <param name="cd">Код в формате CodeTextBox.CodeDataType</param>
    ''' <param name="editClassId">Класс элемента 2 уровня</param>
    ''' <param name="line">Линия в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>Строку с измененным именем, Entrance (если нужна проверка со стороны Писателя) или пустая строка, когда ложная тревога</returns>
    Private Function SeekChild2NameInString(ByRef cd() As CodeTextBox.CodeDataType, ByVal editClassId As Integer, ByVal newName As String, ByVal line As Integer, wordId As Integer) As String
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim className As String = mScript.mainClass(editClassId).Names.Last
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWord(cd, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT _
                        AndAlso IsNothing(f.params(paramNumber - 1).EnumValues) = False AndAlso f.params(paramNumber - 1).EnumValues.Count = 1 Then
                        If f.params(paramNumber - 1).EnumValues(0) = className Then
                            'Строка точно является именем элемента 2 уровня, стоит параметром в функции
                            Dim w As String = cd(line).Code(wordId).Word
                            Return w.Replace(w.Trim, newName)
                        End If
                    Else
                        Return "Entrance"
                    End If
                Else
                    Return "Entrance"
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    If elClassId = editClassId AndAlso paramNumber = 1 Then
                        'Строка точно является именем элемента 2 уровня, стоит первым параметром в свойстве
                        Dim w As String = cd(line).Code(wordId).Word
                        Return w.Replace(w.Trim, newName)
                    End If
                Else
                    Return "Entrance"
                End If
            Else
                Return "Entrance"
            End If
        ElseIf wordId > 0 AndAlso (cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                   cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                   cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) AndAlso _
                               GetLastPropertyReturnElementClassId(cd, wordId - 1, line) = editClassId Then
            Dim w As String = cd(line).Code(wordId).Word
            Return w.Replace(w.Trim, newName)
        ElseIf cd(line).Code.Count > 3 AndAlso cd(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT AndAlso cd(line).Code(1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso wordId = 3 Then
            'Event 'EventName', 'child2Name'
            Dim cId As Integer = -1, elementName As String = ""
            If GetClassIdAndElementNameByString(cd(line).Code(1).Word, cId, elementName, False) = False Then Return "Entrance"
            If cId = editClassId Then
                Dim w As String = cd(line).Code(wordId).Word
                Return w.Replace(w.Trim, newName)
            Else
                Return ""
            End If
        Else
            Return "Entrance"
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Внутрення функция группы Rename Child2. Определяет может ли слово с индексом wordId являться именем элемента 2 уровня класса editClassId
    ''' </summary>
    ''' <param name="exData">Код в формате MatewScript.ExecuteDataType</param>
    ''' <param name="editClassId">Класс элемента 2 уровня</param>
    ''' <param name="line">Линия в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>Строку с измененным именем, Entrance (если нужна проверка со стороны Писателя) или пустая строка, когда ложная тревога</returns>
    Private Function SeekChild2NameInStringEx(ByRef exData As List(Of MatewScript.ExecuteDataType), ByVal editClassId As Integer, ByVal newName As String, ByVal line As Integer, wordId As Integer) As String
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim className As String = mScript.mainClass(editClassId).Names.Last
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWordInExData(exData, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT _
                        AndAlso IsNothing(f.params(paramNumber - 1).EnumValues) = False AndAlso f.params(paramNumber - 1).EnumValues.Count = 1 Then
                        If f.params(paramNumber - 1).EnumValues(0) = className Then
                            'Строка точно является именем элемента 2 уровня, стоит параметром в функции
                            Dim w As String = exData(line).Code(wordId).Word
                            Return w.Replace(w.Trim, newName)
                        End If
                    Else
                        Return "Entrance"
                    End If
                Else
                    Return "Entrance"
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    If elClassId = editClassId AndAlso paramNumber = 1 Then
                        'Строка точно является именем элемента 2 уровня, стоит первым параметром в свойстве
                        Dim w As String = exData(line).Code(wordId).Word
                        Return w.Replace(w.Trim, newName)
                    End If
                Else
                    Return "Entrance"
                End If
            Else
                Return "Entrance"
            End If
        ElseIf wordId > 0 AndAlso (exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                   exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                   exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) AndAlso _
                               GetLastPropertyReturnElementClassIdEx(exData, wordId - 1, line) = editClassId Then
            Dim w As String = exData(line).Code(wordId).Word
            Return w.Replace(w.Trim, newName)
        ElseIf exData(line).Code.Count > 3 AndAlso exData(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT AndAlso exData(line).Code(1).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso wordId = 3 Then
            'Event 'EventName', 'child2Name'
            Dim cId As Integer = -1, elementName As String = ""
            If GetClassIdAndElementNameByString(exData(line).Code(1).Word, cId, elementName, False) = False Then Return "Entrance"
            If cId = editClassId Then
                Dim w As String = exData(line).Code(wordId).Word
                Return w.Replace(w.Trim, newName)
            Else
                Return ""
            End If
        Else
            Return "Entrance"
        End If
        Return ""
    End Function
#End Region

#Region "Check Child2"
    ''' <summary>
    ''' Ищет ссылки на элементы 2 уровня во всех свойствах, функциях, переменных...
    ''' </summary>
    ''' <param name="remClassId">Id класса удаляемого элемента</param>
    ''' <param name="remName">Имя элемента (в кавычках)</param>
    ''' <param name="afterRemoving">Запускается ли процедура после удаления элемента или же это просто поиск</param>
    Public Sub CheckChild2InStruct(ByVal remClassId As Integer, remName As String, Optional ByVal afterRemoving As Boolean = True)
        Dim className As String = mScript.mainClass(remClassId).Names.Last
        If afterRemoving Then
            dlgEntrancies.BeginNewEntrancies("Найдены ссылки на удаленный элемент, а также идентичные его имени строки, значение которым вычислить не удалось. Провертьте их самостоятельно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        Else
            dlgEntrancies.BeginNewEntrancies("Результаты поиска элемента " + remName + " класса " + mScript.mainClass(remClassId).Names.Last + ". Включены также те результаты, значение которых определить невозможно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        End If

        'Прочесывание структуры mainClass на предмет наличия свойств типа ELEMENT нашего класса и имеющим значение = remName.
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств каждого класса
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso p.returnArray(0) = className Then
                        'тип = элемент нужного класса
                        Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                        'Свойство по умолчанию
                        If String.Compare(p.Value, remName, True) = 0 Then
                            dlgEntrancies.NewEntrance(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                        End If
                        If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                            'Свойства 2 уровня
                            For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                If String.Compare(ch.Value, remName, True) = 0 Then
                                    dlgEntrancies.NewEntrance(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                                End If
                                If IsNothing(ch.ThirdLevelProperties) = False Then
                                    'Свойства 3 уровня
                                    For child3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                        If String.Compare(ch.ThirdLevelProperties(child3Id), remName, True) = 0 Then
                                            dlgEntrancies.NewEntrance(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                                        End If
                                    Next child3Id
                                End If
                            Next child2Id
                        End If
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For
                        If pp.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(pp.returnArray) = False AndAlso pp.returnArray.Count = 1 AndAlso pp.returnArray(0) = className Then
                            'тип = элемент нужного класса
                            If String.Compare(ch.Value, remName, True) = 0 Then
                                dlgEntrancies.NewEntrance(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                            End If
                        End If
                    Next pId
                Next aId
            Next i
        End If

        CheckChild2InScripts(remClassId, remName)
        If dlgEntrancies.hasEntrancies Then
            dlgEntrancies.Show()
        End If
    End Sub


    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 2 уровня во всех скриптах
    ''' </summary>
    ''' <param name="remClassId">Класс удаляемого элеента</param>
    ''' <param name="remName">Имя элемента (в кавычках)</param>
    Private Sub CheckChild2InScripts(ByVal remClassId As Integer, ByVal remName As String)
        'Изменения в коде
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    If p.eventId > 0 AndAlso RemoveChild2InCode(p.eventId, remClassId, remName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        RemoveChild2InExPropery(p.Value, remClassId, remName)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RemoveChild2InCode(ch.eventId, remClassId, remName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                            RemoveChild2InExPropery(ch.Value, remClassId, remName)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso RemoveChild2InCode(ch.ThirdLevelEventId(child3Id), remClassId, remName) Then
                                'Значение - скрипт
                                dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                                RemoveChild2InExPropery(ch.ThirdLevelProperties(child3Id), remClassId, remName)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If RemoveChild2InCode(p.eventId, remClassId, remName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        RemoveChild2InExPropery(p.Value, remClassId, remName)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        If ch.eventId > 0 AndAlso RemoveChild2InCode(ch.eventId, remClassId, remName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                            RemoveChild2InExPropery(ch.Value, remClassId, remName)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        RemoveChild2InFunctionHash(remClassId, remName)
        'В глобальных переменных
        RemoveChild2InVariables(remClassId, remName)
        'В событиях изменения свойства
        RemoveChild2InTracking(remClassId, remName)
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 2 порядка в глобальных переменных
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    ''' <param name="remName">Имя элемента</param>
    Private Sub RemoveChild2InVariables(ByVal remClassId As Integer, ByVal remName As String)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return
        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim cd() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    cd = .codeBox.CodeData
                End With
                If IsNothing(cd) Then Continue For
                'В элементе массива - исполняемый код
                For line As Integer = 0 To cd.Count - 1
                    If IsNothing(cd(line).Code) Then Continue For

                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekChild2NameInStringRemoving(cd, remClassId, line, wordId) Then
                                dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), _
                                                          GetWordPosInLine(cd(line), wordId))
                            End If
                        End If
                    Next wordId
                Next line
            Next arrId
        Next varId
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 2 порядка в функциях писателя functionsHash
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    ''' <param name="remName">Имя элемента</param>
    Private Sub RemoveChild2InFunctionHash(ByVal remClassId As Integer, ByVal remName As String)
        If IsNothing(mScript.functionsHash) Then Return
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            'В функции хранится код в исполняемом виде - в ValueExecuteDt, и в виде для редактирования - в ValueDt (при этом второе может быть пустым, если функция объявлена из кода)
            'Первое быть пустым не должно
            If IsNothing(func.ValueDt) = False Then
                Dim cd() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekChild2NameInStringRemoving(cd, remClassId, line, wordId) Then
                                dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), _
                                                          GetWordPosInLine(cd(line), wordId))
                            End If
                        End If
                    Next wordId
                Next line
            End If
        Next funcId
    End Sub

    Private Sub RemoveChild2InTracking(ByVal remClassId As Integer, ByVal remName As String)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return
        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
            Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            If IsNothing(tr.propBeforeContent) = False Then
                Dim cd() As CodeTextBox.CodeDataType = tr.propBeforeContent
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekChild2NameInStringRemoving(cd, remClassId, line, wordId) Then
                                dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), _
                                                          GetWordPosInLine(cd(line), wordId), frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                            End If
                        End If
                    Next wordId
                Next line
            End If

            If IsNothing(tr.propAfterContent) = False Then
                Dim cd() As CodeTextBox.CodeDataType = tr.propAfterContent
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekChild2NameInStringRemoving(cd, remClassId, line, wordId) Then
                                dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), _
                                                          GetWordPosInLine(cd(line), wordId), frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                            End If
                        End If
                    Next wordId
                Next line
            End If
        Next trId
    End Sub


    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 2 уровня в исполняемой строке или сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="remClassId">класс, в котором меняется имя</param>
    ''' <param name="remName">Старое имя элемента</param>
    Private Sub RemoveChild2InExPropery(ByVal strXml As String, ByVal remClassId As Integer, ByVal remName As String)
        If String.IsNullOrEmpty(strXml) Then Return
        Dim className As String = mScript.mainClass(remClassId).Names.Last

        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wordId As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    If SeekChild2NameInStringRemoving(cd, remClassId, line, wordId) Then
                        dlgEntrancies.SetSeekPosInfo(line, mScript.PrepareStringToPrint(remName, Nothing, False), GetWordPosInLine(cd(line), wordId))
                        dlgEntrancies.NewEntrance()
                    End If
                End If
            Next wordId
        Next line
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 2 уровня в содержимом события
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="remClassId">Id класса</param>
    ''' <param name="remName">имя элемента</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function RemoveChild2InCode(ByVal evendId As Integer, ByVal remClassId As Integer, ByVal remName As String) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 1 To exData(line).Code.Count - 1 '1 словом быть не может
                If exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    If SeekChild2NameInStringRemovingEx(exData, remClassId, line, wordId) = True Then wasFound = True
                End If
            Next wordId
        Next line
        Return wasFound
    End Function

    ''' <summary>
    ''' Внутрення функция группы Remove Child2. Определяет может ли слово с индексом wordId являться именем элемента 2 уровня класса remClassId
    ''' </summary>
    ''' <param name="exData">Код в формате MatewScript.ExecuteDataType</param>
    ''' <param name="remClassId">Класс элемента 2 уровня</param>
    ''' <param name="line">Линим в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>True если может</returns>
    Private Function SeekChild2NameInStringRemovingEx(ByRef exData As List(Of MatewScript.ExecuteDataType), ByVal remClassId As Integer, ByVal line As Integer, wordId As Integer) As Boolean
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim className As String = mScript.mainClass(remClassId).Names.Last
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWordInExData(exData, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT _
                        AndAlso IsNothing(f.params(paramNumber - 1).EnumValues) = False AndAlso f.params(paramNumber - 1).EnumValues.Count = 1 Then
                        If f.params(paramNumber - 1).EnumValues(0) = className Then
                            'Строка точно является именем элемента 2 уровня, стоит параметром в функции
                            Return True
                        End If
                    Else
                        Return True 'надо проверить что это
                    End If
                Else
                    Return True 'надо проверить что это
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    If elClassId = remClassId AndAlso paramNumber = 1 Then
                        'Строка точно является именем элемента 2 уровня, стоит первым параметром в свойстве
                        Return True
                    End If
                Else
                    Return True 'надо проверить что это
                End If
            Else
                Return True 'надо проверить что это
            End If
        ElseIf wordId > 0 AndAlso (exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                   exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                   exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) Then
            Dim res As Integer = GetLastPropertyReturnElementClassIdEx(exData, wordId - 1, line)
            If res > -1 AndAlso res <> remClassId Then
                'Это другой класс
            Else
                Return True
            End If
        Else
            Return True 'надо проверить что это
        End If
        Return False
    End Function

    ''' <summary>
    ''' Внутрення функция группы Remove Child2. Определяет может ли слово с индексом wordId являться именем элемента 2 уровня класса remClassId
    ''' </summary>
    ''' <param name="cd">Код в формате CodeTextBox.CodeDataType</param>
    ''' <param name="remClassId">Класс элемента 2 уровня</param>
    ''' <param name="line">Линим в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>True если может</returns>
    Private Function SeekChild2NameInStringRemoving(ByRef cd() As CodeTextBox.CodeDataType, ByVal remClassId As Integer, ByVal line As Integer, wordId As Integer) As Boolean
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim className As String = mScript.mainClass(remClassId).Names.Last
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWord(cd, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT _
                        AndAlso IsNothing(f.params(paramNumber - 1).EnumValues) = False AndAlso f.params(paramNumber - 1).EnumValues.Count = 1 Then
                        If f.params(paramNumber - 1).EnumValues(0) = className Then
                            'Строка точно является именем элемента 2 уровня, стоит параметром в функции
                            Return True
                        End If
                    Else
                        Return True 'надо проверить что это
                    End If
                Else
                    Return True 'надо проверить что это
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    If elClassId = remClassId AndAlso paramNumber = 1 Then
                        'Строка точно является именем элемента 2 уровня, стоит первым параметром в свойстве
                        Return True
                    End If
                Else
                    Return True 'надо проверить что это
                End If
            Else
                Return True 'надо проверить что это
            End If
        ElseIf wordId > 0 AndAlso (cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                   cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                   cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) Then
            Dim res As Integer = GetLastPropertyReturnElementClassId(cd, wordId - 1, line)
            If res > -1 AndAlso res <> remClassId Then
                'Это другой класс
            Else
                Return True
            End If
        Else
            Return True 'надо проверить что это
        End If
        Return False
    End Function
#End Region

#Region "Rename Actions"
    ''' <summary>
    ''' Переименовывает действия во всех свойствах и скриптах текущей локации
    ''' </summary>
    ''' <param name="oldName">Старое имя элемента (в кавычках)</param>
    ''' <param name="newName">Новое имя элемента (в кавычках)</param>
    Public Sub RenameActionInStruct(oldName As String, newName As String)
        Dim editClassId As String = mScript.mainClassHash("A")
        Dim className As String = mScript.mainClass(editClassId).Names.Last
        'Dim newLong As String

        'Прочесывание классов Loc & Act структуры mainClass (только 2 уровень) на предмет наличия свойств типа ELEMENT нашего класса и имеющим значение = старое имя. Если найдем - переименовываем
        For i As Integer = 0 To 1
            Dim classId As Integer = IIf(i = 0, mScript.mainClassHash("L"), editClassId)

            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств каждого класса
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso p.returnArray(0) = className Then
                        'тип = элемент нужного класса
                        Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                        If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                            'Свойства 2 уровня
                            For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                If String.Compare(ch.Value, oldName, True) = 0 Then
                                    ch.Value = newName
                                End If
                            Next child2Id
                        End If
                    End If
                Next pId
            End If
        Next i

        RenameActionInScripts(editClassId, oldName, newName)
    End Sub

    ''' <summary>
    ''' Переименовывает действия во всех скриптах текущей локации
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="oldName">Старое имя элемента (в кавычках)</param>
    ''' <param name="newName">Новое имя элемента (в кавычках)</param>
    Private Sub RenameActionInScripts(ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String)
        'Изменения в коде
        dlgEntrancies.BeginNewEntrancies("Найдены строки, идентичные старому имени действия, значение которых вычислить не удалось. Провертьте их самостоятельно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        For i As Integer = 0 To 1
            Dim classId As Integer = IIf(i = 0, mScript.mainClassHash("L"), editClassId)
            Dim parentId As String = ""
            If classId = editClassId Then
                If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False Then
                    parentId = cPanelManager.ActivePanel.GetChild2Id
                ElseIf currentClassName = "A" Then
                    parentId = currentParentId
                End If
            End If
            'Перебираем классы L & A (только 2 уровень)
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RenameActionInCode(ch.eventId, editClassId, oldName, newName) Then
                            'Значение - скрипт                            
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, parentId)
                            ch.Value = RenameActionInExPropery(ch.Value, editClassId, oldName, newName)
                        End If
                    Next child2Id
                Next pId
            End If
        Next i

        If dlgEntrancies.hasEntrancies Then
            dlgEntrancies.Show()
        End If
    End Sub

    ''' <summary>
    ''' Производит замены имени действия в исполняемой строке или сериализованном коде/длинном тексте текущей локации
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="editClassId">класс, в котором меняется имя</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    ''' <returns>Исполнемую строку с внесенными изменениями</returns>
    Private Function RenameActionInExPropery(ByVal strXml As String, ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String) As String
        If String.IsNullOrEmpty(strXml) Then Return strXml
        Dim className As String = mScript.mainClass(editClassId).Names.Last

        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return strXml
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wordId As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    Dim res As String = SeekChild2NameInString(cd, editClassId, newName, line, wordId)
                    If res = "Entrance" Then
                        dlgEntrancies.SetSeekPosInfo(line)
                        dlgEntrancies.NewEntrance()
                        wasFound = True
                    ElseIf res.Length > 0 Then
                        cd(line).Code(wordId).Word = res
                        wasFound = True
                    End If
                End If
            Next wordId
        Next line
        If wasFound Then
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                'собираем заново исполняемую строку
                strXml = cd(0).StartingSpaces
                For i As Integer = 0 To cd(0).Code.Count - 1
                    strXml += cd(0).Code(i).Word
                Next i
                strXml += cd(0).Comments
            Else
                'сериализаем код с внесенными изменениями
                strXml = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
            End If
        End If
        Return strXml
    End Function

    ''' <summary>
    ''' Переименовывает действия в содержимом события текущей локации
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="editClassId">Id класса</param>
    ''' <param name="oldName">Старое имя</param>
    ''' <param name="newName">Новое имя</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function RenameActionInCode(ByVal evendId As Integer, ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False
        Dim className As String = mScript.mainClass(editClassId).Names.Last

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line)) OrElse IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 1 To exData(line).Code.Count - 1 '1 словом быть не может
                If exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    Dim res As String = SeekChild2NameInStringEx(exData, editClassId, newName, line, wordId)
                    If res = "Entrance" Then
                        wasFound = True
                    ElseIf res.Length > 0 Then
                        exData(line).Code(wordId).Word = res
                        wasFound = True
                    End If
                End If
            Next wordId
        Next line
        Return wasFound
    End Function
#End Region

#Region "Check Actions"
    ''' <summary>
    ''' Переименовывает действия во всех свойствах и скриптах текущей локации
    ''' </summary>
    ''' <param name="remName">Имя действия (в кавычках)</param>
    ''' <param name="removedActionId">Id удаленного действия</param>
    ''' <param name="afterRemoving">Запускается ли процедура после удаления элемента или же это просто поиск</param>
    Public Sub CheckActionInStruct(remName As String, ByVal removedActionId As Integer, Optional ByVal afterRemoving As Boolean = True)
        Dim remClassId As Integer = mScript.mainClassHash("A")
        Dim className As String = mScript.mainClass(remClassId).Names.Last
        If afterRemoving Then
            dlgEntrancies.BeginNewEntrancies("Найдены строки, идентичные имени удаленного действия, значение которых вычислить не удалось. Провертьте их самостоятельно.", dlgEntrancies.EntranciesStyleEnum.Simple, removedActionId)
        Else
            dlgEntrancies.BeginNewEntrancies("Результаты поиска действия " + remName + " локации " + mScript.PrepareStringToPrint(actionsRouter.locationOfCurrentActions, Nothing, False) + ". Включены также те результаты, значение которых определить невозможно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        End If

        'Прочесывание классов Loc & Act структуры mainClass (только 2 уровень) на предмет наличия свойств типа ELEMENT нашего класса и имеющим значение = старое имя. Если найдем - переименовываем
        For i As Integer = 0 To 1
            Dim classId As Integer = IIf(i = 0, mScript.mainClassHash("L"), remClassId)
            Dim parentId As Integer = -1
            If classId = remClassId Then
                If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False Then
                    parentId = cPanelManager.ActivePanel.GetChild2Id
                ElseIf currentClassName = "A" Then
                    parentId = currentParentId
                End If
            End If

            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств каждого класса
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso p.returnArray(0) = className Then
                        'тип = элемент нужного класса
                        Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                        If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                            'Свойства 2 уровня
                            For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                If String.Compare(ch.Value, remName, True) = 0 Then
                                    dlgEntrancies.NewEntrance(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, parentId, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                                End If
                            Next child2Id
                        End If
                    End If
                Next pId
            End If
        Next i

        CheckActionInScripts(remClassId, remName)

        If dlgEntrancies.hasEntrancies Then
            dlgEntrancies.Show()
        End If
    End Sub

    ''' <summary>
    ''' Переименовывает действия во всех скриптах текущей локации
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    ''' <param name="remName">Имя действия (в кавычках)</param>
    Private Sub CheckActionInScripts(ByVal remClassId As Integer, ByVal remName As String)
        'Изменения в коде
        For i As Integer = 0 To 1
            Dim classId As Integer = IIf(i = 0, mScript.mainClassHash("L"), remClassId)
            Dim parentId As Integer = -1
            If classId = remClassId Then
                If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False Then
                    parentId = cPanelManager.ActivePanel.GetChild2Id
                ElseIf currentClassName = "A" Then
                    parentId = currentParentId
                End If
            End If
            'Перебираем классы L & A (только 2 уровень)
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RemoveActionInCode(ch.eventId, remClassId, remName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, parentId)
                            RemoveActionInExPropery(ch.Value, remClassId, remName)
                        End If
                    Next child2Id
                Next pId
            End If
        Next i
    End Sub

    ''' <summary>
    ''' Производит замены имени действия в исполняемой строке или сериализованном коде/длинном тексте текущей локации
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="remClassId">класс, в котором меняется имя</param>
    ''' <param name="remName">Имя действия (в кавычках)</param>
    Private Sub RemoveActionInExPropery(ByVal strXml As String, ByVal remClassId As Integer, ByVal remName As String)
        If String.IsNullOrEmpty(strXml) Then Return
        Dim className As String = mScript.mainClass(remClassId).Names.Last

        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wordId As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    If SeekChild2NameInStringRemoving(cd, remClassId, line, wordId) Then
                        dlgEntrancies.SetSeekPosInfo(line, mScript.PrepareStringToPrint(remName, Nothing, False), GetWordPosInLine(cd(line), wordId))
                        dlgEntrancies.NewEntrance()
                    End If
                End If
            Next wordId
        Next line
    End Sub

    ''' <summary>
    ''' Переименовывает действия в содержимом события текущей локации
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="remClassId">Id класса</param>
    ''' <param name="remName">Имя действия (в кавычках)</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function RemoveActionInCode(ByVal evendId As Integer, ByVal remClassId As Integer, ByVal remName As String) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False
        Dim className As String = mScript.mainClass(remClassId).Names.Last

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line)) OrElse IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 1 To exData(line).Code.Count - 1 '1 словом быть не может
                If exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    If SeekChild2NameInStringRemovingEx(exData, remClassId, line, wordId) Then
                        wasFound = True
                    End If
                End If
            Next wordId
        Next line
        Return wasFound
    End Function
#End Region

#Region "Rename Child3"
    ''' <summary>
    ''' Переименовывает элемент 3 уровня во всех свойствах, функциях, перменных со скриптами и т. д.
    ''' </summary>
    ''' <param name="editClassId">Id класса для замены</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="oldName">Старое имя элемента (в кавычках)</param>
    ''' <param name="newName">Новое имя элемента (в кавычках)</param>
    Public Sub RenameChild3InStruct(ByVal editClassId As Integer, ByVal parentId As Integer, oldName As String, newName As String)
        dlgEntrancies.BeginNewEntrancies("Найдены строки, идентичные старому имени элемента, значение которых вычислить не удалось. Провертьте их самостоятельно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        Dim className As String = mScript.mainClass(editClassId).Names.Last
        Dim parentName As String = mScript.mainClass(editClassId).ChildProperties(parentId)("Name").Value
        'Прочесывание структуры mainClass на предмет наличия свойств типа ELEMENT нашего класса и имеющим значение = parentId.
        'Если таковой имеется, то проверяем все свойства данного класса на предмет наличия строки = oldName
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств каждого класса 
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso p.returnArray(0) = className Then
                        'тип = элемент нужного класса
                        Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                        'Свойство по умолчанию
                        If String.Compare(p.Value, parentName, True) = 0 OrElse p.Value = parentId.ToString Then
                            'Значение свойства по умолчанию - родительский элемент по отношению к нашему
                            'Проверка всех свойств по умолчанию
                            For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                                Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(i).Value
                                Dim ppName As String = mScript.mainClass(classId).Properties.ElementAt(i).Key
                                If String.Compare(pp.Value, oldName, True) = 0 Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, ppName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0)
                                End If
                            Next i
                            Exit For
                        End If
                        If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                            'Свойства 2 уровня
                            For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                If String.Compare(ch.Value, parentName, True) = 0 OrElse ch.Value = parentId.ToString Then
                                    'Значение - родительский элемент по отношению к нашему
                                    'Проверка всех свойств этого элемента
                                    For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                                        Dim ppName As String = mScript.mainClass(classId).Properties.ElementAt(i).Key
                                        Dim ch2 As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(ppName)
                                        If String.Compare(ch2.Value, oldName, True) = 0 Then
                                            dlgEntrancies.NewEntrance(classId, child2Id, -1, ppName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0)
                                        End If
                                    Next i
                                End If
                                If IsNothing(ch.ThirdLevelProperties) = False Then
                                    'Свойства 3 уровня. Значение - родительский элемент по отношению к нашему
                                    For child3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                        If String.Compare(ch.ThirdLevelProperties(child3Id), parentName, True) = 0 OrElse ch.ThirdLevelProperties(child3Id) = parentId.ToString Then
                                            'Проверка всех свойств этого элемента
                                            For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                                                Dim ppName As String = mScript.mainClass(classId).Properties.ElementAt(i).Key
                                                Dim ch2 As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(ppName)
                                                If String.Compare(ch2.ThirdLevelProperties(child3Id), oldName, True) = 0 Then
                                                    dlgEntrancies.NewEntrance(classId, child2Id, child3Id, ppName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0)
                                                End If
                                            Next i
                                        End If
                                    Next child3Id
                                End If
                            Next child2Id
                        End If
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For
                        If pp.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(pp.returnArray) = False AndAlso pp.returnArray.Count = 1 AndAlso pp.returnArray(0) = className Then
                            'тип = элемент нужного класса
                            If String.Compare(ch.Value, parentName, True) = 0 OrElse ch.Value = parentId.ToString Then
                                'Значение - родительский элемент по отношению к нашему
                                'Проверка всех свойств этого элемента
                                For j As Integer = 0 To mScript.mainClass(classAct).Properties.Count - 1
                                    Dim ppName As String = mScript.mainClass(classAct).Properties.ElementAt(j).Key
                                    Dim ch2 As MatewScript.ChildPropertiesInfoType = arrProp(aId)(ppName)
                                    If String.Compare(ch2.Value, oldName, True) = 0 Then
                                        dlgEntrancies.NewEntrance(classAct, aId, -1, ppName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId, 0)
                                    End If
                                Next j
                            End If
                        End If
                    Next pId
                Next aId
            Next i
        End If

        RenameChild3InScripts(editClassId, parentId, oldName, newName)
        If dlgEntrancies.hasEntrancies Then
            dlgEntrancies.Show()
        End If
    End Sub

    ''' <summary>
    ''' Переименовывает элемент 3 уровня во всех скриптах
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="oldName">Старое имя элемента (в кавычках)</param>
    ''' <param name="newName">Новое имя элемента (в кавычках)</param>
    Private Sub RenameChild3InScripts(ByVal editClassId As Integer, ByVal parentId As Integer, ByVal oldName As String, ByVal newName As String)
        'Изменения в коде
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    If p.eventId > 0 AndAlso RenameChild3InCode(p.eventId, editClassId, parentId, oldName, newName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        p.Value = RenameChild3InExPropery(p.Value, editClassId, parentId, oldName, newName)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RenameChild3InCode(ch.eventId, editClassId, parentId, oldName, newName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                            ch.Value = RenameChild3InExPropery(ch.Value, editClassId, parentId, oldName, newName)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso RenameChild3InCode(ch.ThirdLevelEventId(child3Id), editClassId, parentId, oldName, newName) Then
                                'Значение - скрипт
                                dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                                ch.ThirdLevelProperties(child3Id) = RenameChild3InExPropery(ch.ThirdLevelProperties(child3Id), editClassId, parentId, oldName, newName)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If RenameChild3InCode(p.eventId, editClassId, parentId, oldName, newName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        p.Value = RenameChild3InExPropery(p.Value, editClassId, parentId, oldName, newName)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        If ch.eventId > 0 AndAlso RenameChild3InCode(ch.eventId, editClassId, parentId, oldName, newName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                            ch.Value = RenameChild3InExPropery(ch.Value, editClassId, parentId, oldName, newName)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        RenameChild3InFunctionHash(editClassId, parentId, oldName, newName)
        'В глобальных переменных
        RenameChild3InVariables(editClassId, parentId, oldName, newName)
        'В событиях изменения свойств
        RenameChild3InTracking(editClassId, parentId, oldName, newName)
    End Sub

    ''' <summary>
    ''' Производит замены имени элемента 3 порядка в глобальных переменных
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    Private Sub RenameChild3InVariables(ByVal editClassId As Integer, ByVal parentId As Integer, ByVal oldName As String, ByVal newName As String)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return
        Dim className As String = mScript.mainClass(editClassId).Names.Last
        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            Dim wasFoundGlobal As Boolean = False 'были ли изменения хоть в одном элементе массива переменной
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim cd() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    cd = .codeBox.CodeData
                End With
                If IsNothing(cd) Then Continue For
                'В элементе массива - исполняемый код
                Dim wasFound As Boolean = False 'была ли замена в данном элементе массива
                For line As Integer = 0 To cd.Count - 1
                    If IsNothing(cd(line).Code) Then Continue For

                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            Dim res As String = SeekChild3InString(cd, editClassId, parentId, newName, line, wordId)
                            If res = "Entrance" Then
                                dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line)
                            ElseIf res.Length > 0 Then
                                cd(line).Code(wordId).Word = res
                                wasFound = True
                            End If
                            'Dim shouldBeChecked As Boolean = False
                            'Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
                            'If GetCurrentElementWord(cd, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
                            '    If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                            '        Dim f As MatewScript.PropertiesInfoType = Nothing
                            '        If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                            '            If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT _
                            '                AndAlso IsNothing(f.params(paramNumber - 1).EnumValues) = False AndAlso f.params(paramNumber - 1).EnumValues.Count = 1 Then
                            '                If f.params(paramNumber - 1).EnumValues(0) = className Then
                            '                    'Строка точно является именем элемента 2 уровня, стоит параметром в функции
                            '                    wasFound = True
                            '                    Dim w As String = cd(line).Code(wordId).Word
                            '                    cd(line).Code(wordId).Word = w.Replace(w.Trim, newName)
                            '                End If
                            '            Else
                            '                shouldBeChecked = True
                            '            End If
                            '        Else
                            '            shouldBeChecked = True
                            '        End If
                            '    ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                            '        Dim p As MatewScript.PropertiesInfoType = Nothing
                            '        If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                            '            If elClassId = editClassId AndAlso paramNumber = 1 Then
                            '                'Строка точно является именем элемента 2 уровня, стоит первым параметром в свойстве
                            '                wasFound = True
                            '                Dim w As String = cd(line).Code(wordId).Word
                            '                cd(line).Code(wordId).Word = w.Replace(w.Trim, newName)
                            '            End If
                            '        Else
                            '            shouldBeChecked = True
                            '        End If
                            '    Else
                            '        shouldBeChecked = True
                            '    End If
                            'ElseIf wordId > 0 AndAlso (cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                            '                           cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                            '                           cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) AndAlso _
                            '                       GetLastPropertyReturnElementClassId(cd, wordId - 1, line) = editClassId Then
                            '    Dim w As String = cd(line).Code(wordId).Word
                            '    cd(line).Code(wordId).Word = w.Replace(w.Trim, newName)
                            '    wasFound = True
                            'Else
                            '    shouldBeChecked = True
                            'End If
                            'If shouldBeChecked Then
                            '    dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line)
                            'End If
                        End If
                    Next wordId
                Next line
                If wasFound Then
                    wasFoundGlobal = True
                    If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                        'Исполняемая строка. Собираем измененную строку
                        Dim w As String = cd(0).StartingSpaces
                        For i As Integer = 0 To cd(0).Code.Count - 1
                            w += cd(0).Code(i).Word
                        Next i
                        w += cd(0).Comments
                        arrValues(arrId) = WrapString(w)
                    Else
                        'Скрипт или длинный текст. Собираем с внесенными изменениями
                        arrValues(arrId) = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
                    End If
                End If
            Next arrId
            If wasFoundGlobal Then
                'Сохраняем результат
                Dim var As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value
                var.arrValues = arrValues
            End If
        Next varId
    End Sub

    ''' <summary>
    ''' Производит замены имени элемента 3 порядка в функциях писателя functionsHash
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    Private Sub RenameChild3InFunctionHash(ByVal editClassId As Integer, ByVal parentId As Integer, ByVal oldName As String, ByVal newName As String)
        If IsNothing(mScript.functionsHash) Then Return
        Dim className As String = mScript.mainClass(editClassId).Names.Last
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            'В функции хранится код в исполняемом виде - в ValueExecuteDt, и в виде для редактирования - в ValueDt (при этом второе может быть пустым, если функция объявлена из кода)
            'Первое быть пустым не должно
            Dim wasFound As Boolean = False
            If IsNothing(func.ValueExecuteDt) = False Then
                Dim edt As List(Of MatewScript.ExecuteDataType) = func.ValueExecuteDt
                For line As Integer = 0 To edt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(edt(line).Code) Then Continue For
                    For wordId As Integer = 1 To edt(line).Code.Count - 1 '1 словом быть не может
                        If edt(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(edt(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            Dim res As String = SeekChild3InStringEx(edt, editClassId, parentId, newName, line, wordId)
                            If res = "Entrance" Then
                                wasFound = True
                            ElseIf res.Length > 0 Then
                                edt(line).Code(wordId).Word = res
                                wasFound = True
                            End If
                        End If
                    Next wordId
                Next line
            End If

            If Not wasFound Then Continue For 'если не найдено в ValueExecuteDt, то здесь тоже не будет
            If IsNothing(func.ValueDt) = False Then
                Dim cd() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            Dim res As String = SeekChild3InString(cd, editClassId, parentId, newName, line, wordId)
                            If res = "Entrance" Then
                                dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line)
                            ElseIf res.Length > 0 Then
                                cd(line).Code(wordId).Word = res
                                wasFound = True
                            End If
                        End If
                    Next wordId
                Next line
            End If
        Next funcId
    End Sub

    Private Sub RenameChild3InTracking(ByVal editClassId As Integer, ByVal parentId As Integer, ByVal oldName As String, ByVal newName As String)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return

        Dim className As String = mScript.mainClass(editClassId).Names.Last
        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
            Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            If IsNothing(tr.propBeforeContent) = False Then
                If RenameChild3InCode(tr.eventBeforeId, editClassId, parentId, oldName, newName) Then
                    Dim cd() As CodeTextBox.CodeDataType = tr.propBeforeContent
                    For line As Integer = 0 To cd.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(cd(line).Code) Then Continue For
                        For wordId As Integer = 0 To cd(line).Code.Count - 1
                            'перибираем слова в строках кода
                            If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                Dim res As String = SeekChild3InString(cd, editClassId, parentId, newName, line, wordId)
                                If res = "Entrance" Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                                ElseIf res.Length > 0 Then
                                    cd(line).Code(wordId).Word = res
                                End If
                            End If
                        Next wordId
                    Next line
                End If
            End If

            If IsNothing(tr.propAfterContent) = False Then
                If RenameChild3InCode(tr.eventAfterId, editClassId, parentId, oldName, newName) Then
                    Dim cd() As CodeTextBox.CodeDataType = tr.propAfterContent
                    For line As Integer = 0 To cd.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(cd(line).Code) Then Continue For
                        For wordId As Integer = 0 To cd(line).Code.Count - 1
                            'перибираем слова в строках кода
                            If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                                'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                                Dim res As String = SeekChild3InString(cd, editClassId, parentId, newName, line, wordId)
                                If res = "Entrance" Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                                ElseIf res.Length > 0 Then
                                    cd(line).Code(wordId).Word = res
                                End If
                            End If
                        Next wordId
                    Next line
                End If
            End If
        Next trId
    End Sub

    ''' <summary>
    ''' Производит замены имени элемента 3 порядка в исполняемой строке или сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="editClassId">класс, в котором меняется имя</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    ''' <returns>Исполнемую строку с внесенными изменениями</returns>
    Private Function RenameChild3InExPropery(ByVal strXml As String, ByVal editClassId As Integer, ByVal parentId As Integer, ByVal oldName As String, ByVal newName As String) As String
        If String.IsNullOrEmpty(strXml) Then Return strXml

        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return strXml
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wordId As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    Dim res As String = SeekChild3InString(cd, editClassId, parentId, newName, line, wordId)
                    If res = "Entrance" Then
                        dlgEntrancies.SetSeekPosInfo(line)
                        dlgEntrancies.NewEntrance()
                    ElseIf res.Length > 0 Then
                        cd(line).Code(wordId).Word = res
                        wasFound = True
                    End If
                End If
            Next wordId
        Next line
        If wasFound Then
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                'собираем заново исполняемую строку
                strXml = cd(0).StartingSpaces
                For i As Integer = 0 To cd(0).Code.Count - 1
                    strXml += cd(0).Code(i).Word
                Next i
                strXml += cd(0).Comments
            Else
                'сериализаем код с внесенными изменениями
                strXml = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
            End If
        End If
        Return strXml
    End Function

    ''' <summary>
    ''' Переименовывает элементы 3 уровня в содержимом события
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="editClassId">Id класса</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="oldName">Старое имя</param>
    ''' <param name="newName">Новое имя</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function RenameChild3InCode(ByVal evendId As Integer, ByVal editClassId As Integer, ByVal parentId As Integer, ByVal oldName As String, ByVal newName As String) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False
        Dim className As String = mScript.mainClass(editClassId).Names.Last

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 1 To exData(line).Code.Count - 1 '1 словом быть не может
                If exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId).Word.Trim, oldName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    Dim res As String = SeekChild3InStringEx(exData, editClassId, parentId, newName, line, wordId)
                    If res = "Entrance" Then
                        wasFound = True
                    ElseIf res.Length > 0 Then
                        exData(line).Code(wordId).Word = res
                        wasFound = True
                    End If
                End If
            Next wordId
        Next line
        Return wasFound
    End Function

    ''' <summary>
    ''' Внутрення функция группы Rename Child3. Определяет может ли слово с индексом wordId являться именем элемента 3 уровня класса editClassId, дочерним элементом по отношению к parentId
    ''' </summary>
    ''' <param name="cd">Код в формате CodeTextBox.CodeDataType</param>
    ''' <param name="editClassId">Класс элемента 3 уровня</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="newName">Новое имя элемента 3 уровня</param>
    ''' <param name="line">Линия в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>True если может</returns>
    Private Function SeekChild3InString(ByRef cd() As CodeTextBox.CodeDataType, ByVal editClassId As Integer, ByVal parentId As Integer, ByVal newName As String, ByVal line As Integer, _
                                        ByVal wordId As Integer) As String
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim className As String = mScript.mainClass(editClassId).Names.Last
        Dim parentName As String = mScript.mainClass(editClassId).ChildProperties(parentId)("Name").Value
        Dim shouldBeChecked As Boolean = False
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWord(cd, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT2 Then
                        'Тип текущего параметра - элемент 3 уровня
                        'Получаем параметр типа ELEMENT этой же функции
                        Dim elementParam As MatewScript.paramsType = Nothing
                        Dim elementName As String = GetFunctionParamWhichTypeIsElement(cd, line, elementWordId, f.params, elementParam)
                        If String.IsNullOrEmpty(elementName) Then Return ""
                        If IsNothing(elementParam.EnumValues) = False AndAlso elementParam.EnumValues.Count > 0 AndAlso elementParam.EnumValues(0) = className Then
                            If String.Compare(elementName, parentName, True) = 0 Then
                                Dim w As String = cd(line).Code(wordId).Word
                                Return w.Replace(w.Trim, newName)
                            Else
                                Return ""
                            End If
                        Else
                            Return ""
                        End If
                    Else
                        Return "Entrance"
                    End If
                Else
                    Return "Entrance"
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    If elClassId = editClassId AndAlso paramNumber = 2 Then
                        'Строка точно является именем элемента 3 уровня, стоит вторым параметром в свойстве
                        If cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA AndAlso wordId >= 2 Then
                            If (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId - 2).Word.Trim, parentName, True) = 0) OrElse _
                                (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER AndAlso cd(line).Code(wordId - 2).Word.Trim = parentId.ToString) Then
                                Dim w As String = cd(line).Code(wordId).Word
                                Return w.Replace(w.Trim, newName)
                            Else
                                Return ""
                            End If
                        Else
                            Return ""
                        End If
                    End If
                Else
                    Return "Entrance"
                End If
            Else
                Return ""
            End If
        ElseIf wordId = cd(line).Code.Count - 1 AndAlso (cd(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT OrElse cd(line).Code(0).wordType = _
                                                         CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION) Then
            If cd(line).Code.Count > 5 AndAlso cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA AndAlso _
                (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId - 2).Word.Trim, parentName, True) = 0) OrElse _
                (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER AndAlso cd(line).Code(wordId - 2).Word.Trim = parentId.ToString) Then
                Dim w As String = cd(line).Code(wordId).Word
                Return w.Replace(w.Trim, newName)
            Else
                Return ""
            End If
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Внутрення функция группы Rename Child3. Определяет может ли слово с индексом wordId являться именем элемента 3 уровня класса editClassId, дочерним элементом по отношению к parentId
    ''' </summary>
    ''' <param name="cd">Код в формате MatewScript.ExecuteDataType</param>
    ''' <param name="editClassId">Класс элемента 3 уровня</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="newName">Новое имя элемента 3 уровня</param>
    ''' <param name="line">Линия в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>True если может</returns>
    Private Function SeekChild3InStringEx(ByRef cd As List(Of MatewScript.ExecuteDataType), ByVal editClassId As Integer, ByVal parentId As Integer, ByVal newName As String, ByVal line As Integer, _
                                         ByVal wordId As Integer) As String
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim className As String = mScript.mainClass(editClassId).Names.Last
        Dim parentName As String = mScript.mainClass(editClassId).ChildProperties(parentId)("Name").Value
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWordInExData(cd, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT2 Then
                        'Тип текущего параметра - элемент 3 уровня
                        'Получаем параметр типа ELEMENT этой же функции
                        Dim elementParam As MatewScript.paramsType = Nothing
                        Dim elementName As String = GetFunctionParamWhichTypeIsElementEx(cd, line, elementWordId, f.params, elementParam)
                        If String.IsNullOrEmpty(elementName) Then Return ""
                        If IsNothing(elementParam.EnumValues) = False AndAlso elementParam.EnumValues.Count > 0 AndAlso elementParam.EnumValues(0) = className Then
                            If String.Compare(elementName, parentName, True) = 0 Then
                                Dim w As String = cd(line).Code(wordId).Word
                                Return w.Replace(w.Trim, newName)
                            Else
                                Return ""
                            End If
                        Else
                            Return ""
                        End If
                    Else
                        Return "Entrance"
                    End If
                Else
                    Return "Entrance"
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    If elClassId = editClassId AndAlso paramNumber = 2 Then
                        'Строка точно является именем элемента 3 уровня, стоит вторым параметром в свойстве
                        If cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA AndAlso wordId >= 2 Then
                            If (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId - 2).Word.Trim, parentName, True) = 0) OrElse _
                                (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER AndAlso cd(line).Code(wordId - 2).Word.Trim = parentId.ToString) Then
                                Dim w As String = cd(line).Code(wordId).Word
                                Return w.Replace(w.Trim, newName)
                            Else
                                Return ""
                            End If
                        Else
                            Return ""
                        End If
                    End If
                Else
                    Return "Entrance"
                End If
            Else
                Return ""
            End If
        ElseIf wordId = cd(line).Code.Count - 1 AndAlso (cd(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT OrElse cd(line).Code(0).wordType = _
                                                         CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION) Then
            If cd(line).Code.Count > 5 AndAlso cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA AndAlso _
                (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId - 2).Word.Trim, parentName, True) = 0) OrElse _
                (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER AndAlso cd(line).Code(wordId - 2).Word.Trim = parentId.ToString) Then
                Dim w As String = cd(line).Code(wordId).Word
                Return w.Replace(w.Trim, newName)
            Else
                Return ""
            End If
        End If
        Return ""
    End Function
#End Region

#Region "Check Child3"
    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 3 уровня во всех свойствах, функциях, перменных со скриптами и т. д.
    ''' </summary>
    ''' <param name="remClassId">Id класса для замены</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="remName">Имя удаляемого элемента 3 уровня (в кавычках)</param>
    ''' <param name="afterRemoving">Запускается ли процедура после удаления элемента или же это просто поиск</param>
    Public Sub CheckChild3InStruct(ByVal remClassId As Integer, ByVal parentId As Integer, remName As String, Optional ByVal afterRemoving As Boolean = True)
        Dim className As String = mScript.mainClass(remClassId).Names.Last
        Dim parentName As String = mScript.mainClass(remClassId).ChildProperties(parentId)("Name").Value
        If afterRemoving Then
            dlgEntrancies.BeginNewEntrancies("Найдены строки, идентичные старому имени элемента, значение которых вычислить не удалось. Провертьте их самостоятельно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        Else
            dlgEntrancies.BeginNewEntrancies("Результаты поиска элемента " + mScript.PrepareStringToPrint(parentName, Nothing, False) _
                                             + " --> " + mScript.PrepareStringToPrint(remName, Nothing, False) + " класса " + className + ". Включены также те результаты, значение которых определить невозможно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        End If
        'Прочесывание структуры mainClass на предмет наличия свойств типа ELEMENT нашего класса и имеющим значение = parentId.
        'Если таковой имеется, то проверяем все свойства данного класса на предмет наличия строки = oldName
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств каждого класса 
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso p.returnArray(0) = className Then
                        'тип = элемент нужного класса
                        Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                        'Свойство по умолчанию
                        If String.Compare(p.Value, parentName, True) = 0 OrElse p.Value = parentId.ToString Then
                            'Значение свойства по умолчанию - родительский элемент по отношению к нашему
                            'Проверка всех свойств по умолчанию
                            For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                                Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(i).Value
                                Dim ppName As String = mScript.mainClass(classId).Properties.ElementAt(i).Key
                                If String.Compare(pp.Value, remName, True) = 0 Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, ppName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                                End If
                            Next i
                            Exit For
                        End If
                        If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                            'Свойства 2 уровня
                            For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                If String.Compare(ch.Value, parentName, True) = 0 OrElse ch.Value = parentId.ToString Then
                                    'Значение - родительский элемент по отношению к нашему
                                    'Проверка всех свойств этого элемента
                                    For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                                        Dim ppName As String = mScript.mainClass(classId).Properties.ElementAt(i).Key
                                        Dim ch2 As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(ppName)
                                        If String.Compare(ch2.Value, remName, True) = 0 Then
                                            dlgEntrancies.NewEntrance(classId, child2Id, -1, ppName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                                        End If
                                    Next i
                                End If
                                If IsNothing(ch.ThirdLevelProperties) = False Then
                                    'Свойства 3 уровня. Значение - родительский элемент по отношению к нашему
                                    For child3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                        If String.Compare(ch.ThirdLevelProperties(child3Id), parentName, True) = 0 OrElse ch.ThirdLevelProperties(child3Id) = parentId.ToString Then
                                            'Проверка всех свойств этого элемента
                                            For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                                                Dim ppName As String = mScript.mainClass(classId).Properties.ElementAt(i).Key
                                                Dim ch2 As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(ppName)
                                                If String.Compare(ch2.ThirdLevelProperties(child3Id), remName, True) = 0 Then
                                                    dlgEntrancies.NewEntrance(classId, child2Id, child3Id, ppName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                                                End If
                                            Next i
                                        End If
                                    Next child3Id
                                End If
                            Next child2Id
                        End If
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For
                        If pp.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(pp.returnArray) = False AndAlso pp.returnArray.Count = 1 AndAlso pp.returnArray(0) = className Then
                            'тип = элемент нужного класса
                            If String.Compare(ch.Value, parentName, True) = 0 OrElse ch.Value = parentId.ToString Then
                                'Значение - родительский элемент по отношению к нашему
                                'Проверка всех свойств этого элемента
                                For j As Integer = 0 To mScript.mainClass(classAct).Properties.Count - 1
                                    Dim ppName As String = mScript.mainClass(classAct).Properties.ElementAt(j).Key
                                    Dim ch2 As MatewScript.ChildPropertiesInfoType = arrProp(aId)(ppName)
                                    If String.Compare(ch2.Value, remName, True) = 0 Then
                                        dlgEntrancies.NewEntrance(classAct, aId, -1, ppName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId, 0, mScript.PrepareStringToPrint(remName, Nothing, False), 0)
                                    End If
                                Next j
                            End If
                        End If
                    Next pId
                Next aId
            Next i
        End If

        CheckChild3InScripts(remClassId, parentId, remName)
        If dlgEntrancies.hasEntrancies Then
            dlgEntrancies.Show()
        End If
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 3 уровня во всех скриптах
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="remName">Имя удаляемого элемента 3 уровня (в кавычках)</param>
    Private Sub CheckChild3InScripts(ByVal remClassId As Integer, ByVal parentId As Integer, ByVal remName As String)
        'Изменения в коде
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    If p.eventId > 0 AndAlso RemoveChild3InCode(p.eventId, remClassId, parentId, remName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        RemoveChild3InExPropery(p.Value, remClassId, parentId, remName)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RemoveChild3InCode(ch.eventId, remClassId, parentId, remName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                            RemoveChild3InExPropery(ch.Value, remClassId, parentId, remName)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso RemoveChild3InCode(ch.ThirdLevelEventId(child3Id), remClassId, parentId, remName) Then
                                'Значение - скрипт
                                dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                                RemoveChild3InExPropery(ch.ThirdLevelProperties(child3Id), parentId, remClassId, remName)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If RemoveChild3InCode(p.eventId, remClassId, parentId, remName) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        RemoveChild3InExPropery(p.Value, remClassId, parentId, remName)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        If ch.eventId > 0 AndAlso RemoveChild3InCode(ch.eventId, remClassId, parentId, remName) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                            RemoveChild3InExPropery(ch.Value, remClassId, parentId, remName)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        RemoveChild3InFunctionHash(remClassId, parentId, remName)
        'В глобальных переменных
        RemoveChild3InVariables(remClassId, parentId, remName)
        'В событиях изменения свойств
        RemoveChild3InTracking(remClassId, parentId, remName)
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 3 порядка в глобальных переменных
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="remName">Имя удаляемого элемента 3 уровня</param>
    Private Sub RemoveChild3InVariables(ByVal remClassId As Integer, ByVal parentId As Integer, ByVal remName As String)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return
        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            Dim wasFoundGlobal As Boolean = False 'были ли изменения хоть в одном элементе массива переменной
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim cd() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    cd = .codeBox.CodeData
                End With
                If IsNothing(cd) Then Continue For
                'В элементе массива - исполняемый код
                Dim wasFound As Boolean = False 'была ли замена в данном элементе массива
                For line As Integer = 0 To cd.Count - 1
                    If IsNothing(cd(line).Code) Then Continue For

                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekChild3InStringRemoving(cd, remClassId, parentId, remName, line, wordId) Then
                                dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), GetWordPosInLine(cd(line), wordId))
                            End If
                        End If
                    Next wordId
                Next line
                If wasFound Then
                    wasFoundGlobal = True
                    If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                        'Исполняемая строка. Собираем измененную строку
                        Dim w As String = cd(0).StartingSpaces
                        For i As Integer = 0 To cd(0).Code.Count - 1
                            w += cd(0).Code(i).Word
                        Next i
                        w += cd(0).Comments
                        arrValues(arrId) = WrapString(w)
                    Else
                        'Скрипт или длинный текст. Собираем с внесенными изменениями
                        arrValues(arrId) = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
                    End If
                End If
            Next arrId
            If wasFoundGlobal Then
                'Сохраняем результат
                Dim var As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value
                var.arrValues = arrValues
            End If
        Next varId
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 3 порядка в функциях писателя functionsHash
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="remName">Имя удаляемого элемента 3 уровня</param>
    Private Sub RemoveChild3InFunctionHash(ByVal remClassId As Integer, ByVal parentId As Integer, ByVal remName As String)
        If IsNothing(mScript.functionsHash) Then Return
        Dim className As String = mScript.mainClass(remClassId).Names.Last
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            'В функции хранится код в исполняемом виде - в ValueExecuteDt, и в виде для редактирования - в ValueDt (при этом второе может быть пустым, если функция объявлена из кода)
            'Первое быть пустым не должно
            Dim wasFound As Boolean = False
            If IsNothing(func.ValueExecuteDt) = False Then
                Dim edt As List(Of MatewScript.ExecuteDataType) = func.ValueExecuteDt
                For line As Integer = 0 To edt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(edt(line).Code) Then Continue For
                    For wordId As Integer = 1 To edt(line).Code.Count - 1 '1 словом быть не может
                        If edt(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(edt(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekChild3InStringRemovingEx(edt, remClassId, parentId, remName, line, wordId) Then
                                wasFound = True
                            End If
                        End If
                    Next wordId
                Next line
            End If

            If Not wasFound Then Continue For 'если не найдено в ValueExecuteDt, то здесь тоже не будет
            If IsNothing(func.ValueDt) = False Then
                Dim cd() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekChild3InStringRemoving(cd, remClassId, parentId, remName, line, wordId) Then
                                dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), _
                                                          GetWordPosInLine(cd(line), wordId))
                            End If
                        End If
                    Next wordId
                Next line
            End If
        Next funcId
    End Sub

    Private Sub RemoveChild3InTracking(ByVal remClassId As Integer, ByVal parentId As Integer, ByVal remName As String)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return

        Dim className As String = mScript.mainClass(remClassId).Names.Last
        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
            Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            If IsNothing(tr.propBeforeContent) = False Then
                Dim cd() As CodeTextBox.CodeDataType = tr.propBeforeContent
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekChild3InStringRemoving(cd, remClassId, parentId, remName, line, wordId) Then
                                dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), _
                                                          GetWordPosInLine(cd(line), wordId), frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                            End If
                        End If
                    Next wordId
                Next line
            End If

            If IsNothing(tr.propAfterContent) = False Then
                Dim cd() As CodeTextBox.CodeDataType = tr.propAfterContent
                For line As Integer = 0 To cd.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(cd(line).Code) Then Continue For
                    For wordId As Integer = 0 To cd(line).Code.Count - 1
                        'перибираем слова в строках кода
                        If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                            'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                            If SeekChild3InStringRemoving(cd, remClassId, parentId, remName, line, wordId) Then
                                dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, mScript.PrepareStringToPrint(remName, Nothing, False), _
                                                          GetWordPosInLine(cd(line), wordId), frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                            End If
                        End If
                    Next wordId
                Next line
            End If
        Next trId
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 3 порядка в исполняемой строке или сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="remClassId">класс, в котором меняется имя</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="remName">Имя удаляемого элемента 3 уровня</param>
    Private Sub RemoveChild3InExPropery(ByVal strXml As String, ByVal remClassId As Integer, ByVal parentId As Integer, ByVal remName As String)
        If String.IsNullOrEmpty(strXml) Then Return

        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wordId As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                If cd(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    If SeekChild3InStringRemoving(cd, remClassId, parentId, remName, line, wordId) Then
                        dlgEntrancies.SetSeekPosInfo(line, mScript.PrepareStringToPrint(remName, Nothing, False), GetWordPosInLine(cd(line), wordId))
                        dlgEntrancies.NewEntrance()
                    End If
                End If
            Next wordId
        Next line
    End Sub

    ''' <summary>
    ''' Ищет ссылки на удаленные элементы 3 уровня в содержимом события
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="remClassId">Id класса</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="remName">Имя удаляемого элемента 3 уровня</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function RemoveChild3InCode(ByVal evendId As Integer, ByVal remClassId As Integer, ByVal parentId As Integer, ByVal remName As String) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False
        Dim className As String = mScript.mainClass(remClassId).Names.Last

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line)) OrElse IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 1 To exData(line).Code.Count - 1 '1 словом быть не может
                If exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId).Word.Trim, remName, True) = 0 Then
                    'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
                    If SeekChild3InStringRemovingEx(exData, remClassId, parentId, remName, line, wordId) Then wasFound = True
                End If
            Next wordId
        Next line
        Return wasFound
    End Function

    ''' <summary>
    ''' Внутрення функция группы Remove Child3. Определяет может ли слово с индексом wordId являться именем элемента 3 уровня класса editClassId, дочерним элементом по отношению к parentId
    ''' </summary>
    ''' <param name="cd">Код в формате CodeTextBox.CodeDataType</param>
    ''' <param name="remClassId">Класс элемента 3 уровня</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="remName">Имя удаляемого элемента 3 уровня</param>
    ''' <param name="line">Линия в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>True если может</returns>
    Private Function SeekChild3InStringRemoving(ByRef cd() As CodeTextBox.CodeDataType, ByVal remClassId As Integer, ByVal parentId As Integer, ByVal remName As String, ByVal line As Integer, _
                                        ByVal wordId As Integer) As Boolean
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim className As String = mScript.mainClass(remClassId).Names.Last
        Dim parentName As String = mScript.mainClass(remClassId).ChildProperties(parentId)("Name").Value
        Dim shouldBeChecked As Boolean = False
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWord(cd, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT2 Then
                        'Тип текущего параметра - элемент 3 уровня
                        'Получаем параметр типа ELEMENT этой же функции
                        Dim elementParam As MatewScript.paramsType = Nothing
                        Dim elementName As String = GetFunctionParamWhichTypeIsElement(cd, line, elementWordId, f.params, elementParam)
                        If String.IsNullOrEmpty(elementName) Then Return False
                        If IsNothing(elementParam.EnumValues) = False AndAlso elementParam.EnumValues.Count > 0 AndAlso elementParam.EnumValues(0) = className Then
                            If String.Compare(elementName, parentName, True) = 0 Then
                                Return True
                            Else
                                Return False
                            End If
                        Else
                            Return False
                        End If
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    If elClassId = remClassId AndAlso paramNumber = 2 Then
                        'Строка точно является именем элемента 3 уровня, стоит вторым параметром в свойстве
                        If cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA AndAlso wordId >= 2 Then
                            If (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId - 2).Word.Trim, parentName, True) = 0) OrElse _
                                (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER AndAlso cd(line).Code(wordId - 2).Word.Trim = parentId.ToString) Then
                                Return True
                            Else
                                Return False
                            End If
                        Else
                            Return False
                        End If
                    End If
                Else
                    Return True
                End If
            Else
                Return False
            End If
        ElseIf wordId = cd(line).Code.Count - 1 AndAlso (cd(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT OrElse cd(line).Code(0).wordType = _
                                                         CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION) Then
            If cd(line).Code.Count > 5 AndAlso cd(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA AndAlso _
                (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(cd(line).Code(wordId - 2).Word.Trim, parentName, True) = 0) OrElse _
                (cd(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER AndAlso cd(line).Code(wordId - 2).Word.Trim = parentId.ToString) Then
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function

    ''' <summary>
    ''' Внутрення функция группы Remove Child3. Определяет может ли слово с индексом wordId являться именем элемента 3 уровня класса editClassId, дочерним элементом по отношению к parentId
    ''' </summary>
    ''' <param name="exData">Код в формате MatewScript.ExecuteDataType</param>
    ''' <param name="remClassId">Класс элемента 3 уровня</param>
    ''' <param name="parentId">Id родителя</param>
    ''' <param name="remName">Имя удаляемого элемента 3 уровня</param>
    ''' <param name="line">Линия в скрипте, где находится слово</param>
    ''' <param name="wordId">Индекс слова</param>
    ''' <returns>True если может</returns>
    Private Function SeekChild3InStringRemovingEx(ByRef exData As List(Of MatewScript.ExecuteDataType), ByVal remClassId As Integer, ByVal parentId As Integer, ByVal remName As String, ByVal line As Integer, _
                                         ByVal wordId As Integer) As Boolean
        'Совпадение имени. Но может быть что элемент другого класса имеет такое же имя
        Dim className As String = mScript.mainClass(remClassId).Names.Last
        Dim parentName As String = mScript.mainClass(remClassId).ChildProperties(parentId)("Name").Value
        Dim shouldBeChecked As Boolean = False
        Dim paramNumber As Integer = 0, elementWord As CodeTextBox.EditWordType = Nothing, elClassId As Integer = -1, elementWordId As Integer = -1
        If GetCurrentElementWordInExData(exData, line, wordId, paramNumber, elementWord, elClassId, elementWordId) Then
            If elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                Dim f As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Functions.TryGetValue(elementWord.Word.Trim, f) Then
                    If IsNothing(f.params) = False AndAlso paramNumber <= f.params.Count AndAlso f.params(paramNumber - 1).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT2 Then
                        'Тип текущего параметра - элемент 3 уровня
                        'Получаем параметр типа ELEMENT этой же функции
                        Dim elementParam As MatewScript.paramsType = Nothing
                        Dim elementName As String = GetFunctionParamWhichTypeIsElementEx(exData, line, elementWordId, f.params, elementParam)
                        If String.IsNullOrEmpty(elementName) Then Return False
                        If IsNothing(elementParam.EnumValues) = False AndAlso elementParam.EnumValues.Count > 0 AndAlso elementParam.EnumValues(0) = className Then
                            If String.Compare(elementName, parentName, True) = 0 Then
                                Return True
                            Else
                                Return False
                            End If
                        Else
                            Return False
                        End If
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If
            ElseIf elementWord.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(elClassId).Properties.TryGetValue(elementWord.Word.Trim, p) Then
                    If elClassId = remClassId AndAlso paramNumber = 2 Then
                        'Строка точно является именем элемента 3 уровня, стоит вторым параметром в свойстве
                        If exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA AndAlso wordId >= 2 Then
                            If (exData(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId - 2).Word.Trim, parentName, True) = 0) OrElse _
                                (exData(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER AndAlso exData(line).Code(wordId - 2).Word.Trim = parentId.ToString) Then
                                Return True
                            Else
                                Return False
                            End If
                        Else
                            Return False
                        End If
                    End If
                Else
                    Return True
                End If
            Else
                Return False
            End If
        ElseIf wordId = exData(line).Code.Count - 1 AndAlso (exData(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT OrElse exData(line).Code(0).wordType = _
                                                         CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION) Then
            If exData(line).Code.Count > 5 AndAlso exData(line).Code(wordId - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA AndAlso _
                (exData(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso String.Compare(exData(line).Code(wordId - 2).Word.Trim, parentName, True) = 0) OrElse _
                (exData(line).Code(wordId - 2).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER AndAlso exData(line).Code(wordId - 2).Word.Trim = parentId.ToString) Then
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function
#End Region

#Region "Replace Class Name"
    ''' <summary>
    ''' Заменят имя класса на новое во всех свойствах, функциях, перменных со скриптами и т. д.
    ''' </summary>
    ''' <param name="editClassId">Id класса для замены</param>
    ''' <param name="newNames">Имена</param>
    ''' <remarks></remarks>
    Public Sub ReplaceClassNameInStruct(ByVal editClassId As Integer, ByVal newNames() As String)
        Dim newShort As String = newNames(0)
        Dim oldLong As String = mScript.mainClass(editClassId).Names.Last
        Dim newLong As String = newNames.Last

        'Прочесывание структуры mainClass на предмет наличия функций и свойств, содержащих в качестве возвращаемого значения или параметра функции тип ELEMENT нашего переименовываемого класса класса. Если будет такое найдено - очищаем формат
        If oldLong <> newLong Then
            For classId As Integer = 0 To mScript.mainClass.Count - 1
                If IsNothing(mScript.mainClass(classId).Properties) = False Then
                    'проверка свойств каждого класса
                    For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                        Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                        If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso p.returnArray(0) = oldLong Then
                            'тип = элемент переименовываемого класса
                            p.returnArray(0) = newLong
                        End If
                    Next pId
                End If

                If IsNothing(mScript.mainClass(classId).Functions) = False Then
                    'проверка свойств функций класса
                    For fId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                        Dim f As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(fId).Value
                        Dim fName As String = mScript.mainClass(classId).Functions.ElementAt(fId).Key
                        If f.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(f.returnArray) = False AndAlso f.returnArray.Count = 1 AndAlso f.returnArray(0) = oldLong Then
                            'тип = элемент переименовываемого класса.
                            f.returnArray(0) = newLong
                        End If
                        If IsNothing(f.params) = False AndAlso f.params.Count > 0 Then
                            For pId As Integer = 0 To f.params.Count - 1
                                'Проверка параметров функции
                                If f.params(pId).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT AndAlso IsNothing(f.params(pId).EnumValues) = False AndAlso f.params(pId).EnumValues(0) = oldLong Then
                                    'тип параметра = элементу переименовываемого класса.
                                    f.params(pId).EnumValues(0) = newLong
                                End If
                            Next pId
                        End If
                    Next fId
                End If
            Next classId
        End If

        ReplaceClassNameInScripts(editClassId, newShort)
        mScript.MakeMainClassHash()
        mScript.FillFuncAndPropHash()
    End Sub

    ''' <summary>
    ''' Изменяет имя указанного класса во всех скриптах
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="newName">Новое имя елемента</param>
    Private Sub ReplaceClassNameInScripts(ByVal editClassId As Integer, ByVal newName As String)
        'Изменения в коде
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    If p.eventId > 0 AndAlso ReplaceClassNameInCode(p.eventId, editClassId, newName) Then
                        'Значение - скрипт
                        p.Value = ReplaceClassNameInExPropery(p.Value, editClassId, newName)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso ReplaceClassNameInCode(ch.eventId, editClassId, newName) Then
                            'Значение - скрипт
                            ch.Value = ReplaceClassNameInExPropery(ch.Value, editClassId, newName)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso ReplaceClassNameInCode(ch.ThirdLevelEventId(child3Id), editClassId, newName) Then
                                'Значение - скрипт
                                ch.ThirdLevelProperties(child3Id) = ReplaceClassNameInExPropery(ch.ThirdLevelProperties(child3Id), editClassId, newName)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If ReplaceClassNameInCode(p.eventId, editClassId, newName) Then
                        'Значение - скрипт
                        p.Value = ReplaceClassNameInExPropery(p.Value, editClassId, newName)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                        If ch.eventId > 0 AndAlso ReplaceClassNameInCode(ch.eventId, editClassId, newName) Then
                            'Значение - скрипт
                            ch.Value = ReplaceClassNameInExPropery(ch.Value, editClassId, newName)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        ReplaceClassNameInFunctionHash(editClassId, newName)
        'В глобальных переменных
        ReplaceClassNameInVariables(editClassId, newName)
        'В событиях изменения свойства
        ReplaceClassNameInTracking(editClassId, newName)
    End Sub

    ''' <summary>
    ''' Заменяем имя указанного класса в глобальных переменных
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="newName">Новое имя</param>
    Private Sub ReplaceClassNameInVariables(ByVal editClassId As Integer, ByVal newName As String)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return

        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            Dim wasFoundGlobal As Boolean = False 'были ли изменения хоть в одном элементе массива переменной
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim dt() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    dt = .codeBox.CodeData
                End With
                If IsNothing(dt) Then Continue For
                'В элементе массива - исполняемый код
                Dim wasFound As Boolean = False 'была ли замена в данном элементе массива
                For line As Integer = 0 To dt.Count - 1
                    If IsNothing(dt(line).Code) Then Continue For
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        If dt(line).Code(wrd).classId = editClassId AndAlso dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                            'Строка найдена
                            Dim w As String = dt(line).Code(wrd).Word
                            dt(line).Code(wrd).Word = w.Replace(w.Trim, newName)
                            wasFound = True
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                            'Ищем строку вида "'ClassName.EventName'"
                            Dim w As String = dt(line).Code(wrd).Word
                            Dim pos As Integer = w.IndexOf(".")
                            If pos > 1 Then
                                Dim cName As String = w.Substring(1, pos - 1)
                                If mScript.mainClass(editClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                    'Начало строки - имя класса и точка
                                    If mScript.mainClass(editClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                        mScript.mainClass(editClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                        'За точкой - имя свойства или функции данного класса. Проводим замену имени класса
                                        dt(line).Code(wrd).Word = "'" + newName + w.Substring(pos)
                                        wasFound = True
                                    End If
                                End If
                            End If
                        End If
                    Next wrd
                Next line
                If wasFound Then
                    wasFoundGlobal = True
                    If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso dt.Count = 1 Then
                        'Исполняемая строка. Собираем измененную строку
                        Dim w As String = dt(0).StartingSpaces
                        For i As Integer = 0 To dt(0).Code.Count - 1
                            w += dt(0).Code(i).Word
                        Next i
                        w += dt(0).Comments
                        arrValues(arrId) = WrapString(w)
                    Else
                        'Скрипт или длинный текст. Собираем с внесенными изменениями
                        arrValues(arrId) = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(dt)
                    End If
                End If
            Next arrId
            If wasFoundGlobal Then
                'Сохраняем результат
                Dim var As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value
                var.arrValues = arrValues
            End If
        Next varId
    End Sub

    ''' <summary>
    ''' Заменяем имя указанного элемента класса в функциях писателя functionsHash
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="newName">Новое слово</param>
    Private Sub ReplaceClassNameInFunctionHash(ByVal editClassId As Integer, ByVal newName As String)
        If IsNothing(mScript.functionsHash) Then Return
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            'В функции хранится код в исполняемом виде - в ValueExecuteDt, и в виде для редактирования - в ValueDt (при этом второе может быть пустым, если функция объявлена из кода)
            'Первое быть пустым не должно
            Dim wasFound As Boolean = False
            If IsNothing(func.ValueExecuteDt) = False Then
                Dim edt As List(Of MatewScript.ExecuteDataType) = func.ValueExecuteDt
                For line As Integer = 0 To edt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(edt(line).Code) Then Continue For
                    For wrd As Integer = 0 To edt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        If edt(line).Code(wrd).classId = editClassId AndAlso edt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                            'Слово найдено
                            Dim w As String = edt(line).Code(wrd).Word
                            edt(line).Code(wrd).Word = w.Replace(w.Trim, newName)
                            wasFound = True
                        ElseIf edt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                            'Ищем строку вида "'ClassName.EventName'"
                            Dim w As String = edt(line).Code(wrd).Word
                            Dim pos As Integer = w.IndexOf(".")
                            If pos > 1 Then
                                Dim cName As String = w.Substring(1, pos - 1)
                                If mScript.mainClass(editClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                    'Начало строки - имя класса и точка
                                    If mScript.mainClass(editClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                        mScript.mainClass(editClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                        'За точкой - имя свойства или функции данного класса. Проводим замену имени класса
                                        edt(line).Code(wrd).Word = "'" + newName + w.Substring(pos)
                                        wasFound = True
                                    End If
                                End If
                            End If
                        End If
                    Next wrd
                Next line
            End If

            If Not wasFound Then Continue For 'если не найдено в ValueExecuteDt, то здесь тоже не будет
            If IsNothing(func.ValueDt) = False Then
                Dim dt() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To dt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(dt(line).Code) Then Continue For
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        If dt(line).Code(wrd).classId = editClassId AndAlso dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                            'Слово найдено
                            Dim w As String = dt(line).Code(wrd).Word
                            dt(line).Code(wrd).Word = w.Replace(w.Trim, newName)
                            wasFound = True
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                            'Ищем строку вида "'ClassName.EventName'"
                            Dim w As String = dt(line).Code(wrd).Word
                            Dim pos As Integer = w.IndexOf(".")
                            If pos > 1 Then
                                Dim cName As String = w.Substring(1, pos - 1)
                                If mScript.mainClass(editClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                    'Начало строки - имя класса и точка
                                    If mScript.mainClass(editClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                        mScript.mainClass(editClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                        'За точкой - имя свойства или функции данного класса. Проводим замену имени класса
                                        dt(line).Code(wrd).Word = "'" + newName + w.Substring(pos)
                                        wasFound = True
                                    End If
                                End If
                            End If
                        End If
                    Next wrd
                Next line
            End If
        Next funcId
    End Sub

    Private Sub ReplaceClassNameInTracking(ByVal editClassId As Integer, ByVal newName As String)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return

        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
            Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            If IsNothing(tr.propBeforeContent) = False Then
                If ReplaceClassNameInCode(tr.eventBeforeId, editClassId, newName) Then
                    Dim dt() As CodeTextBox.CodeDataType = tr.propBeforeContent
                    For line As Integer = 0 To dt.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(dt(line).Code) Then Continue For
                        For wrd As Integer = 0 To dt(line).Code.Count - 1
                            'перебираем каждое слово с строке
                            If dt(line).Code(wrd).classId = editClassId AndAlso dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                                'Слово найдено
                                Dim w As String = dt(line).Code(wrd).Word
                                dt(line).Code(wrd).Word = w.Replace(w.Trim, newName)
                            ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                                'Ищем строку вида "'ClassName.EventName'"
                                Dim w As String = dt(line).Code(wrd).Word
                                pos = w.IndexOf(".")
                                If pos > 1 Then
                                    Dim cName As String = w.Substring(1, pos - 1)
                                    If mScript.mainClass(editClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                        'Начало строки - имя класса и точка
                                        If mScript.mainClass(editClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                            mScript.mainClass(editClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                            'За точкой - имя свойства или функции данного класса. Проводим замену имени класса
                                            dt(line).Code(wrd).Word = "'" + newName + w.Substring(pos)
                                        End If
                                    End If
                                End If
                            End If
                        Next wrd
                    Next line
                End If
            End If

            If IsNothing(tr.propAfterContent) = False Then
                If ReplaceClassNameInCode(tr.eventAfterId, editClassId, newName) Then
                    Dim dt() As CodeTextBox.CodeDataType = tr.propAfterContent
                    For line As Integer = 0 To dt.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(dt(line).Code) Then Continue For
                        For wrd As Integer = 0 To dt(line).Code.Count - 1
                            'перебираем каждое слово с строке
                            If dt(line).Code(wrd).classId = editClassId AndAlso dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                                'Слово найдено
                                Dim w As String = dt(line).Code(wrd).Word
                                dt(line).Code(wrd).Word = w.Replace(w.Trim, newName)
                            ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                                'Ищем строку вида "'ClassName.EventName'"
                                Dim w As String = dt(line).Code(wrd).Word
                                pos = w.IndexOf(".")
                                If pos > 1 Then
                                    Dim cName As String = w.Substring(1, pos - 1)
                                    If mScript.mainClass(editClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                        'Начало строки - имя класса и точка
                                        If mScript.mainClass(editClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                            mScript.mainClass(editClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                            'За точкой - имя свойства или функции данного класса. Проводим замену имени класса
                                            dt(line).Code(wrd).Word = "'" + newName + w.Substring(pos)
                                        End If
                                    End If
                                End If
                            End If
                        Next wrd
                    Next line
                End If
            End If
        Next trId
    End Sub

    ''' <summary>
    ''' Производит замены имени элемента класса в исполняемой строке или сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="editClassId">класс, в котором меняется имя</param>
    ''' <param name="newName">Новое имя</param>
    ''' <returns>Исполнемую строку с внесенными изменениями</returns>
    Private Function ReplaceClassNameInExPropery(ByVal strXml As String, ByVal editClassId As Integer, ByVal newName As String) As String
        If String.IsNullOrEmpty(strXml) Then Return strXml
        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return strXml
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wrd As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                Dim w As CodeTextBox.EditWordType = cd(line).Code(wrd)
                If w.classId = editClassId AndAlso w.wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                    'заменяемый класс найден
                    Dim strWord As String = w.Word
                    cd(line).Code(wrd).Word = strWord.Replace(strWord.Trim, newName)
                    wasFound = True
                ElseIf w.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                    'Ищем строку вида 'ClassName.EventName'
                    Dim pos As Integer = w.Word.IndexOf(".")
                    If pos > 1 Then
                        Dim cName As String = w.Word.Substring(1, pos - 1)
                        If mScript.mainClass(editClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                            'Начало строки - имя класса и точка
                            If mScript.mainClass(editClassId).Properties.Keys.Contains(w.Word.Substring(pos + 1, w.Word.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                mScript.mainClass(editClassId).Functions.Keys.Contains(w.Word.Substring(pos + 1, w.Word.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                'За точкой - имя свойства или функции данного класса. Проводим замену имени класса
                                cd(line).Code(wrd).Word = "'" + newName + w.Word.Substring(pos)
                                wasFound = True
                            End If
                        End If
                    End If
                End If
            Next wrd
        Next line
        If wasFound Then
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                'собираем заново исполняемую строку
                strXml = cd(0).StartingSpaces
                For i As Integer = 0 To cd(0).Code.Count - 1
                    strXml += cd(0).Code(i).Word
                Next i
                strXml += cd(0).Comments
            Else
                'сериализаем код с внесенными изменениями
                strXml = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
            End If
        End If
        Return strXml
    End Function

    ''' <summary>
    ''' Заменят старое имя класса на новое в содержимом события
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="classId">Id класса</param>
    ''' <param name="newName">Любое новое имя</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function ReplaceClassNameInCode(ByVal evendId As Integer, ByVal classId As Integer, ByVal newName As String) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 0 To exData(line).Code.Count - 1
                If exData(line).Code(wordId).classId = classId AndAlso exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                    'Найден класс для замены
                    Dim w As String = exData(line).Code(wordId).Word
                    exData(line).Code(wordId).Word = w.Replace(w.Trim, newName)
                    wasFound = True
                ElseIf exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                    Dim w As String = exData(line).Code(wordId).Word
                    'Ищем строку вида "'ClassName.EventName'"
                    Dim pos As Integer = w.IndexOf(".")
                    If pos > 1 Then
                        Dim cName As String = w.Substring(1, pos - 1)
                        If mScript.mainClass(classId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                            'Начало строки - имя класса и точка
                            If mScript.mainClass(classId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                mScript.mainClass(classId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                'За точкой - имя свойства или функции данного класса. Проводим замену имени класса
                                exData(line).Code(wordId).Word = "'" + newName + w.Substring(pos)
                                wasFound = True
                            End If
                        End If
                    End If
                End If
            Next wordId
        Next line
        Return wasFound
    End Function

#End Region

#Region "Remove Class Name"
    ''' <summary>
    ''' При удалении класса писателя проверяет все свойства, функции и скрипты. Свойствах и функиях типа Element и таких же параметрах функций меняет тип на USUAL, 
    ''' в скриптах уменьшает номера последующих классов на 1, собирает инфо о наличии кода, содержащего удаляемый класс
    ''' </summary>
    ''' <param name="remClassId">Удаляемый класс</param>
    Public Sub RemoveClassNameInStruct(ByVal remClassId As Integer)
        Dim remName As String = mScript.mainClass(remClassId).Names.Last

        'Прочесывание структуры mainClass на предмет наличия функция и свойств, содержащих в качестве возвращаемого значения или параметра функции тип ELEMENT нашего удаляемого класса. Если будет такое найдено - очищаем формат
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств каждого класса
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso _
                        p.returnArray(0) = mScript.mainClass(remClassId).Names.Last Then
                        'тип = элемент удаляемого класса. Очищаем формат
                        Erase p.returnArray
                        p.returnType = MatewScript.ReturnFunctionEnum.RETURN_USUAL
                        Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                        Dim arrC() As Control = cPanelManager.dictDefContainers(classId).Controls.Find(pName, True)
                        If arrC.Count > 0 AndAlso arrC(0).GetType.Name = "ComboBoxEx" Then cPanelManager.ReplaceComboWithTextBox(arrC(0))
                    End If
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка свойств функций класса
                For fId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    Dim f As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(fId).Value
                    Dim fName As String = mScript.mainClass(classId).Functions.ElementAt(fId).Key
                    If f.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(f.returnArray) = False AndAlso f.returnArray.Count = 1 AndAlso _
                        f.returnArray(0) = mScript.mainClass(remClassId).Names.Last Then
                        'тип = элемент удаляемого класса. Очищаем формат
                        Erase f.returnArray
                        f.returnType = MatewScript.ReturnFunctionEnum.RETURN_USUAL
                    End If
                    If IsNothing(f.params) = False AndAlso f.params.Count > 0 Then
                        For pId As Integer = 0 To f.params.Count - 1
                            'Проверка параметров функции
                            If f.params(pId).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT AndAlso IsNothing(f.params(pId).EnumValues) = False AndAlso _
                                f.params(pId).EnumValues(0) = mScript.mainClass(remClassId).Names.Last Then
                                'тип параметра = элементу удаляемого класса. Очищаем формат
                                Erase f.params(pId).EnumValues
                                f.params(pId).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ANY
                            End If
                        Next pId
                    End If
                Next fId
            End If
        Next classId

        RemoveClassNameInScripts(remClassId) 'если удаляется последний класс, то на 1 уменьшать ничего не надо
    End Sub

    ''' <summary>
    ''' Проверяет наличие удаляемого класса во всех скриптах и составляет список найденных.
    ''' Если удаляется класс, то уменьшает на 1 значения всех классов с индексом большим, чем удаляемый
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    Private Sub RemoveClassNameInScripts(ByVal remClassId As Integer)
        'Изменения в коде
        dlgEntrancies.BeginNewEntrancies("Найден(ы) скрипт(ы), в которых встечаются объекты удаленного класса. Исправьте это вручную.", dlgEntrancies.EntranciesStyleEnum.Simple)
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                    If p.eventId > 0 AndAlso RemoveClassNameInCode(p.eventId, remClassId) Then
                        'Значение - скрипт
                        p.Value = RemoveClassNameInExPropery(p.Value, remClassId)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RemoveClassNameInCode(ch.eventId, remClassId) Then
                            'Значение - скрипт
                            ch.Value = RemoveClassNameInExPropery(ch.Value, remClassId)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso RemoveClassNameInCode(ch.ThirdLevelEventId(child3Id), remClassId) Then
                                'Значение - скрипт
                                ch.ThirdLevelProperties(child3Id) = RemoveClassNameInExPropery(ch.ThirdLevelProperties(child3Id), remClassId)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    dlgEntrancies.SetEntranceDefault(classId, -1, -1, mScript.mainClass(classId).Functions.ElementAt(pId).Key, CodeTextBox.EditWordTypeEnum.W_FUNCTION)

                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If RemoveClassNameInCode(p.eventId, remClassId) Then
                        'Значение - скрипт
                        p.Value = RemoveClassNameInExPropery(p.Value, remClassId)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                        If ch.eventId > 0 AndAlso RemoveClassNameInCode(ch.eventId, remClassId) Then
                            'Значение - скрипт
                            ch.Value = RemoveClassNameInExPropery(ch.Value, remClassId)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        RemoveClassNameInFunctionHash(remClassId)
        'В глобальных переменных
        RemoveClassNameInVariables(remClassId)
        'В событиях изменения свойств
        RemoveClassNameInTracking(remClassId)

        If dlgEntrancies.hasEntrancies Then dlgEntrancies.Show()
    End Sub

    ''' <summary>
    ''' Проверяет наличие удаляемого класса в глобальных переменных и составляет список найденных.
    ''' Если удаляется класс, то уменьшает на 1 значения всех классов с индексом большим, чем удаляемый
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    Private Sub RemoveClassNameInVariables(ByVal remClassId As Integer)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return

        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            Dim wasFoundGlobal As Boolean = False 'были ли изменения хоть в одном элементе массива переменной
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim dt() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    dt = .codeBox.CodeData
                End With
                If IsNothing(dt) Then Continue For
                'В элементе массива - исполняемый код
                Dim wasFound As Boolean = False 'была ли замена в данном элементе массива
                For line As Integer = 0 To dt.Count - 1
                    If IsNothing(dt(line).Code) Then Continue For
                    Dim wasEntrance As Boolean = False
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        If dt(line).Code(wrd).classId > remClassId AndAlso ret <> MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                            dt(line).Code(wrd).classId -= 1
                            wasFound = True
                        ElseIf dt(line).Code(wrd).classId = remClassId AndAlso wasEntrance = False Then
                            dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line)
                            wasEntrance = True
                            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then Exit For
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                            Dim w As String = dt(line).Code(wrd).Word
                            'Ищем строку вида 'ClassName.EventName'
                            Dim pos As Integer = w.IndexOf(".")
                            If pos > 1 Then
                                Dim cName As String = w.Substring(1, pos - 1)
                                If mScript.mainClass(remClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                    'Начало строки - имя класса и точка
                                    If mScript.mainClass(remClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                        mScript.mainClass(remClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                        'За точкой - имя свойства или функции данного класса
                                        dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line)
                                        wasEntrance = True
                                        If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next wrd
                Next line
                If wasFound AndAlso ret <> MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                    wasFoundGlobal = True
                    'Скрипт или длинный текст. Собираем с внесенными изменениями
                    arrValues(arrId) = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(dt)
                End If
            Next arrId
            If wasFoundGlobal Then
                'Сохраняем результат
                Dim var As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value
                var.arrValues = arrValues
            End If
        Next varId
    End Sub

    ''' <summary>
    ''' Проверяет наличие удаляемого класса в функциях писателя functionsHash и составляет список найденных.
    ''' Если удаляется класс, то уменьшает на 1 значения всех классов с индексом большим, чем удаляемый
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    Private Sub RemoveClassNameInFunctionHash(ByVal remClassId As Integer)
        If IsNothing(mScript.functionsHash) Then Return
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            'В функции хранится код в исполняемом виде - в ValueExecuteDt, и в виде для редактирования - в ValueDt (при этом второе может быть пустым, если функция объявлена из кода)
            'Первое быть пустым не должно
            Dim wasFound As Boolean = False
            If IsNothing(func.ValueExecuteDt) = False Then
                Dim edt As List(Of MatewScript.ExecuteDataType) = func.ValueExecuteDt
                For line As Integer = 0 To edt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(edt(line).Code) Then Continue For
                    For wrd As Integer = 0 To edt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        If edt(line).Code(wrd).classId > remClassId Then
                            edt(line).Code(wrd).classId = -1
                            wasFound = True
                        ElseIf edt(line).Code(wrd).classId = remClassId Then
                            wasFound = True
                        ElseIf edt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                            Dim w As String = edt(line).Code(wrd).Word
                            'Ищем строку вида 'ClassName.EventName'
                            Dim pos As Integer = w.IndexOf(".")
                            If pos > 1 Then
                                Dim cName As String = w.Substring(1, pos - 1)
                                If mScript.mainClass(remClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                    'Начало строки - имя класса и точка
                                    If mScript.mainClass(remClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                        mScript.mainClass(remClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                        'За точкой - имя свойства или функции данного класса
                                        wasFound = True
                                    End If
                                End If
                            End If
                        End If
                    Next wrd
                Next line
            End If

            If Not wasFound Then Continue For 'если не найдено в ValueExecuteDt, то здесь тоже не будет
            If IsNothing(func.ValueDt) = False Then
                Dim dt() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To dt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(dt(line).Code) Then Continue For
                    Dim wasEntrance As Boolean = False
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        If dt(line).Code(wrd).classId > remClassId Then
                            dt(line).Code(wrd).classId -= 1
                        ElseIf dt(line).Code(wrd).classId = remClassId AndAlso wasEntrance = False Then
                            dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line)
                            wasEntrance = True
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                            Dim w As String = dt(line).Code(wrd).Word
                            'Ищем строку вида 'ClassName.EventName'
                            Dim pos As Integer = w.IndexOf(".")
                            If pos > 1 Then
                                Dim cName As String = w.Substring(1, pos - 1)
                                If mScript.mainClass(remClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                    'Начало строки - имя класса и точка
                                    If mScript.mainClass(remClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                        mScript.mainClass(remClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                        'За точкой - имя свойства или функции данного класса
                                        dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line)
                                        wasEntrance = True
                                    End If
                                End If
                            End If
                        End If
                    Next wrd
                Next line
            End If
        Next funcId
    End Sub

    Private Sub RemoveClassNameInTracking(ByVal remClassId As Integer)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return

        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
        Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            If IsNothing(tr.propBeforeContent) = False Then
                Dim dt() As CodeTextBox.CodeDataType = (tr.propBeforeContent)
                For line As Integer = 0 To dt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(dt(line).Code) Then Continue For
                    Dim wasEntrance As Boolean = False
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        If dt(line).Code(wrd).classId > remClassId Then
                            dt(line).Code(wrd).classId -= 1
                        ElseIf dt(line).Code(wrd).classId = remClassId AndAlso wasEntrance = False Then
                            dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                            wasEntrance = True
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                            Dim w As String = dt(line).Code(wrd).Word
                            'Ищем строку вида 'ClassName.EventName'
                            pos = w.IndexOf(".")
                            If pos > 1 Then
                                Dim cName As String = w.Substring(1, pos - 1)
                                If mScript.mainClass(remClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                    'Начало строки - имя класса и точка
                                    If mScript.mainClass(remClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                        mScript.mainClass(remClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                        'За точкой - имя свойства или функции данного класса
                                        dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                                        wasEntrance = True
                                    End If
                                End If
                            End If
                        End If
                    Next wrd
                Next line
            End If

            If IsNothing(tr.propAfterContent) = False Then
                Dim dt() As CodeTextBox.CodeDataType = (tr.propAfterContent)
                For line As Integer = 0 To dt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(dt(line).Code) Then Continue For
                    Dim wasEntrance As Boolean = False
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        If dt(line).Code(wrd).classId > remClassId Then
                            dt(line).Code(wrd).classId -= 1
                        ElseIf dt(line).Code(wrd).classId = remClassId AndAlso wasEntrance = False Then
                            dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                            wasEntrance = True
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                            Dim w As String = dt(line).Code(wrd).Word
                            'Ищем строку вида 'ClassName.EventName'
                            pos = w.IndexOf(".")
                            If pos > 1 Then
                                Dim cName As String = w.Substring(1, pos - 1)
                                If mScript.mainClass(remClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                                    'Начало строки - имя класса и точка
                                    If mScript.mainClass(remClassId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                        mScript.mainClass(remClassId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                        'За точкой - имя свойства или функции данного класса
                                        dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                                        wasEntrance = True
                                    End If
                                End If
                            End If
                        End If
                    Next wrd
                Next line
            End If
        Next trId
    End Sub

    ''' <summary>
    ''' Проверяет наличие удаляемого класса в исполняемой строке или сериализованном коде/длинном тексте и составляет список найденных.
    ''' Уменьшает на 1 значения всех классов с индексом большим, чем удаляемый
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="remClassId">класс, в котором меняется имя</param>
    ''' <returns>Исполнемую строку с внесенными изменениями</returns>
    Private Function RemoveClassNameInExPropery(ByVal strXml As String, ByVal remClassId As Integer) As String
        If String.IsNullOrEmpty(strXml) Then Return strXml
        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return strXml
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            Dim wasEntrance As Boolean = False
            For wrd As Integer = 0 To cd(line).Code.Count - 1
                'перебираем слова в строках кода
                Dim w As CodeTextBox.EditWordType = cd(line).Code(wrd)
                If w.classId > remClassId Then
                    cd(line).Code(wrd).classId -= 1
                    wasFound = True
                ElseIf w.classId = remClassId AndAlso wasEntrance = False Then
                    dlgEntrancies.SetSeekPosInfo(line)
                    dlgEntrancies.NewEntrance()
                    wasEntrance = True
                ElseIf w.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                    'Ищем строку вида 'ClassName.EventName'
                    Dim pos As Integer = w.Word.IndexOf(".")
                    If pos > 1 Then
                        Dim cName As String = w.Word.Substring(1, pos - 1)
                        If mScript.mainClass(remClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                            'Начало строки - имя класса и точка
                            If mScript.mainClass(remClassId).Properties.Keys.Contains(w.Word.Substring(pos + 1, w.Word.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                mScript.mainClass(remClassId).Functions.Keys.Contains(w.Word.Substring(pos + 1, w.Word.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                'За точкой - имя свойства или функции данного класса
                                dlgEntrancies.SetSeekPosInfo(line)
                                dlgEntrancies.NewEntrance()
                                wasEntrance = True
                            End If
                        End If
                    End If
                End If
            Next wrd
        Next line

        If wasFound Then
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                'поиск был только для выявления ошибок
                Return strXml
            Else
                'сериализуем код с внесенными изменениями
                strXml = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
            End If
        End If
        Return strXml
    End Function

    ''' <summary>
    ''' Проверяет наличие удаляемого класса в содержимом события.
    ''' Уменшает на 1 значения всех классов с индексом большим, чем удаляемый
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="classId">Id класса</param>
    ''' <returns>True если были произведены замены или найдено вхождение удаляемого элемента, False - если нечего не найдено</returns>
    Private Function RemoveClassNameInCode(ByVal evendId As Integer, ByVal classId As Integer) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 0 To exData(line).Code.Count - 1
                If exData(line).Code(wordId).classId > classId Then
                    exData(line).Code(wordId).classId -= 1
                    wasFound = True
                ElseIf exData(line).Code(wordId).classId = classId Then
                    wasFound = True
                ElseIf exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                    'Ищем строку вида 'ClassName.EventName'
                    Dim w As String = exData(line).Code(wordId).Word
                    Dim pos As Integer = w.IndexOf(".")
                    If pos > 1 Then
                        Dim cName As String = w.Substring(1, pos - 1)
                        If mScript.mainClass(classId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                            'Начало строки - имя класса и точка
                            If mScript.mainClass(classId).Properties.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) OrElse _
                                mScript.mainClass(classId).Functions.Keys.Contains(w.Substring(pos + 1, w.Length - pos - 2), StringComparer.CurrentCultureIgnoreCase) Then
                                'За точкой - имя свойства или функции данного класса
                                wasFound = True
                            End If
                        End If
                    End If
                End If
            Next wordId
        Next line
        Return wasFound
    End Function
#End Region

#Region "Replace Element Name"
    ''' <summary>
    ''' Изменяет имя указанного элемента (функция, свойство, переменная) во всех свойствах, функциях, перменных со скриптами и т. д.
    ''' </summary>
    ''' <param name="editClassId">Id класса для замены</param>
    ''' <param name="oldName">Старое имя элемента (в кавычках)</param>
    ''' <param name="newName">Новое имя элемента (в кавычках)</param>
    ''' <param name="elementType">тип элемента</param>
    Public Sub ReplaceElementNameInStruct(ByVal editClassId As Integer, oldName As String, newName As String, ByVal elementType As CodeTextBox.EditWordTypeEnum)

        If elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
            'Прочесывание структуры mainClass на предмет наличия свойств типа ELEMENT нашего класса и имеющим значение = Variable. Если найдем - переименовываем
            Dim varName As String = WrapString(oldName)
            For classId As Integer = 0 To mScript.mainClass.Count - 1
                If IsNothing(mScript.mainClass(classId).Properties) = False Then
                    'проверка свойств каждого класса
                    For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                        Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                        If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso p.returnArray(0) = "Variable" Then
                            'тип = элемент нужного класса
                            Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                            'Свойство по умолчанию
                            If String.Compare(p.Value, varName, True) = 0 Then
                                p.Value = WrapString(newName)
                            End If
                            If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                                'Свойства 2 уровня
                                For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                    Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                    If String.Compare(ch.Value, varName, True) = 0 Then
                                        ch.Value = WrapString(newName)
                                    End If
                                    If IsNothing(ch.ThirdLevelProperties) = False Then
                                        'Свойства 3 уровня
                                        For child3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                            If String.Compare(ch.ThirdLevelProperties(child3Id), varName, True) = 0 Then
                                                ch.ThirdLevelProperties(child3Id) = WrapString(newName)
                                            End If
                                        Next child3Id
                                    End If
                                Next child2Id
                            End If
                        End If
                    Next pId
                End If
            Next classId

            'В сохраненных действиях
            Dim classAct As Integer = mScript.mainClassHash("A")
            Dim classLoc As Integer = mScript.mainClassHash("L")
            If actionsRouter.hasSavedActions Then
                For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                    Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                    Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                    Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                    If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                    For aId As Integer = 0 To arrProp.Count - 1
                        For pId As Integer = 0 To arrProp(aId).Count - 1
                            Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                            Dim pName As String = arrProp(aId).ElementAt(pId).Key
                            Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                            If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                                pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                            If pp.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(pp.returnArray) = False AndAlso pp.returnArray.Count = 1 AndAlso _
                                pp.returnArray(0) = "Variable" Then
                                'тип = элемент нужного класса
                                If String.Compare(ch.Value, varName, True) = 0 Then
                                    ch.Value = WrapString(newName)
                                End If
                            End If
                        Next pId
                    Next aId
                Next i
            End If
        End If

        ReplaceElementNameInScripts(editClassId, oldName, newName, elementType)
    End Sub

    ''' <summary>
    ''' Изменяет имя указанного элемента (функция, свойство, переменная) во всех скриптах
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    ''' <param name="elementType">тип элемента</param>
    Private Sub ReplaceElementNameInScripts(ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String, ByVal elementType As CodeTextBox.EditWordTypeEnum)
        If elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE Then editClassId = -1
        dlgEntrancies.BeginNewEntrancies("Найдены строки, которые, возможно, также следует переименовать. Сделайте это самостоятельно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        'Изменения в коде
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    If p.eventId > 0 AndAlso ReplaceElementNameInCode(p.eventId, editClassId, oldName, newName, elementType) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                        p.Value = ReplaceElementNameInExPropery(p.Value, editClassId, oldName, newName, elementType)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso ReplaceElementNameInCode(ch.eventId, editClassId, oldName, newName, elementType) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                            ch.Value = ReplaceElementNameInExPropery(ch.Value, editClassId, oldName, newName, elementType)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso ReplaceElementNameInCode(ch.ThirdLevelEventId(child3Id), editClassId, oldName, newName, elementType) Then
                                'Значение - скрипт
                                dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                                ch.ThirdLevelProperties(child3Id) = ReplaceElementNameInExPropery(ch.ThirdLevelProperties(child3Id), editClassId, oldName, newName, elementType)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If ReplaceElementNameInCode(p.eventId, editClassId, oldName, newName, elementType) Then
                        'Значение - скрипт
                        dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_FUNCTION)
                        p.Value = ReplaceElementNameInExPropery(p.Value, editClassId, oldName, newName, elementType)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                        If ch.eventId > 0 AndAlso ReplaceElementNameInCode(ch.eventId, editClassId, oldName, newName, elementType) Then
                            'Значение - скрипт
                            dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                            ch.Value = ReplaceElementNameInExPropery(ch.Value, editClassId, oldName, newName, elementType)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        ReplaceElementNameInFunctionHash(editClassId, oldName, newName, elementType)
        'В глобальных переменных
        ReplaceElementNameInVariables(editClassId, oldName, newName, elementType)
        'В событиях изменения свойств
        ReplaceElementNameInTracking(editClassId, oldName, newName, elementType)

        mScript.FillFuncAndPropHash()

        If dlgEntrancies.hasEntrancies Then dlgEntrancies.Show()
    End Sub

    ''' <summary>
    ''' Заменяем имя указанного элемента (функция, свойство, переменная) в глобальных переменных
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    ''' <param name="elementType">тип элемента</param>
    Private Sub ReplaceElementNameInVariables(ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String, ByVal elementType As CodeTextBox.EditWordTypeEnum)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return

        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            Dim wasFoundGlobal As Boolean = False 'были ли изменения хоть в одном элементе массива переменной
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim dt() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    dt = .codeBox.CodeData
                End With
                If IsNothing(dt) Then Continue For
                'В элементе массива - исполняемый код
                Dim wasFound As Boolean = False 'была ли замена в данном элементе массива
                For line As Integer = 0 To dt.Count - 1
                    If IsNothing(dt(line).Code) Then Continue For
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        Dim strWord As String = dt(line).Code(wrd).Word
                        If dt(line).Code(wrd).classId = editClassId AndAlso dt(line).Code(wrd).wordType = elementType Then
                            'Строка найдена
                            If String.Compare(strWord.Trim, oldName, True) = 0 Then
                                dt(line).Code(wrd).Word = strWord.Replace(strWord.Trim, newName)
                                wasFound = True
                            End If
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(oldName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                            (elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                             strWord.IndexOf(oldName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                            Dim res As String = SeekElementNameInString(dt, strWord, editClassId, oldName, newName, elementType, line, wrd)
                            If res = "#Entrance" Then
                                dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line)
                            ElseIf res.Length > 0 Then
                                wasFound = True
                                dt(line).Code(wrd).Word = res
                            End If
                        End If
                    Next wrd
                Next line
                If wasFound Then
                    wasFoundGlobal = True
                    If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso dt.Count = 1 Then
                        'Исполняемая строка. Собираем измененную строку
                        Dim w As String = dt(0).StartingSpaces
                        For i As Integer = 0 To dt(0).Code.Count - 1
                            w += dt(0).Code(i).Word
                        Next i
                        w += dt(0).Comments
                        arrValues(arrId) = WrapString(w)
                    Else
                        'Скрипт или длинный текст. Собираем с внесенными изменениями
                        arrValues(arrId) = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(dt)
                    End If
                End If
            Next arrId
            If wasFoundGlobal Then
                'Сохраняем результат
                Dim var As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value
                var.arrValues = arrValues
            End If
        Next varId
    End Sub

    ''' <summary>
    ''' Заменяем имя указанного элемента (функция, свойство, переменная) в функциях писателя functionsHash
    ''' </summary>
    ''' <param name="editClassId">Класс для редактирования</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    ''' <param name="elementType">тип элемента</param>
    Private Sub ReplaceElementNameInFunctionHash(ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String, ByVal elementType As CodeTextBox.EditWordTypeEnum)
        If IsNothing(mScript.functionsHash) Then Return
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            'В функции хранится код в исполняемом виде - в ValueExecuteDt, и в виде для редактирования - в ValueDt (при этом второе может быть пустым, если функция объявлена из кода)
            'Первое быть пустым не должно
            Dim wasFound As Boolean = False
            If IsNothing(func.ValueExecuteDt) = False Then
                Dim edt As List(Of MatewScript.ExecuteDataType) = func.ValueExecuteDt
                For line As Integer = 0 To edt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(edt(line).Code) Then Continue For
                    For wrd As Integer = 0 To edt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        Dim strWord As String = edt(line).Code(wrd).Word
                        If edt(line).Code(wrd).classId = editClassId AndAlso edt(line).Code(wrd).wordType = elementType Then
                            'Слово найдено
                            If String.Compare(strWord.Trim, oldName, True) = 0 Then
                                edt(line).Code(wrd).Word = strWord.Replace(strWord.Trim, newName)
                                wasFound = True
                            End If
                        ElseIf edt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(oldName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                            (elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                             strWord.IndexOf(oldName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                            Dim res As String = SeekElementNameInStringEx(edt, strWord, editClassId, oldName, newName, elementType, line, wrd)
                            If res = "#Entrance" Then
                                wasFound = True
                            ElseIf res.Length > 0 Then
                                wasFound = True
                                edt(line).Code(wrd).Word = res
                            End If
                        End If
                    Next wrd
                Next line
            End If

            If Not wasFound Then Continue For 'если не найдено в ValueExecuteDt, то здесь тоже не будет
            If IsNothing(func.ValueDt) = False Then
                Dim dt() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To dt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(dt(line).Code) Then Continue For
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        Dim strWord As String = dt(line).Code(wrd).Word
                        If dt(line).Code(wrd).classId = editClassId AndAlso dt(line).Code(wrd).wordType = elementType Then
                            'Слово найдено
                            If String.Compare(strWord.Trim, oldName, True) = 0 Then
                                dt(line).Code(wrd).Word = strWord.Replace(strWord.Trim, newName)
                                wasFound = True
                            End If
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(oldName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                            (elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                             strWord.IndexOf(oldName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                            Dim res As String = SeekElementNameInString(dt, strWord, editClassId, oldName, newName, elementType, line, wrd)
                            If res = "#Entrance" Then
                                dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line)
                            ElseIf res.Length > 0 Then
                                wasFound = True
                                dt(line).Code(wrd).Word = res
                            End If
                        End If
                    Next wrd
                Next line
            End If
        Next funcId
    End Sub

    Private Sub ReplaceElementNameInTracking(ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String, ByVal elementType As CodeTextBox.EditWordTypeEnum)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return

        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
            Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            If IsNothing(tr.propBeforeContent) = False Then
                If ReplaceElementNameInCodeTracking(tr.eventBeforeId, editClassId, oldName, newName, elementType) Then
                    Dim dt() As CodeTextBox.CodeDataType = tr.propBeforeContent
                    For line As Integer = 0 To dt.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(dt(line).Code) Then Continue For
                        For wrd As Integer = 0 To dt(line).Code.Count - 1
                            'перебираем каждое слово с строке
                            Dim strWord As String = dt(line).Code(wrd).Word
                            If dt(line).Code(wrd).classId = editClassId AndAlso dt(line).Code(wrd).wordType = elementType Then
                                'Слово найдено
                                If String.Compare(strWord.Trim, oldName, True) = 0 Then
                                    dt(line).Code(wrd).Word = strWord.Replace(strWord.Trim, newName)
                                End If
                            ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(oldName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                                (elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                                 strWord.IndexOf(oldName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                                Dim res As String = SeekElementNameInString(dt, strWord, editClassId, oldName, newName, elementType, line, wrd)
                                If res = "#Entrance" Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                                ElseIf res.Length > 0 Then
                                    dt(line).Code(wrd).Word = res
                                End If
                            End If
                        Next wrd
                    Next line
                End If
            End If

            If IsNothing(tr.propAfterContent) = False Then
                'dlgEntrancies.SetEntranceDefault(classId, -1, -1, oldName, elementType, -2, -1, "", -1, frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                If ReplaceElementNameInCodeTracking(tr.eventAfterId, editClassId, oldName, newName, elementType) Then
                    Dim dt() As CodeTextBox.CodeDataType = tr.propAfterContent
                    For line As Integer = 0 To dt.Count - 1
                        'Перебираем каждую линию
                        If IsNothing(dt(line).Code) Then Continue For
                        For wrd As Integer = 0 To dt(line).Code.Count - 1
                            'перебираем каждое слово с строке
                            Dim strWord As String = dt(line).Code(wrd).Word
                            If dt(line).Code(wrd).classId = editClassId AndAlso dt(line).Code(wrd).wordType = elementType Then
                                'Слово найдено
                                If String.Compare(strWord.Trim, oldName, True) = 0 Then
                                    dt(line).Code(wrd).Word = strWord.Replace(strWord.Trim, newName)
                                End If
                            ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(oldName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                                (elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                                 strWord.IndexOf(oldName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                                Dim res As String = SeekElementNameInString(dt, strWord, editClassId, oldName, newName, elementType, line, wrd)
                                If res = "#Entrance" Then
                                    dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, "", -2, frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                                ElseIf res.Length > 0 Then
                                    dt(line).Code(wrd).Word = res
                                End If
                            End If
                        Next wrd
                    Next line
                End If
            End If
        Next trId
    End Sub

    ''' <summary>
    ''' Заменят старое имя элемента (функция, свойство, переменная) на новое в содержимом события. От ReplaceElementNameInCode отличается тем, что не создает точкек вхождения, а вместо этого возвращает True
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="editClassId">Id класса</param>
    ''' <param name="oldName">Старое имя</param>
    ''' <param name="newName">Новое имя</param>
    ''' <param name="elementType">Тип элемента</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function ReplaceElementNameInCodeTracking(ByVal evendId As Integer, ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String, _
                                            ByVal elementType As CodeTextBox.EditWordTypeEnum) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 0 To exData(line).Code.Count - 1
                Dim strWord As String = exData(line).Code(wordId).Word
                If exData(line).Code(wordId).classId = editClassId AndAlso exData(line).Code(wordId).wordType = elementType Then
                    'Найден класс для замены
                    If String.Compare(strWord.Trim, oldName, True) = 0 Then
                        exData(line).Code(wordId).Word = strWord.Replace(strWord.Trim, newName)
                        wasFound = True
                    End If
                ElseIf exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(oldName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                    (elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                     strWord.IndexOf(oldName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                    Dim res As String = SeekElementNameInStringEx(exData, strWord, editClassId, oldName, newName, elementType, line, wordId)
                    If res = "#Entrance" Then
                        wasFound = True
                    ElseIf res.Length > 0 Then
                        wasFound = True
                        exData(line).Code(wordId).Word = res
                    End If
                End If
            Next wordId
        Next line
        Return wasFound
    End Function

    ''' <summary>
    ''' Производит замены имени элемента (функция, свойство, переменная) в исполняемой строке или сериализованном коде/длинном тексте
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="editClassId">класс, в котором меняется имя</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя елемента</param>
    ''' <param name="elementType">тип элемента</param>
    ''' <returns>Исполнемую строку с внесенными изменениями</returns>
    Private Function ReplaceElementNameInExPropery(ByVal strXml As String, ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String, _
                                                   ByVal elementType As CodeTextBox.EditWordTypeEnum) As String
        If String.IsNullOrEmpty(strXml) Then Return strXml
        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return strXml
        Dim wasFound As Boolean = False 'были ли изменения
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wrd As Integer = 0 To cd(line).Code.Count - 1
                'перибираем слова в строках кода
                Dim w As CodeTextBox.EditWordType = cd(line).Code(wrd)
                Dim strWord As String = w.Word
                If w.classId = editClassId AndAlso w.wordType = elementType Then
                    'заменяемый класс найден
                    If String.Compare(strWord.Trim, oldName, True) = 0 Then
                        cd(line).Code(wrd).Word = strWord.Replace(strWord.Trim, newName)
                        wasFound = True
                    End If
                ElseIf w.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(oldName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                    (elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                     strWord.IndexOf(oldName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                    Dim res As String = SeekElementNameInString(cd, strWord, editClassId, oldName, newName, elementType, line, wrd)
                    If res = "#Entrance" Then
                        dlgEntrancies.SetSeekPosInfo(line)
                        dlgEntrancies.NewEntrance()
                    ElseIf res.Length > 0 Then
                        wasFound = True
                        cd(line).Code(wrd).Word = res
                    End If
                End If
            Next wrd
        Next line
        If wasFound Then
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING AndAlso cd.Count = 1 Then
                'собираем заново исполняемую строку
                strXml = cd(0).StartingSpaces
                For i As Integer = 0 To cd(0).Code.Count - 1
                    strXml += cd(0).Code(i).Word
                Next i
                strXml += cd(0).Comments
            Else
                'сериализаем код с внесенными изменениями
                strXml = questEnvironment.codeBoxShadowed.codeBox.SerializeCodeData(cd)
            End If
        End If
        Return strXml
    End Function

    ''' <summary>
    ''' Заменят старое имя элемента (функция, свойство, переменная) на новое в содержимом события
    ''' </summary>
    ''' <param name="evendId">Id события</param>
    ''' <param name="editClassId">Id класса</param>
    ''' <param name="oldName">Старое имя</param>
    ''' <param name="newName">Новое имя</param>
    ''' <param name="elementType">Тип элемента</param>
    ''' <returns>True если были произведены замены, False - если нечего не найдено</returns>
    Private Function ReplaceElementNameInCode(ByVal evendId As Integer, ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String, _
                                             ByVal elementType As CodeTextBox.EditWordTypeEnum) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False
        If mScript.eventRouter.IsExistsAndNotEmpty(evendId, exData) = False Then Return False

        Dim wasFound As Boolean = False
        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 0 To exData(line).Code.Count - 1
                Dim strWord As String = exData(line).Code(wordId).Word
                If exData(line).Code(wordId).classId = editClassId AndAlso exData(line).Code(wordId).wordType = elementType Then
                    'Найден класс для замены
                    If String.Compare(strWord.Trim, oldName, True) = 0 Then
                        exData(line).Code(wordId).Word = strWord.Replace(strWord.Trim, newName)
                        wasFound = True
                    End If
                ElseIf exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(oldName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                    (elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                     strWord.IndexOf(oldName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                    Dim res As String = SeekElementNameInStringEx(exData, strWord, editClassId, oldName, newName, elementType, line, wordId)
                    If res = "#Entrance" Then
                        dlgEntrancies.SetSeekPosInfo(line)
                        dlgEntrancies.NewEntrance()
                    ElseIf res.Length > 0 Then
                        wasFound = True
                        exData(line).Code(wordId).Word = res
                    End If
                End If
            Next wordId
        Next line
        Return wasFound
    End Function


    ''' <summary>
    ''' Внутрення функция для функций группы ReplaceElementNameX. Определяет является ли слово скрипта формата CodeTextBox.CodeDataType строкой вида 'ClassName.EventName' / '[Var.]VarName[x]'.
    ''' Если является, то возвращает слово с измененным именем элемента.
    ''' </summary>
    ''' <param name="cd">Скрипт</param>
    ''' <param name="strWord">Проверяемое слово</param>
    ''' <param name="editClassId">Класс, элемент которого меняется</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя элемента</param>
    ''' <param name="elementType">Тип элемента</param>
    ''' <param name="line">Текущая линия в cd</param>
    ''' <param name="wordId">Id слова strWord в текущей линии</param>
    Private Function SeekElementNameInString(ByRef cd() As CodeTextBox.CodeDataType, ByVal strWord As String, ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String, _
                                             ByVal elementType As CodeTextBox.EditWordTypeEnum, ByVal line As Integer, ByVal wordId As Integer) As String
        'Ищем строку вида 'ClassName.EventName'
        'Если поиск переменной, то это может быть '[Var.]VarName[x]'
        Dim pos As Integer = strWord.IndexOf(".")
        If pos > 1 Then
            'ClassName.EventName' / 'Var.VarName' / 'V.VarName[x]'
            Dim cName As String = strWord.Substring(1, pos - 1)
            If elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
                If strWord.StartsWith("'V.", StringComparison.CurrentCultureIgnoreCase) OrElse strWord.StartsWith("'Var.", StringComparison.CurrentCultureIgnoreCase) Then
                    'Var.VarName' / 'V.VarName[x]'
                    Dim pos2 As Integer = strWord.LastIndexOf("["c)
                    Dim vName As String = ""
                    If pos2 >= 0 Then
                        'V.VarName[x]'
                        vName = strWord.Substring(pos + 1, pos2 - pos - 1)
                        Return strWord.Substring(0, pos + 1) + newName + strWord.Substring(pos2)
                    Else
                        'Var.VarName'
                        vName = strWord.Substring(pos + 1, strWord.Length - pos - 2)
                        Return strWord.Substring(0, pos + 1) + newName + "'"
                    End If
                End If
            Else
                'ClassName.EventName'
                If mScript.mainClass(editClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                    'Начало строки - имя класса и точка, затем - старое имя. Проводим замену элемента
                    Return strWord.Substring(0, pos + 1) + newName + "'"
                End If
            End If
        ElseIf String.Compare(strWord, "'" + oldName + "'", True) = 0 Then
            'полное совпадение имени элемента и текущей строки
            'ElementName' / 'VarName'
            If wordId = 1 AndAlso (cd(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_WRAP OrElse cd(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT OrElse _
                                   cd(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION) Then
                Return "'" + newName + "'"
            Else
                'dlgEntrancies.SetSeekLine(line)
                'dlgEntrancies.NewEntrance()
                Return "#Entrance"
            End If
        ElseIf elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.StartsWith("'" + oldName + "[") Then
            'VarName[x]'
            Return "'" + newName + strWord.Substring(oldName.Length + 1)
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Внутрення функция для функций группы ReplaceElementNameX. Определяет является ли слово скрипта формата MatewScript.ExecuteDataType строкой вида 'ClassName.EventName' / '[Var.]VarName[x]'.
    ''' Если является, то возвращает слово с измененным именем элемента.
    ''' </summary>
    ''' <param name="exDt">Скрипт</param>
    ''' <param name="strWord">Проверяемое слово</param>
    ''' <param name="editClassId">Класс, элемент которого меняется</param>
    ''' <param name="oldName">Старое имя элемента</param>
    ''' <param name="newName">Новое имя элемента</param>
    ''' <param name="elementType">Тип элемента</param>
    ''' <param name="line">Текущая линия в exDt</param>
    ''' <param name="wordId">Id слова strWord в текущей линии</param>
    Private Function SeekElementNameInStringEx(ByRef exDt As List(Of MatewScript.ExecuteDataType), ByVal strWord As String, ByVal editClassId As Integer, ByVal oldName As String, ByVal newName As String, _
                                         ByVal elementType As CodeTextBox.EditWordTypeEnum, ByVal line As Integer, ByVal wordId As Integer) As String
        'Ищем строку вида 'ClassName.EventName'
        'Если поиск переменной, то это может быть '[Var.]VarName[x]'
        Dim pos As Integer = strWord.IndexOf(".")
        If pos > 1 Then
            'ClassName.EventName' / 'Var.VarName' / 'V.VarName[x]'
            Dim cName As String = strWord.Substring(1, pos - 1)
            If elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
                If strWord.StartsWith("'V.", StringComparison.CurrentCultureIgnoreCase) OrElse strWord.StartsWith("'Var.", StringComparison.CurrentCultureIgnoreCase) Then
                    'Var.VarName' / 'V.VarName[x]'
                    Dim pos2 As Integer = strWord.LastIndexOf("["c)
                    Dim vName As String = ""
                    If pos2 >= 0 Then
                        'V.VarName[x]'
                        vName = strWord.Substring(pos + 1, pos2 - pos - 1)
                        Return strWord.Substring(0, pos + 1) + newName + strWord.Substring(pos2)
                    Else
                        'Var.VarName'
                        vName = strWord.Substring(pos + 1, strWord.Length - pos - 2)
                        Return strWord.Substring(0, pos + 1) + newName + "'"
                    End If
                End If
            Else
                'ClassName.EventName'
                If mScript.mainClass(editClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                    'Начало строки - имя класса и точка, затем - старое имя. Проводим замену элемента
                    Return strWord.Substring(0, pos + 1) + newName + "'"
                End If
            End If
        ElseIf String.Compare(strWord, "'" + oldName + "'", True) = 0 Then
            'полное совпадение имени элемента и текущей строки
            'ElementName' / 'VarName'
            If wordId = 1 AndAlso (exDt(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_WRAP OrElse exDt(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT OrElse _
                                   exDt(line).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION) Then
                Return "'" + newName + "'"
            Else
                'dlgEntrancies.SetSeekLine(line)
                'dlgEntrancies.NewEntrance()
                Return "#Entrance"
            End If
        ElseIf elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.StartsWith("'" + oldName + "[") Then
            'VarName[x]'
            Return "'" + newName + strWord.Substring(oldName.Length + 1)
        End If
        Return ""
    End Function
#End Region

#Region "Check Element Name"
    ''' <summary>
    ''' Проверяет наличие удаляемого элемента (функция, свойство, переменная) во всех свойствах, функциях, перменных со скриптами и т. д.
    ''' </summary>
    ''' <param name="remClassId">Id класса для замены</param>
    ''' <param name="remName">Имя удаляемого элемента</param>
    ''' <param name="elementType">тип элемента</param>
    ''' <param name="afterRemoving">Запускается ли процедура после удаления элемента или же это просто поиск</param>
    Public Sub CheckElementNameInStruct(ByVal remClassId As Integer, remName As String, ByVal elementType As CodeTextBox.EditWordTypeEnum, Optional ByVal afterRemoving As Boolean = True)
        If remClassId < -1 Then remClassId = -1
        If afterRemoving Then
            dlgEntrancies.BeginNewEntrancies("Найден(ы) скрипт(ы), в которых встечаются удаленные элементы. Исправьте это вручную.", dlgEntrancies.EntranciesStyleEnum.Simple)
        Else
            dlgEntrancies.BeginNewEntrancies("Результаты поиска переменной " + remName + ". Включены также те результаты, значение которых определить невозможно.", dlgEntrancies.EntranciesStyleEnum.Simple)
        End If

        If elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
            'Прочесывание структуры mainClass на предмет наличия свойств типа ELEMENT нашего класса и имеющим значение = Variable. Если найдем - переименовываем
            Dim varName As String = WrapString(remName)
            For classId As Integer = 0 To mScript.mainClass.Count - 1
                If IsNothing(mScript.mainClass(classId).Properties) = False Then
                    'проверка свойств каждого класса
                    For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                        Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                        If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray.Count = 1 AndAlso p.returnArray(0) = "Variable" Then
                            'тип = элемент нужного класса
                            Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                            'Свойство по умолчанию
                            If String.Compare(p.Value, varName, True) = 0 Then
                                dlgEntrancies.NewEntrance(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, remName, 0)
                            End If
                            If IsNothing(mScript.mainClass(classId).ChildProperties) = False Then
                                'Свойства 2 уровня
                                For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                                    Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                                    If String.Compare(ch.Value, varName, True) = 0 Then
                                        dlgEntrancies.NewEntrance(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, remName, 0)
                                    End If
                                    If IsNothing(ch.ThirdLevelProperties) = False Then
                                        'Свойства 3 уровня
                                        For child3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                            If String.Compare(ch.ThirdLevelProperties(child3Id), varName, True) = 0 Then
                                                dlgEntrancies.NewEntrance(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, 0, remName, 0)
                                            End If
                                        Next child3Id
                                    End If
                                Next child2Id
                            End If
                        End If
                    Next pId
                End If
            Next classId

            'В сохраненных действиях
            Dim classAct As Integer = mScript.mainClassHash("A")
            Dim classLoc As Integer = mScript.mainClassHash("L")
            If actionsRouter.hasSavedActions Then
                For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                    Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                    Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                    Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                    If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                    For aId As Integer = 0 To arrProp.Count - 1
                        For pId As Integer = 0 To arrProp(aId).Count - 1
                            Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                            Dim pName As String = arrProp(aId).ElementAt(pId).Key
                            Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                            If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                                pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                            If pp.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(pp.returnArray) = False AndAlso pp.returnArray.Count = 1 AndAlso _
                                pp.returnArray(0) = "Variable" Then
                                'тип = элемент нужного класса
                                If String.Compare(ch.Value, varName, True) = 0 Then
                                    dlgEntrancies.NewEntrance(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId, 0, remName, 0)
                                End If
                            End If
                        Next pId
                    Next aId
                Next i
            End If
        End If

        CheckElementNameInScripts(remClassId, remName, elementType)
        If dlgEntrancies.hasEntrancies Then dlgEntrancies.Show()
    End Sub

    ''' <summary>
    ''' Проверяет наличие удаляемого элемента (функция, свойство, переменная) во всех скриптах и составляет список найденных.
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    ''' <param name="remName">Имя удаляемого элемента</param>
    ''' <param name="remType">Тип удаляемого элемента</param>
    Private Sub CheckElementNameInScripts(ByVal remClassId As Integer, ByVal remName As String, ByVal remType As CodeTextBox.EditWordTypeEnum)
        'Изменения в коде
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'Перебираем все классы
            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                'проверка свойств
                For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    'Перебираем все значения свойств
                    'Свойство по умолчанию
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(pId).Value
                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(pId).Key
                    dlgEntrancies.SetEntranceDefault(classId, -1, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                    If p.eventId > 0 AndAlso RemoveElementNameInCode(p.eventId, remClassId, remName, remType) Then
                        'Значение - скрипт
                        RemoveElementNameInExPropery(p.Value, remClassId, remName, remType)
                    End If

                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then Continue For
                    For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                        'Перебираем элементы 2 порядка
                        dlgEntrancies.SetEntranceDefault(classId, child2Id, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(pName)
                        If ch.eventId > 0 AndAlso RemoveElementNameInCode(ch.eventId, remClassId, remName, remType) Then
                            'Значение - скрипт
                            RemoveElementNameInExPropery(ch.Value, remClassId, remName, remType)
                        End If
                        If IsNothing(ch.ThirdLevelEventId) Then Continue For

                        For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                            dlgEntrancies.SetEntranceDefault(classId, child2Id, child3Id, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)
                            'Перебираем все элемты 3 порядка
                            If ch.ThirdLevelEventId(child3Id) > 0 AndAlso RemoveElementNameInCode(ch.ThirdLevelEventId(child3Id), remClassId, remName, remType) Then
                                'Значение - скрипт
                                RemoveElementNameInExPropery(ch.ThirdLevelProperties(child3Id), remClassId, remName, remType)
                            End If
                        Next child3Id
                    Next child2Id
                Next pId
            End If

            If IsNothing(mScript.mainClass(classId).Functions) = False Then
                'проверка функций
                For pId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    dlgEntrancies.SetEntranceDefault(classId, -1, -1, mScript.mainClass(classId).Functions.ElementAt(pId).Key, CodeTextBox.EditWordTypeEnum.W_FUNCTION)

                    'Перебираем все функции на предмет содержания кода расширения (или же просто кода функции пользователя)
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(pId).Value
                    If p.eventId <= 0 Then Continue For
                    Dim pName As String = mScript.mainClass(classId).Functions.ElementAt(pId).Key
                    If RemoveElementNameInCode(p.eventId, remClassId, remName, remType) Then
                        'Значение - скрипт
                        RemoveElementNameInExPropery(p.Value, remClassId, remName, remType)
                    End If
                Next pId
            End If
        Next classId

        'В сохраненных действиях
        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")
        If actionsRouter.hasSavedActions Then
            For i As Integer = 0 To actionsRouter.lstActions.Count - 1
                Dim locName As String = actionsRouter.lstActions.ElementAt(i).Key
                Dim locId As Integer = GetSecondChildIdByName(locName, mScript.mainClass(classLoc).ChildProperties)
                Dim arrProp() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(i).Value
                If IsNothing(arrProp) OrElse arrProp.Count = 0 Then Continue For
                For aId As Integer = 0 To arrProp.Count - 1
                    For pId As Integer = 0 To arrProp(aId).Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = arrProp(aId).ElementAt(pId).Value
                        Dim pName As String = arrProp(aId).ElementAt(pId).Key
                        Dim pp As MatewScript.PropertiesInfoType = mScript.mainClass(classAct).Properties(pName)
                        If pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pp.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR OrElse _
                            pp.Hidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY Then Continue For

                        dlgEntrancies.SetEntranceDefault(classAct, aId, -1, pName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, locId)
                        If ch.eventId > 0 AndAlso RemoveElementNameInCode(ch.eventId, remClassId, remName, remType) Then
                            'Значение - скрипт
                            RemoveElementNameInExPropery(ch.Value, remClassId, remName, remType)
                        End If
                    Next pId
                Next aId
            Next i
        End If

        'В функциях
        RemoveElementNameInFunctionHash(remClassId, remName, remType)
        'В глобальных переменных
        RemoveElementNameInVariables(remClassId, remName, remType)
        'В событиях изменения свойства
        RemoveElementNameInTracking(remClassId, remName, remType)

        mScript.FillFuncAndPropHash()
    End Sub

    ''' <summary>
    ''' Проверяет наличие удаляемого элемента (функция, свойство, переменная) в глобальных переменных и составляет список найденных.
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    ''' <param name="remName">Имя удаляемого элемента</param>
    ''' <param name="remType">Тип удаляемого элемента</param>
    Private Sub RemoveElementNameInVariables(ByVal remClassId As Integer, ByVal remName As String, ByVal remType As CodeTextBox.EditWordTypeEnum)
        If IsNothing(mScript.csPublicVariables.lstVariables) Then Return

        For varId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
            'Перебираем все переменные
            Dim varName As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Key
            Dim arrValues() As String = mScript.csPublicVariables.lstVariables.ElementAt(varId).Value.arrValues
            If IsNothing(arrValues) Then Continue For
            For arrId As Integer = 0 To arrValues.Count - 1
                'Перебираем массив переменной
                Dim ret As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(arrValues(arrId))
                If ret = MatewScript.ContainsCodeEnum.NOT_CODE Then Continue For
                Dim dt() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = 0
                    .codeBox.IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                    .Text = mScript.PrepareStringToPrint(arrValues(arrId), Nothing, False)
                    dt = .codeBox.CodeData
                End With
                If IsNothing(dt) Then Continue For
                'В элементе массива - исполняемый код
                For line As Integer = 0 To dt.Count - 1
                    If IsNothing(dt(line).Code) Then Continue For
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        Dim strWord As String = dt(line).Code(wrd).Word
                        If dt(line).Code(wrd).classId = remClassId AndAlso dt(line).Code(wrd).wordType = remType Then
                            If String.Compare(dt(line).Code(wrd).Word.Trim, remName, True) = 0 Then
                                dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line, remName, GetWordPosInLine(dt(line), wrd))
                            End If
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(remName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                            (remType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                             strWord.IndexOf(remName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                            'Ищем строку вида 'ClassName.EventName'. Если поиск переменной, то это может быть '[Var.]VarName[x]'
                            If SeekElementNameInStringRemoving(dt, strWord, remClassId, remName, remType, line, wrd) Then
                                dlgEntrancies.NewEntrance(-2, varId, arrId, varName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, -1, line, remName, GetWordPosInLine(dt(line), wrd))
                            End If
                        End If
                    Next wrd
                Next line
            Next arrId
        Next varId
    End Sub

    ''' <summary>
    ''' Проверяет наличие удаляемого элемента (функция, свойство, переменная) в функциях писателя functionsHash и составляет список найденных.
    ''' </summary>
    ''' <param name="remClassId">Класс для редактирования</param>
    ''' <param name="remName">Имя удаляемого элемента</param>
    ''' <param name="remType">Тип удаляемого элемента</param>
    Private Sub RemoveElementNameInFunctionHash(ByVal remClassId As Integer, ByVal remName As String, ByVal remType As CodeTextBox.EditWordTypeEnum)
        If IsNothing(mScript.functionsHash) Then Return
        For funcId As Integer = 0 To mScript.functionsHash.Count - 1
            Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
            Dim func As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
            If IsNothing(func.ValueDt) = False Then
                Dim dt() As CodeTextBox.CodeDataType = func.ValueDt
                For line As Integer = 0 To dt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(dt(line).Code) Then Continue For
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        Dim strWord As String = dt(line).Code(wrd).Word
                        If dt(line).Code(wrd).classId = remClassId AndAlso dt(line).Code(wrd).wordType = remType Then
                            If String.Compare(dt(line).Code(wrd).Word.Trim, remName, True) = 0 Then
                                dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line, remName, GetWordPosInLine(dt(line), wrd))
                            End If
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(remName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                            (remType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                             strWord.IndexOf(remName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                            'Ищем строку вида 'ClassName.EventName'. Если поиск переменной, то это может быть '[Var.]VarName[x]'
                            If SeekElementNameInStringRemoving(dt, strWord, remClassId, remName, remType, line, wrd) Then
                                dlgEntrancies.NewEntrance(-3, funcId, -1, funcName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, -1, line, remName, GetWordPosInLine(dt(line), wrd))
                            End If
                        End If
                    Next wrd
                Next line
            End If
        Next funcId
    End Sub

    Private Sub RemoveElementNameInTracking(ByVal remClassId As Integer, ByVal remName As String, ByVal remType As CodeTextBox.EditWordTypeEnum)
        If IsNothing(mScript.trackingProperties.lstTrackingProperties) OrElse mScript.trackingProperties.lstTrackingProperties.Count = 0 Then Return

        For trId As Integer = 0 To mScript.trackingProperties.lstTrackingProperties.Count - 1
            Dim tr As cTrackingProperties.TrackingPropertyData = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Value
            Dim strKey As String = mScript.trackingProperties.lstTrackingProperties.ElementAt(trId).Key, classId As Integer, propName As String
            Dim pos As Integer = strKey.IndexOf("."c)
            If pos = -1 Then Continue For
            classId = mScript.mainClassHash(strKey.Substring(0, pos))
            propName = strKey.Substring(pos + 1)

            If IsNothing(tr.propBeforeContent) = False Then
                Dim dt() As CodeTextBox.CodeDataType = tr.propBeforeContent
                For line As Integer = 0 To dt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(dt(line).Code) Then Continue For
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        Dim strWord As String = dt(line).Code(wrd).Word
                        If dt(line).Code(wrd).classId = remClassId AndAlso dt(line).Code(wrd).wordType = remType Then
                            If String.Compare(dt(line).Code(wrd).Word.Trim, remName, True) = 0 Then
                                dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, remName, GetWordPosInLine(dt(line), wrd), frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                            End If
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(remName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                            (remType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                             strWord.IndexOf(remName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                            'Ищем строку вида 'ClassName.EventName'. Если поиск переменной, то это может быть '[Var.]VarName[x]'
                            If SeekElementNameInStringRemoving(dt, strWord, remClassId, remName, remType, line, wrd) Then
                                dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, remName, GetWordPosInLine(dt(line), wrd), frmMainEditor.trackingcodeEnum.EVENT_BEFORE)
                            End If
                        End If
                    Next wrd
                Next line
            End If

            If IsNothing(tr.propAfterContent) = False Then
                Dim dt() As CodeTextBox.CodeDataType = tr.propAfterContent
                For line As Integer = 0 To dt.Count - 1
                    'Перебираем каждую линию
                    If IsNothing(dt(line).Code) Then Continue For
                    For wrd As Integer = 0 To dt(line).Code.Count - 1
                        'перебираем каждое слово с строке
                        Dim strWord As String = dt(line).Code(wrd).Word
                        If dt(line).Code(wrd).classId = remClassId AndAlso dt(line).Code(wrd).wordType = remType Then
                            If String.Compare(dt(line).Code(wrd).Word.Trim, remName, True) = 0 Then
                                dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, remName, GetWordPosInLine(dt(line), wrd), frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                            End If
                        ElseIf dt(line).Code(wrd).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(remName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                            (remType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                             strWord.IndexOf(remName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                            'Ищем строку вида 'ClassName.EventName'. Если поиск переменной, то это может быть '[Var.]VarName[x]'
                            If SeekElementNameInStringRemoving(dt, strWord, remClassId, remName, remType, line, wrd) Then
                                dlgEntrancies.NewEntrance(classId, -1, -1, propName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1, line, remName, GetWordPosInLine(dt(line), wrd), frmMainEditor.trackingcodeEnum.EVENT_AFTER)
                            End If
                        End If
                    Next wrd
                Next line
            End If
        Next trId
    End Sub

    ''' <summary>
    ''' Проверяет наличие удаляемого элемента (функция, свойство, переменная) в исполняемой строке или сериализованном коде/длинном тексте и составляет список найденных.
    ''' </summary>
    ''' <param name="strXml">Исполняемая строка или код</param>
    ''' <param name="remClassId">класс, в котором меняется имя</param>
    ''' <param name="remName">Имя удаляемого элемента</param>
    ''' <param name="remType">Тип удаляемого элемента</param>
    Private Sub RemoveElementNameInExPropery(ByVal strXml As String, ByVal remClassId As Integer, ByVal remName As String, ByVal remType As CodeTextBox.EditWordTypeEnum)
        If String.IsNullOrEmpty(strXml) Then Return
        Dim cd() As CodeTextBox.CodeDataType
        Dim ret As MatewScript.ContainsCodeEnum
        'получаем структуру кода cd
        With questEnvironment.codeBoxShadowed.codeBox
            .Text = ""
            ret = mScript.IsPropertyContainsCode(strXml)
            If ret = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                .IsTextBlockByDefault = False
                .Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            Else
                .IsTextBlockByDefault = (ret = MatewScript.ContainsCodeEnum.LONG_TEXT)
                .LoadCodeFromProperty(strXml)
            End If
            cd = .CodeData
        End With

        If IsNothing(cd) OrElse cd.Count = 0 Then Return
        For line As Integer = 0 To cd.Count - 1
            'Перебираем все строки кода
            If IsNothing(cd(line).Code) Then Continue For
            For wrd As Integer = 0 To cd(line).Code.Count - 1
                'перебираем слова в строках кода
                Dim w As CodeTextBox.EditWordType = cd(line).Code(wrd)
                Dim strWord As String = w.Word
                If w.classId = remClassId AndAlso w.wordType = remType Then
                    If String.Compare(w.Word.Trim, remName, True) = 0 Then
                        dlgEntrancies.SetSeekPosInfo(line, remName, GetWordPosInLine(cd(line), wrd))
                        dlgEntrancies.NewEntrance()
                    End If
                ElseIf w.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(remName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                    (remType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                     strWord.IndexOf(remName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                    'Ищем строку вида 'ClassName.EventName'. Если поиск переменной, то это может быть '[Var.]VarName[x]'
                    If SeekElementNameInStringRemoving(cd, strWord, remClassId, remName, remType, line, wrd) Then
                        dlgEntrancies.SetSeekPosInfo(line, remName, GetWordPosInLine(cd(line), wrd))
                        dlgEntrancies.NewEntrance()
                    End If
                End If
            Next wrd
        Next line
    End Sub

    ''' <summary>
    ''' Проверяет наличие удаляемого элемента (функция, свойство, переменная) в содержимом события.
    ''' </summary>
    ''' <param name="eventId">Id события</param> 
    ''' <param name="remClassId">Id класса</param>
    ''' <param name="remName">Имя удаляемого элемента</param>
    ''' <param name="remType">Тип удаляемого элемента</param>
    ''' <returns>True если найдено вхождение удаляемого элемента, False - если ничего не найдено</returns>
    Private Function RemoveElementNameInCode(ByVal eventId As Integer, ByVal remClassId As Integer, ByVal remName As String, ByVal remType As CodeTextBox.EditWordTypeEnum) As Boolean
        Dim exData As List(Of MatewScript.ExecuteDataType) = Nothing
        'If mScript.eventRouter.lstEvents.TryGetValue(evendId, exData) = False OrElse IsNothing(exData) Then Return False
        If mScript.eventRouter.IsExistsAndNotEmpty(eventId, exData) = False Then Return False

        For line As Integer = 0 To exData.Count - 1
            If IsNothing(exData(line).Code) Then Continue For
            For wordId As Integer = 0 To exData(line).Code.Count - 1
                Dim strWord As String = exData(line).Code(wordId).Word
                If exData(line).Code(wordId).classId = remClassId AndAlso exData(line).Code(wordId).wordType = remType Then
                    If String.Compare(strWord.Trim, remName, True) = 0 Then
                        Return True 'Найдено 
                    End If
                ElseIf exData(line).Code(wordId).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING AndAlso strWord.Length > 1 AndAlso (strWord.EndsWith(remName + "'", StringComparison.CurrentCultureIgnoreCase) OrElse _
                                    (remType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.Chars(strWord.Length - 2) = "]"c AndAlso _
                                     strWord.IndexOf(remName + "[", StringComparison.CurrentCultureIgnoreCase) > -1)) Then
                    'Ищем строку вида 'ClassName.EventName'. Если поиск переменной, то это может быть '[Var.]VarName[x]'
                    If SeekElementNameInStringRemovingEx(exData, strWord, remClassId, remName, remType, line, wordId) Then
                        Return True 'Найдено 
                    End If
                End If
            Next wordId
        Next line
        Return False
    End Function

    ''' <summary>
    ''' Внутрення функция для функций группы RemoveElementNameX. Определяет является ли слово скрипта формата CodeTextBox.CodeDataType строкой вида 'ClassName.EventName' / '[Var.]VarName[x]'.
    ''' Если является, то возвращает True.
    ''' </summary>
    ''' <param name="cd">Скрипт</param>
    ''' <param name="strWord">Проверяемое слово</param>
    ''' <param name="remClassId">Класс, элемент которого меняется</param>
    ''' <param name="remName">Старое имя элемента</param>
    ''' <param name="elementType">Тип элемента</param>
    ''' <param name="line">Текущая линия в cd</param>
    ''' <param name="wordId">Id слова strWord в текущей линии</param>
    Private Function SeekElementNameInStringRemoving(ByRef cd() As CodeTextBox.CodeDataType, ByVal strWord As String, ByVal remClassId As Integer, ByVal remName As String, _
                                             ByVal elementType As CodeTextBox.EditWordTypeEnum, ByVal line As Integer, ByVal wordId As Integer) As Boolean
        'Ищем строку вида 'ClassName.EventName'
        'Если поиск переменной, то это может быть '[Var.]VarName[x]'
        Dim pos As Integer = strWord.IndexOf(".")
        If pos > 1 Then
            'ClassName.EventName' / 'Var.VarName' / 'V.VarName[x]'
            Dim cName As String = strWord.Substring(1, pos - 1)
            If elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
                If strWord.StartsWith("'V.", StringComparison.CurrentCultureIgnoreCase) OrElse strWord.StartsWith("'Var.", StringComparison.CurrentCultureIgnoreCase) Then
                    'Var.VarName' / 'V.VarName[x]'
                    Return True
                End If
            Else
                'ClassName.EventName'
                If mScript.mainClass(remClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                    'Начало строки - имя класса и точка, затем - старое имя.
                    Return True
                End If
            End If
        ElseIf String.Compare(strWord, "'" + remName + "'", True) = 0 Then
            'полное совпадение имени элемента и текущей строки
            'ElementName' / 'VarName'
            Return True
        ElseIf elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.StartsWith("'" + remName + "[") Then
            'VarName[x]'
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' Внутрення функция для функций группы RemoveElementNameX. Определяет является ли слово скрипта формата MatewScript.ExecuteDataType строкой вида 'ClassName.EventName' / '[Var.]VarName[x]'.
    ''' Если является, то возвращает True.
    ''' </summary>
    ''' <param name="exDt">Скрипт</param>
    ''' <param name="strWord">Проверяемое слово</param>
    ''' <param name="remClassId">Класс, элемент которого меняется</param>
    ''' <param name="remName">Старое имя элемента</param>
    ''' <param name="elementType">Тип элемента</param>
    ''' <param name="line">Текущая линия в exDt</param>
    ''' <param name="wordId">Id слова strWord в текущей линии</param>
    Private Function SeekElementNameInStringRemovingEx(ByRef exDt As List(Of MatewScript.ExecuteDataType), ByVal strWord As String, ByVal remClassId As Integer, ByVal remName As String, _
                                         ByVal elementType As CodeTextBox.EditWordTypeEnum, ByVal line As Integer, ByVal wordId As Integer) As Boolean
        'Ищем строку вида 'ClassName.EventName'
        'Если поиск переменной, то это может быть '[Var.]VarName[x]'
        Dim pos As Integer = strWord.IndexOf(".")
        If pos > 1 Then
            'ClassName.EventName' / 'Var.VarName' / 'V.VarName[x]'
            Dim cName As String = strWord.Substring(1, pos - 1)
            If elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
                If strWord.StartsWith("'V.", StringComparison.CurrentCultureIgnoreCase) OrElse strWord.StartsWith("'Var.", StringComparison.CurrentCultureIgnoreCase) Then
                    'Var.VarName' / 'V.VarName[x]'
                    Return True
                End If
            Else
                'ClassName.EventName'
                If mScript.mainClass(remClassId).Names.Contains(cName, StringComparer.CurrentCultureIgnoreCase) Then
                    'Начало строки - имя класса и точка, затем - старое имя.
                    Return True
                End If
            End If
        ElseIf String.Compare(strWord, "'" + remName + "'", True) = 0 Then
            'полное совпадение имени элемента и текущей строки
            'ElementName' / 'VarName'
            Return True
        ElseIf elementType = CodeTextBox.EditWordTypeEnum.W_VARIABLE AndAlso strWord.StartsWith("'" + remName + "[") Then
            'VarName[x]'
            Return True
        End If
        Return False
    End Function
#End Region
End Class