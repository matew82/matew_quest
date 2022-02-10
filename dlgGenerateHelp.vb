Imports System.Windows.Forms

Public Class dlgGenerateHelp
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Public Sub GenerateHelpForFunction(ByVal propName As String, ByVal classId As Integer, ByRef propValue As MatewScript.PropertiesInfoType)
        Dim qt As Char = Chr(34)
        Dim sb As New System.Text.StringBuilder
        'получаем версию Internet Explorer
        Dim ieVersion As FileVersionInfo = FileVersionInfo.GetVersionInfo("c:\windows\system32\ieframe.dll")
        Dim strIEversion As String = "11"
        If IsNothing(ieVersion) = False Then
            strIEversion = ieVersion.FileMajorPart.ToString
        End If

        Dim strReturn As String = "", strReturnType As String = ""
        GetFunctionDescriptionReturnData(propValue.Description, strReturn, strReturnType)

        Dim retFormat As String = strReturnType

        If retFormat.Length = 0 Then
            retFormat = "any"
            Select Case propValue.returnType
                Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                    retFormat = "bool"
                Case MatewScript.ReturnFunctionEnum.RETURN_COLOR
                    retFormat = "color"
                Case MatewScript.ReturnFunctionEnum.RETURN_FUNCTION
                    retFormat = "function"
                Case MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION
                    retFormat = "string/html"
                Case MatewScript.ReturnFunctionEnum.RETURN_ENUM
                    retFormat = "enum"
                Case MatewScript.ReturnFunctionEnum.RETURN_EVENT
                    retFormat = "script"
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO
                    retFormat = "path"
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS
                    retFormat = "path"
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_JS
                    retFormat = "path"
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE
                    retFormat = "path"
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_TEXT
                    retFormat = "path"
                Case MatewScript.ReturnFunctionEnum.RETURN_ELEMENT
                    If IsNothing(propValue.returnArray) = False AndAlso propValue.returnArray.Count > 0 Then
                        retFormat = propValue.returnArray(0).ToLower
                        If retFormat = "magic" Then
                            retFormat = "magic book"
                        ElseIf retFormat = "ability" Then
                            retFormat = "ability set"
                        End If
                    End If
            End Select
        End If

        With sb
            'Шапка
            .AppendLine("<!DOCTYPE html>")
            .AppendLine("<html>")
            .AppendLine("<head>")
            .AppendFormat("<meta http-equiv={0}Content-Type{0} content={0}text/html; charset=windows-1251{0}/>" & vbNewLine, qt)
            .AppendFormat("<meta http-equiv={0}X-UA-Compatible{0} content={0}IE={1}{0} />" & vbNewLine, qt, strIEversion)
            .AppendFormat("<link rel={0}stylesheet{0} type={0}text/css{0} href={0}../help.css{0}/>" & vbNewLine, qt)
            .AppendLine("</head>")
            .AppendLine()
            'Тело
            .AppendLine("<body>")
            'Таблица заголовка
            If String.IsNullOrEmpty(propValue.EditorCaption) OrElse propValue.EditorCaption = propName Then
                .AppendFormat("    <h2 class={0}elementName{0}>{1}</h2>" & vbNewLine, qt, propName)
            Else
                .AppendFormat("    <h2 class={0}elementName{0}>{1} ({2})</h2>" & vbNewLine, qt, propName, propValue.EditorCaption)
            End If
            .AppendFormat("<table class={0}headerTable{0}>" & vbNewLine, qt)
            .AppendLine("<tr>")
            .AppendLine("    <td>")
            Dim aStr As String = ""
            Dim realParamsCount As Integer = 0
            If IsNothing(propValue.params) = False AndAlso propValue.params.Count > 0 Then
                realParamsCount = propValue.params.Count
                For i As Integer = 0 To realParamsCount - 1
                    Dim p As MatewScript.paramsType = propValue.params(i)
                    Dim isRequired As Boolean = False
                    If i <= propValue.paramsMin - 1 Then isRequired = True
                    If p.Type = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY Then
                        aStr += "<i>[Еще параметр1, Еще параметр2, Еще параметр3...]</i>"
                    Else
                        If isRequired Then
                            aStr += p.Name
                        Else
                            aStr += "<i>[" + p.Name + "]</i>"
                        End If
                    End If
                    If i < realParamsCount - 1 Then aStr += ", "
                Next
                aStr = "(" + aStr + ")"
            End If

            .AppendFormat("        <code><i>{0}</i> = {1}.<b><font color=red>{2}</font></b>{3}</code>" & vbNewLine, retFormat, mScript.mainClass(classId).Names.Last, propName, aStr)
            .AppendLine("    </td>")
            .AppendLine("    <td width='250px'>")
            .AppendLine("        <table align='right'>")
            .AppendLine("        <tr>")
            .AppendFormat("            <td class={0}typeCaption{0}>Класс</td>" & vbNewLine, qt)

            If mScript.mainClass(classId).Names.Count = 1 Then
                aStr = mScript.mainClass(classId).Names(0)
            Else
                aStr = mScript.mainClass(classId).Names.Last & " ("
                For i As Integer = mScript.mainClass(classId).Names.Count - 2 To 0 Step -1
                    aStr += mScript.mainClass(classId).Names(i)
                    If i > 0 Then aStr += ", "
                Next i
                aStr += ")"
            End If
            .AppendFormat("                <td class={0}typeInfo{0}>{1}</td>" & vbNewLine, qt, aStr)
            .AppendLine("            </tr>")
            .AppendLine("            <tr>")
            .AppendFormat("            <td class={0}typeCaption{0}>Тип</td>" & vbNewLine, qt)
            .AppendFormat("            <td class={0}typeInfo{0}>Функция</td>" & vbNewLine, qt)
            .AppendLine("        </tr>")
            .AppendLine("        </table>")
            .AppendLine("    </td>")
            .AppendLine("</tr>")
            .AppendLine("</table>")
            .AppendLine()
            'Описание
            .AppendFormat("<p class={0}description{0}>" & vbNewLine, qt)
            .AppendLine("    " & GetFunctionDescription(propValue.Description))
            .AppendLine("</p>")


            Dim isNumericEnum As Boolean = False 'если тип Enum, то является ли он численным
            If realParamsCount > 0 Then
                'Значение параметров у функции нет  - не печатаем
                .AppendLine()
                .AppendLine("<center>")
                .AppendLine("    <h2>Параметры функции</h2>")
                .AppendFormat("    <table class={0}tableParamsFunction{0}>" & vbNewLine, qt)
                .AppendLine("        <tr>")
                .AppendLine("            <th>параметр</th>")
                .AppendLine("            <th>тип</th>")
                .AppendLine("            <th>значение</th>")
                .AppendLine("        </tr>")
                For i As Integer = 0 To realParamsCount - 1
                    Dim p As MatewScript.paramsType = propValue.params(i)
                    Dim isRequired As Boolean = False
                    If i <= propValue.paramsMin - 1 Then isRequired = True
                    Dim isEnum As Boolean = False
                    isNumericEnum = False
                    Dim propArray() As String = GetParamArray(p)
                    If propArray.Count > 0 Then
                        isEnum = True
                        Dim isNum As Boolean = False
                        Dim retStr As String = mScript.PrepareStringToPrint(propArray(0), Nothing, False)
                        If retStr.Length > 0 Then
                            If Integer.TryParse(retStr.Chars(0), Nothing) Then
                                isNum = True
                            ElseIf retStr.Length > 1 AndAlso retStr.Chars(0) = "-"c AndAlso Integer.TryParse(retStr.Chars(1), Nothing) Then
                                isNum = True
                            End If
                        End If
                        If isNum Then
                            'enum численный
                            isNumericEnum = True
                        End If
                    End If

                    .AppendLine("        <tr>")
                    If isRequired Then
                        .AppendFormat("            <td><code>{0}</code></td>" & vbNewLine, p.Name)
                    Else
                        .AppendFormat("            <td><code>[<i>{0}</code>]</i></td>" & vbNewLine, p.Name)
                    End If
                    .AppendFormat("            <td><code>{0}</code></td>" & vbNewLine, GetFunctionParamFormatName(propValue.params, i, isNumericEnum))
                    If Not isEnum Then
                        .AppendFormat("            <td>{0}</td>" & vbNewLine, p.Description)
                    Else
                        aStr = p.Description + "<br/>" + vbNewLine + "                <ul>" + vbNewLine
                        For j As Integer = 0 To propArray.Count - 1
                            Dim curVal As String = mScript.PrepareStringToPrint(propArray(j), Nothing, False)
                            aStr += "                    <li>" + curVal + "</li>" + vbNewLine
                        Next
                        aStr += "                </ul>"
                        .AppendFormat("            <td>{0}" & vbNewLine, aStr)
                        .AppendLine("            </td>")
                    End If
                    .AppendLine("        </tr>")
                Next

                .AppendLine("    </table>")
            Else
                .AppendLine("    <center>")
            End If
            If String.IsNullOrEmpty(strReturn) = False Then
                .AppendLine()
                'Возвращает
                .AppendLine("    <h2>Возвращает</h2>")
                .AppendFormat("    <table class={0}tableParams{0} id={0}functionReturn{0}>" & vbNewLine, qt)
                .AppendLine("        <tr>")
                .AppendFormat("            <td><i><code>{0}</code></i></td>" & vbNewLine, retFormat)
                .AppendFormat("            <td>{0}</td>" & vbNewLine, strReturn)
                .AppendLine("        </tr>")
                .AppendLine("    </table>")
                .AppendLine()
            ElseIf retFormat = "enum" AndAlso IsNothing(propValue.returnArray) = False AndAlso propValue.returnArray.Count > 0 Then
                .AppendLine()
                'возвращаемый enum
                .AppendLine("    <h2>Возвращает</h2>")
                .AppendFormat("    <table class={0}tableParams{0} id={0}functionReturn{0}>" & vbNewLine, qt)


                isNumericEnum = False
                Dim isNum As Boolean = False
                Dim retStr As String = mScript.PrepareStringToPrint(propValue.returnArray(0), Nothing, False)
                If retStr.Length > 0 Then
                    If Integer.TryParse(retStr.Chars(0), Nothing) Then
                        isNum = True
                    ElseIf retStr.Length > 1 AndAlso retStr.Chars(0) = "-"c AndAlso Integer.TryParse(retStr.Chars(1), Nothing) Then
                        isNum = True
                    End If
                End If
                If isNum Then
                    'enum численный
                    isNumericEnum = True
                End If

                If isNumericEnum Then
                    'enum нумерованный
                    For i As Integer = 0 To propValue.returnArray.Count - 1
                        Dim rVal As String = UnWrapString(propValue.returnArray(i)).Trim
                        Dim rNum As String = i.ToString
                        Dim pos As Integer = rVal.IndexOf(" "c)
                        If pos > -1 Then
                            rNum = rVal.Substring(0, pos)
                            rVal = rVal.Substring(pos + 1).TrimStart
                            If rVal.StartsWith("- ") Then rVal = rVal.Substring(2).TrimStart
                        End If

                        .AppendLine("        <tr>")
                        .AppendFormat("            <td><i><code>{0}</code></i></td>" & vbNewLine, rNum)
                        .AppendFormat("            <td>{0}</td>" & vbNewLine, rVal)
                        .AppendLine("        </tr>")
                    Next
                Else
                    'enum не нумерованный
                    For i As Integer = 0 To propValue.returnArray.Count - 1
                        .AppendLine("        <tr>")
                        .AppendFormat("            <td>{0}</td>" & vbNewLine, UnWrapString(propValue.returnArray(i)))
                        .AppendLine("        </tr>")
                    Next
                End If
                .AppendLine("    </table>")
                .AppendLine()
            End If



            'Завершение
            'заготовка для вставки примера скрипта
            .AppendLine()
            .AppendLine("<!--")
            .AppendLine("<h2>Заготовка для вставки примера скрипта</h2>")
            .AppendFormat("<div class={0}codeSampleTitle{0}>Название примера</div>" & vbNewLine, qt)
            .AppendFormat("<div class={0}codeSampleDescription{0}>Описание примера скрипта, вводная к игровой ситуации где это может пригодиться.</div>" & vbNewLine, qt)
            .AppendFormat("<div class={0}codeSample{0}>" & vbNewLine, qt)
            .AppendLine("    Непосредственно скрипт")
            .AppendFormat("    <div class={0}codeSampleNote{0}>... прерываемый различными замечаниями. Так можно разделить, например, скрипты из разных функций.</div>" & vbNewLine, qt)
            .AppendLine("</div>")
            .AppendLine("--->")

            .AppendLine("</center>")
            .AppendLine("</body>")
            .AppendLine("</html>")
        End With

        codeHelp.Text = sb.ToString
        sb.Length = 0
    End Sub

    ''' <summary>
    ''' Возвращает список значений параметра типа ОдинИзВозможных, обрабатыва при необходимости ключевое слово [ByProperty]
    ''' </summary>
    ''' <param name="p">Параметр, список значений которого надо получить</param>
    Private Function GetParamArray(ByRef p As MatewScript.paramsType) As String()
        If p.Type <> MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM Then Return {}
        If IsNothing(p.EnumValues) OrElse p.EnumValues.Count = 0 Then Return {}
        Dim enValue As String = mScript.PrepareStringToPrint(p.EnumValues(0), Nothing, False)
        If enValue.StartsWith("[ByProperty]") Then
            'вместо данного списка надо отобразить список из указанного свойства
            enValue = enValue.Substring("[ByProperty]".Length).Trim  'теперь здесь только имя класса и свойства (напр., O.EquipType)
            Dim arrVal() As String = enValue.Split("."c) '0 - класс, 1 - свойство
            If IsNothing(arrVal) = False AndAlso arrVal.Count = 2 Then
                Dim propTranslateFrom As MatewScript.PropertiesInfoType = Nothing
                Dim cId As Integer = -1, cProp As String = ""
                If mScript.mainClassHash.TryGetValue(arrVal(0), cId) Then
                    mScript.mainClass(cId).Properties.TryGetValue(arrVal(1), propTranslateFrom)
                    If IsNothing(propTranslateFrom.returnArray) = False AndAlso propTranslateFrom.returnArray.Count > 0 Then
                        Return propTranslateFrom.returnArray
                    End If
                End If
            End If
            Return {}
        Else
            Return p.EnumValues
        End If
    End Function

    Private Function GetFunctionParamFormatName(ByRef params() As MatewScript.paramsType, ByVal curParamId As Integer, Optional isNumericEnum As Boolean = False) As String
        Dim p As MatewScript.paramsType = params(curParamId)
        Dim pType As MatewScript.paramsType.paramsTypeEnum = p.Type

        Select Case pType
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_ANY
                Return "any"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_BOOL
                Return "bool"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT
                If IsNothing(p.EnumValues) OrElse p.EnumValues.Count = 0 Then Return "element"
                Return p.EnumValues(0).ToLower
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT2
                For i As Integer = 0 To params.Count - 1
                    If params(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                        If IsNothing(params(i).EnumValues) OrElse params(i).EnumValues.Count = 0 Then Return "sub_element"
                        Dim classId As Integer = 0
                        If mScript.mainClassHash.TryGetValue(params(i).EnumValues(0), classId) = False Then Return "sub_element"
                        Return GetItem3Name(classId).ToLower
                    End If
                Next
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM
                If isNumericEnum Then
                    Return "number (int)"
                Else
                    Return "string"
                End If
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_INTEGER
                Return "number (int)"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_AUDIO
                Return "path (audio)"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_CSS
                Return "path (css)"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_JS
                Return "path (script)"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_PICTURE
                Return "path (image)"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_TEXT
                Return "path (text)"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_SINGLE
                Return "number (float)"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_STRING
                Return "string"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_STRING_OR_NUM
                Return "string / number"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_USER_FUNCTION
                Return "function"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_VARIABLE
                Return "variable"
            Case MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY
                Return "param array"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_PROPERTY
                Return "property name"
            Case MatewScript.paramsType.paramsTypeEnum.PARAM_EVENT
                Return "event name"
        End Select
        Return "any"
    End Function

    Public Sub GenerateHelpForProperty(ByVal propName As String, ByVal classId As Integer, ByRef propValue As MatewScript.PropertiesInfoType)
        Dim qt As Char = Chr(34)
        Dim sb As New System.Text.StringBuilder
        'получаем версию Internet Explorer
        Dim ieVersion As FileVersionInfo = FileVersionInfo.GetVersionInfo("c:\windows\system32\ieframe.dll")
        Dim strIEversion As String = "11"
        If IsNothing(ieVersion) = False Then
            strIEversion = ieVersion.FileMajorPart.ToString
        End If

        Dim retFormat As String = "any"
        Dim retFormatName As String = "Любое значение"
        Select Case propValue.returnType
            Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                retFormat = "bool"
                retFormatName = "True или False (Истина или Ложь)"
            Case MatewScript.ReturnFunctionEnum.RETURN_COLOR
                retFormat = "color"
                retFormatName = "Любой цвет"
            Case MatewScript.ReturnFunctionEnum.RETURN_FUNCTION
                retFormat = "function"
                retFormatName = "Функция Писателя"
            Case MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION
                retFormat = "string/html"
                retFormatName = "Длинный текст с html-разметкой"
            Case MatewScript.ReturnFunctionEnum.RETURN_ENUM
                retFormat = "enum"
                retFormatName = "Один из возможных"
            Case MatewScript.ReturnFunctionEnum.RETURN_EVENT
                retFormat = "script"
                retFormatName = "Скрипт события"
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO
                retFormat = "path"
                retFormatName = "Путь к аудиофайлу относительно папки квеста"
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS
                retFormat = "path"
                retFormatName = "Путь к css-файлу относительно папки квеста"
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_JS
                retFormat = "path"
                retFormatName = "Путь к файлу java-скрипта относительно папки квеста"
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE
                retFormat = "path"
                retFormatName = "Путь к файлу картинки относительно папки квеста"
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_TEXT
                retFormat = "path"
                retFormatName = "Путь к текстовому файлу относительно папки квеста"
            Case MatewScript.ReturnFunctionEnum.RETURN_ELEMENT
                If IsNothing(propValue.returnArray) = False AndAlso propValue.returnArray.Count > 0 Then
                    retFormat = propValue.returnArray(0).ToLower
                    If retFormat = "magic" Then
                        retFormat = "magic book"
                    ElseIf retFormat = "ability" Then
                        retFormat = "ability set"
                    End If
                    retFormatName = "Имя/Id " + retFormat
                End If
        End Select

        With sb
            'Шапка
            .AppendLine("<!DOCTYPE html>")
            .AppendLine("<html>")
            .AppendLine("<head>")
            .AppendFormat("<meta http-equiv={0}Content-Type{0} content={0}text/html; charset=windows-1251{0}/>" & vbNewLine, qt)
            .AppendFormat("<meta http-equiv={0}X-UA-Compatible{0} content={0}IE={1}{0} />" & vbNewLine, qt, strIEversion)
            .AppendFormat("<link rel={0}stylesheet{0} type={0}text/css{0} href={0}../help.css{0}/>" & vbNewLine, qt)
            .AppendLine("</head>")
            .AppendLine()
            'Тело
            .AppendLine("<body>")
            .AppendFormat("<div id={0}HelpConvas{0}>" & vbNewLine, qt)
            'Таблица заголовка
            If String.IsNullOrEmpty(propValue.EditorCaption) OrElse propValue.EditorCaption = propName Then
                .AppendFormat("    <h2 class={0}elementName{0}>{1}</h2>" & vbNewLine, qt, propName)
            Else
                .AppendFormat("    <h2 class={0}elementName{0}>{1} ({2})</h2>" & vbNewLine, qt, propName, propValue.EditorCaption)
            End If
            .AppendFormat("    <table class={0}headerTable{0}>" & vbNewLine, qt)
            .AppendLine("    <tr>")
            .AppendLine("        <td>")
            Dim aStr As String = ""
            If mScript.mainClass(classId).LevelsCount = 1 Then
                aStr = "<span id=" + qt + "propElement1" + qt + ">" + "[Id/Name]" + "</span>"
            ElseIf mScript.mainClass(classId).LevelsCount = 2 Then
                aStr = "<span id=" + qt + "propElement1" + qt + ">" + "[" + GetItem2Name(classId) + " Id/Name" + "<span id=" + qt + "propElement2" + qt + ">, " + GetItem3Name(classId) + " Id/Name</span>]" + "</span>"
            End If
            .AppendFormat("            <code><i>{0}</i> = {1}{2}.<b><font color=red>{3}</font></b></code>" & vbNewLine, retFormat, mScript.mainClass(classId).Names.Last, aStr, propName)
            .AppendLine("        </td>")
            .AppendLine("        <td>")
            .AppendLine("            <table align='right'>")
            .AppendLine("            <tr>")
            .AppendFormat("                <td class={0}typeCaption{0}>Класс</td>" & vbNewLine, qt)

            If mScript.mainClass(classId).Names.Count = 1 Then
                aStr = mScript.mainClass(classId).Names(0)
            Else
                aStr = mScript.mainClass(classId).Names.Last & " ("
                For i As Integer = mScript.mainClass(classId).Names.Count - 2 To 0 Step -1
                    aStr += mScript.mainClass(classId).Names(i)
                    If i > 0 Then aStr += ", "
                Next i
                aStr += ")"
            End If
            .AppendFormat("                <td class={0}typeInfo{0}>{1}</td>" & vbNewLine, qt, aStr)
            .AppendLine("            </tr>")
            .AppendLine("            <tr>")
            .AppendFormat("                <td class={0}typeCaption{0}>Тип</td>" & vbNewLine, qt)
            If propValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                aStr = "Событие"
            ElseIf propValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION Then
                aStr = "Длинный текст"
            Else
                aStr = "Свойство"
            End If
            .AppendFormat("                <td class={0}typeInfo{0}>{1}</td>" & vbNewLine, qt, aStr)
            .AppendLine("            </tr>")
            .AppendLine("            </table>")
            .AppendLine("        </td>")
            .AppendLine("        </tr>")
            .AppendLine("    </table>")
            .AppendLine()
            'Описание
            Dim blnShowOneOnly As Boolean = False
            Dim blnShowLevel1 As Boolean = False
            Dim blnShowLevel2 As Boolean = False
            Dim blnShowLevel3 As Boolean = False
            GetShowDescriptionData(mScript.mainClass(classId).LevelsCount, propValue, blnShowOneOnly, blnShowLevel1, blnShowLevel2, blnShowLevel3)
            If blnShowOneOnly Then
                .AppendFormat("    <p class={0}description{0}>" & vbNewLine, qt)
                .AppendLine("        " & propValue.Description)
                .AppendLine("    </p>")
            Else
                If blnShowLevel1 Then
                    .AppendFormat("    <p class={0}description{0} id={0}Level1{0}>" & vbNewLine, qt)
                    .AppendLine("        " & GetPropertyDescription(propValue.Description, 0))
                    .AppendLine("    </p>")
                End If
                If blnShowLevel2 Then
                    .AppendFormat("    <p class={0}description{0} id={0}Level2{0}>" & vbNewLine, qt)
                    .AppendLine("        " & GetPropertyDescription(propValue.Description, 1))
                    .AppendLine("    </p>")
                End If
                If blnShowLevel3 Then
                    .AppendFormat("    <p class={0}description{0} id={0}Level3{0}>" & vbNewLine, qt)
                    .AppendLine("        " & GetPropertyDescription(propValue.Description, 2))
                    .AppendLine("    </p>")
                End If
            End If

            Dim isNumericEnum As Boolean = False 'если тип Enum, то является ли он численным
            If IsNothing(propValue.Value) = False AndAlso propValue.Value.Length > 0 AndAlso propValue.Value <> "''" Then
                'Значение по умолчанию - не печатаем если значение по умолчанию не установлено
                .AppendLine()
                .AppendLine("    <center>")
                .AppendFormat("        <table class={0}tableParams{0} id={0}tableDefaults{0}>" & vbNewLine, qt)
                .AppendLine("            <tr>")
                .AppendLine("                <td width=50%>Значение по умолчанию</td>")
                'If IsNothing(propValue.Value) OrElse propValue.Value.Length = 0 Then
                '    .AppendFormat("                <td id={0}defaultValue{0}>отсутствует</td>" & vbNewLine, qt)
                If propValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                    .AppendFormat("                <td id={0}defaultValue{0}><code>[скрипт]</code></td>" & vbNewLine, qt)
                ElseIf propValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION Then
                    .AppendFormat("                <td id={0}defaultValue{0}><code>[длинный текст]</code></td>" & vbNewLine, qt)
                ElseIf propValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_ENUM AndAlso IsNothing(propValue.returnArray) = False AndAlso propValue.returnArray.Count > 0 Then
                    Dim retStr As String = mScript.PrepareStringToPrint(propValue.returnArray(0), Nothing, False)
                    Dim isNum As Boolean = False
                    If retStr.Length > 0 Then
                        If Integer.TryParse(retStr.Chars(0), Nothing) Then
                            isNum = True
                        ElseIf retStr.Length > 1 AndAlso retStr.Chars(0) = "-"c AndAlso Integer.TryParse(retStr.Chars(1), Nothing) Then
                            isNum = True
                        End If
                    End If
                    If isNum Then
                        'enum численный
                        retStr = mScript.PrepareStringToPrint(propValue.Value, Nothing, False)
                        Dim retVal As Double = Val(retStr)
                        .AppendFormat("                <td id={0}defaultValue{0}  class={0}selectable{0} rVal={0}{1}{0}>{2}</td>" & vbNewLine, qt, Convert.ToString(retVal, provider_points), retStr)
                        isNumericEnum = True
                    Else
                        'enum строчный
                        retStr = mScript.PrepareStringToPrint(propValue.Value, Nothing, False)
                        .AppendFormat("                <td id={0}defaultValue{0}  class={0}selectable{0} rVal={0}{1}{0}>{1}</td>" & vbNewLine, qt, retStr)
                        isNumericEnum = False
                    End If
                Else
                    Dim retStr As String = mScript.PrepareStringToPrint(propValue.Value, Nothing, False)
                    .AppendFormat("                <td id={0}defaultValue{0}  class={0}selectable{0} rVal={0}{1}{0}>{1}</td>" & vbNewLine, qt, retStr)
                End If
                .AppendLine("            </tr>")
                .AppendLine("        </table>")
            Else
                .AppendLine("    <center>")
            End If

            .AppendLine()
            'Принимаемые значения
            .AppendLine("        <h3>Принимаемые значения</h3>")
            .AppendFormat("        <table class={0}tableParams{0}>" & vbNewLine, qt)

            Select Case propValue.returnType
                Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                    .AppendLine("            <tr>")
                    .AppendLine("                <td><i><code>Истина, верно</code></i></td>")
                    .AppendFormat("                <td class = {0}selectable{0} rVal={0}True{0}><i><code>True</code></i></td>" & vbNewLine, qt)
                    .AppendLine("            </tr>")

                    .AppendLine("            <tr>")
                    .AppendLine("                <td><i><code>Ложь, неверно</code></i></td>")
                    .AppendFormat("                <td class = {0}selectable{0} rVal={0}False{0}><i><code>False</code></i></td>" & vbNewLine, qt)
                    .AppendLine("            </tr>")
                Case MatewScript.ReturnFunctionEnum.RETURN_ENUM
                    If isNumericEnum Then
                        For i As Integer = 0 To propValue.returnArray.Count - 1
                            Dim curText As String = mScript.PrepareStringToPrint(propValue.returnArray(i), Nothing, False)
                            Dim curVal As String = Convert.ToString(Val(curText), provider_points)
                            curText = curText.Substring(curVal.Length).Trim
                            If curText.StartsWith("-") Then curText = curText.Substring(1).Trim

                            .AppendLine("            <tr>")
                            .AppendFormat("                <td><i><code>{0}</code></i></td>" & vbNewLine, curVal)
                            .AppendFormat("                <td rVal={0}{1}{0} class={0}selectable{0}>{2}</td>" & vbNewLine, qt, curVal, curText)
                            .AppendLine("            </tr>")
                        Next i
                    ElseIf IsNothing(propValue.returnArray) = False AndAlso propValue.returnArray.Count > 0 Then
                        For i As Integer = 0 To propValue.returnArray.Count - 1
                            Dim curText As String = mScript.PrepareStringToPrint(propValue.returnArray(i), Nothing, False)
                            .AppendLine("            <tr>")
                            .AppendFormat("                <td><i><code>{0}</code></i></td>" & vbNewLine, (i + 1).ToString)
                            .AppendFormat("                <td rVal={0}{1}{0} class={0}selectable{0}>{1}</td>" & vbNewLine, qt, curText)
                            .AppendLine("            </tr>")
                        Next i
                    Else
                        .AppendLine("            <tr>")
                        .AppendFormat("                <td><i><code>{0}</code></i></td>" & vbNewLine, retFormat)
                        .AppendFormat("                <td>{0}</td>" & vbNewLine, retFormatName)
                        .AppendLine("            </tr>")
                    End If
                Case Else
                    .AppendLine("            <tr>")
                    .AppendFormat("                <td><i><code>{0}</code></i></td>" & vbNewLine, retFormat)
                    .AppendFormat("                <td>{0}</td>" & vbNewLine, retFormatName)
                    .AppendLine("            </tr>")
            End Select
            .AppendLine("        </table>")
            .AppendLine()

            If propValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                'Принимаемые параметры
                .AppendLine("        <h3>Принимаемые параметры</h3>")
                GetShowReceivedParamData(mScript.mainClass(classId).LevelsCount, propValue, blnShowOneOnly, blnShowLevel1, blnShowLevel2, blnShowLevel3)
                If blnShowOneOnly Then
                    'Параметры для всех уровней класса одинаковы
                    .AppendFormat("        <table class={0}tableParams{0}>" & vbNewLine, qt)
                    If (IsNothing(propValue.params) OrElse propValue.params.Count = 0) OrElse (blnShowLevel1 = False AndAlso blnShowLevel2 = False AndAlso blnShowLevel3 = False) Then
                        'параметры не указаны
                        .AppendLine("            <tr>")
                        .AppendLine("                <td><code>нет</code></td>")
                        .AppendLine("            </tr>")
                    Else
                        'параметры указаны
                        For i As Integer = 0 To propValue.params.Count - 1
                            Dim param As MatewScript.paramsType = propValue.params(i)
                            Dim scores As Integer = param.Type
                            If scores >= 16 Then scores -= 16 'вычленили ParamArray
                            If scores >= 8 Then Exit For 'начались параметры Return
                            scores = param.Type
                            .AppendLine("            <tr>")
                            If scores >= 8 Then
                                .AppendFormat("                <td><code>Param[{0}, {1}, ... n]</code></td>" & vbNewLine, i.ToString, (i + 1).ToString)
                            Else
                                .AppendFormat("                <td><code>Param[{0}]</code></td>" & vbNewLine, i.ToString)
                            End If
                            .AppendLine("                <td>" & param.Description & "</td>")
                            .AppendLine("            </tr>")
                        Next
                    End If
                    .AppendLine("        </table>")
                Else
                    'Параметры для разных уровней класса разные
                    .AppendFormat("        <table class={0}tableParams{0} id={0}Level1_1{0}>" & vbNewLine, qt)
                    If blnShowLevel1 Then
                        Dim pos As Integer = 0
                        For i As Integer = 0 To propValue.params.Count - 1
                            Dim param As MatewScript.paramsType = propValue.params(i)
                            Dim scores As Integer = param.Type
                            If scores >= 16 Then scores -= 16 'вычленили ParamArray
                            If scores >= 8 Then Exit For 'начались параметры Return
                            If scores >= 4 Then scores -= 4 '- level3
                            If scores >= 2 Then scores -= 2 ' - level2
                            If scores <> 1 Then Continue For 'параметр не для 1 уровня
                            scores = param.Type
                            .AppendLine("            <tr>")
                            If scores >= 16 Then
                                .AppendFormat("                <td><code>Param[{0}, {1}, ... n]</code></td>" & vbNewLine, pos.ToString, (pos + 1).ToString)
                            Else
                                .AppendFormat("                <td><code>Param[{0}]</code></td>" & vbNewLine, pos.ToString)
                            End If
                            .AppendLine("                <td>" & param.Description & "</td>")
                            .AppendLine("            </tr>")
                            pos += 1
                        Next
                    Else
                        .AppendLine("            <tr>")
                        .AppendLine("                <td><code>нет</code></td>")
                        .AppendLine("            </tr>")
                    End If
                    .AppendLine("        </table>")

                    If mScript.mainClass(classId).LevelsCount >= 1 Then
                        .AppendFormat("        <table class={0}tableParams{0} id={0}Level2_1{0}>" & vbNewLine, qt)
                        If blnShowLevel2 Then
                            Dim pos As Integer = 0
                            For i As Integer = 0 To propValue.params.Count - 1
                                Dim param As MatewScript.paramsType = propValue.params(i)
                                Dim scores As Integer = param.Type
                                If scores >= 16 Then scores -= 16 'вычленили ParamArray
                                If scores >= 8 Then Exit For 'начались параметры Return
                                If scores >= 4 Then scores -= 4 '- level3
                                If scores < 2 Then Continue For 'параметр не для 2 уровня
                                scores = param.Type
                                .AppendLine("            <tr>")
                                If scores >= 16 Then
                                    .AppendFormat("                <td><code>Param[{0}, {1}, ... n]</code></td>" & vbNewLine, pos.ToString, (pos + 1).ToString)
                                Else
                                    .AppendFormat("                <td><code>Param[{0}]</code></td>" & vbNewLine, pos.ToString)
                                End If
                                .AppendLine("                <td>" & param.Description & "</td>")
                                .AppendLine("            </tr>")
                                pos += 1
                            Next
                        Else
                            .AppendLine("            <tr>")
                            .AppendLine("                <td><code>нет</code></td>")
                            .AppendLine("            </tr>")
                        End If
                        .AppendLine("        </table>")
                    End If

                    If mScript.mainClass(classId).LevelsCount >= 2 Then
                        .AppendFormat("        <table class={0}tableParams{0} id={0}Level3_1{0}>" & vbNewLine, qt)
                        If blnShowLevel3 Then
                            Dim pos As Integer = 0
                            For i As Integer = 0 To propValue.params.Count - 1
                                Dim param As MatewScript.paramsType = propValue.params(i)
                                Dim scores As Integer = param.Type
                                If scores >= 16 Then scores -= 16 'вычленили ParamArray
                                If scores >= 8 Then Exit For 'начались параметры Return
                                If scores < 4 Then Continue For 'параметр не для 3 уровня
                                scores = param.Type
                                .AppendLine("            <tr>")
                                If scores >= 16 Then
                                    .AppendFormat("                <td><code>Param[{0}, {1}, ... n]</code></td>" & vbNewLine, pos.ToString, (pos + 1).ToString)
                                Else
                                    .AppendFormat("                <td><code>Param[{0}]</code></td>" & vbNewLine, pos.ToString)
                                End If
                                .AppendLine("                <td>" & param.Description & "</td>")
                                .AppendLine("            </tr>")
                                pos += 1
                            Next
                        Else
                            .AppendLine("            <tr>")
                            .AppendLine("                <td><code>нет</code></td>")
                            .AppendLine("            </tr>")
                        End If
                        .AppendLine("        </table>")
                    End If
                End If 'blnShowOneOnly

                .AppendLine()
                'Возвращаемые параметры
                Dim strLevel(2) As String
                If blnShowOneOnly Then
                    strLevel(0) = "Level1_1"
                    strLevel(1) = "Level2_1"
                    strLevel(2) = "Level3_1"
                Else
                    strLevel(0) = "Level1_2"
                    strLevel(1) = "Level2_2"
                    strLevel(2) = "Level3_2"
                End If
                .AppendLine("        <h3>Возвращаемые параметры</h3>")
                GetShowReturnedParamData(mScript.mainClass(classId).LevelsCount, propValue, blnShowOneOnly, blnShowLevel1, blnShowLevel2, blnShowLevel3)
                If blnShowOneOnly Then
                    'Возвращаемые параметры для всех уровней класса одинаковы
                    .AppendFormat("        <table class={0}tableParams{0}>" & vbNewLine, qt)
                    If (IsNothing(propValue.params) OrElse propValue.params.Count = 0) OrElse (blnShowLevel1 = False AndAlso blnShowLevel2 = False AndAlso blnShowLevel3 = False) Then
                        'возвращаемые параметры не указаны
                        .AppendLine("            <tr>")
                        .AppendLine("                <td><code>нет</code></td>")
                        .AppendLine("            </tr>")
                    Else
                        'возвращаемые параметры указаны
                        For i As Integer = 0 To propValue.params.Count - 1
                            Dim param As MatewScript.paramsType = propValue.params(i)
                            Dim scores As Integer = param.Type
                            If scores >= 16 Then Continue For 'ParamArray - обычный параметр (параметры Return еще не начались)
                            If scores < 8 Then Continue For 'параметры Return еще не начались
                            .AppendLine("            <tr>")
                            Dim rName As String = param.Name
                            If rName.Length > 20 Then rName = "<font size=1>" + rName + "</font>"
                            If String.IsNullOrEmpty(rName) Then rName = "[ничего]"
                            .AppendFormat("                <td><code>Return {0}</code></td>" & vbNewLine, rName)
                            .AppendLine("                <td>" & param.Description & "</td>")
                            .AppendLine("            </tr>")
                        Next
                    End If
                    .AppendLine("        </table>")
                Else
                    'Возвращаемые параметры для разных уровней класса разные
                    .AppendFormat("        <table class={0}tableParams{0} id={0}{1}{0}>" & vbNewLine, qt, strLevel(0))
                    If blnShowLevel1 Then
                        For i As Integer = 0 To propValue.params.Count - 1
                            Dim param As MatewScript.paramsType = propValue.params(i)
                            Dim scores As Integer = param.Type
                            If scores >= 16 Then Continue For 'ParamArray - обычный параметр (параметры Return еще не начались)
                            If scores < 8 Then Continue For 'параметры Return еще не начались
                            scores -= 8
                            If scores >= 4 Then scores -= 4 '- level3
                            If scores >= 2 Then scores -= 2 ' - level2
                            If scores <> 1 Then Continue For 'параметр не для 1 уровня
                            .AppendLine("            <tr>")
                            Dim rName As String = param.Name
                            If String.IsNullOrEmpty(rName) Then rName = "[ничего]"
                            If rName.Length > 20 Then rName = "<font size=1>" + rName + "</font>"
                            .AppendFormat("                <td><code>Return {0}</code></td>" & vbNewLine, rName)
                            .AppendLine("                <td>" & param.Description & "</td>")
                            .AppendLine("            </tr>")
                        Next
                    Else
                        .AppendLine("            <tr>")
                        .AppendLine("                <td><code>нет</code></td>")
                        .AppendLine("            </tr>")
                    End If
                    .AppendLine("        </table>")

                    If mScript.mainClass(classId).LevelsCount >= 1 Then
                        .AppendFormat("        <table class={0}tableParams{0} id={0}{1}{0}>" & vbNewLine, qt, strLevel(1))
                        If blnShowLevel2 Then
                            For i As Integer = 0 To propValue.params.Count - 1
                                Dim param As MatewScript.paramsType = propValue.params(i)
                                Dim scores As Integer = param.Type
                                If scores >= 16 Then Continue For 'ParamArray - обычный параметр (параметры Return еще не начались)
                                If scores < 8 Then Continue For 'параметры Return еще не начались
                                scores -= 8
                                If scores >= 4 Then scores -= 4 '- level3
                                If scores < 2 Then Continue For 'параметр не для 2 уровня
                                .AppendLine("            <tr>")
                                Dim rName As String = param.Name
                                If String.IsNullOrEmpty(rName) Then rName = "[ничего]"
                                If rName.Length > 20 Then rName = "<font size=1>" + rName + "</font>"
                                .AppendFormat("                <td><code>Return {0}</code></td>" & vbNewLine, rName)
                                .AppendLine("                <td>" & param.Description & "</td>")
                                .AppendLine("            </tr>")
                            Next
                        Else
                            .AppendLine("            <tr>")
                            .AppendLine("                <td><code>нет</code></td>")
                            .AppendLine("            </tr>")
                        End If
                        .AppendLine("        </table>")
                    End If

                    If mScript.mainClass(classId).LevelsCount >= 2 Then
                        .AppendFormat("        <table class={0}tableParams{0} id={0}{1}{0}>" & vbNewLine, qt, strLevel(2))
                        If blnShowLevel3 Then
                            For i As Integer = 0 To propValue.params.Count - 1
                                Dim param As MatewScript.paramsType = propValue.params(i)
                                Dim scores As Integer = param.Type
                                If scores >= 16 Then Continue For 'ParamArray - обычный параметр (параметры Return еще не начались)
                                If scores < 8 Then Continue For 'параметры Return еще не начались
                                scores -= 8
                                If scores < 4 Then Continue For 'параметр не для 3 уровня
                                .AppendLine("            <tr>")
                                Dim rName As String = param.Name
                                If rName.Length > 20 Then rName = "<font size=1>" + rName + "</font>"
                                If String.IsNullOrEmpty(rName) Then rName = "[ничего]"
                                .AppendFormat("                <td><code>Return {0}</code></td>" & vbNewLine, rName)
                                .AppendLine("                <td>" & param.Description & "</td>")
                                .AppendLine("            </tr>")
                            Next
                        Else
                            .AppendLine("            <tr>")
                            .AppendLine("                <td><code>нет</code></td>")
                            .AppendLine("            </tr>")
                        End If
                        .AppendLine("        </table>")
                    End If
                End If 'blnShowOneOnly
            End If

            'Завершение
            Select Case propValue.returnType
                Case MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION
                    .AppendFormat("        <div id={0}selectorsContainer{0}></div>" & vbNewLine, qt)
                    .AppendFormat("        <iframe id={0}dataContainer{0} height={0}100%{0}></iframe>" & vbNewLine, qt)
                Case MatewScript.ReturnFunctionEnum.RETURN_ELEMENT, MatewScript.ReturnFunctionEnum.RETURN_FUNCTION
                    .AppendFormat("        <div id={0}placeMarker{0}></div>" & vbNewLine, qt)
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO, MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS, MatewScript.ReturnFunctionEnum.RETURN_PATH_JS, _
                    MatewScript.ReturnFunctionEnum.RETURN_PATH_TEXT
                    .AppendFormat("        <div id={0}placeMarker{0} fitToWidth={0}True{0} style={0}width:100%{0}></div>" & vbNewLine, qt)
                Case MatewScript.ReturnFunctionEnum.RETURN_COLOR
                    .AppendFormat("        <div id={0}dataContainer{0} fitToWidth={0}True{0}>" & vbNewLine, qt)
                    .AppendLine("            <table>")
                    .AppendLine("            <tr>")
                    .AppendLine("                <td>")
                    .AppendFormat("                    <img id={0}colorMap{0}/>" & vbNewLine, qt)
                    .AppendLine("                </td>")
                    .AppendLine("                <td>")
                    .AppendLine("                    <table>")
                    .AppendLine("                    <tr align='center'>")
                    .AppendLine("                        <td>")
                    .AppendFormat("                            прозрачность (от 0 до 1):<br /> <input type={0}text{0} value={0}1{0} id={0}colorTransparency{0}/>" & vbNewLine, qt)
                    .AppendLine("                        </td>")
                    .AppendLine("                    </tr>")
                    .AppendLine("                    <tr align='center'>")
                    .AppendLine("                        <td>")
                    .AppendFormat("                            <span id={0}colorSample{0}></span>" & vbNewLine, qt)
                    .AppendLine("                        </td>")
                    .AppendLine("                    </tr>")
                    .AppendLine("                    </table>")
                    .AppendLine("                </td>")
                    .AppendLine("            </tr>")
                    .AppendLine("            </table>")
                    .AppendLine()
                    .AppendFormat("            <div id={0}selectedColorsContainer{0}></div>" & vbNewLine, qt)
                    .AppendLine("        </div>")
            End Select

            .AppendLine("    </center>")

            If propValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                .AppendFormat("    <div id={0}dataContainer{0} style={0}display: none{0}></div>" & vbNewLine, qt)
            ElseIf propValue.returnType = MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE Then
                .AppendFormat("    <div id={0}sampleContainer{0}></div>" & vbNewLine, qt)
                .AppendLine("    <table>")
                .AppendLine("    <tr>")
                .AppendFormat("        <td id={0}placeMarker{0} fitToWidth={0}True{0} style={0}width:350px;{0}></td>" & vbNewLine, qt)
                .AppendFormat("        <td id={0}dataContainer{0} style={0}display: none;{0}></td>" & vbNewLine, qt)
                .AppendLine("    </tr>")
                .AppendLine("    </table>")
            End If
            .AppendLine("</div>")
            'заготовка для вставки примера скрипта
            .AppendLine()
            .AppendLine("<!--")
            .AppendLine("<h2>Заготовка для вставки примера скрипта</h2>")
            .AppendFormat("<div class={0}codeSampleTitle{0}>Название примера</div>" & vbNewLine, qt)
            .AppendFormat("<div class={0}codeSampleDescription{0}>Описание примера скрипта, вводная к игровой ситуации где это может пригодиться.</div>" & vbNewLine, qt)
            .AppendFormat("<div class={0}codeSample{0}>" & vbNewLine, qt)
            .AppendLine("    Непосредственно скрипт")
            .AppendFormat("    <div class={0}codeSampleNote{0}>... прерываемый различными замечаниями. Так можно разделить, например, скрипты из разных функций.</div>" & vbNewLine, qt)
            .AppendLine("</div>")
            .AppendLine("--->")


            .AppendLine("</body>")
            .AppendLine("</html>")
        End With

        codeHelp.Text = sb.ToString
        sb.Length = 0
    End Sub

    ''' <summary>
    ''' Выясняет для параметров свойства Event следующее: надо ли создавать блок Принимаемых параметров один или несколько; надо ли создавать блок для каждого из уровней.
    ''' </summary>
    ''' <param name="levels">Количество уровней класса</param>
    ''' <param name="prop">Свойство для которого определяем показатели</param>
    ''' <param name="showOneOnly">Ссылка для получения того надо ли создавать один блок или несколько</param>
    ''' <param name="showLevel1">надо ли создавать блок для 1 уровня</param>
    ''' <param name="showLevel2">надо ли создавать блок для 2 уровня</param>
    ''' <param name="showLevel3">надо ли создавать блок для 3 уровня</param>
    Private Sub GetShowReceivedParamData(ByVal levels As Integer, ByRef prop As MatewScript.PropertiesInfoType, ByRef showOneOnly As Boolean, ByRef showLevel1 As Boolean, _
                                       ByRef showLevel2 As Boolean, ByRef showLevel3 As Boolean)
        Dim params() As MatewScript.paramsType = prop.params
        If IsNothing(params) OrElse params.Count = 0 OrElse levels = 0 Then
            showOneOnly = True
            showLevel1 = True
            showLevel2 = False
            showLevel3 = False
            Return
        End If

        showOneOnly = True
        showLevel1 = False
        showLevel2 = False
        showLevel3 = False
        Dim prevVal As Byte = params(0).Type

        For i As Integer = 0 To params.Count - 1
            Dim scores As Integer = params(i).Type
            If scores >= 16 Then scores -= 16 'вычленяем ParamArray
            If scores >= 8 Then Exit For 'начались параметры типа Return - заканчиваем
            If scores <> prevVal Then showOneOnly = False
            If scores >= 4 Then
                scores -= 4
                showLevel3 = True
            End If
            If scores >= 2 Then
                scores -= 2
                showLevel2 = True
            End If
            If scores = 1 Then showLevel1 = True
        Next i

        If Not showOneOnly Then Return 'параметры отличаются - значит надо показывать разные окна принимаемых параметров - выход
        'может возникнуть ситуация, что параметры одинаковы, но показывать одно окно для всех нельзя. Например, для уровня1 - ни одного параметра, а для уровня2 - все.
        'при этом параметры будут иметь равное кол-во scores, но окна должны быть разными
        If levels = 0 Then Return
        If levels = 1 Then
            showOneOnly = (prevVal = 3) 'все параметры - для 1 и 2 уровня
        End If
        'остался третий уровень
        showOneOnly = (prevVal = 7) 'все параметры - для 1, 2 и 3 уровня
    End Sub

    ''' <summary>
    ''' Выясняет для параметров типа RETURN свойства Event следующее: надо ли создавать блок Возвращаемых параметров один или несколько; надо ли создавать блок для каждого из уровней.
    ''' </summary>
    ''' <param name="levels">Количество уровней класса</param>
    ''' <param name="prop">Свойство для которого определяем показатели</param>
    ''' <param name="showOneOnly">Ссылка для получения того надо ли создавать один блок или несколько</param>
    ''' <param name="showLevel1">надо ли создавать блок для 1 уровня</param>
    ''' <param name="showLevel2">надо ли создавать блок для 2 уровня</param>
    ''' <param name="showLevel3">надо ли создавать блок для 3 уровня</param>
    Private Sub GetShowReturnedParamData(ByVal levels As Integer, ByRef prop As MatewScript.PropertiesInfoType, ByRef showOneOnly As Boolean, ByRef showLevel1 As Boolean, _
                                       ByRef showLevel2 As Boolean, ByRef showLevel3 As Boolean)
        Dim params() As MatewScript.paramsType = prop.params
        If IsNothing(params) OrElse params.Count = 0 OrElse levels = 0 Then
            showOneOnly = True
            showLevel1 = True
            showLevel2 = False
            showLevel3 = False
            Return
        End If

        showOneOnly = True
        showLevel1 = False
        showLevel2 = False
        showLevel3 = False
        Dim prevVal As Byte = params.Last.Type
        For i As Integer = params.Count - 1 To 0 Step -1
            Dim scores As Integer = params(i).Type
            If scores >= 16 Then Exit For 'ParamArray - это не параметр типа Return - выход
            If scores < 8 Then Exit For 'начались обычные параметры - заканчиваем
            If scores <> prevVal Then showOneOnly = False
            scores -= 8
            If scores >= 4 Then
                scores -= 4
                showLevel3 = True
            End If
            If scores >= 2 Then
                scores -= 2
                showLevel2 = True
            End If
            If scores = 1 Then showLevel1 = True
        Next i

        If Not showOneOnly Then Return 'параметры отличаются - значит надо показывать разные окна возвращаемых параметров - выход
        'может возникнуть ситуация, что параметры одинаковы, но показывать одно окно для всех нельзя. Например, для уровня1 - ни одного параметра, а для уровня2 - все.
        'при этом параметры будут иметь равное кол-во scores, но окна должны быть разными
        If levels = 0 Then Return
        If prevVal >= 8 Then prevVal -= 8 'вычленяем Return
        If levels = 1 Then
            showOneOnly = (prevVal = 3) 'все параметры - для 1 и 2 уровня
        End If
        'остался третий уровень
        showOneOnly = (prevVal = 7) 'все параметры - для 1, 2 и 3 уровня
    End Sub

    ''' <summary>
    ''' Выясняет для данного свойства следующее: надо ли создавать текст описания один или несколько; надо ли создавать описание для каждого из уровней.
    ''' </summary>
    ''' <param name="levels">Количество уровней класса</param>
    ''' <param name="prop">Свойство для которого определяем показатели</param>
    ''' <param name="showOneOnly">Ссылка для получения того надо ли создавать один блок описания или несколько</param>
    ''' <param name="showLevel1">надо ли создавать блок описания для 1 уровня</param>
    ''' <param name="showLevel2">надо ли создавать блок описания для 2 уровня</param>
    ''' <param name="showLevel3">надо ли создавать блок описания для 3 уровня</param>
    Private Sub GetShowDescriptionData(ByVal levels As Integer, ByRef prop As MatewScript.PropertiesInfoType, ByRef showOneOnly As Boolean, ByRef showLevel1 As Boolean, _
                                       ByRef showLevel2 As Boolean, ByRef showLevel3 As Boolean)
        Select Case levels
            Case 0
                showOneOnly = True
                showLevel1 = True
                showLevel2 = False
                showLevel3 = False
            Case 1
                showLevel3 = False
                Select Case prop.Hidden
                    Case MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL, MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR
                        showOneOnly = True
                        showLevel2 = True
                        showLevel1 = False
                    Case MatewScript.PropertyHiddenEnum.LEVEL1_ONLY
                        showOneOnly = True
                        showLevel1 = True
                        showLevel2 = False
                    Case MatewScript.PropertyHiddenEnum.LEVEL2_ONLY
                        showOneOnly = True
                        showLevel1 = False
                        showLevel2 = True
                    Case Else
                        showOneOnly = False
                        showLevel1 = True
                        showLevel2 = True
                End Select
            Case Else
                Select Case prop.Hidden
                    Case MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL, MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR, MatewScript.PropertyHiddenEnum.LEVEL2_ONLY
                        showOneOnly = True
                        showLevel2 = True
                        showLevel1 = False
                        showLevel3 = False
                    Case MatewScript.PropertyHiddenEnum.LEVEL1_ONLY
                        showOneOnly = True
                        showLevel1 = True
                        showLevel2 = False
                        showLevel3 = False
                    Case MatewScript.PropertyHiddenEnum.LEVEL12_ONLY
                        showOneOnly = False
                        showLevel1 = True
                        showLevel2 = True
                        showLevel3 = False
                    Case MatewScript.PropertyHiddenEnum.LEVEL13_ONLY
                        showOneOnly = False
                        showLevel1 = True
                        showLevel2 = False
                        showLevel3 = True
                    Case MatewScript.PropertyHiddenEnum.LEVEL23_ONLY
                        showOneOnly = False
                        showLevel1 = False
                        showLevel2 = True
                        showLevel3 = True
                    Case MatewScript.PropertyHiddenEnum.LEVEL3_ONLY
                        showOneOnly = True
                        showLevel1 = False
                        showLevel2 = False
                        showLevel3 = True
                    Case Else
                        showOneOnly = False
                        showLevel1 = True
                        showLevel2 = True
                        showLevel3 = True
                End Select
        End Select
    End Sub

    ''' <summary>
    ''' Строка для формулы заголовка
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    Private Function GetItem2Name(ByVal classId As Integer) As String
        Select Case mScript.mainClass(classId).Names(0)
            Case "M"
                Return "menu"
            Case "Map"
                Return "map"
            Case "Med"
                Return "media_list"
            Case "Mg"
                Return "magic_book"
            Case "Ab"
                Return "ablities_set"
            Case "Army"
                Return "army"
            Case Else
                Return "item"
        End Select
    End Function

    ''' <summary>
    ''' Строка для формулы заголовка
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    Private Function GetItem3Name(ByVal classId As Integer) As String
        Select Case mScript.mainClass(classId).Names(0)
            Case "M"
                Return "menu_item"
            Case "Map"
                Return "cell"
            Case "Med"
                Return "audio_file"
            Case "Mg"
                Return "magic"
            Case "Ab"
                Return "ablity"
            Case "Army"
                Return "unit"
            Case Else
                Return "subitem"
        End Select
    End Function

    Private Sub btnShowScript_Click(sender As Object, e As EventArgs) Handles btnShowScript.Click
        dlgExecute.Show(Me)
    End Sub

    Private Sub dlgGenerateHelp_Load(sender As Object, e As EventArgs) Handles Me.Load
        codeHelp.SplitContainerHorizontal.Panel2Collapsed = True
    End Sub
End Class
