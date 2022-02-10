Imports System.Collections.ObjectModel

Public Class cListManagerClass

    Public Enum fileTypesEnum As Byte
        PICTURES = 0
        AUDIO = 1
        TEXT = 2
        CSS = 3
        JS = 4
    End Enum

    ''' <summary>Хранит списки имен элементов всех классов 2 уровня (ключ - Id класса, значение - список имен)</summary>
    Private lstMain As New SortedList(Of Integer, List(Of String))
    Private arrPictures() As String
    Private arrAudio() As String
    Private arrText() As String
    Private arrCSS() As String
    Private arrJS() As String

    ''' <summary>Хранит списки классов и css-файлов (ключ - путь к css относительно папки квеста, значение - имя класса (без кавычек))</summary>
    Private lstCssClasses As New SortedList(Of String, List(Of String))
    ''' <summary>Хранит списки анимаций и css-файлов (ключ - путь к css относительно папки квеста, значение - имя анимации (без кавычек))</summary>
    Private lstCssAnimations As New SortedList(Of String, List(Of String))

    Sub New()
        'UpdateFileLists()
    End Sub

    ''' <summary>
    ''' Возвращает список классов из ресурса, определяемого в файле помощи
    ''' </summary>
    ''' <param name="strSource">Исполняемая строка, возвращающая путь к css-файлу</param>
    ''' <param name="strSourceParams">Параметры, передаваемые исполняемой строке. Ключевые слова this и parent заменяются на имена текущего элемента и его родителя</param>
    ''' <param name="strPattern">Паттерн для фильтрации классов. Будут выбраны только те, которые начинаются на указанную строку</param>
    ''' <returns>Список классов или пустой список</returns>
    Public Function GetCssClassesByPattern(ByVal strSource As String, ByVal strSourceParams As String, ByVal strPattern As String, Optional ByRef CSSFile As String = "") As List(Of String)
        CSSFile = ""
        Dim lstClasses As New List(Of String)
        If String.IsNullOrEmpty(strSource) Then Return lstClasses
        If IsNothing(cPanelManager.ActivePanel) Then Return lstClasses

        Dim arrParams() As String = Nothing
        If String.IsNullOrEmpty(strSourceParams) = False Then
            'параметры к исполняемой строке strSource переданы. Получаем их и заменяем ключевые слова на реальные значения
            arrParams = strSourceParams.Split(","c)
            For i As Integer = 0 To arrParams.Count - 1
                Dim strParam As String = arrParams(i).Trim
                If String.Compare(strParam, "parent", True) = 0 Then
                    'получаем имя родителя текущего элемента
                    If cPanelManager.ActivePanel.child3Name.Length > 0 Then
                        strParam = cPanelManager.ActivePanel.child2Name
                    ElseIf cPanelManager.ActivePanel.classId = mScript.mainClassHash("A") AndAlso cPanelManager.ActivePanel.supraElementName.Length > 0 Then
                        strParam = cPanelManager.ActivePanel.supraElementName
                    Else
                        Return lstClasses
                    End If
                ElseIf String.Compare(strParam, "this", True) = 0 Then
                    'получаем имя текущего элемента
                    If cPanelManager.ActivePanel.child3Name.Length > 0 Then
                        strParam = cPanelManager.ActivePanel.child3Name
                    ElseIf cPanelManager.ActivePanel.child2Name.Length > 0 Then
                        strParam = cPanelManager.ActivePanel.child2Name
                    Else
                        Return lstClasses
                    End If
                End If
                arrParams(i) = strParam
            Next
        End If

        'Выполняем скрипт в исполняемой строке, и получаем путь к css-файлу относительно папки квеста
        strSource = WrapString(strSource)
        strSource = mScript.PrepareStringToPrint(strSource, arrParams, True)
        If strSource = "#Error" Then Return lstClasses
        'Создаем/получаем полные список классов
        CSSFile = strSource
        Dim lstFull As List(Of String) = MakeCssClassesList(strSource)
        If lstFull.Count = 0 Then Return lstClasses
        If String.IsNullOrEmpty(strPattern) Then Return lstFull 'фильтр не указан - возвращаем весь список
        'возвращаем отфильтрованный список
        For i As Integer = 0 To lstFull.Count - 1
            If lstFull(i).StartsWith(strPattern, StringComparison.CurrentCultureIgnoreCase) Then
                lstClasses.Add(lstFull(i))
            End If
        Next
        Return lstClasses
    End Function

    ''' <summary>
    ''' Создает список классов, найденных в данном css-файле (а также в дочерних, переданных с помощью оператора @import)
    ''' </summary>
    ''' <param name="pathToCss">Путь к css-файлу</param>
    ''' <param name="forceUpdate">Принудительно обновить. Если True - то список будет перестроен, если False - то будет возвращен подготовленный заранее список (если он был, иначе создается новый)</param>
    Public Function MakeCssClassesList(ByVal pathToCss As String, Optional forceUpdate As Boolean = False) As List(Of String)
        If lstCssClasses.ContainsKey(pathToCss) Then
            If forceUpdate Then
                lstCssClasses.Remove(pathToCss)
            Else
                Return lstCssClasses(pathToCss)
            End If
        End If

        Dim fullPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, pathToCss)
        'получаем список классв
        Dim lstClasses As New List(Of String)
        ReadCssClasses(fullPath, lstClasses)
        'сохраняем в структуре
        lstCssClasses.Add(pathToCss, lstClasses)
        Return lstClasses
    End Function

    ''' <summary>
    ''' Рекурсивная функция, составляющая список классов из css-файла, включая другие css-файлы, включенные посредством  @import
    ''' </summary>
    ''' <param name="pathToCss">Полный (и заранее проверенный) путь к css-файлу</param>
    ''' <param name="lstClasses">Список, в котором будут собираться имена классов</param>
    Private Sub ReadCssClasses(ByVal pathToCss As String, ByRef lstClasses As List(Of String))
        If FileIO.FileSystem.FileExists(pathToCss) = False Then
            MessageBox.Show("Файл " + pathToCss + " не найден.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        'Считываем файл
        Dim fileContents As String
        Try
            fileContents = My.Computer.FileSystem.ReadAllText(pathToCss, System.Text.Encoding.Default)
        Catch ex As Exception
            MessageBox.Show("Не удалось прочитать файл " + pathToCss + ".", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End Try

        'ищем @import 
        Dim pos As Integer = 0
        Do
            pos = fileContents.IndexOf("@import ", pos)
            If pos = -1 Then Exit Do
            'imports найден
            pos += "@import ".Length
            Dim pos2 As Integer = fileContents.IndexOf(";"c, pos) 'получаем завершение оператора import - ";"
            If pos2 = -1 Then
                Continue Do
            Else
                Dim strPath As String = fileContents.Substring(pos, pos2 - pos).Trim 'получаем путь к дочерней таблице стилей
                If strPath.Length < 3 Then Continue Do
                'убираем кавычки
                Dim ch As Char = strPath.Chars(0)
                If ch = Chr(34) OrElse ch = "'"c Then strPath = strPath.Substring(1)
                ch = strPath.Last
                If ch = Chr(34) OrElse ch = "'"c Then strPath = strPath.Substring(0, strPath.Length - 1)
                'получаем полный путь и запускаем рекурсию
                Dim newPath As String = FileIO.FileSystem.GetFileInfo(pathToCss).DirectoryName
                strPath = FileIO.FileSystem.CombinePath(newPath, strPath)
                ReadCssClasses(strPath, lstClasses)
            End If
        Loop
        'FileIO.FileSystem.GetFileInfo(pathToCss).DirectoryName
        'Ищем классы
        pos = 0
        Do
            pos = fileContents.IndexOf("."c, pos)
            If pos = -1 Then Return

            'нашли точку. Определяем не является ли она началом слова
            Dim pointIsFirstChar As Boolean = False
            If pos > 0 Then
                Dim prevCh As Char = fileContents.Chars(pos - 1)
                If {" "c, vbCr, vbLf, ",", "<", ">"}.Contains(prevCh) Then
                    pointIsFirstChar = True
                End If
            Else
                pointIsFirstChar = True
            End If

            If pointIsFirstChar AndAlso pos < fileContents.Length - 1 Then
                Dim nextCh As Char = fileContents.Chars(pos + 1)
                If IsNumeric(nextCh) Then
                    'это дробное число
                    pos += 1
                    Continue Do
                End If
            End If

            If pointIsFirstChar Then
                'слово начинается с точки. Считаем это классом. Ищем конец слова
                Dim pos2 As Integer = fileContents.IndexOfAny({" "c, ":"c, "{"c, ","c}, pos)
                If pos2 = -1 Then
                    pos += 1
                    Continue Do
                End If
                'получаем имя класса и добавляем его в список
                Dim className As String = fileContents.Substring(pos + 1, pos2 - pos - 1).Trim
                If lstClasses.Contains(className) = False Then lstClasses.Add(className)
                pos = pos2
            Else
                pos += 1
            End If
        Loop
    End Sub

    ''' <summary>
    ''' Создает список анимаций, найденных в данном css-файле (а также в дочерних, переданных с помощью оператора @import)
    ''' </summary>
    ''' <param name="pathToCss">Путь к css-файлу</param>
    ''' <param name="forceUpdate">Принудительно обновить. Если True - то список будет перестроен, если False - то будет возвращен подготовленный заранее список (если он был, иначе создается новый)</param>
    Public Function MakeCssAnimationsList(ByVal pathToCss As String, Optional forceUpdate As Boolean = False) As List(Of String)
        If lstCssAnimations.ContainsKey(pathToCss) Then
            If forceUpdate Then
                lstCssAnimations.Remove(pathToCss)
            Else
                Return lstCssAnimations(pathToCss)
            End If
        End If

        Dim fullPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, pathToCss)
        'получаем список классв
        Dim lstAnimations As New List(Of String)
        ReadCssAnimations(fullPath, lstAnimations)
        'сохраняем в структуре
        lstCssAnimations.Add(pathToCss, lstAnimations)
        Return lstAnimations
    End Function

    ''' <summary>
    ''' Рекурсивная функция, составляющая список анимаций из css-файла, включая другие css-файлы, включенные посредством  @import
    ''' </summary>
    ''' <param name="pathToCss">Полный (и заранее проверенный) путь к css-файлу</param>
    ''' <param name="lstAnimations">Список, в котором будут собираться имена анимаций</param>
    Private Sub ReadCssAnimations(ByVal pathToCss As String, ByRef lstAnimations As List(Of String))
        If FileIO.FileSystem.FileExists(pathToCss) = False Then
            MessageBox.Show("Файл " + pathToCss + " не найден.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        'Считываем файл
        Dim fileContents As String
        Try
            fileContents = My.Computer.FileSystem.ReadAllText(pathToCss, System.Text.Encoding.Default)
        Catch ex As Exception
            MessageBox.Show("Не удалось прочитать файл " + pathToCss + ".", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End Try

        'ищем @import 
        Dim pos As Integer = 0
        Do
            pos = fileContents.IndexOf("@import ", pos)
            If pos = -1 Then Exit Do
            'imports найден
            pos += "@import ".Length
            Dim pos2 As Integer = fileContents.IndexOf(";"c, pos) 'получаем завершение оператора import - ";"
            If pos2 = -1 Then
                Continue Do
            Else
                Dim strPath As String = fileContents.Substring(pos, pos2 - pos).Trim 'получаем путь к дочерней таблице стилей
                If strPath.Length < 3 Then Continue Do
                'убираем кавычки
                Dim ch As Char = strPath.Chars(0)
                If ch = Chr(34) OrElse ch = "'"c Then strPath = strPath.Substring(1)
                ch = strPath.Last
                If ch = Chr(34) OrElse ch = "'"c Then strPath = strPath.Substring(0, strPath.Length - 1)
                'получаем полный путь и запускаем рекурсию
                Dim newPath As String = FileIO.FileSystem.GetFileInfo(pathToCss).DirectoryName
                strPath = FileIO.FileSystem.CombinePath(newPath, strPath)
                ReadCssClasses(strPath, lstAnimations)
            End If
        Loop
        'FileIO.FileSystem.GetFileInfo(pathToCss).DirectoryName
        'Ищем классы
        pos = 0
        Do
            pos = fileContents.IndexOf("@keyframes ", pos)
            If pos = -1 Then Return

            'нашли @keyframes . Определяем не является ли она началом слова
            pos += 11
            Dim pos2 As Integer = fileContents.IndexOfAny({" "c, ":"c, "{"c, ","c}, pos)
            If pos2 = -1 Then
                pos += 1
                Continue Do
            End If
            'получаем имя анимации и добавляем его в список
            Dim aniName As String = fileContents.Substring(pos, pos2 - pos).Trim
            If lstAnimations.Contains(aniName) = False Then lstAnimations.Add(aniName)
            pos = pos2
        Loop
    End Sub

    ''' <summary>
    ''' Возвращает список анимаций из ресурса, определяемого в файле помощи
    ''' </summary>
    ''' <param name="strSource">Исполняемая строка, возвращающая путь к css-файлу</param>
    ''' <param name="strSourceParams">Параметры, передаваемые исполняемой строке. Ключевые слова this и parent заменяются на имена текущего элемента и его родителя</param>
    ''' <param name="strPattern">Паттерн для фильтрации анимаций. Будут выбраны только те, которые начинаются на указанную строку</param>
    ''' <returns>Список анимаций или пустой список</returns>
    Public Function GetCssAnimationsByPattern(ByVal strSource As String, ByVal strSourceParams As String, ByVal strPattern As String, Optional ByRef CSSFile As String = "") As List(Of String)
        CSSFile = ""
        Dim lstAnimations As New List(Of String)
        If String.IsNullOrEmpty(strSource) Then Return lstAnimations
        If IsNothing(cPanelManager.ActivePanel) Then Return lstAnimations

        Dim arrParams() As String = Nothing
        If String.IsNullOrEmpty(strSourceParams) = False Then
            'параметры к исполняемой строке strSource переданы. Получаем их и заменяем ключевые слова на реальные значения
            arrParams = strSourceParams.Split(","c)
            For i As Integer = 0 To arrParams.Count - 1
                Dim strParam As String = arrParams(i).Trim
                If String.Compare(strParam, "parent", True) = 0 Then
                    'получаем имя родителя текущего элемента
                    If cPanelManager.ActivePanel.child3Name.Length > 0 Then
                        strParam = cPanelManager.ActivePanel.child2Name
                    ElseIf cPanelManager.ActivePanel.classId = mScript.mainClassHash("A") AndAlso cPanelManager.ActivePanel.supraElementName.Length > 0 Then
                        strParam = cPanelManager.ActivePanel.supraElementName
                    Else
                        Return lstAnimations
                    End If
                ElseIf String.Compare(strParam, "this", True) = 0 Then
                    'получаем имя текущего элемента
                    If cPanelManager.ActivePanel.child3Name.Length > 0 Then
                        strParam = cPanelManager.ActivePanel.child3Name
                    ElseIf cPanelManager.ActivePanel.child2Name.Length > 0 Then
                        strParam = cPanelManager.ActivePanel.child2Name
                    Else
                        Return lstAnimations
                    End If
                End If
                arrParams(i) = strParam
            Next
        End If

        'Выполняем скрипт в исполняемой строке, и получаем путь к css-файлу относительно папки квеста
        strSource = WrapString(strSource)
        strSource = mScript.PrepareStringToPrint(strSource, arrParams, True)
        If strSource = "#Error" Then Return lstAnimations
        'Создаем/получаем полные список классов
        CSSFile = strSource
        Dim lstFull As List(Of String) = MakeCssAnimationsList(strSource)
        If lstFull.Count = 0 Then Return lstAnimations
        If String.IsNullOrEmpty(strPattern) Then Return lstFull 'фильтр не указан - возвращаем весь список
        'возвращаем отфильтрованный список
        For i As Integer = 0 To lstFull.Count - 1
            If lstFull(i).StartsWith(strPattern, StringComparison.CurrentCultureIgnoreCase) Then
                lstAnimations.Add(lstFull(i))
            End If
        Next
        Return lstAnimations
    End Function


    ''' <summary>
    ''' Возвращает список id всех элементов, найденных в указанном свойстве (на всех 3 уровнях, исходя из cPanelManager.ActivePanel). Свойство должно быть типа LONG_TEXT or CODE
    ''' </summary>
    ''' <param name="strSources">Строка-ресурс из html-файла вида "L.Description,L.Footer" (без кавычек)</param>
    ''' <returns>Список id или пустой список</returns>
    Public Function GetListOfID(ByVal strSources As String) As List(Of String)
        Dim lstId As New List(Of String)
        'Получаем из строки strSource имя свойства и его класс
        Dim classId As Integer, propName As String = ""
        Dim arrSource() As String = Split(strSources, ",")

        For i As Integer = 0 To arrSource.Count - 1
            Dim strSource As String = arrSource(i).Trim
            If GetClassIdAndElementNameByString(WrapString(strSource), classId, propName, False, False) = False Then Continue For

            Dim hidden As MatewScript.PropertyHiddenEnum = mScript.mainClass(classId).Properties(propName).Hidden 'уровни свойства, на которых оно актуально
            Dim res As MatewScript.ContainsCodeEnum
            Dim val As String = ""
            If hidden <> MatewScript.PropertyHiddenEnum.LEVEL2_ONLY AndAlso hidden <> MatewScript.PropertyHiddenEnum.LEVEL23_ONLY AndAlso hidden <> MatewScript.PropertyHiddenEnum.LEVEL3_ONLY Then
                'Получаем свойство 1 уровня
                val = mScript.mainClass(classId).Properties(propName).Value
                res = mScript.IsPropertyContainsCode(val)
                'получаем список id
                If res = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                    ReadIdFromLongText(val, lstId)
                Else
                    If mScript.mainClass(classId).Properties(propName).eventId > 0 Then mScript.eventRouter.MakeElementIdListFromHTMLBlocks(mScript.mainClass(classId).Properties(propName).eventId, lstId)
                End If
            End If

            'Получение значения свойства из элементов 2 и 3 уровня возможно только если они являются текущими на данный момент
            If IsNothing(cPanelManager.ActivePanel) OrElse currentClassName = "Variable" OrElse currentClassName = "Function" Then Continue For
            Dim curClassId As Integer = mScript.mainClassHash(currentClassName)

            Dim child2Id As Integer = -1
            If hidden <> MatewScript.PropertyHiddenEnum.LEVEL1_ONLY AndAlso hidden <> MatewScript.PropertyHiddenEnum.LEVEL13_ONLY AndAlso hidden <> MatewScript.PropertyHiddenEnum.LEVEL3_ONLY Then
                If curClassId = classId AndAlso cPanelManager.ActivePanel.child2Name.Length > 0 Then
                    'Получаем свойство 2 уровня
                    child2Id = cPanelManager.ActivePanel.GetChild2Id
                ElseIf currentClassName = "A" AndAlso classId = mScript.mainClassHash("L") AndAlso currentParentName.Length > 0 Then
                    'Получаем свойство локации ели сейчас выбраны действия
                    child2Id = currentParentId
                End If
            End If

            If child2Id > -1 Then
                'Элемент 2 уровня был найден - получаем из его свойства список id
                val = mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value
                res = mScript.IsPropertyContainsCode(val)
                If res = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                    ReadIdFromLongText(val, lstId)
                Else
                    If mScript.mainClass(classId).ChildProperties(child2Id)(propName).eventId > 0 Then mScript.eventRouter.MakeElementIdListFromHTMLBlocks(mScript.mainClass(classId).ChildProperties(child2Id)(propName).eventId, lstId)
                End If
            End If

            If classId = curClassId AndAlso mScript.mainClass(classId).LevelsCount = 2 AndAlso hidden <> MatewScript.PropertyHiddenEnum.LEVEL1_ONLY AndAlso hidden <> MatewScript.PropertyHiddenEnum.LEVEL12_ONLY _
                AndAlso hidden <> MatewScript.PropertyHiddenEnum.LEVEL2_ONLY AndAlso cPanelManager.ActivePanel.child3Name.Length > 0 Then
                Dim child3Id As Integer = cPanelManager.ActivePanel.GetChild3Id(child2Id)
                If child3Id > -1 Then
                    'Получаем свойство 3 уровня
                    val = mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id)
                    'получаем список id из свойства элемента 3 уровня
                    If mScript.IsPropertyContainsCode(val) = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                        ReadIdFromLongText(val, lstId)
                    Else
                        If mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelEventId(child3Id) > 0 Then mScript.eventRouter.MakeElementIdListFromHTMLBlocks(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelEventId(child3Id), lstId)
                    End If
                End If
            End If
        Next

        'Возвращаем список id
        Return lstId
    End Function

    ''' <summary>
    ''' Из содержимого свойства типа LONG_TEXT собирает список id html-элементов
    ''' </summary>
    ''' <param name="strText">Сериализованна строка длинного html-текста</param>
    ''' <param name="lstId">Список для сохранения id</param>
    ''' <remarks></remarks>
    Private Sub ReadIdFromLongText(ByVal strText As String, ByRef lstId As List(Of String))
        'Ищем строки вида id = "xxx", и сохраняем xxx в список lstId
        Dim pos As Integer = 0
        Do
            'Находим строку, которая начинается на id
            pos = strText.IndexOf(" id", pos, System.StringComparison.CurrentCultureIgnoreCase)
            If pos = -1 Then Exit Do
            pos += 3

            'за id ищем вначале =, а затем кавычку
            Dim equalFound As Boolean = False
            Dim brChar As Char = "" 'для хранения вида кавычки - " или '
            For i As Integer = pos To strText.Length - 1                
                Dim ch As Char = strText.Chars(i)
                If ch = " "c OrElse ch = vbLf OrElse ch = vbCr Then Continue For 'пробелы пропускаем
                If equalFound = False AndAlso ch = "=" Then
                    'вначале нашли =. Далее на нужна кавычка
                    equalFound = True
                ElseIf equalFound AndAlso (ch = "'"c OrElse ch = Chr(34)) Then
                    'после = нашли кавычку.
                    pos = i + 1
                    brChar = ch
                    Exit For
                Else
                    Continue Do
                End If
            Next
            'ищем закрывающую кавычку
            Dim pos2 As Integer = strText.IndexOf(brChar, pos)
            If pos2 = -1 Then Continue Do
            'получаем текст в кавычках - id элемента, и сохраняем в списке
            Dim strId As String = strText.Substring(pos, pos2 - pos)
            If lstId.Contains(strId) = False Then lstId.Add(strId)
            pos = pos2
        Loop
    End Sub


    ''' <summary>Определяет есть ли в папке квеста изображения, аудио или другие файлы</summary>
    Public Function HasFiles(ByVal fType As fileTypesEnum) As Boolean
        Select Case fType
            Case fileTypesEnum.PICTURES
                If IsNothing(arrPictures) OrElse arrPictures.Count = 0 Then Return False
            Case fileTypesEnum.AUDIO
                If IsNothing(arrAudio) OrElse arrAudio.Count = 0 Then Return False
            Case fileTypesEnum.TEXT
                If IsNothing(arrText) OrElse arrText.Count = 0 Then Return False
            Case fileTypesEnum.CSS
                If IsNothing(arrCSS) OrElse arrCSS.Count = 0 Then Return False
            Case fileTypesEnum.JS
                If IsNothing(arrJS) OrElse arrJS.Count = 0 Then Return False
        End Select
        Return True
    End Function

    ''' <summary>Обновляет списки файлов после изменения в директории квеста</summary>
    Public Sub UpdateFileLists()
        Erase arrPictures
        Erase arrAudio
        Erase arrText
        Erase arrCSS
        Erase arrJS

        Dim files As ReadOnlyCollection(Of String)
        files = My.Computer.FileSystem.GetFiles(questEnvironment.QuestPath, FileIO.SearchOption.SearchAllSubDirectories, "*.png", "*.jpg", "*.gif", "*.jpeg", "*.bmp", "*.wmf", "*.mp3", "*.wav", "*.mid", "*.txt", "*.js", "*.css")
        If IsNothing(files) OrElse files.Count = 0 Then Return
        Dim lstPictures As New List(Of String)
        Dim lstAudio As New List(Of String)
        Dim lstText As New List(Of String)
        Dim lstCSS As New List(Of String)
        Dim lstJS As New List(Of String)
        For i As Integer = 0 To files.Count - 1
            Dim fName As String = files(i).Substring(questEnvironment.QuestPath.Length + 1)
            Select Case System.IO.Path.GetExtension(fName)
                Case ".png", ".jpg", ".gif", ".jpeg", ".bmp", ".wmf"
                    lstPictures.Add(fName)
                Case ".mp3", ".wav", ".mid"
                    lstAudio.Add(fName)
                Case ".txt"
                    lstText.Add(fName)
                Case ".css"
                    lstCSS.Add(fName)
                Case ".js"
                    lstJS.Add(fName)
            End Select
        Next i
        arrPictures = lstPictures.ToArray
        arrAudio = lstAudio.ToArray
        arrText = lstText.ToArray
        arrCSS = lstCSS.ToArray
        arrJS = lstJS.ToArray
    End Sub

    ''' <summary>Обновляет все списки</summary>
    Public Sub UpdateElementsLists()
        lstMain = New SortedList(Of Integer, List(Of String))

        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'вставляем только имена второго уровня
            If mScript.mainClass(classId).LevelsCount = 0 Then Continue For
            If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Continue For
            If mScript.mainClass(classId).Properties.ContainsKey("Name") = False Then Continue For
            'Создаем список имен в lstNames
            Dim lstNames As New List(Of String)
            For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(i)("Name")
                If p.Hidden Then Continue For
                Dim pName As String = p.Value
                If String.IsNullOrWhiteSpace(pName) Then Continue For
                pName = pName.Substring(1, pName.Length - 2) 'убирем кавычки
                lstNames.Add(pName)
            Next i
            If IsNothing(lstNames) OrElse lstNames.Count = 0 Then Continue For 'все скрыты
            'собственно, создаем список имен класса classId
            lstNames.Sort()
            lstMain.Add(classId, lstNames)
        Next classId
    End Sub

    ''' <summary>Обновляет список элементов указанного класса</summary>
    ''' <param name="className">Класс, список элементов которого надо обновить</param>
    Public Sub UpdateList(ByVal className As String)
        Dim classId As Integer = -1
        If mScript.mainClassHash.TryGetValue(className, classId) = False Then Return
        lstMain(classId) = New List(Of String)

        'вставляем только имена второго уровня
        If mScript.mainClass(classId).LevelsCount = 0 Then Return
        If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Return
        If mScript.mainClass(classId).Properties.ContainsKey("Name") = False Then Return
        'Создаем спосок имен в lstNames
        Dim lstNames As New List(Of String)
        For i As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
            Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(i)("Name")
            If p.Hidden Then Continue For
            Dim pName As String = p.Value
            If String.IsNullOrWhiteSpace(pName) Then Continue For
            pName = pName.Substring(1, pName.Length - 2) 'убирем кавычки
            lstNames.Add(pName)
        Next i
        If IsNothing(lstNames) OrElse lstNames.Count = 0 Then Return 'все скрыты
        'собственно, создаем список имен класса classId
        lstNames.Sort()
        lstMain(classId) = lstNames
    End Sub

    Public Sub AddNameToList(ByVal className As String, ByVal newName As String)
        Dim classId As Integer = -1
        If mScript.mainClassHash.TryGetValue(className, classId) = False Then Return
        If lstMain.ContainsKey(classId) = False Then lstMain.Add(classId, New List(Of String))
        If lstMain(classId).Contains(newName) Then Return
        lstMain(classId).Add(newName)
        lstMain(classId).Sort()
    End Sub

    Public Sub RemoveNameFromList(ByVal className As String, ByVal elName As String)
        Dim classId As Integer = -1
        If mScript.mainClassHash.TryGetValue(className, classId) = False Then Return
        lstMain(classId).Remove(elName)
    End Sub

    Public Sub RenameElementInList(ByVal classId As Integer, ByVal oldName As String, ByVal newName As String)
        If String.IsNullOrWhiteSpace(oldName) OrElse String.IsNullOrWhiteSpace(newName) Then Return
        'oldName = oldName.Substring(1, oldName.Length - 2)
        'newName = newName.Substring(1, newName.Length - 2)
        If lstMain.ContainsKey(classId) = False Then Return
        Dim res As Integer = lstMain(classId).BinarySearch(oldName)
        If res > -1 Then lstMain(classId)(res) = newName
        lstMain(classId).Sort()
    End Sub

    ''' <summary>Функция заполняет указаный комбо/листбокс списком, который положен согласно имени свойства и его типа</summary>
    ''' <param name="propName">Имя свойства</param>
    ''' <param name="childProp">Ссылка на панель, в котором находится данный контрол</param>
    ''' <param name="lstToFill">Контрол для заполнения</param>
    ''' <param name="dontTouchIfHasArray">Если свойство хранит в returnArray массив, то определяет стоит ли заменять список в соответствующем контроле (поскольку он меняется крайне редко)</param>
    ''' <returns>True если был заполнен, False если пусто</returns>
    Public Function FillListByChildPanel(ByVal propName As String, ByRef childProp As clsPanelManager.clsChildPanel, ByRef lstToFill As Object, Optional dontTouchIfHasArray As Boolean = True) As Boolean
        Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(childProp.classId).Properties(propName)
        Dim retType As MatewScript.ReturnFunctionEnum = p.returnType
        Dim curText As String = lstToFill.Text

        If IsPropertyContainsEnum(p) Then
            If dontTouchIfHasArray Then Return True
            lstToFill.Items.Clear()
            Dim arrToAdd() As String
            ReDim arrToAdd(p.returnArray.Count - 1)
            For i As Integer = 0 To arrToAdd.Count - 1
                arrToAdd(i) = mScript.PrepareStringToPrint(p.returnArray(i), Nothing, False)
            Next
            lstToFill.Items.AddRange(arrToAdd)
            lstToFill.Text = curText
            Return True
        End If
        If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_BOOl AndAlso dontTouchIfHasArray Then Return True

        lstToFill.Items.Clear()
        Select Case retType
            Case MatewScript.ReturnFunctionEnum.RETURN_ELEMENT
                Dim classToShow As Integer = -1
                If IsNothing(p.returnArray) OrElse p.returnArray.Count = 0 Then lstToFill.Text = curText : Return False
                If p.returnArray(0) = "Variable" Then
                    lstToFill.Items.AddRange(mScript.csPublicVariables.GetVariablesList.ToArray)
                    If lstToFill.Items.Count = 0 Then Return False
                Else
                    classToShow = mScript.mainClassHash(p.returnArray(0))
                    If lstMain.ContainsKey(classToShow) = False Then lstToFill.Text = curText : Return False
                    If IsNothing(lstMain(classToShow)) OrElse lstMain(classToShow).Count = 0 Then lstToFill.Text = curText : Return False
                    lstToFill.Items.AddRange(lstMain(classToShow).ToArray)
                End If
                Return True
            Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                lstToFill.Items.AddRange({"True", "False"})
                lstToFill.Text = curText
                Return True
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE
                If IsNothing(arrPictures) OrElse arrPictures.Count = 0 Then lstToFill.Text = curText : Return False
                lstToFill.Items.AddRange(arrPictures)
                lstToFill.Text = curText
                Return True
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO
                If IsNothing(arrAudio) OrElse arrAudio.Count = 0 Then lstToFill.Text = curText : Return False
                lstToFill.Items.AddRange(arrAudio)
                lstToFill.Text = curText
                Return True
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_TEXT
                If IsNothing(arrText) OrElse arrText.Count = 0 Then lstToFill.Text = curText : Return False
                lstToFill.Items.AddRange(arrText)
                lstToFill.Text = curText
                Return True
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS
                If IsNothing(arrCSS) OrElse arrCSS.Count = 0 Then lstToFill.Text = curText : Return False
                lstToFill.Items.AddRange(arrCSS)
                lstToFill.Text = curText
                Return True
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_JS
                If IsNothing(arrJS) OrElse arrJS.Count = 0 Then lstToFill.Text = curText : Return False
                lstToFill.Items.AddRange(arrJS)
                lstToFill.Text = curText
                Return True
            Case MatewScript.ReturnFunctionEnum.RETURN_COLOR
                If questEnvironment.lstSelectedColors.Count = 0 Then lstToFill.Text = curText : Return False
                lstToFill.Items.AddRange(questEnvironment.lstSelectedColors.ToArray)
                lstToFill.Text = curText
                Return True
            Case MatewScript.ReturnFunctionEnum.RETURN_FUNCTION
                If IsNothing(mScript.functionsHash) OrElse mScript.functionsHash.Count = 0 Then lstToFill.Text = curText : Return False
                lstToFill.Items.AddRange(mScript.functionsHash.Keys.ToArray)
                lstToFill.Text = curText
                Return True
        End Select
        lstToFill.Text = curText
        Return False
    End Function

    ''' <summary>
    ''' Возвращает список элементов 2 порядка по имени их класса
    ''' </summary>
    ''' <param name="className">Имя класса, элементы которого надо вернуть (в 'кавычках')</param>
    Public Function FillListByClassName(ByVal className As String) As String()
        If IsNothing(lstMain) OrElse lstMain.Count = 0 Then Return {}
        If className = "Variable" Then Return mScript.csPublicVariables.GetVariablesList(True).ToArray
        Dim classId As Integer = -1
        If mScript.mainClassHash.TryGetValue(className, classId) = False Then Return {}
        Dim lstCopy As List(Of String) = Nothing
        If lstMain.TryGetValue(classId, lstCopy) = False Then Return {}
        If IsNothing(lstCopy) OrElse lstCopy.Count = 0 Then Return {}

        Dim arrNew(lstCopy.Count - 1) As String
        For i As Integer = 0 To lstCopy.Count - 1
            arrNew(i) = WrapString(lstCopy(i))
        Next
        Return arrNew
    End Function

    ''' <summary>
    ''' Создает список элементов 3 уровня
    ''' </summary>
    ''' <param name="classId">класс элементов</param>
    ''' <param name="child2Id">родитель 2 уровня для данных элементов</param>
    ''' <returns>список элементов или пустой список</returns>
    ''' <remarks></remarks>
    Public Function CreateThirdLevelElementsList(ByVal classId As Integer, ByVal child2Id As Integer) As String()
        If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Return {}
        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)("Name")
        If IsNothing(ch.ThirdLevelProperties) OrElse ch.ThirdLevelProperties.Count = 0 Then Return {}
        Return ch.ThirdLevelProperties
    End Function

    ''' <summary>
    ''' Возвращает список функций Писателя
    ''' </summary>
    ''' <param name="wrapNames">Заварчивать ли фукнции в кавычки</param>
    Public Function GetFunctionsList(Optional wrapNames As Boolean = True) As String()
        If IsNothing(mScript.functionsHash) OrElse mScript.functionsHash.Count = 0 Then Return {}
        If Not wrapNames Then Return mScript.functionsHash.Keys
        Dim arr() As String
        Dim arrKeys() = mScript.functionsHash.Keys.ToArray
        ReDim arr(arrKeys.Count - 1)
        For i As Integer = 0 To arrKeys.Count - 1
            arr(i) = "'" + arrKeys(i) + "'"
        Next
        Return arr
    End Function

    ''' <summary>
    ''' Возвращает список всех свойств
    ''' </summary>
    ''' <param name="wrapNames">заворачивать ли в 'кавычки'</param>
    ''' <returns></returns>
    Public Function GetPropertiesList(Optional wrapNames As Boolean = True) As String()
        Dim lastRange() As String = Nothing

        Dim rangeSize As Integer = 0
        For i As Integer = 0 To mScript.mainClass.GetUpperBound(0)
            If IsNothing(mScript.mainClass(i).Properties) Then Continue For
            Dim className As String = mScript.mainClass(i).Names(0)
            For Each prop As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(i).Properties
                If prop.Value.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse prop.Value.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_CODE Then Continue For
                'проверка каждого свойства в каждом классе
                rangeSize += 1
                Array.Resize(Of String)(lastRange, rangeSize)
                If wrapNames Then
                    lastRange(rangeSize - 1) = "'" + className + "." + prop.Key + "'"
                Else
                    lastRange(rangeSize - 1) = className + "." + prop.Key
                End If
            Next
        Next

        Return lastRange
    End Function

    ''' <summary>
    ''' Возвращает список всех свойств-событий
    ''' </summary>
    ''' <param name="wrapNames">заворачивать ли в 'кавычки'</param>
    ''' <returns></returns>
    Public Function GetEventsList(Optional wrapNames As Boolean = True) As String()
        Dim lastRange() As String = Nothing

        Dim rangeSize As Integer = 0
        For i As Integer = 0 To mScript.mainClass.GetUpperBound(0)
            If IsNothing(mScript.mainClass(i).Properties) Then Continue For
            Dim className As String = mScript.mainClass(i).Names(0)
            For Each prop As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(i).Properties
                If prop.Value.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse prop.Value.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_CODE Then Continue For
                'проверка каждого свойства в каждом классе
                If prop.Value.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                    'тип свойства - событие. добавляе его в список.
                    rangeSize += 1
                    Array.Resize(Of String)(lastRange, rangeSize)
                    If wrapNames Then
                        lastRange(rangeSize - 1) = "'" + className + "." + prop.Key + "'"
                    Else
                        lastRange(rangeSize - 1) = className + "." + prop.Key
                    End If
                End If
            Next
        Next

        Return lastRange
    End Function

    ''' <summary>
    ''' Возвращает список файлов указанного типа, взятых в кавычки '', с путями относительно папки квеста
    ''' </summary>
    ''' <param name="filesType">Тип файлов</param>
    ''' <param name="wrapNames">Заварчивать ли файлы в кавычки</param>
    Public Function GetFilesList(ByVal filesType As fileTypesEnum, Optional wrapNames As Boolean = True) As String()
        Dim arrWork() As String = Nothing
        Select Case filesType
            Case fileTypesEnum.PICTURES
                arrWork = arrPictures.ToArray
            Case fileTypesEnum.AUDIO
                arrWork = arrAudio.ToArray
            Case fileTypesEnum.TEXT
                arrWork = arrText.ToArray
            Case fileTypesEnum.CSS
                arrWork = arrCSS.ToArray
            Case fileTypesEnum.JS
                arrWork = arrJS.ToArray
        End Select

        If IsNothing(arrWork) OrElse arrWork.Count = 0 Then Return {}
        If Not wrapNames Then Return arrWork

        For i As Integer = 0 To arrWork.Count - 1
            arrWork(i) = WrapString(arrWork(i))
        Next
        Return arrWork
    End Function

    ''' <summary>
    ''' Создает список выбранных цветов, завернутых в кавычки
    ''' </summary>
    Public Function GetSelectedColorsList() As String()
        If questEnvironment.lstSelectedColors.Count = 0 Then Return {}
        Dim arrWork() As String
        ReDim arrWork(questEnvironment.lstSelectedColors.Count - 1)
        For i As Integer = 0 To arrWork.Count - 1
            arrWork(i) = WrapString(questEnvironment.lstSelectedColors(i))
        Next
        Return arrWork
    End Function

    ''' <summary>
    ''' Создает список папок относительно папки квеста (только тех, в которых есть хоть один файл указанного типа)
    ''' </summary>
    ''' <param name="filesType">Тип файлов</param>
    Public Function CreateFoldersList(ByVal filesType As fileTypesEnum) As String()
        Dim arrWork() As String = Nothing
        Select Case filesType
            Case fileTypesEnum.PICTURES
                arrWork = arrPictures
            Case fileTypesEnum.AUDIO
                arrWork = arrAudio
            Case fileTypesEnum.TEXT
                arrWork = arrText
            Case fileTypesEnum.CSS
                arrWork = arrCSS
            Case fileTypesEnum.JS
                arrWork = arrJS
        End Select

        If IsNothing(arrWork) OrElse arrWork.Count = 0 Then Return {}
        Dim lstFolders As New List(Of String)

        Dim fPath As String
        For i As Integer = 0 To arrWork.Count - 1
            Dim aPos As Integer = arrWork(i).LastIndexOf("\"c)
            If aPos = -1 Then
                fPath = ""
            Else
                fPath = arrWork(i).Substring(0, aPos)
            End If
            If lstFolders.Contains(fPath) = False Then lstFolders.Add(fPath)
        Next

        Return lstFolders.ToArray
    End Function

    ''' <summary>
    ''' Восзвращает список файлов без пути указанного типа из указанной директории
    ''' </summary>
    ''' <param name="filesType">тип файлов</param>
    ''' <param name="folder">папка относительно папки квеста</param>
    Public Function GetFolderFiles(ByVal filesType As fileTypesEnum, ByVal folder As String) As List(Of String)
        Dim arrWork() As String = Nothing
        Select Case filesType
            Case fileTypesEnum.PICTURES
                arrWork = arrPictures
            Case fileTypesEnum.AUDIO
                arrWork = arrAudio
            Case fileTypesEnum.TEXT
                arrWork = arrText
            Case fileTypesEnum.CSS
                arrWork = arrCSS
            Case fileTypesEnum.JS
                arrWork = arrJS
        End Select

        Dim lstFiles As New List(Of String)
        If IsNothing(arrWork) OrElse arrWork.Count = 0 Then Return lstFiles
        Dim fPath As String
        For i As Integer = 0 To arrWork.Count - 1
            Dim aPos As Integer = arrWork(i).LastIndexOf("\"c)
            If aPos = -1 Then
                fPath = ""
            Else
                fPath = arrWork(i).Substring(0, aPos)
            End If
            If fPath = folder Then
                If fPath.Length > 0 Then
                    lstFiles.Add(arrWork(i).Substring(fPath.Length + 1))
                Else
                    lstFiles.Add(arrWork(i).Substring(fPath.Length))
                End If
            End If
        Next
        Return lstFiles
    End Function

    ''' <summary>
    ''' Заменяет в одном из главных html-файлов из папки квеста ссылку на таблицу css.
    ''' </summary>
    ''' <param name="classId">Класс свойства, в котором хранится путь к css</param>
    ''' <param name="child2Id">Id элемента 2 уровня, в свойстве которого хранится путь к css</param>
    ''' <param name="child3Id">Id элемента 3 уровня, в свойстве которого хранится путь к css</param>
    ''' <param name="propName">Имя свойства с названием файла css</param>
    ''' <param name="newValue">Новое значени свойства - путь к файлу css относительно папки квеста (в кавычках)</param>
    Public Sub UpdateCSSInFile(ByVal classId As Integer, ByVal child2Id As Integer, ByVal child3Id As Integer, ByVal propName As String, ByVal newValue As String)
        'Получаем старое значение свойства
        If questEnvironment.EDIT_MODE = False Then Return 'в режиме игры файл не меняем
        Dim oldValue As String
        If child2Id < 0 Then
            oldValue = mScript.mainClass(classId).Properties(propName).Value
        ElseIf child3Id < 0 Then
            oldValue = mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value
        Else
            oldValue = mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id)
        End If
        'Если старое и новое значения совпадают - выход
        If newValue = oldValue Then Return
        If lstCssClasses.ContainsKey(oldValue) Then
            lstCssClasses.Remove(oldValue) 'удаляем список классов из заменяемого css-файла
        End If

        newValue = mScript.PrepareStringToPrint(newValue, Nothing).Replace("\", "/")

        'получаем полный путь к html-файлу
        Dim className As String = mScript.mainClass(classId).Names(0)
        Dim fileToChange As String
        Select Case className
            Case "A", "AW"
                fileToChange = "Actions.html"
            Case "O", "OW"
                fileToChange = "Objects.html"
            Case "D", "DW"
                fileToChange = "Description.html"
            Case "Cm"
                fileToChange = "Command.html"
            Case "B"
                fileToChange = "Battle.html"
            Case "Mg"
                fileToChange = "Magic.html"
            Case "Map"
                fileToChange = "Map.html"
            Case Else
                fileToChange = "Location.html"
        End Select

        Dim filePath As String = questEnvironment.QuestPath + "\" + fileToChange
        If FileIO.FileSystem.FileExists(filePath) = False Then
            MessageBox.Show(String.Format("В папке квеста не найден необходимый для работы файл {0}. Для восстановления работоспособности скопируйте его из [папки игры]\src\defaultFiles", fileToChange), "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        'считываем все его содержимое
        Dim fileContent As String = FileIO.FileSystem.ReadAllText(filePath, System.Text.Encoding.Default)
        Dim pos As Integer = 0
        Do
            pos = fileContent.IndexOf("<link", pos, System.StringComparison.CurrentCultureIgnoreCase)
            If pos = -1 Then Exit Do
            Dim posEnd As Integer = fileContent.IndexOf(">", pos)
            If posEnd = -1 Then
                MessageBox.Show(String.Format("Файл {0} имеет неправильную структуру. Для восстановления работоспособности скопируйте его из [папки игры]\src\defaultFiles", fileToChange), "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            If fileContent.IndexOf("stylesheet", pos, posEnd - pos, StringComparison.CurrentCultureIgnoreCase) = -1 Then Continue Do
            pos = fileContent.IndexOf("href", pos, posEnd - pos, StringComparison.CurrentCultureIgnoreCase)
            'найдена строка, которая начинается на <link и в которой содержится stylesheet и href. После href ищем =, а после равно - кавычку
            If pos = -1 Then Continue Do
            pos += 4
            Dim quotesFound As Boolean = False
            For i As Integer = pos To posEnd
                Dim ch As Char = fileContent.Chars(i)
                If ch = " "c OrElse ch = vbCr OrElse ch = vbLf Then Continue For
                If quotesFound = False AndAlso ch = "="c Then
                    quotesFound = True 'равно найдено
                ElseIf quotesFound AndAlso (ch = Chr(34) OrElse ch = "'") Then
                    'за знаком = найдена первая кавычка, после которой начинается название файла
                    pos = i + 1 'начало имени файла
                    posEnd = fileContent.IndexOf(ch, pos, posEnd - pos) 'конец имени файла
                    If posEnd = -1 Then Continue Do
                    fileContent = fileContent.Substring(0, pos) & newValue & fileContent.Substring(posEnd) 'новое содержимое файла с измененным путем к css
                    FileIO.FileSystem.WriteAllText(filePath, fileContent, False, System.Text.Encoding.Default) 'сохраняем файл с измененным путем

                    Return
                End If
            Next
        Loop

        'Ссылки не найдено - добавляем ее
        pos = fileContent.IndexOf("</head>", 0, System.StringComparison.CurrentCultureIgnoreCase)
        If pos <= 0 Then
            MessageBox.Show(String.Format("Файл {0} имеет неправильную структуру. Для восстановления работоспособности скопируйте его из [папки игры]\src\defaultFiles", fileToChange), "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        'Перед окончанием заголовка добавляем строку вида <link rel='stylesheet' type='text/css' href='css/location.css'/>
        fileContent = fileContent.Substring(0, pos - 1) & "<link rel='stylesheet' type='text/css' href='" & newValue & "'/>" & fileContent.Substring(pos)
        FileIO.FileSystem.WriteAllText(filePath, fileContent, False, System.Text.Encoding.Default)
    End Sub

    ''' <summary>
    ''' Fills lstDest with properties which returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION. For ex., L.Description
    ''' </summary>
    ''' <param name="arrDest">Destination list. Previous content won't be erased</param>
    ''' <remarks></remarks>
    Public Sub MakeLongTextPropertiesList(ByRef arrDest() As String)
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If IsNothing(mScript.mainClass(classId).Properties) Then Continue For
            For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                Dim itm As KeyValuePair(Of String, MatewScript.PropertiesInfoType) = mScript.mainClass(classId).Properties.ElementAt(i)
                If itm.Value.returnType = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION Then
                    If IsNothing(arrDest) Then
                        ReDim arrDest(0)
                    Else
                        ReDim Preserve arrDest(UBound(arrDest) + 1)
                    End If
                    arrDest(UBound(arrDest)) = mScript.mainClass(classId).Names(0) & "." & itm.Key
                End If
            Next i
        Next classId
    End Sub
End Class
