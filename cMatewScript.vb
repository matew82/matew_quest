Imports System.Deployment
Imports System.IO
Imports System.Security.Cryptography
Imports System.ServiceProcess
Imports System.Globalization

Public Class MatewScript
#Region "Declarations"
    Public Enum ContainsCodeEnum As Byte
        NOT_CODE = 0
        CODE = 1
        EXECUTABLE_STRING = 2
        LONG_TEXT = 3
    End Enum
    ''' <summary>
    ''' Типы выражений
    ''' </summary>
    Public Enum WordTypeEnum As Byte
        W_NOTHING = 0
        W_SIMPLE_NUMBER = 1
        W_SIMPLE_STRING = 2
        W_SIMPLE_BOOL = 3
        W_VARIABLE_LOCAL = 4
        W_VARIABLE_PUBLIC = 5 ' С ПОМОЩЬЮ Global
        W_OPERATOR_MATH = 6 '+ - * ^ / \
        W_OPERATOR_STRINGS_MERGER = 7 '&
        W_OPERATOR_COMPARE = 8 '= != < > <> >= <=
        W_OPERATOR_LOGIC = 9 'And, Or, Xor
        W_PROPERTY = 10
        W_FUNCTION = 11
        W_PARAM = 12 'Param[x]
        W_PARAM_COUNT = 13 'ParamCount
        W_EXIT = 14
        W_SWITCH = 15
        W_WRAP = 16
        W_HTML = 17
        W_BLOCK_DOWHILE = 18
        W_BLOCK_FUNCTION = 19
        W_BLOCK_NEWCLASS = 20
        W_BLOCK_EVENT = 21
        W_BLOCK_FOR = 22
        W_BLOCK_IF = 23
        W_REM_CLASS = 24
        W_JUMP = 25
        W_RETURN = 26
        W_MARK = 27
        W_CYCLE_INTERNAL = 28
        W_BREAK = 29
        W_CONTINUE = 30
        'W_VARIABLE_EXAM = 31 ' С ПОМОЩЬЮ ExamOnly
    End Enum

    ''' <summary>
    ''' Класс, содержащий готовый к исполнению код
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ExecuteDataType
        Public Code() As CodeTextBox.EditWordType
        Public lineId As Integer

        Public Sub New()
            lineId = 0
        End Sub

        Public Sub New(ByRef Code() As CodeTextBox.EditWordType, ByVal lineId As Integer)
            Me.Code = Code
            Me.lineId = lineId
        End Sub

        ''' <summary>
        ''' Собирает из массива <see cref="Code"></see> строку
        ''' </summary>
        ''' <returns>Строку кода или пустую строку, если массив Code пуст</returns>
        Public Function CompileString() As String
            If IsNothing(Code) OrElse Code.Length = 0 Then Return ""
            Dim s As New System.Text.StringBuilder
            For i As Integer = 0 To Code.GetUpperBound(0)
                s.Append(Code(i).Word)
            Next
            Return s.ToString
        End Function

        Public Function Clone() As ExecuteDataType
            Dim other As ExecuteDataType = DirectCast(Me.MemberwiseClone, ExecuteDataType)
            If IsNothing(Code) OrElse Code.Count = 0 Then Return other

            Dim tmpArr() As CodeTextBox.EditWordType
            ReDim tmpArr(Code.Count - 1)
            Array.Copy(Code, tmpArr, Code.Length)
            other.Code = tmpArr
            Return other
        End Function
    End Class

    ''' <summary>
    ''' Формат, в котором возвращается результат (строка, число, Да/Нет)
    ''' </summary>
    Public Enum ReturnFormatEnum As Byte
        ORIGINAL = 0
        TO_STRING = 1
        TO_NUMBER = 2
        TO_BOOL = 3
        TO_CODE = 4
        TO_LONG_TEXT = 5
        TO_EXECUTABLE_STRING = 6
        TO_ARRAY = 7
    End Enum

    ''' <summary>Класс глобальных переменных</summary>
    Public csPublicVariables As New cVariable
    ''' <summary>Класс локальных переменных</summary>
    Public csLocalVariables As New cVariable
    ''' <summary>стэк кода (для отображения в информации об ошибке в AnalizeError)</summary>
    Public codeStack As New Stack(Of String)
    ''' <summary>Класс управления событиями, содержащий все скрипты в расшифрованном виде</summary>
    Public eventRouter As EventRouterClass
    ''' <summary>Класс управления событиями изменения свойства</summary>
    Public trackingProperties As New cTrackingProperties

    ''' <summary>Был ли код прерван каким-либо оператором или в случае ошибки</summary>
    Public EXIT_CODE As Boolean = False
    ''' <summary>Текст последней ошибки</summary>
    Public LAST_ERROR As String = ""
    ''' <summary>Индекс текущей строки в masCode. Нужно для отслеживании строки с ошибкой при рекурсивном вызове ExecuteCode</summary>
    Public CURRENT_LINE As Integer = 0
    ''' <summary>Содержимое переданного с помощью #ARRAY массива</summary>
    Public lastArray As cVariable.variableEditorInfoType
    ''' <summary>Основа для формиррования имени переменной из временного массива (для присваивания)</summary>
    Private temporary_system_array_BASE As String = Chr(5) & "#temporary_system_array_"
    ''' <summary>Основа для формиррования имени переменной из временного массива (для проверки, с ' в начале)</summary>
    Private temporary_system_array_BASE_check As String = "'" & Chr(5) & "#temporary_system_array_"
#End Region

#Region "Classes Declarations"
    Public Class funcAndPropHashType
        Enum funcOrPropEnum
            E_PROPERTY = 0
            E_FUNCTION = 1
        End Enum
        Public elementName As String 'имя функции или свойства (в правильном регистре)
        Public classId As Integer 'Id класса в mainClass
        Public elementType As funcOrPropEnum
    End Class

    Public Class paramsType
        Public Enum paramsTypeEnum
            PARAMS_ARRAY = 0
            PARAM_INTEGER = 1
            PARAM_SINGLE = 2
            PARAM_STRING = 3
            PARAM_STRING_OR_NUM = 4
            PARAM_BOOL = 5
            PARAM_ANY = 6
            PARAM_ENUM = 7
            PARAM_USER_FUNCTION = 8
            PARAM_ELEMENT = 9
            PARAM_ELEMENT2 = 10
            PARAM_VARIABLE = 11
            PARAM_PATH_PICTURE = 12
            PARAM_PATH_AUDIO = 13
            PARAM_PATH_TEXT = 14
            PARAM_PATH_CSS = 15
            PARAM_PATH_JS = 16
            PARAM_PROPERTY = 17
            PARAM_EVENT = 18
        End Enum
        Public Name As String
        Public Description As String
        Public Type As paramsTypeEnum
        Public EnumValues() As String

        Public Function Clone() As paramsType
            Dim newObj As paramsType = DirectCast(Me.MemberwiseClone, paramsType)
            If IsNothing(EnumValues) = False Then newObj.EnumValues = EnumValues.Clone
            Return newObj
        End Function
    End Class

    Public Enum ReturnFunctionEnum
        RETURN_USUAL = 0
        RETURN_BOOl = 1
        RETURN_ENUM = 2
        RETURN_EVENT = 3
        RETURN_DESCRIPTION = 4
        RETURN_ELEMENT = 5
        RETURN_PATH_PICTURE = 6
        RETURN_PATH_AUDIO = 7
        RETURN_PATH_TEXT = 8
        RETURN_PATH_CSS = 9
        RETURN_PATH_JS = 10
        RETURN_COLOR = 11
        RETURN_FUNCTION = 12
    End Enum

    Public Enum PropertyHiddenEnum As Byte
        NOT_HIDDEN = 0
        HIDDEN_IN_EDITOR = 1
        HIDDEN_IN_CODE = 2
        HIDDEN_AT_ALL = 3
        LEVEL1_ONLY = 4
        LEVEL2_ONLY = 5
        LEVEL3_ONLY = 6
        LEVEL12_ONLY = 7
        LEVEL13_ONLY = 8
        LEVEL23_ONLY = 9
    End Enum

    'структура для хранения функций и свойств 1-го порядка
    Public Class PropertiesInfoType
        Public Value As String
        Public UserAdded As Boolean
        Public Hidden As PropertyHiddenEnum
        Public helpFile As String
        Public paramsMin As Integer
        Public paramsMax As Integer
        Public Description As String
        Public EditorCaption As String
        Public params() As paramsType
        Public returnType As ReturnFunctionEnum
        Public returnArray() As String
        Public eventId As Integer
        ''' <summary>порядковый номер для отображения в редакторе</summary>
        Public editorIndex As Integer


        Public Function Clone(Optional ByVal duplicateEvents As Boolean = True) As PropertiesInfoType
            Dim other As PropertiesInfoType = DirectCast(Me.MemberwiseClone, PropertiesInfoType)
            If IsNothing(params) OrElse params.Count = 0 Then Return other
            If IsNothing(returnArray) = False Then other.returnArray = returnArray.Clone

            ReDim other.params(params.Count - 1)
            For i As Integer = 0 To params.Count - 1
                other.params(i) = params(i).Clone
            Next i

            If eventId > 0 AndAlso duplicateEvents Then other.eventId = mScript.eventRouter.DuplicateEvent(eventId)
            Return other
        End Function
    End Class

    'структура для хранения свойств объектов 2-го порядка
    Public Class ChildPropertiesInfoType
        Public Value As String 'значение свойства 2-го порядка
        'Public UserAdded As Boolean 'свойство встроенное или добавлено пользователем
        Public Hidden As Boolean
        Public eventId As Integer
        Public ThirdLevelProperties() As String 'массив для хранения значений свойств объектов 3-го порядка
        Public ThirdLevelEventId() As Integer 'массив для хранения id событий (расшифрованного кода), ассоциированного с каждым свойством

        Public Function Clone(Optional ByVal duplicateEvents As Boolean = True) As ChildPropertiesInfoType
            Dim newObj As ChildPropertiesInfoType = DirectCast(Me.MemberwiseClone, ChildPropertiesInfoType)
            If IsNothing(ThirdLevelEventId) = False Then newObj.ThirdLevelEventId = ThirdLevelEventId.Clone
            If IsNothing(ThirdLevelProperties) = False Then newObj.ThirdLevelProperties = ThirdLevelProperties.Clone

            If duplicateEvents Then
                If eventId > 0 Then
                    newObj.eventId = mScript.eventRouter.DuplicateEvent(eventId)
                End If
                If IsNothing(ThirdLevelEventId) = False Then
                    For i As Integer = 0 To ThirdLevelEventId.Count - 1
                        If ThirdLevelEventId(i) > 0 Then newObj.ThirdLevelEventId(i) = mScript.eventRouter.DuplicateEvent(ThirdLevelEventId(i))
                    Next
                End If
            End If

            Return newObj
        End Function
    End Class

    Public Class FunctionInfoType
        ''' <summary>Соодержимое функции если она создана в редакторе</summary>
        Public ValueDt() As CodeTextBox.CodeDataType
        ''' <summary>Соодержимое функции в виде исполняемого кода</summary>
        Public ValueExecuteDt As List(Of ExecuteDataType)
        ''' <summary>Список перменных, привязанных к функции. Ключ - имена переменных (без кавычек)</summary>
        Public Variables As SortedList(Of String, cVariable.variableEditorInfoType)
        ''' <summary>Скрыта функция или нет</summary>
        Public Hidden As Boolean
        ''' <summary>Иконка для отображения в дереве</summary>
        Public Icon As String
        ''' <summary>Группа для отображения в дереве</summary>
        Public Group As String
        ''' <summary>Описание функции</summary>
        Public Description As String

        ''' <summary>Запуск функции</summary>
        Public Function Run(ByRef arrParams() As String) As String
            If IsNothing(Variables) = False AndAlso Variables.Count > 0 Then mScript.csLocalVariables.RestoreVariables(Variables, True)
            Return mScript.ExecuteCode(ValueExecuteDt, arrParams, True)
        End Function

        Public Function Clone() As FunctionInfoType
            Dim other As FunctionInfoType = DirectCast(Me.MemberwiseClone, FunctionInfoType)
            If IsNothing(ValueDt) = False AndAlso ValueDt.Count > 0 Then
                other.ValueDt = CopyCodeDataArray(ValueDt)
            End If
            If IsNothing(ValueExecuteDt) = False AndAlso ValueExecuteDt.Count > 0 Then
                other.ValueExecuteDt = New List(Of ExecuteDataType)
                For i As Integer = 0 To ValueExecuteDt.Count - 1
                    other.ValueExecuteDt.Add(ValueExecuteDt(i).Clone)
                Next
            End If
            Return other
        End Function
    End Class

    ''' <summary>Структура для хранения всей информации об определенном классе (например Code, Location)</summary>
    Public Class MainClassType
        ''' <summary>массив с именами классов</summary>
        Public Names() As String
        ''' <summary>свойства (key = имя свойства, value = PropertiesInfoType)</summary>
        Public Properties As SortedList(Of String, PropertiesInfoType)
        ''' <summary>функции (key = имя функции, value = PropertiesInfoType)</summary>
        Public Functions As SortedList(Of String, PropertiesInfoType)
        ''' <summary>встроенный класс или класс пользователя</summary>
        Public UserAdded As Boolean
        '''<summary>массив для хранения свойств объектов второго и третьего порядков (key = имя свойства, value = ChildPropertiesInfoType,
        '''value.value - свойства объектов 2-го порядка,value.ThirdLevelProperties() - свойства объектов 3-го порядка)</summary>
        Public ChildProperties() As SortedList(Of String, ChildPropertiesInfoType)
        ''' <summary>Количество уровней в классе (от 0 до 2)</summary>
        Public LevelsCount As Byte
        ''' <summary>Файл помощи по классу</summary>
        Public HelpFile As String
        ''' <summary>Свойство для классов 3 уровня (и для действий), которое будет отображаться при выборе узла в поддереве</summary>
        Public DefaultProperty As String

        Public Function Clone(Optional ByVal copyWithChildren As Boolean = True, Optional ByVal duplicateEvents As Boolean = True) As MainClassType
            Dim other As MainClassType = DirectCast(Me.MemberwiseClone, MainClassType)
            If IsNothing(Properties) = False AndAlso Properties.Count > 0 Then
                other.Properties = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
                For i As Integer = 0 To Properties.Count - 1
                    other.Properties.Add(Properties.ElementAt(i).Key, Properties.ElementAt(i).Value.Clone(duplicateEvents))
                Next
            End If

            If IsNothing(Functions) = False AndAlso Functions.Count > 0 Then
                other.Functions = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
                For i As Integer = 0 To Functions.Count - 1
                    other.Functions.Add(Functions.ElementAt(i).Key, Functions.ElementAt(i).Value.Clone(duplicateEvents))
                Next
            End If

            If copyWithChildren = False OrElse IsNothing(ChildProperties) OrElse ChildProperties.Count = 0 Then
                other.ChildProperties = Nothing
                Return other
            End If

            ReDim other.ChildProperties(ChildProperties.Count - 1)
            For i As Integer = 0 To ChildProperties.Count - 1
                other.ChildProperties(i) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
                If IsNothing(ChildProperties(i)) Then Continue For
                For j As Integer = 0 To ChildProperties(i).Count - 1
                    other.ChildProperties(i).Add(ChildProperties(i).ElementAt(j).Key, ChildProperties(i).ElementAt(j).Value.Clone(duplicateEvents))
                Next j
            Next i

            Return other
        End Function
    End Class

    ''' <summary>Массив труктур для хранения всей информации о классах (Code, Location...)</summary>
    Public mainClass() As MainClassType 'главная структура
    ''' <summary>Копия свойств массива структур для хранения всей информации о классах (Code, Location...) для восстановления значений по умолчанию</summary>
    Public mainClassCopy() As MainClassType
    ''' <summary>хэш главной структуры, для ускоренного обращения по имени классов</summary>
    Public mainClassHash As New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
    Public functionsHash As New SortedList(Of String, FunctionInfoType) 'хэш с функциями пользователя, в которых содержится готовый код на исполнение
    Public basicFunctionsHashLevel1 As New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase) 'базовые функции для класса пользователя (AddPropery, RemoveProperty, AddFunction, RemoveFunction, SetProperty)
    Public basicFunctionsHashLevel2 As New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase) 'базовые функции для класса пользователя (Create, Remove, Id, Count, IsExists, AddPropery, RemoveProperty, AddFunction, RemoveFunction, SetProperty)
    Public basicFunctionsHashLevel3 As New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase) 'базовые функции для класса пользователя (Create, Remove, Id, Count, IsExists, AddPropery, RemoveProperty, AddFunction, RemoveFunction, SetProperty)
    ''' <summary>хэш всех функций и свойств всех классов. Ключ выглядит как [первое имя класса]_[имя функции/свойства] - напр. C_Min. См. FillFuncAndPropHash</summary>
    Public funcAndPropHash As New SortedList(Of String, MatewScript.funcAndPropHashType)(StringComparer.CurrentCultureIgnoreCase)
    ''' <summary>Класс битвы</summary>
    Public Battle As New cBattle
#End Region

#Region "Classes"
    ''' <summary>
    ''' Процедура загружает всю информацию о классах и их свойствах и функциях, включая данные для редактора, из classes.xml
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LoadClasses(Optional ByVal fileName As String = "")
        If String.IsNullOrEmpty(fileName) Then fileName = ProgramPath + "\src\classes.xml"
        Dim classesXML As New System.Xml.XmlTextReader(fileName) 'главный xml
        Dim funcXML As System.Xml.XmlReader 'xml для функций и свойств
        Dim paramXML As System.Xml.XmlReader 'xml для параметров функций
        Dim enumXML As System.Xml.XmlReader 'xml для перечисления возможных значений параметров функций и для возможных значений того, что функция или свойство возвращают
        Dim elementName As String 'для хранения имени текущего свойства или функции

        Dim classUBound As Integer = -1 'верхняя граница массива mainClass (чтобы не проверять каждый раз)
        Erase mainClass 'на всякий случай очищаем нашу главную структуру
        While classesXML.ReadToFollowing("CLASS") 'пееходим от класса к классу
            'расширяем массив для добавления инфо о классе
            classUBound += 1
            ReDim Preserve mainClass(classUBound)
            mainClass(classUBound) = New MainClassType
            mainClass(classUBound).LevelsCount = CByte(classesXML.Item("levels")) 'кол-во уровней класса
            mainClass(classUBound).HelpFile = classesXML.Item("helpFile") 'кол-во уровней класса
            mainClass(classUBound).DefaultProperty = classesXML.Item("defaultProperty") 'кол-во уровней класса
            mainClass(classUBound).Names = classesXML.Item("names").Split(","c) 'его имена
            mainClass(classUBound).UserAdded = CBool(classesXML.Item("UserAdded")) 'кол-во уровней класса

            'создаем таблицу для функций
            Dim editorId As Integer = -1
            Dim arrParams() As paramsType = Nothing, arrParamsUBound As Integer = -1 'массив для временного хранения полученных параметров
            mainClass(classUBound).Functions = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
            If classesXML.ReadToFollowing("FUNCTIONS") Then
                funcXML = classesXML.ReadSubtree 'получаем дерево функций данного класса
                While funcXML.ReadToFollowing("function") 'переходим от функции к функции
                    Dim prop As New PropertiesInfoType 'создаем переменную для получения всей инфо о функции
                    prop.helpFile = funcXML.Item("help_file")
                    prop.paramsMax = funcXML.Item("params_max")
                    prop.paramsMin = funcXML.Item("params_min")
                    prop.Description = funcXML.Item("description")
                    editorId += 1
                    prop.editorIndex = editorId
                    elementName = funcXML.Item("name") 'сохраняем имя функции

                    paramXML = funcXML.ReadSubtree 'получаем дерево параметров данной функции
                    Erase arrParams
                    arrParamsUBound = -1 'массив для временного хранения полученных параметров
                    While paramXML.Read 'переходим от параметра к параметру
                        If paramXML.NodeType = System.Xml.XmlNodeType.Element Then
                            If paramXML.Name = "param" Then
                                arrParamsUBound += 1 'расширяем массив параметров и получаем в него всю инфо
                                ReDim Preserve arrParams(arrParamsUBound)
                                Dim par As New paramsType With {.Description = paramXML.Item("description"), .Name = paramXML.Item("name"), _
                                                                               .Type = paramXML.Item("type")}
                                If par.Type = paramsType.paramsTypeEnum.PARAM_ENUM OrElse par.Type = paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                                    'Если тип = PARAM_ENUM, то в этом параметре должно быть перечисление возможных значений
                                    enumXML = paramXML.ReadSubtree 'получаем дерево значений параметра
                                    While enumXML.ReadToFollowing("enum") 'переходим от значения к значению
                                        'расширяем массив значений и получаем его из xml
                                        If IsNothing(par.EnumValues) Then
                                            ReDim par.EnumValues(0)
                                        Else
                                            ReDim Preserve par.EnumValues(par.EnumValues.Length)
                                        End If
                                        par.EnumValues(par.EnumValues.GetUpperBound(0)) = enumXML.Item("value")
                                    End While
                                End If
                                arrParams(arrParamsUBound) = par
                            ElseIf paramXML.Name = "return" Then
                                'если есть return - получаем и его
                                prop.returnType = paramXML.Item("type") 'сохраняем тип возвращаемого значения
                                If prop.returnType = ReturnFunctionEnum.RETURN_ENUM OrElse prop.returnType = ReturnFunctionEnum.RETURN_ELEMENT Then
                                    'если он = RETURN_ENUM, далее идет перечисление его возможных значений
                                    enumXML = paramXML.ReadSubtree 'получаем дерево значений 
                                    While enumXML.ReadToFollowing("enum") 'переходим от значения к значению
                                        'расширяем массив значений и получаем его из xml
                                        If IsNothing(prop.returnArray) Then
                                            ReDim prop.returnArray(0)
                                        Else
                                            ReDim Preserve prop.returnArray(prop.returnArray.Length)
                                        End If
                                        prop.returnArray(prop.returnArray.GetUpperBound(0)) = enumXML.Item("value")
                                    End While
                                End If
                            End If
                        End If
                    End While
                    If IsNothing(arrParams) = False Then prop.params = arrParams 'сохраняем параметры во временной переменной с инфо для функции
                    'создаем новую функцию со всей полученной информацией
                    mainClass(classUBound).Functions.Add(elementName, prop)
                End While
            End If

            'создаем таблицу для свойств
            editorId = -1
            mainClass(classUBound).Properties = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
            If classesXML.ReadToFollowing("PROPERTIES") Then
                funcXML = classesXML.ReadSubtree 'получаем дерево свойств данного класса
                While funcXML.ReadToFollowing("property") 'переходи от свойства к свойству
                    Dim prop As New PropertiesInfoType 'создаем переменную для получения всей инфо о свойстве
                    prop.helpFile = funcXML.Item("help_file")
                    prop.Value = funcXML.Item("def_value")
                    prop.Description = funcXML.Item("description")
                    prop.EditorCaption = funcXML.Item("editor_name")
                    editorId += 1
                    prop.editorIndex = editorId
                    prop.Hidden = [Enum].Parse(GetType(MatewScript.PropertyHiddenEnum), funcXML.Item("hidden"))
                    elementName = funcXML.Item("name") 'сохраняем имя свойства

                    paramXML = funcXML.ReadSubtree 'получаем дерево свойства
                    Erase arrParams
                    arrParamsUBound = -1 'массив для временного хранения полученных параметров
                    Dim blnReturnUsual As Boolean = True
                    While paramXML.Read 'переходим от параметра к параметру
                        If paramXML.NodeType = System.Xml.XmlNodeType.Element Then
                            If paramXML.Name = "param" Then
                                arrParamsUBound += 1 'расширяем массив параметров и получаем в него всю инфо
                                ReDim Preserve arrParams(arrParamsUBound)
                                Dim par As New paramsType With {.Description = paramXML.Item("description"), .Name = paramXML.Item("name"), _
                                                                               .Type = paramXML.Item("type")}
                                If par.Type = paramsType.paramsTypeEnum.PARAM_ENUM OrElse par.Type = paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                                    'Если тип = PARAM_ENUM, то в этом параметре должно быть перечисление возможных значений
                                    enumXML = paramXML.ReadSubtree 'получаем дерево значений параметра
                                    While enumXML.ReadToFollowing("enum") 'переходим от значения к значению
                                        'расширяем массив значений и получаем его из xml
                                        If IsNothing(par.EnumValues) Then
                                            ReDim par.EnumValues(0)
                                        Else
                                            ReDim Preserve par.EnumValues(par.EnumValues.Length)
                                        End If
                                        par.EnumValues(par.EnumValues.GetUpperBound(0)) = enumXML.Item("value")
                                    End While
                                End If
                                arrParams(arrParamsUBound) = par
                            ElseIf paramXML.Name = "return" Then
                                blnReturnUsual = False
                                'если есть return - получаем и его
                                prop.returnType = paramXML.Item("type") 'сохраняем тип возвращаемого значения
                                If prop.returnType = ReturnFunctionEnum.RETURN_ENUM OrElse prop.returnType = ReturnFunctionEnum.RETURN_ELEMENT Then
                                    'если он = RETURN_ENUM, далее идет перечисление его возможных значений
                                    enumXML = paramXML.ReadSubtree 'получаем дерево значений 
                                    While enumXML.ReadToFollowing("enum") 'переходим от значения к значению
                                        'расширяем массив значений и получаем его из xml
                                        If IsNothing(prop.returnArray) Then
                                            ReDim prop.returnArray(0)
                                        Else
                                            ReDim Preserve prop.returnArray(prop.returnArray.Length)
                                        End If
                                        prop.returnArray(prop.returnArray.GetUpperBound(0)) = enumXML.Item("value")
                                    End While
                                End If
                            End If
                        End If
                    End While
                    If blnReturnUsual Then prop.returnType = ReturnFunctionEnum.RETURN_USUAL
                    If IsNothing(arrParams) = False Then prop.params = arrParams 'сохраняем параметры во временной переменной с инфо для свойства
                    'создаем новое свойство со всей полученной информацией
                    mainClass(classUBound).Properties.Add(elementName, prop)
                    ''устанавливаем событие для исполняемых строк и кода
                    Dim cRes As MatewScript.ContainsCodeEnum = IsPropertyContainsCode(prop.Value)
                    If cRes <> MatewScript.ContainsCodeEnum.NOT_CODE Then mScript.eventRouter.SetEventId(classUBound, elementName, prop.Value, cRes, -1, -1)
                End While
            End If

        End While
        classesXML.Close()
        MakeMainClassHash() 'создания хэша для быстрого поиска id класса по любому его имени
        FillBasicFunctionsHash() 'заполняем структуру с базовыми (обязательными) функциями и свойствами для классов 2 и 3 уровней. Например, функции Create, Remove, свойство Name и т. д.
    End Sub

    Public Sub CreateMainClassCopy()
        Dim classesXML As New System.Xml.XmlTextReader(ProgramPath + "\src\classes.xml") 'главный xml
        Dim funcXML As System.Xml.XmlReader 'xml для функций и свойств
        Dim paramXML As System.Xml.XmlReader 'xml для параметров функций
        Dim enumXML As System.Xml.XmlReader 'xml для перечисления возможных значений параметров функций и для возможных значений того, что функция или свойство возвращают
        Dim elementName As String 'для хранения имени текущего свойства или функции

        Dim classUBound As Integer = -1 'верхняя граница массива mainClassCopy (чтобы не проверять каждый раз)
        Erase mainClassCopy 'на всякий случай очищаем нашу главную структуру
        While classesXML.ReadToFollowing("CLASS") 'пееходим от класса к классу
            'расширяем массив для добавления инфо о классе
            classUBound += 1
            ReDim Preserve mainClassCopy(classUBound)
            mainClassCopy(classUBound) = New MainClassType
            mainClassCopy(classUBound).LevelsCount = CByte(classesXML.Item("levels")) 'кол-во уровней класса
            mainClassCopy(classUBound).HelpFile = classesXML.Item("helpFile") 'кол-во уровней класса
            mainClassCopy(classUBound).DefaultProperty = classesXML.Item("defaultProperty") 'кол-во уровней класса
            mainClassCopy(classUBound).Names = classesXML.Item("names").Split(","c) 'его имена
            mainClassCopy(classUBound).UserAdded = CBool(classesXML.Item("UserAdded"))

            'создаем таблицу для функций
            Dim editorId As Integer = -1
            Dim arrParams() As paramsType = Nothing, arrParamsUBound As Integer = -1 'массив для временного хранения полученных параметров
            mainClassCopy(classUBound).Functions = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
            If classesXML.ReadToFollowing("FUNCTIONS") Then
                funcXML = classesXML.ReadSubtree 'получаем дерево функций данного класса
                While funcXML.ReadToFollowing("function") 'переходим от функции к функции
                    Dim prop As New PropertiesInfoType 'создаем переменную для получения всей инфо о функции
                    prop.helpFile = funcXML.Item("help_file")
                    prop.paramsMax = funcXML.Item("params_max")
                    prop.paramsMin = funcXML.Item("params_min")
                    prop.Description = funcXML.Item("description")
                    editorId += 1
                    prop.editorIndex = editorId
                    elementName = funcXML.Item("name") 'сохраняем имя функции

                    paramXML = funcXML.ReadSubtree 'получаем дерево параметров данной функции
                    Erase arrParams
                    arrParamsUBound = -1 'массив для временного хранения полученных параметров
                    While paramXML.Read 'переходим от параметра к параметру
                        If paramXML.NodeType = System.Xml.XmlNodeType.Element Then
                            If paramXML.Name = "param" Then
                                arrParamsUBound += 1 'расширяем массив параметров и получаем в него всю инфо
                                ReDim Preserve arrParams(arrParamsUBound)
                                Dim par As New paramsType With {.Description = paramXML.Item("description"), .Name = paramXML.Item("name"), _
                                                                               .Type = paramXML.Item("type")}
                                If par.Type = paramsType.paramsTypeEnum.PARAM_ENUM OrElse par.Type = paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                                    'Если тип = PARAM_ENUM, то в этом параметре должно быть перечисление возможных значений
                                    enumXML = paramXML.ReadSubtree 'получаем дерево значений параметра
                                    While enumXML.ReadToFollowing("enum") 'переходим от значения к значению
                                        'расширяем массив значений и получаем его из xml
                                        If IsNothing(par.EnumValues) Then
                                            ReDim par.EnumValues(0)
                                        Else
                                            ReDim Preserve par.EnumValues(par.EnumValues.Length)
                                        End If
                                        par.EnumValues(par.EnumValues.GetUpperBound(0)) = enumXML.Item("value")
                                    End While
                                End If
                                arrParams(arrParamsUBound) = par
                            ElseIf paramXML.Name = "return" Then
                                'если есть return - получаем и его
                                prop.returnType = paramXML.Item("type") 'сохраняем тип возвращаемого значения
                                If prop.returnType = ReturnFunctionEnum.RETURN_ENUM OrElse prop.returnType = ReturnFunctionEnum.RETURN_ELEMENT Then
                                    'если он = RETURN_ENUM, далее идет перечисление его возможных значений
                                    enumXML = paramXML.ReadSubtree 'получаем дерево значений 
                                    While enumXML.ReadToFollowing("enum") 'переходим от значения к значению
                                        'расширяем массив значений и получаем его из xml
                                        If IsNothing(prop.returnArray) Then
                                            ReDim prop.returnArray(0)
                                        Else
                                            ReDim Preserve prop.returnArray(prop.returnArray.Length)
                                        End If
                                        prop.returnArray(prop.returnArray.GetUpperBound(0)) = enumXML.Item("value")
                                    End While
                                End If
                            End If
                        End If
                    End While
                    If IsNothing(arrParams) = False Then prop.params = arrParams 'сохраняем параметры во временной переменной с инфо для функции
                    'создаем новую функцию со всей полученной информацией
                    mainClassCopy(classUBound).Functions.Add(elementName, prop)
                    'If mainClassCopy(classUBound).Functions.ContainsKey(elementName) = False Then mainClassCopy(classUBound).Functions.Add(elementName, prop)
                End While
            End If

            'создаем таблицу для свойств
            mainClassCopy(classUBound).Properties = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
            If classesXML.ReadToFollowing("PROPERTIES") Then
                editorId = -1
                funcXML = classesXML.ReadSubtree 'получаем дерево свойств данного класса
                While funcXML.ReadToFollowing("property") 'переходи от свойства к свойству
                    Dim prop As New PropertiesInfoType 'создаем переменную для получения всей инфо о свойстве
                    prop.helpFile = funcXML.Item("help_file")
                    prop.Value = funcXML.Item("def_value")
                    prop.Description = funcXML.Item("description")
                    prop.EditorCaption = funcXML.Item("editor_name")
                    editorId += 1
                    prop.editorIndex = editorId
                    prop.Hidden = [Enum].Parse(GetType(MatewScript.PropertyHiddenEnum), funcXML.Item("hidden"))
                    elementName = funcXML.Item("name") 'сохраняем имя свойства

                    Erase arrParams
                    arrParamsUBound = -1 'массив для временного хранения полученных параметров
                    Dim blnReturnUsual As Boolean = True
                    paramXML = funcXML.ReadSubtree 'получаем дерево свойства, где должен быть узел return
                    While paramXML.Read 'переходим от параметра к параметру
                        If paramXML.NodeType = System.Xml.XmlNodeType.Element Then
                            If paramXML.Name = "param" Then
                                arrParamsUBound += 1 'расширяем массив параметров и получаем в него всю инфо
                                ReDim Preserve arrParams(arrParamsUBound)
                                Dim par As New paramsType With {.Description = paramXML.Item("description"), .Name = paramXML.Item("name"), _
                                                                               .Type = paramXML.Item("type")}
                                If par.Type = paramsType.paramsTypeEnum.PARAM_ENUM OrElse par.Type = paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                                    'Если тип = PARAM_ENUM, то в этом параметре должно быть перечисление возможных значений
                                    enumXML = paramXML.ReadSubtree 'получаем дерево значений параметра
                                    While enumXML.ReadToFollowing("enum") 'переходим от значения к значению
                                        'расширяем массив значений и получаем его из xml
                                        If IsNothing(par.EnumValues) Then
                                            ReDim par.EnumValues(0)
                                        Else
                                            ReDim Preserve par.EnumValues(par.EnumValues.Length)
                                        End If
                                        par.EnumValues(par.EnumValues.GetUpperBound(0)) = enumXML.Item("value")
                                    End While
                                End If
                                arrParams(arrParamsUBound) = par
                            ElseIf paramXML.Name = "return" Then
                                blnReturnUsual = True
                                'если есть return - получаем и его
                                prop.returnType = paramXML.Item("type") 'сохраняем тип возвращаемого значения
                                If prop.returnType = ReturnFunctionEnum.RETURN_ENUM OrElse prop.returnType = ReturnFunctionEnum.RETURN_ELEMENT Then
                                    'если он = RETURN_ENUM, далее идет перечисление его возможных значений
                                    enumXML = paramXML.ReadSubtree 'получаем дерево значений 
                                    While enumXML.ReadToFollowing("enum") 'переходим от значения к значению
                                        'расширяем массив значений и получаем его из xml
                                        If IsNothing(prop.returnArray) Then
                                            ReDim prop.returnArray(0)
                                        Else
                                            ReDim Preserve prop.returnArray(prop.returnArray.Length)
                                        End If
                                        prop.returnArray(prop.returnArray.GetUpperBound(0)) = enumXML.Item("value")
                                    End While
                                End If
                            End If
                        End If
                    End While
                    If blnReturnUsual Then prop.returnType = ReturnFunctionEnum.RETURN_USUAL
                    If IsNothing(arrParams) = False Then prop.params = arrParams 'сохраняем параметры во временной переменной с инфо для свойства
                    'на копии события не делаются. Копируем из основного 
                    prop.eventId = mScript.mainClass(classUBound).Properties(elementName).eventId
                    'Dim cRes As MatewScript.ContainsCodeEnum = IsPropertyContainsCode(prop.Value)
                    'If cRes <> MatewScript.ContainsCodeEnum.NOT_CODE Then
                    '    'mScript.eventRouter.SetEventId(classUBound, elementName, prop.Value, cRes, -1, -1)
                    'End If
                    'создаем новое свойство со всей полученной информацией
                    mainClassCopy(classUBound).Properties.Add(elementName, prop)
                End While
            End If

        End While
        classesXML.Close()
    End Sub

    ''' <summary>
    ''' Создает список, содержащий порядковые номера элементов PropertiesInfoType, отсортированные по editorIndex
    ''' </summary>
    ''' <param name="propList">Ссылка на массив PropertiesInfoType, который надо отсортировать</param>
    ''' <returns>отсортированный массив</returns>
    Public Function CreatePropertiesOrderArray(ByRef propList As SortedList(Of String, PropertiesInfoType)) As Integer()
        Dim arr() As Integer = Nothing
        If IsNothing(propList) OrElse propList.Count = 0 Then Return {}

        ReDim arr(propList.Count - 1)
        For i As Integer = 0 To propList.Count - 1
            Dim freeIndex As Integer = propList.ElementAt(i).Value.editorIndex
            If freeIndex > arr.Count - 1 Then
                'editorIndex за пределами массива. Ищем свободный индекс
                freeIndex = -1
                For j As Integer = 0 To propList.Count - 1
                    Dim blnFound As Boolean = False
                    For u As Integer = 0 To propList.Count - 1
                        If propList.ElementAt(u).Value.editorIndex = j Then
                            blnFound = True
                            Exit For
                        End If
                    Next u
                    If Not blnFound Then
                        freeIndex = j
                        Exit For
                    End If
                Next j
                If freeIndex = -1 Then freeIndex = propList.Count - 1
            End If
            arr(freeIndex) = i
        Next
        Return arr
    End Function

    ''' <summary>
    ''' Сохраняет в xml данные о классах структуры mainClass
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SaveClasses(Optional ByVal fileName As String = "")
        'открываем наш xml-файл на запись
        If String.IsNullOrEmpty(fileName) Then fileName = ProgramPath + "\src\classes.xml"
        Dim classesXML As System.Xml.XmlWriter = System.Xml.XmlWriter.Create(fileName)

        classesXML.WriteStartDocument() 'пишем заголовок документа
        'пишем всю структуру
        classesXML.WriteStartElement("ENGINE")
        For i = 0 To mainClass.GetUpperBound(0)
            'перебираем все классы
            classesXML.WriteStartElement("CLASS")
            classesXML.WriteAttributeString("levels", mainClass(i).LevelsCount.ToString)
            classesXML.WriteAttributeString("names", Join(mainClass(i).Names, ","))
            classesXML.WriteAttributeString("helpFile", mainClass(i).HelpFile)
            classesXML.WriteAttributeString("defaultProperty", mainClass(i).DefaultProperty)
            classesXML.WriteAttributeString("UserAdded", mainClass(i).UserAdded.ToString)

            classesXML.WriteStartElement("FUNCTIONS")
            If IsNothing(mainClass(i).Functions) = False Then
                'пишем функции класса
                Dim arrList = CreatePropertiesOrderArray(mainClass(i).Functions)
                For j As Integer = 0 To arrList.Count - 1
                    Dim func As KeyValuePair(Of String, PropertiesInfoType) = mainClass(i).Functions.ElementAt(arrList(j))
                    classesXML.WriteStartElement("function")
                    classesXML.WriteAttributeString("name", func.Key)
                    classesXML.WriteAttributeString("help_file", func.Value.helpFile)
                    classesXML.WriteAttributeString("description", func.Value.Description)
                    classesXML.WriteAttributeString("params_min", func.Value.paramsMin)
                    classesXML.WriteAttributeString("params_max", func.Value.paramsMax)

                    If IsNothing(func.Value.params) = False Then
                        'параметры функции
                        For Each par As paramsType In func.Value.params
                            classesXML.WriteStartElement("param")
                            classesXML.WriteAttributeString("name", par.Name)
                            classesXML.WriteAttributeString("description", par.Description)
                            classesXML.WriteAttributeString("type", Convert.ToString(par.Type))
                            If IsNothing(par.EnumValues) = False Then
                                'варианты значения параметра функции
                                For Each parEnum As String In par.EnumValues
                                    classesXML.WriteStartElement("enum")
                                    classesXML.WriteAttributeString("value", parEnum)
                                    classesXML.WriteEndElement() 'enum
                                Next
                            End If
                            classesXML.WriteEndElement() 'param
                        Next
                    End If

                    If func.Value.returnType <> ReturnFunctionEnum.RETURN_USUAL Then
                        'что функция возвращает
                        classesXML.WriteStartElement("return")
                        classesXML.WriteAttributeString("type", Convert.ToString(func.Value.returnType))
                        If IsNothing(func.Value.returnArray) = False Then
                            For Each retEnum As String In func.Value.returnArray
                                'пишем варианты того, что функция возвращает
                                classesXML.WriteStartElement("enum")
                                classesXML.WriteAttributeString("value", retEnum)
                                classesXML.WriteEndElement() 'enum
                            Next
                        End If
                        classesXML.WriteEndElement() 'return
                    End If
                    classesXML.WriteEndElement() 'function
                Next j
            End If
            classesXML.WriteEndElement() 'FUNCTIONS

            'пишем свойства класса
            classesXML.WriteStartElement("PROPERTIES")
            If IsNothing(mainClass(i).Properties) = False Then
                Dim arrList = CreatePropertiesOrderArray(mainClass(i).Properties)
                For j As Integer = 0 To arrList.Count - 1
                    Dim prop As KeyValuePair(Of String, PropertiesInfoType) = mainClass(i).Properties.ElementAt(arrList(j))
                    classesXML.WriteStartElement("property")
                    classesXML.WriteAttributeString("name", prop.Key)
                    classesXML.WriteAttributeString("def_value", prop.Value.Value)
                    classesXML.WriteAttributeString("help_file", prop.Value.helpFile)
                    classesXML.WriteAttributeString("description", prop.Value.Description)
                    classesXML.WriteAttributeString("editor_name", prop.Value.EditorCaption)
                    classesXML.WriteAttributeString("hidden", prop.Value.Hidden.ToString)

                    If IsNothing(prop.Value.params) = False Then
                        'параметры свойства
                        For Each par As paramsType In prop.Value.params
                            classesXML.WriteStartElement("param")
                            classesXML.WriteAttributeString("name", par.Name)
                            classesXML.WriteAttributeString("description", par.Description)
                            classesXML.WriteAttributeString("type", Convert.ToString(par.Type))
                            classesXML.WriteEndElement() 'param
                        Next
                    End If

                    If prop.Value.returnType <> ReturnFunctionEnum.RETURN_USUAL Then
                        'что свойство возвращает
                        classesXML.WriteStartElement("return")
                        classesXML.WriteAttributeString("type", Convert.ToString(prop.Value.returnType))
                        If IsNothing(prop.Value.returnArray) = False Then
                            'варианты того, что свойство возвращает
                            For Each retEnum As String In prop.Value.returnArray
                                classesXML.WriteStartElement("enum")
                                classesXML.WriteAttributeString("value", retEnum)
                                classesXML.WriteEndElement() 'enum
                            Next
                        End If
                        classesXML.WriteEndElement() 'return
                    End If
                    classesXML.WriteEndElement() 'property
                Next
            End If
            classesXML.WriteEndElement() 'PROPERTIES
            classesXML.WriteEndElement() 'CLASS
        Next
        classesXML.WriteEndElement() 'ENGINE
        classesXML.Close()
    End Sub

    ''' <summary>Заполняет структуру с базовыми (обязательными) функциями и свойствами для классов 2 и 3 уровней.
    ''' Например, функции Create, Remove, свойство Name и т. д.
    ''' Надо для создания классов пользователя и редактора классов
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FillBasicFunctionsHash()
        Dim par() As paramsType 'для временного хранения свойств функции или свойства
        'структура для классов 1 уровня
        basicFunctionsHashLevel1.Clear()

        'структура для классов 2 уровня
        basicFunctionsHashLevel2.Clear()
        ReDim par(0)
        par(0) = New paramsType
        par(0).Name = "Имя"
        par(0).Description = "имя элемента, который будет создан"
        par(0).Type = paramsType.paramsTypeEnum.PARAM_STRING
        basicFunctionsHashLevel2.Add("Create", New PropertiesInfoType With {.Description = "Создает новый элемент.", .paramsMin = 1, .paramsMax = 1, .params = par, .editorIndex = 0})

        ReDim par(0)
        par(0) = New paramsType
        par(0).Name = "Имя/Id"
        par(0).Description = "имя/id удаляемого элемента"
        par(0).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT
        par(0).EnumValues = {"[Current]"}
        basicFunctionsHashLevel2.Add("Remove", New PropertiesInfoType With {.Description = "Удаляет элемент(ы).", .paramsMin = 0, .paramsMax = 1, .params = par, .editorIndex = 1})

        ReDim par(0)
        par(0) = New paramsType
        par(0).Name = "Имя"
        par(0).Description = "имя элемента, id которого надо получить"
        par(0).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT
        par(0).EnumValues = {"[Current]"}
        basicFunctionsHashLevel2.Add("Id", New PropertiesInfoType With {.Description = "Возвращает id элемента по его имени.", .paramsMin = 1, .paramsMax = 1, .params = par, .editorIndex = 2})

        ReDim par(0)
        par(0) = New paramsType
        par(0).Name = "Имя/Id"
        par(0).Description = "имя/id проверяемого элемента"
        par(0).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT
        par(0).EnumValues = {"[Current]"}
        basicFunctionsHashLevel2.Add("IsExist", New PropertiesInfoType With {.Description = "Определяет существует ли элемент с данным именем / id (или же существует ли хоть один элемент).", _
                                                                             .paramsMin = 0, .paramsMax = 1, .params = par, .returnType = ReturnFunctionEnum.RETURN_BOOl, .editorIndex = 3})

        basicFunctionsHashLevel2.Add("Count", New PropertiesInfoType With {.Description = "Определяет количество элементов.", .paramsMin = 0, .editorIndex = 4})

        'структура для классов 3 уровня
        basicFunctionsHashLevel3.Clear()
        ReDim par(1)
        par(0) = New paramsType
        par(1) = New paramsType
        par(0).Name = "Имя1"
        par(0).Description = "имя/id создаваемого элемента или имя родителя, внутри которого будет создан новый элемент."
        par(0).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT
        par(0).EnumValues = {"[Current]"}
        par(1).Name = "Имя2"
        par(1).Description = "имя элемента, который будет создан внутри элемента-родителя [Имя1]"
        par(1).Type = paramsType.paramsTypeEnum.PARAM_STRING
        basicFunctionsHashLevel3.Add("Create", New PropertiesInfoType With {.Description = "Создает новый элемент.", .paramsMin = 1, .paramsMax = 2, .params = par, .editorIndex = 0})

        ReDim par(1)
        par(0) = New paramsType
        par(1) = New paramsType
        par(0).Name = "Имя1"
        par(0).Description = "имя/id удаляемого элемента или имя родителя, из которого будет удален элемент."
        par(0).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT
        par(0).EnumValues = {"[Current]"}
        par(1).Name = "Имя2"
        par(1).Description = "имя/id элемента, который будет удален из элемента-родителя [Имя1]"
        par(1).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT2
        basicFunctionsHashLevel3.Add("Remove", New PropertiesInfoType With {.Description = "Удаляет элемент(ы).", .paramsMin = 0, .paramsMax = 2, .params = par, .editorIndex = 1})

        ReDim par(1)
        par(0) = New paramsType
        par(1) = New paramsType
        par(0).Name = "Имя1"
        par(0).Description = "имя элемента или имя/id родителя, id элемента которого надо получить"
        par(0).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT
        par(0).EnumValues = {"[Current]"}
        par(1).Name = "Имя2"
        par(1).Description = "имя дочернего элемента внутри элемента-родителя [Имя1], Id которого надо получить."
        par(1).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT2
        basicFunctionsHashLevel3.Add("Id", New PropertiesInfoType With {.Description = "Возвращает id элемента по его имени.", .paramsMin = 1, .paramsMax = 2, .params = par, .editorIndex = 2})

        ReDim par(1)
        par(0) = New paramsType
        par(1) = New paramsType
        par(0).Name = "Имя1"
        par(0).Description = "имя/id проверяемого элемента или имя/id родителя"
        par(0).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT
        par(0).EnumValues = {"[Current]"}
        par(1).Name = "Имя2"
        par(1).Description = "имя/id дочернего элемента внутри элемента-родителя [Имя1], наличие которого определяется."
        par(1).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT2
        basicFunctionsHashLevel3.Add("IsExist", New PropertiesInfoType With {.Description = "Определяет существует ли элемент с данным именем / id (или же существует ли хоть один элемент).", _
                                                                             .paramsMin = 0, .paramsMax = 2, .params = par, .returnType = ReturnFunctionEnum.RETURN_BOOl, .editorIndex = 3})

        ReDim par(0)
        par(0) = New paramsType
        par(0).Name = "Имя_родителя"
        par(0).Description = "имя/id родителя, количество дочерних элементов которого определяем."
        par(0).Type = paramsType.paramsTypeEnum.PARAM_ELEMENT
        par(0).EnumValues = {"[Current]"}
        basicFunctionsHashLevel3.Add("Count", New PropertiesInfoType With {.Description = "Определяет количество элементов.", .paramsMin = 0, .paramsMax = 1, .params = par, .editorIndex = 4})
    End Sub

    ''' <summary>
    ''' Изменяет в параметрах функций типа Element имя класса
    ''' </summary>
    ''' <param name="classId">Класс, в котором производить замены</param>
    ''' <param name="oldClassName">Старое имя класса</param>
    Public Sub UpdateBasicFunctionsParamsWhichIsElements(ByVal classId As Integer, Optional ByVal oldClassName As String = "[Current]")
        If IsNothing(mScript.mainClass(classId).Functions) Then Return
        Dim newClassName As String = mScript.mainClass(classId).Names.Last
        For fId As Integer = 0 To mainClass(classId).Functions.Count - 1
            Dim f As PropertiesInfoType = mainClass(classId).Functions.ElementAt(fId).Value
            If IsNothing(f.params) OrElse f.params.Count = 0 Then Continue For
            For parId As Integer = 0 To f.params.Count - 1
                Dim p As paramsType = f.params(parId)
                If p.Type = paramsType.paramsTypeEnum.PARAM_ELEMENT AndAlso IsNothing(p.EnumValues) = False AndAlso p.EnumValues.Count = 1 AndAlso String.Compare(p.EnumValues(0), oldClassName, True) = 0 Then
                    f.params(parId).EnumValues(0) = newClassName
                End If
            Next parId
        Next fId
    End Sub

    ''' <summary>
    ''' Процедура создает хэш-таблицу, ключом которой являются все имена всех элементов структуры mainClass, а значения - индексы в структуре, соответствующие этим именам
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub MakeMainClassHash()
        On Error GoTo er

        mainClassHash.Clear()
        For i As Integer = 0 To mainClass.Length - 1
            For j As Integer = 0 To mainClass(i).Names.Length - 1
                mainClassHash.Add(mainClass(i).Names(j), i)
            Next
        Next

        Exit Sub
er:
        MessageBox.Show("Ошибка при создании хэша mainClass. Продолжение работы программы невзоможно!", "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        System.Windows.Forms.Application.Exit()
    End Sub

    ''' <summary>
    ''' Функция обрабатывает блок New Class и создает новый класс (класс пользователя) в структуре mainClass.
    ''' </summary>
    ''' <param name="masCode">массив со всем выполняемым кодом</param>
    ''' <param name="className">первое имя класса</param>
    ''' <param name="startPos">положение строки New Class  в массиве masCode()</param>
    ''' <param name="finalPos">индекс последней строки кода внутри блока, которую надо обработать</param>
    ''' <param name="arrParams">параметры, с которыми запущен код</param>
    ''' <returns>пустая строка, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function CreateUserClass(ByRef masCode() As DictionaryEntry, ByVal className As String, ByVal startPos As Integer, ByRef finalPos As Integer, ByRef arrParams() As String) As String
        'Получаем значение первого имени класса
        If className.Length = 0 Then
            LAST_ERROR = "Неверная запись блока New Class. Не указано имя класса."
            Return "#Error"
        End If
        className = GetValue(className, arrParams)
        If className = "#Error" Then Return "#Error"
        className = PrepareStringToPrint(className, arrParams)
        If mainClassHash.ContainsKey(className) Then
            LAST_ERROR = "Класс с именем " + className + " уже существует."
            Return "#Error"
        End If

        'Заполняем начальную структуру
        Dim strResult As String = ""
        Dim classId As Integer = mainClass.Length
        Dim commaIndex As Integer
        ReDim Preserve mainClass(classId)
        mainClass(classId) = New MainClassType
        ReDim mainClass(classId).Names(0)
        mainClass(classId).Names(0) = className 'первое имя
        mainClass(classId).Properties = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
        mainClass(classId).Properties.Add("Name", New PropertiesInfoType With {.Value = "'default'", .Description = "Имя элемента"})
        mainClass(classId).Functions = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase) 'создаем объект с готовым набором базовых функций (Create, Remove, Count и т. д.)
        For i As Integer = 0 To basicFunctionsHashLevel3.Count - 1
            mainClass(classId).Functions.Add(basicFunctionsHashLevel3.ElementAt(i).Key, basicFunctionsHashLevel3.ElementAt(i).Value.Clone)
        Next
        mainClass(classId).LevelsCount = 2 'кол-во уровней (для пользовательских классов = 2 - максимальное)
        mainClass(classId).UserAdded = True 'означает, что это не встроенный класс
        mainClassHash.Add(className, classId) 'добавляем имя в хэш

        Dim curString As String = ""
        Dim lineIndex As Integer = startPos + 1
        'Главный цикл. Перебираем все строки, по не найдем End Class
        Do While lineIndex < masCode.Length
            curString = masCode(lineIndex).Value
            CURRENT_LINE += 1
            If curString = "End Class" Then
                finalPos = lineIndex
                Return ""
            ElseIf curString.StartsWith("Name ") Then 'Еще одно имя класса
                strResult = PrepareStringToPrint(curString.Substring(5), arrParams)
                If mainClassHash.ContainsKey(strResult) Then
                    LAST_ERROR = "Класс с именем " + className + " уже существует."
                    Return "#Error"
                End If
                ReDim Preserve mainClass(classId).Names(mainClass(classId).Names.Length)
                mainClass(classId).Names(mainClass(classId).Names.Length - 1) = strResult
                mainClassHash.Add(strResult, classId)
            ElseIf curString.StartsWith("Prop ") Then 'Свойство
                strResult = curString.Substring(5)
                'допустимая запись:
                'Prop 'myProp'      ,       Prop 'myProp', 10     и     Prop 'myProp' = 10
                'получаем позицию разделителя между именем свойства и его значением по-умолчанию
                commaIndex = strResult.IndexOf(", ")
                If commaIndex = -1 Then commaIndex = strResult.IndexOf(" = ")


                If commaIndex = -1 Then
                    'Значение по-умолчанию не установлено
                    mainClass(classId).Properties.Add(PrepareStringToPrint(strResult, arrParams), New PropertiesInfoType With {.UserAdded = True, .Value = ""})
                Else
                    'Значение по-умолчанию установлено
                    mainClass(classId).Properties.Add(PrepareStringToPrint(strResult.Substring(0, commaIndex), arrParams), New PropertiesInfoType With {.UserAdded = True, .Value = strResult.Substring(commaIndex + 2).TrimStart})
                End If
            ElseIf curString.StartsWith("Func ") Then
                'Функция пользователя. Допустима запись только с указанием обработчика:
                'Func 'myFunc', 'myHandler'         и           Func 'myFunc' = 'myHandler'
                strResult = curString.Substring(5)
                'получаем позицию разделителя между именем функции и ее функцией-обработчиком
                commaIndex = strResult.IndexOf(", ")
                If commaIndex = -1 Then commaIndex = strResult.IndexOf(" = ")
                If commaIndex = -1 Then
                    LAST_ERROR = "За оператором Func не стоит имя обработчика."
                    Return "#Error"
                End If
                mainClass(classId).Functions.Add(PrepareStringToPrint(strResult.Substring(0, commaIndex), arrParams), New PropertiesInfoType With {.UserAdded = True, .Value = strResult.Substring(commaIndex + 2).TrimStart})
            ElseIf curString.StartsWith("Function ") Then
                'Блок функции (создание функции класса без указания обработчика)
                Dim finLine As Integer
                Dim res As String = BlockFunction(masCode, lineIndex, finLine)
                If res = "#Error" Then Return res

                If finLine = lineIndex Then 'пустой блок
                    lineIndex += 2
                    Continue Do
                End If
                'получаем первую и последнюю строку на исполнение относительно начала кода
                Dim blockStart As Integer = lineIndex + 1, blockFinish As Integer = finLine
                Dim funcName As String = GetValue(curString.Substring("Function ".Length).Trim, arrParams)
                If funcName = "#Error" Then
                    Return "#Error"
                End If
                'создаем подмассив, содержащий код фунции
                Dim subArray() As DictionaryEntry
                ReDim subArray(blockFinish - blockStart)
                Array.ConstrainedCopy(masCode, blockStart, subArray, 0, blockFinish - blockStart + 1)
                Dim lineDifference As Long = subArray(0).Key - 1
                Dim sb As New System.Text.StringBuilder
                For i = 0 To subArray.GetUpperBound(0)
                    subArray(i).Key -= lineDifference
                    sb.AppendLine(subArray(i).Value)
                Next
                'создаем исполняемый код
                questEnvironment.codeBoxShadowed.codeBox.IsTextBlockByDefault = False
                questEnvironment.codeBoxShadowed.Text = sb.ToString
                Dim exData As New List(Of ExecuteDataType)
                exData = PrepareBlock(questEnvironment.codeBoxShadowed.codeBox.CodeData)
                questEnvironment.codeBoxShadowed.Text = ""

                'получаем имя функции
                funcName = PrepareStringToPrint(funcName, arrParams)
                Dim pos As Integer = funcName.IndexOf("."c)
                If pos > -1 AndAlso pos < funcName.Length - 1 Then funcName = funcName.Substring(pos + 1)

                'Это функция класса, добавленная Писателем. Вносим соответствующие изменения в структуру mainClass и создаем событие
                If mainClass(classId).Functions.ContainsKey(funcName) = False Then
                    If AddUserFunction({"'" & mainClass(classId).Names(0) & "'", WrapString(funcName)}, arrParams) = "#Error" Then
                        'Если указана пока еще не созданная функция пользователя - создаем ее
                        LAST_ERROR = "Неверный синтаксис блока Function. Не удалось зодать новую функцию " + funcName + " в классе " + mainClass(classId).Names(0)
                        Return "#Error"
                    End If
                End If

                Dim func As PropertiesInfoType = mainClass(classId).Functions(funcName)
                func.Value = "" 'создавать серализованный аналог необязательно
                func.eventId = eventRouter.SetEventId(func.eventId, exData) 'создаем/заменяем событие

                lineIndex = finLine + 1 'устанавливаем текущую позицию за Loop
            Else
                LAST_ERROR = "Недопустимая строка в блоке New Class."
                Return "#Error"
            End If
            lineIndex += 1
        Loop
        'Цикл завершен, а End Class так и не найден. Ошибка.
        LAST_ERROR = "Не найден конец блока End Class."
        Return "#Error"
    End Function

    ''' <summary>
    ''' Функция обрабатывает блок New Class и создает новый класс (класс пользователя) в структуре mainClass.
    ''' </summary>
    ''' <param name="exCode">весь выполняемый код</param>
    ''' <param name="startPos">положение строки New Class  в массиве masCode()</param>
    ''' <param name="finalPos">индекс последней строки кода внутри блока, которую надо обработать</param>
    ''' <param name="arrParams">параметры, с которыми запущен код</param>
    ''' <returns>пустая строка, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function CreateUserClass(ByVal exCode As List(Of ExecuteDataType), ByVal startPos As Integer, ByRef finalPos As Integer, ByRef arrParams() As String) As String
        'Получаем значение первого имени класса
        Dim className As String = 0
        If exCode.Count < 3 Then
            LAST_ERROR = "Неверная запись блока New Class. Не указано имя класса."
            Return "#Error"
        End If

        className = GetValue(exCode(startPos).Code, 2, -1, arrParams)
        If className = "#Error" Then Return "#Error"
        className = PrepareStringToPrint(className, arrParams)
        If mainClassHash.ContainsKey(className) Then
            LAST_ERROR = "Класс с именем " + className + " уже существует."
            Return "#Error"
        End If

        'Заполняем начальную структуру
        Dim strResult As String = ""
        Dim classId As Integer = mainClass.Length
        ReDim Preserve mainClass(classId)
        mainClass(classId) = New MainClassType
        ReDim mainClass(classId).Names(0)
        mainClass(classId).Names(0) = className 'первое имя
        mainClass(classId).Properties = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
        mainClass(classId).Properties.Add("Name", New PropertiesInfoType With {.Value = "'default'", .Description = "Имя элемента"})
        'mainClass(classId).Functions = New SortedList(Of String, PropertiesInfoType)(basicFunctionsHashLevel3) 'создаем объект с готовым набором базовых функций (Create, Remove, Count и т. д.)
        mainClass(classId).Functions = New SortedList(Of String, PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
        For i As Integer = 0 To basicFunctionsHashLevel3.Count - 1
            mainClass(classId).Functions.Add(basicFunctionsHashLevel3.ElementAt(i).Key, basicFunctionsHashLevel3.ElementAt(i).Value.Clone)
        Next
        mainClass(classId).LevelsCount = 2 'кол-во уровней (для пользовательских классов = 2 - максимальное)
        mainClass(classId).UserAdded = True 'означает, что это не встроенный класс
        mainClassHash.Add(className, classId) 'добавляем имя в хэш


        'Главный цикл. Перебираем все строки, пока не найдем End Class
        Dim lineIndex As Integer = startPos + 1
        Dim curCode() As CodeTextBox.EditWordType
        Do While lineIndex < exCode.Count
            CURRENT_LINE += 1
            curCode = exCode(lineIndex).Code
            If curCode.Length < 2 OrElse (curCode(0).wordType <> CodeTextBox.EditWordTypeEnum.W_BLOCK_NEWCLASS AndAlso curCode(0).wordType <> CodeTextBox.EditWordTypeEnum.W_CYCLE_END _
                AndAlso curCode(0).wordType <> CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION) Then
                LAST_ERROR = "Недопустимая строка в блоке New Class."
                Return "#Error"
            ElseIf curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_CYCLE_END AndAlso curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_NEWCLASS Then
                finalPos = lineIndex
                Return ""
            ElseIf curCode(0).Word.Trim = "Name" Then 'Еще одно имя класса
                strResult = PrepareStringToPrint(curCode(1).Word, arrParams)
                If mainClassHash.ContainsKey(strResult) Then
                    LAST_ERROR = "Класс с именем " + className + " уже существует."
                    Return "#Error"
                End If
                ReDim Preserve mainClass(classId).Names(mainClass(classId).Names.Length)
                mainClass(classId).Names(mainClass(classId).Names.Length - 1) = strResult
                mainClassHash.Add(strResult, classId)
            ElseIf curCode(0).Word.Trim = "Prop" Then 'Свойство
                'допустимая запись:
                'Prop 'myProp'      ,       Prop 'myProp', 10     и     Prop 'myProp' = 10
                'имя свойства
                Dim propName As String = PrepareStringToPrint(curCode(1).Word, arrParams)
                If curCode.Length = 2 Then
                    'Значение по-умолчанию не установлено
                    mainClass(classId).Properties.Add(propName, New PropertiesInfoType With {.UserAdded = True, .Value = ""})
                Else
                    'Значение по-умолчанию установлено
                    If curCode.Length <> 4 Then
                        LAST_ERROR = "Направильная запись присвоения значения свойству."
                        Return "#Error"
                    End If
                    mainClass(classId).Properties.Add(propName, New PropertiesInfoType With {.UserAdded = True, .Value = curCode(3).Word})
                End If
            ElseIf curCode(0).Word.Trim = "Func" Then
                'Функция пользователя. Допустима запись только с указанием обработчика:
                'Func 'myFunc', 'myHandler'         и           Func 'myFunc' = 'myHandler'
                If curCode.Length <> 4 Then
                    LAST_ERROR = "За оператором Func не стоит имя обработчика или неправильная запись функции."
                    Return "#Error"
                End If
                Dim funcName As String = PrepareStringToPrint(curCode(1).Word, arrParams)
                mainClass(classId).Functions.Add(funcName, New PropertiesInfoType With {.UserAdded = True, .Value = curCode(3).Word})
            ElseIf curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION Then
                'Блок функции (создание функции класса без указания обработчика)
                Dim finLine As Integer
                Dim res As String = BlockFunction(exCode, lineIndex, finLine)
                If res = "#Error" Then Return res


                'получаем имя функции
                Dim funcName As String = mScript.PrepareStringToPrint(curCode(1).Word, arrParams, True)
                Dim pos As Integer = funcName.IndexOf("."c)
                If pos > -1 AndAlso pos < funcName.Length - 1 Then funcName = funcName.Substring(pos + 1)

                'создаем блок кода функции
                Dim codeCopy As New List(Of ExecuteDataType)
                For i As Integer = lineIndex + 1 To finLine
                    codeCopy.Add(New ExecuteDataType(exCode(i).Code, i - lineIndex - 1))
                Next

                    'Это функция класса, добавленная Писателем. Вносим соответствующие изменения в структуру mainClass и создаем событие
                If mainClass(classId).Functions.ContainsKey(funcName) = False Then
                    If AddUserFunction({mainClass(classId).Names(0), WrapString(funcName)}, arrParams) = "#Error" Then
                        'Если указана пока еще не созданная функция пользователя - создаем ее
                        LAST_ERROR = "Неверный синтаксис блока Function. Не удалось зодать новую функцию " + funcName + " в классе " + mainClass(classId).Names(0)
                        Return "#Error"
                    End If
                End If

                Dim func As PropertiesInfoType = mainClass(classId).Functions(funcName)
                func.Value = "" 'создавать сериализованный аналог необязательно
                func.eventId = eventRouter.SetEventId(func.eventId, codeCopy) 'создаем/заменяем событие

                lineIndex = finLine + 1
            Else
                LAST_ERROR = "Недопустимая строка в блоке New Class."
                Return "#Error"
            End If

            lineIndex += 1
        Loop
        'Цикл завершен, а End Class так и не найден. Ошибка.
        LAST_ERROR = "Не найден конец блока End Class."
        Return "#Error"
    End Function

    ''' <summary>
    ''' Выполняет оператор Rem Class - удаляет класс пользователя
    ''' </summary>
    ''' <param name="className">одно из имен класса на удаление</param>
    ''' <remarks></remarks>
    Private Sub RemoveUserClass(ByVal className As String)
        If mainClassHash.ContainsKey(className) = False Then Exit Sub
        Dim classId As Integer = mainClassHash(className) 'получаем идентификатор класса в mainClass
        If mainClass(classId).UserAdded = False Then Exit Sub 'встроенные классы удалять нельзя

        'самоудаление
        For i = classId To mainClass.GetUpperBound(0) - 1
            mainClass(i) = mainClass(i + 1)
        Next
        ReDim Preserve mainClass(mainClass.GetUpperBound(0) - 1)
        'переделываем хэш (без имен удаленного класса)
        MakeMainClassHash()
    End Sub
#End Region

#Region "Main Functions"
    ''' <summary>
    ''' Выдает информацию об ошибке при выполении кода.
    ''' </summary>
    ''' <param name="errorCode">Структура ExecuteDataType со строкой, в которой произошла ошибка</param>
    ''' <param name="arrParams">параметры, переданные коды при выполнении</param>
    ''' <remarks></remarks>
    Public Overridable Sub AnalizeError(ByVal errorCode As ExecuteDataType, ByRef arrParams() As String)
        EXIT_CODE = True
        csLocalVariables.KillVars()
        If LAST_ERROR.IndexOf("#DON'T_SHOW#") > -1 Then Exit Sub
        Dim errorString As String = errorCode.CompileString
        If errorString.Length > 100 Then errorString = errorString.Substring(100) + "..."
        Dim strErrorText As String
        strErrorText = IIf(LAST_ERROR.Length = 0, "Неизвестная ошибка.", LAST_ERROR)
        strErrorText += vbNewLine + vbNewLine + "Строка № " + Convert.ToString(errorCode.lineId) + ":" + vbNewLine + errorString + vbNewLine

        If IsNothing(codeStack) = False Then
            strErrorText += vbNewLine + "Содержимое стека: " + String.Join("->", codeStack.ToArray.Reverse)
        End If

        strErrorText += vbNewLine + vbNewLine + "Параметры: "
        If IsNothing(arrParams) OrElse arrParams.GetUpperBound(0) = -1 Then
            strErrorText += "нет"
        Else
            For i As Integer = 0 To arrParams.GetUpperBound(0)
                strErrorText += vbNewLine + "Param[" + Convert.ToString(i) + "] = " + arrParams(i)
            Next i
        End If
        MessageBox.Show(strErrorText, "Ошибка при выполнении кода", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    ''' <summary>
    ''' Выдает информацию об ошибке при выполении кода.
    ''' </summary>
    ''' <param name="errorLine">элемент DictionaryEntry с ключом-номером строки и значением-самой строкой, в которой произошла ошибка</param>
    ''' <param name="arrParams">параметры, переданные коды при выполнении</param>
    ''' <remarks></remarks>
    Public Overridable Sub AnalizeError(ByVal errorLine As DictionaryEntry, ByRef arrParams() As String)
        EXIT_CODE = True
        csLocalVariables.KillVars()
        If LAST_ERROR.IndexOf("#DON'T_SHOW#") > -1 Then Exit Sub
        If errorLine.Value.ToString.Length > 100 Then errorLine.Value = errorLine.Value.ToString.Substring(100) + "..."
        Dim strErrorText As String
        strErrorText = IIf(LAST_ERROR.Length = 0, "Неизвестная ошибка.", LAST_ERROR)
        strErrorText += vbNewLine + vbNewLine + "Строка № " + Convert.ToString(errorLine.Key) + ":" + vbNewLine + errorLine.Value + vbNewLine

        If IsNothing(codeStack) = False Then
            strErrorText += vbNewLine + "Содержимое стека: " + String.Join("->", codeStack.ToArray.Reverse)
        End If

        strErrorText += vbNewLine + vbNewLine + "Параметры: "
        If IsNothing(arrParams) OrElse arrParams.GetUpperBound(0) = -1 Then
            strErrorText += "нет"
        Else
            For i As Integer = 0 To arrParams.GetUpperBound(0)
                strErrorText += vbNewLine + "Param[" + Convert.ToString(i) + "] = " + arrParams(i)
            Next i
        End If
        MessageBox.Show(strErrorText, "Ошибка при выполнении кода", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    ''' <summary>
    ''' Создает массив параметров функции или свойства, заключенных в их квадратных и/или кругых скобках, уже просчитанных
    ''' и готовых к передаче, а также возвращает индекс элемента кода, следующего за функцией. В случае ишибки функция возвращает #Error.
    ''' </summary>
    ''' <param name="masCode">массив со структурой кода</param>
    ''' <param name="firstElementId">Id начала свойства/функции в структуре кода. Это может быть название класса или имя 
    ''' свойства/функции</param>
    ''' <param name="funcParams">массив, куда собираютя параметры функции</param>
    ''' <param name="returnFormat">для получения формата, в котором возвращать результат. Если символа #/$ не было, то 
    ''' переменная не меняет свое значение</param>
    ''' <param name="functionPos">Индекс функции/свойства в массиве masCode</param>
    ''' <param name="nextElementAfterFunction">Id элемента, следующего за функцией. Если код закончился, то -1</param>
    ''' <param name="arrParams">массив параметров, с которыми запущен код</param>
    ''' <param name="canBeNoBrackets">может ли здесь находится вызов функции как процедуры (т. е. без скобок, напр. "myFunc x, y")</param>
    ''' <returns>пустую строку, в случае ошибки #Error</returns>
    ''' <remarks></remarks>
    Public Overridable Function GetFunctionParams(ByRef masCode() As CodeTextBox.EditWordType, ByVal firstElementId As Integer, _
                                ByRef funcParams As List(Of String), ByRef returnFormat As ReturnFormatEnum, _
                                ByRef functionPos As Integer, ByRef nextElementAfterFunction As Integer, _
                                ByRef arrParams() As String, Optional ByVal canBeNoBrackets As Boolean = False) As String
        'Class[x, y].#fName(u, v) ...
        'Если перед функцией стоит #/$, то они уже учтены
        funcParams = New List(Of String) 'создаем пустой массив для параметров функции

        Dim curPos As Integer = firstElementId 'текущий элемент кода
        Dim qbBalance As Integer = 0 'баланс [ ]
        Dim obBalance As Integer = 0 'баланс ( )
        Dim strResult As String 'для получения результатов GetValue
        Dim paramStartPos As Integer 'для сохранении позиции начала текущего параметра

        If masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
            'функция начинается с класса
            'Class[x, y].#fName(u, v) ...
            curPos += 1
            If masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                'Class[
                'получаем индекс следующего элемента после [ ]
                Dim qbNextPos As Integer = GetNextElementIdAfterBrackets(masCode, curPos)
                If qbNextPos = -2 Then
                    Return "#Error"
                ElseIf qbNextPos = -1 Then
                    'код заканчивается скобкой ] - ошибка
                    LAST_ERROR = "Ошибка в синтаксисе. Неверная запись функции/свойства."
                    Return "#Error"
                End If

                'в цикле собираем все параметры в список funcParams
                paramStartPos = curPos + 1
                For i As Integer = curPos + 1 To qbNextPos - 2
                    Select Case masCode(i).wordType
                        Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                            qbBalance += 1
                        Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                            qbBalance -= 1
                        Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                            obBalance += 1
                        Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                            obBalance -= 1
                        Case CodeTextBox.EditWordTypeEnum.W_COMMA
                            If qbBalance <> 0 OrElse obBalance <> 0 Then Continue For
                            'кол-во открытых и закрытых скобок равно - значит это запятая этой функции и она разделяет параметры
                            'вычисляем значение параметра
                            strResult = GetValue(masCode, paramStartPos, i - 1, arrParams)
                            If strResult = "#Error" Then Return strResult
                            'сохраняем в список результат
                            funcParams.Add(strResult)
                            paramStartPos = i + 1
                    End Select
                Next
                'добавляем в список последний параметр
                strResult = GetValue(masCode, paramStartPos, qbNextPos - 2, arrParams)
                If strResult = "#Error" Then Return strResult
                funcParams.Add(strResult)

                'проходим точку после скобок [x].
                If masCode(qbNextPos).wordType <> CodeTextBox.EditWordTypeEnum.W_POINT Then
                    LAST_ERROR = "Ошибка в синтаксисе. Неверная запись функции/свойства. За квадратными скобками не стоит точка."
                    Return "#Error"
                End If
                curPos = qbNextPos + 1
                'если есть #/$, то указываем returnFormat
                '#fName(u, v) ...
                If masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Then
                    returnFormat = ReturnFormatEnum.TO_NUMBER
                    curPos += 1
                ElseIf masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING Then
                    returnFormat = ReturnFormatEnum.TO_STRING
                    curPos += 1
                End If
            ElseIf masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_POINT Then
                'квадратных скобок нет - Class.#fName(u, v) ...
                curPos += 1
                'если есть #/$, то указываем returnFormat
                '#fName(u, v) ...
                If masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Then
                    returnFormat = ReturnFormatEnum.TO_NUMBER
                    curPos += 1
                ElseIf masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING Then
                    returnFormat = ReturnFormatEnum.TO_STRING
                    curPos += 1
                End If
            Else
                LAST_ERROR = "Ошибка в синтаксисе. Неверная запись функции/свойства."
                Return "#Error"
            End If
        End If

        'fName(u, v) ... / fName u, v /myVar[u]
        If masCode(curPos).wordType <> CodeTextBox.EditWordTypeEnum.W_FUNCTION AndAlso _
            masCode(curPos).wordType <> CodeTextBox.EditWordTypeEnum.W_PROPERTY AndAlso
            masCode(curPos).wordType <> CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
            'LAST_ERROR = "Ошибка в синтаксисе. Неверная запись функции/свойства."
            'Return "#Error"
            'без имени функции. Например, параметры блока Event (Event 'myEvent', 5, Ture, 'ok')
            functionPos = -1
        Else
            functionPos = curPos
        End If
        curPos += 1

        'в данном месте если были [], то они пройдены. Также в returnFormat получен #/$. Остались только круглые скобки (если они есть)
        '(u, v) ...
        'или же, если это переменная, то еще может быть содержимое квадратных скорок ЗА переменной
        'получаем элемент сразу за скобками
        nextElementAfterFunction = GetNextElementIdAfterBrackets(masCode, curPos)
        If nextElementAfterFunction = -2 Then Return "#Error"
        If curPos >= masCode.Length Then Return "" 'нет параметров - выход
        If canBeNoBrackets = False AndAlso (masCode(curPos).wordType <> CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN _
                                           AndAlso masCode(curPos).wordType <> CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN) Then
            Return "" 'нет параметров - выход
        End If
        Dim lastElement As Integer = 0 'последний элемент кода внутри скобок (если они есть), принадлежащий данной функции
        If nextElementAfterFunction = curPos Then
            If canBeNoBrackets = False Then Return "" 'скобок нет - выход
            lastElement = masCode.GetUpperBound(0)
            If curPos > lastElement Then Return "" 'последний элемент - сама функция. Выход
        ElseIf nextElementAfterFunction = -1 Then
            'последний элемент конец строки
            lastElement = masCode.GetUpperBound(0)
        Else
            lastElement = nextElementAfterFunction - 1
        End If
        If masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN _
            OrElse masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
            'функция имеет скобки - тогда сдвигаем lastElement на 1 влево, убирая последнюю закрывающую скобку
            lastElement -= 1
            'а текущую позицию смещаем на 1 вправо за открывающую скобку
            curPos += 1
        End If

        'в цикле получаем все параметры, вычисляем их и добавляем в список funcParams
        qbBalance = 0
        obBalance = 0
        paramStartPos = curPos
        For i As Integer = curPos To lastElement
            Select Case masCode(i).wordType
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                    qbBalance += 1
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                    qbBalance -= 1
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                    obBalance += 1
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                    obBalance -= 1
                Case CodeTextBox.EditWordTypeEnum.W_COMMA
                    If qbBalance <> 0 OrElse obBalance <> 0 Then Continue For
                    'кол-во открытых и закрытых скобок равно - значит это запятая этой функции и она разделяет параметры
                    'вычисляем значение параметра
                    strResult = GetValue(masCode, paramStartPos, i - 1, arrParams)
                    If strResult = "#Error" Then Return strResult
                    funcParams.Add(strResult)
                    paramStartPos = i + 1
                Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH
                    If masCode(i).Word.TrimEnd.EndsWith("=") Then Return "" '+= / -=
            End Select
        Next
        'получаем последний параметр
        If functionPos > -1 AndAlso masCode(functionPos).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
            If masCode(paramStartPos).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                masCode(paramStartPos).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE OrElse _
                (masCode(paramStartPos).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH AndAlso (masCode(paramStartPos).Word.Trim = "+=" OrElse masCode(paramStartPos).Word.Trim = "-=")) Then
                '= или <> != после свойства. Значит параметров данного свойства больше нет
                Return ""
            End If
        End If
        strResult = GetValue(masCode, paramStartPos, lastElement, arrParams)
        If strResult = "#Error" Then Return strResult
        funcParams.Add(strResult)
        Return ""
    End Function

    ''' <summary>Функция создает массив arrayResult из всех параметров функции/свойства (просчитанных и готовых к передаче), а также возвращает положение закрывающей скобки функции 
    ''' и определяет количество прилегающих к ней справа закрывающих скобок. Например: "func(dd))))" : finalBracketCount = 3.
    ''' При этом, если у функции нет собственных скобок, то посчитать прилегающие к ней ) не представляется возможным (они считаются в GetValue).
    ''' Массив создается из предварительно полученных параметров в квадратных скобках, и из параметров в скобках, идущих за именем функции.
    ''' </summary>
    ''' <param name="strExpression">строка исходного кода, где используется нужная функция</param>
    ''' <param name="arrayResult">массив, который будет хранить параметры нужной функции</param>
    ''' <param name="funcOpenBracketPos">положение в strExpression открывающей скобки функции (если скобок нет, то -1)</param>
    ''' <param name="funcCloseBracketPos">переменная для хранения позиции в строке ), закрывающей функцию</param>
    ''' <param name="QBcontent">содержимое квадратных скобок перед именем функции, предварительно полученных с помощью GetWordType()</param>
    ''' <param name="finalBracketCount">количество скобок ), идущих вслед за закрывающей скобкой функции</param>
    ''' <param name="arrParams">массив с параметрами, с которыми выполняется код</param>
    ''' <param name="getParamsValues">вычислять ли значения параметров (только для квадратных скобок, используется в функции BlockSwitch)</param>
    Public Overridable Function GetFunctionParams(ByVal strExpression As String, ByRef arrayResult() As String, ByVal funcOpenBracketPos As Integer, ByRef funcCloseBracketPos As Integer,
                                       ByVal QBcontent As String, ByRef finalBracketCount As Integer, ByRef arrParams() As String, Optional ByVal getParamsValues As Boolean = True) As Integer
        finalBracketCount = 0
        funcCloseBracketPos = -1
        'Если нет ни круглых, ни квадратных скобок, значит нет и параметров. Выход
        If funcOpenBracketPos = -1 AndAlso QBcontent.Length = 0 Then Return funcOpenBracketPos

        Dim i As Integer, j As Integer
        Dim paramStart As Integer = -1 'начало параметра в строке (от начала строки или от запятой)
        Dim chBr As Integer = -1 'положение '
        Dim chComma As Integer = -1 'положение ,
        Dim chCircBrOpen As Integer = -1 'положение (
        Dim chCircBrClose As Integer = -1 'положение )
        Dim chCircBrBalance As Integer = 0 'на сколько ( больше, чем )

        Erase arrayResult 'на всякий случай очищаем массив для параметров функции
        Dim arrayResultLength As Integer = -1 'длина массива параметров (чтобы не проверять каждый раз)
        Dim qbHidden As String

        '1. Обработка [...] ************************************************
        If QBcontent.Length > 0 Then
            'Квадратные скобки есть
            qbHidden = HideQBcontent(QBcontent) 'прячем вложенные квадратные скобки, чтоб не путались
qbSt:
            'Ищем , и ' (так как встроке '...' могут быть нефункциональные запятые)
            If chBr < i Then
                chBr = qbHidden.IndexOf("'", i)
                If chBr = -1 Then chBr = Integer.MaxValue
            End If
            If chComma < i Then
                chComma = qbHidden.IndexOf(",", i)
                If chComma = -1 Then chComma = Integer.MaxValue
            End If
            If chCircBrOpen < i Then
                chCircBrOpen = qbHidden.IndexOf("(", i)
                If chCircBrOpen = -1 Then chCircBrOpen = Integer.MaxValue
            End If

            'получаем то, что в строке встретилось раньше
            i = {chBr, chComma, chCircBrOpen}.Min

            If chComma = Integer.MaxValue Then
                'запятых нет (или больше нет)
                'Расширяем массив
                arrayResultLength += 1
                ReDim Preserve arrayResult(arrayResultLength)
                'Получаем строку до конца и сохраняем ее значение в новой ячейке массива
                If getParamsValues Then
                    arrayResult(arrayResultLength) = GetValue(QBcontent.Substring(paramStart + 1), arrParams)
                    If arrayResult(arrayResultLength) = "#Error" Then Return -2
                Else
                    arrayResult(arrayResultLength) = QBcontent.Substring(paramStart + 1)
                End If
                GoTo qbNx
            End If

            If QBcontent.Chars(i) = ","c Then
                'Найдена ,
                'Расширяем массив
                arrayResultLength += 1
                ReDim Preserve arrayResult(arrayResultLength)
                'Получаем строку с параметром и сохраняем ее значение в новой ячейке массива
                If getParamsValues Then
                    arrayResult(arrayResultLength) = GetValue(QBcontent.Substring(paramStart + 1, i - paramStart - 1), arrParams)
                    If arrayResult(arrayResultLength) = "#Error" Then Return -2
                Else
                    arrayResult(arrayResultLength) = QBcontent.Substring(paramStart + 1, i - paramStart - 1)
                End If
                'Запоминаем начальное положение следующего параметра
                paramStart = i + 1
                i += 1
            ElseIf QBcontent.Chars(i) = "("c Then
                'Найдена (, которая может быть скобкой внутренней функции(и та может содержать свои запятые, не нужные нам)
                chCircBrBalance = 1 'пока одна открывающая
                'Продвигаемся по строке QBcontent, пока "кол-во открытых" - "кол-во закрытых" не станет равно 0
                Do Until chCircBrBalance = 0
                    i += 1
                    If chCircBrOpen < i Then 'позиция ближайшей (
                        chCircBrOpen = qbHidden.IndexOf("(", i)
                        If chCircBrOpen = -1 Then chCircBrOpen = Integer.MaxValue
                    End If
                    If chCircBrClose < i Then 'позиция ближайшей )
                        chCircBrClose = qbHidden.IndexOf(")", i)
                        If chCircBrClose = -1 Then chCircBrClose = Integer.MaxValue
                    End If
                    If chBr < i Then
                        chBr = qbHidden.IndexOf("'", i) 'позиция ближайшей ' (в строке могут быть нефункциональные скобки)
                        If chBr = -1 Then chBr = Integer.MaxValue
                    End If
                    i = {chCircBrOpen, chCircBrClose, chBr}.Min 'получаем первое, что встретилось
                    If i = Integer.MaxValue Then
                        LAST_ERROR = "Непарное количество круглых скобок внутри квадратных скобок."
                        Return -2
                    End If
                    If QBcontent.Chars(i) = "(" Then
                        chCircBrBalance += 1
                    ElseIf QBcontent.Chars(i) = ")" Then
                        chCircBrBalance -= 1
                    Else
                        'Найдена строка '...'.
qbQuoteSearching2:
                        i = qbHidden.IndexOf("'", i + 1) 'ищем закрывающий "'"
                        If i = -1 Then
                            'закрывающий "'" не найден - ошибка
                            LAST_ERROR = "Не найдена закрывающая кавычка"
                            Return -2
                        End If
                        'Обрабатываем экранированную кавычку /'
                        If qbHidden.Chars(i - 1) = "/"c Then GoTo qbQuoteSearching2
                    End If
                Loop
            Else
                'Найдена строка '...'.
qbQuoteSearching:
                i = qbHidden.IndexOf("'", i + 1) 'ищем закрывающий "'"
                If i = -1 Then
                    'закрывающий "'" не найден - ошибка
                    LAST_ERROR = "Не найдена закрывающая кавычка"
                    Return -2
                End If
                'Обрабатываем экранированную кавычку /'
                If qbHidden.Chars(i - 1) = "/"c Then GoTo qbQuoteSearching
            End If
            'Новый виток цикла
            i = i + 1
            GoTo qbSt
qbNx:
        End If

        '2. Обработка (...) ************************************************
        If funcOpenBracketPos > -1 Then
            qbHidden = HideQBcontent(strExpression)
            'Круглые скобки есть
            i = funcOpenBracketPos + 1 'Поиск с символа, следующего за открывающей скобкой
            chCircBrBalance = 1 'для хранения количества вложенных скобок (одна - открывающая скобка функции - уже есть)
            paramStart = funcOpenBracketPos  'положение в строке начала текущего параметра
            'Для хранения положения последних найденных ( ) ' ,
            Dim chOpn As Integer = -1, chCls As Integer = -1
            chBr = -1 'положение последних найденных ( ) '
            chComma = -1

st:
            'ищем символы ( ) ' ,
            If chOpn < i Then
                chOpn = qbHidden.IndexOf("(", i)
                If chOpn = -1 Then chOpn = Integer.MaxValue
            End If
            If chCls < i Then
                chCls = qbHidden.IndexOf(")", i)
                If chCls = -1 Then chCls = Integer.MaxValue
            End If
            If chBr < i Then
                chBr = qbHidden.IndexOf("'", i)
                If chBr = -1 Then chBr = Integer.MaxValue
            End If
            If chComma < i Then
                chComma = qbHidden.IndexOf(",", i)
                If chComma = -1 Then chComma = Integer.MaxValue
            End If

            'получаем то, что в строке встретилось раньше
            i = {chOpn, chCls, chComma, chBr}.Min

            'ничего не найдено - ошибка
            If i = Integer.MaxValue Then
                LAST_ERROR = "Не найдена закрывающая скобка у функции."
                Return -2
            End If

            If strExpression.Chars(i) = "("c Then
                chCircBrBalance = chCircBrBalance + 1 'увеличиваем количество вложенных скобок
            ElseIf strExpression.Chars(i) = "'"c Then
                'найдена строка '...'.
QuoteSearching:
                i = strExpression.IndexOf("'", i + 1) 'ищем закрывающую "'"
                If i = -1 Then
                    'закрывающая "'" не найдена - ошибка
                    LAST_ERROR = "Не найдена закрывающая кавычка строки."
                    Return -2
                End If
                'Обрабатываем экранированную кавычку /'
                If strExpression.Chars(i - 1) = "/"c Then GoTo QuoteSearching
            ElseIf strExpression.Chars(i) = ")"c Then
                'нашли ")"
                chCircBrBalance = chCircBrBalance - 1 'уменьшаем количество вложенных скобок
                If chCircBrBalance = 0 Then
                    'найдена закрывающая скобка функции
                    funcCloseBracketPos = i 'получаем положение закрывающей скобки
                    'Расширяем массив
                    arrayResultLength += 1
                    ReDim Preserve arrayResult(arrayResultLength)
                    'Получаем строку с параметром и сохраняем ее значение в новой ячейке массива
                    arrayResult(arrayResultLength) = GetValue(strExpression.Substring(paramStart + 1, i - paramStart - 1), arrParams)
                    If arrayResult(arrayResultLength) = "#Error" Then Return -2
                    'Возвращаем положение закрывающей скобки функции
                    For j = i + 1 To strExpression.Length - 1
                        If strExpression.Chars(j) = ")"c Then
                            finalBracketCount += 1
                        Else
                            Exit For
                        End If
                    Next
                    Return i
                End If
            Else
                'Найдена запятая
                'Если количество открывающих и закрывающих скобок не совпадает (qbCnt <> 1) - мы внутри другой (внутренней) функции
                If chCircBrBalance = 1 Then
                    'Расширяем массив
                    arrayResultLength += 1
                    ReDim Preserve arrayResult(arrayResultLength)
                    'Получаем строку с параметром и сохраняем ее значение в новой ячейке массива
                    arrayResult(arrayResultLength) = GetValue(strExpression.Substring(paramStart + 1, i - paramStart - 1), arrParams)
                    If arrayResult(arrayResultLength) = "#Error" Then Return -2
                    'Запоминаем начальное положение следующего параметра
                    paramStart = i + 1
                    i += 1
                End If
            End If
            i = i + 1
            'Новый виток цикла
            GoTo st
        End If

        Return -1
    End Function

    ''' <summary>
    ''' Одна из главных функций. Возвращает результат выражения из переданного кода
    ''' </summary>
    ''' <param name="masCode">код с вычисляемым выражением</param>
    ''' <param name="startPos">индекс первого элемента выражения</param>
    ''' <param name="endPos">индекс последнего элемента выражения. Если -1, то выполнять до конца строки кода</param>
    ''' <param name="arrParams">параметры, с которыми запущен код</param>
    ''' <returns>Пустую строку, в случае ошибки #Error</returns>
    ''' <remarks></remarks>
    Public Overridable Function GetValue(ByRef masCode() As CodeTextBox.EditWordType, ByVal startPos As Integer, _
                                         ByVal endPos As Integer, ByRef arrParams() As String, Optional ByVal fullExecution As Boolean = True) As String
        If startPos < 0 Then
            LAST_ERROR = "Неверные параметры запуска GetValue. startPos < 0."
            Return "#Error"
        ElseIf startPos > masCode.GetUpperBound(0) Then
            LAST_ERROR = "Неверные параметры запуска GetValue. startPos > количества элементов кода."
            Return "#Error"
        End If

        If endPos = -1 OrElse endPos > masCode.GetUpperBound(0) Then
            endPos = masCode.GetUpperBound(0)
        ElseIf endPos < startPos Then
            LAST_ERROR = "Неверные параметры запуска GetValue. endPos < startPos."
            Return "#Error"
        End If

        Dim blnStrings, blnNumbers, blnBooleans, blnCode As Boolean 'есть ли в выражении строки, числа, логические
        Dim blnOperatorsMath, blnOperatorsLogic, blnOperatorsCompare, blnOperatorStringsMerger As Boolean 'есть ли в выражении операторы

        Dim values() As String = Nothing
        Dim valuesLength As Integer = -1 'массив и его размер, куда помещать промежуточный результат (в котором будут
        'только числа, строки, True/False, круглые скобки и операторы
        Dim ovalBracketsBalance As Integer = 0 'баланс открытых и закрытых скобок ( )
        Dim bracketsCount = 0 'общее кол-во пар скобок
        Dim returnFormat As ReturnFormatEnum = ReturnFormatEnum.ORIGINAL 'формат, в котором возвращать выражение
        Dim strResult As String 'для получения промежуточных результатов внутренних функций
        Dim nextPos As Integer 'для получения индекса следующего элемента после функции, свойства и др.
        Dim funcParams As List(Of String) = Nothing  'массив для хранения параметров функции
        Dim i As Integer = startPos

        Do While i <= endPos
            Select Case masCode(i).wordType
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN '(
                    ovalBracketsBalance += 1
                    bracketsCount += 1
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = masCode(i).Word.Trim
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE ')
                    ovalBracketsBalance -= 1
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = masCode(i).Word.Trim
                Case CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER '#
                    returnFormat = ReturnFormatEnum.TO_NUMBER
                Case CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING '$
                    returnFormat = ReturnFormatEnum.TO_STRING
                Case CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL 'True/False
                    If startPos = endPos Then Return masCode(i).Word.Trim
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = masCode(i).Word.Trim
                    blnBooleans = True
                Case CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER 'число
                    If startPos = endPos Then Return masCode(i).Word.Trim
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = masCode(i).Word.Trim
                    blnNumbers = True
                Case CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING 'строка
                    If startPos = endPos Then Return masCode(i).Word.Trim
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = masCode(i).Word.Trim
                    blnStrings = True
                Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE, _
                    CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL '< > =
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = masCode(i).Word.Trim
                    blnOperatorsCompare = True
                Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC 'And Or Xor
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = masCode(i).Word.Trim
                    blnOperatorsLogic = True
                Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH '+ - * / ^
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = masCode(i).Word.Trim
                    blnOperatorsMath = True
                Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER '&
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = masCode(i).Word.Trim
                    blnOperatorStringsMerger = True
                Case CodeTextBox.EditWordTypeEnum.W_PARAM_COUNT 'ParamCount
                    'получаем количество параметров
                    Dim resValue As Integer
                    If IsNothing(arrParams) Then
                        resValue = 0
                    Else
                        resValue = arrParams.Length
                    End If
                    'Конвертируем результат (на случай, если был указан модификатор # или $)
                    strResult = ConvertElement(resValue.ToString, returnFormat, ReturnFormatEnum.TO_NUMBER)
                    If strResult = "#Error" Then Return strResult
                    Select Case returnFormat
                        Case ReturnFormatEnum.ORIGINAL, ReturnFormatEnum.TO_NUMBER
                            blnNumbers = True
                        Case ReturnFormatEnum.TO_STRING
                            blnStrings = True
                        Case ReturnFormatEnum.TO_BOOL
                            blnBooleans = True
                        Case ReturnFormatEnum.TO_CODE
                            blnCode = True
                    End Select
                    'сбрасываем значение returnFormat 
                    returnFormat = ReturnFormatEnum.ORIGINAL
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = strResult
                Case CodeTextBox.EditWordTypeEnum.W_PARAM 'Param[x]
                    '#Param[x + y] ...
                    If IsNothing(arrParams) OrElse arrParams.Length = 0 Then
                        LAST_ERROR = "Нельзя получить значение параметра, так как код был запущен без параметров."
                        Return "#Error"
                    End If

                    'получаем индекс элемента сразу за закрывающей скобкой параметра ]
                    nextPos = GetNextElementIdAfterBrackets(masCode, i + 1)
                    If nextPos = -2 Then Return "#Error"
                    'получаем положение закрыващей скобки ]
                    If nextPos = -1 Then
                        nextPos = masCode.GetUpperBound(0)
                    Else
                        nextPos -= 1
                    End If
                    'получаем номер параметра
                    Dim parNumber As Integer = 0
                    If nextPos > i Then
                        'указан номер параметра (а не просто Param). Получаем номер
                        strResult = GetValue(masCode, i + 2, nextPos - 1, arrParams)
                        If strResult = "#Error" Then Return strResult
                        If Param_GetType(strResult) <> ReturnFormatEnum.TO_NUMBER Then
                            LAST_ERROR = "Неверный формат данных. Номер параметра должен быть числом."
                            Return "#Error"
                        End If
                        parNumber = Convert.ToInt32(strResult)
                    End If
                    'получаем значение Param[x]
                    If parNumber < 0 OrElse parNumber > arrParams.Count - 1 Then
                        LAST_ERROR = "Номер параметра указан не верно."
                        Return "#Error"
                    End If
                    strResult = ConvertElement(arrParams(parNumber), returnFormat)
                    If strResult = "#Error" Then Return strResult
                    Select Case Param_GetType(strResult)
                        Case ReturnFormatEnum.TO_NUMBER
                            blnNumbers = True 'В выражении есть число
                            'strWord = strWord.Replace(","c, "."c)
                        Case ReturnFormatEnum.TO_STRING
                            blnStrings = True 'В выражении есть строка
                        Case ReturnFormatEnum.TO_BOOL
                            blnBooleans = True 'В выражении есть True/False
                        Case ReturnFormatEnum.TO_CODE
                            blnCode = True
                    End Select
                    i = nextPos + 1
                    'сбрасываем значение returnFormat 
                    returnFormat = ReturnFormatEnum.ORIGINAL
                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = strResult
                    Continue Do
                Case CodeTextBox.EditWordTypeEnum.W_CLASS, CodeTextBox.EditWordTypeEnum.W_FUNCTION, _
                    CodeTextBox.EditWordTypeEnum.W_PROPERTY, CodeTextBox.EditWordTypeEnum.W_VARIABLE

                    '-1 переменная, -2 неизвестный класс
                    Dim funcPos As Integer = 0
                    'получаем паметры в квадратных и/или круглых скобках переменной, свойства или фукнции
                    strResult = GetFunctionParams(masCode, i, funcParams, returnFormat, funcPos, nextPos, arrParams)
                    If strResult = "#Error" Then Return strResult

                    If masCode(funcPos).classId = -1 Then
                        'получаем индекс переменной
                        Dim varIndex As Integer = 0, varSignature As String = ""
                        If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then
                            Dim ret As ReturnFormatEnum = Param_GetType(funcParams(0))
                            If ret <> ReturnFormatEnum.TO_NUMBER AndAlso ret <> ReturnFormatEnum.TO_STRING Then
                                LAST_ERROR = "Неверный формат данных. Номер переменной должен быть числом/строкой."
                                Return "#Error"
                            End If
                            If ret = ReturnFormatEnum.TO_NUMBER Then
                                varIndex = Convert.ToInt32(funcParams(0))
                            Else
                                varSignature = mScript.PrepareStringToPrint(funcParams(0), arrParams)
                            End If
                        End If

                        'Получаем значение переменной
                        Dim varResult As String
                        If varSignature.Length = 0 Then
                            varResult = csLocalVariables.GetVariable(masCode(funcPos).Word, varIndex) 'Поиск в локальных переменных
                            If varResult <> "#Error" Then
                            Else 'Локальной переменной не существует
                                varResult = csPublicVariables.GetVariable(masCode(funcPos).Word, varIndex) 'Поиск в переменных исследования
                                If varResult = "#Error" Then
                                    LAST_ERROR = "Попытка получить значение несуществующей переменной. " + masCode(funcPos).Word + "[" + _
                                        varIndex.ToString + "] не существует."
                                    Return varResult
                                End If
                            End If
                        Else
                            varResult = csLocalVariables.GetVariable(masCode(funcPos).Word, varSignature) 'Поиск в локальных переменных
                            If varResult <> "#Error" Then
                            Else 'Локальной переменной не существует
                                varResult = csPublicVariables.GetVariable(masCode(funcPos).Word, varSignature) 'Поиск в переменных исследования
                                If varResult = "#Error" Then
                                    LAST_ERROR = "Попытка получить значение несуществующей переменной. " + masCode(funcPos).Word + "[" + _
                                        varSignature + "] не существует."
                                    Return varResult
                                End If
                            End If
                        End If

                        strResult = varResult
                    Else
                        'уточняем класс, свойство/функцию для пользовательских классов и элементов
                        strResult = ClarifyUserElement(masCode(funcPos))
                        If strResult = "#Error" Then
                            LAST_ERROR = "Свойство/функция " + masCode(funcPos).Word + " не найдена."
                            Return strResult
                        End If

                        If masCode(funcPos).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                            strResult = FunctionRouter(masCode(funcPos).classId, masCode(funcPos).Word, funcParams.ToArray, arrParams)
                        Else
                            strResult = PropertiesRouter(masCode(funcPos).classId, masCode(funcPos).Word, funcParams.ToArray, arrParams)
                        End If
                        If strResult = "#Error" Then Return strResult
                    End If

                    'передан массив?
                    If strResult = "#ARRAY" Then
                        If Not fullExecution AndAlso nextPos = -1 AndAlso valuesLength = -1 Then
                            Return strResult
                        Else
                            strResult = CreateTempArray()
                            If strResult = "#Error" Then Return strResult
                        End If
                    ElseIf strResult = "" AndAlso returnFormat = ReturnFormatEnum.ORIGINAL Then
                        'empty
                        If blnOperatorsMath OrElse blnOperatorsLogic OrElse blnOperatorStringsMerger OrElse (nextPos = -1 OrElse nextPos > endPos) = False OrElse valuesLength <> -1 Then
                            Return _ERROR("Данная операция с пустым значением недопустима.")
                        End If
                        Return ""
                    End If

                    strResult = ConvertElement(strResult, returnFormat)
                    If strResult = "#Error" Then Return strResult
                    Select Case Param_GetType(strResult)
                        Case ReturnFormatEnum.TO_NUMBER
                            blnNumbers = True 'В выражении есть число
                            'strWord = strWord.Replace(","c, "."c)
                        Case ReturnFormatEnum.TO_STRING
                            blnStrings = True 'В выражении есть строка
                        Case ReturnFormatEnum.TO_BOOL
                            blnBooleans = True 'В выражении есть True/False
                        Case ReturnFormatEnum.TO_CODE
                            blnCode = True 'В выражении есть код
                    End Select

                    valuesLength += 1
                    Array.Resize(Of String)(values, valuesLength + 1)
                    values(valuesLength) = strResult

                    'сбрасываем значение returnFormat 
                    returnFormat = ReturnFormatEnum.ORIGINAL
                    If nextPos = -1 Then Exit Do
                    i = nextPos
                    Continue Do

            End Select
            i += 1
        Loop







        'Список values, состоящий только из чисел, строк, True/False, операторов и скобок собран
        'Теперь надо получить окончательный результат получившегося выражения

        Dim typesCount As Integer 'Количество типов в выражении
        'Количество типов надо для подстановки в функцию LogicCompare(). Если тип только один, то эта функция выполнится немного быстрее
        If blnNumbers Then
            typesCount += 1 'количество типов + 1 (числа)
            returnFormat = ReturnFormatEnum.TO_NUMBER 'Передадим в LogicCompare(), что у нас только числа
        End If
        If blnBooleans Then
            typesCount += 1 'количество типов + 1 (логические)
            returnFormat = ReturnFormatEnum.TO_BOOL 'Передадим в LogicCompare(), что у нас только True/False
        End If
        If blnStrings Then
            typesCount += 1 'количество типов + 1 (строки)
            returnFormat = ReturnFormatEnum.TO_STRING 'Передадим в LogicCompare(), что у нас только строки
        End If
        If blnCode Then
            If typesCount > 0 OrElse valuesLength > 0 Then
                LAST_ERROR = "Нельзя выполнять логические и математические операции с кодом."
                Return "#Error"
            End If
            Return values(0)
        End If

        If typesCount = 0 Then Return "" 'Типов нет вообще - значит массив пустой - возвращаем пустую строку
        If typesCount > 1 Then returnFormat = ReturnFormatEnum.ORIGINAL 'Типов несколько - значит LogicCompare() будет определять типы самостоятельно

        If valuesLength = 0 Then Return values(0) 'Массив из одного элемента - значит это и есть готовый результат

        If ovalBracketsBalance <> 0 Then
            'Ошибка: количество открывающих и закрывающих скобок не совпадает!
            LAST_ERROR = "Количество открывающих и закрывающих скобок не совпадает."
            Return "#Error"
        End If

        If blnOperatorsCompare = False AndAlso blnOperatorsLogic = False Then
            'Все операторы исключительно математические (или сложения строк)
            If typesCount > 1 Then
                LAST_ERROR = "Несоответствие типов данных."
                Return "#Error"
            End If
            If blnNumbers Then
                'У нас числа - считаем результат уравнения
                Return Calculate(values, bracketsCount)
            ElseIf blnBooleans Then
                'У нас логические превращаем True в 1 и False в 0
                For i = 0 To values.Count - 1 Step 2
                    values(0) = IIf(values(i) = "True", "1", "0")
                Next
                'Считаем результат уравнения. Если получили 0 - возвращаем False; иначе - True
                Return IIf(Calculate(values, bracketsCount) = 0, "False", "True")
            Else
                'У нас строки - выполняем сложение строк
                strResult = "'" 'Вставляем начальную кавычку
                For i = 0 To values.Count - 1 Step 2
                    strResult += values(i).Substring(1, values(i).Length - 2)
                Next
                Return strResult + "'" 'Возвращаем результат, добавив в конце обратную кавычку
            End If
        Else
            'Выполняем логическое сравнение и возвращаем результат
            Return LogicCompare(values, returnFormat, bracketsCount, blnOperatorsMath Or blnOperatorStringsMerger, blnOperatorsCompare, blnOperatorsLogic)
        End If
        Return ""
    End Function

    Private Function CreateTempArray() As String
        If IsNothing(lastArray) Then Return _ERROR("Попытка обратиться к несществующему массиву.")
        Dim strName As String = ""
        If IsNothing(csLocalVariables.lstVariables) OrElse csLocalVariables.lstVariables.Count = 0 Then
            strName = temporary_system_array_BASE & "1"
        Else
            Dim i As Integer = 1
            Do
                strName = temporary_system_array_BASE & i.ToString
                If csLocalVariables.lstVariables.ContainsKey(strName) = False Then Exit Do
                i += 1
            Loop
        End If

        csLocalVariables.SetVariableInternal(strName, "", 0)
        csLocalVariables.lstVariables(strName).arrValues = lastArray.arrValues
        lastArray = Nothing
        Return "'" & strName & "'"
    End Function

    ''' <summary>
    ''' Одна из главных функций. Получает в strExpression любое выражение и возвращает его результат
    ''' </summary>
    ''' <param name="strExpression">выражение для вычисления</param>
    ''' <param name="arrParams">массив с параметрами, переданными при выполнении кода</param>
    Public Overridable Function GetValue(ByVal strExpression As String, ByRef arrParams() As String, Optional ByVal fullExecution As Boolean = True) As String
        'Если выражение простое (т. е. просто число, строка, True/False или пусто) - возвращаем его же
        If IsSimpleValue(strExpression) Then Return strExpression

        Dim strWord As String 'Строка, хранящее очередное слово в выражении
        Dim currentPosition = 0, savePosition As Integer = 0 'текущее положение в strExpression и еще одна переменная для его резервного хранения
        Dim blnStrings, blnNumbers, blnBooleans, blnCode As Boolean 'есть ли в выражении строки, числа, логические
        Dim blnOperatorsMath, blnOperatorsLogic, blnOperatorsCompare, blnOperatorStringsMerger As Boolean 'есть ли в выражении операторы
        Dim openBracketsCount, closeBracketsCount As Integer 'количество открывающих, закрывающих скобок
        Dim newCloseBracketsCount, funcOpenBracketPos As Integer 'количество закрывающих скобок в конце текущего слова и переменная для получения позиции открывающей скобки функции
        Dim currentClassId As Integer, elementName As String = "" 'переменные для получения Id класса и текущей функции/свойства в структуре mainClass
        Dim currentWordType As MatewScript.WordTypeEnum 'Для хранения значения текущего слова (т. е. число, оператор, функция и т. д.)
        Dim returnFormat As MatewScript.ReturnFormatEnum 'Для получения/преобразования формата (число, строка, логическое)
        Dim elementContent As String = "" 'Для получения соответствующего параметра функции GetWordType()
        Dim funcParams As String() = {} 'Массив для хранения параметров функции/свойства
        Dim arrResult() As String = {}, arrResultLength As Integer = -1 'Результирующий массив, в котором помещаются предварительные результаты и его размер (чтобы каждый раз не проверять)
        Dim funcCloseBracketPos As Integer 'переменная для хранения позиции в строке ), закрывающей функцию

        'В цикле из strExpression получаем слова, получаем их значения (результаты функций, значения свойств, переменных...) и помещаем в результирующий массив arrResult
        'В итоге arrResult может содержать только числа, строки, True/False, операторы и скобки
        Do While Not currentPosition >= strExpression.Length
            If EXIT_CODE Then Return ""

            'Если в начале скобки - сохраняем их в arrResult
            Do While strExpression.Chars(currentPosition) = "("c
                openBracketsCount += 1 'Увеливаем число открывающих скобок
                'Расширяем массив arrResult
                arrResultLength += 1
                ReDim Preserve arrResult(arrResultLength)
                arrResult(arrResultLength) = "(" 'Сохраняем в нем скобку
                currentPosition += 1 'Перемещаем текущую позицию в strExpression на 1 вперед
            Loop

            savePosition = currentPosition 'Сохраним текущее положение

            'Обработка строки (строка начинается с одинарной кавычки)
            If strExpression.Chars(currentPosition) = "'"c Then
                'Начинается строка
closingQuoteSearching:
                'Ищем закрывающую кавычку
                currentPosition = strExpression.IndexOf("'"c, currentPosition + 1)
                If currentPosition = -1 Then 'Не найдена - ошибка
                    LAST_ERROR = "Не найдена закрывающая кавычка стороки."
                    Return "#Error"
                End If
                'Экранированная кавычка - ищем дальше
                If strExpression.Chars(currentPosition - 1) = "/"c Then GoTo closingQuoteSearching
                'Получаем всю строку
                strWord = strExpression.Substring(savePosition, currentPosition - savePosition + 1)
                blnStrings = True 'В выражении есть строки
                'Расширяем массив arrResult
                arrResultLength += 1
                ReDim Preserve arrResult(arrResultLength)
                arrResult(arrResultLength) = strWord 'Сохраняем в нем строку
                currentPosition += 1 'Перемещаем текущую позицию в strExpression на 1 вперед
                'Если в конце закрывающие скобки - сохраняем их в arrResult
                Do While currentPosition < strExpression.Length AndAlso strExpression.Chars(currentPosition) = ")"c
                    closeBracketsCount += 1 'Увеливаем число закрывающих скобок
                    'Расширяем массив arrResult
                    arrResultLength += 1
                    ReDim Preserve arrResult(arrResultLength)
                    arrResult(arrResultLength) = ")" 'Сохраняем в нем скобку
                    currentPosition += 1 'Перемещаем текущую позицию в strExpression на 1 вперед
                Loop
                currentPosition += 1 'Перемещаем текущую позицию в strExpression на 1 вперед (т. к. впереди пробел (или конец выражения - тогда это несущественно))
                Continue Do 'Новый виток цикла
            End If

            'Находим пробел
            currentPosition = HideQBcontent(strExpression).IndexOf(" ", currentPosition + 1)
            'Получаем новое слово
            If currentPosition = -1 Then
                'Пробел не найден - конец выражения
                strWord = strExpression.Substring(savePosition) 'Получаем последнее слово в выражении
                currentPosition = strExpression.Length 'Текущая позиция- конец строки
            Else
                'Получаем очередное слово в выражении
                strWord = strExpression.Substring(savePosition, currentPosition - savePosition)
            End If

            savePosition = currentPosition 'Сохраним текущее положение
            'Определяем, что это за слово (число, оператор, функция...) и в некоторых случаях его результат
            currentWordType = GetWordType(strWord, arrParams, currentClassId, elementName, elementContent, funcOpenBracketPos, newCloseBracketsCount, returnFormat)

            Select Case currentWordType
                Case MatewScript.WordTypeEnum.W_OPERATOR_MATH
                    blnOperatorsMath = True 'В выражении есть математический оператор
                Case MatewScript.WordTypeEnum.W_OPERATOR_COMPARE
                    blnOperatorsCompare = True 'В выражении есть оператор сравнения
                Case MatewScript.WordTypeEnum.W_OPERATOR_LOGIC
                    blnOperatorsLogic = True 'В выражении есть логический оператор сравнения
                Case MatewScript.WordTypeEnum.W_OPERATOR_STRINGS_MERGER
                    blnOperatorStringsMerger = True 'В выражении есть оператор сложения строк
                Case MatewScript.WordTypeEnum.W_SIMPLE_BOOL
                    blnBooleans = True 'В выражении есть True/False
                Case MatewScript.WordTypeEnum.W_SIMPLE_NUMBER
                    blnNumbers = True 'В выражении есть число
                Case MatewScript.WordTypeEnum.W_PARAM_COUNT
                    'ParamCount
                    'Конвертируем результат (на случай, если был указан модификатор # или $)
                    strWord = ConvertElement(elementContent, returnFormat, MatewScript.ReturnFormatEnum.TO_NUMBER)
                    If strWord = "#Error" Then Return strWord
                    Select Case returnFormat
                        Case MatewScript.ReturnFormatEnum.TO_NUMBER
                            blnNumbers = True 'В выражении есть число
                        Case MatewScript.ReturnFormatEnum.TO_STRING
                            blnStrings = True 'В выражении есть строка
                        Case MatewScript.ReturnFormatEnum.TO_BOOL
                            blnBooleans = True 'В выражении есть True/False
                        Case ReturnFormatEnum.TO_CODE
                            blnCode = True
                    End Select
                Case MatewScript.WordTypeEnum.W_PARAM, MatewScript.WordTypeEnum.W_VARIABLE_LOCAL, MatewScript.WordTypeEnum.W_VARIABLE_PUBLIC
                    'Param[x] или переменная. В этом случае elementContent уже содержит ее результат
                    'Конвертируем результат (на случай, если был указан модификатор # или $)
                    strWord = ConvertElement(elementContent, returnFormat)
                    If strWord = "#Error" Then Return strWord
                    Select Case returnFormat
                        Case MatewScript.ReturnFormatEnum.TO_NUMBER
                            blnNumbers = True 'В выражении есть число
                            strWord = strWord.Replace(","c, "."c)
                        Case MatewScript.ReturnFormatEnum.TO_STRING
                            blnStrings = True 'В выражении есть строка
                        Case MatewScript.ReturnFormatEnum.TO_BOOL
                            blnBooleans = True 'В выражении есть True/False
                        Case ReturnFormatEnum.TO_CODE
                            blnCode = True
                    End Select
                Case MatewScript.WordTypeEnum.W_FUNCTION
                    'Функция
                    'Получаем параметры функции и положение ее закрывающей скобки
                    currentPosition = GetFunctionParams(strExpression, funcParams, IIf(funcOpenBracketPos = -1, -1, currentPosition - strWord.Length + funcOpenBracketPos), funcCloseBracketPos, elementContent, newCloseBracketsCount, arrParams)
                    If currentPosition = -2 Then Return "#Error"
                    If currentPosition < 0 Then
                        'Функция без скобок - возвращаем в currentPosition положение пробела вслед за именем функции
                        currentPosition = savePosition
                        'сохраняем кол-во ) в конце функции, если они есть (т. к. GetFunctionParams их не считает, когда у функции нет собственных скобок)
                        Do While strExpression.Chars(currentPosition - 1 - newCloseBracketsCount) = ")"c
                            newCloseBracketsCount += 1
                        Loop
                    Else
                        currentPosition += 1 + newCloseBracketsCount 'Перемещаем текущую позицию в strExpression вперед на 1 + количество закрывающих скобок за функцией
                    End If
                    'Передаем управление FunctionRouter() и получаем результат
                    strWord = FunctionRouter(currentClassId, elementName, funcParams, arrParams)
                    If strWord = "#Error" Then Return "#Error"
                    'Конвертируем результат (на случай, если был указан модификатор # или $)
                    ConvertElement(strWord, returnFormat)
                    Select Case Param_GetType(strWord)
                        Case MatewScript.ReturnFormatEnum.TO_NUMBER
                            blnNumbers = True 'В выражении есть число
                        Case MatewScript.ReturnFormatEnum.TO_STRING
                            blnStrings = True 'В выражении есть строка
                        Case MatewScript.ReturnFormatEnum.TO_BOOL
                            blnBooleans = True 'В выражении есть True/False
                        Case ReturnFormatEnum.TO_CODE
                            blnCode = True
                    End Select
                Case MatewScript.WordTypeEnum.W_PROPERTY
                    'Свойство
                    'Получаем параметры свойства и положение его закрывающей скобки
                    currentPosition = GetFunctionParams(strExpression, funcParams, IIf(funcOpenBracketPos = -1, -1, currentPosition - strWord.Length + funcOpenBracketPos), funcCloseBracketPos, elementContent, newCloseBracketsCount, arrParams)
                    If currentPosition = -2 Then Return "#Error"
                    If currentPosition < 0 Then
                        'Свойство без скобок - возвращаем в currentPosition положение пробела вслед за именем свойства
                        currentPosition = savePosition
                        'сохраняем кол-во ) в конце свойства, если они есть (т. к. GetFunctionParams их не считает, когда у свойства нет собственных скобок)
                        Do While strExpression.Chars(currentPosition - 1 - newCloseBracketsCount) = ")"c
                            newCloseBracketsCount += 1
                        Loop
                    Else
                        currentPosition += 1 + newCloseBracketsCount 'Перемещаем текущую позицию в strExpression вперед на 1 + количество закрывающих скобок за свойством
                    End If
                    'Передаем управление PropertiesRouter() и получаем результат
                    strWord = PropertiesRouter(currentClassId, elementName, funcParams, arrParams)
                    If strWord = "#Error" Then Return "#Error"
                    'Конвертируем результат (на случай, если был указан модификатор # или $)
                    ConvertElement(strWord, returnFormat)
                    Select Case Param_GetType(strWord)
                        Case MatewScript.ReturnFormatEnum.TO_NUMBER
                            blnNumbers = True 'В выражении есть число
                        Case MatewScript.ReturnFormatEnum.TO_STRING
                            blnStrings = True 'В выражении есть строка
                        Case MatewScript.ReturnFormatEnum.TO_BOOL
                            blnBooleans = True 'В выражении есть True/False
                        Case ReturnFormatEnum.TO_CODE
                            blnCode = True
                    End Select
                Case MatewScript.WordTypeEnum.W_NOTHING
                    Return "#Error"
            End Select

            'передан массив?
            If strWord = "#ARRAY" Then
                If Not fullExecution AndAlso currentPosition >= strExpression.Length AndAlso arrResultLength = -1 Then
                    Return strWord
                ElseIf blnOperatorsMath OrElse blnOperatorsLogic OrElse blnOperatorStringsMerger Then
                    Return _ERROR("К массивам можно применять только операторы присваивания.")
                Else
                    strWord = CreateTempArray()
                    blnStrings = True
                    If strWord = "#Error" Then Return strWord
                End If
            ElseIf strWord = "" Then
                'empty
                If blnOperatorsMath OrElse blnOperatorsLogic OrElse blnOperatorStringsMerger OrElse currentPosition < strExpression.Length OrElse arrResultLength <> -1 Then
                    Return _ERROR("Данная операция с пустым значением недопустима.")
                End If
                Return ""
            End If

            'Убираем конечные скобки в текущем слове (если они есть)
            If newCloseBracketsCount > 0 AndAlso (currentWordType = MatewScript.WordTypeEnum.W_SIMPLE_NUMBER OrElse currentWordType = MatewScript.WordTypeEnum.W_SIMPLE_BOOL) Then
                If strWord.EndsWith(StrDup(newCloseBracketsCount, ")")) Then strWord = strWord.Remove(strWord.Length - newCloseBracketsCount)
            End If
            'Расширяем массив arrResult
            arrResultLength += 1
            ReDim Preserve arrResult(arrResultLength)
            arrResult(arrResultLength) = strWord 'Сохраняем результат

            If newCloseBracketsCount > 0 Then
                'Если есть закрывающие скобки - также сохраняем их в arrResult
                ReDim Preserve arrResult(arrResultLength + newCloseBracketsCount) 'Расширяем массив arrResult
                For i As Integer = arrResultLength + 1 To arrResultLength + newCloseBracketsCount
                    closeBracketsCount += 1 'Количество закрывающих скобок (всего, в обработанной части выражения strExpression)
                    arrResult(i) = ")" 'Сохраняем скобку в arrResult
                Next
                arrResultLength += newCloseBracketsCount
            End If
            currentPosition += 1 'Перемещаем текущую позицию в strExpression на 1 вперед            
        Loop 'Новый виток цикла

        'Массив arrResult, состоящий только из чисел, строк, True/False, операторов и скобок собран
        'Теперь надо получить окончательный результат получившегося выражения

        Dim typesCount As Integer 'Количество типов в выражении
        'Количество типов надо для подстановки в функцию LogicCompare(). Если тип только один, то эта функция выполнится немного быстрее
        If blnNumbers Then
            typesCount += 1 'количество типов + 1 (числа)
            returnFormat = MatewScript.ReturnFormatEnum.TO_NUMBER 'Передадим в LogicCompare(), что у нас только числа
        End If
        If blnBooleans Then
            typesCount += 1 'количество типов + 1 (логические)
            returnFormat = MatewScript.ReturnFormatEnum.TO_BOOL 'Передадим в LogicCompare(), что у нас только True/False
        End If
        If blnStrings Then
            typesCount += 1 'количество типов + 1 (строки)
            returnFormat = MatewScript.ReturnFormatEnum.TO_STRING 'Передадим в LogicCompare(), что у нас только строки
        End If
        If blnCode Then
            If typesCount > 0 OrElse arrResultLength > 0 Then
                LAST_ERROR = "Нельзя выполнять логические и математические операции с кодом."
                Return "#Error"
            End If
            Return arrResult(0)
        End If

        If typesCount = 0 Then Return "" 'Типов нет вообще - значит массив пустой - возвращаем пустую строку
        If typesCount > 1 Then returnFormat = MatewScript.ReturnFormatEnum.ORIGINAL 'Типов несколько - значит LogicCompare() будет определять типы самостоятельно

        If arrResultLength = 0 Then Return arrResult(0) 'Массив из одного элемента - значит это и есть готовый результат

        If openBracketsCount <> closeBracketsCount Then
            'Ошибка: количество открывающих и закрывающих скобок не совпадает!
            LAST_ERROR = "Количество открывающих и закрывающих скобок не совпадает."
            Return "#Error"
        End If

        If blnOperatorsCompare = False AndAlso blnOperatorsLogic = False Then
            'Все операторы исключительно математические (или сложения строк)
            If typesCount > 1 Then
                LAST_ERROR = "Несоответствие типов данных."
                Return "#Error"
            End If
            If blnNumbers Then
                'У нас числа - считаем результат уравнения
                Return Calculate(arrResult, openBracketsCount)
            ElseIf blnBooleans Then
                'У нас логические превращаем True в 1 и False в 0
                For i = 0 To arrResultLength Step 2
                    arrResult(i) = IIf(arrResult(i) = "True", "1", "0")
                Next
                'Считаем результат уравнения. Если получили 0 - возвращаем False; иначе - True
                Return IIf(Calculate(arrResult, openBracketsCount) = 0, "False", "True")
            Else
                'У нас строки - выполняем сложение строк
                Dim strResult As String = "'" 'Вставляем начальную кавычку
                For i = 0 To arrResultLength Step 2
                    strResult += arrResult(i).Substring(1, arrResult(i).Length - 2)
                Next
                Return strResult + "'" 'Возвращаем результат, добавив в конце обратную кавычку
            End If
        Else
            'Выполняем логическое сравнение и возвращаем результат
            Return LogicCompare(arrResult, returnFormat, openBracketsCount, blnOperatorsMath Or blnOperatorStringsMerger, blnOperatorsCompare, blnOperatorsLogic)
        End If
    End Function

    ''' <summary>Функция выполняет математические операции. В качестве параметра masExpression() передается массив, в котором отдельно хранятся скобки, числа и математические операторы.
    ''' Проверка корректности массива выполняется на предыдущих этапах
    ''' </summary>
    ''' <param name="masExpression">массив с числами и математическими знаками</param>
    ''' <param name="bracketsCount">количество скобок</param>
    ''' <returns>результат вычисления</returns>
    Public Function Calculate(ByVal masExpression() As String, Optional ByVal bracketsCount As Integer = 0) As String
        On Error GoTo er

        '1********************************************************************************
        'Вычисляем значения в скобках и подставляем результат
        Dim i, j As Integer 'счетчик
        If bracketsCount > 0 Then
            'Скобки есть
            Dim brOpenPos, brClosePos As Integer 'для хранения положения последних найденных открывающей и закрывающей скобок
            Dim brSubArray() As String 'для передачи части исходного массива, который находится между скобками
            For i = 1 To bracketsCount
                brOpenPos = Array.LastIndexOf(masExpression, "(") 'Последняя открывающая скобка в массиве
                brClosePos = Array.IndexOf(masExpression, ")", brOpenPos + 1) 'Первая закрывающая, следующая за ней
                'Создаем массив, размером, равным содержимому между скобками
                ReDim brSubArray(brClosePos - brOpenPos - 2)
                'Копируем фрагмент массива между скобками
                Array.Copy(masExpression, brOpenPos + 1, brSubArray, 0, brClosePos - brOpenPos - 1)
                'Очищаем скопированный фрагмент
                For j = brOpenPos + 1 To brClosePos
                    masExpression(j) = ""
                Next
                'На место открывающей скобки помещаем результат всего выражения в скобках
                If IsNothing(brSubArray) Then
                    masExpression(brOpenPos) = "0"
                Else
                    masExpression(brOpenPos) = Calculate(brSubArray)
                    If masExpression(brOpenPos) = "#Error" Then Return "#Error"
                End If
            Next
        End If

        '2********************************************************************************
        'Константы, используемые для наглядности (определяют выполняемый математический оператор).
        Static OP_POWER As Byte = 1, OP_MULTI As Byte = 2, OP_DIV As Byte = 3, OP_SUM As Byte = 4, OP_MIN As Byte = 5

        'Идея: проходим все знаки в поиске приоритетного (напр., * приоритетнее, чем +). Затем выполняем математическую операцию, 
        'помещаем результат в левую ячейку массива, обнуляем использованные знак и числа и повторяем цикл. Так до тех пор, пока не обнулятся все знаки
        'В результате в первой ячейке массива masExpression получаем результат (остальные ячейки обнуляются)
        Dim strSign As String 'Текущий знак
        Do
            Dim usingSign As Byte = 0, usingSignPos As Integer = 0 'Знак, выбранный приоритетным и его позиция в строке
            For i = 1 To UBound(masExpression) Step 2 'Проходим все знаки
                strSign = masExpression(i) 'Получаем текущий знак
                If Len(strSign) = 0 Then Continue For 'Если знак здесь обнулен - продолжаем цикл
                Select Case strSign
                    Case "^" 'Текущий знак - возведение в степень (приоритет № 1 - наивысший)
                        'Возведение в степень - всегда приоритетный
                        usingSign = OP_POWER 'Устанавливаем действие
                        usingSignPos = i
                        Exit For
                    Case "+", "-" 'сложение или вычитание (приоритет № 3 - наименьший)
                        If usingSign = 0 Then
                            'Сюда попадаем, если знак еще не был установлен
                            usingSign = IIf(strSign = "+", OP_SUM, OP_MIN)
                            usingSignPos = i
                        End If
                    Case Else 'умножение или деление (приоритет № 2)
                        If usingSign >= 4 OrElse usingSign = 0 Then
                            'Сюда попадаем если предыдущий знак + или -, или же еще не был установлен
                            usingSign = IIf(strSign = "*", OP_MULTI, OP_DIV)
                            usingSignPos = i
                        End If
                End Select
            Next

            If usingSignPos = 0 Then Return masExpression(0) 'Не найден ни один знак - значит все операторы выполнены
            'Получаем ближайшее слева необнуленное число. К нему будет применено выбранное (usingSign) математическое действие
            For i = usingSignPos - 1 To 0 Step -2
                If Len(masExpression(i)) > 0 Then Exit For
            Next
            'Выполняем математическое действие, помещая результат в левую ячейку
            '"а = а + в" работает почему-то быстрее, чем "а += в"
            Select Case usingSign
                Case OP_POWER
                    masExpression(i) = Val(masExpression(i)) ^ Val(masExpression(usingSignPos + 1))
                Case OP_MULTI
                    masExpression(i) = Val(masExpression(i)) * Val(masExpression(usingSignPos + 1))
                Case OP_DIV
                    masExpression(i) = Val(masExpression(i)) / Val(masExpression(usingSignPos + 1))
                Case OP_SUM
                    masExpression(i) = Val(masExpression(i)) + Val(masExpression(usingSignPos + 1))
                Case OP_MIN
                    masExpression(i) = Val(masExpression(i)) - Val(masExpression(usingSignPos + 1))
            End Select
            masExpression(i) = masExpression(i).Replace(",", ".")
            'Обнуляем использованный оператор
            masExpression(usingSignPos) = ""
            masExpression(usingSignPos + 1) = ""
        Loop

        Return masExpression(0)
er:
        LAST_ERROR = "Ошибка при математических расчетах."
        Return "#Error"
        'MessageBox.Show("Ошибка при математических расчетах", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
    End Function

    ''' <summary>Функция выполняет операции математического и логического сравнения (а также математические выражения, используя Calculate() ).
    ''' Возвращает "True"/"False" или число/строку, если не было операторов сравнения.
    ''' </summary>
    ''' <param name="masExpression">массив, в котором все значения, операторы и скобки являются отдельными элементами</param>
    ''' <param name="varType">тип значений в массиве (число, строка, логическое). Если равно ORIGINAL, то тип определяется внутри этой функции</param>
    ''' <param name="bracketsCount">количество скобок. Важно только их 0 или больше. Если 0 - проверка на скобки не происходит (для ускорения работы функции)
    ''' Кратность скобок определяется ранее и здесь не происходит.</param>
    ''' <param name="isMathOperators">есть ли в переданном массиве математические операторы.
    ''' Если False - проверка не происходит (для ускорения)</param>
    ''' <param name="isComparisonOperator">есть ли в переданном массиве операторы математического
    ''' сравнения. Если False - проверка не происходит (для ускорения)</param>
    ''' <param name="isLogicOperators">есть ли в переданном массиве операторы логического сравнения. 
    ''' Если False - проверка не происходит (для ускорения)</param>
    ''' <returns>Результат сравнения - True / False</returns>
    Public Function LogicCompare(ByVal masExpression() As String, ByVal varType As MatewScript.ReturnFormatEnum, Optional ByVal bracketsCount As Integer = 0,
                                 Optional ByVal isMathOperators As Boolean = True, Optional ByVal isComparisonOperator As Boolean = True,
                                 Optional ByVal isLogicOperators As Boolean = True) As String

        'isMathOperators, isComparisonOperator, isLogicOperators - есть ли в переданном массиве операторы: математические, математического
        'сравнения и логического сравнения. Если что-либо False - проверка не происходит (для ускорения)

        On Error GoTo er
        'Если массив пустой - ошибка
        If IsNothing(masExpression) Then Return "#Error"

        '1********************************************************************************
        'Получаем результат выражений в скобках и вставляем в исходный массив

        Dim subArray() As String = {} 'для передачи части исходного массива, который находится между скобками
        Dim i, j As Integer 'счетчики
        If bracketsCount > 0 Then
            'Скобки есть
            Dim brOpenPos, brClosePos As Integer 'для хранения положения последних найденных открывающей и закрывающей скобок
            brOpenPos = Array.LastIndexOf(masExpression, "(") 'Последняя открывающая скобка в массиве
            Do While brOpenPos > -1 'Работает, пока скобки не закончатся
                brClosePos = Array.IndexOf(masExpression, ")", brOpenPos + 1) 'Первая закрывающая скобка, следующая за найденной открывающей
                'В конце будущего массива с содержимым скобок, ни в коем случае не должно быть пустых ячеек в конце.
                j = 0
                Do While masExpression(brClosePos - j - 1).Length = 0
                    'Две ячейки перед закрывающей скобкой пустые
                    j += 2
                Loop
                'Теперь j = количеству пустых ячеек перед закрывающей скобкой
                'Создаем массив, размером, равным содержимому между скобками (минус пустые ячейки)
                ReDim subArray(brClosePos - brOpenPos - 2 - j)
                'Копируем фрагмент массива между скобками
                Array.Copy(masExpression, brOpenPos + 1, subArray, 0, brClosePos - brOpenPos - 1 - j)
                'Очищаем скопированный фрагмент
                For j = brOpenPos + 1 To brClosePos
                    masExpression(j) = ""
                Next
                'На место открывающей скобки помещаем результат всего выражения в скобках
                If IsNothing(subArray) Then
                    masExpression(brOpenPos) = "False"
                Else
                    masExpression(brOpenPos) = LogicCompare(subArray, varType, 0, isMathOperators, isComparisonOperator, isLogicOperators)
                    If masExpression(brOpenPos) = "#Error" Then Return "#Error"
                End If
                brOpenPos = Array.LastIndexOf(masExpression, "(") 'следующая последняя открывающая скобка в массиве
            Loop

            'Если конец исходного массива оказался пустым - обрезаем его (обязательно для нормальной работы функции)
            j = 0
            For i = masExpression.GetUpperBound(0) To 0 Step -2
                If masExpression(i).Length = 0 Then
                    'Две последние ячейки массива пустые
                    j += 2
                Else
                    'Ячейки содержат информацию - не обрезать
                    Exit For
                End If
            Next
            'Удаляем конечные пустые ячейки
            If j > 0 Then Array.Resize(masExpression, masExpression.Length - j)
            If masExpression.Length = 1 Then Return masExpression(0) 'Не осталось операторов - результат готов
        End If

        '2********************************************************************************
        'Выполняем математические операторы
        Dim actBegin As Integer = -1 'Хранит положение первого значения, учавствующего в операции
        Dim actEnd As Integer = -1 'Хранит положение последнего оператора, учавствующего в операции
        Dim blnDoAction As Boolean 'Пора ли проводить операцию, или еще не все значения получены
        Dim currentType As MatewScript.ReturnFormatEnum 'Формат значений (строка, число...) - если явно не указано при вызове функции

        If isMathOperators Then
            'Математические операторы есть
            'Выполняем все математические блоки в массиве и результат помещаем в первую (левую) задействованную ячейку
            'В результате получаем скорректированный массив уже без математических операторов, в которых результаты расчетов
            'находятся между операторами сравнения и логическими.
            Static arrOperatorMath() As String = {"-", "&", "*", "/", "\", "^", "+"} 'Массив из математических операторов, отсортированных повозрастанию
            For i = 1 To masExpression.GetUpperBound(0) - 1 Step 2 'Проходим в цикле операторы
                If masExpression(i).Length = 0 Then Continue For 'строка пустая
                blnDoAction = (actBegin > -1) 'Если начало уже найдено, то, возможно, будем выполнять действие
                If Array.BinarySearch(arrOperatorMath, masExpression(i)) > -1 Then 'Оператор математический - подходит
                    actEnd = i 'меняем положение конечного оператора каждый раз, когда находим новый
                    If actBegin = -1 Then actBegin = i - 1 'Если начало еще не найдено - запоминаем
                    blnDoAction = False 'Подходящие операторы могут еще быть - пока не выполняем действие
                End If
                If i = masExpression.GetUpperBound(0) - 1 Then blnDoAction = True 'Если уже последний оператор в строке - пора действовать
                If blnDoAction AndAlso actBegin > -1 Then 'Если actBegin = -1 - нужного оператора так и не нашли
                    'Выполняем действие
                    'Корректируем положение первого значения на случай, если оно пустое
                    '(движемся к началу массива, пока не найдем первую непустую ячейку)
                    Do While masExpression(actBegin).Length = 0
                        actBegin = actBegin - 2
                    Loop
                    If varType = MatewScript.ReturnFormatEnum.ORIGINAL Then 'Тип значений явно не указан
                        currentType = Param_GetType(masExpression(actBegin)) 'Определяем тип первого значения
                        For j = actBegin To actEnd - 1 Step 2
                            'Все остальные значения обязаны быть такими же, как первое
                            If masExpression(j).Length > 0 AndAlso Param_GetType(masExpression(j)) <> currentType Then Return "#Error"
                        Next
                    Else
                        currentType = varType 'Работаем с явно указанным типом без проверок
                    End If
                    'Вычисляем результат
                    If currentType <> MatewScript.ReturnFormatEnum.TO_STRING Then
                        'Для чисел и логических надо создать подмассив для передачи функции Calculate()
                        'Сложение строк проводится прямо в этой функции прямо в массиве masExpression() без создания подмассива
                        'Создаем массив, размером, равным содержимому 
                        ReDim subArray(actEnd - actBegin + 1)
                        'Копируем фрагмент массива между скобками
                        Array.Copy(masExpression, actBegin, subArray, 0, actEnd - actBegin + 2)
                    End If
                    'Выполняем математические операции
                    If IsNothing(subArray) AndAlso currentType <> MatewScript.ReturnFormatEnum.TO_STRING Then
                        masExpression(actBegin) = IIf(currentType = MatewScript.ReturnFormatEnum.TO_NUMBER, "0", "False")
                    Else
                        Select Case currentType
                            Case MatewScript.ReturnFormatEnum.TO_NUMBER
                                masExpression(actBegin) = IIf(IsNothing(subArray), "0", Calculate(subArray)) 'Вызов Calculate()
                                If masExpression(actBegin) = "#Error" Then Return "#Error"
                            Case MatewScript.ReturnFormatEnum.TO_STRING
                                'Выполняем сложение строк
                                'Убираем последнюю кавычку ' в первой строке
                                masExpression(actBegin) = masExpression(actBegin).Remove(masExpression(actBegin).Length - 1)
                                For j = actBegin + 2 To actEnd + 1 Step 2
                                    'Добавляем все последующие строки к первой, убрав в них кавычки ' '
                                    masExpression(actBegin) += masExpression(j).Substring(1, masExpression(j).Length - 2)
                                Next
                                'В конец полученной строки дописываем кавычку ' - готовый результат
                                masExpression(actBegin) += "'"
                            Case MatewScript.ReturnFormatEnum.TO_BOOL
                                'Преобразуем True и False в их числовые эквиваленты и запускаем Calculate()
                                If IsNothing(subArray) Then
                                    masExpression(actBegin) = "False"
                                Else
                                    For j = 0 To subArray.GetUpperBound(0) Step 2
                                        subArray(j) = Convert.ToString(Val(Convert.ToBoolean(subArray(j))))
                                    Next
                                    masExpression(actBegin) = IIf(Calculate(subArray) = 0, "False", "True")
                                    If masExpression(actBegin) = "#Error" Then Return "#Error"
                                End If
                        End Select
                    End If
                    'Очищаем скопированный фрагмент
                    For j = actBegin + 1 To actEnd + 1
                        masExpression(j) = ""
                    Next
                    'Обнуляем положение первого задействованного элемента и продолжаем цикл в поиске нового математического блока
                    actBegin = -1
                End If
            Next

            'Если конец исходного массива оказался пустым - обрезаем его (обязательно для нормальной работы функции)
            j = 0
            For i = masExpression.GetUpperBound(0) To 0 Step -2
                If masExpression(i).Length = 0 Then
                    'Две последние ячейки массива пустые
                    j += 2
                Else
                    'Ячейки содержат информацию - не обрезать
                    Exit For
                End If
            Next
            'Удаляем конечные пустые ячейки
            If j > 0 Then Array.Resize(masExpression, masExpression.Length - j)
            If masExpression.Length = 1 Then Return masExpression(0) 'Не осталось операторов - результат готов
        End If

        '3********************************************************************************
        'Выполняем операторы сравнения
        actBegin = -1 'Хранит положение первого значения, учавствующего в операции
        actEnd = -1 'Хранит положение последнего оператора, учавствующего в операции
        Dim blnResult As Boolean 'Хранит результат операции сравнения
        Dim intElementValue As Double, blnElementValue As Boolean, strElementValue As String 'Массив из операторов сравнения, отсортированных повозрастанию
        If isComparisonOperator Then
            'Операторы сравнения есть
            'Выполняем все блоки сравнения в массиве и результат помещаем в первую (левую) задействованную ячейку
            'В результате получаем скорректированный массив уже без математических операторов и операторов сравнения, в которых результаты расчетов
            'находятся между логическими операторами.
            Static arrOperatorCompare() As String = {"!=", "<", "<=", "<>", "=", ">", ">="}
            For i = 1 To masExpression.GetUpperBound(0) - 1 Step 2 'Проходим в цикле операторы
                If masExpression(i).Length = 0 Then Continue For 'строка пустая
                blnDoAction = (actBegin > -1) 'Если начало уже найдено, то, возможно, будем выполнять действие
                If Array.BinarySearch(arrOperatorCompare, masExpression(i)) > -1 Then 'Оператор сравнения - подходит
                    actEnd = i 'меняем положение конечного оператора каждый раз, когда находим новый
                    If actBegin = -1 Then actBegin = i - 1 'Если начало еще не найдено - запоминаем
                    blnDoAction = False 'Подходящие операторы могут еще быть - пока не выполняем действие
                End If
                If i = masExpression.GetUpperBound(0) - 1 Then blnDoAction = True 'Если уже последний оператор в строке - пора действовать
                If blnDoAction AndAlso actBegin > -1 Then 'Если actBegin = -1 - нужного оператора так и не нашли
                    'Выполняем действие
                    'Корректируем положение первого значения на случай, если оно пустое
                    '(движемся к началу массива, пока не найдем первую непустую ячейку)
                    Do While masExpression(actBegin).Length = 0
                        actBegin = actBegin - 2
                    Loop
                    If varType = MatewScript.ReturnFormatEnum.ORIGINAL Then 'Тип значений явно не указан
                        currentType = Param_GetType(masExpression(actBegin)) 'Определяем тип первого значения
                        For j = actBegin To actEnd - 1 Step 2
                            'Все остальные значения обязаны быть такими же, как первое
                            If masExpression(j).Length > 0 AndAlso Param_GetType(masExpression(j)) <> currentType Then Return "#Error"
                        Next
                    Else
                        currentType = varType 'Работаем с явно указанным типом без проверок
                    End If
                    'Вычисляем результат
                    blnResult = True
                    Select Case currentType
                        Case MatewScript.ReturnFormatEnum.TO_STRING
                            'Работаем со строками
                            strElementValue = masExpression(actBegin) 'Первое значение в блоке сравнения
                            'Сравниваем все последующие значения в блоке с первым. Если хоть раз получили False - значит False, если нет - True
                            For j = actBegin + 2 To actEnd + 1 Step 2
                                Select Case masExpression(j - 1)
                                    Case "="
                                        blnResult = strElementValue = masExpression(j)
                                    Case "<>", "!="
                                        blnResult = strElementValue <> masExpression(j)
                                    Case ">"
                                        blnResult = strElementValue > masExpression(j)
                                    Case "<"
                                        blnResult = strElementValue < masExpression(j)
                                    Case ">="
                                        blnResult = strElementValue >= masExpression(j)
                                    Case "<="
                                        blnResult = strElementValue <= masExpression(j)
                                End Select
                                If blnResult = False Then Exit For
                            Next
                        Case MatewScript.ReturnFormatEnum.TO_NUMBER
                            'Работаем с числами
                            intElementValue = Val(masExpression(actBegin)) 'Первое значение в блоке сравнения
                            'Сравниваем все последующие значения в блоке с первым. Если хоть раз получили False - значит False, если нет - True
                            For j = actBegin + 2 To actEnd + 1 Step 2
                                Select Case masExpression(j - 1)
                                    Case "="
                                        blnResult = intElementValue = Val(masExpression(j))
                                    Case ">"
                                        blnResult = intElementValue > Val(masExpression(j))
                                    Case "<"
                                        blnResult = intElementValue < Val(masExpression(j))
                                    Case ">="
                                        blnResult = intElementValue >= Val(masExpression(j))
                                    Case "<="
                                        blnResult = intElementValue <= Val(masExpression(j))
                                    Case "<>", "!="
                                        blnResult = intElementValue <> Val(masExpression(j))
                                End Select
                                If blnResult = False Then Exit For
                            Next
                        Case MatewScript.ReturnFormatEnum.TO_BOOL
                            'Работаем с логическими
                            blnElementValue = Convert.ToBoolean(masExpression(actBegin)) 'Первое значение в блоке сравнения
                            'Сравниваем все последующие значения в блоке с первым. Если хоть раз получили False - значит False, если нет - True
                            For j = actBegin + 2 To actEnd + 1 Step 2
                                Select Case masExpression(j - 1)
                                    Case "="
                                        blnResult = blnElementValue = Convert.ToBoolean(masExpression(j))
                                    Case ">"
                                        blnResult = blnElementValue > Convert.ToBoolean(masExpression(j))
                                    Case "<"
                                        blnResult = blnElementValue < Convert.ToBoolean(masExpression(j))
                                    Case ">="
                                        blnResult = blnElementValue >= Convert.ToBoolean(masExpression(j))
                                    Case "<="
                                        blnResult = blnElementValue <= Convert.ToBoolean(masExpression(j))
                                    Case "<>", "!="
                                        blnResult = blnElementValue <> Convert.ToBoolean(masExpression(j))
                                End Select
                                If blnResult = False Then Exit For
                            Next
                    End Select
                    'Очищаем участов массива, который только что просчитали
                    For j = actBegin + 1 To actEnd + 1
                        masExpression(j) = ""
                    Next
                    'В первый элемент выбранного блока помещаем результат сравнения
                    masExpression(actBegin) = IIf(blnResult, "True", "False")
                    'Обнуляем положение первого задействованного элемента и продолжаем цикл в поиске нового блока сравнения
                    actBegin = -1
                End If
            Next

            'Если конец исходного массива оказался пустым - обрезаем его (обязательно для нормальной работы функции)
            j = 0
            For i = masExpression.GetUpperBound(0) To 0 Step -2
                If masExpression(i).Length = 0 Then
                    'Две последние ячейки массива пустые
                    j += 2
                Else
                    'Ячейки содержат информацию - не обрезать
                    Exit For
                End If
            Next
            'Удаляем конечные пустые ячейки
            If j > 0 Then Array.Resize(masExpression, masExpression.Length - j)
            If masExpression.Length = 1 Then Return masExpression(0) 'Не осталось операторов - результат готов
        End If

        '4********************************************************************************
        'Выполняем логические операторы
        If isLogicOperators Then
            'Логические операторы есть
            'В данной строке уже нет никаких операторов, кроме логических. Также все значения должны быть только True / False
            actBegin = 0 'Начинаем работать с первого элемента массива
            Dim blnCurrent As Boolean 'Для хранения значения элементов массива, которые сравниваются с первым
            blnResult = False
            Boolean.TryParse(masExpression(actBegin), blnResult) 'Получаем значение первого элемента, с которым будут сравниваться все последующие
            For i = actBegin + 2 To masExpression.GetUpperBound(0) Step 2
                If masExpression(i - 1).Length > 0 Then
                    Boolean.TryParse(masExpression(i), blnCurrent) 'Получаем значение очередного элемента
                    'Выполняем логическое сравнение его и первого элемента
                    Select Case masExpression(i - 1)
                        Case "And"
                            blnResult = (blnResult AndAlso blnCurrent)
                        Case "Or"
                            blnResult = (blnResult OrElse blnCurrent)
                        Case "Xor"
                            blnResult = (blnResult Xor blnCurrent)
                    End Select
                End If
            Next
            masExpression(0) = Convert.ToString(blnResult) 'Сохраняем результат в первой ячейке рабочего массива
        End If

        'Возвращаем результат функции
        Return masExpression(0)

er:
        LAST_ERROR = "Ошибка в логических вычислениях."
        Return "#Error"
    End Function

    ''' <summary>
    ''' Возвращает Id линни, следующей за указанной меткой
    ''' </summary>
    ''' <param name="masCode">ссылка на полный код</param>
    ''' <param name="markName">имя метки без кавычек</param>
    ''' <returns>Id линии за меткой</returns>
    Private Overloads Function GetLineAfterMark(ByRef masCode As List(Of ExecuteDataType), ByVal markName As String) As Integer
        Dim markPos As Integer = 0
        While markPos < masCode.Count - 1
            If masCode(markPos).Code.Length > 0 AndAlso masCode(markPos).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_MARK _
                AndAlso String.Compare(masCode(markPos).Code(0).Word, markName + ":", True) = 0 Then
                Return markPos + 1
            End If
            markPos += 1
        End While
        Return -1
    End Function

    ''' <summary>
    ''' Возвращает Id линни, следующей за указанной меткой
    ''' </summary>
    ''' <param name="masCode">ссылка на полный код</param>
    ''' <param name="markName">имя метки без кавычек</param>
    ''' <returns>Id линии за меткой</returns>
    Private Overloads Function GetLineAfterMark(ByRef masCode() As DictionaryEntry, ByVal markName As String) As Integer
        Dim markPos As Integer = 0
        markName = markName + ":"
        While markPos < masCode.GetUpperBound(0)
            If String.Compare(masCode(markPos).Value, markName, True) = 0 Then
                Return markPos + 1
            End If
            markPos += 1
        End While

        Return -1
    End Function

    ''' <summary>Полулокальная переменная для PrepareForExternalFunction и RestoreAfterExternalFunction. Хранит текщую линию текущего скрипта</summary>
    Private pe_curLine As Integer
    ''' <summary>Полулокальная переменная для PrepareForExternalFunction и RestoreAfterExternalFunction. Хранит копию локальных переменных</summary>
    Private pe_varsCopy As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
    ''' <summary>
    ''' Подготавливает данные (в частности, делает копию локальных переменных) перед запуском внешнего кода
    ''' </summary>
    ''' <param name="funcName">Название внешнего кода для вставки в стек</param>
    ''' <remarks></remarks>
    Public Sub PrepareForExternalFunction(ByVal funcName As String)
        codeStack.Push(funcName)
        pe_curLine = CURRENT_LINE
        'csLocalVariables.CopyVariables(pe_varsCopy)
        'csLocalVariables.KillVars()
        pe_varsCopy = csLocalVariables.lstVariables
        csLocalVariables.lstVariables = New SortedList(Of String, cVariable.variableEditorInfoType)(StringComparer.CurrentCultureIgnoreCase)
    End Sub

    ''' <summary>
    ''' Восстанавливает переменные и все необходимые данные после завершения внешнего кода. Перед началом исполнения внешнего кода небходимо запустить PrepareForExternalFunction
    ''' </summary>
    ''' <param name="wasError">завершен ли внешний код с ошибкой</param>
    ''' <remarks></remarks>
    Public Sub RestoreAfterExternalFunction(ByVal wasError As Boolean)
        codeStack.Pop()
        CURRENT_LINE = pe_curLine
        'csLocalVariables.RestoreVariables(pe_varsCopy)
        'pe_varsCopy.Clear()
        csLocalVariables.lstVariables.Clear()
        csLocalVariables.lstVariables = pe_varsCopy
        pe_varsCopy = Nothing
        If Not wasError Then EXIT_CODE = False
    End Sub

    ''' <summary>Главная функция, выполняющая весь код, предварительно собранный в список ExecuteDataType.
    ''' Он уже не содержит комментариев, разделения строк, пустых строк, лишних пробелов и т. д.
    ''' </summary>
    ''' <param name="masCode">список, каждый элемент которого - элемент ExecuteDataType</param>
    ''' <param name="arrParams">параметры, переданные при выполнении кода</param>
    ''' <param name="isPrimaryCode">является ли вызов первичным (т. е., не рекурсивный ли это вызов из уже выполняемого кода)</param>
    ''' <param name="startPos">первый элемент - строка кода, которую надо выполнить</param>
    ''' <param name="endPos">последняя строка, которую надо выполнить</param>
    ''' <returns>результат, который возвращет код</returns>
    Public Overridable Function ExecuteCode(ByRef masCode As List(Of ExecuteDataType), ByRef arrParams() As String, _
                                    Optional ByVal isPrimaryCode As Boolean = False, Optional ByVal startPos As Integer = 0, _
                                    Optional ByVal endPos As Integer = -1) As String
        If IsNothing(masCode) Then Return ""
        If isPrimaryCode = True Then
            'Если это не рекурсивный вызов - обнуляем передыдущую инфо об ошибках
            EXIT_CODE = False
            LAST_ERROR = ""
        Else
            If EXIT_CODE = True Then Return ""
        End If

        'ВЫПОЛНЯЕМ ПОСТРОЧНО КОД
        Dim lineId As Long = startPos 'порядковый номер (начиная с 0) выполняемой строки кода
        Dim curExecuteItem As ExecuteDataType, strResult As String = "" 'выполнямый элемент из списка masCode и переменная для хранения результата различных вычислений 
        Dim firstWordType As CodeTextBox.EditWordTypeEnum
        Dim returnFormat As MatewScript.ReturnFormatEnum
        Dim lineStartPos As Integer = 0
        Dim nextPos As Integer = 0
        Dim curLineCode() As CodeTextBox.EditWordType

        'Dim firstWordType As MatewScript.WordTypeEnum 'переменная для получения типа первого слова в строке кода, который получается и ф-и GetFirstWordType
        ''переменные для получения данных функции GetFirstWordType (содержимое квадратных скобок, Id функции/свойства в массиве mainClass,
        ''имя функции/свойства/переменной, оставшаяся часть строки кода после первого слова, формат, в который преобразуется результат вычислений перед присвоением его переменной/свойству
        'Dim qbContent As String = "", currentClassId As Integer, ElementId As Integer, elementName As String = "", stringRemain As String = ""
        'Dim funcParams As String() = {} 'Массив для хранения параметров функции/свойства
        'Dim finalBracketCount As Integer = 0 'нефункциональная переменная, необходимая для вызова GetFunctionParams
        'Dim funcCloseBracketPos As Integer 'для получения позиции ), закрывающей функцию/свойство
        'Dim blockFinalPos As Integer, blockFinalOperator As Integer 'переменные для работы с блоками вроде If ... Then
        If endPos = -1 Then endPos = masCode.Count - 1

        'Главный цикл. Перебираем все строки кода
        Do While lineId <= endPos
            CURRENT_LINE = lineId
            curExecuteItem = masCode(lineId) 'получаем код на исполнение
            If IsNothing(curExecuteItem.Code) OrElse curExecuteItem.Code.Length = 0 Then Continue Do
            'получаем первое слово в коде и другие необходимые параметры

            If curExecuteItem.Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Then
                returnFormat = ReturnFormatEnum.TO_NUMBER
                firstWordType = curExecuteItem.Code(1).wordType
                lineStartPos = 1
            ElseIf curExecuteItem.Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING Then
                returnFormat = ReturnFormatEnum.TO_STRING
                firstWordType = curExecuteItem.Code(1).wordType
                lineStartPos = 1
            Else
                returnFormat = ReturnFormatEnum.ORIGINAL
                firstWordType = curExecuteItem.Code(0).wordType
                lineStartPos = 0
            End If

            Select Case firstWordType  'проверяем, что у нас за первое слово и выполняем код соответственно ему
                Case CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT
                    'Event 'propEventName'[, par1, par2]
                    Dim startLine As Integer = lineId
                    Dim finalLine As Integer = -1
                    Dim curPos As Integer = 1
                    curLineCode = masCode(lineId).Code
                    If curLineCode.Length < 2 Then
                        LAST_ERROR = "Неверный синтаксис блока Event."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    'Получаем индекс последней строки в блоке
                    strResult = BlockEvent(masCode, startLine, finalLine)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If

                    'получаем параметры блока event, первый из которых - имя свойства-события
                    Dim funcParams As List(Of String) = Nothing
                    strResult = GetFunctionParams(curLineCode, 0, funcParams, returnFormat, Nothing, Nothing, arrParams, True)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If

                    'получаем имя свойства-события
                    Dim elementName As String = PrepareStringToPrint(funcParams(0), arrParams)
                    If elementName = "#Error" Then
                        LAST_ERROR = "Неправильная запись блока Event. При получении имени события возникла ошибка: " + LAST_ERROR
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    'событие отслеживания свойства?
                    Dim isTracking As frmMainEditor.trackingcodeEnum = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT
                    If elementName.EndsWith(":changed", StringComparison.CurrentCultureIgnoreCase) Then
                        isTracking = frmMainEditor.trackingcodeEnum.EVENT_AFTER
                        elementName = elementName.Substring(0, elementName.Length - 8)
                    ElseIf elementName.EndsWith(":changing", StringComparison.CurrentCultureIgnoreCase) Then
                        isTracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE
                        elementName = elementName.Substring(0, elementName.Length - 9)
                    End If
                    'получаем класс, в котором находится это свойство-событие
                    Dim currentClassId As Integer = -1
                    Dim cPos As Integer = elementName.IndexOf("."c)
                    If cPos > -1 AndAlso cPos < elementName.Length - 1 Then
                        '"L.EventOnEnter"
                        Dim className As String = elementName.Substring(0, cPos)
                        elementName = elementName.Substring(cPos + 1)
                        If mainClassHash.TryGetValue(className, currentClassId) = False Then
                            LAST_ERROR = "Неправильная запись блока Event. При получении имени события возникла ошибка: " + LAST_ERROR
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    Else
                        '"EventOnEnter"
                        For i As Integer = 0 To mainClass.GetUpperBound(0)
                            If mainClass(i).Properties.TryGetValue(elementName, Nothing) Then
                                currentClassId = i
                                Exit For
                            End If
                        Next
                    End If

                    If currentClassId = -1 OrElse mainClass(currentClassId).Properties.TryGetValue(elementName, Nothing) = False Then
                        LAST_ERROR = "Не найдено событие " + elementName + "."
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    'получаем параметры без имени свойства-события
                    funcParams.RemoveAt(0)

                    'создаем структуру codeBox.CodeData() и помещаем туда код события
                    Dim dt() As CodeTextBox.CodeDataType = Nothing
                    Array.Resize(dt, finalLine - lineId) '+ 2)

                    Dim dLvl As Integer = 0
                    For i As Integer = lineId + 1 To finalLine
                        dLvl = i - lineId - 1
                        dt(dLvl) = New CodeTextBox.CodeDataType
                        Array.Resize(dt(dLvl).Code, masCode(i).Code.Length)
                        Array.ConstrainedCopy(masCode(i).Code, 0, dt(dLvl).Code, 0, masCode(i).Code.Length)
                    Next

                    If isTracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE Then
                        'событие отслеживания свойства
                        trackingProperties.AddPropertyBefore(currentClassId, elementName, dt)
                    ElseIf isTracking = frmMainEditor.trackingcodeEnum.EVENT_AFTER Then
                        'событие отслеживания свойства
                        trackingProperties.AddPropertyAfter(currentClassId, elementName, dt)
                    Else
                        'сохраняем событие в соответствующем свойстве с помощью класса eventRouter
                        If eventRouter.SetPropertyWithEvent(currentClassId, elementName, dt, funcParams.ToArray, "", False, False) = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    End If

                    returnFormat = ReturnFormatEnum.ORIGINAL
                    lineId = finalLine + 2 'устанавливаем текущую позицию за End Event
                    Continue Do
                Case CodeTextBox.EditWordTypeEnum.W_HTML
                    If questEnvironment.EDIT_MODE Then Continue Do
                    'Вставляет текст в текущий HTML-документ.
                    'HtmlDocument [Append] [parentId]
                    Dim startLine As Integer = lineId
                    Dim finalLine As Integer = -1
                    Dim curPos As Integer = 1
                    curLineCode = masCode(lineId).Code
                    'Получаем текст и индекс последней строки в блоке
                    strResult = MakeExec(masCode, startLine, finalLine, arrParams)
                    'strResult = PrepareStringToPrint(strResult, arrParams)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If
                    Dim strText As String = strResult ' "'" + strResult.Replace("'", "/'") + "'"

                    'получаем id элемента, куда вставлять текст (если он указан) и замещать или добалять текст
                    Dim isAppend As Boolean = False
                    Dim containerId As String = ""
                    If curLineCode.Length > 1 Then
                        If curLineCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_HTML Then
                            isAppend = True
                            curPos += 1
                        End If
                        If curPos < curLineCode.Length Then
                            containerId = GetValue(curLineCode, curPos, -1, arrParams)
                            If containerId = "#Error" Then
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return "#Error"
                            End If
                            containerId = PrepareStringToPrint(containerId, arrParams)
                        End If
                    End If

                    'Вставляем текст
                    If isAppend = False And containerId.Length = 0 Then
                        'HTML
                        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                        If IsNothing(hDoc) = False Then
                            Dim hEl As HtmlElement = hDoc.GetElementById("MainConvas")
                            If IsNothing(hEl) = False Then hEl.InnerHtml = ""
                        End If
                        PrintTextToMainWindow("", strText, PrintInsertionEnum.APPEND_NEW_BLOCK, "", PrintDataEnum.TEXT, arrParams)
                    Else
                        If containerId.Length = 0 Then
                            'HTML Append
                            PrintTextToMainWindow("", strText, PrintInsertionEnum.APPEND_NEW_BLOCK, "", PrintDataEnum.TEXT, arrParams)
                        Else
                            'HTML 'elementId' или HTML Append 'elementId'
                            If isAppend Then
                                PrintTextToMainWindow(containerId, strText, PrintInsertionEnum.APPEND, "", PrintDataEnum.TEXT, arrParams)
                            Else
                                PrintTextToMainWindow(containerId, strText, PrintInsertionEnum.REPLACE, "", PrintDataEnum.TEXT, arrParams)
                            End If
                        End If
                    End If

                    returnFormat = ReturnFormatEnum.ORIGINAL
                    lineId = finalLine + 2
                    Continue Do
                Case CodeTextBox.EditWordTypeEnum.W_WRAP
                    'Wrap #Var.$a[4 + 5]
                    'Wrap Class[x,y].Property(x,y)

                    Dim startLine As Integer = lineId
                    Dim finalLine As Integer = -1
                    Dim curPos As Integer = 1
                    curLineCode = masCode(lineId).Code
                    If curLineCode.Length < 2 Then
                        LAST_ERROR = "Неверный синтаксис блока Wrap."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If

                    returnFormat = ReturnFormatEnum.ORIGINAL
                    If curLineCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Then
                        curPos += 1
                        returnFormat = ReturnFormatEnum.TO_NUMBER
                    ElseIf curLineCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING Then
                        curPos += 1
                        returnFormat = ReturnFormatEnum.TO_STRING
                    End If

                    If curLineCode.GetUpperBound(0) < curPos Then
                        LAST_ERROR = "Ошибка в синтаксисе блока Wrap. Не указана переменная."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If

                    If curLineCode(curPos).classId >= 0 Then
                        'property changing
                        Dim fParams As New List(Of String), funcPos As Integer = -1, nextEl As Integer = -1

                        If GetFunctionParams(masCode(lineId).Code, curPos, fParams, returnFormat, funcPos, nextEl, arrParams, False) = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If

                        If funcPos < 0 OrElse masCode(lineId).Code(funcPos).wordType <> CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                            LAST_ERROR = "Ошибка в синтаксисе блока Wrap. Свойство не найдено."
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If

                        Dim child2Id As Integer = -1, child3Id As Integer = -1, classId As Integer = curLineCode(curPos).classId
                        If fParams.Count > 0 Then child2Id = GetSecondChildIdByName(fParams(0), mScript.mainClass(classId).ChildProperties)
                        If fParams.Count > 1 AndAlso child2Id > -1 Then child3Id = GetThirdChildIdByName(fParams(1), child2Id, mScript.mainClass(classId).ChildProperties)

                        'Получаем индекс последней строки в блоке
                        strResult = BlockWrap(masCode, startLine, finalLine)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If

                        'создаем структуру codeBox.CodeData() и помещаем туда код события
                        Dim dt() As CodeTextBox.CodeDataType = Nothing
                        Array.Resize(dt, finalLine - lineId) '+ 2)

                        Dim dLvl As Integer = 0
                        For i As Integer = lineId + 1 To finalLine
                            dLvl = i - lineId - 1
                            dt(dLvl) = New CodeTextBox.CodeDataType
                            Array.Resize(dt(dLvl).Code, masCode(i).Code.Length)
                            Array.ConstrainedCopy(masCode(i).Code, 0, dt(dLvl).Code, 0, masCode(i).Code.Length)
                        Next
                        If eventRouter.SetPropertyWithEvent(classId, masCode(lineId).Code(funcPos).Word, dt, child2Id, child3Id) = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If
                        If PropertiesRouterRunSpecific(classId, masCode(lineId).Code(funcPos).Word, "", fParams.ToArray, arrParams, True) = "#Error" Then
                            'do specific actions for properties set (for ex., change description on the screen for L.Description)
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If
                    ElseIf curLineCode(curPos).classId <> -1 AndAlso curLineCode(curPos).wordType <> CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
                        LAST_ERROR = "Ошибка в синтаксисе блока Wrap. Не указана переменная."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    Else
                        'variable changing
                        'Получаем текст и индекс последней строки в блоке
                        strResult = MakeExec(masCode, startLine, finalLine, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If
                        Dim varValue As String = "'" + strResult.Replace("'", "/'") + "'"

                        'получаем название и индекс переменной
                        Dim varName As String = "", varIndex As Integer, varSignature As String = ""
                        strResult = SplitVarFromFullName(curLineCode, curPos, varName, varIndex, varSignature, nextPos, returnFormat, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If

                        'Получаем ссылку на класс переменной
                        Dim cVar As cVariable
                        If csLocalVariables.GetVariable(varName, 0) = "#Error" Then
                            If csPublicVariables.GetVariable(varName, 0) = "#Error" Then
                                cVar = csLocalVariables
                            Else
                                cVar = csPublicVariables
                            End If
                        Else
                            cVar = csLocalVariables
                        End If

                        'сохраняем результат блока в переменной
                        If varSignature.Length = 0 Then
                            cVar.SetVariableInternal(varName, varValue, varIndex)
                        Else
                            varSignature = mScript.PrepareStringToPrint(varSignature, arrParams)
                            If varSignature = "#Error" Then
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return varSignature
                            End If

                            cVar.SetVariableInternal(varName, varValue, varSignature)
                        End If
                    End If

                    returnFormat = ReturnFormatEnum.ORIGINAL
                    lineId = finalLine + 2
                    Continue Do
                Case CodeTextBox.EditWordTypeEnum.W_BLOCK_DOWHILE
                    Dim startLine As Integer = lineId
                    Dim finalLine As Integer = -1
                    curLineCode = masCode(lineId).Code
                    'If curLineCode.Length < 3 Then
                    '    LAST_ERROR = "Неверный синтаксис блока Do While."
                    '    If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                    '    Return "#Error"
                    'End If
                    'Получаем индекс последней строки в блоке
                    strResult = BlockDoWhile(masCode, startLine, finalLine)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If

                    If finalLine <= startLine Then
                        'пустой блок
                        lineId += 1
                        Continue Do
                    End If

                    'Выполняем цикл
                    Dim cycleResult = "True"
                    startLine += 1
                    Do While cycleResult = "True"
                        If curLineCode.Length > 3 Then
                            'блок с условием While
                            cycleResult = GetValue(curLineCode, 2, -1, arrParams)
                        Else
                            'бесконечный блок
                            cycleResult = "True"
                        End If
                        If cycleResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        ElseIf cycleResult = "True" Then
                            'Выполняем блок
                            strResult = ExecuteCode(masCode, arrParams, False, startLine, finalLine)
fromMark:
                            If strResult = "#Error" Then
                                LAST_ERROR = "Ошибка внутри цикла Do While. " + LAST_ERROR
                                If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                Return "#Error"
                            End If
                            If EXIT_CODE Then
                                'Из кода получен сигнал о завершении, но это не связано с ошибкой (strResult <> "#Error")
                                'Это может быть Return, Exit или Break
                                If strResult.StartsWith("#BREAK#") Then
                                    'Это Break
                                    If strResult = "#BREAK#" Then
                                        EXIT_CODE = False 'Break без параметров - продолжаем работу
                                        Exit Do
                                    Else
                                        'Break с параметом = кол-во циклов, из которых надо выйти (напр., #BREAK#3 )
                                        Dim breaksLeft As Integer = Convert.ToInt32(strResult.Substring(7)) - 1 'отнимаем 1 от кол-ва циклов, из которых надо выйти
                                        If breaksLeft <= 0 Then
                                            EXIT_CODE = False 'больше ни откуда выходить не надо
                                            Exit Do
                                        Else
                                            If isPrimaryCode Then
                                                LAST_ERROR = "Некорректное использование оператора Break. Больше нет циклов."
                                                If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                                Return "#Error"
                                            Else
                                                Return "#BREAK#" + breaksLeft.ToString 'выходим и передаем родительской функции ExecuteCode параметр на 1 меньше
                                            End If
                                        End If
                                    End If
                                Else
                                    Return strResult 'Это Return или Exit
                                End If
                            End If
                            If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                                Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                                Dim markLine As Integer = GetLineAfterMark(masCode, markName)
                                If markLine = -1 Then
                                    LAST_ERROR = "Метка " + markName + " не найдена."
                                    AnalizeError(masCode(lineId), arrParams)
                                    Return "#Error"
                                Else
                                    If markLine >= startLine AndAlso markLine <= finalLine Then
                                        'метка в пределах цикла - цикл продолжается
                                        strResult = ExecuteCode(masCode, arrParams, False, markLine, finalLine)
                                        GoTo fromMark
                                    ElseIf isPrimaryCode Then
                                        lineId = markLine
                                        Continue Do
                                    Else
                                        Return strResult
                                    End If
                                End If
                            End If
                            'If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            '    If isPrimaryCode Then
                            '        strResult = strResult.Substring(6) 'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.
                            '        GoTo jmp
                            '    Else
                            '        Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                            '    End If
                            'End If
                            If strResult.StartsWith("#CONTINUE#") Then 'в блоке, который только что выполнился, возник оператор Continue
                                'If isPrimaryCode Then
                                '    LAST_ERROR = "Оператор Continue нельзя использовать вне циклов." + LAST_ERROR
                                '    If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                '    Return "#Error"
                                'Else
                                '    'ничего делать не надо - новая итерация цикла
                                'End If
                            End If
                        End If
                    Loop

                    lineId = finalLine + 2
                    Continue Do
                Case CodeTextBox.EditWordTypeEnum.W_SWITCH
                    'Блок Select Case / Switch
                    Dim startLine As Integer = lineId
                    Dim finalLine As Integer = -1
                    Dim endSelectLine As Integer = -1
                    curLineCode = masCode(lineId).Code
                    If curLineCode.Length < 2 Then
                        LAST_ERROR = "Неверный синтаксис блока Select Case."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    'Получаем начало и конец исполняемого кода внутри блока, а также позицию End Select
                    strResult = BlockSwitch(masCode, startLine, finalLine, endSelectLine, arrParams)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If

                    If finalLine < startLine Then
                        'пустой блок
                        lineId = endSelectLine + 1
                        Continue Do
                    End If

                    'выполняем блок
                    If finalLine > -1 Then
                        strResult = ExecuteCode(masCode, arrParams, False, startLine, finalLine)

                        If strResult = "#Error" Then
                            LAST_ERROR = "ошибка внутри блока Select Case. " + LAST_ERROR
                            If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                            Return "#Error"
                        End If
                        If EXIT_CODE Then
                            If strResult = "#BREAK#" Then
                                LAST_ERROR = "Оператор Break нельзя использовать в блоке Select Case / Switch."
                                If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                Return "#Error"
                            Else
                                Return strResult
                            End If
                        End If
                        If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            If isPrimaryCode Then
                                'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.
                                Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                                lineId = GetLineAfterMark(masCode, markName)
                                If lineId = -1 Then
                                    LAST_ERROR = "Метка " + curExecuteItem.Code(lineStartPos + 1).Word + " не найдена."
                                    AnalizeError(masCode(lineId), arrParams)
                                    Return "#Error"
                                Else
                                    Continue Do
                                End If
                            Else
                                Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                            End If
                        End If
                    End If

                    lineId = endSelectLine + 1
                    Continue Do
                Case CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR
                    'For i = 1 To 10 Step 2
                    Dim startLine As Integer = lineId
                    Dim finalLine As Integer = -1
                    curLineCode = masCode(lineId).Code
                    If curLineCode.Length < 6 Then
                        LAST_ERROR = "Неверный синтаксис блока For ... Next."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    'Получаем индекс последней строки в блоке
                    strResult = BlockFor(masCode, startLine, finalLine)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If

                    If finalLine <= startLine Then
                        'пустой блок
                        lineId += 1
                        Continue Do
                    End If

                    'получаем переменную цикла
                    Dim curPos As Integer = 1
                    returnFormat = ReturnFormatEnum.ORIGINAL
                    If curLineCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Then
                        curPos += 1
                        returnFormat = ReturnFormatEnum.TO_NUMBER
                    ElseIf curLineCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING Then
                        curPos += 1
                        returnFormat = ReturnFormatEnum.TO_STRING
                    End If
                    If curLineCode(curPos).classId <> -1 Then
                        LAST_ERROR = "Неверный синтаксис блока For ... Next. Не найдено переменной цикла."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    'получаем название и индекс переменной
                    Dim varName As String = "", varIndex As Integer, varSignature As String = ""
                    strResult = SplitVarFromFullName(curLineCode, curPos, varName, varIndex, varSignature, nextPos, returnFormat, arrParams)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    ElseIf nextPos = -1 OrElse curLineCode(nextPos).wordType <> CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL Then
                        LAST_ERROR = "Неверный синтаксис блока For ... Next."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    If String.IsNullOrEmpty(varSignature) = False Then
                        varSignature = mScript.PrepareStringToPrint(varSignature, arrParams)
                        If varSignature = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return varSignature
                        End If
                    End If
                    'Получаем ссылку на класс переменной
                    Dim cVar As cVariable
                    If csLocalVariables.GetVariable(varName, 0) = "#Error" Then
                        If csPublicVariables.GetVariable(varName, 0) = "#Error" Then
                            cVar = csLocalVariables
                        Else
                            cVar = csPublicVariables
                        End If
                    Else
                        cVar = csLocalVariables
                    End If

                    'получаем начальное значение переменной цикла
                    Dim cycleStart As Single
                    curPos = nextPos + 1
                    nextPos = -1
                    For i As Integer = curPos + 1 To curLineCode.GetUpperBound(0)
                        If curLineCode(i).Word.Trim = "To" Then
                            nextPos = i
                            Exit For
                        End If
                    Next
                    If nextPos = -1 Then
                        LAST_ERROR = "Неверный синтаксис блока For ... Next. Не найден оператор To."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    strResult = GetValue(curLineCode, curPos, nextPos - 1, arrParams)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    strResult = ConvertElement(strResult, returnFormat)
                    If Param_GetType(strResult) <> ReturnFormatEnum.TO_NUMBER Then
                        LAST_ERROR = "Начальное значение переменной в цикле For ... Next не число."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    cycleStart = Double.Parse(strResult, NumberStyles.Any, provider_points)

                    'получаем конечное значение переменной цикла
                    curPos = nextPos + 1
                    If curLineCode.GetUpperBound(0) < curPos Then
                        LAST_ERROR = "Неверный синтаксис блока For ... Next."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If

                    nextPos = -1
                    For i As Integer = curPos + 1 To curLineCode.GetUpperBound(0)
                        If curLineCode(i).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR Then
                            'найден Step
                            nextPos = i
                            Exit For
                        End If
                    Next

                    Dim cycleEnd As Single, stepValue As Single = 1
                    If nextPos = -1 Then
                        strResult = GetValue(curLineCode, curPos, -1, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If
                        strResult = ConvertElement(strResult, returnFormat)
                        If Param_GetType(strResult) <> ReturnFormatEnum.TO_NUMBER Then
                            LAST_ERROR = "Конечное значение переменной в цикле For ... Next не число."
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If
                        cycleEnd = Double.Parse(strResult, NumberStyles.Any, provider_points)
                    Else
                        strResult = GetValue(curLineCode, curPos, nextPos - 1, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If
                        strResult = ConvertElement(strResult, returnFormat)
                        If Param_GetType(strResult) <> ReturnFormatEnum.TO_NUMBER Then
                            LAST_ERROR = "Конечное значение переменной в цикле For ... Next не число."
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If
                        cycleEnd = Double.Parse(strResult, NumberStyles.Any, provider_points)

                        'получаем шаг цикла
                        strResult = GetValue(curLineCode, nextPos + 1, -1, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If
                        strResult = ConvertElement(strResult, returnFormat)
                        If Param_GetType(strResult) <> ReturnFormatEnum.TO_NUMBER Then
                            LAST_ERROR = "Шаг переменной в цикле For ... Next не число."
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If
                        stepValue = Double.Parse(strResult, NumberStyles.Any, provider_points)
                    End If

                    'Выполняем цикл
                    startLine += 1
                    For i As Single = cycleStart To cycleEnd Step stepValue
                        'присваиваем новое значение перменной цикла
                        If varSignature.Length = 0 Then
                            cVar.SetVariableInternal(varName, i.ToString, varSignature)
                        Else
                            cVar.SetVariableInternal(varName, i.ToString, varIndex)
                        End If

                        'Выполняем блок
                        strResult = ExecuteCode(masCode, arrParams, False, startLine, finalLine)
fromMark2:
                        If strResult = "#Error" Then
                            LAST_ERROR = "Ошибка внутри цикла For ... Next. " + LAST_ERROR
                            If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                            Return "#Error"
                        End If
                        If EXIT_CODE Then
                            'Из кода получен сигнал о завершении, но это не связано с ошибкой (strResult <> "#Error")
                            'Это может быть Return, Exit или Break
                            If strResult.StartsWith("#BREAK#") Then
                                'Это Break
                                If strResult = "#BREAK#" Then
                                    EXIT_CODE = False 'Break без параметров - продолжаем работу
                                    Exit For
                                Else
                                    'Break с параметом = кол-во циклов, из которых надо выйти (напр., #BREAK#3 )
                                    Dim breaksLeft As Integer = Convert.ToInt32(strResult.Substring(7)) - 1 'отнимаем 1 от кол-ва циклов, из которых надо выйти
                                    If breaksLeft <= 0 Then
                                        EXIT_CODE = False 'больше ни откуда выходить не надо
                                        Exit For
                                    Else
                                        If isPrimaryCode Then
                                            LAST_ERROR = "Некорректное использование оператора Break. Больше нет циклов."
                                            If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                            Return "#Error"
                                        Else
                                            Return "#BREAK#" + breaksLeft.ToString 'выходим и передаем родительской функции ExecuteCode параметр на 1 меньше
                                        End If
                                    End If
                                End If
                            Else
                                Return strResult 'Это Return или Exit
                            End If
                        End If

                        If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                            Dim markLine As Integer = GetLineAfterMark(masCode, markName)
                            If markLine = -1 Then
                                LAST_ERROR = "Метка " + markName + " не найдена."
                                AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            Else
                                If markLine >= startLine AndAlso markLine <= finalLine Then
                                    'метка в пределах цикла - цикл продолжается
                                    strResult = ExecuteCode(masCode, arrParams, False, markLine, finalLine)
                                    GoTo fromMark2
                                ElseIf isPrimaryCode Then
                                    lineId = markLine
                                    Continue Do
                                Else
                                    Return strResult
                                End If
                            End If
                        End If


                        'If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                        '    If isPrimaryCode Then
                        '        strResult = strResult.Substring(6) 'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.

                        '        Dim markPos As Integer = 0
                        '        If isPrimaryCode Then
                        '            While markPos < masCode.Count - 1
                        '                If masCode(markPos).Code.Length > 0 AndAlso masCode(markPos).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_MARK _
                        '                    AndAlso masCode(markPos).Code(0).Word = strResult + ":" Then
                        '                    lineId = markPos + 1
                        '                    Continue Do
                        '                End If
                        '                markPos += 1
                        '            End While
                        '            LAST_ERROR = "Метка " + curExecuteItem.Code(lineStartPos + 1).Word + " не найдена."
                        '            AnalizeError(masCode(lineId), arrParams)
                        '            Return "#Error"
                        '        Else
                        '            Return "#JUMP#" + strResult
                        '        End If



                        '        GoTo jmp
                        '    Else
                        '        Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                        '    End If
                        'End If

                        If strResult.StartsWith("#CONTINUE#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            'If isPrimaryCode Then
                            '    LAST_ERROR = "Оператор Continue нелбзя использовать вне циклов." + LAST_ERROR
                            '    If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                            '    Return "#Error"
                            'Else
                            '    'ничего делать не надо - новая итерация цикла
                            'End If
                        End If


                    Next

                    lineId = finalLine + 2
                    Continue Do
                Case CodeTextBox.EditWordTypeEnum.W_BLOCK_IF
                    'Блок If ... Then ... ElseIf ... Else ...End If
                    If curExecuteItem.Code.Length < 3 AndAlso lineStartPos <> 0 Then
                        LAST_ERROR = "Ошибка в ситаксисе блока If ... Then."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If

                    If curExecuteItem.Code(curExecuteItem.Code.GetUpperBound(0)).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_IF Then
                        'многострочный блок
                        'Получаем начало и конец исполняемого кода внутри блока, а также позицию End If
                        Dim startLine As Integer = lineId
                        Dim finalLine As Integer = -1, endIfLine As Integer = -1
                        strResult = BlockIF(masCode, startLine, finalLine, endIfLine, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If

                        If finalLine > startLine Then
                            'если есть блок на выполнение (одно из условий блока If ... Then вернуло True) и этот блок не пустой, то
                            'Выполняем блок
                            strResult = ExecuteCode(masCode, arrParams, False, startLine + 1, finalLine)
                            If strResult = "#Error" Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                            If EXIT_CODE Then
                                If strResult = "#BREAK#" Then
                                    LAST_ERROR = "Оператор Break нельзя использовать в блоке If ... Then."
                                    If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                    Return "#Error"
                                Else
                                    Return strResult
                                End If
                            End If

                            If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                                Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                                Dim markLine As Integer = GetLineAfterMark(masCode, markName)
                                If markLine = -1 Then
                                    LAST_ERROR = "Метка " + markName + " не найдена."
                                    AnalizeError(masCode(lineId), arrParams)
                                    Return "#Error"
                                Else
                                    If markLine >= startLine AndAlso markLine <= finalLine Then
                                        'метка в пределах блока, причем только той части If Then, которая пдошла по условиям - выполняем код только внутри
                                        strResult = ExecuteCode(masCode, arrParams, False, markLine, finalLine)
                                        Return strResult
                                    ElseIf isPrimaryCode Then
                                        lineId = markLine
                                        Continue Do
                                    Else
                                        Return strResult
                                    End If
                                End If
                            End If

                            'If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            '    If isPrimaryCode Then
                            '        strResult = strResult.Substring(6) 'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.
                            '        GoTo jmp
                            '    Else
                            '        Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                            '    End If
                            'End If
                            If strResult.StartsWith("#CONTINUE#") Then
                                If isPrimaryCode Then
                                    LAST_ERROR = "Оператор Continue нельзя использовать за пределами цикла."
                                    If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                    Return "#Error"
                                Else
                                    Return strResult
                                End If
                            End If

                        End If
                        lineId = endIfLine + 1
                        Continue Do
                    Else
                        'блок однострочный
                        Dim thenPos As Integer = -1
                        For i As Integer = lineStartPos + 1 To curExecuteItem.Code.GetUpperBound(0)
                            If curExecuteItem.Code(i).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_IF Then
                                thenPos = i
                                Exit For
                            End If
                        Next
                        If thenPos = -1 Then
                            LAST_ERROR = "Ошибка в ситаксисе блока If. Не найдено ключевое слово Then."
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If

                        'получаем условие и его результат
                        strResult = GetValue(curExecuteItem.Code, lineStartPos + 1, thenPos - 1, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If
                        If strResult = "True" Then
                            'создаем блок кода на выполнение (из того, что за Then)

                            'создаем копию кода, содержащего только данные за словом Then
                            Dim masCopy() As CodeTextBox.EditWordType = Nothing
                            Array.Resize(masCopy, curExecuteItem.Code.Length - thenPos - 1)
                            Array.Copy(curExecuteItem.Code, thenPos + 1, masCopy, 0, masCopy.Length)
                            'создаем блок кода для передачи
                            Dim exCopy As New List(Of ExecuteDataType)
                            exCopy.Add(New ExecuteDataType(masCopy, curExecuteItem.lineId))
                            'Выполняем блок
                            strResult = ExecuteCode(exCopy, arrParams)

                            If strResult = "#Error" Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                            If EXIT_CODE Then
                                If strResult = "#BREAK#" AndAlso isPrimaryCode Then
                                    LAST_ERROR = "Оператор Break нельзя использовать в блоке If ... Then."
                                    If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                    Return "#Error"
                                Else
                                    Return strResult
                                End If
                            End If


                            If strResult.StartsWith("#JUMP#") Then 'в однострочном блоке If Then возник оператор Jump
                                If isPrimaryCode Then
                                    Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                                    Dim markLine As Integer = GetLineAfterMark(masCode, markName)
                                    lineId = markLine
                                    If markLine = -1 Then
                                        LAST_ERROR = "Метка " + markName + " не найдена."
                                        AnalizeError(masCode(lineId), arrParams)
                                        Return "#Error"
                                    Else
                                        Continue Do
                                    End If
                                Else
                                    Return strResult
                                End If
                            End If

                            'If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            '    If isPrimaryCode Then
                            '        strResult = strResult.Substring(6) 'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.
                            '        GoTo jmp
                            '    Else
                            '        Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                            '    End If
                            'End If
                            If strResult = ("#CONTINUE#") Then
                                If isPrimaryCode Then
                                    LAST_ERROR = "Оператор Continue нельзя использовать за пределами цикла."
                                    If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                    Return "#Error"
                                Else
                                    Return strResult
                                End If
                            End If
                        End If
                    End If
                Case CodeTextBox.EditWordTypeEnum.W_VARIABLE, CodeTextBox.EditWordTypeEnum.W_CLASS, _
                    CodeTextBox.EditWordTypeEnum.W_FUNCTION, CodeTextBox.EditWordTypeEnum.W_PROPERTY
                    'присвоение значения переменной / свойству / выполнение функции
                    'myClass[x, y].#myFuncProp(u, v) ...
                    If curExecuteItem.Code(lineStartPos).classId = -1 Then
                        'Присвоение значения переменной
                        'Var.#myVar[x] ...
                        'получаем имя и индекс переменной
                        Dim varName As String = "", varIndex As Integer, varSignature As String = ""
                        strResult = SplitVarFromFullName(curExecuteItem.Code, lineStartPos, varName, varIndex, varSignature, nextPos, returnFormat, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If

                        If nextPos = -1 OrElse (curExecuteItem.Code(nextPos).wordType <> CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL _
                                                AndAlso curExecuteItem.Code(nextPos).wordType <> CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH) Then
                            LAST_ERROR = "После переменной должен стоять оператор присвоения значения."
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If

                        'получаем значение для присваивания
                        Dim varResult As String
                        If curExecuteItem.Code(nextPos).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH Then
                            'Получаем массив кода для операторов += / -=
                            'Например, из "a += 5" получаем "а + 5"
                            'получаем значение переменной
                            If varSignature.Length = 0 Then
                                varResult = csLocalVariables.GetVariable(varName, varIndex)
                            Else
                                varResult = csLocalVariables.GetVariable(varName, varSignature)
                            End If
                            If varResult = "#Error" Then
                                If varSignature.Length = 0 Then
                                    varResult = csPublicVariables.GetVariable(varName, varIndex)
                                Else
                                    varResult = csPublicVariables.GetVariable(varName, varSignature)
                                End If
                            End If
                            If varResult = "#Error" Then
                                LAST_ERROR = "Операторы += и -= не могут быть использованы. Необходимо предварительно инициализировать переменную."
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return "#Error"
                            End If
                            'создаем ячейку кода с этим значением
                            Dim varCode As New CodeTextBox.EditWordType
                            varCode.Word = varResult
                            Select Case Param_GetType(varResult)
                                Case ReturnFormatEnum.TO_NUMBER
                                    varCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER
                                Case ReturnFormatEnum.TO_STRING
                                    varCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING
                                Case ReturnFormatEnum.TO_BOOL
                                    varCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL
                            End Select
                            'создаем копию кода, содержащего только данные для получения значения переменной
                            Dim masCopy() As CodeTextBox.EditWordType = Nothing
                            If curExecuteItem.Code(nextPos).Word.Trim = "+=" Then
                                Array.Resize(masCopy, curExecuteItem.Code.Length + 1 - nextPos)
                                Array.Copy(curExecuteItem.Code, nextPos + 1, masCopy, 2, masCopy.Length - 2)
                                masCopy(1) = New CodeTextBox.EditWordType With {.Word = " + ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH}
                            Else '-=
                                Array.Resize(masCopy, curExecuteItem.Code.Length + 3 - nextPos)
                                Array.Copy(curExecuteItem.Code, nextPos + 1, masCopy, 3, masCopy.Length - 4)
                                masCopy(2) = New CodeTextBox.EditWordType With {.wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, .Word = "( "}
                                masCopy(masCopy.Count - 1) = New CodeTextBox.EditWordType With {.wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, .Word = ") "}
                                masCopy(1) = New CodeTextBox.EditWordType With {.Word = " - ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH}
                            End If

                            masCopy(0) = varCode
                            strResult = GetValue(masCopy, 0, -1, arrParams)
                        Else
                            strResult = GetValue(curExecuteItem.Code, nextPos + 1, -1, arrParams, False)
                        End If
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If

                        If strResult = "#ARRAY" Then
                            Dim var As cVariable.variableEditorInfoType = Nothing
                            If csLocalVariables.lstVariables.TryGetValue(varName, var) = False Then
                                If csPublicVariables.lstVariables.TryGetValue(varName, var) = False Then
                                    csLocalVariables.SetVariableInternal(varName, "", 0)
                                    var = csLocalVariables.lstVariables(varName)
                                End If
                            End If
                            var.arrValues = lastArray.arrValues
                            var.lstSingatures = lastArray.lstSingatures
                            lastArray = Nothing
                        Else
                            strResult = ConvertElement(strResult, returnFormat, ReturnFormatEnum.ORIGINAL)

                            'Присваиваем значение переменной
                            varResult = csLocalVariables.GetVariable(varName, 0)
                            If varSignature.Length = 0 Then
                                If varResult <> "#Error" Then
                                    csLocalVariables.SetVariableInternal(varName, strResult, varIndex)
                                Else
                                    varResult = csPublicVariables.GetVariable(varName, varIndex)
                                    If varResult <> "#Error" Then
                                        csPublicVariables.SetVariableInternal(varName, strResult, varIndex)
                                    Else
                                        csLocalVariables.SetVariableInternal(varName, strResult, varIndex)
                                    End If
                                End If
                            Else
                                varSignature = mScript.PrepareStringToPrint(varSignature, arrParams)
                                If varSignature = "#Error" Then
                                    If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                    Return varSignature
                                End If

                                If varResult <> "#Error" Then
                                    csLocalVariables.SetVariableInternal(varName, strResult, varSignature)
                                Else
                                    varResult = csPublicVariables.GetVariable(varName, varSignature)
                                    If varResult <> "#Error" Then
                                        csPublicVariables.SetVariableInternal(varName, strResult, varSignature)
                                    Else
                                        csLocalVariables.SetVariableInternal(varName, strResult, varSignature)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        'функция или свойство
                        'myClass[x, y].#myFuncProp(u, v) ...
                        'получаем параметры в квадратных и/или круглых скобках, а также 
                        Dim funcParams As New List(Of String)
                        Dim funcPos As Integer = 0
                        strResult = GetFunctionParams(curExecuteItem.Code, lineStartPos, funcParams, returnFormat, funcPos, nextPos, arrParams, True)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If

                        'уточняем класс, свойство/функцию для пользовательских классов и элементов
                        If curExecuteItem.Code(funcPos).classId < 0 Then strResult = ClarifyUserElement(curExecuteItem.Code(funcPos))
                        If strResult = "#Error" Then
                            LAST_ERROR = "Свойство/функция " + curExecuteItem.Code(funcPos).Word + " не найдена."
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If

                        If curExecuteItem.Code(funcPos).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                            'вызов функции как процедуры
                            strResult = FunctionRouter(curExecuteItem.Code(funcPos).classId, curExecuteItem.Code(funcPos).Word.Trim, funcParams.ToArray, arrParams)
                            If strResult = "#Error" Then
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return strResult
                            End If
                        Else
                            'присвоение значения свойству
                            If nextPos = -1 Then
                                LAST_ERROR = "За свойством не стоит опрератор присвоения значения."
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return "#Error"
                            End If
                            If curExecuteItem.Code(nextPos).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL Then
                                'простое присваивание prop = x
                                strResult = GetValue(curExecuteItem.Code, nextPos + 1, -1, arrParams)
                            ElseIf curExecuteItem.Code(nextPos).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH Then
                                'математическое присваивание prop += x
                                'получаем значение свойства
                                Dim propCode As New CodeTextBox.EditWordType

                                propCode.Word = PropertiesRouter(curExecuteItem.Code(funcPos).classId, curExecuteItem.Code(funcPos).Word, funcParams.ToArray, arrParams)
                                If propCode.Word = "#Error" Then
                                    If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                    Return "#Error"
                                End If

                                Select Case Param_GetType(propCode.Word)
                                    Case ReturnFormatEnum.TO_NUMBER
                                        propCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER
                                    Case ReturnFormatEnum.TO_STRING
                                        propCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING
                                    Case ReturnFormatEnum.TO_BOOL
                                        propCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL
                                End Select

                                Dim masCopy() As CodeTextBox.EditWordType = Nothing
                                'создаем копию кода, содержащего только данные для получения значения свойства
                                If curExecuteItem.Code(nextPos).Word.Trim = "+=" Then
                                    Array.Resize(masCopy, curExecuteItem.Code.Length + 1 - nextPos)
                                    Array.Copy(curExecuteItem.Code, nextPos + 1, masCopy, 2, masCopy.Length - 2)
                                    masCopy(1) = New CodeTextBox.EditWordType With {.Word = " + ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH}
                                Else '-=
                                    Array.Resize(masCopy, curExecuteItem.Code.Length + 3 - nextPos)
                                    Array.Copy(curExecuteItem.Code, nextPos + 1, masCopy, 3, masCopy.Length - 4)
                                    masCopy(2) = New CodeTextBox.EditWordType With {.wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, .Word = "( "}
                                    masCopy(masCopy.Count - 1) = New CodeTextBox.EditWordType With {.wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, .Word = ") "}
                                    masCopy(1) = New CodeTextBox.EditWordType With {.Word = " - ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH}
                                End If

                                masCopy(0) = propCode
                                strResult = GetValue(masCopy, 0, -1, arrParams)
                            Else
                                LAST_ERROR = "За свойством не стоит опрератор присвоения значения."
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return "#Error"
                            End If

                            If strResult = "#Error" Then
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return "#Error"
                            ElseIf strResult.StartsWith(temporary_system_array_BASE_check) Then
                                LAST_ERROR = "Свойству нельзя присвоить массив."
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return "#Error"
                            End If
                            strResult = ConvertElement(strResult, returnFormat)
                            strResult = PropertiesRouter(curExecuteItem.Code(funcPos).classId, curExecuteItem.Code(funcPos).Word, funcParams.ToArray, arrParams, PropertiesOperationEnum.PROPERTY_SET, strResult)
                            If strResult = "#Error" Then
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return "#Error"
                            End If
                        End If

                    End If

                    returnFormat = ReturnFormatEnum.ORIGINAL
                Case MatewScript.WordTypeEnum.W_BLOCK_FUNCTION
                    'создаем/обновляем функцию пользователя в массиве functionsHashEx
                    Dim startLine As Integer = lineId
                    Dim finalLine As Integer = -1
                    Dim curPos As Integer = 1
                    curLineCode = masCode(lineId).Code
                    If curLineCode.Length < 2 Then
                        LAST_ERROR = "Неверный синтаксис блока Function."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    'Получаем индекс последней строки в блоке
                    strResult = BlockFunction(masCode, startLine, finalLine)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If

                    'получаем имя функции
                    Dim funcName As String = mScript.PrepareStringToPrint(curLineCode(curPos).Word, arrParams, True)
                    Dim fClassId As Integer = -1
                    Dim isClassFunction As Boolean = GetClassIdAndElementNameByString(curLineCode(curPos).Word, fClassId, funcName, True, False, True)

                    'создаем блок кода функции
                    Dim codeCopy As New List(Of ExecuteDataType)
                    For i As Integer = startLine + 1 To finalLine
                        codeCopy.Add(New ExecuteDataType(masCode(i).Code, i - startLine - 1))
                    Next

                    If isClassFunction Then
                        'Это функция класса, добавленная Писателем. Вносим соответствующие изменения в структуру mainClass и создаем событие
                        If mainClass(fClassId).Functions.ContainsKey(funcName) = False Then
                            If AddUserFunction({mainClass(fClassId).Names(0), WrapString(funcName)}, arrParams) = "#Error" Then
                                'Если указана пока еще не созданная функция пользователя - создаем ее
                                LAST_ERROR = "Неверный синтаксис блока Function. Не удалось зодать новую функцию " + funcName + " в классе " + mainClass(fClassId).Names(0)
                                If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                                Return "#Error"
                            End If
                        End If

                        Dim func As PropertiesInfoType = mainClass(fClassId).Functions(funcName)
                        func.Value = "" 'создавать серализованный аналог необязательно
                        func.eventId = eventRouter.SetEventId(func.eventId, codeCopy) 'создаем/заменяем событие
                    Else
                        'Это обычная функция. Копируем блок функции в хэш functionsEx
                        Dim f As FunctionInfoType
                        If functionsHash.ContainsKey(funcName) Then
                            f = functionsHash(funcName)
                            If IsNothing(f.ValueDt) = False Then Erase f.ValueDt
                            If IsNothing(f.Variables) = False Then f.Variables.Clear()
                            If IsNothing(f.ValueExecuteDt) = False Then f.ValueExecuteDt.Clear()
                            f.ValueExecuteDt = codeCopy
                        Else
                            f = New FunctionInfoType With {.ValueExecuteDt = codeCopy}
                            functionsHash.Add(funcName, f)
                        End If
                        'копируем текущие локальные переменные
                        If IsNothing(csLocalVariables.lstVariables) = False AndAlso csLocalVariables.lstVariables.Count > 0 Then csLocalVariables.CopyVariables(f.Variables)
                    End If

                    returnFormat = ReturnFormatEnum.ORIGINAL
                    lineId = finalLine + 2 'устанавливаем текущую позицию за End Function
                    Continue Do

                Case MatewScript.WordTypeEnum.W_PARAM
                    'присвоение значения параметру, с которыми запущен код
                    'Param[x] ...
                    'получаем индекс параметра
                    Dim pIndex As Integer = 0
                    If curExecuteItem.Code.Length <= 1 + lineStartPos Then
                        LAST_ERROR = "После параметра должен стоять оператор присвоения значения."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If

                    If curExecuteItem.Code(lineStartPos + 1).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                        '[..] - есть квадратные скобки
                        nextPos = GetNextElementIdAfterBrackets(curExecuteItem.Code, lineStartPos + 1)
                        If nextPos = -2 Then
                            LAST_ERROR = "Не найдена закрывающая скобка ]."
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        ElseIf nextPos = -1 Then
                            'за квадратными скобками нет опреатора присвоения
                            LAST_ERROR = "После параметра должен стоять оператор присвоения значения."
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return "#Error"
                        End If
                        'получаем результат выражения в квадратных скобках. Он же будет индексом параметра
                        strResult = GetValue(curExecuteItem.Code, lineStartPos + 2, nextPos - 2, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                            Return strResult
                        End If
                        pIndex = Integer.Parse(strResult, NumberStyles.Any, provider_points)
                    Else
                        'квадратных скобок [] нет - просто Param
                        nextPos = lineStartPos + 1
                    End If

                    If IsNothing(arrParams) OrElse pIndex < 0 OrElse arrParams.Length < pIndex + 1 Then
                        LAST_ERROR = "Параметра с индексом " + CStr(pIndex) + " не существует."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If

                    If curExecuteItem.Code(nextPos).wordType <> CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL _
                        AndAlso curExecuteItem.Code(nextPos).wordType <> CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH Then
                        LAST_ERROR = "После переменной должен стоять оператор присвоения значения."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If

                    'получаем значение для присваивания
                    Dim varResult As String = ""
                    If curExecuteItem.Code(nextPos).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH Then
                        'Получаем массив кода для операторов += / -=
                        'Например, из "Param += 5" получаем "Param + 5"
                        'получаем значение параметра
                        strResult = arrParams(pIndex)
                        'создаем ячейку кода с этим значением
                        Dim varCode As New CodeTextBox.EditWordType
                        varCode.Word = strResult
                        Select Case Param_GetType(strResult)
                            Case ReturnFormatEnum.TO_NUMBER
                                varCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER
                            Case ReturnFormatEnum.TO_STRING
                                varCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING
                            Case ReturnFormatEnum.TO_BOOL
                                varCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL
                        End Select
                        'создаем копию кода, содержащего только данные для получения значения переменной
                        Dim masCopy() As CodeTextBox.EditWordType = Nothing
                        If curExecuteItem.Code(nextPos).Word.Trim = "+=" Then
                            Array.Resize(masCopy, curExecuteItem.Code.Length + 1 - nextPos)
                            Array.Copy(curExecuteItem.Code, nextPos + 1, masCopy, 2, masCopy.Length - 2)
                            masCopy(1) = New CodeTextBox.EditWordType With {.Word = " + ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH}
                        Else '-=
                            Array.Resize(masCopy, curExecuteItem.Code.Length + 3 - nextPos)
                            Array.Copy(curExecuteItem.Code, nextPos + 1, masCopy, 3, masCopy.Length - 4)
                            masCopy(2) = New CodeTextBox.EditWordType With {.wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, .Word = "( "}
                            masCopy(masCopy.Count - 1) = New CodeTextBox.EditWordType With {.wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, .Word = ") "}
                            masCopy(1) = New CodeTextBox.EditWordType With {.Word = " - ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH}
                        End If

                        masCopy(0) = varCode
                        strResult = GetValue(masCopy, 0, -1, arrParams)
                    Else
                        strResult = GetValue(curExecuteItem.Code, nextPos + 1, -1, arrParams)
                    End If
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    ElseIf strResult.StartsWith(temporary_system_array_BASE_check) Then
                        LAST_ERROR = "Нельзя параметру присвоить массив."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return strResult
                    End If
                    strResult = ConvertElement(strResult, returnFormat, ReturnFormatEnum.ORIGINAL)

                    'Присваиваем значение Param
                    arrParams(pIndex) = strResult
                Case MatewScript.WordTypeEnum.W_RETURN
                    'return xxxx
                    If curExecuteItem.Code.Length = 1 Then '0 + lineStartPos Then
                        csLocalVariables.KillVars()
                        If isPrimaryCode = False Then EXIT_CODE = True
                        Return ""
                    Else
                        strResult = GetValue(curExecuteItem.Code, lineStartPos + 1, -1, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            EXIT_CODE = True
                            Return "#Error"
                        End If
                        csLocalVariables.KillVars()
                        If isPrimaryCode = False Then EXIT_CODE = True
                        Return strResult
                    End If
                Case MatewScript.WordTypeEnum.W_EXIT
                    If isPrimaryCode = False Then EXIT_CODE = True
                    csLocalVariables.KillVars()
                    Return ""
                Case MatewScript.WordTypeEnum.W_JUMP
                    If curExecuteItem.Code.Length = 1 Then
                        LAST_ERROR = "За ключевым словом Jump не стоит метка."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    Dim markPos As Integer = 0
                    strResult = PrepareStringToPrint(curExecuteItem.Code(lineStartPos + 1).Word, arrParams)
                    'jmp:
                    If isPrimaryCode Then
                        lineId = GetLineAfterMark(masCode, strResult)
                        If lineId = -1 Then
                            LAST_ERROR = "Метка " + curExecuteItem.Code(lineStartPos + 1).Word + " не найдена."
                            AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        Else
                            Continue Do
                        End If
                    Else
                        Return "#JUMP#" + strResult
                    End If
                Case MatewScript.WordTypeEnum.W_CONTINUE
                    If isPrimaryCode Then
                        LAST_ERROR = "Оператор Continue не может применяться вне цикла."
                        AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    Else
                        Return "#CONTINUE#"
                    End If
                Case MatewScript.WordTypeEnum.W_BREAK
                    'Оператор Break выходит из циклов For...Next & Do While.
                    'Может содержать параметр, из скольких циклов выйти
                    If isPrimaryCode Then
                        LAST_ERROR = "Оператор Break не может применяться вне цикла."
                        AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    Else
                        EXIT_CODE = True
                        If curExecuteItem.Code.Length = 1 Then
                            Return "#BREAK#"
                        Else
                            EXIT_CODE = False
                            strResult = GetValue(curExecuteItem.Code, lineStartPos + 1, -1, arrParams)
                            EXIT_CODE = True
                            If strResult = "#Error" Then Return "#Error"
                            If Param_GetType(strResult) <> MatewScript.ReturnFormatEnum.TO_NUMBER Then
                                LAST_ERROR = "За оператором Break может стоять только число."
                                Return "#Error"
                            End If
                            Return "#BREAK#" + strResult
                        End If
                    End If
                Case MatewScript.WordTypeEnum.W_CYCLE_INTERNAL
                    LAST_ERROR = "Обнаружен оператор блока (" + curExecuteItem.Code(lineStartPos).Word + ") вне этого блока. Возможно, неправильное размещение метки."
                    If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                    Return "#Error"
                Case MatewScript.WordTypeEnum.W_BLOCK_NEWCLASS
                    'добавляем новый класс пользователя
                    Dim blockFinalPos As Integer = -1
                    If CreateUserClass(masCode, lineId, blockFinalPos, arrParams) = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                        Return "#Error"
                    End If
                    lineId = blockFinalPos + 1
                    Continue Do
                Case MatewScript.WordTypeEnum.W_REM_CLASS
                    'удаляем класс пользователя
                    If curExecuteItem.Code.Length < 3 Then
                        LAST_ERROR = "Неверная запись блока Rem Class."
                        If isPrimaryCode Then AnalizeError(curExecuteItem, arrParams)
                        Return "#Error"
                    End If
                    strResult = GetValue(curExecuteItem.Code, 2, -1, arrParams)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    strResult = PrepareStringToPrint(strResult, arrParams)
                    RemoveUserClass(strResult)
            End Select

            lineId += 1 'переходим к новой строке кода и продолжаем цикл
        Loop

        If isPrimaryCode Then csLocalVariables.KillVars()
        Return ""
    End Function

    ''' <summary>Главная функция, выполняющая весь код, предварительно подготовленный в массиве masCode().
    ''' Он уже не содержит комментариев, разделения строк, пустых строк, лишних пробелов и т. д.
    ''' </summary>
    ''' <param name="masCode">массив, каждый элемент которого - строка кода без лишних элементов</param>
    ''' <param name="arrParams">параметры, переданные при выполнении кода</param>
    ''' <param name="isPrimaryCode">является ли вызов первичным (т. е., не рекурсивный ли это вызов из уже выполняемого кода)</param>
    ''' <param name="startPos">первый элемент - строка кода, которую надо выполнить</param>
    ''' <param name="endPos">последняя строка, которую надо выполнить</param>
    ''' <returns>результат, который возвращет код</returns>
    Public Overridable Function ExecuteCode(ByRef masCode() As DictionaryEntry, ByRef arrParams() As String, Optional ByVal isPrimaryCode As Boolean = False, Optional ByVal startPos As Integer = 0, Optional ByVal endPos As Integer = -1) As String
        'masCode() - массив, каждый элемент которого - строка кода без лишних элементов
        'arrParams() - параметры, переданные при выполнении кода
        'isPrimaryCode - является ли вызов первичным (т. е., не рекурсивный ли это вызов из уже выполняемого кода)
        'startPos - первый элемент - строка кода, которую надо выполнить
        'endPos - последняя строка, которую надо выполнить

        If isPrimaryCode = True Then
            'Если это не рекурсивный вызов - обнуляем передыдущую инфо об ошибках
            EXIT_CODE = False
            LAST_ERROR = ""
        Else
            If EXIT_CODE = True Then Return ""
        End If

        'ВЫПОЛНЯЕМ ПОСТРОЧНО КОД
        Dim lineId As Long = startPos 'порядковый номер (начиная с 0) выполняемой строки кода
        Dim curExpression As String, strResult As String = "" 'текст выполняемой строки кода и переменная для хранения результата различных вычислений 
        Dim firstWordType As MatewScript.WordTypeEnum 'переменная для получения типа первого слова в строке кода, который получается и ф-и GetFirstWordType
        'переменные для получения данных функции GetFirstWordType (содержимое квадратных скобок, Id функции/свойства в массиве mainClass,
        'имя функции/свойства/переменной, оставшаяся часть строки кода после первого слова, формат, в который преобразуется результат вычислений перед присвоением его переменной/свойству
        Dim qbContent As String = "", currentClassId As Integer, ElementId As Integer, elementName As String = "", stringRemain As String = "", returnFormat As MatewScript.ReturnFormatEnum
        Dim funcParams As String() = {} 'Массив для хранения параметров функции/свойства
        Dim finalBracketCount As Integer = 0 'нефункциональная переменная, необходимая для вызова GetFunctionParams
        Dim funcCloseBracketPos As Integer 'для получения позиции ), закрывающей функцию/свойство
        Dim blockFinalPos As Integer, blockFinalOperator As Integer 'переменные для работы с блоками вроде If ... Then
        If endPos = -1 Then endPos = masCode.Length - 1
        'Главный цикл. Перебираем все строки кода
        Do While lineId <= endPos
            CURRENT_LINE = lineId
            curExpression = masCode(lineId).Value 'получаем код на исполнение
            'получаем первое слово в коде и другие необходимые параметры
            firstWordType = GetFirstWordType(curExpression, qbContent, currentClassId, elementName, stringRemain, returnFormat, ElementId)
            If firstWordType = MatewScript.WordTypeEnum.W_NOTHING Then 'первое слово не распознано - ошибка
                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                Return "#Error"
            End If
            Select Case firstWordType 'проверяем, что у нас за первое слово и выполняем код соответственно ему
                Case MatewScript.WordTypeEnum.W_VARIABLE_LOCAL, MatewScript.WordTypeEnum.W_VARIABLE_PUBLIC   'присвоение значения переменной
                    'если есть [], получаем результат в них: myVar[2 + 2] = x, qbContent = 4
                    Dim qbContentType As ReturnFormatEnum = ReturnFormatEnum.TO_NUMBER
                    If qbContent.Length > 0 Then
                        If Integer.TryParse(qbContent, Nothing) = False Then
                            qbContent = GetValue(qbContent, arrParams)
                            If qbContent = "#Error" Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        End If
                        qbContentType = Param_GetType(qbContent)
                        If qbContentType <> ReturnFormatEnum.TO_STRING Then
                            If Convert.ToInt32(qbContent) < 0 Then
                                LAST_ERROR = "Индекс элемента массива получился меньше 0."
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        End If
                    Else
                        qbContent = "0"
                    End If

                    If stringRemain.StartsWith("= ") = False AndAlso strResult = "#ARRAY" Then
                        LAST_ERROR = "При работе с массивами допустимы только операции присваивания."
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If

                    'Получаем результат, который надо присвоить переменной, с учетом оператора (=, += или -=)
                    If stringRemain.StartsWith("= ") Then
                        strResult = GetValue(stringRemain.Substring(2), arrParams, False)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    ElseIf stringRemain.StartsWith("+= ") Then
                        strResult = GetValue(elementName + " + " + stringRemain.Substring(3), arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    ElseIf stringRemain.StartsWith("-= ") Then
                        strResult = GetValue(elementName + " - " + stringRemain.Substring(3), arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    Else
                        If firstWordType = MatewScript.WordTypeEnum.W_VARIABLE_PUBLIC Then
                            strResult = "0"
                        Else
                            LAST_ERROR = "Неверный оператор присвоения значения переменной или нераспознанная строка."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    End If

                    'Присваиваем значение переменной
                    If strResult = "#ARRAY" Then
                        'присваиваем переменной массив
                        Dim var As cVariable.variableEditorInfoType = Nothing
                        If csLocalVariables.lstVariables.TryGetValue(elementName, var) = False Then
                            If csPublicVariables.lstVariables.TryGetValue(elementName, var) = False Then
                                csLocalVariables.SetVariableInternal(elementName, "", 0)
                                var = csLocalVariables.lstVariables(elementName)
                            End If
                        End If
                        var.arrValues = mScript.lastArray.arrValues
                        var.lstSingatures = mScript.lastArray.lstSingatures
                        lastArray = Nothing
                    Else
                        Dim setResult As Integer
                        If qbContentType = ReturnFormatEnum.TO_STRING Then
                            qbContent = mScript.PrepareStringToPrint(qbContent, arrParams)
                            If qbContent = "#Error" Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return qbContent
                            End If
                        End If
                        If firstWordType = MatewScript.WordTypeEnum.W_VARIABLE_LOCAL Then
                            'локальная переменная
                            If qbContentType = ReturnFormatEnum.TO_STRING Then
                                setResult = csLocalVariables.SetVariableInternal(elementName, strResult, qbContent, returnFormat)
                            Else
                                setResult = csLocalVariables.SetVariableInternal(elementName, strResult, CInt(qbContent), returnFormat)
                            End If
                        Else
                            'глобальная переменная
                            If qbContentType = ReturnFormatEnum.TO_STRING Then
                                setResult = csPublicVariables.SetVariableInternal(elementName, strResult, qbContent, returnFormat)
                            Else
                                setResult = csPublicVariables.SetVariableInternal(elementName, strResult, CInt(qbContent), returnFormat)
                            End If
                        End If
                        If setResult = -1 Then
                            LAST_ERROR = "Не удалось присвоить значение переменной " + elementName + "."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    End If
                Case MatewScript.WordTypeEnum.W_PROPERTY 'присвоение значения свойству
                    'получаем все параметры из [] и () в массив funcParams
                    funcParams = {}
                    funcCloseBracketPos = -1
                    If qbContent.Length > 0 Or curExpression.EndsWith("(" + stringRemain) Then
                        If curExpression.EndsWith("(" + stringRemain) Then
                            If GetFunctionParams("(" + stringRemain, funcParams, 0, funcCloseBracketPos, qbContent, finalBracketCount, arrParams) = -2 Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        Else
                            If GetFunctionParams(stringRemain, funcParams, -1, funcCloseBracketPos, qbContent, finalBracketCount, arrParams) = -2 Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        End If
                    End If
                    'если свойство имеет (), надо скорректировать оставшуюся строку (убрать из нее содержимое круглых скобок функции)
                    If funcCloseBracketPos > -1 Then stringRemain = stringRemain.Substring(funcCloseBracketPos + 1)

                    'Получаем результат, который надо присвоить свойству, с учетом оператора (=, += или -=)
                    If stringRemain.StartsWith("= ") Then
                        strResult = GetValue(stringRemain.Substring(2), arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        ElseIf strResult.StartsWith(temporary_system_array_BASE_check) Then
                            LAST_ERROR = "Нельзя параметру присвоить массив."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    ElseIf stringRemain.StartsWith("+= ") Then
                        If IsNothing(funcParams) OrElse funcParams.GetUpperBound(0) = -1 Then
                            strResult = GetValue(mainClass(mainClassHash(currentClassId)).Names(0) + "." + elementName + " + " + stringRemain.Substring(3), arrParams)
                        ElseIf funcParams.GetUpperBound(0) = 0 Then
                            strResult = GetValue(mainClass(currentClassId).Names(0) + "[" + funcParams(0) + "]." + elementName + " + " + stringRemain.Substring(3), arrParams)
                        Else
                            strResult = GetValue(mainClass(currentClassId).Names(0) + "[" + funcParams(0) + ", " + funcParams(1) + "]." + elementName + " + " + stringRemain.Substring(3), arrParams)
                        End If
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        ElseIf strResult.StartsWith(temporary_system_array_BASE_check) Then
                            LAST_ERROR = "Нельзя параметру присвоить массив."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    ElseIf stringRemain.StartsWith("-= ") Then
                        If IsNothing(funcParams) OrElse funcParams.GetUpperBound(0) = -1 Then
                            strResult = GetValue(mainClass(currentClassId).Names(0) + "." + elementName + " - " + stringRemain.Substring(3), arrParams)
                        ElseIf funcParams.GetUpperBound(0) = 0 Then
                            strResult = GetValue(mainClass(currentClassId).Names(0) + "[" + funcParams(0) + "]." + elementName + " - " + stringRemain.Substring(3), arrParams)
                        Else
                            strResult = GetValue(mainClass(currentClassId).Names(0) + "[" + funcParams(0) + ", " + funcParams(1) + "]." + elementName + " - " + stringRemain.Substring(3), arrParams)
                        End If
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        ElseIf strResult.StartsWith(temporary_system_array_BASE_check) Then
                            LAST_ERROR = "Нельзя параметру присвоить массив."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    Else
                        LAST_ERROR = "Неверный оператор присвоения значения свойства или нераспознанная строка."
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    'Присваиваем значение свойству
                    If returnFormat <> MatewScript.ReturnFormatEnum.ORIGINAL Then strResult = ConvertElement(strResult, returnFormat)
                    If PropertiesRouter(currentClassId, elementName, funcParams, arrParams, PropertiesOperationEnum.PROPERTY_SET, strResult) = "#Error" Then
                        LAST_ERROR = "Не удалось присвоить значение свойству " + mainClass(currentClassId).Names(mainClass(currentClassId).Names.GetUpperBound(0)) + "." + elementName + "."
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                Case MatewScript.WordTypeEnum.W_BLOCK_IF
                    'Блок If ... Then ... Else If ... Else ...End If
                    strResult = GetValue(qbContent, arrParams) 'результат между If и Then
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If

                    If stringRemain.Length = 0 Then
                        'многострочный блок
                        'Получаем начало и конец исполняемого кода внутри блока, а также позицию End If
                        strResult = BlockIF(masCode, Convert.ToBoolean(strResult), lineId, blockFinalPos, blockFinalOperator, arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        If blockFinalPos = -1 Then 'нет кода для исполнения (ничего не подошло условиям)
                            lineId = blockFinalOperator + 1 'устанавливаем текущую позицию за End If
                            Continue Do
                        End If

                        strResult = ExecuteCode(masCode, arrParams, False, lineId, blockFinalPos) 'выполняем блок
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                            Return "#Error"
                        End If
                        If EXIT_CODE Then
                            If strResult = "#BREAK#" Then
                                LAST_ERROR = "Оператор Break нельзя использоваь в блоке If ... Then."
                                If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                Return "#Error"
                            Else
                                Return strResult
                            End If
                        End If

                        If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                            Dim markLine As Integer = GetLineAfterMark(masCode, markName)
                            If markLine = -1 Then
                                LAST_ERROR = "Метка " + markName + " не найдена."
                                AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            Else
                                If markLine >= lineId AndAlso markLine <= blockFinalPos Then
                                    'метка в пределах блока, причем только той части If Then, которая пдошла по условиям - выполняем код только внутри
                                    strResult = ExecuteCode(masCode, arrParams, False, markLine, blockFinalPos)
                                    Return strResult
                                ElseIf isPrimaryCode Then
                                    lineId = markLine
                                    Continue Do
                                Else
                                    Return strResult
                                End If
                            End If
                        End If

                        'If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                        '    If isPrimaryCode Then
                        '        stringRemain = strResult.Substring(6) 'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.
                        '        GoTo jmp
                        '    Else
                        '        Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                        '    End If
                        'End If

                        If strResult.StartsWith("#CONTINUE#") Then
                            If isPrimaryCode Then
                                LAST_ERROR = "Оператор Continue нельзя использовать за пределами цикла."
                                If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                Return "#Error"
                            Else
                                Return strResult
                            End If
                        End If

                        lineId = blockFinalOperator + 1 'устанавливаем текущую позицию за End If
                        Continue Do
                    Else
                        'блок однострочный
                        If strResult = "False" Then
                            lineId += 1
                            Continue Do
                        End If
                        'создаем блок кода на выполнение (из того, что за Then)
                        Dim externalBlock(0) As DictionaryEntry 'для создания подблока на выполнение из кода после Then 
                        externalBlock(0).Value = stringRemain
                        externalBlock(0).Key = masCode(lineId).Key
                        strResult = ExecuteCode(externalBlock, arrParams) 'Выполняем блок
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        If EXIT_CODE Then
                            If strResult = "#BREAK#" AndAlso isPrimaryCode Then
                                LAST_ERROR = "Оператор Break нельзя использовать в блоке If ... Then."
                                If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                Return "#Error"
                            Else
                                Return strResult
                            End If
                        End If

                        If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            If isPrimaryCode Then
                                Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                                Dim markLine As Integer = GetLineAfterMark(masCode, markName)
                                If markLine = -1 Then
                                    LAST_ERROR = "Метка " + markName + " не найдена."
                                    AnalizeError(masCode(lineId), arrParams)
                                    Return "#Error"
                                Else
                                    lineId = markLine
                                    Continue Do
                                End If
                            Else
                                Return strResult
                            End If
                        End If

                        'If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                        '    If isPrimaryCode Then
                        '        stringRemain = strResult.Substring(6) 'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.
                        '        GoTo jmp
                        '    Else
                        '        Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                        '    End If
                        'End If
                        If strResult.StartsWith("#CONTINUE#") Then
                            If isPrimaryCode Then
                                LAST_ERROR = "Оператор Continue нельзя использовать за пределами цикла."
                                If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                Return "#Error"
                            Else
                                Return strResult
                            End If
                        End If
                    End If
                Case MatewScript.WordTypeEnum.W_BLOCK_FOR
                    'For i = 1 To 10 Step 2
                    'i - elementName 
                    '"1" - qbContent, "10" - stringRemain 
                    '2 - позиция в строке strCodeString сохраняется в ElementId. Если Step нет, то = -1
                    'Получаем иедекс последней строки в блоке
                    strResult = BlockFor(masCode, lineId, blockFinalPos)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    If blockFinalPos = lineId Then 'пустой цикл
                        lineId += 2
                        Continue Do
                    End If
                    'получаем начальное значение переменной цикла
                    Dim minValue As Double
                    If Double.TryParse(qbContent, NumberStyles.Any, provider_points, minValue) = False Then
                        qbContent = GetValue(qbContent, arrParams)
                        If qbContent = "#Error" Then
                            LAST_ERROR = "В цикле For ... Next ошибка при расчете первого значения перменной."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        If Param_GetType(qbContent) <> MatewScript.ReturnFormatEnum.TO_NUMBER Then
                            LAST_ERROR = "В цикле For ... Next первое значение перменной  - не число."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        minValue = Convert.ToDouble(qbContent, provider_points)
                    End If

                    'получаем конечное значение переменной цикла
                    Dim maxValue As Double
                    If Double.TryParse(stringRemain, NumberStyles.Any, provider_points, maxValue) = False Then
                        stringRemain = GetValue(stringRemain, arrParams)
                        If stringRemain = "#Error" Then
                            LAST_ERROR = "В цикле For ... Next ошибка при расчете последнего значения перменной."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        If Param_GetType(stringRemain) <> MatewScript.ReturnFormatEnum.TO_NUMBER Then
                            LAST_ERROR = "В цикле For ... Next последнее значение перменной  - не число."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        maxValue = Convert.ToDouble(stringRemain, provider_points)
                    End If

                    'получаем шаг цикла
                    Dim stepValue As Double = 1
                    If ElementId > -1 Then
                        strResult = curExpression.Substring(ElementId)
                        If Double.TryParse(strResult, NumberStyles.Any, provider_points, stepValue) = False Then
                            strResult = GetValue(strResult, arrParams)
                            If strResult = "#Error" Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                            If Param_GetType(strResult) <> MatewScript.ReturnFormatEnum.TO_NUMBER Then
                                LAST_ERROR = "В цикле For ... Next параметр Step - не число."
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                            stepValue = Convert.ToDouble(strResult, provider_points)
                        End If
                    End If
                    'Получаем из Var.myVar[x] имя и индекс переменной (myVar & x)
                    Dim varName As String = "", varIndex As Integer, varSignature As String = ""
                    If SplitVarFromFullName(elementName, varName, varIndex, varSignature, returnFormat, arrParams) = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    If varSignature.Length > 0 Then
                        varSignature = mScript.PrepareStringToPrint(varSignature, arrParams)
                        If varSignature = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return varSignature
                        End If
                    End If

                    'получаем первую и последнюю строку на исполнение относительно начала кода
                    Dim blockStart As Integer = lineId + 1, blockFinish As Integer = blockFinalPos
                    'переменная локальная или глобальная?
                    Dim isPublicVariable As Boolean = False
                    Dim isExamVariable As Boolean = False
                    If csLocalVariables.GetVariable(varName, 0) = "#Error" Then
                        If csPublicVariables.GetVariable(varName, 0) <> "#Error" Then
                            isPublicVariable = True
                        End If
                    End If
                    'Выполняем цикл

                    For i = minValue To maxValue Step stepValue
                        'присваиваем переменной цикла новое значение
                        If isPublicVariable Then
                            Dim res As Integer
                            If varSignature.Length = 0 Then
                                res = csPublicVariables.SetVariableInternal(varName, i.ToString, varIndex, returnFormat)
                            Else
                                res = csPublicVariables.SetVariableInternal(varName, i.ToString, varSignature, returnFormat)
                            End If
                            If res = -1 Then
                                LAST_ERROR = "Не удалось присвоить переменной " + varName + "[" + varIndex.ToString + "]" + " значение " + i.ToString + "."
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        Else
                            Dim res As Integer
                            If varSignature.Length = 0 Then
                                res = csLocalVariables.SetVariableInternal(varName, i.ToString, varIndex, returnFormat)
                            Else
                                res = csLocalVariables.SetVariableInternal(varName, i.ToString, varSignature, returnFormat)
                            End If
                            If res = -1 Then
                                LAST_ERROR = "Не удалось присвоить переменной " + varName + "[" + varIndex.ToString + "]" + " значение " + i.ToString + "."
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        End If
                        'Выполняем блок кода
                        strResult = ExecuteCode(masCode, arrParams, False, blockStart, blockFinish) 'выполняем блок
fromMark2:
                        If strResult = "#Error" Then
                            LAST_ERROR = "Ошибка внутри цикла For ... Next. " + LAST_ERROR
                            If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                            Return "#Error"
                        End If
                        If EXIT_CODE Then
                            'Из кода получен сигнал о завершении, но это не связано с ошибкой (strResult <> "#Error")
                            'Это может быть Return, Exit или Break
                            If strResult.StartsWith("#BREAK#") Then
                                'Это Break
                                If strResult = "#BREAK#" Then
                                    EXIT_CODE = False 'Break без параметров - продолжаем работу
                                    Exit For
                                Else
                                    'Break с параметом = кол-во циклов, из которых надо выйти (напр., #BREAK#3 )
                                    Dim breaksLeft As Integer = Convert.ToInt32(strResult.Substring(7)) - 1 'отнимаем 1 от кол-ва циклов, из которых надо выйти
                                    If breaksLeft <= 0 Then
                                        EXIT_CODE = False 'больше ни откуда выходить не надо
                                        Exit For
                                    Else
                                        If isPrimaryCode Then
                                            LAST_ERROR = "Некорректное использование оператора Break. Больше нет циклов."
                                            If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                            Return "#Error"
                                        Else
                                            Return "#BREAK#" + breaksLeft.ToString 'выходим и передаем родительской функции ExecuteCode параметр на 1 меньше
                                        End If
                                    End If
                                End If
                            Else
                                Return strResult 'Это Return или Exit
                            End If
                        End If

                        If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                            Dim markLine As Integer = GetLineAfterMark(masCode, markName)
                            If markLine = -1 Then
                                LAST_ERROR = "Метка " + markName + " не найдена."
                                AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            Else
                                If markLine >= blockStart AndAlso markLine <= blockFinish Then
                                    'метка в пределах цикла - цикл продолжается
                                    strResult = ExecuteCode(masCode, arrParams, False, markLine, blockFinish)
                                    GoTo fromMark2
                                ElseIf isPrimaryCode Then
                                    lineId = markLine
                                    Continue Do
                                Else
                                    Return strResult
                                End If
                            End If
                        End If



                        'If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                        '    If isPrimaryCode Then
                        '        stringRemain = strResult.Substring(6) 'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.
                        '        GoTo jmp
                        '    Else
                        '        Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                        '    End If
                        'End If
                    Next i

                    lineId = blockFinalPos + 2 'устанавливаем текущую позицию за Next
                    Continue Do
                Case MatewScript.WordTypeEnum.W_FUNCTION
                    'получаем все параметры из [] и () в массив funcParams
                    funcParams = {}
                    funcCloseBracketPos = -1
                    If curExpression.EndsWith("(" + stringRemain) Then
                        'есть скобки, принадлежащие функции и прилежащие к ней
                        If GetFunctionParams("(" + stringRemain, funcParams, 0, funcCloseBracketPos, qbContent, finalBracketCount, arrParams) = -2 Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    Else
                        'скобок нет
                        If stringRemain.Length > 0 Then
                            'есть параметры за именем функции
                            If GetFunctionParams("(" + stringRemain + ")", funcParams, 0, funcCloseBracketPos, qbContent, finalBracketCount, arrParams) = -2 Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        Else
                            'параметры могут быть только в []
                            If GetFunctionParams("", funcParams, -1, funcCloseBracketPos, qbContent, finalBracketCount, arrParams) = -2 Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        End If
                    End If

                    'выполняем функцию
                    If FunctionRouter(currentClassId, elementName, funcParams, arrParams) = "#Error" Then
                        LAST_ERROR = "Не удалось выполнить функцию " + mainClass(currentClassId).Names(mainClass(currentClassId).Names.GetUpperBound(0)) + "." + elementName + ". " + LAST_ERROR
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                Case MatewScript.WordTypeEnum.W_SWITCH
                    'Блок Select Case / Switch
                    'Получаем начало и конец исполняемого кода внутри блока, а также позицию End Select
                    strResult = BlockSwitch(masCode, stringRemain, lineId, blockFinalPos, blockFinalOperator, arrParams)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    If blockFinalPos = -1 Then 'нет кода для исполнения (ничего не подошло условиям)
                        lineId = blockFinalOperator + 1 'устанавливаем текущую позицию за End Select
                        Continue Do
                    End If

                    strResult = ExecuteCode(masCode, arrParams, False, lineId, blockFinalPos) 'выполняем блок
                    If strResult = "#Error" Then
                        LAST_ERROR = "ошибка внутри блока Select Case. " + LAST_ERROR
                        If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                        Return "#Error"
                    End If
                    If EXIT_CODE Then
                        If strResult = "#BREAK#" Then
                            LAST_ERROR = "Оператор Break нельзя использоваь в блоке Select Case / Switch."
                            If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                            Return "#Error"
                        Else
                            Return strResult
                        End If
                    End If

                    If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                        If isPrimaryCode Then
                            Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                            Dim markLine As Integer = GetLineAfterMark(masCode, markName)
                            If markLine = -1 Then
                                LAST_ERROR = "Метка " + markName + " не найдена."
                                AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            Else
                                lineId = markLine
                                Continue Do
                            End If
                        Else
                            Return strResult
                        End If
                    End If

                    'If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                    '    If isPrimaryCode Then
                    '        stringRemain = strResult.Substring(6) 'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.
                    '        GoTo jmp
                    '    Else
                    '        Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                    '    End If
                    'End If

                    lineId = blockFinalOperator + 1 'устанавливаем текущую позицию за End Select
                    Continue Do
                Case MatewScript.WordTypeEnum.W_BLOCK_DOWHILE
                    'stringRemain - условие
                    'Получаем индекс последней строки в блоке
                    strResult = BlockDoWhile(masCode, lineId, blockFinalPos)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    If blockFinalPos = lineId Then 'пустой цикл
                        lineId += 2
                        Continue Do
                    End If
                    'получаем первую и последнюю строку на исполнение относительно начала кода
                    Dim blockStart As Integer = lineId + 1, blockFinish As Integer = blockFinalPos
                    'Выполняем цикл
                    Do While IIf(stringRemain.Length = 0, "True", GetValue(stringRemain, arrParams)) = "True"
                        strResult = ExecuteCode(masCode, arrParams, False, blockStart, blockFinish) 'выполняем блок
fromMark:
                        If strResult = "#Error" Then
                            LAST_ERROR = "Ошибка внутри цикла Do While. " + LAST_ERROR
                            If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                            Return "#Error"
                        End If
                        If EXIT_CODE Then
                            'Из кода получен сигнал о завершении, но это не связано с ошибкой (strResult <> "#Error")
                            'Это может быть Return, Exit или Break
                            If strResult.StartsWith("#BREAK#") Then
                                'Это Break
                                If strResult = "#BREAK#" Then
                                    EXIT_CODE = False 'Break без параметров - продолжаем работу
                                    Exit Do
                                Else
                                    'Break с параметом = кол-во циклов, из которых надо выйти (напр., #BREAK#3 )
                                    Dim breaksLeft As Integer = Convert.ToInt32(strResult.Substring(7)) - 1 'отнимаем 1 от кол-ва циклов, из которых надо выйти
                                    If breaksLeft <= 0 Then
                                        EXIT_CODE = False 'больше ни откуда выходить не надо
                                        Exit Do
                                    Else
                                        If isPrimaryCode Then
                                            LAST_ERROR = "Некорректное использование оператора Break. Больше нет циклов."
                                            If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                                            Return "#Error"
                                        Else
                                            Return "#BREAK#" + breaksLeft.ToString 'выходим и передаем родительской функции ExecuteCode параметр на 1 меньше
                                        End If
                                    End If
                                End If
                            Else
                                Return strResult 'Это Return или Exit
                            End If
                        End If

                        If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                            Dim markName As String = mScript.PrepareStringToPrint(strResult.Substring(6), Nothing, False)
                            Dim markLine As Integer = GetLineAfterMark(masCode, markName)
                            If markLine = -1 Then
                                LAST_ERROR = "Метка " + markName + " не найдена."
                                AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            Else
                                If markLine >= blockStart AndAlso markLine <= blockFinish Then
                                    'метка в пределах цикла - цикл продолжается
                                    strResult = ExecuteCode(masCode, arrParams, False, markLine, blockFinish)
                                    GoTo fromMark
                                ElseIf isPrimaryCode Then
                                    lineId = markLine
                                    Continue Do
                                Else
                                    Return strResult
                                End If
                            End If
                        End If

                        'If strResult.StartsWith("#JUMP#") Then 'в блоке, который только что выполнился, возник оператор Jump
                        '    If isPrimaryCode Then
                        '        stringRemain = strResult.Substring(6) 'первичновызванная функция - значит здесь надо искать метку. Переходим к поиску.
                        '        GoTo jmp
                        '    Else
                        '        Return strResult 'код не первичный - переходим на уровень выше (к более раннему или первичному)
                        '    End If
                        'End If
                    Loop
                    If LAST_ERROR.Length > 0 Then
                        LAST_ERROR = "ошибка в условии цикла Do While. " + LAST_ERROR
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    lineId = blockFinalPos + 2 'устанавливаем текущую позицию за Loop
                    Continue Do
                Case MatewScript.WordTypeEnum.W_BLOCK_FUNCTION
                    'stringRemain - имя функции
                    'Получаем индекс последней строки в блоке
                    strResult = BlockFunction(masCode, lineId, blockFinalPos)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    If blockFinalPos = lineId Then 'пустой блок
                        lineId += 2
                        Continue Do
                    End If
                    'получаем первую и последнюю строку на исполнение относительно начала кода
                    Dim blockStart As Integer = lineId + 1, blockFinish As Integer = blockFinalPos
                    stringRemain = GetValue(stringRemain, arrParams)
                    If stringRemain = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    'создаем подмассив, содержащий код фунции
                    Dim subArray() As DictionaryEntry
                    ReDim subArray(blockFinish - blockStart)
                    Array.ConstrainedCopy(masCode, blockStart, subArray, 0, blockFinish - blockStart + 1)
                    Dim lineDifference As Long = subArray(0).Key - 1
                    Dim sb As New System.Text.StringBuilder
                    For i = 0 To subArray.GetUpperBound(0)
                        subArray(i).Key -= lineDifference
                        sb.AppendLine(subArray(i).Value)
                    Next
                    'создаем исполняемый код
                    'questEnvironment.codeBoxShadowed.codeBox.IsTextBlockByDefault = False
                    'questEnvironment.codeBoxShadowed.Text = sb.ToString
                    Dim exData As New List(Of ExecuteDataType)
                    exData = PrepareBlock(mScript.eventRouter.ConverTextToCodeData(sb.ToString))
                    'exData = PrepareBlock(questEnvironment.codeBoxShadowed.codeBox.CodeData)
                    'questEnvironment.codeBoxShadowed.Text = ""

                    'получаем имя функции
                    Dim funcName As String = stringRemain
                    Dim fClassId As Integer = -1
                    Dim isClassFunction As Boolean = GetClassIdAndElementNameByString(stringRemain, fClassId, funcName, True, False, True)
                    stringRemain = PrepareStringToPrint(stringRemain, arrParams)

                    If isClassFunction Then
                        'Это функция класса, добавленная Писателем. Вносим соответствующие изменения в структуру mainClass и создаем событие
                        If mainClass(fClassId).Functions.ContainsKey(funcName) = False Then
                            If AddUserFunction({mainClass(fClassId).Names(0), stringRemain}, arrParams) = "#Error" Then
                                'Если указана пока еще не созданная функция пользователя - создаем ее
                                LAST_ERROR = "Неверный синтаксис блока Function. Не удалось зодать новую функцию " + funcName + " в классе " + mainClass(fClassId).Names(0)
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        End If

                        Dim func As PropertiesInfoType = mainClass(fClassId).Functions(funcName)
                        func.Value = "" 'создавать серализованный аналог необязательно
                        func.eventId = eventRouter.SetEventId(func.eventId, exData) 'создаем/заменяем событие
                    Else
                        'Это обычная функция
                        Dim f As FunctionInfoType
                        If IsNothing(functionsHash) OrElse functionsHash.ContainsKey(stringRemain) = False Then
                            'Добавляем новую функцию
                            f = New FunctionInfoType
                            f.ValueExecuteDt = exData
                            functionsHash.Add(stringRemain, f)
                        Else
                            'Замещаем старую функцию новой
                            f = functionsHash(stringRemain)
                            If IsNothing(f.Variables) = False Then f.Variables.Clear()
                            If IsNothing(f.ValueDt) = False Then Erase f.ValueDt
                            If IsNothing(f.ValueExecuteDt) = False Then f.ValueExecuteDt.Clear()
                            f.ValueExecuteDt = exData
                        End If
                        If IsNothing(csLocalVariables.lstVariables) = False AndAlso csLocalVariables.lstVariables.Count > 0 Then csLocalVariables.CopyVariables(f.Variables)
                    End If

                    lineId = blockFinalPos + 2 'устанавливаем текущую позицию за Loop
                    Continue Do
                Case MatewScript.WordTypeEnum.W_PARAM
                    'присвоение значения параметру, с которыми запущен код
                    'если есть [], получаем результат в них: Param[2 + 2] = x, qbContent = 4
                    If qbContent.Length > 0 Then
                        If Integer.TryParse(qbContent, Nothing) = False Then qbContent = GetValue(qbContent, arrParams)
                        If qbContent = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        If qbContent < 0 Then
                            LAST_ERROR = "Индекс элемента массива Param получился меньше 0."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    Else
                        qbContent = "0"
                    End If

                    If IsNothing(arrParams) OrElse arrParams.GetUpperBound(0) < Convert.ToInt32(qbContent) Then
                        LAST_ERROR = "Попытка присвоить значение несуществующему параметру. Param[" + qbContent + "] не существует."
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If

                    'Получаем результат, который надо присвоить Param, с учетом оператора (=, += или -=)
                    If stringRemain.StartsWith("= ") Then
                        strResult = GetValue(stringRemain.Substring(2), arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        ElseIf strResult.StartsWith(temporary_system_array_BASE_check) Then
                            LAST_ERROR = "Нельзя параметру присвоить массив."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    ElseIf stringRemain.StartsWith("+= ") Then
                        strResult = GetValue("Param[" + qbContent + "] + " + stringRemain.Substring(3), arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        ElseIf strResult.StartsWith(temporary_system_array_BASE_check) Then
                            LAST_ERROR = "Нельзя параметру присвоить массив."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    ElseIf stringRemain.StartsWith("-= ") Then
                        strResult = GetValue("Param[" + qbContent + "] - " + stringRemain.Substring(3), arrParams)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        ElseIf strResult.StartsWith(temporary_system_array_BASE_check) Then
                            LAST_ERROR = "Нельзя параметру присвоить массив."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    Else
                        LAST_ERROR = "Неверный оператор присвоения значения Param или нераспознанная строка."
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If

                    'Присваиваем значение Param
                    arrParams(Convert.ToInt32(qbContent)) = ConvertElement(strResult, returnFormat)
                Case MatewScript.WordTypeEnum.W_RETURN
                    If stringRemain.Length = 0 Then
                        csLocalVariables.KillVars()
                        EXIT_CODE = True
                        Return ""
                    Else
                        If Integer.TryParse(stringRemain, Nothing) = False Then
                            stringRemain = GetValue(stringRemain, arrParams)
                            If stringRemain = "#Error" Then
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                EXIT_CODE = True
                                Return "#Error"
                            End If
                        End If
                        csLocalVariables.KillVars()
                        EXIT_CODE = True
                        Return stringRemain
                    End If
                Case MatewScript.WordTypeEnum.W_EXIT
                    EXIT_CODE = True
                    csLocalVariables.KillVars()
                    Return ""
                Case MatewScript.WordTypeEnum.W_JUMP
                    If stringRemain.Trim.Length = 0 Then
                        LAST_ERROR = "За ключевым словом Jump не стоит метка."
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    stringRemain = PrepareStringToPrint(stringRemain, arrParams)
                    If isPrimaryCode Then
                        lineId = GetLineAfterMark(masCode, strResult)
                        If lineId = -1 Then
                            LAST_ERROR = "Метка " + stringRemain + " не найдена."
                            Return "#Error"
                        Else
                            Continue Do
                        End If
                    Else
                        Return "#JUMP#" + stringRemain
                    End If


                    '                    Dim markPos As Integer = 0
                    '                    stringRemain += ":"
                    '                    If isPrimaryCode Then
                    'jmp:
                    '                        While markPos < masCode.GetUpperBound(0)
                    '                            If masCode(markPos).Value = stringRemain Then
                    '                                lineId = markPos + 1
                    '                                Continue Do
                    '                            End If
                    '                            markPos += 1
                    '                        End While
                    '                        LAST_ERROR = "Метка " + stringRemain + " не найдена."
                    '                        AnalizeError(masCode(lineId), arrParams)
                    '                        Return "#Error"
                    '                    Else
                    '                        Return "#JUMP#" + stringRemain
                    '                    End If
                Case MatewScript.WordTypeEnum.W_CONTINUE
                    If isPrimaryCode Then
                        LAST_ERROR = "Оператор Continue не может применяться вне цикла."
                        AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    Else
                        Return "#CONTINUE#"
                    End If
                Case MatewScript.WordTypeEnum.W_BREAK
                    'Оператор Break выходит из циклов For...Next & Do While.
                    'Может содержать параметр, из скольких циклов выйти
                    If isPrimaryCode Then
                        LAST_ERROR = "Оператор Break не может применяться вне цикла."
                        AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    Else
                        EXIT_CODE = True
                        If stringRemain.Length = 0 Then
                            Return "#BREAK#"
                        Else
                            If Integer.TryParse(stringRemain, Nothing) = False Then
                                EXIT_CODE = False
                                stringRemain = GetValue(stringRemain, arrParams)
                                EXIT_CODE = True
                                If stringRemain = "#Error" Then Return "#Error"
                                If Param_GetType(stringRemain) <> MatewScript.ReturnFormatEnum.TO_NUMBER Then
                                    LAST_ERROR = "За оператором Break может стоять только число."
                                    Return "#Error"
                                End If
                            End If
                            Return "#BREAK#" + stringRemain
                        End If
                    End If
                Case MatewScript.WordTypeEnum.W_CYCLE_INTERNAL
                    LAST_ERROR = "Обнаружен оператор блока (" + curExpression + ") вне этого блока. Возможно, неправильное размещение метки."
                    If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                    Return "#Error"
                Case MatewScript.WordTypeEnum.W_HTML
                    If questEnvironment.EDIT_MODE Then Continue Do
                    'Вставляет текст в текущий HTML-документ.
                    'HtmlDocument [Append] [parentId]
                    Dim isAppend As Boolean = False  ' добавить текст или заменить?
                    If stringRemain.StartsWith("Append ") OrElse stringRemain.StartsWith("Append" + Chr(1)) Then
                        isAppend = True
                        stringRemain = stringRemain.Substring(7)
                    End If
                    'получаем id элемента, куда вставлять текст (если он указан)
                    Dim containerId As String = ""
                    If stringRemain.StartsWith(Chr(1)) Then
                        stringRemain = stringRemain.Substring(1)
                    Else
                        blockFinalPos = stringRemain.IndexOf(Chr(1))
                        If blockFinalPos > -1 Then
                            containerId = PrepareStringToPrint(stringRemain.Substring(0, blockFinalPos), arrParams)
                            stringRemain = stringRemain.Substring(blockFinalPos + 1)
                        End If
                    End If
                    'Вставляем текст
                    stringRemain = MakeExec(stringRemain, arrParams) 'выполняем тэги <exec>

                    If isAppend = False And containerId.Length = 0 Then
                        'HTML
                        Dim hDoc As HtmlDocument = frmPlayer.wbMain.Document
                        If IsNothing(hDoc) = False Then
                            Dim hEl As HtmlElement = hDoc.GetElementById("MainConvas")
                            If IsNothing(hEl) = False Then hEl.InnerHtml = ""
                        End If
                        PrintTextToMainWindow("", stringRemain, PrintInsertionEnum.APPEND_NEW_BLOCK, "", PrintDataEnum.TEXT, arrParams)

                    Else
                        If containerId.Length = 0 Then
                            'HTML Append
                            hCurrentDocument.Write(stringRemain)
                            'Dim hChild As HtmlElement = hCurrentDocument.CreateElement("P")
                            'hChild.InnerHtml = stringRemain
                            'hCurrentDocument.Body.AppendChild(hChild)
                        Else
                            'HTML 'elementId' или HTML Append 'elementId'
                            Dim hParent As HtmlElement
                            hParent = hCurrentDocument.GetElementById(containerId)
                            If IsNothing(hParent) Then
                                LAST_ERROR = "HTML-элемента с id = " + containerId + " не найдено."
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                            hParent.InnerHtml = IIf(isAppend, hParent.InnerHtml + stringRemain, stringRemain)

                            If containerId.Length = 0 Then
                                'HTML Append
                                PrintTextToMainWindow("", stringRemain, PrintInsertionEnum.APPEND_NEW_BLOCK, "", PrintDataEnum.TEXT, arrParams)
                            Else
                                'HTML 'elementId' или HTML Append 'elementId'
                                If isAppend Then
                                    PrintTextToMainWindow(containerId, stringRemain, PrintInsertionEnum.APPEND, "", PrintDataEnum.TEXT, arrParams)
                                Else
                                    PrintTextToMainWindow(containerId, stringRemain, PrintInsertionEnum.REPLACE, "", PrintDataEnum.TEXT, arrParams)
                                End If
                            End If

                        End If
                    End If

                    returnFormat = ReturnFormatEnum.ORIGINAL
                    lineId = blockFinalPos + 2
                    Continue Do
                Case MatewScript.WordTypeEnum.W_BLOCK_EVENT 'блок Event
                    'Event 'propEventName'[, par1, par2]
                    'stringRemain - 'propEventName'[, par1, par2]
                    'Получаем индекс последней строки в блоке
                    strResult = BlockEvent(masCode, lineId, blockFinalPos)
                    If strResult = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    'получаем параметры блока event, первый из которых - имя свойства-события
                    If GetFunctionParams("", funcParams, -1, funcCloseBracketPos, stringRemain, finalBracketCount, arrParams) = -2 Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    If IsNothing(funcParams) OrElse funcParams.GetUpperBound(0) = -1 Then
                        LAST_ERROR = "Неправильная запись блока Event. Отсутствует имя события."
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    'получаем имя свойства-события
                    elementName = PrepareStringToPrint(funcParams(0), arrParams)
                    If elementName = "#Error" Then
                        LAST_ERROR = "Неправильная запись блока Event. При получении имени события возникла ошибка: " + LAST_ERROR
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If

                    'событие отслеживания свойства?
                    Dim isTracking As frmMainEditor.trackingcodeEnum = frmMainEditor.trackingcodeEnum.NOT_TRACKING_EVENT
                    If elementName.EndsWith(":changed") Then
                        isTracking = frmMainEditor.trackingcodeEnum.EVENT_AFTER
                        elementName = elementName.Substring(0, elementName.Length - 8)
                    ElseIf elementName.EndsWith(":changing") Then
                        isTracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE
                        elementName = elementName.Substring(0, elementName.Length - 9)
                    End If

                    'получаем класс, в котором находится это свойство-событие
                    currentClassId = -1
                    Dim cPos As Integer = elementName.IndexOf("."c)
                    If cPos > -1 AndAlso cPos < elementName.Length - 1 Then
                        '"L.EventOnEnter"
                        Dim className As String = elementName.Substring(0, cPos)
                        elementName = elementName.Substring(cPos + 1)
                        If mainClassHash.TryGetValue(className, currentClassId) = False Then
                            LAST_ERROR = "Неправильная запись блока Event. При получении имени события возникла ошибка: " + LAST_ERROR
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    Else
                        '"EventOnEnter"
                        For i As Integer = 0 To mainClass.GetUpperBound(0)
                            If mainClass(i).Properties.TryGetValue(elementName, Nothing) Then
                                currentClassId = i
                                Exit For
                            End If
                        Next
                    End If

                    If currentClassId = -1 Then
                        LAST_ERROR = "Не найдено событие " + elementName + "."
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    'получаем параметры без имени свойства-события
                    If funcParams.GetUpperBound(0) = 0 Then
                        Erase funcParams
                    Else
                        For i = 0 To funcParams.GetUpperBound(0) - 1
                            funcParams(i) = funcParams(i + 1)
                        Next
                        ReDim Preserve funcParams(funcParams.GetUpperBound(0) - 1)
                    End If
                    'получаем в строку весь код, который надо вставить в событие
                    Dim eventContent As New System.Text.StringBuilder
                    For i = lineId + 1 To blockFinalPos
                        eventContent.AppendLine(masCode(i).Value)
                    Next

                    Dim cd() As CodeTextBox.CodeDataType = mScript.eventRouter.ConverTextToCodeData(eventContent.ToString)
                    If IsNothing(cd) Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If

                    If isTracking = frmMainEditor.trackingcodeEnum.EVENT_BEFORE Then
                        mScript.trackingProperties.AddPropertyBefore(currentClassId, elementName, cd)
                    ElseIf isTracking = frmMainEditor.trackingcodeEnum.EVENT_AFTER Then
                        mScript.trackingProperties.AddPropertyAfter(currentClassId, elementName, cd)
                    Else
                        mScript.eventRouter.SetPropertyWithEvent(currentClassId, elementName, cd, funcParams, "", False, False, questEnvironment.EDIT_MODE)
                    End If

                    'If PropertiesRouter(currentClassId, elementName, funcParams, arrParams, PropertiesOperationEnum.PROPERTY_SET, eventContent.ToString) = "#Error" Then
                    '    If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                    '    Return "#Error"
                    'End If

                    lineId = blockFinalPos + 2 'устанавливаем текущую позицию за Loop
                    Continue Do
                Case MatewScript.WordTypeEnum.W_WRAP
                    'elementName - имя переменной
                    'stringRemain - текст блока
                    'Получаем из Var.myVar[x] имя и индекс переменной (myVar & x)
                    Dim varName As String = "", varIndex As Integer, varSignature As String = ""
                    If SplitVarFromFullName(elementName, varName, varIndex, varSignature, returnFormat, arrParams) = "#Error" Then
                        'Class[x,y].Property(x,y)

                        strResult = BlockWrap(masCode, lineId, blockFinalPos)
                        If strResult = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        'получаем параметры блока wrap, первый из которых - имя свойства-события
                        If GetFunctionParams("", funcParams, -1, funcCloseBracketPos, stringRemain, finalBracketCount, arrParams) = -2 Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        If IsNothing(funcParams) OrElse funcParams.GetUpperBound(0) = -1 Then
                            LAST_ERROR = "Неправильная запись блока Wrap. Отсутствует имя события."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        'получаем имя свойства-события
                        elementName = PrepareStringToPrint(funcParams(0), arrParams)
                        If elementName = "#Error" Then
                            LAST_ERROR = "Неправильная запись блока Wrap. При получении имени события возникла ошибка: " + LAST_ERROR
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If

                        'получаем класс, в котором находится это свойство-событие
                        currentClassId = -1
                        Dim cPos As Integer = elementName.IndexOf("."c)
                        If cPos > -1 AndAlso cPos < elementName.Length - 1 Then
                            '"L.EventOnEnter"
                            Dim className As String = elementName.Substring(0, cPos)
                            elementName = elementName.Substring(cPos + 1)
                            If mainClassHash.TryGetValue(className, currentClassId) = False Then
                                LAST_ERROR = "Неправильная запись блока Wrap. При получении имени события возникла ошибка: " + LAST_ERROR
                                If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                                Return "#Error"
                            End If
                        Else
                            '"EventOnEnter"
                            For i As Integer = 0 To mainClass.GetUpperBound(0)
                                If mainClass(i).Properties.TryGetValue(elementName, Nothing) Then
                                    currentClassId = i
                                    Exit For
                                End If
                            Next
                        End If

                        If currentClassId = -1 Then
                            LAST_ERROR = "Не найдено событие " + elementName + "."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                        'получаем параметры без имени свойства-события
                        If funcParams.GetUpperBound(0) = 0 Then
                            Erase funcParams
                        Else
                            For i = 0 To funcParams.GetUpperBound(0) - 1
                                funcParams(i) = funcParams(i + 1)
                            Next
                            ReDim Preserve funcParams(funcParams.GetUpperBound(0) - 1)
                        End If
                        'получаем в строку весь код, который надо вставить в событие
                        Dim eventContent As New System.Text.StringBuilder
                        For i = lineId + 1 To blockFinalPos
                            eventContent.AppendLine(masCode(i).Value)
                        Next

                        Dim cd() As CodeTextBox.CodeDataType = mScript.eventRouter.ConverTextToCodeData(eventContent.ToString)
                        If IsNothing(cd) Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If

                        mScript.eventRouter.SetPropertyWithEvent(currentClassId, elementName, cd, funcParams, "", False, False, questEnvironment.EDIT_MODE)

                        lineId = blockFinalPos + 2 'устанавливаем текущую позицию за Loop
                        Continue Do

                    End If
                    If varSignature.Length > 0 Then
                        varSignature = mScript.PrepareStringToPrint(varSignature, arrParams)
                        If varSignature = "#Error" Then
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return varSignature
                        End If
                    End If
                    'переменная локальная или глобальная?
                    Dim isPublicVariable As Boolean = False

                    If csLocalVariables.GetVariable(varName, 0) = "#Error" Then
                        If csPublicVariables.GetVariable(varName, 0) <> "#Error" Then
                            isPublicVariable = True
                        End If
                    End If

                    stringRemain = MakeExec(stringRemain, arrParams) 'выполняем тэги <exec>
                    stringRemain = "'" + stringRemain.Replace("'", "/'") + "'" 'экранируем текст
                    If isPublicVariable Then
                        Dim res As Integer = -1
                        If varSignature.Length = 0 Then
                            res = csPublicVariables.SetVariableInternal(varName, stringRemain, varIndex, returnFormat)
                        Else
                            res = csPublicVariables.SetVariableInternal(varName, stringRemain, varSignature, returnFormat)
                        End If
                        If res = -1 Then
                            LAST_ERROR = "Не удалось присвоить переменной " + varName + "[" + varIndex.ToString + "]" + " значение " + stringRemain + "."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    Else
                        Dim res As Integer = -1
                        If varSignature.Length = 0 Then
                            res = csLocalVariables.SetVariableInternal(varName, stringRemain, varIndex, returnFormat)
                        Else
                            res = csLocalVariables.SetVariableInternal(varName, stringRemain, varSignature, returnFormat)
                        End If
                        If res = -1 Then
                            LAST_ERROR = "Не удалось присвоить переменной " + varName + "[" + varIndex.ToString + "]" + " значение " + stringRemain + "."
                            If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                            Return "#Error"
                        End If
                    End If
                Case MatewScript.WordTypeEnum.W_BLOCK_NEWCLASS
                    'добавляем новый класс пользователя
                    If CreateUserClass(masCode, stringRemain, lineId, blockFinalPos, arrParams) = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(CURRENT_LINE), arrParams)
                        Return "#Error"
                    End If
                    lineId = blockFinalPos + 1
                    Continue Do
                Case MatewScript.WordTypeEnum.W_REM_CLASS
                    'удаляем класс пользователя
                    stringRemain = GetValue(stringRemain, arrParams)
                    If stringRemain = "#Error" Then
                        If isPrimaryCode Then AnalizeError(masCode(lineId), arrParams)
                        Return "#Error"
                    End If
                    stringRemain = PrepareStringToPrint(stringRemain, arrParams)
                    RemoveUserClass(stringRemain)
            End Select
            lineId += 1 'переходим к новой строке кода и продолжаем цикл
        Loop

        If isPrimaryCode Then csLocalVariables.KillVars()
        Return ""
    End Function

    ''' <summary> Функция принимает блок кода в формате <see cref="CodeTextBox.CodeDataType"></see> и превращает ее в 
    ''' список классов <see cref="ExecuteDataType"></see>. При этом выполняется следующее:
    '''1) Объединяется код с _
    '''2) Разделяестся код с ;</summary>
    ''' <param name="srcCodeData">массив кода в формате <see cref="CodeTextBox.CodeDataType"></see></param>
    ''' <returns>список готового к исполнению кодав формате <see cref="ExecuteDataType"></see></returns>
    ''' <remarks></remarks>
    Public Overridable Function PrepareBlock(ByRef srcCodeData() As CodeTextBox.CodeDataType) As List(Of ExecuteDataType)
        If IsNothing(srcCodeData) OrElse srcCodeData.Count = 0 Then Return Nothing
        Dim CodeData() As CodeTextBox.CodeDataType = CopyCodeDataArray(srcCodeData)
        Dim curCode() As CodeTextBox.EditWordType 'для текущей линии кода
        Dim tempCode() As CodeTextBox.EditWordType = Nothing  'временнаЯ структура для разбивки кода на строки после ;
        Dim exList As New List(Of ExecuteDataType) 'список-результат функции
        Dim posStart As Integer 'номер фрагмента кода, следующего после последнего найденного ; (или 0, если ; не найдено)

        'выполняем для свех строк в CodeData
        For i As Integer = 0 To CodeData.GetUpperBound(0)
            curCode = CodeData(i).Code
            If IsNothing(curCode) OrElse curCode.Count = 0 Then Continue For

            If curCode(curCode.Count - 1).wordType = CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                'символ переноса строки _
                If i = CodeData.GetUpperBound(0) OrElse IsNothing(CodeData(i + 1).Code) = True Then
                    'за ним (на новой строке) нет кода - ошибка
                    LAST_ERROR = "За символом переноса строки не найдено никакого кода."
                    AnalizeError(New ExecuteDataType(curCode, i), Nothing)
                    Return Nothing
                End If
                'копируем в curCode всю следующую строку, замещая символ _ в конце. Получается полная строка кода без переноса
                Dim insertPos As Integer = curCode.Length - 1
                If CodeData(i + 1).Code.Length > 1 Then Array.Resize(curCode, curCode.Length + CodeData(i + 1).Code.Length - 1)
                Array.Copy(CodeData(i + 1).Code, 0, curCode, insertPos, CodeData(i + 1).Code.Length)
                'ставим полную строку без переноса на новую строку и продолжаем цикл с начала
                CodeData(i + 1).Code = curCode
                Continue For
            End If

            'ищем ;
            posStart = 0
            For j As Integer = 0 To curCode.Count - 1
                If curCode(j).wordType = CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION Then
                    'найден символ разбивки строк ;
                    'добавляем в exList код до этого символа
                    Array.Resize(tempCode, j - posStart)
                    Array.Copy(curCode, posStart, tempCode, 0, j - posStart)
                    exList.Add(New ExecuteDataType(tempCode, i + 1))
                    'сохраняем положение первого символа за ; на случай, если разбивок ; несколько
                    posStart = j + 1
                ElseIf curCode(j).wordType = CodeTextBox.EditWordTypeEnum.W_ERROR Then
                    'запущена строка с ошибкой - выход
                    LAST_ERROR = "Запущен на исполнение код с ошибкой синтаксиса."
                    AnalizeError(New ExecuteDataType(curCode, i), Nothing)
                    Return Nothing
                End If
            Next

            If posStart = 0 Then
                'добавляем всю строку кода в exList, если не было разбивок ;
                exList.Add(New ExecuteDataType(curCode, i + 1))
            ElseIf curCode.Length <> posStart Then
                'добавляем в exList последний кусок кода, после последнего ;
                Array.Resize(tempCode, curCode.Length - posStart)
                Array.Copy(curCode, posStart, tempCode, 0, curCode.Count - posStart)
                exList.Add(New ExecuteDataType(tempCode, i + 1))
            End If
        Next

        Return exList
    End Function


    ''' <summary>Функция принимает блок кода (строку) и превращает ее в массив DictionaryEntry. При этом выполняется следующее:
    '''1) Объединяются строки с _
    '''2) Разделяются строки с ;
    '''3) Убираются табуляторы
    '''4) Убираются комментарии //
    '''5) Блоки Text/HTML собираются в одной ячейке массива. При этом конечный End Text/HTML убирается
    '''Результирующий массив содержит ключи - номера строк, значения - строки чистого кода</summary>
    ''' <param name="arrBlock">массив кода, разбитого по строкам</param>
    ''' <returns>массив готовых к исполнению строк кода</returns>
    ''' <remarks></remarks>
    Public Overridable Function PrepareBlock(ByVal arrBlock() As String) As DictionaryEntry()
        'Получаем массив кода, разбитого по строкам
        'Dim arrBlock() As String = strBlock.Split(vbNewLine)
        'Объявляем наш конечный массив DictionaryEntry и переменную с его длиной
        Dim block() As DictionaryEntry = {}, blockLength As Integer = -1
        Dim curIndex As Integer = 0 'Номер текущей строки
        Dim curString As String 'Строка, с которой работаем
        'Переменные для хранения положения последнего найденного символа ' // ;
        Dim qPos As Integer = -1
        Dim commPos As Integer = -1
        Dim semiPos As Integer = -1
        'Позиция End Text/HTML
        Dim htmlEndPos As Integer
        'Счетчики и переменные общего назначения
        Dim i As Integer
        Dim strRes As String

        'Главный цикл. Выполняется для каждой строки кода
        Do While i < arrBlock.Length
            'Получаем текущую строку (без начальных и конечных пробелов)
            curString = arrBlock(i).Trim
            'Строка пустая - переходим к новой строке
            If curString.Length = 0 Then
                curIndex += 1 'Увеличиваем номер текущей строки
                i += 1
                Continue Do
            End If

            'Убираем комментарии
            commPos = curString.IndexOf("//") 'Ищем комментарий
            If commPos > -1 Then
                '// найден
                If commPos = 0 Then
                    'Комментарий на всю строку - переход к следующей
                    curIndex += 1 'Увеличиваем номер текущей строки
                    i += 1
                    Continue Do
                End If
                'Символы // могут быть в кавычках. Поэтому ищем последнюю кавычку
                qPos = curString.LastIndexOf("'")
                If qPos = -1 OrElse qPos < commPos Then
                    'Если кавычек нет или они закрываются до символов //, значит это точно комментарий
                    curString = curString.Remove(commPos).TrimEnd  'Убираем комментарий из строки и возможные пробелы в конце кода
                Else
                    'Кавычки есть после //. Значит, или кавычка закомментирована (//ss'ss), или комментарии в кавычках('asd//sas'), или и то, и другое
                    Do While commPos > -1 'Выполняем цикл, пока неясной природы // не закончатся
                        strRes = curString.Remove(commPos) 'Получаем строку найденных досимволов //
                        If IsQuotesCountEven(strRes) Then
                            'До // четное количество кавычек. Значит, это точно комментарий
                            curString = strRes.TrimEnd 'Убираем комментарий из строки и возможные пробелы в конце кода
                            Exit Do 'Выходим из цикла
                        End If
                        commPos = curString.IndexOf("//", commPos + 2) 'Эти символы // в кавычках. Ищем дальше
                    Loop
                End If
            End If

            'Обрабатываем блок HTML
            If curString = "HTML" OrElse curString.StartsWith("HTML ") Then
                'Ищем End Text/HTML
                htmlEndPos = Array.FindIndex(Of String)(arrBlock, i + 1, AddressOf PrepareTextSearch)
                If htmlEndPos = -1 Then 'Не нашли - ошибка
                    '#Error
                    Return Nothing
                End If
                curIndex += 1 'Увеличиваем номер текущей строки
                'Расширяем массив block
                blockLength += 1
                ReDim Preserve block(blockLength)
                block(blockLength).Key = curIndex 'Сохраняем текущий номер строки
                'Помещаем в новую ячейку все содержимое блока Text/HTML, без конечного End Text/HTML

                Dim subArray() As String
                ReDim subArray(htmlEndPos - i - 2)
                Array.ConstrainedCopy(arrBlock, i + 1, subArray, 0, htmlEndPos - i - 1)
                block(blockLength).Value = curString + Chr(1) + String.Join(vbNewLine, subArray)

                curIndex += htmlEndPos - i 'Увеличиваем номер текущей строки
                i = htmlEndPos + 1 'Перемещаем указатель текущей позиции на строку, следующую за блоком
                Continue Do
            ElseIf curString.StartsWith("Wrap ") Then
                'Обрабатываем блок Wrap
                'Ищем End Wrap
                htmlEndPos = Array.FindIndex(Of String)(arrBlock, i + 1, AddressOf PrepareWrapSearch)
                If htmlEndPos = -1 Then 'Не нашли - ошибка
                    '#Error
                    Return Nothing
                End If
                curIndex += 1 'Увеличиваем номер текущей строки
                'Расширяем массив block
                blockLength += 1
                ReDim Preserve block(blockLength)
                block(blockLength).Key = curIndex 'Сохраняем текущий номер строки
                'Помещаем в новую ячейку все содержимое блока Text/HTML, без конечного End Text/HTML

                Dim subArray() As String
                ReDim subArray(htmlEndPos - i - 2)
                Array.ConstrainedCopy(arrBlock, i + 1, subArray, 0, htmlEndPos - i - 1)
                block(blockLength).Value = curString + Chr(1) + String.Join(vbNewLine, subArray)

                curIndex += htmlEndPos - i 'Увеличиваем номер текущей строки
                i = htmlEndPos + 1 'Перемещаем указатель текущей позиции на строку, следующую за блоком
                Continue Do
            End If

            'Обрабатываем символ объединения строк _
            If curString.Last = "_"c Then
                If i = arrBlock.Length - 1 Then 'Найден в последней строке - ошибка
                    '#Error
                    Return Nothing
                End If
                curIndex += 1 'Увеличиваем номер текущей строки
                i += 1 'Перемещаем указатель текущей позиции
                curString = curString.Remove(curString.Length - 1).TrimEnd  'удаляем из строки " _"
                If curString.EndsWith("(") OrElse curString.EndsWith("[") Then
                    arrBlock(i) = curString + arrBlock(i).TrimStart 'Помещаем в следующую строку текущую + следующую
                Else
                    arrBlock(i) = curString + " " + arrBlock(i).TrimStart 'Помещаем в следующую строку текущую + " " + следующую
                End If
                Continue Do
            End If

            'Обрабатываем символ разделения строк
            semiPos = 0
semicolon:
            semiPos = curString.IndexOf(";", semiPos) 'Поиск символа ; в текущей строке, начиная с предыдущей позиции (первый поиск - с начала строки)
            If semiPos > 0 Then
                'Символ ; есть
                strRes = curString.Remove(semiPos) 'Получаем строку до ;
                If IsQuotesCountEven(strRes) Then 'Если в ней четное количество строк, значит символ ; функциональный
                    curIndex += 1 'Увеличиваем номер текущей строки
                    'Расширяем массив block
                    blockLength += 1
                    ReDim Preserve block(blockLength)
                    block(blockLength).Key = curIndex 'Сохраняем текущий номер строки
                    block(blockLength).Value = strRes.TrimEnd 'Записываем строку до символа ;
                    arrBlock(i) = curString.Substring(semiPos + 1) 'Убираем из текущей ячейки массива arrBlock часть строки до ; (с ней будем работать дальше)
                    curIndex -= 1 'Возвращаем номер текущей строки назад (фактически работаем с той же самой строкой)
                    Continue Do 'Новый виток цикла (будет обработана та же строка, но без части до ;)
                Else 'Если в ней НЕчетное количество строк, значит символ ; находится в кавычках
                    semiPos += 1 'Переходим к следующему символу и ищем заново
                    GoTo semicolon
                End If
            End If

final:
            curIndex += 1 'Увеличиваем номер текущей строки
            'Расширяем массив block
            blockLength += 1
            ReDim Preserve block(blockLength)
            block(blockLength).Key = curIndex 'Сохраняем текущий номер строки
            block(blockLength).Value = curString 'Сохраняем содержимое текущей строки
            i += 1 'Перемещаем указатель текущей позиции и новый виток цикла
        Loop

        Return block 'Возвращаем результат
    End Function

    ''' <summary>
    ''' Функция определяет что за выражение ей передано, а также некоторые другие параметры
    ''' </summary>
    ''' <param name="strWord">исходная строка, в которой находится нечто. Что именно - как раз надо узнать. Характеристика strWord: 
    '''- в начале строки точно нет скобок ( 
    '''- содержит выражение до пробела (или до конца строки)
    '''- если есть квадратные скобки [], то содержимое внутри них передается полностью, несмотря на пробелы
    '''- если есть круглые скобки () функции или свойства, то передается только открывающая скобка и текст за ней до первого пробела (не включительно)</param>
    ''' <param name="arrParams">массив параметров, с которыми выполняется код</param>
    ''' <param name="currentClassId">переменная для получения id класса в структуре mainClass</param>
    ''' <param name="ElementName">для получения имени функции или свойста в структуре mainClass.Functions / mainClass.Properties</param>
    ''' <param name="elementContent">получение для функции или свойства - содержимого крадратных скобок []; для переменной, Param[] и ParamCount - готовый результат (значение переменной, параметра, количество параметров)</param>
    ''' <param name="funcOpenBracketPos">положение в strWord открывающей скобки для (функции или свойства)</param>
    ''' <param name="closeBracketsCount">количество закрывающих круглых скобок ) в конце строки. Не работает для функций и свойств, которые заканчиваются собственными "функциональными" скобками</param>
    ''' <param name="returnFormat">указаны ли модификаторы преобразования типа (0 - нет, 1 - в строку $, 2 - в число #)</param>
    ''' <returns>тип выражения</returns>
    ''' <remarks>Таким образом, если в strWord:
    '''ParamCount -  elementContent = количество параметров, currentClassId, ElementId и funcOpenBracketPos не используются;
    '''Param[x] -    elementContent = значению этого параметра, currentClassId, ElementId и funcOpenBracketPos не используются;
    '''Переменная -  elementContent = значению этой переменной, currentClassId, ElementId и funcOpenBracketPos не используются;
    '''Функция или свойство - elementContent = содержимому квадратных скобок, currentClassId - идентификатор класса в mainClass, 
    '''ElementId - идентификатор функции/свойства в mainClass(0).Functions/Properties и funcOpenBracketPos позиция открывающей скобки по отношению к началу строки strWord
    '''Во всех иных случаях elementContent, currentClassId, ElementId и funcOpenBracketPos не используются.</remarks>
    Public Function GetWordType(ByVal strWord As String, ByRef arrParams() As String, ByRef currentClassId As Integer, ByRef ElementName As String,
                                ByRef elementContent As String, ByRef funcOpenBracketPos As Integer, ByRef closeBracketsCount As Integer, ByRef returnFormat As MatewScript.ReturnFormatEnum) As MatewScript.WordTypeEnum
        Dim wordLen As Integer = strWord.Length - 1 'Длина переданной строки (чтобы не проверять каждый раз)
        If wordLen = -1 Then Return MatewScript.WordTypeEnum.W_NOTHING
        elementContent = ""

        'Убираем все скобки в конце строки (если они есть)
        closeBracketsCount = 0
        Do While strWord.Chars(wordLen) = ")"c
            strWord = strWord.Substring(0, wordLen)
            wordLen -= 1
            closeBracketsCount += 1 'Сохраняем количество закрывающих скобок
            If wordLen = -1 Then Return MatewScript.WordTypeEnum.W_NOTHING
        Loop

        'Это простая строка?
        'Закомментированная ниже строка: эта проверка происходит в GetValue, поэтому повторять ее  здесь нет смысла
        'If strWord.Chars(0) = "'"c Then Return MatewScript.WordTypeEnum.W_SIMPLE_STRING

        'Проверка на True/False и все виды операторов
        If wordLen < 6 Then
            Select Case strWord
                Case "True", "False"
                    Return MatewScript.WordTypeEnum.W_SIMPLE_BOOL
                Case "+", "-", "*", "/", "\", "^"
                    Return MatewScript.WordTypeEnum.W_OPERATOR_MATH
                Case "&"
                    Return MatewScript.WordTypeEnum.W_OPERATOR_STRINGS_MERGER
                Case "=", ">", "<", ">=", "<=", "<>", "!="
                    Return MatewScript.WordTypeEnum.W_OPERATOR_COMPARE
                Case "And", "Or", "Xor"
                    Return MatewScript.WordTypeEnum.W_OPERATOR_LOGIC
            End Select
        End If

        'Это простое число?
        If IsNumeric(strWord.Replace(".", ",")) Then Return MatewScript.WordTypeEnum.W_SIMPLE_NUMBER

        'Все указанные ниже значения могут иметь в начале модификатор преобразования типа $ или #
        '$Class[x, y].#Function(z | #Var.$varName[...]
        returnFormat = MatewScript.ReturnFormatEnum.ORIGINAL
        Dim firstChar As Char
        firstChar = strWord.Chars(0)
        If firstChar = "$"c OrElse firstChar = "#"c Then
            'Модификатор найден
            'Убираем его из рабочей строки strWord
            strWord = strWord.Substring(1)
            wordLen -= 1
            'Сохраняем его значение 
            returnFormat = IIf(firstChar = "$", MatewScript.ReturnFormatEnum.TO_STRING, MatewScript.ReturnFormatEnum.TO_NUMBER)
        End If

        'Это ParamCount?
        If strWord = "ParamCount" Then
            elementContent = Convert.ToString(arrParams.Length)
            Return MatewScript.WordTypeEnum.W_PARAM_COUNT
        End If

        'Далее все значения могут иметь квадратные скобки (внутри которых может быть все, что угодно) . Для упрощения работы "прячем" их содержимое, заменив на *
        Dim hideQB As String = HideQBcontent(strWord)
        Dim qbStart As Integer 'Для хранения положения открывающей [ в строке
        'Это Param[x]?
        If wordLen > 5 AndAlso strWord.Substring(0, 6) = "Param[" Then
            'Да, это Param[x]
            Dim paramCount As Integer = arrParams.Length
            If paramCount = 0 Then 'Если параметры не переданы - то в любом случае произойдет обращение к несуществующему параметру
                LAST_ERROR = "Обращение к несуществующему параметру."
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            qbStart = strWord.IndexOf("[") 'Получаем положение открывающей [
            elementContent = strWord.Substring(qbStart + 1, strWord.Length - qbStart - 2) 'Получаем содержимое квадратных скобок
            'Получаем результат содержимого квадратных скобок (если не указан, то выбирем элемент с индексом 0)
            If elementContent.Length = 0 Then elementContent = "0"
            If Integer.TryParse(elementContent, Nothing) = False Then elementContent = GetValue(elementContent, arrParams)
            If elementContent = "#Error" Then Return MatewScript.WordTypeEnum.W_NOTHING
            If IsNumeric(elementContent) = False Then 'Если результат - не число - значит ошибка
                LAST_ERROR = "Обращение к несуществующему параметру. Идентификатор параметра - не число."
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            'Если результат > количества параметров - ошибка
            If paramCount <= Convert.ToInt32(elementContent) Then
                LAST_ERROR = "Обращение к несуществующему параметру."
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            'Сохраняем результат в elementContent 
            elementContent = arrParams(Convert.ToInt32(elementContent))
            Return MatewScript.WordTypeEnum.W_PARAM
        End If

        'Далее могут быть только 1) Функция, 2) Свойство, 3) Переменная
        'Class[x, y].#Function(z | Var.$varName[...]
        Dim currentClassName As String = "" 'Для хранения имени класса (напр. Code)
        funcOpenBracketPos = hideQB.IndexOf("(") 'Сохраняем положение открывающей скобки функции (если таковая имеется)
        If funcOpenBracketPos > -1 AndAlso returnFormat <> MatewScript.ReturnFormatEnum.ORIGINAL Then funcOpenBracketPos += 1 'Корректировка, если строка strWord была укорочена (убран модификатор #/$)
        Dim pointPos As Integer = hideQB.IndexOfAny({"."c, "'"c, "("c}) 'Находим точку
        If pointPos > -1 AndAlso hideQB.Chars(pointPos) <> "."c Then pointPos = -1

        Dim hasQbrackets As Boolean = True
        Dim qbPos As Integer = strWord.IndexOfAny({"["c, "("c, "."c})
        If qbPos > -1 AndAlso strWord.Chars(qbPos) <> "["c Then qbPos = -1
        If qbPos = -1 Then ' hideQB.Equals(strWord) Then
            'Квадратных скобок нет
            'Class.#Function(z | Var.$varName
            If pointPos > -1 Then
                'Точка найдена
                currentClassName = strWord.Substring(0, pointPos) 'Получаем имя класса
                strWord = strWord.Substring(pointPos + 1) 'Убираем его из рабочей строки (вместе с точкой)
            End If
        Else
            'Квадратные скобки найдены
            qbStart = strWord.IndexOf("[") 'Получаем положение открывающей [
            If pointPos > -1 Then
                'Точка найдена
                If qbStart < pointPos Then
                    'Открывающая квадратная скобка раньше точки. Значит, это или функция, или свойство
                    'Class[x, y].#Function(z
                    currentClassName = strWord.Substring(0, qbStart) 'Получаем имя класса
                    elementContent = strWord.Substring(qbStart + 1, pointPos - qbStart - 2) 'Сохраняем содержимое квадратных скобок
                    strWord = strWord.Substring(pointPos + 1) 'Убираем из рабочей строки имя класса и квадратные скобки
                Else
                    'Открывающая квадратная скобка после точки. Значит, это переменная
                    'Var.$varName[...]
                    currentClassName = strWord.Substring(0, pointPos) 'Получаем имя класса 
                    'Имя класса должно быть только V или Var
                    If currentClassName <> "V" AndAlso currentClassName <> "Var" Then
                        LAST_ERROR = "Ошибка синтаксиса. Возможно, вместо квадратных скобок, надо использовать круглые."
                        Return MatewScript.WordTypeEnum.W_NOTHING
                    End If
                    elementContent = strWord.Substring(qbStart + 1, strWord.Length - qbStart - 2) 'Получаем содержимое квадратных скобок
                    strWord = strWord.Substring(pointPos + 1, qbStart - pointPos - 1) 'Убираем из рабочей строки имя класса и квадратные скобки
                    'Убираем модификатор преобразование (если есть) и сохраняем значение
                    '$varName
                    firstChar = strWord.Chars(0)
                    If firstChar = "#" Then
                        returnFormat = MatewScript.ReturnFormatEnum.TO_NUMBER
                        strWord = strWord.Substring(1)
                    ElseIf firstChar = "$" Then
                        returnFormat = MatewScript.ReturnFormatEnum.TO_STRING
                        strWord = strWord.Substring(1)
                    End If
                    GoTo variableChecking 'Переход к обработке переменной
                End If
            Else
                'Точка не найдена (а квадратные скобки найдены). Это может быть только переменная
                'varName[...]
                elementContent = strWord.Substring(qbStart + 1, strWord.Length - qbStart - 2) 'Получаем содержимое квадратных скобок
                strWord = strWord.Remove(qbStart) 'Убираем их из рабочей строки
                GoTo variableChecking 'Переход к обработке переменной
            End If
        End If

        'Теперь в рабочей строке нет имени класса и квадратных скобок
        'На этом месте либо функция или свойство, либо переменная, изначально не имевшая квадратных скобок
        '#Function(z | VarName
        'Убираем модификатор преобразование и сохраняем его значения
        firstChar = strWord.Chars(0)
        If firstChar = "#" Then
            returnFormat = MatewScript.ReturnFormatEnum.TO_NUMBER
            strWord = strWord.Substring(1)
        ElseIf firstChar = "$" Then
            returnFormat = MatewScript.ReturnFormatEnum.TO_STRING
            strWord = strWord.Substring(1)
        End If

        'Function(z | VarName
        Dim bracketPos As Integer = strWord.IndexOf("(") 'Получаем положение открывающей круглой скобки
        If bracketPos > -1 Then strWord = strWord.Remove(bracketPos) 'Убираем ее и все что за ней
        'На этом этапе у нас strWord содержит "очищенные" имя функции, свойства или переменной. Все, дополнительно указанное с ними, сохранено и удалено
        'Function | VarName
        If currentClassName.Length > 0 Then
            'Если имя класса указано
            If currentClassName = "V" OrElse currentClassName = "Var" Then
                'Это точно переменная
                If bracketPos > -1 Then 'Если найдена круглая скобка, то ошибка: их после переменной быть не должно
                    LAST_ERROR = "Ошибка синтаксиса. Возможно, надо использовать квадратные скобки вместо круглых."
                    Return MatewScript.WordTypeEnum.W_NOTHING
                End If
                GoTo variableChecking 'Переход к обработке переменной
            End If

            If mainClassHash.ContainsKey(currentClassName) = False Then 'Имя указанного класса неправильное (не определено)
                LAST_ERROR = "Указано несуществующее имя класса."
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If

            'Function
            ElementName = strWord
            'Сохраняем идентификатор класса в mainClass
            currentClassId = mainClassHash(currentClassName)
            'Ищем функцию в указанном классе
            If mainClass(currentClassId).Functions.ContainsKey(strWord) Then
                'Функция найдена
                Return MatewScript.WordTypeEnum.W_FUNCTION
            End If
            'Ищем свойство в указанном классе
            If mainClass(currentClassId).Properties.ContainsKey(strWord) Then
                'Свойство найдено
                Return MatewScript.WordTypeEnum.W_PROPERTY
            End If

            'В указанном классе нет указанного свойства или функции - ошибка
            LAST_ERROR = "В классе " + currentClassName + " функция/свойство " + strWord + " не найдено."
            Return MatewScript.WordTypeEnum.W_NOTHING
        Else
            ElementName = strWord
            'Имя класса не указано
            'Ищем свойство или функцию с нужным именем по всем классам
            For i As Integer = 0 To mainClass.Length - 1

                If mainClass(i).Functions.ContainsKey(strWord) Then
                    'Функция найдена
                    currentClassId = i 'Сохраняем идентификатор класса
                    Return MatewScript.WordTypeEnum.W_FUNCTION
                End If

                If mainClass(currentClassId).Properties.ContainsKey(strWord) Then
                    'Свойство найдено
                    currentClassId = i 'Сохраняем идентификатор класса
                    Return MatewScript.WordTypeEnum.W_PROPERTY
                End If
            Next
        End If

variableChecking:
        'На этом этапе может быть только переменная, при этом в strWord храниться только ее имя
        'VarName
        'На всякий случай проверяем, нет ли в конце открывающих скобок. Если есть  - неправильный синтаксис
        If bracketPos = 0 Then bracketPos = strWord.IndexOf("(")
        If bracketPos > -1 Then
            LAST_ERROR = "Неправильный синтаксис. Нераспознанная функция/свойство или круглые скобки вместо квадратных после переменной"
            Return MatewScript.WordTypeEnum.W_NOTHING
        End If
        bracketPos = strWord.IndexOf("["c)
        If bracketPos > -1 Then
            'если есть квадратные скобки
            If strWord.Last <> "]"c Then
                LAST_ERROR = "Не найдена закрывающая квадратная скобка переменной."
                Return WordTypeEnum.W_NOTHING
            End If
            elementContent = strWord.Substring(bracketPos + 1, strWord.Length - bracketPos - 2)
            strWord = strWord.Substring(0, bracketPos)
        End If
        'Получаем результат содержимого квадратных скобок (внутри которых должен быть индекс переменной-массива или сигнатура)

        If elementContent.Length = 0 Then elementContent = "0"
        Dim retType As ReturnFormatEnum = ReturnFormatEnum.TO_NUMBER
        If Integer.TryParse(elementContent, Nothing) = False Then
            elementContent = GetValue(elementContent, arrParams)
            retType = Param_GetType(elementContent)
            If elementContent = "#Error" Then Return MatewScript.WordTypeEnum.W_NOTHING 'В квадратных скобках ошибка
            If retType <> ReturnFormatEnum.TO_NUMBER AndAlso retType <> ReturnFormatEnum.TO_STRING Then 'Результат - не число. Ошибка!
                LAST_ERROR = "Индекс переменной - не число/строка."
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
        End If

        'Получаем значение переменной
        Dim varResult As String
        If retType = ReturnFormatEnum.TO_STRING Then
            varResult = csLocalVariables.GetVariable(strWord, elementContent) 'Поиск в локальных переменных
            If varResult <> "#Error" Then
                elementContent = varResult
                Return MatewScript.WordTypeEnum.W_VARIABLE_LOCAL
            Else 'Локальной переменной не существует
                varResult = csPublicVariables.GetVariable(strWord, elementContent) 'Поиск в глобальных переменных
                If varResult = "#Error" Then 'Переменной не существует
                    LAST_ERROR = "Переменной " + strWord + "[" & elementContent & "] не существует."
                    Return MatewScript.WordTypeEnum.W_NOTHING
                End If
                elementContent = varResult
                Return MatewScript.WordTypeEnum.W_VARIABLE_PUBLIC
            End If
        Else
            varResult = csLocalVariables.GetVariable(strWord, Convert.ToInt32(elementContent)) 'Поиск в локальных переменных
            If varResult <> "#Error" Then
                elementContent = varResult
                Return MatewScript.WordTypeEnum.W_VARIABLE_LOCAL
            Else 'Локальной переменной не существует
                varResult = csPublicVariables.GetVariable(strWord, Convert.ToInt32(elementContent)) 'Поиск в глобальных переменных
                If varResult = "#Error" Then 'Переменной не существует
                    LAST_ERROR = "Переменной " + strWord + " не существует."
                    Return MatewScript.WordTypeEnum.W_NOTHING
                End If
                elementContent = varResult
                Return MatewScript.WordTypeEnum.W_VARIABLE_PUBLIC
            End If
        End If

er:
        MessageBox.Show("Критическая ошибка при выполнении GetWordType!", "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return MatewScript.WordTypeEnum.W_NOTHING
    End Function

    ''' <summary>
    ''' Функция определяет значение первого слова в строке кода, а также определяет некоторые другие параметры
    ''' </summary>
    ''' <param name="strCodeString">строка кода, с которой работаем</param>
    ''' <param name="qbContent">переменная для получения содержимого квадратных скобок (в переменных, функциях, свойствах и Param[x])</param>
    ''' <param name="currentClassId">для получения класса (для функций и свойств)</param>
    ''' <param name="elementName">для получения имени переменной, функции, свойства</param>
    ''' <param name="stringRemain">остаток строки после первого слова (без пробела, а также без первой скобки в функциях)</param>
    ''' <param name="returnFormat">для получения формата, в котором выдавать результат (чистло, строка, Да/Нет или не трогать)</param>
    ''' <param name="stepId">только для цикла For, для получения позиции Step</param>
    ''' <returns>тип первого слова</returns>
    ''' <remarks></remarks>
    Public Function GetFirstWordType(ByVal strCodeString As String, ByRef qbContent As String, ByRef currentClassId As Integer, _
                    ByRef elementName As String, ByRef stringRemain As String, ByRef returnFormat As ReturnFormatEnum, _
                    ByRef stepId As Integer) As MatewScript.WordTypeEnum
        If strCodeString.Length = 0 Then Return MatewScript.WordTypeEnum.W_NOTHING
        'сбрасываем на пустые значения
        qbContent = ""
        stringRemain = ""
        returnFormat = MatewScript.ReturnFormatEnum.ORIGINAL

        'Проверка является ли строка одним из операторов
        If strCodeString = "Exit" Then Return MatewScript.WordTypeEnum.W_EXIT
        If strCodeString = "Return" Then Return MatewScript.WordTypeEnum.W_RETURN
        'If strCodeString = "Return" OrElse strCodeString.StartsWith("?") Then Return MatewScript.WordTypeEnum.W_RETURN

        If strCodeString.EndsWith(":") Then Return MatewScript.WordTypeEnum.W_MARK
        If strCodeString = "Continue" Then Return MatewScript.WordTypeEnum.W_CONTINUE
        If strCodeString = "Break" Then Return MatewScript.WordTypeEnum.W_BREAK
        If strCodeString = "Do" Then Return WordTypeEnum.W_BLOCK_DOWHILE
        If strCodeString.StartsWith("HTML" + Chr(1)) Then
            stringRemain = strCodeString.Substring(6)
            Return MatewScript.WordTypeEnum.W_HTML
        End If

        Dim brPos As Integer = -1, operatorPos As Integer = -1, i As Integer = 0
        If strCodeString.StartsWith("If ") Then
            'Для блока If ... Then в qbContent помещаем условие блока, а в stringRemain - все после Then
            Do
                If brPos < i Then 'Поиск '
                    brPos = strCodeString.IndexOf("'", i)
                    If brPos = -1 Then brPos = Integer.MaxValue
                End If
                If operatorPos < i Then 'Поиск оператора Then
                    operatorPos = strCodeString.IndexOf(" Then", i)
                    If operatorPos = -1 Then
                        LAST_ERROR = "Не найдено оператора Then в блоке If ... Then."
                        Return MatewScript.WordTypeEnum.W_NOTHING
                    End If
                End If
                i = Math.Min(brPos, operatorPos)
                If strCodeString.Substring(i, 5) = " Then" Then
                    'Then найдено
                    If i < 3 Then
                        LAST_ERROR = "Отсутствует условие в блоке If ... Then."
                        Return MatewScript.WordTypeEnum.W_NOTHING
                    End If
                    qbContent = strCodeString.Substring(3, i - 3) 'сохраняем условие блока
                    'Если после Then  есть код - сохраняем и его
                    If strCodeString.Length > i + 5 Then stringRemain = strCodeString.Substring(i + 6)
                    Return MatewScript.WordTypeEnum.W_BLOCK_IF
                Else
                    'Найдена строка '...'.
qbQuoteSearching:
                    i = strCodeString.IndexOf("'", i + 1) 'ищем закрывающий "'"
                    If i = -1 Then
                        'закрывающий "'" не найден - ошибка
                        LAST_ERROR = "Не найдена закрывающая кавычка"
                        Return MatewScript.WordTypeEnum.W_NOTHING
                    End If
                    'Обрабатываем экранированную кавычку /'
                    If strCodeString.Chars(i - 1) = "/"c Then GoTo qbQuoteSearching
                End If
                i += 1
            Loop
        End If
        If strCodeString.StartsWith("For ") Then
            'For i = 1 To 10 Step 2
            'i - elementName 
            '1 - qbContent 
            '10 - stringRemain 
            '2 - позиция в строке strCodeString сохраняется в stepId. Если Step нет, то = -1
            i = HideQBcontent(strCodeString).IndexOf(" ", 5)
            If i = -1 Then
                LAST_ERROR = "Неверная запись цикла For ... Next"
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            elementName = strCodeString.Substring(4, i - 4) 'получаем имя переменной
            If strCodeString.Substring(i + 1, 1) <> "=" Then
                LAST_ERROR = "Неверная запись цикла For ... Next"
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            i = i + 3 '...1 To 10 Step 2
            If i > strCodeString.Length Then
                LAST_ERROR = "Неверная запись цикла For ... Next"
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            stepId = i 'сохраняем начало первого значения переменной цикла
            'Поиск То
            Do
                If brPos < i Then 'Поиск '
                    brPos = strCodeString.IndexOf("'", i)
                    If brPos = -1 Then brPos = Integer.MaxValue
                End If
                If operatorPos < i Then 'Поиск оператора To
                    operatorPos = strCodeString.IndexOf(" To ", i)
                    If operatorPos = -1 Then
                        LAST_ERROR = "Не найдено оператора To в цикле For ... Next."
                        Return MatewScript.WordTypeEnum.W_NOTHING
                    End If
                End If
                i = Math.Min(brPos, operatorPos)
                If strCodeString.Substring(i, 4) = " To " Then
                    'To найдено
                    qbContent = strCodeString.Substring(stepId, i - stepId) 'сохраняем первое значение переменной цикла
                    Exit Do
                Else
                    'Найдена строка '...'.
qbQuoteSearching2:
                    i = strCodeString.IndexOf("'", i + 1) 'ищем закрывающий "'"
                    If i = -1 Then
                        'закрывающий "'" не найден - ошибка
                        LAST_ERROR = "Не найдена закрывающая кавычка"
                        Return MatewScript.WordTypeEnum.W_NOTHING
                    End If
                    'Обрабатываем экранированную кавычку /'
                    If strCodeString.Chars(i - 1) = "/"c Then GoTo qbQuoteSearching2
                End If
                i += 1
            Loop

            i = i + 4 '...10 Step 2
            stepId = i 'сохраняем начало второго значения переменной цикла
            'Поиск Step
            Do
                If brPos < i Then 'Поиск '
                    brPos = strCodeString.IndexOf("'", i)
                    If brPos = -1 Then brPos = Integer.MaxValue
                End If
                If operatorPos < i Then 'Поиск оператора Step
                    operatorPos = strCodeString.IndexOf(" Step ", i)
                    If operatorPos = -1 Then
                        stringRemain = strCodeString.Substring(stepId)
                        stepId = -1
                        Return MatewScript.WordTypeEnum.W_BLOCK_FOR
                    End If
                End If
                i = Math.Min(brPos, operatorPos)
                If strCodeString.Substring(i, 6) = " Step " Then
                    'Step найден
                    stringRemain = strCodeString.Substring(stepId, i - stepId) 'сохраняем второе значение переменной цикла
                    stepId = i + 6 'сохраняем позицию условия перебора (...10 Step |2)
                    Return MatewScript.WordTypeEnum.W_BLOCK_FOR
                Else
                    'Найдена строка '...'.
qbQuoteSearching3:
                    i = strCodeString.IndexOf("'", i + 1) 'ищем закрывающий "'"
                    If i = -1 Then
                        'закрывающий "'" не найден - ошибка
                        LAST_ERROR = "Не найдена закрывающая кавычка"
                        Return MatewScript.WordTypeEnum.W_NOTHING
                    End If
                    'Обрабатываем экранированную кавычку /'
                    If strCodeString.Chars(i - 1) = "/"c Then GoTo qbQuoteSearching3
                End If
                i += 1
            Loop

        End If
        If strCodeString.StartsWith("HTML ") Then
            stringRemain = strCodeString.Substring(5)
            Return MatewScript.WordTypeEnum.W_HTML
        End If
        If strCodeString.StartsWith("Wrap ") Then
            i = strCodeString.IndexOf(Chr(1), 5)
            If i = -1 Then
                LAST_ERROR = "ошибка при распознавании блока Wrap."
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            elementName = strCodeString.Substring(5, i - 5)
            stringRemain = strCodeString.Substring(i + 1).Trim(" "c, vbCr, vbLf)
            Return MatewScript.WordTypeEnum.W_WRAP
        End If
        If strCodeString.StartsWith("Jump ") Then
            stringRemain = strCodeString.Substring(5)
            If stringRemain.Length > 1 Then stringRemain = stringRemain.Substring(1, stringRemain.Length - 2).Replace("/'", "'")
            Return MatewScript.WordTypeEnum.W_JUMP
        End If
        If strCodeString.StartsWith("Break ") Then
            stringRemain = strCodeString.Substring(6)
            Return MatewScript.WordTypeEnum.W_BREAK
        End If
        If strCodeString.StartsWith("Event ") Then
            stringRemain = strCodeString.Substring(6)
            Return MatewScript.WordTypeEnum.W_BLOCK_EVENT
        End If
        If strCodeString.StartsWith("Switch ") Then
            stringRemain = strCodeString.Substring(7)
            Return MatewScript.WordTypeEnum.W_SWITCH
        End If
        If strCodeString.StartsWith("Return ") Then
            stringRemain = strCodeString.Substring(7)
            Return MatewScript.WordTypeEnum.W_RETURN
        End If
        If strCodeString.StartsWith("?") Then
            stringRemain = strCodeString.Substring(1).TrimStart
            Return MatewScript.WordTypeEnum.W_RETURN
        End If
        Dim blnOperatorGlobal As Boolean = False
        If strCodeString.StartsWith("Global ") Then
            strCodeString = strCodeString.Substring(7)
            blnOperatorGlobal = True
            GoTo fromGlobal
        End If
        If strCodeString.StartsWith("Do While ") Then
            stringRemain = strCodeString.Substring(9)
            Return MatewScript.WordTypeEnum.W_BLOCK_DOWHILE
        End If
        If strCodeString.StartsWith("Function ") Then
            stringRemain = strCodeString.Substring(9)
            Return MatewScript.WordTypeEnum.W_BLOCK_FUNCTION
        End If
        If strCodeString.StartsWith("New Class ") Then
            stringRemain = strCodeString.Substring(10)
            Return MatewScript.WordTypeEnum.W_BLOCK_NEWCLASS
        End If
        If strCodeString.StartsWith("Rem Class ") Then
            stringRemain = strCodeString.Substring(10)
            Return MatewScript.WordTypeEnum.W_REM_CLASS
        End If
        If strCodeString.StartsWith("Select Case ") Then
            stringRemain = strCodeString.Substring(12)
            Return MatewScript.WordTypeEnum.W_SWITCH
        End If

        If strCodeString.StartsWith("ElseIf") OrElse strCodeString = "Else" OrElse strCodeString = "End" OrElse _
            strCodeString.StartsWith("End ") OrElse strCodeString = "Loop" OrElse strCodeString = "Next" OrElse _
            strCodeString.StartsWith("Next ") OrElse strCodeString.StartsWith("Case ") Then Return MatewScript.WordTypeEnum.W_CYCLE_INTERNAL

fromGlobal:
        'Проверяем не указан ли формат (с помощью # или $), в котором выдавать результат (число или строка)
        If strCodeString.Chars(0) = "#"c Then
            returnFormat = MatewScript.ReturnFormatEnum.TO_NUMBER
            strCodeString = strCodeString.Substring(1)
        ElseIf strCodeString.Chars(0) = "$"c Then
            returnFormat = MatewScript.ReturnFormatEnum.TO_STRING
            strCodeString = strCodeString.Substring(1)
        End If

        'Param[x]?
        If strCodeString.StartsWith("Param[") Then
            i = HideQBcontent(strCodeString).IndexOf(" ")
            If i = -1 Then
                LAST_ERROR = "Не правильное обращение к Param[x]"
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            stringRemain = strCodeString.Substring(i + 1)
            qbContent = strCodeString.Substring(6, i - 7)
            Return MatewScript.WordTypeEnum.W_PARAM
        ElseIf strCodeString.StartsWith("Param ") Then
            stringRemain = strCodeString.Substring(6)
            qbContent = ""
            Return MatewScript.WordTypeEnum.W_PARAM
        End If

        'Далее могут быть только функции, свойства или переменные
        'Class[x, y].#Element(z) = | Var.#myVar[x] =
        'переменные для хранения позиции в строке . [ ( и пробела
        Dim pointPos, qbPos, cbPos, spacePos As Integer
        Dim strQbHidden As String = HideQBcontent(strCodeString) 'строка со спрятанным содержимым квадратных скобок
        Dim currentClassName As String 'для получения имени класса

        pointPos = strQbHidden.IndexOf(".") 'положение точки
        If pointPos = -1 Then pointPos = Integer.MaxValue

        qbPos = strCodeString.IndexOf("[") 'положение [
        If qbPos = -1 Then qbPos = Integer.MaxValue

        spacePos = strQbHidden.IndexOf(" ") 'положение пробела
        If spacePos = -1 Then spacePos = Integer.MaxValue

        cbPos = strQbHidden.IndexOf("(") 'положение (
        If cbPos = -1 Then cbPos = Integer.MaxValue

        If pointPos < qbPos And pointPos < spacePos And pointPos < cbPos Then
            'Если точка находится до квадратных скобок, до пробела и до круглых скобок
            'Значит, есть имя класса и нет [ до точки. 
            'Class.#Element(z) = | Var.#myVar[x] = 
            currentClassName = strCodeString.Substring(0, pointPos) 'получаем имя класса
            'Проверка на операторы преобразования типа $ и # за именем класса
            If strCodeString.Chars(pointPos + 1) = "#"c Then
                returnFormat = MatewScript.ReturnFormatEnum.TO_NUMBER
                pointPos += 1
            ElseIf strCodeString.Chars(pointPos + 1) = "$"c Then
                returnFormat = MatewScript.ReturnFormatEnum.TO_STRING
                pointPos += 1
            End If

            If currentClassName = "V" OrElse currentClassName = "Var" Then
                'Это явно указанная переменная
                'Var.myVar[x] =
                stringRemain = strCodeString.Substring(spacePos + 1) 'получаем остаток строки
                If spacePos < qbPos Then
                    'квадратных скобок нет
                    'myVar =
                    elementName = strCodeString.Substring(pointPos + 1, spacePos - pointPos - 1) 'получаем имя переменной
                    If blnOperatorGlobal Then
                        Return MatewScript.WordTypeEnum.W_VARIABLE_PUBLIC
                    Else
                        Return MatewScript.WordTypeEnum.W_VARIABLE_LOCAL
                    End If
                Else
                    'есть квадратные скобки
                    'myVar[x] =
                    elementName = strCodeString.Substring(pointPos + 1, qbPos - pointPos - 1) 'получаем имя переменной
                    qbContent = strCodeString.Substring(qbPos + 1, spacePos - qbPos - 2) 'получаем содержимое квадратных скобок
                    If blnOperatorGlobal Then
                        Return MatewScript.WordTypeEnum.W_VARIABLE_PUBLIC
                    Else
                        Return MatewScript.WordTypeEnum.W_VARIABLE_LOCAL
                    End If
                End If
            End If
            If blnOperatorGlobal Then
                LAST_ERROR = "За оператором Global должно стоять имя переменной."
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            If Math.Min(spacePos, cbPos) < Integer.MaxValue Then stringRemain = strCodeString.Substring(Math.Min(spacePos, cbPos) + 1) 'получаем остаток строки
            'Function
            'Сохраняем идентификатор класса в mainClass
            currentClassId = mainClassHash(currentClassName)
            'Получаем имя функции/свойства
            If Math.Min(spacePos, cbPos) < Integer.MaxValue Then
                elementName = strCodeString.Substring(pointPos + 1, Math.Min(spacePos, cbPos) - pointPos - 1)
            Else
                elementName = strCodeString.Substring(pointPos + 1)
            End If
            'Ищем функцию в указанном классе
            If mainClass(currentClassId).Functions.ContainsKey(elementName) Then
                'Функция найдена
                Return MatewScript.WordTypeEnum.W_FUNCTION
            End If
            'Ищем свойство в указанном классе
            If mainClass(currentClassId).Properties.ContainsKey(elementName) Then
                'Свойство найдено
                Return MatewScript.WordTypeEnum.W_PROPERTY
            End If

            'В указанном классе нет указанного свойства или функции - ошибка
            LAST_ERROR = "В классе " + currentClassName + " нет указанного свойства или функции."
            Return MatewScript.WordTypeEnum.W_NOTHING
        End If

        spacePos = Math.Min(spacePos, cbPos)
        If qbPos < pointPos And qbPos < spacePos And pointPos < spacePos Then
            If blnOperatorGlobal Then
                LAST_ERROR = "Неправильная запись переменной за оператором Global."
                Return MatewScript.WordTypeEnum.W_NOTHING
            End If
            'есть квадратные скобки до точки
            'Class[x, y].#Element(z) =
            currentClassName = strCodeString.Substring(0, qbPos) 'получаем имя класса
            'получаем содержимое квадратных скобок
            qbContent = strCodeString.Substring(qbPos + 1, pointPos - qbPos - 2)
            'Проверка на операторы преобразования типа $ и # за именем класса
            If strCodeString.Chars(pointPos + 1) = "#"c Then
                returnFormat = MatewScript.ReturnFormatEnum.TO_NUMBER
                pointPos += 1
            ElseIf strCodeString.Chars(pointPos + 1) = "$"c Then
                returnFormat = MatewScript.ReturnFormatEnum.TO_STRING
                pointPos += 1
            End If
            If spacePos < Integer.MaxValue Then stringRemain = strCodeString.Substring(spacePos + 1) 'получаем остаток строки
            'Function
            'Сохраняем идентификатор класса в mainClass
            currentClassId = mainClassHash(currentClassName)
            'Получаем имя функции/свойства
            If spacePos < Integer.MaxValue Then
                elementName = strCodeString.Substring(pointPos + 1, spacePos - pointPos - 1)
            Else
                elementName = strCodeString.Substring(pointPos + 1)
            End If
            'Ищем функцию в указанном классе
            If mainClass(currentClassId).Functions.ContainsKey(elementName) Then
                'Функция найдена
                Return MatewScript.WordTypeEnum.W_FUNCTION
            End If
            'Ищем свойство в указанном классе
            If mainClass(currentClassId).Properties.ContainsKey(elementName) Then
                'Свойство найдено
                Return MatewScript.WordTypeEnum.W_PROPERTY
            End If

            'В указанном классе нет указанного свойства или функции - ошибка
            LAST_ERROR = "В классе " + currentClassName + " нет указанного свойства или функции."
            Return MatewScript.WordTypeEnum.W_NOTHING
        End If

        'имя класса нет
        'Element(z) = | myVar[x] = 
        If spacePos < Integer.MaxValue Then stringRemain = strCodeString.Substring(spacePos + 1) 'получаем остаток строки
        If qbPos < spacePos Then
            'есть квадратные скобки
            'myVar[x] = 
            elementName = strCodeString.Substring(0, qbPos)
            qbContent = strCodeString.Substring(qbPos + 1, spacePos - qbPos - 2)
            If blnOperatorGlobal Then
                Return MatewScript.WordTypeEnum.W_VARIABLE_PUBLIC
            Else
                Return MatewScript.WordTypeEnum.W_VARIABLE_LOCAL
            End If
        End If

        'квадратных скобок нет
        'Element = | myVar = 
        'Имя класса не указано
        If spacePos < Integer.MaxValue Then
            elementName = strCodeString.Substring(0, spacePos)
        Else
            elementName = strCodeString
        End If
        If blnOperatorGlobal = False Then
            'Ищем свойство или функцию с нужным именем по всем классам
            For i = 0 To mainClass.Length - 1
                If mainClass(i).Functions.ContainsKey(elementName) Then
                    'Функция найдена
                    currentClassId = i 'Сохраняем идентификатор класса
                    Return MatewScript.WordTypeEnum.W_FUNCTION
                End If

                If mainClass(i).Properties.ContainsKey(elementName) Then
                    'Свойство найдено
                    currentClassId = i 'Сохраняем идентификатор класса
                    Return MatewScript.WordTypeEnum.W_PROPERTY
                End If
            Next
        End If

        'это переменная
        If blnOperatorGlobal Then
            Return MatewScript.WordTypeEnum.W_VARIABLE_PUBLIC
        Else
            Return MatewScript.WordTypeEnum.W_VARIABLE_LOCAL
        End If
    End Function

#End Region

#Region "New additions functions"
    ''' <summary>
    ''' Возвращает Id элемента в массиве EditWordType, следующего за скобками - как (), так и []
    ''' </summary>
    ''' <param name="masCode">массив <see cref="CodeTextBox.EditWordType"></see></param>
    ''' <param name="startId">начало поиска</param>
    ''' <returns>Id элемента в массиве, -1 если за скобками строка заканчивается и -2 при непарном количестве скобок</returns>
    ''' <remarks></remarks>
    Private Function GetNextElementIdAfterBrackets(ByRef masCode() As CodeTextBox.EditWordType, startId As Integer) As Integer
        If startId >= masCode.Count - 1 Then
            If masCode(masCode.Count - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE Then Return masCode.Count - 1
            Return -1
        End If
        Dim ovalBracketsBalance As Integer = 0
        Dim quadBracketsBalance As Integer = 0
        For i As Integer = startId To masCode.GetUpperBound(0)
            Select Case masCode(i).wordType
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                    ovalBracketsBalance += 1
                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                    ovalBracketsBalance -= 1
                    If ovalBracketsBalance < 0 AndAlso quadBracketsBalance = 0 Then Return i
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                    quadBracketsBalance += 1
                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                    quadBracketsBalance -= 1
                    If quadBracketsBalance < 0 AndAlso ovalBracketsBalance = 0 Then Return i
                Case Else
                    If ovalBracketsBalance = 0 AndAlso quadBracketsBalance = 0 Then Return i
            End Select
        Next
        If ovalBracketsBalance = 0 AndAlso quadBracketsBalance = 0 Then
            Return -1 'конец строки
        Else
            LAST_ERROR = "Непарное количество скобок."
            Return -2 'непарное кол-во скобок - ошибка
        End If

    End Function

    ''' <summary>
    ''' Функция проверяет код с пользовательским элементом, уточняя содержание в нем свойства или функции, а также определяет
    ''' индекс класса. Необходимо для корректной работы с пользовательскими классами и элементами, объявленными прямо в коде,
    ''' когда заранее не известно что это - функция или свойство.
    ''' </summary>
    ''' <param name="code">код для проверки</param>
    ''' <returns>пустую строку, при ошибке #Error</returns>
    ''' <remarks></remarks>
    Private Function ClarifyUserElement(ByRef code As CodeTextBox.EditWordType) As String
        Dim elName As String = code.Word.Trim

        If code.classId >= 0 Then
            'класс встроенный. Функция/свойство могут быть пользовательскими
            Dim cId As Integer = code.classId
            If cId > mainClass.GetUpperBound(0) Then Return "#Error"
            'если вместо функции указано, что это свойство - меняем (и наоборот)
            If code.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION Then
                If mainClass(cId).Functions.ContainsKey(elName) Then Return ""
                If mainClass(cId).Properties.ContainsKey(elName) Then
                    code.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY
                    Return ""
                Else
                    Return "#Error"
                End If
            ElseIf code.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY Then
                If mainClass(cId).Properties.ContainsKey(elName) Then Return ""
                If mainClass(cId).Functions.ContainsKey(elName) Then
                    code.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION
                    Return ""
                Else
                    Return "#Error"
                End If
            Else
                Return "#Error"
            End If
        End If

        'указан пользовательский класс
        For i As Integer = mainClass.GetUpperBound(0) To 0 Step -1
            If mainClass(i).UserAdded = False Then Continue For
            'ищем в пользовательских классах указанный элемент
            If mainClass(i).Functions.ContainsKey(elName) Then
                code.classId = i
                code.wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION
                Return ""
            ElseIf mainClass(i).Properties.ContainsKey(elName) Then
                code.classId = i
                code.wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY
                Return ""
            End If
        Next
        Return "#Error"
    End Function
#End Region

#Region "Additional functions"

    ''' <summary>Главная функция, выполняющая весь код в strCode, запущенный с параметрами arrParams.
    ''' Код может содержать комментарии, символы переноса кода a = _1 и разделения строк a = 1;b = 2.
    ''' Также код должен содержать правильное кол-во пробелов между всеми операторами, а сами операторы должны быть написаны с соблюдением регистра
    ''' </summary>
    ''' <param name="strCode">массив строк с кодом</param>
    ''' <param name="arrParams">параметры, передаваемые коду</param>
    ''' <returns>результат выполнения кода или #Error в случае ошибки.</returns>
    ''' <remarks></remarks>
    Public Function ExecuteString(ByVal strCode() As String, ByRef arrParams() As String) As String
        If IsNothing(strCode) OrElse strCode.GetUpperBound(0) = -1 Then Return ""
        'Возвращаем результат выполнения кода. Перевод кода в нужный формат DictionaryEntry выполняет PrepareBlock,
        'Само выполнение кода происходит в функции Execute
        Return ExecuteCode(PrepareBlock(strCode), arrParams, True)
    End Function

    ''' <summary>
    ''' Определяет в каком формате хранится переданный параметр
    ''' </summary>
    ''' <param name="strParam">Строка с параметром, формат которого надо определить</param>
    ''' <returns></returns>
    Public Function Param_GetType(ByVal strParam As String) As ReturnFormatEnum
        If String.IsNullOrEmpty(strParam) Then Return ReturnFormatEnum.ORIGINAL
        If strParam.Chars(0) = "'"c Then
            If strParam.Length > 1 AndAlso strParam.Chars(1) = "?"c Then Return ReturnFormatEnum.TO_EXECUTABLE_STRING
            Return ReturnFormatEnum.TO_STRING
        End If
        If strParam = "True" OrElse strParam = "False" Then Return ReturnFormatEnum.TO_BOOL
        If strParam = "0" OrElse Val(strParam) <> 0 Then Return ReturnFormatEnum.TO_NUMBER
        Dim ct As ContainsCodeEnum = IsPropertyContainsCode(strParam)
        If ct <> ContainsCodeEnum.NOT_CODE Then
            Select Case ct
                Case ContainsCodeEnum.CODE
                    Return ReturnFormatEnum.TO_CODE
                Case ContainsCodeEnum.LONG_TEXT
                    Return ReturnFormatEnum.TO_LONG_TEXT
                Case ContainsCodeEnum.EXECUTABLE_STRING
                    Return ReturnFormatEnum.TO_EXECUTABLE_STRING
            End Select
        End If
        If strParam = "#ARRAY" Then Return ReturnFormatEnum.TO_ARRAY
        Return ReturnFormatEnum.ORIGINAL
    End Function

    ''' <summary>
    ''' Процедура преобразует содержимое strElement в формат, указанный в formatTo
    ''' </summary>
    ''' <param name="strElement">Строка для преобразования</param>
    ''' <param name="formatTo">в какой формат перевести. Если равен ORIGINAL (т. е. исходному - преобразовывать не надо), то определяется, что это за исходный формат.</param>
    ''' <param name="formatFrom">из какого формата. Если не указан, то он определяется автоматически.</param>
    ''' <returns></returns>
    ''' <remarks>Если formatFrom (формат исходной строки strElement) не указан, то он определяется автоматически.
    ''' Если formatTo (формат, в который преобразовывать) равен ORIGINAL (т. е. исходному - преобразовывать не надо), то определяется, что это за исходный формат.</remarks>
    Public Function ConvertElement(ByRef strElement As String, ByRef formatTo As MatewScript.ReturnFormatEnum, Optional ByVal formatFrom As MatewScript.ReturnFormatEnum = MatewScript.ReturnFormatEnum.ORIGINAL) As String
        On Error GoTo er

        If String.IsNullOrEmpty(strElement) Then
            'Строка пустая - Empty
            strElement = ""
            If formatTo = ReturnFormatEnum.TO_STRING Then
                strElement = "''"
            ElseIf formatTo = ReturnFormatEnum.TO_NUMBER Then
                strElement = "0"
            End If
            Return strElement
        End If
        If strElement.Length = 0 Then Return strElement 'Строка пустая - ее же и возвращаем

        If formatFrom = ReturnFormatEnum.ORIGINAL OrElse formatTo = ReturnFormatEnum.ORIGINAL Then
            'Определяем формат переданной строки
            If strElement = "True" OrElse strElement = "False" Then
                formatFrom = ReturnFormatEnum.TO_BOOL 'True/False
            ElseIf strElement.Chars(0) = "'"c Then
                formatFrom = ReturnFormatEnum.TO_STRING 'Строка
            ElseIf IsPropertyContainsCode(strElement) <> ContainsCodeEnum.NOT_CODE Then
                formatFrom = ReturnFormatEnum.TO_CODE
            Else
                formatFrom = ReturnFormatEnum.TO_NUMBER 'Число
            End If
            If formatTo = ReturnFormatEnum.ORIGINAL Then 'преобразовывать не надо
                formatTo = formatFrom 'сохраняем в formatTo формат исходной строки
                Return strElement
            End If
        End If

        If formatFrom = formatTo Then Return strElement 'Формат из = формату в - выход

        Select Case formatFrom
            Case ReturnFormatEnum.TO_NUMBER
                'Исходный формат - число
                If formatTo = ReturnFormatEnum.TO_STRING Then
                    'Число в строку
                    strElement = "'" + strElement + "'"
                Else
                    'Число в True/False
                    strElement = IIf(Convert.ToBoolean(Val(strElement)), "True", "False")
                End If
            Case ReturnFormatEnum.TO_STRING
                'Исходный формат - строка
                If formatTo = ReturnFormatEnum.TO_NUMBER Then
                    'Строка в число
                    strElement = strElement.Substring(1, strElement.Length - 2).Replace("/'", "'")
                    strElement = Val(strElement)
                    strElement = strElement.Replace(",", ".")
                Else
                    'Строка в True/False
                    strElement = strElement.ToLower
                    If strElement = "'true'" Then
                        strElement = "True"
                    Else
                        strElement = "False"
                    End If
                End If
            Case ReturnFormatEnum.TO_BOOL
                'Исходный формат - True/False
                If formatTo = ReturnFormatEnum.TO_STRING Then
                    'True/False в строку
                    strElement = "'" + strElement + "'"
                Else
                    'True/False в число
                    strElement = IIf(strElement = "True", 1, 0)
                End If
        End Select

        Return strElement
er:
        MessageBox.Show("Ошибка в процедуре ConvertElement!", "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return "#Error"
    End Function

    ''' <summary>Функция заменяет содержимое квадратных скобок [...] на звездочки ***.
    '''Например: A[BBB].Caption = H[CCC].Name =>A*****.Caption = H*****.Name
    ''' </summary>
    ''' <param name="strExpression">Строка, в которой будет происходить замена</param>
    Public Function HideQBcontent(ByVal strExpression As String) As String

        Dim i As Integer = 0, j As Integer = 0
        Dim qbCnt As Integer 'количество вложенных скобок: [..[..[.. = 3
        Dim qStart As Integer 'положение первой скобки в этой группе
        Dim chOpn As Integer = -1, chCls As Integer = -1, chBr As Integer = -1 'положение последних найденных [, ] и '


        'Если нет скобок - не продолжать
        chOpn = strExpression.IndexOf("[", i)
        If chOpn = -1 Then Return strExpression

        strExpression = strExpression.Replace("/'", "**") 'заменяем заблокированные кавычки, чтоб не путались

        'Главный цикл (GoTo немного быстрее Do While)
st:
        'ищем [, ] и '. каждый символ находим только по одному разу
        If chOpn < i Then
            chOpn = strExpression.IndexOf("[", i)
            If chOpn = -1 Then chOpn = Integer.MaxValue
        End If
        If chCls < i Then
            chCls = strExpression.IndexOf("]", i)
            If chCls = -1 Then chCls = Integer.MaxValue
        End If
        If chBr < i Then
            chBr = strExpression.IndexOf("'", i)
            If chBr = -1 Then chBr = Integer.MaxValue
        End If

        'Получаем первый из найденных символов
        j = IIf(chOpn < chCls, chOpn, chCls)
        i = IIf(j < chBr, j, chBr)

        'не найден ни один - конец
        If i = Integer.MaxValue Then Return strExpression

        If strExpression.Chars(i) = "["c Then
            If qbCnt = 0 Then qStart = i 'новая группа скобок [..] Сохраняем положение первой открывающей скобки "["
            qbCnt = qbCnt + 1 'увеличиваем на 1 кол-во вложенных скобок
        ElseIf strExpression.Chars(i) = "'"c Then
            'найдена строка '...'.
qbQuoteSearching:
            i = strExpression.IndexOf("'", i + 1) 'ищем конец строки (закрывающую кавычку "'") и продолжим работать уже за ней
            If i = -1 Then Return strExpression 'не найдена закрывающая кавычка - конец
            'Обрабатываем экранированную кавычку /'
            If strExpression.Chars(i - 1) = "/"c Then GoTo qbQuoteSearching
        Else
            'найдена "]"
            qbCnt = qbCnt - 1 'уменьшаем количество вложенных скобок
            If qbCnt = 0 Then 'последняя скобка в группе?
                If qStart > i Then Return strExpression 'если получилось, что положение первой открывающей кавычки менбше текущего положения - выход
                'заменяем содержимое кавычек (с ними включительно) на ***
                Mid$(strExpression, qStart + 1, i - qStart + 1) = StrDup(i - qStart + 1, "*")
            End If
        End If
        i = i + 1
        GoTo st
nx:
        'возвращаем результат
        Return strExpression
    End Function

    ''' <summary>
    ''' определяет, находится ли в строке простое выражение - строка, число или Да/Нет
    ''' </summary>
    ''' <param name="strExpression">строка кода</param>
    Public Function IsSimpleValue(ByVal strExpression As String) As Boolean
        If strExpression.Length = 0 Then Return True
        If strExpression = "True" OrElse strExpression = "False" Then Return True
        If strExpression.Chars(0) = "'"c AndAlso strExpression.EndsWith("'") AndAlso strExpression.IndexOf(" & ") = -1 AndAlso strExpression.IndexOf(" + ") = -1 Then Return True
        Dim res As Double
        Return Double.TryParse(strExpression, NumberStyles.Any, provider_points, res)
    End Function

    ''' <summary>
    ''' Возвращает имя и индекс переменной из массива masCode
    ''' </summary>
    ''' <param name="masCode">массив кода</param>
    ''' <param name="startIndex">индекс элемента - начала переменной (может быть именем переменной или классом Var)</param>
    ''' <param name="varName">для получения имени переменной</param>
    ''' <param name="varIndex">для получения индекса переменной</param>
    ''' <param name="varSignature">для получения сигнатуры переменной</param>
    ''' <param name="nextElementId">для получения индекса элемента, следующего за переменной</param>
    ''' <param name="returnFormat">формат, в который надо перевести значение переменной. Если символ конвертации #/$ не указан,
    ''' то значение не меняется</param>
    ''' <param name="arrParams">массив параметров, переданных коду</param>
    ''' <returns>пустую строку, в случае ошибки #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function SplitVarFromFullName(ByRef masCode() As CodeTextBox.EditWordType, ByVal startIndex As Integer, _
                                                        ByRef varName As String, ByRef varIndex As Integer, ByRef varSignature As String, _
                                                        ByRef nextElementId As Integer, ByRef returnFormat As ReturnFormatEnum, _
                                                        ByRef arrParams() As String) As String
        'Первый символ конвертации, если он, был, то уже обработан (находится в позиции startIndex он не может)
        'Var.#myVar[x] = 123
        Dim curPos As Integer = startIndex 'текущее положение в коде

        If masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
            'указан класс Var
            curPos += 2
            If curPos >= masCode.Length Then
                LAST_ERROR = "Неверная запись переменной."
                Return "#Error"
            End If
            '#myVar[x] = 123
            'Проверка на символ конвертации
            If masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Then
                returnFormat = ReturnFormatEnum.TO_NUMBER
                curPos += 1
            ElseIf masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING Then
                returnFormat = ReturnFormatEnum.TO_STRING
                curPos += 1
            End If
        End If

        'myVar[x] = 123
        If curPos >= masCode.Length OrElse masCode(curPos).wordType <> CodeTextBox.EditWordTypeEnum.W_VARIABLE Then
            LAST_ERROR = "Неверная запись переменной."
            Return "#Error"
        End If
        'если был класс Var, то он уже обработан и в позиции curPos находится имя переменной
        varName = masCode(curPos).Word.Trim

        'получаем индекс переменной, расположенный в квадратных скобках
        curPos += 1
        If curPos = masCode.Length Then
            'код заканцивается именем переменной (скобок за ней нет)
            varIndex = 0
            nextElementId = -1
        ElseIf masCode(curPos).wordType <> CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
            'за переменной нет квадратных скобок, значит индекс = 0
            varIndex = 0
            nextElementId = curPos
        Else
            'квадратные скобки есть. Индекс переменной - это результат кода между скобками
            'получаем индекс элемента, следующего сразу за закрывающей скобкой
            Dim nextPos As Integer = GetNextElementIdAfterBrackets(masCode, curPos)
            If nextPos = -2 Then Return "#Error"
            nextElementId = nextPos
            'получаем индекс последнего элемента в скобках
            If nextPos = -1 Then
                nextPos = masCode.GetUpperBound(0) - 1
            Else
                nextPos -= 2
            End If
            curPos += 1 'первый элемент в скобках
            If curPos = nextPos AndAlso masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER Then
                If masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER Then
                    'в скобках просто число - незачем выполнять GetValue. Индекс равняется следующему числу
                    varIndex = Convert.ToInt32(masCode(curPos).Word)
                ElseIf masCode(curPos).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING Then
                    'в скобках просто строка - незачем выполнять GetValue. Получаем сигнатуру
                    varSignature = masCode(curPos).Word
                End If
            Else
                'получаем результат кода в скобках
                Dim strResult As String = GetValue(masCode, curPos, nextPos, arrParams)
                If strResult = "#Error" Then Return strResult
                Dim ret As ReturnFormatEnum = Param_GetType(strResult)
                If ret <> ReturnFormatEnum.TO_NUMBER AndAlso ret <> ReturnFormatEnum.TO_STRING Then
                    LAST_ERROR = "Индекс переменной " + varName + " должен быть числом/строкой."
                    Return "#Error"
                End If
                If ret = ReturnFormatEnum.TO_NUMBER Then
                    varIndex = Convert.ToInt32(strResult)
                Else
                    varSignature = strResult
                End If
            End If
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Возвращает имя и индекс переменной по ее полному описанию, переданному в varString (Var.#varName[x])
    ''' </summary>
    ''' <param name="varString">строка с полным описанием переменной (напр., Var.#varName[x])</param>
    ''' <param name="varName">для получения имени переменной</param>
    ''' <param name="varIndex">переменная для получения индекса</param>
    ''' <param name="varSignature">для получения сигнатуры переменной</param>
    ''' <param name="returnFormat">для получения формата, в котором присвоить результат (чистло, строка, Да/Нет или не трогать)</param>
    ''' <param name="arrParams">параметры, переданные коду при выполнении</param>
    ''' <returns>#Error в случае ошибки или ничего</returns>
    ''' <remarks></remarks>
    Protected Overridable Function SplitVarFromFullName(ByVal varString As String, ByRef varName As String, ByRef varIndex As Integer, ByRef varSignature As String, _
                                                        ByRef returnFormat As MatewScript.ReturnFormatEnum, ByRef arrParams() As String) As String

        Dim varClass As String = "" 'для получения имени класса (Var, V или не указан)
        Dim qbContent As String = "" 'для получения содержимого квадратных скобок переменной

        If varString.Length = 0 Then
            LAST_ERROR = "Не указано имя переменной."
            Return ("#Error")
        End If

        'Проверяем не указан ли формат (с помощью # или $), в котором выдавать результат (число или строка)
        returnFormat = MatewScript.ReturnFormatEnum.ORIGINAL
        If varString.Chars(0) = "#"c Then
            returnFormat = MatewScript.ReturnFormatEnum.TO_NUMBER
            varString = varString.Substring(1)
        ElseIf varString.Chars(0) = "$"c Then
            returnFormat = MatewScript.ReturnFormatEnum.TO_STRING
            varString = varString.Substring(1)
        End If

        'Var.#myVar[x]
        'переменные для хранения позиции в строке . [ и пробела
        Dim pointPos, qbPos As Integer
        Dim strQbHidden As String = HideQBcontent(varString) 'строка со спрятанным содержимым квадратных скобок

        pointPos = strQbHidden.IndexOf(".") 'положение точки
        If pointPos = -1 Then pointPos = Integer.MaxValue

        qbPos = varString.IndexOf("[") 'положение [
        If qbPos = -1 Then qbPos = Integer.MaxValue

        If pointPos < qbPos Then
            'Если точка находится до квадратных скобок 
            'Значит, есть имя класса и нет [ до точки. 
            'Var.#myVar[x]
            varClass = varString.Substring(0, pointPos) 'получаем имя класса
            'Проверка на операторы преобразования типа $ и # за именем класса
            If varString.Chars(pointPos + 1) = "#"c Then
                returnFormat = MatewScript.ReturnFormatEnum.TO_NUMBER
                pointPos += 1
            ElseIf varString.Chars(pointPos + 1) = "$"c Then
                returnFormat = MatewScript.ReturnFormatEnum.TO_STRING
                pointPos += 1
            End If

            If varClass = "V" OrElse varClass = "Var" Then
                'Это явно указанная переменная
                'Var.myVar[x]
                If qbPos = Integer.MaxValue Then
                    'квадратных скобок нет
                    'myVar
                    varName = varString.Substring(pointPos + 1, varString.Length - pointPos - 1) 'получаем имя переменной
                    qbContent = "0"
                Else
                    'есть квадратные скобки
                    'myVar[x]
                    varName = varString.Substring(pointPos + 1, qbPos - pointPos - 1) 'получаем имя переменной
                    qbContent = varString.Substring(qbPos + 1, varString.Length - qbPos - 2) 'получаем содержимое квадратных скобок
                End If
            Else
                'указан другой класс
                LAST_ERROR = varString + " - не переменная."
                Return "#Error"
            End If
        ElseIf qbPos < pointPos And pointPos <> Integer.MaxValue Then
            'есть квадратные скобки до точки - неверная запись переменной Var[x].varName
            LAST_ERROR = "Неверная запись перменной."
            Return "#Error"
        ElseIf pointPos = Integer.MaxValue And qbPos <> Integer.MaxValue Then
            'имя класса нет, есть квадратные скобки myVar[x]
            varName = varString.Substring(0, qbPos)
            qbContent = varString.Substring(qbPos + 1, varString.Length - qbPos - 2)
        Else
            'имя класса нет, квадратных скобок нет myVar
            varName = varString
            qbContent = "0"
        End If

        'получаем индекс переменной
        If Integer.TryParse(qbContent, varIndex) = False Then 'для ускорения
            'считаем индекс
            qbContent = GetValue(qbContent, arrParams)
            If qbContent = "#Error" Then Return "#Error"
            Dim ret As ReturnFormatEnum = Param_GetType(qbContent)
            If ret <> MatewScript.ReturnFormatEnum.TO_NUMBER AndAlso ret <> ReturnFormatEnum.TO_STRING Then
                LAST_ERROR = "Индекс переменной - не число/строка."
                Return "#Error"
            End If
            If ret = MatewScript.ReturnFormatEnum.TO_NUMBER Then
                varIndex = Convert.ToDouble(qbContent)
            Else
                varSignature = qbContent
            End If
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Собирает из блоков HTML и Wrap готовый текст, выполнив в нем код между тэгами <exec>...</exec>
    ''' </summary>
    ''' <param name="exCode">код</param>
    ''' <param name="startLine">инлекс линии с началом блока</param>
    ''' <param name="finalLine">последняя линия внутри блока</param>
    ''' <param name="arrParams">параметры, с которыми запущен код</param>
    ''' <param name="checkForBlockEnding">Поиск проводится внутри блока HTML/Wrap и надо искать End HTML/Wrap или же нет</param>
    ''' <param name="selector">Номер селектора, текст которого надо вывести</param>
    ''' <returns>готовую строку, в случае ошибки - #Error</returns>
    Public Overridable Function MakeExec(ByRef exCode As List(Of ExecuteDataType), ByVal startLine As Integer, ByRef finalLine As Integer, _
                                            ByRef arrParams() As String, Optional ByVal checkForBlockEnding As Boolean = True, Optional ByVal selector As Integer = 0) As String
        finalLine = -1 'сбрасываем значение последней строки блока

        Dim curLine As Integer
        If checkForBlockEnding Then
            curLine = startLine + 1 'Текущая линия, которую проверяем. В начале - первая линия после блока
        Else
            curLine = startLine
        End If
        Dim curCode() As CodeTextBox.EditWordType 'код текущей строки
        Dim strResult As New System.Text.StringBuilder

        Dim exCopy As New List(Of ExecuteDataType) 'для хранения и передачи в ExecuteCode данных из <exec>
        Dim execStart As Integer = -1 'индекс в строке кода элемента, сразу после тэга
        Dim insideExec As Boolean = False 'внутри <exec>?
        Dim codeResult As String 'для получения результата ExecuteCode

        Dim strSelector As String = ""
        If selector > 0 Then
            curCode = exCode(curLine).Code
            strSelector = "#" & selector.ToString & ":"
            If IsNothing(curCode) = False AndAlso curCode.Length > 0 AndAlso curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA AndAlso curCode(0).Word.StartsWith("#") AndAlso curCode(0).Word.StartsWith(strSelector) Then
                'селектор уже найден.
            Else
                Dim wasFound As Boolean = False
                For lId As Integer = curLine To exCode.Count - 1
                    curCode = exCode(lId).Code
                    If IsNothing(curCode) OrElse curCode.Count = 0 Then Continue For
                    If checkForBlockEnding AndAlso insideExec = False AndAlso curCode.Length > 1 AndAlso curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_CYCLE_END AndAlso _
                        (curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_HTML OrElse curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_WRAP) Then
                        'нашли End Wrap/HTML
                        'селектор так и не найден - работаем так, как-будто селекторов нет
                        Exit For
                    ElseIf curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA AndAlso curCode(0).Word.StartsWith(strSelector) Then
                        'найден селектор
                        curLine = lId
                        wasFound = True
                        Exit For
                    End If
                Next lId
                If Not wasFound Then selector = 0
            End If

        End If
        startLine = curLine

        'выполняем цикл, пока не найдем End Wrap/HTML
        Do
            If curLine > exCode.Count - 1 Then
                If checkForBlockEnding Then
                    LAST_ERROR = "Не найден конец блока Wrap/HTML."
                    Return "#Error"
                Else
                    Return strResult.ToString
                End If
            End If

            curCode = exCode(curLine).Code 'текущая строка кода

            If IsNothing(curCode) OrElse curCode.Length = 0 Then
                'пустая строка
                If Not insideExec Then strResult.AppendLine()
                curLine += 1
                Continue Do
            End If

            If checkForBlockEnding AndAlso insideExec = False AndAlso curCode.Length > 1 AndAlso curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_CYCLE_END AndAlso _
                (curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_HTML OrElse curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_WRAP) Then
                'нашли End Wrap/HTML
                finalLine = curLine - 1
                Return strResult.ToString 'возвращаем результирующую строку
            End If

            If insideExec Then
                'внутри exec - первый элемент строки по-умолчанию код
                execStart = 0
            Else
                'не в exec - первый элемент строки по-умолчанию не код
                execStart = -1
                If curLine > startLine + 1 Then strResult.AppendLine()
            End If
            'в цикле проверяем каждое слово строки кода
            For i As Integer = 0 To curCode.GetUpperBound(0)
                If curCode(i).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA OrElse curCode(i).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_TAG Then
                    'текущее слово - не код
                    If Not insideExec Then
                        '...и мы не внутри exec - добавляем слово в результирующую строку
                        Dim wrd As String = curCode(i).Word
                        If curLine = startLine AndAlso selector > 0 AndAlso wrd.StartsWith(strSelector) Then
                            wrd = wrd.Substring(strSelector.Length)
                        ElseIf selector > 0 AndAlso wrd.StartsWith("#") Then
                            Dim res As Integer = IsItSelector(wrd, 0)
                            If res > 0 Then
                                'найден новый селектор - конец текста
                                Return strResult.ToString
                            End If
                        End If

                        If wrd.EndsWith("<exec>") Then wrd = wrd.Substring(0, wrd.Length - 6)
                        strResult.Append(wrd)
                    Else
                        '...и мы внутри exec - конец блока exec
                        'html<exec>...</exec>html
                        If i > execStart Then
                            'вырезаем код из exec для выполнения (отделяем его от html)
                            Dim codeCopy() As CodeTextBox.EditWordType = Nothing
                            Array.Resize(codeCopy, i - execStart)
                            Array.Copy(curCode, execStart, codeCopy, 0, codeCopy.Length)
                            'приводим код в вид, пригодный для передачи в ExecuteCode
                            exCopy.Add(New ExecuteDataType(codeCopy, exCode(curLine).lineId))
                        End If

                        'выполняем код
                        Dim curLineCopy As Integer = CURRENT_LINE
                        codeStack.Push("Код <exec>")
                        If questEnvironment.EDIT_MODE Then
                            codeResult = "<code>&lt;скрипт&gt;"
                            For j As Integer = 0 To exCopy.Count - 1
                                If IsNothing(exCopy(j).Code) OrElse exCopy(j).Code.Count = 0 Then Continue For
                                For u As Integer = 0 To exCopy(j).Code.Count - 1
                                    codeResult &= exCopy(j).Code(u).Word
                                Next u
                            Next j
                            codeResult &= "&lt;скрипт/&gt;</code>"
                        Else
                            codeResult = ExecuteCode(exCopy, arrParams)
                        End If
                        codeStack.Pop()

                        If codeResult = "#Error" Then LAST_ERROR = "#DON'T_SHOW#Ошибка в тэге <exec>:" + LAST_ERROR
                        CURRENT_LINE = curLineCopy
                        EXIT_CODE = False

                        'добавляем результат кода в результирующую строку
                        strResult.Append(PrepareStringToPrint(codeResult, arrParams))
                        'и оставшуюся часть строки
                        Dim finalString As String = curCode(i).Word
                        If finalString.StartsWith("</exec>", StringComparison.CurrentCultureIgnoreCase) Then finalString = finalString.Substring(7)
                        If finalString.Length > 0 Then strResult.Append(finalString)

                        exCopy = New List(Of ExecuteDataType)

                        insideExec = False
                    End If
                Else
                    'текущее слово - код
                    If Not insideExec Then
                        '... и мы не внутри exec - начало кода
                        insideExec = True
                        execStart = i
                    End If
                End If
            Next

            If insideExec Then
                'строка кода закончилась, а код не закрылся - значит это многострочный блок exec
                'вставляем в exCopy код из текущей строки (от execStart и до конца)
                Dim codeCopy() As CodeTextBox.EditWordType = Nothing
                Array.Resize(codeCopy, curCode.Length - execStart)
                'Array.Resize(codeCopy, curCode.GetUpperBound(0) - execStart)
                Array.Copy(curCode, execStart, codeCopy, 0, codeCopy.Length)
                exCopy = New List(Of ExecuteDataType)
                exCopy.Add(New ExecuteDataType(codeCopy, exCode(curLine).lineId))
            ElseIf exCode(curLine).Code.Last.Word.TrimEnd.EndsWith("<exec>") Then
                insideExec = True
            End If

            curLine += 1
        Loop

        Return strResult.ToString
    End Function

    ''' <summary>
    ''' Выполняет код между тэгами <exec>...</exec>
    ''' </summary>
    ''' <param name="strText">полный текст, внутри которого могут быть теги exec</param>
    ''' <param name="arrParams">параметры, переданные коду при выполнении</param>
    ''' <returns>результирующую строку, содержащую вместо тэгов exec результат выполнения кода между ними</returns>
    ''' <remarks></remarks>
    Public Overridable Function MakeExec(ByVal strText As String, ByRef arrParams() As String, Optional ByVal selector As Integer = 0) As String
        '<exec>...</exec>
        Dim execStart As Integer = 0
        Dim execEnd As Integer = 0
        Dim strExec As String = ""
        If selector > 0 Then
            Dim strSelector As String = "#" & selector.ToString & ":"
            Dim pos As Integer = strText.IndexOf(strSelector)
            If pos > -1 Then
                'селектор найден
                strText = strText.Substring(pos + strSelector.Length) 'теперь у нас текст, начиная от нужно места (от указанного селектора)

                pos = -1
                Do
                    pos = strText.IndexOf("#", pos + 1)
                    If pos = -1 Then Exit Do
                    Dim res As Integer = IsItSelector(strText, pos)
                    If res = 0 Then Continue Do
                    strText = Left(strText, pos)
                    'теперь у нас только текст нашего селектора
                    Exit Do
                Loop
            End If
        End If
        Do
            execStart = strText.IndexOf("<exec>")
            If execStart = -1 Then Return strText
            execEnd = strText.IndexOf("</exec>", execStart + 6)
            If execEnd = -1 Then execEnd = strText.Length - 1
            strExec = strText.Substring(execStart + 6, execEnd - execStart - 6)

            'Dim arrVars As Dictionary(Of String, Array) = Nothing, arrInfo() As variableEditorInfoType = {}
            'csLocalVariables.CopyVariables(arrVars, arrInfo)
            'csLocalVariables.KillVars()
            Dim curLine As Integer = CURRENT_LINE

            codeStack.Push("Код <exec>")
            Dim strResult As String = ExecuteString(Split(strExec, vbNewLine), arrParams)
            codeStack.Pop()

            If strResult = "#Error" Then LAST_ERROR = "#DON'T_SHOW#Ошибка в тэге <exec>:" + LAST_ERROR
            CURRENT_LINE = curLine
            EXIT_CODE = False
            'csLocalVariables.RestoreVariables(arrVars, arrInfo)

            strExec = PrepareStringToPrint(strResult, arrParams)

            strText = strText.Substring(0, execStart) + strExec + strText.Substring(execEnd + 7)
        Loop
    End Function

    ''' <summary>
    ''' Запускаем функцию пользователя, сохраненную предварительно в массив functionsHash
    ''' </summary>
    ''' <param name="funcName">имя функции</param>
    ''' <param name="funcParams">параметры, передаваемые функции</param>
    ''' <returns>результат функции или #Error в случае ошибки</returns>
    ''' <remarks></remarks>
    Public Function CallFunction(ByVal funcName As String, ByRef funcParams() As String) As String
        Dim arrVars As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
        'создаем копию локальных переменных в arrVars 
        csLocalVariables.CopyVariables(arrVars)
        'удаляем текущие локальные переменные
        csLocalVariables.KillVars()
        'сохраняем текущее значение CURRENT_LINE (номер строки кода, который исполняется)
        Dim curLine As Integer = CURRENT_LINE
        If funcName.StartsWith("'") Then funcName = mScript.PrepareStringToPrint(funcName, funcParams, True)
        If functionsHash.ContainsKey(funcName) Then
            codeStack.Push("Функция " + funcName) 'добвляем стэк
            Dim strResult As String = ExecuteCode(functionsHash(funcName).ValueExecuteDt, funcParams, True) 'выполняем код функции (как первичный)
            codeStack.Pop() 'убираем из стэка
            If strResult = "#Error" Then LAST_ERROR = "#DON'T_SHOW#" + LAST_ERROR
            Return strResult
        Else
            LAST_ERROR = "Функция Писателя " + funcName + " не найдена."
            Return "#Error"
        End If
        'восстанавливаем локальные переменные и положение текущей строки кода
        CURRENT_LINE = curLine
        csLocalVariables.RestoreVariables(arrVars)
    End Function

    ''' <summary>
    ''' Функция возвращает четное ли количество кавычек в ' в строке, игнорируя экранированные /'
    ''' </summary>
    ''' <param name="strExpression">выражение</param>
    Public Function IsQuotesCountEven(ByVal strExpression As String) As Boolean
        Dim qPos As Integer, qCnt As Integer = 0

        qPos = strExpression.IndexOf("'"c) 'Находим первую
        Do While qPos > -1
            'Выполняем цикл, пока кавычки есть
            If Not (qPos > 0 AndAlso strExpression.Chars(qPos - 1) = "/"c) Then qCnt += 1 'Если кавычка не экранированная, то +1
            qPos = strExpression.IndexOf("'"c, qPos + 1) 'Находим новую
        Loop

        Dim i As Decimal = qCnt / 2 'полусумма кавычек будет целым числом, если сумма - четное
        Return (IIf(i = Math.Round(i), True, False))
    End Function

    ''' <summary>
    ''' Функция подготавливает строку к выводу на экран, т. е. убирает кавычки и выполняет исполняемые строки
    ''' </summary>
    ''' <param name="strToPrint">строка, выводимая на экран</param>
    ''' <param name="arrParams">массив параметров, переданных коду</param>
    ''' <param name="blnExecuteString">выполнять ли код, если строка исполняемая</param>
    ''' <returns>строку для вывода на экран</returns>
    ''' <remarks></remarks>
    Public Function PrepareStringToPrint(ByVal strToPrint As String, ByRef arrParams() As String, Optional ByVal blnExecuteString As Boolean = True) As String
        'если параметр пуст или не является строкой - передаем то, что получили
        If String.IsNullOrEmpty(strToPrint) OrElse strToPrint.Chars(0) <> "'"c Then Return strToPrint

        Dim cRes As ContainsCodeEnum
        If blnExecuteString Then cRes = IsPropertyContainsCode(strToPrint)
        'убираем и разблокируем кавычки
        strToPrint = strToPrint.Substring(1, strToPrint.Length - 2).Replace("/'", "'")

        'If blnExecuteString = True AndAlso strToPrint.Length > 0 AndAlso strToPrint.Chars(0) = "?"c Then
        If blnExecuteString = True AndAlso cRes = ContainsCodeEnum.EXECUTABLE_STRING Then
            'исполняемая строка
            strToPrint = strToPrint.Substring(1).Trim
            Return PrepareStringToPrint(GetValue(strToPrint, arrParams), arrParams, True)
        End If
        Return strToPrint
    End Function
#End Region

#Region "Blocks"
    ''' <summary>
    ''' Функция обрабатывает многострочный блок if ... Then и возвращает индекс последней строки кода внутри блока, которую надо обработать,
    ''' а также индекс строки с End If данного блока.
    ''' </summary>
    ''' <param name="exCode">Список со всем исполняемым кодом</param>
    ''' <param name="startLine">Линия, на которой находится If ... Then. Заменяется первой линией внутри блока, которую надо 
    ''' выполнить</param>
    ''' <param name="finalLine">Последняя линия внутри блока, которую надо выполнить. Если блока на выполнение не найдено 
    ''' (не подошло ни одно условие), равен -1</param>
    ''' <param name="endIfLine">Линия с End If</param>
    ''' <param name="arrParams">Параметры, переданные коду при запуске</param>
    ''' <returns>Пустую строку, при ошибке #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockIF(ByRef exCode As List(Of ExecuteDataType), ByRef startLine As Integer, _
                                           ByRef finalLine As Integer, ByRef endIfLine As Integer, ByRef arrParams() As String) As String
        'проверка на правильность первой строки (If ... Then) была проведена ранее и здесь не повторяется
        finalLine = -1
        Dim ifBalance As Integer = 1 'Баланс между If Then / End If (т. е. сколько открыто внутренних блоков If Then).
        'Исходно мы внутри блока
        Dim curLine As Integer = startLine + 1 'Текущая линия, которую проверяем. В начале - первая линия после If ... Then
        Dim curCode() As CodeTextBox.EditWordType 'код текущей строки
        Dim ifCondition As Boolean = False 'Результат последнего условия (True/False)
        'получаем результат условия между If и Then
        Dim strResult As String = GetValue(exCode(startLine).Code, 1, exCode(startLine).Code.GetUpperBound(0) - 1, arrParams)
        If strResult = "#Error" Then Return strResult
        If strResult = "True" Then ifCondition = True

        'Выполняем пока наш блок не закроется (пока не найдем End If НАШЕГО (а не вложенного) блока)
        Do Until ifBalance = 0
            If curLine > exCode.Count - 1 Then
                LAST_ERROR = "Не найден конец блока If."
                Return "#Error"
            End If

            curCode = exCode(curLine).Code
            If IsNothing(curCode) OrElse curCode.Length = 0 Then
                'поустая строка
                curLine += 1
                Continue Do
            End If

            If curCode(0).Word = "Else" Then
                'Else
                If ifBalance = 1 Then
                    If ifCondition = False Then
                        'предыдущие условия возвращали False, значит надо выполнить то, что за Else
                        ifCondition = True
                        startLine = curLine
                    Else
                        'сохраняем последнюю линию выполняемого блока
                        If finalLine = -1 Then finalLine = curLine - 1
                    End If
                End If
            ElseIf curCode(0).Word = "If " AndAlso curCode(curCode.GetUpperBound(0)).Word = "Then" Then
                'вложенный многострочный If ... Then
                ifBalance += 1
            ElseIf curCode.Length = 2 AndAlso curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_CYCLE_END AndAlso curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_IF Then
                'End If
                If ifBalance = 1 Then
                    'это End If нашего блока
                    If ifCondition = True AndAlso finalLine = -1 Then finalLine = curLine - 1
                    endIfLine = curLine
                    Return ""
                End If
                ifBalance -= 1
            ElseIf curCode(0).Word.Trim = "ElseIf" Then
                'ElseIf
                If ifBalance = 1 Then
                    If ifCondition = False Then
                        'до сих пор предыдущие условия возвращали False. Проверяем это условие
                        If curCode.Length = 1 Then
                            LAST_ERROR = "Нет условия после оператора ElseIf."
                            Return "#Error"
                        ElseIf curCode(curCode.GetUpperBound(0)).Word <> "Then" Then
                            LAST_ERROR = "Отсутствует ключевое слово Then в конструкции ElseIf."
                            Return "#Error"
                        End If
                        strResult = GetValue(curCode, 1, curCode.GetUpperBound(0) - 1, arrParams)
                        If strResult = "#Error" Then Return strResult
                        If strResult = "True" Then
                            'условие вернуло True
                            ifCondition = True
                            startLine = curLine
                        End If
                    Else
                        'сохраняем последнюю линию выполняемого блока
                        If finalLine = -1 Then finalLine = curLine - 1
                    End If
                End If
            End If

            curLine += 1 'к новой линии
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает многострочный блок if ... Then и возвращает индекс последней строки кода внутри блока, которую надо обработать,
    ''' а также индекс строки с End If данного блока.
    ''' </summary>
    ''' <param name="masCode">массив со всем выполняемым кодом</param>
    ''' <param name="firstCondition">результат первого условия между If и Then</param>
    ''' <param name="startPos">в начале - индекс строки с обрабатываемым блоком If ... Then. Затем передает индекс первой строки кода внутри блока, которую надо обработать</param>
    ''' <param name="finalPos">индекс последней строки кода внутри блока, которую надо обработать</param>
    ''' <param name="endIfPos">индекс строки с End If данного блока</param>
    ''' <param name="arrParams">параметры, с которыми запущен код</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockIF(ByRef masCode() As DictionaryEntry, ByVal firstCondition As Boolean, _
                             ByRef startPos As Integer, ByRef finalPos As Integer, ByRef endIfPos As Integer, _
                             ByRef arrParams() As String) As String
        finalPos = -1 'сбрасываем значение конца исполняемого блока

        Dim ifBalance As Integer = 1 'Баланс между If Then / End If (т. е. сколько открыто внутренних блоков If Then)
        Dim curLine As Integer = startPos + 1 'номер строки, с которой работаем
        Dim curString As String = "" 'текст (код) строки
        Dim saveStartPos As String = startPos 'переменная для хранения начального значения стартовой позиции (положение самого блока If Then)

        If firstCondition Then startPos += 1 'Если начальное условие соблюдено - начало исполняемого блока сразу за If Then
        'Выполняем цикл, пока баланс между If Then и End If не станет равен 0 (т. е., пока не найдем наш End If)
        Do Until ifBalance = 0
            'ищем первый оператор
            curLine = Array.FindIndex(Of DictionaryEntry)(masCode, curLine, AddressOf BlockIfSearch)
            If curLine = -1 Then
                'Не найден ни один, а баланс еще не 0 - значит ошибка
                LAST_ERROR = "В блоке If ... Then ... End If не найден End If."
                Return "#Error"
            End If

            If masCode(curLine).Value = "End If" Then 'найденный оператор - End If
                ifBalance -= 1 'меняем баланс
                If ifBalance = 0 AndAlso finalPos = -1 Then finalPos = curLine - 1
            ElseIf masCode(curLine).Value.ToString.StartsWith("Else") Then 'найденный оператор - Else / ElseIf
                If ifBalance = 1 Then 'работаем только если мы внутри нашего блока (а не вложенного)
                    If startPos <> saveStartPos Then
                        'если начало блока установлено, а конец еще нет - устанавливаем конец на строку выше Else (или ElseIf)
                        If finalPos = -1 Then finalPos = curLine - 1
                    Else
                        'начало блока еще не установлено
                        curString = masCode(curLine).Value 'текст текущей строки
                        If curString = "Else" Then
                            startPos = curLine + 1 'устанавливаем начало блока сразу вслед за Else
                        Else
                            'опретор ElseIf
                            If curString.EndsWith(" Then") Then curString = curString.Substring(0, curString.Length - 5) 'убираем конечный Then
                            curString = GetValue(curString.Substring(7), arrParams) 'получаем результат условия ElseIf
                            If curString = "#Error" Then Return "#Error"
                            If curString <> "False" Then startPos = curLine + 1 'Если True - условие соблюдено. Устанавливаем начало блока вслед за ElseIf
                        End If
                    End If
                End If
            Else 'найденный оператор - еще один внутренний If ... Then
                ifBalance += 1 'меняем баланс
            End If
            curLine += 1 'будем искать дальше с новой строки
        Loop

        If startPos = saveStartPos Then finalPos = -1 'если начало блока так и найдено - сбрасываем конец блока (нет кода для исполнения)
        If finalPos < startPos Then finalPos = -1 'на случай, если между If ... Then и Else нет ни одной исполняемой строки - сбрасываем конец блока (нет кода для исполнения)
        endIfPos = curLine - 1 'сохраняем позицию конечного End If
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает цикл For ... Next и возвращает индекс последней строки кода внутри цикла
    ''' </summary>
    ''' <param name="exCode">Список со всем исполняемым кодом</param>
    ''' <param name="startLine">индекс строки с обрабатываемым блоком For ... Next</param>
    ''' <param name="finalLine">индекс последней строки кода внутри цикла</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockFor(ByRef exCode As List(Of ExecuteDataType), ByRef startLine As Integer, _
                                             ByRef finalLine As Integer) As String
        finalLine = -1
        Dim forBalance As Integer = 1 'Баланс между If Then / End If (т. е. сколько открыто внутренних блоков If Then).
        'Исходно мы внутри блока
        Dim curLine As Integer = startLine + 1 'Текущая линия, которую проверяем. В начале - первая линия после If ... Then
        Dim curCode() As CodeTextBox.EditWordType 'код текущей строки

        Do Until forBalance = 0
            If curLine > exCode.Count - 1 Then
                LAST_ERROR = "Не найден конец блока If."
                Return "#Error"
            End If

            curCode = exCode(curLine).Code
            If IsNothing(curCode) OrElse curCode.Length = 0 Then
                'поустая строка
                curLine += 1
                Continue Do
            End If

            If curCode(0).Word.Trim = "Next" Then
                If forBalance = 1 Then
                    finalLine = curLine - 1
                    Return ""
                End If
                forBalance -= 1
            ElseIf curCode(0).Word = "For " Then
                forBalance += 1
            End If

            curLine += 1
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает цикл For ... Next и возвращает индекс последней строки кода внутри цикла
    ''' </summary>
    ''' <param name="masCode">массив со всем выполняемым кодом</param>
    ''' <param name="startPos">индекс строки с обрабатываемым блоком For ... Next</param>
    ''' <param name="finalPos">индекс последней строки кода внутри цикла</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockFor(ByRef masCode() As DictionaryEntry, ByVal startPos As Integer, _
                                            ByRef finalPos As Integer) As String
        finalPos = -1 'сбрасываем значение конца исполняемого блока

        Dim forBalance As Integer = 1 'Баланс между For... / Next (т. е. сколько открыто внутренних циклов For ... Next)
        Dim curLine As Integer = startPos + 1 'номер строки, с которой работаем
        Dim curString As String = "" 'текст (код) строки
        Dim saveStartPos As String = startPos 'переменная для хранения начального значения стартовой позиции (положение самого цикла For ... Next)

        startPos += 1 'начало исполняемого блока сразу за For ... Next
        'Выполняем цикл, покка баланс между For... и Next не станет равен 0 (т. е., пока не найдем наш Next)
        Do Until forBalance = 0
            'ищем первый оператор
            curLine = Array.FindIndex(Of DictionaryEntry)(masCode, curLine, AddressOf BlockForSearch)
            If curLine = -1 Then
                'Не найден ни один, а баланс еще не 0 - значит ошибка
                LAST_ERROR = "Не найден конец цикла в блоке For ... Next."
                Return "#Error"
            End If

            If masCode(curLine).Value.ToString.StartsWith("Next") Then 'найденный оператор - Next
                forBalance -= 1 'меняем баланс
                If forBalance = 0 Then
                    finalPos = curLine - 1
                    Return ""
                End If
            Else 'найденный оператор - еще один внутренний For ... Next
                forBalance += 1 'меняем баланс
            End If
            curLine += 1 'будем искать дальше с новой строки
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает блок Select Case / Switch и возвращает индекс первой строки кода внутри блока, которую надо обработать, 
    ''' а также индекс последней строки кода внутри блока, которую надо обработать.
    ''' </summary>
    ''' <param name="exCode">код</param>
    ''' <param name="startLine">в начале - индекс строки с обрабатываемым блоком Select Case. Затем передает индекс первой строки кода внутри блока, которую надо обработать</param>
    ''' <param name="finalLine">индекс последней строки кода внутри блока, которую надо обработать</param>
    ''' <param name="endSelectLine">индекс строки с End Select данного блока</param>
    ''' <param name="arrParams">параметры, с которыми запущен код</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockSwitch(ByVal exCode As List(Of ExecuteDataType), ByRef startLine As Integer, _
                                                ByRef finalLine As Integer, ByRef endSelectLine As Integer, ByRef arrParams() As String) As String
        finalLine = -1
        Dim switchBalance As Integer = 1 'Баланс между Select / End Select (т. е. сколько открыто внутренних блоков Select).
        'Исходно мы внутри блока
        Dim curLine As Integer = startLine + 1 'Текущая линия, которую проверяем. В начале - первая линия после Select
        Dim saveStartLine As String = startLine  'переменная для хранения начального значения стартовой позиции (положение самого блока Select Case)
        Dim curCode() As CodeTextBox.EditWordType 'код текущей строки
        Dim switchQuery As String = "False" 'результат условия (как строка)
        Dim switchQueryCode As New CodeTextBox.EditWordType ''результат условия (как элемент кода)

        'получаем результат условия Swith/Select Case
        curCode = exCode(startLine).Code
        If curCode(0).Word = "Select " Then
            switchQuery = GetValue(curCode, 2, -1, arrParams)
        ElseIf curCode(0).Word = "Switch " Then
            switchQuery = GetValue(curCode, 1, -1, arrParams)
        End If
        If switchQuery = "#Error" Then Return switchQuery
        switchQueryCode.Word = switchQuery
        'сохраняет результат условия как элемент кода
        Select Case Param_GetType(switchQuery)
            Case ReturnFormatEnum.TO_NUMBER
                switchQueryCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER
            Case ReturnFormatEnum.TO_STRING
                switchQueryCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING
            Case ReturnFormatEnum.TO_BOOL
                switchQueryCode.wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL
            Case Else
                LAST_ERROR = "Ошибка в условии блока Slect Case."
                Return "#Error"
        End Select

        'Выполняем пока наш блок не закроется (пока не найдем End Select НАШЕГО (а не вложенного) блока)
        Do Until switchBalance = 0
            If curLine > exCode.Count - 1 Then
                LAST_ERROR = "Не найден конец блока End Select."
                Return "#Error"
            End If

            curCode = exCode(curLine).Code
            If IsNothing(curCode) OrElse curCode.Length = 0 Then
                'пустая строка
                curLine += 1
                Continue Do
            End If

            If curCode.Length > 1 AndAlso curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_CYCLE_END AndAlso curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_SWITCH Then
                'End Select
                switchBalance -= 1
                If switchBalance = 0 Then
                    If finalLine = -1 Then finalLine = curLine - 1
                    Exit Do
                End If
            ElseIf curCode(0).Word.Trim = "Case" AndAlso switchBalance = 1 Then
                'Case ...
                If curCode.Length < 2 Then
                    LAST_ERROR = "Неверный синтаксис блока Select. За оператором Case должно следовать условие."
                    Return "#Error"
                End If
                If startLine <> saveStartLine Then
                    'начало блока установлено, а конец еще нет (выше - блок, подошедший по условию). Сохраняем нижнюю границу блока
                    If finalLine = -1 Then finalLine = curLine - 1
                Else
                    'начало блока еще не установлено (не было блока с подходящим условием)
                    If curCode(1).Word = "Else" Then 'Case Else
                        startLine = curLine + 1 'устанавливаем начало блока сразу вслед за Case Else
                    Else
                        'оператор Case ...
                        'Должно работать Case 4 ... Case <> 4 ... Case 3 To 5, 6, 7, > 10 And x != C[1, 2].Max(5, 67)
                        Dim curCaseCode() As CodeTextBox.EditWordType = Nothing 'будет хранится код с условием за оператором Case, готовый
                        'для передачи GetValue
                        Dim curQueryStart As Integer = 1 'индекс элемента с началом условия (сразу за Case или за запятой)
                        Dim ToPos As Integer = -1 'положение оператора То или -1, если его нет
                        Dim qbBalance As Integer = 0, obBalance As Integer = 0 'баланс скобок. Если не равны 0, то запятые не относятся к перечислению условий Case
                        Dim qResult As String = "" 'для получения результата сравнения условия блока с текущим условием Case
                        'проходим весь код за Case, перебирая все условия (разделяемые запятыми)
                        For i As Integer = 1 To curCode.GetUpperBound(0)
                            Select Case curCode(i).wordType
                                'баланс скобок
                                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                                    obBalance += 1
                                Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                                    obBalance -= 1
                                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                                    qbBalance += 1
                                Case CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                                    qbBalance -= 1
                                Case CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR
                                    'положение оператора То
                                    ToPos = i
                                Case CodeTextBox.EditWordTypeEnum.W_COMMA
                                    If qbBalance <> 0 OrElse obBalance <> 0 Then Continue For
                                    'запятая-разделитель условий Case
                                    If i <= curQueryStart Then
                                        LAST_ERROR = "Ошибка в условии Case."
                                        Return "#Error"
                                    End If
                                    If curCode(curQueryStart).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE OrElse _
                                        curCode(curQueryStart).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL Then
                                        'Case <> 5
                                        If ToPos = -1 Then
                                            '"To" нету. Создаем код "Условие блока + условие Case"
                                            Array.Resize(curCaseCode, 1 + i - curQueryStart)
                                            curCaseCode(0) = switchQueryCode
                                            Array.Copy(curCode, curQueryStart, curCaseCode, 1, curCaseCode.Length - 1)
                                        Else
                                            LAST_ERROR = "Неверная записть условия Case. Оператор То в таком виде использовать нельзя."
                                            Return "#Error"
                                        End If
                                    Else
                                        'Case 5
                                        If ToPos = -1 Then
                                            '"To" нету. Создаем код "Условие блока + "=" + условие Case"
                                            Array.Resize(curCaseCode, 2 + i - curQueryStart)
                                            curCaseCode(0) = switchQueryCode
                                            curCaseCode(1) = New CodeTextBox.EditWordType With {.Word = " = ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL}
                                            Array.Copy(curCode, curQueryStart, curCaseCode, 2, curCaseCode.Length - 2)
                                        Else
                                            'caseQ >= ... And caseQ <= ... (5 новых элементов)
                                            'Есть "То". Создаем код "Условие блока + ">=" + условие Case до То + "And " + Условие блока + "<=" условие Case после То"
                                            Array.Resize(curCaseCode, 5 + i - curQueryStart - 1)
                                            curCaseCode(0) = switchQueryCode
                                            curCaseCode(1) = New CodeTextBox.EditWordType With {.Word = " >= ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE}
                                            Array.Copy(curCode, curQueryStart, curCaseCode, 2, ToPos - curQueryStart)
                                            Dim afterTo As Integer = ToPos - curQueryStart + 2
                                            curCaseCode(afterTo) = New CodeTextBox.EditWordType With {.Word = " And ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC}
                                            curCaseCode(afterTo + 1) = switchQueryCode
                                            curCaseCode(afterTo + 2) = New CodeTextBox.EditWordType With {.Word = " <= ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE}
                                            Array.Copy(curCode, ToPos + 1, curCaseCode, afterTo + 3, i - ToPos - 1)
                                        End If
                                    End If
                                    'получаем результат условия
                                    qResult = GetValue(curCaseCode, 0, -1, arrParams)
                                    If qResult = "#Error" Then Return qResult
                                    'если True - то это подходящий блок
                                    If qResult = "True" Then Exit For
                                    curCaseCode = Nothing
                                    curQueryStart = i + 1 'следующее условие начинается за этой запятой
                                    ToPos = -1
                            End Select
                        Next
                        If qResult <> "True" Then
                            'обрабатываем последнее условие за Case
                            If qbBalance <> 0 OrElse obBalance <> 0 Then
                                LAST_ERROR = "Непарное число скобок."
                                Return "#Error"
                            End If
                            If curCode(curQueryStart).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE OrElse _
                                curCode(curQueryStart).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL Then
                                'Case <> 5
                                If ToPos = -1 Then
                                    '"To" нету. Создаем код "Условие блока + условие Case"
                                    Array.Resize(curCaseCode, 1 + curCode.Length - curQueryStart)
                                    curCaseCode(0) = switchQueryCode
                                    Array.Copy(curCode, curQueryStart, curCaseCode, 1, curCaseCode.Length - 1)
                                Else
                                    LAST_ERROR = "Неверная записть условия Case. Оператор То в таком виде использовать нельзя."
                                    Return "#Error"
                                End If
                            Else
                                'Case 5
                                If ToPos = -1 Then
                                    '"To" нету. Создаем код "Условие блока + "=" + условие Case"
                                    Array.Resize(curCaseCode, 2 + curCode.Length - curQueryStart)
                                    curCaseCode(0) = switchQueryCode
                                    curCaseCode(1) = New CodeTextBox.EditWordType With {.Word = " = ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL}
                                    Array.Copy(curCode, curQueryStart, curCaseCode, 2, curCaseCode.Length - 2)
                                Else
                                    'caseQ >= ... And caseQ <= ... (5 новых элементов)
                                    'Есть "То". Создаем код "Условие блока + ">=" + условие Case до То + "And " + Условие блока + "<=" условие Case после То"
                                    Array.Resize(curCaseCode, 5 + curCode.Length - curQueryStart - 1)
                                    curCaseCode(0) = switchQueryCode
                                    curCaseCode(1) = New CodeTextBox.EditWordType With {.Word = " >= ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE}
                                    Array.Copy(curCode, curQueryStart, curCaseCode, 2, ToPos - curQueryStart)
                                    Dim afterTo As Integer = ToPos - curQueryStart + 2
                                    curCaseCode(afterTo) = New CodeTextBox.EditWordType With {.Word = " And ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC}
                                    curCaseCode(afterTo + 1) = switchQueryCode
                                    curCaseCode(afterTo + 2) = New CodeTextBox.EditWordType With {.Word = " <= ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE}
                                    Array.Copy(curCode, ToPos + 1, curCaseCode, afterTo + 3, curCode.Length - ToPos - 1)

                                    'Array.Resize(curCaseCode, 5 + curCode.Length - curQueryStart)
                                    'curCaseCode(0) = switchQueryCode
                                    'curCaseCode(1) = New CodeTextBox.EditWordType With {.Word = " >= ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE}
                                    'Array.Copy(curCode, 2, curCaseCode, 2, ToPos - curQueryStart)
                                    'Dim afterTo As Integer = ToPos - curQueryStart + 1
                                    'curCaseCode(afterTo) = New CodeTextBox.EditWordType With {.Word = " And ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC}
                                    'curCaseCode(afterTo + 1) = switchQueryCode
                                    'curCaseCode(afterTo + 2) = New CodeTextBox.EditWordType With {.Word = " >= ", .wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE}
                                    'Array.Copy(curCode, ToPos + 1, curCaseCode, afterTo + 3, curCode.Length - ToPos)
                                End If
                            End If
                            'Получаем результат условия
                            qResult = GetValue(curCaseCode, 0, -1, arrParams)
                            If qResult = "#Error" Then Return qResult
                            curCaseCode = Nothing
                        End If
                        If qResult = "True" Then
                            startLine = curLine + 1
                        End If
                    End If
                End If
            End If

            curLine += 1
        Loop

        If startLine = saveStartLine Then finalLine = -1 'если начало блока так и найдено - сбрасываем конец блока (нет кода для исполнения)
        If finalLine < startLine Then finalLine = -1 'на случай, если между Case ... и End Select нет ни одной исполняемой строки - сбрасываем конец блока (нет кода для исполнения)
        endSelectLine = curLine 'сохраняем позицию конечного End Select
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает блок Select Case / Switch и возвращает индекс первой строки кода внутри блока, которую надо обработать, 
    ''' а также индекс последней строки кода внутри блока, которую надо обработать.
    ''' </summary>
    ''' <param name="masCode">массив со всем выполняемым кодом</param>
    ''' <param name="queryToCompare">условие для проверки (уже вычисленное)</param>
    ''' <param name="startPos">в начале - индекс строки с обрабатываемым блоком Select Case. Затем передает индекс первой строки кода внутри блока, которую надо обработать</param>
    ''' <param name="finalPos">индекс последней строки кода внутри блока, которую надо обработать</param>
    ''' <param name="endSelectPos">индекс строки с End Select данного блока</param>
    ''' <param name="arrParams">параметры, с которыми запущен код</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockSwitch(ByRef masCode() As DictionaryEntry, ByVal queryToCompare As String, ByRef startPos As Integer, _
                                 ByRef finalPos As Integer, ByRef endSelectPos As Integer, ByRef arrParams() As String) As String
        'вычисляем условие для проверки
        If Double.TryParse(queryToCompare, NumberStyles.Any, provider_points, Nothing) = False Then
            queryToCompare = GetValue(queryToCompare, arrParams)
            If queryToCompare = "#Error" Then Return "#Error"
        End If
        finalPos = -1 'сбрасываем значение конца исполняемого блока

        Dim SwitchBalance As Integer = 1 'Баланс между Select Case / End Select (т. е. сколько открыто внутренних блоков Switch)
        Dim curLine As Integer = startPos + 1 'номер строки, с которой работаем
        Dim curString As String = "" 'текст (код) строки
        Dim saveStartPos As String = startPos 'переменная для хранения начального значения стартовой позиции (положение самого блока Select Case)
        Dim arrayResult() As String = {} 'для получения результатов финкции GetFunctionParams
        Dim blnStartsFromOperator As Boolean 'начинается ли условие в Case с оператора (>, <= и т. д.)
        Dim posOfOperatorTo As Integer 'Позиция оператора "То" в условии Case 

        'Выполняем цикл, пока баланс между Select Case и End Select не станет равен 0 (т. е., пока не найдем наш End Select)
        Do Until SwitchBalance = 0
            'ищем первый оператор
            curLine = Array.FindIndex(Of DictionaryEntry)(masCode, curLine, AddressOf BlockSwitchSearch)
            If curLine = -1 Then
                'Не найден ни один, а баланс еще не 0 - значит ошибка
                LAST_ERROR = "В блоке Select Case (Switch) не найден End Select (End Switch)."
                Return "#Error"
            End If

            curString = masCode(curLine).Value 'текущая строка
            If curString = "End Select" OrElse curString = "End Switch" Then 'найденный оператор - End Select
                SwitchBalance -= 1 'меняем баланс
                If SwitchBalance = 0 AndAlso finalPos = -1 Then finalPos = curLine - 1 'получили индекс последней строки блока на исполнение
            ElseIf curString.StartsWith("Case ") Then 'найденный оператор - Case 
                If SwitchBalance = 1 Then 'работаем только если мы внутри нашего блока (а не вложенного)
                    If startPos <> saveStartPos Then
                        'если начало блока установлено, а конец еще нет - устанавливаем конец на строку выше Case
                        If finalPos = -1 Then finalPos = curLine - 1
                    Else
                        'начало блока еще не установлено
                        If curString = "Case Else" Then
                            startPos = curLine + 1 'устанавливаем начало блока сразу вслед за Else
                        Else
                            'оператор Case ...
                            curString = curString.Substring(5) 'получаем условие после Case
                            'Должно работать Case 4 ... Case <> 4 ... Case 3 To 5, 6, 7, > 10 And x != C[1, 2].Max(5, 67)
                            'используем функцию GetFunctionParams для разбиения условий, разделенных скобками, в массив arrayResult
                            GetFunctionParams(curString, arrayResult, -1, Nothing, curString, Nothing, Nothing, False)
                            For i As Integer = 0 To arrayResult.GetUpperBound(0)
                                curString = arrayResult(i) 'текущее условие
                                'начинается с оператора сравнения?
                                blnStartsFromOperator = False
                                If curString.StartsWith("> ") OrElse curString.StartsWith("< ") OrElse curString.StartsWith("= ") _
                                    OrElse curString.StartsWith(">= ") OrElse curString.StartsWith("<= ") OrElse curString.StartsWith("<> ") _
                                    OrElse curString.StartsWith("!= ") Then blnStartsFromOperator = True
                                'ищем положение оператора То (5 То 10)
                                If curString.EndsWith("'") = False Then 'в строках его быть не должно
                                    posOfOperatorTo = curString.IndexOf(" To ")
                                End If
                                'получаем результат условия Case
                                If blnStartsFromOperator And posOfOperatorTo = -1 Then
                                    'начинается с оператора сравнения, То нет (напр., != 10)
                                    curString = GetValue(queryToCompare + " " + curString, arrParams)
                                ElseIf posOfOperatorTo = -1 Then
                                    'оператора сравнения нет, То тоже нет (напр., 10)
                                    curString = GetValue(queryToCompare + " = " + curString, arrParams)
                                ElseIf blnStartsFromOperator Then
                                    'начинается с оператора, То есть (напр., != 10 То 22) - ошибка
                                    LAST_ERROR = "Ошибка в условии Case. " + queryToCompare + " " + curString + " вычислить невозможно."
                                    Return "#Error"
                                Else
                                    'оператора нет, "То" есть (напр., 10 То 20)
                                    curString = GetValue(queryToCompare + " >= " + curString.Substring(0, posOfOperatorTo) + " And " + queryToCompare + " <= " + curString.Substring(posOfOperatorTo + 4), arrParams)
                                End If
                                If curString = "#Error" Then
                                    LAST_ERROR = "Ошибка в условии Case. " + queryToCompare + " = " + curString + " вычислить невозможно." + vbNewLine + LAST_ERROR
                                    Return "#Error"
                                End If
                                If curString <> "False" Then
                                    'Если хоть в одном из условий True - все условие соблюдено. Устанавливаем начало блока вслед за Case ...
                                    startPos = curLine + 1
                                    curLine += 1 'выходим из цикла For и продолжаем Do Until с новой строчки
                                    Continue Do
                                End If
                            Next
                        End If
                    End If
                End If
            Else 'найденный оператор - еще один внутренний Select Case
                SwitchBalance += 1 'меняем баланс
            End If
            curLine += 1 'будем искать дальше с новой строки
        Loop

        If startPos = saveStartPos Then finalPos = -1 'если начало блока так и найдено - сбрасываем конец блока (нет кода для исполнения)
        If finalPos < startPos Then finalPos = -1 'на случай, если между Case ... и End Select нет ни одной исполняемой строки - сбрасываем конец блока (нет кода для исполнения)
        endSelectPos = curLine - 1 'сохраняем позицию конечного End Select
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает цикл Do While и возвращает индекс последней строки кода внутри цикла
    ''' </summary>
    ''' <param name="exCode">код</param>
    ''' <param name="startLine">индекс строки с обрабатываемым блоком Do While</param>
    ''' <param name="finalLine">индекс последней строки кода внутри цикла</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockDoWhile(ByRef exCode As List(Of ExecuteDataType), ByRef startLine As Integer, ByRef finalLine As Integer) As String
        finalLine = -1
        Dim whileBalance As Integer = 1 'Баланс между Do While / Loop (т. е. сколько открыто внутренних блоков Do While).
        'Исходно мы внутри блока
        Dim curLine As Integer = startLine + 1 'Текущая линия, которую проверяем. В начале - первая линия после Do While
        Dim curCode() As CodeTextBox.EditWordType 'код текущей строки

        Do Until whileBalance = 0
            If curLine > exCode.Count - 1 Then
                LAST_ERROR = "Не найден конец блока Do While."
                Return "#Error"
            End If

            curCode = exCode(curLine).Code
            If IsNothing(curCode) OrElse curCode.Length = 0 Then
                'поустая строка
                curLine += 1
                Continue Do
            End If

            If curCode(0).Word.Trim = "Loop" Then
                If whileBalance = 1 Then
                    finalLine = curLine - 1
                    Return ""
                End If
                whileBalance -= 1
            ElseIf curCode(0).Word = "While " Then
                whileBalance += 1
            End If

            curLine += 1
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает цикл Do While и возвращает индекс последней строки кода внутри цикла
    ''' </summary>
    ''' <param name="masCode">массив со всем выполняемым кодом</param>
    ''' <param name="startPos">индекс строки с обрабатываемым блоком Do While</param>
    ''' <param name="finalPos">индекс последней строки кода внутри цикла</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockDoWhile(ByRef masCode() As DictionaryEntry, ByVal startPos As Integer, ByRef finalPos As Integer) As String
        finalPos = -1 'сбрасываем значение конца исполняемого блока

        Dim doWhileBalance As Integer = 1 'Баланс между Do While / Loop (т. е. сколько открыто внутренних циклов Do While)
        Dim curLine As Integer = startPos + 1 'номер строки, с которой работаем
        Dim curString As String = "" 'текст (код) строки
        Dim saveStartPos As String = startPos 'переменная для хранения начального значения стартовой позиции (положение самого цикла For ... Next)

        startPos += 1 'начало исполняемого блока сразу за Do While
        'Выполняем цикл, пока баланс между Do While и Loop не станет равен 0 (т. е., пока не найдем наш Loop)
        Do Until doWhileBalance = 0
            'ищем первый оператор
            curLine = Array.FindIndex(Of DictionaryEntry)(masCode, curLine, AddressOf BlockDoWhileSearch)
            If curLine = -1 Then
                'Не найден ни один, а баланс еще не 0 - значит ошибка
                LAST_ERROR = "Не найден конец цикла в блоке Do While."
                Return "#Error"
            End If

            If masCode(curLine).Value = "Loop" Then 'найденный оператор - Loop
                doWhileBalance -= 1 'меняем баланс
                If doWhileBalance = 0 Then
                    finalPos = curLine - 1
                    Return ""
                End If
            Else 'найденный оператор - еще один внутренний Do While
                doWhileBalance += 1 'меняем баланс
            End If
            curLine += 1 'будем искать дальше с новой строки
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает блок Function и возвращает индекс последней строки кода внутри блока
    ''' </summary>
    ''' <param name="exCode">код</param>
    ''' <param name="startLine">индекс строки с обрабатываемым блоком Function</param>
    ''' <param name="finalLine">индекс последней строки кода внутри блока</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockFunction(ByRef exCode As List(Of ExecuteDataType), ByVal startLine As Integer, ByRef finalLine As Integer) As String
        finalLine = -1
        Dim funcBalance As Integer = 1 'Баланс между Function / End Function (т. е. сколько открыто внутренних блоков Function).
        'Исходно мы внутри блока
        Dim curLine As Integer = startLine + 1 'Текущая линия, которую проверяем. В начале - первая линия после Function
        Dim curCode() As CodeTextBox.EditWordType 'код текущей строки

        Do Until funcBalance = 0
            If curLine > exCode.Count - 1 Then
                LAST_ERROR = "Не найден конец блока Function."
                Return "#Error"
            End If

            curCode = exCode(curLine).Code
            If IsNothing(curCode) OrElse curCode.Length = 0 Then
                'поустая строка
                curLine += 1
                Continue Do
            End If

            If curCode.Length = 2 AndAlso curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_CYCLE_END AndAlso _
                curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION Then
                If funcBalance = 1 Then
                    finalLine = curLine - 1
                    Return ""
                End If
                funcBalance -= 1
            ElseIf curCode(0).Word = "Function " Then
                funcBalance += 1
            End If

            curLine += 1
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает блок Function и возвращает индекс последней строки кода внутри блока
    ''' </summary>
    ''' <param name="masCode">массив со всем выполняемым кодом</param>
    ''' <param name="startPos">индекс строки с обрабатываемым блоком Function</param>
    ''' <param name="finalPos">индекс последней строки кода внутри блока</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockFunction(ByRef masCode() As DictionaryEntry, ByVal startPos As Integer, ByRef finalPos As Integer) As String
        finalPos = -1 'сбрасываем значение конца исполняемого блока

        Dim functionBalance As Integer = 1 'Баланс между Function / End Function (т. е. сколько открыто внутренних блоков End Function)
        Dim curLine As Integer = startPos + 1 'номер строки, с которой работаем
        Dim curString As String = "" 'текст (код) строки
        Dim saveStartPos As String = startPos 'переменная для хранения начального значения стартовой позиции (положение самого блока End Function)

        startPos += 1 'начало исполняемого блока сразу за Function
        'Выполняем цикл, пока баланс между Function и End Function не станет равен 0 (т. е., пока не найдем наш End Function)
        Do Until functionBalance = 0
            'ищем первый оператор
            curLine = Array.FindIndex(Of DictionaryEntry)(masCode, curLine, AddressOf BlockFunctionSearch)
            If curLine = -1 Then
                'Не найден ни один, а баланс еще не 0 - значит ошибка
                LAST_ERROR = "Не найден оператор End Function."
                Return "#Error"
            End If

            If masCode(curLine).Value = "End Function" Then 'найденный оператор - End Function
                functionBalance -= 1 'меняем баланс
                If functionBalance = 0 Then
                    finalPos = curLine - 1
                    Return ""
                End If
            Else 'найденный оператор - еще один внутренний Function
                functionBalance += 1 'меняем баланс
            End If
            curLine += 1 'будем искать дальше с новой строки
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает блок Event и возвращает индекс последней строки кода внутри блока
    ''' </summary>
    ''' <param name="exCode">код</param>
    ''' <param name="startLine">индекс строки с обрабатываемым блоком Event</param>
    ''' <param name="finalLine">индекс последней строки кода внутри блока</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockEvent(ByRef exCode As List(Of ExecuteDataType), ByVal startLine As Integer, ByRef finalLine As Integer) As String
        finalLine = -1
        Dim eventBalance As Integer = 1 'Баланс между Event / End Event (т. е. сколько открыто внутренних блоков Event).
        'Исходно мы внутри блока
        Dim curLine As Integer = startLine + 1 'Текущая линия, которую проверяем. В начале - первая линия после Event
        Dim curCode() As CodeTextBox.EditWordType 'код текущей строки

        Do Until eventBalance = 0
            If curLine > exCode.Count - 1 Then
                LAST_ERROR = "Не найден конец блока Event."
                Return "#Error"
            End If

            curCode = exCode(curLine).Code
            If IsNothing(curCode) OrElse curCode.Length = 0 Then
                'поустая строка
                curLine += 1
                Continue Do
            End If

            If curCode.Length = 2 AndAlso curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_CYCLE_END AndAlso _
                curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT Then
                If eventBalance = 1 Then
                    finalLine = curLine - 1
                    Return ""
                End If
                eventBalance -= 1
            ElseIf curCode(0).Word = "Event " Then
                eventBalance += 1
            End If

            curLine += 1
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает блок Event и возвращает индекс последней строки кода внутри блока
    ''' </summary>
    ''' <param name="masCode">массив со всем выполняемым кодом</param>
    ''' <param name="startPos">индекс строки с обрабатываемым блоком Event</param>
    ''' <param name="finalPos">индекс последней строки кода внутри блока</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockEvent(ByRef masCode() As DictionaryEntry, ByVal startPos As Integer, ByRef finalPos As Integer) As String
        finalPos = -1 'сбрасываем значение конца исполняемого блока

        Dim eventBalance As Integer = 1 'Баланс между Event / End Event (т. е. сколько открыто внутренних блоков Event)
        Dim curLine As Integer = startPos + 1 'номер строки, с которой работаем
        Dim curString As String = "" 'текст (код) строки
        Dim saveStartPos As String = startPos 'переменная для хранения начального значения стартовой позиции (положение самого блока Event)

        startPos += 1 'начало исполняемого блока сразу за Event
        'Выполняем цикл, пока баланс между Event и End Event не станет равен 0 (т. е., пока не найдем наш End Event)
        Do Until eventBalance = 0
            'ищем первый оператор
            curLine = Array.FindIndex(Of DictionaryEntry)(masCode, curLine, AddressOf BlockEventSearch)
            If curLine = -1 Then
                'Не найден ни один, а баланс еще не 0 - значит ошибка
                LAST_ERROR = "Не найден оператор End Event."
                Return "#Error"
            End If

            If masCode(curLine).Value = "End Event" Then 'найденный оператор - End Event
                eventBalance -= 1 'меняем баланс
                If eventBalance = 0 Then
                    finalPos = curLine - 1
                    Return ""
                End If
            Else 'найденный оператор - еще один внутренний Event
                eventBalance += 1 'меняем баланс
            End If
            curLine += 1 'будем искать дальше с новой строки
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает блок Wrap и возвращает индекс последней строки кода внутри цикла
    ''' </summary>
    ''' <param name="exCode">код</param>
    ''' <param name="startLine">индекс строки с обрабатываемым блоком Wrap</param>
    ''' <param name="finalLine">индекс последней строки кода внутри цикла</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockWrap(ByRef exCode As List(Of ExecuteDataType), ByVal startLine As Integer, ByRef finalLine As Integer) As String
        finalLine = -1
        Dim wrapBalance As Integer = 1 'Баланс между Wrap / End Wrap (т. е. сколько открыто внутренних блоков Wrap).
        'Исходно мы внутри блока
        Dim curLine As Integer = startLine + 1 'Текущая линия, которую проверяем. В начале - первая линия после Wrap
        Dim curCode() As CodeTextBox.EditWordType 'код текущей строки

        Do Until wrapBalance = 0
            If curLine > exCode.Count - 1 Then
                LAST_ERROR = "Не найден конец блока Wrap."
                Return "#Error"
            End If

            curCode = exCode(curLine).Code
            If IsNothing(curCode) OrElse curCode.Length = 0 Then
                'поустая строка
                curLine += 1
                Continue Do
            End If

            If curCode.Length = 2 AndAlso curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_CYCLE_END AndAlso _
                curCode(1).wordType = CodeTextBox.EditWordTypeEnum.W_WRAP Then
                If wrapBalance = 1 Then
                    finalLine = curLine - 1
                    Return ""
                End If
                wrapBalance -= 1
            ElseIf curCode(0).Word = "Wrap " Then
                wrapBalance += 1
            End If

            curLine += 1
        Loop
        Return ""
    End Function

    ''' <summary>
    ''' Функция обрабатывает блок Wrap и возвращает индекс последней строки кода внутри цикла
    ''' </summary>
    ''' <param name="masCode">массив со всем выполняемым кодом</param>
    ''' <param name="startPos">индекс строки с обрабатываемым блоком Wrap</param>
    ''' <param name="finalPos">индекс последней строки кода внутри цикла</param>
    ''' <returns>Пустую строку, в случае ошибки - #Error</returns>
    ''' <remarks></remarks>
    Protected Overridable Function BlockWrap(ByRef masCode() As DictionaryEntry, ByVal startPos As Integer, ByRef finalPos As Integer) As String
        finalPos = -1 'сбрасываем значение конца исполняемого блока

        Dim wrapBalance As Integer = 1 'Баланс между Wrap / End Wrap (т. е. сколько открыто внутренних циклов Wrap)
        Dim curLine As Integer = startPos + 1 'номер строки, с которой работаем
        Dim curString As String = "" 'текст (код) строки
        Dim saveStartPos As String = startPos 'переменная для хранения начального значения стартовой позиции (положение самого цикла For ... Next)

        startPos += 1 'начало исполняемого блока сразу за Wrap
        'Выполняем цикл, пока баланс между Wrap и End Wrap не станет равен 0 (т. е., пока не найдем наш End Wrap)
        Do Until wrapBalance = 0
            'ищем первый оператор
            curLine = Array.FindIndex(Of DictionaryEntry)(masCode, curLine, AddressOf BlockDoWhileSearch)
            If curLine = -1 Then
                'Не найден ни один, а баланс еще не 0 - значит ошибка
                LAST_ERROR = "Не найден конец цикла в блоке Wrap."
                Return "#Error"
            End If

            If masCode(curLine).Value = "Loop" Then 'найденный оператор - Loop
                wrapBalance -= 1 'меняем баланс
                If wrapBalance = 0 Then
                    finalPos = curLine - 1
                    Return ""
                End If
            Else 'найденный оператор - еще один внутренний Do While
                wrapBalance += 1 'меняем баланс
            End If
            curLine += 1 'будем искать дальше с новой строки
        Loop
        Return ""
    End Function


#End Region

#Region "Block's search functions"
    ''' <summary>
    ''' Функция поиска для блока HTML. Ищет End HTML на этапе подготовки блока
    ''' </summary>
    ''' <param name="s">строка для поиска</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PrepareTextSearch(ByVal s As String) As Boolean
        'Функция поиска для блока Wrap. Ищет End Wrap на этапе подготовки блока
        s = s.Trim
        If s = "End HTML" OrElse s.StartsWith("End HTML //") Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Функция поиска для блока Wrap. Ищет End Wrap на этапе подготовки блока
    ''' </summary>
    ''' <param name="s"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PrepareWrapSearch(ByVal s As String) As Boolean
        'Функция поиска для блока Wrap. Ищет End Wrap на этапе подготовки блока
        s = s.Trim
        If s = "End Wrap" OrElse s.StartsWith("End Wrap //") Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Функция поиска для BlockIF. Ищет операторы блока If ... Then
    ''' </summary>
    ''' <param name="s">ячейка кода для поиска</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function BlockIfSearch(ByVal s As DictionaryEntry) As Boolean
        If s.Value = "End If" OrElse s.Value.ToString.StartsWith("Else") OrElse (s.Value.ToString.StartsWith("If ") AndAlso s.Value.ToString.EndsWith(" Then")) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Функция поиска для BlockFor. Ищет операторы блока For ... Next
    ''' </summary>
    ''' <param name="s">ячейка кода для поиска</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function BlockForSearch(ByVal s As DictionaryEntry) As Boolean
        If s.Value.ToString.StartsWith("For ") OrElse s.Value = "Next" OrElse s.Value.ToString.StartsWith("Next ") Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Функция поиска для BlockSwitch. Ищет операторы блока Select Case / Switch
    ''' </summary>
    ''' <param name="s">ячейка кода для поиска</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function BlockSwitchSearch(ByVal s As DictionaryEntry) As Boolean
        If s.Value = "End Select" OrElse s.Value = "End Switch" OrElse s.Value.ToString.StartsWith("Case ") OrElse s.Value.ToString.StartsWith("Select Case ") OrElse s.Value.ToString.StartsWith("Switch ") Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Функция поиска для BlockSwitch. Ищет операторы блока Select Case / Switch
    ''' </summary>
    ''' <param name="s">ячейка кода для поиска</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function BlockDoWhileSearch(ByVal s As DictionaryEntry) As Boolean
        If s.Value = "Loop" OrElse s.Value.ToString.StartsWith("Do While") Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Функция поиска для BlockFunction. Ищет операторы блока Function
    ''' </summary>
    ''' <param name="s">ячейка кода для поиска</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function BlockFunctionSearch(ByVal s As DictionaryEntry) As Boolean
        If s.Value = "End Function" OrElse s.Value.ToString.StartsWith("Function ") Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Функция поиска для BlockEvent. Ищет операторы блока Event
    ''' </summary>
    ''' <param name="s">ячейка кода для поиска</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function BlockEventSearch(ByVal s As DictionaryEntry) As Boolean
        If s.Value = "End Event" OrElse s.Value.ToString.StartsWith("Event ") Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region

    ''' <summary>
    ''' Определяет, содержит ли свойство скрипт matewEngine (кодированный в формате xml)
    ''' </summary>
    ''' <param name="propContent">строка с содержимым свойства</param>
    Public Function IsPropertyContainsCode(ByVal propContent As String) As ContainsCodeEnum
        If IsNothing(propContent) Then Return ContainsCodeEnum.NOT_CODE
        If propContent.StartsWith("<ArrayOfCodeDataType") Then
            Dim wordTypePos As Integer = propContent.IndexOf("<wordType>", 154)
            If wordTypePos > -1 AndAlso propContent.Length > wordTypePos + 12 Then
                If propContent.Substring(wordTypePos + 10, 11) = "W_HTML_DATA" Then
                    Return ContainsCodeEnum.LONG_TEXT
                End If
            End If
            Return ContainsCodeEnum.CODE
            'My.Computer.FileSystem.WriteAllText("D:\Test.txt", propContent, False)
        Else
            If propContent.Length > 2 Then
                Dim chr As Char = propContent.Chars(1)
                If chr = "?"c Then Return ContainsCodeEnum.EXECUTABLE_STRING
            End If
            Return ContainsCodeEnum.NOT_CODE
        End If
    End Function

    ''' <summary>
    ''' Сериализует объект codeBox.CodeData в xml и возвращает результат в виде xml-текста
    ''' </summary>
    Public Function SerializeCodeData(ByRef CodeData() As CodeTextBox.CodeDataType) As String
        If IsNothing(CodeData) OrElse CodeData.Count = 0 Then Return ""
        If CodeData.Count = 1 Then
            If IsNothing(CodeData(0).Code) OrElse CodeData(0).Code.Length = 0 Then
                If IsNothing(CodeData(0).Comments) OrElse CodeData(0).Comments.Length = 0 Then
                    If IsNothing(CodeData(0).StartingSpaces) OrElse CodeData(0).StartingSpaces.Length = 0 Then
                        Return ""
                    End If
                End If
            End If
        End If
        Dim sb As New System.Text.StringBuilder

        Using xmlWriter As System.Xml.XmlWriter = System.Xml.XmlWriter.Create(sb, New System.Xml.XmlWriterSettings With {.OmitXmlDeclaration = True})
            Dim x As New System.Xml.Serialization.XmlSerializer(CodeData.GetType)
            CheckSymb01(CodeData)
            x.Serialize(xmlWriter, CodeData)
            Return sb.ToString
        End Using
    End Function


    ''' <summary>
    ''' Создает полный клон объекта с помощью xml-сериализации
    ''' </summary>
    Public Function CloneObjectFull(ByRef obj As Object) As Object
        If IsNothing(obj) Then Return Nothing
        Dim sb As New System.Text.StringBuilder

        Using xmlWriter As System.Xml.XmlWriter = System.Xml.XmlWriter.Create(sb, New System.Xml.XmlWriterSettings With {.OmitXmlDeclaration = True})
            Dim x As New System.Xml.Serialization.XmlSerializer(obj.GetType)
            x.Serialize(xmlWriter, obj)
        End Using

        Using xmlMemorySteam As New System.IO.MemoryStream
            Using xmlSteamWriter As New System.IO.StreamWriter(xmlMemorySteam)
                xmlSteamWriter.Write(sb.ToString)
                xmlSteamWriter.Flush()
            End Using

            Using xmlMemorySteam2 As New System.IO.MemoryStream(xmlMemorySteam.GetBuffer)
                Using xmlStreamReader As New System.IO.StreamReader(xmlMemorySteam2)
                    Dim x As New System.Xml.Serialization.XmlSerializer(obj.GetType)
                    Return x.Deserialize(xmlMemorySteam2)
                End Using
            End Using
        End Using

    End Function

    Private Sub CheckSymb01(ByRef CodeData() As CodeTextBox.CodeDataType)
        'удаляет символы с кодом 01 (если таковые имеются)
        For i As Integer = 0 To CodeData.Count - 1
            If IsNothing(CodeData(i).Code) = False Then
                For j As Integer = 0 To CodeData(i).Code.Count - 1
                    If CodeData(i).Code(j).Word.IndexOf(Chr(1)) > 0 Then
                        CodeData(i).Code(j).Word = CodeData(i).Code(j).Word.Replace(Chr(1), "")
                    End If
                Next
            End If
        Next
    End Sub

    ''' <summary>
    ''' Десериализует xml-строку и преобразует ее в массив кода в формате CodeDataType()
    ''' </summary>
    ''' <param name="strXML">xml-строка</param>
    Public Function DeserializeCodeData(ByRef strXML As String) As CodeTextBox.CodeDataType()
        If [String].IsNullOrEmpty(strXML) Then
            Dim dt() As CodeTextBox.CodeDataType = Nothing
            ReDim dt(0)
            Return dt
        End If

        Using xmlMemorySteam As New System.IO.MemoryStream
            Using xmlSteamWriter As New System.IO.StreamWriter(xmlMemorySteam)
                xmlSteamWriter.Write(strXML)
                xmlSteamWriter.Flush()
            End Using

            Using xmlMemorySteam2 As New System.IO.MemoryStream(xmlMemorySteam.GetBuffer)
                Using xmlStreamReader As New System.IO.StreamReader(xmlMemorySteam2)
                    Dim d() As CodeTextBox.CodeDataType
                    ReDim d(0)
                    d(0) = New CodeTextBox.CodeDataType
                    Dim x As New System.Xml.Serialization.XmlSerializer(d.GetType)
                    Return x.Deserialize(xmlMemorySteam2)
                End Using
            End Using

        End Using
    End Function

    ''' <summary>
    ''' Создает элемент второго уровня у выбранного класса
    ''' </summary>
    ''' <param name="mainClassId">Идентификатор класса</param>
    ''' <param name="newName">имя нового элемента</param>
    ''' <returns>идентификатор нового элемента</returns>
    ''' <remarks></remarks>
    Public Function CreateSecondLevelElement(ByVal mainClassId As Integer, ByVal newName As String) As Integer
        Dim newId As Integer = 0
        If IsNothing(mainClass(mainClassId).ChildProperties) = False AndAlso mainClass(mainClassId).ChildProperties.Count > 0 Then
            newId = mainClass(mainClassId).ChildProperties.Count
            For i As Integer = 0 To mainClass(mainClassId).ChildProperties.Count - 1
                If mainClass(mainClassId).ChildProperties(i)("Name").Value = newName Then
                    MsgBox("Элемент с таким именем уже существует!", vbExclamation)
                    Return -1
                End If
            Next
        End If
        Array.Resize(mainClass(mainClassId).ChildProperties, newId + 1)
        mainClass(mainClassId).ChildProperties(newId) = New SortedList(Of String, ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
        For Each pr As KeyValuePair(Of String, PropertiesInfoType) In mainClass(mainClassId).Properties
            Dim chPr As New ChildPropertiesInfoType
            chPr.Value = pr.Value.Value
            chPr.Hidden = pr.Value.Hidden
            mainClass(mainClassId).ChildProperties(newId).Add(pr.Key, chPr)
        Next
        Dim namePr As ChildPropertiesInfoType = mainClass(mainClassId).ChildProperties(newId)("Name")
        namePr.Value = "'" + newName + "'"
        Return newId
    End Function

    ''' <summary>
    ''' Устанавливает значение свойства элемента второго порядка
    ''' </summary>
    ''' <param name="classId">идентификатор класса</param>
    ''' <param name="childId">идентификатор элемента второго порядка</param>
    ''' <param name="propName">имя свойства</param>
    ''' <param name="propValue">новое значение свойства</param>
    Public Sub SetSecondLevelProperty(ByVal classId As Integer, ByVal childId As Integer, ByVal propName As String, ByVal propValue As String)
        If classId = -1 Then
            MsgBox("Класс свойства указан неправильно.", vbExclamation)
            Return
        End If
        If IsNothing(mainClass(classId).ChildProperties) OrElse childId > mainClass(classId).ChildProperties.Count - 1 OrElse childId < 0 Then
            MsgBox("Идентификатор элемента второго уровня указан неправильно.", vbExclamation)
            Return
        End If
        If mainClass(classId).ChildProperties(childId).ContainsKey(propName) = False Then
            MsgBox("Свойство " + propName + " не найдено.", vbExclamation)
            Return
        End If
        Dim pr As ChildPropertiesInfoType = mainClass(classId).ChildProperties(childId)(propName)
        pr.Value = propValue
    End Sub

    ''' <summary>
    ''' Переобразует структуру кода codeBox.CodeData в html-текст
    ''' </summary>
    ''' <param name="cd">codeBox.CodeData для обработки</param>
    ''' <param name="isScript">Создавать структуру для скрипта или для длинного текста</param>
    Public Function ConvertCodeDataToHTML(ByRef cd() As CodeTextBox.CodeDataType, Optional isScript As Boolean = True) As String
        If IsNothing(cd) OrElse cd.Count = 0 Then Return ""

        Dim st As CodeTextBox.StylePresetType 'для получения стиля и цвета слова (в зависимости от его значения)
        Dim blnDrawWord As Boolean 'надо ли менять цвет /стиль текущего слова
        Dim default_st As CodeTextBox.StylePresetType = questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_NOTHING) 'стиль по-умолчанию
        Dim comm_st As CodeTextBox.StylePresetType = questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_COMMENTS) '<!-- -->
        Dim strR, strB, strG, strIB1, strIB2 As String

        strR = Hex(comm_st.style_Color.R)
        If strR.Length = 1 Then strR = "0" + strR
        strG = Hex(comm_st.style_Color.G)
        If strG.Length = 1 Then strG = "0" + strG
        strB = Hex(comm_st.style_Color.B)
        If strB.Length = 1 Then strB = "0" + strB
        Select Case comm_st.font_style
            Case FontStyle.Bold
                strIB1 = "<b>"
                strIB2 = "</b>"
            Case FontStyle.Italic
                strIB1 = "<i>"
                strIB2 = "</i>"
            Case FontStyle.Bold + FontStyle.Bold
                strIB1 = "<b><i>"
                strIB2 = "</i></b>"
            Case Else
                strIB1 = ""
                strIB2 = ""
        End Select
        Dim strCommOpen As String = "<font color='#" + strR + strG + strB + "'>" + strIB1 + "&lt!--"
        Dim strCommClose As String = "--&gt" + strIB2 + "</font>"
        Dim selector_st As CodeTextBox.StylePresetType = questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_RETURN) '#X:

        Dim cur_st As CodeTextBox.StylePresetType = default_st 'стиль фрагмента, который надо будет менять
        Dim isHTMLdata As Boolean = False
        Dim res As New System.Text.StringBuilder
        Dim prevColor As String = ""
        'If showAsCode Then res.Append("<code>")
        For lineId As Integer = 0 To cd.Count - 1
            Dim c As CodeTextBox.CodeDataType = cd(lineId)
            If IsNothing(c) Then Continue For
            If [String].IsNullOrEmpty(c.StartingSpaces) = False Then
                If isScript Then
                    res.Append(c.StartingSpaces.Replace(" ", "&nbsp;"))
                Else
                    res.Append(c.StartingSpaces)
                End If
            End If
            Dim arrWords() As CodeTextBox.EditWordType = c.Code
            If IsNothing(c.Code) = False AndAlso c.Code.Count > 0 Then
                For i As Integer = 0 To arrWords.GetUpperBound(0) 'перебираем все слова в цикле
                    'получаем стиль текущего слова
                    If arrWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_CYCLE_END OrElse arrWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER _
                        OrElse arrWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING Then
                        'если текущее слово - "End" # или $ - берем стиль следующего слова 
                        If i < arrWords.GetUpperBound(0) Then
                            If questEnvironment.codeBoxShadowed.codeBox.styleHash.TryGetValue(arrWords(i + 1).wordType, st) = False Then
                                st = default_st
                            End If
                        Else
                            st = questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_BLOCK_IF)
                        End If
                        blnDrawWord = True
                    ElseIf arrWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_POINT Then
                        'если текущее слово - точка то это или Class. или [...].Function / Property
                        If i > 0 AndAlso arrWords(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS Then
                            'стиль точки = стилю класса
                            st = IIf(arrWords(i - 1).classId = -1, questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_VARIABLE), questEnvironment.codeBoxShadowed.codeBox.styleHash(arrWords(i - 1).wordType))
                        Else
                            'стиль точки = стилю свойства / функции
                            If i < arrWords.GetUpperBound(0) Then
                                If questEnvironment.codeBoxShadowed.codeBox.styleHash.TryGetValue(arrWords(i + 1).wordType, st) = False Then
                                    st = questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                                End If
                            Else
                                st = questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                            End If
                        End If
                        blnDrawWord = True
                    ElseIf arrWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS AndAlso arrWords(i).classId = -1 Then
                        'V / Var
                        st = questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_VARIABLE)
                        blnDrawWord = True
                    ElseIf questEnvironment.codeBoxShadowed.codeBox.styleHash.TryGetValue(arrWords(i).wordType, st) = True Then 'получаем стиль текущего слова
                        blnDrawWord = True
                    Else
                        'такой тип не найден в таблице стилей ExEnvironment.codeBox.styleHash. Закрашивание слова будет по-умолчанию
                        blnDrawWord = False
                        st = default_st
                    End If
                    If isHTMLdata = False Then
                        If arrWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA Then isHTMLdata = True
                    End If

                    Dim sWord As String = arrWords(i).Word
                    If (arrWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA OrElse arrWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING) AndAlso _
                        isScript = True Then
                        Dim stTags As CodeTextBox.StylePresetType = questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_HTML_TAG) 'стиль тэгов
                        strR = Hex(stTags.style_Color.R)
                        If strR.Length = 1 Then strR = "0" + strR
                        strG = Hex(stTags.style_Color.G)
                        If strG.Length = 1 Then strG = "0" + strG
                        strB = Hex(stTags.style_Color.B)
                        If strB.Length = 1 Then strB = "0" + strB
                        Select Case stTags.font_style
                            Case FontStyle.Bold
                                strIB1 = "<b>"
                                strIB2 = "</b>"
                            Case FontStyle.Italic
                                strIB1 = "<i>"
                                strIB2 = "</i>"
                            Case FontStyle.Bold + FontStyle.Bold
                                strIB1 = "<b><i>"
                                strIB2 = "</i></b>"
                            Case Else
                                strIB1 = ""
                                strIB2 = ""
                        End Select
                        Dim repStr1 As String = "<font color='#" + strR + strG + strB + "'>" + strIB1 + "&lt"
                        sWord = sWord.Replace("<", repStr1)
                        sWord = sWord.Replace(repStr1 + "!--", strCommOpen)
                        Dim repStr2 As String = "&gt" + strIB2 + "</font>"
                        sWord = sWord.Replace(">", repStr2)
                        sWord = sWord.Replace("--" + repStr2, strCommClose)
                    End If

                    strR = Hex(st.style_Color.R)
                    If strR.Length = 1 Then strR = "0" + strR
                    strG = Hex(st.style_Color.G)
                    If strG.Length = 1 Then strG = "0" + strG
                    strB = Hex(st.style_Color.B)
                    If strB.Length = 1 Then strB = "0" + strB
                    Select Case st.font_style
                        Case FontStyle.Bold
                            strIB1 = "<b>"
                            strIB2 = "</b>"
                        Case FontStyle.Italic
                            strIB1 = "<i>"
                            strIB2 = "</i>"
                        Case FontStyle.Bold + FontStyle.Bold
                            strIB1 = "<b><i>"
                            strIB2 = "</i></b>"
                        Case Else
                            strIB1 = ""
                            strIB2 = ""
                    End Select


                    If sWord.StartsWith("#") Then
                        'закрашиваем селекторы
                        Dim sLen As Integer = IsItSelector(sWord, 0)
                        Dim selB1 As String, selB2 As String, selR As String, selG As String, selB As String
                        If sLen > 0 Then
                            selR = Hex(selector_st.style_Color.R)
                            If selR.Length = 1 Then selR = "0" + selR
                            selG = Hex(selector_st.style_Color.G)
                            If selG.Length = 1 Then selG = "0" + selG
                            selB = Hex(selector_st.style_Color.B)
                            If selB.Length = 1 Then selB = "0" + selB
                            Select Case selector_st.font_style
                                Case FontStyle.Bold
                                    selB1 = "<b>"
                                    selB2 = "</b>"
                                Case FontStyle.Italic
                                    selB1 = "<i>"
                                    selB2 = "</i>"
                                Case FontStyle.Bold + FontStyle.Bold
                                    selB1 = "<b><i>"
                                    selB2 = "</i></b>"
                                Case Else
                                    selB1 = ""
                                    selB2 = ""
                            End Select
                            sWord = "<font color='#" + selR + selG + selB + "'>" + selB1 + sWord.Substring(0, sLen) + selB2 + "</font>" + sWord.Substring(sLen)
                        End If
                    End If

                    Dim curColor As String = ""
                    'If isScript AndAlso i = 0 Then blnDrawWord = True
                    If isScript = False AndAlso blnDrawWord = False Then
                        If String.IsNullOrEmpty(prevColor) Then
                            res.Append(sWord)
                        Else
                            res.Append("</font>" + sWord)
                            prevColor = ""
                        End If
                    Else
                        curColor = strR + strG + strB
                        If curColor = prevColor AndAlso (isScript AndAlso i = 0) = False Then
                            res.Append(strIB1 + sWord + strIB2)
                        ElseIf prevColor.Length = 0 Then
                            res.Append("<font color='#" + curColor + "'>" + strIB1 + sWord + strIB2)
                        Else
                            res.Append("</font><font color='#" + curColor + "'>" + strIB1 + sWord + strIB2)
                        End If
                        prevColor = curColor
                    End If
                    'res.Append("<font color='#" + curColor + "'>" + strIB1 + sWord + strIB2 + "</font>")
                Next
                If prevColor.Length > 0 AndAlso (isScript OrElse blnDrawWord) Then res.Append("</font>")
            End If
            st = questEnvironment.codeBoxShadowed.codeBox.styleHash(CodeTextBox.EditWordTypeEnum.W_COMMENTS)
            strR = Hex(st.style_Color.R)
            If strR.Length = 1 Then strR = "0" + strR
            strG = Hex(st.style_Color.G)
            If strG.Length = 1 Then strG = "0" + strR
            strB = Hex(st.style_Color.B)
            If strB.Length = 1 Then strB = "0" + strB
            Select Case st.font_style
                Case FontStyle.Bold
                    strIB1 = "<b>"
                    strIB2 = "</b>"
                Case FontStyle.Italic
                    strIB1 = "<i>"
                    strIB2 = "</i>"
                Case FontStyle.Bold + FontStyle.Bold
                    strIB1 = "<b><i>"
                    strIB2 = "</i></b>"
                Case Else
                    strIB1 = ""
                    strIB2 = ""
            End Select
            If [String].IsNullOrEmpty(c.Comments) = False Then res.Append("<font color='#" + strR + strG + strB + "'>" + strIB1 + c.Comments + strIB2 + "</font>")
            If isScript AndAlso lineId < cd.Count - 1 Then res.Append("<br/>" & vbNewLine)
        Next

        'If showAsCode Then res.Append("</code>")

        Return res.ToString
    End Function

    ' ''' <summary>
    ' ''' Возвращает количество селекторов. Делается допуще
    ' ''' </summary>
    ' ''' <param name="cd"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function GetSelectorsCount(ByRef cd() As CodeTextBox.CodeDataType) As Integer
    '    If IsNothing(cd) OrElse cd.Count = 0 Then Return 0
    '    If IsNothing(cd(0).Code) OrElse cd(0).Code.Count = 0 OrElse cd(0).Code(0).wordType <> CodeTextBox.EditWordTypeEnum.W_HTML_DATA OrElse IsItSelector(cd(0).Code(0).Word.Trim, 0) = 0 Then Return 1
    '    For lineId As Integer = cd.Count - 1 To 0 Step -1
    '        If IsNothing(cd(lineId).Code) OrElse cd(lineId).Code.Count = 0 Then Continue For
    '        Dim w As CodeTextBox.EditWordType = cd(lineId).Code(0)
    '        If w.wordType <> CodeTextBox.EditWordTypeEnum.W_HTML_DATA Then Continue For
    '        Dim strWord As String = w.Word.Trim
    '        If strWord.Chars(0) <> "#"c Then Continue For
    '        Dim sLen As Integer = IsItSelector(strWord, 0)
    '        If sLen = 0 Then Continue For
    '        Dim num As Integer = CInt(strWord.Substring(1, sLen - 2))
    '        If num < 1 Then Return 1
    '        Return num
    '    Next lineId

    '    Return 1
    'End Function

    ''' <summary>
    ''' Возвращает фрагмент кода длиного текста по указанному селектору
    ''' </summary>
    ''' <param name="cd">полный код</param>
    ''' <param name="selectorNumber">номер селектора, начиная с 1</param>
    Public Function GetLongTextBySelector(ByRef cd() As CodeTextBox.CodeDataType, ByVal selectorNumber As Integer) As CodeTextBox.CodeDataType()
        If IsNothing(cd) OrElse cd.Count = 0 Then Return cd
        If IsNothing(cd(0).Code) OrElse cd(0).Code.Count = 0 OrElse cd(0).Code(0).wordType <> CodeTextBox.EditWordTypeEnum.W_HTML_DATA OrElse IsItSelector(cd(0).Code(0).Word.Trim, 0) = 0 Then Return cd

        Dim selStart As Integer = -1
        Dim selLength As Integer = 0
        Dim strWord As String
        Dim strSelector As String = "#" + selectorNumber.ToString + ":"
        For lineId As Integer = 0 To cd.Count - 1
            If IsNothing(cd(lineId).Code) OrElse cd(lineId).Code.Count = 0 Then Continue For
            If cd(lineId).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA Then
                strWord = cd(lineId).Code(0).Word.TrimStart
                If strWord.StartsWith(strSelector) Then
                    selStart = lineId
                ElseIf selStart > -1 AndAlso strWord.Chars(0) = "#"c AndAlso IsItSelector(strWord, 0) > 0 Then
                    selLength = lineId - selStart
                    Exit For
                End If
            End If
        Next lineId

        Dim cdCopy() As CodeTextBox.CodeDataType
        If selStart = -1 Then
            cdCopy = CopyCodeDataArray(cd)
            strWord = cd(0).Code(0).Word.TrimStart
            cdCopy(0).Code(0).Word = strWord.Substring(0, strWord.Length - 3).TrimStart  'минус #1:
            Return cdCopy
        End If

        If selLength = 0 Then selLength = cd.Count - selStart + 1
        ReDim cdCopy(selLength - 1)
        Array.Copy(cd, selStart, cdCopy, 0, selLength)
        strWord = cdCopy(0).Code(0).Word.TrimStart
        cdCopy(0).Code(0).Word = strWord.Substring(0, strWord.Length - strSelector.Length).TrimStart  'минус #XX:
        Return cdCopy
    End Function

    ''' <summary>
    ''' Возвращает сортированный список, ключ которого - номер селектора, значение - код
    ''' </summary>
    ''' <param name="cd">полный код в формате Long_Text для извлечения селекторов</param>
    ''' <remarks>Пустой список или список с селекторами</remarks>
    Public Function SplitLongTextBySelectors(ByRef cd() As CodeTextBox.CodeDataType) As SortedList(Of Integer, CodeTextBox.CodeDataType())
        Dim lst As New SortedList(Of Integer, CodeTextBox.CodeDataType())
        If IsNothing(cd) OrElse cd.Count = 0 Then Return lst
        If IsNothing(cd(0).Code) OrElse cd(0).Code.Count = 0 OrElse cd(0).Code(0).wordType <> CodeTextBox.EditWordTypeEnum.W_HTML_DATA OrElse IsItSelector(cd(0).Code(0).Word.TrimStart, 0) = 0 Then
            'код не начинается на селектор #1:. Возвращаем список с единственным селектором
            lst.Add(1, cd)
            Return lst
        End If
        'Dim firstSelectorLength As Integer = 0 'длина строки первого селектора (например, #1: = 3)

        Dim selLength As Integer 'кол-во линий селектора
        'устанавливаем исходные значения
        Dim selStart As Integer = 0 'начальная первая линия селектора
        Dim strSelector As String = cd(0).Code(0).Word.Substring(0, IsItSelector(cd(0).Code(0).Word.TrimStart, 0)) ' "#1:" 'первый найденный селектор
        For lineId As Integer = 1 To cd.Count - 1
            'перебираем все линии в поиске новых селекторов
            If IsNothing(cd(lineId).Code) OrElse cd(lineId).Code.Count = 0 Then Continue For
            If cd(lineId).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA Then
                'селектор может быть только в начале строки. Если он не в первом слове, то его в строке нет
                Dim strWord As String = cd(lineId).Code(0).Word.TrimStart
                If String.IsNullOrEmpty(strWord) = False AndAlso strWord.Chars(0) = "#"c Then
                    Dim selectorLength As Integer = IsItSelector(strWord, 0) 'длина селектора #X: или 0, если это не селектор
                    If selectorLength > 0 Then
                        selLength = lineId - selStart 'количество линий селектора
                        'Копируем часть кода, относящуюся к данному селектору
                        Dim cdCopy() As CodeTextBox.CodeDataType
                        ReDim cdCopy(selLength - 1)
                        Array.Copy(cd, selStart, cdCopy, 0, selLength)
                        'Убираем сам селектор #X: из кода и добавляем в lst
                        strWord = cdCopy(0).Code(0).Word.TrimStart
                        cdCopy(0).Code(0).Word = strWord.Substring(strSelector.Length, strWord.Length - selectorLength).TrimStart
                        lst.Add(Val(strSelector.Substring(1)), cdCopy)
                        'сохраняем новый селектор и его начало
                        strSelector = cd(lineId).Code(0).Word.TrimStart.Substring(0, selectorLength)
                        selStart = lineId
                    End If
                End If
            End If
        Next lineId
        'Остался последний селектор
        selLength = cd.Count - selStart 'количество линий селектора
        'Копируем часть кода, относящуюся к данному селектору
        Dim cdCopyFinal() As CodeTextBox.CodeDataType
        ReDim cdCopyFinal(selLength - 1)
        Array.Copy(cd, selStart, cdCopyFinal, 0, selLength)
        'Убираем сам селектор #X: из кода и добавляем в lst
        Dim strWordFinal As String = cdCopyFinal(0).Code(0).Word.TrimStart
        cdCopyFinal(0).Code(0).Word = strWordFinal.Substring(strSelector.Length, strWordFinal.Length - strSelector.Length).TrimStart
        Dim selNumber As Integer = Val(strSelector.Substring(1))
        If lst.ContainsKey(selNumber) = False Then lst.Add(selNumber, cdCopyFinal)

        Return lst
    End Function

    ''' <summary>
    ''' Определяет находится ли на текущей позиции разделитель описаний #XX:
    ''' </summary>
    ''' <param name="txt">текст</param>
    ''' <param name="curPos">Позиция найденного символа #</param>
    ''' <returns>Длину разделителя #XXX...: или 0, если это что-то другое</returns>
    Public Function IsItSelector(ByVal txt As String, ByVal curPos As Integer) As Integer
        Dim startPos As Integer = curPos 'положение #
        Dim ch As Char
        If curPos > 0 Then
            ch = txt.Chars(curPos - 1)
            If ch <> vbLf AndAlso ch <> vbCr Then Return 0 'разделитель может быть только в начале строки
        End If

        Dim numberFound As Boolean = False 'есть ли между # и : число
        Do
            curPos += 1
            If curPos > txt.Length - 1 Then Return 0
            ch = txt.Chars(curPos) 'следующий символ
            If IsNumeric(ch) Then
                'найдено число после #
                numberFound = True
                Continue Do
            ElseIf ch = ":"c Then
                If Not numberFound Then Return 0 'найдено #: - неправильно, выход
                Return curPos - startPos + 1 'возвращаем длину разделителя
            Else
                Return 0 'найдено что-то другое
            End If
        Loop
    End Function

    ''' <summary>
    ''' Создает хэш сфойств и функций для быстрого поиска и распознавания в кодбоксах
    ''' </summary>
    Public Sub FillFuncAndPropHash()
        funcAndPropHash.Clear()

        Dim strClass As String

        For i As Integer = 0 To mScript.mainClass.GetUpperBound(0)
            strClass = mScript.mainClass(i).Names(0) + "_"
            If IsNothing(mScript.mainClass(i).Properties) = False Then
                For Each prop As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(i).Properties
                    funcAndPropHash.Add(strClass + prop.Key, New MatewScript.funcAndPropHashType With {.classId = i, .elementName = prop.Key, .elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_PROPERTY})
                Next
            End If
            If IsNothing(mScript.mainClass(i).Functions) = False Then
                For Each prop As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(i).Functions
                    If funcAndPropHash.ContainsKey(strClass + prop.Key) Then
                        MsgBox("Класс " + strClass + ". Найдены одноименные свойство и функция " + prop.Key + ". Загрузка продолжается, но работа до переименования может быть некорректной.", MsgBoxStyle.Exclamation)
                        Continue For
                    End If
                    funcAndPropHash.Add(strClass + prop.Key, New MatewScript.funcAndPropHashType With {.classId = i, .elementName = prop.Key, .elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION})
                Next
            End If
        Next
    End Sub

    ''' <summary>
    ''' Класс для работы с событиями движка MatewScript
    ''' </summary>
#Region "EventRouter"
    Public Class EventRouterClass
        ''' <summary>Сортированный список событий</summary>
        Public lstEvents As New SortedList(Of Integer, List(Of MatewScript.ExecuteDataType))
        ''' <summary>Сортированный список перменных, привязанных к событию. Ключ - eventId</summary>
        Public lstVariables As New SortedList(Of Integer, SortedList(Of String, cVariable.variableEditorInfoType))

        Private mScript As MatewScript
        ''' <summary>Индекс последнего сохраненного события (начинается с 1). При удалении события уменьшение индекса не происходит, что гарантирует уникальность каждого eventId</summary>
        Public lastEventId As Integer = 0

        Sub New()
            FillOperators()
        End Sub

        ''' <summary>Создает новый eventId. При удалении события уменьшение индекса не происходит, что гарантирует уникальность каждого eventId</summary>
        Private Function CreateNewId() As Integer
            lastEventId = lastEventId + 1
            Return lastEventId
        End Function

        Private Function CalculateNextId() As Integer
            Return lastEventId + 1
        End Function

        ''' <summary>Возвращает пустой ли список событий</summary>
        Public Function IsEventsListEmpty() As Boolean
            If IsNothing(lstEvents) OrElse lstEvents.Count = 0 Then Return True
            Return False
        End Function

        ''' <summary>Очищает список событий</summary>
        Public Sub Clear()
            lstEvents.Clear()
            lstVariables.Clear()
        End Sub

        ''' <summary>Возвращает структуру кода по eventId</summary>
        Public Function GetExDataByEventId(ByVal eventId As Integer) As List(Of MatewScript.ExecuteDataType)
            Dim res As List(Of MatewScript.ExecuteDataType) = Nothing
            lstEvents.TryGetValue(eventId, res)
            Return res
        End Function

        ''' <summary>
        ''' Определяет иммется ли событие с указанным eventId и, если оно имеется, то не пустое ли оно. Также возвращает структуру события
        ''' </summary>
        ''' <param name="eventId">Id события для получения</param>
        ''' <param name="exData">ссылка для возврата exData</param>
        Public Function IsExistsAndNotEmpty(ByVal eventId As Integer, ByRef exData As List(Of MatewScript.ExecuteDataType)) As Boolean
            If lstEvents.TryGetValue(eventId, exData) Then
                If IsNothing(exData) Then Return False
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Устанавливает обработчик свойства (кода или исполняемой строки) или заменяет его на новый, если таковой уже создан, а также устанавливает все необходимые параметры свойства в mainClass.
        ''' Если строка не является кодом или исполняемой, то просто запускает SetPropertyValue
        ''' </summary>
        ''' <param name="classId">Индекс класса - хозяина обработчика события</param>
        ''' <param name="propertyName">Название события-обработчика</param>
        ''' <param name="propertyValue">Значение свойства. Это может быть код, исполняемая или обычная строка</param>
        ''' <param name="funcParams">Параметры блока Event</param>
        ''' <returns>Пустую строку; в случае ошибки - #Error.</returns>
        Public Overridable Function SetPropertyWithEvent(ByVal classId As Integer, ByVal propertyName As String, ByVal propertyValue As String, _
                                                ByRef funcParams() As String, ByVal isContainsCode As ContainsCodeEnum, ByVal isUserFunction As Boolean, _
                                                Optional ByVal ignoreBattle As Boolean = False) As String
            Dim child2Id As Integer = -1, child3Id As Integer = -1

            If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then
                If GVARS.G_ISBATTLE AndAlso ignoreBattle = False AndAlso mScript.mainClass(classId).Names(0) = "H" Then
                    child2Id = mScript.Battle.GetFighterByName(funcParams(0))
                Else
                    child2Id = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                    If funcParams.Count > 1 AndAlso child2Id > -1 Then child3Id = GetThirdChildIdByName(funcParams(1), child2Id, mScript.mainClass(classId).ChildProperties)
                End If
            End If

            If isContainsCode = ContainsCodeEnum.CODE OrElse isContainsCode = ContainsCodeEnum.LONG_TEXT Then
                'в свойстве код/длинный текст
                Return SetPropertyWithEvent(classId, propertyName, mScript.DeserializeCodeData(propertyValue), child2Id, child3Id, propertyValue, False, isUserFunction, True, ignoreBattle)
            ElseIf isContainsCode = ContainsCodeEnum.NOT_CODE Then
                SetPropertyValue(classId, propertyName, propertyValue, child2Id, child3Id, ignoreBattle)
                Return ""
            End If

            'исполняемая строка
            'получаем codeBox.CodeData исполняемой строки
            Dim strToPrint As String = mScript.PrepareStringToPrint(propertyValue, Nothing, False) 'получаем строку без кавычек (код не исполняем)
            Dim dt() As CodeTextBox.CodeDataType
            With questEnvironment.codeBoxShadowed
                .codeBox.Text = strToPrint
                dt = CopyCodeDataArray(.codeBox.CodeData)
                propertyValue = WrapString(.Text)
                .codeBox.Clear()
            End With
            Return SetPropertyWithEvent(classId, propertyName, dt, child2Id, child3Id, propertyValue, True, isUserFunction, ignoreBattle)
        End Function

        ''' <summary>
        ''' Устанавливает обработчик свойства (кода или исполняемой строки) или заменяет его на новый, если таковой уже создан, а также устанавливает все необходимые параметры свойства в mainClass.
        ''' Если строка не является кодом или исполняемой, то просто запускает SetPropertyValue
        ''' </summary>
        ''' <param name="classId">Индекс класса - хозяина обработчика события</param>
        ''' <param name="propertyName">Название события-обработчика</param>
        ''' <param name="propertyValue">Значение свойства. Это может быть код, исполняемая или обычная строка</param>
        ''' <param name="child2Id">Id элемента 2 порядка, которому устанавливается обработчик, или -1, если устанавливаем глобальный обработчик</param>
        ''' <param name="child3Id">Id элемента 3 порядка, которому устанавливается обработчик, или -1, если устанавливаем обработчик элемента 2 порядка / глобальный обработчик</param>
        ''' <returns>Пустую строку; в случае ошибки - #Error.</returns>
        Public Overridable Function SetPropertyWithEvent(ByVal classId As Integer, ByVal propertyName As String, ByVal propertyValue As String, _
                                                ByVal child2Id As Integer, ByVal child3Id As Integer, ByVal isContainsCode As ContainsCodeEnum, ByVal isUserFunction As Boolean, _
                                                Optional ByVal ignoreBattle As Boolean = False) As String
            If isContainsCode = ContainsCodeEnum.CODE OrElse isContainsCode = ContainsCodeEnum.LONG_TEXT Then
                'в свойстве код
                Return SetPropertyWithEvent(classId, propertyName, mScript.DeserializeCodeData(propertyValue), child2Id, child3Id, propertyValue, isUserFunction, ignoreBattle)
            ElseIf isContainsCode = ContainsCodeEnum.EXECUTABLE_STRING Then
                'исполняемая строка
                'получаем codeBox.CodeData исполняемой строки
                Dim strToPrint As String = mScript.PrepareStringToPrint(propertyValue, Nothing, False) 'получаем строку без кавычек (код не исполняем)
                Dim dt() As CodeTextBox.CodeDataType
                With questEnvironment.codeBoxShadowed
                    .Text = strToPrint
                    dt = CopyCodeDataArray(.codeBox.CodeData)
                    propertyValue = WrapString(.Text)
                    .Text = ""
                End With
                Return SetPropertyWithEvent(classId, propertyName, dt, child2Id, child3Id, propertyValue, True, isUserFunction, ignoreBattle)
            Else
                'простое значение
                SetPropertyValue(classId, propertyName, propertyValue, child2Id, child3Id, ignoreBattle)
                Return ""
            End If
        End Function


        ''' <summary>
        ''' Устанавливает обработчик события или заменяет его на новый, если таковой уже создан, а также устанавливает все необходимые параметры свойства в mainClass.
        ''' </summary>
        ''' <param name="classId">Индекс класса - хозяина обработчика события</param>
        ''' <param name="propertyName">Название события-обработчика</param>
        ''' <param name="dt">Новый код для вставки</param>
        ''' <param name="funcParams">Параметры блока Event</param>
        ''' <param name="xmlData">Сериализованный CodeDataType для вставки в свойство</param>
        ''' <param name="IsExecutabelString">Является содержимое xmlData исполняемой строкой</param>
        ''' <returns>Пустую строку; в случае ошибки - #Error.</returns>
        Public Overridable Function SetPropertyWithEvent(ByVal classId As Integer, ByVal propertyName As String, ByRef dt() As CodeTextBox.CodeDataType, _
                                                ByRef funcParams() As String, Optional ByVal xmlData As String = "", Optional IsExecutabelString As Boolean = False, _
                                                Optional isUserFunction As Boolean = False, Optional makeXml As Boolean = True, Optional ByVal ignoreBattle As Boolean = False) As String

            Dim child2Id As Integer = -1, child3Id As Integer = -1
            If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then
                If GVARS.G_ISBATTLE AndAlso ignoreBattle = False AndAlso mScript.mainClass(classId).Names(0) = "H" Then
                    child2Id = mScript.Battle.GetFighterByName(funcParams(0))
                Else
                    child2Id = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
                    If funcParams.Count > 1 AndAlso child2Id > -1 Then child3Id = GetThirdChildIdByName(funcParams(1), child2Id, mScript.mainClass(classId).ChildProperties)
                End If
            End If

            Return SetPropertyWithEvent(classId, propertyName, dt, child2Id, child3Id, xmlData, IsExecutabelString, isUserFunction, makeXml, ignoreBattle)
        End Function

        ''' <summary>
        ''' Устанавливает обработчик события или заменяет его на новый, если таковой уже создан, а также устанавливает все необходимые параметры свойства в mainClass
        ''' </summary>
        ''' <param name="classId">Индекс класса - хозяина обработчика события</param>
        ''' <param name="propertyName">Название обработчика</param>
        ''' <param name="dt">Новый код для вставки</param>
        ''' <param name="level2Id">Id элемента 2 порядка, которому устанавливается обработчик, или -1, если устанавливаем глобальный обработчик</param>
        ''' <param name="level3Id">Id элемента 3 порядка, которому устанавливается обработчик, или -1, если устанавливаем обработчик элемента 2 порядка / глобальный обработчик</param>
        ''' <param name="xmlData">Сериализованный CodeDataType для вставки в свойство</param>
        ''' <param name="IsExecutabelString">Является содержимое xmlData исполняемой строкой</param>
        ''' <returns>Пустую строку; в случае ошибки - #Error.</returns>
        Public Overridable Function SetPropertyWithEvent(ByVal classId As Integer, ByVal propertyName As String, ByRef dt() As CodeTextBox.CodeDataType, ByVal level2Id As Integer, _
                                                   Optional ByVal level3Id As Integer = -1, Optional ByVal xmlData As String = "", Optional IsExecutabelString As Boolean = False, _
                                                   Optional isUserFunction As Boolean = False, Optional makeXml As Boolean = True, Optional ByVal ignoreBattle As Boolean = False) As String

            If makeXml Then
                If IsExecutabelString = False Then
                    If xmlData.Length = 0 Then xmlData = mScript.SerializeCodeData(dt) 'сериализуем код если это еще не было сделано
                Else
                    Do While xmlData.EndsWith(vbLf + "'")
                        xmlData = xmlData.Substring(0, xmlData.Length - 2) + "'"
                    Loop
                End If
            End If

            Dim eventId As Integer = CalculateNextId() 'id следующего события, если не будет найдено уже установленного
            'устанавливаем свайства события в mainClass
            If level2Id = -1 Then
                'корневой обработчик (1-го уровня)
                Dim propData As MatewScript.PropertiesInfoType = Nothing
                If isUserFunction Then
                    If mScript.mainClass(classId).Functions.TryGetValue(propertyName, propData) = False Then
                        mScript.LAST_ERROR = "Класс " + mScript.mainClass(classId).Names(0) + " не найден."
                        Return "#Error"
                    End If
                Else
                    If mScript.mainClass(classId).Properties.TryGetValue(propertyName, propData) = False Then
                        mScript.LAST_ERROR = "Класс " + mScript.mainClass(classId).Names(0) + " не найден."
                        Return "#Error"
                    End If
                End If
                propData.Value = xmlData
                If propData.eventId > 0 Then
                    eventId = propData.eventId 'событие уже установлено - получаем его id
                Else
                    propData.eventId = eventId 'событие еще не установлено - устанавливаем новый id
                End If
            ElseIf level3Id = -1 Then
                'обработчик объекта 2-го уровня
                If level2Id >= mScript.mainClass(classId).ChildProperties.Count Then
                    mScript.LAST_ERROR = "Свойство " + propertyName + " не найдено."
                    Return "#Error"
                End If
                Dim propData As MatewScript.ChildPropertiesInfoType = Nothing

                If GVARS.G_ISBATTLE AndAlso ignoreBattle = False AndAlso mScript.mainClass(classId).Names(0) = "H" Then
                    If mScript.Battle.Fighters(level2Id).heroProps.TryGetValue(propertyName, propData) = False Then
                        mScript.LAST_ERROR = "Свойство " + propertyName + " не найдено."
                        Return "#Error"
                    End If
                Else
                    If mScript.mainClass(classId).ChildProperties(level2Id).TryGetValue(propertyName, propData) = False Then
                        mScript.LAST_ERROR = "Свойство " + propertyName + " не найдено."
                        Return "#Error"
                    End If
                End If
                propData.Value = xmlData
                If propData.eventId > 0 Then
                    eventId = propData.eventId
                Else
                    propData.eventId = eventId
                End If
            Else
                'обработчик объекта 3-го уровня
                If level2Id >= mScript.mainClass(classId).ChildProperties.Count Then
                    mScript.LAST_ERROR = "Объект класса " + mScript.mainClass(classId).Names(0) + " с Id = " + level2Id.ToString + " не найден."
                    Return "#Error"
                End If
                Dim propData As MatewScript.ChildPropertiesInfoType = Nothing
                If mScript.mainClass(classId).ChildProperties(level2Id).TryGetValue(propertyName, propData) = False Then
                    mScript.LAST_ERROR = "Свойство " + propertyName + " не найдено."
                    Return "#Error"
                End If
                If level3Id >= propData.ThirdLevelProperties.Count Then
                    mScript.LAST_ERROR = "Объект класса " + mScript.mainClass(classId).Names(0) + " с Id = " + level2Id.ToString + ";" + level3Id.ToString + " не найден."
                    Return "#Error"
                End If
                propData.ThirdLevelProperties(level3Id) = xmlData
                If propData.ThirdLevelEventId(level3Id) > 0 Then
                    eventId = propData.ThirdLevelEventId(level3Id)
                Else
                    propData.ThirdLevelEventId(level3Id) = eventId
                End If
            End If

            SetEventId(eventId, dt)
            Return ""
        End Function

        ''' <summary>
        ''' Удаляет событие
        ''' </summary>
        ''' <param name="eventId">Индекс удалемого события</param>
        ''' <remarks></remarks>
        Public Sub RemoveEvent(ByVal eventId As Integer)
            If eventId < 1 AndAlso lstEvents.ContainsKey(eventId) = False Then Return
            'If lstEvents.ContainsKey(eventId) = False Then
            '    MessageBox.Show("RemoveEvent: неправильный идентификатор удаляемого события!")
            '    Return
            'End If
            lstEvents.Remove(eventId)
            'удаляем переменные
            If lstVariables.ContainsKey(eventId) Then
                If IsNothing(lstVariables(eventId)) = False Then lstVariables(eventId).Clear()
                lstVariables.Remove(eventId)
            End If
            'If lastEventId > lstEvents.Count Then lastEventId = lstEvents.Count - 1
        End Sub

        ''' <summary>
        ''' Возвращает Id события свойства или -1 если его нет
        ''' </summary>
        ''' <param name="classId">Id класса события</param>
        ''' <param name="propName">Имя свойства</param>
        ''' <param name="child2Id">Id элемента 2 порядка</param>
        ''' <param name="child3Id">Id элемента 3 порядка</param>
        Public Function GetEventId(ByVal classId As Integer, ByVal propName As String, ByVal child2Id As Integer, ByVal child3Id As Integer) As Integer
            If mScript.mainClass(classId).Properties.ContainsKey(propName) = False Then Return -1

            If child2Id < 0 Then
                Return mScript.mainClass(classId).Properties(propName).eventId
            ElseIf child3Id < 0 Then
                Return mScript.mainClass(classId).ChildProperties(child2Id)(propName).eventId
            Else
                Return mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelEventId(child3Id)
            End If
        End Function

        ''' <summary>
        ''' Дублирует код из события с индексом eventIdToDuptlicate
        ''' </summary>
        ''' <param name="eventIdToDuptlicate">Id события, которое дублируется</param>
        ''' <param name="destId">Id события, в которое помещается копия кода или -1, если надо создать новое</param>
        ''' <returns>Id копии события</returns>
        Public Overridable Function DuplicateEvent(ByVal eventIdToDuptlicate As Integer, Optional destId As Integer = -1) As Integer
            Dim eSrc As List(Of MatewScript.ExecuteDataType) = Nothing

            If eventIdToDuptlicate = -1 OrElse lstEvents.TryGetValue(eventIdToDuptlicate, eSrc) = False Then
                MessageBox.Show("DuplicateEvent: неправильный идентификатор ресурса для дублирования!")
                Return -1
            End If
            If destId > 0 AndAlso lstEvents.ContainsKey(destId) = False Then
                MessageBox.Show("DuplicateEvent: неправильный идентификатор места назначения для дублирования!")
                Return -1
            End If
            'создаем копию
            Dim arrDest() As MatewScript.ExecuteDataType
            If IsNothing(eSrc) OrElse eSrc.Count = 0 Then
                ReDim arrDest(0)
            Else
                ReDim arrDest(eSrc.Count - 1)
                eSrc.CopyTo(arrDest)
                'For i As Integer = 0 To arrDest.Count - 1
                '    arrDest(i) = eSrc(i).Clone
                'Next
            End If
            'вставляем копию в массив
            If destId < 1 Then
                destId = CreateNewId()
                lstEvents.Add(destId, arrDest.ToList)
            Else
                lstEvents(destId) = arrDest.ToList
            End If

            'дублируем переменные
            If lstVariables.ContainsKey(eventIdToDuptlicate) Then
                Dim varSrc As SortedList(Of String, cVariable.variableEditorInfoType) = lstVariables(eventIdToDuptlicate)
                Dim varDest As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                mScript.csLocalVariables.CopyVariables(varDest, varSrc)
                lstVariables.Add(destId, varDest)
            End If

            Return destId
        End Function

        ''' <summary>
        ''' Устанавливает на сохраненое событие новый код либо создает новое событие.
        ''' </summary>
        ''' <param name="eventId">Индекс вставляемого события или -1, если надо создать новый</param>
        ''' <param name="dt">Новый код для вставки</param>
        ''' <remarks>При предварительном вызове SetEvent не является необходимым</remarks>
        Public Overridable Function SetEventId(ByVal eventId As Integer, ByRef dt() As CodeTextBox.CodeDataType) As Integer
            If eventId > 0 AndAlso lstEvents.ContainsKey(eventId) Then
                'нашли сохраненное событие
                lstEvents(eventId) = mScript.PrepareBlock(dt)

                'привязка переменных
                If IsNothing(mScript.csLocalVariables.lstVariables) = False AndAlso mScript.csLocalVariables.lstVariables.Count > 0 Then
                    Dim varCopy As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                    mScript.csLocalVariables.CopyVariables(varCopy)

                    Dim prevVars As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                    If lstVariables.TryGetValue(eventId, prevVars) Then
                        If IsNothing(prevVars) = False Then prevVars.Clear()
                        lstVariables(eventId) = varCopy
                    Else
                        lstVariables.Add(eventId, varCopy)
                    End If
                End If

                Return eventId
            End If
            'добавляем новое событие
            Dim newId As Integer = eventId
            If newId <= 0 Then
                newId = CreateNewId()
            ElseIf newId = CalculateNextId() Then
                newId = CreateNewId()
            End If
            lstEvents.Add(newId, mScript.PrepareBlock(dt))

            'привязка переменных
            If IsNothing(mScript.csLocalVariables.lstVariables) = False AndAlso mScript.csLocalVariables.lstVariables.Count > 0 Then
                Dim varCopy As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                mScript.csLocalVariables.CopyVariables(varCopy)
                lstVariables.Add(newId, varCopy)
            End If

            Return newId
        End Function

        ''' <summary>
        ''' Устанавливает на сохраненое событие новый код либо создает новое событие.
        ''' </summary>
        ''' <param name="eventId">Индекс вставляемого события или -1, если надо создать новый</param>
        ''' <param name="executeDt">Новый код для вставки (уже в готовом виде)</param>
        ''' <remarks>При предварительном вызове SetEvent не является необходимым</remarks>
        Public Overridable Function SetEventId(ByVal eventId As Integer, ByRef executeDt As List(Of ExecuteDataType)) As Integer
            If eventId > 0 AndAlso lstEvents.ContainsKey(eventId) Then
                'нашли сохраненное событие
                lstEvents(eventId) = executeDt

                'привязка переменных
                If IsNothing(mScript.csLocalVariables.lstVariables) = False AndAlso mScript.csLocalVariables.lstVariables.Count > 0 Then
                    Dim varCopy As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                    If lstVariables.TryGetValue(eventId, varCopy) Then
                        varCopy.Clear()
                    Else
                        lstVariables.Add(eventId, varCopy)
                    End If
                    mScript.csLocalVariables.CopyVariables(varCopy)
                End If

                Return eventId
            End If
            'добавляем новое событие
            Dim newId As Integer = eventId
            If newId <= 0 Then
                newId = CreateNewId()
            ElseIf newId = CalculateNextId() Then
                newId = CreateNewId()
            End If
            lstEvents.Add(newId, executeDt)

            'привязка переменных
            If IsNothing(mScript.csLocalVariables.lstVariables) = False AndAlso mScript.csLocalVariables.lstVariables.Count > 0 Then
                Dim varCopy As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                mScript.csLocalVariables.CopyVariables(varCopy)
                lstVariables.Add(newId, varCopy)
            End If

            Return newId
        End Function

        ''' <summary>
        ''' Создает новое событие для указанного свойства и устанавливает в mainClass его eventId. Проверка на наличие уже созданного события для данного свойства не проводится.
        ''' </summary>
        ''' <param name="classId">Индекс класса - хозяина обработчика события</param>
        ''' <param name="propName">Имя свойства</param>
        ''' <param name="propValue">Значение свойства с кодом или исполняемой строкой</param>
        ''' <param name="containsCode">Что именно - код или исполняемая строка</param>
        ''' <param name="child2Id">Имя/ Id элемента 2 порядка или -1)</param>
        ''' <param name="child3Id">Имя/ Id элемента 3 порядка или -1)</param>
        ''' <returns>Id созданного события или -1, если переданные данные некорректны</returns>
        Public Overridable Function SetEventId(ByVal classId As Integer, ByVal propName As String, ByVal propValue As String, ByVal containsCode As ContainsCodeEnum, _
                                               ByVal child2Id As Integer, ByVal child3Id As Integer) As Integer
            If mScript.mainClass(classId).Properties.ContainsKey(propName) = False Then Return -1

            With questEnvironment.codeBoxShadowed.codeBox
                If containsCode = ContainsCodeEnum.CODE OrElse containsCode = ContainsCodeEnum.LONG_TEXT Then
                    .LoadCodeFromProperty(propValue)
                ElseIf containsCode = ContainsCodeEnum.EXECUTABLE_STRING Then
                    .Text = propValue
                Else
                    Return -1
                End If
            End With

            Dim newId As Integer = SetEventId(-1, questEnvironment.codeBoxShadowed.codeBox.CodeData)
            questEnvironment.codeBoxShadowed.Text = ""

            If child2Id = -1 Then
                Dim p As PropertiesInfoType = mScript.mainClass(classId).Properties(propName)
                p.eventId = newId
            ElseIf child3Id = -1 Then
                Dim p As ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(propName)
                p.eventId = newId
            Else
                Dim p As ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)(propName)
                p.ThirdLevelEventId(child3Id) = newId
            End If

            Return newId
        End Function

        ''' <summary>
        ''' Creates list of htlm element IDs in blocks HTML...END HTML inside the code of the event
        ''' </summary>
        ''' <param name="eventId">Event id to search</param>
        ''' <param name="lstId">destination list to put new id's</param>
        ''' <remarks></remarks>
        Public Sub MakeElementIdListFromHTMLBlocks(ByVal eventId As Integer, ByRef lstId As List(Of String))
            Dim dt As List(Of ExecuteDataType) = Nothing
            If eventId < 1 OrElse lstEvents.TryGetValue(eventId, dt) = False Then Return
            If IsNothing(dt) OrElse dt.Count = 0 Then Return

            Dim isInBlock As Boolean = False
            For curLine As Integer = 0 To dt.Count - 1
                If IsNothing(dt(curLine)) OrElse IsNothing(dt(curLine).Code) OrElse dt(curLine).Code.Count = 0 Then Continue For
                Dim curCode() As CodeTextBox.EditWordType = dt(curLine).Code
                If isInBlock = False Then
                    If curCode(0).wordType = CodeTextBox.EditWordTypeEnum.W_HTML AndAlso curCode(0).Word.ToLower.Trim = "html" Then
                        'found HTML ....
                        isInBlock = True
                        Continue For
                    End If
                ElseIf curCode.Count > 1 Then
                    If curCode(0).Word.Trim.ToLower = "end" AndAlso curCode(1).Word.Trim.ToLower = "html" Then
                        'block has finished
                        isInBlock = False
                        Continue For
                    End If
                End If

                If Not isInBlock Then Continue For
                'now we're inside a block HTML
                For i As Integer = 0 To curCode.Count - 1
                    If curCode(i).wordType <> CodeTextBox.EditWordTypeEnum.W_HTML_DATA Then Continue For
                    Dim curText As String = curCode(i).Word, curPos As Integer = 0

                    curPos = curText.IndexOf(" id", curPos, System.StringComparison.CurrentCultureIgnoreCase)
                    If curPos = -1 Then Continue For
                    'string " id ...." has been found
                    Dim resStr As String = curText.Substring(curPos + 3).Replace(" ", "")
                    Dim brChar As Char = resStr.Chars(0)
                    If brChar <> "="c OrElse resStr.Length < 3 Then Continue For
                    brChar = resStr(1)
                    If brChar <> "'"c AndAlso brChar <> Chr(34) Then Continue For

                    'the string is id='xxx...
                    Dim endPos As Integer = resStr.IndexOf(brChar, 2)
                    If endPos = -1 Then Continue For
                    Dim newId As String = resStr.Substring(2, endPos - 2)
                    If lstId.Contains(newId) = False Then lstId.Add(newId)
                Next i
            Next curLine
        End Sub

        '(пока не нужны, но должны работать)
        ' ''' <summary>
        ' ''' Возвращает Id события указанного свойства. Если такового не сущетвует, то возвращает -1
        ' ''' </summary>
        ' ''' <param name="classId">Класс события</param>
        ' ''' <param name="propName">Имя свойства</param>
        ' ''' <param name="child2Id">Id элемента 2 порядка указанного свойства или -1, если возвращаем Id глобального обработчика</param>
        ' ''' <param name="child3Id">Id элемента 3 порядка указанного свойства или -1, если возвращаем Id элемента 2 порядка / глобального обработчика</param>
        ' ''' <returns>Id события. Если Id > 0, то оно гарантированно существует</returns>
        'Public Overridable Function GetPropertyEventId(ByVal classId As Integer, ByVal propName As String, Optional child2Id As Integer = -1, Optional child3Id As Integer = -1) As Integer
        '    If mScript.mainClass(classId).Properties.ContainsKey(propName) = False Then Return -1
        '    Dim eventId As Integer

        '    If child2Id = -1 Then
        '        eventId = mScript.mainClass(classId).Properties(propName).eventId
        '    ElseIf child3Id = -1 Then
        '        eventId = mScript.mainClass(classId).ChildProperties(child2Id)(propName).eventId
        '    Else
        '        eventId = mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelEventId(child3Id)
        '    End If
        '    If eventId = 0 Then eventId = -1
        '    If eventId > 0 Then
        '        If lstEvents.ContainsKey(eventId) Then eventId = -1
        '    End If

        '    Return eventId
        'End Function

        'Public Overridable Function GetPropertyEventId(ByVal classId As Integer, ByVal propName As String, ByRef funcParams() As String) As Integer
        '    If mScript.mainClass(classId).Properties.ContainsKey(propName) = False Then Return -1

        '    Dim child2Id As Integer = -1, child3Id As Integer = -1
        '    If IsNothing(funcParams) = False AndAlso funcParams.Count > 0 Then
        '        child2Id = GetSecondChildIdByName(funcParams(0), mScript.mainClass(classId).ChildProperties)
        '        If funcParams.Count > 1 AndAlso child2Id > -1 Then child3Id = GetThirdChildIdByName(funcParams(1), child2Id, mScript.mainClass(classId).ChildProperties)
        '    End If

        '    Return GetPropertyEventId(classId, propName, child2Id, child3Id)
        'End Function

        ''' <summary>
        ''' Возвращает существует ли событие с указанным идентификатором
        ''' </summary>
        ''' <param name="eventId">Id события для поиска</param>
        Public Function IsEventIdExists(ByVal eventId As Integer) As Boolean
            If eventId < 1 Then Return False
            Return lstEvents.ContainsKey(eventId)
        End Function

        ''' <summary>Запускает событие на выполнение и возвращает его результат</summary>
        ''' <returns>Результат события</returns>
        Public Function RunEvent(ByVal EventId As Integer, ByRef arrParams() As String, ByVal eventName As String, ByVal runAfterScriptEvent As Boolean) As String
            Dim e As List(Of ExecuteDataType) = Nothing
            If EventId > 0 AndAlso lstEvents.TryGetValue(EventId, e) Then
                If IsNothing(e) OrElse e.Count = 0 OrElse (e.Count = 1 AndAlso IsNothing(e(0))) Then Return ""
                'загрузка локальных переменных
                Dim varCopy As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                Dim prevVars As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                mScript.csLocalVariables.CopyVariables(prevVars)
                If lstVariables.TryGetValue(EventId, varCopy) Then mScript.csLocalVariables.RestoreVariables(varCopy, True)
                'запускаем событие и возвращаем результат
                If String.IsNullOrEmpty(eventName) = False Then mScript.codeStack.Push(eventName)
                Dim res As String = mScript.ExecuteCode(e, arrParams, True)

                'запускаем ScriptFinishedEvent
                If runAfterScriptEvent Then
                    Dim classId As Integer = mScript.mainClassHash("Script")
                    Dim after_eventId As Integer = mScript.mainClass(classId).Properties("ScriptFinishedEvent").eventId
                    If after_eventId > 0 AndAlso lstEvents.TryGetValue(after_eventId, e) Then
                        'загрузка локальных переменных ScriptFinishedEvent
                        varCopy = Nothing
                        If lstVariables.TryGetValue(after_eventId, varCopy) Then mScript.csLocalVariables.RestoreVariables(varCopy, True)
                        mScript.codeStack.Push("ScriptFinishedEvent")
                        'запускаем ScriptFinishedEvent и возвращаем результат
                        Dim res2 As String = mScript.ExecuteCode(e, arrParams, True)
                        mScript.codeStack.Pop()
                        If res2 = "#Error" Then Return res2
                    End If
                End If

                mScript.csLocalVariables.RestoreVariables(prevVars)
                If String.IsNullOrEmpty(eventName) = False Then mScript.codeStack.Pop()
                mScript.EXIT_CODE = False
                Return res
            End If
            Return ""
        End Function

        ''' <summary>Запускает событие ScriptFinishedEvent и возвращает его результат</summary>
        ''' <returns>Результат события</returns>
        Public Function RunScriptFinishedEvent(ByRef arrParams() As String) As String
            Dim classId As Integer = mScript.mainClassHash("Script")
            Dim after_eventId As Integer = mScript.mainClass(classId).Properties("ScriptFinishedEvent").eventId
            Dim e As List(Of ExecuteDataType) = Nothing
            If after_eventId > 0 AndAlso lstEvents.TryGetValue(after_eventId, e) Then
                'загрузка локальных переменных ScriptFinishedEvent
                Dim varCopy As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                If lstVariables.TryGetValue(after_eventId, varCopy) Then mScript.csLocalVariables.RestoreVariables(varCopy, True)
                mScript.codeStack.Push("ScriptFinishedEvent")
                'запускаем ScriptFinishedEvent и возвращаем результат
                Dim prevVars As SortedList(Of String, cVariable.variableEditorInfoType) = Nothing
                mScript.csLocalVariables.CopyVariables(prevVars)
                Dim res As String = mScript.ExecuteCode(e, arrParams, True)
                mScript.csLocalVariables.RestoreVariables(prevVars)
                mScript.codeStack.Pop()
                Return res
            End If
            Return ""
        End Function

        Public Sub New(ByRef mScript As MatewScript)
            Me.mScript = mScript
        End Sub

        ''' <summary>
        ''' Стирает все события данного элемента (дочерние элементы не затрагиваются).
        ''' </summary>
        ''' <param name="classId">Индекс класса - хозяина обработчика события</param>
        ''' <param name="child2Id">Id элемента 2 порядка или -1)</param>
        ''' <param name="child3Id">Id элемента 3 порядка или -1)</param>
        ''' <param name="parentId">Id локации (только для действий)</param>
        Public Sub EraseElementEvents(ByVal classId As Integer, ByVal child2Id As Integer, Optional ByVal child3Id As Integer = -1, Optional ByVal parentId As Integer = -1)
            If parentId > -1 AndAlso actionsRouter.GetActiveLocationId <> parentId Then
                'Тут можно все корректно описать и сделать, но, скорее всего, это просто не надо
                MessageBox.Show("Попытка удаления событий действия до открытия соответствующей ему локации", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If
            For pId As Integer = 0 To mScript.mainClass(classId).ChildProperties(child2Id).Count - 1
                'Перебираем все события. Если у них есть скрипты - удаляем
                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id).ElementAt(pId).Value
                Dim eventId As Integer
                If child3Id > -1 Then
                    eventId = ch.ThirdLevelEventId(child3Id)
                Else
                    eventId = ch.eventId
                End If
                RemoveEvent(eventId)
                If child3Id > -1 Then
                    ch.ThirdLevelProperties(child3Id) = -1
                Else
                    ch.eventId = -1
                End If
            Next pId
        End Sub

        Private stopCharArray() As Char = {" "c, "("c, ")"c, "["c, "]"c, "."c, ","c, "'"c, "="c, "+"c, "-"c, "*"c, "^"c, "/"c, "\"c, "<"c, ">"c, "&"c, "!"c, ";"c, "#"c, "$"c, Chr(34), "?"}
        Public Function ConverTextToCodeData(ByVal strCode As String, Optional asLongText As Boolean = False) As CodeTextBox.CodeDataType()

            Dim CodeData() As CodeTextBox.CodeDataType

            'проверка на пустой текст
            strCode = strCode.Trim
            If String.IsNullOrEmpty(strCode) Then
                ReDim CodeData(0)
                CodeData(0) = New CodeTextBox.CodeDataType
                Return CodeData
            End If

            Dim codeLines() As String = Split(strCode, vbNewLine)
            Dim endLine As Integer = codeLines.Count - 1
            ReDim CodeData(endLine)


            Dim curString As String = "" 'текущая строка
            Dim i, j As Integer

            Dim arrayWords() As CodeTextBox.EditWordType = Nothing, arrayWordsUBound As Integer = 0 'массив, в котором храним слова из строки, ее размер
            Dim arrayWordsCopy() As CodeTextBox.EditWordType = Nothing  'массив для сборки строки, разбитой _
            Dim ovalBracketBalance As Integer = 0 'баланс ( и )
            Dim quadBracketBalance As Integer = 0 'баланс [ и ]
            Dim wordStart As Integer 'положение в строке первого символа слова, с которым будем работать
            Dim isPrevStringDissociation As Boolean 'разбита ли строка с помощью _
            Dim execOpenPos As Integer 'для нахождения тега <exec>
            Dim isInExec As CodeTextBox.ExecBlockEnum = CodeTextBox.ExecBlockEnum.NO_EXEC
            Dim isInText As CodeTextBox.TextBlockEnum = asLongText   'находимся ли мы внутри блока HTML или Wrap
            Dim textStartLine As Integer 'индекс первой строки с текстом, идущей после начала блока Wrap или HTML 
            If isInText <> CodeTextBox.TextBlockEnum.NO_TEXT_BLOCK Then textStartLine = 0 'если начало текстового блока до начала кода - textStartLine = первой обрабатываемой строке

            i = 0
            Dim startLine As Integer = 0
            While i <= endLine
                j = -1
                wordStart = 0
                If isInText <> CodeTextBox.TextBlockEnum.NO_TEXT_BLOCK Then
                    'мы внутри блока HTML / Wrap
                    curString = codeLines(i)
                    Dim strLower As String = curString.Trim.Replace("  ", " ").ToLower
                    Dim strSearch As String = IIf(isInText = CodeTextBox.TextBlockEnum.TEXT_HTML, "end html", "end wrap")
                    If strLower.StartsWith(strSearch) AndAlso StripComments(strLower, Nothing, Nothing) <> 2 AndAlso strLower = strSearch Then
                        'StripComments <> 2 выполнено в блоке If Then только чтоб очистить curString от комментариев и пробелов перед третьей проверкой
                        'текущая строка - "End HTML" или "End Wrap" - конец блока
                        isInText = CodeTextBox.TextBlockEnum.NO_TEXT_BLOCK
                        isInExec = CodeTextBox.ExecBlockEnum.NO_EXEC
                    Else
                        'продолжение текстового блока
                        If isInExec = CodeTextBox.ExecBlockEnum.NO_EXEC Then
                            execOpenPos = curString.IndexOf("<exec>", StringComparison.CurrentCultureIgnoreCase)
                            If execOpenPos > -1 Then
                                isInExec = CodeTextBox.ExecBlockEnum.EXEC_SINGLE
                            End If

                            'записываем текстовый блок в массив CodeData
                            ReDim CodeData(i).Code(0)
                            If isInExec Then
                                execOpenPos += "<exec>".Length
                                ReDim arrayWords(20)
                                arrayWordsUBound = 0 'id последней заполненной ячейки массива
                                arrayWords(0).Word = curString.Substring(0, execOpenPos)
                                arrayWords(0).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA
                                j = execOpenPos - 1
                                wordStart = execOpenPos
                            Else
                                CodeData(i).Code(0).Word = codeLines(i)
                                CodeData(i).Code(0).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA
                                i += 1
                                Continue While
                            End If
                        End If
                    End If
                End If

                'определяем, не заканчивается ли строка перед первой, которую мы обрабатываем, на _
                If i = 0 And i > 0 AndAlso isInExec <> CodeTextBox.ExecBlockEnum.EXEC_SINGLE Then
                    curString = codeLines(i - 1)
                    If StripComments(curString, Nothing, Nothing) > -1 Then
                        If curString.Length > 0 AndAlso curString.EndsWith("_") Then
                            'если предыдущая строка - начало текущей (самой первой) - перезапускаем цикл, начиная с предыдущей строки
                            i -= 1
                            startLine = i
                            Continue While
                        End If
                    End If
                End If

                curString = codeLines(i) 'получаем текущую линию

                'Очищаем строку от дополнительных элементов (пробелов в начале, комментариев...)
                If isInExec <> CodeTextBox.ExecBlockEnum.EXEC_SINGLE AndAlso StripComments(curString, CodeData(i).Comments, CodeData(i).StartingSpaces) = -1 Then
                    Return Nothing
                End If
                If IsNothing(CodeData(i).Comments) Then CodeData(i).Comments = ""
                If IsNothing(CodeData(i).StartingSpaces) Then CodeData(i).StartingSpaces = ""

                'в строке были только пробелы и комментарии - переходим сразу к новой строке
                If curString.Length = 0 Then
                    Erase CodeData(i).Code
                    i += 1
                    Continue While
                End If

                'Получаем в массив arrayWords слова из рабочей строки
                If isInExec <> CodeTextBox.ExecBlockEnum.EXEC_SINGLE Then
                    ReDim arrayWords(20) 'расширяем массив слов сразу до Х, чтобы не делать это каждый раз при добавлении слова
                    arrayWordsUBound = -1 'id последней заполненной ячейки массива (-1 - пусто)
                End If

                Do
                    If arrayWordsUBound + 2 > arrayWords.GetUpperBound(0) Then ReDim Preserve arrayWords(arrayWordsUBound + 10) 'расширяем массив если было мало
                    'поиск первого попавшегося символа из массива stopCharArray:
                    'stopCharArray {" ", "(", ")", "[", "]", ".", ",", ,"'", "=", "+", "-", "*", "^", "/", "\", "<", ">", "&", "!", ";", "#", "$", """}
                    j = curString.IndexOfAny(stopCharArray, j + 1)

                    If j = -1 Then
                        'таких символов не найдено - последнее слово
                        If wordStart < curString.Length Then
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart)
                        End If
                        Exit Do
                    End If
                    'в зависимости от того, что нашли, заполняем массив arrayWords
                    Select Case curString.Chars(j)
                        Case " "c
                            If j > wordStart Then
                                'если между предыдущим и текущим оператором было какое-то слово - сохраняем его
                                arrayWordsUBound += 1 'будет заполнена новая ячейка массива
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart) 'заполняем ячейку
                            End If
                            wordStart = j + 1 'начало следующего слова - на символ дальше
                        Case "("c
                            ovalBracketBalance += 1
                            If j > wordStart Then
                                'если есть слово до (, то добавляем его в массив arrayWords
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "("
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN 'устанавливаем тип содержимого
                            wordStart = j + 1
                        Case ")"c
                            ovalBracketBalance -= 1
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = ")"
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                            wordStart = j + 1
                        Case "["c
                            quadBracketBalance += 1
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "["
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                            wordStart = j + 1
                        Case "]"c
                            quadBracketBalance -= 1
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "]"
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                            wordStart = j + 1
                        Case "."c
                            If j > 0 Then '56.34
                                If Integer.TryParse(curString.Chars(j - 1), Nothing) = True Then Continue Do
                            End If
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "."
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_POINT
                            wordStart = j + 1
                        Case ","c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = ", "
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA
                            wordStart = j + 1
                        Case "="c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = " = "
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL
                            wordStart = j + 1
                        Case "?"
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "?"
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_RETURN
                            wordStart = j + 1
                        Case "&"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = " & "
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER
                            wordStart = j + 1
                        Case "#"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "#"
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER
                            wordStart = j + 1
                        Case "$"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "$"
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING
                            wordStart = j + 1
                        Case "+"c, "-"c, "*"c, "^"c, "/"c, "\"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            Dim strOperatorFull As String = curString.Chars(j)
                            If j < curString.Length - 1 Then
                                Dim strUnar As String = curString.Substring(j + 1).TrimStart(" "c, Chr(1)) 'в strUnar - строка за текущим символом, за вычетом левых пробелов и символа каретки
                                If strUnar.Length > 0 AndAlso strUnar.Chars(0) = "="c Then
                                    'получаем полный оператор, если это += или -=
                                    If curString.Chars(j) = "+"c OrElse curString.Chars(j) = "-"c Then strOperatorFull += "="
                                    j = curString.IndexOf("="c, j + 1) - 1
                                ElseIf curString.Chars(j) = "-"c Then
                                    If strUnar.Length > 0 AndAlso Integer.TryParse(strUnar.Chars(0), Nothing) = True Then
                                        If arrayWordsUBound > 0 AndAlso (arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA) Then
                                            'если за знаком "-" идет число, а перед ним - математический оператор или оператор сравнения или ( [ , то 
                                            'это - унарный минус, который относится к следующей за ним цифре. Продолжаем цикл, не сохраняя минус отдельно
                                            wordStart = j
                                            curString = curString.Substring(0, j + 1) + curString.Substring(j + 1).TrimStart
                                            Continue Do
                                        End If
                                        If arrayWordsUBound = -1 Then
                                            strUnar = ""
                                        Else
                                            strUnar = arrayWords(arrayWordsUBound).Word.ToLower
                                        End If
                                        If strUnar = "to" OrElse strUnar = "case" OrElse strUnar = "step" OrElse strUnar = "return" OrElse strUnar = "?" Then
                                            'если за знаком "-" идет число, а перед ним - To Case или  Step, то 
                                            'это - унарный минус, который относится к следующей за ним цифре. Продолжаем цикл, не сохраняя минус отдельно
                                            wordStart = j
                                            curString = curString.Substring(0, j + 1) + curString.Substring(j + 1).TrimStart
                                            Continue Do
                                        End If
                                        If isPrevStringDissociation And j = 0 Then
                                            If CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA Then
                                                'то же, только предыдущая строка заканчивалась на _, а минус - первый символ строки
                                                wordStart = j
                                                curString = curString.Substring(0, j + 1) + curString.Substring(j + 1).TrimStart
                                                Continue Do
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf arrayWordsUBound = -1 AndAlso isPrevStringDissociation AndAlso curString.Chars(j) = "-" Then
                                If CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                 CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse _
                                 CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse _
                                 CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = CodeTextBox.EditWordTypeEnum.W_COMMA Then
                                    'то же, только предыдущая строка заканчивалась на _, а минус - первый символ строки
                                    wordStart = j
                                    curString = curString.Substring(0, j + 1) + curString.Substring(j + 1).TrimStart
                                    Continue Do
                                End If
                            End If
                            arrayWordsUBound += 1
                            If j = 0 Then
                                arrayWords(arrayWordsUBound).Word = strOperatorFull + " "
                            Else
                                arrayWords(arrayWordsUBound).Word = " " + strOperatorFull + " "
                            End If
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH
                            wordStart = j + strOperatorFull.Length
                            j += strOperatorFull.Length - 1
                        Case "<"c, ">"c, "!"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            Dim strOperatorFull As String = curString.Chars(j)
                            If j < curString.Length - 1 Then
                                Dim strNextChar As String = curString.Substring(j + 1).TrimStart(" "c, Chr(1))
                                If strNextChar.Length > 0 AndAlso (strNextChar.First = ">" OrElse strNextChar.First = "=") Then
                                    'получаем полный оператор, если это <=, >=, <>, !=
                                    strOperatorFull += strNextChar.First 'curString.Chars(j + 1)
                                    j = curString.IndexOf(strNextChar.First, j + 1) - 1
                                End If
                            End If
                            If isInExec <> CodeTextBox.ExecBlockEnum.NO_EXEC AndAlso curString.Substring(j).StartsWith("</exec>", StringComparison.CurrentCultureIgnoreCase) Then
                                execOpenPos = curString.IndexOf("<exec>", j + 7, StringComparison.CurrentCultureIgnoreCase)
                                If execOpenPos = -1 Then
                                    arrayWords(arrayWordsUBound).Word = curString.Substring(j)
                                    arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA
                                    isInExec = CodeTextBox.ExecBlockEnum.NO_EXEC
                                    Exit Do
                                Else
                                    arrayWords(arrayWordsUBound).Word = curString.Substring(j, execOpenPos - j + 6)
                                    arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA
                                    j = execOpenPos + 5
                                    wordStart = j + 1
                                    Continue Do
                                End If
                            End If
                            arrayWords(arrayWordsUBound).Word = " " + strOperatorFull + " "
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE
                            wordStart = j + strOperatorFull.Length
                            j += strOperatorFull.Length - 1
                        Case "'"
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                                wordStart = j
                            End If
                            'Найдена строка '...'.
                            Dim closeQuotePos As Integer = j
qbQuoteSearching2:
                            closeQuotePos = curString.IndexOf("'", closeQuotePos + 1) 'ищем закрывающий "'"
                            If closeQuotePos = -1 Then
                                'закрывающий "'" не найден - ошибка
                                mScript.LAST_ERROR = "Не найдена закрывающая кавычка"
                                Return Nothing
                            End If
                            'Обрабатываем экранированную кавычку /'
                            If closeQuotePos > 0 AndAlso curString.Chars(closeQuotePos - 1) = "/"c Then
                                If closeQuotePos <> curString.Length - 1 Then GoTo qbQuoteSearching2
                            End If
                            'закрывающая кавычка найдена
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, closeQuotePos - wordStart + 1)
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING
                            wordStart = closeQuotePos + 1
                            j = closeQuotePos
                        Case Chr(34)
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                                wordStart = j
                            End If
                            'Найдена строка "...".
                            Dim closeQuotePos As Integer = j
qbQuoteSearching3:
                            closeQuotePos = curString.IndexOf(Chr(34), closeQuotePos + 1) 'ищем закрывающий "
                            If closeQuotePos = -1 Then
                                'закрывающий " не найден - ошибка
                                mScript.LAST_ERROR = "Не найдена закрывающая кавычка"
                                Return Nothing
                            End If
                            'Обрабатываем экранированную кавычку /"
                            If curString.Chars(closeQuotePos - 1) = "/"c Then
                                Mid(curString, closeQuotePos + 1, 1) = "'"
                                If closeQuotePos <> curString.Length - 1 Then GoTo qbQuoteSearching3
                            End If
                            'закрывающая кавычка найдена
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "'" + curString.Substring(wordStart + 1, closeQuotePos - wordStart - 1) + "'"
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING
                            wordStart = closeQuotePos + 1
                            j = closeQuotePos
                        Case ";"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "; "
                            arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION
                            wordStart = j + 1
                            If ovalBracketBalance <> 0 Then
                                mScript.LAST_ERROR = "Непарное кол-во открывающих и закрывающих круглых скобок."
                                Return Nothing
                            End If
                            If quadBracketBalance <> 0 Then
                                mScript.LAST_ERROR = "Непарное кол-во открывающих и закрывающих квадратных скобок."
                                Return Nothing
                            End If
                    End Select
                Loop
                If isInExec = CodeTextBox.ExecBlockEnum.EXEC_SINGLE Then isInExec = CodeTextBox.ExecBlockEnum.EXEC_MULTILINE

                'текущая строка заканчивается на _  - устанавливаем последнее слово = W_STRINGS_DISSOCIATION
                If arrayWords(arrayWordsUBound).Word = "_" Then
                    arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION
                ElseIf arrayWords(arrayWordsUBound).Word.EndsWith("_") Then
                    arrayWords(arrayWordsUBound).Word = arrayWords(arrayWordsUBound).Word.Substring(0, arrayWords(arrayWordsUBound).Word.Length - 1)
                    arrayWordsUBound += 1
                    arrayWords(arrayWordsUBound).Word = "_"
                    arrayWords(arrayWordsUBound).wordType = CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION
                End If

                ReDim Preserve arrayWords(arrayWordsUBound) 'убираем пустые значения в конце массива слов

                If arrayWords(arrayWordsUBound).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                    'проверка парности скобок лишь в том случае, если эта строка не заканчивается на _ (или это незавершенная строка)
                    If ovalBracketBalance <> 0 Then
                        mScript.LAST_ERROR = "Непарное кол-во открывающих и закрывающих круглых скобок."
                        Return Nothing
                    End If
                    If quadBracketBalance <> 0 Then
                        mScript.LAST_ERROR = "Непарное кол-во открывающих и закрывающих квадратных скобок."
                        Return Nothing
                    End If
                End If

                'получаем значения оставшихся слов, расставляем пробелы и проверяем синтаксис
                If AnalizeArrayWords(CodeData, arrayWords, isPrevStringDissociation, i) = -1 Then
                    Return Nothing
                End If
                CodeData(i).Code = arrayWords

                'собираем строку заново, уже с правильными пробелами и в правильном регистре
                curString = ""
                For j = 0 To arrayWords.GetUpperBound(0)
                    curString += arrayWords(j).Word
                Next
                'если текущая строка - начало блока HTML / Wrap - устанавливаем isInText = True
                If curString.StartsWith("HTML") OrElse curString.StartsWith("Wrap") Then
                    If isInExec <> CodeTextBox.ExecBlockEnum.NO_EXEC Then
                        mScript.LAST_ERROR = "Недопустимы внутренние блоки Wrap/HTML внури тэга <exec>."
                        Return Nothing
                    End If
                    If arrayWords(0).wordType = CodeTextBox.EditWordTypeEnum.W_HTML Then
                        isInText = CodeTextBox.TextBlockEnum.TEXT_HTML
                        isInExec = CodeTextBox.ExecBlockEnum.NO_EXEC
                        textStartLine = i + 1
                    ElseIf arrayWords(0).wordType = CodeTextBox.EditWordTypeEnum.W_WRAP Then
                        isInText = CodeTextBox.TextBlockEnum.TEXT_WRAP
                        isInExec = CodeTextBox.ExecBlockEnum.NO_EXEC
                        textStartLine = i + 1
                    End If
                End If

                curString = CodeData(i).StartingSpaces + curString + CodeData(i).Comments 'готовая строка со всеми правками

                'если строка заканцивается на _ - устанавливаем isPrevStringDissociation = True,
                If arrayWords(arrayWords.GetUpperBound(0)).wordType = CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                    isPrevStringDissociation = True 'в новом витке цикла будет обозначать, что предыдущая строка закончилась на _
                    If arrayWordsUBound > 0 Then
                        If IsNothing(arrayWordsCopy) OrElse arrayWordsCopy.GetUpperBound(0) = -1 Then 'arrayWordsCopy пуст. Вносим в него arrayWords
                            ReDim arrayWordsCopy(arrayWordsUBound - 1)
                            Array.ConstrainedCopy(arrayWords, 0, arrayWordsCopy, 0, arrayWordsCopy.GetUpperBound(0) + 1)
                        Else 'arrayWordsCopy уже частично заполнен. добавляем в него arrayWords
                            ReDim Preserve arrayWordsCopy(arrayWordsCopy.GetUpperBound(0) + arrayWordsUBound)
                            Array.ConstrainedCopy(arrayWords, 0, arrayWordsCopy, arrayWordsCopy.GetUpperBound(0) - arrayWordsUBound + 1, arrayWordsUBound)
                        End If
                    End If
                Else
                    If isPrevStringDissociation Then
                        'если предыдущая строка заканчивалась на _, то собираем до конца всю разбитую строку и проверяем ее синтаксис
                        ReDim Preserve arrayWordsCopy(arrayWordsCopy.GetUpperBound(0) + arrayWordsUBound + 1)
                        Array.ConstrainedCopy(arrayWords, 0, arrayWordsCopy, arrayWordsCopy.GetUpperBound(0) - arrayWordsUBound, arrayWordsUBound + 1)
                        If AnalizeArrayWords(CodeData, arrayWordsCopy, False, i) = -1 Then
                            Return Nothing
                        End If
                    End If
                    isPrevStringDissociation = False
                    If IsNothing(arrayWordsCopy) = False Then Erase arrayWordsCopy
                End If
                i += 1
            End While
            Return CodeData
        End Function

        Private Function StripComments(ByRef curString As String, ByRef strComment As String, ByRef strStartingSpaces As String) As Integer
            'сохраняем начальные пробелы и убираем их из curString
            Dim j As Integer
            Dim commentPos As Integer
            Dim brPos As Integer
            strStartingSpaces = curString
            curString = curString.TrimStart(" "c, Convert.ToChar(vbTab)) 'убираем начальные пробелы и символы табуляции из curString
            If strStartingSpaces.Length <> curString.Length Then
                'пробелы есть. Сохраняем их в strStartingSpaces
                strStartingSpaces = strStartingSpaces.Substring(0, strStartingSpaces.Length - curString.Length)
            Else
                'пробелов нет
                strStartingSpaces = ""
            End If
            'сохраняем комментарии и убираем их из curString
            strComment = ""
            commentPos = curString.IndexOf("//")
            If commentPos > -1 Then
                'комментарии могут быть (а могут быть и в составе обычных строк 'Hello, W//orld!'
                j = -1
                brPos = -2
                Do
                    'ищем // и '
                    If commentPos < j Then
                        commentPos = curString.IndexOf("//", j + 1)
                        If commentPos = -1 Then commentPos = Integer.MaxValue
                    End If
                    If brPos < j Then
                        brPos = curString.IndexOf("'", j + 1)
                        If brPos = -1 Then brPos = Integer.MaxValue
                    End If
                    j = Math.Min(commentPos, brPos) 'получаем что раньше встречается
                    If j = Integer.MaxValue Then Exit Do 'нет комментария - заканчиваем
                    If j = brPos Then
                        'Найдена строка '...'.
qbQuoteSearching:
                        j = curString.IndexOf("'", j + 1) 'ищем закрывающий "'"
                        If j = -1 Then
                            'закрывающий "'" не найден - ошибка
                            mScript.LAST_ERROR = "Не найдена закрывающая кавычка"
                            Return -1
                        End If
                        'Обрабатываем экранированную кавычку /'
                        If curString.Chars(j - 1) = "/"c Then GoTo qbQuoteSearching
                    Else
                        'комментарий раньше. Сохраняем его и выходим из цикла
                        strComment = " " + curString.Substring(j)
                        If j = 0 Then
                            curString = ""
                        Else
                            curString = curString.Remove(j).TrimEnd
                        End If
                        Exit Do
                    End If
                Loop
            End If
            Return 1
        End Function

        Private arrOperators As New SortedList(Of String, CodeTextBox.EditWordTypeEnum)(StringComparer.CurrentCultureIgnoreCase) 'см. FillOperators
        Private arrOperatorsKeys As New SortedList(Of String, String)(StringComparer.CurrentCultureIgnoreCase) 'key = value

        ''' <summary>
        ''' Функция получает массив arrayWords() из PrepareText  с частично полученными значениями слов и получает значения оставшихся слов, приводит их в правильный регистр
        ''' Предназначена для уже готовой строки (а не такой, которую юзер еще пишет)
        ''' </summary>
        ''' <param name="arrayWords">массив arrayWords() из PrepareText  с частично полученными значениями слов</param>
        ''' <param name="isStringDissociation">заканчивается ли предыдущая строка на _</param>
        ''' <param name="curLine">текущая линия</param>
        Private Function AnalizeArrayWords(ByRef CodeData() As CodeTextBox.CodeDataType, ByRef arrayWords() As CodeTextBox.EditWordType, ByRef isStringDissociation As Boolean, ByVal curLine As Integer) As Integer
            If IsNothing(arrayWords) OrElse arrayWords.GetUpperBound(0) = -1 Then Return 1

            Dim curWord As String 'для текущего слова
            Dim classId As Integer 'для получения Id класса текущего слова (если это класс, функция или свойство)
            Dim strLower As String 'для текущего слова в нижнем регистре
            Dim fp As New MatewScript.funcAndPropHashType

            For i As Integer = 0 To arrayWords.GetUpperBound(0) 'перебираем все слова по одному
                If arrayWords(i).wordType <> CodeTextBox.EditWordTypeEnum.W_NOTHING Then Continue For 'слово уже известно - продолжаем

                curWord = arrayWords(i).Word
                'Это просто число?
                If Double.TryParse(curWord, System.Globalization.NumberStyles.Any, provider_points, Nothing) Then
                    arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER
                    Continue For
                End If
                'Это один из операторов (сохраненный в хэше операторов)?
                If arrOperators.TryGetValue(curWord, arrayWords(i).wordType) Then
                    arrayWords(i).Word = arrOperatorsKeys(curWord)
                    If arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC Then arrayWords(i).Word = " " & arrayWords(i).Word & " "
                    'If i = 1 AndAlso arrayWords(i).wordType = codetextbox.EditWordTypeEnum.W_BLOCK_IF AndAlso String.Compare(arrayWords(i - 1).Word.Trim, "Case", True) = 0 _
                    'AndAlso String.Compare(arrayWords(i).Word.Trim, "Else", True) = 0 Then
                    '    'это Case Else, Else относится не к W_BLOCK_IF, а к W_SWITCH
                    '    arrayWords(i).wordType = codetextbox.EditWordTypeEnum.W_SWITCH
                    'End If

                    Continue For
                End If
                'Это метка?
                If curWord.EndsWith(":") Then
                    If arrayWords.GetUpperBound(0) = 0 Then
                        arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_MARK
                        Continue For
                    Else
                        mScript.LAST_ERROR = "Не допускаются любые другие операторы в строке рядом с меткой."
                        Return -1
                    End If
                End If

                If i < arrayWords.GetUpperBound(0) Then
                    strLower = curWord.ToLower 'слово в нижнем регистре
                    'Это строка вида Class.Element / Class[...].Element?
                    If arrayWords(i + 1).wordType = CodeTextBox.EditWordTypeEnum.W_POINT Then
                        'Class.Element
                        If strLower = "var" OrElse strLower = "v" Then
                            'Var.varName
                            'сохраняем результат слова
                            arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS
                            arrayWords(i).Word = IIf(strLower = "v", "V", "Var")
                            arrayWords(i).classId = -1 'для псевдокласса Var classId = -1
                            arrayWords(i + 1).classId = -1 'ставим к какому классу относится точка 
                            'устанавливаем, что varName - переменная
                            If i + 1 < arrayWords.GetUpperBound(0) Then
                                If arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                    arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_VARIABLE
                                    arrayWords(i + 2).classId = -1
                                    Continue For
                                Else
                                    If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                        arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_VARIABLE
                                        arrayWords(i + 3).classId = -1
                                        Continue For
                                    Else
                                        mScript.LAST_ERROR = "После " + arrayWords(i).Word + ". стоит не имя переменной."
                                        Return -1
                                    End If
                                End If
                            Else
                                mScript.LAST_ERROR = "После " + arrayWords(i).Word + ". не стоит имя переменной."
                                Return -1
                            End If
                        Else
                            If mScript.mainClassHash.TryGetValue(curWord, classId) Then
                                'имя класса найдено в хэше mainClassHash
                                'Class.Element. Сохраняем значения слова класса
                                arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = classId
                                arrayWords(i + 1).classId = classId  'ставим к какому классу относится точка 
                                'Ставим имя класса в правильном регистре
                                For Each strName As String In mScript.mainClass(classId).Names
                                    If strLower = strName.ToLower Then
                                        arrayWords(i).Word = strName
                                        Exit For
                                    End If
                                Next
                                'определяем функцию / свойство найденного класса, следующего за ним через точку (Class.XXX)
                                If i + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                        arrayWords(i + 2).classId = classId
                                        If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(i + 2).Word, fp) Then
                                            'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                            arrayWords(i + 2).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, CodeTextBox.EditWordTypeEnum.W_FUNCTION, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                                            arrayWords(i + 2).Word = fp.elementName
                                        Else
                                            'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                            arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                            'и свойство - не найдя в структуре проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                            arrayWords(i + 3).classId = classId
                                            If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(i + 3).Word, fp) Then
                                                'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                                arrayWords(i + 3).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, CodeTextBox.EditWordTypeEnum.W_FUNCTION, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                                                arrayWords(i + 3).Word = fp.elementName
                                            Else
                                                'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                                arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                                'и свойство - не найдя в структуре проверить точно невозможно
                                            End If
                                            Continue For
                                        Else
                                            mScript.LAST_ERROR = "После " + arrayWords(i).Word + ". стоит не название функции или свойства."
                                            Return -1
                                        End If
                                    End If
                                Else
                                    mScript.LAST_ERROR = "После " + arrayWords(i).Word + ". не стоит название функции или свойства."
                                    Return -1
                                End If
                            Else
                                'имя класса НЕ найдено в хэше mainClassHash (возможно, создан динамически с New Class)
                                arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = -2 'неизвестный класс (возможно, создан динамически с New Class)
                                arrayWords(i + 1).classId = -2 'ставим к какому классу относится точка 
                                If i + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                        'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                        arrayWords(i + 2).classId = -2 'неизвестный класс
                                        If arrayWords.GetUpperBound(0) > i + 2 AndAlso (arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                            arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY 'это свойство
                                        Else
                                            arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                            'и свойство - не зная класса проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                            'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                            arrayWords(i + 3).classId = -2 'неизвестный класс
                                            If arrayWords.GetUpperBound(0) > i + 3 AndAlso (arrayWords(i + 4).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse arrayWords(i + 4).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                                arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY 'это свойство
                                            Else
                                                arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                                'и свойство - не зная класса проверить точно невозможно
                                            End If
                                            Continue For
                                        Else
                                            mScript.LAST_ERROR = "После " + arrayWords(i).Word + ". стоит не название функции или свойства."
                                            Return -1
                                        End If
                                    End If
                                Else
                                    mScript.LAST_ERROR = "После " + arrayWords(i).Word + ". не стоит название функции или свойства."
                                    Return -1
                                End If
                            End If
                        End If
                    ElseIf arrayWords(i + 1).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                        'myVar[...], Code[...].X

                        If (strLower = "var" OrElse strLower = "v") Then
                            mScript.LAST_ERROR = "Неверная запись переменной-массива. Должно быть " + curWord + ".myVar[x]"
                            Return -1
                        End If

                        'ищем закрывающую ], за которой может стоять точка - и тогда за ней - функция или свойство, или что-либо другое (тогда это перменная)
                        Dim qbBalance As Integer = 1, qbClosePos As Integer
                        For j As Integer = i + 2 To arrayWords.GetUpperBound(0)
                            If arrayWords(j).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                qbBalance -= 1
                                If qbBalance = 0 Then
                                    qbClosePos = j
                                    Exit For
                                End If
                            ElseIf arrayWords(j).wordType = CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                                qbBalance += 1
                            End If
                        Next

                        If qbClosePos = arrayWords.GetUpperBound(0) OrElse arrayWords(qbClosePos + 1).wordType <> CodeTextBox.EditWordTypeEnum.W_POINT Then
                            'это переменная (за [] нет точки)
                            arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_VARIABLE
                            arrayWords(i).classId = -1
                            Continue For
                        Else
                            'это функция или свойство
                            If mScript.mainClassHash.TryGetValue(curWord, classId) Then
                                'имя класса найдено в хэше mainClassHash
                                arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = classId
                                arrayWords(qbClosePos + 1).classId = classId  'ставим к какому классу относится точка 
                                'Ставим имя класса в правильном регистре
                                For Each strName As String In mScript.mainClass(classId).Names
                                    If strLower = strName.ToLower Then
                                        arrayWords(i).Word = strName
                                        Exit For
                                    End If
                                Next
                                'определяем функцию / свойство найденного класса, следующего за ним за закрывающей ] (Class[...].XXX)
                                If qbClosePos + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(qbClosePos + 2).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                        arrayWords(qbClosePos + 2).classId = classId
                                        If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(qbClosePos + 2).Word, fp) Then
                                            'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                            arrayWords(qbClosePos + 2).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, CodeTextBox.EditWordTypeEnum.W_FUNCTION, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                                            arrayWords(qbClosePos + 2).Word = fp.elementName
                                        Else
                                            'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                            arrayWords(qbClosePos + 2).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                            'и свойство - не найдя в структуре проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                            arrayWords(qbClosePos + 3).classId = classId
                                            If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(qbClosePos + 3).Word, fp) Then
                                                'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                                arrayWords(qbClosePos + 3).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, CodeTextBox.EditWordTypeEnum.W_FUNCTION, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                                                arrayWords(qbClosePos + 3).Word = fp.elementName
                                            Else
                                                'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                                arrayWords(qbClosePos + 3).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                                'и свойство - не найдя в структуре проверить точно невозможно
                                            End If
                                            Continue For
                                        Else
                                            mScript.LAST_ERROR = "После " + arrayWords(i).Word + "[...]. стоит не название функции или свойства."
                                            Return -1
                                        End If
                                    End If
                                Else
                                    mScript.LAST_ERROR = "После " + arrayWords(i).Word + "[...]. не стоит название функции или свойства."
                                    Return -1
                                End If
                            Else
                                'имя класса НЕ найдено в хэше mainClassHash (возможно, создан динамически с New Class)
                                arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = -2 'неизвестный класс (возможно, создан динамически с New Class)
                                arrayWords(qbClosePos + 1).classId = -2  'ставим к какому классу относится точка 
                                If qbClosePos + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(qbClosePos + 2).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                        'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                        arrayWords(qbClosePos + 2).classId = -2 'неизвестный класс
                                        If arrayWords.GetUpperBound(0) > qbClosePos + 2 AndAlso (arrayWords(qbClosePos + 3).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse arrayWords(qbClosePos + 3).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                            arrayWords(qbClosePos + 2).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY 'это свойство
                                        Else
                                            arrayWords(qbClosePos + 2).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                            'и свойство - не зная класса проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                                            'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                            arrayWords(qbClosePos + 3).classId = -2 'неизвестный класс
                                            If arrayWords.GetUpperBound(0) > qbClosePos + 3 AndAlso (arrayWords(qbClosePos + 4).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL OrElse arrayWords(qbClosePos + 4).wordType = CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                                arrayWords(qbClosePos + 3).wordType = CodeTextBox.EditWordTypeEnum.W_PROPERTY 'это свойство
                                            Else
                                                arrayWords(qbClosePos + 3).wordType = CodeTextBox.EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                                'и свойство - не зная класса проверить точно невозможно
                                            End If
                                            Continue For
                                        Else
                                            mScript.LAST_ERROR = "После " + arrayWords(i).Word + ". стоит не название функции или свойства."
                                            Return -1
                                        End If
                                    End If
                                Else
                                    mScript.LAST_ERROR = "После " + arrayWords(i).Word + ". не стоит название функции или свойства."
                                    Return -1
                                End If
                            End If
                        End If
                    End If
                End If
                'неизвестное слово. Это функция или свойство без имени класса, или переменная
                'ищем в функциях и свойствах
                For j As Integer = 0 To mScript.mainClass.GetUpperBound(0)
                    If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(j).Names(0) + "_" + curWord, fp) Then
                        arrayWords(i).classId = j
                        arrayWords(i).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, CodeTextBox.EditWordTypeEnum.W_FUNCTION, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                        arrayWords(i).Word = fp.elementName
                        Exit For
                    End If
                Next

                If arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_NOTHING Then
                    'это переменная
                    arrayWords(i).wordType = CodeTextBox.EditWordTypeEnum.W_VARIABLE
                    arrayWords(i).classId = -1
                End If
            Next

            'Расставляем правильно пробелы и проверяем на ошибки синтаксиса
            Dim isFirstWord As Boolean, isLastWord As Boolean 'текущее слово первое или последнее? (напр., после Then слово считается первым)
            'перебираем в цикле все слова (кроме последнего)
            For i = 0 To arrayWords.GetUpperBound(0) - 1
                If i = 0 OrElse arrayWords(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION OrElse arrayWords(i - 1).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA OrElse arrayWords(i - 1).Word = "Then " Then
                    'если перед текущим словом стоит Then или ; или <exec> - слово первое
                    If i = 0 And isStringDissociation = True Then
                        If IsNothing(CodeData(curLine - 1).Code) = False AndAlso (IsNothing(CodeData(curLine).Code) = True OrElse (CodeData(curLine).Code.GetUpperBound(0) > 0 AndAlso CodeData(curLine).Code(CodeData(curLine).Code.GetUpperBound(0) - 1).Word = "Then ")) Then
                            isFirstWord = True
                        Else
                            isFirstWord = False
                        End If
                    Else
                        isFirstWord = True
                    End If
                Else
                    isFirstWord = False
                End If
                If arrayWords(i + 1).wordType = CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION OrElse arrayWords(i + 1).Word = "Then" OrElse arrayWords(i + 1).wordType = CodeTextBox.EditWordTypeEnum.W_HTML_DATA Then
                    'если после текущего слова стоит Then или ; - слово последнее
                    isLastWord = True
                Else
                    isLastWord = False
                End If
                'пербираем слова и, в зависимости от их значения и значения слова, следующего за ним, ставим где надо пробел (и проверяем на ошибки синтаксиса)
                Select Case arrayWords(i).wordType 'простое число, строка или True/False
                    Case CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL, CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER, CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается с простого значения."
                            Return -1
                        End If
                        Select Case arrayWords(i + 1).wordType
                            Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE, CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC, CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, CodeTextBox.EditWordTypeEnum.W_COMMA, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION, CodeTextBox.EditWordTypeEnum.W_HTML_DATA
                                '= <> And + & ) ] , ;
                                'ничего
                            Case CodeTextBox.EditWordTypeEnum.W_BLOCK_IF 'Then
                                If arrayWords(i + 1).Word.Trim = "Then" Then
                                    If arrayWords(i + 1).Word.First <> " "c Then arrayWords(i).Word &= " "
                                Else
                                    mScript.LAST_ERROR = "Неверная запись блока If ... Then."
                                    Return -1
                                End If
                            Case CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR  'To or Step
                                If arrayWords(i + 1).Word.Trim = "To" OrElse arrayWords(i + 1).Word.Trim = "Step" Then
                                    If arrayWords(i + 1).Word.First <> " "c Then arrayWords(i).Word &= " "
                                Else
                                    mScript.LAST_ERROR = "Неверная запись блока For ... Next."
                                    Return -1
                                End If
                            Case CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case CodeTextBox.EditWordTypeEnum.W_VARIABLE 'переменная
                        Select Case arrayWords(i + 1).wordType
                            Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE, CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC, CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, CodeTextBox.EditWordTypeEnum.W_COMMA, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN, _
                                CodeTextBox.EditWordTypeEnum.W_HTML_DATA
                                '= <> And + & ) ] , ; [
                                'ничего
                            Case CodeTextBox.EditWordTypeEnum.W_BLOCK_IF 'Then
                                If arrayWords(i + 1).Word = "Then" Then
                                    arrayWords(i).Word &= " "
                                Else
                                    mScript.LAST_ERROR = "Неверная запись блока If ... Then."
                                    Return -1
                                End If
                            Case CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR  'To
                                If arrayWords(i + 1).Word = "To" Then
                                    arrayWords(i).Word &= " "
                                Else
                                    mScript.LAST_ERROR = "Неверная запись блока For ... Next."
                                    Return -1
                                End If
                            Case CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH, CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, _
                        CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE, CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL
                        '+ & <> =
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается с оператора."
                            Return -1
                        End If
                        Select Case arrayWords(i + 1).wordType
                            Case CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL, CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER, _
                                CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING, CodeTextBox.EditWordTypeEnum.W_VARIABLE, _
                                CodeTextBox.EditWordTypeEnum.W_CLASS, CodeTextBox.EditWordTypeEnum.W_PROPERTY, _
                                CodeTextBox.EditWordTypeEnum.W_FUNCTION, CodeTextBox.EditWordTypeEnum.W_PARAM, _
                                CodeTextBox.EditWordTypeEnum.W_PARAM_COUNT, CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION, CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                                CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING
                                'True 5 'xx' myVar Code Prop Func Param ParamCount ( _ $ #
                                'все ок
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается с оператора."
                            Return -1
                        End If
                        Select Case arrayWords(i + 1).wordType
                            Case CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL, CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER, _
                                CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING, CodeTextBox.EditWordTypeEnum.W_VARIABLE, _
                                CodeTextBox.EditWordTypeEnum.W_CLASS, CodeTextBox.EditWordTypeEnum.W_PROPERTY, _
                                CodeTextBox.EditWordTypeEnum.W_FUNCTION, CodeTextBox.EditWordTypeEnum.W_PARAM, _
                                CodeTextBox.EditWordTypeEnum.W_PARAM_COUNT, CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION, CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                                CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING
                                'True 5 'xx' myVar Code Prop Func Param ParamCount ( _ $ #
                                'arrayWords(i).Word = " " & arrayWords(i).Word & " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается со скобки."
                            Return -1
                        End If
                        ')
                        Select Case arrayWords(i + 1).wordType
                            Case CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION, CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC, CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE, CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, _
                                CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                                CodeTextBox.EditWordTypeEnum.W_HTML_DATA, CodeTextBox.EditWordTypeEnum.W_COMMA
                                'ok
                            Case Else
                                If arrayWords(i + 1).Word.First <> " " Then arrayWords(i).Word &= " "
                        End Select
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        '( [ ]
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается со скобки."
                            Return -1
                        End If
                    Case CodeTextBox.EditWordTypeEnum.W_PROPERTY, CodeTextBox.EditWordTypeEnum.W_FUNCTION
                        'Func
                        Select Case arrayWords(i + 1).wordType
                            Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE, CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC, CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION, CodeTextBox.EditWordTypeEnum.W_HTML_DATA, CodeTextBox.EditWordTypeEnum.W_COMMA
                                'ok
                            Case CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR, CodeTextBox.EditWordTypeEnum.W_BLOCK_IF, _
                                CodeTextBox.EditWordTypeEnum.W_CLASS, CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                                CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING, CodeTextBox.EditWordTypeEnum.W_FUNCTION, _
                                CodeTextBox.EditWordTypeEnum.W_PARAM, CodeTextBox.EditWordTypeEnum.W_PARAM_COUNT, _
                                CodeTextBox.EditWordTypeEnum.W_PROPERTY, CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL, _
                                CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER, CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION, CodeTextBox.EditWordTypeEnum.W_VARIABLE
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case CodeTextBox.EditWordTypeEnum.W_PARAM
                        'Param[x]
                        Select Case arrayWords(i + 1).wordType
                            Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE, CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC, CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN, _
                                CodeTextBox.EditWordTypeEnum.W_HTML_DATA, CodeTextBox.EditWordTypeEnum.W_COMMA
                                'ok
                            Case CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR, CodeTextBox.EditWordTypeEnum.W_BLOCK_IF, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case CodeTextBox.EditWordTypeEnum.W_PARAM_COUNT
                        'ParamCount
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается с ParamCount."
                            Return -1
                        End If
                        Select Case arrayWords(i + 1).wordType
                            Case CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE, CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC, CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH, _
                                CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_CONSOLIDATION, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN, _
                                CodeTextBox.EditWordTypeEnum.W_HTML_DATA, CodeTextBox.EditWordTypeEnum.W_COMMA
                                'ok
                            Case CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR, CodeTextBox.EditWordTypeEnum.W_BLOCK_IF, _
                                CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case CodeTextBox.EditWordTypeEnum.W_BREAK, CodeTextBox.EditWordTypeEnum.W_CONTINUE, CodeTextBox.EditWordTypeEnum.W_EXIT, _
                        CodeTextBox.EditWordTypeEnum.W_JUMP, CodeTextBox.EditWordTypeEnum.W_MARK, _
                        CodeTextBox.EditWordTypeEnum.W_RETURN, CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT, _
                        CodeTextBox.EditWordTypeEnum.W_GLOBAL
                        If isFirstWord Then
                            If isLastWord = False Then
                                If arrayWords(i + 1).Word.First <> " " AndAlso arrayWords(i).Word <> "?" Then arrayWords(i).Word &= " "
                            End If
                        Else
                            mScript.LAST_ERROR = "Неправильный синтаксис."
                            Return -1
                        End If
                    Case CodeTextBox.EditWordTypeEnum.W_SWITCH, CodeTextBox.EditWordTypeEnum.W_WRAP, CodeTextBox.EditWordTypeEnum.W_HTML, _
                        CodeTextBox.EditWordTypeEnum.W_CYCLE_END, CodeTextBox.EditWordTypeEnum.W_BLOCK_DOWHILE, _
                        CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION, CodeTextBox.EditWordTypeEnum.W_BLOCK_NEWCLASS, _
                        CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR, CodeTextBox.EditWordTypeEnum.W_BLOCK_IF, CodeTextBox.EditWordTypeEnum.W_REM_CLASS
                        'Select Case Switch Wrap HTML Append End ...
                        If arrayWords(i + 1).Word.First <> " " Then arrayWords(i).Word &= " "
                        'Case codetextbox.EditWordTypeEnum.W_COMMA
                End Select
            Next

            'проверка последнего символа
            If arrayWords(arrayWords.GetUpperBound(0)).wordType <> CodeTextBox.EditWordTypeEnum.W_STRINGS_DISSOCIATION Then 'если последний символ _ , то слово на самом деле не последнее
                Select Case arrayWords(arrayWords.GetUpperBound(0)).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_CLASS, CodeTextBox.EditWordTypeEnum.W_COMMA, CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_NUMBER, CodeTextBox.EditWordTypeEnum.W_CONVERT_TO_STRING, CodeTextBox.EditWordTypeEnum.W_JUMP, CodeTextBox.EditWordTypeEnum.W_OPERATOR_COMPARE, CodeTextBox.EditWordTypeEnum.W_OPERATOR_EQUAL, CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC, CodeTextBox.EditWordTypeEnum.W_OPERATOR_MATH, CodeTextBox.EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_OPEN, CodeTextBox.EditWordTypeEnum.W_POINT, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_OPEN, CodeTextBox.EditWordTypeEnum.W_REM_CLASS
                        mScript.LAST_ERROR = "Неправильный синтаксис."
                        Return -1
                End Select
            End If
            If arrayWords.GetUpperBound(0) = 0 And isStringDissociation = False Then
                'когда в строке только одно слово
                Select Case arrayWords(0).wordType
                    Case CodeTextBox.EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, CodeTextBox.EditWordTypeEnum.W_PARAM_COUNT, CodeTextBox.EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL, CodeTextBox.EditWordTypeEnum.W_SIMPLE_NUMBER, CodeTextBox.EditWordTypeEnum.W_SIMPLE_STRING
                        mScript.LAST_ERROR = "Неправильный синтаксис."
                        Return -1
                End Select
            End If
            Return 1
        End Function

        Public Sub FillOperators()
            arrOperators.Clear()

            arrOperators.Add("True", CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL)
            arrOperators.Add("False", CodeTextBox.EditWordTypeEnum.W_SIMPLE_BOOL)
            arrOperators.Add("And", CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC)
            arrOperators.Add("Or", CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC)
            arrOperators.Add("Xor", CodeTextBox.EditWordTypeEnum.W_OPERATOR_LOGIC)
            arrOperators.Add("ParamCount", CodeTextBox.EditWordTypeEnum.W_PARAM_COUNT)
            arrOperators.Add("Param", CodeTextBox.EditWordTypeEnum.W_PARAM)
            arrOperators.Add("Exit", CodeTextBox.EditWordTypeEnum.W_EXIT)
            arrOperators.Add("Select", CodeTextBox.EditWordTypeEnum.W_SWITCH)
            arrOperators.Add("Switch", CodeTextBox.EditWordTypeEnum.W_SWITCH)
            arrOperators.Add("Case", CodeTextBox.EditWordTypeEnum.W_SWITCH)
            arrOperators.Add("End", CodeTextBox.EditWordTypeEnum.W_CYCLE_END)
            arrOperators.Add("Wrap", CodeTextBox.EditWordTypeEnum.W_WRAP)
            arrOperators.Add("HTML", CodeTextBox.EditWordTypeEnum.W_HTML)
            arrOperators.Add("Append", CodeTextBox.EditWordTypeEnum.W_HTML)
            arrOperators.Add("Do", CodeTextBox.EditWordTypeEnum.W_BLOCK_DOWHILE)
            arrOperators.Add("While", CodeTextBox.EditWordTypeEnum.W_BLOCK_DOWHILE)
            arrOperators.Add("Loop", CodeTextBox.EditWordTypeEnum.W_BLOCK_DOWHILE)
            arrOperators.Add("Function", CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION)
            arrOperators.Add("New", CodeTextBox.EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Class", CodeTextBox.EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Name", CodeTextBox.EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Prop", CodeTextBox.EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Func", CodeTextBox.EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Rem", CodeTextBox.EditWordTypeEnum.W_REM_CLASS)
            arrOperators.Add("For", CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR)
            arrOperators.Add("To", CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR)
            arrOperators.Add("Step", CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR)
            arrOperators.Add("Next", CodeTextBox.EditWordTypeEnum.W_BLOCK_FOR)
            arrOperators.Add("If", CodeTextBox.EditWordTypeEnum.W_BLOCK_IF)
            arrOperators.Add("Then", CodeTextBox.EditWordTypeEnum.W_BLOCK_IF)
            arrOperators.Add("Else", CodeTextBox.EditWordTypeEnum.W_BLOCK_IF)
            arrOperators.Add("ElseIf", CodeTextBox.EditWordTypeEnum.W_BLOCK_IF)
            arrOperators.Add("Jump", CodeTextBox.EditWordTypeEnum.W_JUMP)
            arrOperators.Add("Global", CodeTextBox.EditWordTypeEnum.W_GLOBAL)
            arrOperators.Add("Return", CodeTextBox.EditWordTypeEnum.W_RETURN)
            arrOperators.Add("Break", CodeTextBox.EditWordTypeEnum.W_BREAK)
            arrOperators.Add("Continue", CodeTextBox.EditWordTypeEnum.W_CONTINUE)
            arrOperators.Add("Event", CodeTextBox.EditWordTypeEnum.W_BLOCK_EVENT)
            arrOperators.Add("?", CodeTextBox.EditWordTypeEnum.W_RETURN)

            For Each strKey In arrOperators.Keys
                arrOperatorsKeys.Add(strKey, strKey)
            Next
        End Sub
    End Class
#End Region

    Public Sub New()
        eventRouter = New EventRouterClass(Me)
    End Sub
End Class
