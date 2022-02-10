Imports System.ComponentModel
Imports System.Deployment
Imports System.IO
Imports System.Security.Cryptography
'Imports System.ServiceProcess
Imports System.Globalization

Public Class CodeTextBox
#Region "Enums & Structures"
    Public Enum EditWordTypeEnum
        W_NOTHING = 0
        W_SIMPLE_NUMBER = 1
        W_SIMPLE_STRING = 2
        W_SIMPLE_BOOL = 3
        W_VARIABLE = 4
        W_CLASS = 5 'C, Code ...
        W_OPERATOR_MATH = 6 '+ - * ^ / \
        W_OPERATOR_STRINGS_MERGER = 7 '&
        W_OPERATOR_COMPARE = 8 '!= < > <> >= <=
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
        W_CYCLE_END = 28
        W_BREAK = 29
        W_CONTINUE = 30
        W_OVAL_BRACKET_OPEN = 31 '(
        W_OVAL_BRACKET_CLOSE = 32 ')
        W_QUAD_BRACKET_OPEN = 33 '[
        W_QUAD_BRACKET_CLOSE = 34 ']
        W_POINT = 35 '.
        W_COMMA = 36 ',
        W_OPERATOR_EQUAL = 37 '=
        W_STRINGS_CONSOLIDATION = 38 ' ;
        W_STRINGS_DISSOCIATION = 39 ' _
        W_CONVERT_TO_STRING = 40 '$
        W_CONVERT_TO_NUMBER = 41 ' #
        W_GLOBAL = 42
        W_COMMENTS = 98
        W_ERROR = 99
        W_HTML_TAG = 100
        W_HTML_DATA = 101
    End Enum

    Public Structure EditWordType
        Public Word As String
        Public wordType As EditWordTypeEnum
        Public classId As Integer
    End Structure

    Public Enum TextBlockEnum
        NO_TEXT_BLOCK = 0
        TEXT_HTML = 1
        TEXT_WRAP = 2
    End Enum

    Public Enum ExecBlockEnum
        NO_EXEC = 0
        EXEC_SINGLE = 1
        EXEC_MULTILINE = 2
    End Enum

    Public Structure CodeDataType
        Public StartingSpaces As String
        Public Comments As String
        Public Code() As EditWordType
    End Structure

    Public Structure StylePresetType
        Public style_Color As Drawing.Color
        Public font_style As FontStyle
    End Structure

    Private Structure CheckCodeStructure
        Public AllowedNextWordTypes() As EditWordTypeEnum
        Public bannedNextWords() As String
        Public canBeFirst As Boolean
        Public canBeLast As Boolean
        Public shouldBeFirst As Boolean
        Public shouldBeLast As Boolean
        Public previousWordShouldBe As String
        Public canBeFinalBlockWord As Boolean 'End If
    End Structure
#End Region

#Region "Declarations"

    'для сохранения номера и кода линии, которую сейчас буду редактировать
    Public Declare Function GetScrollPos Lib "User32" (ByVal hWnd As IntPtr, ByVal nBar As Integer) As Integer
    Public Declare Function SetScrollPos Lib "User32" (ByVal hWnd As IntPtr, ByVal nBar As Integer, ByVal nPos As Integer, ByVal bRedraw As Boolean) As Integer
    Private Const SB_HORZ As Integer = &H0
    Private Const SB_VERT As Integer = &H1

    Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As _
       Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Const WM_SetRedraw As Integer = &HB 'запрет/разрешение перерисовки
    Const EM_LINESCROLL As Integer = &HB6 'для возврата прокрутки в точности туда, где она была до подготовки текста
    'Const EM_SETSCROLLPOS As Integer = &H400 + 222
    'Const EM_GETSCROLLPOS As Integer = &H400 + 221
    Const WM_KEYDOWN = &H100
    Const WM_KEYUP = &H101
    'Public Const WM_UNDO = &H304

    Private lastKey As System.Windows.Forms.Keys = 0 'чтоб знать в событии textChanged был ли нажат Enter, Del или Backspace
    Private Const rtbMaxLineLength As Integer = 2730 'максимальновозможная длина одной строки в rtb, после которой срабатывает Wrap
    Public WithEvents helpDocument As HtmlDocument

    Private EXIT_CODE As Boolean = False 'Был ли код прерван каким-либо оператором или в случае ошибки
    Private LAST_ERROR As String = "" 'Текст последней ошибки
    Private CURRENT_LINE As Integer = 0 ' Индекс текущей строки в masCode. Нужно для отслеживании строки с ошибкой при рекурсивном вызове ExecuteCode

    Private hCurrentDocument As HtmlDocument

    Public Shadows Event TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    Public WithEvents codeBox As New RichTextBoxEx With {.Text = "//", .Dock = DockStyle.Fill, .Visible = True, .ContextMenuStrip = mnuRtb, .CausesValidation = True, .EnableAutoDragDrop = False, _
        .Font = New Font("Consolas", 11), .ScrollBars = RichTextBoxScrollBars.ForcedBoth, .WordWrap = False}
#End Region

    ''' <summary>Класс, управляющий Undo и Redo в кодбоксе</summary>
    Public Class UndoClass
        ''' <summary>максимальное число символов в типе CHAR_ENTRY, которое помещается в один item (после которого будет создана новое вхождение undo)</summary>
        Public Shared UNDO_SYMBOLS_IN_CHAR_ENTRY As Integer = 50
        ''' <summary>Сылка на кодбокс, которым управляет данный класс</summary>
        Public codeBox As RichTextBoxEx
        ''' <summary>Включить/отключить события Undo/Redo в данном кодбоксе</summary>
        Public AllowUndo As Boolean = True
        ''' <summary>Для хранения выделенного текста. Он сохраняется во время события KeyDown, а создание нового вхождения Undo - только в KeyPress, когда выделенный текст стерт</summary>
        Public lastSelection As String = ""
        ''' <summary>Хранит список символов, воспринимаемых как конец слова</summary>
        Private wordBounds As List(Of Char)
        ''' <summary>Позиция последего введенного с клавиатуры символа. Если последнее изменение было не с помощью клавиатуры, то равно -1. Надо для присоединения нового введенного символа к последнему
        ''' вхождению Undo, если его тип был CHAR_ENTRY (иначе каждое Undo будет удалять введеные данные по одной букве).</summary>
        Private lastCharPos As Integer = -1
        ''' <summary>Хранит номер первой видимой линии кодбокса (надо в некоторых случаях внутри функций PrepareText/PrepaeTextNonCompeted, поскольку там может обрабатываться текст от первой линии 
        ''' до последней, в результате номер линии в момент создания нового вхождения Undo будет некорректным)</summary>
        Public defaultFirstVisibleLine As Integer = -1
        ''' <summary>Кнопки Undo/Redo - для установлки доступности/недоступности</summary>
        Public btnUndo As ToolStripButton, btnRedo As ToolStripButton

        ''' <summary>Типы undo - ввод символа, изменение выделенного, клавиша Delete, Backspace и полное изменение текста</summary>
        Public Enum UndoTypeEnum As Byte
            CHAR_ENTRY = 0
            SELECTED_CHANGED = 1
            DELETE = 2
            BACKSPACE = 3
            'FULL_TEXT_CHANGED = 4
        End Enum

        ''' <summary>Класс, хранящий ввхождение Undo</summary>
        Public Class UndoItemClass
            ''' <summary>Новый вставляемый в кодбокс текст</summary>
            Public TextNew As String
            ''' <summary>Структура CodeData, необходима только для FULL_TEXT_CHANGED</summary>
            Public cd() As CodeDataType
            ''' <summary>Тип Undo</summary>
            Public itemType As UndoTypeEnum
            ''' <summary>Начало выделенного текста в момент создания вхождения</summary>
            Public SelectionStart As Integer
            ''' <summary>Длина выделенного текста в момент создания вхождения</summary>
            Public SelectionLength As Integer
            ''' <summary>Выделенный текст в момент создания вхождения</summary>
            Public SelectedText As String
            ''' <summary>Номер первой видимой линии</summary>
            Public FirstVisibleLine As Integer
        End Class

        ''' <summary>Список, хранящий все вхождения Undo</summary>
        Private lstUndo As New List(Of UndoItemClass)
        ''' <summary>Список, хранящий все вхождения Redo</summary>
        Private lstRedo As New List(Of UndoItemClass)
        ''' <summary>Выполнется ли в данный момент Undo (необходимо, чтобы не создавались новые вхождения undo во время выполнения старых)</summary>
        Public UndoInProcess As Boolean = False
        ''' <summary>Для сохранения исходного состояния до всех правок</summary>
        Public InitialState As UndoItemClass

        ''' <summary>Подготавливает класс к началу ввода. Текущее состояние кодбокса считается исходным</summary>
        Public Sub BeginNew()
            If Not AllowUndo Then Return
            lstUndo.Clear()
            lstRedo.Clear()
            'AppendItem(UndoTypeEnum.FULL_TEXT_CHANGED)
            InitialState = New UndoItemClass With {.cd = CopyCodeDataArray(codeBox.CodeData), .FirstVisibleLine = 0, .SelectionStart = 0, .SelectionLength = 0, .SelectedText = ""}
            lastCharPos = -1 'очищаем номер последнего введенного символа, так как тип Undo - не ввод с клавиатуры
            If Not IsNothing(btnUndo) Then btnUndo.Enabled = False
            If Not IsNothing(btnRedo) Then btnRedo.Enabled = False
        End Sub

        ''' <summary>
        ''' Добавляет новое вхождение Undo
        ''' </summary>
        ''' <param name="itemType">Тип Undo</param>
        ''' <param name="strText">Соответствующий текст (введенный символ, вставляемый текст и т. д.)</param>
        ''' <param name="selStart">Начало выделения. Пока используется только для BACKSPACE, если была комбинация Ctrl+Backspace</param>
        Public Sub AppendItem(ByVal itemType As UndoTypeEnum, Optional ByVal strText As String = "", Optional ByVal selStart As Integer = -1)
            If Not AllowUndo OrElse UndoInProcess OrElse IsNothing(codeBox) Then Return

            'получаем первую видимую линию
            Dim upperLine As Integer
            If defaultFirstVisibleLine = -1 Then
                'заранее линия не установлена. Значит получаем ту, которая прямо сейчас
                upperLine = codeBox.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = codeBox.GetLineFromCharIndex(upperLine)
            Else
                'верхня линия установлена заранее, так как прямо в этот момент она некорректная
                upperLine = defaultFirstVisibleLine
            End If

            Select Case itemType
                Case UndoTypeEnum.CHAR_ENTRY
                    'простой ввод символа с клавиатуры
                    Dim blnNewItem As Boolean = True
                    If lstUndo.Count > 0 AndAlso lstUndo.Last.SelectedText.Length = 0 AndAlso lstUndo.Last.itemType = UndoTypeEnum.CHAR_ENTRY AndAlso strText.Length = 1 AndAlso _
                        wordBounds.IndexOf(strText) = -1 AndAlso lstUndo.Last.TextNew.Length < UNDO_SYMBOLS_IN_CHAR_ENTRY AndAlso lastSelection.Length = 0 AndAlso codeBox.SelectionStart = lastCharPos + 1 Then
                        'Добавляем введенный символ к предыдущему входению Undo если:
                        'Список Undo не пустой, последнее вхождение не заменяло выделенный текст и оно такого же типа (CHAR_ENTRY), введен один символ (единичной длины), этот символ не входит 
                        'в список разделител слов, длина текста предыдущего вхождения не превышает максимально установленную (на длинные слова используем несколько вхождений Undo),
                        'в данный момент напечатанный символ не заменяет выделенный текст и вводится символ сразу же за предыдущим (а не где-то в новом месте)
                        Dim itm As UndoItemClass = lstUndo.Last
                        itm.TextNew += strText 'добавляем введенный символ в последнее вхождение Undo
                        lstUndo(lstUndo.Count - 1) = itm
                        lastCharPos = codeBox.SelectionStart 'сохраняем текущее положение каретки чтобы точно знать будет ли введен следующий символ сразу за этим (а не в новом месте)
                    Else
                        Dim itm As New UndoItemClass With {.TextNew = strText, .itemType = UndoTypeEnum.CHAR_ENTRY, .FirstVisibleLine = upperLine, _
                                                           .SelectionStart = Math.Max(codeBox.SelectionStart, 0), .SelectionLength = lastSelection.Length, .SelectedText = lastSelection}
                        lstUndo.Add(itm)
                        lastCharPos = codeBox.SelectionStart 'сохраняем текущее положение каретки чтобы точно знать будет ли введен следующий символ сразу за этим (а не в новом месте)
                    End If
                Case UndoTypeEnum.SELECTED_CHANGED
                    'замена/удаление выделенного текста с помощью Cut, Paste, вставка тэгов из меню или програмно
                    Dim itm As New UndoItemClass With {.TextNew = strText, .itemType = UndoTypeEnum.SELECTED_CHANGED, .FirstVisibleLine = upperLine, _
                                                       .SelectionStart = Math.Max(codeBox.SelectionStart, 0), .SelectionLength = lastSelection.Length, .SelectedText = lastSelection}
                    lstUndo.Add(itm)
                    lastCharPos = -1 'очищаем номер последнего введенного символа, так как тип Undo - не ввод с клавиатуры
                Case UndoTypeEnum.DELETE
                    'нажата клавиша Delete
                    Dim itm As New UndoItemClass With {.TextNew = strText, .itemType = itemType, .FirstVisibleLine = upperLine, _
                                                       .SelectionStart = Math.Max(codeBox.SelectionStart, 0), .SelectionLength = lastSelection.Length, .SelectedText = lastSelection}
                    lstUndo.Add(itm)
                    lastCharPos = -1 'очищаем номер последнего введенного символа, так как тип Undo - не ввод с клавиатуры
                Case UndoTypeEnum.BACKSPACE
                    'нажата клавиша Backspace
                    If selStart = -1 Then
                        If codeBox.SelectionLength > 0 Then
                            selStart = codeBox.SelectionStart ' - 1
                        Else
                            selStart = codeBox.SelectionStart - strText.Length
                        End If
                    End If
                    Dim itm As New UndoItemClass With {.TextNew = strText, .itemType = itemType, .FirstVisibleLine = upperLine, _
                                                       .SelectionStart = selStart, .SelectionLength = lastSelection.Length, .SelectedText = lastSelection}
                    'Dim itm As New UndoItemClass With {.TextNew = strText, .itemType = itemType, .FirstVisibleLine = upperLine, _
                    '                                   .SelectionStart = codeBox.SelectionStart - 1, .SelectionLength = lastSelection.Length, .SelectedText = lastSelection}
                    lstUndo.Add(itm)
                    lastCharPos = -1 'очищаем номер последнего введенного символа, так как тип Undo - не ввод с клавиатуры

                    'Case UndoTypeEnum.FULL_TEXT_CHANGED
                    '    'Полная замена текста. На данный момент - только первое вхождение Undo для сохданения исходного состояния
                    '    Dim itm As New UndoItemClass With {.cd = CopyCodeDataArray(codeBox.CodeData), .itemType = UndoTypeEnum.FULL_TEXT_CHANGED, .FirstVisibleLine = 0, .SelectionStart = 0, .SelectionLength = 0, .SelectedText = ""}
                    '    lstUndo.Add(itm)
                    '    lastCharPos = -1 'очищаем номер последнего введенного символа, так как тип Undo - не ввод с клавиатуры
            End Select

            If Not IsNothing(btnUndo) Then btnUndo.Enabled = True
            lstRedo.Clear() 'список Redo больше не актуален - очищаем
            If Not IsNothing(btnRedo) Then btnRedo.Enabled = False
        End Sub

        ''' <summary>Выполняет Undo</summary>
        Public Sub Undo()
            If Not AllowUndo Then Return
            If lstUndo.Count = 0 Then Return
            lastCharPos = -1 'очищаем номер последнего введенного символа, так как после Undo он будет некорректным
            UndoInProcess = True 'запрет на создание новых вхождений Undo пока не завершится данная процедура

            Dim item As UndoItemClass = lstUndo.Last
            Select Case item.itemType
                'отмена ввода символа(-лов) с клавиатуры или отмена замены выделенного текста
                Case UndoTypeEnum.CHAR_ENTRY, UndoTypeEnum.SELECTED_CHANGED
                    With codeBox
                        .Select(item.SelectionStart, item.TextNew.Length)
                        .SelectedText = item.SelectedText
                        .Select(item.SelectionStart, item.SelectedText.Length)
                    End With
                Case UndoTypeEnum.DELETE
                    'отмена удаления с помощью клавиши Delete
                    With codeBox
                        .Select(item.SelectionStart, 0)
                        .SelectedText = item.TextNew
                        .Select(item.SelectionStart, item.SelectedText.Length)
                    End With
                Case UndoTypeEnum.BACKSPACE
                    'отмена удаления с помощью клавиши Backspace
                    With codeBox
                        If item.SelectionLength = 0 Then
                            .Select(Math.Max(item.SelectionStart, 0), 0)
                            .SelectedText = item.TextNew
                            .Select(item.SelectionStart + item.TextNew.Length, item.SelectedText.Length)
                        Else
                            '.Select(item.SelectionStart + 1, 0)
                            '.SelectedText = item.SelectedText + item.TextNew
                            '.Select(item.SelectionStart + 1, item.SelectedText.Length)
                            Dim additional As Integer = 0
                            If item.TextNew.StartsWith(vbLf) Then additional = 1
                            .Select(item.SelectionStart, 0)
                            .SelectedText = item.TextNew + item.SelectedText
                            .Select(item.SelectionStart + additional, item.TextNew.Length + item.SelectedText.Length - additional)
                        End If
                    End With
                    'Case UndoTypeEnum.FULL_TEXT_CHANGED
                    '    'Восстановление исходного состояния текста
                    '    With codeBox
                    '        .LoadCodeFromCodeData(item.cd)
                    '        .Select(item.SelectionStart, item.SelectionLength)
                    '    End With
            End Select

            ScrollToLine(item.FirstVisibleLine) 'Прокручиваем кодбокс так, чтобы первая видимая линия была такой же, как и во время создания данного вхождения Undo
            'удаляем выполненное вхождение Undo и добавляем его в Redo
            lstRedo.Add(item)
            lstUndo.Remove(item)
            If Not IsNothing(btnUndo) Then
                If lstUndo.Count = 0 Then
                    btnUndo.Enabled = False
                Else
                    btnUndo.Enabled = True
                End If
            End If
            If Not IsNothing(btnRedo) Then btnRedo.Enabled = True

            UndoInProcess = False 'отменяем запрет на новые Undo
        End Sub

        ''' <summary>Выполняет Redo</summary>
        Public Sub Redo()
            If Not AllowUndo Then Return
            If lstRedo.Count = 0 Then Return
            lastCharPos = -1 'очищаем номер последнего введенного символа, так как после Redo он будет некорректным
            UndoInProcess = True 'запрет на создание новых вхождений Undo пока не завершится данная процедура

            Dim item As UndoItemClass = lstRedo.Last
            Select Case item.itemType
                'повторение ввода символа(-лов) с клавиатуры или восстановление выделенного текста
                Case UndoTypeEnum.CHAR_ENTRY, UndoTypeEnum.SELECTED_CHANGED
                    With codeBox
                        .Select(item.SelectionStart, item.SelectionLength)
                        .SelectedText = item.TextNew
                    End With
                Case UndoTypeEnum.DELETE
                    'повторение удаления с помощью клавиши Delete
                    With codeBox
                        .Select(item.SelectionStart, item.TextNew.Length)
                        .SelectedText = ""
                    End With
                Case UndoTypeEnum.BACKSPACE
                    'восстановление удаления с помощью клавиши Backspace
                    With codeBox
                        .Select(Math.Max(item.SelectionStart, 0), item.SelectionLength + item.TextNew.Length)
                        .SelectedText = ""

                        'If item.TextNew.Length = 0 Then
                        '    .Select(Math.Max(item.SelectionStart, 0), item.SelectionLength)
                        '    .SelectedText = ""

                        '    '.Select(Math.Max(item.SelectionStart, 0), 0)
                        '    '.SelectedText = item.TextNew
                        '    '.Select(item.SelectionStart + item.TextNew.Length, item.SelectedText.Length)
                        'Else
                        '    'Dim additional As Integer = 0
                        '    'If item.TextNew.StartsWith(vbLf) Then additional = 1
                        '    .Select(Math.Max(item.SelectionStart, 0), item.SelectionLength + item.TextNew.Length)
                        '    .SelectedText = ""


                        '    '.Select(item.SelectionStart, 0)
                        '    '.SelectedText = item.TextNew + item.SelectedText
                        '    '.Select(item.SelectionStart + additional, item.TextNew.Length + item.SelectedText.Length - additional)
                        'End If
                    End With
            End Select

            ScrollToLine(item.FirstVisibleLine) 'Прокручиваем кодбокс так, чтобы первая видимая линия была такой же, как и во время создания данного вхождения Undo
            'удаляем выполненное вхождение Redo и добавляем его в Undo
            lstUndo.Add(item)
            lstRedo.Remove(item)
            If Not IsNothing(btnRedo) Then
                If lstRedo.Count = 0 Then
                    btnRedo.Enabled = False
                Else
                    btnRedo.Enabled = True
                End If
            End If
            If Not IsNothing(btnUndo) Then btnUndo.Enabled = True

            UndoInProcess = False 'отменяем запрет на новые Undo
        End Sub

        ''' <summary>Восстанавливает исходное состояние текста до всех правок</summary>
        Public Sub RestoreInitalState()
            If IsNothing(InitialState) Then
                MessageBox.Show("Исходное состояние текста не было сохранено.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If
            UndoInProcess = True
            With codeBox
                .LoadCodeFromCodeData(InitialState.cd)
                .Select(InitialState.SelectionStart, InitialState.SelectionLength)
            End With
            ScrollToLine(InitialState.FirstVisibleLine)
            UndoInProcess = False
            lstUndo.Clear()
            lstRedo.Clear()
            If Not IsNothing(btnUndo) Then btnUndo.Enabled = False
            If Not IsNothing(btnRedo) Then btnRedo.Enabled = False
        End Sub

        ''' <summary>Проверяет не содержит ли последний Undo слово html/wrap</summary>
        Public Function IsLastUndoContainsTextBlock() As Boolean
            If lstUndo.Count = 0 Then Return False
            Dim strText As String = lstUndo.Last.TextNew.Trim
            If String.Compare(strText, "html", True) = 0 OrElse String.Compare(strText, "wrap", True) = 0 Then Return True
            Return False
        End Function

        ''' <summary>Прокручивает кодбокс таким образом, чтобы указанная линия стала первой видимой</summary>
        Private Sub ScrollToLine(ByVal upperLine As Integer)
            'возвращаем на место скролбар
            If codeBox.TextLength = 0 Then Return
            Dim charPos As Integer = codeBox.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})
            Dim lId As Integer = codeBox.GetLineFromCharIndex(charPos)
            SendMessage(codeBox.Handle, EM_LINESCROLL, 0, upperLine - lId)
        End Sub

        ''' <summary>Сохраняет текущую первую видимую линию для использования в дальнейшем, когда она может измениться</summary>
        Public Sub SetCurrentUpperLineAsDefault()
            If codeBox.TextLength = 0 Then Return
            Dim charPos As Integer = codeBox.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})
            defaultFirstVisibleLine = codeBox.GetLineFromCharIndex(charPos)
        End Sub

        Public Sub New(ByRef rtb As RichTextBoxEx)
            codeBox = rtb
            If IsNothing(codeBox) Then Return
            wordBounds = codeBox.wordBoundsArray.ToList
        End Sub
    End Class

    Public Class RichTextBoxEx
        Inherits RichTextBox

#Region "Declarations"
        Public wbRtbHelp As WebBrowser
        Public lstRtb As ListBox
        Private arrOperators As New SortedList(Of String, EditWordTypeEnum)(StringComparer.CurrentCultureIgnoreCase) 'см. FillOperators
        Private arrOperatorsKeys As New SortedList(Of String, String)(StringComparer.CurrentCultureIgnoreCase) 'key = value
        Private stopCharArray() As Char = {" "c, "("c, ")"c, "["c, "]"c, "."c, ","c, "'"c, "="c, "+"c, "-"c, "*"c, "^"c, "/"c, "\"c, "<"c, ">"c, "&"c, "!"c, ";"c, "#"c, "$"c, Chr(34), "?"}
        Public styleHash As New SortedList(Of EditWordTypeEnum, StylePresetType) 'цвет, выделение слов в RichTextBox в зависимости от значения
        Private checkCodeTypes As New SortedList(Of EditWordTypeEnum, CheckCodeStructure)
        Private checkCodeWords As New SortedList(Of String, CheckCodeStructure)(StringComparer.CurrentCultureIgnoreCase)
        Public wordBoundsArray() As Char = {" "c, "("c, ")"c, "["c, "]"c, "."c, ","c, "'"c, "="c, "+"c, "-"c, "*"c, "^"c, "/"c, "\"c, "<"c, ">"c, "&"c, "!"c, ";"c, "#"c, "$"c, Chr(34), Chr(10), "?"}
        Private wordBoundsArrayNoQuotes() As Char = {" "c, "("c, ")"c, "["c, "]"c, "."c, ","c, "="c, "+"c, "-"c, "*"c, "^"c, "/"c, "\"c, "<"c, ">"c, "&"c, "!"c, ";"c, "#"c, "$"c, Chr(10), "?"}
        'первые слова в строках, после которых надо увеличить начальные пробелы
        Private AddSpacesArray() As EditWordTypeEnum = {EditWordTypeEnum.W_BLOCK_DOWHILE, EditWordTypeEnum.W_BLOCK_EVENT, EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_FUNCTION, EditWordTypeEnum.W_BLOCK_IF, EditWordTypeEnum.W_BLOCK_NEWCLASS, EditWordTypeEnum.W_SWITCH}
        Public lastSelectionStartLine As Integer = 0 'чтоб знать в событии textChanged начало выделения до изменения текста
        Public lastSelectionEndLine As Integer = 0 'чтоб знать в событии textChanged длину выделения до изменения текста
        Public CodeData() As CodeDataType 'стр-ра для хранения всего кода из RichTextBox в распознанном виде
        Public rtbCanRaiseEvents As Boolean = False  'для разрешения/запрета событий в rtb
        Public prevLineContent As CodeDataType = Nothing, prevLineId As Integer = -1
        Private provider_points As New System.Globalization.NumberFormatInfo 'настройки отображения чисел с разделителем точкой (3.14 вместо 3,14)
        Public csUndo As New UndoClass(Me)
#End Region

#Region "Properties"
        ''' <summary>Считать ли содержимое текстовыми(html) данными. Например, для L.Description</summary>
        Public Property IsTextBlockByDefault As Boolean = False

        ''' <summary>
        ''' Показывать ли ошибки кода. Технически при значении True вместо PrepareText будет всегда запускаться PrepareTextNonCompleted
        ''' </summary>
        <System.ComponentModel.Browsable(True)> _
        Public Property DontShowError As Boolean = False

        ''' <summary>Раскрашивать ли слова или оставлять черными</summary>
        <System.ComponentModel.Browsable(True)> _
        Public Property CanDrawWords As Boolean = True

        Public Property CanRaiseCodeEvents() As Boolean
            Get
                Return rtbCanRaiseEvents
            End Get
            Set(ByVal value As Boolean)
                rtbCanRaiseEvents = value
            End Set
        End Property

        Private lcHelpPath As String
        <System.ComponentModel.Browsable(True)> _
        Public Property HelpPath() As String
            Get
                If IsNothing(lcHelpPath) OrElse lcHelpPath.Length = 0 Then lcHelpPath = APP_HELP_PATH
                Return lcHelpPath
            End Get
            Set(ByVal value As String)
                lcHelpPath = value
            End Set
        End Property

        Private rtbHelpFile As String
        <System.ComponentModel.Browsable(True)> _
        Public Property HelpFile() As String
            Get
                If IsNothing(rtbHelpFile) OrElse rtbHelpFile.Length = 0 Then rtbHelpFile = ProgramPathHTML + "/src/rtbHelp.html"
                Return rtbHelpFile
            End Get
            Set(ByVal value As String)
                rtbHelpFile = value
                wbRtbHelp.Navigate(rtbHelpFile)
                Do While IsNothing(wbRtbHelp.Url)
                    System.Windows.Forms.Application.DoEvents()
                Loop
                rtbHelpFile = wbRtbHelp.Url.ToString
            End Set
        End Property

        ''' <summary>Чвет выделения одинаковых переменных, функций ...</summary>
        Public Property HightLightColor As Color = Color.FromArgb(255, 211, 211, 211) ' DEFAULT_COLORS.codeBoxDesignate
#End Region

        ''' <summary>Полулокальная переменная, определяющая выделен ли сейчас фрагмент текста задним фоном</summary>
        Public wasTextSelected As Boolean = False

        ''' <summary>Для события codeBox_TextChanged. True если не надо изменять размер CodeData, поскольку это было сделано в свойстве SelectedText</summary>
        Public IsSelectedTextUsed As Boolean = False
        Public Overrides Property SelectedText As String
            Get
                Return MyBase.SelectedText
            End Get
            Set(value As String)
                If csUndo.UndoInProcess = False AndAlso MyBase.SelectedText <> value Then
                    csUndo.lastSelection = MyBase.SelectedText
                    csUndo.AppendItem(UndoClass.UndoTypeEnum.SELECTED_CHANGED, value)
                    csUndo.lastSelection = ""
                End If

                If IsNothing(mScript.mainClass) OrElse IsNothing(CodeData) Then
                    MyBase.SelectedText = value
                    Return
                End If

                'текущее выделение
                lastSelectionStartLine = GetLineFromCharIndex(SelectionStart)
                lastSelectionEndLine = GetLineFromCharIndex(SelectionStart + SelectionLength)
                Dim selectedLines As Integer = lastSelectionEndLine - lastSelectionStartLine + 1 'количество выделенных линий, которые заменяются вставляемым текстом

                'вставляемый текст. Получаем количество линий во вставляемом тексте
                Dim insertedLines As Integer = 1 'количество вставляемых линий
                If String.IsNullOrEmpty(value) = False Then
                    Dim pos As Integer = -1
                    Do
                        pos = value.IndexOf(vbLf, pos + 1)
                        If pos = -1 Then Exit Do
                        insertedLines += 1
                    Loop
                End If

                If selectedLines = 1 AndAlso selectedLines = insertedLines Then
                    'вставка в пределах одной линии - продолжаем, никаких особых действий не требуется
                    MyBase.SelectedText = value
                    Return
                End If

                IsSelectedTextUsed = True
                Dim linesDifference = insertedLines - selectedLines 'если больше 0, то количество вставляемых линий больше замененных, иначе - вставили линий меньше, чем было выделено
                If linesDifference > 0 Then
                    'линий после вставки стало больше на linesDifference. Расширяем CodeData
                    Dim cdList As List(Of CodeDataType) = CodeData.ToList
                    Dim cdInsert() As CodeDataType
                    ReDim cdInsert(linesDifference - 1)
                    For i As Integer = 0 To linesDifference - 1
                        cdInsert(i) = New CodeDataType
                    Next
                    cdList.InsertRange(lastSelectionStartLine, cdInsert)
                    CodeData = cdList.ToArray
                ElseIf linesDifference < 0 Then
                    'линий после вставки стало меньше - сужаем массив CodeData
                    Dim cdList As List(Of CodeDataType) = CodeData.ToList
                    cdList.RemoveRange(lastSelectionStartLine, -linesDifference)
                    CodeData = cdList.ToArray
                End If

                lastSelectionEndLine = lastSelectionStartLine + insertedLines - 1
                'Размер массива CodeData теперь имеет правильный размер, но содержимое между lastSelectionStartLine и lastSelectionEndLine неправильное.
                'Это уже исправит событие codeBox_TextChanged
                MyBase.SelectedText = value
            End Set
        End Property

#Region "Text Navigation"
        ''' <summary>
        ''' Выполняет проверку измененной строки с учетом изменения текстовых блоков и _. Переданный параметр lineEnd может увеличится, если надо проверять текст дальше (блока HTML/Wrap)
        ''' </summary>
        ''' <param name="rtb">Ссылка на RichTextBox</param>
        ''' <param name="lineStart">первая линия для проверки</param>
        ''' <param name="lineEnd">последняя строка для подготовки, по умолчанию = той, что редактировалась (т. е. строка одна)</param>
        ''' <param name="ignoreLineChanging">если True, то проверка сместилась ли каретка на следующую линию не происходит</param>
        ''' <param name="redrawUnabled">Если True, то включение прорисовки WM_SetRedraw не происходит</param>
        Public Sub CheckPrevLine(ByRef rtb As RichTextBox, Optional ByVal lineStart As Integer = -1, Optional ByVal lineEnd As Integer = -1, _
                                 Optional ByVal ignoreLineChanging As Boolean = False, Optional redrawUnabled As Boolean = False)
            Dim curLine As Integer = rtb.GetLineFromCharIndex(rtb.GetFirstCharIndexOfCurrentLine) 'текущая линия
            If lineStart = -1 Then lineStart = prevLineId

            If CodeData.GetUpperBound(0) < prevLineId Then
                prevLineId = -1
                prevLineContent = Nothing
                Return
            End If

            If prevLineId <> -1 AndAlso (prevLineId <> curLine OrElse ignoreLineChanging) Then
                'prevLineId - Id линии, которая только что редактировалась
                'curLine - Id линии, на которую перешли
                'выполняется если было редактирование и был совершен переход на другую линию...
                If CodeData.GetUpperBound(0) > -1 AndAlso (prevLineContent.Equals(CodeData(prevLineId)) = False OrElse ignoreLineChanging) Then
                    ', а также если структура CodeData не пуста и линия после редактирования изменилась

                    'получаем положение каретки по отношению к первому символу текущей строки, т. к. абсолютное положение каретки изменится
                    Dim selStart As Integer = rtb.SelectionStart - rtb.GetFirstCharIndexFromLine(curLine)
                    Dim selLength As Integer = rtb.SelectionLength
                    If lineEnd = -1 Then lineEnd = prevLineId 'последняя строка для подготовки, по умолчанию = той, что редактировалась (т. е. строка одна)

                    'Проверка на разбиение строк
                    If GetWordFromPrevLineContent(-1).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                        If GetWordFromCodeData(prevLineId, -1).wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                            'было разделение строк _ , а теперь убрали
                            lineEnd += 1 'будем подготавливать всю строку, разбитую с помощью _
                            'если на _ заканчиваются много строк, то получаем № строки за последней _
                            Do While lineEnd < CodeData.GetUpperBound(0) AndAlso IsNothing(CodeData(lineEnd - 1).Code) = False AndAlso CodeData(lineEnd - 1).Code(CodeData(lineEnd - 1).Code.GetUpperBound(0)).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION
                                lineEnd += 1
                            Loop
                        End If
                    End If
                    'Проверка на текстовые блоки
                    Dim scndWord As New EditWordType
                    scndWord = GetWordFromCodeData(prevLineId, 1) 'второе слово в строке, что редактировалась
                    If GetWordFromCodeData(prevLineId, 0).wordType = EditWordTypeEnum.W_CYCLE_END Then
                        If (scndWord.wordType = EditWordTypeEnum.W_HTML OrElse scndWord.wordType = EditWordTypeEnum.W_WRAP) OrElse _
                            (String.Compare(scndWord.Word, "html") = 0 OrElse String.Compare(scndWord.Word, "wrap") = 0) Then
                            'стал End HTML/Wrap (но раньше им не был)
                            lineEnd = CodeData.GetUpperBound(0) 'подготавливаем текст до конца
                        End If

                    End If
                    'scndWord = GetWordFromPrevLineContent(1)
                    If GetWordFromPrevLineContent(0).wordType = EditWordTypeEnum.W_CYCLE_END AndAlso (scndWord.wordType = EditWordTypeEnum.W_HTML OrElse scndWord.wordType = EditWordTypeEnum.W_WRAP) Then
                        'был End HTML/Wrap (а стал чем-то другим)
                        lineEnd = CodeData.GetUpperBound(0) 'подготавливаем текст до конца
                    End If
                    Dim fstWord As New EditWordType
                    fstWord = GetWordFromCodeData(prevLineId, 0) 'первое слово в строке, что редактировалась
                    If fstWord.wordType = EditWordTypeEnum.W_VARIABLE Then
                        If String.Compare(fstWord.Word, "html", True) = 0 Then
                            fstWord.Word = "HTML"
                            fstWord.wordType = EditWordTypeEnum.W_HTML
                        ElseIf String.Compare(fstWord.Word, "wrap", True) Then
                            fstWord.Word = "WRAP"
                            fstWord.wordType = EditWordTypeEnum.W_WRAP
                        End If
                    End If
                    scndWord = GetWordFromPrevLineContent(0) 'первое слово в строке, что редактировалась, до редактирования
                    If (fstWord.wordType = EditWordTypeEnum.W_HTML Or scndWord.wordType = EditWordTypeEnum.W_HTML) Then ' AndAlso fstWord.wordType <> scndWord.wordType Then
                        'Стало началом блока HTML
                        Do While lineEnd < CodeData.GetUpperBound(0) 'Ищем End HTML - это будет концом блока, который надо подготавливать. Если не найдем - до конца текста
                            If GetWordFromCodeData(lineEnd + 1, 0).wordType = EditWordTypeEnum.W_CYCLE_END AndAlso _
                                GetWordFromCodeData(lineEnd + 1, 1).wordType = EditWordTypeEnum.W_HTML AndAlso _
                                GetWordFromCodeData(lineEnd + 1, 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                Exit Do
                            End If
                            lineEnd += 1
                        Loop
                    End If
                    If (fstWord.wordType = EditWordTypeEnum.W_WRAP Or scndWord.wordType = EditWordTypeEnum.W_WRAP) AndAlso fstWord.wordType <> scndWord.wordType Then
                        'Стало началом блока Wrap
                        Do While lineEnd < CodeData.GetUpperBound(0) 'Ищем End Wrap - это будет концом блока, который надо подготавливать. Если не найдем - до конца текста
                            If GetWordFromCodeData(lineEnd + 1, 0).wordType = EditWordTypeEnum.W_CYCLE_END AndAlso _
                                GetWordFromCodeData(lineEnd + 1, 1).wordType = EditWordTypeEnum.W_WRAP AndAlso _
                                GetWordFromCodeData(lineEnd + 1, 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                Exit Do
                            End If
                            lineEnd += 1
                        Loop
                    End If
                    'проверка на измененный exec
                    Dim wasExecOpen As Boolean = False, nowExecOpen As Boolean = False 'был <exec> и есть ли после редактирования
                    Dim wasExecClose As Boolean = False, nowExecClose As Boolean = False 'был </exec> и есть ли после редактирования
                    Dim wasOpenBeforeClose As Boolean = False, nowOpenBeforeClose As Boolean = False 'что встречается раньше: </exec> или <exec> (до и после редактирования)
                    Dim execOpen As Integer, execClose As Integer 'для поиска тегов <exec> и </exec>
                    'ищем в строке после редактирования
                    If IsNothing(CodeData(prevLineId).Code) = False Then
                        For i As Integer = CodeData(prevLineId).Code.GetUpperBound(0) To 0 Step -1
                            If CodeData(prevLineId).Code(i).wordType = EditWordTypeEnum.W_HTML_DATA Then
                                If nowExecOpen = False Then execOpen = CodeData(prevLineId).Code(i).Word.LastIndexOf("<exec>", StringComparison.CurrentCultureIgnoreCase)
                                If nowExecClose = False Then execClose = CodeData(prevLineId).Code(i).Word.LastIndexOf("</exec>", StringComparison.CurrentCultureIgnoreCase)
                                If nowExecOpen = False AndAlso execOpen > -1 Then
                                    '<exec> найден и он не был найден ранее
                                    nowExecOpen = True
                                    If nowExecClose Then
                                        '</exec> найден ранее
                                        nowOpenBeforeClose = True
                                    ElseIf execClose = -1 Then
                                        '</exec> не найден и не был найден ранее
                                        nowOpenBeforeClose = False
                                    ElseIf execClose > execOpen Then
                                        '</exec> не был найден ранее, а сейчас он найден после <exec>
                                        nowOpenBeforeClose = False
                                        nowExecClose = True
                                    Else
                                        '</exec> не был найден ранее, а сейчас он найден до <exec>
                                        nowExecClose = True
                                        nowOpenBeforeClose = True
                                    End If
                                ElseIf nowExecClose = False AndAlso execClose > -1 Then
                                    '</exec> найден, но не был не найден ранее, к тому же
                                    '<exec> либо уже найден, либо не найден в текущей структуре
                                    nowExecClose = True
                                    If nowExecOpen Then
                                        '<exec> найден ранее
                                        nowOpenBeforeClose = False
                                    End If
                                End If
                                If nowExecOpen And nowExecClose Then Exit For
                            End If
                        Next
                    End If
                    'ищем в строке до редактирования
                    If IsNothing(prevLineContent.Code) = False Then
                        For i As Integer = prevLineContent.Code.GetUpperBound(0) To 0 Step -1
                            If prevLineContent.Code(i).wordType = EditWordTypeEnum.W_HTML_DATA Then
                                If wasExecOpen = False Then execOpen = prevLineContent.Code(i).Word.LastIndexOf("<exec>", StringComparison.CurrentCultureIgnoreCase)
                                If wasExecClose = False Then execClose = prevLineContent.Code(i).Word.LastIndexOf("</exec>", StringComparison.CurrentCultureIgnoreCase)
                                If wasExecOpen = False AndAlso execOpen > -1 Then
                                    '<exec> найден и он не был найден ранее
                                    wasExecOpen = True
                                    If wasExecClose Then
                                        '</exec> найден ранее
                                        wasOpenBeforeClose = True
                                    ElseIf execClose = -1 Then
                                        '</exec> не найден и не был найден ранее
                                        wasOpenBeforeClose = False
                                    ElseIf execClose > execOpen Then
                                        '</exec> не был найден ранее, а сейчас он найден после <exec>: <exec>...</exec>
                                        wasOpenBeforeClose = True
                                        wasExecClose = True
                                    Else
                                        '</exec> не был найден ранее, а сейчас он найден до <exec>: </exec>...<exec>
                                        wasExecClose = True
                                        wasOpenBeforeClose = False
                                    End If
                                ElseIf wasExecClose = False AndAlso execClose > -1 Then
                                    '</exec> найден, но не был не найден ранее, к тому же
                                    '<exec> либо уже найден, либо не найден в текущей структуре
                                    wasExecClose = True
                                    If wasExecOpen Then
                                        '<exec> найден ранее
                                        wasOpenBeforeClose = False
                                    End If
                                End If
                                If wasExecOpen And wasExecClose Then Exit For
                            End If
                        Next
                    End If
                    'если относительное расположение тегов <exec> и </exec> изменилось - проверка текста до конца
                    If wasExecOpen <> nowExecOpen OrElse wasExecClose <> nowExecClose OrElse wasOpenBeforeClose <> nowOpenBeforeClose Then lineEnd = CodeData.GetUpperBound(0) 'был изменен тег exec - проверяем весь текст до конца

                    'получаем линию, которая сейчас сверху (чтоб затем вернуть прокрутку вертикального скролла туда, где он был до подготовки текста, т. е. сейчас)
                    Dim upperLine As Integer = rtb.GetLineFromCharIndex(rtb.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5}))

                    If redrawUnabled = False Then SendMessage(rtb.Handle, WM_SetRedraw, 0, 0) 'запрет перерисовки (для ускорения)
                    Dim prevCanRaiseEvents As Boolean = rtbCanRaiseEvents
                    rtbCanRaiseEvents = False 'запрет на выполнение событий в RichTextBox
                    PrepareText(rtb, lineStart, lineEnd, True) 'непосредственно подготовка текста
                    rtbCanRaiseEvents = prevCanRaiseEvents  'снова разрешаем выполнение событий в RichTextBox
                    rtb.SelectionStart = selStart + rtb.GetFirstCharIndexFromLine(curLine) 'возвращаем положение каретки относительно начала текущей строки
                    rtb.SelectionLength = selLength
                    'прокручиваем текст обратно, как было до разрисовки текста
                    SendMessage(rtb.Handle, EM_LINESCROLL, 0, upperLine - rtb.GetLineFromCharIndex(rtb.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})))
                    If redrawUnabled = False Then SendMessage(rtb.Handle, WM_SetRedraw, 1, 0) 'разрешаем перерисовку
                    rtb.Refresh() 'обновляем (т. к. перерисовка была запрещена)
                End If
            End If
        End Sub

        ''' <summary>
        ''' получаем слово из структуры CodeData, находящееся в линии lineId на месте wordId, а если его там нет - пустое слово
        ''' </summary>
        ''' <param name="lineId">Id линии в кодбоксе</param>
        ''' <param name="wordId">Id слова в структуре CodeData</param>
        Public Function GetWordFromCodeData(ByVal lineId As Integer, ByVal wordId As Integer) As EditWordType
            If IsNothing(mScript.mainClass) Then Return Nothing
            '
            If lineId >= CodeData.Count Then Return Nothing
            If IsNothing(CodeData(lineId).Code) Then Return New EditWordType
            If wordId < 0 Then
                wordId = CodeData(lineId).Code.GetUpperBound(0)
            Else
                If wordId > CodeData(lineId).Code.GetUpperBound(0) Then Return New EditWordType
            End If
            Return CodeData(lineId).Code(wordId)
        End Function

        Private Function GetWordFromPrevLineContent(ByVal wordId As Integer) As EditWordType
            If IsNothing(mScript.mainClass) Then Return Nothing
            'получаем слово из prevLineContent, находящееся на месте wordId, а если его там нет - пустое слово
            If IsNothing(prevLineContent) Then Return New EditWordType
            If IsNothing(prevLineContent.Code) Then Return New EditWordType
            If wordId < 0 Then
                wordId = prevLineContent.Code.GetUpperBound(0)
            Else
                If wordId > prevLineContent.Code.GetUpperBound(0) Then Return New EditWordType
            End If
            Return prevLineContent.Code(wordId)
        End Function

        Public Sub SetPrevLine(ByRef rtb As RichTextBox)
            'Сохраняет код и номер текущей строки перед ее редактированием
            If CodeData.GetUpperBound(0) = -1 Then
                prevLineContent = Nothing
                prevLineId = -1
                Return
            End If
            Dim curLine As Integer = rtb.GetLineFromCharIndex(rtb.GetFirstCharIndexOfCurrentLine) 'текущая линия
            If curLine > CodeData.GetUpperBound(0) Then Return

            If curLine < 0 Then
                prevLineContent = Nothing
                prevLineId = -1
                Return
            End If

            If IsNothing(CodeData(curLine)) Then
                MsgBox("Структура CodeData не соответствует тексту!", MsgBoxStyle.Critical)
                prevLineContent = Nothing
                prevLineId = -1
                Return
            End If

            'сохраняем код строки
            prevLineContent = New CodeDataType
            prevLineContent = CodeData(curLine)
            'сохраняем номер строки
            prevLineId = curLine
        End Sub

        ''' <summary>
        ''' процедура выделяет слово, внутри которого находится каретка
        ''' </summary>
        ''' <param name="rtb"></param>
        Public Sub SelectCurrentWord(ByRef rtb As RichTextBox, ByRef lstRtb As ListBox)
            'процедура выделяет слово, внутри которого находится каретка
            Dim wordStart As Integer = 0, wordEnd As Integer = 0
            'получаем начало и конец слова
            If lstRtb.Items.Count > 0 AndAlso lstRtb.Items(0).ToString.Length > 0 AndAlso lstRtb.Items(0).ToString.First = "'"c Then
                'ищем символ, который не может быть продолжением слова (напр., пробел или перевод строки) в списке, где нет '
                'это надо когда в списке находятся строки (т. е. 'xxx') - для коректной замены словом из списка, - символ ' здесь является частью слова, а не его концом
                wordStart = rtb.Text.Substring(0, rtb.SelectionStart).LastIndexOfAny(wordBoundsArrayNoQuotes) + 1
                wordEnd = rtb.Text.IndexOfAny(wordBoundsArrayNoQuotes, rtb.SelectionStart + rtb.SelectionLength)
            Else
                'ищем символ, который не может быть продолжением слова (напр., пробел или перевод строки) в полном списке
                wordStart = rtb.Text.Substring(0, rtb.SelectionStart).LastIndexOfAny(wordBoundsArray) + 1
                wordEnd = rtb.Text.IndexOfAny(wordBoundsArray, rtb.SelectionStart + rtb.SelectionLength)
            End If
            If wordStart = -1 Then wordStart = 0
            If wordEnd = -1 Then
                wordEnd = rtb.TextLength - 1 '
                rtb.Select(wordStart, wordEnd - wordStart + 1)
                Return
            End If
            If wordEnd - wordStart < 0 Then Return 'стоим между словами (а не внутри слова) - ничего не выделяем
            'непосредственно выделяем
            rtb.Select(wordStart, wordEnd - wordStart)
        End Sub

        Public Sub SetCorretSpaces(ByRef rtb As RichTextBoxEx, Optional stopRedrawing As Boolean = True)
            'Процедура расставляет правильные пробелы в начале каждой строки (просматривается весь текст), а также вставляет окончания незавершенных блоков
            If IsNothing(CodeData) OrElse rtb.csUndo.UndoInProcess Then Return

            Dim spaceChar As String = "  " ' vbTab 'пробельный символ или строка, на которую смещается код (например, внутри блока)
            Dim newSpaces As String = "" 'для хранения строки с пробелами, какая должна быть
            Dim oldSpaces As String = "" 'для получения строки с пробелами, которая сейчас
            Dim blnAddSpace As Boolean = False 'надо ли добавить еще один знак к пробельной строке
            Dim firstWord As EditWordType 'для получения первого слова каждой строки
            Dim strWord As String 'для слова из firstWord
            Dim textBlock As TextBlockEnum = TextBlockEnum.NO_TEXT_BLOCK 'для информации не в текстовом ли блоке мы сейчас

            'сохраняем текущие верхнюю видимую строку, выделение и линию
            Dim upperLine As Integer = rtb.GetLineFromCharIndex(rtb.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5}))
            Dim curSelStart As Integer = rtb.SelectionStart - rtb.GetFirstCharIndexOfCurrentLine
            Dim curLine As Integer = rtb.GetLineFromCharIndex(rtb.GetFirstCharIndexOfCurrentLine)
            Dim blockBalance() As Integer, blockId As Integer 'баланс открытия / закрытия блоков
            ReDim blockBalance(AddSpacesArray.GetUpperBound(0)) 'расширяем массив по кол-ву видов блоков
            Dim strAppendEndBlock As String = "" 'строка завершения блока для вставки в rtb
            Dim codeAppendEndBlock() As EditWordType = Nothing  'код завершения блока для вставки в CodeData
            Dim strAppendBlockSpaces As String = "" 'начальные пробелы блока
            Dim appendBlockSelStart As Integer = 0 'куда вставлять завершение блока
            Dim appendBlockId As Integer = -1 'кой именно блок вставляем
            'запрет событий в rtb
            Dim prevCanRaiseEvents As Boolean = rtbCanRaiseEvents
            rtbCanRaiseEvents = False
            If stopRedrawing Then SendMessage(rtb.Handle, WM_SetRedraw, 0, 0) 'запрет перерисовки (для ускорения)

            Dim i As Integer = 0
checkAgain:
            While i <= CodeData.GetUpperBound(0) 'перебираем все строки
                If i > Lines.Count - 1 Then Exit While
                If blnAddSpace Then 'в пробельную строку добавляем еще один знак
                    newSpaces += spaceChar
                    blnAddSpace = False
                End If
                'получаем текущую пробельную строку
                oldSpaces = CodeData(i).StartingSpaces
                If IsNothing(oldSpaces) Then oldSpaces = ""

                If IsNothing(CodeData(i).Code) OrElse CodeData(i).Code.GetUpperBound(0) = -1 Then
                    'в этой строке нет кода
                    If newSpaces = "" AndAlso IsNothing(CodeData(i).Comments) Then i += 1 : Continue While 'нет и начальных пробелов - продолжаем цикл
                    If curLine = i Then
                        'на этой строке стоит каретка. Добавляем перед ней правильную пробельную строку
                        curSelStart = newSpaces.Length 'смещаем каретку 
                        CodeData(i).StartingSpaces = newSpaces 'вставляем правильные пробелы в CodeData
                        rtb.Select(rtb.GetFirstCharIndexFromLine(i), 0)
                        rtb.SelectedText = newSpaces 'вставляем правильные пробелы в rtb

                        If i > 0 AndAlso IsNothing(CodeData(i - 1).Code) = False Then
                            'проверяем предыдущую строку на незавершенный блок
                            firstWord = CodeData(i - 1).Code(0) 'получаем первое слово предыдущей строки
                            strWord = firstWord.Word.Trim
                            blockId = Array.IndexOf(AddSpacesArray, firstWord.wordType) 'поиск блока
                            If (firstWord.wordType = EditWordTypeEnum.W_BLOCK_NEWCLASS AndAlso strWord <> "New") = False AndAlso _
                                strWord <> "Case" AndAlso strWord <> "Else" AndAlso strWord <> "ElseIf" AndAlso _
                                (firstWord.wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0)).Word.Trim <> "Then") = False Then
                                'в предыдущей строке начало блока. Создаем данные для вставки конца блока, если это понадобится
                                Select Case firstWord.wordType 'создаем такст и код для вставки
                                    Case EditWordTypeEnum.W_BLOCK_DOWHILE
                                        strAppendEndBlock = "Loop"
                                        ReDim codeAppendEndBlock(0)
                                        codeAppendEndBlock(0) = New EditWordType With {.Word = strAppendEndBlock, .wordType = EditWordTypeEnum.W_BLOCK_DOWHILE}
                                    Case EditWordTypeEnum.W_BLOCK_FOR
                                        strAppendEndBlock = "Next"
                                        ReDim codeAppendEndBlock(0)
                                        codeAppendEndBlock(0) = New EditWordType With {.Word = strAppendEndBlock, .wordType = EditWordTypeEnum.W_BLOCK_FOR}
                                    Case EditWordTypeEnum.W_BLOCK_NEWCLASS
                                        strAppendEndBlock = "End Class"
                                        ReDim codeAppendEndBlock(1)
                                        codeAppendEndBlock(0) = New EditWordType With {.Word = "End ", .wordType = EditWordTypeEnum.W_CYCLE_END}
                                        codeAppendEndBlock(1) = New EditWordType With {.Word = "Class", .wordType = firstWord.wordType}
                                    Case Else
                                        strAppendEndBlock = "End " + strWord
                                        ReDim codeAppendEndBlock(1)
                                        codeAppendEndBlock(0) = New EditWordType With {.Word = "End ", .wordType = EditWordTypeEnum.W_CYCLE_END}
                                        codeAppendEndBlock(1) = New EditWordType With {.Word = strWord, .wordType = firstWord.wordType}
                                End Select
                                strAppendEndBlock = vbNewLine + newSpaces + strAppendEndBlock 'полная строка с пробелами для вставки
                                strAppendBlockSpaces = newSpaces 'сохраняем строку пробелов на этот момент
                                appendBlockId = blockId 'сохраняем Id блока
                                appendBlockSelStart = rtb.SelectionStart 'сохраняем положение каретки, куда потом вставим блок
                            End If
                        End If
                    End If
                    If oldSpaces <> newSpaces AndAlso IsNothing(CodeData(i).Comments) = False Then
                        CodeData(i).StartingSpaces = newSpaces 'вставляем правильные пробелы в CodeData
                        rtb.Select(rtb.GetFirstCharIndexFromLine(i), oldSpaces.Length)
                        rtb.SelectedText = newSpaces 'вставляем правильные пробелы в rtb
                    End If
                    i += 1
                    Continue While
                End If

                firstWord = CodeData(i).Code(0) 'получаем первое слово текущей строки
                strWord = firstWord.Word.Trim
                If textBlock <> TextBlockEnum.NO_TEXT_BLOCK Then
                    'мы в текстовом блоке
                    If textBlock = TextBlockEnum.TEXT_HTML Then
                        'точнее, в блоке HTML
                        If firstWord.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso CodeData(i).Code.GetUpperBound(0) > 0 _
                            AndAlso CodeData(i).Code(1).wordType = EditWordTypeEnum.W_HTML Then
                            textBlock = TextBlockEnum.NO_TEXT_BLOCK 'End HTML - конец блока
                        Else
                            i += 1
                            Continue While 'текстовый блок продолжается
                        End If
                    ElseIf textBlock = TextBlockEnum.TEXT_WRAP Then
                        'точнее, в блоке Wrap
                        If firstWord.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso CodeData(i).Code.GetUpperBound(0) > 0 _
                            AndAlso CodeData(i).Code(1).wordType = EditWordTypeEnum.W_WRAP Then
                            textBlock = TextBlockEnum.NO_TEXT_BLOCK 'End Wrap - конец блока
                        Else
                            i += 1
                            Continue While 'текстовый блок продолжается
                        End If
                    End If
                End If

                If firstWord.wordType = EditWordTypeEnum.W_CYCLE_END OrElse strWord = "Loop" OrElse strWord = "Next" OrElse
                    strWord = "Case" OrElse strWord = "Else" OrElse strWord = "ElseIf" Then
                    'первое слово End (или Loop и др., тоже конец блока)
                    If newSpaces.Length > spaceChar.Length - 1 Then newSpaces = newSpaces.Remove(newSpaces.Length - spaceChar.Length) 'сдвиг влево
                    If firstWord.wordType = EditWordTypeEnum.W_CYCLE_END Then
                        blockId = Array.IndexOf(AddSpacesArray, CodeData(i).Code(1).wordType)
                        If blockId > -1 Then blockBalance(blockId) -= 1
                    Else
                        Dim pos = Array.IndexOf(AddSpacesArray, firstWord.wordType)
                        If pos > -1 Then
                            blockBalance(pos) -= 1
                            'Case, Else и ElseIf должны сдвинуться влево, на уровень начала блока, но последующий код должн опять вернуться вправо
                            If strWord = "Case" OrElse strWord = "Else" OrElse strWord = "ElseIf" Then blnAddSpace = True
                        End If
                    End If
                ElseIf (firstWord.wordType = EditWordTypeEnum.W_BLOCK_NEWCLASS AndAlso strWord <> "New") = False AndAlso _
                    (firstWord.wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso CodeData(i).Code(CodeData(i).Code.GetUpperBound(0)).Word.Trim <> "Then") = False Then
                    'это блок
                    blockId = Array.IndexOf(AddSpacesArray, firstWord.wordType)
                    If blockId > -1 Then
                        'это блок 
                        blnAddSpace = True 'следующая строка сдвинется вправо
                        blockBalance(blockId) += 1
                    End If
                End If

                If firstWord.wordType = EditWordTypeEnum.W_HTML Then
                    textBlock = TextBlockEnum.TEXT_HTML 'мы вошли в блок HTML
                    i += 1
                    Continue While
                ElseIf firstWord.wordType = EditWordTypeEnum.W_WRAP Then
                    textBlock = TextBlockEnum.TEXT_WRAP 'мы вошли в блок HTML
                    i += 1
                    Continue While
                End If

                If newSpaces = oldSpaces Then i += 1 : Continue While 'правильное кол-во пробелов и текущее совпадают - ничего не делаем
                If curLine = i Then 'если на текущей строке каретка, то смещаем каретку соответственно новой длине пробельной строки
                    curSelStart = newSpaces.Length - oldSpaces.Length
                End If
                CodeData(i).StartingSpaces = newSpaces 'вставляем правильные пробелы в CodeData
                rtb.Select(rtb.GetFirstCharIndexFromLine(i), oldSpaces.Length)
                rtb.SelectedText = newSpaces 'вставляем правильные пробелы в rtb

                i += 1
            End While

            If appendBlockId > -1 AndAlso blockBalance(appendBlockId) = 1 AndAlso strAppendEndBlock.Length > -0 Then
                'для полного баланса не хватает завершения блока - значит его надо вставить
                'вставляем завершения блока в rtb в сохраненную позицию
                rtb.SelectionStart = appendBlockSelStart
                rtb.SelectedText = strAppendEndBlock
                'расширяем CodeData и вставляем в него конец блока
                'ReDim Preserve CodeData(CodeData.Length)
                'For i = CodeData.GetUpperBound(0) To curLine + 1 Step -1
                '    CodeData(i) = CodeData(i - 1) 'сдвигаем инфо в CodeData на 1 вправо (в конец), начиная со строки, следующей за текущей
                'Next
                CodeData(curLine + 1).Code = codeAppendEndBlock
                CodeData(curLine + 1).StartingSpaces = strAppendBlockSpaces
                CodeData(curLine).StartingSpaces = strAppendBlockSpaces
                DrawWords(rtb, codeAppendEndBlock, curLine + 1, strAppendBlockSpaces.Length, False) 'разукрашиваем строку
                appendBlockId = -1
                'возвращаемся и перепроставляем пробелы везде под вновь вставленным концом блока
                i = curLine + 1
                newSpaces = strAppendBlockSpaces
                GoTo checkAgain
            End If

            If stopRedrawing Then SendMessage(rtb.Handle, WM_SetRedraw, 1, 0) 'разрешаем перерисовку
            'прокручиваем текст обратно, как было раньше
            If stopRedrawing Then SendMessage(rtb.Handle, EM_LINESCROLL, 0, upperLine - rtb.GetLineFromCharIndex(rtb.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})))
            rtb.SelectionStart = curSelStart + rtb.GetFirstCharIndexFromLine(curLine)
            rtb.Refresh() 'обновляем (т. к. перерисовка была запрещена)
            rtbCanRaiseEvents = True
        End Sub

#End Region

#Region "Working with Code"
        ''' <summary> Функция приводит текст из RichTextBox в вид, понятный движку (убирает/расставляет пробелы, приводит операторы в правильный регистр и т. д.,
        '''а также закрашивает слова в RichTextBox исходя из того, что они значат (цифры - в один цвет, операторы в другой и т. д.).
        ''' Предназначена для уже готовых фрагментов (а не тех, которые юзер еще пишет)
        ''' </summary>
        ''' <param name="rtb">RichTextBox, с которым работаем</param>
        ''' <param name="startLine">первая строка, с которой работаем</param>
        ''' <param name="endLine">последняя строка, с которой работаем (если -1, то последняя строка)</param>
        ''' <param name="drawText">раскрашивать ли текст в RichTextBox</param>
        ''' <param name="currentWordId">порядковый номер текущего слова (при редактировании строки) или -1, если не известно</param>
        Public Sub PrepareText(ByRef rtb As RichTextBoxEx, Optional ByVal startLine As Integer = 0, Optional ByVal endLine As Integer = -1, _
                           Optional ByVal drawText As Boolean = True, Optional ByRef currentWordId As Integer = -1)
            If DontShowError Then
                PrepareTextNonCompleted(rtb, startLine, endLine, drawText, currentWordId)
                Return
            End If
            currentWordId = -1
            If endLine = -1 Then endLine = rtb.Lines.GetUpperBound(0)
            mScript.LAST_ERROR = "" 'обнуляем ошибки
            PrepareHTMLforErrorInfo(wbRtbHelp)

            If rtb.Lines.GetUpperBound(0) = -1 AndAlso CodeData.GetUpperBound(0) = 0 Then 'проверка на пустой текст
                rtb.ForeColor = Color.Black
                CodeData(0) = New CodeDataType
                Return
            End If

            'проверяем структуру CodeData
            'Эта структура создается здесь, если подготавливается весть текст.
            'Если подготавливается лишь часть текста, то структура должа быть создана заранее (из всего текста)
            If startLine = 0 And (endLine = rtb.Lines.GetUpperBound(0) OrElse rtb.Lines.GetUpperBound(0) = -1) Then
                'подготавливается весь текст. Создаем структуру CodeData
                ReDim CodeData(rtb.Lines.GetUpperBound(0))
            ElseIf IsNothing(CodeData) Then
                'подготавливается только часть текста, а структура CodeData еще не создана - ошибка
                MsgBox("PrepareText: структура CodeData не подготовлена!", vbCritical)
                Return
            ElseIf CodeData.GetUpperBound(0) <> rtb.Lines.GetUpperBound(0) Then
                'подготавливается только часть текста, а размер CodeData не соответствует кол-ву строк в RichTextBox - ошибка

                startLine = 0
                endLine = rtb.Lines.GetUpperBound(0)
                ReDim CodeData(rtb.Lines.GetUpperBound(0))
                'MsgBox("PrepareText: структура CodeData не подготовлена!", vbCritical)
                'Return
            End If
            rtb.csUndo.SetCurrentUpperLineAsDefault() 'для корректной работы Undo сохраняем текущую верхнюю линию кодбокса

            Dim curString As String = "" 'текущая строка
            Dim i, j As Integer

            Dim selStart As Integer = rtb.SelectionStart 'начало выделения в RichTextBox
            Dim selLength As Integer = rtb.SelectionLength 'длина выделения в RichTextBox

            Dim arrayWords() As EditWordType = Nothing, arrayWordsUBound As Integer = 0 'массив, в котором храним слова из строки, ее размер
            Dim arrayWordsCopy() As EditWordType = Nothing  'массив для сборки строки, разбитой _
            Dim ovalBracketBalance As Integer = 0 'баланс ( и )
            Dim quadBracketBalance As Integer = 0 'баланс [ и ]
            Dim wordStart As Integer 'положение в строке первого символа слова, с которым будем работать
            Dim isPrevStringDissociation As Boolean 'разбита ли строка с помощью _
            Dim execOpenPos As Integer 'для нахождения тега <exec>
            Dim isInExec As ExecBlockEnum = ExecBlockEnum.NO_EXEC
            Dim isInText As TextBlockEnum = IsInTextBlock(rtb, startLine, isInExec)  'находимся ли мы внутри блока HTML или Wrap
            Dim textStartLine As Integer 'индекс первой строки с текстом, идущей после начала блока Wrap или HTML 
            If isInText <> TextBlockEnum.NO_TEXT_BLOCK Then textStartLine = startLine 'если начало текстового блока до начала кода - textStartLine = первой обрабатываемой строке

            i = startLine
            If endLine >= rtb.Lines.Length Then endLine = rtb.Lines.Length - 1
            While i <= endLine
                If drawText AndAlso CanDrawWords Then
                    'если это первый виток цикла - сбрасываем цвет и стиль текста на значение по-умолчанию
                    If i = startLine Then
                        rtb.Select(rtb.GetFirstCharIndexFromLine(startLine), rtb.GetFirstCharIndexFromLine(endLine) + rtb.Lines(endLine).Length - rtb.GetFirstCharIndexFromLine(startLine))
                        rtb.SelectionColor = styleHash(EditWordTypeEnum.W_NOTHING).style_Color
                        rtb.SelectionFont = New Drawing.Font(rtb.Font, styleHash(EditWordTypeEnum.W_NOTHING).font_style)
                    End If
                End If

                j = -1
                wordStart = 0
                If isInText <> TextBlockEnum.NO_TEXT_BLOCK Then
                    'мы внутри блока HTML / Wrap
                    curString = rtb.Lines(i)
                    Dim strLower As String = curString.Trim.Replace("  ", " ").ToLower
                    Dim strSearch As String = IIf(isInText = TextBlockEnum.TEXT_HTML, "end html", "end wrap")
                    If strLower.StartsWith(strSearch) AndAlso StripComments(strLower, Nothing, Nothing, False) <> 2 AndAlso strLower = strSearch Then
                        'StripComments <> 2 выполнено в блоке If Then только чтоб очистить curString от комментариев и пробелов перед третьей проверкой
                        'текущая строка - "End HTML" или "End Wrap" - конец блока
                        isInText = TextBlockEnum.NO_TEXT_BLOCK
                        isInExec = ExecBlockEnum.NO_EXEC
                        If drawText AndAlso i - 1 > textStartLine Then DrawHTML(rtb, textStartLine, i - 1)
                    Else
                        'продолжение текстового блока
                        If isInExec = ExecBlockEnum.NO_EXEC Then
                            execOpenPos = curString.IndexOf("<exec>", StringComparison.CurrentCultureIgnoreCase)
                            If execOpenPos > -1 Then
                                isInExec = ExecBlockEnum.EXEC_SINGLE
                            End If

                            If i = endLine And drawText Then DrawHTML(rtb, textStartLine, i) 'дошли до конца и не нашли конец блока - разрисовываем текстовый блок
                            'записываем текстовый блок в массив CodeData
                            ReDim CodeData(i).Code(0)
                            If isInExec Then
                                execOpenPos += "<exec>".Length
                                ReDim arrayWords(20)
                                arrayWordsUBound = 0 'id последней заполненной ячейки массива
                                arrayWords(0).Word = curString.Substring(0, execOpenPos)
                                arrayWords(0).wordType = EditWordTypeEnum.W_HTML_DATA
                                j = execOpenPos - 1
                                wordStart = execOpenPos
                            Else
                                CodeData(i).Code(0).Word = rtb.Lines(i)
                                CodeData(i).Code(0).wordType = EditWordTypeEnum.W_HTML_DATA
                                i += 1
                                Continue While
                            End If
                        End If
                    End If
                End If

                'определяем, не заканчивается ли строка перед первой, которую мы обрабатываем, на _
                If i = startLine And i > 0 AndAlso isInExec <> ExecBlockEnum.EXEC_SINGLE Then
                    curString = rtb.Lines(i - 1)
                    If StripComments(curString, Nothing, Nothing, False) > -1 Then
                        If curString.Length > 0 AndAlso curString.EndsWith("_") Then
                            'если предыдущая строка - начало текущей (самой первой) - перезапускаем цикл, начиная с предыдущей строки
                            i -= 1
                            startLine = i
                            Continue While
                        End If
                    End If
                End If

                curString = rtb.Lines(i) 'получаем текущую линию

                'Очищаем строку от дополнительных элементов (пробелов в начале, комментариев...)
                If isInExec <> ExecBlockEnum.EXEC_SINGLE AndAlso StripComments(curString, CodeData(i).Comments, CodeData(i).StartingSpaces, False) = -1 Then
                    ShowTextError(rtb, i, False)
                    ovalBracketBalance = 0
                    quadBracketBalance = 0
                    isPrevStringDissociation = False
                    Erase arrayWordsCopy
                    i += 1
                    Continue While
                End If
                If IsNothing(CodeData(i).Comments) Then CodeData(i).Comments = ""
                If IsNothing(CodeData(i).StartingSpaces) Then CodeData(i).StartingSpaces = ""

                'в строке были только пробелы и комментарии - переходим сразу к новой строке
                If curString.Length = 0 Then
                    Erase CodeData(i).Code
                    If drawText AndAlso CodeData(i).Comments.Length > 0 Then DrawWords(rtb, Nothing, i, CodeData(i).StartingSpaces.Length, True)
                    i += 1
                    Continue While
                End If

                'Получаем в массив arrayWords слова из рабочей строки
                If isInExec <> ExecBlockEnum.EXEC_SINGLE Then
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
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN 'устанавливаем тип содержимого
                            wordStart = j + 1
                        Case ")"c
                            ovalBracketBalance -= 1
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = ")"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                            wordStart = j + 1
                        Case "["c
                            quadBracketBalance += 1
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "["
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                            wordStart = j + 1
                        Case "]"c
                            quadBracketBalance -= 1
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "]"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
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
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_POINT
                            wordStart = j + 1
                        Case ","c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = ", "
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_COMMA
                            wordStart = j + 1
                        Case "="c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = " = "
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL
                            wordStart = j + 1
                        Case "?"
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "?"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_RETURN
                            wordStart = j + 1
                        Case "&"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = " & "
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER
                            wordStart = j + 1
                        Case "#"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "#"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER
                            wordStart = j + 1
                        Case "$"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "$"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING
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
                                        If arrayWordsUBound > 0 AndAlso (arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_COMMA) Then
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
                                        If strUnar = "to" OrElse strUnar = "case" OrElse strUnar = "step" OrElse strUnar = "return" OrElse strUnar = "?" _
                                            OrElse IsMinusUnar(arrayWords, isPrevStringDissociation, i, arrayWordsUBound) Then
                                            'если за знаком "-" идет число, а перед ним - To Case или  Step, то 
                                            'это - унарный минус, который относится к следующей за ним цифре. Продолжаем цикл, не сохраняя минус отдельно
                                            wordStart = j
                                            curString = curString.Substring(0, j + 1) + curString.Substring(j + 1).TrimStart
                                            Continue Do
                                        End If
                                        If isPrevStringDissociation And j = 0 Then
                                            If CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_MATH OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_BLOCK_FOR OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_COMMA Then
                                                'то же, только предыдущая строка заканчивалась на _, а минус - первый символ строки
                                                wordStart = j
                                                curString = curString.Substring(0, j + 1) + curString.Substring(j + 1).TrimStart
                                                Continue Do
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf arrayWordsUBound = -1 AndAlso isPrevStringDissociation AndAlso curString.Chars(j) = "-" Then
                                If CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_MATH OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                 CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse _
                                 CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse _
                                 CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_COMMA Then
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
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_MATH
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
                            If isInExec <> ExecBlockEnum.NO_EXEC AndAlso curString.Substring(j).StartsWith("</exec>", StringComparison.CurrentCultureIgnoreCase) Then
                                execOpenPos = curString.IndexOf("<exec>", j + 7, StringComparison.CurrentCultureIgnoreCase)
                                If execOpenPos = -1 Then
                                    arrayWords(arrayWordsUBound).Word = curString.Substring(j)
                                    arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_HTML_DATA
                                    isInExec = ExecBlockEnum.NO_EXEC
                                    Exit Do
                                Else
                                    arrayWords(arrayWordsUBound).Word = curString.Substring(j, execOpenPos - j + 6)
                                    arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_HTML_DATA
                                    j = execOpenPos + 5
                                    wordStart = j + 1
                                    Continue Do
                                End If
                            End If
                            arrayWords(arrayWordsUBound).Word = " " + strOperatorFull + " "
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE
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
                                ShowTextError(rtb, i, False)
                                ovalBracketBalance = 0
                                quadBracketBalance = 0
                                isPrevStringDissociation = False
                                Erase arrayWordsCopy
                                i += 1
                                Continue While
                            End If
                            'Обрабатываем экранированную кавычку /'
                            If closeQuotePos > 0 AndAlso curString.Chars(closeQuotePos - 1) = "/"c Then
                                If closeQuotePos <> curString.Length - 1 Then GoTo qbQuoteSearching2
                            End If
                            'закрывающая кавычка найдена
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, closeQuotePos - wordStart + 1)
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_SIMPLE_STRING
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
                                ShowTextError(rtb, i, False)
                                ovalBracketBalance = 0
                                quadBracketBalance = 0
                                isPrevStringDissociation = False
                                Erase arrayWordsCopy
                                i += 1
                                Continue While
                            End If
                            'Обрабатываем экранированную кавычку /"
                            If curString.Chars(closeQuotePos - 1) = "/"c Then
                                Mid(curString, closeQuotePos + 1, 1) = "'"
                                If closeQuotePos <> curString.Length - 1 Then GoTo qbQuoteSearching3
                            End If
                            'закрывающая кавычка найдена
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "'" + curString.Substring(wordStart + 1, closeQuotePos - wordStart - 1) + "'"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_SIMPLE_STRING
                            wordStart = closeQuotePos + 1
                            j = closeQuotePos
                        Case ";"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "; "
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION
                            wordStart = j + 1
                            If ovalBracketBalance <> 0 Then
                                mScript.LAST_ERROR = "Непарное кол-во открывающих и закрывающих круглых скобок."
                                ShowTextError(rtb, i, False)
                                ovalBracketBalance = 0
                                quadBracketBalance = 0
                                isPrevStringDissociation = False
                                Erase arrayWordsCopy
                                i += 1
                                Continue While
                            End If
                            If quadBracketBalance <> 0 Then
                                mScript.LAST_ERROR = "Непарное кол-во открывающих и закрывающих квадратных скобок."
                                ShowTextError(rtb, i, False)
                                ovalBracketBalance = 0
                                quadBracketBalance = 0
                                isPrevStringDissociation = False
                                Erase arrayWordsCopy
                                i += 1
                                Continue While
                            End If
                    End Select
                Loop
                If isInExec = ExecBlockEnum.EXEC_SINGLE Then isInExec = ExecBlockEnum.EXEC_MULTILINE

                'текущая строка заканчивается на _  - устанавливаем последнее слово = W_STRINGS_DISSOCIATION
                If arrayWords(arrayWordsUBound).Word = "_" Then
                    arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION
                ElseIf arrayWords(arrayWordsUBound).Word.EndsWith("_") Then
                    arrayWords(arrayWordsUBound).Word = arrayWords(arrayWordsUBound).Word.Substring(0, arrayWords(arrayWordsUBound).Word.Length - 1)
                    arrayWordsUBound += 1
                    arrayWords(arrayWordsUBound).Word = "_"
                    arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION
                End If

                ReDim Preserve arrayWords(arrayWordsUBound) 'убираем пустые значения в конце массива слов

                If arrayWords(arrayWordsUBound).wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                    'проверка парности скобок лишь в том случае, если эта строка не заканчивается на _ (или это незавершенная строка)
                    If ovalBracketBalance <> 0 Then
                        mScript.LAST_ERROR = "Непарное кол-во открывающих и закрывающих круглых скобок."
                        ShowTextError(rtb, i, False)
                        ovalBracketBalance = 0
                        quadBracketBalance = 0
                        isPrevStringDissociation = False
                        Erase arrayWordsCopy
                        i += 1
                        Continue While
                    End If
                    If quadBracketBalance <> 0 Then
                        mScript.LAST_ERROR = "Непарное кол-во открывающих и закрывающих квадратных скобок."
                        ShowTextError(rtb, i, False)
                        ovalBracketBalance = 0
                        quadBracketBalance = 0
                        isPrevStringDissociation = False
                        Erase arrayWordsCopy
                        i += 1
                        Continue While
                    End If
                End If

                'получаем значения оставшихся слов, расставляем пробелы и проверяем синтаксис
                If AnalizeArrayWords(arrayWords, isPrevStringDissociation, i) = -1 Then
                    ShowTextError(rtb, i, False)
                    ovalBracketBalance = 0
                    quadBracketBalance = 0
                    isPrevStringDissociation = False
                    Erase arrayWordsCopy
                    i += 1
                    Continue While
                End If
                CodeData(i).Code = arrayWords

                'собираем строку заново, уже с правильными пробелами и в правильном регистре
                curString = ""
                For j = 0 To arrayWords.GetUpperBound(0)
                    curString += arrayWords(j).Word
                Next
                'если текущая строка - начало блока HTML / Wrap - устанавливаем isInText = True
                If curString.StartsWith("HTML") OrElse curString.StartsWith("Wrap") Then
                    If isInExec <> ExecBlockEnum.NO_EXEC Then
                        mScript.LAST_ERROR = "Недопустимы внутренние блоки Wrap/HTML внури тэга <exec>."
                        ShowTextError(rtb, i, False)
                        ovalBracketBalance = 0
                        quadBracketBalance = 0
                        isPrevStringDissociation = False
                        Erase arrayWordsCopy
                        i += 1
                        Continue While
                    End If
                    If arrayWords(0).wordType = EditWordTypeEnum.W_HTML Then
                        isInText = TextBlockEnum.TEXT_HTML
                        isInExec = ExecBlockEnum.NO_EXEC
                        textStartLine = i + 1
                    ElseIf arrayWords(0).wordType = EditWordTypeEnum.W_WRAP Then
                        isInText = TextBlockEnum.TEXT_WRAP
                        isInExec = ExecBlockEnum.NO_EXEC
                        textStartLine = i + 1
                    End If
                End If

                curString = CodeData(i).StartingSpaces + curString + CodeData(i).Comments 'готовая строка со всеми правками
                'если исходная строка и строка со всеми правками не совпадают - вставляем исправленную строку в RichTextBox вместо исходной
                If rtb.Lines(i).Equals(curString) = False AndAlso rtb.csUndo.UndoInProcess = False Then
                    rtb.Select(rtb.GetFirstCharIndexFromLine(i), rtb.Lines(i).Length)
                    rtb.SelectedText = curString
                End If
                If drawText Then DrawWords(rtb, arrayWords, i, CodeData(i).StartingSpaces.Length, CodeData(i).Comments.Length > 0) 'разукрашиваем слова
                rtb.Select(selStart, selLength)

                'если строка заканцивается на _ - устанавливаем isPrevStringDissociation = True,
                If arrayWords(arrayWords.GetUpperBound(0)).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                    isPrevStringDissociation = True 'в новом витке цикла будет обозначать, что предыдущая строка закончилась на _
                    'если это не последняя строка в RichtextBox, но последняя в обрабатываемой выборке - расширяем выборку на 1, добавляя строку, следующуя за _
                    If endLine < rtb.Lines.GetUpperBound(0) AndAlso i = endLine Then
                        endLine += 1
                        If drawText Then
                            'если это последняя строка - сбрасываем ее цвет и стиль на значение по-умолчанию
                            rtb.Select(rtb.GetFirstCharIndexFromLine(endLine), rtb.Lines(endLine).Length)
                            rtb.SelectionColor = styleHash(EditWordTypeEnum.W_NOTHING).style_Color
                            rtb.SelectionFont = New Drawing.Font(rtb.Font, styleHash(EditWordTypeEnum.W_NOTHING).font_style)
                        End If
                    End If

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
                        If AnalizeArrayWords(arrayWordsCopy, False, i) = -1 Then
                            ShowTextError(rtb, i, False)
                            ovalBracketBalance = 0
                            quadBracketBalance = 0
                            isPrevStringDissociation = False
                            Erase arrayWordsCopy
                            i += 1
                            Continue While
                        End If
                    End If
                    isPrevStringDissociation = False
                    If IsNothing(arrayWordsCopy) = False Then Erase arrayWordsCopy
                End If
                i += 1
            End While
            rtb.csUndo.defaultFirstVisibleLine = -1 'убираем сохраненное значение верхней линии
        End Sub

        ''' <summary> Функция приводит текст из RichTextBox в вид, понятный движку (убирает/расставляет пробелы, приводит операторы в правильный регистр и т. д.,
        '''а также закрашивает слова в RichTextBox исходя из того, что они значат (цифры - в один цвет, операторы в другой и т. д.)
        ''' </summary>
        ''' <param name="rtb">RichTextBox, с которым работаем</param>
        ''' <param name="startLine">первая строка, с которой работаем</param>
        ''' <param name="endLine">последняя строка, с которой работаем (если -1, то последняя строка)</param>
        ''' <param name="drawText">раскрашивать ли текст в RichTextBox</param>
        ''' <param name="currentWordId">порядковый номер текущего слова (при редактировании строки) или -1, если не известно</param>
        Public Sub PrepareTextNonCompleted(ByRef rtb As RichTextBoxEx, Optional ByVal startLine As Integer = 0, Optional ByVal endLine As Integer = -1, _
                           Optional ByVal drawText As Boolean = True, Optional ByRef currentWordId As Integer = -1)

            currentWordId = -1
            If endLine = -1 Then endLine = rtb.Lines.GetUpperBound(0)
            mScript.LAST_ERROR = "" 'обнуляем ошибки
            PrepareHTMLforErrorInfo(wbRtbHelp)

            If rtb.Lines.GetUpperBound(0) = -1 AndAlso CodeData.GetUpperBound(0) = 0 Then 'проверка на пустой текст
                rtb.ForeColor = Color.Black
                CodeData(0) = New CodeDataType
                Return
            End If

            'проверяем структуру CodeData
            'Эта структура создается здесь, если подготавливается весть текст.
            'Если подготавливается лишь часть текста, то структура должа быть создана заранее (из всего текста)
            If startLine = 0 And (endLine = rtb.Lines.GetUpperBound(0) OrElse rtb.Lines.GetUpperBound(0) = -1) Then
                'подготавливается весь текст. Создаем структуру CodeData
                ReDim CodeData(rtb.Lines.GetUpperBound(0))
            ElseIf IsNothing(CodeData) Then
                'подготавливается только часть текста, а структура CodeData еще не создана - ошибка
                MsgBox("PrepareText: структура CodeData не подготовлена!", vbCritical)
                Return
            ElseIf CodeData.GetUpperBound(0) <> rtb.Lines.GetUpperBound(0) Then
                'подготавливается только часть текста, а размер CodeData не соответствует кол-ву строк в RichTextBox - ошибка
                startLine = 0
                endLine = rtb.Lines.GetUpperBound(0)
                ReDim CodeData(rtb.Lines.GetUpperBound(0))
                'MsgBox("PrepareText: структура CodeData не подготовлена!", vbCritical)
                'Return
            End If
            rtb.csUndo.SetCurrentUpperLineAsDefault() 'для корректной работы Undo сохраняем текущую верхнюю линию кодбокса

            Dim curString As String = "" 'текущая строка
            Dim i, j As Integer

            Dim selStart As Integer = rtb.SelectionStart 'начало выделения в RichTextBox
            Dim selLength As Integer = rtb.SelectionLength 'длина выделения в RichTextBox
            Dim arrayWords() As EditWordType = Nothing, arrayWordsUBound As Integer = 0 'массив, в котором храним слова из строки, ее размер
            Dim arrayWordsCopy() As EditWordType = Nothing  'массив для сборки строки, разбитой _
            Dim ovalBracketBalance As Integer = 0 'баланс ( и )
            Dim quadBracketBalance As Integer = 0 'баланс [ и ]
            Dim wordStart As Integer 'положение в строке первого символа слова, с которым будем работать
            Dim isPrevStringDissociation As Boolean 'разбита ли строка с помощью _
            Dim execOpenPos As Integer 'для нахождения тега <exec>
            Dim isInExec As ExecBlockEnum = ExecBlockEnum.NO_EXEC
            Dim isInText As TextBlockEnum = IsInTextBlock(rtb, startLine, isInExec)  'находимся ли мы внутри блока HTML или Wrap
            Dim textStartLine As Integer 'индекс первой строки с текстом, идущей после начала блока Wrap или HTML 
            If isInText <> TextBlockEnum.NO_TEXT_BLOCK Then textStartLine = startLine 'если начало текстового блока до начала кода - textStartLine = первой обрабатываемой строке

            i = startLine
            If endLine >= rtb.Lines.Length Then endLine = rtb.Lines.Length - 1
            While i <= endLine
                If drawText Then
                    'если это первый виток цикла - сбрасываем цвет и стиль текста на значение по-умолчанию
                    If i = startLine Then
                        rtb.Select(rtb.GetFirstCharIndexFromLine(startLine), rtb.GetFirstCharIndexFromLine(endLine) + rtb.Lines(endLine).Length - rtb.GetFirstCharIndexFromLine(startLine))
                        rtb.SelectionColor = styleHash(EditWordTypeEnum.W_NOTHING).style_Color
                        rtb.SelectionFont = New Drawing.Font(rtb.Font, styleHash(EditWordTypeEnum.W_NOTHING).font_style)
                        rtb.SelectionStart = selStart
                        rtb.SelectionLength = selLength
                    End If
                End If

                j = -1
                wordStart = 0
                If isInText <> TextBlockEnum.NO_TEXT_BLOCK Then
                    'мы внутри блока HTML / Wrap
                    curString = rtb.Lines(i)
                    Dim strLower As String = curString.Trim.Replace("  ", " ").ToLower
                    Dim strSearch As String = IIf(isInText = TextBlockEnum.TEXT_HTML, "end html", "end wrap")
                    If strLower.StartsWith(strSearch) AndAlso StripComments(strLower, Nothing, Nothing, True) <> 2 AndAlso strLower = strSearch Then
                        'StripComments <> 2 выполнено в блоке If Then только чтоб очистить curString от комментариев и пробелов перед третьей проверкой
                        'текущая строка - "End HTML" или "End Wrap" - конец блока
                        isInText = TextBlockEnum.NO_TEXT_BLOCK
                        isInExec = ExecBlockEnum.NO_EXEC
                        If drawText AndAlso i - 1 > textStartLine Then DrawHTML(rtb, textStartLine, i - 1)
                        rtb.SelectionStart = selStart
                        rtb.SelectionLength = selLength
                    Else
                        'продолжение текстового блока
                        If isInExec = ExecBlockEnum.NO_EXEC Then
                            execOpenPos = curString.IndexOf("<exec>", StringComparison.CurrentCultureIgnoreCase)
                            If execOpenPos > -1 Then
                                isInExec = ExecBlockEnum.EXEC_SINGLE
                            End If

                            If i = endLine And drawText Then DrawHTML(rtb, textStartLine, i) 'дошли до конца и не нашли конец блока - разрисовываем текстовый блок
                            rtb.SelectionStart = selStart
                            rtb.SelectionLength = selLength
                            'записываем текстовый блок в массив CodeData
                            ReDim CodeData(i).Code(0)
                            If isInExec <> ExecBlockEnum.NO_EXEC Then
                                execOpenPos += "<exec>".Length
                                ReDim arrayWords(20)
                                arrayWordsUBound = 0 'id последней заполненной ячейки массива
                                arrayWords(0).Word = curString.Substring(0, execOpenPos)
                                arrayWords(0).wordType = EditWordTypeEnum.W_HTML_DATA
                                j = execOpenPos - 1
                                wordStart = execOpenPos
                            Else
                                CodeData(i).Code(0).Word = rtb.Lines(i)
                                CodeData(i).Code(0).wordType = EditWordTypeEnum.W_HTML_DATA
                                i += 1
                                Continue While
                            End If
                        End If
                    End If
                End If

                'определяем, не заканчивается ли строка перед первой, которую мы обрабатываем, на _
                If i = startLine And i > 0 AndAlso isInExec <> ExecBlockEnum.EXEC_SINGLE Then
                    curString = rtb.Lines(i - 1)
                    If StripComments(curString, Nothing, Nothing, True) > -1 Then
                        If curString.Length > 0 AndAlso curString.EndsWith("_") Then
                            'если предыдущая строка - начало текущей (самой первой) - перезапускаем цикл, начиная с предыдущей строки
                            i -= 1
                            startLine = i
                            Continue While
                        End If
                    End If
                End If

                curString = rtb.Lines(i) 'получаем текущую линию

                'Очищаем строку от дополнительных элементов (пробелов в начале, комментариев...)
                If isInExec <> ExecBlockEnum.EXEC_SINGLE AndAlso StripComments(curString, CodeData(i).Comments, CodeData(i).StartingSpaces, True) = -1 Then
                    ShowTextError(rtb, i, True)
                    ovalBracketBalance = 0
                    quadBracketBalance = 0
                    isPrevStringDissociation = False
                    Erase arrayWordsCopy
                    i += 1
                    Continue While
                End If
                If IsNothing(CodeData(i).Comments) Then CodeData(i).Comments = ""
                If IsNothing(CodeData(i).StartingSpaces) Then CodeData(i).StartingSpaces = ""

                'curString = curString.TrimEnd(" "c, Chr(1), vbTab) 'убираем пробелы в конце
                'в строке были только пробелы и комментарии - переходим сразу к новой строке
                If curString.Length = 0 Then
                    Erase CodeData(i).Code
                    If drawText AndAlso CodeData(i).Comments.Length > 0 Then DrawWords(rtb, Nothing, i, CodeData(i).StartingSpaces.Length, True)
                    rtb.SelectionStart = selStart
                    rtb.SelectionLength = selLength
                    i += 1
                    Continue While
                End If

                'Получаем в массив arrayWords слова из рабочей строки
                If isInExec <> ExecBlockEnum.EXEC_SINGLE Then
                    ReDim arrayWords(20) 'расширяем массив слов сразу до Х, чтобы не делать это каждый раз при добавлении слова
                    arrayWordsUBound = -1 'id последней заполненной ячейки массива (-1 - пусто)
                End If

                Do
                    If arrayWordsUBound + 2 > arrayWords.GetUpperBound(0) Then ReDim Preserve arrayWords(arrayWordsUBound + 10) 'расширяем массив если было мало
                    'поиск первого попавшегося символа из массива stopCharArray:
                    'stopCharArray {" ", "(", ")", "[", "]", ".", ",", ,"'", "=", "+", "-", "*", "^", "/", "\", "<", ">", "&", "!", ";", "#", "$", """}
                    'Dim prevPos As Integer = j + 1
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
                            'If j > wordStart Then
                            '    'если между предыдущим и текущим оператором было какое-то слово - сохраняем его
                            '    arrayWordsUBound += 1 'будет заполнена новая ячейка массива
                            '    arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart) 'заполняем ячейку
                            'End If
                            'wordStart = j + 1 'начало следующего слова - на символ дальше

                            If j > wordStart Then
                                'если между предыдущим и текущим оператором было какое-то слово - сохраняем его
                                arrayWordsUBound += 1 'будет заполнена новая ячейка массива
                                For q As Integer = j + 1 To curString.Length - 1
                                    'вносим все пробелы
                                    If curString.Chars(q) = " "c Then
                                        j = j + 1
                                    Else
                                        Exit For
                                    End If
                                Next
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart + 1) 'заполняем ячейку
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
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN 'устанавливаем тип содержимого
                            wordStart = j + 1
                        Case ")"c
                            ovalBracketBalance -= 1
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = ")"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                            wordStart = j + 1
                        Case "["c
                            quadBracketBalance += 1
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "["
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                            wordStart = j + 1
                        Case "]"c
                            quadBracketBalance -= 1
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "]"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
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
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_POINT
                            arrayWords(arrayWordsUBound).classId = -2
                            wordStart = j + 1
                        Case ","c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = ","
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_COMMA
                            wordStart = j + 1
                        Case "="c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            wordStart = wordStart + j
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "=" ' " = "
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL
                            wordStart = j + 1
                        Case "?"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "?"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_RETURN
                            wordStart = j + 1
                        Case "&"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "&"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER
                            wordStart = j + 1
                        Case "#"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "#"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER
                            wordStart = j + 1
                        Case "$"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = "$"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING
                            wordStart = j + 1
                        Case "+"c, "-"c, "*"c, "^"c, "/"c, "\"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            Dim strOperatorFull As String = curString.Chars(j)
                            If j < curString.Length - 1 Then
                                Dim strUnar As String = curString.Substring(j + 1).TrimStart(" "c) 'в strUnar - строка за текущим символом, за вычетом левых пробелов и символа каретки
                                If strUnar.Length > 0 AndAlso strUnar.Chars(0) = "="c Then
                                    'получаем полный оператор, если это += или -=
                                    If curString.Chars(j) = "+"c OrElse curString.Chars(j) = "-"c Then strOperatorFull += "="
                                    j = curString.IndexOf("="c, j + 1) - 1
                                ElseIf curString.Chars(j) = "-"c Then
                                    If strUnar.Length > 0 AndAlso Integer.TryParse(strUnar.Chars(0), Nothing) = True Then
                                        If arrayWordsUBound > 0 AndAlso (arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse _
                                                                         arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_COMMA) Then
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
                                        If strUnar = "to" OrElse strUnar = "case" OrElse strUnar = "step" OrElse strUnar = "return" OrElse strUnar = "?" _
                                            OrElse IsMinusUnar(arrayWords, isPrevStringDissociation, i, arrayWordsUBound) Then
                                            'если за знаком "-" идет число, а перед ним - To Case или  Step
                                            'или же "-" это второй символ в строке, то 
                                            'это - унарный минус, который относится к следующей за ним цифре. Продолжаем цикл, не сохраняя минус отдельно
                                            wordStart = j
                                            curString = curString.Substring(0, j + 1) + curString.Substring(j + 1).TrimStart
                                            Continue Do
                                        End If
                                        If isPrevStringDissociation And j = 0 Then
                                            If CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_MATH OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_BLOCK_FOR OrElse _
                                                                         CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_COMMA Then
                                                'то же, только предыдущая строка заканчивалась на _, а минус - первый символ строки
                                                wordStart = j
                                                curString = curString.Substring(0, j + 1) + curString.Substring(j + 1).TrimStart
                                                Continue Do
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf arrayWordsUBound = -1 AndAlso isPrevStringDissociation AndAlso curString.Chars(j) = "-" Then
                                If CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_MATH OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE OrElse CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                 CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse _
                                 CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse _
                                 CodeData(i - 1).Code(CodeData(i - 1).Code.GetUpperBound(0) - 1).wordType = EditWordTypeEnum.W_COMMA Then
                                    'то же, только предыдущая строка заканчивалась на _, а минус - первый символ строки
                                    wordStart = j
                                    curString = curString.Substring(0, j + 1) + curString.Substring(j + 1).TrimStart
                                    Continue Do
                                End If
                            End If
                            arrayWordsUBound += 1
                            'If j = 0 Then
                            '    arrayWords(arrayWordsUBound).Word = strOperatorFull + " "
                            'Else
                            '    arrayWords(arrayWordsUBound).Word = " " + strOperatorFull + " "
                            'End If
                            arrayWords(arrayWordsUBound).Word = strOperatorFull
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_MATH
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
                            If isInExec <> ExecBlockEnum.NO_EXEC AndAlso curString.Substring(j).StartsWith("</exec>", StringComparison.CurrentCultureIgnoreCase) Then
                                execOpenPos = curString.IndexOf("<exec>", j + 7, StringComparison.CurrentCultureIgnoreCase)
                                If execOpenPos = -1 Then
                                    arrayWords(arrayWordsUBound).Word = curString.Substring(j)
                                    arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_HTML_DATA
                                    isInExec = ExecBlockEnum.NO_EXEC
                                    Exit Do
                                Else
                                    arrayWords(arrayWordsUBound).Word = curString.Substring(j, execOpenPos - j + 6)
                                    arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_HTML_DATA
                                    j = execOpenPos + 5
                                    wordStart = j + 1
                                    Continue Do
                                End If
                            End If
                            arrayWords(arrayWordsUBound).Word = strOperatorFull '" " + strOperatorFull + " "
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE
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
                                'закрывающий "'" не еще не введен
                                closeQuotePos = curString.Length - 1
                                'If selStart > 0 AndAlso rtb.Text.Chars(selStart - 1) = "'"c AndAlso lstRtb.Visible = True Then
                                '    closeQuotePos = Math.Min(curString.Length - 1, j)
                                'End If
                            Else
                                'Обрабатываем экранированную кавычку /'
                                If closeQuotePos > 0 AndAlso curString.Chars(closeQuotePos - 1) = "/"c Then
                                    If closeQuotePos <> curString.Length - 1 Then GoTo qbQuoteSearching2
                                End If
                            End If
                            'закрывающая кавычка найдена
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, closeQuotePos - wordStart + 1)
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_SIMPLE_STRING
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
                                'закрывающий " не еще не введен
                                closeQuotePos = curString.Length - 1
                            End If
                            ''Обрабатываем экранированную кавычку /"
                            'If curString.Chars(closeQuotePos - 1) = "/"c Then
                            '    Mid(curString, closeQuotePos + 1, 1) = "'"
                            '    If nonCompleted = False AndAlso closeQuotePos <> curString.Length - 1 Then GoTo qbQuoteSearching3
                            'End If
                            'закрывающая кавычка найдена
                            arrayWordsUBound += 1
                            If closeQuotePos - wordStart = 0 Then
                                arrayWords(arrayWordsUBound).Word = Chr(34)
                            Else
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, closeQuotePos - wordStart + 1)
                            End If
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_SIMPLE_STRING
                            wordStart = closeQuotePos + 1
                            j = closeQuotePos
                        Case ";"c
                            If j > wordStart Then
                                arrayWordsUBound += 1
                                arrayWords(arrayWordsUBound).Word = curString.Substring(wordStart, j - wordStart)
                            End If
                            arrayWordsUBound += 1
                            arrayWords(arrayWordsUBound).Word = ";"
                            arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION
                            wordStart = j + 1
                    End Select
endSel:
                    'If j > wordStart Then
                    '    'если между предыдущим и текущим оператором было какое-то слово - сохраняем его
                    For q As Integer = j + 1 To curString.Length - 1
                        'вносим все пробелы
                        If curString.Chars(q) = " "c Then
                            j = j + 1
                            wordStart = wordStart + 1
                            arrayWords(arrayWordsUBound).Word += " "
                        Else
                            Exit For
                        End If
                    Next
                    'End If
                    'wordStart = j + 1 'начало следующего слова - на символ дальше

                Loop
                If isInExec = ExecBlockEnum.EXEC_SINGLE Then isInExec = ExecBlockEnum.EXEC_MULTILINE

                'текущая строка заканчивается на _  - устанавливаем последнее слово = W_STRINGS_DISSOCIATION
                If arrayWords(arrayWordsUBound).Word.TrimEnd = "_" Then
                    arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION
                ElseIf arrayWords(arrayWordsUBound).Word.EndsWith("_") Then
                    arrayWords(arrayWordsUBound).Word = arrayWords(arrayWordsUBound).Word.Substring(0, arrayWords(arrayWordsUBound).Word.Length - 1)
                    arrayWordsUBound += 1
                    arrayWords(arrayWordsUBound).Word = "_"
                    arrayWords(arrayWordsUBound).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION
                End If

                ReDim Preserve arrayWords(arrayWordsUBound) 'убираем пустые значения в конце массива слов

                'получаем значения оставшихся слов, расставляем пробелы и проверяем синтаксис
                If AnalizeArrayWordsNonCompleted(arrayWords, isPrevStringDissociation, i) = -1 Then
                    ShowTextError(rtb, i, True)
                    ovalBracketBalance = 0
                    quadBracketBalance = 0
                    isPrevStringDissociation = False
                    Erase arrayWordsCopy
                    i += 1
                    Continue While
                End If
                CodeData(i).Code = arrayWords

                'собираем строку заново, уже с правильными пробелами и в правильном регистре
                curString = ""
                For j = 0 To arrayWords.GetUpperBound(0)
                    curString += arrayWords(j).Word
                Next
                'если текущая строка - начало блока HTML / Wrap - устанавливаем isInText = True
                If curString.StartsWith("HTML") OrElse curString.StartsWith("Wrap") Then
                    If isInExec <> ExecBlockEnum.NO_EXEC Then
                        mScript.LAST_ERROR = "Недопустимы внутренние блоки Wrap/HTML внури тэга <exec>."
                        ShowTextError(rtb, i, True)
                        ovalBracketBalance = 0
                        quadBracketBalance = 0
                        isPrevStringDissociation = False
                        Erase arrayWordsCopy
                        i += 1
                        Continue While
                    End If
                    If arrayWords(0).wordType = EditWordTypeEnum.W_HTML Then
                        isInText = TextBlockEnum.TEXT_HTML
                        isInExec = ExecBlockEnum.NO_EXEC
                        textStartLine = i + 1
                    ElseIf arrayWords(0).wordType = EditWordTypeEnum.W_WRAP Then
                        isInText = TextBlockEnum.TEXT_WRAP
                        isInExec = ExecBlockEnum.NO_EXEC
                        textStartLine = i + 1
                    End If
                End If

                'возвращаем в строку начальные пробелы и комментарии
                'Dim prevCanRaiseEventsValue As Boolean = rtbCanRaiseEvents
                'rtbCanRaiseEvents = False 'запрет событий в RichTextBox (т. к. текст в нем может меняться)

                'каретка не на этой линии или это завершенная строка
                curString = CodeData(i).StartingSpaces + curString + CodeData(i).Comments 'готовая строка со всеми правками
                'если исходная строка и строка со всеми правками не совпадают - вставляем исправленную строку в RichTextBox вместо исходной
                If rtb.Lines(i).Equals(curString) = False AndAlso rtb.csUndo.UndoInProcess = False Then
                    rtb.Select(rtb.GetFirstCharIndexFromLine(i), rtb.Lines(i).Length)
                    rtb.SelectedText = curString
                End If

                'опеределяем на котором слове стоит курсор
                If i = rtb.GetLineFromCharIndex(rtb.SelectionStart) Then
                    Dim prevLength As String = CodeData(i).StartingSpaces.Length + rtb.GetFirstCharIndexOfCurrentLine
                    For q As Integer = 0 To arrayWords.Count - 1
                        prevLength = prevLength + arrayWords(q).Word.Length
                        If selStart <= prevLength Then
                            'If (q = arrayWords.Count - 1 AndAlso (selStart > prevLength - arrayWords(q).Word.Length + arrayWords(q).Word.TrimEnd.Length + 1) OrElse _
                            '    arrayWords(q).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE OrElse arrayWords(q).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN) Then Exit For
                            If (q = arrayWords.Count - 1 AndAlso (selStart > prevLength - arrayWords(q).Word.Length + arrayWords(q).Word.TrimEnd.Length + 1) OrElse _
                                arrayWords(q).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE OrElse arrayWords(q).wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE) Then Exit For

                            currentWordId = q
                            Exit For
                        End If
                    Next
                End If


                If drawText Then DrawWords(rtb, arrayWords, i, CodeData(i).StartingSpaces.Length, CodeData(i).Comments.Length > 0) 'разукрашиваем слова
                rtb.SelectionStart = selStart
                rtb.SelectionLength = selLength
                'rtbCanRaiseEvents = prevCanRaiseEventsValue  'разрешение событий в RichTextBox (точнее, возвращаем то, что было раньше)


                'если строка заканцивается на _ - устанавливаем isPrevStringDissociation = True,
                If arrayWords(arrayWords.GetUpperBound(0)).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                    isPrevStringDissociation = True 'в новом витке цикла будет обозначать, что предыдущая строка закончилась на _
                    'если это не последняя строка в RichtextBox, но последняя в обрабатываемой выборке - расширяем выборку на 1, добавляя строку, следующуя за _
                    If endLine < rtb.Lines.GetUpperBound(0) AndAlso i = endLine Then
                        endLine += 1
                        If drawText Then
                            'если это последняя строка - сбрасываем ее цвет и стиль на значение по-умолчанию
                            rtb.Select(rtb.GetFirstCharIndexFromLine(endLine), rtb.Lines(endLine).Length)
                            rtb.SelectionColor = styleHash(EditWordTypeEnum.W_NOTHING).style_Color
                            rtb.SelectionFont = New Drawing.Font(rtb.Font, styleHash(EditWordTypeEnum.W_NOTHING).font_style)
                            rtb.SelectionStart = selStart
                            rtb.SelectionLength = selLength
                        End If
                    End If
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
                        If AnalizeArrayWordsNonCompleted(arrayWordsCopy, False, i) = -1 Then
                            ShowTextError(rtb, i, True)
                            ovalBracketBalance = 0
                            quadBracketBalance = 0
                            isPrevStringDissociation = False
                            Erase arrayWordsCopy
                            i += 1
                            Continue While
                        End If
                    End If
                    isPrevStringDissociation = False
                    If IsNothing(arrayWordsCopy) = False Then Erase arrayWordsCopy
                End If
                i += 1
            End While
            rtb.csUndo.defaultFirstVisibleLine = -1 'убираем сохраненное значение верхней линии
        End Sub

        Private Function IsMinusUnar(ByVal arrWords() As CodeTextBox.EditWordType, ByVal isStringDissociation As Boolean, ByVal curLine As Integer, ByVal lastWordPos As Integer) As Boolean
            AnalizeArrayWordsNonCompleted(arrWords, isStringDissociation, curLine)
            For i As Integer = lastWordPos To 0 Step -1
                If arrWords(i).wordType = EditWordTypeEnum.W_VARIABLE Then arrWords(i).wordType = EditWordTypeEnum.W_NOTHING
            Next
            'RemoveText -1 LW.RemoveText -1 / Then LW.RemoveText -1 / Then RemoveText -1 / LW[..].RemoveText -1 / Then LW[...].RemoveText -1

            Dim pos As Integer = lastWordPos
            If lastWordPos = 0 Then
                'RemoveText -1
                Return True
            ElseIf arrWords(pos).wordType <> EditWordTypeEnum.W_FUNCTION Then
                Return False
            End If

            Dim lineChanged As Boolean = False, initLine As Integer = curLine, arr() As CodeTextBox.EditWordType = arrWords
            Dim wrd As CodeTextBox.EditWordType = Nothing
            wrd = GetPrevWordWithWordArray(arr, curLine, pos, True)
            If curLine <> initLine Then arr = CodeData(curLine).Code

            If pos = 0 Then
                If IsNothing(wrd.Word) OrElse wrd.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then Return True
                'RemoveText -1 
            ElseIf IsNothing(wrd.Word) Then
                Return False
            ElseIf wrd.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                wrd = GetPrevWordWithWordArray(arr, curLine, pos, False)
                If curLine <> initLine Then arr = CodeData(curLine).Code
                If IsNothing(wrd.Word) Then Return False
            End If

            If String.Compare(wrd.Word.Trim, "Then", True) = 0 Then Return True 'Then RemoveText -1

            'LW. / Then LW. / LW[..]. / Then LW[...].
            If wrd.wordType <> EditWordTypeEnum.W_POINT Then Return False

            'LW / Then LW / LW[..] / Then LW[...]
            wrd = GetPrevWordWithWordArray(arr, curLine, pos, False)
            If IsNothing(wrd.Word) Then Return False

            If wrd.wordType = EditWordTypeEnum.W_CLASS Then
                'LW / Then LW
                If pos > 0 AndAlso String.Compare(arr(pos - 1).Word.Trim, "Then", True) = 0 Then
                    Return True 'Then LW
                ElseIf pos <> 0 Then
                    Return False
                Else
                    wrd = GetPrevWordWithWordArray(arr, curLine, pos, True)
                    If curLine <> initLine Then arr = CodeData(curLine).Code
                    If IsNothing(wrd.Word) OrElse wrd.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                        'LW
                        Return True
                    Else
                        Return False
                    End If
                End If
            ElseIf wrd.wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                'LW[..] / Then LW[...]
                Dim qb As Integer = 1
                Do
                    wrd = GetPrevWordWithWordArray(arr, curLine, pos, False)                    
                    If curLine <> initLine Then
                        If IsNothing(wrd.Word) = False AndAlso wrd.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then Return False
                        arr = CodeData(curLine).Code
                    End If
                    If IsNothing(wrd.Word) Then
                        Return False
                    ElseIf wrd.wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                        qb += 1
                    ElseIf wrd.wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                        qb -= 1
                        If qb = 0 Then Exit Do
                    End If
                Loop
                wrd = GetPrevWordWithWordArray(arr, curLine, pos, False)
                If curLine <> initLine Then arr = CodeData(curLine).Code
                If IsNothing(wrd.Word) OrElse wrd.wordType <> EditWordTypeEnum.W_CLASS Then Return False
                wrd = GetPrevWordWithWordArray(arr, curLine, pos, True)
                If curLine <> initLine Then arr = CodeData(curLine).Code
                If IsNothing(wrd.Word) OrElse wrd.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION OrElse String.Compare(wrd.Word.Trim, "Then", True) = 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function


        ''' <summary>
        ''' Возвращает предыдущее слово из CodeData (с учетом _) или Nothing. Изменяет lineId и wordId. Использует массив частично расшифрованных слов текущей строки из PrepateText
        ''' </summary>
        ''' <param name="lineId">текущая линия</param>
        ''' <param name="wordId">текущее слово</param>
        ''' <param name="ignoreStringDissociation">Если предыдущий символ _, то получать его или сразу перейти ко второму</param>
        Private Function GetPrevWordWithWordArray(ByRef arrWords() As CodeTextBox.EditWordType, ByRef lineId As Integer, ByRef wordId As Integer, ByVal ignoreStringDissociation As Boolean) As CodeTextBox.EditWordType
            Dim checkingArrWords As Boolean = True
1:
            If wordId > 0 Then
                wordId -= 1
                If checkingArrWords Then
                    Return arrWords(wordId)
                Else
                    Return CodeData(lineId).Code(wordId)
                End If
            ElseIf lineId = 0 OrElse IsNothing(CodeData(lineId - 1)) OrElse IsNothing(CodeData(lineId - 1).Code) OrElse CodeData(lineId - 1).Code.Count = 0 OrElse CodeData(lineId - 1).Code.Last.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                Return Nothing
            ElseIf CodeData(lineId - 1).Code.Count = 1 Then
                '_
                lineId -= 1
                wordId = 0
                checkingArrWords = False
                If Not ignoreStringDissociation Then Return CodeData(lineId).Code(wordId)
                GoTo 1
            Else
                lineId -= 1
                If ignoreStringDissociation Then
                    wordId = CodeData(lineId).Code.Count - 2
                Else
                    wordId = CodeData(lineId).Code.Count - 1
                End If
                Return CodeData(lineId).Code(wordId)
            End If
        End Function

        ''' <summary>
        ''' Возвращает предыдущее слово из CodeData (с учетом _) или Nothing. Изменяет lineId и wordId
        ''' </summary>
        ''' <param name="lineId">текущая линия</param>
        ''' <param name="wordId">текущее слово</param>
        ''' <param name="ignoreStringDissociation">Если предыдущий символ _, то получать его или сразу перейти ко второму</param>
        Public Function GetPrevWord(ByRef lineId As Integer, ByRef wordId As Integer, ByVal ignoreStringDissociation As Boolean) As CodeTextBox.EditWordType
1:
            If wordId > 0 Then
                wordId -= 1
                Return CodeData(lineId).Code(wordId)
            ElseIf lineId = 0 OrElse IsNothing(CodeData(lineId - 1)) OrElse IsNothing(CodeData(lineId - 1).Code) OrElse CodeData(lineId - 1).Code.Count = 0 OrElse CodeData(lineId - 1).Code.Last.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                Return Nothing
            ElseIf CodeData(lineId - 1).Code.Count = 1 Then
                '_
                lineId -= 1
                wordId = 0
                If Not ignoreStringDissociation Then Return CodeData(lineId).Code(wordId)
                GoTo 1
            Else
                lineId -= 1
                If ignoreStringDissociation Then
                    wordId = CodeData(lineId).Code.Count - 2
                Else
                    wordId = CodeData(lineId).Code.Count - 1
                End If
                Return CodeData(lineId).Code(wordId)
            End If
        End Function

        '        ''' <summary>
        '        ''' Возвращает следующее слово из CodeData (с учетом _) или Nothing. Изменяет lineId и wordId. Использует массив частично расшифрованных слов текущей строки из PrepateText
        '        ''' </summary>
        '        ''' <param name="lineId">текущая линия</param>
        '        ''' <param name="wordId">текущее слово</param>
        '        ''' <param name="ignoreStringDissociation">Если следующий символ _, то получать его или сразу перейти ко второму</param>
        '        Private Function GetNextWordWithWordArray(ByRef arrWords() As CodeTextBox.EditWordType, ByRef lineId As Integer, ByRef wordId As Integer, ByVal ignoreStringDissociation As Boolean) As CodeTextBox.EditWordType
        '            Dim checkingArrWords As Boolean = True
        '            Dim lastPos As Integer = 0

        '1:
        '            If wordId > 0 Then
        '                wordId -= 1
        '                If checkingArrWords Then
        '                    Return arrWords(wordId)
        '                Else
        '                    Return CodeData(lineId).Code(wordId)
        '                End If
        '            ElseIf lineId = 0 OrElse IsNothing(CodeData(lineId - 1)) OrElse CodeData(lineId - 1).Code.Count = 0 OrElse CodeData(lineId - 1).Code.Last.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
        '                Return Nothing
        '            ElseIf CodeData(lineId - 1).Code.Count = 1 Then
        '                '_
        '                lineId -= 1
        '                wordId = 0
        '                checkingArrWords = False
        '                If Not ignoreStringDissociation Then Return CodeData(lineId).Code(wordId)
        '                GoTo 1
        '            Else
        '                lineId -= 1
        '                If ignoreStringDissociation Then
        '                    wordId = CodeData(lineId).Code.Count - 2
        '                Else
        '                    wordId = CodeData(lineId).Code.Count - 1
        '                End If
        '                Return CodeData(lineId).Code(wordId)
        '            End If
        '        End Function

        ''' <summary>
        ''' Возвращает следующее слово из CodeData (с учетом _) или Nothing. Изменяет lineId и wordId
        ''' </summary>
        ''' <param name="lineId">текущая линия</param>
        ''' <param name="wordId">текущее слово</param>
        ''' <param name="ignoreStringDissociation">Если следующий символ _, то получать его или сразу перейти ко второму</param>
        Public Function GetNextWord(ByRef lineId As Integer, ByRef wordId As Integer, ByVal ignoreStringDissociation As Boolean) As CodeTextBox.EditWordType
            If IsNothing(CodeData) OrElse lineId > CodeData.Count - 1 Then Return Nothing
            If IsNothing(CodeData(lineId).Code) OrElse wordId > CodeData(lineId).Code.Count - 1 Then
                lineId += 1
                wordId = 0
                If IsNothing(CodeData) OrElse lineId > CodeData.Count - 1 OrElse IsNothing(CodeData(lineId).Code) OrElse wordId > CodeData(lineId).Code.Count - 1 Then Return Nothing
            End If
1:
            If wordId < CodeData(lineId).Code.Count - 1 Then
                wordId += 1
                If ignoreStringDissociation AndAlso CodeData(lineId).Code(wordId).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                    wordId += 1
                    GoTo 1
                End If
                Return CodeData(lineId).Code(wordId)
            ElseIf CodeData(lineId).Code.Last.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION OrElse lineId >= CodeData.Count - 1 OrElse IsNothing(CodeData(lineId + 1)) OrElse IsNothing(CodeData(lineId + 1).Code) OrElse CodeData(lineId + 1).Code.Count = 0 Then
                Return Nothing
            ElseIf CodeData(lineId + 1).Code.Count = 1 Then
                lineId += 1
                wordId = 0
                '_
                If CodeData(lineId).Code(wordId).wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION OrElse ignoreStringDissociation = False Then Return CodeData(lineId).Code(wordId)
                GoTo 1
            Else
                lineId += 1
                wordId = 0
                Return CodeData(lineId).Code(wordId)
            End If
        End Function


        ''' <summary>
        ''' Функция устанавливает цвет и стиль слов в RichTextBox в зависимости от значения этих слов. Предварительно стиль текста сброшен на сначение по-умолчанию</summary>
        ''' <param name="rtb">целевой RichTextBox</param>
        ''' <param name="arrWords">массив слов</param>
        ''' <param name="lineId">индекс линии в RichTextBox, которая обрабатывается</param>
        ''' <param name="startingSpacesLength">количество пробельных символов в начале строки</param>
        ''' <param name="isComments">есть ли комментарии в строке</param>
        Private Sub DrawWords(ByRef rtb As RichTextBox, ByRef arrWords() As EditWordType, ByVal lineId As Integer, ByVal startingSpacesLength As Integer, ByVal isComments As Boolean)
            If Not CanDrawWords Then Return
            'Индекс первого символа текущего слова. В начале - начало строки + начальные пробелы
            Dim posInLine As Integer = rtb.GetFirstCharIndexFromLine(lineId) + startingSpacesLength
            'Индекс первого символа, с которого надо начинать изменение стиля
            Dim drawingStart As Integer = posInLine
            Dim st As StylePresetType 'для получения стиля и цвета слова (в зависимости от его значения)
            Dim blnDrawWord As Boolean 'надо ли менять цвет /стиль текущего слова
            Dim default_st As StylePresetType = styleHash(EditWordTypeEnum.W_NOTHING) 'стиль по-умолчанию
            Dim cur_st As StylePresetType = default_st 'стиль фрагмента, который надо будет менять
            Dim isHTMLdata As Boolean = False

            If IsNothing(arrWords) = False Then 'если есть код
                For i As Integer = 0 To arrWords.GetUpperBound(0) 'перебираем все слова в цикле
                    'получаем стиль текущего слова
                    If arrWords(i).wordType = EditWordTypeEnum.W_CYCLE_END OrElse arrWords(i).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER _
                        OrElse arrWords(i).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING Then
                        'если текущее слово - "End" # или $ - берем стиль следующего слова 
                        If i < arrWords.GetUpperBound(0) Then
                            If styleHash.TryGetValue(arrWords(i + 1).wordType, st) = False Then
                                st = default_st
                            End If
                        Else
                            st = styleHash(EditWordTypeEnum.W_BLOCK_IF)
                        End If
                        blnDrawWord = True
                    ElseIf arrWords(i).wordType = EditWordTypeEnum.W_POINT Then
                        'если текущее слово - точка то это или Class. или [...].Function / Property
                        If i > 0 AndAlso arrWords(i - 1).wordType = EditWordTypeEnum.W_CLASS Then
                            'стиль точки = стилю класса
                            st = IIf(arrWords(i - 1).classId = -1, styleHash(EditWordTypeEnum.W_VARIABLE), styleHash(arrWords(i - 1).wordType))
                        Else
                            'стиль точки = стилю свойства / функции
                            If i < arrWords.GetUpperBound(0) Then
                                If styleHash.TryGetValue(arrWords(i + 1).wordType, st) = False Then
                                    st = styleHash(EditWordTypeEnum.W_PROPERTY)
                                End If
                            Else
                                st = styleHash(EditWordTypeEnum.W_PROPERTY)
                            End If
                        End If
                        blnDrawWord = True
                    ElseIf arrWords(i).wordType = EditWordTypeEnum.W_CLASS AndAlso arrWords(i).classId = -1 Then
                        'V / Var
                        st = styleHash(EditWordTypeEnum.W_VARIABLE)
                        blnDrawWord = True
                    ElseIf styleHash.TryGetValue(arrWords(i).wordType, st) = True Then 'получаем стиль текущего слова
                        blnDrawWord = True
                    Else
                        'такой тип не найден в таблице стилей styleHash. Закрашивание слова будет по-умолчанию
                        blnDrawWord = False
                        st = default_st
                    End If
                    If isHTMLdata = False Then
                        If arrWords(i).wordType = EditWordTypeEnum.W_HTML_DATA Then isHTMLdata = True
                    End If

                    If i = 0 And i < arrWords.GetUpperBound(0) Then
                        'если то первое слово (и не единственное) - сохраняем стиль фрагмента
                        cur_st = st
                    ElseIf cur_st.Equals(st) = False OrElse i = arrWords.GetUpperBound(0) Then
                        'стиль нового слова не равен стилю фрагмента или же это последнее слово
                        If drawingStart < posInLine AndAlso cur_st.Equals(default_st) = False Then
                            'если есть фрагмент кода, который надо закрасить - делаем это
                            rtb.Select(drawingStart, posInLine - drawingStart)
                            rtb.SelectionColor = cur_st.style_Color
                            rtb.SelectionFont = New Drawing.Font(rtb.Font.FontFamily, Convert.ToSingle(rtb.Font.Size), cur_st.font_style)
                        End If

                        drawingStart = posInLine 'устанавливаем новое значение начала фрагмента 
                        If blnDrawWord Then
                            'если стиль текущего слова не равен значению по-умолчанию:
                            If i = arrWords.GetUpperBound(0) Then
                                'если это последнее слово - меняем его стиль
                                rtb.Select(posInLine, arrWords(i).Word.Length)
                                rtb.SelectionColor = st.style_Color
                                rtb.SelectionFont = New Drawing.Font(rtb.Font.FontFamily, Convert.ToSingle(rtb.Font.Size), st.font_style)
                            Else
                                'это не последнее слово - сохраняем  его стиль в стиль фрагмента
                                cur_st = st
                            End If
                        Else
                            'стиль текущего слова равен значнию по-умолчанию - устанавливаем стиль фрагмента таким же
                            cur_st = default_st
                        End If
                    End If

                    posInLine += arrWords(i).Word.Length 'величиваем позицию в строке
                Next
            End If

            If isComments Then
                'закрашиваем комментарии
                rtb.Select(posInLine, rtb.GetFirstCharIndexFromLine(lineId) + rtb.Lines(lineId).Length - posInLine + 1)
                rtb.SelectionColor = styleHash(EditWordTypeEnum.W_COMMENTS).style_Color
                rtb.SelectionFont = New Drawing.Font(rtb.Font, styleHash(EditWordTypeEnum.W_COMMENTS).font_style)
            End If
            If isHTMLdata Then DrawHTML(rtb, lineId, lineId)

        End Sub

        ''' <summary>
        ''' Совершает поиск в html-содержимом
        ''' </summary>
        ''' <param name="txt"></param>
        ''' <param name="seekChars"></param>
        ''' <param name="startPos"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function LastIndexOAnyfNoQuotes(ByRef txt As String, ByVal seekChars() As Char, ByVal startPos As Integer) As Integer
            'Dim arr() As Char = {seekChar, "'"c, Chr(34)}
            'Dim pos As Integer = startPos
            'Do
            '    pos = txt.LastIndexOfAny(arr, pos)
            '    If pos = -1 Then Return -1
            '    Dim ch As Char = txt.Chars(pos)
            '    If ch = seekChar Then Return pos

            'Loop
            Return 0
        End Function

        ''' <summary> Переменная полулокального действия. Исходя из последнего введенного символа (события KePress и KeyDown) определяет надо ли разукрашивать html-блок до конца</summary>
        Public htmlDrawToEnd As Boolean = False
        Dim htmlCharArray() As Char = {"<"c, ">"c, "'"c, Chr(34), "#"} 'массив символов для поиска в IndexOfAny
        ''' <summary>Процедура раскрашивает HTML-текст в блоках Wrap и HTML</summary>
        ''' <param name="rtb">Целевой RichTextBox</param>
        ''' <param name="startLine">Начальная строка закрашивания</param>
        ''' <param name="endLine">Последняя строка закрашивания</param>
        Private Sub DrawHTML(ByRef rtb As RichTextBox, ByVal startLine As Integer, ByVal endLine As Integer)
            'Процедура раскрашивает HTML-текст в блоках Wrap & HTML
            Dim drawToEnd As Boolean = htmlDrawToEnd
            htmlDrawToEnd = False
            If Not CanDrawWords Then Return

            Dim stTags As StylePresetType = styleHash(EditWordTypeEnum.W_HTML_TAG) 'стиль тэгов
            Dim stComment As StylePresetType = styleHash(EditWordTypeEnum.W_COMMENTS) 'стиль комментариев
            Dim stText As StylePresetType = styleHash(EditWordTypeEnum.W_NOTHING) 'стиль обычного текста

            Dim startPos As Integer = rtb.GetFirstCharIndexFromLine(startLine) 'первый символ блока
            Dim endPos As Integer = rtb.GetFirstCharIndexFromLine(endLine) + rtb.Lines(endLine).Length 'последний символ блока
            Dim initialEndPos As Integer = endPos 'последний символ блока - копия, так как реальное значение endPos может измениться
            Dim curPos As Integer = startPos ' - 1 'текущая позиция
            Dim txtLength As Integer = rtb.TextLength
            Dim txt As String = rtb.Text 'весь текст из кодбокса

            '''''''''''''''I. При необходимости продлеваем endPos
            If endPos > txtLength - 1 Then
                'endPos - на последнем символе
                endPos = txtLength - 1
            ElseIf drawToEnd Then
                'установлен флаг прорисовки до конца блока. Ищем конец блока или селектор
                Dim final As Integer = FindEndHTMLorWrapOrSelector(rtb.GetLineFromCharIndex(curPos) + 1)
                If final = -1 Then
                    endPos = txtLength - 1 'конец не найден. Значит разрисовываем текст до конца
                Else
                    endPos = rtb.GetFirstCharIndexFromLine(final) - 2 'конец найден. Устанавливаем конец на конец строки перед завершением блока
                End If
            Else
                'Никаких особых флагов, каретка не в конце текста. Устанавливаем конец endPos на ближайший символ закрытия тэга
                Dim newEnd As Integer = txt.IndexOf(">"c, endPos)
                If newEnd = -1 Then
                    'Символов закрытия тэга не нашли. В таком случаем будет разрисовывать до конца блока
                    'Ищем конец блока или селектор
                    Dim final As Integer = FindEndHTMLorWrapOrSelector(rtb.GetLineFromCharIndex(endPos) + 1)
                    If final = -1 Then
                        final = txtLength - 1 'конец не найден. Значит разрисовываем текст до конца
                    Else
                        final = rtb.GetFirstCharIndexFromLine(final) - 1 '-2'конец найден. Устанавливаем конец на конец строки перед завершением блока
                    End If
                    endPos = final
                Else
                    endPos = newEnd 'Устанавливаем конец endPos на ближайший символ закрытия тэга
                End If
            End If

            '''''''''''''''II.А. Совершаем поиск от позиции posStart в направлении начала текста для определяния не находимся ли мы внутри открытых блоков комментариев или тэгов 
            Dim tagOpenPos As Integer = -1 'для хранения позиции символа открытия тэга <
            Dim tagClosed As Boolean = False 'закрыты ли тэги или мы внутри тэга
            Dim commOpenPos As Integer = curPos 'для хранения положения начала комментария <!--, а также для временного хранения положения последних найденных символов в цикле Do .. Loop
            Dim qPos As Integer 'положение ' / "

            If curPos > 0 Then
                Do
                    'Выполняем проверку от startPos в обратном направлении (к началу текста) на предмет наличия незакрытых комментариев и тэгов
                    If commOpenPos <= 0 Then
                        commOpenPos = -1
                        Exit Do
                    End If

                    'htmlCharArray = {"<"c, ">"c, "'"c, Chr(34), "#"} 
                    commOpenPos = txt.LastIndexOfAny(htmlCharArray, commOpenPos - 1)
                    If commOpenPos = -1 Then Exit Do
                    Dim ch As Char = txt.Chars(commOpenPos)
                    If ch = "'"c OrElse ch = Chr(34) Then
                        ''строка. Ищем ее начало строки и продолжаем
                        'qPos = txt.LastIndexOf(ch, commOpenPos - 1)
                        'If qPos > 0 Then
                        '    commOpenPos = qPos - 1
                        '    Continue Do  'в начале (перед строкой) может быть незакрытый тэг или коммент                        
                        'End If
                        ''строка с самого начала - ок
                        'Exit Do
                    ElseIf commOpenPos < txtLength - 5 AndAlso txt.Substring(commOpenPos, 4) = "<!--" Then
                        '<!--
                        'найдено начало комментария. Раньше проверять ничего не надо
                        Exit Do
                    ElseIf ch = "<"c Then
                        '< - начало тэга. Сохраняем положение начала
                        If Not tagClosed Then tagOpenPos = commOpenPos
                    ElseIf ch = "#"c Then
                        If mScript.IsItSelector(txt, commOpenPos) Then
                            '#1: найден селектор. Продолжение поиска в начало также не надо
                            commOpenPos = -1
                            Exit Do
                        End If
                    ElseIf commOpenPos <= txtLength - 3 AndAlso txt.Substring(commOpenPos - 2, 3) = "-->" Then
                        '--> закрытие комментария находится после открытия. Значит комментарий закрыт
                        commOpenPos = -1
                        Exit Do
                    ElseIf tagOpenPos = -1 Then
                        '> закрытие тэга найдено перед открытием. Все тэги закрыты
                        tagClosed = True
                    End If
                Loop
            Else
                commOpenPos = -1
            End If

            '''''''''''''''IIB. Если в IIA было выяснено, что мы внутри комментария / тэга, то ищеи их завершения и закрашиваем соответствующим образом.
            '''''''''''''''Если при этом мы вышли за пределы endPos, то завершаем процедуру, а если нет, то очищаем формат текта от initialEndPos до нового значения endPos и идем дальше в III
            If commOpenPos > -1 Then
                'найден открытый комментарий. Ищем его конец
                curPos = txt.IndexOf("-->", curPos)
                If curPos = -1 Then
                    'комментарий не закрыт. Надо закрасить зеленым все от начала коммента до конца текста или селектора или End HTML / Wrap
                    Dim final As Integer = FindEndHTMLorWrapOrSelector(rtb.GetLineFromCharIndex(commOpenPos) + 1)
                    If final = -1 Then
                        final = txtLength - 1 'конец не найден. Значит разрисовываем текст до конца
                    Else
                        final = rtb.GetFirstCharIndexFromLine(final) - 2 'конец найден. Устанавливаем конец на конец строки перед завершением блока
                    End If
                    rtb.Select(commOpenPos, final - commOpenPos + 1)
                    rtb.SelectionColor = stComment.style_Color
                    rtb.SelectionFont = New Drawing.Font(rtb.Font, stComment.font_style)
                    If final >= endPos Then
                        'VScrollPos = scrlPos
                        Return
                    End If
                    curPos = final
                Else
                    'комментарий закрыт. Надо закрасить зеленым все от начала коммента до селектора или End HTML / Wrap или конца коментария (что встретится раньше)
                    curPos += 3 '+ длина -->
                    Dim final As Integer = FindEndHTMLorWrapOrSelector(rtb.GetLineFromCharIndex(commOpenPos) + 1) 'получаем линию конца блока или начало нового селектора
                    If final = -1 Then
                        final = curPos 'конец блока / селектор не найден. Значит закрашиваем все до -->
                    Else
                        final = Math.Min(rtb.GetFirstCharIndexFromLine(final) - 1, curPos) 'найден конец блока /селектор. Получаем то, что встретилось раньше - конец блока или конец комментария -->
                    End If
                    rtb.Select(commOpenPos, final - commOpenPos + 1)
                    rtb.SelectionColor = stComment.style_Color
                    rtb.SelectionFont = New Drawing.Font(rtb.Font, stComment.font_style)
                    curPos = final
                    If curPos >= endPos Then
                        'VScrollPos = scrlPos
                        Return 'выход если мы за пределами endPos
                    End If
                End If
                'Здесь мы оказываемся в том случае, если не вышли за пределы endPos. Очищаем формат текста от исходного значения endPos = initialPos до EndPos
                rtb.Select(initialEndPos, endPos - initialEndPos + 1)
                rtb.SelectionColor = stText.style_Color
                rtb.SelectionFont = New Drawing.Font(rtb.Font, stText.font_style)
            ElseIf tagOpenPos > -1 Then
                'найден открытый тэг. Ищем закрывающий
                Dim tagEnd As Integer = tagOpenPos
                Do
                    If tagEnd > txtLength - 1 Then
                        tagEnd = -1
                        Exit Do
                    End If

                    tagEnd = txt.IndexOfAny({">"c, "'"c, Chr(34)}, tagEnd + 1)
                    If tagEnd = -1 Then Exit Do
                    Dim ch As Char = txt.Chars(tagEnd)
                    If ch <> ">"c Then
                        'строка
                        qPos = txt.IndexOf(ch, tagEnd + 1)
                        If qPos > 0 Then
                            tagEnd = qPos + 1
                            Continue Do
                        End If
                        'строка до самого конца. Закрывающего тэга не найдено
                        tagEnd = -1
                        Exit Do
                    Else
                        'найден >
                        Exit Do
                    End If
                Loop

                curPos = tagEnd
                If curPos = -1 Then
                    'тэг не закрыт. Надо закрасить синим все от начала коммента до конца текста или селектора или End HTML / Wrap
                    Dim final As Integer = FindEndHTMLorWrapOrSelector(rtb.GetLineFromCharIndex(tagOpenPos) + 1)
                    If final = -1 Then
                        final = txtLength - 1 'конец не найден. Значит разрисовываем текст до конца
                    Else
                        final = rtb.GetFirstCharIndexFromLine(final) - 2 'конец найден. Устанавливаем конец на конец строки перед завершением блока
                    End If
                    rtb.Select(tagOpenPos, final - tagOpenPos + 1)
                    rtb.SelectionColor = stTags.style_Color
                    rtb.SelectionFont = New Drawing.Font(rtb.Font, stTags.font_style)
                    If final >= endPos Then
                        'VScrollPos = scrlPos
                        Return
                    End If
                    curPos = final
                Else
                    'тэг закрыт. Надо закрасить синим все от начала тэга до селектора или End HTML / Wrap или конца тэга (что встретится раньше)
                    curPos += 1
                    Dim final As Integer = FindEndHTMLorWrapOrSelector(rtb.GetLineFromCharIndex(tagOpenPos) + 1) 'получаем линию конца блока или начало нового селектора
                    If final = -1 Then
                        final = curPos 'конец блока / селектор не найден. Значит закрашиваем все до >
                    Else
                        final = Math.Min(rtb.GetFirstCharIndexFromLine(final) - 2, curPos) 'найден конец блока /селектор. Получаем то, что встретилось раньше - конец блока или конец тэга >
                    End If
                    rtb.Select(tagOpenPos, final - tagOpenPos + 1)
                    rtb.SelectionColor = stTags.style_Color
                    rtb.SelectionFont = New Drawing.Font(rtb.Font, stTags.font_style)
                    If curPos >= endPos Then
                        'VScrollPos = scrlPos
                        Return 'выход если мы за пределами endPos
                    End If
                End If
                'Здесь мы оказываемся в том случае, если не вышли за пределы endPos. Очищаем формат текста от исходного значения endPos = initialPos до EndPos
                rtb.Select(initialEndPos, endPos - initialEndPos + 1)
                rtb.SelectionColor = stText.style_Color
                rtb.SelectionFont = New Drawing.Font(rtb.Font, stText.font_style)
            Else
                'мы не были ни в блоке комментариев, ни внутри тэга. Очищаем формат текста от исходного значения endPos = initialPos до EndPos
                curPos = startPos
                rtb.Select(initialEndPos, endPos - initialEndPos + 1)
                rtb.SelectionColor = stText.style_Color
                rtb.SelectionFont = New Drawing.Font(rtb.Font, stText.font_style)
                curPos = startPos - 1
            End If

            'Теперь мы находимся за пределами тэгов и комментариев, но конца проерки проверки endPos еще не достигли.
            '''''''''''''''III. Продвигаемся вперед, пока не дойдем до конца проверки, и закрашиваем текст в ссответствии с его значениями
            tagClosed = True
            Dim qChar As Char 'найдено ' или " ?
            Do ' While curPos <= endPos
                'htmlCharArray = {"<"c, ">"c, "'"c, Chr(34), "#"} 
                curPos = txt.IndexOfAny(htmlCharArray, curPos + 1)
                If curPos = -1 OrElse (curPos > endPos AndAlso tagClosed) Then Exit Do 'ничего больше не найдено или же мы дошли до endPos и при этом не находимся внутри тэга - выход
                Select Case txt.Chars(curPos)
                    Case "<"c
                        If txtLength > curPos + 3 AndAlso txt.Substring(curPos, 4) = "<!--" Then
                            'начало комментария.
                            commOpenPos = curPos
                            'Ищем конец
                            curPos = txt.IndexOf("-->", curPos + 3)
                            If curPos = -1 Then
                                'конец не найден - разукрашиваем зеленым все до конца текста или end html/wrap / selector
                                Dim final As Integer = FindEndHTMLorWrapOrSelector(rtb.GetLineFromCharIndex(commOpenPos) + 1) 'получаем линию конца блока или начало нового селектора
                                If final = -1 Then
                                    final = txtLength - 1 'конец не найден. Значит разрисовываем текст до конца
                                Else
                                    final = rtb.GetFirstCharIndexFromLine(final) - 2
                                End If
                                rtb.Select(commOpenPos, final - commOpenPos + 1)
                                rtb.SelectionColor = stComment.style_Color
                                rtb.SelectionFont = New Drawing.Font(rtb.Font, stComment.font_style)
                                If final >= endPos Then
                                    Return
                                End If
                                curPos = final
                            Else
                                curPos += 2
                                'комментарий закрыт. Надо закрасить зеленым все от начала коммента до селектора или End HTML / Wrap или конца коментария
                                Dim final As Integer = FindEndHTMLorWrapOrSelector(rtb.GetLineFromCharIndex(commOpenPos) + 1) 'получаем линию конца блока или начало нового селектора
                                If final = -1 Then
                                    final = curPos
                                Else
                                    final = Math.Min(rtb.GetFirstCharIndexFromLine(final) - 2, curPos) 'конец найден. Устанавливаем конец на конец строки перед завершением блока
                                End If
                                rtb.Select(commOpenPos, final - commOpenPos + 1)
                                rtb.SelectionColor = stComment.style_Color
                                rtb.SelectionFont = New Drawing.Font(rtb.Font, stComment.font_style)
                                If curPos >= endPos Then
                                    Return 'выход если мы за пределами endPos
                                End If
                            End If
                        Else
                            'начало тэга
                            If tagClosed Then tagOpenPos = curPos 'сохраняем позицию начала
                            tagClosed = False
                        End If
                    Case ">"c
                        If tagClosed = False Then
                            'конец тэга. Устанавливаем стиль тэга - закрашиваем его синим
                            tagClosed = True
                            rtb.Select(tagOpenPos, curPos - tagOpenPos + 1)
                            rtb.SelectionColor = stTags.style_Color
                            rtb.SelectionFont = New Drawing.Font(rtb.Font, stTags.font_style)
                            If curPos >= 5 AndAlso txt.Substring(curPos - 5).StartsWith("<exec>", StringComparison.CurrentCultureIgnoreCase) Then
                                'если это закрылся тэг <exec>, то ищем его конец </exec> и продолжаем уже за ним. Если конец <exec> за пределами endPos - выход
                                curPos = txt.IndexOf("</exec>", curPos + 1, System.StringComparison.CurrentCultureIgnoreCase)
                                If curPos = -1 OrElse curPos > endPos Then
                                    Return
                                End If
                                curPos -= 1
                            End If
                        End If
                    Case "#"
                        Dim sLength As Integer = mScript.IsItSelector(txt, curPos)
                        If sLength > 0 Then
                            'найден селектор #XX: Закрашиваем его соответственно
                            Dim stSelector As StylePresetType = styleHash(EditWordTypeEnum.W_RETURN) 'стиль #1:
                            rtb.Select(curPos, sLength)
                            rtb.SelectionColor = stSelector.style_Color
                            rtb.SelectionFont = New Drawing.Font(rtb.Font, stSelector.font_style)
                            If tagClosed = False Then
                                'Если до селектора был о=незакрытый тэг, то все от начала тэга и до селектора закрашиваем синим.
                                tagClosed = True
                                rtb.Select(tagOpenPos, curPos - tagOpenPos)
                                rtb.SelectionColor = stTags.style_Color
                                rtb.SelectionFont = New Drawing.Font(rtb.Font, stTags.font_style)
                            End If
                            curPos += 2
                        End If
                    Case "'"c, Chr(34)
                        If tagClosed = False Then
                            'строка. Ищем конец строки и продолжаем
                            qChar = txt.Chars(curPos) '  ' или "
                            qPos = txt.IndexOf(qChar, curPos + 1)
                            If qPos > -1 Then
                                curPos = qPos
                            Else
                                'конец строки не найден - выход из цикла
                                Exit Do
                            End If
                        End If
                End Select
            Loop


            If tagClosed = False Then
                'текст закончен незакрытым тэгом. Устанавливаем его стиль.
                rtb.Select(tagOpenPos, endPos - tagOpenPos + 1)
                rtb.SelectionColor = stTags.style_Color
                rtb.SelectionFont = New Drawing.Font(rtb.Font, stTags.font_style)
            End If
            'VScrollPos = scrlPos
        End Sub

        ''' <summary>
        ''' Находит конец блока Wrap / HTML за указанным положением и возвращает id линии или -1, если не найдено
        ''' </summary>
        ''' <param name="lineStart">Id линии, с которой начинать поиск</param>
        Private Function FindEndHTMLorWrapOrSelector(ByVal lineStart As Integer) As Integer
            If IsNothing(CodeData) OrElse lineStart > CodeData.Length - 1 Then Return -1
            Dim sStart As Integer
            For i As Integer = lineStart To CodeData.Length - 1
                Dim cod() As EditWordType = CodeData(i).Code
                If IsNothing(cod) OrElse cod.Length < 1 Then Continue For
                If (cod(0).wordType = EditWordTypeEnum.W_CYCLE_END AndAlso cod.Length >= 2 AndAlso (cod(1).wordType = EditWordTypeEnum.W_HTML OrElse cod(1).wordType = EditWordTypeEnum.W_WRAP)) Then
                    'OrElse (cod(0).wordType = EditWordTypeEnum.W_HTML_DATA AndAlso cod(0).Word.StartsWith("#") AndAlso IsCurWordSelector(cod(0).Word, 0) > 0) Then
                    Return i
                End If
                sStart = GetFirstCharIndexFromLine(i)
                If sStart <= TextLength - 3 AndAlso Text.Chars(sStart) = "#"c AndAlso mScript.IsItSelector(Text, sStart) > 0 Then
                    Return i
                End If
            Next
            Return -1
        End Function

        ''' <summary>
        ''' Устанавливает класс точке, если он по каким-либо причинам не установился ранее (такое возможно, когда функция находится на строчку выше точки | Class[... , _ ...]. |
        ''' Id класса получаем из имени класса (если таковой имеется). Если нет, то класс устанавливается равным 0
        ''' </summary>
        ''' <param name="arrayWords">массив обрабатываемых в данный момент слов в PrepareText</param>
        ''' <param name="pointId">Id точки</param>
        ''' <param name="curLine">линия точки</param>
        Private Sub SetPointClass(ByRef arrayWords() As EditWordType, ByVal pointId As Integer, ByVal curLine As Integer)
            Dim qbCount As Integer = 0
            'ищем в текущем ряде
            Dim wrd As EditWordType
            For i As Integer = pointId - 1 To 0 Step -1
                wrd = arrayWords(i)
                Select Case wrd.wordType
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount += 1
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount -= 1
                    Case EditWordTypeEnum.W_CLASS
                        If qbCount <> 0 Then Continue For
                        arrayWords(pointId).classId = wrd.classId
                        Return
                End Select
            Next i

            'если не нашли, то ищем в рядах выше
            'получаем последнее слово в строке выше
            curLine -= 1
            If curLine <= 0 OrElse IsNothing(CodeData(curLine).Code) OrElse CodeData(curLine).Code.Count = 0 Then
                arrayWords(pointId).classId = 0
                Return
            End If

            Dim pos As Integer = CodeData(curLine).Code.Count - 1
            wrd = CodeData(curLine).Code(pos)
            If wrd.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then 'если это не _, то выход
                arrayWords(pointId).classId = 0
                Return
            End If

            Do
                wrd = GetPrevWord(curLine, pos, True)
                Select Case wrd.wordType
                    Case EditWordTypeEnum.W_NOTHING
                        arrayWords(pointId).classId = 0
                        Return
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount += 1
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount -= 1
                    Case EditWordTypeEnum.W_CLASS
                        If qbCount <> 0 Then Continue Do
                        arrayWords(pointId).classId = wrd.classId
                        Return
                End Select
            Loop
        End Sub

        ''' <summary>
        ''' Функция получает массив arrayWords() из PrepareText  с частично полученными значениями слов и получает значения оставшихся слов, приводит их в правильный регистр
        ''' Предназначена еще не готовой строки (такой, которую юзер еще пишет)
        ''' </summary>
        ''' <param name="arrayWords">массив arrayWords() из PrepareText  с частично полученными значениями слов</param>
        ''' <param name="isStringDissociation">заканчивается ли предыдущая строка на _</param>
        ''' <param name="curLine">текущая линия</param>
        Private Function AnalizeArrayWordsNonCompleted(ByRef arrayWords() As EditWordType, ByRef isStringDissociation As Boolean, ByVal curLine As Integer) As Integer
            If IsNothing(arrayWords) OrElse arrayWords.GetUpperBound(0) = -1 Then Return 1

            Dim curWord As String 'для текущего слова
            Dim curWordNoSpaces As String
            Dim canWordBeExplained As Boolean = False 'может ли новое слово оцениваться как некий оператор или функция (или, может, пользователь написал только часть слова).
            'Например, он пишет переменную maxValue. То есть, начав печатать max он не имеет ввиду функцию Max
            Dim classId As Integer 'для получения Id класса текущего слова (если это класс, функция или свойство)
            Dim strLower As String 'для текущего слова в нижнем регистре
            Dim fp As New MatewScript.funcAndPropHashType

            Dim upperBound As Integer = arrayWords.GetUpperBound(0)
            For i As Integer = 0 To upperBound  'перебираем все слова по одному
                If arrayWords(i).wordType <> EditWordTypeEnum.W_NOTHING Then
                    If arrayWords(i).classId = -2 AndAlso arrayWords(i).wordType = EditWordTypeEnum.W_POINT Then
                        SetPointClass(arrayWords, i, curLine)
                    End If
                    Continue For 'слово уже известно - продолжаем
                End If
                curWord = arrayWords(i).Word
                If IsNothing(curWord) Then
                    upperBound = i - 1
                    Exit For
                End If
                curWordNoSpaces = curWord.Trim
                canWordBeExplained = curWord.EndsWith(" ")
                If canWordBeExplained = False AndAlso i < arrayWords.GetUpperBound(0) Then
                    If arrayWords(i + 1).wordType = EditWordTypeEnum.W_COMMA OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_CLASS OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER OrElse _
                        arrayWords(i + 1).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_FUNCTION OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_GLOBAL _
                        OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                        arrayWords(i + 1).wordType = EditWordTypeEnum.W_OPERATOR_LOGIC OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_OPERATOR_MATH OrElse _
                        arrayWords(i + 1).wordType = EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE OrElse _
                        arrayWords(i + 1).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_PARAM OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_PARAM_COUNT _
                        OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_PROPERTY OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_SIMPLE_BOOL OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER _
                        OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_SIMPLE_STRING OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION OrElse _
                        arrayWords(i + 1).wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_VARIABLE Then
                        canWordBeExplained = True
                    End If
                End If

                'Это просто число?
                If Double.TryParse(curWordNoSpaces, System.Globalization.NumberStyles.Any, provider_points, Nothing) Then
                    arrayWords(i).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER
                    Continue For
                End If
                'Это один из операторов (сохраненный в хэше операторов)?
                If canWordBeExplained AndAlso arrOperators.TryGetValue(curWordNoSpaces, arrayWords(i).wordType) Then
                    Mid(arrayWords(i).Word, arrayWords(i).Word.IndexOf(curWordNoSpaces, StringComparison.CurrentCultureIgnoreCase) + 1) = arrOperatorsKeys(curWordNoSpaces)
                    'If i = 1 AndAlso arrayWords(i).wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso String.Compare(arrayWords(i - 1).Word.Trim, "Case", True) = 0 _
                    '    AndAlso String.Compare(arrayWords(i).Word.Trim, "Else", True) = 0 Then
                    '    'это Case Else, Else относится не к W_BLOCK_IF, а к W_SWITCH
                    '    arrayWords(i).wordType = EditWordTypeEnum.W_SWITCH
                    'End If
                    Continue For
                End If
                'Это метка?
                If curWordNoSpaces.EndsWith(":") Then
                    If arrayWords.GetUpperBound(0) = 0 Then
                        arrayWords(i).wordType = EditWordTypeEnum.W_MARK
                        Continue For
                    Else
                        mScript.LAST_ERROR = "Не допускаются любые другие операторы в строке рядом с меткой."
                        Return -1
                    End If
                End If

                If i < arrayWords.GetUpperBound(0) Then
                    strLower = curWordNoSpaces.ToLower 'слово в нижнем регистре
                    'Это строка вида Class.Element / Class[...].Element?
                    If arrayWords(i + 1).wordType = EditWordTypeEnum.W_POINT Then
                        'Class.Element
                        If strLower = "var" OrElse strLower = "v" Then
                            'Var.varName
                            'сохраняем результат слова
                            arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                            Mid(arrayWords(i).Word, arrayWords(i).Word.IndexOf(curWordNoSpaces, StringComparison.CurrentCultureIgnoreCase) + 1) = IIf(strLower = "v", "V", "Var")
                            arrayWords(i).classId = -1 'для псевдокласса Var classId = -1
                            arrayWords(i + 1).classId = -1 'ставим к какому классу относится точка 
                            'устанавливаем, что varName - переменная
                            If i + 1 < arrayWords.GetUpperBound(0) Then
                                If arrayWords(i + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                    arrayWords(i + 2).wordType = EditWordTypeEnum.W_VARIABLE
                                    arrayWords(i + 2).classId = -1
                                    Continue For
                                Else
                                    If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or _
                                        arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                        arrayWords(i + 3).wordType = EditWordTypeEnum.W_VARIABLE
                                        arrayWords(i + 3).classId = -1
                                        Continue For
                                    End If
                                End If
                            End If
                        Else
                            If mScript.mainClassHash.TryGetValue(curWordNoSpaces, classId) Then
                                'имя класса найдено в хэше mainClassHash
                                'Class.Element. Сохраняем значения слова класса
                                arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = classId
                                arrayWords(i + 1).classId = classId  'ставим к какому классу относится точка 
                                'Ставим имя класса в правильном регистре
                                For Each strName As String In mScript.mainClass(classId).Names
                                    If strLower = strName.ToLower Then
                                        Mid(arrayWords(i).Word, arrayWords(i).Word.IndexOf(curWordNoSpaces, StringComparison.CurrentCultureIgnoreCase) + 1) = strName
                                        Exit For
                                    End If
                                Next
                                'определяем функцию / свойство найденного класса, следующего за ним через точку (Class.XXX)
                                If i + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(i + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                        arrayWords(i + 2).classId = classId
                                        If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(i + 2).Word.Trim, fp) Then
                                            'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                            arrayWords(i + 2).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                                            'arrayWords(i + 2).Word = fp.elementName
                                            Mid(arrayWords(i + 2).Word, arrayWords(i + 2).Word.IndexOf(arrayWords(i + 2).Word.Trim, StringComparison.CurrentCultureIgnoreCase) + 1) = fp.elementName

                                        Else
                                            'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                            'и свойство - не найдя в структуре проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or _
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                            arrayWords(i + 3).classId = classId
                                            If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(i + 3).Word.Trim, fp) Then
                                                'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                                arrayWords(i + 3).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                                                Mid(arrayWords(i + 3).Word, arrayWords(i + 3).Word.IndexOf(arrayWords(i + 3).Word.Trim, StringComparison.CurrentCultureIgnoreCase) + 1) = fp.elementName
                                            Else
                                                'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                                arrayWords(i + 3).wordType = EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                                'и свойство - не найдя в структуре проверить точно невозможно
                                            End If
                                            Continue For
                                        End If
                                    End If
                                End If
                            Else
                                'имя класса НЕ найдено в хэше mainClassHash (возможно, создан динамически с New Class)
                                arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = -2 'неизвестный класс (возможно, создан динамически с New Class)
                                arrayWords(i + 1).classId = -2 'ставим к какому классу относится точка 
                                If i + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(i + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                        'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                        arrayWords(i + 2).classId = -2 'неизвестный класс
                                        If arrayWords.GetUpperBound(0) > i + 2 AndAlso (arrayWords(i + 3).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                            arrayWords(i + 3).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_PROPERTY 'это свойство
                                        Else
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                            'и свойство - не зная класса проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or _
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                            'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                            arrayWords(i + 3).classId = -2 'неизвестный класс
                                            If arrayWords.GetUpperBound(0) > i + 3 AndAlso (arrayWords(i + 4).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                                                                            arrayWords(i + 4).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                                arrayWords(i + 3).wordType = EditWordTypeEnum.W_PROPERTY 'это свойство
                                            Else
                                                arrayWords(i + 3).wordType = EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                                'и свойство - не зная класса проверить точно невозможно
                                            End If
                                            Continue For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    ElseIf arrayWords(i + 1).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                        'myVar[...], Code[...].X

                        'ищем закрывающую ], за которой может стоять точка - и тогда за ней - функция или свойство, или что-либо другое (тогда это перменная)
                        Dim qbBalance As Integer = 1, qbClosePos As Integer
                        For j As Integer = i + 2 To arrayWords.GetUpperBound(0)
                            If arrayWords(j).wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                qbBalance -= 1
                                If qbBalance = 0 Then
                                    qbClosePos = j
                                    Exit For
                                End If
                            ElseIf arrayWords(j).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                                qbBalance += 1
                            End If
                        Next

                        If qbBalance <> 0 Then
                            'если это незавершенная строка и ] не найдена 
                            If mScript.mainClassHash.TryGetValue(curWordNoSpaces, classId) Then
                                'Class[
                                arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = classId
                                'Ставим имя класса в правильном регистре
                                For Each strName As String In mScript.mainClass(classId).Names
                                    If strLower = strName.ToLower Then
                                        Mid(arrayWords(i).Word, arrayWords(i).Word.IndexOf(curWordNoSpaces, StringComparison.CurrentCultureIgnoreCase) + 1) = strName
                                        Exit For
                                    End If
                                Next
                            Else
                                'varName[
                                arrayWords(i).wordType = EditWordTypeEnum.W_VARIABLE
                                Continue For
                            End If
                        ElseIf qbClosePos = arrayWords.GetUpperBound(0) OrElse arrayWords(qbClosePos + 1).wordType <> EditWordTypeEnum.W_POINT Then
                            'это переменная (за [] нет точки) или Param
                            If String.Compare(arrayWords(i).Word.Trim, "Param", True) = 0 Then
                                arrayWords(i).wordType = EditWordTypeEnum.W_PARAM
                            Else
                                arrayWords(i).wordType = EditWordTypeEnum.W_VARIABLE
                            End If
                            arrayWords(i).classId = -1
                            Continue For
                        Else
                            'это функция или свойство
                            If mScript.mainClassHash.TryGetValue(curWordNoSpaces, classId) Then
                                'имя класса найдено в хэше mainClassHash
                                arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = classId
                                arrayWords(qbClosePos + 1).classId = classId  'ставим к какому классу относится точка 
                                'Ставим имя класса в правильном регистре
                                For Each strName As String In mScript.mainClass(classId).Names
                                    If strLower = strName.ToLower Then
                                        Mid(arrayWords(i).Word, arrayWords(i).Word.IndexOf(curWordNoSpaces, StringComparison.CurrentCultureIgnoreCase) + 1) = strName
                                        Exit For
                                    End If
                                Next
                                'определяем функцию / свойство найденного класса, следующего за ним за закрывающей ] (Class[...].XXX)
                                If qbClosePos + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                        arrayWords(qbClosePos + 2).classId = classId
                                        If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(qbClosePos + 2).Word.Trim, fp) Then
                                            'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                            arrayWords(qbClosePos + 2).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                                            Mid(arrayWords(qbClosePos + 2).Word, arrayWords(qbClosePos + 2).Word.IndexOf(arrayWords(qbClosePos + 2).Word.Trim, StringComparison.CurrentCultureIgnoreCase) + 1) = fp.elementName
                                        Else
                                            'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                            arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                            'и свойство - не найдя в структуре проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or _
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                            arrayWords(qbClosePos + 3).classId = classId
                                            If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(qbClosePos + 3).Word.Trim, fp) Then
                                                'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                                arrayWords(qbClosePos + 3).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                                                Mid(arrayWords(qbClosePos + 3).Word, arrayWords(qbClosePos + 3).Word.IndexOf(arrayWords(qbClosePos + 3).Word.Trim, StringComparison.CurrentCultureIgnoreCase) + 1) = fp.elementName
                                            Else
                                                'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                                arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                                'и свойство - не найдя в структуре проверить точно невозможно
                                            End If
                                            Continue For
                                        End If
                                    End If
                                End If
                            Else
                                'имя класса НЕ найдено в хэше mainClassHash (возможно, создан динамически с New Class)
                                arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = -2 'неизвестный класс (возможно, создан динамически с New Class)
                                arrayWords(qbClosePos + 1).classId = -2  'ставим к какому классу относится точка 
                                If qbClosePos + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                        'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                        arrayWords(qbClosePos + 2).classId = -2 'неизвестный класс
                                        If arrayWords.GetUpperBound(0) > qbClosePos + 2 AndAlso (arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                            arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_PROPERTY 'это свойство
                                        Else
                                            arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                            'и свойство - не зная класса проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or _
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                            'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                            arrayWords(qbClosePos + 3).classId = -2 'неизвестный класс
                                            If arrayWords.GetUpperBound(0) > qbClosePos + 3 AndAlso (arrayWords(qbClosePos + 4).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse _
                                                                                                     arrayWords(qbClosePos + 4).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                                arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_PROPERTY 'это свойство
                                            Else
                                                arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                                'и свойство - не зная класса проверить точно невозможно
                                            End If
                                            Continue For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                'неизвестное слово. Это функция или свойство без имени класса, или переменная
                'ищем в функциях и свойствах

                If canWordBeExplained Then
                    For j As Integer = 0 To mScript.mainClass.GetUpperBound(0)
                        If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(j).Names(0) + "_" + curWordNoSpaces, fp) Then
                            arrayWords(i).classId = j
                            arrayWords(i).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                            Mid(arrayWords(i).Word, arrayWords(i).Word.IndexOf(arrayWords(i).Word.Trim, StringComparison.CurrentCultureIgnoreCase) + 1) = fp.elementName
                            Exit For
                        End If
                    Next
                End If

                If arrayWords(i).wordType = EditWordTypeEnum.W_NOTHING Then
                    'это переменная (или Param)
                    If String.Compare(arrayWords(i).Word.Trim, "Param", True) = 0 Then
                        arrayWords(i).wordType = EditWordTypeEnum.W_PARAM
                    Else
                        arrayWords(i).wordType = EditWordTypeEnum.W_VARIABLE
                    End If
                    arrayWords(i).classId = -1
                End If
            Next

            'Расставляем правильно пробелы и проверяем на ошибки синтаксиса
            Dim isFirstWord As Boolean, isLastWord As Boolean 'текущее слово первое или последнее? (напр., после Then слово считается первым)
            'перебираем в цикле все слова (кроме последнего)
            For i = 0 To upperBound - 1
                If i = 0 OrElse arrayWords(i - 1).wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION OrElse arrayWords(i - 1).wordType = EditWordTypeEnum.W_HTML_DATA OrElse _
                    arrayWords(i - 1).Word.TrimStart.ToLower.StartsWith("then ") Then
                    'если перед текущим словом стоит Then или ; или <exec> - слово первое
                    If i = 0 And isStringDissociation = True Then
                        If IsNothing(CodeData(curLine - 1).Code) = False AndAlso IsNothing(CodeData(curLine).Code) = False AndAlso CodeData(curLine).Code.GetUpperBound(0) > 0 AndAlso _
                            CodeData(curLine).Code(CodeData(curLine).Code.GetUpperBound(0) - 1).Word.TrimStart.ToLower.StartsWith("then ") Then
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
                If (arrayWords(i + 1).wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION OrElse arrayWords(i + 1).Word.Trim.ToLower = "then" OrElse _
                    arrayWords(i + 1).wordType = EditWordTypeEnum.W_HTML_DATA) = False Then
                    isLastWord = False
                End If
            Next
            Return 1
        End Function

        ''' <summary>
        ''' Функция получает массив arrayWords() из PrepareText  с частично полученными значениями слов и получает значения оставшихся слов, приводит их в правильный регистр
        ''' Предназначена для уже готовой строки (а не такой, которую юзер еще пишет)
        ''' </summary>
        ''' <param name="arrayWords">массив arrayWords() из PrepareText  с частично полученными значениями слов</param>
        ''' <param name="isStringDissociation">заканчивается ли предыдущая строка на _</param>
        ''' <param name="curLine">текущая линия</param>
        Private Function AnalizeArrayWords(ByRef arrayWords() As EditWordType, ByRef isStringDissociation As Boolean, ByVal curLine As Integer) As Integer
            If IsNothing(arrayWords) OrElse arrayWords.GetUpperBound(0) = -1 Then Return 1

            Dim curWord As String 'для текущего слова
            Dim classId As Integer 'для получения Id класса текущего слова (если это класс, функция или свойство)
            Dim strLower As String 'для текущего слова в нижнем регистре
            Dim fp As New MatewScript.funcAndPropHashType

            For i As Integer = 0 To arrayWords.GetUpperBound(0) 'перебираем все слова по одному
                If arrayWords(i).wordType <> EditWordTypeEnum.W_NOTHING Then Continue For 'слово уже известно - продолжаем

                curWord = arrayWords(i).Word
                'Это просто число?
                If Double.TryParse(curWord, System.Globalization.NumberStyles.Any, provider_points, Nothing) Then
                    arrayWords(i).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER
                    Continue For
                End If
                'Это один из операторов (сохраненный в хэше операторов)?
                If arrOperators.TryGetValue(curWord, arrayWords(i).wordType) Then
                    arrayWords(i).Word = arrOperatorsKeys(curWord)
                    If arrayWords(i).wordType = EditWordTypeEnum.W_OPERATOR_LOGIC Then arrayWords(i).Word = " " & arrayWords(i).Word & " "
                    'If i = 1 AndAlso arrayWords(i).wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso String.Compare(arrayWords(i - 1).Word.Trim, "Case", True) = 0 _
                    'AndAlso String.Compare(arrayWords(i).Word.Trim, "Else", True) = 0 Then
                    '    'это Case Else, Else относится не к W_BLOCK_IF, а к W_SWITCH
                    '    arrayWords(i).wordType = EditWordTypeEnum.W_SWITCH
                    'End If

                    Continue For
                End If
                'Это метка?
                If curWord.EndsWith(":") Then
                    If arrayWords.GetUpperBound(0) = 0 Then
                        arrayWords(i).wordType = EditWordTypeEnum.W_MARK
                        Continue For
                    Else
                        mScript.LAST_ERROR = "Не допускаются любые другие операторы в строке рядом с меткой."
                        Return -1
                    End If
                End If

                If i < arrayWords.GetUpperBound(0) Then
                    strLower = curWord.ToLower 'слово в нижнем регистре
                    'Это строка вида Class.Element / Class[...].Element?
                    If arrayWords(i + 1).wordType = EditWordTypeEnum.W_POINT Then
                        'Class.Element
                        If strLower = "var" OrElse strLower = "v" Then
                            'Var.varName
                            'сохраняем результат слова
                            arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                            arrayWords(i).Word = IIf(strLower = "v", "V", "Var")
                            arrayWords(i).classId = -1 'для псевдокласса Var classId = -1
                            arrayWords(i + 1).classId = -1 'ставим к какому классу относится точка 
                            'устанавливаем, что varName - переменная
                            If i + 1 < arrayWords.GetUpperBound(0) Then
                                If arrayWords(i + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                    arrayWords(i + 2).wordType = EditWordTypeEnum.W_VARIABLE
                                    arrayWords(i + 2).classId = -1
                                    Continue For
                                Else
                                    If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                        arrayWords(i + 3).wordType = EditWordTypeEnum.W_VARIABLE
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
                                arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
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
                                    If arrayWords(i + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                        arrayWords(i + 2).classId = classId
                                        If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(i + 2).Word, fp) Then
                                            'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                            arrayWords(i + 2).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                                            arrayWords(i + 2).Word = fp.elementName
                                        Else
                                            'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                            'и свойство - не найдя в структуре проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                            arrayWords(i + 3).classId = classId
                                            If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(i + 3).Word, fp) Then
                                                'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                                arrayWords(i + 3).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                                                arrayWords(i + 3).Word = fp.elementName
                                            Else
                                                'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                                arrayWords(i + 3).wordType = EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
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
                                arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = -2 'неизвестный класс (возможно, создан динамически с New Class)
                                arrayWords(i + 1).classId = -2 'ставим к какому классу относится точка 
                                If i + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(i + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                        'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                        arrayWords(i + 2).classId = -2 'неизвестный класс
                                        If arrayWords.GetUpperBound(0) > i + 2 AndAlso (arrayWords(i + 3).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse arrayWords(i + 3).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_PROPERTY 'это свойство
                                        Else
                                            arrayWords(i + 2).wordType = EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                            'и свойство - не зная класса проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                            'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                            arrayWords(i + 3).classId = -2 'неизвестный класс
                                            If arrayWords.GetUpperBound(0) > i + 3 AndAlso (arrayWords(i + 4).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse arrayWords(i + 4).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                                arrayWords(i + 3).wordType = EditWordTypeEnum.W_PROPERTY 'это свойство
                                            Else
                                                arrayWords(i + 3).wordType = EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
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
                    ElseIf arrayWords(i + 1).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                        'myVar[...], Code[...].X

                        If (strLower = "var" OrElse strLower = "v") Then
                            mScript.LAST_ERROR = "Неверная запись переменной-массива. Должно быть " + curWord + ".myVar[x]"
                            Return -1
                        End If

                        'ищем закрывающую ], за которой может стоять точка - и тогда за ней - функция или свойство, или что-либо другое (тогда это перменная)
                        Dim qbBalance As Integer = 1, qbClosePos As Integer = -1, qbCloseLine As Integer = curLine
                        Dim lCopy As Integer = curLine, wCopy As Integer = arrayWords.GetUpperBound(0)
                        Dim wrd As EditWordType

                        For j As Integer = i + 2 To arrayWords.GetUpperBound(0)
                            If arrayWords(j).wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                qbBalance -= 1
                                If qbBalance = 0 Then
                                    qbClosePos = j
                                    Exit For
                                End If
                            ElseIf arrayWords(j).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                                qbBalance += 1
                            End If
                        Next

                        If qbClosePos < 0 AndAlso arrayWords.Last.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                            Do
                                wrd = GetNextWord(lCopy, wCopy, True)
                                Select Case wrd.wordType
                                    Case EditWordTypeEnum.W_NOTHING
                                        Exit Do
                                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                                        qbBalance -= 1
                                        If qbBalance = 0 Then
                                            qbClosePos = wCopy
                                            qbCloseLine = lCopy
                                            Exit Do
                                        End If
                                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                                        qbBalance += 1
                                End Select
                            Loop
                        End If
                        If curLine = lCopy AndAlso qbClosePos > arrayWords.Count - 1 Then Exit For 'ошибка (бывает при переносе строк)

                        Dim wrdNext As New EditWordType
                        Dim nextLinePos As Integer, nextWordPos As Integer
                        If qbClosePos >= 0 Then
                            If lCopy = curLine Then
                                wrd = arrayWords(qbClosePos)
                                If qbClosePos = arrayWords.Count - 1 Then
                                    Dim lc As Integer = lCopy, wc As Integer = wCopy
                                    wrdNext = GetNextWord(lc, wc, True)
                                    nextLinePos = lc
                                    nextWordPos = wc
                                Else
                                    wrdNext = arrayWords(qbClosePos + 1)
                                    nextLinePos = curLine
                                    nextWordPos = qbClosePos + 1
                                End If
                            Else
                                wrd = CodeData(lCopy).Code(wCopy)
                                Dim lc As Integer = lCopy, wc As Integer = wCopy
                                wrdNext = GetNextWord(lc, wc, True) 'qbClosePos + 1
                                nextLinePos = lc
                                nextWordPos = wc
                            End If
                        End If

                        If wrdNext.wordType = EditWordTypeEnum.W_NOTHING OrElse wrdNext.wordType <> EditWordTypeEnum.W_POINT Then
                            'If qbClosePos = arrayWords.GetUpperBound(0) OrElse arrayWords(qbClosePos + 1).wordType <> EditWordTypeEnum.W_POINT Then
                            'это переменная (за [] нет точки)
                            arrayWords(i).wordType = EditWordTypeEnum.W_VARIABLE
                            arrayWords(i).classId = -1
                            Continue For
                        Else
                            'это функция или свойство
                            If mScript.mainClassHash.TryGetValue(curWord, classId) Then
                                'имя класса найдено в хэше mainClassHash
                                arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = classId
                                If nextLinePos = curLine Then
                                    arrayWords(qbClosePos + 1).classId = classId  'ставим к какому классу относится точка 
                                ElseIf wrdNext.wordType <> EditWordTypeEnum.W_NOTHING Then
                                    CodeData(nextLinePos).Code(nextWordPos).classId = classId
                                End If
                                'Ставим имя класса в правильном регистре
                                For Each strName As String In mScript.mainClass(classId).Names
                                    If strLower = strName.ToLower Then
                                        arrayWords(i).Word = strName
                                        Exit For
                                    End If
                                Next
                                'определяем функцию / свойство найденного класса, следующего за ним за закрывающей ] (Class[...].XXX)
                                Dim lc As Integer = nextLinePos, wc As Integer = nextWordPos
                                Dim wrdNext2 As EditWordType
                                Dim nextLinePos2 As Integer, nextWordPos2 As Integer
                                If nextLinePos = curLine AndAlso qbClosePos + 1 < arrayWords.GetUpperBound(0) Then
                                    If qbClosePos + 2 <= arrayWords.Count - 1 Then
                                        wrdNext2 = arrayWords(qbClosePos + 2)
                                        nextLinePos2 = curLine
                                        nextWordPos2 = qbClosePos + 2
                                    Else
                                        wrdNext2 = GetNextWord(lc, wc, True)
                                        nextLinePos2 = lc
                                        nextWordPos2 = wc
                                    End If
                                Else
                                    wrdNext2 = GetNextWord(lc, wc, True) 'qbClosePos + 2
                                    nextLinePos2 = lc
                                    nextWordPos2 = wc
                                End If

                                'If qbClosePos + 1 < arrayWords.GetUpperBound(0) Then
                                If curLine = nextLinePos2 AndAlso IsNothing(wrdNext2.Word) = False Then
                                    If arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                        arrayWords(qbClosePos + 2).classId = classId
                                        If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(qbClosePos + 2).Word, fp) Then
                                            'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                            arrayWords(qbClosePos + 2).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                                            arrayWords(qbClosePos + 2).Word = fp.elementName
                                        Else
                                            'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                            arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                            'и свойство - не найдя в структуре проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                            arrayWords(qbClosePos + 3).classId = classId
                                            If mScript.funcAndPropHash.TryGetValue(mScript.mainClass(classId).Names(0) + "_" + arrayWords(qbClosePos + 3).Word, fp) Then
                                                'В хэше funcAndPropHash найдена наша функция / свойство. Устанавливаем что это и ставим в правильный регистр
                                                arrayWords(qbClosePos + 3).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                                                arrayWords(qbClosePos + 3).Word = fp.elementName
                                            Else
                                                'В хэше funcAndPropHash наша функция не найдена - возможно, это функция или свойство, добавленные динамически (AddProperty / AddFunction)
                                                arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_PROPERTY 'на самом деле это может быть и функция, 
                                                'и свойство - не найдя в структуре проверить точно невозможно
                                            End If
                                            Continue For
                                        Else
                                            mScript.LAST_ERROR = "После " + arrayWords(i).Word + "[...]. стоит не название функции или свойства."
                                            Return -1
                                        End If
                                    End If
                                Else
                                    If curLine <> nextLinePos2 Then Continue For
                                    mScript.LAST_ERROR = "После " + arrayWords(i).Word + "[...]. не стоит название функции или свойства."
                                    Return -1
                                End If
                            Else
                                'имя класса НЕ найдено в хэше mainClassHash (возможно, создан динамически с New Class)
                                Dim lc As Integer = nextLinePos, wc As Integer = nextWordPos
                                Dim wrdNext2 As EditWordType
                                Dim nextLinePos2 As Integer, nextWordPos2 As Integer
                                If nextLinePos = curLine AndAlso qbClosePos + 1 < arrayWords.GetUpperBound(0) Then
                                    If qbClosePos + 2 <= arrayWords.Count - 1 Then
                                        wrdNext2 = arrayWords(qbClosePos + 2)
                                    Else
                                        wrdNext2 = GetNextWord(lc, wc, True)
                                    End If

                                Else
                                    wrdNext2 = GetNextWord(lc, wc, True) 'qbClosePos + 2
                                    nextLinePos2 = lc
                                    nextWordPos2 = wc
                                End If


                                arrayWords(i).wordType = EditWordTypeEnum.W_CLASS
                                arrayWords(i).classId = -2 'неизвестный класс (возможно, создан динамически с New Class)
                                arrayWords(qbClosePos + 1).classId = -2  'ставим к какому классу относится точка 
                                If curLine = nextLinePos2 AndAlso IsNothing(wrdNext2.Word) = False Then
                                    'If qbClosePos + 1 < arrayWords.GetUpperBound(0) Then
                                    If arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_NOTHING Then
                                        'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                        arrayWords(qbClosePos + 2).classId = -2 'неизвестный класс
                                        If arrayWords.GetUpperBound(0) > qbClosePos + 2 AndAlso (arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                            arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_PROPERTY 'это свойство
                                        Else
                                            arrayWords(qbClosePos + 2).wordType = EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                            'и свойство - не зная класса проверить точно невозможно
                                        End If
                                        Continue For
                                    Else
                                        If i + 2 < arrayWords.GetUpperBound(0) AndAlso (arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER Or arrayWords(i + 2).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) AndAlso arrayWords(i + 3).wordType = EditWordTypeEnum.W_NOTHING Then
                                            'заполняем данные о функции или свойстве, идущими за этим классом через точку
                                            arrayWords(qbClosePos + 3).classId = -2 'неизвестный класс
                                            If arrayWords.GetUpperBound(0) > qbClosePos + 3 AndAlso (arrayWords(qbClosePos + 4).wordType = EditWordTypeEnum.W_OPERATOR_EQUAL OrElse arrayWords(qbClosePos + 4).wordType = EditWordTypeEnum.W_OPERATOR_COMPARE) Then
                                                arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_PROPERTY 'это свойство
                                            Else
                                                arrayWords(qbClosePos + 3).wordType = EditWordTypeEnum.W_FUNCTION  'на самом деле это может быть и функция, 
                                                'и свойство - не зная класса проверить точно невозможно
                                            End If
                                            Continue For
                                        Else
                                            mScript.LAST_ERROR = "После " + arrayWords(i).Word + ". стоит не название функции или свойства."
                                            Return -1
                                        End If
                                    End If
                                Else
                                    If curLine <> nextLinePos2 Then Continue For
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
                        arrayWords(i).wordType = IIf(fp.elementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY)
                        arrayWords(i).Word = fp.elementName
                        Exit For
                    End If
                Next

                If arrayWords(i).wordType = EditWordTypeEnum.W_NOTHING Then
                    'это переменная
                    arrayWords(i).wordType = EditWordTypeEnum.W_VARIABLE
                    arrayWords(i).classId = -1
                End If
            Next

            'Расставляем правильно пробелы и проверяем на ошибки синтаксиса
            Dim isFirstWord As Boolean, isLastWord As Boolean 'текущее слово первое или последнее? (напр., после Then слово считается первым)
            'перебираем в цикле все слова (кроме последнего)
            For i = 0 To arrayWords.GetUpperBound(0) - 1
                If i = 0 OrElse arrayWords(i - 1).wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION OrElse arrayWords(i - 1).wordType = EditWordTypeEnum.W_HTML_DATA OrElse arrayWords(i - 1).Word = "Then " Then
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
                If arrayWords(i + 1).wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION OrElse arrayWords(i + 1).Word = "Then" OrElse arrayWords(i + 1).wordType = EditWordTypeEnum.W_HTML_DATA Then
                    'если после текущего слова стоит Then или ; - слово последнее
                    isLastWord = True
                Else
                    isLastWord = False
                End If
                'пербираем слова и, в зависимости от их значения и значения слова, следующего за ним, ставим где надо пробел (и проверяем на ошибки синтаксиса)
                Select Case arrayWords(i).wordType 'простое число, строка или True/False
                    Case EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_SIMPLE_STRING
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается с простого значения."
                            Return -1
                        End If
                        Select Case arrayWords(i + 1).wordType
                            Case EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                                EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_COMMA, _
                                EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_HTML_DATA
                                '= <> And + & ) ] , ;
                                'ничего
                            Case EditWordTypeEnum.W_BLOCK_IF 'Then
                                If arrayWords(i + 1).Word.Trim = "Then" Then
                                    If arrayWords(i + 1).Word.First <> " "c Then arrayWords(i).Word &= " "
                                Else
                                    mScript.LAST_ERROR = "Неверная запись блока If ... Then."
                                    Return -1
                                End If
                            Case EditWordTypeEnum.W_BLOCK_FOR  'To or Step
                                If arrayWords(i + 1).Word.Trim = "To" OrElse arrayWords(i + 1).Word.Trim = "Step" Then
                                    If arrayWords(i + 1).Word.First <> " "c Then arrayWords(i).Word &= " "
                                Else
                                    mScript.LAST_ERROR = "Неверная запись блока For ... Next."
                                    Return -1
                                End If
                            Case EditWordTypeEnum.W_STRINGS_DISSOCIATION
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case EditWordTypeEnum.W_VARIABLE 'переменная
                        Select Case arrayWords(i + 1).wordType
                            Case EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                                EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_COMMA, _
                                EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_QUAD_BRACKET_OPEN, _
                                EditWordTypeEnum.W_HTML_DATA
                                '= <> And + & ) ] , ; [
                                'ничего
                            Case EditWordTypeEnum.W_BLOCK_IF 'Then
                                If arrayWords(i + 1).Word = "Then" Then
                                    arrayWords(i).Word &= " "
                                Else
                                    mScript.LAST_ERROR = "Неверная запись блока If ... Then."
                                    Return -1
                                End If
                            Case EditWordTypeEnum.W_BLOCK_FOR  'To
                                If arrayWords(i + 1).Word = "To" Then
                                    arrayWords(i).Word &= " "
                                Else
                                    mScript.LAST_ERROR = "Неверная запись блока For ... Next."
                                    Return -1
                                End If
                            Case EditWordTypeEnum.W_STRINGS_DISSOCIATION
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case EditWordTypeEnum.W_OPERATOR_MATH, EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, _
                        EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL
                        '+ & <> =
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается с оператора."
                            Return -1
                        End If
                        Select Case arrayWords(i + 1).wordType
                            Case EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                                EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE, _
                                EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_PROPERTY, _
                                EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PARAM, _
                                EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                                EditWordTypeEnum.W_STRINGS_DISSOCIATION, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                                EditWordTypeEnum.W_CONVERT_TO_STRING
                                'True 5 'xx' myVar Code Prop Func Param ParamCount ( _ $ #
                                'все ок
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case EditWordTypeEnum.W_OPERATOR_LOGIC
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается с оператора."
                            Return -1
                        End If
                        Select Case arrayWords(i + 1).wordType
                            Case EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                                EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE, _
                                EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_PROPERTY, _
                                EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PARAM, _
                                EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                                EditWordTypeEnum.W_STRINGS_DISSOCIATION, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                                EditWordTypeEnum.W_CONVERT_TO_STRING
                                'True 5 'xx' myVar Code Prop Func Param ParamCount ( _ $ #
                                'arrayWords(i).Word = " " & arrayWords(i).Word & " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается со скобки."
                            Return -1
                        End If
                        ')
                        Select Case arrayWords(i + 1).wordType
                            Case EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_OPERATOR_MATH, _
                                EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, _
                                EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                                EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_COMMA
                                'ok
                            Case Else
                                If arrayWords(i + 1).Word.First <> " " Then arrayWords(i).Word &= " "
                        End Select
                    Case EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_QUAD_BRACKET_OPEN, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        '( [ ]
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается со скобки."
                            Return -1
                        End If
                    Case EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_FUNCTION
                        'Func
                        Select Case arrayWords(i + 1).wordType
                            Case EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                                EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                                EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_COMMA
                                'ok
                            Case EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                                EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                                EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                                EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PARAM_COUNT, _
                                EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_BOOL, _
                                EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_SIMPLE_STRING, _
                                EditWordTypeEnum.W_STRINGS_DISSOCIATION, EditWordTypeEnum.W_VARIABLE
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case EditWordTypeEnum.W_PARAM
                        'Param[x]
                        Select Case arrayWords(i + 1).wordType
                            Case EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                                EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                                EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_QUAD_BRACKET_OPEN, _
                                EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_COMMA
                                'ok
                            Case EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                                EditWordTypeEnum.W_STRINGS_DISSOCIATION
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case EditWordTypeEnum.W_PARAM_COUNT
                        'ParamCount
                        If isFirstWord Then
                            mScript.LAST_ERROR = "Неверная строка кода. Строка начинается с ParamCount."
                            Return -1
                        End If
                        Select Case arrayWords(i + 1).wordType
                            Case EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                                EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                                EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                                EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                                EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_QUAD_BRACKET_OPEN, _
                                EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_COMMA
                                'ok
                            Case EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                                EditWordTypeEnum.W_STRINGS_DISSOCIATION
                                arrayWords(i).Word &= " "
                            Case Else
                                mScript.LAST_ERROR = "Неправильный синтаксис."
                                Return -1
                        End Select
                    Case EditWordTypeEnum.W_BREAK, EditWordTypeEnum.W_CONTINUE, EditWordTypeEnum.W_EXIT, _
                        EditWordTypeEnum.W_JUMP, EditWordTypeEnum.W_MARK, _
                        EditWordTypeEnum.W_RETURN, EditWordTypeEnum.W_BLOCK_EVENT, _
                        EditWordTypeEnum.W_GLOBAL
                        If isFirstWord Then
                            If isLastWord = False Then
                                If arrayWords(i + 1).Word.First <> " " AndAlso arrayWords(i).Word <> "?" Then arrayWords(i).Word &= " "
                            End If
                        Else
                            mScript.LAST_ERROR = "Неправильный синтаксис."
                            Return -1
                        End If
                    Case EditWordTypeEnum.W_SWITCH, EditWordTypeEnum.W_WRAP, EditWordTypeEnum.W_HTML, _
                        EditWordTypeEnum.W_CYCLE_END, EditWordTypeEnum.W_BLOCK_DOWHILE, _
                        EditWordTypeEnum.W_BLOCK_FUNCTION, EditWordTypeEnum.W_BLOCK_NEWCLASS, _
                        EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, EditWordTypeEnum.W_REM_CLASS
                        'Select Case Switch Wrap HTML Append End ...
                        If arrayWords(i + 1).Word.First <> " " Then arrayWords(i).Word &= " "
                        'Case EditWordTypeEnum.W_COMMA
                End Select
            Next

            'проверка последнего символа
            If arrayWords(arrayWords.GetUpperBound(0)).wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then 'если последний символ _ , то слово на самом деле не последнее
                Select Case arrayWords(arrayWords.GetUpperBound(0)).wordType
                    Case EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_CONVERT_TO_NUMBER, EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_JUMP, EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_POINT, EditWordTypeEnum.W_QUAD_BRACKET_OPEN, EditWordTypeEnum.W_REM_CLASS
                        mScript.LAST_ERROR = "Неправильный синтаксис."
                        Return -1
                End Select
            End If
            If arrayWords.GetUpperBound(0) = 0 And isStringDissociation = False Then
                'когда в строке только одно слово
                Select Case arrayWords(0).wordType
                    Case EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_SIMPLE_STRING
                        mScript.LAST_ERROR = "Неправильный синтаксис."
                        Return -1
                End Select
            End If
            Return 1
        End Function

        Public Sub PrepareHTMLforErrorInfo(ByRef wbRtbHelp As WebBrowser)
            'получаем документ html
            If IsNothing(wbRtbHelp.Url) OrElse wbRtbHelp.Url.ToString <> HelpFile Then
                wbRtbHelp.Navigate(HelpFile)
                Do While IsNothing(wbRtbHelp.Url) OrElse wbRtbHelp.Url.ToString <> HelpFile
                    System.Windows.Forms.Application.DoEvents()
                Loop
            End If
            Dim hDoc As HtmlDocument = wbRtbHelp.Document

            hDoc.GetElementById("funcInfo").Style = "display:none"
            Dim hEl As HtmlElement = hDoc.GetElementById("errors")
            hEl.InnerHtml = ""
            hEl.Style = "display:block"
        End Sub

        ''' <summary>
        ''' функция определяет, находится ли текущая строка внутри блока HTML ... End HTML или Wrap ... End Wrap
        ''' </summary>
        ''' <param name="rtb">ссылка на RichTextBox</param>
        ''' <param name="curLineIndex">индекс линии для проверки</param>
        ''' <param name="isInExec">находится ли текущая строка в блоке <exec>...</exec></param>
        Public Function IsInTextBlock(ByRef rtb As RichTextBox, ByVal curLineIndex As Integer, ByRef isInExec As ExecBlockEnum) As TextBlockEnum
            If curLineIndex = 0 Then Return IIf(IsTextBlockByDefault, TextBlockEnum.TEXT_HTML, TextBlockEnum.NO_TEXT_BLOCK) 'текущая строка - первая - нет, не находимся
            isInExec = ExecBlockEnum.NO_EXEC
            If IsNothing(CodeData) Then Return TextBlockEnum.NO_TEXT_BLOCK
            'Перебираем код в структуре CodeData от текущей - 1 до первой

            If IsTextBlockByDefault Then
                'Описание локации. По умолчанию мы в блоке HTML, внутренниеHTML/Wrap  запрещены
                For i As Integer = curLineIndex - 1 To 0 Step -1
                    If i < CodeData.Count - 1 AndAlso IsNothing(CodeData(i).Code) = False Then
                        If i > CodeData.Count - 1 Then Exit For
                        Dim execOpenPos As Integer = CodeData(i).Code(CodeData(i).Code.GetUpperBound(0)).Word.LastIndexOf("<exec>", StringComparison.CurrentCultureIgnoreCase)
                        Dim execClosePos As Integer = CodeData(i).Code(CodeData(i).Code.GetUpperBound(0)).Word.LastIndexOf("</exec>", StringComparison.CurrentCultureIgnoreCase)
                        If execOpenPos > execClosePos Then 'последний открывающий - мы в теге
                            isInExec = ExecBlockEnum.EXEC_MULTILINE
                            Return TextBlockEnum.TEXT_HTML
                        ElseIf execClosePos > -1 Then
                            isInExec = ExecBlockEnum.NO_EXEC
                            Return TextBlockEnum.TEXT_HTML
                        End If
                    End If
                Next
            Else
                'Обычный код, не описание локации
                For i As Integer = curLineIndex - 1 To 0 Step -1
                    If IsNothing(CodeData(i).Code) = False Then
                        If CodeData(i).Code(0).wordType = EditWordTypeEnum.W_CYCLE_END Then
                            If CodeData(i).Code.Length > 1 AndAlso (CodeData(i).Code(1).wordType = EditWordTypeEnum.W_HTML OrElse CodeData(i).Code(1).wordType = EditWordTypeEnum.W_WRAP) Then
                                'End Wrap / End HTML - мы за пределами блока
                                If IsTextBlockByDefault Then Return TextBlockEnum.TEXT_HTML
                                Return TextBlockEnum.NO_TEXT_BLOCK
                            End If
                        ElseIf (CodeData(i).Code(0).wordType = EditWordTypeEnum.W_HTML OrElse CodeData(i).Code(0).wordType = EditWordTypeEnum.W_WRAP) Then
                            'HTML или Wrap - мы в блоке
                            Dim execOpenPos As Integer, execClosePos As Integer
                            'определяем находимся ли мы в <exec>
                            For j = curLineIndex - 1 To i + 1 Step -1 'ищем первую непустую строку, начиная с перед текущей
                                If IsNothing(CodeData(j).Code) Then Continue For
                                If CodeData(j).Code.GetUpperBound(0) > 0 AndAlso CodeData(j).Code(CodeData(j).Code.GetUpperBound(0)).wordType <> EditWordTypeEnum.W_HTML_DATA Then
                                    isInExec = ExecBlockEnum.EXEC_MULTILINE
                                    Exit For
                                End If
                                'ищем последний открывающий и закрывающий тег exec
                                execOpenPos = CodeData(j).Code(CodeData(j).Code.GetUpperBound(0)).Word.LastIndexOf("<exec>", StringComparison.CurrentCultureIgnoreCase)
                                execClosePos = CodeData(j).Code(CodeData(j).Code.GetUpperBound(0)).Word.LastIndexOf("</exec>", StringComparison.CurrentCultureIgnoreCase)
                                If execOpenPos = execClosePos Then 'нет ни одного (оба равны -1)
                                    isInExec = ExecBlockEnum.NO_EXEC
                                    Exit For
                                ElseIf execOpenPos < execClosePos Then 'последним стоит закрывающий - значит мы не в теге
                                    isInExec = ExecBlockEnum.NO_EXEC
                                    Exit For
                                Else 'последний открывающий - мы в теге
                                    isInExec = ExecBlockEnum.EXEC_MULTILINE
                                    Exit For
                                End If
                            Next
                            If CodeData(i).Code(0).wordType = EditWordTypeEnum.W_HTML Then
                                'HTML - мы в блоке
                                Return TextBlockEnum.TEXT_HTML
                            Else
                                'Wrap - мы в блоке
                                Return TextBlockEnum.TEXT_WRAP
                            End If
                        End If
                    End If
                Next
            End If

            'не нашли ни один блок - не находимся
            If IsTextBlockByDefault Then Return TextBlockEnum.TEXT_HTML
            Return TextBlockEnum.NO_TEXT_BLOCK
        End Function

        Private Sub ShowTextError(ByRef rtb As RichTextBox, ByVal strIndex As Integer, ByVal nonCompleted As Boolean, Optional ByVal setStringAsWrong As Boolean = True)
            Dim selStart As Integer = rtb.SelectionStart
            Dim selLength As Integer = rtb.SelectionLength
            rtb.Select(rtb.GetFirstCharIndexFromLine(strIndex), rtb.Lines(strIndex).Length)
            If nonCompleted Then
                rtb.SelectionColor = styleHash(EditWordTypeEnum.W_NOTHING).style_Color
                rtb.SelectionFont = New Font(rtb.Font.FontFamily, Convert.ToSingle(rtb.Font.Size), styleHash(EditWordTypeEnum.W_NOTHING).font_style)
                rtb.SelectionStart = selStart
                rtb.SelectionLength = selLength
            Else
                rtb.SelectionColor = styleHash(EditWordTypeEnum.W_ERROR).style_Color
                rtb.SelectionFont = New Font(rtb.Font.FontFamily, Convert.ToSingle(rtb.Font.Size), styleHash(EditWordTypeEnum.W_ERROR).font_style)
            End If
            If setStringAsWrong Then
                ReDim CodeData(strIndex).Code(0)
                CodeData(strIndex).Code(0).Word = rtb.Lines(strIndex)
                CodeData(strIndex).Code(0).wordType = EditWordTypeEnum.W_ERROR
            End If

            If nonCompleted = False Then
                Dim hDoc As HtmlDocument = wbRtbHelp.Document
                If IsNothing(hDoc) Then Return
                Dim hContainer As HtmlElement = hDoc.GetElementById("errors")
                If IsNothing(hContainer) Then Return
                Dim hErr As HtmlElement
                hErr = hDoc.CreateElement("DIV")
                hErr.InnerHtml = "Строка № " + (strIndex + 1).ToString + ". " + mScript.LAST_ERROR
                hErr.SetAttribute("Line", strIndex.ToString)
                hContainer.AppendChild(hErr)
                'MsgBox(mScript.LAST_ERROR)
            End If
        End Sub

        Private Function StripComments(ByRef curString As String, ByRef strComment As String, ByRef strStartingSpaces As String, ByVal nonCompleted As Boolean) As Integer
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
                            If nonCompleted Then
                                Return 1 'если строка в процессе напичания - ошибку не считаем
                            Else
                                mScript.LAST_ERROR = "Не найдена закрывающая кавычка"
                                Return -1
                            End If
                        End If
                        'Обрабатываем экранированную кавычку /'
                        If curString.Chars(j - 1) = "/"c Then GoTo qbQuoteSearching
                    Else
                        'комментарий раньше. Сохраняем его и выходим из цикла
                        If nonCompleted Then
                            If j = 0 Then
                                strComment = curString.Substring(j)
                            Else
                                Dim strTemp As String = curString.Substring(0, j) 'подсчет пробелов между последним оператором и началом комментария
                                strComment = Space(strTemp.Length - strTemp.TrimEnd.Length) + curString.Substring(j)
                            End If
                        Else
                            strComment = " " + curString.Substring(j)
                        End If
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

#End Region

#Region "Загрузка данных"
        Public Sub FillOperators()
            arrOperators.Clear()

            arrOperators.Add("True", EditWordTypeEnum.W_SIMPLE_BOOL)
            arrOperators.Add("False", EditWordTypeEnum.W_SIMPLE_BOOL)
            arrOperators.Add("And", EditWordTypeEnum.W_OPERATOR_LOGIC)
            arrOperators.Add("Or", EditWordTypeEnum.W_OPERATOR_LOGIC)
            arrOperators.Add("Xor", EditWordTypeEnum.W_OPERATOR_LOGIC)
            arrOperators.Add("ParamCount", EditWordTypeEnum.W_PARAM_COUNT)
            arrOperators.Add("Param", EditWordTypeEnum.W_PARAM)
            arrOperators.Add("Exit", EditWordTypeEnum.W_EXIT)
            arrOperators.Add("Select", EditWordTypeEnum.W_SWITCH)
            arrOperators.Add("Switch", EditWordTypeEnum.W_SWITCH)
            arrOperators.Add("Case", EditWordTypeEnum.W_SWITCH)
            arrOperators.Add("End", EditWordTypeEnum.W_CYCLE_END)
            arrOperators.Add("Wrap", EditWordTypeEnum.W_WRAP)
            arrOperators.Add("HTML", EditWordTypeEnum.W_HTML)
            arrOperators.Add("Append", EditWordTypeEnum.W_HTML)
            arrOperators.Add("Do", EditWordTypeEnum.W_BLOCK_DOWHILE)
            arrOperators.Add("While", EditWordTypeEnum.W_BLOCK_DOWHILE)
            arrOperators.Add("Loop", EditWordTypeEnum.W_BLOCK_DOWHILE)
            arrOperators.Add("Function", EditWordTypeEnum.W_BLOCK_FUNCTION)
            arrOperators.Add("New", EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Class", EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Name", EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Prop", EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Func", EditWordTypeEnum.W_BLOCK_NEWCLASS)
            arrOperators.Add("Rem", EditWordTypeEnum.W_REM_CLASS)
            arrOperators.Add("For", EditWordTypeEnum.W_BLOCK_FOR)
            arrOperators.Add("To", EditWordTypeEnum.W_BLOCK_FOR)
            arrOperators.Add("Step", EditWordTypeEnum.W_BLOCK_FOR)
            arrOperators.Add("Next", EditWordTypeEnum.W_BLOCK_FOR)
            arrOperators.Add("If", EditWordTypeEnum.W_BLOCK_IF)
            arrOperators.Add("Then", EditWordTypeEnum.W_BLOCK_IF)
            arrOperators.Add("Else", EditWordTypeEnum.W_BLOCK_IF)
            arrOperators.Add("ElseIf", EditWordTypeEnum.W_BLOCK_IF)
            arrOperators.Add("Jump", EditWordTypeEnum.W_JUMP)
            arrOperators.Add("Global", EditWordTypeEnum.W_GLOBAL)
            arrOperators.Add("Return", EditWordTypeEnum.W_RETURN)
            arrOperators.Add("Break", EditWordTypeEnum.W_BREAK)
            arrOperators.Add("Continue", EditWordTypeEnum.W_CONTINUE)
            arrOperators.Add("Event", EditWordTypeEnum.W_BLOCK_EVENT)
            arrOperators.Add("?", EditWordTypeEnum.W_RETURN)

            For Each strKey In arrOperators.Keys
                arrOperatorsKeys.Add(strKey, strKey)
            Next
        End Sub

        Public Sub FillStyle()
            styleHash.Clear()

            styleHash.Add(EditWordTypeEnum.W_NOTHING, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.Black})

            Dim st As New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.Navy}
            styleHash.Add(EditWordTypeEnum.W_BLOCK_DOWHILE, st)
            styleHash.Add(EditWordTypeEnum.W_BLOCK_FOR, st)
            styleHash.Add(EditWordTypeEnum.W_BLOCK_FUNCTION, st)
            styleHash.Add(EditWordTypeEnum.W_BLOCK_IF, st)
            styleHash.Add(EditWordTypeEnum.W_SWITCH, st)
            styleHash.Add(EditWordTypeEnum.W_BLOCK_EVENT, st)
            styleHash.Add(EditWordTypeEnum.W_BREAK, st)
            styleHash.Add(EditWordTypeEnum.W_CONTINUE, st)

            styleHash.Add(EditWordTypeEnum.W_BLOCK_NEWCLASS, New StylePresetType With {.font_style = FontStyle.Bold, .style_Color = Color.Brown})
            styleHash.Add(EditWordTypeEnum.W_REM_CLASS, New StylePresetType With {.font_style = FontStyle.Bold, .style_Color = Color.Brown})

            styleHash.Add(EditWordTypeEnum.W_JUMP, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.Maroon})

            styleHash.Add(EditWordTypeEnum.W_GLOBAL, New StylePresetType With {.font_style = FontStyle.Bold, .style_Color = Color.Brown})

            styleHash.Add(EditWordTypeEnum.W_EXIT, New StylePresetType With {.font_style = FontStyle.Bold, .style_Color = Color.OrangeRed})
            styleHash.Add(EditWordTypeEnum.W_RETURN, New StylePresetType With {.font_style = FontStyle.Bold, .style_Color = Color.OrangeRed})

            styleHash.Add(EditWordTypeEnum.W_FUNCTION, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.DarkBlue})
            styleHash.Add(EditWordTypeEnum.W_PROPERTY, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.DarkBlue})
            styleHash.Add(EditWordTypeEnum.W_CLASS, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.DarkBlue})

            styleHash.Add(EditWordTypeEnum.W_HTML, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.Brown})
            styleHash.Add(EditWordTypeEnum.W_WRAP, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.Brown})

            styleHash.Add(EditWordTypeEnum.W_MARK, New StylePresetType With {.font_style = FontStyle.Underline, .style_Color = Color.Maroon})

            styleHash.Add(EditWordTypeEnum.W_SIMPLE_BOOL, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.IndianRed})
            styleHash.Add(EditWordTypeEnum.W_SIMPLE_NUMBER, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.IndianRed})

            styleHash.Add(EditWordTypeEnum.W_SIMPLE_STRING, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.IndianRed})

            styleHash.Add(EditWordTypeEnum.W_VARIABLE, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.DarkRed})
            styleHash.Add(EditWordTypeEnum.W_PARAM, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.DarkRed})
            styleHash.Add(EditWordTypeEnum.W_PARAM_COUNT, New StylePresetType With {.font_style = FontStyle.Regular, .style_Color = Color.DarkRed})

            styleHash.Add(EditWordTypeEnum.W_COMMENTS, New StylePresetType With {.style_Color = Color.DarkGreen, .font_style = FontStyle.Italic})

            styleHash.Add(EditWordTypeEnum.W_HTML_TAG, New StylePresetType With {.style_Color = Color.Blue, .font_style = FontStyle.Regular})

            styleHash.Add(EditWordTypeEnum.W_ERROR, New StylePresetType With {.style_Color = Color.Red, .font_style = FontStyle.Regular})
        End Sub

        Public Sub FillCheckCodeStructure()
            'arrAllowedTypes - типы слов за данным, которые разрешены
            'arrBannedWords - слова (из разрешенных типов), которые запрещены
            Dim arrAllowedTypes() As EditWordTypeEnum
            Dim arrBannedWords() As String

            'Do
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_DOWHILE}
            arrBannedWords = {"Do", "Loop"}
            checkCodeWords.Add("Do", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                                    .canBeLast = True, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'While
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PARAM_COUNT, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_BOOL, _
                               EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_OVAL_BRACKET_OPEN}
            checkCodeWords.Add("While", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                                    .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "Do", .canBeFinalBlockWord = False})
            'Loop
            arrAllowedTypes = {EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_STRINGS_CONSOLIDATION}
            checkCodeWords.Add("Loop", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = True, .shouldBeFirst = True, .shouldBeLast = True, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Event ... End Event
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_BLOCK_EVENT, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = True})
            'For
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PARAM_COUNT, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_VARIABLE}
            checkCodeWords.Add("For", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'To
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_VARIABLE}
            checkCodeWords.Add("To", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                                    .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Step
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_VARIABLE}
            checkCodeWords.Add("Step", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                                    .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Next
            arrAllowedTypes = {EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_STRINGS_CONSOLIDATION, _
                           EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_CLASS, _
                           EditWordTypeEnum.W_CONVERT_TO_NUMBER, EditWordTypeEnum.W_OVAL_BRACKET_OPEN}
            checkCodeWords.Add("Next", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = True, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Function ... End Function
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_BLOCK_FUNCTION, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = True})
            'If
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_SIMPLE_BOOL, _
                               EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_VARIABLE}
            checkCodeWords.Add("If", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = True})
            'Then
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_IF, EditWordTypeEnum.W_BREAK, _
                               EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONTINUE, _
                               EditWordTypeEnum.W_CONVERT_TO_NUMBER, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_EXIT, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_JUMP, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_REM_CLASS, EditWordTypeEnum.W_RETURN, _
                               EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_VARIABLE}
            arrBannedWords = {"Then", "Else", "ElseIf"}
            checkCodeWords.Add("Then", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = False, _
                                    .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Else
            arrAllowedTypes = {EditWordTypeEnum.W_STRINGS_CONSOLIDATION}
            checkCodeWords.Add("Else", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = True, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'ElseIf
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeWords.Add("ElseIf", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'New
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_NEWCLASS}
            arrBannedWords = {"New", "Name", "Prop", "Func"}
            checkCodeWords.Add("New", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Class ... End Class
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_OVAL_BRACKET_OPEN}
            checkCodeWords.Add("Class", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                                    .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = True})
            'Prop
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_OVAL_BRACKET_OPEN}
            checkCodeWords.Add("Prop", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Name
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_OVAL_BRACKET_OPEN}
            checkCodeWords.Add("Name", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Func
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_OVAL_BRACKET_OPEN}
            checkCodeWords.Add("Func", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_CLASS
            arrAllowedTypes = {EditWordTypeEnum.W_POINT, EditWordTypeEnum.W_QUAD_BRACKET_OPEN}
            checkCodeTypes.Add(EditWordTypeEnum.W_CLASS, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Break
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_HTML_DATA, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_STRINGS_CONSOLIDATION, _
                               EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_BREAK, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = True, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'comma
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_COMMA}
            checkCodeTypes.Add(EditWordTypeEnum.W_COMMA, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                                    .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Continue
            arrAllowedTypes = {EditWordTypeEnum.W_STRINGS_CONSOLIDATION}
            checkCodeTypes.Add(EditWordTypeEnum.W_CONTINUE, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = True, .shouldBeFirst = True, .shouldBeLast = True, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_CONVERT_TO_NUMBER
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PARAM_COUNT, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_BOOL, _
                               EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_CONVERT_TO_NUMBER, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_CONVERT_TO_STRING
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PARAM_COUNT, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_BOOL, _
                               EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_CONVERT_TO_STRING, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'End
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_EVENT, EditWordTypeEnum.W_BLOCK_FUNCTION, _
                               EditWordTypeEnum.W_BLOCK_IF, EditWordTypeEnum.W_BLOCK_NEWCLASS, _
                               EditWordTypeEnum.W_HTML, EditWordTypeEnum.W_WRAP}
            arrBannedWords = {"Then", "Else", "ElseIf", "New", "Append"}
            checkCodeTypes.Add(EditWordTypeEnum.W_CYCLE_END, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Exit
            arrAllowedTypes = {EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_STRINGS_CONSOLIDATION}
            checkCodeTypes.Add(EditWordTypeEnum.W_EXIT, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = True, .shouldBeFirst = True, .shouldBeLast = True, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Function
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_FUNCTION, _
                                EditWordTypeEnum.W_BLOCK_IF, EditWordTypeEnum.W_CLASS, _
                               EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_OPERATOR_COMPARE, _
                               EditWordTypeEnum.W_OPERATOR_EQUAL, EditWordTypeEnum.W_OPERATOR_LOGIC, _
                               EditWordTypeEnum.W_OPERATOR_MATH, EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, _
                               EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PARAM_COUNT, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_STRINGS_CONSOLIDATION, _
                               EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_CONVERT_TO_NUMBER, EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_COMMA}
            arrBannedWords = {"If", "Else", "ElseIf", "For", "Next"}
            checkCodeTypes.Add(EditWordTypeEnum.W_FUNCTION, New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                                    .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'HTML
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_HTML, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_VARIABLE}
            arrBannedWords = {"HTML"}
            checkCodeWords.Add("HTML", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                                    .canBeLast = True, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = True})
            'Append
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeWords.Add("Append", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                                    .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "HTML", .canBeFinalBlockWord = False})
            'HTML Data
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_DOWHILE, EditWordTypeEnum.W_BLOCK_EVENT, _
                               EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_FUNCTION, _
                               EditWordTypeEnum.W_BLOCK_IF, EditWordTypeEnum.W_BLOCK_NEWCLASS, _
                               EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_JUMP, _
                               EditWordTypeEnum.W_MARK, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_REM_CLASS, _
                               EditWordTypeEnum.W_RETURN, EditWordTypeEnum.W_SWITCH, _
                               EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_GLOBAL}
            arrBannedWords = {"While", "Loop", "To", "Step", "Next", "Then", "Else", "ElseIf", "Class", "Name", "Func", "Prop", "Case"}
            checkCodeTypes.Add(EditWordTypeEnum.W_HTML_DATA, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Jump
            arrAllowedTypes = {EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_SIMPLE_NUMBER}
            checkCodeTypes.Add(EditWordTypeEnum.W_JUMP, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                            .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Global
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_GLOBAL, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                            .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Mark
            arrAllowedTypes = {EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_STRINGS_CONSOLIDATION}
            checkCodeTypes.Add(EditWordTypeEnum.W_MARK, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                            .canBeLast = True, .shouldBeFirst = True, .shouldBeLast = True, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_OPERATOR_COMPARE
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_OPERATOR_COMPARE, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                            .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_OPERATOR_EQUAL
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_OPERATOR_EQUAL, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                            .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_OPERATOR_LOGIC
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_OPERATOR_LOGIC, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                            .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_OPERATOR_MATH
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_OPERATOR_MATH, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                            .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_OPERATOR_STRINGS_MERGER
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                            .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_OVAL_BRACKET_CLOSE
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                               EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_HTML_DATA, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                               EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_STRINGS_CONSOLIDATION, _
                               EditWordTypeEnum.W_SWITCH}
            arrBannedWords = {"For", "Next", "If", "Else", "ElseIf", "Select", "Switch", "Case"}
            checkCodeTypes.Add(EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = False, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_OVAL_BRACKET_OPEN
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_OVAL_BRACKET_OPEN, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                            .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_PARAM
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_QUAD_BRACKET_OPEN, EditWordTypeEnum.W_STRINGS_CONSOLIDATION}
            arrBannedWords = {"For", "Next", "If", "Else", "ElseIf"}
            checkCodeTypes.Add(EditWordTypeEnum.W_PARAM, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_PARAM_COUNT
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                               EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_HTML_DATA, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                               EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_STRINGS_CONSOLIDATION, _
                               EditWordTypeEnum.W_SWITCH}
            arrBannedWords = {"For", "Next", "If", "Else", "ElseIf", "Select", "Case", "Switch"}
            checkCodeTypes.Add(EditWordTypeEnum.W_PARAM_COUNT, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = False, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_POINT
            arrAllowedTypes = {EditWordTypeEnum.W_CONVERT_TO_NUMBER, EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_POINT, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                            .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_PROPERTY
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                               EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_HTML_DATA, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                               EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_QUAD_BRACKET_OPEN, _
                               EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_SWITCH, EditWordTypeEnum.W_OVAL_BRACKET_OPEN}
            arrBannedWords = {"For", "Next", "If", "Else", "ElseIf", "Select", "Case", "Switch"}
            checkCodeTypes.Add(EditWordTypeEnum.W_PROPERTY, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_QUAD_BRACKET_CLOSE
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                               EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_HTML_DATA, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                               EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_POINT, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_SWITCH}
            arrBannedWords = {"For", "Next", "If", "Else", "ElseIf", "Select", "Case", "Switch"}
            checkCodeTypes.Add(EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = False, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_QUAD_BRACKET_OPEN
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_QUAD_BRACKET_OPEN, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = False, _
                            .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_REM_CLASS
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_NEWCLASS}
            arrBannedWords = {"New", "Name", "Prop", "Func"}
            checkCodeTypes.Add(EditWordTypeEnum.W_REM_CLASS, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                            .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Return
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_HTML_DATA, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PARAM_COUNT, _
                               EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_SIMPLE_BOOL, _
                               EditWordTypeEnum.W_SIMPLE_NUMBER, EditWordTypeEnum.W_SIMPLE_STRING, _
                               EditWordTypeEnum.W_STRINGS_CONSOLIDATION, EditWordTypeEnum.W_VARIABLE}
            checkCodeTypes.Add(EditWordTypeEnum.W_RETURN, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                            .canBeLast = True, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_SIMPLE_BOOL
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                               EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_HTML_DATA, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                               EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_STRINGS_CONSOLIDATION, _
                               EditWordTypeEnum.W_SWITCH}
            arrBannedWords = {"For", "Next", "If", "Else", "ElseIf", "Select", "Case", "Switch"}
            checkCodeTypes.Add(EditWordTypeEnum.W_SIMPLE_BOOL, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = False, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_SIMPLE_NUMBER
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                               EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_HTML_DATA, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                               EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_STRINGS_CONSOLIDATION, _
                               EditWordTypeEnum.W_SWITCH}
            arrBannedWords = {"For", "Next", "If", "Else", "ElseIf", "Select", "Case", "Switch"}
            checkCodeTypes.Add(EditWordTypeEnum.W_SIMPLE_NUMBER, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = False, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_SIMPLE_STRING
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                               EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_HTML_DATA, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                               EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_STRINGS_CONSOLIDATION, _
                               EditWordTypeEnum.W_SWITCH}
            arrBannedWords = {"For", "Next", "If", "Else", "ElseIf", "Select", "Case", "Switch"}
            checkCodeTypes.Add(EditWordTypeEnum.W_SIMPLE_STRING, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = False, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Select
            arrAllowedTypes = {EditWordTypeEnum.W_SWITCH}
            arrBannedWords = {"Select", "Switch"}
            checkCodeWords.Add("Select", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = True})
            'Switch
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE, _
                               EditWordTypeEnum.W_BLOCK_IF}
            arrBannedWords = {"Select", "Switch"}
            checkCodeWords.Add("Switch", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = True})
            'Case
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_NUMBER, _
                               EditWordTypeEnum.W_CONVERT_TO_STRING, EditWordTypeEnum.W_FUNCTION, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_PARAM, _
                               EditWordTypeEnum.W_PARAM_COUNT, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_BOOL, EditWordTypeEnum.W_SIMPLE_NUMBER, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE, _
                               EditWordTypeEnum.W_BLOCK_IF}
            arrBannedWords = {"If", "ElseIf", "Then"}
            checkCodeWords.Add("Case", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'W_VARIABLE
            arrAllowedTypes = {EditWordTypeEnum.W_BLOCK_FOR, EditWordTypeEnum.W_BLOCK_IF, _
                               EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_HTML_DATA, _
                               EditWordTypeEnum.W_OPERATOR_COMPARE, EditWordTypeEnum.W_OPERATOR_EQUAL, _
                               EditWordTypeEnum.W_OPERATOR_LOGIC, EditWordTypeEnum.W_OPERATOR_MATH, _
                               EditWordTypeEnum.W_OPERATOR_STRINGS_MERGER, EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, _
                               EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, EditWordTypeEnum.W_QUAD_BRACKET_OPEN, _
                               EditWordTypeEnum.W_STRINGS_CONSOLIDATION}
            arrBannedWords = {"For", "Next", "If", "Else", "ElseIf"}
            checkCodeTypes.Add(EditWordTypeEnum.W_VARIABLE, New CheckCodeStructure _
                       With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = arrBannedWords, .canBeFirst = True, _
                            .canBeLast = True, .shouldBeFirst = False, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = False})
            'Wrap
            arrAllowedTypes = {EditWordTypeEnum.W_CLASS, EditWordTypeEnum.W_CONVERT_TO_STRING, _
                               EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                               EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PROPERTY, _
                               EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_VARIABLE}
            checkCodeWords.Add("Wrap", New CheckCodeStructure _
                               With {.AllowedNextWordTypes = arrAllowedTypes, .bannedNextWords = Nothing, .canBeFirst = True, _
                                    .canBeLast = False, .shouldBeFirst = True, .shouldBeLast = False, .previousWordShouldBe = "", .canBeFinalBlockWord = True})
        End Sub

#End Region

#Region "CodeData Management"
        ''' <summary>
        ''' Сериализует объект CodeData в xml и возвращает результат в виде xml-текста
        ''' </summary>
        Public Function SerializeCodeData(Optional ByRef cd() As CodeDataType = Nothing) As String
            If IsNothing(cd) Then cd = CodeData
            If IsNothing(cd) OrElse cd.Count = 0 Then Return ""
            If cd.Count = 1 Then
                If IsNothing(cd(0).Code) OrElse cd(0).Code.Length = 0 Then
                    If IsNothing(cd(0).Comments) OrElse cd(0).Comments.Length = 0 Then
                        If IsNothing(cd(0).StartingSpaces) OrElse cd(0).StartingSpaces.Length = 0 Then
                            Return ""
                        End If
                    End If
                End If
            End If
            Dim sb As New System.Text.StringBuilder

            Using xmlWriter As System.Xml.XmlWriter = System.Xml.XmlWriter.Create(sb, New System.Xml.XmlWriterSettings With {.OmitXmlDeclaration = True})
                Dim x As New System.Xml.Serialization.XmlSerializer(cd.GetType)
                'CheckSymb01() - после изменений в PrepareText функция вроде как не нужна
                x.Serialize(xmlWriter, cd)
            End Using
            Return sb.ToString
        End Function

        ''' <summary>
        ''' Десериализует xml-строку и преобразует ее в массив кода в формате CodeDataType()
        ''' </summary>
        ''' <param name="strXML">xml-строка</param>
        Public Function DeserializeCodeData(ByRef strXML As String) As CodeDataType()
            If strXML.Length = 0 Then
                Dim dt() As CodeDataType = Nothing
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
                        Dim x As New System.Xml.Serialization.XmlSerializer(CodeData.GetType)
                        Return x.Deserialize(xmlMemorySteam2)
                    End Using
                End Using

            End Using
        End Function

        ''' <summary>
        ''' Загружает CodeData из сериализованного свойства
        ''' </summary>
        ''' <param name="strXml"></param>
        ''' <remarks></remarks>
        Public Sub LoadCodeFromProperty(ByVal strXml As String)
            SendMessage(Handle, WM_SetRedraw, 0, 0)
            rtbCanRaiseEvents = False
            Clear()
            'десериализуем из xml-строки код в CodeData
            Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(strXml)
            If cRes <> MatewScript.ContainsCodeEnum.CODE AndAlso cRes <> MatewScript.ContainsCodeEnum.LONG_TEXT Then
                'вероятно, содержимое свойства - код в виде обычного текста (или исполняемая строка)
                Text = strXml
                rtbCanRaiseEvents = True
                PrepareText(Me)
                SendMessage(Handle, WM_SetRedraw, 1, 0)
                csUndo.BeginNew()
                Return
            End If
            CodeData = DeserializeCodeData(strXml)
            If CodeData.Length = 1 AndAlso IsNothing(CodeData(0).Code) AndAlso IsNothing(CodeData(0).Comments) AndAlso IsNothing(CodeData(0).StartingSpaces) Then
                Text = strXml
                rtbCanRaiseEvents = True
                PrepareText(Me)
                SendMessage(Handle, WM_SetRedraw, 1, 0)
                csUndo.BeginNew()
                Return
            End If
            Dim sb As New System.Text.StringBuilder
            For i As Integer = 0 To CodeData.Count - 1
                If IsNothing(CodeData(i).StartingSpaces) = False AndAlso CodeData(i).StartingSpaces.Length > 0 Then sb.Append(CodeData(i).StartingSpaces)
                If IsNothing(CodeData(i).Code) = False AndAlso CodeData(i).Code.Length > 0 Then
                    For j As Integer = 0 To CodeData(i).Code.Count - 1
                        sb.Append(CodeData(i).Code(j).Word)
                    Next
                End If
                If IsNothing(CodeData(i).Comments) = False Then
                    If IsNothing(CodeData(i).Code) = False AndAlso CodeData(i).Code.Length > 0 Then
                        sb.Append(CodeData(i).Comments)
                    Else
                        sb.Append(CodeData(i).Comments.TrimStart)
                    End If
                End If
                If i <> CodeData.Count - 1 Then sb.AppendLine()
            Next
            Text = sb.ToString
            SetCorretSpaces(Me, False)
            'после SetCorretSpaces опять ставим rtbCanRaiseEvents и WM_SetRedraw на ускорение
            rtbCanRaiseEvents = False
            SendMessage(Handle, WM_SetRedraw, 0, 0)

            For i As Integer = 0 To CodeData.Count - 1
                Dim stSpacesLength As Integer = 0
                Dim isComm As Boolean = False
                If IsNothing(CodeData(i).Comments) = False AndAlso CodeData(i).Comments.Length > 0 Then isComm = True
                If IsNothing(CodeData(i).StartingSpaces) = False Then stSpacesLength = CodeData(i).StartingSpaces.Length
                DrawWords(Me, CodeData(i).Code, i, stSpacesLength, True)

                If IsNothing(CodeData(i).StartingSpaces) = False Then sb.Append(CodeData(i).StartingSpaces)
                If IsNothing(CodeData(i).Code) = False AndAlso CodeData(i).Code.Length > 0 Then
                    For j As Integer = 0 To CodeData(i).Code.Count - 1
                        sb.Append(CodeData(i).Code(j).Word)
                    Next
                End If
                If IsNothing(CodeData(i).Comments) = False Then sb.Append(CodeData(i).Comments)
                If i <> CodeData.Count - 1 Then sb.AppendLine()
            Next

            rtbCanRaiseEvents = True
            SendMessage(Handle, WM_SetRedraw, 1, 0)
            csUndo.BeginNew()
        End Sub

        ''' <summary>
        ''' Загружаем код из другого объекта CodeData
        ''' </summary>
        ''' <param name="cd">объект CodeData</param>
        ''' <remarks></remarks>
        Public Sub LoadCodeFromCodeData(ByVal cd() As CodeDataType)
            rtbCanRaiseEvents = False
            Visible = False
            System.Windows.Forms.Application.DoEvents()
            Clear()
            If IsNothing(cd) Then
                CodeData = cd
            Else
                CodeData = CopyCodeDataArray(cd)
            End If
            If IsNothing(cd) Then ReDim CodeData(0)
            If CodeData.Length = 1 AndAlso IsNothing(CodeData(0).Code) AndAlso IsNothing(CodeData(0).Comments) AndAlso IsNothing(CodeData(0).StartingSpaces) Then
                rtbCanRaiseEvents = True
                Visible = True
                csUndo.BeginNew()
                Return
            End If
            Dim sb As New System.Text.StringBuilder
            For i As Integer = 0 To CodeData.Count - 1
                If IsNothing(CodeData(i).StartingSpaces) = False AndAlso CodeData(i).StartingSpaces.Length > 0 Then sb.Append(CodeData(i).StartingSpaces)
                If IsNothing(CodeData(i).Code) = False AndAlso CodeData(i).Code.Length > 0 Then
                    For j As Integer = 0 To CodeData(i).Code.Count - 1
                        sb.Append(CodeData(i).Code(j).Word)
                    Next
                End If
                If IsNothing(CodeData(i).Comments) = False Then
                    If IsNothing(CodeData(i).Code) = False AndAlso CodeData(i).Code.Length > 0 Then
                        sb.Append(CodeData(i).Comments)
                    Else
                        sb.Append(CodeData(i).Comments.TrimStart)
                    End If
                End If
                If i <> CodeData.Count - 1 Then sb.AppendLine()
            Next
            SendMessage(Handle, WM_SetRedraw, 0, 0)
            SuspendLayout()
            Text = sb.ToString
            SetCorretSpaces(Me, False)
            'Visible = False
            'Left = -10000
            SendMessage(Handle, WM_SetRedraw, 0, 0)
            rtbCanRaiseEvents = False

            sb.Clear()
            For i As Integer = 0 To CodeData.Count - 1
                Dim stSpacesLength As Integer = 0
                Dim isComm As Boolean = False
                If IsNothing(CodeData(i).Comments) = False AndAlso CodeData(i).Comments.Length > 0 Then isComm = True
                If IsNothing(CodeData(i).StartingSpaces) = False Then stSpacesLength = CodeData(i).StartingSpaces.Length
                DrawWords(Me, CodeData(i).Code, i, stSpacesLength, True)

                If IsNothing(CodeData(i).StartingSpaces) = False Then sb.Append(CodeData(i).StartingSpaces)
                If IsNothing(CodeData(i).Code) = False AndAlso CodeData(i).Code.Length > 0 Then
                    For j As Integer = 0 To CodeData(i).Code.Count - 1
                        sb.Append(CodeData(i).Code(j).Word)
                    Next
                End If
                If IsNothing(CodeData(i).Comments) = False Then sb.Append(CodeData(i).Comments)
                If i <> CodeData.Count - 1 Then sb.AppendLine()
            Next


            ResumeLayout()
            Left = 0
            rtbCanRaiseEvents = True
            Show()
            SendMessage(Handle, WM_SetRedraw, 1, 0)
            csUndo.BeginNew()
        End Sub

#End Region

#Region "Global syntax checking"
        Private Function CheckStringForSyntaxErrors(ByRef rtb As RichTextBox, ByVal curLine As Integer) As Boolean
            'Функция проверяет весь текст на наличие ошибок синтаксиса и возвращает False, если найдена хоть одна ошибка
            'Основан на данных из массивов checkCodeWords и checkCodeTypes

            If IsNothing(CodeData) Then Return False 'пустой код
            If IsNothing(CodeData(curLine).Code) Then Return False 'пустая строка
            'запрет событий в rtb

            Dim curWord, nextWord, prevWord As EditWordType 'текущее, предыдущее и следующее слово, с которым работаем
            Dim isFirstWord, isLastWord As Boolean 'является ли текущеее слово первым или последним в выражении
            Dim checkData As CheckCodeStructure = Nothing 'данные по текущему слову из массива checkCodeWords или checkCodeTypes
            Dim j As Integer = 0 'счетчик

            'перебираем в цикле все слова в текущей строке
checkNext:
            For i As Integer = 0 To CodeData(curLine).Code.GetUpperBound(0)
                curWord = CodeData(curLine).Code(i) 'полуаем текущее слово
                If curWord.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                    'текущее слово "_" (разделение строки). Если оно не последнее в строке - это ошибка. Если последнее - переходим к следующей строке
                    If i < CodeData(curLine).Code.GetUpperBound(0) Then
                        mScript.LAST_ERROR = "Ошибка синтаксиса. разделение строки в неположенном месте."
                        ShowTextError(rtb, curLine, True)
                    End If
                    curLine += 1
                    GoTo checkNext
                End If
                'Текущее слово ";" (объединение строк) - переходим к следующему слову
                If curWord.wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION Then Continue For

                'Получаем слово за текущим
                If i < CodeData(curLine).Code.GetUpperBound(0) Then
                    nextWord = CodeData(curLine).Code(i + 1) 'получаем следующее слово
                    If nextWord.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                        'это слово - "_". Получаем слово, следующее дальше (на следующей строке)
                        j = curLine + 1 'номер следующей строки
nextLine:
                        If j > CodeData.GetUpperBound(0) Then
                            mScript.LAST_ERROR = "Ошибка синтаксиса. Нет кода за разделителем строки."
                            Return False
                        End If
                        If IsNothing(CodeData(j).Code) Then
                            mScript.LAST_ERROR = "Ошибка синтаксиса. Нет кода за разделителем строки."
                            Return False
                        End If
                        nextWord = CodeData(j).Code(0) 'получаем первое слово на следующей строке
                        If nextWord.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                            'это слово - опять "_". Переходим к следующей строке и повторяем действия
                            j += 1
                            GoTo nextLine
                        End If
                    End If
                Else
                    'текущее слово - последнее в выражении
                    nextWord = Nothing
                End If
                'Получаем слово перед текущим
                If i > 0 Then
                    prevWord = CodeData(curLine).Code(i - 1) 'получаем предыдущее слово
                    If prevWord.wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION OrElse prevWord.wordType = EditWordTypeEnum.W_HTML_DATA _
                        OrElse prevWord.Word.Trim = "Then" Then
                        'предыдущее слово ; или Then или тект html - значит текущее слово - первое в выражении (предыдущее слово относится к другому выражению)
                        prevWord = Nothing
                    End If
                Else
                    prevWord = Nothing 'текущее слово - первое в строке
                    If curLine > 0 Then
                        'проверяем, не заканчивается ли предыдущая строка на _
                        j = curLine - 1 'номер предыдущей строки
prevLine:
                        If j < 0 Then
                            mScript.LAST_ERROR = "Ошибка синтаксиса. Нет кода перед разделителем строки."
                            Return False
                        End If
                        If IsNothing(CodeData(j).Code) = False Then
                            prevWord = CodeData(j).Code(CodeData(j).Code.GetUpperBound(0)) 'получаем последнее слово в предыдущей строке
                            If prevWord.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                                'и это слово _
                                If CodeData(j).Code.GetUpperBound(0) = 0 Then
                                    '_  - единственный символ кода в строке. Переходим к предыдущей строке и повторяем действия
                                    j -= 1
                                    prevWord = Nothing
                                    GoTo prevLine
                                End If
                                'получаем слово перед _ - это и есть слово перед текущим
                                prevWord = CodeData(j).Code(CodeData(j).Code.GetUpperBound(0) - 1)
                            Else
                                'последнее слово в предыдущей  строке не _. Тогда предыдущего слова в выражении нет (текущее - первое)
                                prevWord = Nothing
                            End If
                        End If
                    End If
                End If

                'Является ли слово первым
                If IsNothing(prevWord) OrElse prevWord.wordType = EditWordTypeEnum.W_NOTHING OrElse prevWord.wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION _
                    OrElse prevWord.wordType = EditWordTypeEnum.W_HTML_DATA OrElse prevWord.Word.Trim = "Then" Then
                    isFirstWord = True 'слово первое
                Else
                    isFirstWord = False 'слово не первое
                End If
                'Является ли слово последним
                If IsNothing(nextWord) OrElse nextWord.wordType = EditWordTypeEnum.W_NOTHING OrElse nextWord.wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION _
                    OrElse nextWord.wordType = EditWordTypeEnum.W_HTML_DATA OrElse nextWord.Word.Trim = "Then" Then
                    isLastWord = True 'слово последнее
                Else
                    isLastWord = False 'слово не последнее
                End If

                'Получаем стр-ру checkData по текущему слову из checkCodeWords или checkCodeTypes
                If checkCodeWords.TryGetValue(curWord.Word.Trim, checkData) = False Then
                    If checkCodeTypes.TryGetValue(curWord.wordType, checkData) = False Then
                        'слово не найдено - ошибка
                        mScript.LAST_ERROR = "Ошибка синтаксиса."
                        Return False
                    End If
                End If

                If IsNothing(nextWord) = False AndAlso nextWord.wordType <> EditWordTypeEnum.W_NOTHING Then
                    'проверяем, является ли следующее слово разрешенным для текущего. Если нет - ошибка
                    If Array.IndexOf(checkData.AllowedNextWordTypes, nextWord.wordType) = -1 Then
                        mScript.LAST_ERROR = "Ошибка синтаксиса. " + nextWord.Word.Trim + " не может следовать за " + curWord.Word.Trim + "."
                        Return False
                    End If
                    If IsNothing(checkData.bannedNextWords) = False Then
                        If Array.IndexOf(checkData.bannedNextWords, nextWord.Word.Trim) > -1 Then
                            mScript.LAST_ERROR = "Ошибка синтаксиса. " + nextWord.Word.Trim + " не может следовать за " + curWord.Word.Trim + "."
                            Return False
                        End If
                    End If
                End If

                If checkData.canBeFirst = False And isFirstWord Then
                    'текущее слово - первое, но быть им не может - ошибка
                    mScript.LAST_ERROR = "Ошибка синтаксиса." + curWord.Word.Trim + " не может быть первым."
                    Return False
                End If
                If checkData.canBeLast = False And isLastWord Then
                    'текущее слово - последнее, но быть им не может - ошибка
                    If (checkData.canBeFinalBlockWord AndAlso IsNothing(prevWord) = False AndAlso prevWord.wordType <> EditWordTypeEnum.W_NOTHING AndAlso prevWord.wordType = EditWordTypeEnum.W_CYCLE_END) = False Then
                        mScript.LAST_ERROR = "Ошибка синтаксиса." + curWord.Word.Trim + " не может быть последним."
                        Return False
                    End If
                End If
                If checkData.shouldBeFirst AndAlso isFirstWord = False Then
                    'текущее слово может быть только первым, но им не является - ошибка
                    If (checkData.canBeFinalBlockWord AndAlso IsNothing(prevWord) = False AndAlso prevWord.wordType <> EditWordTypeEnum.W_NOTHING AndAlso prevWord.wordType = EditWordTypeEnum.W_CYCLE_END) = False Then
                        mScript.LAST_ERROR = "Ошибка синтаксиса." + curWord.Word.Trim + " может быть только первым."
                        Return False
                    End If
                End If
                If checkData.shouldBeLast AndAlso isLastWord = False Then
                    'текущее слово может быть только последним, но им не является - ошибка
                    mScript.LAST_ERROR = "Ошибка синтаксиса." + curWord.Word.Trim + " может быть только последним."
                    Return False
                End If
                If checkData.previousWordShouldBe.Length > 0 Then
                    'предыдущее слово не является разрешенным - ошибка
                    If IsNothing(prevWord) OrElse prevWord.wordType = EditWordTypeEnum.W_NOTHING OrElse prevWord.Word.Trim <> checkData.previousWordShouldBe Then
                        mScript.LAST_ERROR = "Ошибка синтаксиса." + prevWord.Word.Trim + " не может стоять перед " + curWord.Word.Trim + "."
                        Return False
                    End If
                End If
            Next

            Return True
        End Function

        Public Function CheckTextForSyntaxErrors(ByRef rtb As RichTextBox) As Boolean
            'Функция проверяет весь текст на наличие ошибок синтаксиса и возвращает False, если найдена хоть одна ошибка
            'Основан на данных из массивов checkCodeWords и checkCodeTypes

            If IsNothing(CodeData) Then Return True 'пустой код
            'запрет событий в rtb
            Dim prevCanRaiseEvents As Boolean = rtbCanRaiseEvents
            rtbCanRaiseEvents = False

            Dim curWord, nextWord, prevWord As EditWordType 'текущее, предыдущее и следующее слово, с которым работаем
            Dim isFirstWord, isLastWord As Boolean 'является ли текущеее слово первым или последним в выражении
            Dim curLine As Integer = -1 'текущая линия
            Dim checkData As CheckCodeStructure = Nothing 'данные по текущему слову из массива checkCodeWords или checkCodeTypes
            Dim j As Integer = 0 'счетчик
            Dim noErrors As Boolean = True

            'перебираем в цикле все строки в CodeData
            Do While curLine < CodeData.GetUpperBound(0)
                curLine += 1
                If IsNothing(CodeData(curLine).Code) Then Continue Do
                'перебираем в цикле все слова в текущей строке
                For i As Integer = 0 To CodeData(curLine).Code.GetUpperBound(0)
                    curWord = CodeData(curLine).Code(i) 'получаем текущее слово
                    If curWord.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                        'текущее слово "_" (разделение строки). Если оно не последнее в строке - это ошибка. Если последнее - переходим к следующей строке
                        If i < CodeData(curLine).Code.GetUpperBound(0) Then
                            mScript.LAST_ERROR = "Ошибка синтаксиса. разделение строки в неположенном месте."
                            ShowTextError(rtb, curLine, False)
                            noErrors = False
                        End If
                        Continue Do
                    End If
                    'Текущее слово ";" (объединение строк) - переходим к следующему слову
                    If curWord.wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION Then Continue For

                    'Получаем слово за текущим
                    If i < CodeData(curLine).Code.GetUpperBound(0) Then
                        nextWord = CodeData(curLine).Code(i + 1) 'получаем следующее слово
                        If nextWord.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                            'это слово - "_". Получаем слово, следующее дальше (на следующей строке)
                            j = curLine + 1 'номер следующей строки
nextLine:
                            If j > CodeData.GetUpperBound(0) Then
                                mScript.LAST_ERROR = "Ошибка синтаксиса. Нет кода за разделителем строки."
                                noErrors = False
                                ShowTextError(rtb, curLine, False)
                                Continue Do
                            End If
                            If IsNothing(CodeData(j).Code) Then
                                mScript.LAST_ERROR = "Ошибка синтаксиса. Нет кода за разделителем строки."
                                ShowTextError(rtb, curLine, False)
                                noErrors = False
                                Continue Do
                            End If
                            nextWord = CodeData(j).Code(0) 'получаем первое слово на следующей строке
                            If nextWord.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                                'это слово - опять "_". Переходим к следующей строке и повторяем действия
                                j += 1
                                GoTo nextLine
                            End If
                        End If
                    Else
                        'текущее слово - последнее в выражении
                        nextWord = Nothing
                    End If
                    'Получаем слово перед текущим
                    If i > 0 Then
                        prevWord = CodeData(curLine).Code(i - 1) 'получаем предыдущее слово
                        If prevWord.wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION OrElse prevWord.wordType = EditWordTypeEnum.W_HTML_DATA _
                            OrElse prevWord.Word.Trim = "Then" Then
                            'предыдущее слово ; или Then или тект html - значит текущее слово - первое в выражении (предыдущее слово относится к другому выражению)
                            prevWord = Nothing
                        End If
                    Else
                        prevWord = Nothing 'текущее слово - первое в строке
                        If curLine > 0 Then
                            'проверяем, не заканчивается ли предыдущая строка на _
                            j = curLine - 1 'номер предыдущей строки
prevLine:
                            If j < 0 Then
                                mScript.LAST_ERROR = "Ошибка синтаксиса. Нет кода перед разделителем строки."
                                ShowTextError(rtb, curLine, False)
                                noErrors = False
                                Continue Do
                            End If
                            If IsNothing(CodeData(j).Code) = False Then
                                prevWord = CodeData(j).Code(CodeData(j).Code.GetUpperBound(0)) 'получаем последнее слово в предыдущей строке
                                If prevWord.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                                    'и это слово _
                                    If CodeData(j).Code.GetUpperBound(0) = 0 Then
                                        '_  - единственный символ кода в строке. Переходим к предыдущей строке и повторяем действия
                                        j -= 1
                                        prevWord = Nothing
                                        GoTo prevLine
                                    End If
                                    'получаем слово перед _ - это и есть слово перед текущим
                                    prevWord = CodeData(j).Code(CodeData(j).Code.GetUpperBound(0) - 1)
                                Else
                                    'последнее слово в предыдущей  строке не _. Тогда предыдущего слова в выражении нет (текущее - первое)
                                    prevWord = Nothing
                                End If
                            End If
                        End If
                    End If

                    'Является ли слово первым
                    If IsNothing(prevWord) OrElse prevWord.wordType = EditWordTypeEnum.W_NOTHING OrElse prevWord.wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION _
                        OrElse prevWord.wordType = EditWordTypeEnum.W_HTML_DATA OrElse prevWord.Word.Trim = "Then" Then
                        isFirstWord = True 'слово первое
                    Else
                        isFirstWord = False 'слово не первое
                    End If
                    'Является ли слово последним
                    If IsNothing(nextWord) OrElse nextWord.wordType = EditWordTypeEnum.W_NOTHING OrElse nextWord.wordType = EditWordTypeEnum.W_STRINGS_CONSOLIDATION _
                        OrElse nextWord.wordType = EditWordTypeEnum.W_HTML_DATA OrElse nextWord.Word.Trim = "Then" Then
                        isLastWord = True 'слово последнее
                    Else
                        isLastWord = False 'слово не последнее
                    End If

                    'Получаем стр-ру checkData по текущему слову из checkCodeWords или checkCodeTypes
                    If curWord.wordType = EditWordTypeEnum.W_FUNCTION OrElse curWord.wordType = EditWordTypeEnum.W_PROPERTY Then
                        If checkCodeTypes.TryGetValue(curWord.wordType, checkData) = False Then
                            'слово не найдено - ошибка
                            mScript.LAST_ERROR = "Ошибка синтаксиса."
                            ShowTextError(rtb, curLine, False)
                            noErrors = False
                            Continue Do
                        End If
                    Else
                        If checkCodeWords.TryGetValue(curWord.Word.Trim, checkData) = False Then
                            If checkCodeTypes.TryGetValue(curWord.wordType, checkData) = False Then
                                'слово не найдено - ошибка
                                mScript.LAST_ERROR = "Ошибка синтаксиса."
                                ShowTextError(rtb, curLine, False)
                                noErrors = False
                                Continue Do
                            End If
                        End If
                    End If

                    If IsNothing(nextWord) = False AndAlso nextWord.wordType <> EditWordTypeEnum.W_NOTHING Then
                        'проверяем, является ли следующее слово разрешенным для текущего. Если нет - ошибка
                        If Array.IndexOf(checkData.AllowedNextWordTypes, nextWord.wordType) = -1 Then
                            If curWord.wordType <> EditWordTypeEnum.W_CYCLE_END Then
                                mScript.LAST_ERROR = "Ошибка синтаксиса. " + nextWord.Word.Trim + " не может следовать за " + curWord.Word.Trim + "."
                                ShowTextError(rtb, curLine, False)
                                noErrors = False
                                Continue Do
                            End If
                        End If
                        If IsNothing(checkData.bannedNextWords) = False Then
                            If Array.IndexOf(checkData.bannedNextWords, nextWord.Word.Trim) > -1 Then
                                mScript.LAST_ERROR = "Ошибка синтаксиса. " + nextWord.Word.Trim + " не может следовать за " + curWord.Word.Trim + "."
                                ShowTextError(rtb, curLine, False)
                                noErrors = False
                                Continue Do
                            End If
                        End If
                    End If

                    If checkData.canBeFirst = False And isFirstWord Then
                        'текущее слово - первое, но быть им не может - ошибка
                        mScript.LAST_ERROR = "Ошибка синтаксиса." + curWord.Word.Trim + " не может быть первым."
                        ShowTextError(rtb, curLine, False)
                        noErrors = False
                        Continue Do
                    End If
                    If checkData.canBeLast = False And isLastWord Then
                        'текущее слово - последнее, но быть им не может - ошибка
                        If (checkData.canBeFinalBlockWord AndAlso IsNothing(prevWord) = False AndAlso prevWord.wordType <> EditWordTypeEnum.W_NOTHING AndAlso prevWord.wordType = EditWordTypeEnum.W_CYCLE_END) = False Then
                            mScript.LAST_ERROR = "Ошибка синтаксиса." + curWord.Word.Trim + " не может быть последним."
                            ShowTextError(rtb, curLine, False)
                            noErrors = False
                            Continue Do
                        End If
                    End If
                    If checkData.shouldBeFirst AndAlso isFirstWord = False Then
                        'текущее слово может быть только первым, но им не является - ошибка
                        If (checkData.canBeFinalBlockWord AndAlso IsNothing(prevWord) = False AndAlso prevWord.wordType <> EditWordTypeEnum.W_NOTHING AndAlso prevWord.wordType = EditWordTypeEnum.W_CYCLE_END) = False Then
                            mScript.LAST_ERROR = "Ошибка синтаксиса." + curWord.Word.Trim + " может быть только первым."
                            ShowTextError(rtb, curLine, False)
                            noErrors = False
                            Continue Do
                        End If
                    End If
                    If checkData.shouldBeLast AndAlso isLastWord = False Then
                        'текущее слово может быть только последним, но им не является - ошибка
                        mScript.LAST_ERROR = "Ошибка синтаксиса." + curWord.Word.Trim + " может быть только последним."
                        ShowTextError(rtb, curLine, False)
                        noErrors = False
                        Continue Do
                    End If
                    If checkData.previousWordShouldBe.Length > 0 Then
                        'предыдущее слово не является разрешенным - ошибка
                        If IsNothing(prevWord) OrElse prevWord.wordType = EditWordTypeEnum.W_NOTHING OrElse prevWord.Word.Trim <> checkData.previousWordShouldBe Then
                            mScript.LAST_ERROR = "Ошибка синтаксиса." + prevWord.Word.Trim + " не может стоять перед " + curWord.Word.Trim + "."
                            ShowTextError(rtb, curLine, False)
                            noErrors = False
                            Continue Do
                        End If
                    End If
                Next
            Loop

            rtbCanRaiseEvents = prevCanRaiseEvents 'разрешаем события rtb
            Return noErrors
        End Function

        Public Function CheckCycles(ByRef rtb As RichTextBox) As Boolean
            'Процедура проверяет код на парность открытия / закрытия всех блоков и возвращает False, если есть незакрытые блоки
            If IsNothing(CodeData) Then Return True

            Dim firstWord As EditWordType 'для получения первого слова каждой строки
            Dim strWord As String 'для слова из firstWord
            Dim textBlock As TextBlockEnum = TextBlockEnum.NO_TEXT_BLOCK 'для информации не в текстовом ли блоке мы сейчас
            Dim textBlockStartLine As Integer 'начало текстового блока

            'сохраняем текущие верхнюю видимую строку, выделение и линию
            Dim blockBalance() As Integer, blockId As Integer 'баланс открытия / закрытия блоков
            Dim blockStartLine() As Integer 'положения начала блоков
            ReDim blockBalance(AddSpacesArray.GetUpperBound(0)) 'расширяем массив по кол-ву видов блоков
            ReDim blockStartLine(AddSpacesArray.GetUpperBound(0)) 'расширяем массив по кол-ву видов блоков

            Dim i As Integer = 0
checkAgain:
            While i <= CodeData.GetUpperBound(0) 'перебираем все строки
                If IsNothing(CodeData(i).Code) OrElse CodeData(i).Code.GetUpperBound(0) = -1 Then
                    'в этой строке нет кода
                    i += 1
                    Continue While
                End If

                firstWord = CodeData(i).Code(0) 'получаем первое слово текущей строки
                strWord = firstWord.Word.Trim
                If textBlock <> TextBlockEnum.NO_TEXT_BLOCK Then
                    'мы в текстовом блоке
                    If textBlock = TextBlockEnum.TEXT_HTML Then
                        'точнее, в блоке HTML
                        If firstWord.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso CodeData(i).Code.GetUpperBound(0) > 0 _
                            AndAlso CodeData(i).Code(1).wordType = EditWordTypeEnum.W_HTML Then
                            textBlock = TextBlockEnum.NO_TEXT_BLOCK 'End HTML - конец блока
                        Else
                            i += 1
                            Continue While 'текстовый блок продолжается
                        End If
                    ElseIf textBlock = TextBlockEnum.TEXT_WRAP Then
                        'точнее, в блоке Wrap
                        If firstWord.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso CodeData(i).Code.GetUpperBound(0) > 0 _
                            AndAlso CodeData(i).Code(1).wordType = EditWordTypeEnum.W_WRAP Then
                            textBlock = TextBlockEnum.NO_TEXT_BLOCK 'End Wrap - конец блока
                        Else
                            i += 1
                            Continue While 'текстовый блок продолжается
                        End If
                    End If
                End If

                If firstWord.wordType = EditWordTypeEnum.W_CYCLE_END OrElse strWord = "Loop" OrElse strWord = "Next" Then 'OrElse
                    'strWord = "Case" OrElse strWord = "Else" OrElse strWord = "ElseIf" Then
                    'первое слово End (или Loop и др., тоже конец блока)
                    If firstWord.wordType = EditWordTypeEnum.W_CYCLE_END Then
                        blockId = Array.IndexOf(AddSpacesArray, CodeData(i).Code(1).wordType)
                        If blockId > -1 Then
                            blockBalance(blockId) -= 1
                        End If
                    Else
                        blockBalance(Array.IndexOf(AddSpacesArray, firstWord.wordType)) -= 1
                    End If
                ElseIf (firstWord.wordType = EditWordTypeEnum.W_BLOCK_NEWCLASS AndAlso strWord <> "New") = False AndAlso _
                    (firstWord.wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso CodeData(i).Code(CodeData(i).Code.GetUpperBound(0)).Word.Trim <> "Then") = False Then
                    'это может быть блок
                    blockId = Array.IndexOf(AddSpacesArray, firstWord.wordType)
                    If blockId > -1 Then
                        'это точно блок 
                        If strWord <> "ElseIf" AndAlso strWord <> "Else" AndAlso strWord <> "Case" Then
                            blockBalance(blockId) += 1
                            blockStartLine(blockId) = i
                        End If
                    End If
                End If

                If firstWord.wordType = EditWordTypeEnum.W_HTML Then
                    textBlock = TextBlockEnum.TEXT_HTML 'мы вошли в блок HTML
                    textBlockStartLine = i
                ElseIf firstWord.wordType = EditWordTypeEnum.W_WRAP Then
                    textBlock = TextBlockEnum.TEXT_WRAP 'мы вошли в блок HTML
                    textBlockStartLine = i
                End If

                i += 1
            End While

            'запрет событий в rtb
            Dim prevCanRaiseEvents As Boolean = rtbCanRaiseEvents
            rtbCanRaiseEvents = False
            If textBlock = TextBlockEnum.TEXT_HTML Then
                mScript.LAST_ERROR = "Незакрытый блок HTML."
                ShowTextError(rtb, textBlockStartLine, False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            ElseIf textBlock = TextBlockEnum.TEXT_WRAP Then
                mScript.LAST_ERROR = "Незакрытый блок Wrap."
                ShowTextError(rtb, textBlockStartLine, False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            End If

            blockId = Array.IndexOf(AddSpacesArray, EditWordTypeEnum.W_BLOCK_DOWHILE)
            If blockBalance(blockId) > 0 Then
                mScript.LAST_ERROR = "Незакрытый блок Do While."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            ElseIf blockBalance(blockId) < 0 Then
                mScript.LAST_ERROR = "Лишний(е) оператор(ы) Loop блока Do While."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            End If

            blockId = Array.IndexOf(AddSpacesArray, EditWordTypeEnum.W_BLOCK_EVENT)
            If blockBalance(blockId) > 0 Then
                mScript.LAST_ERROR = "Незакрытый блок Event."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            ElseIf blockBalance(blockId) < 0 Then
                mScript.LAST_ERROR = "Лишний(е) оператор(ы) End Event блока Event."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            End If

            blockId = Array.IndexOf(AddSpacesArray, EditWordTypeEnum.W_BLOCK_FOR)
            If blockBalance(blockId) > 0 Then
                mScript.LAST_ERROR = "Незакрытый блок For ... Next."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            ElseIf blockBalance(blockId) < 0 Then
                mScript.LAST_ERROR = "Лишний(е) оператор(ы) Next блока For."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            End If

            blockId = Array.IndexOf(AddSpacesArray, EditWordTypeEnum.W_BLOCK_FUNCTION)
            If blockBalance(blockId) > 0 Then
                mScript.LAST_ERROR = "Незакрытый блок Function."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            ElseIf blockBalance(blockId) < 0 Then
                mScript.LAST_ERROR = "Лишний(е) оператор(ы) End Function блока Function."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            End If

            blockId = Array.IndexOf(AddSpacesArray, EditWordTypeEnum.W_BLOCK_IF)
            If blockBalance(blockId) > 0 Then
                mScript.LAST_ERROR = "Незакрытый блок If ... Then."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            ElseIf blockBalance(blockId) < 0 Then
                mScript.LAST_ERROR = "Лишний(е) оператор(ы) End If блока If ... Then."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            End If

            blockId = Array.IndexOf(AddSpacesArray, EditWordTypeEnum.W_BLOCK_NEWCLASS)
            If blockBalance(blockId) > 0 Then
                mScript.LAST_ERROR = "Незакрытый блок New Class."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            ElseIf blockBalance(blockId) < 0 Then
                mScript.LAST_ERROR = "Лишний(е) оператор(ы) End Class блока New Class."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            End If

            blockId = Array.IndexOf(AddSpacesArray, EditWordTypeEnum.W_SWITCH)
            If blockBalance(blockId) > 0 Then
                mScript.LAST_ERROR = "Незакрытый блок Select Case (Switch)."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            ElseIf blockBalance(blockId) < 0 Then
                mScript.LAST_ERROR = "Лишний(е) оператор(ы) End Select (Switch) блока Select Case (Switch)."
                ShowTextError(rtb, blockStartLine(blockId), False, False)
                rtbCanRaiseEvents = prevCanRaiseEvents
                Return False
            End If

            rtbCanRaiseEvents = True
            Return True
        End Function

#End Region

#Region "String Seeking"
        ''' <summary>
        ''' Очищает background контрола от предыдущих выделений
        ''' </summary>
        Public Sub ClearPreviousSelection()
            CanRaiseCodeEvents = False 'запрет TextChanged
            'получаем верхнюю видимую линию
            Dim upperLine As Integer = GetCharIndexFromPosition(New Point(5, 5))
            upperLine = GetLineFromCharIndex(upperLine)
            SendMessage(Handle, WM_SetRedraw, 0, 0) 'запрет обновления
            'сохраняем выделение
            Dim selStart As Integer = SelectionStart
            Dim selLength As Integer = SelectionLength

            'стираем предыдущее выделение
            Me.Select(0, TextLength)
            SelectionBackColor = Color.FromArgb(0, 255, 255, 255)

            Me.Select(selStart, selLength) 'восстанавливаем выделение
            'возвращаем на место скролбар
            Dim charPos2 As Integer = GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})
            Dim lId2 As Integer = GetLineFromCharIndex(charPos2)
            SendMessage(Handle, EM_LINESCROLL, 0, upperLine - lId2)
            'разрешаем обновление
            SendMessage(Handle, WM_SetRedraw, 1, 0)
            Refresh()
            CanRaiseCodeEvents = True 'разрешаем TextChanged
            wasTextSelected = False
        End Sub

        ''' <summary>
        ''' Выделяет в тексте искомую строку
        ''' </summary>
        ''' <param name="strSeek">строка для поиска</param>
        ''' <param name="matchCase">учитывать регистр</param>
        ''' <param name="wholeWord">искать слово целиком</param>
        ''' <returns>Индекс первого вхождения или -1</returns>
        Public Function SeekString(ByVal strSeek As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean) As Integer
            If TextLength = 0 Then Return -1
            Dim compar As System.StringComparison = IIf(Not matchCase, StringComparison.CurrentCultureIgnoreCase, StringComparison.CurrentCulture)
            Dim selColor As Color = HightLightColor
            'стираем предыдущее выделение
            CanRaiseCodeEvents = False 'запрет TextChanged
            'получаем верхнюю видимую линию
            Dim upperLine As Integer = GetCharIndexFromPosition(New Point(5, 5))
            upperLine = GetLineFromCharIndex(upperLine)
            SendMessage(Handle, WM_SetRedraw, 0, 0) 'запрет обновления
            'сохраняем выделение
            Dim selStart As Integer = SelectionStart
            Dim selLength As Integer = SelectionLength

            'стираем предыдущее выделение
            Me.Select(0, TextLength)
            SelectionBackColor = Color.FromArgb(0, 255, 255, 255)

            Dim pos As Integer = -1
            Dim firstPos As Integer = -1
            Do
                pos = Text.IndexOf(strSeek, pos + 1, compar) 'ищем дальше
                If pos = -1 Then Exit Do
                If wholeWord Then
                    If pos > 0 AndAlso wordBoundsArray.Contains(Text.Chars(pos - 1)) = False Then
                        Continue Do 'это не слово целиком
                    End If

                    If pos < TextLength - 1 AndAlso wordBoundsArray.Contains(Text.Chars(pos + strSeek.Length)) = False Then
                        Continue Do 'это не слово целиком
                    End If
                End If
                If firstPos = -1 AndAlso pos >= 0 Then firstPos = pos
                'выделяем найденное
                Me.Select(pos, strSeek.Length)
                SelectionBackColor = DEFAULT_COLORS.codeBoxHighLight
            Loop

            Me.Select(selStart, selLength) 'восстанавливаем выделение
            'возвращаем на место скролбар
            Dim charPos2 As Integer = GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})
            Dim lId2 As Integer = GetLineFromCharIndex(charPos2)
            SendMessage(Handle, EM_LINESCROLL, 0, upperLine - lId2)
            'разрешаем обновление
            SendMessage(Handle, WM_SetRedraw, 1, 0)
            Refresh()
            CanRaiseCodeEvents = True 'разрешаем TextChanged
            wasTextSelected = True
            Return firstPos
        End Function

#End Region

#Region "List And Information"

        ''' <summary>
        ''' Процедура отображает информацию по функции, которая печатается
        ''' </summary>
        ''' <param name="classId">Id класса в mainClass</param>
        ''' <param name="elementName">имя функции/свойства</param>
        ''' <param name="elementType">тип (функция или свойство)</param>
        ''' <param name="paramNumber">порядковый номер параметра, который сейчас вводится</param>
        Public Sub ShowElementInfo(Optional ByVal classId As Integer = -1, Optional ByVal elementName As String = "", _
                                Optional ByVal elementType As EditWordTypeEnum = EditWordTypeEnum.W_NOTHING, _
                                Optional ByVal paramNumber As Integer = -1)

            'получаем документ html
            If wbRtbHelp.Url.ToString <> rtbHelpFile Then
                wbRtbHelp.Navigate(HelpFile)
                Do While IsNothing(wbRtbHelp.Url) OrElse wbRtbHelp.Url.ToString <> HelpFile
                    System.Windows.Forms.Application.DoEvents()
                Loop
            End If
            Dim hDoc As HtmlDocument = wbRtbHelp.Document
            My.Application.DoEvents()
            hDoc.GetElementById("funcInfo").Style = "display:block"
            hDoc.GetElementById("errors").Style = "display:none"
            If elementType = EditWordTypeEnum.W_NOTHING Then
                'elementType не указан - очищаем документ
                hDoc.GetElementById("elementDescription").InnerHtml = ""
                hDoc.GetElementById("elementCode").InnerHtml = ""
                hDoc.GetElementById("parameterInfo").InnerHtml = ""
                Return
            End If

            Select Case elementType
                Case EditWordTypeEnum.W_VARIABLE
                    'Здесь вывод инфо о перменной
                    elementName = Trim(elementName)
                    If elementName.StartsWith("'") Then elementName = mScript.PrepareStringToPrint(elementName, Nothing, False)
                    hDoc.GetElementById("elementCode").InnerHtml = "Var." & elementName
                    Dim isPublic As Boolean = False
                    If mScript.csPublicVariables.lstVariables.ContainsKey(elementName) Then
                        'переменная есть в глобальных
                        'Dim arrLocal As List(Of String) = CodeTextBox.MakeVariablesList(Me, False, True).ToList
                        'If arrLocal.Contains(elementName, StringComparer.CurrentCultureIgnoreCase) = False Then
                        '    'локальная переменная, заменяющая шлобальную с таким же именем
                        '    isPublic = True
                        'End If
                        isPublic = True
                    Else
                        'переменная локальная
                    End If

                    If isPublic Then
                        Dim VarDesc As String = mScript.csPublicVariables.lstVariables(elementName).Description
                        If IsNothing(VarDesc) Then VarDesc = ""
                        VarDesc = "Глобальная переменная. " & VarDesc
                        hDoc.GetElementById("elementDescription").InnerHtml = VarDesc
                    Else
                        hDoc.GetElementById("elementDescription").InnerHtml = "Локальная переменная"
                    End If

                    Return
                Case EditWordTypeEnum.W_BLOCK_DOWHILE
                    hDoc.GetElementById("elementDescription").InnerHtml = "Цикл Do ... While. Повторяет событие внутри цикла до тех пор, пока соблюдается условие. <a href='" & _
                        APP_HELP_PATH & "General\Do.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "Do [While Условие]<br/>&nbsp;&nbsp;...<br/>&nbsp;&nbsp;[Continue]</br>&nbsp;&nbsp;[Break [number]]<br/>Loop"
                    Return
                Case EditWordTypeEnum.W_BLOCK_EVENT
                    hDoc.GetElementById("elementDescription").InnerHtml = "Блок Event. Создает событие с указанным именем для указанного элемента. <a href='" & _
                        APP_HELP_PATH & "General\Event.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "Event 'Имя свойства-события' [, Имя/Id элемента, Имя/Id дочернего элемента]<br/>&nbsp;&nbsp;...<br/>End Event"
                    Return
                Case EditWordTypeEnum.W_BLOCK_FOR
                    hDoc.GetElementById("elementDescription").InnerHtml = "Цикл For ... Next. Повторяет скрипт внутри цикла заданное количество раз. <a href='" & _
                        APP_HELP_PATH & "General\For.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "For счетчик = начальное_значение  To конечное_значение  [Step шаг]<br/>&nbsp;&nbsp;[Continue]</br>&nbsp;&nbsp;[Break [number]]<br/>Next [счетчик]"
                    Return
                Case EditWordTypeEnum.W_BLOCK_FUNCTION
                    hDoc.GetElementById("elementDescription").InnerHtml = "Блок Function. Создает простую функцию с указанным именем или же функцию класса. <a href='" & _
                        APP_HELP_PATH & "General\Function.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "<table width='80%'><tr><td>Простая функция</td><td>Функция класса</td></tr><tr><td>Function 'название'<br/>&nbsp;&nbsp;...<br/>End Function</td><td>Function 'имя_класса.название'<br/>&nbsp;&nbsp;...<br/>End Function</td></tr></table>"
                    Return
                Case EditWordTypeEnum.W_BLOCK_IF
                    hDoc.GetElementById("elementDescription").InnerHtml = "Блок If ... Then. Однострочный блок исполняет однострочный код только в том случае, если условие оказалось равным True. Многострочный блок выполняет блок кода под тем условием, которое оказалось равным True. <a href='" & _
                        APP_HELP_PATH & "General\If.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "<table width='80%'><tr><td>Однострочный блок</td><td>Многострочный блок</td></tr><tr><td>If Условие Then однострочный код</td><td>If Условие Then<br/>...<br/>[ElseIf еще условие Then]<br/>...<br/>[Else]<br/>...<br/>End If</td></tr></table>"
                    Return
                Case EditWordTypeEnum.W_BLOCK_NEWCLASS, EditWordTypeEnum.W_REM_CLASS
                    hDoc.GetElementById("elementDescription").InnerHtml = "Блок New Class создает новый класс Писателя. Блок Rem Class удаляет класс. <a href='" & _
                        APP_HELP_PATH & "General\New.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "New Class 'имя_класса'<br/>&nbsp;&nbsp;[Name 'имя_класса']</br>&nbsp;&nbsp;Prop 'имя_свойства'[, значение]<br/></br>&nbsp;&nbsp;Prop 'имя_свойства'[ = значение]<br/>&nbsp;&nbsp;Func 'имя_функции', 'делегат'<br/>&nbsp;&nbsp;Function 'имя_функции'<br/>&nbsp;&nbsp;&nbsp;&nbsp;...<br/>&nbsp;&nbsp;End Function<br/>End Class<br/><br/>Rem Class 'ClassName'"
                    Return
                Case EditWordTypeEnum.W_HTML
                    hDoc.GetElementById("elementDescription").InnerHtml = "Блок HTML. Выводит html-текст, помещенный внутрь блока, в главное окно игры. <a href='" & _
                        APP_HELP_PATH & "General\html.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "HTML [Append 'Id элемента']<br/>&nbsp;&nbsp;...<br/>&nbsp;&nbsp;[&lt;exec&gt;скрипт&lt;/exec&gt;]<br/>&nbsp;&nbsp;...<br/>End HTML"
                    Return
                Case EditWordTypeEnum.W_WRAP
                    hDoc.GetElementById("elementDescription").InnerHtml = "Блок Wrap. Помещает блок html-текста в указанную переменную. <a href='" & _
                        APP_HELP_PATH & "General\Wrap.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "Wrap переменная<br/>&nbsp;&nbsp;...<br/>&nbsp;&nbsp;[&lt;exec&gt;скрипт&lt;/exec&gt;]<br/>&nbsp;&nbsp;...<br/>End Wrap"
                    Return
                Case EditWordTypeEnum.W_JUMP, EditWordTypeEnum.W_MARK
                    hDoc.GetElementById("elementDescription").InnerHtml = "Оператор Jump. Передает управление в строку скрипта, перед которой расположена указанная метка. <a href='" & _
                        APP_HELP_PATH & "General\Jump.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "Jump 'метка'<br/>...<br/><u>метка:</u>"
                    Return
                Case EditWordTypeEnum.W_SWITCH
                    hDoc.GetElementById("elementDescription").InnerHtml = "Блок Select Case / Switch. Выполняет блок кода, значения которого первыми подошли под выражение. <a href='" & _
                        APP_HELP_PATH & "General\Select.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "<table width='80%'><tr><td>Первый вариант записи</td><td>Второй вариант записи</td></tr><tr><td>Select Case Выражение<br/>Case значения<br/>&nbsp;&nbsp;...<br/>Case [другие значения]<br/>&nbsp;&nbsp;...<br/>[Case Else]<br/>&nbsp;&nbsp;...<br/>End Select</td><td>Switch Выражение<br/>Case значения<br/>&nbsp;&nbsp;...<br/>Case [другие значения]<br/>&nbsp;&nbsp;...<br/>[Case Else]<br/>&nbsp;&nbsp;...<br/>End Switch</td></tr></table>"
                    Return
                Case EditWordTypeEnum.W_RETURN
                    hDoc.GetElementById("elementDescription").InnerHtml = "Оператор Return. Совершает немедленное завершение текущего скрипта с возможностью передачи некоторого значения внешнему коду. <a href='" & _
                        APP_HELP_PATH & "General\Return.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "Return [значение]"
                    Return
                Case EditWordTypeEnum.W_EXIT
                    hDoc.GetElementById("elementDescription").InnerHtml = "Оператор Exit. Совершает немедленное завершение текущего скрипта. <a href='" & _
                        APP_HELP_PATH & "General\Return.html'>Подробнее...</a>"
                    hDoc.GetElementById("elementCode").InnerHtml = "Exit"
                    Return
                Case EditWordTypeEnum.W_CONTINUE
                    hDoc.GetElementById("elementDescription").InnerHtml = "Оператор Continue. Прекращает виток текущего цикла и сразу переходит к новой итерации. Используется в циклах <a href='" & _
                        APP_HELP_PATH & "General\Do.html'>Do ... While</a>, <a href='" & APP_HELP_PATH & "General\For.html'>For ... Next</a>."
                    hDoc.GetElementById("elementCode").InnerHtml = "Continue"
                    Return
                Case EditWordTypeEnum.W_BREAK
                    hDoc.GetElementById("elementDescription").InnerHtml = "Оператор Break. Немедленно прекращает текущий цикл(ы). Используется в циклах <a href='" & _
                        APP_HELP_PATH & "General\Do.html'>Do ... While</a>, <a href='" & APP_HELP_PATH & "General\For.html'>For ... Next</a>."
                    hDoc.GetElementById("elementCode").InnerHtml = "Break [number]"
                    Return
            End Select

            If classId > mScript.mainClass.Count - 1 Then Return
            'получаем инфо о нужной функции или свойстве
            Dim func As MatewScript.PropertiesInfoType = Nothing
            If elementName.Length > 0 AndAlso classId > -1 Then
                elementName = elementName.Trim
                If elementType = EditWordTypeEnum.W_FUNCTION Then
                    mScript.mainClass(classId).Functions.TryGetValue(elementName, func)
                Else
                    mScript.mainClass(classId).Properties.TryGetValue(elementName, func)
                End If
            End If

            If IsNothing(func) Then
                'функция/свойство не найдено - очищаем документ html
                hDoc.GetElementById("elementDescription").InnerHtml = ""
                hDoc.GetElementById("elementCode").InnerHtml = ""
                hDoc.GetElementById("parameterInfo").InnerHtml = ""
                Return
            End If

            'выводим описание елемента
            Dim strInnerHTML As String = ""
            Dim hEl As HtmlElement
            hEl = hDoc.GetElementById("elementDescription")
            If IsNothing(func.helpFile) Then func.helpFile = "" 'Return
            If func.helpFile.Length > 0 Then
                'получаем ссылку на файл помощи
                'If func.helpFile.IndexOf(":\") > -1 Then
                '    strInnerHTML = " <a href='" + func.helpFile + "'>Подробнее...</a>"
                'Else
                '    strInnerHTML = " <a href='" + HelpPath + func.helpFile + "'>Подробнее...</a>"
                'End If
                strInnerHTML = " <a href='" + GetHelpPath(func.helpFile) + "'>Подробнее...</a>"
            End If
            'выводим описание со сылкой на помощь
            Dim desc As String = func.Description
            If IsNothing(desc) Then desc = ""
            If elementType = EditWordTypeEnum.W_PROPERTY Then
                If desc.IndexOf("[Level") > -1 Then
                    'При наличии разных описаний для разных уровней выводим все здесь
                    Dim l1 As String = GetPropertyDescription(desc, 0)
                    Dim l2 As String = GetPropertyDescription(desc, 1)
                    Dim l3 As String = ""
                    If mScript.mainClass(classId).LevelsCount = 2 Then
                        l3 = GetPropertyDescription(desc, 2)
                    End If
                    desc = "<table width=100%>"
                    Dim strClassName As String = mScript.mainClass(classId).Names.Last
                    If l1.Length > 0 Then
                        desc += "<tr><td class='elementCode'>" + strClassName + "[<font size='-1'>-1</font>]." + elementName + "</td><td> " + l1 + "</td></tr>"
                    End If
                    If l2.Length > 0 Then
                        desc += "<tr><td class='elementCode'>" + strClassName + "[Name/Id]." + elementName + "</td><td> " + l2 + "</td></tr>"
                    End If
                    If l3.Length > 0 Then
                        desc += "<tr><td class='elementCode'>" + strClassName + "[Name/Id, Name/Id]." + elementName + "</td><td> " + l3 + "</td></tr>"
                    End If
                    desc += "</table>"
                    strInnerHTML = desc + strInnerHTML
                    hDoc.GetElementById("elementDescription").InnerHtml = strInnerHTML
                    hDoc.GetElementById("elementCode").InnerHtml = ""
                    hDoc.GetElementById("parameterInfo").InnerHtml = ""
                    Return
                End If
            ElseIf elementType = EditWordTypeEnum.W_FUNCTION Then
                If desc.IndexOf("[Return]") > -1 Then
                    'Есть ключевое слово [Return]. Обрабатываем его
                    Dim d As String = GetFunctionDescription(desc)
                    Dim r As String = "", rType As String = ""
                    GetFunctionDescriptionReturnData(desc, r, rType)
                    desc = d + "<br/><b>Возвращает</b><code>: " + rType + " </code> " + r
                End If
            End If
            hEl.InnerHtml = desc + strInnerHTML

            'выводим елемент с его параметрами
            strInnerHTML = ""
            If classId > -1 Then
                strInnerHTML = mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.GetUpperBound(0)) + "."
            End If
            strInnerHTML += elementName + " " 'имя элемента
            If IsNothing(func.params) = False Then
                Dim strParam As String = ""
                'вписываем параметры
                For i = 0 To func.params.GetUpperBound(0)
                    If func.params(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY Then
                        strParam = "[" + func.params(i).Name + "1, " + func.params(i).Name + "2, " + func.params(i).Name + "...]"
                    Else
                        If i > func.paramsMin - 1 Then
                            strParam = "[" + func.params(i).Name + "]"
                        Else
                            strParam = func.params(i).Name
                        End If
                    End If
                    'текущий параметр выделяем жирным
                    If paramNumber - 1 = i Then
                        strParam = "<b>" + strParam + "</b>"
                    ElseIf paramNumber > i AndAlso func.params(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY Then
                        strParam = "<b>" + strParam + "</b>"
                    End If
                    strInnerHTML += strParam
                    If i < func.params.GetUpperBound(0) Then strInnerHTML += ", "
                Next
            End If
            hDoc.GetElementById("elementCode").InnerHtml = strInnerHTML

            If elementType = EditWordTypeEnum.W_FUNCTION AndAlso IsNothing(func.params) Then
                hDoc.GetElementById("parameterInfo").InnerHtml = "функция не имеет параметров"
                Return
            ElseIf elementType = EditWordTypeEnum.W_PROPERTY Then
                Return
            End If
            'выводим описание текущего параметра
            strInnerHTML = ""
            Dim paramId As Integer = -1
            Dim paramsDescribed As Integer = 0
            If IsNothing(func.params) = False Then paramsDescribed = func.params.Count
            If func.paramsMax = -1 Then
                'массив параметров
                If paramNumber > paramsDescribed Then
                    'в доп. элементах массива параметров
                    paramId = func.params.GetUpperBound(0)
                    strInnerHTML = func.params(paramId).Description
                Else
                    'в обязательных элементах
                    paramId = paramNumber - 1
                    If paramId >= 0 Then strInnerHTML = func.params(paramId).Description
                End If
            Else
                If paramNumber > func.paramsMax Then
                    'лишний параметр (ошибка)
                    strInnerHTML = "<font color='#ff0000'>Превышено допустимое количество параметров функции.</font>"
                Else
                    paramId = paramNumber - 1
                    If paramId >= 0 Then strInnerHTML = func.params(paramId).Description
                End If
            End If

            If paramId > -1 Then
                Select Case func.params(paramId).Type
                    Case MatewScript.paramsType.paramsTypeEnum.PARAM_ANY
                        strInnerHTML += " (любой тип)"
                    Case MatewScript.paramsType.paramsTypeEnum.PARAM_BOOL
                        strInnerHTML += " (Да/Нет)"
                    Case MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM
                        strInnerHTML += " (один из вариантов)"
                    Case MatewScript.paramsType.paramsTypeEnum.PARAM_INTEGER
                        strInnerHTML += " (целое число)"
                    Case MatewScript.paramsType.paramsTypeEnum.PARAM_SINGLE
                        strInnerHTML += " (число)"
                    Case MatewScript.paramsType.paramsTypeEnum.PARAM_STRING
                        strInnerHTML += " (строка)"
                    Case MatewScript.paramsType.paramsTypeEnum.PARAM_USER_FUNCTION
                        strInnerHTML += " (имя функции Писателя)"
                    Case MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY
                        strInnerHTML += " (массив параметров)"
                End Select
            End If
            hDoc.GetElementById("parameterInfo").InnerHtml = strInnerHTML
        End Sub


        Public Function IsOpenedTextBlockBetweenLines(ByVal startLine As Integer, ByVal endLine As Integer) As Boolean
            'Функция возвращает есть ли в коде между  линиями startLine и endLine включительно открытые блоки текста (HTML или Wrap)
            Dim htmlBalance As Integer = 0, wrapBalance As Integer = 0 'Баланс HTML & Wrap (разница между HTML ... End HTML и Wrap ... End Wrap)
            Dim fstWord As EditWordType
            For i As Integer = startLine To endLine 'проверяем все линии между указанными (включительно)
                fstWord = GetWordFromCodeData(i, 0) 'первое слово текущей линии
                Select Case fstWord.wordType
                    Case EditWordTypeEnum.W_HTML 'начало блока HTML
                        htmlBalance += 1
                    Case EditWordTypeEnum.W_WRAP 'начало блока Wrap
                        wrapBalance += 1
                    Case EditWordTypeEnum.W_CYCLE_END
                        fstWord = GetWordFromCodeData(i, 1)
                        If fstWord.wordType = EditWordTypeEnum.W_HTML Then 'конец блока HTML
                            htmlBalance -= 1
                        ElseIf fstWord.wordType = EditWordTypeEnum.W_WRAP Then 'конец блока Wrap
                            wrapBalance -= 1
                        End If
                End Select
            Next
            Return Not (htmlBalance = 0 And wrapBalance = 0)
        End Function

#End Region
    End Class 'RichTextBoxEx

#Region "Properties"

    ''' <summary>Раскрашивать ли слова или оставлять черными</summary>
    Public Property CanDrawWords As Boolean
        Get
            Return codeBox.CanDrawWords
        End Get
        Set(value As Boolean)
            codeBox.CanDrawWords = value
        End Set
    End Property


    <System.ComponentModel.Browsable(True)> _
    Public Property Multiline() As Boolean
        Get
            Return codeBox.Multiline
        End Get
        Set(ByVal value As Boolean)
            codeBox.Multiline = value
            SplitContainerVertical.Panel1Collapsed = Not value
            'masUndo.Push(codeBox.Rtf)
            'masUndoCode.Push(codeBox.CodeData)
            'MakeNextUndoItem = False
            'timUndo.Stop()
            'timUndo.Start()

            Invalidate()
        End Set
    End Property

    <System.ComponentModel.Browsable(True)> _
    Public Overrides Property Text() As String
        Get
            Return codeBox.Text
        End Get
        Set(ByVal value As String)
            codeBox.rtbCanRaiseEvents = False
            If IsNothing(value) Then value = ""
            If value.Length > rtbMaxLineLength Then
                'Если в тексте присутствуют очень длинные строки - разбиваем их
                Dim pos As Integer = 0
                Dim prevPos As Integer = 0
                Do
                    pos = value.IndexOf(vbCr, pos)
                    If pos = -1 Then
                        If value.Length - prevPos > rtbMaxLineLength Then
                            'вставляем символ разбиения строк и продолжаем с этого места
                            value = value.Substring(0, rtbMaxLineLength + prevPos) + vbCr + value.Substring(rtbMaxLineLength + prevPos)
                            pos = rtbMaxLineLength + 1 + prevPos
                            prevPos = pos
                            Continue Do
                        Else
                            Exit Do
                        End If
                    Else
                        If pos - prevPos > rtbMaxLineLength Then
                            'вставляем символ разбиения строк и продолжаем с этого места
                            value = value.Substring(0, rtbMaxLineLength + prevPos) + vbCr + value.Substring(rtbMaxLineLength + prevPos)
                            pos = rtbMaxLineLength + 1 + prevPos
                        Else
                            pos += 1
                        End If
                        prevPos = pos
                    End If
                Loop
            End If

            With codeBox
                .lastSelectionEndLine = 0
                .lastSelectionStartLine = 0
                .Text = value
                .PrepareText(codeBox)
                .rtbCanRaiseEvents = True
                .csUndo.BeginNew()
            End With

            If codeBox.Visible = False Then codeBox.Visible = True
        End Set
    End Property

    Public ReadOnly Property Lines() As String()
        Get
            Return codeBox.Lines
        End Get
    End Property

    <System.ComponentModel.Browsable(True)> _
    Public Property AutoWordSelection() As Boolean
        Get
            Return codeBox.AutoWordSelection
        End Get
        Set(ByVal value As Boolean)
            codeBox.AutoWordSelection = value
        End Set
    End Property
#End Region

#Region "Нумерация строк"


    Private Sub timDraw_Tick(sender As Object, e As EventArgs) Handles timDraw.Tick
        Dim i As Integer = 0
        Dim cntrl As Control
        For i = SplitContainerVertical.Panel1.Controls.Count - 1 To 0 Step -1
            cntrl = SplitContainerVertical.Panel1.Controls(i)
            If cntrl.GetType.Name = "Label" Then
                SplitContainerVertical.Panel1.Controls.Remove(cntrl)
            End If
        Next
        Dim labelPosY As Integer
        i = 0
        Dim firstVisibleLine As Integer
        firstVisibleLine = codeBox.GetLineFromCharIndex(codeBox.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})) + 1
        Do
            labelPosY = codeBox.Top + codeBox.Font.Height * i - 3
            If labelPosY + 20 > codeBox.Parent.Height Then Exit Do
            Dim lbl As New Label With {.Text = (i + firstVisibleLine).ToString, .Font = codeBox.Font, .Left = 5, .Top = labelPosY + 2, .Padding = New Padding(2)}
            lbl.BringToFront()
            'Dim lbl As New Label With {.Text = (i + firstVisibleLine).ToString, .Font = codeBox.Font, .Left = 5, .Top = labelPosY + 2, .TextAlign = ContentAlignment.MiddleLeft}
            'AddHandler lbl.MouseDown, AddressOf del_lbl_MouseDown
            'AddHandler lbl.MouseHover, AddressOf del_lbl_MouseEnter
            'AddHandler lbl.MouseUp, AddressOf del_lbl_MouseUp
            SplitContainerVertical.Panel1.Controls.Add(lbl)
            i += 1
        Loop
        codeBox.Size = New Size(SplitContainerVertical.Panel2.Width, SplitContainerVertical.Panel2.Height)

        timDraw.Enabled = False
    End Sub

    Private Sub codeBox_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles codeBox.MouseDoubleClick
        codeBox.SelectionLength = 0
        Dim loc As Integer = codeBox.GetCharIndexFromPosition(e.Location)
        If loc > -1 Then codeBox.SelectionStart = loc

        'wordBoundsArray
        Dim strText As String = codeBox.Text
        Dim sLength As Integer = strText.Length
        If sLength = 0 Then Return

        Dim selStart As Integer = codeBox.SelectionStart
        For i As Integer = codeBox.SelectionStart To 0 Step -1
            Dim ch As Char = strText.Chars(i)
            Dim pos As Integer = Array.IndexOf(codeBox.wordBoundsArray, ch)
            If pos > -1 Then
                If i > 0 AndAlso ch = "'" Then
                    If strText.Chars(i - 1) = "\"c Then
                        selStart = i
                        Continue For
                    End If
                ElseIf ch = "\"c AndAlso pos < sLength - 1 Then
                    If strText.Chars(i + 1) = "'"c Then
                        selStart = i
                        Continue For
                    End If
                Else
                    Exit For
                End If
            Else
                selStart = i
            End If
        Next i

        Dim selEnd As Integer = codeBox.SelectionStart
        For i As Integer = codeBox.SelectionStart To sLength - 1
            Dim ch As Char = strText.Chars(i)
            Dim pos As Integer = Array.IndexOf(codeBox.wordBoundsArray, ch)
            If pos > -1 Then
                If i > 0 AndAlso ch = "'" Then
                    If strText.Chars(i - 1) = "\"c Then
                        selEnd = i
                        Continue For
                    End If
                ElseIf ch = "\"c AndAlso pos < sLength - 1 Then
                    If strText.Chars(i + 1) = "'"c Then
                        selEnd = i
                        Continue For
                    End If
                Else
                    Exit For
                End If
            Else
                selEnd = i
            End If
        Next i

        codeBox.Select(selStart, selEnd - selStart + 1)
    End Sub

    'Dim lblSelStartLine As Integer = -1
    'Dim lblSelEndLine As Integer = -1
    'Private Sub del_lbl_MouseDown(sender As Label, e As MouseEventArgs)
    '    If e.Button = Windows.Forms.MouseButtons.Left Then
    '        Dim lineNum As Integer = Val(sender.Text) - 1
    '        If lineNum > -1 AndAlso lineNum <= codeBox.Lines.Count - 1 Then
    '            lblSelStartLine = lineNum
    '            lblSelEndLine = lineNum
    '            Dim startPos As Integer = codeBox.GetFirstCharIndexFromLine(lineNum)
    '            Dim sLength As Integer = codeBox.Lines(lineNum).Length
    '            codeBox.Select(startPos, sLength)
    '        End If
    '    End If
    'End Sub

    'Private Sub del_lbl_MouseEnter(sender As Label, e As EventArgs)
    '    If lblSelStartLine = -1 Then Return
    '    Dim lineNum As Integer = Val(sender.Text) - 1
    '    frmMainEditor.Text = "EndLine: " & lblSelEndLine.ToString & ", lineNum: " & lineNum.ToString
    '    If lineNum = -1 OrElse lineNum > codeBox.Lines.Count - 1 OrElse lblSelEndLine = lineNum Then Return
    '    lblSelEndLine = lineNum
    '    Dim sStart As Integer = codeBox.GetFirstCharIndexFromLine(Math.Min(lblSelEndLine, lblSelStartLine))
    '    Dim endline As Integer = Math.Max(lblSelStartLine, lblSelEndLine)
    '    Dim sEnd As Integer = codeBox.GetFirstCharIndexFromLine(endline) + codeBox.Lines(endline).Length
    '    codeBox.Select(sStart, sEnd - sStart)
    'End Sub

    'Private Sub del_lbl_MouseUp(sender As Label, e As MouseEventArgs)
    '    lblSelEndLine = -1
    '    lblSelStartLine = -1
    'End Sub

    Private Sub rtb_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles codeBox.Resize
        'Расстановка нумерации
        If IsNothing(codeBox.Parent) Then Return
        timDraw.Enabled = True

    End Sub

    Private Sub codeBox_VScroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles codeBox.VScroll
        'смена номеров
        If sender.TextLength = 0 Then Return

        Dim firstCharIndex As Integer = codeBox.GetCharIndexFromPosition(New Point With {.X = 1, .Y = 1})
        Dim curLine As Integer = sender.GetLineFromCharIndex(firstCharIndex) + 1
        Dim firstCharTop As Integer = sender.GetPositionFromCharIndex(firstCharIndex).Y - 2
        Dim labelOffsetY As Integer
        Dim cntrl As Control

        Dim isFirstLable As Boolean = True
        For i As Integer = 0 To SplitContainerVertical.Panel1.Controls.Count - 1
            cntrl = SplitContainerVertical.Panel1.Controls(i)
            If cntrl.GetType.Name = "Label" Then
                If isFirstLable Then
                    labelOffsetY = cntrl.Top - firstCharTop
                    isFirstLable = False
                End If
                cntrl.Text = curLine.ToString
                cntrl.Top -= labelOffsetY
                curLine += 1
            End If
        Next
    End Sub

    Private Sub CheckLabels(ByRef rtb As RichTextBox)
        'Корректировка нумерации строк при редактировании текста
1:
        'получаем объект - первую надпись с номером
        Static firstLable As Control = Nothing
        If IsNothing(firstLable) Then
            Dim cntrl As Control
            For i As Integer = 0 To SplitContainerVertical.Panel1.Controls.Count - 1
                cntrl = SplitContainerVertical.Panel1.Controls(i)
                If cntrl.GetType.Name = "Label" Then
                    firstLable = cntrl
                    Exit For
                End If
            Next
        End If
        'If IsNothing(firstLable) Then
        '    Call timDraw_Tick(timDraw, New System.EventArgs)
        '    GoTo 1
        'End If
        Dim curLine As Integer = 1
        If IsNothing(firstLable) Then Return
        If rtb.TextLength = 0 AndAlso firstLable.Text <> "1" Then
            'удален весь текст - коррекция нумерации (начало нумерации с 1)
            Dim cntrl As Control
            For i As Integer = 0 To SplitContainerVertical.Panel1.Controls.Count - 1
                cntrl = SplitContainerVertical.Panel1.Controls(i)
                If cntrl.GetType.Name = "Label" Then
                    cntrl.Text = curLine.ToString
                    cntrl.Top = rtb.Top + rtb.Font.Height * (curLine - 1) - 3
                    cntrl.BringToFront()
                    curLine += 1
                End If
            Next
            Return
        End If

        Dim firstCharIndex As Integer = rtb.GetCharIndexFromPosition(New Point With {.X = 1, .Y = 1})
        curLine = rtb.GetLineFromCharIndex(firstCharIndex) + 1 'получили номер верхней видимой строки
        'если номер первой видимой строки изменился, вызываем событие VScroll
        If IsNothing(firstLable) = False AndAlso firstLable.Text <> curLine.ToString Then codeBox_VScroll(rtb, New System.EventArgs)
    End Sub
#End Region

#Region "Rtb Events"
    Private Sub codeBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles codeBox.Click
        If IsNothing(mScript.mainClass) Then Return
        If codeBox.rtbCanRaiseEvents = False Then Return
        If lstRtb.Visible Then lstRtb.Hide()
        codeBox.CheckPrevLine(codeBox)
        'если был совершен переход на новую линию - сохраняем информацию о текущем содержимом строки, которую будут редактировать
        If codeBox.prevLineId <> codeBox.GetLineFromCharIndex(codeBox.GetFirstCharIndexOfCurrentLine) Then codeBox.SetPrevLine(codeBox)
    End Sub

    Private Sub codeBox_GotFocus(ByVal sender As RichTextBox, ByVal e As System.EventArgs) Handles codeBox.GotFocus
        If IsNothing(mScript.mainClass) Then Return
        If codeBox.rtbCanRaiseEvents = False Then Return
        codeBox.SetPrevLine(sender) 'сохраняем текущую линию до ее редактирования
    End Sub

    Private Sub codeBox_TextChanged(ByVal sender As RichTextBoxEx, ByVal e As System.EventArgs) Handles codeBox.TextChanged
        If IsNothing(mScript.mainClass) Then Return
        If codeBox.rtbCanRaiseEvents = False Then Return
        If codeBox.wasTextSelected Then codeBox.ClearPreviousSelection()
        Dim wasTextBlock As Boolean = False
        If IsNothing(codeBox.CodeData) = False AndAlso codeBox.CodeData.GetUpperBound(0) >= (codeBox.lastSelectionEndLine - codeBox.lastSelectionStartLine) AndAlso _
            codeBox.lastSelectionStartLine <> codeBox.lastSelectionEndLine AndAlso shortcutType = shortcutTypeEnum.NONE Then
            wasTextBlock = codeBox.IsOpenedTextBlockBetweenLines(codeBox.lastSelectionStartLine, codeBox.lastSelectionEndLine) 'был незакрытый текстовый блок в убранном тексте?
            If codeBox.IsSelectedTextUsed = False Then
                'Текст гарантированно не был вставлен с помощью SelectionText - он изменился из-за обычного ввода с клавиатуры. При этом перед нажатием на кнопку был выделен фрагмент кода
                'в несколько строк (и на данный момент он уже стерт и заменен введенный символом (или просто стерт, если был нажат Del или Backspace)
                'если были выделены символы на нескольких строках, то при редактировании они сотрутся (точнее, уже стерты).
                'Убираем из codeBox.CodeData лишнее
                Dim linesDeleted = codeBox.lastSelectionEndLine - codeBox.lastSelectionStartLine
                For i As Integer = codeBox.lastSelectionStartLine + 1 To codeBox.CodeData.GetUpperBound(0) - linesDeleted
                    codeBox.CodeData(i) = codeBox.CodeData(i + linesDeleted)
                Next
                ReDim Preserve codeBox.CodeData(codeBox.CodeData.GetUpperBound(0) - linesDeleted)
                codeBox.lastSelectionEndLine = codeBox.lastSelectionStartLine
            End If
        ElseIf shortcutType <> shortcutTypeEnum.NONE Then
            'Комбинация Ctrl + Bckspace/Del
            Dim diff As Integer = shortcutLinesBefore - sender.Lines.Count
            If diff > 0 Then
                Dim lstData As List(Of CodeDataType) = sender.CodeData.ToList
                lstData.RemoveRange(sender.GetLineFromCharIndex(sender.SelectionStart), diff)
                sender.CodeData = lstData.ToArray
            End If
            If sender.csUndo.lastSelection.Length > 0 Then
                sender.csUndo.UndoInProcess = True
                sender.Select(sender.SelectionStart, sender.csUndo.lastSelection.Length)
                sender.SelectedText = ""
                sender.csUndo.UndoInProcess = False
            End If
        End If

        Dim currentLine As Integer = codeBox.GetLineFromCharIndex(codeBox.GetFirstCharIndexOfCurrentLine) 'получаем текущую линию
        If lastKey = 0 Then
            SendMessage(codeBox.Handle, WM_SetRedraw, 0, 0)
            If wasTextBlock OrElse (codeBox.csUndo.UndoInProcess AndAlso codeBox.csUndo.IsLastUndoContainsTextBlock) Then
                'если в результате последнего редактирования был удален кусок текстового блока, перерисовываем весь текст от текущей динии и до конца
                codeBox.CheckPrevLine(sender, sender.GetLineFromCharIndex(sender.SelectionStart), sender.Lines.GetUpperBound(0), True, True) 'проверка синтаксиса и раскраска измененных строк
            ElseIf codeBox.csUndo.UndoInProcess Then
                'выполняется Undo надо проверить предыдущую строку (возможно, потребуется перерисовка большого блока)
                codeBox.CheckPrevLine(sender, sender.GetLineFromCharIndex(sender.SelectionStart), sender.GetLineFromCharIndex(sender.SelectionStart + sender.SelectionLength), True, True) 'проверка синтаксиса и раскраска измененных строк
            End If
            'перерисовываем текст с пометкой "незавершенная строка"
            Dim currentWordId As Integer
            Dim upperLine As Integer = sender.GetLineFromCharIndex(sender.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5}))
            If codeBox.IsSelectedTextUsed Then
                codeBox.PrepareText(sender, codeBox.lastSelectionStartLine, codeBox.lastSelectionEndLine, True, currentWordId)
            Else
                codeBox.PrepareTextNonCompleted(sender, sender.GetLineFromCharIndex(sender.SelectionStart), sender.GetLineFromCharIndex(sender.SelectionStart), True, currentWordId)
            End If
            Dim charPos As Integer = sender.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})
            Dim lId As Integer = sender.GetLineFromCharIndex(charPos)
            SendMessage(sender.Handle, EM_LINESCROLL, 0, upperLine - lId)
            SendMessage(codeBox.Handle, WM_SetRedraw, 1, 0)
            codeBox.Refresh()


            Dim curWord As EditWordType = Nothing
            Dim paramNumber As Integer
            Dim classId As Integer
            Dim wordId As Integer = -1, wordLine As Integer = currentLine
            GetCurrentElementWord(currentLine, currentWordId, paramNumber, curWord, classId, wordId, wordLine)
            PrepareRtbList(currentLine, currentWordId)
            'ShowElementInfo(curWord.classId, curWord.Word, curWord.wordType, paramNumber)
            If curWord.wordType <> EditWordTypeEnum.W_NOTHING OrElse (lstRtb.Visible = True AndAlso IsNothing(lstRtb.Tag) = False) = False Then
                If curWord.wordType = EditWordTypeEnum.W_NOTHING AndAlso currentWordId > -1 AndAlso IsNothing(codeBox.CodeData(wordLine).Code) = False Then
                    curWord = codeBox.CodeData(wordLine).Code(currentWordId)
                    If curWord.wordType = EditWordTypeEnum.W_VARIABLE AndAlso curWord.Word.EndsWith(" "c) = False Then
                        curWord.wordType = EditWordTypeEnum.W_NOTHING
                        curWord.Word = Nothing
                        curWord.classId = 0
                    End If
                End If
                codeBox.ShowElementInfo(curWord.classId, curWord.Word, curWord.wordType, paramNumber)
            End If
        ElseIf lastKey = Keys.Return Then
            'Добавляем линию вслед за текущей
            ReDim Preserve codeBox.CodeData(codeBox.CodeData.Length) 'расширяем массив на 1
            For i As Integer = codeBox.CodeData.GetUpperBound(0) To currentLine + 1 Step -1
                codeBox.CodeData(i) = codeBox.CodeData(i - 1) 'сдвигаем инфо в codeBox.CodeData на 1 вправо (в конец), начиная со строки, следующей за текущей
            Next
            codeBox.CheckPrevLine(sender, currentLine - 1, IIf(wasTextBlock, sender.Lines.GetUpperBound(0), currentLine), True) 'проверка синтаксиса и раскраска измененных строк
            If codeBox.SelectionLength = 0 Then
                'codeBox.csUndo.lastSelection = codeBox.SelectedText
                codeBox.SelectionStart -= 1
                codeBox.csUndo.AppendItem(UndoClass.UndoTypeEnum.CHAR_ENTRY, vbLf)
                codeBox.SelectionStart += 1
            Else
                codeBox.csUndo.lastSelection = codeBox.SelectedText
                codeBox.csUndo.AppendItem(UndoClass.UndoTypeEnum.CHAR_ENTRY, vbLf)
            End If
            'если был совершен переход на новую линию - сохраняем информацию о текущем содержимом строки, которую будут редактировать
            If codeBox.prevLineId <> codeBox.GetLineFromCharIndex(codeBox.GetFirstCharIndexOfCurrentLine) Then codeBox.SetPrevLine(codeBox)
            lastKey = 0
            'расставляем правильные начальные пробелы
            codeBox.SetCorretSpaces(codeBox)
        ElseIf lastKey = Keys.Back OrElse lastKey = Keys.Delete Then
            'Удаляем линию за текущей
            If shortcutType = shortcutTypeEnum.NONE Then
                'обычное нажатие Bckspace / Del
                For i As Integer = currentLine + 1 To codeBox.CodeData.GetUpperBound(0) - 1
                    codeBox.CodeData(i) = codeBox.CodeData(i + 1) 'сдвигаем инфо в codeBox.CodeData на 1 вправо (в конец), начиная со строки, следующей за текущей
                Next
                If codeBox.CodeData.Length = 0 Then
                    Erase codeBox.CodeData
                Else
                    ReDim Preserve codeBox.CodeData(codeBox.CodeData.GetUpperBound(0) - 1) 'уменьшаем массив на 1
                End If
            End If

            codeBox.CheckPrevLine(sender, currentLine, IIf(wasTextBlock, sender.Lines.GetUpperBound(0), currentLine + 1), True) 'проверка синтаксиса и раскраска измененных строк
            'если был совершен переход на новую линию - сохраняем информацию о текущем содержимом строки, которую будут редактировать
            If codeBox.prevLineId <> codeBox.GetLineFromCharIndex(codeBox.GetFirstCharIndexOfCurrentLine) Then codeBox.SetPrevLine(codeBox)
            lastKey = 0
        End If

        codeBox.IsSelectedTextUsed = False
        CheckLabels(codeBox)

        'If Not codeBox.csUndo.UndoInProcess Then
        '    'Undo


        '    If MakeNextUndoItem = False And masUndo.Count > 1 Then
        '        masUndo.Pop()
        '        masUndoCode.Pop()
        '    End If
        '    masUndo.Push(codeBox.Rtf)
        '    masUndoCode.Push(CopyCodeDataArray(codeBox.CodeData))
        '    'Form1.Button1_Click(Form1.Button1, New System.EventArgs)
        '    MakeNextUndoItem = False
        '    timUndo.Stop()
        '    timUndo.Start()

        'End If

        RaiseEvent TextChanged(sender, e)
    End Sub

    ''' <summary>Получаем слово у каретки </summary>
    ''' <param name="lineId">для получения Id линии текущего слова</param>
    ''' <param name="wordId">для получения Id текущего слова</param>
    Public Function GetWordInCaret(Optional ByRef lineId As Integer = Nothing, Optional ByRef wordId As Integer = Nothing) As EditWordType
        Dim curLine As Integer = codeBox.GetLineFromCharIndex(codeBox.SelectionStart)
        lineId = curLine
        Dim charId As Integer = codeBox.SelectionStart - codeBox.GetFirstCharIndexOfCurrentLine
        If IsNothing(codeBox.CodeData) OrElse curLine > codeBox.CodeData.Count - 1 Then Return Nothing
        Dim cd As CodeDataType = codeBox.CodeData(curLine)
        Dim curPos As Integer = 0
        If String.IsNullOrEmpty(cd.StartingSpaces) = False Then curPos = cd.StartingSpaces.Length
        If charId < curPos Then Return Nothing 'мы в начальных пробелах
        If IsNothing(cd.Code) OrElse cd.Code.Length = 0 Then Return Nothing

        For wId As Integer = 0 To cd.Code.Length - 1
            curPos += cd.Code(wId).Word.Length
            If charId <= curPos Then
                wordId = wId
                Return cd.Code(wId)
            End If
        Next
        Return Nothing
    End Function

    Private Sub codeBox_MouseClick(sender As RichTextBox, e As MouseEventArgs) Handles codeBox.MouseClick
        If sender.TextLength = 0 Then Return
        Dim selColor As Color = codeBox.HightLightColor
        'стираем предыдущее выделение
        codeBox.ClearPreviousSelection()
        'codeBox.PrepareHTMLforErrorInfo(codeBox.wbRtbHelp)

        Dim curLine As Integer = Nothing
        Dim curWordId As Integer = Nothing
        Dim curWord As EditWordType = GetWordInCaret(curLine, curWordId) 'получаем слово, на котором стоит каретка
        If IsNothing(curWord) OrElse curWord.wordType = EditWordTypeEnum.W_NOTHING Then Return

        If IsNothing(selColor) Then
            'не выделяем цветом - только показываем справку
            Select Case curWord.wordType
                Case EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_BLOCK_DOWHILE, EditWordTypeEnum.W_BLOCK_EVENT, EditWordTypeEnum.W_BLOCK_FOR, _
                    EditWordTypeEnum.W_BLOCK_FUNCTION, EditWordTypeEnum.W_BLOCK_IF, EditWordTypeEnum.W_BLOCK_NEWCLASS, EditWordTypeEnum.W_HTML, EditWordTypeEnum.W_JUMP, EditWordTypeEnum.W_MARK, _
                    EditWordTypeEnum.W_REM_CLASS, EditWordTypeEnum.W_SWITCH, EditWordTypeEnum.W_WRAP, EditWordTypeEnum.W_EXIT, EditWordTypeEnum.W_RETURN, EditWordTypeEnum.W_CONTINUE, EditWordTypeEnum.W_BREAK
                    codeBox.ShowElementInfo(curWord.classId, curWord.Word.Trim, curWord.wordType)
            End Select
            Return
        End If

        If curWord.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(curLine).Code.Count = 2 Then
            curWord = codeBox.CodeData(curLine).Code(1)
            curWordId = 1
        ElseIf curWord.wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso curWordId = 1 AndAlso codeBox.CodeData(curLine).Code(0).wordType = EditWordTypeEnum.W_SWITCH AndAlso String.Compare(curWord.Word.Trim, "Else", True) = 0 Then
            curWord = codeBox.CodeData(curLine).Code(0)
            curWordId = 0
        End If
        Dim strWord As String = curWord.Word.Trim

        Dim upperLine As Integer
        Dim selStart As Integer
        Dim selLength As Integer
        Select Case curWord.wordType
            Case EditWordTypeEnum.W_VARIABLE, EditWordTypeEnum.W_SIMPLE_STRING, EditWordTypeEnum.W_PROPERTY, EditWordTypeEnum.W_FUNCTION
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                For lineId As Integer = 0 To codeBox.CodeData.Length - 1
                    'перебираем все строки
                    If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                    Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                    For wId As Integer = 0 To codeBox.CodeData(lineId).Code.Count - 1
                        'перебираем все слова с троке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(wId)
                        If w.wordType = curWord.wordType AndAlso w.classId = curWord.classId AndAlso String.Compare(w.Word.Trim, strWord, True) = 0 Then
                            'Найдено совпадение                        
                            Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        End If
                        pos += w.Word.Length
                    Next wId
                Next lineId

                If curWord.wordType <> EditWordTypeEnum.W_SIMPLE_STRING Then
                    codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
                End If
            Case EditWordTypeEnum.W_MARK, EditWordTypeEnum.W_JUMP
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                For lineId As Integer = 0 To codeBox.CodeData.Length - 1
                    'перебираем все строки
                    If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                    Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                    For wId As Integer = 0 To codeBox.CodeData(lineId).Code.Count - 1
                        'перебираем все слова с троке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(wId)
                        If w.wordType = EditWordTypeEnum.W_JUMP OrElse w.wordType = EditWordTypeEnum.W_MARK Then
                            'Найдено совпадение                        
                            Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        End If
                        pos += w.Word.Length
                    Next wId
                Next lineId

                codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
            Case EditWordTypeEnum.W_RETURN, EditWordTypeEnum.W_EXIT
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                For lineId As Integer = 0 To codeBox.CodeData.Length - 1
                    'перебираем все строки
                    If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                    Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                    For wId As Integer = 0 To codeBox.CodeData(lineId).Code.Count - 1
                        'перебираем все слова с троке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(wId)
                        If w.wordType = EditWordTypeEnum.W_RETURN OrElse w.wordType = EditWordTypeEnum.W_EXIT Then
                            'Найдено совпадение                        
                            Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        End If
                        pos += w.Word.Length
                    Next wId
                Next lineId
                codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
            Case EditWordTypeEnum.W_PARAM, EditWordTypeEnum.W_PARAM_COUNT
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                For lineId As Integer = 0 To codeBox.CodeData.Length - 1
                    'перебираем все строки
                    If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                    Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                    For wId As Integer = 0 To codeBox.CodeData(lineId).Code.Count - 1
                        'перебираем все слова с троке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(wId)
                        If w.wordType = EditWordTypeEnum.W_PARAM OrElse w.wordType = EditWordTypeEnum.W_PARAM_COUNT Then
                            'Найдено совпадение                        
                            Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        End If
                        pos += w.Word.Length
                    Next wId
                Next lineId
            Case EditWordTypeEnum.W_CLASS
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                For lineId As Integer = 0 To codeBox.CodeData.Length - 1
                    'перебираем все строки
                    If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                    Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                    For wId As Integer = 0 To codeBox.CodeData(lineId).Code.Count - 1
                        'перебираем все слова с троке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(wId)
                        If w.wordType = curWord.wordType AndAlso w.classId = curWord.classId Then
                            'Найдено совпадение классов                      
                            Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        End If
                        pos += w.Word.Length
                    Next wId
                Next lineId
            Case EditWordTypeEnum.W_BLOCK_DOWHILE
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                'выделяем текущее слово
                Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                For wId As Integer = 0 To curWordId - 1
                    pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                Next wId
                Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                sender.Select(firstCharId, strWord.Length) 'длина
                sender.SelectionBackColor = selColor

                Dim seekForward As Boolean = False
                If curWordId < codeBox.CodeData(curLine).Code.Length - 1 AndAlso strWord.ToLower = "do" AndAlso codeBox.CodeData(curLine).Code(curWordId + 1).wordType = EditWordTypeEnum.W_BLOCK_DOWHILE Then
                    'выделяем while
                    firstCharId += 3
                    strWord = codeBox.CodeData(curLine).Code(curWordId + 1).Word.Trim
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = True
                ElseIf curWordId > 0 AndAlso strWord.ToLower = "while" AndAlso codeBox.CodeData(curLine).Code(curWordId - 1).wordType = EditWordTypeEnum.W_BLOCK_DOWHILE Then
                    'выделяем do
                    firstCharId -= 3
                    strWord = codeBox.CodeData(curLine).Code(curWordId - 1).Word.Trim
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = True
                ElseIf strWord.ToLower = "do" Then
                    seekForward = True
                End If

                Dim internalCycles As Integer = 0
                Dim internalCyclesForBreak As Integer = 0
                If seekForward Then
                    'ищем Loop
                    For lineId As Integer = curLine + 1 To codeBox.CodeData.Length - 1
                        'перебираем все строки от следующей после текущей в поиске Loop нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                        pos = 0 'хранит длину обработанной части текущей строки
                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = curWord.wordType Then
                            'Найдено слово из блока
                            Dim wTrim As String = w.Word.Trim
                            If String.Compare(wTrim, "Loop", True) = 0 Then
                                'Найден Loop
                                If internalCycles <> 0 Then
                                    'это часть внутреннего блока
                                    internalCycles -= 1
                                Else
                                    'это часть нашего блока
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                    Exit For
                                End If
                            ElseIf String.Compare(wTrim, "Do", True) = 0 Then
                                'найден внутренний цикл
                                internalCycles += 1
                            End If
                        ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "next", True) <> 0) OrElse _
                            (w.wordType = EditWordTypeEnum.W_SWITCH AndAlso String.Compare(w.Word.Trim, "Case", True) <> 0) Then
                            internalCyclesForBreak += 1
                        ElseIf (w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code.Count = 2 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_SWITCH) OrElse _
                            (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "next", True) = 0) Then
                            internalCyclesForBreak -= 1
                        ElseIf internalCycles = 0 AndAlso internalCyclesForBreak = 0 Then
                            If w.wordType = EditWordTypeEnum.W_CONTINUE OrElse w.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка начинается с Continue / Break
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_CONTINUE OrElse codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка заканчивается на Continue / Break
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 2
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code.Last
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf (codeBox.CodeData(lineId).Code.Length > 2 AndAlso codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2).wordType = EditWordTypeEnum.W_BREAK) Then
                                'Строка заканчивается на Break [Number]
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 3
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2)
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            End If
                        End If
                        pos += w.Word.Length
                    Next lineId
                Else
                    'был выделен Loop. Ищем Do While
                    For lineId As Integer = curLine - 1 To 0 Step -1
                        'перебираем все строки от предыдущей к началу в поиске Do нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                        pos = 0 'хранит длину обработанной части текущей строки
                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = curWord.wordType Then
                            'Найдено слово из блока
                            Dim wTrim As String = w.Word.Trim
                            If String.Compare(wTrim, "Do", True) = 0 Then
                                'Найден Do
                                If internalCycles <> 0 Then
                                    'это часть внутреннего блока
                                    internalCycles -= 1
                                Else
                                    'это Do нашего блока
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor

                                    'выделяем While
                                    If codeBox.CodeData(lineId).Code.Length > 1 Then
                                        firstCharId += 3
                                        strWord = codeBox.CodeData(lineId).Code(1).Word.Trim
                                        sender.Select(firstCharId, strWord.Length) 'длина
                                        sender.SelectionBackColor = selColor
                                    End If
                                    Exit For
                                End If
                            ElseIf String.Compare(wTrim, "Loop", True) = 0 Then
                                'найден внутренний цикл
                                internalCycles += 1
                            End If
                        ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "next", True) <> 0) OrElse _
                            (w.wordType = EditWordTypeEnum.W_SWITCH AndAlso String.Compare(w.Word.Trim, "Case", True) <> 0) Then
                            internalCyclesForBreak -= 1
                        ElseIf (w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code.Count = 2 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_SWITCH) OrElse _
                            (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "next", True) = 0) Then
                            internalCyclesForBreak += 1
                        ElseIf internalCycles = 0 AndAlso internalCyclesForBreak = 0 Then
                            If w.wordType = EditWordTypeEnum.W_CONTINUE OrElse w.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка начинается с Continue / Break
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_CONTINUE OrElse codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка заканчивается на Continue / Break
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 2
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code.Last
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf (codeBox.CodeData(lineId).Code.Length > 2 AndAlso codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2).wordType = EditWordTypeEnum.W_BREAK) Then
                                'Строка заканчивается на Break [Number]
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 3
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2)
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            End If
                        End If
                        pos += w.Word.Length
                    Next lineId
                End If
                codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
            Case EditWordTypeEnum.W_BLOCK_FOR
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                Dim firstCharId As Integer

                Dim seekForward As Boolean = False
                If curWordId = 0 AndAlso strWord.ToLower = "next" Then
                    'выделяем текущее слово
                    If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                    For wId As Integer = 0 To curWordId - 1
                        pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                    Next wId
                    firstCharId = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor

                    seekForward = False
                Else
                    'выделяем For ... To ... Step
                    pos = 0 'хранит длину обработанной части текущей строки                    
                    If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                    For wId As Integer = 0 To codeBox.CodeData(curLine).Code.Count - 1
                        Dim w As EditWordType = codeBox.CodeData(curLine).Code(wId)
                        If w.wordType = EditWordTypeEnum.W_BLOCK_FOR Then
                            firstCharId = sender.GetFirstCharIndexFromLine(curLine) + pos + w.Word.Length - w.Word.TrimStart.Length  'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        End If
                        pos += w.Word.Length
                    Next wId
                    seekForward = True
                End If

                Dim internalCycles As Integer = 0
                Dim internalCyclesForBreak As Integer = 0
                If seekForward Then
                    'ищем Next
                    For lineId As Integer = curLine + 1 To codeBox.CodeData.Length - 1
                        'перебираем все строки от следующей после текущей в поиске Loop нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                        pos = 0 'хранит длину обработанной части текущей строки
                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = curWord.wordType Then
                            'Найдено слово из блока
                            Dim wTrim As String = w.Word.Trim
                            If String.Compare(wTrim, "Next", True) = 0 Then
                                'Найден Next
                                If internalCycles <> 0 Then
                                    'это часть внутреннего блока
                                    internalCycles -= 1
                                Else
                                    'это часть нашего блока
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                    Exit For
                                End If
                            ElseIf String.Compare(wTrim, "For", True) = 0 Then
                                'найден внутренний цикл
                                internalCycles += 1
                            End If
                        ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "loop", True) <> 0) OrElse _
                            (w.wordType = EditWordTypeEnum.W_SWITCH AndAlso String.Compare(w.Word.Trim, "Case", True) <> 0) Then
                            internalCyclesForBreak += 1
                        ElseIf (w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code.Count = 2 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_SWITCH) OrElse _
                            (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "loop", True) = 0) Then
                            internalCyclesForBreak -= 1
                        ElseIf internalCycles = 0 AndAlso internalCyclesForBreak = 0 Then
                            If w.wordType = EditWordTypeEnum.W_CONTINUE OrElse w.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка начинается с Continue / Break
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_CONTINUE OrElse codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка заканчивается на Continue / Break
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 2
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code.Last
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf (codeBox.CodeData(lineId).Code.Length > 2 AndAlso codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2).wordType = EditWordTypeEnum.W_BREAK) Then
                                'Строка заканчивается на Break [Number]
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 3
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2)
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            End If
                        End If
                        pos += w.Word.Length
                    Next lineId
                Else
                    'был выделен Next. Ищем For ... To ... Step
                    For lineId As Integer = curLine - 1 To 0 Step -1
                        'перебираем все строки от предыдущей к началу в поиске For нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = curWord.wordType Then
                            'Найдено слово из блока
                            Dim wTrim As String = w.Word.Trim
                            If String.Compare(wTrim, "For", True) = 0 Then
                                'Найден For
                                If internalCycles <> 0 Then
                                    'это часть внутреннего блока
                                    internalCycles -= 1
                                Else
                                    'это For нашего блока
                                    'выделяем For ... To ... Step
                                    pos = 0 'хранит длину обработанной части текущей строки                    
                                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                                    For wId As Integer = 0 To codeBox.CodeData(lineId).Code.Count - 1
                                        w = codeBox.CodeData(lineId).Code(wId)
                                        If w.wordType = EditWordTypeEnum.W_BLOCK_FOR Then
                                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length  'получаем первый символ выделения
                                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                            sender.SelectionBackColor = selColor
                                        End If
                                        pos += w.Word.Length
                                    Next wId

                                    Exit For
                                End If
                            ElseIf String.Compare(wTrim, "Next", True) = 0 Then
                                'найден внутренний цикл
                                internalCycles += 1
                            End If
                        ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "loop", True) <> 0) OrElse _
                            (w.wordType = EditWordTypeEnum.W_SWITCH AndAlso String.Compare(w.Word.Trim, "Case", True) <> 0) Then
                            internalCyclesForBreak += 1
                        ElseIf (w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code.Count = 2 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_SWITCH) OrElse _
                            (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "loop", True) = 0) Then
                            internalCyclesForBreak -= 1
                        ElseIf internalCycles = 0 AndAlso internalCyclesForBreak = 0 Then
                            pos = 0 'хранит длину обработанной части текущей строки
                            If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                            If w.wordType = EditWordTypeEnum.W_CONTINUE OrElse w.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка начинается с Continue / Break
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_CONTINUE OrElse codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка заканчивается на Continue / Break
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 2
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code.Last
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf (codeBox.CodeData(lineId).Code.Length > 2 AndAlso codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2).wordType = EditWordTypeEnum.W_BREAK) Then
                                'Строка заканчивается на Break [Number]
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 3
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2)
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            End If
                        End If
                    Next lineId
                End If
                codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
            Case EditWordTypeEnum.W_BLOCK_EVENT, EditWordTypeEnum.W_BLOCK_FUNCTION, EditWordTypeEnum.W_HTML, EditWordTypeEnum.W_WRAP
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                'выделяем текущее слово
                Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                For wId As Integer = 0 To curWordId - 1
                    pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                Next wId
                Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                sender.Select(firstCharId, strWord.Length) 'длина
                sender.SelectionBackColor = selColor
                If curWord.wordType = EditWordTypeEnum.W_HTML Then
                    If curWordId = 0 AndAlso String.Compare(strWord, "HTML", True) = 0 AndAlso codeBox.CodeData(curLine).Code.Count > 1 AndAlso codeBox.CodeData(curLine).Code(1).wordType = EditWordTypeEnum.W_HTML Then
                        'выделено HTML
                        'выделяем Append
                        firstCharId += 5
                        strWord = codeBox.CodeData(curLine).Code(1).Word.Trim
                        sender.Select(firstCharId, strWord.Length) 'длина
                        sender.SelectionBackColor = selColor
                    ElseIf curWordId = 1 AndAlso codeBox.CodeData(curLine).Code(1).wordType = EditWordTypeEnum.W_HTML AndAlso String.Compare(strWord, "Append", True) = 0 Then
                        'выделено Append
                        'выделяем HTML
                        firstCharId -= 5
                        strWord = codeBox.CodeData(curLine).Code(0).Word.Trim
                        sender.Select(firstCharId, strWord.Length) 'длина
                        sender.SelectionBackColor = selColor
                    End If
                End If

                Dim seekForward As Boolean = False
                If curWordId = 1 AndAlso codeBox.CodeData(curLine).Code(0).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    'End Event
                    'выделяем End
                    firstCharId -= 4
                    strWord = codeBox.CodeData(curLine).Code(0).Word.Trim
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = False
                Else
                    'Event ...
                    seekForward = True
                End If

                Dim internalCycles As Integer = 0
                If seekForward Then
                    'ищем End Event
                    For lineId As Integer = curLine + 1 To codeBox.CodeData.Length - 1
                        'перебираем все строки от следующей после текущей в поиске End Event нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) OrElse codeBox.CodeData(lineId).Code.Length < 2 Then Continue For 'должно быть два слова - End Event
                        pos = 0 'хранит длину обработанной части текущей строки
                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code(1).wordType = curWord.wordType Then
                            'Найдено слово из блока
                            If internalCycles <> 0 Then
                                'это часть внутреннего блока
                                internalCycles -= 1
                            Else
                                'это End нашего блока
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor

                                'выделяем Event
                                firstCharId += 4
                                strWord = codeBox.CodeData(lineId).Code(1).Word.Trim
                                sender.Select(firstCharId, strWord.Length) 'длина
                                sender.SelectionBackColor = selColor
                                Exit For
                            End If
                        ElseIf w.wordType = curWord.wordType Then
                            'найден внутренний цикл
                            internalCycles += 1
                        End If
                        pos += w.Word.Length
                    Next lineId
                Else
                    'Ищем Event ...
                    For lineId As Integer = curLine - 1 To 0 Step -1
                        'перебираем все строки от предыдущей к началу в поиске Event ... нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                        pos = 0 'хранит длину обработанной части текущей строки
                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = curWord.wordType Then
                            'Найдено слово Event из блока
                            If internalCycles <> 0 Then
                                'это часть внутреннего блока
                                internalCycles -= 1
                            Else
                                'это Event нашего блока
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                                If w.wordType = EditWordTypeEnum.W_HTML AndAlso codeBox.CodeData(lineId).Code.Count > 1 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_HTML Then
                                    'выделяем Append
                                    firstCharId += 5
                                    w = codeBox.CodeData(lineId).Code(1)
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                End If
                                Exit For
                            End If
                        ElseIf w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code.Count = 2 AndAlso codeBox.CodeData(lineId).Code(1).wordType = curWord.wordType Then
                            'найден внутренний цикл
                            internalCycles += 1
                        End If
                        pos += w.Word.Length
                    Next lineId
                End If
                codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
            Case EditWordTypeEnum.W_BLOCK_IF
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                'выделяем текущее слово
                Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                For wId As Integer = 0 To curWordId - 1
                    pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                Next wId
                Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                sender.Select(firstCharId, strWord.Length) 'длина
                sender.SelectionBackColor = selColor

                Dim seekForward As Boolean = False
                Dim multiLine As Boolean = False
                If curWordId = 0 AndAlso String.Compare(strWord, "Then", True) <> 0 Then
                    'выбрано If в начале блока If Then или ElseIf или Else
                    'выделяем then
                    pos = 0 'хранит длину обработанной части текущей строки
                    If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                    Dim blnLineChanged As Boolean = False

                    For wId As Integer = 0 To codeBox.CodeData(curLine).Code.Count - 1
                        If wId > codeBox.CodeData(curLine).Code.Count - 1 Then Exit For
                        Dim w As EditWordType = codeBox.CodeData(curLine).Code(wId)
                        If w.wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso (wId > 0 OrElse blnLineChanged) Then
                            firstCharId = sender.GetFirstCharIndexFromLine(curLine) + pos + w.Word.Length - w.Word.TrimStart.Length  'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                            If wId = codeBox.CodeData(curLine).Code.Count - 1 Then multiLine = True
                            Exit For
                        End If
                        pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                        If wId = codeBox.CodeData(curLine).Code.Count - 1 AndAlso w.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION AndAlso curLine < codeBox.CodeData.Count - 1 Then
                            'стока заканчивается на _ . Переходим ниже
                            wId = -1
                            curLine += 1
                            pos = 0 'хранит длину обработанной части текущей строки
                            If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                            blnLineChanged = True
                        End If
                    Next wId
                    If strWord.StartsWith("Else", StringComparison.CurrentCultureIgnoreCase) Then multiLine = True
                    seekForward = True
                ElseIf strWord.ToLower = "then" Then
                    'выбрано Then
                    'выделяем If или ElseIf
                    pos = 0 'хранит длину обработанной части текущей строки
                    Dim startLine As Integer = curLine
                    Do
                        Dim w As String = codeBox.CodeData(curLine).Code(0).Word.Trim
                        If String.Compare(w, "If", True) = 0 OrElse String.Compare(w, "ElseIf", True) = 0 Then Exit Do
                        If curLine = 0 OrElse IsNothing(codeBox.CodeData(curLine - 1).Code) OrElse codeBox.CodeData(curLine - 1).Code.Last.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then GoTo finish
                        curLine -= 1
                    Loop
                    If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                    firstCharId = sender.GetFirstCharIndexFromLine(curLine) + pos + codeBox.CodeData(curLine).Code(0).Word.Length - codeBox.CodeData(curLine).Code(0).Word.TrimStart.Length  'получаем первый символ выделения
                    sender.Select(firstCharId, codeBox.CodeData(curLine).Code(0).Word.Trim.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = True
                    If curWordId = codeBox.CodeData(startLine).Code.Count - 1 Then multiLine = True
                ElseIf curWordId = 1 AndAlso codeBox.CodeData(curLine).Code(0).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    'выбрано If в конце блока End If
                    'выделяем End
                    pos = 0 'хранит длину обработанной части текущей строки
                    If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                    firstCharId = sender.GetFirstCharIndexFromLine(curLine) + pos + codeBox.CodeData(curLine).Code(0).Word.Length - codeBox.CodeData(curLine).Code(0).Word.TrimStart.Length  'получаем первый символ выделения
                    sender.Select(firstCharId, codeBox.CodeData(curLine).Code(0).Word.Trim.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = False
                    multiLine = True
                End If

                If multiLine Then ' (codeBox.CodeData(curLine).Code(0).Word.Trim.ToLower = "if" AndAlso codeBox.CodeData(curLine).Code.Last.Word.Trim.ToLower <> "then") = False Then
                    'многострочный блок
                    Dim internalCycles As Integer = 0
                    If seekForward Then
                        'ищем либо от If Then, либо от ElseIf Then, лио от Else
                        'ищем ElseIf Then, End If
                        For lineId As Integer = curLine + 1 To codeBox.CodeData.Length - 1
                            'перебираем все строки от следующей после текущей в поиске ElseIf Then, Else, End If нашего цикла
                            If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                            pos = 0 'хранит длину обработанной части текущей строки
                            If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                            'интересуют только первое слово в строке
                            Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                            If w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_BLOCK_IF Then
                                'Найдено End If
                                If internalCycles <> 0 Then
                                    'это часть внутреннего блока
                                    internalCycles -= 1
                                Else
                                    'это End нашего блока
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor

                                    'выделяем If
                                    firstCharId += 4
                                    strWord = codeBox.CodeData(lineId).Code(1).Word.Trim
                                    sender.Select(firstCharId, strWord.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                    Exit For
                                End If
                            ElseIf w.wordType = EditWordTypeEnum.W_BLOCK_IF Then
                                Dim strTrim As String = w.Word.Trim
                                If String.Compare(strTrim, "If", True) = 0 AndAlso String.Compare(codeBox.CodeData(lineId).Code.Last.Word.Trim, "Then", True) = 0 Then
                                    'найден внутренний цикл
                                    internalCycles += 1
                                ElseIf String.Compare(strTrim, "ElseIf", True) = 0 AndAlso internalCycles = 0 Then
                                    'найден ElseIf нашего блока
                                    'выделяем ElseIf
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                    'выделяем Then
                                    pos = 0 'хранит длину обработанной части текущей строки
                                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                                    Dim blnLineChanged As Boolean = False
                                    For wId As Integer = 0 To codeBox.CodeData(lineId).Code.Count - 1
                                        w = codeBox.CodeData(lineId).Code(wId)
                                        If w.wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso (wId > 0 OrElse blnLineChanged) Then
                                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length  'получаем первый символ выделения
                                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                            sender.SelectionBackColor = selColor
                                            Exit For
                                        End If
                                        pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                        If wId = codeBox.CodeData(lineId).Code.Count - 1 AndAlso w.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION AndAlso lineId < codeBox.CodeData.Count - 1 Then
                                            'стока заканчивается на _ . Переходим ниже
                                            wId = -1
                                            lineId += 1
                                            pos = 0 'хранит длину обработанной части текущей строки
                                            If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                                            blnLineChanged = True
                                        End If
                                    Next wId
                                ElseIf String.Compare(strTrim, "Else", True) = 0 AndAlso internalCycles = 0 Then
                                    'найден Else нашего блока. Выделяем
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                End If
                            End If
                            pos += w.Word.Length
                        Next lineId
                    End If

                    'ищем от End if/Else/Else If
                    If seekForward = False OrElse codeBox.CodeData(curLine).Code(0).Word.Trim.StartsWith("Else", StringComparison.CurrentCultureIgnoreCase) Then

                        For lineId As Integer = curLine - 1 To 0 Step -1
                            'перебираем все строки от предыдущей и до конца в поиске ElseIf Then, Else, If Then нашего блока
                            If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                            pos = 0 'хранит длину обработанной части текущей строки
                            If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                            'интересуют только первое слово в строке
                            Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)

                            If w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_BLOCK_IF Then
                                'Найдено End If
                                internalCycles += 1
                            ElseIf w.wordType = EditWordTypeEnum.W_BLOCK_IF Then
                                Dim strTrim As String = w.Word.Trim
                                If String.Compare(strTrim, "If", True) = 0 AndAlso (String.Compare(codeBox.CodeData(lineId).Code.Last.Word.Trim, "Then", True) = 0 OrElse _
                                                                                    codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION) Then
                                    If codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                                        Dim tmpLine As Integer = lineId
                                        Do
                                            If tmpLine = codeBox.CodeData.Count - 1 OrElse IsNothing(codeBox.CodeData(tmpLine + 1).Code) Then Continue For
                                            tmpLine += 1
                                            If codeBox.CodeData(tmpLine).Code.Last.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then Continue Do
                                            If String.Compare(codeBox.CodeData(tmpLine).Code.Last.Word.Trim, "Then", True) = 0 Then
                                                Exit Do
                                            Else
                                                Continue For
                                            End If
                                        Loop
                                    End If
                                    'найден If Then
                                    If internalCycles <> 0 Then
                                        'это часть внутреннего блока
                                        internalCycles -= 1
                                    Else
                                        'это If Then нашего блока. Выделяем If
                                        firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                        sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                        sender.SelectionBackColor = selColor

                                        'выделяем Then
                                        pos = 0 'хранит длину обработанной части текущей строки
                                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                                        Dim blnLineChanged As Boolean = False
                                        For wId As Integer = 0 To codeBox.CodeData(lineId).Code.Count - 1
                                            w = codeBox.CodeData(lineId).Code(wId)
                                            If w.wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso (wId > 0 OrElse blnLineChanged) Then
                                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length  'получаем первый символ выделения
                                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                                sender.SelectionBackColor = selColor
                                                Exit For
                                            End If
                                            pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                            If wId = codeBox.CodeData(lineId).Code.Count - 1 AndAlso w.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION AndAlso lineId < codeBox.CodeData.Count - 1 Then
                                                'стока заканчивается на _ . Переходим ниже
                                                wId = -1
                                                lineId += 1
                                                pos = 0 'хранит длину обработанной части текущей строки
                                                If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                                                blnLineChanged = True
                                            End If
                                        Next wId
                                        Exit For
                                    End If
                                ElseIf String.Compare(strTrim, "ElseIf", True) = 0 AndAlso internalCycles = 0 Then
                                    'найден ElseIf нашего блока
                                    'выделяем ElseIf
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                    'выделяем Then
                                    pos = 0 'хранит длину обработанной части текущей строки
                                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                                    For wId As Integer = 0 To codeBox.CodeData(lineId).Code.Count - 1
                                        w = codeBox.CodeData(lineId).Code(wId)
                                        If w.wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso wId > 0 Then
                                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length  'получаем первый символ выделения
                                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                            sender.SelectionBackColor = selColor
                                            Exit For
                                        End If
                                        pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                    Next wId
                                ElseIf String.Compare(strTrim, "Else", True) = 0 AndAlso internalCycles = 0 Then
                                    'найден Else нашего блока. Выделяем
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                End If
                            End If
                            pos += w.Word.Length
                        Next lineId

                    End If

                End If
                codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
            Case EditWordTypeEnum.W_BLOCK_NEWCLASS
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                'выделяем текущее слово
                Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                For wId As Integer = 0 To curWordId - 1
                    pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                Next wId
                Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                sender.Select(firstCharId, strWord.Length) 'длина
                sender.SelectionBackColor = selColor

                Dim seekForward As Boolean = False
                If curWordId = 1 AndAlso codeBox.CodeData(curLine).Code(0).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    'End Class
                    'выделяем End
                    firstCharId -= 4
                    strWord = codeBox.CodeData(curLine).Code(0).Word.Trim
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = False
                ElseIf curWordId = 0 Then
                    'выделено New/Prop/Func
                    seekForward = True
                Else
                    'выделено Class
                    'выделям New
                    pos = 0
                    If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                    firstCharId = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                    strWord = codeBox.CodeData(curLine).Code(0).Word.Trim
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = False
                End If

                Dim internalCycles As Integer = 0
                If seekForward Then
                    'ищем End Class
                    For lineId As Integer = curLine + 1 To codeBox.CodeData.Length - 1
                        'перебираем все строки от следующей после текущей в поиске End Class нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                        pos = 0 'хранит длину обработанной части текущей строки
                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code(1).wordType = curWord.wordType Then
                            'Найдено End Class
                            If internalCycles <> 0 Then
                                'это часть внутреннего блока
                                internalCycles -= 1
                            Else
                                'это End нашего блока
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor

                                'выделяем Class
                                firstCharId += 4
                                strWord = codeBox.CodeData(lineId).Code(1).Word.Trim
                                sender.Select(firstCharId, strWord.Length) 'длина
                                sender.SelectionBackColor = selColor
                                Exit For
                            End If
                        ElseIf w.wordType = curWord.wordType Then
                            If String.Compare(w.Word.Trim, "New", True) = 0 Then
                                'найден внутренний цикл
                                internalCycles += 1
                            ElseIf internalCycles = 0 Then
                                'Выделяем Func / Prop
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            End If
                        End If
                        pos += w.Word.Length
                    Next lineId
                End If

                If seekForward = False OrElse String.Compare(curWord.Word.Trim, "Class", True) <> 0 Then
                    'Ищем New от EndClass или от внутренних Prop / Func
                    For lineId As Integer = curLine - 1 To 0 Step -1
                        'перебираем все строки от предыдущей к началу в поиске New Class нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                        pos = 0 'хранит длину обработанной части текущей строки
                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = curWord.wordType Then
                            'Найдено слово из блока - New / Prop / Func
                            If String.Compare(w.Word.Trim, "New", True) = 0 Then
                                'Найдено слово из блока - New
                                If internalCycles <> 0 Then
                                    'это часть внутреннего блока
                                    internalCycles -= 1
                                Else
                                    'это New нашего блока
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                    'выделяем Class
                                    firstCharId += 4
                                    w = codeBox.CodeData(lineId).Code(1)
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                    Exit For
                                End If
                            ElseIf internalCycles = 0 Then
                                'Найдено слово из блока - Func / Prop
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            End If
                        ElseIf w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code.Count = 2 AndAlso codeBox.CodeData(lineId).Code(1).wordType = curWord.wordType Then
                            'найден внутренний цикл
                            internalCycles += 1
                        End If
                        pos += w.Word.Length
                    Next lineId
                End If
                codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
            Case EditWordTypeEnum.W_SWITCH
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                'выделяем текущее слово
                Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                For wId As Integer = 0 To curWordId - 1
                    pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                Next wId
                Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                sender.Select(firstCharId, strWord.Length) 'длина
                sender.SelectionBackColor = selColor

                Dim seekForward As Boolean = False
                If curWordId = 1 AndAlso codeBox.CodeData(curLine).Code(0).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    'End Select/Switch
                    'выделяем End
                    firstCharId -= 4
                    strWord = codeBox.CodeData(curLine).Code(0).Word.Trim
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = False
                ElseIf curWordId = 0 AndAlso String.Compare(strWord, "Select", True) = 0 Then
                    'выделено Select
                    'выделяем Case
                    firstCharId += 7
                    strWord = codeBox.CodeData(curLine).Code(1).Word.Trim
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = True
                ElseIf curWordId = 0 AndAlso String.Compare(strWord, "Case", True) = 0 AndAlso codeBox.CodeData(curLine).Code.Length = 2 AndAlso String.Compare(codeBox.CodeData(curLine).Code(1).Word.Trim, "Else", True) = 0 Then
                    'Case Else
                    'Выделяем Else
                    firstCharId += 5
                    strWord = codeBox.CodeData(curLine).Code(1).Word.Trim
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = True
                ElseIf curWordId = 0 Then
                    'выделено Case / Switch
                    seekForward = True
                Else
                    'выделено Case первой строки
                    'выделям Select
                    pos = 0
                    If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                    firstCharId = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                    strWord = codeBox.CodeData(curLine).Code(0).Word.Trim
                    sender.Select(firstCharId, strWord.Length) 'длина
                    sender.SelectionBackColor = selColor
                    seekForward = False
                End If

                Dim internalCycles As Integer = 0
                Dim internalCyclesForBreak = 0
                If seekForward Then
                    'ищем End Select/Switch
                    For lineId As Integer = curLine + 1 To codeBox.CodeData.Length - 1
                        'перебираем все строки от следующей после текущей в поиске End Select/Switch нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                        pos = 0 'хранит длину обработанной части текущей строки
                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code(1).wordType = curWord.wordType Then
                            'Найдено End Select / Switch
                            If internalCycles <> 0 Then
                                'это часть внутреннего блока
                                internalCycles -= 1
                            Else
                                'это End нашего блока
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor

                                'выделяем Select/Switch
                                firstCharId += 4
                                strWord = codeBox.CodeData(lineId).Code(1).Word.Trim
                                sender.Select(firstCharId, strWord.Length) 'длина
                                sender.SelectionBackColor = selColor
                                Exit For
                            End If
                        ElseIf w.wordType = curWord.wordType Then
                            If String.Compare(w.Word.Trim, "Select", True) = 0 OrElse String.Compare(w.Word.Trim, "Switch", True) = 0 Then
                                'найден внутренний цикл
                                internalCycles += 1
                            ElseIf internalCycles = 0 Then
                                'Выделяем Case x
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                                If codeBox.CodeData(lineId).Code.Length = 2 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso String.Compare(codeBox.CodeData(lineId).Code(1).Word.Trim, "Else", True) = 0 Then
                                    'Case Else
                                    firstCharId += 5
                                    strWord = codeBox.CodeData(lineId).Code(1).Word.Trim
                                    sender.Select(firstCharId, strWord.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                End If
                            End If
                        ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "Loop", True) <> 0) OrElse _
                               (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "next", True) <> 0) Then
                            internalCyclesForBreak += 1
                        ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "next", True) = 0) OrElse _
                            (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "loop", True) = 0) Then
                            internalCyclesForBreak -= 1
                        ElseIf internalCycles = 0 AndAlso internalCyclesForBreak = 0 Then
                            If w.wordType = EditWordTypeEnum.W_CONTINUE OrElse w.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка начинается с Continue / Break
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_CONTINUE OrElse codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка заканчивается на Continue / Break
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 2
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code.Last
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf (codeBox.CodeData(lineId).Code.Length > 2 AndAlso codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2).wordType = EditWordTypeEnum.W_BREAK) Then
                                'Строка заканчивается на Break [Number]
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 3
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2)
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            End If
                        End If
                        pos += w.Word.Length
                    Next lineId
                End If

                If seekForward = False OrElse (String.Compare(curWord.Word.Trim, "Case", True) = 0 AndAlso curWordId = 0) Then
                    'Ищем Select / Switch от End Select / Switch или от внутренних Case
                    For lineId As Integer = curLine - 1 To 0 Step -1
                        'перебираем все строки от предыдущей к началу в поиске Select Case / Switch нашего цикла
                        If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                        pos = 0 'хранит длину обработанной части текущей строки
                        If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                        'интересуют только первые слова в строке
                        Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                        If w.wordType = curWord.wordType Then
                            'Найдено слово из блока - New / Prop / Func
                            If String.Compare(w.Word.Trim, "Select", True) = 0 OrElse String.Compare(w.Word.Trim, "Switch", True) = 0 Then
                                'Найдено слово из блока - Select / Switch
                                If internalCycles <> 0 Then
                                    'это часть внутреннего блока
                                    internalCycles -= 1
                                Else
                                    'это Select / Switch нашего блока
                                    firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                    If String.Compare(w.Word.Trim, "Select", True) = 0 Then
                                        'выделяем Case
                                        firstCharId += 7
                                        w = codeBox.CodeData(lineId).Code(1)
                                        sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                        sender.SelectionBackColor = selColor
                                    End If
                                    Exit For
                                End If
                            ElseIf internalCycles = 0 Then
                                'Найдено слово из блока - Case
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                                If codeBox.CodeData(lineId).Code.Length = 2 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_BLOCK_IF AndAlso String.Compare(codeBox.CodeData(lineId).Code(1).Word.Trim, "Else", True) = 0 Then
                                    'Case Else
                                    firstCharId += 5
                                    strWord = codeBox.CodeData(lineId).Code(1).Word.Trim
                                    sender.Select(firstCharId, strWord.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                End If
                            End If
                        ElseIf w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code.Count = 2 AndAlso codeBox.CodeData(lineId).Code(1).wordType = curWord.wordType Then
                            'найден внутренний цикл
                            internalCycles += 1
                        ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "Loop", True) <> 0) OrElse _
                               (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "next", True) <> 0) Then
                            internalCyclesForBreak -= 1
                        ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "next", True) = 0) OrElse _
                               (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "loop", True) = 0) Then
                            internalCyclesForBreak += 1
                        ElseIf internalCycles = 0 AndAlso internalCyclesForBreak = 0 Then
                            If w.wordType = EditWordTypeEnum.W_CONTINUE OrElse w.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка начинается с Continue / Break
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_CONTINUE OrElse codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_BREAK Then
                                'Строка заканчивается на Continue / Break
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 2
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code.Last
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            ElseIf (codeBox.CodeData(lineId).Code.Length > 2 AndAlso codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2).wordType = EditWordTypeEnum.W_BREAK) Then
                                'Строка заканчивается на Break [Number]
                                pos += w.Word.Length
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 3
                                    pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                                Next wId
                                w = codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2)
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            End If
                        End If
                        pos += w.Word.Length
                    Next lineId
                End If
                codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
            Case EditWordTypeEnum.W_BREAK, EditWordTypeEnum.W_CONTINUE
                'внутри циклов Do, For, Select / Switch
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                'выделяем текущее слово
                Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                For wId As Integer = 0 To curWordId - 1
                    pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                Next wId
                Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                sender.Select(firstCharId, strWord.Length) 'длина
                sender.SelectionBackColor = selColor


                Dim internalCycles As Integer = 0
                'Закрашиваем все в прямом порядке
                For lineId As Integer = curLine + 1 To codeBox.CodeData.Length - 1
                    'перебираем все строки от от следующей за текущей в поиске Select Case / Switch / Do / For нашего цикла
                    If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                    pos = 0 'хранит длину обработанной части текущей строки
                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                    'интересуют только первые слова в строке
                    Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                    If (w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code.Count > 1 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_SWITCH) OrElse _
                        (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "Loop") = 0) OrElse _
                        (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "Next") = 0) Then
                        If internalCycles <> 0 Then
                            internalCycles -= 1
                        Else
                            'завершение блока, в который входит Continue / Break
                            'закрашиваем End, Next, Loop
                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                            If codeBox.CodeData(lineId).Code.Count > 1 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_SWITCH Then
                                'закрашиваем Select / Switch после End
                                firstCharId += 4
                                w = codeBox.CodeData(lineId).Code(1)
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                            End If
                            Exit For
                        End If
                    ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "Do", True) = 0) OrElse _
                        (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "For", True) = 0) OrElse _
                         (w.wordType = EditWordTypeEnum.W_SWITCH AndAlso String.Compare(w.Word.Trim, "Case", True) <> 0) Then
                        'начало нового блока Do, For, Select / Switch
                        internalCycles += 1
                    ElseIf w.wordType = EditWordTypeEnum.W_SWITCH AndAlso String.Compare(w.Word.Trim, "Case") = 0 Then
                        'Case X, Y...
                        firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                        sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                        sender.SelectionBackColor = selColor
                    ElseIf internalCycles = 0 Then
                        If w.wordType = EditWordTypeEnum.W_CONTINUE OrElse w.wordType = EditWordTypeEnum.W_BREAK Then
                            'Строка начинается с Continue / Break
                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        ElseIf codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_CONTINUE OrElse codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_BREAK Then
                            'Строка заканчивается на Continue / Break
                            pos += w.Word.Length
                            For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 2
                                pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                            Next wId
                            w = codeBox.CodeData(lineId).Code.Last
                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        ElseIf (codeBox.CodeData(lineId).Code.Length > 2 AndAlso codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2).wordType = EditWordTypeEnum.W_BREAK) Then
                            'Строка заканчивается на Break [Number]
                            pos += w.Word.Length
                            For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 3
                                pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                            Next wId
                            w = codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2)
                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        End If
                    End If
                Next lineId

                'А теперь закрашиваем все в обратном направлении
                For lineId As Integer = curLine - 1 To 0 Step -1
                    'перебираем все строки от предыдущей к началу в поиске Select Case / Switch нашего цикла
                    If IsNothing(codeBox.CodeData(lineId).Code) Then Continue For
                    pos = 0 'хранит длину обработанной части текущей строки
                    If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                    'интересуют только первые слова в строке
                    Dim w As EditWordType = codeBox.CodeData(lineId).Code(0)
                    If (w.wordType = EditWordTypeEnum.W_CYCLE_END AndAlso codeBox.CodeData(lineId).Code.Count > 1 AndAlso codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_SWITCH) OrElse _
                        (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "Loop") = 0) OrElse _
                        (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "Next") = 0) Then
                        internalCycles += 1
                    ElseIf (w.wordType = EditWordTypeEnum.W_BLOCK_DOWHILE AndAlso String.Compare(w.Word.Trim, "Do", True) = 0) OrElse _
                        (w.wordType = EditWordTypeEnum.W_BLOCK_FOR AndAlso String.Compare(w.Word.Trim, "For", True) = 0) OrElse _
                         (w.wordType = EditWordTypeEnum.W_SWITCH AndAlso String.Compare(w.Word.Trim, "Case", True) <> 0) Then
                        'начало нового блока Do, For, Select / Switch
                        If internalCycles <> 0 Then
                            internalCycles -= 1
                        Else
                            'начало блока, в который входит Continue / Break
                            'закрашиваем Do While, For To Step, Select Case / Switch
                            If w.wordType = EditWordTypeEnum.W_BLOCK_FOR Then
                                'закрашиваем For
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                                'Закрашиваем To Step
                                pos += 4 '"For "
                                For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 1
                                    w = codeBox.CodeData(lineId).Code(wId)
                                    If w.wordType = EditWordTypeEnum.W_BLOCK_FOR Then
                                        firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                        sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                        sender.SelectionBackColor = selColor
                                    End If
                                    pos += w.Word.Length
                                Next wId
                            Else
                                'закрашиваем Do / Select / Switch
                                firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                                sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                sender.SelectionBackColor = selColor
                                If codeBox.CodeData(lineId).Code.Count > 1 AndAlso (codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_SWITCH OrElse _
                                                                            codeBox.CodeData(lineId).Code(1).wordType = EditWordTypeEnum.W_BLOCK_DOWHILE) Then
                                    firstCharId += w.Word.Trim.Length + 1
                                    'закрашиваем While / Case
                                    w = codeBox.CodeData(lineId).Code(1)
                                    sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                                    sender.SelectionBackColor = selColor
                                End If
                            End If
                            Exit For
                        End If
                    ElseIf w.wordType = EditWordTypeEnum.W_SWITCH AndAlso String.Compare(w.Word.Trim, "Case") = 0 Then
                        'Case X, Y...
                        firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                        sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                        sender.SelectionBackColor = selColor
                    ElseIf internalCycles = 0 Then
                        If w.wordType = EditWordTypeEnum.W_CONTINUE OrElse w.wordType = EditWordTypeEnum.W_BREAK Then
                            'Строка начинается с Continue / Break
                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        ElseIf codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_CONTINUE OrElse codeBox.CodeData(lineId).Code.Last.wordType = EditWordTypeEnum.W_BREAK Then
                            'Строка заканчивается на Continue / Break
                            pos += w.Word.Length
                            For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 2
                                pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                            Next wId
                            w = codeBox.CodeData(lineId).Code.Last
                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        ElseIf (codeBox.CodeData(lineId).Code.Length > 2 AndAlso codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2).wordType = EditWordTypeEnum.W_BREAK) Then
                            'Строка заканчивается на Break [Number]
                            pos += w.Word.Length
                            For wId As Integer = 1 To codeBox.CodeData(lineId).Code.Count - 3
                                pos += codeBox.CodeData(lineId).Code(wId).Word.Length
                            Next wId
                            w = codeBox.CodeData(lineId).Code(codeBox.CodeData(lineId).Code.Count - 2)
                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                        End If
                    End If
                Next lineId
                codeBox.ShowElementInfo(curWord.classId, strWord, curWord.wordType)
            Case EditWordTypeEnum.W_OVAL_BRACKET_OPEN, EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                'выделяем текущее слово
                Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                Dim wId As Integer
                For wId = 0 To curWordId - 1
                    pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                Next wId
                Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                sender.Select(firstCharId, strWord.Length) 'длина
                sender.SelectionBackColor = selColor

                Dim closeType As EditWordTypeEnum
                If curWord.wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN Then
                    closeType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                Else
                    closeType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                End If

                Dim balance As Integer = 1
                Dim lineId As Integer = curLine
                wId = curWordId + 1
                pos += 1 'открывающая скобка
                Do
                    If wId > codeBox.CodeData(lineId).Code.Count - 1 Then GoTo finish
                    Dim w As EditWordType = codeBox.CodeData(lineId).Code(wId)
                    If wId = codeBox.CodeData(lineId).Code.Count - 1 Then
                        If w.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                            lineId += 1
                            wId = 0
                            If lineId > codeBox.CodeData.Count - 1 OrElse IsNothing(codeBox.CodeData(lineId).Code) Then GoTo finish
                            pos = 0
                            If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                            Continue Do
                        End If
                    End If
                    If w.wordType = curWord.wordType Then
                        balance += 1
                    ElseIf w.wordType = closeType Then
                        balance -= 1
                        If balance = 0 Then
                            'закрашиваем скобку
                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                            Exit Do
                        End If
                    End If
                    pos += w.Word.Length
                    wId += 1
                Loop

            Case EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                codeBox.CanRaiseCodeEvents = False 'запрет TextChanged
                'получаем верхнюю видимую линию
                upperLine = sender.GetCharIndexFromPosition(New Point(5, 5))
                upperLine = sender.GetLineFromCharIndex(upperLine)
                SendMessage(sender.Handle, WM_SetRedraw, 0, 0) 'запрет обновления
                'сохраняем выделение
                selStart = sender.SelectionStart
                selLength = sender.SelectionLength

                'выделяем текущее слово
                Dim pos As Integer = 0 'хранит длину обработанной части текущей строки
                If String.IsNullOrEmpty(codeBox.CodeData(curLine).StartingSpaces) = False Then pos = codeBox.CodeData(curLine).StartingSpaces.Length
                Dim wId As Integer
                For wId = 0 To curWordId - 1
                    pos += codeBox.CodeData(curLine).Code(wId).Word.Length
                Next wId
                Dim firstCharId As Integer = sender.GetFirstCharIndexFromLine(curLine) + pos 'получаем первый символ выделения
                sender.Select(firstCharId, strWord.Length) 'длина
                sender.SelectionBackColor = selColor

                Dim openType As EditWordTypeEnum
                If curWord.wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE Then
                    openType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                Else
                    openType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                End If

                Dim balance As Integer = -1
                Dim lineId As Integer = curLine
                wId = curWordId - 1
                Do
                    If wId < 0 Then
                        If lineId > 0 AndAlso IsNothing(codeBox.CodeData(lineId - 1).Code) = False AndAlso codeBox.CodeData(lineId - 1).Code.Count > 0 AndAlso _
                            codeBox.CodeData(lineId - 1).Code.Last.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                            'переход на строчку вверх
                            lineId -= 1
                            wId = codeBox.CodeData(lineId).Code.Count - 1
                            Continue Do
                        Else
                            GoTo finish
                        End If
                    End If
                    Dim w As EditWordType = codeBox.CodeData(lineId).Code(wId)
                    If w.wordType = curWord.wordType Then
                        balance -= 1
                    ElseIf w.wordType = openType Then
                        balance += 1
                        If balance = 0 Then
                            'закрашиваем скобку
                            pos = 0
                            If String.IsNullOrEmpty(codeBox.CodeData(lineId).StartingSpaces) = False Then pos = codeBox.CodeData(lineId).StartingSpaces.Length
                            For i As Integer = 0 To wId - 1
                                pos += codeBox.CodeData(lineId).Code(i).Word.Length
                            Next i
                            firstCharId = sender.GetFirstCharIndexFromLine(lineId) + pos + w.Word.Length - w.Word.TrimStart.Length 'получаем первый символ выделения
                            sender.Select(firstCharId, w.Word.Trim.Length) 'длина
                            sender.SelectionBackColor = selColor
                            Exit Do
                        End If
                    End If
                    wId -= 1
                Loop
            Case Else
                Return
        End Select
finish:
        sender.Select(selStart, selLength) 'восстанавливаем выделение
        'возвращаем на место скролбар
        Dim charPos2 As Integer = sender.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})
        Dim lId2 As Integer = sender.GetLineFromCharIndex(charPos2)
        SendMessage(sender.Handle, EM_LINESCROLL, 0, upperLine - lId2)
        'разрешаем обновление
        SendMessage(sender.Handle, WM_SetRedraw, 1, 0)
        sender.Refresh()
        codeBox.CanRaiseCodeEvents = True 'разрешаем TextChanged
        codeBox.wasTextSelected = True
    End Sub

    ''' <summary>полулокальна переменная для комбинаций Ctrl+Backspace и Ctrl+Delete. содержит SelectionStart до Ctrl+Backspace или длину всего текста до Ctrl+Del</summary>
    Private shortcutUsedData As Integer = -1
    ''' <summary>копия всего текста до Ctrl+Backspace/Ctrl+Delete</summary>
    Private txtCopy As String
    ''' <summary>Было ли событие KeyUp (надо для Ctrl+Delete для случая, когда эта комбинация зажата и не отпускается)</summary>
    Private noKeyUpEvent As Boolean = False
    Private shortcutLinesBefore As Integer
    Private Enum shortcutTypeEnum As Byte
        NONE = 0
        CTRL_BACKSPACE = 1
        CTRL_DELETE = 2
    End Enum
    Private shortcutType As shortcutTypeEnum = shortcutTypeEnum.NONE

    ''' <summary>Было ли нажатие кнопки, после которой список должен быть убран</summary>
    Private keyHideList As Boolean = False
    Private Sub codeBox_KeyDown(ByVal sender As RichTextBoxEx, ByVal e As System.Windows.Forms.KeyEventArgs) Handles codeBox.KeyDown
        If IsNothing(mScript.mainClass) Then Return
        shortcutType = shortcutTypeEnum.NONE
        If e.KeyCode = Keys.Insert Then
            e.SuppressKeyPress = True
            Return
        End If
        'сохраняем начало и длину выделения до редактирования
        If sender.SelectionLength = 0 Then
            codeBox.lastSelectionEndLine = 0
            codeBox.lastSelectionStartLine = 0
        Else
            'Разделитель строк - vbLf (без vbCr)
            Dim curLine As Integer = sender.GetLineFromCharIndex(sender.GetFirstCharIndexOfCurrentLine)
            codeBox.lastSelectionStartLine = sender.GetLineFromCharIndex(sender.SelectionStart)
            codeBox.lastSelectionEndLine = sender.GetLineFromCharIndex(sender.SelectionStart + sender.SelectionLength)
        End If

        Select Case e.KeyCode
            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down, Keys.PageDown, Keys.PageUp
                If lstRtb.Visible Then
                    lstRtb.Focus()
                    SendMessage(lstRtb.Handle, WM_KEYDOWN, e.KeyCode, 0)
                    SendMessage(lstRtb.Handle, WM_KEYUP, e.KeyCode, 0)
                    e.SuppressKeyPress = True
                End If
            Case Keys.Return
                lastKey = e.KeyCode
                If lstRtb.Visible Then
                    codeBox.SelectCurrentWord(sender, lstRtb)
                    Dim prevVal As Boolean = codeBox.rtbCanRaiseEvents
                    codeBox.rtbCanRaiseEvents = False
                    sender.SelectedText = lstRtb.SelectedItem.ToString
                    codeBox.rtbCanRaiseEvents = prevVal
                    lstRtb.Hide()
                End If
                sender.csUndo.lastSelection = sender.SelectedText
                If e.Control Then
                    'codeBox.csUndo.AppendItem(UndoClass.UndoTypeEnum.CHAR_ENTRY, vbLf + "<BR/>")
                    sender.rtbCanRaiseEvents = False
                    codeBox.SelectedText = "<BR/>"
                    sender.rtbCanRaiseEvents = True
                    If sender.Lines.Count - 1 = sender.lastSelectionEndLine Then
                        ReDim Preserve sender.CodeData(sender.CodeData.Count)
                        sender.CodeData(sender.CodeData.Count - 1) = New CodeDataType With {.Comments = "", .StartingSpaces = ""}
                    End If
                Else
                    'codeBox.csUndo.AppendItem(UndoClass.UndoTypeEnum.CHAR_ENTRY, vbLf)
                End If
            Case Keys.Escape
                If lstRtb.Visible Then lstRtb.Hide()
            Case Keys.Back
                'backspace
                keyHideList = True
                'Проверка не стирается ли символ комментария или селектора. Если да, то надо будет прорисовать текст вплоть до конца
                Dim sStart As Integer = sender.SelectionStart
                If sStart >= 4 AndAlso sender.Text.Substring(sStart - 4, 4) = "<!--" Then
                    codeBox.htmlDrawToEnd = True '<!--
                ElseIf sStart >= 3 AndAlso sender.Text.Substring(sStart - 3, 3) = "<!-" AndAlso sStart <= sender.TextLength - 2 AndAlso sender.Text.Chars(sStart) = "-"c Then
                    codeBox.htmlDrawToEnd = True '<!- -
                ElseIf sStart >= 2 AndAlso sender.Text.Substring(sStart - 2, 2) = "<!" AndAlso sStart <= sender.TextLength - 3 AndAlso sender.Text.Substring(sStart, 2) = "--" Then
                    codeBox.htmlDrawToEnd = True '<! --
                ElseIf sStart >= 1 AndAlso sender.Text.Chars(sStart - 1) = "<"c AndAlso sStart <= sender.TextLength - 3 AndAlso sender.Text.Substring(sStart, 3) = "!--" Then
                    codeBox.htmlDrawToEnd = True '< !--
                    'ElseIf sStart >= 3 AndAlso sender.Text.Chars(sStart - 1) = ":"c AndAlso IsNumeric(sender.Text.Chars(sStart - 2)) AndAlso sStart - sender.GetFirstCharIndexOfCurrentLine < 5 Then
                    'htmlDrawToEnd = True '#1:
                End If

                If sender.SelectionLength = 0 AndAlso sender.SelectionStart > 0 AndAlso Asc(sender.Text.Substring(sender.SelectionStart - 1, 1)) = 10 Then
                    Dim curLine As Integer = sender.GetLineFromCharIndex(sender.SelectionStart)
                    If curLine > 0 AndAlso sender.Lines(curLine).Length + sender.Lines(curLine - 1).Length > rtbMaxLineLength Then
                        'предотвращаем слияние строк, если их общая длина будет больше максимальной
                        e.SuppressKeyPress = True
                        lastKey = 0
                        Return
                    End If
                    lastKey = e.KeyCode
                End If

                sender.csUndo.lastSelection = sender.SelectedText
                If e.Control Then 'AndAlso sender.SelectionLength = 0 Then
                    If sender.SelectionLength > 0 Then
                        'удаление выделенного текста делаем отдельным вхождением Undo
                        sender.SelectedText = ""
                        sender.csUndo.lastSelection = sender.SelectedText
                    End If
                    txtCopy = sender.Text
                    shortcutUsedData = sender.SelectionStart
                    shortcutType = shortcutTypeEnum.CTRL_BACKSPACE
                    shortcutLinesBefore = sender.Lines.Count
                Else
                    If sender.SelectionLength = 0 Then
                        If sender.SelectionStart > 0 Then
                            sender.csUndo.AppendItem(UndoClass.UndoTypeEnum.BACKSPACE, sender.Text.Chars(sender.SelectionStart - 1))
                        End If
                    Else
                        sender.csUndo.AppendItem(UndoClass.UndoTypeEnum.BACKSPACE, "") 'sender.SelectedText)
                    End If
                End If
            Case Keys.Delete
                'Проверка не стирается ли символ комментария или селектора. Если да, то надо будет прорисовать текст вплоть до конца
                Dim sStart As Integer = sender.SelectionStart
                If sStart <= sender.TextLength - 5 AndAlso sender.Text.Substring(sStart, 4) = "<!--" Then
                    codeBox.htmlDrawToEnd = True
                ElseIf sStart >= 1 AndAlso sStart <= sender.TextLength - 4 AndAlso sender.Text.Substring(sStart, 3) = "!--" AndAlso sender.Text.Chars(sStart - 1) = "<"c Then
                    codeBox.htmlDrawToEnd = True
                ElseIf sStart >= 2 AndAlso sStart <= sender.TextLength - 3 AndAlso sender.Text.Substring(sStart, 2) = "--" AndAlso sender.Text.Substring(sStart - 2, 2) = "<!" Then
                    codeBox.htmlDrawToEnd = True
                ElseIf sStart >= 3 AndAlso sStart <= sender.TextLength - 4 AndAlso sender.Text.Chars(sStart) = "-"c AndAlso sender.Text.Substring(sStart - 3, 3) = "<!-" Then
                    codeBox.htmlDrawToEnd = True
                End If

                If sender.SelectionLength = 0 AndAlso sender.SelectionStart < sender.TextLength - 1 AndAlso Asc(sender.Text.Substring(sender.SelectionStart, 1)) = 10 Then
                    Dim curLine As Integer = sender.GetLineFromCharIndex(sender.SelectionStart)
                    If curLine < sender.Lines.Length AndAlso sender.Lines(curLine).Length + sender.Lines(curLine + 1).Length > rtbMaxLineLength Then
                        'предотвращаем слияние строк, если их общая длина будет больше максимальной
                        e.SuppressKeyPress = True
                        lastKey = 0
                        Return
                    End If
                    lastKey = e.KeyCode
                End If
                sender.csUndo.lastSelection = sender.SelectedText

                If e.Control Then
                    'комбинация Ctrl+Delete
                    If sender.SelectionLength > 0 Then
                        'удаление выделенного текста делаем отдельным вхождением Undo
                        sender.SelectedText = ""
                        sender.csUndo.lastSelection = sender.SelectedText
                    End If
                    If noKeyUpEvent Then
                        'комбинация зажата (события KeyUp не было) - выполняем KeyUp сейчас
                        Dim sLength As Integer = shortcutUsedData - sender.TextLength
                        sender.csUndo.AppendItem(UndoClass.UndoTypeEnum.DELETE, txtCopy.Substring(sender.SelectionStart, sLength))
                    End If
                    txtCopy = sender.Text
                    shortcutUsedData = sender.TextLength
                    shortcutType = shortcutTypeEnum.CTRL_DELETE
                    shortcutLinesBefore = sender.Lines.Count
                    noKeyUpEvent = True
                Else
                    If sender.SelectionLength = 0 Then
                        If sender.SelectionStart <= sender.TextLength - 1 Then
                            sender.csUndo.AppendItem(UndoClass.UndoTypeEnum.DELETE, sender.Text.Chars(sender.SelectionStart))
                        End If
                    Else
                        sender.csUndo.AppendItem(UndoClass.UndoTypeEnum.DELETE, sender.SelectedText)
                    End If
                End If
            Case Else
                lastKey = 0
                'не допускаем, чтобы строка стала длиннее ее максимальной длины
                Dim selLineIndex As Integer = sender.GetLineFromCharIndex(sender.SelectionStart)
                If sender.Lines.Count >= selLineIndex Then selLineIndex = sender.Lines.Count - 1
                If sender.Lines.GetUpperBound(0) >= selLineIndex AndAlso selLineIndex > -1 Then
                    If sender.Lines(selLineIndex).Length > rtbMaxLineLength Then
                        If (e.KeyCode = Keys.Home OrElse e.KeyCode = Keys.End OrElse e.KeyCode = Keys.Left OrElse e.KeyCode = Keys.Right OrElse e.KeyCode = Keys.Up OrElse e.KeyCode = Keys.Down OrElse e.KeyCode = Keys.PageDown OrElse e.KeyCode = Keys.PageUp) = False Then
                            SendKeys.Send("{ENTER}")
                            Return
                        End If
                    End If
                End If
                'последняя строка выделена до последнего символа (т. е. код последнего выделенного символа = 10), 
                'значит codeBox.lastSelectionEndLine  будет следующая строка, хотя это не так. Уменьшаем значение codeBox.lastSelectionEndLine на 1
                If codeBox.lastSelectionStartLine <> codeBox.lastSelectionEndLine AndAlso _
                Asc(sender.Text.Chars(sender.SelectionStart + sender.SelectionLength - 1)) = 10 Then codeBox.lastSelectionEndLine -= 1
                'Сохраняем в Undo введенный символ (если он вводимый, а не функциональный)
                'Dim ch As Char = Chr(e.KeyCode)
                'If (Char.IsLetterOrDigit(ch) OrElse Char.IsPunctuation(ch) OrElse Char.IsSeparator(ch) OrElse Char.IsWhiteSpace(ch)) AndAlso _
                '    (e.KeyCode <> Keys.End AndAlso e.KeyCode <> Keys.RWin AndAlso e.KeyCode <> Keys.LWin AndAlso e.KeyCode <> Keys.Apps AndAlso e.KeyCode <> Keys.LaunchApplication1 _
                '     AndAlso e.KeyCode <> Keys.LaunchApplication2 AndAlso e.KeyCode <> Keys.NumLock AndAlso e.KeyCode <> Keys.Clear AndAlso e.KeyCode <> Keys.Insert) AndAlso _
                '     (e.KeyCode >= Keys.F1 AndAlso e.KeyCode <= Keys.F24) = False Then
                '    sender.csUndo.AppendItem(UndoClass.UndoTypeEnum.CHAR_ENTRY, Chr(e.KeyCode))
                'End If
                sender.csUndo.lastSelection = sender.SelectedText
        End Select

    End Sub

    Private Sub codeBox_KeyPress(ByVal sender As RichTextBoxEx, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles codeBox.KeyPress
        If IsNothing(mScript.mainClass) Then Return
        'Проверка на ввод символа (в режиме html), который требует расширения области зарисовки текста вплоть до конца (символы комментариев, селектора #1:)
        Dim chrAsc As Integer = Asc(e.KeyChar)
        '127 - Ctrl+Backspace
        If chrAsc <> Keys.Return AndAlso chrAsc <> 10 AndAlso chrAsc <> Keys.Back AndAlso chrAsc <> 127 Then
            sender.csUndo.AppendItem(UndoClass.UndoTypeEnum.CHAR_ENTRY, e.KeyChar)
        ElseIf chrAsc = 127 Then
            'shortcutUsedData содержит SelectionStart до Ctrl+Backspace
            Dim sLength As Integer = shortcutUsedData - sender.SelectionStart '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! (sender.SelectionStart может длинее txtCopy.length)
            If IsNothing(txtCopy) = False AndAlso sLength > 0 Then sender.csUndo.AppendItem(UndoClass.UndoTypeEnum.BACKSPACE, txtCopy.Substring(sender.SelectionStart, sLength), sender.SelectionStart)
        End If
        codeBox.htmlDrawToEnd = False
        If e.KeyChar = ">"c Then
            Dim sStart As Integer = sender.SelectionStart
            If sStart >= 2 AndAlso sender.Text.Substring(sStart - 2, 2) = "--" Then
                codeBox.htmlDrawToEnd = True
            End If
        ElseIf e.KeyChar = "-"c Then
            Dim sStart As Integer = sender.SelectionStart

            If sStart <= sender.TextLength - 2 AndAlso sStart >= 1 AndAlso sender.Text.Chars(sStart - 1) = "-"c AndAlso sender.Text.Chars(sStart) = ">"c Then
                codeBox.htmlDrawToEnd = True
            ElseIf sStart <= sender.TextLength - 3 AndAlso sender.Text.Substring(sStart, 2) = "->" Then
                codeBox.htmlDrawToEnd = True
            End If
        ElseIf e.KeyChar = ":"c Then
            Dim sStart As Integer = sender.SelectionStart
            Dim curLine As Integer = sender.GetLineFromCharIndex(sStart)
            Dim firstInLine As Integer = sender.GetFirstCharIndexFromLine(curLine)
            If sender.Text.Chars(firstInLine) = "#"c AndAlso sStart - firstInLine < 5 Then
                codeBox.htmlDrawToEnd = True
            End If
        ElseIf e.KeyChar = "#"c Then
            Dim sStart As Integer = sender.SelectionStart
            Dim curLine As Integer = sender.GetLineFromCharIndex(sStart)
            Dim firstInLine As Integer = sender.GetFirstCharIndexFromLine(curLine)
            If sStart = firstInLine AndAlso sStart <= sender.TextLength - 3 AndAlso IsNumeric(sender.Text.Chars(sStart)) AndAlso (sender.Text.Chars(sStart + 1) = ":"c OrElse sender.Text.Chars(sStart + 2) = ":"c) Then
                codeBox.htmlDrawToEnd = True
            End If
        ElseIf IsNumeric(e.KeyChar) Then
            Dim sStart As Integer = sender.SelectionStart
            If (sStart <= sender.TextLength - 2 AndAlso sender.Text.Chars(sStart) = ":"c) OrElse (sStart <= sender.TextLength - 3 AndAlso sender.Text.Chars(sStart + 1) = ":"c AndAlso IsNumeric(sender.Text.Chars(sStart))) Then
                Dim curLine As Integer = sender.GetLineFromCharIndex(sStart)
                Dim firstInLine As Integer = sender.GetFirstCharIndexFromLine(curLine)
                If sender.Text.Chars(firstInLine) = "#"c AndAlso sStart - firstInLine < 5 Then
                    codeBox.htmlDrawToEnd = True
                End If
            End If
        End If

        'codebox.wordBoundsArray = {" "c, "("c, ")"c, "["c, "]"c, "."c, ","c, "'"c, "="c, "+"c, "-"c, "*"c, "^"c, "/"c, "\"c, "<"c, ">"c, "&"c, "!"c, ";"c, "#"c, "$"c, Chr(34), Chr(10), "?"}
        If lstRtb.Visible AndAlso Array.IndexOf(codeBox.wordBoundsArray, e.KeyChar) > -1 Then
            'если введен символ, означающий конец слова (пробел, [, " ...), заменяем последнее введенное слово (или его часть) на полное слово из списка
            Dim lineStart As Integer = codeBox.GetFirstCharIndexOfCurrentLine
            If codeBox.SelectionStart - lineStart + 1 >= 0 Then
                'если виден список и введена кавычка ', при этом она непарная (вводится строка), то игнорируем ее как сигнал к закрытию списка
                If mScript.IsQuotesCountEven(codeBox.Text.Substring(lineStart, codeBox.SelectionStart - lineStart) + e.KeyChar) = False Then Return
            End If

            'виден список и нажат символ, который не является продолжением вводимого слова
            'вместо введеного (вероятно, неполного) слова надо вставить слово, выбранное в списке
            If e.KeyChar = "="c OrElse e.KeyChar = ">"c Then
                'проверка на != <> <= >=   Т. е., проверяем не вводится ли оператор сравнения, состоящий из 2 символов
                Dim prevChar As Char = sender.Text.Substring(0, sender.SelectionStart).Trim.Last
                If prevChar = "!"c And e.KeyChar = "="c Then Return 'Это !=   - продолжается ввод слова
                If prevChar = ">"c OrElse prevChar = "<"c Then Return 'Это <= >= <>   - продолжается ввод слова
            End If

            Dim prevVal As Boolean = codeBox.rtbCanRaiseEvents
            codeBox.rtbCanRaiseEvents = False
            SendMessage(sender.Handle, WM_SetRedraw, 0, 0)
            codeBox.SelectCurrentWord(sender, lstRtb) 'выделяем текущее слово
            Dim selItem As String = lstRtb.SelectedItem.ToString
            If selItem.Length > 0 AndAlso selItem.Last = "'" AndAlso e.KeyChar = "'" Then e.KeyChar = "" 'selItem = selItem.Substring(0, selItem.Length - 1)
            sender.SelectedText = selItem  'заменяем его словом из списка
            SendMessage(sender.Handle, WM_SetRedraw, 1, 0)
            sender.Refresh()
            codeBox.rtbCanRaiseEvents = prevVal
            lstRtb.Hide() 'убираем список
        End If
    End Sub

    Private Sub codeBox_KeyUp(ByVal sender As RichTextBoxEx, ByVal e As System.Windows.Forms.KeyEventArgs) Handles codeBox.KeyUp
        noKeyUpEvent = False
        keyHideList = False
        shortcutType = shortcutTypeEnum.NONE
        If IsNothing(mScript.mainClass) Then Return
        If codeBox.rtbCanRaiseEvents = False Then Return
        sender.csUndo.lastSelection = ""
        Select Case e.KeyCode
            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down, Keys.PageDown, Keys.PageUp
                codeBox.CheckPrevLine(sender)
                'если был совершен переход на новую линию - сохраняем информацию о текущем содержимом строки, которую будут редактировать
                If codeBox.prevLineId <> sender.GetLineFromCharIndex(sender.GetFirstCharIndexOfCurrentLine) Then codeBox.SetPrevLine(sender)
            Case Keys.Delete
                If e.Control Then
                    'Ctrl+Delete
                    Dim sLength As Integer = shortcutUsedData - sender.TextLength
                    sender.csUndo.AppendItem(UndoClass.UndoTypeEnum.DELETE, txtCopy.Substring(sender.SelectionStart, sLength))
                End If
        End Select
    End Sub

    Private Sub CodeTextBox_Load(sender As Object, e As EventArgs) Handles Me.Load
        SplitContainerVertical.FixedPanel = FixedPanel.Panel1
    End Sub

    Private Sub CodeTextBox_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        codeBox.Location = Point.Empty
        codeBox.Size = New Size(SplitContainerVertical.Panel2.Width, SplitContainerVertical.Panel2.Height)
    End Sub

    Private Sub CodeTextBox_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Validating
        If codeBox.rtbCanRaiseEvents = False Then Return
        If lstRtb.Visible Then lstRtb.Hide()
        Dim curLine As Integer = codeBox.GetLineFromCharIndex(codeBox.GetFirstCharIndexOfCurrentLine)
        codeBox.rtbCanRaiseEvents = False
        codeBox.PrepareText(codeBox, curLine, curLine)
        codeBox.PrepareHTMLforErrorInfo(wbRtbHelp)
        If codeBox.CheckTextForSyntaxErrors(codeBox) Then
            codeBox.CheckCycles(codeBox)
        End If
        codeBox.rtbCanRaiseEvents = True
    End Sub

#End Region

#Region "List And Information"
    Private Sub lstRtb_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstRtb.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Return Then
            'Del или Enter - пересылаем это нажатие нашему rtb
            SendMessage(codeBox.Handle, WM_KEYDOWN, e.KeyCode, 0)
            SendMessage(codeBox.Handle, WM_KEYUP, e.KeyCode, 0)
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Escape Then
            'Esc - прячем список
            lstRtb.Visible = False
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub lstRtb_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles lstRtb.KeyPress
        If Asc(e.KeyChar) = 8 Then 'backspace
            'пересылаем это нажатие нашему rtb
            SendMessage(codeBox.Handle, WM_KEYDOWN, 8, 0)
            SendMessage(codeBox.Handle, WM_KEYUP, 8, 0)
        Else
            'введен знак, являющийся новым словом (или разделителем). заменяем введенное пользователем на то, что в списке
            If Array.IndexOf(codeBox.wordBoundsArray, e.KeyChar) > -1 Then
                codeBox.SelectCurrentWord(codeBox, lstRtb)
                Dim selitem As String = lstRtb.SelectedItem.ToString
                If e.KeyChar = "'"c AndAlso selitem.Length > 0 AndAlso selitem.Last = "'"c Then
                    'selitem = selitem.Substring(0, selitem.Length - 1)
                    e.KeyChar = ""
                End If
                codeBox.SelectedText = selitem
                lstRtb.Hide()
            End If
            codeBox.SelectedText = e.KeyChar.ToString
        End If
        codeBox.Focus()
    End Sub

    Private Sub lstRtb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstRtb.SelectedIndexChanged
        'FuncAndProps
        If IsNothing(sender.Tag) OrElse sender.Tag.ToString.Length = 0 Then Return
        Dim arrTag() As String = sender.Tag.ToString.Split(" "c)
        If IsNothing(arrTag) OrElse arrTag.Count = 0 Then Return
        Select Case arrTag(0)
            Case "FuncAndProps"
                'Выводим помощь по функции или свойству
                Dim classId As Integer = CInt(arrTag(1))
                'Dim elName As String = sender.Text
                Dim elName As String
                If sender.SelectedIndex = -1 Then
                    elName = sender.Items(0).ToString
                Else
                    elName = sender.Text
                End If
                Dim isFunc As Boolean = True
                If mScript.mainClass(classId).Properties.ContainsKey(elName) Then isFunc = False
                If isFunc = True AndAlso mScript.mainClass(classId).Functions.ContainsKey(elName) = False Then Return
                codeBox.ShowElementInfo(classId, lstRtb.Text, IIf(isFunc, EditWordTypeEnum.W_FUNCTION, EditWordTypeEnum.W_PROPERTY))
            Case "Variables"
                codeBox.ShowElementInfo(-1, lstRtb.Text, EditWordTypeEnum.W_VARIABLE)
        End Select
    End Sub

    Private Sub lstRtb_VisibleChanged(ByVal sender As ListBox, ByVal e As System.EventArgs) Handles lstRtb.VisibleChanged
        If sender.Visible = False AndAlso codeBox.Visible = True Then codeBox.Focus()
    End Sub

    ''' <summary>
    ''' Процедура локального действия (для PrepareRtbList) наносит последние штрихи в подготовке списка
    ''' </summary>
    ''' <param name="arrValues">список для вставки</param>
    ''' <param name="lastRange">переменная для хранения последнего добавленного списка</param>
    ''' <remarks></remarks>
    Private Sub PrepareListWithRange(ByRef arrValues() As String, ByRef lastRange() As String, Optional ByVal append As Boolean = False)
        If Not append Then
            lstRtb.Items.Clear()
            lstRtb.Items.AddRange(arrValues)
            lastRange = arrValues
        Else
            lstRtb.Items.AddRange(arrValues)
            If lstRtb.Items.Count = 0 Then
                lastRange = Nothing
            Else
                ReDim lastRange(lstRtb.Items.Count - 1)
                lstRtb.Items.CopyTo(lastRange, 0)
            End If
        End If
        lstRtb.Tag = Nothing
        If lstRtb.Items.Count > 0 Then
            lstRtb.Show()
            lstRtb.SelectedIndex = 0
        Else
            lstRtb.Hide()
        End If
    End Sub

    ''' <summary>
    ''' Процедура подготавливает список. т. е. создает его из нужных элементов (например, имен переменных), или же, если ввод слова продолжается, производит поиск введеной части слова в полном списке 
    ''' и выводит только подходящие результаты
    ''' </summary>
    ''' <param name="currentLine">Текущая линия в кодбоксе</param>
    ''' <param name="currentWordId">Id текущего (вводимого/введенного) слова с соответствующей структуре codeBox.CodeData</param> 
    ''' <remarks></remarks>
    Private Sub PrepareRtbList(ByVal currentLine As Integer, ByVal currentWordId As Integer)
        If IsNothing(mScript.mainClass) Then Return
        lstRtb.Tag = Nothing
        If currentWordId = -1 Then
            'ввод за пределами кода (например, в комментарии) - выход
            lstRtb.Hide()
            Return
        End If
        If codeBox.SelectionStart = 0 Then Return

        Static prevLine As Integer = -1, prevWordId As Integer = -1  'линия и слово, с которыми процедура работала прошлый раз
        Static lastRange() As String  'последний созданный список
        'получаем координаты для размещения списка
        Dim curCharPos As Point = codeBox.GetPositionFromCharIndex(codeBox.SelectionStart + codeBox.SelectionLength)
        lstRtb.Left = curCharPos.X + codeBox.Left  'Х - сразу за кареткой
        Dim newPosTop As Integer = curCharPos.Y + codeBox.Top + codeBox.Font.Height + 2 'Y - на 2 пикселя ниже каретки 
        If newPosTop + lstRtb.Height > codeBox.Top + codeBox.Height Then
            lstRtb.Top = curCharPos.Y + codeBox.Top - lstRtb.Height
        Else
            lstRtb.Top = newPosTop
        End If

        'если слова с указанным индексом в указанной строке не существует - выход
        If codeBox.GetWordFromCodeData(currentLine, currentWordId).wordType = EditWordTypeEnum.W_NOTHING Then
            lstRtb.Hide()
            Return
        End If

        Dim curWord As EditWordType = codeBox.CodeData(currentLine).Code(currentWordId) 'указанное (последнее введенное) слово
        Dim chkWord As String = curWord.Word.Trim
        Select Case curWord.wordType
            Case EditWordTypeEnum.W_WRAP 'введено Wrap
                prevLine = -1
                If codeBox.Text.Chars(codeBox.SelectionStart - 1) <> " "c OrElse curWord.Word.Trim <> "Wrap" Then
                    lstRtb.Hide()
                    Return
                End If
                If currentWordId > 0 AndAlso codeBox.CodeData(currentLine).Code(currentWordId - 1).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    'End Wrap - ничего не показываем
                    lstRtb.Hide()
                    Return
                End If
                'создаем новый список из переменных
                Dim arrVars() As String = {}
                PrepareListWithRange(MakeVariablesList(codeBox, False), lastRange)
                cListManager.MakeLongTextPropertiesList(lastRange)
                'сохраняем последние линию и индекс слова 
                prevLine = currentLine
                prevWordId = currentWordId
            Case EditWordTypeEnum.W_BLOCK_DOWHILE 'введено Do
                prevLine = -1
                If codeBox.Text.Chars(codeBox.SelectionStart - 1) <> " "c OrElse curWord.Word.Trim <> "Do" Then
                    lstRtb.Hide()
                    Return
                End If
                'создаем новый список из одного слова
                PrepareListWithRange({"While"}, lastRange)
                'сохраняем последние линию и индекс слова 
                prevLine = currentLine
                prevWordId = currentWordId
            Case EditWordTypeEnum.W_HTML 'введено HTML
                prevLine = -1
                If codeBox.Text.Chars(codeBox.SelectionStart - 1) <> " "c OrElse curWord.Word.Trim <> "HTML" Then
                    lstRtb.Hide()
                    Return
                End If
                If currentWordId > 0 AndAlso codeBox.CodeData(currentLine).Code(currentWordId - 1).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    lstRtb.Hide()
                    Return
                End If
                'создаем новый список из двух слов - "" и "Append"
                PrepareListWithRange({"", "Append"}, lastRange)
                'сохраняем последние линию и индекс слова 
                prevLine = currentLine
                prevWordId = currentWordId
            Case EditWordTypeEnum.W_BLOCK_EVENT 'введено Event
                prevLine = -1
                If codeBox.Text.Chars(codeBox.SelectionStart - 1) <> " "c Then
                    lstRtb.Hide()
                    Return
                End If
                If currentWordId > 0 AndAlso codeBox.CodeData(currentLine).Code(currentWordId - 1).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    lstRtb.Hide()
                    Return
                End If
                'создаем новый список из всех свойств, имеющих тип "событие" (RETURN_EVENT)
                'Erase lastRange
                'Dim rangeSize As Integer = 0
                'For i As Integer = 0 To mScript.mainClass.GetUpperBound(0)
                '    If IsNothing(mScript.mainClass(i).Properties) Then Continue For
                '    Dim className As String = mScript.mainClass(i).Names(0)
                '    For Each prop As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(i).Properties
                '        If prop.Value.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse prop.Value.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_CODE Then Continue For
                '        'проверка каждого свойства в каждом классе
                '        If prop.Value.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                '            'тип свойства - событие. добавляе его в список.
                '            rangeSize += 1
                '            Array.Resize(Of String)(lastRange, rangeSize)
                '            lastRange(rangeSize - 1) = "'" + className + "." + prop.Key + "'"
                '        End If
                '    Next
                'Next
                lastRange = cListManager.GetEventsList(True)
                lstRtb.Items.Clear()
                If IsNothing(lastRange) = False Then lstRtb.Items.AddRange(lastRange)
                If lstRtb.Items.Count > 0 Then
                    lstRtb.Show()
                    lstRtb.SelectedIndex = 0
                Else
                    lstRtb.Hide()
                End If
                'сохраняем последние линию и индекс слова 
                prevLine = currentLine
                prevWordId = currentWordId
            Case EditWordTypeEnum.W_BLOCK_FUNCTION 'введене Function 
                prevLine = -1
                If codeBox.Text.Chars(codeBox.SelectionStart - 1) <> " "c Then
                    lstRtb.Hide()
                    Return
                End If
                If currentWordId > 0 AndAlso codeBox.CodeData(currentLine).Code(currentWordId - 1).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    lstRtb.Hide()
                    Return
                End If
                'создаем новый список из всех функций и функций пользователя
                Erase lastRange
                Dim rangeSize As Integer = 0
                For i As Integer = 0 To mScript.mainClass.GetUpperBound(0)
                    If IsNothing(mScript.mainClass(i).Properties) Then Continue For
                    Dim className As String = mScript.mainClass(i).Names(0)
                    For Each prop As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(i).Functions
                        If prop.Value.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse prop.Value.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_CODE OrElse _
                            prop.Value.UserAdded = False Then Continue For
                        'функций пользователя в каждом классе
                        rangeSize += 1
                        Array.Resize(Of String)(lastRange, rangeSize)
                        lastRange(rangeSize - 1) = "'" + className + "." + prop.Key + "'"
                    Next
                Next
                If IsNothing(mScript.functionsHash) = False AndAlso mScript.functionsHash.Count > 0 Then
                    For i As Integer = 0 To mScript.functionsHash.Count - 1
                        rangeSize += 1
                        Array.Resize(Of String)(lastRange, rangeSize)
                        lastRange(rangeSize - 1) = "'" + mScript.functionsHash.ElementAt(i).Key + "'"
                    Next
                End If
                lstRtb.Items.Clear()
                If IsNothing(lastRange) = False Then lstRtb.Items.AddRange(lastRange)
                If lstRtb.Items.Count > 0 Then
                    lstRtb.Show()
                    lstRtb.SelectedIndex = 0
                Else
                    lstRtb.Hide()
                End If
                'сохраняем последние линию и индекс слова 
                prevLine = currentLine
                prevWordId = currentWordId
            Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN, EditWordTypeEnum.W_OVAL_BRACKET_OPEN, _
                EditWordTypeEnum.W_COMMA, EditWordTypeEnum.W_QUAD_BRACKET_CLOSE, _
                EditWordTypeEnum.W_OVAL_BRACKET_CLOSE, EditWordTypeEnum.W_FUNCTION
                'это может быть или функцией, или свойством (ну, или ни тем, ни другим - тогда выход)
                prevLine = -1
                Dim initialWord As EditWordType = curWord
                Dim initialWordId As Integer = currentWordId
                Dim paramNumber As Integer
                Dim classId As Integer
                Dim funcPos As Integer = -1
                'получаем слово с этой функцией/свойством и номер редактируемого в данный момент параметра
                If GetCurrentElementWord(currentLine, currentWordId, paramNumber, curWord, classId, funcPos) = False AndAlso classId < 0 Then
                    'это ни функция, ни свойство 
                    If codeBox.CodeData(currentLine).Code(0).wordType = EditWordTypeEnum.W_BLOCK_EVENT AndAlso initialWord.wordType = EditWordTypeEnum.W_COMMA Then
                        Dim cId As Integer = -1, eventName As String = ""
                        Dim params() As String = GetEventParamValues(currentLine, cId, eventName)
                        If cId >= 0 Then
                            'это событие, получены класс и его имя
                            If params.Count = 0 AndAlso mScript.mainClass(cId).LevelsCount >= 1 Then
                                'выводим элементы 2 уровня
                                PrepareListWithRange(cListManager.FillListByClassName(mScript.mainClass(cId).Names(0)), lastRange)
                                PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                            ElseIf params.Count = 1 AndAlso mScript.mainClass(cId).LevelsCount = 2 Then
                                'выводим элементы 3 уровня
                                Dim child2Id As Integer = GetSecondChildIdByName(params(0), mScript.mainClass(cId).ChildProperties)
                                If child2Id >= 0 Then
                                    PrepareListWithRange(cListManager.CreateThirdLevelElementsList(cId, child2Id), lastRange)
                                    PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                                Else
                                    lstRtb.Hide()
                                    Return
                                End If
                            End If
                            'сохраняем последние линию и индекс слова 
                            prevLine = currentLine
                            prevWordId = currentWordId
                        Else
                            'это событие, но параметры получить не удается - выход
                            lstRtb.Hide()
                            Return
                        End If
                    ElseIf currentWordId > 2 AndAlso initialWord.wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN AndAlso codeBox.CodeData(currentLine).Code(currentWordId - 1).wordType = EditWordTypeEnum.W_VARIABLE Then
                        'это переменная-массив, за которой поставили [
                        'Здесь ничего не делаем, обработка пойдет ниже
                    Else
                        'это ни функция, ни свойство, ни событие - выход
                        lstRtb.Hide()
                        Return
                    End If
                End If

                If IsNothing(curWord.Word) = False Then curWord.Word = curWord.Word.Trim
                Dim paramStarted As Boolean = True
                If funcPos = currentWordId Then
                    Dim lastChar As Char = codeBox.Text.Chars(codeBox.SelectionStart - 1)
                    If lastChar <> " "c AndAlso lastChar <> "("c Then paramStarted = False
                End If

                If curWord.wordType = EditWordTypeEnum.W_FUNCTION AndAlso curWord.classId >= 0 AndAlso paramStarted Then
                    'если это функция, то делаем следующее:
                    'в func получаем структуру со всеми свойствами функции
                    Dim func As MatewScript.PropertiesInfoType = mScript.mainClass(curWord.classId).Functions(curWord.Word)
                    Dim paramsDescribed As Integer = 0
                    If IsNothing(func.params) = False Then paramsDescribed = func.params.Count
                    'проверка не находится ли paramNumber за пределами кол-ва описанных параметров функции
                    If func.paramsMax = -1 Then
                        'последний параметр описан как "массив параметров" (т. е. кол-во параметров не ограничено)
                        If paramNumber > paramsDescribed Then 'если каретка в одном из этих параметров (а не раньше, в предыдущих параметрах с другим типом) - выход
                            lstRtb.Hide()
                            Return
                        End If
                    ElseIf paramNumber > func.paramsMax Then
                        'paramNumber больше кол-ва описанных параметров (вообще-то это ошибка)
                        lstRtb.Hide()
                        Return
                    End If

                    If paramNumber = 0 Then paramNumber = 1
                    Dim param As MatewScript.paramsType = func.params(paramNumber - 1)
                    Dim isEnum As Boolean = (IsNothing(param.EnumValues) = False AndAlso param.EnumValues.Count > 0)
                    Dim propTranslateFrom As MatewScript.PropertiesInfoType = Nothing
                    Dim isTranslated As Boolean = False
                    If isEnum AndAlso param.EnumValues.Count = 1 Then
                        Dim enValue As String = mScript.PrepareStringToPrint(param.EnumValues(0), Nothing, False)
                        If enValue.StartsWith("[ByProperty]") Then
                            'вместо данного списка надо отобразить список из указанного свойства
                            enValue = enValue.Substring("[ByProperty]".Length) 'теперь здесь только имя класса и свойства (напр., O.EquipType)
                            Dim arrVal() As String = enValue.Split("."c) '0 - класс, 1 - свойство
                            If IsNothing(arrVal) = False AndAlso arrVal.Count = 2 Then
                                Dim cId As Integer = -1, cProp As String = ""
                                If mScript.mainClassHash.TryGetValue(arrVal(0), cId) Then
                                    mScript.mainClass(cId).Properties.TryGetValue(arrVal(1), propTranslateFrom)
                                    If IsNothing(propTranslateFrom.returnArray) OrElse propTranslateFrom.returnArray.Count = 0 Then
                                        propTranslateFrom = Nothing
                                    Else
                                        isTranslated = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Select Case param.Type
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_BOOL
                            'тип параметра - Да/Нет. Составляем списов с этими двумя словами
                            PrepareListWithRange({"True", "False"}, lastRange)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM
                            'тип параметра - список значений. Составляем этот список
                            If Not isEnum Then lstRtb.Hide() : Return
                            If Not isTranslated Then
                                PrepareListWithRange(param.EnumValues, lastRange)
                            Else
                                PrepareListWithRange(propTranslateFrom.returnArray, lastRange)
                            End If
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_USER_FUNCTION
                            'тип параметра - список функций пользователя. Составляем его
                            PrepareListWithRange(cListManager.GetFunctionsList, lastRange)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT
                            'тип параметра - элемент 2 уровня. Составляем его
                            If Not isEnum Then lstRtb.Hide() : Return
                            PrepareListWithRange(cListManager.FillListByClassName(param.EnumValues(0)), lastRange)
                            PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT2
                            'тип параметра - элемент 3 уровня. Составляем его
                            Dim elementParam As MatewScript.paramsType = Nothing
                            'Класс элемента 3 уровня должен быть идентичен классу элемента 2 уровня, указанного в этой же фукнции (базовые принципы построения), который также должен являтся его родителем
                            'Если элемент 2 уровня указан явно (указано его Имя или Id), то получаем его. Иначе el2 будет содержать пустую строку - 
                            'вычислить родителя на этапе написания кода невозможно, соответственно, построить список также невозможно
                            Dim el2 As String = GetFunctionParamWhichTypeIsElement(currentLine, funcPos, func.params, elementParam)
                            If String.IsNullOrEmpty(el2) = False AndAlso IsNothing(elementParam.EnumValues) = False AndAlso elementParam.EnumValues.Count > 0 Then
                                'Элемент 2 уровня успешно получен
                                Dim cId As Integer = mScript.mainClassHash(elementParam.EnumValues(0)) 'получаем его класс
                                Dim child2Id As Integer = GetSecondChildIdByName(el2, mScript.mainClass(cId).ChildProperties) 'получаем его индекс в ChildProperties
                                If child2Id >= 0 Then
                                    PrepareListWithRange(cListManager.CreateThirdLevelElementsList(cId, child2Id), lastRange) 'получаем список элементов 3 уровня
                                    PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                                End If
                            End If
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_AUDIO
                            PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.AUDIO), lastRange)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_CSS
                            PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.CSS), lastRange)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_JS
                            PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.JS), lastRange)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_PICTURE
                            PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.PICTURES), lastRange)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_PATH_TEXT
                            PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.TEXT), lastRange)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_VARIABLE
                            'составляем список переменных    
                            PrepareListWithRange(MakeVariablesList(codeBox, True), lastRange)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_EVENT
                            PrepareListWithRange(cListManager.GetEventsList(True), lastRange)
                        Case MatewScript.paramsType.paramsTypeEnum.PARAM_PROPERTY
                            PrepareListWithRange(cListManager.GetPropertiesList(True), lastRange)
                        Case Else
                            If isEnum Then
                                'тип параметра не предполагает список, однако он был создан Писателем. Составляем этот список
                                If Not isTranslated Then
                                    PrepareListWithRange(param.EnumValues, lastRange)
                                Else
                                    PrepareListWithRange(propTranslateFrom.returnArray, lastRange)
                                End If
                            Else
                                lstRtb.Hide()
                                Return
                            End If
                    End Select
                    'сохраняем последние линию и индекс слова 
                    prevLine = currentLine
                    prevWordId = currentWordId
                ElseIf curWord.wordType = EditWordTypeEnum.W_PROPERTY AndAlso curWord.classId >= 0 AndAlso paramStarted Then
                    'если это свойство, то делаем следующее:
                    'в prop получаем структуру со всеми свойствами свойства
                    Dim prop As MatewScript.PropertiesInfoType = Nothing
                    mScript.mainClass(curWord.classId).Properties.TryGetValue(curWord.Word, prop)
                    If IsNothing(prop) Then
                        lstRtb.Hide()
                        Return
                    End If
                    If paramNumber = 0 Then paramNumber = 1
                    'paramNumber может быть только 1 или 2. Если 1, то там должно быть Имя/Id элемента 2 уровня, если 2 - третьего.
                    If mScript.mainClass(classId).LevelsCount > 0 AndAlso IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso paramNumber = 1 Then
                        'Это свойство 2-х или 3-х уровневого класса, параметр № 1. Создаем список элементов 2 уровня для этого класса
                        PrepareListWithRange(cListManager.FillListByClassName(mScript.mainClass(classId).Names(0)), lastRange)
                        PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                    ElseIf paramNumber = 2 AndAlso currentWordId > 2 AndAlso mScript.mainClass(classId).LevelsCount = 2 Then
                        'Это свойство 3-х уровневого класса, параметр № 2. Создаем список элементов 3 уровня для этого класса
                        'Класс элемента 3 уровня должен быть идентичен классу элемента 2 уровня, указанного первым параметром (сейчас в curWord), который также должен являтся его родителем
                        'Если элемент 2 уровня указан явно (указано его Имя или Id), то получаем его. Иначе вычислить родителя на этапе написания кода невозможно, соответсвенно, 
                        'построить список также невозможно
                        Dim arrValues() = GetPropertyParamValues(currentLine, funcPos) 'в массиве - все параметры свойства. нас интересует первый, в котром, предположительно, Имя/Id элемента 2 уровня
                        If IsNothing(arrValues) = False AndAlso arrValues.Count > 0 AndAlso String.IsNullOrEmpty(arrValues(0)) = False Then
                            'тут можно быть уверенным, что в arrValues(0) находится простая строка/число, которая должна быть именем/Id элемента-родителя 2 уровня
                            Dim child2Id As Integer = GetSecondChildIdByName(arrValues(0), mScript.mainClass(classId).ChildProperties)
                            If child2Id >= 0 Then
                                'элемент 2 уровня получен. Теперь все готово для построения списка элементов 3 уровня
                                PrepareListWithRange(cListManager.CreateThirdLevelElementsList(classId, child2Id), lastRange) 'получаем список элементов 3 уровня
                                PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                            End If
                        Else
                            lstRtb.Hide()
                            Return
                        End If
                    Else 'возвращаемое значение другое - выход
                        lstRtb.Hide()
                        Return
                    End If

                    'сохраняем последние линию и индекс слова 
                    prevLine = currentLine
                    prevWordId = currentWordId
                ElseIf classId >= 0 AndAlso curWord.wordType <> EditWordTypeEnum.W_NOTHING AndAlso curWord.wordType <> EditWordTypeEnum.W_FUNCTION Then
                    'это свойство
                    If mScript.mainClass(classId).LevelsCount > 0 AndAlso IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso paramNumber = 1 Then
                        'Это свойство 2-х или 3-х уровневого класса, параметр № 1. Создаем список элементов 2 уровня для этого класса
                        PrepareListWithRange(cListManager.FillListByClassName(mScript.mainClass(classId).Names(0)), lastRange)
                        PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                    Else 'возвращаемое значение другое - выход
                        lstRtb.Hide()
                        Return
                    End If
                    'сохраняем последние линию и индекс слова 
                    prevLine = currentLine
                    prevWordId = currentWordId
                ElseIf initialWord.wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN AndAlso currentWordId > 0 Then
                    'это либо переменная, либо класс, после которых стоит [
                    curWord = codeBox.CodeData(currentLine).Code(currentWordId - 1)
                    If curWord.wordType = EditWordTypeEnum.W_VARIABLE Then
                        'это переменная-массив, за которой поставили [
                        classId = -1

                        PrepareListWithRange(MakeVariablesSignaturesList(codeBox, curWord.Word.Trim), lastRange)
                    Else
                        'это ClassName[
                        classId = curWord.classId

                        If curWord.classId > -2 AndAlso mScript.mainClass(classId).LevelsCount > 0 AndAlso IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso paramNumber = 1 Then
                            'Это свойство 2-х или 3-х уровневого класса, параметр № 1. Создаем список элементов 2 уровня для этого класса
                            PrepareListWithRange(cListManager.FillListByClassName(mScript.mainClass(classId).Names(0)), lastRange)
                            PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                        Else 'возвращаемое значение другое - выход
                            lstRtb.Hide()
                            Return
                        End If
                    End If
                    'сохраняем последние линию и индекс слова 
                    prevLine = currentLine
                    prevWordId = currentWordId
                ElseIf initialWord.wordType = EditWordTypeEnum.W_COMMA AndAlso currentWordId > 2 AndAlso classId >= 0 AndAlso mScript.mainClass(classId).LevelsCount = 2 Then
                    'Это свойство 3-х уровневого класса, параметр № 2. Создаем список элементов 3 уровня для этого класса
                    '(точнее, это может быть и функция - введено нечто вроде "M['menu1',")
                    curWord = codeBox.CodeData(currentLine).Code(currentWordId - 1) 'здесь должен стоять элемент 2 уровня
                    If curWord.wordType = EditWordTypeEnum.W_SIMPLE_STRING OrElse curWord.wordType = EditWordTypeEnum.W_SIMPLE_NUMBER Then
                        'Класс элемента 3 уровня должен быть идентичен классу элемента 2 уровня, указанного первым параметром (сейчас в curWord), который также должен являтся его родителем
                        'Если элемент 2 уровня указан явно (указано его Имя или Id), то получаем его. Иначе вычислить родителя на этапе написания кода невозможно, соответсвенно, 
                        'построить список также невозможно
                        Dim prevWordType As EditWordTypeEnum = codeBox.CodeData(currentLine).Code(currentWordId - 2).wordType  'слово перед предполагаемым 2 элементом
                        If prevWordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN OrElse prevWordType = EditWordTypeEnum.W_COMMA Then
                            'тут можно быть уверенным, что в CurWord находится простая строка/число, которая должна быть именем/Id элемента-родителя 2 уровня
                            Dim child2Id As Integer = GetSecondChildIdByName(curWord.Word.Trim, mScript.mainClass(classId).ChildProperties)
                            If child2Id >= 0 Then
                                'элемент 2 уровня получен. Теперь все готово для построения списка элементов 3 уровня
                                PrepareListWithRange(cListManager.CreateThirdLevelElementsList(classId, child2Id), lastRange) 'получаем список элементов 3 уровня
                                PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                                'сохраняем последние линию и индекс слова 
                                prevLine = currentLine
                                prevWordId = currentWordId
                            End If
                        End If
                    End If
                End If
            Case EditWordTypeEnum.W_OPERATOR_EQUAL, EditWordTypeEnum.W_OPERATOR_COMPARE 'операторы сравнения
                'если перед ними стоят функция или свойство, возвращающее некий список - здесь этот список должен отобразиться
                prevLine = -1
                If currentWordId = 0 Then
                    lstRtb.Hide()
                    Return
                End If
                curWord = codeBox.CodeData(currentLine).Code(currentWordId - 1) 'в curWord получаем слово перед текущим
                curWord.Word = curWord.Word.Trim

                'В retType получаем тип возвращаемого значения функции/свойства
                Dim retType As MatewScript.ReturnFunctionEnum = MatewScript.ReturnFunctionEnum.RETURN_USUAL
                Dim prop As MatewScript.PropertiesInfoType = Nothing
                If curWord.classId >= 0 AndAlso (curWord.wordType = EditWordTypeEnum.W_PROPERTY OrElse curWord.wordType = EditWordTypeEnum.W_FUNCTION) Then
                    If curWord.wordType = EditWordTypeEnum.W_PROPERTY Then
                        'сразу перед оператором стоит имя свойства. Получаем его тип
                        If mScript.mainClass(curWord.classId).Properties.TryGetValue(curWord.Word, prop) = False Then Return
                    Else
                        'сразу перед оператором стоит имя функции. Получаем ее тип
                        If mScript.mainClass(curWord.classId).Functions.TryGetValue(curWord.Word, prop) = False Then Return
                    End If
                    retType = prop.returnType
                ElseIf curWord.wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE Then
                    'перед оператором стоит ). Ищем имя свойства/функции, которому эта скобка принадлежит
                    Dim bracketBalance As Integer = -1 'текущий баланс открытых/закрытых скобок -1 (одну закрытую уже посчитали)
                    For i As Integer = currentWordId - 2 To 0 Step -1 'поиск от слова перед скобкой и до первого слова в строке
                        If codeBox.CodeData(currentLine).Code(i).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN Then
                            bracketBalance += 1
                            If i > 0 AndAlso bracketBalance = 0 Then
                                curWord = codeBox.CodeData(currentLine).Code(i - 1)
                                If curWord.wordType = EditWordTypeEnum.W_PROPERTY Then
                                    'имя свойства найдено, получаем его тип
                                    prop = mScript.mainClass(curWord.classId).Properties(curWord.Word.Trim)
                                ElseIf curWord.wordType = EditWordTypeEnum.W_FUNCTION Then
                                    'имя функции найдено, получаем ее тип
                                    prop = mScript.mainClass(curWord.classId).Functions(curWord.Word.Trim)
                                End If
                                retType = prop.returnType
                                Exit For
                            End If
                        ElseIf codeBox.CodeData(currentLine).Code(i).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE Then
                            bracketBalance -= 1
                        End If
                    Next
                Else
                    lstRtb.Hide()
                    Return
                End If
                Dim isEnum As Boolean = (IsNothing(prop.returnArray) = False AndAlso prop.returnArray.Count > 0)
                Dim propTranslateFrom As MatewScript.PropertiesInfoType = Nothing
                Dim isTranslated As Boolean = False
                If isEnum AndAlso prop.returnArray.Count = 1 Then
                    Dim enValue As String = mScript.PrepareStringToPrint(prop.returnArray(0), Nothing, False)
                    If enValue.StartsWith("[ByProperty]") Then
                        'вместо данного списка надо отобразить список из указанного свойства
                        enValue = enValue.Substring("[ByProperty]".Length) 'теперь здесь только имя класса и свойства (напр., O.EquipType)
                        Dim arrVal() As String = enValue.Split("."c) '0 - класс, 1 - свойство
                        If IsNothing(arrVal) = False AndAlso arrVal.Count = 2 Then
                            Dim cId As Integer = -1, cProp As String = ""
                            If mScript.mainClassHash.TryGetValue(arrVal(0), cId) Then
                                mScript.mainClass(cId).Properties.TryGetValue(arrVal(1), propTranslateFrom)
                                If IsNothing(propTranslateFrom.returnArray) OrElse propTranslateFrom.returnArray.Count = 0 Then
                                    propTranslateFrom = Nothing
                                Else
                                    isTranslated = True
                                End If
                            End If
                        End If
                    End If
                End If
                Select Case retType
                    Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                        'возвращаемое значение - Да/Нет. Составляем список из этих 2 слов
                        PrepareListWithRange({"True", "False"}, lastRange)
                    Case MatewScript.ReturnFunctionEnum.RETURN_ENUM
                        PrepareListWithRange(prop.returnArray, lastRange)
                        If Not isTranslated Then
                            PrepareListWithRange(prop.returnArray, lastRange)
                        Else
                            PrepareListWithRange(propTranslateFrom.returnArray, lastRange)
                        End If
                    Case MatewScript.ReturnFunctionEnum.RETURN_ELEMENT
                        'тип свойства - элемент 2 уровня. Составляем его
                        If Not isEnum Then lstRtb.Hide() : Return
                        If prop.returnArray(0) = "Variable" Then
                            'составляем список переменных    
                            PrepareListWithRange(MakeVariablesList(codeBox), lastRange)
                        Else
                            PrepareListWithRange(cListManager.FillListByClassName(prop.returnArray(0)), lastRange)
                            PrepareListWithRange(MakeVariablesList(codeBox), lastRange, True)
                        End If
                    Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO
                        PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.AUDIO), lastRange)
                    Case MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS
                        PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.CSS), lastRange)
                    Case MatewScript.ReturnFunctionEnum.RETURN_PATH_JS
                        PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.JS), lastRange)
                    Case MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE
                        PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.PICTURES), lastRange)
                    Case MatewScript.ReturnFunctionEnum.RETURN_PATH_TEXT
                        PrepareListWithRange(cListManager.GetFilesList(cListManagerClass.fileTypesEnum.TEXT), lastRange)
                    Case MatewScript.ReturnFunctionEnum.RETURN_COLOR
                        PrepareListWithRange(cListManager.GetSelectedColorsList, lastRange)
                    Case MatewScript.ReturnFunctionEnum.RETURN_FUNCTION
                        PrepareListWithRange(cListManager.GetFunctionsList, lastRange)
                    Case Else
                        If isEnum Then
                            'тип параметра не предполагает список, однако он был создан Писателем. Составляем этот список
                            If Not isTranslated Then
                                PrepareListWithRange(prop.returnArray, lastRange)
                            Else
                                PrepareListWithRange(propTranslateFrom.returnArray, lastRange)
                            End If
                        Else
                            'возвращаемое значение другое - выход
                            lstRtb.Hide()
                            Return
                        End If
                End Select
                'сохраняем последние линию и индекс слова 
                prevLine = currentLine
                prevWordId = currentWordId
            Case EditWordTypeEnum.W_POINT 'поставлена точка
                prevLine = -1
                If curWord.classId < -1 Then 'неизвестный класс - выход 
                    lstRtb.Hide()
                ElseIf curWord.classId = -1 Then
                    'переменная (т. е. введено V. или Var. )
                    'составляем список переменных    
                    PrepareListWithRange(MakeVariablesList(codeBox), lastRange)
                    lstRtb.Tag = "Variables "
                Else
                    'составляем список функций и свойств класса
                    Dim lst As New List(Of String)
                    Dim pList As SortedList(Of String, MatewScript.PropertiesInfoType) = mScript.mainClass(curWord.classId).Functions
                    If IsNothing(pList) = False AndAlso pList.Count > 0 Then
                        For i As Integer = 0 To pList.Count - 1
                            Dim pValue As MatewScript.PropertiesInfoType = pList.ElementAt(i).Value
                            If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_CODE Then Continue For
                            lst.Add(pList.ElementAt(i).Key)
                        Next i
                    End If
                    pList = mScript.mainClass(curWord.classId).Properties
                    If IsNothing(pList) = False AndAlso pList.Count > 0 Then
                        For i As Integer = 0 To pList.Count - 1
                            Dim pValue As MatewScript.PropertiesInfoType = pList.ElementAt(i).Value
                            If pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pValue.Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_CODE Then Continue For
                            lst.Add(pList.ElementAt(i).Key)
                        Next i
                    End If
                    'lst.AddRange(mScript.mainClass(curWord.classId).Functions.Keys)
                    'lst.AddRange(mScript.mainClass(curWord.classId).Properties.Keys)
                    PrepareListWithRange(lst.ToArray, lastRange)
                    lstRtb.Tag = "FuncAndProps " + curWord.classId.ToString
                End If
                'сохраняем последние линию и индекс слова 
                prevLine = currentLine
                prevWordId = currentWordId
            Case Else
                'введено нечто другое
                Dim lastChar As String = codeBox.Text.Substring(codeBox.SelectionStart - 1).Trim
                If prevLine = currentLine AndAlso prevWordId = currentWordId - 1 AndAlso lastChar.Length > 0 Then
                    If Array.IndexOf(codeBox.wordBoundsArray, lastChar.First) > -1 Then
                        If mScript.IsQuotesCountEven(curWord.Word) Then Return
                    End If
                    'Вводится тоже слово, что было при предыдущем создании списка lastRange
                    If IsNothing(lastRange) Then
                        lstRtb.Hide()
                        Return
                    End If
                    If codeBox.SelectionStart > 0 AndAlso codeBox.Text.Substring(codeBox.SelectionStart - 1, 1) = " " Then
                        'не выводим список, если перед кареткой не стоит пробел
                        lstRtb.Hide()
                        Return
                    End If
                    'создаем новый список, куда войдут все подходящие значения из списка lastRange
                    Dim newRange() As String = {}, newRangeUBound As Integer = -1
                    Dim equalId As Integer = 0 'идентификатор слова, идентичного введеному в rtb. Если такого нет, выделяется первое слово (с индексом 0)
                    For i As Integer = 0 To lastRange.GetUpperBound(0)
                        Dim blnFound As Boolean = False
                        Dim tt As String = frmMainEditor.TranslitText(chkWord)
                        If lastRange(i).ToString.IndexOf(chkWord, StringComparison.CurrentCultureIgnoreCase) > -1 Then
                            blnFound = True
                        ElseIf questEnvironment.UseTranslit AndAlso lastRange(i).ToString.IndexOf(tt, StringComparison.CurrentCultureIgnoreCase) > -1 Then
                            blnFound = True
                        End If
                        If blnFound Then
                            'найдено (напр., введено в rtb "gbo", найдено "MsgBox")
                            newRangeUBound += 1
                            Array.Resize(Of String)(newRange, newRangeUBound + 1) 'расширяем массив
                            newRange(newRangeUBound) = lastRange(i) 'и добавляем подходящее слово
                            'если слово идентично введенному - выдялим его
                            If String.Equals(chkWord, lastRange(i), StringComparison.CurrentCultureIgnoreCase) Then equalId = newRangeUBound
                        End If
                    Next
                    'вносим список в lstRtb
                    lstRtb.Items.Clear()
                    lstRtb.Items.AddRange(newRange)
                    If lstRtb.Items.Count = 0 Then
                        lstRtb.Hide()
                    Else
                        lstRtb.Show()
                        lstRtb.SelectedIndex = equalId
                        lstRtb_SelectedIndexChanged(lstRtb, New EventArgs)
                    End If
                    Return
                Else 'вводится слово, которое мы не обрабатывали - выход
                    If keyHideList Then
                        prevLine = -1
                        lstRtb.Hide()
                        Return
                    End If
                End If
        End Select

        'эта часть процедуры выполняется после создания нового списка и сразу же его отсортировывает в том случае, если за кареткой уже стоит какое-то слово (часть слова)
        If currentWordId < codeBox.CodeData(currentLine).Code.GetUpperBound(0) AndAlso codeBox.CodeData(currentLine).Code(currentWordId + 1).classId = curWord.classId Then
            curWord = codeBox.CodeData(currentLine).Code(currentWordId + 1) 'сохраняем в curWord слово за текущим
            If codeBox.CodeData(currentLine).Code(currentWordId).wordType = EditWordTypeEnum.W_POINT AndAlso curWord.wordType <> EditWordTypeEnum.W_VARIABLE AndAlso curWord.wordType <> EditWordTypeEnum.W_PROPERTY AndAlso curWord.wordType <> EditWordTypeEnum.W_FUNCTION Then
                Return
            End If
            'вводится тоже слово, что было при предыдущем создании списка lastRange
            If IsNothing(lastRange) Then
                Return
            End If
            If (curWord.wordType = EditWordTypeEnum.W_SIMPLE_BOOL OrElse curWord.wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curWord.wordType = EditWordTypeEnum.W_SIMPLE_STRING OrElse curWord.wordType = EditWordTypeEnum.W_VARIABLE) = False Then
                'если следующее слово не Да/Нет, не число, переменная и не строка, то это слово к нашему не относится. Выход без отсортировывания
                Return
            End If
            'создаем новый список, куда войдут все подходящие значения из списка lastRange
            Dim newRange() As String = {}, newRangeUBound As Integer = -1
            Dim equalId As Integer = 0 'идентификатор слова, идентичного введеному в rtb. Если такого нет, выделяется первое слово (с индексом 0)
            For i As Integer = 0 To lastRange.GetUpperBound(0)
                Dim blnFound As Boolean = False
                If lastRange(i).ToString.IndexOf(chkWord, StringComparison.CurrentCultureIgnoreCase) > -1 Then
                    blnFound = True
                ElseIf questEnvironment.UseTranslit AndAlso lastRange(i).ToString.IndexOf(frmMainEditor.TranslitText(chkWord), StringComparison.CurrentCultureIgnoreCase) > -1 Then
                    blnFound = True
                End If
                If blnFound Then
                    'найдено (напр., введено в rtb "gbo", найдено "MsgBox")
                    newRangeUBound += 1
                    Array.Resize(Of String)(newRange, newRangeUBound + 1) 'расширяем массив
                    newRange(newRangeUBound) = lastRange(i) 'и добавляем подходящее слово
                    'если слово идентично введенному - выдялим его
                    If String.Equals(chkWord, lastRange(i), StringComparison.CurrentCultureIgnoreCase) Then equalId = newRangeUBound
                End If
            Next
            'вносим список в lstRtb
            lstRtb.Items.Clear()
            lstRtb.Items.AddRange(newRange)
            If lstRtb.Items.Count = 0 Then
                lstRtb.Hide()
            Else
                lstRtb.Show()
                lstRtb.SelectedIndex = equalId
                lstRtb_SelectedIndexChanged(lstRtb, New EventArgs)
            End If
        ElseIf currentWordId > 0 AndAlso codeBox.CodeData(currentLine).Code(currentWordId - 1).wordType = EditWordTypeEnum.W_POINT AndAlso _
            (curWord.wordType = EditWordTypeEnum.W_FUNCTION OrElse curWord.wordType = EditWordTypeEnum.W_PROPERTY) Then
            Dim lastChar As Char = codeBox.Text.Chars(codeBox.SelectionStart - 1)
            If lastChar = " "c OrElse lastChar = "("c Then Return
            'вводится тоже слово, что было при предыдущем создании списка lastRange
            If IsNothing(lastRange) Then
                Return
            End If
            'создаем новый список, куда войдут все подходящие значения из списка lastRange
            Dim newRange() As String = {}, newRangeUBound As Integer = -1
            Dim equalId As Integer = 0 'идентификатор слова, идентичного введеному в rtb. Если такого нет, выделяется первое слово (с индексом 0)
            For i As Integer = 0 To lastRange.GetUpperBound(0)
                Dim blnFound As Boolean = False
                If lastRange(i).ToString.IndexOf(chkWord, StringComparison.CurrentCultureIgnoreCase) > -1 Then
                    blnFound = True
                ElseIf questEnvironment.UseTranslit AndAlso lastRange(i).ToString.IndexOf(frmMainEditor.TranslitText(chkWord), StringComparison.CurrentCultureIgnoreCase) > -1 Then
                    blnFound = True
                End If
                If blnFound Then
                    'найдено (напр., введено в rtb "gbo", найдено "MsgBox")
                    newRangeUBound += 1
                    Array.Resize(Of String)(newRange, newRangeUBound + 1) 'расширяем массив
                    newRange(newRangeUBound) = lastRange(i) 'и добавляем подходящее слово
                    'если слово идентично введенному - выдялим его
                    If String.Equals(chkWord, lastRange(i), StringComparison.CurrentCultureIgnoreCase) Then equalId = newRangeUBound
                End If
            Next
            'вносим список в lstRtb
            lstRtb.Items.Clear()
            lstRtb.Items.AddRange(newRange)
            If lstRtb.Items.Count < 1 Then
                lstRtb.Hide()
            Else
                If lstRtb.Items.Count = 1 AndAlso String.Compare(chkWord, lstRtb.Items(0), True) = 0 Then
                    lstRtb.Hide()
                Else
                    lstRtb.Show()
                    lstRtb.SelectedIndex = equalId
                    lstRtb_SelectedIndexChanged(lstRtb, New EventArgs)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Получает параметры блока Event. При этом возвращается их содержимое, но только в том случае, если
    ''' они являются простой строкой или числом (то есть, именем или индексом элемента)
    ''' </summary>
    ''' <param name="currentLine">Текущая линия в кодбоксе</param>
    ''' <returns>Содержимое параметра, но только если оно является простой строкой или числом (то есть, именем или индексом элемента). Иначе - пустая строка</returns>
    Private Function GetEventParamValues(ByVal currentLine As Integer, ByRef classId As Integer, ByRef propName As String) As String()
        Dim curCD As CodeDataType = codeBox.CodeData(currentLine)
        Dim lstContent As New List(Of String) 'В этот массив получаем список всех параметров свойства
        Dim codeLen As Integer = curCD.Code.Count
        classId = -1
        If codeLen >= 3 AndAlso curCD.Code(1).wordType = EditWordTypeEnum.W_SIMPLE_STRING AndAlso curCD.Code(2).wordType = EditWordTypeEnum.W_COMMA Then
            'Event '[ClassName].EventName',...
            'Плоучаем класс и имя собтытия
            If Not GetClassIdAndElementNameByString(curCD.Code(1).Word, classId, propName, False, True) Then Return lstContent.ToArray

            If codeLen >= 5 AndAlso (curCD.Code(3).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(3).wordType = EditWordTypeEnum.W_SIMPLE_STRING) AndAlso _
                curCD.Code(4).wordType = EditWordTypeEnum.W_COMMA Then
                'Event '[ClassName].ventName', element1_Name/Id, ...
                lstContent.Add(curCD.Code(3).Word.Trim) 'вставляем первый параметр - элемент 2 уровня
                If mScript.mainClass(classId).LevelsCount = 2 AndAlso codeLen >= 6 AndAlso (curCD.Code(5).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse _
                                                                                          curCD.Code(5).wordType = EditWordTypeEnum.W_SIMPLE_STRING) Then
                    lstContent.Add(curCD.Code(5).Word.Trim) 'вставляем второй параметр - элемент 3 уровня
                End If
                Return lstContent.ToArray
            End If
        End If
        Return lstContent.ToArray
    End Function

    ''' <summary>
    ''' Получает параметры свойства. При этом возвращается их содержимое, но только в том случае, если
    ''' они являются простой строкой или числом (то есть, именем или индексом элемента)
    ''' </summary>
    ''' <param name="currentLine">Текущая линия в кодбоксе</param>
    ''' <param name="propWordId">Id слова в codeBox.CodeData текущей строки, которая является именем свойства, с параметрами которого работаем</param>
    ''' <returns>Содержимое параметра, но только если оно является простой строкой или числом (то есть, именем или индексом элемента). Иначе - пустая строка</returns>
    Private Function GetPropertyParamValues(ByVal currentLine As Integer, ByVal propWordId As Integer) As String()
        Dim curCD As CodeDataType = codeBox.CodeData(currentLine)
        Dim parCount As Integer = 0 'хранит количество полученных параметров
        Dim qbCount As Integer = 0, obCount As Integer = 0 'для хранения разности открытых/закрытых скобок
        Dim lstContent As New List(Of String) 'В этот массив получаем список всех параметров свойства
        If propWordId > 4 AndAlso curCD.Code(propWordId - 1).wordType = EditWordTypeEnum.W_POINT AndAlso curCD.Code(propWordId - 2).wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
            'Имеются начальные скобки []. Возможно Class[param1, ...].PropName

            If curCD.Code(propWordId - 3).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(propWordId - 3).wordType = EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                (propWordId >= 4 AndAlso (curCD.Code(propWordId - 4).wordType = EditWordTypeEnum.W_COMMA OrElse curCD.Code(propWordId - 4).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN)) Then
                'слово перед ] является простой строкой или числом, при этом перед этим словом стоит , или [

                lstContent.Add(curCD.Code(propWordId - 3).Word.Trim)
            Else
                'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                lstContent.Add("")
            End If
            parCount += 1
            'цикл идет от ] в сторону начала строки и начинается внутри [] свойства
            qbCount = 1 '(то есть, первый символ ] уже получен)

            For wId As Integer = propWordId - 3 To 1 Step -1
                Select Case curCD.Code(wId).wordType
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount -= 1
                        If qbCount = 0 Then Exit For 'дошли до открывающей [
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount += 1
                    Case EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        obCount += 1
                    Case EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                        obCount -= 1
                    Case EditWordTypeEnum.W_COMMA
                        'найдена запятая
                        If obCount <> 0 OrElse qbCount <> 1 Then Continue For 'она стоит вне данного свойства (является разделителем внутри вложенной функции или свойства)
                        'данная запятая - разделитель параметров этого свойства. Получаем параметр сразу перед запятой
                        parCount += 1
                        If curCD.Code(wId - 1).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(wId - 1).wordType = EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                            (wId >= 2 AndAlso (curCD.Code(wId - 2).wordType = EditWordTypeEnum.W_COMMA OrElse curCD.Code(wId - 2).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN)) Then
                            'слово перед запятой является простой строкой или числом, при этом перед этим словом стоит , или [
                            lstContent.Insert(0, curCD.Code(wId - 1).Word.Trim) 'вставляем вновь полученный параметр на первое место в lstContent, так как идем от конца к началу
                        Else
                            'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                            lstContent.Insert(0, "")
                        End If
                End Select
            Next wId

            If lstContent.Count > 1 Then
                'Список из 2 параметров (максимум у свойства уже составлен). Возвращаем его
                Return lstContent.ToArray
            End If
        End If

        Dim codeLen As Integer = curCD.Code.Count
        If codeLen >= propWordId + 3 AndAlso curCD.Code(propWordId + 1).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN Then
            '...propName(paramX, ...
            obCount = 1 'то есть, первый символ ( уже получен
            qbCount = 0
            If curCD.Code(propWordId + 2).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(propWordId + 2).wordType = EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                (propWordId <= codeLen - 4 AndAlso (curCD.Code(propWordId + 3).wordType = EditWordTypeEnum.W_COMMA OrElse curCD.Code(propWordId + 3).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE)) Then
                'слово после ( является простой строкой или числом, при этом после этого слова стоит , или )
                lstContent.Add(curCD.Code(propWordId + 2).Word.Trim)
            Else
                'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                lstContent.Add("")
            End If
            parCount += 1
            'цикл идет от ( в сторону конца строки и начинается внутри () свойства
            For wId As Integer = propWordId + 2 To codeLen - 1
                Select Case curCD.Code(wId).wordType
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount += 1
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount -= 1
                    Case EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        obCount -= 1
                        If obCount = 0 Then Exit For 'дошли до закрывающей (
                    Case EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                        obCount += 1
                    Case EditWordTypeEnum.W_COMMA
                        'найдена запятая
                        If obCount <> 0 OrElse qbCount <> 1 Then Continue For 'она стоит вне данного свойства (является разделителем внутри вложенной функции или свойства)
                        'данная запятая - разделитель параметров этом свойстве. Получаем параметр сразу перед запятой

                        If wId = codeLen - 1 Then Exit For
                        parCount += 1
                        If curCD.Code(wId + 1).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(wId + 1).wordType = EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                            (wId <= codeLen - 3 AndAlso (curCD.Code(wId + 2).wordType = EditWordTypeEnum.W_COMMA OrElse curCD.Code(wId + 2).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE)) Then
                            'слово после запятой является простой строкой или числом, при этом после этого слова стоит , или )
                            lstContent.Add(curCD.Code(wId - 1).Word.Trim)
                        Else
                            'содержимое не является простым числом или строкой, на этапе написания кода просчитать итоговое значение содержимого невозможно.
                            lstContent.Add("")
                        End If
                End Select
            Next wId
            'На этом этапе получены в списке lstContent содержатся все параметры свойства (в правильном порядке). Если содержимое параметра - простое число/строка, то оно попало 
            'в данный список lstContent. Если же содержимое иное (не просчитывается на этапе написания кода), то вместо содержимого - пустая строка.
            Return lstContent.ToArray
        End If
        Return lstContent.ToArray
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
    Private Function GetFunctionParamWhichTypeIsElement(ByVal currentLine As Integer, ByVal funcWordId As Integer, ByRef params() As MatewScript.paramsType, _
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

        Dim curCD As CodeDataType = codeBox.CodeData(currentLine)
        Dim parCount As Integer = 0 'хранит количество полученных параметров
        Dim qbCount As Integer = 0, obCount As Integer = 0 'для хранения разности открытых/закрытых скобок
        Dim lstContent As New List(Of String) 'В этот массив получаем список всех параметров функции
        If funcWordId > 4 AndAlso curCD.Code(funcWordId - 1).wordType = EditWordTypeEnum.W_POINT AndAlso curCD.Code(funcWordId - 2).wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
            'Имеются начальные скобки []. Возможно Class[param1, ...].FuncName
            If curCD.Code(funcWordId - 3).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(funcWordId - 3).wordType = EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                (funcWordId >= 4 AndAlso (curCD.Code(funcWordId - 4).wordType = EditWordTypeEnum.W_COMMA OrElse curCD.Code(funcWordId - 4).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN)) Then
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
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount -= 1
                        If qbCount = 0 Then Exit For 'дошли до открывающей [
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount += 1
                    Case EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        obCount += 1
                    Case EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                        obCount -= 1
                    Case EditWordTypeEnum.W_COMMA
                        'найдена запятая
                        If obCount <> 0 OrElse qbCount <> 1 Then Continue For 'она стоит вне данной функции (является разделителем внутри вложенной функции или свойства)
                        'данная запятая - разделитель параметров этой функции. Получаем параметр сразу перед запятой
                        parCount += 1
                        If curCD.Code(wId - 1).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(wId - 1).wordType = EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                            (wId >= 2 AndAlso (curCD.Code(wId - 2).wordType = EditWordTypeEnum.W_COMMA OrElse curCD.Code(wId - 2).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN)) Then
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
        If codeLen >= funcWordId + 3 AndAlso curCD.Code(funcWordId + 1).wordType = EditWordTypeEnum.W_OVAL_BRACKET_OPEN Then
            '...funcName(paramX, ...
            obCount = 1 'то есть, первый символ ( уже получен
            qbCount = 0
            If curCD.Code(funcWordId + 2).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(funcWordId + 2).wordType = EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                (funcWordId <= codeLen - 4 AndAlso (curCD.Code(funcWordId + 3).wordType = EditWordTypeEnum.W_COMMA OrElse curCD.Code(funcWordId + 3).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE)) Then
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
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        qbCount += 1
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        qbCount -= 1
                    Case EditWordTypeEnum.W_OVAL_BRACKET_CLOSE
                        obCount -= 1
                        If obCount = 0 Then Exit For 'дошли до закрывающей (
                    Case EditWordTypeEnum.W_OVAL_BRACKET_OPEN
                        obCount += 1
                    Case EditWordTypeEnum.W_COMMA
                        'найдена запятая
                        If obCount <> 0 OrElse qbCount <> 1 Then Continue For 'она стоит вне данной функции (является разделителем внутри вложенной функции или свойства)
                        'данная запятая - разделитель параметров этой функции. Получаем параметр сразу перед запятой

                        If wId = codeLen - 1 Then Exit For
                        parCount += 1
                        If curCD.Code(wId + 1).wordType = EditWordTypeEnum.W_SIMPLE_NUMBER OrElse curCD.Code(wId + 1).wordType = EditWordTypeEnum.W_SIMPLE_STRING AndAlso _
                            (wId <= codeLen - 3 AndAlso (curCD.Code(wId + 2).wordType = EditWordTypeEnum.W_COMMA OrElse curCD.Code(wId + 2).wordType = EditWordTypeEnum.W_OVAL_BRACKET_CLOSE)) Then
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
    ''' Процедура находит слово функции или свойства, в которой находится каретка, а также номер вводимого параметра
    ''' </summary>
    ''' <param name="currentLine">текущая линия</param>
    ''' <param name="currentWordId">Id текущего слова в линии</param>
    ''' <param name="paramNumber">для получения номера параметра, внутри которого каретка</param>
    ''' <param name="elementWord">для получения слова функции/свойства</param>
    ''' <param name="classId">для получения Id класса функции/свойства</param>
    ''' <param name="elementWordId">Id слова в строке, которая является функцией/свойством</param>
    ''' <returns></returns>
    Private Function GetCurrentElementWord(ByVal currentLine As Integer, ByVal currentWordId As Integer, ByRef paramNumber As Integer, ByRef elementWord As EditWordType, _
                                           ByRef classId As Integer, Optional ByRef elementWordId As Integer = -1, Optional ByRef elementWordLine As Integer = -1) As Boolean
        If IsNothing(mScript.mainClass) Then Return False
        'class[...,X,...].funcName...
        '...funcName(...,X,...

        If currentWordId < 0 Then Return False
        If IsNothing(codeBox.CodeData(currentLine).Code) Then
            If currentLine = 0 OrElse IsNothing(codeBox.CodeData(currentLine - 1).Code) OrElse codeBox.CodeData(currentLine - 1).Code.Count = 0 OrElse _
                codeBox.CodeData(currentLine - 1).Code.Last.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then Return False
            currentLine -= 1
            currentWordId = codeBox.CodeData(currentLine).Code.Count - 1
        End If
        If codeBox.CodeData(currentLine).Code.GetUpperBound(0) < currentWordId Then Return False
        elementWord = New EditWordType 'очищаем слово
        Dim quadBracketBalance As Integer = 0 'баланс [ ]
        Dim ovalBracketBalance As Integer = 0 'баланс 
        Dim paramIfInQuad As Integer = 1 'номер параметра, если мы внутри []
        Dim paramIfInOval As Integer = 1 'номер параметра, если мы внутри () или скобок нет вообще
        classId = -1 'для получения класса элемента (т. е. функции или свойства)
        'Dim wordId As Integer = -1 'Id слова с элементом
        'сначала проходим от текущего слова до первого, определяя внутри каких скобок мы находимся - () или []
        Dim prevLine As Integer = currentLine, initLine As Integer = currentLine
        Dim i As Integer = currentWordId

        Dim lCopy As Integer = currentLine, wCopy As Integer = i
        Dim wrd As CodeTextBox.EditWordType = codeBox.CodeData(currentLine).Code(i)

        If wrd.wordType = EditWordTypeEnum.W_FUNCTION OrElse wrd.wordType = EditWordTypeEnum.W_PROPERTY Then
            'мы стоим прямо на слове
            paramNumber = 1
            elementWord = wrd
            classId = wrd.classId
            elementWordId = currentWordId
            elementWordLine = currentLine
            'Return True
        End If

        Dim wrdPrev As CodeTextBox.EditWordType = codeBox.GetPrevWord(lCopy, wCopy, True)
        Dim wrdNext As CodeTextBox.EditWordType = codeBox.GetNextWord(lCopy, wCopy, True)

        If wrd.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
            wrdNext = wrd
2:
            wrd = codeBox.GetPrevWord(currentLine, i, False)
            If IsNothing(wrd.Word) = False AndAlso prevLine <> currentLine Then
                'перешли к строке выше
                If wrd.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                    prevLine = currentLine
                    GoTo 2
                End If
            End If
            lCopy = currentLine
            wCopy = i
            wrdPrev = codeBox.GetPrevWord(lCopy, wCopy, True)
        End If

        'For i As Integer = currentWordId To 0 Step -1
        Do
            Select Case wrd.wordType
                Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN '[
                    quadBracketBalance += 1
                    If quadBracketBalance = 1 Then
                        'мы внутри [..X..]
                        If wrdPrev.wordType = EditWordTypeEnum.W_NOTHING Then Return False 'перед [ ничего не стоит - выход
                        classId = wrdPrev.classId
                        paramNumber = paramIfInQuad 'получаем номер параметра, внутри которого мы находимся)
                        If classId = -1 Then Return False
                        Exit Do
                    End If
                Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE ']
                    quadBracketBalance -= 1
                Case EditWordTypeEnum.W_FUNCTION
                    If codeBox.CodeData(currentLine).Code.GetUpperBound(0) = currentWordId OrElse (wrdNext.wordType <> EditWordTypeEnum.W_QUAD_BRACKET_OPEN AndAlso _
                                                                                           quadBracketBalance = 0) Then
                        'мы внутри параметров функции (за ее именем), у которой нет круглых скобок
                        lCopy = currentLine
                        wCopy = i
                        Dim wrdPrev2 As CodeTextBox.EditWordType = codeBox.GetPrevWord(lCopy, wCopy, True) 'i - 2
                        wrdPrev2 = codeBox.GetPrevWord(lCopy, wCopy, True)

                        If wrdPrev.wordType <> EditWordTypeEnum.W_NOTHING Then
                            If wrdPrev.wordType <> EditWordTypeEnum.W_POINT Then
                                If wrdPrev.wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso wrdPrev.Word.Trim <> "Then" Then
                                    'Continue Do
                                    GoTo cycleEnd
                                End If
                            ElseIf (wrdPrev.wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER OrElse wrdPrev.wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) _
                                AndAlso wrdPrev.wordType <> EditWordTypeEnum.W_POINT Then

                                If wrdPrev.wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso wrdPrev.Word.Trim <> "Then" Then
                                    'Continue Do
                                    GoTo cycleEnd
                                End If
                            ElseIf IsNothing(wrdPrev2.Word) = False AndAlso wrdPrev2.wordType = EditWordTypeEnum.W_CLASS Then
                                Dim wrdPrev3 As CodeTextBox.EditWordType = codeBox.GetPrevWord(lCopy, wCopy, True) 'i - 3

                                If IsNothing(wrdPrev3.Word) = False Then
                                    If wrdPrev3.wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso wrdPrev3.Word.Trim <> "Then" Then
                                        'Continue Do
                                        GoTo cycleEnd
                                    End If
                                End If
                            ElseIf IsNothing(wrdPrev2.Word) = False AndAlso wrdPrev2.wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                'Class[...].Name
                                Dim internalQbalance As Integer = -1
                                Dim classPos As Integer = -1
                                Dim j As Integer = wCopy
                                Dim prevJ As Integer = j
                                Dim wrdPrev3 As CodeTextBox.EditWordType = codeBox.GetPrevWord(lCopy, j, True) 'i - 3
                                Do
                                    If wrdPrev3.wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                                        internalQbalance += 1
                                        If internalQbalance = 0 Then
                                            classPos = prevJ
                                        End If
                                    ElseIf wrdPrev3.wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                        internalQbalance -= 1
                                    End If

                                    prevJ = j
                                    wrdPrev3 = codeBox.GetPrevWord(lCopy, j, True)
                                    If wrdPrev3.wordType = EditWordTypeEnum.W_NOTHING Then Exit Do
                                Loop

                                wCopy = classPos
                                wrdPrev3 = codeBox.GetPrevWord(lCopy, wCopy, True)
                                If IsNothing(wrdPrev3.Word) = False AndAlso codeBox.CodeData(lCopy).Code(classPos).wordType = EditWordTypeEnum.W_CLASS Then
                                    If wrdPrev3.wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso wrdPrev3.Word.Trim <> "Then" Then
                                        'Continue Do
                                        GoTo cycleEnd
                                    End If
                                End If
                            End If
                        End If
                        If codeBox.CodeData(currentLine).Code.GetUpperBound(0) = currentWordId AndAlso currentLine = initLine AndAlso i = currentWordId AndAlso codeBox.SelectionStart > 0 AndAlso (codeBox.Text.Chars(codeBox.SelectionStart - 1) <> " "c) Then
                            'если мы сразу за именем функции, то выходим, возвращая False (т.к. имя функции, возможно, еще редактируется)
                            Return False
                        End If
                        'мы за функцией, не имеющей круглых скобок
                        classId = wrd.classId
                        paramNumber = paramIfInOval 'получаем номер параметра, без учета возможных параметров в квадратных скобках
                        If classId = -1 Then Return False 'это переменная - выход (на всякий случай, хоть это и синтаксическая ошибка)
                        If wrd.wordType = EditWordTypeEnum.W_PROPERTY OrElse wrd.wordType = EditWordTypeEnum.W_FUNCTION Then
                            elementWord = wrd  'получаем слово элемента и его Id
                            elementWordId = i
                            elementWordLine = currentLine
                        Else
                            Return False
                        End If
                        Exit Do
                    End If
                Case EditWordTypeEnum.W_OVAL_BRACKET_OPEN '(
                    ovalBracketBalance += 1
                    If ovalBracketBalance = 1 Then
                        'мы внутри (..Х..)
                        If wrdPrev.wordType = EditWordTypeEnum.W_NOTHING Then Return False
                        classId = wrdPrev.classId
                        paramNumber = paramIfInOval 'получаем номер параметра, без учета возможных параметров в квадратных скобках
                        If classId = -1 Then Return False 'переменная - выход
                        If wrdPrev.wordType = EditWordTypeEnum.W_PROPERTY OrElse wrdPrev.wordType = EditWordTypeEnum.W_FUNCTION Then
                            elementWord = wrdPrev  'получаем слово элемента и его Id
                            lCopy = currentLine
                            wCopy = i
                            Dim wp As CodeTextBox.EditWordType = codeBox.GetPrevWord(lCopy, wCopy, True)
                            elementWordId = wCopy
                            elementWordLine = lCopy
                        Else
                            Return False
                        End If
                        Exit Do
                    End If
                Case EditWordTypeEnum.W_OVAL_BRACKET_CLOSE ')
                    ovalBracketBalance -= 1
                Case EditWordTypeEnum.W_COMMA
                    'пока мы не знаем, внутри каких скобок находимся, считаем и как для круглых, и как для квадратных
                    'If quadBracketBalance = 0 Then paramIfInQuad += 1
                    'If ovalBracketBalance = 0 Then paramIfInOval += 1
                    If ovalBracketBalance = 0 AndAlso quadBracketBalance = 0 Then
                        paramIfInQuad += 1
                        paramIfInOval += 1
                    End If
            End Select
            'Next
cycleEnd:
            wrdNext = wrd
1:
            wrd = codeBox.GetPrevWord(currentLine, i, False)
            If IsNothing(wrd.Word) Then Exit Do
            If prevLine <> currentLine Then
                'перешли к строке выше
                If wrd.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                    Exit Do
                End If
                prevLine = currentLine
                GoTo 1
            End If
            lCopy = currentLine
            wCopy = i
            wrdPrev = codeBox.GetPrevWord(lCopy, wCopy, True)
        Loop

        currentLine = initLine
        prevLine = initLine

        'Сейчас возможны 2 варианта:
        '1) Мы в [..X..], известен класс, известен номер параметра, не известно слово
        '2) Мы в (..X..), известен класс, не известен номер параметра, известно слово

        If elementWordId = -1 Then
            'вариант 1). Ищем закрывающую квадратную скобку функции, вслед за которой должно идти .funcName
            i = currentWordId + 1
            If i > codeBox.CodeData(currentLine).Code.Count - 1 Then Return False
            lCopy = currentLine
            prevLine = currentLine
            wCopy = i
            wrdPrev = codeBox.GetPrevWord(lCopy, wCopy, True)
            wrd = codeBox.CodeData(currentLine).Code(i)
3:
            If wrd.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                wrd = codeBox.GetNextWord(currentLine, i, False)
                If IsNothing(wrd.Word) = False AndAlso prevLine <> currentLine Then
                    'перешли к строке ниже
                    If wrd.wordType = EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                        prevLine = currentLine
                        GoTo 3
                    End If
                End If
            End If
            lCopy = currentLine
            wCopy = i
            wrdNext = codeBox.GetNextWord(lCopy, wCopy, True)

            quadBracketBalance = 0
            Do
                Select Case wrd.wordType
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        quadBracketBalance += 1
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        quadBracketBalance -= 1
                        If quadBracketBalance = -1 Then
                            'наша закрывающая ] найдена
                            lCopy = currentLine
                            wCopy = i
                            Dim wrdNext2 As CodeTextBox.EditWordType = codeBox.GetNextWord(lCopy, wCopy, True)
                            wrdNext2 = codeBox.GetNextWord(lCopy, wCopy, True) 'i+2

                            If IsNothing(wrdNext2.Word) = False Then
                                If wrdNext2.wordType = EditWordTypeEnum.W_PROPERTY OrElse wrdNext2.wordType = EditWordTypeEnum.W_FUNCTION Then
                                    'элемент найден. Получаем слово
                                    elementWord = wrdNext2
                                    Return True
                                Else
                                    Return False
                                End If
                            End If
                        End If
                End Select

                wrdPrev = wrd
                '4:
                wrd = codeBox.GetNextWord(currentLine, i, False)
                If IsNothing(wrd.Word) Then Exit Do

                If prevLine <> currentLine Then
                    'перешли к строке ниже
                    If wrdPrev.wordType <> EditWordTypeEnum.W_STRINGS_DISSOCIATION Then
                        Exit Do
                    End If
                    prevLine = currentLine
                    'GoTo 4
                End If

                lCopy = currentLine
                wCopy = i
                wrdNext = codeBox.GetNextWord(lCopy, wCopy, True)
            Loop
        Else
            'вариант 2). Не известно кол-во параметров в квадратных скобках
            i = elementWordId
            currentLine = elementWordLine
            lCopy = currentLine
            wCopy = i
            wrdPrev = codeBox.GetPrevWord(lCopy, wCopy, True)

            If wrdPrev.wordType = EditWordTypeEnum.W_NOTHING Then Return True 'слово первое, значит квадратных скобок нет - все известно
            If wrdPrev.wordType <> EditWordTypeEnum.W_POINT Then Return True 'перед именем функции нет точки - квадратных скобок нет - все известно
            'If elementWordId < 3 Then Return True 'ошибка синтаксиса
            Dim wrdPrev2 As CodeTextBox.EditWordType = codeBox.GetPrevWord(lCopy, wCopy, True) 'elementWordId - 2
            If wrdPrev2.wordType = EditWordTypeEnum.W_NOTHING Then Return True 'ошибка синтаксиса
            If wrdPrev2.wordType <> EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then Return True 'квадратных скобок нет - все известно
            quadBracketBalance = -1
            paramNumber += 1
            'проверяем все слова от слова перед закрывающей ] до начала строки в поисках открывающей [ функции
            wrd = codeBox.CodeData(currentLine).Code(elementWordId)
            'получаем elementWordId - 3 с учетом возможных _
            wrd = codeBox.GetPrevWord(lCopy, wCopy, True) 'elementWordId - 3
            currentLine = lCopy
            i = wCopy

            Do
                Select Case wrd.wordType
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        quadBracketBalance += 1
                        If quadBracketBalance = 0 Then Return True 'найдена нужная [ - выход
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        quadBracketBalance -= 1
                    Case EditWordTypeEnum.W_COMMA
                        If quadBracketBalance = -1 Then paramNumber += 1 'найдена запятая этой функции, номер параметра +1
                    Case EditWordTypeEnum.W_NOTHING
                        Exit Do
                End Select

                wrd = codeBox.GetPrevWord(currentLine, i, True)
                If IsNothing(wrd.Word) Then Exit Do
            Loop

        End If
        Return False
    End Function

    Private Function GetCurrentElementWordCopy(ByVal currentLine As Integer, ByVal currentWordId As Integer, ByRef paramNumber As Integer, ByRef elementWord As EditWordType, _
                                       ByRef classId As Integer, Optional ByRef elementWordId As Integer = -1) As Boolean
        If IsNothing(mScript.mainClass) Then Return False
        'class[...,X,...].funcName...
        '...funcName(...,X,...

        If currentWordId < 0 Then Return False
        If IsNothing(codeBox.CodeData(currentLine).Code) Then Return False
        If codeBox.CodeData(currentLine).Code.GetUpperBound(0) < currentWordId Then Return False
        elementWord = New EditWordType 'очищаем слово
        Dim quadBracketBalance As Integer = 0 'баланс [ ]
        Dim ovalBracketBalance As Integer = 0 'баланс 
        Dim paramIfInQuad As Integer = 1 'номер параметра, если мы внутри []
        Dim paramIfInOval As Integer = 1 'номер параметра, если мы внутри () или скобок нет вообще
        classId = -1 'для получения класса элемента (т. е. функции или свойства)
        'Dim wordId As Integer = -1 'Id слова с элементом
        'сначала проходим от текущего слова до первого, определяя внутри каких скобок мы находимся - () или []
        For i As Integer = currentWordId To 0 Step -1
            Select Case codeBox.CodeData(currentLine).Code(i).wordType
                Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN '[
                    quadBracketBalance += 1
                    If quadBracketBalance = 1 Then
                        'мы внутри [..X..]
                        If i = 0 Then Return False 'перед [ ничего не стоит - выход
                        classId = codeBox.CodeData(currentLine).Code(i - 1).classId
                        paramNumber = paramIfInQuad 'получаем номер параметра, внутри которого мы находимся)
                        If classId = -1 Then Return False
                        Exit For
                    End If
                Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE ']
                    quadBracketBalance -= 1
                Case EditWordTypeEnum.W_FUNCTION
                    If codeBox.CodeData(currentLine).Code.GetUpperBound(0) = currentWordId OrElse (codeBox.CodeData(currentLine).Code(i + 1).wordType <> EditWordTypeEnum.W_QUAD_BRACKET_OPEN AndAlso _
                                                                                           quadBracketBalance = 0) Then
                        'мы внутри параметров функции (за ее именем), у которой нет круглых скобок
                        If i > 0 Then
                            If codeBox.CodeData(currentLine).Code(i - 1).wordType <> EditWordTypeEnum.W_POINT Then
                                If codeBox.CodeData(currentLine).Code(i - 1).wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso codeBox.CodeData(currentLine).Code(i - 1).Word.Trim <> "Then" Then
                                    Continue For
                                End If
                            ElseIf (codeBox.CodeData(currentLine).Code(i - 1).wordType = EditWordTypeEnum.W_CONVERT_TO_NUMBER OrElse codeBox.CodeData(currentLine).Code(i - 1).wordType = EditWordTypeEnum.W_CONVERT_TO_STRING) _
                                AndAlso i > 1 AndAlso codeBox.CodeData(currentLine).Code(i - 1).wordType <> EditWordTypeEnum.W_POINT Then

                                If codeBox.CodeData(currentLine).Code(i - 1).wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso codeBox.CodeData(currentLine).Code(i - 1).Word.Trim <> "Then" Then
                                    Continue For
                                End If
                            ElseIf i > 1 AndAlso codeBox.CodeData(currentLine).Code(i - 2).wordType = EditWordTypeEnum.W_CLASS Then
                                'If i = 2 Then Continue For
                                'If codeBox.CodeData(currentLine).Code(i - 3).wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso codeBox.CodeData(currentLine).Code(i - 3).Word.Trim <> "Then" Then
                                '    Continue For
                                'End If

                                If i > 2 Then
                                    If codeBox.CodeData(currentLine).Code(i - 3).wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso codeBox.CodeData(currentLine).Code(i - 3).Word.Trim <> "Then" Then
                                        Continue For
                                    End If
                                End If
                            ElseIf i > 1 AndAlso codeBox.CodeData(currentLine).Code(i - 2).wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                'Class[...].Name
                                Dim internalQbalance As Integer = -1
                                Dim classPos As Integer = -1
                                For j = i - 3 To 0 Step -1
                                    If codeBox.CodeData(currentLine).Code(j).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN Then
                                        internalQbalance += 1
                                        If internalQbalance = 0 Then
                                            classPos = j - 1
                                        End If
                                    ElseIf codeBox.CodeData(currentLine).Code(j).wordType = EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then
                                        internalQbalance -= 1
                                    End If
                                Next
                                If classPos > 0 AndAlso codeBox.CodeData(currentLine).Code(classPos).wordType = EditWordTypeEnum.W_CLASS Then
                                    If codeBox.CodeData(currentLine).Code(classPos - 1).wordType <> EditWordTypeEnum.W_STRINGS_CONSOLIDATION AndAlso codeBox.CodeData(currentLine).Code(classPos - 1).Word.Trim <> "Then" Then
                                        Continue For
                                    End If
                                End If
                            End If
                        End If
                        If codeBox.CodeData(currentLine).Code.GetUpperBound(0) = currentWordId AndAlso i = currentWordId AndAlso codeBox.SelectionStart > 0 AndAlso (codeBox.Text.Chars(codeBox.SelectionStart - 1) <> " "c) Then
                            'если мы сразу за именем функции, то выходим, возвращая False (т.к. имя функции, возможно, еще редактируется)
                            Return False
                        End If
                        'мы за функцией, не имеющей круглых скобок
                        classId = codeBox.CodeData(currentLine).Code(i).classId
                        paramNumber = paramIfInOval 'получаем номер параметра, без учета возможных параметров в квадратных скобках
                        If classId = -1 Then Return False 'это переменная - выход (на всякий случай, хоть это и синтаксическая ошибка)
                        If codeBox.CodeData(currentLine).Code(i).wordType = EditWordTypeEnum.W_PROPERTY OrElse codeBox.CodeData(currentLine).Code(i).wordType = EditWordTypeEnum.W_FUNCTION Then
                            elementWord = codeBox.CodeData(currentLine).Code(i) 'получаем слово элемента и его Id
                            elementWordId = i
                        Else
                            Return False
                        End If
                        Exit For
                    End If
                Case EditWordTypeEnum.W_OVAL_BRACKET_OPEN '(
                    ovalBracketBalance += 1
                    If ovalBracketBalance = 1 Then
                        'мы внутри (..Х..)
                        If i = 0 Then Return False
                        classId = codeBox.CodeData(currentLine).Code(i - 1).classId
                        paramNumber = paramIfInOval 'получаем номер параметра, без учета возможных параметров в квадратных скобках
                        If classId = -1 Then Return False 'переменная - выход
                        If codeBox.CodeData(currentLine).Code(i - 1).wordType = EditWordTypeEnum.W_PROPERTY OrElse codeBox.CodeData(currentLine).Code(i - 1).wordType = EditWordTypeEnum.W_FUNCTION Then
                            elementWord = codeBox.CodeData(currentLine).Code(i - 1) 'получаем слово элемента и его Id
                            elementWordId = i - 1
                        Else
                            Return False
                        End If
                        Exit For
                    End If
                Case EditWordTypeEnum.W_OVAL_BRACKET_CLOSE ')
                    ovalBracketBalance -= 1
                Case EditWordTypeEnum.W_COMMA
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
            For i As Integer = currentWordId To codeBox.CodeData(currentLine).Code.GetUpperBound(0)
                Select Case codeBox.CodeData(currentLine).Code(i).wordType
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        quadBracketBalance += 1
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        quadBracketBalance -= 1
                        If quadBracketBalance = -1 Then
                            'наша закрывающая ] найдена
                            If i + 1 < codeBox.CodeData(currentLine).Code.GetUpperBound(0) Then
                                If codeBox.CodeData(currentLine).Code(i + 2).wordType = EditWordTypeEnum.W_PROPERTY OrElse codeBox.CodeData(currentLine).Code(i + 2).wordType = EditWordTypeEnum.W_FUNCTION Then
                                    'элемент найден. Получаем слово
                                    elementWord = codeBox.CodeData(currentLine).Code(i + 2)
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
            If codeBox.CodeData(currentLine).Code(elementWordId - 1).wordType <> EditWordTypeEnum.W_POINT Then Return True 'перед именем функции нет точки - квадратных скобок нет - все известно
            If elementWordId < 3 Then Return True 'ошибка синтаксиса
            If codeBox.CodeData(currentLine).Code(elementWordId - 2).wordType <> EditWordTypeEnum.W_QUAD_BRACKET_CLOSE Then Return True 'квадратных скобок нет - все известно
            quadBracketBalance = -1
            paramNumber += 1
            'проверяем все слова от слова перед закрывающей ] до начала строки в поисках открывающей [ функции
            For i = elementWordId - 3 To 0 Step -1
                Select Case codeBox.CodeData(currentLine).Code(i).wordType
                    Case EditWordTypeEnum.W_QUAD_BRACKET_OPEN
                        quadBracketBalance += 1
                        If quadBracketBalance = 0 Then Return True 'найдена нужная [ - выход
                    Case EditWordTypeEnum.W_QUAD_BRACKET_CLOSE
                        quadBracketBalance -= 1
                    Case EditWordTypeEnum.W_COMMA
                        If quadBracketBalance = -1 Then paramNumber += 1 'найдена запятая этой функции, номер параметра +1
                End Select
            Next
        End If
        Return False
    End Function

    ''' <summary>
    ''' Процедура создает список перменных на основании глобальных переменных + всех найденных в тексте до текущей строки
    ''' </summary>
    ''' <param name="rtb">ссылка на тестбокс</param>
    ''' <param name="asStrings">передавать их как строки (т. е. в '') или нет</param>
    ''' <param name="onlyLocal">Не добавлять в список глобальные переменные</param>
    Public Shared Function MakeVariablesList(ByRef rtb As RichTextBoxEx, Optional ByVal asStrings As Boolean = False, Optional onlyLocal As Boolean = False) As String()
        'получаем глобальные переменные

        Dim arrVars As List(Of String)
        If onlyLocal Then
            arrVars = New List(Of String)
        Else
            arrVars = mScript.csPublicVariables.GetVariablesList
        End If
        If asStrings Then
            For i As Integer = 0 To arrVars.Count - 1
                arrVars(i) = "'" + arrVars(i) + "'"
            Next
        End If
        'получаем переменные из текста
        Dim curWord As EditWordType
        'Dim qFunctions As Integer = 0, qEvents As Integer = 0 'счетчики для определения где мы - внутри блоков Function / Event или вне (переменные внутри блоков недоступны снаружи)

        For txtLine As Integer = rtb.GetLineFromCharIndex(rtb.SelectionStart) - 1 To 0 Step -1
            If IsNothing(rtb.CodeData(txtLine)) OrElse IsNothing(rtb.CodeData(txtLine).Code) Then Continue For
            Dim finish As Boolean = False
            For wId = 0 To rtb.CodeData(txtLine).Code.GetUpperBound(0)
                'проверяем все слова, от текущей и до первой строки
                curWord = rtb.CodeData(txtLine).Code(wId)
                If curWord.wordType = EditWordTypeEnum.W_VARIABLE Then
                    'If qFunctions <> 0 OrElse qEvents <> 0 Then Continue For
                    Dim newVar As String = curWord.Word.Trim
                    If asStrings Then newVar = "'" + newVar + "'" 'отмечаем как строку
                    'найденное слово - переменная
                    If arrVars.Contains(newVar, StringComparer.CurrentCultureIgnoreCase) Then Continue For 'такая уже есть - продолжаем
                    arrVars.Add(newVar) 'добавляем найденную переменную в список
                    'ElseIf curWord.wordType = EditWordTypeEnum.W_BLOCK_FUNCTION Then
                    '    If wId = 0 Then
                    '        If qFunctions = 0 Then Return arrVars.ToArray 'мы были внутри функции. Все что выше - не интересует (локальные переменные извне блока тут недоступны)
                    '        qFunctions -= 1
                    '    ElseIf rtb.CodeData(txtLine).Code(wId - 1).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    '        qFunctions += 1
                    '    End If
                    'ElseIf curWord.wordType = EditWordTypeEnum.W_BLOCK_EVENT Then
                    '    If wId = 0 Then
                    '        If qEvents = 0 Then Return arrVars.ToArray 'мы были внутри события. Все что выше - не интересует (локальные переменные извне блока тут недоступны)
                    '        qEvents -= 1
                    '    ElseIf rtb.CodeData(txtLine).Code(wId - 1).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    '        qEvents += 1
                    '    End If
                End If
            Next wId
        Next txtLine
        Return arrVars.ToArray
    End Function

    ''' <summary>
    ''' Создает список сигнатур указанной переменной, произведя поиск в глобальных переменных и в коде выше текущей позиции
    ''' </summary>
    ''' <param name="rtb">кодбокс</param>
    ''' <param name="varName">имя переменной, сигнатуры которой ищем</param>
    Private Function MakeVariablesSignaturesList(ByRef rtb As RichTextBox, ByVal varName As String) As String()
        'получаем сигнатуры в глобальных переменных
        Dim arrVars As List(Of String) = mScript.csPublicVariables.CreateListOfVariableSignatures(varName, True)
        If IsNothing(arrVars) Then arrVars = New List(Of String)

        'получаем переменные из текста
        varName = varName.ToLower
        'Dim qFunctions As Integer = 0, qEvents As Integer = 0 'счетчики для определения где мы - внутри блоков Function / Event или вне (переменные внутри блоков недоступны снаружи)
        For txtLine As Integer = rtb.GetLineFromCharIndex(rtb.SelectionStart) - 1 To 0 Step -1
            Dim curWord As EditWordType
            If IsNothing(codeBox.CodeData(txtLine)) OrElse IsNothing(codeBox.CodeData(txtLine).Code) Then Continue For
            For wId = 0 To codeBox.CodeData(txtLine).Code.GetUpperBound(0)
                'проверяем все слова, от текущей и до первой строки
                curWord = codeBox.CodeData(txtLine).Code(wId)
                If curWord.wordType = EditWordTypeEnum.W_VARIABLE AndAlso wId <= codeBox.CodeData(txtLine).Code.GetUpperBound(0) - 2 Then
                    'If qFunctions <> 0 OrElse qEvents <> 0 Then Continue For
                    Dim newVar As String = curWord.Word.Trim.ToLower
                    If newVar = varName AndAlso codeBox.CodeData(txtLine).Code(wId + 1).wordType = EditWordTypeEnum.W_QUAD_BRACKET_OPEN AndAlso _
                        codeBox.CodeData(txtLine).Code(wId + 2).wordType = EditWordTypeEnum.W_SIMPLE_STRING Then
                        'найденное слово - нужная нам переменная с последующей за ней сигнатурой varName['signature'].
                        Dim sig As String = codeBox.CodeData(txtLine).Code(wId + 2).Word.Trim
                        If arrVars.Contains(sig, StringComparer.CurrentCultureIgnoreCase) Then Continue For 'такая уже есть - продолжаем
                        arrVars.Add(sig) 'добавляем найденную переменную в список
                    End If
                    'ElseIf curWord.wordType = EditWordTypeEnum.W_BLOCK_FUNCTION Then
                    '    If wId = 0 Then
                    '        If qFunctions = 0 Then Return arrVars.ToArray 'мы были внутри функции. Все что выше - не интересует (локальные переменные извне блока тут недоступны)
                    '        qFunctions -= 1
                    '    ElseIf codeBox.CodeData(txtLine).Code(wId - 1).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    '        qFunctions += 1
                    '    End If
                    'ElseIf curWord.wordType = EditWordTypeEnum.W_BLOCK_EVENT Then
                    '    If wId = 0 Then
                    '        If qEvents = 0 Then Return arrVars.ToArray 'мы были внутри события. Все что выше - не интересует (локальные переменные извне блока тут недоступны)
                    '        qEvents -= 1
                    '    ElseIf codeBox.CodeData(txtLine).Code(wId - 1).wordType = EditWordTypeEnum.W_CYCLE_END Then
                    '        qEvents += 1
                    '    End If
                End If
            Next wId
        Next txtLine
        Return arrVars.ToArray

    End Function
#End Region

#Region "Undo & Redo"

    Private Sub mnuUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUndo.Click
        'отмена редактирования
        codeBox.csUndo.Undo()
        'If masUndo.Count = 1 Then
        '    'если в массиве masUndo остался только исходный текст - отменять нечего
        '    Return
        'End If

        'If MakeNextUndoItem = False Then
        '    'отмена вызвана до срабатывания таймера timUndo. Сохраняем в masUndo содержимое rtb и останавливаем таймер
        '    timUndo.Stop()
        '    MakeNextUndoItem = True
        '    masUndo.Push(codeBox.Rtf)
        '    masUndoCode.Push(CopyCodeDataArray(codeBox.CodeData))
        'End If

        ''добавляем инфо в masRedo
        'masRedo.Push(masUndo.Pop)
        'masRedoCode.Push(CopyCodeDataArray(masUndoCode.Pop))

        'codebox.rtbCanRaiseEvents = False 'запрет событий в rtb
        'Dim curLine As Integer = codeBox.GetLineFromCharIndex(codeBox.SelectionStart) '№ текущей линии
        'Dim curSelStart As Integer '= codeBox.SelectionStart - codeBox.GetFirstCharIndexOfCurrentLine 'положение каретки относительно начала линии
        'Dim curLineLength As Integer '= codeBox.Lines(curLine).Length 'длина текущей линии
        'If codeBox.Lines.Count = 0 Then
        '    curLineLength = 0
        '    curLine = -1
        'Else
        '    curLineLength = codeBox.Lines(curLine).Length 'длина текущей линии
        '    curLine = codeBox.GetLineFromCharIndex(codeBox.SelectionStart) '№ текущей линии
        'End If
        'codeBox.Rtf = masUndo.Peek 'восстанавливаем содержимое rtb до редактирования из masUndo.
        'codeBox.CodeData = CopyCodeDataArray(masUndoCode.Peek)
        ''В masUndo всегда последним элементом должно быть текущее содержимое rtb, поэтому masUndo.Peek а не masUndo.Pop
        ''восстанавливаем положение каретки до вставки 
        'If curLine > codeBox.Lines.GetUpperBound(0) Then
        '    codeBox.SelectionStart = codeBox.TextLength
        'Else
        '    Dim sStart As Integer ' = codeBox.GetFirstCharIndexFromLine(curLine) + curSelStart + codeBox.Lines(curLine).Length - curLineLength
        '    If curLine = -1 Then
        '        sStart = 0
        '    Else
        '        sStart = codeBox.GetFirstCharIndexFromLine(curLine) + curSelStart + codeBox.Lines(curLine).Length - curLineLength
        '    End If
        '    If sStart < 0 Then sStart = 0
        '    codeBox.SelectionStart = sStart
        'End If
        'codebox.rtbCanRaiseEvents = True 'разрешаем события в rtb
    End Sub

    Private Sub mnuRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRedo.Click
        'повтор редактирования после его отмены (выполнения Undo)
        codeBox.csUndo.Redo()
        'If masRedo.Count = 0 Then
        '    'содержимое mnuRedo пусто - отменять нечего
        '    Return
        'End If
        ''добавляем содержимое последнего элемента в masRedo (= новому содержимому rtb) в массив masUndo
        'masUndo.Push(masRedo.Peek)
        'masUndoCode.Push(CopyCodeDataArray(masRedoCode.Peek))

        'codebox.rtbCanRaiseEvents = False 'запрет событий в rtb
        'Dim curLine As Integer = codeBox.GetLineFromCharIndex(codeBox.SelectionStart) '№ текущей линии
        'Dim curSelStart As Integer = codeBox.SelectionStart - codeBox.GetFirstCharIndexOfCurrentLine 'положение каретки относительно начала линии
        'Dim curLineLength As Integer = codeBox.Lines(curLine).Length 'длина текущей линии
        'codeBox.Rtf = masRedo.Pop 'восстанавливаем содержимое rtb из последнего элемента masRedo
        'codeBox.CodeData = CopyCodeDataArray(masRedoCode.Pop)
        ''восстанавливаем положение каретки до вставки 
        'If curLine > codeBox.Lines.GetUpperBound(0) Then
        '    codeBox.SelectionStart = codeBox.TextLength
        'Else
        '    codeBox.SelectionStart = codeBox.GetFirstCharIndexFromLine(curLine) + curSelStart + codeBox.Lines(curLine).Length - curLineLength
        'End If
        'codebox.rtbCanRaiseEvents = True 'разрешаем события в rtb
    End Sub

#End Region

#Region "Other Context Menu"
    Private Sub mnuPaste_Click(ByVal sender As ToolStripMenuItem, ByVal e As System.EventArgs) Handles mnuPaste.Click
        'Выполняет вставку текста из буфера   
        If Clipboard.ContainsText = False Then Exit Sub

        Dim cbTxt As String = Clipboard.GetText(System.Windows.Forms.TextDataFormat.Text) 'получаем тект из буфера

        'If cbTxt.EndsWith(vbLf) AndAlso cbTxt.EndsWith(vbNewLine) = False Then
        '    cbTxt = cbTxt.Substring(0, cbTxt.Length - 1)
        'End If

        'Предотвращаем вставку слишком длинной строки
1:
        Dim cbMas() As String = Split(cbTxt, vbLf) ' vbNewLine)
        Dim curLineLength As Integer = 0
        If codeBox.Lines.Count > 0 Then curLineLength = codeBox.Lines(codeBox.GetLineFromCharIndex(codeBox.GetFirstCharIndexOfCurrentLine)).Length
        For i = 0 To cbMas.GetUpperBound(0)
            If i = 0 Then
                If curLineLength + cbMas(i).Length >= rtbMaxLineLength Then
                    cbMas(i) = vbNewLine + cbMas(i)
                    cbTxt = Join(cbMas, vbNewLine)
                    GoTo 1
                End If
            End If
            If cbMas(i).Length >= rtbMaxLineLength Then
                cbMas(i) = cbMas(i).Substring(0, rtbMaxLineLength / 2) + vbNewLine + cbMas(i).Substring(rtbMaxLineLength / 2 + 1)
                cbTxt = Join(cbMas, vbNewLine)
                GoTo 1
            End If
        Next

        codeBox.SelectedText = cbTxt

        '        Dim cbLinesCount As Integer = cbMas.GetUpperBound(0) 'получаем кол-во строк в тексте из буфера - 1

        '        If cbLinesCount < 1 Then
        '            'если переноса строки нет - нет нужды в этой процедуре. Вставляем текст и выход
        '            codeBox.SelectedText = cbTxt
        '            Return
        '        End If
        '        'сохраняем начало и длину выделения до редактирования
        '        codeBox.lastSelectionStartLine = codeBox.GetLineFromCharIndex(codeBox.SelectionStart)
        '        codeBox.lastSelectionEndLine = codeBox.GetLineFromCharIndex(codeBox.SelectionStart + codeBox.SelectionLength)

        '        codebox.rtbCanRaiseEvents = False 'запрет событий в rtb
        '        Dim wasTextBlock As Boolean = False
        '        'очищаем предыдущее выделение
        '        If codeBox.lastSelectionStartLine <> codeBox.lastSelectionEndLine Then
        '            codeBox.SelectedText = ""
        '            'если были выделены символы на нескольких строках, то убираем из codeBox.CodeData лишнее            
        '            wasTextBlock = codeBox.IsOpenedTextBlockBetweenLines(codeBox.lastSelectionStartLine, codeBox.lastSelectionEndLine)
        '            Dim linesDeleted = codeBox.lastSelectionEndLine - codeBox.lastSelectionStartLine
        '            For i As Integer = codeBox.lastSelectionStartLine + 1 To codeBox.CodeData.GetUpperBound(0) - linesDeleted
        '                codeBox.CodeData(i) = codeBox.CodeData(i + linesDeleted)
        '            Next
        '            ReDim Preserve codeBox.CodeData(codeBox.CodeData.GetUpperBound(0) - linesDeleted)
        '            codeBox.lastSelectionEndLine = codeBox.lastSelectionStartLine
        '        End If

        '        Dim currentLine As Integer = codeBox.GetLineFromCharIndex(codeBox.GetFirstCharIndexOfCurrentLine)
        '        ReDim Preserve codeBox.CodeData(codeBox.CodeData.GetUpperBound(0) + cbLinesCount) 'расширяем массив на кол-во новых линий
        '        For i As Integer = codeBox.CodeData.GetUpperBound(0) To currentLine + cbLinesCount Step -1
        '            'сдвигаем инфо в codeBox.CodeData вправо (в конец), начиная со строки, следующей за текущей, оставляя пустые места в codeBox.CodeData для вставки текста из буфера
        '            codeBox.CodeData(i) = codeBox.CodeData(i - cbLinesCount)
        '        Next

        '        codeBox.SelectedText = cbTxt 'вставляем текст в rtb
        '        If cbTxt.Length > 0 AndAlso codeBox.prevLineId = -1 Then codeBox.prevLineId = 0
        '        codeBox.CheckPrevLine(codeBox, currentLine - IIf(currentLine = 0, 0, 1), IIf(wasTextBlock, codeBox.Lines.GetUpperBound(0), currentLine + cbLinesCount), True) 'проверка синтаксиса и раскраска измененных строк
        '        'если был совершен переход на новую линию - сохраняем информацию о текущем содержимом строки, которую будут редактировать
        '        If codeBox.prevLineId <> codeBox.GetLineFromCharIndex(codeBox.GetFirstCharIndexOfCurrentLine) Then codeBox.SetPrevLine(codeBox)

        '        codebox.rtbCanRaiseEvents = True 'разрешаем события в rtb
    End Sub


    Private Sub mnuCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopy.Click
        If codeBox.SelectionLength = 0 Then Return
        Clipboard.SetText(codeBox.SelectedText, TextDataFormat.Text)
    End Sub

    Private Sub mnuCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCut.Click
        If codeBox.SelectionLength = 0 Then Return
        Clipboard.SetText(codeBox.SelectedText, TextDataFormat.Text)
        codeBox.SelectedText = ""

        'codebox.rtbCanRaiseEvents = False
        ''сохраняем начало и длину выделения до редактирования
        'codeBox.lastSelectionStartLine = codeBox.GetLineFromCharIndex(codeBox.SelectionStart)
        'codeBox.lastSelectionEndLine = codeBox.GetLineFromCharIndex(codeBox.SelectionStart + codeBox.SelectionLength)

        'codeBox.SelectedText = ""
        'Dim wasTextBlock As Boolean = False
        'If codeBox.lastSelectionStartLine <> codeBox.lastSelectionEndLine Then
        '    'если были выделены символы на нескольких строках, то при редактировании они сотрутся (точнее, уже стерты).
        '    wasTextBlock = codeBox.IsOpenedTextBlockBetweenLines(codeBox.lastSelectionStartLine, codeBox.lastSelectionEndLine) 'был незакрытый текстовый блок в убранном тексте?
        '    'Убираем из codeBox.CodeData лишнее
        '    Dim linesDeleted = codeBox.lastSelectionEndLine - codeBox.lastSelectionStartLine
        '    For i As Integer = codeBox.lastSelectionStartLine + 1 To codeBox.CodeData.GetUpperBound(0) - linesDeleted
        '        codeBox.CodeData(i) = codeBox.CodeData(i + linesDeleted)
        '    Next
        '    ReDim Preserve codeBox.CodeData(codeBox.CodeData.GetUpperBound(0) - linesDeleted)
        '    codeBox.lastSelectionEndLine = codeBox.lastSelectionStartLine
        'End If

        'If codeBox.TextLength > 0 Then
        '    Dim currentLine As Integer = codeBox.GetLineFromCharIndex(codeBox.GetFirstCharIndexOfCurrentLine) 'получаем текущую линию

        '    SendMessage(codeBox.Handle, WM_SetRedraw, 0, 0)
        '    If wasTextBlock Then
        '        'если в результате последнего редактирования был удален кусок текстового блока, перерисовываем весь текст от текущей динии и до конца
        '        codeBox.CheckPrevLine(codeBox, codeBox.GetLineFromCharIndex(codeBox.SelectionStart), codeBox.Lines.GetUpperBound(0), True) 'проверка синтаксиса и раскраска измененных строк
        '    End If
        '    'перерисовываем текст с пометкой "незавершенная строка"
        '    Dim currentWordId As Integer
        '    Dim upperLine As Integer = codeBox.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})
        '    upperLine = codeBox.GetLineFromCharIndex(upperLine)
        '    codeBox.PrepareTextNonCompleted(codeBox, codeBox.GetLineFromCharIndex(codeBox.SelectionStart), codeBox.GetLineFromCharIndex(codeBox.SelectionStart), True, currentWordId)
        '    Dim charPos As Integer = codeBox.GetCharIndexFromPosition(New Point With {.X = 5, .Y = 5})
        '    Dim lId As Integer = codeBox.GetLineFromCharIndex(charPos)
        '    SendMessage(codeBox.Handle, EM_LINESCROLL, 0, upperLine - lId)
        '    SendMessage(codeBox.Handle, WM_SetRedraw, 1, 0)
        '    codeBox.Refresh()
        '    PrepareRtbList(currentLine, currentWordId)

        '    Dim curWord As EditWordType = Nothing
        '    Dim paramNumber As Integer
        '    Dim classId As Integer
        '    GetCurrentElementWord(currentLine, currentWordId, paramNumber, curWord, classId)
        '    codeBox.ShowElementInfo(curWord.classId, curWord.Word, curWord.wordType, paramNumber)
        'End If

        'CheckLabels(codeBox)

        ''подготавливаем Undo
        'If MakeNextUndoItem = False And masUndo.Count > 1 Then
        '    masUndo.Pop()
        '    masUndoCode.Pop()
        'End If
        'masUndo.Push(codeBox.Rtf)
        'masUndoCode.Push(CopyCodeDataArray(codeBox.CodeData))
        'MakeNextUndoItem = False
        'timUndo.Stop()
        'timUndo.Start()

        'codebox.rtbCanRaiseEvents = True

        'RaiseEvent TextChanged(codeBox, e)

    End Sub

    Private Sub mnuSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSelectAll.Click
        codeBox.SelectAll()
    End Sub

#End Region

#Region "Other"
    Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавьте все инициализирующие действия после вызова InitializeComponent().
        Try
            provider_points.NumberDecimalSeparator = "."
        Catch ex As Exception

        End Try

        SplitContainerVertical.Panel2.Controls.Add(codeBox)
        With codeBox
            .lstRtb = lstRtb
            .Controls.Add(lstRtb)
            .wbRtbHelp = wbRtbHelp
            .FillOperators()
            .FillStyle()
            .FillCheckCodeStructure()

            .rtbCanRaiseEvents = False
            .PrepareText(codeBox, 0, -1, True)
            .rtbCanRaiseEvents = True
            .ContextMenuStrip = mnuRtb
            .Visible = True
        End With

    End Sub

    Private Sub helpDocument_Click(ByVal sender As HtmlDocument, ByVal e As HtmlElementEventArgs) Handles helpDocument.Click
        Dim hEl As HtmlElement = sender.GetElementFromPoint(e.OffsetMousePosition)

        Do While IsNothing(hEl) = False
            If hEl.GetAttribute("Line").Length > 0 Then
                Dim curLine As Integer = Convert.ToInt32(hEl.GetAttribute("Line"))
                codeBox.Select(codeBox.GetFirstCharIndexFromLine(curLine), codeBox.Lines(curLine).Length)
                codeBox.Focus()
                codeBox.ScrollToCaret()
            End If
            hEl = hEl.Parent
        Loop
    End Sub

    Private Sub wbRtbHelp_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles wbRtbHelp.DocumentCompleted
        helpDocument = wbRtbHelp.Document
    End Sub


    Private Sub SplitContainerHorizontal_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainerHorizontal.SplitterMoved
        'SplitContainerVertical.Height = SplitContainerHorizontal.Panel1.ClientRectangle.Height
        codeBox.Height = SplitContainerHorizontal.Panel1.ClientRectangle.Height
    End Sub
#End Region

End Class
