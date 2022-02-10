'TOD
'1. Доделать файлы помощи функций
'2. Загрузка всей структуры из папки квеста (пересмотреть все картинки, аудио по умолчанию)
'2. Переменные с eventId
'3. CreateEvent - событие создания дочерних элементов
'4. Проблема при восстановлении после игры - переход не в точности в то же место, где был
'5. CodeTextBox: сбой при удалении Else

'2.Везде ToolTip-ы
'3.Сводная информация (когда ни одно свойство не активно), кнопка для нее
'7.Свойства папок - IsExpanded
'8.Да/Нет свойств сделать чекбоксом
'9.Удалять вкладки если свойства не изменялись (возможно, делать их другого цвета)
'11.Переменные локальные и глобальные - проверка, как действовать в циклах Event, Function
'12.Узлы в разные цвета в зависимости от заполненности свойств (= поумолчанию?) и tracking
'14.Edit/Remove в frmClassEditor - изменение всех связанных структур при удалении/изменении класса, свойства, функции - проверить


Module modMain
#Region "Мелкие классы"
    Public Class clsDEFAULT_COLORS
        'цвета для текста узлов дерева классов (обычный, выбранный и скрытый)
        ''' <summary>Обычный цвет узла</summary>
        Public NodeForeColor As Color = Color.Black
        ''' <summary>Цвет текста выделенного узла</summary>
        Public NodeSelForeColor As Color = Color.Red
        ''' <summary>Цвет скрытого узла</summary>
        Public NodeHiddenForeColor As Color = Color.Gray
        'цвета для текста узлов дерева классов выбранный и обычный
        ''' <summary>Цвет фона выделенного узла</summary>
        Public NodeSelBackColor As Color = Color.LawnGreen
        ''' <summary>Обычный цвет фона узла, кнопки</summary>
        Public ControlTransparentBackground As Color = Color.FromArgb(0, 255, 255, 255)
        ''' <summary>Цвет текста на кнопке события, если событие пустое</summary>
        Public EventButtonEmpty As Color = Color.Gray
        ''' <summary>Цвет текста на кнопке события, если событие заполнено</summary>
        Public EventButtonFilled As Color = Color.DarkGreen
        ''' <summary>Обычный цвет текстбокса со свойством</summary>
        Public TextBoxBackColor As Color = Color.White
        ''' <summary>Цвет текущего текстбокса со свойством</summary>
        Public TextBoxHighlighted As Color = Color.LawnGreen
        ''' <summary>Обычный цвет надписи рядом с текстбоксом со свойством</summary>
        Public LabelForeColor As Color = Color.Black
        ''' <summary>Цвет надписи рядом с текстбоксом в выделенным свойством</summary>
        Public LabelHighLighted As Color = Color.DarkGreen
        ''' <summary>Цвет яркого выделения в кодбоксах</summary>
        Public codeBoxHighLight As Color = Color.Orange
        ''' <summary>Цвет слабого выделения в кодбоксах</summary>
        Public codeBoxDesignate As Color = Color.LightGray
        ''' <summary>Цвет выделения вкладок того же класса</summary>
        Public SameClassToolButton As Color = Color.FromArgb(253, 245, 230)
        ''' <summary>Цвет выделения related вкладок</summary>
        Public RelatedToolButton As Color = Color.FromArgb(176, 196, 222)
    End Class

    ''' <summary>Расширенный класс TextBox для работы со свойствами</summary>
    Public Class TextBoxEx
        Inherits TextBox
        ''' <summary>Ассоциировання с данным текстбоксом окно-панель</summary>
        Public Property childPanel As clsPanelManager.clsChildPanel
        ''' <summary>Разрешать ли события Validating, SelectedIndexChanged ...</summary>
        Public Property AllowEvents As Boolean = True
        ''' <summary>Является ли данный контрол кнопкой для функции пользователя</summary>
        Public Property IsFunctionButton As Boolean = False
        ''' <summary>Положение по Y свойства для панели свойств по умолчанию, 2 уровня и 3 уровня</summary>
        Public Property PropertyControlsTop As New clsPanelManager.cPropertyControlsTop
        ''' <summary>Ссылка на лейбл с названием свойства</summary>
        Public Label As Label
        ''' <summary>Ссылка на кнопку помощи рядом с кнопкой события</summary>
        Public ButtonHelp As ButtonEx
        ''' <summary>Ссылка на кнопку настроек рядом с основным контролом</summary>
        Public ButtonConfig As ButtonEx
        ''' <summary>Нужна ли данному контролу кнопка настройки</summary>
        Public Property hasConfigButton As Boolean = False

        Private Sub VisibleChangedEvent(sender As Object, e As EventArgs) Handles Me.VisibleChanged
            If IsNothing(cPanelManager) Then Return
            Dim visibleState As Boolean = sender.Visible
            If IsNothing(Label) = False Then Label.Visible = visibleState
            If IsNothing(ButtonHelp) = False Then ButtonHelp.Visible = visibleState
            If IsNothing(ButtonConfig) = False Then
                If visibleState = False Then
                    ButtonConfig.Visible = False
                Else
                    Dim mainPanel As clsPanelManager.PanelEx = cPanelManager.dictDefContainers(mScript.mainClassHash(currentClassName))
                    ButtonConfig.Visible = mainPanel.chkConfig.Checked
                End If
            End If
        End Sub


        ''' <summary>
        ''' Изменяет свое положение и положение связанных контролов (лейбла, кнопки помощи...) исходя из значения PropertyControlsTop, а также видимость исходя из значения pHidden
        ''' </summary>
        ''' <param name="childPanel">ссылка на панель, где размещен контрол</param>
        ''' <param name="pHidden">Ссылка на свойство отображения из класса mScript.mainClass().Properties()(Name).Hidden</param>
        ''' <returns>Свойство Visible контрола</returns>
        Public Function ChangeTopAndVisible(ByRef childPanel As clsPanelManager.clsChildPanel, ByRef pHidden As MatewScript.PropertyHiddenEnum) As Boolean
            'Прячем при необходимости
            If pHidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pHidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            If childPanel.child2Name.Length = 0 AndAlso (pHidden = MatewScript.PropertyHiddenEnum.LEVEL2_ONLY OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL3_ONLY _
                                  OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL23_ONLY) Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            If (childPanel.child2Name.Length > 0 AndAlso childPanel.child3Name.Length = 0) AndAlso (pHidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL3_ONLY _
                                                                               OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL13_ONLY) Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            If childPanel.child3Name.Length > 0 AndAlso (pHidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL2_ONLY _
                                  OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL12_ONLY) Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            'контрол должен быть видим. Изменяет положение по высоте
            Dim confOffset As Integer
            Dim lblOffset As Integer
            Dim helpOffset As Integer

            If IsNothing(ButtonConfig) = False Then confOffset = Me.Top - ButtonConfig.Top
            If IsNothing(Label) = False Then lblOffset = Me.Top - Label.Top
            If IsNothing(ButtonHelp) = False Then helpOffset = Me.Top - ButtonHelp.Top

            If childPanel.child2Name.Length = 0 Then
                Me.Top = PropertyControlsTop.TopLevel1
            ElseIf childPanel.child3Name.Length = 0 Then
                Me.Top = PropertyControlsTop.TopLevel2
            Else
                Me.Top = PropertyControlsTop.TopLevel3
            End If

            If IsNothing(ButtonConfig) = False Then ButtonConfig.Top = Me.Top - confOffset
            If IsNothing(Label) = False Then Label.Top = Me.Top - lblOffset
            If IsNothing(ButtonHelp) = False Then ButtonHelp.Top = Me.Top - helpOffset

            Me.Visible = True
            Return True
        End Function

    End Class

    Public Class ComboBoxEx
        Inherits ComboBox
        ''' <summary>Ассоциировання с данным текстбоксом окно-панель</summary>
        Public Property childPanel As clsPanelManager.clsChildPanel
        ''' <summary>Разрешать ли события Validating, SelectedIndexChanged ...</summary>
        Public Property AllowEvents As Boolean = True
        ''' <summary>Является ли данный контрол кнопкой для функции пользователя</summary>
        Public Property IsFunctionButton As Boolean = False
        ''' <summary>Положение по Y свойства для панели свойств по умолчанию, 2 уровня и 3 уровня</summary>
        Public Property PropertyControlsTop As New clsPanelManager.cPropertyControlsTop
        ''' <summary>Ссылка на лейбл с названием свойства</summary>
        Public Label As Label
        ''' <summary>Ссылка на кнопку помощи рядом с кнопкой события</summary>
        Public ButtonHelp As ButtonEx
        ''' <summary>Ссылка на кнопку настроек рядом с основным контролом</summary>
        Public ButtonConfig As ButtonEx
        ''' <summary>Нужна ли данному контролу кнопка настройки</summary>
        Public Property hasConfigButton As Boolean = False

        Private Sub VisibleChangedEvent(sender As Object, e As EventArgs) Handles Me.VisibleChanged
            If IsNothing(cPanelManager) Then Return
            Dim visibleState As Boolean = sender.Visible
            If IsNothing(Label) = False Then Label.Visible = visibleState
            If IsNothing(ButtonHelp) = False Then ButtonHelp.Visible = visibleState
            If IsNothing(ButtonConfig) = False Then
                If visibleState = False Then
                    ButtonConfig.Visible = False
                Else
                    Dim mainPanel As clsPanelManager.PanelEx = cPanelManager.dictDefContainers(mScript.mainClassHash(currentClassName))
                    ButtonConfig.Visible = mainPanel.chkConfig.Checked
                End If
            End If
        End Sub


        ''' <summary>
        ''' Изменяет свое положение и положение связанных контролов (лейбла, кнопки помощи...) исходя из значения PropertyControlsTop, а также видимость исходя из значения pHidden
        ''' </summary>
        ''' <param name="childPanel">ссылка на панель, где размещен контрол</param>
        ''' <param name="pHidden">Ссылка на свойство отображения из класса mScript.mainClass().Properties()(Name).Hidden</param>
        ''' <returns>Свойство Visible контрола</returns>
        Public Function ChangeTopAndVisible(ByRef childPanel As clsPanelManager.clsChildPanel, ByRef pHidden As MatewScript.PropertyHiddenEnum) As Boolean
            'Прячем при необходимости
            If pHidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pHidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            If childPanel.child2Name.Length = 0 AndAlso (pHidden = MatewScript.PropertyHiddenEnum.LEVEL2_ONLY OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL3_ONLY _
                                  OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL23_ONLY) Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            If (childPanel.child2Name.Length > 0 AndAlso childPanel.child3Name.Length = 0) AndAlso (pHidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL3_ONLY _
                                                                               OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL13_ONLY) Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            If childPanel.child3Name.Length > 0 AndAlso (pHidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL2_ONLY _
                                  OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL12_ONLY) Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            'контрол должен быть видим. Изменяет положение по высоте
            Dim confOffset As Integer
            Dim lblOffset As Integer
            Dim helpOffset As Integer

            If IsNothing(ButtonConfig) = False Then confOffset = Me.Top - ButtonConfig.Top
            If IsNothing(Label) = False Then lblOffset = Me.Top - Label.Top
            If IsNothing(ButtonHelp) = False Then helpOffset = Me.Top - ButtonHelp.Top

            If childPanel.child2Name.Length = 0 Then
                Me.Top = PropertyControlsTop.TopLevel1
            ElseIf childPanel.child3Name.Length = 0 Then
                Me.Top = PropertyControlsTop.TopLevel2
            Else
                Me.Top = PropertyControlsTop.TopLevel3
            End If

            If IsNothing(ButtonConfig) = False Then ButtonConfig.Top = Me.Top - confOffset
            If IsNothing(Label) = False Then Label.Top = Me.Top - lblOffset
            If IsNothing(ButtonHelp) = False Then ButtonHelp.Top = Me.Top - helpOffset

            Me.Visible = True
            Return True
        End Function

    End Class

    Public Class ButtonEx
        Inherits Button
        ''' <summary>Ассоциировання с данным текстбоксом окно-панель</summary>
        Public Property childPanel As clsPanelManager.clsChildPanel
        ' ''' <summary>Техническая информация о свойстве, где Key - имя свойства, связанного с данным контролом</summary>
        'Public Property propertyData As KeyValuePair(Of String, MatewScript.PropertiesInfoType)
        ''' <summary>Является ли данный контрол кнопкой для функции пользователя</summary>
        Public Property IsFunctionButton As Boolean = False
        ''' <summary>Положение по Y свойства для панели свойств по умолчанию, 2 уровня и 3 уровня</summary>
        Public Property PropertyControlsTop As New clsPanelManager.cPropertyControlsTop
        ''' <summary>Ссылка на лейбл с названием свойства</summary>
        Public Label As Label
        ''' <summary>Ссылка на кнопку помощи рядом с кнопкой события</summary>
        Public ButtonHelp As ButtonEx
        ''' <summary>Ссылка на кнопку настроек рядом с основным контролом</summary>
        Public ButtonConfig As ButtonEx

        ''' <summary>Для кнопок навигации - к какому классу перейти. Устанавливается в CreatePropertiesControl</summary>
        Public Property NavigateToClassId As Integer
        ''' <summary>Для кнопок навигации - Name элемента-родителя, куда осуществляется переход. Устанавливается динамично</summary>
        Public Property NavigateToParentName As String = ""
        ''' <summary>Для кнопок навигации - осуществляется переход к элементам 3 (например, пункты меню) уровня или нет (например, действия). Устанавливается динамично</summary>
        Public Property NavigateToThirdLevel As Boolean = False

        ''' <summary>Является ли данный контрол кнопкой настройки</summary>
        Public Property IsConfigButton As Boolean = False
        ''' <summary>Нужна ли данному контролу кнопка настройки</summary>
        Public Property hasConfigButton As Boolean = False


        Private Sub VisibleChangedEvent(sender As Object, e As EventArgs) Handles Me.VisibleChanged
            If IsNothing(cPanelManager) Then Return
            Dim visibleState As Boolean = sender.Visible
            If IsNothing(Label) = False Then Label.Visible = visibleState
            If IsNothing(ButtonHelp) = False Then ButtonHelp.Visible = visibleState
            If IsNothing(ButtonConfig) = False Then
                If visibleState = False Then
                    ButtonConfig.Visible = False
                Else
                    Dim mainPanel As clsPanelManager.PanelEx = cPanelManager.dictDefContainers(mScript.mainClassHash(currentClassName))
                    ButtonConfig.Visible = mainPanel.chkConfig.Checked
                End If
            End If
        End Sub

        ''' <summary>
        ''' Изменяет свое положение и положение связанных контролов (лейбла, кнопки помощи...) исходя из значения PropertyControlsTop, а также видимость исходя из значения pHidden
        ''' </summary>
        ''' <param name="childPanel">ссылка на панель, где размещен контрол</param>
        ''' <param name="pHidden">Ссылка на свойство отображения из класса mScript.mainClass().Properties()(Name).Hidden</param>
        ''' <returns>Свойство Visible контрола</returns>
        Public Function ChangeTopAndVisible(ByRef childPanel As clsPanelManager.clsChildPanel, ByRef pHidden As MatewScript.PropertyHiddenEnum) As Boolean
            'Прячем при необходимости
            If pHidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse pHidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            If childPanel.child2Name.Length = 0 AndAlso (pHidden = MatewScript.PropertyHiddenEnum.LEVEL2_ONLY OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL3_ONLY _
                                  OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL23_ONLY) Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            If (childPanel.child2Name.Length > 0 AndAlso childPanel.child3Name.Length = 0) AndAlso (pHidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL3_ONLY _
                                                                               OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL13_ONLY) Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            If childPanel.child3Name.Length > 0 AndAlso (pHidden = MatewScript.PropertyHiddenEnum.LEVEL1_ONLY OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL2_ONLY _
                                  OrElse pHidden = MatewScript.PropertyHiddenEnum.LEVEL12_ONLY) Then
                Me.Visible = False
                VisibleChangedEvent(Me, New EventArgs)
                Return False
            End If
            'контрол должен быть видим. Изменяет положение по высоте
            Dim confOffset As Integer
            Dim lblOffset As Integer
            Dim helpOffset As Integer

            If IsNothing(ButtonConfig) = False Then confOffset = Me.Top - ButtonConfig.Top
            If IsNothing(Label) = False Then lblOffset = Me.Top - Label.Top
            If IsNothing(ButtonHelp) = False Then helpOffset = Me.Top - ButtonHelp.Top

            If childPanel.child2Name.Length = 0 Then
                Me.Top = PropertyControlsTop.TopLevel1
            ElseIf childPanel.child3Name.Length = 0 Then
                Me.Top = PropertyControlsTop.TopLevel2
            Else
                Me.Top = PropertyControlsTop.TopLevel3
            End If

            If IsNothing(ButtonConfig) = False Then ButtonConfig.Top = Me.Top - confOffset
            If IsNothing(Label) = False Then Label.Top = Me.Top - lblOffset
            If IsNothing(ButtonHelp) = False Then ButtonHelp.Top = Me.Top - helpOffset

            Me.Visible = True
            Return True
        End Function
    End Class

    Public Class ContextMenuStripEx
        Inherits ContextMenuStrip
        ''' <summary>Текстбокс/комбокс со свойством, для которого открывается меню</summary>
        Public Property OwnerControl As Control
    End Class


    Public Class cLog
        Public Sub PrintToLog(ByVal aStr As String)
            'frmLog.lstLog.Items.Add(aStr)
        End Sub

        Public Overridable Function GetChildInfo(ByVal classId As Integer, ByVal child2Name As String, ByVal supraName As String, Optional child3Name As String = "") As String
            Dim className As String = ""
            If classId = -2 Then
                className = "Variable"
            ElseIf classId = -3 Then
                className = "Function"
            ElseIf classId = -1 Then
                className = ""
            Else
                className = mScript.mainClass(classId).Names(0)
            End If
            Dim aStr As String = className + " ("
            If child2Name.Length = 0 Then
                aStr += "Def"
            Else
                If classId < 0 Then
                    aStr += child2Name
                Else
                    aStr += mScript.PrepareStringToPrint(child2Name, Nothing, False)
                End If
            End If
            aStr += "). Parent: "
            If supraName.Length = 0 Then
                aStr += "-"
            ElseIf classId >= 0 Then
                aStr += supraName
            End If
            Return aStr
        End Function

        Public Overridable Function GetChildInfo(ByRef childPanel As clsPanelManager.clsChildPanel) As String
            Return GetChildInfo(childPanel.classId, childPanel.child2Name, childPanel.supraElementName, childPanel.child3Name)
        End Function

        Public Function GetChild2Str(ByVal className As String, ByVal child2Id As Integer) As String
            If child2Id = -1 Then Return "Def"
            Dim classId As Integer = mScript.mainClassHash(className)
            If IsNothing(mScript.mainClass(classId).ChildProperties) = False AndAlso mScript.mainClass(classId).ChildProperties.Count > child2Id Then
                Return mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("Name").Value, Nothing, False)
            Else
                Return "[NOT FOUND " + child2Id.ToString + "]"
            End If
        End Function
    End Class
#End Region

    Public Class clsQuestEnvironment

#Region "Данные для редактора"
        Public Enum TreeNodeRenameEnum As Byte
            DoubleClick = 0
            SpacePressed = 1
        End Enum

        Public codeBoxShadowed As New CodeTextBox With {.CanDrawWords = False} ', .DontShowError = True}
        ''' <summary>Переменная для хранения размеров контролов при построении панели свойств</summary>
        Public defPropHeight As Integer = 36
        ''' <summary>Переменная для хранения размеров интервалом между контролами при построении панели свойств</summary>
        Public defPaddingLeft As Integer = 5, defPaddingTop As Integer = 5
        ''' <summary>Размеры кнопок настроек содержимого контролов</summary>
        Public btnConfigSize As New Size(26, 26)
        ''' <summary>При каких условиях переименовывать узел - двойной щелчок или пробел</summary>
        Public treeNodeRename As TreeNodeRenameEnum = TreeNodeRenameEnum.SpacePressed
        ''' <summary>Меню настроек свойств (превратить в скрипт, настроить список...)</summary>
        Public propertiesConfigMenu As ContextMenuStripEx
        ''' <summary>Для временного запрета событий</summary>
        Public EnabledEvents As Boolean = True
        ''' <summary>Меню поддерева (добавить, добавить группу...)</summary>
        Public subTreeMenu As ContextMenuStrip
        ''' <summary>Список цветов, выбранных в свойствах типа Color</summary>
        Public lstSelectedColors As New List(Of String)
        ''' <summary>Короткое имя квеста (оно же - папка квеста)</summary>
        Public DefaultInfoPanelHeight As Integer = 50

        ''' <summary>Короткое имя квеста (оно же - папка квеста)</summary>
        Public UseTranslit As Boolean = True
#End Region

#Region "Переменные Игры"

        ''' <summary>Путь к текущему квесту</summary>
        Public QuestPath As String
        ''' <summary>Короткое имя квеста (оно же - папка квеста)</summary>
        Public QuestShortName As String
        ''' <summary>Режим редактирования (Да/Нет)</summary>
        Public EDIT_MODE As Boolean = True
        ''' <summary>Квест открыт из редактора или нет</summary>
        Public OPENED_FROM_EDITOR As Boolean = True
        Public wbMap As New WebBrowser With {.IsWebBrowserContextMenuEnabled = False, .AllowNavigation = False, .Dock = DockStyle.Fill}
        Public wbMagic As New WebBrowser With {.IsWebBrowserContextMenuEnabled = False, .AllowNavigation = False, .Dock = DockStyle.Fill}

#End Region

        Public Sub SaveQuest()
            'сохраняем структуру классов
            mScript.SaveClasses(My.Computer.FileSystem.CombinePath(QuestPath, "classes.xml"))
            'сохраняем главный файл игры
            SaveQuestGameFile(My.Computer.FileSystem.CombinePath(QuestPath, "game.world"))
            'сохраняем главный для редактора
            SaveQuestEditorFile(My.Computer.FileSystem.CombinePath(QuestPath, "editor.world"))
        End Sub

        Public Sub LoadQuest(ByVal questFolder As String)
            If My.Computer.FileSystem.DirectoryExists(questFolder) = False Then
                MessageBox.Show("Путь к квесту не найден!", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If
            QuestPath = questFolder
            Dim dirInf As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(QuestPath)
            QuestShortName = dirInf.Name

            Dim fileName As String
            fileName = My.Computer.FileSystem.CombinePath(questFolder, "classes.xml")
            If My.Computer.FileSystem.FileExists(fileName) = False Then
                MessageBox.Show("Не найден файл classes.xml, содержащий структуру классов квеста. Дальнейшая загрузка невозможна.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If
            Dim fileNameEditor As String
            fileNameEditor = My.Computer.FileSystem.CombinePath(questFolder, "editor.world")
            If My.Computer.FileSystem.FileExists(fileNameEditor) = False Then
                MessageBox.Show("Не найден файл editor.world, содержащий данные для редактора квеста. Дальнейшая загрузка невозможна.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            mScript.LoadClasses(fileName)
            mScript.MakeMainClassHash()
            mScript.FillFuncAndPropHash()

            fileName = My.Computer.FileSystem.CombinePath(questFolder, "game.world")
            If My.Computer.FileSystem.FileExists(fileName) = False Then
                MessageBox.Show("Не найден файл game.world, содержащий все игровые данные. Дальнейшая загрузка невозможна.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If
            If LoadQuestGameFile(fileName) = False Then Return

            cGroups.PrepareGroups()
            frmMainEditor.LoadTrees()
            cPanelManager.btnDefSettings = frmMainEditor.btnShowSettings
            cListManager.UpdateElementsLists()
            cListManager.UpdateFileLists()
            Dim di As New System.IO.DirectoryInfo(questEnvironment.QuestPath)
            frmMainEditor.fsWatcher.Path = di.Root.FullName

            'frmMainEditor.ChangeCurrentClass("Q")
            Dim activePanel As clsPanelManager.clsChildPanel = Nothing
            If LoadQuestEditorFile(fileNameEditor, activePanel) = False Then Return

            DisableEventsDuringLoading = False
            If IsNothing(activePanel) Then
                frmMainEditor.ChangeCurrentClass("Q")
            Else
                cPanelManager.OpenPanel(activePanel)
            End If

            With frmMainEditor
                'заполняем деревья классов 2 уровня (классы 3 уровня всеравно обновляются при переходе между классами)
                Dim curParentName As String = currentParentName
                currentParentName = ""
                For i As Integer = 0 To .dictTrees.Count - 1
                    Dim classId As Integer = .dictTrees.ElementAt(i).Key
                    If mScript.mainClass(classId).LevelsCount = 1 Then .FillTree(mScript.mainClass(classId).Names(0), .dictTrees.ElementAt(i).Value, .chkShowHidden.Checked)
                Next i
                .FillTreeFunctions(.chkShowHidden.Checked)
                .FillTreeVariables(.chkShowHidden.Checked)
                currentParentName = curParentName
            End With
            DisableEventsDuringLoading = True

            With cPanelManager
                If IsNothing(.lstPanels) = False AndAlso .lstPanels.Count > 0 Then
                    .ReassociateAllNodes()
                End If
            End With
            If IsNothing(activePanel) = False Then
                cPanelManager.ActivePanel = Nothing
                cPanelManager.OpenPanel(activePanel)
            End If

            With frmMainEditor
                'удаление старых кнопок классов пользователя
                For i = .ToolStripMain.Items.Count - 1 To 0 Step -1
                    If .ToolStripMain.Items(i).GetType.Name <> "ToolStripButton" Then Continue For
                    Dim itm As ToolStripButton = .ToolStripMain.Items(i)
                    Dim itmClass As String = itm.Tag.ToString, classId As Integer = -1
                    If String.IsNullOrEmpty(itmClass) Then Continue For

                    Dim blnDelete As Boolean = False 'надо ли удалять кнопку
                    If mScript.mainClassHash.TryGetValue(itmClass, classId) = False Then
                        If itmClass <> "Variable" AndAlso itmClass <> "Function" Then blnDelete = True 'класс кнопки не найден - удаляем
                    ElseIf mScript.mainClass(classId).UserAdded Then
                        blnDelete = True 'пользовательский класс. Необходимо обновление - удаляем
                    End If

                    If blnDelete Then
                        .ToolStripMain.Items.Remove(itm)
                        itm.Dispose()
                    End If
                Next i

                'загрузка кнопок классов пользователся
                Dim img As Image = Nothing
                For cId As Integer = 0 To mScript.mainClass.Count - 1
                    If mScript.mainClass(cId).UserAdded Then
                        'добавляем кнопку
                        Dim fPath As String = QuestPath + "\img\classMenu\" + mScript.mainClass(cId).Names(0) + ".png"
                        If My.Computer.FileSystem.FileExists(fPath) Then
                            img = Image.FromFile(fPath)
                        ElseIf My.Computer.FileSystem.FileExists(Application.StartupPath + "\src\img\classMenu\U.png") Then
                            img = Image.FromFile(Application.StartupPath + "\src\img\classMenu\U.png")
                        End If
                        Dim tsb As New ToolStripButton(mScript.mainClass(cId).Names.Last, img, AddressOf .tsbChangeClass) With _
                            {.Tag = mScript.mainClass(cId).Names(0), .DisplayStyle = ToolStripItemDisplayStyle.Image, .ImageScaling = ToolStripItemImageScaling.None}
                        .ToolStripMain.Items.Insert(.ToolStripMain.Items.IndexOf(.tsbAddUserClass), tsb)
                    End If
                Next cId
            End With

            If actionsRouter.locationOfCurrentActions.Length > 0 AndAlso currentClassName <> "A" AndAlso currentClassName <> "L" Then
                actionsRouter.SaveActions()
            End If
        End Sub
        Public DisableEventsDuringLoading As Boolean = True

        Public Sub ClearAllData()
            With frmMainEditor
                .codeBoxPanel.Hide()
                .codeBox.Tag = Nothing
                .codeBox.Text = ""
                .WBhelp.Hide()
                .WBhelp.Tag = Nothing
            End With

            'закрываем все владки, очищаем cPanelManager
            With cPanelManager
                If IsNothing(.lstPanels) = False AndAlso .lstPanels.Count > 0 Then
                    For i As Integer = 0 To .lstPanels.Count - 1
                        Dim pnl As clsPanelManager.clsChildPanel = .lstPanels.Last
                        .RemovePanel(pnl, False)
                    Next
                End If
                .ActivePanel = Nothing
                .dictDefContainers.Clear()
                .dictLastPanel.Clear()
            End With

            With mScript
                .csLocalVariables.KillVars()
                .csPublicVariables.KillVars()
                '.eventRouter.lstEvents.Clear()
                .eventRouter.Clear()
                .eventRouter.lastEventId = 0
                .functionsHash.Clear()
                .LAST_ERROR = ""
                Erase .mainClass

                'If IsNothing(.mainClassCopy) Then
                '    .LoadClasses()
                'Else
                '    .mainClass = .mainClassCopy.Clone
                '    ReDim .mainClass(.mainClassCopy.Count - 1)
                '    For i As Integer = 0 To .mainClassCopy.Count - 1
                '        .mainClass(i) = .mainClassCopy(i).Clone
                '    Next i
                'End If
                .LoadClasses()
                ReDim .mainClassCopy(.mainClass.Count - 1)
                For i As Integer = 0 To .mainClass.Count - 1
                    .mainClassCopy(i) = .mainClass(i).Clone
                Next i
                .MakeMainClassHash()
            End With

            actionsRouter.locationOfCurrentActions = ""
            actionsRouter.lstActions.Clear()
            actionsRouter.lstSavedActions.Clear()
            cGroups.dictRemoved.Clear()
            cGroups.dictGroups.Clear()
            removedObjects.lstRemoved.Clear()

            'очистка меню от дополнительных классов
            With frmMainEditor.ToolStripMain
                For i As Integer = .Items.Count - 1 To 0 Step -1
                    Dim itm As ToolStripItem = .Items(i)
                    Dim tg As String = itm.Tag
                    If String.IsNullOrEmpty(tg) Then Continue For
                    If mScript.mainClassHash.ContainsKey(tg) = False AndAlso tg <> "Variable" AndAlso tg <> "Function" Then
                        .Items.Remove(itm)
                    ElseIf itm.GetType.Name = "ToolStripButton" Then
                        Dim btn As ToolStripButton = itm
                        btn.Checked = False
                    End If
                Next
            End With

            currentTreeView = Nothing
            With frmMainEditor
                .imgLstGroupIcons.Images.Clear()

                .Controls.Add(iconMenuElements)
                .Controls.Add(iconMenuGroups)
                If IsNothing(.dictTrees) = False Then
                    For i As Integer = 0 To .dictTrees.Count - 1
                        Dim tree As TreeView = .dictTrees.Last.Value
                        .dictTrees.Remove(.dictTrees.Last.Key)
                        tree.Dispose()
                    Next
                End If
                If IsNothing(.splitTreeContainer) = False Then .splitTreeContainer.Dispose()
            End With

            mScript.trackingProperties.Clear()

            Me.QuestPath = ""
            Me.QuestShortName = ""
            currentClassName = ""
            currentParentName = ""
            'colors
            lstSelectedColors.Clear()
            '        mScript.csPublicVariables.SetVariableInternal("ALL", "-1", 0)

        End Sub

        Dim sep1 As String = "|#" & Chr(1) & "|"
        Dim sep2 As String = "|#" & Chr(2) & "|"
        Dim sep3 As String = "|#" & Chr(3) & "|"
        Dim sep4 As String = "|#" & Chr(4) & "|"
        Dim sep5 As String = "|#" & Chr(5) & "|"
        ''' <summary>
        ''' Сохраняет главный файл игры
        ''' </summary>
        ''' <param name="fName">Имя файла для сохранения</param>
        ''' <returns>False в случае ошибки</returns>
        Private Function SaveQuestGameFile(ByVal fName As String) As Boolean
            On Error GoTo er

            Dim txt As New System.Text.StringBuilder
            '1.Опознавательная строка
            txt.Append("Quest file for World Creator 2.00" & sep1)

            '2. Сохраняем eventId свойств по умолчанию
            Dim firstItem As Boolean = True
            For classId As Integer = 0 To mScript.mainClass.Count - 1
                If IsNothing(mScript.mainClass(classId).Properties) OrElse mScript.mainClass(classId).Properties.Count = 0 Then Continue For
                For propId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(propId).Value
                    If p.eventId > 0 Then
                        If firstItem Then
                            firstItem = False
                        Else
                            txt.Append(sep2)
                        End If
                        txt.Append(classId.ToString & sep3 & mScript.mainClass(classId).Properties.ElementAt(propId).Key & sep3 & p.eventId.ToString)
                    End If
                Next propId
            Next classId
            txt.Append(sep1)

            '3. Сохраняем расширения функций классов и их eventId
            firstItem = True
            For classId As Integer = 0 To mScript.mainClass.Count - 1
                If IsNothing(mScript.mainClass(classId).Functions) OrElse mScript.mainClass(classId).Functions.Count = 0 Then Continue For
                For funcId As Integer = 0 To mScript.mainClass(classId).Functions.Count - 1
                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions.ElementAt(funcId).Value
                    If p.eventId > 0 Then
                        If firstItem Then
                            firstItem = False
                        Else
                            txt.Append(sep2)
                        End If
                        txt.Append(classId.ToString & sep3 & mScript.mainClass(classId).Functions.ElementAt(funcId).Key & sep3 & p.eventId.ToString & sep3 & p.Value)
                    End If
                Next funcId
            Next classId
            txt.Append(sep1)

            '4. Сохраняем свойства 2-го и 3-го уровня
            firstItem = True
            For classId As Integer = 0 To mScript.mainClass.Count - 1
                If IsNothing(mScript.mainClass(classId).Properties) OrElse mScript.mainClass(classId).Properties.Count = 0 OrElse mScript.mainClass(classId).LevelsCount = 0 _
                    OrElse IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Continue For
                For chId As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                    For propId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                        Dim propName As String = mScript.mainClass(classId).Properties.ElementAt(propId).Key
                        Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(chId)(propName)
                        If firstItem Then
                            firstItem = False
                        Else
                            txt.Append(sep2)
                        End If
                        'Класс - child2Id - Имя свойства - eventId - Value - Value 3-го уровня (элемент 0, 1, ...) - eventId  3-го уровня (элемент 0, 1, ...)
                        txt.Append(classId.ToString & sep3 & chId.ToString & sep3 & propName.ToString & sep3 & ch.eventId.ToString & sep3 & ch.Value)
                        If IsNothing(ch.ThirdLevelProperties) = False AndAlso ch.ThirdLevelProperties.Count > 0 Then
                            'сохраняем свойства 3 уровня
                            'txt.Append(sep3 & Join(ch.ThirdLevelProperties, sep3))
                            For ch3Id As Integer = 0 To ch.ThirdLevelProperties.Count - 1
                                txt.Append(sep3 & ch.ThirdLevelProperties(ch3Id))
                            Next ch3Id
                            For ch3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                                txt.Append(sep3 & ch.ThirdLevelEventId(ch3Id).ToString)
                            Next ch3Id
                        End If
                    Next propId
                Next chId
            Next classId
            txt.Append(sep1)

            '5. Сохраняем функции
            If IsNothing(mScript.functionsHash) = False AndAlso mScript.functionsHash.Count > 0 Then
                firstItem = True
                For funcId As Integer = 0 To mScript.functionsHash.Count - 1
                    Dim f As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
                    Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2)
                    End If
                    txt.Append(funcName & sep3 & SerializeExData(f.ValueExecuteDt))
                Next funcId
            End If
            txt.Append(sep1)

            '6. Сохраняем переменные, привязанные к функциям
            If IsNothing(mScript.functionsHash) = False AndAlso mScript.functionsHash.Count > 0 Then
                Dim firstFunc As Boolean = True
                For funcId As Integer = 0 To mScript.functionsHash.Count - 1
                    Dim f As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
                    If IsNothing(f.Variables) Then Continue For
                    If firstFunc = False Then txt.Append(sep2)
                    firstFunc = False
                    txt.Append(mScript.functionsHash.ElementAt(funcId).Key & sep3)

                    firstItem = True
                    For vId As Integer = 0 To f.Variables.Count - 1
                        Dim v As cVariable.variableEditorInfoType = f.Variables.ElementAt(vId).Value
                        Dim vName As String = f.Variables.ElementAt(vId).Key
                        If firstItem Then
                            firstItem = False
                        Else
                            txt.Append(sep3)
                        End If
                        txt.Append(vName & sep4)
                        For j As Integer = 0 To v.arrValues.Count - 1
                            If j > 0 Then txt.Append(sep5)
                            If IsNothing(v.arrValues(j)) Then v.arrValues(j) = ""
                            txt.Append(v.arrValues(j))
                        Next j
                        If IsNothing(v.lstSingatures) = False AndAlso v.lstSingatures.Count > 0 Then
                            txt.Append(sep4)
                            For signId As Integer = 0 To v.lstSingatures.Count - 1
                                If signId > 0 Then txt.Append(sep5)
                                txt.Append(v.lstSingatures.Keys(signId))
                            Next
                            For signId As Integer = 0 To v.lstSingatures.Count - 1
                                txt.Append(sep5)
                                txt.Append(v.lstSingatures.ElementAt(signId).Value.ToString)
                            Next signId
                        End If
                    Next vId
                    'If funcId < mScript.functionsHash.Count - 1 Then txt.Append(sep2)
                Next funcId
            End If
            txt.Append(sep1)


            '7. Сохраняем переменные
            Dim varClass As cVariable = mScript.csPublicVariables
            For i = 1 To 2
                If i = 1 Then
                    varClass = mScript.csPublicVariables
                Else
                    varClass = mScript.csLocalVariables
                End If

                If IsNothing(varClass.lstVariables) = False AndAlso varClass.lstVariables.Count > 0 Then
                    firstItem = True
                    For vId As Integer = 0 To varClass.lstVariables.Count - 1
                        Dim v As cVariable.variableEditorInfoType = varClass.lstVariables.ElementAt(vId).Value
                        Dim vName As String = varClass.lstVariables.ElementAt(vId).Key
                        If firstItem Then
                            firstItem = False
                        Else
                            txt.Append(sep2)
                        End If
                        txt.Append(vName & sep3) ' & Join(v.arrValues, sep4))
                        For j As Integer = 0 To v.arrValues.Count - 1
                            If j > 0 Then txt.Append(sep4)
                            If IsNothing(v.arrValues(j)) Then v.arrValues(j) = ""
                            txt.Append(v.arrValues(j))
                        Next j
                        If IsNothing(v.lstSingatures) = False AndAlso v.lstSingatures.Count > 0 Then
                            txt.Append(sep3) '& Join(v.lstSingatures.Keys.ToArray, sep4))
                            For signId As Integer = 0 To v.lstSingatures.Count - 1
                                If signId > 0 Then txt.Append(sep4)
                                txt.Append(v.lstSingatures.Keys(signId))
                            Next
                            For signId As Integer = 0 To v.lstSingatures.Count - 1
                                'If signId > 0 Then
                                txt.Append(sep4)
                                txt.Append(v.lstSingatures.ElementAt(signId).Value.ToString)
                            Next signId
                        End If
                    Next vId
                End If
                txt.Append(sep1)
            Next i

            '8. Сохраняем события eventRouter
            txt.Append(mScript.eventRouter.lastEventId.ToString & sep1)

            If mScript.eventRouter.IsEventsListEmpty = False Then ' IsNothing(mScript.eventRouter.lstEvents) = False AndAlso mScript.eventRouter.lstEvents.Count > 0 Then
                firstItem = True
                For eId As Integer = 0 To mScript.eventRouter.lstEvents.Count - 1
                    Dim eventId As Integer = mScript.eventRouter.lstEvents.ElementAt(eId).Key
                    Dim exData As List(Of MatewScript.ExecuteDataType) = mScript.eventRouter.lstEvents.ElementAt(eId).Value
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2)
                    End If
                    txt.Append(eventId.ToString & sep3 & SerializeExData(exData))
                Next eId
            End If
            txt.Append(sep1)

            '9. Сохраняем действия из actionsRouter
            txt.Append(actionsRouter.locationOfCurrentActions & sep1)
            '(1)-..(2)-Имя локации-(3)-Имя свойства-(4)-Event Id-(4)-Value-(3)-Имя свойства...-
            '(1) - отделяет всю эту структуру от остального кода. (2) - разделение локаций. (3) - разделение действий. (4) - разделение свойств
            If actionsRouter.hasSavedActions Then
                firstItem = True
                For locId As Integer = 0 To actionsRouter.lstActions.Count - 1
                    Dim lName As String = actionsRouter.lstActions.ElementAt(locId).Key
                    Dim childs() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(locId).Value
                    If IsNothing(childs) OrElse childs.Count = 0 Then Continue For
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2)
                    End If
                    txt.Append(lName)

                    For aId As Integer = 0 To childs.Count - 1
                        txt.Append(sep3) 'разделитель действий
                        For pId As Integer = 0 To childs(aId).Count - 1
                            If pId > 0 Then txt.Append(sep4) 'разделитель действий
                            Dim propName As String = childs(aId).ElementAt(pId).Key
                            Dim ch As MatewScript.ChildPropertiesInfoType = childs(aId).ElementAt(pId).Value
                            txt.Append(propName & sep5 & ch.eventId.ToString & sep5 & ch.Value)
                        Next pId
                    Next aId
                Next locId
            End If
            txt.Append(sep1)

            '10. Сохраняем сохраненные действия из actionsRouter
            '(1)-..(2)-Имя локации-(3)-Имя свойства-(4)-Event Id-(4)-Value-(3)-Имя свойства...-
            '(1) - отделяет всю эту структуру от остального кода. (2) - разделение локаций. (3) - разделение действий. (4) - разделение свойств
            If actionsRouter.lstSavedActions.Count > 0 Then
                firstItem = True
                For saveId As Integer = 0 To actionsRouter.lstSavedActions.Count - 1
                    Dim saveName As String = actionsRouter.lstSavedActions.ElementAt(saveId).Key
                    Dim childs() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstSavedActions.ElementAt(saveId).Value
                    If IsNothing(childs) OrElse childs.Count = 0 Then Continue For
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2)
                    End If
                    txt.Append(saveName)

                    For aId As Integer = 0 To childs.Count - 1
                        txt.Append(sep3) 'разделитель действий
                        For pId As Integer = 0 To childs(aId).Count - 1
                            If pId > 0 Then txt.Append(sep4) 'разделитель действий
                            Dim propName As String = childs(aId).ElementAt(pId).Key
                            Dim ch As MatewScript.ChildPropertiesInfoType = childs(aId).ElementAt(pId).Value
                            txt.Append(propName & sep5 & ch.eventId.ToString & sep5 & ch.Value)
                        Next pId
                    Next aId
                Next saveId
            End If
            txt.Append(sep1)

            '11. Сохраняем события изменения свойств
            txt.AppendLine(mScript.trackingProperties.CreateStringForSaveFile)
            txt.Append(sep1)

            '12. Сохраняем переменные, привязанные к событиям
            If IsNothing(mScript.eventRouter.lstVariables) = False AndAlso mScript.eventRouter.lstVariables.Count > 0 Then
                For eId As Integer = 0 To mScript.eventRouter.lstVariables.Count - 1
                    Dim eventId As Integer = mScript.eventRouter.lstVariables.ElementAt(eId).Key
                    Dim vars As SortedList(Of String, cVariable.variableEditorInfoType) = mScript.eventRouter.lstVariables.ElementAt(eId).Value
                    If eId > 0 Then txt.Append(sep2)
                    txt.Append(eventId.ToString & sep3)

                    firstItem = True
                    For vId As Integer = 0 To vars.Count - 1
                        Dim v As cVariable.variableEditorInfoType = vars.ElementAt(vId).Value
                        Dim vName As String = vars.ElementAt(vId).Key
                        If firstItem Then
                            firstItem = False
                        Else
                            txt.Append(sep3)
                        End If
                        txt.Append(vName & sep4)
                        For j As Integer = 0 To v.arrValues.Count - 1
                            If j > 0 Then txt.Append(sep5)
                            If IsNothing(v.arrValues(j)) Then v.arrValues(j) = ""
                            txt.Append(v.arrValues(j))
                        Next j
                        If IsNothing(v.lstSingatures) = False AndAlso v.lstSingatures.Count > 0 Then
                            txt.Append(sep4)
                            For signId As Integer = 0 To v.lstSingatures.Count - 1
                                If signId > 0 Then txt.Append(sep5)
                                txt.Append(v.lstSingatures.Keys(signId))
                            Next
                            For signId As Integer = 0 To v.lstSingatures.Count - 1
                                txt.Append(sep5)
                                txt.Append(v.lstSingatures.ElementAt(signId).Value.ToString)
                            Next signId
                        End If
                    Next vId
                Next eId
            End If


            My.Computer.FileSystem.WriteAllText(fName, txt.ToString, False)
            txt.Clear()
            Return True
er:
            MessageBox.Show(Err.Description, "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Function

        ''' <summary>
        ''' Сохраняет файл с информацией для редактора
        ''' </summary>
        ''' <param name="fName">Имя файла для сохранения</param>
        ''' <returns>False в случае ошибки</returns>
        Private Function SaveQuestEditorFile(ByVal fName As String) As Boolean
            On Error GoTo er

            Dim txt As New System.Text.StringBuilder
            '1.Опознавательная строка
            txt.Append("Editor file for World Creator 2.00" & sep1)

            '2. Сохраняем свойства Hidden элементов 2 уровня
            Dim firstItem As Boolean = True
            For classId As Integer = 0 To mScript.mainClass.Count - 1
                If IsNothing(mScript.mainClass(classId).Properties) OrElse mScript.mainClass(classId).Properties.Count = 0 OrElse mScript.mainClass(classId).LevelsCount = 0 _
                    OrElse IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Continue For
                For chId As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                    For propId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                        Dim propName As String = mScript.mainClass(classId).Properties.ElementAt(propId).Key
                        If Not mScript.mainClass(classId).ChildProperties(chId)(propName).Hidden Then Continue For
                        If firstItem Then
                            firstItem = False
                        Else
                            txt.Append(sep2)
                        End If
                        'Класс - child2Id - Имя свойства (тольк те, где Hidden = True)
                        txt.Append(classId.ToString & sep3 & chId.ToString & sep3 & propName.ToString)
                    Next propId
                Next chId
            Next classId
            txt.Append(sep1)

            '3. Сохраняем данные о функциях
            If IsNothing(mScript.functionsHash) = False AndAlso mScript.functionsHash.Count > 0 Then
                firstItem = True
                For funcId As Integer = 0 To mScript.functionsHash.Count - 1
                    Dim f As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(funcId).Value
                    Dim funcName As String = mScript.functionsHash.ElementAt(funcId).Key
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2)
                    End If
                    txt.Append(funcName & sep3 & f.Description & sep3 & f.Group & sep3 & f.Hidden.ToString & sep3 & f.Icon & sep3 & codeBoxShadowed.codeBox.SerializeCodeData(f.ValueDt))
                Next funcId
            End If
            txt.Append(sep1)

            '4. Сохраняем данные о переменных
            Dim varClass As cVariable = mScript.csPublicVariables
            For i = 1 To 2
                If i = 1 Then
                    varClass = mScript.csPublicVariables
                Else
                    varClass = mScript.csLocalVariables
                End If

                If IsNothing(varClass.lstVariables) = False AndAlso varClass.lstVariables.Count > 0 Then
                    firstItem = True
                    For vId As Integer = 0 To varClass.lstVariables.Count - 1
                        Dim v As cVariable.variableEditorInfoType = varClass.lstVariables.ElementAt(vId).Value
                        Dim vName As String = varClass.lstVariables.ElementAt(vId).Key
                        If firstItem Then
                            firstItem = False
                        Else
                            txt.Append(sep2)
                        End If
                        txt.Append(vName & sep3 & v.Description & sep3 & v.Group & sep3 & v.Icon & sep3 & v.Hidden.ToString)
                    Next vId
                End If
                txt.Append(sep1)
            Next i

            '5. Сохраняем Hidden действий из actionsRouter
            '(1)-..(2)-Имя локации-(3)-Имя свойства-(4)-Event Id-(4)-Value-(3)-Имя свойства...-
            '(1) - отделяет всю эту структуру от остального кода. (2) - разделение локаций. (3) - разделение действий. (4) - разделение свойств
            If actionsRouter.hasSavedActions Then
                firstItem = True
                For locId As Integer = 0 To actionsRouter.lstActions.Count - 1
                    Dim lName As String = actionsRouter.lstActions.ElementAt(locId).Key
                    Dim childs() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = actionsRouter.lstActions.ElementAt(locId).Value
                    If IsNothing(childs) OrElse childs.Count = 0 Then Continue For
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2)
                    End If
                    txt.Append(lName)

                    For aId As Integer = 0 To childs.Count - 1
                        txt.Append(sep3) 'разделитель действий
                        Dim firstItem2 As Boolean = True
                        For pId As Integer = 0 To childs(aId).Count - 1
                            Dim ch As MatewScript.ChildPropertiesInfoType = childs(aId).ElementAt(pId).Value
                            If Not ch.Hidden Then Continue For
                            If firstItem2 Then
                                firstItem2 = False
                            Else
                                txt.Append(sep4) 'разделитель свойств
                            End If
                            txt.Append(childs(aId).ElementAt(pId).Key)
                        Next pId
                    Next aId
                Next locId
            End If
            txt.Append(sep1)

            '6. Сохраняем группы
            If IsNothing(cGroups.dictGroups) = False AndAlso cGroups.dictGroups.Count > 0 Then
                firstItem = True
                For cId As Integer = 0 To cGroups.dictGroups.Count - 1
                    Dim lstGroups As List(Of clsGroups.clsGroupInfo) = cGroups.dictGroups.ElementAt(cId).Value
                    If IsNothing(lstGroups) OrElse lstGroups.Count = 0 Then Continue For
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2) 'разделяет классы
                    End If
                    txt.Append(cGroups.dictGroups.ElementAt(cId).Key)
                    For gId As Integer = 0 To lstGroups.Count - 1
                        Dim g As clsGroups.clsGroupInfo = lstGroups(gId)
                        txt.Append(sep3 & g.Name & sep4 & g.Hidden.ToString & sep4 & g.iconName & sep4 & g.isThirdLevelGroup.ToString & sep4 & g.parentName)
                    Next gId
                Next cId
            End If
            txt.Append(sep1)

            '7. Сохраняем цвета
            If IsNothing(lstSelectedColors) = False AndAlso lstSelectedColors.Count > 0 Then
                firstItem = True
                For cId As Integer = 0 To lstSelectedColors.Count - 1
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2) 'разделяет классы
                    End If
                    txt.Append(lstSelectedColors(cId))
                Next cId
            End If
            txt.Append(sep1)

            '8. Сохраняем вкладки
            If IsNothing(cPanelManager) = False AndAlso cPanelManager.lstPanels.Count > 0 Then
                firstItem = True
                For pId As Integer = 0 To cPanelManager.lstPanels.Count - 1
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2) 'разделяет вкладки
                    End If
                    Dim ch As clsPanelManager.clsChildPanel = cPanelManager.lstPanels(pId)
                    Dim actControlName As String = ""
                    If IsNothing(ch.lcActiveControl) = False Then actControlName = ch.lcActiveControl.Name
                    txt.Append(actControlName & sep3 & ch.child2Name & sep3 & ch.child3Name & sep3 & ch.classId.ToString & sep3 & ch.IsCodeBoxVisible.ToString & sep3 & ch.IsWbVisible.ToString & sep3 & ch.supraElementName)
                Next pId
            End If
            txt.Append(sep1)

            'текущая
            If IsNothing(cPanelManager.ActivePanel) = False Then
                txt.Append(cPanelManager.lstPanels.IndexOf(cPanelManager.ActivePanel).ToString)
            End If
            txt.Append(sep1)

            '9. Сохраняем lastPanels
            If IsNothing(cPanelManager.dictLastPanel) = False AndAlso cPanelManager.dictLastPanel.Count > 0 Then
                firstItem = True
                For cId As Integer = 0 To cPanelManager.dictLastPanel.Count - 1
                    Dim ch As clsPanelManager.clsChildPanel = cPanelManager.dictLastPanel.ElementAt(cId).Value
                    If IsNothing(ch) Then Continue For
                    Dim classId As Integer = cPanelManager.dictLastPanel.ElementAt(cId).Key
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2) 'разделяет вкладки
                    End If
                    txt.Append(classId.ToString & sep3 & cPanelManager.lstPanels.IndexOf(ch).ToString)
                Next cId
            End If
            txt.Append(sep1)

            '10. chkShowHidden
            txt.Append(frmMainEditor.chkShowHidden.Checked.ToString & sep1)
            '11. Codebox HightLight
            If frmMainEditor.tsbHighLightFull.Checked Then
                txt.Append("0")
            ElseIf frmMainEditor.tsbHighLightDesignate.Checked Then
                txt.Append("1")
            Else
                txt.Append("2")
            End If
            txt.Append(sep1)

            '12. showBasicFunctions
            firstItem = True
            For pId As Integer = 0 To cPanelManager.dictDefContainers.Count - 1
                Dim pEx As clsPanelManager.PanelEx = cPanelManager.dictDefContainers.ElementAt(pId).Value
                If pEx.showBasicFunctions Then
                    If firstItem Then
                        firstItem = False
                    Else
                        txt.Append(sep2)
                    End If
                    txt.Append(cPanelManager.dictDefContainers.ElementAt(pId).Key.ToString)
                End If
            Next pId

            My.Computer.FileSystem.WriteAllText(fName, txt.ToString, False)
            txt.Clear()
            Return True
er:
            MessageBox.Show(Err.Description, "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Function

        ''' <summary>
        ''' Загружает главный файл игры
        ''' </summary>
        ''' <param name="fName">Путь к файлу</param>
        Public Function LoadQuestGameFile(ByVal fName As String) As Boolean
            On Error GoTo er
            Dim fileContent As String = My.Computer.FileSystem.ReadAllText(fName)
            Dim arr1() As String = Split(fileContent, sep1)
            '0 - опознавательная строка
            '1 - eventId свойств по умолчанию
            '2 - расширения функций классов и их eventId
            '3 - свойства 2-го и 3-го уровня
            '4 - функции
            '5 - переменные, привязанные к функциям
            '6 - переменные глобальные
            '7 - переменные локальные
            '8 - lastId eventRouter
            '9 - события eventRouter
            '10 - actionsRouter.locationOfCurrentActions
            '11 - действия из actionsRouter
            '12 - сохраненные действия из actionsRouter
            '13 - события изменения свойств
            '14 - переменные, привязанные к событиям

            If arr1.Count <> 15 OrElse arr1(0).StartsWith("Quest file for World Creator ", StringComparison.CurrentCultureIgnoreCase) = False Then
                MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            '1 - eventId свойств по умолчанию
            Dim arr2() As String = Split(arr1(1), sep2)
            'каждый arr2 содержит classId-3-Имя свойства-3-eventId
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                For i As Integer = 0 To arr2.Length - 1
                    Dim arr3() As String = Split(arr2(i), sep3)
                    If arr3.Length <> 3 Then
                        MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                    Dim classId As Integer = CInt(arr3(0))
                    Dim propName As String = arr3(1)
                    Dim eventId As Integer = CInt(arr3(2))
                    mScript.mainClass(classId).Properties(propName).eventId = eventId
                Next i
            End If

            '2 - расширения функций классов и их eventId
            arr2 = Split(arr1(2), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                'каждый arr2 содержит classId-3-Имя функции-3-eventId-3-Value
                For i As Integer = 0 To arr2.Length - 1
                    Dim arr3() As String = Split(arr2(i), sep3)
                    If arr3.Length <> 4 Then
                        MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                    Dim classId As Integer = CInt(arr3(0))
                    Dim funcName As String = arr3(1)
                    Dim eventId As Integer = CInt(arr3(2))
                    mScript.mainClass(classId).Functions(funcName).eventId = eventId
                    mScript.mainClass(classId).Functions(funcName).Value = arr3(3)
                Next i
            End If

            '3 - свойства 2-го и 3-го уровня
            arr2 = Split(arr1(3), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                'каждый arr2 содержит classId-3-child2Id-3-Имя свойства-3-eventId-3-Value
                'если у свойства есть элементы 3 уровня то далее следуют Value1-3-Value2-3-...ValueX-3-eventId1-3-eventId2-3-...-eventIdX
                'ReDim mScript.mainClass(classId).ChildProperties(arr2.Length - 1)
                For i As Integer = 0 To arr2.Length - 1
                    Dim arr3() As String = Split(arr2(i), sep3)
                    If arr3.Length < 5 Then
                        MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                    Dim classId As Integer = CInt(arr3(0))
                    Dim child2Id As Integer = CInt(arr3(1))
                    Dim propName As String = arr3(2)
                    Dim eventId As Integer = CInt(arr3(3))

                    Dim ch As New MatewScript.ChildPropertiesInfoType '= mScript.mainClass(classId).ChildProperties(child2Id)(propName)
                    ch.eventId = eventId
                    ch.Value = arr3(4)
                    If arr3.Length > 5 Then
                        'есть элементы 3 уровня
                        Dim cnt As Integer = (arr3.Length - 5) / 2 'кол-во элементов 3 уровня
                        ReDim ch.ThirdLevelEventId(cnt - 1)
                        ReDim ch.ThirdLevelProperties(cnt - 1)
                        For j As Integer = 5 To 5 + cnt - 1
                            ch.ThirdLevelProperties(j - 5) = arr3(j)
                        Next j
                        For j As Integer = 5 + cnt To arr3.Length - 1
                            ch.ThirdLevelEventId(j - 5 - cnt) = arr3(j)
                        Next j
                    End If
                    If IsNothing(mScript.mainClass(classId).ChildProperties) Then
                        ReDim mScript.mainClass(classId).ChildProperties(child2Id)
                        mScript.mainClass(classId).ChildProperties(child2Id) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
                    ElseIf child2Id > mScript.mainClass(classId).ChildProperties.Count - 1 Then
                        ReDim Preserve mScript.mainClass(classId).ChildProperties(child2Id)
                        mScript.mainClass(classId).ChildProperties(child2Id) = New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
                    End If
                    mScript.mainClass(classId).ChildProperties(child2Id).Add(propName, ch)
                Next i
            End If

            '4 - функции
            arr2 = Split(arr1(4), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                'каждый arr2 содержит Имя функции-3-SerializeExData(f.ValueExecuteDt)
                For i As Integer = 0 To arr2.Count - 1
                    Dim arr3() As String = Split(arr2(i), sep3)
                    If arr3.Length <> 2 Then
                        MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                    Dim funcName As String = arr3(0)
                    Dim f As New MatewScript.FunctionInfoType With {.ValueExecuteDt = DeserializeExData(arr3(1))}
                    mScript.functionsHash.Add(funcName, f)
                Next i
            End If

            '5 переменные, привязанные к функциям
            arr2 = Split(arr1(5), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                'в arr2 - список функций
                Dim arr3() As String
                For fId As Integer = 0 To arr2.Count - 1
                    Dim lstVars As New SortedList(Of String, cVariable.variableEditorInfoType)(StringComparer.CurrentCultureIgnoreCase)
                    arr3 = Split(arr2(fId), sep3)
                    'каждый arr3 содержит Имя переменной-4-массив arrValues(-5-)-4-массив названий сигнатур(-5-) и соответствующих им индексов(-5-)
                    For i As Integer = 1 To arr3.Count - 1
                        Dim arr4() As String = Split(arr3(i), sep4)
                        If arr4.Length < 2 Then
                            MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return False
                        End If
                        Dim vName As String = arr4(0)
                        Dim v As New cVariable.variableEditorInfoType
                        v.arrValues = Split(arr4(1), sep5)
                        If arr4.Length > 2 Then
                            'имеются сигнатуры
                            Dim signAll() As String = Split(arr4(2), sep5)
                            Dim cnt As Integer = signAll.Length / 2
                            v.lstSingatures = New SortedList(Of String, Integer)
                            For j As Integer = 0 To cnt - 1
                                v.lstSingatures.Add(signAll(j), signAll(j + cnt))
                            Next j
                        End If
                        lstVars.Add(vName, v)
                    Next i
                    mScript.functionsHash(arr3(0)).Variables = lstVars
                Next fId
            End If


            '6,7 - переменные глобальные/локальные
            For q = 6 To 7
                arr2 = Split(arr1(q), sep2)
                If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                    'каждый arr2 содержит Имя переменной-3-массив arrValues(-4-)-3-массив названий сигнатур(-4-) и соответствующих им индексов(-4-)
                    For i As Integer = 0 To arr2.Count - 1
                        Dim arr3() As String = Split(arr2(i), sep3)
                        If arr3.Length < 2 Then
                            MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return False
                        End If
                        Dim vName As String = arr3(0)
                        Dim v As New cVariable.variableEditorInfoType
                        v.arrValues = Split(arr3(1), sep4)
                        If arr3.Length > 2 Then
                            'имеются сигнатуры
                            Dim signAll() As String = Split(arr3(2), sep4)
                            Dim cnt As Integer = signAll.Length / 2
                            v.lstSingatures = New SortedList(Of String, Integer)
                            For j As Integer = 0 To cnt - 1
                                v.lstSingatures.Add(signAll(j), signAll(j + cnt))
                            Next j
                        End If
                        If q = 6 Then
                            mScript.csPublicVariables.lstVariables.Add(vName, v)
                        Else
                            mScript.csLocalVariables.lstVariables.Add(vName, v)
                        End If
                    Next i
                End If
            Next q

            '8 - lastId eventRouter
            mScript.eventRouter.Clear()
            mScript.eventRouter.lastEventId = CInt(arr1(8))

            '9 - события eventRouter
            arr2 = Split(arr1(9), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                'каждый arr2 содержит eventId-3-SerializeExData(exData)
                For i As Integer = 0 To arr2.Count - 1
                    Dim arr3() As String = Split(arr2(i), sep3)
                    If arr3.Length <> 2 Then
                        MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                    Dim eventId As Integer = CInt(arr3(0))
                    mScript.eventRouter.SetEventId(eventId, DeserializeExData(arr3(1)))
                    'mScript.eventRouter.lstEvents.Add(eventId, DeserializeExData(arr3(1)))
                Next i
            End If

            '10 - actionsRouter.locationOfCurrentActions
            actionsRouter.locationOfCurrentActions = arr1(10)
            '11 - действия из actionsRouter
            arr2 = Split(arr1(11), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                '(1)-..(2)-Имя локации-(3)-Имя свойства-(4)-Event Id-(4)-Value-(3)-Имя свойства...-
                '(1) - отделяет всю эту структуру от остального кода. (2) - разделение локаций. (3) - разделение действий. (4) - разделение свойств
                For i As Integer = 0 To arr2.Count - 1
                    'arr2 - -2-имя локации-3-все ее действия-2-
                    Dim arr3() As String = Split(arr2(i), sep3) '0 - имя локации, 1,2.., Х - данные о действиях
                    Dim locName As String = arr3(0)
                    Dim childs() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing
                    If arr3.Length > 1 Then
                        ReDim childs(arr3.Length - 2)
                        For aId As Integer = 0 To arr3.Length - 2
                            'arr3 - имя локации-4-действия
                            Dim arr4() As String = Split(arr3(aId + 1), sep4)
                            'каждый arr4 содержит наборы свойств действия Id
                            Dim actCh As New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase) 'набор свойств действия
                            For pId As Integer = 0 To arr4.Length - 1
                                Dim arr5() As String = Split(arr4(pId), sep5)
                                'arr5 - имя свойства-4-eventId-4-value
                                If arr5.Length <> 3 Then
                                    MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Return False
                                End If
                                Dim propName As String = arr5(0)
                                Dim eventId As Integer = arr5(1)
                                Dim ch As New MatewScript.ChildPropertiesInfoType With {.eventId = eventId, .Value = arr5(2)}
                                actCh.Add(propName, ch)
                            Next pId
                            childs(aId) = actCh
                        Next aId
                    End If
                    actionsRouter.lstActions.Add(locName, childs)
                Next i
            End If

            '12 - сохраненные действия из actionsRouter
            arr2 = Split(arr1(12), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                '(1)-..(2)-Имя локации-(3)-Имя свойства-(4)-Event Id-(4)-Value-(3)-Имя свойства...-
                '(1) - отделяет всю эту структуру от остального кода. (2) - разделение локаций. (3) - разделение действий. (4) - разделение свойств
                For i As Integer = 0 To arr2.Count - 1
                    'arr2 - -2-имя локации-3-все ее действия-2-
                    Dim arr3() As String = Split(arr2(i), sep3) '0 - имя локации, 1,2.., Х - данные о действиях
                    Dim saveName As String = arr3(0)
                    Dim childs() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = Nothing
                    If arr3.Length > 1 Then
                        ReDim childs(arr3.Length - 2)
                        For aId As Integer = 0 To arr3.Length - 2
                            'arr3 - имя локации-4-действия
                            Dim arr4() As String = Split(arr3(aId + 1), sep4)
                            'каждый arr4 содержит наборы свойств действия Id
                            Dim actCh As New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase) 'набор свойств действия
                            For pId As Integer = 0 To arr4.Length - 1
                                Dim arr5() As String = Split(arr4(pId), sep5)
                                'arr5 - имя свойства-4-eventId-4-value
                                If arr5.Length <> 3 Then
                                    MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Return False
                                End If
                                Dim propName As String = arr5(0)
                                Dim eventId As Integer = arr5(1)
                                Dim ch As New MatewScript.ChildPropertiesInfoType With {.eventId = eventId, .Value = arr5(2)}
                                actCh.Add(propName, ch)
                            Next pId
                            childs(aId) = actCh
                        Next aId
                    End If
                    actionsRouter.lstSavedActions.Add(saveName, childs)
                Next i
            End If

            '13 - события изменения свойств
            mScript.trackingProperties.LoadDataFromSaveString(arr1(13))

            '14 переменные, привязанные к событиям
            arr2 = Split(arr1(14), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                'в arr2 - список событий
                Dim arr3() As String
                For eId As Integer = 0 To arr2.Count - 1
                    Dim lstVars As New SortedList(Of String, cVariable.variableEditorInfoType)(StringComparer.CurrentCultureIgnoreCase)
                    arr3 = Split(arr2(eId), sep3)
                    'каждый arr3 содержит Имя переменной-4-массив arrValues(-5-)-4-массив названий сигнатур(-5-) и соответствующих им индексов(-5-)
                    If String.IsNullOrEmpty(arr3(1)) Then Continue For
                    For i As Integer = 1 To arr3.Count - 1
                        Dim arr4() As String = Split(arr3(i), sep4)
                        If arr4.Length < 2 Then
                            MessageBox.Show("Главный файл игры имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return False
                        End If
                        Dim vName As String = arr4(0)
                        Dim v As New cVariable.variableEditorInfoType
                        v.arrValues = Split(arr4(1), sep5)
                        If arr4.Length > 2 Then
                            'имеются сигнатуры
                            Dim signAll() As String = Split(arr4(2), sep5)
                            Dim cnt As Integer = signAll.Length / 2
                            v.lstSingatures = New SortedList(Of String, Integer)
                            For j As Integer = 0 To cnt - 1
                                v.lstSingatures.Add(signAll(j), signAll(j + cnt))
                            Next j
                        End If
                        lstVars.Add(vName, v)
                    Next i
                    mScript.eventRouter.lstVariables.Add(arr3(0), lstVars)
                Next eId
            End If

            Return True
er:
            MessageBox.Show(Err.Description, "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Function


        Public Function LoadQuestEditorFile(ByVal fName As String, ByRef activePanel As clsPanelManager.clsChildPanel) As Boolean
            On Error GoTo er
            Dim fileContent As String = My.Computer.FileSystem.ReadAllText(fName)
            Dim arr1() As String = Split(fileContent, sep1)
            '1 - опознавательная строка
            '2 - свойства Hidden элементов 2 уровня
            '3 - данные о функциях
            '4 - переменные глобальные
            '5 - переменные локальные
            '6 - Hidden действий из actionsRouter
            '7 - группы
            '8 - цвета
            '9 - вкладки
            '10 - текущая вкладка
            '11 - lastPanels
            '12 - chkShowHidden
            '13 - Codebox HightLight
            '14. showBasicFunctions

            If arr1.Count <> 14 OrElse arr1(0).StartsWith("Editor file for World Creator ", StringComparison.CurrentCultureIgnoreCase) = False Then
                MessageBox.Show("Файл квеста с информацией для редактора имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            '2 - свойства Hidden элементов 2 уровня
            Dim arr2() As String = Split(arr1(1), sep2)
            'каждый arr2 содержит Класс -3- child2Id -3- Имя свойства (тольк те, где Hidden = True)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                For i As Integer = 0 To arr2.Length - 1
                    Dim arr3() As String = Split(arr2(i), sep3)
                    If arr3.Length <> 3 Then
                        MessageBox.Show("Файл квеста с информацией для редактора имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                    Dim classId As Integer = CInt(arr3(0))
                    Dim child2Id As Integer = CInt(arr3(1))
                    Dim propName As String = arr3(2)
                    mScript.mainClass(classId).ChildProperties(child2Id)(propName).Hidden = True
                Next i
            End If

            '3. данные о функциях
            arr2 = Split(arr1(2), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                For i As Integer = 0 To arr2.Length - 1
                    Dim arr3() As String = Split(arr2(i), sep3)
                    If arr3.Length <> 6 Then
                        MessageBox.Show("Файл квеста с информацией для редактора имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                    'каждый arr2 содержит: имя функции-3-описание-3-группу-3-hidden-3-иконку-3-ValueDt
                    Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(arr3(0))
                    f.Description = arr3(1)
                    f.Group = arr3(2)
                    f.Hidden = CBool(arr3(3))
                    f.Icon = arr3(4)
                    f.ValueDt = codeBoxShadowed.codeBox.DeserializeCodeData(arr3(5))
                Next i
            End If

            '4, 5. переменные
            For i = 3 To 4
                arr2 = Split(arr1(i), sep2)
                Dim varClass As cVariable
                If i = 3 Then
                    varClass = mScript.csPublicVariables
                Else
                    varClass = mScript.csLocalVariables
                End If
                If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                    For j As Integer = 0 To arr2.Length - 1
                        Dim arr3() As String = Split(arr2(j), sep3)
                        If arr3.Length <> 5 Then
                            MessageBox.Show("Файл квеста с информацией для редактора имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return False
                        End If
                        'каждый arr2 содержит: имя переменной-3-описание-3-группу-3-иконку-3-hidden
                        Dim v As cVariable.variableEditorInfoType = varClass.lstVariables(arr3(0))
                        v.Description = arr3(1)
                        v.Group = arr3(2)
                        v.Icon = arr3(3)
                        v.Hidden = CBool(arr3(4))
                    Next j
                End If
            Next i

            '6 - Hidden действий из actionsRouter
            arr2 = Split(arr1(5), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                For i As Integer = 0 To arr2.Length - 1
                    Dim arr3() As String = Split(arr2(i), sep3) '0 - локация, остальные - действия
                    Dim locName As String = arr3(0)
                    For actId As Integer = 0 To arr3.Length - 2
                        Dim arr4() As String = Split(arr3(actId + 1), sep4) 'свойства, которые hidden
                        If (arr4.Length = 1 AndAlso String.IsNullOrEmpty(arr4(0))) = False Then
                            For pId As Integer = 0 To arr4.Length - 1
                                actionsRouter.lstActions(locName)(actId)(arr4(pId)).Hidden = True
                            Next pId
                        End If
                    Next actId
                Next i
            End If

            '7 - группы
            arr2 = Split(arr1(6), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                'arr2 содержит группы разных классов
                For i As Integer = 0 To arr2.Length - 1
                    Dim arr3() As String = Split(arr2(i), sep3) '0 - имя класса, 1 - (имя группы-4-hidden-4-иконка-4-isThirdLevelGroup-4-parentName)
                    Dim className As String = arr3(0)
                    For j As Integer = 0 To arr3.Length - 2
                        Dim arr4() As String = Split(arr3(j + 1), sep4)
                        If arr4.Length = 1 AndAlso String.IsNullOrEmpty(arr4(0)) Then Continue For
                        If arr4.Length <> 5 Then
                            MessageBox.Show("Файл квеста с информацией для редактора имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return False
                        End If
                        Dim g As New clsGroups.clsGroupInfo '= cGroups.dictGroups(className)(j)
                        g.Name = arr4(0)
                        g.Hidden = CBool(arr4(1))
                        g.iconName = arr4(2)
                        g.isThirdLevelGroup = CBool(arr4(3))
                        g.parentName = arr4(4)
                        If cGroups.dictGroups.ContainsKey(className) = False Then cGroups.dictGroups.Add(className, New List(Of clsGroups.clsGroupInfo))
                        cGroups.dictGroups(className).Add(g)
                    Next j
                Next i
            End If

            '8 - цвета
            arr2 = Split(arr1(7), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                lstSelectedColors.AddRange(arr2)
            End If

            '9 - вкладки
            arr2 = Split(arr1(8), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                'каждый arr2 содержит: activeControlName-3-child2Name-3-child3Name-3-classId-3-IsCodeBoxVisible-3-IsWbVisible-3-supraElementName
                For i As Integer = 0 To arr2.Length - 1
                    Dim arr3() As String = Split(arr2(i), sep3)
                    If arr3.Length <> 7 Then
                        MessageBox.Show("Файл квеста с информацией для редактора имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                    Dim child2Name As String = arr3(1)
                    Dim child3Name As String = arr3(2)
                    Dim classId As Integer = CInt(arr3(3))
                    Dim IsCBvisible As Boolean = CBool(arr3(4))
                    Dim IsWBvisible As Boolean = CBool(arr3(5))
                    Dim supra As String = arr3(6)
                    Dim ch As New clsPanelManager.clsChildPanel(classId, child2Name, child3Name, Nothing, supra) With {.IsCodeBoxVisible = IsCBvisible, .IsWbVisible = IsWBvisible}
                    Dim cName As String = arr3(0)
                    If String.IsNullOrEmpty(cName) = False Then
                        If cPanelManager.dictDefContainers.ContainsKey(classId) = False Then cPanelManager.CreatePropertiesControl(classId)
                        Dim pEx As clsPanelManager.PanelEx = cPanelManager.dictDefContainers(classId)
                        Dim c() As Control = pEx.Controls.Find(cName, True)
                        If c.Length > 0 Then ch.lcActiveControl = c(0)
                    End If
                    cPanelManager.lstPanels.Add(ch)
                    cPanelManager.CreateToolButton(ch)
                    ch.toolButton.Checked = False
                Next i
            End If

            '10 - текущая вкладка
            arr2 = Split(arr1(9), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                Dim pId As Integer = CInt(arr2(0))
                If pId >= 0 Then
                    activePanel = cPanelManager.lstPanels(pId)
                End If
            End If

            '11 - lastPanels
            arr2 = Split(arr1(10), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                'каждый arr2 содержит: classId-3-Id панели в lstPanels
                For i As Integer = 0 To arr2.Length - 1
                    Dim arr3() As String = Split(arr2(i), sep3)
                    If arr3.Length <> 2 Then
                        MessageBox.Show("Файл квеста с информацией для редактора имеет неверный формат. Загрузка невозможна.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End If
                    Dim classId As Integer = CInt(arr3(0))
                    Dim pId As Integer = CInt(arr3(1))
                    If pId = -1 Then Continue For
                    If cPanelManager.dictLastPanel.ContainsKey(classId) Then
                        cPanelManager.dictLastPanel(classId) = cPanelManager.lstPanels(pId)
                    Else
                        cPanelManager.dictLastPanel.Add(classId, cPanelManager.lstPanels(pId))
                    End If
                Next i
            End If

            '12. chkShowHidden
            frmMainEditor.chkShowHidden.Checked = CBool(arr1(11))

            '13. Codebox HightLight
            Select Case CInt(arr1(12))
                Case 0
                    frmMainEditor.tsmiHighLightFull.PerformClick()
                Case 1
                    frmMainEditor.tsmiHighLightDesignate.PerformClick()
                Case 2
                    frmMainEditor.tsmiHighLightNo.PerformClick()
            End Select

            '14. showBasicFunctions
            arr2 = Split(arr1(13), sep2)
            If (arr2.Length = 1 AndAlso String.IsNullOrEmpty(arr2(0))) = False Then
                For i As Integer = 0 To arr2.Length - 1
                    Dim cId As Integer = CInt(arr2(i))
                    If cPanelManager.dictDefContainers.ContainsKey(cId) Then
                        Dim pEx As clsPanelManager.PanelEx = cPanelManager.dictDefContainers(cId)
                        cPanelManager.dictDefContainers.Remove(cId)
                        pEx.Dispose()
                    End If
                    cPanelManager.CreatePropertiesControl(cId, True)
                Next i
            End If


            Return True
er:
            MessageBox.Show(Err.Description, "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Function

        ''' <summary>
        ''' сериализует тип ExecuteDataType
        ''' </summary>
        ''' <param name="exData">Список ExecuteDataType для сериализации</param>
        ''' <returns>Сериализованную строку</returns>
        Private Function SerializeExData(ByRef exData As List(Of MatewScript.ExecuteDataType)) As String
            If IsNothing(exData) OrElse exData.Count = 0 Then Return ""
            If exData.Count = 1 Then
                If IsNothing(exData(0)) OrElse IsNothing(exData(0).Code) OrElse exData(0).Code.Length = 0 Then
                    Return ""
                End If
            End If
            Dim sb As New System.Text.StringBuilder

            'Dim arrExData() As MatewScript.ExecuteDataType = exData.ToArray
            Using xmlWriter As System.Xml.XmlWriter = System.Xml.XmlWriter.Create(sb, New System.Xml.XmlWriterSettings With {.OmitXmlDeclaration = True})
                Dim x As New System.Xml.Serialization.XmlSerializer(exData.GetType)
                x.Serialize(xmlWriter, exData)
            End Using
            Return sb.ToString

        End Function

        ''' <summary>
        ''' десериализует тип ExecuteDataType
        ''' </summary>
        ''' <param name="strXML">Строка для десериализации</param>
        Public Function DeserializeExData(ByRef strXML As String) As List(Of MatewScript.ExecuteDataType)
            If String.IsNullOrEmpty(strXML) Then
                Dim exData As New List(Of MatewScript.ExecuteDataType)
                Return exData
            End If

            Using xmlMemorySteam As New System.IO.MemoryStream
                Using xmlSteamWriter As New System.IO.StreamWriter(xmlMemorySteam)
                    xmlSteamWriter.Write(strXML)
                    xmlSteamWriter.Flush()
                End Using

                Using xmlMemorySteam2 As New System.IO.MemoryStream(xmlMemorySteam.GetBuffer)
                    Using xmlStreamReader As New System.IO.StreamReader(xmlMemorySteam2)
                        Dim x As New System.Xml.Serialization.XmlSerializer(GetType(List(Of MatewScript.ExecuteDataType)))
                        Return x.Deserialize(xmlMemorySteam2)
                    End Using
                End Using
            End Using
        End Function

    End Class

    Public DEFAULT_COLORS As New clsDEFAULT_COLORS
    Public provider_points As New System.Globalization.NumberFormatInfo 'настройки отображения чисел с разделителем точкой (3.14 вместо 3,14)
    Public hCurrentDocument As HtmlDocument '???
    Public questEnvironment As New clsQuestEnvironment

    ''' <summary>Имя класса, с которым идет работа в данный момент</summary>
    Public Property currentClassName As String
    ''' <summary>Имя элемента 2-го порядка, с детьми которого идет работа (в случае, если сейчас идет работа с элементами 3-го порядка, иначе равно "")
    ''' То есть, при работе с элементами 2-го уровня БУДЕТ РАВНА "". Исключения - это дочерние элементы, по сути являющиеся элементами 2-го порядка 
    ''' (пока что это только действия - дочерние для локаций, но являются классом 2 уровня)</summary>
    Public Property currentParentName As String = ""
    Public ReadOnly Property currentParentId As Integer
        Get
            If String.IsNullOrEmpty(currentParentName) Then Return -1
            If currentClassName = "Variable" OrElse currentClassName = "Function" Then Return -1
            If currentClassName = "A" Then
                Return GetSecondChildIdByName(currentParentName, mScript.mainClass(mScript.mainClassHash("L")).ChildProperties)
            Else
                Return GetSecondChildIdByName(currentParentName, mScript.mainClass(mScript.mainClassHash(currentClassName)).ChildProperties)
            End If
        End Get
    End Property
    ''' <summary>Объект TreeView, с которым идет работа в данный момент (в зависимости от текущего класса currentClassName)</summary>
    Public Property currentTreeView As TreeView = Nothing
    ''' <summary>Меню для выбора иконок групп</summary>
    Public Property iconMenuGroups As ToolStrip
    ''' <summary>Меню для выбора иконок элементов</summary>
    Public Property iconMenuElements As ToolStrip
    ''' <summary>Класс, содержащий инфо о всех группах</summary>
    Public cGroups As New clsGroups
    ''' <summary>Класс управления списками для отображения в свойствах в списке в коде</summary>
    Public cListManager As New cListManagerClass
    ''' <summary>Класс управления вкладками</summary>
    Public cPanelManager As clsPanelManager
    Public Log As New cLog
    ''' <summary>Класс глобального поиска и замены</summary>
    Public GlobalSeeker As New cGlobalSeeker
    ''' <summary>Класс работы с сохраненными действиями</summary>
    Public actionsRouter As New cActionsRouter
    ''' <summary>Класс для хранения и восстановления удаленных элементов</summary>
    Public removedObjects As New cRemovedObjects
    ''' <summary>Класс для построения у управления картами</summary>
    Public mapManager As New cMapManager

#Region "Temporary"
    ' ''' <summary>
    ' ''' НЕ ИСПОЛЬЗУЕТСЯ? Функция поиска метки (при вызове оператора Jump). Ищет конец функции.
    ' ''' </summary>
    ' ''' <param name="s">ячейка кода для поиска</param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Private Function SearchMark(ByVal s As DictionaryEntry) As Boolean
    '    If s.Value = "End Function" OrElse s.Value.ToString.StartsWith("Function ") Then
    '        Return True
    '    Else
    '        Return False
    '    End If
    'End Function

#End Region

    Private Property classId As Integer

    ''' <summary>
    ''' Создает копию форматированного кода
    ''' </summary>
    ''' <param name="arr">Код, который копируется (из CodeTextBox)</param>
    ''' <returns>Копия кода</returns>
    Public Function CopyCodeDataArray(ByVal arr() As CodeTextBox.CodeDataType) As Array
        Dim tmpArr() As CodeTextBox.CodeDataType
        ReDim tmpArr(arr.GetUpperBound(0))
        Array.Copy(arr, tmpArr, arr.Length)
        Return tmpArr
    End Function

    ''' <summary>
    ''' Десериализует из xml-строки код в обычный текст
    ''' </summary>
    ''' <param name="strXml">Сериализованный код</param>
    Public Function ConvertCodePropertyToText(ByVal strXml As String) As String
        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(strXml)

        If cRes <> MatewScript.ContainsCodeEnum.CODE AndAlso cRes <> MatewScript.ContainsCodeEnum.LONG_TEXT Then
            'вероятно, содержимое свойства - код в виде обычного текста (или исполняемая строка)
            Return strXml
        End If
        'получаем десериализованный код в формате CodeData
        Dim CodeData() As CodeTextBox.CodeDataType = frmMainEditor.codeBox.codeBox.DeserializeCodeData(strXml)

        If CodeData.Length = 1 AndAlso IsNothing(CodeData(0).Code) AndAlso IsNothing(CodeData(0).Comments) AndAlso IsNothing(CodeData(0).StartingSpaces) Then Return ""

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
        Return sb.ToString
    End Function

    ''' <summary>
    ''' Заворачивает строку в кавычки (если надо). Числа и булевые значения не меняет.
    ''' </summary>
    ''' <param name="strToWrap">Строка, которую надо завернуть</param>
    ''' <returns>Если была передана строка без кавычек, то возвращается строка в кавычках. Иначе - исходные данные</returns>
    Public Function WrapString(ByVal strToWrap As String) As String
        If String.IsNullOrEmpty(strToWrap) Then Return "''"
        'If mScript.Param_GetType(strToWrap) <> MatewScript.ReturnFormatEnum.ORIGINAL Then Return strToWrap
        If Double.TryParse(strToWrap, Nothing) Then Return strToWrap
        strToWrap = strToWrap.Replace("'", "/'")
        Return "'" & strToWrap & "'"
    End Function

    Public Function GetParentClassId(ByVal childClassName As String) As Integer
        If childClassName = "A" Then Return mScript.mainClassHash("L")
        Return -1
    End Function

    ''' <summary>
    ''' Возвращает полный путь до указанного файла помощи
    ''' </summary>
    ''' <param name="aPath">Путь до указанного файла - либо полный, либо только от папки программы</param>
    Public Function GetHelpPath(ByVal aPath As String) As String
        If String.IsNullOrWhiteSpace(aPath) Then Return ""
        If aPath.IndexOf(":\") > -1 Then Return aPath
        Dim strPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "Help\" + aPath)
        If FileIO.FileSystem.FileExists(strPath) Then Return strPath
        Return FileIO.FileSystem.CombinePath(APP_HELP_PATH, aPath)
    End Function

    ''' <summary>
    ''' Возвращает содержит ли свойство массив или же нет (оно пустое или в p.returnArray находится имя класса для returnType = RETURN_ELEMENT)
    ''' </summary>
    ''' <param name="p">Свойство для проверки</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsPropertyContainsEnum(ByRef p As MatewScript.PropertiesInfoType) As Boolean
        If IsNothing(p.returnArray) OrElse p.returnArray.Count = 0 Then Return False
        If p.returnArray.Count > 1 Then Return True
        If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso _
            (mScript.mainClassHash.ContainsKey(p.returnArray(0)) OrElse p.returnArray(0) = "Variable") Then Return False
        Return True
    End Function

    ''' <summary>
    ''' Возвращает описание свойства для элемента выбранного уровня
    ''' </summary>
    ''' <param name="strDescription">Полный текс описания</param>
    ''' <param name="Level">Уровень элемента (от 0 до 2), для которого вывести описание</param>
    Public Function GetPropertyDescription(ByVal strDescription As String, Optional ByVal Level As Integer = -1) As String
        If Level < 0 Then
            If IsNothing(cPanelManager) Then Return ""
            Dim ch As clsPanelManager.clsChildPanel = cPanelManager.ActivePanel
            If IsNothing(ch) Then Return ""
            If ch.child2Name.Length = 0 Then
                Level = 0
            ElseIf ch.child3Name.Length = 0 Then
                Level = 1
            Else
                Level = 2
            End If
        End If
        If String.IsNullOrWhiteSpace(strDescription) Then Return strDescription
        '[Level1]
        'Dim strLevel As String = "Level" + (Level + 1).ToString
        Dim pos(3) As Integer
        pos(0) = strDescription.IndexOf("[Level1]")
        pos(1) = strDescription.IndexOf("[Level2]")
        pos(2) = strDescription.IndexOf("[Level3]")

        If pos(0) = -1 AndAlso pos(1) = -1 AndAlso pos(2) = -1 Then Return strDescription
        If pos(Level) = -1 Then
            Return ""
        End If

        Dim sStart As Integer = pos(Level) + "[Level1]".Length
        If Level = 2 OrElse pos(Level + 1) = -1 Then
            Return strDescription.Substring(sStart).Trim
        End If
        Dim sEnd As Integer = pos(Level + 1)
        'Dim desc As String = ""
        Return strDescription.Substring(sStart, sEnd - sStart).Trim
    End Function

    ''' <summary>
    ''' Возвращаем описание функции. Если в описании присутствует ключевое слово Return, то возвращается текст до этого слова
    ''' </summary>
    ''' <param name="strDescription">Строка описания</param>
    Public Function GetFunctionDescription(ByVal strDescription As String) As String
        If String.IsNullOrWhiteSpace(strDescription) Then Return strDescription
        Dim pos As Integer = strDescription.IndexOf("[Return]")

        If pos = -1 Then Return strDescription
        Return strDescription.Substring(0, pos).Trim
    End Function

    ''' <summary>
    ''' Возвращает текст в строке описания функции, следующий за ключевым словом [Return], вычленив из него тип возвращаемого значения (если указан)
    ''' </summary>
    ''' <param name="strDescription">Строка описания функции</param>
    ''' <param name="strReturn">ссылка на строку возвращаемого значения</param>
    ''' <param name="strReturnType">тип возвращаемого значения</param>
    ''' <remarks></remarks>
    Public Sub GetFunctionDescriptionReturnData(ByVal strDescription As String, ByRef strReturn As String, ByRef strReturnType As String)
        strReturn = ""
        strReturnType = ""

        If String.IsNullOrWhiteSpace(strDescription) Then Return
        Dim pos As Integer = strDescription.IndexOf("[Return]")
        If pos = -1 Then Return
        strReturn = strDescription.Substring(pos + "[Return]".Length).Trim
        If strReturn.Length = 0 OrElse strReturn.StartsWith("[") = False Then Return

        pos = strReturn.IndexOf("]"c)
        If pos = -1 Then Return
        strReturnType = strReturn.Substring(1, pos - 1)
        If pos = strReturn.Length - 1 Then
            strReturn = ""
        Else
            strReturn = strReturn.Substring(pos + 1)
        End If
    End Sub

    Public Sub AAA()
        If IsNothing(cPanelManager.lstPanels) Then Return
        For i = 0 To cPanelManager.lstPanels.Count - 1
            If IsNothing(cPanelManager.lstPanels(i)) Then
                MsgBox("Nothing " + i.ToString)
                Exit For
            End If
        Next
    End Sub
End Module




