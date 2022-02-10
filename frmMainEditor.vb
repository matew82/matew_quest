Imports System.Collections.ObjectModel

<System.Runtime.InteropServices.ComVisible(True)>
Public Class frmMainEditor
    ''' <summary> Класс для получения инфо про узел дерева для функции GetChildsByTreeNode</summary>
    Public Class NodeInfo
        ''' <summary>Id класса, к которому принадлежит узел</summary>
        Public Property classId As Integer = -1
        ''' <summary>Принадлежит ли узел дереву узлов 3 уровня (например, содержащих пункты меню) или нет (например, действия для локации)</summary>
        Public Property ThirdLevelNode As Boolean = False
        ''' <summary>Имя соответствующего элемента в структуре mScript.mainClass(classId).ChildProperties </summary>
        Public Property nodeChild2Name As String = ""
        ''' <summary>Имя соответствующего элемента в структуре mScript.mainClass(classId).ChildProperties(nodeChild2Id)("Name").ThirdLevelProperties </summary>
        Public Property nodeChild3Name As String = ""

        Public Function GetChild3Id(Optional ByVal child2Id As Integer = -1) As Integer
            If classId < 0 OrElse String.IsNullOrEmpty(nodeChild3Name) Then Return -1
            If child2Id < 0 Then child2Id = GetChild2Id()
            If child2Id < 0 Then Return -1
            Dim res As Integer = GetThirdChildIdByName(nodeChild3Name, child2Id, mScript.mainClass(classId).ChildProperties)
            If res < 0 Then
                Return -1
            Else
                Return res
            End If
        End Function

        Public Function GetChild2Id() As Integer
            If String.IsNullOrEmpty(nodeChild2Name) Then Return -1
            If classId = -2 Then
                Try
                    Return mScript.csPublicVariables.lstVariables.IndexOfKey(nodeChild2Name)
                Catch ex As Exception
                    Return -1
                End Try
            ElseIf classId = -3 Then
                Try
                    Return mScript.functionsHash.IndexOfKey(nodeChild2Name)
                Catch ex As Exception
                    Return -1
                End Try
            Else
                Dim res As Integer = GetSecondChildIdByName(nodeChild2Name, mScript.mainClass(classId).ChildProperties)
                If res < 0 Then
                    Return -1
                Else
                    Return res
                End If
            End If
        End Function
    End Class

    Dim prevNodeMainTree As TreeNode = Nothing 'предыдущий узел основного дерева (для возврата ему цвета по-умолчанию)
    Dim prevNodeSubTree As TreeNode = Nothing 'предыдущий узел дерева подклассов (для возврата ему цвета по-умолчанию)
    'перемещаемый узел и узел, куда перемещается (для события DragDrop)
    Dim nodeUnderMouseToDrag, nodeUnderMouseToDrop As TreeNode
    Dim dragBoxFromMouseDown As Rectangle 'квадат, при выходе за пределы которого начинается DragDrop

    ''' <summary>Дерево для отображения в свойствах типа RETURN_ELEMENT </summary>
    Public treeProperties As New TreeView With {.AllowDrop = True, .Dock = DockStyle.None, .FullRowSelect = True, .LabelEdit = False, .Scrollable = True, .ShowLines = True, _
                              .ShowPlusMinus = True, .ShowRootLines = True, .Visible = False, .HotTracking = True}
    ''' <summary>Массив с деревьями всех классов</summary>
    Public dictTrees As New Dictionary(Of Integer, TreeView)
    ''' <summary>Дерево для переменных</summary>
    Public treeVariables As TreeView
    ''' <summary>Дерево для функций</summary>
    Public treeFunctions As TreeView

    Public WithEvents WBhelp As New WebBrowser With {.Dock = DockStyle.Fill, .AllowWebBrowserDrop = False, .IsWebBrowserContextMenuEnabled = False, _
                                                     .ScrollBarsEnabled = True, .WebBrowserShortcutsEnabled = True, .Visible = False}
    Private WithEvents hDocument As HtmlDocument

    ''' <summary>Главный кодбокс</summary>
    Public WithEvents codeBox As New CodeTextBox
    ''' <summary>Панель, содержащая кодбокс и меню управления им</summary>
    Public codeBoxPanel As New Panel With {.Dock = DockStyle.Fill, .Visible = False}
    ''' <summary>Чекбокс отображать ли скрытые элементы</summary>
    Public chkShowHidden As CheckBox
    ''' <summary>Элементы управления над главными treeView</summary>
    Private btnAddGroup As Button, btnAddElement As Button, btnDelete As Button, btnDuplicate As Button
    Public btnShowSettings As Button, btnMakeHidden As Button, btnCopyTo As Button
    ''' <summary>Контейнер, в Panel1 которого находится панель управлению деревом, а в Panel2 - само дерево</summary>
    Public splitTreeContainer As SplitContainer
    ''' <summary>Таймер для обновления списков файлов - наиболее оптимальный подход чтобы не обновлять список 100 раз, если добавлено 100 файлов</summary>
    Private WithEvents timUpdateFiles As New Timer With {.Interval = 500, .Enabled = False}
    ''' <summary>Картинка - карта для выбора цвета</summary>
    Public imgSelectColorMap As Bitmap
    ''' <summary>Первичное переименование элемента (без проверки скриптов на наличие такого имени)</summary>
    Private primaryNameSelection As Boolean = False
    Public Enum trackingcodeEnum As Byte
        NOT_TRACKING_EVENT = 0
        EVENT_BEFORE = 1
        EVENT_AFTER = 2
    End Enum
    ''' <summary>Открыт ли сейчас кодбокс с событием редактирования свойства TrackingEvent</summary>
    Public Property trakingEventState As trackingcodeEnum = trackingcodeEnum.NOT_TRACKING_EVENT

#Region "События формы"

    Private Sub frmMainEditor_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub


    Private Sub frmMainEditor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        VBMath.Randomize()

        WBhelp.ObjectForScripting = Me
        pnlVariables.Dock = DockStyle.Fill
        questEnvironment.codeBoxShadowed.codeBox.csUndo.AllowUndo = False
        SplitInner.Panel2.Controls.Add(codeBoxPanel)
        SplitInner.Panel2.Controls.Add(WBhelp)
        'создаем класс cPanelManager, управляющим панелями (окнами) управляния классами
        cPanelManager = New clsPanelManager(SplitInner, WBhelp, ToolStripPanels, btnShowSettings)
        'очищаем все старые данные, загружаем структуру
        questEnvironment.ClearAllData()
        'создаем копию структуры для поддержки восстановления исходных данных
        mScript.CreateMainClassCopy()
        PrepareCodeBoxPanel()
        codeBox.codeBox.csUndo.btnUndo = tsbUndo
        codeBox.codeBox.csUndo.btnRedo = tsbRedo
        'подготавливаем дерево, отображаемое в файлах помощи
        PrepareTreeProperties()
        'загрузка карtинки-карты цветов для возможности выбора цвета из палитры данного рисунка
        Dim mapPath As String = Application.StartupPath + "\src\img\colorMap.bmp"
        If My.Computer.FileSystem.FileExists(mapPath) Then
            imgSelectColorMap = Image.FromFile(mapPath)
        End If

        questEnvironment.LoadQuest(FileIO.FileSystem.CombinePath(Application.StartupPath, "Quests\myQuest"))
        frmMainEditor_Resize(Me, New EventArgs)

        ''''''''''''''''''''''''''''''''''''''''''!!!!!!!!!!!!!!!!!!!!!!!!
        'WBhelp.ObjectForScripting = Me
        'pnlVariables.Dock = DockStyle.Fill
        'questEnvironment.codeBoxShadowed.codeBox.csUndo.AllowUndo = False

        'mScript.LoadClasses()
        'mScript.CreateMainClassCopy()
        'mScript.FillFuncAndPropHash()
        'cGroups.PrepareGroups()
        'questEnvironment.QuestPath = Application.StartupPath + "\Quests\MyQuest"
        'questEnvironment.QuestShortName = "MyQuest"
        'currentClassName = "L"

        'LoadTrees()
        'Dim tree As TreeView = dictTrees(mScript.mainClassHash(currentClassName))
        'currentTreeView = tree

        'SplitInner.Panel2.Controls.Add(codeBoxPanel)
        'cPanelManager = New clsPanelManager(SplitInner, WBhelp, ToolStripPanels, btnShowSettings)
        'tree.BeginUpdate()
        'Call tsmiAddGroup_Click(tsmiAddGroup, New System.EventArgs)

        'Call tsmiAddGroup_Click(tsmiAddGroup, New System.EventArgs)

        'Call tsmiAddGroup_Click(tsmiAddGroup, New System.EventArgs)

        'Call tsmiAddGroup_Click(tsmiAddGroup, New System.EventArgs)

        'Call tsmiAddGroup_Click(tsmiAddGroup, New System.EventArgs)

        'tree.Nodes.Clear()
        'mScript.ExecuteString({"L.Create 'Локация 1'", "L.Create 'Локация 2'", "L.Create 'Локация 3'", "L.Create 'Локация 4'", "L.Create 'Локация 5'", "L.Create 'Локация 6'", "L.Create 'Локация 7'", _
        '                       "L.Create 'Локация 8'", "L.Create 'Локация 9'", "L.Create 'Локация 10'", "L[1].Group = 'Группа 1'", "L[2].Group = 'Группа 1'", "L[3].Group = 'Группа 2'", _
        '                       "L[4].Group = 'Группа 3'", "L[5].Group = 'Группа 3'", "L[6].Group = 'Группа 3'", "L[7].Group = 'Группа 4'", "L[8].Group = 'Группа 5'", "L[9].Group = 'Группа 5'"}, Nothing)
        'mScript.ExecuteString({"A.Add 'Действие 1'", "A.Add 'Действие 2'", "A.Add 'Действие 3'"}, Nothing)
        'tree.EndUpdate()

        'mScript.csPublicVariables.SetVariableInternal("ALL", "-1", 0)

        'splitTreeContainer.Show()
        'FillTree(currentClassName, currentTreeView, True, Nothing, "")
        'FillTreeVariables(chkShowHidden.Checked, Nothing)
        'FillTreeFunctions(chkShowHidden.Checked, Nothing)
        'SplitInner.Panel2.Controls.Add(WBhelp)
        'tree.Show()
        'PrepareCodeBoxPanel()

        'cListManager.UpdateElementsLists()
        'cListManager.UpdateFileLists()

        'Dim di As New System.IO.DirectoryInfo(questEnvironment.QuestPath)
        'fsWatcher.Path = di.Root.FullName
        'PrepareTreeProperties()

        'Dim mapPath As String = Application.StartupPath + "\src\img\colorMap.bmp"
        'If My.Computer.FileSystem.FileExists(mapPath) Then
        '    imgSelectColorMap = Image.FromFile(mapPath)
        'End If

        'codeBox.codeBox.csUndo.btnUndo = tsbUndo
        'codeBox.codeBox.csUndo.btnRedo = tsbRedo
    End Sub
#End Region

#Region "TreeProperties"
    ''' <summary>
    ''' Подготавливает дерево, которое размещается в браузере для удобства выбора значения свойства
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub PrepareTreeProperties()
        WBhelp.Controls.Add(treeProperties)
        treeProperties.ImageList = imgLstGroupIcons
        AddHandler treeProperties.AfterSelect, AddressOf del_treeProperties_AfterSelect

        AddHandler treeProperties.BeforeSelect, Sub(sender2 As Object, e2 As TreeViewCancelEventArgs)
                                                    Dim n As TreeNode = sender2.SelectedNode
                                                    If IsNothing(n) = False Then
                                                        If n.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then
                                                            n.ForeColor = DEFAULT_COLORS.NodeForeColor
                                                            n.BackColor = DEFAULT_COLORS.ControlTransparentBackground
                                                        End If
                                                    End If
                                                    n = e2.Node
                                                    If IsNothing(n) = False Then
                                                        If n.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then
                                                            n.ForeColor = DEFAULT_COLORS.NodeSelForeColor
                                                            n.BackColor = DEFAULT_COLORS.NodeSelBackColor
                                                        End If
                                                    End If
                                                End Sub
        AddHandler treeProperties.DragEnter, AddressOf del_treeProperties_DragEnter
        AddHandler treeProperties.DragOver, AddressOf del_treeProperties_DragOver
        AddHandler treeProperties.DragDrop, AddressOf del_treeProperties_DragDrop
        AddHandler treeProperties.QueryContinueDrag, AddressOf del_treeProperties_QueryContinueDrag
        AddHandler treeProperties.MouseUp, Sub(sender2 As Object, e2 As MouseEventArgs)
                                               draggedFiles = Nothing
                                           End Sub
        AddHandler treeProperties.DoubleClick, AddressOf del_treeProperties_DoubleClick
        AddHandler treeProperties.MouseDown, AddressOf del_treeProperties_MouseDown
    End Sub

    Dim draggedFiles() As String = Nothing
    Private Sub del_treeProperties_DragEnter(sender As Object, e As DragEventArgs)
        'начало операции перетаскивания файлов
        If e.Data.GetDataPresent("FileDrop") = False Then Return
        Dim files() As String = e.Data.GetData("FileDrop")
        If files.Count = 0 Then Return
        'перетаскивают именно файлы
        If files(0).ToLower.StartsWith(questEnvironment.QuestPath.ToLower) Then Return 'из директории квеста не перетаскиваются

        'получаем какие именной файлы можно перетаскивать (в зависимости от returnType текущего элемента)
        Dim classId As Integer
        Dim c As Object = Nothing
        Dim propName As String
        Dim forSubTree As Boolean = False

        If IsNothing(WBhelp.Tag) = False AndAlso WBhelp.Tag.GetType.Name = "TreeNode" Then
            'Свойство по умолчанию для элементов 3 уровня
            forSubTree = True
            Dim ni As NodeInfo = GetNodeInfo(WBhelp.Tag)
            classId = ni.classId
            propName = mScript.mainClass(classId).DefaultProperty
        Else
            'Обычное свойство текущей панели
            If IsNothing(cPanelManager.ActivePanel) Then Return
            classId = cPanelManager.ActivePanel.classId
            c = cPanelManager.ActivePanel.ActiveControl
            If IsNothing(c) Then Return
            propName = c.Name
        End If

        Dim prop As MatewScript.PropertiesInfoType
        If forSubTree = False AndAlso c.IsFunctionButton Then
            prop = mScript.mainClass(classId).Functions(propName)
        Else
            prop = mScript.mainClass(classId).Properties(propName)
        End If

        'получаем список допустимых расширений перетаскиваемых файлов
        Dim arrExtensions() As String
        Select Case prop.returnType
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE
                arrExtensions = {".png", ".jpg", ".gif", ".jpeg", ".bmp", ".wmf"}
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO
                arrExtensions = {".mp3", ".mid", ".wav"}
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_TEXT
                arrExtensions = {".txt"}
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_JS
                arrExtensions = {".js"}
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS
                arrExtensions = {".css"}
            Case Else
                Return
        End Select
        Dim lstExtensions As List(Of String) = arrExtensions.ToList

        ReDim draggedFiles(files.Count - 1)
        Dim fUBound As Integer = -1
        For i As Integer = 0 To files.Count - 1
            If lstExtensions.IndexOf(System.IO.Path.GetExtension(files(i))) > -1 Then
                fUBound += 1
                draggedFiles(fUBound) = files(i)
            End If
        Next
        If fUBound = -1 Then Return
        ReDim Preserve draggedFiles(fUBound)
        'в draggedFiles - только файлы с правильным расширением
        treeProperties.DoDragDrop(e.Data, DragDropEffects.Copy)
    End Sub

    Private Sub del_treeProperties_DragOver(sender As Object, e As DragEventArgs)
        If IsNothing(draggedFiles) Then Return
        If ((e.AllowedEffect And DragDropEffects.Move) = DragDropEffects.Move) Then
            ' By default, the drop action should be move, if allowed.
            e.Effect = DragDropEffects.Move
        Else
            e.Effect = DragDropEffects.None
        End If

        'получаем узел, который под мышью на данный момент
        Dim prevNodeDrop As TreeNode = nodeUnderMouseToDrop
        nodeUnderMouseToDrop = treeProperties.GetNodeAt(treeProperties.PointToClient(New Point(e.X, e.Y)))
        If IsNothing(nodeUnderMouseToDrop) = False AndAlso IsNothing(nodeUnderMouseToDrop.Tag) = False AndAlso nodeUnderMouseToDrop.Tag.ToString = "ITEM" Then _
            nodeUnderMouseToDrop = nodeUnderMouseToDrop.Parent
        If Object.Equals(treeProperties.SelectedNode, nodeUnderMouseToDrop) Then Return
        treeProperties.SelectedNode = nodeUnderMouseToDrop
    End Sub

    Private Sub del_treeProperties_DragDrop(sender As Object, e As DragEventArgs)
        If IsNothing(draggedFiles) Then Return
        If e.Data.GetDataPresent("FileDrop") = False Then Return
        If IsNothing(nodeUnderMouseToDrop) Then Return
        Dim selNode As TreeNode = treeProperties.SelectedNode
        If Object.Equals(selNode, nodeUnderMouseToDrop) = False Then Return
        If e.Effect <> DragDropEffects.Move Then Return

        Dim classId As Integer
        Dim c As Object = Nothing
        Dim propName As String
        Dim forSubTree As Boolean = False

        If IsNothing(WBhelp.Tag) = False AndAlso WBhelp.Tag.GetType.Name = "TreeNode" Then
            'Свойство по умолчанию для элементов 3 уровня
            forSubTree = True
            Dim ni As NodeInfo = GetNodeInfo(WBhelp.Tag)
            classId = ni.classId
            propName = mScript.mainClass(classId).DefaultProperty
        Else
            'Обычное свойство текущей панели
            If IsNothing(cPanelManager.ActivePanel) Then Return
            classId = cPanelManager.ActivePanel.classId
            c = cPanelManager.ActivePanel.ActiveControl
            If IsNothing(c) Then Return
            propName = c.Name
        End If

        Dim prop As MatewScript.PropertiesInfoType
        If forSubTree = False AndAlso c.IsFunctionButton Then
            prop = mScript.mainClass(classId).Functions(propName)
        Else
            prop = mScript.mainClass(classId).Properties(propName)
        End If

        For i As Integer = 0 To draggedFiles.Count - 1
            Dim fName As String = System.IO.Path.GetFileName(draggedFiles(i))
            Dim newPath As String = System.IO.Path.Combine(questEnvironment.QuestPath, selNode.Name, fName)
            If My.Computer.FileSystem.FileExists(newPath) Then
                Dim ret As DialogResult = MessageBox.Show("Файл " + newPath + " уже существует. Заменить его новым?", "MatewQuest", MessageBoxButtons.YesNoCancel)
                If ret = Windows.Forms.DialogResult.Cancel Then
                    Erase draggedFiles
                    Return
                ElseIf ret = Windows.Forms.DialogResult.No Then
                    Continue For
                End If
            End If
            My.Computer.FileSystem.CopyFile(draggedFiles(i), newPath, True)
            Select Case prop.returnType
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE
                    AddHTMLimgToContainer(Nothing, selNode.Name + "\" + fName)
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO
                    Dim n As TreeNode = nodeUnderMouseToDrop.Nodes.Add(selNode.Name + "\" + fName, fName)
                    n.Tag = "ITEM"
            End Select

        Next
    End Sub

    Public Sub del_treeProperties_QueryContinueDrag(sender As Object, e As QueryContinueDragEventArgs)
        'Организация DragDop - перемещение функций и свойств между классами
        If e.EscapePressed Then
            'При нажатии Esc прекращаем перетаскивание
            e.Action = DragAction.Cancel
            nodeUnderMouseToDrag = Nothing
            Erase draggedFiles
        End If
    End Sub

    Private Sub del_treeProperties_DoubleClick(sender As Object, e As EventArgs)
        Dim classId As Integer
        Dim c As Object = Nothing
        Dim propName As String
        Dim forSubTree As Boolean = False

        If IsNothing(WBhelp.Tag) = False AndAlso WBhelp.Tag.GetType.Name = "TreeNode" Then
            'Свойство по умолчанию для элементов 3 уровня
            forSubTree = True
            Dim ni As NodeInfo = GetNodeInfo(WBhelp.Tag)
            classId = ni.classId
            propName = mScript.mainClass(classId).DefaultProperty
        Else
            'Обычное свойство текущей панели
            If IsNothing(cPanelManager.ActivePanel) Then Return
            classId = cPanelManager.ActivePanel.classId
            c = cPanelManager.ActivePanel.ActiveControl
            If IsNothing(c) Then Return
            propName = c.Name
        End If

        Dim prop As MatewScript.PropertiesInfoType
        If forSubTree = False AndAlso c.IsFunctionButton Then
            prop = mScript.mainClass(classId).Functions(propName)
        Else
            prop = mScript.mainClass(classId).Properties(propName)
        End If

        Select Case prop.returnType
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO, MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS, MatewScript.ReturnFunctionEnum.RETURN_PATH_JS, _
                MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE, MatewScript.ReturnFunctionEnum.RETURN_PATH_TEXT
                Dim path As String = questEnvironment.QuestPath + "\" + treeProperties.SelectedNode.Name + "\"
                Process.Start(path)
        End Select
    End Sub

    Private Sub del_treeProperties_AfterSelect(sender As Object, e As TreeViewEventArgs)
        Dim classId As Integer
        Dim child2Id As Integer 'только для forSubTree
        Dim child3Id As Integer 'только для forSubTree
        Dim c As Object = Nothing
        Dim propName As String
        Dim forSubTree As Boolean = False

        If IsNothing(WBhelp.Tag) = False AndAlso WBhelp.Tag.GetType.Name = "TreeNode" Then
            'Свойство по умолчанию для элементов 3 уровня
            forSubTree = True
            Dim ni As NodeInfo = GetNodeInfo(WBhelp.Tag)
            classId = ni.classId
            child2Id = ni.GetChild2Id
            child3Id = ni.GetChild3Id(child2Id)
            propName = mScript.mainClass(classId).DefaultProperty
        Else
            'Обычное свойство текущей панели
            If IsNothing(cPanelManager.ActivePanel) Then Return
            classId = cPanelManager.ActivePanel.classId
            c = cPanelManager.ActivePanel.ActiveControl
            If IsNothing(c) Then Return
            propName = c.Name
            child2Id = cPanelManager.ActivePanel.GetChild2Id
            child3Id = cPanelManager.ActivePanel.GetChild3Id(child2Id)
        End If

        Dim prop As MatewScript.PropertiesInfoType
        If forSubTree = False AndAlso c.IsFunctionButton Then
            prop = mScript.mainClass(classId).Functions(propName)
        Else
            prop = mScript.mainClass(classId).Properties(propName)
        End If

        Select Case prop.returnType
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE
                'в дереве список папок с картинками
                If IsNothing(hDocument) Then Return
                Dim hContainer As HtmlElement = hDocument.GetElementById("dataContainer")
                If IsNothing(hContainer) Then Return
                hContainer.Style = "display:block"
                hContainer.InnerHtml = ""
                If IsNothing(e.Node) Then Return
                Dim shortPath As String = e.Node.Name
                If shortPath.Length > 0 Then shortPath += "\"
                Dim arrPictures() As String = cListManager.GetFilesList(cListManagerClass.fileTypesEnum.PICTURES, False)
                Dim fCount As Integer = 0

                For i As Integer = 0 To arrPictures.Count - 1
                    Dim curStr As String = arrPictures(i)
                    If curStr.StartsWith(shortPath) Then
                        If curStr.Substring(shortPath.Length).IndexOf("\"c) > -1 Then Continue For 'файл из подкаталога данной папки - не включаем
                    Else
                        Continue For
                    End If
                    'Файл из этой папки - включаем
                    AddHTMLimgToContainer(hContainer, arrPictures(i))
                    fCount += 1
                Next
                If fCount = 0 Then hContainer.Style = "display:none"
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO
                If e.Node.Tag <> "ITEM" Then Return
                If forSubTree Then
                    SetPropertyValue(classId, propName, WrapString(e.Node.Name), child2Id, child3Id)
                Else
                    c.Text = e.Node.Name
                End If
                mPlayer.Stop()
                mPlayer.Open(questEnvironment.QuestPath + "\" + e.Node.Name)
                mPlayer.Play()
            Case MatewScript.ReturnFunctionEnum.RETURN_ELEMENT, MatewScript.ReturnFunctionEnum.RETURN_FUNCTION
                If e.Node.Tag <> "ITEM" Then Return
                If forSubTree = False Then
                    c.Text = e.Node.Text
                End If
                SetPropertyValue(classId, propName, WrapString(e.Node.Text), child2Id, child3Id)
            Case Else
                If e.Node.Tag <> "ITEM" Then Return
                Dim fPath As String = e.Node.FullPath
                If fPath.StartsWith(questEnvironment.QuestShortName + "\") Then fPath = fPath.Substring(questEnvironment.QuestShortName.Length + 1)
                If forSubTree = False Then
                    c.Text = fPath
                End If
                SetPropertyValue(classId, propName, WrapString(fPath), child2Id, child3Id)
        End Select
    End Sub

    Private Sub del_treeProperties_MouseDown(sender As Object, e As MouseEventArgs)
        Dim classId As Integer
        Dim child2Id As Integer 'только для forSubTree
        Dim child3Id As Integer 'только для forSubTree
        Dim c As Object = Nothing
        Dim propName As String
        Dim forSubTree As Boolean = False

        If IsNothing(WBhelp.Tag) = False AndAlso WBhelp.Tag.GetType.Name = "TreeNode" Then
            'Свойство по умолчанию для элементов 3 уровня
            forSubTree = True
            Dim ni As NodeInfo = GetNodeInfo(WBhelp.Tag)
            classId = ni.classId
            child2Id = ni.GetChild2Id
            child3Id = ni.GetChild3Id(child2Id)
            propName = mScript.mainClass(classId).DefaultProperty
        Else
            'Обычное свойство текущей панели
            If IsNothing(cPanelManager.ActivePanel) Then Return
            classId = cPanelManager.ActivePanel.classId
            c = cPanelManager.ActivePanel.ActiveControl
            If IsNothing(c) Then Return
            propName = c.Name
        End If

        Dim prop As MatewScript.PropertiesInfoType
        If forSubTree = False AndAlso c.IsFunctionButton Then
            prop = mScript.mainClass(classId).Functions(propName)
        Else
            prop = mScript.mainClass(classId).Properties(propName)
        End If

        Select Case prop.returnType
            Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO
                If mPlayer.PlayState = MediaPlayer.MPPlayStateConstants.mpPaused Then
                    mPlayer.Play()
                Else
                    mPlayer.Pause()
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Добавляет html-элемент изображение (для свойств с returnType = PATH_PICTURE)
    ''' </summary>
    ''' <param name="hContainer">контейнер, в котором размещаются картинки</param>
    ''' <param name="imgPath">путь к картинке относительно папки квеста</param>
    ''' <remarks></remarks>
    Private Sub AddHTMLimgToContainer(ByRef hContainer As HtmlElement, ByVal imgPath As String)
        If IsNothing(hContainer) Then
            hContainer = hDocument.GetElementById("dataContainer")
            If IsNothing(hContainer) Then Return
        End If
        Dim hImg As HtmlElement = hDocument.CreateElement("IMG")
        hImg.SetAttribute("src", questEnvironment.QuestPath + "\" + imgPath)
        hImg.SetAttribute("className", "thumbnail selectable")
        Dim aPos As Integer = imgPath.LastIndexOf("\")
        If aPos = -1 OrElse aPos = imgPath.Length - 1 Then
            hImg.SetAttribute("Title", imgPath)
        Else
            hImg.SetAttribute("Title", imgPath.Substring(aPos + 1))
        End If
        hImg.SetAttribute("rVal", imgPath.Replace("\", "/"))
        hContainer.AppendChild(hImg)

        AddHandler hImg.Click, AddressOf del_html_img_click
    End Sub

    Private Sub del_html_img_click(sender As Object, e As HtmlElementEventArgs)
        If IsNothing(hDocument) Then Return
        Dim oldImg As HtmlElement = sender
        Dim hSample As HtmlElement = hDocument.GetElementById("sampleContainer")
        If IsNothing(hSample) Then Return
        hSample.InnerHtml = ""
        Dim newImg As HtmlElement = hDocument.CreateElement("IMG")
        newImg.SetAttribute("src", oldImg.GetAttribute("src"))
        newImg.SetAttribute("className", "thumbnail")
        hSample.AppendChild(newImg)
    End Sub

    Private Sub frmMainEditor_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        SplitOuter.Location = New Point(0, ToolStripPanels.Bottom)
        SplitOuter.Size = New Size(Me.ClientRectangle.Width, Me.ClientRectangle.Height - SplitOuter.Top)
    End Sub
#End Region

#Region "MainTreeToolStrip"
    ''' <summary>
    ''' Возвращает Id родительского элемента - Id элемента 2 уровня для элементов 3 уровня и Id локации для действий.
    ''' </summary>
    ''' <param name="className">Имя класса, с которым сейчас идет работа ///ОТКЛЮЧЕНО - не используется///</param>
    ''' <param name="tree">Дерево, с которым сейчас идет работа</param>
    Public Function GetParentId(ByVal className As String, ByRef tree As TreeView) As Integer
        If Object.Equals(currentTreeView, tree) Then
            'главное дерево. Возвращаем currentParentId
            Return currentParentId
        Else
            'второстепенное дерево. Возвращаем child2Id
            If IsNothing(cPanelManager.ActivePanel) Then Return -1
            Return cPanelManager.ActivePanel.GetChild2Id
        End If
    End Function

    '''<summary>Возвращает Имя текущего родительского элемента - Id элемента 2 уровня для элементов 3 уровня и имя локации для действий.</summary>
    ''' <param name="tree">Дерево, с которым сейчас идет работа</param>
    Public Function GetParentName(ByRef tree As TreeView) As String
        If Object.Equals(currentTreeView, tree) Then
            'главное дерево. Возвращаем currentParentName
            Return currentParentName
        Else
            'второстепенное дерево. Возвращаем child2Id
            If IsNothing(cPanelManager.ActivePanel) Then Return ""
            With cPanelManager.ActivePanel
                If .classId < 0 Then Return ""
                If .child2Name.Length = 0 Then Return ""
                Return .child2Name
            End With
        End If
    End Function


    ''' <summary>
    ''' Создание новой группы для указанного класса
    ''' </summary>
    Private Sub tsmiAddGroup_Click(sender As Object, e As EventArgs) Handles tsmiAddGroup.Click
        AddGroup(currentClassName, currentTreeView)
    End Sub

    ''' <summary>
    ''' Добавляем новую группу элементам 2 или 3 уровня
    ''' </summary>
    ''' <param name="className">Имя класса, кторому элемнт добавляется</param>
    ''' <param name="tree">Дерево, содержащее узлы данного класса</param>
    Public Sub AddGroup(ByVal className As String, ByRef tree As TreeView)
        Dim defGroupName As String = "Группа "
        Dim thirdLevel As Boolean = IsThirdLevelTree(tree)
        'Создаем новое имя группы
        Dim parentName = GetParentName(tree)
        Dim i As Integer = 1
        Do While cGroups.IsGroupExist(className, defGroupName + CStr(i), thirdLevel, parentName)
            i += 1
        Loop

        Dim groupName As String = defGroupName + CStr(i)
        If cGroups.AddGroup(className, groupName, thirdLevel, False, parentName) = -1 Then
            MessageBox.Show("Не удалось создать группу. Вероятно, группа с таким именем уже существует.")
            Return
        End If
        'Добавляем группы в treeView
        Dim curNode As TreeNode = tree.Nodes.Add(groupName)
        curNode.ImageKey = "groupDefault.png"
        curNode.SelectedImageKey = curNode.ImageKey
        curNode.Tag = "GROUP"
        tree.SelectedNode = curNode
    End Sub

    Private Sub tsmiAddElement_Click(sender As Object, e As EventArgs) Handles tsmiAddElement.Click
        AddElement(currentClassName, currentTreeView)
    End Sub

    ''' <summary>
    ''' Добавляем новый элемент 2 или 3 уровня (Локацию, действие ...)
    ''' </summary>
    ''' <param name="className">Имя класса, кторому элемнт добавляется</param>
    ''' <param name="tree">Дерево, содержащее узлы данного класса</param>
    Public Function AddElement(ByVal className As String, ByRef tree As TreeView, Optional ByRef retreiveObject As Object = Nothing, Optional newName As String = "", Optional showCells As Boolean = True) As String
        'Получаем новое имя
        AAA()
        Dim thirdLevel As Boolean = IsThirdLevelTree(tree)
        If showCells AndAlso currentClassName = "Map" AndAlso thirdLevel Then
            mapManager.BuildMapForCellsEdit()
            Return ""
        End If
        Dim parentId As Integer = GetParentId(className, tree)
        Dim elementName As String = ""
        'Непосредственное создание элемента
        If className = "Variable" Then
            If String.IsNullOrEmpty(newName) Then
                elementName = GetNewDefName(className, -1)
                mScript.csPublicVariables.SetVariableInternal(elementName, "", 0)
            Else
                mScript.csPublicVariables.lstVariables.Add(newName, retreiveObject)
            End If
        ElseIf className = "Function" Then
            'Создаем список панелей, класс которых Function. Им всем потом надо изменить child2Id
            If String.IsNullOrEmpty(newName) Then
                elementName = GetNewDefName(className, -1)
                mScript.functionsHash.Add(elementName, New MatewScript.FunctionInfoType)
            Else
                elementName = newName
                mScript.functionsHash.Add(elementName, retreiveObject)
            End If
        Else
            Dim create As String = "Create"
            If className = "A" OrElse className = "T" Then create = "Add"
            If String.IsNullOrEmpty(newName) = False Then
                'Восстановление удаленного
                elementName = mScript.PrepareStringToPrint(newName, Nothing, False)
                Dim classId As Integer = mScript.mainClassHash(currentClassName)
                If thirdLevel Then
                    'Восстановление элемента 3 уровня
                    If mScript.ExecuteString({className + "." + create + "(" + parentId.ToString + ", " + newName + ")"}, Nothing) = "#Error" Then Return ""
                    Dim chProps As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classId).ChildProperties(parentId)
                    Dim newId As Integer = mScript.mainClass(classId).ChildProperties(parentId)("Name").ThirdLevelProperties.Count - 1
                    Dim src As SortedList(Of String, MatewScript.PropertiesInfoType) = retreiveObject
                    mScript.eventRouter.EraseElementEvents(classId, parentId, newId)
                    For i As Integer = 0 To chProps.Count - 1
                        Dim ch As MatewScript.ChildPropertiesInfoType = chProps.ElementAt(i).Value
                        Dim chName As String = chProps.ElementAt(i).Key
                        If chName = "Name" Then
                            ch.ThirdLevelProperties(newId) = newName
                        Else
                            ch.ThirdLevelProperties(newId) = src(chName).Value
                        End If
                        ch.ThirdLevelEventId(newId) = src(chName).eventId
                    Next i
                Else
                    'Восстановление элемента 2 уровня
                    If mScript.ExecuteString({className + "." + create + "(" + newName + ")"}, Nothing) = "#Error" Then Return ""
                    Dim newId As Integer = mScript.mainClass(classId).ChildProperties.Count - 1
                    mScript.eventRouter.EraseElementEvents(classId, newId)
                    mScript.mainClass(classId).ChildProperties(newId).Clear()
                    mScript.mainClass(classId).ChildProperties(newId) = retreiveObject
                    Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(newId)("Name")
                    ch.Value = newName
                End If
            Else
                'Создание нового
                If thirdLevel Then
                    elementName = GetNewDefName(className, parentId)
                    If elementName.Length = 0 Then Return ""
                    If mScript.ExecuteString({className + "." + create + "(" + parentId.ToString + ", '" + elementName + "')"}, Nothing) = "#Error" Then Return ""
                Else
                    elementName = GetNewDefName(className, -1)
                    If mScript.ExecuteString({className + "." + create + "('" + elementName + "')"}, Nothing) = "#Error" Then Return ""
                End If
            End If
        End If

        AAA()
        '2. Добавляем элемент в treeView
        'Ищем родительский узел (группу)
        Dim parNode As TreeNode = tree.SelectedNode
        If IsNothing(parNode) = False Then
            If parNode.Tag <> "GROUP" Then
                parNode = parNode.Parent
            End If
        End If
        'Добавляем узел с названием элемента
        Dim curNode As TreeNode = Nothing
        If IsNothing(parNode) Then
            curNode = tree.Nodes.Add(elementName) 'группы у элемента нет
        Else
            curNode = parNode.Nodes.Add(elementName) 'добавляем в группу элементов
            Dim shouldRearrange As Boolean = False
            If thirdLevel Then
                mScript.ExecuteString({className + "[" + parentId.ToString + ", '" + elementName + "'].Group = '" + parNode.Text + "'"}, Nothing)
                shouldRearrange = True
            ElseIf className = "Variable" Then
                mScript.csPublicVariables.lstVariables(elementName).Group = parNode.Text
            ElseIf className = "Function" Then
                Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(elementName)
                f.Group = parNode.Text
            Else
                mScript.ExecuteString({className + "['" + elementName + "'].Group = '" + parNode.Text + "'"}, Nothing)
                shouldRearrange = True
            End If
            'перенумеровываем элементы в mainClass и панели
            If parNode.Index < tree.Nodes.Count - 1 AndAlso shouldRearrange Then
                Dim blnShowHidden As Boolean = True
                If Object.Equals(tree, currentTreeView) Then blnShowHidden = chkShowHidden.Checked
                If thirdLevel Then
                    RearrangeClass(tree, mScript.mainClassHash(className), currentParentId, blnShowHidden)
                Else
                    RearrangeClass(tree, mScript.mainClassHash(className), -1, blnShowHidden)
                End If
            End If
        End If
        curNode.ImageKey = GetDefaultIcon(className, thirdLevel)
        curNode.SelectedImageKey = curNode.ImageKey
        tree.BackColor = Color.White
        curNode.Tag = "ITEM"
        If currentClassName = "Variable" Then
            FillTreeVariables(chkShowHidden.Checked, MakeListToExpand(tree))
            tree.SelectedNode = FindItemNodeByText(tree, elementName)
        ElseIf currentClassName = "Function" Then
            FillTreeFunctions(chkShowHidden.Checked, MakeListToExpand(tree))
            tree.SelectedNode = FindItemNodeByText(tree, elementName)
        Else
            If showCells Then tree.SelectedNode = curNode
            'tree.SelectedNode = curNode
        End If
        primaryNameSelection = True
        'добавляем в список
        If Not thirdLevel AndAlso className <> "Variable" AndAlso className <> "Function" Then
            cListManager.AddNameToList(className, elementName)
        End If
        AAA()
        If IsNothing(tree.SelectedNode) = False Then tree.SelectedNode.BeginEdit()

        Return "'" & elementName & "'"
    End Function

    Private Sub tsmiRemoveElement_Click(sender As Object, e As EventArgs) Handles tsmiRemoveElement.Click
        RemoveElement(currentTreeView.SelectedNode)
    End Sub


    ''' <summary>
    ''' Удаляет элемент узел из дерева, ассоциированный с группой или некоторым элементом
    ''' </summary>
    ''' <param name="nodeToRemove">Узел дерева</param>
    Public Sub RemoveElement(ByRef nodeToRemove As TreeNode)
        AAA()
        If IsNothing(nodeToRemove) Then Return
        Dim ni As NodeInfo = GetNodeInfo(nodeToRemove)
        Dim tree As TreeView = nodeToRemove.TreeView
        Dim isEmpty As Boolean = False
        iconMenuElements.Hide()
        iconMenuGroups.Hide()
        If nodeToRemove.Tag = "GROUP" Then
            'удаляем группу
            If nodeToRemove.Nodes.Count > 0 Then
                MessageBox.Show("Нельзя удалить группу пока в ней содержится хоть один элемент!", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If
            Dim cName As String
            Dim parName As String = ""
            If ni.classId = -2 Then
                cName = "Variable"
            ElseIf ni.classId = -3 Then
                cName = "Function"
            Else
                cName = mScript.mainClass(ni.classId).Names(0)
                parName = GetParentName(nodeToRemove.TreeView)
            End If
            If cGroups.RemoveGroup(cName, nodeToRemove.Text, ni.ThirdLevelNode, parName) = -1 Then
                MessageBox.Show("Ошибка при удалении элемента")
                Return
            End If
            nodeToRemove.Remove()
            If tree.Nodes.Count = 0 Then isEmpty = True
        Else
            'Удаляем элемент
            Dim child2Id As Integer = ni.GetChild2Id
            If ni.ThirdLevelNode Then
                removedObjects.AddItem(mScript.mainClass(ni.classId).ChildProperties(child2Id), ni.classId, "'" + nodeToRemove.Text + "'", child2Id)
                If mScript.ExecuteString({mScript.mainClass(ni.classId).Names(0) + "[" + child2Id.ToString + ", '" + nodeToRemove.Text + "'].Remove"}, Nothing) = "#Error" Then Return
                GlobalSeeker.CheckChild3InStruct(ni.classId, child2Id, WrapString(nodeToRemove.Text))
            ElseIf ni.classId = -2 Then
                removedObjects.AddItem(mScript.csPublicVariables.lstVariables(nodeToRemove.Text), -2, nodeToRemove.Text)
                mScript.csPublicVariables.DeleteVariable(nodeToRemove.Text)
                'Ищем удаленныю переменную в скриптах
                GlobalSeeker.CheckElementNameInStruct(-2, nodeToRemove.Text, CodeTextBox.EditWordTypeEnum.W_VARIABLE)
            ElseIf ni.classId = -3 Then
                removedObjects.AddItem(mScript.functionsHash(nodeToRemove.Text), -3, nodeToRemove.Text)
                mScript.functionsHash.Remove(nodeToRemove.Text)
                'Ищем удаленную функцию в скриптах
                GlobalSeeker.CheckFunctionInStruct(WrapString(nodeToRemove.Text))
            Else
                Dim remName As String = WrapString(nodeToRemove.Text)
                If child2Id <= -1 Then
                    MessageBox.Show("Matew Quest", "Action is not found!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                removedObjects.AddItem(mScript.mainClass(ni.classId).ChildProperties(child2Id), ni.classId, remName, GetParentId(currentClassName, tree))

                If mScript.ExecuteString({mScript.mainClass(ni.classId).Names(0) + "[" + remName + "].Remove"}, Nothing) = "#Error" Then Return

                If currentClassName = "L" AndAlso Object.Equals(currentTreeView, tree) Then
                    actionsRouter.RemoveLocation(remName)
                End If
                If ni.classId = mScript.mainClassHash("A") Then
                    GlobalSeeker.CheckActionInStruct(remName, child2Id)
                Else
                    GlobalSeeker.CheckChild2InStruct(ni.classId, remName)
                    'Вносим изменения в группы - изменяем parentId дочерних к удаленного родителя, вносим в список удаленных группы дочерних элементов. То же касается групп действий
                    If mScript.mainClass(ni.classId).LevelsCount = 2 OrElse currentClassName = "L" Then
                        cGroups.RemoveChildrenGroups(currentClassName, remName)
                    End If
                End If
            End If
            'убираем из списка
            If Not ni.ThirdLevelNode AndAlso ni.classId >= 0 Then
                cListManager.RemoveNameFromList(mScript.mainClass(ni.classId).Names(0), nodeToRemove.Text)
            End If
            If ni.classId = -2 Then
                cPanelManager.ReassociateVariablesNodes()
            ElseIf ni.classId = -3 Then
                cPanelManager.ReassociateFunctionsNodes()
            Else
                cPanelManager.ReassociateNodes(ni.classId, tree, GetParentName(tree))
            End If
            cPanelManager.UpdateAfterNodeRemoving(nodeToRemove)
            nodeToRemove.Remove()
            If currentClassName = "L" AndAlso Object.Equals(currentTreeView, tree) AndAlso IsNothing(cPanelManager.ActivePanel) = False Then
                Dim subTree As TreeView = cPanelManager.dictDefContainers(mScript.mainClassHash(currentClassName)).subTree
                actionsRouter.RetreiveActions(cPanelManager.ActivePanel.child2Name)
                FillTree("A", subTree, True, Nothing, cPanelManager.ActivePanel.child2Name)
            End If

            'cPanelManager.RemovePanel(nodeToRemove)
            If tree.Nodes.Count = 0 Then isEmpty = True
        End If
        If currentClassName = "L" AndAlso ni.classId <> mScript.mainClassHash("L") Then
            'обновляем список действий в главном дереве, так как иначе он останется старым
            FillTree("A", dictTrees(mScript.mainClassHash("A")), chkShowHidden.Checked, MakeListToExpand(tree), cPanelManager.ActivePanel.child2Name)
        End If
        'делаем недоступной кнопку удаления
        If isEmpty AndAlso Object.Equals(currentTreeView, tree) = False Then
            'дочернее дерево
            Dim parent As Control = tree.Parent
            Dim btnRemoveSub As Button = parent.Controls.Find("RemoveSubElement", True)(0)
            btnRemoveSub.Enabled = False
            If codeBoxPanel.Visible Then
                codeBoxChangeOwner(Nothing)
                codeBoxPanel.Hide()
            End If
        End If
        AAA()
    End Sub

    Private Sub tsmiDuplicate_Click(sender As Object, e As EventArgs) Handles tsmiDuplicate.Click
        Duplicate(currentTreeView.SelectedNode)
    End Sub

    ''' <summary>
    ''' Создает полную копию элемента, ассоциированного с заданным узлом дерева
    ''' </summary>
    ''' <param name="nodeToDuplicate">Узел дерева для дублирования</param>
    Public Sub Duplicate(ByRef nodeToDuplicate As TreeNode)
        If IsNothing(nodeToDuplicate) Then Return
        If nodeToDuplicate.Tag = "GROUP" Then Return

        Dim ni As NodeInfo = GetNodeInfo(nodeToDuplicate)
        Dim child2Id As Integer = ni.GetChild2Id
        Dim child3Id As Integer = ni.GetChild3Id(child2Id)
        Dim className As String = ""
        If ni.classId = -2 Then
            className = "Variable"
        ElseIf ni.classId = -3 Then
            className = "Function"
        Else
            className = mScript.mainClass(ni.classId).Names(0)
        End If
        If ni.nodeChild2Name.Length = 0 Then Return

        'Получаем новое имя
        Dim elName As String
        If ni.ThirdLevelNode Then
            elName = GetNewDefName(className, IIf(currentParentName.Length < 0, currentParentId, child2Id))
        Else
            elName = GetNewDefName(className, -1)
        End If
        If elName.Length = 0 Then Return
        'Непосредственное создание элемента
        If ni.classId = -2 Then
            'дублирование переменной
            If mScript.csPublicVariables.DuplicateVariable(nodeToDuplicate.Text, elName) = -1 Then Return
        ElseIf ni.classId = -3 Then
            'дублирование функции
            Dim fOld As MatewScript.FunctionInfoType = Nothing
            Try
                fOld = mScript.functionsHash(nodeToDuplicate.Text)
            Catch ex As Exception
                Return
            End Try
            Dim fNew As New MatewScript.FunctionInfoType
            fNew.Description = fOld.Description
            fNew.Group = fOld.Group
            fNew.Hidden = fOld.Hidden
            fNew.Icon = fOld.Icon
            If IsNothing(fOld.ValueDt) = False Then
                fNew.ValueDt = CopyCodeDataArray(fOld.ValueDt)
                fNew.ValueExecuteDt = mScript.PrepareBlock(fNew.ValueDt)
            End If
            mScript.functionsHash.Add(elName, fNew)
        Else
            'создание элемента (пока не дублирование)
            Dim Create As String = "Create"
            If className = "A" Then Create = "Add"
            If ni.ThirdLevelNode Then
                If mScript.ExecuteString({className + "." + Create + "(" + child2Id.ToString + ", '" + elName + "')"}, Nothing) = "#Error" Then Return 'создаем элемент 3 уровня
            Else
                If mScript.ExecuteString({className + "." + Create + "('" + elName + "')"}, Nothing) = "#Error" Then Return 'создаем элемент 2 уровня
                'Если дублируется локация - дублируем и все действия
                If className = "L" Then actionsRouter.DuplicateActions("'" + elName + "'")
            End If
        End If

        '2. Добавляем элемент в treeView
        'Ищем родительский узел (группу)
        Dim tree As TreeView = nodeToDuplicate.TreeView
        Dim parNode As TreeNode = nodeToDuplicate.Parent

        'Добавляем узел с названием элемента
        Dim newNode As TreeNode = Nothing
        If IsNothing(parNode) Then
            newNode = tree.Nodes.Add(elName) 'группы у элемента нет
        ElseIf ni.classId >= 0 Then
            newNode = parNode.Nodes.Add(elName) 'добавляем в группу элементов
            If ni.ThirdLevelNode Then
                mScript.ExecuteString({className + "[" + child2Id.ToString + ", '" + elName + "'].Group = '" + parNode.Text + "'"}, Nothing)
            Else
                mScript.ExecuteString({className + "['" + elName + "'].Group = '" + parNode.Text + "'"}, Nothing)
            End If
        Else
            newNode = parNode.Nodes.Add(elName) 'добавляем в группу элементов (переменные или функции)
        End If
        newNode.ImageKey = nodeToDuplicate.ImageKey
        newNode.SelectedImageKey = nodeToDuplicate.SelectedImageKey
        newNode.Tag = "ITEM"
        'копируем все свойства на новый элемент
        If ni.classId >= 0 Then
            If Not ni.ThirdLevelNode Then
                Dim srcList As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(ni.classId).ChildProperties(child2Id)
                Dim destList As New SortedList(Of String, MatewScript.ChildPropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
                Dim destChildId As Integer = mScript.mainClass(ni.classId).ChildProperties.Count - 1
                For Each srcProp As KeyValuePair(Of String, MatewScript.ChildPropertiesInfoType) In srcList
                    If srcProp.Key = "Name" Then
                        destList.Add(srcProp.Key, New MatewScript.ChildPropertiesInfoType With {.Hidden = srcProp.Value.Hidden, .ThirdLevelProperties = srcProp.Value.ThirdLevelProperties, _
                                                                                         .Value = "'" + elName + "'"})
                    Else

                        'If srcProp.Value.eventId < 1 Then
                        '    destList.Add(srcProp.Key, srcProp.Value)
                        'Else
                        '    Dim newP As MatewScript.ChildPropertiesInfoType = srcProp.Value
                        '    newP.eventId = mScript.eventRouter.DuplicateEvent(newP.eventId, mScript.mainClass(ni.classId).ChildProperties(destChildId)(srcProp.Key).eventId)
                        '    destList.Add(srcProp.Key, newP)
                        'End If
                        destList.Add(srcProp.Key, srcProp.Value.Clone(True))
                    End If
                Next
                mScript.mainClass(ni.classId).ChildProperties(destChildId) = destList
            Else
                Dim newChildId As Integer = mScript.mainClass(ni.classId).ChildProperties(child2Id)("Name").ThirdLevelProperties.Count - 1
                For i As Integer = 0 To mScript.mainClass(ni.classId).ChildProperties(child2Id).Count - 1
                    Dim pName As String = mScript.mainClass(ni.classId).ChildProperties(child2Id).ElementAt(i).Key
                    Dim pValue As MatewScript.ChildPropertiesInfoType = mScript.mainClass(ni.classId).ChildProperties(child2Id).ElementAt(i).Value

                    If pName = "Name" Then
                        Dim prop As MatewScript.ChildPropertiesInfoType = mScript.mainClass(ni.classId).ChildProperties(child2Id)("Name")
                        prop.ThirdLevelProperties(newChildId) = "'" + elName + "'"
                        mScript.mainClass(ni.classId).ChildProperties(child2Id)("Name") = prop
                    Else
                        SetPropertyValue(ni.classId, pName, pValue.ThirdLevelProperties(child3Id), child2Id, newChildId) 'событие само продублируется при необходимости
                    End If
                Next
            End If

            If IsNothing(parNode) = False AndAlso parNode.Index < tree.Nodes.Count - 1 Then
                'перенумеровываем элементы в mainClass и панели
                Dim blnShowHidden As Boolean = True
                If Object.Equals(tree, currentTreeView) Then blnShowHidden = chkShowHidden.Checked
                Dim classId As Integer = mScript.mainClassHash(className)
                If ni.ThirdLevelNode Then
                    RearrangeClass(tree, classId, child2Id, blnShowHidden)
                Else
                    RearrangeClass(tree, classId, -1, blnShowHidden)
                End If
            End If

            tree.SelectedNode = newNode
            'добавляем в список
            If Not ni.ThirdLevelNode Then
                cListManager.AddNameToList(className, elName)
                If currentClassName = "L" AndAlso ni.classId = mScript.mainClassHash("A") Then
                    FillTree("A", dictTrees(mScript.mainClassHash("A")), chkShowHidden.Checked, MakeListToExpand(tree), cPanelManager.ActivePanel.child2Name)
                End If
            End If
        ElseIf ni.classId = -2 Then
            Dim lst As New List(Of String)
            For i As Integer = 0 To tree.Nodes.Count - 1
                If tree.Nodes(i).IsExpanded Then lst.Add(tree.Nodes(i).Text)
            Next
            FillTreeVariables(chkShowHidden.Checked, lst)
            tree.SelectedNode = FindItemNodeByText(tree, elName)
        ElseIf ni.classId = -3 Then
            Dim lst As New List(Of String)
            For i As Integer = 0 To tree.Nodes.Count - 1
                If tree.Nodes(i).IsExpanded Then lst.Add(tree.Nodes(i).Text)
            Next
            FillTreeFunctions(chkShowHidden.Checked, lst)
            tree.SelectedNode = FindItemNodeByText(tree, elName)
        End If

    End Sub

    Private Sub tsmiExpand_Click(sender As Object, e As EventArgs) Handles tsmiExpand.Click
        currentTreeView.ExpandAll()
    End Sub

    Private Sub СвернутьВсеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СвернутьВсеToolStripMenuItem.Click
        currentTreeView.CollapseAll()
    End Sub

    Private Sub tsmiExcludeFromGroup_Click(sender As Object, e As EventArgs) Handles tsmiExcludeFromGroup.Click
        ExcludeFromGroup(currentTreeView)
    End Sub

    ''' <summary>
    ''' Выводит из элемент из группы (если выбран элемент) или выводит все элементы из группы (если выбрана группа)
    ''' </summary>
    ''' <param name="tree">ссылка на дерево</param>
    Public Sub ExcludeFromGroup(ByRef tree As TreeView)
        If IsNothing(tree.SelectedNode) Then Return
        Dim ni As NodeInfo = GetNodeInfo(tree.SelectedNode)
        Dim child2Id As Integer = ni.GetChild2Id

        ''Сохраняем список развернутых узлов для последующего восстановления
        'Dim lstExpanded As New List(Of String)
        'For i As Integer = 0 To tree.Nodes.Count - 1
        '    Dim nod As TreeNode = tree.Nodes(i)
        '    If nod.IsExpanded Then lstExpanded.Add(nod.Text)
        'Next

        Dim excludeAllFromGroup As Boolean = (tree.SelectedNode.Tag = "GROUP")

        Dim nodeText As String
        If Not excludeAllFromGroup Then
            '1 - вывести узел из группы
            If IsNothing(tree.SelectedNode.Parent) Then Return
            nodeText = tree.SelectedNode.Text
            If ni.classId = -2 Then
                Dim v As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables(ni.nodeChild2Name)
                v.Group = ""
            ElseIf ni.classId = -3 Then
                Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(ni.nodeChild2Name)
                f.Group = ""
            Else
                SetPropertyValue(ni.classId, "Group", "", child2Id, ni.GetChild3Id(child2Id)) 'убираем группу из элемента
            End If
        Else
            '2 - вывести из данной группы все узлы
            If tree.SelectedNode.Nodes.Count = 0 Then Return
            nodeText = tree.SelectedNode.Text
            Dim groupName As String = "'" + nodeText + "'"
            If Not ni.ThirdLevelNode Then
                'в дереве элементы 2 порядка, выводим все элементы из группы
                If ni.classId = -2 Then
                    For i As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
                        Dim v As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables.ElementAt(i).Value
                        If v.Group = nodeText Then v.Group = ""
                    Next
                ElseIf ni.classId = -3 Then
                    For i As Integer = 0 To mScript.functionsHash.Count - 1
                        Dim f As MatewScript.FunctionInfoType = mScript.functionsHash.ElementAt(i).Value
                        If f.Group = nodeText Then f.Group = ""
                    Next
                Else
                    For i As Integer = 0 To mScript.mainClass(ni.classId).ChildProperties.Count - 1
                        If mScript.mainClass(ni.classId).ChildProperties(i)("Group").Value = groupName Then
                            SetPropertyValue(ni.classId, "Group", "", i)
                        End If
                    Next
                End If
            Else
                'в дереве элементы 3 порядка, выводим все элементы из группы
                Dim arrValues() As String = mScript.mainClass(ni.classId).ChildProperties(child2Id)("Group").ThirdLevelProperties
                If IsNothing(arrValues) OrElse arrValues.Count = 0 Then Return
                For i As Integer = 0 To arrValues.Count - 1
                    If arrValues(i) = groupName Then
                        SetPropertyValue(ni.classId, "Group", "", child2Id, i)
                    End If
                Next
            End If
        End If

        If Not ni.ThirdLevelNode Then
            'это основные элементы 2-го порядка либо вторичные к ним (например, Действия для класса Локаций)
            If ni.classId = -2 Then
                FillTreeVariables(chkShowHidden.Checked, MakeListToExpand(tree))
            ElseIf ni.classId = -3 Then
                FillTreeFunctions(chkShowHidden.Checked, MakeListToExpand(tree))
            Else
                Dim className As String = mScript.mainClass(ni.classId).Names(0)
                FillTree(className, tree, chkShowHidden.Checked, MakeListToExpand(tree), GetParentName(tree))
                If currentClassName = "L" AndAlso ni.classId = mScript.mainClassHash("A") Then
                    FillTree("A", dictTrees(mScript.mainClassHash("A")), chkShowHidden.Checked, MakeListToExpand(tree), cPanelManager.ActivePanel.child2Name)
                End If

            End If
        Else
            'это элементы 3 порядка
            FillTree(mScript.mainClass(ni.classId).Names(0), tree, True, MakeListToExpand(tree), ni.nodeChild2Name)
        End If
        'перенумеровываем элементы в mainClass и панели
        If ni.classId >= 0 Then
            Dim blnShowHidden As Boolean = True
            If Object.Equals(tree, currentTreeView) Then blnShowHidden = chkShowHidden.Checked
            If ni.ThirdLevelNode Then
                RearrangeClass(tree, ni.classId, child2Id, True)
            Else
                RearrangeClass(tree, ni.classId, -1, blnShowHidden)
            End If
        End If

        'Выделяем 
        Dim n As TreeNode = FindRootNodeByText(nodeText, tree, excludeAllFromGroup)
        If IsNothing(n) = False Then tree.SelectedNode = n
    End Sub


    ''' <summary>
    ''' Возвращает имя класса на русском (из строки ресурсов)
    ''' </summary>
    ''' <param name="classId">Id класса,которое надо получить</param>
    ''' <param name="isThirdLevel">Надо ли получить имя элемента 3-го уровня</param>
    Public Function GetTranslatedName(ByVal classId As Integer, Optional ByVal isThirdLevel As Boolean = False) As String
        If classId = -2 Then
            Return My.Resources.Variables
        ElseIf classId = -3 Then
            Return My.Resources.Functions
        End If
        Dim className As String = mScript.mainClass(classId).Names(0)

        Select Case className
            Case "L"
                Return My.Resources.Locations
            Case "A"
                Return My.Resources.Actions
            Case "O"
                Return My.Resources.Objects
            Case "M"
                If isThirdLevel Then
                    Return My.Resources.MenuItems
                Else
                    Return My.Resources.Menus
                End If
            Case "T"
                Return My.Resources.Timers
            Case "H"
                Return My.Resources.Heroes
            Case "Map"
                If isThirdLevel Then
                    Return My.Resources.MapCells
                Else
                    Return My.Resources.Maps
                End If
            Case "Med"
                If isThirdLevel Then
                    Return My.Resources.MediaFiles
                Else
                    Return My.Resources.newMediaList
                End If
            Case "Mg"
                If isThirdLevel Then
                    Return My.Resources.Magics
                Else
                    Return My.Resources.MagicBooks
                End If
            Case "Ab"
                If isThirdLevel Then
                    Return My.Resources.Abilities
                Else
                    Return My.Resources.AbilitySets
                End If
            Case "Army"
                If isThirdLevel Then
                    Return My.Resources.Units
                Else
                    Return My.Resources.Armies
                End If
            Case Else
                If isThirdLevel Then
                    Return mScript.mainClass(classId).Names.Last + "_subitem"
                Else
                    Return mScript.mainClass(classId).Names.Last
                End If
        End Select
    End Function

    ''' <summary>
    ''' Возвращает имя класса на русском (из строки ресурсов) для надписи вроде "Новая локация"
    ''' </summary>
    ''' <param name="classId">Id класса,которое надо получить</param>
    ''' <param name="isThirdLevel">Надо ли получить имя элемента 3-го уровня</param>
    Public Function GetTranslationAddNew(ByVal classId As Integer, Optional ByVal isThirdLevel As Boolean = False) As String
        If classId = -2 Then
            Return My.Resources.newVariable
        ElseIf classId = -3 Then
            Return My.Resources.newFunction
        End If
        Dim className As String = mScript.mainClass(classId).Names(0)

        Select Case className
            Case "L"
                Return My.Resources.newLocation
            Case "A"
                Return My.Resources.newAction
            Case "O"
                Return My.Resources.newObject
            Case "M"
                If isThirdLevel Then
                    Return My.Resources.newMenuItem
                Else
                    Return My.Resources.newMenu
                End If
            Case "T"
                Return My.Resources.newTimer
            Case "Map"
                If isThirdLevel Then
                    Return My.Resources.newMapCell
                Else
                    Return My.Resources.newMap
                End If
            Case "Med"
                If isThirdLevel Then
                    Return My.Resources.newMediaFile
                Else
                    Return My.Resources.newMediaList
                End If
            Case "H"
                Return My.Resources.newHero
            Case "Mg"
                If isThirdLevel Then
                    Return My.Resources.newMagic
                Else
                    Return My.Resources.newMagicBook
                End If
            Case "Ab"
                If isThirdLevel Then
                    Return My.Resources.newAbility
                Else
                    Return My.Resources.newAbilitySet
                End If
            Case "Army"
                If isThirdLevel Then
                    Return My.Resources.newUnit
                Else
                    Return My.Resources.newArmy
                End If
            Case Else
                Return mScript.mainClass(classId).Names.Last
        End Select
    End Function

    ''' <summary>
    ''' Возвращает стандартное имя для создания нового элемента (локации, действия ..). Возвращает имя в виде строки [Имя_элемента Х]
    ''' </summary>
    ''' <param name="className">Имя класса, для которого получается новое имя</param>
    Public Function GetNewDefName(ByVal className As String, ByVal parentId As Integer) As String
        Dim classId As Integer = -1

        Dim defElementName As String = ""
        Select Case className
            Case "Variable"
                defElementName = "Переменная"
                classId = -2
            Case "Function"
                defElementName = "Функция"
                classId = -3
            Case "L"
                defElementName = "Локация "
            Case "A"
                defElementName = "Действие "
            Case "O"
                defElementName = "Предмет "
            Case "M"
                If parentId >= 0 Then
                    defElementName = "Пункт меню "
                Else
                    defElementName = "Меню "
                End If
            Case "T"
                defElementName = "Счетчик "
            Case "H"
                defElementName = "Персонаж "
            Case "Map"
                If parentId >= 0 Then
                    defElementName = "Клетка карты "
                Else
                    defElementName = "Карта "
                End If
            Case "Med"
                If parentId >= 0 Then
                    defElementName = "Аудиофайл "
                Else
                    defElementName = "Список воспроизведения "
                End If
            Case "Mg"
                If parentId >= 0 Then
                    defElementName = "Магия "
                Else
                    defElementName = "Книга магий "
                End If
            Case "Ab"
                If parentId >= 0 Then
                    defElementName = "Способность "
                Else
                    defElementName = "Набор способностей "
                End If
            Case "Army"
                If parentId >= 0 Then
                    defElementName = "Боец "
                Else
                    defElementName = "Армия "
                End If
            Case Else
                Dim cId As Integer = mScript.mainClassHash(className)
                defElementName = mScript.mainClass(cId).Names.Last + " "
        End Select
        If defElementName.Length = 0 Then
            MessageBox.Show("Не удалось получить новое стандартное имя!")
            Return ""
        End If

        'Создаем новое имя 
        Dim i As Integer = 1
        If classId = -2 Then
            'переменные
            Do While mScript.csPublicVariables.lstVariables.ContainsKey(defElementName + CStr(i))
                i += 1
            Loop
        ElseIf classId = -3 Then
            'функции
            Do While mScript.functionsHash.ContainsKey(defElementName + CStr(i))
                i += 1
            Loop
        Else
            'Получаем classId класса className
            If mScript.mainClassHash.TryGetValue(className, classId) = False Then Return ""

            If parentId >= 0 Then
                Do While GetThirdChildIdByName("'" + defElementName + CStr(i) + "'", parentId, mScript.mainClass(classId).ChildProperties) >= 0
                    i += 1
                Loop
            Else
                Do While GetSecondChildIdByName("'" + defElementName + CStr(i) + "'", mScript.mainClass(classId).ChildProperties) >= 0
                    i += 1
                Loop
            End If
        End If

        Return defElementName + CStr(i)
    End Function
#End Region

#Region "TreeView и панель управления для TreeView"

    ''' <summary> Делегат события при нажати на кнопку "Установки по умолчанию"</summary>
    Public Sub btnShowSettings_Click(sender As Object, e As EventArgs)
        Dim selNode As TreeNode = Nothing
        If IsNothing(currentTreeView) = False Then
            selNode = currentTreeView.SelectedNode
        End If
        If IsNothing(selNode) = False Then
            If selNode.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then selNode.ForeColor = DEFAULT_COLORS.NodeForeColor
            If selNode.BackColor <> DEFAULT_COLORS.ControlTransparentBackground Then selNode.BackColor = DEFAULT_COLORS.ControlTransparentBackground
            currentTreeView.SelectedNode = Nothing
        End If
        sender.BackColor = DEFAULT_COLORS.NodeSelBackColor

        If questEnvironment.EnabledEvents = False Then Return
        Log.PrintToLog("btnShowSettings_Click: " + currentClassName)
        'If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False AndAlso _
        '    cPanelManager.ActivePanel.child2Id > -1 Then actionsRouter.SaveActions(cPanelManager.ActivePanel.child2Id)

        cPanelManager.OpenPanel(Nothing, mScript.mainClassHash(currentClassName), "") 'Открытие настроек по умолчанию
    End Sub

    ''' <summary> Делегат события при нажати на кнопку "Скрыть"</summary>
    Private Sub btnMakeHidden_Click(sender As Object, e As EventArgs)
        If IsNothing(currentTreeView.SelectedNode) Then Return
        'Dim thirdLevel As Boolean = IsThirdLevelTree(currentTreeView)
        Dim ni As NodeInfo = GetNodeInfo(currentTreeView.SelectedNode)
        If currentTreeView.SelectedNode.Tag = "GROUP" Then
            'показать/скрыть группу

            Dim gId As Integer = cGroups.GetGroupIdByName(currentClassName, currentTreeView.SelectedNode.Text, ni.ThirdLevelNode, currentParentName)
            If gId = -1 Then Return
            Dim curG As clsGroups.clsGroupInfo = cGroups.dictGroups(currentClassName)(gId)
            curG.Hidden = Not curG.Hidden
            If curG.Hidden Then
                If ni.classId >= 0 Then
                    'удаляем из списка все элементы с именами = узлам группы
                    For Each n As TreeNode In currentTreeView.SelectedNode.Nodes
                        cListManager.RemoveNameFromList(currentClassName, n.Text)
                    Next
                End If

                If chkShowHidden.Checked Then
                    'скрытые показывать
                    With currentTreeView.SelectedNode
                        .ForeColor = DEFAULT_COLORS.NodeHiddenForeColor
                        .ImageKey = "groupHidden.png"
                        .SelectedImageKey = .ImageKey
                    End With
                Else
                    'скрытые прятать
                    currentTreeView.SelectedNode.Remove()
                End If
            Else
                If ni.classId >= 0 Then
                    'добавляем в список все элементы с именами = узлам группы
                    For Each n As TreeNode In currentTreeView.SelectedNode.Nodes
                        cListManager.AddNameToList(currentClassName, n.Text)
                    Next
                End If
                If chkShowHidden.Checked Then
                    'скрытые показывать
                    With currentTreeView.SelectedNode
                        .ForeColor = DEFAULT_COLORS.NodeSelForeColor
                        .ImageKey = cGroups.dictGroups(currentClassName)(gId).iconName
                        .SelectedImageKey = .ImageKey
                    End With
                End If
            End If
        Else
            'показать/скрыть элемент
            'Dim classId As Integer = mScript.mainClassHash(currentClassName)
            Dim nodeText As String = currentTreeView.SelectedNode.Text
            If ni.classId = -2 Then
                'переменная
                Dim v As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables(nodeText)
                v.Hidden = Not v.Hidden
                If v.Hidden Then
                    If chkShowHidden.Checked Then
                        'скрытые показывать
                        With currentTreeView.SelectedNode
                            .ForeColor = DEFAULT_COLORS.NodeHiddenForeColor
                            .ImageKey = "itemHidden.png"
                            .SelectedImageKey = .ImageKey
                        End With
                    Else
                        'скрытые прятать
                        currentTreeView.SelectedNode.Remove()
                    End If
                Else
                    If chkShowHidden.Checked Then
                        'скрытые показывать
                        With currentTreeView.SelectedNode
                            .ForeColor = DEFAULT_COLORS.NodeSelForeColor
                            Dim iName As String = v.Icon
                            If String.IsNullOrEmpty(iName) Then iName = GetDefaultIcon(ni)
                            .ImageKey = iName
                            .SelectedImageKey = .ImageKey
                        End With
                    End If
                End If
            ElseIf ni.classId = -3 Then
                'функция
                Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(nodeText)
                f.Hidden = Not f.Hidden
                If f.Hidden Then
                    If chkShowHidden.Checked Then
                        'скрытые показывать
                        With currentTreeView.SelectedNode
                            .ForeColor = DEFAULT_COLORS.NodeHiddenForeColor
                            .ImageKey = "itemHidden.png"
                            .SelectedImageKey = .ImageKey
                        End With
                    Else
                        'скрытые прятать
                        currentTreeView.SelectedNode.Remove()
                    End If
                Else
                    'элемент
                    If chkShowHidden.Checked Then
                        'скрытые показывать
                        With currentTreeView.SelectedNode
                            .ForeColor = DEFAULT_COLORS.NodeSelForeColor
                            Dim iName As String = f.Icon
                            If String.IsNullOrEmpty(iName) Then iName = GetDefaultIcon(ni)
                            .ImageKey = iName
                            .SelectedImageKey = .ImageKey
                        End With
                    End If
                End If
            Else
                For i As Integer = 0 To mScript.mainClass(ni.classId).ChildProperties.Count - 1
                    Dim p As MatewScript.ChildPropertiesInfoType = mScript.mainClass(ni.classId).ChildProperties(i)("Name")
                    If mScript.PrepareStringToPrint(p.Value, Nothing) = nodeText Then
                        p.Hidden = Not p.Hidden
                        If p.Hidden Then
                            If chkShowHidden.Checked Then
                                'скрытые показывать
                                With currentTreeView.SelectedNode
                                    .ForeColor = DEFAULT_COLORS.NodeHiddenForeColor
                                    .ImageKey = "itemHidden.png"
                                    .SelectedImageKey = .ImageKey
                                End With
                            Else
                                'скрытые прятать
                                currentTreeView.SelectedNode.Remove()
                            End If
                            'удаляем имя в списке
                            cListManager.RemoveNameFromList(currentClassName, mScript.PrepareStringToPrint(p.Value, Nothing, False))
                        Else
                            If chkShowHidden.Checked Then
                                'скрытые показывать
                                Dim child2Id As Integer, child3Id As Integer = -1
                                If currentParentName.Length > 0 Then
                                    child2Id = GetSecondChildIdByName(currentParentName, mScript.mainClass(ni.classId).ChildProperties)
                                    child3Id = GetThirdChildIdByName("'" + nodeText + "'", child2Id, mScript.mainClass(ni.classId).ChildProperties)
                                Else
                                    child2Id = GetSecondChildIdByName("'" + nodeText + "'", mScript.mainClass(ni.classId).ChildProperties)
                                End If
                                'добавляем имя в список
                                cListManager.AddNameToList(currentClassName, nodeText)
                                With currentTreeView.SelectedNode
                                    .ForeColor = DEFAULT_COLORS.NodeSelForeColor
                                    Dim iName As String = mScript.mainClass(ni.classId).ChildProperties(child2Id)("Icon").Value
                                    iName = iName.Substring(1, iName.Length - 2)
                                    .ImageKey = iName
                                    .SelectedImageKey = .ImageKey
                                End With
                            End If
                        End If
                        Exit For
                    End If
                Next i
            End If
        End If
        If btnMakeHidden.Text = "Скрыть" Then
            btnMakeHidden.Text = "Отобразить"
        Else
            btnMakeHidden.Text = "Скрыть"
        End If
    End Sub

    ''' <summary>
    ''' Делегат события клика по кнопке копировать в...
    ''' </summary>
    ''' <param name="sender">кнопка копировать в...</param>
    ''' <param name="e">событие</param>
    Private Sub btnCopyTo_Click(sender As Object, e As EventArgs)
        If currentParentName.Length = 0 Then
            MessageBox.Show("Копирование в данный момент невозможно: не распознан источник копирования.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        Dim classId As Integer = mScript.mainClassHash(currentClassName)
        If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then
            MessageBox.Show("Копирование в данный момент невозможно: нет ни одного элемента для копирования.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If
        Dim srcChild2 As Integer = currentParentId
        If mScript.mainClass(classId).LevelsCount = 2 AndAlso (IsNothing(mScript.mainClass(classId).ChildProperties(srcChild2)("Name").ThirdLevelProperties) OrElse _
                                                               mScript.mainClass(classId).ChildProperties(srcChild2)("Name").ThirdLevelProperties.Count = 0) Then
            MessageBox.Show("Копирование в данный момент невозможно: нет ни одного элемента для копирования.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        Dim treeDest As TreeView = dlgCopyTo.treeCopy
        treeDest.ImageList = imgLstGroupIcons
        If currentClassName = "A" Then
            classId = mScript.mainClassHash("L")
            FillTree("L", treeDest, chkShowHidden.Checked, Nothing, "", True)
        Else
            FillTree(currentClassName, treeDest, chkShowHidden.Checked, Nothing, "", True)
        End If

        Dim res As DialogResult = dlgCopyTo.ShowDialog(Me)
        If res = Windows.Forms.DialogResult.Cancel OrElse IsNothing(treeDest.SelectedNode) OrElse treeDest.SelectedNode.Tag = "GROUP" Then Return
        Dim txt As String = treeDest.SelectedNode.Text

        Dim blnReplace As Boolean = dlgCopyTo.chkReplace.Checked
        Dim destChild2 As Integer = GetSecondChildIdByName(WrapString(txt), mScript.mainClass(classId).ChildProperties)
        'копирование групп
        cGroups.GroupsCopyTo(currentClassName, currentParentName, WrapString(txt), blnReplace)
        If currentClassName = "A" Then
            actionsRouter.DuplicateActions(WrapString(txt), blnReplace)
        Else
            If blnReplace Then
                'Копирование с заменой. Удаляем все дочерние элементы в элементе назначения.
                If IsNothing(mScript.mainClass(classId).ChildProperties(destChild2)("Name").ThirdLevelProperties) = False Then
                    For i As Integer = mScript.mainClass(classId).ChildProperties(destChild2)("Name").ThirdLevelProperties.Count - 1 To 0 Step -1
                        RemoveObject(classId, {destChild2.ToString, i})
                    Next
                End If
            End If

            Dim cnt As Integer = mScript.mainClass(classId).ChildProperties(srcChild2)("Name").ThirdLevelProperties.Count 'количество копируемых элементов
            Dim initCnt As Integer = 0 'исходное количество элементов в родителе, к которым добавляем cnt новых элементов
            If blnReplace = False Then
                If IsNothing(mScript.mainClass(classId).ChildProperties(destChild2)("Name").ThirdLevelProperties) = False AndAlso mScript.mainClass(classId).ChildProperties(destChild2)("Name").ThirdLevelProperties.Count > 0 Then
                    initCnt = mScript.mainClass(classId).ChildProperties(destChild2)("Name").ThirdLevelProperties.Count
                End If
            End If
            For propId As Integer = 0 To mScript.mainClass(classId).ChildProperties(srcChild2).Count - 1
                'перебираем все свойства. В каждом свойстве выбираем ThirdLevelProperties и ThirdLevelEventId ресурса и копируем их в элемент назначения
                Dim srcP As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(srcChild2).ElementAt(propId).Value 'откуда копируем
                Dim destP As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(destChild2).ElementAt(propId).Value 'куда копируем
                If blnReplace Then
                    'Копирование с заменой. В элементе назначения на данный момент уже пусто
                    'Копируем массивы ThirdLevelProperties и ThirdLevelEventId
                    ReDim destP.ThirdLevelProperties(cnt - 1)
                    Array.Copy(srcP.ThirdLevelProperties, destP.ThirdLevelProperties, cnt)
                    ReDim destP.ThirdLevelEventId(cnt - 1)
                    Array.Copy(srcP.ThirdLevelEventId, destP.ThirdLevelEventId, cnt)
                    'Создаем дубликаты событий
                    For i As Integer = 0 To cnt - 1
                        If destP.ThirdLevelEventId(i) > 0 Then destP.ThirdLevelEventId(i) = mScript.eventRouter.DuplicateEvent(destP.ThirdLevelEventId(i))
                    Next
                Else
                    'Копирование с добавлением к существующим элементам. На данный момент в элементе назначения имеется initCnt своих дочерних элементов
                    'Расширяем массивы ThirdLevelProperties и ThirdLevelEventId
                    ReDim Preserve destP.ThirdLevelProperties(initCnt - 1 + cnt)
                    Array.Copy(srcP.ThirdLevelProperties, 0, destP.ThirdLevelProperties, initCnt, cnt)
                    ReDim Preserve destP.ThirdLevelEventId(initCnt - 1 + cnt)
                    Array.Copy(srcP.ThirdLevelEventId, 0, destP.ThirdLevelEventId, initCnt, cnt)
                    'Создаем дубликаты событий
                    For i As Integer = initCnt To initCnt + cnt - 1
                        If destP.ThirdLevelEventId(i) > 0 Then destP.ThirdLevelEventId(i) = mScript.eventRouter.DuplicateEvent(destP.ThirdLevelEventId(i))
                    Next
                    If mScript.mainClass(classId).ChildProperties(srcChild2).ElementAt(propId).Key = "Name" Then
                        'Если текущее свойство - Name, то проверяем нет ли совпадения имен с элементами, которые были раньше
                        Dim lst As List(Of String) = destP.ThirdLevelProperties.ToList
                        For i As Integer = initCnt To initCnt + cnt - 1
                            If i = 0 Then Continue For
                            Dim pos As Integer = lst.LastIndexOf(lst(i), i - 1)
                            If pos > -1 Then
                                'найден элемент с таким же именем. Создаем новое имя
                                destP.ThirdLevelProperties(i) = WrapString(GetNewDefName(currentClassName, destChild2))
                            End If
                        Next
                    End If
                End If
            Next propId

        End If

        cPanelManager.FindAndOpen(classId, destChild2, -1, -1, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
    End Sub

    ''' <summary>
    ''' Заполняет нужный treeView группами и элементами
    ''' </summary>
    ''' <param name="className">Имя класса, для которого заполняется TreeView</param>
    ''' <returns>0  при успехе, -1 при ошибке</returns>
    Public Function FillTree(ByVal className As String, ByRef tree As TreeView, ByVal showHidden As Boolean, Optional lstToExpand As List(Of String) = Nothing, _
                             Optional ByVal parentName As String = "", Optional forcedBuildSecondLevel As Boolean = False) As Integer
        If IsNothing(tree) Then Return -1
        Dim classId As Integer = mScript.mainClassHash(className)
        Dim parentId As Integer = -1
        If parentName.Length > 0 Then
            If className = "A" Then
                parentId = GetSecondChildIdByName(parentName, mScript.mainClass(mScript.mainClassHash("L")).ChildProperties)
            Else
                parentId = GetSecondChildIdByName(parentName, mScript.mainClass(classId).ChildProperties)
            End If
        End If
        Dim thirdLevel As Boolean = False
        If Not forcedBuildSecondLevel AndAlso parentId <> -2 Then thirdLevel = IsThirdLevelTree(tree, classId)

        If thirdLevel Then showHidden = True
        Dim isMainTree As Boolean = Object.Equals(currentTreeView, tree) 'главное дерево заполняется или второстепенное
        'получаем classId
        tree.BeginUpdate()
        tree.Nodes.Clear()
        'заполняем treeView элементами 2-го порядка
        Dim lstGroups As List(Of clsGroups.clsGroupInfo) = cGroups.dictGroups(className)
        Dim prop() As SortedList(Of String, MatewScript.ChildPropertiesInfoType) = mScript.mainClass(classId).ChildProperties

        Log.PrintToLog("FillTree: " + className + ", " + IIf(isMainTree, "Main", "Sub") + " parentId: " + CStr(parentId))

        If IsNothing(prop) OrElse prop.Count = 0 Then
            tree.EndUpdate()
            'If isMainTree Then cPanelManager.ReassociateNodes(classId, tree, parentId)
            Return 0
        End If

        Dim cnt As Integer = prop.Count
        If thirdLevel Then
            If IsNothing(prop(parentId)("Name").ThirdLevelProperties) = False Then
                cnt = prop(parentId)("Name").ThirdLevelProperties.Count
            Else
                cnt = 0
            End If
            If cnt <= 0 Then
                tree.EndUpdate()
                'If isMainTree Then cPanelManager.ReassociateNodes(classId, tree, parentId)
                Return 0
            End If
        End If
        For i As Integer = 0 To cnt - 1
            Dim imgKey As String
            If showHidden = False AndAlso prop(i)("Name").Hidden Then Continue For
            If thirdLevel = False AndAlso prop(i)("Name").Hidden Then
                imgKey = "'itemHidden.png'"
            ElseIf thirdLevel Then
                imgKey = prop(parentId)("Icon").ThirdLevelProperties(i)
            Else
                imgKey = prop(i)("Icon").Value
            End If
            If String.IsNullOrEmpty(imgKey) OrElse imgKey.Length <= 2 Then imgKey = "'" + GetDefaultIcon(className, thirdLevel) + "'"
            imgKey = imgKey.Substring(1, imgKey.Length - 2)
            Dim n As TreeNode
            Dim gName As String
            If thirdLevel Then
                n = New TreeNode(mScript.PrepareStringToPrint(prop(parentId)("Name").ThirdLevelProperties(i), Nothing)) With {.Tag = "ITEM", .ImageKey = imgKey, .SelectedImageKey = imgKey}
                gName = mScript.PrepareStringToPrint(prop(parentId)("Group").ThirdLevelProperties(i), Nothing)
            Else
                n = New TreeNode(mScript.PrepareStringToPrint(prop(i)("Name").Value, Nothing)) With {.Tag = "ITEM", .ImageKey = imgKey, .SelectedImageKey = imgKey}
                If prop(i)("Name").Hidden Then n.ForeColor = DEFAULT_COLORS.NodeHiddenForeColor
                gName = mScript.PrepareStringToPrint(prop(i)("Group").Value, Nothing)
            End If

            Dim gId As Integer = cGroups.GetGroupIdByName(className, gName, thirdLevel, parentName)
            If gName.Length = 0 OrElse gId = -1 Then
                tree.Nodes.Add(n) 'добавляем узел в корень 
            Else
                If gId = -1 Then
                    MessageBox.Show("Ошибка при заполнении дерева узлами. Группа не найдена.")
                    Continue For
                End If
                If (showHidden = False AndAlso cGroups.dictGroups(className)(gId).Hidden) = False Then
                    Dim groupFound As Boolean = False
                    For j = 0 To tree.Nodes.Count - 1
                        If tree.Nodes(j).Text = gName AndAlso tree.Nodes(j).Tag = "GROUP" Then
                            tree.Nodes(j).Nodes.Add(n) 'добавляем узел к группе (созданной ранее)
                            groupFound = True
                            Exit For
                        End If
                    Next j
                    If Not groupFound Then
                        'узла с такой группой еще не существует - создаем узел группы и добавляем к нему элемент
                        If cGroups.dictGroups(className)(gId).Hidden Then
                            imgKey = "groupHidden.png"
                        Else
                            imgKey = cGroups.dictGroups(className)(gId).iconName
                        End If

                        Dim parNode As New TreeNode(gName) With {.Tag = "GROUP", .ImageKey = imgKey, .SelectedImageKey = .ImageKey}
                        If cGroups.dictGroups(className)(gId).Hidden Then parNode.ForeColor = DEFAULT_COLORS.NodeHiddenForeColor
                        parNode.Nodes.Add(n)
                        tree.Nodes.Add(parNode)
                        If IsNothing(lstToExpand) = False Then
                            If lstToExpand.Contains(parNode.Text) Then parNode.Expand()
                        Else
                            parNode.Expand()
                        End If
                    End If
                End If
            End If
        Next
        'вставляем группы, в которых нет ни одного элемента. Стараемся их вставить недалеко от старого места (точное расположение выяснить невозможно без создания дополнительных регистров)        
        If IsNothing(lstGroups) = False AndAlso lstGroups.Count > 0 Then
            Dim prevGroupNode As TreeNode = Nothing 'последний найденный узел группы, уже вставленный в дерево
            For i As Integer = 0 To lstGroups.Count - 1
                If lstGroups(i).isThirdLevelGroup <> thirdLevel OrElse lstGroups(i).parentName <> parentName Then Continue For
                Dim n As TreeNode = FindRootNodeByText(lstGroups(i).Name, tree) 'если узел даной группы уже существует (он не пустой), то получаем его
                If IsNothing(n) Then
                    If chkShowHidden.Checked = False AndAlso lstGroups(i).Hidden Then Continue For 'скрытый узел пропускаем
                    'узел не существовал. Вставляем его либо в начало, либо перед предыдущим узлом группы
                    Dim imgKey As String
                    If lstGroups(i).Hidden Then
                        imgKey = "groupHidden.png"
                    Else
                        imgKey = lstGroups(i).iconName
                    End If
                    Dim newNode As New TreeNode With {.Text = lstGroups(i).Name, .Tag = "GROUP", .ImageKey = imgKey, .SelectedImageKey = .ImageKey}
                    If IsNothing(prevGroupNode) Then
                        tree.Nodes.Insert(0, newNode)
                    Else
                        tree.Nodes.Insert(prevGroupNode.Index + 1, newNode)
                    End If
                    prevGroupNode = newNode
                Else
                    'узел существовал
                    prevGroupNode = n
                End If
            Next
        End If
        tree.EndUpdate()
        If IsNothing(cPanelManager) = False Then
            If isMainTree Then cPanelManager.ReassociateNodes(classId, tree, parentName)
            If Not isMainTree Then
                Dim btnRemoveSub() As Control = tree.Parent.Controls.Find("RemoveSubElement", True)
                If IsNothing(btnRemoveSub) = False AndAlso btnRemoveSub.Count > 0 Then btnRemoveSub(0).Enabled = False
            End If
        End If
        'tree.BackColor = Color.Green

        Return 0
    End Function

    ''' <summary>
    ''' Заполняет дерево переменных группами и элементами
    ''' </summary>
    ''' <returns>0  при успехе, -1 при ошибке</returns>
    Public Function FillTreeVariables(ByVal showHidden As Boolean, Optional lstToExpand As List(Of String) = Nothing) As Integer
        Dim tree As TreeView = treeVariables
        Dim className As String = "Variable"
        If IsNothing(tree) Then Return -1
        tree.BeginUpdate()
        tree.Nodes.Clear()
        'заполняем treeView переменными
        Dim lstVariables As SortedList(Of String, cVariable.variableEditorInfoType) = mScript.csPublicVariables.lstVariables
        Dim lstGroups As List(Of clsGroups.clsGroupInfo) = cGroups.dictGroups(className)
        Log.PrintToLog("FillTree: " + className)

        If IsNothing(lstVariables) OrElse lstVariables.Count = 0 Then
            tree.EndUpdate()
            Return 0
        End If

        Dim cnt As Integer = lstVariables.Count
        Dim prevGid As Integer = -1
        Dim prevGnodeIndex As Integer = 0
        For i As Integer = 0 To cnt - 1
            Dim imgKey As String
            Dim varName As String = lstVariables.ElementAt(i).Key
            Dim varValue As cVariable.variableEditorInfoType = lstVariables.ElementAt(i).Value

            If showHidden = False AndAlso varValue.Hidden Then Continue For
            If varValue.Hidden Then
                imgKey = "itemHidden.png"
            Else
                imgKey = varValue.Icon
                If String.IsNullOrEmpty(imgKey) Then imgKey = GetDefaultIcon(className, False)
            End If
            Dim n As TreeNode
            Dim gName As String
            n = New TreeNode(varName) With {.Tag = "ITEM", .ImageKey = imgKey, .SelectedImageKey = imgKey}
            gName = varValue.Group
            If IsNothing(gName) Then gName = ""

            Dim gId As Integer = cGroups.GetGroupIdByName(className, gName, False)
            If gName.Length = 0 OrElse gId = -1 Then
                tree.Nodes.Add(n) 'добавляем узел в корень 
            Else
                If gId = -1 Then
                    MessageBox.Show("Ошибка при заполнении дерева узлами. Группа не найдена.")
                    Continue For
                End If
                If (showHidden = False AndAlso cGroups.dictGroups(className)(gId).Hidden) = False Then
                    Dim groupFound As Boolean = False
                    For j = 0 To tree.Nodes.Count - 1
                        If tree.Nodes(j).Text = gName AndAlso tree.Nodes(j).Tag = "GROUP" Then
                            tree.Nodes(j).Nodes.Add(n) 'добавляем узел к группе (созданной ранее)
                            groupFound = True
                            Exit For
                        End If
                    Next j
                    If Not groupFound Then
                        'узла с такой группой еще не существует - создаем узел группы и добавляем к нему элемент
                        If cGroups.dictGroups(className)(gId).Hidden Then
                            imgKey = "groupHidden.png"
                        Else
                            imgKey = cGroups.dictGroups(className)(gId).iconName
                        End If

                        Dim parNode As New TreeNode(gName) With {.Tag = "GROUP", .ImageKey = imgKey, .SelectedImageKey = .ImageKey}
                        If cGroups.dictGroups(className)(gId).Hidden Then parNode.ForeColor = DEFAULT_COLORS.NodeHiddenForeColor
                        parNode.Nodes.Add(n)

                        If gId < prevGid Then
                            tree.Nodes.Insert(prevGnodeIndex, parNode)
                        Else
                            tree.Nodes.Add(parNode)
                        End If
                        prevGid = gId
                        prevGnodeIndex = parNode.Index

                        If IsNothing(lstToExpand) = False Then
                            If lstToExpand.Contains(parNode.Text) Then parNode.Expand()
                        Else
                            parNode.Expand()
                        End If
                    End If
                End If
            End If
        Next
        'вставляем группы, в которых нет ни одного элемента. Стараемся их вставить недалеко от старого места (точное расположение выяснить невозможно без создания дополнительных регистров)        
        If IsNothing(lstGroups) = False AndAlso lstGroups.Count > 0 Then
            Dim prevGroupNode As TreeNode = Nothing 'последний найденный узел группы, уже вставленный в дерево
            For i As Integer = 0 To lstGroups.Count - 1
                Dim n As TreeNode = FindRootNodeByText(lstGroups(i).Name, tree) 'если узел даной группы уже существует (он не пустой), то получаем его
                If IsNothing(n) Then
                    If chkShowHidden.Checked = False AndAlso lstGroups(i).Hidden Then Continue For 'скрытый узел пропускаем
                    'узел не существовал. Вставляем его либо в начало, либо перед предыдущим узлом группы
                    Dim imgKey As String
                    If lstGroups(i).Hidden Then
                        imgKey = "groupHidden.png"
                    Else
                        imgKey = lstGroups(i).iconName
                    End If
                    Dim newNode As New TreeNode With {.Text = lstGroups(i).Name, .Tag = "GROUP", .ImageKey = imgKey, .SelectedImageKey = .ImageKey}
                    If IsNothing(prevGroupNode) Then
                        tree.Nodes.Insert(0, newNode)
                    Else
                        tree.Nodes.Insert(prevGroupNode.Index + 1, newNode)
                    End If
                    prevGroupNode = newNode
                Else
                    'узел существовал
                    prevGroupNode = n
                End If
            Next
        End If
        tree.EndUpdate()
        If IsNothing(cPanelManager) = False Then
            cPanelManager.ReassociateVariablesNodes()
        End If

        Return 0
    End Function

    ''' <summary>
    ''' Заполняет дерево функций группами и элементами
    ''' </summary>
    ''' <returns>0  при успехе, -1 при ошибке</returns>
    Public Function FillTreeFunctions(ByVal showHidden As Boolean, Optional lstToExpand As List(Of String) = Nothing) As Integer
        Dim tree As TreeView = treeFunctions
        Dim className As String = "Function"
        If IsNothing(tree) Then Return -1
        tree.BeginUpdate()
        tree.Nodes.Clear()
        'заполняем treeView функциями
        Dim lstFunctions As SortedList(Of String, MatewScript.FunctionInfoType) = mScript.functionsHash

        Dim lstGroups As List(Of clsGroups.clsGroupInfo) = cGroups.dictGroups(className)
        Log.PrintToLog("FillTree: " + className)

        If IsNothing(lstFunctions) OrElse lstFunctions.Count = 0 Then
            tree.EndUpdate()
            Return 0
        End If

        Dim cnt As Integer = lstFunctions.Count
        Dim prevGid As Integer = -1
        For i As Integer = 0 To cnt - 1
            Dim imgKey As String
            Dim varName As String = lstFunctions.ElementAt(i).Key
            Dim varValue As MatewScript.FunctionInfoType = lstFunctions.ElementAt(i).Value

            If showHidden = False AndAlso varValue.Hidden Then Continue For
            If varValue.Hidden Then
                imgKey = "itemHidden.png"
            Else
                imgKey = varValue.Icon
                If String.IsNullOrEmpty(imgKey) Then imgKey = GetDefaultIcon(className, False)
            End If
            Dim n As TreeNode
            Dim gName As String
            n = New TreeNode(varName) With {.Tag = "ITEM", .ImageKey = imgKey, .SelectedImageKey = imgKey}
            gName = varValue.Group
            If IsNothing(gName) Then gName = ""

            Dim gId As Integer = cGroups.GetGroupIdByName(className, gName, False)
            Dim prevGnodeIndex As Integer = 0
            If gName.Length = 0 OrElse gId = -1 Then
                tree.Nodes.Add(n) 'добавляем узел в корень 
            Else
                If gId = -1 Then
                    MessageBox.Show("Ошибка при заполнении дерева узлами. Группа не найдена.")
                    Continue For
                End If
                If (showHidden = False AndAlso cGroups.dictGroups(className)(gId).Hidden) = False Then
                    Dim groupFound As Boolean = False
                    For j = 0 To tree.Nodes.Count - 1
                        If tree.Nodes(j).Text = gName AndAlso tree.Nodes(j).Tag = "GROUP" Then
                            tree.Nodes(j).Nodes.Add(n) 'добавляем узел к группе (созданной ранее)
                            groupFound = True
                            Exit For
                        End If
                    Next j
                    If Not groupFound Then
                        'узла с такой группой еще не существует - создаем узел группы и добавляем к нему элемент
                        If cGroups.dictGroups(className)(gId).Hidden Then
                            imgKey = "groupHidden.png"
                        Else
                            imgKey = cGroups.dictGroups(className)(gId).iconName
                        End If

                        Dim parNode As New TreeNode(gName) With {.Tag = "GROUP", .ImageKey = imgKey, .SelectedImageKey = .ImageKey}
                        If cGroups.dictGroups(className)(gId).Hidden Then parNode.ForeColor = DEFAULT_COLORS.NodeHiddenForeColor
                        parNode.Nodes.Add(n)

                        If gId < prevGid Then
                            tree.Nodes.Insert(prevGnodeIndex, parNode)
                        Else
                            tree.Nodes.Add(parNode)
                        End If
                        prevGid = gId
                        prevGnodeIndex = parNode.Index

                        If IsNothing(lstToExpand) = False Then
                            If lstToExpand.Contains(parNode.Text) Then parNode.Expand()
                        Else
                            parNode.Expand()
                        End If
                    End If
                End If
            End If
        Next
        'вставляем группы, в которых нет ни одного элемента. Стараемся их вставить недалеко от старого места (точное расположение выяснить невозможно без создания дополнительных регистров)        
        If IsNothing(lstGroups) = False AndAlso lstGroups.Count > 0 Then
            Dim prevGroupNode As TreeNode = Nothing 'последний найденный узел группы, уже вставленный в дерево
            For i As Integer = 0 To lstGroups.Count - 1
                Dim n As TreeNode = FindRootNodeByText(lstGroups(i).Name, tree) 'если узел даной группы уже существует (он не пустой), то получаем его
                If IsNothing(n) Then
                    If chkShowHidden.Checked = False AndAlso lstGroups(i).Hidden Then Continue For 'скрытый узел пропускаем
                    'узел не существовал. Вставляем его либо в начало, либо перед предыдущим узлом группы
                    Dim imgKey As String
                    If lstGroups(i).Hidden Then
                        imgKey = "groupHidden.png"
                    Else
                        imgKey = lstGroups(i).iconName
                    End If
                    Dim newNode As New TreeNode With {.Text = lstGroups(i).Name, .Tag = "GROUP", .ImageKey = imgKey, .SelectedImageKey = .ImageKey}
                    If IsNothing(prevGroupNode) Then
                        tree.Nodes.Insert(0, newNode)
                    Else
                        tree.Nodes.Insert(prevGroupNode.Index + 1, newNode)
                    End If
                    prevGroupNode = newNode
                Else
                    'узел существовал
                    prevGroupNode = n
                End If
            Next
        End If
        tree.EndUpdate()
        If IsNothing(cPanelManager) = False Then
            cPanelManager.ReassociateFunctionsNodes()
        End If

        Return 0
    End Function

    ''' <summary>
    ''' Создает список для функции FillTree
    ''' </summary>
    ''' <param name="tree"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MakeListToExpand(ByRef tree As TreeView) As List(Of String)
        Dim lst As New List(Of String)
        For i As Integer = 0 To tree.Nodes.Count - 1
            If tree.Nodes(i).IsExpanded Then lst.Add(tree.Nodes(i).Text)
        Next
        Return lst
    End Function

    ''' <summary>
    ''' Находит узел дерева 1 уровня по его названию
    ''' </summary>
    ''' <param name="nodeText">Текст узла группы для поиска</param>
    ''' <param name="tree">Дерево</param>
    ''' <param name="seekGroups">Искать в группах или в элементах</param>
    Public Function FindRootNodeByText(ByVal nodeText As String, ByRef tree As TreeView, Optional ByVal seekGroups As Boolean = True) As TreeNode
        For i As Integer = 0 To tree.Nodes.Count - 1
            Dim n As TreeNode = tree.Nodes(i)
            If seekGroups Then
                If n.Tag = "GROUP" AndAlso n.Text = nodeText Then Return n
            Else
                If n.Tag = "ITEM" AndAlso n.Text = nodeText Then Return n

            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Находит узел дерева по его названию (только узел элемента)
    ''' </summary>
    ''' <param name="nodeText">Текст узла группы для поиска</param>
    ''' <param name="tree">Дерево</param>
    Public Function FindItemNodeByText(ByRef tree As TreeView, ByVal nodeText As String) As TreeNode
        For i As Integer = 0 To tree.Nodes.Count - 1
            If tree.Nodes(i).Tag = "GROUP" Then
                For j As Integer = 0 To tree.Nodes(i).Nodes.Count - 1
                    If tree.Nodes(i).Nodes(j).Text = nodeText Then Return tree.Nodes(i).Nodes(j)
                Next
            Else
                If tree.Nodes(i).Text = nodeText Then Return tree.Nodes(i)
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>Загружает все TreeView данными исследования (локации, объекты, меню ...)</summary>
    Public Sub LoadTrees()
        LoadIconsForTrees() 'загружаем в начале иконки для TreeViev
        'Создаем сплит-панель. Сверху - кнопки управления (добавить группу, удалить ...), снизу - TreeView
        splitTreeContainer = New SplitContainer With {.Orientation = Orientation.Horizontal, .Dock = DockStyle.Fill}
        SplitOuter.Panel1.Controls.Add(splitTreeContainer)
        splitTreeContainer.Panel1.BackgroundImage = My.Resources.bg01

        'Создание панели управления
        'Надпись с помощью
        Dim lblClassHelp As New Label With {.Text = "     " + "Локации" + " → ", .AutoSize = True, .Left = 5, .Top = 5, .BackColor = Color.Transparent, .Image = My.Resources.help, .Cursor = Cursors.Help, _
                                   .ImageAlign = ContentAlignment.MiddleLeft, .TextAlign = ContentAlignment.MiddleRight, .ForeColor = Color.IndianRed, _
                                   .Font = New Font(Me.Font.FontFamily, 20, FontStyle.Italic), .Tag = Application.StartupPath + "\Help\Location\about.html"}
        lblClassHelp.Name = "lblClassHelp"
        splitTreeContainer.Panel1.Controls.Add(lblClassHelp)
        AddHandler lblClassHelp.Click, Sub(sender As Object, e As EventArgs)
                                           WBhelp.Navigate(sender.Tag)
                                           WBhelp.Visible = True
                                       End Sub
        'Чекбокс показать скрытые
        chkShowHidden = New CheckBox With {.Checked = True, .Left = 5, .Top = 59, .AutoSize = True, .Text = "Показать скрытые"}
        splitTreeContainer.Panel1.Controls.Add(chkShowHidden)
        AddHandler chkShowHidden.CheckedChanged, Sub(sender As Object, e As EventArgs)
                                                     If IsNothing(currentTreeView) Then Return
                                                     If currentClassName = "Variable" Then
                                                         FillTreeVariables(chkShowHidden.Checked, MakeListToExpand(currentTreeView))
                                                     ElseIf currentClassName = "Function" Then
                                                         FillTreeFunctions(chkShowHidden.Checked, MakeListToExpand(currentTreeView))
                                                     Else
                                                         FillTree(currentClassName, currentTreeView, chkShowHidden.Checked, MakeListToExpand(currentTreeView), currentParentName)
                                                     End If
                                                     If IsNothing(cPanelManager.ActivePanel) = False Then
                                                         Dim nod As TreeNode = cPanelManager.ActivePanel.treeNode
                                                         If IsNothing(nod) = False AndAlso IsNothing(nod.TreeView) = False Then
                                                             currentTreeView.SelectedNode = nod
                                                         Else
                                                             cPanelManager.HidePanel(cPanelManager.ActivePanel)
                                                             cPanelManager.ActivePanel = Nothing
                                                         End If
                                                     End If
                                                 End Sub
        'кнопка сделать скрытым
        btnMakeHidden = New Button With {.Text = "Скрыть", .Width = 150, .Height = 39, .Left = 177, .Top = 50, .Image = My.Resources.hidden, .ImageAlign = ContentAlignment.MiddleLeft, _
                                               .TextAlign = ContentAlignment.MiddleRight}
        splitTreeContainer.Panel1.Controls.Add(btnMakeHidden)
        AddHandler btnMakeHidden.Click, AddressOf btnMakeHidden_Click

        'кнопка Копировать в...
        btnCopyTo = New Button With {.Text = "Копировать в...", .Width = btnMakeHidden.Right - questEnvironment.defPaddingLeft, .Height = 39, _
                                     .Left = questEnvironment.defPaddingLeft, .Top = 50, .Image = My.Resources.CopyTo, .ImageAlign = ContentAlignment.MiddleLeft, _
                                               .TextAlign = ContentAlignment.MiddleRight, .Visible = False}
        splitTreeContainer.Panel1.Controls.Add(btnCopyTo)
        AddHandler btnCopyTo.Click, AddressOf btnCopyTo_Click

        'кнопка добавить группу
        btnAddGroup = New Button With {.Text = "Новая группа", .Width = 101, .Height = 56, .Left = chkShowHidden.Left, .Top = 95, .Image = My.Resources.add_group32, .ImageAlign = ContentAlignment.MiddleLeft, _
                                              .TextAlign = ContentAlignment.MiddleRight}
        splitTreeContainer.Panel1.Controls.Add(btnAddGroup)
        AddHandler btnAddGroup.Click, Sub(sender As Object, e As EventArgs)
                                          Call tsmiAddGroup_Click(tsmiAddGroup, New EventArgs)
                                      End Sub
        'кнопка добавить элемент
        btnAddElement = New Button With {.Text = "New element", .Width = 216, .Height = btnAddGroup.Height, .Left = 112, .Top = btnAddGroup.Top, .Image = My.Resources.add32, _
                                                        .ImageAlign = ContentAlignment.MiddleLeft, .TextAlign = ContentAlignment.MiddleRight}
        splitTreeContainer.Panel1.Controls.Add(btnAddElement)
        AddHandler btnAddElement.Click, Sub(sender As Object, e As EventArgs)
                                            Call tsmiAddElement_Click(tsmiAddElement, New EventArgs)
                                        End Sub
        'кнопка удаления (и группы, и элемента)
        btnDelete = New Button With {.Text = "", .Width = 48, .Height = btnAddElement.Height, .Left = chkShowHidden.Left, .Top = 157, .Image = My.Resources.delete32, .Enabled = False}
        splitTreeContainer.Panel1.Controls.Add(btnDelete)
        AddHandler btnDelete.Click, Sub(sender As Object, e As EventArgs)
                                        Call tsmiRemoveElement_Click(tsmiRemoveElement, New EventArgs)
                                    End Sub
        'кнопка создания дубликата
        btnDuplicate = New Button With {.Text = "", .Size = btnDelete.Size, .Left = 58, .Top = btnDelete.Top, .Image = My.Resources.duplicate32, .Enabled = False}
        splitTreeContainer.Panel1.Controls.Add(btnDuplicate)
        AddHandler btnDuplicate.Click, Sub(sender As Object, e As EventArgs)
                                           Call tsmiDuplicate_Click(tsmiDuplicate, New EventArgs)
                                       End Sub
        'кнопка выбора настроек по умолчанию
        btnShowSettings = New Button With {.Text = "Установки по умолчанию", .Width = btnAddElement.Width, .Height = btnAddGroup.Height, .Left = btnAddElement.Left, .Top = btnDelete.Top, .Image = My.Resources.def, _
                                                        .ImageAlign = ContentAlignment.MiddleLeft, .TextAlign = ContentAlignment.MiddleRight}
        splitTreeContainer.Panel1.Controls.Add(btnShowSettings)
        AddHandler btnShowSettings.Click, AddressOf btnShowSettings_Click

        If splitTreeContainer.Width > 0 Then
            splitTreeContainer.SplitterDistance = btnShowSettings.Bottom + 5
            SplitOuter.SplitterDistance = btnShowSettings.Right + 5
            SplitOuter.Panel1MinSize = SplitOuter.SplitterDistance
            Dim lblHeight As Integer = lblClassHelp.Height
            lblClassHelp.AutoSize = False
            lblClassHelp.Size = New Size(SplitOuter.SplitterDistance, lblHeight)
        End If

        AddHandler SplitOuter.SplitterMoved, Sub(sender As Object, e As SplitterEventArgs)
                                                 Dim w As Integer = sender.Panel1.ClientSize.Width
                                                 lblClassHelp.Width = w - lblClassHelp.Left
                                                 btnShowSettings.Width = w - btnShowSettings.Left
                                                 btnAddElement.Width = w - btnAddElement.Left
                                                 btnMakeHidden.Width = w - btnMakeHidden.Left
                                                 btnCopyTo.Width = w - questEnvironment.defPaddingLeft
                                             End Sub


        'Dim tree As TreeView
        For classId As Integer = 0 To mScript.mainClass.Count - 1
            'If mScript.mainClass(classId).LevelsCount = 0 Then Continue For

            'tree = New TreeView With {.AllowDrop = True, .Dock = DockStyle.Fill, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, _
            '                                              .ShowPlusMinus = True, .ShowRootLines = True, .Visible = False, .ImageList = imgLstGroupIcons, .HotTracking = True}
            'dictTrees.Add(classId, tree)

            ''наполняем дерево событиями размещаем на форме
            'tree.ContextMenuStrip = cmnuMainTree
            'splitTreeContainer.Panel2.Controls.Add(tree)
            'AddHandler tree.AfterLabelEdit, AddressOf tree_AfterLabelEdit
            'AddHandler tree.AfterSelect, AddressOf tree_AfterSelect
            'AddHandler tree.DoubleClick, AddressOf tree_DoubleClick
            'AddHandler tree.DragDrop, AddressOf tree_DragDrop
            'AddHandler tree.DragOver, AddressOf tree_DragOver
            'AddHandler tree.MouseDown, AddressOf tree_MouseDown
            'AddHandler tree.MouseMove, AddressOf tree_MouseMove
            'AddHandler tree.MouseUp, AddressOf tree_MouseUp
            'AddHandler tree.QueryContinueDrag, AddressOf tree_QueryContinueDrag
            'AddHandler tree.VisibleChanged, Sub(sender As Object, e As EventArgs) tree_Visiblechanged(sender, lblClassHelp, btnAddElement)
            'AddHandler tree.BeforeLabelEdit, Sub(sender As Object, e As NodeLabelEditEventArgs) ShowIconMenu(sender)
            'AddHandler tree.BeforeSelect, Sub(sender As Object, e As TreeViewCancelEventArgs)
            '                                  If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False AndAlso _
            '                                      cPanelManager.ActivePanel.child2Id > -1 AndAlso questEnvironment.EnabledEvents Then SaveLocationActions(cPanelManager.ActivePanel.child2Id)
            '                              End Sub
            'AddHandler tree.KeyUp, AddressOf tree_KeyUp
            AppendTree(classId)
        Next classId

        'дерево переменных
        Dim tree As TreeView
        tree = New TreeView With {.AllowDrop = True, .Dock = DockStyle.Fill, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, _
                                                      .ShowPlusMinus = True, .ShowRootLines = True, .Visible = False, .ImageList = imgLstGroupIcons, .HotTracking = True}
        treeVariables = tree
        'наполняем дерево событиями размещаем на форме
        tree.ContextMenuStrip = cmnuMainTree
        splitTreeContainer.Panel2.Controls.Add(tree)
        AddHandler tree.AfterLabelEdit, AddressOf tree_AfterLabelEdit
        AddHandler tree.AfterSelect, AddressOf tree_AfterSelectVariables
        AddHandler tree.DoubleClick, AddressOf tree_DoubleClick
        AddHandler tree.DragDrop, AddressOf tree_DragDrop
        AddHandler tree.DragOver, AddressOf tree_DragOver
        AddHandler tree.MouseDown, AddressOf tree_MouseDown
        AddHandler tree.MouseMove, AddressOf tree_MouseMove
        AddHandler tree.MouseUp, AddressOf tree_MouseUp
        AddHandler tree.QueryContinueDrag, AddressOf tree_QueryContinueDrag
        AddHandler tree.VisibleChanged, Sub(sender As Object, e As EventArgs) tree_VisiblechangedVariables(sender, lblClassHelp, btnAddElement)
        AddHandler tree.BeforeLabelEdit, Sub(sender As Object, e As NodeLabelEditEventArgs) ShowIconMenu(sender)
        AddHandler tree.KeyUp, AddressOf tree_KeyUp
        'дерево функций
        tree = New TreeView With {.AllowDrop = True, .Dock = DockStyle.Fill, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, _
                                              .ShowPlusMinus = True, .ShowRootLines = True, .Visible = False, .ImageList = imgLstGroupIcons, .HotTracking = True}
        treeFunctions = tree
        'наполняем дерево событиями размещаем на форме
        tree.ContextMenuStrip = cmnuMainTree
        splitTreeContainer.Panel2.Controls.Add(tree)
        AddHandler tree.AfterLabelEdit, AddressOf tree_AfterLabelEdit
        AddHandler tree.AfterSelect, AddressOf tree_AfterSelectFunctions
        AddHandler tree.DoubleClick, AddressOf tree_DoubleClick
        AddHandler tree.DragDrop, AddressOf tree_DragDrop
        AddHandler tree.DragOver, AddressOf tree_DragOver
        AddHandler tree.MouseDown, AddressOf tree_MouseDown
        AddHandler tree.MouseMove, AddressOf tree_MouseMove
        AddHandler tree.MouseUp, AddressOf tree_MouseUp
        AddHandler tree.QueryContinueDrag, AddressOf tree_QueryContinueDrag
        AddHandler tree.VisibleChanged, Sub(sender As Object, e As EventArgs) tree_VisiblechangedFunctions(sender, lblClassHelp, btnAddElement)
        AddHandler tree.BeforeLabelEdit, Sub(sender As Object, e As NodeLabelEditEventArgs) ShowIconMenu(sender)
        AddHandler tree.KeyUp, AddressOf tree_KeyUp

        splitTreeContainer.Hide()
        CreateIconsForTreeMenus()
    End Sub

    Public Sub AppendTree(ByVal classId As Integer)
        Dim tree As TreeView
        If mScript.mainClass(classId).LevelsCount = 0 Then Return

        tree = New TreeView With {.AllowDrop = True, .Dock = DockStyle.Fill, .FullRowSelect = True, .LabelEdit = True, .Scrollable = True, .ShowLines = True, _
                                                      .ShowPlusMinus = True, .ShowRootLines = True, .Visible = False, .ImageList = imgLstGroupIcons, .HotTracking = True}
        dictTrees.Add(classId, tree)

        'наполняем дерево событиями размещаем на форме
        tree.ContextMenuStrip = cmnuMainTree
        splitTreeContainer.Panel2.Controls.Add(tree)
        AddHandler tree.AfterLabelEdit, AddressOf tree_AfterLabelEdit
        AddHandler tree.AfterSelect, AddressOf tree_AfterSelect
        AddHandler tree.DoubleClick, AddressOf tree_DoubleClick
        AddHandler tree.DragDrop, AddressOf tree_DragDrop
        AddHandler tree.DragOver, AddressOf tree_DragOver
        AddHandler tree.MouseDown, AddressOf tree_MouseDown
        AddHandler tree.MouseMove, AddressOf tree_MouseMove
        AddHandler tree.MouseUp, AddressOf tree_MouseUp
        AddHandler tree.QueryContinueDrag, AddressOf tree_QueryContinueDrag
        AddHandler tree.VisibleChanged, Sub(sender As Object, e As EventArgs) tree_Visiblechanged(sender, splitTreeContainer.Panel1.Controls("lblClassHelp"), btnAddElement)
        AddHandler tree.BeforeLabelEdit, Sub(sender As Object, e As NodeLabelEditEventArgs) ShowIconMenu(sender)
        'AddHandler tree.BeforeSelect, Sub(sender As Object, e As TreeViewCancelEventArgs)
        '                                  If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False AndAlso _
        '                                      cPanelManager.ActivePanel.child2Id > -1 AndAlso questEnvironment.EnabledEvents Then actionsRouter.SaveActions(cPanelManager.ActivePanel.child2Id)
        '                              End Sub
        AddHandler tree.KeyUp, AddressOf tree_KeyUp
    End Sub

#Region "treeVariables & treeFunctions delegates"
    Private Sub tree_VisiblechangedVariables(sender As TreeView, lblClassHelp As Label, btnAddElement As Button)
        If Not sender.Visible Then Return
        Dim curClassId As Integer = -2

        lblClassHelp.Text = "     " + GetTranslatedName(curClassId)
        lblClassHelp.Tag = Application.StartupPath + "\Help\Var\about.html"
        btnAddElement.Text = GetTranslationAddNew(curClassId, currentParentName.Length > 0)
    End Sub

    Private Sub tree_VisiblechangedFunctions(sender As TreeView, lblClassHelp As Label, btnAddElement As Button)
        If Not sender.Visible Then Return
        Dim curClassId As Integer = -3

        lblClassHelp.Text = "     " + GetTranslatedName(curClassId)
        lblClassHelp.Tag = Application.StartupPath + "\Help\Function\about.html"
        btnAddElement.Text = GetTranslationAddNew(curClassId, currentParentName.Length > 0)
    End Sub

    Public Sub tree_AfterSelectVariables(sender As Object, e As TreeViewEventArgs)
        primaryNameSelection = False
        Dim n As TreeNode = e.Node
        Log.PrintToLog("treeAfterSelectVariables (start)")
        Dim prevNode As TreeNode = prevNodeMainTree
        If IsNothing(prevNode) = False Then
            'возвращаем цвет по-умолчанию предыдущему узлу
            If prevNode.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then prevNode.ForeColor = DEFAULT_COLORS.NodeForeColor
            If prevNode.BackColor <> DEFAULT_COLORS.ControlTransparentBackground Then prevNode.BackColor = DEFAULT_COLORS.ControlTransparentBackground
        End If
        If btnShowSettings.BackColor = DEFAULT_COLORS.NodeSelBackColor Then btnShowSettings.BackColor = btnAddElement.BackColor
        'выделяем цветом текущий выбранный узел
        If n.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then n.ForeColor = DEFAULT_COLORS.NodeSelForeColor
        If n.BackColor <> DEFAULT_COLORS.NodeSelBackColor Then n.BackColor = DEFAULT_COLORS.NodeSelBackColor
        prevNode = n 'сохраняем текущий узел, что бы затем можно было вернуть ему цвет по-умолчанию
        prevNodeMainTree = prevNode

        'Настройка вида панели управления элементами
        If IsNothing(sender.SelectedNode) Then
            btnDelete.Enabled = False
            tsmiRemoveElement.Enabled = False
            btnDuplicate.Enabled = False
            tsmiDuplicate.Enabled = False
            btnMakeHidden.Enabled = False
            Return
        End If

        btnMakeHidden.Enabled = True
        If n.ForeColor = DEFAULT_COLORS.NodeHiddenForeColor Then
            btnMakeHidden.Text = "Отобразить"
        Else
            btnMakeHidden.Text = "Скрыть"
        End If

        If n.Tag = "GROUP" AndAlso n.Nodes.Count > 0 Then
            btnDelete.Enabled = False
            tsmiRemoveElement.Enabled = False
        Else
            btnDelete.Enabled = True
            tsmiRemoveElement.Enabled = True
        End If

        If n.Tag = "GROUP" Then
            btnDuplicate.Enabled = False
            tsmiDuplicate.Enabled = False
        Else
            tsmiDuplicate.Enabled = True
            btnDuplicate.Enabled = True
        End If

        If n.Tag = "GROUP" Then Return
        'Открытие нового окна
        If questEnvironment.EnabledEvents = False Then Return
        Log.PrintToLog("treeAfterSelectVariables: " + n.Text)

        cPanelManager.OpenPanel(currentTreeView.SelectedNode, -2, n.Text, "")
    End Sub

    Public Sub tree_AfterSelectFunctions(sender As Object, e As TreeViewEventArgs)
        primaryNameSelection = False
        Dim n As TreeNode = e.Node
        Log.PrintToLog("treeAfterSelectFunctions (start)")
        Dim prevNode As TreeNode = prevNodeMainTree
        If IsNothing(prevNode) = False Then
            'возвращаем цвет по-умолчанию предыдущему узлу
            If prevNode.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then prevNode.ForeColor = DEFAULT_COLORS.NodeForeColor
            If prevNode.BackColor <> DEFAULT_COLORS.ControlTransparentBackground Then prevNode.BackColor = DEFAULT_COLORS.ControlTransparentBackground
        End If
        If btnShowSettings.BackColor = DEFAULT_COLORS.NodeSelBackColor Then btnShowSettings.BackColor = btnAddElement.BackColor
        'выделяем цветом текущий выбранный узел
        If n.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then n.ForeColor = DEFAULT_COLORS.NodeSelForeColor
        If n.BackColor <> DEFAULT_COLORS.NodeSelBackColor Then n.BackColor = DEFAULT_COLORS.NodeSelBackColor
        prevNode = n 'сохраняем текущий узел, что бы затем можно было вернуть ему цвет по-умолчанию
        prevNodeMainTree = prevNode

        'Настройка вида панели управления элементами
        If IsNothing(sender.SelectedNode) Then
            btnDelete.Enabled = False
            tsmiRemoveElement.Enabled = False
            btnDuplicate.Enabled = False
            tsmiDuplicate.Enabled = False
            btnMakeHidden.Enabled = False
            Return
        End If

        btnMakeHidden.Enabled = True
        If n.ForeColor = DEFAULT_COLORS.NodeHiddenForeColor Then
            btnMakeHidden.Text = "Отобразить"
        Else
            btnMakeHidden.Text = "Скрыть"
        End If

        If n.Tag = "GROUP" AndAlso n.Nodes.Count > 0 Then
            btnDelete.Enabled = False
            tsmiRemoveElement.Enabled = False
        Else
            btnDelete.Enabled = True
            tsmiRemoveElement.Enabled = True
        End If

        If n.Tag = "GROUP" Then
            btnDuplicate.Enabled = False
            tsmiDuplicate.Enabled = False
        Else
            tsmiDuplicate.Enabled = True
            btnDuplicate.Enabled = True
        End If

        If n.Tag = "GROUP" Then Return
        'Открытие нового окна
        If questEnvironment.EnabledEvents = False Then Return
        Log.PrintToLog("treeAfterSelectFunctions: " + n.Text)

        cPanelManager.OpenPanel(currentTreeView.SelectedNode, -3, n.Text, "")
    End Sub

    Public Sub TreeAfterLabelEditVariables(ByRef e As NodeLabelEditEventArgs)
        Dim className As String = "Variable"

        If iconMenuElements.Visible Then iconMenuElements.Hide()
        If iconMenuGroups.Visible Then iconMenuGroups.Hide()

        If String.IsNullOrWhiteSpace(e.Label) Then Return

        If e.Node.Tag = "GROUP" Then
            'Изменяем название группы
            If cGroups.IsGroupExist(className, e.Label, False) AndAlso (e.Node.Text.ToLower <> e.Label.ToLower) Then
                MessageBox.Show("Группа с таким названием уже существует!")
                e.CancelEdit = True
                Return
            End If
            If cGroups.ChangeGroupName(className, e.Node.Text, e.Label, False) = -1 Then e.CancelEdit = True
            Return
        ElseIf e.Node.Tag = "ITEM" Then
            Dim newName As String = e.Label
            Dim oldName As String = e.Node.Text

            If cPanelManager.ActivePanel.child2Name.Length = 0 Then
                'работа с элементами 2 уровня
                MessageBox.Show("Ошибка при переименовании. Не найдена переменная!")
                e.CancelEdit = True
                Return
            End If
            If IsNothing(mScript.csPublicVariables.lstVariables) = False Then
                If mScript.csPublicVariables.lstVariables.ContainsKey(newName) Then
                    MessageBox.Show("Переменная с таким названием уже существует!")
                    e.CancelEdit = True
                    Return
                End If
            End If
            'меняем имя в csPublicVariables
            Dim v As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables(oldName)
            mScript.csPublicVariables.lstVariables.Remove(oldName)
            mScript.csPublicVariables.lstVariables.Add(newName, v)

            Dim ch As clsPanelManager.clsChildPanel
            ch = cPanelManager.GetPanelByNode(e.Node)
            'Изменяем надпись на вкладке
            If IsNothing(ch) = False Then
                cPanelManager.ChangePanelChildsName(ch, oldName, newName)
            End If
            GlobalSeeker.ReplaceElementNameInStruct(-2, oldName, newName, CodeTextBox.EditWordTypeEnum.W_VARIABLE)
            'cPanelManager.ReassociateNodes(-2, treeVariables, -1)
            'перестройка дерева через таймер - иначе глюки
            Dim tim As New Timer With {.Interval = 1}
            AddHandler tim.Tick, Sub(sender2 As Object, e2 As EventArgs)
                                     tim.Stop()
                                     FillTreeVariables(chkShowHidden.Checked, MakeListToExpand(treeVariables))
                                     treeVariables.SelectedNode = FindItemNodeByText(treeVariables, newName)
                                     tim.Dispose()
                                 End Sub
            tim.Start()
        End If
    End Sub

    Public Sub TreeAfterLabelEditFunctions(ByRef e As NodeLabelEditEventArgs)
        Dim className As String = "Function"

        If iconMenuElements.Visible Then iconMenuElements.Hide()
        If iconMenuGroups.Visible Then iconMenuGroups.Hide()

        If String.IsNullOrWhiteSpace(e.Label) Then Return

        If e.Node.Tag = "GROUP" Then
            'Изменяем название группы
            If cGroups.IsGroupExist(className, e.Label, False) AndAlso (e.Node.Text.ToLower <> e.Label.ToLower) Then
                MessageBox.Show("Группа с таким названием уже существует!")
                e.CancelEdit = True
                Return
            End If
            If cGroups.ChangeGroupName(className, e.Node.Text, e.Label, False) = -1 Then e.CancelEdit = True
            Return
        ElseIf e.Node.Tag = "ITEM" Then
            Dim newName As String = e.Label
            Dim oldName As String = e.Node.Text

            If cPanelManager.ActivePanel.child2Name.Length = 0 Then
                'работа с элементами 2 уровня
                MessageBox.Show("Ошибка при переименовании. Не найдена функция!")
                e.CancelEdit = True
                Return
            End If
            If mScript.functionsHash.ContainsKey(newName) Then
                MessageBox.Show("Функция с таким названием уже существует!")
                e.CancelEdit = True
                Return
            End If
            'меняем имя в functionsHash
            Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(oldName)
            mScript.functionsHash.Remove(oldName)
            mScript.functionsHash.Add(newName, f)

            Dim ch As clsPanelManager.clsChildPanel
            ch = cPanelManager.GetPanelByNode(e.Node)
            'Изменяем надпись на вкладке
            If IsNothing(ch) = False Then
                cPanelManager.ChangePanelChildsName(ch, oldName, newName)
            End If
            GlobalSeeker.RenameFunctionsInStruct(WrapString(oldName), WrapString(newName))
            'cPanelManager.ReassociateNodes(-3, treeFunctions, -1)
            'перестройка дерева через таймер - иначе глюки
            Dim tim As New Timer With {.Interval = 1}
            AddHandler tim.Tick, Sub(sender2 As Object, e2 As EventArgs)
                                     tim.Stop()
                                     FillTreeFunctions(chkShowHidden.Checked, MakeListToExpand(treeFunctions))
                                     treeFunctions.SelectedNode = FindItemNodeByText(treeFunctions, newName)
                                     tim.Dispose()
                                 End Sub
            tim.Start()
        End If
    End Sub
#End Region

    Private Sub tree_Visiblechanged(sender As TreeView, lblClassHelp As Label, btnAddElement As Button)
        If Not sender.Visible Then Return
        Dim curClassId As Integer = mScript.mainClassHash(currentClassName)
        If currentParentName.Length = 0 Then
            lblClassHelp.Text = "     " + GetTranslatedName(curClassId) + " → "
        ElseIf currentParentName.Length > 0 Then
            Dim parClassId As Integer = GetParentClassId(currentClassName)
            If parClassId >= 0 Then
                lblClassHelp.Text = mScript.PrepareStringToPrint(currentParentName, _
                                                                 Nothing, False) + " → " + GetTranslatedName(curClassId) + " → "
            ElseIf mScript.mainClass(curClassId).LevelsCount = 2 Then
                lblClassHelp.Text = mScript.PrepareStringToPrint(currentParentName, _
                                                                 Nothing, False) + " → "
            End If
        Else
            lblClassHelp.Text = "     " + GetTranslatedName(curClassId, True) + " → "
        End If
        'Log.PrintToLog("tree.VisibleChanged. Caption: " + lblClassHelp.Text)
        lblClassHelp.Tag = GetHelpPath(mScript.mainClass(curClassId).HelpFile) ' Application.StartupPath + "\Help\" + mScript.mainClass(curClassId).Names.Last + "\about.html"
        btnAddElement.Text = GetTranslationAddNew(curClassId, currentParentName.Length > 0)
    End Sub

    Private Sub tree_AfterLabelEdit(sender As Object, e As NodeLabelEditEventArgs)
        If currentClassName = "Function" Then
            TreeAfterLabelEditFunctions(e)
        ElseIf currentClassName = "Variable" Then
            TreeAfterLabelEditVariables(e)
        Else
            TreeAfterLabelEdit(sender, e, currentClassName)
        End If
    End Sub

    Public Sub TreeAfterLabelEdit(sender As TreeView, ByRef e As NodeLabelEditEventArgs, ByVal className As String)
        If iconMenuElements.Visible Then iconMenuElements.Hide()
        If iconMenuGroups.Visible Then iconMenuGroups.Hide()

        If String.IsNullOrWhiteSpace(e.Label) Then Return
        Dim ni As NodeInfo = GetNodeInfo(e.Node)
        Dim child2Id As Integer = ni.GetChild2Id
        Dim child3Id As Integer = ni.GetChild3Id(child2Id)

        If e.Node.Tag = "GROUP" Then
            'Изменяем название группы
            Dim parentName As String = GetParentName(sender)
            If cGroups.IsGroupExist(className, e.Label, ni.ThirdLevelNode, parentName) AndAlso (e.Node.Text.ToLower <> e.Label.ToLower) Then
                MessageBox.Show("Группа с таким названием уже существует!")
                e.CancelEdit = True
                Return
            End If
            If cGroups.ChangeGroupName(className, e.Node.Text, e.Label, ni.ThirdLevelNode, parentName) = -1 Then e.CancelEdit = True
        ElseIf e.Node.Tag = "ITEM" Then
            Dim newName As String = e.Label
            If ni.nodeChild2Name.Length = 0 Then
                'работа с элементами 2 уровня
                MessageBox.Show("Ошибка при переименовании. Не найден элемент второго порядка!")
                e.CancelEdit = True
                Return
            End If
            If ni.ThirdLevelNode AndAlso ni.nodeChild3Name.Length = 0 Then
                'работа с элементами 3 уровня
                MessageBox.Show("Ошибка при переименовании. Не найден элемент третьего порядка!")
                e.CancelEdit = True
                Return
            End If

            Dim funcParams() As String
            If ni.nodeChild3Name.Length = 0 Then
                ReDim funcParams(0)
                funcParams(0) = "'" + newName + "'"
            Else
                ReDim funcParams(1)
                funcParams(0) = child2Id.ToString
                funcParams(1) = "'" + newName + "'"
            End If

            If ObjectIsExists(ni.classId, funcParams) = "True" Then
                MessageBox.Show("Элемент с таким названием уже существует!")
                e.CancelEdit = True
                Return
            End If
            Dim wNewName As String = WrapString(newName)
            Dim wOldName As String = WrapString(e.Node.Text)
            SetPropertyValue(ni.classId, "Name", wNewName, child2Id, child3Id)

            Dim ch As clsPanelManager.clsChildPanel
            Dim isMainTree As Boolean = Object.Equals(currentTreeView, sender)

            If isMainTree Then
                'Изменение свойства Caption
                Dim propCaption As MatewScript.PropertiesInfoType = Nothing
                If mScript.mainClass(ni.classId).Properties.TryGetValue("Caption", propCaption) Then
                    'существует свойство Caption. Если его значение идентично с прежним значением имени - вставляем в соответствующий контрол (с именем Caption) новое значение
                    Dim captValue As String = ""
                    If ni.ThirdLevelNode Then
                        captValue = mScript.mainClass(ni.classId).ChildProperties(child2Id)("Caption").ThirdLevelProperties(child3Id).ToLower
                    Else
                        captValue = mScript.mainClass(ni.classId).ChildProperties(child2Id)("Caption").Value.ToLower()
                    End If
                    If captValue = WrapString(newName.ToLower) Then 'значение Caption в структуре mainClass уже успело измениться в функции SetPropertyValue
                        Dim mPanel As clsPanelManager.PanelEx = cPanelManager.dictDefContainers(ni.classId)
                        Dim arrC() As Control = mPanel.Controls.Find("Caption", True)
                        If IsNothing(arrC) = False AndAlso arrC.Count > 0 Then
                            arrC(0).Text = newName
                        End If
                    End If
                End If
                If currentClassName = "L" Then
                    'Переименовываем локацию в классе сохраненных действий
                    actionsRouter.RenameLocation(wOldName, wNewName)
                    'Переименовываем родителей в группах действий данной локации
                    cGroups.RenameGroupInChilren("A", wOldName, wNewName)
                ElseIf ni.ThirdLevelNode = False AndAlso mScript.mainClass(ni.classId).LevelsCount = 2 Then
                    'Переименовываем родителей в группах соответствующих элементов 3 уровня
                    cGroups.RenameGroupInChilren(currentClassName, wOldName, wNewName)
                End If
            End If

            'Замена в скриптах
            If primaryNameSelection = False Then
                If ni.classId = mScript.mainClassHash("A") Then
                    GlobalSeeker.RenameActionInStruct(WrapString(e.Node.Text), WrapString(newName))
                ElseIf child3Id < 0 Then
                    GlobalSeeker.RenameChild2InStruct(ni.classId, WrapString(e.Node.Text), WrapString(newName))
                Else
                    GlobalSeeker.RenameChild3InStruct(ni.classId, child2Id, WrapString(e.Node.Text), WrapString(newName))
                End If
            Else
                primaryNameSelection = False
            End If
            'cPanelManager.ReassociateNodes(ni.classId, sender, GetParentId(mScript.mainClass(ni.classId).Names(0), sender))

            If isMainTree Then
                ch = cPanelManager.GetPanelByNode(e.Node)
                'Изменяем надпись на вкладке
                If IsNothing(ch) = False Then
                    'ch.toolButton.Text = cPanelManager.MakeToolButtonText(ch)
                    cPanelManager.ChangePanelChildsName(ch, wOldName, wNewName)
                    'Изменяем надпись вида "Локация1"
                    Dim arrC() As Control = cPanelManager.dictDefContainers(ni.classId).Controls.Find("ElementName", True)
                    If IsNothing(arrC) = False AndAlso arrC.Count > 0 Then
                        arrC(0).Text = newName
                    End If
                End If
                'изменяем имя в списке
                If Not ni.ThirdLevelNode Then
                    cListManager.RenameElementInList(ni.classId, e.Node.Text, newName)
                End If
                Return
            ElseIf ni.ThirdLevelNode Then
                ch = cPanelManager.GetPanelByChildInfo(ni.classId, ni.nodeChild2Name, ni.nodeChild3Name, ni.nodeChild2Name)
            Else
                ch = cPanelManager.GetPanelByChildInfo(ni.classId, ni.nodeChild2Name, -1, cPanelManager.ActivePanel.child2Name)
            End If

            'изменяем имя в списке
            If Not ni.ThirdLevelNode Then
                cListManager.RenameElementInList(ni.classId, e.Node.Text, newName)
            End If

            If IsNothing(ch) = False Then
                'ch.toolButton.Text = cPanelManager.MakeToolButtonText(ch)
                cPanelManager.ChangePanelChildsName(ch, wOldName, wNewName)
                Dim n As TreeNode = ch.treeNode
                If IsNothing(n) = False Then
                    n.Text = e.Label
                End If
            End If
        End If
    End Sub

    Public Sub tree_AfterSelect(sender As Object, e As TreeViewEventArgs)
        primaryNameSelection = False
        Dim n As TreeNode = e.Node
        Log.PrintToLog("treeAfterSelect (start)")
        Dim isMainTree As Boolean = Object.Equals(sender, currentTreeView)
        Dim prevNode As TreeNode = IIf(isMainTree, prevNodeMainTree, prevNodeSubTree)
        If IsNothing(prevNode) = False Then
            'возвращаем цвет по-умолчанию предыдущему узлу
            If prevNode.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then prevNode.ForeColor = DEFAULT_COLORS.NodeForeColor
            If prevNode.BackColor <> DEFAULT_COLORS.ControlTransparentBackground Then prevNode.BackColor = DEFAULT_COLORS.ControlTransparentBackground
        End If
        If isMainTree AndAlso btnShowSettings.BackColor = DEFAULT_COLORS.NodeSelBackColor Then btnShowSettings.BackColor = btnAddElement.BackColor
        'выделяем цветом текущий выбранный узел
        If n.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then n.ForeColor = DEFAULT_COLORS.NodeSelForeColor
        If n.BackColor <> DEFAULT_COLORS.NodeSelBackColor Then n.BackColor = DEFAULT_COLORS.NodeSelBackColor
        prevNode = n 'сохраняем текущий узел, что бы затем можно было вернуть ему цвет по-умолчанию
        If isMainTree Then
            prevNodeMainTree = prevNode
        Else
            prevNodeSubTree = prevNode
        End If

        'Настройка вида панели управления элементами
        If isMainTree Then
            If IsNothing(sender.SelectedNode) Then
                btnDelete.Enabled = False
                tsmiRemoveElement.Enabled = False
                btnDuplicate.Enabled = False
                tsmiDuplicate.Enabled = False
                btnMakeHidden.Enabled = False
                Return
            End If

            If currentParentName.Length > 0 Then
                btnMakeHidden.Enabled = False
            Else
                btnMakeHidden.Enabled = True
            End If
            If n.ForeColor = DEFAULT_COLORS.NodeHiddenForeColor Then
                btnMakeHidden.Text = "Отобразить"
            Else
                btnMakeHidden.Text = "Скрыть"
            End If

            If n.Tag = "GROUP" AndAlso n.Nodes.Count > 0 Then
                btnDelete.Enabled = False
                tsmiRemoveElement.Enabled = False
            Else
                btnDelete.Enabled = True
                tsmiRemoveElement.Enabled = True
            End If

            If n.Tag = "GROUP" Then
                btnDuplicate.Enabled = False
                tsmiDuplicate.Enabled = False
            Else
                tsmiDuplicate.Enabled = True
                btnDuplicate.Enabled = True
            End If

            If n.Tag = "GROUP" Then Return
            'Открытие нового окна
            If questEnvironment.EnabledEvents = False Then Return
            Log.PrintToLog("treeAfterSelect: " + n.Text)

            If IsThirdLevelTree(sender) Then
                cPanelManager.OpenPanel(currentTreeView.SelectedNode, mScript.mainClassHash(currentClassName), currentParentName, "'" + n.Text + "'", currentParentName)
            Else
                cPanelManager.OpenPanel(currentTreeView.SelectedNode, mScript.mainClassHash(currentClassName), "'" + n.Text + "'", "", currentParentName)
            End If
        End If
    End Sub

    Public Sub tree_DoubleClick(sender As Object, e As EventArgs)
        If IsNothing(sender.SelectedNode) OrElse questEnvironment.treeNodeRename = clsQuestEnvironment.TreeNodeRenameEnum.SpacePressed Then Return
        sender.SelectedNode.BeginEdit()
    End Sub

    Private Sub tree_KeyUp(sender As Object, e As KeyEventArgs)
        If IsNothing(sender.SelectedNode) OrElse questEnvironment.treeNodeRename = clsQuestEnvironment.TreeNodeRenameEnum.DoubleClick Then Return
        If e.KeyCode = System.Windows.Forms.Keys.Space Then
            sender.SelectedNode.BeginEdit()
        End If
    End Sub

    Public Sub tree_DragDrop(sender As Object, e As DragEventArgs)
        Dim tree As TreeView = sender
        'Организация DragDop - завршение операции перетаскивания элементов, в т. ч. между группами
        If e.Data.GetDataPresent(GetType(TreeNode)) = False Then Return
        If Not IsNothing(nodeUnderMouseToDrop) Then nodeUnderMouseToDrop.ForeColor = DEFAULT_COLORS.NodeForeColor
        If Object.Equals(nodeUnderMouseToDrag, nodeUnderMouseToDrop) Then Return

        If e.Effect <> DragDropEffects.Move Then Return
        'операция перемещения
        'получаем класс и третий это уровань или второй
        Dim ni As NodeInfo = GetNodeInfo(nodeUnderMouseToDrag)
        Dim classId As Integer = ni.classId
        Dim thirdLevel As Boolean = ni.ThirdLevelNode

        tree.Nodes.Remove(nodeUnderMouseToDrag) 'убираем перемещаемый узел из старого места

        If nodeUnderMouseToDrag.Tag = "GROUP" Then
            'перемещение группы
            If IsNothing(nodeUnderMouseToDrop) Then
                tree.Nodes.Add(nodeUnderMouseToDrag) 'добавляем в конец
            Else
                tree.Nodes.Insert(nodeUnderMouseToDrop.Index, nodeUnderMouseToDrag)
            End If
            cGroups.MoveGroup(currentClassName, nodeUnderMouseToDrag.Text, nodeUnderMouseToDrop.Text, thirdLevel, GetParentName(tree))
        Else
            'перемещение элемента
            Dim newParent As TreeNode = nodeUnderMouseToDrop
            If IsNothing(nodeUnderMouseToDrop) Then
                tree.Nodes.Add(nodeUnderMouseToDrag) 'добавляем в конец
            Else
                If nodeUnderMouseToDrop.Tag = "ITEM" Then
                    newParent = nodeUnderMouseToDrop.Parent
                    If IsNothing(newParent) Then
                        tree.Nodes.Insert(nodeUnderMouseToDrop.Index, nodeUnderMouseToDrag)
                    Else
                        newParent.Nodes.Insert(nodeUnderMouseToDrop.Index, nodeUnderMouseToDrag)
                    End If
                Else
                    newParent.Nodes.Add(nodeUnderMouseToDrag)
                End If
            End If
            'меняем группу
            If Not thirdLevel Then
                If nodeUnderMouseToDrop.Tag = "ITEM" Then
                    'элемент к элементу
                    If ni.classId = -2 Then
                        'variables
                        Dim gName As String = mScript.csPublicVariables.lstVariables(nodeUnderMouseToDrop.Text).Group
                        mScript.csPublicVariables.lstVariables(nodeUnderMouseToDrag.Text).Group = gName
                    ElseIf ni.classId = -3 Then
                        'functions
                        Dim gName As String = mScript.functionsHash(nodeUnderMouseToDrop.Text).Group
                        Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(nodeUnderMouseToDrag.Text)
                        f.Group = gName
                    Else
                        'elements
                        Dim child2Id As Integer = GetSecondChildIdByName("'" + nodeUnderMouseToDrop.Text + "'", mScript.mainClass(ni.classId).ChildProperties)
                        If child2Id >= 0 Then
                            Dim gName As String = mScript.mainClass(ni.classId).ChildProperties(child2Id)("Group").Value
                            child2Id = GetSecondChildIdByName("'" + nodeUnderMouseToDrag.Text + "'", mScript.mainClass(ni.classId).ChildProperties)
                            If child2Id >= 0 Then SetPropertyValue(ni.classId, "Group", gName, child2Id)
                        End If
                    End If
                Else
                    'элемент к группе
                    If ni.classId = -2 Then
                        'variables
                        Dim gName As String = nodeUnderMouseToDrop.Text
                        mScript.csPublicVariables.lstVariables(nodeUnderMouseToDrag.Text).Group = gName
                    ElseIf ni.classId = -3 Then
                        'functions
                        Dim gName As String = nodeUnderMouseToDrop.Text
                        Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(nodeUnderMouseToDrag.Text)
                        f.Group = gName
                    Else
                        'elements
                        Dim gName As String = "'" + nodeUnderMouseToDrop.Text + "'"
                        Dim child2Id As Integer = GetSecondChildIdByName("'" + nodeUnderMouseToDrag.Text + "'", mScript.mainClass(ni.classId).ChildProperties)
                        If child2Id >= 0 Then SetPropertyValue(ni.classId, "Group", gName, child2Id)
                    End If
                End If
            Else
                Dim child2Id As Integer = ni.GetChild2Id
                If nodeUnderMouseToDrop.Tag = "ITEM" Then
                    'элемент к элементу
                    Dim child3Id As Integer = GetThirdChildIdByName("'" + nodeUnderMouseToDrop.Text + "'", child2Id, mScript.mainClass(ni.classId).ChildProperties)
                    If child3Id >= 0 Then
                        Dim gName As String = mScript.mainClass(ni.classId).ChildProperties(child2Id)("Group").ThirdLevelProperties(child3Id)
                        child3Id = GetThirdChildIdByName("'" + nodeUnderMouseToDrag.Text + "'", child2Id, mScript.mainClass(ni.classId).ChildProperties)
                        If child3Id >= 0 Then SetPropertyValue(ni.classId, "Group", gName, child2Id, child3Id)
                        SetPropertyValue(ni.classId, "Group", gName, child2Id, child3Id)
                    End If
                Else
                    'элемент к группе
                    Dim gName As String = nodeUnderMouseToDrop.Text
                    Dim child3Id As Integer = GetThirdChildIdByName("'" + nodeUnderMouseToDrag.Text + "'", child2Id, mScript.mainClass(ni.classId).ChildProperties)
                    If child3Id >= 0 Then SetPropertyValue(ni.classId, "Group", WrapString(gName), child2Id, child3Id)
                End If
            End If
        End If
        If classId = -1 Then Return
        'реорганизуем класс и панели всвязи с изменением порядка следования элементов

        Dim isMainTree As Boolean = Object.Equals(sender, currentTreeView)
        Dim parId As Integer
        If thirdLevel Then
            If isMainTree Then
                parId = currentParentId
            Else
                parId = ni.GetChild2Id
            End If
        Else
            parId = -1
        End If
        If classId > -2 Then
            RearrangeClass(sender, classId, parId, chkShowHidden.Checked)
            tree.SelectedNode = nodeUnderMouseToDrag
        Else
            Dim lstToExpand As New List(Of String)
            For i As Integer = 0 To tree.Nodes.Count - 1
                If tree.Nodes(i).IsExpanded Then lstToExpand.Add(tree.Nodes(i).Text)
            Next
            Dim itmText As String = nodeUnderMouseToDrag.Text
            If classId = -2 Then
                FillTreeVariables(chkShowHidden.Checked, lstToExpand)
            ElseIf classId = -3 Then
                FillTreeFunctions(chkShowHidden.Checked, lstToExpand)
            End If
            tree.SelectedNode = FindItemNodeByText(tree, itmText)
        End If
        If isMainTree Then
            'делаем активным элемент, узел которого только что переносился, получив предварительно его новые характеристики
            ni = GetNodeInfo(nodeUnderMouseToDrag)
            cPanelManager.OpenPanel(nodeUnderMouseToDrag, ni.classId, ni.nodeChild2Name, ni.nodeChild3Name, currentParentName)
        ElseIf currentClassName = "L" AndAlso ni.classId = mScript.mainClassHash("A") Then
            FillTree("A", dictTrees(mScript.mainClassHash("A")), chkShowHidden.Checked, MakeListToExpand(tree), cPanelManager.ActivePanel.child2Name)
        End If
    End Sub

    Public Sub tree_DragOver(sender As Object, e As DragEventArgs)
        'Организация DragDop - перемещение элементов внутри групп и за их пределы
        'Возникает при попадании переносимого узла в область нового узла
        Dim tree As TreeView = sender
        If IsNothing(nodeUnderMouseToDrag) Then Return

        If ((e.AllowedEffect And DragDropEffects.Move) = DragDropEffects.Move) Then
            ' By default, the drop action should be move, if allowed.
            e.Effect = DragDropEffects.Move
        Else
            e.Effect = DragDropEffects.None
        End If

        'получаем узел, который под мышью на данный момент
        Dim prevNodeDrop As TreeNode = nodeUnderMouseToDrop
        nodeUnderMouseToDrop = tree.GetNodeAt(tree.PointToClient(New Point(e.X, e.Y)))

        If IsNothing(nodeUnderMouseToDrop) Then
            If IsNothing(prevNodeDrop) = False AndAlso prevNodeDrop.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then prevNodeDrop.ForeColor = DEFAULT_COLORS.NodeForeColor
        Else
            If nodeUnderMouseToDrag.Tag = "GROUP" Then
                If nodeUnderMouseToDrop.Tag = "ITEM" Then nodeUnderMouseToDrop = nodeUnderMouseToDrop.Parent 'при перемещении групп подсвечиваем только узлы с группами 
            ElseIf currentClassName = "Variable" OrElse currentClassName = "Function" Then
                'Перемещение только между группами - узлы отсортированы
                If nodeUnderMouseToDrop.Tag = "GROUP" Then
                    If Object.Equals(nodeUnderMouseToDrag.Parent, nodeUnderMouseToDrop) Then
                        'текущий узел - узел группы, в которой находится и перетаскиваемый узел - отмена
                        e.Effect = DragDropEffects.None
                    End If
                Else
                    Dim parNode As TreeNode = nodeUnderMouseToDrop.Parent
                    If IsNothing(parNode) OrElse Object.Equals(nodeUnderMouseToDrag.Parent, parNode) Then
                        'узел под мышью не имеет группы или у него таже группа, что и у перетаскиваемого - операция запрещена
                        e.Effect = DragDropEffects.None
                    Else
                        nodeUnderMouseToDrop = parNode 'узел группы
                    End If
                End If
            End If
        End If
        'операция разрешена
        If Object.Equals(prevNodeDrop, nodeUnderMouseToDrop) = False AndAlso IsNothing(prevNodeDrop) = False AndAlso prevNodeDrop.ForeColor = DEFAULT_COLORS.NodeSelForeColor Then
            If prevNodeDrop.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then prevNodeDrop.ForeColor = DEFAULT_COLORS.NodeForeColor
        End If
        If e.Effect <> DragDropEffects.None AndAlso IsNothing(nodeUnderMouseToDrop) = False AndAlso nodeUnderMouseToDrop.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then nodeUnderMouseToDrop.ForeColor = DEFAULT_COLORS.NodeSelForeColor 'подсветка элемента, над который перемещаемый узел в данный момент
    End Sub

    Public Sub tree_MouseDown(sender As Object, e As MouseEventArgs)
        'Организация DragDop - перемещение локаций внутри и между групп
        Dim tree As TreeView = sender
        If tree.Nodes.Count = 0 Then Return

        If Not Object.Equals(currentTreeView, tree) Then
            'если дерево второстепенное, то отображаем кодбокс со стандартным событием
            Dim n As TreeNode = tree.SelectedNode
            If IsNothing(n) = False Then
                If Object.Equals(tree.GetNodeAt(e.Location), n) Then
                    Dim cbPanel As Panel = codeBoxPanel
                    Dim cb As CodeTextBox = codeBox
                    If cbPanel.Visible = False OrElse Object.Equals(cb.Tag, n) = False Then
                        cPanelManager.del_subTree_AfterSelect(tree, New TreeViewEventArgs(n))
                    End If
                End If
            End If
        End If

        nodeUnderMouseToDrag = sender.GetNodeAt(e.X, e.Y) 'получаем перемещаемый узел
        If IsNothing(nodeUnderMouseToDrag) Then
            dragBoxFromMouseDown = Rectangle.Empty
            Exit Sub
        End If
        'получаем квадрат, при выходе за пределы которого начинается операция DragDrop
        Dim dragSize As Size = SystemInformation.DragSize
        dragBoxFromMouseDown = New Rectangle(New Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize)
    End Sub

    Public Sub tree_MouseMove(sender As Object, e As MouseEventArgs)
        'Организация DragDop - перемещение узлов элементов
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If IsNothing(nodeUnderMouseToDrag) OrElse IsNothing(sender.SelectedNode) Then Exit Sub
        If dragBoxFromMouseDown.Equals(Rectangle.Empty) = False And dragBoxFromMouseDown.Contains(e.X, e.Y) = False Then
            If sender.SelectedNode.Equals(nodeUnderMouseToDrag) = False Then
                sender.SelectedNode = nodeUnderMouseToDrag 'выделяем перемещаемый узел в дереве классов
            End If
            'начинаем DragDrop
            Dim dropEffect As DragDropEffects = sender.DoDragDrop(nodeUnderMouseToDrag, DragDropEffects.Move + DragDropEffects.Scroll)
        End If
    End Sub

    Public Sub tree_MouseUp(sender As Object, e As MouseEventArgs)
        'Организация DragDop - перемещение функций и свойств между классами
        'при отпучкании кнопки мыши прекращаем перетаскивание
        dragBoxFromMouseDown = Rectangle.Empty
    End Sub

    Public Sub tree_QueryContinueDrag(sender As Object, e As QueryContinueDragEventArgs)
        'Организация DragDop - перемещение функций и свойств между классами
        If e.EscapePressed Then
            'При нажатии Esc прекращаем перетаскивание
            e.Action = DragAction.Cancel
            For i As Integer = 0 To sender.Nodes.Count - 1
                'возвращаем цвет по-умолчанию всем узлам с классами
                If sender.Nodes(i).ForeColor = DEFAULT_COLORS.NodeSelForeColor Then sender.Nodes(i).ForeColor = DEFAULT_COLORS.NodeForeColor
            Next
            'стираем данные, необходимые для перетаскивания
            nodeUnderMouseToDrag = Nothing
            dragBoxFromMouseDown = Rectangle.Empty
        End If
    End Sub
#End Region

#Region "Icons for TreeViews"
    ''' <summary>
    ''' Показывает меню выбора иконок для группы/элемента
    ''' </summary>
    ''' <param name="tree">ссылка на TreeView</param>
    Public Sub ShowIconMenu(ByRef tree As TreeView)
        If IsNothing(tree.SelectedNode) Then Return
        If tree.SelectedNode.ForeColor = DEFAULT_COLORS.NodeHiddenForeColor Then Exit Sub 'для скрытых не показываем

        'получаем ссылку на меню
        Dim imenu As ToolStrip
        If tree.SelectedNode.Tag = "GROUP" Then
            imenu = iconMenuGroups
        Else
            imenu = iconMenuElements
            ChangeMenuElementsIconsVisible(tree.SelectedNode)
        End If
        'вносим меню внутрь дерева
        If Not Object.Equals(imenu.Container, tree) Then tree.Controls.Add(imenu)
        imenu.Tag = tree.SelectedNode 'в тэге меню хранится ссылка на узел, которому меняется иконка
        CheckIconButton(imenu, imenu.Items(tree.SelectedNode.ImageKey)) 'выделяем кнопку иконки
        imenu.Show()
    End Sub

    ''' <summary>
    ''' Отображает в меню иконок те из них, которые подходят текущему классу, и прячет остальные
    ''' </summary>
    ''' <param name="n">Узел, которому устанавливается иконка</param>
    Private Sub ChangeMenuElementsIconsVisible(ByRef n As TreeNode)
        'получаем имя класса выбранного узла
        Dim isMainTree As Boolean = Object.Equals(currentTreeView, n.TreeView)
        Dim ni As NodeInfo = GetNodeInfo(n)

        Dim className As String = currentClassName
        If currentClassName = "L" AndAlso isMainTree = False Then className = "A"
        If ni.ThirdLevelNode = False Then
            className = className + "_"
        Else
            className = className + "-sub_"
        End If

        Dim imenu As ToolStrip = iconMenuElements
        For i As Integer = 0 To imenu.Items.Count - 1
            'перебираем все иконки в меню
            Dim iName As String = imenu.Items(i).Name 'имя текущей иконки
            If iName.StartsWith(className, StringComparison.CurrentCultureIgnoreCase) OrElse iName.StartsWith("item", StringComparison.CurrentCultureIgnoreCase) Then
                'имя подходит - кнопка видна
                imenu.Items(i).Visible = True
            Else
                'имя не подходит - кнопка невидна
                imenu.Items(i).Visible = False
            End If
        Next
    End Sub

    ''' <summary>
    ''' Возвращает имя иконки класса по умолчанию
    ''' </summary>
    ''' <param name="className">имя класса, для которого надо получить иконку по умолчанию</param>
    ''' <param name="thirdLevel">Для элемента 3 уровня</param>
    Public Overridable Function GetDefaultIcon(ByVal className As String, ByVal thirdLevel As Boolean) As String
        Dim iconName As String
        If thirdLevel Then
            iconName = className.ToLower + "-sub_default.png"
        Else
            iconName = className.ToLower + "_default.png"
        End If
        If iconMenuElements.Items.ContainsKey(iconName) Then Return iconName
        Return "itemDefault.png"
    End Function

    ''' <summary>
    ''' Возвращает имя иконки класса по умолчанию
    ''' </summary>
    ''' <param name="ni">Информация об узле элемента</param>
    Public Overridable Function GetDefaultIcon(ByRef ni As NodeInfo) As String
        Dim iconName As String
        Dim className As String
        If ni.classId >= 0 Then
            className = mScript.mainClass(ni.classId).Names(0)
        ElseIf ni.classId = -2 Then
            className = "variable"
        ElseIf ni.classId = -2 Then
            className = "function"
        Else
            Return ""
        End If
        If ni.ThirdLevelNode Then
            iconName = className.ToLower + "-sub_default.png"
        Else
            iconName = className.ToLower + "_default.png"
        End If
        If iconMenuElements.Items.ContainsKey(iconName) Then Return iconName
        Return "itemDefault.png"
    End Function

    ''' <summary>
    ''' Выделяет кнопку иконки на меню иконок для групп/элементов и убирает выделение с остальных
    ''' </summary>
    ''' <param name="iconMenu">меню иконок</param>
    ''' <param name="currentIconButton">текущая кнопка иконки</param> 
    Private Sub CheckIconButton(ByRef iconMenu As ToolStrip, ByRef currentIconButton As ToolStripButton)
        For i As Integer = 0 To iconMenu.Items.Count - 1
            Dim mi As ToolStripButton = iconMenu.Items(i)
            If IsNothing(currentIconButton) = False AndAlso Object.Equals(mi, currentIconButton) Then
                If mi.Checked = False Then mi.Checked = True
            Else
                If mi.Checked Then mi.Checked = False
            End If
        Next
    End Sub

    ''' <summary> Загружает иконки для TreeViev</summary>
    Private Sub LoadIconsForTrees()
        If IsNothing(imgLstGroupIcons.Images) = False AndAlso imgLstGroupIcons.Images.Count > 0 Then
            'Убираем старые изображения если они загружены
            For i As Integer = imgLstGroupIcons.Images.Count - 1 To 0 Step -1
                imgLstGroupIcons.Images(i).Dispose()
                imgLstGroupIcons.Images.RemoveAt(i)
            Next
        End If

        Dim strPath As String = questEnvironment.QuestPath + "\img\tree_icons" ' Application.StartupPath + "\src\img\tree_icons\"
        If My.Computer.FileSystem.DirectoryExists(strPath) Then
            'получаем список файлов иконок из папки квеста
            Dim quest_files As ReadOnlyCollection(Of String)
            quest_files = My.Computer.FileSystem.GetFiles(strPath, FileIO.SearchOption.SearchTopLevelOnly, "*.png")
            If IsNothing(quest_files) = False AndAlso quest_files.Count > 0 Then
                'загружаем их в контейнер imgGroupIcons с именем(ключом) = имени файла
                For i As Integer = 0 To quest_files.Count - 1
                    Dim f As String = quest_files(i)
                    imgLstGroupIcons.Images.Add(FileIO.FileSystem.GetName(f).ToLower, Image.FromFile(f))
                Next
            End If
        End If

        strPath = Application.StartupPath + "\src\img\tree_icons\"
        If My.Computer.FileSystem.DirectoryExists(strPath) Then
            'получаем список файлов иконок из общей папки
            Dim common_files As ReadOnlyCollection(Of String)
            common_files = My.Computer.FileSystem.GetFiles(strPath, FileIO.SearchOption.SearchTopLevelOnly, "*.png")
            If IsNothing(common_files) = False AndAlso common_files.Count > 0 Then
                'загружаем их в контейнер imgGroupIcons с именем(ключом) = имени файла
                For i As Integer = 0 To common_files.Count - 1
                    Dim f As String = common_files(i)
                    Dim fName As String = FileIO.FileSystem.GetName(f).ToLower
                    If imgLstGroupIcons.Images.ContainsKey(fName) Then Continue For
                    imgLstGroupIcons.Images.Add(fName, Image.FromFile(f))
                Next
            End If
        End If

    End Sub

    ''' <summary>Создает два меню для выбора иконок элементов и групп </summary> 
    Private Sub CreateIconsForTreeMenus()
        'очищаем оба меню
        If Not IsNothing(iconMenuGroups) Then
            iconMenuGroups.Visible = False
            iconMenuGroups.Items.Clear()
        End If
        If Not IsNothing(iconMenuElements) Then
            iconMenuElements.Items.Clear()
            iconMenuElements.Visible = False
        End If
        If imgLstGroupIcons.Images.Count = 0 Then Return

        Dim imenu(1) As ToolStrip 'для хранения меню

        'Создание меню с группами
        Dim strGroup As String = "group" '"item"
        imenu(0) = New ToolStrip With {.CanOverflow = True, .Dock = DockStyle.Right, .ImageList = imgLstGroupIcons, .Visible = False} 'создание меню
        For j As Integer = 0 To imgLstGroupIcons.Images.Keys.Count - 1
            'создание кнопок с иконками
            Dim mi As ToolStripButton
            Dim iconName As String = imgLstGroupIcons.Images.Keys(j).ToLower
            If Not iconName.StartsWith(strGroup) Then Continue For 'имя должно начинаться на "group"
            Dim strFinal As String = iconName.Substring(strGroup.Length)
            If strFinal.ToLower.StartsWith("hidden") Then Continue For 'иконку скрытого элемента не добавляем с именем groupHiden.png
            'непосредственное создание иконки
            mi = imenu(0).Items.Add("", imgLstGroupIcons.Images(iconName))
            AddHandler mi.Click, Sub(sender As Object, e As EventArgs)
                                     ChangeTreeNodeIcon(sender.Name, sender.Owner.Tag)
                                     CheckIconButton(sender.Owner, sender)
                                 End Sub
            mi.Name = iconName
            mi.DisplayStyle = ToolStripItemDisplayStyle.Image
        Next

        'Создание меню с иконками элементов
        Dim strItem As String = "item"
        imenu(1) = New ToolStrip With {.CanOverflow = True, .Dock = DockStyle.Right, .ImageList = imgLstGroupIcons, .Visible = False} 'создание меню
        For j As Integer = 0 To imgLstGroupIcons.Images.Keys.Count - 1
            'создание кнопок с иконками
            Dim mi As ToolStripButton
            Dim iconName As String = imgLstGroupIcons.Images.Keys(j).ToLower
            If iconName.StartsWith(strGroup) OrElse iconName.StartsWith(strItem) Then Continue For 'имя не должно начинаться на "group", "item" вставим вконце
            'непосредственное создание иконки
            mi = imenu(1).Items.Add("", imgLstGroupIcons.Images(iconName))
            AddHandler mi.Click, Sub(sender As Object, e As EventArgs)
                                     ChangeTreeNodeIcon(sender.Name, sender.Owner.Tag)
                                     CheckIconButton(sender.Owner, sender)
                                 End Sub
            mi.Name = iconName
            mi.DisplayStyle = ToolStripItemDisplayStyle.Image
        Next
        'а теперь вставляем иконки, которые начинаются на item (иконки общего назначения)
        For j As Integer = 0 To imgLstGroupIcons.Images.Keys.Count - 1
            'создание кнопок с иконками
            Dim mi As ToolStripButton
            Dim iconName As String = imgLstGroupIcons.Images.Keys(j).ToLower
            If Not iconName.StartsWith(strItem) Then Continue For 'имя должно начинаться на "item"
            Dim strFinal As String = iconName.Substring(strItem.Length)
            If strFinal.ToLower.StartsWith("hidden") Then Continue For 'иконку скрытого элемента не добавляем с именем itemHidden.png
            'непосредственное создание иконки
            mi = imenu(1).Items.Add("", imgLstGroupIcons.Images(iconName))
            AddHandler mi.Click, Sub(sender As Object, e As EventArgs)
                                     ChangeTreeNodeIcon(sender.Name, sender.Owner.Tag)
                                     CheckIconButton(sender.Owner, sender)
                                 End Sub
            mi.Name = iconName
            mi.DisplayStyle = ToolStripItemDisplayStyle.Image
        Next


        'сохраняем меню в соответствующих глобальных переменных
        iconMenuGroups = imenu(0)
        iconMenuElements = imenu(1)
    End Sub

    ''' <summary>
    ''' Изменяет иконку заданного узла (как иконки группы, так и иконки элемента)
    ''' </summary>
    ''' <param name="iconName">Название иконки</param>
    ''' <param name="node">Узел дерева, в котором надо изменить иконку</param>
    ''' <remarks></remarks>
    Private Sub ChangeTreeNodeIcon(ByVal iconName As String, ByRef node As TreeNode)
        If IsNothing(node) Then Return
        'получаем класс и его дочерние элементы
        Dim ni As NodeInfo = GetNodeInfo(node)
        Dim className As String = ""
        If ni.classId = -2 Then
            className = "Variable"
        ElseIf ni.classId = -3 Then
            className = "Function"
        Else
            className = mScript.mainClass(ni.classId).Names(0)
        End If
        'устанавливаем в соответствующее свойство новое имя иконки
        If node.Tag = "GROUP" Then
            If ni.classId < 0 Then
                cGroups.ChangeGroupIcon(className, node.Text, iconName, ni.ThirdLevelNode, "")
            Else
                cGroups.ChangeGroupIcon(className, node.Text, iconName, ni.ThirdLevelNode, GetParentName(node.TreeView))
            End If
        Else 'ITEM
            If ni.classId = -2 Then
                Dim v As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables.ElementAt(ni.GetChild2Id).Value
                'Dim vName As String = mScript.csPublicVariables.lstVariables.ElementAt(ni.nodeChild2Id).Key
                v.Icon = iconName
            ElseIf ni.classId = -3 Then
                Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(ni.nodeChild2Name)
                f.Icon = iconName
            Else
                Dim child2Id As Integer = ni.GetChild2Id
                SetPropertyValue(ni.classId, "Icon", "'" + iconName + "'", child2Id, ni.GetChild3Id(child2Id))
            End If
        End If
        'меняем иконку узла
        node.ImageKey = iconName
        node.SelectedImageKey = node.ImageKey
    End Sub

    ''' <summary>
    ''' Получаем класс и детей 2 и 3 порядка по узлу дерева
    ''' </summary>
    ''' <param name="node">Узел элемента, принадлежность которого надо выяснить</param>
    Public Function GetNodeInfo(ByRef node As TreeNode) As NodeInfo
        Dim ni As New NodeInfo
        Dim tree As TreeView = node.TreeView
        If Object.Equals(tree, currentTreeView) Then
            'имеем дело с главным деревом
            If IsNothing(cPanelManager.ActivePanel) Then
                'выбрана группа, текущей панели нет
                If currentClassName = "Variable" Then
                    ni.classId = -2
                ElseIf currentClassName = "Function" Then
                    ni.classId = -3
                Else
                    ni.classId = mScript.mainClassHash(currentClassName)
                End If
                If currentParentName.Length > 0 Then
                    ni.nodeChild2Name = currentParentName
                Else
                    ni.nodeChild2Name = ""
                End If
                ni.nodeChild3Name = ""
                ni.ThirdLevelNode = currentParentName.Length > 0
                Return ni
            Else
                ni.ThirdLevelNode = (cPanelManager.ActivePanel.child3Name.Length > 0)
            End If
            ni.nodeChild2Name = cPanelManager.ActivePanel.child2Name
            ni.nodeChild3Name = cPanelManager.ActivePanel.child3Name
            ni.classId = cPanelManager.ActivePanel.classId
        Else
            'имеем дело со вспомогательным деревом
            'Получаем класс из свойства NavigateToClassId кнопки перехода к дочерний элементам
            If IsNothing(cPanelManager.ActivePanel) Then
                Select Case currentClassName
                    Case "Function"
                        ni.classId = -3
                        ni.nodeChild2Name = ""
                        ni.nodeChild3Name = ""
                        ni.ThirdLevelNode = False
                    Case "Variable"
                        ni.classId = -2
                        ni.nodeChild2Name = ""
                        ni.nodeChild3Name = ""
                        ni.ThirdLevelNode = False
                    Case Else
                        ni.classId = mScript.mainClassHash(currentClassName)
                        ni.nodeChild2Name = ""
                        ni.nodeChild3Name = ""
                        ni.ThirdLevelNode = False
                End Select
                Return ni
            End If

            Dim NavigationButton As ButtonEx
            Dim curClassId As Integer = cPanelManager.ActivePanel.classId
            If curClassId < 0 Then
                If curClassId < -1 Then
                    'func of var
                    ni.classId = curClassId
                    ni.nodeChild2Name = cPanelManager.ActivePanel.child2Name
                    ni.nodeChild3Name = ""
                    ni.ThirdLevelNode = False
                End If
                Return ni
            End If

            NavigationButton = cPanelManager.dictDefContainers(curClassId).NavigationButton
            If IsNothing(NavigationButton) Then Return ni
            ni.classId = NavigationButton.NavigateToClassId
            ni.ThirdLevelNode = NavigationButton.NavigateToThirdLevel
            If NavigationButton.NavigateToThirdLevel Then
                If IsNothing(cPanelManager.ActivePanel) = False Then
                    ni.nodeChild2Name = cPanelManager.ActivePanel.child2Name
                    ni.nodeChild3Name = "'" + node.Text + "'" ' cPanelManager.ActivePanel.child3Name
                Else
                    ni.nodeChild2Name = ""
                    ni.nodeChild3Name = ""
                End If
                'If ni.nodeChild2Name.Length > 0 Then ni.nodeChild3Id = GetThirdChildIdByName("'" + node.Text + "'", ni.nodeChild2Id, mScript.mainClass(ni.classId).ChildProperties)
            Else
                ni.nodeChild2Name = GetSecondChildIdByName("'" + node.Text + "'", mScript.mainClass(ni.classId).ChildProperties)
            End If
            End If
            Return ni
    End Function

    ''' <summary>Содержит ли дерево элементы 3 порядка</summary>
    ''' <param name="tree">Ссылка на дерево</param>
    Public Function IsThirdLevelTree(ByVal tree As TreeView, Optional ByVal curClassId As Integer = -1) As Boolean
        If Object.Equals(tree, currentTreeView) Then
            If currentClassName = "A" OrElse currentClassName = "Variable" OrElse currentClassName = "Function" OrElse (curClassId > -1 AndAlso mScript.mainClass(curClassId).LevelsCount < 2) Then Return False
            Return currentParentName.Length > 0 'главное дерево
        End If

        If IsNothing(cPanelManager.ActivePanel) Then
            'выбрана группа, текущей панели нет
            'либо создается самый первый элемент
            If curClassId = -1 Then curClassId = mScript.mainClassHash(currentClassName)
            Dim NavigationButton As ButtonEx = cPanelManager.dictDefContainers(curClassId).NavigationButton
            If IsNothing(NavigationButton) Then Return False
            Return NavigationButton.NavigateToThirdLevel
        Else
            'имеем дело со второстепенным деревом
            If curClassId = -1 Then curClassId = cPanelManager.ActivePanel.classId
            If cPanelManager.dictDefContainers.ContainsKey(curClassId) = False Then cPanelManager.CreatePropertiesControl(curClassId)
            Dim NavigationButton As ButtonEx = cPanelManager.dictDefContainers(curClassId).NavigationButton
            If IsNothing(NavigationButton) Then Return False
            Return NavigationButton.NavigateToThirdLevel
        End If

    End Function
#End Region

#Region "CodeBox и ToolStripCodeBox"
    ''' <summary>
    ''' Меняет в кодбоксе содержимое кодом из другого свойства
    ''' </summary>
    ''' <param name="newOwner">Кнопка на панели управления нового свойства</param>
    Public Sub codeBoxChangeOwner(ByRef newOwner As Control)
        Dim btn As ButtonEx = newOwner
        codeBox.Tag = Nothing
        codeBox.Text = ""
        If IsNothing(newOwner) Then Return
        trakingEventState = trackingcodeEnum.NOT_TRACKING_EVENT
        cPanelManager.del_propertyButtonMouseClick(newOwner, New EventArgs)
    End Sub

    Public Sub codeBox_Validating(sender As CodeTextBox, e As System.ComponentModel.CancelEventArgs) Handles codeBox.Validating
        If IsNothing(sender.Tag) Then Return
        Dim curLine As Integer = codeBox.codeBox.GetLineFromCharIndex(codeBox.codeBox.SelectionStart)
        If mScript.LAST_ERROR.Length > 0 Then
            MsgBox("Нельзя сохранить код с ошибками!", vbExclamation)
            e.Cancel = True
            Return
        End If
        Dim strXml As String = ""

        Dim tName As String = sender.Tag.GetType.Name
        Select Case tName
            Case "ButtonEx"
                strXml = sender.codeBox.SerializeCodeData()
                Dim ch As clsPanelManager.clsChildPanel = sender.Tag.childPanel
                If IsNothing(ch) Then Return

                Dim child2Id As Integer = ch.GetChild2Id
                Dim child3Id As Integer = ch.GetChild3Id(child2Id)

                If String.IsNullOrWhiteSpace(strXml) Then
                    sender.Tag.Text = "(пусто)"
                    sender.Tag.ForeColor = DEFAULT_COLORS.EventButtonEmpty
                    'скрипт пустой - удаляем событие
                    If sender.Tag.IsFunctionButton Then
                        Dim eventId As Integer = mScript.mainClass(ch.classId).Functions(sender.Tag.Name).eventId
                        If eventId > 0 Then
                            mScript.eventRouter.RemoveEvent(eventId)
                            mScript.mainClass(ch.classId).Functions(sender.Tag.Name).eventId = 0
                            mScript.mainClass(ch.classId).Functions(sender.Tag.Name).Value = ""
                        End If
                    Else
                        SetPropertyValue(ch.classId, sender.Tag.Name, "", child2Id, child3Id)
                    End If
                Else
                    sender.Tag.Text = "(заполнено)"
                    sender.Tag.ForeColor = DEFAULT_COLORS.EventButtonFilled
                    mScript.eventRouter.SetPropertyWithEvent(ch.classId, sender.Tag.Name, sender.codeBox.CodeData, child2Id, child3Id, strXml, False, sender.Tag.IsFunctionButton)
                End If
            Case "clsChildPanel"
                'Функции Писателя
                Dim ch As clsPanelManager.clsChildPanel = sender.Tag
                Dim f As MatewScript.FunctionInfoType = mScript.functionsHash(ch.child2Name)
                f.ValueDt = CopyCodeDataArray(sender.codeBox.CodeData)
                f.ValueExecuteDt = mScript.PrepareBlock(f.ValueDt)
            Case "TextBoxEx"
                'вводился скрипт для обычного свойства
                Dim ch As clsPanelManager.clsChildPanel = sender.Tag.childPanel
                Dim txtLength As Integer = sender.Text.Trim.Length
                If IsNothing(ch) Then Return

                If trakingEventState = trackingcodeEnum.EVENT_BEFORE Then
                    'событие изменения свойства
                    mScript.trackingProperties.AddPropertyBefore(ch.classId, sender.Tag.Name, codeBox.codeBox.CodeData)
                ElseIf trakingEventState = trackingcodeEnum.EVENT_AFTER Then
                    'событие изменения свойства
                    mScript.trackingProperties.AddPropertyAfter(ch.classId, sender.Tag.Name, codeBox.codeBox.CodeData)
                Else
                    'просто свойство
                    'получаем тип содержимого
                    Dim cRes As MatewScript.ContainsCodeEnum = MatewScript.ContainsCodeEnum.CODE
                    If txtLength > 0 Then
                        If sender.codeBox.IsTextBlockByDefault Then
                            cRes = MatewScript.ContainsCodeEnum.LONG_TEXT
                        Else
                            Dim txt As String = sender.Text
                            If txt.Chars(0) = "?"c Then
                                cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING
                            End If
                        End If
                    Else
                        cRes = MatewScript.ContainsCodeEnum.NOT_CODE
                    End If

                    Dim child2Id As Integer = ch.GetChild2Id
                    Dim child3Id As Integer = ch.GetChild3Id(child2Id)
                    If cRes = MatewScript.ContainsCodeEnum.NOT_CODE Then
                        'кодбокс пустой - восстанавливаем прежний контрол и значение по умолчанию
                        Dim propName As String = sender.Tag.Name
                        mScript.eventRouter.RemoveEvent(mScript.eventRouter.GetEventId(ch.classId, propName, child2Id, child3Id))
                        Dim tb As Control = sender.Tag
                        sender.Tag = Nothing
                        Dim newC As Object = cPanelManager.RestoreDefaultControl(tb, True)
                        If newC.GetType.Name = "TextBoxEx" Then newC.ReadOnly = False
                        newC.Text = ""
                        SetPropertyValue(ch.classId, propName, "", child2Id, child3Id)
                        codeBoxPanel.Hide()
                        Return
                    End If
                    'кодбокс содержит скрипт/длинный текст/исполняемую строку (в любом случае кодбокс не пустой)
                    sender.Tag.ReadOnly = True
                    If cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
                        'исполняемая строка
                        strXml = WrapString(sender.Text)
                        sender.Tag.Text = sender.Text
                    ElseIf cRes = MatewScript.ContainsCodeEnum.LONG_TEXT Then
                        'длинный текст
                        strXml = sender.codeBox.SerializeCodeData()
                        sender.Tag.Text = My.Resources.longText
                    Else
                        'скрипт
                        strXml = sender.codeBox.SerializeCodeData()
                        sender.Tag.Text = My.Resources.script
                    End If

                    mScript.eventRouter.SetPropertyWithEvent(ch.classId, sender.Tag.Name, sender.codeBox.CodeData, child2Id, child3Id, strXml, cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING, False)
                End If
            Case "ComboBoxEx"
                If trakingEventState = trackingcodeEnum.NOT_TRACKING_EVENT Then Return
                'тут может быть только событие редактирования свойства
                Dim ch As clsPanelManager.clsChildPanel = sender.Tag.childPanel
                Dim txtLength As Integer = sender.Text.Trim.Length
                If IsNothing(ch) Then Return
                If trakingEventState = trackingcodeEnum.EVENT_BEFORE Then
                    mScript.trackingProperties.AddPropertyBefore(ch.classId, sender.Tag.Name, codeBox.codeBox.CodeData)
                Else
                    mScript.trackingProperties.AddPropertyAfter(ch.classId, sender.Tag.Name, codeBox.codeBox.CodeData)
                End If
            Case "TreeNode"
                strXml = sender.codeBox.SerializeCodeData()
                Dim ni As NodeInfo = GetNodeInfo(sender.Tag)
                Dim propName As String = mScript.mainClass(ni.classId).DefaultProperty
                If String.IsNullOrEmpty(propName) Then Return
                Dim child2Id As Integer = ni.GetChild2Id
                mScript.eventRouter.SetPropertyWithEvent(ni.classId, propName, sender.codeBox.CodeData, child2Id, ni.GetChild3Id(child2Id), strXml, False, False)
        End Select
    End Sub


    Private Sub tsmiClassEditor_Click(sender As Object, e As EventArgs) Handles tsmiClassEditor.Click
        If codeBoxPanel.Visible Then
            codeBox.codeBox.CheckTextForSyntaxErrors(codeBox.codeBox)
            Call codeBox_Validating(codeBox, New System.ComponentModel.CancelEventArgs)
        ElseIf dgwVariables.Visible Then
            Call dgwVariables_Validating(dgwVariables, New System.ComponentModel.CancelEventArgs)
        End If

        frmClassEditor.Show()
    End Sub

    ''' <summary>
    ''' Подготавливает панель кодбокса - вставляет кнопки управления, назначает события...
    ''' </summary>
    Private Sub PrepareCodeBoxPanel()
        codeBoxPanel.Controls.Add(ToolStripCodeBox)
        ToolStripCodeBox.Dock = DockStyle.Top
        codeBoxPanel.Controls.Add(codeBox)
        AddHandler codeBoxPanel.Resize, Sub(sender2 As Object, e2 As EventArgs)
                                            Dim cbPanel As Panel = sender2
                                            codeBox.Location = New Point(0, ToolStripCodeBox.Bottom)
                                            codeBox.Size = New Size(cbPanel.ClientSize.Width, cbPanel.ClientSize.Height - codeBox.Top)
                                        End Sub
        AddHandler codeBoxPanel.VisibleChanged, Sub(sender As Object, e As EventArgs)
                                                    If IsNothing(cPanelManager) Then Return
                                                    If IsNothing(cPanelManager.ActivePanel) Then Return
                                                    cPanelManager.ActivePanel.IsCodeBoxVisible = sender.Visible
                                                End Sub

        SplitInner.Panel2.Controls.Add(codeBoxPanel)
    End Sub

    ''' <summary>
    ''' Обрамляет выделенный текст в кодбоксе указанными тэгами
    ''' </summary>
    ''' <param name="codeBox">Кодбокс</param>
    ''' <param name="openTag">Открывающий тэг</param>
    ''' <param name="closeTag">Закрывающий тэг</param>
    ''' <remarks></remarks>
    Private Sub WrapWithTag(codeBox As CodeTextBox, ByVal openTag As String, ByVal closeTag As String)
        Dim cb As RichTextBox = codeBox.codeBox
        Dim selStart As Integer = cb.SelectionStart, selLength As Integer = cb.SelectionLength
        cb.SelectedText = openTag + cb.SelectedText + closeTag
        cb.SelectionStart = selStart + openTag.Length
        cb.SelectionLength = selLength
    End Sub

    Private Sub tsmiBR_Click(sender As Object, e As EventArgs) Handles tsmiBR.Click, tsbBR.Click
        codeBox.codeBox.SelectedText = "<BR/>"
    End Sub

    Private Sub tsmiHR_Click(sender As Object, e As EventArgs) Handles tsmiHR.Click, tsbHR.Click
        codeBox.codeBox.SelectedText = "<HR/>"
    End Sub

    Private Sub tsmiB_Click(sender As Object, e As EventArgs) Handles tsmiB.Click, tsbB.Click
        WrapWithTag(codeBox, "<STRONG>", "</STRONG>")
    End Sub

    Private Sub tsmiI_Click(sender As Object, e As EventArgs) Handles tsmiI.Click, tsbI.Click
        WrapWithTag(codeBox, "<EM>", "</EM>")
    End Sub

    Private Sub tsmiU_Click(sender As Object, e As EventArgs) Handles tsmiU.Click, tsbU.Click
        WrapWithTag(codeBox, "<U>", "</U>")
    End Sub

    Private Sub tsmiSup_Click(sender As Object, e As EventArgs) Handles tsmiSup.Click, tsbSup.Click
        WrapWithTag(codeBox, "<SUP>", "</SUP>")
    End Sub

    Private Sub tsmiSub_Click(sender As Object, e As EventArgs) Handles tsmiSub.Click, tsbSub.Click
        WrapWithTag(codeBox, "<SUB>", "</SUB>")
    End Sub

    Private Sub tsmiP_Click(sender As Object, e As EventArgs) Handles tsmiP.Click, tsbP.Click
        WrapWithTag(codeBox, "<P>", "</P>")
    End Sub

    Private Sub tsmiH1_Click(sender As Object, e As EventArgs) Handles tsmiH1.Click, tsbH1.Click
        WrapWithTag(codeBox, "<H1>", "</H1>")
    End Sub

    Private Sub tsmiH2_Click(sender As Object, e As EventArgs) Handles tsmiH2.Click, tsbH2.Click
        WrapWithTag(codeBox, "<H2>", "</H2>")
    End Sub

    Private Sub tsmiH3_Click(sender As Object, e As EventArgs) Handles tsmiH3.Click, tsbH3.Click
        WrapWithTag(codeBox, "<H3>", "</H3>")
    End Sub

    Private Sub tsmiSpan_Click(sender As Object, e As EventArgs) Handles tsmiSpan.Click, tsbSpan.Click
        WrapWithTag(codeBox, "<SPAN>", "</SPAN>")
    End Sub

    Private Sub tsmiLeft_Click(sender As Object, e As EventArgs) Handles tsmiLeft.Click, tsbLeft.Click
        WrapWithTag(codeBox, "<DIV Align=Left>", "</DIV>")
    End Sub

    Private Sub tsmiCenter_Click(sender As Object, e As EventArgs) Handles tsmiCenter.Click, tsbCenter.Click
        WrapWithTag(codeBox, "<DIV Align=Center>", "</DIV>")
    End Sub

    Private Sub tsmiRight_Click(sender As Object, e As EventArgs) Handles tsmiRight.Click, tsbRight.Click
        WrapWithTag(codeBox, "<DIV Align=Right>", "</DIV>")
    End Sub

    Private Sub tsmiJustify_Click(sender As Object, e As EventArgs) Handles tsmiJustify.Click, tsbJustify.Click
        WrapWithTag(codeBox, "<DIV Align=Justify>", "</DIV>")
    End Sub

    Private Sub tsmiIFrame_Click(sender As Object, e As EventArgs) Handles tsmiIFrame.Click
        WrapWithTag(codeBox, "<IFRAME src=" & Chr(34) & "www.Ваш адрес.com" & Chr(34) & " Width=100%>", "</IFRAME>")
    End Sub

    Private Sub tsmiQuotes_Click(sender As Object, e As EventArgs) Handles tsmiQuotes.Click, tsbQuotes.Click
        WrapWithTag(codeBox, "&laquo;", "&raquo;")
    End Sub

    Private Sub tsmiSpace_Click(sender As Object, e As EventArgs) Handles tsmiSpace.Click, tsbSpace.Click
        codeBox.codeBox.SelectedText = "&nbsp;"
    End Sub

    Private Sub tsmiDash_Click(sender As Object, e As EventArgs) Handles tsmiDash.Click, tsbDash.Click
        codeBox.codeBox.SelectedText = "&#150;"
    End Sub

    Private Sub tsmiExecuteScript_Click(sender As Object, e As EventArgs) Handles tsmiExecuteScript.Click
        dlgExecute.ShowDialog(Me)
    End Sub
#End Region

    'Public Sub InjectScript(Script As String)
    '    WBhelp.Document.InvokeScript("eval", New Object() {Script})
    'End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim aStr As String = "-5.5"
        If IsNumeric(aStr.ToString(System.Globalization.NumberFormatInfo.CurrentInfo)) Then
            Dim v As Double = Double.Parse(aStr)
            v = Math.Round(-0.5)
        End If

        'Dim a As Integer = 0
        'Dim aStr As String = currentClassName
        'aStr &= "; parent: " & currentParentName & "; "
        'Dim ch As clsPanelManager.clsChildPanel = cPanelManager.ActivePanel
        'If IsNothing(ch) Then
        '    aStr &= "active: none"
        'Else
        '    aStr &= "active: " & mScript.mainClass(ch.classId).Names(0) & ", ch2: " & ch.child2Name
        '    aStr &= ",ch3: " & ch.child3Name & ",supra: " & ch.supraElementName
        'End If
        'MsgBox(aStr)

        'AddHandler wmp.PlayStateChange, Sub(Result As Integer, aa As Integer)
        '                                    MsgBox(aa.ToString)
        '                                End Sub


        'Do
        '    Application.DoEvents()
        '    Me.Text = wmp.PlayState.ToString
        'Loop
        'Dim arrC() As Object = Me.Controls.Find("LabelPlayAssorted", True)
        'If IsNothing(arrC) = False AndAlso arrC.Count > 0 Then
        '    Dim l As Label = arrC(0)
        '    l.Hide()
        'End If

        'dlgExecute.codeMain.codeBox.Show()
        'Dim img As Image = Image.FromFile("D:\Projects\MatewQuest2\MatewQuest2\bin\Debug\src\img\tree_icons\favicon.ico")
        'img.Save("D:\Projects\MatewQuest2\MatewQuest2\bin\Debug\src\img\tree_icons\favicon.png", System.Drawing.Imaging.ImageFormat.Png)
        Dim i As Integer = 0
        '&times\; #,#.##;&times\; -#,#.##

        'Dim aStr As String
        'aStr.pa

        'aa = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, "D:\src/img/d.txt")
        'If IsNothing(s.Panel1.VerticalScroll) = False Then s.Panel1.VerticalScroll.Value = 0
        'Dim sWatch As New Stopwatch
        'sWatch.Start()
        'cPanelManager = New clsPanelManager(SplitInner, WBhelp, ToolStripPanels, btnShowSettings)
        'MsgBox(sWatch.ElapsedMilliseconds.ToString)
        'sWatch.Stop()


        '19442 ms (def)
        '12867 ms

        'InjectScript(My.Computer.FileSystem.ReadAllText(Environment.CurrentDirectory & "\Plugin.js".ToString))
        'Dim hDoc As HtmlDocument = WBhelp.Document


        'Dim a As Integer = 0
        'Dim t1 As Long, t2 As Long
        'Dim sWatch As New Stopwatch
        'Dim times As Integer = 10000
        'Dim c As ComboBoxEx = Me.Controls.Find("IsDead", True)(0)
        't1 = 0
        'questEnvironment.codeBoxShadowed.LoadCodeFromProperty(mScript.mainClass(1).Properties("LocationExitEvent").Value)
        'Dim cd() As CodeTextBox.CodeDataType = questEnvironment.codeBoxShadowed.codeBox.CodeData
        'sWatch.Start()
        'For i As Integer = 1 To times
        '    mScript.PrepareBlock(cd)
        'Next
        't1 = sWatch.ElapsedTicks
        'sWatch.Reset()
        'sWatch.Stop()



        'MessageBox.Show("Old: " + t1.ToString + "; New: " + t2.ToString)
    End Sub

#Region "HTML"
    Private Sub WBhelp_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WBhelp.DocumentCompleted
        If abortNavigationEvent Then
            abortNavigationEvent = False
            Return
        End If
        treeProperties.Hide()
        hDocument = sender.Document

        If IsNothing(hDocument) Then Return

        Dim myControl2 As Object = Nothing
        Dim classId As Integer
        Dim child2Name As String
        Dim child3Name As String
        Dim child2Id As Integer 'только для forSubTree
        Dim child3Id As Integer 'только для forSubTree 
        Dim forSubTree As Boolean = False

        If IsNothing(WBhelp.Tag) = False AndAlso WBhelp.Tag.GetType.Name = "TreeNode" Then
            'Файл помощи загружается для свойства по умолчанию узла, выбранного в поддереве
            forSubTree = True
            Dim ni As NodeInfo = GetNodeInfo(WBhelp.Tag)
            classId = ni.classId
            child2Name = ni.nodeChild2Name
            child3Name = ni.nodeChild3Name
            child2Id = ni.GetChild2Id
            child3Id = ni.GetChild3Id(child2Id)
            If child2Id < 0 Then Return
        Else
            'Файл помощи загружается для обычного свойства
            If IsNothing(cPanelManager.ActivePanel) Then Return
            With cPanelManager.ActivePanel
                classId = .classId
                child2Name = .child2Name
                child3Name = .child3Name
                child2Id = .GetChild2Id
                child3Id = .GetChild3Id(child2Id)
            End With
            myControl2 = cPanelManager.ActivePanel.ActiveControl
            If IsNothing(myControl2) Then Return
        End If


        'Отображаем описание согласно уровню класса элемента
        Dim El As HtmlElement = hDocument.GetElementById("Level1")
        If IsNothing(El) Then El = hDocument.GetElementById("Level1_1")
        Dim lvlExt As Integer = 0
        Do While Not IsNothing(El)
            If child2Name.Length = 0 Then
                El.Style = "" '"display:block"
            Else
                El.Style = "display:none"
            End If
            lvlExt += 1
            El = hDocument.GetElementById("Level1" + "_" + lvlExt.ToString)
        Loop

        El = hDocument.GetElementById("Level2")
        If IsNothing(El) Then El = hDocument.GetElementById("Level2_1")
        lvlExt = 0
        Do While Not IsNothing(El)
            If child2Name.Length > 0 AndAlso child3Name.Length = 0 Then
                El.Style = "" '"display:block"
            Else
                El.Style = "display:none"
            End If
            lvlExt += 1
            El = hDocument.GetElementById("Level2" + "_" + lvlExt.ToString)
        Loop

        El = hDocument.GetElementById("Level3")
        If IsNothing(El) Then El = hDocument.GetElementById("Level3_1")
        lvlExt = 0
        Do While Not IsNothing(El)
            If child3Name.Length > 0 Then
                El.Style = "" '"display:block"
            Else
                El.Style = "display:none"
            End If
            lvlExt += 1
            El = hDocument.GetElementById("Level3" + "_" + lvlExt.ToString)
        Loop

        'Вносим коррективы в заголовок (формулу свойства) в зависимости от текущего уровня класса 
        If child2Name.Length = 0 Then
            El = hDocument.GetElementById("propElement1")
            If IsNothing(El) = False Then El.Style = "display:none"
        ElseIf child3Name.Length = 0 Then
            El = hDocument.GetElementById("propElement2")
            If IsNothing(El) = False Then El.Style = "display:none"
        End If

        If forSubTree = False AndAlso myControl2.IsFunctionButton Then
            'контрол функции
            '!!!!!!!!!!!!!!!!!!!!!
        Else
            'контрол свойства
            Dim propName As String
            If forSubTree Then
                propName = mScript.mainClass(classId).DefaultProperty
            Else
                propName = myControl2.Name()
            End If
            Dim prop As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties(propName)

            'если есть таблица выбора слота, то устанавливаем текущий слот
            El = hDocument.GetElementById("slotContainer")
            If Not IsNothing(El) Then
                Dim slot As Integer = 0
                If child2Id = -1 Then
                    slot = Val(mScript.PrepareStringToPrint(prop.Value, Nothing, False))
                ElseIf child3Id = -1 Then
                    slot = Val(mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False))
                Else
                    slot = Val(mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False))
                End If
                El = hDocument.GetElementById("slot" & (slot + 1).ToString)
                If IsNothing(El) = False Then El.SetAttribute("className", "selected")
            End If



            'размещаем дополнительные элементы для удобства выбора значения свойства
            Select Case prop.returnType
                Case MatewScript.ReturnFunctionEnum.RETURN_ELEMENT
                    'Свойство возвращает элемент (локацию, меню и т. д.)
                    If IsNothing(prop.returnArray) OrElse prop.returnArray.Count = 0 Then Return
                    Dim treeSrc As TreeView = Nothing
                    Dim retClassName As String = prop.returnArray(0) 'получили одно из имен класса
                    Dim retClassId As Integer = 0
                    If mScript.mainClassHash.TryGetValue(retClassName, retClassId) = False Then
                        If retClassName = "Variable" Then
                            retClassId = -2
                        ElseIf retClassName = "Function" Then
                            retClassId = -3
                        Else
                            Return
                        End If
                    Else
                        retClassName = mScript.mainClass(retClassId).Names(0) 'теперь здесь первое имя класса
                    End If
                    'Размещаем дерево с узлами соответствующего элемента
                    If retClassId >= 0 Then
                        treeSrc = dictTrees(retClassId)
                        If mScript.mainClass(retClassId).LevelsCount = 2 Then
                            '3-уровневый класс. Необходимо обновить дерево, так как там могут быть элементы 3 порядка
                            FillTree(retClassName, treeSrc, chkShowHidden.Checked, Nothing, "", True)
                        End If
                    ElseIf retClassId = -2 Then
                        treeSrc = treeVariables
                    Else
                        treeSrc = treeFunctions
                    End If

                    'Копируем узлы в дерево на браузере
                    treeProperties.BeginUpdate()
                    treeProperties.Nodes.Clear()
                    'treeProperties.ImageList = treeSrc.ImageList
                    For i As Integer = 0 To treeSrc.Nodes.Count - 1
                        treeProperties.Nodes.Add(CType(treeSrc.Nodes(i).Clone, TreeNode))
                        If treeSrc.Nodes(i).IsExpanded Then treeProperties.Nodes(i).Expand()
                    Next i
                    treeProperties.EndUpdate()

                    Dim elRetValues As HtmlElement = hDocument.GetElementById("returnValues") 'html-элемент, который надо спрятать если будет показано дерево (для экономии места по высоте)

                    If treeProperties.Nodes.Count = 0 Then
                        'дерево пустое
                        treeProperties.Hide()
                        If IsNothing(elRetValues) = False Then elRetValues.Style = "display:block"
                    Else
                        'дерево заполнено - отображаем его и устанавливаем положение
                        If IsNothing(elRetValues) = False Then elRetValues.Style = "display:none"
                        treeProperties.Width = Math.Max(treeSrc.Width, 300)
                        treeProperties.Show()
                        WBhelp_Resize(sender, New EventArgs)

                        'выделяем узел с выбранным значением
                        Dim curValue As String
                        If forSubTree Then
                            If child3Name.Length > 0 Then
                                curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False)
                            Else
                                curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False)
                            End If
                        Else
                            curValue = myControl2.Text
                        End If
                        If String.IsNullOrEmpty(curValue) Then Return
                        Dim elId As Integer = 0
                        If Integer.TryParse(curValue, elId) AndAlso retClassId >= 0 Then
                            'в свойстве указано Id элемента. Получаем имя
                            If IsNothing(mScript.mainClass(retClassId).ChildProperties) OrElse elId > mScript.mainClass(retClassId).ChildProperties.Count - 1 Then Return
                            curValue = mScript.PrepareStringToPrint(mScript.mainClass(retClassId).ChildProperties(elId)("Name").Value, Nothing, False)
                        End If
                        Dim n As TreeNode = FindItemNodeByText(treeProperties, curValue)
                        If IsNothing(n) = False Then
                            treeProperties.SelectedNode = n
                            n.BackColor = DEFAULT_COLORS.NodeSelBackColor
                        End If
                    End If
                Case MatewScript.ReturnFunctionEnum.RETURN_FUNCTION
                    'Свойство возвращает функцию Писателя
                    'Размещаем дерево с узлами соответствующего элемента
                    Dim treeSrc As TreeView = treeFunctions

                    'Копируем узлы в дерево на браузере
                    treeProperties.BeginUpdate()
                    treeProperties.Nodes.Clear()
                    'treeProperties.ImageList = treeSrc.ImageList
                    For i As Integer = 0 To treeSrc.Nodes.Count - 1
                        treeProperties.Nodes.Add(CType(treeSrc.Nodes(i).Clone, TreeNode))
                        If treeSrc.Nodes(i).IsExpanded Then treeProperties.Nodes(i).Expand()
                    Next i
                    treeProperties.EndUpdate()

                    Dim elRetValues As HtmlElement = hDocument.GetElementById("returnValues") 'html-элемент, который надо спрятать если будет показано дерево (для экономии места по высоте)

                    If treeProperties.Nodes.Count = 0 Then
                        'дерево пустое
                        treeProperties.Hide()
                        If IsNothing(elRetValues) = False Then elRetValues.Style = "display:block"
                    Else
                        'дерево заполнено - отображаем его и устанавливаем положение
                        If IsNothing(elRetValues) = False Then elRetValues.Style = "display:none"
                        treeProperties.Width = treeSrc.Width
                        treeProperties.Show()
                        WBhelp_Resize(sender, New EventArgs)

                        'выделяем узел с выбранным значением
                        Dim curValue As String
                        If forSubTree Then
                            If child3Name.Length > 0 Then
                                curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False)
                            Else
                                curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False)
                            End If
                        Else
                            curValue = myControl2.Text
                        End If
                        If String.IsNullOrEmpty(curValue) Then Return
                        Dim elId As Integer = 0
                        If Integer.TryParse(curValue, elId) Then
                            'в свойстве указано Id элемента. Получаем имя
                            If IsNothing(mScript.functionsHash) OrElse elId > mScript.functionsHash.Count - 1 Then Return
                            curValue = mScript.functionsHash.ElementAt(elId).Key
                        End If
                        Dim n As TreeNode = FindItemNodeByText(treeProperties, curValue)
                        If IsNothing(n) = False Then
                            treeProperties.SelectedNode = n
                            n.BackColor = DEFAULT_COLORS.NodeSelBackColor
                        End If
                    End If
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE
                    'Путь к изображению
                    If cListManager.HasFiles(cListManagerClass.fileTypesEnum.PICTURES) Then
                        FillTreeWithFolders(treeProperties, cListManagerClass.fileTypesEnum.PICTURES) 'Создаем узлы - папки с изображениями

                        El = hDocument.GetElementById("returnValues")
                        If IsNothing(El) = False Then El.Style = "display:none"
                        WBhelp_Resize(WBhelp, New EventArgs)

                        If treeProperties.Nodes.Count = 0 Then Return
                        Dim curValue As String
                        If forSubTree Then
                            If child3Name.Length > 0 Then
                                curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False)
                            Else
                                curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False)
                            End If
                        Else
                            curValue = myControl2.Text
                        End If
                        'выводим текущую картинку
                        Dim hSample As HtmlElement = hDocument.GetElementById("sampleContainer")
                        If IsNothing(hSample) = False Then
                            Dim newImg As HtmlElement = hDocument.CreateElement("IMG")
                            newImg.SetAttribute("className", "thumbnail")
                            newImg.SetAttribute("src", questEnvironment.QuestPath + "\" + curValue)
                            hSample.AppendChild(newImg)
                        End If

                        'Выбираем папку в которой находится выбранная картинка
                        If String.IsNullOrEmpty(curValue) Then
                            treeProperties.SelectedNode = treeProperties.Nodes(0)
                            Return
                        End If
                        Dim fPath As String
                        Dim aPos As Integer = curValue.LastIndexOfAny({"/"c, "\"c})
                        If aPos = -1 Then
                            fPath = curValue
                        Else
                            fPath = curValue.Substring(0, aPos)
                        End If
                        Dim nCol() As TreeNode = treeProperties.Nodes.Find(fPath.Replace("/", "\"), True)
                        If IsNothing(nCol) = False AndAlso nCol.Count > 0 Then
                            treeProperties.SelectedNode = nCol(0)
                        Else
                            treeProperties.SelectedNode = treeProperties.Nodes(0)
                        End If
                    End If
                Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO, MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS, MatewScript.ReturnFunctionEnum.RETURN_PATH_JS, _
                    MatewScript.ReturnFunctionEnum.RETURN_PATH_TEXT
                    'Путь к аудиофайлу
                    Dim fType As cListManagerClass.fileTypesEnum
                    Select Case prop.returnType
                        Case MatewScript.ReturnFunctionEnum.RETURN_PATH_AUDIO
                            fType = cListManagerClass.fileTypesEnum.AUDIO
                        Case MatewScript.ReturnFunctionEnum.RETURN_PATH_CSS
                            fType = cListManagerClass.fileTypesEnum.CSS
                        Case MatewScript.ReturnFunctionEnum.RETURN_PATH_JS
                            fType = cListManagerClass.fileTypesEnum.JS
                        Case MatewScript.ReturnFunctionEnum.RETURN_PATH_TEXT
                            fType = cListManagerClass.fileTypesEnum.TEXT
                    End Select
                    If cListManager.HasFiles(fType) Then
                        FillTreeWithFolders(treeProperties, fType, True) 'Создаем узлы - папки с аудио

                        El = hDocument.GetElementById("returnValues")
                        If IsNothing(El) = False Then El.Style = "display:none"
                        WBhelp_Resize(WBhelp, New EventArgs)

                        If treeProperties.Nodes.Count = 0 Then Return
                        Dim curValue As String
                        If forSubTree Then
                            If child3Name.Length > 0 Then
                                curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False)
                            Else
                                curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False)
                            End If
                        Else
                            curValue = myControl2.Text
                        End If

                        'Выбираем папку в которой находится выбранный аудиофайл
                        If String.IsNullOrEmpty(curValue) Then
                            treeProperties.SelectedNode = treeProperties.Nodes(0)
                            Return
                        End If
                        Dim nCol() As TreeNode = treeProperties.Nodes.Find(curValue, True)
                        If IsNothing(nCol) = False AndAlso nCol.Count > 0 Then
                            treeProperties.SelectedNode = nCol(0)
                        Else
                            treeProperties.SelectedNode = treeProperties.Nodes(0)
                        End If
                    End If
                Case MatewScript.ReturnFunctionEnum.RETURN_COLOR
                    'Цвет
                    Dim hDataContainer As HtmlElement = hDocument.GetElementById("dataContainer")
                    If IsNothing(hDataContainer) Then Return
                    Dim mapPath As String = Application.StartupPath + "\src\img\colorMap.bmp"
                    If My.Computer.FileSystem.FileExists(mapPath) Then
                        El = hDocument.GetElementById("returnValues")
                        If IsNothing(El) = False Then El.Style = "display:none"

                        'получаем элемент с картой цвета
                        Dim hMap As HtmlElement = hDocument.GetElementById("colorMap")
                        If IsNothing(hMap) Then Return
                        hMap.SetAttribute("src", mapPath)
                        'получаем элемент для вывода сэмпла цвета
                        Dim hSample As HtmlElement = hDocument.GetElementById("colorSample")
                        If IsNothing(hSample) Then Return
                        'выводим в него текущий выбранный цвет
                        Dim cText As String
                        If forSubTree Then
                            If child3Name.Length > 0 Then
                                cText = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False)
                            Else
                                cText = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False)
                            End If
                        Else
                            cText = myControl2.Text
                        End If
                        Dim strAlpha As String = "1"
                        If cText.StartsWith("rgba(") = False Then
                            If cText.StartsWith("#") AndAlso cText.Length = 7 Then
                                Dim r As Byte, g As Byte, b As Byte
                                Byte.TryParse(cText.Substring(1, 2), System.Globalization.NumberStyles.HexNumber, provider_points, r)
                                Byte.TryParse(cText.Substring(3, 2), System.Globalization.NumberStyles.HexNumber, provider_points, g)
                                Byte.TryParse(cText.Substring(5, 2), System.Globalization.NumberStyles.HexNumber, provider_points, b)
                                cText = "rgba(" + r.ToString + "," + g.ToString + "," + b.ToString + ", 1)"
                            Else
                                cText = "rgba(0,0,0,1)"
                            End If
                        Else
                            'get alpha from rgba
                            Dim lastComma As Integer = cText.LastIndexOf(","c)
                            If lastComma > -1 And cText.EndsWith(")") Then
                                strAlpha = cText.Substring(lastComma + 1, cText.Length - lastComma - 2).Trim
                            End If
                        End If
                        hSample.Style = "background-color:" + cText
                        'получаем элемент-текстбокс для ввода прозрачности
                        Dim hTransp As HtmlElement = hDocument.GetElementById("colorTransparency")
                        hTransp.SetAttribute("value", strAlpha)
                        If IsNothing(hTransp) = False Then
                            AddHandler hMap.Click, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                       'устанавливаем цвет пикселя, по которому кликнули
                                                       Dim curCol As Color = imgSelectColorMap.GetPixel(e2.OffsetMousePosition.X, e2.OffsetMousePosition.Y)
                                                       Dim transp As Double = 1
                                                       If IsNothing(transp) = False Then
                                                           'получаем прозрачность
                                                           Double.TryParse(hTransp.GetAttribute("Value").Replace(",", "."), System.Globalization.NumberStyles.Float, provider_points, transp)
                                                       End If
                                                       Dim strRGBA As String = "rgba(" + curCol.R.ToString + "," + curCol.G.ToString + "," + curCol.B.ToString + ", " + Convert.ToString(transp, provider_points) + ")"
                                                       If forSubTree Then
                                                           SetPropertyValue(classId, propName, WrapString(strRGBA), child2Id, child3Id)
                                                       Else
                                                           myControl2.Text = strRGBA
                                                       End If
                                                       hSample.Style = "background-color:" + strRGBA
                                                       hTransp.SetAttribute("value", "1")
                                                   End Sub
                            AddHandler hTransp.LosingFocus, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                                'устанавливаем цвет с новой прозрачностью
                                                                Dim strColor As String = hSample.Style, strColorResult As String = ""
                                                                Dim r As Byte, g As Byte, b As Byte, alpha As Double
                                                                Double.TryParse(hTransp.GetAttribute("Value").Replace(",", "."), System.Globalization.NumberStyles.Float, provider_points, alpha)
                                                                If strColor.StartsWith("background-color: rgba(") Then
                                                                    'color format is RGBA
                                                                    strColor = strColor.Substring("background-color: rgba(".Length)
                                                                    If strColor.EndsWith(")") Then strColor = strColor.Substring(0, strColor.Length - 1)
                                                                    Dim arrColors() As String = strColor.Split(","c)
                                                                    If arrColors.Count <> 4 Then Return
                                                                    Byte.TryParse(arrColors(0).Trim, r)
                                                                    Byte.TryParse(arrColors(1).Trim, g)
                                                                    Byte.TryParse(arrColors(2).Trim, b)
                                                                    If alpha >= 1 Then
                                                                        strColorResult = "rgb(" + r.ToString + "," + g.ToString + "," + b.ToString + ")"
                                                                    Else
                                                                        strColorResult = "rgba(" + r.ToString + "," + g.ToString + "," + b.ToString + ", " + Convert.ToString(alpha, provider_points) + ")"
                                                                    End If
                                                                ElseIf strColor.StartsWith("background-color: rgb(") Then
                                                                    'color format is RGB
                                                                    strColor = strColor.Substring("background-color: rgb(".Length)
                                                                    If strColor.EndsWith(")") Then strColor = strColor.Substring(0, strColor.Length - 1)
                                                                    Dim arrColors() As String = strColor.Split(","c)
                                                                    If arrColors.Count <> 3 Then Return
                                                                    Byte.TryParse(arrColors(0).Trim, r)
                                                                    Byte.TryParse(arrColors(1).Trim, g)
                                                                    Byte.TryParse(arrColors(2).Trim, b)
                                                                    If alpha >= 1 Then
                                                                        strColorResult = "rgb(" + r.ToString + "," + g.ToString + "," + b.ToString + ")"
                                                                    Else
                                                                        strColorResult = "rgba(" + r.ToString + "," + g.ToString + "," + b.ToString + ", " + Convert.ToString(alpha, provider_points) + ")"
                                                                    End If
                                                                ElseIf strColor.StartsWith("background-color: #") Then
                                                                    'color format is #rrggbb                                                                    
                                                                    strColor = strColor.Substring("background-color: #".Length).Trim
                                                                    If alpha >= 1 Then
                                                                        'no transparency
                                                                        strColorResult = "#" & strColor
                                                                    Else
                                                                        'transparency, alpha < 1
                                                                        Byte.TryParse(strColor.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, provider_points, r)
                                                                        Byte.TryParse(strColor.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, provider_points, g)
                                                                        Byte.TryParse(strColor.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, provider_points, b)
                                                                        strColorResult = "rgba(" + r.ToString + "," + g.ToString + "," + b.ToString + ", " + Convert.ToString(alpha, provider_points) + ")"
                                                                    End If
                                                                Else
                                                                    'unknown format
                                                                    Return
                                                                End If

                                                                If forSubTree Then
                                                                    SetPropertyValue(classId, propName, WrapString(strColorResult), child2Id, child3Id)
                                                                Else
                                                                    myControl2.Text = strColorResult
                                                                End If
                                                                hSample.Style = "background-color:" + strColorResult
                                                            End Sub
                        End If
                        'выводим выбранные цвета
                        Dim hColorsContainer As HtmlElement = hDocument.GetElementById("selectedColorsContainer")
                        If IsNothing(hColorsContainer) Then Return
                        For i As Integer = 0 To questEnvironment.lstSelectedColors.Count - 1
                            Dim hColor As HtmlElement = hDocument.CreateElement("SPAN")
                            hColor.SetAttribute("className", "selectedColor")
                            hColor.SetAttribute("Title", questEnvironment.lstSelectedColors(i))
                            hColor.Style = "background-color: " + questEnvironment.lstSelectedColors(i)
                            hColorsContainer.AppendChild(hColor)
                            AddHandler hColor.Click, Sub(sender2 As Object, e2 As HtmlElementEventArgs)
                                                         Dim strColor As String = sender2.GetAttribute("Title")
                                                         If forSubTree Then
                                                             SetPropertyValue(classId, propName, WrapString(strColor), child2Id, child3Id)
                                                         Else
                                                             myControl2.Text = strColor
                                                         End If
                                                         hSample.Style = "background-color:" + strColor
                                                         If strColor.StartsWith("#") OrElse strColor.StartsWith("rgb(") Then
                                                             hTransp.SetAttribute("value", "1")
                                                         ElseIf strColor.StartsWith("rgba(") Then
                                                             Dim val As String = strColor
                                                             Dim pos As Integer = val.LastIndexOf(")"c)
                                                             If pos = -1 Then Return
                                                             val = val.Substring(0, pos)
                                                             pos = val.LastIndexOf(","c)
                                                             If pos = -1 Then Return
                                                             val = val.Substring(pos + 1).Trim
                                                             hTransp.SetAttribute("value", val)
                                                         End If
                                                     End Sub
                        Next

                        Dim fileContents As String
                        fileContents = My.Computer.FileSystem.ReadAllText(Application.StartupPath + "\src\colors.html")


                        Dim hFrame As HtmlElement = hDocument.CreateElement("DIV")
                        hFrame.InnerHtml = fileContents
                        hDataContainer.AppendChild(hFrame)
                        hFrame.Id = "colorsFrame"
                        'Dim wFrame As HtmlWindow = hDocument.Window.Frames("colorsFrame")
                        'Dim hFrame As HtmlWindow = tmpFrame
                        'abortNavigationEvent = True
                        'wFrame.Navigate(Application.StartupPath + "\src\colors.html")
                        'abortNavigationEvent = False
                    End If
                Case Else
                    Dim hPlaceMarker As HtmlElement = hDocument.GetElementById("placeMarker")
                    If IsNothing(hPlaceMarker) Then Return
                    Dim strSpecialType As String = hPlaceMarker.GetAttribute("specialType")
                    If String.IsNullOrEmpty(strSpecialType) Then Return
                    Dim strSource As String
                    Dim strSourceParams As String

                    Select Case strSpecialType.ToLower
                        Case "classesfromcss", "animationsfromcss"
                            Dim isAni As Boolean = (strSpecialType.ToLower = "animationsfromcss")

                            If child2Name.Length = 0 Then
                                strSource = hPlaceMarker.GetAttribute("sourceLevel1")
                                strSourceParams = hPlaceMarker.GetAttribute("sourceParamsLevel1")
                            ElseIf child3Name.Length = 0 Then
                                strSource = hPlaceMarker.GetAttribute("sourceLevel2")
                                strSourceParams = hPlaceMarker.GetAttribute("sourceParamsLevel2")
                            Else
                                strSource = hPlaceMarker.GetAttribute("sourceLevel3")
                                strSourceParams = hPlaceMarker.GetAttribute("sourceParamsLevel3")
                            End If

                            'Отображаем список классов, подгруженных из указанного свойства
                            Dim strPattern As String = hPlaceMarker.GetAttribute("pattern")
                            Dim strCSS As String = ""
                            Dim lstClasses As List(Of String) = Nothing
                            If isAni Then
                                lstClasses = cListManager.GetCssAnimationsByPattern(strSource, strSourceParams, strPattern, strCSS)
                            Else
                                lstClasses = cListManager.GetCssClassesByPattern(strSource, strSourceParams, strPattern, strCSS)
                            End If

                            If lstClasses.Count > 0 Then
                                'Список классов не пуст

                                If hDocument.Window.Frames.Count > 0 Then
                                    Dim frmSample As HtmlWindow = Nothing
                                    Try
                                        frmSample = hDocument.Window.Frames("sampleFrame")
                                        If IsNothing(frmSample) = False Then
                                            'выводим семплы классов
                                            abortNavigationEvent = True
                                            HtmlAppendCSS(frmSample.Document, FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, strCSS))
                                            If isAni Then
                                                Dim hSampleClass As String = hPlaceMarker.GetAttribute("sampleClass")
                                                Dim hEffDurationProp As String = hPlaceMarker.GetAttribute("durationProperty")
                                                Dim hDuration As Integer = 1000
                                                If String.IsNullOrEmpty(hEffDurationProp) = False Then ReadPropertyInt(classId, hEffDurationProp, child2Id, child3Id, hDuration, Nothing)
                                                FillFrameWithAnimations(frmSample, lstClasses, myControl2, forSubTree, propName, classId, child2Id, child3Id, hSampleClass, hDuration)
                                            Else
                                                FillFrameWithClasses(frmSample, lstClasses, myControl2, forSubTree, propName, classId, child2Id, child3Id)
                                            End If
                                        End If
                                    Catch ex As Exception

                                    End Try
                                End If

                                'заполняем дерево
                                'treeProperties.ImageList = imgLstGroupIcons
                                treeProperties.BeginUpdate()
                                treeProperties.Nodes.Clear()
                                For i As Integer = 0 To lstClasses.Count - 1
                                    Dim n As TreeNode = treeProperties.Nodes.Add(lstClasses(i), lstClasses(i), "pattern.png", "pattern.png")
                                    n.Tag = "ITEM"
                                Next
                                treeProperties.ExpandAll()
                                treeProperties.EndUpdate()
                                treeProperties.Show()
                                'прячем строку помощи "Возвращаемое значение", так как здесь лишнее
                                El = hDocument.GetElementById("returnValues")
                                If IsNothing(El) = False Then El.Style = "display:none"
                                WBhelp_Resize(WBhelp, New EventArgs)

                                If treeProperties.Nodes.Count = 0 Then Return
                                Dim curValue As String
                                If forSubTree Then
                                    If child3Name.Length > 0 Then
                                        curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False)
                                    Else
                                        curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False)
                                    End If
                                Else
                                    curValue = myControl2.Text
                                End If

                                'Выделяем выбранный класс
                                If String.IsNullOrEmpty(curValue) Then
                                    'treeProperties.SelectedNode = treeProperties.Nodes(0)
                                    Return
                                End If
                                Dim nCol() As TreeNode = treeProperties.Nodes.Find(curValue, True)
                                If IsNothing(nCol) = False AndAlso nCol.Count > 0 Then
                                    treeProperties.SelectedNode = nCol(0)
                                Else
                                    treeProperties.SelectedNode = Nothing ' treeProperties.Nodes(0)
                                End If
                            End If
                        Case "idfromproperty"
                            strSource = hPlaceMarker.GetAttribute("source")
                            If String.IsNullOrEmpty(strSource) Then Return
                            Dim lstId As List(Of String) = cListManager.GetListOfID(strSource)
                            lstId.Add("DescriptionText")

                            If lstId.Count > 0 Then
                                'Список идентификаторов не пуст
                                'заполняем дерево
                                'treeProperties.ImageList = imgLstGroupIcons
                                treeProperties.BeginUpdate()
                                treeProperties.Nodes.Clear()
                                Dim nImg As String = "element.png"

                                For i As Integer = 0 To lstId.Count - 1
                                    If i = lstId.Count - 1 Then nImg = "element_spec.png"
                                    Dim n As TreeNode = treeProperties.Nodes.Add(lstId(i), lstId(i), nImg, nImg)
                                    n.Tag = "ITEM"
                                Next
                                treeProperties.ExpandAll()
                                treeProperties.EndUpdate()
                                treeProperties.Show()
                                'прячем строку помощи "Возвращаемое значение", так как здесь лишнее
                                El = hDocument.GetElementById("returnValues")
                                If IsNothing(El) = False Then El.Style = "display:none"
                                WBhelp_Resize(WBhelp, New EventArgs)

                                If treeProperties.Nodes.Count = 0 Then Return
                                Dim curValue As String
                                If forSubTree Then
                                    If child3Name.Length > 0 Then
                                        curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id), Nothing, False)
                                    Else
                                        curValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value, Nothing, False)
                                    End If
                                Else
                                    curValue = myControl2.Text
                                End If

                                'Выделяем выбранный класс
                                If String.IsNullOrEmpty(curValue) Then
                                    treeProperties.SelectedNode = treeProperties.Nodes(0)
                                    Return
                                End If
                                Dim nCol() As TreeNode = treeProperties.Nodes.Find(curValue, True)
                                If IsNothing(nCol) = False AndAlso nCol.Count > 0 Then
                                    treeProperties.SelectedNode = nCol(0)
                                Else
                                    treeProperties.SelectedNode = Nothing ' treeProperties.Nodes(0)
                                End If
                            End If
                    End Select
            End Select
        End If
    End Sub

    ''' <summary>
    ''' Заполняет iframe примерами классов
    ''' </summary>
    ''' <param name="frm">фрейм</param>
    ''' <param name="lstClasses">список классов</param>
    ''' <param name="mControl">контрол для вывода названия класса при клике</param>
    ''' <param name="forSubTree">поддерево (свойство по умолчанию)?</param>
    ''' <param name="propName">название свойства</param>
    ''' <param name="classId"></param>
    ''' <param name="child2Id"></param>
    ''' <param name="child3Id"></param>
    Private Sub FillFrameWithClasses(ByRef frm As HtmlWindow, ByRef lstClasses As List(Of String), ByRef mControl As Control, ByVal forSubTree As Boolean, ByVal propName As String, _
                                     ByVal classId As Integer, ByVal child2Id As Integer, ByVal child3Id As Integer)
        Dim hBody As HtmlElement = frm.Document.Body
        If IsNothing(hBody) Then Return
        For i As Integer = 0 To lstClasses.Count - 1
            Dim hFrame As HtmlElement = frm.WindowFrameElement
            Dim itemDisplayInBlock As String = hFrame.GetAttribute("itemDisplayInBlock")

            Dim hSpan As HtmlElement = Nothing
            If itemDisplayInBlock = "True" Then
                Dim hDiv As HtmlElement = hDocument.CreateElement("DIV")
                hDiv.Style = "position:relative;padding:5px;margin:5px;display:block;overflow:auto;"
                hBody.AppendChild(hDiv)

                hSpan = hDocument.CreateElement("DIV")
                hSpan.Style = "cursor:pointer"
                hDiv.AppendChild(hSpan)
            Else
                Dim strWidth As String = hFrame.GetAttribute("itemWidth")
                If String.IsNullOrEmpty(strWidth) Then strWidth = "64px"
                Dim strHeight As String = hFrame.GetAttribute("itemHeight")
                If String.IsNullOrEmpty(strHeight) Then strHeight = "64px"
                hSpan = hDocument.CreateElement("SPAN")
                hSpan.Style = "position:relative;padding:5px;width:" & strWidth & ";height:" & strHeight & ";cursor:pointer;margin:5px;display:inline-block;overflow:hidden;"
                hBody.AppendChild(hSpan)
            End If

            Dim itmName As String = lstClasses(i)
            hSpan.InnerHtml = itmName
            hSpan.SetAttribute("ClassName", itmName)
            hSpan.SetAttribute("Title", itmName)
            Dim controlCopy As Control = mControl
            AddHandler hSpan.Click, Sub(sender As Object, e As HtmlElementEventArgs)
                                        'устанавливаем выбранное значение класса

                                        Dim nodeFound As Boolean = False
                                        If treeProperties.Visible Then
                                            'если открыто дерево с классами, то выбираем соответствующий узел
                                            Dim n As TreeNode = FindItemNodeByText(treeProperties, itmName)
                                            If IsNothing(n) = False Then
                                                nodeFound = True
                                                treeProperties.SelectedNode = n
                                            End If
                                        End If

                                        If Not nodeFound Then
                                            'если узел выбран не был (напр., не отображалось дерево), то устанавливаем новое значение самостоятельно
                                            If forSubTree Then
                                                SetPropertyValue(classId, propName, WrapString(itmName), child2Id, child3Id)
                                            Else
                                                controlCopy.Text = itmName
                                            End If
                                        End If
                                    End Sub
        Next i

    End Sub


    ''' <summary>
    ''' Заполняет iframe примерами анимаций
    ''' </summary>
    ''' <param name="frm">фрейм</param>
    ''' <param name="lstAnimations">список анимаций</param>
    ''' <param name="mControl">контрол для вывода названия класса при клике</param>
    ''' <param name="forSubTree">поддерево (свойство по умолчанию)?</param>
    ''' <param name="propName">название свойства</param>
    ''' <param name="classId"></param>
    ''' <param name="child2Id"></param>
    ''' <param name="child3Id"></param>
    ''' <param name="sampleClass">атрибут sampleClass of html-div-element with id="placeMarker" (from html-help-file)</param>
    Private Sub FillFrameWithAnimations(ByRef frm As HtmlWindow, ByRef lstAnimations As List(Of String), ByRef mControl As Control, ByVal forSubTree As Boolean, ByVal propName As String, _
                                     ByVal classId As Integer, ByVal child2Id As Integer, ByVal child3Id As Integer, ByVal sampleClass As String, ByVal duration As Integer)
        Dim hBody As HtmlElement = frm.Document.Body

        If IsNothing(hBody) Then Return
        For i As Integer = 0 To lstAnimations.Count - 1
            Dim hFrame As HtmlElement = frm.WindowFrameElement
            Dim itemDisplayInBlock As String = hFrame.GetAttribute("itemDisplayInBlock")

            Dim hSpan As HtmlElement = Nothing
            If itemDisplayInBlock = "True" Then
                Dim hDiv As HtmlElement = hDocument.CreateElement("DIV")
                hDiv.Style = "position:relative;padding:5px;margin:5px;display:block;overflow:auto;"
                hBody.AppendChild(hDiv)

                hSpan = hDocument.CreateElement("DIV")
                hSpan.Style = "cursor:pointer"
                hDiv.AppendChild(hSpan)
            Else
                Dim strWidth As String = hFrame.GetAttribute("itemWidth")
                If String.IsNullOrEmpty(strWidth) Then strWidth = "64px"
                Dim strHeight As String = hFrame.GetAttribute("itemHeight")
                If String.IsNullOrEmpty(strHeight) Then strHeight = "64px"
                hSpan = hDocument.CreateElement("SPAN")
                hSpan.Style = "position:relative;padding:5px;width:" & strWidth & ";height:" & strHeight & ";cursor:pointer;margin:5px;display:inline-block;overflow:hidden;"
                hBody.AppendChild(hSpan)
            End If

            Dim itmName As String = lstAnimations(i)
            hSpan.InnerHtml = itmName
            hSpan.Style &= "animation-name:" & itmName & ";animation-duration: " & duration.ToString & "ms;"
            hSpan.SetAttribute("ClassName", sampleClass)
            hSpan.SetAttribute("Title", itmName)
            Dim controlCopy As Control = mControl
            AddHandler hSpan.Click, Sub(sender As Object, e As HtmlElementEventArgs)
                                        'устанавливаем выбранное значение класса

                                        Dim nodeFound As Boolean = False
                                        If treeProperties.Visible Then
                                            'если открыто дерево с классами, то выбираем соответствующий узел
                                            Dim n As TreeNode = FindItemNodeByText(treeProperties, itmName)
                                            If IsNothing(n) = False Then
                                                nodeFound = True
                                                treeProperties.SelectedNode = n
                                            End If
                                        End If

                                        If Not nodeFound Then
                                            'если узел выбран не был (напр., не отображалось дерево), то устанавливаем новое значение самостоятельно
                                            If forSubTree Then
                                                SetPropertyValue(classId, propName, WrapString(itmName), child2Id, child3Id)
                                            Else
                                                controlCopy.Text = itmName
                                            End If
                                        End If
                                    End Sub
        Next i

    End Sub

    ''' <summary>
    ''' Переменная для предотвращения события DocumentCompleted WbHelp-а при загрузке во фрейм
    ''' </summary>
    ''' <remarks></remarks>
    Public abortNavigationEvent As Boolean = False

    ''' <summary>
    ''' Заполняет дерево, отображаемое в веббраузере, папками
    ''' </summary>
    ''' <param name="filesType">тип файлов, папки которых отображаются</param>
    ''' <remarks></remarks>
    Public Sub FillTreeWithFolders(ByRef tree As TreeView, ByVal filesType As cListManagerClass.fileTypesEnum, Optional addFilesToList As Boolean = False)
        tree.BeginUpdate()
        tree.Nodes.Clear()

        Dim igroupDef As String = "groupDefault.png"
        Dim igroupSel As String = "group01.png"
        Dim iFileDef As String
        Dim iFileSel As String
        Select Case filesType
            Case cListManagerClass.fileTypesEnum.AUDIO
                iFileDef = "play.png"
                iFileSel = "pause.png"
            Case Else
                iFileDef = "file.png"
                iFileSel = "file.png"
        End Select

        Dim arrFolders() As String = cListManager.CreateFoldersList(filesType)
        Dim n As TreeNode = Nothing
        n = tree.Nodes.Add("", questEnvironment.QuestShortName, igroupDef, igroupSel) 'Корневой каталог квеста
        Dim lstFiles As List(Of String)
        If addFilesToList Then
            lstFiles = cListManager.GetFolderFiles(filesType, "") 'получаем список файлов корневой директории
            For i As Integer = 0 To lstFiles.Count - 1
                Dim nF As TreeNode = n.Nodes.Add(lstFiles(i), lstFiles(i), iFileDef, iFileSel)
                nF.Tag = "ITEM"
            Next
        End If
        If arrFolders.Count > 0 Then
            For i As Integer = 0 To arrFolders.Count - 1
                'перебираем все папки (пути относительно папки квеста)
                Dim curPath As String = arrFolders(i)
                If tree.Nodes.Find(curPath, True).Count > 0 Then Continue For 'такой каталог уже есть - выход
                Dim arrPath() As String = curPath.Split("\"c) 'разбиваем путь по папкам

                Dim colN As TreeNodeCollection = tree.Nodes(0).Nodes 'получаем набор узлов корневого узла (папка квеста)
                For j As Integer = 0 To arrPath.Length - 1
                    Dim semiPath As String = String.Join("\", arrPath, 0, j + 1) 'часть пути от корня до вложенной папки № j - 1
                    Dim aPos As Integer = colN.IndexOfKey(semiPath)
                    If aPos = -1 Then
                        'узел, ассоциированный с этой частью пути еще не создан - создаем его
                        n = colN.Add(semiPath, arrPath(j), igroupDef, igroupSel)
                        If addFilesToList Then
                            lstFiles = cListManager.GetFolderFiles(filesType, semiPath) 'получаем список файлов текущей директории
                            For u As Integer = 0 To lstFiles.Count - 1
                                Dim nF As TreeNode = n.Nodes.Add(semiPath + "\" + lstFiles(u), lstFiles(u), iFileDef, iFileSel)
                                nF.Tag = "ITEM"
                            Next u
                        End If
                    Else
                        'узел, ассоциированный с этой частью пути уже есть - только получаем этот узел
                        n = colN(aPos)
                    End If
                    colN = n.Nodes 'получаем набор узлов данного узла - ассоциированного с частью пути до вложенной папки № j - 1
                Next j
            Next i
        End If
        tree.ExpandAll()
        tree.EndUpdate()
        tree.Show()
    End Sub

    Private Sub WBhelp_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs) Handles WBhelp.Navigating
        If abortNavigationEvent Then Return
        If IsNothing(hDocument) Then Return 'NEWW

        Dim hSample As HtmlElement = hDocument.GetElementById("colorSample")
        If IsNothing(hSample) = False Then
            'Добавляем цвет в список выбранных
            Dim strColor As String = hSample.Style
            If strColor.StartsWith("background-color: rgb") = False AndAlso strColor.StartsWith("background-color: #") = False Then Return
            strColor = strColor.Substring("background-color: ".Length).Trim
            If strColor.EndsWith(";") Then strColor = strColor.Substring(0, strColor.Length - 1).Trim
            If questEnvironment.lstSelectedColors.IndexOf(strColor) = -1 Then questEnvironment.lstSelectedColors.Add(strColor)
        End If
    End Sub

    Private Sub WBhelp_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles WBhelp.PreviewKeyDown
        If e.Control AndAlso e.KeyCode = Keys.O Then
            e.IsInputKey = True
            tsmiLoad.PerformClick()
        ElseIf e.KeyCode = Keys.F5 AndAlso e.Control = False AndAlso e.Alt = False AndAlso e.Shift = False Then
            e.IsInputKey = True
            tsmiRun.PerformClick()
        End If
    End Sub

    Private Sub WBhelp_Resize(sender As Object, e As EventArgs) Handles WBhelp.Resize
        If treeProperties.Visible Then
            'Размещаем дерево с элементами (если видимо) в окне браузера. Маркером высоты является html-элемент с id="placeMarker"
            treeProperties.Left = (WBhelp.ClientRectangle.Width - treeProperties.Width) / 2
            Dim hDocument As HtmlDocument = WBhelp.Document
            If IsNothing(hDocument) Then Return
            Dim hPlaceMarker As HtmlElement = hDocument.GetElementById("placeMarker")
            If IsNothing(hPlaceMarker) Then Return

            Dim hTop As Integer = 0, hLeft As Integer = 0
            If IsNothing(hPlaceMarker) = False Then
                hTop = hPlaceMarker.OffsetRectangle.Top
                hLeft = hPlaceMarker.OffsetRectangle.Left
                Dim pEl As HtmlElement = hPlaceMarker.OffsetParent
                Do While Not pEl.TagName = "BODY"
                    hTop += pEl.OffsetRectangle.Top
                    hLeft += pEl.OffsetRectangle.Left
                    pEl = pEl.OffsetParent
                Loop
            End If
            treeProperties.Location = New Point((WBhelp.ClientRectangle.Width - treeProperties.Width) / 2, hTop)

            If hPlaceMarker.GetAttribute("fitToWidth") = "True" Then
                treeProperties.Top += 10
                treeProperties.Width = hPlaceMarker.OffsetRectangle.Width
                treeProperties.Left = hLeft
            End If

            treeProperties.Height = WBhelp.ClientRectangle.Height - treeProperties.Top - questEnvironment.defPaddingTop
        End If
    End Sub

    Private Sub WBhelp_VisibleChanged(sender As Object, e As EventArgs) Handles WBhelp.VisibleChanged
        Dim ch As clsPanelManager.clsChildPanel = cPanelManager.ActivePanel
        If IsNothing(ch) Then Return
        ch.IsWbVisible = WBhelp.Visible

        If mPlayer.FileName.Length > 0 Then
            'Прекращаем проигрывание аудио если надо
            mPlayer.Stop()
            mPlayer.FileName = ""
        End If
    End Sub

    Private Sub hDocument_Click(sender As Object, e As HtmlElementEventArgs) Handles hDocument.Click
        If WBhelp.ReadyState <> WebBrowserReadyState.Complete Then Return
        Dim hEl As HtmlElement = hDocument.GetElementFromPoint(e.ClientMousePosition)
        If IsNothing(hEl) Then Return 'NEWW
        Dim classId As Integer
        Dim child2Id As Integer 'только для forSubTree
        Dim child3Id As Integer 'только для forSubTree
        Dim c2 As Object = Nothing
        Dim propName As String
        Dim forSubTree As Boolean = False

        If IsNothing(WBhelp.Tag) = False AndAlso WBhelp.Tag.GetType.Name = "TreeNode" Then
            'Свойство по умолчанию для элементов 3 уровня
            forSubTree = True
            Dim ni As NodeInfo = GetNodeInfo(WBhelp.Tag)
            classId = ni.classId
            child2Id = ni.GetChild2Id
            child3Id = ni.GetChild3Id(child2Id)
            propName = mScript.mainClass(classId).DefaultProperty
        Else
            'Обычное свойство текущей панели
            If IsNothing(cPanelManager.ActivePanel) Then Return
            classId = cPanelManager.ActivePanel.classId
            child2Id = cPanelManager.ActivePanel.GetChild2Id
            child3Id = cPanelManager.ActivePanel.GetChild3Id(child2Id)
            c2 = cPanelManager.ActivePanel.ActiveControl
            If IsNothing(c2) Then Return
            propName = c2.Name
        End If

        Do While Not hEl.TagName = "BODY"
            'если в html-элементе присутствует тэг rVal, то значит внутри него содержится некое значение, которое надо установить текущему свойству
            Dim res As String = hEl.GetAttribute("rVal")
            If res.Length > 0 AndAlso res.ToLower = "[title]" Then res = hEl.GetAttribute("Title")

            If res.Length > 0 Then
                Dim prevValue As String = ""
                If forSubTree Then
                    If child2Id < 0 Then
                        prevValue = mScript.mainClass(classId).Properties(propName).Value
                    ElseIf child3Id < 0 Then
                        prevValue = mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value
                    Else
                        prevValue = mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id)
                    End If
                    prevValue = mScript.PrepareStringToPrint(prevValue, Nothing, False)
                Else
                    prevValue = c2.Text
                End If

                If res <> prevValue Then
                    'устанавливаем свойству новое значение
                    If forSubTree Then
                        Dim wRes As String = WrapString(res)
                        SetPropertyValue(classId, propName, wRes, child2Id, child3Id)
                    Else
                        Dim cIndex As Integer = 0
                        If c2.GetType.Name = "ComboBoxEx" AndAlso Integer.TryParse(res, cIndex) Then
                            Dim cmb As ComboBoxEx = c2
                            If cIndex <= cmb.Items.Count - 1 Then
                                cmb.SelectedIndex = cIndex
                            Else
                                c2.Text = res
                            End If
                        Else
                            c2.Text = res
                        End If
                        SetPropertyValue(classId, propName, WrapString(c2.Text), child2Id, child3Id)
                    End If

                    'Если это свойство вида Picture (и новое значение не пустое), то устанавливаем свойство PictureHover и PictureActive
                    If String.IsNullOrEmpty(res) = False AndAlso mScript.mainClass(classId).Properties(propName).returnType = MatewScript.ReturnFunctionEnum.RETURN_PATH_PICTURE AndAlso _
                        propName.EndsWith("Picture") Then
                        Dim propNewName As String = propName & "Hover"
                        If mScript.mainClass(classId).Properties.ContainsKey(propNewName) Then
                            'в данном классе существует свойство вида PictureHover
                            Dim pos As Integer = res.LastIndexOf(".")
                            If pos > -1 Then
                                Dim newName As String = res.Substring(0, pos) & "Hover" & res.Substring(pos)
                                If FileIO.FileSystem.FileExists(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, newName)) Then
                                    'существует файл [Имя картинки]Hover[.ext]
                                    If forSubTree = False Then
                                        Dim arrCont() As Control = cPanelManager.dictDefContainers(classId).Controls.Find(propNewName, True)
                                        If arrCont.Count > 0 Then
                                            arrCont(0).Text = newName
                                        End If
                                    End If
                                    SetPropertyValue(classId, propNewName, WrapString(newName), child2Id, child3Id)
                                End If
                            End If
                        End If

                        propNewName = propName & "Active"
                        If mScript.mainClass(classId).Properties.ContainsKey(propNewName) Then
                            'в данном классе существует свойство вида PictureActive
                            Dim pos As Integer = res.LastIndexOf(".")
                            If pos > -1 Then
                                Dim newName As String = res.Substring(0, pos) & "Active" & res.Substring(pos)
                                If FileIO.FileSystem.FileExists(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, newName)) Then
                                    'существует файл [Имя картинки]Active[.ext]
                                    If forSubTree = False Then
                                        Dim arrCont() As Control = cPanelManager.dictDefContainers(classId).Controls.Find(propNewName, True)
                                        If arrCont.Count > 0 Then
                                            arrCont(0).Text = newName
                                        End If
                                    End If
                                    SetPropertyValue(classId, propNewName, WrapString(newName), child2Id, child3Id)
                                End If
                            End If
                        End If

                        propNewName = propName & "Current"
                        If mScript.mainClass(classId).Properties.ContainsKey(propNewName) Then
                            'в данном классе существует свойство вида PictureCurrent
                            Dim pos As Integer = res.LastIndexOf(".")
                            If pos > -1 Then
                                Dim newName As String = res.Substring(0, pos) & "Current" & res.Substring(pos)
                                If FileIO.FileSystem.FileExists(FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, newName)) Then
                                    'существует файл [Имя картинки]Current[.ext]
                                    If forSubTree = False Then
                                        Dim arrCont() As Control = cPanelManager.dictDefContainers(classId).Controls.Find(propNewName, True)
                                        If arrCont.Count > 0 Then
                                            arrCont(0).Text = newName
                                        End If
                                    End If
                                    SetPropertyValue(classId, propNewName, WrapString(newName), child2Id, child3Id)
                                End If
                            End If
                        End If
                    End If


                    If hEl.GetAttribute("bgcolor").Length > 0 AndAlso res.Length = 7 AndAlso res.StartsWith("#") Then
                        'надо отобразить выбранный цвет в colorSample
                        Dim hSample As HtmlElement = hDocument.GetElementById("colorSample")
                        Dim hTransp As HtmlElement = hDocument.GetElementById("colorTransparency")
                        If IsNothing(hSample) = False Then
                            Dim r As Byte, g As Byte, b As Byte
                            Byte.TryParse(res.Substring(1, 2), System.Globalization.NumberStyles.HexNumber, provider_points, r)
                            Byte.TryParse(res.Substring(3, 2), System.Globalization.NumberStyles.HexNumber, provider_points, g)
                            Byte.TryParse(res.Substring(5, 2), System.Globalization.NumberStyles.HexNumber, provider_points, b)
                            Dim strARGB As String = "rgba(" + r.ToString + "," + g.ToString + "," + b.ToString + ",1)"
                            hSample.Style = "background-color:" + strARGB
                            If IsNothing(hTransp) = False Then hTransp.SetAttribute("value", "1")
                        End If
                    End If

                    If hEl.GetAttribute("dontAddPrevValue") <> "True" Then
                        'Создаем строку в документе для возврата старого значения
                        'Получаем элемент TBODY
                        Dim hPar As HtmlElement = hEl
                        If hPar.TagName = "TD" Then
                            hPar = hPar.Parent.Parent
                        ElseIf hEl.TagName = "TR" Then
                            hPar = hPar.Parent
                        End If
                        If Not IsHTMLElementContainsRVal(hPar, prevValue) Then
                            'Добавляем нову строку в таблицу
                            Dim tr As HtmlElement = hDocument.CreateElement("TR")
                            hPar.AppendChild(tr)
                            Dim td As HtmlElement = hDocument.CreateElement("TD")
                            td.InnerHtml = "Предыдущее значение"
                            tr.AppendChild(td)
                            td = hDocument.CreateElement("TD")
                            td.SetAttribute("rVal", prevValue)
                            td.SetAttribute("className", "selectable")
                            td.InnerHtml = prevValue
                            tr.AppendChild(td)
                        End If
                    End If
                End If
            End If
            hEl = hEl.Parent
        Loop
    End Sub

    ''' <summary>
    ''' Рекурсивная функция, проверяющая наличие тэга rVal в html-эдементах с заданным значанием
    ''' </summary>
    ''' <param name="tBody">Родительский html-элменет для поиска</param>
    ''' <param name="value">Значение rVal для поиска</param>
    Private Function IsHTMLElementContainsRVal(ByRef tBody As HtmlElement, ByVal value As String) As Boolean
        If IsNothing(tBody) Then Return False
        If tBody.CanHaveChildren AndAlso IsNothing(tBody.Children) = False AndAlso tBody.Children.Count > 0 Then
            For i As Integer = 0 To tBody.Children.Count - 1
                Dim hEl As HtmlElement = tBody.Children(i)
                If hEl.GetAttribute("rVal") = value Then Return True
                If IsHTMLElementContainsRVal(hEl, value) Then Return True
            Next
        End If
        Return False
    End Function

    ''' <summary>
    ''' Пишет текст события на html-страничку
    ''' </summary>
    Public Sub HtmlShowEventText(ByRef relatedControl As Object)
        'Для вывода кода события в html на страничке должен быть элемент с id = HelpConvas. Он послужит родителем для выводимого кода
        Do While Not WBhelp.ReadyState = WebBrowserReadyState.Complete
            Application.DoEvents()
        Loop
        If IsNothing(hDocument) Then Return
        Dim hasConvas As Boolean = True
        Dim elFrame As HtmlElement
        elFrame = hDocument.GetElementById("DataContainer")
        If IsNothing(elFrame) Then
            elFrame = hDocument.GetElementById("HelpConvas")
            hasConvas = False
        End If
        If IsNothing(elFrame) Then Return

        Dim isFunction As Boolean = relatedControl.IsFunctionButton
        'Получаем текст события
        Dim propName As String
        Dim forSubTree As Boolean = False
        Dim classId As Integer
        Dim child2Id As Integer
        Dim child3Id As Integer

        If IsNothing(WBhelp.Tag) = False AndAlso WBhelp.Tag.GetType.Name = "TreeNode" Then
            'Свойство по умолчанию для элементов 3 уровня
            forSubTree = True
            Dim ni As NodeInfo = GetNodeInfo(WBhelp.Tag)
            classId = ni.classId
            child2Id = ni.GetChild2Id
            child3Id = ni.GetChild3Id(child2Id)
            propName = mScript.mainClass(classId).DefaultProperty
        Else
            'Обычное свойство текущей панели
            If IsNothing(cPanelManager.ActivePanel) Then Return
            classId = cPanelManager.ActivePanel.classId
            Dim ch As clsPanelManager.clsChildPanel = relatedControl.childPanel
            child2Id = ch.GetChild2Id
            child3Id = ch.GetChild3Id(child2Id)
            Dim tName As String = relatedControl.GetType.Name
            If tName = "ButtonEx" AndAlso relatedControl.Name.ToString.EndsWith("Help") Then
                'Имя события (если relatedControl - кнопка справки с именем, например, DescriptionHelp) - для всех встроенных свойств-событий
                propName = relatedControl.Name.Substring(0, relatedControl.Name.Length - 4)
            Else
                propName = relatedControl.Name 'Имя свойства
            End If
        End If


        Dim prop As MatewScript.PropertiesInfoType
        If isFunction Then
            prop = mScript.mainClass(classId).Functions(propName)
        Else
            prop = mScript.mainClass(classId).Properties(propName)
        End If
        Dim retType As MatewScript.ReturnFunctionEnum = prop.returnType

        'Получаем текст события в xml
        Dim strXml As String
        If child2Id = -1 Then
            strXml = prop.Value
        ElseIf child3Id = -1 Then
            strXml = mScript.mainClass(classId).ChildProperties(child2Id)(propName).Value
        Else
            strXml = mScript.mainClass(classId).ChildProperties(child2Id)(propName).ThirdLevelProperties(child3Id)
        End If
        Dim cRes As MatewScript.ContainsCodeEnum = mScript.IsPropertyContainsCode(strXml)
        Dim cd() As CodeTextBox.CodeDataType
        If cRes = MatewScript.ContainsCodeEnum.NOT_CODE OrElse cRes = MatewScript.ContainsCodeEnum.EXECUTABLE_STRING Then
            questEnvironment.codeBoxShadowed.Text = mScript.PrepareStringToPrint(strXml, Nothing, False)
            cd = CopyCodeDataArray(questEnvironment.codeBoxShadowed.codeBox.CodeData)
            questEnvironment.codeBoxShadowed.Text = ""
        Else
            cd = mScript.DeserializeCodeData(strXml) 'переводим его в codeBox.CodeData()
        End If
        Dim eventHTML As String
        'выводим код в браузер
        Dim hCode As HtmlElement
        Dim lstSelectors As SortedList(Of Integer, CodeTextBox.CodeDataType()) = Nothing
        Dim iHeight As Integer = 600, classLW As Integer = mScript.mainClassHash("LW")
        If cRes = MatewScript.ContainsCodeEnum.LONG_TEXT Then
            'получаем высоту для iframe
            Dim strHeight As String = mScript.mainClass(classLW).Properties("Height").Value
            If mScript.IsPropertyContainsCode(strHeight) <> MatewScript.ContainsCodeEnum.NOT_CODE Then
                iHeight = CInt(Screen.PrimaryScreen.WorkingArea.Height * 0.7)
            Else
                strHeight = UnWrapString(strHeight)
                If strHeight.EndsWith("%") Then
                    iHeight = CInt(Screen.PrimaryScreen.WorkingArea.Height * (Val(strHeight) / 100))
                Else
                    iHeight = Val(strHeight)
                End If
            End If
            iHeight = Math.Max(iHeight, 600)

            lstSelectors = mScript.SplitLongTextBySelectors(cd)
            If lstSelectors.Count = 0 Then
                'текст пуст
                eventHTML = ""
            ElseIf lstSelectors.Count = 1 Then
                'селектор только один или текст вообще без селекторов
                eventHTML = mScript.ConvertCodeDataToHTML(lstSelectors.ElementAt(0).Value, False) 'конвертируем в html
            Else
                'в тексте есть более одного селектора
                'получаем текст первого
                eventHTML = mScript.ConvertCodeDataToHTML(lstSelectors.ElementAt(0).Value, False) 'конвертируем в html
                Dim hSelContainer As HtmlElement = hDocument.GetElementById("selectorsContainer")
                If IsNothing(hSelContainer) = False Then
                    'создаем кнопки выбора селекторов
                    For i As Integer = 0 To lstSelectors.Count - 1
                        Dim selectorId As Integer = lstSelectors.ElementAt(i).Key
                        Dim btnText As String = "Селектор #" + selectorId.ToString
                        Dim hButton As HtmlElement = hDocument.CreateElement("INPUT")
                        hButton.SetAttribute("Type", "Button")
                        hButton.SetAttribute("Value", btnText)
                        If i = 0 Then
                            hButton.SetAttribute("ClassName", "SelectorButtonActivated")
                        Else
                            hButton.SetAttribute("ClassName", "SelectorButton")
                        End If
                        Dim strText As String = mScript.ConvertCodeDataToHTML(lstSelectors.ElementAt(i).Value, False) 'конвертируем в html
                        AddHandler hButton.Click, Sub(sender As Object, e2 As HtmlElementEventArgs)
                                                      'клик по кнопке выбора селектора
                                                      Dim fr As HtmlWindow = hDocument.Window.Frames("dataContainer")
                                                      If IsNothing(fr) OrElse IsNothing(fr.Document) Then Return
                                                      Dim hMC As HtmlElement = fr.Document.GetElementById("MainConvas")
                                                      If IsNothing(hMC) Then Return
                                                      hMC.InnerHtml = ""
                                                      ShowDescription(hMC, classId, "Description", child2Id, -1, Nothing, selectorId)

                                                      For j = 0 To hSelContainer.All.Count - 1
                                                          Dim hCurBtn As HtmlElement = hSelContainer.All(j)
                                                          If hCurBtn.TagName <> "INPUT" Then Continue For
                                                          hCurBtn.SetAttribute("ClassName", "SelectorButton")
                                                      Next j
                                                      hButton.SetAttribute("ClassName", "SelectorButtonActivated")
                                                      'Dim h As Integer = Math.Max(fr.Document.Body.ScrollRectangle.Height + 15, 600)
                                                      'hMC.Style = "height:" + (h).ToString + "px;"
                                                      'Dim hFr As HtmlElement = hDocument.GetElementById("dataContainer")
                                                      'If IsNothing(hFr) = False Then hFr.Style = "height:" + (h + 25).ToString + "px"
                                                  End Sub
                        hSelContainer.AppendChild(hButton)
                    Next
                End If
            End If
            'eventHTML = mScript.ConvertCodeDataToHTML(cd, False) 'конвертируем в html
            hCode = hDocument.CreateElement("DIV")
        Else
            eventHTML = mScript.ConvertCodeDataToHTML(cd, True) 'конвертируем в html
            hCode = hDocument.CreateElement("CODE")
        End If
        If cRes <> MatewScript.ContainsCodeEnum.LONG_TEXT Then hCode.InnerHtml = eventHTML

        If hasConvas Then
            If String.IsNullOrWhiteSpace(eventHTML) Then
                elFrame.Style = "display:none"
            Else
                elFrame.Style = "display:block"
            End If

            If cRes <> MatewScript.ContainsCodeEnum.LONG_TEXT Then
                elFrame.InnerHtml = ""
                elFrame.AppendChild(hCode)
            Else
                'iFrame
                'elConvas.SetAttribute("src", questEnvironment.QuestPath + "\Location.html")
                Dim frameWnd As HtmlWindow = Nothing
                Try
                    frameWnd = hDocument.Window.Frames(elFrame.Id)
                Catch ex As Exception
                    Return
                End Try
                abortNavigationEvent = True
                frameWnd.Navigate(questEnvironment.QuestPath + "\Location.html")
                Dim cnt As Integer = 0
                Dim locConvas As HtmlElement = Nothing
                Do While IsNothing(locConvas)
                    Application.DoEvents()
                    locConvas = frameWnd.Document.GetElementById("MainConvas")
                    cnt += 1
                    If cnt > 100000 Then Exit Do
                Loop

                If IsNothing(locConvas) Then
                    MessageBox.Show("Ресурс квеста " + questEnvironment.QuestPath + "\Location.html " + "был изменен. Для корректной работы добавьте внутрь файла строку:" + vbNewLine + _
                                    "<div id='MainConvas'></div>", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return
                End If

                Dim strStyle As New System.Text.StringBuilder
                Dim classL As Integer = mScript.mainClassHash("L")
                'Устанавливаем цвет фона
                Dim bgColor As String = ""
                If child2Id = -1 Then
                    bgColor = mScript.mainClass(classL).Properties("BackColor").Value
                Else
                    bgColor = mScript.mainClass(classL).ChildProperties(child2Id)("BackColor").Value
                End If
                bgColor = UnWrapString(bgColor)
                If String.IsNullOrEmpty(bgColor) = False AndAlso bgColor <> "''" Then
                    If mScript.IsPropertyContainsCode(bgColor) = MatewScript.ContainsCodeEnum.NOT_CODE Then
                        strStyle.Append("background:" & bgColor & ";")
                    End If
                End If
                'Устанавливаем цвет текста
                Dim fColor As String = mScript.mainClass(classLW).Properties("TextColor").Value
                If String.IsNullOrEmpty(fColor) = False AndAlso fColor <> "''" Then
                    If mScript.IsPropertyContainsCode(fColor) = MatewScript.ContainsCodeEnum.NOT_CODE Then
                        fColor = UnWrapString(fColor)
                        strStyle.Append("color:" & fColor & ";")
                    End If
                End If
                'Устанавливаем прилегание
                Dim tAlign As String = mScript.mainClass(classLW).Properties("Align").Value
                If String.IsNullOrEmpty(tAlign) = False AndAlso tAlign <> "''" Then
                    If mScript.IsPropertyContainsCode(tAlign) = MatewScript.ContainsCodeEnum.NOT_CODE Then
                        tAlign = UnWrapString(tAlign)
                        strStyle.Append("text-align:" & tAlign & ";")
                    End If
                End If

                'выводим фоновую картинку (если есть)
                Dim backPicture As String = ""
                If classId = mScript.mainClassHash("L") Then
                    'только для описания локаций
                    If child2Id = -1 Then
                        backPicture = mScript.PrepareStringToPrint(mScript.mainClass(classL).Properties("BackPicture").Value, Nothing, False)
                    Else
                        backPicture = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("BackPicture").Value, Nothing, False)
                    End If
                End If
                If String.IsNullOrEmpty(backPicture) = False Then
                    'картинка указана
                    backPicture = (questEnvironment.QuestPath & "/" & backPicture).Replace("\", "/")
                    If My.Computer.FileSystem.FileExists(backPicture) Then
                        'файл картинки существует
                        strStyle.AppendFormat("background-image: url({0});", "'" + backPicture + "'")
                        Dim propValue As String = ""
                        If child2Id < 0 Then
                            propValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).Properties("BackPicStyle").Value, Nothing, False)
                        ElseIf child3Id < 0 Then
                            propValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("BackPicStyle").Value, Nothing, False)
                        Else
                            propValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("BackPicStyle").ThirdLevelProperties(child3Id), Nothing, False)
                        End If
                        '0 простая загрузка, 1 - заполнить, 2 - масштабировать, 3 - размножить, 4 - размножить по Х, 5 - размножить по Y 
                        strStyle.Append("background-repeat:")
                        Dim propInt As Integer = Val(propValue)
                        Select Case propInt
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
                        If propInt = 0 Then
                            'BackPicPos
                            If child2Id < 0 Then
                                propValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).Properties("BackPicPos").Value, Nothing, False)
                            ElseIf child3Id < 0 Then
                                propValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("BackPicPos").Value, Nothing, False)
                            Else
                                propValue = mScript.PrepareStringToPrint(mScript.mainClass(classId).ChildProperties(child2Id)("BackPicPos").ThirdLevelProperties(child3Id), Nothing, False)
                            End If

                            '0 в левом верхнем углу, 1 слева по центру, 2 в левом нижнем углу, 3 сверху по центру, 4 в центре, 5 снизу по центру, 6 в правом верхнем углу, 7 справа по центру, 8 в правом нижнем углу
                            strStyle.Append("background-position:")
                            propInt = Val(propValue)
                            Select Case propInt
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
                        'locConvas.Style = strStyle.ToString
                    End If
                End If
                If strStyle.Length > 0 Then frameWnd.Document.Body.Style = strStyle.ToString

                Dim selector As Integer = -1
                If IsNothing(lstSelectors) = False AndAlso lstSelectors.Count > 0 Then selector = lstSelectors.Keys(0)
                ShowDescription(locConvas, classId, "Description", child2Id, child3Id, Nothing, selector)

                elFrame.Style = "width:100%;height:" & iHeight & "px;padding:0px;"
                'locConvas.AppendChild(hCode)
                'locConvas.Style += "position:relative;left:0px;padding:0px;margin:0x"

                'Try
                '    'elConvas.Style = "height:" + (frameWnd.Document.Body.ScrollRectangle.Height + 15).ToString + "px;"
                '    elConvas.Style = "max-height:1000px;min-height:400px" '+ (frameWnd.Document.Body.ScrollRectangle.Height + 15).ToString + "px;"
                'Catch ex As Exception
                '    MsgBox("Can't set height")
                '    elConvas.Style = "height:400px;"
                'End Try

            End If
        Else
            elFrame.AppendChild(hCode)
        End If
    End Sub
#End Region

    ''' <summary>
    ''' Совершает переход от дочерних элементов к родителю и наоборот. Отличается от ShowWithClassChanging тем, что может быть совершен к дочерним элементам даже если их пока не существует 
    ''' (отображается пустое дерево)
    ''' </summary>
    ''' <param name="newClassName">Имя класса, к которому осуществляется переход</param>
    ''' <param name="isThirdLevel">Осуществляется ли переход к элементам 3 уровня</param>
    ''' <param name="parentForThirdLevel">Родитель для отображаемых элементов (например, Имя меню для подменю или Имя локации для действий)</param>
    ''' <param name="elementToSelect">Имя элемента, который в конце должен стать активным</param>
    Public Sub ChangeCurrentClass(ByVal newClassName As String, Optional ByVal isThirdLevel As Boolean = False, Optional ByVal parentForThirdLevel As String = "", _
                              Optional elementToSelect As String = "")
        Log.PrintToLog("ChangeCurrentClass: " + newClassName + ", selChildName: " + elementToSelect)

        Dim newClassId As Integer = mScript.mainClassHash(newClassName)
        Dim classLevels As Integer = mScript.mainClass(newClassId).LevelsCount
        Dim tree As TreeView = Nothing
        dictTrees.TryGetValue(newClassId, tree) 'получаем новое дерево

        If String.IsNullOrEmpty(currentClassName) AndAlso String.IsNullOrEmpty(actionsRouter.locationOfCurrentActions) = False Then
            'такое возможно при загрузке квеста, если сохранение было на действиях или локациях
            actionsRouter.SaveActions()
        End If
        'предварительно обновляем дерево для классов 3 уровня и классов, имеющих родителя (кроме Actions, они должны обновиться позже)
        If newClassName <> "A" AndAlso IsNothing(tree) = False AndAlso (parentForThirdLevel.Length > 0 OrElse mScript.mainClass(newClassId).LevelsCount = 2) Then
            'для запуска FillTree установливаем глобальные переменные состояния в новые значения, затем возвращаем их значения обратно
            Dim oldTree As TreeView = currentTreeView
            Dim oldParent As String = currentParentName
            currentTreeView = tree
            currentParentName = parentForThirdLevel
            FillTree(newClassName, tree, chkShowHidden.Checked, Nothing, parentForThirdLevel)
            currentTreeView = oldTree
            currentParentName = oldParent
        End If

        If String.IsNullOrWhiteSpace(elementToSelect) = False Then
            'Указан элемент, к которому надо перейти
            If IsNothing(tree) Then
                MsgBox("Не найдено дерево элементов для класса " + newClassName, MsgBoxStyle.Exclamation)
                Return
            End If

            'заполняем дерево при необходмости (при смене родителя)
            Dim n As TreeNode = Nothing
            If newClassName <> "A" Then
                n = FindItemNodeByText(tree, mScript.PrepareStringToPrint(elementToSelect, Nothing, False))
            End If

            'сохраняем действия если сейчас открыта локация
            'If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False AndAlso cPanelManager.ActivePanel.child2Id > -1 Then actionsRouter.SaveActions(cPanelManager.ActivePanel.child2Id)

            'Если удастся получить к какому элементу переходить, то открываем панель данного элемента и выходим (смена класса произойдет в ShowWithClassChanging)
            If IsNothing(n) = False OrElse newClassName = "A" Then
                If isThirdLevel Then
                    'третий уровень элементов
                    cPanelManager.OpenPanel(n, newClassId, parentForThirdLevel, elementToSelect, parentForThirdLevel, False)
                ElseIf classLevels = 2 Then
                    'второй уровень элементов в 3-уровнем классе
                    cPanelManager.OpenPanel(n, newClassId, elementToSelect, "", "", False)
                Else
                    '2 уровень в 2-уровневом классе (может быть Actions для локаций)
                    Dim shouldFillTree As Boolean = False
                    If IsNothing(n) Then shouldFillTree = True
                    Dim ch As clsPanelManager.clsChildPanel = cPanelManager.OpenPanel(n, newClassId, elementToSelect, "", parentForThirdLevel, shouldFillTree)
                    If IsNothing(n) AndAlso IsNothing(ch) = False Then
                        n = ch.treeNode
                    Else
                        Return
                    End If
                End If
                questEnvironment.EnabledEvents = False
                tree_AfterSelect(currentTreeView, New TreeViewEventArgs(n))
                questEnvironment.EnabledEvents = True
                Return
            End If
        ElseIf isThirdLevel = False Then
            'Элемент, который должен стать активным, не указан. Восстанавливаем последний выбранный в этом классе
            'сохраняем действия если сейчас открыта локация
            'If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False AndAlso cPanelManager.ActivePanel.child2Id > -1 Then actionsRouter.SaveActions(cPanelManager.ActivePanel.child2Id)

            Dim newPanel As clsPanelManager.clsChildPanel = Nothing
            cPanelManager.dictLastPanel.TryGetValue(newClassId, newPanel)
            If IsNothing(newPanel) = False Then
                If IsNothing(mScript.mainClass(newClassId).ChildProperties) = False AndAlso mScript.mainClass(newClassId).ChildProperties.Count > 0 Then
                    'Получили к какому элементу переходить - открываем панель данного элемента и выходим (смена класса произойдет в ShowWithClassChanging)
                    cPanelManager.OpenPanel(newPanel.treeNode, newPanel.classId, newPanel.child2Name, newPanel.child3Name, parentForThirdLevel, True)
                    Return
                End If
            End If
        End If
        If classLevels = 0 Then
            cPanelManager.OpenPanel(Nothing, newClassId, "")
            Return
        End If
        'Здесь мы оказываемся только если нет вкладки для перехода. Осуществляем переход к новому классу без выбора вкладки
        CheckClassButton(currentClassName, newClassId)
        'If currentClassName = "A" Then
        '    actionsRouter.UpdateActions(newClassId, -1, parentForThirdLevel)
        'ElseIf currentClassName = "L" Then
        '    actionsRouter.UpdateActions(newClassId, parentForThirdLevel, -1)
        'End If

        actionsRouter.UpdateActions(newClassId, parentForThirdLevel)

        'Сохраняем действия на локации
        'If currentClassName = "A" Then actionsRouter.SaveActions(currentParentId)

        'сохраняем последний выбранный элемент в старом классе (для возможности возврата именно к нему, если надо будет перейти к старому классу без указания конкретного элемента)
        'для элементов 3 уровня и действий необходимости в этом нет
        Dim prevPanel As clsPanelManager.clsChildPanel = Nothing
        If IsNothing(cPanelManager.ActivePanel) = False AndAlso cPanelManager.ActivePanel.child3Name.Length = 0 AndAlso currentParentName.Length = 0 Then
            prevPanel = cPanelManager.SetLastPanel(currentClassName)
        End If

        If currentClassName = "Variable" OrElse currentClassName = "Function" Then
            currentTreeView.Hide()
        ElseIf mScript.mainClass(mScript.mainClassHash(currentClassName)).LevelsCount > 0 Then
            currentTreeView.Hide()
        End If
        btnShowSettings.Enabled = True
        If isThirdLevel OrElse parentForThirdLevel.Length > 0 Then
            btnMakeHidden.Visible = False
            chkShowHidden.Visible = False
            btnCopyTo.Visible = True
        Else
            btnMakeHidden.Visible = True
            chkShowHidden.Visible = True
            btnCopyTo.Visible = False
        End If
        'установка глобальных переменных состояния
        Dim oldClassName As String = currentClassName
        currentClassName = newClassName
        currentParentName = parentForThirdLevel

        'настройка отображения дерева и разных панелей
        If classLevels > 0 AndAlso dictTrees.ContainsKey(mScript.mainClassHash(currentClassName)) Then
            tree = dictTrees(mScript.mainClassHash(currentClassName))
            currentTreeView = tree
            splitTreeContainer.Show()
            tree.Show()
            SplitOuter.Panel1Collapsed = False
        Else
            currentTreeView = Nothing
            splitTreeContainer.Hide()
            SplitOuter.Panel1Collapsed = True
        End If
        SplitInner.Panel1Collapsed = False
        pnlVariables.Hide()
        'очищаем все данные о предыдущем выбранном элементе
        cPanelManager.HidePanel(cPanelManager.ActivePanel)
        cPanelManager.ActivePanel = Nothing
        If IsNothing(tree) = False Then tree.SelectedNode = Nothing
        btnShowSettings.BackColor = DEFAULT_COLORS.ControlTransparentBackground

        If currentClassName = "A" Then FillTree(currentClassName, tree, True, Nothing, parentForThirdLevel)
        'If currentClassName = "A" Then
        '    'Восстанавливаем действия на локации
        '    actionsRouter.RetreiveActions(currentParentId)
        '    FillTree("A", tree, chkShowHidden.Checked, Nothing, currentParentId)
        'End If
        cPanelManager.UpdateToolButtonsColors(cPanelManager.ActivePanel)        
    End Sub

    ''' <summary>
    ''' Совершает переход к переменным. Отличается от ShowWithClassChanging тем, что может быть совершен к переменным даже если их пока не существует 
    ''' (отображается пустое дерево)
    ''' </summary>
    ''' <param name="elementToSelect">Имя переммой, котораф в конце должна стать активной</param>
    Public Sub ChangeCurrentClassToVariables(Optional elementToSelect As String = "")
        Dim newClassName As String = "Variable"
        Log.PrintToLog("ChangeCurrentClass: " + newClassName + ", selChildName: " + elementToSelect)

        Dim newClassId As Integer = -2
        Dim tree As TreeView = treeVariables 'получаем новое дерево

        If String.IsNullOrWhiteSpace(elementToSelect) = False Then
            'Указан элемент, к которому надо перейти
            Dim n As TreeNode = FindItemNodeByText(tree, elementToSelect)
            'сохраняем действия если сейчас открыта локация
            'If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False AndAlso cPanelManager.ActivePanel.child2Id > -1 Then actionsRouter.SaveActions(cPanelManager.ActivePanel.child2Id)

            'Если удастся получить к какому элементу переходить, то открываем панель данного элемента и выходим (смена класса произойдет в ShowWithClassChanging)
            If IsNothing(n) = False Then
                cPanelManager.OpenPanel(n, newClassId, elementToSelect)
                questEnvironment.EnabledEvents = False
                tree_AfterSelect(currentTreeView, New TreeViewEventArgs(n))
                questEnvironment.EnabledEvents = True
                Return
            End If
        Else
            'Элемент, который должен стать активным, не указан. Восстанавливаем последний выбранный в этом классе
            'сохраняем действия если сейчас открыта локация
            'If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False AndAlso cPanelManager.ActivePanel.child2Id > -1 Then actionsRouter.SaveActions(cPanelManager.ActivePanel.child2Id)
            Dim newPanel As clsPanelManager.clsChildPanel = Nothing
            cPanelManager.dictLastPanel.TryGetValue(newClassId, newPanel)
            If IsNothing(newPanel) = False Then
                If IsNothing(mScript.csPublicVariables.lstVariables) = False AndAlso mScript.csPublicVariables.lstVariables.Count > 0 Then
                    'Получили к какому элементу переходить - открываем панель данного элемента и выходим (смена класса произойдет в ShowWithClassChanging)
                    cPanelManager.OpenPanel(newPanel.treeNode, newPanel.classId, newPanel.child2Name)
                    Return
                End If
            End If
        End If
        'Здесь мы оказываемся только если нет вкладки для перехода. Осуществляем переход к новому классу без выбора вкладки
        CheckClassButton(currentClassName, newClassId)

        'Сохраняем действия на локации
        'If currentClassName = "A" Then actionsRouter.SaveActions(currentParentId)
        If currentClassName = "A" OrElse currentClassName = "L" Then actionsRouter.SaveActions()

        'сохраняем последний выбранный элемент в старом классе (для возможности возврата именно к нему, если надо будет перейти к старому классу без указания конкретного элемента)
        'для элементов 3 уровня и действий необходимости в этом нет
        Dim prevPanel As clsPanelManager.clsChildPanel = Nothing
        If IsNothing(cPanelManager.ActivePanel) = False AndAlso cPanelManager.ActivePanel.child3Name.Length = 0 AndAlso currentParentName.Length = 0 Then
            prevPanel = cPanelManager.SetLastPanel(currentClassName)
        End If

        If IsNothing(currentTreeView) = False Then currentTreeView.Hide()
        'установка глобальных переменных состояния
        Dim oldClassName As String = currentClassName
        currentClassName = newClassName
        currentParentName = ""

        'настройка отображения дерева и разных панелей
        btnShowSettings.Enabled = False
        btnMakeHidden.Visible = True
        chkShowHidden.Visible = True
        btnCopyTo.Visible = False
        tree = treeVariables
        currentTreeView = tree
        splitTreeContainer.Show()
        tree.Show()
        SplitOuter.Panel1Collapsed = False
        SplitInner.Panel1Collapsed = True
        pnlVariables.Hide()

        'очищаем все данные о предыдущем выбранном элементе
        cPanelManager.HidePanel(cPanelManager.ActivePanel)
        cPanelManager.ActivePanel = Nothing
        If IsNothing(tree) = False Then tree.SelectedNode = Nothing
        btnShowSettings.BackColor = DEFAULT_COLORS.ControlTransparentBackground
        cPanelManager.UpdateToolButtonsColors(cPanelManager.ActivePanel)
    End Sub

    ''' <summary>
    ''' Совершает переход к функциям. Отличается от ShowWithClassChanging тем, что может быть совершен к функциям даже если их пока не существует 
    ''' (отображается пустое дерево)
    ''' </summary>
    ''' <param name="elementToSelect">Имя переммой, котораф в конце должна стать активной</param>
    Public Sub ChangeCurrentClassToFunctions(Optional elementToSelect As String = "")
        Dim newClassName As String = "Function"
        Log.PrintToLog("ChangeCurrentClass: " + newClassName + ", selChildName: " + elementToSelect)

        Dim newClassId As Integer = -3
        Dim tree As TreeView = treeFunctions  'получаем новое дерево

        If String.IsNullOrWhiteSpace(elementToSelect) = False Then
            'Указан элемент, к которому надо перейти
            Dim n As TreeNode = FindItemNodeByText(tree, elementToSelect)
            'сохраняем действия если сейчас открыта локация
            'If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False AndAlso cPanelManager.ActivePanel.child2Id > -1 Then actionsRouter.SaveActions(cPanelManager.ActivePanel.child2Id)

            'Если удастся получить к какому элементу переходить, то открываем панель данного элемента и выходим (смена класса произойдет в ShowWithClassChanging)
            If IsNothing(n) = False Then
                cPanelManager.OpenPanel(n, newClassId, elementToSelect)
                questEnvironment.EnabledEvents = False
                tree_AfterSelect(currentTreeView, New TreeViewEventArgs(n))
                questEnvironment.EnabledEvents = True
                Return
            End If
        Else
            'Элемент, который должен стать активным, не указан. Восстанавливаем последний выбранный в этом классе
            'сохраняем действия если сейчас открыта локация
            'If currentClassName = "L" AndAlso IsNothing(cPanelManager.ActivePanel) = False AndAlso cPanelManager.ActivePanel.child2Id > -1 Then actionsRouter.SaveActions(cPanelManager.ActivePanel.child2Id)
            Dim newPanel As clsPanelManager.clsChildPanel = Nothing
            cPanelManager.dictLastPanel.TryGetValue(newClassId, newPanel)
            If IsNothing(newPanel) = False Then
                If IsNothing(mScript.functionsHash) = False AndAlso mScript.functionsHash.Count > 0 Then
                    'Получили к какому элементу переходить - открываем панель данного элемента и выходим (смена класса произойдет в ShowWithClassChanging)
                    cPanelManager.OpenPanel(newPanel.treeNode, newPanel.classId, newPanel.child2Name)
                    Return
                End If
            End If
        End If
        'Здесь мы оказываемся только если нет вкладки для перехода. Осуществляем переход к новому классу без выбора вкладки
        CheckClassButton(currentClassName, newClassId)
        'Сохраняем действия на локации
        'If currentClassName = "A" Then actionsRouter.SaveActions(currentParentId)
        If currentClassName = "A" OrElse currentClassName = "L" Then actionsRouter.SaveActions()

        'сохраняем последний выбранный элемент в старом классе (для возможности возврата именно к нему, если надо будет перейти к старому классу без указания конкретного элемента)
        'для элементов 3 уровня и действий необходимости в этом нет
        Dim prevPanel As clsPanelManager.clsChildPanel = Nothing
        If IsNothing(cPanelManager.ActivePanel) = False AndAlso cPanelManager.ActivePanel.child3Name.Length = 0 AndAlso currentParentName.Length = 0 Then
            prevPanel = cPanelManager.SetLastPanel(currentClassName)
        End If

        If IsNothing(currentTreeView) = False Then currentTreeView.Hide()
        'установка глобальных переменных состояния
        Dim oldClassName As String = currentClassName
        currentClassName = newClassName
        currentParentName = ""
        'настройка отображения дерева и разных панелей
        btnShowSettings.Enabled = False
        btnMakeHidden.Visible = True
        chkShowHidden.Visible = True
        btnCopyTo.Visible = False
        tree = treeFunctions
        currentTreeView = tree
        splitTreeContainer.Show()
        tree.Show()
        SplitOuter.Panel1Collapsed = False
        SplitInner.Panel1Collapsed = True
        pnlVariables.Hide()
        'очищаем все данные о предыдущем выбранном элементе
        cPanelManager.HidePanel(cPanelManager.ActivePanel)
        cPanelManager.ActivePanel = Nothing
        If IsNothing(tree) = False Then tree.SelectedNode = Nothing
        btnShowSettings.BackColor = DEFAULT_COLORS.ControlTransparentBackground
        cPanelManager.UpdateToolButtonsColors(cPanelManager.ActivePanel)
    End Sub

    ''' <summary>
    ''' Выделяет кнопнку на главной панели, соответствующую выбранному классу, и снимает выделения со старой
    ''' </summary>
    ''' <param name="oldClassName">Имя старого класса</param>
    ''' <param name="newClassId">Id нового класса</param>
    ''' <remarks></remarks>
    Public Sub CheckClassButton(ByVal oldClassName As String, ByVal newClassId As Integer)
        'Получаем Id старого класса
        Dim oldClassId As Integer = -1
        If oldClassName = "Variable" Then
            oldClassId = -2
        ElseIf oldClassName = "Function" Then
            oldClassId = -3
        Else
            If String.IsNullOrEmpty(oldClassName) Then
                oldClassId = -1
            Else
                oldClassId = mScript.mainClassHash(oldClassName)
            End If
        End If
        'Для действий нет отдельной кнопки - активной должна быть кнопка локаций
        Dim actClassId As Integer = mScript.mainClassHash("A")
        If oldClassId = actClassId Then
            oldClassId = mScript.mainClassHash("L")
        End If
        If newClassId = actClassId Then
            newClassId = mScript.mainClassHash("L")
        End If
        If oldClassId = newClassId Then Return

        For i As Integer = 0 To ToolStripMain.Items.Count - 1
            Dim itm As ToolStripItem = ToolStripMain.Items(i)
            Dim tName As String = itm.GetType.Name
            If tName = "ToolStripButton" Then
                Dim btn As ToolStripButton = itm
                Dim strTag As String = itm.Tag
                If String.IsNullOrEmpty(strTag) Then Continue For
                'Это кнопнка с тэгом. Получаем класс из тэга
                Dim btnClassId As Integer = -1
                If strTag = "Variable" Then
                    btnClassId = -2
                ElseIf strTag = "Function" Then
                    btnClassId = -3
                Else
                    If mScript.mainClassHash.TryGetValue(strTag, btnClassId) = False Then Continue For
                End If
                If btnClassId = oldClassId Then
                    btn.Checked = False
                ElseIf btnClassId = newClassId Then
                    btn.Checked = True
                End If
            ElseIf tName = "ToolStripDropDownButton" Then
                'Это выпадающий список. Продим все его элементы
                Dim cb As ToolStripDropDownButton = itm
                For j As Integer = 0 To cb.DropDownItems.Count - 1
                    itm = cb.DropDownItems(j)
                    If itm.GetType.Name = "ToolStripMenuItem" Then
                        Dim strTag As String = itm.Tag
                        If String.IsNullOrEmpty(strTag) Then Continue For
                        'Это меню с тэгом. Получаем класс из тэга
                        Dim btnClassId As Integer = -1
                        If mScript.mainClassHash.TryGetValue(strTag, btnClassId) = False Then Continue For
                        Dim tsmi As ToolStripMenuItem = itm
                        If btnClassId = oldClassId Then
                            tsmi.Checked = False
                        ElseIf btnClassId = newClassId Then
                            tsmi.Checked = True
                        End If
                    End If
                Next j
            End If
        Next i
    End Sub

    Public Sub tsbChangeClass(sender As Object, e As EventArgs) Handles tsbQ.Click, tsbL.Click, tsbM.Click, tsbObj.Click, tsbT.Click, tsbMap.Click, tsbMg.Click, tsbAb.Click, tsbMed.Click, tsbBat.Click, _
        tsbHer.Click, tsmiLW.Click, tsmiCm.Click, tsmiDW.Click, tsmiOW.Click, tsmiAW.Click, tsmiMgW.Click, tsbString.Click, tsbMath.Click, tsbArr.Click, tsbFile.Click, tsbCode.Click, tsbDate.Click, tsbArmy.Click
        If codeBoxPanel.Visible Then
            'Validating кодбокса
            codeBox.codeBox.CheckTextForSyntaxErrors(codeBox.codeBox)
            Dim ee As New System.ComponentModel.CancelEventArgs
            Call codeBox_Validating(codeBox, ee)
            If ee.Cancel Then Return
            codeBoxChangeOwner(Nothing)
            'ElseIf IsNothing(cPanelManager.ActivePanel) = False AndAlso IsNothing(cPanelManager.ActivePanel.ActiveControl) = False AndAlso cPanelManager.ActivePanel.ActiveControl.Focused Then
            '    'Validating свойства
            '    Dim tName As String
            '    tName = cPanelManager.ActivePanel.ActiveControl.GetType.Name
            '    If tName <> "ButtonEx" Then
            '        Dim ee As New System.ComponentModel.CancelEventArgs
            '        'Call del_propertyValidating(cPanelManager.ActivePanel.ActiveControl, ee)
            '        If ee.Cancel Then Return
            '    End If
        End If


        ChangeCurrentClass(sender.Tag.ToString)
    End Sub

    Private Sub tsbFunc_Click(sender As Object, e As EventArgs) Handles tsbFunc.Click
        If codeBoxPanel.Visible Then
            'Validating кодбокса
            codeBox.codeBox.CheckTextForSyntaxErrors(codeBox.codeBox)
            Dim ee As New System.ComponentModel.CancelEventArgs
            Call codeBox_Validating(codeBox, ee)
            If ee.Cancel Then Return
            codeBoxChangeOwner(Nothing)
        End If

        ChangeCurrentClassToFunctions()
    End Sub

    Private Sub tsbVar_Click(sender As Object, e As EventArgs) Handles tsbVar.Click
        If codeBoxPanel.Visible Then
            'Validating кодбокса
            codeBox.codeBox.CheckTextForSyntaxErrors(codeBox.codeBox)
            Dim ee As New System.ComponentModel.CancelEventArgs
            Call codeBox_Validating(codeBox, ee)
            If ee.Cancel Then Return
            codeBoxChangeOwner(Nothing)
        End If

        ChangeCurrentClassToVariables()
    End Sub

    Private fsWatcherFullPath As New List(Of String)
    Private Sub fsWatcher_Created(sender As Object, e As IO.FileSystemEventArgs) Handles fsWatcher.Created
        If Not e.FullPath.ToLower.StartsWith(questEnvironment.QuestPath.ToLower) Then Return
        'cListManager.UpdateFileLists()
        If timUpdateFiles.Enabled = False Then timUpdateFiles.Start()
        fsWatcherFullPath.Add(e.FullPath)
    End Sub

    Private Sub fsWatcher_Deleted(sender As Object, e As IO.FileSystemEventArgs) Handles fsWatcher.Deleted
        If Not e.FullPath.ToLower.StartsWith(questEnvironment.QuestPath.ToLower) Then Return
        'cListManager.UpdateFileLists()
        If timUpdateFiles.Enabled = False Then timUpdateFiles.Start()
        fsWatcherFullPath.Add(e.FullPath)
    End Sub

    Private Sub fsWatcher_Renamed(sender As Object, e As IO.RenamedEventArgs) Handles fsWatcher.Renamed
        If Not e.FullPath.ToLower.StartsWith(questEnvironment.QuestPath.ToLower) Then Return
        'cListManager.UpdateFileLists()
        If timUpdateFiles.Enabled = False Then timUpdateFiles.Start()
        fsWatcherFullPath.Add(e.FullPath)
    End Sub

    Private Sub timUpdateFiles_Tick(sender As Object, e As EventArgs) Handles timUpdateFiles.Tick
        timUpdateFiles.Stop()
        cListManager.UpdateFileLists()
        For i As Integer = 0 To fsWatcherFullPath.Count - 1
            If fsWatcherFullPath(i).StartsWith(questEnvironment.QuestPath + "\img\tree_icons", StringComparison.CurrentCultureIgnoreCase) Then
                'Обновляем иконки деревьев при изменении папки tree_icons
                LoadIconsForTrees()
                CreateIconsForTreeMenus()
                UpdateTreesAfterIconsChanging()
                Exit For
            End If
        Next
        fsWatcherFullPath.Clear()
    End Sub

    ''' <summary>
    ''' После изменения списка иконок возникает глюк - во всех деревьях иконка выбранного узла становится другой. Исправляем это
    ''' </summary>
    Private Sub UpdateTreesAfterIconsChanging()
        If IsNothing(dictTrees) OrElse dictTrees.Count = 0 Then Return
        For tId As Integer = 0 To dictTrees.Count - 1
            'обновляем деревья элементов
            Dim tree As TreeView = dictTrees.ElementAt(tId).Value
            If tree.Nodes.Count = 0 Then Continue For
            For nId As Integer = 0 To tree.Nodes.Count - 1
                Dim n As TreeNode = tree.Nodes(nId)
                n.SelectedImageKey = n.ImageKey
                If n.Nodes.Count > 0 Then UpdateTreeNodeAfterIconChangingRecursion(n)
            Next nId
            'обновляем дерево переменных
            tree = treeVariables
            If tree.Nodes.Count > 0 Then
                For nId As Integer = 0 To tree.Nodes.Count - 1
                    Dim n As TreeNode = tree.Nodes(nId)
                    n.SelectedImageKey = n.ImageKey
                    If n.Nodes.Count > 0 Then UpdateTreeNodeAfterIconChangingRecursion(n)
                Next nId
            End If
            'обновляем дерево функций
            tree = treeFunctions
            If tree.Nodes.Count > 0 Then
                For nId As Integer = 0 To tree.Nodes.Count - 1
                    Dim n As TreeNode = tree.Nodes(nId)
                    n.SelectedImageKey = n.ImageKey
                    If n.Nodes.Count > 0 Then UpdateTreeNodeAfterIconChangingRecursion(n)
                Next nId
            End If
        Next tId
    End Sub

    ''' <summary>
    ''' Рекурсивная процедура, вызыватеся только из UpdateTreesAfterIconsChanging
    ''' </summary>
    ''' <param name="n">Текущий узел рекурсии</param>
    Private Sub UpdateTreeNodeAfterIconChangingRecursion(ByRef n As TreeNode)
        For i As Integer = 0 To n.Nodes.Count - 1
            Dim childN As TreeNode = n.Nodes(i)
            childN.SelectedImageKey = childN.ImageKey
            If n.Nodes.Count > 0 Then UpdateTreeNodeAfterIconChangingRecursion(childN)
        Next
    End Sub

    Private Sub pnlVariables_ClientSizeChanged(sender As Object, e As EventArgs) Handles pnlVariables.ClientSizeChanged
        txtVariableDescription.Width = pnlVariables.ClientSize.Width - txtVariableDescription.Left - questEnvironment.defPaddingLeft
        dgwVariables.Location = New Point(lblVariableDescription.Left, txtVariableDescription.Bottom + questEnvironment.defPaddingTop)
        dgwVariables.Size = New Size(pnlVariables.ClientSize.Width - dgwVariables.Left - questEnvironment.defPaddingLeft, pnlVariables.ClientSize.Height - dgwVariables.Top - questEnvironment.defPaddingTop)
    End Sub

    Private Sub dgwVariables_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles dgwVariables.Validating
        If IsNothing(sender.Tag) Then Return
        lblVariableError.Text = ""
        Dim ch As clsPanelManager.clsChildPanel = sender.Tag
        Dim v As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables(ch.child2Name)
        Dim vName As String = ch.child2Name
        If IsNothing(v.lstSingatures) = False Then v.lstSingatures.Clear()
        If dgwVariables.Rows.Count = 0 Then
            Erase v.arrValues
            mScript.csPublicVariables.lstVariables(vName) = v
            Return
        End If

        ReDim v.arrValues(dgwVariables.Rows.Count - 1)
        Dim ub As Integer = -1
        For i As Integer = 0 To dgwVariables.Rows.Count - 1
            Dim res As String = dgwVariables.Item(1, i).Value
            Dim sign As String = dgwVariables.Item(0, i).Value
            If String.IsNullOrEmpty(res) Then res = ""
            If String.IsNullOrEmpty(sign) Then sign = ""
            If res.Length = 0 AndAlso sign.Length = 0 AndAlso i = dgwVariables.RowCount - 1 Then Exit For
            ub += 1
            If res.ToLower = "true" Then res = "True"
            If res.ToLower = "false" Then res = "False"
            If res.Length > 0 Then v.arrValues(i) = WrapString(res)


            If String.IsNullOrEmpty(sign) = False Then
                If IsNothing(v.lstSingatures) Then
                    v.lstSingatures = New SortedList(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)
                Else
                    If v.lstSingatures.ContainsKey(sign) Then
                        lblVariableError.Text = "Ошибка при заполнении массива переменной. Сигнатура " + Chr(34) + sign + Chr(34) + " встречается несколько раз. Значения должны быть уникальными!"
                        e.Cancel = True
                        Return
                    End If
                End If
                v.lstSingatures.Add(sign, i)
            End If
        Next
        If ub = -1 Then
            Erase v.arrValues
        ElseIf ub <> v.arrValues.Count - 1 Then
            ReDim Preserve v.arrValues(ub)
        End If

    End Sub

    Private Sub txtVariableDescription_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtVariableDescription.Validating
        If IsNothing(sender.Tag) Then Return
        Dim ch As clsPanelManager.clsChildPanel = sender.Tag
        Dim v As cVariable.variableEditorInfoType = mScript.csPublicVariables.lstVariables(ch.child2Name)
        v.Description = sender.Text
    End Sub

    Private Sub tsbAddUserClass_Click(sender As Object, e As EventArgs) Handles tsbAddUserClass.Click
        Dim strNames As String
        Dim levels As Integer = 3
        Dim img As Image
        Dim helpFile As String
        Dim defProperty As String
        With dlgNewClass
            .PrepareForNewClass()
            Dim res As DialogResult = dlgNewClass.ShowDialog(Me)
            If res = Windows.Forms.DialogResult.Cancel Then Return
            strNames = .txtNewClassName.Text.Trim
            levels = .nudClassLevels.Value
            img = .classIcon
            helpFile = .txtClassHelpFile.Text
            defProperty = .cmbClassDefProperty.Text
        End With

        'Создаем новый класс
        'Получаем Id класса
        Dim classId As Integer = mScript.mainClass.Length
        'Заполняем arrNames именами класса
        Dim arrNames() As String = Split(strNames, ",")
        For i As Integer = 0 To arrNames.GetUpperBound(0)
            arrNames(i) = arrNames(i).Trim
            If arrNames(i).Length = 0 Then
                MsgBox("Ошибка при заполнении имен класса. Вероятно, лишняя запятая.", MsgBoxStyle.Exclamation)
                Return
            End If
            If mScript.mainClassHash.ContainsKey(arrNames(i)) Then
                MsgBox("Имя " + arrNames(i) + " уже существует.", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
        Next
        'сохраняем имена класса в структуре mainClass
        ReDim Preserve mScript.mainClass(classId)
        mScript.mainClass(classId) = New MatewScript.MainClassType
        mScript.mainClass(classId).Names = arrNames
        mScript.mainClass(classId).HelpFile = helpFile
        mScript.mainClass(classId).DefaultProperty = defProperty
        mScript.mainClass(classId).UserAdded = True
        'Сохраняем файл картинки в папке квеста
        If IsNothing(img) = False Then
            If Not My.Computer.FileSystem.DirectoryExists(questEnvironment.QuestPath + "\img\classMenu") Then
                My.Computer.FileSystem.CreateDirectory(questEnvironment.QuestPath + "\img\classMenu")
            End If
            img.Save(questEnvironment.QuestPath + "\img\classMenu\" + mScript.mainClass(classId).Names(0) + ".png", System.Drawing.Imaging.ImageFormat.Png)
        End If

        'заполняем данные о новом классе. Классам 2 и 3 уровня добавляем набор базовых функций
        mScript.mainClass(classId).LevelsCount = levels - 1
        If levels = 1 Then
            mScript.mainClass(classId).Functions = New SortedList(Of String, MatewScript.PropertiesInfoType)(mScript.basicFunctionsHashLevel1)
        ElseIf levels = 2 Then
            mScript.mainClass(classId).Functions = New SortedList(Of String, MatewScript.PropertiesInfoType)(mScript.basicFunctionsHashLevel2)
        Else
            mScript.mainClass(classId).Functions = New SortedList(Of String, MatewScript.PropertiesInfoType)(mScript.basicFunctionsHashLevel3)
        End If
        mScript.mainClass(classId).Properties = New SortedList(Of String, MatewScript.PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
        'Классам 2 и 3 уровня создаем свойство Name
        If levels > 1 Then mScript.mainClass(classId).Properties.Add("Name", New MatewScript.PropertiesInfoType With {.Value = "'default'", .Description = "Имя элемента", .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR, .editorIndex = 0})
        mScript.mainClass(classId).Properties.Add("Group", New MatewScript.PropertiesInfoType With {.Value = "", .Description = "[СЛУЖЕБНОЕ] Для хранения группы элемента", .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL, .editorIndex = 1})
        mScript.mainClass(classId).Properties.Add("Icon", New MatewScript.PropertiesInfoType With {.Value = "", .Description = "[СЛУЖЕБНОЕ] Для хранения иконки элемента", .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL, .editorIndex = 2})
        mScript.MakeMainClassHash() 'обновляем хэш с именами классов
        mScript.UpdateBasicFunctionsParamsWhichIsElements(classId)
        mScript.FillFuncAndPropHash()

        'добавляем кнопку
        If IsNothing(img) Then img = Image.FromFile(Application.StartupPath + "\src\img\classMenu\U.png")
        Dim tsb As New ToolStripButton(mScript.mainClass(classId).Names.Last, img, AddressOf tsbChangeClass) With _
            {.Tag = mScript.mainClass(classId).Names(0), .DisplayStyle = ToolStripItemDisplayStyle.Image, .ImageScaling = ToolStripItemImageScaling.None}

        ToolStripMain.Items.Insert(ToolStripMain.Items.IndexOf(tsbAddUserClass), tsb)

        AppendTree(classId)
        cGroups.dictGroups.Add(mScript.mainClass(classId).Names(0), New List(Of clsGroups.clsGroupInfo))
    End Sub

    Private Sub tsmiFindGlobal_Click(sender As Object, e As EventArgs) Handles tsmiFindGlobal.Click, btnFindClobal.Click
        If frmSeek.Visible Then Return
        frmSeek.Show(Me)
    End Sub

    Private Sub tsmiDefVars_Click(sender As Object, e As EventArgs) Handles tsmiDefVars.Click, tsbDefVars.Click
        'Выводит в кодбокс параметры свойств-событий в виде myVar = Param[x] //Описание. После этого выводит возвращаемые значения
        'Для функций выводит закомментированный список параметров.
        If codeBoxPanel.Visible = False Then Return
        If IsNothing(codeBox.Tag) Then Return

        If trakingEventState = trackingcodeEnum.EVENT_BEFORE Then
            'событие перед изменением свойства
            Dim ch As clsPanelManager.clsChildPanel = codeBox.Tag.childPanel
            Dim txt As New System.Text.StringBuilder
            Dim levels As Integer = mScript.mainClass(ch.classId).LevelsCount
            txt.AppendLine("//Событие до изменения свойства. Возникает каждый раз перед изменением значения данного свойства.")
            If levels = 0 Then
                txt.AppendLine("//Param[0] для данного свойства не актуально")
            Else
                txt.AppendLine("itmId = Param[0] //Id элемента 2 порядка, которому устанавливается данное свойство. Если устанавливается свойство по умолчанию, то -1")
            End If
            If levels < 2 Then
                txt.AppendLine("//Param[1] для данного свойства не актуально")
            Else
                txt.AppendLine("childId = Param[1] //Id элемента 3 порядка, которому устанавливается данное свойство. Если устанавливается свойство по умолчанию или свойство элемента 2 порядка, то -1")
            End If
            txt.AppendLine("val = Param[2] //Устанавливаемое значение свойству. Для передачи свойству иного значения следует изменить Param[2]")
            txt.AppendLine("//Для отмены изменения значения свойства событие должно вернуть False")

            codeBox.codeBox.SelectedText = txt.ToString
            txt.Clear()
            Return
        ElseIf trakingEventState = trackingcodeEnum.EVENT_AFTER Then
            'событие после изменения свойства
            Dim ch As clsPanelManager.clsChildPanel = codeBox.Tag.childPanel
            Dim txt As New System.Text.StringBuilder
            Dim levels As Integer = mScript.mainClass(ch.classId).LevelsCount
            txt.AppendLine("//Событие после свойства. Возникает каждый раз, когда значения данного свойства изменяется.")
            If levels = 0 Then
                txt.AppendLine("//Param[0] для данного свойства не актуально")
            Else
                txt.AppendLine("itmId = Param[0] //Id элемента 2 порядка, которому устанавливается данное свойство. Если устанавливается свойство по умолчанию, то -1")
            End If
            If levels < 2 Then
                txt.AppendLine("//Param[1] для данного свойства не актуально")
            Else
                txt.AppendLine("childId = Param[1] //Id элемента 3 порядка, которому устанавливается данное свойство. Если устанавливается свойство по умолчанию или свойство элемента 2 порядка, то -1")
            End If

            codeBox.codeBox.SelectedText = txt.ToString
            txt.Clear()
            Return
        End If

        Select Case codeBox.Tag.GetType.Name
            Case "ButtonEx", "TextBoxEx", "ComboBoxEx"
                'Dim btn As ButtonEx = codeBox.Tag
                Dim ch As clsPanelManager.clsChildPanel = codeBox.Tag.childPanel
                Dim elName As String = codeBox.Tag.Name
                Dim p As MatewScript.PropertiesInfoType = Nothing
                If codeBox.Tag.IsFunctionButton Then
                    'Это функция
                    If mScript.mainClass(ch.classId).Functions.TryGetValue(elName, p) = False Then Return
                    If IsNothing(p.params) OrElse p.params.Count = 0 Then
                        MessageBox.Show("Функция не имеет принимаемых параметров.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If
                    Dim sb As New System.Text.StringBuilder
                    For i As Integer = 0 To p.params.Count - 1
                        sb.AppendLine("//Param[" + i.ToString + "] - " + p.params(i).Name + ". " + p.params(i).Description)
                    Next i
                    If sb.Length > 0 Then
                        codeBox.codeBox.SelectedText = sb.ToString
                        sb.Clear()
                    End If
                Else
                    'Это свойство
                    If mScript.mainClass(ch.classId).Properties.TryGetValue(elName, p) = False Then Return
                    If p.returnType <> MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                        MessageBox.Show("Данное свойство не является событием.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If
                    If IsNothing(p.params) OrElse p.params.Count = 0 Then
                        MessageBox.Show("Событие не имеет принимаемых параметров.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If
                    Dim sb As New System.Text.StringBuilder
                    Dim isArray As Boolean
                    Dim level1 As Boolean, level2 As Boolean, level3 As Boolean
                    Dim returnFirstly As Boolean = True
                    For i As Integer = 0 To p.params.Count - 1
                        Dim scores As Integer = p.params(i).Type
                        If scores >= 16 Then
                            'Массив параметров
                            scores -= 16
                            isArray = True
                        End If
                        If scores >= 8 Then
                            'Здесь и свсе остальные - Return
                            If returnFirstly Then
                                returnFirstly = False
                                sb.AppendLine("//Возвращает :")
                            End If
                            sb.AppendLine("// " + p.params(i).Name + ". " + p.params(i).Description)
                            Continue For
                        End If
                        If scores >= 4 Then
                            level3 = True
                            scores -= 4
                        Else
                            level3 = False
                        End If
                        If scores >= 2 Then
                            level2 = True
                            scores -= 2
                        Else
                            level2 = False
                        End If
                        If scores = 1 Then
                            level1 = True
                        Else
                            level1 = False
                        End If
                        If (ch.child2Name.Length = 0 AndAlso level1) OrElse (ch.child3Name.Length = 0 AndAlso level2) OrElse (ch.child3Name.Length > 0 AndAlso level3) Then
                            If isArray Then
                                sb.AppendLine("//Param[" + i.ToString + ", " + (i + 1).ToString + ", ..., x] - " + p.params(i).Description)
                            Else
                                sb.AppendLine(p.params(i).Name + " = Param[" + i.ToString + "] //" + p.params(i).Description)
                            End If
                        End If
                    Next
                    If sb.Length > 0 Then
                        codeBox.codeBox.SelectedText = sb.ToString
                        sb.Clear()
                    End If
                End If
        End Select
    End Sub

    Private Sub tsmiRestoreElement_Click(sender As Object, e As EventArgs) Handles tsmiRestoreElement.Click
        If removedObjects.FillListBoxWithItems(dlgRestoreElements.lstElements) = False Then
            MessageBox.Show("Список удаленных элементов пуст.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        dlgRestoreElements.ShowDialog(Me)
    End Sub

    Private Sub tsmiCopy_Click(sender As Object, e As EventArgs) Handles tsmiCopy.Click
        If codeBoxPanel.Visible AndAlso WBhelp.Visible = False Then
            codeBox.mnuCopy.PerformClick()
            Return
        ElseIf WBhelp.Focused Then
            hDocument.ExecCommand("Copy", Nothing, Nothing)
        ElseIf IsNothing(cPanelManager.ActivePanel) = False AndAlso IsNothing(cPanelManager.ActivePanel.ActiveControl) = False Then
            CodeTextBox.SendMessage(cPanelManager.ActivePanel.ActiveControl.Handle, &H301, 0, 0)
        End If

        'Clipboard.SetText(c.codeBox.SelectedText)

    End Sub

    Private Sub tsmiPaste_Click(sender As Object, e As EventArgs) Handles tsmiPaste.Click
        If codeBoxPanel.Visible AndAlso WBhelp.Visible = False Then
            codeBox.mnuPaste.PerformClick()
            Return
        ElseIf IsNothing(cPanelManager.ActivePanel) = False AndAlso IsNothing(cPanelManager.ActivePanel.ActiveControl) = False Then
            CodeTextBox.SendMessage(cPanelManager.ActivePanel.ActiveControl.Handle, &H302, 0, 0)
        End If
    End Sub

    Private Sub tsmiCut_Click(sender As Object, e As EventArgs) Handles tsmiCut.Click
        If codeBoxPanel.Visible AndAlso WBhelp.Visible = False Then
            codeBox.mnuCut.PerformClick()
            Return
        ElseIf IsNothing(cPanelManager.ActivePanel) = False AndAlso IsNothing(cPanelManager.ActivePanel.ActiveControl) = False Then
            CodeTextBox.SendMessage(cPanelManager.ActivePanel.ActiveControl.Handle, &H300, 0, 0)
        End If
    End Sub

    Private Sub tsmiShowAdditionalClasses_Click(sender As Object, e As EventArgs) Handles tsmiShowAdditionalClasses.Click
        tsmiShowAdditionalClasses.Checked = Not tsmiShowAdditionalClasses.Checked
        Dim blnShow As Boolean = tsmiShowAdditionalClasses.Checked
        tsbString.Visible = blnShow
        tsbDate.Visible = blnShow
        tsbMath.Visible = blnShow
        tsbArr.Visible = blnShow
        tsbCode.Visible = blnShow
        tsbFile.Visible = blnShow
    End Sub

    Private Sub tsmiWrap_Click(sender As Object, e As EventArgs) Handles tsmiWrap.Click, tsbWrap.Click
        'Заворачивает выделенный текст в кавычки
        Dim rtb As RichTextBox = codeBox.codeBox
        If rtb.SelectionLength = 0 Then
            Dim sStart As Integer = rtb.SelectionStart
            rtb.SelectedText = "''"
            rtb.SelectionStart = sStart + 1
        Else
            Dim sStart As Integer = rtb.SelectionStart
            Dim txt As String = "'" & rtb.SelectedText.Replace("/'", Chr(1)).Replace("'", "/'").Replace(Chr(1), "/'") & "'"
            rtb.SelectedText = txt
            Application.DoEvents()
            rtb.Select(sStart, txt.Length)
        End If
    End Sub

    Private Sub tsmiComment_Click(sender As Object, e As EventArgs) Handles tsmiComment.Click, tsbComment.Click
        'Вставляет символы комментария // или <!-- -->
        Dim rtb As RichTextBox = codeBox.codeBox
        Dim curline As Integer = rtb.GetLineFromCharIndex(rtb.SelectionStart)
        Dim isExec As CodeTextBox.ExecBlockEnum = CodeTextBox.ExecBlockEnum.NO_EXEC
        Dim isInText As CodeTextBox.TextBlockEnum = codeBox.codeBox.IsInTextBlock(rtb, curline, isExec)

        If isInText Then
            'мы в блоке текста. Комментарии <!-- -->
            'просто вставляем берем выделенный текст в комментарии
            Dim sStart As Integer = rtb.SelectionStart
            Dim sLength As Integer = rtb.SelectionLength
            rtb.SelectedText = "<!--" + rtb.SelectedText + "-->"
            'rtb.SelectionStart = sStart + 4
            rtb.Select(sStart + 4, sLength)
        Else
            'мы в блоке скрипта. Комментарии //
            If rtb.SelectionLength = 0 Then
                'если выделения нет, то ставим // в начале строки
                Dim sStart As Integer = rtb.GetFirstCharIndexOfCurrentLine
                Dim curPos As Integer = rtb.SelectionStart
                rtb.SelectionStart = sStart
                rtb.SelectedText = "//"
                rtb.SelectionStart = curPos + 2
            Else
                'есть выделенный текст
                Dim startLine As Integer = rtb.GetLineFromCharIndex(rtb.SelectionStart)
                Dim endLine As Integer = rtb.GetLineFromCharIndex(rtb.SelectionStart + rtb.SelectionLength)
                If startLine = endLine Then
                    'выделена одна строка - вставляем // перед выделенным
                    rtb.SelectedText = "//" + rtb.SelectedText
                Else
                    'выделено несколько строк. Ставим комментарии в начале каждой строки
                    Dim selStart As Integer = rtb.GetFirstCharIndexOfCurrentLine
                    rtb.Select(selStart, rtb.GetFirstCharIndexFromLine(endLine) + rtb.Lines(endLine).Length - selStart)
                    Dim txt As String = "//" & rtb.SelectedText.Replace(vbLf, vbLf & "//")
                    rtb.SelectedText = txt
                    Application.DoEvents()
                    rtb.Select(selStart, txt.Length)
                End If
            End If
        End If
    End Sub

    Private Sub tsbUncomment_Click(sender As Object, e As EventArgs) Handles tsbUncomment.Click, tsmiUncomment.Click
        'Удаляет символы комментария // или <!-- -->
        Dim rtb As RichTextBox = codeBox.codeBox
        Dim curline As Integer = rtb.GetLineFromCharIndex(rtb.SelectionStart)
        Dim isExec As CodeTextBox.ExecBlockEnum = CodeTextBox.ExecBlockEnum.NO_EXEC
        Dim isInText As CodeTextBox.TextBlockEnum = codeBox.codeBox.IsInTextBlock(rtb, curline, isExec)
        Dim selStart As Integer = rtb.SelectionStart

        If isInText Then
            'мы в блоке текста. Комментарии <!-- -->
            If rtb.SelectionLength = 0 Then
                'если слева/справа комментарии - удаляем. Иначе - ничего
                Dim remStart As Integer = selStart
                Dim remEnd As Integer = selStart
                If selStart > 3 AndAlso rtb.Text.Substring(selStart - 4, 4) = "<!--" Then
                    remStart = selStart - 4
                End If
                If selStart <= rtb.TextLength - 3 AndAlso rtb.Text.Substring(selStart, 3) = "-->" Then
                    remEnd = selStart + 3
                End If
                If remStart <> remEnd Then
                    rtb.Select(remStart, remEnd - remStart)
                    rtb.SelectedText = ""
                End If
            Else
                'удаляем в выделенном тексте все комментарии
                Dim txt As String = rtb.SelectedText.Replace("<!--", "").Replace("-->", "")
                rtb.SelectedText = txt
                Application.DoEvents()
                rtb.Select(selStart, txt.Length)
            End If
        Else
            'мы в блоке скрипта. Комментарии //
            If rtb.SelectionLength = 0 Then
                'если выделения нет, то удаляем // в начале строки (если есть)
                If rtb.Lines(curline).TrimStart.StartsWith("//") Then
                    Dim txt As String = rtb.Lines(curline)
                    Dim pos As Integer = txt.IndexOf("//")
                    If pos = -1 Then Return
                    If pos = 0 Then
                        txt = txt.Substring(pos + 2)
                    Else
                        txt = txt.Substring(0, pos) & txt.Substring(pos + 2)
                    End If
                    selStart = rtb.GetFirstCharIndexOfCurrentLine
                    rtb.Select(selStart, rtb.Lines(curline).Length)
                    rtb.SelectedText = txt
                    Application.DoEvents()
                    rtb.Select(selStart, txt.Length)
                End If
            Else
                'выделено несколько строк. Удаляем комментарии в начале каждой строки
                selStart = rtb.GetFirstCharIndexOfCurrentLine
                Dim endLine As Integer = rtb.GetLineFromCharIndex(rtb.SelectionStart + rtb.SelectionLength)
                rtb.Select(selStart, rtb.GetFirstCharIndexFromLine(endLine) + rtb.Lines(endLine).Length - selStart)
                Dim txt As String = rtb.SelectedText.Replace(vbLf & "//", vbLf)
                If txt.StartsWith("//") Then txt = txt.Substring(2)
                rtb.SelectedText = txt
                Application.DoEvents()
                rtb.Select(selStart, txt.Length)
            End If
        End If
    End Sub

    Private Sub tsbHighLightFull_Click(sender As Object, e As EventArgs) Handles tsbHighLightFull.Click, tsmiHighLightFull.Click
        codeBox.codeBox.HightLightColor = DEFAULT_COLORS.codeBoxHighLight
        codeBox.codeBox.ClearPreviousSelection()
        tsbHighLightFull.Checked = True
        tsbHighLightDesignate.Checked = False
        tsbHighLightNo.Checked = False
        tsmiHighLightFull.Checked = True
        tsmiHighLightDesignate.Checked = False
        tsmiHighLightNo.Checked = False
    End Sub

    Private Sub tsbHighLightDesignate_Click(sender As Object, e As EventArgs) Handles tsbHighLightDesignate.Click, tsmiHighLightDesignate.Click
        codeBox.codeBox.HightLightColor = DEFAULT_COLORS.codeBoxDesignate
        codeBox.codeBox.ClearPreviousSelection()
        tsbHighLightFull.Checked = False
        tsbHighLightDesignate.Checked = True
        tsbHighLightNo.Checked = False
        tsmiHighLightFull.Checked = False
        tsmiHighLightDesignate.Checked = True
        tsmiHighLightNo.Checked = False
    End Sub

    Private Sub tsbHighLightNo_Click(sender As Object, e As EventArgs) Handles tsbHighLightNo.Click, tsmiHighLightNo.Click
        codeBox.codeBox.HightLightColor = Nothing
        codeBox.codeBox.ClearPreviousSelection()
        tsbHighLightFull.Checked = False
        tsbHighLightDesignate.Checked = False
        tsbHighLightNo.Checked = True
        tsmiHighLightFull.Checked = False
        tsmiHighLightDesignate.Checked = False
        tsmiHighLightNo.Checked = True
    End Sub


    Private Sub tsbChars_Click(sender As Object, e As EventArgs) Handles tsbChars.Click, tsmiChars.Click
        Dim res As DialogResult = dlgChars.ShowDialog(Me)
        If res = Windows.Forms.DialogResult.OK Then
            codeBox.codeBox.SelectedText = dlgChars.SelectedText
        End If
    End Sub

    Private Sub tsbColor_Click(sender As Object, e As EventArgs) Handles tsbColor.Click, tsmiColor.Click
        Dim res As DialogResult = dlgColor.ShowDialog(Me)
        If res = Windows.Forms.DialogResult.OK Then
            Dim hUseFont As HtmlElement = dlgColor.wbColor.Document.GetElementById("useFont"), blnFont As Boolean = True
            If IsNothing(hUseFont) = False Then blnFont = CBool(hUseFont.GetAttribute("checked"))
            With codeBox.codeBox
                Dim selStart As Integer = .SelectionStart
                Dim txt As String = ""
                If blnFont Then
                    txt = "<font color=" & Chr(34) & dlgColor.SelectedColor & Chr(34) & ">" & .SelectedText & "</font>"
                Else
                    txt = dlgColor.SelectedColor
                End If

                .SelectedText = txt
                Application.DoEvents()
                .Select(selStart, txt.Length)
                .Focus()
            End With
            'codeBox.codeBox.SelectedText ="<font color=" & Chr(34) & dlgColor.SelectedColor & Chr(34) &
        End If
    End Sub

    Private Sub tsbTable_Click(sender As Object, e As EventArgs) Handles tsbTable.Click, tsmiTable.Click
        Dim res As DialogResult = dlgTable.ShowDialog(Me)
        If res = Windows.Forms.DialogResult.OK Then
            With codeBox.codeBox
                Dim selStart As Integer = .SelectionStart
                .SelectedText = dlgTable.txtResult.Text
                Application.DoEvents()
                .Select(selStart, dlgTable.txtResult.TextLength)
                .Focus()
            End With
        End If
    End Sub

    Private Sub tsbImage_Click(sender As Object, e As EventArgs) Handles tsbImage.Click, tsmiImage.Click
        Dim res As DialogResult = dlgImage.ShowDialog(Me)
        If res = Windows.Forms.DialogResult.OK Then
            codeBox.codeBox.SelectedText = dlgImage.selectedImage
        End If
    End Sub

#Region "Local Seek"

    Private Sub tsbSeek_Click(sender As Object, e As EventArgs) Handles tsbSeek.Click, tsmiFind.Click
        If IsNothing(pnlSeek.Parent) OrElse Object.Equals(pnlSeek.Parent, codeBoxPanel) = False Then
            codeBoxPanel.Controls.Add(pnlSeek)
        End If

        pnlSeek.Location = New Point(codeBoxPanel.ClientSize.Width - pnlSeek.Width - 15, codeBox.Top)
        If tsbSeek.Checked Then
            pnlSeek.Hide()
            tsbSeek.Checked = False
        Else
            pnlSeek.BringToFront()
            pnlSeek.Show()
            tsbSeek.Checked = True
            If String.IsNullOrEmpty(cmbSeek.Text) Then
                cmbSeek.Focus()
            Else
                cmbSeek_TextChanged(cmbSeek, New EventArgs)
            End If
        End If
    End Sub

    Dim pnlSeekClickPos As Point = Nothing
    Dim pnlSeekMoving As Boolean = False
    Private Sub pnlSeek_MouseDown(sender As Object, e As MouseEventArgs) Handles pnlSeek.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then pnlSeekClickPos = e.Location : pnlSeekMoving = True
    End Sub

    Private Sub pnlSeek_MouseMove(sender As Object, e As MouseEventArgs) Handles pnlSeek.MouseMove
        If Not pnlSeekMoving Then Return
        pnlSeek.Location = New Point(pnlSeek.Left + e.X - pnlSeekClickPos.X, pnlSeek.Top + e.Y - pnlSeekClickPos.Y)
    End Sub

    Private Sub pnlSeek_MouseUp(sender As Object, e As MouseEventArgs) Handles pnlSeek.MouseUp
        pnlSeekClickPos = Nothing
        pnlSeekMoving = False
    End Sub

    Private wordBoundsArray() As Char = {" "c, "("c, ")"c, "["c, "]"c, "."c, ","c, "'"c, "="c, "+"c, "-"c, "*"c, "^"c, "/"c, "\"c, "<"c, ">"c, "&"c, "!"c, ";"c, "#"c, "$"c, Chr(34), Chr(10), "?"}
    Dim seekLastPos As Integer = -1

    Private Sub cmbSeek_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbSeek.KeyDown, chkSeekCase.KeyDown, chkSeekWholeWord.KeyDown, btnSeekBackward.KeyDown, btnSeekBackward.KeyDown
        If e.KeyCode = Keys.Return Then
            btnSeekBackward_Click(btnSeekForward, New EventArgs)
        ElseIf e.KeyCode = Keys.Escape Then
            tsbSeek_Click(tsbSeek, New EventArgs)
        End If
    End Sub
    Private Sub cmbSeek_TextChanged(sender As Object, e As EventArgs) Handles cmbSeek.TextChanged, chkSeekCase.CheckedChanged, chkSeekWholeWord.CheckedChanged
        seekLastPos = -1
        If cmbSeek.Text.Length = 0 Then
            cmbSeek.ForeColor = Color.Black
            btnSeekBackward.Enabled = False
            btnSeekForward.Enabled = False
            Return
        End If
        If codeBox.codeBox.SeekString(cmbSeek.Text, chkSeekCase.Checked, chkSeekWholeWord.Checked) = -1 Then
            cmbSeek.ForeColor = Color.Red
            btnSeekBackward.Enabled = False
            btnSeekForward.Enabled = False
        Else
            cmbSeek.ForeColor = Color.DarkGreen
            btnSeekBackward.Enabled = True
            btnSeekForward.Enabled = True
        End If
    End Sub

    Private Sub btnSeekForward_Click(sender As Object, e As EventArgs) Handles btnSeekForward.Click, tsmiFindNext.Click
        Dim strSeek As String = cmbSeek.Text
        If String.IsNullOrEmpty(strSeek) Then Return
        Dim compar As System.StringComparison = IIf(Not chkSeekCase.Checked, StringComparison.CurrentCultureIgnoreCase, StringComparison.CurrentCulture)
begin:
        Dim newPos As Integer = codeBox.Text.IndexOf(strSeek, seekLastPos + 1, compar)
        If chkSeekWholeWord.Checked Then
            If newPos > 0 AndAlso wordBoundsArray.Contains(codeBox.Text.Chars(newPos - 1)) = False Then
                newPos = -1 'это не слово целиком
            End If

            If newPos < codeBox.codeBox.TextLength - 1 AndAlso wordBoundsArray.Contains(codeBox.Text.Chars(newPos + strSeek.Length)) = False Then
                newPos = -1 'это не слово целиком
            End If
        End If

        seekLastPos = newPos
        If seekLastPos = -1 Then
            Dim res As DialogResult = MessageBox.Show("Достигнут конец текта. Продолжить с начала?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res = Windows.Forms.DialogResult.No Then Return
            seekLastPos = -1
            GoTo begin
        Else
            If cmbSeek.Items.Contains(strSeek) = False Then cmbSeek.Items.Add(strSeek)
            codeBox.codeBox.Select(seekLastPos, strSeek.Length)
            codeBox.codeBox.Focus()
        End If
    End Sub

    Private Sub btnSeekBackward_Click(sender As Object, e As EventArgs) Handles btnSeekBackward.Click
        Dim strSeek As String = cmbSeek.Text
        If String.IsNullOrEmpty(strSeek) Then Return
        Dim compar As System.StringComparison = IIf(Not chkSeekCase.Checked, StringComparison.CurrentCultureIgnoreCase, StringComparison.CurrentCulture)
        If seekLastPos = -1 Then seekLastPos = codeBox.codeBox.TextLength
        If seekLastPos = 0 Then
            Dim res As DialogResult = MessageBox.Show("Достигнуто начало текта. Повторить поиск заново?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res = Windows.Forms.DialogResult.No Then Return
            seekLastPos = codeBox.codeBox.TextLength
        End If

begin:
        Dim newPos As Integer = codeBox.Text.LastIndexOf(strSeek, seekLastPos - 1, compar)
        If chkSeekWholeWord.Checked Then
            If newPos > 0 AndAlso wordBoundsArray.Contains(codeBox.Text.Chars(newPos - 1)) = False Then
                newPos = -1 'это не слово целиком
            End If

            If newPos < codeBox.codeBox.TextLength - 1 AndAlso wordBoundsArray.Contains(codeBox.Text.Chars(newPos + strSeek.Length)) = False Then
                newPos = -1 'это не слово целиком
            End If
        End If

        seekLastPos = newPos
        If seekLastPos = -1 Then
            Dim res As DialogResult = MessageBox.Show("Достигнуто начало текта. Повторить поиск заново?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res = Windows.Forms.DialogResult.No Then Return
            seekLastPos = codeBox.codeBox.TextLength
            GoTo begin
        Else
            If cmbSeek.Items.Contains(strSeek) = False Then cmbSeek.Items.Add(strSeek)
            codeBox.codeBox.Select(seekLastPos, strSeek.Length)
            codeBox.codeBox.Focus()
        End If
    End Sub
#End Region

    Dim strKeyboardENG As String = "QWERTYUIOP{}ASDFGHJKL:" + Chr(34) + "ZXCVBNM<>?qwertyuiop[]asdfghjkl;'zxcvbnm,./"
    Private Sub tsmiKeyboard_Click(sender As Object, e As EventArgs) Handles tsmiKeyboard.Click
        Dim selText As String = codeBox.codeBox.SelectedText
        selText = TranslitText(selText)

        With codeBox.codeBox
            Dim selStart As Integer = .SelectionStart
            Dim selLength As Integer = .SelectionLength
            .SelectedText = selText
            .Select(selStart, selLength)
        End With
    End Sub

    Public Function TranslitText(ByVal selText As String) As String
        If String.IsNullOrEmpty(selText) Then Return ""
        Dim strFrom As String = "", strTo As String = ""
        Dim isEnglish As Boolean = True
        For i As Integer = 0 To selText.Length - 1
            If strKeyboardENG.IndexOf(selText.Chars(i)) > -1 Then
                strFrom = strKeyboardENG
                strTo = My.Resources.keyBord
                Exit For
            ElseIf My.Resources.keyBord.IndexOf(selText.Chars(i)) > -1 Then
                strFrom = My.Resources.keyBord
                strTo = strKeyboardENG
                Exit For
            End If
        Next i
        If String.IsNullOrEmpty(strFrom) Then Return selText

        Dim res As New System.Text.StringBuilder
        For i As Integer = 0 To selText.Length - 1
            Dim pos As Integer = strFrom.IndexOf(selText.Chars(i))
            If pos > -1 Then
                res.Append(strTo.Chars(pos))
            Else
                res.Append(selText.Chars(i))
            End If
        Next i
        Return res.ToString
    End Function

    Private Sub tsbList_Click(sender As Object, e As EventArgs) Handles tsbList.Click, tsmiList.Click
        Dim res As DialogResult = dlgListHtml.ShowDialog(Me)
        If res = Windows.Forms.DialogResult.OK Then
            With codeBox.codeBox
                Dim selStart As Integer = .SelectionStart
                Dim txt As String = dlgListHtml.GetListString
                .SelectedText = txt
                Application.DoEvents()
                .Select(selStart, txt.Length)
                .Focus()
            End With
        End If
    End Sub

    Private Sub tsbFontSize_Click(sender As Object, e As EventArgs) Handles tsbFontSize.Click, tsmiFontSize.Click
        With codeBox.codeBox
            Dim selStart As Integer = .SelectionStart
            Dim sel As String = .SelectedText
            Dim txt As String = sel
            If txt.Trim.Length = 0 Then txt = "Пример для тестирования шрифта"
            dlgFontSize.strTest = txt

            Dim res As DialogResult = dlgFontSize.ShowDialog(Me)
            If res = Windows.Forms.DialogResult.OK Then
                txt = dlgFontSize.WrapTextToFont(.SelectedText)
                .SelectedText = txt
                Application.DoEvents()
                .Select(selStart, txt.Length)
                .Focus()
            End If
        End With

    End Sub

    Private Sub tsbMarquee_Click(sender As Object, e As EventArgs) Handles tsbMarquee.Click, tsmiMarquee.Click
        With codeBox.codeBox
            Dim selStart As Integer = .SelectionStart
            Dim sel As String = .SelectedText
            Dim txt As String = sel
            If txt.Trim.Length = 0 Then txt = "Пример бегущей строки"
            dlgMarquee.sampleText = txt

            Dim res As DialogResult = dlgMarquee.ShowDialog(Me)
            If res = Windows.Forms.DialogResult.OK Then
                txt = dlgMarquee.selectedMarquee & sel & "</MARQEE>"
                .SelectedText = txt
                Application.DoEvents()
                .Select(selStart, txt.Length)
                .Focus()
            End If
        End With
    End Sub

    Private Sub tsbExec_Click(sender As Object, e As EventArgs) Handles tsbExec.Click, tsmiExec.Click
        WrapWithTag(codeBox, "<exec>?", "</exec>")
    End Sub

    Private Sub tsmiSelLine_Click(sender As Object, e As EventArgs) Handles tsmiSelLine.Click
        Dim rtb As RichTextBox = codeBox.codeBox
        Dim selStart As Integer = rtb.GetFirstCharIndexOfCurrentLine
        Dim lineStart As Integer = rtb.GetLineFromCharIndex(selStart)
        Dim lineEnd As Integer = rtb.GetLineFromCharIndex(selStart + rtb.SelectionLength)
        Dim selEnd As Integer = rtb.GetFirstCharIndexFromLine(lineEnd) + rtb.Lines(lineEnd).Length
        rtb.Select(selStart, selEnd - selStart)
    End Sub

    Private Sub tsbUndo_Click(sender As Object, e As EventArgs) Handles tsbUndo.Click
        codeBox.codeBox.csUndo.Undo()
    End Sub

    Private Sub tsbRedo_Click(sender As Object, e As EventArgs) Handles tsbRedo.Click
        codeBox.codeBox.csUndo.Redo()
    End Sub

    Private Sub tsbRestoreInitial_Click(sender As Object, e As EventArgs) Handles tsbRestoreInitial.Click
        If MessageBox.Show("Все изменения в данном скрипте будут отменены. Продолжить?", "MatewQuest", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then Return
        codeBox.codeBox.csUndo.RestoreInitalState()
    End Sub

    Private Sub tsbSelector_Click(sender As Object, e As EventArgs) Handles tsbSelector.Click, tsmiSelector.Click
        'Вставка селектора #X:
        With codeBox.codeBox
            If .TextLength = 0 Then
                .SelectedText = "#1:"
                Return
            End If

            Dim txt As String = .Text.Substring(0, .SelectionStart)
            Dim num As Integer = 1
            Do
                Dim pos As Integer = txt.IndexOf("#" & num.ToString & ":")
                If pos = -1 Then Exit Do
                num += 1
            Loop
            .SelectedText = "#" & num.ToString & ":"
        End With
    End Sub

    Private Sub tsmiCopyAsHTML_Click(sender As Object, e As EventArgs) Handles tsmiCopyAsHTML.Click
        If codeBoxPanel.Visible = False OrElse IsNothing(codeBox.codeBox.CodeData) Then Return
        Dim res As String = mScript.ConvertCodeDataToHTML(codeBox.codeBox.CodeData, Not codeBox.codeBox.IsTextBlockByDefault)
        Clipboard.SetText(res, TextDataFormat.Text)
        MessageBox.Show("Скрипт, конвертированный в html, скопирован в буфер обмена.", "MatewScript", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub tsmiUndo_Click(sender As Object, e As EventArgs) Handles tsmiUndo.Click
        If codeBoxPanel.Visible = True Then
            codeBox.codeBox.csUndo.Undo()
            Return
        End If
        Dim ch As clsPanelManager.clsChildPanel = cPanelManager.ActivePanel
        If IsNothing(ch) Then Return
        Dim c As Control = ch.ActiveControl
        If IsNothing(c) Then Return
        Dim tName As String = c.GetType.Name
        If tName = "TextBoxEx" Then
            Dim t As TextBox = c
            If t.CanUndo Then t.Undo()
        End If

    End Sub

    Private Sub tsmiRedo_Click(sender As Object, e As EventArgs) Handles tsmiRedo.Click
        If codeBoxPanel.Visible = True Then
            codeBox.codeBox.csUndo.Redo()
            Return
        End If
        Dim ch As clsPanelManager.clsChildPanel = cPanelManager.ActivePanel
        If IsNothing(ch) Then Return
        Dim c As Control = ch.ActiveControl
        If IsNothing(c) Then Return
        Dim tName As String = c.GetType.Name
        If tName = "TextBoxEx" Then
            Dim t As TextBox = c
            If t.CanUndo Then t.Undo()
        End If
    End Sub

    Private Sub tsmiSave_Click(sender As Object, e As EventArgs) Handles tsmiSave.Click
        If codeBoxPanel.Visible Then
            codeBox.codeBox.CheckTextForSyntaxErrors(codeBox.codeBox)
            Dim ee As New System.ComponentModel.CancelEventArgs
            Call codeBox_Validating(codeBox, ee)
            If ee.Cancel Then Return
        ElseIf dgwVariables.Visible Then
            Call dgwVariables_Validating(dgwVariables, New System.ComponentModel.CancelEventArgs)
        End If
        questEnvironment.SaveQuest()
    End Sub

    Private Sub tsmiLoad_Click(sender As Object, e As EventArgs) Handles tsmiLoad.Click
        questEnvironment.ClearAllData()
        questEnvironment.LoadQuest(My.Computer.FileSystem.CombinePath(Application.StartupPath, "Quests\myQuest"))
    End Sub

    Private Sub dgwVariables_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgwVariables.CellContentClick

    End Sub

    Private Sub tsmiRun_Click(sender As Object, e As EventArgs) Handles tsmiRun.Click
        'questEnvironment.SaveQuest()
        tsmiSave.PerformClick()
        If mScript.LAST_ERROR.Length > 0 Then
            'MessageBox.Show("Нельзя запускать код с ошибками!", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        If currentClassName = "L" OrElse currentClassName = "A" Then actionsRouter.SaveActions()

        modPlayer.SetWindowsContainers()
        If modPlayer.LoadPlayerDocuments() = False Then Return
        questEnvironment.EDIT_MODE = False
        questEnvironment.OPENED_FROM_EDITOR = True
        frmPlayer.Show()
        GVARS.G_PREVLOC = -1
        GVARS.G_CURLOC = -1
        GVARS.TIME_IN_GAME = Now.Ticks
        GVARS.TIME_STARTING = Now.Ticks
        GVARS.TIME_SAVING = Now.Ticks
        GVARS.HOLD_UP_TIMERS = False
        GVARS.G_CURMENU = -1
        GVARS.G_CURLIST = -1
        GVARS.G_CURAUDIO = -1
        GVARS.G_CURMAP = -1
        GVARS.G_CURMAPCELL = -1
        GVARS.G_PREVMAPCELL = -1
        PlayerDate_SetInitialTime(Nothing)
        ObjectWindowPrepareFace(Nothing)
        ObjectWindowPrepareStruct(Nothing)
        'Dim ts As New TimeSpan(Date.Now.Ticks)
        InitTimers()
        mScript.eventRouter.RunEvent(mScript.mainClass(mScript.mainClassHash("Q")).Properties("QuestStartEvent").eventId, Nothing, "QuestStartEvent", True)
        Me.Hide()
        'frmMap.Show()
        'frmMagic.Show()
    End Sub

    ''' <summary>
    ''' Процедура, вызывающая событие LocationHtmlEvent - заглушка для корректной работы при отбражении содержимого Description
    ''' </summary>
    ''' <param name="args"></param>
    Public Sub LocationHtmlEvent(ByVal args As String)

    End Sub

    Private Sub tsbFullScreen_Click(sender As Object, e As EventArgs) Handles tsbFullScreen.Click
        If tsbFullScreen.Checked Then
            tsbFullScreen.Checked = False
            SplitInner.Panel1Collapsed = False
            Dim pInnerVisible As Boolean = True
            Dim classId As Integer = -1
            If mScript.mainClassHash.TryGetValue(currentClassName, classId) Then
                If mScript.mainClass(classId).LevelsCount = 0 Then pInnerVisible = False
            End If

            If pInnerVisible Then
                SplitOuter.Panel1Collapsed = False
            End If
        Else
            tsbFullScreen.Checked = True
            SplitInner.Panel1Collapsed = True
            SplitOuter.Panel1Collapsed = True
        End If
    End Sub

End Class
