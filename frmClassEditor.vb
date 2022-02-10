'Редактор классов (сторуктуры MainClass)
Public Class frmClassEditor
    ''' <summary>
    ''' Класс для временного хранения параметров свойств типа Event
    ''' </summary>
    Private Class PropParamsClass
        'Levelx, T_PARAM_ARRAY и T_RETURN заполняются исходя из значения Property.params().Type:
        'Level1 +1, Level2 +2, Level3 +4, Return +8, ParamArray +16
        Public Enum ParamTypeEnum As Byte
            T_PARAM = 0
            T_PARAM_ARRAY = 1
            T_RETURN = 2
        End Enum
        Public Level1 As Boolean = False
        Public Level2 As Boolean = False
        Public Level3 As Boolean = False
        Public paramType As ParamTypeEnum
        Public paramCaption As String
        Public paramDescription As String
    End Class
    Private arrPropParams As List(Of PropParamsClass)

    Enum SavingClassActionEnum
        NEW_ELEMENT = 0 'Создаем новое свойство/функцию/класс
        EDIT_ELEMENT = 1 'Редактируем свойство/функцию/класс
    End Enum

    ''' <summary>
    ''' Класс для хранения информации и выполненных изменениях в структуре
    ''' </summary>
    ''' <remarks></remarks>
    Private Class EditQueueClass
        Public Enum EditActionTypeEnum As Byte
            EDIT_PROPERTY_NAME = 0
            CREATE_NEW_PROPERTY = 1
            REMOVE_PROPERTY = 2
            EDIT_CLASS_NAME = 3
            CREATE_CLASS = 4
            REMOVE_CLASS = 5
        End Enum
        Public ActionType As EditActionTypeEnum
        Public newValue As String
        Public oldValue As String
        Public classId As Integer
    End Class

    ''' <summary>Id класса выбранного элемента и что это за элемент (свойство или функция)</summary>
    Dim selectedClassId As Integer, selectedElementType As MatewScript.funcAndPropHashType.funcOrPropEnum
    ''' <summary>Имя выбранного элемента и что с ним делаем (новое или редактируем)</summary>
    Dim selectedElementName As String
    ''' <summary>Что делаем - создаем новый или редактируем старый</summary>
    Dim curAction As SavingClassActionEnum
    Dim curFuncParams() As MatewScript.paramsType 'для хранения параметров редактируемой функции
    'цвета для узлов дерева классов (обычный и выбранный)
    Dim defNodeColor As Color, nodeSelColor As Color
    Dim prevNode As TreeNode = Nothing 'предыдущий узел (для возврата ему цвета по-умолчанию)
    'перемещаемый узел и узел, куда перемещается (для события DragDrop)
    Dim nodeUnderMouseToDrag, nodeUnderMouseToDrop As TreeNode
    Dim dragBoxFromMouseDown As Rectangle 'квадат, при выходе за пределы которого начинается DragDrop
    Dim dragNodeType As MatewScript.funcAndPropHashType.funcOrPropEnum 'тип перемещаемого узла (свойство или функция)
    Dim CodeData() As CodeTextBox.CodeDataType 'для хранения копии скриптов перед их редактированием
    Private initialClassesState() As MatewScript.MainClassType
    ''' <summary>Если True, то режим полного редактирования главной структуры (а не структуры квеста)</summary>
    Private DEEP_EDIT_MODE As Boolean = True
    ''' <summary>список выполненных действий, связанных с изменением структуры склассов</summary>
    Private lstQueue As New List(Of EditQueueClass)

    Private Sub DoChanges()
        If lstQueue.Count = 0 Then Return

        For i As Integer = lstQueue.Count - 1 To 0 Step -1
            Dim itm As EditQueueClass = lstQueue(i)
            Select Case itm.ActionType
                Case EditQueueClass.EditActionTypeEnum.CREATE_CLASS
                    mScript.UpdateBasicFunctionsParamsWhichIsElements(itm.classId)
                    cGroups.dictGroups.Add(mScript.mainClass(itm.classId).Names(0), New List(Of clsGroups.clsGroupInfo))
                Case EditQueueClass.EditActionTypeEnum.EDIT_CLASS_NAME
                    Dim arrNames() As String = mScript.mainClass(itm.classId).Names
                    GlobalSeeker.ReplaceClassNameInStruct(itm.classId, arrNames)
                    mScript.UpdateBasicFunctionsParamsWhichIsElements(itm.classId, itm.oldValue)
                    If itm.oldValue <> arrNames(0) Then
                        'Переименовываем класс в группах
                        Dim g As List(Of clsGroups.clsGroupInfo) = cGroups.dictGroups(itm.oldValue)
                        cGroups.dictGroups.Remove(itm.oldValue)
                        cGroups.dictGroups.Add(arrNames(0), g)
                        If cGroups.dictRemoved.ContainsKey(itm.oldValue) Then
                            g = cGroups.dictRemoved(itm.oldValue)
                            cGroups.dictRemoved.Remove(itm.oldValue)
                            cGroups.dictRemoved.Add(arrNames(0), g)
                        End If
                        'переименовываем в событиях изменения свойств
                        mScript.trackingProperties.RenameClass(itm.oldValue, arrNames(0))
                    End If
                Case EditQueueClass.EditActionTypeEnum.REMOVE_CLASS
                    'Прочесывание структуры mainClass на предмет наличия функция и свойств, содержащих в качестве возвращаемого значения или параметра функции тип ELEMENT нашего удаляемого класса. Если будет такое найдено - очищаем формат
                    GlobalSeeker.RemoveClassNameInStruct(itm.classId)

                    mScript.trackingProperties.RemoveClass(mScript.mainClass(itm.classId).Names(0)) 'удаляем из событий изменений свойств
                    'Удаляем группы
                    If cGroups.dictGroups.ContainsKey(mScript.mainClass(itm.classId).Names(0)) Then
                        cGroups.dictGroups.Remove(mScript.mainClass(itm.classId).Names(0))
                    End If
                    'Вносим изменения в удаленные элементы 
                    removedObjects.RemoveClass(itm.classId)

                Case EditQueueClass.EditActionTypeEnum.CREATE_NEW_PROPERTY
                    If mScript.mainClass(itm.classId).Names(0) = "A" Then actionsRouter.AddProperty(itm.newValue)
                    removedObjects.AddProperty(itm.classId, itm.newValue)
                Case EditQueueClass.EditActionTypeEnum.EDIT_PROPERTY_NAME
                    mScript.trackingProperties.RenameProperty(itm.classId, itm.oldValue, itm.newValue)
                    If currentClassName = "A" Then actionsRouter.RenameProperty(itm.oldValue, itm.newValue)
                    removedObjects.RenameProperty(itm.classId, itm.oldValue, itm.newValue)
                    GlobalSeeker.ReplaceElementNameInStruct(itm.classId, itm.oldValue, itm.newValue, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                Case EditQueueClass.EditActionTypeEnum.REMOVE_PROPERTY
                    If mScript.mainClass(itm.classId).Names(0) = "A" Then actionsRouter.RemoveProperty(itm.newValue)
                    removedObjects.RemoveProperty(itm.classId, itm.newValue) 'удаляеем свойство в удаленных элементах

                    'удаляем в дочерних
                    If IsNothing(mScript.mainClass(selectedClassId).ChildProperties) = False AndAlso mScript.mainClass(selectedClassId).ChildProperties.Count > 0 Then
                        For chId As Integer = 0 To mScript.mainClass(selectedClassId).ChildProperties.Count - 1
                            Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(selectedClassId).ChildProperties(chId)(itm.newValue)
                            If ch.eventId > 0 Then mScript.eventRouter.RemoveEvent(ch.eventId)
                            If IsNothing(ch.ThirdLevelProperties) = False AndAlso ch.ThirdLevelProperties.Count > 0 Then
                                For child3Id As Integer = 0 To ch.ThirdLevelEventId.Count - 1
                                    If ch.ThirdLevelEventId(child3Id) > 0 Then mScript.eventRouter.RemoveEvent(ch.ThirdLevelEventId(child3Id))
                                Next child3Id
                            End If
                            mScript.mainClass(selectedClassId).ChildProperties(chId).Remove(itm.newValue)
                        Next chId
                    End If

                    'удаляем из событий изменения свойств
                    mScript.trackingProperties.RemoveProperty(itm.classId, itm.newValue)
                    GlobalSeeker.CheckElementNameInStruct(itm.classId, itm.newValue, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
            End Select
        Next
    End Sub

    ''' <summary> Обновляет  EditorId во всех элементах всех класов исходя из расположения в дереве (действие перед сохранением классов)</summary>
    Private Sub UpdateAllEditorId()
        '"Class " + mScript.mainClass(i).Names(0)
        For i As Integer = 0 To treeClasses.Nodes.Count - 1
            'перебор в цикле всех узлов-классов
            Dim classNode As TreeNode = treeClasses.Nodes(i)
            Dim classId As Integer = -1
            If Not mScript.mainClassHash.TryGetValue(classNode.Tag.ToString.Substring(6), classId) Then Continue For
            Dim funcNode As TreeNode = classNode.Nodes(0)
            Dim propNode As TreeNode = classNode.Nodes(1)
            For j As Integer = 0 To funcNode.Nodes.Count - 1
                'обновление EditorId в функциях класса i
                Dim fName As String = funcNode.Nodes(j).Text
                Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions(fName)
                p.editorIndex = funcNode.Nodes(j).Index
            Next j
            For j As Integer = 0 To propNode.Nodes.Count - 1
                'обновление EditorId в свойствах класса i
                Dim pName As String = propNode.Nodes(j).Text
                Dim p As MatewScript.PropertiesInfoType = Nothing  '= mScript.mainClass(classId).Properties(pName)
                If mScript.mainClass(classId).Properties.TryGetValue(pName, p) = False Then Continue For
                p.editorIndex = propNode.Nodes(j).Index
            Next j
        Next i
    End Sub

    Private Sub frmClassEditor_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'закрытие редактора
        Dim ret As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Сохранить изменения?", MsgBoxStyle.YesNoCancel)
        If ret = MsgBoxResult.Cancel Then
            e.Cancel = True
            Return
        End If

        If ret = MsgBoxResult.Yes Then
            If IsNothing(cPanelManager.lstPanels) = False AndAlso cPanelManager.lstPanels.Count > 0 Then
                'удаляем все панели
                For i As Integer = 0 To cPanelManager.lstPanels.Count - 1
                    cPanelManager.RemovePanel(cPanelManager.lstPanels.Last, False)
                Next
                cPanelManager.dictLastPanel.Clear()
            End If
            'удаляем структуру панелей (она обновится при первом вызове)
            cPanelManager.dictDefContainers.Clear()

            UpdateAllEditorId()
            If DEEP_EDIT_MODE Then
                mScript.SaveClasses() 'сохранение классов
            Else
                mScript.SaveClasses(My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "classes.xml")) 'сохранение классов
            End If
            mScript.FillFuncAndPropHash()
            DoChanges()
        ElseIf ret = MsgBoxResult.No Then
            Erase mScript.mainClass
            mScript.mainClass = initialClassesState
            lstQueue.Clear()
        End If
        e.Cancel = False
    End Sub

    Private Sub frmClassEditor_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'mScript.LoadClasses()
        ReDim initialClassesState(mScript.mainClass.Count - 1)
        For i As Integer = 0 To initialClassesState.Count - 1
            initialClassesState(i) = mScript.mainClass(i).Clone(True, False)
        Next

        splitMain.Dock = DockStyle.Fill
        pnlClass.Dock = DockStyle.Fill
        pnlFunction.Dock = DockStyle.Fill
        pnlProperty.Dock = DockStyle.Fill
        pnlPropParams.Dock = DockStyle.Fill
        splitMain.Dock = DockStyle.Fill
        splitCode.Dock = DockStyle.Fill
        splitCode.Hide()

        ofd.InitialDirectory = APP_HELP_PATH
        FillTreeViewWithClasses(treeClasses) 'заполнение дерева классов
        defNodeColor = treeClasses.Nodes(0).ForeColor
        nodeSelColor = Color.Red
        FillListWithClasses()
    End Sub

    ''' <summary>
    ''' Процедура заполняет дерево классов
    ''' </summary>
    ''' <param name="tree">объект <see cref="TreeView">TrreView</see> для заполнения</param>
    Private Sub FillTreeViewWithClasses(ByRef tree As TreeView)
        tree.Nodes.Clear() 'очищаем дерево от предыдущих классов
        Dim treeClasses, treeFunctions, treeProperties As TreeNode
        'перебираем в цикле все классы
        For i As Integer = 0 To mScript.mainClass.GetUpperBound(0)
            'вставляем узел с именами класса и тэгом вида "Class X", где Х - первое имя класса
            treeClasses = tree.Nodes.Add("Класс " + Join(mScript.mainClass(i).Names, ", "))
            treeClasses.Tag = "Class " + mScript.mainClass(i).Names(0)
            'вставляем функции
            treeFunctions = treeClasses.Nodes.Add("Функции")
            treeFunctions.Tag = "func"
            Dim arrOrder() As Integer = mScript.CreatePropertiesOrderArray(mScript.mainClass(i).Functions)
            For j As Integer = 0 To arrOrder.Count - 1
                treeFunctions.Nodes.Add(mScript.mainClass(i).Functions.ElementAt(arrOrder(j)).Key)
            Next
            'вставляем свойства
            treeProperties = treeClasses.Nodes.Add("Свойства")
            treeProperties.Tag = "prop"
            arrOrder = mScript.CreatePropertiesOrderArray(mScript.mainClass(i).Properties)
            For j As Integer = 0 To arrOrder.Count - 1
                treeProperties.Nodes.Add(mScript.mainClass(i).Properties.ElementAt(arrOrder(j)).Key)
            Next
        Next
    End Sub

#Region "TreeClasses Events"
    Private Sub treeClasses_AfterSelect(ByVal sender As TreeView, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treeClasses.AfterSelect
        'При выборе нового узла:
        '- закрашиваем его красный и возвращаем старый цвет предыдущему узлу
        '- в зависимости что за узел выбран, отображаем и заполняем форму для редактирования класса / свойства / функции
        '- в зависимости что за узел выбран, меняем текст и доступность кнопки "Удалить"
        If IsNothing(e.Node) Then Return
        If IsNothing(prevNode) = False Then prevNode.ForeColor = defNodeColor 'возвращаем цвет по-умолчанию предыдущему узлу
        e.Node.ForeColor = nodeSelColor 'выделяем цветом текущий выбранный узел
        prevNode = e.Node 'сохраняем текущий узел, что бы затем можно было вернуть ему цвет по-умолчанию


        pnlPropParams.Hide()
        Dim strTag As String = e.Node.Tag
        If IsNothing(strTag) = False AndAlso strTag.StartsWith("Class ") Then
            'Выбранный узел - класс
            'Получаем Id класса
            strTag = strTag.Substring(6)
            selectedClassId = mScript.mainClassHash(strTag)
            selectedElementName = ""
            'Заполняем форму редактирования класса
            curAction = SavingClassActionEnum.EDIT_ELEMENT
            txtNewClassName.Text = Join(mScript.mainClass(selectedClassId).Names, ", ")
            nudClassLevels.Value = mScript.mainClass(selectedClassId).LevelsCount + 1
            txtClassHelpFile.Text = mScript.mainClass(selectedClassId).HelpFile
            FillComboWithPropertiesForDefault(selectedClassId, cmbClassDefProperty)
            'отображаем форму
            pnlFunction.Hide()
            pnlProperty.Hide()
            pnlClass.Show()
            'меняем подписть кнопки "Удалить"
            btnRemove.Enabled = True
            btnRemove.Text = "Удалить класс " + strTag
            btnUp.Enabled = False
            btnDn.Enabled = False
        Else
            If e.Node.Parent.Tag = "func" Then
                'Выбранный узел - функция
                'Получаем Id класса
                selectedClassId = mScript.mainClassHash(e.Node.Parent.Parent.Tag.ToString.Substring(6))
                selectedElementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION 'это - функция
                selectedElementName = e.Node.Text 'Имя функции

                'Заполняем форму редактирования функции
                curAction = SavingClassActionEnum.EDIT_ELEMENT
                txtFuncName.Text = selectedElementName
                txtFuncDescription.Text = mScript.mainClass(selectedClassId).Functions(selectedElementName).Description
                txtFuncEditorName.Text = mScript.mainClass(selectedClassId).Functions(selectedElementName).EditorCaption
                cmbFuncHidden.SelectedIndex = mScript.mainClass(selectedClassId).Functions(selectedElementName).Hidden
                txtFuncHelpFile.Text = mScript.mainClass(selectedClassId).Functions(selectedElementName).helpFile
                lstFuncReturnType.SelectedIndex = mScript.mainClass(selectedClassId).Functions(selectedElementName).returnType
                txtFuncResult.Clear()
                If IsNothing(mScript.mainClass(selectedClassId).Functions(selectedElementName).returnArray) = False Then
                    'заполняем текстбокс значениями, которые может возвращать функция
                    For i As Integer = 0 To mScript.mainClass(selectedClassId).Functions(selectedElementName).returnArray.GetUpperBound(0)
                        txtFuncResult.AppendText(mScript.mainClass(selectedClassId).Functions(selectedElementName).returnArray(i) + vbNewLine)
                    Next
                End If
                'заполняем структуру curFuncParams данными о параметрах функции
                curFuncParams = mScript.mainClass(selectedClassId).Functions(selectedElementName).params
                'очищаем список параметров функции
                lstFuncParams.Items.Clear()
                If IsNothing(mScript.mainClass(selectedClassId).Functions(selectedElementName).params) OrElse mScript.mainClass(selectedClassId).Functions(selectedElementName).params.Count = 0 Then
                    'параметров нет
                    pnlFuncParams.Hide()
                    nudFuncMin.Value = 0
                    nudFuncMax.Value = 0
                Else
                    'заполняем список lstFuncParams названиями параметров функции
                    lstFuncParams.BeginUpdate()
                    For i As Integer = 0 To mScript.mainClass(selectedClassId).Functions(selectedElementName).params.GetUpperBound(0)
                        lstFuncParams.Items.Add(mScript.mainClass(selectedClassId).Functions(selectedElementName).params(i).Name)
                    Next
                    If mScript.mainClass(selectedClassId).Functions(selectedElementName).paramsMax = -1 Then
                        'максимальное кол-во параметров неограниченно - значение в nudFuncMax ставим равным кол-ву описанных параметров (хотя он все-равно будет недоступен)
                        nudFuncMax.Value = mScript.mainClass(selectedClassId).Functions(selectedElementName).params.GetUpperBound(0) + 1
                    Else
                        'максимальное кол-во параметров четко определено - значение в nudFuncMax берем из структуры mainClass
                        nudFuncMax.Value = mScript.mainClass(selectedClassId).Functions(selectedElementName).paramsMax
                    End If
                    nudFuncMin.Value = mScript.mainClass(selectedClassId).Functions(selectedElementName).paramsMin
                    pnlFuncParams.Show()
                    If lstFuncParams.Items.Count > 0 Then lstFuncParams.SelectedIndex = 0
                    lstFuncParams.EndUpdate()
                End If

                'отображаем форму
                pnlFunction.Show()
                pnlClass.Hide()
                pnlProperty.Hide()
                'меняем подписть кнопки "Удалить"
                btnRemove.Enabled = True
                btnRemove.Text = "Удалить функцию " + selectedElementName

                btnUp.Enabled = (e.Node.Index > 0)
                btnDn.Enabled = (e.Node.Index < e.Node.Parent.Nodes.Count - 1)
            ElseIf e.Node.Parent.Tag = "prop" Then
                'Выбранный узел - свойство
                'Получаем Id класса
                selectedClassId = mScript.mainClassHash(e.Node.Parent.Parent.Tag.ToString.Substring(6))
                selectedElementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_PROPERTY 'это - свойство
                selectedElementName = e.Node.Text 'Имя свойства
                If Not mScript.mainClass(selectedClassId).Properties.ContainsKey(selectedElementName) Then
                    'e.Node.Remove()
                    Return
                End If
                Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(selectedClassId).Properties(selectedElementName)
                'Заполняем форму редактирования свойства
                curAction = SavingClassActionEnum.EDIT_ELEMENT
                txtPropName.Text = selectedElementName
                txtPropEditorName.Text = p.EditorCaption
                cmbPropHidden.SelectedIndex = p.Hidden
                txtPropDescription.Text = p.Description
                txtPropHelp.Text = p.helpFile
                If p.returnType >= lstPropReturn.Items.Count Then
                    lstPropReturn.SelectedIndex = 0
                Else
                    lstPropReturn.SelectedIndex = p.returnType
                End If
                If p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray(0).Length > 2 Then
                    'returnArray(0) содержит имя класса, элементы котрого надо отобразить (аналог значения returnType = Location -> returnArray(0)="L")
                    Dim selClass As String = p.returnArray(0)
                    Dim selClassId As Integer = -1
                    If mScript.mainClassHash.TryGetValue(selClass, selClassId) Then
                        selClass = mScript.mainClass(selClassId).Names.Last
                        Dim res As Integer = lstPropElementClass.Items.IndexOf(selClass)
                        lstPropElementClass.SelectedIndex = res
                    Else
                        'variable
                        lstPropElementClass.SelectedIndex = lstPropElementClass.Items.IndexOf("Variable")
                    End If
                Else
                    lstPropElementClass.SelectedIndex = -1
                End If
                txtPropReturnArray.Clear()
                If IsNothing(p.returnArray) = False AndAlso p.returnType <> MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then
                    'заполняем текстбокс значениями, которые может возвращать свойство
                    For i As Integer = 0 To p.returnArray.GetUpperBound(0)
                        txtPropReturnArray.AppendText(p.returnArray(i) + vbNewLine)
                    Next
                End If
                Dim curProVal As String = p.Value
                Select Case lstPropReturn.SelectedIndex
                    Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                        rtbProp.Hide()
                        txtProp.Hide()
                        pnlDataType.Hide()
                        pnlPropBool.Show()
                        If curProVal = "True" Then
                            optPropValueTrue.Checked = True
                        Else
                            optPropValueFalse.Checked = True
                        End If
                    Case MatewScript.ReturnFunctionEnum.RETURN_EVENT, MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION
                        pnlDataType.Hide()
                        pnlPropBool.Hide()
                        txtProp.Hide()
                        rtbProp.Show()
                        If lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                            codeProp.codeBox.IsTextBlockByDefault = False
                        Else
                            codeProp.codeBox.IsTextBlockByDefault = True
                        End If
                        codeProp.codeBox.LoadCodeFromProperty(curProVal)
                        rtbProp.Rtf = codeProp.codeBox.Rtf
                    Case Else
                        pnlDataType.Show()
                        pnlPropBool.Hide()
                        If mScript.IsPropertyContainsCode(curProVal) <> MatewScript.ContainsCodeEnum.NOT_CODE Then
                            codeProp.codeBox.LoadCodeFromProperty(curProVal)
                            rtbProp.Rtf = codeProp.codeBox.Rtf
                            optValueCode.Checked = True
                        Else
                            txtProp.Text = curProVal
                            optValueNormal.Checked = True
                        End If
                End Select

                'отображаем форму
                pnlFunction.Hide()
                pnlClass.Hide()
                pnlProperty.Show()
                'меняем подписть кнопки "Удалить"
                btnRemove.Enabled = True
                btnRemove.Text = "Удалить свойство " + selectedElementName
                btnUp.Enabled = (e.Node.Index > 0)
                btnDn.Enabled = (e.Node.Index < e.Node.Parent.Nodes.Count - 1)
                FillPropParamArray(selectedElementName, False)
            Else
                'Выбранный узел - "Функции" / "Свойства"
                'пречем все формы редактирования
                selectedElementName = ""
                pnlFunction.Hide()
                pnlClass.Hide()
                pnlProperty.Hide()
                selectedClassId = mScript.mainClassHash(e.Node.Parent.Tag.ToString.Substring(6))
                'делаем недоступной кнопку "Удалить"
                btnRemove.Enabled = False
                btnRemove.Text = "Удалить ..."
                btnUp.Enabled = False
                btnDn.Enabled = False
            End If
        End If
    End Sub

    Private Sub treeClasses_DragDrop(ByVal sender As TreeView, ByVal e As System.Windows.Forms.DragEventArgs) Handles treeClasses.DragDrop
        'Организация DragDop - перемещение функций и свойств между классами
        For i As Integer = 0 To sender.Nodes.Count - 1
            'возвращаем цвет по-умолчанию всем узлам с классами
            If sender.Nodes(i).ForeColor <> defNodeColor Then sender.Nodes(i).ForeColor = defNodeColor
        Next
        If e.Data.GetDataPresent(GetType(TreeNode)) = False Then Return

        'Получаем узел класса, куда совершается перемещение
        Dim NewClassNode As TreeNode = Nothing
        If IsNothing(nodeUnderMouseToDrop.Tag) Then
            NewClassNode = nodeUnderMouseToDrop.Parent.Parent
        ElseIf nodeUnderMouseToDrop.Tag = "prop" OrElse nodeUnderMouseToDrop.Tag = "func" Then
            NewClassNode = nodeUnderMouseToDrop.Parent
        ElseIf nodeUnderMouseToDrop.Tag.ToString.StartsWith("Class") Then
            NewClassNode = nodeUnderMouseToDrop
        End If
        'получаем Id предыдущего и нового класса перемещаемого элемента
        Dim newClassId As Integer = mScript.mainClassHash(NewClassNode.Tag.ToString.Substring(6))
        Dim oldClassId As Integer = mScript.mainClassHash(nodeUnderMouseToDrag.Parent.Parent.Tag.ToString.Substring(6))

        If dragNodeType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION Then
            'перемещаем функцию
            'получаем инфо о функции из mainClass
            Dim prop As MatewScript.PropertiesInfoType = mScript.mainClass(oldClassId).Functions(nodeUnderMouseToDrag.Text)
            If mScript.mainClass(newClassId).Functions.ContainsKey(nodeUnderMouseToDrag.Text) Then
                'в новом классе уже есть одноименная функция
                Dim res As Microsoft.VisualBasic.MsgBoxResult = MsgBox("В классе " + mScript.mainClass(newClassId).Names(0) + " функция " + nodeUnderMouseToDrag.Text + " уже существует. Заменить ее на новую?", MsgBoxStyle.YesNo)
                If res = MsgBoxResult.No Then Return
                'сохраняем инфо в mainClass
                mScript.mainClass(newClassId).Functions(nodeUnderMouseToDrag.Text) = prop.Clone
                'выделяем узел с измененной функцией
                For i As Integer = 0 To NewClassNode.Nodes(0).Nodes.Count - 1
                    If NewClassNode.Nodes(0).Nodes(i).Text = nodeUnderMouseToDrag.Text Then
                        sender.SelectedNode = NewClassNode.Nodes(0).Nodes(i)
                        Exit For
                    End If
                Next
            Else
                AddUserFunctionByPropData(newClassId, nodeUnderMouseToDrag.Text, prop) 'сохраняем инфо в mainClass
                'mScript.mainClass(newClassId).Functions.Add(nodeUnderMouseToDrag.Text, prop) 
                sender.SelectedNode = NewClassNode.Nodes(0).Nodes.Add(nodeUnderMouseToDrag.Text) 'выделяем новый узел
            End If
            If e.Effect = DragDropEffects.Move Then
                'если операция - перемещение (а не копирование), то удаляем функцию из старого места
                'mScript.mainClass(oldClassId).Functions.Remove(nodeUnderMouseToDrag.Text)
                RemoveUserFunction({mScript.mainClass(oldClassId).Names(0), "'" & nodeUnderMouseToDrag.Text & "'"}, Nothing, True)
                sender.Nodes.Remove(nodeUnderMouseToDrag)
            End If
            'MsgBox("Функция " + nodeUnderMouseToDrag.Text + " перемещена в " + classNode.Text)
        Else
            'перемещаем свойство
            'получаем инфо о свойстве из mainClass
            Dim prop As MatewScript.PropertiesInfoType = mScript.mainClass(oldClassId).Properties(nodeUnderMouseToDrag.Text)
            If mScript.mainClass(newClassId).Properties.ContainsKey(nodeUnderMouseToDrag.Text) Then
                'в новом классе уже есть одноименное свойство
                Dim res As Microsoft.VisualBasic.MsgBoxResult = MsgBox("В классе " + mScript.mainClass(newClassId).Names(0) + " свойство " + nodeUnderMouseToDrag.Text + " уже существует. Заменить его на новое?", MsgBoxStyle.YesNo)
                If res = MsgBoxResult.No Then Return
                'сохраняем инфо в mainClass
                mScript.mainClass(newClassId).Properties(nodeUnderMouseToDrag.Text) = prop.Clone
                'выделяем узел с измененным свойством
                For i As Integer = 0 To NewClassNode.Nodes(0).Nodes.Count - 1
                    If NewClassNode.Nodes(1).Nodes(i).Text = nodeUnderMouseToDrag.Text Then
                        sender.SelectedNode = NewClassNode.Nodes(0).Nodes(i)
                        Exit For
                    End If
                Next
            Else
                AddUserPropertyByPropData(newClassId, nodeUnderMouseToDrag.Text, prop) 'сохраняем инфо в mainClass
                'mScript.mainClass(newClassId).Properties.Add(nodeUnderMouseToDrag.Text, prop.Clone)
                sender.SelectedNode = NewClassNode.Nodes(1).Nodes.Add(nodeUnderMouseToDrag.Text) 'выделяем новый узел
            End If
            If e.Effect = DragDropEffects.Move Then
                'если операция - перемещение (а не копирование), то удаляем свойство из старого места
                'mScript.mainClass(oldClassId).Properties.Remove(nodeUnderMouseToDrag.Text)
                RemoveUserProperty({mScript.mainClass(oldClassId).Names(0), "'" & nodeUnderMouseToDrag.Text & "'"}, Nothing, True)
                sender.Nodes.Remove(nodeUnderMouseToDrag)
            End If
            'MsgBox("Свойство " + nodeUnderMouseToDrag.Text + " перемещено в " + NewClassNode.Text)
        End If
        'http://msdn.microsoft.com/ru-ru/library/system.windows.forms.control.dodragdrop(v=VS.90).aspx#Y1680
    End Sub

    Private Sub treeClasses_DragOver(ByVal sender As TreeView, ByVal e As System.Windows.Forms.DragEventArgs) Handles treeClasses.DragOver
        'Организация DragDop - перемещение функций и свойств между классами
        If IsNothing(nodeUnderMouseToDrag) Then Return

        'Если нажат Ctrl - копирование; если Shift или ничего - перемещение
        If ((e.KeyState And 4) = 4 And _
            (e.AllowedEffect And DragDropEffects.Move) = DragDropEffects.Move) Then
            ' SHIFT KeyState for move.
            e.Effect = DragDropEffects.Move
        ElseIf ((e.KeyState And 8) = 8 And _
            (e.AllowedEffect And DragDropEffects.Copy) = DragDropEffects.Copy) Then
            ' CTL KeyState for copy.
            e.Effect = DragDropEffects.Copy
        ElseIf ((e.AllowedEffect And DragDropEffects.Move) = DragDropEffects.Move) Then
            ' By default, the drop action should be move, if allowed.
            e.Effect = DragDropEffects.Move
        Else
            e.Effect = DragDropEffects.None
        End If

        'получаем узел, который под мышью на данный момент
        nodeUnderMouseToDrop = sender.GetNodeAt(sender.PointToClient(New Point(e.X, e.Y)))
        If IsNothing(nodeUnderMouseToDrop) Then
            'если под мышью нет узла - действие запрещено
            e.Effect = DragDropEffects.None
            'возвращаем цвет по-умолчанию всем узлам с классами
            For i As Integer = 0 To sender.Nodes.Count - 1
                If sender.Nodes(i).ForeColor <> defNodeColor Then sender.Nodes(i).ForeColor = defNodeColor
            Next
            Exit Sub
        End If
        'If e.Effect = DragDropEffects.Move OrElse e.Effect = DragDropEffects.Copy Then
        '    If sender.PointToClient(New Point(e.X, e.Y)).Y + sender.ItemHeight > sender.Height Then e.Effect = DragDropEffects.Scroll
        'End If

        'Если класс элемента, под которым мышь, тот же, что и у перемещаемого элемента - операция запрещена (копировать можно только в другой класс)
        If IsNothing(nodeUnderMouseToDrop.Tag) Then
            If nodeUnderMouseToDrop.Parent.Parent.Tag = nodeUnderMouseToDrag.Parent.Parent.Tag Then
                e.Effect = DragDropEffects.None
            End If
        ElseIf nodeUnderMouseToDrop.Tag = "prop" OrElse nodeUnderMouseToDrop.Tag = "func" Then
            If nodeUnderMouseToDrop.Parent.Tag = nodeUnderMouseToDrag.Parent.Parent.Tag Then
                e.Effect = DragDropEffects.None
            End If
        ElseIf nodeUnderMouseToDrop.Tag.ToString.StartsWith("Class") Then
            If nodeUnderMouseToDrop.Tag = nodeUnderMouseToDrag.Parent.Parent.Tag Then
                e.Effect = DragDropEffects.None
            End If
        Else
            e.Effect = DragDropEffects.None
        End If

        If e.Effect = DragDropEffects.None Then
            'операция запрещена
            For i As Integer = 0 To sender.Nodes.Count - 1
                'возвращаем цвет по-умолчанию всем узлам с классами
                If sender.Nodes(i).ForeColor <> defNodeColor Then sender.Nodes(i).ForeColor = defNodeColor
            Next
        Else
            'операция разрешена
            'получаем узел класса, куда может совершиться перемещение
            Dim classNodeToDrop As TreeNode
            If IsNothing(nodeUnderMouseToDrop.Tag) Then
                classNodeToDrop = nodeUnderMouseToDrop.Parent.Parent
            ElseIf nodeUnderMouseToDrop.Tag.ToString.StartsWith("Class") Then
                classNodeToDrop = nodeUnderMouseToDrop
            ElseIf nodeUnderMouseToDrop.Tag = "func" OrElse nodeUnderMouseToDrop.Tag = "prop" Then
                classNodeToDrop = nodeUnderMouseToDrop.Parent
            Else
                Return
            End If
            'возвращаем цвет по-умолчанию всем узлам с классами, кроме того, куда может совершиться перемещение. Его закрашиваем в nodeSelColor
            For i As Integer = 0 To sender.Nodes.Count - 1
                If sender.Nodes(i).Equals(classNodeToDrop) Then
                    If sender.Nodes(i).ForeColor <> nodeSelColor Then sender.Nodes(i).ForeColor = nodeSelColor
                Else
                    If sender.Nodes(i).ForeColor <> defNodeColor Then sender.Nodes(i).ForeColor = defNodeColor
                End If
            Next
        End If
    End Sub

    Private Sub treeClasses_MouseDown(ByVal sender As TreeView, ByVal e As System.Windows.Forms.MouseEventArgs) Handles treeClasses.MouseDown
        'Организация DragDop - перемещение функций и свойств между классами
        nodeUnderMouseToDrag = sender.GetNodeAt(e.X, e.Y) 'получаем перемещаемый узел
        If IsNothing(nodeUnderMouseToDrag) Then
            dragBoxFromMouseDown = Rectangle.Empty
            Exit Sub
        End If
        If IsNothing(nodeUnderMouseToDrag.Tag) = False Then
            'перемещать можно только узлы со свойствами и функциями, которые не имеют тэгов
            dragBoxFromMouseDown = Rectangle.Empty
            Exit Sub
        End If
        'получаем тип перемещаемого элемента
        If nodeUnderMouseToDrag.Parent.Tag = "func" Then
            dragNodeType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION 'функция
        Else
            dragNodeType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_PROPERTY 'свойство
        End If
        'получаем квадрат, при выходе за пределы которого начинается операция DragDrop
        Dim dragSize As Size = SystemInformation.DragSize
        dragBoxFromMouseDown = New Rectangle(New Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize)

    End Sub

    Private Sub treeClasses_MouseMove(ByVal sender As TreeView, ByVal e As System.Windows.Forms.MouseEventArgs) Handles treeClasses.MouseMove
        'Организация DragDop - перемещение функций и свойств между классами
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If IsNothing(nodeUnderMouseToDrag) Then Exit Sub
        If dragBoxFromMouseDown.Equals(Rectangle.Empty) = False And dragBoxFromMouseDown.Contains(e.X, e.Y) = False Then
            If sender.SelectedNode.Equals(nodeUnderMouseToDrag) = False Then
                sender.SelectedNode = nodeUnderMouseToDrag 'выделяем перемещаемый узел в дереве классов
            End If
            'начинаем DragDrop
            Dim dropEffect As DragDropEffects = sender.DoDragDrop(nodeUnderMouseToDrag, DragDropEffects.Move + DragDropEffects.Copy + DragDropEffects.Scroll)
        End If
    End Sub

    Private Sub treeClasses_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles treeClasses.MouseUp
        'Организация DragDop - перемещение функций и свойств между классами
        'при отпучкании кнопки мыши прекращаем перетаскивание
        dragBoxFromMouseDown = Rectangle.Empty
    End Sub

    Private Sub treeClasses_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs) Handles treeClasses.QueryContinueDrag
        'Организация DragDop - перемещение функций и свойств между классами
        If e.EscapePressed Then
            'При нажатии Esc прекращаем перетаскивание
            e.Action = DragAction.Cancel
            For i As Integer = 0 To sender.Nodes.Count - 1
                'возвращаем цвет по-умолчанию всем узлам с классами
                If sender.Nodes(i).ForeColor <> defNodeColor Then sender.Nodes(i).ForeColor = defNodeColor
            Next
            'стираем данные, необходимые для перетаскивания
            nodeUnderMouseToDrag = Nothing
            dragBoxFromMouseDown = Rectangle.Empty
        End If
    End Sub
#End Region

    Private Sub btnNewClass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewClass.Click
        'Отображаем форму для создания нового класса
        curAction = SavingClassActionEnum.NEW_ELEMENT
        txtNewClassName.Clear()
        nudClassLevels.Value = 3
        pnlFunction.Hide()
        pnlProperty.Hide()
        pnlClass.Show()
    End Sub

    Private Sub blnCreateNewClass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateNewClass.Click
        'Создаем новый или редактируем существующий класс
        Dim strNames As String = txtNewClassName.Text.Trim
        If strNames.Length = 0 Then
            MsgBox("Укажите хоть одно имя класса.", MsgBoxStyle.Exclamation)
            txtNewClassName.Focus()
            Exit Sub
        End If
        'Получаем Id класса
        Dim classId As Integer = -1, oldfirstName As String = ""
        If curAction = SavingClassActionEnum.EDIT_ELEMENT Then
            classId = selectedClassId
            oldfirstName = mScript.mainClass(classId).Names(0)
        Else
            classId = mScript.mainClass.Length
        End If
        'Заполняем arrNames уменами класса
        Dim arrNames() As String = Split(strNames, ",")
        For i As Integer = 0 To arrNames.GetUpperBound(0)
            arrNames(i) = arrNames(i).Trim
            If arrNames(i).Length = 0 Then
                MsgBox("Уберите лишнюю запятую в именах класса.", MsgBoxStyle.Exclamation)
                txtNewClassName.Focus()
                Exit Sub
            End If
            If mScript.mainClassHash.ContainsKey(arrNames(i)) Then
                If mScript.mainClassHash(arrNames(i)) <> classId Then
                    MsgBox("Имя " + arrNames(i) + " уже существует.", MsgBoxStyle.Exclamation)
                    txtNewClassName.Focus()
                    Exit Sub
                End If
            End If
        Next
        'сохраняем имена класса в структуре mainClass
        If curAction = SavingClassActionEnum.NEW_ELEMENT Then
            ReDim Preserve mScript.mainClass(classId)
            mScript.mainClass(classId) = New MatewScript.MainClassType
            If Not DEEP_EDIT_MODE Then
                mScript.mainClass(classId).UserAdded = True
            End If
        Else
            If DEEP_EDIT_MODE = False AndAlso mScript.mainClass(classId).UserAdded = False Then
                'редактирование встроенного класса
                If mScript.mainClass(classId).LevelsCount <> nudClassLevels.Value - 1 Then
                    If MessageBox.Show("Изменение количества уровней встроенного класса может привести к нестабильности работы движка или даже к полной потере работоспособности. Продолжить?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.No Then Return
                End If
                If Join(mScript.mainClass(classId).Names, ", ") <> txtNewClassName.Text.Trim Then
                    If MessageBox.Show("Имена встроенного класса были изменены. Допускается добавление нового имени, но удаление существуещего приведет к потере работоспособности. Продолжить?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.No Then Return
                End If
            End If
        End If
        mScript.mainClass(classId).Names = arrNames
        mScript.mainClass(classId).HelpFile = txtClassHelpFile.Text
        mScript.mainClass(classId).DefaultProperty = cmbClassDefProperty.Text

        If curAction = SavingClassActionEnum.NEW_ELEMENT Then
            'заполняем данные о новом классе. Классам 2 и 3 уровня добавляем набор базовых функций
            mScript.mainClass(classId).LevelsCount = nudClassLevels.Value - 1
            If nudClassLevels.Value = 1 Then
                mScript.mainClass(classId).Functions = New SortedList(Of String, MatewScript.PropertiesInfoType)(mScript.basicFunctionsHashLevel1)
            ElseIf nudClassLevels.Value = 2 Then
                mScript.mainClass(classId).Functions = New SortedList(Of String, MatewScript.PropertiesInfoType)(mScript.basicFunctionsHashLevel2)
            Else
                mScript.mainClass(classId).Functions = New SortedList(Of String, MatewScript.PropertiesInfoType)(mScript.basicFunctionsHashLevel3)
            End If
            mScript.mainClass(classId).Properties = New SortedList(Of String, MatewScript.PropertiesInfoType)(StringComparer.CurrentCultureIgnoreCase)
            'Классам 2 и 3 уровня создаем свойство Name
            If nudClassLevels.Value > 1 Then mScript.mainClass(classId).Properties.Add("Name", New MatewScript.PropertiesInfoType With {.Value = "'default'", .Description = "Имя элемента", _
                    .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR, .editorIndex = 0})
            mScript.mainClass(classId).Properties.Add("Group", New MatewScript.PropertiesInfoType With {.Value = "", .Description = "[СЛУЖЕБНОЕ] Для хранения группы элемента", _
                                                                                                        .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL, .editorIndex = 1})
            mScript.mainClass(classId).Properties.Add("Icon", New MatewScript.PropertiesInfoType With {.Value = "", .Description = "[СЛУЖЕБНОЕ] Для хранения иконки элемента", _
                                                                                                       .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL, .editorIndex = 2})
        Else
            'Редактирование класса
            If mScript.mainClass(classId).LevelsCount <> nudClassLevels.Value - 1 Then
                'уровень класса изменился
                If mScript.mainClass(classId).LevelsCount = 0 Then
                    'Уровень класса с первого стал 2 или 3
                    Dim res As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Количество уровней класса поменялось. Добавить в класс стандартные свойства и функции для класса нового уровня?", MsgBoxStyle.YesNo)
                    If res = MsgBoxResult.Yes Then
                        'заполняем класса бызовыми свойствами и функциями
                        mScript.mainClass(classId).LevelsCount = nudClassLevels.Value - 1
                        If nudClassLevels.Value = 2 Then
                            'уровень стал 2
                            For i As Integer = 0 To mScript.basicFunctionsHashLevel2.Count - 1
                                Dim func As MatewScript.PropertiesInfoType = mScript.basicFunctionsHashLevel2.ElementAt(i).Value.Clone
                                Dim funcName As String = mScript.basicFunctionsHashLevel2.ElementAt(i).Key
                                If mScript.mainClass(classId).Functions.ContainsKey(funcName) = False Then
                                    func.editorIndex = mScript.mainClass(classId).Functions.Count
                                    mScript.mainClass(classId).Functions.Add(funcName, func)
                                End If
                            Next
                            If mScript.mainClass(classId).Properties.ContainsKey("Name") = False Then
                                mScript.mainClass(classId).Properties.Add("Name", New MatewScript.PropertiesInfoType With {.Value = "'default'", .Description = "Имя элемента", _
                                                                          .editorIndex = mScript.mainClass(classId).Properties.Count, .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR})
                            End If
                        Else
                            'уровень стал 3
                            For i As Integer = 0 To mScript.basicFunctionsHashLevel3.Count - 1
                                Dim func As MatewScript.PropertiesInfoType = mScript.basicFunctionsHashLevel3.ElementAt(i).Value.Clone
                                Dim funcName As String = mScript.basicFunctionsHashLevel3.ElementAt(i).Key
                                If mScript.mainClass(classId).Functions.ContainsKey(funcName) = False Then
                                    func.editorIndex = mScript.mainClass(classId).Functions.Count
                                    mScript.mainClass(classId).Functions.Add(funcName, func)
                                End If
                            Next
                            If mScript.mainClass(classId).Properties.ContainsKey("Name") = False Then
                                mScript.mainClass(classId).Properties.Add("Name", New MatewScript.PropertiesInfoType With {.Value = "'default'", .Description = "Имя элемента", _
                                                                          .editorIndex = mScript.mainClass(classId).Properties.Count, .Hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR})
                            End If
                        End If
                    End If
                ElseIf nudClassLevels.Value = 1 Then
                    'уровень класса с 2 или 3 стал первым
                    Dim res As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Количество уровней класса поменялось. Удалить из класса стандартные свойства и функции, не нужные для класса 1 уровня?", MsgBoxStyle.YesNo)
                    If res = MsgBoxResult.Yes Then
                        'удаляем функции и свойства, не нужные классу 1 уровня
                        mScript.mainClass(classId).LevelsCount = nudClassLevels.Value - 1
                        For Each func As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In IIf(nudClassLevels.Value = 2, mScript.basicFunctionsHashLevel2, mScript.basicFunctionsHashLevel3)
                            If mScript.mainClass(classId).Functions.ContainsKey(func.Key) Then
                                mScript.mainClass(classId).Functions.Remove(func.Key)
                            End If
                        Next
                        If mScript.mainClass(classId).Properties.ContainsKey("Name") Then
                            Dim edIndex As Integer = mScript.mainClass(classId).Properties("Name").editorIndex
                            mScript.mainClass(classId).Properties.Remove("Name")
                            If IsNothing(mScript.mainClass(classId).Properties) = False Then
                                For i As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
                                    Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties.ElementAt(i).Value
                                    Dim pName As String = mScript.mainClass(classId).Properties.ElementAt(i).Key
                                    If p.editorIndex > edIndex Then
                                        p.editorIndex -= 1
                                    End If
                                Next i
                            End If
                        End If
                    End If
                    'Добавляем функции 1 класса
                    For Each func As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.basicFunctionsHashLevel1
                        If mScript.mainClass(classId).Functions.ContainsKey(func.Key) = False Then
                            mScript.mainClass(classId).Functions.Add(func.Key, func.Value.Clone)
                        End If
                    Next
                ElseIf nudClassLevels.Value = 2 Then
                    'Уровень класса с третьего стал вторым
                    Dim res As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Количество уровней класса поменялось. Заменить стандартные свойства и функции 3 уровня на соответствующие им 2 уровня?", MsgBoxStyle.YesNo)
                    If res = MsgBoxResult.Yes Then
                        'заменяем базовыем функции третьего класса на аналогичные им второго
                        mScript.mainClass(classId).LevelsCount = nudClassLevels.Value - 1
                        For Each func As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.basicFunctionsHashLevel2
                            If mScript.mainClass(classId).Functions.ContainsKey(func.Key) Then
                                mScript.mainClass(classId).Functions(func.Key) = func.Value.Clone
                            End If
                        Next
                    End If
                Else
                    'Уровень класса со второго стал третьим
                    Dim res As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Количество уровней класса поменялось. Заменить стандартные свойства и функции 2 уровня на соответствующие им 3 уровня?", MsgBoxStyle.YesNo)
                    If res = MsgBoxResult.Yes Then
                        'заменяем базовыем функции второго класса на аналогичные им третьего
                        mScript.mainClass(classId).LevelsCount = nudClassLevels.Value - 1
                        For Each func As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.basicFunctionsHashLevel3
                            If mScript.mainClass(classId).Functions.ContainsKey(func.Key) Then
                                mScript.mainClass(classId).Functions(func.Key) = func.Value.Clone
                            End If
                        Next
                    End If
                End If
            End If
        End If
        mScript.MakeMainClassHash() 'обновляем хэш с именами классов

        If curAction = SavingClassActionEnum.EDIT_ELEMENT Then
            'меням текст и тэг узла редактируемого класса
            treeClasses.SelectedNode.Text = "Класс " + Join(mScript.mainClass(selectedClassId).Names, ", ")
            treeClasses.SelectedNode.Tag = "Class " + mScript.mainClass(selectedClassId).Names(0)
            FillTreeViewWithClasses(treeClasses)

            Dim q As New EditQueueClass With {.ActionType = EditQueueClass.EditActionTypeEnum.EDIT_CLASS_NAME, .classId = selectedClassId, .newValue = mScript.mainClass(selectedClassId).Names(0), .oldValue = oldfirstName}
            lstQueue.Add(q)
        Else
            'добавляем в дерево классов все элементы нового класса
            selectedClassId = mScript.mainClass.GetUpperBound(0)
            Dim newNode As TreeNode = treeClasses.Nodes.Add("Класс " + Join(mScript.mainClass(selectedClassId).Names, ", "))
            newNode.Tag = "Class " + mScript.mainClass(selectedClassId).Names(0)
            Dim treeFunctions As TreeNode = newNode.Nodes.Add("Функции")
            treeFunctions.Tag = "func"
            For Each prop As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(selectedClassId).Functions
                treeFunctions.Nodes.Add(prop.Key)
            Next

            Dim treeProperties As TreeNode = newNode.Nodes.Add("Свойства")
            treeProperties.Tag = "prop"
            For Each prop As KeyValuePair(Of String, MatewScript.PropertiesInfoType) In mScript.mainClass(selectedClassId).Properties
                treeProperties.Nodes.Add(prop.Key)
            Next
            treeClasses.SelectedNode = newNode

            Dim q As New EditQueueClass With {.ActionType = EditQueueClass.EditActionTypeEnum.CREATE_CLASS, .classId = selectedClassId, .newValue = mScript.mainClass(selectedClassId).Names(0), .oldValue = ""}
            lstQueue.Add(q)

        End If

        'очищаем поля ввода и прячем форму
        nudClassLevels.Value = 3
        txtNewClassName.Text = ""
        txtClassHelpFile.Text = ""
        cmbClassDefProperty.Text = ""
        cmbClassDefProperty.Items.Clear()
        pnlClass.Hide()

        FillListWithClasses()
        mScript.FillFuncAndPropHash()
    End Sub

    Private Sub btnNewFunction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewFunction.Click
        'отображаем форму для создания новоой функции
        pnlClass.Hide()
        pnlProperty.Hide()
        pnlFunction.Show()
        curAction = SavingClassActionEnum.NEW_ELEMENT
        lstFuncReturnType.SelectedIndex = 0
        'очищаем все от старых данных
        txtFuncName.Clear()
        txtFuncDescription.Clear()
        txtFuncEditorName.Clear()
        cmbFuncHidden.SelectedIndex = 0
        txtFuncHelpFile.Clear()
        txtFuncResult.Clear()
        nudFuncMin.Value = 0
        nudFuncMax.Value = 0
        lstFuncParams.Items.Clear()
        lstFuncReturnType.SelectedIndex = 0
        Erase curFuncParams 'очищаем структуру curFuncParams
        txtFuncName.Focus()
    End Sub

    Private Sub lstFuncReturnType_SelectedIndexChanged(ByVal sender As ListBox, ByVal e As System.EventArgs) Handles lstFuncReturnType.SelectedIndexChanged
        'Если функция возвращает один из набора вариантов, то показываем поля для ввода этих вариантов. Иначе это поле прячем
        Dim blnShow As Boolean = IIf(sender.SelectedIndex = 2, True, False)
        txtFuncResult.Visible = blnShow
        lblFuncResult.Visible = blnShow
    End Sub

    Private Sub nudFuncMax_ValueChanged(ByVal sender As NumericUpDown, ByVal e As System.EventArgs) Handles nudFuncMax.ValueChanged
        'изменилось значение максимального кол-ва параметров функции
        pnlFuncParams.Visible = IIf(sender.Value = 0, False, True)
        'если минимальное кол-во параметров оказалось больше максимального - делаем их равными
        If sender.Value < nudFuncMin.Value Then nudFuncMin.Value = sender.Value
        If lstFuncParams.Items.Count < sender.Value Then
            'число увеличилось - добавляем параметр
            AddFuncParam()
            lstFuncParams.SelectedIndex = lstFuncParams.Items.Count - 1
        ElseIf lstFuncParams.Items.Count > sender.Value Then
            'число уменьшилось - удаляем параметр
            RemoveFuncParam()
            If lstFuncParams.Items.Count > 0 Then lstFuncParams.SelectedIndex = 0
        End If
    End Sub

    Private Sub nudFuncMin_ValueChanged(ByVal sender As NumericUpDown, ByVal e As System.EventArgs) Handles nudFuncMin.ValueChanged
        'если максимальное кол-во параметров функциии стало меньше минимального - делаем их равными
        If nudFuncMax.Value < sender.Value Then nudFuncMax.Value = sender.Value
    End Sub

    Private Sub AddFuncParam()
        'Добавляем функции новый параметр
        Dim parUBound As Integer = 0
        Dim defParName As String = "Param1"
        If IsNothing(curFuncParams) = False AndAlso curFuncParams.GetUpperBound(0) > -1 Then
            'параметры функции уже существуют
            'В defParName получаем имя по-умолчанию для нового параметра, вида "ParamX", где х - любое число, такое, чтобы имя параметра было уникальным
            parUBound = curFuncParams.Length
            Do
                For i = 0 To parUBound - 1
                    If curFuncParams(i).Name = defParName Then
                        defParName = defParName.Substring(0, 5) + Convert.ToString(Convert.ToInt32(defParName.Substring(5)) + 1)
                        Continue Do
                    End If
                Next
                Exit Do
            Loop
        End If
        'Создаем параматр в структуре curFuncParams со значениями по-умолчанию
        ReDim Preserve curFuncParams(parUBound)
        curFuncParams(parUBound) = New MatewScript.paramsType
        curFuncParams(parUBound).Name = defParName
        curFuncParams(parUBound).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ANY
        'добавляем параметр в список lstFuncParams
        lstFuncParams.Items.Add(defParName)
    End Sub

    Private Sub RemoveFuncParam(Optional ByVal paramIndex As Integer = -1)
        'Удаляет параметр функции
        If paramIndex = -1 Then paramIndex = lstFuncParams.Items.Count - 1
        If paramIndex = -1 Then Exit Sub
        'удаляем параметр из списка
        lstFuncParams.Items.Remove(lstFuncParams.Items(paramIndex))

        If IsNothing(curFuncParams) OrElse curFuncParams.GetUpperBound(0) = 0 Then
            Erase curFuncParams 'параметр был последним - очищаем curFuncParams и выход 
            Exit Sub
        End If
        'удаляем параметр из структуры curFuncParams
        For i = paramIndex To curFuncParams.GetUpperBound(0) - 1
            curFuncParams(i) = curFuncParams(i + 1)
        Next
        ReDim Preserve curFuncParams(curFuncParams.GetUpperBound(0) - 1)
    End Sub

    Private Sub btnParamAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParamAdd.Click
        'увеличение счетчика nudFuncMax приводит к созданию нового параметра функции
        nudFuncMax.Value += 1
    End Sub

    Private Sub btnParamRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParamRemove.Click
        'удаляем выбранный в lstFuncParams параметр функции
        RemoveFuncParam(lstFuncParams.SelectedIndex)
        If nudFuncMax.Value > lstFuncParams.Items.Count Then nudFuncMax.Value = lstFuncParams.Items.Count
        If lstFuncParams.Items.Count > 0 Then lstFuncParams.SelectedIndex = 0
    End Sub

    Private Sub lstFuncParams_SelectedIndexChanged(ByVal sender As ListBox, ByVal e As System.EventArgs) Handles lstFuncParams.SelectedIndexChanged
        'При выборе в lstFuncParams нового элемента выводим в соответствующих контролах инфу о параметре для ее редактирования
        'в зависимости от того, выбран ли первый или последний параметр, меняем доступность кнопок, меняющих позицию параметра
        Dim selIndex As Integer = sender.SelectedIndex
        If selIndex = 0 Then
            btnParamUp.Enabled = False
        Else
            btnParamUp.Enabled = True
        End If
        If selIndex = sender.Items.Count - 1 Then
            btnParamDown.Enabled = False
        Else
            btnParamDown.Enabled = True
        End If
        If selIndex = -1 Then Exit Sub
        'заполняем соответствующие поля данными о параметре из стр-ры curFuncParams
        txtParamName.Text = curFuncParams(selIndex).Name
        txtParamDescription.Text = curFuncParams(selIndex).Description
        Dim parType As MatewScript.paramsType.paramsTypeEnum = curFuncParams(selIndex).Type
        If cmbParamType.Items.Count - 1 >= parType Then
            cmbParamType.SelectedIndex = parType
        Else
            cmbParamType.SelectedIndex = 0
        End If
        If IsNothing(curFuncParams(selIndex).EnumValues) Then
            txtParamsEnum.Clear()
            lstParamsClass.SelectedIndex = -1
        Else
            If curFuncParams(selIndex).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT AndAlso curFuncParams(selIndex).EnumValues(0).Length > 2 Then
                'если параметр является элементом, то выбираем значение из набора
                'txtParamsEnum.Clear()
                Dim selClass As String = curFuncParams(selIndex).EnumValues(0)
                Dim selClassId As Integer = -1
                If mScript.mainClassHash.TryGetValue(selClass, selClassId) Then
                    selClass = mScript.mainClass(selClassId).Names.Last
                    Dim res As Integer = lstParamsClass.Items.IndexOf(selClass)
                    lstParamsClass.SelectedIndex = res
                Else
                    lstParamsClass.SelectedIndex = -1
                End If
            Else
                'если параметр принимает одно из набора возможных значений, заполняем txtParamsEnum этими значениями
                Dim sBuilder As New System.Text.StringBuilder
                For i = 0 To curFuncParams(selIndex).EnumValues.GetUpperBound(0)
                    sBuilder.AppendLine(curFuncParams(selIndex).EnumValues(i))
                Next
                txtParamsEnum.Text = sBuilder.ToString
                lstParamsClass.SelectedIndex = -1
            End If
        End If
    End Sub

    Private Sub btnParamUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParamUp.Click
        'Перемещаем параметр функции на 1 вверх (будет идти раньше)
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex < 1 Then Exit Sub
        'В curFuncParams меняем местами выбранный параметр и предыдущий
        Dim paramCopy As MatewScript.paramsType = curFuncParams(selIndex - 1)
        curFuncParams(selIndex - 1) = curFuncParams(selIndex)
        curFuncParams(selIndex) = paramCopy
        'Таким же образом меняем местами параметры в списке lstFuncParams
        Dim strItemCopy As String = lstFuncParams.GetItemText(lstFuncParams.Items(selIndex - 1))
        lstFuncParams.Items(selIndex - 1) = lstFuncParams.Items(selIndex)
        lstFuncParams.Items(selIndex) = strItemCopy
        lstFuncParams.SelectedIndex = selIndex - 1
    End Sub

    Private Sub btnParamDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParamDown.Click
        'Перемещаем параметр функции на 1 вниз (будет идти позже)
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        If selIndex = lstFuncParams.Items.Count - 1 Then Exit Sub
        'В curFuncParams меняем местами выбранный параметр и последующий
        Dim paramCopy As MatewScript.paramsType = curFuncParams(selIndex + 1)
        curFuncParams(selIndex + 1) = curFuncParams(selIndex)
        curFuncParams(selIndex) = paramCopy
        'Таким же образом меняем местами параметры в списке lstFuncParams
        Dim strItemCopy As String = lstFuncParams.GetItemText(lstFuncParams.Items(selIndex + 1))
        lstFuncParams.Items(selIndex + 1) = lstFuncParams.Items(selIndex)
        lstFuncParams.Items(selIndex) = strItemCopy
        lstFuncParams.SelectedIndex = selIndex + 1
    End Sub

    Private Sub txtParamName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtParamName.TextChanged
        'Сохраняем имя параметра функции
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        curFuncParams(selIndex).Name = sender.Text
        lstFuncParams.Items(selIndex) = sender.Text 'обновляем имя параметра в списке lstFuncParams
    End Sub

    Private Sub txtParamDescription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtParamDescription.TextChanged
        'Сохраняем описание параметра функции
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        curFuncParams(selIndex).Description = sender.Text
    End Sub

    Private Sub cmbParamType_SelectedIndexChanged(ByVal sender As ComboBox, ByVal e As System.EventArgs) Handles cmbParamType.SelectedIndexChanged
        'Сохраняем тип параметра функции
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        curFuncParams(selIndex).Type = sender.SelectedIndex
        'если параметр возвращает одно из набора значений, показываем txtParamsEnum для ввода туда этих значений. Иначе его прячем.
        lblParamEnum.Visible = (curFuncParams(selIndex).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM)
        txtParamsEnum.Visible = lblParamEnum.Visible
        lblParamsClass.Visible = (curFuncParams(selIndex).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT)
        lstParamsClass.Visible = lblParamsClass.Visible
        If curFuncParams(selIndex).Type <> MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT AndAlso curFuncParams(selIndex).Type <> MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM Then
            curFuncParams(selIndex).EnumValues = Nothing
        End If

        'Если тип хоть одного - это массив параметров, то счетчик максимального кол-ва параметров делаем недоступным
        For i As Integer = 0 To curFuncParams.GetUpperBound(0)
            If curFuncParams(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY Then
                nudFuncMax.Enabled = False
                Exit Sub
            End If
        Next
        nudFuncMax.Enabled = True
    End Sub

    Private Sub txtParamsEnum_TextChanged(ByVal sender As TextBox, ByVal e As System.EventArgs) Handles txtParamsEnum.TextChanged
        'Сохраняем набор возможных значений параметра функции
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        Erase curFuncParams(selIndex).EnumValues 'очищаем предыдущие данные в стр-ре curFuncParams
        If sender.Text.Trim.Length = 0 Then Exit Sub

        'Приводим варианты значений в правильный вид и сохраняем в curFuncParams
        Dim strCurLine As String
        Dim parEnumUBound As Integer = -1
        For i As Integer = 0 To sender.Lines.GetUpperBound(0)
            strCurLine = sender.Lines(i).Trim
            If strCurLine.Length = 0 Then Continue For
            parEnumUBound += 1
            ReDim Preserve curFuncParams(selIndex).EnumValues(parEnumUBound)

            If Double.TryParse(strCurLine, Globalization.NumberStyles.Any, provider_points, Nothing) Then
                curFuncParams(selIndex).EnumValues(parEnumUBound) = strCurLine
            Else
                If strCurLine.StartsWith("'") And strCurLine.EndsWith("'") And strCurLine.Length > 1 Then
                    curFuncParams(selIndex).EnumValues(parEnumUBound) = strCurLine
                Else
                    If strCurLine.ToLower = "true" Then
                        curFuncParams(selIndex).EnumValues(parEnumUBound) = "True"
                    ElseIf strCurLine.ToLower = "false" Then
                        curFuncParams(selIndex).EnumValues(parEnumUBound) = "False"
                    Else
                        curFuncParams(selIndex).EnumValues(parEnumUBound) = "'" + strCurLine.Replace("'", "/'") + "'"
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub btnOpenFuncHelpFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenFuncHelpFile.Click, btnOpenPropHelpFile.Click, btnClassHelpFile.Click
        'получение имени файла помощи для класса
        Dim initDir As String
        If DEEP_EDIT_MODE Then
            initDir = APP_HELP_PATH
        Else
            initDir = My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "help") + "\"
        End If
        ofd.InitialDirectory = initDir
        ofd.ShowDialog()
        If ofd.FileName.Length = 0 Then Exit Sub
        Dim strFileName = ofd.FileName
        Dim questInitDir As String = My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "help") + "\"
        If strFileName.StartsWith(APP_HELP_PATH, StringComparison.CurrentCultureIgnoreCase) Then
            strFileName = strFileName.Substring(APP_HELP_PATH.Length)
        ElseIf strFileName.StartsWith(questInitDir, StringComparison.CurrentCultureIgnoreCase) Then
            strFileName = strFileName.Substring(questInitDir.Length)
        End If
        If sender.name = "btnOpenFuncHelpFile" Then
            txtFuncHelpFile.Text = strFileName
        ElseIf sender.name = "btnClassHelpFile" Then
            txtClassHelpFile.Text = strFileName
        Else
            txtPropHelp.Text = strFileName
        End If
    End Sub

    Private Sub cmbParamType_Validating(ByVal sender As ComboBox, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmbParamType.Validating
        If sender.SelectedIndex = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY Then
            If lstFuncParams.SelectedIndex <> lstFuncParams.Items.Count - 1 Then
                'Запрет установки типа параметра функции = PARAMS_ARRAY любому параметру, если он не последний
                MsgBox("Этот тип может иметь только последний параметр.", MsgBoxStyle.Exclamation)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub btnSaveFunction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveFunction.Click
        'Создание новой или редактирование существующей функции
        'проверка корректности данных
        Dim funcName As String = txtFuncName.Text.Trim
        If funcName.Length = 0 Then
            MsgBox("Укажите имя функции.", MsgBoxStyle.Exclamation)
            txtFuncName.Focus()
            Exit Sub
        End If

        If mScript.mainClass(selectedClassId).Functions.ContainsKey(funcName) Then
            If (funcName = selectedElementName And curAction = SavingClassActionEnum.EDIT_ELEMENT) = False Then
                MsgBox("Функция " + funcName + " уже существует. Укажите другое имя.", MsgBoxStyle.Exclamation)
                txtFuncName.Focus()
                Exit Sub
            End If
        End If

        If curAction = SavingClassActionEnum.EDIT_ELEMENT AndAlso DEEP_EDIT_MODE = False Then
            Dim f As MatewScript.PropertiesInfoType = mScript.mainClass(selectedClassId).Functions(selectedElementName)
            If IsNothing(f) = False AndAlso f.UserAdded = False Then
                'редактирование встроенного функции
                If MessageBox.Show("Редактирование встроенной функции. Изменение имени функции, количества и типа параметров, а также результата функции может привести к неработоспособности функции. Продолжить?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.No Then Return
            End If
        End If

        Dim funcEditorName As String = txtFuncEditorName.Text.Trim
        If String.IsNullOrEmpty(funcEditorName) Then funcEditorName = funcName

        Dim paramsUBound As Integer = -1
        If IsNothing(curFuncParams) = False Then
            paramsUBound = curFuncParams.GetUpperBound(0)
            For i As Integer = 0 To paramsUBound
                If curFuncParams(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY AndAlso i <> paramsUBound Then
                    MsgBox("Параметр " + curFuncParams(i).Name + " имеет тип " + Chr(34) + "Массив параметров" + Chr(34) + ", который может иметь лишь последний параметр.", MsgBoxStyle.Exclamation)
                    lstFuncParams.SelectedIndex = i
                    Exit Sub
                ElseIf curFuncParams(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM Then
                    If IsNothing(curFuncParams(i).EnumValues) = True OrElse curFuncParams(i).EnumValues.GetUpperBound(0) = -1 Then
                        MsgBox("Параметр " + curFuncParams(i).Name + " имеет тип " + Chr(34) + "Один из возможных" + Chr(34) + ", для которого необходимо перечисление этих возможных вариантов.", MsgBoxStyle.Exclamation)
                        lstFuncParams.SelectedIndex = i
                        Exit Sub
                    End If
                ElseIf curFuncParams(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                    If IsNothing(curFuncParams(i).EnumValues) = True OrElse curFuncParams(i).EnumValues.GetUpperBound(0) = -1 Then
                        MsgBox("Параметр " + curFuncParams(i).Name + " имеет тип " + Chr(34) + "Элемент" + Chr(34) + ", для которого необходимо выбрать дочерний элемент.", MsgBoxStyle.Exclamation)
                        lstFuncParams.SelectedIndex = i
                        Exit Sub
                    End If
                End If
            Next
        End If
        'Заполняем массив arrReturnEnum значениями, которые возвращает функция
        Dim arrReturnEnum() As String = Nothing, arrReturnEnumUBound As Integer = -1
        If lstFuncReturnType.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ENUM Then
            Dim curStr As String
            For i As Integer = 0 To txtFuncResult.Lines.GetUpperBound(0)
                curStr = txtFuncResult.Lines(i).Trim
                If curStr.Length > 0 Then
                    If Double.TryParse(curStr, Globalization.NumberStyles.Any, provider_points, Nothing) = False Then
                        If curStr.ToLower = "true" Then
                            curStr = "True"
                        ElseIf curStr.ToLower = "false" Then
                            curStr = "False"
                        Else
                            If curStr.StartsWith("'") = False And curStr.EndsWith("'") = False Then
                                curStr = "'" + curStr.Replace("'", "/'") + "'"
                            End If
                        End If
                    End If
                    arrReturnEnumUBound += 1
                    ReDim Preserve arrReturnEnum(arrReturnEnumUBound)
                    arrReturnEnum(arrReturnEnumUBound) = curStr
                End If
            Next
            If arrReturnEnumUBound = -1 Then
                MsgBox("Для результата функции " + Chr(34) + "один из возможных" + Chr(34) + " надо перечислить эти возможные варианты.", MsgBoxStyle.Exclamation)
                txtFuncResult.Focus()
                Exit Sub
            End If
        End If

        'Заполняем func введенными данными о функции
        Dim func As New MatewScript.PropertiesInfoType
        func.Description = txtFuncDescription.Text
        func.helpFile = txtFuncHelpFile.Text
        func.paramsMax = IIf(nudFuncMax.Enabled, nudFuncMax.Value, -1)
        func.paramsMin = nudFuncMin.Value
        func.params = curFuncParams
        func.EditorCaption = funcEditorName
        func.Hidden = cmbFuncHidden.SelectedIndex
        func.returnType = lstFuncReturnType.SelectedIndex

        If func.returnType = MatewScript.ReturnFunctionEnum.RETURN_ENUM Then func.returnArray = arrReturnEnum
        'сохраняем данные в структуре mainClass
        If curAction = SavingClassActionEnum.NEW_ELEMENT Then
            If DEEP_EDIT_MODE = False Then func.UserAdded = True
            'mScript.mainClass(selectedClassId).Functions.Add(funcName, func)
            AddUserFunctionByPropData(selectedClassId, funcName, func)
        Else
            func.UserAdded = mScript.mainClass(selectedClassId).Functions(selectedElementName).UserAdded
            If funcName = selectedElementName Then
                mScript.mainClass(selectedClassId).Functions(funcName) = func
            Else
                'mScript.mainClass(selectedClassId).Functions.Remove(selectedElementName)
                'mScript.mainClass(selectedClassId).Functions.Add(funcName, func)
                RemoveUserFunction({mScript.mainClass(selectedClassId).Names(0), "'" & selectedElementName & "'"}, Nothing, True)
                AddUserFunctionByPropData(selectedClassId, funcName, func)
                selectedElementName = funcName 'меняем имя выбранного элемента selectedElementName на новое
            End If
        End If

        If curAction = SavingClassActionEnum.EDIT_ELEMENT Then
            'меняем имя функции на новое в дереве классов
            treeClasses.SelectedNode.Text = funcName
        Else
            'вносим новую функцию в дерево классов
            Dim newNode As TreeNode
            If IsNothing(treeClasses.SelectedNode.Tag) Then
                If treeClasses.SelectedNode.Parent.Tag = "func" Then
                    newNode = treeClasses.SelectedNode.Parent.Nodes.Add(funcName)
                Else
                    newNode = treeClasses.SelectedNode.Parent.Parent.Nodes(1).Nodes.Add(funcName)
                End If
            ElseIf treeClasses.SelectedNode.Tag.ToString.StartsWith("Class") Then
                newNode = treeClasses.SelectedNode.Nodes(0).Nodes.Add(funcName)
            ElseIf treeClasses.SelectedNode.Tag.ToString = "func" Then
                newNode = treeClasses.SelectedNode.Nodes.Add(funcName)
            ElseIf treeClasses.SelectedNode.Tag.ToString = "prop" Then
                newNode = treeClasses.SelectedNode.Parent.Nodes(0).Nodes.Add(funcName)
            Else
                FillTreeViewWithClasses(treeClasses)
                newNode = treeClasses.Nodes(0)
            End If
            treeClasses.SelectedNode = newNode
        End If
        'прячем форму ввода
        pnlFunction.Hide()
        mScript.FillFuncAndPropHash()
    End Sub

    Private Sub lstPropReturn_SelectedIndexChanged(ByVal sender As ListBox, ByVal e As System.EventArgs) Handles lstPropReturn.SelectedIndexChanged
        lblPropReturnArray.Hide()
        txtPropReturnArray.Hide()
        lblPropElementClass.Hide()
        lstPropElementClass.Hide()
        btnPropParamShow.Hide()
        Select Case lstPropReturn.SelectedIndex
            Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                'тип - True/False
                txtProp.Hide()
                rtbProp.Hide()
                pnlDataType.Hide()
                pnlPropBool.Show()
            Case MatewScript.ReturnFunctionEnum.RETURN_EVENT, MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION
                'тип - событие
                txtProp.Hide()
                If lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                    btnPropParamShow.Show()
                    codeProp.codeBox.IsTextBlockByDefault = False
                Else
                    codeProp.codeBox.IsTextBlockByDefault = True
                End If
                pnlPropBool.Hide()
                pnlDataType.Hide()
                rtbProp.Show()
            Case Else
                'тип - обычный или "Один из вариантов"
                'Если тип возвращаемого значения свойства один из набора вариантов, то отображаем txtPropReturnArray для ввода
                'этих вариантов. Иначе они остаются спрятанными.
                If lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ENUM Then
                    lblPropReturnArray.Show()
                    txtPropReturnArray.Show()
                ElseIf lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then
                    lblPropElementClass.Show()
                    lstPropElementClass.Show()
                End If
                pnlPropBool.Hide()
                pnlDataType.Show()
                If optValueCode.Checked Then
                    rtbProp.Show()
                    txtProp.Hide()
                Else
                    rtbProp.Hide()
                    txtProp.Show()
                End If
        End Select
    End Sub

    Private Sub btnNewProperty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewProperty.Click
        'Отображаем и подготоваливаем форму для ввода нового свойства
        pnlClass.Hide()
        pnlFunction.Hide()
        pnlProperty.Show()
        pnlPropParams.Hide()
        curAction = SavingClassActionEnum.NEW_ELEMENT
        rtbProp.Clear()
        txtProp.Clear()
        codeProp.codeBox.Text = ""
        optPropValueFalse.Checked = True
        optValueNormal.Checked = True
        rtbProp.ForeColor = Color.Black
        lstPropReturn.SelectedIndex = 0
        txtPropName.Clear()
        txtPropEditorName.Clear()
        cmbPropHidden.SelectedIndex = 0
        txtPropDescription.Clear()
        txtPropHelp.Clear()
        txtPropReturnArray.Clear()
        lstPropElementClass.SelectedIndex = -1
        FillPropParamArray("", True)
        'treeClasses.Enabled = False
    End Sub

    Private Sub btnSaveProperty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveProperty.Click
        'Создание нового или редактирование существующего свойства
        'проверка корректности данных
        Dim propName As String = txtPropName.Text.Trim
        If propName.Length = 0 Then
            MsgBox("Укажите имя свойства.", MsgBoxStyle.Exclamation)
            txtPropName.Focus()
            Exit Sub
        End If

        If mScript.mainClass(selectedClassId).Properties.ContainsKey(propName) Then
            If (propName = selectedElementName And curAction = SavingClassActionEnum.EDIT_ELEMENT) = False Then
                MsgBox("Свойство " + propName + " уже существует. Укажите другое имя.", MsgBoxStyle.Exclamation)
                txtPropName.Focus()
                Exit Sub
            End If
        End If

        If curAction = SavingClassActionEnum.EDIT_ELEMENT AndAlso DEEP_EDIT_MODE = False Then
            Dim f As MatewScript.PropertiesInfoType = mScript.mainClass(selectedClassId).Properties(selectedElementName)
            If IsNothing(f) = False AndAlso f.UserAdded = False Then
                'редактирование встроенного свойства
                If MessageBox.Show("Редактирование встроенного свойства. Изменение имени свойста, возвращаемого значения, а также параметров свойства-события может привести к неработоспособности свойства. Продолжить?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.No Then Return
            End If
        End If

        Dim propEditorName As String = txtPropEditorName.Text.Trim
        If String.IsNullOrEmpty(propEditorName) Then propEditorName = propName

        Dim blnReturnCode As Boolean = False 'в свойстве сохраняется скрипт?
        If lstPropReturn.SelectedIndex <> MatewScript.ReturnFunctionEnum.RETURN_BOOl Then
            If (optValueCode.Checked OrElse (lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_EVENT OrElse lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION)) _
                Then blnReturnCode = True 'да, сохраняется скрипт
        End If

        If blnReturnCode AndAlso mScript.LAST_ERROR.Length > 0 Then Return 'код содержит ошибку - выход

        'Заполняем массив arrReturnEnum значениями, которые возвращает свойство
        Dim arrReturnEnum() As String = Nothing, arrReturnEnumUBound As Integer = -1
        If lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ENUM Then
            Dim curStr As String
            For i As Integer = 0 To txtPropReturnArray.Lines.GetUpperBound(0)
                curStr = txtPropReturnArray.Lines(i).Trim
                If curStr.Length > 0 Then
                    If Double.TryParse(curStr, Globalization.NumberStyles.Any, provider_points, Nothing) = False Then
                        If curStr.ToLower = "true" Then
                            curStr = "True"
                        ElseIf curStr.ToLower = "false" Then
                            curStr = "False"
                        Else
                            If curStr.StartsWith("'") = False And curStr.EndsWith("'") = False Then
                                curStr = "'" + curStr.Replace("'", "/'") + "'"
                            End If
                        End If
                    End If
                    arrReturnEnumUBound += 1
                    ReDim Preserve arrReturnEnum(arrReturnEnumUBound)
                    arrReturnEnum(arrReturnEnumUBound) = curStr
                End If
            Next
            If arrReturnEnumUBound = -1 Then
                MsgBox("Для типа возвращаемого значения " + Chr(34) + "один из возможных" + Chr(34) + " надо перечислить эти возможные значения.", MsgBoxStyle.Exclamation)
                txtPropReturnArray.Focus()
                Exit Sub
            End If
        ElseIf lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then
            If lstPropElementClass.SelectedIndex = -1 Then
                MsgBox("Для типа возвращаемого значения " + Chr(34) + "элемент" + Chr(34) + " надо выбрать класс элемента.", MsgBoxStyle.Exclamation)
                lstPropElementClass.Focus()
                Exit Sub
            End If
            ReDim arrReturnEnum(0)
            arrReturnEnum(0) = lstPropElementClass.Text
        End If

        'Заполняем prop введенными данными о свойстве
        'сохраняем данные в структуре mainClass
        Dim prop As New MatewScript.PropertiesInfoType
        prop.Description = txtPropDescription.Text
        prop.helpFile = txtPropHelp.Text
        prop.returnType = lstPropReturn.SelectedIndex
        prop.EditorCaption = propEditorName
        prop.Hidden = cmbPropHidden.SelectedIndex
        If blnReturnCode Then
            'заполняем значение свойства сериализованным в xml кодом
            prop.Value = codeProp.codeBox.SerializeCodeData()
        Else
            If prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_BOOl Then
                If optPropValueTrue.Checked Then
                    prop.Value = "True"
                Else
                    prop.Value = "False"
                End If
            Else
                Dim strDefValue As String = txtProp.Text.Trim
                If strDefValue.Length > 0 Then
                    If Double.TryParse(strDefValue, Globalization.NumberStyles.Any, provider_points, Nothing) Then
                        prop.Value = strDefValue
                    ElseIf strDefValue.ToLower = "true" Then
                        prop.Value = "True"
                    ElseIf strDefValue.ToLower = "false" Then
                        prop.Value = strDefValue = "False"
                    ElseIf strDefValue.Length > 1 AndAlso strDefValue.First = "'"c AndAlso strDefValue.Last = "'"c Then
                        prop.Value = strDefValue
                    Else
                        prop.Value = "'" + strDefValue.Replace("'", "/'") + "'"
                    End If
                Else
                    prop.Value = "''"
                End If
            End If
        End If

        If prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT AndAlso IsNothing(arrPropParams) = False AndAlso arrPropParams.Count > 0 Then
            'Вносим данные о параметрах свойства-события
            ReDim prop.params(arrPropParams.Count - 1)
            For i As Integer = 0 To arrPropParams.Count - 1
                prop.params(i) = New MatewScript.paramsType
                Dim strName As String = arrPropParams(i).paramCaption
                Dim scores As Integer = 0
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then scores += 16
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then scores += 8
                If arrPropParams(i).Level3 Then scores += 4
                If arrPropParams(i).Level2 Then scores += 2
                If arrPropParams(i).Level1 Then scores += 1
                prop.params(i).Type = scores
                prop.params(i).Name = strName
                prop.params(i).Description = arrPropParams(i).paramDescription
            Next
        End If

        If prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_ENUM OrElse prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then prop.returnArray = arrReturnEnum
        If curAction = SavingClassActionEnum.NEW_ELEMENT Then
            If DEEP_EDIT_MODE = False Then prop.UserAdded = True
            modMainCode.AddUserPropertyByPropData(selectedClassId, propName, prop)
            'mScript.mainClass(selectedClassId).Properties.Add(propName, prop)
        Else
            prop.UserAdded = mScript.mainClass(selectedClassId).Properties(selectedElementName).UserAdded
            'If propName = selectedElementName Then
            '    mScript.mainClass(selectedClassId).Properties(propName) = prop
            'Else
            '    RemoveUserProperty({mScript.mainClass(selectedClassId).Names(0), "'" & selectedElementName & "'"}, Nothing, True)
            '    AddUserPropertyByPropData(selectedClassId, propName, prop)
            '    selectedElementName = propName
            'End If

            'RemoveUserProperty({mScript.mainClass(selectedClassId).Names(0), "'" & selectedElementName & "'"}, Nothing, True)
            'AddUserPropertyByPropData(selectedClassId, propName, prop)
            UpdateUserPropertyByPropData(selectedClassId, selectedElementName, propName, prop)

            If selectedElementName <> propName Then
                Dim q As New EditQueueClass With {.ActionType = EditQueueClass.EditActionTypeEnum.EDIT_PROPERTY_NAME, .classId = selectedClassId, .newValue = propName, .oldValue = selectedElementName}
                lstQueue.Add(q)
            End If

            selectedElementName = propName
        End If

        'treeClasses.Enabled = True
        If curAction = SavingClassActionEnum.EDIT_ELEMENT Then
            'меняем имя свойства на новое в дереве классов
            treeClasses.SelectedNode.Text = propName

        Else
            'вносим новое свойство в дерево классов
            Dim newNode As TreeNode
            If IsNothing(treeClasses.SelectedNode.Tag) Then
                If treeClasses.SelectedNode.Parent.Tag = "prop" Then
                    newNode = treeClasses.SelectedNode.Parent.Nodes.Add(propName)
                Else
                    newNode = treeClasses.SelectedNode.Parent.Parent.Nodes(0).Nodes.Add(propName)
                End If
            ElseIf treeClasses.SelectedNode.Tag.ToString.StartsWith("Class") Then
                newNode = treeClasses.SelectedNode.Nodes(1).Nodes.Add(propName)
            ElseIf treeClasses.SelectedNode.Tag.ToString = "func" Then
                newNode = treeClasses.SelectedNode.Parent.Nodes(1).Nodes.Add(propName)
            ElseIf treeClasses.SelectedNode.Tag.ToString = "prop" Then
                newNode = treeClasses.SelectedNode.Nodes.Add(propName)
            Else
                FillTreeViewWithClasses(treeClasses)
                newNode = treeClasses.Nodes(0)
            End If
            treeClasses.SelectedNode = newNode

            Dim q As New EditQueueClass With {.ActionType = EditQueueClass.EditActionTypeEnum.CREATE_NEW_PROPERTY, .classId = selectedClassId, .newValue = propName, .oldValue = ""}
            lstQueue.Add(q)
        End If

        'прячем форму ввода
        pnlProperty.Hide()
        mScript.FillFuncAndPropHash()
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        treeClasses.Focus()
        If selectedElementName.Length = 0 Then
            'Удаляем класс
            If mScript.mainClass.GetUpperBound(0) = 0 Then
                MsgBox("Нельзя удалять все классы.", MsgBoxStyle.Exclamation)
                Return
            ElseIf DEEP_EDIT_MODE = False AndAlso mScript.mainClass(selectedClassId).UserAdded = False Then
                If MessageBox.Show("Удаление встроенного класса может привести к нестабильности работы движка. Продолжить?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.No Then Return
            ElseIf IsNothing(mScript.mainClass(selectedClassId).ChildProperties) = False AndAlso mScript.mainClass(selectedClassId).ChildProperties.Count > 0 Then
                MessageBox.Show("Класс нельзя удалить, так как он имеет дочерние элементы.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If
            'удаляем из структуры mainClass
            For i As Integer = selectedClassId To mScript.mainClass.GetUpperBound(0) - 1
                mScript.mainClass(i) = mScript.mainClass(i + 1)
            Next
            ReDim Preserve mScript.mainClass(mScript.mainClass.GetUpperBound(0) - 1)

            Dim q As New EditQueueClass With {.ActionType = EditQueueClass.EditActionTypeEnum.REMOVE_CLASS, .classId = selectedClassId, .newValue = mScript.mainClass(selectedClassId).Names(0)}
            lstQueue.Add(q)

            mScript.MakeMainClassHash() 'обновляем хэш с именами классов
            'удаляем из дерева классов
            treeClasses.SelectedNode.Remove()
            treeClasses.SelectedNode = treeClasses.Nodes(0)
            FillListWithClasses()
        ElseIf selectedElementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_FUNCTION Then
            If DEEP_EDIT_MODE = False AndAlso mScript.mainClass(selectedClassId).Functions(selectedElementName).UserAdded = False Then
                If MessageBox.Show("Удаление встроенной функции может привести к нестабильности работы движка. Продолжить?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.No Then Return
            End If
            'Удаляем функцию
            'удаляем из структуры mainClass
            mScript.mainClass(selectedClassId).Functions.Remove(selectedElementName)
            'удаляем из дерева классов
            treeClasses.SelectedNode.Remove()
            treeClasses.SelectedNode = treeClasses.Nodes(selectedClassId)
        ElseIf selectedElementType = MatewScript.funcAndPropHashType.funcOrPropEnum.E_PROPERTY Then
            'Удаляем свойство
            If DEEP_EDIT_MODE = False AndAlso mScript.mainClass(selectedClassId).Properties(selectedElementName).UserAdded = False Then
                If MessageBox.Show("Удаление встроенного свойства может привести к нестабильности работы движка. Продолжить?", "MatewQuest", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.No Then Return
            End If

            Dim q As New EditQueueClass With {.ActionType = EditQueueClass.EditActionTypeEnum.REMOVE_PROPERTY, .classId = selectedClassId, .newValue = selectedElementName, .oldValue = ""}
            lstQueue.Add(q)

            'удаляем из структуры mainClass
            mScript.mainClass(selectedClassId).Properties.Remove(selectedElementName)

            'удаляем из дерева классов
            treeClasses.SelectedNode.Remove()
            treeClasses.SelectedNode = treeClasses.Nodes(selectedClassId)
        End If
        mScript.FillFuncAndPropHash()
    End Sub


    Private Sub ExpandAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExpandAllToolStripMenuItem.Click
        treeClasses.ExpandAll() 'развернуть дерево классов
    End Sub

    Private Sub CollapseAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CollapseAllToolStripMenuItem.Click
        treeClasses.CollapseAll() 'свернуть дерево классов
    End Sub

    Private Sub SortToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SortToolStripMenuItem.Click
        treeClasses.Sort() 'сортировать дерево классов
    End Sub

    Private Sub optValueNormal_CheckedChanged(sender As Object, e As EventArgs) Handles optValueNormal.CheckedChanged
        txtProp.Show()
        rtbProp.Hide()
        btnEditCode.Hide()
    End Sub

    Private Sub optValueCode_CheckedChanged(sender As Object, e As EventArgs) Handles optValueCode.CheckedChanged
        txtProp.Hide()
        rtbProp.Show()
        btnEditCode.Show()
    End Sub

    Private Sub btnEditCode_Click(sender As Object, e As EventArgs) Handles btnEditCode.Click
        CodeData = CopyCodeDataArray(codeProp.codeBox.CodeData)
        splitCode.Show()
        splitCode.SplitterDistance = splitCode.Height - btnCodeSave.Height - 30
        codeProp.Dock = DockStyle.Fill
        codeProp.codeBox.Refresh()

    End Sub

    Private Sub rtbProp_VisibleChanged(sender As Object, e As EventArgs) Handles rtbProp.VisibleChanged
        btnEditCode.Visible = rtbProp.Visible
    End Sub

    Private Sub btnCodeCancel_Click(sender As Object, e As EventArgs) Handles btnCodeCancel.Click
        codeProp.codeBox.LoadCodeFromCodeData(CodeData)
        splitCode.Hide()
    End Sub

    Private Sub btnCodeSave_Click(sender As Object, e As EventArgs) Handles btnCodeSave.Click
        If mScript.LAST_ERROR.Length > 0 Then
            MsgBox("Нельзя сохранить код с ошибками!", vbExclamation)
            Return
        End If
        rtbProp.Rtf = codeProp.codeBox.Rtf
        splitCode.Hide()
    End Sub


    Private Sub btnDuplicateFunction_Click(sender As Object, e As EventArgs) Handles btnDuplicateFunction.Click
        'Dim newFuncParams() As MatewScript.paramsType = curFuncParams.
        If IsNothing(curFuncParams) = False Then
            curFuncParams = curFuncParams.Clone
            For i As Integer = 0 To curFuncParams.Count - 1
                curFuncParams(i) = curFuncParams(i).Clone
            Next
        End If
        curAction = SavingClassActionEnum.NEW_ELEMENT
        txtFuncName.Text = "NewFunction"
        If txtFuncName.Visible = False Then Return
        txtFuncName.Focus()
        txtFuncName.SelectAll()
    End Sub

    Private Sub btnDuplicateProperty_Click(sender As Object, e As EventArgs) Handles btnDuplicateProperty.Click
        curAction = SavingClassActionEnum.NEW_ELEMENT
        txtPropName.Text = "NewProperty"
        If txtPropName.Visible = False Then Return
        txtPropName.Focus()
        txtPropName.SelectAll()
    End Sub

    ''' <summary>
    ''' Заполняет listView с названиями классов 2 уровня (для выбора если returnType = ELEMENT)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FillListWithClasses()
        lstPropElementClass.Items.Clear()
        lstParamsClass.Items.Clear()
        If IsNothing(mScript.mainClass) OrElse mScript.mainClass.Count = 0 Then Return
        For i As Integer = 0 To mScript.mainClass.Count - 1
            If mScript.mainClass(i).LevelsCount = 0 Then Continue For
            lstPropElementClass.Items.Add(mScript.mainClass(i).Names.Last)
        Next
        lstPropElementClass.Items.Add("Variable")
        lstParamsClass.Items.AddRange(lstPropElementClass.Items)
    End Sub

    Private Sub btnUp_Click(sender As Object, e As EventArgs) Handles btnUp.Click
        'перемещаем текущий узел на 1 выше
        Dim n As TreeNode = treeClasses.SelectedNode
        If IsNothing(n) Then Return
        If IsNothing(n.Tag) = False Then Return
        pnlProperty.Hide()
        pnlFunction.Hide()

        Dim parN As TreeNode = n.Parent
        If Object.Equals(parN.Nodes(0), n) Then Return 'узел первы - выход
        'переставляем узел в дереве
        Dim newIndex As Integer = n.Index - 1
        n.Remove()
        parN.Nodes.Insert(newIndex, n)
        treeClasses.SelectedNode = n
        'обновляем дов\ступность кнопок
        If n.Index = 0 Then btnUp.Enabled = False
        If n.Index < parN.Nodes.Count - 1 Then
            btnDn.Enabled = True
        Else
            btnDn.Enabled = False
        End If
    End Sub

    Private Sub btnDn_Click(sender As Object, e As EventArgs) Handles btnDn.Click
        'перемещаем текущий узел на 1 ниже
        Dim n As TreeNode = treeClasses.SelectedNode
        If IsNothing(n) Then Return
        If IsNothing(n.Tag) = False Then Return
        pnlProperty.Hide()
        pnlFunction.Hide()

        Dim parN As TreeNode = n.Parent
        Dim lastId As Integer = parN.Nodes.Count - 1
        If Object.Equals(parN.Nodes(lastId), n) Then Return 'узел последний - выход
        'переставляем узел в дереве
        Dim newIndex As Integer = n.Index + 1
        n.Remove()
        parN.Nodes.Insert(newIndex, n)
        treeClasses.SelectedNode = n
        'обновляем доступность кнопок
        If n.Index = lastId Then btnDn.Enabled = False
        If n.Index = 0 Then
            btnUp.Enabled = False
        Else
            btnUp.Enabled = True
        End If
    End Sub

    Private Sub lstParamsClass_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstParamsClass.SelectedIndexChanged
        'Сохраняем класс при значении paramType = ELEMENT
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Return
        If lstParamsClass.SelectedIndex = -1 Then Return
        curFuncParams(selIndex).EnumValues = {lstParamsClass.Text}
    End Sub

    Private Sub lblParamsClass_Click(sender As Object, e As EventArgs) Handles lblParamsClass.Click

    End Sub

    Private Sub btnPropGenerateHelp_Click(sender As Object, e As EventArgs) Handles btnPropGenerateHelp.Click
        'Открывает диалог генерации файла помощи
        Dim dRes As DialogResult = Windows.Forms.DialogResult.No
        Dim fPath As String = ""
        If txtPropHelp.Text.Length > 0 Then
            fPath = GetHelpPath(txtPropHelp.Text)
            If My.Computer.FileSystem.FileExists(fPath) Then
                dRes = MessageBox.Show("К свойству уже присоединен файл помощи. Редактировать его?", "Matew Quest", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If dRes = Windows.Forms.DialogResult.Cancel Then Return
            End If
        End If

        If txtPropName.Text.Length = 0 Then
            MsgBox("Не введено имя свойства!")
            txtPropName.Focus()
            Return
        End If
        If txtPropDescription.Text.Length = 0 Then
            MsgBox("Не введено описание свойства!")
            txtPropDescription.Focus()
            Return
        End If
        'Перед началом сохраняем свойство
        btnSaveProperty_Click(btnSaveProperty, New EventArgs)
        Dim prop As MatewScript.PropertiesInfoType = Nothing
        If mScript.mainClass(selectedClassId).Properties.TryGetValue(txtPropName.Text, prop) = False Then Return
        If dRes = Windows.Forms.DialogResult.No Then
            'Генерируем файл помощи
            dlgGenerateHelp.GenerateHelpForProperty(txtPropName.Text, selectedClassId, prop)
        Else
            'Открываем готовый файл
            dlgGenerateHelp.codeHelp.Text = My.Computer.FileSystem.ReadAllText(fPath, System.Text.Encoding.Default)
        End If
        dRes = dlgGenerateHelp.ShowDialog(Me)
        pnlProperty.Show()

        If dRes = Windows.Forms.DialogResult.OK Then
            'Открываем диалог сохранения файла
            'Выбираем начальную директорию
            If fPath.Length > 0 Then
                Dim fName As String = System.IO.Path.GetFileName(fPath)
                fPath = fPath.Substring(0, fPath.Length - fName.Length)
            ElseIf DEEP_EDIT_MODE Then
                fPath = APP_HELP_PATH
            Else
                fPath = My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "help")
            End If
            sfd.InitialDirectory = fPath
            sfd.FileName = txtPropName.Text.Trim + ".html"
            dRes = sfd.ShowDialog(Me)
            If dRes = Windows.Forms.DialogResult.OK Then
                'Сохраняем файл
                fPath = sfd.FileName
                My.Computer.FileSystem.WriteAllText(fPath, dlgGenerateHelp.codeHelp.Text, False, System.Text.Encoding.Default)
                Dim initDir As String = My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "help") + "\"
                If fPath.StartsWith(APP_HELP_PATH, StringComparison.CurrentCultureIgnoreCase) Then
                    fPath = fPath.Substring(APP_HELP_PATH.Length)
                ElseIf fPath.StartsWith(initDir, StringComparison.CurrentCultureIgnoreCase) Then
                    fPath = fPath.Substring(initDir.Length)
                End If

                fPath = fPath.Replace(APP_HELP_PATH, "")
                txtPropHelp.Text = fPath
            End If
        End If
    End Sub

#Region "Property Params"
    Private curPropParamId As Integer = -1

    ''' <summary>
    ''' Заполняет список параметров свойства типа EVENT, необходимых для генерации файла помощи и вывода типичных переменных
    ''' </summary>
    ''' <param name="propName">Имя свойства</param>
    Private Sub FillPropParamArray(ByVal propName As String, ByVal newProperty As Boolean)
        'paramType:
        '1 - Level1
        '2 - Level2 (3)
        '4 - Level3 (5-7)
        '8 - Return (9-15)
        '16- Array (17-31)
        ClearPropParamArray()
        Dim prop As MatewScript.PropertiesInfoType = Nothing

        If newProperty = False AndAlso mScript.mainClass(selectedClassId).Properties.TryGetValue(propName, prop) = False Then Return
        If newProperty Then
            'Новое свойство - создаем стандартные параметры
            If mScript.mainClass(selectedClassId).LevelsCount = 1 Then
                arrPropParams.Add(New PropParamsClass With {.paramType = PropParamsClass.ParamTypeEnum.T_PARAM, .paramCaption = mScript.mainClass(selectedClassId).Names(0) + "Id", _
                                                            .paramDescription = mScript.mainClass(selectedClassId).Names.Last + " Id", .Level1 = True, .Level2 = True})
            ElseIf mScript.mainClass(selectedClassId).LevelsCount = 2 Then
                arrPropParams.Add(New PropParamsClass With {.paramType = PropParamsClass.ParamTypeEnum.T_PARAM, .paramCaption = mScript.mainClass(selectedClassId).Names(0) + "Id", _
                                                           .paramDescription = mScript.mainClass(selectedClassId).Names.Last + " Id", .Level1 = True, .Level2 = True})
                arrPropParams.Add(New PropParamsClass With {.paramType = PropParamsClass.ParamTypeEnum.T_PARAM, .paramCaption = mScript.mainClass(selectedClassId).Names(0) + "ItemId", _
                                                .paramDescription = mScript.mainClass(selectedClassId).Names.Last + " subitem Id", .Level1 = True, .Level2 = True, .Level3 = True})
            End If
        Else
            If IsNothing(prop.params) OrElse prop.params.Count = 0 Then Return
            'Параметры есть - получаем
            Dim params() As MatewScript.paramsType = prop.params
            For i As Integer = 0 To params.Count - 1
                Dim propParam As New PropParamsClass
                Dim pName As String = params(i).Name
                Dim parScores As Integer = params(i).Type 'params.Type используется для хранения значений Level1, Level2, Level3 & ParamArray
                propParam.paramType = PropParamsClass.ParamTypeEnum.T_PARAM
                If parScores >= 16 Then propParam.paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY : parScores -= 16
                If parScores >= 8 Then propParam.paramType = PropParamsClass.ParamTypeEnum.T_RETURN : parScores -= 8
                If parScores >= 4 Then propParam.Level3 = True : parScores -= 4
                If parScores >= 2 Then propParam.Level2 = True : parScores -= 2
                If parScores = 1 Then propParam.Level1 = True

                propParam.paramCaption = params(i).Name
                propParam.paramDescription = params(i).Description
                arrPropParams.Add(propParam)
            Next
        End If
        'Заполняем список
        btnPropParamArrayAdd.Enabled = True
        If arrPropParams.Count > 0 Then
            lstPropParams.BeginUpdate()
            For i As Integer = 0 To arrPropParams.Count - 1
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM Then
                    lstPropParams.Items.Add("Param[" + i.ToString + "]")
                ElseIf arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
                    lstPropParams.Items.Add("Param[" + i.ToString + ", " + (i + 1).ToString + ", ... n]")
                    btnPropParamArrayAdd.Enabled = False
                Else
                    lstPropParams.Items.Add("Return " + arrPropParams(i).paramCaption)
                End If
            Next
            lstPropParams.EndUpdate()
        End If
        If lstPropParams.Items.Count > 0 Then lstPropParams.SelectedIndex = 0
    End Sub

    ''' <summary>
    ''' Очистка параметров свойств
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearPropParamArray()
        curPropParamId = -1
        lstPropParams.Items.Clear()
        If IsNothing(arrPropParams) = False Then arrPropParams.Clear()

        arrPropParams = New List(Of PropParamsClass)
        btnPropParamRemove.Enabled = False
        txtPropParamDescription.Clear()
        txtPropParamDescription.Enabled = False
        txtPropParamName.Clear()
        txtPropParamName.Enabled = False
        txtPropParamReturnDescription.Clear()
        txtPropParamReturnDescription.Enabled = False
        cmbPropParamReturnType.Text = ""
        cmbPropParamReturnType.Enabled = False
        chkPropParamLevel1.Enabled = False
        chkPropParamLevel2.Enabled = False
        chkPropParamLevel3.Enabled = False
        chkPropParamReturnLevel1.Enabled = False
        chkPropParamReturnLevel2.Enabled = False
        chkPropParamReturnLevel3.Enabled = False

        btnPropParamArrayAdd.Enabled = True
    End Sub

    Private Sub btnPropParamHide_Click(sender As Object, e As EventArgs) Handles btnPropParamHide.Click
        pnlPropParams.Hide()
    End Sub

    Private Sub btnPropParamArrayAdd_Click(sender As Object, e As EventArgs) Handles btnPropParamArrayAdd.Click
        sender.Enabled = False
        AddNewPropParam(PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY)
    End Sub

    ''' <summary>
    ''' Добавляет новый параметр свойства
    ''' </summary>
    ''' <param name="parType">Тип добавляемого параметра</param>
    Private Sub AddNewPropParam(ByVal parType As PropParamsClass.ParamTypeEnum)
        Dim propParam As New PropParamsClass With {.paramType = parType, .paramCaption = "", .paramDescription = "", .Level1 = True}
        If mScript.mainClass(selectedClassId).LevelsCount >= 1 Then propParam.Level2 = True
        If mScript.mainClass(selectedClassId).LevelsCount >= 2 Then propParam.Level3 = True

        Dim insertPos As Integer = -1
        Dim strText As String
        Dim hasParamArray As Boolean = False
        If parType <> PropParamsClass.ParamTypeEnum.T_RETURN Then
            'Если мы вставляем обычный параметр, то он должен следовать перед параметрами типа Return & Param_Array. Для этого получаем insertPos - позицию, куда надо вставить новый паарметр
            For i As Integer = 0 To arrPropParams.Count - 1
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
                    If parType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
                        MsgBox("Массив параметров уже существует и может быть только один!")
                        Return
                    End If
                    hasParamArray = True 'имеется массив параметров
                    insertPos = i
                    Exit For
                ElseIf arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
                    insertPos = i
                    Exit For
                End If
            Next i
        End If

        If insertPos = -1 Then
            'Вставляем в конец
            insertPos = arrPropParams.Count
            arrPropParams.Add(propParam)
        Else
            'Вставляем последним параметром, но перед параметрами типа Return & Param_Array
            arrPropParams.Insert(insertPos, propParam)
        End If

        'Генерирует тескт для листбоска
        If parType = PropParamsClass.ParamTypeEnum.T_PARAM Then
            strText = "Param[" + insertPos.ToString + "]"
        ElseIf parType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
            strText = "Param[" + insertPos.ToString + ", " + (insertPos + 1).ToString + ", ... n]"
        Else
            strText = "Return (ничего)"
        End If
        'Вставляем новый параметр в дистбокс
        If insertPos >= lstPropParams.Items.Count Then
            lstPropParams.Items.Add(strText)
        Else
            lstPropParams.Items.Insert(insertPos, strText)
        End If

        If hasParamArray Then
            'изменяем такст в листбоксе у параметра типа Array
            lstPropParams.Items(insertPos + 1) = "Param[" + (insertPos + 1).ToString + ", " + (insertPos + 2).ToString + ", ... n]"
        End If

        lstPropParams.SelectedIndex = insertPos
    End Sub

    Private Sub btnPropParamAdd_Click(sender As Object, e As EventArgs) Handles btnPropParamAdd.Click
        AddNewPropParam(PropParamsClass.ParamTypeEnum.T_PARAM)
    End Sub

    Private Sub btnPropParamReturnAdd_Click(sender As Object, e As EventArgs) Handles btnPropParamReturnAdd.Click
        AddNewPropParam(PropParamsClass.ParamTypeEnum.T_RETURN)
    End Sub

    Private Sub btnPropParamRemove_Click(sender As Object, e As EventArgs) Handles btnPropParamRemove.Click
        'Удаляем параметр свойства
        If lstPropParams.SelectedIndex = -1 Then Return
        Dim delId As Integer = lstPropParams.SelectedIndex

        If arrPropParams(delId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM Then
            'Изменяем текст в листбоксе для параметра типа Array, а также обычных параметров, если они находились выше удаляемого (индекс -1)
            For i As Integer = 0 To arrPropParams.Count - 1
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
                    lstPropParams.Items(i) = "Param[" + (i - 1).ToString + ", " + i.ToString + ", ... n]"
                    Exit For
                ElseIf arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
                    Exit For
                ElseIf arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM AndAlso i > delId Then
                    lstPropParams.Items(i) = "Param[" + (i - 1).ToString + "]"
                End If
            Next
        ElseIf arrPropParams(delId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
            btnPropParamArrayAdd.Enabled = True
        End If
        'Собственно удаление и выбор ближайшего другого параметра
        arrPropParams.RemoveAt(delId)
        lstPropParams.Items.RemoveAt(delId)
        If arrPropParams.Count = 0 Then
            btnPropParamRemove.Enabled = False
            Return
        ElseIf delId >= arrPropParams.Count - 1 Then
            lstPropParams.SelectedIndex = delId - 1
        Else
            delId -= 1
            lstPropParams.SelectedIndex = delId
        End If
    End Sub

    Private Sub btnPropParamUp_Click(sender As Object, e As EventArgs) Handles btnPropParamUp.Click
        Dim movId As Integer = lstPropParams.SelectedIndex
        If CheckPropParamButtonUp() = False Then Return
        Dim movParam As PropParamsClass = arrPropParams(movId)
        arrPropParams.RemoveAt(movId)
        arrPropParams.Insert(movId - 1, movParam)

        If arrPropParams(movId - 1).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
            Dim strText As String = lstPropParams.Items(movId)
            lstPropParams.Items.RemoveAt(movId)
            lstPropParams.Items.Insert(movId - 1, strText)
        End If
        lstPropParams.SelectedIndex = movId - 1
    End Sub

    Private Sub btnPropParamDn_Click(sender As Object, e As EventArgs) Handles btnPropParamDn.Click
        Dim movId As Integer = lstPropParams.SelectedIndex
        If CheckPropParamButtonDn() = False Then Return
        Dim movParam As PropParamsClass = arrPropParams(movId)
        arrPropParams.RemoveAt(movId)
        arrPropParams.Insert(movId + 1, movParam)

        If arrPropParams(movId + 1).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
            Dim strText As String = lstPropParams.Items(movId)
            lstPropParams.Items.RemoveAt(movId)
            lstPropParams.Items.Insert(movId + 1, strText)
        End If
        lstPropParams.SelectedIndex = movId + 1
    End Sub

    ''' <summary>
    ''' Проверяет доступно ли перемещение текущего параметра вверх, и, если недоступно , то делает его кнопку также недоступной
    ''' </summary>
    Private Function CheckPropParamButtonUp() As Boolean
        Dim curId As Integer = lstPropParams.SelectedIndex
        If curId <= 0 Then
            btnPropParamUp.Enabled = False
            Return False
        End If
        If arrPropParams(curId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
            btnPropParamUp.Enabled = False
            Return False
        End If
        If arrPropParams(curId).paramType = PropParamsClass.ParamTypeEnum.T_RETURN AndAlso arrPropParams(curId - 1).paramType <> PropParamsClass.ParamTypeEnum.T_RETURN Then
            btnPropParamUp.Enabled = False
            Return False
        End If
        btnPropParamUp.Enabled = True
        Return True
    End Function

    ''' <summary>
    ''' Проверяет доступно ли перемещение текущего параметра вниз, и, если недоступно , то делает его кнопку также недоступной
    ''' </summary>
    Private Function CheckPropParamButtonDn() As Boolean
        Dim curId As Integer = lstPropParams.SelectedIndex
        If curId = -1 OrElse curId >= arrPropParams.Count - 1 Then
            btnPropParamDn.Enabled = False
            Return False
        End If
        If arrPropParams(curId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
            btnPropParamDn.Enabled = False
            Return False
        End If
        If arrPropParams(curId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM AndAlso arrPropParams(curId + 1).paramType <> PropParamsClass.ParamTypeEnum.T_PARAM Then
            btnPropParamDn.Enabled = False
            Return False
        End If
        btnPropParamDn.Enabled = True
        Return True
    End Function

    Private Sub lstPropParams_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstPropParams.SelectedIndexChanged
        If lstPropParams.SelectedIndex = -1 Then Return
        curPropParamId = lstPropParams.SelectedIndex
        CheckPropParamButtonDn()
        CheckPropParamButtonUp()

        btnPropParamRemove.Enabled = True
        If arrPropParams(curPropParamId).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
            'Вставляем данные параметра типа Return
            txtPropParamDescription.Clear()
            txtPropParamDescription.Enabled = False
            txtPropParamName.Clear()
            txtPropParamName.Enabled = False
            chkPropParamLevel1.Enabled = False
            chkPropParamLevel2.Enabled = False
            chkPropParamLevel3.Enabled = False

            txtPropParamReturnDescription.Text = arrPropParams(curPropParamId).paramDescription
            txtPropParamReturnDescription.Enabled = True
            cmbPropParamReturnType.Text = arrPropParams(curPropParamId).paramCaption
            cmbPropParamReturnType.Enabled = True
            chkPropParamReturnLevel1.Enabled = True
            chkPropParamReturnLevel1.Checked = arrPropParams(curPropParamId).Level1
            If mScript.mainClass(selectedClassId).LevelsCount >= 1 Then
                chkPropParamReturnLevel2.Enabled = True
                chkPropParamReturnLevel2.Checked = arrPropParams(curPropParamId).Level2
            End If
            If mScript.mainClass(selectedClassId).LevelsCount >= 2 Then
                chkPropParamReturnLevel3.Enabled = True
                chkPropParamReturnLevel3.Checked = arrPropParams(curPropParamId).Level3
            End If

        Else
            'Вставляем данные параметра типа Param / ParamArray
            txtPropParamReturnDescription.Clear()
            txtPropParamReturnDescription.Enabled = False
            cmbPropParamReturnType.Text = ""
            cmbPropParamReturnType.Enabled = False
            chkPropParamReturnLevel1.Enabled = False
            chkPropParamReturnLevel2.Enabled = False
            chkPropParamReturnLevel3.Enabled = False

            txtPropParamDescription.Text = arrPropParams(curPropParamId).paramDescription
            txtPropParamDescription.Enabled = True
            txtPropParamName.Text = arrPropParams(curPropParamId).paramCaption
            txtPropParamName.Enabled = True
            chkPropParamLevel1.Enabled = True
            chkPropParamLevel1.Checked = arrPropParams(curPropParamId).Level1
            If mScript.mainClass(selectedClassId).LevelsCount >= 1 Then
                chkPropParamLevel2.Enabled = True
                chkPropParamLevel2.Checked = arrPropParams(curPropParamId).Level2
            End If
            If mScript.mainClass(selectedClassId).LevelsCount >= 2 Then
                chkPropParamLevel3.Enabled = True
                chkPropParamLevel3.Checked = arrPropParams(curPropParamId).Level3
            End If
        End If
    End Sub

    Private Sub txtPropParamName_TextChanged(sender As Object, e As EventArgs) Handles txtPropParamName.TextChanged
        If curPropParamId = -1 Then Return
        If arrPropParams(curPropParamId).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then Return
        arrPropParams(curPropParamId).paramCaption = sender.Text
    End Sub

    Private Sub txtPropParamDescription_TextChanged(sender As Object, e As EventArgs) Handles txtPropParamDescription.TextChanged
        If curPropParamId = -1 Then Return
        If arrPropParams(curPropParamId).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then Return
        arrPropParams(curPropParamId).paramDescription = sender.Text
    End Sub

    Private Sub cmbPropParamReturnType_TextChanged(sender As Object, e As EventArgs) Handles cmbPropParamReturnType.TextChanged
        If curPropParamId = -1 Then Return
        If arrPropParams(curPropParamId).paramType <> PropParamsClass.ParamTypeEnum.T_RETURN Then Return
        arrPropParams(curPropParamId).paramCaption = sender.Text
        If String.IsNullOrWhiteSpace(sender.Text) Then
            lstPropParams.Items(curPropParamId) = "Return (ничего)"
        Else
            lstPropParams.Items(curPropParamId) = "Return " + sender.Text
        End If
    End Sub

    Private Sub txtPropParamReturnDescription_TextChanged(sender As Object, e As EventArgs) Handles txtPropParamReturnDescription.TextChanged
        If curPropParamId = -1 Then Return
        If arrPropParams(curPropParamId).paramType <> PropParamsClass.ParamTypeEnum.T_RETURN Then Return
        arrPropParams(curPropParamId).paramDescription = sender.Text
    End Sub

    Private Sub btnPropParamsShow_Click(sender As Object, e As EventArgs) Handles btnPropParamShow.Click
        pnlPropParams.BringToFront()
        pnlPropParams.Show()
    End Sub

    Private Sub chkPropParamLevel1_CheckedChanged(sender As Object, e As EventArgs) Handles chkPropParamLevel1.CheckedChanged, chkPropParamLevel2.CheckedChanged, chkPropParamLevel3.CheckedChanged, _
        chkPropParamReturnLevel1.CheckedChanged, chkPropParamReturnLevel2.CheckedChanged, chkPropParamReturnLevel3.CheckedChanged
        If curPropParamId = -1 Then Return
        Dim lvl As Integer = CInt(sender.Tag)
        If lvl = 1 Then
            arrPropParams(curPropParamId).Level1 = sender.Checked
        ElseIf lvl = 2 Then
            arrPropParams(curPropParamId).Level2 = sender.Checked
        Else
            arrPropParams(curPropParamId).Level3 = sender.Checked
        End If
    End Sub

#End Region

    Private Sub btnFuncGenerateHelp_Click(sender As Object, e As EventArgs) Handles btnFuncGenerateHelp.Click
        'Открывает диалог генерации файла помощи функции
        Dim dRes As DialogResult = Windows.Forms.DialogResult.No
        Dim fPath As String = ""
        If txtFuncHelpFile.Text.Length > 0 Then
            fPath = GetHelpPath(txtFuncHelpFile.Text)
            If My.Computer.FileSystem.FileExists(fPath) Then
                dRes = MessageBox.Show("К функции уже присоединен файл помощи. Редактировать его?", "Matew Quest", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If dRes = Windows.Forms.DialogResult.Cancel Then Return
            End If
        End If

        If txtFuncName.Text.Length = 0 Then
            MsgBox("Не введено имя функции!")
            txtFuncName.Focus()
            Return
        End If
        If txtFuncDescription.Text.Length = 0 Then
            MsgBox("Не введено описание функции!")
            txtFuncDescription.Focus()
            Return
        End If
        'Перед началом сохраняем функцию
        btnSaveFunction_Click(btnSaveFunction, New EventArgs)
        Dim prop As MatewScript.PropertiesInfoType = Nothing
        If mScript.mainClass(selectedClassId).Functions.TryGetValue(txtFuncName.Text, prop) = False Then Return
        If dRes = Windows.Forms.DialogResult.No Then
            'Генерируем файл помощи
            dlgGenerateHelp.GenerateHelpForFunction(txtFuncName.Text, selectedClassId, prop)
        Else
            'Открываем готовый файл
            dlgGenerateHelp.codeHelp.Text = My.Computer.FileSystem.ReadAllText(fPath, System.Text.Encoding.Default)
        End If
        dRes = dlgGenerateHelp.ShowDialog(Me)
        pnlFunction.Show()

        If dRes = Windows.Forms.DialogResult.OK Then
            'Открываем диалог сохранения файла
            'Выбираем начальную директорию
            If fPath.Length > 0 Then
                Dim fName As String = System.IO.Path.GetFileName(fPath)
                fPath = fPath.Substring(0, fPath.Length - fName.Length)
            ElseIf DEEP_EDIT_MODE Then
                fPath = APP_HELP_PATH
            Else
                fPath = My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "help")
            End If
            sfd.InitialDirectory = fPath
            sfd.FileName = txtFuncName.Text.Trim + ".html"
            dRes = sfd.ShowDialog(Me)
            If dRes = Windows.Forms.DialogResult.OK Then
                'Сохраняем файл
                fPath = sfd.FileName
                My.Computer.FileSystem.WriteAllText(fPath, dlgGenerateHelp.codeHelp.Text, False, System.Text.Encoding.Default)
                Dim initDir As String = My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "help") + "\"
                If fPath.StartsWith(APP_HELP_PATH, StringComparison.CurrentCultureIgnoreCase) Then
                    fPath = fPath.Substring(APP_HELP_PATH.Length)
                ElseIf fPath.StartsWith(initDir, StringComparison.CurrentCultureIgnoreCase) Then
                    fPath = fPath.Substring(initDir.Length)
                End If
                fPath = fPath.Replace(APP_HELP_PATH, "")
                txtFuncHelpFile.Text = fPath
            End If
        End If
    End Sub

    ''' <summary>
    ''' Заполняет комбо списком свойтв для выбора свойства по умолчанию
    ''' </summary>
    ''' <param name="classId">Id класса</param>
    ''' <param name="cmb">комбобокс</param>
    Private Sub FillComboWithPropertiesForDefault(ByVal classId As Integer, ByRef cmb As ComboBox)
        cmb.Items.Clear()
        If IsNothing(mScript.mainClass(classId).Properties) OrElse mScript.mainClass(classId).Properties.Count = 0 Then Return

        For pId As Integer = 0 To mScript.mainClass(classId).Properties.Count - 1
            Dim hidden As MatewScript.PropertyHiddenEnum = mScript.mainClass(classId).Properties.ElementAt(pId).Value.Hidden
            If hidden = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL OrElse hidden = MatewScript.PropertyHiddenEnum.HIDDEN_IN_EDITOR Then Continue For
            cmb.Items.Add(mScript.mainClass(classId).Properties.ElementAt(pId).Key)
        Next pId
        cmb.Text = mScript.mainClass(classId).DefaultProperty
    End Sub

End Class