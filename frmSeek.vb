Public Class frmSeek

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click, treeElements.DoubleClick
        'Совершает поиск (и замену)
        If dlgEntrancies.Visible Then dlgEntrancies.Hide()
        Dim pnl As clsPanelManager.clsChildPanel = cPanelManager.ActivePanel
        If IsNothing(pnl) = False Then
            cPanelManager.HidePanel(pnl)
            cPanelManager.ActivePanel = Nothing
            If IsNothing(currentTreeView) = False Then currentTreeView.SelectedNode = Nothing
        End If

        If chkSearchElement.Checked Then
            'Поиск указанного элемента
            Dim n As TreeNode = treeElements.SelectedNode
            Dim ent As dlgEntrancies.cEntranciesClass = Nothing
            If IsNothing(n) = False Then ent = n.Tag
            If IsNothing(ent) Then
                MessageBox.Show("Не выбран элемент для поиска!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            Dim classAct As Integer = mScript.mainClassHash("A")

            If ent.child3Id < 0 Then
                If ent.classId = -3 Then
                    'Поиск функции Писателя
                    GlobalSeeker.CheckFunctionInStruct("'" + ent.elementName + "'", False)
                ElseIf ent.classId = -2 Then
                    'Поиск переменной
                    GlobalSeeker.CheckElementNameInStruct(-2, ent.elementName, CodeTextBox.EditWordTypeEnum.W_VARIABLE, True)
                ElseIf ent.classId = classAct Then
                    'Поиск действия
                    Dim curLocId As Integer = -1
                    If String.IsNullOrEmpty(actionsRouter.locationOfCurrentActions) = False Then curLocId = _
                        GetSecondChildIdByName(actionsRouter.locationOfCurrentActions, mScript.mainClass(mScript.mainClassHash("L")).ChildProperties)
                    If curLocId <> ent.parentId Then
                        'Если мы сейчас не на локации - родителе действия, то переходим к выбранному действию
                        Dim ch As clsPanelManager.clsChildPanel = cPanelManager.FindAndOpen(ent.classId, ent.child2Id, -1, ent.parentId, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                        If IsNothing(ch) Then
                            MessageBox.Show("Не удалось совершить переход к выбранным действиям для проведения поиска.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Return
                        End If
                    End If
                    GlobalSeeker.CheckActionInStruct("'" + ent.elementName + "'", mScript.mainClass(classAct).ChildProperties.Count - 1, True)
                Else
                    'Поиск элемента 2 уровня
                    GlobalSeeker.CheckChild2InStruct(ent.classId, "'" + ent.elementName + "'", False)
                End If
            Else
                'Поиск элемента 3 уровня
                Dim nPar As TreeNode = n.Parent
                Dim entPar As dlgEntrancies.cEntranciesClass = Nothing
                If IsNothing(n) = False Then entPar = nPar.Tag
                If IsNothing(entPar) Then
                    MessageBox.Show("Не удалось совершить поиск элемента. Ошибка в структуре.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If
                GlobalSeeker.CheckChild3InStruct(ent.classId, entPar.child2Id, "'" + ent.elementName + "'", False)
            End If
            If dlgEntrancies.hasEntrancies = False Then
                MessageBox.Show("Не найдено ни одного совпадения.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            'Поиск обычной строки
            If GlobalSeeker.FindStringInStruct(txtSearch.Text, chkWholeWord.Checked, chkCase.Checked) = False Then
                MessageBox.Show("Не найдено ни одного совпадения.", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
        If IsNothing(pnl) = False Then
            cPanelManager.OpenPanel(pnl)
        End If
    End Sub

    Private Sub frmSeek_Load(sender As Object, e As EventArgs) Handles Me.Load
        splitMain.FixedPanel = FixedPanel.Panel1
        splitMain.Panel2Collapsed = True
        Me.Height = btnFind.Bottom + questEnvironment.defPaddingTop + (Me.Height - Me.ClientSize.Height)
    End Sub

    Private Sub chkSearchElement_CheckedChanged(sender As Object, e As EventArgs) Handles chkSearchElement.CheckedChanged
        If chkSearchElement.Checked Then
            txtSearch.Enabled = False
            chkCase.Enabled = False
            chkWholeWord.Enabled = False
            splitMain.Panel2Collapsed = False
            Me.Height = 625 + (Me.Height - Me.ClientSize.Height)
            FillTreeWithElements()
        Else
            txtSearch.Enabled = True
            chkCase.Enabled = True
            chkWholeWord.Enabled = True
            splitMain.Panel2Collapsed = True
            Me.Height = btnFind.Bottom + questEnvironment.defPaddingTop + (Me.Height - Me.ClientSize.Height)
        End If
    End Sub

    Private Sub FillTreeWithElements()
        Dim tree As TreeView = treeElements
        tree.BeginUpdate()
        tree.Nodes.Clear()

        Dim classAct As Integer = mScript.mainClassHash("A")
        Dim classLoc As Integer = mScript.mainClassHash("L")

        For classId As Integer = 0 To mScript.mainClass.Count - 1
            If classId = classAct Then Continue For
            If mScript.mainClass(classId).LevelsCount = 0 Then Continue For
            If IsNothing(mScript.mainClass(classId).ChildProperties) OrElse mScript.mainClass(classId).ChildProperties.Count = 0 Then Continue For
            Dim nEl As TreeNode = tree.Nodes.Add(frmMainEditor.GetTranslatedName(classId))
            nEl.NodeFont = New Font(tree.Font, FontStyle.Bold)
            nEl.ForeColor = Color.Blue
            For child2Id As Integer = 0 To mScript.mainClass(classId).ChildProperties.Count - 1
                'Добавляем элементы 2 уровня
                Dim ch As MatewScript.ChildPropertiesInfoType = mScript.mainClass(classId).ChildProperties(child2Id)("Name")
                Dim ch2Name As String = mScript.PrepareStringToPrint(ch.Value, Nothing, False)
                Dim n As TreeNode = nEl.Nodes.Add(ch2Name, ch2Name)
                n.Tag = New dlgEntrancies.cEntranciesClass(classId, child2Id, -1, ch2Name, CodeTextBox.EditWordTypeEnum.W_PROPERTY, -1)

                If classId = classLoc Then
                    'Добавляем действия
                    Dim lstActions As List(Of String) = actionsRouter.GetActionsNames(child2Id)
                    If lstActions.Count = 0 Then Continue For
                    For actId As Integer = 0 To lstActions.Count - 1
                        Dim actName As String = mScript.PrepareStringToPrint(lstActions(actId), Nothing, False)
                        Dim n2 As TreeNode = n.Nodes.Add(actName, actName)
                        n2.Tag = New dlgEntrancies.cEntranciesClass(classAct, actId, -1, actName, CodeTextBox.EditWordTypeEnum.W_PROPERTY, child2Id)
                    Next actId
                ElseIf IsNothing(ch.ThirdLevelProperties) = False AndAlso _
                    ch.ThirdLevelProperties.Count > 0 Then
                    'Добавляем элементы 3 уровня
                    For child3Id = 0 To ch.ThirdLevelProperties.Count - 1
                        Dim ch3Name As String = mScript.PrepareStringToPrint(ch.ThirdLevelProperties(child3Id), Nothing, False)
                        Dim n2 As TreeNode = n.Nodes.Add(ch3Name, ch3Name)
                        n2.Tag = New dlgEntrancies.cEntranciesClass(classId, child2Id, child3Id, ch3Name, CodeTextBox.EditWordTypeEnum.W_PROPERTY)
                    Next child3Id
                End If
            Next child2Id
        Next classId

        'Добавляем функции
        If IsNothing(mScript.functionsHash) = False AndAlso mScript.functionsHash.Count > 0 Then
            Dim nFunc As TreeNode = tree.Nodes.Add("Функции Писателя")
            nFunc.NodeFont = New Font(tree.Font, FontStyle.Bold)
            nFunc.ForeColor = Color.Blue

            For fId As Integer = 0 To mScript.functionsHash.Count - 1
                Dim fName As String = mScript.functionsHash.ElementAt(fId).Key
                Dim n As TreeNode = nFunc.Nodes.Add(fName, fName)
                n.Tag = New dlgEntrancies.cEntranciesClass(-3, fId, -1, fName, CodeTextBox.EditWordTypeEnum.W_BLOCK_FUNCTION)
            Next fId
        End If

        'Добавляем переменные
        If IsNothing(mScript.csPublicVariables.lstVariables) = False AndAlso mScript.csPublicVariables.lstVariables.Count > 0 Then
            Dim nVar As TreeNode = tree.Nodes.Add("Переменные")
            nVar.NodeFont = New Font(tree.Font, FontStyle.Bold)
            nVar.ForeColor = Color.Blue
            For vId As Integer = 0 To mScript.csPublicVariables.lstVariables.Count - 1
                Dim vName As String = mScript.csPublicVariables.lstVariables.ElementAt(vId).Key
                Dim n As TreeNode = nVar.Nodes.Add(vName, vName)
                n.Tag = New dlgEntrancies.cEntranciesClass(-2, vId, -1, vName, CodeTextBox.EditWordTypeEnum.W_VARIABLE)
            Next vId
        End If

        tree.ExpandAll()
        tree.EndUpdate()
    End Sub
End Class