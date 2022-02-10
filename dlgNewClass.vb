Imports System.Windows.Forms

Public Class dlgNewClass
    ''' <summary>
    ''' Иконка класса
    ''' </summary>
    Public classIcon As Bitmap
    ''' <summary>
    ''' Id редактируемого класса
    ''' </summary>
    ''' <remarks></remarks>
    Dim editClassId As Integer = -1

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim strNames As String = txtNewClassName.Text.Trim
        If strNames.Length = 0 Then
            MsgBox("Укажите хоть одно имя класса.", MsgBoxStyle.Exclamation)
            txtNewClassName.Focus()
            Exit Sub
        End If

        Dim arrNames() As String = Split(strNames, ",")
        For i As Integer = 0 To arrNames.GetUpperBound(0)
            arrNames(i) = arrNames(i).Trim
            If arrNames(i).Length = 0 Then
                MsgBox("Уберите лишнюю запятую в именах класса.", MsgBoxStyle.Exclamation)
                txtNewClassName.Focus()
                Exit Sub
            End If
            If arrNames(i).IndexOf(" "c) > 0 Then
                MsgBox("Имя класса не может содержать пробелы.", MsgBoxStyle.Exclamation)
                txtNewClassName.Focus()
                Exit Sub
            End If
            If editClassId = -1 Then
                If mScript.mainClassHash.ContainsKey(arrNames(i)) Then
                    MsgBox("Имя " + arrNames(i) + " уже существует.", MsgBoxStyle.Exclamation)
                    txtNewClassName.Focus()
                    Exit Sub
                End If
            Else
                For q As Integer = 0 To arrNames.Count - 1
                    Dim cls As Integer = -1
                    If mScript.mainClassHash.TryGetValue(arrNames(i), cls) Then
                        If cls <> editClassId Then
                            MsgBox("Имя " + arrNames(i) + " уже существует.", MsgBoxStyle.Exclamation)
                            txtNewClassName.Focus()
                            Exit Sub
                        End If
                    End If
                Next
            End If
        Next

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgNewClass_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ofd.InitialDirectory = Application.StartupPath
    End Sub

    Private Sub btnClassIcon_Click(sender As Object, e As EventArgs) Handles btnClassIcon.Click
        Dim dRes As DialogResult = ofd.ShowDialog(Me)
        If dRes = Windows.Forms.DialogResult.Cancel Then Return

        Dim fs As New System.IO.FileStream(ofd.FileName, IO.FileMode.Open, IO.FileAccess.Read)
        Dim bm As Bitmap = Image.FromStream(fs)
        fs.Close()
        classIcon = New Bitmap(48, 48)
        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(classIcon)
        g.DrawImage(bm, 0, 0, 48, 48)
        pbIcon.Image = classIcon
    End Sub

    Public Sub PrepareForNewClass()
        editClassId = -1
        txtNewClassName.Clear()
        txtClassHelpFile.Text = ""
        cmbClassDefProperty.Items.Clear()
        nudClassLevels.Value = 3
        classIcon = Nothing
        pbIcon.Image = Nothing
        tlpRemove.Hide()
    End Sub

    Public Sub PrepareForEditClass(ByVal classId As Integer)
        editClassId = classId
        txtNewClassName.Text = Join(mScript.mainClass(classId).Names, ", ")
        nudClassLevels.Value = mScript.mainClass(classId).LevelsCount + 1
        txtClassHelpFile.Text = mScript.mainClass(classId).HelpFile
        FillComboWithPropertiesForDefault(classId, cmbClassDefProperty)

        Dim imgPath As String = questEnvironment.QuestPath + "\img\classMenu\" + mScript.mainClass(classId).Names(0) + ".png"
        If My.Computer.FileSystem.FileExists(imgPath) Then
            Dim fs As New System.IO.FileStream(imgPath, IO.FileMode.Open, IO.FileAccess.Read)
            classIcon = Image.FromStream(fs) ' Image.FromFile(imgPath)
            fs.Close()
        Else
            classIcon = Nothing
        End If
        pbIcon.Image = classIcon
        tlpRemove.Show()

    End Sub

    Private Sub btnRemoveClass_Click(sender As Object, e As EventArgs) Handles btnRemoveClass.Click
        Dim dRes As DialogResult = MessageBox.Show("Действие необратимо. Удалить класс со всем содержимым?", "Matew Quest", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If dRes = Windows.Forms.DialogResult.No Then Return
        'Закрываем текущую вкладку
        frmMainEditor.ChangeCurrentClass("Q")
        'Прочесывание структуры mainClass на предмет наличия функция и свойств, содержащих в качестве возвращаемого значения или параметра функции тип ELEMENT нашего удаляемого класса. Если будет такое найдено - очищаем формат
        GlobalSeeker.RemoveClassNameInStruct(editClassId)

        mScript.trackingProperties.RemoveClass(mScript.mainClass(editClassId).Names(0)) 'удаляем из событий изменений свойств

        'Удаляем иконку
        Dim iconPath As String = questEnvironment.QuestPath + "\img\classMenu\" + mScript.mainClass(editClassId).Names(0) + ".png"
        If My.Computer.FileSystem.FileExists(iconPath) Then
            Try
                My.Computer.FileSystem.DeleteFile(iconPath)
            Catch ex As Exception
                MessageBox.Show("Не удалось удалить иконку класса " + iconPath + ". Попробуйте сделать это вручную.", "Matew Quest", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Try
        End If
        'Выводим меню иконок из дерева или поддерева, чтобы они случайно не удалились
        If frmMainEditor.Controls.Contains(iconMenuElements) = False Then frmMainEditor.Controls.Add(iconMenuElements)
        If frmMainEditor.Controls.Contains(iconMenuGroups) = False Then frmMainEditor.Controls.Add(iconMenuGroups)
        'Удаляем дерево
        If frmMainEditor.dictTrees.ContainsKey(editClassId) Then
            frmMainEditor.dictTrees(editClassId).Dispose()
            frmMainEditor.dictTrees.Remove(editClassId)
        End If

        If editClassId < mScript.mainClass.Count - 1 Then
            'пересоздаем список деревьев классов
            Dim dictNew As New Dictionary(Of Integer, TreeView)
            For i As Integer = 0 To frmMainEditor.dictTrees.Count - 1
                Dim oldId As Integer = frmMainEditor.dictTrees.ElementAt(i).Key
                Dim newId As Integer = oldId
                If oldId > editClassId Then newId -= 1
                dictNew.Add(newId, frmMainEditor.dictTrees(oldId))
            Next i
            frmMainEditor.dictTrees.Clear()
            frmMainEditor.dictTrees = dictNew
        End If

        'Удаляем кнопку на тулстрипе
        For i As Integer = 0 To frmMainEditor.ToolStripMain.Items.Count - 1
            Dim tsi As ToolStripItem = frmMainEditor.ToolStripMain.Items(i)
            If IsNothing(tsi.Tag) = False AndAlso tsi.Tag = mScript.mainClass(editClassId).Names(0) Then
                frmMainEditor.ToolStripMain.Items.Remove(tsi)
                Exit For
            End If
        Next
        'Удаляем группы
        If cGroups.dictGroups.ContainsKey(mScript.mainClass(editClassId).Names(0)) Then
            cGroups.dictGroups.Remove(mScript.mainClass(editClassId).Names(0))
        End If
        'Удаляем панели где класс = удаляемому, изменяем Id класса если он больше Id удаляемого
        If IsNothing(cPanelManager.lstPanels) = False AndAlso cPanelManager.lstPanels.Count > 0 Then
            For i As Integer = cPanelManager.lstPanels.Count - 1 To 0 Step -1
                Dim pnl As clsPanelManager.clsChildPanel = cPanelManager.lstPanels(i)
                If pnl.classId = editClassId Then
                    cPanelManager.RemovePanel(pnl, False)
                ElseIf pnl.classId > editClassId Then
                    pnl.classId -= 1
                End If
            Next i
        End If
        'Очищаем lastPanel аналогичным образом
        If cPanelManager.dictLastPanel.ContainsKey(editClassId) Then cPanelManager.dictLastPanel.Remove(editClassId)

        If IsNothing(cPanelManager.dictLastPanel) = False AndAlso cPanelManager.dictLastPanel.Count > 0 Then
            For i As Integer = cPanelManager.dictLastPanel.Count - 1 To 0 Step -1
                Dim lId As Integer = cPanelManager.dictLastPanel.ElementAt(i).Key
                If lId > editClassId Then
                    Dim ch As clsPanelManager.clsChildPanel = cPanelManager.dictLastPanel(lId)
                    cPanelManager.dictLastPanel.Remove(lId)
                    cPanelManager.dictLastPanel.Add(lId - 1, ch)
                End If
            Next i
        End If

        'Удаляем панель
        If IsNothing(cPanelManager.dictDefContainers) = False AndAlso cPanelManager.dictDefContainers.Count > 0 Then
            If cPanelManager.dictDefContainers.ContainsKey(editClassId) Then
                cPanelManager.dictDefContainers(editClassId).Dispose()
                cPanelManager.dictDefContainers.Remove(editClassId)
            End If
            'В списке панелей, где индекс был больше editClassId, уменьшаем на 1
            For i As Integer = cPanelManager.dictDefContainers.Count - 1 To 0 Step -1
                Dim pnl As clsPanelManager.PanelEx = cPanelManager.dictDefContainers.ElementAt(i).Value
                Dim pnlId As Integer = cPanelManager.dictDefContainers.ElementAt(i).Key
                If pnlId > editClassId Then
                    cPanelManager.dictDefContainers.Remove(pnlId)
                    cPanelManager.dictDefContainers.Add(pnlId - 1, pnl)
                End If
            Next
        End If
        'Удаляем из класса mainClass
        Dim lst As List(Of MatewScript.MainClassType) = mScript.mainClass.ToList
        lst.RemoveAt(editClassId)
        mScript.mainClass = lst.ToArray
        'Вносим изменения в удаленные элементы
        removedObjects.RemoveClass(editClassId)
        'Вносим изменения в кнопки навигации
        If IsNothing(cPanelManager.dictDefContainers) = False Then
            For i As Integer = 0 To cPanelManager.dictDefContainers.Count - 1
                Dim pnl As clsPanelManager.PanelEx = cPanelManager.dictDefContainers.ElementAt(i).Value
                If IsNothing(pnl.NavigationButton) Then Continue For
                Dim obj As Object = pnl.NavigationButton
                If obj.NavigateToClassId > editClassId Then
                    obj.NavigateToClassId -= 1
                End If
            Next i
        End If

        mScript.MakeMainClassHash()
        mScript.FillFuncAndPropHash()
        Me.DialogResult = System.Windows.Forms.DialogResult.Abort
        Me.Close()
    End Sub

    Private Sub btnClassHelpFile_Click(sender As Object, e As EventArgs) Handles btnClassHelpFile.Click
        'получение имени файла помощи для класса
        ofd.ShowDialog()
        If ofd.FileName.Length = 0 Then Exit Sub
        Dim strFileName = ofd.FileName
        If strFileName.StartsWith(APP_HELP_PATH) Then strFileName = strFileName.Substring(APP_HELP_PATH.Length)
        txtClassHelpFile.Text = strFileName
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
