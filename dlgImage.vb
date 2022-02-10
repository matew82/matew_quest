Imports System.Windows.Forms

Public Class dlgImage
    Public selectedPath As String = ""
    Public selectedImage As String = ""
    Private hDocument As HtmlDocument
    Dim nodeUnderMouseToDrop As TreeNode, nodeUnderMouseToDrag As TreeNode

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        BuildImageString()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgImage_Load(sender As Object, e As EventArgs) Handles Me.Load
        splitMain.FixedPanel = FixedPanel.Panel2
        wbImages.Navigate(FileIO.FileSystem.CombinePath(Application.StartupPath, "src\editor.html"))
        treePath.ImageList = frmMainEditor.imgLstGroupIcons
    End Sub

    ''' <summary>
    ''' Возвращает строку html-кода img
    ''' </summary>
    Private Sub BuildImageString()
        selectedImage = ""
        If String.IsNullOrEmpty(selectedPath) Then Return

        selectedImage = "<img "

        If txtWidth.TextLength = 0 AndAlso txtHeight.TextLength = 0 AndAlso optFloatNo.Checked Then
            'не указана ни ширина, ни высота, ни прилегание - style= не будет
            selectedImage &= "src=" & Chr(34) & selectedPath & Chr(34) & " />"
            Return
        End If

        'ширину, высоту и прилегание устанавливаем в атрибуте style
        selectedImage &= "style=" & Chr(34)
        If txtWidth.TextLength > 0 Then
            selectedImage &= "width:" & txtWidth.Text & ";"
        End If

        If txtHeight.TextLength > 0 Then
            selectedImage &= "height:" & txtHeight.Text & ";"
        End If

        If optFloatLeft.Checked Then
            selectedImage &= "float:left;"
        ElseIf optFloatRight.Checked Then
            selectedImage &= "float:right;"
        End If
        selectedImage = selectedImage.Substring(0, selectedImage.Length - 1) & Chr(34) & " src=" & Chr(34) & selectedPath & Chr(34) & " />"

    End Sub


    Private Sub wbImages_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbImages.DocumentCompleted
        hDocument = wbImages.Document
        If IsNothing(hDocument) Then Return

        Dim hSample As HtmlElement = hDocument.CreateElement("DIV")
        hSample.Id = "sampleContainer"
        hDocument.Body.AppendChild(hSample)
        Dim hData As HtmlElement = hDocument.CreateElement("DIV")
        hData.Id = "dataContainer"
        hDocument.Body.AppendChild(hData)


        'Путь к изображению
        If cListManager.HasFiles(cListManagerClass.fileTypesEnum.PICTURES) Then
            frmMainEditor.FillTreeWithFolders(treePath, cListManagerClass.fileTypesEnum.PICTURES) 'Создаем узлы - папки с изображениями

            If treePath.Nodes.Count = 0 Then Return
        Else
            treePath.Nodes.Clear()
        End If
    End Sub

    Private Sub treePath_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles treePath.AfterSelect
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
    End Sub

    Private Sub treePath_BeforeSelect(sender As Object, e As TreeViewCancelEventArgs) Handles treePath.BeforeSelect
        Dim n As TreeNode = sender.SelectedNode
        If IsNothing(n) = False Then
            If n.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then
                n.ForeColor = DEFAULT_COLORS.NodeForeColor
                n.BackColor = DEFAULT_COLORS.ControlTransparentBackground
            End If
        End If
        n = e.Node
        If IsNothing(n) = False Then
            If n.ForeColor <> DEFAULT_COLORS.NodeHiddenForeColor Then
                n.ForeColor = DEFAULT_COLORS.NodeSelForeColor
                n.BackColor = DEFAULT_COLORS.NodeSelBackColor
            End If
        End If
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
        selectedPath = oldImg.GetAttribute("rVal")
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
        hImg.SetAttribute("rVal", imgPath.Replace("\", "/"))
        hImg.SetAttribute("className", "thumbnail selectable")
        Dim aPos As Integer = imgPath.LastIndexOf("\")
        If aPos = -1 OrElse aPos = imgPath.Length - 1 Then
            hImg.SetAttribute("Title", imgPath)
        Else
            hImg.SetAttribute("Title", imgPath.Substring(aPos + 1))
        End If
        hContainer.AppendChild(hImg)

        AddHandler hImg.Click, AddressOf del_html_img_click
    End Sub

    Private Sub treePath_DoubleClick(sender As Object, e As EventArgs) Handles treePath.DoubleClick
        Dim path As String = questEnvironment.QuestPath + "\" + treePath.SelectedNode.Name + "\"
        Process.Start(path)
    End Sub

    Dim draggedFiles() As String = Nothing
    Private Sub treePath_DragDrop(sender As Object, e As DragEventArgs) Handles treePath.DragDrop
        If IsNothing(draggedFiles) Then Return
        If e.Data.GetDataPresent("FileDrop") = False Then Return
        If IsNothing(nodeUnderMouseToDrop) Then Return
        Dim selNode As TreeNode = treePath.SelectedNode
        If Object.Equals(selNode, nodeUnderMouseToDrop) = False Then Return
        If e.Effect <> DragDropEffects.Move Then Return

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
            AddHTMLimgToContainer(Nothing, selNode.Name + "\" + fName)
        Next
    End Sub

    Private Sub treePath_DragEnter(sender As Object, e As DragEventArgs) Handles treePath.DragEnter
        'начало операции перетаскивания файлов
        If e.Data.GetDataPresent("FileDrop") = False Then Return
        Dim files() As String = e.Data.GetData("FileDrop")
        If files.Count = 0 Then Return
        'перетаскивают именно файлы
        If files(0).ToLower.StartsWith(questEnvironment.QuestPath.ToLower) Then Return 'из директории квеста не перетаскиваются

        'получаем список допустимых расширений перетаскиваемых файлов
        Dim lstExtensions As List(Of String) = {".png", ".jpg", ".gif", ".jpeg", ".bmp", ".wmf"}.ToList

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
        treePath.DoDragDrop(e.Data, DragDropEffects.Copy)
    End Sub

    Private Sub treePath_DragOver(sender As Object, e As DragEventArgs) Handles treePath.DragOver
        If IsNothing(draggedFiles) Then Return
        If ((e.AllowedEffect And DragDropEffects.Move) = DragDropEffects.Move) Then
            ' By default, the drop action should be move, if allowed.
            e.Effect = DragDropEffects.Move
        Else
            e.Effect = DragDropEffects.None
        End If

        'получаем узел, который под мышью на данный момент
        Dim prevNodeDrop As TreeNode = nodeUnderMouseToDrop
        nodeUnderMouseToDrop = treePath.GetNodeAt(treePath.PointToClient(New Point(e.X, e.Y)))
        If IsNothing(nodeUnderMouseToDrop) = False AndAlso IsNothing(nodeUnderMouseToDrop.Tag) = False AndAlso nodeUnderMouseToDrop.Tag.ToString = "ITEM" Then _
            nodeUnderMouseToDrop = nodeUnderMouseToDrop.Parent
        If Object.Equals(treePath.SelectedNode, nodeUnderMouseToDrop) Then Return
        treePath.SelectedNode = nodeUnderMouseToDrop
    End Sub

    Private Sub treePath_QueryContinueDrag(sender As Object, e As QueryContinueDragEventArgs) Handles treePath.QueryContinueDrag
        'Организация DragDop - перемещение функций и свойств между классами
        If e.EscapePressed Then
            'При нажатии Esc прекращаем перетаскивание
            e.Action = DragAction.Cancel
            nodeUnderMouseToDrag = Nothing
            Erase draggedFiles
        End If
    End Sub
End Class
