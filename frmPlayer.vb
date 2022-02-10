'Imports System.Security.Permissions
'<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
<System.Runtime.InteropServices.ComVisible(True)>
Public Class frmPlayer
    Private WithEvents hDocumentMain As HtmlDocument, hDocumentObjects As HtmlDocument

    Private Sub wbMain_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbMain.DocumentCompleted
        hDocumentMain = wbMain.Document
    End Sub

    Private Sub wbObjects_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbObjects.DocumentCompleted
        hDocumentObjects = wbObjects.Document
    End Sub

    Private Sub frmPlayer_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        TimerRemove(-1)
        timAudio.Enabled = False
        mPlayer.Stop()
        mPlayer.Volume = 0
        mapManager.ClearPreviousBitmap()
        mScript.Battle.StopBattleAbrupt()

        If IsNothing(frmMap) = False Then
            frmMap.CanBeClosed = True
            frmMap.Dispose()
        End If
        If IsNothing(frmMagic) = False Then
            frmMagic.CanBeClosed = True
            frmMagic.Dispose()
        End If

        If questEnvironment.OPENED_FROM_EDITOR Then
            questEnvironment.EDIT_MODE = True
            Dim qPath As String = questEnvironment.QuestPath
            Me.Hide()
            questEnvironment.ClearAllData()
            questEnvironment.LoadQuest(qPath)
            frmMainEditor.Show()
        End If
    End Sub


    Private Sub frmPlayer_Load(sender As Object, e As EventArgs) Handles Me.Load
        wbMain.ObjectForScripting = Me
    End Sub

    Public Sub wbMain_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles wbMain.PreviewKeyDown, wbActions.PreviewKeyDown, wbCommand.PreviewKeyDown, wbDescription.PreviewKeyDown, _
        wbObjects.PreviewKeyDown
        If e.KeyCode = Keys.F5 Then e.IsInputKey = True
    End Sub

    ''' <summary>
    ''' Процедура, вызывающая событие LocationHtmlEvent
    ''' </summary>
    ''' <param name="args"></param>
    Public Sub LocationHtmlEvent(ByVal args As String)
        Dim arr() As String = Split(args, Chr(1))

        For i As Integer = 0 To arr.Count - 1
            Dim carr = arr(i)
            If (String.Compare(carr, "true", True) = 0 OrElse String.Compare(carr, "false", True) = 0 OrElse IsNumeric(carr.Replace(".", ",")) OrElse carr.StartsWith("'")) = False Then
                arr(i) = WrapString(carr)
            End If
        Next

        Dim eventId As Integer = mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(GVARS.G_CURLOC)("LocationHtmlEvent").eventId
        If eventId > 0 Then mScript.eventRouter.RunEvent(eventId, arr, "LocationHtmlEvent", True)
    End Sub

    Private Sub hDocumentMain_Click(sender As Object, e As HtmlElementEventArgs) Handles hDocumentMain.Click
        'Выполняем событие LocationHtmlEvent для элементов, имеющих атрибут send="param1, param2, ..., paramN"
        Dim eventId As Integer = mScript.mainClass(mScript.mainClassHash("L")).ChildProperties(GVARS.G_CURLOC)("LocationHtmlEvent").eventId
        If eventId <= 0 Then Return

        If wbMain.ReadyState <> WebBrowserReadyState.Complete Then Return
        Dim hEl As HtmlElement = hDocumentMain.GetElementFromPoint(e.ClientMousePosition)

        Do
            'перебираем все html-элементы, начиная с кликнутого, и включая всех его родителей
            If IsNothing(hEl) OrElse hEl.TagName = "BODY" Then Exit Do
            Dim strSend As String = hEl.GetAttribute("Send")
            If String.IsNullOrEmpty(strSend) = False Then
                'у html-элемента есть атрибут send
                'получаем строку (param1, param2...)
                Dim pos As Integer = -1
                Dim arr As New List(Of String)
                Dim lastPos As Integer = 0
                Do
                    'получаем массив параметров
                    pos = strSend.IndexOfAny({","c, "'"c}, pos + 1)
                    If pos = -1 Then
                        'последний параметр
                        arr.Add(strSend.Substring(lastPos, strSend.Length - lastPos).Trim)
                        Exit Do
                    End If
                    Dim ch As Char = strSend.Chars(pos)
                    If ch = "," Then
                        'новый параметр
                        arr.Add(strSend.Substring(lastPos, pos - lastPos).Trim)
                        lastPos = pos + 1
                    ElseIf ch = "'"c Then
                        'это строка, в которой могут быть запятые (которые не учитываются)
                        Dim pos2 As Integer = pos
                        Do
                            pos2 = strSend.IndexOf("'"c, pos2 + 1)
                            If pos2 = -1 Then
                                'незакрытая кавычка - ошибка аргумента. Код все-равно надо запустить для получения номера строки с ошибкой
                                arr.Add(strSend.Substring(lastPos, strSend.Length - lastPos).Trim)
                                Exit Do
                            End If
                            If pos2 > 0 AndAlso strSend.Chars(pos2 - 1) = "/"c Then
                                'экранированная кавычка
                                Continue Do
                            End If
                            'найдена закрывающая кавычка
                            pos = pos2
                            Exit Do
                        Loop
                        If pos = -1 Then Exit Do
                    End If
                Loop

                For i As Integer = 0 To arr.Count - 1
                    Dim carr = arr(i)
                    If (String.Compare(carr, "true", True) = 0 OrElse String.Compare(carr, "false", True) = 0 OrElse IsNumeric(carr.Replace(".", ",")) OrElse carr.StartsWith("'")) = False Then
                        arr(i) = WrapString(carr)
                    End If
                Next

                'запуск самого события
                If mScript.eventRouter.RunEvent(eventId, arr.ToArray, "LocationHtmlEvent", True) = "#Error" Then Return
            End If
            hEl = hEl.Parent
        Loop

    End Sub

    Public Sub FlipMagicBookEvent()
        'получаем mBookId
        Dim hDoc As HtmlDocument = questEnvironment.wbMagic.Document
        If IsNothing(hDoc) Then Return
        Dim hBook As HtmlElement = hDoc.GetElementById("magicBook")
        If IsNothing(hBook) Then Return
        If hBook.Children.Count > 0 Then hBook = hBook.Children(0) 'Page
        If hBook.Children.Count > 0 Then hBook = hBook.Children(0) 'MPage magicsPageFirst
        If hBook.Children.Count > 0 Then hBook = hBook.Children(0).Children(0) 'TBODY
        If hBook.Children.Count > 0 Then hBook = hBook.Children(0) 'TR
        If hBook.TagName <> "TR" Then Return
        Dim mBookId As Integer = Val(hBook.GetAttribute("magicBookId"))

        Dim classId As Integer = mScript.mainClassHash("Mg")
        Dim arrs() As String = {mBookId.ToString, "-1"}
        Dim sndFlip As String = ""
        If ReadProperty(classId, "FlipSound", mBookId, -1, sndFlip, arrs) = False Then Return
        sndFlip = UnWrapString(sndFlip)
        If String.IsNullOrEmpty(sndFlip) Then Return

        'убеждаемся в правильном формате пути к звуку
        sndFlip = sndFlip.Replace("\", "/")
        Dim fPath As String = FileIO.FileSystem.CombinePath(questEnvironment.QuestPath, sndFlip)
        If FileIO.FileSystem.FileExists(fPath) = False Then
            MessageBox.Show("Файл " & fPath & " не найден!", "MatewQuest", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'проигрывем звук
        Say(sndFlip, 100, False)
    End Sub
End Class