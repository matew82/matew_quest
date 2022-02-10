Imports System.Windows.Forms

Public Class dlgExecute


    Private Sub chkHTML_CheckedChanged(sender As Object, e As EventArgs) Handles chkHTML.CheckedChanged
        If chkHTML.Checked Then
            Dim eventHTML As String = mScript.ConvertCodeDataToHTML(codeMain.codeBox.CodeData, True) 'конвертируем в html
            codeHTML.Text = eventHTML
            codeMain.Hide()
            codeHTML.Show()
        Else
            codeHTML.Hide()
            codeMain.Show()
        End If
    End Sub

    Private Sub btnExecute_Click(sender As Object, e As EventArgs) Handles btnExecute.Click
        mScript.codeStack.Push("Произвольный скрипт")
        Dim sWatch As New Stopwatch
        'sWatch.Start()
        'Dim res As String = mScript.ExecuteCode(mScript.PrepareBlock(codeMain.codeBox.CodeData), Nothing, True)
        Dim res As String = mScript.ExecuteCode(mScript.PrepareBlock(codeMain.codeBox.Lines), Nothing, True)
        'sWatch.Stop()
        'MsgBox(Format(Math.Round(sWatch.ElapsedTicks / 1000), "#,#"))
        mScript.codeStack.Pop()
        lblResultText.Show()
        res = UnWrapString(res)
        mScript.LAST_ERROR = ""
        mScript.EXIT_CODE = False
        If res.Length = 0 Then
            lblResult.Text = "[empty]"
        Else
            lblResult.Text = res
        End If
        lblResult.Show()
    End Sub

    Private Sub dlgExecute_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        lblResult.Text = ""
        lblResult.Hide()
        lblResult.Hide()
        'chkHTML.Checked = False
    End Sub

    Private Sub dlgExecute_Load(sender As Object, e As EventArgs) Handles Me.Load
        codeHTML.Text = ""
        codeMain.Text = ""
        codeMain.codeBox.Visible = True
    End Sub
End Class
