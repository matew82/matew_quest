<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgExecute
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.splitMain = New System.Windows.Forms.SplitContainer()
        Me.codeMain = New WindowsApplication1.CodeTextBox()
        Me.codeHTML = New WindowsApplication1.CodeTextBox()
        Me.lblResult = New System.Windows.Forms.Label()
        Me.lblResultText = New System.Windows.Forms.Label()
        Me.chkHTML = New System.Windows.Forms.CheckBox()
        Me.btnExecute = New System.Windows.Forms.Button()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'splitMain
        '
        Me.splitMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitMain.Location = New System.Drawing.Point(0, 0)
        Me.splitMain.Name = "splitMain"
        Me.splitMain.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitMain.Panel1
        '
        Me.splitMain.Panel1.Controls.Add(Me.codeMain)
        Me.splitMain.Panel1.Controls.Add(Me.codeHTML)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.lblResult)
        Me.splitMain.Panel2.Controls.Add(Me.lblResultText)
        Me.splitMain.Panel2.Controls.Add(Me.chkHTML)
        Me.splitMain.Panel2.Controls.Add(Me.btnExecute)
        Me.splitMain.Panel2MinSize = 60
        Me.splitMain.Size = New System.Drawing.Size(1395, 632)
        Me.splitMain.SplitterDistance = 568
        Me.splitMain.TabIndex = 0
        '
        'codeMain
        '
        Me.codeMain.AutoWordSelection = False
        Me.codeMain.codeBox.CanDrawWords = True
        Me.codeMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.codeMain.codeBox.DontShowError = False
        Me.codeMain.codeBox.HelpFile = "file:///D:/Projects/MatewQuest2/MatewQuest2/bin/Debug/src/rtbHelp.html"
        Me.codeMain.codeBox.HelpPath = "D:\Projects\q\Help\"
        Me.codeMain.codeBox.IsTextBlockByDefault = False
        Me.codeMain.Location = New System.Drawing.Point(0, 0)
        Me.codeMain.Margin = New System.Windows.Forms.Padding(2)
        Me.codeMain.Multiline = True
        Me.codeMain.Name = "codeMain"
        Me.codeMain.Size = New System.Drawing.Size(1395, 568)
        Me.codeMain.TabIndex = 1
        '
        'codeHTML
        '
        Me.codeHTML.AutoWordSelection = False
        Me.codeHTML.codeBox.CanDrawWords = True
        Me.codeHTML.Dock = System.Windows.Forms.DockStyle.Fill
        Me.codeHTML.codeBox.DontShowError = False
        Me.codeHTML.codeBox.HelpFile = "file:///D:/Projects/MatewQuest2/MatewQuest2/bin/Debug/src/rtbHelp.html"
        Me.codeHTML.codeBox.HelpPath = "D:\Projects\q\Help\"
        Me.codeHTML.codeBox.IsTextBlockByDefault = True
        Me.codeHTML.Location = New System.Drawing.Point(0, 0)
        Me.codeHTML.Margin = New System.Windows.Forms.Padding(2)
        Me.codeHTML.Multiline = True
        Me.codeHTML.Name = "codeHTML"
        Me.codeHTML.Size = New System.Drawing.Size(1395, 568)
        Me.codeHTML.TabIndex = 2
        Me.codeHTML.Visible = False
        '
        'lblResult
        '
        Me.lblResult.AutoSize = True
        Me.lblResult.BackColor = System.Drawing.Color.White
        Me.lblResult.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblResult.ForeColor = System.Drawing.Color.Crimson
        Me.lblResult.Location = New System.Drawing.Point(680, 19)
        Me.lblResult.Name = "lblResult"
        Me.lblResult.Size = New System.Drawing.Size(17, 25)
        Me.lblResult.TabIndex = 3
        Me.lblResult.Text = "-"
        Me.lblResult.Visible = False
        '
        'lblResultText
        '
        Me.lblResultText.AutoSize = True
        Me.lblResultText.Location = New System.Drawing.Point(512, 19)
        Me.lblResultText.Name = "lblResultText"
        Me.lblResultText.Size = New System.Drawing.Size(162, 23)
        Me.lblResultText.TabIndex = 2
        Me.lblResultText.Text = "Результат скрипта: "
        Me.lblResultText.Visible = False
        '
        'chkHTML
        '
        Me.chkHTML.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkHTML.Image = Global.WindowsApplication1.My.Resources.Resources.toHTML32
        Me.chkHTML.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.chkHTML.Location = New System.Drawing.Point(232, 6)
        Me.chkHTML.Name = "chkHTML"
        Me.chkHTML.Size = New System.Drawing.Size(206, 49)
        Me.chkHTML.TabIndex = 1
        Me.chkHTML.Text = "Просмотреть в html"
        Me.chkHTML.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkHTML.UseVisualStyleBackColor = True
        '
        'btnExecute
        '
        Me.btnExecute.Image = Global.WindowsApplication1.My.Resources.Resources.Execute2
        Me.btnExecute.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExecute.Location = New System.Drawing.Point(12, 6)
        Me.btnExecute.Name = "btnExecute"
        Me.btnExecute.Size = New System.Drawing.Size(203, 49)
        Me.btnExecute.TabIndex = 0
        Me.btnExecute.Text = "Выполнить скрипт"
        Me.btnExecute.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnExecute.UseVisualStyleBackColor = True
        '
        'dlgExecute
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 23.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1395, 632)
        Me.Controls.Add(Me.splitMain)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Italic)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MinimumSize = New System.Drawing.Size(1411, 671)
        Me.Name = "dlgExecute"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Выполнить скрипт..."
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel2.ResumeLayout(False)
        Me.splitMain.Panel2.PerformLayout()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents codeMain As WindowsApplication1.CodeTextBox
    Friend WithEvents codeHTML As WindowsApplication1.CodeTextBox
    Friend WithEvents btnExecute As System.Windows.Forms.Button
    Friend WithEvents chkHTML As System.Windows.Forms.CheckBox
    Friend WithEvents lblResult As System.Windows.Forms.Label
    Friend WithEvents lblResultText As System.Windows.Forms.Label

End Class
