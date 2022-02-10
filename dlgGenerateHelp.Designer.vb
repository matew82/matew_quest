<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgGenerateHelp
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
        Me.codeHelp = New WindowsApplication1.CodeTextBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnShowScript = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'splitMain
        '
        Me.splitMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.splitMain.Location = New System.Drawing.Point(0, 0)
        Me.splitMain.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.splitMain.Name = "splitMain"
        Me.splitMain.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitMain.Panel1
        '
        Me.splitMain.Panel1.Controls.Add(Me.codeHelp)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.splitMain.Panel2.Controls.Add(Me.btnCancel)
        Me.splitMain.Size = New System.Drawing.Size(1103, 531)
        Me.splitMain.SplitterDistance = 437
        Me.splitMain.SplitterWidth = 6
        Me.splitMain.TabIndex = 0
        '
        'codeHelp
        '
        Me.codeHelp.AutoWordSelection = False
        Me.codeHelp.codeBox.CanDrawWords = True
        Me.codeHelp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.codeHelp.codeBox.DontShowError = False
        Me.codeHelp.codeBox.HelpFile = "file:///D:/Projects/MatewQuest2/MatewQuest2/bin/Debug/src/rtbHelp.html"
        Me.codeHelp.codeBox.HelpPath = "D:\Projects\q\Help\"
        Me.codeHelp.codeBox.IsTextBlockByDefault = True
        Me.codeHelp.Location = New System.Drawing.Point(0, 0)
        Me.codeHelp.Margin = New System.Windows.Forms.Padding(2)
        Me.codeHelp.Multiline = True
        Me.codeHelp.Name = "codeHelp"
        Me.codeHelp.Size = New System.Drawing.Size(1103, 437)
        Me.codeHelp.TabIndex = 1
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 235.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 90.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnSave, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.btnShowScript, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(774, 10)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(325, 61)
        Me.TableLayoutPanel1.TabIndex = 21
        '
        'btnSave
        '
        Me.btnSave.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnSave.Image = Global.WindowsApplication1.My.Resources.Resources.save32
        Me.btnSave.Location = New System.Drawing.Point(242, 3)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 55)
        Me.btnSave.TabIndex = 0
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnShowScript
        '
        Me.btnShowScript.Image = Global.WindowsApplication1.My.Resources.Resources.execte
        Me.btnShowScript.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnShowScript.Location = New System.Drawing.Point(3, 3)
        Me.btnShowScript.Name = "btnShowScript"
        Me.btnShowScript.Size = New System.Drawing.Size(227, 55)
        Me.btnShowScript.TabIndex = 2
        Me.btnShowScript.Text = "Вставить пример скрипта"
        Me.btnShowScript.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnShowScript.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Image = Global.WindowsApplication1.My.Resources.Resources.delete32
        Me.btnCancel.Location = New System.Drawing.Point(12, 10)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 56)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'dlgGenerateHelp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1103, 531)
        Me.Controls.Add(Me.splitMain)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "dlgGenerateHelp"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Генерация файла помощи"
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel2.ResumeLayout(False)
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents codeHelp As WindowsApplication1.CodeTextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnShowScript As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel

End Class
