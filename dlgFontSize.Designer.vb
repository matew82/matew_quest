<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgFontSize
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.splitMain = New System.Windows.Forms.SplitContainer()
        Me.wbFont = New System.Windows.Forms.WebBrowser()
        Me.chkUnderline = New System.Windows.Forms.CheckBox()
        Me.chkBold = New System.Windows.Forms.CheckBox()
        Me.chkItalic = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbFontName = New System.Windows.Forms.ComboBox()
        Me.cmbPoints = New System.Windows.Forms.ComboBox()
        Me.cmbSize = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(684, 31)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(219, 45)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(4, 5)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(100, 35)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "ОК"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(114, 5)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(100, 35)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Отмена"
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
        Me.splitMain.Panel1.Controls.Add(Me.wbFont)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.chkUnderline)
        Me.splitMain.Panel2.Controls.Add(Me.chkBold)
        Me.splitMain.Panel2.Controls.Add(Me.chkItalic)
        Me.splitMain.Panel2.Controls.Add(Me.Label2)
        Me.splitMain.Panel2.Controls.Add(Me.cmbFontName)
        Me.splitMain.Panel2.Controls.Add(Me.cmbPoints)
        Me.splitMain.Panel2.Controls.Add(Me.cmbSize)
        Me.splitMain.Panel2.Controls.Add(Me.Label1)
        Me.splitMain.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.splitMain.Size = New System.Drawing.Size(907, 365)
        Me.splitMain.SplitterDistance = 267
        Me.splitMain.TabIndex = 1
        '
        'wbFont
        '
        Me.wbFont.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbFont.Location = New System.Drawing.Point(0, 0)
        Me.wbFont.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbFont.Name = "wbFont"
        Me.wbFont.Size = New System.Drawing.Size(907, 267)
        Me.wbFont.TabIndex = 0
        '
        'chkUnderline
        '
        Me.chkUnderline.AutoSize = True
        Me.chkUnderline.Location = New System.Drawing.Point(394, 58)
        Me.chkUnderline.Name = "chkUnderline"
        Me.chkUnderline.Size = New System.Drawing.Size(133, 24)
        Me.chkUnderline.TabIndex = 8
        Me.chkUnderline.Text = "Подчеркнутый"
        Me.chkUnderline.UseVisualStyleBackColor = True
        '
        'chkBold
        '
        Me.chkBold.AutoSize = True
        Me.chkBold.Location = New System.Drawing.Point(394, 36)
        Me.chkBold.Name = "chkBold"
        Me.chkBold.Size = New System.Drawing.Size(90, 24)
        Me.chkBold.TabIndex = 7
        Me.chkBold.Text = "Жирный"
        Me.chkBold.UseVisualStyleBackColor = True
        '
        'chkItalic
        '
        Me.chkItalic.AutoSize = True
        Me.chkItalic.Location = New System.Drawing.Point(394, 13)
        Me.chkItalic.Name = "chkItalic"
        Me.chkItalic.Size = New System.Drawing.Size(78, 24)
        Me.chkItalic.TabIndex = 6
        Me.chkItalic.Text = "Курсив"
        Me.chkItalic.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 20)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Шрифт"
        '
        'cmbFontName
        '
        Me.cmbFontName.FormattingEnabled = True
        Me.cmbFontName.Items.AddRange(New Object() {"Arial, Helvetica, sans-serif", "Arial Black, Gadget, sans-serif", "Comic Sans MS, cursive, sans-serif", "Courier New, Courier, monospace", "Georgia, serif", "Impact, Charcoal, sans-serif", "Lucida Console, Monaco, monospace", "Lucida Sans Unicode, Lucida Grande, sans-serif", "Tahoma, Geneva, sans-serif", "Times New Roman, Times, serif", "Trebuchet MS, Helvetica, sans-serif", "Verdana, Geneva, sans-serif"})
        Me.cmbFontName.Location = New System.Drawing.Point(80, 14)
        Me.cmbFontName.Name = "cmbFontName"
        Me.cmbFontName.Size = New System.Drawing.Size(292, 28)
        Me.cmbFontName.TabIndex = 4
        '
        'cmbPoints
        '
        Me.cmbPoints.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPoints.FormattingEnabled = True
        Me.cmbPoints.Items.AddRange(New Object() {"px", "em", "rem", "vw", "vh", "vmin", "vmax"})
        Me.cmbPoints.Location = New System.Drawing.Point(267, 54)
        Me.cmbPoints.Name = "cmbPoints"
        Me.cmbPoints.Size = New System.Drawing.Size(105, 28)
        Me.cmbPoints.TabIndex = 3
        '
        'cmbSize
        '
        Me.cmbSize.FormattingEnabled = True
        Me.cmbSize.Items.AddRange(New Object() {"2", "4", "6", "8", "9", "10", "11", "12", "13", "14", "15", "16", "18", "20", "22", "24", "26", "30", "35", "40", "45", "50", "60", "70"})
        Me.cmbSize.Location = New System.Drawing.Point(140, 54)
        Me.cmbSize.Name = "cmbSize"
        Me.cmbSize.Size = New System.Drawing.Size(121, 28)
        Me.cmbSize.TabIndex = 2
        Me.cmbSize.Text = "16"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(121, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Размер шрифта"
        '
        'dlgFontSize
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(907, 365)
        Me.Controls.Add(Me.splitMain)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgFontSize"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Подбор шрифта"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel2.ResumeLayout(False)
        Me.splitMain.Panel2.PerformLayout()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents wbFont As System.Windows.Forms.WebBrowser
    Friend WithEvents cmbPoints As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSize As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbFontName As System.Windows.Forms.ComboBox
    Friend WithEvents chkUnderline As System.Windows.Forms.CheckBox
    Friend WithEvents chkBold As System.Windows.Forms.CheckBox
    Friend WithEvents chkItalic As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
