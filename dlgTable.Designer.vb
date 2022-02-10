<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgTable
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
        Me.splitSecond = New System.Windows.Forms.SplitContainer()
        Me.wbTable = New System.Windows.Forms.WebBrowser()
        Me.txtCaption = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.numRows = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.numColumns = New System.Windows.Forms.NumericUpDown()
        Me.txtWidth = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtResult = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        CType(Me.splitSecond, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitSecond.Panel1.SuspendLayout()
        Me.splitSecond.Panel2.SuspendLayout()
        Me.splitSecond.SuspendLayout()
        CType(Me.numRows, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numColumns, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(911, 8)
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
        Me.splitMain.Panel1.Controls.Add(Me.splitSecond)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.splitMain.Size = New System.Drawing.Size(1134, 589)
        Me.splitMain.SplitterDistance = 527
        Me.splitMain.TabIndex = 1
        '
        'splitSecond
        '
        Me.splitSecond.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitSecond.Location = New System.Drawing.Point(0, 0)
        Me.splitSecond.Name = "splitSecond"
        Me.splitSecond.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitSecond.Panel1
        '
        Me.splitSecond.Panel1.Controls.Add(Me.wbTable)
        Me.splitSecond.Panel1.Controls.Add(Me.txtCaption)
        Me.splitSecond.Panel1.Controls.Add(Me.Label4)
        Me.splitSecond.Panel1.Controls.Add(Me.Label1)
        Me.splitSecond.Panel1.Controls.Add(Me.numRows)
        Me.splitSecond.Panel1.Controls.Add(Me.Label3)
        Me.splitSecond.Panel1.Controls.Add(Me.numColumns)
        Me.splitSecond.Panel1.Controls.Add(Me.txtWidth)
        Me.splitSecond.Panel1.Controls.Add(Me.Label2)
        '
        'splitSecond.Panel2
        '
        Me.splitSecond.Panel2.Controls.Add(Me.txtResult)
        Me.splitSecond.Size = New System.Drawing.Size(1134, 527)
        Me.splitSecond.SplitterDistance = 276
        Me.splitSecond.TabIndex = 8
        '
        'wbTable
        '
        Me.wbTable.AllowWebBrowserDrop = False
        Me.wbTable.Location = New System.Drawing.Point(12, 52)
        Me.wbTable.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbTable.Name = "wbTable"
        Me.wbTable.Size = New System.Drawing.Size(1110, 221)
        Me.wbTable.TabIndex = 0
        Me.wbTable.WebBrowserShortcutsEnabled = False
        '
        'txtCaption
        '
        Me.txtCaption.Location = New System.Drawing.Point(699, 15)
        Me.txtCaption.Name = "txtCaption"
        Me.txtCaption.Size = New System.Drawing.Size(423, 28)
        Me.txtCaption.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(613, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 20)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Заголовок"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 20)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Рядов"
        '
        'numRows
        '
        Me.numRows.Location = New System.Drawing.Point(68, 16)
        Me.numRows.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numRows.Name = "numRows"
        Me.numRows.Size = New System.Drawing.Size(120, 28)
        Me.numRows.TabIndex = 1
        Me.numRows.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(413, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 20)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Ширина"
        '
        'numColumns
        '
        Me.numColumns.Location = New System.Drawing.Point(284, 16)
        Me.numColumns.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numColumns.Name = "numColumns"
        Me.numColumns.Size = New System.Drawing.Size(120, 28)
        Me.numColumns.TabIndex = 2
        Me.numColumns.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'txtWidth
        '
        Me.txtWidth.Location = New System.Drawing.Point(487, 15)
        Me.txtWidth.Name = "txtWidth"
        Me.txtWidth.Size = New System.Drawing.Size(120, 28)
        Me.txtWidth.TabIndex = 5
        Me.txtWidth.Text = "100%"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(209, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 20)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Колонок"
        '
        'txtResult
        '
        Me.txtResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtResult.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.txtResult.Location = New System.Drawing.Point(0, 0)
        Me.txtResult.Multiline = True
        Me.txtResult.Name = "txtResult"
        Me.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtResult.Size = New System.Drawing.Size(1134, 247)
        Me.txtResult.TabIndex = 7
        '
        'dlgTable
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(1134, 589)
        Me.Controls.Add(Me.splitMain)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgTable"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Создание таблицы"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel2.ResumeLayout(False)
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.splitSecond.Panel1.ResumeLayout(False)
        Me.splitSecond.Panel1.PerformLayout()
        Me.splitSecond.Panel2.ResumeLayout(False)
        Me.splitSecond.Panel2.PerformLayout()
        CType(Me.splitSecond, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitSecond.ResumeLayout(False)
        CType(Me.numRows, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numColumns, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents wbTable As System.Windows.Forms.WebBrowser
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtWidth As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents numColumns As System.Windows.Forms.NumericUpDown
    Friend WithEvents numRows As System.Windows.Forms.NumericUpDown
    Friend WithEvents splitSecond As System.Windows.Forms.SplitContainer
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    Friend WithEvents txtCaption As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label

End Class
