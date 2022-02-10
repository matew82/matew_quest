<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgImage
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
        Me.splitInner = New System.Windows.Forms.SplitContainer()
        Me.treePath = New System.Windows.Forms.TreeView()
        Me.wbImages = New System.Windows.Forms.WebBrowser()
        Me.optFloatNo = New System.Windows.Forms.RadioButton()
        Me.optFloatRight = New System.Windows.Forms.RadioButton()
        Me.optFloatLeft = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtHeight = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtWidth = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        CType(Me.splitInner, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitInner.Panel1.SuspendLayout()
        Me.splitInner.Panel2.SuspendLayout()
        Me.splitInner.SuspendLayout()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(1033, 12)
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
        Me.splitMain.Panel1.Controls.Add(Me.splitInner)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.optFloatNo)
        Me.splitMain.Panel2.Controls.Add(Me.optFloatRight)
        Me.splitMain.Panel2.Controls.Add(Me.optFloatLeft)
        Me.splitMain.Panel2.Controls.Add(Me.Label2)
        Me.splitMain.Panel2.Controls.Add(Me.txtHeight)
        Me.splitMain.Panel2.Controls.Add(Me.Label1)
        Me.splitMain.Panel2.Controls.Add(Me.txtWidth)
        Me.splitMain.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.splitMain.Size = New System.Drawing.Size(1256, 580)
        Me.splitMain.SplitterDistance = 514
        Me.splitMain.TabIndex = 1
        '
        'splitInner
        '
        Me.splitInner.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitInner.Location = New System.Drawing.Point(0, 0)
        Me.splitInner.Name = "splitInner"
        '
        'splitInner.Panel1
        '
        Me.splitInner.Panel1.Controls.Add(Me.treePath)
        '
        'splitInner.Panel2
        '
        Me.splitInner.Panel2.Controls.Add(Me.wbImages)
        Me.splitInner.Size = New System.Drawing.Size(1256, 514)
        Me.splitInner.SplitterDistance = 356
        Me.splitInner.TabIndex = 0
        '
        'treePath
        '
        Me.treePath.AllowDrop = True
        Me.treePath.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treePath.FullRowSelect = True
        Me.treePath.HotTracking = True
        Me.treePath.Location = New System.Drawing.Point(0, 0)
        Me.treePath.Name = "treePath"
        Me.treePath.Size = New System.Drawing.Size(356, 514)
        Me.treePath.TabIndex = 0
        '
        'wbImages
        '
        Me.wbImages.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbImages.Location = New System.Drawing.Point(0, 0)
        Me.wbImages.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbImages.Name = "wbImages"
        Me.wbImages.Size = New System.Drawing.Size(896, 514)
        Me.wbImages.TabIndex = 0
        '
        'optFloatNo
        '
        Me.optFloatNo.AutoSize = True
        Me.optFloatNo.Checked = True
        Me.optFloatNo.Location = New System.Drawing.Point(406, 18)
        Me.optFloatNo.Name = "optFloatNo"
        Me.optFloatNo.Size = New System.Drawing.Size(137, 24)
        Me.optFloatNo.TabIndex = 7
        Me.optFloatNo.TabStop = True
        Me.optFloatNo.Text = "Без прилегания"
        Me.optFloatNo.UseVisualStyleBackColor = True
        '
        'optFloatRight
        '
        Me.optFloatRight.AutoSize = True
        Me.optFloatRight.Location = New System.Drawing.Point(709, 18)
        Me.optFloatRight.Name = "optFloatRight"
        Me.optFloatRight.Size = New System.Drawing.Size(164, 24)
        Me.optFloatRight.TabIndex = 6
        Me.optFloatRight.Tag = ""
        Me.optFloatRight.Text = "Прилегание вправо"
        Me.optFloatRight.UseVisualStyleBackColor = True
        '
        'optFloatLeft
        '
        Me.optFloatLeft.AutoSize = True
        Me.optFloatLeft.Location = New System.Drawing.Point(549, 18)
        Me.optFloatLeft.Name = "optFloatLeft"
        Me.optFloatLeft.Size = New System.Drawing.Size(154, 24)
        Me.optFloatLeft.TabIndex = 5
        Me.optFloatLeft.Tag = ""
        Me.optFloatLeft.Text = "Прилегание влево"
        Me.optFloatLeft.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(203, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 20)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Высота"
        '
        'txtHeight
        '
        Me.txtHeight.Location = New System.Drawing.Point(277, 17)
        Me.txtHeight.Name = "txtHeight"
        Me.txtHeight.Size = New System.Drawing.Size(100, 28)
        Me.txtHeight.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Ширина"
        '
        'txtWidth
        '
        Me.txtWidth.Location = New System.Drawing.Point(86, 17)
        Me.txtWidth.Name = "txtWidth"
        Me.txtWidth.Size = New System.Drawing.Size(100, 28)
        Me.txtWidth.TabIndex = 1
        '
        'dlgImage
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(1256, 580)
        Me.Controls.Add(Me.splitMain)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgImage"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Вставка изображения"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel2.ResumeLayout(False)
        Me.splitMain.Panel2.PerformLayout()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.splitInner.Panel1.ResumeLayout(False)
        Me.splitInner.Panel2.ResumeLayout(False)
        CType(Me.splitInner, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitInner.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents splitInner As System.Windows.Forms.SplitContainer
    Friend WithEvents treePath As System.Windows.Forms.TreeView
    Friend WithEvents wbImages As System.Windows.Forms.WebBrowser
    Friend WithEvents optFloatNo As System.Windows.Forms.RadioButton
    Friend WithEvents optFloatRight As System.Windows.Forms.RadioButton
    Friend WithEvents optFloatLeft As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtHeight As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtWidth As System.Windows.Forms.TextBox

End Class
