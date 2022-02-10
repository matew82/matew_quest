<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgListHtml
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
        Me.wbList = New System.Windows.Forms.WebBrowser()
        Me.txtStart = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblImage = New System.Windows.Forms.Label()
        Me.cmbImage = New System.Windows.Forms.ComboBox()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(912, 5)
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
        Me.splitMain.Panel1.Controls.Add(Me.wbList)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.txtStart)
        Me.splitMain.Panel2.Controls.Add(Me.Label1)
        Me.splitMain.Panel2.Controls.Add(Me.lblImage)
        Me.splitMain.Panel2.Controls.Add(Me.cmbImage)
        Me.splitMain.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.splitMain.Size = New System.Drawing.Size(1135, 699)
        Me.splitMain.SplitterDistance = 640
        Me.splitMain.TabIndex = 1
        '
        'wbList
        '
        Me.wbList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbList.Location = New System.Drawing.Point(0, 0)
        Me.wbList.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbList.Name = "wbList"
        Me.wbList.Size = New System.Drawing.Size(1135, 640)
        Me.wbList.TabIndex = 0
        '
        'txtStart
        '
        Me.txtStart.Location = New System.Drawing.Point(672, 18)
        Me.txtStart.Name = "txtStart"
        Me.txtStart.Size = New System.Drawing.Size(100, 28)
        Me.txtStart.TabIndex = 4
        Me.txtStart.Text = "1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(595, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 20)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Начать с"
        '
        'lblImage
        '
        Me.lblImage.AutoSize = True
        Me.lblImage.Location = New System.Drawing.Point(12, 20)
        Me.lblImage.Name = "lblImage"
        Me.lblImage.Size = New System.Drawing.Size(157, 20)
        Me.lblImage.TabIndex = 2
        Me.lblImage.Text = "Путь к изображению"
        '
        'cmbImage
        '
        Me.cmbImage.FormattingEnabled = True
        Me.cmbImage.Location = New System.Drawing.Point(175, 18)
        Me.cmbImage.Name = "cmbImage"
        Me.cmbImage.Size = New System.Drawing.Size(405, 28)
        Me.cmbImage.TabIndex = 1
        Me.cmbImage.Text = "img/Default/listItem.png"
        '
        'dlgListHtml
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(1135, 699)
        Me.Controls.Add(Me.splitMain)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgListHtml"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Создание списка"
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
    Friend WithEvents wbList As System.Windows.Forms.WebBrowser
    Friend WithEvents cmbImage As System.Windows.Forms.ComboBox
    Friend WithEvents lblImage As System.Windows.Forms.Label
    Friend WithEvents txtStart As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
