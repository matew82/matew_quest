<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgSetChildrenValue
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
        Me.treeChildren = New System.Windows.Forms.TreeView()
        Me.splitMain = New System.Windows.Forms.SplitContainer()
        Me.lblValueToSet = New System.Windows.Forms.Label()
        Me.splitInner = New System.Windows.Forms.SplitContainer()
        Me.btnCheckAll = New System.Windows.Forms.Button()
        Me.btnUnselectAll = New System.Windows.Forms.Button()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(566, 11)
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
        'treeChildren
        '
        Me.treeChildren.CheckBoxes = True
        Me.treeChildren.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeChildren.FullRowSelect = True
        Me.treeChildren.HotTracking = True
        Me.treeChildren.Location = New System.Drawing.Point(0, 0)
        Me.treeChildren.Name = "treeChildren"
        Me.treeChildren.Size = New System.Drawing.Size(789, 436)
        Me.treeChildren.TabIndex = 1
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
        Me.splitMain.Panel2.Controls.Add(Me.lblValueToSet)
        Me.splitMain.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.splitMain.Panel2MinSize = 60
        Me.splitMain.Size = New System.Drawing.Size(789, 555)
        Me.splitMain.SplitterDistance = 490
        Me.splitMain.TabIndex = 2
        '
        'lblValueToSet
        '
        Me.lblValueToSet.AutoSize = True
        Me.lblValueToSet.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblValueToSet.Location = New System.Drawing.Point(12, 23)
        Me.lblValueToSet.Name = "lblValueToSet"
        Me.lblValueToSet.Size = New System.Drawing.Size(17, 20)
        Me.lblValueToSet.TabIndex = 2
        Me.lblValueToSet.Text = "я"
        '
        'splitInner
        '
        Me.splitInner.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitInner.Location = New System.Drawing.Point(0, 0)
        Me.splitInner.Name = "splitInner"
        Me.splitInner.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitInner.Panel1
        '
        Me.splitInner.Panel1.Controls.Add(Me.btnUnselectAll)
        Me.splitInner.Panel1.Controls.Add(Me.btnCheckAll)
        '
        'splitInner.Panel2
        '
        Me.splitInner.Panel2.Controls.Add(Me.treeChildren)
        Me.splitInner.Size = New System.Drawing.Size(789, 490)
        Me.splitInner.TabIndex = 2
        '
        'btnCheckAll
        '
        Me.btnCheckAll.Location = New System.Drawing.Point(3, 12)
        Me.btnCheckAll.Name = "btnCheckAll"
        Me.btnCheckAll.Size = New System.Drawing.Size(159, 33)
        Me.btnCheckAll.TabIndex = 0
        Me.btnCheckAll.Text = "Выделить все"
        Me.btnCheckAll.UseVisualStyleBackColor = True
        '
        'btnUnselectAll
        '
        Me.btnUnselectAll.Location = New System.Drawing.Point(168, 12)
        Me.btnUnselectAll.Name = "btnUnselectAll"
        Me.btnUnselectAll.Size = New System.Drawing.Size(159, 33)
        Me.btnUnselectAll.TabIndex = 1
        Me.btnUnselectAll.Text = "Снять выделение"
        Me.btnUnselectAll.UseVisualStyleBackColor = True
        '
        'dlgSetChildrenValue
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(789, 555)
        Me.Controls.Add(Me.splitMain)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgSetChildrenValue"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "dlgSetChildrenValue"
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
    Friend WithEvents treeChildren As System.Windows.Forms.TreeView
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents lblValueToSet As System.Windows.Forms.Label
    Friend WithEvents splitInner As System.Windows.Forms.SplitContainer
    Friend WithEvents btnCheckAll As System.Windows.Forms.Button
    Friend WithEvents btnUnselectAll As System.Windows.Forms.Button

End Class
