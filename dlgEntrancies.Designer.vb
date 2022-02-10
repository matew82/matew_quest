<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgEntrancies
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
        Me.components = New System.ComponentModel.Container()
        Me.splitMain = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.lstBoxEntrancies = New System.Windows.Forms.ListBox()
        Me.lstBoxEntranciesChecked = New System.Windows.Forms.CheckedListBox()
        Me.cmnuListBoxChecked = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiCheckAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiUnCheckAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.chkReplace = New System.Windows.Forms.CheckBox()
        Me.btnReplace = New System.Windows.Forms.Button()
        Me.txtReplace = New System.Windows.Forms.TextBox()
        Me.btnGo = New System.Windows.Forms.Button()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.cmnuListBoxChecked.SuspendLayout()
        Me.SuspendLayout()
        '
        'splitMain
        '
        Me.splitMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitMain.IsSplitterFixed = True
        Me.splitMain.Location = New System.Drawing.Point(0, 0)
        Me.splitMain.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.splitMain.Name = "splitMain"
        Me.splitMain.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitMain.Panel1
        '
        Me.splitMain.Panel1.Controls.Add(Me.SplitContainer1)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.chkReplace)
        Me.splitMain.Panel2.Controls.Add(Me.btnReplace)
        Me.splitMain.Panel2.Controls.Add(Me.txtReplace)
        Me.splitMain.Panel2.Controls.Add(Me.btnGo)
        Me.splitMain.Size = New System.Drawing.Size(890, 514)
        Me.splitMain.SplitterDistance = 433
        Me.splitMain.SplitterWidth = 6
        Me.splitMain.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lblInfo)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lstBoxEntrancies)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lstBoxEntranciesChecked)
        Me.SplitContainer1.Size = New System.Drawing.Size(890, 433)
        Me.SplitContainer1.SplitterDistance = 64
        Me.SplitContainer1.TabIndex = 2
        '
        'lblInfo
        '
        Me.lblInfo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblInfo.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblInfo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblInfo.Location = New System.Drawing.Point(0, 0)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(890, 64)
        Me.lblInfo.TabIndex = 1
        Me.lblInfo.Text = "Label1"
        Me.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lstBoxEntrancies
        '
        Me.lstBoxEntrancies.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstBoxEntrancies.FormattingEnabled = True
        Me.lstBoxEntrancies.ItemHeight = 20
        Me.lstBoxEntrancies.Location = New System.Drawing.Point(0, 0)
        Me.lstBoxEntrancies.Name = "lstBoxEntrancies"
        Me.lstBoxEntrancies.Size = New System.Drawing.Size(890, 365)
        Me.lstBoxEntrancies.TabIndex = 0
        '
        'lstBoxEntranciesChecked
        '
        Me.lstBoxEntranciesChecked.ContextMenuStrip = Me.cmnuListBoxChecked
        Me.lstBoxEntranciesChecked.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstBoxEntranciesChecked.FormattingEnabled = True
        Me.lstBoxEntranciesChecked.Location = New System.Drawing.Point(0, 0)
        Me.lstBoxEntranciesChecked.Name = "lstBoxEntranciesChecked"
        Me.lstBoxEntranciesChecked.Size = New System.Drawing.Size(890, 365)
        Me.lstBoxEntranciesChecked.TabIndex = 1
        Me.lstBoxEntranciesChecked.ThreeDCheckBoxes = True
        '
        'cmnuListBoxChecked
        '
        Me.cmnuListBoxChecked.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiCheckAll, Me.tsmiUnCheckAll})
        Me.cmnuListBoxChecked.Name = "cmnuListBoxChecked"
        Me.cmnuListBoxChecked.Size = New System.Drawing.Size(170, 48)
        '
        'tsmiCheckAll
        '
        Me.tsmiCheckAll.Name = "tsmiCheckAll"
        Me.tsmiCheckAll.Size = New System.Drawing.Size(169, 22)
        Me.tsmiCheckAll.Text = "Выделить все"
        '
        'tsmiUnCheckAll
        '
        Me.tsmiUnCheckAll.Name = "tsmiUnCheckAll"
        Me.tsmiUnCheckAll.Size = New System.Drawing.Size(169, 22)
        Me.tsmiUnCheckAll.Text = "Снять выделение"
        '
        'chkReplace
        '
        Me.chkReplace.AutoSize = True
        Me.chkReplace.Location = New System.Drawing.Point(12, 21)
        Me.chkReplace.Name = "chkReplace"
        Me.chkReplace.Size = New System.Drawing.Size(120, 24)
        Me.chkReplace.TabIndex = 4
        Me.chkReplace.Text = "Заменить на:"
        Me.chkReplace.UseVisualStyleBackColor = True
        '
        'btnReplace
        '
        Me.btnReplace.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReplace.Enabled = False
        Me.btnReplace.Location = New System.Drawing.Point(376, 12)
        Me.btnReplace.Name = "btnReplace"
        Me.btnReplace.Size = New System.Drawing.Size(195, 37)
        Me.btnReplace.TabIndex = 3
        Me.btnReplace.Text = "Заменить"
        Me.btnReplace.UseVisualStyleBackColor = True
        '
        'txtReplace
        '
        Me.txtReplace.Enabled = False
        Me.txtReplace.Location = New System.Drawing.Point(138, 19)
        Me.txtReplace.Name = "txtReplace"
        Me.txtReplace.Size = New System.Drawing.Size(232, 28)
        Me.txtReplace.TabIndex = 1
        '
        'btnGo
        '
        Me.btnGo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGo.Location = New System.Drawing.Point(683, 12)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(195, 37)
        Me.btnGo.TabIndex = 0
        Me.btnGo.Text = "Перейти"
        Me.btnGo.UseVisualStyleBackColor = True
        '
        'dlgEntrancies
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(890, 514)
        Me.Controls.Add(Me.splitMain)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "dlgEntrancies"
        Me.Text = "Найденные вхождения"
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel2.ResumeLayout(False)
        Me.splitMain.Panel2.PerformLayout()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.cmnuListBoxChecked.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents btnGo As System.Windows.Forms.Button
    Friend WithEvents lstBoxEntrancies As System.Windows.Forms.ListBox
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents lstBoxEntranciesChecked As System.Windows.Forms.CheckedListBox
    Friend WithEvents cmnuListBoxChecked As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiCheckAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiUnCheckAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents chkReplace As System.Windows.Forms.CheckBox
    Friend WithEvents btnReplace As System.Windows.Forms.Button
    Friend WithEvents txtReplace As System.Windows.Forms.TextBox
End Class
