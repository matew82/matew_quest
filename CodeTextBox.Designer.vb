<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CodeTextBox
    Inherits System.Windows.Forms.UserControl

    'Пользовательский элемент управления (UserControl) переопределяет метод Dispose для очистки списка компонентов.
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
        Me.SplitContainerHorizontal = New System.Windows.Forms.SplitContainer()
        Me.SplitContainerVertical = New System.Windows.Forms.SplitContainer()
        Me.lstRtb = New System.Windows.Forms.ListBox()
        Me.wbRtbHelp = New System.Windows.Forms.WebBrowser()
        Me.mnuRtb = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuUndo = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRedo = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuSelectAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.timDraw = New System.Windows.Forms.Timer(Me.components)
        CType(Me.SplitContainerHorizontal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerHorizontal.Panel1.SuspendLayout()
        Me.SplitContainerHorizontal.Panel2.SuspendLayout()
        Me.SplitContainerHorizontal.SuspendLayout()
        CType(Me.SplitContainerVertical, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerVertical.SuspendLayout()
        Me.mnuRtb.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainerHorizontal
        '
        Me.SplitContainerHorizontal.CausesValidation = False
        Me.SplitContainerHorizontal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainerHorizontal.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainerHorizontal.Name = "SplitContainerHorizontal"
        Me.SplitContainerHorizontal.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainerHorizontal.Panel1
        '
        Me.SplitContainerHorizontal.Panel1.Controls.Add(Me.SplitContainerVertical)
        '
        'SplitContainerHorizontal.Panel2
        '
        Me.SplitContainerHorizontal.Panel2.Controls.Add(Me.lstRtb)
        Me.SplitContainerHorizontal.Panel2.Controls.Add(Me.wbRtbHelp)
        Me.SplitContainerHorizontal.Size = New System.Drawing.Size(441, 360)
        Me.SplitContainerHorizontal.SplitterDistance = 281
        Me.SplitContainerHorizontal.TabIndex = 1
        '
        'SplitContainerVertical
        '
        Me.SplitContainerVertical.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainerVertical.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainerVertical.Name = "SplitContainerVertical"
        '
        'SplitContainerVertical.Panel1
        '
        Me.SplitContainerVertical.Panel1.CausesValidation = False
        Me.SplitContainerVertical.Panel1.Font = New System.Drawing.Font("Consolas", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        '
        'SplitContainerVertical.Panel2
        '
        Me.SplitContainerVertical.Panel2.BackColor = System.Drawing.SystemColors.Control
        Me.SplitContainerVertical.Size = New System.Drawing.Size(441, 281)
        Me.SplitContainerVertical.SplitterDistance = 25
        Me.SplitContainerVertical.TabIndex = 0
        '
        'lstRtb
        '
        Me.lstRtb.Font = New System.Drawing.Font("Consolas", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lstRtb.FormattingEnabled = True
        Me.lstRtb.ItemHeight = 18
        Me.lstRtb.Items.AddRange(New Object() {"мропо", "одлолдд"})
        Me.lstRtb.Location = New System.Drawing.Point(148, 3)
        Me.lstRtb.Name = "lstRtb"
        Me.lstRtb.Size = New System.Drawing.Size(197, 76)
        Me.lstRtb.Sorted = True
        Me.lstRtb.TabIndex = 10
        Me.lstRtb.Visible = False
        '
        'wbRtbHelp
        '
        Me.wbRtbHelp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbRtbHelp.Location = New System.Drawing.Point(0, 0)
        Me.wbRtbHelp.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbRtbHelp.Name = "wbRtbHelp"
        Me.wbRtbHelp.Size = New System.Drawing.Size(441, 75)
        Me.wbRtbHelp.TabIndex = 9
        '
        'mnuRtb
        '
        Me.mnuRtb.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCopy, Me.mnuCut, Me.mnuPaste, Me.ToolStripSeparator1, Me.mnuUndo, Me.mnuRedo, Me.ToolStripSeparator2, Me.mnuSelectAll})
        Me.mnuRtb.Name = "mnuRtb"
        Me.mnuRtb.Size = New System.Drawing.Size(165, 148)
        Me.mnuRtb.Text = "Menu"
        '
        'mnuCopy
        '
        Me.mnuCopy.Name = "mnuCopy"
        Me.mnuCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.mnuCopy.Size = New System.Drawing.Size(164, 22)
        Me.mnuCopy.Text = "Copy"
        '
        'mnuCut
        '
        Me.mnuCut.Name = "mnuCut"
        Me.mnuCut.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.mnuCut.Size = New System.Drawing.Size(164, 22)
        Me.mnuCut.Text = "Cut"
        '
        'mnuPaste
        '
        Me.mnuPaste.Name = "mnuPaste"
        Me.mnuPaste.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.mnuPaste.Size = New System.Drawing.Size(164, 22)
        Me.mnuPaste.Text = "Paste"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(161, 6)
        '
        'mnuUndo
        '
        Me.mnuUndo.Name = "mnuUndo"
        Me.mnuUndo.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Z), System.Windows.Forms.Keys)
        Me.mnuUndo.Size = New System.Drawing.Size(164, 22)
        Me.mnuUndo.Text = "Undo"
        '
        'mnuRedo
        '
        Me.mnuRedo.Name = "mnuRedo"
        Me.mnuRedo.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.mnuRedo.Size = New System.Drawing.Size(164, 22)
        Me.mnuRedo.Text = "Redo"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(161, 6)
        '
        'mnuSelectAll
        '
        Me.mnuSelectAll.Name = "mnuSelectAll"
        Me.mnuSelectAll.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.mnuSelectAll.Size = New System.Drawing.Size(164, 22)
        Me.mnuSelectAll.Text = "Select All"
        '
        'timDraw
        '
        Me.timDraw.Interval = 200
        '
        'CodeTextBox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.SplitContainerHorizontal)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "CodeTextBox"
        Me.Size = New System.Drawing.Size(441, 360)
        Me.SplitContainerHorizontal.Panel1.ResumeLayout(False)
        Me.SplitContainerHorizontal.Panel2.ResumeLayout(False)
        CType(Me.SplitContainerHorizontal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerHorizontal.ResumeLayout(False)
        CType(Me.SplitContainerVertical, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerVertical.ResumeLayout(False)
        Me.mnuRtb.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainerHorizontal As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainerVertical As System.Windows.Forms.SplitContainer
    Friend WithEvents lstRtb As System.Windows.Forms.ListBox
    Friend WithEvents mnuRtb As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCut As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuPaste As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuUndo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRedo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuSelectAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents wbRtbHelp As System.Windows.Forms.WebBrowser
    Friend WithEvents timDraw As System.Windows.Forms.Timer

End Class
