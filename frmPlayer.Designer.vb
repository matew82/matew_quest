<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPlayer
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
        Me.mnuMain = New System.Windows.Forms.MenuStrip()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИграToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiNewGame = New System.Windows.Forms.ToolStripMenuItem()
        Me.СохранитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЗагрузитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.НачатьЗановоToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.splitHorizontal = New System.Windows.Forms.SplitContainer()
        Me.split12 = New System.Windows.Forms.SplitContainer()
        Me.wbMain = New System.Windows.Forms.WebBrowser()
        Me.wbObjects = New System.Windows.Forms.WebBrowser()
        Me.split345 = New System.Windows.Forms.SplitContainer()
        Me.split35 = New System.Windows.Forms.SplitContainer()
        Me.wbActions = New System.Windows.Forms.WebBrowser()
        Me.wbCommand = New System.Windows.Forms.WebBrowser()
        Me.wbDescription = New System.Windows.Forms.WebBrowser()
        Me.mnuMain.SuspendLayout()
        CType(Me.splitHorizontal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitHorizontal.Panel1.SuspendLayout()
        Me.splitHorizontal.Panel2.SuspendLayout()
        Me.splitHorizontal.SuspendLayout()
        CType(Me.split12, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.split12.Panel1.SuspendLayout()
        Me.split12.Panel2.SuspendLayout()
        Me.split12.SuspendLayout()
        CType(Me.split345, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.split345.Panel1.SuspendLayout()
        Me.split345.Panel2.SuspendLayout()
        Me.split345.SuspendLayout()
        CType(Me.split35, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.split35.Panel1.SuspendLayout()
        Me.split35.Panel2.SuspendLayout()
        Me.split35.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuMain
        '
        Me.mnuMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.ИграToolStripMenuItem})
        Me.mnuMain.Location = New System.Drawing.Point(0, 0)
        Me.mnuMain.Name = "mnuMain"
        Me.mnuMain.Size = New System.Drawing.Size(996, 24)
        Me.mnuMain.TabIndex = 0
        Me.mnuMain.Text = "MenuStrip1"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(12, 20)
        '
        'ИграToolStripMenuItem
        '
        Me.ИграToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiNewGame, Me.СохранитьToolStripMenuItem, Me.ЗагрузитьToolStripMenuItem, Me.НачатьЗановоToolStripMenuItem})
        Me.ИграToolStripMenuItem.Name = "ИграToolStripMenuItem"
        Me.ИграToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
        Me.ИграToolStripMenuItem.Text = "Игра"
        '
        'tsmiNewGame
        '
        Me.tsmiNewGame.Name = "tsmiNewGame"
        Me.tsmiNewGame.Size = New System.Drawing.Size(154, 22)
        Me.tsmiNewGame.Text = "Новая"
        '
        'СохранитьToolStripMenuItem
        '
        Me.СохранитьToolStripMenuItem.Name = "СохранитьToolStripMenuItem"
        Me.СохранитьToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.СохранитьToolStripMenuItem.Text = "Сохранить"
        '
        'ЗагрузитьToolStripMenuItem
        '
        Me.ЗагрузитьToolStripMenuItem.Name = "ЗагрузитьToolStripMenuItem"
        Me.ЗагрузитьToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.ЗагрузитьToolStripMenuItem.Text = "Загрузить"
        '
        'НачатьЗановоToolStripMenuItem
        '
        Me.НачатьЗановоToolStripMenuItem.Name = "НачатьЗановоToolStripMenuItem"
        Me.НачатьЗановоToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.НачатьЗановоToolStripMenuItem.Text = "Начать заново"
        '
        'splitHorizontal
        '
        Me.splitHorizontal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitHorizontal.Location = New System.Drawing.Point(0, 24)
        Me.splitHorizontal.Name = "splitHorizontal"
        Me.splitHorizontal.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitHorizontal.Panel1
        '
        Me.splitHorizontal.Panel1.Controls.Add(Me.split12)
        '
        'splitHorizontal.Panel2
        '
        Me.splitHorizontal.Panel2.Controls.Add(Me.split345)
        Me.splitHorizontal.Size = New System.Drawing.Size(996, 572)
        Me.splitHorizontal.SplitterDistance = 339
        Me.splitHorizontal.TabIndex = 1
        '
        'split12
        '
        Me.split12.Dock = System.Windows.Forms.DockStyle.Fill
        Me.split12.Location = New System.Drawing.Point(0, 0)
        Me.split12.Name = "split12"
        '
        'split12.Panel1
        '
        Me.split12.Panel1.Controls.Add(Me.wbMain)
        '
        'split12.Panel2
        '
        Me.split12.Panel2.Controls.Add(Me.wbObjects)
        Me.split12.Size = New System.Drawing.Size(996, 339)
        Me.split12.SplitterDistance = 549
        Me.split12.TabIndex = 0
        '
        'wbMain
        '
        Me.wbMain.AllowNavigation = False
        Me.wbMain.AllowWebBrowserDrop = False
        Me.wbMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbMain.IsWebBrowserContextMenuEnabled = False
        Me.wbMain.Location = New System.Drawing.Point(0, 0)
        Me.wbMain.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbMain.Name = "wbMain"
        Me.wbMain.Size = New System.Drawing.Size(549, 339)
        Me.wbMain.TabIndex = 0
        '
        'wbObjects
        '
        Me.wbObjects.AllowNavigation = False
        Me.wbObjects.AllowWebBrowserDrop = False
        Me.wbObjects.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbObjects.IsWebBrowserContextMenuEnabled = False
        Me.wbObjects.Location = New System.Drawing.Point(0, 0)
        Me.wbObjects.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbObjects.Name = "wbObjects"
        Me.wbObjects.Size = New System.Drawing.Size(443, 339)
        Me.wbObjects.TabIndex = 1
        '
        'split345
        '
        Me.split345.Dock = System.Windows.Forms.DockStyle.Fill
        Me.split345.Location = New System.Drawing.Point(0, 0)
        Me.split345.Name = "split345"
        '
        'split345.Panel1
        '
        Me.split345.Panel1.Controls.Add(Me.split35)
        '
        'split345.Panel2
        '
        Me.split345.Panel2.Controls.Add(Me.wbDescription)
        Me.split345.Size = New System.Drawing.Size(996, 229)
        Me.split345.SplitterDistance = 549
        Me.split345.TabIndex = 0
        '
        'split35
        '
        Me.split35.Dock = System.Windows.Forms.DockStyle.Fill
        Me.split35.Location = New System.Drawing.Point(0, 0)
        Me.split35.Name = "split35"
        Me.split35.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'split35.Panel1
        '
        Me.split35.Panel1.Controls.Add(Me.wbActions)
        '
        'split35.Panel2
        '
        Me.split35.Panel2.Controls.Add(Me.wbCommand)
        Me.split35.Size = New System.Drawing.Size(549, 229)
        Me.split35.SplitterDistance = 177
        Me.split35.TabIndex = 0
        '
        'wbActions
        '
        Me.wbActions.AllowNavigation = False
        Me.wbActions.AllowWebBrowserDrop = False
        Me.wbActions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbActions.IsWebBrowserContextMenuEnabled = False
        Me.wbActions.Location = New System.Drawing.Point(0, 0)
        Me.wbActions.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbActions.Name = "wbActions"
        Me.wbActions.Size = New System.Drawing.Size(549, 177)
        Me.wbActions.TabIndex = 2
        '
        'wbCommand
        '
        Me.wbCommand.AllowNavigation = False
        Me.wbCommand.AllowWebBrowserDrop = False
        Me.wbCommand.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbCommand.IsWebBrowserContextMenuEnabled = False
        Me.wbCommand.Location = New System.Drawing.Point(0, 0)
        Me.wbCommand.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbCommand.Name = "wbCommand"
        Me.wbCommand.Size = New System.Drawing.Size(549, 48)
        Me.wbCommand.TabIndex = 4
        '
        'wbDescription
        '
        Me.wbDescription.AllowNavigation = False
        Me.wbDescription.AllowWebBrowserDrop = False
        Me.wbDescription.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbDescription.IsWebBrowserContextMenuEnabled = False
        Me.wbDescription.Location = New System.Drawing.Point(0, 0)
        Me.wbDescription.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbDescription.Name = "wbDescription"
        Me.wbDescription.Size = New System.Drawing.Size(443, 229)
        Me.wbDescription.TabIndex = 3
        '
        'frmPlayer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(996, 596)
        Me.Controls.Add(Me.splitHorizontal)
        Me.Controls.Add(Me.mnuMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MainMenuStrip = Me.mnuMain
        Me.MaximizeBox = False
        Me.Name = "frmPlayer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmPlayer"
        Me.mnuMain.ResumeLayout(False)
        Me.mnuMain.PerformLayout()
        Me.splitHorizontal.Panel1.ResumeLayout(False)
        Me.splitHorizontal.Panel2.ResumeLayout(False)
        CType(Me.splitHorizontal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitHorizontal.ResumeLayout(False)
        Me.split12.Panel1.ResumeLayout(False)
        Me.split12.Panel2.ResumeLayout(False)
        CType(Me.split12, System.ComponentModel.ISupportInitialize).EndInit()
        Me.split12.ResumeLayout(False)
        Me.split345.Panel1.ResumeLayout(False)
        Me.split345.Panel2.ResumeLayout(False)
        CType(Me.split345, System.ComponentModel.ISupportInitialize).EndInit()
        Me.split345.ResumeLayout(False)
        Me.split35.Panel1.ResumeLayout(False)
        Me.split35.Panel2.ResumeLayout(False)
        CType(Me.split35, System.ComponentModel.ISupportInitialize).EndInit()
        Me.split35.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents mnuMain As System.Windows.Forms.MenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ИграToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiNewGame As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents СохранитьToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ЗагрузитьToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents НачатьЗановоToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents splitHorizontal As System.Windows.Forms.SplitContainer
    Friend WithEvents split12 As System.Windows.Forms.SplitContainer
    Friend WithEvents wbMain As System.Windows.Forms.WebBrowser
    Friend WithEvents wbObjects As System.Windows.Forms.WebBrowser
    Friend WithEvents split345 As System.Windows.Forms.SplitContainer
    Friend WithEvents split35 As System.Windows.Forms.SplitContainer
    Friend WithEvents wbActions As System.Windows.Forms.WebBrowser
    Friend WithEvents wbCommand As System.Windows.Forms.WebBrowser
    Friend WithEvents wbDescription As System.Windows.Forms.WebBrowser
End Class
