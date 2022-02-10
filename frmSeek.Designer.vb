<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSeek
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
        Me.chkCase = New System.Windows.Forms.CheckBox()
        Me.chkWholeWord = New System.Windows.Forms.CheckBox()
        Me.treeElements = New System.Windows.Forms.TreeView()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.chkSearchElement = New System.Windows.Forms.CheckBox()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.splitMain = New System.Windows.Forms.SplitContainer()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkCase
        '
        Me.chkCase.AutoSize = True
        Me.chkCase.Location = New System.Drawing.Point(17, 47)
        Me.chkCase.Name = "chkCase"
        Me.chkCase.Size = New System.Drawing.Size(159, 24)
        Me.chkCase.TabIndex = 0
        Me.chkCase.Text = "С учетом регистра"
        Me.chkCase.UseVisualStyleBackColor = True
        '
        'chkWholeWord
        '
        Me.chkWholeWord.AutoSize = True
        Me.chkWholeWord.Location = New System.Drawing.Point(228, 47)
        Me.chkWholeWord.Name = "chkWholeWord"
        Me.chkWholeWord.Size = New System.Drawing.Size(135, 24)
        Me.chkWholeWord.TabIndex = 1
        Me.chkWholeWord.Text = "Слово целиком"
        Me.chkWholeWord.UseVisualStyleBackColor = True
        '
        'treeElements
        '
        Me.treeElements.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeElements.Location = New System.Drawing.Point(0, 0)
        Me.treeElements.Name = "treeElements"
        Me.treeElements.Size = New System.Drawing.Size(707, 505)
        Me.treeElements.TabIndex = 3
        '
        'txtSearch
        '
        Me.txtSearch.Location = New System.Drawing.Point(17, 13)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(582, 28)
        Me.txtSearch.TabIndex = 4
        '
        'chkSearchElement
        '
        Me.chkSearchElement.AutoSize = True
        Me.chkSearchElement.Location = New System.Drawing.Point(414, 47)
        Me.chkSearchElement.Name = "chkSearchElement"
        Me.chkSearchElement.Size = New System.Drawing.Size(154, 24)
        Me.chkSearchElement.TabIndex = 5
        Me.chkSearchElement.Text = "Поиск элемента..."
        Me.chkSearchElement.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFind.Image = Global.WindowsApplication1.My.Resources.Resources.find
        Me.btnFind.Location = New System.Drawing.Point(607, 12)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(88, 59)
        Me.btnFind.TabIndex = 6
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'splitMain
        '
        Me.splitMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitMain.IsSplitterFixed = True
        Me.splitMain.Location = New System.Drawing.Point(0, 0)
        Me.splitMain.Name = "splitMain"
        Me.splitMain.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitMain.Panel1
        '
        Me.splitMain.Panel1.Controls.Add(Me.txtSearch)
        Me.splitMain.Panel1.Controls.Add(Me.btnFind)
        Me.splitMain.Panel1.Controls.Add(Me.chkWholeWord)
        Me.splitMain.Panel1.Controls.Add(Me.chkSearchElement)
        Me.splitMain.Panel1.Controls.Add(Me.chkCase)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.treeElements)
        Me.splitMain.Size = New System.Drawing.Size(707, 586)
        Me.splitMain.SplitterDistance = 77
        Me.splitMain.TabIndex = 7
        '
        'frmSeek
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(707, 586)
        Me.Controls.Add(Me.splitMain)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MinimumSize = New System.Drawing.Size(723, 0)
        Me.Name = "frmSeek"
        Me.Text = "frmSeek"
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel1.PerformLayout()
        Me.splitMain.Panel2.ResumeLayout(False)
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents chkCase As System.Windows.Forms.CheckBox
    Friend WithEvents chkWholeWord As System.Windows.Forms.CheckBox
    Friend WithEvents treeElements As System.Windows.Forms.TreeView
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents chkSearchElement As System.Windows.Forms.CheckBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
End Class
