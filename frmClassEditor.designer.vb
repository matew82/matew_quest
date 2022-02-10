<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmClassEditor
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
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Класс C, Code")
        Me.splitMain = New System.Windows.Forms.SplitContainer()
        Me.treeClasses = New System.Windows.Forms.TreeView()
        Me.treeClassesMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExpandAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CollapseAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SortToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.btnDn = New System.Windows.Forms.Button()
        Me.btnUp = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnNewProperty = New System.Windows.Forms.Button()
        Me.btnNewFunction = New System.Windows.Forms.Button()
        Me.btnNewClass = New System.Windows.Forms.Button()
        Me.pnlProperty = New System.Windows.Forms.Panel()
        Me.btnPropParamShow = New System.Windows.Forms.Button()
        Me.btnPropGenerateHelp = New System.Windows.Forms.Button()
        Me.lstPropElementClass = New System.Windows.Forms.ListBox()
        Me.lblPropElementClass = New System.Windows.Forms.Label()
        Me.cmbPropHidden = New System.Windows.Forms.ComboBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtPropEditorName = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnDuplicateProperty = New System.Windows.Forms.Button()
        Me.btnEditCode = New System.Windows.Forms.Button()
        Me.txtProp = New System.Windows.Forms.TextBox()
        Me.pnlPropBool = New System.Windows.Forms.Panel()
        Me.optPropValueFalse = New System.Windows.Forms.RadioButton()
        Me.optPropValueTrue = New System.Windows.Forms.RadioButton()
        Me.rtbProp = New System.Windows.Forms.RichTextBox()
        Me.pnlDataType = New System.Windows.Forms.Panel()
        Me.optValueCode = New System.Windows.Forms.RadioButton()
        Me.optValueNormal = New System.Windows.Forms.RadioButton()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtPropReturnArray = New System.Windows.Forms.TextBox()
        Me.lblPropReturnArray = New System.Windows.Forms.Label()
        Me.txtPropHelp = New System.Windows.Forms.TextBox()
        Me.btnOpenPropHelpFile = New System.Windows.Forms.Button()
        Me.lstPropReturn = New System.Windows.Forms.ListBox()
        Me.txtPropDescription = New System.Windows.Forms.TextBox()
        Me.txtPropName = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.btnSaveProperty = New System.Windows.Forms.Button()
        Me.pnlFunction = New System.Windows.Forms.Panel()
        Me.btnFuncGenerateHelp = New System.Windows.Forms.Button()
        Me.cmbFuncHidden = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtFuncEditorName = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.btnDuplicateFunction = New System.Windows.Forms.Button()
        Me.pnlFuncParams = New System.Windows.Forms.Panel()
        Me.lstParamsClass = New System.Windows.Forms.ListBox()
        Me.lblParamsClass = New System.Windows.Forms.Label()
        Me.lstFuncParams = New System.Windows.Forms.ListBox()
        Me.btnParamDown = New System.Windows.Forms.Button()
        Me.btnParamUp = New System.Windows.Forms.Button()
        Me.btnParamRemove = New System.Windows.Forms.Button()
        Me.btnParamAdd = New System.Windows.Forms.Button()
        Me.txtParamsEnum = New System.Windows.Forms.TextBox()
        Me.lblParamEnum = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmbParamType = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtParamDescription = New System.Windows.Forms.TextBox()
        Me.txtParamName = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtFuncResult = New System.Windows.Forms.TextBox()
        Me.lblFuncResult = New System.Windows.Forms.Label()
        Me.txtFuncHelpFile = New System.Windows.Forms.TextBox()
        Me.btnOpenFuncHelpFile = New System.Windows.Forms.Button()
        Me.lstFuncReturnType = New System.Windows.Forms.ListBox()
        Me.nudFuncMax = New System.Windows.Forms.NumericUpDown()
        Me.nudFuncMin = New System.Windows.Forms.NumericUpDown()
        Me.txtFuncDescription = New System.Windows.Forms.TextBox()
        Me.txtFuncName = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnSaveFunction = New System.Windows.Forms.Button()
        Me.pnlClass = New System.Windows.Forms.Panel()
        Me.cmbClassDefProperty = New System.Windows.Forms.ComboBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.txtClassHelpFile = New System.Windows.Forms.TextBox()
        Me.btnClassHelpFile = New System.Windows.Forms.Button()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.btnCreateNewClass = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.nudClassLevels = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtNewClassName = New System.Windows.Forms.TextBox()
        Me.pnlPropParams = New System.Windows.Forms.Panel()
        Me.chkPropParamReturnLevel3 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamReturnLevel2 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamReturnLevel1 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamLevel3 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamLevel2 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamLevel1 = New System.Windows.Forms.CheckBox()
        Me.cmbPropParamReturnType = New System.Windows.Forms.ComboBox()
        Me.btnPropParamArrayAdd = New System.Windows.Forms.Button()
        Me.btnPropParamHide = New System.Windows.Forms.Button()
        Me.btnPropParamReturnAdd = New System.Windows.Forms.Button()
        Me.btnPropParamDn = New System.Windows.Forms.Button()
        Me.btnPropParamUp = New System.Windows.Forms.Button()
        Me.btnPropParamRemove = New System.Windows.Forms.Button()
        Me.btnPropParamAdd = New System.Windows.Forms.Button()
        Me.txtPropParamReturnDescription = New System.Windows.Forms.TextBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.txtPropParamDescription = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.txtPropParamName = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.lstPropParams = New System.Windows.Forms.ListBox()
        Me.ofd = New System.Windows.Forms.OpenFileDialog()
        Me.splitCode = New System.Windows.Forms.SplitContainer()
        Me.codeProp = New WindowsApplication1.CodeTextBox()
        Me.btnCodeCancel = New System.Windows.Forms.Button()
        Me.btnCodeSave = New System.Windows.Forms.Button()
        Me.sfd = New System.Windows.Forms.SaveFileDialog()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        Me.treeClassesMenu.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.pnlProperty.SuspendLayout()
        Me.pnlPropBool.SuspendLayout()
        Me.pnlDataType.SuspendLayout()
        Me.pnlFunction.SuspendLayout()
        Me.pnlFuncParams.SuspendLayout()
        CType(Me.nudFuncMax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudFuncMin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlClass.SuspendLayout()
        CType(Me.nudClassLevels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPropParams.SuspendLayout()
        CType(Me.splitCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitCode.Panel1.SuspendLayout()
        Me.splitCode.Panel2.SuspendLayout()
        Me.splitCode.SuspendLayout()
        Me.SuspendLayout()
        '
        'splitMain
        '
        Me.splitMain.Location = New System.Drawing.Point(75, 0)
        Me.splitMain.Name = "splitMain"
        '
        'splitMain.Panel1
        '
        Me.splitMain.Panel1.Controls.Add(Me.treeClasses)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.SplitContainer2)
        Me.splitMain.Size = New System.Drawing.Size(1203, 659)
        Me.splitMain.SplitterDistance = 397
        Me.splitMain.TabIndex = 0
        '
        'treeClasses
        '
        Me.treeClasses.AllowDrop = True
        Me.treeClasses.ContextMenuStrip = Me.treeClassesMenu
        Me.treeClasses.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeClasses.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.treeClasses.ForeColor = System.Drawing.Color.Navy
        Me.treeClasses.Location = New System.Drawing.Point(0, 0)
        Me.treeClasses.Name = "treeClasses"
        TreeNode1.Name = "Узел0"
        TreeNode1.Text = "Класс C, Code"
        Me.treeClasses.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode1})
        Me.treeClasses.Size = New System.Drawing.Size(397, 659)
        Me.treeClasses.TabIndex = 6
        '
        'treeClassesMenu
        '
        Me.treeClassesMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExpandAllToolStripMenuItem, Me.CollapseAllToolStripMenuItem, Me.SortToolStripMenuItem})
        Me.treeClassesMenu.Name = "ContextMenuStrip1"
        Me.treeClassesMenu.Size = New System.Drawing.Size(148, 70)
        '
        'ExpandAllToolStripMenuItem
        '
        Me.ExpandAllToolStripMenuItem.Name = "ExpandAllToolStripMenuItem"
        Me.ExpandAllToolStripMenuItem.Size = New System.Drawing.Size(147, 22)
        Me.ExpandAllToolStripMenuItem.Text = "Раскрыть все"
        '
        'CollapseAllToolStripMenuItem
        '
        Me.CollapseAllToolStripMenuItem.Name = "CollapseAllToolStripMenuItem"
        Me.CollapseAllToolStripMenuItem.Size = New System.Drawing.Size(147, 22)
        Me.CollapseAllToolStripMenuItem.Text = "Скрыть все"
        '
        'SortToolStripMenuItem
        '
        Me.SortToolStripMenuItem.Name = "SortToolStripMenuItem"
        Me.SortToolStripMenuItem.Size = New System.Drawing.Size(147, 22)
        Me.SortToolStripMenuItem.Text = "Сортировать"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.AutoScroll = True
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnDn)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnUp)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnRemove)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnNewProperty)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnNewFunction)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnNewClass)
        Me.SplitContainer2.Panel1.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.pnlFunction)
        Me.SplitContainer2.Panel2.Controls.Add(Me.pnlProperty)
        Me.SplitContainer2.Panel2.Controls.Add(Me.pnlClass)
        Me.SplitContainer2.Panel2.Controls.Add(Me.pnlPropParams)
        Me.SplitContainer2.Size = New System.Drawing.Size(802, 659)
        Me.SplitContainer2.SplitterDistance = 104
        Me.SplitContainer2.TabIndex = 2
        '
        'btnDn
        '
        Me.btnDn.Image = Global.WindowsApplication1.My.Resources.Resources.arrow_down
        Me.btnDn.Location = New System.Drawing.Point(710, 54)
        Me.btnDn.Name = "btnDn"
        Me.btnDn.Size = New System.Drawing.Size(72, 36)
        Me.btnDn.TabIndex = 14
        Me.btnDn.UseVisualStyleBackColor = True
        '
        'btnUp
        '
        Me.btnUp.Image = Global.WindowsApplication1.My.Resources.Resources.arrow_up
        Me.btnUp.Location = New System.Drawing.Point(710, 12)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(72, 36)
        Me.btnUp.TabIndex = 13
        Me.btnUp.UseVisualStyleBackColor = True
        '
        'btnRemove
        '
        Me.btnRemove.Enabled = False
        Me.btnRemove.Location = New System.Drawing.Point(369, 54)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(335, 36)
        Me.btnRemove.TabIndex = 12
        Me.btnRemove.Text = "Удалить"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'btnNewProperty
        '
        Me.btnNewProperty.Location = New System.Drawing.Point(28, 12)
        Me.btnNewProperty.Name = "btnNewProperty"
        Me.btnNewProperty.Size = New System.Drawing.Size(335, 36)
        Me.btnNewProperty.TabIndex = 11
        Me.btnNewProperty.Text = "Новое свойство"
        Me.btnNewProperty.UseVisualStyleBackColor = True
        '
        'btnNewFunction
        '
        Me.btnNewFunction.Location = New System.Drawing.Point(28, 54)
        Me.btnNewFunction.Name = "btnNewFunction"
        Me.btnNewFunction.Size = New System.Drawing.Size(335, 36)
        Me.btnNewFunction.TabIndex = 10
        Me.btnNewFunction.Text = "Новая функция"
        Me.btnNewFunction.UseVisualStyleBackColor = True
        '
        'btnNewClass
        '
        Me.btnNewClass.Location = New System.Drawing.Point(369, 12)
        Me.btnNewClass.Name = "btnNewClass"
        Me.btnNewClass.Size = New System.Drawing.Size(335, 36)
        Me.btnNewClass.TabIndex = 8
        Me.btnNewClass.Text = "Создать новый класс"
        Me.btnNewClass.UseVisualStyleBackColor = True
        '
        'pnlProperty
        '
        Me.pnlProperty.Controls.Add(Me.btnPropParamShow)
        Me.pnlProperty.Controls.Add(Me.btnPropGenerateHelp)
        Me.pnlProperty.Controls.Add(Me.lstPropElementClass)
        Me.pnlProperty.Controls.Add(Me.lblPropElementClass)
        Me.pnlProperty.Controls.Add(Me.cmbPropHidden)
        Me.pnlProperty.Controls.Add(Me.Label15)
        Me.pnlProperty.Controls.Add(Me.txtPropEditorName)
        Me.pnlProperty.Controls.Add(Me.Label14)
        Me.pnlProperty.Controls.Add(Me.btnDuplicateProperty)
        Me.pnlProperty.Controls.Add(Me.btnEditCode)
        Me.pnlProperty.Controls.Add(Me.txtProp)
        Me.pnlProperty.Controls.Add(Me.pnlPropBool)
        Me.pnlProperty.Controls.Add(Me.rtbProp)
        Me.pnlProperty.Controls.Add(Me.pnlDataType)
        Me.pnlProperty.Controls.Add(Me.Label13)
        Me.pnlProperty.Controls.Add(Me.txtPropReturnArray)
        Me.pnlProperty.Controls.Add(Me.lblPropReturnArray)
        Me.pnlProperty.Controls.Add(Me.txtPropHelp)
        Me.pnlProperty.Controls.Add(Me.btnOpenPropHelpFile)
        Me.pnlProperty.Controls.Add(Me.lstPropReturn)
        Me.pnlProperty.Controls.Add(Me.txtPropDescription)
        Me.pnlProperty.Controls.Add(Me.txtPropName)
        Me.pnlProperty.Controls.Add(Me.Label19)
        Me.pnlProperty.Controls.Add(Me.Label22)
        Me.pnlProperty.Controls.Add(Me.Label23)
        Me.pnlProperty.Controls.Add(Me.Label24)
        Me.pnlProperty.Controls.Add(Me.btnSaveProperty)
        Me.pnlProperty.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.pnlProperty.Location = New System.Drawing.Point(28, -15)
        Me.pnlProperty.Name = "pnlProperty"
        Me.pnlProperty.Size = New System.Drawing.Size(771, 563)
        Me.pnlProperty.TabIndex = 12
        Me.pnlProperty.Visible = False
        '
        'btnPropParamShow
        '
        Me.btnPropParamShow.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamShow.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnPropParamShow.Image = Global.WindowsApplication1.My.Resources.Resources.edit32
        Me.btnPropParamShow.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPropParamShow.Location = New System.Drawing.Point(23, 211)
        Me.btnPropParamShow.Name = "btnPropParamShow"
        Me.btnPropParamShow.Size = New System.Drawing.Size(143, 43)
        Me.btnPropParamShow.TabIndex = 36
        Me.btnPropParamShow.Text = "Параметры..."
        Me.btnPropParamShow.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPropParamShow.UseVisualStyleBackColor = True
        Me.btnPropParamShow.Visible = False
        '
        'btnPropGenerateHelp
        '
        Me.btnPropGenerateHelp.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropGenerateHelp.ForeColor = System.Drawing.Color.Navy
        Me.btnPropGenerateHelp.Image = Global.WindowsApplication1.My.Resources.Resources.generateHelp32
        Me.btnPropGenerateHelp.Location = New System.Drawing.Point(237, 496)
        Me.btnPropGenerateHelp.Name = "btnPropGenerateHelp"
        Me.btnPropGenerateHelp.Size = New System.Drawing.Size(52, 43)
        Me.btnPropGenerateHelp.TabIndex = 35
        Me.btnPropGenerateHelp.UseVisualStyleBackColor = True
        '
        'lstPropElementClass
        '
        Me.lstPropElementClass.FormattingEnabled = True
        Me.lstPropElementClass.ItemHeight = 20
        Me.lstPropElementClass.Location = New System.Drawing.Point(232, 313)
        Me.lstPropElementClass.Name = "lstPropElementClass"
        Me.lstPropElementClass.Size = New System.Drawing.Size(512, 164)
        Me.lstPropElementClass.Sorted = True
        Me.lstPropElementClass.TabIndex = 34
        Me.lstPropElementClass.Visible = False
        '
        'lblPropElementClass
        '
        Me.lblPropElementClass.AutoSize = True
        Me.lblPropElementClass.Location = New System.Drawing.Point(271, 293)
        Me.lblPropElementClass.Name = "lblPropElementClass"
        Me.lblPropElementClass.Size = New System.Drawing.Size(452, 20)
        Me.lblPropElementClass.TabIndex = 33
        Me.lblPropElementClass.Text = "Одно из имен класса, которому принадлежит данный элемент"
        Me.lblPropElementClass.Visible = False
        '
        'cmbPropHidden
        '
        Me.cmbPropHidden.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPropHidden.FormattingEnabled = True
        Me.cmbPropHidden.Items.AddRange(New Object() {"Свойство открыто", "Скрыть в редакторе", "Скрыть в коде", "Скрыть полностью", "Только в настройках по умолчанию", "Только для элементов 2 уровня", "Только для элементов 3 уровня", "Только для 1 и элементов  2 уровня", "Только для 1 и элементов  3 уровня", "Только для элементов  2 и 3 уровня"})
        Me.cmbPropHidden.Location = New System.Drawing.Point(495, 49)
        Me.cmbPropHidden.Name = "cmbPropHidden"
        Me.cmbPropHidden.Size = New System.Drawing.Size(257, 28)
        Me.cmbPropHidden.TabIndex = 31
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(392, 52)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(63, 20)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "Скрыть"
        '
        'txtPropEditorName
        '
        Me.txtPropEditorName.Location = New System.Drawing.Point(132, 49)
        Me.txtPropEditorName.Name = "txtPropEditorName"
        Me.txtPropEditorName.Size = New System.Drawing.Size(254, 28)
        Me.txtPropEditorName.TabIndex = 29
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(10, 52)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(126, 20)
        Me.Label14.TabIndex = 28
        Me.Label14.Text = "Имя в редакторе"
        '
        'btnDuplicateProperty
        '
        Me.btnDuplicateProperty.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnDuplicateProperty.ForeColor = System.Drawing.Color.Navy
        Me.btnDuplicateProperty.Image = Global.WindowsApplication1.My.Resources.Resources.duplicate32
        Me.btnDuplicateProperty.Location = New System.Drawing.Point(595, 499)
        Me.btnDuplicateProperty.Name = "btnDuplicateProperty"
        Me.btnDuplicateProperty.Size = New System.Drawing.Size(52, 43)
        Me.btnDuplicateProperty.TabIndex = 27
        Me.btnDuplicateProperty.UseVisualStyleBackColor = True
        '
        'btnEditCode
        '
        Me.btnEditCode.Location = New System.Drawing.Point(701, 216)
        Me.btnEditCode.Margin = New System.Windows.Forms.Padding(2)
        Me.btnEditCode.Name = "btnEditCode"
        Me.btnEditCode.Size = New System.Drawing.Size(35, 29)
        Me.btnEditCode.TabIndex = 26
        Me.btnEditCode.Text = "..."
        Me.btnEditCode.UseVisualStyleBackColor = True
        '
        'txtProp
        '
        Me.txtProp.Location = New System.Drawing.Point(192, 164)
        Me.txtProp.Margin = New System.Windows.Forms.Padding(2)
        Me.txtProp.Name = "txtProp"
        Me.txtProp.Size = New System.Drawing.Size(558, 28)
        Me.txtProp.TabIndex = 25
        '
        'pnlPropBool
        '
        Me.pnlPropBool.Controls.Add(Me.optPropValueFalse)
        Me.pnlPropBool.Controls.Add(Me.optPropValueTrue)
        Me.pnlPropBool.Location = New System.Drawing.Point(193, 207)
        Me.pnlPropBool.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlPropBool.Name = "pnlPropBool"
        Me.pnlPropBool.Size = New System.Drawing.Size(110, 60)
        Me.pnlPropBool.TabIndex = 24
        '
        'optPropValueFalse
        '
        Me.optPropValueFalse.AutoSize = True
        Me.optPropValueFalse.Checked = True
        Me.optPropValueFalse.Location = New System.Drawing.Point(26, 32)
        Me.optPropValueFalse.Margin = New System.Windows.Forms.Padding(2)
        Me.optPropValueFalse.Name = "optPropValueFalse"
        Me.optPropValueFalse.Size = New System.Drawing.Size(59, 24)
        Me.optPropValueFalse.TabIndex = 1
        Me.optPropValueFalse.TabStop = True
        Me.optPropValueFalse.Text = "False"
        Me.optPropValueFalse.UseVisualStyleBackColor = True
        '
        'optPropValueTrue
        '
        Me.optPropValueTrue.AutoSize = True
        Me.optPropValueTrue.Location = New System.Drawing.Point(26, 4)
        Me.optPropValueTrue.Margin = New System.Windows.Forms.Padding(2)
        Me.optPropValueTrue.Name = "optPropValueTrue"
        Me.optPropValueTrue.Size = New System.Drawing.Size(58, 24)
        Me.optPropValueTrue.TabIndex = 0
        Me.optPropValueTrue.Text = "True"
        Me.optPropValueTrue.UseVisualStyleBackColor = True
        '
        'rtbProp
        '
        Me.rtbProp.Cursor = System.Windows.Forms.Cursors.Default
        Me.rtbProp.Location = New System.Drawing.Point(192, 162)
        Me.rtbProp.Name = "rtbProp"
        Me.rtbProp.ReadOnly = True
        Me.rtbProp.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
        Me.rtbProp.Size = New System.Drawing.Size(559, 114)
        Me.rtbProp.TabIndex = 22
        Me.rtbProp.Text = ""
        Me.rtbProp.WordWrap = False
        '
        'pnlDataType
        '
        Me.pnlDataType.Controls.Add(Me.optValueCode)
        Me.pnlDataType.Controls.Add(Me.optValueNormal)
        Me.pnlDataType.Location = New System.Drawing.Point(10, 187)
        Me.pnlDataType.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlDataType.Name = "pnlDataType"
        Me.pnlDataType.Size = New System.Drawing.Size(173, 88)
        Me.pnlDataType.TabIndex = 23
        '
        'optValueCode
        '
        Me.optValueCode.Appearance = System.Windows.Forms.Appearance.Button
        Me.optValueCode.Location = New System.Drawing.Point(14, 48)
        Me.optValueCode.Margin = New System.Windows.Forms.Padding(2)
        Me.optValueCode.Name = "optValueCode"
        Me.optValueCode.Size = New System.Drawing.Size(144, 29)
        Me.optValueCode.TabIndex = 1
        Me.optValueCode.Text = "Скрипт"
        Me.optValueCode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.optValueCode.UseVisualStyleBackColor = True
        '
        'optValueNormal
        '
        Me.optValueNormal.Appearance = System.Windows.Forms.Appearance.Button
        Me.optValueNormal.Checked = True
        Me.optValueNormal.Location = New System.Drawing.Point(12, 14)
        Me.optValueNormal.Margin = New System.Windows.Forms.Padding(2)
        Me.optValueNormal.Name = "optValueNormal"
        Me.optValueNormal.Size = New System.Drawing.Size(146, 29)
        Me.optValueNormal.TabIndex = 0
        Me.optValueNormal.TabStop = True
        Me.optValueNormal.Text = "Простое значение"
        Me.optValueNormal.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.optValueNormal.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(10, 165)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(181, 20)
        Me.Label13.TabIndex = 21
        Me.Label13.Text = "Значение по-умолчанию"
        '
        'txtPropReturnArray
        '
        Me.txtPropReturnArray.Location = New System.Drawing.Point(229, 315)
        Me.txtPropReturnArray.Multiline = True
        Me.txtPropReturnArray.Name = "txtPropReturnArray"
        Me.txtPropReturnArray.Size = New System.Drawing.Size(515, 163)
        Me.txtPropReturnArray.TabIndex = 20
        Me.txtPropReturnArray.Visible = False
        '
        'lblPropReturnArray
        '
        Me.lblPropReturnArray.AutoSize = True
        Me.lblPropReturnArray.Location = New System.Drawing.Point(277, 292)
        Me.lblPropReturnArray.Name = "lblPropReturnArray"
        Me.lblPropReturnArray.Size = New System.Drawing.Size(437, 20)
        Me.lblPropReturnArray.TabIndex = 19
        Me.lblPropReturnArray.Text = "Варианты возвращаемого значения (каждый с новой строки)"
        Me.lblPropReturnArray.Visible = False
        '
        'txtPropHelp
        '
        Me.txtPropHelp.Location = New System.Drawing.Point(495, 16)
        Me.txtPropHelp.Name = "txtPropHelp"
        Me.txtPropHelp.Size = New System.Drawing.Size(221, 28)
        Me.txtPropHelp.TabIndex = 16
        '
        'btnOpenPropHelpFile
        '
        Me.btnOpenPropHelpFile.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnOpenPropHelpFile.Location = New System.Drawing.Point(722, 17)
        Me.btnOpenPropHelpFile.Name = "btnOpenPropHelpFile"
        Me.btnOpenPropHelpFile.Size = New System.Drawing.Size(30, 24)
        Me.btnOpenPropHelpFile.TabIndex = 15
        Me.btnOpenPropHelpFile.Text = "..."
        Me.btnOpenPropHelpFile.UseVisualStyleBackColor = True
        '
        'lstPropReturn
        '
        Me.lstPropReturn.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstPropReturn.FormattingEnabled = True
        Me.lstPropReturn.ItemHeight = 20
        Me.lstPropReturn.Items.AddRange(New Object() {"Обычный", "Да / Нет", "Один из возможных", "Событие", "Длинный текст", "Элемент", "Путь к картинке", "Путь к аудиофайлу", "Путь к текстовому файлу", "Путь к таблице стилей", "Путь к скрипту", "Цвет", "Функция Писателя"})
        Me.lstPropReturn.Location = New System.Drawing.Point(12, 315)
        Me.lstPropReturn.Name = "lstPropReturn"
        Me.lstPropReturn.Size = New System.Drawing.Size(214, 224)
        Me.lstPropReturn.TabIndex = 14
        '
        'txtPropDescription
        '
        Me.txtPropDescription.Location = New System.Drawing.Point(132, 84)
        Me.txtPropDescription.Multiline = True
        Me.txtPropDescription.Name = "txtPropDescription"
        Me.txtPropDescription.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtPropDescription.Size = New System.Drawing.Size(619, 71)
        Me.txtPropDescription.TabIndex = 11
        '
        'txtPropName
        '
        Me.txtPropName.Location = New System.Drawing.Point(132, 17)
        Me.txtPropName.Name = "txtPropName"
        Me.txtPropName.Size = New System.Drawing.Size(254, 28)
        Me.txtPropName.TabIndex = 10
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(31, 292)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(177, 20)
        Me.Label19.TabIndex = 9
        Me.Label19.Text = "Возвращаемое значение"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(392, 19)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(106, 20)
        Me.Label22.TabIndex = 6
        Me.Label22.Text = "Файл помощи"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(10, 87)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(79, 20)
        Me.Label23.TabIndex = 5
        Me.Label23.Text = "Описание"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(10, 19)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(104, 20)
        Me.Label24.TabIndex = 4
        Me.Label24.Text = "Имя свойства"
        '
        'btnSaveProperty
        '
        Me.btnSaveProperty.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnSaveProperty.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSaveProperty.Image = Global.WindowsApplication1.My.Resources.Resources.editProp32
        Me.btnSaveProperty.Location = New System.Drawing.Point(677, 499)
        Me.btnSaveProperty.Name = "btnSaveProperty"
        Me.btnSaveProperty.Size = New System.Drawing.Size(64, 43)
        Me.btnSaveProperty.TabIndex = 3
        Me.btnSaveProperty.UseVisualStyleBackColor = True
        '
        'pnlFunction
        '
        Me.pnlFunction.BackColor = System.Drawing.SystemColors.Control
        Me.pnlFunction.Controls.Add(Me.btnFuncGenerateHelp)
        Me.pnlFunction.Controls.Add(Me.cmbFuncHidden)
        Me.pnlFunction.Controls.Add(Me.Label16)
        Me.pnlFunction.Controls.Add(Me.txtFuncEditorName)
        Me.pnlFunction.Controls.Add(Me.Label17)
        Me.pnlFunction.Controls.Add(Me.btnDuplicateFunction)
        Me.pnlFunction.Controls.Add(Me.pnlFuncParams)
        Me.pnlFunction.Controls.Add(Me.txtFuncResult)
        Me.pnlFunction.Controls.Add(Me.lblFuncResult)
        Me.pnlFunction.Controls.Add(Me.txtFuncHelpFile)
        Me.pnlFunction.Controls.Add(Me.btnOpenFuncHelpFile)
        Me.pnlFunction.Controls.Add(Me.lstFuncReturnType)
        Me.pnlFunction.Controls.Add(Me.nudFuncMax)
        Me.pnlFunction.Controls.Add(Me.nudFuncMin)
        Me.pnlFunction.Controls.Add(Me.txtFuncDescription)
        Me.pnlFunction.Controls.Add(Me.txtFuncName)
        Me.pnlFunction.Controls.Add(Me.Label8)
        Me.pnlFunction.Controls.Add(Me.Label7)
        Me.pnlFunction.Controls.Add(Me.Label6)
        Me.pnlFunction.Controls.Add(Me.Label5)
        Me.pnlFunction.Controls.Add(Me.Label4)
        Me.pnlFunction.Controls.Add(Me.Label3)
        Me.pnlFunction.Controls.Add(Me.btnSaveFunction)
        Me.pnlFunction.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.pnlFunction.Location = New System.Drawing.Point(17, -27)
        Me.pnlFunction.Name = "pnlFunction"
        Me.pnlFunction.Size = New System.Drawing.Size(763, 561)
        Me.pnlFunction.TabIndex = 11
        Me.pnlFunction.Visible = False
        '
        'btnFuncGenerateHelp
        '
        Me.btnFuncGenerateHelp.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnFuncGenerateHelp.ForeColor = System.Drawing.Color.Navy
        Me.btnFuncGenerateHelp.Image = Global.WindowsApplication1.My.Resources.Resources.generateHelp32
        Me.btnFuncGenerateHelp.Location = New System.Drawing.Point(19, 511)
        Me.btnFuncGenerateHelp.Name = "btnFuncGenerateHelp"
        Me.btnFuncGenerateHelp.Size = New System.Drawing.Size(52, 43)
        Me.btnFuncGenerateHelp.TabIndex = 36
        Me.btnFuncGenerateHelp.UseVisualStyleBackColor = True
        '
        'cmbFuncHidden
        '
        Me.cmbFuncHidden.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFuncHidden.FormattingEnabled = True
        Me.cmbFuncHidden.Items.AddRange(New Object() {"Функция открыта", "Скрыть в редакторе", "Скрыть в коде", "Скрыть полностью"})
        Me.cmbFuncHidden.Location = New System.Drawing.Point(495, 53)
        Me.cmbFuncHidden.Name = "cmbFuncHidden"
        Me.cmbFuncHidden.Size = New System.Drawing.Size(257, 28)
        Me.cmbFuncHidden.TabIndex = 35
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(388, 56)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(63, 20)
        Me.Label16.TabIndex = 34
        Me.Label16.Text = "Скрыть"
        '
        'txtFuncEditorName
        '
        Me.txtFuncEditorName.Location = New System.Drawing.Point(135, 53)
        Me.txtFuncEditorName.Name = "txtFuncEditorName"
        Me.txtFuncEditorName.Size = New System.Drawing.Size(250, 28)
        Me.txtFuncEditorName.TabIndex = 33
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(10, 57)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(126, 20)
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "Имя в редакторе"
        '
        'btnDuplicateFunction
        '
        Me.btnDuplicateFunction.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnDuplicateFunction.ForeColor = System.Drawing.Color.Navy
        Me.btnDuplicateFunction.Image = Global.WindowsApplication1.My.Resources.Resources.duplicate32
        Me.btnDuplicateFunction.Location = New System.Drawing.Point(549, 509)
        Me.btnDuplicateFunction.Name = "btnDuplicateFunction"
        Me.btnDuplicateFunction.Size = New System.Drawing.Size(83, 47)
        Me.btnDuplicateFunction.TabIndex = 22
        Me.btnDuplicateFunction.UseVisualStyleBackColor = True
        '
        'pnlFuncParams
        '
        Me.pnlFuncParams.Controls.Add(Me.lstParamsClass)
        Me.pnlFuncParams.Controls.Add(Me.lblParamsClass)
        Me.pnlFuncParams.Controls.Add(Me.lstFuncParams)
        Me.pnlFuncParams.Controls.Add(Me.btnParamDown)
        Me.pnlFuncParams.Controls.Add(Me.btnParamUp)
        Me.pnlFuncParams.Controls.Add(Me.btnParamRemove)
        Me.pnlFuncParams.Controls.Add(Me.btnParamAdd)
        Me.pnlFuncParams.Controls.Add(Me.txtParamsEnum)
        Me.pnlFuncParams.Controls.Add(Me.lblParamEnum)
        Me.pnlFuncParams.Controls.Add(Me.Label12)
        Me.pnlFuncParams.Controls.Add(Me.cmbParamType)
        Me.pnlFuncParams.Controls.Add(Me.Label11)
        Me.pnlFuncParams.Controls.Add(Me.txtParamDescription)
        Me.pnlFuncParams.Controls.Add(Me.txtParamName)
        Me.pnlFuncParams.Controls.Add(Me.Label10)
        Me.pnlFuncParams.Controls.Add(Me.Label9)
        Me.pnlFuncParams.Location = New System.Drawing.Point(362, 169)
        Me.pnlFuncParams.Name = "pnlFuncParams"
        Me.pnlFuncParams.Size = New System.Drawing.Size(398, 334)
        Me.pnlFuncParams.TabIndex = 21
        Me.pnlFuncParams.Visible = False
        '
        'lstParamsClass
        '
        Me.lstParamsClass.FormattingEnabled = True
        Me.lstParamsClass.ItemHeight = 20
        Me.lstParamsClass.Location = New System.Drawing.Point(186, 221)
        Me.lstParamsClass.Name = "lstParamsClass"
        Me.lstParamsClass.Size = New System.Drawing.Size(208, 104)
        Me.lstParamsClass.Sorted = True
        Me.lstParamsClass.TabIndex = 36
        Me.lstParamsClass.Visible = False
        '
        'lblParamsClass
        '
        Me.lblParamsClass.AutoSize = True
        Me.lblParamsClass.Font = New System.Drawing.Font("Palatino Linotype", 9.0!)
        Me.lblParamsClass.Location = New System.Drawing.Point(190, 201)
        Me.lblParamsClass.Name = "lblParamsClass"
        Me.lblParamsClass.Size = New System.Drawing.Size(185, 17)
        Me.lblParamsClass.TabIndex = 35
        Me.lblParamsClass.Text = "Одно из имен класса элемента"
        Me.lblParamsClass.Visible = False
        '
        'lstFuncParams
        '
        Me.lstFuncParams.FormattingEnabled = True
        Me.lstFuncParams.ItemHeight = 20
        Me.lstFuncParams.Location = New System.Drawing.Point(3, 30)
        Me.lstFuncParams.Name = "lstFuncParams"
        Me.lstFuncParams.Size = New System.Drawing.Size(172, 244)
        Me.lstFuncParams.TabIndex = 19
        '
        'btnParamDown
        '
        Me.btnParamDown.Enabled = False
        Me.btnParamDown.Font = New System.Drawing.Font("Palatino Linotype", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnParamDown.ForeColor = System.Drawing.Color.Navy
        Me.btnParamDown.Location = New System.Drawing.Point(94, 288)
        Me.btnParamDown.Name = "btnParamDown"
        Me.btnParamDown.Size = New System.Drawing.Size(33, 33)
        Me.btnParamDown.TabIndex = 32
        Me.btnParamDown.Text = "↓"
        Me.btnParamDown.UseVisualStyleBackColor = True
        '
        'btnParamUp
        '
        Me.btnParamUp.Enabled = False
        Me.btnParamUp.Font = New System.Drawing.Font("Palatino Linotype", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnParamUp.ForeColor = System.Drawing.Color.Navy
        Me.btnParamUp.Location = New System.Drawing.Point(55, 288)
        Me.btnParamUp.Name = "btnParamUp"
        Me.btnParamUp.Size = New System.Drawing.Size(33, 33)
        Me.btnParamUp.TabIndex = 31
        Me.btnParamUp.Text = "↑"
        Me.btnParamUp.UseVisualStyleBackColor = True
        '
        'btnParamRemove
        '
        Me.btnParamRemove.Font = New System.Drawing.Font("Palatino Linotype", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnParamRemove.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnParamRemove.Location = New System.Drawing.Point(133, 288)
        Me.btnParamRemove.Name = "btnParamRemove"
        Me.btnParamRemove.Size = New System.Drawing.Size(42, 33)
        Me.btnParamRemove.TabIndex = 30
        Me.btnParamRemove.Text = "-"
        Me.btnParamRemove.UseVisualStyleBackColor = True
        '
        'btnParamAdd
        '
        Me.btnParamAdd.Font = New System.Drawing.Font("Palatino Linotype", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnParamAdd.ForeColor = System.Drawing.Color.Green
        Me.btnParamAdd.Location = New System.Drawing.Point(7, 289)
        Me.btnParamAdd.Name = "btnParamAdd"
        Me.btnParamAdd.Size = New System.Drawing.Size(42, 33)
        Me.btnParamAdd.TabIndex = 29
        Me.btnParamAdd.Text = "+"
        Me.btnParamAdd.UseVisualStyleBackColor = True
        '
        'txtParamsEnum
        '
        Me.txtParamsEnum.Location = New System.Drawing.Point(187, 221)
        Me.txtParamsEnum.Multiline = True
        Me.txtParamsEnum.Name = "txtParamsEnum"
        Me.txtParamsEnum.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtParamsEnum.Size = New System.Drawing.Size(203, 100)
        Me.txtParamsEnum.TabIndex = 28
        '
        'lblParamEnum
        '
        Me.lblParamEnum.AutoSize = True
        Me.lblParamEnum.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblParamEnum.Location = New System.Drawing.Point(183, 200)
        Me.lblParamEnum.Name = "lblParamEnum"
        Me.lblParamEnum.Size = New System.Drawing.Size(202, 17)
        Me.lblParamEnum.TabIndex = 27
        Me.lblParamEnum.Text = "Варианты (каждый с новой строки)"
        Me.lblParamEnum.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(183, 166)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(36, 20)
        Me.Label12.TabIndex = 26
        Me.Label12.Text = "Тип"
        '
        'cmbParamType
        '
        Me.cmbParamType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbParamType.FormattingEnabled = True
        Me.cmbParamType.Items.AddRange(New Object() {"Массив параметров", "Целое число", "Любое число", "Строка", "Строка или число", "Да / Нет", "Что угодно", "Один из возможных", "Функция", "Дочерний элемент", "Элемент 3 уровня", "Переменная", "Путь к картинке", "Путь к аудиофайлу", "Путь к текстовому файлу", "Путь к таблице стилей", "Путь к скрипту", "Свойство", "Событие"})
        Me.cmbParamType.Location = New System.Drawing.Point(223, 169)
        Me.cmbParamType.Name = "cmbParamType"
        Me.cmbParamType.Size = New System.Drawing.Size(167, 28)
        Me.cmbParamType.TabIndex = 25
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(250, 58)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(79, 20)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "Описание"
        '
        'txtParamDescription
        '
        Me.txtParamDescription.Location = New System.Drawing.Point(185, 81)
        Me.txtParamDescription.Multiline = True
        Me.txtParamDescription.Name = "txtParamDescription"
        Me.txtParamDescription.Size = New System.Drawing.Size(205, 82)
        Me.txtParamDescription.TabIndex = 23
        '
        'txtParamName
        '
        Me.txtParamName.Location = New System.Drawing.Point(181, 27)
        Me.txtParamName.Name = "txtParamName"
        Me.txtParamName.Size = New System.Drawing.Size(209, 28)
        Me.txtParamName.TabIndex = 22
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(198, 4)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(156, 20)
        Me.Label10.TabIndex = 21
        Me.Label10.Text = "Название параметра"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Palatino Linotype", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label9.Location = New System.Drawing.Point(20, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(121, 26)
        Me.Label9.TabIndex = 20
        Me.Label9.Text = "Параметры"
        '
        'txtFuncResult
        '
        Me.txtFuncResult.Location = New System.Drawing.Point(19, 364)
        Me.txtFuncResult.Multiline = True
        Me.txtFuncResult.Name = "txtFuncResult"
        Me.txtFuncResult.Size = New System.Drawing.Size(328, 139)
        Me.txtFuncResult.TabIndex = 20
        Me.txtFuncResult.Visible = False
        '
        'lblFuncResult
        '
        Me.lblFuncResult.AutoSize = True
        Me.lblFuncResult.Location = New System.Drawing.Point(15, 328)
        Me.lblFuncResult.Name = "lblFuncResult"
        Me.lblFuncResult.Size = New System.Drawing.Size(342, 20)
        Me.lblFuncResult.TabIndex = 19
        Me.lblFuncResult.Text = "Варианты результата (каждый с новой строки)"
        Me.lblFuncResult.Visible = False
        '
        'txtFuncHelpFile
        '
        Me.txtFuncHelpFile.Location = New System.Drawing.Point(495, 16)
        Me.txtFuncHelpFile.Name = "txtFuncHelpFile"
        Me.txtFuncHelpFile.Size = New System.Drawing.Size(221, 28)
        Me.txtFuncHelpFile.TabIndex = 16
        '
        'btnOpenFuncHelpFile
        '
        Me.btnOpenFuncHelpFile.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnOpenFuncHelpFile.Location = New System.Drawing.Point(722, 17)
        Me.btnOpenFuncHelpFile.Name = "btnOpenFuncHelpFile"
        Me.btnOpenFuncHelpFile.Size = New System.Drawing.Size(30, 24)
        Me.btnOpenFuncHelpFile.TabIndex = 15
        Me.btnOpenFuncHelpFile.Text = "..."
        Me.btnOpenFuncHelpFile.UseVisualStyleBackColor = True
        '
        'lstFuncReturnType
        '
        Me.lstFuncReturnType.FormattingEnabled = True
        Me.lstFuncReturnType.ItemHeight = 20
        Me.lstFuncReturnType.Items.AddRange(New Object() {"Обычный", "Да / Нет", "Один из возможных"})
        Me.lstFuncReturnType.Location = New System.Drawing.Point(167, 250)
        Me.lstFuncReturnType.Name = "lstFuncReturnType"
        Me.lstFuncReturnType.Size = New System.Drawing.Size(180, 64)
        Me.lstFuncReturnType.TabIndex = 14
        '
        'nudFuncMax
        '
        Me.nudFuncMax.Location = New System.Drawing.Point(294, 205)
        Me.nudFuncMax.Name = "nudFuncMax"
        Me.nudFuncMax.Size = New System.Drawing.Size(53, 28)
        Me.nudFuncMax.TabIndex = 13
        '
        'nudFuncMin
        '
        Me.nudFuncMin.Location = New System.Drawing.Point(294, 168)
        Me.nudFuncMin.Name = "nudFuncMin"
        Me.nudFuncMin.Size = New System.Drawing.Size(53, 28)
        Me.nudFuncMin.TabIndex = 12
        '
        'txtFuncDescription
        '
        Me.txtFuncDescription.Location = New System.Drawing.Point(135, 91)
        Me.txtFuncDescription.Multiline = True
        Me.txtFuncDescription.Name = "txtFuncDescription"
        Me.txtFuncDescription.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtFuncDescription.Size = New System.Drawing.Size(616, 72)
        Me.txtFuncDescription.TabIndex = 11
        '
        'txtFuncName
        '
        Me.txtFuncName.Location = New System.Drawing.Point(135, 16)
        Me.txtFuncName.Name = "txtFuncName"
        Me.txtFuncName.Size = New System.Drawing.Size(250, 28)
        Me.txtFuncName.TabIndex = 10
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(11, 269)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(148, 20)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Результат функции"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(11, 207)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(250, 20)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Максимальное кол-во параметров"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(11, 170)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(245, 20)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Минимальное кол-во параметров"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(388, 19)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(106, 20)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Файл помощи"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 94)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(79, 20)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Описание"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(107, 20)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Имя функции"
        '
        'btnSaveFunction
        '
        Me.btnSaveFunction.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnSaveFunction.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSaveFunction.Image = Global.WindowsApplication1.My.Resources.Resources.editFunc32
        Me.btnSaveFunction.Location = New System.Drawing.Point(664, 509)
        Me.btnSaveFunction.Name = "btnSaveFunction"
        Me.btnSaveFunction.Size = New System.Drawing.Size(87, 47)
        Me.btnSaveFunction.TabIndex = 3
        Me.btnSaveFunction.UseVisualStyleBackColor = True
        '
        'pnlClass
        '
        Me.pnlClass.Controls.Add(Me.cmbClassDefProperty)
        Me.pnlClass.Controls.Add(Me.Label29)
        Me.pnlClass.Controls.Add(Me.txtClassHelpFile)
        Me.pnlClass.Controls.Add(Me.btnClassHelpFile)
        Me.pnlClass.Controls.Add(Me.Label28)
        Me.pnlClass.Controls.Add(Me.btnCreateNewClass)
        Me.pnlClass.Controls.Add(Me.Label2)
        Me.pnlClass.Controls.Add(Me.nudClassLevels)
        Me.pnlClass.Controls.Add(Me.Label1)
        Me.pnlClass.Controls.Add(Me.txtNewClassName)
        Me.pnlClass.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.pnlClass.Location = New System.Drawing.Point(16, 41)
        Me.pnlClass.Name = "pnlClass"
        Me.pnlClass.Size = New System.Drawing.Size(763, 451)
        Me.pnlClass.TabIndex = 0
        Me.pnlClass.Visible = False
        '
        'cmbClassDefProperty
        '
        Me.cmbClassDefProperty.FormattingEnabled = True
        Me.cmbClassDefProperty.Location = New System.Drawing.Point(304, 141)
        Me.cmbClassDefProperty.Name = "cmbClassDefProperty"
        Me.cmbClassDefProperty.Size = New System.Drawing.Size(424, 28)
        Me.cmbClassDefProperty.TabIndex = 22
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(32, 143)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(273, 20)
        Me.Label29.TabIndex = 21
        Me.Label29.Text = "Свойство по умолчанию для 3 уровня"
        '
        'txtClassHelpFile
        '
        Me.txtClassHelpFile.Location = New System.Drawing.Point(411, 87)
        Me.txtClassHelpFile.Name = "txtClassHelpFile"
        Me.txtClassHelpFile.Size = New System.Drawing.Size(277, 28)
        Me.txtClassHelpFile.TabIndex = 19
        '
        'btnClassHelpFile
        '
        Me.btnClassHelpFile.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnClassHelpFile.Location = New System.Drawing.Point(696, 89)
        Me.btnClassHelpFile.Name = "btnClassHelpFile"
        Me.btnClassHelpFile.Size = New System.Drawing.Size(30, 24)
        Me.btnClassHelpFile.TabIndex = 18
        Me.btnClassHelpFile.Text = "..."
        Me.btnClassHelpFile.UseVisualStyleBackColor = True
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(300, 91)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(106, 20)
        Me.Label28.TabIndex = 17
        Me.Label28.Text = "Файл помощи"
        '
        'btnCreateNewClass
        '
        Me.btnCreateNewClass.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCreateNewClass.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnCreateNewClass.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnCreateNewClass.Image = Global.WindowsApplication1.My.Resources.Resources.userClass32
        Me.btnCreateNewClass.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCreateNewClass.Location = New System.Drawing.Point(597, 398)
        Me.btnCreateNewClass.Name = "btnCreateNewClass"
        Me.btnCreateNewClass.Size = New System.Drawing.Size(151, 42)
        Me.btnCreateNewClass.TabIndex = 2
        Me.btnCreateNewClass.Text = "Сохранить"
        Me.btnCreateNewClass.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnCreateNewClass.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(36, 89)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 20)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Уровни класса"
        '
        'nudClassLevels
        '
        Me.nudClassLevels.Location = New System.Drawing.Point(188, 87)
        Me.nudClassLevels.Maximum = New Decimal(New Integer() {3, 0, 0, 0})
        Me.nudClassLevels.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudClassLevels.Name = "nudClassLevels"
        Me.nudClassLevels.Size = New System.Drawing.Size(57, 28)
        Me.nudClassLevels.TabIndex = 1
        Me.nudClassLevels.Value = New Decimal(New Integer() {3, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(33, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(221, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Имена класса (через запятую)"
        '
        'txtNewClassName
        '
        Me.txtNewClassName.Location = New System.Drawing.Point(304, 42)
        Me.txtNewClassName.Name = "txtNewClassName"
        Me.txtNewClassName.Size = New System.Drawing.Size(422, 28)
        Me.txtNewClassName.TabIndex = 0
        '
        'pnlPropParams
        '
        Me.pnlPropParams.Controls.Add(Me.chkPropParamReturnLevel3)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamReturnLevel2)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamReturnLevel1)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamLevel3)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamLevel2)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamLevel1)
        Me.pnlPropParams.Controls.Add(Me.cmbPropParamReturnType)
        Me.pnlPropParams.Controls.Add(Me.btnPropParamArrayAdd)
        Me.pnlPropParams.Controls.Add(Me.btnPropParamHide)
        Me.pnlPropParams.Controls.Add(Me.btnPropParamReturnAdd)
        Me.pnlPropParams.Controls.Add(Me.btnPropParamDn)
        Me.pnlPropParams.Controls.Add(Me.btnPropParamUp)
        Me.pnlPropParams.Controls.Add(Me.btnPropParamRemove)
        Me.pnlPropParams.Controls.Add(Me.btnPropParamAdd)
        Me.pnlPropParams.Controls.Add(Me.txtPropParamReturnDescription)
        Me.pnlPropParams.Controls.Add(Me.Label27)
        Me.pnlPropParams.Controls.Add(Me.Label26)
        Me.pnlPropParams.Controls.Add(Me.Label25)
        Me.pnlPropParams.Controls.Add(Me.txtPropParamDescription)
        Me.pnlPropParams.Controls.Add(Me.Label21)
        Me.pnlPropParams.Controls.Add(Me.txtPropParamName)
        Me.pnlPropParams.Controls.Add(Me.Label20)
        Me.pnlPropParams.Controls.Add(Me.Label18)
        Me.pnlPropParams.Controls.Add(Me.lstPropParams)
        Me.pnlPropParams.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.pnlPropParams.Location = New System.Drawing.Point(6, 32)
        Me.pnlPropParams.Name = "pnlPropParams"
        Me.pnlPropParams.Size = New System.Drawing.Size(748, 452)
        Me.pnlPropParams.TabIndex = 2
        Me.pnlPropParams.Visible = False
        '
        'chkPropParamReturnLevel3
        '
        Me.chkPropParamReturnLevel3.AutoSize = True
        Me.chkPropParamReturnLevel3.Location = New System.Drawing.Point(467, 301)
        Me.chkPropParamReturnLevel3.Name = "chkPropParamReturnLevel3"
        Me.chkPropParamReturnLevel3.Size = New System.Drawing.Size(97, 24)
        Me.chkPropParamReturnLevel3.TabIndex = 46
        Me.chkPropParamReturnLevel3.Tag = "3"
        Me.chkPropParamReturnLevel3.Text = "Уровень 3"
        Me.chkPropParamReturnLevel3.UseVisualStyleBackColor = True
        '
        'chkPropParamReturnLevel2
        '
        Me.chkPropParamReturnLevel2.AutoSize = True
        Me.chkPropParamReturnLevel2.Location = New System.Drawing.Point(365, 301)
        Me.chkPropParamReturnLevel2.Name = "chkPropParamReturnLevel2"
        Me.chkPropParamReturnLevel2.Size = New System.Drawing.Size(97, 24)
        Me.chkPropParamReturnLevel2.TabIndex = 45
        Me.chkPropParamReturnLevel2.Tag = "2"
        Me.chkPropParamReturnLevel2.Text = "Уровень 2"
        Me.chkPropParamReturnLevel2.UseVisualStyleBackColor = True
        '
        'chkPropParamReturnLevel1
        '
        Me.chkPropParamReturnLevel1.AutoSize = True
        Me.chkPropParamReturnLevel1.Location = New System.Drawing.Point(261, 301)
        Me.chkPropParamReturnLevel1.Name = "chkPropParamReturnLevel1"
        Me.chkPropParamReturnLevel1.Size = New System.Drawing.Size(97, 24)
        Me.chkPropParamReturnLevel1.TabIndex = 44
        Me.chkPropParamReturnLevel1.Tag = "1"
        Me.chkPropParamReturnLevel1.Text = "Уровень 1"
        Me.chkPropParamReturnLevel1.UseVisualStyleBackColor = True
        '
        'chkPropParamLevel3
        '
        Me.chkPropParamLevel3.AutoSize = True
        Me.chkPropParamLevel3.Location = New System.Drawing.Point(467, 105)
        Me.chkPropParamLevel3.Name = "chkPropParamLevel3"
        Me.chkPropParamLevel3.Size = New System.Drawing.Size(97, 24)
        Me.chkPropParamLevel3.TabIndex = 43
        Me.chkPropParamLevel3.Tag = "3"
        Me.chkPropParamLevel3.Text = "Уровень 3"
        Me.chkPropParamLevel3.UseVisualStyleBackColor = True
        '
        'chkPropParamLevel2
        '
        Me.chkPropParamLevel2.AutoSize = True
        Me.chkPropParamLevel2.Location = New System.Drawing.Point(365, 105)
        Me.chkPropParamLevel2.Name = "chkPropParamLevel2"
        Me.chkPropParamLevel2.Size = New System.Drawing.Size(97, 24)
        Me.chkPropParamLevel2.TabIndex = 42
        Me.chkPropParamLevel2.Tag = "2"
        Me.chkPropParamLevel2.Text = "Уровень 2"
        Me.chkPropParamLevel2.UseVisualStyleBackColor = True
        '
        'chkPropParamLevel1
        '
        Me.chkPropParamLevel1.AutoSize = True
        Me.chkPropParamLevel1.Location = New System.Drawing.Point(261, 105)
        Me.chkPropParamLevel1.Name = "chkPropParamLevel1"
        Me.chkPropParamLevel1.Size = New System.Drawing.Size(97, 24)
        Me.chkPropParamLevel1.TabIndex = 41
        Me.chkPropParamLevel1.Tag = "1"
        Me.chkPropParamLevel1.Text = "Уровень 1"
        Me.chkPropParamLevel1.UseVisualStyleBackColor = True
        '
        'cmbPropParamReturnType
        '
        Me.cmbPropParamReturnType.FormattingEnabled = True
        Me.cmbPropParamReturnType.Items.AddRange(New Object() {"True", "False", "[что либо другое или ничего]", "[любое число, в т. ч. 0]", "[любая строка]"})
        Me.cmbPropParamReturnType.Location = New System.Drawing.Point(255, 252)
        Me.cmbPropParamReturnType.Name = "cmbPropParamReturnType"
        Me.cmbPropParamReturnType.Size = New System.Drawing.Size(230, 28)
        Me.cmbPropParamReturnType.TabIndex = 40
        '
        'btnPropParamArrayAdd
        '
        Me.btnPropParamArrayAdd.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamArrayAdd.ForeColor = System.Drawing.Color.Green
        Me.btnPropParamArrayAdd.Location = New System.Drawing.Point(607, 55)
        Me.btnPropParamArrayAdd.Name = "btnPropParamArrayAdd"
        Me.btnPropParamArrayAdd.Size = New System.Drawing.Size(118, 41)
        Me.btnPropParamArrayAdd.TabIndex = 39
        Me.btnPropParamArrayAdd.Text = "+ param Array"
        Me.btnPropParamArrayAdd.UseVisualStyleBackColor = True
        '
        'btnPropParamHide
        '
        Me.btnPropParamHide.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamHide.ForeColor = System.Drawing.Color.Navy
        Me.btnPropParamHide.Image = Global.WindowsApplication1.My.Resources.Resources.ok32
        Me.btnPropParamHide.Location = New System.Drawing.Point(627, 388)
        Me.btnPropParamHide.Name = "btnPropParamHide"
        Me.btnPropParamHide.Size = New System.Drawing.Size(83, 47)
        Me.btnPropParamHide.TabIndex = 38
        Me.btnPropParamHide.UseVisualStyleBackColor = True
        '
        'btnPropParamReturnAdd
        '
        Me.btnPropParamReturnAdd.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamReturnAdd.ForeColor = System.Drawing.Color.Green
        Me.btnPropParamReturnAdd.Location = New System.Drawing.Point(494, 245)
        Me.btnPropParamReturnAdd.Name = "btnPropParamReturnAdd"
        Me.btnPropParamReturnAdd.Size = New System.Drawing.Size(106, 41)
        Me.btnPropParamReturnAdd.TabIndex = 37
        Me.btnPropParamReturnAdd.Text = "+ Return x"
        Me.btnPropParamReturnAdd.UseVisualStyleBackColor = True
        '
        'btnPropParamDn
        '
        Me.btnPropParamDn.Enabled = False
        Me.btnPropParamDn.Font = New System.Drawing.Font("Palatino Linotype", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamDn.ForeColor = System.Drawing.Color.Navy
        Me.btnPropParamDn.Location = New System.Drawing.Point(105, 388)
        Me.btnPropParamDn.Name = "btnPropParamDn"
        Me.btnPropParamDn.Size = New System.Drawing.Size(33, 33)
        Me.btnPropParamDn.TabIndex = 36
        Me.btnPropParamDn.Text = "↓"
        Me.btnPropParamDn.UseVisualStyleBackColor = True
        '
        'btnPropParamUp
        '
        Me.btnPropParamUp.Enabled = False
        Me.btnPropParamUp.Font = New System.Drawing.Font("Palatino Linotype", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamUp.ForeColor = System.Drawing.Color.Navy
        Me.btnPropParamUp.Location = New System.Drawing.Point(66, 388)
        Me.btnPropParamUp.Name = "btnPropParamUp"
        Me.btnPropParamUp.Size = New System.Drawing.Size(33, 33)
        Me.btnPropParamUp.TabIndex = 35
        Me.btnPropParamUp.Text = "↑"
        Me.btnPropParamUp.UseVisualStyleBackColor = True
        '
        'btnPropParamRemove
        '
        Me.btnPropParamRemove.Font = New System.Drawing.Font("Palatino Linotype", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamRemove.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnPropParamRemove.Location = New System.Drawing.Point(144, 388)
        Me.btnPropParamRemove.Name = "btnPropParamRemove"
        Me.btnPropParamRemove.Size = New System.Drawing.Size(42, 33)
        Me.btnPropParamRemove.TabIndex = 34
        Me.btnPropParamRemove.Text = "-"
        Me.btnPropParamRemove.UseVisualStyleBackColor = True
        '
        'btnPropParamAdd
        '
        Me.btnPropParamAdd.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamAdd.ForeColor = System.Drawing.Color.Green
        Me.btnPropParamAdd.Location = New System.Drawing.Point(491, 55)
        Me.btnPropParamAdd.Name = "btnPropParamAdd"
        Me.btnPropParamAdd.Size = New System.Drawing.Size(110, 41)
        Me.btnPropParamAdd.TabIndex = 33
        Me.btnPropParamAdd.Text = "+ Param[x]"
        Me.btnPropParamAdd.UseVisualStyleBackColor = True
        '
        'txtPropParamReturnDescription
        '
        Me.txtPropParamReturnDescription.Location = New System.Drawing.Point(255, 352)
        Me.txtPropParamReturnDescription.Name = "txtPropParamReturnDescription"
        Me.txtPropParamReturnDescription.Size = New System.Drawing.Size(473, 28)
        Me.txtPropParamReturnDescription.TabIndex = 10
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(255, 328)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(232, 20)
        Me.Label27.TabIndex = 9
        Me.Label27.Text = "Действия при данном значении"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(257, 229)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(177, 20)
        Me.Label26.TabIndex = 8
        Me.Label26.Text = "Возвращаемое значение"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.Maroon
        Me.Label25.Location = New System.Drawing.Point(252, 198)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(471, 21)
        Me.Label25.TabIndex = 6
        Me.Label25.Text = "Настройка параметров, возвращаемых свойством - событием"
        '
        'txtPropParamDescription
        '
        Me.txtPropParamDescription.Location = New System.Drawing.Point(252, 156)
        Me.txtPropParamDescription.Name = "txtPropParamDescription"
        Me.txtPropParamDescription.Size = New System.Drawing.Size(473, 28)
        Me.txtPropParamDescription.TabIndex = 5
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(252, 132)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(79, 20)
        Me.Label21.TabIndex = 4
        Me.Label21.Text = "Описание"
        '
        'txtPropParamName
        '
        Me.txtPropParamName.Location = New System.Drawing.Point(256, 62)
        Me.txtPropParamName.Name = "txtPropParamName"
        Me.txtPropParamName.Size = New System.Drawing.Size(229, 28)
        Me.txtPropParamName.TabIndex = 3
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(251, 38)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(234, 20)
        Me.Label20.TabIndex = 2
        Me.Label20.Text = "Рекомендуемое имя переменной"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Maroon
        Me.Label18.Location = New System.Drawing.Point(252, 11)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(470, 21)
        Me.Label18.TabIndex = 1
        Me.Label18.Text = "Настройка параметров, принимаемых свойством - событием"
        '
        'lstPropParams
        '
        Me.lstPropParams.FormattingEnabled = True
        Me.lstPropParams.ItemHeight = 20
        Me.lstPropParams.Location = New System.Drawing.Point(11, 18)
        Me.lstPropParams.Name = "lstPropParams"
        Me.lstPropParams.Size = New System.Drawing.Size(230, 364)
        Me.lstPropParams.TabIndex = 0
        '
        'ofd
        '
        Me.ofd.Filter = "(html-файлы)|*.html;*.htm"
        '
        'splitCode
        '
        Me.splitCode.Location = New System.Drawing.Point(27, 51)
        Me.splitCode.Margin = New System.Windows.Forms.Padding(2)
        Me.splitCode.MinimumSize = New System.Drawing.Size(75, 325)
        Me.splitCode.Name = "splitCode"
        Me.splitCode.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitCode.Panel1
        '
        Me.splitCode.Panel1.Controls.Add(Me.codeProp)
        Me.splitCode.Panel1.Font = New System.Drawing.Font("Palatino Linotype", 11.25!)
        Me.splitCode.Panel1MinSize = 300
        '
        'splitCode.Panel2
        '
        Me.splitCode.Panel2.Controls.Add(Me.btnCodeCancel)
        Me.splitCode.Panel2.Controls.Add(Me.btnCodeSave)
        Me.splitCode.Size = New System.Drawing.Size(381, 369)
        Me.splitCode.SplitterDistance = 318
        Me.splitCode.SplitterWidth = 3
        Me.splitCode.TabIndex = 1
        Me.splitCode.Visible = False
        '
        'codeProp
        '
        Me.codeProp.AutoWordSelection = False
        Me.codeProp.CanDrawWords = True
        Me.codeProp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.codeProp.Location = New System.Drawing.Point(0, 0)
        Me.codeProp.Margin = New System.Windows.Forms.Padding(2)
        Me.codeProp.Multiline = True
        Me.codeProp.Name = "codeProp"
        Me.codeProp.Size = New System.Drawing.Size(381, 318)
        Me.codeProp.TabIndex = 0
        '
        'btnCodeCancel
        '
        Me.btnCodeCancel.Font = New System.Drawing.Font("Palatino Linotype", 11.25!)
        Me.btnCodeCancel.Image = Global.WindowsApplication1.My.Resources.Resources.delete32
        Me.btnCodeCancel.Location = New System.Drawing.Point(11, 3)
        Me.btnCodeCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCodeCancel.Name = "btnCodeCancel"
        Me.btnCodeCancel.Size = New System.Drawing.Size(64, 43)
        Me.btnCodeCancel.TabIndex = 1
        Me.btnCodeCancel.UseVisualStyleBackColor = True
        '
        'btnCodeSave
        '
        Me.btnCodeSave.Font = New System.Drawing.Font("Palatino Linotype", 11.25!)
        Me.btnCodeSave.Image = Global.WindowsApplication1.My.Resources.Resources.save32
        Me.btnCodeSave.Location = New System.Drawing.Point(79, 3)
        Me.btnCodeSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCodeSave.Name = "btnCodeSave"
        Me.btnCodeSave.Size = New System.Drawing.Size(66, 43)
        Me.btnCodeSave.TabIndex = 0
        Me.btnCodeSave.UseVisualStyleBackColor = True
        '
        'sfd
        '
        Me.sfd.Filter = "HTML files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*"
        '
        'frmClassEditor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1279, 671)
        Me.Controls.Add(Me.splitCode)
        Me.Controls.Add(Me.splitMain)
        Me.Name = "frmClassEditor"
        Me.Text = "Редактор классов"
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel2.ResumeLayout(False)
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.treeClassesMenu.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.pnlProperty.ResumeLayout(False)
        Me.pnlProperty.PerformLayout()
        Me.pnlPropBool.ResumeLayout(False)
        Me.pnlPropBool.PerformLayout()
        Me.pnlDataType.ResumeLayout(False)
        Me.pnlFunction.ResumeLayout(False)
        Me.pnlFunction.PerformLayout()
        Me.pnlFuncParams.ResumeLayout(False)
        Me.pnlFuncParams.PerformLayout()
        CType(Me.nudFuncMax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudFuncMin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlClass.ResumeLayout(False)
        Me.pnlClass.PerformLayout()
        CType(Me.nudClassLevels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPropParams.ResumeLayout(False)
        Me.pnlPropParams.PerformLayout()
        Me.splitCode.Panel1.ResumeLayout(False)
        Me.splitCode.Panel2.ResumeLayout(false)
        CType(Me.splitCode,System.ComponentModel.ISupportInitialize).EndInit
        Me.splitCode.ResumeLayout(false)
        Me.ResumeLayout(false)

End Sub
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents treeClasses As System.Windows.Forms.TreeView
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnNewClass As System.Windows.Forms.Button
    Friend WithEvents btnNewProperty As System.Windows.Forms.Button
    Friend WithEvents btnNewFunction As System.Windows.Forms.Button
    Friend WithEvents pnlClass As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents nudClassLevels As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtNewClassName As System.Windows.Forms.TextBox
    Friend WithEvents btnCreateNewClass As System.Windows.Forms.Button
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents pnlProperty As System.Windows.Forms.Panel
    Friend WithEvents txtPropReturnArray As System.Windows.Forms.TextBox
    Friend WithEvents lblPropReturnArray As System.Windows.Forms.Label
    Friend WithEvents txtPropHelp As System.Windows.Forms.TextBox
    Friend WithEvents btnOpenPropHelpFile As System.Windows.Forms.Button
    Friend WithEvents lstPropReturn As System.Windows.Forms.ListBox
    Friend WithEvents txtPropDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtPropName As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents btnSaveProperty As System.Windows.Forms.Button
    Friend WithEvents pnlFunction As System.Windows.Forms.Panel
    Friend WithEvents pnlFuncParams As System.Windows.Forms.Panel
    Friend WithEvents lstFuncParams As System.Windows.Forms.ListBox
    Friend WithEvents btnParamDown As System.Windows.Forms.Button
    Friend WithEvents btnParamUp As System.Windows.Forms.Button
    Friend WithEvents btnParamRemove As System.Windows.Forms.Button
    Friend WithEvents btnParamAdd As System.Windows.Forms.Button
    Friend WithEvents txtParamsEnum As System.Windows.Forms.TextBox
    Friend WithEvents lblParamEnum As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cmbParamType As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtParamDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtParamName As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtFuncResult As System.Windows.Forms.TextBox
    Friend WithEvents lblFuncResult As System.Windows.Forms.Label
    Friend WithEvents txtFuncHelpFile As System.Windows.Forms.TextBox
    Friend WithEvents btnOpenFuncHelpFile As System.Windows.Forms.Button
    Friend WithEvents lstFuncReturnType As System.Windows.Forms.ListBox
    Friend WithEvents nudFuncMax As System.Windows.Forms.NumericUpDown
    Friend WithEvents nudFuncMin As System.Windows.Forms.NumericUpDown
    Friend WithEvents txtFuncDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtFuncName As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnSaveFunction As System.Windows.Forms.Button
    Friend WithEvents treeClassesMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExpandAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CollapseAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SortToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents rtbProp As System.Windows.Forms.RichTextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtProp As System.Windows.Forms.TextBox
    Friend WithEvents pnlPropBool As System.Windows.Forms.Panel
    Friend WithEvents optPropValueFalse As System.Windows.Forms.RadioButton
    Friend WithEvents optPropValueTrue As System.Windows.Forms.RadioButton
    Friend WithEvents pnlDataType As System.Windows.Forms.Panel
    Friend WithEvents optValueCode As System.Windows.Forms.RadioButton
    Friend WithEvents optValueNormal As System.Windows.Forms.RadioButton
    Friend WithEvents btnEditCode As System.Windows.Forms.Button
    Friend WithEvents splitCode As System.Windows.Forms.SplitContainer
    Friend WithEvents codeProp As CodeTextBox
    Friend WithEvents btnCodeSave As System.Windows.Forms.Button
    Friend WithEvents btnCodeCancel As System.Windows.Forms.Button
    Friend WithEvents btnDuplicateFunction As System.Windows.Forms.Button
    Friend WithEvents btnDuplicateProperty As System.Windows.Forms.Button
    Friend WithEvents cmbPropHidden As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtPropEditorName As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents cmbFuncHidden As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtFuncEditorName As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents lblPropElementClass As System.Windows.Forms.Label
    Friend WithEvents lstPropElementClass As System.Windows.Forms.ListBox
    Friend WithEvents btnUp As System.Windows.Forms.Button
    Friend WithEvents btnDn As System.Windows.Forms.Button
    Friend WithEvents lstParamsClass As System.Windows.Forms.ListBox
    Friend WithEvents lblParamsClass As System.Windows.Forms.Label
    Friend WithEvents btnPropGenerateHelp As System.Windows.Forms.Button
    Friend WithEvents sfd As System.Windows.Forms.SaveFileDialog
    Friend WithEvents pnlPropParams As System.Windows.Forms.Panel
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents txtPropParamDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtPropParamName As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents lstPropParams As System.Windows.Forms.ListBox
    Friend WithEvents btnPropParamDn As System.Windows.Forms.Button
    Friend WithEvents btnPropParamUp As System.Windows.Forms.Button
    Friend WithEvents btnPropParamRemove As System.Windows.Forms.Button
    Friend WithEvents btnPropParamAdd As System.Windows.Forms.Button
    Friend WithEvents txtPropParamReturnDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents btnPropParamHide As System.Windows.Forms.Button
    Friend WithEvents btnPropParamReturnAdd As System.Windows.Forms.Button
    Friend WithEvents btnPropParamArrayAdd As System.Windows.Forms.Button
    Friend WithEvents btnPropParamShow As System.Windows.Forms.Button
    Friend WithEvents cmbPropParamReturnType As System.Windows.Forms.ComboBox
    Friend WithEvents chkPropParamReturnLevel3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamReturnLevel2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamReturnLevel1 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamLevel3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamLevel2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamLevel1 As System.Windows.Forms.CheckBox
    Friend WithEvents btnFuncGenerateHelp As System.Windows.Forms.Button
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents txtClassHelpFile As System.Windows.Forms.TextBox
    Friend WithEvents btnClassHelpFile As System.Windows.Forms.Button
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents cmbClassDefProperty As System.Windows.Forms.ComboBox
End Class
