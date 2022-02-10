<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMainEditor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMainEditor))
        Me.SplitOuter = New System.Windows.Forms.SplitContainer()
        Me.pnlSeek = New System.Windows.Forms.Panel()
        Me.btnSeekBackward = New System.Windows.Forms.Button()
        Me.btnFindClobal = New System.Windows.Forms.Button()
        Me.btnSeekForward = New System.Windows.Forms.Button()
        Me.chkSeekWholeWord = New System.Windows.Forms.CheckBox()
        Me.chkSeekCase = New System.Windows.Forms.CheckBox()
        Me.cmbSeek = New System.Windows.Forms.ComboBox()
        Me.SplitInner = New System.Windows.Forms.SplitContainer()
        Me.pnlVariables = New System.Windows.Forms.Panel()
        Me.lblVariableError = New System.Windows.Forms.Label()
        Me.dgwVariables = New System.Windows.Forms.DataGridView()
        Me.colSignature = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colValue = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtVariableDescription = New System.Windows.Forms.TextBox()
        Me.lblVariableDescription = New System.Windows.Forms.Label()
        Me.cmnuMainTree = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiAddGroup = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiAddElement = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiRemoveElement = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiDuplicate = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiExcludeFromGroup = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator13 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiExpand = New System.Windows.Forms.ToolStripMenuItem()
        Me.СвернутьВсеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLoad = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.СохранитьКакToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СвойстваКвестаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiRun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.ПоследниеФайлыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВыходИзРедактораToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВыходИзПрограммыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiUndo = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiRedo = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiSelLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiComment = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiUncomment = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiKeyboard = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiWrap = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiDefVars = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiExec = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ФорматированиеТекстаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiB = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiI = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiU = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSup = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSub = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLeft = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiCenter = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiRight = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiJustify = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiP = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiH1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiH2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiH3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSpan = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiFontSize = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiColor = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSpecialChars = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiHR = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiBR = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSpace = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiDash = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiQuotes = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiEuro = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSelector = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiChars = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiImage = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiTable = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiList = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiMarquee = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiIFrame = New System.Windows.Forms.ToolStripMenuItem()
        Me.СервисToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiFind = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiFindNext = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiFindGlobal = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.НастройкиПрограммыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiExecuteScript = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiCopyAsHTML = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiClassEditor = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiRestoreElement = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmigHighLight = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiHighLightFull = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiHighLightDesignate = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiHighLightNo = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПомощьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiShowAdditionalClasses = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПомощьToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripPanels = New System.Windows.Forms.ToolStrip()
        Me.ttMain = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolStripCodeBox = New System.Windows.Forms.ToolStrip()
        Me.tsbDefVars = New System.Windows.Forms.ToolStripButton()
        Me.tsbUndo = New System.Windows.Forms.ToolStripButton()
        Me.tsbRedo = New System.Windows.Forms.ToolStripButton()
        Me.tsbWrap = New System.Windows.Forms.ToolStripButton()
        Me.tsbComment = New System.Windows.Forms.ToolStripButton()
        Me.tsbUncomment = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator10 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbBR = New System.Windows.Forms.ToolStripButton()
        Me.tsbHR = New System.Windows.Forms.ToolStripButton()
        Me.tsbB = New System.Windows.Forms.ToolStripButton()
        Me.tsbI = New System.Windows.Forms.ToolStripButton()
        Me.tsbU = New System.Windows.Forms.ToolStripButton()
        Me.tsbSup = New System.Windows.Forms.ToolStripButton()
        Me.tsbSub = New System.Windows.Forms.ToolStripButton()
        Me.tsbLeft = New System.Windows.Forms.ToolStripButton()
        Me.tsbCenter = New System.Windows.Forms.ToolStripButton()
        Me.tsbRight = New System.Windows.Forms.ToolStripButton()
        Me.tsbJustify = New System.Windows.Forms.ToolStripButton()
        Me.tsbTable = New System.Windows.Forms.ToolStripButton()
        Me.tsbColor = New System.Windows.Forms.ToolStripButton()
        Me.tsbFontSize = New System.Windows.Forms.ToolStripButton()
        Me.tsbImage = New System.Windows.Forms.ToolStripButton()
        Me.tsbMarquee = New System.Windows.Forms.ToolStripButton()
        Me.tsbList = New System.Windows.Forms.ToolStripButton()
        Me.tsbSpan = New System.Windows.Forms.ToolStripButton()
        Me.tsbP = New System.Windows.Forms.ToolStripButton()
        Me.tsbH1 = New System.Windows.Forms.ToolStripButton()
        Me.tsbH2 = New System.Windows.Forms.ToolStripButton()
        Me.tsbH3 = New System.Windows.Forms.ToolStripButton()
        Me.tsbExec = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbSeek = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator12 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbSelector = New System.Windows.Forms.ToolStripButton()
        Me.tsbSpace = New System.Windows.Forms.ToolStripButton()
        Me.tsbDash = New System.Windows.Forms.ToolStripButton()
        Me.tsbQuotes = New System.Windows.Forms.ToolStripButton()
        Me.tsbChars = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator17 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbgHighLight = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsbHighLightFull = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsbHighLightDesignate = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsbHighLightNo = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsbRestoreInitial = New System.Windows.Forms.ToolStripButton()
        Me.tsbFullScreen = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripMain = New System.Windows.Forms.ToolStrip()
        Me.tsbQ = New System.Windows.Forms.ToolStripButton()
        Me.tsbString = New System.Windows.Forms.ToolStripButton()
        Me.tsbMath = New System.Windows.Forms.ToolStripButton()
        Me.tsbDate = New System.Windows.Forms.ToolStripButton()
        Me.tsbArr = New System.Windows.Forms.ToolStripButton()
        Me.tsbFile = New System.Windows.Forms.ToolStripButton()
        Me.tsbCode = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbL = New System.Windows.Forms.ToolStripButton()
        Me.tsbObj = New System.Windows.Forms.ToolStripButton()
        Me.tsbM = New System.Windows.Forms.ToolStripButton()
        Me.tsbT = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator14 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbMap = New System.Windows.Forms.ToolStripButton()
        Me.tsbMed = New System.Windows.Forms.ToolStripButton()
        Me.tsbHer = New System.Windows.Forms.ToolStripButton()
        Me.tsbMg = New System.Windows.Forms.ToolStripButton()
        Me.tsbAb = New System.Windows.Forms.ToolStripButton()
        Me.tsbArmy = New System.Windows.Forms.ToolStripButton()
        Me.tsbBat = New System.Windows.Forms.ToolStripButton()
        Me.tsddbWindows = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsmiCm = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLW = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiDW = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiOW = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiAW = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiMgW = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator15 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbFunc = New System.Windows.Forms.ToolStripButton()
        Me.tsbVar = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator16 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbAddUserClass = New System.Windows.Forms.ToolStripButton()
        Me.imgLstGroupIcons = New System.Windows.Forms.ImageList(Me.components)
        Me.fsWatcher = New System.IO.FileSystemWatcher()
        CType(Me.SplitOuter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitOuter.Panel1.SuspendLayout()
        Me.SplitOuter.Panel2.SuspendLayout()
        Me.SplitOuter.SuspendLayout()
        Me.pnlSeek.SuspendLayout()
        CType(Me.SplitInner, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitInner.Panel2.SuspendLayout()
        Me.SplitInner.SuspendLayout()
        Me.pnlVariables.SuspendLayout()
        CType(Me.dgwVariables, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cmnuMainTree.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.ToolStripCodeBox.SuspendLayout()
        Me.ToolStripMain.SuspendLayout()
        CType(Me.fsWatcher, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitOuter
        '
        resources.ApplyResources(Me.SplitOuter, "SplitOuter")
        Me.SplitOuter.Name = "SplitOuter"
        '
        'SplitOuter.Panel1
        '
        Me.SplitOuter.Panel1.Controls.Add(Me.pnlSeek)
        '
        'SplitOuter.Panel2
        '
        Me.SplitOuter.Panel2.Controls.Add(Me.SplitInner)
        '
        'pnlSeek
        '
        Me.pnlSeek.Controls.Add(Me.btnSeekBackward)
        Me.pnlSeek.Controls.Add(Me.btnFindClobal)
        Me.pnlSeek.Controls.Add(Me.btnSeekForward)
        Me.pnlSeek.Controls.Add(Me.chkSeekWholeWord)
        Me.pnlSeek.Controls.Add(Me.chkSeekCase)
        Me.pnlSeek.Controls.Add(Me.cmbSeek)
        resources.ApplyResources(Me.pnlSeek, "pnlSeek")
        Me.pnlSeek.Name = "pnlSeek"
        '
        'btnSeekBackward
        '
        resources.ApplyResources(Me.btnSeekBackward, "btnSeekBackward")
        Me.btnSeekBackward.Image = Global.WindowsApplication1.My.Resources.Resources.arrow_left
        Me.btnSeekBackward.Name = "btnSeekBackward"
        Me.ttMain.SetToolTip(Me.btnSeekBackward, resources.GetString("btnSeekBackward.ToolTip"))
        Me.btnSeekBackward.UseVisualStyleBackColor = True
        '
        'btnFindClobal
        '
        Me.btnFindClobal.Image = Global.WindowsApplication1.My.Resources.Resources.radar
        resources.ApplyResources(Me.btnFindClobal, "btnFindClobal")
        Me.btnFindClobal.Name = "btnFindClobal"
        Me.ttMain.SetToolTip(Me.btnFindClobal, resources.GetString("btnFindClobal.ToolTip"))
        Me.btnFindClobal.UseVisualStyleBackColor = True
        '
        'btnSeekForward
        '
        resources.ApplyResources(Me.btnSeekForward, "btnSeekForward")
        Me.btnSeekForward.Image = Global.WindowsApplication1.My.Resources.Resources.arrow_right
        Me.btnSeekForward.Name = "btnSeekForward"
        Me.ttMain.SetToolTip(Me.btnSeekForward, resources.GetString("btnSeekForward.ToolTip"))
        Me.btnSeekForward.UseVisualStyleBackColor = True
        '
        'chkSeekWholeWord
        '
        resources.ApplyResources(Me.chkSeekWholeWord, "chkSeekWholeWord")
        Me.chkSeekWholeWord.Name = "chkSeekWholeWord"
        Me.chkSeekWholeWord.UseVisualStyleBackColor = True
        '
        'chkSeekCase
        '
        resources.ApplyResources(Me.chkSeekCase, "chkSeekCase")
        Me.chkSeekCase.Name = "chkSeekCase"
        Me.chkSeekCase.UseVisualStyleBackColor = True
        '
        'cmbSeek
        '
        Me.cmbSeek.FormattingEnabled = True
        resources.ApplyResources(Me.cmbSeek, "cmbSeek")
        Me.cmbSeek.Name = "cmbSeek"
        '
        'SplitInner
        '
        resources.ApplyResources(Me.SplitInner, "SplitInner")
        Me.SplitInner.Name = "SplitInner"
        '
        'SplitInner.Panel2
        '
        Me.SplitInner.Panel2.Controls.Add(Me.pnlVariables)
        '
        'pnlVariables
        '
        Me.pnlVariables.Controls.Add(Me.lblVariableError)
        Me.pnlVariables.Controls.Add(Me.dgwVariables)
        Me.pnlVariables.Controls.Add(Me.txtVariableDescription)
        Me.pnlVariables.Controls.Add(Me.lblVariableDescription)
        resources.ApplyResources(Me.pnlVariables, "pnlVariables")
        Me.pnlVariables.Name = "pnlVariables"
        '
        'lblVariableError
        '
        resources.ApplyResources(Me.lblVariableError, "lblVariableError")
        Me.lblVariableError.ForeColor = System.Drawing.Color.Red
        Me.lblVariableError.Name = "lblVariableError"
        '
        'dgwVariables
        '
        Me.dgwVariables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgwVariables.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colSignature, Me.colValue})
        resources.ApplyResources(Me.dgwVariables, "dgwVariables")
        Me.dgwVariables.Name = "dgwVariables"
        '
        'colSignature
        '
        Me.colSignature.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.colSignature.FillWeight = 20.0!
        resources.ApplyResources(Me.colSignature, "colSignature")
        Me.colSignature.Name = "colSignature"
        '
        'colValue
        '
        Me.colValue.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.colValue.FillWeight = 80.0!
        resources.ApplyResources(Me.colValue, "colValue")
        Me.colValue.Name = "colValue"
        '
        'txtVariableDescription
        '
        resources.ApplyResources(Me.txtVariableDescription, "txtVariableDescription")
        Me.txtVariableDescription.Name = "txtVariableDescription"
        '
        'lblVariableDescription
        '
        resources.ApplyResources(Me.lblVariableDescription, "lblVariableDescription")
        Me.lblVariableDescription.Name = "lblVariableDescription"
        '
        'cmnuMainTree
        '
        Me.cmnuMainTree.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiAddGroup, Me.tsmiAddElement, Me.tsmiRemoveElement, Me.tsmiDuplicate, Me.tsmiExcludeFromGroup, Me.ToolStripSeparator13, Me.tsmiExpand, Me.СвернутьВсеToolStripMenuItem})
        Me.cmnuMainTree.Name = "cmnuTreeLocations"
        resources.ApplyResources(Me.cmnuMainTree, "cmnuMainTree")
        '
        'tsmiAddGroup
        '
        Me.tsmiAddGroup.Image = Global.WindowsApplication1.My.Resources.Resources.add_group26
        Me.tsmiAddGroup.Name = "tsmiAddGroup"
        resources.ApplyResources(Me.tsmiAddGroup, "tsmiAddGroup")
        '
        'tsmiAddElement
        '
        Me.tsmiAddElement.Image = Global.WindowsApplication1.My.Resources.Resources.add26
        Me.tsmiAddElement.Name = "tsmiAddElement"
        resources.ApplyResources(Me.tsmiAddElement, "tsmiAddElement")
        '
        'tsmiRemoveElement
        '
        Me.tsmiRemoveElement.Image = Global.WindowsApplication1.My.Resources.Resources.delete26
        Me.tsmiRemoveElement.Name = "tsmiRemoveElement"
        resources.ApplyResources(Me.tsmiRemoveElement, "tsmiRemoveElement")
        '
        'tsmiDuplicate
        '
        Me.tsmiDuplicate.Image = Global.WindowsApplication1.My.Resources.Resources.duplicate26
        Me.tsmiDuplicate.Name = "tsmiDuplicate"
        resources.ApplyResources(Me.tsmiDuplicate, "tsmiDuplicate")
        '
        'tsmiExcludeFromGroup
        '
        Me.tsmiExcludeFromGroup.Image = Global.WindowsApplication1.My.Resources.Resources.ungroup26
        Me.tsmiExcludeFromGroup.Name = "tsmiExcludeFromGroup"
        resources.ApplyResources(Me.tsmiExcludeFromGroup, "tsmiExcludeFromGroup")
        '
        'ToolStripSeparator13
        '
        Me.ToolStripSeparator13.Name = "ToolStripSeparator13"
        resources.ApplyResources(Me.ToolStripSeparator13, "ToolStripSeparator13")
        '
        'tsmiExpand
        '
        Me.tsmiExpand.Image = Global.WindowsApplication1.My.Resources.Resources.minus26
        Me.tsmiExpand.Name = "tsmiExpand"
        resources.ApplyResources(Me.tsmiExpand, "tsmiExpand")
        '
        'СвернутьВсеToolStripMenuItem
        '
        Me.СвернутьВсеToolStripMenuItem.Image = Global.WindowsApplication1.My.Resources.Resources.plus26
        Me.СвернутьВсеToolStripMenuItem.Name = "СвернутьВсеToolStripMenuItem"
        resources.ApplyResources(Me.СвернутьВсеToolStripMenuItem, "СвернутьВсеToolStripMenuItem")
        '
        'Button2
        '
        resources.ApplyResources(Me.Button2, "Button2")
        Me.Button2.Name = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.СервисToolStripMenuItem, Me.ПомощьToolStripMenuItem, Me.ПомощьToolStripMenuItem1})
        resources.ApplyResources(Me.MenuStrip1, "MenuStrip1")
        Me.MenuStrip1.Name = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenToolStripMenuItem, Me.tsmiLoad, Me.tsmiSave, Me.СохранитьКакToolStripMenuItem, Me.СвойстваКвестаToolStripMenuItem, Me.ToolStripSeparator8, Me.tsmiRun, Me.ToolStripSeparator9, Me.ПоследниеФайлыToolStripMenuItem, Me.ВыходИзРедактораToolStripMenuItem, Me.ВыходИзПрограммыToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        resources.ApplyResources(Me.FileToolStripMenuItem, "FileToolStripMenuItem")
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        resources.ApplyResources(Me.OpenToolStripMenuItem, "OpenToolStripMenuItem")
        '
        'tsmiLoad
        '
        Me.tsmiLoad.Name = "tsmiLoad"
        resources.ApplyResources(Me.tsmiLoad, "tsmiLoad")
        '
        'tsmiSave
        '
        Me.tsmiSave.Name = "tsmiSave"
        resources.ApplyResources(Me.tsmiSave, "tsmiSave")
        '
        'СохранитьКакToolStripMenuItem
        '
        Me.СохранитьКакToolStripMenuItem.Name = "СохранитьКакToolStripMenuItem"
        resources.ApplyResources(Me.СохранитьКакToolStripMenuItem, "СохранитьКакToolStripMenuItem")
        '
        'СвойстваКвестаToolStripMenuItem
        '
        Me.СвойстваКвестаToolStripMenuItem.Name = "СвойстваКвестаToolStripMenuItem"
        resources.ApplyResources(Me.СвойстваКвестаToolStripMenuItem, "СвойстваКвестаToolStripMenuItem")
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        resources.ApplyResources(Me.ToolStripSeparator8, "ToolStripSeparator8")
        '
        'tsmiRun
        '
        Me.tsmiRun.Name = "tsmiRun"
        resources.ApplyResources(Me.tsmiRun, "tsmiRun")
        '
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        resources.ApplyResources(Me.ToolStripSeparator9, "ToolStripSeparator9")
        '
        'ПоследниеФайлыToolStripMenuItem
        '
        Me.ПоследниеФайлыToolStripMenuItem.Name = "ПоследниеФайлыToolStripMenuItem"
        resources.ApplyResources(Me.ПоследниеФайлыToolStripMenuItem, "ПоследниеФайлыToolStripMenuItem")
        '
        'ВыходИзРедактораToolStripMenuItem
        '
        Me.ВыходИзРедактораToolStripMenuItem.Name = "ВыходИзРедактораToolStripMenuItem"
        resources.ApplyResources(Me.ВыходИзРедактораToolStripMenuItem, "ВыходИзРедактораToolStripMenuItem")
        '
        'ВыходИзПрограммыToolStripMenuItem
        '
        Me.ВыходИзПрограммыToolStripMenuItem.Name = "ВыходИзПрограммыToolStripMenuItem"
        resources.ApplyResources(Me.ВыходИзПрограммыToolStripMenuItem, "ВыходИзПрограммыToolStripMenuItem")
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiUndo, Me.tsmiRedo, Me.tsmiCopy, Me.tsmiCut, Me.tsmiPaste, Me.ToolStripSeparator5, Me.tsmiSelLine, Me.tsmiComment, Me.tsmiUncomment, Me.tsmiKeyboard, Me.tsmiWrap, Me.tsmiDefVars, Me.tsmiExec, Me.ToolStripSeparator2, Me.ФорматированиеТекстаToolStripMenuItem, Me.tsmiSpecialChars, Me.tsmiImage, Me.tsmiTable, Me.tsmiList, Me.tsmiMarquee, Me.tsmiIFrame})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        resources.ApplyResources(Me.EditToolStripMenuItem, "EditToolStripMenuItem")
        '
        'tsmiUndo
        '
        Me.tsmiUndo.Image = Global.WindowsApplication1.My.Resources.Resources.undo
        resources.ApplyResources(Me.tsmiUndo, "tsmiUndo")
        Me.tsmiUndo.Name = "tsmiUndo"
        '
        'tsmiRedo
        '
        Me.tsmiRedo.Image = Global.WindowsApplication1.My.Resources.Resources.redo
        resources.ApplyResources(Me.tsmiRedo, "tsmiRedo")
        Me.tsmiRedo.Name = "tsmiRedo"
        '
        'tsmiCopy
        '
        Me.tsmiCopy.Image = Global.WindowsApplication1.My.Resources.Resources.copy
        resources.ApplyResources(Me.tsmiCopy, "tsmiCopy")
        Me.tsmiCopy.Name = "tsmiCopy"
        '
        'tsmiCut
        '
        Me.tsmiCut.Image = Global.WindowsApplication1.My.Resources.Resources.cut
        resources.ApplyResources(Me.tsmiCut, "tsmiCut")
        Me.tsmiCut.Name = "tsmiCut"
        '
        'tsmiPaste
        '
        Me.tsmiPaste.Image = Global.WindowsApplication1.My.Resources.Resources.paste
        resources.ApplyResources(Me.tsmiPaste, "tsmiPaste")
        Me.tsmiPaste.Name = "tsmiPaste"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        resources.ApplyResources(Me.ToolStripSeparator5, "ToolStripSeparator5")
        '
        'tsmiSelLine
        '
        resources.ApplyResources(Me.tsmiSelLine, "tsmiSelLine")
        Me.tsmiSelLine.Name = "tsmiSelLine"
        '
        'tsmiComment
        '
        Me.tsmiComment.Image = Global.WindowsApplication1.My.Resources.Resources.comment
        resources.ApplyResources(Me.tsmiComment, "tsmiComment")
        Me.tsmiComment.Name = "tsmiComment"
        '
        'tsmiUncomment
        '
        Me.tsmiUncomment.Image = Global.WindowsApplication1.My.Resources.Resources.uncomment
        resources.ApplyResources(Me.tsmiUncomment, "tsmiUncomment")
        Me.tsmiUncomment.Name = "tsmiUncomment"
        '
        'tsmiKeyboard
        '
        Me.tsmiKeyboard.Image = Global.WindowsApplication1.My.Resources.Resources.abc
        resources.ApplyResources(Me.tsmiKeyboard, "tsmiKeyboard")
        Me.tsmiKeyboard.Name = "tsmiKeyboard"
        '
        'tsmiWrap
        '
        resources.ApplyResources(Me.tsmiWrap, "tsmiWrap")
        Me.tsmiWrap.Name = "tsmiWrap"
        '
        'tsmiDefVars
        '
        Me.tsmiDefVars.Name = "tsmiDefVars"
        resources.ApplyResources(Me.tsmiDefVars, "tsmiDefVars")
        '
        'tsmiExec
        '
        Me.tsmiExec.Image = Global.WindowsApplication1.My.Resources.Resources.exec
        resources.ApplyResources(Me.tsmiExec, "tsmiExec")
        Me.tsmiExec.Name = "tsmiExec"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        resources.ApplyResources(Me.ToolStripSeparator2, "ToolStripSeparator2")
        '
        'ФорматированиеТекстаToolStripMenuItem
        '
        Me.ФорматированиеТекстаToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiB, Me.tsmiI, Me.tsmiU, Me.tsmiSup, Me.tsmiSub, Me.ToolStripSeparator3, Me.ToolStripMenuItem3, Me.tsmiP, Me.tsmiH1, Me.tsmiH2, Me.tsmiH3, Me.tsmiSpan, Me.tsmiFontSize, Me.tsmiColor})
        resources.ApplyResources(Me.ФорматированиеТекстаToolStripMenuItem, "ФорматированиеТекстаToolStripMenuItem")
        Me.ФорматированиеТекстаToolStripMenuItem.Name = "ФорматированиеТекстаToolStripMenuItem"
        '
        'tsmiB
        '
        Me.tsmiB.Image = Global.WindowsApplication1.My.Resources.Resources.b
        resources.ApplyResources(Me.tsmiB, "tsmiB")
        Me.tsmiB.Name = "tsmiB"
        '
        'tsmiI
        '
        Me.tsmiI.Image = Global.WindowsApplication1.My.Resources.Resources.i
        resources.ApplyResources(Me.tsmiI, "tsmiI")
        Me.tsmiI.Name = "tsmiI"
        '
        'tsmiU
        '
        Me.tsmiU.Image = Global.WindowsApplication1.My.Resources.Resources.u
        resources.ApplyResources(Me.tsmiU, "tsmiU")
        Me.tsmiU.Name = "tsmiU"
        '
        'tsmiSup
        '
        Me.tsmiSup.Image = Global.WindowsApplication1.My.Resources.Resources.sup
        resources.ApplyResources(Me.tsmiSup, "tsmiSup")
        Me.tsmiSup.Name = "tsmiSup"
        '
        'tsmiSub
        '
        Me.tsmiSub.Image = Global.WindowsApplication1.My.Resources.Resources.sub2
        resources.ApplyResources(Me.tsmiSub, "tsmiSub")
        Me.tsmiSub.Name = "tsmiSub"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        resources.ApplyResources(Me.ToolStripSeparator3, "ToolStripSeparator3")
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiLeft, Me.tsmiCenter, Me.tsmiRight, Me.tsmiJustify})
        resources.ApplyResources(Me.ToolStripMenuItem3, "ToolStripMenuItem3")
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        '
        'tsmiLeft
        '
        Me.tsmiLeft.Image = Global.WindowsApplication1.My.Resources.Resources.left
        resources.ApplyResources(Me.tsmiLeft, "tsmiLeft")
        Me.tsmiLeft.Name = "tsmiLeft"
        '
        'tsmiCenter
        '
        Me.tsmiCenter.Image = Global.WindowsApplication1.My.Resources.Resources.center
        resources.ApplyResources(Me.tsmiCenter, "tsmiCenter")
        Me.tsmiCenter.Name = "tsmiCenter"
        '
        'tsmiRight
        '
        Me.tsmiRight.Image = Global.WindowsApplication1.My.Resources.Resources.right
        resources.ApplyResources(Me.tsmiRight, "tsmiRight")
        Me.tsmiRight.Name = "tsmiRight"
        '
        'tsmiJustify
        '
        Me.tsmiJustify.Image = Global.WindowsApplication1.My.Resources.Resources.just
        resources.ApplyResources(Me.tsmiJustify, "tsmiJustify")
        Me.tsmiJustify.Name = "tsmiJustify"
        '
        'tsmiP
        '
        Me.tsmiP.Image = Global.WindowsApplication1.My.Resources.Resources.p
        resources.ApplyResources(Me.tsmiP, "tsmiP")
        Me.tsmiP.Name = "tsmiP"
        Me.tsmiP.Tag = "<P>"
        '
        'tsmiH1
        '
        Me.tsmiH1.Image = Global.WindowsApplication1.My.Resources.Resources.h1
        resources.ApplyResources(Me.tsmiH1, "tsmiH1")
        Me.tsmiH1.Name = "tsmiH1"
        '
        'tsmiH2
        '
        Me.tsmiH2.Image = Global.WindowsApplication1.My.Resources.Resources.h2
        resources.ApplyResources(Me.tsmiH2, "tsmiH2")
        Me.tsmiH2.Name = "tsmiH2"
        '
        'tsmiH3
        '
        Me.tsmiH3.Image = Global.WindowsApplication1.My.Resources.Resources.h3
        resources.ApplyResources(Me.tsmiH3, "tsmiH3")
        Me.tsmiH3.Name = "tsmiH3"
        '
        'tsmiSpan
        '
        Me.tsmiSpan.Image = Global.WindowsApplication1.My.Resources.Resources.span
        resources.ApplyResources(Me.tsmiSpan, "tsmiSpan")
        Me.tsmiSpan.Name = "tsmiSpan"
        '
        'tsmiFontSize
        '
        Me.tsmiFontSize.Image = Global.WindowsApplication1.My.Resources.Resources.size
        resources.ApplyResources(Me.tsmiFontSize, "tsmiFontSize")
        Me.tsmiFontSize.Name = "tsmiFontSize"
        '
        'tsmiColor
        '
        Me.tsmiColor.Image = Global.WindowsApplication1.My.Resources.Resources.color
        resources.ApplyResources(Me.tsmiColor, "tsmiColor")
        Me.tsmiColor.Name = "tsmiColor"
        '
        'tsmiSpecialChars
        '
        Me.tsmiSpecialChars.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiHR, Me.tsmiBR, Me.tsmiSpace, Me.tsmiDash, Me.tsmiQuotes, Me.tsmiEuro, Me.tsmiSelector, Me.ToolStripSeparator4, Me.tsmiChars})
        resources.ApplyResources(Me.tsmiSpecialChars, "tsmiSpecialChars")
        Me.tsmiSpecialChars.Name = "tsmiSpecialChars"
        '
        'tsmiHR
        '
        Me.tsmiHR.Image = Global.WindowsApplication1.My.Resources.Resources.hr
        resources.ApplyResources(Me.tsmiHR, "tsmiHR")
        Me.tsmiHR.Name = "tsmiHR"
        '
        'tsmiBR
        '
        Me.tsmiBR.Image = Global.WindowsApplication1.My.Resources.Resources.br
        resources.ApplyResources(Me.tsmiBR, "tsmiBR")
        Me.tsmiBR.Name = "tsmiBR"
        '
        'tsmiSpace
        '
        Me.tsmiSpace.Image = Global.WindowsApplication1.My.Resources.Resources.space
        resources.ApplyResources(Me.tsmiSpace, "tsmiSpace")
        Me.tsmiSpace.Name = "tsmiSpace"
        '
        'tsmiDash
        '
        Me.tsmiDash.Image = Global.WindowsApplication1.My.Resources.Resources.dash
        resources.ApplyResources(Me.tsmiDash, "tsmiDash")
        Me.tsmiDash.Name = "tsmiDash"
        '
        'tsmiQuotes
        '
        Me.tsmiQuotes.Image = Global.WindowsApplication1.My.Resources.Resources.qBrackets
        resources.ApplyResources(Me.tsmiQuotes, "tsmiQuotes")
        Me.tsmiQuotes.Name = "tsmiQuotes"
        '
        'tsmiEuro
        '
        Me.tsmiEuro.Image = Global.WindowsApplication1.My.Resources.Resources.e
        resources.ApplyResources(Me.tsmiEuro, "tsmiEuro")
        Me.tsmiEuro.Name = "tsmiEuro"
        '
        'tsmiSelector
        '
        Me.tsmiSelector.Image = Global.WindowsApplication1.My.Resources.Resources.selector
        Me.tsmiSelector.Name = "tsmiSelector"
        resources.ApplyResources(Me.tsmiSelector, "tsmiSelector")
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        resources.ApplyResources(Me.ToolStripSeparator4, "ToolStripSeparator4")
        '
        'tsmiChars
        '
        Me.tsmiChars.Image = Global.WindowsApplication1.My.Resources.Resources.chrOther
        resources.ApplyResources(Me.tsmiChars, "tsmiChars")
        Me.tsmiChars.Name = "tsmiChars"
        '
        'tsmiImage
        '
        Me.tsmiImage.Image = Global.WindowsApplication1.My.Resources.Resources.img
        resources.ApplyResources(Me.tsmiImage, "tsmiImage")
        Me.tsmiImage.Name = "tsmiImage"
        '
        'tsmiTable
        '
        Me.tsmiTable.Image = Global.WindowsApplication1.My.Resources.Resources.table
        resources.ApplyResources(Me.tsmiTable, "tsmiTable")
        Me.tsmiTable.Name = "tsmiTable"
        '
        'tsmiList
        '
        Me.tsmiList.Image = Global.WindowsApplication1.My.Resources.Resources.list
        resources.ApplyResources(Me.tsmiList, "tsmiList")
        Me.tsmiList.Name = "tsmiList"
        '
        'tsmiMarquee
        '
        Me.tsmiMarquee.Image = Global.WindowsApplication1.My.Resources.Resources.marq
        resources.ApplyResources(Me.tsmiMarquee, "tsmiMarquee")
        Me.tsmiMarquee.Name = "tsmiMarquee"
        '
        'tsmiIFrame
        '
        Me.tsmiIFrame.Image = Global.WindowsApplication1.My.Resources.Resources.url
        resources.ApplyResources(Me.tsmiIFrame, "tsmiIFrame")
        Me.tsmiIFrame.Name = "tsmiIFrame"
        '
        'СервисToolStripMenuItem
        '
        Me.СервисToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiFind, Me.tsmiFindNext, Me.tsmiFindGlobal, Me.ToolStripSeparator6, Me.НастройкиПрограммыToolStripMenuItem, Me.tsmiExecuteScript, Me.tsmiCopyAsHTML, Me.ToolStripSeparator7, Me.tsmiClassEditor, Me.tsmiRestoreElement, Me.tsmigHighLight})
        Me.СервисToolStripMenuItem.Name = "СервисToolStripMenuItem"
        resources.ApplyResources(Me.СервисToolStripMenuItem, "СервисToolStripMenuItem")
        '
        'tsmiFind
        '
        Me.tsmiFind.Image = Global.WindowsApplication1.My.Resources.Resources.find
        resources.ApplyResources(Me.tsmiFind, "tsmiFind")
        Me.tsmiFind.Name = "tsmiFind"
        '
        'tsmiFindNext
        '
        resources.ApplyResources(Me.tsmiFindNext, "tsmiFindNext")
        Me.tsmiFindNext.Name = "tsmiFindNext"
        '
        'tsmiFindGlobal
        '
        Me.tsmiFindGlobal.Image = Global.WindowsApplication1.My.Resources.Resources.radar
        resources.ApplyResources(Me.tsmiFindGlobal, "tsmiFindGlobal")
        Me.tsmiFindGlobal.Name = "tsmiFindGlobal"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        resources.ApplyResources(Me.ToolStripSeparator6, "ToolStripSeparator6")
        '
        'НастройкиПрограммыToolStripMenuItem
        '
        resources.ApplyResources(Me.НастройкиПрограммыToolStripMenuItem, "НастройкиПрограммыToolStripMenuItem")
        Me.НастройкиПрограммыToolStripMenuItem.Name = "НастройкиПрограммыToolStripMenuItem"
        '
        'tsmiExecuteScript
        '
        resources.ApplyResources(Me.tsmiExecuteScript, "tsmiExecuteScript")
        Me.tsmiExecuteScript.Name = "tsmiExecuteScript"
        '
        'tsmiCopyAsHTML
        '
        resources.ApplyResources(Me.tsmiCopyAsHTML, "tsmiCopyAsHTML")
        Me.tsmiCopyAsHTML.Name = "tsmiCopyAsHTML"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        resources.ApplyResources(Me.ToolStripSeparator7, "ToolStripSeparator7")
        '
        'tsmiClassEditor
        '
        Me.tsmiClassEditor.Name = "tsmiClassEditor"
        resources.ApplyResources(Me.tsmiClassEditor, "tsmiClassEditor")
        '
        'tsmiRestoreElement
        '
        Me.tsmiRestoreElement.Name = "tsmiRestoreElement"
        resources.ApplyResources(Me.tsmiRestoreElement, "tsmiRestoreElement")
        '
        'tsmigHighLight
        '
        Me.tsmigHighLight.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiHighLightFull, Me.tsmiHighLightDesignate, Me.tsmiHighLightNo})
        Me.tsmigHighLight.Image = Global.WindowsApplication1.My.Resources.Resources.highlight
        resources.ApplyResources(Me.tsmigHighLight, "tsmigHighLight")
        Me.tsmigHighLight.Name = "tsmigHighLight"
        '
        'tsmiHighLightFull
        '
        Me.tsmiHighLightFull.Image = Global.WindowsApplication1.My.Resources.Resources.highlight
        resources.ApplyResources(Me.tsmiHighLightFull, "tsmiHighLightFull")
        Me.tsmiHighLightFull.Name = "tsmiHighLightFull"
        '
        'tsmiHighLightDesignate
        '
        Me.tsmiHighLightDesignate.Checked = True
        Me.tsmiHighLightDesignate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.tsmiHighLightDesignate.Image = Global.WindowsApplication1.My.Resources.Resources.highlight2
        resources.ApplyResources(Me.tsmiHighLightDesignate, "tsmiHighLightDesignate")
        Me.tsmiHighLightDesignate.Name = "tsmiHighLightDesignate"
        '
        'tsmiHighLightNo
        '
        Me.tsmiHighLightNo.Image = Global.WindowsApplication1.My.Resources.Resources.highlight3
        resources.ApplyResources(Me.tsmiHighLightNo, "tsmiHighLightNo")
        Me.tsmiHighLightNo.Name = "tsmiHighLightNo"
        '
        'ПомощьToolStripMenuItem
        '
        Me.ПомощьToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiShowAdditionalClasses})
        Me.ПомощьToolStripMenuItem.Name = "ПомощьToolStripMenuItem"
        resources.ApplyResources(Me.ПомощьToolStripMenuItem, "ПомощьToolStripMenuItem")
        '
        'tsmiShowAdditionalClasses
        '
        Me.tsmiShowAdditionalClasses.Name = "tsmiShowAdditionalClasses"
        resources.ApplyResources(Me.tsmiShowAdditionalClasses, "tsmiShowAdditionalClasses")
        '
        'ПомощьToolStripMenuItem1
        '
        Me.ПомощьToolStripMenuItem1.Name = "ПомощьToolStripMenuItem1"
        resources.ApplyResources(Me.ПомощьToolStripMenuItem1, "ПомощьToolStripMenuItem1")
        '
        'ToolStripPanels
        '
        Me.ToolStripPanels.ImageScalingSize = New System.Drawing.Size(32, 32)
        resources.ApplyResources(Me.ToolStripPanels, "ToolStripPanels")
        Me.ToolStripPanels.Name = "ToolStripPanels"
        '
        'ToolStripCodeBox
        '
        resources.ApplyResources(Me.ToolStripCodeBox, "ToolStripCodeBox")
        Me.ToolStripCodeBox.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbDefVars, Me.tsbUndo, Me.tsbRedo, Me.tsbWrap, Me.tsbComment, Me.tsbUncomment, Me.ToolStripSeparator10, Me.tsbBR, Me.tsbHR, Me.tsbB, Me.tsbI, Me.tsbU, Me.tsbSup, Me.tsbSub, Me.tsbLeft, Me.tsbCenter, Me.tsbRight, Me.tsbJustify, Me.tsbTable, Me.tsbColor, Me.tsbFontSize, Me.tsbImage, Me.tsbMarquee, Me.tsbList, Me.tsbSpan, Me.tsbP, Me.tsbH1, Me.tsbH2, Me.tsbH3, Me.tsbExec, Me.ToolStripSeparator11, Me.tsbSeek, Me.ToolStripSeparator12, Me.tsbSelector, Me.tsbSpace, Me.tsbDash, Me.tsbQuotes, Me.tsbChars, Me.ToolStripSeparator17, Me.tsbgHighLight, Me.tsbRestoreInitial, Me.tsbFullScreen})
        Me.ToolStripCodeBox.Name = "ToolStripCodeBox"
        '
        'tsbDefVars
        '
        Me.tsbDefVars.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbDefVars.Image = Global.WindowsApplication1.My.Resources.Resources.defVars
        resources.ApplyResources(Me.tsbDefVars, "tsbDefVars")
        Me.tsbDefVars.Name = "tsbDefVars"
        '
        'tsbUndo
        '
        Me.tsbUndo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbUndo.Image = Global.WindowsApplication1.My.Resources.Resources.undo
        resources.ApplyResources(Me.tsbUndo, "tsbUndo")
        Me.tsbUndo.Name = "tsbUndo"
        '
        'tsbRedo
        '
        Me.tsbRedo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbRedo.Image = Global.WindowsApplication1.My.Resources.Resources.redo
        resources.ApplyResources(Me.tsbRedo, "tsbRedo")
        Me.tsbRedo.Name = "tsbRedo"
        '
        'tsbWrap
        '
        Me.tsbWrap.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        resources.ApplyResources(Me.tsbWrap, "tsbWrap")
        Me.tsbWrap.Name = "tsbWrap"
        '
        'tsbComment
        '
        Me.tsbComment.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbComment.Image = Global.WindowsApplication1.My.Resources.Resources.comment
        resources.ApplyResources(Me.tsbComment, "tsbComment")
        Me.tsbComment.Name = "tsbComment"
        '
        'tsbUncomment
        '
        Me.tsbUncomment.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbUncomment.Image = Global.WindowsApplication1.My.Resources.Resources.uncomment
        resources.ApplyResources(Me.tsbUncomment, "tsbUncomment")
        Me.tsbUncomment.Name = "tsbUncomment"
        '
        'ToolStripSeparator10
        '
        Me.ToolStripSeparator10.Name = "ToolStripSeparator10"
        resources.ApplyResources(Me.ToolStripSeparator10, "ToolStripSeparator10")
        '
        'tsbBR
        '
        Me.tsbBR.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbBR.Image = Global.WindowsApplication1.My.Resources.Resources.br
        resources.ApplyResources(Me.tsbBR, "tsbBR")
        Me.tsbBR.Name = "tsbBR"
        '
        'tsbHR
        '
        Me.tsbHR.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbHR.Image = Global.WindowsApplication1.My.Resources.Resources.hr
        resources.ApplyResources(Me.tsbHR, "tsbHR")
        Me.tsbHR.Name = "tsbHR"
        '
        'tsbB
        '
        Me.tsbB.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbB.Image = Global.WindowsApplication1.My.Resources.Resources.b
        resources.ApplyResources(Me.tsbB, "tsbB")
        Me.tsbB.Name = "tsbB"
        '
        'tsbI
        '
        Me.tsbI.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbI.Image = Global.WindowsApplication1.My.Resources.Resources.i
        resources.ApplyResources(Me.tsbI, "tsbI")
        Me.tsbI.Name = "tsbI"
        '
        'tsbU
        '
        Me.tsbU.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbU.Image = Global.WindowsApplication1.My.Resources.Resources.u
        resources.ApplyResources(Me.tsbU, "tsbU")
        Me.tsbU.Name = "tsbU"
        '
        'tsbSup
        '
        Me.tsbSup.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbSup.Image = Global.WindowsApplication1.My.Resources.Resources.sup
        resources.ApplyResources(Me.tsbSup, "tsbSup")
        Me.tsbSup.Name = "tsbSup"
        '
        'tsbSub
        '
        Me.tsbSub.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbSub.Image = Global.WindowsApplication1.My.Resources.Resources.sub2
        resources.ApplyResources(Me.tsbSub, "tsbSub")
        Me.tsbSub.Name = "tsbSub"
        '
        'tsbLeft
        '
        Me.tsbLeft.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbLeft.Image = Global.WindowsApplication1.My.Resources.Resources.left
        resources.ApplyResources(Me.tsbLeft, "tsbLeft")
        Me.tsbLeft.Name = "tsbLeft"
        '
        'tsbCenter
        '
        Me.tsbCenter.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbCenter.Image = Global.WindowsApplication1.My.Resources.Resources.center
        resources.ApplyResources(Me.tsbCenter, "tsbCenter")
        Me.tsbCenter.Name = "tsbCenter"
        '
        'tsbRight
        '
        Me.tsbRight.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbRight.Image = Global.WindowsApplication1.My.Resources.Resources.right
        resources.ApplyResources(Me.tsbRight, "tsbRight")
        Me.tsbRight.Name = "tsbRight"
        '
        'tsbJustify
        '
        Me.tsbJustify.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbJustify.Image = Global.WindowsApplication1.My.Resources.Resources.just
        resources.ApplyResources(Me.tsbJustify, "tsbJustify")
        Me.tsbJustify.Name = "tsbJustify"
        '
        'tsbTable
        '
        Me.tsbTable.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbTable.Image = Global.WindowsApplication1.My.Resources.Resources.table
        resources.ApplyResources(Me.tsbTable, "tsbTable")
        Me.tsbTable.Name = "tsbTable"
        '
        'tsbColor
        '
        Me.tsbColor.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbColor.Image = Global.WindowsApplication1.My.Resources.Resources.color
        resources.ApplyResources(Me.tsbColor, "tsbColor")
        Me.tsbColor.Name = "tsbColor"
        '
        'tsbFontSize
        '
        Me.tsbFontSize.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbFontSize.Image = Global.WindowsApplication1.My.Resources.Resources.size
        resources.ApplyResources(Me.tsbFontSize, "tsbFontSize")
        Me.tsbFontSize.Name = "tsbFontSize"
        '
        'tsbImage
        '
        Me.tsbImage.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbImage.Image = Global.WindowsApplication1.My.Resources.Resources.img
        resources.ApplyResources(Me.tsbImage, "tsbImage")
        Me.tsbImage.Name = "tsbImage"
        '
        'tsbMarquee
        '
        Me.tsbMarquee.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbMarquee.Image = Global.WindowsApplication1.My.Resources.Resources.marq
        resources.ApplyResources(Me.tsbMarquee, "tsbMarquee")
        Me.tsbMarquee.Name = "tsbMarquee"
        '
        'tsbList
        '
        Me.tsbList.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbList.Image = Global.WindowsApplication1.My.Resources.Resources.list
        resources.ApplyResources(Me.tsbList, "tsbList")
        Me.tsbList.Name = "tsbList"
        '
        'tsbSpan
        '
        Me.tsbSpan.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbSpan.Image = Global.WindowsApplication1.My.Resources.Resources.span
        resources.ApplyResources(Me.tsbSpan, "tsbSpan")
        Me.tsbSpan.Name = "tsbSpan"
        '
        'tsbP
        '
        Me.tsbP.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbP.Image = Global.WindowsApplication1.My.Resources.Resources.p
        resources.ApplyResources(Me.tsbP, "tsbP")
        Me.tsbP.Name = "tsbP"
        '
        'tsbH1
        '
        Me.tsbH1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbH1.Image = Global.WindowsApplication1.My.Resources.Resources.h1
        resources.ApplyResources(Me.tsbH1, "tsbH1")
        Me.tsbH1.Name = "tsbH1"
        '
        'tsbH2
        '
        Me.tsbH2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbH2.Image = Global.WindowsApplication1.My.Resources.Resources.h2
        resources.ApplyResources(Me.tsbH2, "tsbH2")
        Me.tsbH2.Name = "tsbH2"
        '
        'tsbH3
        '
        Me.tsbH3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbH3.Image = Global.WindowsApplication1.My.Resources.Resources.h3
        resources.ApplyResources(Me.tsbH3, "tsbH3")
        Me.tsbH3.Name = "tsbH3"
        '
        'tsbExec
        '
        Me.tsbExec.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbExec.Image = Global.WindowsApplication1.My.Resources.Resources.exec
        resources.ApplyResources(Me.tsbExec, "tsbExec")
        Me.tsbExec.Name = "tsbExec"
        '
        'ToolStripSeparator11
        '
        Me.ToolStripSeparator11.Name = "ToolStripSeparator11"
        resources.ApplyResources(Me.ToolStripSeparator11, "ToolStripSeparator11")
        '
        'tsbSeek
        '
        Me.tsbSeek.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbSeek.Image = Global.WindowsApplication1.My.Resources.Resources.find
        resources.ApplyResources(Me.tsbSeek, "tsbSeek")
        Me.tsbSeek.Name = "tsbSeek"
        '
        'ToolStripSeparator12
        '
        Me.ToolStripSeparator12.Name = "ToolStripSeparator12"
        resources.ApplyResources(Me.ToolStripSeparator12, "ToolStripSeparator12")
        '
        'tsbSelector
        '
        Me.tsbSelector.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbSelector.Image = Global.WindowsApplication1.My.Resources.Resources.selector
        resources.ApplyResources(Me.tsbSelector, "tsbSelector")
        Me.tsbSelector.Name = "tsbSelector"
        '
        'tsbSpace
        '
        Me.tsbSpace.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbSpace.Image = Global.WindowsApplication1.My.Resources.Resources.space
        resources.ApplyResources(Me.tsbSpace, "tsbSpace")
        Me.tsbSpace.Name = "tsbSpace"
        '
        'tsbDash
        '
        Me.tsbDash.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbDash.Image = Global.WindowsApplication1.My.Resources.Resources.dash
        resources.ApplyResources(Me.tsbDash, "tsbDash")
        Me.tsbDash.Name = "tsbDash"
        '
        'tsbQuotes
        '
        Me.tsbQuotes.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbQuotes.Image = Global.WindowsApplication1.My.Resources.Resources.qBrackets
        resources.ApplyResources(Me.tsbQuotes, "tsbQuotes")
        Me.tsbQuotes.Name = "tsbQuotes"
        '
        'tsbChars
        '
        Me.tsbChars.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbChars.Image = Global.WindowsApplication1.My.Resources.Resources.chrOther
        resources.ApplyResources(Me.tsbChars, "tsbChars")
        Me.tsbChars.Name = "tsbChars"
        '
        'ToolStripSeparator17
        '
        Me.ToolStripSeparator17.Name = "ToolStripSeparator17"
        resources.ApplyResources(Me.ToolStripSeparator17, "ToolStripSeparator17")
        '
        'tsbgHighLight
        '
        Me.tsbgHighLight.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbgHighLight.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbHighLightFull, Me.tsbHighLightDesignate, Me.tsbHighLightNo})
        Me.tsbgHighLight.Image = Global.WindowsApplication1.My.Resources.Resources.highlight
        resources.ApplyResources(Me.tsbgHighLight, "tsbgHighLight")
        Me.tsbgHighLight.Name = "tsbgHighLight"
        '
        'tsbHighLightFull
        '
        Me.tsbHighLightFull.Image = Global.WindowsApplication1.My.Resources.Resources.highlight
        Me.tsbHighLightFull.Name = "tsbHighLightFull"
        resources.ApplyResources(Me.tsbHighLightFull, "tsbHighLightFull")
        '
        'tsbHighLightDesignate
        '
        Me.tsbHighLightDesignate.Checked = True
        Me.tsbHighLightDesignate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.tsbHighLightDesignate.Image = Global.WindowsApplication1.My.Resources.Resources.highlight2
        Me.tsbHighLightDesignate.Name = "tsbHighLightDesignate"
        resources.ApplyResources(Me.tsbHighLightDesignate, "tsbHighLightDesignate")
        '
        'tsbHighLightNo
        '
        Me.tsbHighLightNo.Image = Global.WindowsApplication1.My.Resources.Resources.highlight3
        Me.tsbHighLightNo.Name = "tsbHighLightNo"
        resources.ApplyResources(Me.tsbHighLightNo, "tsbHighLightNo")
        '
        'tsbRestoreInitial
        '
        Me.tsbRestoreInitial.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbRestoreInitial.Image = Global.WindowsApplication1.My.Resources.Resources.restoreInitial26
        resources.ApplyResources(Me.tsbRestoreInitial, "tsbRestoreInitial")
        Me.tsbRestoreInitial.Name = "tsbRestoreInitial"
        '
        'tsbFullScreen
        '
        Me.tsbFullScreen.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbFullScreen.Image = Global.WindowsApplication1.My.Resources.Resources.fullscreen
        resources.ApplyResources(Me.tsbFullScreen, "tsbFullScreen")
        Me.tsbFullScreen.Name = "tsbFullScreen"
        '
        'ToolStripMain
        '
        Me.ToolStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbQ, Me.tsbString, Me.tsbMath, Me.tsbDate, Me.tsbArr, Me.tsbFile, Me.tsbCode, Me.ToolStripSeparator1, Me.tsbL, Me.tsbObj, Me.tsbM, Me.tsbT, Me.ToolStripSeparator14, Me.tsbMap, Me.tsbMed, Me.tsbHer, Me.tsbMg, Me.tsbAb, Me.tsbArmy, Me.tsbBat, Me.tsddbWindows, Me.ToolStripSeparator15, Me.tsbFunc, Me.tsbVar, Me.ToolStripSeparator16, Me.tsbAddUserClass})
        resources.ApplyResources(Me.ToolStripMain, "ToolStripMain")
        Me.ToolStripMain.Name = "ToolStripMain"
        '
        'tsbQ
        '
        Me.tsbQ.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbQ.Image = Global.WindowsApplication1.My.Resources.Resources.Q
        resources.ApplyResources(Me.tsbQ, "tsbQ")
        Me.tsbQ.Name = "tsbQ"
        Me.tsbQ.Tag = "Q"
        '
        'tsbString
        '
        Me.tsbString.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbString.Image = Global.WindowsApplication1.My.Resources.Resources.S
        resources.ApplyResources(Me.tsbString, "tsbString")
        Me.tsbString.Name = "tsbString"
        Me.tsbString.Tag = "S"
        '
        'tsbMath
        '
        Me.tsbMath.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbMath.Image = Global.WindowsApplication1.My.Resources.Resources.Math
        resources.ApplyResources(Me.tsbMath, "tsbMath")
        Me.tsbMath.Name = "tsbMath"
        Me.tsbMath.Tag = "Math"
        '
        'tsbDate
        '
        Me.tsbDate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbDate.Image = Global.WindowsApplication1.My.Resources.Resources.D
        resources.ApplyResources(Me.tsbDate, "tsbDate")
        Me.tsbDate.Name = "tsbDate"
        Me.tsbDate.Tag = "D"
        '
        'tsbArr
        '
        Me.tsbArr.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbArr.Image = Global.WindowsApplication1.My.Resources.Resources.Arr
        resources.ApplyResources(Me.tsbArr, "tsbArr")
        Me.tsbArr.Name = "tsbArr"
        Me.tsbArr.Tag = "Arr"
        '
        'tsbFile
        '
        Me.tsbFile.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbFile.Image = Global.WindowsApplication1.My.Resources.Resources.File
        resources.ApplyResources(Me.tsbFile, "tsbFile")
        Me.tsbFile.Name = "tsbFile"
        Me.tsbFile.Tag = "File"
        '
        'tsbCode
        '
        Me.tsbCode.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbCode.Image = Global.WindowsApplication1.My.Resources.Resources.Code1
        resources.ApplyResources(Me.tsbCode, "tsbCode")
        Me.tsbCode.Name = "tsbCode"
        Me.tsbCode.Tag = "Code"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        resources.ApplyResources(Me.ToolStripSeparator1, "ToolStripSeparator1")
        '
        'tsbL
        '
        Me.tsbL.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbL.Image = Global.WindowsApplication1.My.Resources.Resources.Loc
        resources.ApplyResources(Me.tsbL, "tsbL")
        Me.tsbL.Name = "tsbL"
        Me.tsbL.Tag = "L"
        '
        'tsbObj
        '
        Me.tsbObj.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbObj.Image = Global.WindowsApplication1.My.Resources.Resources.Obj48
        resources.ApplyResources(Me.tsbObj, "tsbObj")
        Me.tsbObj.Name = "tsbObj"
        Me.tsbObj.Tag = "O"
        '
        'tsbM
        '
        Me.tsbM.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbM.Image = Global.WindowsApplication1.My.Resources.Resources.Menu
        resources.ApplyResources(Me.tsbM, "tsbM")
        Me.tsbM.Name = "tsbM"
        Me.tsbM.Tag = "M"
        '
        'tsbT
        '
        Me.tsbT.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbT.Image = Global.WindowsApplication1.My.Resources.Resources.Tim
        resources.ApplyResources(Me.tsbT, "tsbT")
        Me.tsbT.Name = "tsbT"
        Me.tsbT.Tag = "T"
        '
        'ToolStripSeparator14
        '
        Me.ToolStripSeparator14.Name = "ToolStripSeparator14"
        resources.ApplyResources(Me.ToolStripSeparator14, "ToolStripSeparator14")
        '
        'tsbMap
        '
        Me.tsbMap.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbMap.Image = Global.WindowsApplication1.My.Resources.Resources.Map
        resources.ApplyResources(Me.tsbMap, "tsbMap")
        Me.tsbMap.Name = "tsbMap"
        Me.tsbMap.Tag = "Map"
        '
        'tsbMed
        '
        Me.tsbMed.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbMed.Image = Global.WindowsApplication1.My.Resources.Resources.Med
        resources.ApplyResources(Me.tsbMed, "tsbMed")
        Me.tsbMed.Name = "tsbMed"
        Me.tsbMed.Tag = "Med"
        '
        'tsbHer
        '
        Me.tsbHer.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbHer.Image = Global.WindowsApplication1.My.Resources.Resources.Her
        resources.ApplyResources(Me.tsbHer, "tsbHer")
        Me.tsbHer.Name = "tsbHer"
        Me.tsbHer.Tag = "H"
        '
        'tsbMg
        '
        Me.tsbMg.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbMg.Image = Global.WindowsApplication1.My.Resources.Resources.Mg
        resources.ApplyResources(Me.tsbMg, "tsbMg")
        Me.tsbMg.Name = "tsbMg"
        Me.tsbMg.Tag = "Mg"
        '
        'tsbAb
        '
        Me.tsbAb.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbAb.Image = Global.WindowsApplication1.My.Resources.Resources.Ab
        resources.ApplyResources(Me.tsbAb, "tsbAb")
        Me.tsbAb.Name = "tsbAb"
        Me.tsbAb.Tag = "Ab"
        '
        'tsbArmy
        '
        Me.tsbArmy.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbArmy.Image = Global.WindowsApplication1.My.Resources.Resources.Army
        resources.ApplyResources(Me.tsbArmy, "tsbArmy")
        Me.tsbArmy.Name = "tsbArmy"
        Me.tsbArmy.Tag = "Army"
        '
        'tsbBat
        '
        Me.tsbBat.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbBat.Image = Global.WindowsApplication1.My.Resources.Resources.Bat
        resources.ApplyResources(Me.tsbBat, "tsbBat")
        Me.tsbBat.Name = "tsbBat"
        Me.tsbBat.Tag = "B"
        '
        'tsddbWindows
        '
        Me.tsddbWindows.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsddbWindows.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiCm, Me.tsmiLW, Me.tsmiDW, Me.tsmiOW, Me.tsmiAW, Me.tsmiMgW})
        Me.tsddbWindows.Image = Global.WindowsApplication1.My.Resources.Resources.Windows
        resources.ApplyResources(Me.tsddbWindows, "tsddbWindows")
        Me.tsddbWindows.Name = "tsddbWindows"
        '
        'tsmiCm
        '
        Me.tsmiCm.Image = Global.WindowsApplication1.My.Resources.Resources.Cm
        resources.ApplyResources(Me.tsmiCm, "tsmiCm")
        Me.tsmiCm.Name = "tsmiCm"
        Me.tsmiCm.Tag = "Cm"
        '
        'tsmiLW
        '
        Me.tsmiLW.Image = Global.WindowsApplication1.My.Resources.Resources.LW
        resources.ApplyResources(Me.tsmiLW, "tsmiLW")
        Me.tsmiLW.Name = "tsmiLW"
        Me.tsmiLW.Tag = "LW"
        '
        'tsmiDW
        '
        Me.tsmiDW.Image = Global.WindowsApplication1.My.Resources.Resources.DW
        resources.ApplyResources(Me.tsmiDW, "tsmiDW")
        Me.tsmiDW.Name = "tsmiDW"
        Me.tsmiDW.Tag = "DW"
        '
        'tsmiOW
        '
        Me.tsmiOW.Image = Global.WindowsApplication1.My.Resources.Resources.OW
        resources.ApplyResources(Me.tsmiOW, "tsmiOW")
        Me.tsmiOW.Name = "tsmiOW"
        Me.tsmiOW.Tag = "OW"
        '
        'tsmiAW
        '
        Me.tsmiAW.Image = Global.WindowsApplication1.My.Resources.Resources.AW
        resources.ApplyResources(Me.tsmiAW, "tsmiAW")
        Me.tsmiAW.Name = "tsmiAW"
        Me.tsmiAW.Tag = "AW"
        '
        'tsmiMgW
        '
        Me.tsmiMgW.Image = Global.WindowsApplication1.My.Resources.Resources.MgW
        resources.ApplyResources(Me.tsmiMgW, "tsmiMgW")
        Me.tsmiMgW.Name = "tsmiMgW"
        Me.tsmiMgW.Tag = "MgW"
        '
        'ToolStripSeparator15
        '
        Me.ToolStripSeparator15.Name = "ToolStripSeparator15"
        resources.ApplyResources(Me.ToolStripSeparator15, "ToolStripSeparator15")
        '
        'tsbFunc
        '
        Me.tsbFunc.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbFunc.Image = Global.WindowsApplication1.My.Resources.Resources.F
        resources.ApplyResources(Me.tsbFunc, "tsbFunc")
        Me.tsbFunc.Name = "tsbFunc"
        Me.tsbFunc.Tag = "Function"
        '
        'tsbVar
        '
        Me.tsbVar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbVar.Image = Global.WindowsApplication1.My.Resources.Resources.V
        resources.ApplyResources(Me.tsbVar, "tsbVar")
        Me.tsbVar.Name = "tsbVar"
        Me.tsbVar.Tag = "Variable"
        '
        'ToolStripSeparator16
        '
        Me.ToolStripSeparator16.Name = "ToolStripSeparator16"
        resources.ApplyResources(Me.ToolStripSeparator16, "ToolStripSeparator16")
        '
        'tsbAddUserClass
        '
        Me.tsbAddUserClass.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbAddUserClass.Image = Global.WindowsApplication1.My.Resources.Resources.user_classes
        resources.ApplyResources(Me.tsbAddUserClass, "tsbAddUserClass")
        Me.tsbAddUserClass.Name = "tsbAddUserClass"
        Me.tsbAddUserClass.Tag = ""
        '
        'imgLstGroupIcons
        '
        Me.imgLstGroupIcons.ColorDepth = System.Windows.Forms.ColorDepth.Depth24Bit
        resources.ApplyResources(Me.imgLstGroupIcons, "imgLstGroupIcons")
        Me.imgLstGroupIcons.TransparentColor = System.Drawing.Color.Transparent
        '
        'fsWatcher
        '
        Me.fsWatcher.EnableRaisingEvents = True
        Me.fsWatcher.IncludeSubdirectories = True
        Me.fsWatcher.SynchronizingObject = Me
        '
        'frmMainEditor
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ToolStripCodeBox)
        Me.Controls.Add(Me.ToolStripPanels)
        Me.Controls.Add(Me.ToolStripMain)
        Me.Controls.Add(Me.SplitOuter)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "frmMainEditor"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SplitOuter.Panel1.ResumeLayout(False)
        Me.SplitOuter.Panel2.ResumeLayout(False)
        CType(Me.SplitOuter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitOuter.ResumeLayout(False)
        Me.pnlSeek.ResumeLayout(False)
        Me.pnlSeek.PerformLayout()
        Me.SplitInner.Panel2.ResumeLayout(False)
        CType(Me.SplitInner, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitInner.ResumeLayout(False)
        Me.pnlVariables.ResumeLayout(False)
        Me.pnlVariables.PerformLayout()
        CType(Me.dgwVariables, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cmnuMainTree.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ToolStripCodeBox.ResumeLayout(False)
        Me.ToolStripCodeBox.PerformLayout()
        Me.ToolStripMain.ResumeLayout(False)
        Me.ToolStripMain.PerformLayout()
        CType(Me.fsWatcher, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmnuMainTree As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiAddGroup As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiAddElement As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SplitOuter As System.Windows.Forms.SplitContainer
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripPanels As System.Windows.Forms.ToolStrip
    Friend WithEvents SplitInner As System.Windows.Forms.SplitContainer
    Friend WithEvents ttMain As System.Windows.Forms.ToolTip
    Friend WithEvents tsmiRemoveElement As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiDuplicate As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ToolStripCodeBox As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbBR As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbHR As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbB As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbI As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbU As System.Windows.Forms.ToolStripButton
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbSup As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbSub As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsmiSpecialChars As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiDash As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiQuotes As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSpace As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiEuro As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiImage As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiMarquee As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiIFrame As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiChars As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiUndo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiRedo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiCut As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiPaste As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiSelLine As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiUncomment As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiComment As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiKeyboard As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiWrap As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ФорматированиеТекстаToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiB As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiI As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiU As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSup As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSub As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItem3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLeft As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiRight As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiCenter As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiJustify As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiP As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiH1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiH2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiH3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSpan As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiFontSize As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiColor As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiHR As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiBR As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiTable As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiList As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents СервисToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiFind As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiFindNext As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiFindGlobal As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents НастройкиПрограммыToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiExecuteScript As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiCopyAsHTML As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ПомощьToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ПомощьToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator7 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiClassEditor As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLoad As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents СохранитьКакToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents СвойстваКвестаToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator8 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiRun As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ПоследниеФайлыToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВыходИзРедактораToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВыходИзПрограммыToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbUndo As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbRestoreInitial As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbWrap As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbComment As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbUncomment As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator10 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbLeft As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbCenter As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbRight As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbJustify As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbTable As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbColor As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbFontSize As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbImage As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbMarquee As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbSpan As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbList As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbP As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbH1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbH2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbH3 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbSeek As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator12 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbSpace As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbDash As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbQuotes As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbExec As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbChars As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsmiExcludeFromGroup As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator13 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiExpand As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents СвернутьВсеToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents imgLstGroupIcons As System.Windows.Forms.ImageList
    Friend WithEvents ToolStripMain As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbQ As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbL As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbObj As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbM As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbT As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator14 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbMap As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbMed As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbHer As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbMg As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbAb As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbBat As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsddbWindows As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsmiCm As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLW As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiDW As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiOW As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiAW As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiMgW As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents fsWatcher As System.IO.FileSystemWatcher
    Friend WithEvents ToolStripSeparator15 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbVar As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbFunc As System.Windows.Forms.ToolStripButton
    Friend WithEvents pnlVariables As System.Windows.Forms.Panel
    Friend WithEvents txtVariableDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblVariableDescription As System.Windows.Forms.Label
    Friend WithEvents dgwVariables As System.Windows.Forms.DataGridView
    Friend WithEvents colSignature As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents lblVariableError As System.Windows.Forms.Label
    Friend WithEvents ToolStripSeparator16 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbAddUserClass As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbDefVars As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsmiDefVars As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiRestoreElement As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbCode As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbString As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbMath As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbDate As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbArr As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbFile As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsmiShowAdditionalClasses As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbgHighLight As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsbHighLightDesignate As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbHighLightFull As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbHighLightNo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmigHighLight As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiHighLightFull As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiHighLightDesignate As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiHighLightNo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents pnlSeek As System.Windows.Forms.Panel
    Friend WithEvents chkSeekWholeWord As System.Windows.Forms.CheckBox
    Friend WithEvents chkSeekCase As System.Windows.Forms.CheckBox
    Friend WithEvents cmbSeek As System.Windows.Forms.ComboBox
    Friend WithEvents btnFindClobal As System.Windows.Forms.Button
    Friend WithEvents btnSeekForward As System.Windows.Forms.Button
    Friend WithEvents btnSeekBackward As System.Windows.Forms.Button
    Friend WithEvents tsmiExec As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbRedo As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator17 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiSelector As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbSelector As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbFullScreen As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbArmy As System.Windows.Forms.ToolStripButton

End Class
