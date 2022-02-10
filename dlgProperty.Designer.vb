<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgProperty
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
        Me.btnPropGenerateHelp = New System.Windows.Forms.Button()
        Me.txtPropName = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.pnlProperty = New System.Windows.Forms.Panel()
        Me.btnPropParamShow = New System.Windows.Forms.Button()
        Me.lstPropElementClass = New System.Windows.Forms.ListBox()
        Me.lblPropElementClass = New System.Windows.Forms.Label()
        Me.cmbPropHidden = New System.Windows.Forms.ComboBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtPropEditorName = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
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
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.chkDetails = New System.Windows.Forms.CheckBox()
        Me.ofd = New System.Windows.Forms.OpenFileDialog()
        Me.splitCode = New System.Windows.Forms.SplitContainer()
        Me.codeProp = New WindowsApplication1.CodeTextBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnCodeSave = New System.Windows.Forms.Button()
        Me.btnCodeCancel = New System.Windows.Forms.Button()
        Me.pnlPropParams = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnPropParamHide = New System.Windows.Forms.Button()
        Me.chkPropParamReturnLevel3 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamReturnLevel2 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamReturnLevel1 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamLevel3 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamLevel2 = New System.Windows.Forms.CheckBox()
        Me.chkPropParamLevel1 = New System.Windows.Forms.CheckBox()
        Me.cmbPropParamReturnType = New System.Windows.Forms.ComboBox()
        Me.btnPropParamArrayAdd = New System.Windows.Forms.Button()
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
        Me.sfd = New System.Windows.Forms.SaveFileDialog()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.pnlProperty.SuspendLayout()
        Me.pnlPropBool.SuspendLayout()
        Me.pnlDataType.SuspendLayout()
        CType(Me.splitCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitCode.Panel1.SuspendLayout()
        Me.splitCode.Panel2.SuspendLayout()
        Me.splitCode.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.pnlPropParams.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 34.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.btnPropGenerateHelp, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(580, 598)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(206, 55)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Image = Global.WindowsApplication1.My.Resources.Resources.editProp32
        Me.OK_Button.Location = New System.Drawing.Point(138, 5)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(64, 45)
        Me.OK_Button.TabIndex = 0
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Image = Global.WindowsApplication1.My.Resources.Resources.delete32
        Me.Cancel_Button.Location = New System.Drawing.Point(71, 5)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(59, 45)
        Me.Cancel_Button.TabIndex = 1
        '
        'btnPropGenerateHelp
        '
        Me.btnPropGenerateHelp.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnPropGenerateHelp.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropGenerateHelp.ForeColor = System.Drawing.Color.Navy
        Me.btnPropGenerateHelp.Image = Global.WindowsApplication1.My.Resources.Resources.generateHelp32
        Me.btnPropGenerateHelp.Location = New System.Drawing.Point(4, 5)
        Me.btnPropGenerateHelp.Name = "btnPropGenerateHelp"
        Me.btnPropGenerateHelp.Size = New System.Drawing.Size(59, 45)
        Me.btnPropGenerateHelp.TabIndex = 36
        Me.btnPropGenerateHelp.UseVisualStyleBackColor = True
        '
        'txtPropName
        '
        Me.txtPropName.Location = New System.Drawing.Point(123, 16)
        Me.txtPropName.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtPropName.Name = "txtPropName"
        Me.txtPropName.Size = New System.Drawing.Size(434, 28)
        Me.txtPropName.TabIndex = 10
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(11, 19)
        Me.Label24.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(104, 20)
        Me.Label24.TabIndex = 4
        Me.Label24.Text = "Имя свойства"
        '
        'pnlProperty
        '
        Me.pnlProperty.Controls.Add(Me.btnPropParamShow)
        Me.pnlProperty.Controls.Add(Me.lstPropElementClass)
        Me.pnlProperty.Controls.Add(Me.lblPropElementClass)
        Me.pnlProperty.Controls.Add(Me.cmbPropHidden)
        Me.pnlProperty.Controls.Add(Me.Label15)
        Me.pnlProperty.Controls.Add(Me.txtPropEditorName)
        Me.pnlProperty.Controls.Add(Me.Label14)
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
        Me.pnlProperty.Controls.Add(Me.Label19)
        Me.pnlProperty.Controls.Add(Me.Label22)
        Me.pnlProperty.Controls.Add(Me.Label23)
        Me.pnlProperty.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.pnlProperty.Location = New System.Drawing.Point(17, 58)
        Me.pnlProperty.Name = "pnlProperty"
        Me.pnlProperty.Size = New System.Drawing.Size(771, 535)
        Me.pnlProperty.TabIndex = 15
        Me.pnlProperty.Visible = False
        '
        'btnPropParamShow
        '
        Me.btnPropParamShow.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamShow.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnPropParamShow.Image = Global.WindowsApplication1.My.Resources.Resources.edit32
        Me.btnPropParamShow.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPropParamShow.Location = New System.Drawing.Point(14, 217)
        Me.btnPropParamShow.Name = "btnPropParamShow"
        Me.btnPropParamShow.Size = New System.Drawing.Size(143, 43)
        Me.btnPropParamShow.TabIndex = 37
        Me.btnPropParamShow.Text = "Параметры..."
        Me.btnPropParamShow.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPropParamShow.UseVisualStyleBackColor = True
        Me.btnPropParamShow.Visible = False
        '
        'lstPropElementClass
        '
        Me.lstPropElementClass.FormattingEnabled = True
        Me.lstPropElementClass.ItemHeight = 20
        Me.lstPropElementClass.Location = New System.Drawing.Point(225, 322)
        Me.lstPropElementClass.Name = "lstPropElementClass"
        Me.lstPropElementClass.Size = New System.Drawing.Size(512, 184)
        Me.lstPropElementClass.Sorted = True
        Me.lstPropElementClass.TabIndex = 36
        Me.lstPropElementClass.Visible = False
        '
        'lblPropElementClass
        '
        Me.lblPropElementClass.AutoSize = True
        Me.lblPropElementClass.Location = New System.Drawing.Point(256, 299)
        Me.lblPropElementClass.Name = "lblPropElementClass"
        Me.lblPropElementClass.Size = New System.Drawing.Size(452, 20)
        Me.lblPropElementClass.TabIndex = 35
        Me.lblPropElementClass.Text = "Одно из имен класса, которому принадлежит данный элемент"
        Me.lblPropElementClass.Visible = False
        '
        'cmbPropHidden
        '
        Me.cmbPropHidden.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPropHidden.FormattingEnabled = True
        Me.cmbPropHidden.Items.AddRange(New Object() {"Свойство открыто", "Скрыть в редакторе", "Скрыть в коде", "Скрыть полностью", "Только в настройках по умолчанию", "Только для элементов 2 уровня", "Только для элементов 3 уровня", "Только для 1 и элементов  2 уровня", "Только для 1 и элементов  3 уровня", "Только для элементов  2 и 3 уровня"})
        Me.cmbPropHidden.Location = New System.Drawing.Point(459, 15)
        Me.cmbPropHidden.Name = "cmbPropHidden"
        Me.cmbPropHidden.Size = New System.Drawing.Size(291, 28)
        Me.cmbPropHidden.TabIndex = 31
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(392, 19)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(63, 20)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "Скрыть"
        '
        'txtPropEditorName
        '
        Me.txtPropEditorName.Location = New System.Drawing.Point(132, 16)
        Me.txtPropEditorName.Name = "txtPropEditorName"
        Me.txtPropEditorName.Size = New System.Drawing.Size(254, 28)
        Me.txtPropEditorName.TabIndex = 29
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(10, 19)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(126, 20)
        Me.Label14.TabIndex = 28
        Me.Label14.Text = "Имя в редакторе"
        '
        'btnEditCode
        '
        Me.btnEditCode.Location = New System.Drawing.Point(701, 218)
        Me.btnEditCode.Margin = New System.Windows.Forms.Padding(2)
        Me.btnEditCode.Name = "btnEditCode"
        Me.btnEditCode.Size = New System.Drawing.Size(35, 29)
        Me.btnEditCode.TabIndex = 26
        Me.btnEditCode.Text = "..."
        Me.btnEditCode.UseVisualStyleBackColor = True
        '
        'txtProp
        '
        Me.txtProp.Location = New System.Drawing.Point(189, 170)
        Me.txtProp.Margin = New System.Windows.Forms.Padding(2)
        Me.txtProp.Name = "txtProp"
        Me.txtProp.Size = New System.Drawing.Size(558, 28)
        Me.txtProp.TabIndex = 25
        '
        'pnlPropBool
        '
        Me.pnlPropBool.Controls.Add(Me.optPropValueFalse)
        Me.pnlPropBool.Controls.Add(Me.optPropValueTrue)
        Me.pnlPropBool.Location = New System.Drawing.Point(189, 179)
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
        Me.rtbProp.Location = New System.Drawing.Point(189, 175)
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
        Me.pnlDataType.Location = New System.Drawing.Point(11, 196)
        Me.pnlDataType.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlDataType.Name = "pnlDataType"
        Me.pnlDataType.Size = New System.Drawing.Size(150, 88)
        Me.pnlDataType.TabIndex = 23
        '
        'optValueCode
        '
        Me.optValueCode.Appearance = System.Windows.Forms.Appearance.Button
        Me.optValueCode.Location = New System.Drawing.Point(4, 48)
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
        Me.optValueNormal.Location = New System.Drawing.Point(2, 14)
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
        Me.Label13.Location = New System.Drawing.Point(7, 172)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(181, 20)
        Me.Label13.TabIndex = 21
        Me.Label13.Text = "Значение по-умолчанию"
        '
        'txtPropReturnArray
        '
        Me.txtPropReturnArray.Location = New System.Drawing.Point(225, 322)
        Me.txtPropReturnArray.Multiline = True
        Me.txtPropReturnArray.Name = "txtPropReturnArray"
        Me.txtPropReturnArray.Size = New System.Drawing.Size(515, 196)
        Me.txtPropReturnArray.TabIndex = 20
        Me.txtPropReturnArray.Visible = False
        '
        'lblPropReturnArray
        '
        Me.lblPropReturnArray.AutoSize = True
        Me.lblPropReturnArray.Location = New System.Drawing.Point(260, 299)
        Me.lblPropReturnArray.Name = "lblPropReturnArray"
        Me.lblPropReturnArray.Size = New System.Drawing.Size(437, 20)
        Me.lblPropReturnArray.TabIndex = 19
        Me.lblPropReturnArray.Text = "Варианты возвращаемого значения (каждый с новой строки)"
        Me.lblPropReturnArray.Visible = False
        '
        'txtPropHelp
        '
        Me.txtPropHelp.Location = New System.Drawing.Point(132, 49)
        Me.txtPropHelp.Name = "txtPropHelp"
        Me.txtPropHelp.Size = New System.Drawing.Size(578, 28)
        Me.txtPropHelp.TabIndex = 16
        '
        'btnOpenPropHelpFile
        '
        Me.btnOpenPropHelpFile.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnOpenPropHelpFile.Location = New System.Drawing.Point(721, 52)
        Me.btnOpenPropHelpFile.Name = "btnOpenPropHelpFile"
        Me.btnOpenPropHelpFile.Size = New System.Drawing.Size(30, 24)
        Me.btnOpenPropHelpFile.TabIndex = 15
        Me.btnOpenPropHelpFile.Text = "..."
        Me.btnOpenPropHelpFile.UseVisualStyleBackColor = True
        '
        'lstPropReturn
        '
        Me.lstPropReturn.FormattingEnabled = True
        Me.lstPropReturn.ItemHeight = 20
        Me.lstPropReturn.Items.AddRange(New Object() {"Обычный", "Да / Нет", "Один из возможных", "Событие", "Длинный текст", "Элемент", "Путь к картинке", "Путь к аудиофайлу", "Путь к текстовому файлу", "Путь к таблице стилей", "Путь к скрипту", "Цвет", "Функция Писателя"})
        Me.lstPropReturn.Location = New System.Drawing.Point(12, 322)
        Me.lstPropReturn.Name = "lstPropReturn"
        Me.lstPropReturn.Size = New System.Drawing.Size(180, 204)
        Me.lstPropReturn.TabIndex = 14
        '
        'txtPropDescription
        '
        Me.txtPropDescription.Location = New System.Drawing.Point(132, 84)
        Me.txtPropDescription.Multiline = True
        Me.txtPropDescription.Name = "txtPropDescription"
        Me.txtPropDescription.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtPropDescription.Size = New System.Drawing.Size(619, 75)
        Me.txtPropDescription.TabIndex = 11
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(12, 296)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(177, 20)
        Me.Label19.TabIndex = 9
        Me.Label19.Text = "Возвращаемое значение"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(10, 52)
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
        'chkDetails
        '
        Me.chkDetails.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkDetails.Location = New System.Drawing.Point(564, 6)
        Me.chkDetails.Name = "chkDetails"
        Me.chkDetails.Size = New System.Drawing.Size(224, 46)
        Me.chkDetails.TabIndex = 16
        Me.chkDetails.Text = "Детальная настройка..."
        Me.chkDetails.UseVisualStyleBackColor = True
        '
        'ofd
        '
        Me.ofd.Filter = "(html-файлы)|*.html;*.htm"
        '
        'splitCode
        '
        Me.splitCode.Location = New System.Drawing.Point(760, 16)
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
        Me.splitCode.Panel2.Controls.Add(Me.TableLayoutPanel2)
        Me.splitCode.Size = New System.Drawing.Size(381, 388)
        Me.splitCode.SplitterDistance = 328
        Me.splitCode.SplitterWidth = 3
        Me.splitCode.TabIndex = 17
        Me.splitCode.Visible = False
        '
        'codeProp
        '
        Me.codeProp.AutoWordSelection = False
        Me.codeProp.codeBox.CanDrawWords = True
        Me.codeProp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.codeProp.codeBox.DontShowError = False
        Me.codeProp.codeBox.HelpFile = "file:///D:/Projects/MatewQuest2/MatewQuest2/bin/Debug/src/rtbHelp.html"
        Me.codeProp.codeBox.HelpPath = "D:\Projects\q\Help\"
        Me.codeProp.codeBox.IsTextBlockByDefault = False
        Me.codeProp.Location = New System.Drawing.Point(0, 0)
        Me.codeProp.Margin = New System.Windows.Forms.Padding(2)
        Me.codeProp.Multiline = True
        Me.codeProp.Name = "codeProp"
        Me.codeProp.Size = New System.Drawing.Size(381, 328)
        Me.codeProp.TabIndex = 0
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.btnCodeSave, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.btnCodeCancel, 0, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(267, 22)
        Me.TableLayoutPanel2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(110, 45)
        Me.TableLayoutPanel2.TabIndex = 2
        '
        'btnCodeSave
        '
        Me.btnCodeSave.Font = New System.Drawing.Font("Palatino Linotype", 11.25!)
        Me.btnCodeSave.Image = Global.WindowsApplication1.My.Resources.Resources.save32
        Me.btnCodeSave.Location = New System.Drawing.Point(57, 2)
        Me.btnCodeSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCodeSave.Name = "btnCodeSave"
        Me.btnCodeSave.Size = New System.Drawing.Size(42, 41)
        Me.btnCodeSave.TabIndex = 0
        Me.btnCodeSave.UseVisualStyleBackColor = True
        '
        'btnCodeCancel
        '
        Me.btnCodeCancel.Font = New System.Drawing.Font("Palatino Linotype", 11.25!)
        Me.btnCodeCancel.Image = Global.WindowsApplication1.My.Resources.Resources.delete32
        Me.btnCodeCancel.Location = New System.Drawing.Point(2, 2)
        Me.btnCodeCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCodeCancel.Name = "btnCodeCancel"
        Me.btnCodeCancel.Size = New System.Drawing.Size(42, 41)
        Me.btnCodeCancel.TabIndex = 1
        Me.btnCodeCancel.UseVisualStyleBackColor = True
        '
        'pnlPropParams
        '
        Me.pnlPropParams.Controls.Add(Me.TableLayoutPanel3)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamReturnLevel3)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamReturnLevel2)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamReturnLevel1)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamLevel3)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamLevel2)
        Me.pnlPropParams.Controls.Add(Me.chkPropParamLevel1)
        Me.pnlPropParams.Controls.Add(Me.cmbPropParamReturnType)
        Me.pnlPropParams.Controls.Add(Me.btnPropParamArrayAdd)
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
        Me.pnlPropParams.Location = New System.Drawing.Point(26, 88)
        Me.pnlPropParams.Name = "pnlPropParams"
        Me.pnlPropParams.Size = New System.Drawing.Size(748, 452)
        Me.pnlPropParams.TabIndex = 18
        Me.pnlPropParams.Visible = False
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel3.ColumnCount = 1
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.btnPropParamHide, 0, 0)
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(633, 388)
        Me.TableLayoutPanel3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(98, 55)
        Me.TableLayoutPanel3.TabIndex = 47
        '
        'btnPropParamHide
        '
        Me.btnPropParamHide.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnPropParamHide.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnPropParamHide.ForeColor = System.Drawing.Color.Navy
        Me.btnPropParamHide.Image = Global.WindowsApplication1.My.Resources.Resources.ok32
        Me.btnPropParamHide.Location = New System.Drawing.Point(7, 4)
        Me.btnPropParamHide.Name = "btnPropParamHide"
        Me.btnPropParamHide.Size = New System.Drawing.Size(83, 47)
        Me.btnPropParamHide.TabIndex = 38
        Me.btnPropParamHide.UseVisualStyleBackColor = True
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
        'sfd
        '
        Me.sfd.Filter = "HTML files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*"
        '
        'dlgProperty
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(801, 655)
        Me.Controls.Add(Me.pnlPropParams)
        Me.Controls.Add(Me.splitCode)
        Me.Controls.Add(Me.chkDetails)
        Me.Controls.Add(Me.pnlProperty)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.txtPropName)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(817, 160)
        Me.Name = "dlgProperty"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Редактирование свойства"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.pnlProperty.ResumeLayout(False)
        Me.pnlProperty.PerformLayout()
        Me.pnlPropBool.ResumeLayout(False)
        Me.pnlPropBool.PerformLayout()
        Me.pnlDataType.ResumeLayout(False)
        Me.splitCode.Panel1.ResumeLayout(False)
        Me.splitCode.Panel2.ResumeLayout(False)
        CType(Me.splitCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitCode.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.pnlPropParams.ResumeLayout(False)
        Me.pnlPropParams.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents txtPropName As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents pnlProperty As System.Windows.Forms.Panel
    Friend WithEvents cmbPropHidden As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtPropEditorName As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnEditCode As System.Windows.Forms.Button
    Friend WithEvents txtProp As System.Windows.Forms.TextBox
    Friend WithEvents pnlPropBool As System.Windows.Forms.Panel
    Friend WithEvents optPropValueFalse As System.Windows.Forms.RadioButton
    Friend WithEvents optPropValueTrue As System.Windows.Forms.RadioButton
    Friend WithEvents rtbProp As System.Windows.Forms.RichTextBox
    Friend WithEvents pnlDataType As System.Windows.Forms.Panel
    Friend WithEvents optValueCode As System.Windows.Forms.RadioButton
    Friend WithEvents optValueNormal As System.Windows.Forms.RadioButton
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtPropReturnArray As System.Windows.Forms.TextBox
    Friend WithEvents lblPropReturnArray As System.Windows.Forms.Label
    Friend WithEvents txtPropHelp As System.Windows.Forms.TextBox
    Friend WithEvents btnOpenPropHelpFile As System.Windows.Forms.Button
    Friend WithEvents lstPropReturn As System.Windows.Forms.ListBox
    Friend WithEvents txtPropDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents chkDetails As System.Windows.Forms.CheckBox
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents splitCode As System.Windows.Forms.SplitContainer
    Friend WithEvents codeProp As WindowsApplication1.CodeTextBox
    Friend WithEvents btnCodeCancel As System.Windows.Forms.Button
    Friend WithEvents btnCodeSave As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents lstPropElementClass As System.Windows.Forms.ListBox
    Friend WithEvents lblPropElementClass As System.Windows.Forms.Label
    Friend WithEvents pnlPropParams As System.Windows.Forms.Panel
    Friend WithEvents chkPropParamReturnLevel3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamReturnLevel2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamReturnLevel1 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamLevel3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamLevel2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkPropParamLevel1 As System.Windows.Forms.CheckBox
    Friend WithEvents cmbPropParamReturnType As System.Windows.Forms.ComboBox
    Friend WithEvents btnPropParamArrayAdd As System.Windows.Forms.Button
    Friend WithEvents btnPropParamHide As System.Windows.Forms.Button
    Friend WithEvents btnPropParamReturnAdd As System.Windows.Forms.Button
    Friend WithEvents btnPropParamDn As System.Windows.Forms.Button
    Friend WithEvents btnPropParamUp As System.Windows.Forms.Button
    Friend WithEvents btnPropParamRemove As System.Windows.Forms.Button
    Friend WithEvents btnPropParamAdd As System.Windows.Forms.Button
    Friend WithEvents txtPropParamReturnDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents txtPropParamDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtPropParamName As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents lstPropParams As System.Windows.Forms.ListBox
    Friend WithEvents btnPropParamShow As System.Windows.Forms.Button
    Friend WithEvents btnPropGenerateHelp As System.Windows.Forms.Button
    Friend WithEvents sfd As System.Windows.Forms.SaveFileDialog
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel

End Class
