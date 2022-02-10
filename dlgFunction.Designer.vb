<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgFunction
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
        Me.pnlFunction = New System.Windows.Forms.Panel()
        Me.cmbFuncHidden = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtFuncEditorName = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
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
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.chkDetails = New System.Windows.Forms.CheckBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtFuncName = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnFuncGenerateHelp = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.ofd = New System.Windows.Forms.OpenFileDialog()
        Me.sfd = New System.Windows.Forms.SaveFileDialog()
        Me.pnlFunction.SuspendLayout()
        Me.pnlFuncParams.SuspendLayout()
        CType(Me.nudFuncMax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudFuncMin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlFunction
        '
        Me.pnlFunction.BackColor = System.Drawing.SystemColors.Control
        Me.pnlFunction.Controls.Add(Me.cmbFuncHidden)
        Me.pnlFunction.Controls.Add(Me.Label16)
        Me.pnlFunction.Controls.Add(Me.txtFuncEditorName)
        Me.pnlFunction.Controls.Add(Me.Label17)
        Me.pnlFunction.Controls.Add(Me.pnlFuncParams)
        Me.pnlFunction.Controls.Add(Me.txtFuncResult)
        Me.pnlFunction.Controls.Add(Me.lblFuncResult)
        Me.pnlFunction.Controls.Add(Me.txtFuncHelpFile)
        Me.pnlFunction.Controls.Add(Me.btnOpenFuncHelpFile)
        Me.pnlFunction.Controls.Add(Me.lstFuncReturnType)
        Me.pnlFunction.Controls.Add(Me.nudFuncMax)
        Me.pnlFunction.Controls.Add(Me.nudFuncMin)
        Me.pnlFunction.Controls.Add(Me.txtFuncDescription)
        Me.pnlFunction.Controls.Add(Me.Label8)
        Me.pnlFunction.Controls.Add(Me.Label7)
        Me.pnlFunction.Controls.Add(Me.Label6)
        Me.pnlFunction.Controls.Add(Me.Label5)
        Me.pnlFunction.Controls.Add(Me.Label4)
        Me.pnlFunction.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.pnlFunction.Location = New System.Drawing.Point(19, 56)
        Me.pnlFunction.Name = "pnlFunction"
        Me.pnlFunction.Size = New System.Drawing.Size(763, 560)
        Me.pnlFunction.TabIndex = 12
        Me.pnlFunction.Visible = False
        '
        'cmbFuncHidden
        '
        Me.cmbFuncHidden.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFuncHidden.FormattingEnabled = True
        Me.cmbFuncHidden.Items.AddRange(New Object() {"Функция открыта", "Скрыть в редакторе", "Скрыть в коде", "Скрыть полностью"})
        Me.cmbFuncHidden.Location = New System.Drawing.Point(456, 17)
        Me.cmbFuncHidden.Name = "cmbFuncHidden"
        Me.cmbFuncHidden.Size = New System.Drawing.Size(295, 28)
        Me.cmbFuncHidden.TabIndex = 35
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(391, 20)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(63, 20)
        Me.Label16.TabIndex = 34
        Me.Label16.Text = "Скрыть"
        '
        'txtFuncEditorName
        '
        Me.txtFuncEditorName.Location = New System.Drawing.Point(135, 16)
        Me.txtFuncEditorName.Name = "txtFuncEditorName"
        Me.txtFuncEditorName.Size = New System.Drawing.Size(250, 28)
        Me.txtFuncEditorName.TabIndex = 33
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(10, 20)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(126, 20)
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "Имя в редакторе"
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
        Me.pnlFuncParams.Location = New System.Drawing.Point(362, 172)
        Me.pnlFuncParams.Name = "pnlFuncParams"
        Me.pnlFuncParams.Size = New System.Drawing.Size(398, 374)
        Me.pnlFuncParams.TabIndex = 21
        Me.pnlFuncParams.Visible = False
        '
        'lstParamsClass
        '
        Me.lstParamsClass.FormattingEnabled = True
        Me.lstParamsClass.ItemHeight = 20
        Me.lstParamsClass.Location = New System.Drawing.Point(183, 243)
        Me.lstParamsClass.Name = "lstParamsClass"
        Me.lstParamsClass.Size = New System.Drawing.Size(208, 124)
        Me.lstParamsClass.Sorted = True
        Me.lstParamsClass.TabIndex = 38
        Me.lstParamsClass.Visible = False
        '
        'lblParamsClass
        '
        Me.lblParamsClass.AutoSize = True
        Me.lblParamsClass.Font = New System.Drawing.Font("Palatino Linotype", 9.0!)
        Me.lblParamsClass.Location = New System.Drawing.Point(187, 223)
        Me.lblParamsClass.Name = "lblParamsClass"
        Me.lblParamsClass.Size = New System.Drawing.Size(185, 17)
        Me.lblParamsClass.TabIndex = 37
        Me.lblParamsClass.Text = "Одно из имен класса элемента"
        Me.lblParamsClass.Visible = False
        '
        'lstFuncParams
        '
        Me.lstFuncParams.FormattingEnabled = True
        Me.lstFuncParams.ItemHeight = 20
        Me.lstFuncParams.Location = New System.Drawing.Point(3, 30)
        Me.lstFuncParams.Name = "lstFuncParams"
        Me.lstFuncParams.Size = New System.Drawing.Size(172, 304)
        Me.lstFuncParams.TabIndex = 19
        '
        'btnParamDown
        '
        Me.btnParamDown.Enabled = False
        Me.btnParamDown.Font = New System.Drawing.Font("Palatino Linotype", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnParamDown.ForeColor = System.Drawing.Color.Navy
        Me.btnParamDown.Location = New System.Drawing.Point(90, 337)
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
        Me.btnParamUp.Location = New System.Drawing.Point(51, 337)
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
        Me.btnParamRemove.Location = New System.Drawing.Point(129, 337)
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
        Me.btnParamAdd.Location = New System.Drawing.Point(3, 338)
        Me.btnParamAdd.Name = "btnParamAdd"
        Me.btnParamAdd.Size = New System.Drawing.Size(42, 33)
        Me.btnParamAdd.TabIndex = 29
        Me.btnParamAdd.Text = "+"
        Me.btnParamAdd.UseVisualStyleBackColor = True
        '
        'txtParamsEnum
        '
        Me.txtParamsEnum.Location = New System.Drawing.Point(187, 243)
        Me.txtParamsEnum.Multiline = True
        Me.txtParamsEnum.Name = "txtParamsEnum"
        Me.txtParamsEnum.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtParamsEnum.Size = New System.Drawing.Size(203, 127)
        Me.txtParamsEnum.TabIndex = 28
        '
        'lblParamEnum
        '
        Me.lblParamEnum.AutoSize = True
        Me.lblParamEnum.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblParamEnum.Location = New System.Drawing.Point(184, 223)
        Me.lblParamEnum.Name = "lblParamEnum"
        Me.lblParamEnum.Size = New System.Drawing.Size(202, 17)
        Me.lblParamEnum.TabIndex = 27
        Me.lblParamEnum.Text = "Варианты (каждый с новой строки)"
        Me.lblParamEnum.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(183, 189)
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
        Me.cmbParamType.Location = New System.Drawing.Point(223, 192)
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
        Me.txtParamDescription.Size = New System.Drawing.Size(205, 105)
        Me.txtParamDescription.TabIndex = 23
        '
        'txtParamName
        '
        Me.txtParamName.Location = New System.Drawing.Point(181, 29)
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
        Me.txtFuncResult.Location = New System.Drawing.Point(18, 367)
        Me.txtFuncResult.Multiline = True
        Me.txtFuncResult.Name = "txtFuncResult"
        Me.txtFuncResult.Size = New System.Drawing.Size(328, 180)
        Me.txtFuncResult.TabIndex = 20
        Me.txtFuncResult.Visible = False
        '
        'lblFuncResult
        '
        Me.lblFuncResult.AutoSize = True
        Me.lblFuncResult.Location = New System.Drawing.Point(14, 331)
        Me.lblFuncResult.Name = "lblFuncResult"
        Me.lblFuncResult.Size = New System.Drawing.Size(342, 20)
        Me.lblFuncResult.TabIndex = 19
        Me.lblFuncResult.Text = "Варианты результата (каждый с новой строки)"
        Me.lblFuncResult.Visible = False
        '
        'txtFuncHelpFile
        '
        Me.txtFuncHelpFile.Location = New System.Drawing.Point(135, 53)
        Me.txtFuncHelpFile.Name = "txtFuncHelpFile"
        Me.txtFuncHelpFile.Size = New System.Drawing.Size(580, 28)
        Me.txtFuncHelpFile.TabIndex = 16
        '
        'btnOpenFuncHelpFile
        '
        Me.btnOpenFuncHelpFile.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnOpenFuncHelpFile.Location = New System.Drawing.Point(721, 54)
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
        Me.lstFuncReturnType.Location = New System.Drawing.Point(166, 253)
        Me.lstFuncReturnType.Name = "lstFuncReturnType"
        Me.lstFuncReturnType.Size = New System.Drawing.Size(180, 64)
        Me.lstFuncReturnType.TabIndex = 14
        '
        'nudFuncMax
        '
        Me.nudFuncMax.Location = New System.Drawing.Point(293, 217)
        Me.nudFuncMax.Name = "nudFuncMax"
        Me.nudFuncMax.Size = New System.Drawing.Size(53, 28)
        Me.nudFuncMax.TabIndex = 13
        '
        'nudFuncMin
        '
        Me.nudFuncMin.Location = New System.Drawing.Point(293, 180)
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
        Me.txtFuncDescription.Size = New System.Drawing.Size(616, 70)
        Me.txtFuncDescription.TabIndex = 11
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(10, 272)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(148, 20)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Результат функции"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(10, 219)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(250, 20)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Максимальное кол-во параметров"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(10, 182)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(245, 20)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Минимальное кол-во параметров"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(14, 56)
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
        'chkDetails
        '
        Me.chkDetails.Appearance = System.Windows.Forms.Appearance.Button
        Me.chkDetails.Location = New System.Drawing.Point(568, 5)
        Me.chkDetails.Name = "chkDetails"
        Me.chkDetails.Size = New System.Drawing.Size(224, 46)
        Me.chkDetails.TabIndex = 19
        Me.chkDetails.Text = "Детальная настройка..."
        Me.chkDetails.UseVisualStyleBackColor = True
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(15, 18)
        Me.Label24.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(107, 20)
        Me.Label24.TabIndex = 17
        Me.Label24.Text = "Имя функции"
        '
        'txtFuncName
        '
        Me.txtFuncName.Location = New System.Drawing.Point(127, 15)
        Me.txtFuncName.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtFuncName.Name = "txtFuncName"
        Me.txtFuncName.Size = New System.Drawing.Size(434, 28)
        Me.txtFuncName.TabIndex = 18
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnFuncGenerateHelp, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 2, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(500, 622)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(282, 55)
        Me.TableLayoutPanel1.TabIndex = 20
        '
        'btnFuncGenerateHelp
        '
        Me.btnFuncGenerateHelp.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnFuncGenerateHelp.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnFuncGenerateHelp.ForeColor = System.Drawing.Color.Navy
        Me.btnFuncGenerateHelp.Image = Global.WindowsApplication1.My.Resources.Resources.generateHelp32
        Me.btnFuncGenerateHelp.Location = New System.Drawing.Point(21, 6)
        Me.btnFuncGenerateHelp.Name = "btnFuncGenerateHelp"
        Me.btnFuncGenerateHelp.Size = New System.Drawing.Size(52, 43)
        Me.btnFuncGenerateHelp.TabIndex = 37
        Me.btnFuncGenerateHelp.UseVisualStyleBackColor = True
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Image = Global.WindowsApplication1.My.Resources.Resources.delete32
        Me.Cancel_Button.Location = New System.Drawing.Point(103, 5)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(75, 45)
        Me.Cancel_Button.TabIndex = 1
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Image = Global.WindowsApplication1.My.Resources.Resources.editFunc32
        Me.OK_Button.Location = New System.Drawing.Point(197, 5)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(75, 45)
        Me.OK_Button.TabIndex = 0
        '
        'ofd
        '
        Me.ofd.Filter = "(html-файлы)|*.html;*.htm"
        '
        'sfd
        '
        Me.sfd.Filter = "HTML files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*"
        '
        'dlgFunction
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(801, 676)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.chkDetails)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.txtFuncName)
        Me.Controls.Add(Me.pnlFunction)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(817, 160)
        Me.Name = "dlgFunction"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Редактирование функции"
        Me.pnlFunction.ResumeLayout(False)
        Me.pnlFunction.PerformLayout()
        Me.pnlFuncParams.ResumeLayout(False)
        Me.pnlFuncParams.PerformLayout()
        CType(Me.nudFuncMax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudFuncMin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pnlFunction As System.Windows.Forms.Panel
    Friend WithEvents cmbFuncHidden As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtFuncEditorName As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
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
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents chkDetails As System.Windows.Forms.CheckBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtFuncName As System.Windows.Forms.TextBox
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lstParamsClass As System.Windows.Forms.ListBox
    Friend WithEvents lblParamsClass As System.Windows.Forms.Label
    Friend WithEvents btnFuncGenerateHelp As System.Windows.Forms.Button
    Friend WithEvents sfd As System.Windows.Forms.SaveFileDialog

End Class
