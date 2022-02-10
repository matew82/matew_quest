<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgNewClass
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(dlgNewClass))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.nudClassLevels = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtNewClassName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.pbIcon = New System.Windows.Forms.PictureBox()
        Me.btnClassIcon = New System.Windows.Forms.Button()
        Me.ofd = New System.Windows.Forms.OpenFileDialog()
        Me.tlpRemove = New System.Windows.Forms.TableLayoutPanel()
        Me.btnRemoveClass = New System.Windows.Forms.Button()
        Me.cmbClassDefProperty = New System.Windows.Forms.ComboBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.txtClassHelpFile = New System.Windows.Forms.TextBox()
        Me.btnClassHelpFile = New System.Windows.Forms.Button()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.nudClassLevels, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tlpRemove.SuspendLayout()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(785, 475)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45.0!))
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
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 107)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 20)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Уровни класса"
        '
        'nudClassLevels
        '
        Me.nudClassLevels.Location = New System.Drawing.Point(162, 105)
        Me.nudClassLevels.Maximum = New Decimal(New Integer() {3, 0, 0, 0})
        Me.nudClassLevels.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudClassLevels.Name = "nudClassLevels"
        Me.nudClassLevels.Size = New System.Drawing.Size(57, 28)
        Me.nudClassLevels.TabIndex = 6
        Me.nudClassLevels.Value = New Decimal(New Integer() {3, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(221, 20)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Имена класса (через запятую)"
        '
        'txtNewClassName
        '
        Me.txtNewClassName.Location = New System.Drawing.Point(268, 33)
        Me.txtNewClassName.Name = "txtNewClassName"
        Me.txtNewClassName.Size = New System.Drawing.Size(741, 28)
        Me.txtNewClassName.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(12, 65)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(595, 21)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Рекомендовано первым именем писать наиболее краткую форму, последним - полную."
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(12, 176)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(997, 294)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = resources.GetString("Label4.Text")
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pbIcon
        '
        Me.pbIcon.Location = New System.Drawing.Point(962, 93)
        Me.pbIcon.Name = "pbIcon"
        Me.pbIcon.Size = New System.Drawing.Size(48, 48)
        Me.pbIcon.TabIndex = 11
        Me.pbIcon.TabStop = False
        '
        'btnClassIcon
        '
        Me.btnClassIcon.Location = New System.Drawing.Point(748, 101)
        Me.btnClassIcon.Name = "btnClassIcon"
        Me.btnClassIcon.Size = New System.Drawing.Size(208, 34)
        Me.btnClassIcon.TabIndex = 12
        Me.btnClassIcon.Text = "Иконка класса..."
        Me.btnClassIcon.UseVisualStyleBackColor = True
        '
        'ofd
        '
        Me.ofd.FileName = "OpenFileDialog1"
        Me.ofd.Filter = "PNG files (*.png)|*.png"
        '
        'tlpRemove
        '
        Me.tlpRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tlpRemove.ColumnCount = 1
        Me.tlpRemove.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpRemove.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.tlpRemove.Controls.Add(Me.btnRemoveClass, 0, 0)
        Me.tlpRemove.Location = New System.Drawing.Point(16, 475)
        Me.tlpRemove.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.tlpRemove.Name = "tlpRemove"
        Me.tlpRemove.RowCount = 1
        Me.tlpRemove.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpRemove.Size = New System.Drawing.Size(124, 45)
        Me.tlpRemove.TabIndex = 13
        '
        'btnRemoveClass
        '
        Me.btnRemoveClass.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnRemoveClass.Image = Global.WindowsApplication1.My.Resources.Resources.delete26
        Me.btnRemoveClass.Location = New System.Drawing.Point(12, 5)
        Me.btnRemoveClass.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnRemoveClass.Name = "btnRemoveClass"
        Me.btnRemoveClass.Size = New System.Drawing.Size(100, 35)
        Me.btnRemoveClass.TabIndex = 0
        '
        'cmbClassDefProperty
        '
        Me.cmbClassDefProperty.FormattingEnabled = True
        Me.cmbClassDefProperty.Location = New System.Drawing.Point(284, 145)
        Me.cmbClassDefProperty.Name = "cmbClassDefProperty"
        Me.cmbClassDefProperty.Size = New System.Drawing.Size(424, 28)
        Me.cmbClassDefProperty.TabIndex = 27
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(12, 147)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(273, 20)
        Me.Label29.TabIndex = 26
        Me.Label29.Text = "Свойство по умолчанию для 3 уровня"
        '
        'txtClassHelpFile
        '
        Me.txtClassHelpFile.Location = New System.Drawing.Point(395, 105)
        Me.txtClassHelpFile.Name = "txtClassHelpFile"
        Me.txtClassHelpFile.Size = New System.Drawing.Size(277, 28)
        Me.txtClassHelpFile.TabIndex = 25
        '
        'btnClassHelpFile
        '
        Me.btnClassHelpFile.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.btnClassHelpFile.Location = New System.Drawing.Point(680, 107)
        Me.btnClassHelpFile.Name = "btnClassHelpFile"
        Me.btnClassHelpFile.Size = New System.Drawing.Size(30, 24)
        Me.btnClassHelpFile.TabIndex = 24
        Me.btnClassHelpFile.Text = "..."
        Me.btnClassHelpFile.UseVisualStyleBackColor = True
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(284, 109)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(106, 20)
        Me.Label28.TabIndex = 23
        Me.Label28.Text = "Файл помощи"
        '
        'dlgNewClass
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(1021, 538)
        Me.Controls.Add(Me.cmbClassDefProperty)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.txtClassHelpFile)
        Me.Controls.Add(Me.btnClassHelpFile)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.tlpRemove)
        Me.Controls.Add(Me.btnClassIcon)
        Me.Controls.Add(Me.pbIcon)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.nudClassLevels)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNewClassName)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Font = New System.Drawing.Font("Palatino Linotype", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgNewClass"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Создание нового класса"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.nudClassLevels, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbIcon, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tlpRemove.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents nudClassLevels As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtNewClassName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents pbIcon As System.Windows.Forms.PictureBox
    Friend WithEvents btnClassIcon As System.Windows.Forms.Button
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents tlpRemove As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnRemoveClass As System.Windows.Forms.Button
    Friend WithEvents cmbClassDefProperty As System.Windows.Forms.ComboBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents txtClassHelpFile As System.Windows.Forms.TextBox
    Friend WithEvents btnClassHelpFile As System.Windows.Forms.Button
    Friend WithEvents Label28 As System.Windows.Forms.Label

End Class
