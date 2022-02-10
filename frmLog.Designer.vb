<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLog
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
        Me.lstLog = New System.Windows.Forms.ListBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lstLog
        '
        Me.lstLog.FormattingEnabled = True
        Me.lstLog.ItemHeight = 19
        Me.lstLog.Location = New System.Drawing.Point(18, 18)
        Me.lstLog.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.lstLog.Name = "lstLog"
        Me.lstLog.Size = New System.Drawing.Size(1110, 650)
        Me.lstLog.TabIndex = 0
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(1136, 18)
        Me.btnClear.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(112, 34)
        Me.btnClear.TabIndex = 1
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(1136, 60)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 34)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Mark"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmLog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1256, 679)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.lstLog)
        Me.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmLog"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lstLog As System.Windows.Forms.ListBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
