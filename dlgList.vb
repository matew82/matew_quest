Imports System.Windows.Forms

Public Class dlgList
    ''' <summary>Id класса, свойство которого реактируется </summary>
    Public classId As Integer = -1
    ''' <summary>Новый список</summary>
    Public newList() As String = Nothing
    ''' <summary>Имя создаваемого/редактируемого свойства</summary>
    Public propertyName As String = ""

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        'Создание нового или редактирование существующего списка свойства
        'проверка корректности данных
        If classId < 0 Then
            MsgBox("Не указан класс, диалоговое окно открыто некорректно.", MsgBoxStyle.Exclamation)
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
            Return
        End If

        Erase newList
        'Заполняем массив arrReturnEnum значениями, которые возвращает свойство
        Dim arrReturnEnum() As String = Nothing, arrReturnEnumUBound As Integer = -1
        Dim curStr As String
        For i As Integer = 0 To txtPropReturnArray.Lines.GetUpperBound(0)
            curStr = txtPropReturnArray.Lines(i).Trim
            If curStr.Length > 0 Then
                If Double.TryParse(curStr, Globalization.NumberStyles.Any, provider_points, Nothing) = False Then
                    If curStr.ToLower = "true" Then
                        curStr = "True"
                    ElseIf curStr.ToLower = "false" Then
                        curStr = "False"
                    Else
                        If curStr.StartsWith("'") = False And curStr.EndsWith("'") = False Then
                            curStr = "'" + curStr.Replace("'", "/'") + "'"
                        End If
                    End If
                End If
                arrReturnEnumUBound += 1
                ReDim Preserve arrReturnEnum(arrReturnEnumUBound)
                arrReturnEnum(arrReturnEnumUBound) = curStr
            End If
        Next

        'сохраняем данные в структуре mainClass
        Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties(propertyName)
        If arrReturnEnumUBound > 0 Then
            newList = arrReturnEnum
            p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ENUM
        Else
            Dim retType As MatewScript.ReturnFunctionEnum = MatewScript.ReturnFunctionEnum.RETURN_USUAL
            Dim pCopy As MatewScript.PropertiesInfoType = Nothing
            If classId < mScript.mainClassCopy.Count AndAlso mScript.mainClassCopy(classId).Properties.TryGetValue(propertyName, pCopy) Then
                If pCopy.returnType <> MatewScript.ReturnFunctionEnum.RETURN_ENUM Then retType = pCopy.returnType
            End If
            p.returnType = retType
        End If
        p.returnArray = newList
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgList_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub PrepareData(ByVal classId As Integer, ByVal propName As String)
        Me.classId = classId
        If classId < 0 OrElse String.IsNullOrWhiteSpace(propName) Then
            MessageBox.Show("Неверные параметры открытия диалогового окна!")
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Return
        End If
        propertyName = propName
        Me.Text = "Редактирование списка свойства " + propName + " класса " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.Count - 1)

        'Заполняем форму редактирования свойства
        txtPropReturnArray.Clear()
        If IsNothing(mScript.mainClass(classId).Properties(propertyName).returnArray) = False Then
            'заполняем текстбокс значениями, которые может возвращать свойство
            For i As Integer = 0 To mScript.mainClass(classId).Properties(propertyName).returnArray.GetUpperBound(0)
                txtPropReturnArray.AppendText(mScript.PrepareStringToPrint(mScript.mainClass(classId).Properties(propertyName).returnArray(i), Nothing, False) + vbNewLine)
            Next
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Восстановление списка по умолчанию
        'проверка корректности данных
        If classId < 0 Then
            MsgBox("Не указан класс, диалоговое окно открыто некорректно.", MsgBoxStyle.Exclamation)
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
            Return
        End If

        Erase newList
        'Заполняем массив arrReturnEnum значениями, которые возвращает свойство
        Dim arrReturnEnum() As String = Nothing
        Dim pCopy As MatewScript.PropertiesInfoType = Nothing
        If mScript.mainClassCopy(classId).Properties.TryGetValue(propertyName, pCopy) = False Then
            If mScript.mainClass(classId).Properties(propertyName).UserAdded Then
                MsgBox("Не найдена копия свойства " + propertyName + ".", vbExclamation)
            Else
                MsgBox("Свойство " + propertyName + " не является встроенным и имеет сохраненной копии - оно добавлено Писателем.", vbInformation)
            End If
            Return
        End If

        If IsNothing(pCopy.returnArray) OrElse pCopy.returnArray.Count = 0 Then
            MsgBox("Свойство " + propertyName + " изначально не являлось списком.", vbInformation)
            Return
        End If
        ReDim arrReturnEnum(pCopy.returnArray.Count - 1)
        Array.ConstrainedCopy(pCopy.returnArray, 0, arrReturnEnum, 0, arrReturnEnum.Count)

        'сохраняем данные в структуре mainClass
        Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties(propertyName)
        newList = arrReturnEnum
        p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ENUM
        p.returnArray = newList
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class
