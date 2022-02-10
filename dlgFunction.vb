Imports System.Windows.Forms

Public Class dlgFunction
    ''' <summary>для хранения копии скриптов перед их редактированием </summary>
    Dim CodeData() As CodeTextBox.CodeDataType
    ''' <summary>Id класса, функция которого создается/реактируется </summary>
    Public classId As Integer = -1
    ''' <summary>Создается ли новая функция или редактируется существующая </summary>
    Public isNewFunction As Boolean = False
    ''' <summary>Все характеристики новой функции</summary>
    Public newFunctionData As MatewScript.PropertiesInfoType
    ''' <summary>Имя создаваемой/редактируемой функции</summary>
    Public newFunctionName As String = ""
    ''' <summary>для хранения параметров редактируемой функции</summary>
    Dim curFuncParams() As MatewScript.paramsType

    Public heightWithPanel As Integer = 715
    Public heightWithoutPanel As Integer = 160

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgFunction_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ofd.InitialDirectory = APP_HELP_PATH
        Me.Height = heightWithoutPanel
    End Sub

    Private Sub ClearData()
        'отображаем форму для создания новоой функции
        lstFuncReturnType.SelectedIndex = 0
        'очищаем все от старых данных
        txtFuncName.Clear()
        txtFuncDescription.Clear()
        txtFuncEditorName.Clear()
        cmbFuncHidden.SelectedIndex = 0
        txtFuncHelpFile.Clear()
        txtFuncResult.Clear()
        nudFuncMin.Value = 0
        nudFuncMax.Value = 0
        lstFuncParams.Items.Clear()
        lstFuncReturnType.SelectedIndex = 0
        Erase curFuncParams 'очищаем структуру curFuncParams
        txtFuncName.Focus()
    End Sub

    Private Sub btnOpenFuncHelpFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenFuncHelpFile.Click
        'получение имени файла помощи для функции или свойства
        Dim initDir As String = My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "help") + "\"
        ofd.InitialDirectory = initDir
        ofd.ShowDialog()
        If ofd.FileName.Length = 0 Then Exit Sub
        Dim strFileName = ofd.FileName
        'If strFileName.StartsWith(APP_HELP_PATH) Then strFileName = strFileName.Substring(APP_HELP_PATH.Length)
        If strFileName.StartsWith(APP_HELP_PATH, StringComparison.CurrentCultureIgnoreCase) Then
            strFileName = strFileName.Substring(APP_HELP_PATH.Length)
        ElseIf strFileName.StartsWith(initDir, StringComparison.CurrentCultureIgnoreCase) Then
            strFileName = strFileName.Substring(initDir.Length)
        End If
        txtFuncHelpFile.Text = strFileName
    End Sub

    Private Sub chkDetails_CheckedChanged(sender As Object, e As EventArgs) Handles chkDetails.CheckedChanged
        pnlFunction.Visible = chkDetails.Checked
        If pnlFunction.Visible Then
            Me.Height = heightWithPanel
        Else
            Me.Height = heightWithoutPanel
        End If
    End Sub

    Private Sub dlgFunction_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Not Me.Visible Then Return
        If chkDetails.Checked Then
            Me.Height = heightWithPanel
        Else
            Me.Height = heightWithoutPanel
        End If
        txtFuncName.Focus()
    End Sub

    Public Sub PrepareData(ByVal classId As Integer, ByVal isNewFunction As Boolean, Optional ByVal funcName As String = "")
        Me.classId = classId
        Me.isNewFunction = isNewFunction
        If classId < 0 OrElse (isNewFunction = False AndAlso String.IsNullOrWhiteSpace(funcName)) Then
            MessageBox.Show("Неверные параметры открытия диалогового окна!")
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
            Return
        End If
        newFunctionName = funcName

        FillListWithClasses()
        If isNewFunction Then
            Me.Text = "Создание новой функции класса " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.Count - 1)
            OK_Button.Image = My.Resources.addFunc32
            ClearData()
            chkDetails.Checked = False
            Return
        Else
            Me.Text = "Редактирование функции " + funcName + " класса " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.Count - 1)
            OK_Button.Image = My.Resources.editFunc32
            chkDetails.Checked = True
        End If

        'Заполняем форму редактирования функции
        Dim fData As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Functions(newFunctionName)
        txtFuncName.Text = newFunctionName
        txtFuncDescription.Text = fData.Description
        txtFuncEditorName.Text = fData.EditorCaption
        cmbFuncHidden.SelectedIndex = fData.Hidden
        txtFuncHelpFile.Text = fData.helpFile
        lstFuncReturnType.SelectedIndex = fData.returnType
        txtFuncResult.Clear()
        If IsNothing(fData.returnArray) = False Then
            'заполняем текстбокс значениями, которые может возвращать функция
            For i As Integer = 0 To fData.returnArray.GetUpperBound(0)
                txtFuncResult.AppendText(fData.returnArray(i) + vbNewLine)
            Next
        End If
        'заполняем структуру curFuncParams данными о параметрах функции
        curFuncParams = fData.params
        'очищаем список параметров функции
        lstFuncParams.Items.Clear()
        If IsNothing(fData.params) OrElse fData.params.Count = 0 Then
            'параметров нет
            pnlFuncParams.Hide()
            nudFuncMin.Value = 0
            nudFuncMax.Value = 0
        Else
            'заполняем список lstFuncParams названиями параметров функции
            lstFuncParams.BeginUpdate()
            For i As Integer = 0 To fData.params.GetUpperBound(0)
                lstFuncParams.Items.Add(fData.params(i).Name)
            Next
            If fData.paramsMax = -1 Then
                'максимальное кол-во параметров неограниченно - значение в nudFuncMax ставим равным кол-ву описанных параметров (хотя он все-равно будет недоступен)
                nudFuncMax.Value = fData.params.GetUpperBound(0) + 1
            Else
                'максимальное кол-во параметров четко определено - значение в nudFuncMax берем из структуры mainClass
                nudFuncMax.Value = fData.paramsMax
            End If
            nudFuncMin.Value = fData.paramsMin
            pnlFuncParams.Show()
            If lstFuncParams.Items.Count > 0 Then lstFuncParams.SelectedIndex = 0
            lstFuncParams.EndUpdate()
        End If

    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        SaveFunction()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    ''' <summary>
    ''' Сохраняет введенные данные
    ''' </summary>
    Private Sub SaveFunction()
        'Создание новой или редактирование существующей функции
        'проверка корректности данных
        If classId < 0 Then
            MsgBox("Не указан класс, диалоговое окно открыто некорректно.", MsgBoxStyle.Exclamation)
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
            Return
        End If
        Dim funcName As String = txtFuncName.Text.Trim
        If funcName.Length = 0 Then
            MsgBox("Укажите имя функции.", MsgBoxStyle.Exclamation)
            txtFuncName.Focus()
            Exit Sub
        End If

        If mScript.mainClass(classId).Functions.ContainsKey(funcName) Then
            If isNewFunction OrElse (funcName <> newFunctionName) Then
                MsgBox("Функция " + funcName + " уже существует. Укажите другое имя.", MsgBoxStyle.Exclamation)
                txtFuncName.Focus()
                Exit Sub
            End If
        End If

        Dim funcEditorName As String = txtFuncEditorName.Text.Trim
        If String.IsNullOrEmpty(funcEditorName) Then funcEditorName = funcName

        If cmbFuncHidden.SelectedIndex = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL Then
            'Выбрано Скрыть полностью
            Dim dRes As MsgBoxResult = MsgBox("В настройках сокрытия функции установлено значение [Скрыть полностью]. Функция создастся, но никаких следов ее существования Вы не увидите. Тем не менее, она будет работать и может буть удалена с помощью строки " + _
                   mScript.mainClass(classId).Names(0) + ".RemoveFunction('" + funcName + "'). Продолжить?", MsgBoxStyle.YesNo + MsgBoxStyle.Information)
            If dRes = MsgBoxResult.No Then
                cmbFuncHidden.Focus()
                Return
            End If
        End If

        Dim paramsUBound As Integer = -1
        If IsNothing(curFuncParams) = False Then
            paramsUBound = curFuncParams.GetUpperBound(0)
            For i As Integer = 0 To paramsUBound
                If curFuncParams(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY AndAlso i <> paramsUBound Then
                    MsgBox("Параметр " + curFuncParams(i).Name + " имеет тип " + Chr(34) + "Массив параметров" + Chr(34) + ", который может иметь лишь последний параметр.", MsgBoxStyle.Exclamation)
                    lstFuncParams.SelectedIndex = i
                    Exit Sub
                ElseIf curFuncParams(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM Then
                    If IsNothing(curFuncParams(i).EnumValues) = True OrElse curFuncParams(i).EnumValues.GetUpperBound(0) = -1 Then
                        MsgBox("Параметр " + curFuncParams(i).Name + " имеет тип " + Chr(34) + "Один из возможных" + Chr(34) + ", для которого необходимо перечисление этих возможных вариантов.", MsgBoxStyle.Exclamation)
                        lstFuncParams.SelectedIndex = i
                        Exit Sub
                    End If
                ElseIf curFuncParams(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT Then
                    If IsNothing(curFuncParams(i).EnumValues) = True OrElse curFuncParams(i).EnumValues.GetUpperBound(0) = -1 Then
                        MsgBox("Параметр " + curFuncParams(i).Name + " имеет тип " + Chr(34) + "Элемент" + Chr(34) + ", для которого необходимо выбрать дочерний элемент.", MsgBoxStyle.Exclamation)
                        lstFuncParams.SelectedIndex = i
                        Exit Sub
                    End If
                End If
            Next
        End If

        'Заполняем массив arrReturnEnum значениями, которые возвращает функция
        Dim arrReturnEnum() As String = Nothing, arrReturnEnumUBound As Integer = -1
        If lstFuncReturnType.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ENUM Then
            Dim curStr As String
            For i As Integer = 0 To txtFuncResult.Lines.GetUpperBound(0)
                curStr = txtFuncResult.Lines(i).Trim
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
            If arrReturnEnumUBound = -1 Then
                MsgBox("Для результата функции " + Chr(34) + "один из возможных" + Chr(34) + " надо перечислить эти возможные варианты.", MsgBoxStyle.Exclamation)
                txtFuncResult.Focus()
                Exit Sub
            End If
        End If

        'Заполняем func введенными данными о функции
        Dim func As New MatewScript.PropertiesInfoType
        func.Description = txtFuncDescription.Text
        func.helpFile = txtFuncHelpFile.Text
        func.paramsMax = IIf(nudFuncMax.Enabled, nudFuncMax.Value, -1)
        func.paramsMin = nudFuncMin.Value
        func.params = curFuncParams
        func.EditorCaption = funcEditorName
        func.Hidden = cmbFuncHidden.SelectedIndex
        func.returnType = lstFuncReturnType.SelectedIndex
        If func.returnType = MatewScript.ReturnFunctionEnum.RETURN_ENUM OrElse func.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then func.returnArray = arrReturnEnum

        func.UserAdded = True
        newFunctionData = func
        newFunctionName = funcName
    End Sub

    Private Sub txtParamName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtParamName.TextChanged
        'Сохраняем имя параметра функции
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        curFuncParams(selIndex).Name = sender.Text
        lstFuncParams.Items(selIndex) = sender.Text 'обновляем имя параметра в списке lstFuncParams
    End Sub

    Private Sub txtParamDescription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtParamDescription.TextChanged
        'Сохраняем описание параметра функции
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        curFuncParams(selIndex).Description = sender.Text
    End Sub

    Private Sub txtParamsEnum_TextChanged(ByVal sender As TextBox, ByVal e As System.EventArgs) Handles txtParamsEnum.TextChanged
        'Сохраняем набор возможных значений параметра функции
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        Erase curFuncParams(selIndex).EnumValues 'очищаем предыдущие данные в стр-ре curFuncParams
        If sender.Text.Trim.Length = 0 Then Exit Sub

        'Приводим варианты значений в правильный вид и сохраняем в curFuncParams
        Dim strCurLine As String
        Dim parEnumUBound As Integer = -1
        For i As Integer = 0 To sender.Lines.GetUpperBound(0)
            strCurLine = sender.Lines(i).Trim
            If strCurLine.Length = 0 Then Continue For
            parEnumUBound += 1
            ReDim Preserve curFuncParams(selIndex).EnumValues(parEnumUBound)

            If Double.TryParse(strCurLine, Globalization.NumberStyles.Any, provider_points, Nothing) Then
                curFuncParams(selIndex).EnumValues(parEnumUBound) = strCurLine
            Else
                If strCurLine.StartsWith("'") And strCurLine.EndsWith("'") And strCurLine.Length > 1 Then
                    curFuncParams(selIndex).EnumValues(parEnumUBound) = strCurLine
                Else
                    If strCurLine.ToLower = "true" Then
                        curFuncParams(selIndex).EnumValues(parEnumUBound) = "True"
                    ElseIf strCurLine.ToLower = "false" Then
                        curFuncParams(selIndex).EnumValues(parEnumUBound) = "False"
                    Else
                        curFuncParams(selIndex).EnumValues(parEnumUBound) = "'" + strCurLine.Replace("'", "/'") + "'"
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub cmbParamType_SelectedIndexChanged(ByVal sender As ComboBox, ByVal e As System.EventArgs) Handles cmbParamType.SelectedIndexChanged
        'Сохраняем тип параметра функции
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        curFuncParams(selIndex).Type = sender.SelectedIndex
        'если параметр возвращает одно из набора значений, показываем txtParamsEnum для ввода туда этих значений. Иначе его прячем.
        lblParamEnum.Visible = (curFuncParams(selIndex).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM)
        txtParamsEnum.Visible = lblParamEnum.Visible
        lblParamsClass.Visible = (curFuncParams(selIndex).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT)
        lstParamsClass.Visible = lblParamsClass.Visible
        If curFuncParams(selIndex).Type <> MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT AndAlso curFuncParams(selIndex).Type <> MatewScript.paramsType.paramsTypeEnum.PARAM_ENUM Then
            curFuncParams(selIndex).EnumValues = Nothing
        End If

        'Если тип хоть одного - это массив параметров, то счетчик максимального кол-ва параметров делаем недоступным
        For i As Integer = 0 To curFuncParams.GetUpperBound(0)
            If curFuncParams(i).Type = MatewScript.paramsType.paramsTypeEnum.PARAMS_ARRAY Then
                nudFuncMax.Enabled = False
                Exit Sub
            End If
        Next
        nudFuncMax.Enabled = True
    End Sub

    Private Sub lstFuncParams_SelectedIndexChanged(ByVal sender As ListBox, ByVal e As System.EventArgs) Handles lstFuncParams.SelectedIndexChanged
        'При выборе в lstFuncParams нового элемента выводим в соответствующих контролах инфу о параметре для ее редактирования
        'в зависимости от того, выбран ли первый или последний параметр, меняем доступность кнопок, меняющих позицию параметра
        Dim selIndex As Integer = sender.SelectedIndex
        If sender.SelectedIndex = 0 Then
            btnParamUp.Enabled = False
        Else
            btnParamUp.Enabled = True
        End If
        If selIndex = sender.Items.Count - 1 Then
            btnParamDown.Enabled = False
        Else
            btnParamDown.Enabled = True
        End If
        If selIndex = -1 Then Exit Sub
        'заполняем соответствующие поля данными о параметре из стр-ры curFuncParams
        txtParamName.Text = curFuncParams(selIndex).Name
        txtParamDescription.Text = curFuncParams(selIndex).Description
        cmbParamType.SelectedIndex = curFuncParams(selIndex).Type
        If IsNothing(curFuncParams(selIndex).EnumValues) Then
            txtParamsEnum.Clear()
            lstParamsClass.SelectedIndex = -1
        Else
            If curFuncParams(selIndex).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ELEMENT AndAlso curFuncParams(selIndex).EnumValues(0).Length > 2 Then
                'если параметр является элементом, то выбираем значение из набора
                'txtParamsEnum.Clear()
                Dim selClass As String = curFuncParams(selIndex).EnumValues(0)
                Dim selClassId As Integer = -1
                If mScript.mainClassHash.TryGetValue(selClass, selClassId) Then
                    selClass = mScript.mainClass(selClassId).Names.Last
                    Dim res As Integer = lstParamsClass.Items.IndexOf(selClass)
                    lstParamsClass.SelectedIndex = res
                Else
                    lstParamsClass.SelectedIndex = -1
                End If
            Else
                'если параметр принимает одно из набора возможных значений, заполняем txtParamsEnum этими значениями
                Dim sBuilder As New System.Text.StringBuilder
                For i = 0 To curFuncParams(sender.SelectedIndex).EnumValues.GetUpperBound(0)
                    sBuilder.AppendLine(curFuncParams(sender.SelectedIndex).EnumValues(i))
                Next
                txtParamsEnum.Text = sBuilder.ToString
                lstParamsClass.SelectedIndex = -1
            End If

        End If
    End Sub

    Private Sub lstFuncReturnType_SelectedIndexChanged(ByVal sender As ListBox, ByVal e As System.EventArgs) Handles lstFuncReturnType.SelectedIndexChanged
        'Если функция возвращает один из набора вариантов, то показываем поля для ввода этих вариантов. Иначе это поле прячем
        Dim blnShow As Boolean = IIf(sender.SelectedIndex = 2, True, False)
        txtFuncResult.Visible = blnShow
        lblFuncResult.Visible = blnShow
    End Sub

    Private Sub nudFuncMax_ValueChanged(ByVal sender As NumericUpDown, ByVal e As System.EventArgs) Handles nudFuncMax.ValueChanged
        'изменилось значение максимального кол-ва параметров функции
        pnlFuncParams.Visible = IIf(sender.Value = 0, False, True)
        'если минимальное кол-во параметров оказалось больше максимального - делаем их равными
        If sender.Value < nudFuncMin.Value Then nudFuncMin.Value = sender.Value
        If lstFuncParams.Items.Count < sender.Value Then
            'число увеличилось - добавляем параметр
            AddFuncParam()
            lstFuncParams.SelectedIndex = lstFuncParams.Items.Count - 1
        ElseIf lstFuncParams.Items.Count > sender.Value Then
            'число уменьшилось - удаляем параметр
            RemoveFuncParam()
            If lstFuncParams.Items.Count > 0 Then lstFuncParams.SelectedIndex = 0
        End If
    End Sub

    Private Sub nudFuncMin_ValueChanged(ByVal sender As NumericUpDown, ByVal e As System.EventArgs) Handles nudFuncMin.ValueChanged
        'если максимальное кол-во параметров функциии стало меньше минимального - делаем их равными
        If nudFuncMax.Value < sender.Value Then nudFuncMax.Value = sender.Value
    End Sub

    Private Sub AddFuncParam()
        'Добавляем функции новый параметр
        Dim parUBound As Integer = 0
        Dim defParName As String = "Param1"
        If IsNothing(curFuncParams) = False AndAlso curFuncParams.GetUpperBound(0) > -1 Then
            'параметры функции уже существуют
            'В defParName получаем имя по-умолчанию для нового параметра, вида "ParamX", где х - любое число, такое, чтобы имя параметра было уникальным
            parUBound = curFuncParams.Length
            Do
                For i = 0 To parUBound - 1
                    If curFuncParams(i).Name = defParName Then
                        defParName = defParName.Substring(0, 5) + Convert.ToString(Convert.ToInt32(defParName.Substring(5)) + 1)
                        Continue Do
                    End If
                Next
                Exit Do
            Loop
        End If
        'Создаем параматр в структуре curFuncParams со значениями по-умолчанию
        ReDim Preserve curFuncParams(parUBound)
        curFuncParams(parUBound) = New MatewScript.paramsType
        curFuncParams(parUBound).Name = defParName
        curFuncParams(parUBound).Type = MatewScript.paramsType.paramsTypeEnum.PARAM_ANY
        'добавляем параметр в список lstFuncParams
        lstFuncParams.Items.Add(defParName)
    End Sub

    Private Sub RemoveFuncParam(Optional ByVal paramIndex As Integer = -1)
        'Удаляет параметр функции
        If paramIndex = -1 Then paramIndex = lstFuncParams.Items.Count - 1
        If paramIndex = -1 Then Exit Sub
        'удаляем параметр из списка
        lstFuncParams.Items.Remove(lstFuncParams.Items(paramIndex))

        If IsNothing(curFuncParams) OrElse curFuncParams.GetUpperBound(0) = 0 Then
            Erase curFuncParams 'параметр был последним - очищаем curFuncParams и выход 
            Exit Sub
        End If
        'удаляем параметр из структуры curFuncParams
        For i = paramIndex To curFuncParams.GetUpperBound(0) - 1
            curFuncParams(i) = curFuncParams(i + 1)
        Next
        ReDim Preserve curFuncParams(curFuncParams.GetUpperBound(0) - 1)
    End Sub

    Private Sub btnParamAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParamAdd.Click
        'увеличение счетчика nudFuncMax приводит к созданию нового параметра функции
        nudFuncMax.Value += 1
    End Sub

    Private Sub btnParamRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParamRemove.Click
        'удаляем выбранный в lstFuncParams параметр функции
        RemoveFuncParam(lstFuncParams.SelectedIndex)
        If nudFuncMax.Value > lstFuncParams.Items.Count Then nudFuncMax.Value = lstFuncParams.Items.Count
        If lstFuncParams.Items.Count > 0 Then lstFuncParams.SelectedIndex = 0
    End Sub

    Private Sub btnParamUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParamUp.Click
        'Перемещаем параметр функции на 1 вверх (будет идти раньше)
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex < 1 Then Exit Sub
        'В curFuncParams меняем местами выбранный параметр и предыдущий
        Dim paramCopy As MatewScript.paramsType = curFuncParams(selIndex - 1)
        curFuncParams(selIndex - 1) = curFuncParams(selIndex)
        curFuncParams(selIndex) = paramCopy
        'Таким же образом меняем местами параметры в списке lstFuncParams
        Dim strItemCopy As String = lstFuncParams.GetItemText(lstFuncParams.Items(selIndex - 1))
        lstFuncParams.Items(selIndex - 1) = lstFuncParams.Items(selIndex)
        lstFuncParams.Items(selIndex) = strItemCopy
        lstFuncParams.SelectedIndex = selIndex - 1
    End Sub

    Private Sub btnParamDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParamDown.Click
        'Перемещаем параметр функции на 1 вниз (будет идти позже)
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Exit Sub
        If selIndex = lstFuncParams.Items.Count - 1 Then Exit Sub
        'В curFuncParams меняем местами выбранный параметр и последующий
        Dim paramCopy As MatewScript.paramsType = curFuncParams(selIndex + 1)
        curFuncParams(selIndex + 1) = curFuncParams(selIndex)
        curFuncParams(selIndex) = paramCopy
        'Таким же образом меняем местами параметры в списке lstFuncParams
        Dim strItemCopy As String = lstFuncParams.GetItemText(lstFuncParams.Items(selIndex + 1))
        lstFuncParams.Items(selIndex + 1) = lstFuncParams.Items(selIndex)
        lstFuncParams.Items(selIndex) = strItemCopy
        lstFuncParams.SelectedIndex = selIndex + 1
    End Sub

    Private Sub lstParamsClass_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstParamsClass.SelectedIndexChanged
        'Сохраняем класс при значении paramType = ELEMENT
        Dim selIndex As Integer = lstFuncParams.SelectedIndex
        If selIndex = -1 Then Return
        If lstParamsClass.SelectedIndex = -1 Then Return
        curFuncParams(selIndex).EnumValues = {lstParamsClass.Text}
    End Sub

    ''' <summary>
    ''' Заполняет listView с названиями классов 2 уровня (для выбора если returnType = ELEMENT)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FillListWithClasses()
        lstParamsClass.Items.Clear()
        If IsNothing(mScript.mainClass) OrElse mScript.mainClass.Count = 0 Then Return
        For i As Integer = 0 To mScript.mainClass.Count - 1
            If mScript.mainClass(i).LevelsCount = 0 Then Continue For
            lstParamsClass.Items.Add(mScript.mainClass(i).Names.Last)
        Next
        lstParamsClass.Items.Add("Variable")
    End Sub

    Private Sub btnFuncGenerateHelp_Click(sender As Object, e As EventArgs) Handles btnFuncGenerateHelp.Click
        'Открывает диалог генерации файла помощи
        Dim dRes As DialogResult = Windows.Forms.DialogResult.No
        Dim fPath As String = ""
        If txtFuncHelpFile.Text.Length > 0 Then
            fPath = GetHelpPath(txtFuncHelpFile.Text)
            If My.Computer.FileSystem.FileExists(fPath) Then
                dRes = MessageBox.Show("К функции уже присоединен файл помощи. Редактировать его?", "Matew Quest", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If dRes = Windows.Forms.DialogResult.Cancel Then Return
            End If
        End If

        If txtFuncName.Text.Length = 0 Then
            MsgBox("Не введено имя функции!")
            txtFuncName.Focus()
            Return
        End If
        If txtFuncDescription.Text.Length = 0 Then
            MsgBox("Не введено описание функции!")
            txtFuncDescription.Focus()
            Return
        End If
        'Перед началом сохраняем свойство
        SaveFunction()
        'Dim prop As MatewScript.PropertiesInfoType = Nothing
        'If mScript.mainClass(classId).Properties.TryGetValue(txtPropName.Text, prop) = False Then Return
        If dRes = Windows.Forms.DialogResult.No Then
            'Генерируем файл помощи
            dlgGenerateHelp.GenerateHelpForFunction(txtFuncName.Text, classId, newFunctionData)
        Else
            'Открываем готовый файл
            dlgGenerateHelp.codeHelp.Text = My.Computer.FileSystem.ReadAllText(fPath, System.Text.Encoding.Default)
        End If
        dRes = dlgGenerateHelp.ShowDialog(Me)
        pnlFunction.Show()

        If dRes = Windows.Forms.DialogResult.OK Then
            'Открываем диалог сохранения файла
            'Выбираем начальную директорию
            If fPath.Length > 0 Then
                Dim fName As String = System.IO.Path.GetFileName(fPath)
                fPath = fPath.Substring(0, fPath.Length - fName.Length)
            Else
                fPath = APP_HELP_PATH
            End If
            sfd.InitialDirectory = fPath
            dRes = sfd.ShowDialog(Me)
            If dRes = Windows.Forms.DialogResult.OK Then
                'Сохраняем файл
                fPath = sfd.FileName
                My.Computer.FileSystem.WriteAllText(fPath, dlgGenerateHelp.codeHelp.Text, False, System.Text.Encoding.Default)
                Dim initDir As String = My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "help") + "\"
                If fPath.StartsWith(APP_HELP_PATH, StringComparison.CurrentCultureIgnoreCase) Then
                    fPath = fPath.Substring(APP_HELP_PATH.Length)
                ElseIf fPath.StartsWith(initDir, StringComparison.CurrentCultureIgnoreCase) Then
                    fPath = fPath.Substring(initDir.Length)
                End If
                'fPath = fPath.Replace(APP_HELP_PATH, "")
                txtFuncHelpFile.Text = fPath
            End If
        End If
    End Sub
End Class
