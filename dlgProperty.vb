Imports System.Windows.Forms

Public Class dlgProperty
    ''' <summary>
    ''' Класс для временного хранения параметров свойств типа Event
    ''' </summary>
    Private Class PropParamsClass
        'Levelx, T_PARAM_ARRAY и T_RETURN заполняются исходя из значения Property.params().Type:
        'Level1 +1, Level2 +2, Level3 +4, Return +8, ParamArray +16
        Public Enum ParamTypeEnum As Byte
            T_PARAM = 0
            T_PARAM_ARRAY = 1
            T_RETURN = 2
        End Enum
        Public Level1 As Boolean = False
        Public Level2 As Boolean = False
        Public Level3 As Boolean = False
        Public paramType As ParamTypeEnum
        Public paramCaption As String
        Public paramDescription As String
    End Class
    Private arrPropParams As List(Of PropParamsClass)

    ''' <summary>для хранения копии скриптов перед их редактированием </summary>
    Dim CodeData() As CodeTextBox.CodeDataType
    ''' <summary>Id класса, свойство которого создается/реактируется </summary>
    Public classId As Integer = -1
    ''' <summary>Создается ли новое свойство или редактируется существующее </summary>
    Public isNewProperty As Boolean = False
    ''' <summary>Все характеристики нового свойства</summary>
    Public newPropertyData As MatewScript.PropertiesInfoType
    ''' <summary>Имя создаваемого/редактируемого свойства</summary>
    Public newPropertyName As String = ""
    Public heightWithPanel As Integer = 700
    Public heightWithoutPanel As Integer = 160

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not SaveProperty() Then Return
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    ''' <summary>
    ''' Сохраняет введенные данные
    ''' </summary>
    Private Function SaveProperty() As Boolean
        'Создание нового или редактирование существующего свойства
        'проверка корректности данных
        If classId < 0 Then
            MsgBox("Не указан класс, диалоговое окно открыто некорректно.", MsgBoxStyle.Exclamation)
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
            Return False
        End If

        Dim propName As String = txtPropName.Text.Trim
        If propName.Length = 0 Then
            MsgBox("Укажите имя свойства.", MsgBoxStyle.Exclamation)
            txtPropName.Focus()
            Return False
        End If

        If mScript.mainClass(classId).Properties.ContainsKey(propName) Then
            If isNewProperty OrElse (propName <> newPropertyName) Then
                MsgBox("Свойство " + propName + " уже существует. Укажите другое имя.", MsgBoxStyle.Exclamation)
                txtPropName.Focus()
                Return False
            End If
        End If

        If cmbPropHidden.SelectedIndex = MatewScript.PropertyHiddenEnum.HIDDEN_AT_ALL Then
            'Выбрано Скрыть полностью
            Dim dRes As MsgBoxResult = MsgBox("В настройках сокрытия свойства установлено значение [Скрыть полностью]. Свойство создастся, но никаких следов его существования Вы не увидите. Тем не менее, оно будет работать и может буть удалено с помощью строки " + _
                   mScript.mainClass(classId).Names(0) + ".RemoveProperty('" + propName + "'). Продолжить?", MsgBoxStyle.YesNo + MsgBoxStyle.Information)
            If dRes = MsgBoxResult.No Then
                cmbPropHidden.Focus()
                Return False
            End If
        End If

        Dim propEditorName As String = txtPropEditorName.Text.Trim
        If String.IsNullOrEmpty(propEditorName) Then propEditorName = propName

        Dim blnReturnCode As Boolean = False 'в свойстве сохраняется скрипт?
        If lstPropReturn.SelectedIndex <> MatewScript.ReturnFunctionEnum.RETURN_BOOl Then
            If (optValueCode.Checked OrElse (lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_EVENT OrElse lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION)) _
                Then blnReturnCode = True 'да, сохраняется скрипт
        End If

        If blnReturnCode AndAlso mScript.LAST_ERROR.Length > 0 Then Return False 'код содержит ошибку - выход

        'Заполняем массив arrReturnEnum значениями, которые возвращает свойство
        Dim arrReturnEnum() As String = Nothing, arrReturnEnumUBound As Integer = -1
        If lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ENUM Then
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
            If arrReturnEnumUBound = -1 Then
                MsgBox("Для типа возвращаемого значения " + Chr(34) + "один из возможных" + Chr(34) + " надо перечислить эти возможные значения.", MsgBoxStyle.Exclamation)
                txtPropReturnArray.Focus()
                Return False
            End If
        ElseIf lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then
            If lstPropElementClass.SelectedIndex = -1 Then
                MsgBox("Для типа возвращаемого значения " + Chr(34) + "элемент" + Chr(34) + " надо выбрать класс элемента.", MsgBoxStyle.Exclamation)
                lstPropElementClass.Focus()
                Return False
            End If
            ReDim arrReturnEnum(0)
            arrReturnEnum(0) = lstPropElementClass.Text
        End If

        'Заполняем prop введенными данными о свойстве
        'сохраняем данные в структуре mainClass
        Dim prop As New MatewScript.PropertiesInfoType
        prop.Description = txtPropDescription.Text
        prop.helpFile = txtPropHelp.Text
        prop.returnType = lstPropReturn.SelectedIndex
        prop.EditorCaption = propEditorName
        prop.Hidden = Math.Max(cmbPropHidden.SelectedIndex, 0)
        If blnReturnCode Then
            'заполняем значение свойства сериализованным в xml кодом
            prop.Value = codeProp.codeBox.SerializeCodeData()
        Else
            If prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_BOOl Then
                If optPropValueTrue.Checked Then
                    prop.Value = "True"
                Else
                    prop.Value = "False"
                End If
            Else
                Dim strDefValue As String = txtProp.Text.Trim
                If strDefValue.Length > 0 Then
                    If Double.TryParse(strDefValue, Globalization.NumberStyles.Any, provider_points, Nothing) Then
                        prop.Value = strDefValue
                    ElseIf strDefValue.ToLower = "true" Then
                        prop.Value = "True"
                    ElseIf strDefValue.ToLower = "false" Then
                        prop.Value = strDefValue = "False"
                    ElseIf strDefValue.Length > 1 AndAlso strDefValue.First = "'"c AndAlso strDefValue.Last = "'"c Then
                        prop.Value = strDefValue
                    Else
                        prop.Value = "'" + strDefValue.Replace("'", "/'") + "'"
                    End If
                Else
                    prop.Value = "''"
                End If
            End If
        End If

        If prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_EVENT AndAlso IsNothing(arrPropParams) = False AndAlso arrPropParams.Count > 0 Then
            'Вносим данные о параметрах свойства-события
            ReDim prop.params(arrPropParams.Count - 1)
            For i As Integer = 0 To arrPropParams.Count - 1
                prop.params(i) = New MatewScript.paramsType
                Dim strName As String = arrPropParams(i).paramCaption
                Dim scores As Integer = 0
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then scores += 16
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then scores += 8
                If arrPropParams(i).Level3 Then scores += 4
                If arrPropParams(i).Level2 Then scores += 2
                If arrPropParams(i).Level1 Then scores += 1
                prop.params(i).Type = scores
                prop.params(i).Name = strName
                prop.params(i).Description = arrPropParams(i).paramDescription
            Next
        End If

        If prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_ENUM OrElse prop.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then prop.returnArray = arrReturnEnum
        prop.UserAdded = True

        newPropertyData = prop
        newPropertyName = propName
        Return True
    End Function

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgProperty_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        splitCode.Dock = DockStyle.Fill
        splitCode.Hide()

        ofd.InitialDirectory = APP_HELP_PATH
        Me.Height = heightWithoutPanel
    End Sub

    Public Sub PrepareData(ByVal classId As Integer, ByVal isNewProperty As Boolean, Optional ByVal propName As String = "")
        Me.classId = classId
        Me.isNewProperty = isNewProperty
        pnlPropParams.Hide()
        pnlPropParams.Location = pnlProperty.Location
        pnlPropParams.Size = pnlProperty.Size
        If classId < 0 OrElse (isNewProperty = False AndAlso String.IsNullOrWhiteSpace(propName)) Then
            MessageBox.Show("Неверные параметры открытия диалогового окна!")
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
            Return
        End If
        newPropertyName = propName
        FillListWithClasses()

        If isNewProperty Then
            Me.Text = "Создание нового свойства класса " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.Count - 1)
            OK_Button.Image = My.Resources.addProp32
            ClearData()
            FillPropParamArray(propName, isNewProperty)
            chkDetails.Checked = False
            Return
        Else
            Me.Text = "Редактирование свойства " + propName + " класса " + mScript.mainClass(classId).Names(mScript.mainClass(classId).Names.Count - 1)
            OK_Button.Image = My.Resources.editProp32
            chkDetails.Checked = True
        End If

        'Заполняем форму редактирования свойства
        Dim p As MatewScript.PropertiesInfoType = mScript.mainClass(classId).Properties(newPropertyName)
        txtPropName.Text = newPropertyName
        txtPropEditorName.Text = p.EditorCaption
        cmbPropHidden.SelectedIndex = p.Hidden
        txtPropDescription.Text = p.Description
        txtPropHelp.Text = p.helpFile
        lstPropReturn.SelectedIndex = p.returnType
        txtPropReturnArray.Clear()
        FillPropParamArray(propName, isNewProperty)
        If IsNothing(p.returnArray) = False AndAlso p.returnType <> MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then
            'заполняем текстбокс значениями, которые может возвращать свойство
            For i As Integer = 0 To p.returnArray.GetUpperBound(0)
                txtPropReturnArray.AppendText(p.returnArray(i) + vbNewLine)
            Next
        ElseIf p.returnType = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT AndAlso IsNothing(p.returnArray) = False AndAlso p.returnArray(0).Length > 2 Then
            'returnArray(0) содержит имя класса, элементы котрого надо отобразить (аналог значения returnType = Location -> returnArray(0)="L")
            Dim selClass As String = p.returnArray(0)
            Dim selClassId As Integer = -1
            If mScript.mainClassHash.TryGetValue(selClass, selClassId) Then
                selClass = mScript.mainClass(selClassId).Names.Last
                Dim res As Integer = lstPropElementClass.Items.IndexOf(selClass)
                lstPropElementClass.SelectedIndex = res
            Else
                lstPropElementClass.SelectedIndex = -1
            End If
        Else
            lstPropElementClass.SelectedIndex = -1
        End If
        Dim curProVal As String = p.Value
        Select Case lstPropReturn.SelectedIndex
            Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                rtbProp.Hide()
                txtProp.Hide()
                pnlDataType.Hide()
                pnlPropBool.Show()
                If curProVal = "True" Then
                    optPropValueTrue.Checked = True
                Else
                    optPropValueFalse.Checked = True
                End If
            Case MatewScript.ReturnFunctionEnum.RETURN_EVENT, MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION
                pnlDataType.Hide()
                pnlPropBool.Hide()
                txtProp.Hide()
                rtbProp.Show()
                If lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                    codeProp.codeBox.IsTextBlockByDefault = False
                Else
                    codeProp.codeBox.IsTextBlockByDefault = True
                End If
                codeProp.codeBox.LoadCodeFromProperty(curProVal)
                rtbProp.Rtf = codeProp.codeBox.Rtf
            Case Else
                pnlDataType.Show()
                pnlPropBool.Hide()
                If mScript.IsPropertyContainsCode(curProVal) <> MatewScript.ContainsCodeEnum.NOT_CODE Then
                    codeProp.codeBox.LoadCodeFromProperty(curProVal)
                    rtbProp.Rtf = codeProp.codeBox.Rtf
                    optValueCode.Checked = True
                Else
                    txtProp.Text = curProVal
                    optValueNormal.Checked = True
                End If
        End Select
    End Sub

    Private Sub ClearData()
        'Отображаем и подготоваливаем форму для ввода нового свойства
        chkDetails.Checked = False
        rtbProp.Clear()
        txtProp.Clear()
        codeProp.codeBox.Text = ""
        optPropValueFalse.Checked = True
        optValueNormal.Checked = True
        rtbProp.ForeColor = Color.Black
        lstPropReturn.SelectedIndex = 0
        txtPropName.Clear()
        txtPropEditorName.Clear()
        cmbPropHidden.SelectedIndex = 0
        txtPropDescription.Clear()
        txtPropHelp.Clear()
        txtPropReturnArray.Clear()
        lstPropReturn.SelectedIndex = 0
        lstPropElementClass.SelectedIndex = -1
    End Sub

    Private Sub btnOpenPropHelpFile_Click(sender As Object, e As EventArgs) Handles btnOpenPropHelpFile.Click
        'получение имени файла помощи для функции или свойства
        Dim initDir As String= My.Computer.FileSystem.CombinePath(questEnvironment.QuestPath, "help") + "\"
        ofd.InitialDirectory = initDir
        ofd.ShowDialog()
        If ofd.FileName.Length = 0 Then Exit Sub
        Dim strFileName = ofd.FileName
        If strFileName.StartsWith(APP_HELP_PATH, StringComparison.CurrentCultureIgnoreCase) Then
            strFileName = strFileName.Substring(APP_HELP_PATH.Length)
        ElseIf strFileName.StartsWith(initDir, StringComparison.CurrentCultureIgnoreCase) Then
            strFileName = strFileName.Substring(initDir.Length)
        End If
        txtPropHelp.Text = strFileName
    End Sub

    Private Sub rtbProp_VisibleChanged(sender As Object, e As EventArgs) Handles rtbProp.VisibleChanged
        btnEditCode.Visible = rtbProp.Visible
    End Sub

    Private Sub btnEditCode_Click(sender As Object, e As EventArgs) Handles btnEditCode.Click
        CodeData = CopyCodeDataArray(codeProp.codeBox.CodeData) 'создание копии кода перед редактированием
        splitCode.Show()
        splitCode.SplitterDistance = splitCode.Height - btnCodeSave.Height - 30
        codeProp.Dock = DockStyle.Fill
        codeProp.Show()
        codeProp.codeBox.Refresh()
    End Sub

    Private Sub btnCodeCancel_Click(sender As Object, e As EventArgs) Handles btnCodeCancel.Click
        codeProp.codeBox.LoadCodeFromCodeData(CodeData)
        splitCode.Hide()
    End Sub

    Private Sub btnCodeSave_Click(sender As Object, e As EventArgs) Handles btnCodeSave.Click
        If mScript.LAST_ERROR.Length > 0 Then
            MsgBox("Нельзя сохранить код с ошибками!", vbExclamation)
            Return
        End If
        rtbProp.Rtf = codeProp.codeBox.Rtf
        splitCode.Hide()
    End Sub

    Private Sub optValueNormal_CheckedChanged(sender As Object, e As EventArgs) Handles optValueNormal.CheckedChanged
        txtProp.Show()
        rtbProp.Hide()
        btnEditCode.Hide()
    End Sub

    Private Sub optValueCode_CheckedChanged(sender As Object, e As EventArgs) Handles optValueCode.CheckedChanged
        txtProp.Hide()
        rtbProp.Show()
        btnEditCode.Show()
    End Sub

    Private Sub chkDetails_CheckedChanged(sender As Object, e As EventArgs) Handles chkDetails.CheckedChanged
        pnlProperty.Visible = chkDetails.Checked
        If pnlProperty.Visible Then
            Me.Height = heightWithPanel
        Else
            Me.Height = heightWithoutPanel
        End If
    End Sub

    Private Sub lstPropReturn_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstPropReturn.SelectedIndexChanged
        lblPropReturnArray.Hide()
        txtPropReturnArray.Hide()
        lblPropElementClass.Hide()
        lstPropElementClass.Hide()
        btnPropParamShow.Hide()
        Select Case lstPropReturn.SelectedIndex
            Case MatewScript.ReturnFunctionEnum.RETURN_BOOl
                'тип - True/False
                txtProp.Hide()
                rtbProp.Hide()
                pnlDataType.Hide()
                pnlPropBool.Show()
            Case MatewScript.ReturnFunctionEnum.RETURN_EVENT, MatewScript.ReturnFunctionEnum.RETURN_DESCRIPTION
                'тип - событие
                txtProp.Hide()
                If lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_EVENT Then
                    btnPropParamShow.Show()
                    codeProp.codeBox.IsTextBlockByDefault = False
                Else
                    codeProp.codeBox.IsTextBlockByDefault = True
                End If
                pnlPropBool.Hide()
                pnlDataType.Hide()
                rtbProp.Show()
            Case Else
                'тип - обычный или "Один из вариантов"
                'Если тип возвращаемого значения свойства один из набора вариантов, то отображаем txtPropReturnArray для ввода
                'этих вариантов. Иначе они остаются спрятанными.
                If lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ENUM Then
                    lblPropReturnArray.Show()
                    txtPropReturnArray.Show()
                ElseIf lstPropReturn.SelectedIndex = MatewScript.ReturnFunctionEnum.RETURN_ELEMENT Then
                    lblPropElementClass.Show()
                    lstPropElementClass.Show()
                End If
                pnlPropBool.Hide()
                pnlDataType.Show()
                If optValueCode.Checked Then
                    rtbProp.Show()
                    txtProp.Hide()
                Else
                    rtbProp.Hide()
                    txtProp.Show()
                End If
        End Select
    End Sub

    Private Sub dlgProperty_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Not Me.Visible Then Return
        If chkDetails.Checked Then
            Me.Height = heightWithPanel
        Else
            Me.Height = heightWithoutPanel
        End If
        txtPropName.Focus()
    End Sub

    Private Sub FillListWithClasses()
        lstPropElementClass.Items.Clear()
        If IsNothing(mScript.mainClass) OrElse mScript.mainClass.Count = 0 Then Return
        For i As Integer = 0 To mScript.mainClass.Count - 1
            If mScript.mainClass(i).LevelsCount = 0 Then Continue For
            lstPropElementClass.Items.Add(mScript.mainClass(i).Names.Last)
        Next
        lstPropElementClass.Items.Add("Variable")
    End Sub


#Region "Property Params"
    Private curPropParamId As Integer = -1

    ''' <summary>
    ''' Заполняет список параметров свойства типа EVENT, необходимых для генерации файла помощи и вывода типичных переменных
    ''' </summary>
    ''' <param name="propName">Имя свойства</param>
    Private Sub FillPropParamArray(ByVal propName As String, ByVal newProperty As Boolean)
        'paramType:
        '1 - Level1
        '2 - Level2 (3)
        '4 - Level3 (5-7)
        '8 - Return (9-15)
        '16- Array (17-31)
        ClearPropParamArray()
        Dim prop As MatewScript.PropertiesInfoType = Nothing

        If newProperty = False AndAlso mScript.mainClass(classId).Properties.TryGetValue(propName, prop) = False Then Return
        If newProperty Then
            'Новое свойство - создаем стандартные параметры
            If mScript.mainClass(classId).LevelsCount = 1 Then
                arrPropParams.Add(New PropParamsClass With {.paramType = PropParamsClass.ParamTypeEnum.T_PARAM, .paramCaption = mScript.mainClass(classId).Names(0) + "Id", _
                                                            .paramDescription = mScript.mainClass(classId).Names.Last + " Id", .Level1 = True, .Level2 = True})
            ElseIf mScript.mainClass(classId).LevelsCount = 2 Then
                arrPropParams.Add(New PropParamsClass With {.paramType = PropParamsClass.ParamTypeEnum.T_PARAM, .paramCaption = mScript.mainClass(classId).Names(0) + "Id", _
                                                           .paramDescription = mScript.mainClass(classId).Names.Last + " Id", .Level1 = True, .Level2 = True})
                arrPropParams.Add(New PropParamsClass With {.paramType = PropParamsClass.ParamTypeEnum.T_PARAM, .paramCaption = mScript.mainClass(classId).Names(0) + "ItemId", _
                                                .paramDescription = mScript.mainClass(classId).Names.Last + " subitem Id", .Level1 = True, .Level2 = True, .Level3 = True})
            End If
        Else
            If IsNothing(prop.params) OrElse prop.params.Count = 0 Then Return
            'Параметры есть - получаем
            Dim params() As MatewScript.paramsType = prop.params
            For i As Integer = 0 To params.Count - 1
                Dim propParam As New PropParamsClass
                Dim pName As String = params(i).Name
                Dim parScores As Integer = params(i).Type 'params.Type используется для хранения значений Level1, Level2, Level3 & ParamArray
                propParam.paramType = PropParamsClass.ParamTypeEnum.T_PARAM
                If parScores >= 16 Then propParam.paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY : parScores -= 16
                If parScores >= 8 Then propParam.paramType = PropParamsClass.ParamTypeEnum.T_RETURN : parScores -= 8
                If parScores >= 4 Then propParam.Level3 = True : parScores -= 4
                If parScores >= 2 Then propParam.Level2 = True : parScores -= 2
                If parScores = 1 Then propParam.Level1 = True

                propParam.paramCaption = params(i).Name
                propParam.paramDescription = params(i).Description
                arrPropParams.Add(propParam)
            Next
        End If
        'Заполняем список
        btnPropParamArrayAdd.Enabled = True
        If arrPropParams.Count > 0 Then
            lstPropParams.BeginUpdate()
            For i As Integer = 0 To arrPropParams.Count - 1
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM Then
                    lstPropParams.Items.Add("Param[" + i.ToString + "]")
                ElseIf arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
                    lstPropParams.Items.Add("Param[" + i.ToString + ", " + (i + 1).ToString + ", ... n]")
                    btnPropParamArrayAdd.Enabled = False
                Else
                    lstPropParams.Items.Add("Return " + arrPropParams(i).paramCaption)
                End If
            Next
            lstPropParams.EndUpdate()
        End If
        If lstPropParams.Items.Count > 0 Then lstPropParams.SelectedIndex = 0
    End Sub

    ''' <summary>
    ''' Очистка параметров свойств
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearPropParamArray()
        curPropParamId = -1
        lstPropParams.Items.Clear()
        If IsNothing(arrPropParams) = False Then arrPropParams.Clear()

        arrPropParams = New List(Of PropParamsClass)
        btnPropParamRemove.Enabled = False
        txtPropParamDescription.Clear()
        txtPropParamDescription.Enabled = False
        txtPropParamName.Clear()
        txtPropParamName.Enabled = False
        txtPropParamReturnDescription.Clear()
        txtPropParamReturnDescription.Enabled = False
        cmbPropParamReturnType.Text = ""
        cmbPropParamReturnType.Enabled = False
        chkPropParamLevel1.Enabled = False
        chkPropParamLevel2.Enabled = False
        chkPropParamLevel3.Enabled = False
        chkPropParamReturnLevel1.Enabled = False
        chkPropParamReturnLevel2.Enabled = False
        chkPropParamReturnLevel3.Enabled = False

        btnPropParamArrayAdd.Enabled = True
    End Sub

    Private Sub btnPropParamHide_Click(sender As Object, e As EventArgs) Handles btnPropParamHide.Click
        pnlPropParams.Hide()
    End Sub

    Private Sub btnPropParamArrayAdd_Click(sender As Object, e As EventArgs) Handles btnPropParamArrayAdd.Click
        sender.Enabled = False
        AddNewPropParam(PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY)
    End Sub

    ''' <summary>
    ''' Добавляет новый параметр свойства
    ''' </summary>
    ''' <param name="parType">Тип добавляемого параметра</param>
    Private Sub AddNewPropParam(ByVal parType As PropParamsClass.ParamTypeEnum)
        Dim propParam As New PropParamsClass With {.paramType = parType, .paramCaption = "", .paramDescription = "", .Level1 = True}
        If mScript.mainClass(classId).LevelsCount >= 1 Then propParam.Level2 = True
        If mScript.mainClass(classId).LevelsCount >= 2 Then propParam.Level3 = True

        Dim insertPos As Integer = -1
        Dim strText As String
        Dim hasParamArray As Boolean = False
        If parType <> PropParamsClass.ParamTypeEnum.T_RETURN Then
            'Если мы вставляем обычный параметр, то он должен следовать перед параметрами типа Return & Param_Array. Для этого получаем insertPos - позицию, куда надо вставить новый паарметр
            For i As Integer = 0 To arrPropParams.Count - 1
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
                    If parType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
                        MsgBox("Массив параметров уже существует и может быть только один!")
                        Return
                    End If
                    hasParamArray = True 'имеется массив параметров
                    insertPos = i
                    Exit For
                ElseIf arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
                    insertPos = i
                    Exit For
                End If
            Next i
        End If

        If insertPos = -1 Then
            'Вставляем в конец
            insertPos = arrPropParams.Count
            arrPropParams.Add(propParam)
        Else
            'Вставляем последним параметром, но перед параметрами типа Return & Param_Array
            arrPropParams.Insert(insertPos, propParam)
        End If

        'Генерирует тескт для листбоска
        If parType = PropParamsClass.ParamTypeEnum.T_PARAM Then
            strText = "Param[" + insertPos.ToString + "]"
        ElseIf parType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
            strText = "Param[" + insertPos.ToString + ", " + (insertPos + 1).ToString + ", ... n]"
        Else
            strText = "Return (ничего)"
        End If
        'Вставляем новый параметр в дистбокс
        If insertPos >= lstPropParams.Items.Count Then
            lstPropParams.Items.Add(strText)
        Else
            lstPropParams.Items.Insert(insertPos, strText)
        End If

        If hasParamArray Then
            'изменяем такст в листбоксе у параметра типа Array
            lstPropParams.Items(insertPos + 1) = "Param[" + (insertPos + 1).ToString + ", " + (insertPos + 2).ToString + ", ... n]"
        End If

        lstPropParams.SelectedIndex = insertPos
    End Sub

    Private Sub btnPropParamAdd_Click(sender As Object, e As EventArgs) Handles btnPropParamAdd.Click
        AddNewPropParam(PropParamsClass.ParamTypeEnum.T_PARAM)
    End Sub

    Private Sub btnPropParamReturnAdd_Click(sender As Object, e As EventArgs) Handles btnPropParamReturnAdd.Click
        AddNewPropParam(PropParamsClass.ParamTypeEnum.T_RETURN)
    End Sub

    Private Sub btnPropParamRemove_Click(sender As Object, e As EventArgs) Handles btnPropParamRemove.Click
        'Удаляем параметр свойства
        If lstPropParams.SelectedIndex = -1 Then Return
        Dim delId As Integer = lstPropParams.SelectedIndex

        If arrPropParams(delId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM Then
            'Изменяем текст в листбоксе для параметра типа Array, а также обычных параметров, если они находились выше удаляемого (индекс -1)
            For i As Integer = 0 To arrPropParams.Count - 1
                If arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
                    lstPropParams.Items(i) = "Param[" + (i - 1).ToString + ", " + i.ToString + ", ... n]"
                    Exit For
                ElseIf arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
                    Exit For
                ElseIf arrPropParams(i).paramType = PropParamsClass.ParamTypeEnum.T_PARAM AndAlso i > delId Then
                    lstPropParams.Items(i) = "Param[" + (i - 1).ToString + "]"
                End If
            Next
        ElseIf arrPropParams(delId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
            btnPropParamArrayAdd.Enabled = True
        End If
        'Собственно удаление и выбор ближайшего другого параметра
        arrPropParams.RemoveAt(delId)
        lstPropParams.Items.RemoveAt(delId)
        If arrPropParams.Count = 0 Then
            btnPropParamRemove.Enabled = False
            Return
        ElseIf delId >= arrPropParams.Count - 1 Then
            lstPropParams.SelectedIndex = delId
        Else
            delId -= 1
            lstPropParams.SelectedIndex = delId
        End If
    End Sub

    Private Sub btnPropParamUp_Click(sender As Object, e As EventArgs) Handles btnPropParamUp.Click
        Dim movId As Integer = lstPropParams.SelectedIndex
        If CheckPropParamButtonUp() = False Then Return
        Dim movParam As PropParamsClass = arrPropParams(movId)
        arrPropParams.RemoveAt(movId)
        arrPropParams.Insert(movId - 1, movParam)

        If arrPropParams(movId - 1).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
            Dim strText As String = lstPropParams.Items(movId)
            lstPropParams.Items.RemoveAt(movId)
            lstPropParams.Items.Insert(movId - 1, strText)
        End If
        lstPropParams.SelectedIndex = movId - 1
    End Sub

    Private Sub btnPropParamDn_Click(sender As Object, e As EventArgs) Handles btnPropParamDn.Click
        Dim movId As Integer = lstPropParams.SelectedIndex
        If CheckPropParamButtonDn() = False Then Return
        Dim movParam As PropParamsClass = arrPropParams(movId)
        arrPropParams.RemoveAt(movId)
        arrPropParams.Insert(movId + 1, movParam)

        If arrPropParams(movId + 1).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
            Dim strText As String = lstPropParams.Items(movId)
            lstPropParams.Items.RemoveAt(movId)
            lstPropParams.Items.Insert(movId + 1, strText)
        End If
        lstPropParams.SelectedIndex = movId + 1
    End Sub

    ''' <summary>
    ''' Проверяет доступно ли перемещение текущего параметра вверх, и, если недоступно , то делает его кнопку также недоступной
    ''' </summary>
    Private Function CheckPropParamButtonUp() As Boolean
        Dim curId As Integer = lstPropParams.SelectedIndex
        If curId <= 0 Then
            btnPropParamUp.Enabled = False
            Return False
        End If
        If arrPropParams(curId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
            btnPropParamUp.Enabled = False
            Return False
        End If
        If arrPropParams(curId).paramType = PropParamsClass.ParamTypeEnum.T_RETURN AndAlso arrPropParams(curId - 1).paramType <> PropParamsClass.ParamTypeEnum.T_RETURN Then
            btnPropParamUp.Enabled = False
            Return False
        End If
        btnPropParamUp.Enabled = True
        Return True
    End Function

    ''' <summary>
    ''' Проверяет доступно ли перемещение текущего параметра вниз, и, если недоступно , то делает его кнопку также недоступной
    ''' </summary>
    Private Function CheckPropParamButtonDn() As Boolean
        Dim curId As Integer = lstPropParams.SelectedIndex
        If curId = -1 OrElse curId >= arrPropParams.Count - 1 Then
            btnPropParamDn.Enabled = False
            Return False
        End If
        If arrPropParams(curId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM_ARRAY Then
            btnPropParamDn.Enabled = False
            Return False
        End If
        If arrPropParams(curId).paramType = PropParamsClass.ParamTypeEnum.T_PARAM AndAlso arrPropParams(curId + 1).paramType <> PropParamsClass.ParamTypeEnum.T_PARAM Then
            btnPropParamDn.Enabled = False
            Return False
        End If
        btnPropParamDn.Enabled = True
        Return True
    End Function

    Private Sub lstPropParams_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstPropParams.SelectedIndexChanged
        If lstPropParams.SelectedIndex = -1 Then Return
        curPropParamId = lstPropParams.SelectedIndex
        CheckPropParamButtonDn()
        CheckPropParamButtonUp()

        btnPropParamRemove.Enabled = True
        If arrPropParams(curPropParamId).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then
            'Вставляем данные параметра типа Return
            txtPropParamDescription.Clear()
            txtPropParamDescription.Enabled = False
            txtPropParamName.Clear()
            txtPropParamName.Enabled = False
            chkPropParamLevel1.Enabled = False
            chkPropParamLevel2.Enabled = False
            chkPropParamLevel3.Enabled = False

            txtPropParamReturnDescription.Text = arrPropParams(curPropParamId).paramDescription
            txtPropParamReturnDescription.Enabled = True
            cmbPropParamReturnType.Text = arrPropParams(curPropParamId).paramCaption
            cmbPropParamReturnType.Enabled = True
            chkPropParamReturnLevel1.Enabled = True
            chkPropParamReturnLevel1.Checked = arrPropParams(curPropParamId).Level1
            If mScript.mainClass(classId).LevelsCount >= 1 Then
                chkPropParamReturnLevel2.Enabled = True
                chkPropParamReturnLevel2.Checked = arrPropParams(curPropParamId).Level2
            End If
            If mScript.mainClass(classId).LevelsCount >= 2 Then
                chkPropParamReturnLevel3.Enabled = True
                chkPropParamReturnLevel3.Checked = arrPropParams(curPropParamId).Level3
            End If

        Else
            'Вставляем данные параметра типа Param / ParamArray
            txtPropParamReturnDescription.Clear()
            txtPropParamReturnDescription.Enabled = False
            cmbPropParamReturnType.Text = ""
            cmbPropParamReturnType.Enabled = False
            chkPropParamReturnLevel1.Enabled = False
            chkPropParamReturnLevel2.Enabled = False
            chkPropParamReturnLevel3.Enabled = False

            txtPropParamDescription.Text = arrPropParams(curPropParamId).paramDescription
            txtPropParamDescription.Enabled = True
            txtPropParamName.Text = arrPropParams(curPropParamId).paramCaption
            txtPropParamName.Enabled = True
            chkPropParamLevel1.Enabled = True
            chkPropParamLevel1.Checked = arrPropParams(curPropParamId).Level1
            If mScript.mainClass(classId).LevelsCount >= 1 Then
                chkPropParamLevel2.Enabled = True
                chkPropParamLevel2.Checked = arrPropParams(curPropParamId).Level2
            End If
            If mScript.mainClass(classId).LevelsCount >= 2 Then
                chkPropParamLevel3.Enabled = True
                chkPropParamLevel3.Checked = arrPropParams(curPropParamId).Level3
            End If
        End If
    End Sub

    Private Sub txtPropParamName_TextChanged(sender As Object, e As EventArgs) Handles txtPropParamName.TextChanged
        If curPropParamId = -1 Then Return
        If arrPropParams(curPropParamId).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then Return
        arrPropParams(curPropParamId).paramCaption = sender.Text
    End Sub

    Private Sub txtPropParamDescription_TextChanged(sender As Object, e As EventArgs) Handles txtPropParamDescription.TextChanged
        If curPropParamId = -1 Then Return
        If arrPropParams(curPropParamId).paramType = PropParamsClass.ParamTypeEnum.T_RETURN Then Return
        arrPropParams(curPropParamId).paramDescription = sender.Text
    End Sub

    Private Sub cmbPropParamReturnType_TextChanged(sender As Object, e As EventArgs) Handles cmbPropParamReturnType.TextChanged
        If curPropParamId = -1 Then Return
        If arrPropParams(curPropParamId).paramType <> PropParamsClass.ParamTypeEnum.T_RETURN Then Return
        arrPropParams(curPropParamId).paramCaption = sender.Text
        If String.IsNullOrWhiteSpace(sender.Text) Then
            lstPropParams.Items(curPropParamId) = "Return (ничего)"
        Else
            lstPropParams.Items(curPropParamId) = "Return " + sender.Text
        End If
    End Sub

    Private Sub txtPropParamReturnDescription_TextChanged(sender As Object, e As EventArgs) Handles txtPropParamReturnDescription.TextChanged
        If curPropParamId = -1 Then Return
        If arrPropParams(curPropParamId).paramType <> PropParamsClass.ParamTypeEnum.T_RETURN Then Return
        arrPropParams(curPropParamId).paramDescription = sender.Text
    End Sub

    Private Sub btnPropParamsShow_Click(sender As Object, e As EventArgs) Handles btnPropParamShow.Click
        pnlPropParams.BringToFront()
        pnlPropParams.Show()
    End Sub

    Private Sub chkPropParamLevel1_CheckedChanged(sender As Object, e As EventArgs) Handles chkPropParamLevel1.CheckedChanged, chkPropParamLevel2.CheckedChanged, chkPropParamLevel3.CheckedChanged, _
        chkPropParamReturnLevel1.CheckedChanged, chkPropParamReturnLevel2.CheckedChanged, chkPropParamReturnLevel3.CheckedChanged
        If curPropParamId = -1 Then Return
        Dim lvl As Integer = CInt(sender.Tag)
        If lvl = 1 Then
            arrPropParams(curPropParamId).Level1 = sender.Checked
        ElseIf lvl = 2 Then
            arrPropParams(curPropParamId).Level2 = sender.Checked
        Else
            arrPropParams(curPropParamId).Level3 = sender.Checked
        End If
    End Sub

#End Region

    Private Sub btnPropGenerateHelp_Click(sender As Object, e As EventArgs) Handles btnPropGenerateHelp.Click
        'Открывает диалог генерации файла помощи
        Dim dRes As DialogResult = Windows.Forms.DialogResult.No
        Dim fPath As String = ""
        If txtPropHelp.Text.Length > 0 Then
            fPath = GetHelpPath(txtPropHelp.Text)
            If My.Computer.FileSystem.FileExists(fPath) Then
                dRes = MessageBox.Show("К свойству уже присоединен файл помощи. Редактировать его?", "Matew Quest", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If dRes = Windows.Forms.DialogResult.Cancel Then Return
            End If
        End If

        If txtPropName.Text.Length = 0 Then
            MsgBox("Не введено имя свойства!")
            txtPropName.Focus()
            Return
        End If
        If txtPropDescription.Text.Length = 0 Then
            MsgBox("Не введено описание свойства!")
            txtPropDescription.Focus()
            Return
        End If
        'Перед началом сохраняем свойство
        SaveProperty()
        'Dim prop As MatewScript.PropertiesInfoType = Nothing
        'If mScript.mainClass(classId).Properties.TryGetValue(txtPropName.Text, prop) = False Then Return
        If dRes = Windows.Forms.DialogResult.No Then
            'Генерируем файл помощи
            dlgGenerateHelp.GenerateHelpForProperty(txtPropName.Text, classId, newPropertyData)
        Else
            'Открываем готовый файл
            dlgGenerateHelp.codeHelp.Text = My.Computer.FileSystem.ReadAllText(fPath, System.Text.Encoding.Default)
        End If
        dRes = dlgGenerateHelp.ShowDialog(Me)
        pnlProperty.Show()

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
                txtPropHelp.Text = fPath
            End If
        End If
    End Sub
End Class
