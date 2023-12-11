Public Class clsGateInward
    Public Const Formtype As String = "MIGTIN"
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim Header, AttachLine As SAPbouiCOM.DBDataSource
    Dim objRS, Recordset As SAPbobsCOM.Recordset
    Dim strSQL, strQuery As String
    Dim objMatrix, oattachMatrix As SAPbouiCOM.Matrix
    Dim objHeader As SAPbouiCOM.DBDataSource
    Dim objLine As SAPbouiCOM.DBDataSource
    Dim objlink As SAPbouiCOM.LinkedButton
    Dim objedit As SAPbouiCOM.EditText
    Dim FinDate(2) As String



    Public Sub LoadScreen()
        Try
            objForm = objAddOn.objUIXml.LoadScreenXML("Inward.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
            objMatrix = objForm.Items.Item("36").Specific
            oattachMatrix = objForm.Items.Item("mtxattach").Specific
            'oContLabMatrix = objForm.Items.Item("74").Specific
            objHeader = objForm.DataSources.DBDataSources.Item("@MIGTIN")
            objLine = objForm.DataSources.DBDataSources.Item("@MIGTIN1")
            AttachLine = objForm.DataSources.DBDataSources.Item("@MIGTIN2")
            InitForm(objForm.UniqueID)
            objAddOn.objGenFunc.setReport(Formtype, "Gate Inward", objForm.TypeCount)
            objForm.Visible = True
            ManageAttributes()
            objForm.Items.Item("52C1").Left = objForm.Items.Item("52A").Left + objForm.Items.Item("52A").Width + 3
            objForm.Items.Item("52C1").Top = objForm.Items.Item("52A").Top
            objForm.Items.Item("70").Specific.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("4").Specific.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("20").Specific.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("51").Specific.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("63").Click()
            'objMatrix.AutoResizeColumns()

            'GE_Inward_GRPO_Draft = objAddOn.objGenFunc.getSingleValue("Select ""U_GEIGRPD"" from OADM")
            'If GE_Inward_GRPO_Draft = "N" Then
            '    If Not objAddOn.objApplication.Menus.Item("GT").SubMenus.Exists(clsGEToGRPO.Formtype.ToString) Then objAddOn.CreateMenu("", 3, "Gate Entry GRPO", SAPbouiCOM.BoMenuType.mt_STRING, clsGEToGRPO.Formtype, objAddOn.objApplication.Menus.Item("GT"))
            'Else
            '    If objAddOn.objApplication.Menus.Item("GT").SubMenus.Exists(clsGEToGRPO.Formtype.ToString) Then objAddOn.objApplication.Menus.Item("GT").SubMenus.RemoveEx(clsGEToGRPO.Formtype.ToString)
            'End If

            '********************** Dynamic UDF Creation in Line Level of Matrix **************************************
            'objAddOn.objApplication.SendKeys("^+(S)")

            'objAddOn.objApplication.SendKeys("1")
            'objAddOn.objApplication.SendKeys("(^+B)")
            'objAddOn.objApplication.SendKeys("1")
            'If objAddOn.HANA = True Then
            '    strSQL = "Select ""USERID"",""TPLId"" from OUSR Where ""USER_CODE""='" & objAddOn.objCompany.UserName & "'"
            '    strQuery = "Select '@' || ""SonName"" ""TableName"" from UDO1 Where ""Code"" = '" & objForm.BusinessObject.Type & "'"
            'Else
            '    strSQL = "Select USERID,TPLId from OUSR Where USER_CODE='" & objAddOn.objCompany.UserName & "'"
            '    strQuery = "Select '@' + SonName TableName from UDO1 Where Code = '" & objForm.BusinessObject.Type & "'"
            'End If

            'Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Recordset.DoQuery(strQuery)

            'Dim Table_Matrix As Dictionary(Of String, String) = New Dictionary(Of String, String)()
            'Dim MatrixIDs As List(Of String) = New List(Of String)()
            'Dim MatID As String
            'Table_Matrix.Add("36", "@MIGTIN1")
            ''Table_Matrix.Add("72", "@MIGTIN1")
            'Table_Matrix.Add("mtxattach", "@MIGTIN2")
            'If Recordset.RecordCount > 0 Then
            '    For i As Integer = 0 To Recordset.RecordCount - 1
            '        If Not Table_Matrix.ContainsValue(Convert.ToString(Recordset.Fields.Item("TableName").Value)) Then Continue For
            '        For Each pair In Table_Matrix
            '            If pair.Value = Convert.ToString(Recordset.Fields.Item("TableName").Value) Then
            '                MatrixIDs.Add(pair.Key)
            '            End If
            '        Next
            '        MatID = String.Join(",", MatrixIDs)
            '        objAddOn.Create_Dynamic_LineTable_UDF(objForm, Convert.ToString(Recordset.Fields.Item("TableName").Value), objForm.TypeEx, String.Format("'{0}'", String.Join("','", MatrixIDs)))
            '        'strSQL = Table_Matrix(Convert.ToString(Recordset.Fields.Item("TableName").Value))
            '        Recordset.MoveNext()
            '    Next
            'End If
            ''********************** Dynamic UDF END **************************************
            ''******Header UDF Hiding using Form Preferences***************
            'objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'objRS.DoQuery(strSQL)
            'objAddOn.Update_UserFormSettings_UDF(objForm, "-" + objForm.TypeEx, Convert.ToInt32(objRS.Fields.Item("USERID").Value), Convert.ToInt32(objRS.Fields.Item("TPLId").Value))


        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Load Screen: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pval.ItemUID = "" Then Exit Sub
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objCombo = objForm.Items.Item("8").Specific
            'objAddOn.objApplication.MessageBox("Cont Mat Count: " & oContLabMatrix.VisualRowCount & "Mat Count: " & objMatrix.VisualRowCount)
            If pval.BeforeAction Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pval.ItemUID = "10" Then
                            If objCombo.Selected.Value = "PO" Or objCombo.Selected.Value = "SP" Then 'Or objCombo.Selected.Value = "GR"
                                If objAddOn.HANA Then
                                    ChooseFromList_Filteration(FormUID, "BP_CFL", "CardType", "S", "select distinct ""CardCode"" from OPOR where ""DocStatus""='O'")
                                Else
                                    ChooseFromList_Filteration(FormUID, "BP_CFL", "CardType", "S", "select distinct CardCode from OPOR where DocStatus='O'")
                                End If
                            ElseIf objCombo.Selected.Value = "DR" Then
                                If objAddOn.HANA Then
                                    ChooseFromList_Filteration(FormUID, "BP_CFL", "CardType", "C", "select distinct ""CardCode"" from ODLN where ""DocStatus""='O' and ""U_TransTyp""='RDC'")
                                Else
                                    ChooseFromList_Filteration(FormUID, "BP_CFL", "CardType", "C", "select distinct CardCode from ODLN where DocStatus='O' and U_TransTyp='RDC'")
                                End If
                            ElseIf objCombo.Selected.Value = "SR" Then 'objCombo.Selected.Value = "WI" Or objCombo.Selected.Value = "SC" Or objCombo.Selected.Value = "DR" Or objCombo.Selected.Value = "IN"
                                If objAddOn.HANA Then
                                    ChooseFromList_Filteration(FormUID, "BP_CFL", "CardType", "C", "select distinct ""CardCode"" from OINV")
                                Else
                                    ChooseFromList_Filteration(FormUID, "BP_CFL", "CardType", "C", "select distinct CardCode from OINV")
                                End If
                                'ElseIf objCombo.Selected.Value = "MI" Then
                            Else
                                Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("BP_CFL")
                                Dim oEmptyConds As New SAPbouiCOM.Conditions
                                oCFL.SetConditions(oEmptyConds)
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        If pval.ItemUID = "36" And (pval.ColUID = "9" Or pval.ColUID = "29A") Then ' line total calculation
                            If objForm.Items.Item("8").Specific.Selected.Value = "SP" Then
                                If CDbl(objMatrix.Columns.Item("29A").Cells.Item(pval.Row).Specific.value) > CDbl(objMatrix.Columns.Item("12").Cells.Item(pval.Row).Specific.value) Then
                                    objAddOn.objApplication.SetStatusBarMessage("Service Price exceeds Line Total", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    BubbleEvent = False
                                End If
                            Else
                                If Not QtyValidation(FormUID, pval.Row) Then
                                    BubbleEvent = False
                                End If
                            End If

                            If objMatrix.Columns.Item("4").Cells.Item(pval.Row).Specific.String = "" Then Exit Sub
                            If objMatrix.Columns.Item("9").Cells.Item(pval.Row).Specific.String = "" Then objMatrix.Columns.Item("9").Cells.Item(pval.Row).Click() : Exit Sub
                            If CDbl(objMatrix.Columns.Item("9").Cells.Item(pval.Row).Specific.String) <= 0 Then
                                objAddOn.objApplication.StatusBar.SetText("Value in ""Quantity"" cannot be zero.  Line: " & pval.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                objMatrix.Columns.Item("9").Cells.Item(pval.Row).Specific.String = CDbl(1)
                                'objMatrix.Columns.Item("9").Cells.Item(pval.Row).Click() : BubbleEvent = False : Exit Sub
                            End If
                        ElseIf pval.ItemUID = "36" And pval.ColUID = "16" Then 'Contract Matrix
                            objAddOn.objGenFunc.Matrix_Addrow(objMatrix, "16", "0")
                        End If
                        If (pval.ItemUID = "8" Or pval.ItemUID = "10") And pval.ItemChanged = True Then
                            If objMatrix.VisualRowCount > 0 Then objMatrix.Clear() : If objForm.Items.Item("10").Specific.String = "" Then objForm.Items.Item("12").Specific.String = ""
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pval.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If Not Validate(FormUID) Then
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ItemUID = "52C" Or pval.ItemUID = "52C1" Then
                            If objCombo.Selected.Value = "PO" Or objCombo.Selected.Value = "SP" Then
                                'Dim TEntry As String = objAddOn.objGenFunc.getSingleValue("Select 1 from ODRF where ""ObjType""='20' and ifnull(""DocStatus"",'')='O' and ""DocEntry"" in (" & objForm.Items.Item("52B").Specific.String & ")")
                                'objlink = objForm.Items.Item("52C").Specific
                                'objlink.LinkedObject = "-1"
                                'If TEntry <> "" Then
                                '    objlink.LinkedObject = "112"
                                '    objForm.Items.Item("52A").Specific.Caption = "Target Draft"
                                'Else
                                '    objlink.LinkedObject = "20"
                                '    objForm.Items.Item("52A").Specific.Caption = "Target DocEntry"
                                'End If
                                CreateMySimpleForm("ViewData", "Goods Receipt PO List", "ODRF", "OPDN", "20")
                            ElseIf objCombo.Selected.Value = "MI" Then
                                CreateMySimpleForm("ViewData", "Goods Receipt List", "ODRF", "OIGN", "59")
                            ElseIf objCombo.Selected.Value = "SR" Then
                                CreateMySimpleForm("ViewData", "A/R Credit Memo List", "ODRF", "ORIN", "14")
                            ElseIf objCombo.Selected.Value = "DR" Then
                                CreateMySimpleForm("ViewData", "Return List", "ODRF", "ORDN", "16")
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        Dim ColItem As SAPbouiCOM.Column
                        If pval.ItemUID = "36" Then
                            If pval.ColUID = "2" Then
                                ColItem = objMatrix.Columns.Item("2")
                                objlink = ColItem.ExtendedObject
                                If objMatrix.Columns.Item("1A").Cells.Item(pval.Row).Specific.String = "13" Then 'A/R Invoice
                                    objlink.LinkedObjectType = "13"
                                ElseIf objMatrix.Columns.Item("1A").Cells.Item(pval.Row).Specific.String = "15" Then 'Delivery
                                    objlink.LinkedObjectType = "15"
                                ElseIf objMatrix.Columns.Item("1A").Cells.Item(pval.Row).Specific.String = "22" Then 'Purchase Order
                                    objlink.LinkedObjectType = "22"
                                ElseIf objMatrix.Columns.Item("1A").Cells.Item(pval.Row).Specific.String = "67" Then 'Inventory Transfer
                                    objlink.LinkedObjectType = "67"
                                Else
                                    BubbleEvent = False
                                End If
                            ElseIf (pval.ColUID = "16" Or pval.ColUID = "22" Or pval.ColUID = "23" Or pval.ColUID = "24") Then
                                BubbleEvent = False
                                Dim objtempform As SAPbouiCOM.Form
                                'strSQL = oContLabMatrix.Columns.Item(pval.ColUID).DataBind.Alias
                                If pval.ColUID = "16" Then
                                    strSQL = objLine.GetValue("U_LabID", pval.Row - 1)
                                ElseIf pval.ColUID = "22" Then
                                    strSQL = objLine.GetValue("U_SupID", pval.Row - 1)
                                ElseIf pval.ColUID = "23" Then
                                    strSQL = objLine.GetValue("U_TimeKpID", pval.Row - 1)
                                ElseIf pval.ColUID = "24" Then
                                    strSQL = objLine.GetValue("U_SecID", pval.Row - 1)
                                End If
                                If strSQL <> "" Then
                                    objAddOn.objApplication.Menus.Item("3590").Activate()  ' Employee Master screen
                                    objtempform = objAddOn.objApplication.Forms.ActiveForm
                                    objtempform.Freeze(True)
                                    objtempform.Items.Item("257000006").Specific.String = strSQL '33
                                    objtempform.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    objtempform.Freeze(False)
                                End If
                            End If

                            'ElseIf pval.ItemUID = "36" And (pval.ColUID = "2" Or pval.ColUID = "8" Or pval.ColUID = "9" Or pval.ColUID = "10") Then

                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                        'If pval.InnerEvent = True And pval.ItemUID = "1" Then BubbleEvent = False : Exit Sub
                        If pval.ItemUID = "1" And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            objCombo = objForm.Items.Item("8").Specific
                            'If objCombo.Selected.Value = "CL" Then
                            '    oContLabMatrix.Item.Visible = True
                            '    objMatrix.Item.Visible = False
                            'Else
                            '    oContLabMatrix.Item.Visible = False
                            '    objMatrix.Item.Visible = True
                            'End If
                            'If objForm.PaneLevel = 1 Then objAddOn.objGenFunc.RemoveLastrow(objMatrix, "4")
                            'If objForm.PaneLevel = 2 Then objAddOn.objGenFunc.RemoveLastrow(oContLabMatrix, "1")
                        End If
                        If (pval.ItemUID = "51" Or pval.ItemUID = "25") And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            BubbleEvent = False
                        End If
                        If pval.ItemUID = "51B" Then BubbleEvent = False
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub 'objForm.Items.Item("25").Specific.Selected.Value = "O" Or
                        'If pval.ItemUID = "21" Then
                        '    BubbleEvent = False
                        'End If
                        If pval.ItemUID = "21" Then '(pval.ItemUID <> "51" Or pval.ItemUID <> "55") Or
                            BubbleEvent = False
                        End If

                        If pval.ItemUID = "36" And pval.ColUID = "9" Then
                            If objForm.Items.Item("25").Specific.Selected.Value = "C" Then
                                BubbleEvent = False
                            End If
                        End If
                        If pval.CharPressed = 9 Or pval.CharPressed = 8 Then Exit Sub
                        If pval.ItemUID = "36" And pval.ColUID = "21B" Or pval.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_CTRL Then
                            BubbleEvent = False
                        End If
                        'Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        '    If pval.ItemUID = "8" Then
                        '        objCombo = objForm.Items.Item("8").Specific
                        '        If pval.ItemChanged = False Then Exit Sub
                        '        objForm.Freeze(True)

                        '        'If objCombo.Selected.Value = "CL" Then
                        '        '    If oContLabMatrix.VisualRowCount > 0 And oContLabMatrix.Columns.Item("1").Cells.Item(1).Specific.string <> "" Then
                        '        '        If objAddOn.objApplication.MessageBox("Changes will cause the data to be deleted. Continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                        '        '        oContLabMatrix.Clear()
                        '        '    End If
                        '        'Else
                        '        '    'If objForm.PaneLevel = 1 Then Exit Sub
                        '        '    If objMatrix.VisualRowCount > 0 Then
                        '        '        If oContLabMatrix.Item.Visible = True Then If objAddOn.objApplication.MessageBox("Changes will cause the data to be deleted. Continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                        '        '        objMatrix.Clear()
                        '        '    End If
                        '        'End If
                        '        objForm.Freeze(False)
                        '    End If

                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pval.ItemUID = "37" Then ' copy From
                            If objForm.Items.Item("25").Specific.Selected.Value = "O" Then CopyFrom(FormUID)
                        ElseIf pval.ItemUID = "38" Then
                            CopyTo(FormUID)
                        ElseIf pval.ItemUID = "1" And pval.ActionSuccess And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            InitForm(FormUID)
                        ElseIf pval.ItemUID = "36" Then
                            If pval.ColUID = "21A" Then
                                If objMatrix.Columns.Item("21A").Cells.Item(pval.Row).Specific.Checked = False Then
                                    objMatrix.Columns.Item("21B").Cells.Item(pval.Row).Specific.String = ""
                                    objMatrix.CommonSetting.SetCellEditable(pval.Row, 25, True)
                                End If
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        objMatrix.AutoResizeColumns()
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If pval.ItemUID = "36" And (pval.ColUID = "9" Or pval.ColUID = "11") Then ' line total calculation
                            LineTotalCalc(FormUID, pval.Row)
                        End If
                        If pval.ItemUID = "23" Then 'DocDate
                            If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                            objedit = objForm.Items.Item("23").Specific
                            Try
                                'If objAddOn.HANA Then
                                '    strSQL = objAddOn.objGenFunc.getSingleValue("Select ""FinancYear"" ""SelYear"" from OACP Where ""PeriodCat""=(Select Top 1 ""Category"" From OFPR Where ""Indicator""=(Select ""Indicator"" from NNM1 Where ""Series""='" & objForm.Items.Item("20").Specific.Selected.Value & "'))")
                                'Else
                                '    strSQL = objAddOn.objGenFunc.getSingleValue("Select FinancYear SelYear from OACP Where PeriodCat=(Select Top 1 Category From OFPR Where Indicator=(Select Indicator from NNM1 Where Series='" & objForm.Items.Item("20").Specific.Selected.Value & "'))")
                                'End If
                                If FinDate(1) = "" Then Exit Sub
                                If FinDate(0) <> FinDate(1) Then 'Year(Now)
                                    objAddOn.objApplication.MessageBox("Newly entered posting date relates to another posting period. Do you want to Continue?", 2, "Yes", "No")
                                    objCombo = objForm.Items.Item("20").Specific
                                    For i As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                                        objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next
                                    If objAddOn.HANA Then
                                        strSQL = "select ""Series"",""SeriesName"" From NNM1 where ""ObjectCode""='" & Formtype & "' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & objedit.Value & "' between ""F_RefDate"" and ""T_RefDate"")  "
                                    Else
                                        strSQL = "select Series,SeriesName From NNM1 where ObjectCode='" & Formtype & "' and Indicator=(select Top 1 Indicator  from OFPR where '" & objedit.Value & "' between F_RefDate and T_RefDate)  "
                                    End If
                                    objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    objRS.DoQuery(strSQL)
                                    If objRS.RecordCount > 0 Then
                                        For Rec As Integer = 0 To objRS.RecordCount - 1
                                            objCombo.ValidValues.Add(objRS.Fields.Item(0).Value.ToString, objRS.Fields.Item(1).Value.ToString)
                                            objRS.MoveNext()
                                        Next
                                    End If
                                    If objCombo.ValidValues.Count > 0 Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                End If
                            Catch ex As Exception
                                'objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pval.ItemUID = "10" Or (pval.ItemUID = "36" And pval.ColUID = "16" Or pval.ColUID = "22" Or pval.ColUID = "23" Or pval.ColUID = "24") Then 'Partyid
                            ChooseFromList(FormUID, pval)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pval.ItemUID = "51" Then
                            objCombo = objForm.Items.Item("51").Specific
                            If objCombo.Selected.Value = "2" Then
                                objForm.Items.Item("53").Specific.String = DateTime.Now.ToString("HH:mm")
                                objForm.Items.Item("55").Click()
                            End If
                        ElseIf pval.ItemUID = "8" Then
                            objCombo = objForm.Items.Item("8").Specific
                            objForm.Items.Item("8A").Specific.String = GetTransaction_Type(FormUID, objCombo.Selected.Value)
                            FieldSetup(FormUID, objCombo.Selected.Value)
                            If objCombo.Selected.Value = "PO" Then
                                objForm.Items.Item("51A").Visible = True
                                objForm.Items.Item("51B").Visible = True

                                'objForm.PaneLevel = 1
                                'oContLabMatrix.Item.Visible = False
                                'objMatrix.Item.Visible = True
                            ElseIf objCombo.Selected.Value = "CL" Then
                                '    'objForm.PaneLevel = 2
                                '    oContLabMatrix.Item.Visible = True
                                '    objMatrix.Item.Visible = False
                                '    objForm.Settings.MatrixUID = "74"
                                objAddOn.objGenFunc.Matrix_Addrow(objMatrix, "16", "0")
                            Else
                                objForm.Items.Item("51A").Visible = False
                                objForm.Items.Item("51B").Visible = False
                                'objForm.PaneLevel = 1
                                'oContLabMatrix.Item.Visible = False
                                'objMatrix.Item.Visible = True
                                'objForm.Settings.MatrixUID = "36"
                            End If

                            If pval.ItemChanged = True Then objForm.Items.Item("10").Specific.String = "" : objForm.Items.Item("12").Specific.String = ""
                        ElseIf pval.ItemUID = "20" Then
                            If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                            objCombo = objForm.Items.Item("20").Specific
                            objHeader = objForm.DataSources.DBDataSources.Item("@MIGTIN")
                            objHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetDocNum(Formtype, CInt(objForm.Items.Item("20").Specific.Selected.value)))
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        If pval.ItemUID = "mtxattach" Then
                            If pval.ActionSuccess Then objAddOn.objGenFunc.OpenAttachment(oattachMatrix, AttachLine, pval.Row)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pval.ItemUID = "36" Then
                            objMatrix.SelectRow(pval.Row, True, False)
                        ElseIf pval.ItemUID = "Btnbrowse" Then
                            If objForm.Items.Item(pval.ItemUID).Enabled = False Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                            If objAddOn.objGenFunc.SetAttachMentFile(objForm, objHeader, oattachMatrix, AttachLine) = False Then
                                BubbleEvent = False
                            End If
                            If oattachMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) = -1 Then
                                objForm.Items.Item("Btndisp").Enabled = False
                                objForm.Items.Item("Btndel").Enabled = False
                            End If
                        ElseIf pval.ItemUID = "Btndisp" Then
                            If objForm.Items.Item(pval.ItemUID).Enabled = False Then Exit Sub
                            If pval.ActionSuccess Then objAddOn.objGenFunc.OpenAttachment(oattachMatrix, AttachLine, pval.Row)
                        ElseIf pval.ItemUID = "Btndel" Then
                            If objForm.Items.Item(pval.ItemUID).Enabled = False Then Exit Sub
                            If pval.ActionSuccess Then
                                objAddOn.objGenFunc.DeleteRowAttachment(objForm, oattachMatrix, AttachLine, oattachMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder))
                            End If
                        ElseIf pval.ItemUID = "mtxattach" Then
                            oattachMatrix.SelectRow(pval.Row, True, False)
                            If pval.Row > 0 Then
                                If oattachMatrix.IsRowSelected(pval.Row) Then
                                    objForm.Items.Item("Btndisp").Enabled = True
                                    objForm.Items.Item("Btndel").Enabled = True
                                End If
                            End If
                            'ElseIf pval.ItemUID = "63" Then
                            'objCombo = objForm.Items.Item("8").Specific
                            'If objCombo.Selected.Value = "CL" Then
                            '    'objForm.PaneLevel = 2
                            '    oContLabMatrix.Item.Visible = True
                            '    objMatrix.Item.Visible = False
                            '    objForm.Settings.MatrixUID = "74" : oContLabMatrix.AutoResizeColumns()
                            'Else
                            '    'objForm.PaneLevel = 1
                            '    'oContLabMatrix.Item.Visible = False
                            '    'objMatrix.Item.Visible = True
                            '    objForm.Settings.MatrixUID = "36" : objMatrix.AutoResizeColumns()
                            'End If
                        ElseIf pval.ItemUID = "64" Then
                            objForm.PaneLevel = 3
                            objForm.Settings.MatrixUID = "mtxattach"
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pval.ItemUID = "36" Then
                            Dim ColID As Integer = objMatrix.GetCellFocus().ColumnIndex
                            If pval.CharPressed = 38 And pval.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then  'up
                                objMatrix.SetCellFocus(pval.Row - 1, ColID)
                                objMatrix.SelectRow(pval.Row - 1, True, False)
                            ElseIf pval.CharPressed = 40 And pval.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'down
                                objMatrix.SetCellFocus(pval.Row + 1, ColID)
                                objMatrix.SelectRow(pval.Row + 1, True, False)
                            End If
                            'If pval.CharPressed <> 9 Or pval.CharPressed <> 8 Or pval.CharPressed <> 36 Then Exit Sub
                            'If pval.ItemUID = "36" And pval.ColUID = "21B" Or pval.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_CTRL Then
                            '    BubbleEvent = False
                            'End If
                        End If
                        If pval.ItemUID = "36" And pval.ColUID = "21B" And (pval.CharPressed = 37 Or pval.CharPressed = 39 Or pval.CharPressed = 9) Then
                            If objMatrix.Columns.Item("21B").Cells.Item(pval.Row).Specific.string = "" And objMatrix.Columns.Item("21A").Cells.Item(pval.Row).Specific.Checked = True Then
                                objAddOn.objMatDetails.LoadScreen(FormUID, pval.Row)
                            Else
                                Exit Sub
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        If pval.ItemUID = "23" Then 'DocDate
                            If pval.ItemChanged = True And pval.InnerEvent = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If FinDate(1) = "" Then FinDate(1) = Year(Now) Else FinDate(1) = FinDate(0)
                                objedit = objForm.Items.Item("23").Specific
                                Dim DocDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                FinDate(0) = DocDate.Year
                                If FinDate(0) = FinDate(1) Then FinDate(1) = ""
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If pval.ItemUID = "36" And pval.ColUID = "21B" Then
                            objMatrix = objForm.Items.Item("36").Specific
                            objAddOn.objMatDetails.LoadScreen(FormUID, pval.Row, objMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.String)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                        If pval.ItemUID = "36" And pval.ColUID = "21B" Then
                            objForm.EnableMenu("773", False)
                        Else
                            objForm.EnableMenu("773", True)
                        End If
                End Select

            End If
        Catch ex As Exception
            'objForm.Freeze(False)
        Finally
            objForm.Freeze(False)
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            objMatrix = objForm.Items.Item("36").Specific
            objCombo = objForm.Items.Item("8").Specific
            If BusinessObjectInfo.BeforeAction Then
                Try
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If Validate(BusinessObjectInfo.FormUID) Then
                                If (objCombo.Selected.Value = "PO" Or objCombo.Selected.Value = "SP") And GE_Inward_GRPO_Draft = "Y" Then
                                    If Create_GoodsReceipt_PO(BusinessObjectInfo.FormUID) = False Then
                                        BubbleEvent = False
                                        objAddOn.objApplication.StatusBar.SetText("GRPO Draft Not Created...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                                If objCombo.Selected.Value = "MI" Then
                                    If Create_GoodsReceipt(BusinessObjectInfo.FormUID) = False Then
                                        BubbleEvent = False
                                        objAddOn.objApplication.StatusBar.SetText("Goods Receipt Draft Not Created...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                ElseIf objCombo.Selected.Value = "SR" Then
                                    If Create_ARCreditMemo(BusinessObjectInfo.FormUID) = False Then
                                        BubbleEvent = False
                                        objAddOn.objApplication.StatusBar.SetText("A/R CreditMemo Draft Not Created...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                ElseIf objCombo.Selected.Value = "DR" Then
                                    If Create_Delivery_Return(BusinessObjectInfo.FormUID) = False Then
                                        BubbleEvent = False
                                        objAddOn.objApplication.StatusBar.SetText("Return Draft Not Created...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                            Else
                                BubbleEvent = False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                            'If Not Validate(BusinessObjectInfo.FormUID) Then
                            '    BubbleEvent = False
                            'End If
                    End Select

                Catch ex As Exception
                    BubbleEvent = False
                End Try
            Else
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If (objCombo.Selected.Value = "PO" Or objCombo.Selected.Value = "SP") And GE_Inward_GRPO_Draft = "Y" And BusinessObjectInfo.ActionSuccess = True Then
                            objHeader = objForm.DataSources.DBDataSources.Item("@MIGTIN")
                            If objAddOn.HANA Then
                                strSQL = "Update ""@MIGTIN"" Set ""U_Prostat""='3' where ""DocEntry""='" & objHeader.GetValue("DocEntry", 0) & "'"
                                objRS.DoQuery(strSQL)
                                strSQL = "Update ODRF Set ""U_gever""='" & objHeader.GetValue("DocEntry", 0) & "' where ""ObjType""='20' and ""DocEntry""=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                                strSQL = "Update T1 Set T1.""U_geentry""='" & objHeader.GetValue("DocEntry", 0) & "' from ODRF T0 join DRF1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""ObjType""='20' and T0.""DocEntry""=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                            Else
                                strSQL = "Update [@MIGTIN] Set U_Prostat='3' where DocEntry='" & objHeader.GetValue("DocEntry", 0) & "'"
                                objRS.DoQuery(strSQL)
                                strSQL = "Update ODRF Set U_gever='" & objHeader.GetValue("DocEntry", 0) & "' where ObjType='20' and DocEntry=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                                strSQL = "Update T1 Set T1.U_geentry='" & objHeader.GetValue("DocEntry", 0) & "' from ODRF T0 join DRF1 T1 on T0.DocEntry=T1.DocEntry where T0.ObjType='20' and T0.DocEntry=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                            End If
                        End If
                        If objCombo.Selected.Value = "SR" Then 'A/R Credit Memo
                            If objAddOn.HANA Then
                                strSQL = "Update ODRF Set ""U_gever""='" & objHeader.GetValue("DocEntry", 0) & "' where ""ObjType""='14' and ""DocEntry""=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                                strSQL = "Update T1 Set T1.""U_geentry""='" & objHeader.GetValue("DocEntry", 0) & "' from ODRF T0 join DRF1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""ObjType""='14' and T0.""DocEntry""=" & objForm.Items.Item("52B").Specific.String & " "
                                objRS.DoQuery(strSQL)
                            Else
                                strSQL = "Update ODRF Set U_gever='" & objHeader.GetValue("DocEntry", 0) & "' where ObjType='14' and DocEntry=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                                strSQL = "Update T1 Set T1.U_geentry='" & objHeader.GetValue("DocEntry", 0) & "' from ODRF T0 join DRF1 T1 on T0.DocEntry=T1.DocEntry where T0.ObjType='14' and T0.DocEntry=" & objForm.Items.Item("52B").Specific.String & " "
                                objRS.DoQuery(strSQL)
                            End If

                        ElseIf objCombo.Selected.Value = "DR" Then
                            If objAddOn.HANA Then
                                strSQL = "Update ODRF Set ""U_gever""='" & objHeader.GetValue("DocEntry", 0) & "' where ""ObjType""='16' and ""DocEntry""=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                                strSQL = "Update T1 Set T1.""U_geentry""='" & objHeader.GetValue("DocEntry", 0) & "' from ODRF T0 join DRF1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""ObjType""='16' and T0.""DocEntry""=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                            Else
                                strSQL = "Update ODRF Set U_gever='" & objHeader.GetValue("DocEntry", 0) & "' where ObjType='16' and DocEntry=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                                strSQL = "Update T1 Set T1.U_geentry='" & objHeader.GetValue("DocEntry", 0) & "' from ODRF T0 join DRF1 T1 on T0.DocEntry=T1.DocEntry where T0.ObjType='16' and T0.DocEntry=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                            End If

                        ElseIf objCombo.Selected.Value = "MI" Then
                            If objAddOn.HANA Then
                                strSQL = "Update ODRF Set ""U_gever""='" & objHeader.GetValue("DocEntry", 0) & "' where ""ObjType""='59' and ""DocEntry""=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                                strSQL = "Update T1 Set T1.""U_geentry""='" & objHeader.GetValue("DocEntry", 0) & "' from ODRF T0 join DRF1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""ObjType""='59' and T0.""DocEntry""=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                            Else
                                strSQL = "Update ODRF Set U_gever='" & objHeader.GetValue("DocEntry", 0) & "' where ObjType='59' and DocEntry=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                                strSQL = "Update T1 Set T1.U_geentry='" & objHeader.GetValue("DocEntry", 0) & "' from ODRF T0 join DRF1 T1 on T0.DocEntry=T1.DocEntry where T0.ObjType='59' and T0.DocEntry=" & objForm.Items.Item("52B").Specific.String & ""
                                objRS.DoQuery(strSQL)
                            End If
                        ElseIf objCombo.Selected.Value = "CL" Then
                            strQuery = ""
                            For index = 1 To objMatrix.VisualRowCount
                                If objMatrix.Columns.Item("21A").Cells.Item(index).Specific.Checked = True Then
                                    If strQuery = "" Then strQuery = objMatrix.Columns.Item("21B").Cells.Item(index).Specific.String Else strQuery = strQuery + "," + objMatrix.Columns.Item("21B").Cells.Item(index).Specific.String
                                    'strQuery = strQuery + objMatrix.Columns.Item("21B").Cells.Item(index).Specific.String + ","
                                End If
                            Next
                            'strQuery = strQuery.Remove(strQuery.Length - 1, 1)
                            If objAddOn.HANA Then
                                strSQL = "Update ""@MIMATDET"" Set ""U_GIEntry""='" & objHeader.GetValue("DocEntry", 0) & "' Where ""DocEntry"" in (" & strQuery & ")"
                                objRS.DoQuery(strSQL)
                            Else
                                strSQL = "Update [@MIMATDET] Set U_GIEntry='" & objHeader.GetValue("DocEntry", 0) & "' Where DocEntry in (" & strQuery & ")"
                                objRS.DoQuery(strSQL)
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        objCombo = objForm.Items.Item("8").Specific
                        FieldSetup(BusinessObjectInfo.FormUID, objCombo.Selected.Value)
                        'If objCombo.Selected.Value = "CL" Then
                        '    objForm.PaneLevel = 2
                        objAddOn.objGenFunc.Matrix_Addrow(objMatrix, "16", "0")
                        'Else
                        '    objForm.PaneLevel = 1
                        'End If
                        Dim Fieldsize As Size = TextRenderer.MeasureText(objForm.Items.Item("12").Specific.String, New Font("Arial", 12.0F))
                        If Fieldsize.Width <= 140 Then
                            objForm.Items.Item("12").Width = objForm.Items.Item("10").Width '140
                        Else
                            objForm.Items.Item("12").Width = Fieldsize.Width
                        End If
                        objMatrix.AutoResizeColumns()
                        objCombo = objForm.Items.Item("51B").Specific
                        objForm.Items.Item("51B").Enabled = True
                        If objCombo.Selected Is Nothing Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index) : objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE : objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If objMatrix.Columns.Item("14").Cells.Item(i).Specific.string = "C" Then
                                objMatrix.CommonSetting.SetCellEditable(i, 7, False)
                                objMatrix.CommonSetting.SetCellEditable(i, 11, False)
                                objMatrix.CommonSetting.SetCellEditable(i, 15, False)
                            Else
                                objMatrix.CommonSetting.SetCellEditable(i, 7, True)
                                objMatrix.CommonSetting.SetCellEditable(i, 11, True)
                                objMatrix.CommonSetting.SetCellEditable(i, 15, True)
                            End If
                        Next

                        If objForm.Items.Item("25").Specific.Selected.Value = "C" Then
                            objForm.Items.Item("51").Enabled = True
                            objForm.Items.Item("53").Enabled = True
                            objForm.Items.Item("37").Enabled = False
                            objMatrix.Item.Enabled = False
                            If objForm.Items.Item("51").Specific.Selected.Value = "2" And objForm.Items.Item("53").Specific.String <> "" Then
                                objForm.Items.Item("51").Enabled = False
                                objForm.Items.Item("53").Enabled = False
                            End If
                        Else
                            If objForm.Items.Item("51B").Specific.Selected.Value = "0" Then objMatrix.Item.Enabled = True Else objMatrix.Item.Enabled = False
                        End If

                End Select
            End If
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            objMatrix = objForm.Items.Item("36").Specific
            Select Case pVal.MenuUID
                Case "1284"
                    If pVal.BeforeAction = True Then
                        If objAddOn.objApplication.MessageBox("Cancelling of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                        objCombo = objForm.Items.Item("8").Specific
                        If objCombo.Selected.Value = "PO" Then
                            Remove_Draft_Transaction(objForm.UniqueID, SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes, objForm.Items.Item("52B").Specific.String)
                        End If
                    End If
                Case "1286"
                    If pVal.BeforeAction = True Then
                        If objAddOn.objApplication.MessageBox("Closing of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                    End If
                Case "1282"
                    If pVal.BeforeAction = False Then InitForm(objAddOn.objApplication.Forms.ActiveForm.UniqueID)
                    objMatrix.Item.Enabled = True
                    'Case "1289"
                    '    If pVal.BeforeAction = False Then Me.UpdateMode()
                Case "1293"  'delete Row
                    'For i As Integer = objMatrix.VisualRowCount To 1 Step -1
                    '    objMatrix.Columns.Item("0").Cells.Item(i).Specific.String = i
                    'Next
                    If pVal.BeforeAction = True Then
                        If objForm.ActiveItem = "36" Then
                            If objMatrix.VisualRowCount = 1 Then BubbleEvent = False : Exit Sub
                            'DeleteRow(objMatrix, "@MIGTIN1")
                        ElseIf objForm.ActiveItem = "74" Then
                            'If oContLabMatrix.VisualRowCount = 1 Then BubbleEvent = False : Exit Sub
                            'DeleteRow(oContLabMatrix, "@MIGTIN1")
                        End If
                    Else
                        If objForm.PaneLevel = "1" Then
                            If objMatrix.VisualRowCount = 1 Then BubbleEvent = False : Exit Sub
                            DeleteRow(objMatrix, "@MIGTIN1")
                        ElseIf objForm.PaneLevel = "2" Then
                            'If oContLabMatrix.VisualRowCount = 1 Then BubbleEvent = False : Exit Sub
                            'DeleteRow(oContLabMatrix, "@MIGTIN1")
                        End If
                    End If


                Case "1281"
                    objMatrix.Item.Enabled = False
                    objForm.Items.Item("57").Enabled = True
                    objForm.Items.Item("37").Enabled = False
                    objForm.Items.Item("38").Enabled = False
                    objForm.Items.Item("8A").Enabled = True
                    objForm.Items.Item("Btnbrowse").Enabled = False
                    objForm.Items.Item("Btndisp").Enabled = False
                    objForm.Items.Item("Btndel").Enabled = False
                    objForm.ActiveItem = "21"

            End Select
        Catch ex As Exception
            'objAddOn.objApplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Public Sub InitForm(ByVal FormUID As String)
        Try
            LoadType(FormUID)
            LoadSeries(FormUID)
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("36").Specific
            objMatrix.Columns.Item("9").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("12").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Init Form: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Public Sub LoadSeries(ByVal FormUID As String)
        Try
            Dim StrDocNum
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objForm.Items.Item("55").Specific.String = "Created By " & objAddOn.objCompany.UserName & " on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")
            '-----------Load Branch---------------
            If BranchFlag = "Y" Then
                objCombo = objForm.Items.Item("70").Specific
                If objCombo.ValidValues.Count = 0 Then
                    If objAddOn.HANA Then
                        strSQL = "Select ""BPLId"",""BPLName"" from OBPL Where ""BPLId"" in (Select T0.""BPLId"" from OBPL T0 join USR6 T1 on T0.""BPLId""=T1.""BPLId"" where T1.""UserCode""='" & objAddOn.objCompany.UserName & "' and T0.""Disabled""<>'Y') Order by ""BPLName"""
                    Else
                        strSQL = "Select BPLId,BPLName from OBPL Where BPLId in (Select T0.BPLId from OBPL T0 join USR6 T1 on T0.BPLId=T1.BPLId where T1.UserCode='" & objAddOn.objCompany.UserName & "' and T0.Disabled<>'Y') Order by BPLName "
                    End If
                    objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objRS.DoQuery(strSQL)
                    'objCombo.ValidValues.Add("-1", "All")
                    While Not objRS.EoF
                        objCombo.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
                        objRS.MoveNext()
                    End While
                    Try
                        If DefaultBranch = "" Then
                            objAddOn.objApplication.Menus.Item("11010").Activate()
                            Dim tempmatrix As SAPbouiCOM.Matrix
                            tempmatrix = objAddOn.objApplication.Forms.ActiveForm.Items.Item("1320000003").Specific
                            DefaultBranch = tempmatrix.Columns.Item("1320000005").Cells.Item(tempmatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder)).Specific.String
                            objAddOn.objApplication.Forms.ActiveForm.Close()
                        End If
                    Catch ex As Exception
                    End Try
                    objRS = Nothing
                End If
                If DefaultBranch = "" Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index) Else objCombo.Select(DefaultBranch, SAPbouiCOM.BoSearchKey.psk_ByDescription)
            Else
                objForm.Items.Item("70").Enabled = False
            End If
            '---------------- Load locations ------------
            objCombo = objForm.Items.Item("4").Specific
            If objCombo.ValidValues.Count = 0 Then
                If objAddOn.HANA Then
                    strSQL = "select ""Code"", ""Location"" from OLCT order by ""Location"" "
                Else
                    strSQL = "select Code, Location from OLCT order by Location "
                End If

                objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRS.DoQuery(strSQL)
                While Not objRS.EoF
                    objCombo.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
                    objRS.MoveNext()
                End While

                objRS = Nothing
            End If
            'objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            objCombo.Select(objCombo.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index)
            objForm.Items.Item("16").Specific.String = DateTime.Now.ToString("HH:mm") 'DateTime.Now.ToShortTimeString
            '----------------Load series --------------
            objCombo = objForm.Items.Item("20").Specific
            objCombo.ValidValues.LoadSeries(Formtype, SAPbouiCOM.BoSeriesMode.sf_Add)
            If objCombo.ValidValues.Count > 0 Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Try
                StrDocNum = objForm.BusinessObject.GetNextSerialNumber(Trim(objForm.Items.Item("20").Specific.Selected.value), objForm.BusinessObject.Type)
            Catch ex As Exception
                objAddOn.objApplication.MessageBox("To generate this document, first define the numbering series in the Administration module")
                Exit Sub
            End Try
            objHeader = objForm.DataSources.DBDataSources.Item("@MIGTIN")
            objHeader.SetValue("DocNum", 0, StrDocNum)
            'objHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetDocNum(Formtype, CInt(objForm.Items.Item("20").Specific.Selected.value)))
            objForm.Items.Item("23").Specific.String = "A" ' current date
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objCombo = objForm.Items.Item("8").Specific
                objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                objForm.Items.Item("8A").Specific.String = GetTransaction_Type(FormUID, objCombo.Selected.Value)
                If objCombo.Selected.Value = "PO" Then
                    objCombo = objForm.Items.Item("51B").Specific
                    objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                End If
            End If
            '------------ Load Security-------------
            objCombo = objForm.Items.Item("6").Specific
            If objCombo.ValidValues.Count = 0 Then
                If objAddOn.HANA Then
                    strSQL = "SELECT T0.""empID"", T0.""firstName"" || ' ' || T0.""lastName"" as ""empName"", T1.""Name"" FROM OHEM T0 INNER JOIN OUDP T1 ON T0.""dept"" = T1.""Code"" WHERE UPPER(T1.""Name"") ='SECURITY' ;"
                Else
                    strSQL = "SELECT T0.[empID], T0.[firstName] + ' ' + T0.[lastName] as empName, T1.[Name] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.[dept] = T1.[Code] WHERE UPPER(T1.[Name]) ='SECURITY'"
                End If

                objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRS.DoQuery(strSQL)
                While Not objRS.EoF
                    objCombo.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
                    objRS.MoveNext()
                End While
                objRS = Nothing
            End If
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Load Series: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LoadType(ByVal FormUID As String)
        Try

            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objCombo = objForm.Items.Item("8").Specific
            If objCombo.ValidValues.Count = 0 Then

                ''objCombo.ValidValues.Add("PO", "Purchase Order") '--
                ''objCombo.ValidValues.Add("SR", "Sales Return with Invoice") '--
                ''objCombo.ValidValues.Add("DR", "Customer Delivery Return") '--
                ''objCombo.ValidValues.Add("GR", "GRN")
                ''objCombo.ValidValues.Add("SC", "Sales Credit Memo")
                ''objCombo.ValidValues.Add("WI", "Sales Return without Invoice")
                ''objCombo.ValidValues.Add("DC", "Returnable DC")
                ''objCombo.ValidValues.Add("JR", "JobOrder Repair")
                ''objCombo.ValidValues.Add("SP", "Scrap Receipt")
                ''objCombo.ValidValues.Add("WM", "Without Process Material")
                ''objCombo.ValidValues.Add("RW", "Job Order Rework ")
                ''objCombo.ValidValues.Add("ST", "Stock Transfer")
                ''objCombo.ValidValues.Add("SO", "Service Order")
                ''objCombo.ValidValues.Add("JO", "Job Order Regular")
                ''objCombo.ValidValues.Add("JW", "Job Rework")
                ''objCombo.ValidValues.Add("CP", "Cash Purchase")
                ''objCombo.ValidValues.Add("MI", "Material Inward")
                ''objCombo.ValidValues.Add("HR", "Service Invoice HR")
                '''objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                objCombo.ValidValues.Add("PO", "Purchase Order")
                objCombo.ValidValues.Add("SR", "Sales Return")
                objCombo.ValidValues.Add("DR", "Returnable DC")
                objCombo.ValidValues.Add("MI", "Material Inward")
                objCombo.ValidValues.Add("CL", "Contract Labour")
                objCombo.ValidValues.Add("SP", "Service PO")
            End If

            objCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'objCombo = objForm.Items.Item("51B").Specific
            'If objCombo.ValidValues.Count = 0 Then
            '    objCombo.ValidValues.Add("0", "Open")
            '    objCombo.ValidValues.Add("1", "GE Canceled")
            '    objCombo.ValidValues.Add("2", "GE To GRPO Created")
            '    objCombo.ValidValues.Add("3", "GRPO Draft Created")
            '    objCombo.ValidValues.Add("4", "GRPO Created")
            '    objCombo.ValidValues.Add("5", "GRPO Canceled")
            '    objCombo.ValidValues.Add("6", "Closed")
            'End If

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Load Type: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ChooseFromList(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent)
        Try
            Dim CFLEvent As SAPbouiCOM.ChooseFromListEvent
            CFLEvent = pval
            Dim datatable As SAPbouiCOM.DataTable
            objLine = objForm.DataSources.DBDataSources.Item("@MIGTIN1")
            datatable = CFLEvent.SelectedObjects()

            If CFLEvent.ChooseFromListUID = "BP_CFL" Then
                Try
                    objHeader = objForm.DataSources.DBDataSources.Item("@MIGTIN")
                    objHeader.SetValue("U_partyid", 0, datatable.GetValue("CardCode", 0))
                    objHeader.SetValue("U_partynm", 0, datatable.GetValue("CardName", 0))
                Catch ex As Exception
                End Try
            ElseIf CFLEvent.ChooseFromListUID = "cfleid" Then
                'datatable = CFLEvent.SelectedObjects()
                Try
                    objForm.Freeze(True)
                    objMatrix.FlushToDataSource()
                    objLine.SetValue("U_LabID", pval.Row - 1, datatable.GetValue("empID", 0))
                    objLine.SetValue("U_LabName", pval.Row - 1, datatable.GetValue("firstName", 0) + " " + datatable.GetValue("lastName", 0)) ' datatable.GetValue("lastName", 0) + ", " + datatable.GetValue("firstName", 0))
                    objLine.SetValue("U_PFNo", pval.Row - 1, datatable.GetValue("U_PFNo", 0))
                    objLine.SetValue("U_ESINo", pval.Row - 1, datatable.GetValue("U_ESINo", 0))
                    strSQL = DateTime.Now.ToString("HH:mm").ToString.Replace(":", "")
                    'objLine.SetValue("U_InTime", pval.Row - 1, strSQL)

                    objMatrix.LoadFromDataSource()
                    objMatrix.Columns.Item("19").Cells.Item(pval.Row).Specific.value = strSQL ' DateTime.Now.ToString("HH:mm")
                    objMatrix.Columns.Item("16").Cells.Item(pval.Row).Click()
                    objForm.Freeze(False)
                Catch ex As Exception
                    objForm.Freeze(False)
                End Try
            ElseIf CFLEvent.ChooseFromListUID = "cfleid1" Then
                Try
                    objForm.Freeze(True)
                    objMatrix.FlushToDataSource()
                    objLine.SetValue("U_SupID", pval.Row - 1, datatable.GetValue("empID", 0))
                    objLine.SetValue("U_SupName", pval.Row - 1, datatable.GetValue("firstName", 0) + " " + datatable.GetValue("lastName", 0))
                    objMatrix.LoadFromDataSource()
                    objMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click()
                    objForm.Freeze(False)
                Catch ex As Exception
                    objForm.Freeze(False)
                End Try
            ElseIf CFLEvent.ChooseFromListUID = "cfleid2" Then
                Try
                    objForm.Freeze(True)
                    objMatrix.FlushToDataSource()
                    objLine.SetValue("U_TimeKpID", pval.Row - 1, datatable.GetValue("empID", 0))
                    objLine.SetValue("U_TimeKeepNam", pval.Row - 1, datatable.GetValue("firstName", 0) + " " + datatable.GetValue("lastName", 0))
                    objMatrix.LoadFromDataSource()
                    objMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click()
                    objForm.Freeze(False)
                Catch ex As Exception
                    objForm.Freeze(False)
                End Try
            ElseIf CFLEvent.ChooseFromListUID = "cfleid3" Then
                Try
                    objForm.Freeze(True)
                    objMatrix.FlushToDataSource()
                    objLine.SetValue("U_SecID", pval.Row - 1, datatable.GetValue("empID", 0))
                    objLine.SetValue("U_SecName", pval.Row - 1, datatable.GetValue("firstName", 0) + " " + datatable.GetValue("lastName", 0))
                    objMatrix.LoadFromDataSource()
                    objMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click()
                    objForm.Freeze(False)
                Catch ex As Exception
                    objForm.Freeze(False)
                End Try
            End If
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
        Catch ex As Exception
            objForm.Freeze(False)
        End Try

    End Sub

    Private Sub ChooseFromList_Filteration(ByVal FormUID As String, ByVal CFLID As String, ByVal ColAlias As String, ByVal ColValue As String, ByVal Query As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item(CFLID) '"EMP_1"
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            oCond = oConds.Add
            oCond.Alias = ColAlias ' "Active"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = ColValue ' "Y"
            If Query <> "" Then
                rsetCFL.DoQuery(Query)
                If rsetCFL.RecordCount > 0 Then
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    For i As Integer = 0 To rsetCFL.RecordCount - 1
                        If i = rsetCFL.RecordCount - 1 Then
                            oCond = oConds.Add
                            oCond.Alias = rsetCFL.Fields.Item(0).Name
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = rsetCFL.Fields.Item(0).Value
                        Else
                            oCond = oConds.Add
                            oCond.Alias = rsetCFL.Fields.Item(0).Name
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = rsetCFL.Fields.Item(0).Value
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        End If
                        rsetCFL.MoveNext()
                    Next
                End If
            End If
            oCFL.SetConditions(oConds)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LineTotalCalc(ByVal FormUID As String, ByVal RowID As Integer)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objLine = objForm.DataSources.DBDataSources.Item("@MIGTIN1")
        'objMatrix.GetLineData(RowID)
        Dim linetotal As Double

        'linetotal = CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) * CDbl(objLine.GetValue("U_unitpric", RowID - 1))
        linetotal = CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) * CDbl(objMatrix.Columns.Item("11").Cells.Item(RowID).Specific.value)

        objMatrix.Columns.Item("12").Cells.Item(RowID).Specific.value = CStr(linetotal)
        'objLine.SetValue("U_linetot", RowID - 1, linetotal)
        ' MsgBox(CStr(objLine.GetValue("U_linetot", RowID - 1)))
        'objMatrix.SetLineData(RowID)
        objForm.Update()
        objForm.Refresh()
    End Sub

    Private Function Validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try
            If BranchFlag = "Y" Then
                If objForm.Items.Item("70").Specific.Selected.Value Is Nothing Then
                    objAddOn.objApplication.SetStatusBarMessage("Please select Branch")
                    objForm.Items.Item("70").Click()
                    Return False
                End If
                objedit = objForm.Items.Item("23").Specific
                If objAddOn.HANA Then
                    strSQL = objAddOn.objGenFunc.getSingleValue("select 1 as ""Status"" From NNM1 where ""ObjectCode""='" & Formtype & "' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & objedit.Value & "' between ""F_RefDate"" and ""T_RefDate"") and ""BPLId"" is not null ")
                Else
                    strSQL = objAddOn.objGenFunc.getSingleValue("select 1 as Status From NNM1 where ObjectCode='" & Formtype & "' and Indicator=(select Top 1 Indicator  from OFPR where '" & objedit.Value & "' between F_RefDate and T_RefDate) and BPLId is not null")
                End If
                If strSQL <> "" Then
                    If objAddOn.HANA Then
                        strSQL = objAddOn.objGenFunc.getSingleValue("select 1 as ""Status"" From NNM1 where ""ObjectCode""='" & Formtype & "' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & objedit.Value & "' between ""F_RefDate"" and ""T_RefDate"") and ""BPLId""='" & objForm.Items.Item("70").Specific.Selected.Value & "' and ""Series""='" & objForm.Items.Item("20").Specific.Selected.Value & "'")
                    Else
                        strSQL = objAddOn.objGenFunc.getSingleValue("select 1 as Status From NNM1 where ObjectCode='" & Formtype & "' and Indicator=(select Top 1 Indicator  from OFPR where '" & objedit.Value & "' between F_RefDate and T_RefDate) and BPLId='" & objForm.Items.Item("70").Specific.Selected.Value & "' and Series='" & objForm.Items.Item("20").Specific.Selected.Value & "'")
                    End If
                    If strSQL = "" Then objAddOn.objApplication.SetStatusBarMessage("Cannot add transaction; numbering series assigned to another branch [Gate Entry Inward - Series] , '" & objForm.Items.Item("20").Specific.Selected.Description & "'") : Return False
                End If
            End If

            If Trim(objForm.Items.Item("4").Specific.Value) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please select Location name")
                objForm.Items.Item("4").Click()
                Return False
            ElseIf Trim(objForm.Items.Item("23").Specific.String) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please fill up Date")
                objForm.Items.Item("23").Click()
                Return False
            ElseIf Trim(objForm.Items.Item("6").Specific.Value) = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Please select Security name")
                objForm.Items.Item("6").Click()
                Return False
            End If
            objCombo = objForm.Items.Item("8").Specific
            If objCombo.Selected.Value = "PO" Or objCombo.Selected.Value = "DR" Or objCombo.Selected.Value = "GR" Or objCombo.Selected.Value = "IN" Or objCombo.Selected.Value = "SC" Or objCombo.Selected.Value = "WI" Then
                If Trim(objForm.Items.Item("10").Specific.String) = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Please fill up Party details")
                    objForm.Items.Item("10").Click()
                    Return False
                    'ElseIf Trim(objForm.Items.Item("14").Specific.String) = "" Then
                    '    objAddOn.objApplication.SetStatusBarMessage("Please fill up No of packages")
                    '    Return False
                ElseIf Trim(objForm.Items.Item("16").Specific.String) = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Please fill up In time")
                    objForm.Items.Item("16").Click()
                    Return False
                    'ElseIf Trim(objForm.Items.Item("18").Specific.String) = "" Then
                    '    objAddOn.objApplication.SetStatusBarMessage("Please fill up LR Number")
                    '    Return False

                    'ElseIf Trim(objForm.Items.Item("27").Specific.String) = "" Then
                    '    objAddOn.objApplication.SetStatusBarMessage("Please fill up Gate Entry Number")
                    '    Return False

                ElseIf Trim(objForm.Items.Item("31").Specific.String) = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Please fill up Vehicle Number")
                    objForm.Items.Item("31").Click()
                    Return False

                ElseIf Trim(objForm.Items.Item("33").Specific.String) = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Please fill up Transporter Name")
                    objForm.Items.Item("33").Click()
                    Return False
                End If
            End If

            If objAddOn.HANA Then
                strSQL = objAddOn.objGenFunc.getSingleValue("select count(*) from ""@MIGTIN"" where ""U_supinvno""='" & Trim(objForm.Items.Item("25B").Specific.String) & "' and ""U_partyid""='" & Trim(objForm.Items.Item("10").Specific.String) & "' and ""DocNum""<>'" & Trim(objForm.Items.Item("21").Specific.String) & "' and ""Series""='" & Trim(objForm.Items.Item("20").Specific.Selected.Value) & "' ")
            Else
                strSQL = objAddOn.objGenFunc.getSingleValue("select count(*) from [@MIGTIN] where U_supinvno='" & Trim(objForm.Items.Item("25B").Specific.String) & "' and U_partyid='" & Trim(objForm.Items.Item("10").Specific.String) & "' and DocNum<>'" & Trim(objForm.Items.Item("21").Specific.String) & "' and Series='" & Trim(objForm.Items.Item("20").Specific.Selected.Value) & "'")
            End If
            If CInt(strSQL) >= 1 Then
                objAddOn.objApplication.SetStatusBarMessage("Duplicate party Invoice No found... ")
                Return False
            End If
            If objAddOn.HANA Then
                strSQL = objAddOn.objGenFunc.getSingleValue("select count(*) from ""@MIGTIN"" where ""U_supdcno""='" & Trim(objForm.Items.Item("12B").Specific.String) & "' and ""U_partyid""='" & Trim(objForm.Items.Item("10").Specific.String) & "' and ""DocNum""<>'" & Trim(objForm.Items.Item("21").Specific.String) & "' and ""Series""='" & Trim(objForm.Items.Item("20").Specific.Selected.Value) & "'")
            Else
                strSQL = objAddOn.objGenFunc.getSingleValue("select count(*) from [@MIGTIN] where U_supdcno='" & Trim(objForm.Items.Item("12B").Specific.String) & "' and U_partyid='" & Trim(objForm.Items.Item("10").Specific.String) & "' and DocNum<>'" & Trim(objForm.Items.Item("21").Specific.String) & "' and Series='" & Trim(objForm.Items.Item("20").Specific.Selected.Value) & "'")
            End If
            If CInt(strSQL) >= 1 Then
                objAddOn.objApplication.SetStatusBarMessage("Duplicate party DC No found... ")
                Return False
            End If
            objMatrix = objForm.Items.Item("36").Specific
            If objCombo.Selected.Value = "CL" Then
                'If oContLabMatrix.RowCount = 0 Then
                '    objAddOn.objApplication.SetStatusBarMessage("Minimum one Line Item is Required.. ")
                '    Return False
                'Else
                '    If oContLabMatrix.Columns.Item("1").Cells.Item(1).Specific.value = "" Then
                '        objAddOn.objApplication.SetStatusBarMessage("Minimum one Line Item is Required.. ")
                '        Return False
                '    End If
                'End If
            Else
                If objMatrix.RowCount = 0 Then
                    objAddOn.objApplication.SetStatusBarMessage("Minimum one Line Item is Required.. ")
                    Return False
                Else
                    If objMatrix.Columns.Item("1").Cells.Item(1).Specific.value = "" Then
                        objAddOn.objApplication.SetStatusBarMessage("Minimum one Line Item is Required.. ")
                        Return False
                    End If
                End If
            End If


        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage("Validate: " & ex.Message)
            Return False
        End Try
        Return True
    End Function

    Private Function QtyValidation(ByVal FormUID As String, ByVal RowID As Integer) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCombo = objForm.Items.Item("8").Specific
        Select Case (UCase(Trim(objCombo.Value)))
            Case "DC", "JR", "SR", "WM", "RW", "ST", "SO", "CP", "JO", "JW", "MI", "HR"
                Return True
        End Select
        objMatrix = objForm.Items.Item("36").Specific
        objLine = objForm.DataSources.DBDataSources.Item("@MIGTIN1")
        objMatrix.GetLineData(RowID)
        If objCombo.Selected.Value = "PO" Then
            If objMatrix.Columns.Item("4").Cells.Item(RowID).Specific.value.ToString.ToUpper.StartsWith("RM") Then
                If CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) > (CDbl(objLine.GetValue("U_pendqty", RowID - 1)) + CDbl(objLine.GetValue("U_pendqty", RowID - 1)) * 0.1) Then
                    objAddOn.objApplication.SetStatusBarMessage("Quantity exceeds pending quantity with 10% tolerance", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
            Else
                If CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) > CDbl(objLine.GetValue("U_pendqty", RowID - 1)) Then
                    objAddOn.objApplication.SetStatusBarMessage("Quantity exceeds pending quantity", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
            End If
        Else
            If CDbl(objMatrix.Columns.Item("9").Cells.Item(RowID).Specific.value) > CDbl(objLine.GetValue("U_pendqty", RowID - 1)) Then
                objAddOn.objApplication.SetStatusBarMessage("Quantity exceeds pending quantity", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            End If
        End If


        Return True
    End Function

    Private Sub OpenColumns(ByVal FormUID As String)
        objMatrix = objForm.Items.Item("36").Specific
        Select Case (UCase(Trim(objCombo.Value)))
            Case "DC", "JR", "SR", "WM", "RW", "ST", "SO", "CP", "JO", "JW", "MI", "HR"
                objMatrix.Columns.Item("11").Editable = True
            Case Else
                objMatrix.Columns.Item("11").Editable = False
        End Select

    End Sub

    Private Sub CopyTo(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        '  objCombo = objForm.Items.Item("49").Specific
        objCombo = objForm.Items.Item("8").Specific
        If Trim(objCombo.Value) = "" Or Trim(objCombo.Value) = "CL" Then Exit Sub
        If Not objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            objAddOn.objApplication.MessageBox("Form should be in OK mode")
            Exit Sub
        End If
        Select Case Trim(objCombo.Value)
            Case "GR"
                objAddOn.objApplication.MessageBox("Please generate GRPO document with copy from PO and Do GE Verification")
            Case "CP"
                CopyToGRN(FormUID)
            Case "SC"
                CopyToARCreditMemo(FormUID) ' 179,180,721,940,141
            Case "WI" ' 
                CopyToSalesReturn(FormUID) ' 180
            Case "MI"
                CopyToGoodsReceipt(FormUID) '721
            Case "DC", "JR", "SP", "WM", "RW", "ST", "SO", "JO", "JW"
                CopyToStockTransfer(FormUID) '940
            Case "HR"
                CopyToAPInvoice(FormUID) '141
        End Select
    End Sub

    Private Sub CopyFrom(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objHeader = objForm.DataSources.DBDataSources.Item("@MIGTIN")
        objCombo = objForm.Items.Item("8").Specific
        If Trim(objCombo.Value) = "" Or Trim(objCombo.Value) = "CL" Then Exit Sub
        ' If objForm.Items.Item("10").Specific.string <> "" Then

        If objForm.Items.Item("10").Specific.string = "" Then
            If objCombo.Selected.Value = "PO" Or objCombo.Selected.Value = "SR" Or objCombo.Selected.Value = "DR" Or objCombo.Selected.Value = "GR" Or objCombo.Selected.Value = "IN" Or objCombo.Selected.Value = "SC" Or objCombo.Selected.Value = "WI" Then
                objAddOn.objApplication.MessageBox("Please select Party id") : Exit Sub
            End If
        End If
        OpenColumns(FormUID)
        objAddOn.objItemDetails.LoadScreen(Formtype, objForm.TypeCount, objCombo.Value, objForm.Items.Item("10").Specific.string, objHeader.GetValue("U_cutdate", 0), objHeader.GetValue("DocEntry", 0))
        ' Else
        'objAddOn.objApplication.MessageBox("Please select Party id")
        'End If

    End Sub

    Private Sub CopyToStockTransfer(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("3080")
        'Matrix 23; form 940
        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("940", 1)
        CopyToForm.Items.Item("3").Specific.String = objHeader.GetValue("U_partyid", 0)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("23").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            ' CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub

    Private Sub CopyToARCreditMemo(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("2085")

        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("179", 1)
        CopyToForm.Items.Item("4").Specific.String = objHeader.GetValue("U_partyid", 0)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("38").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            ' CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub

    Private Sub CopyToGRN(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        Try


            objAddOn.objApplication.ActivateMenuItem("2306")

            Dim CopyToForm As SAPbouiCOM.Form
            CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("143", 1)
            CopyToForm.Items.Item("4").Specific.String = objHeader.GetValue("U_partyid", 0)
            Dim CTMatrix As SAPbouiCOM.Matrix
            CTMatrix = CopyToForm.Items.Item("38").Specific
            CopyToForm.Items.Item("TVer").Specific.String = objHeader.GetValue("DocEntry", 0)
            For i As Integer = 1 To objMatrix.RowCount
                CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
                ' CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
                CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
                CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
                CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
                CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
                'Dim objGRN As SAPbobsCOM.Documents
                'objGRN = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                'objGRN.Lines.BaseType = objMatrix.Columns.Item("1A").Cells.Item(i).Specific.String
                'objGRN.Lines.BaseEntry = objMatrix.Columns.Item("2").Cells.Item(i).Specific.String
                'objGRN.Lines.BaseLine = objMatrix.Columns.Item("3").Cells.Item(i).Specific.String

            Next
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub CopyToSalesReturn(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("2052")

        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("180", 1)
        CopyToForm.Items.Item("4").Specific.String = objHeader.GetValue("U_partyid", 0)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("38").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            'CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub

    Private Sub CopyToGoodsReceipt(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("3078")

        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("721", 1)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("13").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            'CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("9").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub

    Private Sub CopyToAPInvoice(ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("36").Specific
        objAddOn.objApplication.ActivateMenuItem("2308")

        Dim CopyToForm As SAPbouiCOM.Form
        CopyToForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("141", 1)
        CopyToForm.Items.Item("4").Specific.String = objHeader.GetValue("U_partyid", 0)
        Dim CTMatrix As SAPbouiCOM.Matrix
        CTMatrix = CopyToForm.Items.Item("39").Specific

        For i As Integer = 1 To objMatrix.RowCount
            CTMatrix.Columns.Item("1").Cells.Item(i).Specific.String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.String
            'CTMatrix.Columns.Item("2").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("10").Cells.Item(i).Specific.String = objMatrix.Columns.Item("11").Cells.Item(i).Specific.String
            CTMatrix.Columns.Item("U_getype").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("U_type", 0))
            CTMatrix.Columns.Item("U_gedocno").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocNum", 0))
            CTMatrix.Columns.Item("U_geentry").Cells.Item(i).Specific.String = Trim(objHeader.GetValue("DocEntry", 0))
        Next

    End Sub

    Private Sub viewMode(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

    End Sub

    Private Sub ManageAttributes()
        Try
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "20", True, True, False) 'Series
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "25", True, True, False) 'Status
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "21", True, True, False) 'Doc Num
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "4", True, True, False) 'Location
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "6", True, True, False) 'Security Name
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "8", True, True, False) 'Type
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "10", True, True, False) 'Party Id
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "23", True, True, False) 'Date
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "51B", False, True, False) 'Process Status
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "52B", False, True, False) 'Target Doc
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "12", False, True, False) 'Party Name
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "52C1", False, False, True) 'link Tran
            objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "70", True, True, False) 'Branch
            'objForm.Items.Item("51B").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "12", True, True, False) 'Party Name
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "12B", True, True, False) 'Party DC No
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "12D", True, True, False) 'Party DC Date
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "14", True, True, False) 'No of Packages
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "16", True, True, False) 'In Time
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "18", True, True, False) 'LR No            '
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "25B", True, True, False) 'Party Inv No
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "25D", True, True, False) 'Party Inv Date
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "27", True, True, False) 'Gate Entry No
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "29", True, True, False) 'Weight Challan
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "31", True, True, False) 'Vehicle No
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "33", True, True, False)  'Transporter Name
            'objAddOn.objGenFunc.SetAutomanagedattribute_Editable(objForm, "35", True, True, False)  'LR Date

        Catch ex As Exception

        End Try
    End Sub

    ''Private Sub updatePO(ByVal FormUID As String)

    ''    objForm = objAddOn.objApplication.Forms.Item(FormUID)
    ''    objCombo = objForm.Items.Item("20").Specific
    ''    If objCombo.Value = "PO" Then
    ''        Dim objPO As SAPbobsCOM.Documents
    ''        objPO = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
    ''        For i As Integer = 1 To objMatrix.RowCount
    ''            Dim POEntry As Integer = objHeader.GetValue("DocEntry", 0)
    ''            If objPO.GetByKey(POEntry) Then
    ''                objPO.Lines.SetCurrentLine(CInt(objMatrix.Columns.Item("0").Cells.Item(i).Specific.String))
    ''                objPO.Lines.UserFields.Fields.Item("U_getype").Value = objHeader.GetValue("U_type", 0)
    ''                objPO.Lines.UserFields.Fields.Item("U_gedocno").Value = objHeader.GetValue("DocNum", 0)
    ''                objPO.Lines.UserFields.Fields.Item("U_geentry").Value = objHeader.GetValue("DocEntry", 0)
    ''            End If
    ''        Next

    ''    End If
    ''End Sub

    Sub RightClickEvent(ByRef EventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(objForm.UniqueID)
            objMatrix = objForm.Items.Item("36").Specific
            If EventInfo.BeforeAction Then
                objForm.EnableMenu("1283", False)
                objForm.EnableMenu("1285", False)
                objForm.EnableMenu("1286", False)
                Select Case EventInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                        Try
                            If EventInfo.ItemUID = "" Then Exit Try
                            If objForm.Items.Item(EventInfo.ItemUID).Specific.String <> "" Then
                                objForm.EnableMenu("772", True)  'Copy
                            ElseIf objForm.Items.Item(EventInfo.ItemUID).Specific.String = "" Then
                                objForm.EnableMenu("773", True)  'Paste
                            End If
                        Catch ex As Exception
                            objMatrix = objForm.Items.Item(EventInfo.ItemUID).Specific
                            If EventInfo.Row <= 0 Then If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then objForm.EnableMenu("772", True) : objForm.EnableMenu("784", True) : Exit Try
                            If objMatrix.Columns.Item(EventInfo.ColUID).Cells.Item(EventInfo.Row).Specific.String <> "" Then
                                objForm.EnableMenu("772", True)  'Copy
                            ElseIf objMatrix.Columns.Item(EventInfo.ColUID).Cells.Item(EventInfo.Row).Specific.String = "" Then
                                objForm.EnableMenu("773", True)  'Paste
                            End If
                        End Try

                        Select Case EventInfo.ItemUID
                            Case "36", "74"
                                If EventInfo.ItemUID = "36" And EventInfo.ColUID = "0" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And EventInfo.Row > 0 Then
                                    objForm.EnableMenu("1293", True)
                                    'ElseIf EventInfo.ItemUID = "74" And EventInfo.ColUID = "0" And EventInfo.Row > 0 Then
                                    'oContLabMatrix = objForm.Items.Item("74").Specific
                                    'If oContLabMatrix.VisualRowCount = 1 Then objForm.EnableMenu("1293", False) Else objForm.EnableMenu("1293", True)
                                End If
                                If EventInfo.ItemUID = "36" And EventInfo.ColUID = "21B" Then
                                    objForm.EnableMenu("773", False)
                                End If
                            Case ""
                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                                If objForm.Items.Item("52B").Specific.String <> "" Then
                                    If objAddOn.HANA Then
                                        strQuery = objAddOn.objGenFunc.getSingleValue("Select  Case when T0.DocStatus='O' then 0 Else 1 End [Draft Status] from ODRF T0 left join OPDN T1 On T0.DocEntry=T1.draftKey and T0.ObjType=T1.ObjType Where T0.DocEntry=" & objForm.Items.Item("52B").Specific.String & "")
                                    Else
                                        strQuery = objAddOn.objGenFunc.getSingleValue("Select  Case when T0.DocStatus='O' then 0 Else 1 End [Draft Status] from ODRF T0 left join OPDN T1 On T0.DocEntry=T1.draftKey and T0.ObjType=T1.ObjType Where T0.DocEntry=" & objForm.Items.Item("52B").Specific.String & "")
                                    End If
                                    If strQuery = "0" And objForm.Items.Item("51B").Specific.Selected.Value = "3" And objForm.Items.Item("25").Specific.Selected.Value = "O" Then
                                        objForm.EnableMenu("1284", True) 'Cancel
                                        objForm.EnableMenu("1286", True) 'Close
                                    Else
                                        objForm.EnableMenu("1284", False) 'Cancel
                                        objForm.EnableMenu("1286", False) 'Close
                                    End If
                                End If

                                'If objForm.Items.Item("52B").Specific.String = "" And objForm.Items.Item("51B").Specific.Selected.Value = "0" And objForm.Items.Item("25").Specific.Selected.Value = "O" Then
                                '    objForm.EnableMenu("1284", True) 'Cancel
                                '    objForm.EnableMenu("1286", True) 'Close
                                'Else
                                '    objForm.EnableMenu("1284", False) 'Cancel
                                '    objForm.EnableMenu("1286", False) 'Close
                                'End If
                        End Select
                End Select
            Else
                objForm.EnableMenu("772", False)
                objForm.EnableMenu("773", False)
                objForm.EnableMenu("784", False)
                objForm.EnableMenu("1293", False)
                objForm.EnableMenu("1283", False)
                objForm.EnableMenu("1284", False)
                objForm.EnableMenu("1286", False)
            End If
        Catch ex As Exception
            'objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub

    Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
        Try
            Dim DBSource As SAPbouiCOM.DBDataSource
            'objMatrix = objform.Items.Item("20").Specific
            objMatrix.FlushToDataSource()
            DBSource = objForm.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
            For i As Integer = 1 To objMatrix.VisualRowCount
                objMatrix.GetLineData(i)
                DBSource.Offset = i - 1
                DBSource.SetValue("LineId", DBSource.Offset, i)
                objMatrix.SetLineData(i)
                objMatrix.FlushToDataSource()
            Next
            DBSource.RemoveRecord(DBSource.Size - 1)
            objMatrix.LoadFromDataSource()

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        Finally
        End Try
    End Sub

    Private Function Create_GoodsReceipt_PO(ByVal FormUID As String)
        Try
            'If objForm.Items.Item("2A").Enabled = False Then Return False

            Dim objedit As SAPbouiCOM.EditText
            Dim objGRPO As SAPbobsCOM.Documents
            Dim DocEntry, strQuery As String
            Dim Lineflag As Boolean = False
            Dim Row As Integer = 1
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("36").Specific
            If objForm.Items.Item("52B").Specific.String <> "" Then Return True

            objGRPO = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objAddOn.objApplication.StatusBar.SetText("Creating Goods Receipt PO draft. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objedit = objForm.Items.Item("23").Specific
            Dim DocDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            For MatRow As Integer = 1 To objMatrix.VisualRowCount
                DocEntry = DocEntry + objMatrix.Columns.Item("2").Cells.Item(MatRow).Specific.String + ","
            Next
            DocEntry = DocEntry.Remove(DocEntry.Length - 1)
            If objAddOn.HANA Then
                strQuery = objAddOn.objGenFunc.getSingleValue("Select Distinct 1 as ""Status"" from POR1 A where A.""DocEntry"" in (" & DocEntry & ") and A.""LineStatus""='O' ")
            Else
                strQuery = objAddOn.objGenFunc.getSingleValue("Select Distinct 1 as Status from POR1 A where A.DocEntry in (" & DocEntry & ") and A.LineStatus='O' ")
            End If

            'strQuery = objAddOn.objGenFunc.getSingleValue("Select Distinct 1 as ""Status"" from ""@MIGTIN1"" B join POR1 A on B.""U_basentry""=A.""DocEntry"" and B.""U_itemcode""=A.""ItemCode"" where  B.""DocEntry""=" & objHeader.GetValue("DocEntry", 0) & " and A.""LineStatus""='C' ")
            If strQuery = "" Then objAddOn.objApplication.StatusBar.SetText("PO Status Closed for this Transaction...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
            'strQuery = "Select B.* from ""@MIGTIN"" A join ""@MIGTIN1"" B on A.""DocEntry""=B.""DocEntry"" where A.""U_trgtentry"" is null"

            'objRS.DoQuery(strQuery)
            'If objRS.RecordCount = 0 Then objAddOn.objApplication.StatusBar.SetText("No Records Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
            If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()

            objGRPO.DocDate = DocDate
            'objGRPO.JournalMemo = "Auto-Gen-> " & Now.ToString
            objGRPO.UserFields.Fields.Item("U_GateRem").Value = "From Gate Inward Auto-Gen-> " & Now.ToString
            objGRPO.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes
            If objForm.Items.Item("8").Specific.Selected.Value = "PO" Then
                objGRPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
            Else
                objGRPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
            End If

            ''objGRPO.UserFields.Fields.Item("U_GEGR").Value = objHeader.GetValue("DocEntry", 0)
            'objGRPO.UserFields.Fields.Item("U_gever").Value = GEEntry
            ''objGRPO.UserFields.Fields.Item("U_GEEntry").Value = GEEntry
            If objAddOn.HANA Then
                strQuery = "Select ""BPLId"" from OBPL where ""Disabled""='N' and ""MainBPL""='Y'" 'Branch
            Else
                strQuery = "Select BPLId from OBPL where Disabled='N' and MainBPL='Y'" 'Branch
            End If
            strQuery = objAddOn.objGenFunc.getSingleValue(strQuery)
            If strQuery <> "" Then objGRPO.BPL_IDAssignedToInvoice = strQuery
            If objMatrix.VisualRowCount > 0 Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If objGRPO.CardCode = "" Then objGRPO.CardCode = Trim(objForm.Items.Item("10").Specific.String)
                    If objForm.Items.Item("8").Specific.Selected.Value = "PO" Then
                        objGRPO.Lines.ItemCode = Trim(objMatrix.Columns.Item("4").Cells.Item(i).Specific.String)
                        objGRPO.Lines.Quantity = CDbl(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String) ' CDbl(Matrix0.Columns.Item("grnqty").Cells.Item(i).Specific.String) ' CDbl(objRs.Fields.Item("GRN Qty").Value.ToString) 
                        objGRPO.Lines.UserFields.Fields.Item("U_GateQty").Value = CDbl(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String)
                    Else
                        objGRPO.Lines.ItemDescription = Trim(objMatrix.Columns.Item("5").Cells.Item(i).Specific.String)
                        objGRPO.Lines.AccountCode = Trim(objMatrix.Columns.Item("27").Cells.Item(i).Specific.String)
                        objGRPO.Lines.TaxCode = Trim(objMatrix.Columns.Item("29").Cells.Item(i).Specific.String)
                        objGRPO.Lines.LineTotal = Trim(objMatrix.Columns.Item("29A").Cells.Item(i).Specific.String)
                    End If
                    objGRPO.Lines.BaseType = 22
                    objGRPO.Lines.BaseEntry = CInt(objMatrix.Columns.Item("2").Cells.Item(i).Specific.String) ' CInt(objRs.Fields.Item("PO Entry").Value.ToString)
                    objGRPO.Lines.BaseLine = CInt(objMatrix.Columns.Item("3").Cells.Item(i).Specific.String)
                    'objGRPO.Lines.UnitPrice = Trim(objRS.Fields.Item("Price").Value.ToString)
                    'objGRPO.Lines.WarehouseCode = Trim(objRS.Fields.Item("WhsCode").Value.ToString)
                    objGRPO.Lines.Add()
                Next

            End If

            If objGRPO.Add() <> 0 Then
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objAddOn.objApplication.SetStatusBarMessage("GRPO: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox("GRPO: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                Return False
            Else
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                DocEntry = objAddOn.objCompany.GetNewObjectKey()
                objForm.Items.Item("52B").Specific.String = DocEntry
                objAddOn.objApplication.StatusBar.SetText("Draft GRPO Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return True
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objGRPO)
            GC.Collect()
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub CreateMySimpleForm(ByVal FormID As String, ByVal FormTitle As String, ByVal DraftHeader As String, ByVal TranHeader As String, ByVal LinkedID As String)
        Dim oCreationParams As SAPbouiCOM.FormCreationParams
        Dim objTempForm As SAPbouiCOM.Form
        Dim objrs As SAPbobsCOM.Recordset
        Try
            objAddOn.objApplication.Forms.Item(FormID).Visible = True
        Catch ex As Exception
            oCreationParams = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            oCreationParams.UniqueID = FormID
            objTempForm = objAddOn.objApplication.Forms.AddEx(oCreationParams)
            objTempForm.Title = FormTitle
            objTempForm.Left = 400
            objTempForm.Top = 100
            objTempForm.ClientHeight = 200 '335
            objTempForm.ClientWidth = 500
            objTempForm.Left = objForm.Left + 100
            objTempForm.Top = objForm.Top + 100
            objTempForm = objAddOn.objApplication.Forms.Item(FormID)
            Dim oitm As SAPbouiCOM.Item

            Dim oGrid As SAPbouiCOM.Grid
            oitm = objTempForm.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oitm.Top = 30
            oitm.Left = 2
            oitm.Width = 490
            oitm.Height = 120
            oGrid = objTempForm.Items.Item("Grid").Specific
            objTempForm.DataSources.DataTables.Add("DataTable")
            oitm = objTempForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitm.Top = objTempForm.Items.Item("Grid").Top + objTempForm.Items.Item("Grid").Height + 10
            oitm.Left = 2
            Dim str_sql As String = ""
            If objForm.Items.Item("52B").Specific.String = "" Then objAddOn.objApplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Sub
            'If objAddOn.HANA Then
            '    str_sql = "select T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OPDN T0 where T0.""U_gever""='" & objForm.Items.Item("57").Specific.String & "' and ""CANCELED""='N'"
            'Else
            '    str_sql = "select T0.DocEntry,T0.DocNum,T0.DocDate from OPDN T0 where T0.U_gever='" & objForm.Items.Item("57").Specific.String & "' and CANCELED='N'"
            'End If

            If objAddOn.HANA Then
                str_sql = "Select Case when T0.""DocStatus""='O' then T0.""DocEntry"" Else T1.""DocEntry"" End ""DocEntry"",Case when T0.""DocStatus""='O' then T0.""DocNum"" Else T1.""DocNum"" End ""DocNum"","
                str_sql += vbCrLf + "Case when T0.""DocStatus""='O' then T0.""DocDate"" Else T1.""DocDate"" End ""DocDate"",Case when T0.""DocStatus""='O' then '112' Else T1.""ObjType"" End ""ObjType"""
                str_sql += vbCrLf + "from " & DraftHeader & " T0 left join " & TranHeader & " T1 On T0.""DocEntry""=T1.""draftKey"" and T0.""ObjType""=T1.""ObjType"" where ifnull(T0.""U_gever"",'')='" & objForm.Items.Item("57").Specific.String & "' and ifnull(T1.""CANCELED"",'N')='N'"

            Else
                str_sql = "Select Case when T0.DocStatus='O' then T0.DocEntry Else T1.DocEntry End DocEntry,Case when T0.DocStatus='O' then T0.DocNum Else T1.DocNum End DocNum,"
                str_sql += vbCrLf + "Case when T0.DocStatus='O' then T0.DocDate Else T1.DocDate End DocDate,Case when T0.DocStatus='O' then '112' Else T1.ObjType End ObjType"
                str_sql += vbCrLf + "from " & DraftHeader & " T0 left join " & TranHeader & " T1 On T0.DocEntry=T1.draftKey and T0.ObjType=T1.ObjType where isnull(T0.U_gever,'')='" & objForm.Items.Item("57").Specific.String & "' and isnull(T1.CANCELED,'N')='N'"
            End If

            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(str_sql)
            If objrs.RecordCount = 0 Then
                If objAddOn.HANA Then
                    str_sql = "Select T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from " & TranHeader & " T0  where ifnull(T0.""U_gever"",'')='" & objForm.Items.Item("57").Specific.String & "' and ifnull(T0.""CANCELED"",'N')='N'"
                Else
                    str_sql = "Select T0.DocEntry,T0.DocNum,T0.DocDate from " & TranHeader & " T0  where isnull(T0.U_gever,'')='" & objForm.Items.Item("57").Specific.String & "' and isnull(T0.CANCELED,'N')='N'"
                End If
                objrs.DoQuery(str_sql)
                If objrs.RecordCount = 0 Then
                    objAddOn.objApplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objrs = Nothing
                    Exit Sub
                End If
            End If
            Dim objDT As SAPbouiCOM.DataTable
            objDT = objTempForm.DataSources.DataTables.Item("DataTable")
            objDT.Clear()
            objDT.ExecuteQuery(str_sql)
            objTempForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(str_sql)

            oGrid.DataTable = objTempForm.DataSources.DataTables.Item("DataTable")
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).TitleObject.Sortable = True
                oGrid.Columns.Item(i).Editable = False
            Next
            oGrid.Rows.SelectedRows.Add(0)
            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            Dim col As SAPbouiCOM.EditTextColumn
            col = oGrid.Columns.Item(0)
            col.LinkedObjectType = LinkedID

            objTempForm.Visible = True
            objTempForm.Update()
            'bModal = True
            'FormName = "ViewD"
        End Try
    End Sub

    Private Function Create_ARCreditMemo(ByVal FormUID As String)
        Try
            'If objForm.Items.Item("2A").Enabled = False Then Return False

            Dim objedit As SAPbouiCOM.EditText
            Dim objSalesReturn As SAPbobsCOM.Documents
            Dim DocEntry, strQuery, DocStatus As String
            Dim Lineflag As Boolean = False
            Dim Row As Integer = 1
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("36").Specific
            If objForm.Items.Item("52B").Specific.String <> "" Then Return True

            objSalesReturn = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objAddOn.objApplication.StatusBar.SetText("Creating A/R CreditMemo draft. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objedit = objForm.Items.Item("23").Specific
            Dim DocDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            For MatRow As Integer = 1 To objMatrix.VisualRowCount
                DocEntry = DocEntry + objMatrix.Columns.Item("2").Cells.Item(MatRow).Specific.String + ","
            Next
            DocEntry = DocEntry.Remove(DocEntry.Length - 1)
            'strQuery = objAddOn.objGenFunc.getSingleValue("Select Distinct 1 as ""Status"" from INV1 A where A.""DocEntry"" in (" & DocEntry & ") and A.""LineStatus""='C' ")

            'If DocStatus = "C" Then objAddOn.objApplication.StatusBar.SetText("A/R Invoice Status Closed for this Transaction...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
            ''strQuery = "Select B.* from ""@MIGTIN"" A join ""@MIGTIN1"" B on A.""DocEntry""=B.""DocEntry"" where A.""U_trgtentry"" is null"
            'objRS.DoQuery(strQuery)
            'If objRS.RecordCount = 0 Then objAddOn.objApplication.StatusBar.SetText("No Records Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
            If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()

            objSalesReturn.DocDate = DocDate
            'objSalesReturn.JournalMemo = "Auto-Gen-> " & Now.ToString
            objSalesReturn.UserFields.Fields.Item("U_GateRem").Value = "From Gate Inward Auto-Gen-> " & Now.ToString
            objSalesReturn.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes
            objSalesReturn.GSTTransactionType = SAPbobsCOM.GSTTransactionTypeEnum.gsttrantyp_GSTTaxInvoice
            If objSalesReturn.CardCode = "" Then objSalesReturn.CardCode = Trim(objForm.Items.Item("10").Specific.String)
            If objAddOn.HANA Then
                strQuery = "Select ""BPLId"" from OBPL where ""Disabled""='N' and ""MainBPL""='Y'" 'Branch
            Else
                strQuery = "Select BPLId from OBPL where Disabled='N' and MainBPL='Y'" 'Branch
            End If
            strQuery = objAddOn.objGenFunc.getSingleValue(strQuery)
            If strQuery <> "" Then objSalesReturn.BPL_IDAssignedToInvoice = strQuery
            If objMatrix.VisualRowCount > 0 Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objSalesReturn.Lines.ItemCode = Trim(objMatrix.Columns.Item("4").Cells.Item(i).Specific.String)
                    objSalesReturn.Lines.Quantity = CDbl(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String) ' CDbl(Matrix0.Columns.Item("grnqty").Cells.Item(i).Specific.String) ' CDbl(objRs.Fields.Item("GRN Qty").Value.ToString) 
                    'objGRPO.Lines.AccountCode = Trim(objRS.Fields.Item("AcctCode").Value.ToString)
                    'objGRPO.Lines.TaxCode = Trim(objRS.Fields.Item("TaxCode").Value.ToString)
                    If objAddOn.HANA Then
                        DocStatus = objAddOn.objGenFunc.getSingleValue("Select A.""DocStatus"" from OINV A where A.""DocEntry"" = " & CInt(objMatrix.Columns.Item("2").Cells.Item(i).Specific.String) & "")
                    Else
                        DocStatus = objAddOn.objGenFunc.getSingleValue("Select A.DocStatus from OINV A where A.DocEntry = " & CInt(objMatrix.Columns.Item("2").Cells.Item(i).Specific.String) & "")
                    End If

                    If DocStatus = "O" Then
                        objSalesReturn.Lines.BaseType = 13
                        objSalesReturn.Lines.BaseEntry = CInt(objMatrix.Columns.Item("2").Cells.Item(i).Specific.String)
                        objSalesReturn.Lines.BaseLine = CInt(objMatrix.Columns.Item("3").Cells.Item(i).Specific.String)
                    End If
                    'objGRPO.Lines.UnitPrice = Trim(objRS.Fields.Item("Price").Value.ToString)
                    'objGRPO.Lines.WarehouseCode = Trim(objRS.Fields.Item("WhsCode").Value.ToString)
                    objSalesReturn.Lines.UserFields.Fields.Item("U_GateQty").Value = CDbl(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String)
                    objSalesReturn.Lines.Add()
                Next

            End If

            If objSalesReturn.Add() <> 0 Then
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objAddOn.objApplication.SetStatusBarMessage("A/R CreditMemo: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox("A/R CreditMemo: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                Return False
            Else
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                DocEntry = objAddOn.objCompany.GetNewObjectKey()
                objForm.Items.Item("52B").Specific.String = DocEntry
                objAddOn.objApplication.StatusBar.SetText("Draft A/R CreditMemo Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return True
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objSalesReturn)
            GC.Collect()
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function Create_Delivery_Return(ByVal FormUID As String)
        Try

            Dim objedit As SAPbouiCOM.EditText
            Dim objReturn As SAPbobsCOM.Documents
            Dim DocEntry, strQuery As String
            Dim Lineflag As Boolean = False
            Dim Row As Integer = 1
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("36").Specific
            If objForm.Items.Item("52B").Specific.String <> "" Then Return True

            objReturn = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objAddOn.objApplication.StatusBar.SetText("Creating Return draft. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objedit = objForm.Items.Item("23").Specific
            Dim DocDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            'For MatRow As Integer = 1 To objMatrix.VisualRowCount
            '    DocEntry = DocEntry + objMatrix.Columns.Item("2").Cells.Item(MatRow).Specific.String + ","
            'Next
            'DocEntry = DocEntry.Remove(DocEntry.Length - 1)
            'strQuery = objAddOn.objGenFunc.getSingleValue("Select Distinct 1 as ""Status"" from DLN1 A where A.""DocEntry"" in (" & DocEntry & ") and A.""LineStatus""='C' ")
            ''strQuery = objAddOn.objGenFunc.getSingleValue("Select Distinct 1 as ""Status"" from ""@MIGTIN1"" B join POR1 A on B.""U_basentry""=A.""DocEntry"" and B.""U_itemcode""=A.""ItemCode"" where  B.""DocEntry""=" & objHeader.GetValue("DocEntry", 0) & " and A.""LineStatus""='C' ")
            'If strQuery = "1" Then objAddOn.objApplication.StatusBar.SetText("PO Status Closed for this Transaction...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
            ''strQuery = "Select B.* from ""@MIGTIN"" A join ""@MIGTIN1"" B on A.""DocEntry""=B.""DocEntry"" where A.""U_trgtentry"" is null"

            ''objRS.DoQuery(strQuery)
            ''If objRS.RecordCount = 0 Then objAddOn.objApplication.StatusBar.SetText("No Records Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
            If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()

            objReturn.DocDate = DocDate
            'objGRPO.JournalMemo = "Auto-Gen-> " & Now.ToString
            objReturn.UserFields.Fields.Item("U_GateRem").Value = "From Gate Inward Auto-Gen-> " & Now.ToString
            objReturn.DocObjectCode = SAPbobsCOM.BoObjectTypes.oReturns
            ''objGRPO.UserFields.Fields.Item("U_GEGR").Value = objHeader.GetValue("DocEntry", 0)
            'objGRPO.UserFields.Fields.Item("U_gever").Value = GEEntry
            ''objGRPO.UserFields.Fields.Item("U_GEEntry").Value = GEEntry
            If objAddOn.HANA Then
                strQuery = "Select ""BPLId"" from OBPL where ""Disabled""='N' and ""MainBPL""='Y'" 'Branch
            Else
                strQuery = "Select BPLId from OBPL where Disabled='N' and MainBPL='Y'" 'Branch
            End If

            strQuery = objAddOn.objGenFunc.getSingleValue(strQuery)
            If strQuery <> "" Then objReturn.BPL_IDAssignedToInvoice = strQuery
            objReturn.CardCode = Trim(objForm.Items.Item("10").Specific.String)
            If objMatrix.VisualRowCount > 0 Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objReturn.Lines.ItemCode = Trim(objMatrix.Columns.Item("4").Cells.Item(i).Specific.String)
                    objReturn.Lines.Quantity = CDbl(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String) ' CDbl(Matrix0.Columns.Item("grnqty").Cells.Item(i).Specific.String) ' CDbl(objRs.Fields.Item("GRN Qty").Value.ToString) 
                    'objGRPO.Lines.AccountCode = Trim(objRS.Fields.Item("AcctCode").Value.ToString)
                    'objGRPO.Lines.TaxCode = Trim(objRS.Fields.Item("TaxCode").Value.ToString)
                    objReturn.Lines.BaseType = 15
                    objReturn.Lines.BaseEntry = CInt(objMatrix.Columns.Item("2").Cells.Item(i).Specific.String) ' CInt(objRs.Fields.Item("PO Entry").Value.ToString)
                    objReturn.Lines.BaseLine = CInt(objMatrix.Columns.Item("3").Cells.Item(i).Specific.String)
                    'objGRPO.Lines.UnitPrice = Trim(objRS.Fields.Item("Price").Value.ToString)
                    'objGRPO.Lines.WarehouseCode = Trim(objRS.Fields.Item("WhsCode").Value.ToString)
                    objReturn.Lines.UserFields.Fields.Item("U_GateQty").Value = CDbl(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String)
                    objReturn.Lines.Add()
                Next

            End If

            If objReturn.Add() <> 0 Then
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objAddOn.objApplication.SetStatusBarMessage("Return: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox("Return: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                Return False
            Else
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                DocEntry = objAddOn.objCompany.GetNewObjectKey()
                objForm.Items.Item("52B").Specific.String = DocEntry
                objAddOn.objApplication.StatusBar.SetText("Draft Return Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return True
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objReturn)
            GC.Collect()
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function Create_GoodsReceipt(ByVal FormUID As String)
        Try

            Dim objedit As SAPbouiCOM.EditText
            Dim objGoodsReceipt As SAPbobsCOM.Documents
            Dim DocEntry, strQuery As String
            Dim Lineflag As Boolean = False
            Dim Row As Integer = 1
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("36").Specific
            If objForm.Items.Item("52B").Specific.String <> "" Then Return True

            objGoodsReceipt = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objAddOn.objApplication.StatusBar.SetText("Creating Goods Receipt draft. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objedit = objForm.Items.Item("23").Specific
            Dim DocDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)

            If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()

            objGoodsReceipt.DocDate = DocDate
            'objGoodsReceipt.JournalMemo = "Auto-Gen-> " & Now.ToString
            objGoodsReceipt.UserFields.Fields.Item("U_GateRem").Value = "From Gate Inward Auto-Gen-> " & Now.ToString
            objGoodsReceipt.UserFields.Fields.Item("U_PartyId").Value = Trim(objForm.Items.Item("10").Specific.String)
            objGoodsReceipt.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry
            If objAddOn.HANA Then
                strQuery = "Select ""BPLId"" from OBPL where ""Disabled""='N' and ""MainBPL""='Y'" 'Branch
            Else
                strQuery = "Select BPLId from OBPL where Disabled='N' and MainBPL='Y'" 'Branch
            End If
            strQuery = objAddOn.objGenFunc.getSingleValue(strQuery)
            If strQuery <> "" Then objGoodsReceipt.BPL_IDAssignedToInvoice = strQuery
            If objMatrix.VisualRowCount > 0 Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    'objGoodsReceipt.CardCode = Trim(objForm.Items.Item("10").Specific.String)
                    objGoodsReceipt.Lines.ItemCode = Trim(objMatrix.Columns.Item("4").Cells.Item(i).Specific.String)
                    objGoodsReceipt.Lines.Quantity = CDbl(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String) ' CDbl(Matrix0.Columns.Item("grnqty").Cells.Item(i).Specific.String) ' CDbl(objRs.Fields.Item("GRN Qty").Value.ToString) 
                    'objGRPO.Lines.AccountCode = Trim(objRS.Fields.Item("AcctCode").Value.ToString)
                    'objGRPO.Lines.TaxCode = Trim(objRS.Fields.Item("TaxCode").Value.ToString)
                    objGoodsReceipt.Lines.UnitPrice = CDbl(objMatrix.Columns.Item("11").Cells.Item(i).Specific.String)
                    'objGRPO.Lines.WarehouseCode = Trim(objRS.Fields.Item("WhsCode").Value.ToString)
                    objGoodsReceipt.Lines.UserFields.Fields.Item("U_GateQty").Value = CDbl(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String)
                    objGoodsReceipt.Lines.Add()
                Next

            End If

            If objGoodsReceipt.Add() <> 0 Then
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objAddOn.objApplication.SetStatusBarMessage("Goods Receipt: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox("Goods Receipt: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                Return False
            Else
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                DocEntry = objAddOn.objCompany.GetNewObjectKey()
                objForm.Items.Item("52B").Specific.String = DocEntry
                objAddOn.objApplication.StatusBar.SetText("Draft Goods Receipt Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return True
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsReceipt)
            GC.Collect()
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function GetTransaction_Type(ByVal FormUID As String, ByVal Type As String)
        Try
            Select Case Type
                Case "PO"
                    Type = "Purchase Order"
                Case "SR", "IN"
                    Type = "A/R Invoice"
                Case "DR"
                    Type = "Delivery"
                Case "MI"
                    Type = "Item Master"
                Case "JO", "DC"
                    Type = "Inventory Transfer"
                Case "SP"
                    Type = "Purchase Order Service"
                Case "CL"
                    Type = "Contract Labour"
                Case Else
                    Type = ""
            End Select
            Return Type
        Catch ex As Exception

        End Try
    End Function

    Private Function Remove_Draft_Transaction(ByVal FormUID As String, ByVal DocObjCode As SAPbobsCOM.BoObjectTypes, ByVal DraftEntry As String)
        Try
            Dim oDrafts As SAPbobsCOM.Documents
            oDrafts = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            oDrafts.DocObjectCode = DocObjCode 'SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes
            If oDrafts.GetByKey(DraftEntry) Then
                If oDrafts.Remove() <> 0 Then
                    'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objAddOn.objApplication.SetStatusBarMessage("Drafts: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    objAddOn.objApplication.MessageBox("Drafts: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                    Return False
                Else
                    'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    'DocEntry = objAddOn.objCompany.GetNewObjectKey()
                    'objForm.Items.Item("52B").Specific.String = DocEntry
                    objAddOn.objApplication.StatusBar.SetText("Draft Removed Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If objAddOn.HANA Then
                        strQuery = "Update ""@MIGTIN"" Set ""U_trgtentry""=null Where ""DocEntry""=" & objForm.Items.Item("57").Specific.String & ""
                    Else
                        strQuery = "Update [@MIGTIN] Set U_trgtentry=null Where DocEntry=" & objForm.Items.Item("57").Specific.String & ""
                    End If
                    Recordset.DoQuery(strQuery)
                    Return True
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDrafts)
                GC.Collect()
            End If


        Catch ex As Exception

        End Try
    End Function

    Private Sub FieldSetup(ByVal FormUID As String, ByVal Type As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("36").Specific
            objForm.Freeze(True)
            If Type = "CL" Then
                If objMatrix.Columns.Item("1A").Visible = False Then Exit Sub
                objMatrix.Columns.Item("1A").Visible = False
                objMatrix.Columns.Item("1").Visible = False
                objMatrix.Columns.Item("2").Visible = False
                objMatrix.Columns.Item("3").Visible = False
                objMatrix.Columns.Item("4").Visible = False
                objMatrix.Columns.Item("5").Visible = False
                objMatrix.Columns.Item("6").Visible = False
                objMatrix.Columns.Item("7").Visible = False
                objMatrix.Columns.Item("7A").Visible = False
                objMatrix.Columns.Item("8").Visible = False
                objMatrix.Columns.Item("9").Visible = False
                objMatrix.Columns.Item("10").Visible = False
                objMatrix.Columns.Item("11").Visible = False
                objMatrix.Columns.Item("12").Visible = False
                objMatrix.Columns.Item("13").Visible = False
                objMatrix.Columns.Item("14").Visible = False

                objMatrix.Columns.Item("16").Visible = True
                objMatrix.Columns.Item("19").Visible = True
                objMatrix.Columns.Item("20").Visible = True
                objMatrix.Columns.Item("21").Visible = True
                objMatrix.Columns.Item("22").Visible = True
                objMatrix.Columns.Item("23").Visible = True
                objMatrix.Columns.Item("24").Visible = True
                objMatrix.Columns.Item("25").Visible = True
                objMatrix.Columns.Item("26").Visible = True
                objMatrix.Columns.Item("21A").Visible = True
                objMatrix.Columns.Item("21B").Visible = True

                objMatrix.Columns.Item("27").Visible = False
                objMatrix.Columns.Item("28").Visible = False
                objMatrix.Columns.Item("29").Visible = False
                objMatrix.Columns.Item("29A").Visible = False
                objMatrix.Columns.Item("30").Visible = False
                objMatrix.Columns.Item("31").Visible = False
                objMatrix.Columns.Item("32").Visible = False
            ElseIf Type = "SP" Then
                'If objMatrix.Columns.Item("16").Visible = False Then Exit Sub
                objMatrix.Columns.Item("16").Visible = False
                objMatrix.Columns.Item("19").Visible = False
                objMatrix.Columns.Item("20").Visible = False
                objMatrix.Columns.Item("21").Visible = False
                objMatrix.Columns.Item("22").Visible = False
                objMatrix.Columns.Item("23").Visible = False
                objMatrix.Columns.Item("24").Visible = False
                objMatrix.Columns.Item("25").Visible = False
                objMatrix.Columns.Item("26").Visible = False
                objMatrix.Columns.Item("21A").Visible = False
                objMatrix.Columns.Item("21B").Visible = False
                objMatrix.Columns.Item("4").Visible = False
                objMatrix.Columns.Item("5").Visible = False
                objMatrix.Columns.Item("6").Visible = False
                objMatrix.Columns.Item("7").Visible = False
                objMatrix.Columns.Item("7A").Visible = False
                objMatrix.Columns.Item("8").Visible = False
                objMatrix.Columns.Item("9").Visible = False
                objMatrix.Columns.Item("10").Visible = False

                objMatrix.Columns.Item("1A").Visible = True
                objMatrix.Columns.Item("1").Visible = True
                objMatrix.Columns.Item("2").Visible = True
                objMatrix.Columns.Item("3").Visible = True
                objMatrix.Columns.Item("11").Visible = True
                objMatrix.Columns.Item("12").Visible = True
                objMatrix.Columns.Item("13").Visible = True
                objMatrix.Columns.Item("14").Visible = True

                objMatrix.Columns.Item("27").Visible = True
                objMatrix.Columns.Item("28").Visible = True
                objMatrix.Columns.Item("29").Visible = True
                objMatrix.Columns.Item("29A").Visible = True
                objMatrix.Columns.Item("30").Visible = True
                objMatrix.Columns.Item("31").Visible = True
                objMatrix.Columns.Item("32").Visible = True
            Else
                If objMatrix.Columns.Item("16").Visible = False Then Exit Sub
                objMatrix.Columns.Item("16").Visible = False
                objMatrix.Columns.Item("19").Visible = False
                objMatrix.Columns.Item("20").Visible = False
                objMatrix.Columns.Item("21").Visible = False
                objMatrix.Columns.Item("22").Visible = False
                objMatrix.Columns.Item("23").Visible = False
                objMatrix.Columns.Item("24").Visible = False
                objMatrix.Columns.Item("25").Visible = False
                objMatrix.Columns.Item("26").Visible = False
                objMatrix.Columns.Item("21A").Visible = False
                objMatrix.Columns.Item("21B").Visible = False

                objMatrix.Columns.Item("1A").Visible = True
                objMatrix.Columns.Item("1").Visible = True
                objMatrix.Columns.Item("2").Visible = True
                objMatrix.Columns.Item("3").Visible = True
                objMatrix.Columns.Item("4").Visible = True
                objMatrix.Columns.Item("5").Visible = True
                objMatrix.Columns.Item("6").Visible = True
                objMatrix.Columns.Item("7").Visible = True
                objMatrix.Columns.Item("7A").Visible = True
                objMatrix.Columns.Item("8").Visible = True
                objMatrix.Columns.Item("9").Visible = True
                objMatrix.Columns.Item("10").Visible = True
                objMatrix.Columns.Item("11").Visible = True
                objMatrix.Columns.Item("12").Visible = True
                objMatrix.Columns.Item("13").Visible = True
                objMatrix.Columns.Item("14").Visible = True

                objMatrix.Columns.Item("27").Visible = False
                objMatrix.Columns.Item("28").Visible = False
                objMatrix.Columns.Item("29").Visible = False
                objMatrix.Columns.Item("29A").Visible = False
                objMatrix.Columns.Item("30").Visible = False
                objMatrix.Columns.Item("31").Visible = False
                objMatrix.Columns.Item("32").Visible = False
            End If

            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
        End Try
    End Sub

End Class

