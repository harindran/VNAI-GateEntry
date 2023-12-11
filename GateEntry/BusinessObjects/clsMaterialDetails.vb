Public Class clsMaterialDetails

    Public Const Formtype = "MIMATDET"
    Dim objForm, objGIForm As SAPbouiCOM.Form
    Dim odbdsHeader, odbdsLine As SAPbouiCOM.DBDataSource
    Dim objRS As SAPbobsCOM.Recordset
    Dim objMatrix, objGIMatrix As SAPbouiCOM.Matrix
    Dim strSQL As String
    Dim Row As Integer
    Dim strRejDetails As String = ""
    Dim objCombo As SAPbouiCOM.ComboBox

    Public Sub LoadScreen(ByRef GIFormUID As String, ByVal RowID As Integer, Optional MatEntry As String = "")
        Try
            objForm = objAddOn.objUIXml.LoadSingleScreenXML("MaterialDetails.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
            objMatrix = objForm.Items.Item("13").Specific
            If MatEntry <> "" Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("4A").Enabled = True
                objForm.Items.Item("4A").Specific.String = MatEntry
                objForm.ActiveItem = "6"
                objForm.Items.Item("4A").Enabled = False
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Items.Item("18").Enabled = False
            Else
                objGIForm = objAddOn.objApplication.Forms.Item(GIFormUID)
                Row = RowID
                odbdsHeader = objForm.DataSources.DBDataSources.Item("@MIMATDET")
                odbdsLine = objForm.DataSources.DBDataSources.Item("@MIMATDET1")

                objCombo = objForm.Items.Item("18").Specific
                objCombo.ValidValues.LoadSeries(Formtype, SAPbouiCOM.BoSeriesMode.sf_Add)
                If objCombo.ValidValues.Count > 0 Then objCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Try
                    strSQL = objForm.BusinessObject.GetNextSerialNumber(Trim(objForm.Items.Item("18").Specific.Selected.value), objForm.BusinessObject.Type)
                Catch ex As Exception
                    objAddOn.objApplication.MessageBox("To generate this document, first define the numbering series in the Administration module")
                    Exit Sub
                End Try
                If strSQL <> "" Then odbdsHeader.SetValue("DocNum", 0, strSQL)
                'If objAddOn.HANA Then
                '    odbdsHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetNextDocEntryValue("@MIMATDET"))
                'Else
                '    odbdsHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetNextDocEntryValue("[@MIMATDET]"))
                'End If
                objAddOn.objGenFunc.Matrix_Addrow(objMatrix, "1", "0")
                objForm.Items.Item("6").Specific.Active = True
                objForm.Items.Item("6").Specific.String = "A"
                objGIMatrix = objGIForm.Items.Item("36").Specific
                odbdsHeader.SetValue("U_LabID", 0, objGIMatrix.Columns.Item("15").Cells.Item(RowID).Specific.string)
                odbdsHeader.SetValue("U_LabName", 0, objGIMatrix.Columns.Item("16").Cells.Item(RowID).Specific.string)
                odbdsHeader.SetValue("U_GINo", 0, objGIForm.Items.Item("21").Specific.String)
                objMatrix.Columns.Item("2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objMatrix.Columns.Item("1").Cells.Item(1).Click()
            End If
            bModal = True
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Load Screen: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub DeleteEmptyRowInFormDataEvent(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal ColumnUID As String, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource)
        Try
            If oMatrix.VisualRowCount > 0 Then
                If oMatrix.Columns.Item(ColumnUID).Cells.Item(oMatrix.VisualRowCount).Specific.Value.Equals("") Then
                    oMatrix.DeleteRow(oMatrix.VisualRowCount)
                    oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1)
                    oMatrix.FlushToDataSource()
                End If
            End If
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Delete Empty RowIn Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            If pVal.BeforeAction Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            If validate(FormUID) = False Then
                                System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                                Exit Sub
                            Else
                                strRejDetails = odbdsHeader.GetValue("DocEntry", 0)
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        If pVal.ItemUID = "13" And pVal.ColUID = "2" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If objMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific.String = "" Then objMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific.String = "0"
                            If objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.String <> "" And CDbl(objMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific.String) <= 0 Then
                                objAddOn.objApplication.StatusBar.SetText("Value in ""Quantity"" cannot be zero.  Line: " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                objMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific.String = CDbl(1)
                            End If
                        End If

                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then
                            If objGIMatrix.Columns.Item("21B").Cells.Item(Row).Specific.String <> "" Then
                                objForm.Close()
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then

                        End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If pVal.ItemUID = "13" And pVal.ColUID = "1" Then
                            objMatrix = objForm.Items.Item("13").Specific
                            objMatrix.Columns.Item("2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                            If objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String <> "" Then
                                objAddOn.objGenFunc.Matrix_Addrow(objMatrix, "1", "0")
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = "18" Then
                            If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                            objCombo = objForm.Items.Item("18").Specific
                            Dim StrDocNum = objForm.BusinessObject.GetNextSerialNumber(Trim(objForm.Items.Item("18").Specific.Selected.value), objForm.BusinessObject.Type)
                            objForm.DataSources.DBDataSources.Item("@MIMATDET").SetValue("DocNum", 0, StrDocNum)
                        End If

                End Select
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            objMatrix = objForm.Items.Item("13").Specific
            odbdsHeader = objForm.DataSources.DBDataSources.Item("@MIMATDET")
            odbdsLine = objForm.DataSources.DBDataSources.Item("@MIMATDET1")
            If BusinessObjectInfo.BeforeAction Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                        If validate(objForm.UniqueID) = False Then
                            System.Media.SystemSounds.Asterisk.Play()
                            BubbleEvent = False
                            Exit Sub
                        Else
                            DeleteEmptyRowInFormDataEvent(objMatrix, "3", odbdsLine)
                            objGIMatrix.Columns.Item("21B").Cells.Item(Row).Specific.String = strRejDetails
                            objGIMatrix.CommonSetting.SetCellEditable(Row, 25, False)
                        End If
                End Select
            Else
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        strRejDetails = odbdsHeader.GetValue("DocEntry", 0)
                        objGIMatrix.Columns.Item("21B").Cells.Item(Row).Specific.String = strRejDetails
                        If (objGIForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or objGIForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then objGIForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE : objGIForm.Items.Item("1").Click()
                        objGIMatrix.CommonSetting.SetCellEditable(Row, 25, False)
                End Select
            End If
        Catch ex As Exception

        End Try


    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "1284"
                Case "1289", "1290", "1291", "1288", "1282"
                    If pVal.BeforeAction = True Then
                        BubbleEvent = False
                    End If
                Case "1281" 'Find Mode
                    If pVal.BeforeAction = False Then
                        objForm.Items.Item("8").Enabled = True
                        objForm.Items.Item("10").Enabled = True
                        objForm.Items.Item("4").Enabled = True
                        objForm.Items.Item("6").Enabled = True
                        objForm.Items.Item("12").Enabled = True
                    End If

                Case "1293"  'delete Row
                Case "ditem"
                    objForm = objAddOn.objApplication.Forms.ActiveForm
                    objGIMatrix = objForm.Items.Item("13").Specific
                    If objAddOn.ZB_row > 0 Then
                        objGIMatrix.DeleteRow(objAddOn.ZB_row)
                    End If
            End Select

        Catch ex As Exception
            ' objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub

    Private Function validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        'objForm = objAddOn.objApplication.Forms.ActiveForm
        objMatrix = objForm.Items.Item("13").Specific
        'odbdsHeader = objForm.DataSources.DBDataSources.Item("@MIREJDET")
        If objMatrix.Columns.Item("1").Cells.Item(1).Specific.string = "" Then
            objAddOn.objApplication.MessageBox("Please update the Material Details!!!",, "OK")
            objAddOn.objApplication.StatusBar.SetText("Please update the Material Details!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        'For i As Integer = 1 To objMatrix.RowCount
        '    If objMatrix.Columns.Item("1").Cells.Item(i).Specific.string <> "" Then
        '        If CDbl(objMatrix.Columns.Item("2").Cells.Item(i).Specific.string) > 0 And objMatrix.Columns.Item("3").Cells.Item(i).Specific.string.trim = "" Then
        '            objAddOn.objApplication.MessageBox("Please Update the Reason!!!",, "OK")
        '            Return False
        '        End If
        '    End If
        'Next

        Return True
    End Function

End Class
