Imports System.IO

Imports System.Threading
Namespace Mukesh.SBOLib

    Public Class GeneralFunctions
        Private objCompany As SAPbobsCOM.Company
        Private strThousSep As String = ","
        Private strDecSep As String = "."
        Private intQtyDec As Integer = 3
        Dim BankFileName = ""
        Dim strsql As String


        Public Sub New(ByVal Company As SAPbobsCOM.Company)
            Dim objRS As SAPbobsCOM.Recordset
            objCompany = Company
            If objCompany.Connected Then
                objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRS.DoQuery("SELECT * FROM OADM")
                If Not objRS.EoF Then
                    strThousSep = objRS.Fields.Item("ThousSep").Value
                    strDecSep = objRS.Fields.Item("DecSep").Value
                    intQtyDec = objRS.Fields.Item("QtyDec").Value
                End If
            End If
        End Sub

        Public Function DateCompare(ByVal Date1 As Date, ByVal Date2 As Date) As Integer
            Return Date.Compare(Date1, Date2)
        End Function

        Public Function GetDateTimeValue(ByVal SBODaMIPLAGNTMASring As String) As DateTime
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            objBridge.Format_StringToDate("")
            Return objBridge.Format_StringToDate(SBODaMIPLAGNTMASring).Fields.Item(0).Value
        End Function

        Public Function GetSBODateString(ByVal DateVal As DateTime) As String
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Return objBridge.Format_DateToString(DateVal).Fields.Item(0).Value
        End Function

        Public Function GetSBODaMIPLAGNTMASring(ByVal DateVal As DateTime) As String
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Return objBridge.Format_DateToString(DateVal).Fields.Item(0).Value
        End Function

        Public Function GetQtyValue(ByVal QtyString As String) As Double
            Dim dblValue As Double
            QtyString = QtyString.Replace(strThousSep, "")
            QtyString = QtyString.Replace(strDecSep, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
            dblValue = Convert.ToDouble(QtyString)
            Return dblValue
        End Function

        Public Function GetQtyString(ByVal QtyVal As Double) As String
            GetQtyString = QtyVal.ToString()
            GetQtyString.Replace(",", strDecSep)
        End Function

        Public Function GetCode(ByVal sTableName As String) As String
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery("SELECT Top 1 Code FROM " & sTableName + " ORDER BY Convert(INT,Code) DESC")
            If Not objRS.EoF Then
                Return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString()) + 1
            Else
                GetCode = "1"
            End If
        End Function

        Public Function GetDocNum(ByVal sUDOName As String, ByVal Series As Integer) As String
            Dim StrSQL As String
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If objAddOn.HANA Then
                '  StrSQL = "select ""AutoKey"" from ONNM where ""ObjectCode""='" & sUDOName & "'"
                If Series = 0 Then
                    StrSQL = " select  ""NextNumber""  from NNM1 where ""ObjectCode""='" & sUDOName & "'"
                Else
                    StrSQL = " select  ""NextNumber""  from NNM1 where ""ObjectCode""='" & sUDOName & "' and ""Series"" = " & Series
                End If

            Else
                StrSQL = "select Autokey from onnm where objectcode='" & sUDOName & "'"
            End If

            objRS.DoQuery(StrSQL)
            objRS.MoveFirst()
            If Not objRS.EoF Then
                Return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString())
            Else
                GetDocNum = "1"
            End If
        End Function

        'Public Function GetDocNum_Mbook(ByVal sUDOName As String) As String
        '    objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '    StrSQL = "select Autokey from onnm where objectcode='" & sUDOName & "'"
        '    objRS.DoQuery(StrSQL)
        '    objRS.MoveFirst()
        '    objAddOn.objApplication.MessageBox(objRS.RecordCount)
        '    If objRS.RecordCount > 0 Then
        '        Return objRS.Fields.Item(0).Value.ToString
        '    Else
        '        Return "1"
        '    End If
        'End Function

        Sub setEdittextCFL(ByVal oForm As SAPbouiCOM.Form, ByVal UId As String, ByVal strCFL_ID As String, ByVal strCFL_Obj As String, ByVal strCFL_Alies As String)
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromListCreationParams
                oCFL = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
                oCFL.UniqueID = strCFL_ID
                oCFL.ObjectType = strCFL_Obj
                oForm.ChooseFromLists.Add(oCFL)

                Dim txt As SAPbouiCOM.EditText = oForm.Items.Item(UId).Specific
                txt.ChooseFromListUID = strCFL_ID
                txt.ChooseFromListAlias = strCFL_Alies

            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText("Set EditText CFL Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub

        Function SetAttachMentFile(ByVal oForm As SAPbouiCOM.Form, ByVal oDBDSHeader As SAPbouiCOM.DBDataSource, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource) As Boolean
            Try
                If objCompany.AttachMentPath.Length <= 0 Then
                    objAddOn.objApplication.StatusBar.SetText("Attchment folder not defined, or Attchment folder has been changed or removed. [Message 131-102]")
                    Return False
                End If

                Dim strFileName As String = FindFile()
                If strFileName.Equals("") = False Then
                    Dim FileExist() As String = strFileName.Split("\")
                    Dim FileDestPath As String = objCompany.AttachMentPath & FileExist(FileExist.Length - 1)

                    If File.Exists(FileDestPath) Then
                        Dim LngRetVal As Long = objAddOn.objApplication.MessageBox("A file with this name already exists,would you like to replace this?  " & FileDestPath & " will be replaced.", 1, "Yes", "No")
                        If LngRetVal <> 1 Then Return False
                    End If
                    Dim fileNameExt() As String = FileExist(FileExist.Length - 1).Split(".")
                    Dim ScrPath As String = objCompany.AttachMentPath
                    ScrPath = ScrPath.Substring(0, ScrPath.Length - 1)
                    Dim TrgtPath As String = strFileName.Substring(0, strFileName.LastIndexOf("\"))

                    oMatrix.AddRow()
                    oMatrix.FlushToDataSource()
                    oDBDSAttch.Offset = oDBDSAttch.Size - 1
                    oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset, oMatrix.VisualRowCount)
                    oDBDSAttch.SetValue("U_trgtPath", oDBDSAttch.Offset, ScrPath)
                    oDBDSAttch.SetValue("U_SrcPath", oDBDSAttch.Offset, TrgtPath)
                    oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, fileNameExt(0))
                    oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, fileNameExt(1))
                    oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, GetServerDate())
                    oMatrix.SetLineData(oDBDSAttch.Size)
                    oMatrix.FlushToDataSource()
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
                Return True
            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText("Set AttachMent File Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            Finally
            End Try
        End Function

        Sub OpenAttachment(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal PvalRow As Integer)
            Try
                If PvalRow <= oMatrix.VisualRowCount And PvalRow <> 0 Then
                    Dim RowIndex As Integer = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) - 1
                    Dim strServerPath, strClientPath As String

                    strServerPath = Trim(oDBDSAttch.GetValue("U_trgtPath", RowIndex)) + "\" + Trim(oDBDSAttch.GetValue("U_FileName", RowIndex)) + "." + Trim(oDBDSAttch.GetValue("U_FileExt", RowIndex))
                    strClientPath = Trim(oDBDSAttch.GetValue("U_SrcPath", RowIndex)) + "\" + Trim(oDBDSAttch.GetValue("U_FileName", RowIndex)) + "." + Trim(oDBDSAttch.GetValue("U_FileExt", RowIndex))
                    'Open Attachment File
                    OpenFile(strServerPath, strClientPath)
                End If

            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText("OpenAttachment Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub

        Public Sub OpenFile(ByVal ServerPath As String, ByVal ClientPath As String)
            Try
                Dim oProcess As System.Diagnostics.Process = New System.Diagnostics.Process
                Try
                    oProcess.StartInfo.FileName = ServerPath
                    oProcess.Start()
                Catch ex1 As Exception
                    Try
                        oProcess.StartInfo.FileName = ClientPath
                        oProcess.Start()
                    Catch ex2 As Exception
                        objAddOn.objApplication.StatusBar.SetText("" & ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
                Finally
                End Try
            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText("" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub

        Sub DeleteRowAttachment(ByVal oForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal SelectedRowID As Integer)
            Try
                oDBDSAttch.RemoveRecord(SelectedRowID - 1)
                oMatrix.DeleteRow(SelectedRowID)
                oMatrix.FlushToDataSource()

                For i As Integer = 1 To oMatrix.VisualRowCount
                    oMatrix.GetLineData(i)
                    oDBDSAttch.Offset = i - 1

                    oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset, i)
                    oDBDSAttch.SetValue("U_trgtPath", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("TrgtPath").Cells.Item(i).Specific.Value))
                    oDBDSAttch.SetValue("U_SrcPath", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("Path").Cells.Item(i).Specific.Value))
                    oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("FileName").Cells.Item(i).Specific.Value))
                    oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("FileExt").Cells.Item(i).Specific.Value))
                    oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("Date").Cells.Item(i).Specific.Value))
                    oMatrix.SetLineData(i)
                    oMatrix.FlushToDataSource()
                Next
                'oDBDSAttch.RemoveRecord(oDBDSAttch.Size - 1)
                oMatrix.LoadFromDataSource()

                oForm.Items.Item("Btndisp").Enabled = False
                oForm.Items.Item("Btndel").Enabled = False

                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText("DeleteRowAttachment Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub

        Function GetServerDate() As String
            Try
                Dim rsetBob As SAPbobsCOM.SBObob = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
                Dim rsetServerDate As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                rsetServerDate = rsetBob.Format_StringToDate(objAddOn.objApplication.Company.ServerDate())

                Return CDate(rsetServerDate.Fields.Item(0).Value).ToString("yyyyMMdd")
                'Return "20120215"
            Catch ex As Exception
                Return ""
            Finally
            End Try
        End Function

        Function LoadComboBoxSeries(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal UDOID As String) As Boolean
            Try
                oComboBox.ValidValues.LoadSeries(UDOID, SAPbouiCOM.BoSeriesMode.sf_Add)
                oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText("LoadComboBoxSeries Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            Finally
            End Try
            Return True
        End Function

        Public Function FindFile() As String
            Dim ShowFolderBrowserThread As Threading.Thread
            Try
                ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)

                If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                    ShowFolderBrowserThread.Start()
                ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                    ShowFolderBrowserThread.Start()
                    ShowFolderBrowserThread.Join()
                End If

                While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                    System.Windows.Forms.Application.DoEvents()
                    'ShowFolderBrowserThread.Sleep(100)
                    Thread.Sleep(100)
                End While

                If BankFileName <> "" Then
                    Return BankFileName
                End If

            Catch ex As Exception

                objAddOn.objApplication.MessageBox("File Find  Method Failed : " & ex.Message)
            End Try
            Return ""
        End Function

        Public Sub ShowFolderBrowser()
            Dim MyProcs() As System.Diagnostics.Process
            Dim OpenFile As New OpenFileDialog
            Try
                OpenFile.Multiselect = False
                OpenFile.Filter = "All files(*.)|*.*" '   "|*.*"
                Dim filterindex As Integer = 0
                Try
                    filterindex = 0
                Catch ex As Exception
                End Try
                OpenFile.FilterIndex = filterindex
                'OpenFile.RestoreDirectory = True
                OpenFile.InitialDirectory = objCompany.AttachMentPath ' "\\newton.tmicloud.net\DB4SHARE\OEC_TEST\Attachments\"
                MyProcs = Process.GetProcessesByName("SAP Business One")

                ' *******  Modified on 09-Mar-2012 By parthiban ********

                ' If two or more company opened at the same time,  Dialog is  not opening 
                ' Changed Conditon   to >= 1
                ' Added Condition --If comname(1).ToString.Trim.ToUpper = com Then -- to open dialog
                ' only for this company

                'If MyProcs.Length = 1 Then
                If MyProcs.Length >= 1 Then

                    For i As Integer = 0 To MyProcs.Length - 1
                        Dim comname As String() = MyProcs(i).MainWindowTitle.ToString.Split("-")
                        If comname(1) = "" Then Continue For
                        'Open dialog only for the company where the button is clicked
                        Dim com As String = objCompany.CompanyName.ToString.Trim.ToUpper
                        If comname(1).ToString.Trim.ToUpper = com Then
                            Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
                            Dim ret As Windows.Forms.DialogResult = OpenFile.ShowDialog(MyWindow)
                            If ret = Windows.Forms.DialogResult.OK Then

                                BankFileName = OpenFile.FileName
                                'OpenFile.Dispose()
                            Else
                                System.Windows.Forms.Application.ExitThread()
                            End If
                        End If
                    Next
                Else
                End If
            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText(ex.Message)
                BankFileName = ""
            Finally
                OpenFile.Dispose()
            End Try
        End Sub

        Public Function getSingleValue(ByVal StrSQL As String) As String
            Try
                Dim rset As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strReturnVal As String = ""
                rset.DoQuery(StrSQL)
                Return IIf(rset.RecordCount > 0, rset.Fields.Item(0).Value.ToString(), "")
            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText(" Get Single Value Function Failed :  " & ex.Message + StrSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return ""
            End Try
        End Function

        Function DoQuery(ByVal strSql As String) As SAPbobsCOM.Recordset
            Try
                Dim rsetCode As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsetCode.DoQuery((strSql))

                Return rsetCode
            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText("Execute Query Function Failed:" & ex.Message + strSql, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return Nothing
            Finally
            End Try
        End Function

        Public Class WindowWrapper

            Implements System.Windows.Forms.IWin32Window
            Private _hwnd As IntPtr

            Public Sub New(ByVal handle As IntPtr)
                _hwnd = handle
            End Sub

            Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
                Get
                    Return _hwnd
                End Get
            End Property

        End Class

        Public Sub SetAutomanagedattribute_Editable(ByVal oform As SAPbouiCOM.Form, ByVal fieldname As String, ByVal add As Boolean, ByVal find As Boolean, ByVal update As Boolean)

            If add = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If find = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If update Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If
        End Sub

        Public Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub setReport(ByVal FormType As String, ByVal Layout As String, ByVal FormCount As Integer)
            Try
                Dim objform As SAPbouiCOM.Form
                'objform = objaddon.objapplication.Forms.Item(FormUID)
                objform = objAddOn.objApplication.Forms.GetForm(FormType, FormCount) '"SUBCTPO"
                Dim rptTypeService As SAPbobsCOM.ReportTypesService
                'Dim newType As SAPbobsCOM.ReportType
                Dim newtypesParam As SAPbobsCOM.ReportTypesParams
                rptTypeService = objAddOn.objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
                newtypesParam = rptTypeService.GetReportTypeList
                Dim TypeCode As String
                If objAddOn.HANA Then
                    TypeCode = getSingleValue("Select ""CODE"" from RTYP where ""NAME""='" & Layout & "'")
                Else
                    TypeCode = getSingleValue("Select CODE from RTYP where NAME='" & Layout & "'")
                End If
                objform.ReportType = TypeCode
                'Dim i As Integer
                'For i = 0 To newtypesParam.Count - 1
                '    If newtypesParam.Item(i).TypeName = FormType And newtypesParam.Item(i).MenuID = FormType Then
                '        objform.ReportType = newtypesParam.Item(i).TypeCode
                '        Exit For
                '    End If
                'Next i
            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText("setReport Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Public Sub AddToPermissionTree(ByVal Name As String, ByVal PermissionID As String, ByVal FormType As String, ByVal ParentID As String, ByVal AddPermission As Char)
            Try
                Dim RetVal As Long
                Dim ErrMsg As String = ""
                Dim oPermission As SAPbobsCOM.UserPermissionTree
                Dim objBridge As SAPbobsCOM.SBObob
                Dim objrs As SAPbobsCOM.Recordset
                If ParentID <> "" Then

                    If objAddOn.HANA = True Then
                        strsql = objAddOn.objGenFunc.getSingleValue("Select 1 as ""Status"" from OUPT Where ""AbsId""='" & ParentID & "'")
                    Else
                        strsql = objAddOn.objGenFunc.getSingleValue("Select 1 as Status from OUPT Where AbsId='" & ParentID & "'")
                    End If

                    If strsql = "" Then Return
                End If

                oPermission = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
                objBridge = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
                objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs = objBridge.GetUserList()

                If oPermission.GetByKey(PermissionID) = False Then
                    oPermission.Name = Name
                    oPermission.PermissionID = PermissionID
                    oPermission.UserPermissionForms.FormType = FormType
                    If ParentID <> "" Then oPermission.ParentID = ParentID
                    oPermission.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
                    RetVal = oPermission.Add()
                    Dim temp_int As Integer = CInt((RetVal))
                    Dim temp_string As String = ErrMsg
                    objAddOn.objCompany.GetLastError(temp_int, temp_string)

                    If RetVal <> 0 Then
                        objAddOn.objApplication.StatusBar.SetText("AddToPermissionTree: " & temp_string, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        If AddPermission = "N"c Then Return
                        For i As Integer = 0 To objrs.RecordCount
                            If objAddOn.HANA = True Then
                                strsql = "Select ""USERID"" from OUSR Where ""USER_CODE""='" & Convert.ToString(objrs.Fields.Item(0).Value) & "'"
                            Else
                                strsql = "Select USERID from OUSR Where USER_CODE='" & Convert.ToString(objrs.Fields.Item(0).Value) & "'"
                            End If
                            strsql = objAddOn.objGenFunc.getSingleValue(strsql)
                            objAddOn.objGenFunc.AddPermissionToUsers(Convert.ToInt32(strsql), PermissionID)
                            objrs.MoveNext()
                        Next
                    End If
                    'Else
                    '    oPermission.Remove()
                End If

            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText("Permission: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Sub AddPermissionToUsers(ByVal UserCode As Integer, ByVal PermissionID As String)
            Try
                Dim oUser As SAPbobsCOM.Users = Nothing
                Dim lRetCode As Integer
                Dim sErrMsg As String = ""
                oUser = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)

                If oUser.GetByKey(UserCode) = True Then
                    oUser.UserPermission.Add()
                    oUser.UserPermission.SetCurrentLine(0)
                    oUser.UserPermission.PermissionID = PermissionID
                    oUser.UserPermission.Permission = SAPbobsCOM.BoPermission.boper_Full
                    lRetCode = oUser.Update()
                    objAddOn.objCompany.GetLastError(lRetCode, sErrMsg)

                    If lRetCode <> 0 Then
                        objAddOn.objApplication.StatusBar.SetText("AddPermissionToUser: " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If

            Catch ex As Exception
            End Try
        End Sub

        Public Sub Matrix_Addrow(ByVal omatrix As SAPbouiCOM.Matrix, Optional ByVal colname As String = "", Optional ByVal rowno_name As String = "", Optional ByVal Error_Needed As Boolean = False)
            Try
                Dim addrow As Boolean = False

                If omatrix.VisualRowCount = 0 Then addrow = True : GoTo addrow
                If colname = "" Then addrow = True : GoTo addrow
                If omatrix.Columns.Item(colname).Cells.Item(omatrix.VisualRowCount).Specific.string <> "" Then addrow = True : GoTo addrow

addrow:
                If addrow = True Then
                    omatrix.AddRow(1)
                    omatrix.ClearRowData(omatrix.VisualRowCount)
                    If rowno_name <> "" Then omatrix.Columns.Item("0").Cells.Item(omatrix.VisualRowCount).Specific.string = omatrix.VisualRowCount
                Else
                    If Error_Needed = True Then objAddOn.objApplication.SetStatusBarMessage("Already Empty Row Available", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
            Catch ex As Exception

            End Try
        End Sub

    End Class

End Namespace