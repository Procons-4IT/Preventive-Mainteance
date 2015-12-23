'PM Date Logic
'Update PMSchedule in Service Contract.
'Validate AR Invoice Creation
'Work Flow Testing
'Testing
'Delivery
'New Class file for Service.
'Check Box for No Bill
'Combo Box for Status.

Public Class clsPreventiveMaintenance
    Inherits clsBase

    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private objForm As SAPbouiCOM.Form
    Private oEditText As SAPbouiCOM.EditText
    Private oMode As SAPbouiCOM.BoFormMode
    Private InvForConsumedItems, count As Integer
    Private blnFlag As Boolean = False
    Private oMatrix As SAPbouiCOM.Matrix
    Private oCombo As SAPbouiCOM.ComboBox
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String = String.Empty
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines As SAPbouiCOM.DBDataSource
    Private RowtoDelete As Integer
    Private MatrixId As String
    Public intSelectedMatrixrow As Integer = 0
    Dim oGrid As SAPbouiCOM.Grid
    Dim strDateFormat As String = "yyyy-MM-dd"
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OPMT) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oForm = oApplication.Utilities.LoadForm(xml_Z_OPMT, frm_Z_OPMT)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            dataBind(oForm)
            loadComboColumn(oForm)
            addChooseFromListConditions(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, False)
            oForm.EnableMenu(mnu_DELETE_ROW, False)
            oForm.EnableMenu("520", False)
            oForm.Settings.EnableRowFormat = True
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub LoadForm(ByVal strRef As String)
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OPMT) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oForm = oApplication.Utilities.LoadForm(xml_Z_OPMT, frm_Z_OPMT)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            dataBind(oForm)
            loadComboColumn(oForm)
            addChooseFromListConditions(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, False)
            oForm.EnableMenu(mnu_DELETE_ROW, False)
            oForm.Settings.EnableRowFormat = True
            oForm.PaneLevel = 1
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("6").Specific.value = strRef
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OPMT Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then
                                    If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                        If Not validation(oForm) Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        If Not validation_PM(oForm) Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Sure you want to Add Preventive Maintanence Document ?", 2, "Yes", "No", "")
                                        If _retVal = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf (oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                        oMatrix = oForm.Items.Item("3").Specific

                                        Dim blnValidate As Boolean = True
                                        Dim blnValidate1 As Boolean = True
                                        For index As Integer = 1 To oMatrix.RowCount
                                            Dim dblAmount As Double = CDbl(oApplication.Utilities.getMatrixValues(oMatrix, "V_11", index).ToString())
                                            Dim dblBAmount As Double = CDbl(oApplication.Utilities.getMatrixValues(oMatrix, "V_13", index).ToString())

                                            If Not CType(oMatrix.Columns.Item("V_12").Cells().Item(index).Specific, SAPbouiCOM.CheckBox).Checked Then

                                                'Need to Add @ Edit Condition.
                                                If dblBAmount > dblAmount Then
                                                    'oApplication.Utilities.Message("Bill Amount Should be Less then Or Equal to Amount...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    'BubbleEvent = False
                                                    blnValidate = False
                                                    Exit For
                                                ElseIf dblBAmount < dblAmount Then
                                                    blnValidate1 = False
                                                    Exit For
                                                ElseIf dblBAmount < 0 Then
                                                    oApplication.Utilities.Message("Bill Amount Should be Greater than Zero...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    BubbleEvent = False
                                                    ' blnValidate = False
                                                    Exit For
                                                ElseIf dblBAmount = 0 Then
                                                    oApplication.Utilities.Message("Please check No Bill if Bill amount is zero...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    'BubbleEvent = False
                                                    blnValidate = False
                                                    Exit For
                                                End If
                                                'ElseIf CType(oMatrix.Columns.Item("V_12").Cells().Item(index).Specific, SAPbouiCOM.CheckBox).Checked Then
                                                '    If dblBAmount <> 0 Then
                                                '        oApplication.Utilities.Message("Bill Amount Should be Equal to Zero if No Bill...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                '        BubbleEvent = False
                                                '        Exit For
                                                '    End If
                                            End If
                                        Next
                                        If Not blnValidate Then
                                            Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Bill Amount should be Less than or Equal to the Service Contract Amount, Wish to proceed ?", 2, "Yes", "No", "")
                                            If _retVal = 2 Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        Else
                                            If Not blnValidate1 Then
                                                Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Bill Amount Lesser than Service Contract Amount, Wish to proceed ?", 2, "Yes", "No", "")
                                                If _retVal = 2 Then
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "34" Then
                                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oApplication.Utilities.Message("This action support only in OK Mode", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Exit Sub
                                    End If
                                    If Not validation_AR(oForm) Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim strDocEntry As String = CType(oForm.Items.Item("25").Specific, SAPbouiCOM.EditText).Value
                                        strQuery = "Select U_InvDE,U_BLDueDt From [@Z_OPMT] Where DocEntry ='" + strDocEntry.Trim() + "'"
                                        oRecordSet.DoQuery(strQuery)
                                        If Not oRecordSet.EoF Then
                                            oApplication.Utilities.Message("Inside the RecordSet....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            If oRecordSet.Fields.Item(0).Value.ToString().Length = 0 Then
                                                Dim dtBlDueDate As DateTime = Convert.ToDateTime(oRecordSet.Fields.Item(1).Value)
                                                oApplication.Utilities.Message("Inside the Condition....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Wish to create the Sale Order ?", 2, "Continue", "Cancel", "")
                                                If _retVal = 2 Then
                                                    Exit Sub
                                                End If

                                                If (oApplication.Utilities.AddARInvoice(oForm, dtBlDueDate)) Then
                                                    oApplication.SBO_Application.MessageBox("Sales Order Created Successfully...")
                                                    oApplication.SBO_Application.Menus.Item(mnu_ADD).Activate()
                                                End If
                                            Else
                                                oApplication.Utilities.Message("Sales Order Already Generated....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "7" Then
                                    'oEditText = oForm.Items.Item("7").Specific
                                    'filterEventLabelChooseFromList(oForm, oEditText.ChooseFromListUID)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    If pVal.ItemUID = "3" And pVal.ColUID = "V_13" And pVal.Row > -1 Then
                                        oMatrix = oForm.Items.Item("3").Specific
                                        Dim dblAmount As Double = CDbl(oApplication.Utilities.getMatrixValues(oMatrix, "V_11", pVal.Row).ToString())
                                        Dim dblBAmount As Double = CDbl(oApplication.Utilities.getMatrixValues(oMatrix, "V_13", pVal.Row).ToString())
                                        If dblBAmount > dblAmount Then
                                            Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Bill Amount Should be Less then Or Equal to Service Contract Amount, Wish to proceed ?", 2, "Yes", "No", "")
                                            If _retVal = 2 Then
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_13", pVal.Row, dblAmount)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        ElseIf dblBAmount < dblAmount Then
                                            Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Bill Amount Lesser than Service Contract Amount , Wish to proceed ?", 2, "Yes", "No", "")
                                            If _retVal = 2 Then
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_13", pVal.Row, dblAmount)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "14"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "15"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                    Case "1"
                                        If pVal.Action_Success Then
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                initialize(oForm)
                                            End If
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPMT")
                                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PMT1")
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim strContractID, strCardCode, strCardName, strNBillDate, strStartDt, strEndDt, strPMScheduleNo, strContact, strConDesc As String
                                Dim intFrequency As Integer
                                Dim strPMStartDt, strPMEndDt As String

                                Dim sCHFL_ID As String
                                Dim intChoice As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "13" Then

                                            strContractID = oDataTable.GetValue("ContractID", 0)
                                            strCardCode = oDataTable.GetValue("CstmrCode", 0)
                                            strCardName = oDataTable.GetValue("CstmrName", 0)
                                            strNBillDate = Convert.ToDateTime(oDataTable.GetValue("U_NBillDt", 0)).ToString("yyyyMMdd")
                                            strStartDt = Convert.ToDateTime(oDataTable.GetValue("StartDate", 0)).ToString("yyyyMMdd")
                                            strEndDt = Convert.ToDateTime(oDataTable.GetValue("EndDate", 0)).ToString("yyyyMMdd")
                                            strPMScheduleNo = oDataTable.GetValue("U_PMSchNo", 0)
                                            intFrequency = CInt(oDataTable.GetValue("U_Freqency", 0))
                                            strContact = oDataTable.GetValue("CntctCode", 0)
                                            strConDesc = oDataTable.GetValue("Descriptio", 0)

                                            oDBDataSource.SetValue("U_SConNo", oDBDataSource.Offset, strContractID)
                                            oDBDataSource.SetValue("U_CardCode", oDBDataSource.Offset, strCardCode)
                                            oDBDataSource.SetValue("U_CardName", oDBDataSource.Offset, strCardName)
                                            oDBDataSource.SetValue("U_SEStartDt", oDBDataSource.Offset, strStartDt)
                                            oDBDataSource.SetValue("U_SEEndDt", oDBDataSource.Offset, strEndDt)
                                            oDBDataSource.SetValue("U_BLDueDt", oDBDataSource.Offset, strNBillDate)
                                            oDBDataSource.SetValue("U_PMSchNo", oDBDataSource.Offset, CInt(strPMScheduleNo) + 1)

                                            'oForm.DataSources.UserDataSources.Item("udsContID").ValueEx = strContact
                                            oForm.DataSources.UserDataSources.Item("udsContDes").ValueEx = strConDesc

                                            calculateDates(oForm, intFrequency, strPMStartDt, strPMEndDt)
                                            oDBDataSource.SetValue("U_PMStartDt", oDBDataSource.Offset, strPMStartDt)
                                            oDBDataSource.SetValue("U_PMEndDt", oDBDataSource.Offset, strPMEndDt)


                                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            strQuery = " Select T0.ItemCode,T0.ItemName,T0.InsID,T0.ManufSN,T0.InternalSN,Convert(VarChar(8),T0.StartDate,112) As StartDate, "
                                            strQuery += " Convert(VarChar(8),T0.EndDate,112) As EndDate,Convert(VarChar(8),T0.TermDate,112) As TermDate,U_QuarterAmt, "
                                            strQuery += " (T2.lastName +','+T2.firstName) As 'technician' From CTR1 T0 JOIN OINS T1 On  T0.InsID = T1.InsID "
                                            strQuery += " LEFT OUTER JOIN OHEM T2 ON T1.technician = T2.EmpID "
                                            strQuery += " Where ContractID = '" + strContractID + "'"
                                            strQuery += " Order by T0.Line "
                                            oRecordSet.DoQuery(strQuery)
                                            If Not oRecordSet.EoF Then
                                                oMatrix = oForm.Items.Item("3").Specific
                                                oMatrix.Clear()
                                                oMatrix.FlushToDataSource()
                                                oMatrix.LoadFromDataSource()
                                                Dim intAddRows As Integer = oRecordSet.RecordCount
                                                If intAddRows > 1 Then
                                                    'intAddRows -= 1
                                                    oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                                End If
                                                oMatrix.FlushToDataSource()
                                                Dim index As Integer = 1
                                                While Not oRecordSet.EoF
                                                    oDBDataSourceLines.SetValue("LineId", index - 1, (pVal.Row + index).ToString())
                                                    oDBDataSourceLines.SetValue("U_ItemCode", index - 1, oRecordSet.Fields.Item("ItemCode").Value)
                                                    oDBDataSourceLines.SetValue("U_ItemName", index - 1, oRecordSet.Fields.Item("ItemName").Value)
                                                    oDBDataSourceLines.SetValue("U_InsID", index - 1, oRecordSet.Fields.Item("InsID").Value)
                                                    oDBDataSourceLines.SetValue("U_MSerialNo", index - 1, oRecordSet.Fields.Item("ManufSN").Value)
                                                    oDBDataSourceLines.SetValue("U_SerialNo", index - 1, oRecordSet.Fields.Item("InternalSN").Value)
                                                    oDBDataSourceLines.SetValue("U_SEStartDt", index - 1, oRecordSet.Fields.Item("StartDate").Value)
                                                    oDBDataSourceLines.SetValue("U_SEEndDt", index - 1, oRecordSet.Fields.Item("EndDate").Value)
                                                    oDBDataSourceLines.SetValue("U_TerDt", index - 1, oRecordSet.Fields.Item("TermDate").Value)
                                                    oDBDataSourceLines.SetValue("U_Amount", index - 1, oRecordSet.Fields.Item("U_QuarterAmt").Value)
                                                    If oRecordSet.Fields.Item("TermDate").Value.ToString = "" Then
                                                        oDBDataSourceLines.SetValue("U_BillAmt", index - 1, oRecordSet.Fields.Item("U_QuarterAmt").Value)
                                                    Else
                                                        Dim dtEnddate As Integer = CInt(oRecordSet.Fields.Item("EndDate").Value)
                                                        Dim dtTerDate As Integer = CInt(oRecordSet.Fields.Item("TermDate").Value)
                                                        If dtEnddate <= dtTerDate Then
                                                            oDBDataSourceLines.SetValue("U_BillAmt", index - 1, oRecordSet.Fields.Item("U_QuarterAmt").Value)
                                                        End If
                                                    End If
                                                    oDBDataSourceLines.SetValue("U_Technician", index - 1, oRecordSet.Fields.Item("technician").Value)
                                                    oRecordSet.MoveNext()
                                                    index += 1
                                                End While
                                                oMatrix.LoadFromDataSource()
                                                oMatrix.FlushToDataSource()
                                            End If
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Try
                                        reDrawForm(oForm)
                                    Catch ex As Exception

                                    End Try
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.BeforeAction
                Case True
                    Select Case pVal.MenuUID
                        Case mnu_CANCEL, mnu_CLOSE
                            BubbleEvent = False
                        Case mnu_PMCancel
                            If oApplication.SBO_Application.MessageBox("Do You Want to Cancel Preventive Maintainence Document?", 2, "Yes", "No") = 2 Then
                                BubbleEvent = False
                                Exit Sub
                            Else
                                Dim DocEntry As String = oDBDataSource.GetValue("DocEntry", oDBDataSource.Offset)
                                changeStatus(oForm, "L")
                                oApplication.Utilities.CancelServiceCall(oForm, DocEntry)
                                oApplication.SBO_Application.Menus.Item(mnu_ADD).Activate()
                            End If
                        Case mnu_PMFSC
                            If oApplication.SBO_Application.MessageBox("Do You Want to Recreate Failed Service Call?", 2, "Yes", "No") = 2 Then
                                BubbleEvent = False
                                Exit Sub
                            Else
                                Dim DocEntry As String
                                Dim strQuery As String = String.Empty
                                Dim oRecordSet As SAPbobsCOM.Recordset
                                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                DocEntry = oDBDataSource.GetValue("DocEntry", oDBDataSource.Offset)
                                strQuery = "Select U_ItemCode,T1.U_CardCode,T0.LineId,U_TerDt,U_MSerialNo,U_SerialNo,U_InsID,U_DocDate,U_SConNo From [@Z_PMT1] T0 JOIN [@Z_OPMT] T1 On T0.DocEntry = T1.DocEntry "
                                strQuery += " Where T1.DocEntry = '" + DocEntry + "'"
                                strQuery += " And ((U_TerDt Is Null) OR (ISNULL(U_TerDt,'') = '')  AND (U_TerDt > U_DocDate)) "
                                strQuery += " And ISNULL(T0.U_SCallNo,'') = '' "
                                strQuery += " AND (T1.U_DocDate Between T0.U_SEStartDt And T0.U_SEEndDt) "
                                'strQuery += " AND (T1.U_PMEndDt Between T1.U_SEStartDt And T1.U_SEEndDt) "
                                oRecordSet.DoQuery(strQuery)
                                If oRecordSet.RecordCount <= 0 Then
                                    oApplication.Utilities.Message("No Failed Records Found...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                Else
                                    oApplication.Utilities.BookServiceCall(DocEntry, oForm)
                                    oApplication.SBO_Application.Menus.Item(mnu_ADD).Activate()
                                End If
                            End If
                    End Select
                Case False
                    Select Case pVal.MenuUID
                        Case mnu_Z_OPMT
                            LoadForm()
                        Case mnu_ADD
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            initialize(oForm)
                            clearSource(oForm)
                            'EnableControls(oForm)
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            If oForm.TypeEx = frm_Z_OPMT Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        If BusinessObjectInfo.ActionSuccess = True Then
                            Select Case BusinessObjectInfo.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                                    oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                                    Dim DocEntry As String = oXmlDoc.SelectSingleNode("/Preventive_MaintenanceParams/DocEntry").InnerText
                                    oApplication.Utilities.BookServiceCall(DocEntry, oForm)
                                    oApplication.Utilities.updateServiceContract(DocEntry)
                                Case (SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
                                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPMT")
                                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PMT1")

                                    If (oDBDataSource.GetValue("U_PMStatus", oDBDataSource.Offset) = "L" _
                                        Or oDBDataSource.GetValue("U_PMStatus", oDBDataSource.Offset) = "C") Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                    Else
                                        Dim strInvoice As String = CType(oForm.Items.Item("20").Specific, SAPbouiCOM.EditText).Value
                                        If strInvoice.Length > 0 Then
                                            oForm.Items.Item("34").Enabled = False
                                        Else
                                            oForm.Items.Item("34").Enabled = True
                                        End If
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    End If
                            End Select
                        End If
                End Select
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        If oForm.TypeEx = frm_Z_OPMT Then
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oMenus As SAPbouiCOM.Menus
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data
            oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

            If (eventInfo.BeforeAction = True) Then
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If Not oMenuItem.SubMenus.Exists(mnu_PMCancel) Then
                            oCreationPackage.Checked = False
                            oCreationPackage.Enabled = True
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.UniqueID = mnu_PMCancel
                            oCreationPackage.String = "PM Cancellation"
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If

                        If Not oMenuItem.SubMenus.Exists(mnu_PMFSC) Then
                            oCreationPackage.Checked = False
                            oCreationPackage.Enabled = True
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.UniqueID = mnu_PMFSC
                            oCreationPackage.String = "Recreate PM Failed Service Call"
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                If oMenuItem.SubMenus.Exists(mnu_PMCancel) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_PMCancel)
                End If
                If oMenuItem.SubMenus.Exists(mnu_PMFSC) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_PMFSC)
                End If
                If oMenuItem.SubMenus.Exists(mnu_CLOSE) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_CLOSE)
                End If
                If oMenuItem.SubMenus.Exists(mnu_CANCEL) Then
                    oMenuItem.SubMenus.RemoveEx(mnu_CANCEL)
                End If
            End If
        End If
    End Sub
#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPMT")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PMT1")

            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select IsNull(MAX(DocEntry)+1,1) From [@Z_OPMT]")
            If Not oRecordSet.EoF Then
                oDBDataSource.SetValue("DocNum", 0, oRecordSet.Fields.Item(0).Value.ToString())
            End If
            oDBDataSource.SetValue("U_DocDate", 0, System.DateTime.Now.ToString("yyyyMMdd"))

            oForm.Update()
            MatrixId = "3"
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub dataBind(ByVal oForm As SAPbouiCOM.Form)
        Try
            'oForm.DataSources.UserDataSources.Add("udsContID", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100)
            oForm.DataSources.UserDataSources.Add("udsContDes", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 250)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub clearSource(ByVal aForm As SAPbouiCOM.Form)
        Try
            'aForm.DataSources.UserDataSources.Item("udsContID").ValueEx = ""
            aForm.DataSources.UserDataSources.Item("udsContDes").ValueEx = ""
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim blnStatus As Boolean = False
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPMT")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PMT1")

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If oApplication.Utilities.getEditTextvalue(aForm, "13") = "" Then
                oApplication.Utilities.Message("Enter Service Contract No...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Dim blnRowAdded As Boolean = False
            If oMatrix.RowCount = 0 Then
                oApplication.Utilities.Message("No Row Details Found...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                For index As Integer = 1 To oMatrix.RowCount
                    If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index).ToString() <> "" Then
                        blnRowAdded = True
                        Exit For
                    End If
                Next
            End If

            If Not blnRowAdded Then
                oApplication.Utilities.Message("No Row Details Found...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception

        End Try
    End Function

    Private Function validation_PM(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPMT")
            oMatrix.LoadFromDataSource()

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select 1 As 'Return',T0.DocEntry From [@Z_OPMT] T0 "
            strQuery += " Where "
            strQuery += " U_SConNo = '" + oDBDataSource.GetValue("U_SConNo", 0).Trim() + "' And T0.DocEntry <> '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            strQuery += " And T0.U_PMStatus = 'O'"
            oRecordSet.DoQuery(strQuery)

            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("Preventive Maintenance Document Already exist for this Service Contract : " & oDBDataSource.GetValue("U_SConNo", 0).Trim() & "  with open status...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Function validation_AR(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPMT")
            oMatrix.LoadFromDataSource()

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select U_CStatus From [@Z_PMT1] T0 "
            strQuery += " Where "
            strQuery += " T0.DocEntry = '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            strQuery += " And U_CStatus <> '-1'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("Some of the Call Status is Not Closed...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select U_BillAmt From [@Z_PMT1] T0 "
            strQuery += " Where "
            strQuery += " T0.DocEntry = '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            strQuery += " And ISNULL(U_NoBill,'N') = 'N'"
            strQuery += " And ISNULL(U_BillAmt,0) < 0 "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("Bill Amount should be greater than zero...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select U_BillAmt From [@Z_PMT1] T0 "
            strQuery += " Where "
            strQuery += " T0.DocEntry = '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            strQuery += " And ISNULL(U_NoBill,'N') = 'N'"
            strQuery += " And ISNULL(U_BillAmt,0) = 0 "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("No Bill Should be selected, if the Bill Amount is Zero... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            'oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'strQuery = "Select U_BillAmt From [@Z_PMT1] T0 "
            'strQuery += " Where "
            'strQuery += " T0.DocEntry = '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            'strQuery += " And ISNULL(U_NoBill,'N') = 'Y'"
            'strQuery += " And ISNULL(U_BillAmt,0) <> 0 "
            'oRecordSet.DoQuery(strQuery)
            'If Not oRecordSet.EoF Then
            '    oApplication.Utilities.Message("Bill Amount Should be Zero When No Bill is Selected...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select U_BillAmt From [@Z_PMT1] T0 "
            strQuery += " Where "
            strQuery += " T0.DocEntry = '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            strQuery += " And ISNULL(U_NoBill,'N') = 'N'"
            strQuery += " And ISNULL(U_BillAmt,0) > 0 "
            oRecordSet.DoQuery(strQuery)
            If oRecordSet.EoF Then
                oApplication.Utilities.Message("No Record Exist to Create A/R Invoice...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub loadComboColumn(ByVal oForm As SAPbouiCOM.Form)
        Try
            oCombo = oMatrix.Columns.Item("V_10").Cells().Item(1).Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select statusID,Name  From OSCS "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("statusID").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PMT1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines.Size
                oDBDataSourceLines.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PMT1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    Else
                        oMatrix.AddRow(1, oMatrix.RowCount + 1)
                        oMatrix.ClearRowData(oMatrix.RowCount)
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines.Size
                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Try
            Select Case aForm.PaneLevel
                Case "0"
                    oMatrix = aForm.Items.Item("11").Specific
                    oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_PMT1")
            End Select
            oMatrix.FlushToDataSource()
            For introw As Integer = 1 To oMatrix.RowCount
                If oMatrix.IsRowSelected(introw) Then
                    oMatrix.DeleteRow(introw)
                    oDBDataSourceLines.RemoveRecord(introw - 1)
                    oMatrix.FlushToDataSource()
                    For count As Integer = 1 To oDBDataSourceLines.Size
                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
                    Next
                    Select Case aForm.PaneLevel
                        Case "0"
                            oMatrix = aForm.Items.Item("3").Specific
                            oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_PMT1")
                            AssignLineNo(aForm)
                    End Select
                    oMatrix.LoadFromDataSource()
                    If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPMT")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PMT1")

            If Me.MatrixId = "3" Then
                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PMT1")
            End If

            Me.RowtoDelete = intSelectedMatrixrow
            oDBDataSourceLines.RemoveRecord(Me.RowtoDelete - 1)
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines.Size - 1
                oDBDataSourceLines.SetValue("LineId", count - 1, count)
            Next

            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub calculateDates(ByVal oForm As SAPbouiCOM.Form, ByVal intFreqency As Integer, ByRef strPMStartDt As String, ByRef strPMEndDt As String)
        Try
            Dim strFromDate As String = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value
            Dim strToDate As String = CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value
            Dim strPMScheNo As String = CType(oForm.Items.Item("16").Specific, SAPbouiCOM.EditText).Value
            Dim dtSFD As DateTime = Convert.ToDateTime(strFromDate.Substring(0, 4) + "-" + strFromDate.Substring(4, 2) + "-" + strFromDate.Substring(6, 2))
            Dim dtSTD As DateTime = Convert.ToDateTime(strToDate.Substring(0, 4) + "-" + strToDate.Substring(4, 2) + "-" + strToDate.Substring(6, 2))

            Dim dtPMStartDate As DateTime
            Dim dtPMEndDate As DateTime
            If intFreqency = 0 Then
                intFreqency = 1
            End If
            Dim intNoOfSch As Integer = DateDiff(DateInterval.Month, dtSFD, dtSTD) / intFreqency
            dtPMStartDate = dtSFD.AddMonths(intFreqency * (CInt(strPMScheNo) - 1))
            Dim intTMonth As Integer = -((intFreqency * intNoOfSch) - (intFreqency * CInt(strPMScheNo)))
            dtPMEndDate = dtSTD.AddMonths(intTMonth)
            strPMStartDt = dtPMStartDate.ToString("yyyyMMdd")
            strPMEndDt = dtPMEndDate.ToString("yyyyMMdd")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            oForm.Items.Item("4").Width = oForm.Width - 30
            oForm.Items.Item("4").Height = oForm.Items.Item("3").Height + 10

            'oForm.Items.Item("41").Top = oForm.Items.Item("27").Top + oForm.Items.Item("27").Height + 5 'Resource Consolidation
            'oForm.Items.Item("51").Top = oForm.Items.Item("41").Top + oForm.Items.Item("41").Height + 5 'Resource Grid
            'oForm.Items.Item("51").Height = (oForm.Items.Item("29").Height / 2) - 30
            'oForm.Items.Item("51").Width = oForm.Items.Item("3").Width

            'oForm.Items.Item("43").Top = oForm.Items.Item("51").Top + oForm.Items.Item("51").Height + 5 'Documents
            'oForm.Items.Item("52").Top = oForm.Items.Item("43").Top + oForm.Items.Item("43").Height + 5 'Resource Grid
            'oForm.Items.Item("52").Height = (oForm.Items.Item("29").Height / 2) - 30
            'oForm.Items.Item("52").Width = oForm.Items.Item("3").Width

            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Function changeStatus(ByVal oForm As SAPbouiCOM.Form, ByVal strStatus As String) As Boolean
        Dim _retVal As Boolean = False
        Try
            Dim strDocEntry As String = CType(oForm.Items.Item("25").Specific, SAPbouiCOM.EditText).Value
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim oGeneralDataParams As SAPbobsCOM.GeneralDataParams
            Dim strQuery As String = String.Empty
            oCompanyService = oApplication.Company.GetCompanyService()
            Try
                oGeneralService = oCompanyService.GetGeneralService("Z_OPMT")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralDataParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralDataParams.SetProperty("DocEntry", strDocEntry)
                oGeneralData = oGeneralService.GetByParams(oGeneralDataParams)
                oGeneralData.SetProperty("U_PMStatus", strStatus)
                Dim intPMSchedule As Integer = CInt(oGeneralData.GetProperty("U_PMSchNo"))
                oGeneralData.SetProperty("U_PMSchNo", (intPMSchedule - 1).ToString())
                oGeneralService.Update(oGeneralData)
                _retVal = True
            Catch ex As Exception
                Throw ex
            End Try
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists

            oCFL = oCFLs.Item("CFL_1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "A"
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub filterEventLabelChooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal strCFLID As String)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            oCFLs = oForm.ChooseFromLists
            oCFL = oCFLs.Item(strCFLID)
            oCons = oCFL.GetConditions()
            If oCons.Count = 0 Then
                oCon = oCons.Add()
            Else
                oCon = oCons.Item(0)
            End If
            oCon.Alias = "U_EType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = oForm.Items.Item("6").Specific.value
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class

