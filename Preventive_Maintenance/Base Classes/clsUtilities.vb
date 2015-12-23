Imports System.Xml
Imports System.Collections.Specialized
Imports System.IO
Imports SAPbobsCOM


Public Class clsUtilities

    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer
    Public strSFilePath As String = String.Empty
    Public strDFilePath As String = String.Empty
    Private strFilepath As String = String.Empty
    Private strFileName As String = String.Empty

    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub

#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect() <> 0 Then
                Throw New Exception("Cannot connect to company database. ")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Genral Functions"

    Public Sub assignLineNo(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
                aGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
            aform.Freeze(False)
        Catch ex As Exception

        End Try
    End Sub

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            Throw ex
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

#End Region

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            Throw ex
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            Throw ex
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()
        Catch ex As Exception
            Throw ex
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

    Public Sub setEditText(ByVal aForm As SAPbouiCOM.Form, ByVal aItem As String, ByVal aValue As String)
        Dim oEdit As SAPbouiCOM.EditText
        oEdit = aForm.Items.Item(aItem).Specific
        oEdit.String = aValue
    End Sub

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function getLocalCurrency(ByVal strCurrency As String) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select Maincurrncy from OADM")
        Return oTemp.Fields.Item(0).Value
    End Function

#Region "Get ExchangeRate"
    Public Function getExchangeRate(ByVal strCurrency As String) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dblExchange As Double
        If GetCurrency("Local") = strCurrency Then
            dblExchange = 1
        Else
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery("Select isNull(Rate,0) from ORTT where convert(nvarchar(10),RateDate,101)=Convert(nvarchar(10),getdate(),101) and currency='" & strCurrency & "'")
            dblExchange = oTemp.Fields.Item(0).Value
        End If
        Return dblExchange
    End Function

    Public Function getExchangeRate(ByVal strCurrency As String, ByVal dtdate As Date) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSql As String
        Dim dblExchange As Double
        If GetCurrency("Local") = strCurrency Then
            dblExchange = 1
        Else
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSql = "Select isNull(Rate,0) from ORTT where ratedate='" & dtdate.ToString("yyyy-MM-dd") & "' and currency='" & strCurrency & "'"
            oTemp.DoQuery(strSql)
            dblExchange = oTemp.Fields.Item(0).Value
        End If
        Return dblExchange
    End Function
#End Region

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

#Region "Get DocCurrency"
    Public Function GetDocCurrency(ByVal aDocEntry As Integer) As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select DocCur from OINV where docentry=" & aDocEntry)
        Return oTemp.Fields.Item(0).Value
    End Function
#End Region

#Region "GetEditTextValues"
    Public Function getEditTextvalue(ByVal aForm As SAPbouiCOM.Form, ByVal strUID As String) As String
        Dim oEditText As SAPbouiCOM.EditText
        oEditText = aForm.Items.Item(strUID).Specific
        Return oEditText.Value
    End Function
#End Region

#Region "Get Currency"
    Public Function GetCurrency(ByVal strChoice As String, Optional ByVal aCardCode As String = "") As String
        Dim strCurrQuery, Currency As String
        Dim oTempCurrency As SAPbobsCOM.Recordset
        If strChoice = "Local" Then
            strCurrQuery = "Select MainCurncy from OADM"
        Else
            strCurrQuery = "Select Currency from OCRD where CardCode='" & aCardCode & "'"
        End If
        oTempCurrency = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempCurrency.DoQuery(strCurrQuery)
        Currency = oTempCurrency.Fields.Item(0).Value
        Return Currency
    End Function

#End Region

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

#Region "AddControls"
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal dblWidth As Double = 0, Optional ByVal dblTop As Double = 0, Optional ByVal Hight As Double = 0, Optional ByVal Enable As Boolean = True)
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        Dim oEditText As SAPbouiCOM.EditText
        Dim ofolder As SAPbouiCOM.Folder
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 5
                    .Top = objOldItem.Top
                ElseIf position.ToUpper = "DOWN" Then
                    If ItemUID = "edWork" Then
                        .Left = objOldItem.Left + 40
                    Else
                        .Left = objOldItem.Left
                    End If
                    .Top = objOldItem.Top + objOldItem.Height + 3

                    .Width = objOldItem.Width
                    .Height = objOldItem.Height
                ElseIf position.ToUpper = "COPY" Then
                    .Top = objOldItem.Top
                    .Left = objOldItem.Left
                    .Height = objOldItem.Height
                    .Width = objOldItem.Width
                End If
            End If
            .FromPane = fromPane
            .ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            .LinkTo = linkedUID
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            objNewItem.Width = objOldItem.Width '+ 50
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
            ofolder = objNewItem.Specific
            ofolder.Caption = strCaption
            ofolder.GroupWith(linkedUID)
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption

        End If
        If dblWidth <> 0 Then
            objNewItem.Width = dblWidth
        End If

        If dblTop <> 0 Then
            objNewItem.Top = objNewItem.Top + dblTop
        End If
        If Hight <> 0 Then
            objNewItem.Height = objNewItem.Height + Hight
        End If
    End Sub
#End Region

#Region "Set / Get Values from Matrix"
    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value = strvalue
    End Sub
#End Region

#Region "Add Condition CFL"
    Public Sub AddConditionCFL(ByVal FormUID As String, ByVal strQuery As String, ByVal strQueryField As String, ByVal sCFL As String)
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim Conditions As SAPbouiCOM.Conditions
        Dim oCond As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Dim sDocEntry As New ArrayList()
        Dim sDocNum As ArrayList
        Dim MatrixItem As ArrayList
        sDocEntry = New ArrayList()
        sDocNum = New ArrayList()
        MatrixItem = New ArrayList()

        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            oCFL = oCFLs.Item(sCFL)

            Dim oRec As SAPbobsCOM.Recordset
            oRec = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRec.DoQuery(strQuery)
            oRec.MoveFirst()

            Try
                If oRec.EoF Then
                    sDocEntry.Add("")
                Else
                    While Not oRec.EoF
                        Dim DocNum As String = oRec.Fields.Item(strQueryField).Value.ToString()
                        If DocNum <> "" Then
                            sDocEntry.Add(DocNum)
                        End If
                        oRec.MoveNext()
                    End While
                End If
            Catch generatedExceptionName As Exception
                Throw
            End Try

            'If IsMatrixCondition = True Then
            '    Dim oMatrix As SAPbouiCOM.Matrix
            '    oMatrix = DirectCast(oForm.Items.Item(Matrixname).Specific, SAPbouiCOM.Matrix)

            '    For a As Integer = 1 To oMatrix.RowCount
            '        If a <> pVal.Row Then
            '            MatrixItem.Add(DirectCast(oMatrix.Columns.Item(columnname).Cells.Item(a).Specific, SAPbouiCOM.EditText).Value)
            '        End If
            '    Next
            '    If removelist = True Then
            '        For xx As Integer = 0 To MatrixItem.Count - 1
            '            Dim zz As String = MatrixItem(xx).ToString()
            '            If sDocEntry.Contains(zz) Then
            '                sDocEntry.Remove(zz)
            '            End If
            '        Next
            '    End If
            'End If

            'oCFLs = oForm.ChooseFromLists
            'oCFLCreationParams = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            'If systemMatrix = True Then
            '    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = Nothing
            '    oCFLEvento = DirectCast(pVal, SAPbouiCOM.IChooseFromListEvent)
            '    Dim sCFL_ID As String = Nothing
            '    sCFL_ID = oCFLEvento.ChooseFromListUID
            '    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
            'Else
            '    oCFL = oForm.ChooseFromLists.Item(sCHUD)
            'End If

            Conditions = New SAPbouiCOM.Conditions()
            oCFL.SetConditions(Conditions)
            Conditions = oCFL.GetConditions()
            oCond = Conditions.Add()
            oCond.BracketOpenNum = 2
            For i As Integer = 0 To sDocEntry.Count - 1
                If i > 0 Then
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCond = Conditions.Add()
                    oCond.BracketOpenNum = 1
                End If

                oCond.[Alias] = strQueryField
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = sDocEntry(i).ToString()
                If i + 1 = sDocEntry.Count Then
                    oCond.BracketCloseNum = 2
                Else
                    oCond.BracketCloseNum = 1
                End If
            Next

            oCFL.SetConditions(Conditions)


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Open Files"
    Public Sub OpenFile(ByVal strPath As String)
        Try
            If File.Exists(strPath) Then
                Dim process As New System.Diagnostics.Process
                Dim filestart As New System.Diagnostics.ProcessStartInfo(strPath)
                filestart.UseShellExecute = True
                filestart.WindowStyle = ProcessWindowStyle.Normal
                process.StartInfo = filestart
                process.Start()
            End If
        Catch ex As Exception

        End Try
    End Sub
#End Region

    Public Function createEMMainAuthorization() As Boolean
        Try
            Dim RetVal As Long
            Dim mUserPermission As SAPbobsCOM.UserPermissionTree
            mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
            '//Mandatory field, which is the key of the object.
            '//The partner namespace must be included as a prefix followed by _
            mUserPermission.PermissionID = "Pre_Main"
            '//The Name value that will be displayed in the General Authorization Tree
            mUserPermission.Name = "Preventive Maintenance Addon"
            '//The permission that this object can get
            mUserPermission.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
            '//In case the level is one, there Is no need to set the FatherID parameter.
            '   mUserPermission.Levels = 1
            RetVal = mUserPermission.Add
            If RetVal = 0 Or -2035 Then
                Return True
            Else
                MsgBox(oApplication.Company.GetLastErrorDescription)
                Return False
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function addChildAuthorization(ByVal aChildID As String, ByVal aChildiDName As String, ByVal aorder As Integer, ByVal aFormType As String, ByVal aParentID As String, ByVal Permission As SAPbobsCOM.BoUPTOptions) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim mUserPermission As SAPbobsCOM.UserPermissionTree
        mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)

        mUserPermission.PermissionID = aChildID
        mUserPermission.Name = aChildiDName
        mUserPermission.Options = Permission ' SAPbobsCOM.BoUPTOptions.bou_FullReadNone

        '//For level 2 and up you must set the object's father unique ID
        'mUserPermission.Level
        mUserPermission.ParentID = aParentID
        mUserPermission.UserPermissionForms.DisplayOrder = aorder
        '//this object manages forms
        ' If aFormType <> "" Then
        mUserPermission.UserPermissionForms.FormType = aFormType
        ' End If

        RetVal = mUserPermission.Add
        If RetVal = 0 Or RetVal = -2035 Then
            Return True
        Else
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End If


    End Function

    Public Sub AuthorizationCreation()
        addChildAuthorization("PM_Trans", "Transactions", 2, "", "Pre_Main", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        'Transaction
        addChildAuthorization("Z_OPTM", "Preventive Maintanence", 3, "frm_Z_OPTM", "PM_Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
    End Sub

    Public Function validateAuthorization(ByVal aUserId As String, ByVal aFormUID As String) As Boolean
        Dim oAuth As SAPbobsCOM.Recordset
        oAuth = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim struserid As String
        '    Return False
        struserid = oApplication.Company.UserName
        oAuth.DoQuery("select * from UPT1 where FormId='" & aFormUID & "'")
        If (oAuth.RecordCount <= 0) Then
            Return True
        Else
            Dim st As String
            st = oAuth.Fields.Item("PermId").Value
            st = "Select * from USR3 where PermId='" & st & "' and UserLink=" & aUserId
            oAuth.DoQuery(st)
            If oAuth.RecordCount > 0 Then
                If oAuth.Fields.Item("Permission").Value = "N" Then
                    Return False
                End If
                Return True
            Else
                Return True
            End If

        End If

        Return True

    End Function

    Public Sub AssignSerialNo(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intRow As Integer = 1 To aMatrix.RowCount
            aMatrix.Columns.Item("SlNo").Cells.Item(intRow).Specific.value = intRow
        Next
        aform.Freeze(False)
    End Sub

    Public Sub AssignRowNo(ByVal aMatrix As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intRow As Integer = 0 To aMatrix.DataTable.Rows.Count - 1
            aMatrix.RowHeaders.SetText(intRow, intRow + 1)
        Next
        aform.Freeze(False)
    End Sub

#Region "ValidateCode"
    Public Function ValidateCode(ByVal aCode As String, ByVal aModule As String) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strqry As String = ""
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aModule = "Z_OEVT" Then
            strqry = "Select * from ""@Z_OEVT"" where ""U_Code""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Event Type Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "Z_OEVL" Then
            strqry = "Select * from ""@Z_OEVL"" where ""U_Code""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Event Level Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "Z_OFUS" Then
            strqry = "Select * from ""@Z_OFUS"" where ""U_Code""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Function Space Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "Z_OMUT" Then
            strqry = "Select * from ""@Z_OMUT"" where ""U_Code""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Menu Type Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        End If
        Return False
    End Function
#End Region

    Public Function AddEventResources(ByVal oForm As SAPbouiCOM.Form, ByVal strMenu As String) As String
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGenDataChild As SAPbobsCOM.GeneralData
        Dim oGenDataCollection As SAPbobsCOM.GeneralDataCollection
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim strCode As String = String.Empty
        oCompanyService = oApplication.Company.GetCompanyService()
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OEBR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGenDataChild = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim intCode As Integer = getMaxCode("@Z_OEBR", "DocEntry")
            strCode = String.Format("{0:000000000}", intCode)
            oGeneralData.SetProperty("U_Reference", strCode)
            oGeneralData.SetProperty("U_MCode", strMenu)

            oGenDataCollection = oGeneralData.Child("Z_EBR1")

            strQuery = "Select U_ItemCode,U_ItemName,U_Quantity,U_Price,U_RevType,U_Remarks From [@Z_OMET] T0 JOIN [@Z_MET1] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T0.U_Code = '" + strMenu + "' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oGenDataChild = oGenDataCollection.Add()
                    oGenDataChild.SetProperty("U_ItemCode", oRecordSet.Fields.Item("U_ItemCode").Value.ToString)
                    oGenDataChild.SetProperty("U_ItemName", oRecordSet.Fields.Item("U_ItemName").Value.ToString)
                    oGenDataChild.SetProperty("U_Quantity", oRecordSet.Fields.Item("U_Quantity").Value.ToString)
                    oGenDataChild.SetProperty("U_Price", oRecordSet.Fields.Item("U_Price").Value.ToString)
                    oGenDataChild.SetProperty("U_RevType", oRecordSet.Fields.Item("U_RevType").Value.ToString)
                    oGenDataChild.SetProperty("U_Remarks", oRecordSet.Fields.Item("U_Remarks").Value.ToString)
                    oRecordSet.MoveNext()
                End While
            End If

            oGeneralService.Add(oGeneralData)
            Return strCode
        Catch ex As Exception
            Throw ex
        End Try
        Return strCode
    End Function

    Public Sub RemoveEventResources(ByVal oForm As SAPbouiCOM.Form, ByVal strReference As String)
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oGeneralDataParams As SAPbobsCOM.GeneralDataParams
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim strDocEntry As String = String.Empty
        oCompanyService = oApplication.Company.GetCompanyService()
        Try
            oGeneralService = oCompanyService.GetGeneralService("Z_OEBR")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralDataParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select DocEntry From [@Z_OEBR] Where U_Reference = '" + strReference + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strDocEntry = oRecordSet.Fields.Item(0).Value
                oGeneralDataParams.SetProperty("DocEntry", strDocEntry)
                oGeneralService.Delete(oGeneralDataParams)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function AddQuotation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = False
        Dim oQuotation As SAPbobsCOM.Documents
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim intStatus As Integer
        Try
            oQuotation = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strDocEntry As String = CType(oForm.Items.Item("16").Specific, SAPbouiCOM.EditText).Value

            oQuotation.CardCode = oForm.Items.Item("24").Specific.value
            oQuotation.CardName = oForm.Items.Item("_24").Specific.value
            oQuotation.NumAtCard = oForm.Items.Item("13").Specific.value
            oQuotation.DocDate = System.DateTime.Now
            oQuotation.TaxDate = System.DateTime.Now
            oQuotation.DocDueDate = System.DateTime.Now
            oQuotation.SalesPersonCode = CType(oForm.Items.Item("_12").Specific, SAPbouiCOM.ComboBox).Selected.Value
            oQuotation.DiscountPercent = CDbl(oForm.Items.Item("48").Specific.value)
            oQuotation.Comments = "Event Booking"
            oQuotation.UserFields.Fields.Item("U_BookNo").Value = strDocEntry

            strQuery = " Select T0.U_ItemCode,T0.U_ItemName,T0.U_Quantity,T0.U_Price,T0.U_LineTotal,T0.U_RevType,T0.U_Remarks "
            strQuery += " From [@Z_EBR1] T0 JOIN [@Z_OEBR] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " JOIN [@Z_EBQ1] T2 On T2.U_Reference = T1.U_Reference  "
            strQuery += " Where T2.DocEntry = '" + strDocEntry + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    oQuotation.Lines.ItemCode = oRecordSet.Fields.Item("U_ItemCode").Value
                    oQuotation.Lines.Quantity = oRecordSet.Fields.Item("U_Quantity").Value
                    oQuotation.Lines.UnitPrice = oRecordSet.Fields.Item("U_Price").Value
                    'oQuotation.Lines.TaxCode = "CST@4"
                    'oQuotation.Lines.LineTotal = oRecordSet.Fields.Item("U_LineTotal").Value
                    oQuotation.Lines.CostingCode = oRecordSet.Fields.Item("U_RevType").Value
                    oQuotation.Lines.Add()
                    oRecordSet.MoveNext()
                End While
                intStatus = oQuotation.Add
                If intStatus = 0 Then
                    Dim strQutotation As String = oApplication.Company.GetNewObjectKey()
                    _retVal = True
                    strQuery = "Update [@Z_OEBQ] Set U_SalesQ = '" + strQutotation + "' Where DocEntry = '" + strDocEntry + "'"
                    oRecordSet.DoQuery(strQuery)
                Else
                    _retVal = False
                    Throw New Exception(oApplication.Company.GetLastErrorDescription())
                End If
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function AddOrder(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = False
        Try
            Dim oOrder As SAPbobsCOM.Documents
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim strQuery As String = String.Empty
            Dim intStatus As Integer
            Try
                oOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strDocEntry As String = CType(oForm.Items.Item("16").Specific, SAPbouiCOM.EditText).Value
                Dim strQuotation As String = CType(oForm.Items.Item("69").Specific, SAPbouiCOM.EditText).Value

                oOrder.CardCode = oForm.Items.Item("24").Specific.value
                oOrder.CardName = oForm.Items.Item("_24").Specific.value
                oOrder.NumAtCard = oForm.Items.Item("13").Specific.value
                oOrder.DocDate = System.DateTime.Now
                oOrder.TaxDate = System.DateTime.Now
                oOrder.DocDueDate = System.DateTime.Now
                oOrder.SalesPersonCode = CType(oForm.Items.Item("_12").Specific, SAPbouiCOM.ComboBox).Selected.Value
                oOrder.DiscountPercent = CDbl(oForm.Items.Item("48").Specific.value)
                oOrder.Comments = "Event Booking"
                oOrder.UserFields.Fields.Item("U_BookNo").Value = strDocEntry

                strQuery = " Select ItemCode,Dscription,Quantity,Price,WhsCode,LineTotal,DocEntry,LineNum,TaxCode,OcrCode From QUT1 Where DocEntry = '" + strQuotation + "' "
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    While Not oRecordSet.EoF
                        oOrder.Lines.ItemCode = oRecordSet.Fields.Item("ItemCode").Value
                        oOrder.Lines.Quantity = oRecordSet.Fields.Item("Quantity").Value
                        oOrder.Lines.UnitPrice = oRecordSet.Fields.Item("Price").Value
                        oOrder.Lines.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value
                        oOrder.Lines.TaxCode = oRecordSet.Fields.Item("TaxCode").Value
                        'oOrder.Lines.LineTotal = oRecordSet.Fields.Item("LineTotal").Value
                        oOrder.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oQuotations
                        oOrder.Lines.BaseEntry = oRecordSet.Fields.Item("DocEntry").Value
                        oOrder.Lines.BaseLine = oRecordSet.Fields.Item("LineNum").Value
                        oOrder.Lines.CostingCode = oRecordSet.Fields.Item("OcrCode").Value
                        oOrder.Lines.Add()
                        oRecordSet.MoveNext()
                    End While
                    intStatus = oOrder.Add
                    If intStatus = 0 Then
                        _retVal = True
                        Dim strOrder As String = oApplication.Company.GetNewObjectKey()
                        strQuery = "Update [@Z_OEBQ] Set U_SalesO = '" + strOrder + "' Where DocEntry = '" + strDocEntry + "'"
                        oRecordSet.DoQuery(strQuery)
                    Else
                        _retVal = False
                        Throw New Exception(oApplication.Company.GetLastErrorDescription())
                    End If
                End If
                Return _retVal
            Catch ex As Exception
                Throw ex
            End Try
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function BookServiceCall(ByVal strPMRef As String, ByVal oForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oCERecordSet As SAPbobsCOM.Recordset
        Dim oPTRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim strOrigin As String = String.Empty
        Dim strProblem As String = String.Empty
        Dim strCallType As String = String.Empty
        Try
            Dim oServiceCall As SAPbobsCOM.ServiceCalls
            oServiceCall = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
            Dim oUpdateRecord As SAPbobsCOM.Recordset
            oUpdateRecord = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCERecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oPTRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            strQuery = "Select originID From OSCO Where U_ConPM = 'Y'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strOrigin = oRecordSet.Fields.Item("originID").Value
            End If

            strQuery = "Select prblmTypID From OSCP Where U_ConPM = 'Y'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strProblem = oRecordSet.Fields.Item("prblmTypID").Value
            End If

            strQuery = "Select calltypeID From OSCT Where U_ConPM = 'Y'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strCallType = oRecordSet.Fields.Item("calltypeID").Value
            End If

            Dim strFile As String = "\Service_Creation_" + System.DateTime.Now.ToString("yyyyMMddmmss") + ".txt"

            strQuery = "Select U_ItemCode,T1.U_CardCode,T0.LineId,U_TerDt,U_MSerialNo,U_SerialNo,U_InsID,U_DocDate,U_SConNo From [@Z_PMT1] T0 JOIN [@Z_OPMT] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T1.DocEntry = '" + strPMRef + "'"
            strQuery += " And ((U_TerDt Is Null) OR (ISNULL(U_TerDt,'') = '') AND (U_TerDt > U_DocDate)) "
            strQuery += " And ISNULL(T0.U_SCallNo,'') = '' "
            strQuery += " AND (T1.U_DocDate Between T0.U_SEStartDt And T0.U_SEEndDt) "
            'strQuery += " AND (T1.U_PMEndDt Between T1.U_SEStartDt And T1.U_SEEndDt) "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim dtDocDate As DateTime = oRecordSet.Fields.Item("U_DocDate").Value
                While Not oRecordSet.EoF
                    If 1 = 1 Then

                        oServiceCall.CustomerCode = oRecordSet.Fields.Item("U_CardCode").Value
                        oServiceCall.ItemCode = oRecordSet.Fields.Item("U_ItemCode").Value

                        If oForm.DataSources.UserDataSources.Item("udsContDes").ValueEx.Trim().Length > 0 Then
                            oServiceCall.Subject = oForm.DataSources.UserDataSources.Item("udsContDes").ValueEx.Trim()
                        Else
                            oServiceCall.Subject = " PM Ref : " + strPMRef + ""
                        End If

                        oServiceCall.Priority = SAPbobsCOM.BoSvcCallPriorities.scp_Medium
                        oServiceCall.Status = -3

                        oServiceCall.ManufacturerSerialNum = oRecordSet.Fields.Item("U_MSerialNo").Value
                        oServiceCall.InternalSerialNum = oRecordSet.Fields.Item("U_SerialNo").Value
                        oServiceCall.ContractID = CInt(oRecordSet.Fields.Item("U_SConNo").Value)

                        If strOrigin.Trim().Length > 0 Then
                            oServiceCall.Origin = strOrigin
                        End If

                        If strProblem.Trim().Length > 0 Then
                            oServiceCall.ProblemType = strProblem
                        End If

                        If strCallType.Trim().Length > 0 Then
                            oServiceCall.CallType = strCallType
                        End If

                        oServiceCall.UserFields.Fields.Item("U_PMRef").Value = strPMRef

                        strQuery = " Select ContactCod,technician,U_Queue,U_Remarks From OINS Where insID = '" + oRecordSet.Fields.Item("U_InsID").Value + "'"
                        oCERecordSet.DoQuery(strQuery)
                        If Not oCERecordSet.EoF Then
                            If oCERecordSet.Fields.Item("ContactCod").Value.ToString().Length > 0 Then
                                If oCERecordSet.Fields.Item("ContactCod").Value <> "0" Then
                                    oServiceCall.ContactCode = oCERecordSet.Fields.Item("ContactCod").Value
                                Else
                                    oServiceCall.ContactCode = -1
                                End If
                            End If
                            If oCERecordSet.Fields.Item("technician").Value.ToString().Length > 0 Then
                                If oCERecordSet.Fields.Item("technician").Value <> "0" Then
                                    oServiceCall.TechnicianCode = oCERecordSet.Fields.Item("technician").Value
                                End If
                            End If
                            If oCERecordSet.Fields.Item("U_Queue").Value.ToString().Length > 0 Then
                                oServiceCall.BelongsToAQueue = BoYesNoEnum.tYES
                                oServiceCall.Queue = oCERecordSet.Fields.Item("U_Queue").Value.ToString()
                            End If

                            strQuery = " Select [Text] From OPDT Where TextCode = '" + oCERecordSet.Fields.Item("U_Remarks").Value.ToString() + "'"
                            oPTRecordSet.DoQuery(strQuery)
                            If Not oPTRecordSet.EoF Then
                                oServiceCall.Description = oPTRecordSet.Fields.Item("Text").Value.ToString()
                            End If

                        End If

                        Dim intStatus As Integer = oServiceCall.Add()

                        If intStatus = 0 Then
                            Dim strSerNo As String = oApplication.Company.GetNewObjectKey()
                            strQuery = "Update [@Z_PMT1] Set U_SCallNo = '" + strSerNo + "'"
                            strQuery += " ,U_CStatus = -3 "
                            strQuery += " ,U_CDate = '" + System.DateTime.Now.ToString("MM-dd-yyyy") + "'"
                            strQuery += " Where DocEntry = '" + strPMRef + "'"
                            strQuery += " And LineId = '" + oRecordSet.Fields.Item("LineId").Value.ToString() + "'"
                            oUpdateRecord.DoQuery(strQuery)
                            Trace_ServiceCall("Service Item Code : " + oRecordSet.Fields.Item("U_ItemCode").Value.ToString() + " -->Success", strFile)
                        Else
                            'MessageBox.Show(oApplication.Company.GetLastErrorDescription())
                            Trace_ServiceCall("Service Item Code : " + oRecordSet.Fields.Item("U_ItemCode").Value.ToString() + "-->ERROR ERRORCODE :" + oApplication.Company.GetLastErrorCode().ToString() + " ERRORDESC : " + oApplication.Company.GetLastErrorDescription().ToString(), strFile)
                        End If
                    End If
                    oRecordSet.MoveNext()
                End While

                Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() + strFile
                If (File.Exists(strPath)) Then
                    System.Diagnostics.Process.Start(strPath)
                End If

            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function CancelServiceCall(ByVal oForm As SAPbouiCOM.Form, ByVal strPMRef As String)
        Dim _retVal As Boolean = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty

        Try
            Dim oServiceCall As SAPbobsCOM.ServiceCalls
            oServiceCall = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            strQuery = " Select U_SCallNo From [@Z_PMT1] T0 JOIN [@Z_OPMT] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T1.DocEntry = '" + strPMRef + "'"
            strQuery += " And U_SCallNo <> '' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                While Not oRecordSet.EoF
                    If oServiceCall.GetByKey(oRecordSet.Fields.Item(0).Value.ToString()) Then
                        oServiceCall.Resolution = "PM CANCELED"
                        oServiceCall.Status = 1
                        Dim intStatus As Integer = oServiceCall.Update()
                    End If
                    oRecordSet.MoveNext()
                End While
            End If

            strQuery = "Update [@Z_PMT1] Set "
            strQuery += " U_CStatus = '1'"
            strQuery += " Where DocEntry = '" + strPMRef + "'"
            oRecordSet.DoQuery(strQuery)

            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function AddARInvoice(ByVal oForm As SAPbouiCOM.Form, ByVal dtBDueDate As DateTime) As Boolean
        Dim _retVal As Boolean = False
        Dim oInvoice As SAPbobsCOM.Documents
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Dim intStatus As Integer
        Dim strOcrCode, strocrCode2, strCOGSOCRCODE, strCOGSOCRCODE2 As String
        Try
            oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select * from [@Z_OOCR]")
            If oRecordSet.RecordCount > 0 Then
                strOcrCode = oRecordSet.Fields.Item("U_OcrCode").Value
                strocrCode2 = oRecordSet.Fields.Item("U_OcrCode2").Value
                strCOGSOCRCODE = oRecordSet.Fields.Item("U_CogsOcrCode").Value
                strCOGSOCRCODE2 = oRecordSet.Fields.Item("U_CogsOcrCode2").Value
            Else
                strOcrCode = ""
                strocrCode2 = ""
                strCOGSOCRCODE = ""
                strCOGSOCRCODE2 = ""
            End If

            Dim strContract As String = CType(oForm.Items.Item("13").Specific, SAPbouiCOM.EditText).Value
            Dim strDescription As String = String.Empty
            strQuery = "Select Descriptio From OCTR Where ContractID = '" + strContract + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strDescription = oRecordSet.Fields.Item("Descriptio").Value.ToString()
            End If

            Dim strDocEntry As String = CType(oForm.Items.Item("25").Specific, SAPbouiCOM.EditText).Value

            oInvoice.CardCode = oForm.Items.Item("7").Specific.value
            oInvoice.CardName = oForm.Items.Item("_7").Specific.value
            oInvoice.NumAtCard = oForm.Items.Item("24").Specific.value
            oInvoice.DocDate = dtBDueDate.ToString("yyyy-MM-dd")
            oInvoice.TaxDate = dtBDueDate.ToString("yyyy-MM-dd")
            oInvoice.DocDueDate = dtBDueDate.ToString("yyyy-MM-dd")
            oInvoice.Comments = strDescription

            oInvoice.UserFields.Fields.Item("U_PMRef").Value = strDocEntry
            Try
                oInvoice.UserFields.Fields.Item("U_OrderType").Value = "PMContract"
            Catch ex As Exception

            End Try
            Try
                oInvoice.UserFields.Fields.Item("U_Ordertype").Value = "PMContract"
            Catch ex As Exception

            End Try
            oInvoice.UserFields.Fields.Item("U_QuType").Value = "Service"
            oInvoice.UserFields.Fields.Item("U_Contract").Value = strDescription

            oApplication.Utilities.Message("Header Values Set....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            strQuery = " Select ISNULL(T0.U_BillAmt,0) As U_BillAmt "
            strQuery += " From [@Z_PMT1] T0 JOIN [@Z_OPMT] T1 On T0.DocEntry = T1.DocEntry "
            strQuery += " Where T1.DocEntry = '" + strDocEntry.Trim() + "'"
            strQuery += " And U_BillAmt > 0 "
            strQuery += " And ISNULL(U_NoBill,'N') = 'N' "
            strQuery += " AND  ISNULL(T0.U_SCallNo,'') <> ''  "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("Inside Record Set in AddOrder ....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                While Not oRecordSet.EoF
                    oApplication.Utilities.Message("Loops through Recordset Lines ....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oInvoice.Lines.ItemCode = "IS502938"
                    oInvoice.Lines.LineTotal = CDbl(oRecordSet.Fields.Item("U_BillAmt").Value)
                    If strOcrCode <> "" Then
                        oInvoice.Lines.CostingCode = strOcrCode
                    End If
                    If strocrCode2 <> "" Then
                        oInvoice.Lines.CostingCode2 = strocrCode2
                    End If
                    If strCOGSOCRCODE <> "" Then
                        oInvoice.Lines.COGSCostingCode = strCOGSOCRCODE
                    End If
                    If strCOGSOCRCODE2 <> "" Then
                        oInvoice.Lines.COGSCostingCode2 = strCOGSOCRCODE2
                    End If
                    oInvoice.Lines.Add()
                    oRecordSet.MoveNext()
                End While

                intStatus = oInvoice.Add
                oApplication.Utilities.Message("After Order Add ....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If intStatus = 0 Then
                    oApplication.Utilities.Message("Order Generated Successfully ....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Dim strInvoice As String = oApplication.Company.GetNewObjectKey()
                    _retVal = True
                    If oInvoice.GetByKey(strInvoice) Then
                        changeStatus(oForm, "C", strInvoice, oInvoice.DocNum, oInvoice.DocDate)
                    End If
                Else
                    oApplication.Utilities.Message("Order Generated Failed....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription() + "-" + oApplication.Company.GetLastErrorCode().ToString(), SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    _retVal = False
                    Throw New Exception(oApplication.Company.GetLastErrorDescription())
                End If
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function changeStatus(ByVal oForm As SAPbouiCOM.Form, ByVal strStatus As String, ByVal strInvDE As String, ByVal strInvDN As String, ByVal dtInvDate As DateTime) As Boolean
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
                oGeneralData.SetProperty("U_InvDE", strInvDE)
                oGeneralData.SetProperty("U_InvDN", strInvDN)
                oGeneralData.SetProperty("U_InvDt", dtInvDate)
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

    Public Function updateRowPM(ByVal strSerNo As String)
        Dim _retVal As Boolean = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String
        Try
            Dim oServiceCall As SAPbobsCOM.ServiceCalls
            oServiceCall = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
            Dim oUpdateRecord As SAPbobsCOM.Recordset
            oUpdateRecord = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oServiceCall.GetByKey(strSerNo) Then
                Dim strPMRef As String = oServiceCall.UserFields.Fields.Item("U_PMRef").Value
                strQuery = "Select U_CStatus,T0.U_Technician,LineId From [@Z_PMT1] T0 "
                strQuery += " Where T0.DocEntry = '" + strPMRef + "'"
                strQuery += " And T0.U_SCallNo = '" + strSerNo + "'"
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    Dim strEmpName As String = getEmployeeName(oServiceCall.TechnicianCode.ToString())
                    If (oServiceCall.Status.ToString() <> oRecordSet.Fields.Item("U_CStatus").Value.ToString() Or _
                         strEmpName <> oRecordSet.Fields.Item("U_Technician").Value.ToString()) Then
                        strQuery = "Update [@Z_PMT1] Set "
                        strQuery += " U_CStatus = '" + oServiceCall.Status.ToString() + "'"
                        strQuery += " , U_Technician = '" + strEmpName + "'"
                        strQuery += " Where DocEntry = '" + strPMRef + "'"
                        strQuery += " And LineId = '" + oRecordSet.Fields.Item("LineId").Value.ToString() + "'"
                        oUpdateRecord.DoQuery(strQuery)
                    End If
                End If
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getEmployeeName(ByVal strEmpID As String) As String
        Dim _retVal As String = String.Empty
        Dim strQuery As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            Dim oUpdateRecord As SAPbobsCOM.Recordset
            oUpdateRecord = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select (T2.lastName +','+T2.firstName) As 'technician' From OHEM T2 Where EmpID = '" + strEmpID + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                _retVal = oRecordSet.Fields.Item(0).Value.ToString()
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function updateServiceContract(ByVal strPMRef As String)
        Dim _retVal As Boolean = False
        Dim strQuery As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            Dim oUpdateRecord As SAPbobsCOM.Recordset
            oUpdateRecord = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select U_SConNo,U_PMSchNo From [@Z_OPMT] T0 "
            strQuery += " Where T0.DocEntry = '" + strPMRef + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strQuery = "Update OCTR Set "
                strQuery += " U_PMSchNo = '" + oRecordSet.Fields.Item("U_PMSchNo").Value.ToString() + "'"
                strQuery += " Where ContractID = '" + oRecordSet.Fields.Item("U_SConNo").Value.ToString() + "'"
                oUpdateRecord.DoQuery(strQuery)
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub OpenFileDialogBox(ByVal oForm As SAPbouiCOM.Form, ByVal strPath As String, ByVal strFile As String)
        Dim _retVal As String = String.Empty
        Try
            FileOpen()
            CType(oForm.Items.Item(strPath).Specific, SAPbouiCOM.EditText).Value = strFilepath
            strFileName = Path.GetFileName(strFilepath)
            CType(oForm.Items.Item(strFile).Specific, SAPbouiCOM.EditText).Value = strFileName
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "FileOpen"
    Private Sub FileOpen()
        Try
            Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
            mythr.SetApartmentState(Threading.ApartmentState.STA)
            mythr.Start()
            mythr.Join()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ShowFileDialog()
        Try
            Dim oDialogBox As New OpenFileDialog
            Dim strMdbFilePath As String
            Dim oProcesses() As Process
            Try
                oProcesses = Process.GetProcessesByName("SAP Business One")
                If oProcesses.Length <> 0 Then
                    For i As Integer = 0 To oProcesses.Length - 1
                        Dim MyWindow As New clsListener.WindowWrapper(oProcesses(i).MainWindowHandle)
                        If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                            strMdbFilePath = oDialogBox.FileName
                            strFilepath = oDialogBox.FileName
                        Else
                        End If
                    Next
                End If
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

    Public Function getPicturePath()
        Dim _retVal As String = String.Empty
        Try
            Dim oCompanyService As SAPbobsCOM.CompanyService
            oCompanyService = oApplication.Company.GetCompanyService
            _retVal = oCompanyService.GetPathAdmin().PicturesFolderPath
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Public Sub Trace_ServiceCall(ByVal strContent As String, ByVal strFile As String)
        Try
            Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() + strFile
            If Not File.Exists(strPath) Then
                Dim fileStream As FileStream
                fileStream = New FileStream(strPath, FileMode.Create, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                sw.WriteLine(strContent)
                sw.Flush()
                sw.Close()
            Else
                Dim fileStream As FileStream
                fileStream = New FileStream(strPath, FileMode.Append, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                sw.WriteLine(strContent)
                sw.Flush()
                sw.Close()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
