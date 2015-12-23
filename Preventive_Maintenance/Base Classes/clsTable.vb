Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try

            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OUSR" Or strTab = "OITW" Or strTab = "RDR1" Or strTab = "OINV" Or strTab = "INV1" Or strTab = "OWOR" Or strTab = "ORDR" Or strTab = "OCLG" Or strTab = "ORCT" Or strTab = "OCTR" Or strTab = "CTR1" Or strTab = "OSCL" Or strTab = "OINS" Or strTab = "OSCO") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            If Not (TableName = "ORDR" Or TableName = "OITM" Or TableName = "OCTR" Or TableName = "CTR1" Or TableName = "OSCO" Or TableName = "OSCP" Or TableName = "OSCT") Then
                TableName = "@" + TableName
            End If

            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            Else
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            objUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()

        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)
            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                Optional ByVal sFind3 As String = "", _
                                Optional ByVal sFind4 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document)

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" And sFind3 <> "" And sFind3 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(2)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind3
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(3)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind4
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************

    Public Sub CreateTables()
        Try
            oApplication.SBO_Application.StatusBar.SetText("Initializing Database...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oApplication.Company.StartTransaction()


            AddFields("OINV", "PMRef", "PM Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OSCL", "PMRef", "PM Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            ' addField("OCTR", "FQType", "Frequency Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "M,Q,H,Y", "Month,Quarter,Half Yearly,Yearly", "M")
            addField("OCTR", "Freqency", "Frequency", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OCTR", "NBillDt", "Next Bill Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OCTR", "PMSchNo", "PM Schedule No", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("CTR1", "QuarterAmt", "Quarter Amt", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("OINS", "Queue", "Queue", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OINS", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("OSCO", "ConPM", "Consider In PM", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "NO,YES", "N")
            addField("OSCP", "ConPM", "Consider In PM", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "NO,YES", "N")
            addField("OSCT", "ConPM", "Consider In PM", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "NO,YES", "N")

            AddTables("Z_OPMT", "Preventive Maintenance", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OPMT", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OPMT", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OPMT", "SConNo", "Service Contract No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("Z_OPMT", "SEStartDt", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPMT", "SEEndDt", "End Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPMT", "BLDueDt", "Bill Due Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPMT", "PMSchNo", "PM Schedule No", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPMT", "PMStartDt", "PM Start Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPMT", "PMEndDt", "PM End Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_OPMT", "InvDE", "Invoice Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OPMT", "InvDN", "Invoice No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("Z_OPMT", "InvDt", "Invoice Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_OPMT", "PMStatus", "PM Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,L,C", "Open,Cancelled,Closed", "O")
            addField("Z_OPMT", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            AddTables("Z_PMT1", "Preventive Maintenance Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PMT1", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PMT1", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PMT1", "InsID", "Customer Eq Card", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PMT1", "MSerialNo", "Manufacture Serial No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 32)
            AddFields("Z_PMT1", "MSerialNo", "Manufacture Serial No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 32)
            AddFields("Z_PMT1", "SerialNo", "Serial No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 32)
            addField("Z_PMT1", "SEStartDt", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_PMT1", "SEEndDt", "End Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("Z_PMT1", "TerDt", "Termination Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_PMT1", "SCallNo", "Service Call No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("Z_PMT1", "CDate", "Create Date", SAPbobsCOM.BoFieldTypes.db_Date, 10, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            AddFields("Z_PMT1", "Technician", "Technician", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PMT1", "CStatus", "Call Status", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("Z_PMT1", "Amount", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("Z_PMT1", "NoBill", "No Bill.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            addField("Z_PMT1", "BillAmt", "Bill Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("Z_PMT1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddTables("Z_OOCR", "PM Cost Center", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OOCR", "OcrCode", "Dimention 1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_OOCR", "OcrCode2", "Dimention 2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_OOCR", "CogsOcrCode", "Cogs Dimention 1", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_OOCR", "CogsOcrCode2", "Cogs Dimention 2", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)

            'addField("ORDR", "OrderType", "Order Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,P", "None,PMContract", "N")
            'addField("ORDR", "QuType", "Qu Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,S", "None,Service", "N")
            'AddFields("ORDR", "Contract", "Contract", SAPbobsCOM.BoFieldTypes.db_Alpha, , 254)

            '---- User Defined Object
            CreateUDO()

            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.SBO_Application.StatusBar.SetText("Database creation completed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub CreateUDO()
        Try
            AddUDO("Z_OPMT", "Preventive_Maintenance", "Z_OPMT", "DocNum", "U_CardCode", "U_SConNo", "U_PMSchNo", "Z_PMT1", SAPbobsCOM.BoUDOObjType.boud_Document) ' Event Booking - Document
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
