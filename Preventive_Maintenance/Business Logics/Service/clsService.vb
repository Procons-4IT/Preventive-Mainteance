Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System
Public Class clsService
    Inherits clsBase
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private oMatrix As SAPbouiCOM.Matrix

    Public Sub New()
        MyBase.New()
    End Sub

#Region "Validation"
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Return True
    End Function
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.MenuUID
            Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            Case mnu_ADD
        End Select
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Service Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.BeforeAction
                Case True

                Case False
                    If BusinessObjectInfo.ActionSuccess Then
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                Dim strObjectKey As String = BusinessObjectInfo.ObjectKey
                                Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                                oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                                Dim strSerCall As String = oXmlDoc.SelectSingleNode("/ServiceCallParams/ServiceCallID").InnerText
                                oApplication.Utilities.updateRowPM(strSerCall)
                        End Select
                    End If
            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
