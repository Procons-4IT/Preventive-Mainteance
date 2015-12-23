Public Class clsStart

    Shared Sub Main()
        Try
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Try

                oApplication = New clsListener
                oApplication.Utilities.Connect()
                'oApplication.SetFilter()
                With oApplication.Company.GetCompanyService
                    CompanyDecimalSeprator = .GetAdminInfo.DecimalSeparator
                    CompanyThousandSeprator = .GetAdminInfo.ThousandsSeparator
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            End Try
            oApplication.Utilities.CreateTables()
            oApplication.Utilities.createEMMainAuthorization()
            oApplication.Utilities.AuthorizationCreation()
            oApplication.Utilities.AddRemoveMenus("Menu.xml")
            'oMenuItem = oApplication.SBO_Application.Menus.Item("mnu_Z_EM")
            'oMenuItem.Image = Application.StartupPath & "\EM.PNG"
            oApplication.Utilities.Message("Preventive Maintenance Addon Connected successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.NotifyAlert()
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class
