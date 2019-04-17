Imports ESRI.ArcGIS.esriSystem
Namespace My
    Partial Friend Class MyApplication
        Private m_AOLicenseInitializer As LicenseInitializer = New Statistika1061.LicenseInitializer()

        Private Sub MyApplication_Startup(sender As Object, e As ApplicationServices.StartupEventArgs) Handles Me.Startup
            'ESRI License Initializer generated code.
            m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeBasic},
            New esriLicenseExtensionCode() {})
        End Sub

        Private Sub MyApplication_Shutdown(sender As Object, e As EventArgs) Handles Me.Shutdown
            'ESRI License Initializer generated code.
            'Do not make any call to ArcObjects after ShutDownApplication()
            m_AOLicenseInitializer.ShutdownApplication()
        End Sub
    End Class
End Namespace
