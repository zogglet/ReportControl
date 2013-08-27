
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        report_ctrl.ReportServerURL = System.Configuration.ConfigurationManager.AppSettings("ReportURL")
        report_ctrl.ReportPath = System.Configuration.ConfigurationManager.AppSettings("ReportPath")
        report_ctrl.ReportName = "HourlyVolumesReport"
    End Sub
End Class
