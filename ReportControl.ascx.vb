Imports Microsoft.Reporting.WebForms
Imports System.IO
Imports System.IO.Packaging

Partial Class ReportControl
    Inherits System.Web.UI.UserControl

    Private _reportServerURL As String
    Private _reportPath As String
    Private _reportName As String

    Private _rv As ReportViewer
    Private _exportPath As String

    Dim warnings As Warning() = Nothing
    Dim streamIDs As String() = Nothing
    Dim mimeType As String = Nothing
    Dim encoding As String = Nothing
    Dim ext As String = Nothing
    Dim bytes As Byte()

    Private _param_lbl As Label
    Private _param_txt As TextBox
    Private _param_ddl As DropDownList
    Private _param_rbList As RadioButtonList
    Private _preview_img As Image

    Private msgScript As String = ""

#Region "Exposed Contained Controls"

    '***** The contained label/prompt for the format selection
    Public Property FormatLabel() As Label
        Get
            Return format_lbl
        End Get
        Set(ByVal value As Label)
            format_lbl = value
        End Set
    End Property

    '***** The contained DropDownList for the selectable formats
    Public Property FormatDDL() As DropDownList
        Get
            Return formats_ddl
        End Get
        Set(ByVal value As DropDownList)
            formats_ddl = value
        End Set
    End Property

    '***** The contained Button to preview the report
    Public Property PreviewButton() As Button
        Get
            Return preview_btn
        End Get
        Set(ByVal value As Button)
            preview_btn = value
        End Set
    End Property

    '***** The contained Button to print a previewed report
    Public Property PrintButton() As Button
        Get
            Return print_btn
        End Get
        Set(ByVal value As Button)
            print_btn = value
        End Set
    End Property

    '***** The contained Button to render the report
    Public Property RenderButton() As Button
        Get
            Return render_btn
        End Get
        Set(ByVal value As Button)
            render_btn = value
        End Set
    End Property

    '***** The contained RadioButtonList for the actions to take
    Public Property ActionsRBList() As RadioButtonList
        Get
            Return action_rbList
        End Get
        Set(ByVal value As RadioButtonList)
            action_rbList = value
        End Set
    End Property

    '***** The contained button for performing an action after the report has rendered
    Public Property ActionButton() As Button
        Get
            Return action_btn
        End Get
        Set(ByVal value As Button)
            action_btn = value
        End Set
    End Property

    '***** The contained button to cancel the action taken after the report has rendered
    Public Property CancelButton() As Button
        Get
            Return cancel_btn
        End Get
        Set(ByVal value As Button)
            cancel_btn = value
        End Set
    End Property


#End Region

#Region "ReportControl Properties"

    '***** Enables/disables the control
    Public WriteOnly Property Enabled() As Boolean
        Set(ByVal value As Boolean)
            formats_ddl.Enabled = value
            rendered_pnl.Enabled = value
            params_div.Visible = value
            format_cVal.Enabled = value
        End Set
    End Property


    '***** The URL of the server where the report resides. Required in order to use this control.
    Public Property ReportServerURL() As String
        Get
            Return _reportServerURL
        End Get
        Set(ByVal value As String)
            _reportServerURL = value
        End Set
    End Property

    '***** The directory on the server where the report resides. Required in order to use this control.
    Public Property ReportPath() As String
        Get
            Return _reportPath
        End Get
        Set(ByVal value As String)
            _reportPath = value
        End Set
    End Property

    '***** The name of the report. Required in order to use this control.
    Public Property ReportName() As String
        Get
            Return _reportName
        End Get
        Set(ByVal value As String)
            _reportName = value
        End Set
    End Property

#End Region



    '*********************************************************************************************

    Protected Sub cancel_btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cancel_btn.Click
        rendered_pnl.Visible = False
        formats_ddl.SelectedValue = -1
        preview_pnl.Visible = False
        preview_btn.Visible = False

        If File.Exists(_exportPath & "\" & fileRef_lit.Text) Then
            File.Delete(_exportPath & "\" & fileRef_lit.Text)
        End If

        If Directory.Exists(_exportPath & "\Pages\") Then
            Dim files() As String = Directory.GetFiles(_exportPath & "\Pages\", "*")
            For i As Integer = 0 To files.Length - 1
                File.Delete(files(i))
            Next
            'Directory.Delete(_exportPath & "\Pages\")
        End If

        For i As Integer = 0 To CType(Session("ParamsTable"), Table).Rows(0).Cells.Count - 1
            For Each ctrl As WebControl In CType(Session("ParamsTable"), Table).Rows(0).Cells(i).Controls
                Select Case ctrl.GetType().ToString()
                    Case "System.Web.UI.WebControls.TextBox"
                        CType(ctrl, TextBox).Text = ""
                    Case "System.Web.UI.WebControls.DropDownList"
                        CType(ctrl, DropDownList).SelectedIndex = 0
                    Case "System.Web.UI.WebControls.RadioButtonList"
                        CType(ctrl, RadioButtonList).SelectedIndex = 0
                End Select
            Next

        Next

    End Sub


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        _exportPath = Page.MapPath(".")

        _rv = New ReportViewer()
        _rv.ServerReport.ReportPath = _reportPath & _reportName
        _rv.ServerReport.ReportServerUrl = New Uri(_reportServerURL)

        If Not IsPostBack Then
            If File.Exists(_exportPath & "\" & fileRef_lit.Text) Then
                File.Delete(_exportPath & "\" & fileRef_lit.Text)
            End If

            If Directory.Exists(_exportPath & "\Pages\") Then
                Dim files() As String = Directory.GetFiles(_exportPath & "\Pages\", "*")
                For i As Integer = 0 To files.Length - 1
                    File.Delete(files(i))
                Next
                'Directory.Delete(_exportPath & "\Pages\")
            End If
        End If



        Try
            If _rv.ServerReport.GetParameters().Count > 0 Then

                For Each p As ReportParameterInfo In _rv.ServerReport.GetParameters()

                    If p.IsQueryParameter = True Then
                        params_div.Visible = True
                        Exit For
                    End If
                Next

                Dim params_tbl As Table = New Table
                params_tbl.ID = "params_tbl"
                params_div.Controls.Add(params_tbl)

                Dim row As TableRow = New TableRow()
                Dim cell As TableCell = Nothing

                cell = New TableCell()
                cell.Controls.Add(New LiteralControl("<span style='color:#000000;font-weight:bold;'>Parameters:</span>&nbsp;&nbsp;"))
                row.Cells.Add(cell)

                For Each p As ReportParameterInfo In _rv.ServerReport.GetParameters()

                    'Only parameters requiring input
                    If p.IsQueryParameter = True Then

                        cell = New TableCell()

                        _param_lbl = New Label()
                        _param_lbl.ID = p.Name & "_lbl"
                        _param_lbl.Text = p.Name & ": "
                        _param_lbl.ForeColor = Drawing.Color.Black

                        If p.ValidValues IsNot Nothing Then

                            If p.ValidValues.Count > 1 Then

                                Dim listItem As ListItem

                                'Boolean parameters
                                If p.DataType = ParameterDataType.Boolean Then

                                    _param_rbList = New RadioButtonList()
                                    _param_rbList.ID = p.Name & "_rbList"
                                    _param_rbList.RepeatDirection = RepeatDirection.Horizontal
                                    _param_rbList.ForeColor = Drawing.Color.Black
                                    _param_lbl.AssociatedControlID = _param_rbList.ID

                                    For i As Integer = 0 To p.ValidValues.Count - 1
                                        listItem = New ListItem()
                                        listItem.Text = p.ValidValues(i).Label
                                        listItem.Value = IIf(p.ValidValues(i).Value = Nothing, Nothing, p.ValidValues(i).Value)
                                        listItem.Selected = (i = 0)
                                        _param_rbList.Items.Add(listItem)
                                    Next

                                    Session(_param_rbList.ID) = _param_rbList
                                    cell.Controls.Add(_param_lbl)
                                    cell.Controls.Add(_param_rbList)

                                    'Parameters with multiple available values
                                Else
                                    _param_ddl = New DropDownList()
                                    _param_ddl.ID = p.Name & "_ddl"
                                    _param_lbl.AssociatedControlID = _param_ddl.ID

                                    formatInput(_param_ddl)

                                    For i As Integer = 0 To p.ValidValues.Count - 1
                                        listItem = New ListItem()
                                        listItem.Text = p.ValidValues(i).Label
                                        listItem.Value = IIf(p.ValidValues(i).Value = Nothing, Nothing, p.ValidValues(i).Value)
                                        _param_ddl.Items.Add(listItem)
                                    Next

                                    Session(_param_ddl.ID) = _param_ddl
                                    cell.Controls.Add(_param_lbl)
                                    cell.Controls.Add(_param_ddl)

                                End If
                            End If

                            'Parameters with one valid value (thus with no validValues)
                        Else
                            _param_txt = New TextBox()
                            _param_txt.ID = p.Name & "_txt"
                            _param_txt.Width = 90
                            _param_lbl.AssociatedControlID = _param_txt.ID

                            formatInput(_param_txt)

                            'So that I can access it elsewhere
                            Session(_param_txt.ID) = _param_txt

                            cell.Controls.Add(_param_lbl)
                            cell.Controls.Add(_param_txt)

                        End If

                        row.Cells.Add(cell)
                    End If


                    'Add default values
                    If p.Values.Count > 0 Then
                        Dim rp() As ReportParameter = {New ReportParameter(p.Name, p.Values.ToArray, p.IsQueryParameter)}
                        _rv.ServerReport.SetParameters(rp)
                    End If

                Next

                params_tbl.Rows.Add(row)
                Session("ParamsTable") = params_tbl
            End If

        Catch ex As Exception
            msgScript = "<script language='javascript'>" & _
                                          "alert('Error:\n" & ex.Message & "');" & _
                                          "</script>"
            Page.ClientScript.RegisterStartupScript(Me.GetType(), "msgScript", msgScript)
        End Try

    End Sub


    Protected Sub render_btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles render_btn.Click

        renderReport(False)
        rendered_pnl.Visible = True

    End Sub


    Private Sub renderReport(ByVal preview As Boolean)

        Dim exportFile As String = ""
        Dim matchFound As Boolean = False

        Try
            _rv = New ReportViewer()

            _rv.ServerReport.ReportPath = _reportPath & _reportName
            _rv.ServerReport.ReportServerUrl = New Uri(_reportServerURL)

            '**** Configure parameters
            For Each p As ReportParameterInfo In _rv.ServerReport.GetParameters()

                If p.IsQueryParameter Then
                    If p.ValidValues IsNot Nothing Then

                        If p.ValidValues.Count > 1 Then

                            'Boolean Parameters
                            If p.DataType = ParameterDataType.Boolean Then
                                If p.AllowBlank = False Then

                                    For i As Integer = 0 To p.ValidValues.Count - 1
                                        If CType(Session(p.Name & "_rbList"), RadioButtonList).SelectedValue = p.ValidValues(i).Value Then
                                            matchFound = True
                                            Exit For
                                        End If
                                    Next

                                    If matchFound = True Then
                                        matchFound = False
                                        Dim rp() As ReportParameter = {New ReportParameter(p.Name, CBool(CType(Session(p.Name & "_rbList"), RadioButtonList).SelectedValue), True)}
                                        _rv.ServerReport.SetParameters(rp)
                                    Else
                                        If p.Nullable = False Then
                                            Throw New ArgumentException("The parameter, " & p.Name & ", requires a valid value.")
                                        Else
                                            Dim rp() As ReportParameter = {New ReportParameter(p.Name, New String() {Nothing}, True)}
                                            _rv.ServerReport.SetParameters(rp)
                                        End If

                                    End If

                                Else
                                    Dim rp() As ReportParameter = {New ReportParameter(p.Name, CBool(CType(Session(p.Name & "_rbList"), RadioButtonList).SelectedValue), True)}
                                    _rv.ServerReport.SetParameters(rp)
                                End If

                                'Parameters with multiple available values
                            Else
                                If p.AllowBlank = False Then

                                    For i As Integer = 0 To p.ValidValues.Count - 1
                                        If CType(Session(p.Name & "_ddl"), DropDownList).SelectedValue = p.ValidValues(i).Value Then
                                            matchFound = True
                                            Exit For
                                        End If
                                    Next

                                    If matchFound = True Then
                                        matchFound = False
                                        Dim rp() As ReportParameter = {New ReportParameter(p.Name, CType(Session(p.Name & "_ddl"), DropDownList).SelectedValue, True)}
                                        _rv.ServerReport.SetParameters(rp)
                                    Else
                                        If p.Nullable = False Then
                                            Throw New ArgumentException("The parameter, " & p.Name & ", requires a valid value.")
                                        Else
                                            Dim rp() As ReportParameter = {New ReportParameter(p.Name, New String() {Nothing}, True)}
                                            _rv.ServerReport.SetParameters(rp)
                                        End If

                                    End If
                                Else
                                    Dim rp() As ReportParameter = {New ReportParameter(p.Name, CType(Session(p.Name & "_ddl"), DropDownList).SelectedValue, True)}
                                    _rv.ServerReport.SetParameters(rp)
                                End If

                            End If

                        End If

                        'Parameters with one valid value (thus with no validValues)
                    Else

                        If p.AllowBlank = False Then
                            If CType(Session(p.Name & "_txt"), TextBox).Text.Trim.Length = 0 Then
                                If p.Nullable = False Then
                                    Throw New ArgumentException("The parameter, " & p.Name & ", requires a valid value.")
                                Else
                                    Dim rp() As ReportParameter = {New ReportParameter(p.Name, New String() {Nothing}, True)}
                                    _rv.ServerReport.SetParameters(rp)
                                End If
                            Else
                                Select Case p.DataType
                                    Case ParameterDataType.DateTime
                                        Dim rp() As ReportParameter = {New ReportParameter(p.Name, CDate(CType(Session(p.Name & "_txt"), TextBox).Text.Trim), True)}
                                        _rv.ServerReport.SetParameters(rp)
                                    Case ParameterDataType.Integer
                                        Dim rp() As ReportParameter = {New ReportParameter(p.Name, CInt(CType(Session(p.Name & "_txt"), TextBox).Text.Trim), True)}
                                        _rv.ServerReport.SetParameters(rp)
                                    Case ParameterDataType.String
                                        Dim rp() As ReportParameter = {New ReportParameter(p.Name, CType(Session(p.Name & "_txt"), TextBox).Text.Trim, True)}
                                        _rv.ServerReport.SetParameters(rp)
                                End Select
                            End If
                        Else
                            Select Case p.DataType
                                Case ParameterDataType.DateTime
                                    Dim rp() As ReportParameter = {New ReportParameter(p.Name, CDate(CType(Session(p.Name & "_txt"), TextBox).Text.Trim), True)}
                                    _rv.ServerReport.SetParameters(rp)
                                Case ParameterDataType.Integer
                                    Dim rp() As ReportParameter = {New ReportParameter(p.Name, CInt(CType(Session(p.Name & "_txt"), TextBox).Text.Trim), True)}
                                    _rv.ServerReport.SetParameters(rp)
                                Case ParameterDataType.String
                                    Dim rp() As ReportParameter = {New ReportParameter(p.Name, CType(Session(p.Name & "_txt"), TextBox).Text.Trim, True)}
                                    _rv.ServerReport.SetParameters(rp)
                            End Select
                        End If

                    End If

                End If

            Next

            '****** Do the rendering and writing of the file

            'Delete any prior rendered reports
            If File.Exists(_exportPath & "\" & fileRef_lit.Text) Then
                File.Delete(_exportPath & "\" & fileRef_lit.Text)
            End If

            If Directory.Exists(_exportPath & "\Pages\") Then
                Dim files() As String = Directory.GetFiles(_exportPath & "\Pages\", "*")
                For i As Integer = 0 To files.Length - 1
                    File.Delete(files(i))
                Next
                'Directory.Delete(_exportPath & "\Pages\")
            End If


            If preview Then
                renderReportPages(True)
                renderPreview(preview_pnl, 725)
            Else
                'Render and zip if Image is chosen
                If formats_ddl.SelectedValue = "IMAGE" Then
                    renderReportPages(False)

                    Dim newZipPath As String = _exportPath & "\" & _reportName & "_allPages.zip"
                    If File.Exists(newZipPath) Then
                        File.Delete(newZipPath)
                    End If

                    zipFile(newZipPath)

                    exportFile = _reportName & "_allPages.zip"

                    Dim files() As String = Directory.GetFiles(_exportPath & "\Pages\", "*")
                    For i As Integer = 0 To files.Length - 1
                        File.Delete(files(i))
                    Next
                    'Directory.Delete(_exportPath & "\Pages\")


                Else
                    bytes = _rv.ServerReport.Render(formats_ddl.SelectedValue, Nothing, mimeType, encoding, ext, streamIDs, warnings)

                    Select Case formats_ddl.SelectedValue
                        Case "EXCEL"
                            exportFile = _reportName & ".xls"
                        Case "WORD"
                            exportFile = _reportName & ".doc"
                        Case "PDF"
                            exportFile = _reportName & ".pdf"
                    End Select

                    Dim fStream As New FileStream(_exportPath & "\" & exportFile, FileMode.Create)
                    fStream.Write(bytes, 0, bytes.Length)
                    fStream.Close()
                End If
            End If

            fileRef_lit.Text = exportFile

            preview_pnl.Visible = preview
            print_btn.Visible = preview
            rendered_pnl.Visible = Not preview

        Catch ex As ArgumentException
            msgScript = "<script language='javascript'>" & _
                                          "alert('Error:\n" & ex.Message & "');" & _
                                          "</script>"
            Page.ClientScript.RegisterStartupScript(Me.GetType(), "msgScript", msgScript)
        Catch ex As Exception
            msgScript = "<script language='javascript'>" & _
                                          "alert('Error:\n" & ex.Message & "');" & _
                                          "</script>"
            Page.ClientScript.RegisterStartupScript(Me.GetType(), "msgScript", msgScript)
        End Try

    End Sub

    Private Sub renderPreview(ByVal ctrl As Control, ByVal width As Unit)

        Dim files() As String = Directory.GetFiles(_exportPath & "\Pages\", "*")

        For i As Integer = 0 To files.Count - 1

            _preview_img = New Image()

            With _preview_img
                .ID = _reportName & "_img_" & i
                .ImageUrl = "Pages" & String.Concat("/", IO.Path.GetFileName(files(i)))
                .Width = width
                .BorderColor = Drawing.ColorTranslator.FromHtml("#333333")
                .BorderStyle = BorderStyle.Solid
                .BorderWidth = 2
            End With
            ctrl.Controls.Add(_preview_img)
            ctrl.Controls.Add(New LiteralControl("<br /><br />"))

        Next
    End Sub

    Private Sub zipFile(ByVal zipPath As String)

        Dim zip As Package = ZipPackage.Open(zipPath, FileMode.Create, FileAccess.ReadWrite)

        Dim files() As String = Directory.GetFiles(_exportPath & "\Pages\", "*")
        For i As Integer = 0 To files.Length - 1

            Dim zipUri As String = String.Concat("/", IO.Path.GetFileName(files(i)))
            Dim partUri As New Uri(zipUri, UriKind.Relative)

            Dim pkgPart As PackagePart = zip.CreatePart(partUri, Net.Mime.MediaTypeNames.Application.Zip, CompressionOption.Normal)

            Dim bytes As Byte() = File.ReadAllBytes(files(i))
            pkgPart.GetStream().Write(bytes, 0, bytes.Length)
            pkgPart.GetStream().Close()
        Next

        zip.Close()

    End Sub

    'Render needed to obtain number of pages
    Private Function getTotalReportPages() As Integer
        bytes = _rv.ServerReport.Render("IMAGE", "<DeviceInfo><PageHeight>8.5in</PageHeight><PageWidth>14in</PageWidth><OutputFormat>EMF</OutputFormat></DeviceInfo>", mimeType, encoding, ext, streamIDs, warnings)
        Return streamIDs.Length
    End Function

    Private Sub renderReportPages(ByVal preview As Boolean)

        Directory.CreateDirectory(_exportPath & "\Pages\")

        For i As Integer = 0 To getTotalReportPages()
            bytes = _rv.ServerReport.Render("IMAGE", "<DeviceInfo><PageHeight>8.5in</PageHeight><PageWidth>14in</PageWidth><OutputFormat>" & IIf(preview, "PNG", "TIFF") & "</OutputFormat><StartPage>" & i & "</StartPage></DeviceInfo>", mimeType, encoding, ext, streamIDs, warnings)

            Dim fStream As New FileStream(_exportPath & "\Pages\" & _reportName & "_" & i + 1 & IIf(preview, ".png", ".tif"), FileMode.Create)
            fStream.Write(bytes, 0, bytes.Length)
            fStream.Close()
        Next

    End Sub

    Protected Sub action_btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles action_btn.Click
        Select Case action_rbList.SelectedValue
            Case "Open"
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "newWindow", "window.open('" & fileRef_lit.Text & "','_blank');", True)
            Case "Save"
                saveContent(formats_ddl.SelectedValue)
        End Select
    End Sub

    Private Sub saveContent(ByVal type As String)

        Select Case type
            Case "IMAGE"
                Response.ContentType = "application/x-zip-compressed"
            Case "PDF"
                Response.ContentType = "application/pdf"
            Case "EXCEL"
                Response.ContentType = "application/vnd.ms-excel"
            Case "WORD"
                Response.ContentType = "application/msword"
        End Select
        Response.AppendHeader("Content-Disposition", "attachment; filename=" & fileRef_lit.Text)
        Response.TransmitFile(fileRef_lit.Text)
        Response.End()

    End Sub

    Private Sub printElement(ByVal ctrl As WebControl)

        Dim stringWrite As StringWriter = New StringWriter()
        Dim htmlWrite As HtmlTextWriter = New HtmlTextWriter(stringWrite)
        Dim form As HtmlForm = New HtmlForm()
        Dim page As Page = New Page()

        Dim backBtn As Button = New Button()

        form.Attributes.Add("runat", "server")
        page.EnableEventValidation = False

        With backBtn
            .ID = "back_btn"
            .Attributes.Add("onclick", "history.back(); return false")
            .Text = "Back"
            .ValidationGroup = "Nothing"
            .BorderColor = Drawing.ColorTranslator.FromHtml("#000000")
            .BackColor = Drawing.ColorTranslator.FromHtml("#333333")
            .BorderStyle = BorderStyle.Solid
            .BorderWidth = 2
            .ForeColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
            .Font.Bold = True
            .Font.Size = 9
            .Font.Name = "Verdana"
        End With


        page.Controls.Add(form)
        form.Controls.Add(New LiteralControl("<span style='font-family:Verdana, Sans-Serif;font-size:16px;font-weight:bold;display:block;text-align:center;padding-bottom: 10px;'>" & _reportName & "</span>"))
        form.Controls.Add(backBtn)
        form.Controls.Add(New LiteralControl("<br /><br />"))
        form.Controls.Add(ctrl)

        page.DesignerInitialize()
        page.RenderControl(htmlWrite)

        With HttpContext.Current.Response
            .Clear()
            .Write(stringWrite.ToString())
            .Write("<script>window.print();</script>")
            .Flush()
            .End()
        End With


    End Sub

    Private Sub formatInput(ByVal ctrl As WebControl)
        With ctrl
            .BackColor = Drawing.ColorTranslator.FromHtml("#eeeeee")
            .Font.Size = 8
            .Font.Name = "Verdana"
            .BorderColor = Drawing.ColorTranslator.FromHtml("#000000")
            .BorderStyle = BorderStyle.Solid
            .BorderWidth = 1
        End With
    End Sub

    Protected Sub preview_btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles preview_btn.Click
        renderReport(True)

    End Sub

    Protected Sub print_btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles print_btn.Click
        renderPreview(printPreview_pnl, Unit.Percentage(100))
        printPreview_pnl.Visible = True
        printElement(printPreview_pnl)
    End Sub
End Class
