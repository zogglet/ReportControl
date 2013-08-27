<%@ Control Language="VB" classname="zog.ReportControl" AutoEventWireup="false" CodeFile="ReportControl.ascx.vb" Inherits="ReportControl" %>

<div style="background:#cccccc;padding:5px;font-family:Verdana, Sans-Serif;font-size:11px;">
    <table border="0" width="100%">
        <tr>
            <td colspan="2">
                <%--Paramter inputs are rendered here--%>
                <div id="params_div" runat="server" visible="false" style="padding-bottom:5px;"></div>
            </td>
        </tr>
        <tr>
            <td style="width:50%;background:#a7a7a7;padding:2px;">
               
                <table style="vertical-align:middle;text-align:left;">
                    <tr>
                        <td>
                            <asp:Label ID="format_lbl" runat="server" AssociatedControlID="formats_ddl" Text="Select Output Format: " Font-Bold="true" ForeColor="#000000" />
                        </td>
                        <td>
                             <asp:DropDownList ID="formats_ddl" runat="server" 
                                BackColor="#eeeeee" Font-Size="11px" Font-Names="Verdana, Sans-Serif" borderColor="#000000" BorderStyle="Solid" BorderWidth="1px" >
                                <asp:ListItem Text="-- Select Format --" Value="-1" />
                                <asp:ListItem Text="PDF" Value="PDF" />
                                <asp:ListItem Text="Excel" Value="EXCEL" />
                                <asp:ListItem Text="Word" Value="WORD" />
                                <asp:ListItem Text="Image (TIFF)" Value="IMAGE" />
                             </asp:DropDownList>

                        </td>
                        <td>
                            <asp:Button ID="preview_btn" runat="server" Text="Preview/Print" BorderColor="#000000" BorderStyle="Solid" BorderWidth="2px" BackColor="#333333" 
                                ForeColor="#ffffff" Font-Bold="true" Font-Size="11px" Font-Names="Verdana, Sans-Serif" />
                        </td>
                        <td>
                            <asp:Button ID="render_btn" runat="server" ValidationGroup="RenderVGroup" Text="Render Report &raquo;" 
                                BorderColor="#000000" BorderStyle="Solid" BorderWidth="2px" BackColor="#333333" 
                                ForeColor="#ffffff" Font-Bold="true" Font-Size="11px" Font-Names="Verdana, Sans-Serif" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="4">
                            <asp:CompareValidator ID="format_cVal" runat="server" ValidationGroup="RenderVGroup" ControlToValidate="formats_ddl" Font-Bold="true" ForeColor="#990000" Operator="NotEqual" ValueToCompare="-1" ErrorMessage="Please select a format for export." Display="Dynamic" />
                            
                            <%--Since I need a reference across postbacks to the exported report file, and since I cannot effectively store it in
                            session since the session gets cleared during the deletion of 15 or more files (which could happen when purging the /Pages/ directory if the report 
                            contained 15 or more pages, causing the application to recycle, I am storing the value in this literal.--%>
                            <asp:Literal ID="fileRef_lit" runat="server" Visible="false" />
                        </td>
                    </tr>
                </table>
            </td>
            <td style="background:#a7a7a7;padding:2px;">

                <asp:Panel ID="rendered_pnl" runat="server" Visible="false">
                     <table>
                        <tr>
                            <td>
                                <asp:RadioButtonList ID="action_rbList" runat="server" RepeatDirection="Horizontal" Font-Bold="true" ForeColor="#000000">
                                    <asp:ListItem Text="Open/View" Value="Open" />
                                    <asp:ListItem Text="Save" Value="Save" Selected="True" />
                                </asp:RadioButtonList>
                            </td>
                            
                            <td>
                                <asp:Button ID="action_btn" runat="server" Text="Go &raquo;" BorderColor="#000000" BorderStyle="Solid" BorderWidth="2px" BackColor="#333333" 
                                    ForeColor="#ffffff" Font-Bold="true" Font-Size="11px" Font-Names="Verdana, Sans-Serif" ValidationGroup="EmailVGroup" />
                                &nbsp;<asp:Button ID="cancel_btn" runat="server" Text="Cancel" BorderColor="#000000" BorderStyle="Solid" BorderWidth="2px" BackColor="#333333" 
                                    ForeColor="#ffffff" Font-Bold="true" Font-Size="11px" Font-Names="Verdana, Sans-Serif"/>
                            </td>
                        </tr>
                    </table>
                    
                </asp:Panel>
                        
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <span style="display:block;width:750px;padding:2px 0 4px 0;">
                     <asp:Button ID="print_btn" runat="server" Visible="false" Text="Print" BorderColor="#000000" BorderStyle="Solid" BorderWidth="2px" BackColor="#333333" 
                        ForeColor="#ffffff" Font-Bold="true" Font-Size="11px" Font-Names="Verdana, Sans-Serif" />
                </span>
                <%--Preview pages are rendered here--%>
                <asp:Panel ID="preview_pnl" runat="server" Visible="false" style="overflow:auto;height:600px;width:750px;padding:2px;"></asp:Panel>
            </td>
        </tr>
    </table>
</div>

<asp:Panel id="printPreview_pnl" runat="server" visible="false"></asp:Panel>