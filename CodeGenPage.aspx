<%@ Page Language="VB" AutoEventWireup="true" CodeFile="CodeGenPage.aspx.vb" Inherits="CodeGenPage" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Code Previewer and Generator</title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
        <asp:UpdateProgress ID="UpdateProgress1" runat="server" DisplayAfter="100" DynamicLayout="true" AssociatedUpdatePanelID="UpdatePanel1">
            <ProgressTemplate>
            <div align="center" style="font-family:Tahoma;font-size:large;font-weight:bold;">
                L O A D I N G . . . .
            </div>
            </ProgressTemplate>
        </asp:UpdateProgress>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
            <table>
                <tr>
                    <td valign="top">
                    <asp:ListBox ID="ListBox1" runat="server" BackColor="whitesmoke" Font-Names="Tahoma" Font-Size="Smaller" Height="600" DataSourceID="TableList" DataTextField="TABLE_NAME"
                        DataValueField="TABLE_NAME" SelectionMode="multiple" AutoPostBack="True" ></asp:ListBox><br />
                        <asp:CheckBox ID="chkOverwriteFiles" runat="server" Text="Overwrite Files?" Checked="true" Font-Names="Tahoma" Font-Size="Smaller"   /><br />
                        <asp:LinkButton ID="lbtnGenerateCode" runat="server" Text="Generate Code" Font-Names="Tahoma" Font-Size="Smaller" OnClientClick="if (confirm('Are you sure you would like to generate the class and interface files for the selected objects?'));return true;"  />
                    </td>
                    <td valign="top">
                        <asp:Panel ScrollBars="Auto" Height="600" Width="900" runat="server" BackColor=WhiteSmoke Wrap="false" ID="PanelCode" Font-Names="Courier New" Font-Size="x-Small" BorderColor=DarkGray BorderStyle="solid" BorderWidth="1">
                        </asp:Panel>  
                    </td>
                </tr>
            </table>
            </ContentTemplate>
        </asp:UpdatePanel>
        <asp:ObjectDataSource ID="TableList" runat="server" SelectMethod="GetDistinctTables" TypeName="NewLeaf.CodeGenerator"></asp:ObjectDataSource>
        &nbsp;
    </form>
</body>
</html>
