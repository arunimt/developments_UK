<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CommercialSensitiveView.aspx.cs" Inherits="SensitiveViewAddInWeb.Pages.CommercialSensitiveView" %>

<!DOCTYPE html>

<html>
<head>
    <title></title>
    <script type="text/javascript">
        // Set the style of the client web part page to be consistent with the host web.
        (function () {
            'use strict';

            var hostUrl = '';
            var link = document.createElement('link');
            link.setAttribute('rel', 'stylesheet');
            if (document.URL.indexOf('?') != -1) {
                var params = document.URL.split('?')[1].split('&');
                for (var i = 0; i < params.length; i++) {
                    var p = decodeURIComponent(params[i]);
                    if (/^SPHostUrl=/i.test(p)) {
                        hostUrl = p.split('=')[1];
                        link.setAttribute('href', hostUrl + '/_layouts/15/defaultcss.ashx');
                        break;
                    }
                }
            }
            if (hostUrl == '') {
                link.setAttribute('href', '/_layouts/15/1033/styles/themable/corev15.css');
            }
            document.head.appendChild(link);
        })();
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div></div>
        <div id="divSensitive" runat="server">
            <div>
                <div style="width: 40%; display: inline-block"></div>
                <div style="display: inline-block">
                    <asp:Label ID="Label1" runat="server" Text="Filter sensitive documents"></asp:Label>
                </div>
            </div>
            <div></div>
            <div>
                <div style="width: 30%; display: inline-block"></div>
                <div style="display: inline-block; padding-right: 10px">
                    <asp:DropDownList ID="ddlFilter" runat="server">
                        <asp:ListItem>Select Filter</asp:ListItem>
                        <asp:ListItem>FileName</asp:ListItem>
                        <asp:ListItem>Title</asp:ListItem>
                        <asp:ListItem>ContractID</asp:ListItem>
                        <asp:ListItem>DocumentType</asp:ListItem>
                        <asp:ListItem>Counterparty</asp:ListItem>
                        <asp:ListItem>Effective Date</asp:ListItem>
                        <asp:ListItem>MasterAgreementNumber</asp:ListItem>
                        <asp:ListItem>Library</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div style="display: inline-block; padding-right: 10px">
                    <asp:TextBox ID="txtSearch" runat="server"></asp:TextBox>
                </div>
                <div style="display: inline-block">
                    <asp:Button ID="btnSearch" runat="server" Text="Search" OnClick="BtnSearch_Click" />
                </div>
            </div>
            <div></div>
            <div></div>
            <div>
                <asp:GridView ID="GridSensitiveView" runat="server" BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" AllowPaging="True" OnPageIndexChanging="GridSensitiveView_PageIndexChanging">
                    <FooterStyle BackColor="White" ForeColor="#000066" />
                    <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                    <RowStyle ForeColor="#000066" />
                    <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                    <SortedAscendingCellStyle BackColor="#F1F1F1" />
                    <SortedAscendingHeaderStyle BackColor="#007DBB" />
                    <SortedDescendingCellStyle BackColor="#CAC9C9" />
                    <SortedDescendingHeaderStyle BackColor="#00547E" />
                </asp:GridView>

            </div>
        </div>
        <div id="divError" runat="server">
            <asp:Label ID="lblError" runat="server" ForeColor="#FF3300"></asp:Label>
        </div>
    </form>
</body>
</html>
