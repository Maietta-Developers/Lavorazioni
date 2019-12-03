<%@ Page Language="C#" AutoEventWireup="true" CodeFile="amzShowComunicazioni.aspx.cs" Inherits="amzShowComunicazioni" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Amazon - <%=OPERAZIONE %></title>
    <style type="text/css">
        body {
            font-family: Cambria;
        }
        #imgMerchant {
            vertical-align: middle;
        }
    </style>
</head>
    
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center;">
        <h1>Amazon - <%=OPERAZIONE %></h1>
        <hr />
        <asp:Image ID="imgTopLogo" runat="server" ImageUrl="~/pics/mcs-logo.png" /><br /><hr /><br /><br />
        <br /><br />
        <asp:Table runat="server" ID="tabComs" HorizontalAlign="Center" CellPadding="10" Width="700px">
            <asp:TableRow Font-Bold="true" BackColor="LightGray">
                <asp:TableCell>Mercato:</asp:TableCell>
                <asp:TableCell>Comunicazione:</asp:TableCell>
            </asp:TableRow>
            <asp:TableRow BackColor="LightGray">
                <asp:TableCell>
                    <asp:Image id="imgMerchant" Width="60" Height="40" runat="server" Visible="false" />&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:DropDownList ID="dropMerchant" runat="server" AutoPostBack="true" OnSelectedIndexChanged="dropMerchant_SelectedIndexChanged" ></asp:DropDownList></asp:TableCell>
                <asp:TableCell><asp:DropDownList ID="dropComs" runat="server" AutoPostBack="true" OnSelectedIndexChanged="dropComs_SelectedIndexChanged"></asp:DropDownList></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2"><hr /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2" HorizontalAlign="Left" ><asp:Label ID="labIDCom" runat="server" Font-Bold="true"></asp:Label></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left"><asp:Label ID="labSubject" runat="server"></asp:Label></asp:TableCell>
                <asp:TableCell>
                    <asp:CheckBox ID="chkInvoice" Enabled="false" runat="server" Text="Con ricevuta" Visible="false" />
                    &nbsp;&nbsp;<asp:CheckBox ID="chkAttach" Enabled="false" runat="server" Text="Con allegato" Visible="false" /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2">
                    <asp:Label ID="labTesto" runat="server"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <br /><br />
    </div>
    </form>
</body>
</html>
