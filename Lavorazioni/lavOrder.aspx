<%@ Page Language="C#" AutoEventWireup="true" CodeFile="lavOrder.aspx.cs" Inherits="lavOrder" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <link href="Style/dettaglio.css" rel="stylesheet" type="text/css" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Lavorazioni - Ordine</title>
    
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center; font-family: Cambria;">
    <asp:Table ID="LavTab" runat="server" HorizontalAlign="Center" CellPadding="0" >
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3">
                <asp:Table ID="Table1" runat="server" Width="100%">
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="2" Font-Size="Small" Width="30%" HorizontalAlign="Left" Font-Bold="true" >
                            <h1>Lavorazione - <%= LAVID %></h1>
                        </asp:TableCell>
                        <asp:TableCell ColumnSpan="2" Font-Size="Small">
                            <b><%=Account %>&nbsp;-&nbsp;<%=TipoAccount %></b>
                            <br /><asp:Label ID="labRefresh" runat="server" Font-Size="X-Small"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Center">
                            <asp:Button ID="btnLogOut" runat="server" Text="LogOut" OnClick="btnLogOut_Click" Font-Size="Small" Width="70px" /><br /><br />
                            <asp:Button ID="btnHome" runat="server" Text="Home" OnClick="btnHome_Click" Font-Size="Small" Width="70px"></asp:Button>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:TableCell>                        
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3">
                <hr />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3" HorizontalAlign="Left">
                <asp:Table runat="server" Width="100%" ID="tabRivCl" >
                    <asp:TableRow>
                        <asp:TableCell Font-Size="Smaller" HorizontalAlign="Right" Width="30px" >
                            Rivenditore:&nbsp;&nbsp;
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Left">
                            <asp:Label ID="labRiv" runat="server" Font-Bold="true"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Font-Size="Small" HorizontalAlign="Right">
                            Cliente:&nbsp;&nbsp;
                        </asp:TableCell>
                        <asp:TableCell  HorizontalAlign="Left">
                            <asp:Label ID="labCliente" runat="server" Font-Bold="true"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3">
                <hr />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3" >
                <asp:Table runat="server" Width="100%" ID="tabName">
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="2"  HorizontalAlign="Left" Font-Size="X-Large">
                            Lavoro: &nbsp;&nbsp;
                            <asp:Label ID="labNomeLav" runat="server" Font-Bold="true" ></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="70%" HorizontalAlign="Left">
                            <asp:Label ID="labInserimento" runat="server" Font-Size="Smaller"></asp:Label><br />
                            <asp:Label ID="labConsegna" runat="server" Font-Size="Smaller"></asp:Label>
                            <br />
                        </asp:TableCell>
                        <asp:TableCell Font-Size="Smaller" HorizontalAlign="Left">
                            Inserito da:&nbsp;&nbsp;<asp:Label ID="labPropriet" runat="server" Font-Bold="true" ></asp:Label><br />
                            Approvato da:&nbsp;&nbsp;<asp:Label ID="labApprov" runat="server" Font-Bold="true" ></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell>
                <asp:Label ID="labInfo" runat="server" Font-Bold="true"></asp:Label></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3">
                <hr /><br />
                <asp:GridView ID="gridProdotti" runat="server" OnRowDataBound="gridProdotti_RowDataBound" Width="100%"></asp:GridView>
                <br />
                <hr /><br />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell>
                <asp:Label runat="server" ID="labPostB" Font-Bold="true" ForeColor="Red"></asp:Label></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell>
                <asp:Button ID="btnMakeOrder" runat="server" OnClick="btnMakeOrder_Click" Text="Crea ordine" Width="200" />
                <br /><br />
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
    </div>
    </form>
</body>
</html>
