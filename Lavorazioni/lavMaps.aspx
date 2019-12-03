<%@ Page Language="C#" AutoEventWireup="true" CodeFile="lavMaps.aspx.cs" Inherits="lavMaps" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Lavorazioni - Maps</title>
    <style type="text/css">
        #tabMaps {
            border-collapse:collapse;
            width: 1000px;
        }
        .tabProd {
            width: 100%;
            background-color: #009FE3;
        }
        a {
            text-decoration: none;
            color: black;
        }

        #trSearch {
            width: 100%;
        }

        #tabIntest {
            width: 900px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center; font-family: Cambria;">
        <asp:Table ID="tabIntest" runat="server" HorizontalAlign="Center">
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2" Width="60%" Font-Bold="true" >
                    <h1>Lavorazioni - Maps <%= CODE %></h1>
                </asp:TableCell>
                <asp:TableCell Font-Size="Small">
                    <b><%=Account %>&nbsp;-&nbsp;<%=TipoAccount %>&nbsp;-&nbsp;
                    
                    <asp:DropDownList ID="dropTypeOper" runat="server" Visible="false" AutoPostBack="true" Font-Size="X-Small" ></asp:DropDownList></b>
                    <br /><br />
                    <asp:Button ID="btnLogOut" runat="server" Text="LogOut" OnClick="btnLogOut_Click" Font-Size="Small" Width="70px" />
                    &nbsp;-&nbsp;
                    <asp:HyperLink runat="server" ID="hylGoLav" Text="Vai a lavorazioni" Font-Bold="true" Font-Size="Small" ></asp:HyperLink>
                    <br /><br />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="3">
                    <br />
                    <asp:Image ID="imgTopLogo" runat="server" />
                    <br /><br /><hr /><br />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow ID="trSearch" runat="server">
                <asp:TableCell Width="100%" ColumnSpan="3">
                    <asp:Panel ID="panSearch" runat="server" Width="100%" DefaultButton="btnFindCode">
                        <asp:Table ID="tabSearch" runat="server" Width="100%" >
                            <asp:TableRow>
                                <asp:TableCell>Trova Codice</asp:TableCell>
                                <asp:TableCell><asp:TextBox ID="txFindCode" runat="server"></asp:TextBox></asp:TableCell>
                                <asp:TableCell><asp:Button ID="btnFindCode" runat="server" Text="Trova" Width="200" OnClick="btnFindCode_Click" /></asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:Panel>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow ID="trBar" runat="server">
                <asp:TableCell ColumnSpan="3">
                    <br /><hr /><br />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Table id ="tabMaps" runat="server" HorizontalAlign="Center" CellPadding="4"></asp:Table>
    </div>
    </form>
</body>
</html>
