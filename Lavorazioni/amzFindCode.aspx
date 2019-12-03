<%@ Page Language="C#" AutoEventWireup="true" CodeFile="amzFindCode.aspx.cs" Inherits="amzFindCode" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Amazon - <%=OPERAZIONE %> SKU - <%=COUNTRY %></title>
    <style>
        a {
            text-decoration: none;
            color: black;
        }
    </style>
    <script type="text/javascript">
        function checkOrdNum() {
            var code = document.getElementById("txFindCode").value;
            if (code.length <= 3) {
                alert("Stringa di ricerca troppo corta!");
                return (false);
            }
            else
                return (true);
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center; font-family: Cambria;">
    <h2>Amazon - Ricerca codice - <%=COUNTRY %></h2>
    <hr />
    <asp:Table runat="server" ID="mainTab" HorizontalAlign="Center" CellPadding="0">
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3"><asp:Image ID="imgTopLogo" runat="server" /><br /><hr /><br /> </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3" HorizontalAlign="Left"><asp:Label ID="labReturn" Text="Torna a panoramica" runat="server" Font-Size="Small" Font-Bold="true"></asp:Label>
                <br /><br /><br /></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell>Cerco:</asp:TableCell>
            <asp:TableCell><asp:RadioButton ID="rdbFindBySku" runat="server" GroupName="rdgFindG" Text="Codice SKU Amazon" Checked="true"></asp:RadioButton></asp:TableCell>
            <asp:TableCell><asp:RadioButton ID="rdbFindByCodiceMa" runat="server" GroupName="rdgFindG" Text="Codice Maietta"></asp:RadioButton></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="1"><br />Codice:</asp:TableCell>
            <asp:TableCell ColumnSpan="2"><br /><asp:TextBox ID="txFindCode" runat="server" Width="100%"></asp:TextBox></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3"><asp:Button ID="btnFindCode" runat="server" Text="Cerca!" Width="200px" OnClientClick="return checkOrdNum();" OnClick="btnFindCode_Click" /></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3" HorizontalAlign="Center">
                <br />
                <asp:GridView ID="gridResult" runat="server" OnRowDataBound="gridResult_RowDataBound" CellPadding="5" EmptyDataText = "Nessuna associazione trovata." ></asp:GridView>
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
    </div>
    </form>
</body>
</html>
