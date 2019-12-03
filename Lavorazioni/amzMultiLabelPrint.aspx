<%@ Page Language="C#" AutoEventWireup="true" CodeFile="amzMultiLabelPrint.aspx.cs" Inherits="amzLabelPrint" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="Style/amzLabelPrint.css" rel="stylesheet" type="text/css" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Amazon - Stampa Etichette</title>
    <script type="text/javascript">
        function makeSign(chkB) {
            var mycell = chkB.parentElement;
            if (mycell.style.background == "#FFFFFF" || mycell.style.background == "" ||  mycell.style.background == "rgb(255, 255, 255)")
                mycell.style.background = "#808080";
            else
                mycell.style.background = "#FFFFFF";
        }

        function makeGrey(chkB) {
            var mycell = chkB.parentElement;
            var table = document.getElementById("tabPaper");
            for (var i = 1, row; row = table.rows[i]; i++) {
                for (var j = 0, cell; cell = row.cells[j]; j++)
                    cell.style.background = "#FFFFFF";
            }
            mycell.style.background = "#808080";
        }

        function checkClick() {
            var c = 0;
            var table = document.getElementById("tabPaper");
            for (var i = 2, row; row = table.rows[i]; i++) {
                for (var j = 1, cell; cell = row.cells[j]; j++)
                    if (cell.childNodes[0].checked)
                        c++;
            }
            if (c == '<%=numAddr%>')
                return (true);
            else {
                alert('Devi selezionare ' + '<%=numAddr%>' + ' etichette sul foglio!');
                return (false);
            }
        }

        function checkNum() {
            var c = 0;
            var table = document.getElementById("tabPaper");
            for (var i = 2, row; row = table.rows[i]; i++) {
                for (var j = 1, cell; cell = row.cells[j]; j++)
                    if (cell.childNodes[0].checked)
                        c++;
            }
            if (c == '<%=numAddr%>')
                return (true);
            else {
                alert('Devi selezionare ' + '<%=numAddr%>' + ' etichette sul foglio!');
                return (false);
            }
        }

        function checkPrefix() {
            var txt = document.getElementById("txDownloadList");
            if (txt.value.trim() != "" && !isNaN(parseInt(txt.value.substring(txt.value.length - <%=VARNUM%>))))
                window.open("download.aspx?bprefix=" + txt.value.substring(0, txt.value.length - <%=VARNUM%>) + "&start=" + txt.value.substring(txt.value.length - <%=VARNUM %>) + "&vettID=<%=posteID%>", '_blank');
            else
                alert("Prefisso errato!");
        }
    </script>
    <style type="text/css">
        /*.numberCircle {
            width: 120px;
            line-height: 40px;
            border-radius: 50%;
            text-align: center;
            font-size: 16px;
            border: 2px solid #666;
            color: green;
        }*/
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center; font-family: Cambria;">
        <asp:Table ID="tabInfo" runat="server"  HorizontalAlign="Center" CellPadding="4" BorderWidth="1" Width="1000px">
            <asp:TableRow>
                <asp:TableCell ColumnSpan="1" >
                    <h2>Amazon - Stampa Etichette - <%=COUNTRY %></h2></asp:TableCell>
                <asp:TableCell Font-Size="Small">
                    <b><%=Account %>&nbsp;-&nbsp;<%=TipoAccount %>&nbsp;-&nbsp;
                    <asp:DropDownList ID="dropTypeOper" runat="server" Visible="false" AutoPostBack="true" Font-Size="X-Small" ></asp:DropDownList></b>
                    <br /><br />
                    <asp:Button ID="btnLogOut" runat="server" Text="LogOut" OnClick="btnLogOut_Click" Font-Size="Small" Width="70px" />
                    &nbsp;-&nbsp;
                    <asp:Label runat="server" ID="labGoLav" Text="Vai a lavorazioni" Font-Bold="true" Font-Size="Small"></asp:Label><br />
                    <br /><br />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2">
                    <h2>Amazon - Stampa Etichette</h2>
                    <hr />
                    <asp:Image ID="imgTopLogo" runat="server" ImageUrl="~/pics/mcs-logo.png" /><br /><hr />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Right">
                    <asp:Label ID="labInfoBollino" runat="server" Visible="false" Text="" Font-Bold="true"></asp:Label><br />
                    <asp:Label ID="labDownloadList" runat="server" Visible="false" Text="Numero primo bollino"></asp:Label>
                    &nbsp;&nbsp;
                    <asp:TextBox ID="txDownloadList" runat="server" Visible="false" Width="200"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Left">
                    <br />
                    <asp:HyperLink ID="hylDownloadList" runat="server" Visible="false" Text="Scarica lista" Font-Bold="true" NavigateUrl="javascript:checkPrefix();"></asp:HyperLink>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow BorderWidth="1">
                <asp:TableCell ColumnSpan ="2">
                    <asp:Label ID="labOrderID" runat="server" Text="Ordine n#: " Font-Bold="true" Font-Size="Medium"></asp:Label></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2"><hr /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow BorderWidth="1" HorizontalAlign="center">
                <asp:TableCell Font-Bold="true"><asp:Label ID="labDest" runat="server"></asp:Label><br />
                    da stampare: <%=numAddr%>
                </asp:TableCell>
                <asp:TableCell ColumnSpan ="1" HorizontalAlign="Left">
                    <asp:Label ID="labAddress" runat ="server" Font-Size="Medium" Font-Bold="true" ></asp:Label>
                    <asp:Table ID="tabAddr" runat="server" >

                    </asp:Table>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2"><hr /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow BorderWidth="1">
                <asp:TableCell HorizontalAlign="Center">
                    <asp:Table ID="tabPaperSize" runat="server">
                        <asp:TableRow>
                            <asp:TableCell>Tipo etichetta:<br /><br /></asp:TableCell>
                            <asp:TableCell><asp:DropDownList ID="dropLabels" runat="server" AutoPostBack="true" OnSelectedIndexChanged="dropLabels_SelectedIndexChanged"></asp:DropDownList><br /><br /></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell HorizontalAlign="Right">Etichetta larghezza:</asp:TableCell>
                            <asp:TableCell HorizontalAlign="Left"><asp:TextBox ID="txLabW" runat="server" Width="60" Enabled="false"></asp:TextBox>mm<br /></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell HorizontalAlign="Right"><br />Etichetta altezza:</asp:TableCell>
                            <asp:TableCell HorizontalAlign="Left"><br /><asp:TextBox ID="txLabH" runat="server" Width="60" Enabled="false"></asp:TextBox>mm</asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell HorizontalAlign="Right"><br />Margine alto:<br /></asp:TableCell>
                            <asp:TableCell HorizontalAlign="Left"><br /><asp:TextBox ID="txMarginTop" runat="server" Width="60" Enabled="false"></asp:TextBox>mm</asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell HorizontalAlign="Right"><br />Margine sinistro:<br /></asp:TableCell>
                            <asp:TableCell HorizontalAlign="Left"><br /><asp:TextBox ID="txMarginLeft" runat="server" Width="60" Enabled="false"></asp:TextBox>mm</asp:TableCell>                            
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell HorizontalAlign="Right"><br />Etichette per colonna:<br /></asp:TableCell>
                            <asp:TableCell HorizontalAlign="Left"><br /><asp:TextBox ID="txLabRiga" runat="server" Width="60" Enabled="false"></asp:TextBox></asp:TableCell>                            
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell HorizontalAlign="Right"><br />Etichette per riga:<br /></asp:TableCell>
                            <asp:TableCell HorizontalAlign="Left"><br /><asp:TextBox ID="txLabColonna" runat="server" Width="60" Enabled="false"></asp:TextBox></asp:TableCell>                            
                        </asp:TableRow> 
                        <asp:TableRow>
                            <asp:TableCell HorizontalAlign="Right"><br />Margine tra righe:<br /></asp:TableCell>
                            <asp:TableCell HorizontalAlign="Left"><br /><asp:TextBox ID="txMargInfraRighe" runat="server" Width="60" Enabled="false"></asp:TextBox>mm</asp:TableCell>                            
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell HorizontalAlign="Right"><br />MArgine tra colonne:<br /></asp:TableCell>
                            <asp:TableCell HorizontalAlign="Left"><br /><asp:TextBox ID="txMargInfraCol" runat="server" Width="60" Enabled="false"></asp:TextBox>mm</asp:TableCell>                            
                        </asp:TableRow>
                    </asp:Table>
                </asp:TableCell>
                <asp:TableCell ColumnSpan="1" HorizontalAlign="Left">
                    <asp:Table runat="server" ID="tabPaper" CellPadding="8">
                       
                    </asp:Table>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow BorderWidth="1">
                <asp:TableCell ColumnSpan="2"><br /><br /><asp:CheckBox ID="chkHorizontal" runat="server" Text="Stampa orizzontale" Checked="false" Enabled="false" /></asp:TableCell>
            </asp:TableRow>
        </asp:Table>

        <br />
        <asp:CheckBox ID="chkSetShipped" runat="server" Text="Segna lavorazioni come spedite" Checked="true" />&nbsp;&nbsp;&nbsp;
        <asp:CheckBox ID="chkSetInTime" runat="server" Text="Rimuovi dai ritadi" Checked="true" />
        <br />
        <br />
        <asp:Button ID="btnPrint" Text="Crea PDF!" runat="server" OnClick="btnPrint_Click" OnClientClick="return (checkClick());" /><br /><br />
        <asp:HyperLink ID="hypHome" Text="home" runat="server" Visible="true" Font-Bold="true" Font-Size="Medium" ></asp:HyperLink>
    </div>
    </form>
</body>
</html>

 