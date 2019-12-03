<%@ Page Language="C#" AutoEventWireup="true" CodeFile="amzShipDownload.aspx.cs" Inherits="amzShipDownload" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
     <link href="Style/amzPanoramica.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Amazon - Esporta Spedizioni</title>
    <script type="text/javascript">
        function SelTutti(myCheck) {
            var table = document.getElementById("gvShips");
            if (myCheck.checked) 
                for (var i = 1, row, cell; row = table.rows[i]; i++) {
                    row.cells[0].childNodes[0].checked = true;
                }
            else
                for (var i = 1, row, cell; row = table.rows[i]; i++) {
                    row.cells[0].childNodes[0].checked = false;
                }
        }

        function checkPrime(chk) {
            var drop = document.getElementById("dropStato");
            if (chk.childNodes[0].checked) {
                drop.selectedIndex = 2;
            }
            else {
                drop.selectedIndex = 0;
            }
            return (false);
        }

        function CheckClick() {
            var table = document.getElementById("gvShips");
            for (var i = 1, row, cell; row = table.rows[i]; i++) {
                if (row.cells[0].childNodes[0].checked == true)
                    return (true);
            }
            alert("Nessun ordine selezionato.");
            return (false);
        }

        function remove(rowBtn) {
            var rowIn = rowBtn.parentElement.parentElement.rowIndex;
            return (confirm('Elimina la riga \'' + document.getElementById('gvCsv').rows[rowIn].cells[0].textContent + '\' ?'));
            //if (r == true)
            //    document.getElementById('gvCsv').deleteRow(rowIn);
            //return (false);
        }

        function checkOrdNum() {
            var orderid = document.getElementById("txNumOrdine").value;
            var x = 0;
            //407-5502706-3680330
            if (orderid.length == 19 && orderid.substring(3, 4) == "-" && orderid.substring(11, 12) == "-" &&
                parseInt(orderid.substring(0, 3), x) && parseInt(orderid.substring(4, 11), x) &&
                parseInt(orderid.substring(12, 19), x))
                return (true);
            else {
                var table = document.getElementById("tabAmazon");
                table.innerHTML = "";
                alert("Formato del numero d'ordine non valido!");
                return (false);
            }
        }

        function checkInvoice() {
            var invoice = document.getElementById("txInvoice").value;
            if (!isNaN(parseInt(invoice)) && parseInt(invoice) > 0) {
                return (true);
            }
            else {
                alert("Inserire un numero intero per la ricevuta!");
                return (false);
            }
        }

        function checkList() {
            var lista = document.getElementById("dropListFiltro");
            if (lista.options[lista.selectedIndex].value.trim() == "")
                return (false);
        }

        function checkFile() {
            var file = document.getElementById("fupOrderList");
            if (file != null && file.value != "")
                return (true);
            else {
                alert("Devi scegliere un file da importare!");
                return (false);
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align:center; font-family:Cambria;">
    <asp:Table ID="tabHeader" runat="server" HorizontalAlign="Center" CellPadding="0" >
        <asp:TableRow>
            <asp:TableCell ColumnSpan="2" >
                <h1>Amazon - Esporta Spedizioni - <%=COUNTRY %></h1></asp:TableCell>
            <asp:TableCell Font-Size="Small">
                <b><%=Account %>&nbsp;-&nbsp;<%=TipoAccount %>&nbsp;-&nbsp;
                <asp:DropDownList ID="dropTypeOper" runat="server" Visible="false" AutoPostBack="true" Font-Size="X-Small" 
                    OnSelectedIndexChanged="dropTypeOper_SelectedIndexChanged"></asp:DropDownList></b>
                <br /><br />
                <asp:Button ID="btnLogOut" runat="server" Text="LogOut" OnClick="btnLogOut_Click" Font-Size="Small" Width="70px" />
                &nbsp;-&nbsp;
                <asp:Label runat="server" ID="labGoLav" Text="Vai a lavorazioni" Font-Bold="true" Font-Size="Small"></asp:Label>
                &nbsp;-&nbsp;
                <asp:Label runat="server" ID="labGoPanoramica" Text="Vai a Panoramica" Font-Bold="true" Font-Size="Small" Visible="true"></asp:Label><br />
                <br />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Center" Width="25%">
                <asp:Calendar ID="calFrom" runat="server" Width="290" OnDayRender="cal_DayRender"></asp:Calendar>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Center" Width="25%">
                <asp:Calendar ID="calTo" runat="server" Width="290" OnDayRender="cal_DayRender"></asp:Calendar>
            </asp:TableCell>
            <asp:TableCell RowSpan="2" Width="45%"> 
                <asp:Table ID="Table1" runat="server" Width="100%" CssClass="tabControl" >
                    <asp:TableRow><asp:TableCell ColumnSpan="2"><hr /></asp:TableCell></asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right" Font-Bold="true">Stato:</asp:TableCell>
                            <asp:TableCell><asp:DropDownList ID="dropStato" runat ="server" Width="130"></asp:DropDownList>
                            &nbsp;&nbsp;<asp:CheckBox ID="chkPrime" Text="<img src='pics/prime.png' width='110px' style='vertical-align: middle;' />" runat="server" onchange="return checkPrime(this);" />

                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right" Font-Bold="true">Ordina:</asp:TableCell>
                        <asp:TableCell><asp:DropDownList ID="dropOrdina" runat ="server" Width="130"></asp:DropDownList></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right" Font-Bold="true"><asp:Label ID="labResults" runat="server">Risultati per pagina: </asp:Label></asp:TableCell>
                        <asp:TableCell><asp:DropDownList runat="server" ID="dropResults" Width="130">
                                <asp:ListItem Text="20" Value="20"></asp:ListItem>
                                <asp:ListItem Text="50" Value="50"></asp:ListItem>
                                <asp:ListItem Text="100" Value="100"></asp:ListItem>
                            </asp:DropDownList></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><b>Ricerca num. ordine:</b></asp:TableCell>
                        <asp:TableCell ><asp:Panel ID="panFindSingleOrder" runat="server" DefaultButton="btnFindSingleOrder"><asp:TextBox ID="txNumOrdine"  Width="220" runat="server"></asp:TextBox>
                            &nbsp;&nbsp;<asp:Button ID="btnFindSingleOrder" runat="server" OnClientClick="return (checkOrdNum());" OnClick="btnFindSingleOrder_Click" Text="Cerca" /></asp:Panel></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><b><%=invPrefix %></b></asp:TableCell>
                        <asp:TableCell ><asp:Panel ID="panInvoice" runat="server" DefaultButton="btnFindInvoice"><asp:TextBox ID="txInvoice"  Width="220" runat="server"></asp:TextBox>
                            &nbsp;&nbsp;<asp:Button ID="btnFindInvoice" runat="server" OnClientClick="return (checkInvoice());" OnClick="btnFindInvoice_Click" Text="Cerca" /></asp:Panel></asp:TableCell>
                    </asp:TableRow>
                     <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><b>Carica da lista:</b></asp:TableCell>
                        <asp:TableCell ><asp:DropDownList ID="dropListFiltro" runat="server"  Width="220" ></asp:DropDownList>
                            &nbsp;&nbsp;<asp:Button ID="btnFindOrderList" runat="server" OnClientClick="return (checkList());" OnClick="btnFindOrderList_Click" Text="Carica" /></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><b>Carica da file:</b></asp:TableCell>
                        <asp:TableCell ><asp:FileUpload ID="fupOrderList" runat="server" Width="220" />
                            &nbsp;&nbsp;<asp:Button ID="btnFindOrderFile" runat="server" OnClientClick="return (checkFile());" OnClick="btnFindOrderFile_Click" Text="Carica" /></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><b>Esporta per:</b></asp:TableCell>
                        <asp:TableCell ><asp:DropDownList runat="server" ID="dropVett" Width="130" /></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow><asp:TableCell ColumnSpan="2"><hr /></asp:TableCell></asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell><br />
                            <asp:RadioButton id="rdbDataAcq" Text="Data Acquisto" GroupName="rdgData" runat="server" /><br />
                            <asp:RadioButton id="rdbDataMod" Text="Data Concluso" Checked="true" GroupName="rdgData" runat="server" />
                        </asp:TableCell>
                        <asp:TableCell><br />
                            <asp:RadioButton id="rdbConLav" Text="Solo lavorazione" GroupName="rdgLav" runat="server" /><br />
                            <asp:RadioButton id="rdbSoloPartenza" Text="Solo partenza auto" GroupName="rdgLav" runat="server" /><br />
                            <asp:RadioButton id="rdbTuttiLav" Text="Tutti" Checked="true" GroupName="rdgLav" runat="server" />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><br /><b>Filtra per:</b></asp:TableCell>
                        <asp:TableCell ><br /><asp:DropDownList runat="server" ID="dropVettFiltro" Width="130" /></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell HorizontalAlign="Right" VerticalAlign="Bottom">
                            <br />
                            <asp:CheckBox ID="chkSoloReady" runat="server" Checked="true" Visible="true" Text="Solo LAVORO PRONTO" />
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Left" VerticalAlign="Bottom">
                            <br />
                            <asp:CheckBox ID="chkSoloMov" runat="server" Checked="true" Visible="true" Text="Solo ricevute emesse" />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan ="2">
                <br />
                <asp:CheckBox runat="server" ID="chkForceReload" Checked="false" Text="Aggiorna da Amazon" Font-Bold="true" Font-Size="Small" />
                <br /><br />
                <asp:Button ID="btnShowSped" runat="server" Text="Applica filtri" OnClick="btnShowSped_Click" Width="180" Height="40" Font-Size ="Small" Font-Bold="true" Font-Names="Cambria"/>
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
        <br />
        <hr />
        <br />
        <asp:GridView ID="gvShips" runat="server" HorizontalAlign="Center" AutoGenerateColumns="false" AlternatingRowStyle-CssClass="alternateRow" RowStyle-CssClass="normalRow" CellPadding="6" >

        </asp:GridView>
        <br />
        <asp:ImageButton ID="imbNextPag" runat="server" ImageUrl="~/pics/next.png" Width="132" Height="68" Visible="false" />
        <br />
        <asp:Button ID="btnAddOrderList" runat="server" Text="Aggiungi ordini" OnClientClick="return (CheckClick());" OnClick="btnAddOrderList_Click" Visible="false" />

        <hr />
         <asp:GridView ID="gvCsv" runat="server" AutoGenerateColumns="false" EmptyDataText="Nessun ordine aggiunto." HorizontalAlign="Center" CellPadding="4" OnRowCommand="gvCsv_RowCommand"  >
            
        </asp:GridView>
        <br />
        <asp:CheckBox ID="chkSetShipped" runat="server" Text="Segna lavorazioni come spedite" Checked="true" Visible="false" />&nbsp;&nbsp;&nbsp;
        <asp:CheckBox ID="chkSetInTime" runat="server" Text="Rimuovi dai ritadi" Checked="true" Visible="false" />
        <br />
        <br />
        <asp:Button ID="btnMakeFile" runat="server" Text="Esporta Spedizioni" OnClick="btnMakeFile_Click" Visible="false" />
        <hr />
        <asp:Label ID="labOrdersCount" runat="server" Visible="false" Font-Bold="true" Font-Size="Small"></asp:Label>
    </div>
    </form>
</body>
</html>
