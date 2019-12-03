<%@ Page Language="C#" AutoEventWireup="true" CodeFile="amzPanoramica.aspx.cs" Inherits="amzPanoramica" UICulture="it-IT" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="Style/amzPanoramica.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>
    <title>Amazon - Panoramica</title>
    <script type="text/javascript">

        function changeCarrier(orderid)
        {
            return (confirm("Puoi solo cambiare il corriere memorizzato per l'ordine " + orderid + ", vuoi continuare?"));
        }

        function confermaSend(text) {
            if (confirm("Desideri inviare il messaggio a " + text + " ?"))
                return (true);
            else
                return (false);
        }

        function changeHref(dropChanged) {
            var hyl = "hylSend#" + dropChanged.id.split('#')[1];
            var newValue = dropChanged.options[dropChanged.selectedIndex].value;
            var elem = document.getElementById(hyl);

            /*if (!elem.hasAttribute('tiporisposta')) {
                elem.setAttribute('tiporisposta', newValue);    
            }*/

            $(elem).attr('href', elem.href.replace(/tiporisposta=[^&]+/, 'tiporisposta=' + newValue));
            //elem.href = elem.getAttribute('data-href')+val));
            //return (false);
        }

        function checkResults() {
            if (!isNaN(parseInt(document.getElementById('txResults').value))) {
                return (true);
            }
            else {
                alert("Inserire un numero intero di risultati per pagina.");
                return (false);
            }
        }

        function redirectSku(linkB) {
            var dest = linkB.search.split(',')[0];
            document.location = "addSkuItem.aspx" + dest;
            return (false);
        }

        function addRow(btnRiga, numRows) {
            var table = document.getElementById("tabAmazon");
            var rowIn = btnRiga.parentElement.parentElement.parentElement.rowIndex
            if (table.rows[rowIn + 1].style.display == "table-row") {
                for (var i = rowIn + 1; i<rowIn + 1 + numRows; i++){
                    $(table.rows[i]).hide();
                    $(table.rows[i]).css("display", "none");
                    table.rows[i].style.visibility = 'hidden';
                    btnRiga.childNodes[0].src = 'pics/downarrow.png';
                }
            }
            else if (table.rows[rowIn + 1].style.display == "none") {
                for (var i = rowIn + 1; i < rowIn + 1 + numRows; i++) {
                    $(table.rows[i]).show();
                    $(table.rows[i]).css("display", "table-row");
                    table.rows[i].style.visibility = 'visible';
                    btnRiga.childNodes[0].src = 'pics/uparrow.png';
                }
            }
        }

        function checkClick() {
            var table = document.getElementById("tabAmazon");
            var count = 0;
            for (var i = 1, row, cell; row = table.rows[i]; i++) {
                if (row.cells.length == 11) {
                    cell = row.cells[10];
                    if (cell.childNodes[0].checked) {
                        count++;
                    }
                }
            }
            if (count == 0) {
                alert('Devi selezionare almeno una etichetta da stampare!');
            }
            else if (count > '<%=numAddr%>') {
                if (confirm("Selezionando piu di " + <%=numAddr%> + " etichette NON POTRAI SCEGLIERE LE POSIZIONI. Vuoi continuare?") == true)
                    return (true);
                else
                    return (false);
            }
            else if (count <= '<%=numAddr%>') {
                return (true);
            }
            return (false);
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

        function checkEmail() {
            var buyeremail = document.getElementById("txEmailMkt").value;
            
            var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
            if (!re.test(buyeremail)) {
                alert("Formato email non valido!");
                return (false);
            }
            else
                return true;
        }

        function checkList() {
            var lista = document.getElementById("dropListFiltro");
            if (lista.options[lista.selectedIndex].value.trim() == "")
                return (false);
        }

        function askDelete(numOrd) {
            var risp = confirm("Sei sicuro di voler eliminare le movimentazioni dell'ordine " + numOrd + "?");
            if (risp)
                return (true);
            else
                return (false);
        }

        function checkPrime(chk) {
            var drop = document.getElementById("dropStato");
            //var chkCarr = document.getElementById("rdbDataCarrello");
            //var chkConcl = document.getElementById("rdbDataConcluso");
            var dropDataSearch = document.getElementById("dropDataSearch");
            var dropOrdina = document.getElementById("dropOrdina");

            if (chk.childNodes[0].checked) {
                drop.selectedIndex = '<%=((int)AmazonOrder.OrderStatus.STATO_SPEDIZIONE.SPEDITO)%>';
                dropOrdina.value = '<%=((int)AmazonOrder.Order.OrderComparer.ComparisonType.Data_Concluso).ToString()%>';
                dropDataSearch.value = '<%=((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()%>';
            }
            else {
                drop.selectedIndex = '<%=((int)AmazonOrder.OrderStatus.STATO_SPEDIZIONE.DA_SPEDIRE)%>';
                dropOrdina.value = '<%=((int)AmazonOrder.Order.OrderComparer.ComparisonType.Data_Spedizione).ToString()%>';
                dropDataSearch.value = '<%=((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()%>';
            }
            return (false);
        }

        function enableAll() {
            var sel = document.getElementById("chkMagaEnable").checked;
            var table = document.getElementById("tabAmazon");
            var count = 0;
            for (var i = 1, row, cell; row = table.rows[i]; i++) {
                if (row.cells.length == 11) {
                    cell = row.cells[10];
                    if (sel)
                        cell.childNodes[0].checked = true;
                    else
                        cell.childNodes[0].checked = false;
                }
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
    <div id="wrapper" style="text-align:center; font-family:Cambria; ">
    <asp:Table ID="tabHeader" runat="server" HorizontalAlign="Center" CellPadding="0"  >
        <asp:TableRow>
            <asp:TableCell ColumnSpan="2" >
                <h1>Amazon - Panoramica - <%=COUNTRY %></h1></asp:TableCell>
            <asp:TableCell Font-Size="Small" HorizontalAlign="Center" CssClass="tabControl">
                <b><%=Account %>&nbsp;-&nbsp;<%=TipoAccount %>&nbsp;-&nbsp;
                <asp:DropDownList ID="dropTypeOper" runat="server" Visible="false" AutoPostBack="true" Font-Size="X-Small" 
                    OnSelectedIndexChanged="dropTypeOper_SelectedIndexChanged"></asp:DropDownList></b>
                <br /><br />
                <asp:Button ID="btnLogOut" runat="server" Text="LogOut" OnClick="btnLogOut_Click" Font-Size="Small" Width="70px" />
                &nbsp;-&nbsp;
                <asp:Label runat="server" ID="labGoLav" Text="Vai a lavorazioni" Font-Bold="true" Font-Size="Small"></asp:Label>
                &nbsp;-&nbsp;
                <asp:Label runat="server" ID="labGoDownShip" Text="Vai a Spedizioni" Font-Bold="true" Font-Size="Small" Visible="false"></asp:Label>
                <asp:Label runat="server" ID="labGoShipLog" Text="Vai a Sped.Logistica" Font-Bold="true" Font-Size="Small" Visible="false"></asp:Label>
                &nbsp;-&nbsp;
                <asp:Label runat="server" ID="labGoBarC" Text="Vai a BarCode" Font-Bold="true" Font-Size="Small" Visible="false"></asp:Label>
                <br />
                <asp:Label runat="server" ID="labGoConvert" Text="Vai a Tracking" Font-Bold="true" Font-Size="Small" Visible="false"></asp:Label>
                <br />
                <br />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Center" Width="25%" >
                <asp:Calendar ID="calFrom" runat="server" Width="290" OnDayRender ="cal_DayRender"></asp:Calendar>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Center" Width="25%">
                <asp:Calendar ID="calTo" runat="server" Width="290" OnDayRender="cal_DayRender"></asp:Calendar>
            </asp:TableCell>
            <asp:TableCell RowSpan="2" Width="45%">
                <asp:Table runat="server" Width="100%" CssClass="tabControl" >
                    <asp:TableRow><asp:TableCell ColumnSpan="2"><hr /></asp:TableCell></asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right" Font-Bold="true" Width="40%">Stato:</asp:TableCell>
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
                        <asp:TableCell ><asp:Panel ID="panFindSingleOrder" runat="server" DefaultButton="btnFindSingleOrder"><asp:TextBox ID="txNumOrdine" Width="220" runat="server"></asp:TextBox>
                            &nbsp;&nbsp;<asp:Button ID="btnFindSingleOrder" runat="server" OnClientClick="return (checkOrdNum());" OnClick="btnFindSingleOrder_Click" Text="Cerca" /></asp:Panel></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><b>Ricerca email marketplace:</b></asp:TableCell>
                        <asp:TableCell ><asp:Panel ID="panFindEmail" runat="server" DefaultButton="btnFindOrderMkt"><asp:TextBox ID="txEmailMkt" Width="220" runat="server"></asp:TextBox>
                            &nbsp;&nbsp;<asp:Button ID="btnFindOrderMkt" runat="server" OnClientClick="return (checkEmail());" OnClick="btnFindOrderMkt_Click" Text="Cerca" /></asp:Panel></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><b><%=invPrefix %></b></asp:TableCell>
                        <asp:TableCell ><asp:Panel ID="panInvoice" runat="server" DefaultButton="btnFindInvoice"><asp:TextBox ID="txInvoice" runat="server" Width="220"></asp:TextBox>
                            &nbsp;&nbsp;<asp:Button ID="btnFindInvoice" runat="server"  OnClientClick="return (checkInvoice());" OnClick="btnFindInvoice_Click" Text="Cerca" /></asp:Panel></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><b>Carica da lista:</b></asp:TableCell>
                        <asp:TableCell ><asp:DropDownList ID="dropListFiltro" runat="server"  Width="220"></asp:DropDownList>
                            &nbsp;&nbsp;<asp:Button ID="btnFindOrderList" runat="server" OnClientClick="return (checkList());" OnClick="btnFindOrderList_Click" Text="Carica" /></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell HorizontalAlign="Right"><b>Carica da file:</b></asp:TableCell>
                        <asp:TableCell ><asp:FileUpload ID="fupOrderList" runat="server" Width="220" />
                            &nbsp;&nbsp;<asp:Button ID="btnFindOrderFile" runat="server" OnClientClick="return (checkFile());" OnClick="btnFindOrderFile_Click" Text="Carica" /></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow><asp:TableCell ColumnSpan="2"><hr /></asp:TableCell></asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="2">
                            <asp:Label ID="labFindCode" Text="Cerchi un codice SKU/Maietta?" runat="server"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow><asp:TableCell ColumnSpan="2"><hr /></asp:TableCell></asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell HorizontalAlign="Right">
                            <asp:Label ID="labFilter" Text="Filtra per:" runat="server" Font-Bold="true"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Left">
                            <asp:DropDownList ID="dropDataSearch" runat="server" ></asp:DropDownList>&nbsp;&nbsp;&nbsp;&nbsp;
                            <asp:DropDownList ID ="dropTipoSearch" runat="server"></asp:DropDownList>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="trInvoicePrime" runat="server" Visible="false">
                        <asp:TableCell ColumnSpan="2" >
                            <br />
                            <asp:Label ID="labInvoicePrime" runat="server" Text ="Emetti ricevute" Visible ="false"></asp:Label>&nbsp;
                            <asp:ImageButton ID="imbInvoicePrime" ImageUrl="~/pics/amazon-logo-orange.png" OnClick="imbInvoicePrime_Click" OnClientClick="skm_LockScreen('Attendi...');" Visible="false" runat="server" />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow  ID="trPrintLabels" runat="server">
                        <asp:TableCell HorizontalAlign="Right" VerticalAlign="Bottom">
                            <br />
                            <asp:Button ID ="btnPrintLabels" OnClientClick="return (checkClick());" Text="Stampa etichette" runat="server"  />
                            <br />
                            <br />
                            <asp:Label ID="labVettFiltro" runat="server" Text="Filtra per:" Font-Bold="true"></asp:Label>
                            <br />
                            <br />
                            <span style="text-align:left;"><asp:CheckBox ID="chkSoloReady" runat="server" Checked="false" Visible="false" Text="Solo LAVORO PRONTO" /></span>
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Left" VerticalAlign="Bottom">
                            <br />
                            <asp:DropDownList ID="dropLabels" runat="server" AutoPostBack="true" OnSelectedIndexChanged="dropLabels_SelectedIndexChanged"></asp:DropDownList>
                            <br />
                            <br />
                            <asp:DropDownList ID="dropVettFiltro" runat="server" OnSelectedIndexChanged="dropLabels_SelectedIndexChanged" Width="150"></asp:DropDownList>
                            <br />
                            <br />
                            <asp:CheckBox ID="chkSoloMov" runat="server" Checked="false" Visible="false" Text="Solo ricevute emesse" />
                        </asp:TableCell>
                    </asp:TableRow>
                    
                </asp:Table>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow VerticalAlign="Top">
            <asp:TableCell >
                <asp:CheckBox runat="server" ID="chkForceReload" Checked="false" Text="Aggiorna da Amazon" Font-Bold="true" Font-Size="Small" />
                <br /><br />
                <asp:Button ID="btnApplica" runat="server" Text="Applica Filtri" OnClick="btnApplica_Click" Width="180" Height="40" Font-Size="Small" Font-Bold="true" Font-Names="cambria"  />
            </asp:TableCell>
            <asp:TableCell ID="tdDelay" runat="server" Visible="true">
                <span onclick="setAllList();" style="font-size: 13px; font-weight:bold;">
                <asp:Image runat="server" ID="imgAllDelay" ImageUrl="pics/delayRed.png" Width="25px" Height="25px" CssClass="clockTime" />&nbsp;&nbsp;&nbsp;
                    Aggiungi tutti ai ritardi
                    </span>
                <hr />
                <span onclick="removeAllList();" style="font-size: 13px; font-weight:bold;">
                <asp:Image runat="server" ID="imgAllInTime" ImageUrl="pics/delayGreen.png" Width="25px" Height="25px" CssClass="clockTime" />&nbsp;&nbsp;&nbsp;
                    Rimuovi tutti dai ritardi
                    </span>
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
        <br />
        <hr />
        <br />
        <asp:Table ID="tabAmazon" runat="server" HorizontalAlign="Center">

        </asp:Table>
        <br />
        <hr />
        <br />
        
        <asp:HyperLink runat="server" NavigateUrl="~/amzShowComunicazioni.aspx" Target="_blank" Font-Bold="false" Font-Size="Small" Text="Mostra comunicazioni"></asp:HyperLink>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <asp:LinkButton ID="lkbOrderSospesi" runat="server" Visible="false" Font-Bold="true" Font-Size="Small" OnClick="lkbOrderSospesi_Click"></asp:LinkButton>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="labOrdersCount" runat="server" Visible="false" Font-Bold="true" Font-Size="Small"></asp:Label>
        <br />
        <asp:Label ID="labDoubleBuyer" runat="server" Text=""></asp:Label>
        <br />
        <asp:ImageButton ID="imbPrevPag" runat="server" ImageUrl="~/pics/back.png" Width="132" Height="68" Visible="false" CssClass="ImgMiddle"  />
        &nbsp;&nbsp;
        <asp:Label runat="server" ID="labNowPage" Text="" Font-Bold="true" Font-Size="Medium" CssClass="Rounded" Visible="false"></asp:Label>
        &nbsp;&nbsp;
        <asp:ImageButton ID="imbNextPag" runat="server" ImageUrl="~/pics/next.png" Width="132" Height="68" Visible="false" CssClass="ImgMiddle" />
        <div id="skm_LockPane" class="LockOff"></div> 
        <asp:ScriptManager ID="ScriptMgr" runat="server" EnablePageMethods="true"></asp:ScriptManager>
        <script type="text/javascript">

            function skm_LockScreen(str) {
                var lock = document.getElementById('skm_LockPane');
                if (lock)
                    lock.className = 'LockOn';

                lock.innerHTML = str;
            }

            function openLav(numord, rowInd) {
                PageMethods.OpenLavorazione(numord, rowInd);
            }

            function importOrder(numord, data, span, secondCell) {
                //var res =
                PageMethods.ImportOrder(numord, data);
                //if (res)
                span.childNodes[0].style.display = "none";
                span.parentElement.childNodes[0].childNodes[0].width = span.parentElement.childNodes[0].childNodes[0].height = 35;
                if (parseBool(secondCell))
                    span.parentElement.childNodes[1].childNodes[0].width = span.parentElement.childNodes[1].childNodes[0].height = 35;
            }

            function parseBool(val) { return val === true || val === "true" || val === "True" }
            
            function amzAddList(orderID, insert, listName, btnCall) {
                changeShowList(btnCall, parseBool(insert));
                changeAttrList(orderID, !parseBool(insert), listName, btnCall);
                PageMethods.AddRemoveFromList(insert, orderID, listName);
            }

            function changeShowList(btnCall, insert) {
                var src = btnCall.childNodes[0].src;

                if (parseBool(insert)) {
                    btnCall.childNodes[0].src = src.replace("Green", "Red");
                    
                }
                else {
                    btnCall.childNodes[0].src = src.replace("Red", "Green");
                }
            }

            function changeAttrList(orderID, insert, listName, btnCall) {
                btnCall.attributes[0].value = "amzAddList(\"" + orderID + "\", \"" + insert + "\", \"" + listName + "\", this);";
            }

            function removeAllList() {
                if (!confirm("Vuoi rimuovere dai ritardi tutta la lista?"))
                    return (false);

                var ordNum, btnCall;
                var table = document.getElementById("tabAmazon");
                var count = 0;
                for (var i = 1, row, cell; row = table.rows[i]; i++) {
                    cell = row.cells[0];
                    if (cell.innerText.trim() != "") {
                        ordNum = cell.innerText.trim();
                        btnCall = cell.childNodes[0].childNodes[6]
                        changeShowList(btnCall, false);
                        amzAddList(ordNum, false, "delay", btnCall);
                    }
                }
            }

            function setAllList() {
                if (!confirm("Vuoi aggiungere ai ritardi tutta la lista?"))
                    return (false);
                var ordNum, btnCall;
                var table = document.getElementById("tabAmazon");
                var count = 0;
                for (var i = 1, row, cell; row = table.rows[i]; i++) {
                    cell = row.cells[0];
                    if (cell.innerText.trim() != "") {
                        ordNum = cell.innerText.trim();
                        btnCall = cell.childNodes[0].childNodes[6]
                        changeShowList(btnCall, true);
                        amzAddList(ordNum, true, "delay", btnCall);
                    }
                }
            }

            //function amzAddList(orderID, insert, listName, btnCall) {
            function SetCancel(numord, insert, listname, btnCall) {
                changeShowList(btnCall, parseBool(insert));
                changeAttr(numord, !parseBool(insert), listname, btnCall);
                PageMethods.AddRemoveCanceled(insert, numord);
            }

            function changeAttr(orderID, insert, listname, btnCall) {
                btnCall.attributes[0].value = "SetCancel(\"" + orderID + "\", \"" + insert + "\", \"" + listname + "\", this);";
                    //"amzAddList(\"" + orderID + "\", \"" + insert + "\", \"" + listName + "\", this);";
            }

        </script>
    </div>
    </form>
</body>
</html>
