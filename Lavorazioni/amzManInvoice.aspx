<%@ Page Language="C#" AutoEventWireup="true" CodeFile="amzManInvoice.aspx.cs" Inherits="amzManInvoice" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <link href="Style/freeInvoice.css" rel="stylesheet" type="text/css" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Lavorazioni - Ricevuta Amazon</title>
    <script type="text/javascript">
        function selectDrop() {
            var drop = document.getElementById('dropCodes');
            var dropIn = document.getElementById('dropCodes').selectedIndex;
            var max = document.getElementById('dropCodes').length;
            var txt = document.getElementById('txCodeFind').value;

            var index = 0;
            for (i = 0; i < max; i++) {
                index = dropIn + i + 1;
                if (index >= max)
                    index -= max;
                if (drop.item(index).text.toLowerCase().indexOf(txt.toLowerCase()) != -1) {
                    drop.selectedIndex = index;
                    dropChanges();
                    return (false);
                }
            }
            return (false);
        }

        function dropChanges() {
            var drop = document.getElementById('dropCodes');
            var lab = document.getElementById('labDescSelected');
            lab.textContent = drop.selectedOptions[0].text;
            return (false);
        }

        function checkProd() {
            if (document.getElementById('dropCodes').selectedIndex != 0)
                return (true);
            else
                return (false);
        }

        function makeTotal(btnOk) {
            var rowInd = parseInt(btnOk.split('_')[2]) + 1;
            var row = document.getElementById('gridProducts').rows[rowInd];
            var strCosto = row.cells[1].childNodes[0].value.replace(',', '.');
            var strQt = row.cells[2].childNodes[0].value.replace(',', '.');

            if (isNaN(strCosto))
                row.cells[3].textContent = 'costo non valido';
            else if (isNaN(strQt) || !Number.isInteger(parseFloat(strQt))) {
                row.cells[3].textContent = 'quantità non valida';
            }
            else {
                var a = parseFloat(row.cells[1].childNodes[0].value.replace(',', '.'));
                var c = parseInt(row.cells[2].childNodes[0].value.replace(',', '.'));
                row.cells[3].textContent = (a * c).toFixed(2);
            }
            return (false);
        }

        function removeRow(btnRem) {
            var rowInd = parseInt(btnRem.split('_')[2]) + 1;
            var r = confirm('Elimina la riga \'' + document.getElementById('gridProducts').rows[rowInd].cells[0].textContent + '\' ?');
            if (r == true)
                document.getElementById('gridProducts').deleteRow(rowInd);
            //var row = document.getElementById('gridProducts').rows[rowInd];
            return (false);
        }

        function checkForm() {
            var nomov = document.getElementById("chkMovimenta");
            var comm = document.getElementById("chkSendRisp");
            var risp = document.getElementById("dropRisposte");
            if (comm.checked && risp.selectedOptions[0].value == "0") {
                alert('Comunicazione non valida!');
                return (false);
            }
            else if (document.getElementById('txNumOrd').value == "" || isNaN(parseInt(document.getElementById('txInvoiceNum').value))) {
                alert('Numero di ricevuta non valido!');
                return (false);
            }
            else if (nomov.checked && document.getElementById('gridProducts').rows.length <= 1)
            {
                alert('Nessun prodotto aggiunto!');
                return (false);
            }
            else if (document.getElementById('txNumOrd').value != "" && !isNaN(parseInt(document.getElementById('txInvoiceNum').value)) && nomov.checked &&
                document.getElementById('gridProducts').rows.length > 1)
            {
                //document.form1.action = 'download.aspx';
                //document.form1.method = 'post';
                //document.form1.submit();
                //document.forms[0].method = 'post';
                //document.forms[0].action = 'download.aspx';
                //document.forms[0].submit();
                var table = document.getElementById('gridProducts');
                for (var i = 1, row, cell; row = table.rows[i]; i++) {
                    if (row.cells[3].innerText == '' || isNaN(row.cells[3].innerText)) {
                        alert('Devi confermare ogni riga!');
                        return (false);
                    }
                }
                skm_LockScreen('Attendi...');
                return (true);
            }
            else if (document.getElementById('txNumOrd').value != "" && !isNaN(parseInt(document.getElementById('txInvoiceNum').value)) && !nomov.checked) {
                skm_LockScreen('Attendi...');
                return (true);
            }
            else {
                return (false);
            }

            function skm_LockScreen(str) {
                var lock = document.getElementById('skm_LockPane');
                if (lock)
                    lock.className = 'LockOn';

                lock.innerHTML = str;
            }
        }
    </script>
    <style type="text/css">
        .LockOff { 
            display: none; 
            visibility: hidden; 
        } 

        .LockOn { 
            display: block; 
            visibility: visible; 
            position: absolute; 
            z-index: 999; 
            top: 0px; 
            left: 0px; 
            width: 105%; 
            height: 105%; 
            background-color: #ccc; 
            text-align: center; 
            padding-top: 20%; 
            filter: alpha(opacity=75); 
            opacity: 0.75; 
        } 
    </style>
</head>

<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
    <div id="wrapper" style="text-align: center; font-family: Arial;">
        <h2>Amazon - Ricevuta/Scarico Manuale - <%=COUNTRY %></h2>
        <hr />
        <asp:Table runat="server" ID="mainTab" HorizontalAlign="Center" CellPadding="0">
            <asp:TableRow>
                <asp:TableCell ColumnSpan="3"><asp:Image ID="imgTopLogo" runat="server" /><br /><hr /><br /><br /> </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left" Width="20%"><asp:Label ID="labNumOrd" runat="server" Text="Ordine Nr.:"></asp:Label></asp:TableCell>
                <asp:TableCell ColumnSpan="1" HorizontalAlign="Left"><asp:TextBox ID="txNumOrd" runat="server" Width="200" Enabled="false" ></asp:TextBox></asp:TableCell>
                <asp:TableCell HorizontalAlign="Center" RowSpan="3" Width="55%">
                    <asp:Calendar ID="calInvoiceData" runat="server" OnSelectionChanged="calInvoiceData_SelectionChanged" OnDayRender="calInvoiceData_DayRender" ></asp:Calendar></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left"><asp:Label ID="labInvoiceNum" runat="server" Text="Ricevuta Nr.:"></asp:Label></asp:TableCell>
                <asp:TableCell ColumnSpan="2" HorizontalAlign="Left"><b><%=AmzInvoicePrefix %></b><asp:TextBox ID="txInvoiceNum" runat="server" Width="50" Enabled="false" Text="0">
                    </asp:TextBox></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left">
                    <asp:Label ID="labData" runat="server" Text="Data:"></asp:Label><br /><br />
                    <asp:Label ID="labVettore" runat="server" Text="Vettore"></asp:Label>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Left">
                    <asp:Label ID="labDefinitiveData" runat="server" Font-Bold="true"></asp:Label> <br /><br />
                    <asp:DropDownList ID="dropVettori" runat="server"></asp:DropDownList>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left" ColumnSpan="2">
                    <br />
                    <asp:CheckBox ID="chkRegalo" Checked="false" runat="server" Text="Crea ANCHE ricevuta regalo" Font-Bold="true" Font-Size ="Small" />
                </asp:TableCell>
                <asp:TableCell ></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left" ColumnSpan="2">
                    <br />
                    <asp:CheckBox ID="chkOpenLav" Checked="false" runat="server" Text="Apri ANCHE lavorazione" Font-Bold="true" Font-Size ="Small" />
                </asp:TableCell>
                <asp:TableCell ><br />
                    <asp:CheckBox ID="chkMovimenta" Checked="true" runat="server" Text="Crea ANCHE movimentazioni" Font-Bold="true" Font-Size ="Small" /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="3" HorizontalAlign="Left">
                    <br />
                    <asp:CheckBox ID="chkSendRisp" Checked="true" runat="server" Text="Invia ANCHE la comunicazione:" Font-Bold="true" Font-Size ="Small" />
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:DropDownList ID="dropRisposte" runat="server" ></asp:DropDownList>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="3">
                    <br /><hr /><br />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        
        <asp:UpdatePanel ID="updPan1" runat="server" >
            <ContentTemplate>
                <asp:TextBox ID="txCodeFind" runat="server"></asp:TextBox>
                <asp:Button ID="btnFindCode" runat="server" Text ="Trova" OnClientClick="return(selectDrop());" />
                <asp:DropDownList ID="dropCodes" runat="server" onchange="return(dropChanges());" AutoPostBack="true" 
                    OnDataBound="dropCodes_DataBound" OnDataBinding="dropCodes_DataBinding"></asp:DropDownList>&nbsp;&nbsp;
                <asp:Button ID="btnAddProd" runat="server" OnClick="btnAddProd_Click" Text="Aggiungi" OnClientClick="return(checkProd());" /><br />
                <asp:Label ID="labDescSelected" runat="server" Font-Bold="true" Font-Size="Small"></asp:Label><br /><br />
                <asp:GridView ID="gridProducts" runat="server" AutoGenerateColumns="false" EmptyDataText="Nessun prodotto aggiunto." HorizontalAlign="Center" 
                    OnRowDataBound="gridProducts_RowDataBound" >
                    <Columns>
                        <asp:BoundField HeaderText="Prodotto" DataField="Prodotto" />
                        <asp:BoundField HeaderText="Costo"  />
                        <asp:BoundField HeaderText="Quantità" />
                        <asp:BoundField HeaderText="Totale" />
                        <asp:ButtonField HeaderText="Conferma" />
                        <asp:ButtonField HeaderText="Rimuovi" />
                        <asp:BoundField HeaderText="codprod" DataField="codprod" HeaderStyle-CssClass="cellCode" />
                        <asp:BoundField HeaderText="codforn" DataField="codforn" HeaderStyle-CssClass="cellCode" />
                    </Columns>
                </asp:GridView>
                <br /><br />
            </ContentTemplate>
        </asp:UpdatePanel>
        <span style="font-family: Cambria;"><b>Indicare la quantità da scaricare e il prezzo del singolo pezzo.</b></span><br /><br />
        <asp:Button ID="btnMakePdf" OnClick="btnMakePdf_Click"  OnClientClick="if(!checkForm()) return false;" Text="Crea PDF!" runat="server" Width="150" />
        
        <div id="skm_LockPane" class="LockOff"></div> 
    </div>
    </form>
</body>
</html>
