<%@ Page Language="C#" AutoEventWireup="true" CodeFile="amzAutoInvoice.aspx.cs" Inherits="amzAutoInvoice" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link href="Style/manualInvoice.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Amazon - Emissione Ricevuta</title>
    <script type="text/javascript">
        function checkForm(buttonID) {
            var myObject;
            if (document.getElementById('txNumOrd').value == "" || isNaN(parseInt(document.getElementById('txInvoiceNum').value))) {
                alert("Campi non corretti!");
                return (false);
            }
            else {
                //document.getElementById('btnMakePdf').disabled = true;
                //return (true);
                skm_LockScreen('Attendi...');
                document.forms[0].submit();
                //window.setTimeout("disableButton('" + window.event.srcElement.id + "')", 0);
            }
        }

        function disableButton(buttonID) {
            document.getElementById(buttonID).disabled = true;
        }

        function skm_LockScreen(str) {
            var lock = document.getElementById('skm_LockPane');
            if (lock)
                lock.className = 'LockOn';

            lock.innerHTML = str;
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
    <div id="wrapper" style="text-align: center; font-family: Cambria;">
    <h2>Amazon - Emissione Ricevuta - <%=COUNTRY %></h2>
    <hr />
    <asp:Table runat="server" ID="mainTab" HorizontalAlign="Center" CellPadding="10" CellSpacing="0">
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3"><asp:Image ID="imgTopLogo" runat="server" /><br /><hr /><br /><br /> </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left" Width="35%"><asp:Label ID="labNumOrd" runat="server" Text="Ordine Nr.:"></asp:Label></asp:TableCell>
            <asp:TableCell ColumnSpan="1" HorizontalAlign="Left"><asp:TextBox ID="txNumOrd" runat="server" Width="200" Enabled="false" ></asp:TextBox></asp:TableCell>
            <asp:TableCell RowSpan="5" HorizontalAlign="Right" Width="55%">
                <asp:Calendar ID="calDataInvoice" runat="server" OnSelectionChanged="calDataInvoice_SelectionChanged"  SelectionMode="Day" OnDayRender="calDataInvoice_DayRender"></asp:Calendar>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left"><asp:Label ID="labInvoiceNum" runat="server" Text="Ricevuta Nr.:"></asp:Label></asp:TableCell>
            <asp:TableCell ColumnSpan="2" HorizontalAlign="Left"><b><%=AmzInvoicePrefix %></b><asp:TextBox ID="txInvoiceNum" runat="server" Width="50" Enabled="false" Text="0"></asp:TextBox></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="labData" runat="server" Text="Data ricevuta:"></asp:Label></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="labDataScelta" runat="server" Font-Bold="true" ></asp:Label></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="labVettore" runat="server" Font-Bold="false" Text="Vettore:"></asp:Label></asp:TableCell>
            <asp:TableCell HorizontalAlign ="Left" >
                <asp:DropDownList ID="dropVettori" runat="server"></asp:DropDownList></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="1" HorizontalAlign="Left">
                <asp:CheckBox ID="chkRegalo" Checked="false" runat="server" Text="Crea ANCHE ricevuta regalo" Font-Bold="true" Font-Size ="Small" /></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:CheckBox ID="chkMakeEcmScheda" Checked="true" runat="server" Text="Crea la scheda su ECM" Font-Bold="true" Font-Size ="Small" /></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="2" HorizontalAlign="Left">
                <asp:CheckBox ID="chkMovimenta" Checked="false" runat="server" Text="Crea ANCHE movimentazioni di magazzino" Font-Bold="true" Font-Size ="Small" /></asp:TableCell>
            <asp:TableCell></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="1" HorizontalAlign="Left">
                <asp:CheckBox ID="chkSendRisp" Checked="true" runat="server" Text="Invia la comunicazione:" Font-Bold="true" Font-Size ="Small" /></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:DropDownList ID="dropRisposte" runat="server" ></asp:DropDownList>
            </asp:TableCell>
            
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell >
                <asp:HiddenField ID="dataInvoiceHidden" runat="server" Value="0" />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Center" ColumnSpan="3">
                <br /><br />
                <asp:Button ID="btnMakePdf"  OnClick="btnMakePdf_Click" OnClientClick="return (checkForm(this));" Text="Crea PDF!" runat="server" Width="150" Font-Size="Medium" Font-Bold="true" Font-Names="Cambria" />
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
    <div id="skm_LockPane" class="LockOff"></div> 
    </div>
    </form>
</body>
</html>
