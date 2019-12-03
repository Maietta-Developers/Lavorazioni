<%@ Page Language="C#" AutoEventWireup="true" CodeFile="lavDettaglio.aspx.cs" Inherits="lavDettaglio" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <link href="Style/dettaglio.css" rel="stylesheet" type="text/css" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Lavorazioni - Dettaglio</title>
    
    <script src="js/jquery-1.6.2.min.js" type="text/javascript"></script>
    <script src="js/jquery.magnifier.js" type="text/javascript" ></script>
    <script src="js/jquery.dynDateTime.min.js" type="text/javascript"></script>
    <script src="js/calendar-en.min.js" type="text/javascript"></script>
    
    <link href="Style/calendar-blue.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="js/tinymce/tinymce.js"></script>
    
    <script>
        jQuery(document).ready(function(){
            jQuery("#hide").click(function () {
                jQuery("#divListaStati").slideUp("slow");
                document.getElementById("show").style.visibility = 'visible';
                document.getElementById("hide").style.visibility = 'hidden';
                
            });
        });
        jQuery(document).ready(function(){
            jQuery("#show").click(function () {
                jQuery("#divListaStati").slideDown("slow");
                document.getElementById("show").style.visibility = 'hidden';
                document.getElementById("hide").style.visibility = 'visible';
                jQuery('html, body').animate({ scrollTop: jQuery("#down").offset().top }, 'slow');
            });
        });        
    </script>
    <script type="text/javascript">
        jQuery(document).ready(function () {
            jQuery("#<%=txDatetime.ClientID %>").dynDateTime({
                showsTime: true,
                //ifFormat: "%Y/%m/%d %H:%M",
                ifFormat: "%d/%m/%Y",
                daFormat: "%l;%M %p, %e %m, %Y",
                align: "BR",
                electric: false,
                singleClick: false,
                displayArea: ".siblings('.dtcDisplayArea')",
                button: ".next()"
            });
        });
    </script>
    <script type="text/javascript">
        tinymce.init({
            mode: "textareas", selector: "#txDescrizione", resize: false, width: 670, statusbar: false, language: 'it', menubar: false, toolbar: false, readonly: 1,
            auto_focus: true, content_css: "Style/dettaglio.css"
        });
    </script>
    <script type="text/javascript">
        tinyMCE.init({
            mode: "textareas", selector: "#txAddDescr", menubar: false, resize: false, width: 670, statusbar: false, language: 'it', 
        });
    </script>
    <script type="text/javascript">
        window.onload = function () {
            //document.getElementById("divListaStati").style.visibility = 'hidden';
            jQuery("#divListaStati").slideUp("slow");
            document.getElementById("show").style.visibility = 'visible';
            document.getElementById("hide").style.visibility = 'hidden';
        }

        function btnVisible() {
            if (document.getElementById("divListaStati").style.visibility == 'visible')
                document.getElementById("divListaStati").style.visibility = 'hidden';
            else
                document.getElementById("divListaStati").style.visibility = 'visible';
        }

        function checkInteger() {
            var txt = document.forms[0].txGiorniLav.value;
            if (txt == "" || isNaN(txt)) {
                return false;
            } else {
                return true;
            }
        }

        function chkSePostback(chk){
            if (chk.childNodes[0].checked)
                return false;
            else {
                //document.forms['form1'].submit();
                return true;
            }
        }

        function checkAttesa() {
            var table = document.getElementById("gridProdotti");
            if (document.getElementById("dropStato").selectedOptions[0].value != '<%= LAVSTATORIC %>'){
                return true;
            }
            for (var i = 1, row; row = table.rows[i]; i++) {
                if (row.cells[0].childNodes[1].childNodes[0].checked) {
                //(row.cells[0].childNodes[0].checked){
                    return true;
                }
            }
            alert("Impossibile eseguire senza aver selezionato il prodotto mancante.");
            return false;
        }
        function confirmNoAttach() {
            var ok = confirm("Desideri che la lavorazione NON richieda allegati?");
            if (ok)
                return true;
            else
                return false;
        }
        function confirmResetAttach() {
            var ok = confirm("Desideri che la lavorazione RICHIEDA allegati?");
            if (ok)
                return true;
            else
                return false;
        }

        

    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center; font-family: Cambria;">
    <asp:Table ID="LavTab" runat="server" HorizontalAlign="Center" CellPadding="0" >
        <asp:TableRow>
            <asp:TableCell ColumnSpan="3">
                <asp:Table ID="Table1" runat="server" Width="100%">
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="2" Font-Size="Small" Width="40%" HorizontalAlign="Left" Font-Bold="true" >
                            <h1>Lavorazione - <%= LAVID %></h1>
                            <br /><b><%=COUNTRY %></b>
                        </asp:TableCell>
                        <asp:TableCell ColumnSpan="2" Font-Size="Small"><b>
                            <%=Account %>&nbsp;-&nbsp;<%=TipoAccount %>&nbsp;&nbsp;-&nbsp;&nbsp;<asp:DropDownList ID="dropTypeOper" runat="server" Visible="false" AutoPostBack="true" Font-Size="X-Small" 
                                OnSelectedIndexChanged="dropTypeOper_SelectedIndexChanged"></asp:DropDownList></b>
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
                            <asp:hyperlink id="hylBrowseRiv" runat="server" Target="_blank">
                                    <asp:image id="imgBrowseRiv" runat="server" imageurl="~/pics/browse.png"  />
                            </asp:hyperlink>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Font-Size="Small" HorizontalAlign="Right">
                            Cliente:&nbsp;&nbsp;
                        </asp:TableCell>
                        <asp:TableCell  HorizontalAlign="Left">
                            <asp:Label ID="labCliente" runat="server" Font-Bold="true"></asp:Label>
                            <asp:hyperlink id="hylBrowseCli" runat="server" Target="_blank">
                                    <asp:image id="imgBrowseCli" runat="server" imageurl="~/pics/browse.png"  />
                            </asp:hyperlink>
                            
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell HorizontalAlign="Right"><asp:Label ID="labPdf" Font-Size="Small" runat="server">Info:&nbsp;&nbsp;</asp:Label></asp:TableCell>
                        <asp:TableCell><asp:Label ID="labLinkPdf" runat="server" Font-Bold="true" Visible="false" Text="Ciao"></asp:Label></asp:TableCell>
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
            <asp:TableCell Font-Size="Small" ColumnSpan="3" Width="100%">
                <hr /><br />
                <asp:GridView ID="gridProdotti" runat="server" OnRowDataBound="gridProdotti_RowDataBound" Width="100%" AutoGenerateColumns="False">
                    <Columns>
                        <asp:TemplateField HeaderText="Sel." ShowHeader="false">
                            <ItemTemplate>
                            <asp:CheckBox ID="chk" runat="server" Checked='<%# Convert.ToBoolean(Eval("Sel")) %>' 
                                OnCheckedChanged="cb_CheckedChanged" onchange="return chkSePostback(this);" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Img." HeaderText="Img."  Visible="True" />
                        <asp:BoundField DataField="Codice" HeaderText="Codice"  Visible="True" />
                        <asp:BoundField DataField="Qt." HeaderText="Qt."  Visible="True" />
                        <asp:BoundField DataField="Info" HeaderText="Info"  Visible="True" />
                        <asp:BoundField DataField="Prz." HeaderText="Prz."  Visible="True" />
                        <asp:BoundField DataField="Descrizione" HeaderText="Descrizione"  Visible="True" />
                        <asp:BoundField DataField="Disp." HeaderText="Disp."  Visible="True" />
                        <asp:BoundField DataField="IDP" HeaderText="IDP"  Visible="true" />
                        <asp:BoundField DataField="LocalZ" HeaderText="Maps"  Visible="true" />
                    </Columns>
                </asp:GridView>
                <br />
                
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left" ColumnSpan="3" Font-Size="Small" Font-Bold="true">
                <asp:HyperLink ID="hylMapsAll" runat="server" Visible="false" ><asp:Image ID="imgMapAll" runat="server" ImageUrl="~/pics/maps-pin.png" Width="20" Height="20" />&nbsp;Localizza tutti</asp:HyperLink>
                <br />
                <hr /><br />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell ColumnSpan="2">
                <asp:Table runat="server" CellPadding="5" >
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell RowSpan="4" Font-Size="Small" Font-Bold="true" ><asp:Label ID="labDescDownload" runat="server"></asp:Label><br />
                            <asp:TextBox ID="txDescrizione" Width="670" Height="200" TextMode="MultiLine" runat="server" ReadOnly="true"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell Font-Size="Small" Font-Bold="true" VerticalAlign="Top">Operatore<br />
                            <asp:DropDownList ID="dropOperatore" runat="server" Width="200"></asp:DropDownList></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left"><asp:TableCell Font-Size="Small" Font-Bold="true" VerticalAlign="Top">Obiettivo<br />
                            <asp:DropDownList ID="dropObiettivo" Width="200" runat="server"></asp:DropDownList></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell Font-Size="Small" Font-Bold="true" VerticalAlign="Top">Tipo di stampa<br />
                            <asp:DropDownList ID="dropTipoStampa" Width="200" runat="server"></asp:DropDownList></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ID="tdMacchina" Font-Size="Small" Font-Bold="true" VerticalAlign="Top" >Macchina<br />
                            <asp:DropDownList ID="dropMacchina" Width="200" runat="server"></asp:DropDownList></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow HorizontalAlign="Left">
                        <asp:TableCell Font-Size="Small" Font-Bold="true" RowSpan="2" >
                            Aggiungi descrizione<br />
                            <asp:TextBox ID="txAddDescr" Width="670" Height="100" TextMode="MultiLine" runat="server" 
                                MaxLength="<%# LavClass.SchedaLavoro.DESCR_MAX %>"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell ID="tdPriorita" Font-Size="Small" Font-Bold="true" VerticalAlign="Top">Priorit&agrave;<br />
                            <asp:DropDownList ID="dropPriorita" Width="200" runat="server"></asp:DropDownList></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell>
                            <asp:Button ID="btnUpdateScheda" runat="server" Font-Bold="true" Text="Aggiorna" OnClick="btnUpdateScheda_Click" Width="90%" Height="30px" />
                            <hr />
                            <asp:RadioButton ID="rdbUpdateScheda" runat="server" Text="Scheda" GroupName="rdgUpdate" Font-Size="Small" Font-Bold="true" Checked="true" />&nbsp;&nbsp;
                            <asp:RadioButton ID="rdbUpdateStatoScheda" runat="server" Text="Stato & Scheda" GroupName="rdgUpdate" Font-Size="Small" Font-Bold="true"/>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Font-Size="Small" Font-Bold="true" ColumnSpan="2" >
                            Note&nbsp;&nbsp;
                            <asp:TextBox ID="txNote" runat="server" TextMode="SingleLine" Width="85%" MaxLength="<%# LavClass.SchedaLavoro.NOTE_MAX %>"></asp:TextBox>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow ID="rwInfoPb" runat="server">
            <asp:TableCell ColumnSpan="3">
                <hr /><br />
                <asp:Label ID="labInfoPostb" runat="server" Text="prova1" ForeColor="Red" Font-Bold="true"></asp:Label><br /><br />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow runat="server" ID="rowInfo" Visible="false">
            <asp:TableCell ColumnSpan="3" HorizontalAlign="Center" Font-Italic="true" Font-Size="Smaller" Font-Bold="true">
                <asp:Label ID="labInfoApprova" runat="server"></asp:Label></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell Font-Size="Small" Width="50%" HorizontalAlign="Center" BackColor="LightGray" CssClass="cellDownGrid" >
                <asp:RadioButton ID="rdbUpComuni" runat="server" GroupName="grpUpload" Text="Upload File Comuni" />
                &nbsp;&nbsp;&nbsp;<asp:RadioButton ID="rdbUpLavorazione" runat="server" GroupName="grpUpload" Text ="Upload File Lavorazione" Checked="true" />
                <br />
                <br /><asp:FileUpload ID="fupCarica" runat="server" AllowMultiple="true" />
                <br /><br />
                <asp:Button ID="btnEmptyAllegato" runat="server" Width="100" OnClick="btnUpload_Click"  />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="btnUpload" Text ="Carica" OnClick="btnUpload_Click" Width="200px" runat="server"/>
                <br /><asp:Label ID="Span1" runat="server"></asp:Label>
                <hr />
                <h3>ALLEGATI:</h3><hr />
                <asp:GridView ID="gridAllegatiComuni" runat="server" OnRowDataBound="gridAllegatiCumuni_RowDataBound" Width="400px" HorizontalAlign="Center"></asp:GridView>
                <br /><hr /><br />
                <asp:GridView ID="gridAllegati" runat="server" OnRowDataBound="gridAllegati_RowDataBound" Width="400px" HorizontalAlign="Center"></asp:GridView>
                <br />
                <asp:Table ID="tabDownload" runat="server" HorizontalAlign="Center" CellPadding="10" Width="100%">
                    <asp:TableRow>
                        <asp:TableCell HorizontalAlign="Center" Font-Bold="true" Font-Size="Smaller">
                            <asp:ImageButton Width="40" Height="40" ImageUrl="pics/zip2.png" ID="btnSendAttachLav" runat="server" AlternateText="Scarica selezionati" OnClick="btnSendAttachLav_Click" />
                            <br /><asp:Label ID="labLinkDSel" runat="server" Text="Scarica Selezionati"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell  Font-Bold="true" Font-Size="Smaller" HorizontalAlign="Center">
                            <asp:HyperLink ID="hylDownZipLav" ImageUrl="pics/zip.png" runat="server" ImageHeight="40" ImageWidth="40"></asp:HyperLink>
                            <br /><asp:Label ID="labLinkDLav" runat="server" Text="Scarica Lavorazione"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell  Font-Bold="true" Font-Size="Smaller" HorizontalAlign="Center">
                            <asp:HyperLink ID="hylDownZipCom" ImageUrl="pics/zip.png" runat="server" ImageHeight="40" ImageWidth="40"></asp:HyperLink>
                            <br /><asp:Label ID="labLinkDCom" runat="server" Text="Scarica Comuni"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:TableCell>
            <asp:TableCell BackColor="LightGray" Width="50%" VerticalAlign="Top" CssClass="cellDownGrid" >
                <asp:Table runat="server" HorizontalAlign="Center" Width="100%" ID="tabStatus">
                    <asp:TableRow>
                        <asp:TableCell>
                            <h3><asp:Label ID="labCurrentStatus" runat="server"></asp:Label></h3>
                            <asp:Label ID="labCurrentStData" runat="server" Font-Size="Smaller" Font-Bold="true"></asp:Label><br />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
                <br />
                <hr /><br />
                <asp:DropDownList ID="dropStato" runat="server"></asp:DropDownList><br /><br />
                <asp:Button ID="btnAddStorico" runat="server" text="Aggiungi Storico" OnClick="btnAddStorico_Click" OnClientClick="return checkAttesa();" />
                <br /><br /><hr /><br />
                <asp:Label ID="labGiorniLav" runat="server" Text="" Font-Size="Small" Font-Bold="true"></asp:Label><br />
                <asp:TextBox ID="txGiorniLav" Width="50px" runat="server"></asp:TextBox>
                &nbsp;&nbsp;<asp:Button ID="btnSaveGiorniLav" runat="server" OnClick="btnSaveGiorniLav_Click" Text="Salva" OnClientClick="return (checkInteger());" />
                <asp:Panel runat="server" ID="panScaricoMag" Visible="false" HorizontalAlign="Center">
                    <br />
                    <hr />
                    <asp:Table runat="server" ID ="tabScaricoMag" HorizontalAlign="Center">
                        <asp:TableRow>
                            <asp:TableCell>
                                Data ricevuta:</asp:TableCell>
                            <asp:TableCell>
                                Numero ricevuta:</asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell>
                                <asp:TextBox ID="txDatetime" runat="server" ReadOnly="true" Width="90"></asp:TextBox>
                                <img src="pics/calender.png" style="vertical-align: middle;" />
                            </asp:TableCell>
                            <asp:TableCell>
                                <asp:TextBox ID="txNumInvoice" runat="server"  Width="90"></asp:TextBox>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="2">
                                <br />
                                <asp:Label ID="labScaricoMag" runat="server" Font-Bold="true" Font-Size="Small" Text="Crea movimentazione su MAFRA"></asp:Label><br />
                                <asp:Button runat="server" ID="btnScaricoMag" Text="Crea movimentazioni"  OnClientClick="return (checkDispLav());" />
                                &nbsp;&nbsp; <asp:CheckBox runat="server" ID ="chkScaricoMag" Checked="true" Text="Crea ordine MAGA" />
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                </asp:Panel>
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
        <div id="skm_LockPane" class="LockOff"></div> 
        <asp:ScriptManager ID="ScriptMgr" runat="server" EnablePageMethods="true"></asp:ScriptManager>
        <br />
    <div>
        <asp:Button ID="btnMakeOrder" runat="server" Text="Vai a Creazione Ordine"  />
    </div>
        <br />
    <div>
        <asp:Button ID="btnShowListaStati" runat="server" Text="Mostra Stati" OnClientClick="btnVisible(); return(false);" Visible="false" />
        <div id="show" onmouseover="document.getElementById('show').style.cursor = 'pointer';"><b>Mostra stati</b></div>
        <div id="hide" onmouseover="document.getElementById('hide').style.cursor = 'pointer';"><b>Nascondi stati</b></div>
    </div>
        <br />
    <div id="divListaStati" style="text-align: center; font-family: Cambria;">
        <asp:Table ID="tabStati" runat="server" HorizontalAlign="Center" Width="40%" CellSpacing="5">
                
        </asp:Table>
    </div>
    </div>
        <div id="down"></div>
    <script type="text/javascript">

        function checkDispLav()
        {
            var txtDt = document.getElementById('txDatetime');
            PageMethods.checkDispProds("<%=LAVID%>", txtDt.value, onSuccess);
            return (false);
        }

        function onSuccess(result)
        {
            if (result == "")
            {
                skm_LockScreen("Attendi...");
            }
            else
            {
                alert("Prodotto " + result + " non disponibile!");
            }
        }


        function skm_LockScreen(str) {
            var lock = document.getElementById('skm_LockPane');
            var txtDt = document.getElementById('txDatetime');
            var txtInv = document.getElementById('txNumInvoice');
            var chkMaga = document.getElementById('chkScaricoMag');
                

            if (txtDt.value.trim() == "")
            {
                alert("E' necessario scegliere una data ricevuta valida!")
                return (false);
            }
            else if (txtInv.value.trim() == "" || isNaN(parseInt(txtInv.value)))
            {
                alert("E' necessario inserire un numero di ricevuta valido!")
                return (false);
            }

            if (lock)
                lock.className = 'LockOn';

            lock.innerHTML = str;

            PageMethods.ScaricoLavorazione("<%=LAVID%>", "<%=TOKEN%>", "<%=MERCHID%>", txtDt.value, txtInv.value, chkMaga.checked.toString(), onSucceed);
            return (false);
        }

        function onSucceed(result) {
            window.location = result;
            return (false);
        }

        function onError(result) {
        }
    </script>
    </form>
</body>
</html>
