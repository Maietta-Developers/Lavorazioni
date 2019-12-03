<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Lavorazioni.aspx.cs" Inherits="Lavorazioni" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Lavorazioni - Panoramica</title>
    
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <link href="Style/DropStyle.css" rel="stylesheet" type="text/css" />
    <link href="Style/panoramica.css" rel="stylesheet" type="text/css" />
    <meta http-equiv="refresh" content="180" />

    <script type="text/javascript">
        window.onload = function () {
            document.getElementById("btnSendAttachLav").style.visibility = 'hidden';
            document.getElementById("labLinkDSel").style.visibility = 'hidden';
        }

        function btnVisible(lavid) {
            var loadLav = document.getElementById("hidLavProd").value;

            if (document.getElementById("btnSendAttachLav").style.visibility == 'hidden') {
                document.getElementById("btnSendAttachLav").style.visibility = 'visible';
                document.getElementById("labLinkDSel").style.visibility = 'visible';
            }
            else {
                if (loadLav == lavid) {
                    document.getElementById("btnSendAttachLav").style.visibility = 'hidden';
                    document.getElementById("labLinkDSel").style.visibility = 'hidden';
                }
            }
        }

        function scrollBottom() {
            window.scrollTo(0, document.body.scrollHeight);
        }

        function scrollTop() {
            window.scrollTo(0, 0);
        }

        function checkInteger(val) {
            //var txt = document.forms[0].txGoToLav.value;
            var txt = document.forms[0][val].value;
            if (txt == "" || isNaN(txt)) {
                alert("Inserisci un numero valido!");
                return false;
            } else {
                window.location.href = "lavDettaglio.aspx?token=" + getParameterByName('token') + "&merchantId=" + getParameterByName('merchantId') + "&idlav=" + txt;
                return true;
            }
        }

        function checkOrdNum() {
            var orderid = document.getElementById("txGoToOrder").value;
            var x = 0;
            //407-5502706-3680330
            if (orderid.length == 19 && orderid.substring(3, 4) == "-" && orderid.substring(11, 12) == "-" &&
                parseInt(orderid.substring(0, 3), x) && parseInt(orderid.substring(4, 11), x) &&
                parseInt(orderid.substring(12, 19), x))
                return (true);
            else {
                alert("Formato del numero d'ordine non valido!");
                return (false);
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server"  ></asp:ScriptManager>
    <div id="wrapper" style="text-align:center; font-family:Cambria;">
        <asp:Table HorizontalAlign="Center" runat="server" ID="MainTab"  CellPadding="5" >
            <asp:TableRow >
                <asp:TableCell HorizontalAlign="Left">
                    <asp:Table runat="server" CellPadding="4">
                        <asp:TableRow>
                            <asp:TableCell Font-Bold="true" Font-Size="Small">
                                <asp:CheckBox ID="chkSoloCommerciale" runat="server" AutoPostBack="true" Text="Solo mie lavorazioni" /></asp:TableCell>
                            <asp:TableCell RowSpan="5" VerticalAlign="Middle">
                                <asp:HyperLink ID="hypRefresh" runat="server" ImageUrl="~/pics/refresh.png" ImageWidth="40px" ></asp:HyperLink>                                
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow  runat="server" ID="trApprovate">
                            <asp:TableCell Font-Bold="true" Font-Size="Small">
                                <asp:CheckBox runat="server" ID="chkSoloApprovate" AutoPostBack="true" Text="Solo approvate" /></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell Font-Bold="true" Font-Size="Small">
                                <asp:CheckBox runat="server" ID="chkSoloInevase" AutoPostBack="true" Text="Solo aperte" Checked="true" Enabled="false" /></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow  runat="server" ID="trStati" >
                            <asp:TableCell Font-Bold="true" Font-Size="Small">
                                <asp:CheckBox ID="chkSoloMieiStati" runat="server" AutoPostBack="true" Text="Solo miei stati" /></asp:TableCell>
                        </asp:TableRow><asp:TableRow>
                            <asp:TableCell Font-Bold="true" Font-Size="Small">
                                <asp:CheckBox ID="chkMostraSospesi" runat="server" AutoPostBack="true" Text="Mostra sospesi" Checked="false" /></asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                </asp:TableCell>
                <asp:TableCell ColumnSpan ="3"><h1>Lavorazioni - Panoramica - <%=COUNTRY %></h1> </asp:TableCell>
                <asp:TableCell Font-Size="Small">
                    <b><%=Account %>&nbsp;-&nbsp;<%=TipoAccount %>&nbsp;-&nbsp;
                    <asp:DropDownList ID="dropTypeOper" runat="server" Visible="false" AutoPostBack="true" Font-Size="X-Small" 
                        OnSelectedIndexChanged="dropTypeOper_SelectedIndexChanged"></asp:DropDownList></b>
                    <br />Ultimo Refresh: <%=DateTime.Now %><br /><br />
                    <asp:Button ID="btnLogOut" runat="server" Text="LogOut" OnClick="btnLogOut_Click" Font-Size="Small" Width="70px" />&nbsp;-&nbsp;
                    <asp:Label runat="server" ID="labGoLav" Text="Vai a Modifica stati" Font-Bold="true" Font-Size="Small"></asp:Label>
                    <br />
                    <br />
                    <asp:Image ID="imgMaps" runat="server" Width="25" Height="25" ImageUrl="~/pics/maps-pin.png" />
                    <asp:HyperLink ID="hylMaps"  Font-Size="Small" Font-Bold="true" runat="server" Text="Localizza prodotto" Target="_self"></asp:HyperLink>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="5">
                    <hr />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow Font-Bold="true" >
                <asp:TableCell ColumnSpan="5">
                    <asp:Table runat="server" Width="100%">
                        <asp:TableRow>
                            <asp:TableCell Font-Size="Small">Operatore:<br />
                                <asp:DropDownList ID="DropOperatoreV" runat="server" AutoPostBack="true" AppendDataBoundItems="true" OnSelectedIndexChanged="OperatoreV_SelectedIndexChanged" CssClass="dropdown1">
                                    <asp:ListItem Text="Tutti" Value="0"></asp:ListItem></asp:DropDownList></asp:TableCell>
                            <asp:TableCell Font-Size="Small">Tipo Lavoro:<br />
                                <asp:DropDownList ID="DropObiettiviV" runat="server" AutoPostBack="true" AppendDataBoundItems="true" CssClass="dropdown1">
                                    <asp:ListItem Text="Tutti" Value="0"></asp:ListItem></asp:DropDownList></asp:TableCell>
                            <asp:TableCell Font-Size="Small">Tipo stampa:<br />
                                <asp:DropDownList ID="DropTipoStampa" runat="server" AutoPostBack="true" AppendDataBoundItems="true" CssClass="dropdown1">
                                    <asp:ListItem Text="Tutte" Value="0"></asp:ListItem></asp:DropDownList></asp:TableCell>
                            <asp:TableCell Font-Size="Small">Macchina:<br />
                                <asp:DropDownList ID="DropMacchina" runat="server" AutoPostBack="true" AppendDataBoundItems="true" CssClass="dropdown1">
                                    <asp:ListItem Text="Tutte" Value="0"></asp:ListItem></asp:DropDownList></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Font-Size="Small">Stato Lavoro:<br />
                            <asp:DropDownList ID="dropStato" runat="server" AppendDataBoundItems="true" AutoPostBack="true" CssClass="dropdown1">
                                <asp:ListItem Text="Tutti" Value="0"></asp:ListItem></asp:DropDownList></asp:TableCell>
                        <asp:TableCell Font-Size="Small">Rivenditore:<br />
                            <asp:DropDownList ID="DropRivenditori" runat="server" AppendDataBoundItems="true" AutoPostBack="true" CssClass="dropdown1">
                                </asp:DropDownList></asp:TableCell>
                        <asp:TableCell Font-Size="Small">Priorit&agrave;:<br />
                            <asp:DropDownList ID="DropPriorita" runat="server" AutoPostBack="true" AppendDataBoundItems="true" CssClass="dropdownWH">
                                <asp:ListItem Text="Tutte" Value="0"></asp:ListItem></asp:DropDownList>
                            </asp:TableCell>
                        <asp:TableCell Font-Size="Small">
                            <br />
                            <asp:CheckBox ID="chkSortDate" runat="server" Text="Ordina per data" AutoPostBack="true" /></asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="5"></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        
        <br />
        <div style="text-align:center;">
            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                <ContentTemplate>
                    <asp:GridView CssClass="Row1" ID="LavGrid" CellPadding="6" runat="server" OnRowDataBound="LavGrid_RowDataBound" HorizontalAlign="Center" Font-Size="Small">
                    </asp:GridView>
                </ContentTemplate>
            </asp:UpdatePanel>
        </div>
        <br />
        <br />
        <div style="text-align:center;">
            <asp:UpdatePanel ID="updPan1" runat="server">
                <ContentTemplate>
                    <asp:Table ID="InfoTab" runat="server" Font-Size="Small" HorizontalAlign="Center" CellPadding="6" Width="1200px">
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="2" Font-Bold="true" Font-Size="Medium">
                                <%=numeroLav %></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell Font-Bold="true" Font-Italic="true">Allegati:</asp:TableCell>
                            <asp:TableCell Font-Bold="true" Font-Italic="true">Prodotti:</asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell BackColor="LightGray" Width="550px" BorderStyle="NotSet">
                                <br />
                                 <asp:GridView ID="GridLavAttachComuni" runat="server" Width="100%" Height="100%" EmptyDataText = "Nessun file caricato." OnRowDataBound="GridLavAttachComuni_RowDataBound" >
                                    </asp:GridView>
                                <hr />
                                <asp:GridView ID="GridLavAttach" runat="server" Width="100%" Height="100%" EmptyDataText = "Nessun file caricato." OnRowDataBound="GridLavAttach_RowDataBound" >
                                    </asp:GridView>
                                <br />
                            </asp:TableCell>
                            <asp:TableCell BackColor="LightGray" Width="550px" BorderStyle="NotSet">
                                <asp:GridView ID="GridLavProds" runat="server" Width="100%" Height="100%" EmptyDataText = "Nessun prodotto associato." OnRowDataBound="GridLavProds_RowDataBound">
                                    </asp:GridView></asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                    <asp:HiddenField ID="hidLavProd" runat="server" />
                </ContentTemplate>
            </asp:UpdatePanel>
            <div id="divBtnAttach">
                <asp:Table runat="server" HorizontalAlign="Center" Width="1200px" BackColor="LightGray">
                    <asp:TableRow>
                        <asp:TableCell Width="50%"  Font-Bold="true" Font-Size="X-Small">
                            <asp:ImageButton ID="btnSendAttachLav" runat="server" ImageUrl="pics/zip2.png" Width="45" Height="45"
                                Text="Scarica selezionati" OnClick="btnSendAttachLav_Click" Visible="true" />
                            <br /><asp:Label ID="labLinkDSel" runat="server" Text="Scarica Selezionati"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </div>
        </div>
        <br />
        <asp:Table ID="tabFooter" runat="server" HorizontalAlign="Center" CellPadding="12" Font-Size="Smaller" >
            <asp:TableRow>
                <asp:TableCell Width="20%" Font-Bold="true" >
                    <asp:Label ID="labLinkPalette" runat="server" > </asp:Label><br /><br />
                    <asp:Label ID="labTotRighe" runat="server" > </asp:Label>
                </asp:TableCell>
                <asp:TableCell Width="20%" Font-Bold="true" >
                    <asp:HyperLink NavigateUrl="javascript:scrollTop();" ID="hylScroll" ImageHeight="35" ImageWidth="35" runat="server" ImageUrl="pics/uparrow.png">
                        </asp:HyperLink>
                    <br />torna su
                </asp:TableCell>
                <asp:TableCell Width="20%">
                    <asp:Panel ID="panGoLav" runat="server" DefaultButton="btnGoToLav">
                        <br /><b>Vai alla lavorazione:</b><br />
                        <asp:TextBox ID="txGoToLav" runat="server" Width="50"></asp:TextBox>&nbsp;&nbsp;
                        <asp:Button ID="btnGoToLav" Text="Vai alla lavorazione" runat="server"  OnClick="btnGoToLav_Click"
                            OnClientClick="return (checkInteger('txGoToLav'));"  />        
                    </asp:Panel>   
                </asp:TableCell>
                <asp:TableCell Width="20%">
                    <asp:Panel ID="panGoOrder" runat="server" DefaultButton="btnGoToOrder">
                        <br /><b>Vai all'ordine Amazon:</b><br />
                        <asp:TextBox ID="txGoToOrder" runat="server" Width="160"></asp:TextBox>&nbsp;&nbsp;
                        <asp:Button ID="btnGoToOrder" Text="Vai" runat="server"  OnClick="btnGoToOrder_Click"
                            OnClientClick="return (checkOrdNum());"  />        
                    </asp:Panel> 
                </asp:TableCell>
                <asp:TableCell Width="20%">
                    <asp:Panel ID="panGoMCS" runat="server" DefaultButton="btnGoToMCS">
                        <br /><b>Vai all'ordine MCS:</b><br />
                        <asp:TextBox ID="txGoToMCS" runat="server" Width="50"></asp:TextBox>&nbsp;&nbsp;
                        <asp:Button ID="btnGoToMCS" Text="Vai" runat="server"  OnClick="btnGoToMCS_Click"
                            OnClientClick="return (checkInteger('txGoToMCS'));"  />        
                    </asp:Panel> 
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="4" >
                    <asp:HyperLink NavigateUrl="~/amzPanoramica.aspx" ID="hylAmazon" ImageHeight="40" ImageWidth="304" runat="server" ImageUrl="~/pics/amazon.png"></asp:HyperLink>
                </asp:TableCell>
                <asp:TableCell ColumnSpan="1"><asp:Label ID="labVersion" runat="server"></asp:Label></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        
    </div>            
    </form>
</body>
</html>

