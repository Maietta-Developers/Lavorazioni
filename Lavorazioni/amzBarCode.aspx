<%@ Page Language="C#" AutoEventWireup="true" CodeFile="amzBarCode.aspx.cs" Inherits="amzBarCode" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>
    <script type="text/javascript" src="js/jquery.magnifier.js"></script>

    <title>Amazon - <%=OPERAZIONE %></title>
    
    <style type="text/css">
        body {
            font-family: Cambria;
        }

        a {
            text-decoration: none;
            color: black;
        }

        .invisible {
            display: none;
        }
    </style>
    
    <script type="text/javascript">
        function checkFields() {
            var shipCode = document.getElementById("txShipCode").value;
            if (shipCode.length != 10 && shipCode.length != 12) {
                alert("Codice non valido!");
                return false;
            }
            else if (shipCode.length > 12) {
                alert("Codice spedizione troppo lungo!");
                return false;
            }
            else {
                return true;
            }
        }

        function checkFile() {
            var model = document.getElementById('fupPackageModel').value;
            if (model == "" || !model.endsWith('.tsv')) {
                alert("File non valido!");
                return (false);
            }
            return (true);
        }

        function changeLink(obj) {
            var hyl = document.getElementById(obj.id.replace("txPrint", "hylPrint"));
            var qt = obj.value;
            var pos = hyl.href.indexOf("&labQt=");
            var link = hyl.href.substr(0, pos);
            if (!isNaN(qt))
                hyl.href = link + "&labQt=" + qt;
            return (false);
        }
            
        function checkLavx() {
            var table = document.getElementById("gridCheckItems");
            var count = 0;
            for (var i = 1, row, cell; row = table.rows[i]; i++) {
                cell = row.cells[11];
                cellMod = row.cells[1];
                
                if (cell.childNodes[1].checked && cellMod.innerText.trim().indexOf("(aggiungi)") != -1) {
                    alert('Hai selezionato un prodotto non ancora associato!');
                    return (false);
                }
                if (cell.childNodes[1].checked) {
                    count++;
                }
            }
            
            if (count == 0) {
                alert('Devi selezionare almeno un prodotto per la lavorazione!');
            }
            else 
                return (true);

            return (false);
        }
    
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center;">
        <asp:Table ID="tabInfo" runat="server"  HorizontalAlign="Center" CellPadding="4"  Width="1200px">
            <asp:TableRow>
                <asp:TableCell ColumnSpan="1" Width="70%" >
                    <h1>Amazon - <%=OPERAZIONE %></h1></asp:TableCell>
                <asp:TableCell Font-Size="Small">
                    <b><%=Account %>&nbsp;-&nbsp;<%=TipoAccount %></b><br /><br />
                    <asp:Button ID="btnLogOut" runat="server" Text="LogOut" OnClick="btnLogOut_Click" Font-Size="Small" Width="70px" />
                    &nbsp;-&nbsp;
                    <asp:Button ID="btnHome" runat="server" Text="Home" OnClick="btnHome_Click" Font-Size="Small" Width="70px"></asp:Button>
                    <br /><br />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <hr />
        <asp:Image ID="imgTopLogo" runat="server" ImageUrl="~/pics/mcs-logo.png" /><br /><hr /><br /><br />
        <asp:Table ID="tabBig" runat="server" Width="1300px" HorizontalAlign="Center">
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left">
                    <asp:Label id="labNumProds" runat="server" Font-Bold="true" Font-Names="Cambria" Font-Size="Medium"></asp:Label>
                    <br />
                    <asp:Label ID="labGoToCsv" runat="server" Font-Bold="true" Font-Names="Cambria" Font-Size="Medium"></asp:Label>
                    <br />
                    <asp:Label ID="labGoToCsvFull" runat="server" Font-Bold="true" Font-Names="Cambria" Font-Size="Medium"></asp:Label>
                    <br />
                    <asp:Label ID="labGoToShipPackaging" runat="server" Font-Bold="true" Font-Names="Cambria" Font-Size="Medium"></asp:Label>&nbsp;&nbsp;&nbsp;
                    <asp:FileUpload ID="fupPackageModel" runat="server" Visible="false" Font-Bold="true" Font-Names="Cambria" />&nbsp;
                    <asp:Button ID="btnPackageUpload" runat="server" Visible="false" Font-Bold="true" Font-Names="Cambria" Text="Carica -> Scarica" OnClick="btnPackageUpload_Click" />
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Center" BorderColor="Black" >
                    <asp:Label ID="labGoToLav" runat="server" Visible="false"></asp:Label>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right">Spedizione ID:</asp:TableCell>
                <asp:TableCell HorizontalAlign="Left"><asp:TextBox ID="txShipCode" runat="server"></asp:TextBox></asp:TableCell>
                <asp:TableCell>
                    <asp:Label id="labLabs" runat="server" Text="Etichette:&nbsp;&nbsp;"></asp:Label>
                    <asp:DropDownList ID="dropLabels" runat="server" AutoPostBack="true" OnSelectedIndexChanged="dropLabels_SelectedIndexChanged"></asp:DropDownList></asp:TableCell>
                <asp:TableCell HorizontalAlign="Left">
                    <asp:Button Text="Carica!" ID="btnFindShips" runat="server" OnClick="btnFindShips_Click" OnClientClick="return checkFields();" Width="100" Font-Bold="true" Height="30" Font-Size="Medium" /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="6">
                    <br /><br />
                    <asp:GridView ID="gridShipItems" runat="server" AutoGenerateColumns="false" Width="100%" CellPadding="5" OnRowDataBound="gridShipItems_RowDataBound" AlternatingRowStyle-BackColor="LightGray">
                       <Columns>
                            <asp:ImageField DataImageUrlField="imageUrl" HeaderText ="Image" ControlStyle-Width="70" ControlStyle-Height = "70" Visible="true" >
                            </asp:ImageField>
                            <asp:BoundField DataField="SellerSKU" HeaderText="SKU"  Visible="True" ControlStyle-Width="80px" ItemStyle-Width="80px" ItemStyle-Wrap="false" />
                            <asp:BoundField DataField="FNSKU" HeaderText="Codice" ItemStyle-Font-Bold="true" ItemStyle-Font-Size="Medium"  Visible="True" />
                            <asp:BoundField DataField="title" HeaderText="Titolo"  Visible="True" ItemStyle-Wrap="true" />
                            <asp:BoundField DataField="codmaie" HeaderText="Cod.Maietta" ItemStyle-Font-Size="Smaller"  ItemStyle-Wrap="false"/>
                            <asp:BoundField DataField="quantita" HeaderText="Spedire" ItemStyle-Font-Bold="true" ItemStyle-Font-Size="Medium"  Visible="True" />
                            <asp:TemplateField  ControlStyle-Width="50px" ItemStyle-Width="100px" ControlStyle-Height="15px" HeaderText="Stampa">
                                <ItemTemplate>
                                    <asp:TextBox id="txPrint" runat="server" onchange="return changeLink(this);"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="x">
                                <ItemTemplate>
                                    <asp:HyperLink ID="hylPrint" runat="server" Target="_blank">
                                        <asp:Image ID="imgHylPrint" ImageUrl="pics/label.png" width="35" height="35" runat="server" />
                                    </asp:HyperLink>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField  ControlStyle-Width="200px" ItemStyle-Width="200px" ControlStyle-Height="15px" HeaderText="Note">
                                <ItemTemplate>
                                    <asp:TextBox id="txNote" runat="server" MaxLength="200" Width="200" ></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                    <asp:GridView ID="gridCheckItems" runat="server" AutoGenerateColumns="false" Width="100%" CellPadding="5" OnRowDataBound="gridCheckItems_RowDataBound" AlternatingRowStyle-BackColor="LightGray">
                       <Columns>
                            <asp:ImageField DataImageUrlField="imageUrl" HeaderText ="Image" ControlStyle-Width="70" ControlStyle-Height = "70" Visible="true" >
                            </asp:ImageField>
                            <asp:BoundField DataField="SellerSKU" HeaderText="SKU"  Visible="True" ControlStyle-Width="80px" ItemStyle-Width="80px" ItemStyle-Wrap="false" />
                            <asp:BoundField DataField="FNSKU" HeaderText="Codice" ItemStyle-Font-Bold="true" ItemStyle-Font-Size="Medium"  Visible="True" />
                            <asp:BoundField DataField="title" HeaderText="Titolo"  Visible="True" ControlStyle-Width="400px" ItemStyle-Width="400px" ItemStyle-Wrap="true" />
                            <asp:BoundField DataField="codmaie" HeaderText="Cod.Maietta" ItemStyle-Font-Size="Smaller"  ItemStyle-Wrap="false"/>
                            <asp:BoundField DataField="dispon" HeaderText="Disp." ItemStyle-Font-Size="Smaller" ItemStyle-Wrap="false" /> 
                            <asp:BoundField DataField="quantita" HeaderText="Spedire" ItemStyle-Font-Bold="true"  ItemStyle-Font-Size="Medium"  Visible="True" ItemStyle-Wrap="false" />
                            <asp:BoundField DataField="IDs" HeaderText="IDS" ItemStyle-Font-Bold="true"  ItemStyle-Font-Size="Medium"
                                ControlStyle-CssClass="invisible" HeaderStyle-CssClass="invisible"  ItemStyle-CssClass="invisible"  />
                           <asp:BoundField DataField="vett_risp" HeaderText="Info" ItemStyle-Font-Size="Smaller"  ItemStyle-Wrap="false"/>
                           <asp:TemplateField HeaderText="Lav." ItemStyle-Wrap="false" ItemStyle-Font-Size="Smaller">
                                <ItemTemplate>
                                    <asp:CheckBox ID='lavChk' runat="server" Checked='<%# Eval("lavChk") %>' Enabled="false"  />
                                </ItemTemplate>
                            </asp:TemplateField>
                           <asp:BoundField DataField="qtSca" HeaderText="Qt.Sca" ItemStyle-Font-Size="Smaller"  ItemStyle-Wrap="false" />
                           <asp:TemplateField HeaderText="X" ItemStyle-Wrap="false" ItemStyle-Font-Size="Smaller">
                                <ItemTemplate>
                                    <asp:CheckBox ID='chkOpenLav' runat="server" Checked='<%# Eval("lavorazione") %>'  />
                                </ItemTemplate>
                            </asp:TemplateField>
                           <asp:TemplateField  ControlStyle-Width="50px" ItemStyle-Width="50px" ControlStyle-Height="15px" HeaderText="Qt.Lav.">
                                <ItemTemplate>
                                    <asp:TextBox id="txQtLav" runat="server" ></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>

                        </Columns>
                    </asp:GridView>
                </asp:TableCell>

            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="6">
                    <br />
                    <asp:Panel ID="panBtn" runat ="server" Visible="false">
                        <asp:HyperLink ID="hylPrintAll" Text="Stampa tutto default" runat="server" Visible ="false"  Target="_blank" Font-Bold="true" />
                        <br /><br />
                        <asp:Button ID="btnSaveNote" Text="Salva tutte le note" runat="server" OnClick="btnSaveNote_Click" Height="30" Width="150" Visible="false"  />
                    </asp:Panel>
                    <asp:Button ID="btnMakeLav" Text="Crea lavorazione" runat="server" OnClick="btnMakeLav_Click" OnClientClick="return checkLavx();" Height="30" Width="150" Visible="false"  />
                </asp:TableCell>
            </asp:TableRow>
            
        </asp:Table>
    </div>
    </form>
</body>
</html>
