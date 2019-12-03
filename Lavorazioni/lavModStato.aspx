<%@ Page Language="C#" AutoEventWireup="true" CodeFile="lavModStato.aspx.cs" Inherits="lavModStato" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <script src="js/jquery-1.6.2.min.js" type="text/javascript"></script>
    <script src="js/jquery.dynDateTime.min.js" type="text/javascript"></script>
    <script src="js/calendar-en.min.js" type="text/javascript"></script>
    <link href="Style/calendar-blue.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

    <title>Lavorazioni - Modifica Stato</title>
    <style type="text/css">
        a {
            text-decoration: none;
            color: black;
        }

        #MainTab {
            width: 1300px;
        }

        .Hide
        {
            display: none;
        }
    </style>

    <script type="text/javascript">
        function enableAll() {
            var sel = document.getElementById("gvLav_chkboxSelectAll");
            var table = document.getElementById("gvLav");
            var count = 0;
            for (var i = 1, row, cell; row = table.rows[i]; i++) {
                    cell = row.cells[<%=colChk %>];
                    cell.childNodes[1].checked = sel.checked;
            }
        }

        function checkDate() {
            var datetx = document.getElementById("txDatetime");
            var dropSource = document.getElementById("dropSourceStato");

            if (dropSource.selectedIndex == 0)
                return(false);
            else if (datetx.value.trim() == "")
            {
                alert("Scegliere una data limite.")
                dropSource.selectedIndex = 0;
                return (false);
            }
            {
                document.forms[0].submit();
            }
            return (true);
        }

        function checkState() {
            var dropTarget = document.getElementById("dropTargetStato");
            var dropSource = document.getElementById("dropSourceStato");
            var datetx = document.getElementById("txDatetime");
            var table = document.getElementById("gvLav");

            if (table.rows.lenght < 1 || dropTarget.options[dropTarget.selectedIndex].value == "-1" || 
                dropTarget.options[dropTarget.selectedIndex].value == dropSource.options[dropSource.selectedIndex].value ||
                datetx.value.trim() == "")
                return (false);
            return (true);
        }
    </script>
        
    <script type="text/javascript">
        $(document).ready(function () {
            $("#<%=txDatetime.ClientID %>").dynDateTime({
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
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align:center; font-family:Cambria;">
        <asp:Table HorizontalAlign="Center" runat="server" ID="MainTab"  CellPadding="5" CellSpacing="3" >
            <asp:TableRow >
                <asp:TableCell ColumnSpan ="4"><h2>Lavorazioni - Modifica Stato</h2> </asp:TableCell>
                <asp:TableCell Font-Size="Small">
                    <b><%=Account %>&nbsp;-&nbsp;<%=TipoAccount %>
                        <br /><br />
                        <asp:Label runat="server" ID="labGoHome" Text="Vai a lavorazioni" Font-Bold="true" Font-Size="Small"></asp:Label><br />
                    <br />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="5">
                    <hr />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    mostra fino alla data:<br />
                    <asp:TextBox ID="txDatetime" runat="server" ReadOnly = "true" Width="90"></asp:TextBox>
                    <img src="pics/calender.png" style="vertical-align: middle;" />
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right">Stato da cambiare:</asp:TableCell>
                <asp:TableCell HorizontalAlign="Left">
                    <asp:DropDownList ID="dropSourceStato" runat="server" AutoPostBack="true" onchange="return checkDate();"  OnSelectedIndexChanged="dropSourceStato_SelectedIndexChanged" ></asp:DropDownList>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right">Modifica in:</asp:TableCell>
                <asp:TableCell HorizontalAlign="Left">
                    <asp:DropDownList ID="dropTargetStato" runat="server"></asp:DropDownList>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="5">
                    <asp:GridView ID="gvLav" runat="server" OnRowDataBound="gvLav_RowDataBound" AutoGenerateColumns="false" HeaderStyle-BackColor="LightGray" HeaderStyle-Font-Bold="true"
                         HeaderStyle-Font-Size="Small" Width="100%"  >
                        <Columns>
                            <asp:BoundField DataField="ID" HeaderText="ID" ItemStyle-CssClass="Hide" HeaderStyle-CssClass="Hide" />
                            <asp:BoundField DataField="ID" HeaderText="LinkID" ItemStyle-Font-Bold="true" />
                            <asp:BoundField DataField="Riv.Cod." HeaderText="Riv.Cod." />
                            <asp:BoundField DataField="Rivenditore" HeaderText="Rivenditore" />
                            <asp:BoundField DataField="Cliente" HeaderText="Cliente" />
                            <asp:BoundField DataField="Nome" HeaderText="Nome" />
                            <asp:BoundField DataField="Proprietario" HeaderText="Proprietario" />
                            <asp:BoundField DataField="Ultimo Stato" HeaderText="Ultimo Stato" />
                            <asp:BoundField DataField="USID"  ItemStyle-CssClass="Hide" HeaderStyle-CssClass="Hide" />
                            <asp:TemplateField HeaderText="X" ItemStyle-Wrap="false" ItemStyle-Font-Size="Smaller">
                                <HeaderTemplate>
                                    <asp:CheckBox ID="chkboxSelectAll" runat="server" onclick="enableAll();" />
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkID" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="5">
                    <asp:Button ID="btnChangeStato" runat="server" OnClick="btnChangeStato_Click" OnClientClick="return (checkState());" Text="Cambia stati" Width="200" Visible="false" />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </div>
    </form>
</body>
</html>
