<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AddSkuItem.aspx.cs" Inherits="AddSkuItem" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link href="Style/amzAddSkuItem.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Amazon - <%=OPERAZIONE %> SKU - <%=COUNTRY_TITLE %></title>
    <script type="text/javascript">
        function checkFields() {
            var table = document.getElementById("tabCodes");
            for (var i = 1, row; row = table.rows[i]; i++) {
                if (row.cells[0].childNodes[1].value.substring(0, 2) == "@@")
                    continue;
                if (row.cells[1].childNodes[0].value == "" || row.cells[2].childNodes[0].value == "" || row.cells[4].value == "")
                { return false; }
                else if (row.cells[2].childNodes[0].value == "" || isNaN(row.cells[2].childNodes[0].value))
                { return false; }
                else if (row.cells[4].childNodes[0].value == "" || isNaN(row.cells[4].childNodes[0].value))
                { return false; }
            }

            for (var i = 1, row; row = table.rows[i]; i++) {
                if ((i+1) < table.rows.length && row.cells[0].childNodes[1].value == table.rows[i + 1].cells[0].childNodes[1].value &&
                    row.cells[2].childNodes[0].value != table.rows[i + 1].cells[2].childNodes[0].value)
                {
                    alert("Il tipo di risposta di ogni SKU deve essere lo stesso.");
                    return false;
                }
            }

            var input = document.createElement("input");
            input.setAttribute("type", "hidden");
            input.setAttribute("name", "hidRowsCount");
            input.setAttribute("id", "hidRowsCount");
            input.setAttribute("value", document.getElementById("tabCodes").rows.length.toString());
            document.getElementById("wrapper").appendChild(input);

            //document.getElementById("wrapper").value = table.rows.lenght.tostring();
            return (true);
        }

        function addRow(btnRiga) {
            var table = document.getElementById("tabCodes");
            var rowIn = btnRiga.parentElement.parentElement.rowIndex
            var prevRow = table.rows[rowIn];
            var row = table.insertRow(rowIn + 1);
            row.style.backgroundColor = prevRow.style.backgroundColor;

            var cell1 = row.insertCell(0);
            cell1.innerHTML = prevRow.cells[0].innerHTML;
            cell1.childNodes[0].id = cell1.childNodes[0].id + "#c1";
            cell1.childNodes[0].name = cell1.childNodes[0].name + "#c1";
            cell1.childNodes[1].id = cell1.childNodes[1].id + "#c1";
            cell1.childNodes[1].name = cell1.childNodes[1].name + "#c1";
            
            var cell2 = row.insertCell(1);
            cell2.innerHTML = prevRow.cells[1].innerHTML;
            cell2.childNodes[0].id = cell2.childNodes[0].id + "#c1";
            cell2.childNodes[0].name = cell2.childNodes[0].name + "#c1";

            var cell3 = row.insertCell(2);
            cell3.innerHTML = prevRow.cells[2].innerHTML;
            cell3.childNodes[0].id = cell3.childNodes[0].id + "#c1";
            cell3.childNodes[0].name = cell3.childNodes[0].name + "#c1";

            var cell4 = row.insertCell(3);
            cell4.innerHTML = prevRow.cells[3].innerHTML;
            cell4.childNodes[0].id = cell4.childNodes[0].id + "#c1";
            cell4.childNodes[0].name = cell4.childNodes[0].name + "#c1";

            var cell5 = row.insertCell(4);
            cell5.innerHTML = prevRow.cells[4].innerHTML;
            cell5.childNodes[0].id = cell5.childNodes[0].id + "#c1";
            cell5.childNodes[0].name = cell5.childNodes[0].name + "#c1";

            var cell6 = row.insertCell(5);
            cell6.innerHTML = prevRow.cells[5].innerHTML;
            cell6.childNodes[0].id = cell6.childNodes[0].id + "#c1";
            cell6.childNodes[0].name = cell6.childNodes[0].name + "#c1";

            var cell7 = row.insertCell(6);
            cell7.innerHTML = prevRow.cells[6].innerHTML;
            cell7.childNodes[0].id = cell7.childNodes[0].id + "#c1";
            cell7.childNodes[0].name = cell7.childNodes[0].name + "#c1";

            var cell8 = row.insertCell(7);
            cell8.innerHTML = prevRow.cells[7].innerHTML;
            cell8.childNodes[0].id = cell8.childNodes[0].id + "#c1";
            cell8.childNodes[0].name = cell8.childNodes[0].name + "#c1";
            prevRow.cells[7].innerHTML = "";

            return (false);
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center;">
        <h2>Amazon - SKU <%=OPERAZIONE %></h2>
        <hr />
        <asp:Image ID="imgTopLogo" runat="server" ImageUrl="~/pics/mcs-logo.png" /><br /><hr /><br /><br />
        <asp:Label ID="labOrderID" runat="server" Text="Ordine n#: " Font-Bold="true" Font-Size="Medium"></asp:Label>
        <br /><br />
        <asp:Table runat="server" ID="tabCodes" HorizontalAlign="Center" CellPadding="10">
        </asp:Table>
        <br /><br />
        <asp:Label ID="labRedCode" Text ="" runat="server"></asp:Label>
        <br /><br />
        <asp:Button ID="btnSaveCodes" runat="server" Text="Salva" Width="150" OnClientClick="return checkFields();" OnClick="btnSaveCodes_Click" />
    </div>
    </form>
</body>
</html>
