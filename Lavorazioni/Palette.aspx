<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Palette.aspx.cs" Inherits="Palette" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Lavorazioni - Palette</title>
    
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <link href="Style/palette.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <div id="wrapper" style="text-align:center; font-family:Cambria; ">
            <h2>Palette Stato Lavori</h2>
            <asp:Table runat="server" ID="tabPalette" Width="1000px" HorizontalAlign="Center" CellPadding="8">
                <asp:TableRow Font-Bold="true" Font-Size="Medium" BorderWidth="10" BorderStyle="Inset" BorderColor="Black">
                    <asp:TableCell>stato:
                        <br /><hr />
                    </asp:TableCell>
                    <asp:TableCell>chi inserisce questo stato:
                        <br /><hr />
                    </asp:TableCell>
                    <asp:TableCell>chi vede questo stato:
                        <br /><hr />
                    </asp:TableCell>
                    <asp:TableCell>cosa vedi tu:
                        <br /><hr />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <br />
            <h2>Palette Priorit&agrave; Lavori</h2>
            <asp:Table runat="server" ID="tabPriorita" Width="400px" HorizontalAlign="Center" CellPadding="8">
            
            </asp:Table>
            <br />
            <h2>Stato Stampanti</h2>
            <asp:Table runat="server" ID="tabPrinters" HorizontalAlign="Center"  Width="500px"  CellPadding="8" Font-Size="small">

            </asp:Table>
        </div>
    </form>
</body>
</html>
