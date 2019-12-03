<%@ Page Language="C#" AutoEventWireup="true" CodeFile="amzConvertTrack.aspx.cs" Inherits="amzConvertTrack" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Amazon - Conversione Tracking</title>
    <style type="text/css">
        a {
            text-decoration: none;
            color: black;
        }

        #trLoad {
            width: 100%;
            border: 0px;

        }
        #trLoad td {
            border: 1px solid lightgray;
        }

        #tabIntest {
            width: 1000px;
        }
    </style>

    <script type="text/javascript">
        function checkFile() {
            var file = document.getElementById("fuTracking").value;
            if (file != "") {
                return (true);
            }
            else
                return (false);

        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center; font-family: Cambria;">
        <asp:Table ID="tabIntest" runat="server" HorizontalAlign="Center">
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2" Width="60%" Font-Bold="true" >
                    <h1>Amazon - Conversione Tracking</h1>
                </asp:TableCell>
                <asp:TableCell Font-Size="Small">
                    <b><%=Account %>&nbsp;-&nbsp;<%=TipoAccount %>&nbsp;-&nbsp;
                    
                    <asp:DropDownList ID="dropTypeOper" runat="server" Visible="false" AutoPostBack="true" Font-Size="X-Small" ></asp:DropDownList></b>
                    <br /><br />
                    <asp:Button ID="btnLogOut" runat="server" Text="LogOut" OnClick="btnLogOut_Click" Font-Size="Small" Width="70px" />
                    &nbsp;-&nbsp;
                    <asp:HyperLink runat="server" ID="hylGoLav" Text="Vai a panoramica" Font-Bold="true" Font-Size="Small" ></asp:HyperLink>
                    <br /><br />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="3">
                    <br />
                    <asp:Image ID="imgTopLogo" runat="server" />
                    <br /><br /><hr /><br />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow ID="trLoad" runat="server">
                <asp:TableCell Width="100%" ColumnSpan="3">
                    <asp:Panel ID="panLoad" runat="server" Width="100%" DefaultButton="btnConvert">
                        <asp:Table ID="tabLoad" runat="server" Width="100%" >
                            <asp:TableRow>
                                <asp:TableCell>Scegli file:</asp:TableCell>
                                <asp:TableCell><asp:FileUpload ID="fuTracking" runat="server" AllowMultiple="false" /></asp:TableCell>
                                <asp:TableCell>Vettore:</asp:TableCell>
                                <asp:TableCell><asp:DropDownList ID="dropVett" runat="server" ></asp:DropDownList></asp:TableCell>
                                <asp:TableCell><asp:Button ID="btnConvert" runat="server" Text="Converti" Width="200" OnClick="btnConvert_Click" OnClientClick="return(checkFile());" /></asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:Panel>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow ID="trBar" runat="server">
                <asp:TableCell ColumnSpan="3">
                    <br /><hr /><br />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </div>
    </form>
</body>
</html>
