<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Send.aspx.cs" Inherits="Send" ValidateRequest="false" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Lavorazioni - Invio allegati</title>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <link href="Style/send.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div id="wrapper" style="text-align: center; font-family: Cambria; ">
        <asp:Table ID="tabIntest" runat="server" HorizontalAlign="Center">
            <asp:TableRow>
                <asp:TableCell Font-Size="Small" Width="60%" HorizontalAlign="Left" Font-Bold="true" >
                    <h2>Lavorazione - <%= LAVID %> - Invio allegato</h2>
                    <br />
                    <asp:Image runat="server" Width="24" Height="15" ID="imgChMerch" />&nbsp;&nbsp;&nbsp;
                    <asp:DropDownList runat="server" ID="dropChMerch" AutoPostBack="true" ></asp:DropDownList>
                </asp:TableCell>
                <asp:TableCell Font-Size="Small"><b>
                    <%=Account %>&nbsp;-&nbsp;<%=TipoAccount %></b>
                    <br /><asp:Label ID="labRefresh" runat="server" Font-Size="X-Small"></asp:Label>
                    <br /><b><%=COUNTRY %></b>&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Button ID="btnHome" runat="server" Text="Home" OnClick="btnHome_Click" Font-Size="Small" Width="70px" style='vertical-align: top;'></asp:Button>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2">
                    <hr />
                    <br />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
            <asp:TableCell ColumnSpan="3"><asp:Image ID="imgTopLogo" runat="server" /><br /><br /><hr /><br /> </asp:TableCell>
        </asp:TableRow>
        </asp:Table>
        <asp:Table ID="tabMail" runat="server" HorizontalAlign="Center" >
            
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2"><asp:Label ID="labStatus" runat="server"></asp:Label><br />
                    <asp:Label runat="server" ID="labBack"></asp:Label><br />
                    <br /><hr /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2" HorizontalAlign="Left">
                    <br />
                    <asp:CheckBox ID="chkCopiaC" runat="server" Text="Invia una copia all'operatore" Checked="false" />
                    <br /><br />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell Width="25%" HorizontalAlign="Left">Mittente:</asp:TableCell>
                <asp:TableCell Width="75%" HorizontalAlign="Right" ><asp:TextBox ID="txFrom" runat="server" Width="90%" /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left">Destinatario:</asp:TableCell>
                <asp:TableCell HorizontalAlign="Right"><asp:TextBox ID="txTo" runat="server" Width="90%" /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left">Oggetto:</asp:TableCell>
                <asp:TableCell HorizontalAlign="Right"><asp:TextBox ID="txSubject" runat="server" Width="90%" /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2" HorizontalAlign="Center"><br />Messaggio:<br />
                    <asp:TextBox ID="txMessage" runat="server" TextMode="MultiLine" Rows="10" Width="90%" /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2"><hr /></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Left">Allegato:</asp:TableCell>
                <asp:TableCell HorizontalAlign="Right"><asp:Label ID="labAttach" runat="server"></asp:Label></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2">
                    <asp:Image ID="imgAttach" runat="server" />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2">
                    <hr /><br />
                    <asp:CheckBox ID="chkSetSendBozza" runat="server" Text="" Font-Size="Small" />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2" >
                    <br />
                    <asp:Button ID="btnSend" runat="server" Text="Invia" OnClick="btnSend_Click" Width="120px" /></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </div>
    </form>
</body>
</html>
