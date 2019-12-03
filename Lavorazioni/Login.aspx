<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Login.aspx.cs" Inherits="Login" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="Style/login.css" rel="stylesheet" type="text/css" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Lavorazioni - Login</title>
    <script type="text/javascript">
        function enable() {
            var table = document.getElementById("tabLogin");
            if (table.rows[4].style.visibility == "visible") {
                //$(table.rows[4]).hide();
                //$(table.rows[4]).css("display", "none");
                table.rows[4].style.visibility = 'hidden';
            }
            else {
                //$(table.rows[4]).show();
                //$(table.rows[4]).css("display", "table-row");
                table.rows[4].style.visibility = 'visible';
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div style="text-align:center; width: 100%;">
        <div class="box">
        <div class="content">
            
            <asp:Table runat="server" Width="100%" ID="tabLogin">
                <asp:TableRow>
                    <asp:TableCell HorizontalAlign="Left" >
                        <asp:Label ID="labUserName" runat="server" AssociatedControlID="txUserName">nome utente:</asp:Label></asp:TableCell>
                    <asp:TableCell ColumnSpan="2" ><asp:TextBox ID="txUserName" runat="server" Width="100%"></asp:TextBox></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell Width="20%" HorizontalAlign="Left"><asp:Label ID="labPassword" runat="server" AssociatedControlID="txPassword">password:</asp:Label></asp:TableCell>
                    <asp:TableCell Width="50%"><asp:TextBox ID="txPassword" runat="server" TextMode="Password" Width="100%"></asp:TextBox></asp:TableCell>
                    <asp:TableCell Width="30%">
                        <asp:Label ID="labYear" runat="server" Text="Anno: " Width="40%" Font-Bold="true"></asp:Label><asp:DropDownList ID="dropYear" runat="server" Width="60%" ></asp:DropDownList></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell ColumnSpan="3">
                        <asp:CheckBox ID="chkRememberMe" runat="server" CssClass="checkbox" Text="ricordami" Checked="true" /></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell ColumnSpan="3">
                        <br />
                        <asp:CheckBox ID="chkAmazon" runat="server" CssClass="checkbox" Text="<img src='pics/amazon-logo-orange.png' width='200px' style='vertical-align: middle;' />" Checked="false" onclick="enable();"/></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow ID="rowMarkets" runat="server">

                </asp:TableRow>
            </asp:Table><br />
            <asp:ImageButton ID="LoginButton" runat="server" CommandName="Login" ImageUrl="pics/login3.png" ImageAlign="Middle" ValidationGroup="Accesso" 
                OnClick="LoginButton_Click" Width="250px" CssClass="btnLogin"   />
        </div>
        </div>
    </div>
    </form>
</body>
</html>
