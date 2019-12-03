<%@ Page Language="C#" AutoEventWireup="true" CodeFile="lavShowFolder.aspx.cs" Inherits="lavShowFolder" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server"><link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <link href="Style/dettaglio.css" rel="stylesheet" type="text/css" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Lavorazioni - Mostra Cartella</title>
    
    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>
    <script type="text/javascript" src="js/jquery.magnifier.js"></script>

    <script type="text/javascript" src="js/tinymce/tinymce.js"></script>
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
                            <h1>Lavorazioni - Mostra cartella</h1>
                            <h4><%= LAVID %> </h4>
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
            <asp:TableCell ColumnSpan ="3">
                <asp:TreeView ID="trvDirectories" runat="server" ImageSet="XPFileExplorer" NodeIndent="15">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <NodeStyle Font-Names="Tahoma" Font-Size="8pt" ForeColor="Black" HorizontalPadding="2px"
                        NodeSpacing="0px" VerticalPadding="2px" ImageUrl="~/pics/downarrow.png" ></NodeStyle>
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px"
                        VerticalPadding="0px" />
                </asp:TreeView>
            </asp:TableCell>
        </asp:TableRow>
        </asp:Table>
        </div>
    </form>
</body>
</html>
