<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Download.aspx.cs" Inherits="Download" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="shortcut icon" type="image/x-icon" href="pics/data-icon.ico" />
    <title>Lavorazioni - Download - <%=COUNTRY %></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button ID="btnPB" OnClick="btnPB_Click" Visible="false" runat="server" />
        <asp:ScriptManager ID="ScriptMgr" runat="server" EnablePageMethods="true"></asp:ScriptManager>
    </div>
    </form>
</body>
</html>
