<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;
        <asp:TextBox ID="txt" runat="server" Height="94px" TextMode="MultiLine" Width="239px"></asp:TextBox>
        <asp:TextBox ID="txtComp" runat="server" Height="94px" TextMode="MultiLine" Width="239px"></asp:TextBox>
        <asp:Button ID="btnDesc" runat="server" Text="Desc" />
        <asp:Button ID="btnComp" runat="server" Text="Comp" />
    <asp:GridView ID="grd" runat="server"></asp:GridView>
        &nbsp;
    </div>
    </form>
</body>
</html>
