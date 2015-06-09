<%@ Page Language="VB" AutoEventWireup="false" CodeFile="alterar_senha.aspx.vb" Inherits="alterar_senha" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="frmMain" runat="server" method="post">
    <div id="corpo_documento">
        <table id="maintable" border="0" cellpadding="0" cellspacing="1" style="width: 700px">
            <tr>
                <td id="TituloJanela" colspan="1">
                    <asp:Label ID="lblTitulo" runat="server" Text="Alterar Senha"></asp:Label></td>
            </tr>
            <tr>
                <td align="center" valign="middle" style="height: 185px">
                    <table style="width: 100%" border="0" cellpadding="3" cellspacing="1">
                        <tr>
                            <td align="right" style="width: 339px">
                                <asp:Label ID="lblUsuUsuario" runat="server" Text="Usuário"></asp:Label></td>
                            <td style="text-align: left;">
                                <asp:TextBox ID="txtUsuUsuario" runat="server" ReadOnly="True" MaxLength="10"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td align="right" style="width: 339px">
                                <asp:Label ID="lblUsuSenha" runat="server" Text="Senha Atual" ></asp:Label></td>
                            <td style="text-align: left;">
                                <asp:TextBox ID="txtUsuSenha" runat="server" MaxLength="15" TextMode="Password" EnableViewState="False"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td align="right" style="width: 339px">
                                <asp:Label ID="lblUsuSenhaNova" runat="server" Text="Senha Nova" Width="129px" ></asp:Label></td>
                            <td style="text-align: left;">
                                <asp:TextBox ID="txtUsuSenhaNova" runat="server" MaxLength="15" EnableViewState="False" TextMode="Password"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td align="right" style="width: 339px">
                                <asp:Label ID="lblUsuSenhaConfirmacao" runat="server" Text="Confirmação" ></asp:Label></td>
                            <td style="text-align: left;">
                                <asp:TextBox ID="txtUsuSenhaNovaRetype" runat="server" MaxLength="15" EnableViewState="False" TextMode="Password"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" style="height: 28px">
                                <asp:Button ID="btnEnviar" runat="server" Text="Enviar" Height="22px" Width="90px" /></td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2">
                    <asp:CompareValidator ID="CompareValidator" runat="server" ControlToCompare="txtUsuSenhaNova"
                        ControlToValidate="txtUsuSenhaNovaRetype" Display="Dynamic" ErrorMessage="Nova senha e confirmação não conferem."></asp:CompareValidator><asp:RequiredFieldValidator ID="RequiredFieldValidator" runat="server" ControlToValidate="txtUsuSenha"
                        Display="Dynamic" ErrorMessage="Digite a Senha Antiga"></asp:RequiredFieldValidator></td>
                        </tr>
                    </table>
                    </td>
            </tr>
        </table>    
    </div>
    
    </form>
</body>
</html>
