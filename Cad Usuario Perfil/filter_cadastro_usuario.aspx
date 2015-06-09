<%@ Page StylesheetTheme="corp" Title="Wizard RptView 1.0" Language="VB" Inherits="PageX" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            With filter                
                .AddCol("UsuCodigo", "Código", "Código", , 5, , , "cadastro_usuario.aspx")
                .AddCol("UsuNomeCompleto", "Nome", "Nome do Usuário", , 50)
                .AddCol("UsuUsuario", "Login", "Login do usuário", , 30)
                .AddCol("UsuRamal", "Ramal", "Ramal do Usuário", , 10)
                .AddCol("AreNome", "Setor", "Setor", , 50)
                
                .FilterPagePreLoad = True
                .FilterStrConnection = ConnectionStrings("cnnStrAcessoFiltro").ToString
                .FilterType = ClsFilter.PFilterType.FullFilter                
                .FilterTableName = "VW_USUARIO"                
                .FilterReturnFormName = "cadastro_usuario.aspx"
                .FilterPageSize = 20
                .FilterOrderByCols = True
                .Visible = True
            End With
        End If
        '---------------------------------------------------------------
    End Sub
       
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
</head>
<body>
    <form id="form1" runat="server">
        <div id="corpo_documento">
            <table id="maintable" border="0" cellpadding="0" cellspacing="1" style="width: 700px">
                <tr>
                    <td id="TituloJanela">
                        <asp:Label ID="lblTitulo" runat="server">Usuários</asp:Label></td>
                </tr>
                <tr>
                    <td align="Left">
                        <crp:ClsFilter ID="filter" runat="server" Width="99%" />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
