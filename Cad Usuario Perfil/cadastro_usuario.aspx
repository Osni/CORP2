<%@ Page Language="VB" AutoEventWireup="false" CodeFile="cadastro_usuario.aspx.vb"
    Inherits="cadastro_usuario" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="frmMain" runat="server">
        <div>
            <table border="0" cellpadding="0" id="maintable" cellspacing="1" style="width: 700px;">
                <tr>
                    <th id="TituloJanela" colspan="2" valign="middle" style="height: 28px">
                        Cadastro de Usuários</th>
                </tr>
                <tr>
                    <td align="center" valign="top" style="height: 316px">
                        <table id="tabDadosUsuario" border="0" cellspacing="1" width="100%">
                            <tr>
                                <td align="center" colspan="2" valign="middle" style="height: 11px; text-align: center" nowrap=nowrap>
                                    <crp:ToolBar ID="tb" runat="server" BorderColor="#E0E0E0" BorderedCells="False" BorderStyle="Solid"
                                        BorderWidth="1px" Padding="3" Spacing="3">
                                        <ToolButtons>
                                            <%--<crp:ToolButton ID="btnSalvar" DisabledImageUrl="~/imagens/imgBtn/toolbar_save_d.gif"
                                                ImageUrl="~/imagens/imgBtn/toolbar_save.gif" OnClientClick="if(!confirm(&quot;Confirma Dados?&quot;)) return false;"
                                                ToolTip="Salvar Usu&#225;rio" CausesValidation="True" CommandArgument="" CommandName=""
                                                Enabled="True" ImageAlign="NotSet" PostBackUrl="" RedirectURL="" Text="" ValidationGroup="">
                                            </crp:ToolButton>
--%>
                                            <crp:ToolButton ID="btnPesquisar" ImageUrl="~/imagens/imgBtn/toolbar_filter.gif"
                                                OnClientClick="javascript: window.location.href = &quot;filter_cadastro_usuario.aspx&quot;; return false;"
                                                ToolTip="Pesquisar Usu&#225;rios" CausesValidation="True" CommandArgument=""
                                                CommandName="" DisabledImageUrl="" Enabled="True" ImageAlign="NotSet" PostBackUrl=""
                                                RedirectURL="" Text="" ValidationGroup=""></crp:ToolButton>
                                            <%--<crp:ToolButton ID="btnEditar" DisabledImageUrl="~/imagens/imgBtn/toolbar_update_d.gif"
                                                Enabled="False" ImageUrl="~/imagens/imgBtn/toolbar_update.gif" OnClientClick="if(!confirm(&quot;Confirma Dados?&quot;)) return false;"
                                                ToolTip="Atualizar Usu&#225;rio" CausesValidation="True" CommandArgument="" CommandName=""
                                                ImageAlign="NotSet" PostBackUrl="" RedirectURL="" Text="" ValidationGroup=""></crp:ToolButton>
--%>
                                            <%--<crp:ToolButton ID="btnExcluir" DisabledImageUrl="~/imagens/imgBtn/toolbar_del_d.gif"
                                                Enabled="False" ImageUrl="~/imagens/imgBtn/toolbar_del.gif" OnClientClick="if(!confirm(&quot;Excluir Usu&#225;rio?&quot;)) return false;"
                                                ToolTip="Excluir Usu&#225;rio" CausesValidation="True" CommandArgument="" CommandName=""
                                                ImageAlign="NotSet" PostBackUrl="" RedirectURL="" Text="" ValidationGroup=""></crp:ToolButton>
                                            <crp:ToolButton ID="btnLimpar" CausesValidation="True" CommandArgument="" CommandName=""
                                                DisabledImageUrl="~/imagens/imgBtn/toolbar_new_d.gif" ImageUrl="~/imagens/imgBtn/toolbar_new.gif"
                                                OnClientClick="javascript: window.location.href = &quot;cadastro_usuario.aspx&quot;; return false;"
                                                ToolTip="Limpar Tela" Enabled="True" ImageAlign="NotSet" PostBackUrl="" RedirectURL=""
                                                Text="" ValidationGroup=""></crp:ToolButton>--%>
                                        </ToolButtons>
                                    </crp:ToolBar>
                                    Pesquisar
                                </td>
                            </tr>
                            <tr>
                                <th align="right" style="width: 418px">
                                    Código:</th>
                                <td align="left">
                                    <asp:TextBox ID="txtUsuCodigo" runat="server" ReadOnly="True" MaxLength="10" BackColor="#F0F0F0"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <th align="right" style="width: 418px">
                                    <span style="color: #cc0000">*</span> Nome:</th>
                                <td align="left">
                                    <asp:TextBox ID="txtUsuNomeCompleto" runat="server" MaxLength="100" Width="318px"
                                        BackColor="#F0F0F0" Enabled="False"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <th align="right" style="width: 418px; height: 5px;">
                                    &nbsp;<span style="color: #cc0000">*</span> Usuário:</th>
                                <td align="left" style="height: 5px">
                                    <asp:TextBox ID="txtUsuUsuario" runat="server" MaxLength="20" BackColor="#F0F0F0"
                                        Enabled="False"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <th align="right" style="width: 418px">
                                    <asp:Label ID="lblUsuSenha" runat="server" Text='<span style="color:#cc0000">*</span> Senha'></asp:Label></th>
                                <td style="width: 186px;" align="left">
                                    <asp:TextBox ID="txtUsuSenha" runat="server" MaxLength="20" TextMode="Password" BackColor="#F0F0F0"
                                        Enabled="False"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <th align="right" style="width: 418px">
                                    Ramal:</th>
                                <td align="left" style="width: 186px;">
                                    <asp:TextBox ID="txtUsuRamal" runat="server" MaxLength="20" BackColor="#F0F0F0" Enabled="False"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <th align="right" style="width: 418px">
                                    &nbsp;<span style="color: #cc0000">*</span> Área:</th>
                                <td align="left">
                                    <asp:DropDownList ID="cboAreCodigo" runat="server" Width="196px" BackColor="#F0F0F0"
                                        Enabled="False">
                                    </asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    &nbsp;</td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <table id="tabPerfil">
                                        <tr>
                                            <th colspan="3" id="TituloTopico">
                                                Perfil de acesso aos objetos
                                            </th>
                                        </tr>
                                        <tr>
                                            <td colspan="3" style="text-align: left">
                                                Perfis cadastrados
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3" style="text-align: left">
                                                <asp:DropDownList ID="cboGruCodigo" runat="server" Width="411px" AutoPostBack="true">
                                                </asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="width: 290px; text-align: left" valign="top">
                                                <crp:ClsControlBox ID="lstDisponiveis" runat="server" MultiSelect="True" AutoGenerateColumns="False"
                                                    BoxTitle="Objetos Disponíveis" Height="215px" BoxStyle="ListBox" Width="308px"></crp:ClsControlBox>
                                            </td>
                                            <td style="width: 218px">
                                                <asp:ImageButton ID="btnMoveSelecaoObjetos" runat="server" ToolTip="Dá permissão de acesso aos objetos selecionados."
                                                    ImageUrl="~/imagens/move_select_right.gif" BorderColor="Transparent" BackColor="Transparent">
                                                </asp:ImageButton>
                                                <asp:ImageButton ID="btnMoveTodosObjetos" runat="server" ToolTip="Dá permissão de acesso aos objetos selecionados."
                                                    ImageUrl="~/imagens/move_all_right.gif" BorderColor="Transparent" BackColor="Transparent">
                                                </asp:ImageButton>
                                                <asp:ImageButton ID="btnRetornaTodosObjetos" runat="server" ToolTip="Remove permissão de acesso a todos os objetos."
                                                    ImageUrl="~/imagens/move_all_left.gif" BorderColor="Transparent" BackColor="Transparent">
                                                </asp:ImageButton>
                                                <asp:ImageButton ID="btnRetornaSelecaoObjetos" runat="server" ToolTip="Remove permissão de acesso aos objetos selecionados."
                                                    ImageUrl="~/imagens/move_select_left.gif" BorderColor="Transparent" BackColor="Transparent">
                                                </asp:ImageButton>
                                            </td>
                                            <td style="width: 247px" valign="top">
                                                <crp:ClsControlBox ID="lstAtribuidos" runat="server" MultiSelect="True" AutoGenerateColumns="False"
                                                    BoxTitle="Objetos Atribuidos" Height="213px" BoxStyle="ListBox" Width="308px"></crp:ClsControlBox>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
