<%@ Page StylesheetTheme="acesso" EnableEventValidation="true" Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">       
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)                
        Dim oAcesso As New ClsAcesso        
        If Not IsPostBack Then            
            txtAcessoSenha.Attributes.Add("onkeyup", "k(this, 'enviar')")
            txtAcessoUsuario.Attributes.Add("onkeyup", "k(this, 'enviar')")
            With oAcesso
                If Not .AcessoStatusLogado Then
                    .AcessoTarget = "CORPO"
                    '.AcessoDeniedRedirect = "about:blank"
                    '.AcessoPageLogin = "AcessoLogin.aspx"
                    .AcessoAplCodigo = 37
                End If
            End With
        End If
    End Sub
    
    Protected Sub imgBtnLogin_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)        
        Dim oAcesso As New ClsAcesso                
        '---------------------------------------        
        With oAcesso                
            .cnnStr = ConnectionStrings("cnnStrAcesso").ToString 
            .AcessoAplCodigo = 37
            .AcessoUsuario = txtAcessoUsuario.Text
            .AcessoSenha = txtAcessoSenha.Text                        
            .VerificaLogin()
        End With        
        '---------------------------------------
    End Sub
    
</script>
    

<script language="javascript">
<!--
    Mensagens();
    VeryFrame();
    
        function VeryFrame(){
            if(self.parent.frames.length>0) self.parent.location="AcessoLogin.aspx"
        }
        
        function Mensagens(){    
        <% If Session("AcessoMensagens") IsNot Nothing Then    %>
        <%       If Session("AcessoMensagens") <> "" Then       %>
                    alert('<%=Replace(Session("AcessoMensagens"),"'","\'") %>');
        <% 	    End If  %>
        <% 	End If	%>   
        }
                
-->
</script>
    
    
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
        <META HTTP-EQUIV="Pragma" CONTENT="no-cache">
        <META HTTP-EQUIV="Expires" CONTENT="-1">
        <META HTTP-EQUIV="Cache-control" CONTENT="no-store">
		<title>Acesso ao Sistema</title>				    
		<script language="javascript" src="AcessoLogin.js" type="text/javascript" id="js" ></script>
</head>

	<body  onload="SetFoco()" bgColor="#FFFFFF" topmargin="0" leftmargin="0" >		
		<form id="form1" runat="server">
            <asp:HiddenField ID="IteMsgCodigo" runat="server" />
		<table align=center cellspacing="0" cellpadding="0" class="menu_hor" border="0" ID="Table1" style="width: 96%; height: 231px">
			<tr>
				<td align="center" bgcolor="#FFFFFF" style="height: 318px; width: 1003px;">					
					<div class="caixa" style="WIDTH: 350px ; height: 133px ">													
							<table bgcolor="#F7F7F7" border='1' cellpadding='1' cellspacing='0' width='100%' bordercolor='lightgrey'	style='BORDER-COLLAPSE: collapse' ID="Table3">			
								<tr >
								<td   align=center title="Acesso ao Sistema" style="height: 16px">
                                    <span style="font-size: 16pt">Acesso</span></td>
								</tr>
							</table>														
							<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0" style="height:100px" >				
								<TR><TD Class='Menu' style="width: 350px; height: 17px" > 
																
								<TABLE  align="center" id="Table4" cellSpacing="0" cellPadding="0" width="98%" border="0">				
                                    <tr>
                                        <td class="Menu" style="width: 274px; height: 14px; text-align: left">
                                        </td>
                                        <td class="Menu" style="width: 316px; height: 14px; text-align: left">
                                        </td>
                                        <td class="Menu" style="height: 14px; text-align: left">
                                        </td>
                                        <td class="Menu" style="width: 71px; height: 14px; text-align: left">
                                        </td>
                                        <td class="Menu" style="width: 296px; height: 14px; text-align: left">
                                        </td>
                                    </tr>
								<tr>
                                    <td class="Menu" style="width: 274px; height: 24px; text-align: left">
                                    </td>
									<td Class='Menu' style="width: 316px; text-align: right; height: 24px;">Usuário:</td>
                                    <td class="Menu" style="height: 24px; text-align: left">
                                    </td>
									<td Class='Menu' style="width: 71px; text-align: left; height: 24px;">                                        
									    <asp:TextBox runat="server"  ToolTip="Usuário de Acesso ao Sistema" ID="txtAcessoUsuario" CssClass="campos" Width="109px"></asp:TextBox>
									</td>
                                    <td class="Menu" style="width: 296px; height: 24px; text-align: left">
                                    </td>
								</tr>
								<tr>
                                    <td class="Menu" style="width: 274px; height: 28px; text-align: left">
                                    </td>
									<td Class='Menu' style="width: 316px; text-align: right; height: 28px;">
                                        Senha:</td>
                                    <td class="Menu" style="height: 28px; text-align: left">
                                    </td>
									<td Class='Menu' style="width: 71px; text-align: left; height: 28px;">                                        
									<asp:TextBox runat="server" TextMode=Password ID="txtAcessoSenha" CssClass="campos" ToolTip="Senha de Acesso ao Sistema" Width="109px"></asp:TextBox>
									</td>
                                    <td class="Menu" style="width: 296px; height: 28px; text-align: left">
                                    </td>
								</tr>
								<tr>
                                    <td class="Menu" style="width: 274px; height: 28px">
                                    </td>
									<td Class='Menu' style="height: 28px; width: 316px;">&nbsp;</td>                                    
                                    <td runat="server" class="Menu" onmouseout="MenuPrincipal(this,1)" onmouseover="MenuPrincipal(this,2)"
                                        style="height: 28px" title="Enviar Login">
                                    </td>
									<td  id="tdLogin" runat="server"  onmouseover="MenuPrincipal(this,2)" onmouseout="MenuPrincipal(this,1)" Class='Menu' title="Enviar Login" style="width: 71px; height: 28px">
									<asp:ImageButton ID="imgBtnLogin" ImageUrl="~/imagens/login.gif" runat="server" OnClick="imgBtnLogin_Click" BackColor="Transparent" BorderColor="Transparent" />&nbsp;</td>
                                    <td runat="server" class="Menu" onmouseout="MenuPrincipal(this,1)" onmouseover="MenuPrincipal(this,2)"
                                        style="width: 296px; height: 28px" title="Enviar Login">
                                    </td>
								</tr>																
								</table></td></tr>
								<TR><TD Class='Menu' style="width: 350px; height: 15px" > 
								</td>
								</tr>
							</table>						
					</div>
				</td>
			</tr>
		</table>		
		<p style="FONT-FAMILY: verdana; FONT-SIZE: 10px" align="center">Irmandade da Santa 
			Casa de Misericórdia de São Paulo</p>		
		</form>		
	</body>
</html>