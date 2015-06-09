<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default2.aspx.vb" Inherits="Default2" %>

<%@ Register Assembly="CORP" Namespace="CORP.NET" TagPrefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;
        &nbsp;&nbsp;&nbsp;&nbsp;<br />
        <asp:Button ID="Button1" runat="server" Text="Button" /><br />
        <br />
        &nbsp;<table style="width: 377px" border="1">
            <tr>
                <td style="width: 3px">
                    descrição&nbsp;da&nbsp;tabela</td>
                <td style="width: 352px">
                    <cc1:ClsAutoComplete ID="acpTabela" runat="server" Width="494px" ColumnText="TabDescricao" ColumnValue="TabCodigo" ColumnWhere="TabDescricao" ConnectionString="Provider=SQLOLEDB.1;Password=sistema;Persist Security Info=True;User ID=antonio;Initial Catalog=db_doacaoativos;Data Source=mbr_coreme" TableSelect="Tabela" AutoCompleteStyle="AutoComplete" ButtonCaption="Pesquisar" ButtonStyle="SimpleButton" CaptionStyle="AlignLeft" ShowCaption="False" />
                </td>
            </tr>
            <tr>
                <td style="width: 3px">
                    campo&nbsp;de&nbsp;teste</td>
                <td style="width: 352px">
                    <asp:DropDownList ID="DropDownList1" runat="server" Width="320px">
                        <asp:ListItem Value="Valor 1">Texto 1</asp:ListItem>
                        <asp:ListItem Value="Valor 2">Texto 2</asp:ListItem>
                    </asp:DropDownList></td>
            </tr>
            <tr>
                <td style="width: 3px">
                    descricao&nbsp;dos&nbsp;ítens</td>
                <td style="width: 352px"><cc1:ClsAutoComplete ID="acpItensTabela" runat="server" Width="202px" ColumnText="IteTabNome" ColumnValue="IteTabCodigo" ColumnWhere="IteTabNome" ConnectionString="Provider=SQLOLEDB.1;Password=sistema;Persist Security Info=True;User ID=antonio;Initial Catalog=db_doacaoativos;Data Source=mbr_coreme" TableSelect="Itens_Tabela" CaptionStyle="AlignLeft" AutoCompleteStyle="ButtonClick" ButtonCaption="Pesquisar" />
                </td>
            </tr>
        </table>
        <br />
        &nbsp;para valor tabela:
        <asp:Label ID="lblValueTab" runat="server" Width="143px"></asp:Label><br />
        &nbsp;para texto tabela:&nbsp;
        <asp:Label ID="lblTextTab" runat="server" Width="308px"></asp:Label><br />
        <br />
        &nbsp;para valor itens_tabela:
        <asp:Label ID="lblValueItens" runat="server" Width="211px"></asp:Label><br />
        &nbsp;para texto itens_tabela:&nbsp;
        <asp:Label ID="lblTextItens" runat="server" Width="374px"></asp:Label><br />
        <br />
        <br />
    </div>
    </form>
</body>
</html>
