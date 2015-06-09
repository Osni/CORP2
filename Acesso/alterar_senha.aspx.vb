Imports System.Web.Configuration.WebConfigurationManager

Partial Class alterar_senha
    Inherits PageX

    Private ClsDB As New ClsDB
    Private ClsSQL As ClsSQL

    Private oTool As New ClsTools

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            txtUsuUsuario.Text = Session("AcessoUsuUsuario")
        End If
    End Sub

    Protected Sub btnEnviar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEnviar.Click
        Dim strSenha As String
        Dim strSQL As StringBuilder
        Dim reader As Data.OleDb.OleDbDataReader
        Try
            ClsDB.sConStr = ConnectionStrings("cnnStrAcesso").ToString()
            ClsSQL = New ClsSQL
            With ClsSQL
                .sTable = "usuario"
                .AddCol("UsuCodigo") '
                .AddCol("UsuUsuario")
                .AddCol("UsuSenha")
                reader = ClsDB.GetDataReader(.GetSELECT("UsuCodigo = " & Session("AcessoUsuCodigo")))
                If reader.Read() Then
                    strSenha = reader("UsuSenha")
                    If strSenha = txtUsuSenha.Text Then
                        strSQL = New StringBuilder
                        With strSQL
                            .Append("UPDATE usuario SET UsuSenha = '")
                            .Append(txtUsuSenhaNova.Text)
                            .Append("' WHERE UsuCodigo = ")
                            .Append(Session("AcessoUsuCodigo"))
                        End With
                        ClsDB.SetCommandSQL(strSQL.ToString)
                    Else                        
                        oTool.ShowMessage("Senha atual não confere.", _
                                          "Alteração de Senha", Me.Page, ClsTools.TMsgStyleIcon.MSG_ERROR)
                        Exit Sub
                    End If
                Else                    
                    oTool.ShowMessage("Problemas ao localizar usuário no banco!", _
                                      "Alteração de Senha", Me.Page, ClsTools.TMsgStyleIcon.MSG_ERROR)
                    Exit Sub
                End If
            End With

            oTool.ShowMessage("Senha alterada com sucesso!", "Alteração de Senha", _
                               Me.Page, ClsTools.TMsgStyleIcon.MSG_INFORMATION, _
                               "", "document.location.href='logo.aspx';")

        Catch ex As Exception            
            oTool.ShowMessage("Erro:" & ex.Message, "Erro!", Me.Page, ClsTools.TMsgStyleIcon.MSG_ERROR)
        End Try
    End Sub

End Class
