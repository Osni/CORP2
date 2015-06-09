
Partial Class cadastro_usuario
    Inherits PageX


#Region "Fields"
    Private cnnAcesso As New OleDbConnection
    Private traAcesso As OleDbTransaction
    Private oTool As New ClsTools
    Private oDBAcesso As New ClsDB
    Private strSQL As String = String.Empty
    Private intUsuCodigo As Integer
    Private intAcessoAplCodigo As Integer
    Private intAcessoUsuCodigo As Integer
#End Region

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '--------------------------------------------------------
        intUsuCodigo = Val(Request("UsuCodigo"))
        intAcessoAplCodigo = Val(Session("AcessoAplCodigo"))
        intAcessoUsuCodigo = Val(Session("AcessoUsuCodigo"))
        oDBAcesso.sConStr = ConnectionStrings("cnnStrAcesso").ToString
        '--------------------------------------------------------
        If Not IsPostBack Then
            Try
                '--------------------------------------------------------
                'Carregando Areas
                strSQL = "SELECT AreCodigo=0, AreNome='[SELECIONE]' UNION "
                strSQL &= "SELECT AreCodigo, AreNome FROM VW_AREA"
                oDBAcesso.AddCombo(strSQL, "AreCodigo", "AreNome", cboAreCodigo)
                '--------------------------------------------------------
                'Carregando Perfil
                strSQL = "SELECT GruCodigo=0, GruNome='[SELECIONE]' UNION "
                strSQL &= "SELECT GruCodigo, GruNome FROM GRUPO WHERE AplCodigo = " & intAcessoAplCodigo.ToString
                oDBAcesso.AddCombo(strSQL, "GruCodigo", "GruNome", cboGruCodigo)
                '--------------------------------------------------------
                'Carregando Listas
                If intUsuCodigo <> 0 Then
                    Call CriaListas()
                    Call CarregaDadosUsuario()
                    '--------------------------------------------------------
                    Call CarregaListas()
                Else
                    Throw New Exception("Usuário não definido!")
                End If
                '--------------------------------------------------------
            Catch ex As Exception
                Throw ex
            End Try
        End If
    End Sub

    Protected Sub tb_OnClickMe(ByRef Btn As CORP.NET.ToolButton, ByVal Index As Integer) Handles tb.OnClickMe
        Dim strMensagem As String = String.Empty
        Select Case Index
            Case 0
                '    strMensagem = ValidaCampos()
                '    If strMensagem <> String.Empty Then
                '        oTool.ShowMessage("<span style='color:#0000ff'><b>Favor informar:</b></span> <br />" & strMensagem, "Atenção", Me.Page, ClsTools.TMsgStyleIcon.MSG_WARNING)
                '    Else
                '        Call Salvar()
                '    End If
                'Case 2
                'strMensagem = ValidaCampos()
                If strMensagem <> String.Empty Then
                    oTool.ShowMessage("<span style='color:#0000ff'><b>Favor informar:</b></span> <br />" & strMensagem, "Atenção", Me.Page, ClsTools.TMsgStyleIcon.MSG_WARNING)
                Else
                    Call Editar()
                End If
                'Case 3
                '    Call Excluir()
        End Select
    End Sub

    Protected Sub cboGruCodigo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGruCodigo.SelectedIndexChanged
        Call CarregaListas()
    End Sub

    Protected Sub btnMoveSelecaoDetalhe_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnMoveSelecaoObjetos.Click
        Call Mover(lstDisponiveis, lstAtribuidos)
    End Sub

    Protected Sub btnMoveTodosDetalhe_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnMoveTodosObjetos.Click
        Call Mover(lstDisponiveis, lstAtribuidos, True)
    End Sub

    Protected Sub btnRetornaTodosDetalhe_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnRetornaTodosObjetos.Click
        Call Mover(lstAtribuidos, lstDisponiveis, True)
    End Sub

    Protected Sub btnRetornaSelecaoDetalhe_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnRetornaSelecaoObjetos.Click
        Call Mover(lstAtribuidos, lstDisponiveis)
    End Sub

#End Region

#Region "Methods"

    Private Sub CarregaDadosUsuario()
        '--------------------------------------------------------
        strSQL = "SELECT UsuCodigo, UsuUsuario, UsuNomeCompleto, AreCodigo, UsuRamal "
        strSQL &= "FROM VW_USUARIO WHERE UsuCodigo = " & intUsuCodigo.ToString
        '--------------------------------------------------------
        With oDBAcesso
            .GetDataReader(strSQL)
            If .dtr.Read Then
                txtUsuCodigo.Text = .dtr("UsuCodigo")
                txtUsuNomeCompleto.Text = .dtr("UsuNomeCompleto")
                txtUsuUsuario.Text = .dtr("UsuUsuario")
                txtUsuRamal.Text = .dtr("UsuRamal")
                cboAreCodigo.SelectedValue = .dtr("AreCodigo")
                '------------------------------------------
                .dtr.Close()
                'Call MudarModo(False)
            Else
                Throw New Exception("Problemas ao carregar Usuário!")
            End If
        End With
        '--------------------------------------------------------
    End Sub

    Private Function ValidaCampos(Optional ByVal Insert As Boolean = True) As String
        Dim strMensagem As String = String.Empty

        'If txtUsuNomeCompleto.Text.Trim = String.Empty Then _
        '    strMensagem = "Nome Completo"

        'If txtUsuUsuario.Text.Trim = String.Empty Then
        '    strMensagem &= IIf(strMensagem.Trim <> String.Empty, "<BR />", "") & "Usuário"
        'End If

        'If cboAreCodigo.SelectedValue = 0 Then
        '    strMensagem &= IIf(strMensagem.Trim <> String.Empty, "<BR />", "") & "Área"
        'End If

        'If Insert And txtUsuSenha.Text.Trim = String.Empty Then
        '    strMensagem &= IIf(strMensagem.Trim <> String.Empty, "<BR />", "") & "Senha"
        'End If

        'If txtUsuRamal.Text.Trim = String.Empty Then
        '    strMensagem &= IIf(strMensagem.Trim <> String.Empty, "<BR />", "") & "Ramal"
        'End If

        'If lstAtribuidos.RowCount = 0 Then
        '    strMensagem &= IIf(strMensagem.Trim <> String.Empty, "<BR />", "") & "Perfil de acesso aos objetos"
        'End If

        Return strMensagem
    End Function

    Private Sub Salvar()
        Dim blnConcluido As Boolean = False
        Try
            '----------------------------------------            
            cnnAcesso = oDBAcesso.GetOpenDB()
            traAcesso = cnnAcesso.BeginTransaction
            '----------------------------------------
            If Val(txtUsuCodigo.Text) = 0 Then
                '----------------------------------------
                strSQL = "SELECT * FROM VW_USUARIO WHERE UsuUsuario = '" & txtUsuUsuario.Text.Trim & "'"
                dta = New DataTable
                dta = oDBAcesso.GetDataTable(strSQL)
                '----------------------------------------
                If dta.Rows.Count <> 0 Then
                    Throw New Exception("Usuário já cadastrado!")
                Else
                    '----------------------------------------
                    'Gravando informações do usuário no Módulo de Acesso
                    strSQL = "INSERT INTO USUARIO (UsuUsuario, UsuNomeCompleto, UsuSenha, UsuRamal, AreCodigo) VALUES "
                    strSQL &= "('" & txtUsuUsuario.Text.Trim.Replace("'", "''") & "', '" & txtUsuNomeCompleto.Text.Trim.Replace("'", "''") & "', '"
                    strSQL &= txtUsuSenha.Text.Trim.Replace("'", "''") & "', '" & txtUsuRamal.Text.Trim.Replace("'", "''") & "', " & cboAreCodigo.SelectedValue & ")"
                    '----------------------------------------
                    intUsuCodigo = oDBAcesso.SetCommandSQLReturnMax(strSQL, "db_acesso..USUARIO", "UsuCodigo", cnnAcesso, traAcesso)
                    '----------------------------------------                
                End If
            Else
                intUsuCodigo = txtUsuCodigo.Text
            End If
            '----------------------------------------            
            Call GravaPerfil()
            '----------------------------------------
            traAcesso.Commit()
            '----------------------------------------
            txtUsuCodigo.Text = intUsuCodigo
            blnConcluido = True
            '----------------------------------------
        Catch ex As Exception
            Try : traAcesso.Rollback() : Catch : End Try
            oTool.ShowMessage("Erro:<br />" & ex.Message, "Erro", Me.Page)
        Finally
            If blnConcluido Then
                Call ClearCache()
                oTool.ShowMessage("Concluído!", "O.S.", Me.Page, ClsTools.TMsgStyleIcon.MSG_INFORMATION, , "javascript:window.location.href='filter_cadastro_usuario.aspx';")
            End If
        End Try
    End Sub

    Private Sub Editar()
        Dim blnConcluido As Boolean = False
        Try
            '=======================================
            'ATUALIZANDO NO MÓDULO DE ACESSO
            cnnAcesso = oDBAcesso.GetOpenDB
            traAcesso = cnnAcesso.BeginTransaction
            '----------------------------------------
            'If txtUsuNomeCompleto.Enabled Then
            '    strSQL = " UPDATE USUARIO SET " & vbCrLf
            '    strSQL &= "   UsuUsuario = '" & txtUsuUsuario.Text.Trim.Replace("'", "''") & "', " & vbCrLf
            '    strSQL &= "   UsuNomeCompleto = '" & txtUsuNomeCompleto.Text.Trim.Replace("'", "''") & "', " & vbCrLf
            '    If txtUsuSenha.Text.Trim <> String.Empty Then
            '        strSQL &= "   UsuSenha =  '" & txtUsuSenha.Text.Trim.Replace("'", "''") & "', " & vbCrLf
            '    End If
            '    strSQL &= "   UsuRamal = '" & txtUsuRamal.Text.Trim.Replace("'", "''") & "', " & vbCrLf
            '    strSQL &= "   AreCodigo = " & cboAreCodigo.SelectedValue & ", " & vbCrLf
            '    strSQL &= "   UsuCodigoAltera = " & Session("AcessoUsuCodigo")
            '    strSQL &= " WHERE UsuCodigo = " & txtUsuCodigo.Text
            '    oDBAcesso.SetCommandSQL(strSQL, cnnAcesso, traAcesso)
            'End If
            '----------------------------------------
            'Gravando Perfil de Acesso
            Call GravaPerfil()
            '----------------------------------------
            traAcesso.Commit()
            blnConcluido = True
            '----------------------------------------
        Catch ex As Exception
            Try : traAcesso.Rollback() : Catch : End Try
            oTool.ShowMessage("Erro:<br />" & ex.Message, "Erro", Me.Page)
        Finally
            If blnConcluido Then oTool.ShowMessage("Atualizado com sucesso!", "Cadastro", Me.Page, ClsTools.TMsgStyleIcon.MSG_INFORMATION)
        End Try
    End Sub

    'Private Sub Excluir()
    '    Try
    '        strSQL = "UPDATE USUARIO SET "
    '        strSQL &= "   Ativo = 'N', " & vbCrLf
    '        strSQL &= "   UsuCodigoAltera = " & Session("AcessoUsuCodigo") & ", " & vbCrLf
    '        strSQL &= "   Data = GETDATE() " & vbCrLf
    '        strSQL &= " WHERE UsuCodigo = " & txtUsuCodigo.Text
    '        '----------------------------------------
    '        oDBAcesso.SetCommandSQL(strSQL)
    '    Catch ex As Exception
    '        oTool.ShowMessage("Erro:<br />" & ex.Message, "Erro", Me.Page)
    '    End Try
    'End Sub

    'Private Sub MudarModo(ByVal blnInclusao As Boolean)
    '    'btnEditar.Enabled = Not blnInclusao
    '    btnSalvar.Enabled = blnInclusao
    '    'btnExcluir.Enabled = Not blnInclusao
    '    lblUsuSenha.Text = IIf(blnInclusao, "<span style=""color:#cc0000"">*</span>&nbsp;", "") & "Senha"
    'End Sub

    Private Sub GravaPerfil()

        oDBAcesso.SetCommandSQL("DELETE FROM PERFIL WHERE UsuCodigo = " & intUsuCodigo.ToString & _
                          " AND AplCodigo = " & AplCodigo, cnnAcesso, traAcesso)

        For i As Integer = 0 To lstAtribuidos.RowCount - 1

            strSQL = "INSERT INTO PERFIL ("
            strSQL &= "UsuCodigo, "
            strSQL &= "ObjCodigo, "
            strSQL &= "AplCodigo, "
            strSQL &= "TpPerCodigo, "
            strSQL &= "UsuCodigoAltera "
            strSQL &= ") VALUES ("
            strSQL &= intUsuCodigo.ToString & ", "
            strSQL &= lstAtribuidos.ColumnValue("ObjCodigo", i).ToString & ", "
            strSQL &= AplCodigo & ", "
            strSQL &= "1, "
            strSQL &= Session("AcessoUsuCodigo") & ")"

            oDBAcesso.SetCommandSQL(strSQL, cnnAcesso, traAcesso)
        Next

    End Sub

    Private Sub CarregaListas()
        'If intUsuCodigo <> 0 Then
        Dim strSQLDe As String = String.Empty
        Dim strSQLPara As String = String.Empty
        '------------------------------------------
        If cboGruCodigo.SelectedValue <> 0 Then
            'Lista disponíveis
            strSQLDe = "SELECT o.ObjCodigo, o.ObjNomeMenu FROM "
            strSQLDe &= "VW_OBJETO o INNER JOIN GRUPO_OBJETO g "
            strSQLDe &= "ON(o.ObjCodigo = g.ObjCodigo)"
            strSQLDe &= "WHERE AplCodigo = " & AplCodigo & " AND GruCodigo = " & cboGruCodigo.SelectedValue & "  AND o.ObjCodigo NOT IN "
            strSQLDe &= "(SELECT ObjCodigo FROM  VW_PERFIL WHERE UsuCodigo = " & intUsuCodigo & " AND AplCodigo = " & AplCodigo & ")"

            'Lista atribuídos
            strSQLPara = "SELECT o.ObjCodigo, o.ObjNomeMenu " & vbCrLf
            strSQLPara &= "FROM " & vbCrLf
            strSQLPara &= "VW_OBJETO o INNER JOIN VW_PERFIL p ON(o.ObjCodigo = p.ObjCodigo)" & vbCrLf
            strSQLPara &= "WHERE p.AplCodigo = " & AplCodigo & " AND p.UsuCodigo = " & intUsuCodigo
        Else
            'Lista disponíveis
            strSQLDe = "SELECT ObjCodigo, ObjNomeMenu FROM VW_OBJETO WHERE "
            strSQLDe &= " AplCodigo = " & AplCodigo & " AND  "
            strSQLDe &= " ObjCodigo NOT IN ("
            strSQLDe &= " SELECT ObjCodigo FROM PERFIL "
            strSQLDe &= "WHERE AplCodigo = " & AplCodigo
            strSQLDe &= " AND UsuCodigo = " & intUsuCodigo.ToString & ")"

            'Lista atribuídos
            strSQLPara = "SELECT ObjCodigo, ObjNomeMenu FROM VW_OBJETO WHERE "
            strSQLPara &= " AplCodigo = " & AplCodigo & " AND "
            strSQLPara &= " ObjCodigo IN ("
            strSQLPara &= " SELECT ObjCodigo FROM PERFIL "
            strSQLPara &= "WHERE AplCodigo = " & AplCodigo
            strSQLPara &= " AND UsuCodigo = " & intUsuCodigo.ToString & ")"
        End If
        '-------------------------------------------
        'Carregando Lista Disponíveis
        With lstDisponiveis
            .CommandText = strSQLDe
            .BindSource()
        End With
        'Carregando Lista Atribuídos
        With lstAtribuidos
            .CommandText = strSQLPara
            .BindSource()
        End With
        '-------------------------------------------
        'Else
        '    'Preenche lista de disponíveis
        '    If cboGruCodigo.SelectedValue <> 0 Then
        '        strSQL = "SELECT o.ObjCodigo, o.ObjNomeMenu FROM "
        '        strSQL &= "VW_OBJETO o INNER JOIN GRUPO_OBJETO g "
        '        strSQL &= "ON(o.ObjCodigo = g.ObjCodigo)"
        '        strSQL &= "WHERE AplCodigo = " & AplCodigo & " AND "
        '        strSQL &= "GruCodigo = " & cboGruCodigo.SelectedValue & "  AND o.ObjCodigo NOT IN "
        '        strSQL &= "(SELECT ObjCodigo FROM  VW_PERFIL WHERE UsuCodigo = " & intUsuCodigo & " AND AplCodigo = " & AplCodigo & ")"
        '    Else
        '        strSQL = "SELECT ObjCodigo, ObjNomeMenu FROM VW_OBJETO WHERE "
        '        strSQL &= " AplCodigo = " & AplCodigo
        '    End If
        '    'Carrega Lista Disponíveis
        '    With lstDisponiveis
        '        .CommandText = strSQL
        '        .BindSource()
        '    End With

        '    lstAtribuidos.ClearRows()
        'End If
    End Sub

    Private Sub CriaListas()
        With lstDisponiveis
            .AddColumn("ObjCodigo", "", False)
            .AddColumn("ObjNomeMenu", "Descrição")
            .AutoGenerateColumns = False
            .BoxTitle = "Objetos disponíveis"
            .ConnectionString = oDBAcesso.sConStr
        End With

        With lstAtribuidos
            .AddColumn("ObjCodigo", "", False)
            .AddColumn("ObjNomeMenu", "Descrição")
            .AutoGenerateColumns = False
            .BoxTitle = "Objetos atribuídos"
            .ConnectionString = oDBAcesso.sConStr
        End With
    End Sub

    Private Sub Mover(ByRef cxOrigem As ClsControlBox, _
                      ByRef cxDestino As ClsControlBox, _
                      Optional ByVal AllItems As Boolean = False)

        '----------------------------
        Dim lngRowIndex As Long = 0
        Dim blnExecutou As Boolean
        '----------------------------
        While lngRowIndex < cxOrigem.RowCount
            If AllItems Then
                blnExecutou = True
                cxDestino.ImportRow(cxOrigem.Row(0))
            Else
                If cxOrigem.IsSelectedRow(lngRowIndex) Then
                    blnExecutou = True
                    cxDestino.ImportRow(cxOrigem.Row(lngRowIndex))
                    If lngRowIndex > 0 Then lngRowIndex -= 1
                Else
                    lngRowIndex += 1
                End If
            End If
        End While
        '----------------------------
        If blnExecutou Then Call Editar()
        '----------------------------
    End Sub

#End Region

End Class
