Imports System.Text
Imports System.Web
Imports System.Xml
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.UI.WebControls
Imports System.Web.SessionState
Imports System.Net
Imports System.IO
Imports System.Threading



Public Class ClsAcesso
    Inherits WebControl


    Dim strSQL As StringBuilder

    Private cnn As OleDbConnection
    Private cmd As OleDbCommand
    Private dtr As OleDbDataReader
    Private dta As OleDbDataAdapter
    Private dts As DataSet
    Private tab As DataTable
    Private row As DataRow

    Private Session As HttpSessionState
    Private _AcessoObjeto As New Hashtable
    Private _PerCodigo As New Hashtable
    Private _cnnStr As String = String.Empty


    Public Sub New()
        Dim Response As HttpResponse = HttpContext.Current.Response
        Session = HttpContext.Current.Session

        If Session("AcessoObjeto") IsNot Nothing Then
            _AcessoObjeto = CType(Session("AcessoObjeto"), Hashtable)
        End If
    End Sub

#Region "Properties"
    Public Property cnnStr() As String
        Get
            If _cnnStr.Trim = "" Then
                _cnnStr = "Provider=SQLOLEDB.1;Password=adsadmin;Persist Security Info=True;User ID=netclient;Initial Catalog=netclient;Data Source=187.45.197.65"
            End If
            Return _cnnStr
        End Get
        Set(ByVal value As String)
            _cnnStr = value
        End Set
    End Property

    Public Property AcessoMsgScript() As Boolean
        Get
            If Session("AcessoMsgScript") Is Nothing Then Session("AcessoMsgScript") = True
            Return Session("AcessoMsgScript")
        End Get
        Set(ByVal value As Boolean)
            Session("AcessoMsgScript") = value
        End Set
    End Property

    Public Property AcessoAplCodigo() As String
        Get
            Return CheckTypeProperty(Session("AcessoAplCodigo"))            
        End Get
        Set(ByVal value As String)
            Session("AcessoAplCodigo") = value
        End Set
    End Property
    Public Property AcessoAplicativo() As String
        Get
            Return CheckTypeProperty(Session("AcessoAplicativo"))
        End Get
        Set(ByVal value As String)
            Session("AcessoAplicativo") = value
        End Set
    End Property

    Public WriteOnly Property PerCodigo() As Hashtable
        Set(ByVal value As Hashtable)
            Session("PerCodigo") = value
        End Set
    End Property

    Public ReadOnly Property AcessoObjeto(ByVal Key As String) As Hashtable
        Get
            Try
                Return _AcessoObjeto.Item(Key)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property


    Public ReadOnly Property AcessoObjeto(ByVal Index As Integer) As String
        Get
            Try
                Dim o As Object = Session("AcessoObjeto")
                If Not o Is Nothing Then
                    _AcessoObjeto = o
                    Return _AcessoObjeto.Item(Index)
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property
    Private WriteOnly Property AcessoObjeto() As Hashtable
        Set(ByVal value As Hashtable)
            Session("AcessoObjeto") = value
        End Set
    End Property
    Public Property AcessoUsuCodigo() As String
        Get
            Return CheckTypeProperty(Session("AcessoUsuCodigo"))
        End Get
        Set(ByVal value As String)
            Session("AcessoUsuCodigo") = value
        End Set
    End Property
    Public Property AcessoUsuUsuario() As String
        Get
            Return CheckTypeProperty(Session("AcessoUsuUsuario"))
        End Get
        Set(ByVal value As String)
            Session("AcessoUsuUsuario") = value
        End Set
    End Property
    Public Property AcessoStatusLogado() As Boolean
        Get
            Dim o As Object = Session("AcessoStatusLogado")
            If o Is Nothing Then Session("AcessoStatusLogado") = False
            Return Session("AcessoStatusLogado")
        End Get
        Set(ByVal value As Boolean)
            Session("AcessoStatusLogado") = value
        End Set
    End Property
    Public Property AcessoPageLogin() As String
        Get
            Dim o As String = Session("AcessoPageLogin")
            If o IsNot Nothing Then
                Return o
            Else
                Return "AcessoLogin.aspx"
            End If
        End Get
        Set(ByVal value As String)
            Session("AcessoPageLogin") = value
        End Set
    End Property
    Public Property AcessoDefaultPage() As String
        Get
            Dim o As String = Session("AcessoDefaultPage")
            If o IsNot Nothing Then
                Return o
            Else
                Return "default.aspx"
            End If
        End Get
        Set(ByVal value As String)
            Session("AcessoDefaultPage") = value
        End Set
    End Property
    Public Property AcessoUsuario() As String
        Get
            Return CheckTypeProperty(Session("AcessoUsuario"))
        End Get
        Set(ByVal value As String)
            Session("AcessoUsuario") = value
        End Set
    End Property
    Public Property AcessoSenha() As String
        Get
            Return CheckTypeProperty(Session("AcessoSenha"))
        End Get
        Set(ByVal value As String)
            Session("AcessoSenha") = value
        End Set
    End Property
    Public Property AcessoMensagem() As String
        Get
            Return CheckTypeProperty(Session("AcessoMensagens"))
        End Get
        Set(ByVal value As String)
            Session("AcessoMensagens") = value
        End Set
    End Property

    Public Property AcessoTarget() As String
        Get
            Return CheckTypeProperty(Session("AcessoTarget"))
        End Get
        Set(ByVal value As String)
            Session("AcessoTarget") = value
        End Set
    End Property

    Public Property AcessoDeniedRedirect() As String
        Get
            Return CheckTypeProperty(Session("AcessoDeniedRedirect"))
        End Get
        Set(ByVal value As String)
            Session("AcessoDeniedRedirect") = value
        End Set
    End Property

    Public Property AcessoEmpresa() As String
        Get
            Return CheckTypeProperty(Session("AcessoEmpresa"))
        End Get
        Set(ByVal value As String)
            Session("AcessoEmpresa") = value
        End Set
    End Property
#End Region


#Region "Methods"
    Public Function CheckTypeProperty(ByRef o As Object)
        If o IsNot Nothing Then
            Return CStr(o)
        Else
            Return ""
        End If
    End Function
    Private Sub AcessoObjetoAdd(ByVal Key As String, ByVal Value As String)
        _AcessoObjeto.Add(Key, Value)
        AcessoObjeto = _AcessoObjeto
    End Sub

    Private Sub PerCodigoAdd(ByVal Key As String, ByVal Value As String)
        _PerCodigo.Add(Key, Value)
        PerCodigo = _PerCodigo
    End Sub

    Public Sub VerificaLogin()
        Dim strRetMens As String = String.Empty
        Dim blnLogou As Boolean = False
        Dim Response As HttpResponse = HttpContext.Current.Response
        Try
            '------------------------------------
            If AcessoUsuario.Trim = String.Empty Then
                Throw New Exception("Usuário não informado!")
            ElseIf AcessoSenha.Trim = String.Empty Then
                Throw New Exception("Senha não informada!")
            End If
            '------------------------------------
            If AcessoStatusLogado Then
                _AcessoObjeto.Clear()
                AcessoObjeto = _AcessoObjeto
            End If
            '------------------------------------
            strRetMens = CheckCommandSQL()
            If strRetMens <> "" Then
                Throw New Exception(strRetMens)
            End If
            '------------------------------------
            cnn = New OleDbConnection(cnnStr)
            '------------------------------------
            strSQL = New StringBuilder
            With strSQL
                .Append("SELECT DISTINCT UsuCodigo, UsuUsuario, UsuSenha, EMPRESA ").Append(vbCrLf)
                .Append("FROM VW_USUARIO_PERFIL ").Append(vbCrLf)
                .Append("WHERE UsuUsuario = '").Append(AcessoUsuario & "'")
                If AcessoAplCodigo.Trim <> String.Empty Then
                    .Append(" AND AplCodigo = ").Append(AcessoAplCodigo).Append(vbCrLf)
                End If
            End With
            '------------------------------------
            cmd = New OleDbCommand(strSQL.ToString, cnn)
            dta = New OleDbDataAdapter(cmd)
            dts = New DataSet
            '------------------------------------
            Try
                cnn.Open()
                dta.Fill(dts, "VW_USUARIO_PERFIL")
            Finally
                cnn.Close()
            End Try
            '------------------------------------
            With dts
                If .Tables("VW_USUARIO_PERFIL").Rows.Count > 0 Then
                    row = .Tables("VW_USUARIO_PERFIL").Rows(0)
                    If row("UsuSenha").ToString.ToUpper <> AcessoSenha.ToUpper Then
                        Throw New Exception("Usuário ou senha inválido(s)!")
                    Else
                        '------------------------------------
                        'Informações do Usuário
                        AcessoUsuCodigo = row("UsuCodigo")
                        AcessoUsuUsuario = row("UsuUsuario")
                        AcessoEmpresa = row("EMPRESA")
                        Call LogarUsuario()
                        blnLogou = True
                    End If
                Else
                    Throw New Exception("Acesso negado!")
                End If
            End With
            '------------------------------------
        Catch ex As Exception
            AcessoMensagem = ex.Message
            If AcessoMsgScript Then
                With Response
                    .Clear()
                    .Write("<script language='javascript' type='text/javascript' id='j1'>")
                    .Write("    alert('" & ex.Message.Replace("'", " ").Replace(vbCrLf, " ") & "');")
                    .Write("    window.history.back(-1);")
                    .Write("</script>")
                End With
            Else
                AcessoMensagem = ex.Message
            End If
        Finally
            If blnLogou Then Response.Redirect(AcessoDefaultPage)
        End Try
    End Sub

    Private Sub LogarUsuario()
        '------------------------------------
        cnn = New OleDbConnection(cnnStr)
        '------------------------------------
        strSQL = New StringBuilder
        With strSQL
            .Append("SELECT").Append(vbCrLf)
            .Append("   AplCodigo,").Append(vbCrLf)
            .Append("   PerCodigo, ").Append(vbCrLf)
            .Append("   AplNome,").Append(vbCrLf)
            .Append("   UsuCodigo,").Append(vbCrLf)
            .Append("   UsuUsuario,").Append(vbCrLf)
            .Append("   UsuNomeCompleto,").Append(vbCrLf)
            .Append("   UsuSenha,").Append(vbCrLf)
            .Append("   ObjCodigo, ").Append(vbCrLf)
            .Append("   ObjNome,").Append(vbCrLf)
            .Append("   UsuNomeCompleto ").Append(vbCrLf)
            .Append("FROM VW_USUARIO_PERFIL").Append(vbCrLf)
            .Append("WHERE UsuCodigo = ").Append(AcessoUsuCodigo)
            .Append(" AND AplCodigo = ").Append(AcessoAplCodigo)
        End With
        '------------------------------------
        cmd = New OleDbCommand(strSQL.ToString, cnn)
        dta = New OleDbDataAdapter(cmd)
        dts = New DataSet()
        '------------------------------------
        Try
            cnn.Open()
            dta.Fill(dts, "VW_USUARIO_PERFIL")
        Finally
            cnn.Close()
        End Try
        '------------------------------------
        With dts.Tables(0)
            If .Rows.Count > 0 Then
                '-------------------------------------
                AcessoAplicativo = .Rows(0)("AplNome")
                Session("AcessoUsuNomeCompleto") = dts.Tables(0).Rows(0)("UsuNomeCompleto")
                '-------------------------------------
                For Each row In .Rows
                    AcessoObjetoAdd(row("ObjCodigo"), row("ObjNome").ToString.ToUpper)
                    PerCodigoAdd(row("ObjNome").ToString.ToUpper, row("PerCodigo"))
                Next row
                '-------------------------------------
                AcessoStatusLogado = True
                '-------------------------------------
            Else
                Throw New Exception("Erro ao carregar perfil do usuário!")
            End If
        End With
        '-------------------------------------
    End Sub

    Public Sub ChecaLogin(ByVal strPagina As String, Optional ByVal adEstatistica As Boolean = True)
        Dim Response As HttpResponse = HttpContext.Current.Response
        Dim pCodigo As Hashtable
        Dim codPerfil As Integer

        If AcessoUsuCodigo Is Nothing Then
            Response.Redirect(AcessoPageLogin)
        ElseIf AcessoUsuCodigo = String.Empty Then
            Response.Redirect(AcessoPageLogin)
        Else
            If AcessoStatusLogado Then
                If Not _AcessoObjeto.ContainsValue(strPagina.Trim.ToUpper) Then

                    With Response
                        .Clear()
                        .Write("<script language='javascript' type='text/javascript' id='j1'>")
                        .Write("    alert('Seu perfil não dá acesso à página requisitada.');")
                        If AcessoTarget.Trim <> String.Empty Then
                            .Write("    window.target = """ & AcessoTarget & """;")
                            If AcessoDeniedRedirect.Trim <> String.Empty Then
                                .Write("    window.location.href = """ & AcessoDeniedRedirect & """;")
                            Else
                                .Write("    window.history.go(-1);")
                            End If
                        Else
                            .Write("    window.history.go(-1);")
                        End If
                        .Write("</script>")
                    End With
                Else

                    'Insere Estatísticas
                    '--------------------------------------------
                    If adEstatistica Then

                        pCodigo = Session("PerCodigo")

                        If pCodigo.ContainsKey(strPagina.ToUpper) Then
                            codPerfil = pCodigo(strPagina.ToUpper)
                            sendStats(codPerfil)

                        End If
                    End If


                End If
            Else
                    Response.Redirect(AcessoPageLogin)
            End If
        End If
    End Sub

    Private Function CheckCommandSQL() As String
        'Checar Comando arbitrario 
        Dim vCmdArbitrario As Object
        Dim vArrUser As Object
        Dim vArrPws As Object
        Try
            vArrUser = Split(AcessoUsuario, " ")
            vArrPws = Split(AcessoSenha, " ")
            vCmdArbitrario = Split("select#insert#update#delete#drop#--#'", "#")
            For i As Integer = 0 To UBound(vCmdArbitrario)
                For l As Integer = 0 To UBound(vArrUser)
                    If LCase(vCmdArbitrario(i)) = LCase(vArrUser(l)) Then Throw New Exception("Comando Arbitrário")
                Next
                For l As Integer = 0 To UBound(vArrPws)
                    If LCase(vCmdArbitrario(i)) = LCase(vArrPws(l)) Then Throw New Exception("Comando Arbitrário")
                Next
            Next
            Return String.Empty
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    'Microsoft.XMLHTTP Equivalent
    Private Sub sendStats(ByVal perCodigo As Integer)
        Try
            Dim sParams As String
            Dim sBrowser As String
            Dim sVersao As String
            Dim sSistema As String

            With HttpContext.Current.Request.Browser
                sBrowser = .Browser
                sVersao = .MajorVersion
                sSistema = .Platform
            End With

            If InStr(UCase(sSistema), "WIN") And Not InStr(UCase(sSistema), "WINDOWS") Then
                sSistema = Replace(sSistema, "Win", "Windows ")
            End If

            sParams = perCodigo & "#" & sBrowser & "#" & sVersao & "#" & sSistema & "#" & obterNomePC()
            sParams = CorpCripto.ToHex(sParams)

            Dim httpRequest As HttpWebRequest = HttpWebRequest.Create(New Uri("http://10.0.0.238/estatistica/insereDados.aspx?sParams=" & sParams))
            httpRequest.Method = "GET"
            httpRequest.ContentType = "text/html"

            httpRequest.BeginGetResponse(New AsyncCallback(AddressOf processaResposta), "")

        Catch ex As Exception

        End Try
    End Sub

    Private Sub processaResposta(ByVal asynchronousResult As IAsyncResult)

    End Sub

    ''' <summary>
    ''' Recebe o IP e retorna o Nome do computador na rede
    ''' </summary>
    Public Function obterNomePC() As String
        Dim host As System.Net.IPHostEntry
        Dim strComputerName As String
        Dim ip As String

        ip = HttpContext.Current.Request.ServerVariables("REMOTE_ADDR")

        If String.IsNullOrEmpty(ip) Then
            strComputerName = "Desconhecido"
        Else
            Try
                'host = System.Net.Dns.GetHostEntry(ip)
                host = System.Net.Dns.GetHostByAddress(ip)
                strComputerName = host.HostName.ToUpper

            Catch ex As Exception
                strComputerName = ip
            End Try
        End If
        Return strComputerName
    End Function
#End Region

End Class


