Imports Microsoft.VisualBasic
Imports System.Xml

Public Class ClsPublic

#Region "Fields"
    '------------------------------------------
    'Código do HC
    Public Const HOSPITAL_CENTRAL As Int16 = 1
    '------------------------------------------
    'NÍVEIS DE ACESSO
    Public Const USUARIO As Char = "S"
    Public Const ESPECIALISTA As Char = "E"
    Public Const CALLCENTER As Char = "C"
    '------------------------------------------
    Public Shared strSQL As String = String.Empty
    Public Shared strbSQL As StringBuilder
    Public Shared TIPO_ACESSO As Char = String.Empty
    Public Shared AplCodigo As Integer
    '------------------------------------------
    Public Shared oDB As New ClsDB
    Public Shared oTool As New ClsTools
    '------------------------------------------
    Public Shared dta As DataTable
    Public Shared adp As OleDbDataAdapter
    Public Shared cnn As OleDbConnection
    Public Shared tra As OleDbTransaction
    Public Shared cmd As OleDbCommand
#End Region

#Region "Enumerations"

    Public Enum TpMsg
        Erro
        Info
    End Enum

    Public Enum TOPERACAO_BOTAO
        ENCAMINHAR = 1
        COMENTAR = 2
        BAIXAR = 3
    End Enum

    Public Enum TRETUNIDADE
        ID
        CODIGO
    End Enum

    Public Enum TTABELA
        TIPO_OS = 1
        CONTROLE_EQUIPAMENTO = 2
        TIPO_SOFTWARE = 3
        STATUS_OS = 5
        NIVEL_ACESSO_HELP = 6
        NIVEL_SATISFACAO = 7
        STATUS_ENCAMINHAMENTO = 8
    End Enum

    Public Enum TTIPO_OS 'TIPO_OS = 1        
        HARDWARE = 1
        SOFTWARE = 2
    End Enum

    Public Enum TTIPO_CONTROLE_EQUIP 'CONROLE_EQUIPAMENTO = 2
        SAI = 3
        PATRIMONIO = 4
        NUMSERIE = 5
        EXTERNO = 6
    End Enum

    Public Enum TTIPO_SOFTWARE
        SOFTWARE_INTERNO = 7
        SOFTWARE_TERCEIROS = 8
    End Enum

    Public Enum TSTATUS_OS
        ABERTA_USUARIO = 1
        ABERTA_CALL_CENTER = 2
        ABERTA_ESPECIALIZADO = 3
        SOLUCAO_CALL_CENTER = 4
        ENCERRAMENTO_CALL_CENTER = 5
        ENCERRAMENTO_USUARIO = 6
        ENCAMINHAMENTO_PE_CC = 7
        SOLUCAO_ESPECIALIZADO = 8
        ENCERRAMENTO_PE = 9
        INSERCAO_COMENTARIO = 10
        ENCAMINHAMENTO_PE_PE = 11
    End Enum

    Public Enum TSTATUS_GERAL_OS
        EM_ABERTO = 9
        EM_ANALISE = 20
        ENCERRADO_PI = 10
        ENCERRADO_US = 11
    End Enum

    Public Enum TTIPO_STATUS
        ABERTURA
        ENCAMINHAMENTO
        SOLUCAO
        COMENTARIO
        ENCERRAMENTO
    End Enum

    Public Enum TTIPO_NIVEL_ACESSO
        NIVEL_ACESSO_HELP = 6
    End Enum

    Public Enum TACESSO_HELP
        ADMINISTRADOR = 12
        TECNICO = 13
    End Enum

    Public Enum TSTATUS_ENCAMINHAMENTO
        ATUAL = 17
        NAO_ATUAL = 18
    End Enum

#End Region

#Region "Methods"

    Public Shared Function Quote(ByVal strValor As String) As String
        Return "'" & strValor.Trim.Replace("'", "''") & "'"
    End Function

    Public Shared Function RetStatus(ByVal Tipo As TTIPO_STATUS) As Short
        If TIPO_ACESSO = CALLCENTER Then
            If Tipo = TTIPO_STATUS.ABERTURA Then Return TSTATUS_OS.ABERTA_CALL_CENTER
            If Tipo = TTIPO_STATUS.ENCAMINHAMENTO Then Return TSTATUS_OS.ENCAMINHAMENTO_PE_CC
            If Tipo = TTIPO_STATUS.ENCERRAMENTO Then Return TSTATUS_OS.ENCERRAMENTO_CALL_CENTER
            If Tipo = TTIPO_STATUS.SOLUCAO Then Return TSTATUS_OS.SOLUCAO_CALL_CENTER
            If Tipo = TTIPO_STATUS.COMENTARIO Then Return TSTATUS_OS.INSERCAO_COMENTARIO
        ElseIf TIPO_ACESSO = USUARIO Then
            If Tipo = TTIPO_STATUS.ABERTURA Then Return TSTATUS_OS.ABERTA_USUARIO
            If Tipo = TTIPO_STATUS.ENCERRAMENTO Then Return TSTATUS_OS.ENCERRAMENTO_USUARIO
        ElseIf TIPO_ACESSO = ESPECIALISTA Then
            If Tipo = TTIPO_STATUS.ABERTURA Then Return TSTATUS_OS.ABERTA_ESPECIALIZADO
            If Tipo = TTIPO_STATUS.ENCAMINHAMENTO Then Return TSTATUS_OS.ENCAMINHAMENTO_PE_PE
            If Tipo = TTIPO_STATUS.ENCERRAMENTO Then Return TSTATUS_OS.ENCERRAMENTO_PE
            If Tipo = TTIPO_STATUS.COMENTARIO Then Return TSTATUS_OS.INSERCAO_COMENTARIO
            If Tipo = TTIPO_STATUS.SOLUCAO Then Return TSTATUS_OS.SOLUCAO_ESPECIALIZADO
        End If
    End Function

    Public Shared Function RetUnidades(Optional ByVal TipoRetorno As TRETUNIDADE = TRETUNIDADE.ID) As String
        Dim Session As HttpSessionState = HttpContext.Current.Session
        Dim strUnidades As String = String.Empty
        Dim colUnidades As List(Of String) = Session("UNIDADES")
        Dim intIndice As Int16 = TipoRetorno
        Dim strQuote As String = IIf(TipoRetorno = TRETUNIDADE.CODIGO, "'", "")
        '----------------------------------------
        For Each s As String In colUnidades
            strUnidades &= IIf(strUnidades <> String.Empty, ", ", "") & strQuote & s.Split("|")(intIndice) & strQuote
        Next
        '----------------------------------------
        Return strUnidades
        '----------------------------------------
    End Function

    Public Shared Sub ClearCache()
        Dim Response As HttpResponse = HttpContext.Current.Response
        With Response.Cache
            .SetNoServerCaching()
            .SetCacheability(System.Web.HttpCacheability.NoCache)
            .SetNoStore()
            .SetExpires(New DateTime(1900, 1, 1, 0, 0, 0, 0))
        End With
    End Sub

    Public Shared Sub CarregaPerifericos(ByRef grd As ClsGrid, ByRef Pagina As Object)
        Try
            '------------------------------------             
            For Each dtr As DataRow In dta.Rows
                With grd
                    .NewRow()
                    .AddRowCol("PER_ID", dtr("PER_ID"))
                    .AddRowCol("PER_DESCRICAO", dtr("PER_DESCRICAO"))
                    .AddRowCol("PET_DESCRICAO", dtr("PET_DESCRICAO"))
                    .AppendTableRow()
                End With
            Next
            '------------------------------------           
        Catch ex As Exception
            oTool.ShowMessage("Erro:" & vbCrLf & ex.Message, "Erro", Pagina)
        End Try
    End Sub


    Public Shared Sub GravaStatusOS(ByVal strORS_ID As String, _
                                    ByVal TipoStatus As TTIPO_STATUS, _
                                    Optional ByVal blnTransaction As Boolean = False)

        Dim Session As HttpSessionState = HttpContext.Current.Session
        '-----------------------------------
        strbSQL = New StringBuilder()
        With strbSQL
            .AppendLine("INSERT INTO TAB_STATUS_ORDEM_SERVICO (")
            .AppendLine("   SOS_ORDEM_SERVICO, ")
            .AppendLine("   SOS_STATUS, ")
            .AppendLine("   SOS_USUARIO_UNIDADE ")

            .AppendLine(") VALUES (")

            .Append(strORS_ID).AppendLine(", ")
            .Append(RetStatus(TipoStatus)).AppendLine(", ")
            .Append(Session("UsuarioUnidadeOrigem")).AppendLine(")")
        End With
        '---------------------------------------------
        If blnTransaction Then
            oDB.SetCommandSQL(strbSQL.ToString, cnn, tra)
        Else
            oDB.SetCommandSQL(strbSQL.ToString)
        End If
        '---------------------------------------------
    End Sub

    Public Shared Function RetStatusOS(ByVal intORS_ID As Integer) As TSTATUS_GERAL_OS
        Dim dtaStatus As New DataTable
        '------------------------------------
        strSQL = "SELECT ORS_STATUS FROM TAB_ORDEM_SERVICO "
        strSQL &= "WHERE ORS_ID = " & intORS_ID.ToString
        '------------------------------------
        dtaStatus = oDB.GetDataTable(strSQL)
        '------------------------------------
        Return CType(dtaStatus.Rows(0)("ORS_STATUS"), TSTATUS_GERAL_OS)
        '------------------------------------
    End Function

    Public Shared Sub Msg(ByRef ctr As Object, ByVal sMsg As String, Optional ByVal tm As TpMsg = TpMsg.Info)
        With ctr
            Select Case tm
                Case TpMsg.Info : .ForeColor = System.Drawing.Color.Blue
                Case TpMsg.Erro : .ForeColor = System.Drawing.Color.Red
            End Select
            .Text = sMsg
        End With
    End Sub
    '##########################################################################################
    Public Shared Sub LimpaCampos(ByRef Form As Object)
        For Each ctrl As Object In Form.Controls
            Select Case ctrl.GetType.ToString()
                Case "System.Web.UI.WebControls.TextBox"
                    CType(ctrl, TextBox).Text = String.Empty

                Case "System.Web.UI.WebControls.DropDownList"
                    CType(ctrl, DropDownList).SelectedIndex = -1

                Case "System.Web.UI.WebControls.RadioButton"
                    CType(ctrl, RadioButton).Checked = False

                Case "System.Web.UI.WebControls.CheckBox"
                    CType(ctrl, CheckBox).Checked = False
            End Select
        Next
    End Sub
    '##########################################################################################
    Public Shared Sub ToolBarState(ByRef pToolBar As Object, ByVal blnInclusao As Boolean)
        With pToolBar
            .ToolButtons(0).Enabled = blnInclusao 'Salvar
            .ToolButtons(2).Enabled = Not blnInclusao 'Editar
            .ToolButtons(3).Enabled = Not blnInclusao 'Excluir
        End With
    End Sub
    '##########################################################################################
    Public Shared Function GetTagXML(ByVal sTagName As String, ByVal sTagNameFind As String, ByVal sPath As String) As String
        Dim xDoc As New XmlDocument
        Dim xXmlNodeList As XmlNodeList
        Dim sResult As String = ""
        Dim i As Integer
        xDoc.Load(sPath)
        xXmlNodeList = xDoc.GetElementsByTagName(sTagName)
        With xXmlNodeList
            If .Count > 0 Then
                With .Item(0).ChildNodes
                    For i = 0 To .Count - 1
                        If sTagNameFind <> "" Then
                            If LCase(sTagNameFind) = LCase(.Item(i).Name) Then
                                sResult = .Item(i).InnerText
                                Exit For
                            End If
                        End If
                    Next
                End With
            End If
        End With
        Return sResult
    End Function
    '##########################################################################################
    Public Shared Function SetMoneyDB(ByVal sVal As String) As String
        If sVal <> "" Then
            SetMoneyDB = Replace(Replace(sVal, ".", ""), ",", ".")
        Else
            SetMoneyDB = "Null "
        End If
    End Function
    '##########################################################################################
    Public Shared Function SetNumericDB(ByVal sVal As String) As String
        If sVal <> "" Then
            SetNumericDB = sVal
        Else
            SetNumericDB = "Null "
        End If
    End Function
    '##########################################################################################
    Public Shared Function RPA(ByVal sExpresion As String) As String
        RPA = Replace(sExpresion, "'", "''")
    End Function
    '##########################################################################################
    Public Shared Function RPADB(ByVal sExpresion As String) As String
        If sExpresion <> "" Then
            RPADB = "'" & Replace(sExpresion, "'", "''") & "'"
        Else
            RPADB = "Null "
        End If
    End Function
    '#####################################################################################
    Public Shared Function GetDateFormatDB(ByVal dt As String, Optional ByVal tp As String = "2") As String
        Dim sGetDateFormatDB As String = ""
        If dt <> "" Then
            Select Case tp
                Case "1" 'DD/MM/YYYY
                    sGetDateFormatDB = "'" & Right("0" & Day(dt), 2) & "/" & Right("0" & Month(dt), 2) & "/" & Year(dt) & "'"
                Case "2" 'YYYY-MM-DD
                    sGetDateFormatDB = "'" & Year(dt) & "-" & Right("0" & Month(dt), 2) & "-" & Right("0" & Day(dt), 2) & "'"
            End Select
        Else
            sGetDateFormatDB = "Null "
        End If
        Return sGetDateFormatDB
    End Function
    '#####################################################################################
    Public Shared Function GetDateFormat(ByVal dt As String, ByVal tp As String) As String
        Dim sGetDateFormat As String = ""
        If dt <> "" Then
            Select Case tp
                Case 1 'DD/MM/YYYY
                    sGetDateFormat = Right("0" & Day(dt), 2) & "/" & Right("0" & Month(dt), 2) & "/" & Year(dt)
                Case 2 'YYYY-MM-DD
                    sGetDateFormat = Year(dt) & "-" & Right("0" & Month(dt), 2) & "-" & Right("0" & Day(dt), 2)
            End Select
        End If
        Return sGetDateFormat
    End Function

#End Region

End Class
