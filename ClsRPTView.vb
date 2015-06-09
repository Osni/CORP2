Imports System
Imports System.IO
Imports System.Data
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Collections.Specialized
Imports System.Collections.ObjectModel
Imports Microsoft.VisualBasic


#Region "RPTVIEW"

<ToolboxData("<{0}:ClsRPTView runat=server></{0}:ClsRPTView>")> _
Public Class ClsRPTView
    Inherits WebControl

#Region "Constructores/Destructores"

    Public Sub New()
        oDB = New ClsDB
        '-------------------------------
        'Alinhamento das colunas
        With HASH_TAlinhamento
            .Add(CShort(Coluna.TAlinhamento.CENTER), "center")
            .Add(CShort(Coluna.TAlinhamento.LEFT), "left")
            .Add(CShort(Coluna.TAlinhamento.RIGHT), "right")
        End With
    End Sub

#End Region


#Region "Fields"
    '----------------------------------------
    Private dta As DbDataAdapter
    Private oDB As ClsDB
    Private mTableDataSorce As DataTable
    '----------------------------------------
    Private mRelCodigo As Integer = 0
    Private mEstCodigo As Integer = 0
    Private mGruRelCodigo As Integer = 0
    Private mOriDadCodigo As Integer = 0
    Private mNumPagina As Short = 1
    Private mNumLinhas As Short = 0
    Private mMaxQtdeLinhas As Short = 0
    '---------------------------------------- 
    Private bNotSetResumo As Boolean = False
    '---------------------------------------- 
    Private mRelRodape As String = String.Empty
    Private mRelNome As String = String.Empty
    Private mRelSubNome As String = String.Empty
    Private mRelURLExterno As String = String.Empty
    Private mRelURLLogo As String = String.Empty
    Private strMantemStringExcel As String = String.Empty
    '----------------------------------------             
    Private mIteTabCodigoTipoConsulta As Short = 0
    Private mIteTabCodigoEstiloGrupo As Short = 0
    Private mIteTabCodigoTipoRel As Short = 0
    Private mIteTabCodigoOrientacao As Short = 0
    Private mIteTabCodigoGerarRPT As Short = 0
    '----------------------------------------    
    Private mCellsFormat As New Collection
    Private mColunaGrupo As New Collection(Of String)
    Private mColunasDetalhe As New Colunas
    Private mColuna As Coluna
    '----------------------------------------            
    Private HASH_TAlinhamento As New Hashtable
    '----------------------------------------
    Private mQuebrarPorGrupo As Boolean = False
    Private mRelatorioPaginado As Boolean = False
    Private mQuebrarPaginaExportacao As Boolean = False
    '----------------------------------------
    Private oFunc As New ClsFunctions
    '----------------------------------------

#End Region


#Region "Enumerations"

    Public Enum TPageLayout As Short
        QTD_LINHAS_PAISAGEM = 33
        QTD_LINHAS_RETRATO = 50
        '----------------------------------------
        QTD_CARAC_PAISAGEM = 150
        QTD_CARAC_RETRATO = 90
        '----------------------------------------        
    End Enum

    Public Enum TTipoRel As Short
        INTERNO = 5
        EXTERNO = 6
    End Enum

    Public Enum TTipoConsulta
        SQL = 1
        STORED_PROCEDURE = 2
    End Enum

    '??????????????????????????????????
    Private Enum TEstiloGrupo As Short
        VERTICAL = 3
        HORIZONTAL = 4
    End Enum
    '??????????????????????????????????

    Public Enum TOrientacao
        CUSTOMIZADA = 0
        PAISAGEM = 7
        RETRATO = 8
    End Enum

    Public Enum TGerarRPT
        SIM = 9
        NÃO = 10
    End Enum

#End Region


#Region "Events"

    Private Sub ClsRPTView_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        Call GenerateRptStyle()
    End Sub

    Private Sub ClsRPTView_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        If RelatorioPaginado Then
            Dim Response As HttpResponse = HttpContext.Current.Response
            With Response
                .Write("<script>" & vbCrLf)
                .Write(" try { " & vbCrLf)
                .Write("   var hdnMax = document.getElementById('hdnMaxPage');" & vbCrLf)
                .Write("   var hdnPage = document.getElementById('hdnPage');" & vbCrLf)
                .Write("   var txtPage = document.getElementById('txtNumPag');" & vbCrLf)
                .Write("   var lblMax = document.getElementById('lblMaxPage');" & vbCrLf)
                .Write("   txtPage.value = hdnPage.value;" & vbCrLf)
                .Write("   lblMax.innerText = hdnMax.value;" & vbCrLf)
                .Write("}catch(e){}" & vbCrLf)
                .Write("</script>" & vbCrLf)
            End With
        End If
    End Sub

    Protected Overrides Sub RenderContents(ByVal output As HtmlTextWriter)
        Me.Style.Clear()
        If mColunasDetalhe.Count = 0 Then
            Me.Controls.Clear()
            If Me.DesignMode Then Me.Controls.Add(New LiteralControl("<div style=""width:100%"">" & Me.ClientID & "</div>"))
        Else
            If Exportar Then
                Call GerarExportacao()  'Caso seja exportação
            Else
                Call ShowRelatorio()    'Constroi o relatório               
            End If
        End If
        MyBase.RenderContents(output)
    End Sub

#End Region


#Region "Classes"

    Friend Class MyStyle
        Inherits Style

        Private mStyles As Hashtable

        ''' <summary>
        ''' Instância de um objeto Style com os atributos definidos em pStyle.
        ''' </summary>
        ''' <param name="pStyles">Hashtable contendos atributos para o objeto style.
        ''' Key = atributo style; Value = Valor do atributo style.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pStyles As Hashtable)
            Me.New()
            mStyles = pStyles
        End Sub

        Public Sub New()
            MyBase.New()
        End Sub

        Protected Overrides Sub FillStyleAttributes(ByVal attributes As System.Web.UI.CssStyleCollection, ByVal urlResolver As System.Web.UI.IUrlResolutionService)
            MyBase.FillStyleAttributes(attributes, urlResolver)
            For Each s As String In mStyles.Keys
                attributes(s) = mStyles(s)
            Next
        End Sub
    End Class

    Friend Class TextWriterCorp
        Inherits TextWriter

        Public Overrides ReadOnly Property Encoding() As System.Text.Encoding
            Get
                Return New System.Text.ASCIIEncoding
            End Get
        End Property
    End Class

#End Region


#Region "Properties"

    ''' <summary>
    ''' Ocultar a caixa de parâmetros ao carregar o relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OcultarParamAoCarregar() As Boolean
        Get
            If ViewState("OcultarParamAoCarregar") Is Nothing Then ViewState("OcultarParamAoCarregar") = False
            Return ViewState("OcultarParamAoCarregar")
        End Get
        Set(ByVal value As Boolean)
            ViewState("OcultarParamAoCarregar") = value
        End Set
    End Property

    ''' <summary>
    ''' Permite que o relatório agrupado seja exportado para Excel 
    ''' também agrupado. Valor padrão é "Falso".
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ExportarAgrupado() As Boolean
        Get
            If ViewState("ExportarAgrupado") Is Nothing Then ViewState("ExportarAgrupado") = False
            Return ViewState("ExportarAgrupado")
        End Get
        Set(ByVal value As Boolean)
            ViewState("ExportarAgrupado") = value
        End Set
    End Property

    ''' <summary>
    ''' Se, ao realizar uma exportação, sejam inseridas quebras de
    ''' página no arquivo exportado, de acordo com a quantidade de linhas 
    ''' definida no relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property QuebrarPaginaExportacao() As Boolean
        Get
            Dim o As Object = ViewState("QuebrarPaginaExportacao")
            '-----------------------------------------
            If o Is Nothing Then ViewState("QuebrarPaginaExportacao") = False
            Return ViewState("QuebrarPaginaExportacao")
        End Get
        Set(ByVal value As Boolean)
            ViewState("QuebrarPaginaExportacao") = value
        End Set
    End Property

    ''' <summary>
    ''' Se o relatório possui resumo.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property TemResumo() As Boolean
        Get
            Dim s As Object = ViewState("TemResumo")
            If s Is Nothing Then ViewState("TemResumo") = False
            Return ViewState("TemResumo")
        End Get
        Set(ByVal value As Boolean)
            ViewState("TemResumo") = value
        End Set
    End Property

    ''' <summary>
    ''' Se é uma exportação ou uma exibição de relatório no navegador.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Exportar() As Boolean
        Get
            Dim o As Object = ViewState("Exportar")
            '-----------------------------------------
            If o Is Nothing Then ViewState("Exportar") = False
            Return ViewState("Exportar")
        End Get
        Set(ByVal value As Boolean)
            ViewState("Exportar") = value
        End Set
    End Property

    ''' <summary>
    ''' Retorna e configura a quantidade máxima de caracteres de 
    ''' acordo com a orientação do papel.
    ''' </summary>    
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Caso seja escolhida uma orientação customizada e não seja 
    ''' informado valor para MaxQtdeLinhas, será assumido quantidade de linhas
    ''' do formato Paisagem.</remarks>
    Public Property MaxQtdeLinhas() As Short
        Get
            Select Case CType(mIteTabCodigoOrientacao, TOrientacao)
                Case TOrientacao.PAISAGEM
                    Return TPageLayout.QTD_LINHAS_PAISAGEM

                Case TOrientacao.RETRATO
                    Return TPageLayout.QTD_LINHAS_RETRATO

                Case TOrientacao.CUSTOMIZADA
                    If mMaxQtdeLinhas <> 0 Then mMaxQtdeLinhas = TPageLayout.QTD_LINHAS_PAISAGEM
                    Return mMaxQtdeLinhas

            End Select
        End Get
        Set(ByVal value As Short)
            If CType(mIteTabCodigoOrientacao, TOrientacao) = TOrientacao.CUSTOMIZADA Then
                mMaxQtdeLinhas = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Controle de Número de linhas da página do relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property NumLinhas() As Short
        Get
            Return mNumLinhas
        End Get
        Set(ByVal value As Short)
            mNumLinhas = value
        End Set
    End Property

    ''' <summary>
    ''' Número da página atual.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property NumPagina() As Short
        Get
            Return mNumPagina
        End Get
        Set(ByVal value As Short)
            mNumPagina = value
        End Set
    End Property

    ''' <summary>
    ''' Código do relatório a ser carregado pela aplicação.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RelCodigo() As Integer
        Get
            Return mRelCodigo
        End Get
        Set(ByVal value As Integer)
            mRelCodigo = value
        End Set
    End Property

    ''' <summary>
    ''' Código do grupo em que o relatório está gravado.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property GruRelCodigo() As Integer
        Get
            Return mGruRelCodigo
        End Get
        Set(ByVal value As Integer)
            mGruRelCodigo = value
        End Set
    End Property

    ''' <summary>
    ''' Código da origem de dados.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CodigoOrigemDados() As Integer
        Get
            Return mOriDadCodigo
        End Get
        Set(ByVal value As Integer)
            mOriDadCodigo = value
        End Set
    End Property

    ''' <summary>
    ''' Valor verdadeiro/falso indicando se o relatório será exibido com 
    ''' características que permitam páginação, 
    ''' exibindo inicialmente a primeira página e ocultando
    ''' as páginas seguintes.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RelatorioPaginado() As Boolean
        Get
            Return mRelatorioPaginado
        End Get
        Set(ByVal value As Boolean)
            mRelatorioPaginado = value
        End Set
    End Property

    ''' <summary>
    ''' Se o relatório irá realizar quebra de página por grupo.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property QuebrarPorGrupo() As Boolean
        Get
            Return mQuebrarPorGrupo
        End Get
        Set(ByVal value As Boolean)
            mQuebrarPorGrupo = value
        End Set
    End Property

    ''' <summary>
    ''' Permite obter/definir o provider de dados framework será utilizado.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Provider() As ClsDB.T_PROVIDER
        Get
            If ViewState("Provider") Is Nothing Then ViewState("Provider") = ClsDB.T_PROVIDER.OLEDB
            Return ViewState("Provider")
        End Get
        Set(ByVal value As ClsDB.T_PROVIDER)            
            ViewState("Provider") = value
        End Set
    End Property

    ''' <summary>
    ''' String de Conexão para a base de dados desejada.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ConnectionString() As String
        Get
            Return TrataViewState("OriDadStringConexao")
        End Get
        Set(ByVal value As String)
            ViewState("OriDadStringConexao") = value
        End Set
    End Property

    ''' <summary>
    ''' Cláusula Where para o comando SQL de carga dos dados do relatório.    
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ClausulaWhere() As String
        Get
            Return TrataViewState("Where")
        End Get
        Set(ByVal value As String)
            ViewState("Where") = value
        End Set
    End Property

    ''' <summary>
    ''' Título a ser exibido no cabeçalho do relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TituloRelatorio() As String
        Get
            Return TrataViewState("RelTitulo")
        End Get
        Set(ByVal value As String)
            ViewState("RelTitulo") = value
        End Set
    End Property

    ''' <summary>
    ''' Subtítulo a ser exibido no cabeçalho do relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SubTituloRelatorio() As String
        Get
            Return TrataViewState("RelSubTitulo")
        End Get
        Set(ByVal value As String)
            ViewState("RelSubTitulo") = value
        End Set
    End Property

    ''' <summary>
    ''' Nome do Relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property NomeRelatorio() As String
        Get
            Return mRelNome
        End Get
        Set(ByVal value As String)
            mRelNome = value
        End Set
    End Property

    ''' <summary>
    ''' Comando SQL para carregar as informações do relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SQLQuery() As String
        Get
            Return TrataViewState("RelSQL")
        End Get
        Set(ByVal value As String)
            ViewState("RelSQL") = value
        End Set
    End Property

    ''' <summary>
    ''' Caminho do relatório externo.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property URLExterno() As String
        Get
            Return mRelURLExterno
        End Get
        Set(ByVal value As String)
            mRelURLExterno = value
        End Set
    End Property

    ''' <summary>
    ''' Texto a ser exibido juntamente com a data atual e o número de página.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TextoRodape() As String
        Get
            Return mRelRodape
        End Get
        Set(ByVal value As String)
            mRelRodape = value
        End Set
    End Property

    ''' <summary>
    ''' Permite obter/definir o tipo de consulta, se SQL ou Stored Procedure.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TipoConsulta() As TTipoConsulta
        Get
            Return TrataProp(mIteTabCodigoTipoConsulta, TTipoConsulta.SQL)
        End Get
        Set(ByVal value As TTipoConsulta)
            Select Case value
                Case TTipoConsulta.SQL, TTipoConsulta.STORED_PROCEDURE
                    mIteTabCodigoTipoConsulta = value

                Case Else
                    Throw New Exception("Valor informado não corresponde é inválido!")

            End Select
        End Set
    End Property

    '??????????????????????????????????
    Private Property EstiloGrupo() As TEstiloGrupo
        Get
            Return TrataProp(mIteTabCodigoEstiloGrupo, TEstiloGrupo.VERTICAL)
        End Get
        Set(ByVal value As TEstiloGrupo)
            Select Case value
                Case TEstiloGrupo.HORIZONTAL, TEstiloGrupo.VERTICAL
                    mIteTabCodigoEstiloGrupo = value

                Case Else
                    Throw New Exception("Valor informado não corresponde é inválido!")

            End Select
        End Set
    End Property
    '??????????????????????????????????

    ''' <summary>
    ''' Permite obter/definir  se o relatório é externo ou interno.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TipoRelatorio() As TTipoRel
        Get
            Return TrataProp(mIteTabCodigoTipoRel, TTipoRel.INTERNO)
        End Get
        Set(ByVal value As TTipoRel)
            Select Case value
                Case TTipoRel.EXTERNO, TTipoRel.INTERNO
                    mIteTabCodigoTipoRel = value
                Case Else
                    Throw New Exception("Valor informado não corresponde é inválido!")

            End Select
        End Set
    End Property

    ''' <summary>
    ''' Permite configurar a orientação do papel.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Orientacao() As TOrientacao
        Get
            Return TrataProp(mIteTabCodigoOrientacao, TOrientacao.PAISAGEM)
        End Get
        Set(ByVal value As TOrientacao)
            Select Case value
                Case TOrientacao.PAISAGEM, TOrientacao.RETRATO, TOrientacao.CUSTOMIZADA
                    mIteTabCodigoOrientacao = value

                Case Else
                    Throw New Exception("Valor informado não corresponde é inválido!")

            End Select
        End Set
    End Property

    '''' <summary>    
    '''' </summary>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Property GerarRPT() As TGerarRPT
    '    Get
    '        Return TrataProp(mIteTabCodigoGerarRPT, TGerarRPT.NÃO)
    '    End Get
    '    Set(ByVal value As TGerarRPT)
    '        Select Case value
    '            Case TGerarRPT.NÃO, TGerarRPT.SIM
    '                mIteTabCodigoGerarRPT = value

    '            Case Else
    '                Throw New Exception("Valor informado não corresponde é inválido!")

    '        End Select
    '    End Set
    'End Property

    ''' <summary>
    ''' Caminho da imagem do cabeçalho do relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Caso não seja informado, o título do relatório irá tomar
    ''' todo o espaço do cabeçalho</remarks>
    Public Property URLLogo() As String
        Get
            Return mRelURLLogo
        End Get
        Set(ByVal value As String)
            mRelURLLogo = value
        End Set
    End Property

    Public Property ToolTipRelatorio() As String
        Get
            Return TrataViewState("RelToolTip")
        End Get
        Set(ByVal value As String)
            ViewState("RelToolTip") = value
        End Set
    End Property

    Public Property ColunasDetalhe() As Colunas
        Get
            Return mColunasDetalhe
        End Get
        Set(ByVal value As Colunas)
            mColunasDetalhe = value
        End Set
    End Property

    ''' <summary>
    ''' Permite informar uma fonte de dados (DataTable) 
    ''' para carga do relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TableDataSorce() As DataTable
        Get
            If mTableDataSorce Is Nothing Then
                mTableDataSorce = New System.Data.DataTable()
            End If
            Return mTableDataSorce
        End Get
        Set(ByVal value As DataTable)
            mTableDataSorce = value
        End Set
    End Property

    Private Function TrataViewState(ByVal pKey As String) As String
        '----------------------------------
        Dim o As Object = ViewState(pKey)
        '----------------------------------
        If o IsNot Nothing Then
            Return o
        Else
            Return String.Empty
        End If
        '----------------------------------
    End Function

#End Region


#Region "Methods"

    Private Sub GenerateRptStyle()
        Dim pgHandler As Page = CType(Me.Context.CurrentHandler, Page)
        Dim isStyle As UI.IStyleSheet = pgHandler.Header.StyleSheet
        Dim newStyle As Style
        Dim colStyles As Hashtable
        '----------------------------------------
        'FORMATAÇÃO DO CABEÇALHO
        newStyle = New Style
        With newStyle
            .Width = Unit.Percentage(100)
            .Height = Unit.Pixel(45)
            .BorderStyle = BorderStyle.None
        End With
        isStyle.CreateStyleRule(newStyle, Nothing, "#tabRptHeader")
        '----------------------------------------              
        'FORMATAÇÃO DE PÁGINA
        newStyle = New Style
        newStyle.Width = Unit.Percentage(100)
        isStyle.CreateStyleRule(newStyle, Nothing, "#rptTabDetalhe")
        'RÓTULO DO DETALHE
        colStyles = New Hashtable
        colStyles.Add("vertical-align", "middle")
        newStyle = New MyStyle(colStyles)
        With isStyle
            .CreateStyleRule(newStyle, Nothing, "#rptTabDetalhe th")
            .CreateStyleRule(newStyle, Nothing, "#rptTabDetalhe td")
        End With
        '----------------------------------------              
        'FORMATAÇÃO DE PÁGINA
        newStyle = New Style
        With newStyle
            .Width = Unit.Percentage(100)
            .BorderStyle = BorderStyle.None
        End With
        '-----------------------------------        
        isStyle.CreateStyleRule(newStyle, Nothing, ".rptPaginas")
        isStyle.CreateStyleRule(newStyle, Nothing, "#rptPagina")
        '----------------------------------------              
        'LINHA EM BRANCO
        newStyle = New Style
        With newStyle
            .BorderStyle = BorderStyle.Solid
            .BorderWidth = Unit.Pixel(1)
            .BorderColor = Drawing.ColorTranslator.FromHtml("#FFFFFF")
        End With
        isStyle.CreateStyleRule(newStyle, Nothing, "#rptEmpty")
        '----------------------------------------
        'RODAPE
        newStyle = New Style
        newStyle.Width = Unit.Percentage(100)
        isStyle.CreateStyleRule(newStyle, Nothing, "#tabRptRodape")
        '----------------------------------------
        'QUEBRA DE PÁGINA
        colStyles = New Hashtable
        colStyles.Add("page-break-after", "always")
        newStyle = New MyStyle(colStyles)
        isStyle.CreateStyleRule(newStyle, Nothing, ".rptQuebra")
        '----------------------------------------
        Call GenerateStyleDetail()
        '----------------------------------------
    End Sub

    Private Sub GenerateStyleDetail()
        Dim pgHandler As Page = CType(Me.Context.CurrentHandler, Page)
        Dim isStyle As UI.IStyleSheet = pgHandler.Header.StyleSheet
        Dim newStyle As MyStyle
        Dim colStyles As Hashtable
        Dim intCellCount As Integer
        Dim strAlign As String
        Dim strWidth As String
        '----------------------------------------
        For Each colDet As Coluna In mColunasDetalhe.Values
            If colDet.TipoColuna = Coluna.TTipoColuna.DETALHE Then
                colStyles = New Hashtable
                If colDet.QuebrarTexto = Coluna.TQuebrarTexto.NAO Then
                    colStyles.Add("white-space", "nowrap")
                End If
                '----------------------------------------                
                strAlign = HASH_TAlinhamento(CShort(colDet.Alinhamento)) & ""
                strWidth = colDet.ColumnSize.ToString
                '----------------------------------------
                If strAlign.Trim <> String.Empty Then
                    colStyles.Add("text-align", strAlign)
                End If
                '----------------------------------------
                If strWidth.Trim <> String.Empty Then
                    colStyles.Add("width", strWidth)
                End If
                '----------------------------------------
                newStyle = New MyStyle(colStyles)
                intCellCount += 1
                isStyle.CreateStyleRule(newStyle, Nothing, "#cRpt" & intCellCount.ToString)
                If colDet.TipoResumo = Coluna.TTipoResumo.SOMA Then
                    isStyle.CreateStyleRule(newStyle, Nothing, "#" & colDet.Nome & "_s")
                End If
            End If
            '----------------------------------------
            If colDet.HeaderStyle.Count > 0 Then
                newStyle = New MyStyle(colDet.HeaderStyle)
                isStyle.CreateStyleRule(newStyle, Nothing, "#" & colDet.Nome)
            End If
        Next
        '----------------------------------------
    End Sub

    Private Function TrataProp(ByVal intValProp As Integer, ByVal intValDefault As Integer) As Integer
        If intValProp = 0 Then
            Return intValDefault
        Else
            Return intValProp
        End If
    End Function

    Private Sub SetColunasGrupo()
        For Each col As Coluna In ColunasDetalhe.Values
            If col.TipoColuna = Coluna.TTipoColuna.GRUPO Then mColunaGrupo.Add(col.Nome)
        Next
    End Sub

    ''' <summary>
    ''' Realiza a quebra de página do relatório.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function QuebraPagina() As Object
        '-------------------------------------------
        mNumLinhas = 0
        mNumPagina += 1
        '-------------------------------------------        
        Return New LiteralControl("<p class=""rptQuebra"">&nbsp;</p>")
    End Function

    ''' <summary>
    ''' Gera o cabeçalho do relatório.
    ''' </summary>
    ''' <returns>Retorna uma tabela com as informações do cabeçalho. 
    ''' Caso tenha sido passado uma imagem de logo, essa será incluída no relatório ao
    ''' lado superior esquerdo.</returns>
    ''' <remarks></remarks>
    Private Function SetCabecalho() As TableRow
        Dim tabCabec As New Table
        Dim imgLogo As New Image
        Dim Row As New TableRow
        '--------------------------------------------------        
        With tabCabec
            .ID = "tabRptHeader"
            .CellPadding = 0
            .CellSpacing = 0
        End With
        '--------------------------------------------------
        With Row.Cells
            If mRelURLLogo.Trim <> String.Empty And Not Exportar Then
                With imgLogo
                    .ImageUrl = URLLogo
                    .BorderStyle = WebControls.BorderStyle.None
                End With
                .Add(oFunc.AppendToCell(imgLogo, strWidth:="5%"))
            End If
            '--------------------------------------------------            
            .Add(oFunc.RetCellFormat("", IIf(TituloRelatorio.Trim = String.Empty, "Relatório", TituloRelatorio) & IIf(SubTituloRelatorio.Trim <> String.Empty, "<BR/><span id='rptSubTitulo'>" & SubTituloRelatorio & "</span>", ""), HorizontalAlign.Center, "95%"))
        End With
        '--------------------------------------------------
        tabCabec.Rows.Add(Row)
        '--------------------------------------------------
        Row = New TableRow
        Row.Cells.Add(oFunc.AppendToCell(tabCabec, "rptTitulo"))
        '--------------------------------------------------
        mNumLinhas += 3         'incrementa linhas
        '--------------------------------------------------
        Return Row
    End Function

    ''' <summary>
    ''' Gera o grupo de acordo com a linha passada.
    ''' </summary>
    ''' <param name="row">DataRow contendo as informações do grupo.</param>
    ''' <returns>Uma tabela contendo as informações do grupo.</returns>
    ''' <remarks></remarks>
    Private Function SetGrupo(ByVal row As DataRow) As TableRow
        Dim RowGrupo As New TableRow
        Dim CellGrupo As New TableHeaderCell
        Dim lngMaxLength As Long = 0
        '------------------------------------------
        RowGrupo.ID = "tabRptGrupo"
        '------------------------------------------
        For Each NameCol As String In mColunaGrupo
            If ColunasDetalhe(NameCol).ColumnMaxLength > 0 Then
                CellGrupo.Text &= IIf(CellGrupo.Text.Trim <> String.Empty, " - ", "") & IIf(ColunasDetalhe(NameCol).Titulo.Trim <> "", ColunasDetalhe(NameCol).Titulo.Trim & ": ", "") & _
                                Mid(row(NameCol).ToString, 1, ColunasDetalhe(NameCol).ColumnMaxLength)
            Else
                CellGrupo.Text &= IIf(CellGrupo.Text.Trim <> String.Empty, " - ", "") & IIf(ColunasDetalhe(NameCol).Titulo.Trim <> "", ColunasDetalhe(NameCol).Titulo.Trim & ": ", "") & row(NameCol).ToString
            End If
        Next
        '------------------------------------------
        mNumLinhas += 1          'incrementa linhas
        '------------------------------------------
        RowGrupo.Cells.Add(CellGrupo)
        '------------------------------------------        
        Return RowGrupo
    End Function

    ''' <summary>
    ''' Gera o rodapé da página.
    ''' </summary>
    ''' <returns>Retorna uma tabela contendo as informações do rodapé.</returns>
    ''' <remarks></remarks>
    Private Function SetRodape(Optional ByVal blnFecharBorderUp As Boolean = False) As Table
        Dim RowRodape As New TableRow
        Dim TabRodape As New Table
        '------------------------------------
        With RowRodape.Cells
            .Add(oFunc.RetCellFormat("", "Pág.: " & mNumPagina & "&nbsp;&nbsp;", , "33%"))
            .Add(oFunc.RetCellFormat("", FormatDateTime(Date.Now, DateFormat.ShortDate), HorizontalAlign.Center, "33%"))
            .Add(oFunc.RetCellFormat("", "&nbsp;", , "34%"))
        End With
        '------------------------------------
        mNumLinhas += 1 'incrementa linhas
        '------------------------------------
        With TabRodape
            .ID = "tabRptRodape"
            .CellPadding = 0
            .CellSpacing = 0
            .Rows.Add(RowRodape)
        End With
        '------------------------------------
        Return TabRodape
    End Function

    ''' <summary>
    '''  Retorna uma tabela com as informações do resumo parcial ou total.
    ''' </summary>    
    ''' <returns>Linha contendo as informações de resumo.</returns>
    ''' <remarks></remarks>
    Private Function SetResumo() As TableRow
        If bNotSetResumo Then Return Nothing
        '------------------------------------
        Dim RowResumo As New TableRow
        Dim Cell As TableCell
        Dim intColResumo As Integer
        '------------------------------------
        RowResumo.ID = "ResumoColuna"
        '------------------------------------   
        For Each col As Coluna In mColunasDetalhe.Values
            If Not mColunaGrupo.Contains(col.Nome) Then
                intColResumo += 1
                With RowResumo.Cells
                    If col.TipoResumo = Coluna.TTipoResumo.RESUMO Then
                        .Add(oFunc.RetCellFormat("", col.RotuloResumoPag))
                    ElseIf col.TipoResumo = Coluna.TTipoResumo.SOMA Then
                        Cell = oFunc.RetCellFormat(col.Nome & "_s", IIf(col.Formato.Trim <> "", FormatarColuna(col.ResumoSubTotal, col.Formato), col.ResumoSubTotal))
                        .Add(Cell)
                    ElseIf TemResumo Then
                        .Add(oFunc.RetCellFormat("", ""))
                    End If
                    col.LimpaResumoParcial()
                End With
            End If
        Next
        '--------------------------
        If RowResumo.Cells.Count = 0 Then
            bNotSetResumo = True
            Return Nothing
        Else
            mNumLinhas += 1 'incrementa linhas
            Return RowResumo
        End If
        '--------------------------
    End Function

    ''' <summary>
    ''' Checa se tem alguma coluna resumo
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ChecaTemResumo() As Boolean
        For Each col As Coluna In mColunasDetalhe.Values
            If col.TipoResumo = Coluna.TTipoResumo.RESUMO Or col.TipoResumo = Coluna.TTipoResumo.SOMA Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    '''  Formata a linha de informações do resumo final.
    ''' </summary>    
    ''' <returns>Linha HTML contendo as informações de resumo total.</returns>
    ''' <remarks></remarks>
    Private Function SetResumoFinal() As TableRow
        Dim RowResumo As New TableRow
        Dim Cell As TableCell
        '------------------------------------ 
        RowResumo.ID = "ResumoColuna"
        '------------------------------------        
        For Each col As Coluna In mColunasDetalhe.Values
            If Not mColunaGrupo.Contains(col.Nome) Then
                With RowResumo.Cells
                    If col.TipoResumo = Coluna.TTipoResumo.RESUMO Then
                        .Add(oFunc.RetCellFormat("", col.RotuloResumoFinal))
                    ElseIf col.TipoResumo = Coluna.TTipoResumo.SOMA Then
                        Cell = oFunc.RetCellFormat("", IIf(col.Formato.Trim <> "", FormatarColuna(col.ResumoTotal, col.Formato), col.ResumoTotal))
                        .Add(Cell)
                    ElseIf TemResumo Then
                        .Add(oFunc.RetCellFormat("", ""))
                    End If
                End With
            End If
        Next
        '---------------------------------
        If RowResumo.Cells.Count = 0 Then
            Return Nothing
        Else
            mNumLinhas += 1 'incrementa linhas
            Return RowResumo
        End If
        '---------------------------------
    End Function

    ''' <summary>
    ''' Formata e atribui valores às células de uma linha HTML que irá compor o detalhe do relatório.
    ''' </summary>
    ''' <param name="Row">Linha atual do <B>DataTable</B> corrente com as informações do relatório.</param>
    ''' <returns>Linha HTML com as informações já formatadas.</returns>
    ''' <remarks></remarks>
    Private Function SetDetalhe(ByVal Row As DataRow) As TableRow
        Dim RowDetalhe As New TableRow
        Dim oDado As Object
        Dim CellDetalhe As TableCell
        Dim intCellCount As Integer
        '----------------------------------------
        For Each colDet As Coluna In mColunasDetalhe.Values
            With colDet
                '----------------------------------------------
                'Se for coluna de grupo, captura próximo elemento
                If .TipoColuna = Coluna.TTipoColuna.GRUPO Then Continue For
                '----------------------------------------------
                oDado = Row(.Nome)
                If Not IsDBNull(oDado) Then
                    '------------------------------------------
                    'Se foi definido um Tamanho Máximo, 
                    'aplica truncando se necessário                           
                    If .ColumnMaxLength <> 0 Then oDado = Mid(oDado.ToString(), 1, .ColumnMaxLength)
                    '--------------------------------------------
                    'Aplica o resumo da coluna, se houver
                    If .TipoResumo = Coluna.TTipoResumo.SOMA Then colDet.ResumoSubTotal = oDado
                    '-------------------------------------------
                    'Formatando a Coluna, se necessário                    
                    If .Formato.Trim <> String.Empty Then
                        Dim oValor As Object = GetDadoTipado(oDado, colDet.TipoDado)
                        oDado = FormatarColuna(oValor, .Formato)
                    End If
                    '-------------------------------------------
                Else
                    oDado = "&nbsp"
                End If
                CellDetalhe = New TableCell
                intCellCount += 1
                CellDetalhe.ID = "cRpt" & intCellCount.ToString
                CellDetalhe.Text = IIf(oDado.ToString.Trim = "", "&nbsp;", strMantemStringExcel & " " & oDado.ToString)
                CellDetalhe.ToolTip = .ToolTip
                RowDetalhe.Cells.Add(CellDetalhe)
            End With
        Next
        '------------------------------------------
        mNumLinhas += 1     'incrementa linhas
        '------------------------------------------
        Return RowDetalhe
    End Function

    Private Function GetDadoTipado(ByVal oDado As Object, ByVal Tipo As Coluna.TTipoDado) As Object
        Select Case Tipo
            Case Coluna.TTipoDado.DATA
                If IsDate(oDado) Then
                    Return CDate(oDado)
                Else
                    Return Nothing
                End If

            Case Coluna.TTipoDado.MONEY
                If IsNumeric(oDado) Then
                    If TypeOf oDado Is String Then
                        Return Val(oDado.ToString.Replace(".", "").Replace(",", "."))
                    Else
                        Return Val(oDado)
                    End If
                Else
                    Return Val(0)
                End If

            Case Coluna.TTipoDado.NUMERICO

                If IsNumeric(oDado) Then
                    Return CInt(oDado)
                Else
                    Return Val(0)
                End If

            Case Else
                Return CStr(oDado)

        End Select

    End Function

    Private Function SetRowEmpty() As TableRow
        Dim RowEmpty As New TableRow
        Dim CellEmpty As New TableCell
        '--------------------------
        RowEmpty.ID = "rptEmpty"
        '--------------------------
        CellEmpty.Text = "&nbsp;"
        RowEmpty.Cells.Add(CellEmpty)
        '--------------------------
        mNumLinhas += 1 'incrementa linhas
        '--------------------------
        Return RowEmpty
    End Function

    Private Function GeraNovaPagina() As Table
        Dim tabPagina As Table
        tabPagina = New Table
        With tabPagina
            .ID = "rptPagina"
            .CellPadding = 0
            .CellSpacing = 0
            'With .Style
            '    .Add("width", "100%")
            '    .Add("border", "none")
            'End With
        End With
        Return tabPagina
    End Function

    ''' <summary>
    ''' Tabela de fundo para as páginas. 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Tabela gerada como parâmetro para paginação. 
    ''' Seu ID é gerado da seguinte forma 'rptPagina&lt;número da página&gt;'
    ''' </remarks>
    Private Function GeraNovaPaginaFundo() As Table
        Dim tabPaginaFundo As New Table
        '-----------------------------------
        With tabPaginaFundo
            .ID = "rptPagina" & NumPagina.ToString
            .CssClass = "rptPaginas"
            .CellPadding = 0
            .CellSpacing = 0
            '-----------------------------------            
            If mRelatorioPaginado And NumPagina > 1 And Not Exportar Then
                .Style.Add("display", "none")
            End If
            '-----------------------------------
        End With
        '-----------------------------------
        Return tabPaginaFundo
    End Function

    Private Function GeraNovoDetalhe() As Table
        Dim tabDetalhe As New Table
        Dim row As New TableRow
        '----------------------------
        With tabDetalhe
            .ID = "rptTabDetalhe"
            .CellPadding = 0
            .CellSpacing = 1
        End With
        '----------------------------
        For Each col As Coluna In mColunasDetalhe.Values
            If col.TipoColuna <> Coluna.TTipoColuna.GRUPO Then _
                row.Cells.Add(oFunc.RetCellFormat(col.Nome, col.Titulo, , , , , , "H", , col.ToolTip))
        Next
        '----------------------------
        tabDetalhe.Rows.Add(row)
        '----------------------------
        mNumLinhas += 1    'incrementa linhas
        '----------------------------
        Return tabDetalhe
    End Function

    Private Sub InicializaResumo()
        For Each Col As Coluna In ColunasDetalhe.Values
            Col.LimpaTodoResumo()
        Next
    End Sub

    ''' <summary>
    ''' Formata uma expressão de acordo com a máscara informada.
    ''' </summary>
    ''' <param name="oColValue">Valor a ser formatado.</param>
    ''' <param name="strFormat">Máscara a ser aplicada.</param>
    ''' <returns>Valor formatado.</returns>
    ''' <remarks>Caso um erro ocorra na tentativa de 
    ''' formatar o valor, o valor passado será retornado sem formatação
    ''' </remarks>
    Private Function FormatarColuna(ByVal oColValue As Object, ByVal strFormat As String)
        Try
            Return Strings.Format(oColValue, strFormat)
        Catch ex As Exception
            Return oColValue
        End Try
    End Function

    ''' <summary>
    ''' Constroe relatório com base nos parâmetros e definições 
    ''' de propriedades informadas.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ShowRelatorio() As Object
        Dim tabPagina As Table
        Dim tabPaginaFundo As Table
        Dim tabDetalhe As Table
        Dim rowResumo As TableRow
        '----------------------------------------------
        Dim oResponse As System.Web.HttpResponse = System.Web.HttpContext.Current.Response
        '----------------------------------------------
        Dim strChkParmGrupoAtual As String = String.Empty
        Dim strParmGrupoAtual As String = String.Empty
        Dim strParmGrupoAnt As String = String.Empty
        Dim bInicioGrupo As Boolean = True
        '----------------------------------------------
        'Try        
        '------------------------------------------
        'Verifica se há colunas criadas
        If mColunasDetalhe.Count = 0 Then Throw New Exception("Nenhuma coluna foi criada.")
        '------------------------------------------
        'Preenche a tabela de dados para o relatório            
        If TableDataSorce.Rows.Count = 0 Then Call PreencheTable()
        If TableDataSorce.Rows.Count = 0 Then
            Me.Controls.Add(New LiteralControl("<div id=""rptMsg"" style=""color: red;"">Nenhuma informação foi gerada.</div>"))
            Return False
        End If
        '------------------------------------------
        strMantemStringExcel = IIf(Exportar, "'", "")
        '------------------------------------------
        'Checa se tem resumo
        TemResumo = ChecaTemResumo()
        '------------------------------------------
        'Preenche colunas de Grupo
        Call SetColunasGrupo()
        '------------------------------------------
        'Gera primeira página e página de fundo
        tabPaginaFundo = GeraNovaPaginaFundo()
        tabPagina = GeraNovaPagina()
        '------------------------------------------
        'Adiciona primeiro cabeçalho na primeira página
        tabPagina.Rows.Add(SetCabecalho())
        '------------------------------------------
        If mTableDataSorce.Rows.Count > 0 Then
            '------------------------------------------
            Call InicializaResumo()
            'Verifica se relatório não é por grupo
            If mColunaGrupo.Count = 0 Then
                '------------------------------------------
                tabPagina.Rows.Add(SetRowEmpty())
                tabDetalhe = GeraNovoDetalhe()
                '------------------------------------------
                For Each dtRow As DataRow In mTableDataSorce.Rows
                    '------------------------------------------         
                    'Caso o usuário desista do processo
                    If Not oResponse.IsClientConnected Then Return False
                    '------------------------------------------
                    tabDetalhe.Rows.Add(SetDetalhe(dtRow))
                    '------------------------------------------
                    If mNumLinhas = MaxQtdeLinhas - IIf(bNotSetResumo, 1, 2) Then
                        '--------------------------       
                        'Resumo Parcial                    
                        rowResumo = SetResumo()
                        If rowResumo IsNot Nothing Then tabDetalhe.Rows.Add(rowResumo)
                        '--------------------------       
                        tabPagina.Rows.Add(oFunc.AppendToRow(tabDetalhe, , , "100%")) 'Insere a tabela detalhe
                        tabPagina.Rows.Add(oFunc.AppendToRow(SetRodape(), , , "100%"))
                        '--------------------------
                        tabPaginaFundo.Rows.Add(oFunc.AppendToRow(tabPagina, , , "100%"))
                        tabPaginaFundo.Rows.Add(oFunc.AppendToRow(QuebraPagina()))         'Quebra de Página                        
                        Me.Controls.Add(tabPaginaFundo)
                        'Me.Controls.Add(QuebraPagina())         'Quebra de Página                        
                        '--------------------------
                        tabPaginaFundo = GeraNovaPaginaFundo()  'cria nova página fundo
                        tabPagina = GeraNovaPagina()            'cria nova página
                        tabDetalhe = GeraNovoDetalhe()          'cria novo detalhe
                        '--------------------------
                        With tabPagina.Rows
                            .Add(SetCabecalho())                'Cabeçalho
                            .Add(SetRowEmpty)                   'Linha em branco (Grupo)    
                        End With
                        '--------------------------
                    End If
                Next
                '------------------------------------------ 
                'Parcial
                rowResumo = SetResumo()
                If rowResumo IsNot Nothing Then tabDetalhe.Rows.Add(rowResumo)
                'Total
                rowResumo = SetResumoFinal()
                If rowResumo IsNot Nothing Then tabDetalhe.Rows.Add(rowResumo)
                '------------------------------------------
            Else
                '------------------------------------------
                Dim dtRowNewGrupo As DataRow
                Dim lRow As Long = 0
                Dim strNewGrupoParmGrupoAnt As String = String.Empty
                '------------------------------------------
                'Relatório por grupo
                tabDetalhe = GeraNovoDetalhe()
                '------------------------------------------
                For Each dtRow As DataRow In mTableDataSorce.Rows
                    '------------------------------------------         
                    'Caso o usuário desista do processo
                    If Not oResponse.IsClientConnected Then Return False
                    '------------------------------------------
                    'Seta sempre na linha seguinte
                    lRow += 1
                    '----------------------------------------------
                    'Cria parâmetro para checar o grupo atual
                    strParmGrupoAtual = String.Empty
                    For Each cgr As String In mColunaGrupo
                        strParmGrupoAtual &= IIf(strParmGrupoAtual.Trim = "", dtRow(cgr), "|" & dtRow(cgr))
                    Next
                    '-------------------------------------------------
                    If strParmGrupoAtual <> strParmGrupoAnt Then
                        '-------------------------------------------                        
                        strParmGrupoAnt = strParmGrupoAtual
                        '-------------------------------------------
                        If Not bInicioGrupo Then
                            '-------------------------------------------
                            If QuebrarPorGrupo Then
                                '------------------------------------------
                                'Total
                                rowResumo = SetResumoFinal()
                                If rowResumo IsNot Nothing Then tabDetalhe.Rows.Add(rowResumo)
                                '------------------------------------------
                                InicializaResumo()
                                '------------------------------------------
                                tabPagina.Rows.Add(oFunc.AppendToRow(tabDetalhe, , , "100%"))
                                '-------------------------------------------
                                'Verifica se há a necessidade de
                                'complementar o relatório com linhas brancas
                                While mNumLinhas > 0 And mNumLinhas < MaxQtdeLinhas - 1
                                    tabPagina.Rows.Add(SetRowEmpty())
                                End While
                                '----------------------------------
                                tabPagina.Rows.Add(oFunc.AppendToRow(SetRodape(True), , , "100%")) 'Rodapé
                                tabPaginaFundo.Rows.Add(oFunc.AppendToRow(tabPagina))
                                tabPaginaFundo.Rows.Add(oFunc.AppendToRow(QuebraPagina()))
                                '----------------------------------
                                With Me.Controls
                                    .Add(tabPaginaFundo)
                                    '.Add(QuebraPagina()) 'Quebra de Página
                                End With
                                '----------------------------------
                                'Gera nova página 
                                tabPaginaFundo = GeraNovaPaginaFundo()
                                tabPagina = GeraNovaPagina()
                                '----------------------------------
                                'Adiciona cabeçalho
                                tabPagina.Rows.Add(SetCabecalho())
                            Else
                                '------------------------------------------
                                'Parcial
                                rowResumo = SetResumo()
                                If rowResumo IsNot Nothing Then tabDetalhe.Rows.Add(rowResumo)
                                '------------------------------------------
                                With tabPagina.Rows
                                    .Add(oFunc.AppendToRow(tabDetalhe, , , "100%"))
                                    .Add(SetRowEmpty())             'Linha Branca
                                End With
                                '------------------------------------------
                            End If
                        Else
                            bInicioGrupo = False
                        End If
                        '--------------------------------------    
                        If mNumLinhas >= (MaxQtdeLinhas - IIf(bNotSetResumo, 3, 4)) Then
                            While mNumLinhas > 0 And mNumLinhas < MaxQtdeLinhas - 1
                                tabPagina.Rows.Add(SetRowEmpty())
                            End While
                            '----------------------------------
                            tabPagina.Rows.Add(oFunc.AppendToRow(SetRodape(True), , , "100%")) 'Rodapé
                            tabPaginaFundo.Rows.Add(oFunc.AppendToRow(tabPagina))
                            tabPaginaFundo.Rows.Add(oFunc.AppendToRow(QuebraPagina()))
                            '----------------------------------
                            With Me.Controls
                                .Add(tabPaginaFundo)
                                '.Add(QuebraPagina()) 'Quebra de Página
                            End With
                            '----------------------------------
                            'Gera nova página 
                            tabPaginaFundo = GeraNovaPaginaFundo()
                            tabPagina = GeraNovaPagina()
                            '----------------------------------
                            'Adiciona cabeçalho
                            tabPagina.Rows.Add(SetCabecalho())
                        End If
                        '--------------------------------------
                        tabPagina.Rows.Add(SetGrupo(dtRow)) 'Grupo
                        tabDetalhe = GeraNovoDetalhe()
                        '--------------------------------------
                    End If
                    '------------------------------------------
                    'Adiciona linhas no detalhe
                    tabDetalhe.Rows.Add(SetDetalhe(dtRow))
                    '------------------------------------------
                    If (mNumLinhas >= (MaxQtdeLinhas - IIf(bNotSetResumo, 1, 2))) And lRow < mTableDataSorce.Rows.Count Then
                        '--------------------------
                        'Parcial                        
                        rowResumo = SetResumo()
                        If rowResumo IsNot Nothing Then tabDetalhe.Rows.Add(rowResumo)
                        '--------------------------
                        With tabPagina.Rows
                            .Add(oFunc.AppendToRow(tabDetalhe, , , "100%"))
                            .Add(oFunc.AppendToRow(SetRodape(), , , "100%"))
                            .Add(oFunc.AppendToRow(QuebraPagina()))
                        End With
                        '--------------------------
                        tabPaginaFundo.Rows.Add(oFunc.AppendToRow(tabPagina))
                        '--------------------------
                        Me.Controls.Add(tabPaginaFundo)     'Insere página no relatório                        
                        '--------------------------
                        'cria nova página e detalhe
                        tabPaginaFundo = GeraNovaPaginaFundo()
                        tabPagina = GeraNovaPagina()
                        tabDetalhe = GeraNovoDetalhe()
                        tabPagina.Rows.Add(SetCabecalho())
                        '--------------------------
                        '*****************************************************
                        'verifica se o grupo é o mesmo na proxima pagina                            
                        dtRowNewGrupo = mTableDataSorce.Rows(lRow)
                        strChkParmGrupoAtual = String.Empty
                        For Each cgr As String In mColunaGrupo
                            strChkParmGrupoAtual &= IIf(strChkParmGrupoAtual.Trim = "", dtRowNewGrupo(cgr), "|" & dtRowNewGrupo(cgr))
                        Next
                        '--------------------------
                        If strChkParmGrupoAtual <> strParmGrupoAnt Then
                            strParmGrupoAtual = strChkParmGrupoAtual
                            strParmGrupoAnt = strChkParmGrupoAtual
                            tabPagina.Rows.Add(SetGrupo(dtRowNewGrupo)) 'Grupo
                        Else
                            tabPagina.Rows.Add(SetGrupo(dtRow)) 'Grupo
                        End If
                        '*****************************************************
                    End If
                    '------------------------------------------
                Next
                '------------------------------------------             
                If QuebrarPorGrupo Then
                    '--------------------------
                    'Total                    
                    rowResumo = SetResumo()
                    If rowResumo IsNot Nothing Then tabDetalhe.Rows.Add(rowResumo)
                    '--------------------------
                Else
                    '--------------------------
                    'Parcial                    
                    rowResumo = SetResumo()
                    If rowResumo IsNot Nothing Then tabDetalhe.Rows.Add(rowResumo)
                    'Total                    
                    rowResumo = SetResumoFinal()
                    If rowResumo IsNot Nothing Then tabDetalhe.Rows.Add(rowResumo)
                    '--------------------------
                End If
                '------------------------------------------
            End If
            '------------------------------------------
            'Insere detalhe na página               
            tabPagina.Rows.Add(oFunc.AppendToRow(tabDetalhe))
            '------------------------------------------
            'Verifica se há a necessidade de 
            'complementar o relatório com linhas brancas
            While mNumLinhas > 0 And mNumLinhas <= MaxQtdeLinhas - 1
                tabPagina.Rows.Add(SetRowEmpty())
            End While
            '------------------------------------
            tabPagina.Rows.Add(oFunc.AppendToRow(SetRodape(True), , , "100%")) 'Rodapé
            '------------------------------------
            tabPaginaFundo.Rows.Add(oFunc.AppendToRow(tabPagina))
            '------------------------------------------            
            With Me.Controls
                .Add(tabPaginaFundo)
                'Se for relatório paginado, insere
                'HiddenField para controle da última página
                'If RelatorioPaginado Then
                With Page.ClientScript
                    .RegisterHiddenField("hdnMaxPage", NumPagina)
                    .RegisterHiddenField("hdnPage", "1")
                End With
                'End If
            End With
            '------------------------------------------
            If OcultarParamAoCarregar Then
                With New StringBuilder
                    .AppendLine("<script language='javascript'>")
                    .AppendLine("try {")
                    .AppendLine("   lProcHeight = 0;")
                    .AppendLine("   hs = false;")
                    .AppendLine("   y = document.getElementById('imgHideShow');")
                    .AppendLine("   x = document.getElementById('divParametros');")
                    .AppendLine("   w = document.getElementById(""rptParamTitulo"");")
                    .AppendLine("   lHeight = document.getElementById('rptParametros').offsetHeight;")
                    .AppendLine("   ShowCaixa();")
                    .AppendLine("}catch(e){}")
                    .AppendLine("</script>")
                    HttpContext.Current.Response.Write(.ToString)
                End With
            End If
            '------------------------------------------
        End If
        'Catch ex As Exception

        'End Try
        Return True
    End Function

    Private Function RetSQL() As String
        Dim sSQL As String = String.Empty
        Dim sSQLNEW As String = String.Empty
        RetSQL = String.Empty
        Select Case TipoConsulta
            Case TTipoConsulta.SQL
                If ClausulaWhere.Trim <> String.Empty Then
                    sSQLNEW = SQLQuery
                    If sSQLNEW.ToUpper.Contains(" ORDER BY ") And Not sSQLNEW.ToUpper.Contains(" TOP ") Then sSQLNEW = sSQLNEW.ToUpper.Replace("SELECT", "SELECT TOP 100 PERCENT ")
                    sSQL = "SELECT * FROM (" & sSQLNEW & ") a WHERE " & ClausulaWhere
                    Return sSQL
                Else
                    Return SQLQuery
                End If

            Case TTipoConsulta.STORED_PROCEDURE
                Return SQLQuery & ClausulaWhere

        End Select
    End Function

    ''' <summary>
    ''' Preenche o DataTable de dados, quando nenhum DataTable 
    ''' foi passado anteriormente. 
    ''' </summary>
    ''' <remarks>Um comando Sql e uma String de Conexão
    ''' são obrigatórios para a utilização desse método.</remarks>
    Private Sub PreencheTable()
        Try
            'dta = New OleDb.OleDbDataAdapter(RetSQL(), New System.Data.OleDb.OleDbConnection(ConnectionString))
            oDB = New ClsDB(ConnectionString, Provider)
            dta = oDB.GetDataAdapter(RetSQL())
            dta.Fill(mTableDataSorce)
        Catch ex As Exception
            Me.Controls.Add(New LiteralControl("<div id=""rptMsg"" style=""color: red;"">Erro: <br/>" & ex.Message & "</div>"))
        Finally
            Try : dta.Dispose() : Catch : End Try
        End Try
    End Sub

    Private Sub GerarExportacao()
        Dim excel As New ClsGetExcelFile
        Dim Response As HttpResponse = HttpContext.Current.Response
        '--------------------------------
        If TableDataSorce.Rows.Count = 0 Then Call PreencheTable()
        If TableDataSorce.Rows.Count = 0 Then
            Page.Controls.Clear()
            Response.Clear()
            Me.Controls.Add(New LiteralControl("<div id=""rptMsg"" style=""color: red;"">Nenhuma informação foi gerada.</div>"))
        Else
            'Captura as colunas definidas como grupo
            Call SetColunasGrupo()
            With excel
                .Titulo = TituloRelatorio
                If QuebrarPaginaExportacao Then .LinhasPorPagina = MaxQtdeLinhas
                .DataSource = mTableDataSorce
                '-----------------------------------
                'Colunas grupo
                If ExportarAgrupado Then
                    For Each s As String In mColunaGrupo
                        .AddGroupColumn(mColunasDetalhe(s).Nome)
                    Next
                End If
                '-----------------------------------
                'Colunas Detalhe
                For Each s As String In mColunasDetalhe.Keys
                    .AddColumnTitle(mColunasDetalhe(s).Nome, mColunasDetalhe(s).Titulo)
                Next
                '-----------------------------------
                'Gera a planilha
                .GenerateXLS()
                '-----------------------------------                
            End With
            '-----------------------------------                
            'Enviando informações
            Dim aBytes() As Byte = CType(excel.GetStream, MemoryStream).ToArray
            With Response
                .Clear()
                .AddHeader("Content-Disposition", "attachment; filename=" & excel.FileName)
                .AddHeader("Content-Length", aBytes.Length.ToString())
                .ContentType = "application/vnd.ms-excel"
                .BinaryWrite(aBytes)
                .End()
            End With
            '-----------------------------------                
        End If
        '--------------------------------
    End Sub


#End Region

End Class


Public Class Colunas
    Inherits Dictionary(Of String, Coluna)

    Private mColuna As Coluna

    Public Overloads Function Add(ByVal ColunaNome As String) As Coluna
        mColuna = New Coluna()
        '-------------------------------
        With mColuna
            .Nome = ColunaNome
            .Titulo = ColunaNome
            .Alinhamento = Coluna.TAlinhamento.LEFT
        End With
        '-------------------------------
        MyBase.Add(ColunaNome, mColuna)
        '-------------------------------
        Return MyBase.Item(ColunaNome)
    End Function

    Public Overloads Function Add(ByVal pColuna As Coluna) As Coluna
        '-------------------------------
        MyBase.Add(pColuna.Nome, pColuna)
        '-------------------------------
        Return MyBase.Item(pColuna.Nome)
    End Function

End Class

Public Class Coluna

#Region "Fields"

    Private mFormato As String = String.Empty
    Private mToolTip As String = String.Empty
    Private mNome As String = String.Empty
    Private mTitulo As String = String.Empty
    Private mRotuloResumoPag As String = String.Empty
    Private mRotuloResumoFinal As String = String.Empty
    '----------------------------------------
    Private mColumnSize As Unit
    '----------------------------------------
    Private mTipoColuna As Short = 0
    Private mColumnMaxLength As Short = 0
    Private mNumLinhasGrupo As Short = 0   'FUTURE..........
    Private mTipoResumo As Short = 0
    '----------------------------------------
    Private mAlinhamento As TAlinhamento
    '----------------------------------------
    Private mResumoSubTotal As Double = 0
    Private mResumoTotal As Double = 0
    Private mQuebrarTexto As TQuebrarTexto = TQuebrarTexto.NAO
    Private mTipoDado As TTipoDado
    '----------------------------------------
    Private HASH_TTipoColuna As Hashtable
    Private HASH_TAlinhamento As Hashtable
    Private HASH_TTipoResumo As Hashtable
    Private HASH_TTipoDado As Hashtable
    '----------------------------------------
    Private colHeaderStyle As Hashtable

#End Region

#Region "Constructors/Destructors"

    Public Sub New()
        '--------------------------
        HASH_TTipoColuna = New Hashtable
        With HASH_TTipoColuna
            .Add(CShort(TTipoColuna.DETALHE), TTipoColuna.DETALHE)
            .Add(CShort(TTipoColuna.GRUPO), TTipoColuna.GRUPO)
        End With
        '--------------------------
        HASH_TAlinhamento = New Hashtable
        With HASH_TAlinhamento
            .Add(CShort(Coluna.TAlinhamento.CENTER), "center")
            .Add(CShort(Coluna.TAlinhamento.LEFT), "left")
            .Add(CShort(Coluna.TAlinhamento.RIGHT), "right")
        End With
        '--------------------------
        HASH_TTipoResumo = New Hashtable
        With HASH_TTipoResumo
            .Add(CShort(TTipoResumo.SEM_RESUMO), TTipoResumo.SEM_RESUMO)
            .Add(CShort(TTipoResumo.SOMA), TTipoResumo.SOMA)
            .Add(CShort(TTipoResumo.RESUMO), TTipoResumo.RESUMO)
        End With
        '--------------------------
        HASH_TTipoDado = New Hashtable
        With HASH_TTipoResumo
            .Add(CShort(TTipoDado.CARACTER), TTipoDado.CARACTER)
            .Add(CShort(TTipoDado.DATA), TTipoDado.DATA)
            .Add(CShort(TTipoDado.MONEY), TTipoDado.MONEY)
            .Add(CShort(TTipoDado.NUMERICO), TTipoDado.NUMERICO)
        End With
    End Sub

#End Region

#Region "Enumerations"

    Public Enum TTipoResumo As Short
        SEM_RESUMO = 8
        SOMA = 9
        RESUMO = 10
    End Enum

    Public Enum TAlinhamento As Short
        LEFT = 1
        CENTER = 2
        RIGHT = 7
    End Enum

    Public Enum TTipoColuna As Short
        GRUPO = 13
        DETALHE = 14
    End Enum

    Public Enum TQuebrarTexto As Short
        SIM = 35
        NAO = 36
    End Enum

    Public Enum TTipoDado As Short
        CARACTER = 27
        NUMERICO = 28
        DATA = 29
        MONEY = 30
    End Enum

#End Region

#Region "Properties"

    ''' <summary>
    ''' Tipo de dado da coluna
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TipoDado() As TTipoDado
        Get
            Return mTipoDado
        End Get
        Set(ByVal value As TTipoDado)
            mTipoDado = value
        End Set
    End Property

    ''' <summary>
    ''' Hashtable configurações de stylesheet para a coluna do relatório (Property/Value).
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property HeaderStyle() As Hashtable
        Get
            If colHeaderStyle Is Nothing Then colHeaderStyle = New Hashtable
            Return colHeaderStyle
        End Get
        Set(ByVal value As Hashtable)
            colHeaderStyle = value
        End Set
    End Property

    ''' <summary>
    ''' Largura da coluna.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnSize() As Unit
        Get
            Return mColumnSize
        End Get
        Set(ByVal value As Unit)
            mColumnSize = value
        End Set
    End Property

    ''' <summary>
    ''' Quantidade de caracteres máxima para a coluna.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnMaxLength() As Integer
        Get
            Return mColumnMaxLength
        End Get
        Set(ByVal value As Integer)
            mColumnMaxLength = value
        End Set
    End Property

    ''' <summary>
    ''' Formato a ser aplicado na coluna
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Formato() As String
        Get
            Return mFormato
        End Get
        Set(ByVal value As String)
            mFormato = value
        End Set
    End Property

    ''' <summary>
    ''' Texto de dica para a coluna do relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ToolTip() As String
        Get
            Return mToolTip
        End Get
        Set(ByVal value As String)
            mToolTip = value
        End Set
    End Property

    ''' <summary>
    ''' Nome do Relatório    
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Nome() As String
        Get
            Return mNome
        End Get
        Set(ByVal value As String)
            mNome = value
        End Set
    End Property

    ''' <summary>
    ''' Rótulo do resumo final
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RotuloResumoFinal() As String
        Get
            Return mRotuloResumoFinal
        End Get
        Set(ByVal value As String)
            mRotuloResumoFinal = value
        End Set
    End Property

    ''' <summary>
    ''' Rótulo do Resumo da página.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RotuloResumoPag() As String
        Get
            Return mRotuloResumoPag
        End Get
        Set(ByVal value As String)
            mRotuloResumoPag = value
        End Set
    End Property

    ''' <summary>
    ''' Títulod do relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Titulo() As String
        Get
            Return mTitulo.Replace(" ", "&nbsp;")
        End Get
        Set(ByVal value As String)
            mTitulo = value
        End Set
    End Property

    ''' <summary>
    ''' Tipo da coluna. Se de Grupo ou de Detalhe.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TipoColuna() As TTipoColuna
        Get
            If mTipoColuna = 0 Then mTipoColuna = TTipoColuna.DETALHE
            Return mTipoColuna
        End Get
        Set(ByVal value As TTipoColuna)
            mTipoColuna = TrataProp(value, HASH_TTipoColuna)
        End Set
    End Property

    ''' <summary>
    ''' Alinhamento da Coluna.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Alinhamento() As TAlinhamento
        Get
            If mAlinhamento = 0 Then mAlinhamento = TAlinhamento.LEFT
            Return mAlinhamento
        End Get
        Set(ByVal value As TAlinhamento)
            mAlinhamento = TrataProp(value, HASH_TAlinhamento)
        End Set
    End Property

    ''' <summary>
    ''' Se a coluna será Rótulo do Resumo, Conterá o Resumo ao 
    ''' final do detalhe ou sem valor para resumo
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TipoResumo() As TTipoResumo
        Get
            If mTipoResumo = 0 Then mTipoResumo = TTipoResumo.SEM_RESUMO
            Return mTipoResumo
        End Get
        Set(ByVal value As TTipoResumo)
            mTipoResumo = TrataProp(value, HASH_TTipoResumo)
        End Set
    End Property

    ''' <summary>
    ''' Se a coluna irá quebrar o texto (Wrap).
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property QuebrarTexto() As TQuebrarTexto
        Get
            Return mQuebrarTexto
        End Get
        Set(ByVal value As TQuebrarTexto)
            mQuebrarTexto = value
        End Set
    End Property

    ''' <summary>
    ''' Resumo Subtotal
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Property ResumoSubTotal() As Double
        Get
            Return mResumoSubTotal
        End Get
        Set(ByVal value As Double)
            mResumoSubTotal += value
            mResumoTotal += value
        End Set
    End Property

    ''' <summary>
    ''' Resumo Total do relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Property ResumoTotal() As Double
        Get
            Return mResumoTotal
        End Get
        Set(ByVal value As Double)
            mResumoTotal = value
        End Set
    End Property

    Private Function TrataProp(ByVal ValProp As Short, ByRef hsh As Hashtable) As Integer
        If hsh.ContainsKey(ValProp) Then
            Return ValProp
        Else
            Throw New Exception("Valor informado é inválido.")
        End If
    End Function

    ''' <summary>
    ''' Limpa todos os valores do resumo, tanto parcial quanto total.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub LimpaTodoResumo()
        mResumoSubTotal = 0
        mResumoTotal = 0
    End Sub

    ''' <summary>
    ''' Limpa somente o resumo parcial.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub LimpaResumoParcial()
        mResumoSubTotal = 0
    End Sub

#End Region

End Class

#End Region


'**********************************************************************************
'**********************************************************************************


#Region "RPTPARAM"

<System.Serializable()> _
<ToolboxData("<{0}:ClsRPTParam runat=server></{0}:ClsRPTParam>")> _
Public Class ClsRPTParam
    Inherits WebControl
    Implements IPostBackEventHandler

#Region "Constructors"

    Public Sub New()
        With HASH_TEstilo
            .Add(CShort(TEstilo.LATERAL), TEstilo.LATERAL)
            .Add(CShort(TEstilo.TOPO), TEstilo.TOPO)
        End With

        With HASH_TEstiloCaption
            .Add(CShort(TEstiloCaption.LEFT), TEstiloCaption.LEFT)
            .Add(CShort(TEstiloCaption.TOP), TEstiloCaption.TOP)
        End With

        With HASH_Operadores
            .Add(CShort(ColunaParam.TOperador.UNDEFINED), "")
            .Add(CShort(ColunaParam.TOperador.IGUAL), "=")
            .Add(CShort(ColunaParam.TOperador.MAIOR), ">")
            .Add(CShort(ColunaParam.TOperador.MENOR), "<")
            .Add(CShort(ColunaParam.TOperador.MAIOR_IGUAL), ">=")
            .Add(CShort(ColunaParam.TOperador.MENOR_IGUAL), "<=")
            .Add(CShort(ColunaParam.TOperador.A_PARTIR_DE), "Like '#%'")
            .Add(CShort(ColunaParam.TOperador.TERMINADO_EM), "Like '%#'")
            .Add(CShort(ColunaParam.TOperador.CONTENDO), "Like '%#%'")
            .Add(CShort(ColunaParam.TOperador.DENTRO_DE), "IN(#)")
        End With

        With HASH_OperadoresNum
            .Add(CShort(ColunaParam.TOperador.UNDEFINED), "")
            .Add(CShort(ColunaParam.TOperador.IGUAL), "=")
            .Add(CShort(ColunaParam.TOperador.MAIOR), ">")
            .Add(CShort(ColunaParam.TOperador.MENOR), "<")
            .Add(CShort(ColunaParam.TOperador.MAIOR_IGUAL), ">=")
            .Add(CShort(ColunaParam.TOperador.MENOR_IGUAL), "<=")
            .Add(CShort(ColunaParam.TOperador.DENTRO_DE), "IN(#)")
        End With

        With HASH_TradOperadores
            .Add(CShort(ColunaParam.TOperador.IGUAL), "IGUAL")
            .Add(CShort(ColunaParam.TOperador.MAIOR), "MAIOR")
            .Add(CShort(ColunaParam.TOperador.MENOR), "MENOR")
            .Add(CShort(ColunaParam.TOperador.MAIOR_IGUAL), "MAIOR OU IGUAL")
            .Add(CShort(ColunaParam.TOperador.MENOR_IGUAL), "MENOR OU IGUAL")
            .Add(CShort(ColunaParam.TOperador.A_PARTIR_DE), "A PARTIR DE")
            .Add(CShort(ColunaParam.TOperador.TERMINADO_EM), "TERMINANDO COM")
            .Add(CShort(ColunaParam.TOperador.CONTENDO), "CONTENDO")
            .Add(CShort(ColunaParam.TOperador.DENTRO_DE), "DENTRO DE")
        End With

        With HASH_TipoConsulta
            .Add(CShort(ClsRPTView.TTipoConsulta.SQL), ClsRPTView.TTipoConsulta.SQL)
            .Add(CShort(ClsRPTView.TTipoConsulta.STORED_PROCEDURE), ClsRPTView.TTipoConsulta.STORED_PROCEDURE)
        End With

    End Sub

#End Region


#Region "Fields"

    Private dta As DbDataAdapter
    Private oDB As ClsDB
    Private mTableDataSorce As DataTable
    '----------------------------------------
    Private oFunc As New ClsFunctions
    '----------------------------------------  
    Private mColunasParam As New ColunasParam
    Private mEstilo As TEstilo = TEstilo.TOPO    
    '--------------------------------------------
    Private HASH_TEstilo As New Hashtable
    Private HASH_TEstiloCaption As New Hashtable
    Private HASH_Operadores As New Hashtable
    Private HASH_OperadoresNum As New Hashtable
    Private HASH_OperadoresText As New Hashtable
    Private HASH_TradOperadores As New Hashtable
    Private HASH_TipoConsulta As New Hashtable
    '--------------------------------------------    

#End Region


#Region "Enumerations"

    Public Enum TEstilo As Short
        TOPO
        LATERAL
    End Enum

    Public Enum TOption
        OPT_EXPORTAR
        OPT_VISUALIZAR
    End Enum

    Public Enum TEstiloCaption
        TOP = 33
        LEFT = 34
    End Enum

#End Region


#Region "Events"

    Public Event ParamClick(ByVal pOption As TOption)

    Protected Overrides Sub RenderContents(ByVal output As HtmlTextWriter)
        Me.Controls.Clear()
        Me.Style.Clear()
        Me.ControlStyle.Width = Unit.Percentage(100)
        Me.Style.Item("width") = "100%"
        Me.Width = Unit.Percentage(100)
        If Me.DesignMode Then
            Me.Controls.Add(New LiteralControl("<div style=""width:100%"">" & Me.ClientID & "</div>"))
        Else
            Call ShowParametros()
        End If
        MyBase.RenderContents(output)
    End Sub

    Private Sub RaisePostBackEvent(ByVal eventArgument As String) Implements System.Web.UI.IPostBackEventHandler.RaisePostBackEvent
        Dim aParam As Array
        Call SalvaValores()

        aParam = eventArgument.Split("$")
        If aParam(0) = 0 Then
            RaiseEvent ParamClick(aParam(1))
        Else
            Dim Response As HttpResponse = HttpContext.Current.Response
            Response.Redirect(URL)
        End If
    End Sub

    Private Sub SalvaValores()
        '------------------------------------------------------
        If mColunasParam.Count = 0 Then Exit Sub 'mColunasParam = ColunasParam
        '------------------------------------------------------
        'Captura o contexto da página
        Dim Request As System.Web.HttpRequest = System.Web.HttpContext.Current.Request
        'Dim strValor As String
        '------------------------------------------------------
        'Salva os operadores 
        Dim IndiceColuna As Integer = 0
        Dim strNomeCpo As String
        For Each col As ColunaParam In mColunasParam.Values
            strNomeCpo = col.Nome & "_" & IndiceColuna
            If Request(strNomeCpo) IsNot Nothing Then
                'If col.TipoColuna = ColunaParam.TTipoColuna.CARACTER And col.TipoCampo = ColunaParam.TTipoCampo.LISTA_MULTISELECT Then
                '    strValor = Request(strNomeCpo).ToString.Trim
                '    If strValor <> String.Empty Then
                '        col.Value = "'" & strValor.Replace(",", "','") & "'"
                '    Else
                '        col.Value = String.Empty
                '    End If
                'Else
                col.Value = Request(strNomeCpo)
                'End If
            End If
            '--------------------------------
            If Request(strNomeCpo & "_Operador") <> "" Then col.OperadorDefinido = Request(strNomeCpo & "_Operador")
            '--------------------------------
            IndiceColuna += 1
            '--------------------------------
        Next
        '------------------------------------------------------
    End Sub

    Protected Overrides Function SaveControlState() As Object
        Dim c As New Dictionary(Of String, ColunaParam)

        For Each col As ColunaParam In mColunasParam.Values
            c.Add(c.Count, col)
        Next

        Return c
    End Function

    Protected Overrides Sub LoadControlState(ByVal savedState As Object)
        Dim c As New Dictionary(Of String, ColunaParam)
        c = savedState

        For Each col As ColunaParam In c.Values
            mColunasParam.Add(mColunasParam.Count.ToString, col)
        Next

        Call SalvaValores()
    End Sub

    Protected Overrides Sub OnInit(ByVal e As System.EventArgs)
        Page.RegisterRequiresControlState(Me)
        MyBase.OnInit(e)
        With Page.ClientScript            
            '.RegisterClientScriptInclude(Page.GetType, "functions_valida", "http://10.0.0.238/corpnet2/forms.js")
            '.RegisterClientScriptInclude(Page.GetType, "functions", "http://10.0.0.238/corpnet2/functionsRPT.js")
            '.RegisterClientScriptInclude(Page.GetType, "functions_valida", "forms.js")
            '.RegisterClientScriptInclude(Page.GetType, "functions", "functionsRPT.js")

            .RegisterClientScriptInclude(Page.GetType, "functions_valida", "http://localhost/corpnet2/forms.js")
            .RegisterClientScriptInclude(Page.GetType, "functions", "http://localhost/corpnet2/functionsRPT.js")

            .RegisterStartupScript(Me.GetType, "initialize", "manageform.Initialize(false);manageform.evtValidacaoOk = function(){ rptView.Esperando();}; rptView.Initialize();", True)
        End With
    End Sub

#End Region


#Region "Properties"

    ''' <summary>
    ''' Atribui/Recebe valor lógico para exibir ou não o botão "Visualizar Relatório".
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ShowVisualizarRel() As Boolean
        Get
            If ViewState("ShowVisualizarRel") Is Nothing Then ViewState("ShowVisualizarRel") = True
            Return ViewState("ShowVisualizarRel")
        End Get
        Set(ByVal value As Boolean)
            ViewState("ShowVisualizarRel") = value
        End Set
    End Property

    ''' <summary>
    ''' Atribui/Recebe valor lógico para exibir ou não o botão "Imprimir Relatório".
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ShowImprimirRelatorio() As Boolean
        Get
            If ViewState("ShowImprimirRelatorio") Is Nothing Then ViewState("ShowImprimirRelatorio") = True
            Return ViewState("ShowImprimirRelatorio")
        End Get
        Set(ByVal value As Boolean)
            ViewState("ShowImprimirRelatorio") = value
        End Set
    End Property

    ''' <summary>
    ''' Atribui/Recebe valor lógico para exibir ou não o botão "Exportar Relatório".
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ShowExportarRelatorio() As Boolean
        Get
            If ViewState("ShowExportarRelatorio") Is Nothing Then ViewState("ShowExportarRelatorio") = True
            Return ViewState("ShowExportarRelatorio")
        End Get
        Set(ByVal value As Boolean)
            ViewState("ShowExportarRelatorio") = value
        End Set
    End Property

    ''' <summary>
    ''' Atribui/Recebe valor lógico para exibir ou não o botão "Voltar Página".
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ShowVoltarPagina() As Boolean
        Get
            If ViewState("ShowVoltarPagina") Is Nothing Then ViewState("ShowVoltarPagina") = True
            Return ViewState("ShowVoltarPagina")
        End Get
        Set(ByVal value As Boolean)
            ViewState("ShowVoltarPagina") = value
        End Set
    End Property

    ''' <summary>
    ''' Permite gerenciar as colunas que serão utilizadas como
    ''' parâmetros para o relatório.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColunasParam() As ColunasParam
        Get
            Return mColunasParam
        End Get
        Set(ByVal value As ColunasParam)
            mColunasParam = value
        End Set
    End Property

    Public Property TipoConsulta() As ClsRPTView.TTipoConsulta
        Get
            Dim o As Object = ViewState("TipoConsulta")
            If o Is Nothing Then ViewState("TipoConsulta") = ClsRPTView.TTipoConsulta.SQL
            Return ViewState("TipoConsulta")
        End Get
        Set(ByVal value As ClsRPTView.TTipoConsulta)
            ViewState("TipoConsulta") = TrataProp(value, HASH_TipoConsulta)
        End Set
    End Property

    ''' <summary>
    ''' Define um destino ao clicar do botão voltar. Caso
    ''' não seja informado nenhum valor, será definido como o 
    ''' voltar do histórico de navegação.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property URL() As String
        Get
            Dim o As Object = ViewState("URL")
            If o Is Nothing Then ViewState("URL") = String.Empty
            Return ViewState("URL")
        End Get
        Set(ByVal value As String)
            ViewState("URL") = value
        End Set
    End Property

    ''' <summary>
    ''' Texto de dica para o botão voltar.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ToolTipBackButton() As String
        Get
            Dim o As Object = ViewState("ToolTipButtonBack")
            If o Is Nothing Then ViewState("ToolTipButtonBack") = "Voltar para página anterior"
            Return ViewState("ToolTipButtonBack")
        End Get
        Set(ByVal value As String)
            ViewState("ToolTipButtonBack") = value
        End Set
    End Property

    ''' <summary>
    ''' String de Conexão para a base de dados desejada.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Obsolete("Utilize a String de Conexão do campo parâmetro ao invés deste que é geral.")> _
    Public Property ConnectionString() As String
        Get
            Return TrataViewState("ConnectionString")
        End Get
        Set(ByVal value As String)
            ViewState("ConnectionString") = value
        End Set
    End Property

    ''' <summary>
    ''' Título a ser exibido na caixa de parâmetros.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TituloRel() As String
        Get
            Return TrataViewState("TituloRel")
        End Get
        Set(ByVal value As String)
            ViewState("TituloRel") = value
        End Set
    End Property

    ''' <summary>
    ''' Indica a característica da caixa de parâmetros (TOPO ou CAIXA LATERAL).
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EstiloParametros() As TEstilo
        Get
            Dim o As Object = ViewState("Estilo")
            If o IsNot Nothing Then
                Return o
            Else
                Return TEstilo.TOPO
            End If
        End Get
        Set(ByVal value As TEstilo)
            ViewState("Estilo") = TrataProp(value, HASH_TEstilo)
        End Set
    End Property

    Public Property RelatorioPaginado() As Boolean
        Get
            Dim o As Object = ViewState("RelatorioPaginado")
            If o IsNot Nothing Then
                Return o
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            ViewState("RelatorioPaginado") = value
        End Set
    End Property

    Public Overrides Property Width() As System.Web.UI.WebControls.Unit
        Get
            MyBase.Width = Unit.Percentage(100)
            Return MyBase.Width
        End Get
        Set(ByVal value As System.Web.UI.WebControls.Unit)
            MyBase.Width = Unit.Percentage(100)  'value
        End Set
    End Property

    ''' <summary>
    ''' Configura a posição dos rótulos dos parâmetros como TOP ou LEFT.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EstiloCaption() As TEstiloCaption
        Get
            Dim o As Object = ViewState("EstiloCaption")
            If o IsNot Nothing Then
                Return o
            Else
                Return TEstiloCaption.LEFT
            End If
        End Get
        Set(ByVal value As TEstiloCaption)
            ViewState("EstiloCaption") = TrataProp(value, HASH_TEstiloCaption)
        End Set
    End Property

    Private Function TrataProp(ByVal ValProp As Short, ByRef hsh As Hashtable) As Integer
        If hsh.ContainsKey(ValProp) Then
            Return ValProp
        Else
            Throw New Exception("Valor informado é inválido.")
        End If
    End Function

    Private Function TrataViewState(ByVal pKey As String) As String
        '----------------------------------
        Dim o As Object = ViewState(pKey)
        '----------------------------------
        If o IsNot Nothing Then
            Return o
        Else
            Return String.Empty
        End If
        '----------------------------------
    End Function

    ''' <summary>
    ''' Imagem a ser exibida para validação dos campos.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ImgAlerta() As String
        Get
            If ViewState("ImgAlerta") Is Nothing Then ViewState("ImgAlerta") = "http://10.0.0.238/corpnet2/imagens/alerta.jpg"
            Return ViewState("ImgAlerta")
        End Get
        Set(ByVal value As String)
            ViewState("ImgAlerta") = value
        End Set
    End Property

#End Region


#Region "Methods"

    Public Function GetSQL(ByVal sSQL As String) As String
        Dim sWhere As String = GetWhere()
        If TipoConsulta = ClsRPTView.TTipoConsulta.SQL Then
            If sWhere.Trim <> String.Empty Then
                sSQL = sSQL & IIf(sSQL.ToUpper.Contains("WHERE"), " AND ", " WHERE ") & sWhere
                Return sSQL
            Else
                Return sSQL
            End If
        Else
            Return sSQL & " " & sWhere
        End If
    End Function

    Private Function SetCellImg(ByVal strId As String, _
                                ByVal strPath As String, _
                                ByVal pSize As String, _
                                Optional ByVal pAtributo As String = "", _
                                Optional ByVal Alinhamento As HorizontalAlign = HorizontalAlign.Center, _
                                Optional ByVal ToolTip As String = "", _
                                Optional ByVal VAlinhamento As VerticalAlign = VerticalAlign.NotSet) As TableCell

        'Dim imgToolbar As New ImageButton
        ''--------------------------
        'With imgToolbar
        '    .ID = strId
        '    .ImageUrl = strPath
        '    .BorderStyle = WebControls.BorderStyle.None
        '    .ToolTip = ToolTip
        '    If pAtributo.Trim <> String.Empty Then
        '        .OnClientClick = "javascript:return " & pAtributo
        '    End If
        'End With
        '--------------------------

        'Dim imgToolbar As New LinkButton
        Dim imgToolbar As New Image
        '--------------------------
        With imgToolbar
            .ID = strId
            .ImageUrl = strPath            
            With .Style
                .Add("cursor", "hand")
                .Add("cursor", "pointer")
            End With
            .BorderStyle = WebControls.BorderStyle.None
            .ToolTip = ToolTip
            If pAtributo.Trim <> String.Empty Then
                .Attributes.Add("onclick", "javascript:return " & pAtributo)
            End If
        End With
        '--------------------------
        Return oFunc.AppendToCell(imgToolbar, , Alinhamento, pSize, "0", , , "N", VAlinhamento)
        '--------------------------
    End Function

    Private Function GetImgAlerta(ByVal NomeCampo As String) As Image
        Dim imgAlert As New Image
        With imgAlert
            .ID = NomeCampo
            .ImageUrl = Me.ImgAlerta
            .Style.Add("visibility", "hidden")
        End With
        Return imgAlert
    End Function

    Public Sub ShowParametros()
        Dim tabGeral As New Table
        Dim CellGeral As TableCell
        Dim RowGeral As New TableRow
        Dim tabParametros As New Table
        Dim tabTitulo As New Table
        Dim Row As TableRow
        Dim oField As New Object
        Dim imgAlert As Image
        '--------------------------
        With tabGeral
            .ID = "tabParamGeral"
            .CellPadding = 0
            .CellSpacing = 0
            With .Style
                .Add("width", "100%")
                .Add("height", "0px")
                .Add("border", "0px")
            End With
        End With    
        '--------------------------
        ' Se o estilo de exibição é para 
        ' Relatórios(externos) com os parâmetros
        ' apresentados no topo então
        With tabTitulo
            'If mEstilo = TEstilo.TOPO Then
            '--------------------------
            .ID = "rptParamTitulo"
            .Style.Add("width", "100%")
            '--------------------------
            Row = New TableRow
            With Row.Cells
                '----------------------------------------
                ' Se o titulo da caixa de parâmetros 
                ' não foi informado, atribui um por default
                If TituloRel.Trim = String.Empty Then TituloRel = "&nbsp;"
                '----------------------------------------
                .Add(oFunc.RetCellFormat("", TituloRel, HorizontalAlign.Center, "95%", TypeCell:="H"))
                '.Add(oFunc.RetCellFormat("", TituloRel, HorizontalAlign.Center, "", , , , "H"))
                '----------------------------------------
                ' Visualizar                
                If ShowVisualizarRel Then
                    Page.ClientScript.RegisterStartupScript(Page.GetType, "btn_loadexport", "function LoadVisualizar(){ if (rptView.ChecaOperador()){ rptView.Exportando = false; " & Page.ClientScript.GetPostBackEventReference(New PostBackOptions(Me, "0$" & TOption.OPT_VISUALIZAR)) & " }; }", True)
                    .Add(SetCellImg("imgShowRel", "~\imagens\visualizar.gif", "10px", "LoadVisualizar()", , "Visualizar relatório"))
                    CType(Row.Cells(1).Controls(Row.Cells(1).Controls.Count - 1), Image).Attributes.Add("newcorp", "true")
                End If
                '----------------------------------------
                ' Imprimir
                If ShowImprimirRelatorio Then
                    .Add(SetCellImg("imgPrintRel", "~\imagens\print.gif", "10px", "rptView.imprimir(" & RelatorioPaginado.ToString.ToLower & ");", , "Imprimir relatório"))
                End If
                '----------------------------------------
                ' Exportarl
                If ShowExportarRelatorio Then
                    .Add(SetCellImg("imgExportar", "~\imagens\excel.gif", "10px", "LoadExport()", , "Exportar para Excel"))
                    Page.ClientScript.RegisterStartupScript(Page.GetType, "btn_export", "function LoadExport(){ if(rptView.ChecaOperador()){ rptView.Exportando = true; " & Page.ClientScript.GetPostBackEventReference(Me, "0$" & TOption.OPT_EXPORTAR, True) & " }; }", True)
                    With CType(Row.Cells(Row.Cells.Count - 1).Controls(Row.Cells(Row.Cells.Count - 1).Controls.Count - 1), Image).Attributes
                        .Add("newcorp", "true")
                    End With
                End If
                '----------------------------------
                ' Verifica se o relatório é paginado 
                ' e exibe os botões para paginação
                If RelatorioPaginado Then
                    '----------------------------------
                    Dim txtNumPage As New TextBox
                    Dim lblMaxPage As New Label
                    Dim lblTextoMaxPage As New Label
                    '----------------------------------
                    .Add(oFunc.RetCellFormat("rptDivisor1", "|", , "30px"))
                    With txtNumPage
                        .ID = "txtNumPag"
                        .ToolTip = "Número da página"
                        .Text = "0"
                        .Style.Add("text-align", "center")
                        .Attributes.Add("onkeypress", "if (event.keyCode == 13) return rptView.gotoPage(this);")
                        .Width = Unit.Pixel(20)
                    End With
                    '----------------------------------
                    lblTextoMaxPage.Text = "&nbsp;de&nbsp;"
                    '----------------------------------
                    With lblMaxPage
                        .ID = "lblMaxPage"
                        .ToolTip = "Número total de páginas"
                        .Style.Add("text-align", "center")
                        .Width = Unit.Pixel(20)
                        .Text = "0"
                    End With
                    '----------------------------------
                    .Add(SetCellImg("imgFirst", "~\imagens\first.ico", "10px", "rptView.movepage(-1);", , "Primeira página"))    'Primeira página
                    .Add(SetCellImg("imgPrior", "~\imagens\previous.ico", "10px", "rptView.movepage(0);", , "Página anterior"))  'Anterior 
                    .Add(oFunc.AppendToCell(txtNumPage))
                    .Add(oFunc.AppendToCell(lblTextoMaxPage))
                    .Add(oFunc.AppendToCell(lblMaxPage))
                    .Add(SetCellImg("imgNext", "~\imagens\next.ico", "10px", "rptView.movepage(1);", , "Próxima página"))       'Próxima                            
                    .Add(SetCellImg("imgLast", "~\imagens\last.ico", "10px", "rptView.movepage(2);", , "Última página"))        'Última                        
                End If
                '----------------------------------------
                If ShowVoltarPagina Then
                    .Add(oFunc.RetCellFormat("rptDivisor2", "|", , "30px"))
                    If URL.Trim <> String.Empty Then
                        .Add(SetCellImg("imgHistoryBack", "~\imagens\volta.gif", "0px", Page.ClientScript.GetPostBackEventReference(Me, "1$" & URL), , ToolTipBackButton)) 'Página anterior
                    Else
                        .Add(SetCellImg("imgHistoryBack", "~\imagens\volta.gif", "0px", "rptView.previouspage();", , ToolTipBackButton)) 'Página anterior
                    End If
                End If
                '----------------------------------------
            End With
            '--------------------------
            .Rows.Add(Row)
            '--------------------------
            'Else
            '#######################################
            '#######################################
            '######   ########     #######    ######
            '##### ### ####### ##### #### #### #####
            '#### ##### ###### #  ## #### ##########
            '### ####### ##### #### ##### ##########
            '###         ##### #  ## #### ##########
            '### ####### ##### ##### ##### ### #####
            '### ####### #####      #######   ######
            '#######################################
            'End If
        End With
        '-----------------------------------
        ' Acrescenta o título        
        CellGeral = New TableCell
        With CellGeral
            .Style.Add("height", "auto")
            With .Controls
                .Add(tabTitulo)
                .Add(New LiteralControl("<div id=""divParametros"" style=""position:relative;overflow:auto;"">"))
            End With
        End With
        '---------------------------------        
        With tabParametros
            .ID = "rptParametros"
            .CellPadding = 0
            .CellSpacing = 2
            With .Style
                .Add("width", "100%")
                .Add("height", "0px")
            End With
            '------------------------------
            With .Rows()
                '-----------------------------------
                Dim IndiceColuna As Integer = 0
                For Each colPrm As ColunaParam In ColunasParam.Values                  
                    '----------------------------
                    Row = New TableRow                
                    '----------------------------
                    With colPrm
                        '----------------------------
                        ' Rótulo do campo
                        If EstiloCaption = TEstiloCaption.LEFT Then
                            Row.Cells.Add(oFunc.AppendToCell(New LiteralControl("<span>" & .Titulo.Replace(" ", "&nbsp;") & "</span>"), Alignment:=HorizontalAlign.Right, strWidth:="40%", strHeight:="0px", VAlignment:=VerticalAlign.Top, NoWrap:=True))
                            ' Insere imagem para mensagens de alerta
                            Row.Cells(0).Controls.AddAt(0, GetImgAlerta("img" & colPrm.Nome & "_" & IndiceColuna))                            
                        Else
                            Row.Cells.Add(oFunc.RetCellFormat("", "&nbsp;", , "33%", "0px", VAlignment:=VerticalAlign.Top))
                            Row.Cells.Add(oFunc.AppendToCell(New LiteralControl("<span>" & .Titulo.Replace(" ", "&nbsp;") & "</span>"), strHeight:="auto", intColspan:=3, VAlignment:=VerticalAlign.Top, NoWrap:=True))
                            Row.Cells(1).Controls.Add(GetImgAlerta("img" & colPrm.Nome & "_" & IndiceColuna))
                            ' Insere imagem para mensagens de alerta
                            tabParametros.Rows.Add(Row)
                            Row = New TableRow
                            Row.Cells.Add(oFunc.RetCellFormat("", "&nbsp;", , "40%", "0px", VAlignment:=VerticalAlign.Top))
                        End If
                        '----------------------------
                        ' Tipo de campo para exibição dos dados
                        Select Case .TipoCampo
                            Case ColunaParam.TTipoCampo.TEXT
                                '----------------------------------------
                                oField = New TextBox
                                oField.ID = colPrm.Nome & "_" & IndiceColuna
                                oField.Width = colPrm.Width
                                oField.Text = colPrm.Value
                                '----------------------------------------

                            Case ColunaParam.TTipoCampo.COMBO, ColunaParam.TTipoCampo.LISTA_MULTISELECT
                                '----------------------------------------
                                If colPrm.TipoCampo = ColunaParam.TTipoCampo.LISTA_MULTISELECT Then
                                    oField = New ListBox
                                    oField.SelectionMode = ListSelectionMode.Multiple
                                    oField.Rows = .ListBoxLinhas
                                Else
                                    oField = New DropDownList
                                End If
                                '-------------------------------------
                                With oField
                                    '---------------------------------
                                    .ID = colPrm.Nome & "_" & IndiceColuna
                                    .Width = colPrm.Width
                                    '---------------------------------
                                    ' Carregando as informações do combo/lista.
                                    If colPrm.TipoLista = ColunaParam.TTipoLista.SQL Then  'Informações do banco
                                        '-----------------------------
                                        Dim dt As New DataTable
                                        '-----------------------------
                                        ' POR QUESTÕES DE COMPATIBILIDADE (HAHAHA...P*@^.%)                                        
                                        If colPrm.ConnectionString.Trim <> String.Empty Then
                                            oDB = New ClsDB(colPrm.ConnectionString, colPrm.Provider)
                                        Else
                                            oDB = New ClsDB(ConnectionString, ClsDB.T_PROVIDER.OLEDB) ' MANTENDO A JACA FU...
                                        End If
                                        '-----------------------------
                                        dta = oDB.GetDataAdapter(colPrm.Parametros)
                                        dta.Fill(dt)
                                        '-----------------------------
                                        If dt.Rows.Count > 0 Then
                                            .DataSource = dt
                                            .DataValueField = dt.Columns(0).ColumnName
                                            .DataTextField = dt.Columns(1).ColumnName
                                            .DataBind()
                                        End If
                                        '-----------------------------
                                    Else            'de uma lista de valores
                                        '-----------------------------
                                        Dim aSep As Array = colPrm.Separador.ToString.ToCharArray
                                        Dim aItens As Array
                                        '-----------------------------
                                        ' Quebra a lista de parâmetros passada
                                        ' de acordo com o primeiro separador
                                        aItens = colPrm.Parametros.Split(aSep(0))
                                        '-----------------------------
                                        ' Se foi passado texto e valor, ou seja,
                                        ' mais de um separador, então
                                        If aSep.Length = 2 Then
                                            Dim aItens2 As Array
                                            '---------------------------------
                                            For i As Integer = 0 To aItens.Length - 1
                                                aItens2 = aItens(i).ToString.Split(aSep(1))
                                                .Items.Add(New ListItem(aItens2(1), aItens2(0)))
                                            Next
                                            '---------------------------------
                                        Else
                                            '---------------------------------
                                            For i As Integer = 0 To aItens.Length - 1
                                                .Items.Add(New ListItem(aItens(i)))
                                            Next
                                            '---------------------------------
                                        End If
                                        '---------------------------------
                                    End If
                                    '---------------------------------
                                    ' insere um ítem em branco na primeira linha
                                    .Items.Insert(0, "")
                                    '---------------------------------
                                    If colPrm.Value.Trim <> "" Then
                                        Call RestauraSelecionados(colPrm, oField)
                                    End If
                                    '---------------------------------
                                End With
                                '---------------------------------
                        End Select
                        '-----------------------------------
                        ' Adiciona o campo  
                        oField.ToolTip = colPrm.ToolTip
                        oField.Attributes.Add("rpt_param", colPrm.Nome)
                        Row.Cells.Add(oFunc.AppendToCell(oField, , HorizontalAlign.Left, "10%", NoWrap:=True))
                        '-----------------------------------
                        ' Inserindo validações 
                        Call InsereValidacao(colPrm, oField)
                        '-----------------------------------
                        ' Montando os operadores
                        Dim cboOperador As New DropDownList
                        '-----------------------------------
                        With cboOperador
                            .ID = colPrm.Nome & "_" & IndiceColuna & "_Operador"                            
                            .Width = Unit.Pixel(125)
                            '---------------------------------------
                            ' Atrela o parametro ao seu operador
                            oField.Attributes.Add("rpt_param", .ID)
                            '---------------------------------------
                        End With
                        '-----------------------------------
                        If .TipoCampo <> ColunaParam.TTipoCampo.LISTA_MULTISELECT Then
                            If .Operador = ColunaParam.TOperador.UNDEFINED Then
                                '-----------------------------------
                                ' Verifica o tipo de operadores a serem carregados
                                Dim hash As Hashtable = IIf(.TipoColuna = ColunaParam.TTipoColuna.NUMERICO, _
                                                             HASH_OperadoresNum, HASH_Operadores)
                                Dim lstitem As ListItem
                                '----------------------------------
                                For Each k As Short In hash.Keys
                                    If k <> "0" Then ' DIFERENTE DE UNDEFINED
                                        '----------------------------------
                                        lstitem = New ListItem
                                        '----------------------------------
                                        If .OperadorDefinido = k Then lstitem.Selected = True
                                        '----------------------------------
                                        With lstitem
                                            '.Text = hash(k)
                                            .Text = HASH_TradOperadores(k)
                                            .Value = k
                                        End With
                                        '----------------------------------
                                        cboOperador.Items.Add(lstitem)
                                    End If
                                Next
                                '-----------------------------------
                                cboOperador.Items.Insert(0, "")
                                '-----------------------------------
                                ' Impõe obrigatoriedade em definir o operador
                                With cboOperador.Attributes
                                    .Add("valida", IIf(colPrm.Obrigatorio = ColunaParam.TObrigatorio.SIM, "true", "false"))
                                    .Add("valida_msg", "Informe o operador para " & colPrm.Titulo)
                                End With
                            Else
                                cboOperador.Style.Add("display", "none")
                                cboOperador.Items.Add(New ListItem(HASH_TradOperadores(CShort(.Operador)), .Operador))
                            End If
                        Else
                            cboOperador.Style.Add("display", "none")
                            cboOperador.Items.Add(New ListItem(HASH_TradOperadores(CShort(ColunaParam.TOperador.DENTRO_DE)), ColunaParam.TOperador.DENTRO_DE))
                        End If
                        '-----------------------------------
                        ' Insere o campo operadores
                        Row.Cells.Add(oFunc.AppendToCell(cboOperador, Alignment:=HorizontalAlign.Left, strWidth:="40%"))
                        With Row.Cells(Row.Cells.Count - 2).Controls
                            .Add(New LiteralControl("<span id='rptCelFormato'>&nbsp;&nbsp;" & colPrm.DicaFormatoColuna.Replace(" ", "&nbsp;") & "</span>"))
                        End With
                        '-----------------------------------                      
                        ' Insere a linha parâmetro
                        tabParametros.Rows.Add(Row)
                        '-----------------------------------
                    End With
                    '-----------------------------------
                    IndiceColuna += 1
                    '-----------------------------------
                Next
            End With
        End With
        '---------------------------------        
        With CellGeral.Controls
            .Add(tabParametros)
            .Add(New LiteralControl("</div>"))
        End With
        RowGeral.Cells.Add(CellGeral)
        '---------------------------------                
        RowGeral.Cells.Add(SetCellImg("imgHideShow", "~\imagens\fechar.gif", "0px", "rptView.ocultaexibeparam('divParametros',document.getElementById('rptParametros').offsetHeight, this);", HorizontalAlign.Right, "Ocultar", VerticalAlign.Top)) ' Exibir - ocultar
        tabGeral.Rows.Add(RowGeral)
        Me.Controls.Add(tabGeral)
        '---------------------------------        
    End Sub

    Private Sub InsereValidacao(ByRef col As ColunaParam, ByRef oField As Object)
        '-----------------------------------
        With col
            ' Verifica se o campo é obrigatório
            If .Obrigatorio = ColunaParam.TObrigatorio.SIM Then
                oField.Attributes.Add("valida", "true")
            Else
                oField.Attributes.Add("valida", "false")
            End If
            oField.Attributes.Add("valida_msg", "Campo " & .Titulo & " é obrigatório!")
            '-----------------------------------
            ' Verifica se existe um valor máximo
            If .ValorMaximo.Trim <> String.Empty Then
                oField.Attributes.Add("valor_max", .ValorMaximo)
                oField.Attributes.Add("valor_max_msg", "Excedeu valor máximo: " & .ValorMaximo)
            End If
            '-----------------------------------
            ' Verifica se existe um valor mínimo
            If .ValorMinimo.Trim <> String.Empty Then
                oField.Attributes.Add("valor_min", .ValorMaximo)
                oField.Attributes.Add("valor_min_msg", "Excedeu valor mínimo: " & .ValorMinimo)
            End If
            '-----------------------------------
            Select Case .TipoColuna
                Case ColunaParam.TTipoColuna.DATA
                    oField.Attributes.Add("validavalor", "DATE")

                Case ColunaParam.TTipoColuna.MONEY
                    oField.Attributes.Add("validavalor", "NUMEROREAL")

                Case ColunaParam.TTipoColuna.NUMERICO
                    oField.Attributes.Add("validavalor", "NUMERO")

            End Select
        End With
    End Sub

    Private Sub RestauraSelecionados(ByRef colPrm As ColunaParam, ByRef objSelect As Object)
        Dim aItens As Array = colPrm.Value.Split(",")
        '---------------------------------
        If colPrm.TipoCampo = ColunaParam.TTipoCampo.COMBO Then
            Dim obj As DropDownList = objSelect
            If colPrm.TipoColuna = ColunaParam.TTipoColuna.CARACTER Then
                obj.SelectedValue = RetSemAspas(aItens(0))
            Else
                obj.SelectedValue = aItens(0)
            End If
        Else
            Dim obj As ListBox = objSelect
            Dim intIndice As Int16 = 1
            Dim itemList As ListItem
            '---------------------------------
            With obj.Items
                If colPrm.TipoColuna = ColunaParam.TTipoColuna.CARACTER Then
                    For Each sItem As String In aItens
                        itemList = .FindByValue(RetSemAspas(sItem))
                        Call SelecionaItem(obj, itemList, intIndice)
                    Next
                Else
                    For Each sItem As String In aItens
                        itemList = .FindByValue(sItem)
                        Call SelecionaItem(obj, itemList, intIndice)
                    Next
                End If
            End With
            '---------------------------------
        End If
        '---------------------------------
    End Sub

    Private Sub SelecionaItem(ByRef lst As ListBox, ByRef itemList As ListItem, ByRef intIndice As Int16)
        With lst.Items
            .Remove(itemList)
            .Insert(intIndice, itemList)
            itemList.Selected = True
            intIndice += 1
        End With
    End Sub

    Private Function RetSemAspas(ByVal strValor As String) As String
        strValor = Replace(strValor.ToString, "'", "", Count:=1)
        strValor = IIf(Right(strValor, 1) = "'", Mid(strValor, 1, strValor.Length - 1), strValor)
        Return strValor
    End Function

    Public Function GetWhere() As String
        If TipoConsulta = ClsRPTView.TTipoConsulta.SQL Then
            Return ConstroiSQL()
        Else
            Return ConstroiPrmProc()
        End If
    End Function

    Private Function ConstroiSQL() As String
        Dim strWhere As New StringBuilder
        Dim strAnd As String = " AND "
        Dim strMensagem As String = String.Empty
        '-------------------------------------
        For Each col As ColunaParam In mColunasParam.Values
            '-------------------------------------
            strMensagem = ClsTools.CheckCommandSQL(col.Value)
            If strMensagem <> String.Empty Then Throw New Exception(strMensagem)
            '-------------------------------------
            With strWhere
                If col.IsValidParam() Then
                    '------------------------------------------
                    If .ToString.Trim <> String.Empty Then .Append(strAnd)
                    '------------------------------------------
                    If col.TipoCampo = ColunaParam.TTipoCampo.LISTA_MULTISELECT Then
                        Dim aValue As Array
                        Dim strValorListBox As New StringBuilder
                        '---------------------------------
                        aValue = col.Value.Split(",")
                        '---------------------------------
                        If aValue.Length > 1 Then
                            With strValorListBox
                                '---------------------------------
                                .Append(col.Nome).Append(" IN(")
                                If col.TipoColuna = ColunaParam.TTipoColuna.CARACTER Then
                                    .Append(col.Value)
                                Else
                                    For Each s As String In aValue
                                        If .ToString.Trim <> col.Nome & " IN(" Then .Append(",")
                                        .Append(FormatarValor(s, col))
                                        '.Append(col.ValorFormatado)
                                    Next
                                End If
                                .Append(")")
                                '---------------------------------
                            End With
                            .Append(strValorListBox.ToString)
                            '---------------------------------
                        Else
                            .Append(col.Nome).Append(" ").Append(col.IncluirOperador())
                        End If
                    Else
                        .Append(col.Nome).Append(" ").Append(col.IncluirOperador())
                    End If
                    '---------------------------------                    
                End If
            End With
        Next
        '---------------------------------        
        If strWhere.ToString.Trim <> String.Empty Then
            Return " (" & strWhere.ToString & ") "
        Else
            Return String.Empty
        End If
        '---------------------------------
    End Function

    Private Function FormatarValor(ByVal mValue As Object, ByVal coluna As ColunaParam) As String
        Dim oData As New Object
        With coluna
            Select Case .TipoColuna
                Case ColunaParam.TTipoColuna.CARACTER
                    Return "'" & mValue.Trim & "'"
                Case ColunaParam.TTipoColuna.DATA
                    '--------------------------------------------------
                    If mValue.Trim <> String.Empty Then
                        If .OperadorDefinido = ColunaParam.TOperador.MENOR_IGUAL Then
                            oData = DateAdd(DateInterval.Day, 1, CDate(mValue))
                        End If
                        '--------------------------------------------------
                        If .FormatoColuna.Trim() = "" Then
                            Return "'" & Format(oData, "yyyy-MM-dd") & "'"
                        Else
                            Return "'" & Format(oData, .FormatoColuna) & "'"
                        End If
                    Else
                        Return "NULL"
                    End If
                    '--------------------------------------------------
                Case ColunaParam.TTipoColuna.MONEY, ColunaParam.TTipoColuna.NUMERICO
                    Return mValue.ToString.Replace(".", "").Replace(",", ".")
                Case Else
                    Return mValue
            End Select
        End With
    End Function

    Private Function ConstroiPrmProc()
        Dim strWhere As New StringBuilder
        Dim strMensagem As String = String.Empty
        '---------------------------------        
        For Each col As ColunaParam In mColunasParam.Values
            '---------------------------------        
            strMensagem = ClsTools.CheckCommandSQL(col.Value)
            If strMensagem <> String.Empty Then Throw New Exception(strMensagem)
            '---------------------------------        
            With strWhere
                If .ToString.Trim <> String.Empty Then .Append(", ")
                .Append("@").Append(col.Nome).Append("=").Append(col.ValorFormatado(True))
            End With
        Next
        '---------------------------------        
        Return " " & strWhere.ToString & " "
        '---------------------------------
    End Function

#End Region


End Class

<System.Serializable()> _
Public Class ColunasParam
    Inherits Dictionary(Of String, ColunaParam)


    Private mColunaParam As ColunaParam

    ''' <summary>
    ''' Adiciona uma nova coluna ao objeto de parâmetros
    ''' </summary>
    ''' <param name="ColunaParamNome">Nome da coluna (Mesmo nome da coluna no banco de dados
    ''' caso esteja sendo usado o retorno da cláusula where pelo ClsRPTParam)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>    
    Public Overloads Function Add(ByVal ColunaParamNome As String) As ColunaParam
        Dim strKeyColunaParam As String = MyBase.Count.ToString
        mColunaParam = New ColunaParam()
        '-------------------------------
        With mColunaParam
            .Nome = ColunaParamNome
            .Titulo = ColunaParamNome
            .TipoColuna = ColunaParam.TTipoColuna.CARACTER
        End With
        '-------------------------------
        MyBase.Add(strKeyColunaParam, mColunaParam)
        '-------------------------------
        Return MyBase.Item(strKeyColunaParam)
    End Function

    ''' <summary>
    ''' Adiciona uma nova coluna ao objeto de parâmetros
    ''' </summary>
    ''' <param name="pColunaParam">Um Objeto ColunaParam</param> 
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function Add(ByVal pColunaParam As ColunaParam) As ColunaParam
        Dim strKeyColunaParam As String = MyBase.Count.ToString
        '-------------------------------
        MyBase.Add(strKeyColunaParam, pColunaParam)
        '-------------------------------
        Return MyBase.Item(strKeyColunaParam)
    End Function

End Class


<System.Serializable()> _
Public Class ColunaParam

#Region "Fields"

    Private mParametros As String = String.Empty
    Private mDivColParametros As String = String.Empty
    Private mToolTip As String = String.Empty
    Private mNome As String = String.Empty
    Private mTitulo As String = String.Empty
    Private mValue As String = String.Empty
    Private mNomeColValue As String = String.Empty
    Private mNomeColText As String = String.Empty
    Private mFormatoColuna As String = String.Empty
    Private mTipFormatoColuna As String = String.Empty
    Private mConnectionString As String = String.Empty
    Private mValorMaximo As String = String.Empty
    Private mValorMinimo As String = String.Empty
    '----------------------------------------
    Private mWidth As Unit = Unit.Pixel(70)
    Private mHeight As Unit = Unit.Pixel(10)
    '----------------------------------------
    Private mOperador As TOperador = TOperador.UNDEFINED
    Private mOperadorDefinido As TOperador = TOperador.UNDEFINED
    Private mTipoLista As TTipoLista = TTipoLista.SQL
    Private mTipoColuna As TTipoColuna = TTipoColuna.CARACTER
    Private mTipoCampo As TTipoCampo = TTipoCampo.TEXT
    Private mProvider As ClsDB.T_PROVIDER = ClsDB.T_PROVIDER.OLEDB
    Private mObrigatorio As TObrigatorio = TObrigatorio.NAO
    '----------------------------------------        
    ''' <summary>
    ''' Se o parâmetro for do tipo CARACTER, ele será 
    ''' retornado ou não delimitado com aspas simples.
    ''' </summary>
    ''' <remarks></remarks>
    Private mValorDelimitado As Boolean
    '----------------------------------------
    Private HASH_TTipoColuna As Hashtable
    Private HASH_TTipoLista As Hashtable
    Private HASH_TTipoCampo As Hashtable
    Private HASH_Operadores As Hashtable
    Private HASH_Obrigatorio As Hashtable
    '----------------------------------------
    Private mListBoxLinhas As Short = 5

#End Region

#Region "Constructors"

    Public Sub New()
        '--------------------------
        HASH_TTipoColuna = New Hashtable
        With HASH_TTipoColuna
            .Add(CShort(TTipoColuna.CARACTER), TTipoColuna.CARACTER)
            .Add(CShort(TTipoColuna.DATA), TTipoColuna.DATA)
            .Add(CShort(TTipoColuna.MONEY), TTipoColuna.MONEY)
            .Add(CShort(TTipoColuna.NUMERICO), TTipoColuna.NUMERICO)
        End With
        '--------------------------
        HASH_TTipoLista = New Hashtable
        With HASH_TTipoLista
            .Add(CShort(TTipoLista.UNDEFINED), TTipoLista.UNDEFINED)
            .Add(CShort(TTipoLista.LISTA), TTipoLista.LISTA)
            .Add(CShort(TTipoLista.SQL), TTipoLista.SQL)
        End With
        '--------------------------
        HASH_TTipoCampo = New Hashtable
        With HASH_TTipoCampo
            .Add(CShort(TTipoCampo.COMBO), TTipoCampo.COMBO)
            .Add(CShort(TTipoCampo.TEXT), TTipoCampo.TEXT)
            .Add(CShort(TTipoCampo.LISTA_MULTISELECT), TTipoCampo.LISTA_MULTISELECT)
        End With
        '--------------------------
        HASH_Operadores = New Hashtable
        With HASH_Operadores
            .Add(CShort(ColunaParam.TOperador.UNDEFINED), "")
            .Add(CShort(ColunaParam.TOperador.IGUAL), " = ")
            .Add(CShort(ColunaParam.TOperador.MAIOR), " > ")
            .Add(CShort(ColunaParam.TOperador.MENOR), " < ")
            .Add(CShort(ColunaParam.TOperador.MAIOR_IGUAL), " >= ")
            .Add(CShort(ColunaParam.TOperador.MENOR_IGUAL), " <= ")
            .Add(CShort(ColunaParam.TOperador.A_PARTIR_DE), " Like '#%' ")
            .Add(CShort(ColunaParam.TOperador.TERMINADO_EM), " Like '%#' ")
            .Add(CShort(ColunaParam.TOperador.CONTENDO), " Like '%#%' ")
            .Add(CShort(ColunaParam.TOperador.DENTRO_DE), " IN(#) ")
        End With
        '--------------------------
        HASH_Obrigatorio = New Hashtable
        With HASH_Operadores
            .Add(CShort(TObrigatorio.SIM), TObrigatorio.SIM)
            .Add(CShort(TObrigatorio.NAO), TObrigatorio.NAO)
        End With
        '--------------------------
    End Sub

#End Region

#Region "Enumerations"

    Public Enum TTipoColuna As Short
        CARACTER = 27
        NUMERICO = 28
        DATA = 29
        MONEY = 30
    End Enum

    Public Enum TTipoLista
        UNDEFINED = 0
        SQL = 11
        LISTA = 12
    End Enum

    Public Enum TTipoCampo
        TEXT = 22
        COMBO = 25
        LISTA_MULTISELECT = 26
    End Enum

    Public Enum TOperador As Short
        UNDEFINED = 0
        IGUAL = 13
        MAIOR = 14
        MENOR = 15
        MAIOR_IGUAL = 16
        MENOR_IGUAL = 17
        A_PARTIR_DE = 18
        TERMINADO_EM = 19
        CONTENDO = 20
        DENTRO_DE = 21
    End Enum

    Public Enum TObrigatorio
        SIM = 31
        NAO = 32
    End Enum    

#End Region

#Region "Properties"

    ''' <summary>
    ''' Valor mínimo que o parâmetro pode assumir
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ValorMinimo() As String
        Get
            Return mValorMinimo
        End Get
        Set(ByVal value As String)
            mValorMinimo = value
        End Set
    End Property

    ''' <summary>
    ''' Valor máximo que o parâmetro pode assumir
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ValorMaximo() As String
        Get
            Return mValorMaximo
        End Get
        Set(ByVal value As String)
            mValorMaximo = value
        End Set
    End Property

    ''' <summary>
    ''' Largura do campo parâmetro
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Width() As Unit
        Get
            Return mWidth
        End Get
        Set(ByVal value As Unit)
            mWidth = value
        End Set
    End Property

    ''' <summary>
    ''' Altura do campo parâmetro
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Height() As Unit
        Get
            Return mHeight
        End Get
        Set(ByVal value As Unit)
            mHeight = value
        End Set
    End Property

    ''' <summary>
    ''' Valor informado pelo usuário para o campo parâmetro.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Value() As String
        Get
            Dim strValor As String = String.Empty
            '------------------------------------
            If mValorDelimitado AndAlso (mValue <> String.Empty And mTipoCampo = TTipoCampo.LISTA_MULTISELECT) Then
                strValor = "'" & mValue.Replace(",", "','") & "'"
            Else
                strValor = mValue
            End If
            '------------------------------------
            Return strValor
            '------------------------------------
        End Get
        Set(ByVal value As String)
            mValue = value
        End Set
    End Property

    ''' <summary>
    ''' Texto de dica para a coluna parâmetro.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ToolTip() As String
        Get
            Return mToolTip
        End Get
        Set(ByVal value As String)
            mToolTip = value
        End Set
    End Property

    ''' <summary>
    ''' Nome da coluna parâmetro. 
    ''' Deve ser o mesmo da coluna do relatório caso
    ''' esteja sendo usado o retorno da cláusula where pela ClsRPTParam.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Nome() As String
        Get
            Return mNome
        End Get
        Set(ByVal value As String)
            mNome = value
        End Set
    End Property

    ''' <summary>    
    ''' Rótulo da coluna parâmetro.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Titulo() As String
        Get
            Return mTitulo
        End Get
        Set(ByVal value As String)
            mTitulo = value
        End Set
    End Property

    ''' <summary>
    ''' Define o provider do Framework que será utilizado 
    ''' carregar o campo se o mesmo for do tipo type = [ select-one | select-multiple 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Provider() As ClsDB.T_PROVIDER
        Get
            Return mProvider
        End Get
        Set(ByVal value As ClsDB.T_PROVIDER)
            mProvider = value
        End Set
    End Property

    ''' <summary>
    ''' String de Conexão para a coluna parâmetro.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ConnectionString() As String
        Get
            Return mConnectionString
        End Get
        Set(ByVal value As String)
            mConnectionString = value
        End Set
    End Property

    ''' <summary>
    ''' Nome da coluna que será o identificador em um 
    ''' ListBox ou DropDownList, caso a propriedade TipoLista for SQL.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Caso não informado, assume NomeColText. 
    ''' Se ambos não forem informados, será assumido a ordem dos 
    ''' campos da Query informada.
    ''' </remarks>
    Public Property NomeColValue() As String
        Get
            Return mNomeColValue
        End Get
        Set(ByVal value As String)
            mNomeColValue = value
        End Set
    End Property

    ''' <summary>
    ''' Nome da coluna que será exibida em um 
    ''' ListBox ou DropDownList, caso a propriedade TipoLista for SQL.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Caso não informado, assume NomeColValue. 
    ''' Se ambos não forem informados, será assumido a ordem dos 
    ''' campos da Query informada.
    ''' </remarks>
    Public Property NomeColText() As String
        Get
            Return mNomeColText
        End Get
        Set(ByVal value As String)
            mNomeColText = value
        End Set
    End Property

    ''' <summary>    
    ''' Operadores utilizados para a construção da cláusula WHERE.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' <list type="bullet">
    ''' <item>
    ''' <description>1 - = Igual</description>
    ''' </item>
    ''' <item>
    ''' <description>2 - &gt; Maior </description>
    ''' </item>
    ''' <item>
    ''' 3 - &lt; Menor
    ''' </item>
    ''' <item>
    ''' 4 - &gt;= Maior ou Igual
    ''' </item>
    ''' <item>
    ''' 5 - &lt;= Menor ou Igual
    ''' </item>
    ''' <item>
    ''' 6 - Like '%#'
    ''' </item>
    ''' <item>
    ''' 7 - Like ''#%'
    ''' </item>
    ''' <item>
    ''' 8 - Like '%#%'
    ''' </item>
    ''' </list>
    ''' </remarks>
    Public Property Operador() As TOperador
        Get
            Return mOperador
        End Get
        Set(ByVal value As TOperador)
            mOperador = TrataProp(value, HASH_Operadores)
            mOperadorDefinido = mOperador
        End Set
    End Property

    Friend Property OperadorDefinido() As TOperador
        Get
            Return mOperadorDefinido
        End Get
        Set(ByVal value As TOperador)
            mOperadorDefinido = TrataProp(value, HASH_Operadores)
        End Set
    End Property

    ''' <summary>
    ''' Obriga preenchimento do campo 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Obrigatorio() As TObrigatorio
        Get
            Return mObrigatorio
        End Get
        Set(ByVal value As TObrigatorio)
            mObrigatorio = value
        End Set
    End Property

    ''' <summary>
    ''' Separador(es) utilizado(s) para divisão dos valores passados 
    ''' como lista.
    ''' Ex.: ",#"
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Separador() As String
        Get
            Return mDivColParametros
        End Get
        Set(ByVal value As String)
            mDivColParametros = value
        End Set
    End Property

    ''' <summary>    
    ''' Tipo-Lista
    '''    Parâmetro = 2005,2006,2007
    '''    Div-Col-Parametro = ,    
    ''' Tipo-Lista 
    '''    SQL = SELECT colValue, colText FROM Tabela    
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Parametros() As String
        Get
            Return mParametros
        End Get
        Set(ByVal value As String)
            mParametros = value
        End Set
    End Property

    ''' <summary>
    ''' Se é uma coluna do tipo Caracter, Numérica, Data, Etc.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TipoColuna() As TTipoColuna
        Get
            Return mTipoColuna
        End Get
        Set(ByVal value As TTipoColuna)
            mTipoColuna = TrataProp(value, HASH_TTipoColuna)
        End Set
    End Property


    ''' <summary>
    ''' Se, quando uma lista de seleção única ou múltipla, a fonte
    ''' de dados é uma lista delimitada ou um fonte de dados SQL.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TipoLista() As TTipoLista
        Get
            Return mTipoLista
        End Get
        Set(ByVal value As TTipoLista)
            mTipoLista = TrataProp(value, HASH_TTipoLista)
        End Set
    End Property

    ''' <summary>
    ''' Define o tipo de campo, se TEXTO, COMBO, LISTA.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TipoCampo() As TTipoCampo
        Get
            Return mTipoCampo
        End Get
        Set(ByVal value As TTipoCampo)
            mTipoCampo = TrataProp(value, HASH_TTipoCampo)
        End Set
    End Property

    ''' <summary>
    ''' Define o número de linhas que terá a Lista.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ListBoxLinhas() As Short
        Get
            Return mListBoxLinhas
        End Get
        Set(ByVal value As Short)
            mListBoxLinhas = value
        End Set
    End Property

    ''' <summary>
    ''' Informação para exibição na tela do formato 
    ''' de entrada esperado.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DicaFormatoColuna() As String
        Get
            Return mTipFormatoColuna
        End Get
        Set(ByVal value As String)
            mTipFormatoColuna = value
        End Set
    End Property

    ''' <summary>
    ''' Define a formatação da coluna parâmetro.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FormatoColuna() As String
        Get
            Return mFormatoColuna
        End Get
        Set(ByVal value As String)
            mFormatoColuna = value
        End Set
    End Property

    Private Function TrataProp(ByVal ValProp As Short, ByRef hsh As Hashtable) As Integer
        If hsh.ContainsKey(ValProp) Then
            Return ValProp
        Else
            Throw New Exception("Valor informado é inválido.")
        End If
    End Function

    ''' <summary>
    ''' Se o parâmetro for do tipo CARACTER, ele será 
    ''' retornado ou não delimitado com aspas simples.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ValorDelimitado() As Boolean
        Get
            Return mValorDelimitado
        End Get
        Set(ByVal value As Boolean)
            mValorDelimitado = value
        End Set
    End Property

#End Region

#Region "Methods"

    'Public Function ValorFormatado() As String
    '    Select Case mTipoColuna
    '        Case ColunaParam.TTipoColuna.CARACTER
    '            Return "'" & mValue.Trim & "'"
    '        Case ColunaParam.TTipoColuna.DATA
    '            If mFormatoColuna.Trim() = "" Then
    '                Return "'" & Format(CDate(mValue), "yyyy-MM-dd") & IIf(TipoColuna = TTipoColuna.DATA And mOperadorDefinido = TOperador.MENOR_IGUAL, " 23:59:59.998", "") & "'"
    '            Else
    '                Return "'" & Format(CDate(mValue), mFormatoColuna) & IIf(TipoColuna = TTipoColuna.DATA And mOperadorDefinido = TOperador.MENOR_IGUAL, " 23:59:59.998", "") & "'"
    '            End If
    '        Case ColunaParam.TTipoColuna.MONEY, ColunaParam.TTipoColuna.NUMERICO
    '            Return mValue.ToString.Replace(".", "").Replace(",", ".")
    '        Case Else
    '            Return mValue
    '    End Select
    'End Function

    Public Function ValorFormatado(Optional ByVal IsProc As Boolean = False) As String
        Dim oData As New Object
        Select Case mTipoColuna
            Case ColunaParam.TTipoColuna.CARACTER
                Return "'" & mValue.Trim & "'"
            Case ColunaParam.TTipoColuna.DATA
                '--------------------------------------------------
                If mValue.Trim <> String.Empty Then
                    If mOperadorDefinido = TOperador.MENOR_IGUAL And Not IsProc Then
                        oData = DateAdd(DateInterval.Day, 1, CDate(mValue))
                    Else
                        oData = CDate(mValue)
                    End If
                    '--------------------------------------------------
                    If mFormatoColuna.Trim() = "" Then
                        Return "'" & Format(oData, "yyyy-MM-dd") & "'"
                    Else
                        Return "'" & Format(oData, mFormatoColuna) & "'"
                    End If
                Else
                    Return "NULL"
                End If
                '--------------------------------------------------
            Case ColunaParam.TTipoColuna.MONEY, ColunaParam.TTipoColuna.NUMERICO
                Return mValue.ToString.Replace(".", "").Replace(",", ".")
            Case Else
                Return mValue
        End Select
    End Function

    Protected Friend Function IncluirOperador() As String
        Select Case mOperadorDefinido
            Case TOperador.A_PARTIR_DE, TOperador.CONTENDO, TOperador.TERMINADO_EM, TOperador.DENTRO_DE
                Return HASH_Operadores(CShort(mOperadorDefinido)).Replace("#", Value.Trim)

            Case Else
                Dim shtOperador As TOperador = mOperadorDefinido
                If shtOperador = TOperador.MENOR_IGUAL And mTipoColuna = TTipoColuna.DATA Then
                    shtOperador = TOperador.MENOR
                End If
                Return HASH_Operadores(CShort(shtOperador)) & IIf(mTipoCampo = TTipoCampo.TEXT, ValorFormatado, Value.Trim)
        End Select

    End Function

    Protected Friend Function IsValidParam() As Boolean
        '---------------------------------------
        Select Case mTipoColuna
            Case ColunaParam.TTipoColuna.CARACTER
                If mValue.Trim <> String.Empty Then Return True

            Case ColunaParam.TTipoColuna.DATA
                If IsDate(mValue) Then Return True

            Case ColunaParam.TTipoColuna.MONEY, ColunaParam.TTipoColuna.NUMERICO
                If mTipoCampo = TTipoCampo.TEXT Then
                    If IsNumeric(mValue) Then Return True
                Else
                    If mValue.Trim <> String.Empty Then Return True
                End If

        End Select
        '---------------------------------------
        Return False
        '---------------------------------------
    End Function

#End Region

End Class

#End Region


'**********************************************************************************
'**********************************************************************************


Friend Class ClsFunctions

    ''' <summary>
    ''' Insere um controle em uma célula e a retorna.
    ''' </summary>
    ''' <param name="ctl">Controle HTML a ser inserido.</param>
    ''' <param name="TypeCell">Tipo da célula: 'H' - header, 'N' - normal.</param>
    ''' <param name="Alignment">Alinhamento Horizontal na célula.</param>
    ''' <param name="intColspan">Número de células que a célula irá se expandir.</param> 
    ''' <param name="intRowSpan">Número de linhas que a célula irá ocupar.</param>
    ''' <param name="strHeight">Altura da célula.</param>
    ''' <param name="strWidth">Largura da célula.</param>
    ''' <returns>Uma célula com um controle inserido.</returns>
    ''' <remarks></remarks>
    Public Function AppendToCell(ByVal ctl As Object, _
                   Optional ByVal ID As String = "", _
                   Optional ByVal Alignment As HorizontalAlign = HorizontalAlign.NotSet, _
                   Optional ByVal strWidth As String = "", _
                   Optional ByVal strHeight As String = "", _
                   Optional ByVal intRowSpan As Integer = 1, _
                   Optional ByVal intColspan As Integer = 1, _
                   Optional ByVal TypeCell As String = "N", _
                   Optional ByVal VAlignment As VerticalAlign = VerticalAlign.NotSet, _
                   Optional ByVal NoWrap As Boolean = False) As TableCell

        Dim Cell As TableCell
        '-----------------------
        If TypeCell = "H" Then
            Cell = New TableHeaderCell
        Else
            Cell = New TableCell
        End If
        '-----------------------
        With Cell
            If ID.Trim <> String.Empty Then .ID = ID
            If intColspan > 1 Then .ColumnSpan = intColspan
            If intRowSpan > 1 Then .RowSpan = intRowSpan
            If Alignment <> HorizontalAlign.NotSet Then .HorizontalAlign = Alignment
            If VAlignment <> VerticalAlign.NotSet Then .VerticalAlign = VAlignment
            If strWidth.Trim <> String.Empty Then .Style.Add("width", strWidth)
            If strHeight.Trim <> String.Empty Then .Style.Add("height", strHeight)
            If NoWrap Then
                .Style.Add("white-space", "nowrap")
                .Attributes.Add("nowrap", "nowrap")
            End If
            .Controls.Add(ctl)
        End With
        '-----------------------
        Return Cell
    End Function

    Public Function AppendToRow(ByVal ctl As Object, _
                   Optional ByVal ID As String = "", _
                   Optional ByVal Alignment As HorizontalAlign = HorizontalAlign.NotSet, _
                   Optional ByVal strWidth As String = "", _
                   Optional ByVal strHeight As String = "", _
                   Optional ByVal intRowSpan As Integer = 1, _
                   Optional ByVal intColspan As Integer = 1, _
                   Optional ByVal TypeCell As String = "N", _
                   Optional ByVal VAlignment As VerticalAlign = VerticalAlign.NotSet, _
                   Optional ByVal NoWrap As Boolean = False) As TableRow

        Dim Row As New TableRow
        Row.Cells.Add(AppendToCell(ctl, ID, Alignment, strWidth, strHeight, intRowSpan, intColspan, TypeCell, VAlignment, NoWrap))
        Return Row
    End Function

    Public Function RetCellFormat(ByVal strId As String, ByVal strTexto As String, _
                   Optional ByVal Alignment As HorizontalAlign = HorizontalAlign.Left, _
                   Optional ByVal strWidth As String = "", _
                   Optional ByVal strHeight As String = "", _
                   Optional ByVal intRowSpan As Integer = 1, _
                   Optional ByVal intColspan As Integer = 1, _
                   Optional ByVal TypeCell As String = "N", _
                   Optional ByVal VAlignment As VerticalAlign = VerticalAlign.NotSet, _
                   Optional ByVal ToolTip As String = "", _
                   Optional ByVal NoWrap As Boolean = False) As TableCell
        Dim Cell As TableCell
        '--------------------------------
        If TypeCell.ToUpper = "N" Then
            Cell = New TableCell
        Else
            Cell = New TableHeaderCell
        End If
        '--------------------------------
        With Cell
            If strId.Trim <> String.Empty Then .ID = strId
            If ToolTip.Trim <> String.Empty Then .ToolTip = ToolTip
            If VAlignment <> VerticalAlign.NotSet Then .VerticalAlign = VAlignment
            .Text = IIf(strTexto.Trim = "", "&nbsp;", strTexto.Replace(vbCrLf, "<BR />"))
            If strWidth.Trim <> String.Empty Then .Style.Add("width", strWidth)
            If strHeight.Trim <> String.Empty Then .Style.Add("height", strHeight)
            If Alignment <> HorizontalAlign.NotSet Then .HorizontalAlign = Alignment
            If intColspan > 1 Then .ColumnSpan = intColspan
            If intRowSpan > 1 Then .RowSpan = intRowSpan
            If NoWrap Then
                .Style.Add("white-space", "nowrap")
                .Attributes.Add("nowrap", "nowrap")
            End If

        End With
        Return Cell
    End Function

End Class

