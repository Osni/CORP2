Imports System
Imports System.Data
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Collections.Specialized

<DefaultProperty("Text"), ToolboxData("<{0}:ClsFilter runat=server></{0}:ClsFilter>")> <System.Serializable()> _
Public Class ClsFilter
    Inherits WebControl : Implements IPostBackDataHandler : Implements IPostBackEventHandler
    '############################################################################
    Public Event FILTER_CellClick(ByVal sText As String, ByVal sValue As String)
    '############################################################################
#Region "Fields"
    Private _FilterViewState As Collection
    Private _FilterCols As Collection
    Private _ClsFilterCols As ClsFilterCols
    Private i As Integer
    Private l As Long
    Private _ds As DataSet
    Private _PageFirst As Long
    Private _PageNext As Long
    Private _PagePrevious As Long
    Private _PageLast As Long
    Private _PageLastRows As Long
    Private _PageQtde As Long
    Private _PageAtiva As Long
    Private FilterPageRowPosition As Long
    Private _TotalRows As Long
    Private _GetSQLWhere As String
    Private _FilterLastOrderByColName As String
    Private _FilterOrderByColName As String
    Private _FilterColOrderByPos As String
    Private _GetSQLOrderBy As String = ""
    Private _FILTER_VIEWSTATE As String = ""
    Private _FilterState As System.Data.PropertyCollection
    Private _FilterColText As String
    Private _FilterColValue As String
    Private _FilterSimpleText As String
    Private _FilterColSimpleText As String
    Private _FilterColSimpleValue As String
    Private _FilterNotVisible As Boolean
    Private _FilterQueryString As DataTable
    Private _FilterOptionClear As Boolean    
    Private x As Integer = 0
#End Region
    '#######################################################################
#Region "Constructors"
    Public Sub New()
        _FilterQueryString = New DataTable
        With _FilterQueryString.Columns
            .Add("key")
            .Add("value")
        End With
        _FilterCols = New Collection()
        _ClsFilterCols = New ClsFilterCols()
        _FilterState = New System.Data.PropertyCollection
    End Sub
#End Region
    '#######################################################################
#Region "Filter Eventos"
    '############################################################################
    Protected Overrides Sub CreateChildControls()
        Dim crt As LiteralControl
        Dim sNameID As String = ""
        If FilterType = PFilterType.SimpleFilter Then
            sNameID = Me.ClientID & "_FILTER_LIST"
            crt = New LiteralControl("<input style=""width:90%"" type=""text"" name=""" & sNameID & "_TEXT"" id=""" & sNameID & "_TEXT"" value=""" & _FilterColSimpleText & """ />&nbsp;&nbsp;&nbsp;")
            Me.Controls.Add(crt)
            crt = New LiteralControl("<a href=#><img onclick=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "filter_list_click") & """ title='List Filtro'  id='filter_list_click_img' name='filter_list_click_img' src='imagens/list_filter.gif' border=0></a>")
            Me.Controls.Add(crt)
            crt = New LiteralControl("<input type=""hidden"" name=""" & sNameID & "_VALUE"" id=""" & sNameID & "_VALUE"" value=""" & _FilterColSimpleValue & """  />")
            Me.Controls.Add(crt)
        End If
    End Sub
    '############################################################################
    Private Sub ClsFilter_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        '-------------------------------------------------
        Dim pgHandler As Page = CType(Me.Context.CurrentHandler, Page)
        Dim isStyle As UI.IStyleSheet = pgHandler.Header.StyleSheet
        Dim newStyle As Style
        '----------------------------------------
        Me.Page.Form.Attributes.Add("autocomplete", "off")
        '----------------------------------------
    End Sub
    '############################################################################
    Protected Overrides Sub RenderContents(ByVal output As HtmlTextWriter)
        If FilterPagePreLoad = True Or FilterType = PFilterType.FullFilter Then
            ShowFilter()
            'Me.Controls.Add(ShowFilter())
        End If
        'MyBase.RenderContents(output)
    End Sub
    '#################################################################################
    Public Function LoadPostData(ByVal PostDataKey As String, ByVal Values As NameValueCollection) As Boolean Implements IPostBackDataHandler.LoadPostData
        '_value = Int32.Parse(values(UniqueID))
        Return False
    End Function
    Public Sub RaisePostDataChangedEvent() Implements IPostBackDataHandler.RaisePostDataChangedEvent
        ' Part of the IPostBackDataHandler contract.  Invoked if we ever returned true from the
        ' LoadPostData method (indicates that we want a change notification raised).  Since we
        ' always return false, this method is just a no-op.
    End Sub
    Public Sub RaisePostBackEvent(ByVal EventArgument As String) Implements IPostBackEventHandler.RaisePostBackEvent
        Dim sEventPart As Object
        sEventPart = Split(EventArgument, "$")
        '---------------------------------------------------------
        _GetSQLWhere = GetSQLWhere()
        '---------------------------------------------------------
        If sEventPart(0).ToString().Substring(0, Len("filter_move")) = "filter_move" Then
            _PageQtde = sEventPart(2)
            _TotalRows = sEventPart(3)
            _PageLast = _PageQtde
        End If
        Select Case sEventPart(0)
            Case "filter_move_first"
                _PageLast = _PageQtde
                _PagePrevious = 0
                _PageNext = 0
                FilterPageRowPosition = 0
                _PageAtiva = 1
            Case "filter_move_next"
                _PageNext = sEventPart(1)
                _PageNext += 1
                If _PageNext >= _PageQtde Then
                    _PageNext -= 1
                End If
                FilterPageRowPosition = _PageNext * FilterPageSize
                _PagePrevious = _PageNext
                _PageAtiva = _PageNext + 1
            Case "filter_move_previous"
                _PagePrevious = sEventPart(1)
                _PagePrevious -= 1
                If _PagePrevious < 0 Then
                    _PagePrevious += 1
                End If
                FilterPageRowPosition = _PagePrevious * FilterPageSize
                _PageNext = _PagePrevious
                _PageAtiva = _PageNext + 1
            Case "filter_move_last"
                _PageNext = _PageLast - 1
                _PagePrevious = _PageLast - 1
                FilterPageRowPosition = Int(_TotalRows / FilterPageSize) * FilterPageSize
                If FilterPageRowPosition >= _TotalRows Then
                    FilterPageRowPosition = (Int(_TotalRows / FilterPageSize) - 1) * FilterPageSize
                End If
                _PageAtiva = _PageLast
            Case "filter_click_aplica"
                _TotalRows = 0
                _PageNext = 0
                _PagePrevious = 0
                _PageQtde = 0
                _PageAtiva = 0
                FilterPagePreLoad = True
            Case "filter_click_novo"
                GetClearSQLSelect()
                _GetSQLWhere = ""
                FilterPagePreLoad = False
            Case "filter_click_retornar"
            Case "filter_click_print"
            Case "filter_click_excel"
                ExportarXLS()
            Case "filter_orderby"
                GetSQLOrderBy(sEventPart(1), sEventPart(2), sEventPart(3))
            Case "filter_cellclick"
                _FilterColSimpleText = sEventPart(1)
                _FilterColSimpleValue = sEventPart(2)
                RaiseEvent FILTER_CellClick(sEventPart(1), sEventPart(2))
                _FilterNotVisible = True
            Case "filter_list_click"
                Call GetFilterSimpleText()
        End Select
    End Sub
#End Region
#Region "Filter Property"

    Public Enum PFilterType As Integer
        FullFilter = 1
        SimpleFilter = 2
    End Enum
    <Bindable(True), Category("Appearance"), DefaultValue(""), Localizable(True)> Property FilterColText() As String
        Get
            Dim s As String = CStr(ViewState("FilterColText"))
            If s Is Nothing Then
                Return String.Empty
            Else
                Return s
            End If
        End Get
        Set(ByVal Value As String)
            ViewState("FilterColText") = Value
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(""), Localizable(True)> Property FilterColValue() As String
        Get
            Dim s As String = CStr(ViewState("FilterColValue"))
            If s Is Nothing Then
                Return String.Empty
            Else
                Return s
            End If
        End Get
        Set(ByVal Value As String)
            ViewState("FilterColValue") = Value
        End Set
    End Property
    Public Property FilterProvider() As ClsDB.T_PROVIDER
        Get
            If ViewState("FilterProvider") Is Nothing Then ViewState("FilterProvider") = ClsDB.T_PROVIDER.SQL
            Return ViewState("FilterProvider")
        End Get
        Set(ByVal value As ClsDB.T_PROVIDER)
            ViewState("FilterProvider") = value
        End Set
    End Property
    Public Property FilterCol() As ClsFilterCols
        Get
            Return _ClsFilterCols
        End Get
        Set(ByVal value As ClsFilterCols)
            _ClsFilterCols = value
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(10), Localizable(True)> Property FilterPageSize() As Integer
        Get
            Dim s As String = CStr(ViewState("FilterPageSize"))
            If s Is Nothing Then
                Return 0
            Else
                Return CInt(s)
            End If
        End Get
        Set(ByVal Value As Integer)
            ViewState("FilterPageSize") = Value
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(False), Localizable(True)> Property FilterPagePreLoad() As Boolean
        Get
            Dim s As String = CStr(ViewState("FilterPagePreLoad"))
            If s Is Nothing Then
                Return False
            Else
                Return CBool(s)
            End If
        End Get
        Set(ByVal Value As Boolean)
            ViewState("FilterPagePreLoad") = Value
        End Set
    End Property
    Public Property FilterStateView() As Collection
        Get
            Dim o As Object = ViewState("FilterStateView")
            If (IsNothing(o)) Then
                Return Nothing
            Else
                Return o
            End If
        End Get
        Set(ByVal Value As Collection)
            ViewState.Add("FilterStateView", Value)
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(""), Localizable(True)> Property FilterParamWhere() As String
        Get
            Dim s As String = CStr(ViewState("FilterParamWhere"))
            If s Is Nothing Then
                Return String.Empty
            Else
                Return s
            End If
        End Get
        Set(ByVal Value As String)
            ViewState("FilterParamWhere") = Value
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(ClsFilter.PFilterType.FullFilter), Localizable(True)> Property FilterType() As PFilterType
        Get
            Dim s As String = CStr(ViewState("FilterType"))
            If s Is Nothing Then
                Return 0
            Else
                Return CInt(s)
            End If
        End Get
        Set(ByVal Value As PFilterType)
            ViewState("FilterType") = CInt(Value)
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(False), Localizable(True)> Property FilterOrderByCols() As Boolean
        Get
            Dim s As String = CStr(ViewState("FilterOrderByCols"))
            If s Is Nothing Then
                Return False
            Else
                Return CBool(s)
            End If
        End Get
        Set(ByVal Value As Boolean)
            ViewState("FilterOrderByCols") = Value
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(""), Localizable(True)> Property FilterTableName() As String
        Get
            Dim s As String = CStr(ViewState("FilterTableName"))
            If s Is Nothing Then
                Return String.Empty
            Else
                Return s
            End If
        End Get
        Set(ByVal Value As String)
            ViewState("FilterTableName") = Value
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(""), Localizable(True)> Property FilterStrConnection() As String
        Get
            Dim s As String = CStr(ViewState("FilterStrConnection"))
            If s Is Nothing Then
                Return String.Empty
            Else
                Return s
            End If
        End Get
        Set(ByVal Value As String)
            ViewState("FilterStrConnection") = Value
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(""), Localizable(True)> Property FilterReturnFormName() As String
        Get
            Dim s As String = CStr(ViewState("FilterReturnFormName"))
            If s Is Nothing Then
                Return String.Empty
            Else
                Return s
            End If
        End Get
        Set(ByVal Value As String)
            ViewState("FilterReturnFormName") = Value
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue("Filter"), Localizable(True)> Property FilterTitle() As String
        Get
            Dim s As String = CStr(ViewState("FilterTitle"))
            If s Is Nothing Then
                Return String.Empty
            Else
                Return s
            End If
        End Get
        Set(ByVal Value As String)
            ViewState("FilterTitle") = Value
        End Set
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue("Filter"), Localizable(True)> Property FilterInitialOrderBy() As String
        Get
            Dim s As String = CStr(ViewState("FilterInitialOrderBy"))
            If s Is Nothing Then
                Return String.Empty
            Else
                Return s
            End If
        End Get
        Set(ByVal Value As String)
            ViewState("FilterInitialOrderBy") = Value
        End Set
    End Property

    <Bindable(True), Category("Appearance"), DefaultValue("Filter"), Localizable(True)> Property FilterJsUrl() As String
        Get
            Dim s As String = CStr(ViewState("FilterJsUrl"))
            If s Is Nothing Then
                'Return "http://10.0.0.238/corpnet2/"
                Return MyPath()
            Else
                Return s
            End If
        End Get
        Set(ByVal Value As String)
            ViewState("FilterJsUrl") = Value
        End Set
    End Property    
    <Bindable(True), Category("Appearance"), DefaultValue(False), Localizable(True)> _
    Property FilterEncryptValue() As Boolean
        Get
            Dim s As String = CStr(ViewState("FilterEncryptValue"))
            If s Is Nothing Then
                Return False
            Else
                Return CBool(s)
            End If
        End Get
        Set(ByVal Value As Boolean)
            ViewState("FilterEncryptValue") = Value
        End Set
    End Property
#End Region
    '#######################################################################
#Region "Filter Method"
    Private Sub GetFilterSimpleText()
        Dim oRequest As System.Web.HttpRequest = System.Web.HttpContext.Current.Request
        _FilterSimpleText = oRequest(Me.ClientID & "_FILTER_LIST_TEXT")
        FilterParamWhere = FilterColText & " LIKE '%" & _FilterSimpleText.Replace("'", "''") & "%'"
    End Sub
    Public Function NullDB(ByRef pExpress As Object, Optional ByVal pReturn As Object = "") As Object
        Return IIf(IsDBNull(pExpress), pReturn, pExpress)
    End Function
    Private Sub GetSQLOrderBy(ByVal sCol As String, ByVal sPos As String, ByVal sLastCol As String)
        _FilterOrderByColName = GetViewStateFilter("ColOrderBy")
        If sCol <> _FilterOrderByColName Then
            _FilterLastOrderByColName = sCol
            _FilterColOrderByPos = "ASC"
        Else
            If sPos = "ASC" Then
                _FilterColOrderByPos = "DESC"
            Else
                _FilterColOrderByPos = "ASC"
            End If
        End If
        SetViewStateFilter("ColOrderBy", sCol)
        _FilterOrderByColName = sCol
        _GetSQLOrderBy = " ORDER BY " & _FilterOrderByColName & " " & _FilterColOrderByPos
    End Sub
    Private Function GetPositionRowPage() As Long
        If _TotalRows = 0 And FilterPagePreLoad Then
            Dim sSQL As String
            Dim ClsDb As New ClsDB(FilterStrConnection, FilterProvider)
            Dim strCount As String = 0
            sSQL = "SELECT Count(*) as total FROM " & FilterTableName & " " & _GetSQLWhere
            strCount = ClsDb.GetDataTable(sSQL).Rows(0)(0).ToString
            _TotalRows = IIf(strCount = String.Empty, 0, strCount)
            _PageLast = _TotalRows
            If _TotalRows Mod FilterPageSize = 0 Then
                _PageQtde = Int(_TotalRows / FilterPageSize)
            Else
                _PageQtde = Int(_TotalRows / FilterPageSize) + 1
            End If
            FilterPageRowPosition = 0
            _PageAtiva = 1
        End If
        Return _PageQtde
    End Function
    Private Function GetFilterData(ByVal sSQL As String) As System.Data.DataView
        If FilterStrConnection = "" Then
            Throw New System.Exception("Uma String de Conexão é obrigatório em <FilterStrConnection>")
        Else
            _PageLast = GetPositionRowPage()
            '----------------------------------
            Dim ClsDB As New ClsDB(FilterStrConnection, FilterProvider)
            _ds = New System.Data.DataSet
            '----------------------------------
            With ClsDB.GetDataAdapter()
                .SelectCommand = ClsDB.GetCommand()
                .SelectCommand.CommandText = sSQL
                .SelectCommand.Connection = ClsDB.GetConnection()
                .Fill(_ds, FilterPageRowPosition, FilterPageSize, FilterTableName)
            End With
            '----------------------------------
        End If
        Return _ds.Tables(FilterTableName).DefaultView
    End Function
    Public Function AddCol(ByVal sName As String, _
                            ByVal sLabel As String, _
                            Optional ByVal sTitle As String = "", _
                            Optional ByVal bVisible As Boolean = True, _
                            Optional ByVal iSize As Integer = 10, _
                            Optional ByVal eTypeDB As ClsFilterCols.TypeDB = ClsFilterCols.TypeDB.STRING_T, _
                            Optional ByVal sTextValue As String = "", _
                            Optional ByVal sPageURLDestino As String = "", _
                            Optional ByVal sPageURLColVar As String = "", _
                            Optional ByVal Style As ClsFilterCols.TStyle = ClsFilterCols.TStyle.FilterField) As Boolean
        _FilterCols.Add(New ClsFilterCols(sName, sLabel, sTitle, bVisible, iSize, eTypeDB, sTextValue, sPageURLDestino, sPageURLColVar, Style).FilterColsReadOnly, sName)
        Return True
    End Function
    Public Function AddCol(ByVal pClsFCols As System.Data.PropertyCollection) As Boolean
        _FilterCols.Add(pClsFCols, pClsFCols("Name"))
        Return True
    End Function
    Private Function ShowFilter() As Object
        Dim Tabela As Table = New Table()
        Dim div As HtmlGenericControl = New HtmlGenericControl()
        Dim TableRow As TableRow
        Dim TableRowFiltros As TableRow
        Dim TableCell As TableCell
        Dim iCol As Integer = 0
        Dim iRow As Long = 0
        Dim sTxtName As String
        Dim sTxtSize As String
        Dim sTxtTitle As String
        Dim iColspan As Integer
        Dim sLink As String
        Dim sLinkSpace As String = Replace(Space(2), " ", "&nbsp;")
        Dim df As System.Data.DataView
        Dim rw As System.Data.DataRowView
        Dim sPageURLDestino As String
        Dim sPageURLColVar As String
        Dim ArrCol As ArrayList
        Dim sbControles As StringBuilder
        Dim tw As New IO.StringWriter
        Dim ht As New HtmlTextWriter(tw)
        Dim sFieldList As String = String.Empty
        '------------------------------------------
        If _FilterCols.Count > 0 And Not _FilterNotVisible Then
            '---------------------------------------------------
            sLink = "<a href=#>"
            If Not _FilterOptionClear Then
                '---------------------------------------------------     
                If FilterType = PFilterType.SimpleFilter Then
                    _GetSQLWhere = GetSQLWhere()
                End If
                '---------------------------------------------------     
                If _GetSQLWhere = "" Then
                    If FilterParamWhere <> "" Then
                        _GetSQLWhere = GetSQLWhere()
                    End If
                End If
            End If
            '---------------------------------------------------  
            df = GetFilterData(GetSQLSelect() & _GetSQLWhere & IIf(_GetSQLOrderBy <> String.Empty, _GetSQLOrderBy, IIf(FilterInitialOrderBy <> "", " ORDER BY " & FilterInitialOrderBy, "")))
            '---------------------------------------------------  
            iColspan = _FilterCols.Count
            '------------------------------------------
            With Tabela
                .ID = "filter_table"
                .CssClass = "tblFilter"
                .CellPadding = 0
                .CellSpacing = 0
            End With

            '------------------------------------------
            TableRow = New TableRow()
            TableCell = New TableCell()
            sbControles = New StringBuilder
            '------------------------------------------
            With sbControles
                .AppendLine("<div class=""divToolbar"">")
                .AppendLine("<div class=""tButtons"">")
                .AppendLine("<ul>")
                .AppendLine("<li><img src=""img/pagina1/pixToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li onclick=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "filter_click_novo") & """ id='filter_button_novo' name='filter_button_novo' ><img src=""img/pagina1/btn/btnNovo.gif"" alt=""Novo"" /></li>")
                .AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li onclick=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "filter_click_excel") & """ id='filter_button_excel' name='filter_button_excel' ><img src=""img/pagina1/btn/btnExcel.gif"" alt=""Excel"" /></li>")
                .AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li onclick=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "filter_click_aplica") & """ id='filter_button_aplicar' name='filter_button_aplicar'><img src=""img/pagina1/btn/btnFiltro.gif"" alt=""filtro""  /></li>")
                .AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li onclick=""javascript:window.print()"" title='Imprimir Filtro' id='filter_button_print' name='filter_button_print'><img src=""img/pagina1/btn/btnImprimir.gif"" alt=""Imprimir"" /></li>")
                .AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li onclick=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "filter_move_first$0$" & _PageQtde & "$" & _TotalRows) & """ id='filter_button_first' name='filter_button_first' ><img src=""img/pagina1/btn/btnPrimeiro.gif"" alt=""Primeiro"" /></li>")
                .AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li onclick=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "filter_move_previous$" & _PagePrevious & "$" & _PageQtde & "$" & _TotalRows) & """ id='filter_button_previous' name='filter_button_previous'><img src=""img/pagina1/btn/btnAnterior.gif"" alt=""Anterior"" /></li>")
                .AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li onclick=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "filter_move_next$" & _PageNext & "$" & _PageQtde & "$" & _TotalRows) & """ id='filter_button_next' name='filter_button_next'><img src=""img/pagina1/btn/btnProximo.gif"" alt=""Próximo"" /></li>")
                .AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li onclick=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "filter_move_last$" & _PageLast & "$" & _PageQtde & "$" & _TotalRows) & """ id='filter_button_last' name='filter_button_last'><img src=""img/pagina1/btn/btnUltimo.gif"" alt=""Ultimo""  /></li>")
                .AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li onclick=""javascript:window.location.href='" & FilterReturnFormName & "'"" ><img src=""img/pagina1/btn/volta.gif"" alt=""Voltar""  /></li>")
                .AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                .AppendLine("<li>")
                .AppendLine("Pagina " & _PageAtiva & " de " & _PageQtde & sLinkSpace & " &nbsp; - &nbsp;")
                .AppendLine("Qtde Registros (" & _TotalRows & ")")
                .AppendLine("</li>")
                .AppendLine("</ul>")
                .AppendLine("</div>")
                .AppendLine("</div>")
            End With            
            'Nome das Colunas        
            TableRow = New TableRow()
            For i = 1 To iColspan
                TableCell = New TableHeaderCell()
                TableCell.ID = "filter_col_" & _FilterCols(i)("Name").ToString()
                If CBool(_FilterCols(i)("Visible")) Then
                    If FilterOrderByCols And FilterPagePreLoad And CType(_FilterCols(i)("Style"), ClsFilterCols.TStyle) = ClsFilterCols.TStyle.FilterField Then
                        With Page.ClientScript
                            TableCell.Text = "<a href=""javascript:" & .GetPostBackEventReference(Me, "filter_orderby$" & _FilterCols(i)("Name").ToString() & "$" & _FilterColOrderByPos & "$" & _FilterLastOrderByColName) & """ title='Ordernar'>" & _FilterCols(i)("Label").ToString() & "</a>" & vbCrLf
                        End With
                    Else
                        TableCell.Text = _FilterCols(i)("Label").ToString()
                    End If
                    TableCell.ToolTip = _FilterCols(i)("Title").ToString()
                    With TableRow
                        .Cells.Add(TableCell)
                        .CssClass = "trHeader"
                    End With
                    Tabela.Rows.Add(TableRow)
                End If
            Next
            '---------------------------------------------------
            'Campos Para Filtro
            If FilterType = PFilterType.FullFilter Then
                TableRow = New TableRow()
                TableRowFiltros = New TableRow()
                For i = 1 To iColspan
                    TableCell = New TableCell()
                    If CBool(_FilterCols(i)("Visible")) Then
                        sTxtName = _FilterCols(i)("Name").ToString()
                        sTxtSize = _FilterCols(i)("Size").ToString()
                        sTxtTitle = _FilterCols(i)("Title").ToString()
                        If CType(_FilterCols(i)("Style"), ClsFilterCols.TStyle) = ClsFilterCols.TStyle.FilterField Then
                            sFieldList &= "'filter_txt_" & sTxtName & "', "
                            TableCell.Text = "<input title='" & sTxtTitle & "' type=text id='filter_txt_" & sTxtName & "' name='filter_txt_" & sTxtName & "' size='" & sTxtSize & "' value='" & _FilterCols(i)("TextValue").ToString() & "'"">" & vbCrLf
                        Else
                            With TableCell
                                .Style.Add("width", sTxtSize & "px")
                                .Text = "&nbsp;"
                            End With
                        End If
                        With TableRow
                            .CssClass = "trCamposFiltro"
                            .Cells.Add(TableCell)
                        End With
                    End If
                Next
            End If
            Tabela.Rows.Add(TableRow)
            '--------------------------------------------------
            If FilterType = PFilterType.FullFilter Then
                If FilterPagePreLoad Then
                    For Each rw In df
                        x += 1
                        TableRow = New TableRow()
                        For i = 1 To df.Table.Columns.Count
                            If CBool(_FilterCols(i)("Visible")) Then
                                sPageURLDestino = _FilterCols(i)("PageURLDestino")
                                sPageURLColVar = _FilterCols(i)("PageURLColVar")
                                If sPageURLDestino <> "" Then
                                    TableCell = New TableCell()
                                    If sPageURLColVar = "" Then
                                        sPageURLColVar = _FilterCols(i)("Name")
                                    End If
                                    '-----------------------------
                                    If FilterEncryptValue Then                                        
                                        TableRow.Attributes.Add("onclick", "window.location.href=""" & sPageURLDestino & "?" & sPageURLColVar & "=" & CorpCripto.EncryptString(rw(i - 1)) & GetQueryString(rw) & """")
                                    Else
                                        TableRow.Attributes.Add("onclick", "window.location.href=""" & sPageURLDestino & "?" & sPageURLColVar & "=" & rw(i - 1) & GetQueryString(rw) & """")
                                    End If
                                    TableCell.Text = NullDB(rw(i - 1))
                                    '-----------------------------
                                    TableRow.Cells.Add(TableCell)
                                Else
                                    TableCell = New TableCell()
                                    TableCell.Text = NullDB(rw(i - 1))
                                    TableRow.Cells.Add(TableCell)
                                End If
                            End If
                        Next
                        '-----------------------------
                        With TableRow
                            If x Mod 2 <> 0 Then
                                .Attributes.Add("onmouseover", "jsFilter.mouseOverOut(event);")
                                .CssClass = "darkTD"
                            Else
                                .Attributes.Add("onmouseover", "jsFilter.mouseOverOut(event);")
                                .CssClass = "lightTD"
                            End If
                        End With
                        '-----------------------------
                        Tabela.Rows.Add(TableRow)
                    Next
                End If
            Else
                If FilterPagePreLoad Then
                    For Each rw In df
                        ArrCol = New ArrayList
                        TableRow = New TableRow()
                        '------------------------------------
                        For i = 1 To df.Table.Columns.Count
                            If CBool(_FilterCols(i)("Visible")) Then
                                Select Case _FilterCols(i)("Name").ToString().ToLower()
                                    Case FilterColText.ToLower() : _FilterColText = rw(i - 1)
                                    Case FilterColValue.ToLower() : _FilterColValue = rw(i - 1)
                                End Select
                                ArrCol.Add(rw(i - 1))
                            End If
                        Next
                        '---------------------------------------------------          
                        TableCell = New TableCell()
                        TableCell.Text = "<a href=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "filter_cellclick$" & _FilterColText.Replace("""", "") & "$" & _FilterColValue) & """>" & ArrCol(0) & "</a>" & vbCrLf
                        TableRow.Cells.Add(TableCell)
                        '---------------------------------------------------          
                        For i = 1 To ArrCol.Count - 1
                            TableCell = New TableCell()
                            TableCell.Text = ArrCol(i)
                            TableRow.Cells.Add(TableCell)
                        Next
                        Tabela.Rows.Add(TableRow)
                    Next
                End If
                '---------------------------------------------------          
            End If
            'TableCell = New TableCell
            'TableRow = New TableRow
            ''------------------------------------------
            'If FilterType = PFilterType.FullFilter Then
            '    'Botões
            '    TableCell.Attributes.Add("colspan", iColspan)
            '    TableCell.Text = sbControles.ToString()
            'End If
            ''---------------------------------------------------
            ''Paginação |< < > >|  -  Paginas (9) 0...999  -  Qtde Registros ( 999 )                
            ''TableCell.Text &= sLinkSpace & " | " & sLinkSpace & " Paginas(<font id=""filter_pages"" color=red>" & _PageAtiva & "</font>) 1..." & _PageQtde & sLinkSpace & " | " & sLinkSpace & "Qtde Registros (<font id=""filter_records"" color=blue>" & _TotalRows & "</font>)" & sLinkSpace & vbCrLf
            'If FilterType = PFilterType.SimpleFilter Then
            '    TableCell.Attributes.Add("colspan", iColspan)
            'End If
            ''End With
            'TableRow.Cells.Add(TableCell)
            'Tabela.Rows.Add(TableRow)
            Tabela.RenderControl(ht)

            With Page.Response
                .Write(sbControles.ToString)
                .Write("<div style=""margin-left:5px;padding:4px"">")
                .Write(tw.ToString)
                .Write("</div>")
                .Write(sbControles.ToString)
            End With
            '---------------------------------------------------
        End If
        '---------------------------------------------------        
        Page.ClientScript.RegisterStartupScript(Page.GetType, "js_filter_init", _
                       "jsFilter.init([" & Mid(sFieldList, 1, sFieldList.Length - 2) & "]);", True)
        '---------------------------------------------------
        'Return Tabela

        div.Controls.Add(Tabela)
        Return div


    End Function
    '#####################################################################################    
    Private Sub GetClearSQLSelect()
        For i = 1 To _FilterCols.Count
            _FilterCols(i)("TextValue") = ""
        Next
        _FilterOptionClear = True
    End Sub
    '#####################################################################################
    'Monta o select SQL apartir do array de campos
    Private Function GetSQLSelect() As String
        Dim ClsSQL As New ClsSQL()
        With ClsSQL
            .sTable = FilterTableName
            For i = 1 To _FilterCols.Count
                .AddCol(_FilterCols(i)("Name").ToString())
            Next
        End With
        GetSQLSelect = ClsSQL.GetSELECT()
        Return GetSQLSelect
    End Function
    'Monta o Where apartir dos valor de campos da pagina
    '############################################################################################
    Private Function GetSQLWhere() As String
        Dim sStrText As String
        Dim sColName As String
        Dim sWhere As String = ""
        Dim iType As Integer
        Dim Request As System.Web.HttpRequest = System.Web.HttpContext.Current.Request
        With Request
            For i = 0 To .Form.Count - 1
                If Not IsNothing(.Form.Keys(i)) Then
                    sColName = .Form.Keys(i).ToString() 'Nome
                    If Mid(sColName, 1, 10).ToLower() = "filter_txt" Then
                        sStrText = .Form.Item(sColName).ToString() 'Valor
                        sColName = sColName.Substring(11) 'Nome Coluna DB
                        iType = _FilterCols(sColName)("Type") 'Tipo
                        _FilterCols(sColName)("TextValue") = sStrText
                        If sStrText.Trim() <> "" Then
                            sWhere &= sColName & " " & TrataSinal(sStrText, iType) & " AND "
                        End If
                    End If
                End If
            Next
        End With
        If sWhere <> "" Then
            If FilterParamWhere <> "" Then
                sWhere = " WHERE " & sWhere.Substring(0, sWhere.Length - 4) & " AND (" & FilterParamWhere & ")"
            Else
                sWhere = " WHERE " & sWhere.Substring(0, sWhere.Length - 4)
            End If
        Else
            If FilterParamWhere <> "" Then
                sWhere = " WHERE " & FilterParamWhere
            End If
        End If
        Return sWhere
    End Function
    'Trata cada campo retornando o valor e operador da cláusula Where
    '############################################################################################
    Private Function TrataSinal(ByVal sItem As String, ByVal iTipoDado As ClsFilterCols.TypeDB) As String
        Dim i As Integer
        Dim iSinal As Integer
        Dim sSinal As Object
        Dim bSinal As Boolean
        Dim sReturn As String = ""
        '-------------------------------------------------
        Dim sMensagem As String = ClsTools.CheckCommandSQL(sItem)
        If sMensagem <> String.Empty Then Throw New Exception(sMensagem)
        '-------------------------------------------------
        iSinal = -1
        sSinal = Split("%x<>x>=x<=x<x>", "x")
        sItem = sItem.Replace("'", "''")
        For i = 0 To UBound(sSinal)
            If InStr(sItem, sSinal(i)) <> 0 Then
                iSinal = i
                bSinal = True
                Exit For
            End If
        Next
        If bSinal = True Then
            Select Case iSinal
                Case 0
                    sReturn = " LIKE '" & sItem & "'"
                Case Else
                    sReturn = sItem
            End Select
        Else
            Select Case iTipoDado
                Case ClsFilterCols.TypeDB.NUMERIC_T, ClsFilterCols.TypeDB.MONEY_T
                    sReturn = " = " & sItem
                Case ClsFilterCols.TypeDB.STRING_T
                    If UCase(sItem) = "IS NOT NULL" Then
                        sReturn = " IS NOT NULL"
                    ElseIf UCase(sItem) = "IS NULL" Then
                        sReturn = " IS NULL"
                    ElseIf IsNumeric(UCase(sItem)) Then
                        sReturn = " = '" & sItem & "'"
                    ElseIf UCase(sItem).Contains(">") And IsNumeric(UCase(sItem).Replace(">", "")) Then
                        sReturn = " > '" & sItem & "'"
                    ElseIf UCase(sItem).Contains("<") And IsNumeric(UCase(sItem).Replace("<", "")) Then
                        sReturn = " < '" & sItem & "'"
                    ElseIf IsDate(UCase(sItem)) Then
                        sReturn = " between '" & Format(CDate(sItem), "yyyy-MM-dd") & " 00:00' and '" & Format(CDate(sItem), "yyyy-MM-dd") & " 23:59'"
                    ElseIf UCase(sItem).Contains("*") Then
                        sReturn = " LIKE '%" & sItem.Replace("*", "") & "%'"
                    Else
                        sReturn = " LIKE '%" & sItem & "%'"
                    End If
                Case ClsFilterCols.TypeDB.DATE_T
                    sReturn = " = '" & Format(CDate(sItem), "yyyy-MM-dd") & "'"
                Case ClsFilterCols.TypeDB.DATE_TIME_T
                    sReturn = " = '" & Format(CDate(sItem), "yyyy-MM-dd H:mm:ss") & "'"
            End Select
        End If
        Return sReturn
    End Function
    Public Sub ExportarXLS()
        '------------------------------------------------------------
        Dim TableDataSorce As New System.Data.DataTable
        Dim oDB As New ClsDB(FilterStrConnection, FilterProvider)
        '------------------------------------------------------------
        With oDB.GetDataAdapter()
            .SelectCommand.CommandText = GetSQLSelect() & " " & _GetSQLWhere
            .Fill(TableDataSorce)
        End With
        '------------------------------------------------------------
        Dim excel As New ClsGetExcelFile
        Dim Response As HttpResponse = HttpContext.Current.Response
        '--------------------------------
        If TableDataSorce.Rows.Count = 0 Then
            Page.Controls.Clear()
            Response.Clear()
            Me.Controls.Add(New LiteralControl("<div id=""rptMsg"" style=""color: red;"">Nenhuma informação foi gerada.</div>"))
        Else
            With excel
                .DataSource = TableDataSorce
                'Colunas Detalhe                
                For Each col As PropertyCollection In _FilterCols
                    If CType(col("Style"), ClsFilterCols.TStyle) = ClsFilterCols.TStyle.FilterField Then
                        .AddColumnTitle(col("Name"), col("Label"))
                    End If
                Next
                '-----------------------------------
                'Gera a planilha
                .GenerateXLS()
            End With
            '-----------------------------------                
            'Enviando informações
            Dim aBytes() As Byte = CType(excel.GetStream, IO.MemoryStream).ToArray
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

    Public Sub AddQueryString(ByVal sName As String, Optional ByVal sValue As String = "")
        If sName.Trim <> "" Then
            Dim row As DataRow
            If sName.Contains(" ") Then
                Throw New Exception("Nome da Chave inválida!")
            Else
                With _FilterQueryString
                    row = .NewRow()
                    row("key") = sName.Trim
                    row("value") = sValue.Replace("  ", " ").Trim
                    .Rows.Add(row)
                End With
            End If
        Else
            Throw New Exception("Informe Nome da variável!")
        End If
    End Sub

    Private Function GetQueryString(Optional ByVal rowView As DataRowView = Nothing) As String
        Dim sQString As New StringBuilder
        Dim row As DataRow
        With _FilterQueryString
            For i As Integer = 0 To .Rows.Count - 1
                row = .Rows(i)
                With sQString
                    If CheckFieldExists(row("key"), rowView) Then
                        .Append("&").Append(row("key")).Append("=").Append(rowView(row("key")).ToString.Replace(" ", "+"))
                    Else
                        .Append("&").Append(row("key")).Append("=").Append(row("value").ToString.Replace(" ", "+"))
                    End If
                End With
            Next
        End With
        Return sQString.ToString()
    End Function

    Private Function CheckFieldExists(ByVal strName As String, ByVal Row As DataRowView) As Boolean
        Try
            CheckFieldExists = True
            Dim strGet As String = Row(strName)
        Catch ex As Exception
            CheckFieldExists = False
        End Try
        Return CheckFieldExists
    End Function

#End Region
    '#######################################################################  
#Region "Filter Method Hidden"
    Private Function MyPath() As String
        Dim strURI As String = HttpContext.Current.Request.Url.AbsoluteUri
        strURI = strURI.Replace(HttpContext.Current.Request.Url.Query, "")
        Dim strFile As String = HttpContext.Current.Request.AppRelativeCurrentExecutionFilePath.Replace("~/", "")
        Return strURI.Replace(strFile, "") & "js/"
    End Function
    Private Function SetViewStateFilter(ByVal sKey As String, ByVal sValue As String) As Boolean
        Dim sRetVld As Object = Nothing
        If _FilterState.Count > 0 Then
            sRetVld = _FilterState(sKey)
        End If
        If Not sRetVld Is Nothing Then
            _FilterState(sKey) = sValue
        Else
            _FilterState.Add(sKey, sValue)
        End If
    End Function
    Private Function GetViewStateFilter(ByVal sKey As String) As String
        Return _FilterState(sKey)
    End Function
    Protected Overrides Sub OnInit(ByVal e As System.EventArgs)
        SetViewStateFilter("ColOrderBy", "")
        With Page
            With .ClientScript
                .RegisterClientScriptInclude(Me.Page.GetType, "js_forms", FilterJsUrl & "forms.js")
                .RegisterClientScriptInclude(Me.Page.GetType, "js_filter", FilterJsUrl & "filter.js")
            End With
            .RegisterRequiresControlState(Me)
        End With
        MyBase.OnInit(e)
    End Sub
    Protected Overrides Function SaveControlState() As Object
        _FilterViewState = New Collection()
        _FilterViewState.Add(_FilterCols, "FilterCols")
        _FilterViewState.Add(_FilterState, "FilterState")
        _FilterViewState.Add(_FilterQueryString, "FilterQueryString")
        Return Me._FilterViewState
    End Function
    Protected Overrides Sub LoadControlState(ByVal savedState As Object)
        _FilterViewState = New Collection()
        _FilterViewState = CType(savedState, Collection)
        _FilterCols = _FilterViewState("FilterCols")
        _FilterState = _FilterViewState("FilterState")
        _FilterQueryString = _FilterViewState("FilterQueryString")
    End Sub
#End Region

End Class
