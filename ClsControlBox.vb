Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls


<DefaultProperty(""), ToolboxData("<{0}:ClsControlBox runat=server></{0}:ClsControlBox>")> _
Public Class ClsControlBox : Inherits WebControl : Implements IPostBackEventHandler

    Public Sub New()
        BuildTables()
    End Sub


#Region "Fields"

    Private dts As New DataSet
    Private cnn As New OleDb.OleDbConnection
    Private adp As OleDb.OleDbDataAdapter
    Private strText As String = String.Empty

#End Region


#Region "Properties"

    ''' <summary>
    ''' Permite configurar o controle para múltiplas seleções
    ''' ou apenas uma linha selecionada. (CheckBox ou Radio respectivamente)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Bindable(True), Description("Se permite seleções múltiplas ou de apenas um item."), Category("Behavior"), DefaultValue(""), Localizable(True)> _
        Public Property MultiSelect() As Boolean
        Get
            Dim o As Object = ViewState("_MultiSelect") = True
            Return ViewState("_MultiSelect")
        End Get
        Set(ByVal value As Boolean)
            ViewState("_MultiSelect") = value
        End Set
    End Property

    ''' <summary>
    ''' Se as colunas serão geradas a partir do comando SQL/Stored Procedure
    ''' passado em CommantText.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Bindable(True), Description("Se as colunas serão geradas automaticamente."), Category("Data"), DefaultValue(""), Localizable(True)> _
    Public Property AutoGenerateColumns() As Boolean
        Get
            Dim s As Object = ViewState("_AutoGenerateColumns")
            If s Is Nothing Then ViewState("_AutoGenerateColumns") = True
            Return ViewState("_AutoGenerateColumns")
        End Get
        Set(ByVal value As Boolean)
            ViewState("_AutoGenerateColumns") = value
        End Set
    End Property


    ''' <summary>
    ''' Estilo do controle. (TEXTAREA é o valor default).
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Permite exibir caixas de texto e listas com mais de 
    ''' uma coluna, permitindo rotulá-las.</remarks>
    <Bindable(True), Description("Estilo da Caixa. Se TextArea ou ListBox."), Category("Appearance"), DefaultValue(""), Localizable(True)> _
    Public Property BoxStyle() As TBoxStyle
        Get
            Dim s As Object = ViewState("_Style")
            If s Is Nothing Then ViewState("_Style") = TBoxStyle.TextArea
            Return ViewState("_Style")
        End Get
        Set(ByVal Value As TBoxStyle)
            Select Case Value
                Case TBoxStyle.ListBox, TBoxStyle.TextArea
                    ViewState("_Style") = Value
                Case Else
                    Throw New Exception("Valor atribuído é inválido!")
            End Select
        End Set
    End Property

    ''' <summary>
    ''' Permite definir/recuperar o título da caixa.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Bindable(True), Description("Título da caixa do controle."), Category("Appearance"), DefaultValue(""), Localizable(True)> _
    Public Property BoxTitle() As String
        Get
            Dim s As Object = ViewState("_Title")
            If s Is Nothing Then ViewState("_Title") = String.Empty
            Return ViewState("_Title")
        End Get
        Set(ByVal Value As String)
            ViewState("_Title") = Value
        End Set
    End Property

    ''' <summary>
    ''' Retorna a fonte de dados para manipulação externa.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    <Bindable(False)> _
    Public ReadOnly Property DataSource() As DataSet
        Get
            Return dts
        End Get
    End Property

    ''' <summary>
    ''' Retorna/Altera o valor em uma coluna no DataSource.
    ''' </summary>
    ''' <value></value>
    ''' <param name="ColumnName">Nome da Coluna a ser afetada.</param>
    ''' <param name="RowIndex">Índice da Linha a ser afetada.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnValue(ByVal ColumnName As String, ByVal RowIndex As Long) As String
        Get
            If dts.Tables("dtaLista").Columns.Contains(ColumnName) Then
                Return dts.Tables("dtaLista").Rows(RowIndex)(ColumnName).ToString
            Else
                Throw New Exception("Coluna não encontrada no DataSource!")
            End If
        End Get
        Set(ByVal Value As String)
            If dts.Tables("dtaLista").Columns.Contains(ColumnName) Then
                dts.Tables("dtaLista").Rows(RowIndex)(ColumnName) = Value
            Else
                Throw New Exception("Coluna não encontrada no DataSource!")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Retorna/Altera o valor em uma coluna no DataSource.
    ''' </summary>
    ''' <value></value>
    ''' <param name="ColumnIndex ">Índice da Coluna a ser afetada.</param>
    ''' <param name="RowIndex">Índice da Linha a ser afetada.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnValue(ByVal ColumnIndex As Integer, ByVal RowIndex As Long) As String
        Get
            Return dts.Tables("dtaLista").Rows(RowIndex)(ColumnIndex).ToString
        End Get
        Set(ByVal Value As String)
            dts.Tables("dtaLista").Rows(RowIndex)(ColumnIndex) = Value
        End Set
    End Property

    ''' <summary>
    ''' Permite postar o formulário quando um ítem for selecionado. 
    ''' O default é False.
    ''' </summary>
    ''' <value>True/False</value>
    ''' <returns>True/False</returns>
    ''' <remarks></remarks>
    <Bindable(True), Description("Automaticamente vai ao servidor ao clicar de um item."), Category("Behavior"), DefaultValue(""), Localizable(True)> _
    Public Property AutoPostBack() As Boolean
        Get
            Dim s As Object = ViewState("_AutoPostBack")
            If s Is Nothing Then ViewState("_AutoPostBack") = False
            Return ViewState("_AutoPostBack")
        End Get
        Set(ByVal value As Boolean)
            ViewState("_AutoPostBack") = value
        End Set
    End Property

    ''' <summary>
    ''' Conteúdo texto do controle estilo TEXTAREA.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Bindable(True), Description("O valor do texto do objeto."), Category("Appearance"), DefaultValue(""), Localizable(True)> _
    Public Property Text() As String
        Get
            Dim s As Object = ViewState("_Text")
            If s Is Nothing Then ViewState("_Text") = String.Empty
            Return ViewState("_Text")
        End Get
        Set(ByVal value As String)
            ViewState("_Text") = value
        End Set
    End Property

    ''' <summary>
    ''' Retorna se a linha está selecionada.
    ''' </summary>
    ''' <param name="rowIndex">Índice da linha desejada.</param>
    ''' <value></value>
    ''' <returns>True/False</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsSelectedRow(ByVal rowIndex As Long) As Boolean
        Get
            Return dts.Tables("dtaLista").Rows(rowIndex)("Selected")
        End Get
    End Property

    ''' <summary>
    ''' Retorna as linhas selecionadas.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SelectedRows() As DataRow()
        Get
            Return dts.Tables("dtaLista").Select("Selected")
        End Get
    End Property

    ''' <summary>
    ''' Retorna o número de colunas da fonte de dados
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ColumnCount() As Long
        Get
            Return dts.Tables("dtaLista").Columns.Count
        End Get
    End Property

    ''' <summary>
    ''' Retorna o número de linhas da fonte de dados.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    <Bindable(False)> _
    Public ReadOnly Property RowCount() As Long
        Get
            Return dts.Tables("dtaLista").Rows.Count
        End Get
    End Property

    ''' <summary>
    ''' Retorna uma Columa da fonte de dados
    ''' </summary>
    ''' <param name="Index">Índice da Coluna.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Column(ByVal Index As Integer) As DataColumn
        Get
            Return dts.Tables("dtaLista").Columns(Index)
        End Get
    End Property

    ''' <summary>
    ''' Permite alterar a propriedade das colunas.
    ''' </summary>
    ''' <param name="Index">Índice da coluna.</param>
    ''' <param name="PropertyName">Propriedade a ser alterada.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnProperty(ByVal Index As Integer, ByVal PropertyName As String) As String
        Get
            '------------------------------------------------------
            If Not dts.Tables("dtaColumnProperty").Columns.Contains(PropertyName) Then Throw New Exception("Propriedade inválida!")
            '------------------------------------------------------
            Dim dtCol As DataColumn = dts.Tables("dtaLista").Columns(Index)
            Dim dtRow() As DataRow = dts.Tables("dtaColumnProperty").Select("ColumnName = '" & dtCol.ColumnName & "'")
            '------------------------------------------------------
            If dtRow.Length = 0 Then
                Throw New Exception("Coluna """ & dtCol.ColumnName & """ não foi encontrada.")
            Else
                Return dtRow(0).Item(PropertyName)
            End If
        End Get
        Set(ByVal value As String)
            '------------------------------------------------------
            If Not dts.Tables("dtaColumnProperty").Columns.Contains(PropertyName) Then Throw New Exception("Propriedade inválida!")
            '------------------------------------------------------
            Dim dtCol As DataColumn = dts.Tables("dtaLista").Columns(Index)
            Dim dtRow() As DataRow = dts.Tables("dtaColumnProperty").Select("ColumnName = '" & dtCol.ColumnName & "'")
            '------------------------------------------------------
            If dtRow.Length = 0 Then
                Throw New Exception("Coluna """ & dtCol.ColumnName & """ não foi encontrada.")
            Else
                dtRow(0).Item(PropertyName) = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Permite alterar a propriedade das colunas.
    ''' </summary>
    ''' <param name="ColumnName">Nome da coluna.</param>
    ''' <param name="PropertyName">Propriedade a ser alterada.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnProperty(ByVal ColumnName As String, ByVal PropertyName As String) As String
        Get
            '------------------------------------------------------
            If Not dts.Tables("dtaColumnProperty").Columns.Contains(PropertyName) Then Throw New Exception("Propriedade inválida!")
            '------------------------------------------------------            
            Dim dtRow() As DataRow = dts.Tables("dtaColumnProperty").Select("ColumnName = '" & ColumnName & "'")
            '------------------------------------------------------
            If dtRow.Length = 0 Then
                Throw New Exception("Coluna """ & ColumnName & """ não foi encontrada.")
            Else
                Return dtRow(0).Item(PropertyName)
            End If
        End Get
        Set(ByVal value As String)
            '------------------------------------------------------
            If Not dts.Tables("dtaColumnProperty").Columns.Contains(PropertyName) Then Throw New Exception("Propriedade inválida!")
            '------------------------------------------------------            
            Dim dtRow() As DataRow = dts.Tables("dtaColumnProperty").Select("ColumnName = '" & ColumnName & "'")
            '------------------------------------------------------
            If dtRow.Length = 0 Then
                Throw New Exception("Coluna """ & ColumnName & """ não foi encontrada.")
            Else
                dtRow(0).Item(PropertyName) = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Retorna uma Columa da fonte de dados
    ''' </summary>
    ''' <param name="Name">Nome da Coluna.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Column(ByVal Name As String) As DataColumn
        Get
            Return dts.Tables("dtaLista").Columns(Name)
        End Get
    End Property

    ''' <summary>
    ''' Retorna uma linha determinada por "RowIndex".
    ''' </summary>
    ''' <param name="RowIndex">Índice da linha escolhida.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Row(ByVal RowIndex As Long) As DataRow
        Get
            Return dts.Tables("dtaLista").Rows(RowIndex)
        End Get
    End Property

    ''' <summary>
    ''' Comando SQL/Stored Procedure para preenchimento da tabela.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Bindable(True), Description("Comando SQL/Stored Procedure para preenchimento da tabela."), Category("Data"), DefaultValue(""), Localizable(True)> _
    Public Property CommandText() As String
        Get
            Dim s As Object = ViewState("_CommandText")
            If s Is Nothing Then ViewState("_CommandText") = String.Empty
            Return ViewState("_CommandText")
        End Get
        Set(ByVal value As String)
            ViewState("_CommandText") = value
        End Set
    End Property

    ''' <summary>
    ''' Permite informar qual objeto de conexão deverá ser usado
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    '''     
    Public WriteOnly Property ActiveConnection() As OleDb.OleDbConnection
        Set(ByVal value As OleDb.OleDbConnection)
            cnn = value
        End Set
    End Property

    ''' <summary>
    ''' Permite definir a string de conexão.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Bindable(True), Description("String de conexão da fonte de dados."), Category("Data"), DefaultValue(""), Localizable(True)> _
    Public Property ConnectionString() As String
        Get
            Dim s As Object = ViewState("_ConnectionString")
            If s Is Nothing Then ViewState("_ConnectionString") = String.Empty
            Return ViewState("_ConnectionString")
        End Get
        Set(ByVal value As String)
            ViewState("_ConnectionString") = value
        End Set
    End Property

#End Region


#Region "Enumerations"

    Public Enum TBoxStyle
        TextArea
        ListBox
    End Enum

#End Region


#Region "Methods"

    ''' <summary>
    ''' Importa as colunas de um outro controle de mesma Classe Base.
    ''' </summary>
    ''' <param name="DataSource"></param>
    ''' <remarks></remarks>
    Public Sub ImportColumns(ByRef DataSource As DataSet)
        Dim dtPropEx As DataTable = DataSource.Tables("dtaColumnProperty")
        Dim dtListaEx As DataTable = DataSource.Tables("dtaLista")
        Dim dtListaIn As DataTable = dts.Tables("dtaLista")
        Dim dtRowProp() As DataRow

        With dtListaIn
            .Columns.Clear()
            For Each dtCol As DataColumn In dtListaEx.Columns
                If Not .Columns.Contains(dtCol.ColumnName) Then
                    '------------------------------------------------------------
                    If dtPropEx IsNot Nothing Then
                        dtRowProp = dtPropEx.Select("ColumnName = '" & dtCol.ColumnName & "'")
                        If dtRowProp.Length > 0 Then CloneRowProp(dtRowProp(0))
                    End If
                    '------------------------------------------------------------
                    .Columns.Add(dtCol.ColumnName)
                    '------------------------------------------------------------
                End If
            Next
        End With
    End Sub

    ''' <summary>
    ''' Adiciona colunas para controle estilo LISTBOX.
    ''' </summary>
    ''' <param name="ColumnName">Nome da coluna.</param>
    ''' <param name="ColumnTitle"></param>
    ''' <remarks></remarks>
    Public Sub AddColumn(ByVal ColumnName As String, _
                        ByVal ColumnTitle As String, _
                        Optional ByVal Visible As Boolean = True, _
                        Optional ByVal Width As String = "")
        '------------------------------------------------
        dts.Tables("dtaLista").Columns.Add(ColumnName)
        '------------------------------------------------
        AddProperty(ColumnName, ColumnTitle, Visible, Width)
        '------------------------------------------------
    End Sub

    Private Sub AddProperty(ByVal ColumnName As String, _
                        ByVal ColumnTitle As String, _
                        Optional ByVal Visible As Boolean = True, _
                        Optional ByVal Width As String = "")
        '------------------------------------------------
        Dim dtrow As DataRow
        Dim dtaProp As DataTable = dts.Tables("dtaColumnProperty")
        '------------------------------------------------
        If dtaProp.Select("ColumnName = '" & ColumnName & "'").Length = 0 Then
            dtrow = dtaProp.NewRow
            With dtrow
                .Item("ColumnName") = ColumnName
                .Item("ColumnTitle") = ColumnTitle
                .Item("Visible") = Visible
                .Item("Width") = Width
            End With
            '------------------------------------------------
            dtaProp.Rows.Add(dtrow)
        End If
    End Sub

    ''' <summary>
    ''' Retorna uma nova linha com as características da
    ''' fonte de dados do controle.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NewRow() As DataRow
        Return dts.Tables("dtaLista").NewRow
    End Function

    ''' <summary>
    ''' Adiciona uma nova linha à fonte de dados.
    ''' </summary>
    ''' <param name="dtRow">DataRow a ser inserido.</param>
    ''' <remarks></remarks>
    Public Sub CloneRow(ByVal dtRow As DataRow)
        Dim dtr As DataRow
        '--------------------------
        dtRow("Selected") = False
        '--------------------------
        With dts.Tables("dtaLista")
            dtr = .NewRow
            For Each dtc As DataColumn In dtRow.Table.Columns
                If Not dts.Tables("dtaLista").Columns.Contains(dtc.ColumnName) Then _
                    AddColumn(dtc.ColumnName, dtc.ColumnName)
                dtr(dtc.ColumnName) = dtRow(dtc.ColumnName)
            Next
            .Rows.Add(dtr)
        End With
    End Sub

    ''' <summary>
    ''' Adiciona nova linha à fonte de dados, tentando transferí-la de sua origem.
    ''' </summary>
    ''' <param name="dtRow">DataRow a ser inserido.</param>
    ''' <remarks></remarks>
    Public Sub ImportRow(ByVal dtRow As DataRow)
        Dim dtr As DataRow
        '--------------------------
        dtRow("Selected") = False
        '--------------------------
        With dts.Tables("dtaLista")
            dtr = .NewRow
            For Each dtc As DataColumn In dtRow.Table.Columns
                If Not dts.Tables("dtaLista").Columns.Contains(dtc.ColumnName) Then _
                    AddColumn(dtc.ColumnName, dtc.ColumnName)
                dtr(dtc.ColumnName) = dtRow(dtc.ColumnName)
            Next
            .Rows.Add(dtr)
            Try
                dtRow.Table.Rows.Remove(dtRow)
            Catch : End Try
        End With
    End Sub

    Private Sub CloneRowProp(ByRef dtRow As DataRow)
        Dim dtr As DataRow
        '--------------------------
        With dts.Tables("dtaColumnProperty")
            dtr = .NewRow
            For Each dtc As DataColumn In dtRow.Table.Columns
                dtr(dtc.ColumnName) = dtRow(dtc.ColumnName)
            Next
            .Rows.Add(dtr)
        End With
    End Sub

    ''' <summary>
    ''' Remove uma linha especificada por "Index".
    ''' </summary>
    ''' <param name="Index">Índice da linha a ser excluída (base 0).</param>
    ''' <remarks></remarks>
    Public Sub RemoveRow(ByVal Index As Long)
        dts.Tables("dtaLista").Rows.RemoveAt(Index)
    End Sub

    ''' <summary>
    ''' Remove uma linha especificada.
    ''' </summary>
    ''' <param name="Row">Linha a ser excluída.</param>
    ''' <remarks></remarks>
    Public Sub RemoveRow(ByVal Row As DataRow)
        dts.Tables("dtaLista").Rows.Remove(Row)
    End Sub

    ''' <summary>
    ''' Seleciona todas as linhas do ListBox.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SelectAllRows(Optional ByVal Selected As Boolean = True)
        For Each dtr As DataRow In dts.Tables("dtaLista").Rows
            dtr.BeginEdit()
            dtr("Selected") = Selected
            dtr.EndEdit()
        Next
    End Sub

    Public Sub ShowDesignMode()
        Dim dtRow As DataRow
        '----------------------------------------
        With dts.Tables("dtaLista")

            '------------------------
            .Rows.Clear()
            .Columns.Clear()
            '------------------------

            Dim c As New DataColumn
            c.DefaultValue = True
            c.ReadOnly = False
            c.ColumnName = "Selected"
            .Columns.Add(c)

            .Columns.Add("COLUMN1")

            dtRow = dts.Tables("dtaColumnProperty").NewRow
            dtRow("ColumnName") = "COLUMN1"
            dtRow("ColumnTitle") = "Col 1"
            dtRow("Visible") = True
            dtRow("Width") = "50%"
            dts.Tables("dtaColumnProperty").Rows.Add(dtRow)

            .Columns.Add("COLUMN2")

            dtRow = dts.Tables("dtaColumnProperty").NewRow
            dtRow("ColumnName") = "COLUMN2"
            dtRow("ColumnTitle") = "Col 2"
            dtRow("Visible") = True
            dtRow("Width") = "50%"
            dts.Tables("dtaColumnProperty").Rows.Add(dtRow)

            '-----------------------------------
            dtRow = .NewRow
            dtRow("COLUMN1") = "Column 1 Row 1"
            dtRow("COLUMN2") = "Column 2 Row 1"
            .Rows.Add(dtRow)
            '-----------------------------------
            dtRow = .NewRow
            dtRow("COLUMN1") = "Column 1 Row 2"
            dtRow("COLUMN2") = "Column 2 Row 2"
            .Rows.Add(dtRow)
            '-----------------------------------
        End With
    End Sub

    Private Sub RenderCustomControl()
        If BoxStyle = TBoxStyle.ListBox Then
            BuildListBox()
        Else
            BuildTextBox()
        End If
    End Sub

    Private Sub BuildListBox()
        Dim tabCaixaLista As New Table
        Dim tabCaixaListaDetalhe As New Table
        Dim tabCaixaListaDetalheLinhas As New Table
        Dim pnlCaixaListaDetalheLinhas As New Panel

        Dim dtRowProp() As DataRow
        Dim dtRow As DataRow
        Dim hshColsWidth As New Hashtable
        Dim hshColsVisible As New Hashtable
        Dim objSelecao As New Object

        Dim row As TableRow
        Dim cel As TableCell

        Dim blnTemColunas As Boolean = False

        '------------------------------------------------------
        With tabCaixaLista
            .ID = "tabCaixaLista"
            .CellPadding = 0
            .CellSpacing = 0
            .Width = Me.Width
        End With

        With tabCaixaListaDetalhe
            .ID = "tabCaixaListaDetalhe"
            .CellPadding = 0
            .CellSpacing = 0
            .Width = Unit.Percentage(100)
            .Attributes.Add("class", "rowNormal")
        End With

        With tabCaixaListaDetalheLinhas
            .ID = "tabCaixaListaDetalheLinhas"
            .Width = Unit.Percentage(94)
            .CellSpacing = 0
            .CellPadding = 0
        End With
        '---------------------------------------------
        'Titulo da caixa
        If BoxTitle.Trim <> String.Empty Then
            row = New TableRow
            cel = New TableHeaderCell
            With cel
                .ID = "tituloLst"
                .Text = BoxTitle.Trim.Replace(" ", "&nbsp;")
            End With
            row.Cells.Add(cel)
            tabCaixaLista.Rows.Add(row)
        End If
        '---------------------------------------------
        'Montando cabeçalho das colunas
        row = New TableRow
        cel = New TableHeaderCell
        With cel
            .Width = Unit.Pixel(1)
            .ID = "titulocol"
            .Text = "&nbsp;"
        End With
        row.Cells.Add(cel)
        For Each dtcol As DataColumn In dts.Tables("dtaLista").Columns
            '---------------------------------------------
            dtRowProp = dts.Tables("dtaColumnProperty").Select("ColumnName = '" & dtcol.ColumnName & "'")
            '---------------------------------------------
            If CType(dtRowProp(0)("Visible"), Boolean) Then
                '------------------------------------
                blnTemColunas = True
                cel = New TableHeaderCell
                With cel
                    .ID = "titulocol"
                    .Text = dtRowProp(0)("ColumnTitle").ToString()
                    .Style.Add("width", IIf(dtRowProp(0)("Width") = String.Empty, "auto", dtRowProp(0)("Width")))
                    '-------------------------------------------------
                    'HashTable com informações de Largura para uso posterior 
                    hshColsWidth.Add(dtcol.ColumnName, dtRowProp(0)("Width"))
                    hshColsVisible.Add(dtcol.ColumnName, dtRowProp(0)("Visible"))
                    '-------------------------------------------------
                End With
                row.Cells.Add(cel)
                '------------------------------------
            End If
            '---------------------------------------------
        Next
        tabCaixaListaDetalheLinhas.Rows.Add(row)
        '---------------------------------------------
        'Montando o detalhe da caixa
        For lngIndex As Long = 0 To dts.Tables("dtaLista").Rows.Count - 1
            '-------------------------------------------------
            'Carregando a linha atual
            dtRow = dts.Tables("dtaLista").Rows(lngIndex)
            '-------------------------------------------------
            'Configura os atributos da tag "TR"
            row = New TableRow
            '-------------------------------------------------
            With row
                .ID = "row_" & Me.ClientID & "_" & lngIndex.ToString
                .Attributes.Add("class", IIf(dts.Tables("dtaLista").Rows(lngIndex)("Selected"), "rowSelected", "rowNormal"))
            End With
            '-------------------------------------------------
            'Preenchendo as colunas da linha atual        
            objSelecao = RetCheckBox(lngIndex)
            cel = New TableCell
            cel.Width = Unit.Pixel(1)
            cel.Controls.Add(objSelecao)
            row.Cells.Add(cel)
            For Each dtcol As DataColumn In dts.Tables("dtaLista").Columns
                '---------------------------------------------
                If CType(hshColsVisible(dtcol.ColumnName), Boolean) Then
                    cel = New TableCell
                    With cel
                        If hshColsWidth(dtcol.ColumnName).ToString.Trim <> String.Empty Then
                            .Style.Add("width", hshColsWidth(dtcol.ColumnName))
                        End If
                        .Text = "<label for='" & objSelecao.id & "'>" & IIf(IsDBNull(dtRow(dtcol.ColumnName)), "&nbsp;", dtRow(dtcol.ColumnName)) & "</label>"
                    End With
                    row.Cells.Add(cel)
                End If
                '---------------------------------------------
            Next
            tabCaixaListaDetalheLinhas.Rows.Add(row)
            '-------------------------------------------------
        Next
        '-------------------------------------------------
        With pnlCaixaListaDetalheLinhas
            .ID = "dvCaixaListaDetalheLinhas"
            With .Style
                .Add("overflow", "scroll")
                '.Add("position", "relative")
                .Add("height", (CInt(Me.Height.Value * 0.8)).ToString & "px")
                .Add("width", (CInt(Me.Width.Value * 0.98)).ToString & "px")
                '.Add("height", Me.Height.ToString)
                '.Add("width", Me.Width.ToString)
            End With
            .Attributes.Add("valign", "top")
            .Controls.Add(tabCaixaListaDetalheLinhas)
        End With
        row = New TableRow
        cel = New TableCell
        cel.ColumnSpan = dts.Tables("dtaLista").Columns.Count + 1
        cel.VerticalAlign = VerticalAlign.Top
        cel.Controls.Add(pnlCaixaListaDetalheLinhas)
        row.Cells.Add(cel)
        tabCaixaListaDetalhe.Rows.Add(row)
        '------------------------------------------------------------------------
        row = New TableRow
        cel = New TableCell
        cel.Controls.Add(tabCaixaListaDetalhe)
        row.Cells.Add(cel)
        tabCaixaLista.Rows.Add(row)
        '------------------------------------------------------------------------
        Me.Controls.Add(tabCaixaLista)
        Call SaveControlState()
        '------------------------------------------------------------------------
    End Sub

    Private Sub BuildTextBox()
        Dim tabCaixaLista As New Table
        Dim row As TableRow
        Dim cel As TableCell
        Dim txt As New TextBox
        '-----------------------------------
        With tabCaixaLista
            .ID = "tabCaixaLista"
            .CellPadding = 0
            .CellSpacing = 0
            .Width = Me.Width
            .Height = Me.Height
        End With
        '-----------------------------------
        With txt
            .ID = "txtTextBox_" & Me.ClientID
            .Width = Unit.Percentage(100)
            .Height = Me.Height
            .TextMode = TextBoxMode.MultiLine
            txt.Text = strText
        End With
        '-----------------------------------
        If BoxTitle.Trim <> String.Empty Then
            row = New TableRow
            cel = New TableHeaderCell
            With cel
                .ID = "tituloLst"
                .Text = BoxTitle
            End With
            row.Cells.Add(cel)
            tabCaixaLista.Rows.Add(row)
        End If
        '-----------------------------------
        row = New TableRow
        cel = New TableHeaderCell
        cel.Style.Add("height", Me.Height.ToString)
        cel.Controls.Add(txt)
        row.Cells.Add(cel)
        tabCaixaLista.Rows.Add(row)
        '-----------------------------------
        Me.Controls.Add(tabCaixaLista)
        '-----------------------------------
    End Sub

    Private Function RetCheckBox(ByVal lngRow As Long) As Object
        Dim oSelect As Object
        '-----------------------------------
        If MultiSelect Then
            Dim chkSelect As New CheckBox
            chkSelect.ID = "sel_" & Me.ClientID & "_" & lngRow.ToString()
            oSelect = chkSelect
        Else
            Dim radSelect As New RadioButton
            radSelect.GroupName = "sel_" & Me.ClientID
            radSelect.ID = "sel_" & Me.ClientID & "_" & lngRow.ToString
            oSelect = radSelect
        End If
        '-----------------------------------
        With oSelect.Attributes
            If AutoPostBack Then
                .Add("onclick", "javascript:" & Page.ClientScript.GetPostBackEventReference(Me, lngRow.ToString))
            Else
                .Add("onclick", "javascript:return atualiza_valor(this, 'row_" & Me.ClientID & "_" & lngRow.ToString & "', " & dts.Tables("dtaLista").Rows.Count.ToString & ");")
            End If
            oSelect.Checked = dts.Tables("dtaLista").Rows(lngRow)("Selected")
        End With
        '-----------------------------------
        Return oSelect
        '-----------------------------------
    End Function

    Private Sub LoadFieldValue()
        Dim Request As System.Web.HttpRequest = System.Web.HttpContext.Current.Request
        Dim o As Object
        '----------------------------------------------------------------------
        If BoxStyle = TBoxStyle.ListBox Then
            '------------------------------------------------------------------
            If MultiSelect Then
                For i As Long = 0 To dts.Tables("dtaLista").Rows.Count - 1
                    o = Request(RetornaName(True) & "sel_" & Me.ClientID & "_" & i.ToString)
                    dts.Tables("dtaLista").Rows(i)("Selected") = (o IsNot Nothing)
                Next
            Else
                o = Request("sel_" & Me.ClientID)
                If o IsNot Nothing Then
                    Dim aObj As Array = o.ToString.Split("_")
                    For i As Long = 0 To dts.Tables("dtaLista").Rows.Count - 1
                        dts.Tables("dtaLista").Rows(i)("Selected") = (aObj(2) = i)
                    Next
                End If
            End If
            '------------------------------------------------------------------
        Else
            '------------------------------------------------------------------
            o = Request("txtTextBox_" & Me.ClientID)
            '------------------------------------------------------------------
            If o IsNot Nothing Then
                strText = CType(o, String)
            Else
                strText = String.Empty
            End If
            '------------------------------------------------------------------
        End If
        '------------------------------------------------------------------
    End Sub

    Private Function RetornaName(ByVal blnUniqueID As Boolean) As String
        If Me.Parent.GetType.ToString = "System.Web.UI.WebControls.ContentPlaceHolder" Then
            If blnUniqueID 
                Return Me.NamingContainer.UniqueID & "$"
            Else
                Return Me.NamingContainer.ClientID & "_"
            End If
        End If
        Return ""
    End Function

    Private Sub BuildTables()
        With dts.Tables
            .Add("dtaLista")
            .Item("dtaLista").Columns.Add("Selected")
            '------------------------------------------
            .Add("dtaColumnProperty")
            With .Item("dtaColumnProperty").Columns
                .Add("ColumnName")
                .Add("ColumnTitle")
                .Add("Visible")
                .Add("Width")
            End With
            '------------------------------------------
            AddProperty("Selected", "Selected", False)
            '------------------------------------------
        End With
    End Sub

    Private Function RetScript() As String
        Dim strScript As New StringBuilder
        '--------------------------------------------------------
        With strScript
            .AppendLine()
            .AppendLine("function atualiza_valor(chk, row_name, qtd) {")
            .AppendLine("    /*----------------------------------------*/")
            .AppendLine("    var row = document.getElementById('" & RetornaName(False) & "' + row_name);")
            .AppendLine("    /*----------------------------------------*/")
            .AppendLine("    if (chk.type == 'radio') {")
            .AppendLine("        /*----------------------------------------*/")
            .AppendLine("        row.className = 'rowSelected';")
            .AppendLine("        /*----------------------------------------*/")
            .AppendLine("        var aNome = row_name.split('_');")
            .AppendLine("        var sNome = aNome[0] + '_' + aNome[1];")
            .AppendLine("        var sSeq  = aNome[2];    ")
            .AppendLine("        var unrow;")
            .AppendLine("        /*----------------------------------------*/    ")
            .AppendLine("        for(var i=0; i < qtd; i++) {         ")
            .AppendLine("           if (aNome[2] != i){")
            .AppendLine("                unrow = document.getElementById(aNome[0] + '_' + aNome[1] + '_' + i);")
            .AppendLine("                if (unrow) unrow.className = ""rowNormal"";")
            .AppendLine("           }")
            .AppendLine("        }")
            .AppendLine("    }else{     ")
            .AppendLine("       if (chk.checked) ")
            .AppendLine("           row.className = ""rowSelected"";")
            .AppendLine("       else")
            .AppendLine("           row.className = ""rowNormal"";")
            .AppendLine("    }")
            .AppendLine("    /*----------------------------------------*/")
            .AppendLine("}")

            '.AppendLine("function atualiza_valor(chk, row_name) {")
            '.AppendLine("    /*----------------------------------------*/")
            '.AppendLine("    var row = document.getElementById(row_name);")
            '.AppendLine("    /*----------------------------------------*/")
            '.AppendLine("       if (chk.checked) ")
            '.AppendLine("           row.className = ""rowSelected"";")
            '.AppendLine("       else")
            '.AppendLine("           row.className = ""rowNormal"";")
            '.AppendLine("    /*----------------------------------------*/")
            '.AppendLine("}")
        End With
        '--------------------------------------------------------
        Return strScript.ToString()
        '--------------------------------------------------------
    End Function

    Public Sub BindSource()
        '----------------------------------------
        If AutoGenerateColumns Then
            Call BindSourceGenerateColumns()
        Else
            If dts.Tables("dtaLista").Columns.Count = 0 Then
                Throw New Exception("Nenhuma coluna foi encontrada!")
            Else
                Call BindSourceWithoutGenerateColumns()
            End If
        End If
        '----------------------------------------
        Call LoadFieldValue()
        '----------------------------------------        
        Call AddColumnsProperty()
        '----------------------------------------
    End Sub

    Private Sub BindSourceGenerateColumns()
        '----------------------------------------
        If cnn.State = ConnectionState.Closed Then
            adp = New OleDb.OleDbDataAdapter(CommandText, ConnectionString)
        Else
            adp = New OleDb.OleDbDataAdapter(CommandText, cnn)
        End If
        '----------------------------------------
        dts.Tables("dtaLista").Rows.Clear()
        adp.Fill(dts.Tables("dtaLista"))
        Call ClearSelections()
        '----------------------------------------
    End Sub

    Private Sub BindSourceWithoutGenerateColumns()
        Dim dta As New DataTable
        Dim dtr As DataRow
        '--------------------------------------
        If cnn.State = ConnectionState.Closed Then
            adp = New OleDb.OleDbDataAdapter(CommandText, ConnectionString)
        Else
            adp = New OleDb.OleDbDataAdapter(CommandText, cnn)
        End If
        '--------------------------------------
        dts.Tables("dtaLista").Rows.Clear()
        adp.Fill(dta)
        '--------------------------------------
        With dts.Tables("dtaLista")
            For Each dtrow As DataRow In dta.Rows
                dtr = .NewRow()
                For Each dtc As DataColumn In .Columns
                    If dtc.ColumnName <> "Selected" Then
                        dtr(dtc.ColumnName) = dtrow(dtc.ColumnName)
                    Else
                        dtr(dtc.ColumnName) = False
                    End If
                Next
                .Rows.Add(dtr)
            Next
        End With
        '--------------------------------------
    End Sub

    Private Sub AddColumnsProperty()
        For Each c As DataColumn In dts.Tables("dtaLista").Columns
            AddProperty(c.ColumnName, c.ColumnName)
        Next
    End Sub

    Public Sub MoveRow(ByVal IndexRow As Long, ByVal NewIndex As Long)
        Dim dtNewRow As DataRow
        Dim dtOldRow As DataRow
        Dim dtaLista As DataTable = dts.Tables("dtaLista")
        '------------------------------------------------
        If IndexRow = NewIndex Then Exit Sub
        '------------------------------------------------
        dtNewRow = dtaLista.NewRow
        dtOldRow = dtaLista.Rows(IndexRow)
        '------------------------------------------------
        dtNewRow.BeginEdit()
        For Each dtcol As DataColumn In dtaLista.Columns
            dtNewRow(dtcol.ColumnName) = dtOldRow(dtcol.ColumnName)
        Next
        dtNewRow.EndEdit()
        '------------------------------------------------
        With dtaLista.Rows
            If IndexRow > NewIndex Then
                .Remove(dtOldRow)
                .InsertAt(dtNewRow, NewIndex)
            Else
                .InsertAt(dtNewRow, NewIndex + 1)
                .Remove(dtOldRow)
            End If
        End With
        '------------------------------------------------
    End Sub

    ''' <summary>
    ''' Remove a(s) seleção(ões) existentes.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearSelections()
        Dim dtr() As DataRow = dts.Tables("dtaLista").Select("Selected = True Or Selected Is NULL")
        Dim lngQtdeRows As Long = dtr.Length
        '-------------------------------------------------------
        For i As Integer = 0 To lngQtdeRows - 1
            dtr(i).BeginEdit()
            dtr(i)("Selected") = False
            dtr(i).EndEdit()
        Next
        '-------------------------------------------------------
    End Sub

    ''' <summary>
    ''' Remove todas as linhas da lista
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearRows()
        CommandText = String.Empty
        dts.Tables("dtaLista").Clear()
    End Sub

#End Region


#Region "Events"

    ''' <summary>
    ''' Evento disparado quando um ítem da lista é clicado (AutoPostBack = true).
    ''' </summary>
    ''' <param name="RowClicked">A linha que foi clicada.</param>
    ''' <remarks></remarks>
    Public Event RowClicked(ByRef RowClicked As DataRow, ByVal RowIndex As Long)

    Protected Overrides Sub RenderContents(ByVal writer As System.Web.UI.HtmlTextWriter)
        '------------------------------
        Me.Controls.Clear()
        '------------------------------
        If Me.DesignMode And BoxStyle = TBoxStyle.ListBox Then ShowDesignMode()
        Call RenderCustomControl()
        '------------------------------ 
        MyBase.RenderContents(writer)
        '------------------------------
    End Sub

    Protected Overrides Function SaveControlState() As Object
        '----------------------------------------
        If BoxStyle = TBoxStyle.ListBox Then
            Return dts
        Else
            Return MyBase.SaveControlState()
        End If
        '----------------------------------------
    End Function

    Protected Overrides Sub LoadControlState(ByVal savedState As Object)
        If BoxStyle = TBoxStyle.ListBox Then
            dts = savedState
        Else
            MyBase.LoadControlState(savedState)
        End If
        '----------------------------------------
        Call LoadFieldValue()
        '----------------------------------------
    End Sub

    Protected Overrides Sub OnInit(ByVal e As System.EventArgs)
        Page.RegisterRequiresControlState(Me)
        If Not Page.ClientScript.IsClientScriptBlockRegistered("cxScriptBox") Then _
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "cxScriptBox", RetScript, True)
        '----------------------------------------        
        MyBase.OnInit(e)
    End Sub

    Public Sub RaisePostBackEvent(ByVal eventArgument As String) Implements System.Web.UI.IPostBackEventHandler.RaisePostBackEvent
        RaiseEvent RowClicked(dts.Tables("dtaLista").Rows(eventArgument), eventArgument)
    End Sub

#End Region


End Class
