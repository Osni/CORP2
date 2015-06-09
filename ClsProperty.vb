Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls

<DefaultProperty(""), ToolboxData("<{0}:ClsProperty runat=server></{0}:ClsProperty>")> _
Public Class ClsProperty
    Inherits WebControl

    Public Sub New()
        ds = New DataSet        
        mPropertyTable = New DataTable("TAB_PROPERTIES")
        With mPropertyTable.Columns
            .Add("ProName")
            .Add("ProCaption")            
            .Add("ProStyle", System.Type.GetType("System.Int16"))
            .Add("ProDataType")
            .Add("ProDefaultValue")
            .Add("ProSource")
            .Add("ProColumnText")
            .Add("ProColumnValue")
            .Add("ProListDelimiter")
            .Add("ProSourceList")
            .Add("ProActualValue")
            .Add("ProReadOnly")
        End With
    End Sub

#Region "Enumarations"

    Public Enum TFieldStyle As Short
        DROPDOWN
        TEXTBOX
    End Enum

    Public Enum TDataType As Short
        TEXT
        NUMERIC_INT
        NUMERIC_FLOAT
        DATETIME
    End Enum

#End Region


#Region "Fields"
    Private ds As DataSet
    Private mPropertyTable As DataTable
    Private ClsDB As New ClsDB
#End Region


#Region "Properties"

    Public Property PropertyBoxTitle() As String
        Get
            Dim o As Object = ViewState("_PropertyBoxTitle")
            If o Is Nothing Then ViewState("_PropertyBoxTitle") = String.Empty
            Return ViewState("_PropertyBoxTitle")
        End Get
        Set(ByVal value As String)
            ViewState("_PropertyBoxTitle") = value
        End Set
    End Property


    Public Property PropertyValue(ByVal PropertyName As String) As Object
        Get
            Dim row As DataRow = GetProperty(PropertyName)
            '---------------------------------------------
            If row IsNot Nothing Then
                Return row("ProActualValue")
            Else
                Throw New Exception("Propriedade não encontrada!")
            End If
        End Get
        Set(ByVal value As Object)
            Dim row As DataRow = GetProperty(PropertyName)
            '---------------------------------------------
            If row IsNot Nothing Then
                row("ProActualValue") = value
            Else
                Throw New Exception("Propriedade não encontrada!")
            End If
        End Set
    End Property

    Public ReadOnly Property PropertyNames() As Dictionary(Of String, String).KeyCollection
        Get
            Dim colNames As New Dictionary(Of String, String)

            For Each dtrow As DataRow In mPropertyTable.Rows
                colNames.Add(dtrow("ProName"), dtrow("ProName"))
            Next
            Return colNames.Keys
        End Get
    End Property

#End Region


#Region "Events"

    Protected Overrides Sub RenderContents(ByVal output As HtmlTextWriter)
        Me.Controls.Clear()
        If Me.DesignMode Then
            Me.Controls.Add(New LiteralControl("<div valign=middle style=""text-align:center;"">" & Me.ClientID & "</div>"))
        Else
            RenderProperties()
        End If
        MyBase.RenderContents(output)
    End Sub

    Protected Overrides Function SaveControlState() As Object
        If ds.Tables("TAB_PROPERTIES") IsNot Nothing Then ds.Tables.Remove("TAB_PROPERTIES")
        ds.Tables.Add(mPropertyTable)
        Return ds
    End Function

    Protected Overrides Sub LoadControlState(ByVal savedState As Object)
        ds = savedState
        mPropertyTable = ds.Tables("TAB_PROPERTIES")
        Call AttribuiValores()
    End Sub

    Protected Overrides Sub OnInit(ByVal e As System.EventArgs)
        Page.RegisterRequiresControlState(Me)
        MyBase.OnInit(e)
    End Sub
#End Region


#Region "Methods"

    ''' <summary>
    ''' Adiciona propriedade com característica de TextBox.
    ''' </summary>
    ''' <param name="PropertyName">Nome da Propriedade</param>
    ''' <param name="PropertyCaption">Rótulo que será exibido.</param>
    ''' <param name="DefaultValue">Valor Padrão.</param>
    ''' <param name="IsReadOnly">Se o campo é somente leitura.</param> 
    ''' <param name="TypeData">O tipo de dado de entrada.</param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal PropertyName As String, _
               ByVal PropertyCaption As String, _
               Optional ByVal TypeData As TDataType = TDataType.TEXT, _
               Optional ByVal DefaultValue As Object = Nothing, _
               Optional ByVal IsReadOnly As Boolean = False)

        Dim dtRow As DataRow = mPropertyTable.NewRow()

        dtRow("ProName") = PropertyName
        dtRow("ProCaption") = PropertyCaption
        dtRow("ProStyle") = TFieldStyle.TEXTBOX
        dtRow("ProDefaultValue") = DefaultValue
        dtRow("ProSource") = ""
        dtRow("ProActualValue") = IIf(IsNothing(DefaultValue), "", DefaultValue)
        dtRow("ProReadOnly") = IsReadOnly

        mPropertyTable.Rows.Add(dtRow)

    End Sub

    ''' <summary>
    ''' Adiciona propriedade com características de Dropdownlist
    ''' tendo como fonte de dados um DataTable.
    ''' </summary>
    ''' <param name="PropertyName">Nome da Propriedade</param>
    ''' <param name="PropertyCaption">Rótulo que será exibido.</param>
    ''' <param name="DefaultValue">Valor Padrão.</param>
    ''' <param name="DataSource">DataTable contendo as informações.</param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal PropertyName As String, _
                   ByVal PropertyCaption As String, _
                   ByVal DataSource As DataTable, _
                   ByVal ColumnValue As String, _
                   ByVal ColumnText As String, _
                   Optional ByVal DefaultValue As Object = Nothing)


        Dim dtRow As DataRow = mPropertyTable.NewRow()
        '--------------------------------------------------------
        If DataSource.TableName.Trim = String.Empty Then DataSource.TableName = "TAB_" & PropertyName
        '--------------------------------------------------------
        dtRow("ProName") = PropertyName
        dtRow("ProCaption") = PropertyCaption
        dtRow("ProStyle") = TFieldStyle.DROPDOWN
        dtRow("ProDefaultValue") = DefaultValue
        dtRow("ProColumnText") = ColumnText
        dtRow("ProColumnValue") = ColumnValue
        dtRow("ProSource") = DataSource.TableName
        dtRow("ProActualValue") = IIf(IsNothing(DefaultValue), "", DefaultValue)
        dtRow("ProReadOnly") = False
        '--------------------------------------------------------
        mPropertyTable.Rows.Add(dtRow)
        ds.Tables.Add(DataSource)
        '--------------------------------------------------------
    End Sub

    ''' <summary>
    ''' Adiciona propriedade com características de Dropdownlist
    ''' tendo como fonte de dados um DataTable.
    ''' </summary>
    ''' <param name="PropertyName">Nome da Propriedade</param>
    ''' <param name="PropertyCaption">Rótulo que será exibido.</param>
    ''' <param name="DefaultValue">Valor Padrão.</param>        
    ''' <param name="SourceList">Lista de valores separados por um delimitador.</param>
    ''' <param name="ListDelimiter">Delimitador utilizado para separarar os ítens da lista.</param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal PropertyName As String, _
               ByVal PropertyCaption As String, _
               ByVal SourceList As String, _
               ByVal ListDelimiter As String, _
               Optional ByVal DefaultValue As Object = Nothing)

        Dim dtRow As DataRow = mPropertyTable.NewRow()

        dtRow("ProName") = PropertyName
        dtRow("ProCaption") = PropertyCaption
        dtRow("ProStyle") = TFieldStyle.DROPDOWN
        dtRow("ProDefaultValue") = DefaultValue
        dtRow("ProListDelimiter") = ListDelimiter
        dtRow("ProSourceList") = SourceList
        dtRow("ProActualValue") = IIf(IsNothing(DefaultValue), "", DefaultValue)
        dtRow("ProReadOnly") = False

        mPropertyTable.Rows.Add(dtRow)
    End Sub

    Private Sub RenderProperties()
        Dim tabProperty As New Table
        Dim tabPropertyRows As New Table
        Dim row As New TableRow
        Dim cell As TableCell
        Dim txtProperty As TextBox
        Dim cboProperty As DropDownList
        Dim objProperty As Object
        Dim tamCpo As Long
        '-----------------------------------------------
        'Configurações básicas da tabela Geral de propriedades
        With tabProperty
            .ID = "tablePropertyBox"
            .Width = Me.Width
            .Height = Me.Height
            .CellSpacing = 0
            .CellPadding = 0
            tamCpo = (Me.Width.Value / 2) - 5
        End With
        '-----------------------------------------------
        'Configurações básicas da tabela de propriedades
        With tabPropertyRows
            .ID = "tablePropertyRows"
            .Width = Unit.Percentage(100)
            .CellSpacing = 0
            .CellPadding = 0
        End With
        '-----------------------------------------------
        'Título da caixa de propriedades
        row = New TableRow
        cell = New TableHeaderCell
        With cell
            .ID = "PropertyBoxTitle"
            .Text = PropertyBoxTitle
            .ColumnSpan = 2            
        End With
        row.Cells.Add(cell)
        tabProperty.Rows.Add(row)
        '-----------------------------------------------
        For Each dtRow As DataRow In mPropertyTable.Rows
            '-----------------------------------
            row = New TableRow
            '-----------------------------------
            'Rótulo da propriedade
            cell = New TableHeaderCell
            cell.Style.Add("width", "50%")
            cell.Text = dtRow("ProCaption")
            row.Cells.Add(cell)
            '-----------------------------------
            'Campo da propriedade
            cell = New TableCell
            cell.Style.Add("width", "50%")
            '-----------------------------------
            If CType(dtRow("ProStyle"), TFieldStyle) = TFieldStyle.TEXTBOX Then
                '---------------------------------------------
                txtProperty = New TextBox
                With txtProperty
                    .ID = GetFieldName(dtRow)
                    .Width = Unit.Pixel(tamCpo)
                    .ReadOnly = dtRow("ProReadOnly")
                End With
                objProperty = txtProperty
                '---------------------------------------------
                '???????????
                'dtRow("ProDataType")  'validação do tipo de informação que será inserida
                '???????????
                If ClsDB.NullDB(dtRow("ProActualValue")) <> String.Empty Then
                    txtProperty.Text = dtRow("ProActualValue")
                Else
                    txtProperty.Text = ClsDB.NullDB(dtRow("ProDefaultValue"))
                End If
                '---------------------------------------------
            Else
                '---------------------------------------------
                Dim cboItem As ListItem
                '---------------------------------------------
                cboProperty = New DropDownList
                With cboProperty
                    .ID = GetFieldName(dtRow)
                    .Width = Unit.Pixel(tamCpo)
                End With
                objProperty = cboProperty
                '---------------------------------------------
                With cboProperty.Items
                    '---------------------------------------------
                    If Not IsDBNull(dtRow("ProSource")) Then
                        For Each sourcerow As DataRow In ds.Tables(dtRow("ProSource").ToString).Rows
                            cboItem = New ListItem
                            cboItem.Text = sourcerow(dtRow("ProColumnText")).ToString()
                            cboItem.Value = sourcerow(dtRow("ProColumnValue")).ToString()
                            cboItem.Selected = IsSelected(dtRow, cboItem.Value)
                            .Add(cboItem)
                        Next
                    Else
                        Dim aDelimiters As Array = dtRow("ProListDelimiter").ToString.Trim.ToCharArray
                        Dim aItens As Array
                        '------------------------------------------------
                        aItens = dtRow("ProSourceList").ToString.Trim.Split(aDelimiters(0))
                        '------------------------------------------------                    
                        If aDelimiters.Length = 2 Then
                            Dim aItens2 As Array
                            '---------------------------------
                            For i As Integer = 0 To aItens.Length - 1
                                cboItem = New ListItem
                                aItens2 = aItens(i).ToString.Split(aDelimiters(1))
                                cboItem.Text = aItens2(1)
                                cboItem.Value = aItens2(0)
                                cboItem.Selected = IsSelected(dtRow, cboItem.Value)
                                .Add(cboItem)
                            Next
                            '---------------------------------
                        Else
                            '---------------------------------
                            For i As Integer = 0 To aItens.Length - 1
                                cboItem = New ListItem
                                cboItem.Text = aItens(i)                                
                                cboItem.Selected = IsSelected(dtRow, cboItem.Value)
                                .Add(cboItem)
                            Next
                            '---------------------------------
                        End If
                        '-------------------------------------
                    End If
                    '-----------------------------------------
                End With
            End If
            '-----------------------------------
            cell.Controls.Add(objProperty)
            row.Cells.Add(cell)
            tabPropertyRows.Rows.Add(row)
            '-----------------------------------
        Next
        '-----------------------------------
        row = New TableRow
        cell = New TableCell
        With cell
            .VerticalAlign = VerticalAlign.Top
            .Controls.Add(tabPropertyRows)
            .Height = Unit.Percentage(80)
        End With
        row.Cells.Add(cell)
        tabProperty.Rows.Add(row)
        '-----------------------------------
        row = New TableRow
        cell = New TableHeaderCell
        With cell
            .ID = "PropertyBoxFooter"
            .Text = "&nbsp;"            
            .ColumnSpan = 2
        End With
        row.Cells.Add(cell)
        tabProperty.Rows.Add(row)
        '-----------------------------------
        Me.Controls.Add(tabProperty)
        '-----------------------------------
    End Sub

    Private Sub AttribuiValores()
        Dim Request As System.Web.HttpRequest = System.Web.HttpContext.Current.Request
        '------------------------------------------------------
        For Each dtRow As DataRow In mPropertyTable.Rows
            dtRow("ProActualValue") = Request(GetFieldName(dtRow))
        Next
    End Sub

    Private Function GetFieldName(ByRef dtRow As DataRow) As String
        Dim strFieldName As String
        '*************************************************
        'Regra
        '    <TIPO>_PRO_<NOME_PRO>_<ClientID>
        '*************************************************
        strFieldName = IIf(CType(dtRow("ProStyle"), TFieldStyle) = TFieldStyle.DROPDOWN, "CBO_", "TXT_") & "PRO_"
        strFieldName &= dtRow("ProName").ToString.ToUpper() & "_" & Me.ClientID.ToUpper()
        Return strFieldName
    End Function

    Private Function GetFieldName(ByVal ProName As String, ByVal ProStyle As TFieldStyle) As String
        Dim strFieldName As String
        '*************************************************
        'Regra
        '    <TIPO>_PRO_<NOME_PRO>_<ClientID>
        '*************************************************
        strFieldName = IIf(ProStyle = TFieldStyle.DROPDOWN, "CBO_", "TXT_") & "PRO_"
        strFieldName &= ProName.ToUpper() & "_" & Me.ClientID.ToUpper()
        Return strFieldName
    End Function

    Private Function IsSelected(ByVal dtRow As DataRow, ByVal Value As Object) As Boolean

        If ClsDB.NullDB(dtRow("ProActualValue")) <> String.Empty Then
            Return (Value.Trim = ClsDB.NullDB(dtRow("ProActualValue")).ToString.Trim)
        Else
            Return (Value.Trim = ClsDB.NullDB(dtRow("ProDefaultValue")).ToString.Trim)
        End If

    End Function

    Private Function GetProperty(ByVal PropertyName As String) As DataRow
        Dim row() As DataRow = mPropertyTable.Select("ProName = '" & PropertyName.Trim & "'")
        '------------------------------------------
        If row.Length > 0 Then
            Return row(0)
        Else
            Return Nothing
        End If
    End Function

#End Region

End Class
