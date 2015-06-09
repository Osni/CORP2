Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Data



<ToolboxData("<{0}:ClsAutoComplete runat=server></{0}:ClsAutoComplete>")> _
Public Class ClsAutoComplete
    Inherits WebControl


#Region "Fields"

    Private strValue As String = String.Empty
    Private strText As String = String.Empty

#End Region


#Region "Events"

    Protected Overrides Sub RenderContents(ByVal writer As System.Web.UI.HtmlTextWriter)
        Me.Controls.Clear()
        Call RenderAutoComplete()
        MyBase.RenderContents(writer)
    End Sub

    Protected Overrides Sub OnInit(ByVal e As System.EventArgs)
        With Page
            .RegisterRequiresControlState(Me)
            With .ClientScript
                If Not .IsClientScriptIncludeRegistered("scrpt_AutoComplete") Then
                    If ScriptURL <> "" Then
                        .RegisterClientScriptInclude(Page.GetType, "scrpt_AutoComplete", ScriptURL)
                    End If
                End If
            End With
        End With
        '--------------------------------------------------------
        'Recuperando os valores do componente
        Dim Request As HttpRequest = HttpContext.Current.Request
        '-------------------------------------------------------        
        If Request("hdn" & Me.ClientID) IsNot Nothing Then
            strValue = Request("hdn" & Me.ClientID)
            If strValue.Trim <> String.Empty Then
                If Request(Me.ClientID & "_txt") IsNot Nothing Then strText = Request(Me.ClientID & "_txt")
            Else
                strText = String.Empty
            End If
        End If
        '--------------------------------------------------------
        MyBase.OnInit(e)
    End Sub

    Protected Overrides Function SaveControlState() As Object
        Return MyBase.SaveControlState()
    End Function

    Protected Overrides Sub LoadControlState(ByVal savedState As Object)
        MyBase.LoadControlState(savedState)
    End Sub

#End Region


#Region "Enumerations"

    Public Enum TACStyle As Short
        AutoComplete = 1
        ButtonClick = 2
        Enter = 3
    End Enum

    Public Enum TCaptionStyle As Short
        AlignTop
        AlignLeft
    End Enum

    Public Enum TButtonStyle As Short
        ImageButton
        SimpleButton
    End Enum

#End Region


#Region "Properties"

    ''' <summary>
    ''' Largura da caixa de Listagem.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ListWidth() As Unit
        Get
            If ViewState("ListWidth") Is Nothing Then ViewState("ListWidth") = Me.Width
            Return ViewState("ListWidth")
        End Get
        Set(ByVal value As Unit)
            ViewState("ListWidth") = value
        End Set
    End Property

    ''' <summary>
    ''' Método utilizado para o envio da requisição ajax. GET ou POST.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Method() As String
        Get
            If ViewState("Method") Is Nothing Then ViewState("Method") = "GET"
            Return ViewState("Method")
        End Get
        Set(ByVal value As String)
            If value.ToUpper <> "GET" And value.ToUpper <> "POST" Then
                Throw New Exception("Método '" & value & "' não é valido.")
            Else
                ViewState("Method") = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Provider de dados utilizado.
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

    Public Property ScriptURL() As String
        Get
            If ViewState("_ScriptURL") Is Nothing Then ViewState("_ScriptURL") = "corp_autocomplete.js"
            Return ViewState("_ScriptURL")
        End Get
        Set(ByVal value As String)
            ViewState("_ScriptURL") = value
        End Set
    End Property

    'Private ReadOnly Property txtID() As String
    '    Get
    '        If Me.NamingContainer Is Nothing Then
    '            If ViewState("txtID") Is Nothing Then ViewState("txtID") = "txt" & Me.ClientID
    '        Else
    '            If ViewState("txtID") Is Nothing Then ViewState("txtID") = Me.NamingContainer.ClientID & "$txt" & Me.ClientID
    '        End If
    '        txtID = ViewState("txtID")
    '    End Get
    'End Property

    ''' <summary>
    ''' String de conexão da fonte de dados.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ConnectionString() As String
        Get
            If ViewState("_ConnectionString") Is Nothing Then ViewState("_ConnectionString") = String.Empty
            Return ViewState("_ConnectionString")
        End Get
        Set(ByVal value As String)
            ViewState("_ConnectionString") = value
        End Set
    End Property

    ''' <summary>
    ''' Tabela de origem das informações.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TableSelect() As String
        Get
            If ViewState("_TableSelect") Is Nothing Then ViewState("_TableSelect") = String.Empty
            Return ViewState("_TableSelect")
        End Get
        Set(ByVal value As String)
            ViewState("_TableSelect") = value
        End Set
    End Property

    ''' <summary>
    ''' Coluna da tabela informada em "TableSelect" correspondente ao texto a ser exibido.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnText() As String
        Get
            If ViewState("_ColumnText") Is Nothing Then ViewState("_ColumnText") = String.Empty
            Return ViewState("_ColumnText")
        End Get
        Set(ByVal value As String)
            ViewState("_ColumnText") = value
        End Set
    End Property

    ''' <summary>
    ''' Coluna da tabela informada em "TableSelect" correspondente ao valor a ser gravado.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnValue() As String
        Get
            If ViewState("_ColumnValue") Is Nothing Then ViewState("_ColumnValue") = String.Empty
            Return ViewState("_ColumnValue")
        End Get
        Set(ByVal value As String)
            ViewState("_ColumnValue") = value
        End Set
    End Property

    ''' <summary>
    ''' Coluna da tabela informada em TableSelect para compor a cláusula WHERE.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnWhere() As String
        Get
            If ViewState("_ColumnWhere") Is Nothing Then ViewState("_ColumnWhere") = String.Empty
            Return ViewState("_ColumnWhere")
        End Get
        Set(ByVal value As String)
            ViewState("_ColumnWhere") = value
        End Set
    End Property


    ''' <summary>
    ''' String contida no campo texto.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Text() As String
        Get
            ViewState("_Text") = strText
            Return ViewState("_Text")
        End Get
        Set(ByVal value As String)
            strText = value
            ViewState("_Text") = value
        End Set
    End Property

    ''' <summary>
    ''' String contida no campo valor.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Value() As String
        Get            
            ViewState("_Value") = strValue
            Return ViewState("_Value")
        End Get
        Set(ByVal value As String)
            strValue = value
            ViewState("_Value") = value
        End Set
    End Property

    Public Property ShowCaption() As Boolean
        Get
            If ViewState("_ShowCaption") Is Nothing Then ViewState("_ShowCaption") = False
            Return ViewState("_ShowCaption")
        End Get
        Set(ByVal value As Boolean)
            ViewState("_ShowCaption") = value
        End Set
    End Property

    ''' <summary>
    ''' Rótulo para o controle.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Caption() As String
        Get
            If ViewState("_Caption") Is Nothing Then ViewState("_Caption") = "Caption" & Me.ClientID
            Return ViewState("_Caption")
        End Get
        Set(ByVal value As String)
            ViewState("_Caption") = value
        End Set
    End Property

    ''' <summary>
    ''' Estilo do rótulo. AlignTop ou AlignLeft(Default).
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CaptionStyle() As TCaptionStyle
        Get
            If ViewState("_CaptionStyle") Is Nothing Then ViewState("_CaptionStyle") = TCaptionStyle.AlignLeft
            Return ViewState("_CaptionStyle")
        End Get
        Set(ByVal value As TCaptionStyle)
            Select Case value
                Case TCaptionStyle.AlignLeft, TCaptionStyle.AlignTop
                    ViewState("_CaptionStyle") = value
                Case Else
                    Throw New Exception("Estilo informado é inválido!")
            End Select

        End Set
    End Property

    ''' <summary>
    ''' Imagem para exibição no botão. Para estilo ButtonClick.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ButtonImgSrc() As String
        Get
            If ViewState("_ImgSrc") Is Nothing Then ViewState("_ImgSrc") = String.Empty
            Return ViewState("_ImgSrc")
        End Get
        Set(ByVal value As String)
            ViewState("_ImgSrc") = value
        End Set
    End Property

    ''' <summary>
    ''' Rótulo a ser exibido no botão. Para estilo ButtonClick.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ButtonCaption() As String
        Get
            If ViewState("_ButtonCaption") Is Nothing Then ViewState("_ButtonCaption") = "btn" & Me.ClientID
            Return ViewState("_ButtonCaption")
        End Get
        Set(ByVal value As String)
            ViewState("_ButtonCaption") = value
        End Set
    End Property

    ''' <summary>
    ''' Estilo do botão. Default SimpleButton (botão com rótulo de texto).
    ''' A outra opção é ImageButton.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ButtonStyle() As TButtonStyle
        Get
            If ViewState("_ButtonStyle") Is Nothing Then ViewState("_ButtonStyle") = TButtonStyle.SimpleButton
            Return ViewState("_ButtonStyle")
        End Get
        Set(ByVal value As TButtonStyle)
            Select Case value
                Case TButtonStyle.ImageButton, TButtonStyle.SimpleButton
                    ViewState("_ButtonStyle") = value
                Case Else
                    Throw New Exception("Estilo informado é inválido!")
            End Select
        End Set
    End Property

    ''' <summary>
    ''' Quantidade de mínima caracteres para acionar para ativar a consulta. Default 3 caracteres.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property KeySize() As Short
        Get
            If ViewState("_KeySize") Is Nothing Then ViewState("_KeySize") = 3
            Return ViewState("_KeySize")
        End Get
        Set(ByVal value As Short)
            ViewState("_KeySize") = value
        End Set
    End Property

    ''' <summary>
    ''' Estilo do componente. AutoComplete é ativado conforme a digitação; 
    ''' ButtonClick é acionado com um clique no botão. Default é AutoComplete.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AutoCompleteStyle() As TACStyle
        Get
            If ViewState("_AutoCompleteStyle") Is Nothing Then ViewState("_AutoCompleteStyle") = TACStyle.AutoComplete
            Return ViewState("_AutoCompleteStyle")
        End Get
        Set(ByVal value As TACStyle)
            Select Case value
                Case TACStyle.AutoComplete, TACStyle.ButtonClick, TACStyle.Enter
                    ViewState("_AutoCompleteStyle") = value
                Case Else
                    Throw New Exception("Estilo informado é inválido!")
            End Select
        End Set
    End Property

    ''' <summary>
    ''' Define a cor de fundo do campo de texto.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Property BackColor() As System.Drawing.Color
        Get
            Dim o As Object = ViewState("_BackColor")
            If o Is Nothing Then ViewState("_BackColor") = System.Drawing.Color.White
            Return ViewState("_BackColor")
        End Get
        Set(ByVal value As System.Drawing.Color)
            ViewState("_BackColor") = value
        End Set
    End Property


#End Region


#Region "Methods"


    Private Sub RenderAutoComplete()
        Dim tabAC As New Table
        Dim rowAC As TableRow
        Dim celAC As TableCell
        Dim txtAC As New TextBox
        Dim btnAC As Object
        Dim strParCrip As String
        Dim strParam As New StringBuilder
        '----------------------------------------------------------------------------
        With tabAC
            .ID = "tabAutoComplete_" & Me.ClientID
            .CssClass = "tabAutoComplete"
            .BorderStyle = WebControls.BorderStyle.None
            .Style.Add("padding", "0 0 0 0")
            .Style.Add("margin", "0 0 0 0")
            .CellPadding = 0
            .CellSpacing = 0
            .Style.Add("width", "0px")
            .Style.Add("height", "0px")
        End With
        '----------------------------------------------------------------------------
        rowAC = New TableRow
        '----------------------------------------------------------------------------
        If ShowCaption Then
            celAC = New TableHeaderCell
            '----------------------------------------------------------------------------
            With celAC
                .Text = Caption.Replace(" ", "&nbsp;")
                .VerticalAlign = VerticalAlign.Middle
                .BorderStyle = WebControls.BorderStyle.None
                .Style.Add("width", "100%")
                .ColumnSpan = IIf(CaptionStyle = TCaptionStyle.AlignTop And AutoCompleteStyle = TACStyle.ButtonClick, 2, 1)
            End With
            '----------------------------------------------------------------------------
            rowAC.Cells.Add(celAC)
            '----------------------------------------------------------------------------
            If CaptionStyle = TCaptionStyle.AlignTop Then
                tabAC.Rows.Add(rowAC)
                rowAC = New TableRow
            End If
        End If
        '----------------------------------------------------------------------------
        celAC = New TableCell
        '----------------------------------------------------------------------------
        With txtAC
            .Attributes.Add("autocomplete", "off")
            .ID = Me.ClientID & "_txt"
            .CssClass = "txtAutC"
            .Text = Text
            .Width = Me.Width
            .BackColor = BackColor
            .Style.Add("position", "relative")
        End With
        '----------------------------------------------------------------------------
        With celAC
            .VerticalAlign = VerticalAlign.Middle
            .Controls.Add(txtAC)
            .BorderStyle = WebControls.BorderStyle.None
            .Style.Add("padding", "0 0 0 0")
            'If AutoCompleteStyle = TACStyle.ButtonClick Then
            '    .Controls.Add(New LiteralControl("<br><div id=""div" & txtAC.ID & """ class=""msgAutC"" style=""display:none""></div>"))
            'Else
            '    .Controls.Add(New LiteralControl("<div id=""divMsg" & txtAC.ID & """ class=""msgAutC"" style=""display:none""></div><br><div id=""div" & txtAC.ID & """ class=""msgAutC"" style=""display:none""></div>"))
            'End If
        End With
        rowAC.Cells.Add(celAC)
        '----------------------------------------------------------------------------
        If AutoCompleteStyle = TACStyle.ButtonClick Then
            '------------------------------------------------------------------------
            If ButtonStyle = TButtonStyle.ImageButton Then
                btnAC = New ImageButton
                With btnAC
                    .cssClass = "btnAutC"
                    .ImageAlign = ImageAlign.AbsMiddle
                    .ImageUrl = ButtonImgSrc
                    .ID = "btn" & Me.ClientID
                    .style.Add("position", "relative")
                    .Width = Unit.Pixel(24)
                    .Height = Unit.Pixel(24)
                End With
            Else
                btnAC = New Button
                With btnAC
                    .cssClass = "btnAutC"
                    .style.Add("position", "relative")
                    .Text = ButtonCaption
                    .ID = "btn" & Me.ClientID
                End With
            End If
            '------------------------------------------------------------------------
            celAC = New TableCell
            With celAC
                .BorderStyle = WebControls.BorderStyle.None
                .VerticalAlign = VerticalAlign.Middle
                .Style.Add("padding", "0 0 0 0")
                .Controls.Add(btnAC)
            End With
            'celAC.Controls.Add(New LiteralControl("<div id=""divMsg" & txtAC.ID & """ class=""msgAutC"" style=""display: none""></div>"))
            rowAC.Cells.Add(celAC)
            '------------------------------------------------------------------------
        End If
        '----------------------------------------------------------------------------
        tabAC.Rows.Add(rowAC)
        '----------------------------------------------------------------------------
        With Me.Controls
            .Add(tabAC)
            .Add(New LiteralControl("<input type=hidden id=""hdn" & Me.ClientID & """ name=""hdn" & Me.ClientID & """ value=""" & Value & """> "))
            '------------------------------------------------------------------------
            strParCrip = CorpCripto.EncryptString(ConnectionString & "$" & TableSelect & "$" & ColumnValue & "$" & ColumnText & "$" & ColumnWhere & "$" & Provider)
            '------------------------------------------------------------------------
            With strParam
                .AppendLine("<script>")                
                .AppendLine("aut.add(""" & txtAC.ClientID & """, ""hdn" & Me.ClientID & """, """ & strParCrip & """ , " & KeySize & "," & AutoCompleteStyle & ", """ & Method & """, """ & ListWidth.ToString & """" & IIf(AutoCompleteStyle = TACStyle.ButtonClick, ", ""btn" & Me.ClientID & """", "") & ");")
                .AppendLine("</script>")
            End With
            '------------------------------------------------------------------------
            .Add(New LiteralControl(strParam.ToString))
            '------------------------------------------------------------------------
        End With
        '----------------------------------------------------------------------------
    End Sub

    ''' <summary>
    ''' Limpa conteúdo do campo.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()
        strText = String.Empty
        strValue = String.Empty
    End Sub

#End Region


End Class
