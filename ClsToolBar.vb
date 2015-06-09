Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.ComponentModel.Design
Imports System.Reflection
Imports System.Security


<ToolboxData("<{0}:ToolBar runat=server></{0}:ToolBar>"), _
ParseChildren(True, "ToolButtons"), _
DefaultProperty("ToolButtons")> _
Public Class ToolBar
    Inherits WebControl
    Implements INamingContainer, IPostBackEventHandler

#Region "Fields"
    Private btns As New ToolButtons
#End Region

#Region "Events"

    'Public Event OnClickMe(ByVal Index As Integer, ByVal ID As String)
    Public Event OnClickMe(ByRef Btn As ToolButton, ByVal Index As Integer)

    Protected Overrides Sub RenderContents(ByVal writer As System.Web.UI.HtmlTextWriter)
        Me.Controls.Clear()
        Me.Height = Unit.Pixel(1)
        If Me.DesignMode And btns.Count = 0 Then
            Me.Controls.Add(New LiteralControl("<div style='font-family:Verdana;font-size:10pt;'>No buttons was added yet.</div>"))
        Else
            Call RenderMe()
        End If
        MyBase.RenderContents(writer)
    End Sub

    Protected Overrides Function SaveViewState() As Object
        Return btns
    End Function

    Protected Overrides Sub LoadViewState(ByVal savedState As Object)
        btns = CType(savedState, ToolButtons)
    End Sub

    Protected Overrides Sub OnInit(ByVal e As System.EventArgs)
        Page.RequiresControlState(Me)
        MyBase.OnInit(e)
    End Sub

    Public Sub RaisePostBackEvent(ByVal eventArgument As String) Implements System.Web.UI.IPostBackEventHandler.RaisePostBackEvent
        Dim aArguments As Array
        Dim Response As HttpResponse = HttpContext.Current.Response
        aArguments = eventArgument.Split("$")
        Try
            'RaiseEvent OnClickMe(aArguments(0), aArguments(1))
            RaiseEvent OnClickMe(btns(aArguments(0)), aArguments(0))
        Finally
            Dim strRedirect As String = String.Empty
            strRedirect = btns(aArguments(0)).RedirectURL.Trim()
            If strRedirect <> String.Empty Then
                Response.Redirect(strRedirect)
            End If
        End Try
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' Coleção dos botões do toolbar. Permite configurar/acessar
    ''' cada elemento.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>    
    <Description("Coleção dos botões do toolbar. Permite configurar/acessar cada elemento.")> _
    <Browsable(True), _
     PersistenceMode(PersistenceMode.InnerProperty), _
     DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)> _
    Public ReadOnly Property ToolButtons() As ToolButtons
        Get
            Me.EnsureChildControls()
            Return btns
        End Get
    End Property

    ''' <summary>
    ''' Permite configurar o distanciamento entre as bordas e o conteúdo.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Permite configurar o distanciamento entre as bordas e o conteúdo.")> _
    Public Property Padding() As Integer
        Get
            Dim o As Object = ViewState("Padding")
            If IsNothing(o) Then ViewState("Padding") = 5
            Return ViewState("Padding")
        End Get
        Set(ByVal value As Integer)
            ViewState("Padding") = value
        End Set
    End Property

    ''' <summary>
    ''' Permite configurar o distanciamento entre cada elemento.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Permite configurar o distanciamento entre cada elemento.")> _
    Public Property Spacing() As Integer
        Get
            Dim o As Object = ViewState("Spacing")
            If IsNothing(o) Then ViewState("Spacing") = 0
            Return ViewState("Spacing")
        End Get
        Set(ByVal value As Integer)
            ViewState("Spacing") = value
        End Set
    End Property

    ''' <summary>
    ''' Permite configurar as bordas do toolbar.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>    
    <Description("Permite configurar as bordas do toolbar.")> _
    Public Shadows Property BorderStyle() As BorderStyle
        Get
            Dim o As Object = ViewState("BorderStyle")
            If IsNothing(o) Then ViewState("BorderStyle") = BorderStyle.NotSet
            Return ViewState("BorderStyle")
        End Get
        Set(ByVal value As BorderStyle)
            ViewState("BorderStyle") = value
        End Set
    End Property

    ''' <summary>
    ''' Permite configurar as bordas do toolbar.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>    
    <Description("Permite configurar as bordas do toolbar.")> _
    Public Shadows Property BorderWidth() As Unit
        Get
            Dim o As Object = ViewState("BorderWidth")
            If IsNothing(o) Then ViewState("BorderWidth") = Unit.Pixel(1)
            Return ViewState("BorderWidth")
        End Get
        Set(ByVal value As Unit)
            ViewState("BorderWidth") = value
        End Set
    End Property

    ''' <summary>
    ''' Permite configurar a cor das bordas do toolbar.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Permite configurar a cor das bordas do toolbar.")> _
    Public Shadows Property BorderColor() As System.Drawing.Color
        Get
            Dim o As Object = ViewState("BorderColor")
            If IsNothing(o) Then ViewState("BorderColor") = System.Drawing.Color.White
            Return ViewState("BorderColor")
        End Get
        Set(ByVal value As System.Drawing.Color)
            ViewState("BorderColor") = value
        End Set
    End Property

    ''' <summary>
    ''' Se ao redor dos botões haverá bordas.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Se ao redor dos botões haverá bordas.")> _
    Public Property BorderedCells() As Boolean
        Get
            Dim o As Object = ViewState("BorderedCells")
            If IsNothing(o) Then ViewState("BorderedCells") = False
            Return ViewState("BorderedCells")
        End Get
        Set(ByVal value As Boolean)
            ViewState("BorderedCells") = value
        End Set
    End Property

#End Region

#Region "Methods"

    Private Sub RenderMe()
        Dim btn As New ToolButton
        'Dim imgbtn As LinkButton
        Dim imgBtn As String = String.Empty
        Dim tabToolBar As New Table
        Dim cell As TableCell
        Dim row As New TableRow
        '------------------------------
        With tabToolBar
            .ID = "tabToolBar"
            .CellPadding = Padding
            .CellSpacing = Spacing
            .BorderStyle = BorderStyle
            .BorderWidth = BorderWidth
            .BorderColor = BorderColor
        End With
        '------------------------------    
        For i As Integer = 0 To btns.Count - 1
            '------------------------------
            btn = btns(i)
            '------------------------------
            'Configura a celula
            cell = New TableCell
            If BorderedCells Then
                With cell
                    .BorderStyle = BorderStyle
                    .BorderWidth = BorderWidth
                    .BorderColor = BorderColor
                End With
            End If
            ''------------------------------
            'Checa imagem padão
            If btn.ImageUrl = String.Empty Then
                'btn.ImageUrl = "http://10.0.0.238:9090/corpnet/imgBtn/toolbar_empty.gif"
                'btn.DisabledImageUrl = "http://10.0.0.238:9090/corpnet/imgBtn/toolbar_empty.gif"
            Else
                If btn.DisabledImageUrl.Trim = String.Empty Then
                    btn.DisabledImageUrl = btn.ImageUrl
                End If
            End If
            '-------------------------------            
            'Configura o botão
            If btn.Enabled Then                
                imgBtn = "<img src='" & btn.ImageUrl.Replace("~/", "") & "' border=0 style=""cursor:pointer;cursor:hand;"" onclick =""javascript:" & _
                        btn.OnClientClick.Replace(Chr(34), "&quot;") & _
                        IIf(btn.PostBackUrl.Trim <> "", "WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions(&quot;" & Me.ClientID & "$" & btn.ID & "&quot;, &quot;&quot;, false, &quot;&quot;, &quot;" & btn.PostBackUrl.Replace("~/", "") & "&quot;, false, true)); return false;", "") & ";" & _
                        Page.ClientScript.GetPostBackEventReference(Me, i & "$" & btn.ID, True) & """>"
            Else
                imgBtn = "<img src='" & btn.DisabledImageUrl.Replace("~/", "") & "' border=0>"
            End If
            '------------------------------
            'imgbtn = New LinkButton
            'imgbtn.BackColor = Drawing.Color.Transparent            
            '------------------------------
            If btn.ID = String.Empty Then btn.ID = Me.ClientID & "ToolButton" & i.ToString()
            'If btn.ImageUrl = String.Empty Then btn.ImageUrl = "http://10.0.0.238:9090/corpnet/imgBtn/toolbar_empty.gif"
            '------------------------------
            'With imgbtn
            '    .ID = btn.ID
            '    .PostBackUrl = btn.PostBackUrl
            '    .OnClientClick = btn.OnClientClick & Page.ClientScript.GetPostBackEventReference(Me, i & "$" & imgbtn.ID, True)
            '    .CausesValidation = btn.CausesValidation
            '    .CommandArgument = btn.CommandArgument
            '    .CommandName = btn.CommandName
            '    .Enabled = btn.Enabled
            '    If Not btn.Enabled And btn.DisabledImageUrl.Trim <> String.Empty Then
            '        .Text = "<img src='" & btn.DisabledImageUrl.Replace("~/", "") & "' border=0>"
            '    Else
            '        .Text = "<img src='" & btn.ImageUrl.Replace("~/", "") & "' border=0>"
            '    End If
            '    .ToolTip = btn.ToolTip
            '    .ValidationGroup = btn.ValidationGroup
            '------------------------------            
            'cell.Controls.Add(imgBtn)
            cell.Text = imgBtn
            row.Cells.Add(cell)
            '------------------------------
            'End With
        Next
        '------------------------------
        tabToolBar.Rows.Add(row)
        Me.Controls.Add(tabToolBar)
        '------------------------------
    End Sub

#End Region

End Class

<ToolboxItem(False), Serializable()> _
Public Class ToolButtons
    Inherits List(Of ToolButton)

End Class

<ToolboxData("<{0}:ToolButton runat=server></{0}:ToolButton>"), _
ToolboxItem(False), Serializable()> _
Public Class ToolButton    

#Region "Fields"
    Private mToolTip As String
    Private mImageAlign As ImageAlign
    Private mOnClientClick As String
    Private mImageUrl As String
    Private mCausesValidation As Boolean = True
    Private mCommandArgument As String
    Private mCommandName As String
    Private mEnabled As Boolean = True
    Private mPostBackUrl As String
    Private mText As String
    Private mValidationGroup As String
    Private mID As String
    Private mDisabledImageUrl As String
    Private mRedirectURL As String
#End Region

#Region "Properties"

    ''' <summary>
    ''' Sets or retrieves a value that indicates the string that identifies the object. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Sets or retrieves a value that indicates the string that identifies the object.")> _
    Public Property ID() As String
        Get
            Return mID
        End Get
        Set(ByVal value As String)
            mID = value
        End Set
    End Property

    ''' <summary>
    ''' Adds a rich HTML ToolTip control to the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '''     
    <Description("Adds a rich HTML ToolTip control to the document.")> _
    <DefaultValue("")> _
    Public Property ToolTip() As String
        Get
            Return IIf(IsNothing(mToolTip), String.Empty, mToolTip)
        End Get
        Set(ByVal value As String)
            mToolTip = value
        End Set
    End Property

    ''' <summary>
    ''' Specifies the alignment of an image in relation to the text of a Web page.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Specifies the alignment of an image in relation to the text of a Web page.")> _
    <DefaultValue(ImageAlign.NotSet)> _
    Public Property ImageAlign() As ImageAlign
        Get
            Return IIf(IsNothing(mImageAlign), ImageAlign.NotSet, mImageAlign)
        End Get
        Set(ByVal value As ImageAlign)
            mImageAlign = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the client-side script that executes when a Button control's Click event is raised. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Gets or sets the client-side script that executes when a Button control's Click event is raised.")> _
    <DefaultValue("")> _
    Public Property OnClientClick() As String
        Get
            Return IIf(IsNothing(mOnClientClick), String.Empty, mOnClientClick)
        End Get
        Set(ByVal value As String)
            mOnClientClick = value
        End Set
    End Property

    ''' <summary>
    ''' Imagem de exibição do botão.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Imagem de exibição do botão.")> _
    <Browsable(True)> _
    <Bindable(True)> _
    <Editor(GetType(System.Web.UI.Design.ImageUrlEditor), GetType(System.Drawing.Design.UITypeEditor))> _
    <DefaultValue("http://10.0.0.238/corpnet/imgBtn/toolbar_empty.gif")> _
    Public Property ImageUrl() As String
        Get
            Return mImageUrl
        End Get
        Set(ByVal value As String)
            mImageUrl = value
        End Set
    End Property

    ''' <summary>
    ''' Imagem de exibição caso o botão esteja desativado. Caso não seja informado,
    ''' a imagem informada por "ImageUrl" será exibida.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Browsable(True)> _
    <Bindable(True)> _
    <Editor(GetType(System.Web.UI.Design.ImageUrlEditor), GetType(System.Drawing.Design.UITypeEditor))> _
    <Description("Imagem de exibição caso o botão esteja desativado. Caso não seja informado, a imagem informada por ""ImageUrl"" será exibida.")> _
    <DefaultValue("http://10.0.0.238/corpnet/imgBtn/toolbar_empty.gif")> _
    Public Property DisabledImageUrl() As String
        Get
            Return IIf(IsNothing(mDisabledImageUrl), String.Empty, mDisabledImageUrl)
        End Get
        Set(ByVal value As String)
            mDisabledImageUrl = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets a value indicating whether validation is performed when the Button control is clicked. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Gets or sets a value indicating whether validation is performed when the Button control is clicked.")> _
    <DefaultValue(True)> _
    Public Property CausesValidation() As Boolean
        Get
            Return mCausesValidation
        End Get
        Set(ByVal value As Boolean)
            mCausesValidation = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets an optional parameter passed to the Command event along with the associated CommandName.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Gets or sets an optional parameter passed to the Command event along with the associated CommandName.")> _
    <DefaultValue("")> _
    Public Property CommandArgument() As String
        Get
            Return IIf(IsNothing(mCommandArgument), String.Empty, mCommandArgument)
        End Get
        Set(ByVal value As String)
            mCommandArgument = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the command name associated with the Button control that is passed to the Command event.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Gets or sets the command name associated with the Button control that is passed to the Command event")> _
    <DefaultValue("")> _
    Public Property CommandName() As String
        Get
            Return IIf(IsNothing(mCommandName), String.Empty, mCommandName)
        End Get
        Set(ByVal value As String)
            mCommandName = value
        End Set
    End Property


    ''' <summary>
    ''' Gets or sets the URL of the page to post to from the current page when the Button control is clicked. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Gets or sets the URL of the page to post to from the current page when the Button control is clicked.")> _
    <Browsable(True)> _
    <Bindable(True)> _
    <Editor(GetType(System.Web.UI.Design.UrlEditor), GetType(System.Drawing.Design.UITypeEditor))> _
    <DefaultValue("")> _
    Public Property PostBackUrl() As String
        Get
            Return IIf(IsNothing(mPostBackUrl), String.Empty, mPostBackUrl)
        End Get
        Set(ByVal value As String)
            mPostBackUrl = value
        End Set
    End Property

    ''' <summary>
    ''' Texto alternativo caso a imagem de exibição não esteja disponível.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Texto alternativo caso a imagem de exibição não esteja disponível.")> _
    <DefaultValue("")> _
    Public Property Text() As String
        Get
            Return IIf(IsNothing(mText), String.Empty, mText)
        End Get
        Set(ByVal value As String)
            mText = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the group of controls for which the PostBackOptions object causes validation when it posts back to the server.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Gets or sets the group of controls for which the PostBackOptions object causes validation when it posts back to the server.")> _
    <DefaultValue("")> _
    Public Property ValidationGroup() As String
        Get
            Return IIf(IsNothing(mValidationGroup), String.Empty, mValidationGroup)
        End Get
        Set(ByVal value As String)
            mValidationGroup = value
        End Set
    End Property

    ''' <summary>
    ''' Habilita/Desabilita Botão.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Description("Habilita/Desabilita Botão.")> _
    <DefaultValue(True)> _
    Public Property Enabled() As Boolean
        Get
            Return mEnabled
        End Get
        Set(ByVal value As Boolean)
            mEnabled = value
        End Set
    End Property

    <Description("Destino para que, ao clicar, o site seja redirecionado.")> _
    <Browsable(True)> _
    <Bindable(True)> _
    <Editor(GetType(System.Web.UI.Design.UrlEditor), GetType(System.Drawing.Design.UITypeEditor))> _
    <DefaultValue("")> _
    Public Property RedirectURL() As String
        Get
            Return IIf(IsNothing(mRedirectURL), String.Empty, mRedirectURL)
        End Get
        Set(ByVal value As String)
            mRedirectURL = value
        End Set
    End Property
#End Region

End Class