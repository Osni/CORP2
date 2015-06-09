Imports System
Imports System.Text
Imports System.ComponentModel
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Data
Imports System.Collections.Specialized


<DefaultProperty("Text"), ToolboxData("<{0}:ClsGrid runat=server></{0}:ClsGrid>")> _
Public Class ClsGrid
    Inherits WebControl : Implements IPostBackEventHandler
    '------------------------------------------
    Public Event GRID_Click(ByVal GridEvent As String, ByVal GridSendRow As System.Data.DataRow, ByVal iIdGridRow As Integer)
    '------------------------------------------
    Private _ViewState As Collection
    Private _ColsViewState As Collection
    Private Tabela As Table
    Private iCount As Integer
    Private sTableName As String = String.Empty
    Private dtaTabela As System.Data.DataTable
    Private dtcColuna As System.Data.DataColumn
    Private dtcColunas As System.Data.DataColumnCollection
    Private dtrLinha As System.Data.DataRow
    Private i As Integer
    Private l As Long
    Public ExPrt As ClsExtendCols
    '#################################################################################
    Protected Overrides Sub RenderContents(ByVal output As HtmlTextWriter)
        With Me.Controls
            .Clear()
            .Add(ShowTable())
        End With
        MyBase.RenderContents(output)
    End Sub

    Public Sub New()
        dtaTabela = New System.Data.DataTable
        _ColsViewState = New Collection
        ExPrt = New ClsExtendCols
    End Sub
#Region "Grid Eventos"
    '#################################################################################
    Public Sub RaisePostBackEvent(ByVal EventArgument As String) Implements IPostBackEventHandler.RaisePostBackEvent
        Dim sEventPart As Object
        sEventPart = Split(EventArgument, "_")
        Select Case sEventPart(0) & "_" & sEventPart(1)
            Case "GRID_REMOVE"
                RaiseEvent GRID_Click(sEventPart(4) & "_" & sEventPart(3), GRID_GetRowEvent(sEventPart(2)), sEventPart(2))
                GRID_RemoveRowEvent(sEventPart(2))
            Case "GRID_CLICK"
                Select Case sEventPart(3)
                    Case "ExEventImg"
                        GRID_UpdateRow(sEventPart(2))
                    Case "ExEventButton"
                        GRID_UpdateRow(sEventPart(2))
                End Select
                RaiseEvent GRID_Click(sEventPart(4) & "_" & sEventPart(3), GRID_GetRowEvent(sEventPart(2)), sEventPart(2))
        End Select
    End Sub
#End Region
    '#################################################################################
#Region "Grid Property"
    '#################################################################################

    Public Property Coluna(ByVal lngNumRow As Long, ByVal intColInd As Integer) As String
        Get
            Try
                Return dtaTabela.Rows(lngNumRow).Item(intColInd).ToString()
            Catch ex As Exception
                Throw ex
            End Try
        End Get
        Set(ByVal value As String)
            Try
                With dtaTabela.Rows(lngNumRow)
                    .BeginEdit()
                    .Item(intColInd) = value
                    .EndEdit()
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Set
    End Property

    Public Property Coluna(ByVal lngNumRow As Long, ByVal strNameCol As String) As String
        Get
            Return dtaTabela.Rows(lngNumRow).Item(strNameCol).ToString
        End Get
        Set(ByVal value As String)
            Try
                With dtaTabela.Rows(lngNumRow)
                    .BeginEdit()
                    .Item(strNameCol) = value
                    .EndEdit()
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Set
    End Property

    Public ReadOnly Property GetTabela() As System.Data.DataTable
        Get
            Return dtaTabela
        End Get
    End Property
    <Bindable(True), Category("Appearance"), DefaultValue(""), Localizable(True)> Property Text() As String
        Get
            Dim s As String = CStr(ViewState("Text"))
            If s Is Nothing Then
                Return String.Empty
            Else
                Return s
            End If
        End Get

        Set(ByVal Value As String)
            ViewState("Text") = Value
        End Set
    End Property
    <Bindable(True), Category("Behavior"), DefaultValue("Grid1"), Localizable(True)> Public Property TableName()
        Get
            If dtaTabela Is Nothing Then
                Return ""
            Else
                Return dtaTabela.TableName
            End If
        End Get
        Set(ByVal value)
            dtaTabela.TableName = value
        End Set
    End Property
    Public Property bRowsRemove()
        Get
            Dim o As Object = ViewState("bRowsRemove")
            If (IsNothing(o)) Then
                Return False
            Else
                Return o
            End If
        End Get
        Set(ByVal Value As Object)
            AddTableCol("GRID_ID_COL_DEL", "Del")
            ViewState("bRowsRemove") = Value
        End Set
    End Property
    Public Property GridStateView()
        Get
            Dim o As Object = ViewState("GridStateView")
            If (IsNothing(o)) Then
                Return Nothing
            Else
                Return o
            End If
        End Get
        Set(ByVal Value)
            ViewState.Add("GridStateView", Value)
        End Set
    End Property
#End Region
    '#################################################################################
#Region "Grid Method"
    Private Function GRID_UpdateRow(ByVal iIdGridRow As Integer) As Boolean
        Dim iCol As Integer
        Dim col As System.Data.DataColumn
        Dim exp As System.Data.PropertyCollection
        Dim sRt As String
        For iCol = 1 To dtaTabela.Columns.Count - 2
            col = dtaTabela.Columns(iCol)
            '--------------------------------------------------------
            If col.ExtendedProperties.Count > 0 Then
                exp = col.ExtendedProperties
                Select Case CType(exp("TYPE"), System.Int16)
                    Case 1, 4, 5
                        Try
                            With dtaTabela.Rows(iIdGridRow)
                                .BeginEdit()
                                sRt = GRID_GetRequest("grid_" & exp("ID") & "_" & iIdGridRow)
                                .Item(exp("ID")) = sRt
                                .EndEdit()
                            End With
                        Catch ex As Exception
                            Throw New System.Exception(ex.Message)
                        End Try
                End Select
            End If
        Next
        Return True
    End Function
    Private Function GRID_GetRequest(ByVal sId As String) As String
        Dim Request As System.Web.HttpRequest = System.Web.HttpContext.Current.Request
        Dim oReq As Object
        Dim sRet As String = ""
        oReq = Request(sId)
        If Not oReq Is Nothing Then
            sRet = oReq '(sId).ToString()
        End If
        Return sRet
    End Function
    '#################################################################################
    Public Function GRID_RemoveRowEvent(ByVal iIDRow As Integer) As Boolean
        Dim drv As DataRowView
        Try
            drv = dtaTabela.DefaultView(iIDRow)
            drv.Row.Delete()
            Return True
        Catch ex As Exception
            Return False
            Throw New System.Exception(ex.Message)
        End Try
    End Function
    '#################################################################################
    Public Function GRID_SetEditRowEvent(ByVal GridRowEdit As System.Data.DataRow, ByVal iIdGridRow As Integer) As Boolean
        Try
            With dtaTabela.Rows(iIdGridRow)
                .BeginEdit()
                For i = 1 To GridRowEdit.ItemArray.Length - 1
                    .Item(i) = GridRowEdit.Item(i).ToString()
                Next
                .EndEdit()
            End With
            Return True
        Catch ex As Exception
            Return False
            Throw New System.Exception(ex.Message)
        End Try
    End Function
    '#################################################################################
    Public Function GRID_GetEditRowEvent(ByVal iIDRow As Integer) As DataRow
        Dim drv As DataRowView
        Try
            drv = dtaTabela.DefaultView(iIDRow)
            Return drv.Row
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Function
    '#################################################################################
    Public Function GRID_GetRowEvent(ByVal iIDRow As Integer) As DataRow
        Dim drv As DataRowView
        Try
            drv = dtaTabela.DefaultView(iIDRow)
            Return drv.Row
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Function

    '==========================================================================
    ' Verificar se item já foi inserido no grid
    '==========================================================================
    Public Function ItemExists(ByVal strParamOut As String, _
                                ByVal strParamIn As String) As Boolean
        For iCount As Integer = 0 To dtaTabela.Rows.Count - 1
            If dtaTabela.Rows(iCount).Item(strParamIn).ToString.Trim.ToLower = _
               strParamOut.Replace("'", " ").Trim.ToLower Then Return True
        Next
        Return False
    End Function



    '==========================================================================
    ' Remover linhas do grid
    '==========================================================================
    Public Sub ClearGrid()
        While GRID_RemoveRowEvent(0)
        End While
    End Sub

    ''' <summary>Função para criar a tabela principal</summary>
    ''' <param name="sNome">Nome da tabela</param>    
    ''' <returns>Um string de confirmação com o nome da coluna</returns>    
    ''' <remarks>Esta função tem que ser a primeira função a ser chamada </remarks>
    Public Function CriaTabela(Optional ByVal sNome As String = "") As String
        Dim sRet As String = String.Empty
        Try
            '----------------------------------------
            'Inicializa a tabela
            dtaTabela = New DataTable
            '----------------------------------------
            If Trim(sNome) = "" And Trim(sTableName) = "" Then
                Return "Erro: Nome da tabela não definido."
            ElseIf Trim(sNome) = "" Then
                dtaTabela.TableName = sTableName
            Else
                dtaTabela.TableName = sNome
            End If
            '-----------------------------------------
            'Linha Identity                
            sRet = AddTableCol("GRID_ID_COL", "&nbsp;", , "System.Int32")
            If Trim(sRet) = "" Then
                With dtaTabela.Columns("GRID_ID_COL")
                    .Unique = True
                    .ReadOnly = True
                    .AutoIncrement = True
                    .AutoIncrementSeed = 0
                    .AutoIncrementStep = 1
                End With
            Else
                Return sRet
            End If
            '-----------------------------------------
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function AddRowCol(ByVal sColName As String, ByVal sValue As String) As String
        Dim sRet As String = String.Empty
        Try
            dtrLinha.Item(sColName) = sValue
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Sub NewRow()
        Dim sRet As String = String.Empty
        Try
            dtrLinha = dtaTabela.NewRow()
        Catch ex As Exception
        End Try
    End Sub

    Public Function AppendTableRow() As String
        Try
            dtaTabela.Rows.Add(dtrLinha)
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function AddTableCol(ByVal sColName As String, _
                           Optional ByVal sColCaption As String = "", _
                           Optional ByVal lngTamanho As Long = 10, _
                           Optional ByVal sColType As String = "System.String", _
                           Optional ByVal pColExtend As ClsExtendCols = Nothing, _
                           Optional ByVal bUnique As Boolean = False, _
                           Optional ByVal bVisible As Boolean = True) As String

        Dim cl As System.Data.PropertyCollection
        Dim eK As IEnumerator
        Try
            '-----------------------------------------
            dtcColuna = New System.Data.DataColumn
            '-----------------------------------------
            With dtcColuna
                .ColumnName = sColName
                .DataType = Type.GetType(sColType)
                .Unique = bUnique
                dtcColuna.ExtendedProperties.Add("VISIBLE", bVisible)
                If Not pColExtend Is Nothing Then
                    With pColExtend
                        For Each cl In .ExCols
                            eK = cl.Keys.GetEnumerator()
                            eK.Reset()
                            While eK.MoveNext()
                                '-----------------------------------------                                
                                If sColType = "System.Object" And eK.Current.ToString() = "DATA_TABLE" Then
                                    _ColsViewState.Add(cl("DATA_TABLE"), sColName)
                                End If
                                '-----------------------------------------
                                dtcColuna.ExtendedProperties.Add(eK.Current.ToString(), cl(eK.Current))
                            End While
                        Next
                    End With
                End If
                If sColType = "System.String" Then
                    .MaxLength = lngTamanho
                End If
                '-------------------------------------
                If Trim(sColCaption) <> "" Then
                    .Caption = sColCaption
                Else
                    .Caption = sColName
                End If
                '-------------------------------------
            End With
            '-----------------------------------------
            dtaTabela.Columns.Add(dtcColuna)
            dtcColuna.Dispose()
            ExPrt = New ClsExtendCols
            '-----------------------------------------
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function ShowTable() As Object
        Dim sGrid As String = String.Empty
        Dim iCol As Integer
        Dim iRow As Integer
        Dim TableRow As TableRow
        Dim TableCell As TableCell
        Dim col As DataColumn
        Dim sIDCol As String = String.Empty
        Dim sIDColName As String = String.Empty
        Dim sField As String = String.Empty
        Dim exp As System.Data.PropertyCollection
        Dim sComboBox As StringBuilder
        Dim dtTab As System.Data.DataTable
        Dim dtRow As System.Data.DataRow
        Dim sChecked As String
        Dim iDelRow As Integer
        Try
            iDelRow = IIf(bRowsRemove = True, 2, 1)
            '---------------------------------------------------
            Tabela = New Table()
            '---------------------------------------------------
            With Tabela
                .ID = "grid_" & TableName
                .Width = Unit.Percentage(100)
            End With
            '---------------------------------------------------
            If dtaTabela.Columns.Count > 0 Then
                '---------------------------------------------------
                TableRow = New TableRow()
                '---------------------------------------------------                    
                For iCol = 0 To dtaTabela.Columns.Count - 1
                    col = dtaTabela.Columns(iCol)
                    If CBool(col.ExtendedProperties("VISIBLE")) Then
                        TableCell = New TableHeaderCell()
                        TableCell.Text = col.Caption()
                        TableRow.Cells.Add(TableCell)
                    End If
                Next
                '---------------------------------------------------
                Tabela.Rows.Add(TableRow)
                '---------------------------------------------------                    
                Dim strParImpar As String = String.Empty
                For iRow = 0 To dtaTabela.Rows.Count - 1
                    TableRow = New TableRow()
                    '--------------------------------------------------------
                    If (iRow Mod 2) = 0 Then
                        strParImpar = "grdLinhaPar"
                    Else
                        strParImpar = "grdLinhaImpar"
                    End If
                    TableRow.CssClass = strParImpar
                    '--------------------------------------------------------
                    'Primeira Coluna                    
                    'TableCell = New TableCell()
                    TableCell = New TableHeaderCell()
                    TableCell.Text = iRow + 1
                    TableRow.Cells.Add(TableCell)
                    '--------------------------------------------------------
                    For iCol = 1 To dtaTabela.Columns.Count - iDelRow
                        '--------------------------------------------------------
                        sField = dtaTabela.Rows(iRow).Item(iCol).ToString()
                        col = dtaTabela.Columns(iCol)
                        '--------------------------------------------------------
                        'Colunas Extendidas
                        If CBool(col.ExtendedProperties("VISIBLE")) Then
                            If col.ExtendedProperties.Count > 1 Then
                                exp = col.ExtendedProperties
                                TableCell = New TableCell()
                                sIDColName = exp("ID") & "_" & iRow 'Nome do controle
                                Select Case CInt(exp("TYPE"))
                                    Case ClsExtendCols.ExType.TextBox
                                        sIDColName = " name='grid_" & sIDColName & "' id='grid_" & sIDColName & "' "
                                        TableCell.Text = "<input " & exp("ATRIBUTO") & " type=text maxlength='" & exp("MAXLEN") & "' size='" & exp("SIZE") & "' " & sIDColName & "  value='" & sField & "'"">"
                                    Case ClsExtendCols.ExType.CheckBox
                                        sIDColName = " name='grid_" & sIDColName & "' id='grid_" & sIDColName & "' "
                                        If sField = "" Then
                                            sField = "on" : sChecked = ""
                                        Else
                                            If sField = "on" Then
                                                sChecked = "checked"
                                            Else
                                                sChecked = ""
                                            End If
                                        End If
                                        TableCell.Text = "<input " & exp("ATRIBUTO") & " type=checkbox " & sIDColName & " value='" & sField & "' " & sChecked & " >"
                                    Case ClsExtendCols.ExType.ComboBox
                                        '------------------------------------------------------------
                                        sComboBox = New StringBuilder()
                                        sIDColName = " name='grid_" & sIDColName & "' id='grid_" & sIDColName & "' "
                                        sComboBox.Append("<select " & exp("ATRIBUTO") & " " & sIDColName & " >")
                                        '--------------------------------------------
                                        dtTab = CType(_ColsViewState(exp("ID")), Data.DataTable)
                                        '--------------------------------------------
                                        For Each dtRow In dtTab.Rows
                                            If dtRow(0).ToString() = sField Then
                                                sComboBox.Append("<option selected value=" & dtRow(0).ToString() & ">" & dtRow(1).ToString() & "</option>")
                                            Else
                                                sComboBox.Append("<option value=" & dtRow(0).ToString() & ">" & dtRow(1).ToString() & "</option>")
                                            End If
                                        Next
                                        sComboBox.Append("</select>")
                                        TableCell.Text = sComboBox.ToString()
                                        '------------------------------------------------------------
                                    Case ClsExtendCols.ExType.Button
                                        sIDColName = " name='grid_" & sIDColName & "' id='grid_" & sIDColName & "' "
                                        TableCell.Text = "<input " & exp("ATRIBUTO") & " title='" & exp("TOOLTIPTEXT") & "' size='" & exp("SIZE") & "' " & sIDColName & " type=button value=" & exp("LABEL") & " OnClick=""javascript:" & exp("SCRIPT") & " " & Page.ClientScript.GetPostBackEventReference(Me, "GRID_CLICK_" & iRow & "_" & exp("GRID_EVENT") & "_" & exp("ID")) & """>"
                                    Case ClsExtendCols.ExType.Img
                                        sIDColName = " name='grid_" & sIDColName & "' id='grid_" & sIDColName & "' "
                                        TableCell.Text = "<center><a href=#><img " & exp("ATRIBUTO") & " style=""text-align:center;cursor:hand"" " & sIDColName & " title='" & exp("TOOLTIPTEXT") & "' src='" & exp("SRC") & "' border=0 OnClick=""javascript:" & exp("SCRIPT") & " " & Page.ClientScript.GetPostBackEventReference(Me, "GRID_CLICK_" & iRow & "_" & exp("GRID_EVENT") & "_" & exp("ID")) & """></a></center>"
                                End Select
                                TableRow.Cells.Add(TableCell)
                            Else
                                TableCell = New TableCell()
                                If IsNumeric(sField) Then TableCell.HorizontalAlign = HorizontalAlign.Right
                                TableCell.Text = IIf(sField.Trim = "", "&nbsp;", sField)
                                TableRow.Cells.Add(TableCell)
                            End If
                        End If
                    Next
                    'Remover Row
                    If bRowsRemove Then
                        TableCell = New TableCell()
                        TableCell.Text = "<center><a href=#><img style=""cursor:hand"" id=""grid_img_remove"" name=""grid_img_remove"" title=""Remover Item ('" & iRow & "')"" src='GRID_IMAGENS/grid_br.gif' border=0 OnClick=""javascript:" & Page.ClientScript.GetPostBackEventReference(Me, "GRID_REMOVE_" & iRow & "_OnRemover_Remover") & """></a></center>"
                        TableRow.Cells.Add(TableCell)
                    End If
                    '--------------------------------------------------------
                    Tabela.Rows.Add(TableRow)
                Next
                '---------------------------------------------------
                Return Tabela
            Else
                TableRow = New TableRow()
                Tabela.Rows.Add(TableRow)
                TableCell = New TableCell()
                TableCell.Text = Me.ClientID
                TableRow.Cells.Add(TableCell)
                Return Tabela
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
#Region "Control State"
    Protected Overrides Sub OnInit(ByVal e As System.EventArgs)
        Page.RegisterRequiresControlState(Me)
        MyBase.OnInit(e)
    End Sub
    Protected Overrides Function SaveControlState() As Object
        _ViewState = New Collection()
        _ViewState.Add(dtaTabela, "dtaTabela")
        _ViewState.Add(_ColsViewState, "ColsViewState")
        Return Me._ViewState
    End Function
    Protected Overrides Sub LoadControlState(ByVal savedState As Object)
        _ViewState = New Collection()
        _ViewState = CType(savedState, Collection)
        dtaTabela = _ViewState("dtaTabela")
        _ColsViewState = _ViewState("ColsViewState")
    End Sub
#End Region
#End Region
End Class

'##########################################################################################
#Region "Class Extended"
Public Class ClsExtendCols
    Public Enum ExType As Integer
        TextBox = 1
        Button = 2
        Img = 3
        ComboBox = 4
        CheckBox = 5
    End Enum
    Public ExCols As New Collection
    Private prt As System.Data.PropertyCollection
    Public Sub ComboBox(ByVal sID As String, _
                        ByVal dtTab As System.Data.DataTable, _
                        Optional ByVal sATRIBUTO As String = "")
        prt = New System.Data.PropertyCollection()
        With prt
            .Add("ID", sID)
            .Add("DATA_TABLE", dtTab)
            .Add("ATRIBUTO", sATRIBUTO)
            .Add("TYPE", CInt(ExType.ComboBox))
        End With
        ExCols.Add(prt)
    End Sub

    Public Sub ExTextBox(ByVal sID As String, _
                        Optional ByVal sSIZE As Integer = 10, _
                        Optional ByVal sMAXLEN As Integer = 10, _
                        Optional ByVal sTOOLTIPTEXT As String = "", _
                        Optional ByVal sATRIBUTO As String = "")
        prt = New System.Data.PropertyCollection()
        With prt
            .Add("ID", sID)
            .Add("SIZE", sSIZE)
            .Add("MAXLEN", sMAXLEN)
            .Add("TOOLTIPTEXT", sTOOLTIPTEXT)
            .Add("ATRIBUTO", sATRIBUTO)
            .Add("TYPE", CInt(ExType.TextBox))
        End With
        ExCols.Add(prt)
    End Sub
    Public Sub ExCheckBox(ByVal sID As String, _
                        Optional ByVal sLABEL As String = "", _
                        Optional ByVal sTOOLTIPTEXT As String = "", _
                        Optional ByVal sATRIBUTO As String = "")
        prt = New System.Data.PropertyCollection()
        With prt
            .Add("ID", sID)
            .Add("LABEL", sLABEL)
            .Add("TOOLTIPTEXT", sTOOLTIPTEXT)
            .Add("ATRIBUTO", sATRIBUTO)
            .Add("TYPE", CInt(ExType.CheckBox))
        End With
        ExCols.Add(prt)
    End Sub

    Public Sub ExButton(ByVal sID As String, _
                        ByVal sLABEL As String, _
                        Optional ByVal sSIZE As Integer = 10, _
                        Optional ByVal sTOOLTIPTEXT As String = "", _
                        Optional ByVal sATRIBUTO As String = "", _
                        Optional ByVal sSCRIPT As String = "")
        prt = New System.Data.PropertyCollection()
        With prt
            .Add("ID", sID)
            .Add("LABEL", sLABEL)
            .Add("SIZE", sSIZE)
            .Add("TOOLTIPTEXT", sTOOLTIPTEXT)
            .Add("GRID_EVENT", "ExEventButton")
            .Add("ATRIBUTO", sATRIBUTO)
            .Add("SCRIPT", sSCRIPT)
            .Add("TYPE", CInt(ExType.Button))
        End With
        ExCols.Add(prt)
    End Sub
    Public Sub ExImg(ByVal sID As String, _
                        ByVal sSRC As String, _
                        Optional ByVal sTOOLTIPTEXT As String = "", _
                        Optional ByVal sATRIBUTO As String = "", _
                        Optional ByVal sSCRIPT As String = "")
        prt = New System.Data.PropertyCollection()
        With prt
            .Add("ID", sID)
            .Add("SRC", sSRC)
            .Add("TOOLTIPTEXT", sTOOLTIPTEXT)
            .Add("GRID_EVENT", "ExEventImg")
            .Add("ATRIBUTO", sATRIBUTO)
            .Add("SCRIPT", sSCRIPT)
            .Add("TYPE", CInt(ExType.Img))
        End With
        ExCols.Add(prt)
    End Sub
End Class
#End Region





