Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Text
Imports CORP.NET.ClsExcel.CellAlignment
Imports CORP.NET.ClsExcel.CellFont
Imports CORP.NET.ClsExcel.CellHiddenLocked
Imports CORP.NET.ClsExcel.FontFormatting
Imports CORP.NET.ClsExcel.MarginTypes
Imports CORP.NET.ClsExcel.ValueTypes
'==============================================================

Public Class ClsGetExcelFile


#Region "Fields"
    Private strm As Stream

    Private mLimiteLinhas As Integer = 0
    Private mXLSGenerateType As ClsExcel.GenerateType = ClsExcel.GenerateType.ToMemory
    Private mLinhasPorPagina As Short = 0
    Private mDataSource As New DataTable

    Private mFileName As String = String.Empty
    Private mTitulo As String = String.Empty
    Private mTextoCabecalho As String = String.Empty
    Private mTextoRodape As String = String.Empty

    Private mProtegerPlanilha As Boolean = False
    Private mGroupColumn As New List(Of String)
    Private mTitleColumn As New Dictionary(Of String, String)

    Private oTool As New ClsTools

    Public Const FORMATNORMALCELTEXT As Byte = _
                        ClsExcel.CellAlignment.xlsLeftBorder + _
                        ClsExcel.CellAlignment.xlsRightBorder + _
                        ClsExcel.CellAlignment.xlsTopBorder + _
                        ClsExcel.CellAlignment.xlsBottomBorder + _
                        ClsExcel.CellAlignment.xlsLeftAlign

    Public Const FORMATNORMALCELNUMBER As Byte = _
                    ClsExcel.CellAlignment.xlsLeftBorder + _
                    ClsExcel.CellAlignment.xlsRightBorder + _
                    ClsExcel.CellAlignment.xlsTopBorder + _
                    ClsExcel.CellAlignment.xlsBottomBorder + _
                    ClsExcel.CellAlignment.xlsRightAlign


    Public Const FORMATROTULOCEL As Byte = _
                        ClsExcel.CellAlignment.xlsLeftBorder + _
                        ClsExcel.CellAlignment.xlsRightBorder + _
                        ClsExcel.CellAlignment.xlsTopBorder + _
                        ClsExcel.CellAlignment.xlsBottomBorder + _
                        ClsExcel.CellAlignment.xlsCentreAlign

    Private hshDeParaDBTypes As New Hashtable

#End Region


#Region "Properties"

    ''' <summary>
    ''' Permite bloquear a planilha para edição.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ProtegerPlanilha() As Boolean
        Get
            Return mProtegerPlanilha
        End Get
        Set(ByVal value As Boolean)
            mProtegerPlanilha = value
        End Set
    End Property

    ''' <summary>
    ''' Permite definir o nome do arquivo.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FileName() As String
        Get
            If mFileName.Trim = String.Empty Then mFileName = "ARQ_" & Guid.NewGuid.ToString & ".xls"
            Return mFileName
        End Get
        Set(ByVal value As String)
            mFileName = value
        End Set
    End Property


    ''' <summary>
    ''' Permite definir o número de linhas para cada inserção de quebra
    ''' de página. Caso não seja definido nenhum valor, 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LinhasPorPagina() As Short
        Get
            Return mLinhasPorPagina
        End Get
        Set(ByVal value As Short)
            mLinhasPorPagina = value
        End Set
    End Property


    ''' <summary>
    ''' Permite recuperar o fluxo de informações. Caso seja 
    ''' definido a forma de geração para ToFile, será retornado
    ''' um objeto Stream com as características de um FileStream,
    ''' se for ToMemory, um MemoryStream.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetStream() As Stream
        Get
            Return strm
        End Get
    End Property


    ''' <summary>
    ''' Permite definir a forma de geração das informações.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property XLSGenerateType() As ClsExcel.GenerateType
        Get
            Return mXLSGenerateType
        End Get
        Set(ByVal value As ClsExcel.GenerateType)
            mXLSGenerateType = value
        End Set
    End Property


    ''' <summary>
    ''' Permite informar qual a fonte de dados para geração do arquivo.
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property DataSource() As DataTable
        Get
            Return mDataSource
        End Get
        Set(ByVal value As DataTable)
            mDataSource = value
        End Set
    End Property

    ''' <summary>
    ''' Permite adicionar uma linha de título para o arquivo.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Titulo() As String
        Get
            Return mTitulo.Trim
        End Get
        Set(ByVal value As String)
            mTitulo = value
        End Set
    End Property

    ''' <summary>
    ''' Permite definir o texto que será impresso no rodapé.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TextoRodape() As String
        Get
            Return mTextoRodape.Trim
        End Get
        Set(ByVal value As String)
            mTextoCabecalho = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Permite definir o texto que será impresso no cabeçalho.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TextoCabecalho() As String
        Get
            Return mTextoCabecalho
        End Get
        Set(ByVal value As String)
            mTextoCabecalho = value
        End Set
    End Property

#End Region


#Region "Events"
    Public Sub New()
        With hshDeParaDBTypes
            'ValueTypes.xlsInteger
            'ValueTypes.xlsNumber
            'ValueTypes.xlsText            
            .Add("System.Int16", xlsInteger)
            .Add("System.Boolean", xlsInteger)

            .Add("System.Int32", xlsNumber)
            .Add("System.Decimal", xlsNumber)
            .Add("System.DateTime", xlsNumber)
            .Add("System.Double", xlsNumber)
            .Add("System.Single", xlsNumber)

            .Add("System.TimeSpan", xlsText)
            .Add("System.String", xlsText)
        End With
    End Sub
#End Region


#Region "Methods"

    Public Sub GenerateXLS()
        Dim excel As ClsExcel
        Dim hshLinhasQuebra As New Hashtable
        '----------------------------------------------
        'Checa se existem linhas
        If DataSource.Rows.Count = 0 Then Exit Sub
        '----------------------------------------------
        excel = New ClsExcel(FileName, mXLSGenerateType)
        With excel
            '----------------------------------------------
            'specify whether to print the gridlines or not
            'this should come before the setting of fonts and margins
            .PrintGridLines = True
            '----------------------------------------------
            'it is a good idea to set margins, fonts and column widths
            'prior to writing any text/numerics to the spreadsheet. These
            'should come before setting the fonts.
            .SetMargin(xlsTopMargin, 1.5)   'set to 1.5 inches
            .SetMargin(xlsLeftMargin, 1.5)
            .SetMargin(xlsRightMargin, 1.5)
            .SetMargin(xlsBottomMargin, 1.5)
            '----------------------------------------------
            'to insert a Horizontal Page Break you need to specify the row just
            'after where you want the page break to occur. You can insert as many
            'page breaks as you wish (in any order).
            For Each intNumLinha As Integer In RetQuebras()
                .InsertHorizPageBreak(intNumLinha)
            Next
            '----------------------------------------------
            'set a default row height for the entire spreadsheet (1/20th of a point)
            .SetDefaultRowHeight(14)
            '----------------------------------------------
            'Up to 4 fonts can be specified for the spreadsheet. This is a
            'limitation of the Excel 2.1 format. For each value written to the
            'spreadsheet you can specify which font to use.
            '----------------------------------------------
            .SetFont("Verdana", 11, xlsBold)                 'font0   -> Grupo
            .SetFont("Verdana", 9, xlsNoFormat)              'font1    -> Dados
            .SetFont("Verdana", 9, xlsBold)                  'font2   -> Rótulo Colunas
            .SetFont("Verdana", 13, xlsBold)                 'font3   -> Cabeçalho
            '----------------------------------------------
            'Column widths are specified in Excel as 1/256th of a character.
            '----------------------------------------------
            'set any header or footer that you want to print on
            'every page. This text will be centered at the top and/or
            'bottom of each page. The font will always be the font that
            'is specified as font0, therefore you should only set the
            'header/footer after specifying the fonts through SetFont.
            .SetHeader(TextoCabecalho)
            .SetFooter(TextoRodape)
            '------------------------------------------------
            'Lendo as informações do DataSource            
            If mGroupColumn.Count > 0 Then
                Call MakeGroup(excel)
            Else
                Call MakeSimple(excel)
            End If
            '--------------------------------------------
            .ProtectSpreadsheet = ProtegerPlanilha
            '--------------------------------------------
            'Finally, close the spreadsheet
            .SetEOF()
            '--------------------------------------------
            strm = .GetStream
        End With

    End Sub

    Private Sub MakeGroup(ByRef ex As ClsExcel)
        Dim intNumLinha As Integer = 1
        Dim intLinha As Integer = 1
        Dim intColuna As Integer = 0
        Dim strGrupoAnt As String = String.Empty
        Dim strGrupoAtual As String = String.Empty
        Dim intDataType As Integer
        '--------------------------------------------
        With ex
            'Aplica Título e rótulo das colunas
            Call SetTitleGroup(ex, mDataSource.Rows(0), intLinha)
            strGrupoAtual = RetGrupo(mDataSource.Rows(0))
            strGrupoAnt = strGrupoAtual
            '--------------------------------------------
            For Each dtr As DataRow In mDataSource.Rows
                '--------------------------------------------                
                intColuna = 0
                '--------------------------------------------
                strGrupoAtual = RetGrupo(dtr)
                If strGrupoAtual <> strGrupoAnt Then
                    strGrupoAnt = strGrupoAtual
                    Call SetGrupo(ex, dtr, intLinha)
                    intNumLinha += 2
                End If
                '--------------------------------------------
                Dim value As Object
                If mTitleColumn.Count > 0 Then
                    'For Each dtc As DataColumn In mDataSource.Columns
                    For Each ColumnName As String In mTitleColumn.Keys
                        If Not mGroupColumn.Contains(ColumnName.ToLower) And mTitleColumn.ContainsKey(ColumnName.ToLower) Then
                            intColuna += 1

                            value = IIf(IsDBNull(dtr(ColumnName)), "", dtr(ColumnName))
                            intDataType = hshDeParaDBTypes(value.GetType.ToString)

                            If intDataType = 0 Then
                                If value <= 32767 Then
                                    .WriteInteger(xlsFont1, intLinha, intColuna, value, xlsNormal, FORMATNORMALCELNUMBER)
                                Else
                                    .WriteNumber(xlsFont1, intLinha, intColuna, value, , FORMATNORMALCELNUMBER)
                                End If
                            ElseIf intDataType = 1 Then
                                If value.GetType.ToString = "System.DateTime" Then
                                    Dim intFormat As Int16 = 20 'dd/mm/yy hh:mm
                                    If CType(value, System.DateTime).Hour = 0 And _
                                       CType(value, DateTime).Minute = 0 And _
                                       CType(value, DateTime).Second = 0 Then intFormat = 12
                                    .WriteDate(xlsFont1, intLinha, intColuna, value, intFormat, FORMATNORMALCELNUMBER)
                                Else
                                    .WriteNumber(xlsFont1, intLinha, intColuna, value, , FORMATNORMALCELNUMBER)
                                End If
                            Else
                                .WriteText(xlsFont1, intLinha, intColuna, oTool.RemTags(value.ToString), , FORMATNORMALCELTEXT)
                            End If
                        End If
                    Next
                Else
                    For Each dtc As DataColumn In mDataSource.Columns
                        If Not mGroupColumn.Contains(dtc.ColumnName.ToLower) Then
                            intColuna += 1

                            value = IIf(IsDBNull(dtr(dtc.ColumnName)), "", dtr(dtc.ColumnName))
                            intDataType = hshDeParaDBTypes(value.GetType.ToString)

                            If intDataType = 0 Then
                                .WriteInteger(xlsFont1, intLinha, intColuna, value, xlsNormal, FORMATNORMALCELNUMBER)
                            ElseIf intDataType = 1 Then
                                If value.GetType.ToString = "System.DateTime" Then
                                    Dim intFormat As Int16 = 20 'dd/mm/yy hh:mm
                                    If CType(value, System.DateTime).Hour = 0 And _
                                       CType(value, DateTime).Minute = 0 And _
                                       CType(value, DateTime).Second = 0 Then intFormat = 12
                                    .WriteDate(xlsFont1, intLinha, intColuna, value, intFormat, FORMATNORMALCELNUMBER)
                                Else
                                    .WriteNumber(xlsFont1, intLinha, intColuna, value, , FORMATNORMALCELNUMBER)
                                End If
                            Else
                                .WriteText(xlsFont1, intLinha, intColuna, oTool.RemTags(value.ToString), , FORMATNORMALCELTEXT)
                            End If

                        End If
                    Next
                End If
                '--------------------------------------------
                intNumLinha += 1
                intLinha += 1
                '--------------------------------------------
                If intNumLinha > LinhasPorPagina And LinhasPorPagina > 0 Then
                    Call SetTitleGroup(ex, dtr, intLinha)
                    If Titulo <> "" Then intNumLinha = 3 Else intNumLinha = 2
                End If
                '--------------------------------------------
            Next
            '--------------------------------------------
        End With
    End Sub

    Private Sub SetGrupo(ByRef ex As ClsExcel, ByRef dtr As DataRow, ByRef lRow As Integer)
        Dim strTextoGrupo As String = String.Empty
        Dim strTitleColumn As String = String.Empty
        If mGroupColumn.Count > 0 Then
            For Each strColName As String In mGroupColumn
                If mTitleColumn.ContainsKey(strColName) Then
                    strTitleColumn = IIf(mTitleColumn(strColName).Trim <> String.Empty, mTitleColumn(strColName) & ": ", "")
                    strTextoGrupo &= IIf(strTextoGrupo.Trim <> String.Empty, " - ", "") & strTitleColumn & dtr(strColName)
                Else
                    strTextoGrupo &= IIf(strTextoGrupo.Trim <> String.Empty, " - ", "") & dtr(strColName)
                End If
            Next
            ex.WriteValue(xlsText, xlsFont0, FORMATNORMALCELTEXT, xlsNormal, lRow, 1, strTextoGrupo)
        End If
        '--------------------------------------------
        lRow += 1
        '--------------------------------------------
        Call SetTitleColumns(ex, lRow)
    End Sub

    Private Sub MakeSimple(ByRef ex As ClsExcel)
        Dim intNumLinha As Integer = 1
        Dim intLinha As Integer = 1
        Dim intColuna As Integer = 0
        Dim intDataType As Integer
        '--------------------------------------------
        With ex
            'Aplica Título e rótulo das colunas
            Call SetTitle(ex, mDataSource.Rows(0), intLinha)
            If Titulo <> String.Empty Then intNumLinha += 1
            '--------------------------------------------
            For Each dtr As DataRow In mDataSource.Rows
                '--------------------------------------------                
                intColuna = 0
                '--------------------------------------------
                Dim value As Object
                If mTitleColumn.Count > 0 Then

                    For Each ColumnName As String In mTitleColumn.Keys
                        intColuna += 1
                        value = IIf(IsDBNull(dtr(ColumnName)), "", dtr(ColumnName))
                        intDataType = hshDeParaDBTypes(value.GetType.ToString)

                        If intDataType = 0 Then
                            .WriteInteger(xlsFont1, intLinha, intColuna, value, xlsNormal, FORMATNORMALCELNUMBER)
                        ElseIf intDataType = 1 Then
                            If value.GetType.ToString = "System.DateTime" Then
                                Dim intFormat As Int16 = 20 'dd/mm/yy hh:mm
                                If CType(value, System.DateTime).Hour = 0 And _
                                   CType(value, DateTime).Minute = 0 And _
                                   CType(value, DateTime).Second = 0 Then intFormat = 12
                                .WriteDate(xlsFont1, intLinha, intColuna, value, intFormat, FORMATNORMALCELNUMBER)
                            Else
                                .WriteNumber(xlsFont1, intLinha, intColuna, value, , FORMATNORMALCELNUMBER)
                            End If
                        Else
                            .WriteText(xlsFont1, intLinha, intColuna, oTool.RemTags(value.ToString), , FORMATNORMALCELTEXT)
                        End If
                    Next
                Else
                    For Each dtc As DataColumn In mDataSource.Columns
                        intColuna += 1
                        value = IIf(IsDBNull(dtr(dtc.ColumnName)), "", dtr(dtc.ColumnName))
                        intDataType = hshDeParaDBTypes(value.GetType.ToString)

                        If intDataType = 0 Then
                            .WriteInteger(xlsFont1, intLinha, intColuna, value, xlsNormal, FORMATNORMALCELNUMBER)
                        ElseIf intDataType = 1 Then
                            If value.GetType.ToString = "System.DateTime" Or value.GetType.ToString = "System.TimeSpan" Then
                                Dim intFormat As Int16 = 20 'dd/mm/yy hh:mm
                                If CType(value, System.DateTime).Hour = 0 And _
                                   CType(value, DateTime).Minute = 0 And _
                                   CType(value, DateTime).Second = 0 Then intFormat = 12
                                .WriteDate(xlsFont1, intLinha, intColuna, value, intFormat, FORMATNORMALCELNUMBER)
                            Else
                                .WriteNumber(xlsFont1, intLinha, intColuna, value, , FORMATNORMALCELNUMBER)
                            End If
                        Else
                            .WriteText(xlsFont1, intLinha, intColuna, oTool.RemTags(value.ToString), , FORMATNORMALCELTEXT)
                        End If
                    Next
                End If
                '--------------------------------------------
                intNumLinha += 1
                intLinha += 1
                '--------------------------------------------
                If intNumLinha = mLimiteLinhas And LinhasPorPagina > 0 Then
                    Call SetTitle(ex, dtr, intLinha)
                    If Titulo <> String.Empty Then intNumLinha = 2 Else intNumLinha = 1
                End If
                '--------------------------------------------
            Next
            '--------------------------------------------
        End With
    End Sub

    Private Sub SetTitleGroup(ByRef ex As ClsExcel, ByRef dtr As DataRow, ByRef lRow As Integer)
        '------------------------------------------------------------        
        'Gravando o título
        If Titulo <> String.Empty Then
            ex.WriteValue(xlsText, xlsFont3, xlsLeftAlign, xlsNormal, lRow, 1, Titulo)
            lRow += 1
        End If
        '------------------------------------------------------------
        Call SetGrupo(ex, dtr, lRow)
        '------------------------------------------------------------
    End Sub

    Private Sub SetTitleColumns(ByRef ex As ClsExcel, ByRef lRow As Integer)
        Dim lCol As Integer = 0
        'Gravando rótulo das colunas
        If mGroupColumn.Count > 0 Then
            If mTitleColumn.Count > 0 Then
                For Each ColumnName As String In mTitleColumn.Keys
                    If Not mGroupColumn.Contains(ColumnName) Then
                        lCol += 1
                        ex.WriteValue(xlsText, xlsFont2, FORMATROTULOCEL, xlsNormal, lRow, lCol, mTitleColumn(ColumnName))
                    End If
                Next
            Else
                For Each dtc As DataColumn In mDataSource.Columns
                    If Not mGroupColumn.Contains(dtc.ColumnName.ToLower) Then
                        lCol += 1
                        ex.WriteValue(xlsText, xlsFont2, FORMATROTULOCEL, xlsNormal, lRow, lCol, dtc.ColumnName)
                    End If
                Next
            End If
        Else
            If mTitleColumn.Count > 0 Then
                For Each ColumnName As String In mTitleColumn.Keys
                    If mTitleColumn.ContainsKey(ColumnName.ToLower) Then
                        lCol += 1
                        ex.WriteValue(xlsText, xlsFont2, FORMATROTULOCEL, xlsNormal, lRow, lCol, mTitleColumn(ColumnName))
                    End If
                Next
            Else
                For Each dtc As DataColumn In mDataSource.Columns
                    lCol += 1
                    ex.WriteValue(xlsText, xlsFont2, FORMATROTULOCEL, xlsNormal, lRow, lCol, dtc.ColumnName)
                Next
            End If
        End If
        '------------------------------------------------------------
        lRow += 1
        '------------------------------------------------------------
    End Sub

    Private Sub SetTitle(ByRef ex As ClsExcel, ByRef dtr As DataRow, ByRef lRow As Integer)
        '------------------------------------------------------------        
        'Gravando o título
        If Titulo <> String.Empty Then
            ex.WriteValue(xlsText, xlsFont3, xlsLeftAlign, xlsNormal, lRow, 1, Titulo)
            lRow += 1
        End If
        '------------------------------------------------------------
        'Gravando rótulo das colunas
        Call SetTitleColumns(ex, lRow)
        '------------------------------------------------------------
    End Sub

    Public Sub AddGroupColumn(ByVal strColName As String)
        mGroupColumn.Add(strColName.ToLower)
    End Sub

    Public Sub AddColumnTitle(ByVal strColName As String, ByVal strTitle As String)
        mTitleColumn.Add(strColName.ToLower, strTitle.Replace("&nbsp;", " "))
    End Sub

    Private Function RetQuebras() As List(Of Integer)
        Dim Quebras As New List(Of Integer)
        Dim strGrupoAnt As String = String.Empty
        Dim strGrupoAtual As String = String.Empty
        Dim intLinhasLidas As Integer = 0
        Dim intNumLinha As Integer = 0
        '--------------------------------
        'Define o número de quebras
        If LinhasPorPagina > 0 Then
            '================================
            'Definindo o máximo de linhas/página                                    
            '--------------------------------
            mLimiteLinhas = LinhasPorPagina
            '--------------------------------
            'Titulo
            If Titulo <> String.Empty Then
                mLimiteLinhas += 1
                intNumLinha += 1
                intLinhasLidas += 1
            End If
            '--------------------------------
            If mGroupColumn.Count = 0 Then
                intNumLinha += 1
                intLinhasLidas += 1
            End If
            'Rótulo das colunas
            mLimiteLinhas += 1
            '================================
            'Definindo quebras de página
            For i As Integer = 0 To mDataSource.Rows.Count - 1
                '--------------------------------
                If mGroupColumn.Count > 0 Then
                    strGrupoAtual = RetGrupo(mDataSource.Rows(i))
                    If strGrupoAnt <> strGrupoAtual Then
                        intLinhasLidas += 2
                        intNumLinha += 2
                        strGrupoAnt = strGrupoAtual
                    End If
                End If
                '--------------------------------
                intLinhasLidas += 1
                intNumLinha += 1
                '--------------------------------
                If intLinhasLidas = mLimiteLinhas Then
                    Quebras.Add(intNumLinha + 1)
                    intLinhasLidas = 0
                    strGrupoAnt = ""
                    If Titulo <> String.Empty Then
                        intNumLinha += 1
                        intLinhasLidas += 1
                    End If
                    If mGroupColumn.Count = 0 Then
                        intNumLinha += 1
                        intLinhasLidas += 1
                    End If
                End If
                '--------------------------------
            Next
            '================================
        End If
        Return Quebras
    End Function

    Private Function RetNumGrupos() As Integer
        Dim strGrupoAnt As String = String.Empty
        Dim strGrupoAtual As String = String.Empty
        Dim intNumGrupos As Integer = 0
        '----------------------------------------
        For Each dtr As DataRow In mDataSource.Rows
            strGrupoAtual = RetGrupo(dtr)
            If strGrupoAnt <> strGrupoAtual Then intNumGrupos += 1
        Next
        '----------------------------------------
    End Function

    Private Function RetGrupo(ByRef dtr As DataRow) As String
        Dim strValorGrupo As String = String.Empty
        For Each s As String In mGroupColumn
            strValorGrupo &= dtr(s)
        Next
        Return strValorGrupo
    End Function

#End Region


End Class

'==============================================================
'==============================================================

Public Class ClsExcel

#Region "Fields"
    Private XLSGenerateType As GenerateType = GenerateType.ToFile
    'Private fs As MemoryStream 'FileStream
    Private strm As Stream
    Private writer As BinaryWriter
    Private strFileName As String = "arq.xls"
    Private BEG_FILE_MARKER As BEG_FILE_RECORD
    Private END_FILE_MARKER As END_FILE_RECORD
    Private HORIZ_PAGE_BREAK As HPAGE_BREAK_RECORD
    'create an array that will hold the rows where a horizontal page
    'break will be inserted just before.
    Private HorizPageBreakRows() As Int16
    Private NumHorizPageBreaks As Int16
#End Region


#Region "Enumerations"

    'enum to handle the various types of values that can be written
    'to the excel file.
    Public Enum ValueTypes
        xlsInteger = 0
        xlsNumber = 1
        xlsText = 2
    End Enum

    'enum to hold cell alignment
    Public Enum CellAlignment
        xlsGeneralAlign = 0
        xlsLeftAlign = 1
        xlsCentreAlign = 2
        xlsRightAlign = 3
        xlsFillCell = 4
        xlsLeftBorder = 8
        xlsRightBorder = 16
        xlsTopBorder = 32
        xlsBottomBorder = 64
        xlsShaded = 128
    End Enum

    'enum to handle selecting the font for the cell
    Public Enum CellFont
        'used by rgbAttr2
        'bits 0-5 handle the *picture* formatting, not bold/underline etc...
        'bits 6-7 handle the font number
        xlsFont0 = 0
        xlsFont1 = 64
        xlsFont2 = 128
        xlsFont3 = 192
    End Enum

    Public Enum CellHiddenLocked
        'used by rgbAttr1
        'bits 0-5 must be zero
        'bit 6 locked/unlocked
        'bit 7 hidden/not hidden
        xlsNormal = 0
        xlsLocked = 64
        xlsHidden = 128
    End Enum

    'set up variables to hold the spreadsheet's layout
    Public Enum MarginTypes
        xlsLeftMargin = 38
        xlsRightMargin = 39
        xlsTopMargin = 40
        xlsBottomMargin = 41
    End Enum

    Public Enum FontFormatting
        'add these enums together. For example: xlsBold + xlsUnderline
        xlsNoFormat = 0
        xlsBold = 1
        xlsItalic = 2
        xlsUnderline = 4
        xlsStrikeout = 8
    End Enum

    Public Enum GenerateType
        ToMemory
        ToFile
    End Enum

#End Region


#Region "Structures"

    Private Structure FONT_RECORD
        Dim opcode As Int16  '49
        Dim length As Int16  '5+len(fontname)
        Dim FontHeight As Int16
        'bit0 bold, bit1 italic, bit2 underline, bit3 strikeout, bit4-7 reserved
        Dim FontAttributes1 As Byte
        Dim FontAttributes2 As Byte  'reserved - always 0
        Dim FontNameLength As Byte
    End Structure

    Private Structure PASSWORD_RECORD
        Dim opcode As Int16  '47
        Dim length As Int16  'len(password)
    End Structure

    Private Structure HEADER_FOOTER_RECORD
        Dim opcode As Int16  '20 Header, 21 Footer
        Dim length As Int16  '1+len(text)
        Dim TextLength As Byte
    End Structure

    Private Structure PROTECT_SPREADSHEET_RECORD
        Dim opcode As Int16  '18
        Dim length As Int16  '2
        Dim Protect As Int16
    End Structure

    Private Structure FORMAT_COUNT_RECORD
        Dim opcode As Int16  '1f
        Dim length As Int16 '2
        Dim Count As Int16
    End Structure

    Private Structure FORMAT_RECORD
        Dim opcode As Int16  '1e
        Dim length As Int16  '1+len(format)
        Dim FormatLenght As Byte 'len(format)
    End Structure '+ followed by the Format-Picture

    Private Structure COLWIDTH_RECORD
        Dim opcode As Int16  '36
        Dim length As Int16  '4
        Dim col1 As Byte       'first column
        Dim col2 As Byte       'last column
        Dim ColumnWidth As Int16   'at 1/256th of a character
    End Structure

    'Beginning Of File record
    Private Structure BEG_FILE_RECORD
        Dim opcode As Int16
        Dim length As Int16
        Dim version As Int16
        Dim ftype As Int16
    End Structure

    'End Of File record
    Private Structure END_FILE_RECORD
        Dim opcode As Int16
        Dim length As Int16
    End Structure

    'true/false to print gridlines
    Private Structure PRINT_GRIDLINES_RECORD
        Dim opcode As Int16
        Dim length As Int16
        Dim PrintFlag As Int16
    End Structure

    'Integer record
    Private Structure tInteger
        Dim opcode As Int16
        Dim length As Int16
        Dim Row As Int16     'unsigned integer
        Dim col As Int16
        'rgbAttr1 handles whether cell is hidden and/or locked
        Dim rgbAttr1 As Byte
        'rgbAttr2 handles the Font# and Formatting assigned to this cell
        Dim rgbAttr2 As Byte
        'rgbAttr3 handles the Cell Alignment/borders/shading
        Dim rgbAttr3 As Byte
        Dim intValue As Int16  'the actual integer value
    End Structure

    'Number record
    Private Structure tNumber
        Dim opcode As Int16
        Dim length As Int16
        Dim Row As Int16
        Dim col As Int16
        Dim rgbAttr1 As Byte
        Dim rgbAttr2 As Byte
        Dim rgbAttr3 As Byte
        Dim NumberValue As Double  '8 Bytes
    End Structure

    'Label (Text) record
    Private Structure tText
        Dim opcode As Int16
        Dim length As Int16
        Dim Row As Int16
        Dim col As Int16
        Dim rgbAttr1 As Byte
        Dim rgbAttr2 As Byte
        Dim rgbAttr3 As Byte
        Dim TextLength As Byte
    End Structure

    Private Structure MARGIN_RECORD_LAYOUT
        Dim opcode As Int16
        Dim length As Int16
        Dim MarginValue As Double  '8 bytes
    End Structure

    Private Structure HPAGE_BREAK_RECORD
        Dim opcode As Int16
        Dim length As Int16
        Dim NumPageBreaks As Int16
    End Structure

    Private Structure DEF_ROWHEIGHT_RECORD
        Dim opcode As Int16
        Dim length As Int16
        Dim RowHeight As Int16
    End Structure

    Private Structure ROW_HEIGHT_RECORD
        Dim opcode As Int16  '08
        Dim length As Int16  'should always be 16 bytes
        Dim RowNumber As Int16
        Dim FirstColumn As Int16
        Dim LastColumn As Int16
        Dim RowHeight As Int16  'written to file as 1/20ths of a point
        Dim internal As Int16
        Dim DefaultAttributes As Byte  'set to zero for no default attributes
        Dim FileOffset As Int16
        Dim rgbAttr1 As Byte
        Dim rgbAttr2 As Byte
        Dim rgbAttr3 As Byte
    End Structure


#End Region


#Region "Events"

    Public Sub New(Optional ByVal pXLSGenerateType As GenerateType = GenerateType.ToFile)
        'Set up default values for records
        'These should be the values that 
        'are the same for every record of these types
        With BEG_FILE_MARKER  'beginning of file
            .opcode = 9
            .length = 4
            .version = 2
            .ftype = 10
        End With

        With END_FILE_MARKER  'end of file marker            
            .opcode = 10
        End With

        XLSGenerateType = pXLSGenerateType
        strFileName = "ARQ_" & Guid.NewGuid.ToString & ".xls"

        Call CreateFile(strFileName)
    End Sub

    Public Sub New(ByVal _FileName As String, Optional ByVal pXLSGenerateType As GenerateType = GenerateType.ToFile)
        'Set up default values for records
        'These should be the values that 
        'are the same for every record of these types
        With BEG_FILE_MARKER  'beginning of file
            .opcode = 9
            .length = 4
            .version = 2
            .ftype = 10
        End With

        With END_FILE_MARKER  'end of file marker            
            .opcode = 10
        End With

        XLSGenerateType = pXLSGenerateType

        If _FileName <> String.Empty Then
            strFileName = _FileName
        Else
            strFileName = "ARQ_" & Guid.NewGuid.ToString & ".xls"
        End If

        Call CreateFile(strFileName)
    End Sub

#End Region


#Region "Properties"

    Public ReadOnly Property GetStream() As Stream
        Get
            Return strm
        End Get
    End Property

    Public Property FileName() As String
        Get
            Return strFileName
        End Get
        Set(ByVal value As String)
            strFileName = value
        End Set
    End Property

    Public WriteOnly Property PrintGridLines() As Boolean
        Set(ByVal value As Boolean)
            'Try
            Dim GRIDLINES_RECORD As PRINT_GRIDLINES_RECORD
            With GRIDLINES_RECORD
                .opcode = 43
                .length = 2
                .PrintFlag = IIf(value, 1, 0)
                writer.Write(.opcode)
                writer.Write(.length)
                writer.Write(.PrintFlag)
            End With
            'Catch : End Try
        End Set
    End Property

    Public WriteOnly Property ProtectSpreadsheet() As Boolean
        Set(ByVal value As Boolean)
            'Try
            Dim PROTECT_RECORD As PROTECT_SPREADSHEET_RECORD

            With PROTECT_RECORD
                .opcode = 18
                .length = 2
                .Protect = IIf(value, 1, 0)
            End With
            writer.Write(PROTECT_RECORD.opcode)
            writer.Write(PROTECT_RECORD.length)
            writer.Write(PROTECT_RECORD.Protect)
            'Catch : End Try
        End Set
    End Property

#End Region


#Region "Methods"

    Private Function CreateFile(ByVal FileName As String) As Boolean
        '        Try
        If XLSGenerateType = GenerateType.ToFile Then
            'Apaga o arquivo, caso já exista        
            If File.Exists(FileName) Then File.Delete(FileName)
            strm = New FileStream(FileName, FileMode.CreateNew, FileAccess.Write, FileShare.Read)
        Else
            strm = New MemoryStream()
        End If

        writer = New BinaryWriter(strm)

        With BEG_FILE_MARKER
            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.version)
            writer.Write(.ftype)
        End With

        Call WriteDefaultFormats()

        'create the Horizontal Page Break array
        ReDim HorizPageBreakRows(0)
        NumHorizPageBreaks = 0

        Return True

        'Catch ex As Exception
        'Return False
        'End Try
    End Function

    Public Function WriteDefaultFormats() As Int16
        Dim cFORMAT_COUNT_RECORD As FORMAT_COUNT_RECORD
        Dim cFORMAT_RECORD As FORMAT_RECORD

        Dim lIndex As Integer
        Dim l As Integer

        Dim aFormat(0 To 23) As String
        Dim q As String


        q = Chr(34)

        aFormat(0) = "General"
        aFormat(1) = "0"
        aFormat(2) = "0.00"
        aFormat(3) = "#,##0"
        aFormat(4) = "#,##0.00"
        aFormat(5) = "#,##0\ " & q & "$" & q & ";\-#,##0\ " & q & "$" & q
        aFormat(6) = "#,##0\ " & q & "$" & q & ";[Red]\-#,##0\ " & q & "$" & q
        aFormat(7) = "#,##0.00\ " & q & "$" & q & ";\-#,##0.00\ " & q & "$" & q
        aFormat(8) = "#,##0.00\ " & q & "$" & q & ";[Red]\-#,##0.00\ " & q & "$" & q
        aFormat(9) = "0%"
        aFormat(10) = "0.00%"
        aFormat(11) = "0.00E+00"
        aFormat(12) = "dd/mm/yy"
        aFormat(13) = "dd/\ mmm\ yy"
        aFormat(14) = "dd/\ mmm"
        aFormat(15) = "mmm\ yy"
        aFormat(16) = "h:mm\ AM/PM"
        aFormat(17) = "h:mm:ss\ AM/PM"
        aFormat(18) = "hh:mm"
        aFormat(19) = "hh:mm:ss"
        aFormat(20) = "dd/mm/yy\ hh:mm"
        aFormat(21) = "##0.0E+0"
        aFormat(22) = "mm:ss"
        aFormat(23) = "@"

        With cFORMAT_COUNT_RECORD
            .opcode = &H1F
            .length = &H2
            .Count = UBound(aFormat)

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.Count)
        End With

        For lIndex = LBound(aFormat) To UBound(aFormat)
            l = Len(aFormat(lIndex))
            With cFORMAT_RECORD
                .opcode = &H1E
                .length = CInt(l + 1)
                .FormatLenght = l
                writer.Write(.opcode)
                writer.Write(.length)
                writer.Write(.FormatLenght)
            End With

            'Then the actual format
            Dim b As Byte, a As Integer
            For a = 1 To l
                b = Asc(Mid$(aFormat(lIndex), a, 1))
                writer.Write(b)
            Next
        Next lIndex

        Exit Function

    End Function

    Public Function SetEOF() As Boolean
        'Try

        'write the horizontal page breaks if necessary
        If NumHorizPageBreaks > 0 Then
            'the Horizontal Page Break array must be in sorted order.
            'Use a simple Bubble sort because the size of this array would
            'be pretty small most of the time. A QuickSort would probably
            'be overkill.
            Dim lLoop1 As Integer
            Dim lLoop2 As Integer
            Dim lTemp As Integer

            For lLoop1 = UBound(HorizPageBreakRows) To LBound(HorizPageBreakRows) Step -1
                For lLoop2 = LBound(HorizPageBreakRows) + 1 To lLoop1
                    If HorizPageBreakRows(lLoop2 - 1) > HorizPageBreakRows(lLoop2) Then
                        lTemp = HorizPageBreakRows(lLoop2 - 1)
                        HorizPageBreakRows(lLoop2 - 1) = HorizPageBreakRows(lLoop2)
                        HorizPageBreakRows(lLoop2) = lTemp
                    End If
                Next lLoop2
            Next lLoop1

            'write the Horizontal Page Break Record
            With HORIZ_PAGE_BREAK
                .opcode = 27
                .length = 2 + (NumHorizPageBreaks * 2)
                .NumPageBreaks = NumHorizPageBreaks

                writer.Write(.opcode)
                writer.Write(.length)
                writer.Write(.NumPageBreaks)
            End With

            'now write the actual page break values
            'the MKI function is standard in other versions of BASIC but
            'VisualBasic does not have it. A KnowledgeBase article explains
            'how to recreate it (albeit using 16-bit API, I switched it
            'to 32-bit).
            For x As Integer = 1 To UBound(HorizPageBreakRows)
                writer.Write(CType(HorizPageBreakRows(x), System.Int16))
                'writer.Write(MKI(HorizPageBreakRows(x)))
            Next
        End If

        With END_FILE_MARKER
            writer.Write(.opcode)
            writer.Write(.length)
        End With

        Return True

        'Catch ex As Exception
        '
        '       Return False
        '      End Try
    End Function

    Public Function InsertHorizPageBreak(ByVal lrow As Integer) As Boolean
        Dim Row As Integer
        'Try
        'the row and column values are written to the excel file as
        'unsigned integers. Therefore, must convert the longs to integer.
        If lrow > 32767 Then
            Row = CInt(lrow - 65536)
        Else
            Row = CInt(lrow) - 1    'rows/cols in Excel binary file are zero based
        End If

        NumHorizPageBreaks = NumHorizPageBreaks + 1
        ReDim Preserve HorizPageBreakRows(NumHorizPageBreaks)

        HorizPageBreakRows(NumHorizPageBreaks) = Row

        Return True

        'Catch ex As Exception
        'Return False
        'End Try
    End Function

    Public Function WriteValue(ByVal ValueType As ValueTypes, _
                               ByVal CellFontUsed As CellFont, _
                               ByVal Alignment As CellAlignment, _
                               ByVal HiddenLocked As CellHiddenLocked, _
                               ByVal lRow As Integer, _
                               ByVal lCol As Integer, _
                               ByVal value As Object, _
                               Optional ByVal CellFormat As Integer = 0) As Boolean

        'Try
        Dim Row As Integer
        Dim Col As Integer
        'the row and column values are written to the excel file as
        'unsigned integers. Therefore, must convert the longs to integer.


        If lRow > 32767 Then
            Row = CInt(lRow - 65536)
        Else
            Row = CInt(lRow) - 1    'rows/cols in Excel binary file are zero based
        End If

        If lCol > 32767 Then
            Col = CInt(lCol - 65536)
        Else
            Col = CInt(lCol) - 1    'rows/cols in Excel binary file are zero based
        End If

        Select Case ValueType
            Case ValueTypes.xlsInteger
                Dim INTEGER_RECORD As tInteger
                With INTEGER_RECORD
                    .opcode = 2
                    .length = 9
                    .Row = Row
                    .col = Col
                    .rgbAttr1 = CByte(HiddenLocked)
                    .rgbAttr2 = CByte(CellFontUsed + CellFormat)
                    .rgbAttr3 = CByte(Alignment)
                    .intValue = CInt(value)

                    writer.Write(.opcode)
                    writer.Write(.length)
                    writer.Write(.Row)
                    writer.Write(.col)
                    writer.Write(.rgbAttr1)
                    writer.Write(.rgbAttr2)
                    writer.Write(.rgbAttr3)
                    writer.Write(.intValue)
                End With

            Case ValueTypes.xlsNumber
                Dim NUMBER_RECORD As tNumber
                With NUMBER_RECORD
                    .opcode = 3
                    .length = 15
                    .Row = Row
                    .col = Col
                    .rgbAttr1 = CByte(HiddenLocked)
                    .rgbAttr2 = CByte(CellFontUsed + CellFormat)
                    .rgbAttr3 = CByte(Alignment)
                    If value.GetType().ToString = "System.DateTime" Then
                        .NumberValue = value.ToOAdate
                    Else
                        .NumberValue = value
                    End If
                    writer.Write(.opcode)
                    writer.Write(.length)
                    writer.Write(.Row)
                    writer.Write(.col)
                    writer.Write(.rgbAttr1)
                    writer.Write(.rgbAttr2)
                    writer.Write(.rgbAttr3)
                    writer.Write(.NumberValue)
                End With

            Case ValueTypes.xlsText
                Dim b As Byte
                Dim lngStringLength As Integer
                Dim strValor As String
                Dim TEXT_RECORD As tText

                strValor = CStr("" & value)
                lngStringLength = Len(strValor)

                With TEXT_RECORD
                    .opcode = 4
                    .length = 10
                    'Length of the text portion of the record
                    .TextLength = lngStringLength
                    'Total length of the record
                    .length = 8 + lngStringLength
                    .Row = Row
                    .col = Col
                    .rgbAttr1 = CByte(HiddenLocked)
                    .rgbAttr2 = CByte(CellFontUsed + CellFormat)
                    .rgbAttr3 = CByte(Alignment)
                    'Put record header
                    writer.Write(.opcode)
                    writer.Write(.length)
                    writer.Write(.Row)
                    writer.Write(.col)
                    writer.Write(.rgbAttr1)
                    writer.Write(.rgbAttr2)
                    writer.Write(.rgbAttr3)
                    writer.Write(.TextLength)
                    'Then the actual string data
                    For a As Integer = 1 To lngStringLength
                        b = Asc(Mid(strValor, a, 1))
                        writer.Write(b)
                    Next
                End With
        End Select

        Return True

        'Catch ex As Exception

        'Return False
        'End Try
    End Function

    Public Function WriteText(ByVal CellFontUsed As CellFont, _
                               ByVal lRow As Integer, _
                               ByVal lCol As Integer, _
                               ByVal value As Object, _
                               Optional ByVal CellFormat As Integer = 0, _
                               Optional ByVal Alignment As CellAlignment = xlsLeftAlign) As Boolean

        'Try
        Dim Row As Integer
        Dim Col As Integer
        Dim b As Byte
        Dim lngStringLength As Integer
        Dim strValor As String
        Dim TEXT_RECORD As tText

        If lRow > 32767 Then
            Row = CInt(lRow - 65536)
        Else
            Row = CInt(lRow) - 1    'rows/cols in Excel binary file are zero based
        End If

        If lCol > 32767 Then
            Col = CInt(lCol - 65536)
        Else
            Col = CInt(lCol) - 1    'rows/cols in Excel binary file are zero based
        End If

        strValor = CStr("" & value)
        lngStringLength = Len(strValor)

        With TEXT_RECORD
            .opcode = 4
            .length = 10
            'Length of the text portion of the record
            .TextLength = lngStringLength
            'Total length of the record
            .length = 8 + lngStringLength
            .Row = Row
            .col = Col
            .rgbAttr1 = CByte(CellHiddenLocked.xlsNormal)
            .rgbAttr2 = CByte(CellFontUsed + CellFormat)
            .rgbAttr3 = CByte(Alignment)
            'Put record header
            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.Row)
            writer.Write(.col)
            writer.Write(.rgbAttr1)
            writer.Write(.rgbAttr2)
            writer.Write(.rgbAttr3)
            writer.Write(.TextLength)
            'Then the actual string data
            For a As Integer = 1 To lngStringLength
                b = Asc(Mid(strValor, a, 1))
                writer.Write(b)
            Next
        End With

        Return True

        'Catch ex As Exception
        '   Return False
        'End Try

    End Function

    Public Function WriteInteger(ByVal CellFontUsed As CellFont, _
                               ByVal lRow As Integer, _
                               ByVal lCol As Integer, _
                               ByVal value As Object, _
                               Optional ByVal CellFormat As Integer = 0, _
                               Optional ByVal Alignment As CellAlignment = xlsRightAlign) As Boolean
        'Try
        Dim Row As Integer
        Dim Col As Integer
        Dim INTEGER_RECORD As tInteger

        'the row and column values are written to the excel file as
        'unsigned integers. Therefore, must convert the longs to integer.
        If lRow > 32767 Then
            Row = CInt(lRow - 65536)
        Else
            Row = CInt(lRow) - 1    'rows/cols in Excel binary file are zero based
        End If

        If lCol > 32767 Then
            Col = CInt(lCol - 65536)
        Else
            Col = CInt(lCol) - 1    'rows/cols in Excel binary file are zero based
        End If

        With INTEGER_RECORD
            .opcode = 2
            .length = 9
            .Row = Row
            .col = Col
            .rgbAttr1 = CByte(CellHiddenLocked.xlsNormal)
            .rgbAttr2 = CByte(CellFontUsed + CellFormat)
            .rgbAttr3 = CByte(Alignment)
            .intValue = CInt(value)

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.Row)
            writer.Write(.col)
            writer.Write(.rgbAttr1)
            writer.Write(.rgbAttr2)
            writer.Write(.rgbAttr3)
            writer.Write(.intValue)
        End With

        Return True

        'Catch ex As Exception
        'Return False
        'End Try

    End Function

    Public Function WriteNumber(ByVal CellFontUsed As CellFont, _
                               ByVal lRow As Integer, _
                               ByVal lCol As Integer, _
                               ByVal value As Object, _
                               Optional ByVal CellFormat As Integer = 0, _
                               Optional ByVal Alignment As CellAlignment = xlsRightAlign) As Boolean
        'Try
        Dim Row As Integer
        Dim Col As Integer
        Dim NUMBER_RECORD As tNumber


        'the row and column values are written to the excel file as
        'unsigned integers. Therefore, must convert the longs to integer.
        If lRow > 32767 Then
            Row = CInt(lRow - 65536)
        Else
            Row = CInt(lRow) - 1    'rows/cols in Excel binary file are zero based
        End If

        If lCol > 32767 Then
            Col = CInt(lCol - 65536)
        Else
            Col = CInt(lCol) - 1    'rows/cols in Excel binary file are zero based
        End If

        With NUMBER_RECORD
            .opcode = 3
            .length = 15
            .Row = Row
            .col = Col
            .rgbAttr1 = CByte(CellHiddenLocked.xlsNormal)
            .rgbAttr2 = CByte(CellFontUsed + CellFormat)
            .rgbAttr3 = CByte(Alignment)
            .NumberValue = value

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.Row)
            writer.Write(.col)
            writer.Write(.rgbAttr1)
            writer.Write(.rgbAttr2)
            writer.Write(.rgbAttr3)
            writer.Write(.NumberValue)
        End With

        Return True

        'Catch ex As Exception

        'Return False
        'End Try
    End Function

    Public Function WriteDate(ByVal CellFontUsed As CellFont, _
                               ByVal lRow As Integer, _
                               ByVal lCol As Integer, _
                               ByVal value As Object, _
                               Optional ByVal CellFormat As Integer = 20, _
                               Optional ByVal Alignment As CellAlignment = xlsRightAlign) As Boolean
        'Try
        Dim Row As Integer
        Dim Col As Integer
        Dim NUMBER_RECORD As tNumber


        'the row and column values are written to the excel file as
        'unsigned integers. Therefore, must convert the longs to integer.
        If lRow > 32767 Then
            Row = CInt(lRow - 65536)
        Else
            Row = CInt(lRow) - 1    'rows/cols in Excel binary file are zero based
        End If

        If lCol > 32767 Then
            Col = CInt(lCol - 65536)
        Else
            Col = CInt(lCol) - 1    'rows/cols in Excel binary file are zero based
        End If

        With NUMBER_RECORD
            .opcode = 3
            .length = 15
            .Row = Row
            .col = Col
            .rgbAttr1 = CByte(CellHiddenLocked.xlsNormal)
            .rgbAttr2 = CByte(CellFontUsed + CellFormat)
            .rgbAttr3 = CByte(Alignment)

            If value.GetType.ToString = "System.TimeSpan" Then
                .NumberValue = CDate(value.ToString).ToOADate
            Else
                .NumberValue = value.ToOAdate
            End If

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.Row)
            writer.Write(.col)
            writer.Write(.rgbAttr1)
            writer.Write(.rgbAttr2)
            writer.Write(.rgbAttr3)
            writer.Write(.NumberValue)
        End With

        Return True

        'Catch ex As Exception

        'Return False
        'End Try
    End Function

    Public Function SetMargin(ByVal Margin As MarginTypes, ByVal MarginValue As Double) As Boolean
        'Try
        Dim MarginRecord As MARGIN_RECORD_LAYOUT 'write the spreadsheet's layout information (in inches)

        With MarginRecord
            .opcode = Margin
            .length = 8
            .MarginValue = MarginValue 'in inches

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.MarginValue)
        End With

        Return True

        'Catch ex As Exception
        '    Return False
        'End Try
    End Function

    Public Function SetColumnWidth(ByVal FirstColumn As Byte, ByVal LastColumn As Byte, _
                                   ByVal WidthValue As Int16) As Boolean

        'Try

        Dim COLWIDTH As COLWIDTH_RECORD

        With COLWIDTH
            .opcode = 36
            .length = 4
            .col1 = FirstColumn - 1
            .col2 = LastColumn - 1
            .ColumnWidth = WidthValue * 256  'values are specified as 1/256 of a character

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.col1)
            writer.Write(.col2)
            writer.Write(.ColumnWidth)
        End With

        Return True

        'Catch ex As Exception
        'Return False
        'End Try
    End Function

    Public Function SetFont(ByVal FontName As String, _
                            ByVal FontHeight As Int16, _
                            ByVal FontFormat As FontFormatting) As Boolean

        'Try

        'you can set up to 4 fonts in the spreadsheet file. When writing a value such
        'as a Text or Number you can specify one of the 4 fonts (numbered 0 to 3)

        Dim FONTNAME_RECORD As FONT_RECORD
        Dim lngLengthFontName As String = Len(FontName)
        Dim b As Byte

        With FONTNAME_RECORD
            .opcode = 49
            .length = 5 + lngLengthFontName
            .FontHeight = FontHeight * 20
            .FontAttributes1 = CByte(FontFormat)  'bold/underline etc...
            .FontAttributes2 = CByte(0) 'reserved-always zero!!
            .FontNameLength = CByte(Len(FontName))

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.FontHeight)
            writer.Write(.FontAttributes1)
            writer.Write(.FontAttributes2)
            writer.Write(.FontNameLength)
        End With

        'Then the actual font name data            
        For a As Integer = 1 To lngLengthFontName
            b = Asc(Mid(FontName, a, 1))
            writer.Write(b)
        Next

        Return True

        'Catch ex As Exception
        ' Return False
        'End Try
    End Function

    Public Function SetHeader(ByVal HeaderText As String) As Boolean
        'Try
        Dim HEADER_RECORD As HEADER_FOOTER_RECORD
        Dim lngLenHeaderText As Integer = Len(HeaderText)
        Dim b As Byte

        With HEADER_RECORD
            .opcode = 20
            .length = 1 + lngLenHeaderText
            .TextLength = CByte(Len(HeaderText))

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.TextLength)
        End With

        'Then the actual Header text
        For a As Integer = 1 To lngLenHeaderText
            b = Asc(Mid(HeaderText, a, 1))
            writer.Write(b)
        Next

        Return True

        'Catch ex As Exception
        'Return False
        'End Try
    End Function

    Public Function SetFooter(ByVal FooterText As String) As Boolean

        'Try
        Dim FOOTER_RECORD As HEADER_FOOTER_RECORD
        Dim lngFooterText As Integer
        Dim b As Byte

        lngFooterText = Len(FooterText)

        With FOOTER_RECORD
            .opcode = 21
            .length = 1 + lngFooterText
            .TextLength = CByte(Len(FooterText))

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.TextLength)
        End With

        'Then the actual Header text
        For a As Integer = 1 To lngFooterText
            b = Asc(Mid(FooterText, a, 1))
            writer.Write(b)
        Next

        Return True

        'Catch ex As Exception
        'Return False
        'End Try
    End Function

    Public Function SetFilePassword(ByVal PasswordText As String) As Boolean
        'Try

        Dim FILE_PASSWORD_RECORD As PASSWORD_RECORD
        Dim lngLenPassWordText As Integer
        Dim b As Byte

        lngLenPassWordText = Len(PasswordText)

        With FILE_PASSWORD_RECORD
            .opcode = 47
            .length = lngLenPassWordText

            writer.Write(.opcode)
            writer.Write(.length)
        End With

        'Then the actual Password text            
        For a As Integer = 1 To lngLenPassWordText
            b = Asc(Mid(PasswordText, a, 1))
            writer.Write(b)
        Next

        Return True
        'Catch ex As Exception
        '    Return True
        'End Try
    End Function

    Public Function SetDefaultRowHeight(ByVal HeightValue As Int16) As Boolean
        '        Try

        Dim DEFHEIGHT As DEF_ROWHEIGHT_RECORD
        'Height is defined in units of 1/20th of a point. Therefore, a 10-point font
        'would be 200 (i.e. 200/20 = 10). This function takes a HeightValue such as
        '14 point and converts it the correct size before writing it to the file.        

        With DEFHEIGHT
            .opcode = 37
            .length = 2
            .RowHeight = HeightValue * 20  'convert points to 1/20ths of point            

            writer.Write(DEFHEIGHT.opcode)
            writer.Write(DEFHEIGHT.length)
            writer.Write(DEFHEIGHT.RowHeight)
        End With

        Return True
        'Catch ex As Exception
        '    Return False
        'End Try
    End Function

    Public Function SetRowHeight(ByVal lRow As Integer, ByVal HeightValue As Int16) As Boolean
        'Try

        Dim Row As Integer
        Dim ROWHEIGHTREC As ROW_HEIGHT_RECORD
        'the row and column values are written to the excel file as
        'unsigned integers. Therefore, must convert the longs to integer.

        If lRow > 32767 Then
            Row = CInt(lRow - 65536)
        Else
            Row = CInt(lRow) - 1    'rows/cols in Excel binary file are zero based
        End If

        'Height is defined in units of 1/20th of a point. Therefore, a 10-point font
        'would be 200 (i.e. 200/20 = 10). This function takes a HeightValue such as
        '14 point and converts it the correct size before writing it to the file.
        With ROWHEIGHTREC
            .opcode = 8
            .length = 16
            .RowNumber = Row
            .FirstColumn = 0
            .LastColumn = 256
            .RowHeight = HeightValue * 20 'convert points to 1/20ths of point
            .internal = 0
            .DefaultAttributes = 0
            .FileOffset = 0
            .rgbAttr1 = 0
            .rgbAttr2 = 0
            .rgbAttr3 = 0

            writer.Write(.opcode)
            writer.Write(.length)
            writer.Write(.RowNumber)
            writer.Write(.FirstColumn)
            writer.Write(.LastColumn)
            writer.Write(.RowHeight)
            writer.Write(.internal)
            writer.Write(.DefaultAttributes)
            writer.Write(.FileOffset)
            writer.Write(.rgbAttr1)
            writer.Write(.rgbAttr2)
            writer.Write(.rgbAttr3)
        End With

        Return True
        'Catch ex As Exception
        '    Return True
        'End Try
    End Function


#End Region

End Class

'==============================================================