Imports System.Data.OracleClient

#Region "CRT"
Public Class ClsCRT
    Public Enum TypeMsg
        Alert = 0 'OK
        Confirm = 1 'Confirm
    End Enum
    Public Function MsgBox(ByVal sMsg As String, Optional ByVal TType As TypeMsg = TypeMsg.Alert) As String
        Dim sRet As String = String.Empty
        sMsg = Replace(sMsg, "'", "")
        Select Case TType
            Case TypeMsg.Alert
                sRet = "<script language='javascript' ><!-- " & vbCrLf & "alert('" & sMsg & "') " & vbCrLf & "--></script>"
            Case TypeMsg.Confirm
                sRet = "<script language='javascript' ><!-- " & vbCrLf & "function Confirmar(){ if (confirm('" & sMsg & "')) document.forms[0].submit(); else return false;}" & vbCrLf & " --></script>"
        End Select
        Return sRet
    End Function
End Class
#End Region
#Region "SQL"
Public Class ClsSQL
    '-----------------------------------------
    Public cCols As New Collection
    Public sTable As String
    Public sColPrimaryKeyName As String
    Public sColPrimaryKeyValue As String
    Public sWHERE As String
    '-----------------------------------------
    Private sColName As String
    Private sColValue As String
    Private sColType As String
    Private i As Integer
    Private l As Long
    Public Enum TypeSQL
        EMPTY_T = 0 'Nada
        STRING_T = 1 'String
        DATE_YMD_T = 2 'Data YYYY-MM-DD
        DATE_YMD_HMS_T = 3 'Data YYYY-MM-DD HH:MM:SS
        DATE_DMY_T = 4 'Data DD/MM/YYYY
        DATE_DMY__HMS_T = 5 'Data DD/MM/YYYY  HH:MM:SS
        NUMERIC_T = 6 '
        MONEY_T = 7 'Limpa os ponto e substitui vírgula por ponto
    End Enum
    Public Function AddCol(ByVal sName As String, _
                           Optional ByVal sValue As Object = "", _
                           Optional ByVal TType As TypeSQL = TypeSQL.EMPTY_T) As Boolean
        Dim arr As New ArrayList
        Try
            arr.Add(sName)
            arr.Add(sValue)
            arr.Add(TType)
            cCols.Add(arr)
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Function SetQSQL(ByVal sSQL As String) As String
        Dim arr As New ArrayList
        With arr
            .Add(",")
            .Add("SELECT ")
            .Add("FROM ")
            .Add("WHERE ")
            .Add("ORDER BY ")
            .Add("GROUP BY ")
            .Add("HAVING ")
        End With
        For i = 0 To arr.Count - 1
            sSQL = Replace(sSQL, arr(i), arr(i) & vbCrLf)
        Next
        Return sSQL
    End Function
    Private Function GetValueTType(ByVal sValue As Object, ByVal Type As TypeSQL) As String
        Dim sRet As String = String.Empty
        Select Case Type 'Type TypeSQL
            Case TypeSQL.NUMERIC_T
                sRet = sValue  'Value
            Case TypeSQL.STRING_T, TypeSQL.EMPTY_T
                sRet = "'" & sValue & "'"
            Case TypeSQL.DATE_YMD_T
                sRet = "'" & Format(CDate(sValue), "yyyy-MM-dd") & "'"
            Case TypeSQL.DATE_DMY_T
                sRet = "'" & Format(CDate(sValue), "dd/MM/yyyy") & "'"
            Case TypeSQL.DATE_YMD_HMS_T 'yyyy-MM-dd H:mm:ss
                sRet = "'" & Format(CDate(sValue), "yyyy-MM-dd H:mm:ss") & "'"
            Case TypeSQL.DATE_DMY__HMS_T
                sRet = "'" & Format(CDate(sValue), "dd/MM/yyyy H:mm:ss") & "'"
            Case TypeSQL.MONEY_T
                sRet = Replace(Replace(sValue, ".", ""), ",", ".")
        End Select
        Return sRet
    End Function
    Public Function GetINSERT() As String
        sColName = String.Empty
        sColValue = String.Empty
        For i = 1 To cCols.Count
            sColName &= cCols.Item(i)(0) & ", " 'Name
            sColValue &= GetValueTType(cCols.Item(i)(1), cCols.Item(i)(2)) & ", " 'Value
        Next
        sColName = Left(sColName, Len(sColName) - 2)
        sColValue = Left(sColValue, Len(sColValue) - 2)
        Return "INSERT  INTO " & sTable & "( " & sColName & ") VALUES (" & sColValue & ")"
    End Function
    Public Function GetUPDATE() As String
        sColValue = String.Empty
        For i = 1 To cCols.Count
            sColValue &= cCols.Item(i)(0) & "=" & GetValueTType(cCols.Item(i)(1), cCols.Item(i)(2)) & ", " 'Value
        Next
        sColValue = Left(sColValue, Len(sColValue) - 2)
        Return "UPDATE " & sTable & " SET " & sColValue & " WHERE " & sColPrimaryKeyName & "=" & sColPrimaryKeyValue
    End Function
    Public Function GetDELETE() As String
        If sColPrimaryKeyName <> "" And sColPrimaryKeyValue <> "" Then
            Return "DELETE FROM " & sTable & " WHERE " & sColPrimaryKeyName & "=" & sColPrimaryKeyValue
        Else
            Return "DELETE FROM " & sTable & " WHERE " & sWHERE
        End If
    End Function
    Public Function GetCANCEL(Optional ByVal sSET As String = "Ativo='N'") As String
        If sColPrimaryKeyName <> "" And sColPrimaryKeyValue <> "" Then
            Return "UPDATE " & sTable & " SET " & sSET & " WHERE " & sColPrimaryKeyName & "=" & sColPrimaryKeyValue
        Else
            Return "UPDATE " & sTable & " SET " & sSET & " WHERE " & sWHERE
        End If
    End Function
    Public Function GetSELECT(Optional ByVal sColsWHERE As String = "") As String
        sColName = String.Empty
        For i = 1 To cCols.Count
            sColName &= cCols.Item(i)(0) & ", " 'Name
        Next
        sColName = Left(sColName, Len(sColName) - 2)
        If sColsWHERE <> "" Then sWHERE = sColsWHERE
        sColName = "SELECT " & sColName & " FROM " & sTable & IIf(sWHERE = "", "", " WHERE " & sWHERE)
        Return SetQSQL(sColName)
    End Function
End Class
Public Class ClsSQLParam
    Private mCmd As System.Data.Common.DbCommand
    Private mColuna As New Collection
    Private mColunaWhere As New Collection
    Private mTpSQL As TpSQL
    Private mTabela As String = ""
    Protected mPipeCommend As String = ""
    Public mStrSQL As New System.Text.StringBuilder
    Public Delegate Function EventoExterno(ByRef Objeto As Object, ByVal Valor As Object) As Object
    Public Overridable Property PipeCommend() As String
        Get
            Return IIf(mPipeCommend = "", "?", mPipeCommend)
        End Get
        Set(ByVal value As String)
            mPipeCommend = value
        End Set
    End Property
    Public Property Tabela() As String
        Get
            Return mTabela
        End Get
        Set(ByVal value As String)
            mTabela = value
        End Set
    End Property
    Public ReadOnly Property Command() As DbCommand
        Get
            Return mCmd
        End Get
    End Property
    Structure TpSQL
        Public Coluna As String
        Public Value As Object
        Public TType As System.Data.DbType
        Public EventoExterno As EventoExterno
    End Structure
    Public Sub AddColuna(ByVal Coluna As String, _
                   ByRef Value As Object, _
                   Optional ByVal TType As System.Data.DbType = DbType.String, _
                   Optional ByRef Evt As EventoExterno = Nothing)
        mTpSQL = New TpSQL
        With mTpSQL
            .Coluna = Coluna
            .Value = Value
            .TType = TType
            .EventoExterno = Evt
        End With
        mColuna.Add(mTpSQL, Coluna)
    End Sub
    Public Sub AddColunaWhere(ByVal Coluna As String, _
                   ByRef Value As Object, _
                   Optional ByVal TType As System.Data.DbType = DbType.String, _
                   Optional ByRef Evt As EventoExterno = Nothing)
        mTpSQL = New TpSQL
        With mTpSQL
            .Coluna = Coluna
            .Value = Value
            .TType = TType
            .EventoExterno = Evt
        End With
        mColunaWhere.Add(mTpSQL)
    End Sub
    Public Function SALVAR() As DbCommand
        Dim mSQLInsert As New System.Text.StringBuilder
        Dim mSQLInsertParam As New System.Text.StringBuilder
        Dim pr As System.Data.Common.DbParameter
        Dim sASPAS As String = " "
        '--------------------------------------
        For Each el As TpSQL In mColuna
            mSQLInsert.AppendLine(sASPAS & el.Coluna)
            If TypeOf mCmd Is System.Data.SqlClient.SqlCommand Then
                mSQLInsertParam.AppendLine(sASPAS & PipeCommend & el.Coluna)
            Else
                mSQLInsertParam.AppendLine(sASPAS & PipeCommend)
            End If
            sASPAS = ","
            '--------------------------------------
            pr = mCmd.CreateParameter
            With pr
                .ParameterName = el.Coluna
                .DbType = el.TType
                .Value = GetValueControle(el.Value, el.TType, el.EventoExterno)
            End With
            mCmd.Parameters.Add(pr)
        Next
        '--------------------------------------
        With mStrSQL
            .AppendLine("INSERT INTO " & mTabela)
            .AppendLine("(" & mSQLInsert.ToString & ")")
            .AppendLine("VALUES ")
            .AppendLine("(" & mSQLInsertParam.ToString & ")")
        End With
        '--------------------------------------
        mCmd.CommandText = mStrSQL.ToString
        Return mCmd
    End Function
    Public Function PESQUISAR() As DbCommand
        Dim mSQLWhere As New System.Text.StringBuilder
        Dim mSQLSelect As New System.Text.StringBuilder
        Dim pr As System.Data.Common.DbParameter
        Dim sASPAS As String = " "
        Dim iSeq As Int16 = 0
        '--------------------------------------
        For Each el As TpSQL In mColuna
            mSQLSelect.AppendLine(sASPAS & el.Coluna)
            sASPAS = ","
        Next
        sASPAS = ""
        For Each el As TpSQL In mColunaWhere
            If TypeOf mCmd Is System.Data.SqlClient.SqlCommand Then
                iSeq += 1
                mSQLWhere.AppendLine(sASPAS & el.Coluna & "=" & PipeCommend & el.Coluna & iSeq)
            Else
                mSQLWhere.AppendLine(sASPAS & el.Coluna & "=" & PipeCommend)
            End If
            sASPAS = " and "
            pr = mCmd.CreateParameter
            With pr
                If TypeOf mCmd Is System.Data.SqlClient.SqlCommand Then
                    .ParameterName = el.Coluna & iSeq
                Else
                    .ParameterName = el.Coluna
                End If
                .DbType = el.TType
                .Value = GetValueControle(el.Value, el.TType, el.EventoExterno)
            End With
            mCmd.Parameters.Add(pr)
        Next
        '--------------------------------------
        With mStrSQL
            .AppendLine("SELECT ")
            .Append(mSQLSelect.ToString)
            .AppendLine("FROM " & mTabela)
            .AppendLine("WHERE " & mSQLWhere.ToString)
        End With
        '--------------------------------------
        mCmd.CommandText = mStrSQL.ToString
        Return mCmd
    End Function
    Public Function CONTROLS(ByVal tbTableSource As DataTable) As Boolean
        If tbTableSource.Rows.Count = 1 Then
            For Each el As TpSQL In mColuna
                Call SetValueControle(tbTableSource.Rows(0)(el.Coluna), el.Value, el.TType, el.EventoExterno)
            Next
            Return True
        Else
            Return False
        End If
    End Function
    Public Function ALTERAR() As DbCommand
        Dim mSQLAlterar As New System.Text.StringBuilder
        Dim mSQLWhere As New System.Text.StringBuilder
        Dim pr As System.Data.Common.DbParameter
        Dim iSeq As Int16 = 0
        Dim sASPAS As String = " "
        '--------------------------------------
        For Each el As TpSQL In mColuna
            'Verificar se outros DBs tem alguma outra diferença ??????????????????
            'Somenete Para o SqlClient ??????????????????
            If TypeOf mCmd Is System.Data.SqlClient.SqlCommand Then
                iSeq += 1
                mSQLAlterar.AppendLine(sASPAS & el.Coluna & "=" & PipeCommend & el.Coluna & iSeq)
            Else
                mSQLAlterar.AppendLine(sASPAS & el.Coluna & "=" & PipeCommend)
            End If
            sASPAS = ","
            '--------------------------------------
            pr = mCmd.CreateParameter
            With pr
                If TypeOf mCmd Is System.Data.SqlClient.SqlCommand Then
                    .ParameterName = PipeCommend & el.Coluna & iSeq
                Else
                    .ParameterName = el.Coluna
                End If
                .DbType = el.TType
                .Value = GetValueControle(el.Value, el.TType, el.EventoExterno)
            End With
            mCmd.Parameters.Add(pr)
        Next
        sASPAS = ""
        '--------------------------------------
        For Each el As TpSQL In mColunaWhere
            If TypeOf mCmd Is System.Data.SqlClient.SqlCommand Then
                iSeq += 1
                mSQLWhere.AppendLine(sASPAS & el.Coluna & "=" & PipeCommend & el.Coluna & iSeq)
            Else
                mSQLWhere.AppendLine(sASPAS & el.Coluna & "=" & PipeCommend)
            End If
            sASPAS = " and "
            pr = mCmd.CreateParameter
            With pr
                If TypeOf mCmd Is System.Data.SqlClient.SqlCommand Then
                    .ParameterName = el.Coluna & iSeq
                Else
                    .ParameterName = el.Coluna
                End If
                .DbType = el.TType
                .Value = GetValueControle(el.Value, el.TType, el.EventoExterno)
            End With
            mCmd.Parameters.Add(pr)
        Next
        '--------------------------------------
        With mStrSQL
            .AppendLine("UPDATE " & mTabela)
            .AppendLine("SET ")
            .AppendLine(mSQLAlterar.ToString)
            .AppendLine("WHERE ")
            .AppendLine(mSQLWhere.ToString)
        End With
        '--------------------------------------
        mCmd.CommandText = mStrSQL.ToString
        Return mCmd
    End Function
    Private Function GetValueControle(ByRef Crt As Object, ByVal TType As System.Data.DbType, ByRef evt As Object) As Object
        Dim oRet As Object = Nothing
        Select Case Crt.GetType.ToString()
            Case "System.Web.UI.WebControls.TextBox"
                oRet = CType(Crt, UI.WebControls.TextBox).Text
            Case "System.Web.UI.WebControls.DropDownList"
                oRet = CType(Crt, UI.WebControls.DropDownList).SelectedValue.ToString
            Case "System.Web.UI.WebControls.ListBox"
                oRet = CType(Crt, UI.WebControls.ListBox).SelectedValue.ToString
            Case "System.Web.UI.WebControls.RadioButton"
                oRet = CType(Crt, UI.WebControls.RadioButton).Checked
            Case "System.Web.UI.WebControls.CheckBox"
                oRet = CType(Crt, UI.WebControls.CheckBox).Checked
            Case Else
                If evt IsNot Nothing Then
                    oRet = evt.Invoke(Crt, oRet)
                Else
                    oRet = Crt
                End If
        End Select
        Return ChkTipoDB(oRet, TType)
    End Function
    Public Overridable Function ChkTipoDB(ByVal Crt As Object, ByVal TType As System.Data.DbType) As Object
        Select Case TType
            Case DbType.DateTime, DbType.DateTime2, DbType.DateTimeOffset
                If Not IsDate(Crt) Then
                    Crt = DBNull.Value
                End If
            Case DbType.Currency, DbType.Decimal, DbType.Double, DbType.Int16, DbType.Int32, DbType.Int64, DbType.Single
                If Not IsNumeric(Crt) Then
                    Crt = DBNull.Value
                End If
        End Select
        Return Crt
    End Function

    Private Sub SetValueControle(ByVal Value As Object, ByRef Crt As Object, ByVal TType As System.Data.DbType, ByRef evt As Object)
        Value = IIf(IsDBNull(Value), String.Empty, Value)
        Select Case Crt.GetType.ToString()
            Case "System.Web.UI.WebControls.TextBox"
                CType(Crt, UI.WebControls.TextBox).Text = Value
            Case "System.Web.UI.WebControls.DropDownList"
                CType(Crt, UI.WebControls.DropDownList).SelectedValue = Value
            Case "System.Web.UI.WebControls.ListBox"
                CType(Crt, UI.WebControls.ListBox).SelectedValue = Value
            Case "System.Web.UI.WebControls.RadioButton"
                CType(Crt, UI.WebControls.RadioButton).Checked = Value
            Case "System.Web.UI.WebControls.CheckBox"
                CType(Crt, UI.WebControls.CheckBox).Checked = Value
            Case Else
                If evt IsNot Nothing Then
                    evt.Invoke(Crt, Value)
                Else
                    Crt = Value
                End If
        End Select
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="cmd">Command</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal cmd As System.Data.Common.DbCommand)
        mCmd = cmd
    End Sub
End Class
#End Region

Public Class ClsDB
    Implements IDisposable

#Region "Enums"
    Public Enum T_PROVIDER
        SQL = 0
        OLEDB = 1
        ODBC = 2
        FIREBIRDCLIENT = 3
        ORA = 4
        DB2 = 5
    End Enum
#End Region

#Region "Fields"
    '---------------------------------------------------------
    Public dts As System.Data.DataSet
    Public dtr As DbDataReader
    Public dta As DbDataAdapter
    '---------------------------------------------------------
    Private mFactory As DbProviderFactory
    Private mHasProviders As New Hashtable
    Private con As DbConnection
    Private dtTab As DataTable
    Private cmd As DbCommand
    Private tra As DbTransaction
    Private bCloseConn As Boolean
    Private mConnectionString As String
    Private mProviderName As String
    Private strXML As String
    Private mClsSQL As ClsSQLParam
    '---------------------------------------------------------
#End Region


#Region "Properties"
    ''' <summary>
    ''' Obsoleto. Usar ConnectionString.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Obsolete("Obsoleto. Usar ConnectionString.")> _
    Public Property sConStr() As String
        Get
            Return ConnectionString
        End Get
        Set(ByVal value As String)
            ConnectionString = value
        End Set
    End Property
    Public Property ProviderName(Optional ByVal ePROVIDER As T_PROVIDER = T_PROVIDER.OLEDB) As String
        Get
            mProviderName = mHasProviders(ePROVIDER)
            Return mProviderName
        End Get
        Set(ByVal value As String)
            mProviderName = value
        End Set
    End Property
    Public Property ConnectionString() As String
        Get
            Return mConnectionString
        End Get
        Set(ByVal value As String)
            mConnectionString = value
        End Set
    End Property
    Public ReadOnly Property GetConnection() As DbConnection
        Get
            Dim Conexao As DbConnection = mFactory.CreateConnection
            Conexao.ConnectionString = ConnectionString
            Return Conexao
        End Get
    End Property
    Public ReadOnly Property GetParameter(ByVal NomeParametro As String, ByVal Valor As Object) As DbParameter
        Get
            Dim prm As DbParameter = mFactory.CreateParameter
            prm.ParameterName = NomeParametro
            prm.Value = Valor
            Return prm
        End Get
    End Property
    Public ReadOnly Property GetCommand() As DbCommand
        Get
            Dim cmd As DbCommand = mFactory.CreateCommand
            cmd.Connection = GetConnection
            Return cmd
        End Get
    End Property
    Public Property SQLParam() As ClsSQLParam
        Get
            If mClsSQL Is Nothing Then
                mClsSQL = New ClsSQLParam(Me.GetCommand())
            End If
            Return mClsSQL
        End Get
        Set(ByVal value As ClsSQLParam)
            mClsSQL = value
        End Set
    End Property
#End Region

#Region "Constructors"

    Public Sub New()
        mHasProviders = New Hashtable
        With mHasProviders
            .Add(CShort(T_PROVIDER.OLEDB), "System.Data.OleDb")
            .Add(CShort(T_PROVIDER.ODBC), "System.Data.Odbc")
            .Add(CShort(T_PROVIDER.SQL), "System.Data.SqlClient")
            .Add(CShort(T_PROVIDER.FIREBIRDCLIENT), "FirebirdSql.Data.FirebirdClient")
            .Add(CShort(T_PROVIDER.ORA), "System.Data.OracleClient")
            .Add(CShort(T_PROVIDER.DB2), "IBM.Data.DB2")
        End With
        Me.mProviderName = mHasProviders(CShort(1))
        mFactory = DbProviderFactories.GetFactory(mProviderName)
    End Sub

    Public Sub New(ByVal ConnectionString As String, ByVal psProviderName As String)
        Me.New()
        Me.mProviderName = psProviderName
        Me.ConnectionString = ConnectionString
        mFactory = DbProviderFactories.GetFactory(mProviderName)
    End Sub

    Public Sub New(ByVal ConnectionString As String, Optional ByVal pProviderName As T_PROVIDER = T_PROVIDER.OLEDB)
        Me.New()
        Me.mProviderName = mHasProviders(CShort(pProviderName))
        Me.ConnectionString = ConnectionString
        mFactory = DbProviderFactories.GetFactory(mProviderName)
    End Sub
#End Region
#Region "Methods"
    Public Function GetDataTable(ByVal sSQL As String, _
                                 Optional ByRef pCon As DbConnection = Nothing, _
                                 Optional ByRef pTra As DbTransaction = Nothing, _
                                 Optional ByVal appendDataSet As Boolean = False) As DataTable
        dtTab = New DataTable
        dta = GetDataAdapter()
        '----------------------
        Try
            With dta.SelectCommand
                .CommandText = sSQL
                If pCon IsNot Nothing Then .Connection = pCon
                If pTra IsNot Nothing Then .Transaction = pTra
            End With
            '----------------------
            If appendDataSet Then
                If dts Is Nothing Then dts = New DataSet
            Else
                dts = New DataSet
            End If
            '----------------------
            dts.Tables.Add(dtTab)
            dta.Fill(dtTab)
            '----------------------
            Return dtTab
            '----------------------
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function GetDataTable(ByRef cmd As DbCommand, _
                                    Optional ByRef pCon As DbConnection = Nothing, _
                                    Optional ByRef pTra As DbTransaction = Nothing, _
                                    Optional ByVal appendDataSet As Boolean = False) As DataTable
        dtTab = New DataTable
        dta = GetDataAdapter()
        '----------------------
        Try
            dta.SelectCommand = cmd
            With dta.SelectCommand
                If pCon IsNot Nothing Then .Connection = pCon
                If pTra IsNot Nothing Then .Transaction = pTra
            End With
            '----------------------
            If appendDataSet Then
                If dts Is Nothing Then dts = New DataSet
            Else
                dts = New DataSet
            End If
            '----------------------
            dts.Tables.Add(dtTab)
            dta.Fill(dtTab)
            '----------------------
            Return dtTab
            '----------------------
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function GetDataReader(ByVal sSQL As String, Optional ByRef pcon As DbConnection = Nothing, Optional ByRef ptra As DbTransaction = Nothing) As DbDataReader
        con = GetConnection
        cmd = GetCommand
        With cmd
            .CommandText = sSQL
            If pcon IsNot Nothing Then .Connection = pcon Else .Connection = con
            If ptra IsNot Nothing Then .Transaction = ptra
        End With
        If pcon Is Nothing Then
            con.Open()
            dtr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
        Else
            dtr = cmd.ExecuteReader()
        End If
        Return dtr
    End Function
    Public Function GetDataReader(ByRef cmd As DbCommand, Optional ByRef pcon As DbConnection = Nothing, Optional ByRef ptra As DbTransaction = Nothing) As DbDataReader
        con = GetConnection
        With cmd
            If pcon IsNot Nothing Then .Connection = pcon Else .Connection = con
            If ptra IsNot Nothing Then .Transaction = ptra
        End With
        If pcon Is Nothing Then
            con.Open()
            dtr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
        Else
            dtr = cmd.ExecuteReader()
        End If
        Return dtr
    End Function
    Public Function GetDataAdapter() As DbDataAdapter
        dta = mFactory.CreateDataAdapter
        dta.SelectCommand = GetCommand
        Return dta
    End Function
    Public Function GetDataAdapter(ByVal sSQL As String) As DbDataAdapter
        Call GetDataAdapter()
        dta.SelectCommand.CommandText = sSQL
        Return dta
    End Function
    Public Function GetDataAdapter(ByRef cmd As DbCommand) As DbDataAdapter
        Call GetDataAdapter()
        dta.SelectCommand = cmd
        Return dta
    End Function
    Public Function ExecuteQueryReturnMax(ByVal sSQL As String, _
                                          ByVal Tabela As String, ByVal Campo As String, _
                                          Optional ByRef pCon As DbConnection = Nothing, _
                                          Optional ByRef pTra As DbTransaction = Nothing) As Object
        Dim oExecuteQueryReturnMax As Object
        con = GetConnection
        cmd = GetCommand
        With cmd
            .CommandText = sSQL
            If pCon IsNot Nothing Then .Connection = pCon Else .Connection = con
            If pTra IsNot Nothing Then .Transaction = pTra
        End With
        If pCon Is Nothing Then
            Try
                con.Open()
                tra = con.BeginTransaction
                cmd.Transaction = tra
                cmd.ExecuteNonQuery()
                oExecuteQueryReturnMax = GetDataTable("SELECT MAX(" & Campo & ") FROM " & Tabela, con, tra).Rows(0)(0)
                If oExecuteQueryReturnMax Is Nothing Then
                    ExecuteQueryReturnMax = 0
                Else
                    ExecuteQueryReturnMax = oExecuteQueryReturnMax
                End If
                tra.Commit()
                Return ExecuteQueryReturnMax
            Catch ex As Exception
                Try : tra.Rollback() : Catch : End Try
                Throw ex
            Finally
                con.Close()
            End Try
        Else
            With cmd
                .Connection = pCon
                .Transaction = pTra
                Try
                    .ExecuteNonQuery()
                    Return GetDataTable("SELECT MAX(" & Campo & ") FROM " & Tabela, pCon, pTra).Rows(0)(0)
                Catch ex As Exception
                    Throw ex
                End Try
            End With
        End If
    End Function
    Public Function ExecuteQueryReturnMax(ByRef cmd As DbCommand, _
                                         ByVal Tabela As String, ByVal Campo As String, _
                                         Optional ByRef pCon As DbConnection = Nothing, _
                                         Optional ByRef pTra As DbTransaction = Nothing) As Object
        Dim oExecuteQueryReturnMax As Object
        Dim cmdReturn As DbCommand = GetCommand
        con = GetConnection
        With cmd
            If pCon IsNot Nothing Then .Connection = pCon Else .Connection = con
            If pTra IsNot Nothing Then .Transaction = pTra
        End With
        If pCon Is Nothing Then
            Try
                con.Open()
                tra = con.BeginTransaction
                cmd.Transaction = tra
                cmd.ExecuteNonQuery()
                '-----------------------------------------------------------------
                cmdReturn.CommandText = "SELECT MAX(" & Campo & ") FROM " & Tabela
                oExecuteQueryReturnMax = GetDataTable(cmdReturn, con, tra).Rows(0)(0)
                '-----------------------------------------------------------------
                If oExecuteQueryReturnMax Is Nothing Then
                    ExecuteQueryReturnMax = 0
                Else
                    ExecuteQueryReturnMax = oExecuteQueryReturnMax
                End If
                tra.Commit()
                Return ExecuteQueryReturnMax
            Catch ex As Exception
                Try : tra.Rollback() : Catch : End Try
                Throw ex
            Finally
                con.Close()
            End Try
        Else
            With cmd
                .Connection = pCon
                .Transaction = pTra
                Try
                    .ExecuteNonQuery()
                    '-----------------------------------------------------------------
                    cmdReturn.CommandText = "SELECT MAX(" & Campo & ") FROM " & Tabela
                    '-----------------------------------------------------------------
                    Return GetDataTable(cmdReturn, pCon, pTra).Rows(0)(0)
                Catch ex As Exception
                    Throw ex
                End Try
            End With
        End If
    End Function
    Public Function ExecuteQuery(ByVal sSQL As String, Optional ByRef pcon As DbConnection = Nothing, Optional ByRef ptra As DbTransaction = Nothing) As Integer
        con = GetConnection
        cmd = GetCommand
        With cmd
            .CommandText = sSQL
            If pcon IsNot Nothing Then .Connection = pcon Else .Connection = con
            If ptra IsNot Nothing Then .Transaction = ptra
        End With
        If pcon Is Nothing Then
            Try
                con.Open()
                Return cmd.ExecuteNonQuery
            Catch ex As Exception
                Throw ex
            Finally
                con.Close()
            End Try
        Else
            Return cmd.ExecuteNonQuery
        End If
    End Function
    Public Function ExecuteQuery(ByRef cmd As DbCommand, Optional ByRef pcon As DbConnection = Nothing, Optional ByRef ptra As DbTransaction = Nothing) As Integer
        con = GetConnection
        With cmd
            If pcon IsNot Nothing Then .Connection = pcon Else .Connection = con
            If ptra IsNot Nothing Then .Transaction = ptra
        End With
        If pcon Is Nothing Then
            Try
                con.Open()
                Return cmd.ExecuteNonQuery
            Catch ex As Exception
                Throw ex
            Finally
                con.Close()
            End Try
        Else
            Return cmd.ExecuteNonQuery
        End If
    End Function
    Public Function GetSchema(Optional ByVal collectionName As String = "", Optional ByVal restrictionValues() As String = Nothing) As DataTable
        con = GetConnection
        con.ConnectionString = ConnectionString
        Try
            con.Open()
            If String.IsNullOrEmpty(collectionName) Then
                Return con.GetSchema()
            Else
                If restrictionValues IsNot Nothing Then
                    Return con.GetSchema(collectionName, restrictionValues)
                Else
                    Return con.GetSchema(collectionName)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            con.Close()
        End Try
    End Function

    '#####################################################################################
    Public Function GetSQLCon(Optional ByVal sUser As String = "sa", _
                            Optional ByVal sPassWord As String = "", _
                            Optional ByVal sCatalog As String = "mastes", _
                            Optional ByVal sServer As String = ".", _
                            Optional ByVal sProvider As String = "SQLOLEDB") As String
        sConStr = "Provider=" & sProvider & ";User ID=" & sUser & _
                                        ";password=" & sPassWord & _
                                        ";Initial Catalog=" & sCatalog & _
                                        ";server=" & sServer
        Return sConStr
        '"Provider=SQLOLEDB;password=senha;user id=usuario;Initial Catalog=banco;server=servidor"
    End Function
    '#####################################################################################
    Public Function GetOpenDB(Optional ByVal psConStr As String = "") As DbConnection
        Try
            If Trim(psConStr) <> "" Then
                sConStr = psConStr
            End If
            If sConStr = "" Then
                Throw New System.Exception("Uma string ex: 'Provider=SQLOLEDB;password=senha;user id=usuario;Initial Catalog=banco;server=servidor' de conexão é obrigatória")
            End If
            con = GetConnection()
            con.Open()
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return con
    End Function
    '#####################################################################################
    Private Function GetNewCommand(ByVal pCommandText As String, _
                                    Optional ByRef pCon As DbConnection = Nothing, _
                                    Optional ByRef pTra As DbTransaction = Nothing) As DbCommand
        Dim bConnTrans As Boolean
        Try
            bConnTrans = CheckConnTrans(pCon, pTra)
            cmd = GetCommand
            With cmd
                .CommandTimeout = 0
                .CommandText = pCommandText
                If bConnTrans Then
                    .Transaction = pTra
                End If
                .Connection = pCon
            End With
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        Return cmd
    End Function
    '#####################################################################################
    Public Function NullDB(ByRef pExpress As Object, Optional ByVal pReturn As Object = "") As Object
        Return IIf(IsDBNull(pExpress), pReturn, pExpress)
    End Function
    '#####################################################################################
    Public Function GetDataXML(ByVal sSQL As String, _
                               Optional ByRef pCon As DbConnection = Nothing) As String

        Try
            If Not pCon Is Nothing Then
                con = pCon
                bCloseConn = False
            Else
                con = GetOpenDB()
                bCloseConn = True
            End If
            dta = GetDataAdapter(sSQL)
            dts = New DataSet
            dta.Fill(dts)
            strXML = dts.GetXml()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            If bCloseConn Then
                con.Dispose()
                dtTab.Dispose()
                dts.Dispose()
            End If
        End Try
        Return strXML
    End Function
    '#####################################################################################
    Public Function GetDataXML(ByRef cmd As DbCommand, _
                               Optional ByRef pCon As DbConnection = Nothing) As String

        Try
            If Not pCon Is Nothing Then
                con = pCon
                bCloseConn = False
            Else
                con = GetOpenDB()
                bCloseConn = True
            End If
            dta = GetDataAdapter(cmd)
            dts = New DataSet
            dta.Fill(dts)
            strXML = dts.GetXml()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            If bCloseConn Then
                con.Dispose()
                dtTab.Dispose()
                dts.Dispose()
            End If
        End Try
        Return strXML
    End Function
    '#####################################################################################
    ''' <summary>
    ''' Obsoleto. Utilize ExecuteQuery(sSQL, pcon, ptra)
    ''' </summary>
    ''' <param name="sSQL"></param>
    ''' <param name="pCon"></param>
    ''' <param name="pTra"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetCommandSQL(ByVal sSQL As String, _
                                Optional ByRef pCon As DbConnection = Nothing, _
                                Optional ByRef pTra As DbTransaction = Nothing) As Integer
        Return ExecuteQuery(sSQL, pCon, pTra)
    End Function
    Public Function SetCommandSQL(ByRef cmd As DbCommand, _
                            Optional ByRef pCon As DbConnection = Nothing, _
                            Optional ByRef pTra As DbTransaction = Nothing) As Integer
        Return ExecuteQuery(cmd, pCon, pTra)
    End Function
    '#####################################################################################
    Public Function SetCommandSQLReturn(ByVal sSQLExecuteReturn As String, _
                                        ByVal sSQLExecute As String, _
                                        Optional ByRef pCon As DbConnection = Nothing, _
                                        Optional ByRef pTra As DbTransaction = Nothing) As DbDataReader
        Dim bConnTrans As Boolean
        Try
            bConnTrans = CheckConnTrans(pCon, pTra)
            If Not pCon Is Nothing Then
                con = pCon
                bCloseConn = False
            Else
                con = GetConnection()
                bCloseConn = True
            End If
            If bCloseConn Then con.Open()
            If bConnTrans Then tra = pTra Else  : tra = con.BeginTransaction()
            If Trim(sSQLExecute) <> "" Then
                ExecuteQuery(sSQLExecute, con, tra)
            End If
            If sSQLExecuteReturn <> "" Then
                dtr = GetDataReader(sSQLExecuteReturn, con, tra)
                bCloseConn = False
            End If
            If Not bConnTrans Then tra.Commit()
            '------------------------------------------------
        Catch ex As DbException
            Try : tra.Rollback() : Catch : End Try
            Throw New Exception(ex.Message)
        Finally
            If Not bConnTrans Then
                con.Close()
                tra.Dispose()
                con.Dispose()
            End If
            cmd.Dispose()
        End Try
        Return dtr
    End Function
    '#####################################################################################
    ''' <summary>
    ''' Obsoleto. Utilize o método ExecuteQueryReturnMax(sSQL, Tabela, Campo, pCon, pTra).
    ''' </summary>
    ''' <param name="sSQLExecute"></param>
    ''' <param name="sTableMaxReturn"></param>
    ''' <param name="sColMaxReturn"></param>
    ''' <param name="pCon"></param>
    ''' <param name="pTra"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetCommandSQLReturnMax(ByVal sSQLExecute As String, _
                                            ByVal sTableMaxReturn As String, _
                                            ByVal sColMaxReturn As String, _
                                            Optional ByRef pCon As DbConnection = Nothing, _
                                            Optional ByRef pTra As DbTransaction = Nothing) As String
        Return ExecuteQueryReturnMax(sSQLExecute, sTableMaxReturn, sColMaxReturn, pCon, pTra)
    End Function
    Public Function SetCommandSQLReturnMax(ByRef cmdSQLExecute As DbCommand, _
                                        ByVal sTableMaxReturn As String, _
                                        ByVal sColMaxReturn As String, _
                                        Optional ByRef pCon As DbConnection = Nothing, _
                                        Optional ByRef pTra As DbTransaction = Nothing) As String
        Return ExecuteQueryReturnMax(cmdSQLExecute, sTableMaxReturn, sColMaxReturn, pCon, pTra)
    End Function
    '#####################################################################################
    Public Function AddCombo(ByVal sSQL As String, _
                            ByVal sDataValueField As String, _
                            ByVal sDataTextField As String, _
                            ByRef cbo As Object, _
                            Optional ByVal sTextoSemSelecao As String = "[Selecione]") As Boolean
        Try
            '------------------------------------------------------------
            If Not String.IsNullOrEmpty(sSQL.Trim) Then
                dtr = GetDataReader(sSQL)
                With cbo
                    .DataSource = dtr
                    .DataTextField = sDataTextField
                    .DataValueField = sDataValueField
                    .DataBind()
                    If sTextoSemSelecao IsNot Nothing Then .Items.Insert(0, New UI.WebControls.ListItem(sTextoSemSelecao, ""))
                End With
            Else
                Throw New Exception("Um Comando SQL É nescessário")
            End If
            '------------------------------------------------------------
        Catch ex As DbException
            Throw New Exception(ex.Message)
        Finally
            If dtr IsNot Nothing Then dtr.Dispose()
        End Try
        Return True
    End Function
    '#####################################################################################
    Public Function AddCombo(ByRef cmd As DbCommand, _
                            ByVal sDataValueField As String, _
                            ByVal sDataTextField As String, _
                            ByRef cbo As Object, _
                            Optional ByVal sTextoSemSelecao As String = "[Selecione]") As Boolean
        Try
            '------------------------------------------------------------
            If Not cmd Is Nothing Then
                dtr = GetDataReader(cmd)
                With cbo
                    .DataSource = dtr
                    .DataTextField = sDataTextField
                    .DataValueField = sDataValueField
                    .DataBind()
                    If sTextoSemSelecao <> "" Then .Items.Insert(0, New UI.WebControls.ListItem(sTextoSemSelecao, ""))
                End With
            Else
                Throw New Exception("Um Comando SQL É nescessário")
            End If
            '------------------------------------------------------------
        Catch ex As DbException
            Throw New Exception(ex.Message)
        Finally
            If dtr IsNot Nothing Then dtr.Dispose()
        End Try
        Return True
    End Function
    Private Function CheckConnTrans(ByRef pCon As DbConnection, ByRef pTra As DbTransaction) As Boolean
        Dim bReturn As Boolean
        If Not pTra Is Nothing Then
            bReturn = True
            If pCon Is Nothing Then
                bReturn = False
                Throw New System.Exception("Quando uma transação é passada como paramêtro, uma conexão é obrigatórias")
            End If
        Else
            bReturn = False
        End If
        Return bReturn
    End Function
#End Region


#Region " IDisposable Support "
    ' To detect redundant calls
    Private disposedValue As Boolean = False
    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free unmanaged resources when explicitly called
                Call DisposeObj(dts) 'System.Data.DataSet
                Call DisposeObj(dtr) 'DbDataReader
                Call DisposeObj(dta) 'DbDataAdapter                        
                Call DisposeObj(con) 'DbConnection
                Call DisposeObj(dtTab) 'DataTable
                Call DisposeObj(cmd) 'DbCommand
                Call DisposeObj(tra) 'DbTransaction
            End If
            ' TODO: free shared unmanaged resources
        End If
        '----------------------------
        Me.disposedValue = True
        '----------------------------
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        '----------------------------
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
        '----------------------------
    End Sub
    'Dispose a object when 
    Private Sub DisposeObj(ByRef obj As Object)
        If obj IsNot Nothing Then
            Try : obj.Dispose() : Catch : End Try
        End If
    End Sub
#End Region

End Class
