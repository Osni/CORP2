<%@ Page Language="VB" Debug="true" EnableEventValidation="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CORP.NET" %>


<script runat="server">

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call SetList()
    End Sub
      
    Private Function SetList() As Boolean
        
        Dim URL As String = ""
        Dim ArrURL As Object
        '----------------------------------------
        Dim ConnectionString As String = ""
        Dim TableSelect As String = ""
        Dim ColumnValue As String = ""
        Dim ColumnText As String = ""
        Dim ColumnWhere As String = ""
        Dim ColumnValueWhere As String = ""
        '----------------------------------------
        Dim sSQL As String = ""
        '----------------------------------------
        URL = IIf(Request("autoCompleteURL") Is Nothing, "", Request("autoCompleteURL"))
        If URL = "" Then Return False
        '----------------------------------------
        URL = CorpCripto.DecryptString(URL)
        '----------------------------------------
        ArrURL = Split(URL, "$")
        '----------------------------------------
        ConnectionString = ArrURL(0)
        TableSelect = ArrURL(1)
        ColumnValue = ArrURL(2)
        ColumnText = ArrURL(3)
        ColumnWhere = ArrURL(4)                       
        ColumnValueWhere = CorpCripto.HexToString(Request("prm"))
        '----------------------------------------        
        sSQL = "SELECT " & ColumnValue & ", " & ColumnText & " FROM " & TableSelect & IIf(ColumnWhere <> "", " WHERE " & ColumnWhere & " LIKE '%" & ColumnValueWhere.Trim & "%'", "")
        Dim adp As Data.OleDb.OleDbDataAdapter = New Data.OleDb.OleDbDataAdapter(sSQL, New Data.OleDb.OleDbConnection(ConnectionString))
        Dim dts As New Data.DataSet
        Dim dt As New Data.DataTable        
        '----------------------------------------
        adp.Fill(dts, 1, 15, TableSelect)
        dt = dts.Tables(0)
        '----------------------------------------
        If dt.Rows.Count > 0 Then
            Response.Write("<select size=6 id=""autoCompleteList"">")
            For i As Integer = 0 To dt.Rows.Count - 1
                Response.Write("<option value=""" & dt.Rows(i)(ColumnValue) & """" & IIf(i = 0, " SELECTED", "") & ">" & dt.Rows(i)(ColumnText) & "</option>")
            Next
            Response.Write("</select>")
        Else
            Response.Write("&&&")
        End If
        Return True
    End Function

</script>

<html>
<head runat="server"></head>
</html>