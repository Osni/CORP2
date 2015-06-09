Imports System.IO
Imports System.IO.Compression
Partial Class _Default
    Inherits PageX

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Dim dta As New Data.DataTable
            Dim dtr As Data.DataRow
            '----------------------------------------
            With dta.Columns
                .Add("col1")
                .Add("col2")
                .Add("col3")
                .Add("col4")
                .Add("col5")
                .Add("col6")
            End With
            '----------------------------------------
            For i As Integer = 1 To 300
                dtr = dta.NewRow
                For Each dtc As Data.DataColumn In dta.Columns
                    dtr(dtc) = " valor " & dtc.ColumnName & " " & i.ToString
                Next
                dta.Rows.Add(dtr)
            Next
            '----------------------------------------
            grd.DataSource = dta
            grd.DataBind()
            ViewState("dta") = dta
            '----------------------------------------
        End If
    End Sub

    Protected Sub btnComp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnComp.Click
        txtComp.Text = ClassCompactarViewState.CompactarViewState(txt.Text)
    End Sub

    Protected Sub btnDesc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDesc.Click
        txt.Text = ClassCompactarViewState.DescompactarViewState(txtComp.Text)
    End Sub
End Class

Public Class PageX
    Inherits Page

    Protected Overrides Function LoadPageStateFromPersistenceMedium() As Object
        Dim viewState As String = Request.Form("__VSTATE")
        Dim bytes As Byte() = Convert.FromBase64String(viewState)
        'DESCOMPACTAR VIEWSTATE
        bytes = ClassCompactarViewState.DescompactarViewState(bytes)
        Dim formatter As LosFormatter = New LosFormatter()
        Return formatter.Deserialize(Convert.ToBase64String(bytes))
    End Function


    Protected Overrides Sub SavePageStateToPersistenceMedium(ByVal state As Object)
        Dim formatter As LosFormatter = New LosFormatter()
        Dim writer As StringWriter = New StringWriter()

        formatter.Serialize(writer, state)

        Dim viewStateString As String = writer.ToString()
        Dim bytes As Byte() = Convert.FromBase64String(viewStateString)

        'COMPACTAR VIEWSTATE
        bytes = ClassCompactarViewState.CompactarViewState(bytes)
        ClientScript.RegisterHiddenField("__VSTATE", Convert.ToBase64String(bytes))
    End Sub

End Class

Public Class ClassCompactarViewState    

    Public Shared Function CompactarViewState(ByVal bytes() As Byte) As Byte()
        Dim MSsaida As New MemoryStream
        Dim gzip As New GZipStream(MSsaida, CompressionMode.Compress, True)

        gzip.Write(bytes, 0, bytes.Length)
        gzip.Close()

        Return MSsaida.ToArray()
    End Function

    Public Shared Function CompactarViewState(ByVal texto As String) As String
        Dim MSsaida As New MemoryStream
        Dim gzip As New GZipStream(MSsaida, CompressionMode.Compress, True)
        Dim bytes() As Byte = System.Text.Encoding.ASCII.GetBytes(texto)

        gzip.Write(bytes, 0, bytes.Length)
        gzip.Close()

        Return System.Text.Encoding.ASCII.GetString(MSsaida.ToArray())
    End Function


    Public Shared Function DescompactarViewState(ByVal bytes As Byte()) As Byte()
        Dim MSentrada As MemoryStream = New MemoryStream()

        MSentrada.Write(bytes, 0, bytes.Length)
        MSentrada.Position = 0

        Dim gzip As GZipStream = New GZipStream(MSentrada, CompressionMode.Decompress, True)
        Dim MSsaida As MemoryStream = New MemoryStream()
        Dim buffer(64) As Byte
        Dim leitura As Integer = -1

        leitura = gzip.Read(buffer, 0, buffer.Length)

        While (leitura > 0)
            MSsaida.Write(buffer, 0, leitura)
            leitura = gzip.Read(buffer, 0, buffer.Length)
        End While

        gzip.Close()

        Return MSsaida.ToArray()
    End Function


    Public Shared Function DescompactarViewState(ByVal texto As String) As String
        Dim MSentrada As MemoryStream = New MemoryStream()
        Dim bytes() As Byte = System.Text.Encoding.ASCII.GetBytes(texto)

        MSentrada.Write(bytes, 0, bytes.Length)
        MSentrada.Position = 0

        Dim gzip As GZipStream = New GZipStream(MSentrada, CompressionMode.Decompress, True)
        Dim MSsaida As MemoryStream = New MemoryStream()
        Dim buffer(64) As Byte
        Dim leitura As Integer = -1

        leitura = gzip.Read(buffer, 0, buffer.Length)

        While (leitura > 0)
            MSsaida.Write(buffer, 0, leitura)
            leitura = gzip.Read(buffer, 0, buffer.Length)
        End While

        gzip.Close()

        Return System.Text.Encoding.ASCII.GetString(MSsaida.ToArray)
    End Function



End Class