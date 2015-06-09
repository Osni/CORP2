Imports System.IO
Imports System.IO.Compression


Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox2.Text = ClassCompactarViewState.CompactarViewState(TextBox1.Text)
        Label1.Text = TextBox1.Text.Length
        Label2.Text = TextBox2.Text.Length
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox3.Text = ClassCompactarViewState.DescompactarViewState(TextBox2.Text)
        Label3.Text = TextBox3.Text.Length
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = TextBox3.Text Then
            MsgBox("OK", MsgBoxStyle.Information)
        Else
            MsgBox("Ferrou", MsgBoxStyle.Critical)
        End If
    End Sub

End Class

Public Class ClassCompactarViewState

    Public Shared Function CompactarViewState(ByVal texto As String) As String
        Dim MSsaida As New MemoryStream
        Dim gzip As New GZipStream(MSsaida, CompressionMode.Compress, True)
        Dim bytes() As Byte = System.Text.UTF8Encoding.UTF8.GetBytes(texto)

        gzip.Write(bytes, 0, bytes.Length)
        gzip.Close()

        Return Convert.ToBase64String(MSsaida.ToArray())
    End Function

    Public Shared Function DescompactarViewState(ByVal texto As String) As String
        Dim MSentrada As MemoryStream = New MemoryStream()
        Dim bytes() As Byte = Convert.FromBase64String(texto)

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

        Return System.Text.UTF8Encoding.UTF8.GetString(MSsaida.ToArray)
    End Function


End Class