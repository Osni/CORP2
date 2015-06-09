Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.VisualBasic
Imports ClsExcel.ValueTypes
Imports System.IO

Partial Class _Default
    Inherits System.Web.UI.Page


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click


    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim Response As HttpResponse = HttpContext.Current.Response
        'Dim aBytes() As Byte = CType(strm, MemoryStream).ToArray

        'With Response
        '    .Clear()
        '    .AddHeader("Content-Disposition", "attachment; filename=" & strFileName)
        '    .AddHeader("Content-Length", aBytes.Length.ToString())
        '    .ContentType = "application/vnd.ms-excel"
        '    .BinaryWrite(aBytes)
        'End With

        Dim excel As New ClsGetExcelFile
        Dim adp As OleDbDataAdapter = New OleDbDataAdapter("SELECT uniCodigo, accCodigoInterno, accNome, Data FROM db_CentroCusto..CC_AREA_CENTRO_CUSTO order by uniCodigo", ConnectionStrings("cnnStr").ToString)

        With excel
            adp.Fill(.DataSource)

            .XLSGenerateType = ClsExcel.GenerateType.ToMemory

            .AddColumnTitle("uniCodigo", "Cód. Unidade")
            .AddColumnTitle("accCodigoInterno", "Cód.CC")
            .AddColumnTitle("accNome", "Descrição CC")
            .AddColumnTitle("Data", "Data Inserção")

            .AddGroupColumn("uniCodigo")

            .FileName = "ovo do pato.xls"
            .Titulo = "Relatório de Centros de Custo"
            .LinhasPorPagina = 100
            .GenerateXLS()
        End With

        Dim aBytes() As Byte = CType(excel.GetStream, MemoryStream).ToArray

        With Response
            .Clear()
            .AddHeader("Content-Disposition", "attachment; filename=" & excel.FileName)
            .AddHeader("Content-Length", aBytes.Length.ToString())
            .ContentType = "application/vnd.ms-excel"
            .BinaryWrite(aBytes)
          	.End
        End With

    End Sub
End Class
