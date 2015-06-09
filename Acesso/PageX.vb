Imports Microsoft.VisualBasic

Public Class PageX
    Inherits System.Web.UI.Page

    Protected Sub PageInitAccess(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load        
        Dim oAcesso As New ClsAcesso
        Dim vURLPage As Object
        Dim sPage As String
        vURLPage = Request.AppRelativeCurrentExecutionFilePath.Split("/")
        sPage = vURLPage(UBound(vURLPage))
        '-------------------------------------------
        oAcesso.ChecaLogin(sPage)
        '-------------------------------------------
    End Sub
End Class
