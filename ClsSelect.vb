Imports System.IO
Imports System.Net
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls


Public Class ClsSelect
    Inherits CompositeControl

    Private Sub ClsSelect_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        With Page.ClientScript
            If Not .IsClientScriptIncludeRegistered("script_select") Then
                .RegisterClientScriptInclude(Me.GetType, "script_select", "http://10.0.0.238/corpnet/select.js")
                .RegisterStartupScript(Me.GetType, "init_select", "javascript:select_search.Initialize();", True)
            End If
        End With
    End Sub

End Class
