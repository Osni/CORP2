
Partial Class Default2
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        lblTextTab.Text = acpTabela.Text
        lblValueTab.Text = acpTabela.Value

        lblTextItens.Text = acpItensTabela.Text
        lblValueItens.Text = acpItensTabela.Value

    End Sub

End Class
