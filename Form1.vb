Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim classeDB As ClsDB
        classeDB = New ClsDB("User ID= admin;password= 311097;Data Source =  127.0.0.1; Integrated Security=no;", ClsDB.T_PROVIDER.ORA)
        classeDB.GetOpenDB()

    End Sub
End Class