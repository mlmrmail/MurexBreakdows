Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fr As New FileReader
        Try
            dgv.DataSource = fr.ReadExcelFile().Tables(0)
        Catch ex As Exception
            MsgBox(ex)
        End Try
    End Sub
End Class
