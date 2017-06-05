Imports System.Data.OleDb

Public Class FileReader

    Private filePath As String = String.[Empty]

    Public Sub New()
        openFile()

        'If IO.File.Exists(filePath) = False Then
        '    Throw New Exception("The File doesn't exist")
        'End If

    End Sub

    Private Sub openFile()
        Dim ofd As New OpenFileDialog
        ofd.Title = "Select the file to analyse ..."
        ofd.Filter = "Excel 2007 Files (*.xls)|*.xls | Excel 2017+ Files (*.xlsb)|*.xlsb"
        If ofd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            filePath = ofd.FileName
        End If
    End Sub



    Public Function ReadExcelFile() As DataSet

        Dim conn As String = String.Empty
        Dim ext As String = IO.Path.GetExtension(filePath)
        Dim ds As New DataSet()


        If ext = ".xls" Then
            conn = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath +
                        ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"
        ElseIf ext = ".xlsb" Then
            conn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath +
                    ";Extended Properties='Excel 12.0;HDR=NO';"
        End If

        Dim adapter As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM [Generated from Murex$]", conn)
        adapter.Fill(ds)


        ReadExcelFile = ds

    End Function

End Class
