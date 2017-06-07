Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel


Public Class FileReader

    Private filePath As String = String.[Empty]

    Private appXL As Excel.Application
    Private wbXl As Excel.Workbook
    Private shXL As Excel.Worksheet


    Public Sub New()
        openDialogFile()
        openExcelFile()
        cleanInputFile()
    End Sub

    Private Sub openDialogFile()
        Dim ofd As New OpenFileDialog
        ofd.Title = "Select the file to analyse ..."
        ofd.Filter = "Excel 2007 Files (*.xls)|*.xls | Excel 2017+ Files (*.xlsb)|*.xlsb"
        ofd.FilterIndex = 2
        If ofd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            filePath = ofd.FileName
        End If
    End Sub

    Private Sub openExcelFile()
        If Not IO.File.Exists(filePath) Then
            Throw New Exception("The file specified does not exist or is empty !")
        Else
            appXL = CreateObject("Excel.Application")
            appXL.Visible = True
            wbXl = appXL.Workbooks.Open(filePath)
        End If
    End Sub

    Private Sub cleanInputFile(Optional sheetName As String = "Generated from Murex")

        shXL = wbXl.Sheets(sheetName)
        shXL.Activate()

        Dim colPos As Short = 1

        For col = 65 To (65 + getThelastColumn() - 3)
            Dim curPos As Integer = getTheFirstRow() + 1
            Dim maxPos As Integer = getTheLastRow()
            Dim tmpVal As Object

            For i = curPos To maxPos
                If shXL.Cells(i, colPos).MergeCells = True And shXL.Cells(i, colPos).MergeArea.Columns.Count = 1 Then
                    Dim totRows As Integer = (i + shXL.Cells(i, colPos).MergeArea.Cells.Count) - 1
                    tmpVal = shXL.Range(Chr(col) + i.ToString).Value
                    shXL.Range(Chr(col) + i.ToString).UnMerge()
                    For j = i To totRows
                        shXL.Cells(j, colPos).value = tmpVal
                    Next j
                    i = totRows
                End If
            Next i

            colPos += 1

        Next col



    End Sub

    Private Function getThelastColumn(Optional sheetName As String = "Generated from Murex") As UInteger

        shXL = wbXl.Sheets(sheetName)
        shXL.Activate()
        Return shXL.UsedRange.Columns.Count

    End Function

    Private Function getTheFirstColumn(Optional sheetName As String = "Generated from Murex") As UInteger
        shXL = wbXl.Sheets(sheetName)
        shXL.Activate()
        Return shXL.UsedRange.Column
    End Function

    Private Function getTheLastRow(Optional sheetName As String = "Generated from Murex") As UInteger

        shXL = wbXl.Sheets(sheetName)
        shXL.Activate()
        Return shXL.UsedRange.Rows.Count

    End Function

    Private Function getTheFirstRow(Optional sheetName As String = "Generated from Murex") As UInteger
        shXL = wbXl.Sheets(sheetName)
        shXL.Activate()
        Return shXL.UsedRange.Row
    End Function

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

        With ds.Tables(0)

            .Columns(0).ColumnName = "txDate"
            .Columns(1).ColumnName = "Instrument"
            .Columns(2).ColumnName = "Portfolio"
            .Columns(3).ColumnName = "Currency"
            .Columns(4).ColumnName = "Flows"

        End With

        ReadExcelFile = ds

    End Function

    Public Function getListOfRows(ds As DataTable) As List(Of SourceDataStructure)

        Dim tmp As New List(Of SourceDataStructure)
        Dim currData As New SourceDataStructure
        Dim d As Date

        For i = 1 To ds.Rows.Count - 1

            d = Date.FromOADate(ds(i).Item("TxDate"))
            currData.TxDate = d
            currData.InstrumentDesc = ds(i).Item("Instrument")
            currData.Portf = ds(i).Item("Portfolio")
            currData.Curr = ds(i).Item("Currency")
            currData.Flow = ds(i).Item("Flows")
            tmp.Add(currData)
        Next

        Return tmp



    End Function

    Private Sub cleanRessources()

        ' MsgBox("cleaning ressources ...")
        releaseObject(shXL)
        releaseObject(wbXl)
        releaseObject(appXL)

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        cleanRessources()
        MyBase.Finalize()
    End Sub
End Class
