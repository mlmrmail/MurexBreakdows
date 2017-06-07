Public Class dataManagement

    Public Function getListOfBreakdowns(dt As DataTable) As List(Of String)
        Dim param As New List(Of String)

        For i = 0 To dt.Columns.Count - 1

            param.Add(dt.Rows(0).Item(i).ToString)

        Next

        Return param

    End Function

    Public Function getListOfCurrencies(dt As DataTable) As List(Of String)
        Dim currencies As New List(Of String)

        For i = 1 To dt.Rows.Count - 1

            If Not currencies.Contains(dt.Rows(i).Item(3).ToString) Then
                currencies.Add(dt.Rows(i).Item(3).ToString)
            End If

        Next

        Return currencies
    End Function



End Class
