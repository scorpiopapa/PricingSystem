Public Class Log
    Public Shared Sub WriteLine(msg As String, ParamArray args() As String)
        Console.WriteLine(String.Format(msg, args))
    End Sub

    Public Shared Sub WriteLine(ex As Exception)
        Console.WriteLine(ex)
    End Sub

    Public Shared Sub WriteLine(table As DataTable)
        For Each row As DataRow In table.Rows
            For Each column As DataColumn In table.Columns
                Dim value = row(column)
                Dim v As String
                If DBNull.Value.Equals(value) Then
                    v = ""
                Else
                    v = row(column).ToString
                End If
                Console.Write(column.ColumnName + "=" + v + " ")
            Next
            Console.WriteLine("")
        Next
    End Sub
End Class
