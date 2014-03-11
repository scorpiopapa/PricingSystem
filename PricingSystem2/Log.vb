Imports System.IO
Public Class Log
#If DEBUG Then
    Private Shared ReadOnly LOG_FILE As String = "c:\price.log"
#Else
    Private Shared ReadOnly LOG_FILE As String = Application.StartupPath + "\price.log"
#End If

    Public Shared Sub WriteLine(msg As String, ParamArray args() As String)
        Console.WriteLine(String.Format(msg, args))
        WriteToFile(String.Format(msg, args))
    End Sub

    Public Shared Sub WriteLine(msg() As String)
        For Each m As String In msg
            Console.WriteLine(m)
            WriteToFile(m)
        Next
    End Sub

    Public Shared Sub WriteLine(ex As Exception)
        Console.WriteLine(ex)
        WriteToFile(ex.ToString)
        'WriteToFile(ex.Source)
        'WriteToFile(ex.Message)
        'WriteToFile(ex.StackTrace)
        'My.Computer.FileSystem.WriteAllText("c:\price.log", ex.StackTrace, True)

        'With My.Computer.FileSystem


        'End With

        'Dim fso As New System.IO.WatcherChangeTypes



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

    Private Shared Sub WriteToFile(line As String)
        Dim writer As New StreamWriter(LOG_FILE, True)
        With writer
            .WriteLine(line)
            .Close()
            .Dispose()
        End With
    End Sub
End Class
