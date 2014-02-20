Public Class Log
    Public Shared Sub WriteLine(msg As String, ParamArray args() As String)
        Console.WriteLine(String.Format(msg, args))
    End Sub

    Public Shared Sub WriteLine(ex As Exception)
        Console.WriteLine(ex)
    End Sub
End Class
