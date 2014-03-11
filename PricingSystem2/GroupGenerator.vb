Public Class GroupGenerator

    Public Shared Function Group(data As List(Of List(Of String))) As Dictionary(Of String, List(Of String))
        Dim dic As New Dictionary(Of String, List(Of String))

        For Each l As List(Of String) In data
            Dim key As String = l(0)
            Dim values As List(Of String)

            If dic.ContainsKey(key) Then
                values = dic.Item(key)
            Else
                values = New List(Of String)
                dic.Add(key, values)
            End If

            values.AddRange(l)
            values.RemoveAt(0)
        Next

        Return dic
    End Function
End Class
