Public Class DBUtils

    Public Shared Function BuildClause(elements As List(Of String), types As List(Of String)) As String
        Dim clause As String = ""

        For i As Integer = 0 To elements.Count - 1
            Dim columnName As String = elements(i)
            Dim type As String = types(i)

            Dim typeDef As String

            If type = NUMBER_TYPE Then
                typeDef = "double"
            ElseIf type = DATE_TYPE Then
                typeDef = "datetime"
            Else
                typeDef = "varchar(255)"
            End If
            clause = clause + columnName + " " + typeDef + ","
        Next

        clause = clause.Substring(0, clause.Length - 1)

        Return clause
    End Function
End Class
