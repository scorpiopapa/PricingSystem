Imports System.Data.OleDb
Public Class DBUtils

    Public Shared Function BuildClause(elements As List(Of String), types As List(Of String)) As String
        Dim clause As String = ""

        For i As Integer = 0 To elements.Count - 1
            Dim columnName As String = elements(i)
            Dim type As String = types(i)

            Dim typeDef As String = GetDataType(type)

            'If type = NUMBER_TYPE Then
            '    typeDef = "double"
            'ElseIf type = DATE_TYPE Then
            '    typeDef = "datetime"
            'Else
            '    typeDef = "varchar(255)"
            'End If

            'typeDef = GetDataTyApe(type)
            clause = clause + columnName + " " + typeDef + ","
        Next

        clause = clause.Substring(0, clause.Length - 1)

        Return clause
    End Function

    Public Shared Function GetDataType(type As String) As String
        Dim typeDef As String

        If type = NUMBER_TYPE Then
            typeDef = "double"
        ElseIf type = DATE_TYPE Then
            typeDef = "datetime"
        Else
            typeDef = "varchar(255)"
        End If

        Return typeDef
    End Function

    Public Shared Function GetDataColumnType(type As String) As System.Type
        Dim typeDef As System.Type

        If type = NUMBER_TYPE Then
            typeDef = DBConstants.INT
        ElseIf type = DATE_TYPE Then
            typeDef = DBConstants.DT
        Else
            typeDef = DBConstants.STR
        End If

        Return typeDef
    End Function

    'Public Shared Function QueryTable(sql As String, conn As OleDbConnection, Optional tran As OleDbTransaction = Nothing) As DataTable
    '    Dim cmd As New OleDbCommand(sql, conn)
    '    If Not IsNothing(tran) Then
    '        cmd.Transaction = tran
    '    End If

    '    Dim adapter As New OleDbDataAdapter(cmd)
    '    Dim table As New DataTable
    '    adapter.Fill(table)

    '    Return table
    'End Function
End Class
