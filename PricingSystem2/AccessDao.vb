Imports System.Data.OleDb
Public Class AccessDao

    Private conn As OleDbConnection
    Private tran As OleDbTransaction

    Public Sub New(conn As OleDbConnection)
        Me.conn = conn
        conn.Open()
      
    End Sub

    Public Sub New(conn As OleDbConnection, Optional ByRef tran As OleDbTransaction = Nothing)
        Me.conn = conn
        Me.tran = tran
    End Sub

    Public Function DataExists(tableName As String, columnName As String, value As String, type As OleDbType) As Boolean
        Dim sql As String = String.Format("select * from {0} where {1} = ?", tableName, columnName, value)
        Dim cmd As OleDbCommand = CreateCommand(sql)

        cmd.Parameters.Add(New OleDbParameter("@" + columnName, type) With {.Value = value})
        Dim adapter As New OleDbDataAdapter(cmd)
        Dim table As New DataTable
        adapter.Fill(table)

        Return table.Rows.Count > 0
    End Function

    Public Function SelectTable(sql As String) As DataTable
        Dim cmd As OleDbCommand = CreateCommand(sql)

        Dim adapter As New OleDbDataAdapter(cmd)
        Dim table As New DataTable
        adapter.Fill(table)

        Return table
    End Function

    Public Function SelectTableForUpdate(sql As String) As OleDbDataAdapter
        Dim cmd As OleDbCommand = CreateCommand(sql)

        Dim adapter As New OleDbDataAdapter(cmd) With {.MissingSchemaAction = MissingSchemaAction.AddWithKey}
        Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)

        Log.WriteLine(builder.GetDeleteCommand().CommandText)
        Log.WriteLine(builder.GetInsertCommand().CommandText)

        Return adapter
    End Function

    Public Sub InsertTable(sql As String)
        Dim cmd As OleDbCommand = CreateCommand(sql)

        cmd.ExecuteNonQuery()

    End Sub
    'Public Function SelectTable(sql As String) As DataTable
    '    Dim cmd As OleDbCommand = CreateCommand(sql)

    '    Dim adapter As New OleDbDataAdapter(cmd)
    '    Dim table As New DataTable

    '    'adapter.SelectCommand = cmd
    '    'adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
    '    adapter.Fill(table)

    '    Return table
    'End Function

    'Public Sub SaveOrUpdate(sql As String)
    '    Dim cmd As OleDbCommand = CreateCommand(sql)

    '    Dim adapter As New OleDbDataAdapter(cmd)
    '    With adapter
    '        .SelectCommand = cmd
    '        .MissingSchemaAction = MissingSchemaAction.AddWithKey

    '        Dim builder As New OleDbCommandBuilder(adapter)

    '        .DeleteCommand = builder.DataAdapter.DeleteCommand
    '        .InsertCommand = builder.DataAdapter.InsertCommand
    '        .UpdateCommand = builder.DataAdapter.UpdateCommand
    '    End With
    'End Sub
    'Public Function SelectTable(tableName As String) As DataTable
    '    Dim sql As String = String.Format("select * from {0}", tableName)
    '    Log.WriteLine("SelectTable -> [{0}]", sql)

    '    Dim cmd As OleDbCommand = CreateCommand(sql)

    '    Dim adapter As New OleDbDataAdapter(cmd)
    '    Dim table As New DataTable

    '    adapter.Fill(table)

    '    Return table
    'End Function

    'Public Function SelectDistinctColumn(tableName As String, columnName As String) As DataTable
    '    Dim sql As String = String.Format("select distinct {0} from {1} order by {0}", columnName, tableName)
    '    Log.WriteLine("SelectTable -> [{0}]", sql)

    '    Dim cmd As OleDbCommand = CreateCommand(sql)

    '    Dim adapter As New OleDbDataAdapter(cmd)
    '    Dim table As New DataTable

    '    adapter.Fill(table)

    '    Return table
    'End Function

    Private Function CreateCommand(sql As String) As OleDbCommand
        Dim cmd As New OleDbCommand(sql, conn)

        If Not IsNothing(tran) Then
            cmd.Transaction = tran
        End If

        Return cmd
    End Function
End Class
