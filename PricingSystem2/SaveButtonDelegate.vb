Imports System.Data.OleDb
Public Class SaveButtonDelegate

    Private sql As String
    Private table As DataTable
    'Private view As DataGridView

    Public Sub New(tableName As String, table As DataTable)
        Me.sql = String.Format("select * from {0}", tableName)
        Me.table = table
        'Me.view = view
    End Sub

    Public Sub SaveButtonDelegate(sender As Object, e As EventArgs)

        'For i As Integer = 0 To view.Rows.Count - 2
        '    Dim row As DataGridViewRow = view.Rows(i)

        '    With row
        '        .Cells(ITEM_ID_COLUMN).Value = Guid.NewGuid.ToString
        '        .Cells(MATAURE_DATE_COLUMN).Value = FormUtils.FormatDate(ImportForm.DateTimePicker1.Value)
        '        .Cells(PROJECT_NAME_COLUMN).Value = FormUtils.ProjectName
        '        .Cells(BID_COMPANY).Value = MainForm.ComName
        '    End With
        'Next

        For Each row As DataRow In table.Rows
            row(ITEM_ID_COLUMN) = GetValue(row(ITEM_ID_COLUMN), Guid.NewGuid.ToString)
            row(MATAURE_DATE_COLUMN) = GetValue(row(MATAURE_DATE_COLUMN), FormUtils.FormatDate(ImportForm.DateTimePicker1.Value))
            row(PROJECT_NAME_COLUMN) = GetValue(row(PROJECT_NAME_COLUMN), FormUtils.ProjectName)
            row(BID_COMPANY) = GetValue(row(BID_COMPANY), MainForm.ComName)
        Next

        Using conn As New OleDbConnection(CONNECTION_STRING)
            Dim dao As New AccessDao(conn)

            Dim adapter As OleDbDataAdapter = dao.SelectTableForUpdate(sql)

            'Log.WriteLine(table)
            adapter.Update(table)

            FormUtils.ShowInfoMessage(REPORT_DATA_UPDATE_DONE)
        End Using
    End Sub

    Private Function GetValue(source As Object, target As Object) As Object
        Dim ret As Object

        If DBNull.Value.Equals(source) Then
            ret = target
        Else
            ret = source
        End If

        Return ret
    End Function
End Class
