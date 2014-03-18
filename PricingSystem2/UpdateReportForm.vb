Imports System.Data.OleDb
Public Class UpdateReportForm

    Private Shared ReadOnly ADD_TEXT = "添加"
    Private Shared ReadOnly UPDATE_TEXT = "修改"
    Private Shared ReadOnly BIND_COLUMN_NAME = "report_name"

    Private sql As String = "select * from report_master where report_name = '{0}'"

    Private ReadOnly Property DataGridViewSql
        Get
            Return String.Format(sql, ComboBox1.Text.Trim)
        End Get
    End Property
    Private Sub ReportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.Button2.Text = ADD_TEXT

        RefreshComboBox()
        RefreshDataGridView()

    End Sub

    Private Sub RefreshDataGridView()
        Dim table As New DataTable

        Using conn As New OleDbConnection(CONNECTION_STRING)
            'conn.Open()

            'Dim cmd As New OleDbCommand(DataGridViewSql, conn)
            'Dim adapter As New OleDbDataAdapter(cmd)

            Dim dao As New AccessDao(conn)

            table = dao.SelectTable(DataGridViewSql)
            'If table.Rows.Count = 0 Then
            '    table.NewRow()
            'End If

            'adapter.Fill(table)

            With DataGridView1
                .DataSource = table
                .Columns(0).Visible = False
                .Columns(1).Visible = False
            End With
        End Using
    End Sub

    Private Sub RefreshComboBox()
        Using conn As New OleDbConnection(CONNECTION_STRING)
            Dim table As New DataTable

            Dim sql As String = "select distinct report_name from report_master order by report_name"
            Dim cmd As New OleDbCommand(sql, conn)
            Dim adapter As New OleDbDataAdapter(cmd)

            adapter.Fill(table)

            'table.Rows.Add("")
            FormUtils.BindComboBoxData(Me.ComboBox1, table, ReportMasterTable.REPORT_NAME)

        End Using
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Log.WriteLine("[" + ComboBox1.Text + "]")

        'If ignoreSelectionChange Then
        '    ignoreSelectionChange = False
        '    Exit Sub
        'End If

        'ignoreSelectionChange = True
        RefreshDataGridView()
    End Sub


    'update
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Using conn As New OleDbConnection(CONNECTION_STRING)
            conn.Open()

            Dim cmd As New OleDbCommand(DataGridViewSql, conn)
            Dim adapter As New OleDbDataAdapter(cmd) With {.MissingSchemaAction = MissingSchemaAction.AddWithKey}
            Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)

            Dim table As DataTable = DataGridView1.DataSource
            'Dim table As New DataTable

            For Each row As DataRow In table.Rows
                row(ReportMasterTable.REPORT_NAME) = ComboBox1.Text
            Next

            adapter.Update(table)
        End Using

        RefreshDataGridView()

        'ignoreSelectionChange = True
        RefreshComboBox()

        FormUtils.ShowInfoMessage(FormUtils.REPORT_UPDATE_DONE)
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        'With ComboBox1
        '    If .Text.Trim = "" Then
        '        Me.Button2.Text = ADD_TEXT
        '    Else
        '        If .Items.Contains(.Text) Then
        '            Me.Button2.Text = UPDATE_TEXT
        '        Else
        '            Me.Button2.Text = ADD_TEXT
        '        End If
        '    End If
        'End With
    End Sub


End Class