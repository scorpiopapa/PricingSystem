Imports System.Data.OleDb
Public Class AddReportForm

    Private Shared ReadOnly DEFAULT_COLUMN_NAMES() As String = {"品目名称", "报价", "报价备注", "审价", "审价备注", MATAURE_DATE_COLUMN, PROJECT_NAME_COLUMN, BID_COMPANY}

    Private Sub AddReportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            RefreshDataGridView()
        Catch ex As Exception
            Log.WriteLine(ex)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If TextBox1.Text.Trim = "" Then
            FormUtils.ShowErrorMessage(FormUtils.EMPTY_REPORT_NAME)
            Exit Sub
        End If

        ' check report existents
        Using conn As New OleDbConnection(CONNECTION_STRING)
            'conn.Open()

            Dim dao As New AccessDao(conn)
            If dao.DataExists(ReportMasterTable.TABLE_NAME, ReportMasterTable.REPORT_NAME, TextBox1.Text.Trim, OleDbType.Char) Then
                FormUtils.ShowErrorMessage(FormUtils.REPORT_EXIST)
                Exit Sub
            End If
        End Using

        ' save report
        Using conn As New OleDbConnection(CONNECTION_STRING)
            Dim tran As OleDbTransaction = Nothing
#If Not Debug Then
            Try
#End If
                conn.Open()
            tran = conn.BeginTransaction

                Dim sql As String = String.Format("select * from report_master where report_name = '{0}'", TextBox1.Text)
                Dim cmd As New OleDbCommand(sql, conn, tran)
                Dim adapter As New OleDbDataAdapter(cmd) With {.MissingSchemaAction = MissingSchemaAction.AddWithKey}
                Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)

                Dim table As DataTable = DataGridView1.DataSource

                For Each row As DataRow In table.Rows
                    row(ReportMasterTable.REPORT_NAME) = TextBox1.Text
                Next

                adapter.Update(table)

                'create report table
                Dim columnNames As New List(Of String)
                Dim types As New List(Of String)
                For Each reportRow As DataGridViewRow In DataGridView1.Rows
                    Dim c As String = reportRow.Cells(ReportMasterTable.COLUMN_NAME).Value
                    If Not Empty(c) Then
                        columnNames.Add(c)
                    End If

                    Dim t As String = reportRow.Cells(ReportMasterTable.COLUMN_TYPE).Value
                    If Not Empty(t) Then
                        types.Add(t)
                    End If
                Next

                Dim clause As String = DBUtils.BuildClause(columnNames, types)
            'Dim ddl As String = String.Format("create table {0} (item_id autoincrement(1,1) primary key, {1})", TextBox1.Text, clause)
            Dim ddl As String = String.Format("create table {0} (item_id varchar(50) primary key, {1})", TextBox1.Text, clause)
            cmd = New OleDbCommand(ddl, conn, tran)
                cmd.ExecuteNonQuery()

                tran.Commit()

                FormUtils.ShowInfoMessage(FormUtils.REPORT_ADD_DONE)

                TextBox1.Text = ""

                RefreshDataGridView()
                'table.Rows.Clear()
#If Not Debug Then
            Catch ex As Exception
                Log.WriteLine(ex)
                If Not IsNothing(tran) Then
                    tran.Rollback()
                End If

                FormUtils.ShowErrorMessage(ex.Message)
            End Try
#End If
        End Using
    End Sub

    Private Sub RefreshDataGridView()
        Dim table As New DataTable

        Using conn As New OleDbConnection(CONNECTION_STRING)
            conn.Open()

            Dim cmd As New OleDbCommand("select * from report_master where report_name = ''", conn)
            Dim adapter As New OleDbDataAdapter(cmd)

            With table.Columns
                .Add(New DataColumn(ReportMasterTable.REPORT_ID, INT))
                .Add(New DataColumn(ReportMasterTable.REPORT_NAME, STR))
                .Add(New DataColumn(ReportMasterTable.COLUMN_NAME, STR))
                .Add(New DataColumn(ReportMasterTable.COLUMN_TYPE, STR))
                .Add(New DataColumn(ReportMasterTable.ORDER, INT))
                .Add(New DataColumn(ReportMasterTable.FORMULA, STR))
                .Add(New DataColumn(ReportMasterTable.SHOW, STR))
            End With

            Dim order As Integer = 1
            For Each Name As String In DEFAULT_COLUMN_NAMES
                Dim row As DataRow = table.NewRow()
                row(ReportMasterTable.COLUMN_NAME) = Name

                Dim type As String
                If Name = "报价" OrElse Name = "审价" Then
                    type = NUMBER_TYPE
                ElseIf Name = "到期日" Then
                    type = DATE_TYPE
                Else
                    type = TEXT_TYPE
                End If

                row(ReportMasterTable.COLUMN_TYPE) = type
                row(ReportMasterTable.ORDER) = order

                order = order + 1

                table.Rows.Add(row)
            Next

            adapter.Fill(table)

            With DataGridView1
                .DataSource = table
                .Columns(0).Visible = False
                .Columns(1).Visible = False
            End With
        End Using
    End Sub

End Class