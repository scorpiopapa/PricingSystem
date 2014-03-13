Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Public Class ExportReportForm
    Private Sub ExportReportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ListView1
            .Items.Clear()
            .View = View.Details
            .FullRowSelect = True
            .GridLines = True
            .MultiSelect = True
            .Columns(0).Width = .Width
        End With

        Using conn As New OleDbConnection(CONNECTION_STRING)
            conn.Open()

            Dim sql As String = String.Format("select distinct {0} from {1} order by {0}", ReportMasterTable.REPORT_NAME, ReportMasterTable.TABLE_NAME)
            Dim cmd As New OleDbCommand(sql, conn)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim table As New DataTable

            adapter.Fill(table)

            For Each row As DataRow In table.Rows
                Dim item As New ListViewItem(row(0).ToString)
                ListView1.Items.Add(item)
            Next
            'With ComboBox1
            '    .DisplayMember = ReportMasterTable.REPORT_NAME
            '    .DataSource = table
            'End With
        End Using
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ListView1.SelectedItems.Count = 0 Then
            FormUtils.ShowErrorMessage(EMPTY_REPORT_LIST)
            Exit Sub
        End If

        Dim fileName As String = ""
        'Dim excelApp As Excel.Application = Nothing
        'Dim wb As Excel.Workbook = Nothing

        With SaveFileDialog1
            .Filter = ExcelFilter()

            Dim result As DialogResult = .ShowDialog()
            If result <> Windows.Forms.DialogResult.OK Then
                Exit Sub
            End If

            fileName = .FileName
        End With

        Dim reportNames As New List(Of String)
        For i As Integer = 0 To ListView1.SelectedItems.Count - 1
            reportNames.Add(ListView1.SelectedItems(i).Text)
        Next

        FormUtils.ExportReportTemplate(reportNames, fileName)

    End Sub


End Class