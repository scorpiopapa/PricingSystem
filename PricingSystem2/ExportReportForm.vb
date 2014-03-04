Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Public Class ExportReportForm
    Private Sub ExportReportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()

        Using conn As New OleDbConnection(CONNECTION_STRING)
            conn.Open()

            Dim sql As String = String.Format("select distinct {0} from {1} order by {0}", ReportMasterTable.REPORT_NAME, ReportMasterTable.TABLE_NAME)
            Dim cmd As New OleDbCommand(sql, conn)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim table As New DataTable

            adapter.Fill(table)
            With ComboBox1
                .DisplayMember = ReportMasterTable.REPORT_NAME
                .DataSource = table
            End With
        End Using
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ListBox1.Items.Count = 0 Then
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
        For i As Integer = 0 To ListBox1.Items.Count - 1
            reportNames.Add(ListBox1.Items(i))
        Next

        FormUtils.ExportReportTemplate(reportNames, fileName)

        'Try
        '            Using conn As New OleDbConnection(CONNECTION_STRING)
        '                excelApp = New Excel.Application
        '#If DEBUG Then
        '                excelApp.Visible = True
        '#End If

        '                wb = excelApp.Workbooks.Open(BLANK_FILE, [ReadOnly]:=True)
        '                For i As Integer = 0 To ListBox1.Items.Count - 1
        '                    Dim sql As String = String.Format("select {0} from {1} where {2} = '{3}' order by {4}", _
        '                                                      ReportMasterTable.COLUMN_NAME, ReportMasterTable.TABLE_NAME, _
        '                                                      ReportMasterTable.REPORT_NAME, ListBox1.Items(i), ReportMasterTable.ORDER)
        '                    Dim cmd As New OleDbCommand(sql, conn)
        '                    Dim adapter As New OleDbDataAdapter(cmd)
        '                    Dim table As New DataTable

        '                    adapter.Fill(table)

        '                    Dim sht As Excel.Worksheet

        '                    If i = 0 Then
        '                        sht = wb.Worksheets(1)
        '                    Else
        '                        sht = wb.Worksheets.Add
        '                    End If

        '                    With sht
        '                        .Name = ListBox1.Items(i)

        '                        .Cells(FormUtils.REPORT_NAME_ROW, START_COLUMN).value = .Name

        '                        Dim endColumn As Integer = START_COLUMN + table.Rows.Count

        '                        Dim r As Excel.Range = .Range(.Cells(FormUtils.REPORT_NAME_ROW, START_COLUMN), .Cells(FormUtils.REPORT_NAME_ROW, endColumn))
        '                        r.Merge()
        '                        r.HorizontalAlignment = Excel.Constants.xlCenter
        '                        r.Font.Bold = True
        '                        r.Font.Size = REPORT_NAME_FONT_SIZE

        '                        Dim bidPriceColumn As Integer
        '                        Dim offerPriceColumn As Integer

        '                        For j As Integer = 0 To table.Rows.Count
        '                            Dim curColumn As Integer = START_COLUMN + j
        '                            r = .Cells(HEAD_ROW, curColumn)

        '                            If j = 0 Then
        '                                r.Value = NO_TEXT
        '                            Else
        '                                r.Value = table.Rows(j - 1)(0)
        '                            End If
        '                            r.Font.Bold = True
        '                            r.HorizontalAlignment = Excel.Constants.xlCenter

        '                            If r.Value = BID_PRICE_TEXT Then
        '                                bidPriceColumn = curColumn
        '                            ElseIf r.Value = OFFER_PRICE_TEXT Then
        '                                offerPriceColumn = curColumn
        '                            End If

        '                            r = .Columns(curColumn)
        '                            r.EntireColumn.AutoFit()
        '                        Next

        '                        r = .Cells(MARKER_ROW, START_COLUMN)
        '                        r.Value = MARKER_TEXT1
        '                        r.HorizontalAlignment = Excel.Constants.xlLeft

        '                        r = .Cells(MARKER_ROW, endColumn)
        '                        r.Value = MARKER_TEXT2
        '                        r.HorizontalAlignment = Excel.Constants.xlRight

        '                        r = .Cells(SUMMARY_ROW, START_COLUMN + 1)
        '                        r.Value = SUMMARY_TEXT

        '                        AddSumFormula(.Cells(SUMMARY_ROW, bidPriceColumn))
        '                        AddSumFormula(.Cells(SUMMARY_ROW, offerPriceColumn))

        '                        r = .Range(.Cells(HEAD_ROW, START_COLUMN), .Cells(SUMMARY_ROW, endColumn))
        '                        FormUtils.DrawGrid(r)
        '                    End With
        '                Next
        '            End Using

        '            wb.SaveAs(fileName)
        '            wb.Close()
        '            excelApp.Quit()

        '            With My.Computer.FileSystem
        '                .CopyFile(DB_FILE, .GetParentPath(fileName) + "\" + .GetName(DB_FILE), True)
        '            End With
        '        Catch ex As Exception
        '            Log.WriteLine(ex)

        '            If Not IsNothing(wb) Then
        '                wb.Close(False)
        '            End If

        '            If Not IsNothing(excelApp) Then
        '                excelApp.Quit()
        '            End If

        '            ShowErrorMessage(ex.Message)
        '        Finally
        '            wb = Nothing
        '            excelApp = Nothing
        '        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not ListBox1.Items.Contains(ComboBox1.Text) Then
            ListBox1.Items.Add(ComboBox1.Text)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'With ListBox1
        '    If .SelectedItems.Count > 0 Then

        '    End If
        'End With
    End Sub
End Class