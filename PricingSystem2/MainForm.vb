﻿Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Public Class MainForm

    Public LoginUserName As String
    Public LoginForm As Form
    Public ComName As String

    Private Shared ReadOnly TITLE = "报价管理系统"

    Private Sub 明细表管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 明细表管理ToolStripMenuItem.Click
        ShowForm(AddReportForm)

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        UpdateReportForm.ShowDialog(Me)

    End Sub

    Private Sub MainForm_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        LoginForm.Dispose()
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Label1.Text = "操作员：" + LoginUserName
            Label2.Text = "单位名称：" + ComName

            If TypeOf (LoginForm) Is SellerLogin Then
                ReportManage.Visible = False
                Me.Text = TITLE + "-报价单位版"
            Else
                GenerateReportTemplate.Visible = False
                Me.Text = TITLE + "-审价单位版"
            End If

            With TreeView1
                .HideSelection = False
            End With

            With DataGridView1
                .ReadOnly = True
            End With

            Log.WriteLine("form load - before refresh tree")
            RefreshTreeView()
        Catch ex As Exception
            Log.WriteLine(ex)
        End Try

    End Sub

    Private Sub InitTreeView()

    End Sub

    Private Sub ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem.Click
        ShowForm(ExportReportForm)

    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        ShowForm(AddSummaryReportForm)
    End Sub


    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        ShowForm(UpdateSummaryReportForm)

    End Sub

    Private Sub ShowForm(f As Form)
        With f
            .StartPosition = FormStartPosition.CenterParent
            .MaximizeBox = False
            .ShowDialog(Me)
        End With
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        SwitchTabPage()
        SearchData()
    End Sub

    Private Sub TreeView1_DoubleClick(sender As Object, e As EventArgs) Handles TreeView1.DoubleClick
        SwitchTabPage()
        SearchData()
    End Sub

    Private Sub SwitchTabPage()
        If TreeView1.SelectedNode.Text.Trim = REPORT_SUMMARY_TEXT Then
            TabControl1.SelectedIndex = 0
        ElseIf TreeView1.SelectedNode.Nodes.Count = 0 Then
            With TabControl1
                Dim title As String = TreeView1.SelectedNode.Text
                For Each t As TabPage In .TabPages
                    If title = t.Text Then
                        .SelectedTab = t
                        Exit Sub
                    End If
                Next
                Dim tp As TabPage = CreateTabPage(title)
                .TabPages.Add(tp)
                AddTabControls(tp)
                .SelectedTab = tp
            End With
        End If
    End Sub

    Private Sub SearchData()
        If TreeView1.SelectedNode.Index = 0 Then
            Exit Sub
        End If

        Dim baseSql As String = ""
        Dim sql As String
        'Dim hasReport As Boolean = True

        Using conn As New OleDbConnection(CONNECTION_STRING)
            sql = String.Format("select distinct {0} from {1}", ReportMasterTable.REPORT_NAME, ReportTreeTable.TABLE_NAME)
            Dim dao As New AccessDao(conn)
            Dim table As DataTable = dao.SelectTable(sql)

            For i As Integer = 0 To table.Rows.Count - 1
                baseSql = baseSql & String.Format("select {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7} from {8}", _
                                                  MATAURE_DATE_COLUMN, PROJECT_NAME_COLUMN, BID_COMPANY, ITEM_NAME_COLUMN, _
                                                  BID_PRICE_COLUMN, BID_COMMENT_COLUMN, OFFER_PRICE_COLUMN, OFFER_COMMENT_COLUMN,
                                                  table.Rows(0)(0))
                If i < table.Rows.Count - 1 Then
                    baseSql = baseSql & " union "
                End If
            Next
        End Using

        If baseSql <> "" Then
            With DataGridView1
                For i As Integer = .Columns.Count - 1 To 0 Step -1
                    .Columns.RemoveAt(i)
                Next
            End With

            Using conn As New OleDbConnection(CONNECTION_STRING)
                Dim dao As New AccessDao(conn)
                Dim table As DataTable = Nothing

                With TreeView1
                    'If .SelectedNode.Index = 0 Then
                    '    table = dao.SelectTable(baseSql)
                    '    DataGridView1.DataSource = table
                    'End If
                End With
            End Using
        End If
    End Sub

    Private Function CreateTabPage(title As String) As TabPage
        Dim tp As New TabPage(title)

        Dim view As New DataGridView
        Dim btn As New Button With {.Text = "保存"}
        Dim deleteButton As New Button With {.Text = "删除"}
        Dim moveUpButton As New Button With {.Text = "上移"}
        Dim moveDownButton As New Button With {.Text = "下移"}
        Dim group As New Panel

        With tp.Controls
            .Add(view)
            .Add(group)
            group.Controls.Add(btn)
            group.Controls.Add(deleteButton)
            group.Controls.Add(moveUpButton)
            group.Controls.Add(moveDownButton)
        End With

        view.Dock = DockStyle.Fill
        group.Width = 100
        group.Dock = DockStyle.Right

        Dim padding As Integer = 20

        btn.Top = padding
        btn.Width = 65
        btn.Left = (group.Width - btn.Width) / 2

        With deleteButton
            .Top = btn.Top + btn.Height + padding
            .Width = btn.Width
            .Left = btn.Left
        End With

        With moveUpButton
            .Top = deleteButton.Top + deleteButton.Height + padding
            .Width = btn.Width
            .Left = btn.Left
        End With

        With moveDownButton
            .Top = moveUpButton.Top + moveUpButton.Height + padding
            .Width = btn.Width
            .Left = btn.Left
        End With

        Dim cs As New List(Of Object)
        With cs
            .Add(view)
            .Add(btn)
            '.Add(deleteButton)
        End With

        tp.Tag = cs

        Return tp
    End Function

    Private Sub AddTabControls(tp As TabPage)
        Dim table As DataTable

        Using conn As New OleDbConnection(CONNECTION_STRING)
            Dim dao As New AccessDao(conn)
            Dim sql As String = String.Format("select * from {0}", tp.Text)
            table = dao.SelectTable(sql)

            'view.DataSource = table

            If table.Rows.Count = 0 Then
                'table.NewRow()
                'Dim sql2 As String = String.Format("select {0}, {1} from {2} where {3} = '{4}' order by {5}", _
                '                            ReportMasterTable.COLUMN_NAME, ReportMasterTable.COLUMN_TYPE, ReportMasterTable.TABLE_NAME, _
                '                            ReportMasterTable.REPORT_NAME, tp.Text, ReportMasterTable.ORDER)
                'Dim table2 As DataTable = dao.SelectTable(sql2)

                'With view
                '    .Columns.Add(New DataGridViewTextBoxColumn() With {.HeaderText = ITEM_ID, .DataPropertyName = ITEM_ID, .ValueType = DBConstants.INT})

                '    For Each row As DataRow In table2.Rows
                '        Dim dbvalue = row(0)
                '        Dim c As New DataGridViewTextBoxColumn() With {.HeaderText = dbvalue, .DataPropertyName = row(0), .ValueType = DBUtils.GetDataColumnType(row(1))}
                '        .Columns.Add(c)
                '    Next
                'End With
            End If
        End Using

        Dim cs As List(Of Object) = tp.Tag
        Dim view As DataGridView = CType(cs(0), DataGridView)
        Dim btn As Button = CType(cs(1), Button)

        With view
            .DataSource = table
            .Columns(ITEM_ID_COLUMN).Visible = False
            '到期日
            .Columns(MATAURE_DATE_COLUMN).Visible = False
            '所属项目
            .Columns(PROJECT_NAME_COLUMN).Visible = False
            '报价单位
            .Columns(BID_COMPANY).Visible = False
        End With

        Dim handler As New SaveButtonDelegate(tp.Text, table)
        AddHandler btn.Click, AddressOf handler.SaveButtonDelegate
    End Sub

    Private Sub TreeView1_MouseUp(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseUp
        With TreeView1
            If e.Button = Windows.Forms.MouseButtons.Right AndAlso CanExport() Then
                .SelectedNode.ContextMenuStrip = ContextMenuStrip1
            End If
        End With
    End Sub

    Private Function CanExport() As Boolean
        Dim ret As Boolean = False

        With TreeView1
            If .SelectedNode.Nodes.Count = 0 Then
                ' report name selected
                ret = True
            ElseIf .SelectedNode.FullPath.Split("\").Count = 4 Then
                ' project name selected
                ret = True
            End If
        End With

        Return ret
    End Function
    Private Sub RefreshTreeView()
        Using conn As New OleDbConnection(CONNECTION_STRING)
            Dim sql As String = String.Format("select * from {0} order by {1} desc, {2}, {3}, {4}", _
                                              ReportTreeTable.TABLE_NAME, ReportTreeTable.YEAR, _
                                              ReportTreeTable.PROJECT_NAME, ReportTreeTable.COMPANY_NAME, ReportTreeTable.REPORT_NAME)
            Log.WriteLine("execute sql {0}", sql)

            Dim dao As New AccessDao(conn)
            Dim table As DataTable = dao.SelectTable(sql)

            Dim data As New List(Of List(Of String))
            For Each row As DataRow In table.Rows
                Dim l As New List(Of String)
                With l
                    .Add(row(ReportTreeTable.YEAR))
                    .Add(row(ReportTreeTable.COMPANY_NAME))
                    .Add(row(ReportTreeTable.PROJECT_NAME))
                    .Add(row(ReportTreeTable.REPORT_NAME))
                End With

                data.Add(l)
            Next

            Log.WriteLine("generate tree...")
            With TreeView1
                For Each n As TreeNode In .Nodes(0).Nodes
                    n.Remove()
                Next

                Dim years As Dictionary(Of String, List(Of List(Of String))) = Group(data)

                ' add year
                For Each y As String In years.Keys
                    Dim yearNode As TreeNode = .Nodes(0).Nodes.Add(y)
                    Dim companyValues As List(Of List(Of String)) = years.Item(y)
                    Dim companies As Dictionary(Of String, List(Of List(Of String))) = Group(companyValues)

                    ' add company
                    For Each company As String In companies.Keys
                        Dim companyNode As TreeNode = yearNode.Nodes.Add(company)
                        Dim projectValues As List(Of List(Of String)) = companies.Item(company)
                        Dim projects As Dictionary(Of String, List(Of List(Of String))) = Group(projectValues)

                        ' add project
                        For Each project As String In projects.Keys
                            Dim projectNode As TreeNode = companyNode.Nodes.Add(project)
                            Dim reportValues As List(Of List(Of String)) = projects.Item(project)
                            Dim reports As Dictionary(Of String, List(Of List(Of String))) = Group(reportValues)

                            ' add report
                            For Each report As String In reports.Keys
                                projectNode.Nodes.Add(report)
                            Next
                        Next
                    Next
                Next

                .ExpandAll()
                .SelectedNode = .Nodes(0)
            End With
            Log.WriteLine("generate tree done")
        End Using
    End Sub

    'rem 导入模板
    Private Sub ToolStripMenuItem0_Click(sender As Object, e As EventArgs) Handles GenerateReportTemplate.Click
        ImportForm.ShowDialog(Me)

        Dim year As Integer = FormUtils.ReportYear
        Dim projectName As String = FormUtils.ProjectName

        Dim fileName As String

        With OpenFileDialog1
            .Filter = ExcelFilter()

            Dim result As DialogResult = .ShowDialog()

            If result <> Windows.Forms.DialogResult.OK Then
                Exit Sub
            End If

            fileName = .FileName
        End With

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application
            wb = excelApp.Workbooks.Open(fileName)

            UpdateReportTreeTable(wb, year, projectName, ComName)

            wb.Close(False)
            excelApp.Quit()

            wb = Nothing
            excelApp = Nothing

            RefreshTreeView()

        Catch ex As Exception
            Log.WriteLine(ex)

            FormUtils.ShowErrorMessage(ex.Message)

            If Not IsNothing(wb) Then
                wb.Close(False)
                wb = Nothing
            End If

            If Not IsNothing(excelApp) Then
                excelApp.Quit()
                excelApp = Nothing
            End If
        End Try
    End Sub

    Private Sub UpdateReportTreeTable(wb As Excel.Workbook, year As Integer, Optional pname As String = Nothing, Optional cname As String = Nothing)
        Log.WriteLine("start to update report tree table...")
        Dim tran As OleDbTransaction = Nothing

        Using conn As New OleDbConnection(CONNECTION_STRING)
            Try
                conn.Open()
                tran = conn.BeginTransaction

                Dim dao As New AccessDao(conn, tran)
                Dim sql As String

                For Each sht As Excel.Worksheet In wb.Worksheets
                    sql = String.Format("select * from {0} where year={1} and company_name='{2}' and project_name='{3}' and report_name='{4}'", _
                                        ReportTreeTable.TABLE_NAME, year, cname, pname, sht.Name)
                    Log.WriteLine("execute sql {0}", sql)
                    Dim table As DataTable = dao.SelectTable(sql)

                    If table.Rows.Count = 0 Then
                        sql = String.Format("insert into {0} values({1},'{2}','{3}','{4}')", _
                                                          ReportTreeTable.TABLE_NAME, year, cname, pname, sht.Name)
                        Log.WriteLine("execute insert table by {0}", sql)
                        dao.InsertTable(sql)
                    End If
                Next

                Log.WriteLine("before commit")
                tran.Commit()
            Catch ex As Exception
                If Not IsNothing(tran) Then
                    tran.Rollback()
                End If

                Throw ex
            End Try

        End Using

        Log.WriteLine("end of update report tree table...")
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ExpotReportMenu.Click
        ExportReport()
    End Sub

    Private Sub ExportReport()
        Dim folderName As String

        With FolderBrowserDialog1
            Dim result As DialogResult = .ShowDialog()

            If result <> Windows.Forms.DialogResult.OK Then
                Exit Sub
            End If

            folderName = .SelectedPath
        End With

        Dim reportNames As List(Of String) = FormUtils.FindReportNames(TreeView1.SelectedNode)

        ExportReportData(folderName, reportNames)
    End Sub

    Private Sub ExportReportData(folderName As String, names As List(Of String))
        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing
        'Dim sht As Excel.Worksheet = Nothing

#If Not Debug Then
        Try
#End If
        'Dim projectName As String = TreeView1.SelectedNode.FullPath.Split("\")(3)
        Dim projectName As String = FormUtils.ProjectName

        Dim fileName As String

        If TypeOf LoginForm Is SellerLogin Then
            fileName = folderName + "\" + String.Format("{0} {1} {2}年度报价.xls", ComName, projectName, FormUtils.ReportYear)
        Else
            Dim ns() As String = TreeView1.SelectedNode.FullPath.Split("\")
            fileName = folderName + "\" + String.Format("{0} {1} {2}年度报价.xls", ns(2), ns(3), FormUtils.ReportYear)
        End If

        FormUtils.ExportReportTemplate(names, fileName, False)

        excelApp = New Excel.Application
#If DEBUG Then
        excelApp.Visible = True
#End If
        wb = excelApp.Workbooks.Open(fileName)

        Using conn As New OleDbConnection(CONNECTION_STRING)
            Dim sql As String
            Dim dao As New AccessDao(conn)
            Dim table As DataTable

            For i As Integer = 0 To names.Count - 1
                sql = String.Format("select {0} from {1} where {2}='{3}' order by {4}", _
                                ReportMasterTable.COLUMN_NAME, ReportMasterTable.TABLE_NAME, _
                                ReportMasterTable.REPORT_NAME, names(i), ReportMasterTable.ORDER)
                table = dao.SelectTable(sql)

                Dim elements As New List(Of String)
                For Each row As DataRow In table.Rows
                    elements.Add(row(0))
                Next

                Dim clause As String = DBUtils.BuildClause(elements)
                clause = ITEM_ID_COLUMN + "," + clause
                sql = String.Format("select {0} from {1}", clause, names(i))
                table = dao.SelectTable(sql)

                Dim sht As Excel.Worksheet = wb.Worksheets(names(i))
                Dim curRow As Integer
                For j As Integer = 0 To table.Rows.Count - 1
                    Dim row As DataRow = table.Rows(j)
                    curRow = DATA_START_ROW + j

                    With sht
                        .Cells(DATA_START_ROW + j, EXCEL_ITEM_ID_COLUMN).value = row(ITEM_ID_COLUMN)
                        .Cells(curRow, START_COLUMN).value = j + 1

                        For k As Integer = 1 To table.Columns.Count - 1
                            .Cells(curRow, START_COLUMN + k).value = row(k)
                        Next
                    End With
                Next

                Dim r As Excel.Range = sht.Range(sht.Cells(DATA_START_ROW, START_COLUMN), sht.Cells(curRow, START_COLUMN + table.Columns.Count - 1))
                FormUtils.DrawGrid(r)
            Next
        End Using

        wb.Close(True)
        excelApp.Quit()

        wb = Nothing
        excelApp = Nothing

        With My.Computer.FileSystem
            Dim newName As String = fileName.Split(".")(0)
            .CopyFile(fileName, newName, True)
        End With

        FormUtils.ShowInfoMessage(FormUtils.REPORT_DATA_EXPORT_DONE)
#If Not Debug Then
        Catch ex As Exception
            Log.WriteLine(ex)

            FormUtils.ShowErrorMessage(ex.Message)

            If Not IsNothing(wb) Then
                wb.Close(False)
                wb = Nothing
            End If

            If Not IsNothing(excelApp) Then
                excelApp.Quit()
                excelApp = Nothing
            End If
        End Try
#End If
    End Sub


    Private Sub ImportDataMenu_Click(sender As Object, e As EventArgs) Handles ImportDataMenu.Click
        Log.WriteLine("import report data...")
        Dim fileName As String

        With OpenFileDialog1
            .Filter = ExcelFilter()

            Dim result As DialogResult = .ShowDialog()

            If result <> Windows.Forms.DialogResult.OK Then
                Exit Sub
            End If

            fileName = .FileName
        End With

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing
        Dim tran As OleDbTransaction = Nothing

#If Not Debug Then
        Try
#End If
            excelApp = New Excel.Application
#If DEBUG Then
            excelApp.Visible = True
#End If
            wb = excelApp.Workbooks.Open(fileName, [ReadOnly]:=True)
            Log.WriteLine("open excel for import data")

            'Try


            Using conn As New OleDbConnection(CONNECTION_STRING)
                conn.Open()
                tran = conn.BeginTransaction

                Dim dao As New AccessDao(conn, tran)
                Dim sql As String

                For Each sht As Excel.Worksheet In wb.Worksheets
                    sql = String.Format("select * from {0}", sht.Name)
                    Log.WriteLine("execute sql {0}", sql)

                    Dim table As DataTable = dao.SelectTable(sql)

                    Dim itemIds As New List(Of String)
                    For Each row As DataRow In table.Rows
                        itemIds.Add(row(ITEM_ID_COLUMN))
                    Next

                    Dim excelItemIds As New List(Of String)

                    With sht
                        Dim curRow As Integer = DATA_START_ROW

                        While Trim(.Cells(curRow, EXCEL_ITEM_ID_COLUMN).value) <> ""
                            Dim itemId As String = Trim(.Cells(curRow, EXCEL_ITEM_ID_COLUMN).value)
                            excelItemIds.Add(itemId)

                            Dim condition As String = String.Format("{0}='{1}'", ITEM_ID_COLUMN, itemId)
                            Dim targetRows() As DataRow = table.Select(condition)

                            Dim targetRow As DataRow
                            If IsNothing(targetRows) OrElse targetRows.Length = 0 Then
                                targetRow = table.NewRow
                                targetRow(ITEM_ID_COLUMN) = itemId
                                table.Rows.Add(targetRow)
                            Else
                                targetRow = targetRows(0)
                            End If

                            Dim sql2 As String = String.Format("select count(*) from {0} where {1}='{2}'", _
                                                               ReportMasterTable.TABLE_NAME, ReportMasterTable.REPORT_NAME, .Name)
                            Dim columnCount As Integer = dao.SelectTable(sql2).Rows(0)(0)

                            For i As Integer = START_COLUMN + 1 To START_COLUMN + columnCount
                                Dim columnName As String = Trim(.Cells(HEAD_ROW, i).value)
                                targetRow(columnName) = DBUtils.ToDBValue(Trim(.Cells(curRow, i).value))
                                Log.WriteLine("set column {0} value to [{1}]", columnName, DBUtils.ToDBValue(Trim(.Cells(curRow, i).value)).ToString)
                            Next

                            curRow = curRow + 1
                        End While
                    End With

                    For Each dbId As String In itemIds
                        If Not excelItemIds.Contains(dbId) Then
                            ' remove row
                            For i As Integer = table.Rows.Count - 1 To 0 Step -1
                                Log.WriteLine(table)
                                If table.Rows(i)(ITEM_ID_COLUMN) = dbId Then
                                    table.Rows(i).Delete()
                                End If
                            Next
                        End If
                    Next

                    Log.WriteLine("before update table")
                    Dim adapter As OleDbDataAdapter = dao.SelectTableForUpdate(sql)
                    adapter.Update(table)
                Next

                Log.WriteLine("before commit")
                tran.Commit()
            End Using
            'Catch ex As Exception
            '    Log.WriteLine(ex)
            'End Try

            Log.WriteLine("parsing tree...")
            Dim parts() As String = fileName.Split(" ")
            Log.WriteLine("parts is")
            Log.WriteLine(parts)

            Dim year As Integer = CInt(parts(parts.Length - 1).Substring(0, 4))
            Log.WriteLine("year is {0}", year)

            Dim c() As String = parts(parts.Length - 3).Split("\")
            Dim cname As String = c(c.Length - 1)

            Log.WriteLine("before update report tree table")
            UpdateReportTreeTable(wb, year, parts(parts.Length - 2), cname)
            RefreshTreeView()

            FormUtils.ShowInfoMessage(REPORT_DATA_IMPORT_DONE)

            wb.Close(False)
            excelApp.Quit()

            wb = Nothing
            excelApp = Nothing

#If Not Debug Then
        Catch ex As Exception
            Log.WriteLine(ex)

            'If Not IsNothing(tran) Then
            '    tran.Rollback()
            'End If
            FormUtils.ShowErrorMessage(ex.Message)

            If Not IsNothing(wb) Then
                wb.Close(False)
                wb = Nothing
            End If

            If Not IsNothing(excelApp) Then
                excelApp.Quit()
                excelApp = Nothing
            End If
        End Try
#End If

    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs)
        'SwitchTabPage()
        'TabControl1.SelectedIndex = 0
        'TreeView1.SelectedNode = TreeView1.Nodes(0)
    End Sub

    Private Sub 导出ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ExportDataToolStripMenuItem.Click
        If IsNothing(TreeView1.SelectedNode) OrElse Not CanExport() Then
            FormUtils.ShowErrorMessage(EMPTY_EXPORT_ITEM)
            Exit Sub
        End If


        ExportReport()

    End Sub

    Private Sub DataQueryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataQueryToolStripMenuItem.Click
        TreeView1.SelectedNode = TreeView1.Nodes(0)
    End Sub

    Private Sub 发布模板ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 发布模板ToolStripMenuItem.Click

    End Sub
End Class
