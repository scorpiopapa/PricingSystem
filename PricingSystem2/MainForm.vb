Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Public Class MainForm

    '#If DEBUG Then
    '    Private REPORT_FILE As String = "..\..\..\报价单位A 项目1 2014年报价.xls"
    '    Private REPORT_FILE_BAK As String = "..\..\..\报价单位A 项目1 2014年报价"
    '#Else
    '    Private REPORT_FILE As String = Application.StartupPath + "\config\报价单位A 项目1 2014年报价.xls"
    '    Private REPORT_FILE_BAK As String = Application.StartupPath + "\config\报价单位A 项目1 2014年报价"
    '    Private TEMPLATE_FILE As String = Application.StartupPath + "\config\temp.xls"
    '#End If

    Public LoginUserName As String
    Public LoginForm As Form
    Public ComName As String

    'Private view As New DataGridView

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
        Label1.Text = LoginUserName
        Label2.Text = ComName

        If TypeOf (LoginForm) Is SellerLogin Then
            ReportManage.Visible = False
        Else
            ImportTemplate.Visible = False
        End If

        With TreeView1
            .HideSelection = False
            .ExpandAll()
            .SelectedNode = .Nodes(0)

            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(1).ContextMenuStrip = ContextMenuStrip1

            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(2).ContextMenuStrip = ContextMenuStrip1
            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(2).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(2).Nodes(1).ContextMenuStrip = ContextMenuStrip1

            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(3).ContextMenuStrip = ContextMenuStrip1
            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(3).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(3).Nodes(0).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            '.Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(3).Nodes(0).Nodes(1).ContextMenuStrip = ContextMenuStrip1
        End With
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
    End Sub

    Private Sub TreeView1_DoubleClick(sender As Object, e As EventArgs) Handles TreeView1.DoubleClick
        SwitchTabPage()
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
    Private Function CreateTabPage(title As String) As TabPage
        Dim tp As New TabPage(title)

        Dim view As New DataGridView
        Dim btn As New Button With {.Text = "保存"}
        Dim group As New Panel

        With tp.Controls
            .Add(view)
            .Add(group)
            group.Controls.Add(btn)
        End With

        view.Dock = DockStyle.Fill
        group.Width = 100
        group.Dock = DockStyle.Right

        btn.Top = 20
        btn.Width = 65
        btn.Left = (group.Width - btn.Width) / 2

        Dim cs As New List(Of Object)
        cs.Add(view)
        cs.Add(btn)
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

        Dim handler As New SaveButtonDelegate(tp.Text, table, view)
        AddHandler btn.Click, AddressOf handler.SaveButtonDelegate
    End Sub

    Private Sub SaveButtonDelegate(sender As Object, e As EventArgs)

    End Sub

    Private Sub TreeView1_MouseUp(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If TreeView1.SelectedNode.Nodes.Count = 0 Then
                TreeView1.SelectedNode.ContextMenuStrip = ContextMenuStrip1
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem0_Click(sender As Object, e As EventArgs) Handles ImportTemplate.Click
        ImportForm.ShowDialog(Me)

        Dim year As Integer = ImportForm.DateTimePicker1.Value.Year
        Dim projectName As String = ImportForm.TextBox1.Text.Trim

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

            Dim n As TreeNode
            With TreeView1
                n = .Nodes(0).Nodes.Add(year & YEAR_TEXT)
                n = n.Nodes.Add(projectName)
            End With

            For Each sht As Excel.Worksheet In wb.Worksheets
                n.Nodes.Add(sht.Name)
            Next

            TreeView1.ExpandAll()

            wb.Close(False)
            excelApp.Quit()

            wb = Nothing
            excelApp = Nothing

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

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        With FolderBrowserDialog1
            Dim result As DialogResult = .ShowDialog()
            If result = Windows.Forms.DialogResult.OK Then
                Dim folderName As String = .SelectedPath

                My.Computer.FileSystem.CopyFile(REPORT_FILE, folderName + "\" + My.Computer.FileSystem.GetName(REPORT_FILE), True)
                My.Computer.FileSystem.CopyFile(REPORT_FILE_BAK, folderName + "\" + My.Computer.FileSystem.GetName(REPORT_FILE_BAK), True)

                Dim excelApp As New Excel.Application
                excelApp.Visible = True
                excelApp.Workbooks.Open(folderName + "\" + My.Computer.FileSystem.GetName(REPORT_FILE))
            End If
        End With
    End Sub
End Class
