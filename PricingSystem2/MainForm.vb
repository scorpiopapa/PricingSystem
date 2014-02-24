Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Public Class MainForm

#If DEBUG Then
    Private REPORT_FILE As String = "..\..\..\报价单位A 项目1 2014年报价.xls"
    Private REPORT_FILE_BAK As String = "..\..\..\报价单位A 项目1 2014年报价"
#Else
    Private REPORT_FILE As String = Application.StartupPath + "\报价单位A 项目1 2014年报价.xls"
    Private REPORT_FILE_BAK As String = Application.StartupPath + "\报价单位A 项目1 2014年报价"
    Private TEMPLATE_FILE As String = Application.StartupPath + "\temp.xls"
#End If

    Public LoginUserName As String
    Public LoginForm As Form

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

        If TypeOf (LoginForm) Is SellerLogin Then
            ToolStripMenuItem8.Visible = False
        End If

        With TreeView1
            .HideSelection = False
            .ExpandAll()
            .SelectedNode = .Nodes(0)

            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(1).ContextMenuStrip = ContextMenuStrip1

            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(2).ContextMenuStrip = ContextMenuStrip1
            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(2).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(2).Nodes(1).ContextMenuStrip = ContextMenuStrip1

            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(3).ContextMenuStrip = ContextMenuStrip1
            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(3).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(3).Nodes(0).Nodes(0).ContextMenuStrip = ContextMenuStrip1
            .Nodes(0).Nodes(0).Nodes(0).Nodes(0).Nodes(3).Nodes(0).Nodes(1).ContextMenuStrip = ContextMenuStrip1
        End With


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
        If TreeView1.SelectedNode.Text.EndsWith("表") Then
            With TabControl1
                Dim tp As TabPage = CreateTabPage(TreeView1.SelectedNode.Text.Trim)
                .TabPages.Add(tp)
                .SelectedTab = tp
            End With
        ElseIf TreeView1.SelectedNode.Text.Trim = "报表总览" Then
            TabControl1.SelectedIndex = 0
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

        Using conn As New OleDbConnection(CONNECTION_STRING)
            conn.Open()

            Dim sql As String = "select * from test_table"
            Dim cmd As New OleDbCommand(sql, conn)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim table As New DataTable

            adapter.Fill(table)
            view.DataSource = table
        End Using

        Return tp
    End Function

    Private Sub TreeView1_MouseUp(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseUp
        'If e.Button = Windows.Forms.MouseButtons.Right Then
        '    If TreeView1.SelectedNode.Text.Contains("项目") Then
        '        TreeView1.SelectedNode.ContextMenuStrip = ContextMenuStrip1
        '    End If
        'End If
    End Sub

    Private Sub ToolStripMenuItem0_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem0.Click
        OpenFileDialog1.ShowDialog()
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
