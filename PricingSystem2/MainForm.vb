Public Class MainForm

    Private Sub 明细表管理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 明细表管理ToolStripMenuItem.Click
        ShowForm(AddReportForm)

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        UpdateReportForm.ShowDialog(Me)

    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With TreeView1
            .HideSelection = False
            .ExpandAll()
            .SelectedNode = .Nodes(0)
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
                .TabPages.Add(CreateTabPage(TreeView1.SelectedNode.Text.Trim))
            End With
        End If
    End Sub

    Private Function CreateTabPage(title As String) As TabPage
        Dim tp As New TabPage(title)

        Dim view As New DataGridView
        Dim btn As New Button With {.Text = "保存"}

        With tp.Controls
            .Add(view)
            .Add(btn)
        End With

        With view
            .Left = 13
            .Top = 22
        End With

        With btn
            .Left = view.Left + 10
            .Top = view.Top
            .Width = 64
            .Height = 33

        End With
        'btn.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        view.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right


        Return tp
    End Function

End Class
