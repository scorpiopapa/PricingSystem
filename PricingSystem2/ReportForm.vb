Public Class ReportForm

    Private Shared ReadOnly ADD_TEXT = "添加"
    Private Shared ReadOnly UPDATE_TEXT = "修改"
    Private Shared ReadOnly BIND_COLUMN_NAME = "report_name"

    Private dbService As New AccessService

    Private Sub ReportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Button2.Text = ADD_TEXT

        Dim table As DataTable = dbService.FindAllReports
        FormUtils.InitComboBox(Me.ComboBox1, table, BIND_COLUMN_NAME)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class