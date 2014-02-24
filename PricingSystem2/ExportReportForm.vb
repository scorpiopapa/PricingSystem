Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Public Class ExportReportForm
#If DEBUG Then
    Private TEMPLATE_FILE As String = "..\..\..\报价模板.xls"
#Else
    Private TEMPLATE_FILE As String = Application.StartupPath + "\config\报价模板.xls"
#End If
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
        With SaveFileDialog1
            .Filter = String.Format("Excel 2003 (*{0})|*{0}", ".xls")

            Dim result As DialogResult = .ShowDialog()
            If result = Windows.Forms.DialogResult.OK Then
                Dim fileName As String = .FileName

                My.Computer.FileSystem.CopyFile(TEMPLATE_FILE, .FileName, True)
                My.Computer.FileSystem.CopyFile(DB_FILE, My.Computer.FileSystem.GetParentPath(.FileName) + "\" + My.Computer.FileSystem.GetName(DB_FILE), True)

                Dim excelApp As New Excel.Application
                excelApp.Visible = True
                excelApp.Workbooks.Open(.FileName)

            End If
        End With
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