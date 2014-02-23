Imports Microsoft.Office.Interop
Public Class ExportReportForm
#If DEBUG Then
    Private TEMPLATE_FILE As String = "..\..\..\报价模板.xls"
#Else
    Private TEMPLATE_FILE As String = Application.StartupPath + "\报价模板.xls"
#End If
    Private Sub ExportReportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
End Class