Public Class ImportForm


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text.Trim = "" Then
            FormUtils.ShowErrorMessage(EMPTY_PROJECT_NAME)
        Else
            Me.Hide()
        End If
    End Sub
End Class