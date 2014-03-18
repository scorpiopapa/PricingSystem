Public Class SellerLogin

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox3.Text.Trim = "" Then
            FormUtils.ShowErrorMessage(EMPTY_COM_NAME)
        Else
            FormUtils.ShowMainForm(Me)
        End If
    End Sub
End Class