Public Class frmProvideActualDataMonth

    Private Sub valueMonth_TextChanged(sender As Object, e As EventArgs) Handles valueMonth.TextChanged

    End Sub

    Private Sub frmProvideActualDataMonth_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim tmpMonth As Integer = Date.Now.Month - 1

        If tmpMonth < 1 Then
            tmpMonth = 12
        End If

        valueMonth.Text = tmpMonth.ToString("#0")


    End Sub
End Class