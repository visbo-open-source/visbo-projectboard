Public Class frmPassword

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        DialogResult = Windows.Forms.DialogResult.OK
        MyBase.Close()
    End Sub

    Private Sub languageSettings()
        If englishLanguage Then
            With Me
                .Text = "provide Password"
                .Label1.Text = "Password:"
            End With
        End If
    End Sub


    Private Sub frmPassword_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call languageSettings()
    End Sub
End Class