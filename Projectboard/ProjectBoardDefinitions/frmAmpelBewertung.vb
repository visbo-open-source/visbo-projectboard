Public Class frmAmpelBewertung

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        DialogResult = System.Windows.Forms.DialogResult.OK
        MyBase.Close()
    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()
    End Sub

    Private Sub erlaeuterung_TextChanged(sender As Object, e As EventArgs) Handles erlaeuterung.TextChanged

    End Sub

    Private Sub ampelGruen_CheckedChanged(sender As Object, e As EventArgs) Handles ampelGruen.CheckedChanged

    End Sub

    Private Sub frmAmpelBewertung_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub
End Class