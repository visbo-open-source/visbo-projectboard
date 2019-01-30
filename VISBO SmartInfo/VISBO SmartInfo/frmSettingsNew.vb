Public Class frmSettingsNew


    Private Sub ValidateFontSize()

        If IsNumeric(fontSize.Text) Then
            If CDbl(fontSize.Text) > 3 Then
                ' alles ok 
            Else
                fontSize.Text = CStr(schriftGroesse)
            End If
        Else
            fontSize.Text = CStr(schriftGroesse)
        End If

    End Sub



    Private Sub fontSize_LostFocus(sender As Object, e As EventArgs) Handles fontSize.LostFocus
        Call ValidateFontSize()
    End Sub

    Private Sub frmSettingsNew_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        fontSize.Text = schriftGroesse.ToString
        chkbxEditable.Checked = smartChartsAreEditable
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        schriftGroesse = CDbl(fontSize.Text)
        smartChartsAreEditable = chkbxEditable.Checked
    End Sub

    Private Sub fontSize_Enter(sender As Object, e As EventArgs) Handles fontSize.Enter
        Call ValidateFontSize()
    End Sub
End Class