Public Class frmEditDates
    Public IsMilestone As Boolean = False

    Private Sub frmEditDates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call languageSettings()

        ' tk 29.11.23 
        chkbx_adjustChilds.Checked = awinSettings.autoAjustChilds
        chkbxAutoDistr.Checked = Not awinSettings.noNewCalculation

        If IsMilestone Then
            chkbx_adjustChilds.Visible = False
            chkbxAutoDistr.Visible = False
        End If

    End Sub

    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            btn_Cancel.Text = "Cancel"
            btn_OK.Text = "OK"
            chkbx_adjustChilds.Text = "auto adjust children"
            chkbxAutoDistr.Text = "auto adjust resource costs"

            If IsMilestone Then
                Text = "Edit Milestone Date"
            Else
                Text = "Edit Phase Dates"
            End If

        Else
            btn_Cancel.Text = "Cancel"
            btn_OK.Text = "OK"
            If IsMilestone Then
                Text = "Edit Milestone Date"
            Else
                Text = "Edit Phase Dates"
            End If
        End If

    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click

        DialogResult = Windows.Forms.DialogResult.OK
        MyBase.Close()

    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click

        DialogResult = Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

    Private Sub enddatePicker_ValueChanged(sender As Object, e As EventArgs) Handles enddatePicker.ValueChanged
        startdatePicker.MaxDate = enddatePicker.Value
    End Sub

    Private Sub chkbxAutoDistr_CheckedChanged(sender As Object, e As EventArgs) Handles chkbxAutoDistr.CheckedChanged
        ' tk 29.11.23
        awinSettings.noNewCalculation = Not chkbxAutoDistr.Checked
    End Sub

    Private Sub chkbx_adjustChilds_CheckedChanged(sender As Object, e As EventArgs) Handles chkbx_adjustChilds.CheckedChanged
        awinSettings.autoAjustChilds = chkbx_adjustChilds.Checked
    End Sub

    Private Sub startdatePicker_ValueChanged(sender As Object, e As EventArgs) Handles startdatePicker.ValueChanged
        enddatePicker.MinDate = startdatePicker.Value
    End Sub
End Class