Public Class frmEditDates
    Public IsMilestone As Boolean = False
    Public allowedDateLeft As Date
    Public allowedDateRight As Date

    Private Sub frmEditDates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call languageSettings()

        If allowedDateLeft > Date.MinValue Then
            startdatePicker.MinDate = allowedDateLeft
            enddatePicker.MinDate = allowedDateRight
        End If

        If allowedDateRight > Date.MinValue Then
            enddatePicker.MinDate = allowedDateLeft
            enddatePicker.MaxDate = allowedDateRight
        End If

        If IsMilestone Then
            startdatePicker.Enabled = False
        Else
            startdatePicker.Enabled = True
        End If
    End Sub

    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            btn_Cancel.Text = "Cancel"
            btn_OK.Text = "OK"
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
        Dim allIsOk As Boolean = False

        If IsMilestone Then
            If DateDiff(DateInterval.Day, allowedDateLeft, CDate(enddatePicker.Value)) >= 0 And DateDiff(DateInterval.Day, allowedDateRight, CDate(enddatePicker.Value)) <= 0 Then
                allIsOk = True
            End If
        Else
            ' es handelt sich um eine Phase
            If DateDiff(DateInterval.Day, allowedDateLeft, CDate(startdatePicker.Value)) >= 0 And DateDiff(DateInterval.Day, allowedDateRight, CDate(startdatePicker.Value)) <= 0 And
                DateDiff(DateInterval.Day, allowedDateLeft, CDate(enddatePicker.Value)) >= 0 And DateDiff(DateInterval.Day, allowedDateRight, CDate(enddatePicker.Value)) <= 0 And
                DateDiff(DateInterval.Day, CDate(startdatePicker.Value), CDate(enddatePicker.Value)) >= 0 Then
                allIsOk = True
            End If
        End If

        If allIsOk Then
            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()
        End If


    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click

        DialogResult = Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

    Private Sub enddatePicker_ValueChanged(sender As Object, e As EventArgs) Handles enddatePicker.ValueChanged
        If IsMilestone Then
            startdatePicker.Value = enddatePicker.Value
        End If
    End Sub
End Class