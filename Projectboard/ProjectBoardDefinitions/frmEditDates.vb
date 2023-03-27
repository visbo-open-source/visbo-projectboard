Public Class frmEditDates
    Public IsMilestone As Boolean = False
    Public allowedDateLeft As Date
    Public allowedDateRight As Date
    Public maxPossibleStartDate As Date

    Private Sub frmEditDates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call languageSettings()

        If visboClient = divClients(client.VisboSPE) Then
            chkbx_adjustChilds.Visible = False
            chkbxAutoDistr.Visible = False
        Else
            awinSettings.autoAjustChilds = chkbx_adjustChilds.Checked
        End If

        If allowedDateLeft > Date.MinValue Then
            If startdatePicker.Enabled Then
                startdatePicker.MinDate = allowedDateLeft
                startdatePicker.MaxDate = allowedDateRight
            End If
        End If

        If allowedDateRight > Date.MinValue Then
            enddatePicker.MinDate = allowedDateLeft
            If allowedDateRight > allowedDateLeft Then
                enddatePicker.MaxDate = allowedDateRight
            Else
                enddatePicker.MaxDate = allowedDateLeft
            End If

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
        Dim allIsOk As Boolean = False

        If IsMilestone Then
            If DateDiff(DateInterval.Day, allowedDateLeft, CDate(enddatePicker.Value)) >= 0 And DateDiff(DateInterval.Day, allowedDateRight, CDate(enddatePicker.Value)) <= 0 Then
                allIsOk = True
            End If
        Else
            ' es handelt sich um eine Phase

            'If visboClient = divClients(client.VisboSPE) Then
            If DateDiff(DateInterval.Day, CDate(startdatePicker.Value), CDate(enddatePicker.Value)) >= 0 Then
                If awinSettings.autoAjustChilds Then
                    allIsOk = True
                Else
                    Dim diffdays As Long = DateDiff(DateInterval.Day, CDate(maxPossibleStartDate), CDate(startdatePicker.Value))
                    If Not IsNothing(maxPossibleStartDate) And DateDiff(DateInterval.Day, CDate(maxPossibleStartDate), CDate(startdatePicker.Value)) <= 0 Then

                        allIsOk = True
                    Else
                        Dim errMsg As String = "Start-Datum darf nicht nach dem Start-Datum der untergeordneten Phase liegen ..."
                        If awinSettings.englishLanguage Then
                            errMsg = "start-date should be earlier or equal to start-date of the child ..."
                        End If
                        Call MsgBox(errMsg)
                    End If

                End If

            Else
                Dim errMsg As String = "Ende-Datum darf nicht vor dem Start-Datum liegen ..."
                If awinSettings.englishLanguage Then
                    errMsg = "end-date should be later or equal to start-date ..."
                End If
                Call MsgBox(errMsg)
            End If

            'Else
            '    ' visboClient other
            '    If DateDiff(DateInterval.Day, CDate(startdatePicker.Value), CDate(enddatePicker.Value)) >= 0 Then
            '        allIsOk = True
            '    Else
            '        Dim errMsg As String = "Ende-Datum darf nicht vor dem Start-Datum liegen ..."
            '        If awinSettings.englishLanguage Then
            '            errMsg = "end-date should be later or equal to start-date ..."
            '        End If
            '        Call MsgBox(errMsg)
            '    End If
            'End If


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

    Private Sub chkbxAutoDistr_CheckedChanged(sender As Object, e As EventArgs) Handles chkbxAutoDistr.CheckedChanged

    End Sub

    Private Sub chkbx_adjustChilds_CheckedChanged(sender As Object, e As EventArgs) Handles chkbx_adjustChilds.CheckedChanged
        awinSettings.autoAjustChilds = chkbx_adjustChilds.Checked
    End Sub
End Class