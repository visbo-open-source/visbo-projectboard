Imports System.Windows.Forms

Public Class frmProvideDate
    Private Sub clsFrmProvideDate_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call getFrmPosition(PTfrm.other, Top, Left)

        If awinSettings.englishLanguage Then
            Me.Text = "Provide Date"
        Else
            Me.Text = "Datum angeben"
        End If

        newDateValue.MaxDate = Date.Now.Date
        newDateValue.MinDate = StartofCalendar.Date
        If StartofCalendar < awinSettings.meDateForLastPlan Then
            newDateValue.Value = awinSettings.meDateForLastPlan
        Else
            newDateValue.Value = StartofCalendar.AddMonths(1)
        End If


    End Sub

    Private Sub ok_btn_Click(sender As Object, e As EventArgs) Handles ok_btn.Click
        awinSettings.meDateForLastPlan = newDateValue.Value
        MyBase.Close()
    End Sub

    Private Sub frmProvideDate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try
    End Sub
End Class