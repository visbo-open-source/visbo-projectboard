
Imports System.Windows.Forms
Imports ProjectBoardDefinitions

Public Class frmdemoModusDate
    Public oldHistoryDate As Date

    Private Sub frmdemoModusDate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try
    End Sub
    Private Sub frmdemoModusDate_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call getFrmPosition(PTfrm.other, Top, Left)

        oldHistoryDate = historicDate
        DateTimeHistory.Value = historicDate

    End Sub
    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        DialogResult = System.Windows.Forms.DialogResult.OK
        MyBase.Close()
    End Sub
    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

        historicDate = oldHistoryDate
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub
    Private Sub DateTimeHistory_ValueChanged(sender As Object, e As EventArgs) Handles DateTimeHistory.ValueChanged

        Dim hours As Integer, minutes As Integer
        hours = Date.Now.Hour
        minutes = Date.Now.Minute

        If DateDiff(DateInterval.Second, oldHistoryDate, DateTimeHistory.Value) >= 0 Then
            historicDate = DateTimeHistory.Value.Date.AddHours(hours).AddMinutes(minutes)

        Else
            Call MsgBox("eingegebenes Datum muss später als '" & historicDate & " sein")
        End If


    End Sub

End Class