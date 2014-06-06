
Imports System.Windows.Forms
Imports ProjectBoardDefinitions

Public Class frmdemoModusDate
    Public oldHistoryDate As Date

    Private Sub frmdemoModusDate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

    End Sub
    Private Sub frmdemoModusDate_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

        If DateDiff(DateInterval.Second, oldHistoryDate, DateTimeHistory.Value) >= 0 Then
            historicDate = DateTimeHistory.Value
        Else
            Call MsgBox("eingegebenes Datum muss später als '" & historicDate & " sein")
        End If


    End Sub

End Class