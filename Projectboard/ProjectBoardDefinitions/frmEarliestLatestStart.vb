Imports System.Windows.Forms

Public Class frmEarliestLatestStart

    Public setStartRange As frmEarliestLatestStart

    Private Sub frmEarliestLatestStart_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        'frmCoord(PTfrm.eingabeProj, PTpinfo.top) = Me.Top
        'frmCoord(PTfrm.eingabeProj, PTpinfo.left) = Me.Left

    End Sub
    Private Sub frmEarliestLatestStart_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        With Me

            ' Die Dauer des Projekts soll gleich der Dauer der Vorlage sein.

            '.EarliestStart.Value = Date.Now.AddMonths(1)
            '.LatestStart.Value = Date.Now.AddDays(vorlagenDauer - 1).AddMonths(1)

            '.selectedMonth.Value = DateDiff(DateInterval.Month, StartofCalendar, Date.Now) + 2

        End With
    End Sub
    Private Sub AbbruchButton_Click(sender As Object, e As EventArgs) Handles AbbruchButton.Click

        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub
    Private Sub EarliestStart_ValueChanged(sender As Object, e As EventArgs) Handles EarliestStart.ValueChanged

        'Call MsgBox("Earliest Start Value Changed")

        aktearliestStart.Text = CType(EarliestStart.Value, String)

    End Sub
    Private Sub LatestStart_ValueChanged(sender As Object, e As EventArgs) Handles LatestStart.ValueChanged

        'Call MsgBox("Latest Start Value Changed")

        aktlatestStart.Text = CType(LatestStart.Value, String)
    End Sub
End Class