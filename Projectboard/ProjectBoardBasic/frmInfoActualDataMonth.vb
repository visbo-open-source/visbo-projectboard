Imports ProjectBoardDefinitions
Public Class frmInfoActualDataMonth
    Public Sub MonatJahr_ValueChanged(sender As Object, e As EventArgs) Handles MonatJahr.ValueChanged

    End Sub

    Private Sub okBtn_Click(sender As Object, e As EventArgs) Handles okBtn.Click

    End Sub

    Private Sub cancelBtn_Click(sender As Object, e As EventArgs) Handles cancelBtn.Click

    End Sub

    Private Sub frmInfoActualDataMonth_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MonatJahr.MinDate = StartofCalendar
        MonatJahr.MaxDate = Date.Now
    End Sub
End Class