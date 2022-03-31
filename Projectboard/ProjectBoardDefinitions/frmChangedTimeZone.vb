Imports System.ComponentModel

Public Class frmChangedTimeZone

    Public doNotShowAgain As Boolean = False

    Private Sub frmChangedTimeZone_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If awinSettings.englishLanguage Then
            messageBox.Text = "The following charts show only capacity and planned cost values calculated with the current Organisation   " & vbCrLf & tsOfcurrentOrga.ToString("MMMM dd, yyyy") & " - " & tsOfnextOrga.ToString("MMMM dd, yyyy")
        Else
            messageBox.Text = "In den folgenden Charts werden nur Werte während der Gültigkeit der aktuellen Organisation angezeigt !   " & vbCrLf & tsOfcurrentOrga.ToString("dd.MM.yyyy") & " - " & tsOfnextOrga.ToString("dd.MM.yyyy")
        End If
        doNotShowAgain = notAgain
    End Sub

    Private Sub showNoMore_CheckedChanged(sender As Object, e As EventArgs) Handles showNoMore.CheckedChanged
        If showNoMore.Checked Then
            doNotShowAgain = True
        Else
            doNotShowAgain = False
        End If
    End Sub

    Private Sub frmChangedTimeZone_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        notAgain = doNotShowAgain
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        MyBase.Close()
    End Sub

End Class