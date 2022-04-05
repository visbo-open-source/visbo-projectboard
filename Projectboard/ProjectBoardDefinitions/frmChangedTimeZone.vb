Imports System.ComponentModel

Public Class frmChangedTimeZone

    Public doNotShowAgain As Boolean = False

    Private Sub frmChangedTimeZone_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If awinSettings.englishLanguage Then
            'messageBox.Text = "The following charts only show values based on the latest organization !   " & vbCrLf & tsOfcurrentOrga.ToString("MMMM dd, yyyy") & " - " & tsOfnextOrga.ToString("MMMM dd, yyyy")
            messageBox.Text = "The following charts only show values based on the latest organization, " & vbCrLf & "valid from: " & tsOfcurrentOrga.ToString("MMMM dd, yyyy")
        Else
            messageBox.Text = "In den folgenden Charts werden nur Werte auf Basis der neuesten Organisation angezeigt, " & vbCrLf & "gültig ab: " & tsOfcurrentOrga.ToString("dd.MM.yyyy")
            doNotShowAgain = notAgain
        End If

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