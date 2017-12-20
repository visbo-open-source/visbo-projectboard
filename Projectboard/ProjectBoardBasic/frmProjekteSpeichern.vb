Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic

Public Class frmProjekteSpeichern

    Private Sub JAButton_Click(sender As Object, e As EventArgs) Handles JAButton.Click

    End Sub
    Private Sub NEINButton_Click(sender As Object, e As EventArgs) Handles NEINButton.Click

    End Sub

    Private Sub frmProjekteSpeichern_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call LanguageSettings()

    End Sub

    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            Label1.Text = "Would you like to store your changes?"
            JAButton.Text = "Yes"
            NEINButton.Text = "No"
        End If
    End Sub
End Class