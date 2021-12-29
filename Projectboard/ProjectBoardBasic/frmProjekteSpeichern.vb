Imports ProjectboardReports
Imports ProjectBoardDefinitions
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic
Imports System.Windows.Forms

Public Class frmProjekteSpeichern

    Private Sub JAButton_Click(sender As Object, e As EventArgs) Handles JAButton.Click

    End Sub
    Private Sub NEINButton_Click(sender As Object, e As EventArgs) Handles NEINButton.Click

    End Sub

    Private Sub frmProjekteSpeichern_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call getFrmPosition(PTfrm.other, Top, Left)

        Call languageSettings()

    End Sub

    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            Label1.Text = "Would you like to store your changes?"
            JAButton.Text = "Yes"
            NEINButton.Text = "No"
        End If
    End Sub

    Private Sub frmProjekteSpeichern_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try

    End Sub
End Class