Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports MongoDbAccess



Public Class frmAuthentication

    'Public loginResult As Integer = 0

    Private Sub benutzer_KeyDown(sender As Object, e As KeyEventArgs) Handles benutzer.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
            maskedPwd.Focus()
        End If

    End Sub

    Private Sub maskedPwd_ContextMenuStripChanged(sender As Object, e As EventArgs) Handles maskedPwd.ContextMenuStripChanged

    End Sub

    Private Sub maskedPwd_LostFocus(sender As Object, e As EventArgs) Handles maskedPwd.LostFocus

        Dim pwd As String
        Dim user As String
        Dim projexist As Boolean

        user = benutzer.Text
        pwd = maskedPwd.Text


        Try
            Dim request As New Request(awinSettings.databaseName, user, pwd)
            projexist = request.projectNameAlreadyExists("TestProjekt", "v1")
            dbUsername = benutzer.Text
            dbPasswort = maskedPwd.Text
            messageBox.Text = ""
            DialogResult = System.Windows.Forms.DialogResult.OK
        Catch ex As Exception
            messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
            benutzer.Text = ""
            maskedPwd.Text = ""
            dbUsername = benutzer.Text
            dbPasswort = maskedPwd.Text
            benutzer.Focus()
            DialogResult = System.Windows.Forms.DialogResult.Retry
        End Try

    End Sub
    Private Sub maskedPwd_KeyDown(sender As Object, e As KeyEventArgs) Handles maskedPwd.KeyDown

        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then

            Dim pwd As String
            Dim user As String
            Dim projexist As Boolean

            user = benutzer.Text
            pwd = maskedPwd.Text


            Try
                Dim request As New Request(awinSettings.databaseName, user, pwd)
                projexist = request.projectNameAlreadyExists("TestProjekt", "v1")
                dbUsername = benutzer.Text
                dbPasswort = maskedPwd.Text
                messageBox.Text = ""
                DialogResult = System.Windows.Forms.DialogResult.OK
            Catch ex As Exception
                messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
                benutzer.Text = ""
                maskedPwd.Text = ""
                dbUsername = benutzer.Text
                dbPasswort = maskedPwd.Text
                benutzer.Focus()
                DialogResult = System.Windows.Forms.DialogResult.Retry
            End Try

        End If


    End Sub



    Private Sub maskedPwd_TextChanged(sender As Object, e As EventArgs) Handles maskedPwd.TextChanged

    End Sub



    Private Sub frmAuthentication_FormClosed(sender As Object, e As EventArgs) Handles MyBase.FormClosed
        Dim username_sav As String = dbUsername
        Dim dbPasswort_sav As String = dbPasswort

    End Sub


    Private Sub abbrButton_Click(sender As Object, e As EventArgs) Handles abbrButton.Click

    End Sub
End Class