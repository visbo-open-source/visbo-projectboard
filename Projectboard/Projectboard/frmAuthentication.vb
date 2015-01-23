Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports MongoDbAccess




Public Class frmAuthentication


    Private Sub benutzer_KeyDown(sender As Object, e As KeyEventArgs) Handles benutzer.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
            maskedPwd.Focus()
        End If

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
            username = benutzer.Text
            password = maskedPwd.Text
            messageBox.Text = ""
            DialogResult = System.Windows.Forms.DialogResult.OK
        Catch ex As Exception
            messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
            benutzer.Text = ""
            maskedPwd.Text = ""
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
                username = benutzer.Text
                password = maskedPwd.Text
                messageBox.Text = ""
                DialogResult = System.Windows.Forms.DialogResult.OK
            Catch ex As Exception
                messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
                benutzer.Text = ""
                maskedPwd.Text = ""
                benutzer.Focus()
                DialogResult = System.Windows.Forms.DialogResult.Retry
            End Try

        End If


    End Sub



    Private Sub maskedPwd_TextChanged(sender As Object, e As EventArgs) Handles maskedPwd.TextChanged
     
    End Sub


 
End Class