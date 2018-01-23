Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports MongoDbAccess
Imports System.Windows.Forms



Public Class frmAuthentication

    ' öffentliche Variable, ob userNamePWD gemerkt werden soll

    'Public loginResult As Integer = 0

    Private Sub benutzer_KeyDown(sender As Object, e As KeyEventArgs) Handles benutzer.KeyDown
        'If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
        '    maskedPwd.Focus()
        'End If

    End Sub

    Private Sub maskedPwd_ContextMenuStripChanged(sender As Object, e As EventArgs) Handles maskedPwd.ContextMenuStripChanged

    End Sub

    Private Sub maskedPwd_LostFocus(sender As Object, e As EventArgs) Handles maskedPwd.LostFocus

        'Dim pwd As String
        'Dim user As String
        'Dim projexist As Boolean

        'user = benutzer.Text
        'pwd = maskedPwd.Text


        'Try
        '    Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, user, pwd)
        '    projexist = request.projectNameAlreadyExists("TestProjekt", "v1", Date.Now)
        '    dbUsername = benutzer.Text
        '    dbPasswort = maskedPwd.Text
        '    messageBox.Text = ""
        '    DialogResult = System.Windows.Forms.DialogResult.OK
        'Catch ex As Exception
        '    messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
        '    benutzer.Text = ""
        '    maskedPwd.Text = ""
        '    dbUsername = benutzer.Text
        '    dbPasswort = maskedPwd.Text
        '    benutzer.Focus()
        '    DialogResult = System.Windows.Forms.DialogResult.Retry
        'End Try

    End Sub
    Private Sub maskedPwd_KeyDown(sender As Object, e As KeyEventArgs) Handles maskedPwd.KeyDown

        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then

            Dim pwd As String
            Dim user As String

            user = benutzer.Text
            pwd = maskedPwd.Text


            Try
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, user, pwd)
                Dim ok As Boolean = request.createIndicesOnce()
                If Not ok Then
                    messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
                    benutzer.Text = ""
                    maskedPwd.Text = ""
                    dbUsername = benutzer.Text
                    dbPasswort = maskedPwd.Text
                    benutzer.Focus()
                    DialogResult = System.Windows.Forms.DialogResult.Retry
                Else
                    '' ''projexist = request.projectNameAlreadyExists("TestProjekt", "v1", Date.Now)

                    dbUsername = benutzer.Text
                    dbPasswort = maskedPwd.Text
                    messageBox.Text = ""
                    DialogResult = System.Windows.Forms.DialogResult.OK
                End If

            Catch ex As Exception

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

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim pwd As String
        Dim user As String

        user = benutzer.Text
        pwd = maskedPwd.Text
        messageBox.Text = ""

        Try         ' dieser Try Catch dauert so lange, da beim Request ein TimeOut von 30000ms eingestellt ist
            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, user, pwd)
            Dim ok As Boolean = request.createIndicesOnce()
            If Not ok Then
                messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
                benutzer.Text = ""
                maskedPwd.Text = ""
                dbUsername = benutzer.Text
                dbPasswort = maskedPwd.Text
                benutzer.Focus()
                DialogResult = System.Windows.Forms.DialogResult.Retry
            Else
                dbUsername = benutzer.Text
                dbPasswort = maskedPwd.Text
                messageBox.Text = ""
                DialogResult = System.Windows.Forms.DialogResult.OK
                ' hier werden einmalig alle Projekte in die WriteProtections Collection eingetragen
                Dim initOK As Integer = request.initWriteProtectionsOnce(dbUsername)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmAuthentication_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        chbx_remember.Checked = awinSettings.rememberUserPwd
    End Sub

    Private Sub chbx_remember_CheckedChanged(sender As Object, e As EventArgs) Handles chbx_remember.CheckedChanged

        If chbx_remember.Checked Then
            awinSettings.rememberUserPwd = True
        Else
            awinSettings.rememberUserPwd = False
        End If
    End Sub
End Class