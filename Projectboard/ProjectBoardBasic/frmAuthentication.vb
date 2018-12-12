Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
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


    End Sub
    Private Sub maskedPwd_KeyDown(sender As Object, e As KeyEventArgs) Handles maskedPwd.KeyDown

        Dim err As New clsErrorCodeMsg

        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then

            Dim pwd As String
            Dim user As String

            user = benutzer.Text
            pwd = maskedPwd.Text


            Try
                'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, user, pwd)
                'Dim ok As Boolean = Request.createIndicesOnce()

                If IsNothing(databaseAcc) Then
                    Dim hrequest As New DBAccLayer.Request
                    databaseAcc = hrequest
                End If

                Dim ok As Boolean = CType(databaseAcc, DBAccLayer.Request).login(awinSettings.databaseURL, awinSettings.databaseName, user, pwd, err)

                If Not ok Then
                    If awinSettings.englishLanguage Then
                        messageBox.Text = "Wrong username or password!"
                    Else
                        messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
                    End If

                    benutzer.Text = ""
                    maskedPwd.Text = ""
                    dbUsername = benutzer.Text
                    dbPasswort = maskedPwd.Text
                    benutzer.Focus()
                    DialogResult = System.Windows.Forms.DialogResult.Retry
                Else
                    '' ''projexist = CType(mongoDBAcc, Request).projectNameAlreadyExists("TestProjekt", "v1", Date.Now)

                    dbUsername = benutzer.Text
                    dbPasswort = maskedPwd.Text
                    messageBox.Text = ""
                    DialogResult = System.Windows.Forms.DialogResult.OK

                    ' jett wird request public gemacht ..
                    ' mongoDBAcc = Request

                    ' UR: 07.07.2018: sollte für WebServerAcc eigentlich nicht benötigt werden
                    ' mongoDBAcc = token
                End If

            Catch ex As Exception
                Throw New ArgumentException(ex.Message)
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

        Dim err As New clsErrorCodeMsg


        Dim pwd As String
        Dim user As String

        user = benutzer.Text
        pwd = maskedPwd.Text
        messageBox.Text = ""

        Try         ' dieser Try Catch dauert so lange, da beim Request ein TimeOut von 30000ms eingestellt ist
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, user, pwd)
            'Dim ok As Boolean = Request.createIndicesOnce()

            If IsNothing(databaseAcc) Then
                Dim hrequest As New DBAccLayer.Request
                databaseAcc = hrequest
            End If

            Dim ok As Boolean = CType(databaseAcc, DBAccLayer.Request).login(awinSettings.databaseURL, awinSettings.databaseName, user, pwd, err)

            If Not ok Then
                If awinSettings.englishLanguage Then
                    messageBox.Text = "Wrong username or password!"
                Else
                    messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
                End If
                benutzer.Text = ""
                maskedPwd.Text = ""
                dbUsername = benutzer.Text
                dbPasswort = maskedPwd.Text
                benutzer.Focus()
                DialogResult = System.Windows.Forms.DialogResult.Retry
            Else
                ' login am Rest-Server/mongoDB hat funktioniert

                ' UR: 07.07.2018: sollte für WebServerAcc eigentlich nicht benötigt werden
                dbUsername = benutzer.Text
                dbPasswort = maskedPwd.Text

                If awinSettings.rememberUserPwd Then

                    ' Username Passwort verschlüsselt merken
                    Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)
                    awinSettings.userNamePWD = visboCrypto.verschluessleUserPwd(dbUsername, dbPasswort)

                End If

                messageBox.Text = ""
                DialogResult = System.Windows.Forms.DialogResult.OK


                '' UR: 07.07.2018: sollte für WebServerAcc eigentlich nicht benötigt werden
                '' hier werden einmalig alle Projekte in die WriteProtections Collection eingetragen
                ' Dim initOK As Integer = CType(mongoDBAcc, MongoDbAccess.Request).initWriteProtectionsOnce(dbUsername)

            End If
        Catch ex As Exception
            If awinSettings.englishLanguage Then
                messageBox.Text = "Wrong username or password!"
            Else
                messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
            End If
            benutzer.Text = ""
            maskedPwd.Text = ""
            dbUsername = benutzer.Text
            dbPasswort = maskedPwd.Text
            benutzer.Focus()
            DialogResult = System.Windows.Forms.DialogResult.Retry
        End Try
    End Sub

    Private Sub frmAuthentication_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            If awinSettings.visboServer Then
                pwforgotten.Visible = True
                pwforgotten.Enabled = True
            Else
                pwforgotten.Visible = False
                pwforgotten.Enabled = False
            End If

            If awinSettings.englishLanguage Then
                Label2.Text = "Username"
                Label3.Text = "Password"
                chbx_remember.Text = "Remember Me"
                abbrButton.Text = "Cancel"
                pwforgotten.Text = "Password forgotten"
            Else
                Label2.Text = "Benutzername"
                Label3.Text = "Passwort"
                chbx_remember.Text = "Passwort speichern"
                abbrButton.Text = "Abbrechen"
                pwforgotten.Text = "Passwort vergessen"
            End If

            Dim cipherText As String = awinSettings.userNamePWD
            Dim pwd As String = ""
            Dim user As String = ""

            If awinSettings.rememberUserPwd Then

                Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)

                user = visboCrypto.getUserNameFromCipher(cipherText)
                pwd = visboCrypto.getPwdFromCipher(cipherText)

                chbx_remember.Checked = True
            Else
                chbx_remember.Checked = False
            End If

            benutzer.Text = user
            maskedPwd.Text = pwd


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

    End Sub

    Private Sub chbx_remember_CheckedChanged(sender As Object, e As EventArgs) Handles chbx_remember.CheckedChanged

        If chbx_remember.Checked Then
            awinSettings.rememberUserPwd = True
        Else
            awinSettings.rememberUserPwd = False
        End If
    End Sub

    Private Sub benutzer_TextChanged(sender As Object, e As EventArgs) Handles benutzer.TextChanged

    End Sub

    Private Sub pwforgotten_Click(sender As Object, e As EventArgs) Handles pwforgotten.Click
        Try

            Dim hrequest As New DBAccLayer.Request
            databaseAcc = hrequest
            Dim pwd As String
            Dim user As String

            user = benutzer.Text
            pwd = maskedPwd.Text

            If user = "" Then

                If awinSettings.englishLanguage Then
                    Call MsgBox("Please enter your username")
                Else
                    Call MsgBox("Bitte geben Sie den Benutzer ein")
                End If

            Else

                Dim ok As Boolean = CType(databaseAcc, DBAccLayer.Request).pwforgotten(awinSettings.databaseURL, awinSettings.databaseName, user)

                If ok Then
                    If awinSettings.englishLanguage Then
                        Call MsgBox("You'll receive an email with the link to reset your password!")
                    Else
                        Call MsgBox("Sie erhalten eine Email mit dem Link zum Reset Ihres Passwortes")
                    End If
                Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Error: Please contact your administrator")
                    Else
                        Call MsgBox("Fehler: Bitte kontaktieren Sie ihren Administrator")
                    End If
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub
End Class