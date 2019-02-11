Imports ProjectBoardDefinitions
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Public Class frmProxyAuth
    Public proxyURL As String
    Public domain As String
    Public user As String
    Public pwd As String
    Private Sub benutzer_TextChanged(sender As Object, e As EventArgs) Handles benutzer.TextChanged

    End Sub
    Private Sub maskedPwd_KeyDown(sender As Object, e As KeyEventArgs) Handles maskedPwd.KeyDown

        Dim err As New clsErrorCodeMsg

        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then

            domain = domainBox.Text
            user = benutzer.Text
            pwd = maskedPwd.Text


            Try
                'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, user, pwd)
                'Dim ok As Boolean = Request.createIndicesOnce()

                'If IsNothing(databaseAcc) Then
                '    Dim hrequest As New DBAccLayer.Request
                '    databaseAcc = hrequest
                'End If

                'Dim ok As Boolean = CType(databaseAcc, DBAccLayer.Request).login(awinSettings.databaseURL, awinSettings.databaseName, user, pwd, err)

                'If Not ok Then
                '    If awinSettings.englishLanguage Then
                '        messageBox.Text = "Wrong username or password!"
                '    Else
                '        messageBox.Text = "Benutzername oder Passwort fehlerhaft!"
                '    End If

                '    benutzer.Text = ""
                '    maskedPwd.Text = ""
                '    dbUsername = benutzer.Text
                '    dbPasswort = maskedPwd.Text
                '    benutzer.Focus()
                '    DialogResult = System.Windows.Forms.DialogResult.Retry
                'Else
                '    '' ''projexist = CType(mongoDBAcc, Request).projectNameAlreadyExists("TestProjekt", "v1", Date.Now)

                '    dbUsername = benutzer.Text
                '    dbPasswort = maskedPwd.Text
                '    messageBox.Text = ""
                '    DialogResult = System.Windows.Forms.DialogResult.OK

                '    ' jett wird request public gemacht ..
                '    ' mongoDBAcc = Request

                '    ' UR: 07.07.2018: sollte für WebServerAcc eigentlich nicht benötigt werden
                '    ' mongoDBAcc = token
                'End If

            Catch ex As Exception
                Throw New ArgumentException(ex.Message)
            End Try

        End If


    End Sub
    Private Sub maskedPwd_TextChanged(sender As Object, e As EventArgs) Handles maskedPwd.TextChanged

    End Sub

    Private Sub frmProxyAuth_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        proxyURLbox.Text = proxyURL
        domainBox.Text = domain
        benutzer.Text = user
        maskedPwd.Text = pwd
    End Sub

    Private Sub frmProxyAuth_FormClosed(sender As Object, e As EventArgs) Handles MyBase.FormClosed
        'domain = ""
        'user = ""
        'pwd = ""
    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click
        domain = ""
        user = ""
        pwd = ""
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Dim uriResult As Uri = Nothing
        Dim uriOK As Boolean = False
        DialogResult = DialogResult.OK

        If proxyURLbox.Text <> "" Then
            uriOK = Uri.TryCreate(proxyURLbox.Text, UriKind.Absolute, uriResult)
            If uriOK Then
                uriOK = uriOK And (uriResult.Scheme = Uri.UriSchemeHttp)
            End If
        End If



        If uriOK Then
            proxyURL = proxyURLbox.Text
        Else
            messageBox.Text = "no valid Proxy-URL"
            DialogResult = DialogResult.Retry
        End If
        domain = domainBox.Text
        user = benutzer.Text
        pwd = maskedPwd.Text
        If user = "" Or pwd = "" Then
            messageBox.Text = "Username/Passwort für Proxy eingeben!"
            DialogResult = DialogResult.Retry
        End If

    End Sub

    Private Sub messageBox_TextChanged(sender As Object, e As EventArgs) Handles messageBox.TextChanged

    End Sub

    Private Sub Domain_TextChanged(sender As Object, e As EventArgs) Handles domainBox.TextChanged

    End Sub
End Class