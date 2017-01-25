Imports ProjectBoardDefinitions
Imports MongoDbAccess
Public Class frmCreateNewVariant

    Public multiSelect As Boolean = False

    Private Sub frmCreateNewVariant_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.createVariant, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.createVariant, PTpinfo.left) = Me.Left

    End Sub

    ''' <summary>
    ''' setzt in Abhängigkeit von menuCult die Texte der Formular-Felder 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub languageSettings()

        'If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
        If awinSettings.englishLanguage Then
            ' auf Englisch darstellen 
            Me.Text = "Create new Variant"
            lblNeueVariante.Text = "New Variant"
            lblDescription.Text = "Short description"
            infoText.Text = "The new variant will be created on base of this project-variant:"
            Label3.Text = "Project:"
            Label4.Text = "Variant:"
        End If

    End Sub
    Private Sub frmCreateNewVariant_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Top = frmCoord(PTfrm.createVariant, PTpinfo.top)
        Me.Left = frmCoord(PTfrm.createVariant, PTpinfo.left)

        txtDescription.Text = ""

        Call languageSettings()

        If multiSelect Then

            If awinSettings.englishLanguage Then
                infoText.Text = "den oben angegebenen Namen für alle selektierten Projekte verwenden"
            Else
                infoText.Text = "use the above given variantname for all selected projects"
            End If
            Label3.Visible = False
            Label4.Visible = False
            projektName.Visible = False
            variantenName.Visible = False
        End If

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim key As String
        Dim ok As Boolean = False

        key = calcProjektKey(Me.projektName.Text, Me.newVariant.Text)

        If Not noDB Then
            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

            If request.pingMongoDb() Then

                If Not _
                    (request.projectNameAlreadyExists(projectname:=Me.projektName.Text, variantname:=Me.newVariant.Text, storedAtorBefore:=Date.Now) Or _
                     AlleProjekte.Containskey(key)) Then

                    ' Projekt-Variante existiert noch nicht in der DB, kann also eingetragen werden
                    ok = True
                Else
                    Dim msgTxt As String
                    If awinSettings.englishLanguage Then
                        msgTxt = "Projekt (Variante) " & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " & _
                                "existiert bereits in DB!"
                    Else
                        msgTxt = "Project (Variant) " & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " & _
                                "does already exist in DB!"
                    End If
                    Call MsgBox(msgTxt)
                End If

            Else
                Dim msgTxt As String
                If awinSettings.englishLanguage Then
                    msgTxt = "Datenbank- Verbindung ist unterbrochen!"
                Else
                    msgTxt = "no database connection!"
                End If
                Call MsgBox(msgTxt)

            End If
        Else
            ' es wird ohne Datenbank gearbeitet
            If Not AlleProjekte.Containskey(key) Then
                ' Projekt-Variante existiert noch nicht in der Session, kann also eingetragen werden
                ok = True
            Else
                Dim msgTxt As String
                If awinSettings.englishLanguage Then
                    msgTxt = " Projekt (Variante) '" & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " & _
                            "existiert bereits in der Session!"
                Else
                    msgTxt = "Project (Variant) " & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " & _
                            "does already exist in session!"
                End If
                Call MsgBox(msgTxt)

            End If
        End If
      
        If ok Then
            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()
        End If

    End Sub

    Private Sub newVariant_TextChanged(sender As Object, e As EventArgs) Handles newVariant.TextChanged

    End Sub

    Private Sub infoText_Click(sender As Object, e As EventArgs) Handles infoText.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub
End Class