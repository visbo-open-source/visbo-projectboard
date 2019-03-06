Imports ProjectBoardDefinitions
Imports DBAccLayer
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

        Dim err As New clsErrorCodeMsg

        Dim key As String
        Dim ok As Boolean = False

        If isValidVariantName(Me.newVariant.Text) Then
            key = calcProjektKey(Me.projektName.Text, Me.newVariant.Text)


            If Not noDB Then
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

            If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                    If Not _
                       (CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(projectname:=Me.projektName.Text,
                                                                                        variantname:=Me.newVariant.Text,
                                                                                        storedAtorBefore:=Date.Now, err:=err) _
                       Or AlleProjekte.Containskey(key)) Then

                        ' Projekt-Variante existiert noch nicht in der DB, kann also eingetragen werden
                        ok = True
                    Else
                        Dim msgTxt As String
                        If awinSettings.englishLanguage Then
                            msgTxt = "Projekt (Variante) " & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " &
                                "existiert bereits in DB!"
                        Else
                            msgTxt = "Project (Variant) " & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " &
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
                        msgTxt = " Projekt (Variante) '" & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " &
                            "existiert bereits in der Session!"
                    Else
                        msgTxt = "Project (Variant) " & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " &
                            "does already exist in session!"
                    End If
                    Call MsgBox(msgTxt)

                End If
            End If
        Else
            Dim msgTxt As String
            If awinSettings.englishLanguage Then
                msgTxt = " Varianten-Name ist ein System-Name und damit nicht zugelassen - bitte wählen Sie einen anderen Namen!"
            Else
                msgTxt = "Variant-Name is system-name and not therefore not allowed - please use another name!"
            End If
            Call MsgBox(msgTxt)
        End If

        If ok Then
            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()
        End If

    End Sub

    ''' <summary>
    ''' bestimmt, ob es sich um einen gültigen Varianten-Namen handelt; der Varianten-Name darf nicht einer der konstanten und festgelegten System-Varianten eines Projektes sein .. 
    ''' </summary>
    ''' <param name="vName"></param>
    ''' <returns></returns>
    Private Function isValidVariantName(ByVal vName As String) As Boolean
        Dim tmpResult As Boolean = False

        If Not IsNothing(vName) Then
            If vName.Trim.Length > 0 Then
                ' tk 2.3.19 Die Überprüfung muss auf ptVariantFixNames gehen 
                tmpResult = Not [Enum].GetNames(GetType(ptVariantFixNames)).Contains(vName.Trim)
            End If
        End If

        isValidVariantName = tmpResult

    End Function

End Class