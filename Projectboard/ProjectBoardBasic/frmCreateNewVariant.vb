Imports ProjectBoardDefinitions
Imports MongoDbAccess
Public Class frmCreateNewVariant

    Public multiSelect As Boolean = False

    Private Sub frmCreateNewVariant_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.createVariant, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.createVariant, PTpinfo.left) = Me.Left

    End Sub

    Private Sub frmCreateNewVariant_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Top = frmCoord(PTfrm.createVariant, PTpinfo.top)
        Me.Left = frmCoord(PTfrm.createVariant, PTpinfo.left)

        txtDescription.Text = ""

        If multiSelect Then
            infoText.Text = "den oben angegebenen Namen für alle selektierten Projekte verwenden"
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
                    Call MsgBox(" Projekt (Variante) '" & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " & _
                                "existiert bereits in DB!")
                End If

            Else
                Call MsgBox("Datenbank- Verbindung ist unterbrochen !")

            End If
        Else
            ' es wird ohne Datenbank gearbeitet
            If Not AlleProjekte.Containskey(key) Then
                ' Projekt-Variante existiert noch nicht in der Session, kann also eingetragen werden
                ok = True
            Else
                Call MsgBox(" Projekt (Variante) '" & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " & _
                            "existiert bereits in der Session!")
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