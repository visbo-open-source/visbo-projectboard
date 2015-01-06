Imports ProjectBoardDefinitions
Imports MongoDbAccess
Public Class frmCreateNewVariant


    Private Sub frmCreateNewVariant_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Dim request As New Request(awinSettings.databaseName)
        Dim key As String
        Dim ok As Boolean = False

        key = calcProjektKey(Me.projektName.Text, Me.newVariant.Text)

        If request.pingMongoDb() Then

            If Not _
                (request.projectNameAlreadyExists(projectname:=Me.projektName.Text, variantname:=Me.newVariant.Text) Or _
                 AlleProjekte.ContainsKey(key)) Then

                ' Projekt-Variante existiert noch nicht in der DB, kann also eingetragen werden
                ok = True
            Else
                Call MsgBox(" Projekt (Variante) '" & Me.projektName.Text & "( " & Me.newVariant.Text & " ) " & _
                            "existiert bereits !")
            End If

        Else

            Call MsgBox("Datenbank- Verbindung ist unterbrochen !")


        End If
        If ok Then
            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()
        End If

    End Sub
End Class