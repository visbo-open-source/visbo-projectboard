Imports MongoDbAccess
Imports ProjectBoardDefinitions

Public Class frmRenameProject

    Private Sub Ok_Button_Click(sender As Object, e As EventArgs) Handles Ok_Button.Click

        If newName.Text = oldName.Text Then
            Call MsgBox("keine Unterschiede ...")
        ElseIf newName.Text.Length < 2 Then
            Call MsgBox("Name muss mindestens 2 Zeichen lang sein ...")
        Else
            Try
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                Dim projExist As Boolean = request.projectNameAlreadyExists(newName.Text, "")

                ' muss gemacht werden, weil es auch Projekte geben kann, die nur als Varianten existieren ...
                Dim listOfVariants As Collection = request.retrieveVariantNamesFromDB(newName.Text)

                If projExist Or listOfVariants.Count > 0 Then
                    ' es existiert bereits .. 
                    Call MsgBox("Name existiert bereits in der Datenbank")
                Else
                    DialogResult = System.Windows.Forms.DialogResult.OK
                End If
            Catch ex As Exception

            End Try

        End If
    End Sub

    Private Sub frmRenameProject_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
    End Sub
End Class