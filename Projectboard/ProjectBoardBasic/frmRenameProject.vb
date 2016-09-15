Imports MongoDbAccess
Imports ProjectBoardDefinitions

Public Class frmRenameProject

    Private Sub Ok_Button_Click(sender As Object, e As EventArgs) Handles Ok_Button.Click

        ' es ist wichtig, dass keine führenden Blanks zugelassen sind ... 
        ' das ist in lostFocus_newName geregelt 

        If IsNothing(newName.Text) Then
            Call MsgBox("bitte Namen eingeben")
        Else
            If newName.Text = oldName.Text Then
                Call MsgBox("keine Unterschiede ...")
            ElseIf newName.Text.Trim.Length < 2 Then
                Call MsgBox("Name muss mindestens 2 Zeichen lang sein ...")
            Else
                Try
                    Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                    Dim projExist As Boolean = request.projectNameAlreadyExists(newName.Text, "", Date.Now)

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
        End If
        
    End Sub

    Private Sub frmRenameProject_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub newName_LostFocus(sender As Object, e As EventArgs) Handles newName.LostFocus
        If Not IsNothing(newName) Then
            newName.Text = newName.Text.Trim
        End If
    End Sub

    
End Class