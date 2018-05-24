Imports DBAccLayer
Imports ProjectBoardDefinitions

Public Class frmRenameProject

    Private Sub Ok_Button_Click(sender As Object, e As EventArgs) Handles Ok_Button.Click

        ' es ist wichtig, dass keine führenden Blanks zugelassen sind ... 
        ' das ist in lostFocus_newName geregelt 
        Dim msgtxt As String = ""

        If IsNothing(newName.Text) Then
            If awinSettings.englishLanguage Then
                msgtxt = "please input non-empty name"
            Else
                msgtxt = "bitte Namen eingeben"
            End If
            Call MsgBox(msgtxt)
        Else
            If newName.Text = oldName.Text Then
                If awinSettings.englishLanguage Then
                    msgtxt = "no differences in name ..."
                Else
                    msgtxt = "keine Unterschiede ..."
                End If
                Call MsgBox(msgtxt)
            ElseIf newName.Text.Trim.Length < 2 Then
                If awinSettings.englishLanguage Then
                    msgtxt = "Name has to be at least 2 characters ..."
                Else
                    msgtxt = "Name muss mindestens 2 Zeichen lang sein ..."
                End If
                Call MsgBox(msgtxt)

            ElseIf Not isValidProjectName(newName.Text.Trim) Then
                If awinSettings.englishLanguage Then
                    msgtxt = "Name must not contain any #, (, ) characters ..."
                Else
                    msgtxt = "Name darf keine #, (, ) Zeichen enthalten  ..."
                End If
                Call MsgBox(msgtxt)

            Else
                Try
                    'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                    Dim projExist As Boolean = CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(newName.Text, "", Date.Now)

                    ' muss gemacht werden, weil es auch Projekte geben kann, die nur als Varianten existieren ...
                    Dim listOfVariants As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveVariantNamesFromDB(newName.Text)

                    If projExist Or listOfVariants.Count > 0 Then
                        ' es existiert bereits .. 
                        If awinSettings.englishLanguage Then
                            msgtxt = "Name does already exist in database ..."
                        Else
                            msgtxt = "Name existiert bereits in der Datenbank"
                        End If
                        Call MsgBox(msgtxt)
                    Else
                        DialogResult = System.Windows.Forms.DialogResult.OK
                    End If
                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        msgtxt = "Error when renaming: " & ex.Message
                    Else
                        msgtxt = "Fehler bei Rename: " & ex.Message
                    End If
                    Call MsgBox(msgtxt)
                    DialogResult = System.Windows.Forms.DialogResult.Cancel
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