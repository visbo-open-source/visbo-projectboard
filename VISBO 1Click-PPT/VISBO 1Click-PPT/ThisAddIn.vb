Imports System
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports Microsoft.VisualBasic
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions


Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ' no actions

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        'Call MsgBox("byebye BHTCAddIn")
        Try
            If Not fehlerBeimLoad Then

                If Not IsNothing(appInstance.Workbooks(myCustomizationFile)) Then
                    ' hier wird die Datei Projekt Tafel Customizations als aktives workbook wieder geschlossen ....

                    If awinSettings.visboDebug Then
                        Call MsgBox("Anzahl Missing-Milestones: " & missingMilestoneDefinitions.Count & vbLf &
                               "Anzahl Missing-Phasen: " & missingPhaseDefinitions.Count)
                    End If

                    appInstance.Workbooks(myCustomizationFile).Close(SaveChanges:=False)    ' CustomizationFile wird ohne Abspeichern von Änderungen geschlossen
                End If

                If Not IsNothing(appInstance.Workbooks(myLogfile)) Then
                    ' Schließen des LogFiles
                    Call logfileSchliessen()
                End If

                My.Settings.rememberUserPWD = awinSettings.rememberUserPwd
                If awinSettings.rememberUserPwd Then
                    My.Settings.userNamePWD = awinSettings.userNamePWD
                Else
                    My.Settings.userNamePWD = ""
                End If
                My.Settings.Save()

                'appInstance.ScreenUpdating = True
                'Application.Quit()

            End If

            Try
                appInstance.Quit()
            Catch ex As Exception

            End Try

            Try
                ' die Excel Instanz zumachen
                pseudoappInstance.DisplayAlerts = False
                pseudoappInstance.Quit()

            Catch ex As Exception

            End Try

        Catch ex As Exception
            If awinSettings.englishLanguage Then
                Throw New ArgumentException("Error closing the Customization-Files")
            Else
                Throw New ArgumentException("Fehler beim Schließen des Customization-Files")
            End If

        End Try
    End Sub

    Private Sub ThisAddIn_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        Try
            If Not fehlerBeimLoad Then

                If Not IsNothing(appInstance.Workbooks(myCustomizationFile)) Then
                    ' hier wird die Datei Projekt Tafel Customizations als aktives workbook wieder geschlossen ....

                    If awinSettings.visboDebug Then
                        Call MsgBox("Anzahl Missing-Milestones: " & missingMilestoneDefinitions.Count & vbLf &
                               "Anzahl Missing-Phasen: " & missingPhaseDefinitions.Count)
                    End If

                    appInstance.Workbooks(myCustomizationFile).Close(SaveChanges:=False)    ' CustomizationFile wird ohne Abspeichern von Änderungen geschlossen
                End If

                If Not IsNothing(appInstance.Workbooks(myLogfile)) Then
                    ' Schließen des LogFiles
                    Call logfileSchliessen()
                End If


                'appInstance.ScreenUpdating = True
                'Application.Quit()

            End If
        Catch ex As Exception
            If awinSettings.englishLanguage Then
                Throw New ArgumentException("Error closing the Customization-Files")
            Else
                Throw New ArgumentException("Fehler beim Schließen des Customization-Files")
            End If
        End Try
    End Sub
End Class
