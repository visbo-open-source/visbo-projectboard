Imports System
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports Microsoft.VisualBasic
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions


Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ' war nur zu Testzwecken

        ' ''Call MsgBox("XML write TEst anfang")
        ' ''Try
        ' ''    Call xmltestwrite2()
        ' ''    Call xmltestread2()
        ' ''Catch ex As Exception
        ' ''    Call MsgBox("XML write TEst Fehler")
        ' ''End Try


        ''Call MsgBox("Load VISBO Report Testversion")



        Try

            awinSettings.databaseURL = My.Settings.mongoDBURL
            awinSettings.databaseName = My.Settings.mongoDBname
            awinSettings.globalPath = My.Settings.globalPath
            awinSettings.awinPath = My.Settings.awinPath
            awinSettings.visboTaskClass = My.Settings.TaskClass
            awinSettings.visboAbbreviation = My.Settings.VISBOAbbreviation
            awinSettings.visboAmpel = My.Settings.VISBOAmpel
            awinSettings.visboAmpelText = My.Settings.VISBOAmpelText
            awinSettings.visboresponsible = My.Settings.VISBOresponsible
            awinSettings.visbodeliverables = My.Settings.VISBOdeliverables
            awinSettings.visbopercentDone = My.Settings.VISBOpercentDone
            awinSettings.visboDebug = My.Settings.VISBODebug

            dbUsername = ""
            dbPasswort = ""

            '09.11.2016: ur: Call awinsetTypenNEW("BHTC")
            Call awinsetTypen("BHTC")

            StartofCalendar = StartofCalendar.AddMonths(-12)

        Catch ex As Exception

            Call MsgBox(ex.Message)

        Finally

        End Try

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        'Call MsgBox("byebye BHTCAddIn")
        Try
            If Not fehlerBeimLoad Then

                If Not IsNothing(appInstance.Workbooks(myCustomizationFile)) Then
                    ' hier wird die Datei Projekt Tafel Customizations als aktives workbook wieder geschlossen ....

                    If awinSettings.visboDebug Then
                        Call MsgBox("Anzahl Missing-Milestones: " & missingMilestoneDefinitions.Count & vbLf & _
                               "Anzahl Missing-Phasen: " & missingPhaseDefinitions.Count)
                    End If
                   
                    appInstance.Workbooks(myCustomizationFile).Close(SaveChanges:=False)    ' CustomizationFile wird ohne Abspeichern von Änderungen geschlossen
                End If

                If Not IsNothing(appInstance.Workbooks(myLogfile)) Then
                    ' Schließen des LogFiles
                    Call logfileSchliessen()
                End If

                appInstance.ScreenUpdating = True
                Application.Quit()

            End If
        Catch ex As Exception
            Throw New ArgumentException("Fehler beim Schließen des Customization-Files")
        End Try
    End Sub

End Class
