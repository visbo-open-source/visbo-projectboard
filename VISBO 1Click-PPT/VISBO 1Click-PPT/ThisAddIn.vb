Imports System
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Windows.Forms
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions
Imports Microsoft.Office.Interop.MSProject
Imports Exception = System.Exception

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



        ''Try

        ''    awinSettings.databaseURL = My.Settings.mongoDBURL
        ''    awinSettings.databaseName = My.Settings.mongoDBname
        ''    awinSettings.globalPath = My.Settings.globalPath
        ''    awinSettings.awinPath = My.Settings.awinPath
        ''    awinSettings.visboTaskClass = My.Settings.TaskClass
        ''    awinSettings.visboAbbreviation = My.Settings.VISBOAbbreviation
        ''    awinSettings.visboAmpel = My.Settings.VISBOAmpel
        ''    awinSettings.visboAmpelText = My.Settings.VISBOAmpelText
        ''    awinSettings.visboresponsible = My.Settings.VISBOresponsible
        ''    awinSettings.visbodeliverables = My.Settings.VISBOdeliverables
        ''    awinSettings.visbopercentDone = My.Settings.VISBOpercentDone
        ''    awinSettings.visboDebug = My.Settings.VISBODebug
        ''    awinSettings.visboMapping = My.Settings.VISBOMapping
        ''    awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
        ''    If awinSettings.rememberUserPwd Then
        ''        awinSettings.userNamePWD = My.Settings.userNamePWD
        ''    End If

        ''    dbUsername = ""
        ''    dbPasswort = ""

        ''    '09.11.2016: ur: Call awinsetTypenNEW("BHTC")
        ''    Call awinsetTypen("BHTC")

        ''    StartofCalendar = StartofCalendar.AddMonths(-12)

        ''Catch ex As Exception

        ''    Call MsgBox(ex.Message)

        ''Finally

        ''End Try

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
            Catch ex As System.Exception

            End Try

            Try
                If Not IsNothing(pseudoappInstance) Then
                    ' die Excel Instanz zumachen
                    pseudoappInstance.DisplayAlerts = False
                    pseudoappInstance.Quit()
                End If
            Catch ex As System.Exception

            End Try

        Catch ex As System.Exception
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
        Catch ex As System.Exception
            If awinSettings.englishLanguage Then
                Throw New ArgumentException("Error closing the Customization-Files")
            Else
                Throw New ArgumentException("Fehler beim Schließen des Customization-Files")
            End If
        End Try
    End Sub

    Private Sub Application_ProjectBeforePublish(pj As Project, ByRef Cancel As Boolean) Handles Application.ProjectBeforePublish
        Try
            If Not awinsetTypen_Performed Then

                Try
                    pseudoappInstance = New Microsoft.Office.Interop.Excel.Application

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
                    awinSettings.visboMapping = My.Settings.VISBOMapping
                    awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
                    If awinSettings.rememberUserPwd Then
                        awinSettings.userNamePWD = My.Settings.userNamePWD
                    End If

                    dbUsername = ""
                    dbPasswort = ""

                    '09.11.2016: ur: Call awinsetTypenNEW("BHTC")
                    Call awinsetTypen("BHTC")

                    StartofCalendar = StartofCalendar.AddMonths(-12)


                Catch ex As Exception

                    Call MsgBox(ex.Message)

                Finally

                End Try
                awinsetTypen_Performed = True
            End If



            If fehlerBeimLoad Then
                If awinSettings.englishLanguage Then

                    Call MsgBox("Report of one single project cannot be executed,  " & vbLf & " 'VISBO 1Click-PPT AddIn' couldn't be loaded correctly!")
                Else

                    Call MsgBox("Einzelprojekt Report kann nicht ausgeführt werden,  " & vbLf & "da der 'VISBO 1Click-PPT AddIn' nicht korrekt geladen wurde!")
                End If

            Else

                Dim reportAuswahl As New frmReportProfil
                Dim hierarchiefenster As New frmHierarchySelection
                Dim hproj As New clsProjekt
                Dim mapProj As clsProjekt = Nothing
                Dim aktuellesDatum = Date.Now
                Dim validDatum As Date = "29.Feb.2016"
                Dim filename As String = ""

                '' ''If MsgBox("Lizenz prüfen?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                '' ''    ' ''    ''If aktuellesDatum > validDatum Then

                ' Testen, ob der User die passende Lizenz besitzt
                Dim user As String = myWindowsName
                Dim komponente As String = LizenzKomponenten(PTSWKomp.Swimlanes2)     ' Swimlanes2

                ' Lesen des Lizenzen-Files


                Dim lizenzen As clsLicences = XMLImportLicences(licFileName)

                ' Prüfen der Lizenzen
                If lizenzen.validLicence(user, komponente) Then

                    '' Set cursor as hourglass
                    Cursor.Current = Cursors.WaitCursor

                    'Call MsgBox("EPReport_Click")

                    ' Laden des aktuell geladenen Projektes und des eventuell gemappten
                    Call awinImportMSProject("BHTC", filename, hproj, mapProj, aktuellesDatum)

                    If Not IsNothing(hproj) Then
                        If hproj.name <> "" And Not IsNothing(hproj.name) Then
                            Try
                                ' Message ob Speichern erfolgt ist nur anzeigen, wenn visboMapping nicht definiert ist
                                If awinSettings.visboMapping <> "" Then
                                    Call speichereProjektToDB(hproj)
                                Else
                                    Call speichereProjektToDB(hproj, True)
                                End If

                            Catch ex As System.Exception
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Error saving of the original project to DB ")
                                Else
                                    Call MsgBox("Fehler beim Speichern des Original Projektes in DB")
                                End If
                            End Try
                        End If
                    End If

                    If Not IsNothing(mapProj) Then
                        If mapProj.name <> "" And Not IsNothing(mapProj.name) Then
                            Try
                                Call speichereProjektToDB(mapProj, True)
                            Catch ex As System.Exception
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Error saving of the mapped project to DB ")
                                Else
                                    Call MsgBox("Fehler beim Speichern des Mapping Projektes in DB")
                                End If
                            End Try
                        End If
                    Else
                        If awinSettings.visboMapping <> "" Then
                            If awinSettings.englishLanguage Then
                                Call MsgBox("Error mapping the project: no TMS - project created")
                            Else
                                Call MsgBox("Fehler beim  Mapping dieses Projektes: Kein TMS-project erstellt")
                            End If
                        End If
                    End If

                    '' Set cursor as Default
                    Cursor.Current = Cursors.Default

                Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox(" Please, contact your system administrator")
                    Else
                        Call MsgBox(" Bitte kontaktieren Sie ihren Systemadministrator")
                    End If

                End If

            End If

        Catch ex As System.Exception
            If awinSettings.englishLanguage Then
                Call MsgBox("Error with message:  " & ex.Message)
            Else
                Call MsgBox("Fehler mit Message:  " & ex.Message)
            End If

        End Try
    End Sub

    Private Sub Application_ProjectBeforeSave(pj As Project, SaveAsUi As Boolean, ByRef Cancel As Boolean) Handles Application.ProjectBeforeSave
        Try
            If Not awinsetTypen_Performed Then
                Try
                    pseudoappInstance = New Microsoft.Office.Interop.Excel.Application

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
                    awinSettings.visboMapping = My.Settings.VISBOMapping
                    awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
                    If awinSettings.rememberUserPwd Then
                        awinSettings.userNamePWD = My.Settings.userNamePWD
                    End If

                    dbUsername = ""
                    dbPasswort = ""

                    '09.11.2016: ur: Call awinsetTypenNEW("BHTC")
                    Call awinsetTypen("BHTC")

                    StartofCalendar = StartofCalendar.AddMonths(-12)


                Catch ex As Exception

                    Call MsgBox(ex.Message)

                Finally

                End Try
                awinsetTypen_Performed = True
            End If



            If fehlerBeimLoad Then
                If awinSettings.englishLanguage Then

                    Call MsgBox("Report of one single project cannot be executed,  " & vbLf & " 'VISBO 1Click-PPT AddIn' couldn't be loaded correctly!")
                Else

                    Call MsgBox("Einzelprojekt Report kann nicht ausgeführt werden,  " & vbLf & "da der 'VISBO 1Click-PPT AddIn' nicht korrekt geladen wurde!")
                End If

            Else

                Dim reportAuswahl As New frmReportProfil
                Dim hierarchiefenster As New frmHierarchySelection
                Dim hproj As New clsProjekt
                Dim mapProj As clsProjekt = Nothing
                Dim aktuellesDatum = Date.Now
                Dim validDatum As Date = "29.Feb.2016"
                Dim filename As String = ""

                '' ''If MsgBox("Lizenz prüfen?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                '' ''    ' ''    ''If aktuellesDatum > validDatum Then

                ' Testen, ob der User die passende Lizenz besitzt
                Dim user As String = myWindowsName
                Dim komponente As String = LizenzKomponenten(PTSWKomp.Swimlanes2)     ' Swimlanes2

                ' Lesen des Lizenzen-Files


                Dim lizenzen As clsLicences = XMLImportLicences(licFileName)

                ' Prüfen der Lizenzen
                If lizenzen.validLicence(user, komponente) Then

                    '' Set cursor as hourglass
                    Cursor.Current = Cursors.WaitCursor

                    'Call MsgBox("EPReport_Click")

                    ' Laden des aktuell geladenen Projektes und des eventuell gemappten
                    Call awinImportMSProject("BHTC", filename, hproj, mapProj, aktuellesDatum)

                    If Not IsNothing(hproj) Then
                        If hproj.name <> "" And Not IsNothing(hproj.name) Then
                            Try
                                ' Message ob Speichern erfolgt ist nur anzeigen, wenn visboMapping nicht definiert ist
                                If awinSettings.visboMapping <> "" Then
                                    Call speichereProjektToDB(hproj)
                                Else
                                    Call speichereProjektToDB(hproj, True)
                                End If

                            Catch ex As System.Exception
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Error saving of the original project to DB ")
                                Else
                                    Call MsgBox("Fehler beim Speichern des Original Projektes in DB")
                                End If
                            End Try
                        End If
                    End If

                    If Not IsNothing(mapProj) Then
                        If mapProj.name <> "" And Not IsNothing(mapProj.name) Then
                            Try
                                Call speichereProjektToDB(mapProj, True)
                            Catch ex As System.Exception
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Error saving of the mapped project to DB ")
                                Else
                                    Call MsgBox("Fehler beim Speichern des Mapping Projektes in DB")
                                End If
                            End Try
                        End If
                    Else
                        If awinSettings.visboMapping <> "" Then
                            If awinSettings.englishLanguage Then
                                Call MsgBox("Error mapping the project: no TMS - project created")
                            Else
                                Call MsgBox("Fehler beim  Mapping dieses Projektes: Kein TMS-project erstellt")
                            End If
                        End If
                    End If

                    '' Set cursor as Default
                    Cursor.Current = Cursors.Default

                Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox(" Please, contact your system administrator")
                    Else
                        Call MsgBox(" Bitte kontaktieren Sie ihren Systemadministrator")
                    End If

                End If

            End If

        Catch ex As System.Exception
            If awinSettings.englishLanguage Then
                Call MsgBox("Error with message:  " & ex.Message)
            Else
                Call MsgBox("Fehler mit Message:  " & ex.Message)
            End If

        End Try
    End Sub

End Class
