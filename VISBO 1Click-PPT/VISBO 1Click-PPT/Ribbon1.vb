
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions

Public Class Ribbon1


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
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
            awinSettings.visboServer = My.Settings.VISBOServer
            awinSettings.proxyURL = My.Settings.proxyServerURL
            awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
            If awinSettings.rememberUserPwd Then
                awinSettings.userNamePWD = My.Settings.userNamePWD
            End If

            dbUsername = ""
            dbPasswort = ""

            Call awinsetTypen("BHTC")

            StartofCalendar = StartofCalendar.AddMonths(-12)


            If awinSettings.englishLanguage Then
                DBspeichern.Label = "Save to DB"
                EinzelprojektReport.Label = "Report of one Project"
                Einstellung.Label = "Settings"
            Else
                DBspeichern.Label = "Speichern in DB"
                EinzelprojektReport.Label = "Einzelprojekt Report"
                Einstellung.Label = "Einstellungen"
            End If


        Catch ex As Exception

            Call MsgBox(ex.Message)

        Finally

        End Try

    End Sub

    Private Sub EinzelprojektReport_Click(sender As Object, e As RibbonControlEventArgs) Handles EinzelprojektReport.Click

        Try
            If Not awinsetTypen_Performed Then

                '' Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor
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
                    awinSettings.visboServer = My.Settings.VISBOServer
                    awinSettings.proxyURL = My.Settings.proxyServerURL
                    awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
                    If awinSettings.rememberUserPwd Then
                        awinSettings.userNamePWD = My.Settings.userNamePWD
                    End If

                    dbUsername = ""
                    dbPasswort = ""

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

                    Call MsgBox("Report of one single project is not executable,  " & vbLf & " 'VISBO 1Click-PPT AddIn' couldn't be loaded correctly!")
                Else

                    Call MsgBox("Einzelprojekt Report kann nicht ausgeführt werden,  " & vbLf & "da der 'VISBO 1Click-PPT AddIn' nicht korrekt geladen wurde!")
                End If

            Else


                Dim reportAuswahl As New frmReportProfil
                Dim hierarchiefenster As New frmHierarchySelection
                Dim returnvalue As DialogResult
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

                    ' Laden des aktuell geladenen Projektes
                    Call awinImportMSProject("BHTC", filename, hproj, mapProj, aktuellesDatum)

                    If Not IsNothing(hproj) Then
                        If hproj.name <> "" And Not IsNothing(hproj.name) Then
                            Try
                                Call speichereProjektToDB(hproj)
                            Catch ex As Exception
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Error saving the project to DB ")
                                Else
                                    Call MsgBox("Fehler beim Speichern in DB")
                                End If

                            End Try
                        End If
                    End If

                    '' Set cursor as default
                    Cursor.Current = Cursors.Default

                    If Not IsNothing(mapProj) Then
                        If mapProj.name <> "" And Not IsNothing(mapProj.name) Then
                            Try
                                Call speichereProjektToDB(mapProj)
                            Catch ex As Exception
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Error saving the project to DB ")
                                Else
                                    Call MsgBox("Fehler beim Speichern in DB")
                                End If
                            End Try
                        End If

                        reportAuswahl.calledFrom = "MS Project"
                        reportAuswahl.hproj = mapProj
                        reportAuswahl.calledFrom = "MS Project"
                        returnvalue = reportAuswahl.ShowDialog
                    Else
                        If Not IsNothing(hproj) Then

                            reportAuswahl.calledFrom = "MS Project"
                            reportAuswahl.hproj = hproj
                            reportAuswahl.calledFrom = "MS Project"
                            returnvalue = reportAuswahl.ShowDialog
                        End If
                    End If

                Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox("User " & myWindowsName & " doesn't have any License!" _
                                    & vbLf & " Please, contact your system administrator")
                    Else
                        Call MsgBox("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                                & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")

                    End If

                End If



            End If


        Catch ex As Exception
            If awinSettings.englishLanguage Then
                Call MsgBox(" Please, contact your system administrator")
                Throw New ArgumentException(" Please, contact your system administrator")
            Else
                Call MsgBox(" Bitte kontaktieren Sie ihren Systemadministrator")
                Throw New ArgumentException(" Bitte kontaktieren Sie ihren Systemadministrator")
            End If


        End Try
    End Sub

    Private Sub DBspeichern_Click(sender As Object, e As RibbonControlEventArgs) Handles DBspeichern.Click
        Try
            If Not awinsetTypen_Performed Then
                '' Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor
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
                    awinSettings.visboServer = My.Settings.VISBOServer
                    awinSettings.proxyURL = My.Settings.proxyServerURL
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
                awinsetTypen_Performed = True
            End If



            If fehlerBeimLoad Then
                If awinSettings.englishLanguage Then

                    Call MsgBox("Report of one single project cannot be executed,  " & vbLf & " 'VISBO 1Click-PPT AddIn' couldn't be loaded correctly!")
                Else

                    Call MsgBox("Einzelprojekt Report kann nicht ausgeführt werden,  " & vbLf & "da der 'VISBO 1Click-PPT AddIn' nicht korrekt geladen wurde!")
                End If

            Else
                '' Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor


                ' Dim reportAuswahl As New frmReportProfil
                ' Dim hierarchiefenster As New frmHierarchySelection
                Dim hproj As New clsProjekt
                Dim mapProj As clsProjekt = Nothing
                Dim aktuellesDatum = Date.Now
                'Dim validDatum As Date = "29.Feb.2016"
                Dim filename As String = ""

                ' Testen, ob der User die passende Lizenz besitzt
                Dim user As String = myWindowsName
                Dim komponente As String = LizenzKomponenten(PTSWKomp.Swimlanes2)     ' Swimlanes2

                ' Lesen des Lizenzen-Files
                Dim lizenzen As clsLicences = XMLImportLicences(licFileName)

                ' Prüfen der Lizenzen
                If lizenzen.validLicence(user, komponente) Then


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

                            Catch ex As Exception
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
                            Catch ex As Exception
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

        Catch ex As Exception

            If awinSettings.englishLanguage Then
                Call MsgBox("Error with message:  " & ex.Message)
            Else
                Call MsgBox("Fehler mit Message:  " & ex.Message)
            End If

        End Try

    End Sub

    Private Sub Einstellung_Click(sender As Object, e As RibbonControlEventArgs) Handles Einstellung.Click

    End Sub

    Private Sub Ribbon1_Close(sender As Object, e As EventArgs) Handles Me.Close
        My.Settings.Save()
    End Sub
End Class
