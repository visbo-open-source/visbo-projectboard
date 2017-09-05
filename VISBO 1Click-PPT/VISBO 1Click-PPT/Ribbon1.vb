
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub EinzelprojektReport_Click(sender As Object, e As RibbonControlEventArgs) Handles EinzelprojektReport.Click

        Try
            If fehlerBeimLoad Then
                Call MsgBox("Einzelprojekt Report kann nicht ausgeführt werden,  " & vbLf & "da der 'VISBO 1Click-PPT AddIn' nicht korrekt geladen wurde!")
            Else

                Dim reportAuswahl As New frmReportProfil
                Dim hierarchiefenster As New frmHierarchySelection
                Dim returnvalue As DialogResult
                Dim hproj As New clsProjekt
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


                    'Call MsgBox("EPReport_Click")

                    ' Laden des aktuell geladenen Projektes
                    Call awinImportMSProject("BHTC", filename, hproj, aktuellesDatum)

                    If hproj.name <> "" And Not IsNothing(hproj.name) Then
                        reportAuswahl.calledFrom = "MS Project"
                        reportAuswahl.hproj = hproj
                        reportAuswahl.calledFrom = "MS Project"
                        returnvalue = reportAuswahl.ShowDialog
                    End If
                Else
                    Call MsgBox("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                                & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")

                End If


                '' ''    Else    ' ohne Lizenzprüfung

                '' ''    ' Laden des aktuell geladenen Projektes
                '' ''    Call awinImportMSProject("BHTC", filename, hproj, aktuellesDatum)

                '' ''    If hproj.name <> "" And Not IsNothing(hproj.name) Then
                '' ''        reportAuswahl.hproj = hproj
                '' ''        returnvalue = reportAuswahl.ShowDialog
                '' ''    End If

                '' ''End If ' Ende if Lizenzprüfung


            End If


        Catch ex As Exception

            Throw New ArgumentException(" Bitte kontaktieren Sie ihren Systemadministrator")
            Call MsgBox(" Bitte kontaktieren Sie ihren Systemadministrator")
        End Try
    End Sub

    Private Sub DBspeichern_Click(sender As Object, e As RibbonControlEventArgs) Handles DBspeichern.Click
        Try


            If fehlerBeimLoad Then
                Call MsgBox("Einzelprojekt Report kann nicht ausgeführt werden,  " & vbLf & "da der 'VISBO 1Click-PPT AddIn' nicht korrekt geladen wurde!")
            Else

                Dim reportAuswahl As New frmReportProfil
                Dim hierarchiefenster As New frmHierarchySelection
                Dim hproj As New clsProjekt
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


                    'Call MsgBox("EPReport_Click")

                    ' Laden des aktuell geladenen Projektes
                    Call awinImportMSProject("BHTC", filename, hproj, aktuellesDatum)

                    If hproj.name <> "" And Not IsNothing(hproj.name) Then
                        Try
                            ' LOGIN in DB machen
                            If awinSettings.databaseURL <> "" And awinSettings.databaseName <> "" Then

                                noDB = False

                                If dbUsername = "" Or dbPasswort = "" Then

                                    ' ur: 23.01.2015: Abfragen der Login-Informationen
                                    loginErfolgreich = loginProzedur()


                                    If Not loginErfolgreich Then
                                        Call logfileSchreiben("LOGIN cancelled ...", "", -1)
                                        Call MsgBox("LOGIN cancelled ...")
                                    Else
                                        Dim speichernInDBOk As Boolean = False
                                        Dim identical As Boolean = False
                                        Try
                                            speichernInDBOk = storeSingleProjectToDB(hproj, identical)
                                            If speichernInDBOk Then
                                                If Not identical Then
                                                    Call MsgBox("Projekt '" & hproj.name & "' wurde erfolgreich in der Datenbank gespeichert")
                                                Else
                                                    Call MsgBox("Projekt '" & hproj.name & "' ist identisch mit der aktuellen Version in der DB")
                                                End If
                                            Else
                                                Call MsgBox("Fehler beim Speichern des aktuell geladenen Projektes")
                                            End If


                                        Catch ex As Exception
                                            Throw New ArgumentException("Fehler beim Speichern von Projekt: " & hproj.name)
                                        End Try

                                    End If

                                Else

                                    If testLoginInfo_OK(dbUsername, dbPasswort) Then
                                        Dim speichernInDBOk As Boolean
                                        Dim identical As Boolean = False
                                        Try
                                            speichernInDBOk = storeSingleProjectToDB(hproj, identical)
                                            If speichernInDBOk Then
                                                If Not identical Then
                                                    Call MsgBox("Projekt '" & hproj.name & "' wurde erfolgreich in der Datenbank gespeichert")
                                                Else
                                                    Call MsgBox("Projekt '" & hproj.name & "' ist identisch mit der aktuellen Version in der DB")
                                                End If
                                            Else
                                                Call MsgBox("Fehler beim Speichern des aktuell geladenen Projektes")
                                            End If
                                        Catch ex As Exception
                                            Throw New ArgumentException("Fehler beim Speichern von Projekt: " & hproj.name)
                                        End Try
                                    Else
                                        Call MsgBox("LOGIN fehlerhaft ...")
                                    End If

                                End If


                            End If


                        Catch ex As Exception
                            Call MsgBox(ex.Message)
                        End Try
                      
                    End If
                Else
                    Call MsgBox("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                                & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")

                End If


                '' ''    Else    ' ohne Lizenzprüfung

                '' ''    ' Laden des aktuell geladenen Projektes
                '' ''    Call awinImportMSProject("BHTC", filename, hproj, aktuellesDatum)

                '' ''    If hproj.name <> "" And Not IsNothing(hproj.name) Then
                '' ''        reportAuswahl.hproj = hproj
                '' ''        returnvalue = reportAuswahl.ShowDialog
                '' ''    End If

                '' ''End If ' Ende if Lizenzprüfung


            End If

        Catch ex As Exception
            Call MsgBox("Fehler mit Message:  " & ex.Message)
        End Try
    End Sub

    Private Sub Einstellung_Click(sender As Object, e As RibbonControlEventArgs) Handles Einstellung.Click

    End Sub
End Class
