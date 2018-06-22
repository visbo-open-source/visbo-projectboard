
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
                                Call MsgBox("Fehler beim Speichern in DB")
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
                                Call MsgBox("Fehler beim Speichern in DB")
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
                    Call MsgBox("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                                & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")

                End If



            End If


        Catch ex As Exception
            Call MsgBox(" Bitte kontaktieren Sie ihren Systemadministrator")
            Throw New ArgumentException(" Bitte kontaktieren Sie ihren Systemadministrator")

        End Try
    End Sub

    Private Sub DBspeichern_Click(sender As Object, e As RibbonControlEventArgs) Handles DBspeichern.Click
        Try


            If fehlerBeimLoad Then
                Call MsgBox("Einzelprojekt Report kann nicht ausgeführt werden,  " & vbLf & "da der 'VISBO 1Click-PPT AddIn' nicht korrekt geladen wurde!")
            Else


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

                    '' Set cursor as hourglass
                    Cursor.Current = Cursors.WaitCursor

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
                                Call MsgBox("Fehler beim Speichern von OriginalProjekt in DB")
                            End Try
                        End If
                    End If

                    If Not IsNothing(mapProj) Then
                        If mapProj.name <> "" And Not IsNothing(mapProj.name) Then
                            Try
                                Call speichereProjektToDB(mapProj, True)
                            Catch ex As Exception
                                Call MsgBox("Fehler beim Speichern des MappedProjekt in DB")
                            End Try
                        End If
                    End If

                    '' Set cursor as Default
                    Cursor.Current = Cursors.Default

                Else
                    Call MsgBox("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                                & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")

                End If

            End If

        Catch ex As Exception

            Call MsgBox("Fehler mit Message:  " & ex.Message)
        End Try

    End Sub

    Private Sub Einstellung_Click(sender As Object, e As RibbonControlEventArgs) Handles Einstellung.Click

    End Sub

    Private Sub Ribbon1_Close(sender As Object, e As EventArgs) Handles Me.Close
        My.Settings.Save()
    End Sub
End Class
