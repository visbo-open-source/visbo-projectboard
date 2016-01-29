Imports ClassLibrary1
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions
Imports System.ComponentModel

Public Class frmReportProfil
    Public reportProfil As New clsReport
    Public hproj As clsProjekt
    Public profileBearbeiten As New frmHierarchySelection

    Private Sub RepProfilListbox_load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer

        Dim minDate As Date = Date.MaxValue
        Dim maxDate As Date = Date.MinValue

        Dim anzproj As Integer = ShowProjekte.Count
        ' alle geladenen Projekte in ReportProfil aufnehmen
        For i = 1 To anzproj

            Dim hproj As clsProjekt = ShowProjekte.getProject(i)

            If DateDiff(DateInterval.Day, minDate, hproj.startDate) < 0 Then
                minDate = hproj.startDate

                If minDate < StartofCalendar Then
                    minDate = StartofCalendar
                End If
            End If

            If DateDiff(DateInterval.Day, maxDate, hproj.endeDate) > 0 Then
                maxDate = hproj.endeDate
            End If

        Next

        vonDate.Value = minDate
        bisDate.Value = maxDate

        ' hier müssen die ReportProfile aus dem Directory ausgelesen werden und zur Auswahl angeboten werden

        Dim dirName As String
        Dim dateiName As String
        Dim profilName As String = ""

        dirName = awinPath & "requirements\ReportProfile"
        

        If My.Computer.FileSystem.DirectoryExists(dirName) Then


            Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirName)

            For k As Integer = 1 To listOfFiles.Count

                dateiName = listOfFiles.Item(k - 1)
                If dateiName.Contains(".xml") Then

                    Try
                        
                        Dim hstr() As String
                        hstr = Split(dateiName, ".xml", 2)
                        Dim hhstr() As String
                        hhstr = Split(hstr(0), "\")
                        profilName = hhstr(hhstr.Length - 1)
                        RepProfilListbox.Items.Add(profilName)

                    Catch ex As Exception

                    End Try

                End If

            Next k
            RepProfilListbox.SelectedIndex = 0

        Else
            Throw New ArgumentException("Fehler: es existiert kein ReportProfil")

        End If
        'For i = 0 To 30

        '    Try

        '        RepProfilListbox.Items.Add("aaa" & CStr(i))

        '    Catch ex As Exception

        '    End Try

        '    RepProfilListbox.SelectedItem = RepProfilListbox.Items.Count
        'Next i

    End Sub
    Private Sub RepProfilListbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RepProfilListbox.SelectedIndexChanged


        Dim reportProfilName As String = RepProfilListbox.Text

        Call MsgBox("Lesen des XML-Files" & reportProfilName & ".xml")

        ' Einlesen des ausgewählten ReportProfils
        reportProfil = XMLImportReportProfil(reportProfilName)


    End Sub

    Private Sub vonDate_ValueChanged(sender As Object, e As EventArgs) Handles vonDate.ValueChanged
        Dim reportVondate As Date = vonDate.Value
    End Sub

    Private Sub bisDate_ValueChanged(sender As Object, e As EventArgs) Handles bisDate.ValueChanged

        Dim reportBisdate As Date = bisDate.Value
        'Call MsgBox("Fehler: Endedatum des Reports liegt von dem BeginnDatum" & vbLf & "Bitte korrigieren Sie das")

    End Sub

    Private Sub ReportErstellen_Click(sender As Object, e As EventArgs) Handles ReportErstellen.Click


        Dim tmpSortedList As New SortedList(Of String, String)
        If RepProfilListbox.Text <> "" Then

            Dim reportProfilName As String = RepProfilListbox.Text

            Call MsgBox("Lesen des XML-Files" & reportProfilName & ".xml")

            ' Einlesen des ausgewählten ReportProfils
            reportProfil = XMLImportReportProfil(reportProfilName)

            Call MsgBox("ReportErstellen")


            Dim anzproj As Integer = ShowProjekte.Count
            ' alle geladenen Projekte in ReportProfil aufnehmen
            ' ''For i = 1 To anzproj

            ' ''    Dim hilfsproj As clsProjekt = ShowProjekte.getProject(i)
            ' ''    reportProfil.Projects.Add(i, hilfsproj.name)

            ' ''Next

            Call MsgBox("Es wurden " & CStr(anzproj) & " Projekte in eingelesen." & vbLf _
                         & "Report wird für das aktuell geladene Projekt erstellt: " & hproj.name)

            reportProfil.Projects.Clear()
            reportProfil.Projects.Add(1, hproj.name)



            With awinSettings

                .drawProjectLine = True
                .eppExtendedMode = reportProfil.ExtendedMode
                .mppExtendedMode = reportProfil.ExtendedMode
                .mppOnePage = reportProfil.OnePage
                .mppShowAllIfOne = reportProfil.AllIfOne
                .mppShowAmpel = reportProfil.Ampeln
                .mppShowLegend = reportProfil.Legend
                .mppShowMsDate = reportProfil.MSDate
                .mppShowMsName = reportProfil.MSName
                .mppShowPhDate = reportProfil.PhDate
                .mppShowPhName = reportProfil.PhName
                .mppShowProjectLine = reportProfil.ProjectLine
                .mppSortiertDauer = reportProfil.SortedDauer
                .mppVertikalesRaster = reportProfil.VLinien
                .mppFullyContained = reportProfil.FullyContained
                ' .mppShowHorizontals = reportProfil.ShowHorizontals
                ' .mppUseAbbreviation = reportProfil.UseAbbreviation
                ' .mppUseOriginalNames = reportProfil.UseOriginalNames
            End With


            Call MsgBox("Report erstellen mit Projekt " & hproj.name & "von " & vonDate.Value.ToString & " bis " & bisDate.Value.ToString & " Reportprofil " & reportProfilName)
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

            ' Alternativ ohne Background Worker

            BackgroundWorker3.RunWorkerAsync(reportProfil)
        Else
            Call MsgBox("Es wurde noch kein Report-Profil ausgewählt !")
        End If
    End Sub

    Private Sub changeProfil_Click(sender As Object, e As EventArgs) Handles changeProfil.Click


        Dim tmpcollection As New Collection
        Dim tmpSortedList As New SortedList(Of String, String)
        tmpcollection.Add("name1", "name1")
        tmpcollection.Add("name2", "name2")
        tmpcollection.Add("name3", "name3")
        reportProfil.Phases = copyColltoSortedList(tmpcollection)


        ' '' ''Dim returnvalue As System.Windows.Forms.DialogResult

        ' '' ''With profileBearbeiten
        ' '' ''    .Text = "Report-Profil bearbeiten"
        ' '' ''    returnvalue = .ShowDialog()
        ' '' ''End With

        ' '' ''If returnvalue = System.Windows.Forms.DialogResult.OK Then



        ' '' ''    ' reportProfil gemäß auswahl ändern und abspeichern



        ' '' ''    reportProfil.PPTTemplate = "\\KOYTEK-NAS\backup\Projekt-Tafel Folder\BHTC\requirements\ReportTemplatesProject\Alle PlanElemente Querformat DIN A3.pptx"
        ' '' ''    ' ''reportProfil.name = reportProfilName
        ' '' ''    '' ''reportProfil.Phases = New SortedList(Of String, String)
        ' '' ''    ' ''For i = 5 To 1 Step -1
        ' '' ''    ' ''    reportProfil.Phases.Add("name" & i.ToString, "name" & i.ToString)
        ' '' ''    ' ''Next
        ' '' ''    reportProfil.Phases = tmpSortedList
        ' '' ''    ' ''reportProfil.Milestones = tmpSortedList
        ' '' ''    ' ''reportProfil.BUs = tmpSortedList
        ' '' ''    ' ''reportProfil.Typs = tmpSortedList
        ' '' ''    ' ''reportProfil.Roles = tmpSortedList
        ' '' ''    ' ''reportProfil.Costs = tmpSortedList
        ' '' ''    ' ''reportProfil.VonDate = vonDate.Value
        ' '' ''    ' ''reportProfil.BisDate = bisDate.Value
        ' '' ''    ' ''reportProfil.PPTTemplate = "\\KOYTEK-NAS\backup\Projekt-Tafel Folder\BHTC\requirements\ReportTemplatesProject\Alle PlanElemente Querformat DIN A3.pptx"
        ' '' ''    ' ''reportProfil.OnePage = True
        ' '' ''    ' ''reportProfil.AllIfOne = False
        ' '' ''    ' ''reportProfil.Ampeln = True
        ' '' ''    ' ''reportProfil.CalendarVonDate = Date.Now
        ' '' ''    ' ''reportProfil.CalendarBisDate = Date.MaxValue
        ' '' ''    ' ''reportProfil.ExtendedMode = True
        ' '' ''    ' ''reportProfil.isMpp = False
        ' '' ''    reportProfil.Legend = False
        ' '' ''    ' ''reportProfil.MSDate = False
        ' '' ''    ' ''reportProfil.MSName = False
        ' '' ''    ' ''reportProfil.PhDate = False
        ' '' ''    ' ''reportProfil.PhName = True
        ' '' ''    ' ''reportProfil.ProjectLine = True
        ' '' ''    ' ''reportProfil.SortedDauer = False
        ' '' ''    reportProfil.VLinien = True
        ' '' ''    ' ''reportProfil.FullyContained = False
        ' '' ''    ' Datei mit ReportProfil schreiben

        ' '' ''    Call XMLExportReportProfil(reportProfil)
        ' '' ''    'Call XMLreportProfilExport(reportProfil, xmlReportsfile, False, protokoll)
        ' '' ''End If


    End Sub

    Private Sub BackgroundWorker3_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        ' ''Dim vorlagenDateiName As String = CType(e.Argument, String)
        Dim reportProfil As clsReport = CType(e.Argument, clsReport)
        Dim zeilenhoehe As Double = 0.0     ' zeilenhöhe muss für alle Projekte gleich sein, daher mit übergeben
        Dim legendFontSize As Single = 0.0  ' FontSize der Legenden der Schriftgröße des Projektnamens angepasst

        Dim selectedPhases As New Collection
        Dim selectedMilestones As New Collection
        Dim selectedRoles As New Collection
        Dim selectedCosts As New Collection
        Dim selectedBUs As New Collection
        Dim selectedTypes As New Collection

        selectedPhases = copySortedListtoColl(reportProfil.Phases)
        selectedMilestones = copySortedListtoColl(reportProfil.Milestones)
        selectedRoles = copySortedListtoColl(reportProfil.Roles)
        selectedCosts = copySortedListtoColl(reportProfil.Costs)
        selectedBUs = copySortedListtoColl(reportProfil.BUs)
        selectedTypes = copySortedListtoColl(reportProfil.Typs)


        ' Report wird von Projekt hproj, das vor Aufruf des Formulars in hproj gespeichert wurde erzeugt

        showRangeLeft = CInt(DateDiff(DateInterval.Month, StartofCalendar, vonDate.Value))
        showRangeRight = CInt(DateDiff(DateInterval.Month, StartofCalendar, bisDate.Value))

        Try

            If reportProfil.PPTTemplate.Contains(RepProjectVorOrdner) Then

                Call createPPTSlidesFromProject(hproj, reportProfil.PPTTemplate, _
                                                selectedPhases, selectedMilestones, _
                                                selectedRoles, selectedCosts, _
                                                selectedBUs, selectedTypes, True, _
                                                True, zeilenhoehe, _
                                                legendFontSize, _
                                                worker, e)


                ' ''Call createPPTReportFromProjects(vorlagenDateiName, _
                ' ''                                   selectedPhases, selectedMilestones, _
                ' ''                                   selectedRoles, selectedCosts, _
                ' ''                                   selectedBUs, selectedTyps, _
                ' ''                                   worker, e)
            Else

                ''Call createPPTSlidesFromConstellation(reportProfil.PPTTemplate, _
                ''                                reportProfil.Phases, reportProfil.Milestones, _
                ''                                reportProfil.Roles, reportProfil.Costs, _
                ''                                reportProfil.BUs, reportProfil.Typs, True, _
                ''                                worker, e)
            End If


        Catch ex As Exception
            Call MsgBox("Fehler " & ex.Message)
        End Try

    End Sub

    Private Sub BackgroundWorker3_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker3.ProgressChanged

        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusLabel.Text = CType(re.Result, String)

    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted

        ' ''With Me.AbbrButton
        ' ''    .Text = ""
        ' ''    .Visible = False
        ' ''    .Enabled = False
        ' ''    .Left = .Left + 40
        ' ''End With


        Me.statusLabel.Text = "...done"
        Me.statusLabel.Visible = True
        Me.ReportErstellen.Visible = True
        Me.ReportErstellen.Enabled = True
        Me.RepProfilListbox.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Arrow

        ' hier evt. noch schließen und Abspeichern des Reports von PPT

    End Sub

End Class