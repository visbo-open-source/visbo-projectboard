Imports ClassLibrary1
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions
Imports System.ComponentModel

Public Class frmReportProfil

    Public reportProfil As New clsReport
    Public hproj As clsProjekt
    Public profileBearbeiten As New frmHierarchySelection

    Private Sub frmReportProfil_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Not BGworkerReportBHTC.IsBusy Then
            MyBase.Close()
        Else
            Select Case MsgBox("Wollen Sie das Fenster wirklich schließen?", vbQuestion Or vbYesNo Or vbDefaultButton2, "beenden ?")
                Case vbYes
                    Me.Dispose() 'Fenster wird geschlossen

                Case vbNo
                    e.Cancel = True 'Fenster wird nicht geschlossen
            End Select
        End If

    End Sub
     

    Private Sub RepProfilListbox_load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer

        '' ''Dim minDate As Date = Date.MaxValue
        '' ''Dim maxDate As Date = Date.MinValue

        '' ''Dim anzproj As Integer = ShowProjekte.Count
        ' '' '' alle geladenen Projekte in ReportProfil aufnehmen
        '' ''For i = 1 To anzproj

        '' ''    Dim hhproj As clsProjekt = ShowProjekte.getProject(i)

        '' ''    If DateDiff(DateInterval.Day, minDate, hhproj.startDate) < 0 Then
        '' ''        minDate = hhproj.startDate

        '' ''        If minDate < StartofCalendar Then
        '' ''            minDate = StartofCalendar
        '' ''        End If
        '' ''    End If

        '' ''    If DateDiff(DateInterval.Day, maxDate, hhproj.endeDate) > 0 Then
        '' ''        maxDate = hhproj.endeDate
        '' ''    End If

        '' ''Next

        vonDate.Value = hproj.startDate
        bisDate.Value = hproj.endeDate

        ' hier müssen die ReportProfile aus dem Directory ausgelesen werden und zur Auswahl angeboten werden

        Dim dirName As String
        Dim dateiName As String
        Dim profilName As String = ""

        dirName = awinPath & ReportProfileOrdner


        If My.Computer.FileSystem.DirectoryExists(dirName) Then


            Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirName)


            ' Existiert kein ReportProfil.XML, so wird ein Dummy.xml erzeugt und anschließend eingelesen

            If listofFiles.count < 1 Then

                ' erzeuge ein Dummy-ReportPRofil

                Dim dmyRepProfil As New clsreport
                '' 'Call createDummyReportProfil(dmyRepProfil)

                dmyRepProfil.Projects.Clear()
                dmyRepProfil.Projects.Add(1, hproj.name)

                dmyRepProfil.VonDate = vonDate.Value
                dmyRepProfil.BisDate = bisDate.Value

                ' Schreiben des Dummy ReportProfils
                Call XMLExportReportProfil(dmyRepProfil)

                'erneut Files auf Directory lesen
                listOfFiles = My.Computer.FileSystem.GetFiles(dirName)

            End If

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

            If listofFiles.count > 0 Then
                RepProfilListbox.SelectedIndex = 0
            End If


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

        ''Call MsgBox("Lesen des XML-Files " & reportProfilName & ".xml")

        ' '' Einlesen des ausgewählten ReportProfils
        reportProfil = XMLImportReportProfil(reportProfilName)
        '' ''If Not IsNothing(reportProfil) Then
        '' ''    vonDate.Value = reportProfil.VonDate
        '' ''    bisDate.Value = reportProfil.BisDate
        '' ''End If


        reportProfil.Projects.Clear()
        reportProfil.Projects.Add(1, hproj.name)


        ' für BHTC immer true
        reportProfil.ExtendedMode = True
        ' für BHTC immer false
        reportProfil.Ampeln = False
        reportProfil.AllIfOne = False
        reportProfil.FullyContained = False
        reportProfil.SortedDauer = False
        reportProfil.ProjectLine = False
        reportProfil.UseOriginalNames = False


    End Sub

    Private Sub vonDate_ValueChanged(sender As Object, e As EventArgs) Handles vonDate.ValueChanged

        If Not IsNothing(reportProfil) Then
            reportProfil.VonDate = vonDate.Value
            reportProfil.BisDate = bisDate.Value
        End If
    End Sub

    Private Sub bisDate_ValueChanged(sender As Object, e As EventArgs) Handles bisDate.ValueChanged

        If Not IsNothing(reportProfil) Then
            reportProfil.VonDate = vonDate.Value
            reportProfil.BisDate = bisDate.Value
        End If

        'Call MsgBox("Fehler: Endedatum des Reports liegt von dem BeginnDatum" & vbLf & "Bitte korrigieren Sie das")

    End Sub

    Private Sub ReportErstellen_Click(sender As Object, e As EventArgs) Handles ReportErstellen.Click


        Dim tmpSortedList As New SortedList(Of String, String)

        If RepProfilListbox.Text <> "" Then

            Dim reportProfilName As String = RepProfilListbox.Text

            'Call MsgBox("Lesen des XML-Files " & reportProfilName & ".xml")

            ' Einlesen des ausgewählten ReportProfils
            reportProfil = XMLImportReportProfil(reportProfilName)

            If Not IsNothing(reportProfil) Then

                'Call MsgBox("ReportErstellen")

                reportProfil.VonDate = vonDate.Value
                reportProfil.BisDate = bisDate.Value

                Dim anzproj As Integer = ShowProjekte.Count
                ' alle geladenen Projekte in ReportProfil aufnehmen
                ' ''For i = 1 To anzproj

                ' ''    Dim hilfsproj As clsProjekt = ShowProjekte.getProject(i)
                ' ''    reportProfil.Projects.Add(i, hilfsproj.name)

                ' ''Next

                'Call MsgBox("Es wurden " & CStr(anzproj) & " Projekte in  ShowProjekte eingelesen." & vbLf _
                '        & "Report wird für das aktuell geladene Projekt erstellt: " & hproj.name)

                reportProfil.Projects.Clear()
                reportProfil.Projects.Add(1, hproj.name)

                ' für BHTC immer true
                reportProfil.ExtendedMode = True
                ' für BHTC immer false
                reportProfil.Ampeln = False
                reportProfil.AllIfOne = False
                reportProfil.FullyContained = False
                reportProfil.SortedDauer = False
                reportProfil.ProjectLine = False
                reportProfil.UseOriginalNames = False



                'Call MsgBox("Report erstellen mit Projekt " & hproj.name & "von " & vonDate.Value.ToString & " bis " & bisDate.Value.ToString & " Reportprofil " & reportProfilName)
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

                Me.statusLabel.Visible = True
                Me.statusLabel.Text = "...started"


                BGworkerReportBHTC.RunWorkerAsync(reportProfil)

            Else
                Call MsgBox("ausgewähltes Report-Profil enthält Fehler !")
            End If

        Else
            Call MsgBox("Es wurde noch kein Report-Profil ausgewählt !")
        End If
    End Sub

    Private Sub changeProfil_Click(sender As Object, e As EventArgs) Handles changeProfil.Click


        ''ist bereits erfolgt ''
        '' '' Einlesen des ausgewählten ReportProfils 
        '' '' ''reportProfil = XMLImportReportProfil(RepProfilListbox.Text)

        If Not IsNothing(reportProfil) Then

            reportProfil.Projects.Clear()
            reportProfil.Projects.Add(1, hproj.name)

            reportProfil.VonDate = vonDate.Value
            reportProfil.BisDate = bisDate.Value

        End If


        Me.statusLabel.Visible = False

        ' frmHierarchySelection aufrufen für BHTC
        Call PBBBHTCHierarchySelAction("BHTC", reportProfil)


        'RepVorlagenListBox neu aufbauen, falls ein oder mehrere ReportProfile gespeichert wurden.
        ' hier müssen die ReportProfile erneut aus dem Directory ausgelesen werden und zur Auswahl angeboten werden

        Dim selectedItem As Object = RepProfilListbox.SelectedItem

        RepProfilListbox.Items.Clear()  ' entfernt alle elemente aus Listbox um sie dann neu aufzubauen

        Dim dirName As String
        Dim dateiName As String
        Dim profilName As String = ""

        dirName = awinPath & ReportProfileOrdner


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
            RepProfilListbox.SelectedItem = selectedItem

        Else
            Throw New ArgumentException("Fehler: es existiert kein ReportProfil")

        End If
        'RepVorlagenListBox ist nun  neu aufgebaut

    End Sub




    Private Sub BGworkerReportBHTC_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BGworkerReportBHTC.ProgressChanged

        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusLabel.Text = CType(re.Result, String)

    End Sub

    Private Sub BGworkerReportBHTC_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGworkerReportBHTC.RunWorkerCompleted

        '' ''With Me.AbbrButton
        '' ''    .Text = ""
        '' ''    .Visible = False
        '' ''    .Enabled = False
        '' ''    .Left = .Left + 40
        '' ''End With


        Me.statusLabel.Text = "...done"
        Me.statusLabel.Visible = True
        Me.ReportErstellen.Visible = True
        Me.ReportErstellen.Enabled = True
        Me.RepProfilListbox.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Arrow

        ' hier evt. noch schließen und Abspeichern des Reports von PPT

    End Sub

    Private Sub BGworkerReportBHTC_DoWork(sender As Object, e As DoWorkEventArgs) Handles BGworkerReportBHTC.DoWork



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

        ' für BHTC immer true
        reportProfil.ExtendedMode = True
        ' für BHTC immer false
        reportProfil.Ampeln = False
        reportProfil.AllIfOne = False
        reportProfil.FullyContained = False
        reportProfil.SortedDauer = False
        reportProfil.ProjectLine = False
        reportProfil.UseOriginalNames = False

        With awinSettings

            .drawProjectLine = True
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
            .mppShowHorizontals = reportProfil.ShowHorizontals
            .mppUseAbbreviation = reportProfil.UseAbbreviation
            .mppUseOriginalNames = reportProfil.UseOriginalNames
            .mppKwInMilestone = reportProfil.KwInMilestone
        End With


        ' Report wird von Projekt hproj, das vor Aufruf des Formulars in hproj gespeichert wurde erzeugt

        showRangeLeft = getColumnOfDate(vonDate.Value)
        showRangeRight = getColumnOfDate(bisDate.Value)

        Try
            Dim vorlagendateiname As String = awinPath & RepProjectVorOrdner & "\" & reportProfil.PPTTemplate

            If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

                Dim projname As String = reportProfil.Projects.ElementAt(0).Value

                Dim hproj As clsProjekt = ShowProjekte.getProject(projname)


                Call createPPTSlidesFromProject(hproj, vorlagendateiname, _
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
End Class