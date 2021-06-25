Imports ProjectboardReports
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions
Imports System.ComponentModel

Public Class frmReportProfil

    Public zeitraum_definiert As Boolean = False
    Public reportProfil As New clsReportAll

    ' für calledfrom ="MS Project"
    Public hproj As clsProjekt
    Public profileBearbeiten As New frmHierarchySelection


    ' an der aufrufenden Stelle muss hier entweder "Multiprojekt-Tafel" oder
    ' "MS Project" stehen. 
    Public calledFrom As String

    ' Liste aller vorhandenen ReportProfile
    Friend listofProfils As New SortedList(Of String, clsReportAll)


    Private Sub defineFrmButtonVisibility()

        If Me.calledFrom = "MS Project" Then

            ' für BHTC-Report wird diese Auswahlmöglichkeit derzeit nicht benötigt
            Me.rdbEPreports.Enabled = False
            Me.rdbEPreports.Visible = False
            Me.rdbMPreports.Enabled = False
            Me.rdbMPreports.Visible = False
            Me.changeProfil.Visible = True
            Me.changeProfil.Enabled = True
            Me.zeitLabel.Visible = True
            Me.vonDate.Visible = True
            Me.bisDate.Visible = True
            Me.statusLabel.Visible = True
        Else
            Me.rdbEPreports.Enabled = True
            Me.rdbEPreports.Visible = True
            Me.rdbMPreports.Enabled = True
            Me.rdbMPreports.Visible = True
            Me.changeProfil.Visible = False
            Me.changeProfil.Enabled = False

            Me.zeitLabel.Visible = False
            Me.vonDate.Visible = False
            Me.bisDate.Visible = False
            Me.statusLabel.Visible = False
        End If

        If awinSettings.englishLanguage Then
            ' auf Englisch darstellen 
            Me.Text = "Select a Report Profile to create a report"
            rdbEPreports.Text = "Singleproject Reports"
            rdbMPreports.Text = "Multiproject Reports"
            ReportErstellen.Text = "Create Report"
            changeProfil.Text = "Modify Report Profile"
            zeitLabel.Text = "Timespan:"
        End If



    End Sub


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

        Call defineFrmButtonVisibility()

        zeitraum_definiert = (showRangeLeft > 0 And showRangeRight > showRangeLeft)


        If Me.calledFrom = "MS Project" Then


            Try

                '' ''Dim i As Integer

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

                    If listOfFiles.Count < 1 Then

                        ' erzeuge ein Dummy-ReportPRofil

                        Dim dmyRepProfil As New clsReport
                        '' 'Call createDummyReportProfil(dmyRepProfil)

                        dmyRepProfil.Projects.Clear()
                        dmyRepProfil.Projects.Add(1, hproj.name)

                        dmyRepProfil.calcRepVonBis(vonDate.Value, bisDate.Value)


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

                    If listOfFiles.Count > 0 Then
                        RepProfilListbox.SelectedIndex = 0
                    End If


                Else
                    Dim msgTxt As String = ""
                    If awinSettings.englishLanguage Then
                        msgTxt = "Error: no Report Profile existing!"
                    Else
                        msgTxt = "Fehler: es existiert kein ReportProfil"
                    End If
                    Throw New ArgumentException(msgTxt)

                End If

                Me.statusLabel.Visible = False

            Catch ex As Exception
                'Call MsgBox(ex.Message)
                Me.statusLabel.Text = ex.Message
                Me.statusLabel.Visible = True
            End Try

        ElseIf Me.calledFrom = "Multiprojekt-Tafel" Then



            Try
                If currentReportProfil.name = "Last" Then
                    ' Profil von letztem Report unter Name "Last" speichern
                    Call XMLExportReportProfil(currentReportProfil)

                End If
            Catch ex As Exception

            End Try
            Try

                ' hier müssen die ReportProfile aus dem Directory ausgelesen werden und zur Auswahl angeboten werden

                Dim dirName As String
                Dim dateiName As String
                Dim profilName As String = ""

                dirName = awinPath & ReportProfileOrdner


                If My.Computer.FileSystem.DirectoryExists(dirName) Then


                    Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirName)


                    ' Existiert kein ReportProfil.XML, so wird ein Dummy.xml erzeugt und anschließend eingelesen

                    If listOfFiles.Count < 1 Then

                        Dim msgTxt As String
                        If awinSettings.englishLanguage Then
                            msgTxt = "no Report Profile existing!"
                        Else
                            msgTxt = " Es existiert noch kein Report-Profil! "
                        End If
                        Call MsgBox(msgTxt)

                    Else

                        For k As Integer = 1 To listOfFiles.Count

                            dateiName = listOfFiles.Item(k - 1)
                            If dateiName.Contains(".xml") Then

                                Try

                                    Dim hstr() As String
                                    hstr = Split(dateiName, ".xml", 2)
                                    Dim hhstr() As String
                                    hhstr = Split(hstr(0), "\")
                                    profilName = hhstr(hhstr.Length - 1)

                                    Dim hreportAll As clsReportAll = XMLImportReportProfil(profilName)

                                    If listofProfils.ContainsKey(profilName) Then
                                        listofProfils.Remove(profilName)
                                    End If
                                    listofProfils.Add(profilName, hreportAll)


                                Catch ex As Exception
                                    'Throw New ArgumentException("ReportProfil '" & profilName & "' konnte nicht eingelesen werden!")
                                    Dim msgTxt As String
                                    If awinSettings.englishLanguage Then
                                        msgTxt = "Report Profile '" & profilName & "' could not be read!"
                                    Else
                                        msgTxt = "ReportProfil '" & profilName & "' konnte nicht eingelesen werden!"
                                    End If
                                    Call MsgBox(msgTxt)
                                End Try

                            End If

                        Next k

                        ' anzeige löschen
                        RepProfilListbox.Items.Clear()

                        ' Anzeigen der Profile, abhängig vom gecheckten Radiobutton

                        ' Report mit Constellation - Multiprojektreport
                        If rdbMPreports.Checked Then

                            If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                                ' alles ok , bereits gesetzt 
                                zeitraum_definiert = True
                            Else
                                zeitraum_definiert = False
                                showRangeLeft = ShowProjekte.getMinMonthColumn
                                showRangeRight = ShowProjekte.getMaxMonthColumn
                                Call awinShowtimezone(showRangeLeft, showRangeRight, True)
                            End If

                            For Each kvp In listofProfils

                                If kvp.Value.isMpp Then
                                    ' Profil profilName in Auswahl eintragen
                                    RepProfilListbox.Items.Add(kvp.Value.name)

                                End If
                            Next

                        End If

                        ' Einzelprojektreport
                        If rdbEPreports.Checked Then

                            If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                                ' alles ok , bereits gesetzt 
                                zeitraum_definiert = True
                            Else
                                zeitraum_definiert = False
                                showRangeLeft = 0
                                showRangeRight = 0
                            End If

                            For Each kvp In listofProfils

                                If Not kvp.Value.isMpp Then
                                    ' Profil profilName in Auswahl eintragen
                                    RepProfilListbox.Items.Add(kvp.Value.name)

                                End If
                            Next

                        End If


                            If listOfFiles.Count > 0 Then
                                '  RepProfilListbox.SelectedIndex = 1
                            End If

                    End If


                Else
                    Dim errTxt As String
                    If awinSettings.englishLanguage Then
                        errTxt = "no Report Profile existing!"
                    Else
                        errTxt = " Es existiert noch kein Report-Profil! "
                    End If
                    Throw New ArgumentException(errTxt)

                End If

                'Me.zeitLabel.Visible = False
                'Me.vonDate.Visible = False
                'Me.bisDate.Visible = False
                '' ur: 18.02.2019: für Allianz entfernt, da keine Phasen und Meilensteine ausgewählt werden, höchsten Rollen
                ''Me.changeProfil.Visible = True
                'Me.changeProfil.Visible = False
                'Me.statusLabel.Visible = False

            Catch ex As Exception
                'Call MsgBox(ex.Message)
                Me.statusLabel.Text = ex.Message
                Me.statusLabel.Visible = True
            End Try

            If selectedProjekte.Count > 0 Then
                Me.rdbEPreports.Checked = True
            Else
                Me.rdbMPreports.Checked = True
            End If


        End If

    End Sub
    Private Sub RepProfilListbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RepProfilListbox.SelectedIndexChanged


        Dim reportProfilName As String = RepProfilListbox.Text
        Dim IndSelItem As Integer = RepProfilListbox.SelectedIndex

        ''Call MsgBox("Lesen des XML-Files " & reportProfilName & ".xml")

        If Me.calledFrom = "MS Project" Then
            ' '' Einlesen des ausgewählten ReportProfils
            reportProfil = XMLImportReportProfil(reportProfilName)

            If Not IsNothing(reportProfil) Then

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

            End If
        ElseIf Me.calledFrom = "Multiprojekt-Tafel" Then

            ' '' Einlesen des ausgewählten ReportProfils
            reportProfil = XMLImportReportProfil(reportProfilName)
            currentReportProfil = reportProfil

            If Not IsNothing(reportProfil) Then
                ToolTipProfil.Show(reportProfil.description, RepProfilListbox, 6000)
            End If

        End If



    End Sub


    Private Sub vonDate_ValueChanged(sender As Object, e As EventArgs) Handles vonDate.ValueChanged


    End Sub

    Private Sub bisDate_ValueChanged(sender As Object, e As EventArgs) Handles bisDate.ValueChanged


    End Sub

    Private Sub ReportErstellen_Click(sender As Object, e As EventArgs) Handles ReportErstellen.Click

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        If Me.calledFrom = "MS Project" Then

            Try
                Dim noPhExist As Boolean = True
                Dim noMSExist As Boolean = True
                Dim tmpSortedList As New SortedList(Of String, String)

                If RepProfilListbox.Text <> "" Then

                    Dim reportProfilName As String = RepProfilListbox.Text

                    'Call MsgBox("Lesen des XML-Files " & reportProfilName & ".xml")

                    ' Einlesen des ausgewählten ReportProfils
                    reportProfil = XMLImportReportProfil(reportProfilName)


                    ' Test, ob die in reportProfil definierten Meilenstein und Phasen in hproj enthalten sind

                    If Not (reportProfil.Phases.Count = 0 And reportProfil.Milestones.Count = 0) Then

                        For Each kvp As KeyValuePair(Of String, String) In reportProfil.Phases
                            noPhExist = noPhExist And Not hproj.containsPhase(kvp.Key, True)
                        Next

                        For Each kvp As KeyValuePair(Of String, String) In reportProfil.Milestones
                            noMSExist = noMSExist And Not hproj.containsMilestone(kvp.Key, True)
                        Next
                    Else
                        noPhExist = False
                        noMSExist = False
                    End If


                    If noPhExist And noMSExist Then
                        Dim msgTxt As String = "Achtung: Projekt '" & hproj.name & _
                            "' enthält die ausgewählten Phasen und Meilensteine nicht!"
                        If awinSettings.englishLanguage Then
                            msgTxt = "Warning: Project '" & hproj.name & _
                            "' does not contain the selected phases nor selected milestones!"
                        End If
                        Call MsgBox(msgTxt)
                    Else

                        If Not IsNothing(reportProfil) Then

                            'Call MsgBox("ReportErstellen")
                            Try
                                reportProfil.calcRepVonBis(vonDate.Value, bisDate.Value)
                            Catch ex As Exception
                                Throw New ArgumentException(ex.Message)
                            End Try


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

                            'Call PPTstarten()
                            ''ur: 02042019
                            ''Call BGworkerReportBHTC_new(reportProfil)

                            ''Me.statusLabel.Text = "...finished"


                            BGworkerReportBHTC.RunWorkerAsync(reportProfil)


                        Else
                            Dim msgTxt As String = "ausgewähltes Report-Profil enthält Fehler !"
                            If awinSettings.englishLanguage Then
                                msgTxt = "errors when reading the selected Report Profile !"
                            End If
                            Call MsgBox(msgTxt)
                        End If
                    End If
                Else
                    Dim msgTxt As String = "Es wurde noch kein Report-Profil ausgewählt !"
                    If awinSettings.englishLanguage Then
                        msgTxt = "please select a Report Profile!"
                    End If

                    Call MsgBox(msgTxt)

                End If

            Catch ex As Exception
                'Call MsgBox(ex.Message)
                Me.statusLabel.Text = ex.Message
                Me.statusLabel.Visible = True
            End Try


        ElseIf Me.calledFrom = "Multiprojekt-Tafel" Then
            Try
                If RepProfilListbox.Text <> "" And ShowProjekte.Count > 0 Then

                    Dim reportProfilName As String = RepProfilListbox.Text

                    ' Einlesen des ausgewählten ReportProfils
                    reportProfil = XMLImportReportProfil(reportProfilName)

                    ' ausgewähltes ReportPRofil in current-Variable speichern
                    currentReportProfil = reportProfil

                    If Not IsNothing(reportProfil) Then

                        If reportProfil.isMpp Then


                            If Not (showRangeLeft > 0 And showRangeRight > showRangeLeft) Then  ' Zeitraum wurde nicht gesetzt

                                ' Es muss ein Zeitraum ausgewählt sein
                                Me.statusLabel.Visible = True
                                Dim msgTxt As String = "bitte zuerst einen Zeitraum auswählen!"
                                If awinSettings.englishLanguage Then
                                    msgTxt = "please select a timespan!"
                                End If
                                Me.statusLabel.Text = msgTxt

                                Call MsgBox(msgTxt)
                                ' MyBase.Close()

                            Else

                                ' alles ok , Zeitraum bereits gesetzt 
                                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

                                Me.statusLabel.Visible = True
                                Me.statusLabel.Text = "...started"
                                Me.ReportErstellen.Visible = False
                                Me.ReportErstellen.Enabled = False

                                BGWorkerReportGen.RunWorkerAsync(reportProfil)
                            End If


                        Else   ' Profil für Einzelprojekt-Bericht ausgewählt
                            ' Es muss mindestens ein Projekt selektiert sein
                            If selectedProjekte.Count < 1 Then

                                Me.statusLabel.Visible = True

                                Dim msgTxt As String = "bitte zuerst Projekte selektieren!"
                                If awinSettings.englishLanguage Then
                                    msgTxt = "please select one or more projects!"
                                End If

                                Me.statusLabel.Text = msgTxt

                                Call MsgBox(msgTxt)
                                'MyBase.Close()
                            Else
                                Me.statusLabel.Visible = True
                                Me.statusLabel.Text = "...started"
                                Me.ReportErstellen.Visible = False
                                Me.ReportErstellen.Enabled = False

                                BGWorkerReportGen.RunWorkerAsync(reportProfil)
                            End If

                        End If
                    End If

                Else
                    Dim msgTxt As String = "Es wurde noch kein Report-Profil ausgewählt ! oder " & vbLf & _
                        "Es sind keine Projekte geladen !"
                    If awinSettings.englishLanguage Then
                        msgTxt = "no Report Profile or no projects selected!"
                    End If
                    Call MsgBox(msgTxt)

                End If


            Catch ex As Exception
                'Call MsgBox(ex.Message)
                Me.statusLabel.Text = ex.Message
                Me.statusLabel.Visible = True
            End Try

        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU

    End Sub

    Private Sub changeProfil_Click(sender As Object, e As EventArgs) Handles changeProfil.Click

        Try
            If Me.calledFrom = "MS Project" Then


                ''ist bereits erfolgt ''
                '' '' Einlesen des ausgewählten ReportProfils 
                '' '' ''reportProfil = XMLImportReportProfil(RepProfilListbox.Text)

                If Not IsNothing(reportProfil) Then

                    reportProfil.Projects.Clear()
                    reportProfil.Projects.Add(1, hproj.name)

                    Try
                        reportProfil.calcRepVonBis(vonDate.Value, bisDate.Value)
                    Catch ex As Exception
                        'Call MsgBox(ex.Message)
                        Me.statusLabel.Text = ex.Message
                        Me.statusLabel.Visible = True
                        Exit Sub
                    End Try




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

                Else
                    Throw New ArgumentException("Fehler: es ist kein ReportProfil geladen")

                End If    ' von if not isnothing(reportProfil)


            ElseIf Me.calledFrom = "Multiprojekt-Tafel" Then

                ''ist bereits erfolgt ''
                '' '' Einlesen des ausgewählten ReportProfils 
                '' '' ''reportAllProfil = XMLImportReportAllProfil(reportProfilName)


                ' TO DO :nun muss das Formular zum Ändern gefüllt und angezeigt werden

                If Not IsNothing(reportProfil) Then
                    ToolTipProfil.Show(reportProfil.description, RepProfilListbox, 6000)


                    Me.statusLabel.Visible = False

                    ' frmHierarchySelection aufrufen für BHTC
                    Call PBBBHTCHierarchySelAction("Multiprojekt-Tafel", reportProfil)


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

                        ' Existiert kein ReportProfil.XML, so wird ein Dummy.xml erzeugt und anschließend eingelesen

                        If listOfFiles.Count < 1 Then

                            Dim msgTxt As String = " Es existiert noch kein Report-Profil! "
                            If awinSettings.englishLanguage Then
                                msgTxt = "no Report Profile existing!"
                            End If
                            Call MsgBox(msgTxt)

                        Else

                            For k As Integer = 1 To listOfFiles.Count

                                dateiName = listOfFiles.Item(k - 1)
                                If dateiName.Contains(".xml") Then

                                    Try

                                        Dim hstr() As String
                                        hstr = Split(dateiName, ".xml", 2)
                                        Dim hhstr() As String
                                        hhstr = Split(hstr(0), "\")
                                        profilName = hhstr(hhstr.Length - 1)

                                        Dim hreportAll As clsReportAll = XMLImportReportProfil(profilName)

                                        If listofProfils.ContainsKey(profilName) Then
                                            listofProfils.Remove(profilName)
                                        End If
                                        listofProfils.Add(profilName, hreportAll)


                                    Catch ex As Exception
                                        'Throw New ArgumentException("ReportProfil '" & profilName & "' konnte nicht eingelesen werden!")
                                        Dim msgTxt As String = "ReportProfil '" & profilName & "' konnte nicht eingelesen werden!"
                                        If awinSettings.englishLanguage Then
                                            msgTxt = "Report Profile '" & profilName & "' could not be read!"
                                        End If
                                        Call MsgBox(msgTxt)
                                    End Try

                                End If

                            Next k

                            ' anzeige löschen
                            RepProfilListbox.Items.Clear()

                            ' Anzeigen der Profile, abhängig vom gecheckten Radiobutton

                            ' Report mit Constellation - Multiprojektreport
                            If rdbMPreports.Checked Then

                                For Each kvp In listofProfils

                                    If kvp.Value.isMpp Then
                                        ' Profil profilName in Auswahl eintragen
                                        RepProfilListbox.Items.Add(kvp.Value.name)

                                    End If
                                Next

                            End If

                            ' Einzelprojektreport
                            If rdbEPreports.Checked Then

                                For Each kvp In listofProfils

                                    If Not kvp.Value.isMpp Then
                                        ' Profil profilName in Auswahl eintragen
                                        RepProfilListbox.Items.Add(kvp.Value.name)

                                    End If
                                Next

                            End If


                        End If
                        'RepVorlagenListBox ist nun  neu aufgebaut


                    Else
                        Throw New ArgumentException("Fehler: Directory ReportProfile existiert nicht")

                    End If
                Else
                    Throw New ArgumentException("Fehler: es ist kein ReportProfil geladen")
                End If

            End If   ' end of Me.calledFrom = "Multiprojekt-Tafel"

        Catch ex As Exception
            'Call MsgBox(ex.Message)
            Me.statusLabel.Text = ex.Message
            Me.statusLabel.Visible = True
        End Try
    End Sub
    ' ur: 02042019 wird nicht mehr aufgerufen, da es von der Antwortzeit keinen Effekt gehabt hat.
    'Private Sub BGworkerReportBHTC_new(ByVal reportProfil As clsReportAll)



    '    'Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
    '    Dim worker As BackgroundWorker = nothing
    '    ' ''Dim vorlagenDateiName As String = CType(e.Argument, String)
    '    'Dim reportProfil As clsReportAll = CType(e.Argument, clsReportAll)
    '    Dim zeilenhoehe As Double = 0.0     ' zeilenhöhe muss für alle Projekte gleich sein, daher mit übergeben
    '    Dim legendFontSize As Single = 0.0  ' FontSize der Legenden der Schriftgröße des Projektnamens angepasst


    '    Dim selectedPhases As New Collection
    '    Dim selectedMilestones As New Collection
    '    Dim selectedRoles As New Collection
    '    Dim selectedCosts As New Collection
    '    Dim selectedBUs As New Collection
    '    Dim selectedTypes As New Collection

    '    selectedPhases = copySortedListtoColl(reportProfil.Phases)
    '    selectedMilestones = copySortedListtoColl(reportProfil.Milestones)
    '    selectedRoles = copySortedListtoColl(reportProfil.Roles)
    '    selectedCosts = copySortedListtoColl(reportProfil.Costs)
    '    selectedBUs = copySortedListtoColl(reportProfil.BUs)
    '    selectedTypes = copySortedListtoColl(reportProfil.Typs)

    '    ' für BHTC immer true
    '    'reportProfil.ExtendedMode = True
    '    '' für BHTC immer false
    '    'reportProfil.Ampeln = False
    '    'reportProfil.AllIfOne = False
    '    'reportProfil.FullyContained = False
    '    'reportProfil.SortedDauer = False
    '    'reportProfil.ProjectLine = False
    '    'reportProfil.UseOriginalNames = False

    '    With awinSettings

    '        ' tk Änderung 5.4. wird für Darstellung Projekt auf Multiprojekt-Tafel benötigt; hier nicht setzen 
    '        '.drawProjectLine = True
    '        .mppExtendedMode = reportProfil.ExtendedMode
    '        .mppOnePage = reportProfil.OnePage
    '        .mppShowAllIfOne = reportProfil.AllIfOne
    '        .mppShowAmpel = reportProfil.Ampeln
    '        .mppShowLegend = reportProfil.Legend
    '        .mppShowMsDate = reportProfil.MSDate
    '        .mppShowMsName = reportProfil.MSName
    '        .mppShowPhDate = reportProfil.PhDate
    '        .mppShowPhName = reportProfil.PhName
    '        .mppShowProjectLine = reportProfil.ProjectLine
    '        .mppSortiertDauer = reportProfil.SortedDauer
    '        .mppVertikalesRaster = reportProfil.VLinien
    '        .mppFullyContained = reportProfil.FullyContained
    '        .mppShowHorizontals = reportProfil.ShowHorizontals
    '        .mppUseAbbreviation = reportProfil.UseAbbreviation
    '        .mppUseOriginalNames = reportProfil.UseOriginalNames
    '        .mppKwInMilestone = reportProfil.KwInMilestone
    '    End With


    '    ' Report wird von Projekt hproj, das vor Aufruf des Formulars in hproj gespeichert wurde erzeugt

    '    showRangeLeft = getColumnOfDate(vonDate.Value)
    '    showRangeRight = getColumnOfDate(bisDate.Value)

    '    Try
    '        Dim vorlagendateiname As String = awinPath & RepProjectVorOrdner & "\" & reportProfil.PPTTemplate

    '        If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

    '            Dim projname As String = reportProfil.Projects.ElementAt(0).Value

    '            Dim hproj As clsProjekt = ShowProjekte.getProject(projname)


    '            Call createPPTSlidesFromProject(hproj, vorlagendateiname,
    '                                            selectedPhases, selectedMilestones,
    '                                            selectedRoles, selectedCosts,
    '                                            selectedBUs, selectedTypes, True,
    '                                            True, zeilenhoehe, legendFontSize,
    '                                            worker, Nothing)

    '        Else


    '        End If


    '    Catch ex As Exception
    '        Call MsgBox("Fehler: " & vbLf & ex.Message)
    '    End Try

    'End Sub




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
        Dim reportProfil As clsReportAll = CType(e.Argument, clsReportAll)
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
        'reportProfil.ExtendedMode = True
        '' für BHTC immer false
        'reportProfil.Ampeln = False
        'reportProfil.AllIfOne = False
        'reportProfil.FullyContained = False
        'reportProfil.SortedDauer = False
        'reportProfil.ProjectLine = False
        'reportProfil.UseOriginalNames = False

        With awinSettings

            ' tk Änderung 5.4. wird für Darstellung Projekt auf Multiprojekt-Tafel benötigt; hier nicht setzen 
            '.drawProjectLine = True
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
                                                True, zeilenhoehe, legendFontSize, _
                                                worker, e)

            Else


            End If


        Catch ex As Exception
            Call MsgBox("Fehler: " & vbLf & ex.Message)
        End Try

    End Sub

    Private Sub BGWorkerReportGen_DoWork(sender As Object, e As DoWorkEventArgs) Handles BGWorkerReportGen.DoWork



        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim reportProfil As clsReportAll = CType(e.Argument, clsReportAll)
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

        With awinSettings

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
            .mppProjectsWithNoMPmayPass = reportProfil.projectsWithNoMPmayPass

        End With




        Try
            If Not reportProfil.isMpp Then


                Dim vorlagendateiname As String = awinPath & RepProjectVorOrdner & "\" & reportProfil.PPTTemplate
                If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

                    ' Alle selektierten Projekte reporten
                    '' ''For Each kvp In selectedProjekte.Liste

                    '' ''    hproj = kvp.Value



                    '' ''    Call createPPTSlidesFromProject(hproj, vorlagendateiname, _
                    '' ''                                    selectedPhases, selectedMilestones, _
                    '' ''                                    selectedRoles, selectedCosts, _
                    '' ''                                    selectedBUs, selectedTypes, True, _
                    '' ''                                    True, zeilenhoehe, legendFontSize, _
                    '' ''                                    worker, e)



                    '' ''Next
                    appInstance.EnableEvents = False
                    'appInstance.ScreenUpdating = False

                    Call createPPTReportFromProjects(vorlagendateiname, _
                                                     selectedPhases, selectedMilestones, _
                                                     selectedRoles, selectedCosts, _
                                                     selectedBUs, selectedTypes, _
                                                     worker, e)

                End If
            Else

                If Not (showRangeLeft > 0 And showRangeRight > showRangeLeft) Then

                    showRangeLeft = getColumnOfDate(reportProfil.VonDate)
                    showRangeRight = getColumnOfDate(reportProfil.BisDate)

                End If

                Dim vorlagendateiname As String = awinPath & RepPortfolioVorOrdner & "\" & reportProfil.PPTTemplate
                If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

                    Call createPPTSlidesFromConstellation(vorlagendateiname, _
                                                          selectedPhases, selectedMilestones, _
                                                          selectedRoles, selectedCosts, _
                                                          selectedBUs, selectedTypes, True, _
                                                          worker, e)

                End If

            End If



        Catch ex As Exception
            Call MsgBox("Fehler: " & vbLf & ex.Message)
        End Try

        appInstance.EnableEvents = True
    End Sub

    Private Sub BGWorkerReportGen_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BGWorkerReportGen.ProgressChanged

        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusLabel.Text = CType(re.Result, String)

    End Sub

    Private Sub BGWorkerReportGen_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGWorkerReportGen.RunWorkerCompleted

        Me.statusLabel.Text = "...done"
        Me.statusLabel.Visible = True
        Me.ReportErstellen.Visible = True
        Me.ReportErstellen.Enabled = True
        Me.RepProfilListbox.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Arrow

        ' hier evt. noch schließen und Abspeichern des Reports von PPT

        appInstance.ScreenUpdating = True

    End Sub

    Private Sub rdbEPreports_CheckedChanged(sender As Object, e As EventArgs) Handles rdbEPreports.CheckedChanged


        If rdbEPreports.Checked Then

            If Me.calledFrom = "Multiprojekt-Tafel" Then
                Try

                    RepProfilListbox.Items.Clear()

                    For Each kvp In listofProfils

                        If Not kvp.Value.isMpp Then
                            ' Profil profilName in Auswahl eintragen
                            RepProfilListbox.Items.Add(kvp.Value.name)

                        End If
                    Next

                    ' es wurde noch kein Projekt selektiert
                    If selectedProjekte.Count < 1 Then

                        Me.statusLabel.Visible = True

                        Dim msgTxt As String = "bitte zuerst Projekte selektieren!"
                        If awinSettings.englishLanguage Then
                            msgTxt = "please select one or more projects!"
                        End If

                        Me.statusLabel.Text = msgTxt

                        Call MsgBox(msgTxt)
                    End If

                Catch ex As Exception
                    'Throw New ArgumentException("Fehler beim Filtern")
                    Me.statusLabel.Text = ex.Message
                    Me.statusLabel.Visible = True
                End Try


                'Me.zeitLabel.Visible = False
                'Me.vonDate.Visible = False
                'Me.bisDate.Visible = False
                'Me.changeProfil.Visible = True
                'Me.statusLabel.Visible = False
            End If

        End If

    End Sub


    Private Sub rdbMPreports_CheckedChanged(sender As Object, e As EventArgs) Handles rdbMPreports.CheckedChanged

        If rdbMPreports.Checked Then

            If Me.calledFrom = "Multiprojekt-Tafel" Then
                Try

                    RepProfilListbox.Items.Clear()

                    For Each kvp In listofProfils

                        If kvp.Value.isMpp Then
                            ' Profil profilName in Auswahl eintragen
                            RepProfilListbox.Items.Add(kvp.Value.name)

                        End If
                    Next

                    ' es ist kein Zeitraum selektiert
                    If Not (showRangeLeft > 0 And showRangeRight > showRangeLeft) Then
                        showRangeLeft = ShowProjekte.getMinMonthColumn
                        showRangeRight = ShowProjekte.getMaxMonthColumn
                        Call awinShowtimezone(showRangeLeft, showRangeRight, True)
                    End If


                Catch ex As Exception
                    'Throw New ArgumentException("Fehler beim Filtern")
                    Me.statusLabel.Text = ex.Message
                    Me.statusLabel.Visible = True
                End Try


                'Me.zeitLabel.Visible = False
                'Me.vonDate.Visible = False
                'Me.bisDate.Visible = False
                'Me.changeProfil.Visible = True
                'Me.statusLabel.Visible = False

            End If
        Else
            If zeitraum_definiert = False Then
                ' zu Beginn war kein zeitraum definiert, also zurücksetzen
                Call awinShowtimezone(showRangeLeft, showRangeRight, False)
                showRangeLeft = 0
                showRangeRight = 0
            End If
        End If

    End Sub


End Class