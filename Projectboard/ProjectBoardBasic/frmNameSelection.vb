Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports ClassLibrary1
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms

Public Class frmNameSelection

    ' kann von ausserhalb gesetzt werden; gibt an ob das ganze Portfolio angezeigt werden soll
    ' oder nur die selektierten Projekte 


    Friend menuOption As Integer
    Friend actionCode As Integer


    ' hier steht ggf die ButtonID drin
    Friend ribbonButtonID As String = ""
    


    Private allMilestones As New Collection
    Private allPhases As New Collection
    Private allCosts As New Collection
    Private allRoles As New Collection
    Private allBUs As New Collection
    Private allTyps As New Collection


    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection
    Private selectedCosts As New Collection
    Private selectedRoles As New Collection
    Private selectedBUs As New Collection
    Private selectedTyps As New Collection

    'Private sKeyMilestones As String = ""
    'Private sKeyPhases As String = ""
    'Private sKeyCosts As String = ""
    'Private sKeyRoles As String = ""

    Private backgroundRunning As Boolean = False

    Private Enum PTauswahlTyp
        phase = 0
        meilenstein = 1
        Rolle = 2
        Kostenart = 3
        BusinessUnit = 4
        ProjektTyp = 5
    End Enum


    ' bestimmt, ob und wie die einzelnen Formular Elemente in Abhängigkeit von menuoption angezeigt werden sollen 
    Private Sub defineFrmButtonVisibility()

        If awinSettings.englishLanguage Then
            chkbxOneChart.Text = "all in one chart"
            Label1.Text = ""
            einstellungen.Text = "Settings"
            labelPPTVorlage.Text = "Powerpoint Template"
            AbbrButton.Text = "Reset Selection"
        End If


        With Me
            If .menuOption = PTmenue.filterdefinieren Then
                If awinSettings.englishLanguage Then
                    .Text = "define Database Filter"
                Else
                    .Text = "Datenbank Filter definieren"
                End If


                If .actionCode = PTTvActions.loadPV Or _
                    .actionCode = PTTvActions.loadPVS Or _
                    .actionCode = PTTvActions.delAllExceptFromDB Or _
                    .actionCode = PTTvActions.delFromDB Then
                    If awinSettings.englishLanguage Then
                        .OKButton.Text = "Apply"
                    Else
                        .OKButton.Text = "Anwenden"
                    End If

                Else
                    If awinSettings.englishLanguage Then
                        .OKButton.Text = "Store"
                    Else
                        .OKButton.Text = "Speichern"
                    End If

                End If

                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .rdbRoles.Enabled = True
                .rdbCosts.Enabled = True

                .rdbBU.Visible = True
                .pictureBU.Visible = True

                .rdbTyp.Visible = True
                .pictureTyp.Visible = True

                .einstellungen.Visible = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports 
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False
                .einstellungen.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True
                If awinSettings.englishLanguage Then
                    .filterLabel.Text = "Name of Filter"
                Else
                    .filterLabel.Text = "Name des Filters"
                End If


                ' Auswahl Speichern
                .auswSpeichern.Visible = False
                .auswSpeichern.Enabled = False

            ElseIf menuOption = PTmenue.sessionFilterDefinieren Then

                If awinSettings.englishLanguage Then
                    .Text = "Define Session Filter"
                    .OKButton.Text = "Apply"
                    .filterLabel.Text = "Name of Filter"
                Else
                    .Text = "Session Filter definieren"
                    .OKButton.Text = "Anwenden"
                    .filterLabel.Text = "Name des Filters"
                End If


                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .rdbRoles.Enabled = True
                .rdbCosts.Enabled = True

                .rdbBU.Visible = True
                .pictureBU.Visible = True

                .rdbTyp.Visible = True
                .pictureTyp.Visible = True

                .einstellungen.Visible = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports 
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False
                .einstellungen.Visible = False

                ' Filter
                .filterDropbox.Visible = False
                .filterLabel.Visible = False


                ' Auswahl Speichern
                .auswSpeichern.Visible = False
                .auswSpeichern.Enabled = False


            ElseIf menuOption = PTmenue.visualisieren Then

                If awinSettings.englishLanguage Then
                    .Text = "Visualize Phases and Milestones"
                    .OKButton.Text = "Visualize"
                    .filterLabel.Text = "Selection"
                Else
                    .Text = "Phasen- und Meilensteine visualisieren"
                    .OKButton.Text = "Anzeigen"
                    .filterLabel.Text = "Auswahl"
                End If


                .statusLabel.Text = ""
                .statusLabel.Visible = True


                .rdbBU.Visible = False
                .pictureBU.Visible = False
                .rdbTyp.Visible = False
                .pictureTyp.Visible = False
                .rdbRoles.Visible = False
                .pictureRoles.Visible = False
                .rdbCosts.Visible = False
                .pictureCosts.Visible = False

                ' Leistbarkeits-Charts
                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports 
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False
                .einstellungen.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True


            ElseIf menuOption = PTmenue.leistbarkeitsAnalyse Then

                If awinSettings.englishLanguage Then
                    If ribbonButtonID = "PTMEC1" Then
                        .Text = "Create Role-/Cost-Charts"
                    Else
                        .Text = "Create Feasibility Charts"
                    End If
                    .filterLabel.Text = "Selection"
                    .OKButton.Text = "Create Charts"
                Else
                    If ribbonButtonID = "PTMEC1" Then
                        .Text = "Rollen-/Kosten-Charts erstellen"
                    Else
                        .Text = "Leistbarkeits-Charts erstellen"
                    End If
                    .filterLabel.Text = "Auswahl"
                    .OKButton.Text = "Charts erstellen"
                End If



                .statusLabel.Text = ""
                .statusLabel.Visible = True

                If ribbonButtonID = "PTMEC1" Then
                    .rdbPhases.Visible = False
                    .picturePhasen.Visible = False
                    .rdbMilestones.Visible = False
                    .pictureMilestones.Visible = False
                End If

                .rdbBU.Visible = False
                .pictureBU.Visible = False
                .rdbTyp.Visible = False
                .pictureTyp.Visible = False

                .rdbRoles.Visible = True
                .pictureRoles.Visible = True
                .rdbCosts.Visible = True
                .pictureCosts.Visible = True

                ' Leistbarkeits-Charts
                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = True

                ' Reports 
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True


            ElseIf menuOption = PTmenue.einzelprojektReport Then

                If awinSettings.englishLanguage Then
                    .Text = "Create Project-/Variant Report"
                    .OKButton.Text = "Create Report"
                    .filterLabel.Text = "Selection"
                Else
                    .Text = "Projekt-Varianten Report erzeugen"
                    .OKButton.Text = "Bericht erstellen"
                    .filterLabel.Text = "Auswahl"
                End If


                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .rdbRoles.Enabled = False
                .rdbCosts.Enabled = False

                .rdbBU.Enabled = False
                .rdbBU.Visible = False
                .pictureBU.Visible = False

                .rdbTyp.Enabled = False
                .rdbTyp.Visible = False
                .pictureTyp.Visible = False


                .einstellungen.Visible = True

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                .repVorlagenDropbox.Visible = True
                .labelPPTVorlage.Visible = True

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True



            ElseIf menuOption = PTmenue.multiprojektReport Then

                If awinSettings.englishLanguage Then
                    .Text = "Create Multiproject Reports"
                    .OKButton.Text = "Create Report"
                    .filterLabel.Text = "Selection"
                Else
                    .Text = "Multiprojekt Reports erzeugen"
                    .OKButton.Text = "Bericht erstellen"
                    .filterLabel.Text = "Auswahl"
                End If


                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .rdbRoles.Enabled = True
                .rdbCosts.Enabled = True

                .rdbBU.Enabled = False
                .rdbBU.Visible = False
                .pictureBU.Visible = False

                .rdbTyp.Enabled = False
                .rdbTyp.Visible = False
                .pictureTyp.Visible = False


                .einstellungen.Visible = True

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                .repVorlagenDropbox.Visible = True
                .labelPPTVorlage.Visible = True

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True


            ElseIf menuOption = PTmenue.excelExport Then

                If awinSettings.englishLanguage Then
                    .Text = "Export to Excel"
                    .OKButton.Text = "Export"
                    .filterLabel.Text = "Selection"
                Else
                    .Text = "Export nach Excel"
                    .OKButton.Text = "Export"
                    .filterLabel.Text = "Auswahl"
                End If


                .statusLabel.Text = ""

                .rdbRoles.Enabled = False
                .rdbCosts.Enabled = False

                .rdbBU.Visible = True
                .pictureBU.Visible = True

                .rdbTyp.Visible = True
                .pictureTyp.Visible = True

                .einstellungen.Visible = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True


            ElseIf menuOption = PTmenue.vorlageErstellen Then

                If awinSettings.englishLanguage Then
                    .Text = "Create modular templates"
                    .OKButton.Text = "Create Template"
                    .filterLabel.Text = "Selection"
                Else
                    .Text = "modulare Vorlagen erzeugen"
                    .OKButton.Text = "Vorlage erstellen"
                    .filterLabel.Text = "Auswahl"
                End If


                .statusLabel.Text = ""

                .rdbRoles.Enabled = False
                .rdbCosts.Enabled = False

                .rdbBU.Visible = False
                .pictureBU.Visible = False

                .rdbTyp.Visible = False
                .pictureTyp.Visible = False

                .einstellungen.Visible = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True


            ElseIf menuOption = PTmenue.meilensteinTrendanalyse Then

                If awinSettings.englishLanguage Then
                    .Text = "Create Milestone Trendanalysis"
                    .OKButton.Text = "Show"
                    .headerLine.Text = "Milestones"
                Else
                    .Text = "Meilenstein Trendanalyse erzeugen"
                    .OKButton.Text = "Anzeigen"
                    .headerLine.Text = "Meilensteine"
                End If


                .statusLabel.Text = ""
                .statusLabel.Visible = True



                .picturePhasen.Visible = False
                .rdbPhases.Visible = False
                .rdbPhases.Checked = False
                .rdbPhases.Enabled = False

                .pictureMilestones.Visible = False
                .rdbMilestones.Visible = False
                .rdbMilestones.Checked = True
                .rdbMilestones.Enabled = False

                .pictureRoles.Visible = False
                .rdbRoles.Visible = False
                .rdbRoles.Checked = False
                .rdbRoles.Enabled = False

                .pictureCosts.Visible = False
                .rdbCosts.Visible = False
                .rdbCosts.Checked = False
                .rdbCosts.Enabled = False

                .rdbBU.Visible = False
                .pictureBU.Visible = False

                .rdbTyp.Visible = False
                .pictureTyp.Visible = False

                .einstellungen.Visible = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False

                .auswSpeichern.Visible = False


            End If

        End With

    End Sub


    ''' <summary>
    ''' Koordinaten merken 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmShowPlanElements_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.listselP, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.listselP, PTpinfo.left) = Me.Left

        awinSettings.isHryNameFrmActive = False

        ' Falls einzelne Projekte selektiert waren, so wird diese Selection hier aufgehoben
        If selectedProjekte.Count > 0 And visboZustaende.projectBoardMode = ptModus.graficboard Then
            Call awinDeSelect()
        End If


    End Sub

    ''' <summary>
    ''' wird zu Beginn, als "Lade-Routine" für das Formular aufgerufen; besetzt unter anderem die Selection Collections aus dem letzten Filter
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmShowPlanElements_Load(sender As Object, e As EventArgs) Handles Me.Load

        If frmCoord(PTfrm.listselP, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.listselP, PTpinfo.top))
            Me.Left = CInt(frmCoord(PTfrm.listselP, PTpinfo.left))
        Else
            Me.Top = 60
            Me.Left = 100
        End If


        ' hier kommt jetzt, welche Buttons sollen sichtbar sein ... 
        Call defineFrmButtonVisibility()

        awinSettings.isHryNameFrmActive = True

        statusLabel.Text = ""
        statusLabel.Visible = True


        ' jetzt werden anhand des letzten Filters die Collections gesetzt 

        Call retrieveSelections("Last", menuOption, selectedBUs, selectedTyps, _
                            selectedPhases, selectedMilestones, _
                            selectedRoles, selectedCosts)


        ' jetzt werden die ProjektReport- bzw. PortfolioReport-Vorlagen ausgelesen 
        ' in letztem Fall werden nur die mit Multiprojekt angezeigt 
        ' ur:27.07.2015: für menuOption = filterdefinieren, werden hier die in der Datenbank vorhandenen Filter zur Auswahl angezeigt

        Call frmHryNameReadPPTVorlagen(Me.menuOption, repVorlagenDropbox)


        ' die Filter einlesen
        Call frmHryNameReadFilterVorlagen(Me.menuOption, filterDropbox)

        ' alle definierten Filter in ComboBox anzeigen
        If Me.menuOption = PTmenue.filterdefinieren Or _
            Me.menuOption = PTmenue.sessionFilterDefinieren Then

            For Each kvp As KeyValuePair(Of String, clsFilter) In filterDefinitions.Liste
                filterDropbox.Items.Add(kvp.Key)
            Next
            Me.rdbPhases.Checked = True

        ElseIf Me.menuOption = PTmenue.meilensteinTrendanalyse Then
            Me.rdbMilestones.Checked = True

            For Each element As String In selectedMilestones
                If Not selNameListBox.Items.Contains(element) Then
                    selNameListBox.Items.Add(element)
                End If
            Next
        Else

            For Each kvp As KeyValuePair(Of String, clsFilter) In selFilterDefinitions.Liste
                filterDropbox.Items.Add(kvp.Key)
            Next
            If Me.rdbPhases.Visible Then
                Me.rdbPhases.Checked = True
            Else
                Me.rdbRoles.Checked = True
            End If

        End If



    End Sub

    ''' <summary>
    ''' Behandlung OK Button drücken
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim filterName As String = ""
        Dim lastFilter As String = "Last"

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        Dim formerEoU As Boolean = enableOnUpdate
        enableOnUpdate = False

        statusLabel.Text = ""

        ' hier muss jetzt noch der aktuelle rdb ausgelesen werden ..
        If Me.rdbPhases.Checked = True Then

            selectedPhases.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedPhases.Contains(element) Then
                    selectedPhases.Add(element, element)
                End If
            Next


        ElseIf Me.rdbMilestones.Checked = True Then

            selectedMilestones.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedMilestones.Contains(element) Then
                    selectedMilestones.Add(element, element)
                End If
            Next

        ElseIf rdbRoles.Checked = True Then

            selectedRoles.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedRoles.Contains(element) Then
                    selectedRoles.Add(element, element)
                End If
            Next

        ElseIf rdbCosts.Checked = True Then

            selectedCosts.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedCosts.Contains(element) Then
                    selectedCosts.Add(element, element)
                End If
            Next

        ElseIf rdbBU.Checked = True Then

            selectedBUs.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedBUs.Contains(element) Then
                    selectedBUs.Add(element, element)
                End If
            Next

        ElseIf rdbTyp.Checked = True Then

            selectedTyps.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedTyps.Contains(element) Then
                    selectedTyps.Add(element, element)
                End If
            Next
        End If

        If Me.menuOption = PTmenue.filterdefinieren Or _
            Me.menuOption = PTmenue.sessionFilterDefinieren Then

            If Not IsNothing(filterDropbox.Text) Then
                If filterDropbox.Text.Trim.Length > 0 Then
                    filterName = filterDropbox.Text.Trim
                    Call storeFilter(filterName, menuOption, selectedBUs, selectedTyps, _
                                                   selectedPhases, selectedMilestones, _
                                                   selectedRoles, selectedCosts, False)
                End If
            End If

        End If



        ' jetzt wird der letzte Filter gespeichert ..
        Call storeFilter(lastFilter, menuOption, selectedBUs, selectedTyps, _
                                                   selectedPhases, selectedMilestones, _
                                                   selectedRoles, selectedCosts, False)

        ''''
        ''
        ''
        ' jetzt kommt die Fall-Unterscheidung 
        ''
        ''
        ''''

        Dim validOption As Boolean
        If Me.menuOption = PTmenue.visualisieren Or Me.menuOption = PTmenue.einzelprojektReport Or _
            Me.menuOption = PTmenue.excelExport Or Me.menuOption = PTmenue.multiprojektReport Or _
            Me.menuOption = PTmenue.vorlageErstellen Then
            validOption = True
        ElseIf showRangeRight - showRangeLeft >= minColumns - 1 Then
            validOption = True
        Else
            validOption = False
        End If


        If Me.menuOption = PTmenue.multiprojektReport Or Me.menuOption = PTmenue.einzelprojektReport Then

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And validOption Then

                Dim vorlagenDateiName As String

                If Me.menuOption = PTmenue.multiprojektReport Then
                    vorlagenDateiName = awinPath & RepPortfolioVorOrdner & _
                                    "\" & repVorlagenDropbox.Text
                Else

                    vorlagenDateiName = awinPath & RepProjectVorOrdner & _
                                    "\" & repVorlagenDropbox.Text
                End If

                ' Prüfen, ob die Datei überhaupt existirt 
                If repVorlagenDropbox.Text.Length = 0 Then
                    Call MsgBox("bitte PPT Vorlage auswählen !")
                ElseIf My.Computer.FileSystem.FileExists(vorlagenDateiName) Then

                    Try
                        rdbMilestones.Enabled = False
                        rdbPhases.Enabled = False
                        rdbRoles.Enabled = False
                        rdbCosts.Enabled = False
                        filterBox.Enabled = False
                        nameListBox.Enabled = False
                        OKButton.Enabled = False
                        repVorlagenDropbox.Enabled = False
                        AbbrButton.Cursor = Cursors.Arrow

                        statusLabel.Text = ""
                        statusLabel.Visible = True

                        Me.Cursor = Cursors.WaitCursor
                        AbbrButton.Text = "Abbrechen"

                        'Call PPTstarten()

                        backgroundRunning = True

                        BackgroundWorker1.RunWorkerAsync(vorlagenDateiName)

                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                Else

                    Call MsgBox("bitte PPT Vorlage auswählen !")

                End If




            Else
                Call MsgBox("bitte mindestens ein Element selektieren bzw. " & vbLf & _
                             "einen Zeitraum angeben ...")
            End If

        Else

            ' die Aktion Subroutine aufrufen 
            '
            ' hier unterscheiden, was denn gewünscht wird; dann nur das übergeben 
            Dim tmpCollection As New Collection
            If rdbPhases.Checked Or rdbMilestones.Checked Then
                Call frmHryNameActions(Me.menuOption, selectedPhases, selectedMilestones, _
                            tmpCollection, tmpCollection, Me.chkbxOneChart.Checked, filterName)
            ElseIf rdbRoles.Checked Then
                Call frmHryNameActions(Me.menuOption, tmpCollection, tmpCollection, _
                            selectedRoles, tmpCollection, Me.chkbxOneChart.Checked, filterName)
            ElseIf rdbCosts.Checked Then
                Call frmHryNameActions(Me.menuOption, tmpCollection, tmpCollection, _
                            tmpCollection, selectedCosts, Me.chkbxOneChart.Checked, filterName)
            End If
            

        End If



        appInstance.EnableEvents = formerEE
        enableOnUpdate = formerEoU


        ' bei bestimmten Menu-Optionen das Formular dann schliessen 
        If Me.menuOption = PTmenue.excelExport Or _
            menuOption = PTmenue.filterdefinieren Or _
            menuOption = PTmenue.sessionFilterDefinieren Or _
            menuOption = PTmenue.leistbarkeitsAnalyse Or _
            (menuOption = PTmenue.meilensteinTrendanalyse And selectedMilestones.Count > 0) Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            MyBase.Close()
        Else
            ' geänderte Auswahl/Filterliste neu anzeigen
            filterDropbox.Items.Clear()
            For Each kvp As KeyValuePair(Of String, clsFilter) In selFilterDefinitions.Liste
                filterDropbox.Items.Add(kvp.Key)
            Next

        End If

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

    End Sub

    ''' <summary>
    ''' stellt ggf den vorherigen Zustand an vor-selektierten Items wieder her
    ''' ebenso den Searchkey 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbPhases_CheckedChanged(sender As Object, e As EventArgs) Handles rdbPhases.CheckedChanged

        'Dim i As Integer
        statusLabel.Text = ""
        filterBox.Enabled = True

        If Me.rdbPhases.Checked Then
            ' clear Listbox1 
            If awinSettings.englishLanguage Then
                headerLine.Text = "Phases"
            Else
                headerLine.Text = "Phasen"
            End If

            nameListBox.Items.Clear()
            selNameListBox.Items.Clear()
            filterBox.Text = ""


            If Me.menuOption = PTmenue.sessionFilterDefinieren Then
                ' immer die AlleProjekte hernehmen 
                If selectedProjekte.Count > 0 Then
                    allPhases = selectedProjekte.getPhaseNames
                ElseIf AlleProjekte.Count > 0 Then
                    allPhases = AlleProjekte.getPhaseNames
                Else
                    ' in der Session ist noch nichts, deswegen gbt es nichts zu definieren ... 
                    allPhases.Clear()
                End If

            ElseIf Me.menuOption = PTmenue.filterdefinieren Then
                ' 
                If selectedProjekte.Count > 0 Then
                    allPhases = selectedProjekte.getPhaseNames
                Else
                    ' eigentlich sollten hier alle Phasen der Datenbank stehen ... 
                    For i As Integer = 1 To PhaseDefinitions.Count
                        Dim tmpName As String = PhaseDefinitions.getPhaseDef(i).name
                        If Not allPhases.Contains(tmpName) Then
                            allPhases.Add(tmpName, tmpName)
                        End If
                    Next
                End If

            Else
                ' alle anderen Optionen
                If selectedProjekte.Count > 0 Then
                    allPhases = selectedProjekte.getPhaseNames
                ElseIf ShowProjekte.Count > 0 Then
                    allPhases = ShowProjekte.getPhaseNames
                Else
                    For i As Integer = 1 To PhaseDefinitions.Count
                        Dim tmpName As String = PhaseDefinitions.getPhaseDef(i).name
                        If Not allPhases.Contains(tmpName) Then
                            allPhases.Add(tmpName, tmpName)
                        End If
                    Next
                End If

            End If


            Call rebuildFormerState(PTauswahlTyp.phase)


        Else
            ' Merken, was ggf. das Filterkriterium war 
            'sKeyPhases = filterBox.Text

            ' Merken welches die selektierten Phasen waren 
            selectedPhases.Clear()
            For Each element As String In selNameListBox.Items
                selectedPhases.Add(element, element)
            Next

        End If

    End Sub

    Private Sub rdbMilestones_CheckedChanged(sender As Object, e As EventArgs) Handles rdbMilestones.CheckedChanged

        statusLabel.Text = ""
        filterBox.Enabled = True

        If Me.rdbMilestones.Checked Then
            ' clear Listbox1 
            If awinSettings.englishLanguage Then
                headerLine.Text = "Milestones"
            Else
                headerLine.Text = "Meilensteine"
            End If

            nameListBox.Items.Clear()
            selNameListBox.Items.Clear()

            filterBox.Text = ""

            If Me.menuOption = PTmenue.sessionFilterDefinieren Then
                ' immer die AlleProjekte hernehmen 
                If selectedProjekte.Count > 0 Then
                    allMilestones = selectedProjekte.getMilestoneNames
                ElseIf AlleProjekte.Count > 0 Then
                    allMilestones = AlleProjekte.getMilestoneNames
                Else
                    ' in der Session ist noch nichts, deswegen gbt es nichts zu definieren ... 
                    allMilestones.Clear()
                End If

            ElseIf Me.menuOption = PTmenue.filterdefinieren Then
                ' 
                If selectedProjekte.Count > 0 Then
                    allMilestones = selectedProjekte.getMilestoneNames
                Else
                    ' eigentlich sollten hier alle Meilensteine der Datenbank stehen ... 
                    For i As Integer = 1 To MilestoneDefinitions.Count
                        Dim tmpName As String = MilestoneDefinitions.getMilestoneDef(i).name
                        If Not allMilestones.Contains(tmpName) Then
                            allMilestones.Add(tmpName, tmpName)
                        End If
                    Next
                End If

            Else
                ' alle anderen Optionen
                If selectedProjekte.Count > 0 Then
                    allMilestones = selectedProjekte.getMilestoneNames
                ElseIf ShowProjekte.Count > 0 Then
                    allMilestones = ShowProjekte.getMilestoneNames
                Else
                    For i As Integer = 1 To MilestoneDefinitions.Count
                        Dim tmpName As String = MilestoneDefinitions.getMilestoneDef(i).name
                        If Not allMilestones.Contains(tmpName) Then
                            allMilestones.Add(tmpName, tmpName)
                        End If
                    Next
                End If

            End If



            Call rebuildFormerState(PTauswahlTyp.meilenstein)

        Else

            ' Merken welches die selektierten Meilensteine waren 
            selectedMilestones.Clear()
            For Each element As String In selNameListBox.Items
                selectedMilestones.Add(element, element)
            Next


        End If
    End Sub

    ''' <summary>
    ''' zeigt alle Rollen an, unabhängig davon ob sie in den Projekten vorkommen oder nicht 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbRoles_CheckedChanged(sender As Object, e As EventArgs) Handles rdbRoles.CheckedChanged


        statusLabel.Text = ""
        filterBox.Enabled = True

        If RoleDefinitions.Count = 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("no roles types defined! ")
            Else
                Call MsgBox("es sind keine Rollen definiert !")
            End If

        Else
            If Me.rdbRoles.Checked Then
                ' clear Listbox1 
                If awinSettings.englishLanguage Then
                    headerLine.Text = "Roles/Names"
                Else
                    headerLine.Text = "Rollen/Namen"
                End If

                nameListBox.Items.Clear()
                selNameListBox.Items.Clear()
                filterBox.Text = ""
                'chkbxOneChart.Text = "Alles in einem Chart"


                ' jetzt nur die Rollen anbieten, die auch vorkommen 
                If Me.menuOption = PTmenue.sessionFilterDefinieren Then
                    ' immer die AlleProjekte hernehmen 
                    If selectedProjekte.Count > 0 Then
                        allRoles = selectedProjekte.getRoleNames
                    ElseIf AlleProjekte.Count > 0 Then
                        allRoles = AlleProjekte.getRoleNames
                    Else
                        ' in der Session ist noch nichts, deswegen gbt es nichts zu definieren ... 
                        allRoles.Clear()
                    End If

                ElseIf Me.menuOption = PTmenue.filterdefinieren Then
                    ' 
                    If selectedProjekte.Count > 0 Then
                        allRoles = selectedProjekte.getRoleNames
                    Else
                        ' eigentlich sollten hier alle Rollen der Datenbank stehen ... 
                        For i As Integer = 1 To RoleDefinitions.Count
                            Dim tmpName As String = RoleDefinitions.getRoledef(i).name
                            If Not allRoles.Contains(tmpName) Then
                                allRoles.Add(tmpName, tmpName)
                            End If
                        Next
                    End If

                Else
                    ' alle anderen Optionen
                    If selectedProjekte.Count > 0 Then
                        allRoles = selectedProjekte.getRoleNames(True)
                    ElseIf ShowProjekte.Count > 0 Then
                        allRoles = ShowProjekte.getRoleNames(True)
                    Else
                        For i As Integer = 1 To RoleDefinitions.Count
                            Dim tmpName As String = RoleDefinitions.getRoledef(i).name
                            If Not allRoles.Contains(tmpName) Then
                                allRoles.Add(tmpName, tmpName)
                            End If
                        Next
                    End If

                End If



                Call rebuildFormerState(PTauswahlTyp.Rolle)

            Else
                ' Merken, was ggf. das Filterkriterium war 
                'sKeyRoles = filterBox.Text

                ' Merken welches die selektierten Phasen waren 
                selectedRoles.Clear()
                For Each element As String In selNameListBox.Items
                    selectedRoles.Add(element, element)
                Next

            End If
        End If

    End Sub

    ''' <summary>
    ''' wenn Radio-Button Kosten gedrückt wird 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbCosts_CheckedChanged(sender As Object, e As EventArgs) Handles rdbCosts.CheckedChanged

        statusLabel.Text = ""
        filterBox.Enabled = True

        If CostDefinitions.Count = 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("no cost types defined!")
            Else
                Call MsgBox("es sind keine Kostenarten definiert !")
            End If

        Else
            If Me.rdbCosts.Checked Then
                ' clear Listbox1 
                If awinSettings.englishLanguage Then
                    headerLine.Text = "Cost Types"
                Else
                    headerLine.Text = "Kostenarten"
                End If

                nameListBox.Items.Clear()
                selNameListBox.Items.Clear()
                filterBox.Text = ""
                'chkbxOneChart.Text = "Alles in einem Chart"

                ' jetzt nur die Kosten anbieten, die auch vorkommen 
                If Me.menuOption = PTmenue.sessionFilterDefinieren Then
                    ' immer die AlleProjekte hernehmen 
                    If selectedProjekte.Count > 0 Then
                        allCosts = selectedProjekte.getCostNames
                    ElseIf AlleProjekte.Count > 0 Then
                        allCosts = AlleProjekte.getCostNames()
                    Else
                        ' in der Session ist noch nichts, deswegen gbt es nichts zu definieren ... 
                        allCosts.Clear()
                    End If

                ElseIf Me.menuOption = PTmenue.filterdefinieren Then
                    ' 
                    If selectedProjekte.Count > 0 Then
                        allCosts = selectedProjekte.getCostNames
                    Else
                        ' eigentlich sollten hier alle Rollen der Datenbank stehen ... 
                        For i As Integer = 1 To CostDefinitions.Count - 1
                            Dim tmpName As String = CostDefinitions.getCostdef(i).name
                            If Not allCosts.Contains(tmpName) Then
                                allCosts.Add(tmpName, tmpName)
                            End If
                        Next
                    End If

                Else
                    ' alle anderen Optionen
                    If selectedProjekte.Count > 0 Then
                        allCosts = selectedProjekte.getCostNames
                    ElseIf ShowProjekte.Count > 0 Then
                        allCosts = ShowProjekte.getCostNames
                    Else
                        For i As Integer = 1 To CostDefinitions.Count - 1
                            Dim tmpName As String = CostDefinitions.getCostdef(i).name
                            If Not allCosts.Contains(tmpName) Then
                                allCosts.Add(tmpName, tmpName)
                            End If
                        Next
                    End If

                End If


                Call rebuildFormerState(PTauswahlTyp.Kostenart)

            Else

                ' Merken welches die selektierten Phasen waren 
                selectedCosts.Clear()
                'For Each element As String In ListBox1.SelectedItems
                For Each element As String In selNameListBox.Items
                    selectedCosts.Add(element, element)
                Next

            End If
        End If

    End Sub

    ''' <summary>
    ''' Behandlung Radio Button Business Unit drücken 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbBU_CheckedChanged(sender As Object, e As EventArgs) Handles rdbBU.CheckedChanged

        
        statusLabel.Text = ""
        filterBox.Enabled = True

        If businessUnitDefinitions.Count = 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("no Business Units defined!")
            Else
                Call MsgBox("es sind keine Business Units definiert !")
            End If

        Else
            If Me.rdbBU.Checked Then
                ' clear Listbox1 
                headerLine.Text = "Business Units"
                nameListBox.Items.Clear()
                selNameListBox.Items.Clear()
                filterBox.Text = ""

                ' jetzt nur die BUs anbieten, die auch vorkommen 
                If Me.menuOption = PTmenue.sessionFilterDefinieren Then
                    ' immer die AlleProjekte hernehmen 
                    If selectedProjekte.Count > 0 Then
                        allBUs = selectedProjekte.getBUNames
                    ElseIf AlleProjekte.Count > 0 Then
                        allBUs = AlleProjekte.getBUNames()
                    Else
                        ' in der Session ist noch nichts, deswegen gbt es nichts zu definieren ... 
                        allBUs.Clear()
                    End If

                ElseIf Me.menuOption = PTmenue.filterdefinieren Then
                    ' 
                    If selectedProjekte.Count > 0 Then
                        allBUs = selectedProjekte.getBUNames
                    Else
                        ' eigentlich sollten hier alle Rollen der Datenbank stehen ... 
                        For i As Integer = 1 To businessUnitDefinitions.Count
                            Dim tmpName As String = CStr(businessUnitDefinitions.ElementAt(i - 1).Value.name)
                            If Not allBUs.Contains(tmpName) Then
                                allBUs.Add(tmpName, tmpName)
                            End If
                        Next
                    End If

                Else
                    ' alle anderen Optionen
                    If selectedProjekte.Count > 0 Then
                        allBUs = selectedProjekte.getBUNames
                    ElseIf ShowProjekte.Count > 0 Then
                        allBUs = ShowProjekte.getBUNames
                    Else
                        For i As Integer = 1 To businessUnitDefinitions.Count
                            Dim tmpName As String = CStr(businessUnitDefinitions.ElementAt(i - 1).Value.name)
                            If Not allBUs.Contains(tmpName) Then
                                allBUs.Add(tmpName, tmpName)
                            End If
                        Next
                    End If

                End If

                Call rebuildFormerState(PTauswahlTyp.BusinessUnit)

            Else

                ' Merken welches die selektierten Phasen waren 
                selectedBUs.Clear()

                For Each element As String In selNameListBox.Items
                    selectedBUs.Add(element, element)
                Next

            End If
        End If


    End Sub

    Private Sub rdbTyp_CheckedChanged(sender As Object, e As EventArgs) Handles rdbTyp.CheckedChanged


        statusLabel.Text = ""
        filterBox.Enabled = True

        If Projektvorlagen.Count = 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("no project templates defined!")
            Else
                Call MsgBox("es sind keine Projektvorlagen definiert !")
            End If

        Else
            If Me.rdbTyp.Checked Then
                ' clear Listbox1 
                If awinSettings.englishLanguage Then
                    headerLine.Text = "Project-Templates"
                Else
                    headerLine.Text = "Projektvorlagen"
                End If

                nameListBox.Items.Clear()
                selNameListBox.Items.Clear()

                filterBox.Text = ""
                'chkbxOneChart.Text = "Alles in einem Chart"

                ' jetzt nur die BUs anbieten, die auch vorkommen 
                If Me.menuOption = PTmenue.sessionFilterDefinieren Then
                    ' immer die AlleProjekte hernehmen 
                    If selectedProjekte.Count > 0 Then
                        allTyps = selectedProjekte.getTypNames
                    ElseIf AlleProjekte.Count > 0 Then
                        allTyps = AlleProjekte.getTypNames
                    Else
                        ' in der Session ist noch nichts, deswegen gbt es nichts zu definieren ... 
                        allTyps.Clear()
                    End If

                ElseIf Me.menuOption = PTmenue.filterdefinieren Then
                    ' 
                    If selectedProjekte.Count > 0 Then
                        allTyps = selectedProjekte.getTypNames
                    Else
                        ' eigentlich sollten hier alle Rollen der Datenbank stehen ... 
                        For i As Integer = 1 To Projektvorlagen.Count
                            Dim tmpName As String = CStr(Projektvorlagen.Liste.ElementAt(i - 1).Key)
                            If Not allTyps.Contains(tmpName) Then
                                allTyps.Add(tmpName, tmpName)
                            End If
                        Next
                    End If

                Else
                    ' alle anderen Optionen
                    If selectedProjekte.Count > 0 Then
                        allTyps = selectedProjekte.getTypNames
                    ElseIf ShowProjekte.Count > 0 Then
                        allTyps = ShowProjekte.getTypNames
                    Else
                        For i As Integer = 1 To Projektvorlagen.Count
                            Dim tmpName As String = CStr(Projektvorlagen.Liste.ElementAt(i - 1).Key)
                            If Not allTyps.Contains(tmpName) Then
                                allTyps.Add(tmpName, tmpName)
                            End If
                        Next
                    End If

                End If


                Call rebuildFormerState(PTauswahlTyp.ProjektTyp)

            Else

                ' Merken welches die selektierten Phasen waren 
                selectedTyps.Clear()

                For Each element As String In selNameListBox.Items
                    selectedTyps.Add(element, element)
                Next

            End If
        End If


    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click


        If (menuOption = PTmenue.multiprojektReport Or menuOption = PTmenue.einzelprojektReport) _
            And backgroundRunning Then

            rdbMilestones.Enabled = True
            rdbPhases.Enabled = True
            rdbRoles.Enabled = True
            rdbCosts.Enabled = True
            filterBox.Enabled = True
            nameListBox.Enabled = True
            OKButton.Enabled = True
            If awinSettings.englishLanguage Then
                AbbrButton.Text = "Reset Selection"
                statusLabel.Text = "Report creation cancelled"
            Else
                AbbrButton.Text = "Zurücksetzen"
                statusLabel.Text = "Berichterstellung wurde beendet"
            End If

            repVorlagenDropbox.Enabled = True


            Me.Cursor = Cursors.Arrow
            backgroundRunning = False

            Try
                Me.BackgroundWorker1.CancelAsync()
            Catch ex As Exception
                backgroundRunning = True

            End Try


        Else
            nameListBox.SelectedItems.Clear()
            selNameListBox.Items.Clear()
            filterBox.Text = ""

            If rdbPhases.Checked Then
                selectedPhases.Clear()
            ElseIf rdbMilestones.Checked Then
                selectedMilestones.Clear()
            ElseIf rdbRoles.Checked Then
                selectedRoles.Clear()
            ElseIf rdbCosts.Checked Then
                selectedCosts.Clear()
            ElseIf rdbBU.Checked Then
                selectedBUs.Clear()
            Else
                selectedTyps.Clear()
            End If
        End If


        'MyBase.Close()

    End Sub


    ''' <summary>
    ''' wenn etwas in der Such-Maske eingegeben wird: prüfen, Listbox1 entsprechend ausdünnen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub filterBox_TextChanged(sender As Object, e As EventArgs) Handles filterBox.TextChanged

        Dim suchstr As String = filterBox.Text
        Dim currentNames As New Collection

        If rdbPhases.Checked Then
            currentNames = allPhases
        ElseIf rdbMilestones.Checked Then
            currentNames = allMilestones
        ElseIf rdbRoles.Checked Then
            currentNames = allRoles
        ElseIf rdbCosts.Checked Then
            currentNames = allCosts
        ElseIf rdbBU.Checked Then
            currentNames = allBUs
        ElseIf rdbTyp.Checked Then
            currentNames = allTyps
        End If


        If filterBox.Text = "" Then
            nameListBox.Items.Clear()
            For Each s As String In currentNames
                nameListBox.Items.Add(s)
            Next
        Else
            nameListBox.Items.Clear()
            For Each s As String In currentNames
                If s.Contains(suchstr) Then
                    nameListBox.Items.Add(s)
                End If
            Next
        End If

    End Sub


    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub pictureRoles_Click(sender As Object, e As EventArgs) Handles pictureRoles.Click
        If Me.rdbRoles.Checked = False Then
            rdbRoles.Checked = True
        Else
            rdbRoles.Checked = False
        End If
    End Sub

    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub picturePhasen_Click(sender As Object, e As EventArgs) Handles picturePhasen.Click
        If Me.rdbPhases.Checked = False Then
            rdbPhases.Checked = True
        Else
            rdbPhases.Checked = False
        End If
    End Sub

    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub pictureMilestones_Click(sender As Object, e As EventArgs) Handles pictureMilestones.Click
        If Me.rdbMilestones.Checked = False Then
            rdbMilestones.Checked = True
        Else
            rdbMilestones.Checked = False
        End If
    End Sub

    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub pictureCosts_Click(sender As Object, e As EventArgs) Handles pictureCosts.Click
        If Me.rdbCosts.Checked = False Then
            Me.rdbCosts.Checked = True
        Else
            Me.rdbCosts.Checked = False
        End If
    End Sub

    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub pictureBU_Click(sender As Object, e As EventArgs) Handles pictureBU.Click

        If Me.rdbBU.Checked = False Then
            Me.rdbBU.Checked = True
        Else
            Me.rdbBU.Checked = False
        End If

    End Sub

    ''' <summary>
    ''' stellt den vorherigen Zustand wieder her: welche Werte waren bereits für die betreffende 
    ''' Kategorie ausgewählt
    ''' dabei wird auf die in dieser Klasse definierten Variablen selectedphases, allphases, ... zugegriffen 
    ''' </summary>
    ''' <param name="typ"></param>
    ''' <remarks></remarks>
    Private Sub rebuildFormerState(ByVal typ As Integer)

        'Dim searchkey As String = ""
        Dim tmpCollection As New Collection
        Dim i As Integer
        Dim listOfNames As New Collection

        Select Case typ
            Case PTauswahlTyp.phase
                'searchkey = sKeyPhases
                tmpCollection = selectedPhases
                listOfNames = allPhases

            Case PTauswahlTyp.meilenstein
                'searchkey = sKeyMilestones
                tmpCollection = selectedMilestones
                listOfNames = allMilestones

            Case PTauswahlTyp.Rolle
                'searchkey = sKeyRoles
                tmpCollection = selectedRoles
                listOfNames = allRoles

            Case PTauswahlTyp.Kostenart
                'searchkey = sKeyCosts
                tmpCollection = selectedCosts
                listOfNames = allCosts

            Case PTauswahlTyp.BusinessUnit
                tmpCollection = selectedBUs
                listOfNames = allBUs

            Case PTauswahlTyp.ProjektTyp
                tmpCollection = selectedTyps
                listOfNames = allTyps

        End Select


        For i = 1 To listOfNames.Count
            If Not nameListBox.Items.Contains(listOfNames.Item(i)) Then
                nameListBox.Items.Add(listOfNames.Item(i))
            End If
        Next


        ' Filter Box Test setzen 
        filterBox.Text = ""

        ' jetzt prüfen, ob selected... bereits etwas enthält
        ' wenn ja, dann werden diese Items in Listbox2 dargestellt 
        selNameListBox.Items.Clear()
        For Each element As String In tmpCollection
            If Not selNameListBox.Items.Contains(element) Then
                selNameListBox.Items.Add(element)
            End If

        Next

    End Sub

    Private Sub BackgroundWorker1_Disposed(sender As Object, e As EventArgs) Handles BackgroundWorker1.Disposed



    End Sub


    ''' <summary>
    ''' Hintergrund Prozess - wird nur für die Report Erzeugung benötigt 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim vorlagenDateiName As String = CType(e.Argument, String)

        currentReportProfil.name = "Last"
        currentReportProfil.Phases = copyColltoSortedList(selectedPhases)
        currentReportProfil.Milestones = copyColltoSortedList(selectedMilestones)
        currentReportProfil.Roles = copyColltoSortedList(selectedRoles)
        currentReportProfil.Costs = copyColltoSortedList(selectedCosts)
        currentReportProfil.Typs = copyColltoSortedList(selectedTyps)
        currentReportProfil.BUs = copyColltoSortedList(selectedBUs)

        currentReportProfil.CalendarVonDate = StartofCalendar

        Try
            With awinSettings

                If .mppSortiertDauer Then
                    .mppShowAllIfOne = True
                End If

                currentReportProfil.ProjectLine = .mppShowProjectLine
                currentReportProfil.AllIfOne = .mppShowAllIfOne
                currentReportProfil.Ampeln = .mppShowAmpel
                currentReportProfil.UseAbbreviation = .mppUseAbbreviation

                currentReportProfil.PhName = .mppShowPhName
                currentReportProfil.PhDate = .mppShowPhDate
                currentReportProfil.MSName = .mppShowMsName
                currentReportProfil.MSDate = .mppShowMsDate
                currentReportProfil.UseAbbreviation = .mppUseAbbreviation
                currentReportProfil.KwInMilestone = .mppKwInMilestone


                currentReportProfil.VLinien = .mppVertikalesRaster
                currentReportProfil.ShowHorizontals = .mppShowHorizontals
                currentReportProfil.Legend = .mppShowLegend
                currentReportProfil.OnePage = .mppOnePage

                currentReportProfil.SortedDauer = .mppSortiertDauer
                currentReportProfil.ExtendedMode = .mppExtendedMode
                currentReportProfil.FullyContained = .mppFullyContained

                currentReportProfil.projectsWithNoMPmayPass = .mppProjectsWithNoMPmayPass

                ' Dateiname eliminieren, ohne Pfadangaben im ReportProfil speichern
                Dim hstr() As String
                hstr = Split(vorlagenDateiName, "\")
                currentReportProfil.PPTTemplate = hstr(hstr.Length - 1)

                If vorlagenDateiName.Contains(RepPortfolioVorOrdner) Then

                    ' Multiprojekt-Bericht
                    currentReportProfil.isMpp = True

                    ' für Multiprojekt-Report muss ein Time-Range angegeben sein
                    Dim vonDate As Date = getDateofColumn(showRangeLeft, False)
                    Dim bisDate As Date = getDateofColumn(showRangeRight, True)
                    Try
                        currentReportProfil.calcRepVonBis(vonDate, bisDate)
                    Catch ex As Exception
                        Throw New ArgumentException(ex.Message)
                    End Try


                    Call createPPTSlidesFromConstellation(vorlagenDateiName, _
                                                      selectedPhases, selectedMilestones, _
                                                      selectedRoles, selectedCosts, _
                                                      selectedBUs, selectedTyps, True, _
                                                      worker, e)
                Else
                    ' Einzelprojekt-Bericht

                    currentReportProfil.isMpp = False

                    ' für Einzelprojekt-Bericht ist kein Time-Range erforderlich => keine Fehlermeldung
                    Try
                        currentReportProfil.calcRepVonBis(StartofCalendar, StartofCalendar)
                    Catch ex As Exception

                    End Try


                    Call createPPTReportFromProjects(vorlagenDateiName, _
                                                     selectedPhases, selectedMilestones, _
                                                     selectedRoles, selectedCosts, _
                                                     selectedBUs, selectedTyps, _
                                                     worker, e)
                End If


            End With
        Catch ex As Exception
            Call MsgBox("Fehler " & ex.Message)
        End Try



    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusLabel.Text = CType(re.Result, String)

    End Sub

    ''' <summary>
    ''' wird durchlaufen, wenn der Hintergrund Prozess mit dem Erstellen der Multiprojektsicht fertig ist 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Me.statusLabel.Text = "...done"
        Me.rdbMilestones.Enabled = True
        Me.rdbPhases.Enabled = True
        Me.rdbRoles.Enabled = True
        Me.rdbCosts.Enabled = True
        Me.filterBox.Enabled = True
        Me.nameListBox.Enabled = True
        Me.OKButton.Enabled = True
        Me.repVorlagenDropbox.Enabled = True
        Me.Cursor = Cursors.Arrow
        Me.statusLabel.Visible = True
        Me.AbbrButton.Text = "Zurücksetzen"


        backgroundRunning = False
    End Sub


    ''' <summary>
    ''' ruft das Formular auf, um die Einstellungen für das Multireporting vorzunehmen  
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub einstellungen_Click(sender As Object, e As EventArgs) Handles einstellungen.Click

        Dim mppFrm As New frmMppSettings
        Dim dialogreturn As DialogResult

        mppFrm.calledfrom = "frmShowPlanElements"
        dialogreturn = mppFrm.ShowDialog


    End Sub

    ''' <summary>
    ''' fügt das selektierte Element der Listbox2 hinzu
    ''' es muss unterschieden werden: 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub addButton_Click(sender As Object, e As EventArgs) Handles addButton.Click

        Dim i As Integer
        Dim element As Object
        Dim sammelCollection As New Collection


        For i = 1 To nameListBox.SelectedItems.Count
            element = nameListBox.SelectedItems.Item(i - 1)
            If selNameListBox.Items.Contains(element) Then
                ' nichts tun 
            Else
                selNameListBox.Items.Add(element)
            End If
        Next

        nameListBox.SelectedItems.Clear()


        ' ur: 30.07.2015: gilt nicht für filterdefinieren: Konsistenzbedingungen einhalten: 

        If Me.menuOption = PTmenue.visualisieren Then

            If (rdbPhases.Checked = True Or rdbMilestones.Checked = True) And selNameListBox.Items.Count > 0 Then
                selectedCosts.Clear()
                selectedRoles.Clear()
            ElseIf rdbRoles.Checked = True Then
                selectedCosts.Clear()
                selectedMilestones.Clear()
                selectedPhases.Clear()
            ElseIf rdbCosts.Checked = True Then
                selectedRoles.Clear()
                selectedMilestones.Clear()
                selectedPhases.Clear()
            End If

        ElseIf Me.menuOption = PTmenue.meilensteinTrendanalyse Then

            selectedRoles.Clear()
            selectedCosts.Clear()
            selectedPhases.Clear()

        End If



    End Sub

    ''' <summary>
    ''' entfernt ein Item aus der Listbox2 - die ausgewöhlten Elemente 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub removeButton_Click(sender As Object, e As EventArgs) Handles removeButton.Click
        Dim i As Integer
        Dim element As Object
        Dim removeCollection As New Collection



        For i = 1 To selNameListBox.SelectedItems.Count
            element = selNameListBox.SelectedItems.Item(i - 1)
            removeCollection.Add(element)
        Next

        For Each element In removeCollection
            selNameListBox.Items.Remove(element)
        Next


    End Sub
    '''' ur: 3.08.2015: wurde mit der Auswahl aus der Hierarchie frmHierarchySelection ersetzt

    ' '' ''Private Sub selNameListBox_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles selNameListBox.MouseDoubleClick

    ' '' ''    Dim elemName As String = ""
    ' '' ''    Dim childName As String
    ' '' ''    Dim sammelCollection As Collection
    ' '' ''    Dim anzahl As Integer = selNameListBox.SelectedItems.Count

    ' '' ''    If rdbPhases.Checked Then

    ' '' ''        If anzahl = 1 Then
    ' '' ''            elemName = CStr(selNameListBox.SelectedItem)
    ' '' ''        ElseIf anzahl > 1 Then
    ' '' ''            elemName = CStr(selNameListBox.SelectedItems.Item(1))
    ' '' ''        End If

    ' '' ''        ' das Element rausnehmen 
    ' '' ''        selNameListBox.Items.Remove(selNameListBox.SelectedItem)

    ' '' ''        sammelCollection = ShowProjekte.getPhasesOfPhase(elemName)

    ' '' ''        For i As Integer = 1 To sammelCollection.Count

    ' '' ''            childName = CStr(sammelCollection.Item(i))
    ' '' ''            If Not selNameListBox.Items.Contains(childName) Then
    ' '' ''                selNameListBox.Items.Add(childName)
    ' '' ''            End If

    ' '' ''        Next

    ' '' ''        If sammelCollection.Count > 0 Then
    ' '' ''            ' dann wurde eine Ersetzung vorgenommen 
    ' '' ''            ' jetzt muss bestimmt werden, ob Farben geändert werden müssen 

    ' '' ''        End If

    ' '' ''    End If


    ' '' ''End Sub

    Private Sub filterDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles filterDropbox.SelectedIndexChanged

        If Me.menuOption = PTmenue.filterdefinieren Then

            Dim fName As String = filterDropbox.SelectedItem.ToString
            ' wird nicht benötigt: ur: 29.07.2015 Dim filter As clsFilter = filterDefinitions.retrieveFilter(fName)

            ' jetzt werden anhand des Filters "fName" die Collections gesetzt 
            Call retrieveSelections(fName, menuOption, selectedBUs, selectedTyps, _
                                    selectedPhases, selectedMilestones, _
                                    selectedRoles, selectedCosts)
            If Me.rdbPhases.Checked Then
                Call rebuildFormerState(PTauswahlTyp.phase)
            ElseIf Me.rdbMilestones.Checked Then
                Call rebuildFormerState(PTauswahlTyp.meilenstein)
            ElseIf Me.rdbRoles.Checked Then
                Call rebuildFormerState(PTauswahlTyp.Rolle)
            ElseIf Me.rdbCosts.Checked Then
                Call rebuildFormerState(PTauswahlTyp.Kostenart)
            ElseIf Me.rdbTyp.Checked Then
                Call rebuildFormerState(PTauswahlTyp.ProjektTyp)
            ElseIf Me.rdbBU.Checked Then
                Call rebuildFormerState(PTauswahlTyp.BusinessUnit)
            End If

            '  Call MsgBox("in filterDropbox")
        Else

            Dim fName As String = filterDropbox.SelectedItem.ToString
            ' wird nicht benötigt: ur: 29.07.2015 Dim filter As clsFilter = filterDefinitions.retrieveFilter(fName)

            ' jetzt werden anhand des Filters "fName" die Collections gesetzt 
            Call retrieveSelections(fName, menuOption, selectedBUs, selectedTyps, _
                                    selectedPhases, selectedMilestones, _
                                    selectedRoles, selectedCosts)
            If Me.rdbPhases.Checked Then
                Call rebuildFormerState(PTauswahlTyp.phase)
            ElseIf Me.rdbMilestones.Checked Then
                Call rebuildFormerState(PTauswahlTyp.meilenstein)
            ElseIf Me.rdbRoles.Checked Then
                Call rebuildFormerState(PTauswahlTyp.Rolle)
            ElseIf Me.rdbCosts.Checked Then
                Call rebuildFormerState(PTauswahlTyp.Kostenart)
            ElseIf Me.rdbTyp.Checked Then
                Call rebuildFormerState(PTauswahlTyp.ProjektTyp)
            ElseIf Me.rdbBU.Checked Then
                Call rebuildFormerState(PTauswahlTyp.BusinessUnit)
            End If

        End If

    End Sub

    Private Sub auswSpeichern_Click(sender As Object, e As EventArgs) Handles auswSpeichern.Click

        Dim filterName As String = ""
        Dim lastFilter As String = "Last"
        appInstance.EnableEvents = False
        enableOnUpdate = False

        statusLabel.Text = ""

        ' hier muss jetzt noch der aktuelle rdb ausgelesen werden ..
        If Me.rdbPhases.Checked = True Then

            selectedPhases.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedPhases.Contains(element) Then
                    selectedPhases.Add(element, element)
                End If
            Next


        ElseIf Me.rdbMilestones.Checked = True Then

            selectedMilestones.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedMilestones.Contains(element) Then
                    selectedMilestones.Add(element, element)
                End If
            Next

        ElseIf rdbRoles.Checked = True Then

            selectedRoles.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedRoles.Contains(element) Then
                    selectedRoles.Add(element, element)
                End If
            Next

        ElseIf rdbCosts.Checked = True Then

            selectedCosts.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedCosts.Contains(element) Then
                    selectedCosts.Add(element, element)
                End If
            Next

        ElseIf rdbBU.Checked = True Then

            selectedBUs.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedBUs.Contains(element) Then
                    selectedBUs.Add(element, element)
                End If
            Next

        ElseIf rdbTyp.Checked = True Then

            selectedTyps.Clear()
            For Each element As String In selNameListBox.Items
                If Not selectedTyps.Contains(element) Then
                    selectedTyps.Add(element, element)
                End If
            Next
        End If

        If Me.menuOption = PTmenue.filterdefinieren Or _
            Me.menuOption = PTmenue.sessionFilterDefinieren Then

            filterName = filterDropbox.Text
            ' jetzt wird der Filter unter dem Namen filterName gespeichert ..
            Call storeFilter(filterName, menuOption, selectedBUs, selectedTyps, _
                                                   selectedPhases, selectedMilestones, _
                                                   selectedRoles, selectedCosts, False)

        ElseIf Me.menuOption = PTmenue.visualisieren Then

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0) And _
                (selectedRoles.Count > 0 Or selectedCosts.Count > 0) Then
                Call MsgBox("es können nur entweder Phasen / Meilensteine oder Rollen oder Kosten angezeigt werden")
            Else
                filterName = filterDropbox.Text
                ' jetzt wird der Filter unter dem Namen filterName gespeichert ..
                Call storeFilter(filterName, menuOption, selectedBUs, selectedTyps, _
                                                       selectedPhases, selectedMilestones, _
                                                       selectedRoles, selectedCosts, False)
            End If

        Else    ' alle anderen PTmenues

            filterName = filterDropbox.Text
            ' jetzt wird der Filter unter dem Namen filterName gespeichert ..
            Call storeFilter(filterName, menuOption, selectedBUs, selectedTyps, _
                                                   selectedPhases, selectedMilestones, _
                                                   selectedRoles, selectedCosts, False)
        End If



        ' jetzt wird der letzte Filter gespeichert ..
        Call storeFilter(lastFilter, menuOption, selectedBUs, selectedTyps, _
                                                   selectedPhases, selectedMilestones, _
                                                   selectedRoles, selectedCosts, False)

        ' geänderte Auswahl/Filterliste neu anzeigen
        If Not (Me.menuOption = PTmenue.filterdefinieren) Then
            filterDropbox.Items.Clear()
            For Each kvp As KeyValuePair(Of String, clsFilter) In selFilterDefinitions.Liste
                filterDropbox.Items.Add(kvp.Key)
            Next

        End If


    End Sub

    Private Sub nameListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles nameListBox.SelectedIndexChanged

    End Sub
End Class