Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports System.Windows.Forms
Imports System.ComponentModel


Public Class frmHierarchySelection


    Private hry As clsHierarchy
    Public repProfil As clsReportAll



    Private allMilestones As New Collection
    Private allPhases As New Collection
    Private allCosts As New Collection
    Private allRoles As New Collection
    Private allBUs As New Collection
    Private allTyps As New Collection


    Private auswahl As Integer = 0
    Private lastAuswahl As Integer = -1
    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection
    Private selectedCosts As New Collection
    Private selectedRoles As New Collection
    Private selectedBUs As New Collection
    Private selectedTyps As New Collection

    Private Enum PTauswahlTyp
        phase = 0
        meilenstein = 1
        Rolle = 2
        Kostenart = 3
        BusinessUnit = 4
        ProjektTyp = 5
    End Enum

    ' hier steht drin, mit welcher Menue-Option das Ganze aufgerufen wurde 
    Friend menuOption As Integer

    ' hier steht ggf die ButtonID drin
    Friend ribbonButtonID As String = ""

    ' an der aufrufenden Stelle muss hier entweder "Multiprojekt-Tafel" oder
    ' "MS Project" stehen. 
    Friend calledFrom As String


    Private Sub defineFrmButtonVisibility()

        If awinSettings.englishLanguage Then
            hryStufenLabel.Text = "nr of parents to be considered"
            chkbxOneChart.Text = "all in one chart"
            statusLabel.Text = ""
            einstellungen.Text = "Settings"
            labelPPTVorlage.Text = "Powerpoint Template"
            AbbrButton.Text = "Reset Selection"
        End If

        With Me

            ' Änderung tk: die Hierarchie soll, wie bisher nur bei BHTC nie sichtbar sein; 
            ' der Default Value auf 50 
            '' '' 
            ' ''.hryStufenLabel.Visible = True
            ' ''.hryStufen.Visible = True
            ' ''.hryStufen.Value = 0

            '  Änderung ur:
            ' frmHierarchySelection gemischt mit frmnameSelection

            .hryStufenLabel.Visible = False
            .hryStufen.Visible = False
            .hryStufen.Value = 50

            filterBox.Enabled = False
            filterBox.Visible = False


            If .menuOption = PTmenue.filterdefinieren Then

                If awinSettings.englishLanguage Then
                    .Text = "define Database Filter"
                    .OKButton.Text = "Store"
                    .filterLabel.Text = "Name of Filter"
                Else
                    .Text = "Datenbank Filter definieren"
                    .OKButton.Text = "Speichern"
                    .filterLabel.Text = "Name des Filters"
                End If
                '.rdbNameList.Enabled = False
                '.rdbNameList.Visible = False
                '.rdbNameList.Checked = False

                .rdbNameList.Enabled = True
                .rdbNameList.Visible = True
                .rdbNameList.Checked = True


                '.rdbProjStruktProj.Enabled = False
                '.rdbProjStruktProj.Visible = False
                '.rdbProjStruktProj.Checked = False

                '.rdbProjStruktTyp.Enabled = False
                '.rdbProjStruktTyp.Visible = False
                '.rdbProjStruktTyp.Checked = False

                .rdbProjStruktProj.Enabled = True
                .rdbProjStruktProj.Visible = True
                .rdbProjStruktProj.Checked = False

                .rdbProjStruktTyp.Enabled = True
                .rdbProjStruktTyp.Visible = True
                .rdbProjStruktTyp.Checked = False

                '.rdbPhases.Visible = False
                '.rdbPhases.Checked = False
                '.picturePhasen.Visible = False
                .rdbPhases.Visible = True
                .rdbPhases.Checked = True
                .picturePhasen.Visible = True

                '.rdbPhaseMilest.Visible = True
                .rdbPhaseMilest.Visible = False
                .rdbPhaseMilest.Checked = False

                '.picturePhaseMilest.Visible = True
                .picturePhaseMilest.Visible = False

                .rdbMilestones.Visible = True
                .rdbMilestones.Checked = False
                .pictureMilestones.Visible = True

                .rdbRoles.Visible = True
                .rdbRoles.Checked = False
                .pictureRoles.Visible = True

                .rdbCosts.Visible = True
                .rdbCosts.Checked = False
                .pictureCosts.Visible = True

                .rdbBU.Visible = True
                .pictureBU.Visible = True

                .rdbTyp.Visible = True
                .pictureTyp.Visible = True

                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .AbbrButton.Visible = False
                .AbbrButton.Enabled = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports 
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False
                .einstellungen.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True

                ' Auswahl Speichern
                .auswSpeichern.Visible = False
                .auswSpeichern.Enabled = False

                ' Auswahl Laden
                .auswLaden.Visible = False
                .auswLaden.Enabled = False

                .einstellungen.Visible = False

            ElseIf .menuOption = PTmenue.visualisieren Then

                If awinSettings.englishLanguage Then
                    .Text = "Visualize Phases & Milestones"
                    .OKButton.Text = "Visualize"
                    .filterLabel.Text = "Selection"
                    .auswSpeichern.Text = "Store"
                    .auswLaden.Text = "Load"
                    .AbbrButton.Text = "Cancel"
                    .chkbxOneChart.Text = "all in one chart"
                    .rdbNameList.Text = "List"
                    .rdbProjStruktProj.Text = "Project-Structure by Project"
                    .rdbProjStruktTyp.Text = "Project-Structure by Type"
                Else
                    .Text = "Phasen / Meilensteine visualisieren"
                    .OKButton.Text = "Anzeigen"
                    .filterLabel.Text = "Auswahl"
                    .auswSpeichern.Text = "Speichern"
                    .auswLaden.Text = "Laden"
                    .AbbrButton.Text = "Abbrechen"
                    .chkbxOneChart.Text = "alles in einem Chart"
                    .rdbNameList.Text = "Liste"
                    .rdbProjStruktProj.Text = "Projekt-Struktur (Projekt)"
                    .rdbProjStruktTyp.Text = "Projekt-Struktur (Typ)"
                End If

                .rdbNameList.Enabled = True
                .rdbNameList.Visible = True
                .rdbNameList.Checked = True

                .rdbProjStruktProj.Enabled = True
                .rdbProjStruktProj.Visible = True
                .rdbProjStruktProj.Checked = False

                .rdbProjStruktTyp.Enabled = True
                .rdbProjStruktTyp.Visible = True
                .rdbProjStruktTyp.Checked = False

                .rdbPhases.Visible = True
                .rdbPhases.Checked = True
                .picturePhasen.Visible = True

                .rdbMilestones.Visible = True
                .rdbMilestones.Checked = False
                .pictureMilestones.Visible = True

                .rdbBU.Visible = False
                .pictureBU.Visible = False

                .rdbTyp.Visible = False
                .pictureTyp.Visible = False

                .rdbRoles.Visible = False
                .pictureRoles.Visible = False

                .rdbCosts.Visible = False
                .pictureCosts.Visible = False

                .rdbPhaseMilest.Visible = False
                .picturePhaseMilest.Visible = False

                .AbbrButton.Visible = False
                .AbbrButton.Enabled = False
                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False
                .einstellungen.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True



            ElseIf .menuOption = PTmenue.leistbarkeitsAnalyse Then


                If awinSettings.englishLanguage Then
                    .Text = "Create Feasibility Charts"
                    .filterLabel.Text = "Selection"
                    .OKButton.Text = "Create Charts"
                Else
                    .Text = "Leistbarkeits-Charts erstellen"
                    .filterLabel.Text = "Auswahl"
                    .OKButton.Text = "Charts erstellen"
                End If

                
                .rdbNameList.Enabled = False
                .rdbNameList.Visible = False
                .rdbNameList.Checked = False

                '.rdbNameList.Enabled = True
                '.rdbNameList.Visible = True
                '.rdbNameList.Checked = True


                .rdbProjStruktProj.Enabled = False
                .rdbProjStruktProj.Visible = False
                .rdbProjStruktProj.Checked = False

                .rdbProjStruktTyp.Enabled = False
                .rdbProjStruktTyp.Visible = False
                .rdbProjStruktTyp.Checked = False

                '.rdbProjStruktProj.Enabled = True
                '.rdbProjStruktProj.Visible = True
                '.rdbProjStruktProj.Checked = False

                '.rdbProjStruktTyp.Enabled = True
                '.rdbProjStruktTyp.Visible = True
                '.rdbProjStruktTyp.Checked = False

                .rdbPhases.Visible = False
                .rdbPhases.Checked = False
                .picturePhasen.Visible = False
                '.rdbPhases.Visible = True
                '.rdbPhases.Checked = True
                '.picturePhasen.Visible = True

                .rdbPhaseMilest.Visible = True
                '.rdbPhaseMilest.Visible = False
                .rdbPhaseMilest.Checked = False

                .picturePhaseMilest.Visible = True
                '.picturePhaseMilest.Visible = False

                .rdbMilestones.Visible = True
                .rdbMilestones.Checked = False
                .pictureMilestones.Visible = True

                .rdbBU.Visible = False
                .pictureBU.Visible = False

                .rdbTyp.Visible = False
                .pictureTyp.Visible = False

                .rdbRoles.Visible = True
                .rdbRoles.Checked = True
                .pictureRoles.Visible = True

                .rdbCosts.Visible = True
                .rdbCosts.Checked = False
                .pictureCosts.Visible = True

                .AbbrButton.Visible = False
                .AbbrButton.Enabled = False
                .statusLabel.Text = ""
                .statusLabel.Visible = True


                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = True

                ' Reports
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False
                .einstellungen.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True


            ElseIf .menuOption = PTmenue.einzelprojektReport Then

                If awinSettings.englishLanguage Then
                    .Text = "Create Project-/Variant Report"
                    .OKButton.Text = "Create Report"
                    .filterLabel.Text = "Selection"
                Else
                    .Text = "Projekt-Varianten Report erzeugen"
                    .OKButton.Text = "Bericht erstellen"
                    .filterLabel.Text = "Auswahl"
                End If

                .rdbNameList.Enabled = True
                .rdbNameList.Visible = True
                .rdbNameList.Checked = True

                .rdbProjStruktProj.Enabled = True
                .rdbProjStruktProj.Visible = True
                .rdbProjStruktProj.Checked = False

                .rdbProjStruktTyp.Enabled = True
                .rdbProjStruktTyp.Visible = True
                .rdbProjStruktTyp.Checked = False

                .rdbPhases.Visible = True
                .rdbPhases.Checked = True
                .picturePhasen.Visible = True

                .rdbMilestones.Visible = True
                .rdbMilestones.Checked = False
                .pictureMilestones.Visible = True

                .rdbBU.Visible = False
                .pictureBU.Visible = False

                .rdbTyp.Visible = False
                .pictureTyp.Visible = False

                .rdbRoles.Visible = False
                .pictureRoles.Visible = False

                .rdbCosts.Visible = False
                .pictureCosts.Visible = False

                .rdbPhaseMilest.Visible = False
                .picturePhaseMilest.Visible = False

                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .AbbrButton.Visible = False
                .AbbrButton.Enabled = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False


                ' Reports
                .repVorlagenDropbox.Visible = True
                .labelPPTVorlage.Visible = True
                .einstellungen.Visible = True

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True

            ElseIf .menuOption = PTmenue.multiprojektReport Then

                If awinSettings.englishLanguage Then
                    .Text = "Create Multiproject Reports"
                    .OKButton.Text = "Create Report"
                    .filterLabel.Text = "Selection"
                Else
                    .Text = "Multiprojekt Reports erzeugen"
                    .OKButton.Text = "Bericht erstellen"
                    .filterLabel.Text = "Auswahl"
                End If

                '.rdbNameList.Enabled = False
                '.rdbNameList.Visible = False
                '.rdbNameList.Checked = False

                .rdbNameList.Enabled = True
                .rdbNameList.Visible = True
                .rdbNameList.Checked = True


                '.rdbProjStruktProj.Enabled = False
                '.rdbProjStruktProj.Visible = False
                '.rdbProjStruktProj.Checked = False

                '.rdbProjStruktTyp.Enabled = False
                '.rdbProjStruktTyp.Visible = False
                '.rdbProjStruktTyp.Checked = False

                .rdbProjStruktProj.Enabled = True
                .rdbProjStruktProj.Visible = True
                .rdbProjStruktProj.Checked = False

                .rdbProjStruktTyp.Enabled = True
                .rdbProjStruktTyp.Visible = True
                .rdbProjStruktTyp.Checked = False

                '.rdbPhases.Visible = False
                '.rdbPhases.Checked = False
                '.picturePhasen.Visible = False
                .rdbPhases.Visible = True
                .rdbPhases.Checked = True
                .picturePhasen.Visible = True

                '.rdbPhaseMilest.Visible = True
                .rdbPhaseMilest.Visible = False
                .rdbPhaseMilest.Checked = False

                '.picturePhaseMilest.Visible = True
                .picturePhaseMilest.Visible = False

                .rdbMilestones.Visible = True
                .rdbMilestones.Checked = False
                .pictureMilestones.Visible = True

                .rdbBU.Visible = False
                .pictureBU.Visible = False

                .rdbTyp.Visible = False
                .pictureTyp.Visible = False

                .rdbRoles.Visible = True
                .rdbRoles.Checked = False
                .pictureRoles.Visible = True

                .rdbCosts.Visible = True
                .rdbCosts.Checked = False
                .pictureCosts.Visible = True

                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .AbbrButton.Visible = False
                .AbbrButton.Enabled = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports
                .repVorlagenDropbox.Visible = True
                .labelPPTVorlage.Visible = True
                .einstellungen.Visible = True

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True


            ElseIf .menuOption = PTmenue.excelExport Then

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

                .AbbrButton.Visible = False
                .AbbrButton.Enabled = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True

                .einstellungen.Visible = False

            ElseIf .menuOption = PTmenue.vorlageErstellen Then

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

                .AbbrButton.Visible = False
                .AbbrButton.Enabled = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False
                .einstellungen.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True


            ElseIf .menuOption = PTmenue.reportBHTC Or _
                .menuOption = PTmenue.reportMultiprojektTafel Then



                If awinSettings.englishLanguage Then
                    .Text = "Create Project Report"
                    .OKButton.Text = "Create Report"
                    .filterLabel.Text = "Name of Report-Profile"
                Else
                    .Text = "Projekt-Report erzeugen"
                    .OKButton.Text = "Bericht erstellen"
                    .filterLabel.Text = "Name Report-Profil"
                End If

                ' temporäre Ausblendung von Rollen und Kosten im Modus BHTC 
                If .menuOption = PTmenue.reportBHTC Then

                    .rdbNameList.Enabled = False
                    .rdbNameList.Visible = False
                    .rdbNameList.Checked = False


                    .rdbProjStruktProj.Enabled = False
                    .rdbProjStruktProj.Visible = False
                    .rdbProjStruktProj.Checked = True

                    .rdbProjStruktTyp.Enabled = False
                    .rdbProjStruktTyp.Visible = False
                    .rdbProjStruktTyp.Checked = False


                    .rdbPhases.Visible = False
                    .rdbPhases.Checked = False
                    .picturePhasen.Visible = False

                    .rdbPhaseMilest.Visible = False
                    .rdbPhaseMilest.Checked = True
                    .picturePhaseMilest.Visible = False

                    .rdbMilestones.Visible = False
                    .rdbMilestones.Checked = False
                    .pictureMilestones.Visible = False

                    .rdbBU.Visible = False
                    .pictureBU.Visible = False

                    .rdbTyp.Visible = False
                    .pictureTyp.Visible = False

                    .rdbRoles.Visible = False
                    .rdbRoles.Checked = False
                    .pictureRoles.Visible = False

                    .rdbCosts.Visible = False
                    .rdbCosts.Checked = False
                    .pictureCosts.Visible = False
                Else
                    If .menuOption = PTmenue.reportMultiprojektTafel Then

                        .rdbNameList.Enabled = False
                        .rdbNameList.Visible = False
                        .rdbNameList.Checked = False


                        .rdbProjStruktProj.Enabled = True
                        .rdbProjStruktProj.Visible = True
                        .rdbProjStruktProj.Checked = True

                        .rdbProjStruktTyp.Enabled = True
                        .rdbProjStruktTyp.Visible = True
                        .rdbProjStruktTyp.Checked = False


                        .rdbPhases.Visible = False
                        .rdbPhases.Checked = False
                        .picturePhasen.Visible = False

                        .rdbPhaseMilest.Visible = True
                        .rdbPhaseMilest.Checked = True
                        .picturePhaseMilest.Visible = True

                        .rdbMilestones.Visible = False
                        .rdbMilestones.Checked = False
                        .pictureMilestones.Visible = False

                        .rdbBU.Visible = False
                        .pictureBU.Visible = False

                        .rdbTyp.Visible = False
                        .pictureTyp.Visible = False

                        .rdbRoles.Visible = False
                        .rdbRoles.Checked = False
                        .pictureRoles.Visible = False

                        .rdbCosts.Visible = False
                        .rdbCosts.Checked = False
                        .pictureCosts.Visible = False
                    End If

                End If
                ' Ende temporäre Anpassung 

                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .AbbrButton.Visible = False
                .AbbrButton.Enabled = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                .hryStufenLabel.Visible = False
                .hryStufen.Value = 50
                .hryStufen.Visible = False

                ' Reports
                .repVorlagenDropbox.Visible = True
                .labelPPTVorlage.Visible = True
                .einstellungen.Visible = True

                ' Filter
                .filterDropbox.DropDownStyle = ComboBoxStyle.Simple
                .filterDropbox.Visible = True
                .filterLabel.Visible = True

          
            End If

        End With


    End Sub


    Private Sub frmHierarchySelection_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.listselP, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.listselP, PTpinfo.left) = Me.Left

        awinSettings.isHryNameFrmActive = False
        If appInstance.ScreenUpdating = False Then
            appInstance.ScreenUpdating = True
        End If

        If appInstance.EnableEvents = False Then
            appInstance.EnableEvents = True
        End If

        If Not enableOnUpdate Then
            enableOnUpdate = True
        End If

    End Sub

    Private Sub frmHierarchySelection_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If frmCoord(PTfrm.listselP, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.listselP, PTpinfo.top))
            Me.Left = CInt(frmCoord(PTfrm.listselP, PTpinfo.left))
        Else
            Me.Top = 60
            Me.Left = 100
        End If

        Cursor = Cursors.WaitCursor

        awinSettings.isHryNameFrmActive = True

        ' Button Visibility und Texte definieren 
        Call defineFrmButtonVisibility()



        If Me.calledFrom = "MS-Project" Then
            Call retrieveProfilSelection(filterDropbox.Text, PTmenue.reportBHTC, selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts, repProfil)
            If IsNothing(repProfil) Then
                Throw New ArgumentException("Fehler beim Lesen des ausgewählten ReportProfils")
            End If
        Else   ' calledFrom = "Multiprojekt-Tafel"

            ' ur: 23.06.2017
            ' hier soll immer mit leeren Selektionen begonnen werden
            selectedMilestones.Clear()
            selectedPhases.Clear()
            'Call retrieveSelections("Last", PTmenue.visualisieren, selectedBUs, selectedTyps, selectedPhases, _
            '                        selectedMilestones, selectedRoles, selectedCosts)
            ' tk 8.4.17
            ' hier werden nur Phasen und Meilensteine selektiert: deswegen dürfen hier die anderen Collections nicht zählen
            selectedBUs.Clear()
            selectedTyps.Clear()
            selectedRoles.Clear()
            selectedCosts.Clear()

        End If

        Select Case Me.menuOption

            Case PTmenue.leistbarkeitsAnalyse

                'Me.rdbRoles.Checked = True
                'Me.rdbCosts.Checked = False
                '' Rollen oder Kosten hierarchisch darstellen

                Dim topNodes As List(Of Integer) = RoleDefinitions.getTopLevelNodeIDs

            Case PTmenue.visualisieren

                ' ur: 11.09.2017: beginnt mit ProjektStruktur
                'auswahl = selectionTyp(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)
                auswahl = PTProjektType.projekt

                Select Case auswahl
                    Case PTProjektType.nameList

                        Me.rdbNameList.Checked = True
                        Me.rdbPhases.Checked = True

                        Call buildHryTreeViewNew(PTProjektType.nameList)


                    Case PTProjektType.vorlage

                        Me.rdbProjStruktTyp.Checked = True

                        Call buildHryTreeViewNew(PTProjektType.vorlage)
                        '' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
                        'If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                        '    hryTreeView.ExpandAll()
                        'End If

                    Case PTProjektType.projekt

                        Me.rdbProjStruktProj.Checked = True

                        Call buildHryTreeViewNew(PTProjektType.projekt)
                        '' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
                        'If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                        '    hryTreeView.ExpandAll()
                        'End If

                    Case Else
                        selectedPhases.Clear()
                        selectedMilestones.Clear()
                        selectedBUs.Clear()
                        selectedTyps.Clear()
                        selectedRoles.Clear()
                        selectedCosts.Clear()

                        Me.rdbNameList.Checked = True
                        Me.rdbPhases.Checked = True

                        Call buildHryTreeViewNew(PTProjektType.nameList)

                End Select

                ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
                If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                    hryTreeView.ExpandAll()
                End If


            Case PTmenue.reportMultiprojektTafel
                ' ur: 11.09.2017: beginnt mit ProjektStruktur
                'auswahl = selectionTyp(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)

                Call retrieveProfilSelection(filterDropbox.Text, PTmenue.reportMultiprojektTafel, selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts, repProfil)
                If IsNothing(repProfil) Then
                    Throw New ArgumentException("Fehler beim Lesen des ausgewählten ReportProfils")
                End If

                auswahl = PTProjektType.projekt

                Select Case auswahl
                    Case PTProjektType.nameList

                        Me.rdbNameList.Checked = True
                        Me.rdbPhases.Checked = True

                        Call buildHryTreeViewNew(PTProjektType.nameList)


                    Case PTProjektType.vorlage

                        Me.rdbProjStruktTyp.Checked = True

                        Call buildHryTreeViewNew(PTProjektType.vorlage)
                        '' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
                        'If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                        '    hryTreeView.ExpandAll()
                        'End If

                    Case PTProjektType.projekt

                        'Me.rdbProjStruktProj.Checked = True

                        Call buildHryTreeViewNew(PTProjektType.projekt)
                        '' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
                        'If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                        '    hryTreeView.ExpandAll()
                        'End If

                    Case Else
                        selectedPhases.Clear()
                        selectedMilestones.Clear()
                        selectedBUs.Clear()
                        selectedTyps.Clear()
                        selectedRoles.Clear()
                        selectedCosts.Clear()

                        Me.rdbNameList.Checked = True
                        Me.rdbPhases.Checked = True

                        Call buildHryTreeViewNew(PTProjektType.nameList)

                End Select

                    ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
                    If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                        hryTreeView.ExpandAll()
                    End If

        End Select


        Cursor = Cursors.Default

        ' die Vorlagen  einlesen

        If IsNothing(repProfil) Then
            Call frmHryNameReadPPTVorlagen(Me.menuOption, repVorlagenDropbox)
        Else
            Call frmHryNameReadPPTVorlagen(Me.menuOption, repVorlagenDropbox, repProfil.isMpp)
        End If


        ' die Filter einlesen

        If Not (Me.menuOption = PTmenue.reportBHTC Or Me.menuOption = PTmenue.reportMultiprojektTafel) Then

            Call frmHryNameReadFilterVorlagen(Me.menuOption, filterDropbox)

            ' alle definierten Filter in ComboBox anzeigen
            If Me.menuOption = PTmenue.filterdefinieren Then

                For Each kvp As KeyValuePair(Of String, clsFilter) In filterDefinitions.Liste
                    filterDropbox.Items.Add(kvp.Key)
                Next

            Else

                For Each kvp As KeyValuePair(Of String, clsFilter) In selFilterDefinitions.Liste
                    filterDropbox.Items.Add(kvp.Key)
                Next

            End If

        Else
            '       Me.menuOption = PTmenue.reportBHTC
            '
            If Me.calledFrom = "MS-Project" Then

                If Not IsNothing(repProfil) Then
                    If My.Computer.FileSystem.FileExists(awinPath & RepProjectVorOrdner & "\" & repProfil.PPTTemplate) Then
                        repVorlagenDropbox.Text = repProfil.PPTTemplate
                    Else
                        repVorlagenDropbox.Text = ""
                    End If
                End If

            Else
                If Not IsNothing(repProfil) Then
                    If repProfil.isMpp Then
                        If My.Computer.FileSystem.FileExists(awinPath & RepPortfolioVorOrdner & "\" & repProfil.PPTTemplate) Then
                            repVorlagenDropbox.Text = repProfil.PPTTemplate
                        Else
                            repVorlagenDropbox.Text = ""
                        End If
                    Else
                        If My.Computer.FileSystem.FileExists(awinPath & RepProjectVorOrdner & "\" & repProfil.PPTTemplate) Then
                            repVorlagenDropbox.Text = repProfil.PPTTemplate
                        Else
                            repVorlagenDropbox.Text = ""
                        End If
                    End If

                End If
            End If


        End If



    End Sub



    ''' <summary>
    ''' Behandlung Drücken OK Button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim anzahlKnoten As Integer
        Dim selectedNode As TreeNode
        Dim tmpNode As TreeNode
        Dim filterName As String = ""
        Dim element As String
        Dim type As Integer = -1
        Dim pvName As String = ""

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        Dim formerEoU As Boolean = enableOnUpdate
        enableOnUpdate = False

        statusLabel.Text = ""


        anzahlKnoten = hryTreeView.Nodes.Count
        selectedNode = hryTreeView.SelectedNode


        If Me.rdbNameList.Checked Then

            ' hier muss jetzt noch der aktuelle rdb ausgelesen werden ..
            If Me.rdbPhases.Checked = True Then

                selectedPhases.Clear()
                With hryTreeView
                    For px As Integer = 1 To anzahlKnoten
                        tmpNode = .Nodes.Item(px - 1)
                        If tmpNode.Checked Then
                            ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll
                            If Not selectedPhases.Contains(tmpNode.Name) Then
                                selectedPhases.Add(tmpNode.Name, tmpNode.Name)
                            End If
                        End If
                    Next
                End With



            ElseIf Me.rdbMilestones.Checked = True Then

                selectedMilestones.Clear()
                With hryTreeView
                    For px As Integer = 1 To anzahlKnoten
                        tmpNode = .Nodes.Item(px - 1)
                        If tmpNode.Checked Then
                            ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll
                            If Not selectedMilestones.Contains(tmpNode.Name) Then
                                selectedMilestones.Add(tmpNode.Name, tmpNode.Name)
                            End If
                        End If
                    Next
                End With
            End If


        ElseIf Me.rdbProjStruktProj.Checked Or Me.rdbProjStruktTyp.Checked Then

            ' Radiobutton Projekt-Struktur  wurde geklickt

            selectedPhases.Clear()
            selectedMilestones.Clear()


            With hryTreeView

                Dim hry As clsHierarchy = Nothing
                For px As Integer = 1 To anzahlKnoten

                    tmpNode = .Nodes.Item(px - 1)

                    ' jetzt muss das Projekt, die Projekt-Vorlage ermittelt werden 
                    ' und daraus die Hierarchie 
                    If tmpNode.Level = 0 Then
                        hry = getHryFromNode(tmpNode)
                        type = getTypeFromNode(tmpNode)
                        pvName = getPVnameFromNode(tmpNode)
                        If tmpNode.Checked And Not subNodesSelected(tmpNode) Then

                            Dim tmpBreadcrumb As String = hry.getBreadCrumb(rootPhaseName, CInt(hryStufen.Value))
                            Dim elemName As String = elemNameOfElemID(rootPhaseName)
                            Dim selElem As String = calcHryFullname(elemName, tmpBreadcrumb, getPVkennungFromNode(tmpNode))
                            If Not selectedPhases.Contains(selElem) Then
                                selectedPhases.Add(selElem, selElem)
                            End If

                        End If
                    End If


                    If tmpNode.Checked And Not IsNothing(hry) And tmpNode.Level > 0 Then
                        ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll 
                        Dim filterbyLevel0 As Boolean = topNodeIsSelected(tmpNode)
                        Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
                        Dim elemName As String = elemNameOfElemID(tmpNode.Name)
                        If filterbyLevel0 Then
                            element = calcHryFullname(elemName, tmpBreadcrumb, getPVkennungFromNode(tmpNode))
                        Else
                            element = calcHryFullname(elemName, tmpBreadcrumb)
                        End If


                        If elemIDIstMeilenstein(tmpNode.Name) Then
                            If Not selectedMilestones.Contains(element) Then
                                selectedMilestones.Add(element, element)
                            End If
                        Else
                            If Not selectedPhases.Contains(element) Then
                                selectedPhases.Add(element, element)
                            End If

                        End If

                    End If


                    If tmpNode.Nodes.Count > 0 Then
                        Call pickupCheckedItems(tmpNode, hry)
                    End If

                Next

            End With

        ElseIf rdbRoles.Checked = True Then '?????

            ' Radiobutton Rollen wurde geklickt

            selectedRoles.Clear()
        
            With hryTreeView

                For px As Integer = 1 To anzahlKnoten

                    tmpNode = .Nodes.Item(px - 1)

                    If tmpNode.Checked 

                        If Not selectedRoles.Contains(tmpNode.Text) Then
                            selectedRoles.Add(tmpNode.Text, tmpNode.Text)
                        End If

                    End If


                    If tmpNode.Nodes.Count > 0 Then
                        Call pickupCheckedRoleItems(tmpNode)
                    End If

                Next

            End With



        ElseIf rdbCosts.Checked = True Then


            selectedCosts.Clear()
            With hryTreeView
                For px As Integer = 1 To anzahlKnoten
                    tmpNode = .Nodes.Item(px - 1)
                    If tmpNode.Checked Then
                        ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll
                        If Not selectedCosts.Contains(tmpNode.Name) Then
                            selectedCosts.Add(tmpNode.Name, tmpNode.Name)
                        End If
                    End If
                Next
            End With


        ElseIf rdbBU.Checked = True Then

            selectedBUs.Clear()
            With hryTreeView
                For px As Integer = 1 To anzahlKnoten
                    tmpNode = .Nodes.Item(px - 1)
                    If tmpNode.Checked Then
                        ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll
                        If Not selectedBUs.Contains(tmpNode.Name) Then
                            selectedBUs.Add(tmpNode.Name, tmpNode.Name)
                        End If
                    End If
                Next
            End With


        ElseIf rdbTyp.Checked = True Then

            selectedTyps.Clear()
            With hryTreeView
                For px As Integer = 1 To anzahlKnoten
                    tmpNode = .Nodes.Item(px - 1)
                    If tmpNode.Checked Then
                        ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll
                        If Not selectedTyps.Contains(tmpNode.Name) Then
                            selectedTyps.Add(tmpNode.Name, tmpNode.Name)
                        End If
                    End If
                Next
            End With

        End If

     

        If Me.menuOption = PTmenue.filterdefinieren Then

            filterName = filterDropbox.Text
            ' jetzt wird der Filter unter dem Namen filterName gespeichert ..
            Call storeFilter(filterName, menuOption, selectedBUs, selectedTyps, _
                                                   selectedPhases, selectedMilestones, _
                                                   selectedRoles, selectedCosts, False)
            ' tk 18.11.15 braucht man nicht, weil hier nur Phasen / Meilensteine ausgewählt werden können
            'ElseIf Me.menuOption = PTmenue.visualisieren Then

            '    If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0) And _
            '        (selectedRoles.Count > 0 Or selectedCosts.Count > 0) Then
            '        Call MsgBox("es können nur entweder Phasen / Meilensteine oder Rollen oder Kosten angezeigt werden")
            '        ''Else
            '        ''    filterName = filterDropbox.Text
            '        ''    ' jetzt wird der Filter unter dem Namen filterName gespeichert ..
            '        ''    Call storeFilter(filterName, menuOption, selectedBUs, selectedTyps, _
            '        ''                                           selectedPhases, selectedMilestones, _
            '        ''                                           selectedRoles, selectedCosts, False)
            '    End If

            ''Else    ' alle anderen PTmenues

            ''    filterName = filterDropbox.Text
            ''    ' jetzt wird der Filter unter dem Namen filterName gespeichert ..
            ''    Call storeFilter(filterName, menuOption, selectedBUs, selectedTyps, _
            ''                                           selectedPhases, selectedMilestones, _
            ''                                           selectedRoles, selectedCosts, False)
        End If

        ' jetzt wird der letzte Filter gespeichert ..
        Dim lastfilter As String = "Last"
        If Not (Me.menuOption = PTmenue.reportBHTC Or Me.menuOption = PTmenue.reportMultiprojektTafel) Then
            Call storeFilter(lastfilter, menuOption, selectedBUs, selectedTyps, _
                                                   selectedPhases, selectedMilestones, _
                                                   selectedRoles, selectedCosts, True)
        End If


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
            Me.menuOption = PTmenue.vorlageErstellen Or _
            Me.menuOption = PTmenue.reportBHTC Or Me.menuOption = PTmenue.reportMultiprojektTafel Then
            validOption = True
        ElseIf showRangeRight - showRangeLeft >= minColumns - 1 Then
            validOption = True
        Else
            validOption = False
        End If

        If Me.menuOption = PTmenue.multiprojektReport Or Me.menuOption = PTmenue.einzelprojektReport Or _
            Me.menuOption = PTmenue.reportBHTC Or Me.menuOption = PTmenue.reportMultiprojektTafel Then

            If ((selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0 Or selectedTyps.Count > 0) _
                    And validOption) Or _
                    (Me.menuOption = PTmenue.reportBHTC And validOption) Then

                Dim vorlagenDateiName As String

                If Me.menuOption = PTmenue.multiprojektReport Then
                    vorlagenDateiName = awinPath & RepPortfolioVorOrdner & _
                                    "\" & repVorlagenDropbox.Text
                ElseIf Me.menuOption = PTmenue.einzelprojektReport Then

                    vorlagenDateiName = awinPath & RepProjectVorOrdner & _
                                    "\" & repVorlagenDropbox.Text

                Else

                    If Not IsNothing(repProfil) Then
                        If repProfil.isMpp Then
                            vorlagenDateiName = awinPath & RepPortfolioVorOrdner & _
                                    "\" & repVorlagenDropbox.Text
                        Else

                            vorlagenDateiName = awinPath & RepProjectVorOrdner & _
                                            "\" & repVorlagenDropbox.Text
                        End If
                    Else
                        ' im zweifelsfall werden die Portfolio Vorlagen angezeigt
                        vorlagenDateiName = awinPath & RepPortfolioVorOrdner & _
                                                            "\" & repVorlagenDropbox.Text
                    End If
                End If

                ' Prüfen, ob die Datei überhaupt existirt 
                If repVorlagenDropbox.Text.Length = 0 Then
                    ' Call MsgBox("bitte PPT Vorlage auswählen !")
                    If awinSettings.englishLanguage Then
                        Me.statusLabel.Text = "please select a PPT template !"
                    Else
                        Me.statusLabel.Text = "bitte PPT Vorlage auswählen !"
                    End If

                    Me.statusLabel.Visible = True
                ElseIf My.Computer.FileSystem.FileExists(vorlagenDateiName) Then

                    Try

                        OKButton.Enabled = False
                        OKButton.Visible = False
                        repVorlagenDropbox.Enabled = False

                        With AbbrButton
                            .Cursor = Cursors.Arrow
                            .Enabled = True
                            .Visible = True
                            .Left = OKButton.Left
                            .Top = OKButton.Top
                        End With


                        statusLabel.Text = ""
                        statusLabel.Visible = True

                        Me.Cursor = Cursors.WaitCursor
                        If awinSettings.englishLanguage Then
                            AbbrButton.Text = "Cancel"
                        Else
                            AbbrButton.Text = "Abbrechen"
                        End If


                        ' Alternativ ohne Background Worker
                        If Me.menuOption = PTmenue.reportBHTC Or Me.menuOption = PTmenue.reportMultiprojektTafel Then

                            'Call MsgBox("Report erstellen mit Projekt " & repProfil.VonDate.ToString & " bis " & repProfil.BisDate.ToString & " Reportprofil " & repProfil.name)

                            If menuOption = PTmenue.reportMultiprojektTafel Then
                                If Not repProfil.isMpp And selectedProjekte.Count < 1 Then
                                    Throw New ArgumentException("Zum Erstellen des Reports muss ein Projekt ausgewählt sein")
                                ElseIf repProfil.isMpp And _
                                    Not (showRangeLeft > 0 And showRangeRight > showRangeLeft) Then  ' Zeitraum wurde nicht gesetzt
                                    Throw New ArgumentException("Zum Erstellen des Reports muss ein ein Zeitraum gesetzt sein")
                                End If
                            End If


                            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

                            repProfil.PPTTemplate = repVorlagenDropbox.Text

                            BackgroundWorker3.RunWorkerAsync(repProfil)

                        Else

                            'Call PPTstarten()
                            BackgroundWorker1.RunWorkerAsync(vorlagenDateiName)

                        End If


                    Catch ex As Exception
                        Me.Cursor = System.Windows.Forms.Cursors.Arrow
                        Call MsgBox(ex.Message)
                        appInstance.EnableEvents = formerEE
                        enableOnUpdate = formerEoU
                        MyBase.Close()
                    End Try

                Else

                    If awinSettings.englishLanguage Then
                        Me.statusLabel.Text = "please select a PPT template !"
                    Else
                        Me.statusLabel.Text = "bitte PPT Vorlage auswählen !"
                    End If

                    Me.statusLabel.Visible = True
                End If


            Else
                'Call MsgBox("bitte mindestens ein Element selektieren bzw. " & vbLf & "einen Zeitraum angeben ...")
                If awinSettings.englishLanguage Then
                    Me.statusLabel.Text = "please select at least one planelement resp. " & vbLf & _
                             "provide a timespan ..."
                Else
                    Me.statusLabel.Text = "bitte mindestens ein Element selektieren bzw. " & vbLf & _
                             "einen Zeitraum angeben ..."
                End If

                Me.statusLabel.Visible = True
            End If

        Else
            ' die Aktion Subroutine aufrufen 
            ' hier können nur Phasen / Meilensteine ausgewählt werden; 
            Dim tmpCollection As New Collection
            If rdbPhases.Checked Or rdbMilestones.Checked _
                Or rdbRoles.Checked Or rdbCosts.Checked Then
                Call frmHryNameActions(Me.menuOption, selectedPhases, selectedMilestones, _
                            selectedRoles, selectedCosts, Me.chkbxOneChart.Checked, filterName)
                ''ElseIf rdbRoles.Checked Then
                ''    Call frmHryNameActions(Me.menuOption, tmpCollection, tmpCollection, _
                ''                selectedRoles, tmpCollection, Me.chkbxOneChart.Checked, filterName)
                ''ElseIf rdbCosts.Checked Then
                ''    Call frmHryNameActions(Me.menuOption, tmpCollection, tmpCollection, _
                ''                tmpCollection, selectedCosts, Me.chkbxOneChart.Checked, filterName)
            Else
                Call frmHryNameActions(Me.menuOption, selectedPhases, selectedMilestones, _
                                tmpCollection, tmpCollection, Me.chkbxOneChart.Checked, lastfilter)
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

    Private Sub einstellungen_Click(sender As Object, e As EventArgs) Handles einstellungen.Click

        Dim mppFrm As New frmMppSettings
        Dim dialogreturn As DialogResult


        If Me.menuOption = PTmenue.reportBHTC Or Me.menuOption = PTmenue.reportMultiprojektTafel Then
            mppFrm.calledfrom = "frmBHTC"

            With awinSettings

                If Not IsNothing(repProfil) Then
                    ' tk Änderung 5.4. wird für Darstellung Projekt auf Multiprojekt-Tafel benötigt; hier nicht setzen 
                    '.drawProjectLine = True
                    .mppOnePage = repProfil.OnePage
                    .mppShowLegend = repProfil.Legend
                    .mppShowMsDate = repProfil.MSDate
                    .mppShowMsName = repProfil.MSName
                    .mppShowPhDate = repProfil.PhDate
                    .mppShowPhName = repProfil.PhName
                    .mppVertikalesRaster = repProfil.VLinien
                    .mppShowHorizontals = repProfil.ShowHorizontals
                    .mppUseAbbreviation = repProfil.UseAbbreviation
                    .mppKwInMilestone = repProfil.KwInMilestone

                    ' für BHTC immer true
                    .mppExtendedMode = repProfil.ExtendedMode
                    ' für BHTC immer false
                    .mppShowAmpel = repProfil.Ampeln
                    .mppShowAllIfOne = repProfil.AllIfOne
                    .mppFullyContained = repProfil.FullyContained
                    .mppSortiertDauer = repProfil.SortedDauer
                    .mppShowProjectLine = repProfil.ProjectLine
                    .mppUseOriginalNames = repProfil.UseOriginalNames

                End If


            End With
        Else
            mppFrm.calledfrom = "frmShowPlanElements"

        End If



        dialogreturn = mppFrm.ShowDialog


        If Me.menuOption = PTmenue.reportBHTC Or Me.menuOption = PTmenue.reportMultiprojektTafel Then

            With awinSettings

                ' tk Änderung 5.4. wird für Darstellung Projekt auf Multiprojekt-Tafel benötigt; hier nicht setzen 
                '.drawProjectLine = True

                If Not IsNothing(repProfil) Then

                    repProfil.ExtendedMode = .mppExtendedMode
                    repProfil.OnePage = .mppOnePage
                    repProfil.AllIfOne = .mppShowAllIfOne
                    repProfil.Ampeln = .mppShowAmpel
                    repProfil.Legend = .mppShowLegend
                    repProfil.MSDate = .mppShowMsDate
                    repProfil.MSName = .mppShowMsName
                    repProfil.PhDate = .mppShowPhDate
                    repProfil.PhName = .mppShowPhName
                    repProfil.ProjectLine = .mppShowProjectLine
                    repProfil.SortedDauer = .mppSortiertDauer
                    repProfil.VLinien = .mppVertikalesRaster
                    repProfil.FullyContained = .mppFullyContained
                    repProfil.ShowHorizontals = .mppShowHorizontals
                    repProfil.UseAbbreviation = .mppUseAbbreviation
                    repProfil.UseOriginalNames = .mppUseOriginalNames
                    repProfil.KwInMilestone = .mppKwInMilestone

                End If

            End With
        End If

    End Sub

    ''' <summary>
    ''' gibt zurück, ob das Projekt / die Vorlage selektiert ist: dann wirkt das als Filter 
    ''' </summary>
    ''' <param name="node"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function topNodeIsSelected(ByVal node As TreeNode) As Boolean
        Dim curNode As TreeNode = node
        Dim tmpResult As Boolean = False

        If Not IsNothing(curNode) Then
            ' gehe auf den root-Knoten
            Do While Not IsNothing(curNode.Parent)
                curNode = curNode.Parent
            Loop
            tmpResult = curNode.Checked
        End If

        topNodeIsSelected = tmpResult

    End Function

    ''' <summary>
    ''' gibt die Hierarchie des Root-Knotens des betreffenden Knotens zurück 
    ''' </summary>
    ''' <param name="node"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getHryFromNode(ByVal node As TreeNode) As clsHierarchy
        Dim tmpResult As clsHierarchy = Nothing


        Dim pvName As String = getPVnameFromNode(node)
        Dim type As Integer = getTypeFromNode(node)

        If type = PTProjektType.vorlage Then

            If Projektvorlagen.Contains(pvName) Then
                tmpResult = Projektvorlagen.getProject(pvName).hierarchy
            End If

        Else
            If ShowProjekte.contains(pvName) Then
                tmpResult = ShowProjekte.getProject(pvName).hierarchy
            End If

        End If

        getHryFromNode = tmpResult
    End Function

    ''' <summary>
    ''' gibt den pvname des fullpaths zurück ... 
    ''' </summary>
    ''' <param name="tmpNode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getPVnameFromNode(ByVal tmpNode As TreeNode) As String
        Dim tmpResult As String = ""
        Dim curNode As TreeNode = tmpNode

        ' gehe auf den root-Knoten
        Do While Not IsNothing(curNode.Parent)
            curNode = curNode.Parent
        Loop

        If curNode.Name.StartsWith("P:") Or _
            curNode.Name.StartsWith("V:") Then

            Dim tmpStr() As String = curNode.Name.Split(New Char() {CChar(":")})
            If tmpStr.Length >= 2 Then
                tmpResult = tmpStr(1)
            End If

        End If

        getPVnameFromNode = tmpResult

    End Function

    ''' <summary>
    ''' gibt den Type zurück: 0=Vorlage, 1=Projekt
    ''' </summary>
    ''' <param name="tmpNode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getTypeFromNode(ByVal tmpNode As TreeNode) As Integer
        Dim tmpResult As Integer = -1

        Dim curNode As TreeNode = tmpNode

        ' gehe auf den root-Knoten
        Do While Not IsNothing(curNode.Parent)
            curNode = curNode.Parent
        Loop

        If curNode.Name.StartsWith("V:") Then
            tmpResult = PTProjektType.vorlage
        ElseIf curNode.Name.StartsWith("P:") Then
            tmpResult = PTProjektType.projekt
        End If


        getTypeFromNode = tmpResult

    End Function

    ''' <summary>
    ''' liefert die gesamte Kennung zurück , 
    ''' wird für den Aufbau der Item-Einträge in selectedPhases, selectedMilestones benötigt 
    ''' </summary>
    ''' <param name="tmpNode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getPVkennungFromNode(ByVal tmpNode As TreeNode) As String
        Dim tmpResult As String = ""

        If Not IsNothing(tmpNode) Then
            Dim curNode As TreeNode = tmpNode

            ' gehe auf den root-Knoten
            Do While Not IsNothing(curNode.Parent)
                curNode = curNode.Parent
            Loop

            tmpResult = curNode.Name
        End If

        getPVkennungFromNode = tmpResult
    End Function
    ''' <summary>
    ''' prüft, ob einer dem Knoten tmpNode untergeordneten Knoten selektiert ist
    ''' </summary>
    ''' <param name="tmpNode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function subNodesSelected(ByVal tmpNode As TreeNode) As Boolean
        Dim curNode As TreeNode
        Dim i As Integer = 1
        Dim erg As Boolean = False


        With tmpNode

            While i <= .Nodes.Count And Not erg

                curNode = .Nodes.Item(i - 1)
                If curNode.Checked Then
                    erg = True
                Else
                    erg = subNodesSelected(curNode)
                End If
                i = i + 1

            End While
        End With
        subNodesSelected = erg

    End Function

    Private Sub hryTreeView_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles hryTreeView.AfterCheck
        Dim oNode As TreeNode
        Dim hnode As TreeNode
        Dim anzCheckedNodes As Integer = 0

        oNode = e.Node

        If rdbRoles.Checked Then
            ' keine Sonderbehandlung der obersten Knoten bei Rollen-Hierarchie
        Else


            If Not rdbNameList.Checked Then

                If oNode.Level = 0 Then

                    If Not oNode.Checked Then
                        Call unCheck(oNode)
                    End If

                Else
                    hnode = oNode

                    ' finde den obersten Node
                    While Not IsNothing(hnode.Parent)
                        hnode = hnode.Parent
                    End While

                    If oNode.Checked Then

                        ' Wenn oberster Node nicht gecheckt, dann check ihn
                        If hnode.Level = 0 And Not hnode.Checked Then
                            hnode.Checked = True
                        End If

                    Else ' not oNode.checked 

                        Dim allUnselected As Boolean

                        If hnode.Level = 0 And hnode.Checked Then

                            allUnselected = Not subNodesSelected(hnode)

                            'If Not subNodesSelected(hnode) Then
                            If allUnselected Then
                                hnode.Checked = False
                            End If
                        End If

                    End If

                End If

            End If

        End If    ' Ende von If rdbroles.checked


    End Sub

    Private Sub hryTreeView_BeforeCheck(sender As Object, e As TreeViewCancelEventArgs) Handles hryTreeView.BeforeCheck
      
    End Sub


    Private Sub hryTreeView_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles hryTreeView.BeforeExpand

        Dim node As TreeNode
        Dim childNode As TreeNode
        Dim placeholder As TreeNode
        Dim elemID As String
        Dim hryNode As clsHierarchyNode
        Dim anzChilds As Integer
        Dim childNameID As String
        Dim PVname As String = getPVnameFromNode(e.Node)
        Dim type As Integer = getTypeFromNode(e.Node)
        Dim curHry As clsHierarchy
        Dim vorlElem As String = ""

        Dim childRole As clsRollenDefinition

        node = e.Node
        elemID = node.Name

        If rdbRoles.Checked Then
            ' Rollen expandieren

            If Not IsNothing(node.Tag) Then

                ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
                If node.Tag = "P" Then

                    node.Tag = "X"

                    ' Löschen von Platzhalter
                    node.Nodes.Clear()

                    Dim nodelist As New SortedList(Of Integer, String)
                    Try
                        nodelist = RoleDefinitions.getRoleDefByID(CInt(node.Name)).getSubRoleIDs
                        anzChilds = nodelist.Count
                    Catch ex As Exception
                        anzChilds = 0
                    End Try



                    With hryTreeView
                        .CheckBoxes = True

                        For i As Integer = 0 To anzChilds - 1
                            childRole = RoleDefinitions.getRoleDefByID(nodelist.ElementAt(i).Key)
                            Dim childName As String = childRole.name
                            Dim childID As Integer = childRole.UID

                            If allRoles.Contains(childName) Then

                                childNode = node.Nodes.Add(childName)
                                childNode.Name = childID.ToString
                                childNode.Text = childName


                                If selectedRoles.Contains(childName) Then
                                    childNode.Checked = True
                                End If

                                Dim anzSubRolesOFChild As Integer
                                Try
                                    anzSubRolesOFChild = RoleDefinitions.getRoleDefByID(childID).getSubRoleIDs.Count
                                Catch ex As Exception
                                    anzSubRolesOFChild = 0
                                End Try

                                If anzSubRolesOFChild > 0 Then
                                    childNode.Tag = "P"


                                    placeholder = childNode.Nodes.Add("-")
                                    placeholder.Tag = "P"
                                Else
                                    childNode.Tag = "X"
                                End If

                            End If
                        Next


                    End With
                End If
            End If


        Else
            ' Projekte expandieren
            If type = PTProjektType.vorlage Then
                curHry = Projektvorlagen.getProject(PVname).hierarchy
            Else
                curHry = ShowProjekte.getProject(PVname).hierarchy
            End If


            If Not IsNothing(node.Tag) Then

                ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
                If node.Tag = "P" Then

                    node.Tag = "X"

                    ' Löschen von Platzhalter
                    node.Nodes.Clear()

                    hryNode = curHry.nodeItem(elemID)

                    anzChilds = hryNode.childCount

                    With hryTreeView
                        .CheckBoxes = True

                        For i As Integer = 1 To anzChilds

                            childNameID = hryNode.getChild(i)
                            childNode = node.Nodes.Add(elemNameOfElemID(childNameID))
                            childNode.Name = childNameID


                            Dim tmpBreadcrumb As String = curHry.getBreadCrumb(childNameID, CInt(hryStufen.Value))
                            Dim elemName As String = elemNameOfElemID(childNameID)
                            Dim ele As String = calcHryFullname(elemName, tmpBreadcrumb)

                            ' gehe auf den root-Knoten
                            Dim topNode As TreeNode = node
                            Do While Not IsNothing(topNode.Parent)
                                topNode = topNode.Parent
                            Loop
                            Dim pvElem As String = "[" & topNode.Name & "]" & ele

                            If Projektvorlagen.Contains(topNode.Text) Then
                                Dim vproj As clsProjektvorlage = Projektvorlagen.getProject(topNode.Text)
                            End If

                            If ShowProjekte.contains(topNode.Text) Then

                                Dim hproj As clsProjekt = ShowProjekte.getProject(topNode.Text)
                                vorlElem = "[V:" & hproj.VorlagenName & "]" & ele
                            End If


                            If elemIDIstMeilenstein(childNameID) Then
                                childNode.BackColor = System.Drawing.Color.Azure
                                If selectedMilestones.Contains(ele) Or selectedMilestones.Contains(pvElem) _
                                    Or selectedMilestones.Contains(vorlElem) Or selectedMilestones.Contains(elemName) Then
                                    childNode.Checked = True
                                End If
                            Else
                                If selectedPhases.Contains(ele) Or selectedPhases.Contains(pvElem) _
                                   Or selectedPhases.Contains(vorlElem) Or selectedPhases.Contains(elemName) Then
                                    childNode.Checked = True
                                End If
                            End If



                            If curHry.nodeItem(childNameID).childCount > 0 Then
                                childNode.Tag = "P"
                                placeholder = childNode.Nodes.Add("-")
                                placeholder.Tag = "P"
                            Else
                                childNode.Tag = "X"
                            End If


                        Next

                    End With


                End If

            End If
        End If




    End Sub

    ' ''' <summary>
    ' ''' baut den TreeView für die Hierarchie auf , Treeview enthält sowohl Meilensteine als auch Phasen
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Private Sub buildHryTreeView()

    '    Dim hryNode As clsHierarchyNode
    '    Dim anzChilds As Integer
    '    Dim childNameID As String
    '    Dim nodeLevel0 As TreeNode
    '    Dim nodeLevel1 As TreeNode

    '    With hryTreeView
    '        .Nodes.Clear()
    '    End With

    '    If hry.count >= 1 Then
    '        hryNode = hry.nodeItem(rootPhaseName)

    '        anzChilds = hryNode.childCount

    '        With hryTreeView
    '            .CheckBoxes = True

    '            For i As Integer = 1 To anzChilds

    '                childNameID = hryNode.getChild(i)
    '                nodeLevel0 = .Nodes.Add(elemNameOfElemID(childNameID))
    '                nodeLevel0.Name = childNameID

    '                Dim tmpBreadcrumb As String = hry.getBreadCrumb(childNameID, CInt(hryStufen.Value))
    '                Dim elemName As String = elemNameOfElemID(childNameID)
    '                Dim element As String = calcHryFullname(elemName, tmpBreadcrumb)


    '                If elemIDIstMeilenstein(childNameID) Then
    '                    nodeLevel0.BackColor = System.Drawing.Color.Azure
    '                    If selectedMilestones.Contains(element) Or selectedMilestones.Contains(elemName) Then
    '                        nodeLevel0.Checked = True
    '                    End If
    '                Else

    '                    If selectedPhases.Contains(element) Or selectedPhases.Contains(elemName) Then
    '                        nodeLevel0.Checked = True
    '                    End If
    '                End If


    '                If hry.nodeItem(childNameID).childCount > 0 Then
    '                    nodeLevel0.Tag = "P"
    '                    nodeLevel1 = nodeLevel0.Nodes.Add("-")
    '                    nodeLevel1.Tag = "P"
    '                Else
    '                    nodeLevel0.Tag = "X"
    '                End If


    '            Next

    '        End With

    '    Else
    '        If awinSettings.englishLanguage Then
    '            Call MsgBox("there is no hierarchy")
    '        Else
    '            Call MsgBox("es ist keine Hierarchie gegeben")
    '        End If

    '    End If
    'End Sub

    ''' <summary>
    ''' baut den TreeView für die Hierarchie auf , Treeview enthält Projekt-Vorlagen oder Projekte, dann 
    ''' Meilensteine als auch Phasen
    '''
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub buildHryTreeViewNew(ByVal auswahl As Integer)

        Dim topLevel As TreeNode
        Dim kennung As String ' V: für Vorlagen, P: für Projekte
        Dim hry As clsHierarchy
        Dim checkProj As Boolean = False
        Dim projekteToLook As clsProjekte

        With hryTreeView
            .Nodes.Clear()
            .CheckBoxes = True


            If auswahl = PTProjektType.vorlage Then

                ' alle Templates zeigen 
                kennung = "V:"

                If selectedProjekte.Count > 0 And ShowProjekte.Count > 0 Then
                    If menuOption = PTmenue.multiprojektReport Then
                        projekteToLook = ShowProjekte
                    ElseIf menuOption = PTmenue.einzelprojektReport Then
                        projekteToLook = selectedProjekte
                    ElseIf menuOption = PTmenue.leistbarkeitsAnalyse Then
                        projekteToLook = ShowProjekte
                    ElseIf menuOption = PTmenue.reportMultiprojektTafel Then
                        If repProfil.isMpp Then
                            projekteToLook = ShowProjekte
                        Else
                            projekteToLook = selectedProjekte
                        End If
                    Else
                        projekteToLook = ShowProjekte
                    End If
                Else
                    If selectedProjekte.Count > 0 Then
                        projekteToLook = selectedProjekte
                    ElseIf ShowProjekte.Count > 0 Then
                        projekteToLook = ShowProjekte
                    Else
                        projekteToLook = ShowProjekte
                    End If
                End If


                For Each kvp As KeyValuePair(Of String, clsProjektvorlage) In Projektvorlagen.Liste

                    If projekteToLook.getTypNames().Contains(kvp.Key) Then

                        If kvp.Value.hierarchy.count > 0 Then
                            topLevel = .Nodes.Add(kvp.Key)
                            topLevel.Name = kennung & kvp.Key
                            topLevel.Text = kvp.Key

                            hry = kvp.Value.hierarchy

                            Dim projVorlage As clsProjektvorlage = Projektvorlagen.getProject(kvp.Key)
                            Dim nodeToCheck As Boolean = False

                            If selectedPhases.Count > 0 Then
                                nodeToCheck = projVorlage.containsAnyPhasesOfCollection(selectedPhases)
                            Else
                                nodeToCheck = False
                            End If

                            If selectedMilestones.Count > 0 Then
                                nodeToCheck = nodeToCheck Or projVorlage.containsAnyMilestonesOfCollection(selectedMilestones)
                            Else
                                nodeToCheck = nodeToCheck Or False
                            End If

                            If nodeToCheck Then
                                topLevel.Checked = True
                            End If

                            Call buildProjectSubTreeView(topLevel, hry)
                        End If
                    End If

                Next
            ElseIf auswahl = PTProjektType.projekt Then

                ' alle selektierten Projekte zeigen 
                kennung = "P:"

                If selectedProjekte.Count > 0 And ShowProjekte.Count > 0 Then
                    If menuOption = PTmenue.multiprojektReport Then
                        projekteToLook = ShowProjekte
                    ElseIf menuOption = PTmenue.einzelprojektReport Then
                        projekteToLook = selectedProjekte
                    ElseIf menuOption = PTmenue.leistbarkeitsAnalyse Then
                        projekteToLook = ShowProjekte
                    ElseIf menuOption = PTmenue.reportMultiprojektTafel Then
                        If repProfil.isMpp Then
                            projekteToLook = ShowProjekte
                        Else
                            projekteToLook = selectedProjekte
                        End If
                    Else
                        projekteToLook = ShowProjekte
                    End If
                Else
                    If selectedProjekte.Count > 0 Then
                        projekteToLook = selectedProjekte
                    ElseIf ShowProjekte.Count > 0 Then
                        projekteToLook = ShowProjekte
                    Else
                        projekteToLook = ShowProjekte
                    End If
                End If


                For Each kvp As KeyValuePair(Of String, clsProjekt) In projekteToLook.Liste

                    If kvp.Value.hierarchy.count > 0 Then
                        topLevel = .Nodes.Add(kvp.Key)
                        topLevel.Name = kennung & kvp.Key
                        topLevel.Text = kvp.Key
                        hry = kvp.Value.hierarchy

                        If selectedPhases.Count > 0 Or selectedMilestones.Count > 0 Then
                            ' überprüfen, ob das Projekt irgend eine der selektierten Phasen oder Meilensteine enthält
                            Dim hproj As clsProjekt = projekteToLook.getProject(kvp.Key)
                            Dim tmpcollection As New Collection
                            Dim newFil As New clsFilter("tmp", tmpcollection, tmpcollection, _
                                                        selectedPhases, selectedMilestones, tmpcollection, tmpcollection)
                            If newFil.doesNotBlock(hproj) Then
                                topLevel.Checked = True
                            End If
                        End If

                        Call buildProjectSubTreeView(topLevel, hry)
                    End If

                Next


            ElseIf auswahl = PTProjektType.nameList Then

                'alle Phasen der selektierten Projekte zeigen, je nach menuOption

                If Me.rdbPhases.Checked Then
                    ' clear Listbox1 
                    If awinSettings.englishLanguage Then
                        headerLine.Text = "Phases"
                    Else
                        headerLine.Text = "Phasen"
                    End If

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

                ElseIf Me.rdbMilestones.Checked Then

                    'alle Meilensteine der selektierten Projekte zeigen, je nach menuOption

                    statusLabel.Text = ""
                    filterBox.Enabled = True

                    ' clear Listbox1 
                    If awinSettings.englishLanguage Then
                        headerLine.Text = "Milestones"
                    Else
                        headerLine.Text = "Meilensteine"
                    End If

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
                    ' hier müssen noch Rollen, Kosten, Bu, Typ bearbeitet werden
                End If



            Else
                ' alle Projekte zeigen 
                kennung = "P:"
                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    If kvp.Value.hierarchy.count > 0 Then
                        topLevel = .Nodes.Add(kvp.Key)
                        topLevel.Name = kennung & kvp.Key
                        topLevel.Text = kvp.Key
                        hry = kvp.Value.hierarchy

                        Call buildProjectSubTreeView(topLevel, hry)
                    End If

                Next
            End If

        End With

    End Sub

    ''' <summary>
    ''' baut die Projekt-Struktur unterhalb der Projekt-Vorlage bzw des Projektes 
    ''' </summary>
    ''' <param name="topNode"></param>
    ''' <param name="hry"></param>
    ''' <remarks></remarks>
    '''
    Private Sub buildProjectSubTreeView(ByRef topNode As TreeNode, ByVal hry As clsHierarchy)
        Dim hryNode As clsHierarchyNode
        Dim anzChilds As Integer
        Dim childNameID As String

        Dim nodeLevel0 As TreeNode
        Dim nodeLevel1 As TreeNode

        Dim vorlElem As String = ""

        If hry.count >= 1 Then
            hryNode = hry.nodeItem(rootPhaseName)

            anzChilds = hryNode.childCount

            With topNode

                For i As Integer = 1 To anzChilds

                    childNameID = hryNode.getChild(i)
                    nodeLevel0 = .Nodes.Add(elemNameOfElemID(childNameID))
                    nodeLevel0.Name = childNameID

                    Dim pvName As String = getPVnameFromNode(topNode)
                    Dim tmpBreadcrumb As String = hry.getBreadCrumb(childNameID, CInt(hryStufen.Value))
                    Dim elemName As String = elemNameOfElemID(childNameID)
                    Dim element As String = calcHryFullname(elemName, tmpBreadcrumb)
                    Dim projElem As String = "[" & topNode.Name & "]" & element



                    If Projektvorlagen.Contains(topNode.Text) Then
                        Dim vproj As clsProjektvorlage = Projektvorlagen.getProject(topNode.Text)
                    End If

                    If ShowProjekte.contains(topNode.Text) Then

                        Dim hproj As clsProjekt = ShowProjekte.getProject(topNode.Text)
                        vorlElem = "[V:" & hproj.VorlagenName & "]" & element
                    End If

                    If elemIDIstMeilenstein(childNameID) Then
                        nodeLevel0.BackColor = System.Drawing.Color.Azure
                        If selectedMilestones.Contains(element) Or selectedMilestones.Contains(projElem) _
                            Or selectedMilestones.Contains(vorlElem) Or selectedMilestones.Contains(elemName) Then
                            nodeLevel0.Checked = True
                        End If
                    Else

                        If selectedPhases.Contains(element) Or selectedPhases.Contains(projElem) _
                            Or selectedPhases.Contains(vorlElem) Or selectedPhases.Contains(elemName) Then
                            nodeLevel0.Checked = True
                        End If
                    End If


                    If hry.nodeItem(childNameID).childCount > 0 Then
                        nodeLevel0.Tag = "P"


                        nodeLevel1 = nodeLevel0.Nodes.Add("-")
                        nodeLevel1.Tag = "P"
                    Else
                        nodeLevel0.Tag = "X"
                    End If


                Next

            End With

        Else
            ' nichts tun ...
        End If
    End Sub

    ''' <summary>
    ''' sammelt alle selektierten Phasen und Meilensteine des Strukturbaums treeView in selPhases und/oder
    ''' selMilestones auf, egal ob es ein Baum mit ProjektStuktur oder VorlagenStruktur ist
    ''' </summary>
    ''' <param name="treeView"></param>
    ''' <param name="selphases"></param>
    ''' <param name="selMilestones"></param>
    ''' <remarks></remarks>
    Private Sub pickupCheckedProjStructItems(ByVal treeView As TreeView, ByRef selphases As Collection, ByRef selMilestones As Collection)

        Dim anzahlKnoten As Integer
        Dim tmpNode As TreeNode
        Dim element As String
        Dim type As Integer = -1
        Dim pvName As String = ""

        ' löschen der aktuell selektierten Phasen und Meilensteine und neu einlesen vom Treeview
        selphases.Clear()
        selMilestones.Clear()

        anzahlKnoten = treeView.Nodes.Count

        With hryTreeView

            Dim hry As clsHierarchy = Nothing
            For px As Integer = 1 To anzahlKnoten

                tmpNode = .Nodes.Item(px - 1)

                ' jetzt muss das Projekt, die Projekt-Vorlage ermittelt werden 
                ' und daraus die Hierarchie 
                If tmpNode.Level = 0 Then
                    hry = getHryFromNode(tmpNode)
                    type = getTypeFromNode(tmpNode)
                    pvName = getPVnameFromNode(tmpNode)

                    '' RootPhasename in selectedPhases aufnehmen
                    If tmpNode.Checked And Not subNodesSelected(tmpNode) Then

                        Dim tmpBreadcrumb As String = hry.getBreadCrumb(rootPhaseName, CInt(hryStufen.Value))
                        Dim elemName As String = elemNameOfElemID(rootPhaseName)
                        Dim selElem As String = calcHryFullname(elemName, tmpBreadcrumb, getPVkennungFromNode(tmpNode))
                        If Not selphases.Contains(selElem) Then
                            selphases.Add(selElem, selElem)
                        End If

                    End If
                End If


                If tmpNode.Checked And Not IsNothing(hry) And tmpNode.Level > 0 Then
                    ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll 
                    Dim filterbyLevel0 As Boolean = topNodeIsSelected(tmpNode)
                    Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
                    Dim elemName As String = elemNameOfElemID(tmpNode.Name)
                    If filterbyLevel0 Then
                        element = calcHryFullname(elemName, tmpBreadcrumb, getPVkennungFromNode(tmpNode))
                    Else
                        element = calcHryFullname(elemName, tmpBreadcrumb)
                    End If


                    If elemIDIstMeilenstein(tmpNode.Name) Then
                        If Not selMilestones.Contains(element) Then
                            selMilestones.Add(element, element)
                        End If
                    Else
                        If Not selphases.Contains(element) Then
                            selphases.Add(element, element)
                        End If

                    End If

                End If

                If tmpNode.Nodes.Count > 0 Then
                    Call pickupCheckedItems(tmpNode, hry)
                End If

            Next

        End With
    End Sub


    ''' <summary>
    ''' gibt alle Namen von Knoten, die "gecheckt" sind, in der nameList zurück  
    ''' wird rekursiv aufgerufen 
    ''' Achtung: wenn es Endlos Zyklen gibt, dann ist hier eine Endlos-Schleife ! 
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Private Sub pickupCheckedItems(ByVal node As TreeNode, ByVal hry As clsHierarchy)

        Dim tmpNode As TreeNode
        Dim element As String

        If IsNothing(node) Then
            ' nichts tun
        Else

            Dim anzahlKnoten As Integer = node.Nodes.Count

            With node

                For px As Integer = 1 To anzahlKnoten

                    tmpNode = .Nodes.Item(px - 1)

                    If tmpNode.Checked Then


                        ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll 

                        Dim filterByLevel0 As Boolean = topNodeIsSelected(tmpNode)
                        Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
                        Dim elemName As String = elemNameOfElemID(tmpNode.Name)

                        If filterByLevel0 Then
                            element = calcHryFullname(elemName, tmpBreadcrumb, getPVkennungFromNode(tmpNode))
                        Else
                            element = calcHryFullname(elemName, tmpBreadcrumb)
                        End If

                        If elemIDIstMeilenstein(tmpNode.Name) Then
                            If Not selectedMilestones.Contains(element) Then
                                selectedMilestones.Add(element, element)
                            End If
                        Else
                            If Not selectedPhases.Contains(element) Then
                                selectedPhases.Add(element, element)
                            End If

                        End If

                    End If


                    If tmpNode.Nodes.Count > 0 Then
                        Call pickupCheckedItems(tmpNode, hry)
                    End If

                Next

            End With

        End If

    End Sub


    ''' <summary>
    ''' gibt alle Namen von Knoten, die "gecheckt" sind, in der selectedRoles-Liste zurück  
    ''' wird rekursiv aufgerufen 
    ''' Achtung: wenn es Endlos Zyklen gibt, dann ist hier eine Endlos-Schleife ! 
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Public Sub pickupCheckedRoleItems(ByVal node As TreeNode)
        Dim tmpNode As TreeNode
        Dim element As String

        If IsNothing(node) Then
            ' nichts tun
        Else

            Dim anzahlKnoten As Integer = node.Nodes.Count

            With node

                For px As Integer = 1 To anzahlKnoten

                    tmpNode = .Nodes.Item(px - 1)

                    If tmpNode.Checked Then

                        element = tmpNode.Text
                        If Not selectedRoles.Contains(element) Then
                            selectedRoles.Add(element, element)
                        End If


                    End If


                    If tmpNode.Nodes.Count > 0 Then
                        Call pickupCheckedRoleItems(tmpNode)
                    End If

                Next

            End With

        End If
    End Sub
    ''' <summary>
    ''' gibt alle Namen von Knoten, die "gecheckt" sind, in der selectedElems zurück   
    ''' </summary>
    ''' <param name="tree"></param>
    ''' <param name="selectedElems"></param>
    ''' <remarks></remarks>
    Private Sub pickupCheckedListItems(ByVal tree As TreeView, ByRef selectedElems As Collection)

        ' Merken welches die selektierten Phasen waren 
        selectedElems.Clear()
        For Each tN As TreeNode In tree.Nodes
            If tN.Checked Then
                If Not selectedElems.Contains(tN.Name) Then
                    selectedElems.Add(tN.Name, tN.Name)
                End If
            End If
        Next
    End Sub


    Private Sub hryTreeView_KeyPress(sender As Object, e As KeyPressEventArgs) Handles hryTreeView.KeyPress

        Dim initialNode As TreeNode = hryTreeView.SelectedNode
        Dim newMode As Boolean

        If e.KeyChar = "a" Or e.KeyChar = "A" Then
            ' Selektiere alle Unter-Knoten 
            With hryTreeView.SelectedNode
                .Expand()
                newMode = Not .Nodes.Item(0).Checked
                For i As Integer = 1 To .Nodes.Count
                    .Nodes.Item(i - 1).Checked = newMode
                Next
            End With

            'hryTreeView.SelectedNode = initialNode

        ElseIf e.KeyChar = "m" Or e.KeyChar = "M" Then
            ' selektiere/de-selektiere Meilensteine  
            With hryTreeView.SelectedNode
                .Expand()
                Dim ix As Integer = 1
                Dim fertig As Boolean = False
                While ix <= .Nodes.Count And Not fertig
                    If elemIDIstMeilenstein(.Nodes.Item(ix - 1).Name) Then
                        newMode = Not .Nodes.Item(ix - 1).Checked
                        For i As Integer = ix To .Nodes.Count
                            If elemIDIstMeilenstein(.Nodes.Item(i - 1).Name) Then
                                .Nodes.Item(i - 1).Checked = newMode
                            End If
                        Next
                        fertig = True
                    Else
                        ix = ix + 1
                    End If
                End While
            End With

            'hryTreeView.SelectedNode = initialNode

        ElseIf e.KeyChar = "p" Or e.KeyChar = "P" Then
            ' selektiere/de-selektiere Phasen
            With hryTreeView.SelectedNode
                .Expand()
                Dim ix As Integer = 1
                Dim fertig As Boolean = False
                While ix <= .Nodes.Count And Not fertig
                    If Not elemIDIstMeilenstein(.Nodes.Item(ix - 1).Name) Then
                        newMode = Not .Nodes.Item(ix - 1).Checked
                        For i As Integer = ix To .Nodes.Count
                            If Not elemIDIstMeilenstein(.Nodes.Item(i - 1).Name) Then
                                .Nodes.Item(i - 1).Checked = newMode
                            End If
                        Next
                        fertig = True
                    Else
                        ix = ix + 1
                    End If
                End While
            End With
        End If

        ' kennzeichnen, daß keine weitere Behandlung , insbesondere nicht die Standard-Behandlung notwendig ist 
        e.Handled = True
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

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

        ' Änderung von Thomas: 24.11.2016
        ' ''Dim vonDate As Date = getDateofColumn(showRangeLeft, False)
        ' ''Dim bisDate As Date = getDateofColumn(showRangeRight, True)

        ' ''If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
        ' ''    vonDate = getDateofColumn(showRangeLeft, False)
        ' ''    bisDate = getDateofColumn(showRangeRight, True)
        ' ''Else
        ' ''    vonDate = StartofCalendar
        ' ''    bisDate = StartofCalendar
        ' ''End If

        ' ''Try
        ' ''    currentReportProfil.calcRepVonBis(vonDate, bisDate)
        ' ''Catch ex As Exception
        ' ''    Throw New ArgumentException(ex.Message)
        ' ''End Try


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

                ' VorlagenDateiname eliminieren, ohne Pfadangaben im ReportProfil speichern
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
            Dim msgTxt As String = "Fehler " & ex.Message
            If awinSettings.englishLanguage Then
                msgTxt = "Error: " & ex.Message
            End If
            Call MsgBox(msgTxt)
        End Try

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusLabel.Text = CType(re.Result, String)

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        With Me.AbbrButton
            .Text = ""
            .Visible = False
            .Enabled = False
            .Left = .Left + 40
        End With


        Me.statusLabel.Text = "...done"
        Me.statusLabel.Visible = True
        Me.OKButton.Visible = True
        Me.OKButton.Enabled = True
        Me.repVorlagenDropbox.Enabled = True
        Me.Cursor = Cursors.Arrow



    End Sub
    ''' <summary>
    ''' uncheckt alle Selektionen im gesamten treeView
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''
    Private Sub SelectionReset_Click(sender As Object, e As EventArgs) Handles SelectionReset.Click


        Dim curNode As TreeNode
        With hryTreeView


            For i As Integer = 1 To .Nodes.Count
                curNode = .Nodes.Item(i - 1)
                If curNode.Checked Then
                    curNode.Checked = False
                End If
                If curNode.Nodes.Count > 0 Then
                    Call unCheck(curNode)
                End If
            Next


        End With

    End Sub

    ''' <summary>
    ''' setzt alle Knoten im TreeView auf unchecked
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Private Sub unCheck(ByRef node As TreeNode)
        Dim curNode As TreeNode

        With node

            For i As Integer = 1 To .Nodes.Count
                curNode = .Nodes.Item(i - 1)
                If curNode.Checked Then
                    curNode.Checked = False
                End If
                If curNode.Nodes.Count > 0 Then
                    Call unCheck(curNode)
                End If
            Next

        End With

    End Sub

    Private Sub SelectionSet_Click(sender As Object, e As EventArgs) Handles SelectionSet.Click

        Dim curNode As TreeNode
        With hryTreeView


            For i As Integer = 1 To .Nodes.Count
                curNode = .Nodes.Item(i - 1)
                If Not curNode.Checked Then
                    curNode.Checked = True
                End If
                If curNode.Nodes.Count > 0 Then
                    Call Check(curNode)
                End If
            Next


        End With
    End Sub

    ''' <summary>
    ''' setzt alle Knoten im TreeView auf checked
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Private Sub Check(ByRef node As TreeNode)
        Dim curNode As TreeNode

        With node

            For i As Integer = 1 To .Nodes.Count
                curNode = .Nodes.Item(i - 1)
                If Not curNode.Checked Then
                    curNode.Checked = True
                End If
                If curNode.Nodes.Count > 0 Then
                    Call Check(curNode)
                End If
            Next

        End With

    End Sub

    ''' <summary>
    ''' expandiert den kompletten Baum
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub expandCompletely_Click(sender As Object, e As EventArgs) Handles expandCompletely.Click



        With hryTreeView
            .ExpandAll()
        End With

    End Sub

    ''' <summary>
    ''' minimiert die dargestellte Baum-Struktur (collapse)  
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub collapseCompletely_Click(sender As Object, e As EventArgs) Handles collapseCompletely.Click

        With hryTreeView
            .CollapseAll()
        End With

    End Sub

    Private Sub filterDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles filterDropbox.SelectedIndexChanged

    End Sub

    Private Sub auswSpeichern_Click(sender As Object, e As EventArgs) Handles auswSpeichern.Click

        Dim anzahlKnoten As Integer
        Dim selectedNode As TreeNode
        Dim tmpNode As TreeNode
        Dim filterName As String = ""
        Dim element As String
        Dim type As Integer = -1
        Dim pvName As String = ""


        If Me.rdbNameList.Checked Then

            Dim lastFilter As String = "Last"
            appInstance.EnableEvents = False
            enableOnUpdate = False

            statusLabel.Text = ""

            anzahlKnoten = hryTreeView.Nodes.Count
            selectedNode = hryTreeView.SelectedNode

            ' hier muss jetzt noch der aktuelle rdb ausgelesen werden ..
            If Me.rdbPhases.Checked = True Then

                selectedPhases.Clear()
                With hryTreeView
                    For px As Integer = 1 To anzahlKnoten
                        tmpNode = .Nodes.Item(px - 1)
                        If tmpNode.Checked Then
                            ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll
                            If Not selectedPhases.Contains(tmpNode.Name) Then
                                selectedPhases.Add(tmpNode.Name, tmpNode.Name)
                            End If
                        End If
                    Next
                End With


                'selectedPhases.Clear()
                'For Each element As String In selNameListBox.Items
                '    If Not selectedPhases.Contains(element) Then
                '        selectedPhases.Add(element, element)
                '    End If
                'Next


            ElseIf Me.rdbMilestones.Checked = True Then

                selectedMilestones.Clear()
                With hryTreeView
                    For px As Integer = 1 To anzahlKnoten
                        tmpNode = .Nodes.Item(px - 1)
                        If tmpNode.Checked Then
                            ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll
                            If Not selectedMilestones.Contains(tmpNode.Name) Then
                                selectedMilestones.Add(tmpNode.Name, tmpNode.Name)
                            End If
                        End If
                    Next
                End With

       
            ElseIf rdbBU.Checked = True Then

                selectedBUs.Clear()
                'For Each element As String In selNameListBox.Items
                '    If Not selectedBUs.Contains(element) Then
                '        selectedBUs.Add(element, element)
                '    End If
                'Next

            ElseIf rdbTyp.Checked = True Then

                selectedTyps.Clear()
                'For Each element As String In selNameListBox.Items
                '    If Not selectedTyps.Contains(element) Then
                '        selectedTyps.Add(element, element)
                '    End If
                'Next
            End If
       

        ElseIf Me.rdbProjStruktProj.Checked Or Me.rdbProjStruktTyp.Checked Then

            ' Radiobutton Projekt-Struktur  wurde geklickt

            appInstance.EnableEvents = False
            enableOnUpdate = False

            statusLabel.Text = ""


            anzahlKnoten = hryTreeView.Nodes.Count
            selectedNode = hryTreeView.SelectedNode

            selectedPhases.Clear()
            selectedMilestones.Clear()

            With hryTreeView

                For px As Integer = 1 To anzahlKnoten

                    tmpNode = .Nodes.Item(px - 1)
                    Dim hry As clsHierarchy = getHryFromNode(tmpNode)

                    ' jetzt muss das Projekt, die Projekt-Vorlage ermittelt werden 
                    ' und daraus die Hierarchie 
                    If tmpNode.Level = 0 Then
                        hry = getHryFromNode(tmpNode)
                        type = getTypeFromNode(tmpNode)
                        pvName = getPVnameFromNode(tmpNode)
                        If tmpNode.Checked And Not subNodesSelected(tmpNode) Then

                            Dim tmpBreadcrumb As String = hry.getBreadCrumb(rootPhaseName, CInt(hryStufen.Value))
                            Dim elemName As String = elemNameOfElemID(rootPhaseName)
                            Dim selElem As String = calcHryFullname(elemName, tmpBreadcrumb, getPVkennungFromNode(tmpNode))
                            If Not selectedPhases.Contains(selElem) Then
                                selectedPhases.Add(selElem, selElem)
                            End If

                        End If
                    End If


                    If tmpNode.Checked And Not IsNothing(hry) And tmpNode.Level > 0 Then
                        ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll 
                        Dim filterbyLevel0 As Boolean = topNodeIsSelected(tmpNode)
                        Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
                        Dim elemName As String = elemNameOfElemID(tmpNode.Name)
                        If filterbyLevel0 Then
                            element = calcHryFullname(elemName, tmpBreadcrumb, getPVkennungFromNode(tmpNode))
                        Else
                            element = calcHryFullname(elemName, tmpBreadcrumb)
                        End If


                        'Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
                        'Dim elemName As String = elemNameOfElemID(tmpNode.Name)
                        'element = calcHryFullname(elemName, tmpBreadcrumb)

                        If elemIDIstMeilenstein(tmpNode.Name) Then
                            If Not selectedMilestones.Contains(element) Then
                                selectedMilestones.Add(element, element)
                            End If
                        Else
                            If Not selectedPhases.Contains(element) Then
                                selectedPhases.Add(element, element)
                            End If

                        End If

                    End If


                    If tmpNode.Nodes.Count > 0 Then
                        Call pickupCheckedItems(tmpNode, hry)
                    End If

                Next

            End With


        ElseIf rdbCosts.Checked = True Then

            selectedCosts.Clear()

            With hryTreeView
                For px As Integer = 1 To anzahlKnoten
                    tmpNode = .Nodes.Item(px - 1)
                    If tmpNode.Checked Then
                        ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll
                        If Not selectedCosts.Contains(tmpNode.Name) Then
                            selectedCosts.Add(tmpNode.Name, tmpNode.Name)
                        End If
                    End If
                Next
            End With


        ElseIf Me.rdbRoles.Checked = True Then

            anzahlKnoten = hryTreeView.Nodes.Count

            ' Merken welches die selektierten Rollen waren 
            ' Radiobutton Rollen wurde geklickt

            selectedRoles.Clear()

            With hryTreeView

                For px As Integer = 1 To anzahlKnoten

                    tmpNode = .Nodes.Item(px - 1)

                    If tmpNode.Checked Then

                        If Not selectedRoles.Contains(tmpNode.Text) Then
                            selectedRoles.Add(tmpNode.Text, tmpNode.Text)
                        End If

                    End If


                    If tmpNode.Nodes.Count > 0 Then
                        Call pickupCheckedRoleItems(tmpNode)
                    End If

                Next

            End With

        End If


        If Not (Me.menuOption = PTmenue.reportBHTC Or _
            Me.menuOption = PTmenue.reportMultiprojektTafel) Then

            If Me.menuOption = PTmenue.filterdefinieren Then

                filterName = filterDropbox.Text
                ' jetzt wird der Filter unter dem Namen filterName gespeichert ..
                Call storeFilter(filterName, menuOption, selectedBUs, selectedTyps, _
                                                       selectedPhases, selectedMilestones, _
                                                       selectedRoles, selectedCosts, False)
            ElseIf Me.menuOption = PTmenue.visualisieren Then

                If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0) And _
                    (selectedRoles.Count > 0 Or selectedCosts.Count > 0) Then
                    If awinSettings.englishLanguage Then
                        Call MsgBox("either phases/milestones or Roles/cost may be selected ...")
                    Else
                        Call MsgBox("es können nur entweder Phasen / Meilensteine oder Rollen oder Kosten angezeigt werden")
                    End If

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
            Dim lastfilter As String = "Last"
            Call storeFilter(lastfilter, menuOption, selectedBUs, selectedTyps, _
                                                       selectedPhases, selectedMilestones, _
                                                       selectedRoles, selectedCosts, True)

            ' geänderte Auswahl/Filterliste neu anzeigen
            If Not (Me.menuOption = PTmenue.filterdefinieren) Then
                filterDropbox.Items.Clear()
                For Each kvp As KeyValuePair(Of String, clsFilter) In selFilterDefinitions.Liste
                    filterDropbox.Items.Add(kvp.Key)
                Next

            End If


        ElseIf Me.menuOption = PTmenue.reportBHTC Or _
            Me.menuOption = PTmenue.reportMultiprojektTafel Then


            ' ''statusLabel.Text = ""


            ' ''anzahlKnoten = hryTreeView.Nodes.Count
            ' ''selectedNode = hryTreeView.SelectedNode

            ' ''selectedPhases.Clear()
            ' ''selectedMilestones.Clear()

            ' ''With hryTreeView

            ' ''    For px As Integer = 1 To anzahlKnoten

            ' ''        tmpNode = .Nodes.Item(px - 1)
            ' ''        Dim hry As clsHierarchy = getHryFromNode(tmpNode)

            ' ''        If tmpNode.Checked Then
            ' ''            ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll 

            ' ''            Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
            ' ''            Dim elemName As String = elemNameOfElemID(tmpNode.Name)
            ' ''            element = calcHryFullname(elemName, tmpBreadcrumb)

            ' ''            If elemIDIstMeilenstein(tmpNode.Name) Then
            ' ''                If Not selectedMilestones.Contains(element) Then
            ' ''                    selectedMilestones.Add(element, element)
            ' ''                End If
            ' ''            Else
            ' ''                If Not selectedPhases.Contains(element) Then
            ' ''                    selectedPhases.Add(element, element)
            ' ''                End If

            ' ''            End If

            ' ''        End If


            ' ''        If tmpNode.Nodes.Count > 0 Then
            ' ''            Call pickupCheckedItems(tmpNode, hry)
            ' ''        End If

            ' ''    Next

            ' ''End With


            Dim vorlagenDateiName As String
            If Not repProfil.isMpp Then
                vorlagenDateiName = awinPath & RepProjectVorOrdner & _
                                    "\" & repVorlagenDropbox.Text
            Else
                vorlagenDateiName = awinPath & RepPortfolioVorOrdner & _
                                   "\" & repVorlagenDropbox.Text
            End If


            ' Prüfen, ob die Datei überhaupt existirt 
            If repVorlagenDropbox.Text.Length = 0 Then

                ' Call MsgBox("bitte PPT Vorlage auswählen !")
                If awinSettings.englishLanguage Then
                    Me.statusLabel.Text = "please select a PPT template !"
                Else
                    Me.statusLabel.Text = "bitte PPT Vorlage auswählen !"
                End If

                Me.statusLabel.Visible = True

            ElseIf My.Computer.FileSystem.FileExists(vorlagenDateiName) Then

                ' pptTemplatename speichern
                repProfil.PPTTemplate = repVorlagenDropbox.Text

                If filterDropbox.Text.Length <> 0 Then

                    ' Name der ReportProfils speichern
                    repProfil.name = filterDropbox.Text

                    Call storeReportProfil(menuOption, selectedBUs, selectedTyps, _
                                                               selectedPhases, selectedMilestones, _
                                                               selectedRoles, selectedCosts, repProfil)


                Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox("please provide a name for the Report Profile")
                        Me.statusLabel.Text = "please provide a name for the Report Profile"
                    Else
                        Call MsgBox("Bitte geben Sie einen Namen für diese Report-Profil an")
                        Me.statusLabel.Text = "Bitte geben Sie einen Namen für diese Report-Profil an"
                    End If

                    Me.statusLabel.Visible = True
                End If



            Else

                'Call MsgBox("bitte PPT Vorlage auswählen !")
                If awinSettings.englishLanguage Then
                    Me.statusLabel.Text = "please select a PPT template !"
                Else
                    Me.statusLabel.Text = "bitte PPT Vorlage auswählen !"
                End If

                Me.statusLabel.Visible = True

            End If

            If awinSettings.englishLanguage Then
                Me.statusLabel.Text = "Report-Profile '" & repProfil.name & "' stored"
            Else
                Me.statusLabel.Text = "ReportProfil '" & repProfil.name & "' gespeichert"
            End If

            Me.statusLabel.Visible = True

        Else
            'Call MsgBox("nicht reportBHTC aber auch reportBHTC: also eigentlich nicht möglich")
        End If



    End Sub


    ''' <summary>
    ''' Laden der Auswahl, das sind vorallem Filter
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub auswLaden_Click(sender As Object, e As EventArgs) Handles auswLaden.Click

        Dim missingProjCollection As Collection

        If Me.menuOption = PTmenue.filterdefinieren Then

            Dim fName As String


            Try
                fName = filterDropbox.SelectedItem.ToString
                ' wird nicht benötigt: ur: 29.07.2015 Dim filter As clsFilter = filterDefinitions.retrieveFilter(fName)

                ' jetzt werden anhand des Filters "fName" die Collections gesetzt 
                Call retrieveSelections(fName, menuOption, selectedBUs, selectedTyps, _
                                        selectedPhases, selectedMilestones, _
                                        selectedRoles, selectedCosts)

                auswahl = selectionTyp(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)

                missingProjCollection = checkFilter(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)

                If auswahl = PTProjektType.nameList Then
                    Me.rdbNameList.Checked = True

                ElseIf auswahl = PTProjektType.projekt Then
                    Me.rdbProjStruktProj.Checked = True

                ElseIf auswahl = PTProjektType.vorlage Then
                    Me.rdbProjStruktTyp.Checked = True
                Else
                    Me.rdbProjStruktProj.Checked = True
                End If

                Call buildHryTreeViewNew(auswahl)

                If auswahl = PTProjektType.projekt Or auswahl = PTProjektType.vorlage Then

                    ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
                    If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                        hryTreeView.ExpandAll()
                    End If

                End If

                If selectedRoles.Count > 0 Then
                    Me.rdbRoles.Checked = True
                    Call buildTreeViewRolle()
                End If

                If selectedCosts.Count > 0 Then

                End If
                Cursor = Cursors.Default
            Catch ex As Exception

            End Try


        ElseIf Me.menuOption = PTmenue.reportBHTC Then

            'neuer Profil-Name in Klasse repProfil speichern
            repProfil.name = filterDropbox.SelectedItem.ToString



        Else


            Dim fName As String

            Try
                fName = filterDropbox.SelectedItem.ToString
                ' wird nicht benötigt: ur: 29.07.2015 Dim filter As clsFilter = filterDefinitions.retrieveFilter(fName)

                ' jetzt werden anhand des Filters "fName" die Collections gesetzt 
                Call retrieveSelections(fName, menuOption, selectedBUs, selectedTyps, _
                                        selectedPhases, selectedMilestones, _
                                        selectedRoles, selectedCosts)

                If selectedPhases.Count > 0 Or selectedMilestones.Count > 0 Then


                    auswahl = selectionTyp(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, _
                                           selectedRoles, selectedCosts)

                    missingProjCollection = checkFilter(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, _
                                                        selectedRoles, selectedCosts)

                    If auswahl = PTProjektType.nameList Then
                        Me.rdbNameList.Checked = True

                    ElseIf auswahl = PTProjektType.projekt Then
                        Me.rdbProjStruktProj.Checked = True

                    ElseIf auswahl = PTProjektType.vorlage Then
                        Me.rdbProjStruktTyp.Checked = True
                    Else
                        Me.rdbProjStruktProj.Checked = True
                    End If

                    Call buildHryTreeViewNew(auswahl)

                    If auswahl = PTProjektType.projekt Or auswahl = PTProjektType.vorlage Then

                        ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
                        If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                            hryTreeView.ExpandAll()
                        End If

                    End If
                End If

                If selectedRoles.Count > 0 Then
                    Me.rdbRoles.Checked = True
                    Call buildTreeViewRolle()
                End If

                If selectedCosts.Count > 0 Then

                End If

                Cursor = Cursors.Default
            Catch ex As Exception

            End Try


        End If

    End Sub



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        calledFrom = "Multiprojekt-Tafel"

        ' Add any initialization after the InitializeComponent() call.

    End Sub



    Private Sub BackgroundWorker3_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        ' ''Dim vorlagenDateiName As String = CType(e.Argument, String)

        ' ReportProfil ist nun in reportProfil komplett enthalten
        Dim reportProfil As clsReportAll = CType(e.Argument, clsReportAll)

        Dim zeilenhoehe As Double = 0.0     ' zeilenhöhe muss für alle Projekte gleich sein, daher mit übergeben
        Dim legendFontSize As Single = 0.0  ' FontSize der Legenden der Schriftgröße des Projektnamens angepasst



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
            .mppProjectsWithNoMPmayPass = reportProfil.projectsWithNoMPmayPass
        End With



        Try
            If Not reportProfil.isMpp Then


                Dim vorlagendateiname As String = awinPath & RepProjectVorOrdner & "\" & reportProfil.PPTTemplate
                If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

                    appInstance.EnableEvents = False
                    'appInstance.ScreenUpdating = False

                    If selectedProjekte.Count < 1 Then
                        Dim projname As String = reportProfil.Projects.ElementAt(0).Value
                        Dim hproj As clsProjekt = ShowProjekte.getProject(projname)
                        selectedProjekte.Add(hproj, False)
                    End If

                    Call createPPTReportFromProjects(vorlagendateiname, _
                                                     selectedPhases, selectedMilestones, _
                                                     selectedRoles, selectedCosts, _
                                                     selectedBUs, selectedTyps, _
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
                                                          selectedBUs, selectedTyps, True, _
                                                          worker, e)

                End If

            End If



        Catch ex As Exception
            Call MsgBox("Fehler: " & vbLf & ex.Message)
        End Try

        ' '' '' Report wird von Projekt hproj, das vor Aufruf des Formulars in hproj gespeichert wurde erzeugt

        '' ''showRangeLeft = getColumnOfDate(reportProfil.VonDate)
        '' ''showRangeRight = getColumnOfDate(reportProfil.BisDate)

        '' ''Try
        '' ''    Dim vorlagendateiname As String = awinPath & RepProjectVorOrdner & "\" & reportProfil.PPTTemplate

        '' ''    If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

        '' ''        Dim projname As String = reportProfil.Projects.ElementAt(0).Value

        '' ''        Dim hproj As clsProjekt = ShowProjekte.getProject(projname)

        '' ''        Call createPPTSlidesFromProject(hproj, vorlagendateiname, _
        '' ''                                        selectedPhases, selectedMilestones, _
        '' ''                                        selectedRoles, selectedCosts, _
        '' ''                                        selectedBUs, selectedTyps, True, _
        '' ''                                        True, zeilenhoehe, _
        '' ''                                        legendFontSize, _
        '' ''                                        worker, e)


        '' ''        ' ''Call createPPTReportFromProjects(vorlagenDateiName, _
        '' ''        ' ''                                   selectedPhases, selectedMilestones, _
        '' ''        ' ''                                   selectedRoles, selectedCosts, _
        '' ''        ' ''                                   selectedBUs, selectedTyps, _
        '' ''        ' ''                                   worker, e)
        '' ''    Else

        '' ''        ''Call createPPTSlidesFromConstellation(reportProfil.PPTTemplate, _
        '' ''        ''                                reportProfil.Phases, reportProfil.Milestones, _
        '' ''        ''                                reportProfil.Roles, reportProfil.Costs, _
        '' ''        ''                                reportProfil.BUs, reportProfil.Typs, True, _
        '' ''        ''                                worker, e)
        '' ''    End If


        '' ''Catch ex As Exception
        '' ''    Call MsgBox("Fehler: " & vbLf & ex.Message)
        '' ''End Try

    End Sub

    Private Sub BackgroundWorker3_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker3.ProgressChanged


        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusLabel.Text = CType(re.Result, String)

    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted

        With Me.AbbrButton
            .Text = ""
            .Visible = False
            .Enabled = False
            .Left = .Left + 40
        End With


        Me.statusLabel.Text = "...done"
        Me.statusLabel.Visible = True
        Me.OKButton.Visible = True
        Me.OKButton.Enabled = True
        Me.OKButton.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Arrow

        ' hier evt. noch schließen und Abspeichern des Reports von PPT

    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

        If (menuOption = PTmenue.reportBHTC Or _
            menuOption = PTmenue.reportMultiprojektTafel) Then

            If awinSettings.englishLanguage Then
                statusLabel.Text = "Report Creation cancelled"
            Else
                statusLabel.Text = "Berichterstellung wurde beendet"
            End If

            Try
                Me.BackgroundWorker3.CancelAsync()
            Catch ex As Exception

            End Try

        End If

    End Sub

    Private Sub rdbNameList_CheckedChanged(sender As Object, e As EventArgs) Handles rdbNameList.CheckedChanged


        'Dim i As Integer
        statusLabel.Text = ""
        filterBox.Enabled = True

        If Me.rdbNameList.Checked Then

            If selectedPhases.Count = 0 And _
               selectedMilestones.Count = 0 Then

                auswahl = PTProjektType.nameList
            Else
                auswahl = selectionTyp(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)
            End If

            Select Case auswahl

                Case PTProjektType.nameList

                    Me.rdbMilestones.Visible = True
                    Me.rdbPhases.Visible = True
                    Me.pictureMilestones.Visible = True
                    Me.picturePhasen.Visible = True
                    If Not (Me.rdbMilestones.Checked Or Me.rdbPhases.Checked) Then
                        Me.rdbPhases.Checked = True
                    End If
                    Me.rdbPhaseMilest.Visible = False
                    Me.picturePhaseMilest.Visible = False

                    Call buildHryTreeViewNew(auswahl)


                Case PTProjektType.vorlage

                    Me.rdbMilestones.Visible = False
                    Me.rdbPhases.Visible = False
                    Me.pictureMilestones.Visible = False
                    Me.picturePhasen.Visible = False
                    Me.rdbPhaseMilest.Visible = True
                    Me.picturePhaseMilest.Visible = True
                    If Not Me.rdbPhaseMilest.Checked Then
                        Me.rdbPhaseMilest.Checked = True
                    End If

                    Dim result As MsgBoxResult

                    If awinSettings.englishLanguage Then
                        result = MsgBox("You really want to deselect the elements?", MsgBoxStyle.YesNo, "Deselect the elements?")
                    Else
                        result = MsgBox("Sollen die ausgewählten Elemente wirklich de-selektiert werden?", MsgBoxStyle.YesNo, "Elemente wirklich deselektieren?")
                    End If

                    If result = MsgBoxResult.Yes Then

                        selectedPhases.Clear()
                        selectedMilestones.Clear()

                        Me.rdbMilestones.Visible = True
                        Me.rdbPhases.Visible = True
                        Me.pictureMilestones.Visible = True
                        Me.picturePhasen.Visible = True
                        If Not (Me.rdbMilestones.Checked Or Me.rdbPhases.Checked) Then
                            Me.rdbPhases.Checked = True
                        End If
                        Me.rdbPhaseMilest.Visible = False
                        Me.picturePhaseMilest.Visible = False
                        Me.rdbNameList.Checked = True

                        Call buildHryTreeViewNew(PTProjektType.nameList)

                    Else
                        Call buildHryTreeViewNew(PTProjektType.vorlage)
                        Me.rdbProjStruktTyp.Checked = True

                        'If awinSettings.englishLanguage Then
                        '    statusLabel.Text = "only as Project-Structur possible"
                        'Else
                        '    statusLabel.Text = "Elemente können nur in der Projekt-Struktur angezeigt werden"
                        'End If
                    End If



                Case PTProjektType.projekt

                    Me.rdbMilestones.Visible = False
                    Me.rdbPhases.Visible = False
                    Me.pictureMilestones.Visible = False
                    Me.picturePhasen.Visible = False
                    Me.rdbPhaseMilest.Visible = True
                    Me.picturePhaseMilest.Visible = True
                    If Not Me.rdbPhaseMilest.Checked Then
                        Me.rdbPhaseMilest.Checked = True
                    End If


                    Dim result As MsgBoxResult

                    If awinSettings.englishLanguage Then
                        result = MsgBox("You really want to deselect the elements?", MsgBoxStyle.YesNo, "Deselect the elements?")
                    Else
                        result = MsgBox("Sollen die ausgewählten Elemente wirklich de-selektiert werden?", MsgBoxStyle.YesNo, "Elemente wirklich deselektieren?")
                    End If

                    If result = MsgBoxResult.Yes Then

                        selectedPhases.Clear()
                        selectedMilestones.Clear()
                        Call buildHryTreeViewNew(PTProjektType.nameList)

                        Me.rdbMilestones.Visible = True
                        Me.rdbPhases.Visible = True
                        Me.pictureMilestones.Visible = True
                        Me.picturePhasen.Visible = True
                        If Not (Me.rdbMilestones.Checked Or Me.rdbPhases.Checked) Then
                            Me.rdbPhases.Checked = True
                        End If
                        Me.rdbPhaseMilest.Visible = False
                        Me.picturePhaseMilest.Visible = False
                        Me.rdbNameList.Checked = True
                    Else
                        Call buildHryTreeViewNew(PTProjektType.projekt)
                        Me.rdbProjStruktProj.Checked = True

                        If awinSettings.englishLanguage Then
                            statusLabel.Text = "only as Project-Structur possible"
                        Else
                            statusLabel.Text = "Elemente können nur in der Projekt-Struktur angezeigt werden"
                        End If
                    End If



                Case Else
                    selectedPhases.Clear()
                    selectedMilestones.Clear()
                    selectedBUs.Clear()
                    selectedTyps.Clear()
                    selectedRoles.Clear()
                    selectedCosts.Clear()

                    Me.rdbNameList.Checked = True
                    Me.rdbPhases.Checked = True

                    Call buildHryTreeViewNew(PTProjektType.nameList)

            End Select

            ' ''If lastAuswahl <> auswahl Then
            ' ''    'Call buildHryTreeViewNew(auswahl)
            ' ''    Call MsgBox("lastAuswahl=" & lastAuswahl.ToString & vbLf & "auswahl=" & auswahl.ToString)
            ' ''End If


        Else
            '??? lastAuswahl = PTProjektType.nameList

            auswahl = selectionTyp(selectedBUs, selectedTyps, _
                                   selectedPhases, selectedMilestones, selectedRoles, selectedCosts)
            If auswahl = PTProjektType.nameList Then

                If rdbPhases.Checked Then

                    ' Merken welches die selektierten Phasen waren 
                    Call pickupCheckedListItems(hryTreeView, selectedPhases)

                ElseIf rdbMilestones.Checked Then

                    ' Merken welches die selektierten Meilensteine waren 
                    Call pickupCheckedListItems(hryTreeView, selectedMilestones)

                End If

            Else
                ' nothing to do
            End If
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
        Dim listOfNames As New Collection
        Dim toplevel As TreeNode

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

        With hryTreeView
            .Nodes.Clear()
            'Dim kennung As String = "PH:"
            For Each ele As String In listOfNames
                If listOfNames.Count > 0 Then
                    toplevel = .Nodes.Add(ele)
                    toplevel.Name = ele
                    toplevel.Text = ele
                End If
            Next
        End With

        ' Filter Box Text setzen 
        filterBox.Text = ""

        ' jetzt prüfen, ob selected... bereits etwas enthält
        ' wenn ja, dann werden diese Items im Tree bereits selektiert markiert
        With hryTreeView
            Dim anzNodes As Integer = .Nodes.Count
            Dim tmpNode As TreeNode
            Dim passt As Boolean = True
            Dim bc As String = ""
            Dim eleName As String = ""

            For Each element As String In tmpCollection
                ' nachsehen ob in element 'P:' oder 'V:' enthalten sind
                Dim hstr1() As String = Split(element, "P:", )
                Dim hstr2() As String = Split(element, "V:", )
                passt = passt And (hstr1.Length = 1) And (hstr2.Length = 1)
            Next
            If passt Then
                For Each element As String In tmpCollection
                    For n As Integer = 1 To anzNodes
                        tmpNode = .Nodes.Item(n - 1)
                        Call splitHryFullnameTo2(element, eleName, bc, PTProjektType.nameList, "")
                        If tmpNode.Name = eleName Then
                            tmpNode.Checked = True
                        End If
                    Next
                Next
            Else
                Me.statusLabel.Text = "nur für Projekt-Sturktur (Projekt) geeignet"
            End If

        End With

    End Sub

    Private Sub rdbProjStruktProj_CheckedChanged(sender As Object, e As EventArgs) Handles rdbProjStruktProj.CheckedChanged

        Dim auswahl As Integer = -1

        If rdbProjStruktProj.Checked Then

            If Me.menuOption <> PTmenue.reportBHTC Then
                Me.rdbMilestones.Visible = False
                Me.rdbPhases.Visible = False
                Me.pictureMilestones.Visible = False
                Me.picturePhasen.Visible = False

                Me.rdbPhaseMilest.Visible = True
                Me.picturePhaseMilest.Visible = True
                If Not Me.rdbPhaseMilest.Checked Then
                    Me.rdbPhaseMilest.Checked = True
                End If
            End If
            


            ' clear Listbox1 
            If awinSettings.englishLanguage Then
                headerLine.Text = "Phases/Milestones"
            Else
                headerLine.Text = "Phasen/Meilensteine"
            End If

            filterBox.Visible = False
            filterBox.Text = ""

            If selectedPhases.Count = 0 And _
                selectedMilestones.Count = 0 Then

                auswahl = PTProjektType.projekt
            Else
                auswahl = selectionTyp(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)
            End If

            Select Case auswahl
                Case PTProjektType.nameList

                    Call buildHryTreeViewNew(PTProjektType.projekt)

                Case PTProjektType.vorlage

                    Call buildHryTreeViewNew(PTProjektType.projekt)

                Case PTProjektType.projekt

                    Call buildHryTreeViewNew(auswahl)

                Case Else
                    selectedPhases.Clear()
                    selectedMilestones.Clear()
                    selectedBUs.Clear()
                    selectedTyps.Clear()
                    selectedRoles.Clear()
                    selectedCosts.Clear()

                    Call buildHryTreeViewNew(PTProjektType.projekt)

            End Select

            ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
            If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                hryTreeView.ExpandAll()
            End If

        Else
            ' Merken der Projekte/Phasen und Meilensteine
            Call pickupCheckedProjStructItems(hryTreeView, selectedPhases, selectedMilestones)
        End If

    End Sub

    Private Sub rdbProjStruktTyp_CheckedChanged(sender As Object, e As EventArgs) Handles rdbProjStruktTyp.CheckedChanged

        Dim auswahl As Integer = -1

        If rdbProjStruktTyp.Checked Then

            Me.rdbMilestones.Visible = False
            Me.rdbPhases.Visible = False
            Me.pictureMilestones.Visible = False
            Me.picturePhasen.Visible = False
            Me.rdbPhaseMilest.Visible = True
            Me.picturePhaseMilest.Visible = True
            If Not Me.rdbPhaseMilest.Checked Then
                Me.rdbPhaseMilest.Checked = True
            End If

            ' clear Listbox1 
            If awinSettings.englishLanguage Then
                headerLine.Text = "Phases/Milestones"
            Else
                headerLine.Text = "Phasen/Meilensteine"
            End If

            filterBox.Visible = False
            filterBox.Text = ""

            If selectedPhases.Count = 0 And _
                 selectedMilestones.Count = 0 Then
                auswahl = PTProjektType.vorlage
            Else
                auswahl = selectionTyp(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)
            End If

            Select Case auswahl
                Case PTProjektType.nameList

                    Call buildHryTreeViewNew(PTProjektType.vorlage)


                Case PTProjektType.vorlage

                    Call buildHryTreeViewNew(auswahl)


                Case PTProjektType.projekt

                    Me.rdbProjStruktProj.Checked = True

                    'Call buildHryTreeViewNew(auswahl)
                    Dim result As MsgBoxResult

                    If awinSettings.englishLanguage Then
                        result = MsgBox("You really want to deselct the elements?", MsgBoxStyle.YesNo, "Deselect the elements?")
                    Else
                        result = MsgBox("Sollen die ausgewählten Elemente wirklich de-selektiert werden?", MsgBoxStyle.YesNo, "Elemente wirklich deselektieren?")
                    End If

                    If result = MsgBoxResult.Yes Then
                        selectedPhases.Clear()
                        selectedMilestones.Clear()
                       
                        Call buildHryTreeViewNew(PTProjektType.vorlage)

                        Me.rdbProjStruktTyp.Checked = True

                    Else
                        Call buildHryTreeViewNew(PTProjektType.projekt)
                    End If

                    'If awinSettings.englishLanguage Then
                    '    statusLabel.Text = "only as Project-Structur possible"
                    'Else
                    '    statusLabel.Text = "Elemente können nur in Projekt-Struktur angezeigt werden"
                    'End If


                Case Else

                    Call MsgBox("eigentlich Fehler !!!")
                    selectedPhases.Clear()
                    selectedMilestones.Clear()
                    selectedBUs.Clear()
                    selectedTyps.Clear()
                    selectedRoles.Clear()
                    selectedCosts.Clear()


                    Call buildHryTreeViewNew(PTProjektType.vorlage)

            End Select

            ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
            If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                hryTreeView.ExpandAll()
            End If
        Else

            Call pickupCheckedProjStructItems(hryTreeView, selectedPhases, selectedMilestones)

        End If

    End Sub

    ''' <summary>
    '''  Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub picturePhaseMilest_Click(sender As Object, e As EventArgs) Handles picturePhaseMilest.Click
        If Me.rdbPhaseMilest.Checked = False Then
            rdbPhaseMilest.Checked = True
        Else
            rdbPhaseMilest.Checked = False
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
            Me.rdbPhases.Checked = True
        Else
            Me.rdbPhases.Checked = False
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
            Me.rdbMilestones.Checked = True
        Else
            Me.rdbMilestones.Checked = False
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

    Private Sub rdbPhases_CheckedChanged(sender As Object, e As EventArgs) Handles rdbPhases.CheckedChanged

        'Dim i As Integer
        statusLabel.Text = ""
        filterBox.Enabled = True

        If Me.rdbNameList.Checked Then

            If Me.rdbPhases.Checked Then
                ' clear Listbox1 
                If awinSettings.englishLanguage Then
                    headerLine.Text = "Phases"
                Else
                    headerLine.Text = "Phasen"
                End If

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

                ' Merken welches die selektierten Phasen waren 
                Call pickupCheckedListItems(hryTreeView, selectedPhases)

                ' ''selectedPhases.Clear()
                ' ''For Each tN As TreeNode In hryTreeView.Nodes
                ' ''    If tN.Checked Then
                ' ''        selectedPhases.Add(tN.Name, tN.Name)
                ' ''    End If
                ' ''Next

            End If


        End If

    End Sub

    Private Sub rdbMilestones_CheckedChanged(sender As Object, e As EventArgs) Handles rdbMilestones.CheckedChanged

        statusLabel.Text = ""
        filterBox.Enabled = True

        If Me.rdbNameList.Checked Then

            If Me.rdbMilestones.Checked Then
                ' clear Listbox1 
                If awinSettings.englishLanguage Then
                    headerLine.Text = "Milestones"
                Else
                    headerLine.Text = "Meilensteine"
                End If

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
                Call pickupCheckedListItems(hryTreeView, selectedMilestones)

                ' ''selectedMilestones.Clear()
                ' ''For Each tN As TreeNode In hryTreeView.Nodes
                ' ''    If tN.Checked Then
                ' ''        selectedMilestones.Add(tN.Name, tN.Name)
                ' ''    End If
                ' ''Next

            End If


        End If

    End Sub

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

                With Me

                    'Anzeigen der erforderlichen Buttons
                    .rdbPhaseMilest.Visible = True
                    .rdbPhaseMilest.Checked = False
                    .picturePhaseMilest.Visible = True

                    ' Ausblenden der nicht clickbaren Buttons
                    .rdbNameList.Enabled = False
                    .rdbNameList.Visible = False
                    .rdbNameList.Checked = False

                    .rdbProjStruktProj.Enabled = False
                    .rdbProjStruktProj.Visible = False
                    .rdbProjStruktProj.Checked = False

                    .rdbProjStruktTyp.Enabled = False
                    .rdbProjStruktTyp.Visible = False
                    .rdbProjStruktTyp.Checked = False

                    .rdbPhases.Visible = False
                    .rdbPhases.Checked = False
                    .picturePhasen.Visible = False

                    .rdbMilestones.Visible = False
                    .rdbMilestones.Checked = False
                    .pictureMilestones.Visible = False

                End With


                ' clear Listbox1 
                If awinSettings.englishLanguage Then
                    headerLine.Text = "Roles/Names"
                Else
                    headerLine.Text = "Rollen/Namen"
                End If

                filterBox.Text = ""


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
                        allRoles = selectedProjekte.getRoleNames
                    ElseIf ShowProjekte.Count > 0 Then
                        allRoles = ShowProjekte.getRoleNames
                    Else
                        For i As Integer = 1 To RoleDefinitions.Count
                            Dim tmpName As String = RoleDefinitions.getRoledef(i).name
                            If Not allRoles.Contains(tmpName) Then
                                allRoles.Add(tmpName, tmpName)
                            End If
                        Next
                    End If

                End If

                Call buildTreeViewRolle()


            Else

                Dim anzahlKnoten As Integer = hryTreeView.Nodes.Count
                Dim tmpnode As TreeNode

                ' Merken welches die selektierten Rollen waren 
                ' Radiobutton Rollen wurde geklickt

                'selectedRoles.Clear()

                With hryTreeView

                    For px As Integer = 1 To anzahlKnoten

                        tmpnode = .Nodes.Item(px - 1)

                        If tmpnode.Checked Then

                            If Not selectedRoles.Contains(tmpnode.Text) Then
                                selectedRoles.Add(tmpnode.Text, tmpnode.Text)
                            End If

                        End If


                        If tmpnode.Nodes.Count > 0 Then
                            Call pickupCheckedRoleItems(tmpnode)
                        End If

                    Next

                End With


                End If
            End If
    End Sub

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


                With Me

                    'Anzeigen der erforderlichen Buttons
                    .rdbPhaseMilest.Visible = True
                    .rdbPhaseMilest.Checked = False
                    .picturePhaseMilest.Visible = True

                    ' Ausblenden der nicht clickbaren Buttons
                    .rdbNameList.Enabled = False
                    .rdbNameList.Visible = False
                    .rdbNameList.Checked = False

                    .rdbProjStruktProj.Enabled = False
                    .rdbProjStruktProj.Visible = False
                    .rdbProjStruktProj.Checked = False

                    .rdbProjStruktTyp.Enabled = False
                    .rdbProjStruktTyp.Visible = False
                    .rdbProjStruktTyp.Checked = False

                    .rdbPhases.Visible = False
                    .rdbPhases.Checked = False
                    .picturePhasen.Visible = False

                    .rdbMilestones.Visible = False
                    .rdbMilestones.Checked = False
                    .pictureMilestones.Visible = False

                End With

                ' clear Listbox1 
                If awinSettings.englishLanguage Then
                    headerLine.Text = "Cost Types"
                Else
                    headerLine.Text = "Kostenarten"
                End If

                filterBox.Text = ""

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

                ' Merken welches die selektierten Kosten waren 
                Call pickupCheckedListItems(hryTreeView, selectedCosts)

            End If
        End If
    End Sub

    Private Sub rdbBU_CheckedChanged(sender As Object, e As EventArgs) Handles rdbBU.CheckedChanged

    End Sub

    Private Sub rdbTyp_CheckedChanged(sender As Object, e As EventArgs) Handles rdbTyp.CheckedChanged

    End Sub

    Private Sub rdbPhaseMilest_CheckedChanged(sender As Object, e As EventArgs) Handles rdbPhaseMilest.CheckedChanged

        If Me.menuOption <> PTmenue.reportBHTC Then
            If rdbPhaseMilest.Checked Then

                ' Visibility der Buttons anpassen an die Auswahl
                With Me
                    .rdbNameList.Enabled = True
                    .rdbNameList.Visible = True
                    .rdbNameList.Checked = False

                    .rdbProjStruktProj.Enabled = True
                    .rdbProjStruktProj.Visible = True
                    '.rdbProjStruktProj.Checked = True

                    .rdbProjStruktTyp.Enabled = True
                    .rdbProjStruktTyp.Visible = True
                    ' .rdbProjStruktTyp.Checked = False

                    .rdbPhases.Visible = False
                    .rdbPhases.Checked = False
                    .picturePhasen.Visible = False

                    .rdbMilestones.Visible = False
                    .rdbMilestones.Checked = False
                    .pictureMilestones.Visible = False

                    '.rdbBU.Visible = False
                    '.pictureBU.Visible = False

                    '.rdbTyp.Visible = False
                    '.pictureTyp.Visible = False

                    '.rdbRoles.Visible = True
                    '.pictureRoles.Visible = True

                    '.rdbCosts.Visible = True
                    '.pictureCosts.Visible = True

                    .rdbPhaseMilest.Visible = True
                    .picturePhaseMilest.Visible = True

                End With


                ''ur: 20170905: nicht erforderlich
                ''auswahl = selectionTyp(selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)
                If Me.rdbProjStruktProj.Checked Then
                    Call buildHryTreeViewNew(PTProjektType.projekt)
                ElseIf Me.rdbProjStruktTyp.Checked Then
                    Call buildHryTreeViewNew(PTProjektType.vorlage)
                Else
                    Me.rdbProjStruktProj.Checked = True
                    'Call buildHryTreeViewNew(PTProjektType.projekt)
                End If


            Else

            End If
        End If
        
    End Sub

    Public Sub buildTreeViewRolle()


        Dim topLevelNode As TreeNode
        Dim checkProj As Boolean = False

        With hryTreeView

            .Nodes.Clear()
            .CheckBoxes = True


            ' alle Rollen in geladenen Projekte zeigen 

            If allRoles.Count > 0 Then
                Dim topNodes As List(Of Integer) = RoleDefinitions.getTopLevelNodeIDs


                For i = 0 To topNodes.Count - 1
                    Dim role As clsRollenDefinition = RoleDefinitions.getRoleDefByID(topNodes.ElementAt(i))
                    topLevelNode = .Nodes.Add(role.name)
                    topLevelNode.Name = role.UID.ToString
                    topLevelNode.Text = role.name
                    If selectedRoles.Contains(role.name) Then
                        topLevelNode.Checked = True
                    End If

                    Dim listOfChildIDs As New SortedList(Of Integer, String)
                    Try
                        listOfChildIDs = role.getSubRoleIDs
                    Catch ex As Exception

                    End Try

                    If listOfChildIDs.Count > 0 Then
                        For ii As Integer = 0 To listOfChildIDs.Count - 1
                            Call buildRoleSubTreeView(topLevelNode, listOfChildIDs.ElementAt(ii).Key)
                        Next
                    End If

                    'Call buildRoleSubTreeView(topLevel, roleHry.nodeItem(topNodes.Item(i)).childs)
                Next
            End If


        End With
    End Sub
    ''' <summary>
    ''' baut den Rollen-SubtreeView für die Rolle mit der ID roleUID auf. 
    ''' es wird ein neuer Knoten unterhalb des des parent-Knotens aufgebaut 
    ''' wenn dieser Child-Node seinerseits Kinder enthält, wird wiederum buildRoleSubTreeView aufgerufen ... 
    ''' </summary>
    ''' <param name="parentNode"></param>
    ''' <param name="roleUid"></param>
    ''' <remarks></remarks>
    Public Sub buildRoleSubTreeView(ByRef parentNode As TreeNode, ByVal roleUid As Integer)


        Dim currentRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(roleUid)
        Dim childIds As SortedList(Of Integer, String) = currentRole.getSubRoleIDs
        Dim doItAnyWay As Boolean = False
        Dim listOfroleNames As Collection = ShowProjekte.getRoleNames()

        ' wenn die vorhandenen Rollen als Kind oder Kindeskind von currentRole vorkommen, dann doItAnyWay
        If currentRole.isCombinedRole Then
            If currentRole.hasAnyOfThemAsChild(listOfroleNames) Then
                doItAnyWay = True
            End If
        End If

        If ShowProjekte.getRoleNames().Contains(currentRole.name) Or doItAnyWay Then

            Dim newNode As TreeNode
            With parentNode
                newNode = .Nodes.Add(currentRole.name)
                newNode.Name = roleUid.ToString
                newNode.Text = currentRole.name
                If selectedRoles.Contains(currentRole.name) Then
                    newNode.Checked = True
                End If
            End With

            For i = 0 To childIds.Count - 1

                Call buildRoleSubTreeView(newNode, childIds.ElementAt(i).Key)

            Next
        End If

    End Sub

    ' ''' <summary>
    ' ''' old _ deprecated
    ' ''' </summary>
    ' ''' <param name="topNode"></param>
    ' ''' <param name="nodeList"></param>
    ' ''' <remarks></remarks>
    'Public Sub buildRoleSubTreeView(ByRef topNode As TreeNode, ByRef nodeList As List(Of Integer))

    '    Dim anzChilds As Integer
    '    Dim child As clsRollenDefinition
    '    Dim childChildList As New List(Of Integer)
    '    Dim childChildExist As Boolean = False

    '    Dim nodeLevel0 As TreeNode
    '    Dim nodeLevel1 As TreeNode


    '    anzChilds = nodeList.Count

    '    With topNode

    '        For i As Integer = 0 To anzChilds - 1

    '            child = RoleDefinitions.getRoledef(nodeList(i))
    '            Dim childName As String = child.name
    '            Dim childID As Integer = child.UID

    '            If allRoles.Contains(childName) Then

    '                nodeLevel0 = .Nodes.Add(childName)
    '                nodeLevel0.Name = childID.ToString
    '                nodeLevel0.Text = childName


    '                If selectedRoles.Contains(childName) Then
    '                    nodeLevel0.Checked = True
    '                End If


    '                If RoleDefinitions.hierarchy.nodeItem(childID).childs.Count > 0 Then

    '                    childChildList = RoleDefinitions.hierarchy.nodeItem(childID).childs
    '                    ' prüfen, ob role childName ein Kind hat, das in allRoles enthalten ist
    '                    Dim k As Integer = 0
    '                    While (k <= childChildList.Count - 1) And Not childChildExist
    '                        childChildExist = childChildExist Or allRoles.Contains(RoleDefinitions.getRoledef(childChildList(k)).name)
    '                        k = k + 1
    '                    End While
    '                    If childChildExist Then
    '                        nodeLevel0.Tag = "P"


    '                        nodeLevel1 = nodeLevel0.Nodes.Add("-")
    '                        nodeLevel1.Tag = "P"
    '                    Else
    '                        nodeLevel0.Tag = "X"
    '                    End If

    '                End If
    '            End If

    '        Next

    '    End With


    'End Sub

    Private Sub hryTreeView_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles hryTreeView.AfterSelect

    End Sub

    Private Sub pictureTyp_Click(sender As Object, e As EventArgs) Handles pictureTyp.Click

    End Sub
End Class