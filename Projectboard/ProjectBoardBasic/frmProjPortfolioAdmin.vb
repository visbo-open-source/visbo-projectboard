Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports DBAccLayer
Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' wird verwendet um Portfolios zu definieren, Varianten zu aktivieren, Projekte und Varianten zu laden, zu aktivieren und zu löschen 
''' </summary>
''' <remarks></remarks>
Public Class frmProjPortfolioAdmin


    Private currentBrowserConstellation As New clsConstellation
    ' wenn Filter erstmalig aufgebaut wird , dann wird browserConstellationSav gemerkt ... 
    ' ur: 31.08.2017: Variable wird nun global defnieiert in Module.vb
    ' Private beforeFilterConstellation As clsConstellation = Nothing
    ' PlusMinus Saving 
    Private browserConstellationSavPM As clsConstellation = Nothing
    ' wenn aus der Datenbank schnell gelesen werden soll ..

    ' die  pvNameslistRaw enthält alle Varianten sowohl die Basis-Variante, die pfv-Variante und alle anderen 
    Private pvNamesListRaw As New SortedList(Of String, String)
    ' die pvNamesList enthält nur die BasisVariante+alle anderen bzw. die pfv-Variante plus alle anderen
    Private pvNamesList As New SortedList(Of String, String)

    Private quickList As Boolean
    Private lastIndexChecked As Integer = -1
    Private lastLevelChecked As Integer = -1

    Private earliestDate As Date
    ' tk 17.1.19 verbrät viel zu viel Speicherplatz 
    'Private projektHistorien As New clsProjektDBInfos
    Private stopRecursion As Boolean = False
    Private constellationName As String = ""

    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection
    Private selectedCosts As New Collection
    Private selectedRoles As New Collection
    Private selectedBUs As New Collection
    Private selectedTyps As New Collection

    Private toolTippsAreShowing As Integer

    ' tk 14.6.2020 wenn ActionKennung gleich selectPRojectasTemplate 
    Public selProjectAsTemplate As clsProjekt = Nothing

    ' um den Fehler im bestimmeNode zu umgehen 


    ''' <summary>
    ''' welche ToolTipps sollen im Browser Fenster gezeigt werden 
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum ptPPAtooltipps
        description = 0
        dependencies = 1
        scenarioReferences = 2
        protectedBy = 3
    End Enum

    ''' <summary>
    ''' Auflistung der ShowAttributes 
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum ptPPAshowAttributes
        all = 0
        show = 1
        noShow = 2
    End Enum

    ' wird an der aufrufenden Stelle gesetzt; steuert, was mit den ausgewählten ELementen geschieht
    Public aKtionskennung As Integer

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub

    Private Sub frmProjPortfolioAdmin_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

    End Sub

    Private Sub frmDefineEditPortfolio_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed


        frmCoord(PTfrm.eingabeProj, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.eingabeProj, PTpinfo.left) = Me.Left
        'projektHistorien.clear()


        If aKtionskennung = PTTvActions.chgInSession Then
            awinSettings.isChangePortfolioFrmActive = False
        End If

        If aKtionskennung = PTTvActions.delFromSession Or _
        aKtionskennung = PTTvActions.activateV Or _
        aKtionskennung = PTTvActions.loadPV Then

            ' 27.3.17 die letzte Editor Zusammenstellung nicht speichern; damit aucn im load nicht abfragen ... 
            'currentBrowserConstellation.constellationName = calcLastEditorScenarioName() ' wird damit jetzt auf Last & dbusername gesetzt 
            'projectConstellations.update(currentBrowserConstellation)

        End If


        ' Maus auf Normalmodus zurücksetzen
        'appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault

    End Sub

    Private Sub defineButtonVisibility()

        Dim versionenOffset As Integer = 20


        With Me

            ' bei Beginn immer disabled
            .deleteFilterIcon.Enabled = False

            ' Text des Versions-Feldes
            If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                .lblStandvom.Text = "Version at:"
            End If

            If aKtionskennung = PTTvActions.activateV Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Variante aktivieren"
                Else
                    .Text = "Activate Variant"
                End If

                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = False
                .SelectionReset.Visible = False

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False
                .OKButton.Visible = False

                '.lblVersionen1.Visible = False
                '.lblVersionen2.Visible = False
                '.versionsToKeep.Visible = False

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False

                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.loadProjectAsTemplate Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Projekt als Vorlage wählen"
                Else
                    .Text = "Select project to be template"
                End If

                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = False
                .SelectionReset.Visible = False

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False
                .OKButton.Visible = True
                If awinSettings.englishLanguage Then
                    .OKButton.Text = "Select as template"
                Else
                    .OKButton.Text = "als Vorlage wählen"
                End If


                '.lblVersionen1.Visible = False
                '.lblVersionen2.Visible = False
                '.versionsToKeep.Visible = False

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False

                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.chgInSession Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Portfolio "
                Else
                    .Text = "Portfolio "
                End If

                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = True
                .deleteFilterIcon.Visible = True
                If Not IsNothing(beforeFilterConstellation) Then

                    .deleteFilterIcon.Enabled = True

                    ' Das DeleteFilterIcon mit Bild versehen 
                    Me.deleteFilterIcon.Image = My.Resources.funnel_delete
                    Me.deleteFilterIcon.Enabled = True
                Else

                End If



                .dropboxScenarioNames.Visible = True

                .OKButton.Visible = True
                '.OKButton.Text = "Szenario speichern"

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    If storeToDBasWell.Checked Then
                        .OKButton.Text = "in Session und DB speichern"
                    Else
                        .OKButton.Text = "in Session speichern"
                    End If
                Else
                    If storeToDBasWell.Checked Then
                        .OKButton.Text = "Save to Session and DB"
                    Else
                        .OKButton.Text = "Save to Session"
                    End If
                End If


                Dim testName As String = .OKButton.Name

                '.lblVersionen1.Visible = False
                '.lblVersionen2.Visible = False
                '.versionsToKeep.Visible = False

                onlyActive.Visible = True
                onlyInactive.Visible = True
                backToInit.Visible = False

                If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                    storeToDBasWell.Visible = True
                Else
                    storeToDBasWell.Visible = False
                End If


                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.deleteV Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Variante löschen"
                Else
                    .Text = "Delete Variant"
                End If

                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    OKButton.Text = "aus Session Löschen"
                Else
                    .OKButton.Text = "Delete from Session"
                End If

                .OKButton.Visible = True


                '.lblVersionen1.Visible = False
                '.lblVersionen2.Visible = False
                '.versionsToKeep.Visible = False

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False
                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.delFromDB Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Projekte, Varianten bzw. Snapshots in der Datenbank löschen"
                Else
                    .Text = "Delete projects, variants, timestamps from DB"
                End If


                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False

                .OKButton.Visible = True
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .OKButton.Text = "in DB Löschen"
                Else
                    .OKButton.Text = "Delete from DB"
                End If


                '.lblVersionen1.Visible = False
                '.lblVersionen2.Visible = False
                '.versionsToKeep.Visible = False

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False
                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.delAllExceptFromDB Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Alle Zeitstempel löschen ausser ... "
                Else
                    .Text = "Delete all timestamps except ..."
                End If


                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False

                .OKButton.Visible = True
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .OKButton.Text = "in DB Löschen"
                Else
                    .OKButton.Text = "Delete from DB"
                End If


                '.lblVersionen1.Visible = True
                '.lblVersionen2.Visible = True

                'If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                '    lblVersionen1.Text = "delete all except"
                '    lblVersionen2.Text = "different versions"
                'End If

                '.versionsToKeep.Visible = True
                '.versionsToKeep.Value = 3
                '.lblVersionen1.Top = .lblVersionen1.Top + versionenOffset
                '.lblVersionen2.Top = .lblVersionen2.Top + versionenOffset
                '.versionsToKeep.Top = .versionsToKeep.Top + versionenOffset
                .dropboxScenarioNames.Top = .dropboxScenarioNames.Top - versionenOffset

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False
                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.delFromSession Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Projekte, Varianten aus der Session löschen"
                Else
                    .Text = "Delete projects, variants from Session"
                End If

                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False

                .OKButton.Visible = True
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .OKButton.Text = "Löschen"
                Else
                    .OKButton.Text = "Delete"
                End If


                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False
                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.loadPV Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Projekte und Varianten in die Session laden "
                Else
                    .Text = "Load projects and variants to the session "
                End If

                .requiredDate.Visible = True
                .lblStandvom.Visible = True

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = True
                .deleteFilterIcon.Visible = True

                .dropboxScenarioNames.Visible = False


                .OKButton.Visible = True
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .OKButton.Text = "Laden"
                Else
                    .OKButton.Text = "Load"
                End If


                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False
                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.loadPVinPPT Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "1 Projekt bzw. Projekt-Variante wählen"
                Else
                    .Text = "Select one project / project-variant"
                End If

                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = False
                .SelectionReset.Visible = False

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = True
                .deleteFilterIcon.Visible = True

                .dropboxScenarioNames.Visible = False


                .OKButton.Visible = True

                If awinSettings.englishLanguage Then
                    .OKButton.Text = "Auswählen"
                Else
                    .OKButton.Text = "Select"
                End If

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False
                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.loadMultiPVinPPT Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Projekte bzw. Projekt-Varianten wählen"
                Else
                    .Text = "Select one or more projects / project-variants"
                End If

                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = False
                .SelectionReset.Visible = False

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = True
                .deleteFilterIcon.Visible = True

                .dropboxScenarioNames.Visible = False


                .OKButton.Visible = True

                If awinSettings.englishLanguage Then
                    .OKButton.Text = "Auswählen"
                Else
                    .OKButton.Text = "Select"
                End If

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False
                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.loadPVS Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Projekte und Varianten in die Session laden "
                Else
                    .Text = "Load projects and variants to the session "
                End If

                .requiredDate.Visible = True
                .lblStandvom.Visible = True

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = True
                .deleteFilterIcon.Visible = True

                .dropboxScenarioNames.Visible = False


                .OKButton.Visible = True
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .OKButton.Text = "Laden"
                Else
                    .OKButton.Text = "Load"
                End If

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False
                chkbxPermanent.Visible = False

            ElseIf aKtionskennung = PTTvActions.setWriteProtection Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Schreibschutz für Projekt-Varianten"
                Else
                    .Text = "Write Protections for Project Variants"
                End If

                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False


                .OKButton.Visible = False
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .OKButton.Text = ""
                Else
                    .OKButton.Text = ""
                End If

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False
                chkbxPermanent.Visible = False


            End If

        End With


    End Sub

    Private Sub frmProjPortfolioAdmin_InputLanguageChanging(sender As Object, e As InputLanguageChangingEventArgs) Handles Me.InputLanguageChanging

    End Sub


    Private Sub frmDefineEditPortfolio_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim err As New clsErrorCodeMsg

        ' erstmal den WaitCursor zeigen ... 
        Me.Cursor = Cursors.Default
        lastIndexChecked = -1

        '' den hilfetext setzen ...
        'If awinSettings.englishLanguage Then
        '    Me.portfolioBrowserHelp.SetHelpString(TreeViewProjekte, "HelpMessage TreeView" & vbLf &
        '                                          "das ist die 1.Zeile " & vbLf &
        '                                          "das ist die zweite Zeile")
        '    Me.portfolioBrowserHelp.SetShowHelp(TreeViewProjekte, True)
        'End If

        If frmCoord(PTfrm.eingabeProj, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.top))
        End If

        If frmCoord(PTfrm.eingabeProj, PTpinfo.left) > 0 Then
            Me.Left = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.left))
        End If

        ' was sollen die ToolTipps zeigen ? 
        If aKtionskennung = PTTvActions.setWriteProtection Then
            toolTippsAreShowing = ptPPAtooltipps.protectedBy
        ElseIf aKtionskennung = PTTvActions.delFromDB Then
            toolTippsAreShowing = ptPPAtooltipps.scenarioReferences
        Else
            toolTippsAreShowing = ptPPAtooltipps.description
        End If


        ' bestimmen, ob es sich um quicklist handelt ...
        If aKtionskennung = PTTvActions.loadPV Or
            aKtionskennung = PTTvActions.loadPVInPPT Or
            aKtionskennung = PTTvActions.loadMultiPVInPPT Or
            aKtionskennung = PTTvActions.delFromDB Or
            aKtionskennung = PTTvActions.delAllExceptFromDB Then
            quickList = True
        Else
            quickList = False
        End If


        ' je nachdem, wie die Aktionskennung ist: setzen der Button Visibilitäten 
        Call defineButtonVisibility()

        ' wie heisst das aktuelle Szenario ? 
        If aKtionskennung <> PTTvActions.loadPVInPPT And aKtionskennung <> PTTvActions.loadMultiPVInPPT Then
            Me.Text = Me.Text & ": " & currentConstellationName
        End If


        ' neuer Ansatz
        If Not quickList Then
            If projectConstellations.Contains(currentConstellationName) And AlleProjekte.Count > 0 Then
                ' tk 23.2.19 - wenn eine Constellation geladen wird und als Summary angezeigt werden soll, dann war hier bisher die Liste de rProjekte des Portfolios drin ..!? 
                'currentBrowserConstellation = projectConstellations.getConstellation(currentConstellationName).copy()
                currentBrowserConstellation = currentSessionConstellation.copy()

            ElseIf AlleProjekte.Count > 0 Then

                currentBrowserConstellation = currentSessionConstellation.copy()

            End If
        End If

        ' jetzt die Korrektheitsprüfung ...
        If awinSettings.visboDebug And aKtionskennung = PTTvActions.chgInSession Then
            currentBrowserConstellation.checkAndCorrectYourself()
        End If


        ' jetzt die vorkommenden Timestamps auslesen 
        ' aber nicht bei allen Aktionskennungen 

        If aKtionskennung = PTTvActions.chgInSession Or
            aKtionskennung = PTTvActions.delFromSession Or
            aKtionskennung = PTTvActions.deleteV Or
            aKtionskennung = PTTvActions.activateV Then

        Else

            Try

                Dim tCollection As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveZeitstempelFromDB()

                If tCollection.Count >= 1 Then
                    earliestDate = tCollection.Item(tCollection.Count).date.addhours(23).addminutes(59)
                Else
                    earliestDate = Date.Now.Date.AddHours(23).AddMinutes(59)
                End If


            Catch ex As Exception

            End Try

            ' jetzt ist dropBoxTimeStamps.selecteditem = Nothing ..
        End If


        ' Maus auf Wartemodus setzen
        'appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

        If aKtionskennung = PTTvActions.chgInSession Then

            Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, err)

            For Each kvp1 As KeyValuePair(Of String, String) In dbPortfolioNames
                dropboxScenarioNames.Items.Add(kvp1.Key)
            Next

            For Each kvp2 As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste
                If kvp2.Key <> "Start" Then
                    If Not dbPortfolioNames.ContainsKey(kvp2.Key) Then
                        dropboxScenarioNames.Items.Add(kvp2.Key)
                    End If

                End If
            Next
        Else
            ' nichts weiter tun 
            ' die Filter-NAmen müssen aktuell nicht ausgelesen werden 
        End If



        stopRecursion = True
        Dim storedAtOrBefore As Date = Date.Now.Date.AddHours(23).AddMinutes(59)
        requiredDate.Value = storedAtOrBefore


        ' hier wird jetzt die Browser Gesamt-Liste bestimmt  
        If aKtionskennung = PTTvActions.loadPV Or
            aKtionskennung = PTTvActions.loadPVInPPT Or
            aKtionskennung = PTTvActions.loadMultiPVInPPT Or
            aKtionskennung = PTTvActions.delFromDB Or
            aKtionskennung = PTTvActions.delAllExceptFromDB Then

            ' hier wird jetzt die Raw-List geholt, d.h die enthält neben allen anderen Varianten auch die Basis- und Vorgabe-(PFV)Variante 

            If aKtionskennung = PTTvActions.delFromDB Then
                pvNamesListRaw = buildPvNamesList(storedAtOrBefore, True)
            Else
                pvNamesListRaw = buildPvNamesList(storedAtOrBefore, False)
            End If

            If awinSettings.loadPFV Or (awinSettings.filterPFV And aKtionskennung = PTTvActions.loadPV) Then
                pvNamesList = reduceRawListTo(pvNamesListRaw, True)
            ElseIf aKtionskennung = PTTvActions.loadPVInPPT Or aktionskennung = PTTvActions.loadMultiPVInPPT Then
                pvNamesList = pvNamesListRaw
            Else
                pvNamesList = reduceRawListTo(pvNamesListRaw, False)
            End If

            quickList = True

            If pvNamesList.Count = 0 Then
                Dim txtmsg As String
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    txtmsg = "keine Projekte in der Datenbank vorhanden ..."
                Else
                    txtmsg = "there are no projects in database ..."
                End If
                Call MsgBox(txtmsg)
            End If

        ElseIf aKtionskennung = PTTvActions.setWriteProtection Then

            If AlleProjekte.Count > 0 Then
                pvNamesList.Clear()
            Else

                ' hier wird jetzt die Raw-List geholt, d.h die enthält neben allen anderen Varianten auch die Basis- und Vorgabe-(PFV)Variante 
                pvNamesListRaw = buildPvNamesList(storedAtOrBefore)
                pvNamesList = reduceRawListTo(pvNamesListRaw, awinSettings.loadPFV)
                quickList = True

                If pvNamesList.Count = 0 Then
                    Dim txtmsg As String
                    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                        txtmsg = "keine Projekte in der Datenbank vorhanden ..."
                    Else
                        txtmsg = "there are no projects in database ..."
                    End If
                    Call MsgBox(txtmsg)
                End If
            End If
        End If

        stopRecursion = True
        Call updateTreeview(currentBrowserConstellation, pvNamesList, aKtionskennung, quickList)
        'Call buildTreeview(projektHistorien, TreeViewProjekte, pvNamesList, currentBrowserConstellation, _
        '                   aKtionskennung, quickList, _
        '                   storedAtOrBefore)



        stopRecursion = False

        If AlleProjekte.liste.Count < 1 And pvNamesList.Count < 1 Then
            'If browserAlleProjekte.liste.Count < 1 And pvNamesList.Count < 1 Then
            ' nichts in der Datenbank ...
            DialogResult = Windows.Forms.DialogResult.OK
        End If

        ' Maus auf Normalmodus zurücksetzen
        'appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Me.Cursor = Cursors.Arrow

        ' Fokus auf was unverdächtiges setzen 
        dropboxScenarioNames.Focus()




    End Sub

    ''' <summary>
    ''' die Rawliste enthält alle Varianten, inkl der Basis- wie der Vorgabe-Varianten
    ''' mit dieser Funktion wird die Liste entweder bereinigt um die Basis- oder die PFV-Variante 
    ''' </summary>
    ''' <param name="completeList"></param>
    ''' <param name="showPFV"></param>
    ''' <returns></returns>
    Private Function reduceRawListTo(ByVal completeList As SortedList(Of String, String), ByVal showPFV As Boolean) As SortedList(Of String, String)
        Dim tmpResult As New SortedList(Of String, String)
        Dim ausschluss As String = ""

        If showPFV Then
            ausschluss = ""
        Else
            ausschluss = ptVariantFixNames.pfv.ToString
        End If

        For Each kvp As KeyValuePair(Of String, String) In completeList

            Dim vName As String = getVariantnameFromKey(kvp.Key)

            If Not vName = ausschluss Then
                tmpResult.Add(kvp.Key, kvp.Value)
            End If

        Next

        reduceRawListTo = tmpResult
    End Function


    Private Sub TreeViewProjekte_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles TreeViewProjekte.AfterCheck

        Dim node As TreeNode
        Dim schluessel As String = ""
        'Dim selCollection As SortedList(Of Date, String)
        'Dim timeStamp As Date
        Dim treeLevel As Integer
        'Dim i As Integer, j As Integer
        'Dim childNode As TreeNode
        'Dim parentNode As TreeNode
        Dim currentIndex As Integer
        Dim shiftKeywasPressed As Boolean = False

        Dim considerDependencies As Boolean
        If allDependencies.projectCount > 0 Then
            considerDependencies = True
        Else
            considerDependencies = False
        End If

        ' Andernfalls wird die Check Routine endlos aufgerufen ...
        If stopRecursion Then
            Exit Sub
        End If

        node = e.Node
        treeLevel = node.Level
        currentIndex = node.Index


        ' das Szenario wird im Falle activateV und chgInSession verändert ... 
        ' das muss hier vermerkt werden ...
        If aKtionskennung = PTTvActions.chgInSession Or
            aKtionskennung = PTTvActions.activateV Then
            If currentConstellationName <> calcLastSessionScenarioName() Then
                currentConstellationName = calcLastSessionScenarioName()
            End If

            Dim preText As String
            If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                preText = "Portfolio "
            Else
                preText = "Portfolio "
            End If

            Me.Text = preText & currentConstellationName
        End If

        If My.Computer.Keyboard.ShiftKeyDown Then
            shiftKeywasPressed = True
        End If

        ' hier wird jetzt sichergestellt, daß nur die nach der aktuellen Aktion gültigen Checks gesetzt werden können
        ' vor allem muss überall dort, wo das Szenario mit diesem Check verändert wird, das currentBrowserSzenario geupdated werden ...
        ' mit Click in TreeView wird verändert: Activate Variant, ChgInSession 

        Dim checkMode As Boolean = node.Checked

        If aKtionskennung = PTTvActions.delFromDB Or
            aKtionskennung = PTTvActions.delAllExceptFromDB Or
            aKtionskennung = PTTvActions.loadPV Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 


                    ' prüfen, ob Mauskey gedrückt war ...
                    If shiftKeywasPressed Then
                        If validMultiSelection(lastIndexChecked, currentIndex) Then
                            If lastIndexChecked < 0 Then
                                lastIndexChecked = 0
                            End If

                            Dim lb As Integer = lastIndexChecked
                            Dim ub As Integer = currentIndex
                            If lastIndexChecked > currentIndex Then
                                lb = currentIndex
                                ub = lastIndexChecked
                            End If

                            For h = lb To ub
                                Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(h)

                                If tmpNode.Level = treeLevel Then
                                    ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                                    tmpNode.Checked = checkMode
                                    Call doAfterCheckAction(aKtionskennung, treeLevel, tmpNode, considerDependencies)
                                End If

                            Next
                        Else
                            Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
                        End If

                    Else

                        Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)

                    End If




                Case 1 ' Variante ist selektiert / nicht selektiert

                    ' prüfen, ob Mauskey gedrückt war ...
                    If shiftKeywasPressed Then
                        If validMultiSelection(lastIndexChecked, currentIndex) Then
                            If lastIndexChecked < 0 Then
                                lastIndexChecked = 0
                            End If

                            Dim lb As Integer = lastIndexChecked
                            Dim ub As Integer = currentIndex
                            If lastIndexChecked > currentIndex Then
                                lb = currentIndex
                                ub = lastIndexChecked
                            End If

                            For h = lb To ub
                                Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(h)

                                If tmpNode.Level = treeLevel Then
                                    ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                                    tmpNode.Checked = checkMode
                                    Call doAfterCheckAction(aKtionskennung, treeLevel, tmpNode, considerDependencies)
                                End If

                            Next
                        Else
                            Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
                        End If

                    Else
                        Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
                    End If

                Case 2 ' Snapshot ist selektiert / nicht selektiert 

                    ' prüfen, ob Mauskey gedrückt war ...
                    If shiftKeywasPressed Then
                        If validMultiSelection(lastIndexChecked, currentIndex) Then
                            If lastIndexChecked < 0 Then
                                lastIndexChecked = 0
                            End If

                            Dim lb As Integer = lastIndexChecked
                            Dim ub As Integer = currentIndex
                            If lastIndexChecked > currentIndex Then
                                lb = currentIndex
                                ub = lastIndexChecked
                            End If

                            For h = lb To ub
                                Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(h)

                                If tmpNode.Level = treeLevel Then
                                    ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                                    tmpNode.Checked = checkMode
                                    Call doAfterCheckAction(aKtionskennung, treeLevel, tmpNode, considerDependencies)
                                End If

                            Next
                        Else
                            Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
                        End If

                    Else

                        Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)

                    End If

            End Select

            stopRecursion = False

        ElseIf aKtionskennung = PTTvActions.loadPVInPPT Or aKtionskennung = PTTvActions.loadMultiPVInPPT Then

            stopRecursion = True
            Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
            stopRecursion = False

        ElseIf aKtionskennung = PTTvActions.delFromSession Or
              aKtionskennung = PTTvActions.deleteV Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    ' prüfen, ob Mauskey gedrückt war ...
                    If shiftKeywasPressed Then

                        If validMultiSelection(lastIndexChecked, currentIndex) Then
                            If lastIndexChecked < 0 Then
                                lastIndexChecked = 0
                            End If

                            Dim lb As Integer = lastIndexChecked
                            Dim ub As Integer = currentIndex
                            If lastIndexChecked > currentIndex Then
                                lb = currentIndex
                                ub = lastIndexChecked
                            End If

                            For h = lb To ub
                                Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(h)

                                If tmpNode.Level = treeLevel Then
                                    ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                                    tmpNode.Checked = checkMode
                                    Call doAfterCheckAction(aKtionskennung, treeLevel, tmpNode, considerDependencies)
                                End If

                            Next
                        Else
                            Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
                        End If


                    Else

                        Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)

                    End If

                Case 1 ' Variante ist selektiert / nicht selektiert

                    ' prüfen, ob Mauskey gedrückt war ...
                    If shiftKeywasPressed Then
                        If validMultiSelection(lastIndexChecked, currentIndex) Then
                            If lastIndexChecked < 0 Then
                                lastIndexChecked = 0
                            End If

                            Dim lb As Integer = lastIndexChecked
                            Dim ub As Integer = currentIndex
                            If lastIndexChecked > currentIndex Then
                                lb = currentIndex
                                ub = lastIndexChecked
                            End If

                            For h = lb To ub
                                Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(h)

                                If tmpNode.Level = treeLevel Then
                                    ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                                    tmpNode.Checked = checkMode
                                    Call doAfterCheckAction(aKtionskennung, treeLevel, tmpNode, considerDependencies)
                                End If

                            Next
                        Else
                            Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
                        End If

                    Else

                        Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)

                    End If


            End Select

            stopRecursion = False

        ElseIf aKtionskennung = PTTvActions.setWriteProtection Then

            stopRecursion = True

            If Not noDB Then

                Select Case treeLevel

                    Case 0 ' Projekt ist selektiert / nicht selektiert 

                        ' prüfen, ob Mauskey gedrückt war ...
                        If shiftKeywasPressed Then
                            If validMultiSelection(lastIndexChecked, currentIndex) Then
                                If lastIndexChecked < 0 Then
                                    lastIndexChecked = 0
                                End If

                                Dim lb As Integer = lastIndexChecked
                                Dim ub As Integer = currentIndex
                                If lastIndexChecked > currentIndex Then
                                    lb = currentIndex
                                    ub = lastIndexChecked
                                End If

                                For h = lb To ub
                                    Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(h)

                                    If tmpNode.Level = treeLevel Then
                                        ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                                        tmpNode.Checked = checkMode
                                        Call doAfterCheckAction(aKtionskennung, treeLevel, tmpNode, considerDependencies)
                                    End If

                                Next
                            Else
                                Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
                            End If

                        Else

                            Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)

                        End If


                    Case 1 ' Variante ist selektiert / nicht selektiert

                        ' prüfen, ob Mauskey gedrückt war ...
                        If shiftKeywasPressed Then
                            If validMultiSelection(lastIndexChecked, currentIndex) Then
                                If lastIndexChecked < 0 Then
                                    lastIndexChecked = 0
                                End If

                                Dim lb As Integer = lastIndexChecked
                                Dim ub As Integer = currentIndex
                                If lastIndexChecked > currentIndex Then
                                    lb = currentIndex
                                    ub = lastIndexChecked
                                End If

                                For h = lb To ub
                                    Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(h)

                                    If tmpNode.Level = treeLevel Then
                                        ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                                        tmpNode.Checked = checkMode
                                        Call doAfterCheckAction(aKtionskennung, treeLevel, tmpNode, considerDependencies)
                                    End If

                                Next
                            Else
                                Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
                            End If

                        Else

                            Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)

                        End If


                End Select

            Else
                ' zurücknehmen
                node.Checked = Not node.Checked
            End If

            stopRecursion = False

        ElseIf aKtionskennung = PTTvActions.activateV Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    ' bei Aktivieren kann man Projekt nicht selektieren 
                    node.Checked = False

                Case 1 ' Variante ist selektiert / nicht selektiert

                    ' ein Multiselect macht hier keinen Sinn ...
                    Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)

                    Dim projektNode As TreeNode = node.Parent
                    Dim pName As String = getProjectNameOfTreeNode(projektNode.Text)

                    ' jetzt die Charts , Einzel- wie Multiprojekt-Charts aktualisieren 
                    Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                    Call aktualisiereCharts(hproj, True)
                    Call awinNeuZeichnenDiagramme(2)

            End Select

            stopRecursion = False

        ElseIf aKtionskennung = PTTvActions.chgInSession Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    ' prüfen, ob Mauskey gedrückt war ...
                    If shiftKeywasPressed Then
                        If validMultiSelection(lastIndexChecked, currentIndex) Then
                            If lastIndexChecked < 0 Then
                                lastIndexChecked = 0
                            End If

                            Dim lb As Integer = lastIndexChecked
                            Dim ub As Integer = currentIndex
                            If lastIndexChecked > currentIndex Then
                                lb = currentIndex
                                ub = lastIndexChecked
                            End If

                            For h = lb To ub
                                Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(h)

                                If tmpNode.Level = treeLevel Then
                                    ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                                    tmpNode.Checked = checkMode
                                    Call doAfterCheckAction(aKtionskennung, treeLevel, tmpNode, considerDependencies)
                                End If

                            Next
                        Else
                            Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
                        End If

                    Else

                        Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)

                    End If

                    ' jetzt müssen die Portfolio Diagramme neu gezeichnet werden 
                    Call awinNeuZeichnenDiagramme(2)

                Case 1 ' Variante ist selektiert / nicht selektiert

                    ' eine Multiprojekt Selektion ist hier nicht erlaubt ...
                    Dim projektNode As TreeNode = node.Parent
                    Dim pName As String = getProjectNameOfTreeNode(projektNode.Text)

                    Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)


                    ' jetzt muss das bisherige aus ShowProjekte rausgenommen werden 
                    If ShowProjekte.contains(pName) And projektNode.Checked Then

                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                        Call aktualisiereCharts(hproj, True)
                        Call awinNeuZeichnenDiagramme(2)

                    End If

            End Select

            stopRecursion = False
        ElseIf aKtionskennung = PTTvActions.loadProjectAsTemplate Then
            stopRecursion = True

            Select Case treeLevel
                Case 0
                    If node.Checked Then
                        Dim anzKnoten As Integer = TreeViewProjekte.Nodes.Count
                        For i As Integer = 1 To anzKnoten
                            Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(i - 1)
                            If tmpNode.Name <> node.Name And tmpNode.Text <> node.Text Then
                                tmpNode.Checked = False
                            End If
                        Next
                    End If

                    Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)

                Case 1

                    Call doAfterCheckAction(aKtionskennung, treeLevel, node, considerDependencies)
            End Select


            stopRecursion = False

        End If

        ' merken , wo zum letzten Mal geklickt wurde ....
        lastLevelChecked = treeLevel
        lastIndexChecked = currentIndex

    End Sub

    ''' <summary>
    ''' führt die Aktion aus .. wird jetzt benötigt, um mit Shift mehrere Aktionen gleichzeitig durchführen zu können 
    ''' </summary>
    ''' <param name="actionCode"></param>
    ''' <param name="TreeLevel"></param>
    ''' <param name="node"></param>
    ''' <param name="considerDependencies"></param>
    ''' <remarks></remarks>
    Private Sub doAfterCheckAction(ByVal actionCode As Integer, ByVal TreeLevel As Integer, ByVal node As TreeNode,
                                       ByVal considerDependencies As Boolean)

        Dim err As New clsErrorCodeMsg

        Dim childNode As TreeNode
        Dim parentNode As TreeNode

        If actionCode = PTTvActions.delFromDB Or
            actionCode = PTTvActions.delAllExceptFromDB Then


            Select Case TreeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    Dim checkMode As Boolean = node.Checked

                    For i = 1 To node.Nodes.Count
                        ' Schleife über alle Varianten
                        childNode = node.Nodes.Item(i - 1)
                        childNode.Checked = checkMode
                        For j = 1 To childNode.Nodes.Count
                            ' Schleife über alle TimeStamps 
                            childNode.Nodes.Item(j - 1).Checked = checkMode
                        Next
                    Next

                Case 1 ' Variante ist selektiert / nicht selektiert

                    ' nach unten: das Gleiche 
                    For i = 1 To node.Nodes.Count
                        childNode = node.Nodes.Item(i - 1)
                        childNode.Checked = node.Checked
                    Next
                    ' nach oben 

                    If node.Checked = False Then
                        node.Parent.Checked = False
                    End If

                    ' wenn mit diesem Knoten jetzt alle geckecked/unchecked sind, soll auch parent wieder gesetzt werden 
                    If node.Checked = True Then
                        parentNode = node.Parent
                        Dim allchecked As Boolean = True
                        For i = 1 To parentNode.Nodes.Count
                            allchecked = allchecked And parentNode.Nodes.Item(i - 1).Checked
                        Next
                        If allchecked Then
                            parentNode.Checked = True
                        End If
                    End If

                Case 2 ' Snapshot ist selektiert / nicht selektiert 
                    If node.Checked = False Then
                        node.Parent.Checked = False
                        parentNode = node.Parent
                        parentNode.Parent.Checked = False
                    End If

                    If node.Checked = True Then
                        parentNode = node.Parent
                        Dim allchecked As Boolean = True
                        For i = 1 To parentNode.Nodes.Count
                            allchecked = allchecked And parentNode.Nodes.Item(i - 1).Checked
                        Next
                        If allchecked Then
                            ' jetzt wird bewusst Rekursion angestossen, damit das nach oben weitergeht ...
                            stopRecursion = False
                            parentNode.Checked = True
                        End If
                    End If

            End Select



        ElseIf actionCode = PTTvActions.delFromSession Or
              actionCode = PTTvActions.deleteV Then


            Select Case TreeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    For i = 1 To node.Nodes.Count
                        childNode = node.Nodes.Item(i - 1)
                        childNode.Checked = node.Checked
                        For j = 1 To childNode.Nodes.Count
                            childNode.Nodes.Item(j - 1).Checked = node.Checked
                        Next
                    Next

                Case 1 ' Variante ist selektiert / nicht selektiert

                    ' nach unten: das Gleiche 
                    For i = 1 To node.Nodes.Count
                        childNode = node.Nodes.Item(i - 1)
                        childNode.Checked = node.Checked
                    Next
                    ' nach oben 

                    If node.Checked = False Then
                        node.Parent.Checked = False
                    End If

                    ' wenn mit diesem Knoten jetzt alle gesetzt sind, soll auch parent wieder gesetzt werden 
                    If node.Checked = True Then
                        parentNode = node.Parent
                        Dim allchecked As Boolean = True
                        For i = 1 To parentNode.Nodes.Count
                            allchecked = allchecked And parentNode.Nodes.Item(i - 1).Checked
                        Next
                        If allchecked Then
                            parentNode.Checked = True
                        End If
                    End If


            End Select


        ElseIf actionCode = PTTvActions.setWriteProtection Then

            If Not noDB Then

                'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)

                Select Case TreeLevel

                    Case 0 ' Projekt ist selektiert / nicht selektiert 

                        If node.Nodes.Count = 0 Then
                            ' es gibt nur die eine Projekt-Variante 
                            Dim pName As String = getProjectNameOfTreeNode(node.Text)
                            Dim vName As String = ""
                            Dim variantNames As Collection = AlleProjekte.getVariantNames(pName, False)
                            If variantNames.Count > 0 Then
                                vName = CStr(variantNames.Item(1))
                            End If


                            If setNodeWriteProtections(node, PTTreeNodeTyp.project, pName, vName, node.Checked) Then
                                ' erfolgreich ..
                                ' es wurde bereits Node Apperance inkl Check-Status geklärt
                            Else
                                ' nicht zugelassen , also wieder zurücknehmen 

                                ' wenn node gecheckt wurde, aber das Projekt gar nicht existiert ...
                                If Not CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, vName, Date.Now, err) Then
                                    If awinSettings.englishLanguage Then
                                        Call MsgBox(pName & ", " & vName & "not yet stored in database ... " & vbLf &
                                                    "please store at database before protecting ...")
                                    Else
                                        Call MsgBox(pName & ", " & vName & "bitte erst in Datenbank speichern ... " & vbLf &
                                                    "dann schützen ...")
                                    End If
                                End If

                                node.Checked = Not node.Checked
                                writeProtections.upsert(CType(databaseAcc, DBAccLayer.Request).getWriteProtection(pName, vName, err))
                                Call bestimmeNodeAppearance(node, aKtionskennung, PTTreeNodeTyp.project, pName, vName)
                            End If


                        Else
                            ' es gibt mehrere Projekt-Varianten 
                            Dim atleastOneError As Boolean = False
                            Dim pName As String = getProjectNameOfTreeNode(node.Text)
                            Dim vName As String = ""

                            For i = 1 To node.Nodes.Count
                                childNode = node.Nodes.Item(i - 1)
                                ' darf es ge- bzw. entcheckt werden ? 
                                vName = getVariantNameOfTreeNode(childNode.Text)

                                If setNodeWriteProtections(childNode, PTTreeNodeTyp.pVariant, pName, vName, node.Checked) Then
                                    ' erfolgreich ..
                                    ' es wurde bereits Node Apperance inkl Check-Status geklärt
                                Else

                                    If Not CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, vName, Date.Now, err) Then
                                        If awinSettings.englishLanguage Then
                                            Call MsgBox(pName & ", " & vName & " not yet stored in database ... " & vbLf &
                                                        "please store at database before protecting ...")
                                        Else
                                            Call MsgBox(pName & ", " & vName & "bitte erst in Datenbank speichern ... " & vbLf &
                                                        "dann schützen ...")
                                        End If
                                    End If

                                    ' nicht zugelassen , also alles unverändert lassen  
                                    atleastOneError = True
                                    writeProtections.upsert(CType(databaseAcc, DBAccLayer.Request).getWriteProtection(pName, vName, err))
                                    Call bestimmeNodeAppearance(childNode, aKtionskennung, PTTreeNodeTyp.pVariant, pName, vName)
                                End If


                            Next

                            ' jetzt korrigieren, wenn eines der Kinder nicht auf den gleichen Check-Status gesetzt werden konnte
                            If atleastOneError And node.Checked Then
                                node.Checked = Not node.Checked
                            End If

                            ' jetzt für den Sammel-Knoten die Appearance bestimmen , der Name der Variante ist in diesem Fall egal, weil sich 
                            ' die Appearance ohnehin daran orientiert, was der Status der Kinder ist 
                            Call bestimmeNodeAppearance(node, aKtionskennung, PTTreeNodeTyp.project, pName, "")

                        End If


                    Case 1 ' Variante ist selektiert / nicht selektiert

                        ' ANfang 

                        parentNode = node.Parent
                        ' darf es ge- bzw. entcheckt werden ? 
                        Dim pName As String = getProjectNameOfTreeNode(parentNode.Text)
                        Dim vName As String = getVariantNameOfTreeNode(node.Text)

                        If setNodeWriteProtections(node, PTTreeNodeTyp.pVariant, pName, vName, node.Checked) Then
                            ' erfolgreich ..
                            ' es wurde bereits Node Apperance inkl Check-Status geklärt
                        Else
                            If Not CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, vName, Date.Now, err) Then
                                If awinSettings.englishLanguage Then
                                    Call MsgBox(pName & ", " & vName & " not yet stored in database ... " & vbLf &
                                                "please store at database before protecting ...")
                                Else
                                    Call MsgBox(pName & ", " & vName & " bitte erst in Datenbank speichern ... " & vbLf &
                                                "dann schützen ...")
                                End If
                            End If

                            ' nicht zugelassen , also alles unverändert lassen  
                            node.Checked = Not node.Checked
                            writeProtections.upsert(CType(databaseAcc, DBAccLayer.Request).getWriteProtection(pName, vName, err))
                            Call bestimmeNodeAppearance(node, aKtionskennung, PTTreeNodeTyp.pVariant, pName, vName)
                        End If

                        ' nach oben checken, ob jetzt das Projekt entsprechend gesetzt weren muss 
                        parentNode = node.Parent
                        If node.Checked = False Then
                            parentNode.Checked = False
                            Call bestimmeNodeAppearance(parentNode, aKtionskennung, PTTreeNodeTyp.project, pName, vName)

                        ElseIf node.Checked = True Then
                            Dim allchecked As Boolean = True
                            For i = 1 To parentNode.Nodes.Count
                                allchecked = allchecked And parentNode.Nodes.Item(i - 1).Checked
                            Next
                            If allchecked Then
                                parentNode.Checked = True
                                Call bestimmeNodeAppearance(parentNode, aKtionskennung, PTTreeNodeTyp.project, pName, vName)
                            End If
                        End If


                End Select

            Else
                ' zurücknehmen
                node.Checked = Not node.Checked
            End If

        ElseIf actionCode = PTTvActions.loadPV Then


            Select Case TreeLevel
                Case 0 ' Project Node was checked / unchecked 
                    Dim projektNode As TreeNode = node

                    If projektNode.Checked = True Then
                        ' eine muss gechecked werden
                        Dim variantNameLookingFor As String = ""

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And awinSettings.loadPFV Then
                            variantNameLookingFor = ptVariantFixNames.pfv.ToString
                        Else
                            variantNameLookingFor = ""
                        End If

                        Dim found As Boolean = False
                        For i = 0 To projektNode.Nodes.Count - 1
                            If getVariantNameOfTreeNode(projektNode.Nodes.Item(i).Text) = variantNameLookingFor Then
                                projektNode.Nodes.Item(i).Checked = True
                                found = True
                            Else
                                projektNode.Nodes.Item(i).Checked = False
                            End If
                        Next

                        If Not found Then
                            ' einfach die erste Variante auf checked setzen 
                            If projektNode.Nodes.Count > 0 Then
                                projektNode.Nodes.Item(0).Checked = True
                            End If

                        End If

                    Else
                        ' alle Varianten müssen unchecked werden 
                        ' alle anderen Varianten auf Unchecked setzen 
                        For i = 0 To projektNode.Nodes.Count - 1
                            projektNode.Nodes.Item(i).Checked = False
                        Next
                    End If

                Case 1 ' Variant Node was checked / unchecked 
                    Dim projektNode As TreeNode = node.Parent
                    Dim variantNode As TreeNode = node

                    If variantNode.Checked = True Then
                        If projektNode.Checked = False Then
                            projektNode.Checked = True
                        End If
                    Else
                        ' wenn es die letzte Variante war, die unchecked wurde 
                        Dim anzVariantsChecked As Integer = 0
                        For i = 0 To projektNode.Nodes.Count - 1
                            If projektNode.Nodes.Item(i).Checked = True Then
                                anzVariantsChecked = anzVariantsChecked + 1
                            End If
                        Next
                        If anzVariantsChecked = 0 Then
                            ' den Projekt-Knoten auch wieder zurücksetzen
                            projektNode.Checked = False
                        End If
                    End If

            End Select

        ElseIf actionCode = PTTvActions.loadPVInPPT Then
            ' nur single Selection erlaubt 


            Select Case TreeLevel
                Case 0 ' Project Node was checked / unchecked 
                    Dim projektNode As TreeNode = node

                    If projektNode.Checked = True Then
                        ' eine muss gechecked werden
                        Dim variantNameLookingFor As String = ""

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And awinSettings.loadPFV Then
                            variantNameLookingFor = ptVariantFixNames.pfv.ToString
                        Else
                            variantNameLookingFor = ""
                        End If

                        Dim found As Boolean = False
                        For i = 0 To projektNode.Nodes.Count - 1
                            If getVariantNameOfTreeNode(projektNode.Nodes.Item(i).Text) = variantNameLookingFor Then
                                projektNode.Nodes.Item(i).Checked = True
                                found = True
                            Else
                                projektNode.Nodes.Item(i).Checked = False
                            End If
                        Next

                        If Not found Then
                            ' einfach die erste Variante auf checked setzen 
                            If projektNode.Nodes.Count > 0 Then
                                projektNode.Nodes.Item(0).Checked = True
                            End If

                        End If

                        ' jetzt müssen alle anderen PRoject Nodes unchecked werden 

                        Call uncheckExcept(node.Name)
                        'For i = 1 To TreeViewProjekte.Nodes.Count
                        '    Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(i - 1)
                        '    If tmpNode.Level = 0 And tmpNode.Name <> node.Name Then
                        '        If tmpNode.Checked Then
                        '            tmpNode.Checked = False
                        '            ' dann auch alle Varianten unchecken ... 
                        '            Dim anzV As Integer = tmpNode.Nodes.Count

                        '            For vi As Integer = 1 To anzV
                        '                If tmpNode.Nodes.Item(vi - 1).Checked Then
                        '                    tmpNode.Nodes.Item(vi - 1).Checked = False
                        '                End If
                        '            Next
                        '        End If
                        '    End If
                        'Next


                    Else
                        ' alle Varianten müssen unchecked werden 
                        ' alle anderen Varianten auf Unchecked setzen 
                        For i = 0 To projektNode.Nodes.Count - 1
                            projektNode.Nodes.Item(i).Checked = False
                        Next
                    End If

                Case 1 ' Variant Node was checked / unchecked 
                    Dim projektNode As TreeNode = node.Parent
                    Dim variantNode As TreeNode = node

                    If variantNode.Checked = True Then

                        ' jetzt alle anderen Varianten auf unchecked setzen
                        Dim anzV As Integer = projektNode.Nodes.Count
                        For vi As Integer = 1 To anzV
                            If projektNode.Nodes.Item(vi - 1).Text <> variantNode.Text Then
                                projektNode.Nodes.Item(vi - 1).Checked = False
                            End If
                        Next

                        If projektNode.Checked = False Then
                            projektNode.Checked = True

                            Call uncheckExcept(projektNode.Name)
                            ' jetzt müssen hier noch alle anderen ProjektNodes und ggf selektierten Varianten auf Un-checked gesetzt werden 
                        End If


                    Else
                        ' wenn es die letzte Variante war, die unchecked wurde 
                        Dim anzVariantsChecked As Integer = 0
                        For i = 0 To projektNode.Nodes.Count - 1
                            If projektNode.Nodes.Item(i).Checked = True Then
                                anzVariantsChecked = anzVariantsChecked + 1
                            End If
                        Next
                        If anzVariantsChecked = 0 Then
                            ' den Projekt-Knoten auch wieder zurücksetzen
                            projektNode.Checked = False
                        End If
                    End If

            End Select

        ElseIf actionCode = PTTvActions.loadMultiPVInPPT Then
            ' Mehrfach Selektion erlaubt  


            Select Case TreeLevel
                Case 0 ' Project Node was checked / unchecked 
                    Dim projektNode As TreeNode = node

                    If projektNode.Checked = True Then
                        ' eine muss gechecked werden
                        Dim variantNameLookingFor As String = ""

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And awinSettings.loadPFV Then
                            variantNameLookingFor = ptVariantFixNames.pfv.ToString
                        Else
                            variantNameLookingFor = ""
                        End If

                        Dim found As Boolean = False
                        For i = 0 To projektNode.Nodes.Count - 1
                            If getVariantNameOfTreeNode(projektNode.Nodes.Item(i).Text) = variantNameLookingFor Then
                                projektNode.Nodes.Item(i).Checked = True
                                found = True
                            Else
                                projektNode.Nodes.Item(i).Checked = False
                            End If
                        Next

                        If Not found Then
                            ' einfach die erste Variante auf checked setzen 
                            If projektNode.Nodes.Count > 0 Then
                                projektNode.Nodes.Item(0).Checked = True
                            End If

                        End If

                    Else
                        ' alle Varianten müssen unchecked werden 
                        ' alle anderen Varianten auf Unchecked setzen 
                        For i = 0 To projektNode.Nodes.Count - 1
                            projektNode.Nodes.Item(i).Checked = False
                        Next
                    End If

                Case 1 ' Variant Node was checked / unchecked 
                    Dim projektNode As TreeNode = node.Parent
                    Dim variantNode As TreeNode = node

                    If variantNode.Checked = True Then
                        If projektNode.Checked = False Then
                            projektNode.Checked = True
                        End If
                    Else
                        ' wenn es die letzte Variante war, die unchecked wurde 
                        Dim anzVariantsChecked As Integer = 0
                        For i = 0 To projektNode.Nodes.Count - 1
                            If projektNode.Nodes.Item(i).Checked = True Then
                                anzVariantsChecked = anzVariantsChecked + 1
                            End If
                        Next
                        If anzVariantsChecked = 0 Then
                            ' den Projekt-Knoten auch wieder zurücksetzen
                            projektNode.Checked = False
                        End If
                    End If

            End Select

        ElseIf actionCode = PTTvActions.activateV Then

            Select Case TreeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    ' bei Aktivieren kann man Projekt nicht selektieren 
                    node.Checked = False

                Case 1 ' Variante ist selektiert / nicht selektiert


                    Dim projektNode As TreeNode = node.Parent
                    Dim selectedVariantName As String = node.Text
                    Dim pName As String = getProjectNameOfTreeNode(projektNode.Text)

                    ' es kann immer nur eine Variante selektiert sein; wenn die bisher aktive de-selektiert wird, 
                    ' wird Standard auf checked gesetzt 

                    If node.Checked = True Then

                        ' alle anderen Varianten auf Unchecked setzen 
                        For i = 0 To projektNode.Nodes.Count - 1
                            If projektNode.Nodes.Item(i).Text <> selectedVariantName Then
                                projektNode.Nodes.Item(i).Checked = False
                            End If
                        Next

                        ' jetzt die selektierte Variante ins ShowProjekte stecken und aktualisieren ... 
                        selectedVariantName = getVariantNameOfTreeNode(node.Text)



                    Else

                        ' die Standard Variante auf Checked setzen 
                        ' bzw. was besser ist, den ersten Child-Knoten 
                        ' das funktioniert nämlich auch dann, wenn keine Variante mit Name "" existiert 
                        If projektNode.Nodes.Count > 0 Then
                            projektNode.Nodes.Item(0).Checked = True
                            selectedVariantName = getVariantNameOfTreeNode(projektNode.Nodes.Item(0).Text)
                        Else
                            ' darf eigentlich gar nicht vorkommen 
                            selectedVariantName = ""
                        End If

                    End If

                    ' jetzt das Browser Szenario aktualisieren 
                    currentBrowserConstellation.updateShowAttributes(pName, selectedVariantName, node.Checked)

                    ' jetzt die Variante aktivieren 
                    Call replaceProjectVariant(pName, selectedVariantName, True, True, 0)

                    ' jetzt den Text des ParentNodes aktualisieren  
                    Call bestimmeNodeAppearance(projektNode, aKtionskennung, PTTreeNodeTyp.project, pName, selectedVariantName)

            End Select


        ElseIf actionCode = PTTvActions.chgInSession Then


            Select Case TreeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    Dim pName As String = getProjectNameOfTreeNode(node.Text)
                    Dim variantNames As Collection = AlleProjekte.getVariantNames(pName, False)
                    Dim selectedVariantName As String = ""

                    If variantNames.Count > 0 Then
                        selectedVariantName = CStr(variantNames.Item(1))
                    End If

                    If node.Checked Then
                        ' wurde neu hinzugefügt 
                        ' war bereits vorher irgendwann mal eine Variante gewählt ?
                        Dim selectionExisted As Boolean = False
                        For j = 1 To node.Nodes.Count
                            childNode = node.Nodes.Item(j - 1)
                            If childNode.Checked Then
                                selectionExisted = True
                                selectedVariantName = getVariantNameOfTreeNode(childNode.Text)
                            End If
                        Next

                        If Not selectionExisted And node.Nodes.Count > 0 Then
                            childNode = node.Nodes.Item(0)
                            childNode.Checked = True
                            selectedVariantName = getVariantNameOfTreeNode(childNode.Text)
                        End If

                        ' jetzt das Browser Szenario aktualisieren 
                        Call currentBrowserConstellation.updateShowAttributes(pName, selectedVariantName, node.Checked)

                        Call putProjectInShow(pName:=pName,
                                          vName:=selectedVariantName, considerDependencies:=considerDependencies,
                                          upDateDiagrams:=False,
                                          myConstellation:=currentBrowserConstellation)

                        ' jetzt muss das Projekt aus AlleProjekte auch in ShowProjekte transferiert werden 

                        ' jetzt muss noch die TreeView ggf angepasst werden, wenn considerDependencies true ist 
                        If considerDependencies Then
                            ' ggf. die Projekte einblenden, von denen dieses Projekt abhängt 
                            Dim toDoListe As Collection = allDependencies.passiveListe(pName, PTdpndncyType.inhalt)
                            If toDoListe.Count > 0 Then
                                For Each mprojectName As String In toDoListe
                                    Call activateMasterProject(mprojectName)
                                Next

                            End If
                        End If
                    Else
                        ' wurde abgewählt 
                        Call currentBrowserConstellation.updateShowAttributes(pName, Nothing, False)

                        Call putProjectInNoShow(pName, considerDependencies, False)

                        If considerDependencies Then
                            Dim toDoListe As Collection = allDependencies.activeListe(pName, PTdpndncyType.inhalt)
                            If toDoListe.Count > 0 Then
                                For Each dprojectName As String In toDoListe
                                    Call deactivateDependentProject(dprojectName)
                                Next

                            End If
                        End If


                    End If

                    ' jetzt den Text des Projekt-Knotens aktualisieren  
                    Call bestimmeNodeAppearance(node, aKtionskennung, PTTreeNodeTyp.project, pName, selectedVariantName)

                Case 1 ' Variante ist selektiert / nicht selektiert


                    Dim projektNode As TreeNode = node.Parent
                    Dim selectedVariantName As String = node.Text
                    Dim pName As String = getProjectNameOfTreeNode(projektNode.Text)

                    ' es kann immer nur eine Variante selektiert sein; wenn die bisher aktive de-selektiert wird, 
                    ' wird Standard auf checked gesetzt 

                    If node.Checked = True Then

                        ' alle anderen Varianten auf Unchecked setzen 
                        For i = 0 To projektNode.Nodes.Count - 1
                            If projektNode.Nodes.Item(i).Text <> selectedVariantName Then
                                projektNode.Nodes.Item(i).Checked = False
                            End If
                        Next

                        ' jetzt die selektierte Variante ins ShowProjekte stecken und aktualisieren ... 
                        selectedVariantName = getVariantNameOfTreeNode(node.Text)

                        Call currentBrowserConstellation.updateShowAttributes(pName, selectedVariantName, True)


                    Else

                        ' die Standard Variante auf Checked setzen 
                        ' bzw. was besser ist, den ersten Child-Knoten 
                        ' das funktioniert nämlich auch dann, wenn keine Variante mit Name "" existiert 
                        If projektNode.Nodes.Count > 0 Then
                            projektNode.Nodes.Item(0).Checked = True
                            selectedVariantName = getVariantNameOfTreeNode(projektNode.Nodes.Item(0).Text)
                        Else
                            ' darf eigentlich gar nicht vorkommen 
                            selectedVariantName = ""
                        End If

                        Call currentBrowserConstellation.updateShowAttributes(pName, selectedVariantName, False)


                    End If

                    ' jetzt muss das bisherige aus ShowProjekte rausgenommen werden 
                    If ShowProjekte.contains(pName) And projektNode.Checked Then

                        Call replaceProjectVariant(pName, selectedVariantName, False, True, 0)

                    End If

                    ' jetzt den Text des ParentNodes aktualisieren  
                    Call bestimmeNodeAppearance(projektNode, aKtionskennung, PTTreeNodeTyp.project, pName, selectedVariantName)

            End Select
        ElseIf actionCode = PTTvActions.loadProjectAsTemplate Then

            Select Case TreeLevel
                Case 0 ' Projekt ist selektiert / nicht selektiert 
                    Dim pName As String = getProjectNameOfTreeNode(node.Text)
                    Dim variantNames As Collection = AlleProjekte.getVariantNames(pName, False)
                    Dim selectedVariantName As String = ""

                    If variantNames.Count > 0 Then
                        selectedVariantName = CStr(variantNames.Item(1))
                    End If

                    If node.Checked Then
                        ' wurde neu hinzugefügt 
                        ' war bereits vorher irgendwann mal eine Variante gewählt ?


                        Dim selectionExisted As Boolean = False
                        For j = 1 To node.Nodes.Count
                            childNode = node.Nodes.Item(j - 1)
                            If childNode.Checked Then
                                selectionExisted = True
                                selectedVariantName = getVariantNameOfTreeNode(childNode.Text)
                            End If
                        Next

                        If Not selectionExisted And node.Nodes.Count > 0 Then
                            childNode = node.Nodes.Item(0)
                            childNode.Checked = True
                            selectedVariantName = getVariantNameOfTreeNode(childNode.Text)
                        End If

                    Else
                        ' wurde abgewählt 
                        ' nichts weiter tun 
                    End If

                    ' jetzt den Text des Projekt-Knotens aktualisieren  
                    Call bestimmeNodeAppearance(node, aKtionskennung, PTTreeNodeTyp.project, pName, selectedVariantName)


                Case 1 ' Variante ist selektiert / nicht selektiert 

                    Dim projektNode As TreeNode = node.Parent
                    Dim selectedVariantName As String = node.Text
                    Dim pName As String = getProjectNameOfTreeNode(projektNode.Text)

                    ' es kann immer nur eine Variante selektiert sein; wenn die bisher aktive de-selektiert wird, 
                    ' wird Standard auf checked gesetzt 

                    If node.Checked = True Then

                        ' alle anderen Varianten auf Unchecked setzen 
                        For i = 0 To projektNode.Nodes.Count - 1
                            If projektNode.Nodes.Item(i).Text <> selectedVariantName Then
                                projektNode.Nodes.Item(i).Checked = False
                            End If
                        Next

                    Else

                        ' die Standard Variante auf Checked setzen 
                        ' bzw. was besser ist, den ersten Child-Knoten 
                        ' das funktioniert nämlich auch dann, wenn keine Variante mit Name "" existiert 
                        If projektNode.Nodes.Count > 0 Then
                            projektNode.Nodes.Item(0).Checked = True
                            selectedVariantName = getVariantNameOfTreeNode(projektNode.Nodes.Item(0).Text)
                        Else
                            ' darf eigentlich gar nicht vorkommen 
                            selectedVariantName = ""
                        End If



                    End If


                    ' jetzt den Text des ParentNodes aktualisieren  
                    Call bestimmeNodeAppearance(projektNode, aKtionskennung, PTTreeNodeTyp.project, pName, selectedVariantName)

            End Select

        End If
    End Sub

    ''' <summary>
    ''' prüft, ob es sich um eine gültige Multi-Selection handelt: Beginn und Ende müssen auf der gleichen Stufe sein
    ''' wenn von Stufe 2 oder 3 gestartet wird, es darf nur innerhalb der aktuellen Parent-Struktur sein; also von Variante 1/ Project 2
    ''' bis Variante 7 / Project 8 darf nicht gewählt werden 
    ''' </summary>
    ''' <param name="lastIXChecked"></param>
    ''' <param name="currentIX"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function validMultiSelection(ByVal lastIXChecked As Integer, ByVal currentIX As Integer) As Boolean
        Dim tmpResult As Boolean = True


        If lastIXChecked < 0 Then
            lastIXChecked = 0
        End If

        Dim lb As Integer = lastIXChecked
        Dim ub As Integer = currentIX
        If lastIndexChecked > currentIX Then
            lb = currentIX
            ub = lastIXChecked
        End If

        Dim vglLevel As Integer = TreeViewProjekte.Nodes.Item(lb).Level

        If TreeViewProjekte.Nodes.Item(ub).Level <> vglLevel Then
            tmpResult = False
        Else
            For h = lb To ub
                If TreeViewProjekte.Nodes.Item(h).Level > vglLevel Then
                    tmpResult = False
                    Exit For
                End If
            Next
        End If

        validMultiSelection = tmpResult
    End Function
    ''' <summary>
    ''' aktiviert das Master-Projekt, wenn es nicht schon aktiviert ist ...
    ''' </summary>
    ''' <param name="mprojectName"></param>
    ''' <remarks></remarks>
    Private Sub activateMasterProject(ByVal mprojectName As String)

        For i As Integer = 1 To TreeViewProjekte.GetNodeCount(False)
            Dim curItem As TreeNode = TreeViewProjekte.Nodes.Item(i - 1)
            Dim curItemA As TreeNode = TreeViewProjekte.Nodes.Item(mprojectName)

            If getProjectNameOfTreeNode(curItem.Text) = mprojectName Then
                If curItem.Checked Then
                    ' nichts tun 
                Else
                    stopRecursion = False
                    curItem.Checked = True
                    stopRecursion = True
                End If
            End If



        Next


    End Sub

    ''' <summary>
    ''' de-aktiviert das abhängige Projekt, wenn es nicht schon de-aktiviert ist 
    ''' </summary>
    ''' <param name="dprojectName"></param>
    ''' <remarks></remarks>
    Private Sub deactivateDependentProject(ByVal dprojectName As String)

        For i As Integer = 1 To TreeViewProjekte.GetNodeCount(False)
            Dim curItem As TreeNode = TreeViewProjekte.Nodes.Item(i - 1)

            If getProjectNameOfTreeNode(curItem.Text) = dprojectName Then
                If curItem.Checked Then
                    stopRecursion = False
                    curItem.Checked = False
                    stopRecursion = True
                Else
                    ' nichts tun
                End If
            End If

        Next

    End Sub
    ''' <summary>
    ''' wird aufgerufen, wenn ein TreeItem (eine Zeile) selektiert wird 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TreeViewProjekte_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeViewProjekte.AfterSelect

        Dim node As TreeNode = e.Node
        Dim treeLevel As Integer = node.Level
        Dim projectName As String
        Dim variantName As String = ""
        Dim toolTippText As String = "-"
        Dim hproj As clsProjekt

        If aKtionskennung = PTTvActions.loadPVInPPT Or aKtionskennung = PTTvActions.loadMultiPVInPPT Then
            Exit Sub
        End If

        If treeLevel = 0 Then
            projectName = getProjectNameOfTreeNode(node.Text)

            'Dim variantNames As Collection = AlleProjekte.getVariantNames(projectName, False)
            Dim variantNames As Collection = currentBrowserConstellation.getVariantNames(projectName, False)
            If variantNames.Count > 0 Then
                variantName = variantNames.Item(1)
                If aKtionskennung = PTTvActions.chgInSession Then
                    ' welche Varianten sind gecheckt ? 
                    Dim checkedVariantNames As Collection = Me.getNamesOfChildNodes(node, True)
                    If checkedVariantNames.Count > 0 Then
                        variantName = checkedVariantNames.Item(1)
                    End If
                End If
            Else
                variantName = ""
            End If

            hproj = AlleProjekte.getProject(projectName, variantName)

            If Not IsNothing(hproj) Then
                toolTippText = getToolTippText(hproj, "", "", treeLevel, variantNames.Count)

                ' Anzeige der aktualisierten Charts und Phasen- bzw Milestone Infor Formulare 
                'Call aktualisierePMSForms(hproj)
                Call aktualisiereCharts(hproj, True)
            Else
                toolTippText = getToolTippText(Nothing, projectName, variantName, treeLevel, variantNames.Count)
            End If


        ElseIf treeLevel = 1 Then
            Dim projectNode As TreeNode = node.Parent
            If Not IsNothing(projectNode) Then

                projectName = getProjectNameOfTreeNode(projectNode.Text)
                variantName = getVariantNameOfTreeNode(node.Text)
                hproj = AlleProjekte.getProject(projectName, variantName)

                If Not IsNothing(hproj) Then

                    toolTippText = getToolTippText(hproj, "", "", treeLevel, 0)

                    ' Anzeige der aktualisierten Charts und Phasen- bzw Milestone Infor Formulare 
                    'Call aktualisierePMSForms(hproj)
                    Call aktualisiereCharts(hproj, True)

                Else
                    toolTippText = getToolTippText(Nothing, projectName, variantName, treeLevel, 1)
                End If

            End If
        End If

        ToolTipStand.Show(toolTippText, TreeViewProjekte, 6000)


    End Sub

    ''' <summary>
    ''' liefert für die übergebenen Parameter den TooltippText
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="level"></param>
    ''' <param name="anzahlVariants"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getToolTippText(ByVal hproj As clsProjekt, ByVal pName As String, ByVal vName As String,
                                     ByVal level As Integer, ByVal anzahlVariants As Integer) As String

        Dim tmpText As String = ""
        Dim allowedLength As Integer = 70

        If Not IsNothing(hproj) Then
            pName = hproj.name
            vName = hproj.variantName
        End If

        ' Projekt-Stufe
        If level = 0 Then

            If allDependencies.projectCount > 0 And toolTippsAreShowing = ptPPAtooltipps.dependencies Then
                tmpText = allDependencies.getDependencyInfos(pName)

            ElseIf toolTippsAreShowing = ptPPAtooltipps.protectedBy Then

                If anzahlVariants = 1 Then
                    Dim pvName As String = calcProjektKey(pName, vName)
                    tmpText = writeProtections.getProtectionText(pvName)
                Else
                    tmpText = ""
                End If

            ElseIf toolTippsAreShowing = ptPPAtooltipps.scenarioReferences Then
                tmpText = projectConstellations.getSzenarioNamesWith(pName, "$ALL")

            ElseIf Not IsNothing(hproj) Then
                If hproj.description.Length > 0 Then
                    tmpText = hproj.description
                    If tmpText.Length > allowedLength Then
                        tmpText = tmpText.Substring(0, allowedLength) & "..."
                    End If
                End If
            End If


        ElseIf level = 1 Then
            ' Varianten-Level

            If toolTippsAreShowing = ptPPAtooltipps.protectedBy Then

                Dim pvName As String = calcProjektKey(pName, vName)

                Dim lastUser As String = ""
                Dim zeitpunkt As Date
                lastUser = writeProtections.lastModifiedBy(pvName)
                zeitpunkt = writeProtections.changeDate(pvName)

                If writeProtections.isProtected(pvName) Then
                    Dim permanent As String = ""
                    If writeProtections.isPermanentProtected(pvName) Then
                        permanent = "permanent "
                    End If
                    If awinSettings.englishLanguage Then
                        tmpText = permanent & "protected by: " & lastUser & ", at: " & zeitpunkt.ToString
                    Else
                        tmpText = permanent & "geschützt von: " & lastUser & ", am: " & zeitpunkt.ToString
                    End If

                Else
                    If awinSettings.englishLanguage Then
                        tmpText = "no protection"
                    Else
                        tmpText = "nicht geschützt"
                    End If
                End If

            ElseIf toolTippsAreShowing = ptPPAtooltipps.scenarioReferences Then
                tmpText = projectConstellations.getSzenarioNamesWith(pName, vName)

            ElseIf Not IsNothing(hproj) Then

                If hproj.variantDescription.Length > 0 Then
                        tmpText = hproj.variantDescription
                        If tmpText.Length > allowedLength Then
                            tmpText = tmpText.Substring(0, allowedLength) & "..."
                        End If
                    End If
                End If

            ElseIf level = 2 Then
                ' noch kein ToolTippText verfügbar
            End If



            getToolTippText = tmpText

    End Function



    Private Sub TreeViewProjekte_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeViewProjekte.BeforeExpand


        Dim err As New clsErrorCodeMsg

        Dim selectedNode As New TreeNode
        Dim variantNode As New TreeNode
        Dim nodeTimeStamp As New TreeNode
        Dim projName As String = ""
        Dim variantName As String = ""
        'Dim hliste As SortedList(Of Date, String)
        Dim nodeLevel As Integer
        Dim variantListe As Collection
        Dim hproj As New clsProjekt
        Dim key As String

        'Platzhalter ...
        ' '' jetzt für chgInSession und activateV prüfen, welche Projekte denn im Show sind ... 
        ''If aKtionskennung = PTTvActions.chgInSession Or _
        ''    aKtionskennung = PTTvActions.activateV Then

        ''    ' es muss nur was gecheckt werden, wenn das Projekt im Show ist 
        ''    If projectIsShown And (CStr(variantNames.Item(iv)) = shownVariant) Then
        ''        tmpNodeLevel1.Checked = True
        ''    End If

        ''End If
        ' Platzhalter Ende ... 

        If Not noDB And aKtionskennung = PTTvActions.setWriteProtection Then
            ' jetzt die writeProtections neu bestimmen 
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)
        End If

        selectedNode = e.Node
        nodeLevel = e.Node.Level

        ' Projekt-Ebene
        If nodeLevel = 0 Then


            projName = getProjectNameOfTreeNode(selectedNode.Text)

            ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
            If selectedNode.Tag = "P" Then

                'Call MsgBox("sollte eigentlich gar nicht mehr vorkommen ...")
                ' Inhalte der Sub-Nodes müssen neu aufgebaut werden 
                If quickList Then
                    variantListe = getVariantListeFromPVNames(pvNamesList, projName)
                Else
                    variantListe = currentBrowserConstellation.getVariantNames(projName, True)
                End If

                ' Löschen von Platzhalter
                selectedNode.Nodes.Clear()

                ' wird nur im Fall loadPV benötigt ... wird baer hier besetzt, weil das sonst in der Schleife ständig neu ausgerechnet wird 
                ' es wird entweder by default die Basis Variante geladen, oder die pfv-Variante, sofern das so gesetzt ist und die customUserRole = PMGR
                Dim variantNameLookedFor As String = ""
                If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And awinSettings.loadPFV = True Then
                    variantNameLookedFor = ptVariantFixNames.pfv.ToString
                Else
                    variantNameLookedFor = ""
                End If

                ' Eintragen der zum Projekt gehörenden Varianten
                For Each variantName In variantListe
                    variantNode = selectedNode.Nodes.Add(CType(variantName, String))

                    ' jetzt muss gecheckt werden , ob es sich um das Aktivieren handelt oder nicht
                    If aKtionskennung = PTTvActions.activateV Or
                        aKtionskennung = PTTvActions.chgInSession Then
                        stopRecursion = True
                        If getVariantNameOfTreeNode(variantName) = hproj.variantName Then
                            variantNode.Checked = True
                        Else
                            variantNode.Checked = False
                        End If
                        stopRecursion = False

                    ElseIf aKtionskennung = PTTvActions.loadPV Then
                        ' es wird by default nur eine Projekt-Variante selektiert ...

                        key = calcProjektKey(pName:=projName, variantName:=variantName)

                        stopRecursion = True

                        ' es wird nur gecheckt, wenn selectedNode.checked 
                        If selectedNode.Checked = True Then
                            ' entscheiden, welche Variante gecheckt wird 
                            If variantName = variantNameLookedFor Then
                                variantNode.Checked = True
                            End If
                        Else
                            ' auf jeden Fall nicht checken ..
                            variantNode.Checked = False
                        End If

                        stopRecursion = False

                    ElseIf aKtionskennung = PTTvActions.delAllExceptFromDB Then
                        ' es können alle Elemente selektiert werden ...
                        stopRecursion = True
                        variantNode.Checked = selectedNode.Checked
                        stopRecursion = False

                    ElseIf aKtionskennung = PTTvActions.setWriteProtection And Not noDB Then
                        ' dieser Zweig wird nie mehr betreten ... 

                        variantName = getVariantNameOfTreeNode(variantName)

                        Dim pvName As String = calcProjektKey(projName, variantName)
                        stopRecursion = True
                        variantNode.Checked = writeProtections.isProtected(pvName)
                        Call bestimmeNodeAppearance(variantNode, aKtionskennung, PTTreeNodeTyp.pVariant, projName, variantName)
                        stopRecursion = False

                    Else
                        stopRecursion = True
                        variantNode.Checked = selectedNode.Checked
                        stopRecursion = False
                    End If



                    If aKtionskennung = PTTvActions.delFromDB Or
                        aKtionskennung = PTTvActions.loadPVS Then
                        ' Einfügen eines Platzhalters macht nur Sinn bei Snapshots löschen bzw. Snapshots laden 

                        variantNode.Tag = "P"
                        variantNode.Nodes.Add("()")
                    Else
                        variantNode.Tag = "X"
                    End If


                Next

                selectedNode.Tag = "X"


            Else
                ' und das hier sollte der Standard-Fall sein ...
                ' einfach im Falle chgInSession bzw. ActivateV die Variante aktiv setzen, die im showprojekte gezeigt wird 
                If (selectedNode.Checked And aKtionskennung = PTTvActions.chgInSession) Or
                    aKtionskennung = PTTvActions.activateV Then

                    If ShowProjekte.contains(projName) Then
                        hproj = ShowProjekte.getProject(projName)
                        variantName = "(" & hproj.variantName & ")"
                        stopRecursion = True
                        For Each tmpNode As TreeNode In selectedNode.Nodes
                            tmpNode.Checked = (tmpNode.Text = variantName)
                        Next
                        stopRecursion = False
                    End If

                ElseIf aKtionskennung = PTTvActions.setWriteProtection Then
                    stopRecursion = True

                    For Each tmpNode As TreeNode In selectedNode.Nodes
                        variantName = getVariantNameOfTreeNode(tmpNode.Text)
                        Dim pvName As String = calcProjektKey(projName, variantName)
                        tmpNode.Checked = writeProtections.isProtected(pvName)
                        Call bestimmeNodeAppearance(tmpNode, aKtionskennung, PTTreeNodeTyp.pVariant, projName, variantName)
                    Next

                    stopRecursion = False
                End If
            End If



        ElseIf nodeLevel = 1 And
            (aKtionskennung = PTTvActions.delFromDB Or aKtionskennung = PTTvActions.loadPVS) Then

            ' hier wurde eine Variante selektiert ...

            If selectedNode.Tag = "P" Then

                selectedNode.Tag = "X"
                projName = getProjectNameOfTreeNode(selectedNode.Parent.Text)
                variantName = getVariantNameOfTreeNode(selectedNode.Text)

                'hliste = projektHistorien.getTimeStamps(calcProjektKey(projName, variantName))

                'If hliste.Count = 0 Then

                If Not noDB Then

                    'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                    If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then
                    Else
                        Dim msgText As String
                        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                            msgText = "Datenbank-Verbindung ist unterbrochen!"
                        Else
                            msgText = "no connection to Database!"
                        End If
                        Call MsgBox(msgText)
                    End If

                    ' Lesen der TimeStamp Snapshots für ProjNAme, variantName 
                    Try
                        If Not projekthistorie Is Nothing Then
                            projekthistorie.clear()
                        Else
                            projekthistorie = New clsProjektHistorie
                        End If

                        projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=projName, variantName:=variantName,
                                                                            storedEarliest:=Date.MinValue, storedLatest:=requiredDate.Value, err:=err)

                    Catch ex As Exception
                        projekthistorie.clear()
                    End Try

                End If

                If projekthistorie.Count > 0 Then

                    'projektHistorien.Remove(projName & "#" & variantName, Date.MinValue) 'Platzhalter wieder entfernen
                    selectedNode.Nodes.Clear()  ' Löschen von Platzhalter

                    ' Aufbau der Listen 
                    'projektHistorien.Add(projekthistorie)

                    stopRecursion = True
                    ' Eintragen der zur Projekt-Variante gehörenden TimeStamps
                    ' aber nur dann, wenn Sie nicht nach dem required date liegen 
                    For Each kvp1 As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                        nodeTimeStamp = selectedNode.Nodes.Add(CType(kvp1.Value.timeStamp, String))
                        nodeTimeStamp.Checked = selectedNode.Checked
                    Next kvp1
                    stopRecursion = False

                Else

                    If projekthistorie.Count = 0 Then
                        ' keine ProjektHistorie vorhanden
                        'projektHistorien.Remove(projName & "#" & variantName, Date.MinValue) 'Platzhalter wieder entfernen
                        selectedNode.Nodes.Clear()  ' Löschen von Platzhalter
                    End If
                End If




                'End If

            End If


        End If





    End Sub

    ''' <summary>
    ''' liefert den Namen der Variante zurück, bereinigt um die öffnende und schließende Klammer
    ''' </summary>
    ''' <param name="nodeText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getVariantNameOfTreeNode(ByVal nodeText As String) As String
        Dim tmpstr() As String
        Dim vName As String = ""

        Try
            tmpstr = nodeText.Split(New Char() {CChar("("), CChar(")")}, 3)
            If tmpstr.Length = 3 Then
                vName = tmpstr(1).Trim
            End If
        Catch ex As Exception

        End Try

        getVariantNameOfTreeNode = vName

    End Function

    ''' <summary>
    ''' liefert den Namen des Projektes zurück, kann folgendermaßen aussehen project (variantname) L
    ''' </summary>
    ''' <param name="nodeText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getProjectNameOfTreeNode(ByVal nodeText As String) As String
        Dim tmpstr() As String
        Dim pName As String = ""

        Try
            If nodeText.EndsWith(" /D") Then
                nodeText = nodeText.Substring(0, nodeText.Length - 3)
            ElseIf nodeText.EndsWith(" /L") Then
                nodeText = nodeText.Substring(0, nodeText.Length - 3)
            ElseIf nodeText.EndsWith(" /LD") Then
                nodeText = nodeText.Substring(0, nodeText.Length - 4)
            End If

            tmpstr = nodeText.Split(New Char() {CChar("("), CChar(")")}, 3)
            If tmpstr.Length >= 1 Then
                pName = tmpstr(0).Trim
            End If
        Catch ex As Exception

        End Try

        getProjectNameOfTreeNode = pName

    End Function

    ''' <summary>
    ''' wird bei Auslösen des "Aktionsbuttons" ausgeführt; 
    ''' in Abhängigkeit von Aktionskennung 
    ''' dieser Button kann im Fall activate Variante gar nicht aktiviert werden, weil unsichtbar
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim projektNode As TreeNode, variantNode As TreeNode, timeStampNode As TreeNode
        Dim anzahlProjekte As Integer
        Dim anzahlVarianten As Integer
        Dim anzahlTimeStamps As Integer
        Dim pname As String = "", variantName As String = "", timestamp As Date
        ' nimmt den Namen auf , der in Powerpoint selektiert wird 
        Dim pptPname As String = ""
        Dim pptVname As String = ""
        'Dim hproj As clsProjekt
        Dim portfolioZeile As Integer = 2
        Dim storedAtOrBefore As Date
        Dim considerDependencies As Boolean

        Dim outPutCollection As New Collection
        Dim outPutHeader As String = ""
        Dim outPutExplanation As String = ""

        Dim calledFromPPT As Boolean = (aKtionskennung = PTTvActions.loadPVInPPT) Or (aKtionskennung = PTTvActions.loadMultiPVInPPT)

        ' Cursor auf Wait-Cursor setzen ... 
        Me.Cursor = Cursors.WaitCursor


        If allDependencies.projectCount > 0 Then
            considerDependencies = True
        Else
            considerDependencies = False
        End If

        ' ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        ' ''Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

        If IsNothing(requiredDate.Value) Then
            storedAtOrBefore = Date.Now.Date.AddHours(23).AddMinutes(59)
        Else

            storedAtOrBefore = requiredDate.Value

        End If

        ' Bestimmen der Überschrift des Output Headers, falls es irgendwelche Meldungen gibt
        If aKtionskennung = PTTvActions.delFromDB Or
            aKtionskennung = PTTvActions.delAllExceptFromDB Then

            If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                outPutHeader = "Löschen von Projekt-Varianten in der Datenbank "
                outPutExplanation = "Meldungen: "
            Else
                outPutHeader = "Delete Project-Variants in Database"
                outPutExplanation = "Messages"
            End If

        ElseIf aKtionskennung = PTTvActions.delFromSession Or
            aKtionskennung = PTTvActions.deleteV Then

            If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                outPutHeader = "Löschen von Projekt-Varianten in der Session "
                outPutExplanation = "Meldungen: "
            Else
                outPutHeader = "Delete Project-Variants in Session"
                outPutExplanation = "Messages"
            End If

        ElseIf aKtionskennung = PTTvActions.loadPV Then

            If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                outPutHeader = "Laden von Projekt-Varianten aus der Datenbank "
                outPutExplanation = "Meldungen: "
            Else
                outPutHeader = "Load Project-Variants from Database"
                outPutExplanation = "Messages"
            End If

        ElseIf aKtionskennung = PTTvActions.loadPVInPPT Then

            If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                outPutHeader = "Selektieren Sie ein Projekt oder eine Projekt-Variante "
                outPutExplanation = "Meldungen: "
            Else
                outPutHeader = "Select a project or project-variant from Database"
                outPutExplanation = "Messages"
            End If

        ElseIf aKtionskennung = PTTvActions.loadMultiPVInPPT Then

            If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                outPutHeader = "Selektieren ein oder mehrere Projekte oder Projekt-Varianten "
                outPutExplanation = "Meldungen: "
            Else
                outPutHeader = "Select one or several projects or project-variants "
                outPutExplanation = "Messages"
            End If

        End If




        Dim p As Integer, v As Integer, t As Integer

        If aKtionskennung = PTTvActions.delFromDB Or
            aKtionskennung = PTTvActions.delAllExceptFromDB Or
            aKtionskennung = PTTvActions.delFromSession Or
            aKtionskennung = PTTvActions.deleteV Or
            aKtionskennung = PTTvActions.loadPV Or
            aKtionskennung = PTTvActions.loadPVInPPT Or
            aKtionskennung = PTTvActions.loadMultiPVInPPT Then

            ' alle anderen Aktionen wie Projekte aus Datenbank löschen , aus Session löschen, aus Datenbank laden  ... 
            With TreeViewProjekte
                anzahlProjekte = .Nodes.Count

                For p = 1 To anzahlProjekte

                    projektNode = .Nodes.Item(p - 1)
                    pname = getProjectNameOfTreeNode(projektNode.Text)

                    If projektNode.Checked Then
                        ' Aktion auf allen Varianten und Timestamps 
                        ' Schleife über alle Varianten: 
                        ' lösche in Datenbank pname#vname

                        'anzahlVarianten = projektNode.Nodes.Count
                        Dim variantListe As New Collection

                        If quickList Then
                            variantListe = getVariantListeFromPVNames(pvNamesList, pname)
                        Else
                            variantListe = currentBrowserConstellation.getVariantNames(pname, True)
                        End If

                        anzahlVarianten = variantListe.Count

                        If aKtionskennung = PTTvActions.delFromSession Then

                            Call awinDeleteProjectInSession(pname, considerDependencies)

                            ' jetzt in der currentBrowserConstellation ändern 
                            Try
                                Dim variantNames As Collection = currentBrowserConstellation.getVariantNames(pname, False)
                                For i As Integer = 1 To variantNames.Count
                                    Dim tmpKey As String = calcProjektKey(pname, CStr(variantNames.Item(i)))
                                    currentBrowserConstellation.remove(tmpKey)
                                Next
                            Catch ex As Exception

                            End Try

                        ElseIf aKtionskennung = PTTvActions.delAllExceptFromDB Then

                            For v = 1 To anzahlVarianten

                                variantName = getVariantNameOfTreeNode(CStr(variantListe.Item(v)))
                                'Call deleteCompleteProjectVariant(outPutCollection,
                                'pname, variantName, aKtionskennung, versionsToKeep.Value)

                            Next


                        ElseIf aKtionskennung = PTTvActions.delFromDB Then

                            ' ur: 20190716 Versuch ein ganzes Projekt zu löschen - nur möglich, wenn nicht in Portfolio enthalten
                            '              dieser Check muss noch gemacht werden

                            'Dim err As New clsErrorCodeMsg
                            'Dim erledigt As Boolean = CType(databaseAcc, DBAccLayer.Request).removeCompleteProjectFromDB(pname, err)
                            Call deleteCompleteProjectFromDB(outPutCollection, pname)
                            'For v = 1 To anzahlVarianten
                            '    variantName = getVariantNameOfTreeNode(CStr(variantListe.Item(v)))
                            '    ' Fehler-Behandlung, d.h auch Abfrage ob PName#vName referenziert in Szenario ist, passiert dort drin ... 
                            '    Call deleteCompleteProjectVariant(outPutCollection,
                            '                                     pname, variantName, aKtionskennung)
                            'Next


                        ElseIf aKtionskennung = PTTvActions.loadPV Then

                            Dim hproj As clsProjekt = Nothing
                            If ShowProjekte.Count > 0 Then
                                If ShowProjekte.contains(pname) Then
                                    hproj = ShowProjekte.getProject(pname)
                                End If
                            End If


                            ' da hier manuell Projekte hinzu kommen, muss der Sort-Type auf customTF gesetzt werden 
                            currentBrowserConstellation.sortCriteria = ptSortCriteria.customTF
                            currentSessionConstellation.sortCriteria = ptSortCriteria.customTF


                            Dim variantNameLookingFor As String = ""
                            If (myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And awinSettings.loadPFV) Or awinSettings.filterPFV Then
                                variantNameLookingFor = ptVariantFixNames.pfv.ToString
                            End If

                            ' jetzt muss geprüft werden, welcher Name tatsächlich ins Show gesteckt werden soll 
                            Dim nameOfFirstChecked As String = ""
                            Dim firstTime As Boolean = True
                            Dim found As Boolean = False


                            For i As Integer = 0 To anzahlVarianten - 1

                                If projektNode.Nodes.Count > 0 Then
                                    If projektNode.Nodes.Item(i).Checked = True Then

                                        Dim curVariantName As String = getVariantNameOfTreeNode(projektNode.Nodes.Item(i).Text)

                                        If firstTime Then
                                            nameOfFirstChecked = curVariantName
                                            firstTime = False
                                        End If

                                        If curVariantName = variantNameLookingFor Then
                                            found = True
                                        End If

                                    End If
                                Else
                                    Dim curVariantName As String = getVariantNameOfTreeNode(CStr(variantListe.Item(i + 1)))

                                    If firstTime Then
                                        nameOfFirstChecked = curVariantName
                                        firstTime = False
                                    End If

                                    If curVariantName = variantNameLookingFor Then
                                        found = True
                                    End If
                                End If

                            Next

                            Dim showVariantName As String
                            If found Then
                                showVariantName = variantNameLookingFor
                            Else
                                showVariantName = nameOfFirstChecked
                            End If


                            For v = 1 To anzahlVarianten

                                variantName = getVariantNameOfTreeNode(CStr(variantListe.Item(v)))

                                variantNode = Nothing
                                If projektNode.Nodes.Count > 0 Then
                                    variantNode = projektNode.Nodes.Item(v - 1)
                                End If



                                Dim weitermachen As Boolean = False

                                If Not IsNothing(variantNode) Then
                                    weitermachen = variantNode.Checked
                                Else
                                    weitermachen = True
                                End If

                                If weitermachen Then
                                    Dim showAttribute As Boolean
                                    If IsNothing(hproj) Then
                                        showAttribute = (variantName = showVariantName)
                                    Else
                                        If variantName = hproj.variantName Then
                                            showAttribute = True
                                        Else
                                            showAttribute = False
                                        End If
                                    End If

                                    ' laden der Projekt-Variante 
                                    ' wenn gefiltert wird, dann wird pfv geladen als als Planungs-Version in AllePRojekte gesteckt 

                                    Call loadProjectfromDB(outPutCollection, pname, variantName, showAttribute, storedAtOrBefore, calledFromPPT)

                                    ' das für Powerpoint ausgewählte Projekt 
                                    If aKtionskennung = PTTvActions.loadPVInPPT Then
                                        pptPname = pname
                                        pptVname = variantName
                                    Else
                                        ' in load wird das pfv als Basis-Variante abgespeichert , deswegen muss jetzt variantName der Basis-Varianten-Name sein
                                        If awinSettings.filterPFV And variantName = ptVariantFixNames.pfv.ToString Then
                                            variantName = ""
                                        End If
                                    End If


                                    If currentBrowserConstellation.contains(calcProjektKey(pname, variantName), False) Then
                                        ' nichts tun , ist schon drin 
                                        currentBrowserConstellation.getItem(calcProjektKey(pname, variantName)).show = showAttribute
                                    Else
                                        Dim cItem As New clsConstellationItem
                                        ' tk 28.12.18 , um nachher das Attribut setzen zu können
                                        Dim tmpProj As clsProjekt = getProjektFromSessionOrDB(pname, variantName, AlleProjekte, Date.Now)

                                        With cItem
                                            .projectName = pname
                                            .variantName = variantName
                                            .show = showAttribute
                                            If Not IsNothing(tmpProj) Then
                                                .projectTyp = CType(tmpProj.projectType, ptPRPFType).ToString
                                            End If
                                        End With
                                        currentBrowserConstellation.add(cItem)
                                    End If
                                End If



                            Next


                        ElseIf aKtionskennung = PTTvActions.loadPVInPPT Or aKtionskennung = PTTvActions.loadMultiPVInPPT Then

                            Dim hproj As clsProjekt = Nothing
                            If ShowProjekte.Count > 0 Then
                                If ShowProjekte.contains(pname) Then
                                    hproj = ShowProjekte.getProject(pname)
                                End If
                            End If


                            ' da hier manuell Projekte hinzu kommen, muss der Sort-Type auf customTF gesetzt werden 
                            currentBrowserConstellation.sortCriteria = ptSortCriteria.customTF
                            currentSessionConstellation.sortCriteria = ptSortCriteria.customTF


                            Dim variantNameLookingFor As String = ""


                            ' jetzt muss geprüft werden, welcher Name tatsächlich ins Show gesteckt werden soll 
                            Dim nameOfFirstChecked As String = ""
                            Dim firstTime As Boolean = True
                            Dim found As Boolean = False


                            For i As Integer = 0 To anzahlVarianten - 1

                                If projektNode.Nodes.Count > 0 Then
                                    If projektNode.Nodes.Item(i).Checked = True Then

                                        Dim curVariantName As String = getVariantNameOfTreeNode(projektNode.Nodes.Item(i).Text)

                                        If firstTime Then
                                            nameOfFirstChecked = curVariantName
                                            firstTime = False
                                        End If

                                        If curVariantName = variantNameLookingFor Then
                                            found = True
                                        End If

                                    End If
                                Else
                                    Dim curVariantName As String = getVariantNameOfTreeNode(CStr(variantListe.Item(i + 1)))

                                    If firstTime Then
                                        nameOfFirstChecked = curVariantName
                                        firstTime = False
                                    End If

                                    If curVariantName = variantNameLookingFor Then
                                        found = True
                                    End If
                                End If

                            Next

                            Dim showVariantName As String
                            If found Then
                                showVariantName = variantNameLookingFor
                            Else
                                showVariantName = nameOfFirstChecked
                            End If


                            For v = 1 To anzahlVarianten

                                variantName = getVariantNameOfTreeNode(CStr(variantListe.Item(v)))

                                variantNode = Nothing
                                If projektNode.Nodes.Count > 0 Then
                                    variantNode = projektNode.Nodes.Item(v - 1)
                                End If



                                Dim weitermachen As Boolean = False

                                If Not IsNothing(variantNode) Then
                                    weitermachen = variantNode.Checked
                                Else
                                    weitermachen = True
                                End If

                                If weitermachen Then
                                    Dim showAttribute As Boolean
                                    If IsNothing(hproj) Then
                                        showAttribute = (variantName = showVariantName)
                                    Else
                                        If variantName = hproj.variantName Then
                                            showAttribute = True
                                        Else
                                            showAttribute = False
                                        End If
                                    End If

                                    ' laden der Projekt-Variante 
                                    ' wenn gefiltert wird, dann wird pfv geladen als als Planungs-Version in AllePRojekte gesteckt 

                                    Call loadProjectfromDB(outPutCollection, pname, variantName, showAttribute, storedAtOrBefore, calledFromPPT)

                                    pptPname = pname
                                    pptVname = variantName


                                    If currentBrowserConstellation.contains(calcProjektKey(pname, variantName), False) Then
                                        ' nichts tun , ist schon drin 
                                        currentBrowserConstellation.getItem(calcProjektKey(pname, variantName)).show = showAttribute
                                    Else
                                        Dim cItem As New clsConstellationItem
                                        ' tk 28.12.18 , um nachher das Attribut setzen zu können
                                        Dim tmpProj As clsProjekt = getProjektFromSessionOrDB(pname, variantName, AlleProjekte, Date.Now)

                                        With cItem
                                            .projectName = pname
                                            .variantName = variantName
                                            .show = showAttribute
                                            If Not IsNothing(tmpProj) Then
                                                .projectTyp = CType(tmpProj.projectType, ptPRPFType).ToString
                                            End If
                                        End With
                                        currentBrowserConstellation.add(cItem)
                                    End If
                                End If



                            Next


                        End If



                    ElseIf projektNode.Tag = "X" And aKtionskennung <> PTTvActions.loadPVInPPT And
                        aKtionskennung <> PTTvActions.loadMultiPVInPPT Then

                        anzahlVarianten = projektNode.Nodes.Count
                        Dim first As Boolean = True

                        For v = 1 To anzahlVarianten
                            variantNode = projektNode.Nodes.Item(v - 1)
                            variantName = getVariantNameOfTreeNode(variantNode.Text)


                            If variantNode.Checked Then
                                ' Aktion auf allen Timestamps
                                ' lösche in Datenbank das Objekt mit DB-Namen pname#vname

                                If aKtionskennung = PTTvActions.delFromDB Then

                                    ' Fehler Check, ob in Szenario refernziert , passiert in der Routine 
                                    Call deleteCompleteProjectVariant(outPutCollection,
                                                                      pname, variantName, aKtionskennung)



                                ElseIf aKtionskennung = PTTvActions.delAllExceptFromDB Then

                                    ' hier muss ja gar kein Check auf Szenario referenz erfolgen, da ohnehin immer min 2 Stände behalten werdne  
                                    'Call deleteCompleteProjectVariant(outPutCollection,
                                    '                                  pname, variantName, aKtionskennung, versionsToKeep.Value)

                                ElseIf aKtionskennung = PTTvActions.delFromSession Or
                                        aKtionskennung = PTTvActions.deleteV Then

                                    Call awinDeleteProjectInSession(pName:=pname, considerDependencies:=considerDependencies, vName:=variantName)

                                    ' jetzt in der currentBrowserConstellation ändern 
                                    Dim tmpKey As String = calcProjektKey(pname, variantName)
                                    currentBrowserConstellation.remove(tmpKey)


                                ElseIf aKtionskennung = PTTvActions.loadPV Or aKtionskennung = PTTvActions.loadPVInPPT Or
                                                        aKtionskennung = PTTvActions.loadMultiPVInPPT Then

                                    ' da hier manuell Projekte hinzu kommen, muss der Sort-Type auf customTF gesetzt werden 
                                    currentBrowserConstellation.sortCriteria = ptSortCriteria.customTF
                                    currentSessionConstellation.sortCriteria = ptSortCriteria.customTF

                                    Call loadProjectfromDB(outPutCollection, pname, variantName, first, storedAtOrBefore, calledFromPPT)

                                    ' das für Powerpoint ausgewählte Projekt 
                                    If aKtionskennung = PTTvActions.loadPVInPPT Then
                                        pptPname = pname
                                        pptVname = variantName
                                    End If

                                    first = False

                                    If currentBrowserConstellation.contains(calcProjektKey(pname, variantName), False) Then
                                        ' nichts tun , ist schon drin 
                                    Else
                                        Dim cItem As New clsConstellationItem

                                        ' tk 28.12.18 , um nachher das Attribut setzen zu können
                                        Dim tmpProj As clsProjekt = getProjektFromSessionOrDB(pname, variantName, AlleProjekte, Date.Now)

                                        With cItem
                                            .projectName = pname
                                            .variantName = variantName
                                            If Not IsNothing(tmpProj) Then
                                                .projectTyp = CType(tmpProj.projectType, ptPRPFType).ToString
                                            End If
                                            .show = (v = 1)
                                        End With
                                        currentBrowserConstellation.add(cItem)
                                    End If

                                End If


                            ElseIf aKtionskennung = PTTvActions.delFromDB Or
                                    aKtionskennung = PTTvActions.loadPVS Then

                                anzahlTimeStamps = variantNode.Nodes.Count
                                Dim firstTS As Boolean = True
                                For t = 1 To anzahlTimeStamps
                                    timeStampNode = variantNode.Nodes.Item(t - 1)

                                    If timeStampNode.Checked Then
                                        ' Aktion auf diesem timestamp

                                        timestamp = CType(timeStampNode.Text, Date)
                                        If aKtionskennung = PTTvActions.delFromDB Then
                                            Call deleteProjectVariantTimeStamp(outPutCollection,
                                                                               pname, variantName, timestamp, firstTS)
                                        Else
                                            ' Aktion für LoadPVS : aber hier gibt es wahrscheinlich gar keinen OK-Button
                                        End If

                                    End If
                                Next
                            End If

                        Next
                    End If

                Next


                ' tk 7.10.19
                ' damit dieses Formular weiderverwendbar ist, müssen die Projectboard spezifischen Sachen raus. 

                'If aKtionskennung = PTTvActions.loadPV Or
                '    aKtionskennung = PTTvActions.delFromSession Then
                '    Call awinNeuZeichnenDiagramme(2)
                'End If

            End With

            ' bei Aktionen loadPV, delFromSession muss der currentConstellationName aktualisiert werden 
            If aKtionskennung = PTTvActions.delFromSession Or
                aKtionskennung = PTTvActions.loadPV Or
                aKtionskennung = PTTvActions.loadPVInPPT Or
                aKtionskennung = PTTvActions.loadMultiPVInPPT Or
                aKtionskennung = PTTvActions.deleteV Then
                If currentConstellationName <> calcLastSessionScenarioName() Then
                    currentConstellationName = calcLastSessionScenarioName()
                End If

                'Call storeSessionConstellation("Last")
            End If


            ' jetzt ggf die Outputs anzeigen 
            If outPutCollection.Count > 0 Then
                Call showOutPut(outPutCollection,
                                outPutHeader,
                                outPutExplanation)
            End If

            ' Cursor auf Normal-Cursor setzen ... 
            Me.Cursor = Cursors.Arrow

            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()

        ElseIf aKtionskennung = PTTvActions.loadProjectAsTemplate Then

            With TreeViewProjekte

                anzahlProjekte = .Nodes.Count

                For p = 1 To anzahlProjekte
                    projektNode = .Nodes.Item(p - 1)

                    If projektNode.Checked Then
                        pname = getProjectNameOfTreeNode(projektNode.Text)
                        variantName = ""

                        anzahlVarianten = projektNode.Nodes.Count
                        For v = 1 To anzahlVarianten
                            variantNode = projektNode.Nodes.Item(v - 1)
                            If variantNode.Checked Then
                                variantName = getVariantNameOfTreeNode(variantNode.Text)
                            End If
                        Next

                        ' checken, ob es existiert, sonst weitermachen 
                        selProjectAsTemplate = AlleProjekte.getProject(pname, variantName)
                        If Not IsNothing(selProjectAsTemplate) Then
                            Exit For
                        End If

                    End If
                Next

                ' Cursor auf Normal-Cursor setzen ... 
                Me.Cursor = Cursors.Arrow

                DialogResult = Windows.Forms.DialogResult.OK
                MyBase.Close()
            End With



        ElseIf aKtionskennung = PTTvActions.chgInSession Then

            If dropboxScenarioNames.Text <> "" Then


                currentConstellationName = dropboxScenarioNames.Text
                'currentBrowserConstellation.constellationName = currentConstellationName

                Dim toStoreConstellation As clsConstellation =
                    currentBrowserConstellation.copy(currentConstellationName)


                ' Korrektheitsprüfung
                ' testen 

                If awinSettings.visboDebug Then
                    toStoreConstellation.checkAndCorrectYourself()
                End If

                ' hier war vorher .update
                ' jetzt muss die Constellation upgedated werden ... 
                ' hier muss 
                Dim budget As Double = projectConstellations.getBudgetOfLoadedPortfolios
                projectConstellations.update(toStoreConstellation)

                Dim txtMsg1 As String = ""
                Dim txtMsg2 As String = ""
                If storeToDBasWell.Checked Then
                    Dim errMsg As New clsErrorCodeMsg
                    'Dim dbConstellations As clsConstellations = CType(databaseAcc, DBAccLayer.Request).retrieveConstellationsFromDB(Date.Now, errMsg)
                    Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)

                    Call storeSingleConstellationToDB(outPutCollection, toStoreConstellation, dbPortfolioNames)

                    ' jetzt ggf die Outputs anzeigen 

                    If outPutCollection.Count > 0 Then

                        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                            txtMsg1 = "Speichern Portfolio " & toStoreConstellation.constellationName
                            txtMsg2 = "folgende Informationen:"
                        Else
                            txtMsg1 = "Save Portfolio " & toStoreConstellation.constellationName
                            txtMsg2 = "following messages:"
                        End If
                        Call showOutPut(outPutCollection, txtMsg1, txtMsg2)

                    Else
                        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                            txtMsg1 = "ok, " & currentConstellationName & " in Datenbank und Session gespeichert"
                        Else
                            txtMsg1 = "ok, " & currentConstellationName & " stored in Session and database"
                        End If
                        Call MsgBox(txtMsg1)
                    End If
                Else

                    ' jetzt das Union Projekt errechnen ... 
                    ' jetzt muss das Summary Projekt zur Constellation erzeugt und gespeichert werden
                    Try

                        If budget = 0 Then
                            budget = -1
                        End If

                        Dim tmpVariantName As String = ""
                        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                            tmpVariantName = ptVariantFixNames.pfv.ToString
                        End If

                        Dim oldSummaryP As clsProjekt = getProjektFromSessionOrDB(toStoreConstellation.constellationName, tmpVariantName, AlleProjekte, Date.Now)

                        If Not IsNothing(oldSummaryP) Then
                            'budget = oldSummaryP.budgetWerte.Sum
                            budget = oldSummaryP.Erloes
                        Else
                            budget = toStoreConstellation.getBudgetOfShownProjects
                        End If

                        Dim sproj As clsProjekt = calcUnionProject(toStoreConstellation, False, Date.Now.Date.AddHours(23).AddMinutes(59), budget:=budget)

                        Dim skey As String = calcProjektKey(sproj.name, sproj.variantName)
                        If AlleProjekte.Containskey(skey) Then
                            AlleProjekte.Remove(skey)
                        End If

                        If Not AlleProjekte.Containskey(skey) Then
                            AlleProjekte.Add(sproj)
                        End If

                    Catch ex As Exception

                    End Try


                    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                        txtMsg1 = "ok, " & currentConstellationName & " in Session gespeichert"
                    Else
                        txtMsg1 = "ok, " & currentConstellationName & " stored in Session"
                    End If
                    Call MsgBox(txtMsg1)
                End If

                ' jetzt das EIngabe Feld wieder zurücksetzen 
                dropboxScenarioNames.Text = ""


            End If


            ' im Fesnter bleiben ... 
            'DialogResult = Windows.Forms.DialogResult.OK
            'MyBase.Close()
        Else
            Dim txtMsg As String = ""
            If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                txtMsg = "nicht unterstützte Option in ProjPortfolioAdmin Formular ..."
            Else
                txtMsg = "not supported option in form ProjPortfolioAdmin ..."
            End If

            Call MsgBox(txtMsg)

        End If

        ' jetzt muss das für Projectboard bzw. Aufruf von Powerpoint gestezt werden 

        If aKtionskennung <> PTTvActions.loadProjectAsTemplate Then
            ' jetzt muss die Caption neu gesetzt werden ...
            If Not IsNothing(projectboardWindows(PTwindows.mpt)) Then
                Try
                    projectboardWindows(PTwindows.mpt).Caption = bestimmeWindowCaption(PTwindows.mpt)
                Catch ex As Exception

                End Try
            End If
        End If

        ' tk 16.11.20 braucht man nicht mehr ..
        'If aKtionskennung = PTTvActions.loadPVInPPT Then
        '    ' selectedProjekte setzen, die werden nämlich dann in PPT abgefragt
        '    selectedProjekte.Clear(False)
        '    Dim hproj As clsProjekt = AlleProjekte.getProject(pptPname, pptVname)

        '    If Not IsNothing(hproj) Then
        '        selectedProjekte.Add(hproj, False)
        '    End If

        'ElseIf aKtionskennung <> PTTvActions.loadProjectAsTemplate Then
        '    ' jetzt muss die Caption neu gesetzt werden ...
        '    If Not IsNothing(projectboardWindows(PTwindows.mpt)) Then
        '        Try
        '            projectboardWindows(PTwindows.mpt).Caption = bestimmeWindowCaption(PTwindows.mpt)
        '        Catch ex As Exception

        '        End Try
        '    End If
        'End If



        ' Cursor auf Normal-Cursor setzen ... 
        Me.Cursor = Cursors.Arrow


    End Sub


    ''' <summary>
    ''' alle dargestellten Elemente im ProjektTree selektieren 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SelectionSet_Click(sender As Object, e As EventArgs) Handles SelectionSet.Click

        Dim err As New clsErrorCodeMsg

        Dim projectNode As TreeNode

        stopRecursion = True

        ' jetzt im formular den Mauszeiger auf Warten ... setzen 
        Me.Cursor = Cursors.WaitCursor

        With TreeViewProjekte

            ' die Behandlung von chgInSession ist etwas anders, weil sofort eine Aktion erfolgen muss ... 
            If aKtionskennung = PTTvActions.chgInSession Then



                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)
                    Dim pName As String = getProjectNameOfTreeNode(projectNode.Text)

                    ' jetzt muss die Behandlung kommen, was denn gemacht werden soll 
                    ' ############ ChgInSession ####################################

                    ' das Projekt muss in Showprojekte, aber nur wenn es nicht bereits gecheckt war 
                    Dim variantName As String = ""

                    If Not projectNode.Checked Then
                        projectNode.Checked = True

                        ' ermittle die gecheckte Variante 
                        Dim checkedVariants As Collection = getNamesOfChildNodes(projectNode, True)
                        If checkedVariants.Count = 0 Then
                            ' dann muss der Varianten-Name entsprechend gesetzt werden  
                            Dim tmpCollection As Collection = AlleProjekte.getVariantNames(pName, True)
                            If tmpCollection.Count > 0 Then
                                variantName = getVariantNameOfTreeNode(CStr(tmpCollection.Item(1)))
                            Else
                                variantName = ""
                            End If


                        ElseIf checkedVariants.Count = 1 Then
                            variantName = getVariantNameOfTreeNode(CStr(checkedVariants.Item(1)))

                        ElseIf checkedVariants.Count > 1 Then
                            variantName = getVariantNameOfTreeNode(CStr(checkedVariants.Item(1)))
                            For k As Integer = 1 To projectNode.Nodes.Count
                                If getVariantNameOfTreeNode(projectNode.Nodes.Item(k - 1).Text) = variantName Then
                                    projectNode.Nodes.Item(k - 1).Checked = True
                                Else
                                    projectNode.Nodes.Item(k - 1).Checked = False
                                End If
                            Next
                        End If

                        ' jetzt muss das Show-Attribut entsprechend gesetzt werden 
                        Call currentBrowserConstellation.updateShowAttributes(pName, variantName, True)


                        ' jetzt muss das Projekt in Showprojekte eingetragen werden bzw. das alte zuvor gelöscht werden 
                        If ShowProjekte.contains(pName) Then
                            ShowProjekte.Remove(pName)
                        End If

                        Dim key As String = calcProjektKey(pName, variantName)
                        Dim hproj As clsProjekt = AlleProjekte.getProject(key)

                        ShowProjekte.Add(hproj)

                    Else
                        ' nichts tun , denn das Projekt wird bereits angezeigt und ist in Showprojekte drin 
                    End If

                Next

                ' jetzt muss die Plan-Tafel gelöscht werden 
                Call awinClearPlanTafel()

                ' jetzt muss die Plan-Tafel neu gezeichnet werden 
                'Call awinZeichnePlanTafelNeu(True)
                Call awinZeichnePlanTafel(True)

                ' jetzt müssen die Diagramme neu gezeichnet werden 
                Call awinNeuZeichnenDiagramme(2)


            ElseIf aKtionskennung = PTTvActions.deleteV Or
                aKtionskennung = PTTvActions.activateV Then
                ' nichts tun, Alle Selektieren macht bei diesen keinen Sinn 

            ElseIf aKtionskennung = PTTvActions.delFromDB Then
                ' wenn ein Stand angegeben ist , dann sollen alle mit diesem Stand markiert werden 
                'If Not IsNothing(dropBoxTimeStamps.SelectedItem) Then

                '    Dim lookForTimestamp As Date = CDate(dropBoxTimeStamps.SelectedItem)
                '    Dim vergleichsString As String = lookForTimestamp.ToString

                '    ' jetzt wird der TreeView komplett expanded ...
                '    stopRecursion = False
                '    .ExpandAll()
                '    stopRecursion = True

                '    For i As Integer = 1 To .Nodes.Count
                '        projectNode = .Nodes.Item(i - 1)

                '        For v As Integer = 1 To projectNode.Nodes.Count
                '            Dim variantNode As TreeNode = projectNode.Nodes.Item(v - 1)

                '            For t As Integer = 1 To variantNode.Nodes.Count
                '                Dim tsNode As TreeNode = variantNode.Nodes.Item(t - 1)
                '                If tsNode.Text = vergleichsString Then
                '                    tsNode.Checked = True
                '                    variantNode.Checked = False
                '                    projectNode.Checked = False

                '                    If Not projectNode.IsExpanded Then
                '                        projectNode.Expand()
                '                    End If

                '                    If Not variantNode.IsExpanded Then
                '                        variantNode.Expand()
                '                    End If
                '                End If
                '            Next
                '        Next

                '    Next
                'Else
                Dim txtMsg As String = ""
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    txtMsg = "beim Löschen kann nicht alles selektiert werden ..."
                Else
                    txtMsg = "not allowed to select all when deleting ..."
                End If
                Call MsgBox(txtMsg)

            ElseIf aKtionskennung = PTTvActions.setWriteProtection Then
                ' für jeden Knoten prüfen, ob er bereits geschützt ist
                ' dann nichts machen 
                ' andernfalls prüfen, ob er von mir geschützt werden kann 
                ' wenn ja, dann schützen 

                'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)

                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)

                    If projectNode.Checked Then
                        ' nichts machen, ist ja schon gecheckt / geschützt 
                    Else
                        ' noch nicht geschützt
                        ' jetzt prüfen, ob man es überhaupt schützen darf 
                        Dim pName As String = getProjectNameOfTreeNode(projectNode.Text)
                        Dim vName As String = ""
                        ' holt die Varianten-Namen ohne Klammer ... 
                        Dim variantNames As Collection = AlleProjekte.getVariantNames(pName, False)
                        If variantNames.Count > 0 Then
                            vName = CStr(variantNames.Item(1))
                        End If

                        Dim atLeastOneFailed As Boolean = False

                        ' hier prüfen, ob alle Childs geschützt werden können ... 
                        If projectNode.Nodes.Count > 0 Then
                            ' alle Varianten schützen, die man schützen kann 
                            For iv As Integer = 1 To projectNode.Nodes.Count
                                Dim variantNode As TreeNode = projectNode.Nodes.Item(iv - 1)
                                vName = getVariantNameOfTreeNode(variantNode.Text)

                                If setNodeWriteProtections(variantNode, PTTreeNodeTyp.pVariant, pName, vName, True) Then
                                    ' erfolgreich ..
                                    ' es wurde bereits Node Apperance inkl Check-Status geklärt
                                Else
                                    ' nicht zugelassen , also nichts machen  
                                    writeProtections.upsert(CType(databaseAcc, DBAccLayer.Request).getWriteProtection(pName, vName, err))
                                    Call bestimmeNodeAppearance(variantNode, aKtionskennung, PTTreeNodeTyp.pVariant, pName, vName)

                                    atLeastOneFailed = True
                                End If

                            Next

                            ' jetzt muss noch die Behandlung für das Projekt selber kommen 
                            ' dazu reicht aber, die NodeAppearance zu setzen ..
                            If atLeastOneFailed Then
                                projectNode.Checked = False
                            Else
                                projectNode.Checked = True
                            End If
                            ' vname ist hier nicht wichtig ... 
                            Call bestimmeNodeAppearance(projectNode, aKtionskennung, PTTreeNodeTyp.project, pName, "")

                        Else
                            ' es gibt keine Childs 
                            ' keine Varianten im Baum , aber in variantNames muss mindestens ein Element sein 
                            If setNodeWriteProtections(projectNode, PTTreeNodeTyp.project, pName, vName, True) Then
                                ' erfolgreich ..
                                projectNode.Checked = True
                            Else
                                ' nicht zugelassen , also nichts machen  
                                writeProtections.upsert(CType(databaseAcc, DBAccLayer.Request).getWriteProtection(pName, vName, err))
                                Call bestimmeNodeAppearance(projectNode, aKtionskennung, PTTreeNodeTyp.project, pName, vName)

                            End If

                        End If
                    End If



                Next


            Else
                ' in allen anderen Fällen: loadPV, loadPVS, delAllExceptFromDB, delFromSession

                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)
                    If Not projectNode.Checked Then
                        projectNode.Checked = True
                    End If
                    If projectNode.Nodes.Count > 0 Then
                        Call Check(projectNode)
                    End If
                Next

            End If

        End With

        If aKtionskennung = PTTvActions.chgInSession Or
            aKtionskennung = PTTvActions.activateV Then

            If currentConstellationName <> calcLastSessionScenarioName() Then
                currentConstellationName = calcLastSessionScenarioName()

                Dim preText As String = "Portfolio "
                If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                    preText = "Portfolio "
                End If

                Me.Text = preText & currentConstellationName
            End If

        End If

        Me.Cursor = Cursors.Default
        stopRecursion = False

    End Sub

    ''' <summary>
    ''' setzt oder released den Schreibschutz - aber nur wenn der Nutzer das überhaupt darf 
    ''' </summary>
    ''' <param name="tmpNode">der betroffene Knoten</param>
    ''' <param name="pName">der Projekt-Name</param>
    ''' <param name="vName">der Varianten-NAme</param>
    ''' <param name="writeProtect">true, wenn geschützt werden soll
    ''' false, wenn der Schutz aufgehoben werden soll</param>
    ''' <remarks></remarks>
    Private Function setNodeWriteProtections(ByRef tmpNode As TreeNode, ByVal treeNodeType As Integer,
                                             ByVal pName As String, ByVal vName As String,
                                             ByVal writeProtect As Boolean) As Boolean

        Dim err As New clsErrorCodeMsg

        Dim pvName As String = calcProjektKey(pName, vName)

        Dim wpItem As New clsWriteProtectionItem(pvName, ptWriteProtectionType.project,
                                                dbUsername, Me.chkbxPermanent.Checked, writeProtect)

        If CType(databaseAcc, DBAccLayer.Request).setWriteProtection(wpItem, err) Then
            ' alles in Ordnung : es ist jetzt geschützt bzw. released
            ' dann checken, dann in WriteProtections aktualisieren, dann Appearance setzen ...
            tmpNode.Checked = writeProtect
            Call writeProtections.upsert(wpItem)
            Call bestimmeNodeAppearance(tmpNode, aKtionskennung, treeNodeType, pName, vName)
            setNodeWriteProtections = True
        Else
            Select Case err.errorCode
                Case 409
                    If awinSettings.englishLanguage Then
                        Call MsgBox("project is protected by another user")
                    Else
                        Call MsgBox("Projekt ist bereits von einem anderen User geschützt ")
                    End If

                Case Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Error : Protection doesn't work")
                    Else
                        Call MsgBox("Fehler : der Schutz/die Veränderung ist schiefgegangen")
                    End If

            End Select

            setNodeWriteProtections = False
        End If
    End Function

    ''' <summary>
    ''' gibt die Namen der Kind-Knoten wieder, 
    ''' alle
    ''' nur die, die gecheckt sind  
    ''' nur die, die nicht gecheckt sind 
    ''' </summary>
    ''' <param name="curNode">der aktuelle Knoten </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getNamesOfChildNodes(ByVal curNode As TreeNode, ByVal checkState As Boolean, Optional considerAll As Boolean = False) As Collection
        Dim tmpCollection As New Collection

        Dim childNode As TreeNode

        With curNode

            For i As Integer = 1 To .Nodes.Count

                childNode = .Nodes.Item(i - 1)

                If considerAll Then
                    tmpCollection.Add(childNode.Name)
                Else
                    If childNode.Checked = checkState Then
                        tmpCollection.Add(childNode.Name)
                    End If
                End If
            Next

        End With

        getNamesOfChildNodes = tmpCollection
    End Function

    ''' <summary>
    ''' setzt alle Knoten im node auf checked
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

    ''' <summary>
    ''' unchecks all nodes in TreeView Structure, excep node with Name nodeName
    ''' </summary>
    ''' <param name="nodeName"></param>
    Private Sub uncheckExcept(ByVal nodeName As String)
        For i As Integer = 1 To TreeViewProjekte.Nodes.Count
            Dim tmpNode As TreeNode = TreeViewProjekte.Nodes.Item(i - 1)
            If tmpNode.Level = 0 And tmpNode.Name <> nodeName Then
                If tmpNode.Checked Then
                    tmpNode.Checked = False
                    ' dann auch alle Varianten unchecken ... 
                    Dim anzV As Integer = tmpNode.Nodes.Count

                    For vi As Integer = 1 To anzV
                        If tmpNode.Nodes.Item(vi - 1).Checked Then
                            tmpNode.Nodes.Item(vi - 1).Checked = False
                        End If
                    Next
                End If
            End If
        Next
    End Sub


    Private Sub expandCompletely_Click(sender As Object, e As EventArgs) Handles expandCompletely.Click

        With TreeViewProjekte
            .Cursor = Cursors.WaitCursor
            .Visible = False
            .ExpandAll()
            .Visible = True
            .Cursor = Cursors.Default
        End With

    End Sub

    Private Sub collapseCompletely_Click(sender As Object, e As EventArgs) Handles collapseCompletely.Click
        With TreeViewProjekte
            .CollapseAll()
        End With
    End Sub

    ''' <summary>
    ''' de-selektiert alle Knoten im Formular ProjektPortfolio 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SelectionReset_Click(sender As Object, e As EventArgs) Handles SelectionReset.Click

        Dim err As New clsErrorCodeMsg

        Dim projectNode As TreeNode

        stopRecursion = True

        Me.Cursor = Cursors.WaitCursor

        With TreeViewProjekte

            ' die Behandlung von chgInSession ist etwas anders, weil sofort eine Aktion erfolgen muss ... 

            If aKtionskennung = PTTvActions.chgInSession Then

                Try
                    For i As Integer = 1 To .Nodes.Count
                        projectNode = .Nodes.Item(i - 1)
                        Dim pName As String = getProjectNameOfTreeNode(projectNode.Text)

                        ' jetzt muss die Behandlung kommen, was denn gemacht werden soll 
                        ' ############ ChgInSession ####################################

                        ' das Projekt muss in Showprojekte, aber nur wenn es nicht bereits gecheckt war 
                        Dim variantName As String = ""

                        If projectNode.Checked Then

                            projectNode.Checked = False

                            ' jetzt müssen die Show Attribute und die Zeilen neu gesetzt werden ...
                            currentBrowserConstellation.updateShowAttributes(pName, Nothing, False)

                        Else
                            ' nichts tun , denn das Projekt wird bereits angezeigt und ist in Showprojekte drin 
                        End If

                    Next

                    ' jetzt muss die Plan-Tafel gelöscht werden 
                    Call awinClearPlanTafel()

                    ' jetzt muss Showprojekte gelöscht werden 
                    ShowProjekte.Clear()


                    ' jetzt müssen die Diagramme neu gezeichnet werden 
                    Call awinNeuZeichnenDiagramme(2)

                Catch ex As Exception
                    Dim a As Integer = 0
                    Call MsgBox("Fehler: " & ex.Message)
                End Try


            ElseIf aKtionskennung = PTTvActions.activateV Then
                ' nichts tun, Alle Resetten macht bei diesen keinen Sinn 

            ElseIf aKtionskennung = PTTvActions.setWriteProtection Then
                ' für jeden Knoten prüfen, ob der Schutz aufgehoben ist 
                ' dann nichts machen 
                ' andernfalls prüfen, ob der Schutz von mir aufgehoben werden kann 
                ' wenn ja, dann aufheben 

                'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)

                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)

                    ' es muss nur betrachtet werden, wenn der Knoten gesetzt ist 
                    ' oder der Knoten Kinder hat , die ihrerseits geschützt sein können 
                    If projectNode.Checked Or projectNode.Nodes.Count > 0 Then

                        Dim pName As String = getProjectNameOfTreeNode(projectNode.Text)
                        Dim vName As String = ""
                        ' holt die Varianten-Namen ohne Klammer ... 
                        Dim variantNames As Collection = AlleProjekte.getVariantNames(pName, False)
                        If variantNames.Count > 0 Then
                            vName = CStr(variantNames.Item(1))
                        End If

                        Dim atLeastOneFailed As Boolean = False

                        ' hier prüfen, ob alle Childs aufgehoben werden können ... 
                        If projectNode.Nodes.Count > 0 Then
                            ' alle Varianten aufheben, die man aufheben kann 
                            For iv As Integer = 1 To projectNode.Nodes.Count

                                Dim variantNode As TreeNode = projectNode.Nodes.Item(iv - 1)

                                If variantNode.Checked Then
                                    ' nur dann kann ein Schutz aufgehoben werden ...
                                    vName = getVariantNameOfTreeNode(variantNode.Text)
                                    If setNodeWriteProtections(variantNode, PTTreeNodeTyp.pVariant, pName, vName, False) Then
                                        ' erfolgreich aufgehoben ..
                                        ' es wurde bereits Node Apperance inkl Check-Status geklärt
                                    Else
                                        ' Aufheben nicht zugelassen , also nichts machen  
                                        writeProtections.upsert(CType(databaseAcc, DBAccLayer.Request).getWriteProtection(pName, vName, err))
                                        Call bestimmeNodeAppearance(variantNode, aKtionskennung, PTTreeNodeTyp.pVariant, pName, vName)

                                        atLeastOneFailed = True
                                    End If

                                End If


                            Next

                            ' jetzt muss noch die Behandlung für das Projekt selber kommen 
                            ' dazu reicht aber, die NodeAppearance zu setzen ..
                            If projectNode.Checked Then

                                Dim allChildsAreProtected As Boolean = (projectNode.Nodes.Count > 0)
                                For ti = 1 To projectNode.Nodes.Count
                                    allChildsAreProtected = allChildsAreProtected And projectNode.Nodes.Item(ti - 1).Checked
                                Next
                                If Not allChildsAreProtected Then
                                    projectNode.Checked = False
                                End If
                            End If

                            ' vname ist hier nicht wichtig ... 
                            Call bestimmeNodeAppearance(projectNode, aKtionskennung, PTTreeNodeTyp.project, pName, "")

                        Else
                            ' es gibt keine Childs 
                            ' keine Varianten im Baum , aber in variantNames muss mindestens ein Element sein 
                            If setNodeWriteProtections(projectNode, PTTreeNodeTyp.project, pName, vName, False) Then
                                ' erfolgreich ..
                            Else
                                ' nicht zugelassen , also nichts machen  
                                writeProtections.upsert(CType(databaseAcc, DBAccLayer.Request).getWriteProtection(pName, vName, err))
                                Call bestimmeNodeAppearance(projectNode, aKtionskennung, PTTreeNodeTyp.project, pName, vName)
                            End If


                        End If
                        'End If

                    Else
                        ' es muss nichts gemacht werden ... 
                    End If




                Next

            Else
                ' auch in den Fällen deleteV
                ' in allen anderen Fällen: loadPV, loadPViInPPTloadPVS, delFromDB, delAllExceptFromDB, delFromSession
                ' alle de-selektieren; h 
                Call deSelectNodes()
                'For i As Integer = 1 To .Nodes.Count
                '    projectNode = .Nodes.Item(i - 1)
                '    If projectNode.Checked Then
                '        projectNode.Checked = False
                '    End If

                '    If projectNode.Nodes.Count > 0 Then
                '        Call unCheck(projectNode)
                '    End If
                'Next

            End If

        End With

        If aKtionskennung = PTTvActions.chgInSession Or
            aKtionskennung = PTTvActions.activateV Then

            If currentConstellationName <> calcLastSessionScenarioName() Then
                currentConstellationName = calcLastSessionScenarioName()

                Dim preText As String = "Portfolio "
                If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                    preText = "Portfolio "
                End If
                Me.Text = preText & currentConstellationName
            End If

        End If

        Me.Cursor = Cursors.Default
        stopRecursion = False


    End Sub

    ''' <summary>
    ''' de-selektiert alle Knoten 
    ''' ausser dem Knoten, dessen node.fullpath angegeben ist ; 
    ''' optional werden alle auf unchecked gesetzt
    ''' </summary>
    Private Sub deSelectNodes(ByVal Optional exceptNodeFullPath As String = "")

        Dim formerStopRecursion As Boolean = stopRecursion
        stopRecursion = True
        Dim projectNode As TreeNode

        With TreeViewProjekte

            If exceptNodeFullPath = "" Then
                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)

                    If projectNode.Checked Then
                        projectNode.Checked = False
                    End If

                    If projectNode.Nodes.Count > 0 Then
                        Call unCheck(projectNode)
                    End If
                Next
            Else
                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)

                    If projectNode.Checked And projectNode.FullPath <> exceptNodeFullPath Then
                        projectNode.Checked = False
                    End If

                    If projectNode.Nodes.Count > 0 Then
                        Call unCheck(projectNode)
                    End If

                Next
            End If

        End With

        stopRecursion = formerStopRecursion

    End Sub


    ''' <summary>
    ''' reduziert anhand des definierten Filters die aktuelle Gesamtliste und baut den Treeview wieder auf 
    ''' 
    '''
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub filterIcon_Click(sender As Object, e As EventArgs) Handles filterIcon.Click

        Dim err As New clsErrorCodeMsg

        'Dim filterFormular As New frmNameSelection
        Dim filterFormular As New frmHierarchySelection
        Dim considerDependencies As Boolean
        Dim zeitraumVon As Date = StartofCalendar
        Dim zeitraumBis As Date = StartofCalendar.AddYears(20)
        Dim storedGestern As Date = StartofCalendar

        ' hier ist der einzige Grund für browserAlleProjekte: es muss etwas da sein, wo reingeladen werden kann 
        ' wenn auf der Datenbank gefiltert werden soll - und das geht nur , in dem etwas geladen wird ... 
        Dim browserAlleProjekte As New clsProjekteAlle


        awinSettings.useHierarchy = True

        If currentConstellationName <> calcLastSessionScenarioName() Then
            currentConstellationName = calcLastSessionScenarioName()

            Dim preText As String = "Portfolio "
            If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                preText = "Portfolio "
            End If
            Me.Text = preText & currentConstellationName
        End If


        If IsNothing(beforeFilterConstellation) Then
            beforeFilterConstellation = currentBrowserConstellation.copy()
        End If

        Dim storedAtOrBefore As Date
        If IsNothing(requiredDate) Then
            storedAtOrBefore = Date.Now.Date.AddHours(23).AddMinutes(59)
        Else
            storedAtOrBefore = requiredDate.Value
        End If

        If allDependencies.projectCount > 0 Then
            considerDependencies = True
        Else
            considerDependencies = False
        End If

        Me.Cursor = Cursors.WaitCursor

        ' jetzt erst mal überprüfen, ob quicklist = true ..
        If quickList Or
            aKtionskennung = PTTvActions.delFromDB Or
            aKtionskennung = PTTvActions.delAllExceptFromDB Or
            aKtionskennung = PTTvActions.loadPV Then

            If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                ' es ist ein Zeitraum definiert 
                zeitraumVon = getDateofColumn(showRangeLeft, False)
                zeitraumBis = getDateofColumn(showRangeRight, True)
            End If
            ' es muss die Gesamtliste aufgebaut werden ... das dauert jetzt erst mal 
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            'Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

            Dim pname As String = ""
            Dim variantName As String = ""

            'jetzt wird die aktuelleGesamtListe aufgebaut; sobald die mal aufgebaut wurde, muss sie nicht wieder aufgebaut werden ... 
            ' tk das applyFilter wird nachher gemacht , ausnahmslos für alle 
            If Not browserAlleProjekte.Count = 0 Then
                browserAlleProjekte.Clear(False)
            End If

            If awinSettings.loadPFV Or (awinSettings.filterPFV And aKtionskennung = PTTvActions.loadPV) Then
                variantName = ptVariantFixNames.pfv.ToString
            End If
            browserAlleProjekte.liste = CType(databaseAcc, DBAccLayer.Request).retrieveProjectsFromDB(pname, variantName, "", zeitraumVon, zeitraumBis, storedGestern, storedAtOrBefore, True, err)


        Else
            ' browserAlleProjekte bestimmen  
            browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)
        End If

        With filterFormular
            If aKtionskennung = PTTvActions.loadPV Or
                aKtionskennung = PTTvActions.loadPVS Or
                aKtionskennung = PTTvActions.delAllExceptFromDB Or
                aKtionskennung = PTTvActions.delFromDB Then
                ' damit im Filterformular unterschieden werden kann, ob der Aufruf aus dem ProjPortfolioAdmin Formular erfolgte ...
                'tk 9.9.18 
                '.actionCode = aKtionskennung
                .menuOption = PTmenue.filterdefinieren
            Else
                ' tk 9.9.18
                '.actionCode = aKtionskennung
                .menuOption = PTmenue.sessionFilterDefinieren
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then

                stopRecursion = True

                'Me.Cursor = Cursors.WaitCursor
                Dim filter As clsFilter = filterDefinitions.retrieveFilter("Last")
                Dim ok As Boolean

                If aKtionskennung = PTTvActions.loadPV Or
                    aKtionskennung = PTTvActions.delAllExceptFromDB Or
                    aKtionskennung = PTTvActions.delFromDB Then

                    Dim removeList As New Collection


                    For Each kvp As KeyValuePair(Of String, String) In pvNamesList

                        Dim tmpkey As String = kvp.Key
                        If tmpkey.Contains("#") Then
                            ' alles ok 
                        Else
                            tmpkey = calcProjektKey(kvp.Key, "")
                        End If

                        Dim hproj As clsProjekt = browserAlleProjekte.getProject(tmpkey)

                        If Not filter.isEmpty Then
                            If Not IsNothing(hproj) Then
                                ok = filter.doesNotBlock(hproj)
                            Else
                                ok = False
                            End If

                        Else
                            ok = True
                        End If

                        If Not ok Then
                            ' in RemoveListe aufnehmen - diese Projekte werden nachher alle aus aktuelleGesamtliste rausgenommen 
                            Try

                                If Not removeList.Contains(kvp.Key) Then
                                    removeList.Add(kvp.Key, kvp.Key)
                                End If

                            Catch ex As Exception

                            End Try
                        Else

                        End If

                    Next

                    ' jetzt die Liste bereinigen ...
                    For Each tmpPvName As String In removeList
                        pvNamesList.Remove(tmpPvName)
                    Next

                    If removeList.Count > 0 Then
                        Call updateTreeview(currentBrowserConstellation, pvNamesList,
                                            aKtionskennung, quickList)

                    End If

                ElseIf aKtionskennung = PTTvActions.chgInSession Then

                    Dim removeList As New Collection


                    For Each kvp As KeyValuePair(Of String, clsConstellationItem) In currentBrowserConstellation.Liste
                        'For Each kvp As KeyValuePair(Of String, clsProjekt) In aktuelleGesamtListe.liste
                        Dim hproj As clsProjekt = browserAlleProjekte.getProject(kvp.Key)

                        If Not filter.isEmpty Then
                            If Not IsNothing(hproj) Then
                                ok = filter.doesNotBlock(hproj)
                            Else
                                ok = False
                            End If

                        Else
                            ok = True
                        End If

                        If Not ok Then
                            ' in RemoveListe aufnehmen - diese Projekte werden nachher alle aus aktuelleGesamtliste rausgenommen 
                            Try

                                If Not removeList.Contains(kvp.Key) Then
                                    removeList.Add(kvp.Key, kvp.Key)
                                End If

                            Catch ex As Exception

                            End Try
                        Else

                        End If

                    Next

                    ' jetzt die Liste bereinigen ...
                    For Each tmpPvName As String In removeList
                        currentBrowserConstellation.remove(tmpPvName)
                    Next

                    If currentBrowserConstellation.sortCriteria = ptSortCriteria.customTF Then
                        ' jetzt wird das SortCriteria umgesetzt, weil andernfalls, bei customTF, die 
                        ' Zeilen unverändert bleiben ... 
                        currentBrowserConstellation.sortCriteria = ptSortCriteria.customListe
                    End If

                    ' jetzt müssen die tfZeile neu besetzt werden;
                    '  nach standard, d.h 0 bedeutet einfach sortiert nach Name 
                    ' tk 21.3.17: ab jetzt nicht mehr .... jetzt wird ja in der _sortlist alles mitgeführt 
                    ''currentBrowserConstellation.setTfZeilen(0)

                    If removeList.Count > 0 Then
                        Call updateTreeview(currentBrowserConstellation, pvNamesList,
                                            aKtionskennung, quickList)

                        If aKtionskennung = PTTvActions.chgInSession Then
                            ' erst am Ende alle Diagramme neu machen ...

                            Dim tmpConstellation As New clsConstellations
                            tmpConstellation.Add(currentBrowserConstellation)

                            Call showConstellations(constellationsToShow:=tmpConstellation,
                                                    clearBoard:=True, clearSession:=False, storedAtOrBefore:=storedAtOrBefore)

                            ''If aKtionskennung = PTTvActions.chgInSession Then
                            ''    Call awinNeuZeichnenDiagramme(2)
                            ''End If

                        End If

                    End If


                Else

                    Call MsgBox("nicht unterstützte Option")

                End If


                'Call buildTreeview(projektHistorien, TreeViewProjekte, aktuelleGesamtListe, pvNamesList, _
                '                   aKtionskennung, quickList, Me.filterIsActive, storedAtOrBefore)
                stopRecursion = False

                ' Das DeleteFilterIcon mit Bild versehen 
                Me.deleteFilterIcon.Image = My.Resources.funnel_delete
                Me.deleteFilterIcon.Enabled = True

            End If
        End With

        Me.Cursor = Cursors.Default


    End Sub



    'Private Sub dropBoxTimeStamps_SelectedIndexChanged(sender As Object, e As EventArgs)

    '    'Dim browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

    '    stopRecursion = True

    '    Me.Cursor = Cursors.WaitCursor

    '    Dim storedAtOrBefore As Date
    '    If IsNothing(dropBoxTimeStamps.SelectedItem) Then
    '        storedAtOrBefore = Date.Now
    '    Else
    '        storedAtOrBefore = CDate(dropBoxTimeStamps.SelectedItem)
    '    End If

    '    If aKtionskennung = PTTvActions.loadPV Or _
    '        aKtionskennung = PTTvActions.delFromDB Then

    '        pvNamesList = buildPvNamesList(storedAtOrBefore)
    '        quickList = True
    '    End If

    '    Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, aKtionskennung, quickList)

    '    'Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, aKtionskennung, False)
    '    'Call buildTreeview(projektHistorien, TreeViewProjekte, pvNamesList, currentBrowserConstellation, _
    '    '                   aKtionskennung, quickList, storedAtOrBefore)

    '    stopRecursion = False

    '    Me.Cursor = Cursors.Default

    '    ' Fokus an TreeViewPRojekte geben 
    '    TreeViewProjekte.Focus()

    'End Sub


    Private Sub dropboxScenarioNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dropboxScenarioNames.SelectedIndexChanged

    End Sub

    Private Sub deleteFilterIcon_Click(sender As Object, e As EventArgs) Handles deleteFilterIcon.Click

        Me.Cursor = Cursors.WaitCursor
        Dim storedAtOrBefore As Date
        If IsNothing(requiredDate.Value) Then
            storedAtOrBefore = Date.Now
        Else
            storedAtOrBefore = CDate(requiredDate.Value)
        End If

        If quickList Then

            ' hier wird jetzt die Raw-List geholt, d.h die enthält neben allen anderen Varianten auch die Basis- und Vorgabe-(PFV)Variante 
            pvNamesListRaw = buildPvNamesList(storedAtOrBefore)
            pvNamesList = reduceRawListTo(pvNamesListRaw, awinSettings.loadPFV)

            stopRecursion = True
            Call updateTreeview(currentBrowserConstellation, pvNamesList, aKtionskennung, quickList)
            stopRecursion = False
        Else

            currentBrowserConstellation = beforeFilterConstellation.copy
            'Dim browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

            beforeFilterConstellation = Nothing

            ' jetzt das entzsprechende Szenario wieder laden 
            Dim tmpConstellation As New clsConstellations
            tmpConstellation.Add(currentBrowserConstellation)




            Call showConstellations(constellationsToShow:=tmpConstellation,
                                    clearBoard:=True, clearSession:=False, storedAtOrBefore:=storedAtOrBefore)

            '' neu Zeichnen der Diagramme
            'Call awinNeuZeichnenDiagramme(2)


            ' jetzt muss der Last-Filter zurückgesetzt werden 
            Dim emptyCollection As New Collection
            Dim fName As String = "Last"

            Dim lastFilter As New clsFilter(fName, emptyCollection, emptyCollection, emptyCollection,
                                            emptyCollection, emptyCollection, emptyCollection)
            filterDefinitions.storeFilter(fName, lastFilter)

            stopRecursion = True
            Call updateTreeview(currentBrowserConstellation, pvNamesList, aKtionskennung, False)
            'Call buildTreeview(projektHistorien, TreeViewProjekte, browserAlleProjekte, pvNamesList, _
            '                   aKtionskennung, quickList, Me.filterIsActive, storedAtOrBefore)
            stopRecursion = False

        End If


        ' Das DeleteFilterIcon mit Bild versehen 
        Me.deleteFilterIcon.Image = Nothing
        Me.deleteFilterIcon.Enabled = False

        Me.Cursor = Cursors.Arrow


    End Sub

    'Private Sub dropBoxTimeStamps_MouseHover(sender As Object, e As EventArgs)
    '    ToolTipStand.Show("Angabe des Referenzdatums, zu dem die Projekte geladen werden; Default ist immer der aktuelle Stand", dropBoxTimeStamps, 2000)
    'End Sub

    Private Sub requiredDate_MouseHover(sender As Object, e As EventArgs) Handles requiredDate.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "Angabe des Referenzdatums, zu dem die Projekte geladen werden; Default ist immer der aktuelle Stand"
        Else
            ttText = "Reference-Date: all projects are selected with timestamp at or before that date; Default is current date & time"
        End If
        ToolTipStand.Show(ttText, requiredDate, 2000)
    End Sub

    Private Sub SelectionSet_MouseHover(sender As Object, e As EventArgs) Handles SelectionSet.MouseHover
        Dim ttText As String = ""

        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            If aKtionskennung = PTTvActions.chgInSession Then
                ttText = "alle Projekte auswählen"
            ElseIf aKtionskennung = PTTvActions.delFromDB Then
                ttText = "alle Projekte und Projekt-Varianten auswählen, die den oben ausgewählten Zeitstempel haben"
            ElseIf aKtionskennung = PTTvActions.loadPV Or aKtionskennung = PTTvActions.loadPVS Then
                ttText = "alle Projekte auswählen"
            End If
        Else
            If aKtionskennung = PTTvActions.chgInSession Then
                ttText = "select all projects"
            ElseIf aKtionskennung = PTTvActions.delFromDB Then
                ttText = "select all projects and variants with timestamps at above mentioned date"
            ElseIf aKtionskennung = PTTvActions.loadPV Or aKtionskennung = PTTvActions.loadPVS Then
                ttText = "select all projects"
            End If
        End If


        ToolTipStand.Show(ttText, SelectionSet, 2000)

    End Sub

    Private Sub SelectionReset_MouseHover(sender As Object, e As EventArgs) Handles SelectionReset.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "alle Elemente de-selektieren"
        Else
            ttText = "de-select all elements"
        End If
        ToolTipStand.Show(ttText, SelectionReset, 2000)
    End Sub

    Private Sub collapseCompletely_MouseHover(sender As Object, e As EventArgs) Handles collapseCompletely.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "Struktur ganz zusammenklappen"
        Else
            ttText = "collapse structure completely"
        End If
        ToolTipStand.Show(ttText, SelectionReset, 2000)
    End Sub

    Private Sub expandCompletely_MouseHover(sender As Object, e As EventArgs) Handles expandCompletely.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "Struktur ganz öffnen"
        Else
            ttText = "expand structure completely"
        End If
        ToolTipStand.Show(ttText, expandCompletely, 2000)
    End Sub

    Private Sub filterIcon_MouseHover(sender As Object, e As EventArgs) Handles filterIcon.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "Filter definieren und anwenden"
        Else
            ttText = "Define and apply filter"
        End If
        ToolTipStand.Show(ttText, filterIcon, 2000)
    End Sub

    Private Sub deleteFilterIcon_MouseHover(sender As Object, e As EventArgs) Handles deleteFilterIcon.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "Filter löschen und zurücksetzen"
        Else
            ttText = "Reset Filter"
        End If
        ToolTipStand.Show(ttText, deleteFilterIcon, 2000)
    End Sub


    Private Sub dropboxScenarioNames_MouseHover(sender As Object, e As EventArgs) Handles dropboxScenarioNames.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "Portfolio-Name auswählen oder neuen Namen eingeben"
        Else
            ttText = "Select portfolio name and/or edit new name"
        End If
        ToolTipStand.Show(ttText, deleteFilterIcon, 2000)
    End Sub

    Private Sub onlyActive_MouseHover(sender As Object, e As EventArgs) Handles onlyActive.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "Filter auf angezeigte Projekte und Projekt-Varianten"
        Else
            ttText = "Filter: only selected projects and variants"
        End If
        ToolTipStand.Show(ttText, deleteFilterIcon, 2000)
    End Sub

    Private Sub onlyInactive_MouseHover(sender As Object, e As EventArgs) Handles onlyInactive.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "Filter auf nicht angezeigte Projekte und Projekt-Varianten"
        Else
            ttText = "Filter: only un-selected projects and variants"
        End If
        ToolTipStand.Show(ttText, deleteFilterIcon, 2000)
    End Sub

    Private Sub backToInit_MouseHover(sender As Object, e As EventArgs) Handles backToInit.MouseHover
        Dim ttText As String = ""
        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            ttText = "Filter auf angezeigte / nicht angezeigte Projekte und Projekt-Varianten aufheben"
        Else
            ttText = "Reset Filter selected/un-selected projects"
        End If
        ToolTipStand.Show(ttText, deleteFilterIcon, 2000)
    End Sub



    ''' <summary>
    ''' reduziert die Constellation auf alle Projekt-Varianten mit Attribut Show 
    ''' macht nur Sinn bei chgInSession; wird also nur von dort aus aufgerufen ... 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub onlyActive_Click(sender As Object, e As EventArgs) Handles onlyActive.Click

        If currentConstellationName <> calcLastSessionScenarioName() Then
            currentConstellationName = calcLastSessionScenarioName()

            Dim preText As String = "Portfolio "
            If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                preText = "Portfolio "
            End If
            Me.Text = preText & currentConstellationName
        End If

        Me.Cursor = Cursors.WaitCursor
        Call modifyTreeviewToShowAttribute(showKennung:=ptPPAshowAttributes.show)
        Me.Cursor = Cursors.Default

        backToInit.Visible = True
        onlyInactive.Visible = False

    End Sub

    Private Sub onlyInactive_Click(sender As Object, e As EventArgs) Handles onlyInactive.Click

        If currentConstellationName <> calcLastSessionScenarioName() Then
            currentConstellationName = calcLastSessionScenarioName()

            Dim preText As String = "Portfolio "
            If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                preText = "Portfolio "
            End If
            Me.Text = preText & currentConstellationName
        End If

        Me.Cursor = Cursors.WaitCursor
        Call modifyTreeviewToShowAttribute(showKennung:=ptPPAshowAttributes.noShow)
        Me.Cursor = Cursors.Default

        backToInit.Visible = True
        onlyActive.Visible = False

    End Sub

    Private Sub backToInit_Click(sender As Object, e As EventArgs) Handles backToInit.Click

        Me.Cursor = Cursors.WaitCursor
        Call modifyTreeviewToShowAttribute(showKennung:=ptPPAshowAttributes.all)
        Me.Cursor = Cursors.Default

        browserConstellationSavPM = Nothing
        onlyActive.Visible = True
        onlyInactive.Visible = True
        backToInit.Visible = False

    End Sub

    ''' <summary>
    ''' reduziert den TreeView auf die Projekte mit requiredShowAttribute 
    ''' </summary>
    ''' <param name="showKennung">gibt an, ob alle, nur Show, oder nur noShow gezeigt werden soll</param>
    ''' <param name="appliesToVariantsAsWell"></param>
    ''' <remarks></remarks>
    Private Sub modifyTreeviewToShowAttribute(ByVal showKennung As Integer,
                                                  Optional ByVal appliesToVariantsAsWell As Boolean = True)

        Dim requiredShowAttribute As Boolean = True

        If showKennung = ptPPAshowAttributes.all Then
            ' keine Relevanz für requiredShowAttribute, einfach alle  
        ElseIf showKennung = ptPPAshowAttributes.show Then
            requiredShowAttribute = True
        ElseIf showKennung = ptPPAshowAttributes.noShow Then
            requiredShowAttribute = False
        Else
            Exit Sub
        End If

        Dim storedAtOrBefore As Date = Date.Now

        Dim anzPVsBefore As Integer = currentBrowserConstellation.count

        ' jetzt wird die CurrentConstellation entsprechend neu bestimmt  ... 
        If showKennung = ptPPAshowAttributes.all Then
            If Not IsNothing(browserConstellationSavPM) Then
                currentBrowserConstellation = browserConstellationSavPM.copy
            End If
        Else
            If IsNothing(browserConstellationSavPM) Then
                browserConstellationSavPM = currentBrowserConstellation.copy
            End If

            If appliesToVariantsAsWell Then
                ' dieser Befehl behält nur die Projekt-Varianten mit showAttribute = requiredShowAttribute 
                Call currentBrowserConstellation.reduceToElementsWith(showAttribute:=requiredShowAttribute)
            Else
                ' dieser Befehl behält alle Projekt-Varianten von Projekte mit ShowAttribute = requiredShowAttribute  
                Call currentBrowserConstellation.reduceToProjectsWith(requiredShowAttribute:=requiredShowAttribute)
            End If
        End If



        ' jetzt wird die CurrentBrowserConstellation entsprechend reduziert 
        If currentBrowserConstellation.count <> anzPVsBefore Then

            ' die Positionierung entsprechend Standard setzen ...
            ' tk 21.3.17 jetzt nicht mehr 
            ' currentBrowserConstellation.setTfZeilen(0)

            Dim tmpConstellation As New clsConstellations
            tmpConstellation.Add(currentBrowserConstellation)

            ' auf der Multiprojekt-Tafel entsprechend anzeigen 
            Call showConstellations(constellationsToShow:=tmpConstellation,
                                    clearBoard:=True, clearSession:=False, storedAtOrBefore:=storedAtOrBefore)

            ' den TreeView updaten ... 
            stopRecursion = True
            Call updateTreeview(currentBrowserConstellation, pvNamesList,
                                            aKtionskennung, quickList)
            stopRecursion = False

            ' '' die Diagramme aktualisieren 
            ''If aKtionskennung = PTTvActions.chgInSession Then
            ''    Call awinNeuZeichnenDiagramme(2)
            ''End If

        End If

    End Sub


    Private Sub storeToDBasWell_CheckedChanged(sender As Object, e As EventArgs) Handles storeToDBasWell.CheckedChanged

        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
            If storeToDBasWell.Checked Then
                Me.OKButton.Text = "in Session und DB speichern"
            Else
                Me.OKButton.Text = "in Session speichern"
            End If
        Else
            If storeToDBasWell.Checked Then
                Me.OKButton.Text = "Save to Session and DB"
            Else
                Me.OKButton.Text = "Save to Session"
            End If
        End If

    End Sub

    Private Sub showPFV_CheckedChanged(sender As Object, e As EventArgs)

        If stopRecursion Then
            Exit Sub
        End If

        stopRecursion = True

        pvNamesList = reduceRawListTo(pvNamesListRaw, awinSettings.loadPFV)
        Call updateTreeview(currentBrowserConstellation, pvNamesList, aKtionskennung, quickList)

        stopRecursion = False


        ' Fokus an TreeViewPRojekte geben 
        TreeViewProjekte.Focus()

    End Sub


    Private Sub requiredDate_ValueChanged(sender As Object, e As EventArgs) Handles requiredDate.ValueChanged

        If stopRecursion Then
            Exit Sub
        End If

        stopRecursion = True

        Dim storedAtOrBefore As Date

        If Not IsNothing(requiredDate) Then

            If requiredDate.Value >= earliestDate And requiredDate.Value <= Date.Now Then

                storedAtOrBefore = requiredDate.Value.Date.AddHours(23).AddMinutes(59)

            ElseIf requiredDate.Value > Date.Now Then

                requiredDate.Value = Date.Now
                storedAtOrBefore = requiredDate.Value.Date.AddHours(23).AddMinutes(59)

            Else

                Dim msgText As String = "es gibt vor dem " & earliestDate.ToShortDateString & " keine Projekte in der Datenbank "
                If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                    msgText = "there are no projects at or before " & earliestDate.ToShortDateString & " in the database"
                End If

                Call MsgBox(msgText)

                requiredDate.Value = earliestDate.Date.AddHours(23).AddMinutes(59)
                storedAtOrBefore = earliestDate.Date.AddHours(23).AddMinutes(59)
            End If

        Else
            requiredDate.Value = Date.Now.Date.AddHours(23).AddMinutes(59)
            storedAtOrBefore = Date.Now.Date.AddHours(23).AddMinutes(59)
        End If

        If aKtionskennung = PTTvActions.loadPV Or
            aKtionskennung = PTTvActions.delFromDB Then

            pvNamesListRaw = buildPvNamesList(storedAtOrBefore)
            pvNamesList = reduceRawListTo(pvNamesListRaw, awinSettings.loadPFV)
            quickList = True

        End If

        Call updateTreeview(currentBrowserConstellation, pvNamesList, aKtionskennung, quickList)

        stopRecursion = False


        ' Fokus an TreeViewPRojekte geben 
        TreeViewProjekte.Focus()
    End Sub



    Private Sub OKButton_MouseHover(sender As Object, e As EventArgs) Handles OKButton.MouseHover
        Me.Cursor = Cursors.Default
    End Sub

    ''' <summary>
    ''' aktualisiert bzw. baut die TreeView gemäß der aktuelleGesamtListe bzw. der pvNamesList neu auf
    ''' Rahmenbedingung: stopRecursion ist immer False, wenn Update TreeView aufgerufen wird 
    ''' </summary>
    ''' <param name="constellation"></param>
    ''' <param name="pvNamesList"></param>
    ''' <param name="aKtionskennung"></param>
    ''' <param name="quickList"></param>
    ''' <remarks></remarks>
    Private Sub updateTreeview(ByVal constellation As clsConstellation,
                                  ByVal pvNamesList As SortedList(Of String, String),
                                  ByVal aKtionskennung As Integer,
                                  ByVal quickList As Boolean)

        Dim err As New clsErrorCodeMsg

        Dim projectNode As TreeNode
        Dim zeitraumVon As Date = StartofCalendar
        Dim zeitraumbis As Date = StartofCalendar.AddYears(20)
        'Dim storedHeute As Date = Now
        Dim storedGestern As Date = StartofCalendar
        Dim pname As String = ""
        Dim variantName As String = ""
        Dim loadErrorMsg As String = ""

        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            ' es ist ein Zeitraum definiert 
            zeitraumVon = getDateofColumn(showRangeLeft, False)
            zeitraumbis = getDateofColumn(showRangeRight, True)
        End If

        ' steuert, ob erstmal nur Projekt-Namen, Varianten-Namen gelesen werden 
        ' geht wesentlich schneller, wenn es sich um eine Datenbank mit sehr vielen Projekten handelt ... 


        With TreeViewProjekte
            .Nodes.Clear()
        End With

        If Not IsNothing(constellation) Or pvNamesList.Count >= 1 Then

            If Not noDB And aKtionskennung = PTTvActions.setWriteProtection Then
                'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)
            End If

            With TreeViewProjekte

                .CheckBoxes = True

                Dim projektliste As SortedList(Of String, String)

                If quickList Then
                    projektliste = New SortedList(Of String, String)
                    For Each kvp As KeyValuePair(Of String, String) In pvNamesList
                        Dim tmpName As String = kvp.Key
                        If tmpName.Contains("#") Then
                            Dim tmpStr() As String = tmpName.Split(New Char() {CChar("#")})
                            If Not projektliste.ContainsKey(tmpStr(0)) Then
                                projektliste.Add(tmpStr(0), tmpStr(0))
                            End If
                        Else
                            If Not projektliste.ContainsKey(tmpName) Then
                                projektliste.Add(tmpName, tmpName)
                            End If
                        End If
                    Next

                Else
                    ' hole die Namen aus der sortierten Liste, nicht aus der cItem-Liste 
                    projektliste = constellation.getProjectNames(False)
                End If

                Dim showPname As Boolean



                For Each kvp As KeyValuePair(Of String, String) In projektliste

                    showPname = True
                    pname = kvp.Value ' der key ist das sortier-Kriterium, kann pName sein, aber auch was ganz anderes 

                    Dim hproj As clsProjekt = Nothing
                    Dim variantNames As Collection

                    If quickList Then
                        variantNames = getVariantListeFromPVNames(pvNamesList, pname)
                    Else
                        variantNames = constellation.getVariantNames(pname, True)
                    End If


                    If ShowProjekte.contains(pname) Then
                        hproj = ShowProjekte.getProject(pname)
                        'shownVariant = "(" & hproj.variantName & ")"
                        'projectIsShown = True

                    ElseIf AlleProjekte.Count > 0 Then
                        Dim tmpList As Collection = AlleProjekte.getVariantNames(pname, False)

                        If tmpList.Count > 0 Then
                            variantName = CStr(tmpList.Item(1))
                            hproj = AlleProjekte.getProject(pname, variantName)
                        End If

                    End If



                    ' im Falle activate Variante / Portfolio definieren: nur die Projekte anzeigen, die auch tatsächlich mehrere Varianten haben 
                    If aKtionskennung = PTTvActions.activateV Or aKtionskennung = PTTvActions.deleteV Then
                        If constellation.getVariantZahl(pname) <= 1 Then
                            showPname = False
                        End If
                    End If

                    If showPname Then


                        projectNode = .Nodes.Add(pname)
                        'projectNode.Text = pname
                        ' das wird jetzt über bestimmeCheckStatus gemacht bzw. über bestimmeNodeAppearance
                        ''If aKtionskennung = PTTvActions.chgInSession Or _
                        ''    aKtionskennung = PTTvActions.activateV Then

                        ''    If projectIsShown Then
                        ''        projectNode.Checked = True
                        ''        If aKtionskennung = PTTvActions.chgInSession Then
                        ''            projectNode.Text = pname & " (" & hproj.variantName & ")"
                        ''        End If
                        ''    End If
                        ''ElseIf aKtionskennung = PTTvActions.setWriteProtection Then
                        ''    ' setzen der Checked Informationen 
                        ''End If

                        ' damit kann evtl direkt auf den Node zugegriffen werden ...
                        projectNode.Name = pname



                        If Not IsNothing(hproj) Then
                            variantName = hproj.variantName
                        End If

                        ' Platzhalter einfügen; wird für alle Aktionskennungen benötigt

                        If variantNames.Count > 1 Or
                            aKtionskennung = PTTvActions.delFromDB Then

                            Dim vName As String = variantName
                            projectNode.Tag = "X"
                            For iv As Integer = 1 To variantNames.Count
                                vName = CStr(variantNames.Item(iv))
                                Dim vNameStripped As String = ""
                                Dim tmpStr() As String = vName.Split(New Char() {CChar("("), CChar(")")})
                                If tmpStr.Length = 1 Then
                                    vNameStripped = tmpStr(0)
                                ElseIf tmpStr.Length >= 3 Then
                                    vNameStripped = tmpStr(1).Trim
                                End If

                                Dim variantNode As TreeNode = projectNode.Nodes.Add(vName)
                                'variantNode.Text = vName

                                If aKtionskennung = PTTvActions.delFromDB Then
                                    variantNode.Tag = "P"
                                    Dim tmpNodeLevel2 As TreeNode = variantNode.Nodes.Add("Platzhalter-Datum")
                                Else
                                    variantNode.Tag = "X"
                                End If

                                Call bestimmeNodeCheckStatus(variantNode, aKtionskennung, PTTreeNodeTyp.pVariant,
                                                             pname, vNameStripped)
                                Call bestimmeNodeAppearance(variantNode, aKtionskennung, PTTreeNodeTyp.pVariant, pname, vNameStripped)

                            Next

                        Else
                            projectNode.Tag = "X"
                        End If

                        Call bestimmeNodeCheckStatus(projectNode, aKtionskennung, PTTreeNodeTyp.project,
                                                      pname, variantName)
                        Call bestimmeNodeAppearance(projectNode, aKtionskennung, PTTreeNodeTyp.project, pname, variantName)

                    End If

                Next

            End With
        Else
            Call MsgBox(loadErrorMsg)
        End If


    End Sub

    ''' <summary>
    ''' bestimmt in Abhängigkeit von Aktionskennung den Checkstatus, den das Projekt bzw. die Projekt-Variante haben soll 
    ''' </summary>
    ''' <param name="currentNode"></param>
    ''' <param name="aktionskennung"></param>
    ''' <param name="nodeTyp"></param>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <remarks></remarks>
    Private Sub bestimmeNodeCheckStatus(ByRef currentNode As TreeNode,
                                           ByVal aktionskennung As Integer, ByVal nodeTyp As Integer,
                                           ByVal pName As String, ByVal vName As String)

        Dim hproj As clsProjekt
        Dim shownVariant As String

        If aktionskennung = PTTvActions.chgInSession Or
                            aktionskennung = PTTvActions.activateV Then
            Dim projectIsShown As Boolean = False

            If ShowProjekte.contains(pName) Then
                hproj = ShowProjekte.getProject(pName)
                shownVariant = hproj.variantName

                If nodeTyp = PTTreeNodeTyp.project Then
                    ' setze den Projekt-Node
                    currentNode.Checked = True

                ElseIf nodeTyp = PTTreeNodeTyp.pVariant Then
                    If shownVariant = vName Then
                        currentNode.Checked = True
                    Else
                        currentNode.Checked = False
                    End If


                End If
                projectIsShown = True
            Else
                If nodeTyp = PTTreeNodeTyp.project Then
                    currentNode.Checked = False
                Else
                    ' keine Veränderung am CheckStatus vornehmen 
                End If
            End If


        ElseIf aktionskennung = PTTvActions.setWriteProtection Then
            ' setzen der Checked Informationen 
            If nodeTyp = PTTreeNodeTyp.project Then
                If currentNode.Nodes.Count = 0 Then
                    ' es geht um den WriteProtections-Status des einen pName, vName Projektes 
                    Dim variantNames As Collection = AlleProjekte.getVariantNames(pName, False)
                    Dim activeVariantName As String = vName
                    If variantNames.Count = 1 Then
                        activeVariantName = CStr(variantNames.Item(1))
                    End If
                    Dim pvName As String = calcProjektKey(pName, activeVariantName)
                    If writeProtections.isProtected(pvName) Then
                        currentNode.Checked = True
                    Else
                        currentNode.Checked = False
                    End If
                Else
                    ' der Check-Status ergibt sich aus der Betrachtung der Child-Nodes 
                    ' child-Nodes unterschiedlich: project-Node nicht gecheckt 
                    ' child Nodes alle gecheckt: project Node gecheckt 
                    Dim atleastOneIsDifferent As Boolean = False
                    Dim checkStatus As Boolean = False
                    For i As Integer = 1 To currentNode.Nodes.Count
                        Dim childNode As TreeNode = currentNode.Nodes.Item(i - 1)
                        If i = 1 Then
                            checkStatus = childNode.Checked
                        ElseIf childNode.Checked <> checkStatus Then
                            atleastOneIsDifferent = True
                            Exit For
                        End If
                    Next
                    If atleastOneIsDifferent Then
                        currentNode.Checked = False
                    Else
                        currentNode.Checked = checkStatus
                    End If
                End If
            ElseIf nodeTyp = PTTreeNodeTyp.pVariant Then
                ' es geht um den WriteProtections-Status des einen pName, vName Projektes
                Dim pvName As String = calcProjektKey(pName, vName)
                If writeProtections.isProtected(pvName) Then
                    currentNode.Checked = True
                Else
                    currentNode.Checked = False
                End If
            End If

        End If


    End Sub

    ''' <summary>
    ''' bestimmt das Erscheinungsbild des Knoten in Abhängigkeiten von aktionskennung und dem Check-Zustand des Knoten
    ''' ausserdem wird berücksichtigt, ob der Knoten isoliert betrachtet werden soll oder 
    ''' sein Erscheinunngsbild in Abhängigkeit von den Child Knoten gesetzt werden soll 
    ''' </summary>
    ''' <param name="currentNode">der übergebene Node</param>
    ''' <param name="aktionskennung">mit welcher Aktionskennung wurde der Portfolio Browser aufgerufe</param>
    ''' <param name="nodeTyp">project, variant, timestamp</param>
    ''' <param name="pName">der Projekt Name</param>
    ''' <param name="vName">der Varianten Name</param>
    ''' <remarks></remarks>
    Private Sub bestimmeNodeAppearance(ByRef currentNode As TreeNode,
                                              ByVal aktionskennung As Integer, ByVal nodeTyp As Integer,
                                              ByVal pName As String, ByVal vName As String)


        'Dim fontProtectedbyOther As System.Drawing.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Italic)

        'Dim fontPermanentProtected As System.Drawing.Font = awinSettings.protectedPermanentFont
        'Dim fontNormal As System.Drawing.Font = awinSettings.normalFont

        'Dim colorProtectedByMe As System.Drawing.Color = awinSettings.protectedByMeColor
        'Dim colorProtectedByOther As System.Drawing.Color = awinSettings.protectedByOtherColor
        'Dim colorNormal As System.Drawing.Color = awinSettings.normalColor
        'Dim colorNoShow As System.Drawing.Color = awinSettings.noShowColor

        ' hier auf Normal Font setzen ; in den TreeView Eigenschaften ist die Schriftgröße auf 12 gesetzt, 
        ' um sicherzustellen, dass der Text immer vollständig angezeigt wird 
        ' ur rausgenommen, weil es zu Fehlern führt  
        ' currentNode.NodeFont = awinSettings.normalFont


        If nodeTyp = PTTreeNodeTyp.project Then

            If aktionskennung = PTTvActions.chgInSession Then
                If vName = "" Then
                    currentNode.Text = pName

                Else
                    currentNode.Text = pName & " (" & vName & ")"
                End If
            End If

            If allDependencies.projectCount > 0 Then
                ' es gibt irgendwelche Dependencies, die Lead-Projekte, abhängigen Projekte 
                ' und sowohl-als-auch-Projekte werden farblich markiert  

                ' die Projekte suchen, von denen dieses Projekt abhängt 
                Dim passivListe As Collection = allDependencies.passiveListe(pName, PTdpndncyType.inhalt)
                Dim aktivListe As Collection = allDependencies.activeListe(pName, PTdpndncyType.inhalt)

                If passivListe.Count > 0 And aktivListe.Count = 0 Then
                    ' ist nur abhängiges Projekt ...
                    'currentNode.ForeColor = Color.Gray
                    If Not currentNode.Text.EndsWith(" /D") Then
                        currentNode.Text = currentNode.Text & " /D"
                    End If


                ElseIf passivListe.Count = 0 And aktivListe.Count > 0 Then
                    ' hat abhängige Projekte  
                    'currentNode.ForeColor = Color.OrangeRed
                    If Not currentNode.Text.EndsWith(" /L") Then
                        currentNode.Text = currentNode.Text & " /L"
                    End If


                ElseIf passivListe.Count > 0 And aktivListe.Count > 0 Then
                    ' ist abhängig und hat abhängige Projekte 
                    'currentNode.ForeColor = Color.Orange
                    If Not currentNode.Text.EndsWith(" /LD") Then
                        currentNode.Text = currentNode.Text & " /LD"
                    End If

                End If

            End If

            If aktionskennung = PTTvActions.setWriteProtection And Not noDB Then

                If (currentNode.Nodes.Count = 0) Then
                    Dim pvName As String = calcProjektKey(pName, vName)
                    If currentNode.Checked Then

                        If dbUsername = writeProtections.lastModifiedBy(pvName) Then
                            ' entsprechend kennzeichnen 
                            currentNode.ForeColor = awinSettings.protectedByMeColor
                        Else
                            ' entsprechend kennzeichnen 
                            currentNode.ForeColor = awinSettings.protectedByOtherColor
                        End If

                        If writeProtections.isPermanentProtected(pvName) Then
                            currentNode.NodeFont = awinSettings.protectedPermanentFont
                        Else
                            currentNode.NodeFont = awinSettings.normalFont
                        End If
                    Else
                        ' entsprechend kennzeichnen 
                        currentNode.ForeColor = awinSettings.normalColor
                        currentNode.NodeFont = awinSettings.normalFont
                    End If
                Else

                    currentNode.ForeColor = awinSettings.normalColor
                    currentNode.NodeFont = awinSettings.normalFont

                    ' wenn alle Varianten drunter geschützt / nicht geschützt sind: entsprechend setzen 
                    Call adjustNodeAppearanceToChilds(currentNode)

                End If


            ElseIf aktionskennung = PTTvActions.delFromDB And Not noDB Then
                If (currentNode.Nodes.Count = 0) Then

                    If notReferencedByAnyPortfolio(pName, vName) And
                        Not writeProtections.isProtected(calcProjektKey(pName, vName)) Then
                        ' kann gelöscht werden  
                        currentNode.ForeColor = awinSettings.normalColor
                    Else
                        currentNode.ForeColor = awinSettings.noShowColor
                    End If
                Else
                    currentNode.ForeColor = awinSettings.normalColor

                    ' wenn alle Varianten drunter geschützt / nicht geschützt sind: entsprechend setzen 
                    Call adjustNodeAppearanceToChilds(currentNode)
                End If

            ElseIf aktionskennung = PTTvActions.delFromSession Then
                ' alle  deutlicher zeigen, die im NoShow sind 
                If currentNode.Nodes.Count = 0 Then
                    If ShowProjekte.contains(pName) Then
                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                        If hproj.variantName = vName Then
                            currentNode.ForeColor = awinSettings.noShowColor
                        Else
                            currentNode.ForeColor = awinSettings.normalColor
                        End If
                    Else
                        currentNode.ForeColor = awinSettings.normalColor
                    End If
                Else
                    ' hat Kinder ...
                    currentNode.ForeColor = awinSettings.normalColor

                    ' wenn alle Varianten drunter geschützt / nicht geschützt sind: entsprechend setzen 
                    Call adjustNodeAppearanceToChilds(currentNode)
                End If


            ElseIf aktionskennung = PTTvActions.loadPV Then
                ' alle  markieren, die noch nicht geladen sind, ob im Show oder NoShow  
                If currentNode.Nodes.Count = 0 Then
                    Dim tmpKey As String = calcProjektKey(pName, vName)
                    If AlleProjekte.Containskey(tmpKey) Then
                        Dim hproj As clsProjekt = AlleProjekte.getProject(tmpKey)
                        If Not IsNothing(hproj) Then
                            currentNode.ForeColor = awinSettings.noShowColor
                        Else
                            currentNode.ForeColor = awinSettings.normalColor
                        End If
                    Else
                        currentNode.ForeColor = awinSettings.normalColor
                    End If
                Else
                    ' hat Kinder ...
                    currentNode.ForeColor = awinSettings.normalColor

                    ' wenn alle Varianten drunter geschützt / nicht geschützt sind: entsprechend setzen 
                    Call adjustNodeAppearanceToChilds(currentNode)
                End If



            Else
                currentNode.ForeColor = awinSettings.normalColor
            End If

        ElseIf nodeTyp = PTTreeNodeTyp.pVariant Then

            ' der Current Node Text für die Variante ist schon gesetzt ...
            ' in updateTreeView

            If aktionskennung = PTTvActions.setWriteProtection And Not noDB Then
                Dim pvName As String = calcProjektKey(pName, vName)

                If currentNode.Checked Then
                    If dbUsername = writeProtections.lastModifiedBy(pvName) Then
                        ' entsprechend kennzeichnen 
                        currentNode.ForeColor = awinSettings.protectedByMeColor
                    Else
                        ' entsprechend kennzeichnen 
                        currentNode.ForeColor = awinSettings.protectedByOtherColor
                    End If

                    If writeProtections.isPermanentProtected(pvName) Then
                        currentNode.NodeFont = awinSettings.protectedPermanentFont
                    Else
                        currentNode.NodeFont = awinSettings.normalFont
                    End If
                Else
                    ' entsprechend kennzeichnen 
                    currentNode.ForeColor = awinSettings.normalColor
                    currentNode.NodeFont = awinSettings.normalFont
                End If



            ElseIf aktionskennung = PTTvActions.delFromDB And Not noDB Then

                If notReferencedByAnyPortfolio(pName, vName) Then
                    ' kann gelöscht werden  
                    currentNode.ForeColor = awinSettings.normalColor
                Else
                    currentNode.ForeColor = awinSettings.noShowColor
                End If


            ElseIf aktionskennung = PTTvActions.delFromSession Then

                If ShowProjekte.contains(pName) Then
                    Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                    If hproj.variantName = vName Then
                        currentNode.ForeColor = awinSettings.noShowColor
                    Else
                        currentNode.ForeColor = awinSettings.normalColor
                    End If
                Else
                    currentNode.ForeColor = awinSettings.normalColor
                End If


            ElseIf aktionskennung = PTTvActions.loadPV Then

                ' alle  markieren, die noch nicht geladen sind, ob im Show oder NoShow  
                Dim tmpKey As String = calcProjektKey(pName, vName)
                If AlleProjekte.Containskey(tmpKey) Then
                    Dim hproj As clsProjekt = AlleProjekte.getProject(tmpKey)
                    If Not IsNothing(hproj) Then
                        currentNode.ForeColor = awinSettings.noShowColor
                    Else
                        currentNode.ForeColor = awinSettings.normalColor
                    End If
                Else
                    currentNode.ForeColor = awinSettings.normalColor
                End If

            Else
                currentNode.ForeColor = Drawing.Color.Black
            End If

        End If
        ' Berücksichtigung der Abhängigkeiten im TreeView ...


    End Sub


    ''' <summary>
    ''' passt die ForeColor des Eltern-Knoten an die Child-Nodes an, sofern die alle gleich sind 
    ''' </summary>
    ''' <param name="currentNode"></param>
    ''' <remarks></remarks>
    Private Sub adjustNodeAppearanceToChilds(ByRef currentNode As TreeNode)

        Dim atLeastOneDifferenceInColor As Boolean = False
        Dim atLeastOneDifferenceInFont As Boolean = False

        'Dim colorNormal As System.Drawing.Color = Drawing.Color.Black
        'Dim fontNormal As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, System.Drawing.FontStyle.Regular)

        Dim colorNormal As System.Drawing.Color = awinSettings.normalColor
        Dim fontNormal As System.Drawing.Font = awinSettings.normalFont

        Dim currentColor As System.Drawing.Color = currentNode.ForeColor
        Dim currentFont As System.Drawing.Font = currentNode.NodeFont

        For iv As Integer = 1 To currentNode.Nodes.Count

            Dim variantNode As TreeNode = currentNode.Nodes.Item(iv - 1)

            If iv = 1 Then
                currentColor = variantNode.ForeColor
                currentFont = variantNode.NodeFont
            Else
                If currentColor.ToArgb = variantNode.ForeColor.ToArgb Then
                    ' identisch ...
                Else
                    atLeastOneDifferenceInColor = True
                End If

                If Not IsNothing(currentFont) And Not IsNothing(variantNode.NodeFont) Then
                    If currentFont.Equals(variantNode.NodeFont) Then
                        ' identisch 
                    Else
                        atLeastOneDifferenceInFont = True
                    End If
                ElseIf IsNothing(currentFont) And IsNothing(currentFont) Then
                    ' identisch 
                Else
                    atLeastOneDifferenceInFont = True
                End If

            End If
        Next

        If Not atLeastOneDifferenceInColor Then
            currentNode.ForeColor = currentColor
        Else
            currentNode.ForeColor = colorNormal
        End If

        If Not atLeastOneDifferenceInFont Then
            currentNode.NodeFont = currentFont
        Else
            currentNode.NodeFont = fontNormal
        End If

    End Sub

    Private Sub OKButton_DockChanged(sender As Object, e As EventArgs) Handles OKButton.DockChanged

    End Sub
End Class