Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports MongoDbAccess
Imports System.Windows.Forms

''' <summary>
''' wird verwendet um Portfolios zu definieren, Varianten zu aktivieren, Projekte und Varianten zu laden, zu aktivieren und zu löschen 
''' </summary>
''' <remarks></remarks>
Public Class frmProjPortfolioAdmin


    Private currentBrowserConstellation As New clsConstellation
    ' wenn Filter erstmalig aufgebaut wird , dann wird browserConstellationSav gemerkt ... 
    Private browserConstellationSav As clsConstellation = Nothing
    ' PlusMinus Saving 
    Private browserConstellationSavPM As clsConstellation = Nothing
    ' wenn aus der Datenbank schnell gelesen werden soll ..
    Private pvNamesList As New SortedList(Of String, String)
    Private quickList As Boolean

    Private earliestDate As Date
    Private projektHistorien As New clsProjektDBInfos
    Private stopRecursion As Boolean = False
    Private constellationName As String = ""

    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection
    Private selectedCosts As New Collection
    Private selectedRoles As New Collection
    Private selectedBUs As New Collection
    Private selectedTyps As New Collection

    Private toolTippsAreShowing As Integer

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
    Friend aKtionskennung As Integer

    Private Sub frmProjPortfolioAdmin_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

    End Sub

    Private Sub frmDefineEditPortfolio_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed


        frmCoord(PTfrm.eingabeProj, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.eingabeProj, PTpinfo.left) = Me.Left
        projektHistorien.clear()

        ' jetzt das aktuelle Szenario als Last speichern ... 
        If aKtionskennung = PTTvActions.chgInSession Or _
            aKtionskennung = PTTvActions.delFromSession Or _
            aKtionskennung = PTTvActions.activateV Or _
            aKtionskennung = PTTvActions.loadPV Then

            currentBrowserConstellation.constellationName = "Last"
            projectConstellations.update(currentBrowserConstellation)

        End If
        

        ' Maus auf Normalmodus zurücksetzen
        appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault

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

                .lblVersionen1.Visible = False
                .lblVersionen2.Visible = False
                .versionsToKeep.Visible = False

                onlyActive.Visible = False
                onlyInactive.Visible = False
                backToInit.Visible = False

                storeToDBasWell.Visible = False

                chkbxPermanent.Visible = False


            ElseIf aKtionskennung = PTTvActions.chgInSession Then

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    .Text = "Szenario "
                Else
                    .Text = "Scenario "
                End If

                .requiredDate.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = True
                .deleteFilterIcon.Visible = True

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
                        .OKButton.Text = "Store to Session and DB"
                    Else
                        .OKButton.Text = "Store to Session"
                    End If
                End If


                Dim testName As String = .OKButton.Name

                .lblVersionen1.Visible = False
                .lblVersionen2.Visible = False
                .versionsToKeep.Visible = False

                onlyActive.Visible = True
                onlyInactive.Visible = True
                backToInit.Visible = False

                storeToDBasWell.Visible = True

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
                    .Text = "aus Session Löschen"
                Else
                    .OKButton.Text = "Delete from Session"
                End If

                .OKButton.Visible = True


                .lblVersionen1.Visible = False
                .lblVersionen2.Visible = False
                .versionsToKeep.Visible = False

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


                .requiredDate.Visible = True
                .lblStandvom.Visible = True

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


                .lblVersionen1.Visible = False
                .lblVersionen2.Visible = False
                .versionsToKeep.Visible = False

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


                .lblVersionen1.Visible = True
                .lblVersionen2.Visible = True

                If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                    lblVersionen1.Text = "delete all except"
                    lblVersionen2.Text = "different versions"
                End If

                .versionsToKeep.Visible = True
                .versionsToKeep.Value = 3
                .lblVersionen1.Top = .lblVersionen1.Top + versionenOffset
                .lblVersionen2.Top = .lblVersionen2.Top + versionenOffset
                .versionsToKeep.Top = .versionsToKeep.Top + versionenOffset
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

                .filterIcon.Visible = True
                .deleteFilterIcon.Visible = True

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
                chkbxPermanent.Visible = True


            End If

        End With


    End Sub


    Private Sub frmDefineEditPortfolio_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Dim browserAlleProjekte As New clsProjekteAlle

        If frmCoord(PTfrm.eingabeProj, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.top))
        End If

        If frmCoord(PTfrm.eingabeProj, PTpinfo.left) > 0 Then
            Me.Left = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.left))
        End If



        ' was sollen die ToolTipps zeigen ? 
        If aKtionskennung = PTTvActions.setWriteProtection Then
            toolTippsAreShowing = ptPPAtooltipps.protectedBy
        Else
            toolTippsAreShowing = ptPPAtooltipps.description
        End If


        ' bestimmen, ob es sich um quicklist handelt ...
        If aKtionskennung = PTTvActions.loadPV Or _
            aKtionskennung = PTTvActions.delFromDB Or _
            aKtionskennung = PTTvActions.delAllExceptFromDB Then
            quickList = True
        Else
            quickList = False
        End If


        ' je nachdem, wie die Aktionskennung ist: setzen der Button Visibilitäten 
        Call defineButtonVisibility()

        ' wie heisst das aktuelle Szenario ? 
        Me.Text = Me.Text & ": " & currentConstellationName

        ' jetzt muss bestimmt werden , was die aktuelle SessionConstellation ist 
        If projectConstellations.Contains(currentConstellationName) And AlleProjekte.Count > 0 Then
            currentBrowserConstellation = projectConstellations.getConstellation(currentConstellationName).copy("Last")
            'browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

        ElseIf projectConstellations.Contains("Last") And AlleProjekte.Count > 0 Then
            currentBrowserConstellation = projectConstellations.getConstellation("Last")
            'browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

        ElseIf AlleProjekte.Count > 0 Then
            'browserAlleProjekte = AlleProjekte.createCopy
            'currentBrowserConstellation = New clsConstellation(browserAlleProjekte, Nothing, "Last", ptSzenarioConsider.all)
            currentBrowserConstellation = New clsConstellation(AlleProjekte, Nothing, "Last", ptSzenarioConsider.all)
        End If

        ' jetzt die Korrektheitsprüfung ...
        If awinSettings.visboDebug Then
            currentBrowserConstellation.checkAndCorrectYourself()
        End If


        ' jetzt die vorkommenden Timestamps auslesen 
        ' aber nicht bei allen Aktionskennungen 

        If aKtionskennung = PTTvActions.chgInSession Or _
            aKtionskennung = PTTvActions.delFromSession Or _
            aKtionskennung = PTTvActions.deleteV Or _
            aKtionskennung = PTTvActions.activateV Then

        Else

            Try
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                Dim tCollection As Collection = request.retrieveZeitstempelFromDB()

                If tCollection.Count >= 1 Then
                    earliestDate = tCollection.Item(tCollection.Count).date.addhours(23).addminutes(59)
                Else
                    earliestDate = Date.Now.Date.AddHours(23).AddMinutes(59)
                End If


                'dropBoxTimeStamps.Items.Clear()

                'For k As Integer = 1 To tCollection.Count
                '    Dim tmpDate As Date = CDate(tCollection.Item(k))
                '    dropBoxTimeStamps.Items.Add(tmpDate)
                'Next

            Catch ex As Exception

            End Try

            ' jetzt ist dropBoxTimeStamps.selecteditem = Nothing ..
        End If


        ' Maus auf Wartemodus setzen
        appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

        If aKtionskennung = PTTvActions.chgInSession Then

            For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste
                If kvp.Key <> "Start" Then
                    dropboxScenarioNames.Items.Add(kvp.Key)
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
        If aKtionskennung = PTTvActions.loadPV Or _
            aKtionskennung = PTTvActions.delFromDB Or _
            aKtionskennung = PTTvActions.delAllExceptFromDB Then

            pvNamesList = buildPvNamesList(storedAtOrBefore)
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

                pvNamesList = buildPvNamesList(storedAtOrBefore)
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
        Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, aKtionskennung, quickList)
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
        appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault

        ' Fokus auf was unverdächtiges setzen 
        dropboxScenarioNames.Focus()

    End Sub


    Private Sub TreeViewProjekte_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles TreeViewProjekte.AfterCheck

        Dim node As TreeNode
        Dim schluessel As String = ""
        'Dim selCollection As SortedList(Of Date, String)
        'Dim timeStamp As Date
        Dim treeLevel As IntegerquickList
        Dim i As Integer, j As Integer
        Dim childNode As TreeNode
        Dim parentNode As TreeNode

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

        ' das Szenario wird im Falle activateV und chgInSession verändert ... 
        ' das muss hier vermerkt werden ...
        If aKtionskennung = PTTvActions.chgInSession Or _
            aKtionskennung = PTTvActions.activateV Then
            If Not currentConstellationName.EndsWith("(*)") Then
                currentConstellationName = currentConstellationName & " (*)"
            End If

            Dim preText As String
            If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                preText = "Szenario "
            Else
                preText = "Scenario "
            End If

            Me.Text = preText & currentConstellationName
        End If



        ' hier wird jetzt sichergestellt, daß nur die nach der aktuellen Aktion gültigen Checks gesetzt werden können
        ' vor allem muss überall dort, wo das Szenario mit diesem Check verändert wird, das currentBrowserSzenario geupdated werden ...
        ' mit Click in TreeView wird verändert: Activate Variant, ChgInSession 

        If aKtionskennung = PTTvActions.delFromDB Or _
            aKtionskennung = PTTvActions.delAllExceptFromDB Or _
            aKtionskennung = PTTvActions.loadPV Then

            stopRecursion = True

            Select Case treeLevel

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

            stopRecursion = False


        ElseIf aKtionskennung = PTTvActions.delFromSession Or _
              aKtionskennung = PTTvActions.deleteV Then

            stopRecursion = True

            Select Case treeLevel

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

            stopRecursion = False

        ElseIf aKtionskennung = PTTvActions.setWriteProtection Then

            stopRecursion = True

            If Not noDB Then

                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                writeProtections.liste = request.retrieveWriteProtectionsFromDB(AlleProjekte)

                Select Case treeLevel

                    Case 0 ' Projekt ist selektiert / nicht selektiert 

                        If node.Nodes.Count = 0 Then
                            ' es gibt nur die eine Projekt-Variante 
                            Dim pName As String = getProjectNameOfTreeNode(node.Text)
                            Dim vName As String = ""
                            Dim variantNames As Collection = AlleProjekte.getVariantNames(pName, False)
                            vName = CStr(variantNames.Item(1))
                            Dim pvName As String = calcProjektKey(pName, vName)

                            If Not (writeProtections.isProtected(pvName)) Or _
                                (writeProtections.isProtected(pvName) And dbUsername = writeProtections.wasProtectedBy(pvName)) Then

                                ' jetzt in der Datenbank setzen 
                                Dim wpItem As New clsWriteProtectionItem

                                With wpItem
                                    .isProtected = node.Checked
                                    .permanent = Me.chkbxPermanent.Checked
                                    .userName = dbUsername
                                    .pvName = pvName
                                End With

                                If request.setWriteProtection(wpItem) Then
                                    ' erfolgreich 
                                    ' nichts tun 
                                Else
                                    ' nicht erfolgreich
                                    node.Checked = Not node.Checked
                                End If
                                ' Liste aktualisieren ...
                                writeProtections.liste = request.retrieveWriteProtectionsFromDB(AlleProjekte)
                                Call bestimmeNodeAppearance(node, aKtionskennung, PTTreeNodeTyp.project, pName, vName, True)


                            Else ' protected und falscher User

                                ' nicht erfolgreich
                                node.Checked = Not node.Checked
                            End If

                        Else
                            ' es gibt mehrere Projekt-Varianten 
                            Dim atleastOneError As Boolean = False
                            For i = 1 To node.Nodes.Count
                                childNode = node.Nodes.Item(i - 1)
                                ' darf es ge- bzw. entcheckt werden ? 
                                Dim pName As String = getProjectNameOfTreeNode(node.Text)
                                Dim vName As String = getVariantNameOfTreeNode(childNode.Text)
                                Dim pvName As String = calcProjektKey(pName, vName)

                                If Not (writeProtections.isProtected(pvName)) Or _
                                    (writeProtections.isProtected(pvName) And dbUsername = writeProtections.wasProtectedBy(pvName)) Then

                                    ' jetzt in der Datenbank setzen 
                                    Dim wpItem As New clsWriteProtectionItem

                                    With wpItem
                                        .isProtected = node.Checked
                                        .permanent = Me.chkbxPermanent.Checked
                                        .userName = dbUsername
                                        .pvName = pvName
                                    End With

                                    If request.setWriteProtection(wpItem) Then
                                        ' erfolgreich 
                                        childNode.Checked = node.Checked
                                    Else
                                        ' nicht erfolgreich
                                        ' keine Änderung von childNode.checked ... 
                                    End If
                                    ' Liste aktualisieren ...
                                    writeProtections.liste = request.retrieveWriteProtectionsFromDB(AlleProjekte)
                                    Call bestimmeNodeAppearance(childNode, aKtionskennung, PTTreeNodeTyp.pVariant, pName, vName, False)
                                Else
                                    atleastOneError = True
                                End If
                            Next

                            ' jetzt korrigieren, wenn eines der Kinder nicht auf den gleichen Check-Status gesetzt werden konnte
                            If atleastOneError And node.Nodes.Count > 0 Then
                                node.Checked = Not node.Checked
                            End If

                        End If


                    Case 1 ' Variante ist selektiert / nicht selektiert

                        ' ANfang 

                        parentNode = node.Parent
                        ' darf es ge- bzw. entcheckt werden ? 
                        Dim pName As String = getProjectNameOfTreeNode(parentNode.Text)
                        Dim vName As String = getVariantNameOfTreeNode(node.Text)
                        Dim pvName As String = calcProjektKey(pName, vName)

                        If Not (writeProtections.isProtected(pvName)) Or _
                            (writeProtections.isProtected(pvName) And dbUsername = writeProtections.wasProtectedBy(pvName)) Then

                            ' jetzt in der Datenbank setzen 
                            Dim wpItem As New clsWriteProtectionItem

                            With wpItem
                                .isProtected = node.Checked
                                .permanent = Me.chkbxPermanent.Checked
                                .userName = dbUsername
                                .pvName = pvName
                            End With

                            If request.setWriteProtection(wpItem) Then
                                ' erfolgreich 
                                ' keine Änderung von node.checked nötig 
                            Else
                                ' nicht erfolgreich
                                node.Checked = Not node.Checked
                            End If
                            ' Liste aktualisieren ...
                            writeProtections.liste = request.retrieveWriteProtectionsFromDB(AlleProjekte)
                            Call bestimmeNodeAppearance(node, aKtionskennung, PTTreeNodeTyp.pVariant, pName, vName, False)
                        End If

                        ' Ende


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

                    ' jetzt die Variante aktivieren 
                    Call replaceProjectVariant(pName, selectedVariantName, True, True, 0)

                    ' jetzt das Browser Szenario aktualsieren 
                    currentBrowserConstellation.updateShowAttributes(pName)

                    ' jetzt die Charts , Einzel- wie Multiprojekt-Charts aktualisieren 
                    Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                    Call aktualisiereCharts(hproj, True)
                    Call awinNeuZeichnenDiagramme(2)

                    ' jetzt den Text des ParentNodes aktualisieren  
                    Call bestimmeNodeAppearance(projektNode, aKtionskennung, PTTreeNodeTyp.project, pName, selectedVariantName, False)

            End Select

            stopRecursion = False

        ElseIf aKtionskennung = PTTvActions.chgInSession Then

            stopRecursion = True

            Select Case treeLevel

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

                        Call putProjectInShow(pName, selectedVariantName, considerDependencies, False)
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

                    
                    ' jetzt das Browser Szenario aktualisieren 
                    currentBrowserConstellation.updateShowAttributes()

                    ' jetzt müssen die Portfolio Diagramme neu gezeichnet werden 
                    Call awinNeuZeichnenDiagramme(2)

                    ' jetzt den Text des Projekt-Knotens aktualisieren  
                    Call bestimmeNodeAppearance(node, aKtionskennung, PTTreeNodeTyp.project, pName, selectedVariantName, node.Nodes.Count = 0)


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

                        ' '' die Standard Variante auf Checked setzen 
                        ''For i = 0 To projektNode.Nodes.Count - 1
                        ''    If projektNode.Nodes.Item(i).Text = "()" Then
                        ''        projektNode.Nodes.Item(i).Checked = True
                        ''    End If
                        ''Next

                        ' jetzt die selektierte Variante ins ShowProjekte stecken und aktualisieren ... 
                        ' aber nur, wenn es nicht vorher schon die leere Variante war 


                    End If

                    ' jetzt muss das bisherige aus ShowProjekte rausgenommen werden 
                    If ShowProjekte.contains(pName) And projektNode.Checked Then

                        Call replaceProjectVariant(pName, selectedVariantName, False, True, 0)

                        ' jetzt das Browser Szenario aktualsieren 
                        currentBrowserConstellation.updateShowAttributes(pName)

                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                        Call aktualisiereCharts(hproj, True)
                        Call awinNeuZeichnenDiagramme(2)

                    End If

                    ' jetzt den Text des ParentNodes aktualisieren  
                    Call bestimmeNodeAppearance(projektNode, aKtionskennung, PTTreeNodeTyp.project, pName, selectedVariantName, False)

            End Select

            stopRecursion = False

        End If


    End Sub

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
    Private Sub TreeViewProjekte_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeViewProjekte.AfterSelect

        Dim node As TreeNode = e.Node
        Dim treeLevel As Integer = node.Level
        Dim projectName As String
        Dim variantName As String = ""
        Dim toolTippText As String = "-"
        Dim hproj As clsProjekt


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

            'If IsNothing(hproj) And variantNames.Count > 0 Then
            '    variantName = CStr(variantNames.Item(1))
            '    hproj = AlleProjekte.getProject(projectName, variantName)
            'End If

            ' jetzt muss bestimmt werden, was als ToolTipp Text angezeigt werden soll 
            If allDependencies.projectCount > 0 And toolTippsAreShowing = ptPPAtooltipps.dependencies Then
                toolTippText = allDependencies.getDependencyInfos(projectName)

            ElseIf toolTippsAreShowing = ptPPAtooltipps.protectedBy Then

                If variantNames.Count = 1 Then
                    Dim pvName As String = calcProjektKey(projectName, projectName)
                    Dim lastUser As String = ""
                    Dim zeitpunkt As Date
                    lastUser = writeProtections.wasProtectedBy(pvName)
                    zeitpunkt = writeProtections.changeDate(pvName)

                    If writeProtections.isProtected(pvName) Then
                        toolTippText = "protected by: " & lastUser & ", at: " & zeitpunkt.ToShortDateString
                    Else
                        toolTippText = "no protection"
                    End If

                Else
                    toolTippText = ""
                End If
                

            Else
                If Not IsNothing(hproj) Then
                    If hproj.description.Length > 0 Then
                        toolTippText = hproj.description
                    End If
                End If
            End If

            ' tk, 2.1.17 Anzeige der Info zu diesem Projekt ... 
            If Not IsNothing(hproj) Then
                Call aktualisierePMSForms(hproj)
                Call aktualisiereCharts(hproj, True)
            End If


        ElseIf treeLevel = 1 Then
            Dim projectNode As TreeNode = node.Parent
            If Not IsNothing(projectNode) Then

                projectName = getProjectNameOfTreeNode(projectNode.Text)
                variantName = getVariantNameOfTreeNode(node.Text)
                hproj = AlleProjekte.getProject(projectName, variantName)

                If Not IsNothing(hproj) Then

                    If toolTippsAreShowing = ptPPAtooltipps.protectedBy Then


                        Dim pvName As String = calcProjektKey(projectName, projectName)
                        Dim lastUser As String = ""
                        Dim zeitpunkt As Date
                        lastUser = writeProtections.wasProtectedBy(pvName)
                        zeitpunkt = writeProtections.changeDate(pvName)

                        If writeProtections.isProtected(pvName) Then
                            toolTippText = "protected by: " & lastUser & ", at: " & zeitpunkt.ToShortDateString
                        Else
                            toolTippText = "no protection"
                        End If


                    Else
                        If hproj.variantDescription.Length > 0 Then
                            toolTippText = hproj.variantDescription
                        End If
                    End If


                    ' Anzeige der aktualisierten Charts und Phasen- bzw Milestone Infor Formulare 
                    Call aktualisierePMSForms(hproj)
                    Call aktualisiereCharts(hproj, True)

                End If

            End If
        End If

        ToolTipStand.Show(toolTippText, TreeViewProjekte, 6000)


    End Sub

    Private Sub TreeViewProjekte_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeViewProjekte.BeforeExpand

        ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim selectedNode As New TreeNode
        Dim variantNode As New TreeNode
        Dim nodeTimeStamp As New TreeNode
        Dim projName As String = ""
        Dim variantName As String = ""
        Dim hliste As SortedList(Of Date, String)
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

        If Not noDB Then
            ' jetzt die writeProtections neu bestimmen 
            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            writeProtections.liste = request.retrieveWriteProtectionsFromDB(AlleProjekte)
        End If

        selectedNode = e.Node
        nodeLevel = e.Node.Level

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

                ' Eintragen der zum Projekt gehörenden Varianten
                For Each variantName In variantListe
                    variantNode = selectedNode.Nodes.Add(CType(variantName, String))

                    ' jetzt muss gecheckt werden , ob es sich um das Aktivieren handelt oder nicht
                    If aKtionskennung = PTTvActions.activateV Or _
                        aKtionskennung = PTTvActions.chgInSession Then
                        stopRecursion = True
                        If getVariantNameOfTreeNode(variantName) = hproj.variantName Then
                            variantNode.Checked = True
                        Else
                            variantNode.Checked = False
                        End If
                        stopRecursion = False

                    ElseIf aKtionskennung = PTTvActions.loadPV Then
                        ' es können alle Elemente selektiert werden ...

                        key = calcProjektKey(pName:=projName, variantName:=variantName)

                        stopRecursion = True
                        ' soll gesetzt sein, wenn es entweder bereits geladen ist oder aber sowieso alle geladen werden sollen
                        If AlleProjekte.Containskey(key) Or selectedNode.Checked = True Then
                            variantNode.Checked = True
                        Else
                            variantNode.Checked = False
                        End If
                        stopRecursion = False

                    ElseIf aKtionskennung = PTTvActions.delAllExceptFromDB Then
                        ' es können alle Elemente selektiert werden ...
                        stopRecursion = True
                        variantNode.Checked = selectedNode.Checked
                        stopRecursion = False

                    ElseIf aKtionskennung = PTTvActions.setWriteProtection And Not noDB Then
                        
                        variantName = getVariantNameOfTreeNode(variantName)

                        Dim pvName As String = calcProjektKey(projName, variantName)
                        stopRecursion = True
                        variantNode.Checked = writeProtections.isProtected(pvName)
                        Call bestimmeNodeAppearance(variantNode, aKtionskennung, PTTreeNodeTyp.pVariant, projName, variantName, selectedNode.Nodes.Count = 1)
                        stopRecursion = False

                    Else
                        stopRecursion = True
                        variantNode.Checked = selectedNode.Checked
                        stopRecursion = False
                    End If



                    If aKtionskennung = PTTvActions.delFromDB Or _
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
                If (selectedNode.Checked And aKtionskennung = PTTvActions.chgInSession) Or _
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
                        Call bestimmeNodeAppearance(tmpNode, aKtionskennung, PTTreeNodeTyp.pVariant, projName, variantName, selectedNode.Nodes.Count = 1)
                    Next

                    stopRecursion = False
                End If
            End If



        ElseIf nodeLevel = 1 And _
            (aKtionskennung = PTTvActions.delFromDB Or aKtionskennung = PTTvActions.loadPVS) Then

            ' hier wurde eine Variante selektiert ...

            If selectedNode.Tag = "P" Then

                selectedNode.Tag = "X"
                projName = getProjectNameOfTreeNode(selectedNode.Parent.Text)
                variantName = getVariantNameOfTreeNode(selectedNode.Text)

                hliste = projektHistorien.getTimeStamps(calcProjektKey(projName, variantName))

                If hliste.Count = 0 Then

                    If Not noDB Then

                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                        If request.pingMongoDb() Then
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

                            projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=projName, variantName:=variantName, _
                                                                             storedEarliest:=Date.MinValue, storedLatest:=Date.Now)

                        Catch ex As Exception
                            projekthistorie.clear()
                        End Try

                    End If

                    If projekthistorie.Count > 0 Then

                        projektHistorien.Remove(projName & "#" & variantName, Date.MinValue) 'Platzhalter wieder entfernen
                        selectedNode.Nodes.Clear()  ' Löschen von Platzhalter

                        ' Aufbau der Listen 
                        projektHistorien.Add(projekthistorie)

                        stopRecursion = True
                        ' Eintragen der zur Projekt-Variante gehörenden TimeStamps
                        For Each kvp1 As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                            nodeTimeStamp = selectedNode.Nodes.Add(CType(kvp1.Value.timeStamp, String))
                            nodeTimeStamp.Checked = selectedNode.Checked
                        Next kvp1
                        stopRecursion = False

                    Else

                        If projekthistorie.Count = 0 Then
                            ' keine ProjektHistorie vorhanden
                            projektHistorien.Remove(projName & "#" & variantName, Date.MinValue) 'Platzhalter wieder entfernen
                            selectedNode.Nodes.Clear()  ' Löschen von Platzhalter
                        End If
                    End If




                End If

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
        Dim pname As String, variantName As String, timestamp As Date
        'Dim hproj As clsProjekt
        Dim portfolioZeile As Integer = 2
        Dim storedAtOrBefore As Date
        Dim considerDependencies As Boolean

        Dim outPutCollection As New Collection
        Dim outPutHeader As String = ""
        Dim outPutExplanation As String = ""


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
        If aKtionskennung = PTTvActions.delFromDB Then

            If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                outPutHeader = "Projekt-Varianten können nicht gelöscht werden !"
                outPutExplanation = "folgende Projekt-Varianten werden aktuell in Szenarien referenziert" & vbLf & _
                                    "und können daher nicht gelöscht werden:"
            Else
                outPutHeader = "Project-Variants can not be deleted !"
                outPutExplanation = "following project-variants are referenced in scenarios " & vbLf & _
                                    "and con not be deleted:"
            End If


        End If




        Dim p As Integer, v As Integer, t As Integer

        If aKtionskennung = PTTvActions.delFromDB Or _
            aKtionskennung = PTTvActions.delAllExceptFromDB Or _
            aKtionskennung = PTTvActions.delFromSession Or _
            aKtionskennung = PTTvActions.deleteV Or _
            aKtionskennung = PTTvActions.loadPV Then

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
                                Call deleteCompleteProjectVariant(outPutCollection, _
                                                                  pname, variantName, aKtionskennung, versionsToKeep.Value)

                            Next


                        ElseIf aKtionskennung = PTTvActions.delFromDB Then


                            For v = 1 To anzahlVarianten

                                variantName = getVariantNameOfTreeNode(CStr(variantListe.Item(v)))
                                ' Fehler-Behandlung, d.h auch Abfrage ob PName#vName referenziert in Szenario ist, passiert dort drin ... 
                                Call deleteCompleteProjectVariant(outPutCollection, _
                                                                  pname, variantName, aKtionskennung)


                            Next


                        ElseIf aKtionskennung = PTTvActions.loadPV Then

                            Dim hproj As clsProjekt = Nothing
                            If ShowProjekte.Count > 0 Then
                                If ShowProjekte.contains(pname) Then
                                    hproj = ShowProjekte.getProject(pname)
                                End If
                            End If

                            For v = 1 To anzahlVarianten

                                'variantNode = projektNode.Nodes.Item(v - 1)
                                'variantName = getVariantNameOf(variantNode.Text)
                                variantName = getVariantNameOfTreeNode(CStr(variantListe.Item(v)))

                                Dim showAttribute As Boolean
                                If IsNothing(hproj) Then
                                    If v = 1 Then
                                        showAttribute = True
                                    Else
                                        showAttribute = False
                                    End If
                                Else
                                    If variantName = hproj.variantName Then
                                        showAttribute = True
                                    Else
                                        showAttribute = False
                                    End If
                                End If

                                Call loadProjectfromDB(outPutCollection, pname, variantName, showAttribute, storedAtOrBefore)

                                If currentBrowserConstellation.contains(calcProjektKey(pname, variantName), False) Then
                                    ' nichts tun , ist schon drin 
                                    currentBrowserConstellation.getItem(calcProjektKey(pname, variantName)).show = showAttribute
                                Else
                                    Dim cItem As New clsConstellationItem
                                    With cItem
                                        .projectName = pname
                                        .variantName = variantName
                                        .show = showAttribute
                                    End With
                                    currentBrowserConstellation.add(cItem)
                                End If


                            Next



                        End If




                    ElseIf projektNode.Tag = "X" Then

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
                                    Call deleteCompleteProjectVariant(outPutCollection, _
                                                                      pname, variantName, aKtionskennung)



                                ElseIf aKtionskennung = PTTvActions.delAllExceptFromDB Then

                                    ' hier muss ja gar kein Check auf Szenario referenz erfolgen, da ohnehin immer min 2 Stände behalten werdne  
                                    Call deleteCompleteProjectVariant(outPutCollection, _
                                                                      pname, variantName, aKtionskennung, versionsToKeep.Value)

                                ElseIf aKtionskennung = PTTvActions.delFromSession Or _
                                        aKtionskennung = PTTvActions.deleteV Then

                                    Call awinDeleteProjectInSession(pName:=pname, considerDependencies:=considerDependencies, vName:=variantName)

                                    ' jetzt in der currentBrowserConstellation ändern 
                                    Dim tmpKey As String = calcProjektKey(pname, variantName)
                                    currentBrowserConstellation.remove(tmpKey)


                                ElseIf aKtionskennung = PTTvActions.loadPV Then

                                    Call loadProjectfromDB(outPutCollection, pname, variantName, first, storedAtOrBefore)
                                    first = False

                                    If currentBrowserConstellation.contains(calcProjektKey(pname, variantName), False) Then
                                        ' nichts tun , ist schon drin 
                                    Else
                                        Dim cItem As New clsConstellationItem
                                        With cItem
                                            .projectName = pname
                                            .variantName = variantName
                                            .show = (v = 1)
                                        End With
                                        currentBrowserConstellation.add(cItem)
                                    End If


                                End If


                            ElseIf aKtionskennung = PTTvActions.delFromDB Or _
                                    aKtionskennung = PTTvActions.loadPVS Then

                                anzahlTimeStamps = variantNode.Nodes.Count
                                Dim firstTS As Boolean = True
                                For t = 1 To anzahlTimeStamps
                                    timeStampNode = variantNode.Nodes.Item(t - 1)

                                    If timeStampNode.Checked Then
                                        ' Aktion auf diesem timestamp

                                        timestamp = CType(timeStampNode.Text, Date)
                                        If aKtionskennung = PTTvActions.delFromDB Then
                                            Call deleteProjectVariantTimeStamp(outPutCollection, _
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

                If aKtionskennung = PTTvActions.loadPV Or _
                    aKtionskennung = PTTvActions.delFromSession Then
                    Call awinNeuZeichnenDiagramme(2)
                End If

            End With

            ' bei Aktionen loadPV, delFromSession muss der currentConstellationName aktualisiert werden 
            If aKtionskennung = PTTvActions.delFromSession Or _
                aKtionskennung = PTTvActions.loadPV Or _
                aKtionskennung = PTTvActions.deleteV Then
                If Not currentConstellationName.EndsWith("(*)") Then
                    currentConstellationName = currentConstellationName & " (*)"
                End If

                Call storeSessionConstellation("Last")
            End If


            ' jetzt ggf die Outputs anzeigen 
            If outPutCollection.Count > 0 Then
                Call showOutPut(outPutCollection, _
                                outPutHeader, _
                                outPutExplanation)
            End If

            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()

        ElseIf aKtionskennung = PTTvActions.chgInSession Then

            If dropboxScenarioNames.Text <> "" Then


                currentConstellationName = dropboxScenarioNames.Text
                'currentBrowserConstellation.constellationName = currentConstellationName

                Dim toStoreConstellation As clsConstellation = _
                    currentBrowserConstellation.copy(currentConstellationName)

                ' Korrektheitsprüfung
                If awinSettings.visboDebug Then
                    toStoreConstellation.checkAndCorrectYourself()
                End If

                projectConstellations.update(toStoreConstellation)

                Dim txtMsg1 As String = ""
                Dim txtMsg2 As String = ""
                If storeToDBasWell.Checked Then
                    Call storeSingleConstellationToDB(outPutCollection, toStoreConstellation)

                    ' jetzt ggf die Outputs anzeigen 

                    If outPutCollection.Count > 0 Then

                        If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                            txtMsg1 = "Speichern Szenario " & toStoreConstellation.constellationName
                            txtMsg2 = "folgende Probleme sind aufgetreten:"
                        Else
                            txtMsg1 = "Store Scenario " & toStoreConstellation.constellationName
                            txtMsg2 = "following problems occurred:"
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



    End Sub


    ''' <summary>
    ''' alle dargestellten Elemente im ProjektTree selektieren 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SelectionSet_Click(sender As Object, e As EventArgs) Handles SelectionSet.Click


        Dim projectNode As TreeNode

        stopRecursion = True

        With TreeViewProjekte

            ' die Behandlung von chgInSession ist etwas anders, weil sofort eine Aktion erfolgen muss ... 

            If aKtionskennung = PTTvActions.chgInSession Then

                ' jetzt im formular den Mauszeiger auf Warten ... setzen 
                Me.Cursor = Cursors.WaitCursor

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

                ' jetzt müssen die Show Attribute und die Zeilen neu gesetzt werden ...
                currentBrowserConstellation.updateShowAttributes()

                ' jetzt muss die Plan-Tafel gelöscht werden 
                Call awinClearPlanTafel()

                ' jetzt muss die Plan-Tafel neu gezeichnet werden 
                Call awinZeichnePlanTafelNeu(True)

                ' jetzt müssen die Diagramme neu gezeichnet werden 
                Call awinNeuZeichnenDiagramme(2)

                Me.Cursor = Cursors.Default

            ElseIf aKtionskennung = PTTvActions.deleteV Or _
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
                    txtMsg = "beim Löschen nicht zulässig ..."
                Else
                    txtMsg = "not allowed option ..."
                End If
                Call MsgBox(txtMsg)


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

        If aKtionskennung = PTTvActions.chgInSession Or _
            aKtionskennung = PTTvActions.activateV Then
            If Not currentConstellationName.EndsWith("(*)") Then
                currentConstellationName = currentConstellationName & " (*)"
                Dim preText As String = "Szenario "
                If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                    preText = "Scenario "
                End If

                Me.Text = preText & currentConstellationName
            End If


        End If

        stopRecursion = False

    End Sub

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

        Dim projectNode As TreeNode

        stopRecursion = True

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

                    If projectNode.Checked Then

                        projectNode.Checked = False

                    Else
                        ' nichts tun , denn das Projekt wird bereits angezeigt und ist in Showprojekte drin 
                    End If

                Next

                ' jetzt muss die Plan-Tafel gelöscht werden 
                Call awinClearPlanTafel()

                ' jetzt muss Showprojekte gelöscht werden 
                ShowProjekte.Clear()

                ' jetzt müssen die Show Attribute und die Zeilen neu gesetzt werden ...
                currentBrowserConstellation.updateShowAttributes()

                ' jetzt müssen die Diagramme neu gezeichnet werden 
                Call awinNeuZeichnenDiagramme(2)

            ElseIf aKtionskennung = PTTvActions.activateV Then
                ' nichts tun, Alle Resetten macht bei diesen keinen Sinn 

            Else
                ' auch in den Fällen deleteV
                ' in allen anderen Fällen: loadPV, loadPVS, delFromDB, delAllExceptFromDB, delFromSession

                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)
                    If projectNode.Checked Then
                        projectNode.Checked = False
                    End If

                    If projectNode.Nodes.Count > 0 Then
                        Call unCheck(projectNode)
                    End If
                Next

            End If

        End With

        If aKtionskennung = PTTvActions.chgInSession Or _
            aKtionskennung = PTTvActions.activateV Then
            If Not currentConstellationName.EndsWith("(*)") Then
                currentConstellationName = currentConstellationName & " (*)"

                Dim preText As String = "Szenario "
                If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                    preText = "Scenario "
                End If
                Me.Text = preText & currentConstellationName
            End If


        End If

        Me.Cursor = Cursors.Default
        stopRecursion = False


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

        Dim filterFormular As New frmNameSelection
        Dim considerDependencies As Boolean
        Dim zeitraumVon As Date = StartofCalendar
        Dim zeitraumBis As Date = StartofCalendar.AddYears(20)
        Dim storedGestern As Date = StartofCalendar

        ' hier ist der einzige Grund für browserAlleProjekte: es muss etwas da sein, wo reingeladen werden kann 
        ' wenn auf der Datenbank gefiltert werden soll - und das geht nur , in dem etwas geladen wird ... 
        Dim browserAlleProjekte As New clsProjekteAlle

        If Not currentConstellationName.EndsWith("(*)") Then
            currentConstellationName = currentConstellationName & " (*)"

            Dim preText As String = "Szenario "
            If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                preText = "Scenario "
            End If
            Me.Text = preText & currentConstellationName
        End If


        If IsNothing(browserConstellationSav) Then
            browserConstellationSav = currentBrowserConstellation.copy
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
        If quickList Or _
            aKtionskennung = PTTvActions.delFromDB Or _
            aKtionskennung = PTTvActions.delAllExceptFromDB Or _
            aKtionskennung = PTTvActions.loadPV Then

            If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                ' es ist ein Zeitraum definiert 
                zeitraumVon = getDateofColumn(showRangeLeft, False)
                zeitraumBis = getDateofColumn(showRangeRight, True)
            End If
            ' es muss die Gesamtliste aufgebaut werden ... das dauert jetzt erst mal 
            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            'Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

            Dim pname As String = ""
            Dim variantName As String = ""

            'jetzt wird die aktuelleGesamtListe aufgebaut; sobald die mal aufgebaut wurde, muss sie nicht wieder aufgebaut werden ... 
            ' tk das applyFilter wird nachher gemacht , ausnahmslos für alle 
            If Not browserAlleProjekte.Count = 0 Then
                browserAlleProjekte.Clear()
            End If
            browserAlleProjekte.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumBis, storedGestern, storedAtOrBefore, True)
            quickList = False

        Else
            ' browserAlleProjekte bestimmen  
            browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)
        End If

        With filterFormular
            If aKtionskennung = PTTvActions.loadPV Or _
                aKtionskennung = PTTvActions.loadPVS Or _
                aKtionskennung = PTTvActions.delAllExceptFromDB Or _
                aKtionskennung = PTTvActions.delFromDB Then
                ' damit im Filterformular unterschieden werden kann, ob der Aufruf aus dem ProjPortfolioAdmin Formular erfolgte ...
                .actionCode = aKtionskennung
                .menuOption = PTmenue.filterdefinieren
            Else
                .actionCode = aKtionskennung
                .menuOption = PTmenue.sessionFilterDefinieren
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then

                stopRecursion = True

                Me.Cursor = Cursors.WaitCursor
                Dim filter As clsFilter = filterDefinitions.retrieveFilter("Last")
                Dim ok As Boolean

                If aKtionskennung = PTTvActions.loadPV Or _
                    aKtionskennung = PTTvActions.delAllExceptFromDB Or _
                    aKtionskennung = PTTvActions.delFromDB Or _
                    aKtionskennung = PTTvActions.chgInSession Then

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

                    ' jetzt müssen die tfZeile neu besetzt werden;
                    '  nach standard, d.h 0 bedeutet einfach sortiert nach Name 
                    currentBrowserConstellation.setTfZeilen(0)

                    If removeList.Count > 0 Then
                        Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, _
                                            aKtionskennung, quickList)

                        If aKtionskennung = PTTvActions.chgInSession Then
                            ' erst am Ende alle Diagramme neu machen ...


                            If removeList.Count > 0 Then
                                Dim tmpConstellation As New clsConstellations
                                tmpConstellation.Add(currentBrowserConstellation)

                                Call showConstellations(constellationsToShow:=tmpConstellation, _
                                                        clearBoard:=True, clearSession:=False, storedAtOrBefore:=storedAtOrBefore)

                                If aKtionskennung = PTTvActions.chgInSession Then
                                    Call awinNeuZeichnenDiagramme(2)
                                End If

                            End If
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

        currentBrowserConstellation = browserConstellationSav.copy
        'Dim browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)
        browserConstellationSav = Nothing

        ' jetzt das entzsprechende Szenario wieder laden 
        Dim tmpConstellation As New clsConstellations
        tmpConstellation.Add(currentBrowserConstellation)

        Dim storedAtOrBefore As Date
        If IsNothing(requiredDate.Value) Then
            storedAtOrBefore = Date.Now
        Else
            storedAtOrBefore = requiredDate.Value
        End If


        Call showConstellations(constellationsToShow:=tmpConstellation, _
                                clearBoard:=True, clearSession:=False, storedAtOrBefore:=storedAtOrBefore)

        ' neu Zeichnen der Diagramme
        Call awinNeuZeichnenDiagramme(2)



        Me.Cursor = Cursors.WaitCursor

        

        ' jetzt muss der Last-Filter zurückgesetzt werden 
        Dim emptyCollection As New Collection
        Dim fName As String = "Last"

        Dim lastFilter As New clsFilter(fName, emptyCollection, emptyCollection, emptyCollection, _
                                        emptyCollection, emptyCollection, emptyCollection)
        filterDefinitions.storeFilter(fName, lastFilter)

        stopRecursion = True
        Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, aKtionskennung, False)
        'Call buildTreeview(projektHistorien, TreeViewProjekte, browserAlleProjekte, pvNamesList, _
        '                   aKtionskennung, quickList, Me.filterIsActive, storedAtOrBefore)
        stopRecursion = False

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
            ttText = "Szenario-Name auswählen oder neuen Namen eingeben"
        Else
            ttText = "Select scenario Name and/or edit new name"
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
    
    Private Sub ToolTipStand_Popup(sender As Object, e As PopupEventArgs) Handles ToolTipStand.Popup

    End Sub

    ''' <summary>
    ''' reduziert die Constellation auf alle Projekt-Varianten mit Attribut Show 
    ''' macht nur Sinn bei chgInSession; wird also nur von dort aus aufgerufen ... 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub onlyActive_Click(sender As Object, e As EventArgs) Handles onlyActive.Click

        If Not currentConstellationName.EndsWith("(*)") Then
            currentConstellationName = currentConstellationName & " (*)"

            Dim preText As String = "Szenario "
            If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                preText = "Scenario "
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

        If Not currentConstellationName.EndsWith("(*)") Then
            currentConstellationName = currentConstellationName & " (*)"

            Dim preText As String = "Szenario "
            If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                preText = "Scenario "
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
    Private Sub modifyTreeviewToShowAttribute(ByVal showKennung As Integer, _
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
            currentBrowserConstellation.setTfZeilen(0)

            Dim tmpConstellation As New clsConstellations
            tmpConstellation.Add(currentBrowserConstellation)

            ' auf der Multiprojekt-Tafel entsprechend anzeigen 
            Call showConstellations(constellationsToShow:=tmpConstellation, _
                                    clearBoard:=True, clearSession:=False, storedAtOrBefore:=storedAtOrBefore)

            ' den TreeView updaten ... 
            stopRecursion = True
            Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, _
                                            aKtionskennung, quickList)
            stopRecursion = False

            ' die Diagramme aktualisieren 
            If aKtionskennung = PTTvActions.chgInSession Then
                Call awinNeuZeichnenDiagramme(2)
            End If

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
                Me.OKButton.Text = "Store to Session and DB"
            Else
                Me.OKButton.Text = "Store to Session"
            End If
        End If

    End Sub

    

    Private Sub requiredDate_ValueChanged(sender As Object, e As EventArgs) Handles requiredDate.ValueChanged

        stopRecursion = True

        Me.Cursor = Cursors.WaitCursor

        Dim storedAtOrBefore As Date

        If Not IsNothing(requiredDate) Then

            If requiredDate.Value >= earliestDate Then
                requiredDate.Value = requiredDate.Value.Date.AddHours(23).AddMinutes(59)
                storedAtOrBefore = requiredDate.Value
            Else

                Dim msgText As String = "es gibt vor dem " & earliestDate.ToShortDateString & " keine Projekte in der Datenbank "
                If menuCult.Name <> ReportLang(PTSprache.deutsch).Name Then
                    msgText = "there are no projects at or before " & earliestDate.ToShortDateString & " in the database"
                End If
                
                Call MsgBox(msgText)

                requiredDate.Value = Date.Now.Date.AddHours(23).AddMinutes(59)
                storedAtOrBefore = Date.Now.Date.AddHours(23).AddMinutes(59)
            End If
            
        Else
            requiredDate.Value = Date.Now.Date.AddHours(23).AddMinutes(59)
            storedAtOrBefore = Date.Now.Date.AddHours(23).AddMinutes(59)
        End If

        If aKtionskennung = PTTvActions.loadPV Or _
            aKtionskennung = PTTvActions.delFromDB Then

            pvNamesList = buildPvNamesList(storedAtOrBefore)
            quickList = True
        End If

        Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, aKtionskennung, quickList)

        stopRecursion = False

        Me.Cursor = Cursors.Default

        ' Fokus an TreeViewPRojekte geben 
        TreeViewProjekte.Focus()
    End Sub

   
End Class