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
    
    ' wenn aus der Datenbank schnell gelesen werden soll ..
    Private pvNamesList As New SortedList(Of String, String)
    Private quickList As Boolean

    Private projektHistorien As New clsProjektDBInfos
    Private stopRecursion As Boolean = False
    Private constellationName As String = ""

    Private filterIsActive As Boolean = False
    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection
    Private selectedCosts As New Collection
    Private selectedRoles As New Collection
    Private selectedBUs As New Collection
    Private selectedTyps As New Collection

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

        ' je nachdem, ob es überhaupt Abhäbgigkeiten gibt, wird das angezeigt ..
        If allDependencies.projectCount > 0 Then
            Me.LblToolTipps.Visible = True
            Me.rdbTTDescription.Visible = True
            Me.rdbTTDescription.Checked = True

            Me.rdbTTDependencies.Visible = True
        Else
            Me.LblToolTipps.Visible = False
            Me.rdbTTDescription.Visible = False
            Me.rdbTTDescription.Checked = True

            Me.rdbTTDependencies.Visible = False
        End If

        With Me

            ' bei Beginn immer disabled
            .deleteFilterIcon.Enabled = False

            If aKtionskennung = PTTvActions.activateV Then

                .Text = "Variante aktivieren"

                .dropBoxTimeStamps.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = False
                .SelectionReset.Visible = False

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False
                .OKButton.Visible = False


            ElseIf aKtionskennung = PTTvActions.chgInSession Then
                '.Text = "Zusammenstellung im Szenario ändern"
                .Text = "Modify Multi-Project Scenario "

                .dropBoxTimeStamps.Visible = False
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
                .OKButton.Text = "Store Scenario"
                Dim testName As String = .OKButton.Name


            ElseIf aKtionskennung = PTTvActions.deleteV Then

                .Text = "Variante löschen"

                .dropBoxTimeStamps.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False

                .OKButton.Visible = True
                .OKButton.Text = "Löschen"

            ElseIf aKtionskennung = PTTvActions.delFromDB Then

                .Text = "Projekte, Varianten bzw. Snapshots in der Datenbank löschen"

                .dropBoxTimeStamps.Visible = True
                .lblStandvom.Visible = True

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False

                .OKButton.Visible = True
                .OKButton.Text = "Löschen"


            ElseIf aKtionskennung = PTTvActions.delFromSession Then
                .Text = "Projekte, Varianten aus der Session löschen"

                .dropBoxTimeStamps.Visible = False
                .lblStandvom.Visible = False

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = False
                .deleteFilterIcon.Visible = False

                .dropboxScenarioNames.Visible = False

                .OKButton.Visible = True
                .OKButton.Text = "Löschen"

            ElseIf aKtionskennung = PTTvActions.loadPV Then

                .Text = "Projekte und Varianten in die Session laden "

                .dropBoxTimeStamps.Visible = True
                .lblStandvom.Visible = True

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = True
                .deleteFilterIcon.Visible = True

                .dropboxScenarioNames.Visible = False


                .OKButton.Visible = True
                .OKButton.Text = "Laden"


            ElseIf aKtionskennung = PTTvActions.loadPVS Then

                .Text = "Projekte und Varianten in die Session laden "

                .dropBoxTimeStamps.Visible = True
                .lblStandvom.Visible = True

                .SelectionSet.Visible = True
                .SelectionReset.Visible = True

                .collapseCompletely.Visible = True
                .expandCompletely.Visible = True

                .filterIcon.Visible = True
                .deleteFilterIcon.Visible = True

                .dropboxScenarioNames.Visible = False


                .OKButton.Visible = True
                .OKButton.Text = "Laden"

            End If

        End With


    End Sub


    Private Sub frmDefineEditPortfolio_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim browserAlleProjekte As New clsProjekteAlle

        If frmCoord(PTfrm.eingabeProj, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.top))
        End If

        If frmCoord(PTfrm.eingabeProj, PTpinfo.left) > 0 Then
            Me.Left = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.left))
        End If

        ' bestimmen, ob es sich um quicklist handelt ...
        If aKtionskennung = PTTvActions.loadPV Or _
            aKtionskennung = PTTvActions.delFromDB Then
            quickList = True
        Else
            quickList = False
        End If


        ' je nachdem, wie die Aktionskennung ist: setzen der Button Visibilitäten 
        Call defineButtonVisibility()

        ' jetzt muss bestimmt werden , was die aktuelle SessionConstellation ist 
        If projectConstellations.Contains(currentConstellation) And AlleProjekte.Count > 0 Then
            currentBrowserConstellation = projectConstellations.getConstellation(currentConstellation)
            browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

        ElseIf projectConstellations.Contains("Last") And AlleProjekte.Count > 0 Then
            currentBrowserConstellation = projectConstellations.getConstellation("Last")
            browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

        ElseIf AlleProjekte.Count > 0 Then
            browserAlleProjekte = AlleProjekte.createCopy
            currentBrowserConstellation = New clsConstellation(browserAlleProjekte, Nothing, "currentBrowser", ptSzenarioConsider.all)

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
                'Dim heute As String = Date.Now.ToString

                dropBoxTimeStamps.Items.Clear()

                For k As Integer = 1 To tCollection.Count
                    Dim tmpDate As Date = CDate(tCollection.Item(k))
                    dropBoxTimeStamps.Items.Add(tmpDate)
                Next

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
        Dim storedAtOrBefore As Date
        If IsNothing(dropBoxTimeStamps.SelectedItem) Then
            storedAtOrBefore = Date.Now
        Else
            storedAtOrBefore = CDate(dropBoxTimeStamps.SelectedItem)
        End If

        ' hier wird jetzt die Browser Gesamt-Liste bestimmt  
        Call buildTreeview(projektHistorien, TreeViewProjekte, browserAlleProjekte, pvNamesList, _
                           aKtionskennung, quickList, _
                           Me.filterIsActive, storedAtOrBefore)



        stopRecursion = False

        If browserAlleProjekte.liste.Count < 1 And pvNamesList.Count < 1 Then
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
        Dim treeLevel As Integer
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



        ' hier wird jetzt sichergestellt, daß nur die nach der aktuellen Aktion gültigen Checks gesetzt werden können
        ' vor allem muss überall dort, wo das Szenario mit diesem Check verändert wird, das currentBrowserSzenario geupdated werden ...
        ' mit Click in TreeView wird verändert: Activate Variant, ChgInSession 

        If aKtionskennung = PTTvActions.delFromDB Or _
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


        ElseIf aKtionskennung = PTTvActions.activateV Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    ' bei Aktivieren kann man Projekt nicht selektieren 
                    node.Checked = False

                Case 1 ' Variante ist selektiert / nicht selektiert


                    Dim projektNode As TreeNode = node.Parent
                    Dim selectedVariantName As String = node.Text
                    Dim pName As String = projektNode.Text

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
                        selectedVariantName = getVariantNameOf(node.Text)



                    Else

                        ' die Standard Variante auf Checked setzen 
                        ' bzw. was besser ist, den ersten Child-Knoten 
                        ' das funktioniert nämlich auch dann, wenn keine Variante mit Name "" existiert 
                        If projektNode.Nodes.Count > 0 Then
                            projektNode.Nodes.Item(0).Checked = True
                            selectedVariantName = getVariantNameOf(projektNode.Nodes.Item(0).Text)
                        Else
                            ' darf eigentlich gar nicht vorkommen 
                            selectedVariantName = ""
                        End If



                    End If

                    ' jetzt das CurrentBrowser Szenario aktualisieren 


                    ' jetzt die Variante aktivieren 
                    Call replaceProjectVariant(pName, selectedVariantName, True, True, 0)

                    ' jetzt das Browser Szenario aktualsieren 
                    currentBrowserConstellation.updateShowAttributes(pName)

                    ' jetzt die Charts , Einzel- wie Multiprojekt-Charts aktualisieren 
                    Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                    Call aktualisiereCharts(hproj, False)
                    Call awinNeuZeichnenDiagramme(2)




            End Select

            stopRecursion = False

        ElseIf aKtionskennung = PTTvActions.chgInSession Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    Dim pName As String = node.Text
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
                                selectedVariantName = getVariantNameOf(childNode.Text)
                            End If
                        Next

                        If Not selectionExisted And node.Nodes.Count > 0 Then
                            childNode = node.Nodes.Item(0)
                            childNode.Checked = True
                            selectedVariantName = getVariantNameOf(childNode.Text)
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

                    ' jetzt das Browser Szenario aktualsieren 
                    currentBrowserConstellation.updateShowAttributes(pName)

                    ' jetzt müssen die Portfolio Diagramme neu gezeichnet werden 
                    Call awinNeuZeichnenDiagramme(2)

                Case 1 ' Variante ist selektiert / nicht selektiert


                    Dim projektNode As TreeNode = node.Parent
                    Dim selectedVariantName As String = node.Text
                    Dim pName As String = projektNode.Text

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
                        selectedVariantName = getVariantNameOf(node.Text)



                    Else

                        ' die Standard Variante auf Checked setzen 
                        ' bzw. was besser ist, den ersten Child-Knoten 
                        ' das funktioniert nämlich auch dann, wenn keine Variante mit Name "" existiert 
                        If projektNode.Nodes.Count > 0 Then
                            projektNode.Nodes.Item(0).Checked = True
                            selectedVariantName = getVariantNameOf(projektNode.Nodes.Item(0).Text)
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
                        Call aktualisiereCharts(hproj, False)
                        Call awinNeuZeichnenDiagramme(2)

                    End If



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

            If curItem.Text = mprojectName Then
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

            If curItem.Text = dprojectName Then
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
            projectName = node.Text

            Dim variantNames As Collection = AlleProjekte.getVariantNames(projectName, False)
            variantName = ""


            hproj = AlleProjekte.getProject(projectName, variantName)
            If IsNothing(hproj) And variantNames.Count > 0 Then
                variantName = CStr(variantNames.Item(1))
                hproj = AlleProjekte.getProject(projectName, variantName)
            End If

            ' jetzt muss bestimmt werden, was als ToolTipp Text angezeigt werden soll 
            If allDependencies.projectCount > 0 And rdbTTDependencies.Checked Then
                toolTippText = allDependencies.getDependencyInfos(projectName)
            Else
                If Not IsNothing(hproj) Then
                    If hproj.description.Length > 0 Then
                        toolTippText = hproj.description
                    End If
                End If
            End If



        ElseIf treeLevel = 1 Then
            Dim projectNode As TreeNode = node.Parent
            If Not IsNothing(projectNode) Then

                projectName = projectNode.Text
                variantName = getVariantNameOf(node.Text)
                hproj = AlleProjekte.getProject(projectName, variantName)

                If Not IsNothing(hproj) Then

                    If hproj.variantDescription.Length > 0 Then
                        toolTippText = hproj.variantDescription
                    End If
                End If

            End If
        End If

        ToolTipStand.Show(toolTippText, TreeViewProjekte, 6000)


    End Sub

    Private Sub TreeViewProjekte_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeViewProjekte.BeforeExpand

        ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim node As New TreeNode
        Dim nodeVariant As New TreeNode
        Dim nodeTimeStamp As New TreeNode
        Dim projName As String = ""
        Dim variantName As String = ""
        Dim hliste As SortedList(Of Date, String)
        Dim nodeLevel As Integer
        Dim variantListe As Collection
        Dim hproj As New clsProjekt
        Dim key As String

        Dim browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

        node = e.Node
        nodeLevel = node.Level

        If nodeLevel = 0 Then

            projName = node.Text

            ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
            If node.Tag = "P" Then
                ' Inhalte der Sub-Nodes müssen neu aufgebaut werden 
                If quickList Then
                    variantListe = getVariantListeFromPVNames(pvNamesList, projName)
                Else
                    variantListe = browserAlleProjekte.getVariantNames(projName, True)
                End If

                ' hproj wird benötigt, um herauszufinden, welche Variante gerade aktiv ist
                If aKtionskennung = PTTvActions.activateV Or _
                    (aKtionskennung = PTTvActions.chgInSession And node.Checked) Then
                    hproj = ShowProjekte.getProject(projName)
                ElseIf aKtionskennung = PTTvActions.chgInSession Then
                    ' jetzt erst noch die Variante bestimmen ... 
                    variantName = ""
                    For j As Integer = 1 To node.Nodes.Count
                        nodeVariant = node.Nodes.Item(j - 1)
                        If nodeVariant.Checked Then
                            variantName = nodeVariant.Text
                        End If
                    Next

                    Dim tmpKey As String = calcProjektKey(projName, variantName)
                    hproj = AlleProjekte.getProject(tmpKey)

                End If


                ' Löschen von Platzhalter
                node.Nodes.Clear()

                ' Eintragen der zum Projekt gehörenden Varianten
                For Each variantName In variantListe
                    nodeVariant = node.Nodes.Add(CType(variantName, String))

                    ' jetzt muss gecheckt werden , ob es sich um das Aktivieren handelt oder nicht
                    If aKtionskennung = PTTvActions.activateV Or _
                        aKtionskennung = PTTvActions.chgInSession Then
                        stopRecursion = True
                        If getVariantNameOf(variantName) = hproj.variantName Then
                            nodeVariant.Checked = True
                        Else
                            nodeVariant.Checked = False
                        End If
                        stopRecursion = False

                    ElseIf aKtionskennung = PTTvActions.loadPV Then

                        key = calcProjektKey(pName:=projName, variantName:=variantName)

                        stopRecursion = True
                        ' soll gesetzt sein, wenn es entweder bereits geladen ist oder aber sowieso alle geladen werden sollen
                        If AlleProjekte.Containskey(key) Or node.Checked = True Then
                            nodeVariant.Checked = True
                        Else
                            nodeVariant.Checked = False
                        End If
                        stopRecursion = False

                    Else
                        nodeVariant.Checked = node.Checked
                    End If



                    If aKtionskennung = PTTvActions.delFromDB Or _
                        aKtionskennung = PTTvActions.loadPVS Then
                        ' Einfügen eines Platzhalters macht nur Sinn bei Snapshots löschen bzw. Snapshots laden 

                        nodeVariant.Tag = "P"
                        nodeVariant.Nodes.Add("()")
                    Else
                        nodeVariant.Tag = "X"
                    End If


                Next

                node.Tag = "X"



            End If



        ElseIf nodeLevel = 1 And _
            (aKtionskennung = PTTvActions.delFromDB Or aKtionskennung = PTTvActions.loadPVS) Then


            If node.Tag = "P" Then

                node.Tag = "X"
                projName = node.Parent.Text
                variantName = getVariantNameOf(node.Text)

                hliste = projektHistorien.getTimeStamps(calcProjektKey(projName, variantName))

                If hliste.Count = 0 Then

                    If Not noDB Then

                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                        If request.pingMongoDb() Then
                        Else
                            Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
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
                        node.Nodes.Clear()  ' Löschen von Platzhalter

                        ' Aufbau der Listen 
                        projektHistorien.Add(projekthistorie)


                        ' Eintragen der zur Projekt-Variante gehörenden TimeStamps
                        For Each kvp1 As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                            nodeTimeStamp = node.Nodes.Add(CType(kvp1.Value.timeStamp, String))
                            nodeTimeStamp.Checked = node.Checked
                        Next kvp1


                    Else

                        If projekthistorie.Count = 0 Then
                            ' keine ProjektHistorie vorhanden
                            projektHistorien.Remove(projName & "#" & variantName, Date.MinValue) 'Platzhalter wieder entfernen
                            node.Nodes.Clear()  ' Löschen von Platzhalter
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
    Private Function getVariantNameOf(ByVal nodeText As String) As String
        Dim tmpstr() As String
        Dim vName As String = ""

        tmpstr = nodeText.Split(New Char() {CChar("("), CChar(")")}, 3)
        If tmpstr.Length = 3 Then
            vName = tmpstr(1)
        End If

        getVariantNameOf = vName

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

        Dim browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

        If allDependencies.projectCount > 0 Then
            considerDependencies = True
        Else
            considerDependencies = False
        End If

        ' ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        ' ''Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

        If IsNothing(dropBoxTimeStamps.SelectedItem) Then
            storedAtOrBefore = Date.Now.AddDays(1)
        Else

            storedAtOrBefore = CDate(dropBoxTimeStamps.SelectedItem)

        End If

        Dim p As Integer, v As Integer, t As Integer

        If aKtionskennung = PTTvActions.delFromDB Or _
            aKtionskennung = PTTvActions.delFromSession Or _
            aKtionskennung = PTTvActions.deleteV Or _
            aKtionskennung = PTTvActions.loadPV Then

            ' alle anderen Aktionen wie Projekte aus Datenbank löschen , aus Session löschen, aus Datenbank laden  ... 
            With TreeViewProjekte
                anzahlProjekte = .Nodes.Count

                For p = 1 To anzahlProjekte

                    projektNode = .Nodes.Item(p - 1)
                    pname = projektNode.Text

                    If projektNode.Checked Then
                        ' Aktion auf allen Varianten und Timestamps 
                        ' Schleife über alle Varianten: 
                        ' lösche in Datenbank pname#vname

                        'anzahlVarianten = projektNode.Nodes.Count
                        Dim variantListe As New Collection

                        If quickList Then
                            variantListe = getVariantListeFromPVNames(pvNamesList, pname)
                        Else
                            variantListe = browserAlleProjekte.getVariantNames(pname, True)
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
                            


                        ElseIf aKtionskennung = PTTvActions.delFromDB Then


                            For v = 1 To anzahlVarianten

                                'variantNode = projektNode.Nodes.Item(v - 1)
                                'variantName = getVariantNameOf(variantNode.Text)
                                variantName = getVariantNameOf(CStr(variantListe.Item(v)))
                                Call deleteCompleteProjectVariant(pname, variantName, aKtionskennung)

                            Next


                        ElseIf aKtionskennung = PTTvActions.loadPV Then


                            For v = 1 To anzahlVarianten

                                'variantNode = projektNode.Nodes.Item(v - 1)
                                'variantName = getVariantNameOf(variantNode.Text)
                                variantName = getVariantNameOf(CStr(variantListe.Item(v)))

                                If v = 1 Then
                                    Call loadProjectfromDB(pname, variantName, True, storedAtOrBefore)
                                Else
                                    Call loadProjectfromDB(pname, variantName, False, storedAtOrBefore)
                                End If

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
                                


                            Next



                        End If




                    ElseIf projektNode.Tag = "X" Then

                        anzahlVarianten = projektNode.Nodes.Count
                        Dim first As Boolean = True

                        For v = 1 To anzahlVarianten
                            variantNode = projektNode.Nodes.Item(v - 1)
                            variantName = getVariantNameOf(variantNode.Text)


                            If variantNode.Checked Then
                                ' Aktion auf allen Timestamps
                                ' lösche in Datenbank das Objekt mit DB-Namen pname#vname

                                If aKtionskennung = PTTvActions.delFromDB Then
                                    Call deleteCompleteProjectVariant(pname, variantName, aKtionskennung)

                                ElseIf aKtionskennung = PTTvActions.delFromSession Or _
                                        aKtionskennung = PTTvActions.deleteV Then

                                    Call awinDeleteProjectInSession(pName:=pname, considerDependencies:=considerDependencies, vName:=variantName)

                                    ' jetzt in der currentBrowserConstellation ändern 
                                    Dim tmpKey As String = calcProjektKey(pname, variantName)
                                    currentBrowserConstellation.remove(tmpKey)


                                ElseIf aKtionskennung = PTTvActions.loadPV Then

                                    Call loadProjectfromDB(pname, variantName, first, storedAtOrBefore)
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
                                            Call deleteProjectVariantTimeStamp(pname, variantName, timestamp, firstTS)
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

            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()

        ElseIf aKtionskennung = PTTvActions.chgInSession Then

            If dropboxScenarioNames.Text <> "" Then

                currentConstellation = dropboxScenarioNames.Text
                currentBrowserConstellation.constellationName = currentConstellation
                projectConstellations.update(currentBrowserConstellation)
                ' alt 15.12.16
                'currentConstellation = dropboxScenarioNames.Text
                'Call storeSessionConstellation(currentConstellation)
            End If

            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()
        Else
            Call MsgBox("nicht unterstützte Option in ProjPortfolio Admin Formular ...")
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
                    Dim pName As String = projectNode.Text

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
                                variantName = getVariantNameOf(CStr(tmpCollection.Item(1)))
                            Else
                                variantName = ""
                            End If


                        ElseIf checkedVariants.Count = 1 Then
                            variantName = getVariantNameOf(CStr(checkedVariants.Item(1)))

                        ElseIf checkedVariants.Count > 1 Then
                            variantName = getVariantNameOf(CStr(checkedVariants.Item(1)))
                            For k As Integer = 1 To projectNode.Nodes.Count
                                If getVariantNameOf(projectNode.Nodes.Item(k - 1).Text) = variantName Then
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
                If Not IsNothing(dropBoxTimeStamps.SelectedItem) Then

                    Dim lookForTimestamp As Date = CDate(dropBoxTimeStamps.SelectedItem)
                    Dim vergleichsString As String = lookForTimestamp.ToString

                    ' jetzt wird der TreeView komplett expanded ...
                    stopRecursion = False
                    .ExpandAll()
                    stopRecursion = True

                    For i As Integer = 1 To .Nodes.Count
                        projectNode = .Nodes.Item(i - 1)

                        For v As Integer = 1 To projectNode.Nodes.Count
                            Dim variantNode As TreeNode = projectNode.Nodes.Item(v - 1)

                            For t As Integer = 1 To variantNode.Nodes.Count
                                Dim tsNode As TreeNode = variantNode.Nodes.Item(t - 1)
                                If tsNode.Text = vergleichsString Then
                                    tsNode.Checked = True
                                    variantNode.Checked = False
                                    projectNode.Checked = False

                                    If Not projectNode.IsExpanded Then
                                        projectNode.Expand()
                                    End If

                                    If Not variantNode.IsExpanded Then
                                        variantNode.Expand()
                                    End If
                                End If
                            Next
                        Next

                    Next
                Else
                    Call MsgBox("nur aktiv in Verbindung mit einem ausgewählten Stand")
                End If

            Else
                ' in allen anderen Fällen: loadPV, loadPVS, delFromDB, delFromSession

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

        If aKtionskennung = PTTvActions.chgInSession Then
            currentBrowserConstellation.updateShowAttributes()
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
                    Dim pName As String = projectNode.Text

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

                ' jetzt müssen die Diagramme neu gezeichnet werden 
                Call awinNeuZeichnenDiagramme(2)

            ElseIf aKtionskennung = PTTvActions.activateV Then
                ' nichts tun, Alle Resetten macht bei diesen keinen Sinn 

            Else
                ' auch in den Fällen deleteV
                ' in allen anderen Fällen: loadPV, loadPVS, delFromDB, delFromSession

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

        If aKtionskennung = PTTvActions.chgInSession Then
            currentBrowserConstellation.updateShowAttributes()
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

        Dim browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

        If IsNothing(browserConstellationSav) Then
            browserConstellationSav = currentBrowserConstellation.copy
        End If

        Dim storedAtOrBefore As Date
        If IsNothing(dropBoxTimeStamps.SelectedItem) Then
            storedAtOrBefore = Date.Now
        Else
            storedAtOrBefore = CDate(dropBoxTimeStamps.SelectedItem)
        End If

        If allDependencies.projectCount > 0 Then
            considerDependencies = True
        Else
            considerDependencies = False
        End If

        Me.filterIsActive = True

        Me.Cursor = Cursors.WaitCursor

        ' jetzt erst mal überprüfen, ob quicklist = true ..
        If quickList Or _
            aKtionskennung = PTTvActions.delFromDB Or _
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
                    aKtionskennung = PTTvActions.delFromDB Then

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

                    If removeList.Count > 0 Then
                        Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, _
                                            aKtionskennung, quickList)
                        ' muss bei Loadpv, delfromDB nicht gemacht werden ! 
                        'Call awinNeuZeichnenDiagramme(2)
                    End If


                Else
                    ' hier geht es um chgInSession, ...

                    ' als erstes aktuelles Szenario speichern 
                    'If browserConstellationOF.count = 0 Then
                    '    browserConstellationOF = New clsConstellation(AlleProjekte, Nothing, "browserOF", ptSzenarioConsider.all)
                    'End If

                    Dim noShowNames As Collection = getProjectNamesNotFittingToFilter("Last")

                    If noShowNames.Count > 0 Then
                        browserAlleProjekte.Clear()

                        For Each noShowName As String In noShowNames
                            Call putProjectInNoShow(noShowName, considerDependencies, False)
                        Next

                        ' jetzt die aktuelleGesamtListe aufbauen 
                        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

                            If Not filter.isEmpty Then
                                ok = filter.doesNotBlock(kvp.Value)
                            Else
                                ok = True
                            End If

                            If ok Then
                                ' in aktuelleGesamtListe aufnehmen - 
                                Try

                                    If Not browserAlleProjekte.Containskey(kvp.Key) Then
                                        browserAlleProjekte.Add(kvp.Key, kvp.Value)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else

                            End If

                        Next

                        Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, _
                                       aKtionskennung, False)

                        ' erst am Ende alle Diagramme neu machen ...
                        If noShowNames.Count > 0 Then
                            Call awinNeuZeichnenDiagramme(2)
                        End If
                    End If


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



    Private Sub dropBoxTimeStamps_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dropBoxTimeStamps.SelectedIndexChanged

        Dim browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)

        stopRecursion = True

        Me.Cursor = Cursors.WaitCursor

        Dim storedAtOrBefore As Date
        If IsNothing(dropBoxTimeStamps.SelectedItem) Then
            storedAtOrBefore = Date.Now
        Else
            storedAtOrBefore = CDate(dropBoxTimeStamps.SelectedItem)
        End If

        'Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, aKtionskennung, False)
        Call buildTreeview(projektHistorien, TreeViewProjekte, browserAlleProjekte, pvNamesList, _
                           aKtionskennung, quickList, Me.filterIsActive, storedAtOrBefore)

        stopRecursion = False

        Me.Cursor = Cursors.Default

        ' Fokus an TreeViewPRojekte geben 
        TreeViewProjekte.Focus()

    End Sub


    Private Sub dropboxScenarioNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dropboxScenarioNames.SelectedIndexChanged

    End Sub

    Private Sub deleteFilterIcon_Click(sender As Object, e As EventArgs) Handles deleteFilterIcon.Click

        currentBrowserConstellation = browserConstellationSav.copy
        Dim browserAlleProjekte = AlleProjekte.createCopy(filteredBy:=currentBrowserConstellation)
        browserConstellationSav = Nothing


        stopRecursion = True

        Me.Cursor = Cursors.WaitCursor

        Me.filterIsActive = False

        Dim storedAtOrBefore As Date
        If IsNothing(dropBoxTimeStamps.SelectedItem) Then
            storedAtOrBefore = Date.Now
        Else
            storedAtOrBefore = CDate(dropBoxTimeStamps.SelectedItem)
        End If

        ' jetzt muss der Last-Filter zurückgesetzt werden 
        Dim emptyCollection As New Collection
        Dim fName As String = "Last"

        Dim lastFilter As New clsFilter(fName, emptyCollection, emptyCollection, emptyCollection, _
                                        emptyCollection, emptyCollection, emptyCollection)
        filterDefinitions.storeFilter(fName, lastFilter)

        Call updateTreeview(TreeViewProjekte, currentBrowserConstellation, pvNamesList, aKtionskennung, False)
        'Call buildTreeview(projektHistorien, TreeViewProjekte, browserAlleProjekte, pvNamesList, _
        '                   aKtionskennung, quickList, Me.filterIsActive, storedAtOrBefore)
        stopRecursion = False

        ' Das DeleteFilterIcon mit Bild versehen 
        Me.deleteFilterIcon.Image = Nothing
        Me.deleteFilterIcon.Enabled = False

        Me.Cursor = Cursors.Arrow


    End Sub

    Private Sub dropBoxTimeStamps_MouseHover(sender As Object, e As EventArgs) Handles dropBoxTimeStamps.MouseHover
        ToolTipStand.Show("welcher Planungs-Stand soll geladen werden? Default ist immer der aktuelle Stand", dropBoxTimeStamps, 2000)
    End Sub


    Private Sub SelectionSet_MouseHover(sender As Object, e As EventArgs) Handles SelectionSet.MouseHover

        If aKtionskennung = PTTvActions.chgInSession Then
            ToolTipStand.Show("alle Projekte anzeigen", SelectionSet, 2000)
        ElseIf aKtionskennung = PTTvActions.delFromDB Then
            ToolTipStand.Show("alle Projekte und Projekt-Varianten auswählen, die den oben ausgewählten Zeitstempel haben", SelectionSet, 2000)
        ElseIf aKtionskennung = PTTvActions.loadPV Or aKtionskennung = PTTvActions.loadPVS Then
            ToolTipStand.Show("alle Projekte auswählen", SelectionSet, 2000)
        End If

    End Sub

    Private Sub SelectionReset_MouseHover(sender As Object, e As EventArgs) Handles SelectionReset.MouseHover
        ToolTipStand.Show("alle Elemente de-seletieren", SelectionReset, 2000)
    End Sub

    Private Sub collapseCompletely_MouseHover(sender As Object, e As EventArgs) Handles collapseCompletely.MouseHover
        ToolTipStand.Show("Baum-Struktur zusammenklappen", collapseCompletely, 2000)
    End Sub

    Private Sub expandCompletely_MouseHover(sender As Object, e As EventArgs) Handles expandCompletely.MouseHover
        ToolTipStand.Show("Baum-Struktur vollständig öffnen", expandCompletely, 2000)
    End Sub

    Private Sub filterIcon_MouseHover(sender As Object, e As EventArgs) Handles filterIcon.MouseHover
        ToolTipStand.Show("Filter definieren und anwenden", filterIcon, 2000)
    End Sub

    Private Sub deleteFilterIcon_MouseHover(sender As Object, e As EventArgs) Handles deleteFilterIcon.MouseHover
        ToolTipStand.Show("Filter löschen und zurücksetzen", deleteFilterIcon, 2000)
    End Sub


    Private Sub dropboxScenarioNames_MouseHover(sender As Object, e As EventArgs) Handles dropboxScenarioNames.MouseHover
        ToolTipStand.Show("Szenario-Name auswählen oder neuen Namen eingeben", dropboxScenarioNames, 2000)
    End Sub

    
    Private Sub TreeViewProjekte_MouseHover(sender As Object, e As EventArgs) Handles TreeViewProjekte.MouseHover

    End Sub

    Private Sub ToolTipStand_Popup(sender As Object, e As PopupEventArgs) Handles ToolTipStand.Popup

    End Sub

    Private Sub rdbTTDescription_CheckedChanged(sender As Object, e As EventArgs) Handles rdbTTDescription.CheckedChanged

    End Sub

    Private Sub rdbTTDescription_MouseHover(sender As Object, e As EventArgs) Handles rdbTTDescription.MouseHover
        ToolTipStand.Show("ToolTip in Projekt-Struktur zeigt die Projekt-Beschreibung", rdbTTDescription, 2000)
    End Sub

    Private Sub rdbTTDependencies_CheckedChanged(sender As Object, e As EventArgs) Handles rdbTTDependencies.CheckedChanged
        ToolTipStand.Show("ToolTip in Projekt-Struktur zeigt die Projekt-Abhängigkeiten", rdbTTDependencies, 2000)
    End Sub
End Class