Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports System.Windows.Forms
Imports System.ComponentModel

Public Class frmHierarchySelection

    Private hry As clsHierarchy
    Public repProfil As clsReport

    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection
    Private selectedCosts As New Collection
    Private selectedRoles As New Collection
    Private selectedBUs As New Collection
    Private selectedTyps As New Collection

    Friend menuOption As Integer



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

        hry = New clsHierarchy

        If menuOption = PTmenue.filterdefinieren Then
            For Each kvp As KeyValuePair(Of String, clsProjektvorlage) In Projektvorlagen.Liste
                Dim hproj As New clsProjekt
                kvp.Value.copyAttrTo(hproj)
                kvp.Value.copyTo(hproj)
                Call addToSuperHierarchy(hry, hproj)
            Next
        Else
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                Call addToSuperHierarchy(hry, kvp.Value)
            Next
        End If

        If Not Me.menuOption = PTmenue.reportBHTC Then
            Call retrieveSelections("Last", PTmenue.visualisieren, selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)
        Else

            Call retrieveProfilSelection(filterDropbox.Text, PTmenue.reportBHTC, selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts, repProfil)
            If IsNothing(repProfil) Then
                Throw New ArgumentException("Fehler beim Lesen des áusgewählten ReportProfils")
            End If
        End If


        Call buildHryTreeView()

        ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
        If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
            hryTreeView.ExpandAll()
        End If

        Cursor = Cursors.Default

        ' die Vorlagen  einlesen
        Call frmHryNameReadPPTVorlagen(Me.menuOption, repVorlagenDropbox)

        ' die Filter einlesen

        If Not Me.menuOption = PTmenue.reportBHTC Then
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
            If Not IsNothing(repProfil.PPTTemplate) Then
                If My.Computer.FileSystem.FileExists(awinPath & RepProjectVorOrdner & "\" & repProfil.PPTTemplate) Then
                    repVorlagenDropbox.Text = repProfil.PPTTemplate
                Else
                    repVorlagenDropbox.Text = ""
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

                If tmpNode.Checked Then
                    ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll 

                    Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
                    Dim elemName As String = elemNameOfElemID(tmpNode.Name)
                    element = calcHryFullname(elemName, tmpBreadcrumb)

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
                    Call pickupCheckedItems(tmpNode)
                End If

            Next

        End With

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
        If Not Me.menuOption = PTmenue.reportBHTC Then
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
            Me.menuOption = PTmenue.reportBHTC Then
            validOption = True
        ElseIf showRangeRight - showRangeLeft > 5 Then
            validOption = True
        Else
            validOption = False
        End If

        If Me.menuOption = PTmenue.multiprojektReport Or Me.menuOption = PTmenue.einzelprojektReport Or _
            Me.menuOption = PTmenue.reportBHTC Then

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
                        AbbrButton.Text = "Abbrechen"

                        ' Alternativ ohne Background Worker
                        If Me.menuOption = PTmenue.reportBHTC Then


                            Call MsgBox("Report erstellen mit Projekt " & repProfil.VonDate.ToString & " bis " & repProfil.BisDate.ToString & " Reportprofil " & repProfil.name)
                            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

                            repProfil.PPTTemplate = repVorlagenDropbox.Text

                            BackgroundWorker3.RunWorkerAsync(repProfil)


                        Else
                            BackgroundWorker1.RunWorkerAsync(vorlagenDateiName)
                        End If


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
            ' hier können nur Phasen / Meilensteine ausgewählt werden; 
            Dim tmpCollection As New Collection
            Call frmHryNameActions(Me.menuOption, selectedPhases, selectedMilestones, _
                            tmpCollection, tmpCollection, Me.chkbxOneChart.Checked, lastfilter)
        End If

        appInstance.EnableEvents = True
        enableOnUpdate = True

        ' bei bestimmten Menu-Optionen das Formuzlar dann schliessen 
        If Me.menuOption = PTmenue.excelExport Or menuOption = PTmenue.filterdefinieren Or Me.menuOption = PTmenue.reportBHTC Then
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

        mppFrm.calledfrom = "frmShowPlanElements"

        If Me.menuOption = PTmenue.reportBHTC Then
            With awinSettings

                .drawProjectLine = True
                .eppExtendedMode = repProfil.ExtendedMode
                .mppExtendedMode = repProfil.ExtendedMode
                .mppOnePage = repProfil.OnePage
                .mppShowAllIfOne = repProfil.AllIfOne
                .mppShowAmpel = repProfil.Ampeln
                .mppShowLegend = repProfil.Legend
                .mppShowMsDate = repProfil.MSDate
                .mppShowMsName = repProfil.MSName
                .mppShowPhDate = repProfil.PhDate
                .mppShowPhName = repProfil.PhName
                .mppShowProjectLine = repProfil.ProjectLine
                .mppSortiertDauer = repProfil.SortedDauer
                .mppVertikalesRaster = repProfil.VLinien
                .mppFullyContained = repProfil.FullyContained
                ' .mppShowHorizontals = repProfil.ShowHorizontals
                ' .mppUseAbbreviation = repProfil.UseAbbreviation
                ' .mppUseOriginalNames = repProfil.UseOriginalNames
            End With
        End If


        dialogreturn = mppFrm.ShowDialog


        If Me.menuOption = PTmenue.reportBHTC Then

            With awinSettings

                .drawProjectLine = True
                repProfil.ExtendedMode = .eppExtendedMode
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
                'repProfil.ShowHorizontals = .mppShowHorizontals
                'repProfil.UseAbbreviation = .mppUseAbbreviation
                'repProfil.UseOriginalNames = .mppUseOriginalNames
            End With
        End If

    End Sub


    Private Sub hryTreeView_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles hryTreeView.BeforeExpand

        Dim node As TreeNode
        Dim childNode As TreeNode
        Dim placeholder As TreeNode
        Dim elemID As String
        Dim hryNode As clsHierarchyNode
        Dim anzChilds As Integer
        Dim childNameID As String

        node = e.Node
        elemID = node.Name


        ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
        If node.Tag = "P" Then

            node.Tag = "X"

            ' Löschen von Platzhalter
            node.Nodes.Clear()

            hryNode = hry.nodeItem(elemID)

            anzChilds = hryNode.childCount

            With hryTreeView
                .CheckBoxes = True

                For i As Integer = 1 To anzChilds

                    childNameID = hryNode.getChild(i)
                    childNode = node.Nodes.Add(elemNameOfElemID(childNameID))
                    childNode.Name = childNameID


                    Dim tmpBreadcrumb As String = hry.getBreadCrumb(childNameID, CInt(hryStufen.Value))
                    Dim elemName As String = elemNameOfElemID(childNameID)
                    Dim ele As String = calcHryFullname(elemName, tmpBreadcrumb)


                    If elemIDIstMeilenstein(childNameID) Then
                        childNode.BackColor = System.Drawing.Color.Azure
                        If selectedMilestones.Contains(ele) Or selectedMilestones.Contains(elemName) Then
                            childNode.Checked = True
                        End If
                    Else
                        If selectedPhases.Contains(ele) Or selectedPhases.Contains(elemName) Then
                            childNode.Checked = True
                        End If
                    End If



                    If hry.nodeItem(childNameID).childCount > 0 Then
                        childNode.Tag = "P"
                        placeholder = childNode.Nodes.Add("-")
                        placeholder.Tag = "P"
                    Else
                        childNode.Tag = "X"
                    End If


                Next

            End With


        End If

    End Sub

    ''' <summary>
    ''' baut den TreeView für die Hierarchie auf , Treeview enthält sowohl Meilensteine als auch Phasen
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub buildHryTreeView()

        Dim hryNode As clsHierarchyNode
        Dim anzChilds As Integer
        Dim childNameID As String
        Dim nodeLevel0 As TreeNode
        Dim nodeLevel1 As TreeNode

        With hryTreeView
            .Nodes.Clear()
        End With

        If hry.count >= 1 Then
            hryNode = hry.nodeItem(rootPhaseName)

            anzChilds = hryNode.childCount

            With hryTreeView
                .CheckBoxes = True

                For i As Integer = 1 To anzChilds

                    childNameID = hryNode.getChild(i)
                    nodeLevel0 = .Nodes.Add(elemNameOfElemID(childNameID))
                    nodeLevel0.Name = childNameID

                    Dim tmpBreadcrumb As String = hry.getBreadCrumb(childNameID, CInt(hryStufen.Value))
                    Dim elemName As String = elemNameOfElemID(childNameID)
                    Dim element As String = calcHryFullname(elemName, tmpBreadcrumb)


                    If elemIDIstMeilenstein(childNameID) Then
                        nodeLevel0.BackColor = System.Drawing.Color.Azure
                        If selectedMilestones.Contains(element) Or selectedMilestones.Contains(elemName) Then
                            nodeLevel0.Checked = True
                        End If
                    Else

                        If selectedPhases.Contains(element) Or selectedPhases.Contains(elemName) Then
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
            Call MsgBox("es ist keine Hierarchie gegeben")
        End If
    End Sub



    ''' <summary>
    ''' gibt alle Namen von Knoten, die "gecheckt" sind, in der nameList zurück  
    ''' wird rekursiv aufgerufen 
    ''' Achtung: wenn es Endlos Zyklen gibt, dann ist hier eine Endlos-Schleife ! 
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Private Sub pickupCheckedItems(ByVal node As TreeNode)

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

                        Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
                        Dim elemName As String = elemNameOfElemID(tmpNode.Name)
                        element = calcHryFullname(elemName, tmpBreadcrumb)

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
                        Call pickupCheckedItems(tmpNode)
                    End If

                Next

            End With

        End If

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

        Try
            With awinSettings

                If vorlagenDateiName.Contains(RepPortfolioVorOrdner) Then
                    Call createPPTSlidesFromConstellation(vorlagenDateiName, _
                                                      selectedPhases, selectedMilestones, _
                                                      selectedRoles, selectedCosts, _
                                                      selectedBUs, selectedTyps, True, _
                                                      worker, e)
                Else
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

        If Me.menuOption = PTmenue.filterdefinieren Then

            Dim fName As String = filterDropbox.SelectedItem.ToString
            ' wird nicht benötigt: ur: 29.07.2015 Dim filter As clsFilter = filterDefinitions.retrieveFilter(fName)

            ' jetzt werden anhand des Filters "fName" die Collections gesetzt 
            Call retrieveSelections(fName, menuOption, selectedBUs, selectedTyps, _
                                    selectedPhases, selectedMilestones, _
                                    selectedRoles, selectedCosts)

            Call buildHryTreeView()

            ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
            If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                hryTreeView.ExpandAll()
            End If

            Cursor = Cursors.Default

        ElseIf Me.menuOption = PTmenue.reportBHTC Then

            'neuer Profil-Name in Klasse repProfil speichern
            repProfil.name = filterDropbox.SelectedItem.ToString



        Else


            Dim fName As String = filterDropbox.SelectedItem.ToString
            ' wird nicht benötigt: ur: 29.07.2015 Dim filter As clsFilter = filterDefinitions.retrieveFilter(fName)

            ' jetzt werden anhand des Filters "fName" die Collections gesetzt 
            Call retrieveSelections(fName, menuOption, selectedBUs, selectedTyps, _
                                    selectedPhases, selectedMilestones, _
                                    selectedRoles, selectedCosts)

            Call buildHryTreeView()

            ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
            If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
                hryTreeView.ExpandAll()
            End If

            Cursor = Cursors.Default

        End If

    End Sub

    Private Sub auswSpeichern_Click(sender As Object, e As EventArgs) Handles auswSpeichern.Click

        Dim anzahlKnoten As Integer
        Dim selectedNode As TreeNode
        Dim tmpNode As TreeNode
        Dim filterName As String = ""
        Dim element As String

        If Not Me.menuOption = PTmenue.reportBHTC Then


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

                    If tmpNode.Checked Then
                        ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll 

                        Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
                        Dim elemName As String = elemNameOfElemID(tmpNode.Name)
                        element = calcHryFullname(elemName, tmpBreadcrumb)

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
                        Call pickupCheckedItems(tmpNode)
                    End If

                Next

            End With

            If Me.menuOption = PTmenue.filterdefinieren Then

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


        ElseIf Me.menuOption = PTmenue.reportBHTC Then


            statusLabel.Text = ""


            anzahlKnoten = hryTreeView.Nodes.Count
            selectedNode = hryTreeView.SelectedNode

            selectedPhases.Clear()
            selectedMilestones.Clear()

            With hryTreeView

                For px As Integer = 1 To anzahlKnoten

                    tmpNode = .Nodes.Item(px - 1)

                    If tmpNode.Checked Then
                        ' nur dann muss ja geprüft werden, ob das Element aufgenommen werden soll 

                        Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufen.Value))
                        Dim elemName As String = elemNameOfElemID(tmpNode.Name)
                        element = calcHryFullname(elemName, tmpBreadcrumb)

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
                        Call pickupCheckedItems(tmpNode)
                    End If

                Next

            End With


            Dim vorlagenDateiName As String

            vorlagenDateiName = awinPath & RepProjectVorOrdner & _
                                "\" & repVorlagenDropbox.Text

            ' Prüfen, ob die Datei überhaupt existirt 
            If repVorlagenDropbox.Text.Length = 0 Then
                Call MsgBox("bitte PPT Vorlage auswählen !")
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
                    Call MsgBox("Bitte geben Sie einen Namen für diese Report-Profil an")
                End If



            Else

                Call MsgBox("bitte PPT Vorlage auswählen !")

            End If


        Else
            Call MsgBox("nicht reportBHTC aber auch reportBHTC: also eigentlich nicht möglich")
        End If



    End Sub


    Private Sub BackgroundWorker3_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        ' ''Dim vorlagenDateiName As String = CType(e.Argument, String)

        ' ReportProfil ist nun in reportProfil komplett enthalten
        Dim reportProfil As clsReport = CType(e.Argument, clsReport)

        Dim zeilenhoehe As Double = 0.0     ' zeilenhöhe muss für alle Projekte gleich sein, daher mit übergeben
        Dim legendFontSize As Single = 0.0  ' FontSize der Legenden der Schriftgröße des Projektnamens angepasst

        Dim selectedPhases As New Collection
        Dim selectedMilestones As New Collection
        Dim selectedRoles As New Collection
        Dim selectedCosts As New Collection
        Dim selectedBUs As New Collection
        Dim selectedTypes As New Collection

        Call msgbox("Backgroundworker3_DoWork")

        selectedPhases = copySortedListtoColl(reportProfil.Phases)
        selectedMilestones = copySortedListtoColl(reportProfil.Milestones)
        selectedRoles = copySortedListtoColl(reportProfil.Roles)
        selectedCosts = copySortedListtoColl(reportProfil.Costs)
        selectedBUs = copySortedListtoColl(reportProfil.BUs)
        selectedTypes = copySortedListtoColl(reportProfil.Typs)

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


        ' Report wird von Projekt hproj, das vor Aufruf des Formulars in hproj gespeichert wurde erzeugt

        showRangeLeft = CInt(DateDiff(DateInterval.Month, StartofCalendar, reportProfil.VonDate))
        showRangeRight = CInt(DateDiff(DateInterval.Month, StartofCalendar, reportProfil.BisDate))

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

    Private Sub BackgroundWorker3_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker3.ProgressChanged


        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusLabel.Text = CType(re.Result, String)

    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted

        '' ''With Me.AbbrButton
        '' ''    .Text = ""
        '' ''    .Visible = False
        '' ''    .Enabled = False
        '' ''    .Left = .Left + 40
        '' ''End With


        Me.statusLabel.Text = "...done"
        Me.statusLabel.Visible = True
        Me.OKButton.Visible = True
        Me.OKButton.Enabled = True
        Me.OKButton.Enabled = True
        Me.Cursor = System.Windows.Forms.Cursors.Arrow

        ' hier evt. noch schließen und Abspeichern des Reports von PPT

    End Sub
End Class