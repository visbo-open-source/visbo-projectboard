Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports System.ComponentModel

Public Class frmHierarchySelection

    Private hry As clsHierarchy

    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection
    Private selectedCosts As New Collection
    Private selectedRoles As New Collection
    Private selectedBUs As New Collection
    Private selectedTyps As New Collection

    Friend menuOption As Integer

    Private Sub labelPPTVorlage_Click(sender As Object, e As EventArgs) Handles labelPPTVorlage.Click

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

        Call retrieveSelections("Last", PTmenue.visualisieren, selectedBUs, selectedTyps, selectedPhases, selectedMilestones, selectedRoles, selectedCosts)

        Call buildHryTreeView()

        ' wenn es selektierte Phasen oder Meilensteine schon gibt, so wird die Hierarchie aufgeklappt angezeigt
        If selectedMilestones.Count > 0 Or selectedPhases.Count > 0 Then
            hryTreeView.ExpandAll()
        End If

        Cursor = Cursors.Default

        ' die Vorlagen einlesen
        Call frmHryNameReadPPTVorlagen(Me.menuOption, repVorlagenDropbox)

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


        ' jetzt wird der letzte Filter gespeichert ..
        Dim filterName As String = "Last"
        Call storeFilter(filterName, menuOption, selectedBUs, selectedTyps, _
                                                   selectedPhases, selectedMilestones, _
                                                   selectedRoles, selectedCosts, True)


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
        ElseIf showRangeRight - showRangeLeft > 5 Then
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
            Call frmHryNameActions(Me.menuOption, selectedPhases, selectedMilestones, _
                            selectedRoles, selectedCosts, Me.chkbxOneChart.Checked, filterName)
        End If

        appInstance.EnableEvents = True
        enableOnUpdate = True

        ' bei bestimmten Menu-Optionen das Formuzlar dann schliessen 
        If Me.menuOption = PTmenue.excelExport Or menuOption = PTmenue.filterdefinieren Then
            MyBase.Close()
        End If


    End Sub

    Private Sub einstellungen_Click(sender As Object, e As EventArgs) Handles einstellungen.Click

        Dim mppFrm As New frmMppSettings
        Dim dialogreturn As DialogResult

        mppFrm.calledfrom = "frmShowPlanElements"
        dialogreturn = mppFrm.ShowDialog

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
                    Dim element As String = calcHryFullname(elemName, tmpBreadcrumb)


                    If elemIDIstMeilenstein(childNameID) Then
                        childNode.BackColor = System.Drawing.Color.Azure
                        If selectedMilestones.Contains(element) Or selectedMilestones.Contains(elemName) Then
                            childNode.Checked = True
                        End If
                    Else
                        If selectedPhases.Contains(element) Or selectedPhases.Contains(elemName) Then
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

    Private Sub hryStufen_ValueChanged(sender As Object, e As EventArgs) Handles hryStufen.ValueChanged

    End Sub

    Private Sub hryTreeView_DoubleClick(sender As Object, e As EventArgs) Handles hryTreeView.DoubleClick

        

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
End Class