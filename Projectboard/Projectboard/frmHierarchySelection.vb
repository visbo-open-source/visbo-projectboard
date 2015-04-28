Imports ProjectBoardDefinitions
Public Class frmHierarchySelection

    Private hry As clsHierarchy

    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection

    Friend menuOption As Integer

    Private Sub labelPPTVorlage_Click(sender As Object, e As EventArgs) Handles labelPPTVorlage.Click

    End Sub

    Private Sub frmHierarchySelection_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        hry = New clsHierarchy

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            Call addToSuperHierarchy(hry, kvp.Value)
        Next

        Call buildHryTreeView()

        ' die Vorlagen einlesen
        Call frmHryNameReadPPTVorlagen(Me.menuOption, repVorlagenDropbox)

    End Sub

    

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim anzahlKnoten As Integer
        Dim selectedNode As TreeNode
        Dim tmpNode As TreeNode

        Dim element As String


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


        Dim a As Integer = 1


    End Sub

    Private Sub einstellungen_Click(sender As Object, e As EventArgs) Handles einstellungen.Click

        Dim mppFrm As New frmMppSettings
        Dim dialogreturn As DialogResult

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


                    If elemIDIstMeilenstein(childNameID) Then
                        childNode.BackColor = System.Drawing.Color.Azure
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

                    If elemIDIstMeilenstein(childNameID) Then
                        nodeLevel0.BackColor = System.Drawing.Color.Azure

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
        Call MsgBox("Doppel-Klick")
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

    Private Sub hryTreeView_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles hryTreeView.MouseDoubleClick
        Call MsgBox("Mouse Doppel-Klick")
    End Sub
End Class