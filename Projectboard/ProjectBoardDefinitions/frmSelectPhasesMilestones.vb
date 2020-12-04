Imports System.Windows.Forms
Public Class frmSelectPhasesMilestones

    Private hry As clsHierarchy

    'Private allMilestones As New Collection
    'Private allPhases As New Collection

    Public selectedMilestones As New Collection
    Public selectedPhases As New Collection

    ' steuert ob die showrangeLEft und showrangeRight Daten gezeigt werden 
    Public Property addElementMode As Boolean


    Private dontFire As Boolean = False
    Dim hryStufenValue As Integer = 50

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        _addElementMode = False
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub

    Private Sub frmSelectPhasesMilestones_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        If frmCoord(PTfrm.listselP, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.listselP, PTpinfo.top))
            Me.Left = CInt(frmCoord(PTfrm.listselP, PTpinfo.left))
        Else
            Me.Top = 60
            Me.Left = 100
        End If

        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            vonDate.MinDate = getDateofColumn(showRangeLeft, False).AddMonths(-3)
            vonDate.MaxDate = getDateofColumn(showRangeRight, True).AddMonths(-1)
            vonDate.Value = getDateofColumn(showRangeLeft, False)

            bisDate.MinDate = getDateofColumn(showRangeLeft, False).AddMonths(1)
            bisDate.MaxDate = getDateofColumn(showRangeRight, True).AddMonths(24)
            bisDate.Value = getDateofColumn(showRangeRight, True)
        End If


        ' Button Visibility und Texte definieren 
        Call defineFrmButtonVisibility()

        ' hier soll immer mit leeren Selektionen begonnen werden
        selectedMilestones.Clear()
        selectedPhases.Clear()

        Call buildHryTreeViewInPPT(PTItemType.projekt)



    End Sub


    ''' <summary>
    ''' bestimmt die Benennung der Buttons und des Formulars in Abhängigkeit von dt / engl.
    ''' </summary>
    Private Sub defineFrmButtonVisibility()

        If AlleProjekte.Count > 1 Then
            rdbProjStruktProj.Visible = True
            rdbProjStruktTyp.Visible = True
        Else
            rdbProjStruktProj.Visible = False
            rdbProjStruktTyp.Visible = False
        End If

        zeitLabel.Visible = Not addElementMode
        vonDate.Visible = Not addElementMode
        bisDate.Visible = Not addElementMode

        If awinSettings.englishLanguage Then
            zeitLabel.Text = "Timeframe"
            einstellungen.Text = "Settings"
            Me.Text = "Selection of projects, phases, milestones"
            Me.Ok_Button.Text = "Confirm selection"
        Else
            zeitLabel.Text = "Zeitraum"
            einstellungen.Text = "Einstellungen"
            Me.Text = "Auswahl von Projekten, Phasen, Meilensteinen"
            Me.Ok_Button.Text = "Auswahl bestätigen"
        End If

    End Sub

    ''' <summary>
    ''' baut den TreeView aus Projekte, Phasen udn Meilensteinen auf 
    ''' </summary>
    ''' <param name="auswahl"></param>
    Private Sub buildHryTreeViewInPPT(ByVal auswahl As PTItemType)

        Dim topLevel As TreeNode
        Dim kennung As String ' V: für Vorlagen, P: für Projekte, C: für Kategorien/Darstellungsklassen
        Dim hry As clsHierarchy
        Dim checkProj As Boolean = False
        ' das kann später verwendet werden, um auf Basis aller geladenen Projekte die verschiedenen Vorlagen anzuzeigen ...
        'Dim projekteToLook As clsProjekte = ShowProjekte

        With TreeViewProjects
            .Nodes.Clear()
            .CheckBoxes = True

            ' aktuell wird nur auf Liste de rProjekte abgestellt ... 
            ' alle Projekte zeigen 

            If auswahl = PTItemType.vorlage Then

                ' das Projekt als Vorlage zeigen, das zuvor in der Projekt-Ansicht gezeigt wurde ... 

                kennung = "V:"


                For Each kvp As KeyValuePair(Of String, clsProjekt) In selectedProjekte.Liste

                    If kvp.Value.hierarchy.count > 0 Then
                        topLevel = .Nodes.Add(kvp.Key)
                        topLevel.Name = kennung & kvp.Key
                        topLevel.Text = kvp.Key

                        hry = kvp.Value.hierarchy

                        Dim nodeToCheck As Boolean = False

                        If selectedPhases.Count > 0 Then
                            nodeToCheck = kvp.Value.containsAnyPhasesOfCollection(selectedPhases)
                        Else
                            nodeToCheck = False
                        End If

                        If selectedMilestones.Count > 0 Then
                            nodeToCheck = nodeToCheck Or kvp.Value.containsAnyMilestonesOfCollection(selectedMilestones)
                        Else
                            nodeToCheck = nodeToCheck Or False
                        End If

                        If nodeToCheck Then
                            topLevel.Checked = True
                        End If

                        Call buildProjectSubTreeViewInPPT(topLevel, hry)
                    End If


                Next

            Else
                kennung = "P:"
                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    If kvp.Value.hierarchy.count > 0 Then
                        topLevel = .Nodes.Add(kvp.Key)
                        topLevel.Name = kennung & kvp.Key
                        topLevel.Text = kvp.Key
                        hry = kvp.Value.hierarchy

                        If selectedPhases.Count > 0 Or selectedMilestones.Count > 0 Then
                            ' überprüfen, ob das Projekt irgend eine der selektierten Phasen oder Meilensteine enthält

                            Dim tmpcollection As New Collection
                            Dim newFil As New clsFilter("tmp", tmpcollection, tmpcollection,
                                                        selectedPhases, selectedMilestones, tmpcollection, tmpcollection)
                            If newFil.doesNotBlock(kvp.Value) Then
                                topLevel.Checked = True
                            End If
                        End If

                        Call buildProjectSubTreeViewInPPT(topLevel, hry)
                    End If

                Next
            End If


            ' tk 13.10.19 in Projectboard : in buildHryTreeView kann in Abhängigkeit von auswahl gewählt werden 
            ' das fehlt hier ... 

        End With

    End Sub

    ''' <summary>
    ''' baut dei Subtree-Struktur auf 
    ''' </summary>
    ''' <param name="topNode"></param>
    ''' <param name="hry"></param>
    Private Sub buildProjectSubTreeViewInPPT(ByRef topNode As TreeNode, ByVal hry As clsHierarchy)

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

                    Dim categoryElem As String = ""

                    childNameID = hryNode.getChild(i)
                    nodeLevel0 = .Nodes.Add(elemNameOfElemID(childNameID))
                    nodeLevel0.Name = childNameID


                    Dim isMilestone As Boolean = elemIDIstMeilenstein(childNameID)
                    Dim cMilestone As clsMeilenstein = Nothing
                    Dim cPhase As clsPhase = Nothing

                    Dim tmpBreadcrumb As String = hry.getBreadCrumb(childNameID, CInt(hryStufenValue))
                    Dim elemName As String = elemNameOfElemID(childNameID)
                    Dim element As String = calcHryFullname(elemName, tmpBreadcrumb)
                    Dim projElem As String = "[" & topNode.Name & "]" & element


                    ' tk, 3.12.17 wird doch gar nicht verwendet ..?
                    'If Projektvorlagen.Contains(topNode.Text) Then
                    '    Dim vproj As clsProjektvorlage = Projektvorlagen.getProject(topNode.Text)
                    'End If

                    If ShowProjekte.contains(topNode.Text) Then

                        Dim hproj As clsProjekt = ShowProjekte.getProject(topNode.Text)
                        vorlElem = "[V:" & hproj.VorlagenName & "]" & element

                        If isMilestone Then
                            cMilestone = hproj.getMilestoneByID(childNameID)
                            ' bool'sche Wert gibtz an, ob es sich um einen Meilenstein handelt 
                            categoryElem = calcHryCategoryName(cMilestone.appearance, True)
                        Else
                            cPhase = hproj.getPhaseByID(childNameID)
                            ' bool'sche Wert gibt an, ob es sich um einen Meilenstein handelt
                            categoryElem = calcHryCategoryName(cPhase.appearance, False)
                        End If
                    End If

                    If elemIDIstMeilenstein(childNameID) Then
                        nodeLevel0.BackColor = System.Drawing.Color.Azure
                        If selectedMilestones.Contains(element) Or selectedMilestones.Contains(projElem) _
                            Or selectedMilestones.Contains(vorlElem) Or selectedMilestones.Contains(elemName) Or
                            selectedMilestones.Contains(categoryElem) Then
                            nodeLevel0.Checked = True
                        End If
                    Else

                        If selectedPhases.Contains(element) Or selectedPhases.Contains(projElem) _
                            Or selectedPhases.Contains(vorlElem) Or selectedPhases.Contains(elemName) Or
                            selectedPhases.Contains(categoryElem) Then
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

    Private Sub TreeViewProjects_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeViewProjects.BeforeExpand
        Dim node As TreeNode
        'Dim parentNode As TreeNode = Nothing
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

        Dim hryStufenValue As Integer = 50


        'Dim childRole As clsRollenDefinition

        node = e.Node
        elemID = node.Name


        ' es kann sich hier um die PRojekt- und die Vorlagen Struktur handeln, diese Struktur soll hier exoandiert werden 
        If type = PTItemType.vorlage Then
            curHry = selectedProjekte.getProject(1).hierarchy
        Else
            curHry = ShowProjekte.getProject(getPnameFromKey(PVname)).hierarchy
        End If


        If Not IsNothing(node.Tag) Then

            ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
            If node.Tag = "P" Then

                node.Tag = "X"

                ' Löschen von Platzhalter
                node.Nodes.Clear()

                hryNode = curHry.nodeItem(elemID)

                anzChilds = hryNode.childCount

                With TreeViewProjects
                    .CheckBoxes = True

                    For i As Integer = 1 To anzChilds

                        childNameID = hryNode.getChild(i)
                        childNode = node.Nodes.Add(elemNameOfElemID(childNameID))
                        childNode.Name = childNameID


                        Dim tmpBreadcrumb As String = curHry.getBreadCrumb(childNameID, CInt(hryStufenValue))
                        Dim elemName As String = elemNameOfElemID(childNameID)
                        Dim ele As String = calcHryFullname(elemName, tmpBreadcrumb)

                        ' gehe auf den root-Knoten
                        Dim topNode As TreeNode = node
                        Do While Not IsNothing(topNode.Parent)
                            topNode = topNode.Parent
                        Loop
                        Dim pvElem As String = "[" & topNode.Name & "]" & ele


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

    End Sub

    Private Sub TreeViewProjects_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles TreeViewProjects.AfterCheck

        Dim oNode As TreeNode
        Dim hnode As TreeNode
        Dim anzCheckedNodes As Integer = 0

        If Not dontFire Then
            oNode = e.Node


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

    Private Sub TreeViewProjects_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TreeViewProjects.KeyPress

        Dim initialNode As TreeNode = TreeViewProjects.SelectedNode
        Dim checkMode As Boolean

        dontFire = True
        Try
            If e.KeyChar = "a" Or e.KeyChar = "A" Then
                ' nur unmittelbare Kind-Knoten werden checked / unchecked 
                With TreeViewProjects.SelectedNode
                    '.Expand()
                    If .Nodes.Count > 0 Then
                        checkMode = Not .Nodes.Item(0).Checked
                        For i As Integer = 1 To .Nodes.Count
                            .Nodes.Item(i - 1).Checked = checkMode
                        Next
                    End If

                End With

            ElseIf e.KeyChar = "m" Or e.KeyChar = "M" Then
                ' selektiere/de-selektiere Meilensteine  
                With TreeViewProjects.SelectedNode
                    .Expand()
                    Dim ix As Integer = 1
                    Dim fertig As Boolean = False
                    While ix <= .Nodes.Count And Not fertig
                        If elemIDIstMeilenstein(.Nodes.Item(ix - 1).Name) Then
                            checkMode = Not .Nodes.Item(ix - 1).Checked
                            For i As Integer = ix To .Nodes.Count
                                If elemIDIstMeilenstein(.Nodes.Item(i - 1).Name) Then
                                    .Nodes.Item(i - 1).Checked = checkMode
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
                With TreeViewProjects.SelectedNode
                    .Expand()
                    Dim ix As Integer = 1
                    Dim fertig As Boolean = False
                    While ix <= .Nodes.Count And Not fertig
                        If Not elemIDIstMeilenstein(.Nodes.Item(ix - 1).Name) Then
                            checkMode = Not .Nodes.Item(ix - 1).Checked
                            For i As Integer = ix To .Nodes.Count
                                If Not elemIDIstMeilenstein(.Nodes.Item(i - 1).Name) Then
                                    .Nodes.Item(i - 1).Checked = checkMode
                                End If
                            Next
                            fertig = True
                        Else
                            ix = ix + 1
                        End If
                    End While
                End With
            End If
        Catch ex As Exception
            dontFire = False
        End Try

        dontFire = False

        ' kennzeichnen, daß keine weitere Behandlung , insbesondere nicht die Standard-Behandlung notwendig ist 
        e.Handled = True

    End Sub

    Private Sub Ok_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click

        ' showRangeLeft und showrange Right bestimmen
        showRangeLeft = getColumnOfDate(vonDate.Value)
        showRangeRight = getColumnOfDate(bisDate.Value)


        Dim anzahlKnoten As Integer
        Dim selectedNode As TreeNode = Nothing
        Dim tmpNode As TreeNode

        Dim element As String
        Dim type As Integer = -1
        Dim pvName As String = ""

        selectedPhases.Clear()
        selectedMilestones.Clear()

        anzahlKnoten = TreeViewProjects.Nodes.Count

        With TreeViewProjects

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

                        Dim tmpBreadcrumb As String = hry.getBreadCrumb(rootPhaseName, CInt(hryStufenValue))
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
                    Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufenValue))
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

        MyBase.Close()

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
                        Dim tmpBreadcrumb As String = hry.getBreadCrumb(tmpNode.Name, CInt(hryStufenValue))
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

        If type = PTItemType.vorlage Then

            ' jetzt anders ... 
            If selectedProjekte.contains(getPnameFromKey(pvName)) Then
                tmpResult = selectedProjekte.getProject(getPnameFromKey(pvName)).hierarchy
            End If

            'If Projektvorlagen.Contains(pvName) Then
            '    tmpResult = Projektvorlagen.getProject(pvName).hierarchy
            'End If

        Else
            If ShowProjekte.contains(getPnameFromKey(pvName)) Then
                tmpResult = ShowProjekte.getProject(getPnameFromKey(pvName)).hierarchy
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

        If curNode.Name.StartsWith("P:") Or
            curNode.Name.StartsWith("V:") Then

            Dim tmpStr() As String = curNode.Name.Split(New Char() {CChar(":")}, 2)
            If tmpStr.Length >= 2 Then
                tmpResult = tmpStr(1)
            End If

        End If

        If AlleProjekte.Count > 0 Then
            Dim tmpList As Collection = AlleProjekte.getVariantNames(tmpResult, False)

            If tmpList.Count > 0 Then
                Dim variantName As String = CStr(tmpList.Item(1))
                tmpResult = calcProjektKey(tmpResult, variantName)
                Dim hproj As clsProjekt = AlleProjekte.getProject(tmpResult, variantName)
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
            tmpResult = PTItemType.vorlage
        ElseIf curNode.Name.StartsWith("P:") Then
            tmpResult = PTItemType.projekt
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

    ''' <summary>
    ''' setzt alle Knoten auf Selected
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SelectionSet_Click(sender As Object, e As EventArgs) Handles SelectionSet.Click


        Dim curNode As TreeNode
        With TreeViewProjects


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
    ''' setzt alle Knoten auf unselected 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub resetSelections_Click(sender As Object, e As EventArgs) Handles resetSelections.Click

        Dim curNode As TreeNode
        With TreeViewProjects


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
    ''' minimiert die TreeView Struktur, faltet sie zusammen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub collapseTree_Click(sender As Object, e As EventArgs) Handles collapseTree.Click

        With TreeViewProjects
            .CollapseAll()
        End With

    End Sub

    ''' <summary>
    ''' entfaltet den Baum vollständig
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub expandTree_Click(sender As Object, e As EventArgs) Handles expandTree.Click

        With TreeViewProjects
            .ExpandAll()
        End With
    End Sub

    Private Sub TreeViewProjects_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeViewProjects.AfterSelect

        ' nur etwas machen, wenn man im Modus Projekte ist, nur dann kann ja ein Projekt ausgewählt werden, das dann als Vorlage dienen soll ... 
        If rdbProjStruktProj.Checked = True And e.Node.Level = 0 Then
            Dim node As TreeNode = e.Node
            Dim elemID As String = node.Name
            Dim PVname As String = getPVnameFromNode(e.Node)

            If selectedProjekte.Count > 0 Then
                selectedProjekte.Clear(False)
            End If

            Dim pName As String = getPnameFromKey(PVname)
            If ShowProjekte.contains(pName) Then
                Dim hproj As clsProjekt = ShowProjekte.getProject(getPnameFromKey(PVname))
                selectedProjekte.Add(hproj, False)
            End If
        End If


    End Sub

    Private Sub rdbProjStruktTyp_CheckedChanged(sender As Object, e As EventArgs) Handles rdbProjStruktTyp.CheckedChanged

        If rdbProjStruktTyp.Checked Then
            Call buildHryTreeViewInPPT(PTItemType.vorlage)
        Else
            selectedPhases.Clear()
            selectedMilestones.Clear()
        End If

    End Sub

    Private Sub rdbProjStruktProj_CheckedChanged(sender As Object, e As EventArgs) Handles rdbProjStruktProj.CheckedChanged

        If rdbProjStruktProj.Checked Then
            Call buildHryTreeViewInPPT(PTItemType.projekt)
        Else
            selectedPhases.Clear()
            selectedMilestones.Clear()
        End If

    End Sub

    Private Sub vonDate_ValueChanged(sender As Object, e As EventArgs) Handles vonDate.ValueChanged
        If Not dontFire Then
            dontFire = True
            vonDate.Value = vonDate.Value.AddDays(-1 * vonDate.Value.Day + 1)
        Else
            dontFire = True
        End If
    End Sub

    Private Sub bisDate_ValueChanged(sender As Object, e As EventArgs) Handles bisDate.ValueChanged
        If Not dontFire Then
            dontFire = True
            bisDate.Value = bisDate.Value.AddDays(-1 * bisDate.Value.Day + 1).AddMonths(1).AddDays(-1)
        Else
            dontFire = False
        End If
    End Sub

    Private Sub einstellungen_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles einstellungen.LinkClicked
        Dim mppFrm As New frmMppSettings
        Dim dialogreturn As DialogResult

        mppFrm.calledfrom = "frmSelectPPTTempl"

        dialogreturn = mppFrm.ShowDialog

    End Sub

End Class