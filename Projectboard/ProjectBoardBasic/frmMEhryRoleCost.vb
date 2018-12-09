Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports ClassLibrary1
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Public Class frmMEhryRoleCost

    ' tk 9.12.18 enthält die Rollen, die beim Load des Formulars in der Projekt-Phase enthalten sind   
    Private initialRolesOfPhase As New SortedList(Of String, String)

    ' tk 9.12.18 enthält die Kosten, die beim Load des Formulars in der Projektphase enthalten sind 
    Private initialCostsOfPhase As New SortedList(Of String, String)

    ' tk 9.1218 enthält den identifier in Form roleUID;teamUID 
    Private currentRoleIDentifier As String

    ' das sind die Rollen, die am Ende, also wenn der ok Button gedrückt wird aufgesammelt werden  
    Private selectedRC As New Collection
    Public ergItems As New Collection

    ' der Projekt-Name in der Zeile 
    Public pName As String

    ' der Varianten-NAme in der Zeile
    Public vName As String

    ' der Phasen-Name in der Zeile 
    Public phaseName As String

    ' der Rollen-Kosten Name in der Zeile 
    Public rcName As String

    ' der ggf dazugehörende Team-Name 
    Public teamName As String

    ' die PhaseNameID der Zeile  
    Public phaseNameID As String

    ' das in der Zeile aktive Projekt
    Public hproj As clsProjekt

    Friend existingRoleFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75, System.Drawing.FontStyle.Regular)
    Friend normalRoleFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75, System.Drawing.FontStyle.Regular)
    Friend normalRoleColor As System.Drawing.Color = System.Drawing.Color.Black
    Friend existingRoleColor As System.Drawing.Color = System.Drawing.Color.DimGray


    Private Sub frmMEhryRoleCost_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If frmCoord(PTfrm.rolecostME, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.top))
            Me.Left = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.left))
        Else
            Me.Top = 60
            Me.Left = 100
        End If

        ' welche Rollen & Kosten sind in der aktuellen Phase drin ... 
        initialRolesOfPhase = hproj.getRoleIDs(phaseNameID)
        initialCostsOfPhase = hproj.getCostIDs(phaseNameID)

        ' wie lautet der Identifier der aktuellen Zeile: setzet sich zusammen aus roleuid;teamid
        Dim tmpRoleUID As Integer
        Dim isTeamMember As Boolean = False

        Dim tmpteamUid As Integer
        If RoleDefinitions.containsName(teamName) Then
            isTeamMember = True
            tmpteamUid = RoleDefinitions.getRoledef(teamName).UID
        End If

        If RoleDefinitions.containsName(rcName) Then
            tmpRoleUID = RoleDefinitions.getRoledef(rcName).UID
            currentRoleIDentifier = RoleDefinitions.bestimmeRoleNodeName(tmpRoleUID, isTeamMember, tmpteamUid)
        Else
            currentRoleIDentifier = ""
        End If

        Call buildMERoleTree()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim anzahlKnoten As Integer = hryRoleCost.Nodes.Count
        Dim tmpnode As TreeNode

        selectedRC.Clear()

        ' einsammeln der Rollen und Kosten die selektiert wurden
        With hryRoleCost

            For px As Integer = 1 To anzahlKnoten

                tmpnode = .Nodes.Item(px - 1)

                If tmpnode.Checked Then

                    If Not selectedRC.Contains(tmpnode.Text) Then
                        selectedRC.Add(tmpnode.Text, tmpnode.Text)
                    End If

                End If


                If tmpnode.Nodes.Count > 0 Then
                    Call pickupMECheckedRoleItems(tmpnode)
                End If

            Next

        End With

        Dim anzahlcheckedRoles As Integer = selectedRC.Count


        For i = 1 To anzahlcheckedRoles

            If Not IsNothing(hproj) Then

                If IsNothing(hproj.getPhaseByID(phaseNameID).getRole(selectedRC.Item(i))) _
                And IsNothing(hproj.getPhaseByID(phaseNameID).getCost(selectedRC.Item(i))) Then

                    ''Call massEditZeileEinfügen("")
                    ergItems.Add(selectedRC.Item(i))
                Else
                    If rcName = selectedRC.Item(i) Then
                        ergItems.Add(selectedRC.Item(i))
                    End If
                End If
            Else
                ergItems.Add(selectedRC.Item(i))
            End If


        Next

        DialogResult = System.Windows.Forms.DialogResult.OK
        MyBase.Close()

    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

    Private Sub frmMEhryRoleCost_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.rolecostME, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.rolecostME, PTpinfo.left) = Me.Left


    End Sub



    Public Sub buildMERoleTree()


        'Dim hPhase As clsPhase = Nothing
        'If Not IsNothing(hproj) Then
        '    hPhase = hproj.getPhaseByID(phaseNameID)
        'End If

        Dim topLevelNode As TreeNode
        Dim checkProj As Boolean = False

        With hryRoleCost

            .Nodes.Clear()
            .CheckBoxes = True


            ' alle Rollen zeigen 

            If RoleDefinitions.Count > 0 Then
                Dim topNodes As List(Of Integer) = RoleDefinitions.getTopLevelNodeIDs

                ' wenn die Sicht eingeschränkt werden soll ... 
                If Not IsNothing(awinSettings.isRestrictedToOrgUnit) Then
                    If awinSettings.isRestrictedToOrgUnit.Length > 0 Then

                        If RoleDefinitions.containsName(awinSettings.isRestrictedToOrgUnit) Then

                            topNodes.Clear()
                            topNodes.Add(RoleDefinitions.getRoledef(awinSettings.isRestrictedToOrgUnit).UID)

                        End If

                    End If
                End If


                For i = 0 To topNodes.Count - 1

                    Dim role As clsRollenDefinition = RoleDefinitions.getRoleDefByID(topNodes.ElementAt(i))

                    topLevelNode = .Nodes.Add(role.name)
                    topLevelNode.Text = role.name


                    Dim nrTag As New clsNodeRoleTag
                    With nrTag
                        If role.getSubRoleCount > 0 Then
                            .pTag = "P"
                            topLevelNode.Nodes.Clear()
                            topLevelNode.Nodes.Add("-")
                        Else
                            .pTag = "X"
                        End If
                    End With


                    ' tk 6.12.18 jetzt kommen ggf an einen Knoten noch diese Informationen

                    If role.isTeam Then
                        ' toplevelNode kann nur Team sein, nicht Team-Member
                        nrTag.isTeam = True
                        nrTag.isTeamMember = False
                    End If

                    topLevelNode.Tag = nrTag

                    topLevelNode.Name = RoleDefinitions.bestimmeRoleNodeName(role.UID, nrTag.isTeamMember, nrTag.membershipID)

                    ' ist die Rolle bereits in der Phase, die in der Zeile dargestellt wird ? 
                    If initialRolesOfPhase.ContainsKey(topLevelNode.Name) Then
                        topLevelNode.Checked = True
                    End If

                    ' hier muss gecheckt werden, ob irgendwelche existierende Kind-Rollen unterhalb der aktuellen topNode sind 
                    ' Diese sollen dann als kursiv dargestellt werden, die aktuelle Rolle als gecheckt markiert sein

                    If RoleDefinitions.hasAnyChildParentRelationsship(initialRolesOfPhase, role.UID) Then

                        ' entsprechend kennzeichnen 
                        topLevelNode.NodeFont = existingRoleFont
                        topLevelNode.ForeColor = existingRoleColor

                    End If

                    ' 9.12.18 nicht mehr nötig, da jetzt selektiv, wie der User den BAum entfaltet, aufgebaut wird 
                    'Dim listOfChildIDs As New SortedList(Of Integer, Double)
                    'Try
                    '    listOfChildIDs = role.getSubRoleIDs
                    'Catch ex As Exception

                    'End Try

                    'If listOfChildIDs.Count > 0 Then
                    '    For ii As Integer = 0 To listOfChildIDs.Count - 1
                    '        Call buildMESubRoleTree(topLevelNode, listOfChildIDs.ElementAt(ii).Key)
                    '    Next
                    'End If

                Next
            End If

            If CostDefinitions.Count > 1 Then

                For i = 1 To CostDefinitions.Count - 1
                    Dim cost As clsKostenartDefinition = CostDefinitions.getCostdef(i)

                    topLevelNode = .Nodes.Add(cost.name)
                    topLevelNode.Text = cost.name
                    topLevelNode.Name = cost.name
                    '
                    ' 9.12.18 neuer Stuff 
                    '
                    Dim nrTag As New clsNodeRoleTag
                    With nrTag
                        .pTag = "X"
                        .isRole = False
                    End With

                    topLevelNode.Tag = nrTag


                    ' ist die Rolle bereits in der Phase, die in der Zeile dargestellt wird ? 
                    If initialCostsOfPhase.ContainsKey(cost.name) Then
                        topLevelNode.Checked = True
                    End If


                    ' 9.12.18 alter Stuff
                    'topLevelNode.Text = cost.name

                    'If Not IsNothing(hPhase) Then
                    '    If Not IsNothing(hPhase.getCost(cost.name)) Then

                    '        ' entsprechend kennzeichnen 
                    '        topLevelNode.NodeFont = existingRoleFont
                    '        topLevelNode.ForeColor = existingRoleColor

                    '        If cost.name = rcName Then
                    '            topLevelNode.Checked = True
                    '        End If

                    '    End If
                    'End If


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
    ''' <param name="currentRoleUid"></param>
    ''' <remarks></remarks>
    Public Sub buildMESubRoleTree(ByRef parentNode As TreeNode, ByVal currentRoleUid As Integer)


        Dim currentRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(currentRoleUid)
        Dim childIds As SortedList(Of Integer, Double) = currentRole.getSubRoleIDs

        Dim currentNode As TreeNode
        Dim childNode As TreeNode = Nothing

        currentNode = parentNode.Nodes.Add(currentRole.name)
        currentNode.Text = currentRole.name


        Dim nrTag As New clsNodeRoleTag
        If currentRole.isTeam Then

            nrTag = New clsNodeRoleTag
            With nrTag
                .isTeam = True
                .isTeamMember = False
            End With

        ElseIf currentRole.getTeamIDs.Count > 0 And CType(parentNode.Tag, clsNodeRoleTag).isTeam Then

            nrTag = New clsNodeRoleTag
            With nrTag
                .isTeam = False
                .isTeamMember = True
                .membershipID = CInt(parentNode.Name)
                .membershipPrz = RoleDefinitions.getMembershipPrz(CInt(parentNode.Name), currentRoleUid)
            End With
        End If


        If childIds.Count > 0 Then
            currentNode.Nodes.Clear()
            currentNode.Nodes.Add("-")
            nrTag.pTag = "P"
        Else
            nrTag.pTag = "X"
        End If

        currentNode.Tag = nrTag

        currentNode.Name = RoleDefinitions.bestimmeRoleNodeName(currentRoleUid, nrTag.isTeamMember, nrTag.membershipID)

        ' ist die Rolle bereits in der Phase, die in der Zeile dargestellt wird ? 
        If initialRolesOfPhase.ContainsKey(currentNode.Name) Then
            currentNode.Checked = True
        End If

        ' hier muss gecheckt werden, ob irgendwelche existierende Kind-Rollen unterhalb der aktuellen topNode sind 
        ' Diese sollen dann als kursiv dargestellt werden, die aktuelle Rolle als gecheckt markiert sein

        If RoleDefinitions.hasAnyChildParentRelationsship(initialRolesOfPhase, currentRoleUid) Then

            ' entsprechend kennzeichnen 
            currentNode.NodeFont = existingRoleFont
            currentNode.ForeColor = existingRoleColor

        End If


        'Dim hPhase As clsPhase = Nothing
        'If Not IsNothing(hproj) Then
        '    hPhase = hproj.getPhaseByID(phaseNameID)
        'End If


        '' ---- altes Vorgehen 9.12.18 
        'Dim doItAnyWay As Boolean = False

        'With parentNode

        '    currentNode = .Nodes.Add(currentRole.name)
        '    currentNode.Name = currentRoleUid.ToString
        '    currentNode.Text = currentRole.name

        '    ' hier muss gecheckt werden, welche Rollen in dem Projekt und dieser Phase, in der der Doppelclick erfolgte
        '    ' vergeben sind. Diese sollen dann als kursiv dargestellt werden, die aktuelle Rolle als gecheckt markiert sein

        '    If Not IsNothing(hPhase) Then
        '        If Not IsNothing(hPhase.getRole(currentRole.name)) Then

        '            ' entsprechend kennzeichnen
        '            currentNode.NodeFont = existingRoleFont
        '            currentNode.ForeColor = existingRoleColor

        '            If currentRole.name = rcName Then
        '                currentNode.Checked = True
        '            End If

        '        End If
        '    End If


        'End With

        'For i = 0 To childIds.Count - 1

        '    Call buildMESubRoleTree(currentNode, childIds.ElementAt(i).Key)

        'Next
        ''End If

    End Sub

    ''' <summary>
    ''' gibt alle Namen von Knoten, die "gecheckt" sind, in der selectedRoles-Liste zurück  
    ''' wird rekursiv aufgerufen 
    ''' Achtung: wenn es Endlos Zyklen gibt, dann ist hier eine Endlos-Schleife ! 
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Public Sub pickupMECheckedRoleItems(ByVal node As TreeNode)
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
                        If Not selectedRC.Contains(element) Then
                            selectedRC.Add(element, element)
                        End If


                    End If


                    If tmpNode.Nodes.Count > 0 Then
                        Call pickupMECheckedRoleItems(tmpNode)
                    End If

                Next

            End With

        End If
    End Sub

    Private Sub hryRoleCost_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles hryRoleCost.BeforeExpand

        Dim node As TreeNode
        Dim anzChilds As Integer
        Dim elemID As String

        node = e.Node
        elemID = node.Name

        ' Rollen expandieren

        If Not IsNothing(node.Tag) Then

            'parentNode = node.Parent

            Dim nrTag As clsNodeRoleTag = CType(node.Tag, clsNodeRoleTag)
            ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
            If nrTag.pTag = "P" Then

                nrTag.pTag = "X"

                ' Löschen von Platzhalter
                node.Nodes.Clear()

                Dim nodelist As New SortedList(Of Integer, Double)
                Try
                    Dim teamID As Integer
                    nodelist = RoleDefinitions.getRoleDefByIDKennung(CInt(node.Name), teamID).getSubRoleIDs
                    anzChilds = nodelist.Count
                Catch ex As Exception
                    anzChilds = 0
                End Try


                With hryRoleCost
                    .CheckBoxes = True
                End With

                For i As Integer = 0 To anzChilds - 1

                    Call buildMESubRoleTree(node, nodelist.ElementAt(i).Key)

                Next

            End If

        End If

    End Sub
End Class