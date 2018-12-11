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

    ' das sind die Rollen, die dazu gekommen sind, also noch nicht in der initialRolesOfPhase waren 
    Public rolesToAdd As New Collection

    ' das sind die Rollen, die weggefallen sind, also bereits in der InitialRolesOfPhase waren 
    Public rolesToDelete As New Collection

    ' das sind die Kosten, die dazu gekommen sind, also noch nicht in der initialCostsOfPhase waren
    Public costsToAdd As New Collection

    ' das sind die Kosten, die weggefallen sind, also also bereits in der InitialRolesOfPhase waren 
    Public costsToDelete As New Collection

    ' der Projekt-Name in der Zeile 
    Public pName As String

    ' der Varianten-NAme in der Zeile
    Public vName As String

    ' der Phasen-Name in der Zeile 
    Public phaseName As String

    ' der Rollen-Kosten Name in der Zeile 
    Public rcName As String

    ' die Rollen-ID in der form roleUid;teamID oder roleUid.tostring bzw. costuid.tostring 
    Public rcNameID As String

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
        ' der wird bereits beim Right Click ermittelt und steht in rcNameID - siehe oben ...


        Call buildMERoleTree()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim anzahlKnoten As Integer = hryRoleCost.Nodes.Count
        Dim tmpnode As TreeNode


        ' 1. bestimmen der Rollen und Kosten, die gecheckt sind ... -> in rolesToadd bzw. costsToAdd
        ' 2. alles aus initialRolesOfPhase, die nicht mehr in rolesToadd sind : in rolesTobeDeleted übernehmen ; same for costs
        ' 3. alle, die bereits in initialRolesOfPhase sind, aus rolesToAdd rausnehmen , same for costs 


        With hryRoleCost

            ' Schritt 1: bestimmen der Rollen und Kosten, die gecheckt sind ... -> in rolesToadd bzw. costsToAdd
            For px As Integer = 1 To anzahlKnoten

                tmpnode = .Nodes.Item(px - 1)

                If tmpnode.Checked Then

                    If CType(tmpnode.Tag, clsNodeRoleTag).isRole Then
                        If Not rolesToAdd.Contains(tmpnode.Name) Then
                            rolesToAdd.Add(tmpnode.Name, tmpnode.Name)
                        End If
                    Else
                        If Not costsToAdd.Contains(tmpnode.Name) Then
                            costsToAdd.Add(tmpnode.Name, tmpnode.Name)
                        End If
                    End If

                End If

                If tmpnode.Nodes.Count > 0 Then
                    Call pickupMECheckedRoleCostItems(tmpnode)
                End If

            Next

        End With

        ' Schritt 2: alles aus initialRolesOfPhase, die nicht mehr in rolesToadd sind : in rolesTobeDeleted übernehmen 
        ' Schritt 2 - Rollen 
        For Each kvp As KeyValuePair(Of String, String) In initialRolesOfPhase

            If Not rolesToAdd.Contains(kvp.Key) Then
                rolesToDelete.Add(kvp.Key, kvp.Key)
            End If

        Next

        ' Schritt 2 - Kosten 
        For Each kvp As KeyValuePair(Of String, String) In initialCostsOfPhase

            If Not costsToAdd.Contains(kvp.Key) Then
                costsToDelete.Add(kvp.Key, kvp.Key)
            End If

        Next

        ' Ende Schritt 2
        ' 

        ' Schritt 3: alle, die bereits in initialRolesOfPhase sind, aus rolesToAdd rausnehmen, same for costs 
        For Each kvp As KeyValuePair(Of String, String) In initialRolesOfPhase

            If rolesToAdd.Contains(kvp.Key) Then
                rolesToAdd.Remove(kvp.Key)
            End If

        Next

        ' Schritt 2 - Kosten 
        For Each kvp As KeyValuePair(Of String, String) In initialCostsOfPhase

            If Not costsToAdd.Contains(kvp.Key) Then
                costsToDelete.Remove(kvp.Key)
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

                    topLevelNode.Name = RoleDefinitions.bestimmeRoleNodeName(role.UID, nrTag.membershipID)

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

        currentNode.Name = RoleDefinitions.bestimmeRoleNodeName(currentRoleUid, nrTag.membershipID)

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
    ''' gibt alle Namen von Knoten, die "gecheckt" sind, in der rolesToAdd- bzw. costsToAdd-Liste zurück  
    ''' wird rekursiv aufgerufen 
    ''' Achtung: wenn es Endlos Zyklen gibt, dann ist hier eine Endlos-Schleife ! 
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Public Sub pickupMECheckedRoleCostItems(ByVal node As TreeNode)
        Dim tmpNode As TreeNode

        If IsNothing(node) Then
            ' nichts tun
        Else

            Dim anzahlKnoten As Integer = node.Nodes.Count

            With node

                For px As Integer = 1 To anzahlKnoten

                    tmpNode = .Nodes.Item(px - 1)

                    If tmpNode.Checked Then

                        If CType(tmpNode.Tag, clsNodeRoleTag).isRole Then
                            If Not rolesToAdd.Contains(tmpNode.Name) Then
                                rolesToAdd.Add(tmpNode.Name, tmpNode.Name)
                            End If
                        Else
                            If Not costsToAdd.Contains(tmpNode.Name) Then
                                costsToAdd.Add(tmpNode.Name, tmpNode.Name)
                            End If
                        End If

                    End If

                    If tmpNode.Nodes.Count > 0 Then
                        Call pickupMECheckedRoleCostItems(tmpNode)
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

    Private Sub hryRoleCost_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles hryRoleCost.AfterCheck

        Dim node As TreeNode = e.Node

        If node.Checked = False Then
            Dim checkItem As String = node.Name
            ' es wurde unchecked ... das ist nur erlaubt, wenn die betreffende Rolle nicht bereits initial in der Phase war ... 
            If CType(node.Tag, clsNodeRoleTag).isRole Then

                If Not initialRolesOfPhase.ContainsKey(checkItem) Then
                    ' alles ok 
                Else
                    ' hier prüfen, ob es für diese Rolle in dieser Phase Istdaten gibt, denn darf nicht rausgenommen werden 
                    node.Checked = True
                    Call MsgBox("Rolle hat bereits Ist-Daten und kann deshalb nicht mehr gelöscht werden ...")
                End If
            Else
                If Not initialCostsOfPhase.ContainsKey(checkItem) Then
                    ' alles ok , weil checken, dann un-checken schon erlaubt ist 
                Else
                    node.Checked = True
                    Call MsgBox("bitte verwenden Sie zum Löschen einer Kostenart den Befehl 'Zeile löschen'.")
                End If
            End If
        End If

    End Sub
End Class