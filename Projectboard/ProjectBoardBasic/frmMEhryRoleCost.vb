Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports ClassLibrary1
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Public Class frmMEhryRoleCost

    ' stellt sicher, dass der Check-/Uncheck Event nicht endlos aufgerufen wird ... 
    Dim dontFireInCheck As Boolean = False

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

    ' die PhaseNameID der Zeile  
    Public phaseNameID As String

    ' der Rollen-Kosten Name in der Zeile 
    Public rcName As String

    ' die Rollen-ID in der form roleUid;teamID oder roleUid.tostring bzw. costuid.tostring 
    Public rcNameID As String

    ' der ggf dazugehörende Team-Name 
    Public skillName As String

    ' gibt an, was gezeigt werden soll 
    Public showSkillsOnly As Boolean



    ' das in der Zeile aktive Projekt
    Public hproj As clsProjekt

    Friend existingRoleFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75, System.Drawing.FontStyle.Regular)
    Friend normalRoleFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75, System.Drawing.FontStyle.Regular)
    Friend normalRoleColor As System.Drawing.Color = System.Drawing.Color.Black
    Friend existingRoleColor As System.Drawing.Color = System.Drawing.Color.DimGray

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub

    Private Sub frmMEhryRoleCost_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If frmCoord(PTfrm.rolecostME, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.top))
            Me.Left = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.left))
        Else
            Me.Top = 60
            Me.Left = 100
        End If

        If IsNothing(hproj) Then
            Me.Text = "Auswahl Rollen/Kosten für Excel-Export"
        Else

            ' welche Rollen & Kosten sind in der aktuellen Phase drin ... 
            initialRolesOfPhase = hproj.getRoleIDs(phaseNameID)
            initialCostsOfPhase = hproj.getCostIDs(phaseNameID)

            Dim tmpPhaseName As String = phaseName
            If phaseNameID = rootPhaseName Then
                tmpPhaseName = "gesamte Projektphase"
            Else
                tmpPhaseName = "Phase " & phaseName
            End If

            If awinSettings.englishLanguage Then
                If phaseName.Length > 40 Then
                    Me.Text = "Select Resources/Skills/Costs for  " & tmpPhaseName.Substring(0, 39)
                Else
                    Me.Text = "Select Resources/Skills/Costs for  " & tmpPhaseName
                End If
            Else
                If phaseName.Length > 40 Then
                    Me.Text = "Auswahl Ressourcen/Skills/Kosten für " & tmpPhaseName.Substring(0, 39)
                Else
                    Me.Text = "Auswahl Ressourcen/Skills/Kosten für " & tmpPhaseName
                End If
            End If


        End If


        ' wie lautet der Identifier der aktuellen Zeile: setzet sich zusammen aus roleuid;teamid
        ' der wird bereits beim Right Click ermittelt und steht in rcNameID - siehe oben ...
        If showSkillsOnly Then
            Call buildMESkillTree()
        Else
            Call buildMERoleTree()
        End If


    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click


        ' die rolesToAdd, rolesToDelete, costToAdd und costToDelete sind beriets bestimmt - 
        ' in Checked_changed ...
        'Dim anzahlKnoten As Integer = hryRoleCost.Nodes.Count
        'Dim tmpnode As TreeNode


        '' 1. bestimmen der Rollen und Kosten, die gecheckt sind ... -> in rolesToadd bzw. costsToAdd
        '' 2. alles aus initialRolesOfPhase, die nicht mehr in rolesToadd sind : in rolesTobeDeleted übernehmen ; same for costs
        '' 3. alle, die bereits in initialRolesOfPhase sind, aus rolesToAdd rausnehmen , same for costs 


        'With hryRoleCost

        '    ' Schritt 1: bestimmen der Rollen und Kosten, die gecheckt sind ... -> in rolesToadd bzw. costsToAdd
        '    For px As Integer = 1 To anzahlKnoten

        '        tmpnode = .Nodes.Item(px - 1)

        '        If tmpnode.Checked Then

        '            If CType(tmpnode.Tag, clsNodeRoleTag).isRole Then
        '                If Not rolesToAdd.Contains(tmpnode.Name) Then
        '                    rolesToAdd.Add(tmpnode.Name, tmpnode.Name)
        '                End If
        '            Else
        '                If Not costsToAdd.Contains(tmpnode.Name) Then
        '                    costsToAdd.Add(tmpnode.Name, tmpnode.Name)
        '                End If
        '            End If

        '        End If

        '        If tmpnode.Nodes.Count > 0 Then
        '            Call pickupMECheckedRoleCostItems(tmpnode)
        '        End If

        '    Next

        'End With

        '' Schritt 2: alles aus initialRolesOfPhase, die nicht mehr in rolesToadd sind : in rolesTobeDeleted übernehmen 
        '' Schritt 2 - Rollen 
        'For Each kvp As KeyValuePair(Of String, String) In initialRolesOfPhase

        '    If Not rolesToAdd.Contains(kvp.Key) Then
        '        rolesToDelete.Add(kvp.Key, kvp.Key)
        '    End If

        'Next

        '' Schritt 2 - Kosten 
        'For Each kvp As KeyValuePair(Of String, String) In initialCostsOfPhase

        '    If Not costsToAdd.Contains(kvp.Key) Then
        '        costsToDelete.Add(kvp.Key, kvp.Key)
        '    End If

        'Next

        '' Ende Schritt 2
        '' 

        '' Schritt 3: alle, die bereits in initialRolesOfPhase sind, aus rolesToAdd rausnehmen, same for costs 
        'For Each kvp As KeyValuePair(Of String, String) In initialRolesOfPhase

        '    If rolesToAdd.Contains(kvp.Key) Then
        '        rolesToAdd.Remove(kvp.Key)
        '    End If

        'Next

        '' Schritt 2 - Kosten 
        'For Each kvp As KeyValuePair(Of String, String) In initialCostsOfPhase

        '    If Not costsToAdd.Contains(kvp.Key) Then
        '        costsToDelete.Remove(kvp.Key)
        '    End If

        'Next

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

    ''' <summary>
    ''' wird aufgerufen, um den SkillTree aufzubauen 
    ''' </summary>
    Public Sub buildMESkillTree()

        Dim topLevelNode As TreeNode
        Dim checkProj As Boolean = False

        Dim curRole As clsRollenDefinition = Nothing

        With hryRoleCost

            .Nodes.Clear()
            .CheckBoxes = True

            ' alle Skills zeigen 

            If RoleDefinitions.Count > 0 Then
                Dim topNodes As List(Of Integer) = RoleDefinitions.getTopLevelTeamIDs
                'Dim orgaUnitID As Integer = -1



                ' wenn bereits eine Skill-Id steht, dann soll ab hier begonnen werden 
                If skillName <> "" Then
                    If RoleDefinitions.containsNameOrID(skillName) Then
                        topNodes.Clear()
                        Dim teamID As Integer = -1
                        Dim restrictedSkillID As Integer = RoleDefinitions.getRoledef(skillName).UID
                        topNodes.Add(restrictedSkillID)
                    End If

                End If

                ' wenn rcname belegt ist : prüfen, ob diese Skill überhaupt dazu passt 
                If rcName <> "" Then
                    curRole = RoleDefinitions.getRoledef(rcName)
                End If


                For i = 0 To topNodes.Count - 1

                    Dim skill As clsRollenDefinition = RoleDefinitions.getRoleDefByID(topNodes.ElementAt(i))
                    Dim weitermachen As Boolean = True
                    ' erst prüfen, ob die Rolle überhaupt zu den aktiven Rollen zählt, also im Zeitraum aktiv ist 

                    If Not IsNothing(curRole) Then
                        weitermachen = RoleDefinitions.getCommonChildsOfParents(curRole.UID, skill.UID).Count > 0
                    End If


                    topLevelNode = .Nodes.Add(skill.name)
                    topLevelNode.Text = skill.name

                    If weitermachen Then
                        Dim nrTag As New clsNodeRoleTag
                        With nrTag
                            If skill.getSubRoleCount > 0 And Not isAggregationRole(skill) And Not skill.isSkillLeaf Then
                                .pTag = "P"
                                topLevelNode.Nodes.Clear()
                                topLevelNode.Nodes.Add("-")
                            Else
                                .pTag = "X"
                            End If
                        End With

                        ' tk 6.12.18 jetzt kommen ggf an einen Knoten noch diese Informationen
                        ' toplevelNode kann nur Team sein, nicht Team-Member
                        nrTag.isSkill = True
                        nrTag.isRole = False
                        nrTag.isTeamMember = False


                        topLevelNode.Tag = nrTag

                        ' tk 11.10.20
                        ' hier muss unterschieden werden , ob SkillName = "" ist oder nicht 
                        'topLevelNode.Name = RoleDefinitions.bestimmeRoleNameID(skill.UID, orgaUnitID)
                        topLevelNode.Name = skill.name
                    End If


                Next
            End If


        End With
    End Sub


    ''' <summary>
    ''' es werden hier nur die Organisations-Einheiten angezeigt ... 
    ''' die Skills kommen im buidMESkillTree / buildMESubSkillTree
    ''' </summary>
    Public Sub buildMERoleTree()


        Dim topLevelNode As TreeNode
        Dim checkProj As Boolean = False

        With hryRoleCost

            .Nodes.Clear()
            .CheckBoxes = True


            ' alle Rollen zeigen 
            If visboZustaende.meModus = ptModus.massEditRessSkills Then

                If RoleDefinitions.Count > 0 Then
                    Dim topNodes As List(Of Integer) = RoleDefinitions.getTopLevelNodeIDs


                    ' wenn bereits eine Orga-Id steht, dann soll ab hier begonnen werden 
                    Dim done As Boolean = False
                    If rcNameID <> "" Then
                        If RoleDefinitions.containsNameOrID(rcNameID, strongTest:=False) Then
                            topNodes.Clear()
                            Dim teamID As Integer = -1
                            Dim restrictedToOrgaID As Integer = RoleDefinitions.getRoleDefByIDKennung(rcNameID, teamID).UID
                            topNodes.Add(restrictedToOrgaID)
                        End If

                    ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Or
                    myCustomUserRole.customUserRole = ptCustomUserRoles.InternalViewer Then

                        If myCustomUserRole.specifics.Length > 0 Then
                            If RoleDefinitions.containsNameOrID(myCustomUserRole.specifics, strongTest:=False) Then

                                topNodes.Clear()
                                Dim teamID As Integer = -1
                                Dim restrictedToOrgaID As Integer = RoleDefinitions.parseRoleNameID(myCustomUserRole.specifics, teamID)
                                topNodes.Add(restrictedToOrgaID)

                                ' tk 11.10.20 nein, di emüssen heir überhaupt nicht angeteigt werdne
                                ' hier müssen jetzt auch die Skillgruppen angezeigt werden 
                                'Dim topLevelTeams As List(Of Integer) = RoleDefinitions.getTopLevelTeamIDs

                                'For Each topTeamID As Integer In topLevelTeams
                                '    Dim listOFCommonChildIds As List(Of Integer) = RoleDefinitions.getCommonChildsOfParents(topTeamID, restrictedToOrgaID)
                                '    If listOFCommonChildIds.Count > 0 Then
                                '        If Not topNodes.Contains(topTeamID) Then
                                '            topNodes.Add(topTeamID)
                                '        End If
                                '    End If

                                'Next

                            End If
                        End If
                    End If

                    For i = 0 To topNodes.Count - 1

                        Dim role As clsRollenDefinition = RoleDefinitions.getRoleDefByID(topNodes.ElementAt(i))

                        ' erst prüfen, ob die Rolle überhaupt zu den aktiven Rollen zählt, also im Zeitraum aktiv ist 
                        If role.isActiveRole And Not role.isSkill Then
                            topLevelNode = .Nodes.Add(role.name)
                            topLevelNode.Text = role.name


                            Dim nrTag As New clsNodeRoleTag
                            With nrTag
                                If role.getSubRoleCount > 0 And Not isAggregationRole(role) Then
                                    .pTag = "P"
                                    topLevelNode.Nodes.Clear()
                                    topLevelNode.Nodes.Add("-")
                                Else
                                    .pTag = "X"
                                End If
                            End With


                            topLevelNode.Tag = nrTag

                            ' tk 11.10.20
                            ' hier muss unterschieden werden , ob SkillName = "" ist oder nicht 
                            'topLevelNode.Name = RoleDefinitions.bestimmeRoleNameID(role.UID, nrTag.membershipID)
                            topLevelNode.Name = role.name

                            ' tk 11.10.20
                            ' ist die Rolle bereits in der Phase, die in der Zeile dargestellt wird ? 
                            'If initialRolesOfPhase.ContainsKey(topLevelNode.Name) Then
                            '    dontFireInCheck = True
                            '    topLevelNode.Checked = True
                            'End If

                            '' hier muss gecheckt werden, ob irgendwelche existierende Kind-Rollen unterhalb der aktuellen topNode sind 
                            '' Diese sollen dann als kursiv dargestellt werden, die aktuelle Rolle als gecheckt markiert sein

                            'If RoleDefinitions.hasAnyChildParentRelationsship(initialRolesOfPhase, role.UID) Then

                            '    ' entsprechend kennzeichnen 
                            '    topLevelNode.NodeFont = existingRoleFont
                            '    topLevelNode.ForeColor = existingRoleColor

                            'End If
                        End If

                    Next
                End If

            ElseIf visboZustaende.meModus = ptModus.massEditCosts Then

                If Not (myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager) Then
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


                        Next
                    End If
                End If

            End If




        End With
    End Sub
    Public Sub buildMESubSkillTree(ByRef parentNode As TreeNode, ByVal currentSkillID As Integer)

        Dim currentSkill As clsRollenDefinition = RoleDefinitions.getRoleDefByID(currentSkillID)

        Dim currentRole As clsRollenDefinition = Nothing
        If rcName <> "" Then
            currentRole = RoleDefinitions.getRoledef(rcName)
        End If

        If currentSkill.isActiveRole Then
            Dim childIds As SortedList(Of Integer, Double) = currentSkill.getSubRoleIDs

            Dim currentNode As TreeNode
            Dim childNode As TreeNode = Nothing

            Dim weiterMachen As Boolean = False

            If IsNothing(currentRole) Then
                weiterMachen = True
            Else
                weiterMachen = RoleDefinitions.roleHasSkill(currentRole.UID, currentSkill.UID)
            End If

            If weiterMachen Then
                currentNode = parentNode.Nodes.Add(currentSkill.name)
                currentNode.Text = currentSkill.name


                Dim nrTag As New clsNodeRoleTag


                If childIds.Count > 0 And Not isAggregationRole(currentSkill) And Not currentSkill.isSkillLeaf Then
                    ' hier muss - im Falle einer customUserRole = Portfolio Mgr bei der "letzten" Stufe abgebrochen werden
                    ' die dürfen also nicht die Personen sehen ... aber nur , wenn 
                    currentNode.Nodes.Clear()
                    currentNode.Nodes.Add("-")
                    nrTag.pTag = "P"
                Else
                    nrTag.pTag = "X"
                End If

                currentNode.Tag = nrTag

                'currentNode.Name = RoleDefinitions.bestimmeRoleNameID(currentRoleUid, nrTag.membershipID)
                currentNode.Name = currentSkill.name
            End If


        Else
            ' nur dazu da, um im Falle einer inaktiven Rolle jetzt im Debug anhalten zu können ... 
            Dim a As Integer = 2
        End If



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

        Dim currentSkill As clsRollenDefinition = Nothing
        If skillName <> "" Then
            currentSkill = RoleDefinitions.getRoledef(skillName)
        End If

        If currentRole.isActiveRole Then
            Dim childIds As SortedList(Of Integer, Double) = currentRole.getSubRoleIDs

            Dim currentNode As TreeNode
            Dim childNode As TreeNode = Nothing
            Dim weiterMachen As Boolean = False

            If IsNothing(currentSkill) Then
                weiterMachen = True
            Else
                weiterMachen = RoleDefinitions.roleHasSkill(currentRoleUid, currentSkill.UID)
            End If
            ' wenn eine Skill angegeben ist, dann darf der nur aufgenommen werden, wenn er die Skill hat 

            If weiterMachen Then
                currentNode = parentNode.Nodes.Add(currentRole.name)
                currentNode.Text = currentRole.name

                ' tk hier muss unterschieden werden, ob man Skills zeigt oder ob man Hierarchie zeigt ...
                Dim nrTag As New clsNodeRoleTag

                If childIds.Count > 0 And Not isAggregationRole(currentRole) Then
                    ' hier muss - im Falle einer customUserRole = Portfolio Mgr bei der "letzten" Stufe abgebrochen werden
                    ' die dürfen also nicht die Personen sehen ... aber nur , wenn 
                    currentNode.Nodes.Clear()
                    currentNode.Nodes.Add("-")
                    nrTag.pTag = "P"
                Else
                    nrTag.pTag = "X"
                End If

                currentNode.Tag = nrTag

                'currentNode.Name = RoleDefinitions.bestimmeRoleNameID(currentRoleUid, nrTag.membershipID)
                currentNode.Name = currentRole.name

                ' ist die Rolle bereits in der Phase, die in der Zeile dargestellt wird ? 
                'If initialRolesOfPhase.ContainsKey(currentNode.Name) Then
                '    dontFireInCheck = True
                '    currentNode.Checked = True
                'End If

                '' hier muss gecheckt werden, ob irgendwelche existierende Kind-Rollen unterhalb der aktuellen topNode sind 
                '' Diese sollen dann als kursiv dargestellt werden, die aktuelle Rolle als gecheckt markiert sein

                'If RoleDefinitions.hasAnyChildParentRelationsship(initialRolesOfPhase, currentRoleUid) Then

                '    ' entsprechend kennzeichnen 
                '    currentNode.NodeFont = existingRoleFont
                '    currentNode.ForeColor = existingRoleColor

                'End If
            End If

        Else
            ' nur dazu da, um im Falle einer inaktiven Rolle jetzt im Debug anhalten zu können ... 
            Dim a As Integer = 2
        End If



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

                    'Dim teamID As Integer
                    'Dim curRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(node.Name, teamID)
                    Dim curRole As clsRollenDefinition = RoleDefinitions.getRoledef(node.Name)
                    nodelist = curRole.getSubRoleIDs

                    ' tk 11.10.20 
                    'If myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager And Not curRole.isSkill Then

                    '    Dim virtualChilds As Integer() = RoleDefinitions.getVirtualChildIDs(curRole.UID, True)
                    '    If Not IsNothing(virtualChilds) Then
                    '        For Each vcID As Integer In virtualChilds
                    '            If Not nodelist.ContainsKey(vcID) Then
                    '                nodelist.Add(vcID, 1.0)
                    '            End If
                    '        Next
                    '    End If

                    'Else
                    '    ' wenn es sich um einen Ressourcen Manager handelt, kommen jetzt nur die Kinder zurück, die auch zum Ressort des 
                    '    ' Ressource Managers gehören 
                    '    nodelist = curRole.getSubRoleIDs
                    'End If


                    anzChilds = nodelist.Count
                Catch ex As Exception
                    anzChilds = 0
                End Try


                With hryRoleCost
                    .CheckBoxes = True
                End With

                For i As Integer = 0 To anzChilds - 1

                    If showSkillsOnly Then
                        Call buildMESubSkillTree(node, nodelist.ElementAt(i).Key)
                    Else
                        Call buildMESubRoleTree(node, nodelist.ElementAt(i).Key)
                    End If


                Next

            End If

        End If

    End Sub

    Private Sub hryRoleCost_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles hryRoleCost.AfterCheck

        Dim node As TreeNode = e.Node

        If dontFireInCheck Then
            dontFireInCheck = False
        Else
            Dim checkItem As String = node.Name
            ' un-Checked ...
            If node.Checked = False Then

                ' es wurde unchecked ... webb sie bereits in initialRoles/initialcosts ist, dann muss sie in toDeleteRoles / to deleteCosts..
                If CType(node.Tag, clsNodeRoleTag).isRole Then

                    If Not initialRolesOfPhase.ContainsKey(checkItem) Then
                        ' aus rolesToAdd raustun: sie wurde gecheckt, dann unchecked  
                        If rolesToAdd.Contains(checkItem) Then
                            rolesToAdd.Remove(checkItem)
                        End If
                    Else
                        ' hier prüfen, ob es für diese Rolle in dieser Phase Istdaten gibt, denn darf nicht rausgenommen werden 
                        Dim sumActualValues As Double = 0.0

                        If IsNothing(hproj) Then
                            ' im Falle Excel Export etc. 
                        Else
                            sumActualValues = hproj.getPhaseRCActualValues(phaseNameID, checkItem, True, False).Sum
                        End If

                        If sumActualValues > 0 Then
                            Call MsgBox("Rolle hat bereits Ist-Daten und kann deshalb nicht mehr gelöscht werden ...")
                            dontFireInCheck = True
                            node.Checked = True
                        Else
                            ' initialroles enthält sie: also muss sie in rolesToDelete
                            If Not rolesToDelete.Contains(checkItem) Then
                                rolesToDelete.Add(checkItem, checkItem)
                            End If
                        End If

                    End If
                Else
                    If Not initialCostsOfPhase.ContainsKey(checkItem) Then
                        ' aus costsToAdd raustun: sie wurde gecheckt, dann unchecked  
                        If costsToAdd.Contains(checkItem) Then
                            costsToAdd.Remove(checkItem)
                        End If
                    Else
                        ' prüfen, ob die Rolle Istdaten enthält ? 
                        Dim sumActualValues As Double = 0.0

                        If IsNothing(hproj) Then
                            ' kann im Fall Excel Export sein ...
                        Else
                            sumActualValues = hproj.getPhaseRCActualValues(phaseNameID, checkItem, False, True).Sum
                        End If

                        If sumActualValues > 0 Then
                            Call MsgBox("Kostenart hat bereits Ist-Daten und kann deshalb nicht mehr gelöscht werden ...")
                            dontFireInCheck = True
                            node.Checked = True ' nimmt die de-selection wieder zurück 
                        Else
                            ' initialcosts enthält sie: also muss sie in costsToDelete
                            If Not costsToDelete.Contains(checkItem) Then
                                costsToDelete.Add(checkItem, checkItem)
                            End If
                        End If
                    End If
                End If
            Else
                ' Check des Knoten
                ' prüfen, ob die Phase überhaupt noch Zukunfts-Monate, also Forecast Monate hat, 
                ' in denen was eingegeben werden darf  
                Dim hasStillForecastMonthsOrOtherwiseOK As Boolean = True
                If IsNothing(hproj) Then
                    ' es kann bei Excel Export weitergemacht werden 
                Else
                    hasStillForecastMonthsOrOtherwiseOK = hproj.isPhaseWithForecastMonths(phaseNameID)
                End If

                If hasStillForecastMonthsOrOtherwiseOK Then

                    ' jetzt koommt die Behandlung für Check-.Role bzw Check-Cost 
                    If CType(node.Tag, clsNodeRoleTag).isRole Then

                        If Not initialRolesOfPhase.ContainsKey(checkItem) Then
                            ' in rolesToAdd reintun:   
                            If Not rolesToAdd.Contains(checkItem) Then
                                rolesToAdd.Add(checkItem, checkItem)
                            End If
                        Else
                            ' wurde unchecked, dann checked 
                            If rolesToDelete.Contains(checkItem) Then
                                rolesToDelete.Remove(checkItem)
                            End If

                        End If
                    Else
                        If Not initialCostsOfPhase.ContainsKey(checkItem) Then
                            ' is costsToAdd reintun: 
                            If Not costsToAdd.Contains(checkItem) Then
                                costsToAdd.Add(checkItem, checkItem)
                            End If
                        Else
                            ' wurde unchecked, jetzt wieder checked 
                            If costsToDelete.Contains(checkItem) Then
                                costsToDelete.Remove(checkItem)
                            End If

                        End If
                    End If
                Else
                    ' es gibt einen Fall, wo das trotzdem gehen soll 
                    If CType(node.Tag, clsNodeRoleTag).isRole And rolesToDelete.Contains(checkItem) Then
                        rolesToDelete.Remove(checkItem)
                    ElseIf Not CType(node.Tag, clsNodeRoleTag).isRole And costsToDelete.Contains(checkItem) Then
                        costsToDelete.Remove(checkItem)
                    Else
                        Call MsgBox("es gibt in dieser Phase keine Forecast Monate mehr ..." & vbLf &
                                "deshalb wird die Selektion wieder zurückgenommen ...")
                        dontFireInCheck = True
                        node.Checked = False ' nimmt die de-selection wieder zurück 
                    End If
                End If

            End If

        End If

    End Sub


End Class