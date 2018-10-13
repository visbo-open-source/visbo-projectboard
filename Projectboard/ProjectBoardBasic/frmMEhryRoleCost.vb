Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports ClassLibrary1
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Public Class frmMEhryRoleCost

    Private selectedRC As New Collection
    Public ergItems As New Collection

    Public pName As String
    Public vName As String
    Public phaseName As String
    Public rcName As String
    Public phaseNameID As String
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

        'If Not IsNothing(pName) And pName <> "" Then
        '    hproj = ShowProjekte.getProject(pName)
        'End If

        Call buildMERoleTree()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim anzahlKnoten As Integer = hryRoleCost.Nodes.Count
        Dim tmpnode As TreeNode

        selectedRC.Clear()

        ' einsammeln der Rollen und Kosten die selektiert wurden
        With hryRoleCost

            For px As Integer = 1 To anzahlKnoten

                tmpNode = .Nodes.Item(px - 1)

                If tmpNode.Checked Then

                    If Not selectedRC.Contains(tmpNode.Text) Then
                        selectedRC.Add(tmpNode.Text, tmpNode.Text)
                    End If

                End If


                If tmpNode.Nodes.Count > 0 Then
                    Call pickupMECheckedRoleItems(tmpnode)
                End If

            Next

        End With

        Dim anzahlcheckedRoles As Integer = selectedRC.Count
        Dim anzahlNewRoles As Integer = 0

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

    Private Sub hryRoleCost_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles hryRoleCost.AfterSelect

    End Sub

    Public Sub buildMERoleTree()


        Dim hPhase As clsPhase = Nothing
        If Not IsNothing(hproj) Then
            hPhase = hproj.getPhaseByID(phaseNameID)
        End If

        Dim topLevelNode As TreeNode
        Dim checkProj As Boolean = False

        With hryRoleCost

            .Nodes.Clear()
            .CheckBoxes = True


            ' alle Rollen zeigen 

            If RoleDefinitions.Count > 0 Then
                Dim topNodes As List(Of Integer) = RoleDefinitions.getTopLevelNodeIDs

                ' wenn die Sicht eingeschränkt werden soll ... 
                If awinSettings.isRestrictedToOrgUnit.Length > 0 Then

                    If RoleDefinitions.containsName(awinSettings.isRestrictedToOrgUnit) Then

                        topNodes.Clear()
                        topNodes.Add(RoleDefinitions.getRoledef(awinSettings.isRestrictedToOrgUnit).UID)

                    End If

                End If

                For i = 0 To topNodes.Count - 1
                    Dim role As clsRollenDefinition = RoleDefinitions.getRoleDefByID(topNodes.ElementAt(i))

                    topLevelNode = .Nodes.Add(role.name)
                    topLevelNode.Name = role.UID.ToString
                    topLevelNode.Text = role.name


                    ' hier muss gecheckt werden, welche Rollen in dem Projekt und dieser Phase, in der der Doppelclick erfolgte
                    ' vergeben sind. Diese sollen dann als kursiv dargestellt werden, die aktuelle Rolle als gecheckt markiert sein

                    ' tk 30.5
                    If Not IsNothing(hphase) Then
                        If Not IsNothing(hphase.getRole(role.name)) Then

                            ' entsprechend kennzeichnen 
                            topLevelNode.NodeFont = existingRoleFont
                            topLevelNode.ForeColor = existingRoleColor

                            If role.name = rcName Then
                                topLevelNode.Checked = True
                            End If

                        End If
                    End If



                    Dim listOfChildIDs As New SortedList(Of Integer, Double)
                    Try
                        listOfChildIDs = role.getSubRoleIDs
                    Catch ex As Exception

                    End Try

                    If listOfChildIDs.Count > 0 Then
                        For ii As Integer = 0 To listOfChildIDs.Count - 1
                            Call buildMESubRoleTree(topLevelNode, listOfChildIDs.ElementAt(ii).Key)
                        Next
                    End If

                Next
            End If
            If CostDefinitions.Count > 0 Then
                For i = 1 To CostDefinitions.Count - 1
                    Dim cost As clsKostenartDefinition = CostDefinitions.getCostdef(i)

                    topLevelNode = .Nodes.Add(cost.name)
                    topLevelNode.Name = cost.UID.ToString
                    topLevelNode.Text = cost.name

                    If Not IsNothing(hphase) Then
                        If Not IsNothing(hphase.getCost(cost.name)) Then

                            ' entsprechend kennzeichnen 
                            topLevelNode.NodeFont = existingRoleFont
                            topLevelNode.ForeColor = existingRoleColor

                            If cost.name = rcName Then
                                topLevelNode.Checked = True
                            End If

                        End If
                    End If


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
    ''' <param name="roleUid"></param>
    ''' <remarks></remarks>
    Public Sub buildMESubRoleTree(ByRef parentNode As TreeNode, ByVal roleUid As Integer)


        Dim currentRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(roleUid)


        Dim hPhase As clsPhase = Nothing
        If Not IsNothing(hproj) Then
            hPhase = hproj.getPhaseByID(phaseNameID)
        End If

        Dim childIds As SortedList(Of Integer, Double) = currentRole.getSubRoleIDs
        Dim doItAnyWay As Boolean = False

        Dim newNode As TreeNode
        With parentNode

            newNode = .Nodes.Add(currentRole.name)
            newNode.Name = roleUid.ToString
            newNode.Text = currentRole.name

            ' hier muss gecheckt werden, welche Rollen in dem Projekt und dieser Phase, in der der Doppelclick erfolgte
            ' vergeben sind. Diese sollen dann als kursiv dargestellt werden, die aktuelle Rolle als gecheckt markiert sein

            If Not IsNothing(hPhase) Then
                If Not IsNothing(hPhase.getRole(currentRole.name)) Then

                    ' entsprechend kennzeichnen
                    newNode.NodeFont = existingRoleFont
                    newNode.ForeColor = existingRoleColor

                    If currentRole.name = rcName Then
                        newNode.Checked = True
                    End If

                End If
            End If


        End With

        For i = 0 To childIds.Count - 1

            Call buildMESubRoleTree(newNode, childIds.ElementAt(i).Key)

        Next
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

End Class