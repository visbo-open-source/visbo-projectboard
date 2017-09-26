Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports ClassLibrary1
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Public Class frmMEhryRoleCost

    Private allCosts As New Collection
    Private allRoles As New Collection

    Public pName As String
    Public vName As String
    Public phaseName As String
    Public rcName As String
    Public phaseNameID As String

    Private Sub frmMEhryRoleCost_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If frmCoord(PTfrm.rolecostME, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.top))
            Me.Left = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.left))
        Else
            Me.Top = 60
            Me.Left = 100
        End If

        Call buildMERoleTree()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

    End Sub

    Private Sub frmMEhryRoleCost_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        frmCoord(PTfrm.rolecostME, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.rolecostME, PTpinfo.left) = Me.Left
    End Sub

    Private Sub hryRoleCost_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles hryRoleCost.AfterSelect

    End Sub

    Public Sub buildMERoleTree()


        Dim topLevelNode As TreeNode
        Dim checkProj As Boolean = False

        With hryRoleCost

            .Nodes.Clear()
            .CheckBoxes = True


            ' alle Rollen zeigen 

            If RoleDefinitions.Count > 0 Then
                Dim topNodes As List(Of Integer) = RoleDefinitions.getTopLevelNodeIDs


                For i = 0 To topNodes.Count - 1
                    Dim role As clsRollenDefinition = RoleDefinitions.getRoleDefByID(topNodes.ElementAt(i))
                    topLevelNode = .Nodes.Add(role.name)
                    topLevelNode.Name = role.UID.ToString
                    topLevelNode.Text = role.name
                    ' hier muss gecheckt werden, welche Rollen in dem Projekt und dieser Phase, in der der Doppelclick erfolgte
                    ' vergeben sind. Diese sollen dann als gecheckt markiert sein

                    ' ''If selectedRoles.Contains(role.name) Then
                    ' ''    topLevelNode.Checked = True
                    ' ''End If

                    Dim listOfChildIDs As New SortedList(Of Integer, String)
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
        Dim childIds As SortedList(Of Integer, String) = currentRole.getSubRoleIDs
        Dim doItAnyWay As Boolean = False
        Dim listOfroleNames As Collection = ShowProjekte.getRoleNames()

        ' wenn die vorhandenen Rollen als Kind oder Kindeskind von currentRole vorkommen, dann doItAnyWay
        If currentRole.isCombinedRole Then
            'ur: vorübergehend: 

            'If currentRole.hasAnyOfThemAsChild(listOfroleNames) Then
            '    doItAnyWay = True
            'End If
        End If

        If ShowProjekte.getRoleNames().Contains(currentRole.name) Or doItAnyWay Then

            Dim newNode As TreeNode
            With parentNode
                newNode = .Nodes.Add(currentRole.name)
                newNode.Name = roleUid.ToString
                newNode.Text = currentRole.name
                ' hier muss gecheckt werden, welche Rollen in dem Projekt und dieser Phase, in der der Doppelclick erfolgte
                ' vergeben sind. Diese sollen dann als gecheckt markiert sein

                '' ''If selectedRoles.Contains(currentRole.name) Then
                '' ''    newNode.Checked = True
                '' ''End If
            End With

            For i = 0 To childIds.Count - 1

                Call buildMESubRoleTree(newNode, childIds.ElementAt(i).Key)

            Next
        End If

    End Sub

End Class