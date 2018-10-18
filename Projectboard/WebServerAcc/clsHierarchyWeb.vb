Imports ProjectBoardDefinitions
Public Class clsHierarchyWeb

    Public allNodes As List(Of clsHryNodeWeb)

    ''' <summary>
    ''' kopiert aus einem HSP-Element in ein Web-Element
    ''' </summary>
    ''' <param name="hry"></param>
    ''' <remarks></remarks>
    Sub copyFrom(ByVal hry As clsHierarchy)

        Dim hryNodeWeb As clsHryNodeWeb

        Dim hryNode As clsHierarchyNode
        Dim elemID As String


        For i = 1 To hry.count

            hryNodeWeb = New clsHryNodeWeb

            elemID = hry.getIDAtIndex(i)
            If elemID = rootPhaseName Then
                elemID = rootPhaseNameDB
            End If
            If elemID.Contains(punktName) Then
                elemID = elemID.Replace(punktName, punktNameDB)
            End If
            hryNode = hry.nodeItem(i)

            hryNodeWeb.hryNodeKey = elemID
            hryNodeWeb.hryNode.copyFrom(hryNode)

            Me.allNodes.Add(hryNodeWeb)

        Next


    End Sub

    ''' <summary>
    ''' kopiert aus einem Web Element in ein HSP Element 
    ''' </summary>
    ''' <param name="hry"></param>
    ''' <remarks></remarks>
    Sub copyTo(ByRef hry As clsHierarchy)

        Dim hryNode As clsHierarchyNode
        Dim elemID As String
        Dim hryNodeDB As clsHierarchyNodeDB

        'For i = 1 To Me.allNodes.Count
        For Each node As clsHryNodeWeb In Me.allNodes

            hryNode = New clsHierarchyNode

            'elemID = Me.allNodes.ElementAt(i - 1).Key
            elemID = node.hryNodeKey
            If elemID = rootPhaseNameDB Then
                elemID = rootPhaseName
            End If
            If elemID.Contains(punktNameDB) Then
                elemID = elemID.Replace(punktNameDB, punktName)
            End If

            hryNodeDB = node.hryNode
            hryNodeDB.copyTo(hryNode)

            hry.copyNode(hryNode, elemID)

        Next

    End Sub
    Sub New()
        'allNodes = New Dictionary(Of String, clsHierarchyNodeDB)
        allNodes = New List(Of clsHryNodeWeb)
    End Sub
End Class
