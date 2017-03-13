Public Class clsHierarchyDB
    Public allNodes As SortedList(Of String, clsHierarchyNodeDB)

    ''' <summary>
    ''' kopiert aus einem HSP-Element in ein DB-Element
    ''' </summary>
    ''' <param name="hry"></param>
    ''' <remarks></remarks>
    Sub copyFrom(ByVal hry As clsHierarchy)

        Dim hryNode As clsHierarchyNode
        Dim elemID As String
        Dim hryNodeDB As clsHierarchyNodeDB

        For i = 1 To hry.count

            hryNodeDB = New clsHierarchyNodeDB

            elemID = hry.getIDAtIndex(i)
            If elemID = rootPhaseName Then
                elemID = rootPhaseNameDB
            End If
            If elemID.Contains(punktName) Then
                elemID = elemID.Replace(punktName, punktNameDB)
            End If
            hryNode = hry.nodeItem(i)
            hryNodeDB.copyFrom(hryNode)

            Me.allNodes.Add(elemID, hryNodeDB)

        Next

    End Sub

    ''' <summary>
    ''' kopiert aus einem DB Element in ein HSP Element 
    ''' </summary>
    ''' <param name="hry"></param>
    ''' <remarks></remarks>
    Sub copyTo(ByRef hry As clsHierarchy)

        Dim hryNode As clsHierarchyNode
        Dim elemID As String
        Dim hryNodeDB As clsHierarchyNodeDB

        For i = 1 To Me.allNodes.Count

            hryNode = New clsHierarchyNode

            elemID = Me.allNodes.ElementAt(i - 1).Key
            If elemID = rootPhaseNameDB Then
                elemID = rootPhaseName
            End If
            If elemID.Contains(punktNameDB) Then
                elemID = elemID.Replace(punktNameDB, punktName)
            End If
            hryNodeDB = Me.allNodes.ElementAt(i - 1).Value
            hryNodeDB.copyTo(hryNode)

            hry.copyNode(hryNode, elemID)

        Next

    End Sub

    Sub New()
        allNodes = New SortedList(Of String, clsHierarchyNodeDB)
    End Sub
End Class
