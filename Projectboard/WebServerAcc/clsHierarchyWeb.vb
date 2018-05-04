
Imports ProjectBoardDefinitions
<Serializable()>
Public Class clsHierarchyWeb


    'Public allNodes As SortedList(Of String, clsHierarchyNodeDB)
    Public Property allNodes As Dictionary(Of String, clsHierarchyNodeDB)
        Get
            allNodes = _allNodes
        End Get
        Set(value As Dictionary(Of String, clsHierarchyNodeDB))

        End Set
    End Property
    Private ReadOnly Property _allNodes As Dictionary(Of String, clsHierarchyNodeDB)



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

        'For i = 1 To Me.allNodes.Count
        For Each kvp As KeyValuePair(Of String, clsHierarchyNodeDB) In Me.allNodes

            hryNode = New clsHierarchyNode

            'elemID = Me.allNodes.ElementAt(i - 1).Key
            elemID = kvp.Key
            If elemID = rootPhaseNameDB Then
                elemID = rootPhaseName
            End If
            If elemID.Contains(punktNameDB) Then
                elemID = elemID.Replace(punktNameDB, punktName)
            End If
            'hryNodeDB = Me.allNodes.ElementAt(i - 1).Value
            hryNodeDB = kvp.Value
            hryNodeDB.copyTo(hryNode)

            hry.copyNode(hryNode, elemID)

        Next

    End Sub

    'Sub New()
    '    allNodes = New Dictionary(Of String, clsHierarchyNodeDB)
    '    'allNodes = New SortedList(Of String, clsHierarchyNodeDB)
    'End Sub

    ''Public allNodes As List(Of clsHrchyNodeWeb)

    ''''' <summary>
    ''''' kopiert aus einem HSP-Element in ein DB-Element
    ''''' </summary>
    ''''' <param name="hry"></param>
    ''''' <remarks></remarks>
    ''Sub copyFrom(ByVal hry As clsHierarchy)

    ''    Dim hryNode As clsHierarchyNode
    ''    Dim elemID As String
    ''    Dim hryNodeDB As clsHierarchyNodeDB

    ''    For i = 1 To hry.count

    ''        hryNodeDB = New clsHierarchyNodeDB

    ''        elemID = hry.getIDAtIndex(i)
    ''        If elemID = rootPhaseName Then
    ''            elemID = rootPhaseNameDB
    ''        End If
    ''        If elemID.Contains(punktName) Then
    ''            elemID = elemID.Replace(punktName, punktNameDB)
    ''        End If
    ''        hryNode = hry.nodeItem(i)
    ''        hryNodeDB.copyFrom(hryNode)
    ''        Dim hnodeWeb As New clsHrchyNodeWeb
    ''        hnodeWeb.elemId = elemID
    ''        hnodeWeb.node = hryNodeDB
    ''        Me.allNodes.Add(hnodeWeb)

    ''    Next

    ''End Sub

    ''''' <summary>
    ''''' kopiert aus einem DB Element in ein HSP Element 
    ''''' </summary>
    ''''' <param name="hry"></param>
    ''''' <remarks></remarks>
    ''Sub copyTo(ByRef hry As clsHierarchy)

    ''    Dim hryNode As clsHierarchyNode
    ''    Dim elemID As String
    ''    Dim hryNodeDB As clsHierarchyNodeDB

    ''    For i = 1 To Me.allNodes.Count

    ''        hryNode = New clsHierarchyNode

    ''        elemID = Me.allNodes.ElementAt(i - 1).elemId
    ''        If elemID = rootPhaseNameDB Then
    ''            elemID = rootPhaseName
    ''        End If
    ''        If elemID.Contains(punktNameDB) Then
    ''            elemID = elemID.Replace(punktNameDB, punktName)
    ''        End If
    ''        hryNodeDB = Me.allNodes.ElementAt(i - 1).node
    ''        hryNodeDB.copyTo(hryNode)

    ''        hry.copyNode(hryNode, elemID)

    ''    Next

    ''End Sub

    ''Sub New()
    ''    allNodes = New List(Of clsHrchyNodeWeb)
    ''End Sub
    ''Public Property count As Integer


End Class
