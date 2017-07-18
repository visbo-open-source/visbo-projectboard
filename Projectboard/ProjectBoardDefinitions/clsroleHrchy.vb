Public Class clsroleHrchy

    Private _allroleNodes As SortedList(Of Integer, clsroleNode)

    ''' <summary>
    ''' gibt die Gesamt-Anzahl Elemente = Anzahl Phasen plus Anzahl Meilensteine zurück  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property count As Integer
        Get
            count = Me._allroleNodes.Count
        End Get
    End Property

    ''' <summary>
    ''' gibt an, ob dieser Schlüssel bereits in der Liste vorhanden ist
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsKey(ByVal uniqueID As Integer) As Boolean
        Get
            containsKey = _allroleNodes.ContainsKey(uniqueID)
        End Get
    End Property

 
    ''' <summary>
    ''' gibt die ParentID zurück, die die Rolle roleID hat
    ''' wenn das Element gar nicht existiert wird -1 zurückgegeben; ebenso, wenn es keinen PArent gibt 
    ''' </summary>
    ''' <param name="roleID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getParentIDOfID(ByVal roleID As Integer) As Integer
        Get
            If _allroleNodes.ContainsKey(roleID) Then
                getParentIDOfID = _allroleNodes.Item(roleID).roleParent
            Else
                getParentIDOfID = -1
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt den Hierarchie Knoten zurück , der die angegebene uniqueID hat
    ''' VORSICHT: liefert Nothing zurück , wenn die uniqueID nicht existiert 
    ''' </summary>
    ''' <param name="id"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property nodeItem(ByVal id As Integer) As clsroleNode
        Get
            If _allroleNodes.ContainsKey(id) Then
                nodeItem = _allroleNodes.Item(id)
            Else
                nodeItem = Nothing
            End If
        End Get
    End Property
    ''' <summary>
    ''' gibt eine Collection zurück, die alle roleIds enthält mit Level = level
    ''' </summary>
    ''' <param name="level"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property nodes(ByVal level As Integer) As Collection
        Get
            Dim tmpcollection As New Collection
            For Each kvp As KeyValuePair(Of Integer, clsroleNode) In _allroleNodes
                If kvp.Value.level = level Then
                    tmpcollection.Add(kvp.Key)
                End If
            Next
            nodes = tmpcollection
        End Get
    End Property

    ''' <summary>
    ''' findet die rollen der obersten Ebene heraus
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property toplevelNodes As List(Of Integer)
        Get
            Dim tmpList As New List(Of Integer)

            For i = 1 To _allroleNodes.Count
                If _allroleNodes.Item(i).level = 0 And _
                    Not tmpList.Contains(_allroleNodes.Item(i).roleId) Then
                    tmpList.Add(_allroleNodes.Item(i).roleId)
                End If
            Next
            toplevelNodes = tmpList
        End Get
    End Property

    ''' <summary>
    ''' gibt den den Toplevel Knoten zu Rolle mit ID = id zurück
    ''' </summary>
    ''' <param name="id"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTopNode(ByVal id As Integer) As clsroleNode
        Get

            Dim aktNode As clsroleNode = _allroleNodes.Item(id)

            Dim hNode As clsroleNode = aktNode
            While hNode.level > 0
                hNode = _allroleNodes.Item(_allroleNodes.Item(hNode.roleId).roleParent)
            End While
            getTopNode = hNode
        End Get
    End Property

    ''' <summary>
    ''' Hinzufügen einer Rolle in die Hierarchie der Rollen
    ''' </summary>
    ''' <param name="rolle"></param>
    ''' <remarks></remarks>
    Public Sub add(ByVal rolle As clsroleNode)

        Dim newroleNode As New clsroleNode
        If Not _allroleNodes.ContainsKey(rolle.roleId) Then

            newroleNode.roleId = rolle.roleId
            newroleNode.level = rolle.level
            newroleNode.roleParent = rolle.roleParent
            newroleNode.childs = rolle.childs
           
            If rolle.roleParent > 0 Then
                If Not _allroleNodes.Item(rolle.roleParent).childs.Contains(rolle.roleId) Then
                    ' akt. rolle zu den Kinder der ParentRolle hinzufügen
                    _allroleNodes.Item(rolle.roleParent).childs.Add(rolle.roleId)
                End If
            End If
            ' akt. Rolle in die Liste aller Rollen einfügen
            _allroleNodes.Add(newroleNode.roleId, newroleNode)
        End If

    End Sub

    Public Sub New()
        _allroleNodes = New SortedList(Of Integer, clsroleNode)
    End Sub


End Class
