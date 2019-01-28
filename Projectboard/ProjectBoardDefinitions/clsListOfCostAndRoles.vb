''' <summary>
''' wird analog der Hierarchie für Meilensteine / Phasen für Rollen verwendet , um pro Projekt eine schnell auswertbare Liste zu haben, in welchen Phasen welche Rollen vorkommen 
''' wird verwendet um die Zugriffszeiten beim Aufbau von createPrcCollection Diagrammen drastisch zu reduzieren ...
''' </summary>
''' <remarks></remarks>
Public Class clsListOfCostAndRoles

    ''' <summary>
    ''' der erste schlüssel ist die RoleUID, dann kommt eine SortedList mit teamID ( ohne teamID = -1 und Collection mit PhaseNameIDs 
    ''' </summary>
    ''' <remarks></remarks>
    Private _listOfRoles As SortedList(Of Integer, SortedList(Of Integer, Collection))
    Private _listOfCosts As SortedList(Of Integer, Collection)

    ''' <summary>
    ''' gibt zurück, ob die roleUID, optional in der Eigenschaft als teamMEmber in der PhaseNAmeID auftaucht ..
    ''' </summary>
    ''' <param name="phaseNameID"></param>
    ''' <param name="roleUID"></param>
    ''' <param name="teamID"></param>
    ''' <returns></returns>
    Public ReadOnly Property phaseContainsRoleID(ByVal phaseNameID As String, ByVal roleUID As Integer, Optional ByVal teamID As Integer = -1) As Boolean

        Get
            Dim tmpResult As Boolean = False
            Dim found As Boolean = False

            If _listOfRoles.ContainsKey(roleUID) Then
                ' nur dann gibt es überhaupt irgendetwas zu dieser Rolle in dem Projekt 
                Dim memberList As SortedList(Of Integer, Collection) = _listOfRoles.Item(roleUID)

                If teamID = -1 Then
                    ' alle durchgehen 
                    For Each kvp As KeyValuePair(Of Integer, Collection) In memberList
                        If kvp.Value.Contains(phaseNameID) Then
                            found = True
                            Exit For
                        End If
                    Next

                Else
                    Dim listOfPhases As Collection = memberList.Item(teamID)
                    If Not IsNothing(listOfPhases) Then
                        If listOfPhases.Count > 0 Then
                            found = listOfPhases.Contains(phaseNameID)
                        End If
                    End If
                End If
            End If

            phaseContainsRoleID = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt zurück, ob die CostUID, in der PhaseNameID auftaucht ..
    ''' </summary>
    ''' <param name="phaseNameID"></param>
    ''' <param name="costUID"></param>
    ''' <returns></returns>
    Public ReadOnly Property phaseContainsCost(ByVal phaseNameID As String, ByVal costUID As Integer) As Boolean

        Get
            Dim tmpResult As Boolean = False
            Dim found As Boolean = False

            If _listOfCosts.ContainsKey(costUID) Then

                Dim listOfPhases As Collection = _listOfCosts.Item(costUID)

                If Not IsNothing(listOfPhases) Then
                    If listOfPhases.Count > 0 Then
                        found = listOfPhases.Contains(phaseNameID)
                    End If
                End If

            End If

            phaseContainsCost = tmpResult
        End Get

    End Property



    ''' <summary>
    ''' gibt die Phasen zurück, die diese Rolle enthalten 
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhasesWithRole(ByVal roleName As String, Optional ByVal teamID As Integer = -1) As Collection
        Get
            Dim phaseCollection As New Collection
            Dim role As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)

            If Not IsNothing(role) Then

                ' dann handelt es sich schon mal um eine gültige Rolle ...

                Dim roleUID As Integer = role.UID
                If _listOfRoles.ContainsKey(roleUID) Then
                    Dim memberList As SortedList(Of Integer, Collection) = _listOfRoles.Item(roleUID)

                    If teamID = -1 Then
                        ' alle holen 
                        phaseCollection = memberList.ElementAt(0).Value

                        ' falls es jetzt mehrere geben sollte, weil dieselbe PErson in diesem Projekt für mehrere Teams arbeitet ...
                        For i = 1 To memberList.Count - 1
                            Dim tmpCollection As Collection = memberList.ElementAt(i).Value
                            For Each phNameID As String In tmpCollection
                                If Not phaseCollection.Contains(phNameID) Then
                                    phaseCollection.Add(phNameID, phNameID)
                                End If
                            Next
                        Next


                    ElseIf teamID > 0 Then
                        ' nur die holen, die die Rolle in seiner Eigenschaft als Team-Member enthalten 
                        If memberList.ContainsKey(teamID) Then
                            phaseCollection = memberList.Item(teamID)
                        End If
                    End If

                Else
                    ' nichts tun, phaseCollection ist  eine leere Collection 
                End If

            End If

            getPhasesWithRole = phaseCollection
        End Get
    End Property

    '''' <summary>
    '''' gibt die Phasen zurück, die eine der Rollen aus der Collection enthält
    '''' wenn considerSubRoles = true, dann auch die Phasen, die eine oder mehrere SubRoles einer der Rollen aus der Collection enthalten 
    '''' </summary>
    '''' <param name="roleCollection"></param>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public ReadOnly Property getPhasesWithRoles(ByVal roleCollection As Collection) As Collection
    '    Get
    '        Dim phaseCollection As New Collection
    '        'Dim subRoleCollection As Collection

    '        If roleCollection.Count > 0 Then

    '            For Each roleName As String In roleCollection
    '                Dim role As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)
    '                Dim teilphaseCollection As Collection

    '                Dim roleUID As Integer = role.UID
    '                If _listOfRoles.ContainsKey(roleUID) Then
    '                    teilphaseCollection = _listOfRoles.Item(roleUID)
    '                Else
    '                    teilphaseCollection = New Collection
    '                End If

    '                ' jetzt muss teilphaseCollection mit phaseCollection gemerged werden ...
    '                For Each phaseName As String In teilphaseCollection
    '                    If Not phaseCollection.Contains(phaseName) Then
    '                        phaseCollection.Add(phaseCollection, phaseName)
    '                    End If
    '                Next

    '            Next
    '        End If

    '        getPhasesWithRoles = phaseCollection
    '    End Get
    'End Property


    ''' <summary>
    ''' gibt die Phasen zurück, die diese Kostenart enthalten 
    ''' </summary>
    ''' <param name="costName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhasesWithCost(ByVal costName As String) As Collection
        Get
            Dim phaseCollection As New Collection
            Dim cost As clsKostenartDefinition = CostDefinitions.getCostdef(costName)

            If Not IsNothing(cost) Then

                ' dann handelt es sich schon mal um eine gültige Kostenart ...

                Dim costUID As Integer = cost.UID
                If _listOfCosts.ContainsKey(costUID) Then
                    phaseCollection = _listOfCosts.Item(costUID)
                Else
                    ' nichts tun, tmpCollection ist bereits eine leere Collection 
                End If


            End If

            getPhasesWithCost = phaseCollection
        End Get
    End Property

    '''' <summary>
    '''' gibt die Phasen zurück, die eine der Kostenarten aus der Collection enthält
    '''' </summary>
    '''' <param name="costCollection"></param>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public ReadOnly Property getPhasesWithCosts(ByVal costCollection As Collection) As Collection
    '    Get
    '        Dim phaseCollection As New Collection

    '        If costCollection.Count > 0 Then

    '            For Each costName As String In costCollection

    '                Dim teilphaseCollection As Collection = Me.getPhasesWithCost(costName)

    '                ' jetzt muss teilphaseCollection mit phaseCollection gemerged werden ...
    '                For Each phaseName As String In teilphaseCollection
    '                    If Not phaseCollection.Contains(phaseName) Then
    '                        phaseCollection.Add(phaseCollection, phaseName)
    '                    End If
    '                Next

    '            Next
    '        End If

    '        getPhasesWithCosts = phaseCollection
    '    End Get
    'End Property

    ''' <summary>
    ''' liefert eine sortierte Collection mit allen vorkommenden Role-NameIDs zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleNameIDs As Collection
        Get
            Dim tmpCollection As New Collection
            Dim aufnehmen As Boolean = False

            For Each kvp As KeyValuePair(Of Integer, SortedList(Of Integer, Collection)) In _listOfRoles

                aufnehmen = False
                If RoleDefinitions.containsUid(kvp.Key) Then

                    For Each tkvp As KeyValuePair(Of Integer, Collection) In kvp.Value

                        Dim tmpRoleNameID As String = ""
                        If tkvp.Key = -1 Then
                            aufnehmen = True
                        ElseIf RoleDefinitions.containsUid(tkvp.Key) Then
                            aufnehmen = True
                        End If

                        If aufnehmen Then

                            tmpRoleNameID = RoleDefinitions.bestimmeRoleNameID(kvp.Key, tkvp.Key)

                            If Not tmpCollection.Contains(tmpRoleNameID) Then
                                tmpCollection.Add(tmpRoleNameID, tmpRoleNameID)
                            End If


                        End If

                    Next


                End If

            Next

            getRoleNameIDs = tmpCollection

        End Get
    End Property

    Public ReadOnly Property getRoleNames As Collection
        Get
            Dim tmpCollection As New Collection

            For Each kvp As KeyValuePair(Of Integer, SortedList(Of Integer, Collection)) In _listOfRoles

                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(kvp.Key)

                If Not IsNothing(tmpRole) Then
                    If Not tmpCollection.Contains(tmpRole.name) Then
                        tmpCollection.Add(tmpRole.name, tmpRole.name)
                    End If

                End If

            Next

            getRoleNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' liefert einen Array an UIDs zurück 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getRoleUIDs As Integer()
        Get

            Dim tmpResult() As Integer = Nothing
            If _listOfRoles.Count > 0 Then

                ReDim tmpResult(_listOfRoles.Count - 1)

                For i As Integer = 0 To _listOfRoles.Count - 1
                    tmpResult(i) = _listOfRoles.ElementAt(i).Key
                Next

            End If

            getRoleUIDs = tmpResult

        End Get
    End Property

    ''' <summary>
    ''' liefert eine sortierte Collection mit allen vorkommenden Kostenarten zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostNames As Collection
        Get
            Dim tmpCollection As New Collection

            For Each kvp As KeyValuePair(Of Integer, Collection) In _listOfCosts

                If kvp.Key >= 1 And kvp.Key <= CostDefinitions.Count Then
                    Dim costName As String = CostDefinitions.getCostdef(kvp.Key).name
                    tmpCollection.Add(costName, costName)
                End If

            Next

            getCostNames = tmpCollection
        End Get
    End Property

    ''' <summary>
    ''' ergänzt den Vermerk, dass Rolle mit roleUID in Phase mit Name phaseNameID vorkommt 
    ''' </summary>
    ''' <param name="roleUID"></param>
    ''' <param name="phaseNameID"></param>
    ''' <remarks></remarks>
    Public Sub addRP(ByVal roleUID As Integer, ByVal phaseNameID As String, ByVal Optional teamID As Integer = -1)

        If _listOfRoles.ContainsKey(roleUID) Then
            Dim memberlist As SortedList(Of Integer, Collection) = _listOfRoles.Item(roleUID)
            If memberlist.ContainsKey(teamID) Then
                Dim listOfPhases As Collection = memberlist.Item(teamID)

                If Not listOfPhases.Contains(phaseNameID) Then
                    listOfPhases.Add(phaseNameID, phaseNameID)
                Else
                    ' nichts tun , Phase ist schon drin ...
                End If
            Else
                Dim listOfPhases = New Collection
                listOfPhases.Add(phaseNameID, phaseNameID)
                memberlist.Add(teamID, listOfPhases)
            End If


        Else
            Dim memberlist As New SortedList(Of Integer, Collection)
            Dim listOfPhases = New Collection
            listOfPhases.Add(phaseNameID, phaseNameID)
            memberlist.Add(teamID, listOfPhases)

            _listOfRoles.Add(roleUID, memberlist)
        End If
    End Sub

    ''' <summary>
    ''' ergänzt den Vermerk, dass Kostenart mit costUID in Phase mit Name phaseNameID vorkommt 
    ''' </summary>
    ''' <param name="costUID"></param>
    ''' <param name="phaseNameID"></param>
    ''' <remarks></remarks>
    Public Sub addCP(ByVal costUID As Integer, ByVal phaseNameID As String)

        If _listOfCosts.ContainsKey(costUID) Then
            Dim listOfPhases As Collection = _listOfCosts.Item(costUID)
            If Not listOfPhases.Contains(phaseNameID) Then
                listOfPhases.Add(phaseNameID, phaseNameID)
            Else
                ' nichts tun , Phase ist schon drin ...
            End If
        Else
            Dim listOfPhases = New Collection
            listOfPhases.Add(phaseNameID, phaseNameID)
            _listOfCosts.Add(costUID, listOfPhases)
        End If
    End Sub

    ''' <summary>
    ''' löscht den Vermerk, dass Rolle mit roleUID, ggf teamID  in Phase mit Name phaseNameID ist 
    ''' </summary>
    ''' <param name="roleUID"></param>
    ''' <param name="phaseNameID"></param>
    ''' <remarks></remarks>
    Public Sub removeRP(ByVal roleUID As Integer, ByVal phaseNameID As String, ByVal Optional teamID As Integer = -1, ByVal Optional deleteAll As Boolean = False)

        If _listOfRoles.ContainsKey(roleUID) Then

            Dim memberships As SortedList(Of Integer, Collection) = _listOfRoles.Item(roleUID)

            If deleteAll Then

                Dim deleteList As New Collection
                For Each kvp As KeyValuePair(Of Integer, Collection) In memberships

                    If kvp.Value.Contains(phaseNameID) Then
                        kvp.Value.Remove(phaseNameID)
                        ' merken fpr folgendes löschen ...
                        If kvp.Value.Count = 0 Then
                            deleteList.Add(kvp.Key)
                        End If
                    End If
                Next

                For Each delTeamID As Integer In deleteList
                    memberships.Remove(delTeamID)
                Next

                If memberships.Count = 0 Then
                    _listOfRoles.Remove(roleUID)
                End If

            Else
                If memberships.ContainsKey(teamID) Then
                    Dim phList As Collection = memberships.Item(teamID)
                    If phList.Contains(phaseNameID) Then
                        phList.Remove(phaseNameID)
                    Else
                        ' nichts tun
                    End If

                    If phList.Count = 0 Then
                        memberships.Remove(teamID)
                    End If

                    If memberships.Count = 0 Then
                        _listOfRoles.Remove(roleUID)
                    End If

                End If

            End If

        Else
            ' nichts tun, Rolle gibt es nicht mehr  
        End If

    End Sub

    ''' <summary>
    ''' löscht den Vermerk, dass Kostenart mit costUID in Phase mit NAme phaseNameID ist 
    ''' </summary>
    ''' <param name="costUID"></param>
    ''' <param name="phaseNameID"></param>
    ''' <remarks></remarks>
    Public Sub removeCP(ByVal costUID As Integer, ByVal phaseNameID As String)

        If _listOfCosts.ContainsKey(costUID) Then
            Dim listOfPhases As Collection = _listOfCosts.Item(costUID)
            If listOfPhases.Contains(phaseNameID) Then
                listOfPhases.Remove(phaseNameID)
                If listOfPhases.Count = 0 Then
                    _listOfCosts.Remove(costUID)
                End If
            Else
                ' nichts tun , die Phase enthält die Rolle nicht 
            End If
        Else
            ' nichts tun, Rolle gibt es nicht mehr  
        End If

    End Sub

    Public Sub New()
        _listOfRoles = New SortedList(Of Integer, SortedList(Of Integer, Collection))
        _listOfCosts = New SortedList(Of Integer, Collection)
    End Sub

End Class
