''' <summary>
''' wird analog der Hierarchie für Meilensteine / Phasen für Rollen verwendet , um pro Projekt eine schnell auswertbare Liste zu haben, in welchen Phasen welche Rollen vorkommen 
''' wird verwendet um die Zugriffszeiten beim Aufbau von createPrcCollection Diagrammen drastisch zu reduzieren ...
''' </summary>
''' <remarks></remarks>
Public Class clsListOfCostAndRoles

    ''' <summary>
    ''' der erste schlüssel ist die RoleUID, dann kommt eine Liste mit PhaseNameID und Phasen-Nummern 
    ''' </summary>
    ''' <remarks></remarks>
    Private _listOfRoles As SortedList(Of Integer, Collection)
    Private _listOfCosts As SortedList(Of Integer, Collection)


    ''' <summary>
    ''' gibt die Phasen zurück, die diese Rolle enthalten 
    ''' wenn considerSubRoles = true, dann auch die Phasen, die eine oder mehrere SubRoles enthalten 
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <param name="considerSubroles"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhasesWithRole(ByVal roleName As String, _
                                                   Optional ByVal considerSubroles As Boolean = False) As Collection
        Get
            Dim phaseCollection As New Collection
            Dim role As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)

            If Not IsNothing(role) Then

                ' dann handelt es sich schon mal um eine gültige Rolle ...

                If role.isCombinedRole And considerSubroles Then
                    Dim roleCollection As New Collection
                    roleCollection.Add(roleName, roleName)
                    phaseCollection = Me.getPhasesWithRoles(roleCollection, considerSubroles)
                Else
                    Dim roleUID As Integer = role.UID
                    If _listOfRoles.ContainsKey(roleUID) Then
                        phaseCollection = _listOfRoles.Item(roleUID)
                    Else
                        ' nichts tun, tmpCollection ist bereits eine leere Collection 
                    End If

                End If



            End If

            getPhasesWithRole = phaseCollection
        End Get
    End Property

    ''' <summary>
    ''' gibt die Phasen zurück, die eine der Rollen aus der Collection enthält
    ''' wenn considerSubRoles = true, dann auch die Phasen, die eine oder mehrere SubRoles einer der Rollen aus der Collection enthalten 
    ''' </summary>
    ''' <param name="roleCollection"></param>
    ''' <param name="considerSubRoles"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhasesWithRoles(ByVal roleCollection As Collection, _
                                                    Optional ByVal considerSubRoles As Boolean = False) As Collection
        Get
            Dim phaseCollection As New Collection
            Dim subRoleCollection As Collection

            If roleCollection.Count > 0 Then

                For Each roleName As String In roleCollection
                    Dim role As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)
                    Dim teilphaseCollection As Collection

                    If role.isCombinedRole And considerSubRoles Then
                        subRoleCollection = RoleDefinitions.getSubRoleNamesOf(roleName)
                        teilphaseCollection = Me.getPhasesWithRoles(roleCollection, considerSubRoles)
                    Else
                        Dim roleUID As Integer = role.UID
                        If _listOfRoles.ContainsKey(roleUID) Then
                            teilphaseCollection = _listOfRoles.Item(roleUID)
                        Else
                            teilphaseCollection = New Collection
                        End If

                    End If

                    ' jetzt muss teilphaseCollection mit phaseCollection gemerged werden ...
                    For Each phaseName As String In teilphaseCollection
                        If Not phaseCollection.Contains(phaseName) Then
                            phaseCollection.Add(phaseCollection, phaseName)
                        End If
                    Next

                Next
            End If

            getPhasesWithRoles = phaseCollection
        End Get
    End Property


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

    ''' <summary>
    ''' gibt die Phasen zurück, die eine der Kostenarten aus der Collection enthält
    ''' </summary>
    ''' <param name="costCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhasesWithCosts(ByVal costCollection As Collection) As Collection
        Get
            Dim phaseCollection As New Collection

            If costCollection.Count > 0 Then

                For Each costName As String In costCollection

                    Dim teilphaseCollection As Collection = Me.getPhasesWithCost(costName)

                    ' jetzt muss teilphaseCollection mit phaseCollection gemerged werden ...
                    For Each phaseName As String In teilphaseCollection
                        If Not phaseCollection.Contains(phaseName) Then
                            phaseCollection.Add(phaseCollection, phaseName)
                        End If
                    Next

                Next
            End If

            getPhasesWithCosts = phaseCollection
        End Get
    End Property

    ''' <summary>
    ''' liefert eine sortierte Collection mit allen vorkommenden Role-Names zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleNames As Collection
        Get
            Dim tmpCollection As New Collection

            For Each kvp As KeyValuePair(Of Integer, Collection) In _listOfRoles

                If kvp.Key >= 1 And kvp.Key <= RoleDefinitions.Count Then
                    Dim roleName As String = RoleDefinitions.getRoledef(kvp.Key).name
                    tmpCollection.Add(roleName, roleName)
                End If

            Next

            getRoleNames = tmpCollection

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
    Public Sub addRP(ByVal roleUID As Integer, ByVal phaseNameID As String)

        If _listOfRoles.ContainsKey(roleUID) Then
            Dim listOfPhases As Collection = _listOfRoles.Item(roleUID)
            If Not listOfPhases.Contains(phaseNameID) Then
                listOfPhases.Add(phaseNameID, phaseNameID)
            Else
                ' nichts tun , Phase ist schon drin ...
            End If
        Else
            Dim listOfPhases = New Collection
            listOfPhases.Add(phaseNameID, phaseNameID)
            _listOfRoles.Add(roleUID, listOfPhases)
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
    ''' löscht den Vermerk, dass Rolle mit roleUID in Phase mit NAme phaseNameID ist 
    ''' </summary>
    ''' <param name="roleUID"></param>
    ''' <param name="phaseNameID"></param>
    ''' <remarks></remarks>
    Public Sub removeRP(ByVal roleUID As Integer, ByVal phaseNameID As String)

        If _listOfRoles.ContainsKey(roleUID) Then
            Dim listOfPhases As Collection = _listOfRoles.Item(roleUID)
            If listOfPhases.Contains(phaseNameID) Then
                listOfPhases.Remove(phaseNameID)
                If listOfPhases.Count = 0 Then
                    _listOfRoles.Remove(roleUID)
                End If
            Else
                ' nichts tun , die Phase enthält die Rolle nicht 
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
        _listOfRoles = New SortedList(Of Integer, Collection)
        _listOfCosts = New SortedList(Of Integer, Collection)
    End Sub

End Class
