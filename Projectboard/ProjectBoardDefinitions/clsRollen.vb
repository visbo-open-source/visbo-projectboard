''' <summary>
''' Die Rollen müssen immer in der customization file in der ursprünglichen Reihenfolge aufgeführt sein; 
''' ein Name kann umbenannt werden , aber er darf auf keinen Fall mit einer anderen ID im Customization File versehen werden 
''' 
''' </summary>
''' <remarks></remarks>
Public Class clsRollen


    Private _allRollen As SortedList(Of Integer, clsRollenDefinition)
    ' ist eine sortierte Liste von Namen der Rollen und ihrer zugehörigen ID 
    ' wird benötigt, um das Ganze zu beschelunigen
    Private _allNames As SortedList(Of String, Integer)

    Private _topLevelNodeIDs As List(Of Integer)

    Public Sub Add(roledef As clsRollenDefinition)

        ' Änderung tk: umgestellt auf 
        If Not _allRollen.ContainsKey(roledef.UID) Then
            _allRollen.Add(roledef.UID, roledef)
            If Not _allNames.ContainsKey(roledef.name) Then
                _allNames.Add(roledef.name, roledef.UID)
            Else
                Throw New ArgumentException(roledef.name & " existiert bereits")
            End If

        Else
            Throw New ArgumentException(roledef.UID.ToString & " existiert bereits")
        End If

    End Sub

    Public ReadOnly Property liste() As SortedList(Of Integer, clsRollenDefinition)
        Get
            liste = _allRollen
        End Get
    End Property


    ''' <summary>
    ''' gibt eine Liste aller Rollen zurück
    ''' der Value hat hier keine Bedeutung 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getAllIDs() As SortedList(Of Integer, Double)
        Get
            Dim tmpValue As Double = 1.0
            Dim tmpResult As New SortedList(Of Integer, Double)
            For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In _allRollen
                tmpResult.Add(kvp.Key, tmpValue)
            Next
            getAllIDs = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste aller Team IDs zurück 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getAllTeamIDs() As SortedList(Of Integer, Double)
        Get
            Dim tmpValue As Double = 1.0
            Dim tmpResult As New SortedList(Of Integer, Double)
            For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In _allRollen
                If kvp.Value.isTeam Then
                    If Not tmpResult.ContainsKey(kvp.Key) Then
                        tmpResult.Add(kvp.Key, tmpValue)
                    End If
                End If
            Next

            getAllTeamIDs = tmpResult

        End Get
    End Property

    ''' <summary>
    ''' gibt den Standard TopNode Name zurück, das ist der erste vorkommende Top Node 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getDefaultTopNodeName() As String
        Get
            Dim tmpName As String = ""
            If Not IsNothing(_topLevelNodeIDs) Then
                If _topLevelNodeIDs.Count > 0 Then
                    tmpName = _allRollen.Item(_topLevelNodeIDs.First).name
                End If
            End If
            getDefaultTopNodeName = tmpName
        End Get
    End Property
    ''' <summary>
    ''' gibt die Toplevel NodeIds zurück ...
    ''' Level 0 ist die erste Ebene, Level 1 die zweite. Weitere werden aktuell nicht unterstützt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTopLevelNodeIDs(Optional ByVal Level As Integer = 0) As List(Of Integer)
        Get
            Dim returnResult As New List(Of Integer)
            If Level = 0 Then
                returnResult = _topLevelNodeIDs
            ElseIf Level = 1 Then
                For Each roleID As Integer In _topLevelNodeIDs
                    Dim subroleList As SortedList(Of Integer, Double) = _allRollen.Item(roleID).getSubRoleIDs()
                    For Each srKvP As KeyValuePair(Of Integer, Double) In subroleList
                        If Not returnResult.Contains(srKvP.Key) Then
                            returnResult.Add(srKvP.Key)
                        End If
                    Next
                Next
            Else
                ' noch nicht implementiert - damit etwas zurückgegeben wird ... 
                ' leere Liste
            End If

            getTopLevelNodeIDs = returnResult

        End Get
    End Property

    ''' <summary>
    ''' gibt die Toplevel NodeIds zurück ...
    ''' Level 0 ist die erste Ebene, Level 1 die zweite. Weitere werden aktuell nicht unterstützt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTopLevelNodeNames(Optional ByVal Level As Integer = 0) As Collection
        Get
            Dim returnResult As New Collection

            If Level = 0 Then
                For Each roleID As Integer In _topLevelNodeIDs
                    If _allRollen.ContainsKey(roleID) Then
                        If Not returnResult.Contains(_allRollen.Item(roleID).name) Then
                            returnResult.Add(_allRollen.Item(roleID).name)
                        End If
                    End If
                Next

            ElseIf Level = 1 Then

                For Each roleID As Integer In _topLevelNodeIDs

                    Dim subroleList As SortedList(Of Integer, Double) = _allRollen.Item(roleID).getSubRoleIDs()

                    If subroleList.Count > 0 Then
                        For Each srKvP As KeyValuePair(Of Integer, Double) In subroleList
                            If _allRollen.ContainsKey(srKvP.Key) Then
                                If Not returnResult.Contains(_allRollen.Item(srKvP.Key).name) Then
                                    returnResult.Add(_allRollen.Item(srKvP.Key).name)
                                End If
                            End If
                        Next
                    Else
                        If _allRollen.ContainsKey(roleID) Then
                            If Not returnResult.Contains(_allRollen.Item(roleID).name) Then
                                returnResult.Add(_allRollen.Item(roleID).name)
                            End If
                        End If
                    End If

                Next

            Else
                ' noch nicht implementiert - damit etwas zurückgegeben wird ... 
                ' leere Liste
            End If

            getTopLevelNodeNames = returnResult

        End Get
    End Property

    ''' <summary>
    ''' gibt die eindeutige Liste an SammelRollen bzw. EinzelRollen wieder, die keiner Sammelrolle angehören 
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getUniqueRoleList() As Collection
        Get
            Dim tmpCollection As New Collection
            Dim sammelRollen As New Collection

            For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In _allRollen

                If kvp.Value.isCombinedRole Then

                    If Not sammelRollen.Contains(kvp.Value.name) Then
                        sammelRollen.Add(kvp.Value.name, kvp.Value.name)
                    End If

                End If

                ' jetzt die Rolle / Sammelrolle in tmpCollection aufnehmen 
                If Not tmpCollection.Contains(kvp.Value.name) Then
                    tmpCollection.Add(kvp.Value.name, kvp.Value.name)
                End If

            Next

            ' jetzt die Behandlung Sammelrolle machen 
            For Each sammelRolle As String In sammelRollen
                Dim subRoleList As SortedList(Of Integer, Double) = Me.getSubRoleIDsOf(roleName:=sammelRolle,
                                                                     type:=PTcbr.all)

                For Each srKvP As KeyValuePair(Of Integer, Double) In subRoleList

                    Dim subRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(srKvP.Key)
                    If Not IsNothing(subRole) Then
                        If ((tmpCollection.Contains(CStr(subRole.name))) And (subRole.name <> sammelRolle)) Then
                            tmpCollection.Remove(CStr(subRole.name))
                        End If
                    End If

                Next
            Next

            getUniqueRoleList = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Collection zurück, die nur die Rollen enthält , die keine Sammelrollen sind
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBasicRoles As Collection
        Get
            Dim tmpCollection As New Collection

            For r As Integer = 1 To _allRollen.Count
                Dim tmpRole As clsRollenDefinition = _allRollen.ElementAt(r - 1).Value
                If Not tmpRole.isCombinedRole Then
                    tmpCollection.Add(tmpRole.name, tmpRole.name)
                End If
            Next

            getBasicRoles = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Collection zurück, die nur die Rollen enthält , die Sammelrollen sind
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSummaryRoles(Optional ByVal roleName As String = Nothing) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim removeList As New Collection

            For r As Integer = 1 To _allRollen.Count
                Dim tmpRole As clsRollenDefinition = _allRollen.ElementAt(r - 1).Value
                If tmpRole.isCombinedRole Then
                    If IsNothing(roleName) Then
                        tmpCollection.Add(tmpRole.name, tmpRole.name)
                    ElseIf tmpRole.name <> roleName Then
                        tmpCollection.Add(tmpRole.name, tmpRole.name)
                    End If
                End If
            Next

            If Not IsNothing(roleName) Then

                Dim initialRole As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)
                If Not IsNothing(initialRole) Then

                    For sr As Integer = 1 To tmpCollection.Count
                        Dim tmpRole As clsRollenDefinition = Me.getRoledef(CStr(tmpCollection.Item(sr)))
                        Dim subRoleNames As SortedList(Of Integer, Double) = Me.getSubRoleIDsOf(tmpRole.name, PTcbr.all)

                        If Not subRoleNames.ContainsKey(initialRole.UID) Then
                            removeList.Add(tmpRole.name, tmpRole.name)
                        End If
                    Next

                    For rm As Integer = 1 To removeList.Count
                        tmpCollection.Remove(CStr(removeList.Item(rm)))
                    Next

                End If


            End If

            getSummaryRoles = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt zu der angegebenen Rolle die "Sammel-Rolle" zurück, die die Rolle als direkte Sub-Role enthält 
    ''' leerer String, wenn keine Sammel-Rolle existiert, die die angegebene Rolle enthält  
    ''' </summary>
    ''' <param name="roleUID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getParentRoleOf(ByVal roleUID As Integer) As clsRollenDefinition
        Get

            Dim sammelRollen As Collection = Me.getSummaryRoles
            Dim found As Boolean = False
            Dim i As Integer = 1
            Dim parentRole As clsRollenDefinition = Nothing

            If _allRollen.ContainsKey(roleUID) Then
                While Not found And i <= sammelRollen.Count

                    Dim tmpRole As clsRollenDefinition = Me.getRoledef(CStr(sammelRollen.Item(i)))
                    If Not IsNothing(tmpRole) Then
                        Dim subRoleIDs As SortedList(Of Integer, Double) = tmpRole.getSubRoleIDs
                        If subRoleIDs.ContainsKey(roleUID) Then
                            found = True
                            parentRole = tmpRole
                        Else
                            i = i + 1
                        End If
                    Else
                        i = i + 1
                    End If

                End While
            Else
                ' nichts tun ... 
            End If

            getParentRoleOf = parentRole

        End Get
    End Property

    Public Function hasAnyChildParentRelationsship(ByVal roleNameID As String, ByVal summaryRoleID As Integer) As Boolean
        Dim tmpResult As Boolean = False
        Dim teamID As Integer = -1

        Dim roleID As Integer = RoleDefinitions.parseRoleNameID(roleNameID, teamID)
        If roleID = summaryRoleID Then
            tmpResult = True

        Else
            Dim sRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(summaryRoleID)
            If Not IsNothing(sRole) Then
                Dim alleChildIDs As SortedList(Of String, Double) = RoleDefinitions.getSubRoleNameIDsOf(sRole.name, type:=PTcbr.all)
                If alleChildIDs.Count > 0 Then
                    tmpResult = alleChildIDs.ContainsKey(roleNameID)
                End If
            End If
        End If

        hasAnyChildParentRelationsship = tmpResult
    End Function

    ''' <summary>
    ''' gibt true zurück, wenn roleID irgendwo unterhalb der Hierarchy von summaryRoleID zu finden ist ..
    ''' das gilt für Team-Member ebenso wie für Orga-Mitglieder
    ''' </summary>
    ''' <param name="roleID"></param>
    ''' <param name="summaryRoleID"></param>
    ''' <returns></returns>
    Public Function hasAnyChildParentRelationsship(ByVal roleID As Integer, ByVal summaryRoleID As Integer) As Boolean

        Dim tmpResult As Boolean = False

        If roleID = summaryRoleID Then
            tmpResult = True
        Else
            Dim sRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(summaryRoleID)
            If Not IsNothing(sRole) Then
                Dim alleChildIDs As SortedList(Of Integer, Double) = RoleDefinitions.getSubRoleIDsOf(sRole.name, type:=PTcbr.all)
                If alleChildIDs.Count > 0 Then
                    tmpResult = alleChildIDs.ContainsKey(roleID)
                End If
            End If
        End If


        hasAnyChildParentRelationsship = tmpResult

    End Function

    ''' <summary>
    ''' Input ist eine sortierte Liste mit roleIds der Form roleUId;teamUid und die Uid einer Sammelrolle
    ''' True, wenn irgendeine roleID Kind der summaryRoleID ist 
    ''' </summary>
    ''' <param name="roleIDs"></param>
    ''' <param name="summaryRoleID"></param>
    ''' <returns></returns>
    Public Function hasAnyChildParentRelationsship(ByVal roleIDs As SortedList(Of String, String), ByVal summaryRoleID As Integer) As Boolean
        Dim tmpResult As Boolean = False

        For Each kvp As KeyValuePair(Of String, String) In roleIDs
            Dim teamID As Integer = -1
            Dim sRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(kvp.Value, teamID)

            If hasAnyChildParentRelationsship(sRole.UID, summaryRoleID) Then
                tmpResult = True
                Exit For
            End If

        Next

        hasAnyChildParentRelationsship = tmpResult

    End Function

    ''' <summary>
    ''' gibt true zurück, wenn die angegebene Rolle / Kostenart ein Kind oder Kindeskind eines der Elemente ist
    ''' oder das Element selber ist 
    ''' </summary>
    ''' <param name="roleCostName"></param>
    ''' <param name="roleCostCollection"></param>
    ''' <returns></returns>
    Public Function hasAnyChildParentRelationsship(ByVal roleCostName As String, ByVal roleCostCollection As Collection) As Boolean

        Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(roleCostName)
        Dim isRole As Boolean = RoleDefinitions.containsName(roleCostName)

        Dim iscost As Boolean = False
        Dim found As Boolean = False
        Dim ix As Integer = 1
        Dim myIDs As SortedList(Of Integer, Double)

        If isRole Then

            ' ist es eine Gruppe ...
            If tmpRole.isTeam Then
                myIDs = Me.getSubRoleIDsOf(roleCostName)
            Else
                myIDs = New SortedList(Of Integer, Double)
                Dim myUID As Integer = RoleDefinitions.getRoledef(roleCostName).UID
                myIDs.Add(myUID, 1.0)
            End If

            If roleCostCollection.Contains(roleCostName) Then
                found = True
            Else
                Do While Not found And ix <= roleCostCollection.Count

                    Dim parentName As String = CStr(roleCostCollection.Item(ix))

                    If RoleDefinitions.containsName(parentName) Then
                        Dim myUID As Integer = RoleDefinitions.getRoledef(roleCostName).UID
                        Dim childIDs As SortedList(Of Integer, Double) = Me.getSubRoleIDsOf(parentName)
                        Dim myIX As Integer = 0
                        Do While Not found And myIX <= myIDs.Count - 1
                            found = childIDs.ContainsKey(myIDs.ElementAt(myIX).Key)
                            If Not found Then
                                myIX = myIX + 1
                            End If
                        Loop

                    End If

                    If Not found Then
                        ix = ix + 1
                    End If

                Loop
            End If

        Else
            ' nichts tun, foudn = false lassen
        End If

        hasAnyChildParentRelationsship = found

    End Function
    ''' <summary>
    ''' bestimmt für eine Rolle im TreeView den Namen, der setzt sich zusammen aus RoleUid und ggf Membership Kennung 
    ''' </summary>
    ''' <param name="roleUID"></param>
    ''' <param name="teamID"></param>
    ''' <returns></returns>
    Public Function bestimmeRoleNameID(ByVal roleUID As Integer, ByVal teamID As Integer) As String
        ' der Name wird bestimmt, je nachdem ob es sich um eine normale Orga-Einheit , ein Team oder ein Team-Member handelt 

        Dim tmpResult As String = ""
        Dim ok As Boolean = True
        Dim isTeamMember As Boolean = (teamID > 0)

        If teamID > 0 Then
            ok = _allRollen.ContainsKey(teamID)
        End If

        If _allRollen.ContainsKey(roleUID) And ok Then
            Dim nodeName As String = roleUID.ToString

            If isTeamMember Then
                nodeName = roleUID.ToString & ";" & teamID.ToString
            Else
                nodeName = roleUID.ToString
            End If

            tmpResult = nodeName

        Else
            tmpResult = ""
        End If

        bestimmeRoleNameID = tmpResult

    End Function

    ''' <summary>
    ''' bestimmt den rollen-ID-String in der Form: roleUid;teamUid
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <param name="teamName"></param>
    ''' <returns></returns>
    Public Function bestimmeRoleNameID(ByVal roleName As String, ByVal teamName As String) As String
        Dim tmpResult As String = ""
        Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)

        Try
            If Not IsNothing(tmpRole) Then

                If teamName.Length > 0 Then
                    Dim tmpRoleTeam As clsRollenDefinition = RoleDefinitions.getRoledef(teamName)
                    If Not IsNothing(tmpRoleTeam) Then
                        If tmpRoleTeam.getSubRoleIDs.ContainsKey(tmpRole.UID) Then
                            tmpResult = RoleDefinitions.bestimmeRoleNameID(tmpRole.UID, tmpRoleTeam.UID)
                        Else
                            Dim dummy As Integer = -1
                            tmpResult = RoleDefinitions.bestimmeRoleNameID(tmpRole.UID, dummy)
                        End If
                    End If
                Else
                    tmpResult = RoleDefinitions.bestimmeRoleNameID(tmpRole.UID, -1)
                End If

            End If
        Catch ex As Exception
            tmpResult = ""
        End Try

        bestimmeRoleNameID = tmpResult
    End Function

    ''' <summary>
    ''' ähnlich wie getSubroleIDsOf , gibt die NameIDs in der Form roleUid;teamUid zurück  
    ''' </summary>
    ''' <param name="roleNameID">wird in der Form uid;teamId übergeben</param>
    ''' <param name="type"></param>
    ''' <param name="excludedNames">jeder Eintrag muss in der Form uid;teamID sein</param>
    ''' <returns></returns>
    Public ReadOnly Property getSubRoleNameIDsOf(ByVal roleNameID As String,
                                               Optional ByVal type As Integer = PTcbr.all,
                                               Optional ByVal excludedNames As Collection = Nothing) As SortedList(Of String, Double)
        Get

            ' hier muss überprüft werden, ob die myCollection Sammelrollen enthält 
            ' wenn ja, werden die alle solange um die enthaltenen Sammelrollen ergänzt, bis keine Sammelrolle mehr in der Collection drin ist  
            ' die Sammelrollen werden am Schluss wieder aufgenommen, weil sie ja als Platzhalter Rollen ihre Bedarfs-Werte auch mit geben müssen 

            Dim sammelRollenCollection As New SortedList(Of String, Double)
            Dim realCollection As New SortedList(Of String, Double)
            Dim addToRealCollection As New SortedList(Of String, Double)
            Dim noUntreatedCombinedRole As Boolean = False
            Dim teamID As Integer = -1
            Dim initialRole As clsRollenDefinition = getRoleDefByIDKennung(roleNameID, teamID)


            If Not IsNothing(initialRole) Then


                ' initial besetzen, um es in Gang zu setzen
                'realCollection.Add(roleName, roleName)

                realCollection.Add(roleNameID, 1.0)

                Do Until noUntreatedCombinedRole

                    noUntreatedCombinedRole = True

                    For Each kvp As KeyValuePair(Of String, Double) In realCollection

                        Dim roleDef As clsRollenDefinition = getRoleDefByIDKennung(kvp.Key, teamID)

                        If Not IsNothing(roleDef) Then

                            If roleDef.isCombinedRole Then

                                Dim curTeamID As Integer = -1

                                If roleDef.isTeam Then
                                    curTeamID = roleDef.UID
                                End If

                                If Not sammelRollenCollection.ContainsKey(kvp.Key) Then

                                    noUntreatedCombinedRole = False
                                    ' dann wurde sie nicht schon mal ersetzt  und die Kinder müssen aufgenommen werden  
                                    sammelRollenCollection.Add(kvp.Key, kvp.Value)

                                    Dim listofSubRoles As SortedList(Of Integer, Double) = roleDef.getSubRoleIDs

                                    If Not IsNothing(listofSubRoles) Then

                                        For Each srkvp As KeyValuePair(Of Integer, Double) In listofSubRoles

                                            Dim tmpKey As String = bestimmeRoleNameID(srkvp.Key, curTeamID)
                                            If Not realCollection.ContainsKey(tmpKey) And Not addToRealCollection.ContainsKey(tmpKey) Then
                                                addToRealCollection.Add(tmpKey, srkvp.Value)

                                            ElseIf addToRealCollection.ContainsKey(tmpkey) Then
                                                ' addieren, aber Gesamt-Summe darf nie größer 1 sein
                                                Dim newValue As Double = addToRealCollection(tmpKey) + srkvp.Value
                                                If newValue > 1.0 Then
                                                    newValue = 1.0
                                                End If
                                                addToRealCollection(tmpKey) = newValue
                                            End If


                                        Next

                                    Else
                                        ' darf eigentlich nicht sein , aber ist im Fehlerfall notwenig, um Endlos schleife zu verhindern 
                                        noUntreatedCombinedRole = True
                                    End If

                                End If

                            End If
                        End If


                    Next

                    ' jetzt müssen die addToRealCollection Items übertragen werden 
                    For Each kvp As KeyValuePair(Of String, Double) In addToRealCollection
                        If Not realCollection.ContainsKey(kvp.Key) Then
                            realCollection.Add(kvp.Key, kvp.Value)
                        Else
                            Dim newValue As Double = realCollection(kvp.Key) + kvp.Value
                            If newValue > 1.0 Then
                                newValue = 1.0
                            End If
                            realCollection(kvp.Key) = newValue
                        End If
                    Next

                    addToRealCollection.Clear()

                Loop

                ' jetzt müssen die realCollections ggf noch bereinigt werden: die Namen der Sammelrollen müssen raus

                If type = PTcbr.all Then
                    ' nichts tun - realCollections enthält schon alles 

                ElseIf type = PTcbr.placeholders Then
                    realCollection = sammelRollenCollection

                ElseIf type = PTcbr.realRoles Then
                    For Each cRKvp As KeyValuePair(Of String, Double) In sammelRollenCollection
                        If realCollection.ContainsKey(cRKvp.Key) Then
                            realCollection.Remove(cRKvp.Key)
                        End If
                    Next

                Else
                    ' nichts tun - realCollection enthält alles  
                End If


                If Not IsNothing(excludedNames) Then
                    ' jetzt müssen aus realCollection alle Namen raus, die in excludedNames drin sind ... 
                    For Each exclName As String In excludedNames

                        Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(exclName, teamID)

                        If Not IsNothing(tmpRole) Then
                            If realCollection.ContainsKey(exclName) And exclName <> roleNameID Then
                                realCollection.Remove(exclName)
                            End If
                        End If

                    Next
                End If
            End If


            getSubRoleNameIDsOf = realCollection


        End Get
    End Property


    ''' <summary>
    ''' gibt in einer eindeutigen Liste die Namen aller vorkommenden SubRoleIDs in einer sortierten Liste integer, double zurück, das heisst alle Platzhalter und die realen Rollen , oder nur die Platzhalter oder nur die realen Rollen  
    ''' es werden also alle Rollen-IDs zurückgegeben, Platzhalter und Basis Rollen, oder nur eine Kategorie davon 
    ''' wenn die excludedNames angegeben sind, dann werden nur die Rollen aufgenommen, die nicht in den excluded Names drin sind. 
    ''' Das stellt sicher, dass im Falle einer Ressourcen Auswertung Rollen nicht dopplet gezählt werden, weil sie einmal als Sammerolle gewertet werden, einmal als explizit angegebene Rolle 
    ''' 
    ''' das funktioniert auch über mehrstufige Sammelrollen, also wenn Fig2 FIG22, FIG23, enthält, die wiederum Engineering enthalten, die wiederum Namen enthalten
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSubRoleIDsOf(ByVal roleName As String,
                                               Optional ByVal type As Integer = PTcbr.all,
                                               Optional ByVal excludedNames As Collection = Nothing) As SortedList(Of Integer, Double)

        Get

            ' hier muss überprüft werden, ob die myCollection Sammelrollen enthält 
            ' wenn ja, werden die alle solange um die enthaltenen Sammelrollen ergänzt, bis keine Sammelrolle mehr in der Collection drin ist  
            ' die Sammelrollen werden am Schluss wieder aufgenommen, weil sie ja als Platzhalter Rollen ihre Bedarfs-Werte auch mit geben müssen 

            Dim sammelRollenCollection As New SortedList(Of Integer, Double)
            Dim realCollection As New SortedList(Of Integer, Double)
            Dim addToRealCollection As New SortedList(Of Integer, Double)
            Dim noUntreatedCombinedRole As Boolean = False
            Dim initialRole As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)

            If Not IsNothing(initialRole) Then


                ' initial besetzen, um es in Gang zu setzen
                'realCollection.Add(roleName, roleName)
                realCollection.Add(initialRole.UID, 1.0)

                Do Until noUntreatedCombinedRole

                    noUntreatedCombinedRole = True

                    For Each kvp As KeyValuePair(Of Integer, Double) In realCollection

                        Dim roleDef As clsRollenDefinition = Me.getRoleDefByID(kvp.Key)

                        If Not IsNothing(roleDef) Then

                            If roleDef.isCombinedRole Then

                                If Not sammelRollenCollection.ContainsKey(kvp.Key) Then

                                    noUntreatedCombinedRole = False
                                    ' dann wurde sie nicht schon mal ersetzt  und die Kinder müssen aufgenommen werden  
                                    sammelRollenCollection.Add(kvp.Key, kvp.Value)

                                    Dim listofSubRoles As SortedList(Of Integer, Double) = roleDef.getSubRoleIDs

                                    If Not IsNothing(listofSubRoles) Then

                                        For Each srkvp As KeyValuePair(Of Integer, Double) In listofSubRoles


                                            If Not realCollection.ContainsKey(srkvp.Key) And Not addToRealCollection.ContainsKey(srkvp.Key) Then
                                                addToRealCollection.Add(srkvp.Key, srkvp.Value)

                                            ElseIf addToRealCollection.ContainsKey(srkvp.Key) Then
                                                ' addieren, aber Gesamt-Summe darf nie größer 1 sein
                                                Dim newValue As Double = addToRealCollection(srkvp.Key) + srkvp.Value
                                                If newValue > 1.0 Then
                                                    newValue = 1.0
                                                End If
                                                addToRealCollection(srkvp.Key) = newValue
                                            End If


                                        Next

                                    Else
                                        ' darf eigentlich nicht sein , aber ist im Fehlerfall notwenig, um Endlos schleife zu verhindern 
                                        noUntreatedCombinedRole = True
                                    End If

                                End If

                            End If
                        End If


                    Next

                    ' jetzt müssen die addToRealCollection Items übertragen werden 
                    For Each kvp As KeyValuePair(Of Integer, Double) In addToRealCollection
                        If Not realCollection.ContainsKey(kvp.Key) Then
                            realCollection.Add(kvp.Key, kvp.Value)
                        Else
                            Dim newValue As Double = realCollection(kvp.Key) + kvp.Value
                            If newValue > 1.0 Then
                                newValue = 1.0
                            End If
                            realCollection(kvp.Key) = newValue
                        End If
                    Next

                    addToRealCollection.Clear()

                Loop

                ' jetzt müssen die realCollections ggf noch bereinigt werden: die Namen der Sammelrollen müssen raus

                If type = PTcbr.all Then
                    ' nichts tun - realCollections enthält schon alles 

                ElseIf type = PTcbr.placeholders Then
                    realCollection = sammelRollenCollection

                ElseIf type = PTcbr.realRoles Then
                    For Each cRKvp As KeyValuePair(Of Integer, Double) In sammelRollenCollection
                        If realCollection.ContainsKey(cRKvp.Key) Then
                            realCollection.Remove(cRKvp.Key)
                        End If
                    Next

                Else
                    ' nichts tun - realCollection enthält alles  
                End If


                If Not IsNothing(excludedNames) Then
                    ' jetzt müssen aus realCollection alle Namen raus, die in excludedNames drin sind ... 
                    For Each exclName As String In excludedNames
                        Dim teamID As Integer = -1
                        Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(exclName, teamID)

                        If Not IsNothing(tmpRole) Then
                            If realCollection.ContainsKey(tmpRole.UID) And tmpRole.name <> roleName Then
                                realCollection.Remove(tmpRole.UID)
                            End If
                        End If

                    Next
                End If
            End If


            getSubRoleIDsOf = realCollection


        End Get
    End Property

    ''' <summary>
    ''' liefert true zurück, wenn alle Rollendefinitionen der einen Liste identisch mit der anderen sind
    ''' </summary>
    ''' <param name="vglDefinitionen"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglDefinitionen As clsRollen)
        Get
            Dim stillIdentical = True

            If Me.Count = vglDefinitionen.Count Then
                Dim i As Integer = 0
                Do While i < _allRollen.Count And stillIdentical
                    stillIdentical = _allRollen.ElementAt(i).Value.isIdenticalTo(vglDefinitionen.getRoledef(i + 1))
                    i = i + 1
                Loop

            Else
                stillIdentical = False
            End If

            isIdenticalTo = stillIdentical
        End Get
    End Property

    '
    '
    '
    Public ReadOnly Property Count() As Integer

        Get

            Count = _allRollen.Count

        End Get

    End Property

    ''' <summary>
    ''' prüft ob name in der Collection enthalten ist
    ''' </summary>
    ''' <param name="name">Typ String</param>
    ''' <value></value>
    ''' <returns>wahr, wenn name enthalten ist; falsch, sonst</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsName(name As String) As Boolean
        Get
            Dim found As Boolean = False
            If IsNothing(name) Then
                ' found bleibt auf false
            Else
                found = _allNames.ContainsKey(name)

                ' tk geändert 29.5.18
                'Dim ix As Integer = 0
                'Do While ix <= _allRollen.Count - 1 And Not found
                '    If _allRollen.ElementAt(ix).Value.name = name Then
                '        found = True
                '    Else
                '        ix = ix + 1
                '    End If
                'Loop
            End If

            containsName = found
        End Get
    End Property

    Public ReadOnly Property containsNameID(nameID As String) As Boolean
        Get
            Dim tmpResult As Boolean = False
            Dim teamID As Integer = -1
            Dim roleUID As Integer = parseRoleNameID(nameID, teamID)

            If nameID.Contains(";") And teamID = -1 Then
                ' nicht ok 
            ElseIf roleUID = -1 Then
                ' nicht ok 
            ElseIf roleUID > 0 And teamID = -1 Then
                ' alles ok 
                tmpResult = True
            ElseIf roleUID > 0 And teamID > 0 Then
                ' alles ok 
                tmpResult = True
            End If

            containsNameID = tmpResult

        End Get
    End Property



    ''' <summary>
    ''' gibt zurück, ob der Key bereits enthalten ist 
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsUid(uid As Integer) As Boolean
        Get

            containsUid = _allRollen.ContainsKey(uid)

        End Get
    End Property


    ''' <summary>
    ''' gibt die Rollen-Definition mit angegebenem Namen zurück 
    ''' </summary>
    ''' <param name="myitem"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoledef(ByVal myitem As String) As clsRollenDefinition

        Get
            Dim tmpValue As clsRollenDefinition = Nothing

            Dim found As Boolean = _allNames.ContainsKey(myitem)

            If found Then
                tmpValue = _allRollen.Item(_allNames.Item(myitem))
            End If
            'Dim ix As Integer = 0

            'Do While ix <= _allRollen.Count - 1 And Not found
            '    If _allRollen.ElementAt(ix).Value.name = myitem Then
            '        found = True
            '        tmpValue = _allRollen.ElementAt(ix).Value
            '    Else
            '        ix = ix + 1
            '    End If
            'Loop

            getRoledef = tmpValue


        End Get

    End Property

    ''' <summary>
    ''' 1 gibt das erste Element zurück, AnzahlItems das letzte 
    ''' </summary>
    ''' <param name="myitem"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoledef(ByVal myitem As Integer) As clsRollenDefinition

        Get


            If myitem > 0 And myitem <= _allRollen.Count Then
                getRoledef = _allRollen.ElementAt(myitem - 1).Value
            Else
                getRoledef = Nothing
            End If


        End Get

    End Property

    ''' <summary>
    ''' bestimmt aus dem übergebenen SelectedRolesItem Angaben wie RoleUID, ggf. die zugehörige TeamID
    ''' </summary>
    ''' <param name="selRoleItem"></param>
    ''' <param name="teamID"></param>
    ''' <returns></returns>
    Public Function parseRoleNameID(ByVal selRoleItem As String, ByRef teamID As Integer) As Integer
        ' der Name wird bestimmt, je nachdem ob es sich um eine normale Orga-Einheit , ein Team oder ein Team-Member handelt 

        Dim tmpStr() As String = selRoleItem.Split(New Char() {CChar(";")})
        Dim roleID As Integer = -1

        ' Vorbesetzung von teamID 
        teamID = -1

        ' die RoleUID bestimmen 
        If IsNumeric(tmpStr(0)) Then
            roleID = CInt(tmpStr(0))
            If _allRollen.ContainsKey(roleID) Then
                ' alles ok 
            Else
                roleID = -1
            End If
        Else
            If _allNames.ContainsKey(tmpStr(0)) Then
                roleID = _allNames.Item(tmpStr(0))
            Else
                roleID = -1
            End If
        End If

        ' bestimme teamID
        If tmpStr.Length = 2 Then
            ' hat noch Team Info  
            If IsNumeric(tmpStr(1)) Then
                teamID = CInt(tmpStr(1))
                If _allRollen.ContainsKey(teamID) Then
                    ' alles ok 
                Else
                    teamID = -1
                End If
            Else
                If _allNames.ContainsKey(tmpStr(1)) Then
                    teamID = _allNames.Item(tmpStr(1))
                Else
                    teamID = -1
                End If
            End If
        End If


        parseRoleNameID = roleID

    End Function
    ''' <summary>
    ''' gibt aus dem String die RollenDefinition zurück 
    ''' ebenso die evtl vorhandene Team-Zugehörigkeit
    ''' </summary>
    ''' <param name="idK"></param>
    ''' <returns></returns>
    Public Function getRoleDefByIDKennung(ByVal idK As String, ByRef teamID As Integer) As clsRollenDefinition

        teamID = -1
        Try
            getRoleDefByIDKennung = getRoleDefByID(parseRoleNameID(idK, teamID))
        Catch ex As Exception
            getRoleDefByIDKennung = Nothing
        End Try

    End Function
    ''' <summary>
    ''' gibt die Rolle mit der entsprechenden ID zurück ...
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleDefByID(ByVal uid As Integer) As clsRollenDefinition
        Get
            If _allRollen.ContainsKey(uid) Then
                getRoleDefByID = _allRollen.Item(uid)
            Else
                getRoleDefByID = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt zu gegebenem Team und Team-Member den Prozentsatz zurück, mit dem das Team-Member seine Kapa zur Verfügung stellt 
    ''' </summary>
    ''' <param name="parentUID"></param>
    ''' <param name="childID"></param>
    ''' <returns></returns>
    Public ReadOnly Property getMembershipPrz(ByVal parentUID As Integer, ByVal childID As Integer) As Double
        Get
            Dim tmpResult As Double = 0.0

            If _allRollen.ContainsKey(parentUID) And _allRollen.ContainsKey(childID) Then
                ' nur dann es einen Wert geben 
                Dim parentRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(parentUID)

                If parentRole.getSubRoleIDs.ContainsKey(childID) Then
                    tmpResult = parentRole.getSubRoleIDs.Item(childID)
                End If
            End If

            getMembershipPrz = tmpResult
        End Get
    End Property

    Public Sub New()

        _allRollen = New SortedList(Of Integer, clsRollenDefinition)
        _allNames = New SortedList(Of String, Integer)
        _topLevelNodeIDs = New List(Of Integer)

    End Sub

    ''' <summary>
    ''' baut die Hierarchie der Rollen auf; dabei muss hier nur der bzw. die Top Nodes aufgenommen werden 
    ''' in den clsRoleNode Elementen sind bereits die Kinder verzeichnet 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub buildTopNodes()
        ' TopKnoten aufbauen
        Dim i As Integer = 1
        Dim currentRole As clsRollenDefinition
        Dim hparent As New clsRollenDefinition

        'For i = 1 To _allRollen.Count

        While (i <= _allRollen.Count)

            ' Level 0 Knoten
            currentRole = _allRollen.ElementAt(i - 1).Value
            Dim parentRole As clsRollenDefinition = Me.getParentRoleOf(currentRole.UID)

            If IsNothing(parentRole) Then
                If Not Me._topLevelNodeIDs.Contains(currentRole.UID) Then
                    ' aufnehmen als Top Level Node ...
                    Me._topLevelNodeIDs.Add(currentRole.UID)
                End If
            End If

            i = i + 1

        End While

    End Sub


End Class
