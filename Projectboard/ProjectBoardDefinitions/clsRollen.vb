''' <summary>
''' Die Rollen sind in der Orga-Struktur als sortierte Liste von RollenDefinitionen drin. 
''' Sortier-Key nach der uid (integer) der RollenDefinition. 
''' 
''' </summary>
''' <remarks></remarks>
Public Class clsRollen

    ' Liste ist nach UID sortiert
    Private _allRollen As SortedList(Of Integer, clsRollenDefinition)

    ' ist eine sortierte Liste von Namen und Aliases der Rollen und ihrer zugehörigen ID 
    ' wird benötigt, um das Ganze zu beschleunigen
    Private _allNames As SortedList(Of String, Integer)

    Private _topLevelNodeIDs As List(Of Integer)

    ' tk 4.5.19 eingeführt, um Teams den sie umfassenden Organisations-Einheiten zuordnen zu können
    ' im key ist die ID der Organisationseinheit, in der Liste sind die IDs der Teams, die virtuell zu dieser Orga-Einheit gehören 
    Private _orgaSkillChilds As SortedList(Of Integer, List(Of Integer))

    Private _topLevelSkillParents As List(Of Integer)

    ' tk 25.7.19 in welcher relativen Position im Baum sind die einzelnen IDs
    ' Key = NameID, value = relative Position
    Private _positionIndices As SortedList(Of String, Integer)

    ''' <summary>
    ''' wird aktuell nur in ImportMSProject benötigt .. wird gebraucht, um  unbekannte Rolen mit UID in missingRoleDefinitions aufzunehmen ..
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getFreeRoleID() As Integer
        Get
            getFreeRoleID = _allRollen.Last.Key + 1
        End Get
    End Property

    Public Sub Add(roledef As clsRollenDefinition)

        Dim errMsg As String = ""
        ' Änderung tk: umgestellt auf 

        If Not _allRollen.ContainsKey(roledef.UID) Then
            _allRollen.Add(roledef.UID, roledef)

            If Not _allNames.ContainsKey(roledef.name) Then
                _allNames.Add(roledef.name, roledef.UID)

                ' jetzt müssen noch die Alias-Namen aufgenommen werden, sofern es welche gibt ... 
                If Not IsNothing(roledef.aliases) Then
                    If roledef.aliases(0) <> "" Then
                        For Each aliasItem As String In roledef.aliases
                            If Not _allNames.ContainsKey(aliasItem) Then
                                _allNames.Add(aliasItem, roledef.UID)
                            Else
                                If awinSettings.englishLanguage Then
                                    errMsg = aliasItem & " already exists"
                                Else
                                    errMsg = aliasItem & " existiert bereits"
                                End If

                                Throw New ArgumentException(errMsg)
                            End If
                        Next
                    End If
                End If
            Else

                If awinSettings.englishLanguage Then
                    errMsg = roledef.name & " already exists"
                Else
                    errMsg = roledef.name & " existiert bereits"
                End If

                Throw New ArgumentException(errMsg)
            End If

        Else
            If awinSettings.englishLanguage Then
                errMsg = roledef.UID.ToString & " already exists"
            Else
                errMsg = roledef.UID.ToString & " existiert bereits"
            End If

            Throw New ArgumentException(errMsg)
        End If

    End Sub


    ''' <summary>
    ''' löscht die Rollendefinition roledef aus der Liste der Rollen einer Organisation
    ''' </summary>
    ''' <param name="roledef"></param>
    Public Sub remove(roledef As clsRollenDefinition)

        Dim errMsg As String = ""
        ' Änderung tk: umgestellt auf 

        If _allRollen.ContainsKey(roledef.UID) Then
            _allRollen.Remove(roledef.UID)

            If _allNames.ContainsKey(roledef.name) Then
                _allNames.Remove(roledef.name)

                ' jetzt müssen noch die Alias-Namen aufgenommen werden, sofern es welche gibt ... 
                If Not IsNothing(roledef.aliases) Then
                    If roledef.aliases(0) <> "" Then
                        For Each aliasItem As String In roledef.aliases
                            If _allNames.ContainsKey(aliasItem) Then
                                _allNames.Remove(aliasItem)
                            Else
                                If awinSettings.englishLanguage Then
                                    errMsg = aliasItem & " doesn't exists"
                                Else
                                    errMsg = aliasItem & " existiert nicht"
                                End If

                                Throw New ArgumentException(errMsg)
                            End If
                        Next
                    End If
                End If
            Else

                If awinSettings.englishLanguage Then
                    errMsg = roledef.name & " doesn't exists"
                Else
                    errMsg = roledef.name & " existiert nicht"
                End If

                Throw New ArgumentException(errMsg)
            End If

        Else
            If awinSettings.englishLanguage Then
                errMsg = roledef.UID.ToString & " doesn't exists"
            Else
                errMsg = roledef.UID.ToString & " existiert nicht"
            End If

            Throw New ArgumentException(errMsg)
        End If

    End Sub

    ''' <summary>
    ''' erstellt die virtuellen Zuordnungen von Teams zu ihren Organisations-Einheiten
    ''' die virtuelle Organisations- oder Eltern-Einheit ist die, die alle Team Member als Eltern umfasst
    ''' </summary>
    Public Sub buildOrgaSkillChilds()

        Dim errMsg As String = ""
        Dim alleTeams As SortedList(Of Integer, Double) = getAllSkillIDs

        _orgaSkillChilds = New SortedList(Of Integer, List(Of Integer))
        _topLevelSkillParents = New List(Of Integer)

        For Each kvp As KeyValuePair(Of Integer, Double) In alleTeams

            Try
                Dim parentArray As Integer() = getParentArray(getRoleDefByID(kvp.Key))

                ' jetzt jeden Parent, der nicht schon als Team gekennzeichnet ist, als team-Parent kennzeichnen 
                ' in parentArray(0) steht das Team-Element selber ... deshalb start der Schleife ab i=1
                For i As Integer = 1 To parentArray.Length - 1
                    Dim parentRole As clsRollenDefinition = getRoleDefByID(parentArray(i))
                    If Not parentRole.isSkill Then
                        parentRole.isSkillParent = True
                    End If
                Next


                If Not IsNothing(parentArray) Then
                    If parentArray.Length > 1 Then
                        If Not _topLevelSkillParents.Contains(parentArray.Last) Then
                            _topLevelSkillParents.Add(parentArray.Last)
                        End If
                    End If
                End If

            Catch ex As Exception
                errMsg = "Problems with team-structure ... Code 23879"
            End Try

            Try
                Dim commonParent As clsRollenDefinition = getContainingRoleOfSkillMembers(kvp.Key)

                If Not IsNothing(commonParent) Then
                    If _orgaSkillChilds.ContainsKey(commonParent.UID) Then

                        If Not _orgaSkillChilds.Item(commonParent.UID).Contains(kvp.Key) Then
                            _orgaSkillChilds.Item(commonParent.UID).Add(kvp.Key)
                        End If

                    Else
                        Dim otc As New List(Of Integer)
                        otc.Add(kvp.Key)
                        _orgaSkillChilds.Add(commonParent.UID, otc)
                    End If
                End If
            Catch ex As Exception
                errMsg = "Problems with team-structure ... Code 23880"
            End Try


        Next

        If errMsg <> "" Then
            Call MsgBox(errMsg)
        End If

    End Sub

    ''' <summary>
    ''' gibt zu einer Organisations-Einheit die virtuellen Childs zurück, das sind alle Skill-Gruppen, deren Mitglieder
    ''' diese Organisations-Einheit als kleinsten gemeinsamen Parent haben 
    ''' wenn auch angegeben, werden alle virtuellen Kinder / Enkel zurückgebracht 
    ''' gibt nothing zurück, wenn es keine gibt ..
    ''' </summary>
    ''' <param name="roleID"></param>
    ''' <returns></returns>
    Public ReadOnly Property getVirtualChildIDs(ByVal roleID As Integer, Optional ByVal inclSubRoles As Boolean = True) As Integer()
        Get
            Dim virtualChilds() As Integer = Nothing
            Dim ergebnisListe As New List(Of Integer)
            Try

                If Not _allRollen.Item(roleID).isSkill Then
                    If _allRollen.Item(roleID).isCombinedRole Then

                        If inclSubRoles Then
                            Dim roleName As String = getRoleDefByID(roleID).name
                            Dim subRoleIDs As SortedList(Of Integer, Double) = getSubRoleIDsOf(roleName)

                            For Each kvp As KeyValuePair(Of Integer, Double) In subRoleIDs

                                If Not IsNothing(_orgaSkillChilds) Then
                                    If _orgaSkillChilds.ContainsKey(kvp.Key) Then

                                        Dim teilErgebnis As List(Of Integer) = _orgaSkillChilds.Item(kvp.Key)
                                        For Each srID As Integer In teilErgebnis
                                            If Not ergebnisListe.Contains(srID) Then
                                                ergebnisListe.Add(srID)
                                            End If
                                        Next
                                    End If
                                End If

                            Next

                            virtualChilds = ergebnisListe.ToArray

                        Else

                            If Not IsNothing(_orgaSkillChilds) Then
                                If _orgaSkillChilds.ContainsKey(roleID) Then
                                    virtualChilds = _orgaSkillChilds.Item(roleID).ToArray
                                End If
                            End If

                        End If

                    End If
                End If


            Catch ex As Exception

            End Try

            getVirtualChildIDs = virtualChilds

        End Get
    End Property

    ''' <summary>
    ''' input is a roleID;SkillID String
    ''' </summary>
    ''' <param name="rcNameID"></param>
    ''' <returns>returns true if role has skill</returns>
    Public ReadOnly Property isValidCombination(ByVal rcNameID As String) As Boolean
        Get
            Dim tmpResult As Boolean = False
            Dim skillID As Integer = -1
            Try
                Dim roleID As Integer = RoleDefinitions.parseRoleNameID(rcNameID, skillID)

                If roleID > 0 And skillID > 0 Then
                    tmpResult = roleHasSkill(roleID, skillID)
                ElseIf roleID > 0 And skillID = -1 Then
                    tmpResult = True
                End If

            Catch ex As Exception
                tmpResult = False
            End Try
            isValidCombination = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt zurück ob die angegebene Rolle die Skill hat; Rolle kann eine OrgaUnit oder Person sein  
    ''' </summary>
    ''' <param name="roleID"></param>
    ''' <param name="skillID"></param>
    ''' <returns></returns>
    Public ReadOnly Property roleHasSkill(ByVal roleID As Integer, ByVal skillID As Integer) As Boolean
        Get
            Dim result As Boolean = False
            Dim curRole As clsRollenDefinition = getRoleDefByID(roleID)
            Dim curSkill As clsRollenDefinition = getRoleDefByID(skillID)

            If Not IsNothing(curRole) And Not IsNothing(curSkill) Then
                Try
                    If Not curRole.isSkill And curSkill.isSkill Then
                        If curRole.isCombinedRole Then
                            result = getCommonChildsOfParents(roleID, skillID).Count > 0
                        Else
                            result = curRole.getSkillIDs.ContainsKey(skillID)
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If

            roleHasSkill = result

        End Get
    End Property


    ''' <summary>
    ''' gibt zurück, ob die angegebene Rolle die Skill hat; Rolle kann eine OrgaUnit oder Person sein  
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <param name="skillName"></param>
    ''' <returns></returns>
    Public ReadOnly Property roleHasSkill(ByVal roleName As String, ByVal skillName As String) As Boolean
        Get
            Dim result As Boolean = False
            Dim curRole As clsRollenDefinition = getRoledef(roleName)
            Dim curSkill As clsRollenDefinition = getRoledef(skillName)

            If Not IsNothing(curRole) And Not IsNothing(curSkill) Then
                Try
                    If Not curRole.isSkill And curSkill.isSkill Then
                        If curRole.isCombinedRole Then
                            result = getCommonChildsOfParents(curRole.UID, curSkill.UID).Count > 0
                        Else
                            result = curRole.getSkillIDs.ContainsKey(curSkill.UID)
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If

            roleHasSkill = result

        End Get
    End Property

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
    Public ReadOnly Property getAllSkillIDs() As SortedList(Of Integer, Double)
        Get
            Dim tmpValue As Double = 1.0
            Dim tmpResult As New SortedList(Of Integer, Double)

            For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In _allRollen
                If kvp.Value.isSkill Then
                    If Not tmpResult.ContainsKey(kvp.Key) Then
                        tmpResult.Add(kvp.Key, tmpValue)
                    End If
                End If
            Next

            getAllSkillIDs = tmpResult

        End Get
    End Property


    ''' <summary>
    ''' gibt die Rolle aus der hierarchischen Organisation zurück, die alle Team-Members der teamID  enthält 
    ''' </summary>
    ''' <param name="skillID"></param>
    ''' <returns></returns>
    Public ReadOnly Property getContainingRoleOfSkillMembers(ByVal skillID As Integer) As clsRollenDefinition
        Get
            Dim tmpContainingRole As clsRollenDefinition = Nothing
            Dim listOfTopLevelNodeIds As List(Of Integer) = getTopLevelNodeIDs

            Try
                Dim skillRole As clsRollenDefinition = getRoleDefByID(skillID)
                If Not IsNothing(skillRole) Then
                    Dim allTeamMembers As SortedList(Of Integer, Double) = getSubRoleIDsOf(skillRole.name, type:=PTcbr.realRoles)


                    For Each kvp As KeyValuePair(Of Integer, Double) In allTeamMembers

                        ' tk 16.11.20 wenn eine realrole eine Skill ist, heisst das, dass diese Skill keine Team-Members hat ...
                        Dim checkRole As clsRollenDefinition = getRoleDefByID(kvp.Key)

                        ' nur untersuchen, wenn es nicht die Rolle selber ist und die chckRole keine Skill ohne Team-MEmber
                        If kvp.Key <> skillID And Not checkRole.isSkill Then
                            If IsNothing(tmpContainingRole) Then
                                tmpContainingRole = Me.getParentRoleOf(kvp.Key)
                            Else
                                tmpContainingRole = getCommonParent(tmpContainingRole, Me.getParentRoleOf(kvp.Key))
                            End If
                        End If

                    Next

                    ' Fehler abfangen, solange es Teams gibt, die kein isSkill Attribut haben 
                    If IsNothing(tmpContainingRole) Then
                        tmpContainingRole = getRoleDefByID(getTopLevelNodeIDs.First)
                    End If
                End If


            Catch ex As Exception

            End Try

            ' tk 16.11.20 wenn tmpContainingRole 

            getContainingRoleOfSkillMembers = tmpContainingRole
        End Get
    End Property


    ''' <summary>
    ''' gibt zu zwei Rollen die (Groß-)Eltern-Rolle zurück, die beide Rollen enthält 
    ''' </summary>
    ''' <param name="role1"></param>
    ''' <param name="role2"></param>
    ''' <returns></returns>
    Private Function getCommonParent(ByVal role1 As clsRollenDefinition, ByVal role2 As clsRollenDefinition) As clsRollenDefinition
        Dim tmpRole As clsRollenDefinition = Nothing

        Try
            If IsNothing(role1) Then
                tmpRole = role2
            ElseIf IsNothing(role2) Then
                tmpRole = role1
            Else
                ' beide sind ungleich Nothing
                If role1.UID = role2.UID Then
                    tmpRole = role1
                Else
                    ' beide sind ungleich Nothing und nicht identisch
                    Dim parentArray1() As Integer = getParentArray(role1)
                    Dim parentArray2() As Integer = getParentArray(role2)

                    Dim pA1() As Integer = Nothing
                    Dim pA2() As Integer = Nothing

                    ' jetzt wird mit aufsteigendem Index nach einem gemeinsamen Eltern-Teil gesucht 
                    If parentArray1.Count <= parentArray2.Count Then
                        pA1 = parentArray1
                        pA2 = parentArray2
                    Else
                        pA1 = parentArray2
                        pA2 = parentArray1
                    End If

                    ' jetzt ist pA1 der kleinere, höchstenns gleiche Array 
                    Dim ix1 As Integer = 0
                    Dim ix2 As Integer = 0

                    Dim found As Boolean = False
                    Do While ix1 <= pA1.Count - 1 And Not found
                        ix2 = 0
                        found = (pA1(ix1) = pA2(ix2))

                        Do While ix2 <= pA2.Count - 1 And Not found
                            found = (pA1(ix1) = pA2(ix2))
                            If Not found Then
                                ix2 = ix2 + 1
                            End If
                        Loop

                        If Not found Then
                            ix1 = ix1 + 1
                        End If
                    Loop

                    If found Then
                        tmpRole = getRoleDefByID(pA1(ix1))
                    Else
                        tmpRole = Nothing
                    End If

                End If
            End If
        Catch ex As Exception

        End Try


        getCommonParent = tmpRole

    End Function

    ''' <summary>
    ''' gibt den anzuzeigenden Indeltlevel an 
    ''' 0: nicht gefunden 
    ''' 1: ist bereits auf top Level
    ''' 2, .. Kind bzw. Kindeskind ...
    ''' </summary>
    ''' <param name="roleNameID"></param>
    ''' <returns></returns>
    Public Function getRoleIndent(ByVal roleNameID As String) As Integer

        Dim tmpResult As Integer = 0

        Dim teamID As Integer = -1
        Dim tmpRole As clsRollenDefinition = getRoleDefByIDKennung(roleNameID, teamID)

        If Not IsNothing(tmpRole) Then
            Try
                tmpResult = getParentArray(tmpRole).Count
            Catch ex As Exception

            End Try
        End If

        getRoleIndent = tmpResult

    End Function


    ''' <summary>
    ''' gibt die Eltern-/Groß-Eltern-ID bis zum höchsten Knoten zurück
    ''' wenn Role Nothing ist, kommt Nothing zurück 
    ''' die Liste enthält auch die Kind-Rolle als erstes Element , wenn also nur ein Element drin ist, dann ist es bereits ein TopLevelNode 
    ''' </summary>
    ''' <param name="role"></param>
    ''' <returns></returns>
    Public Function getParentArray(ByVal role As clsRollenDefinition, ByVal Optional includingMySelf As Boolean = True) As Integer()

        Dim tmpList As New List(Of Integer)

        If Not IsNothing(role) Then

            If includingMySelf Then
                tmpList.Add(role.UID)
            End If


            Dim parentRole As clsRollenDefinition = getParentRoleOf(role.UID)
            Do While Not IsNothing(parentRole)
                tmpList.Add(parentRole.UID)
                parentRole = getParentRoleOf(parentRole.UID)
            Loop

            ' jetzt muss die Liste in einen Array gewandelt werdne
            getParentArray = tmpList.ToArray
        Else
            getParentArray = Nothing
        End If


    End Function

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
    ''' gibt die Toplevel NodeIds zurück ...    ''' 
    ''' Level 0 ist die erste Ebene, Level 1 die zweite. Weitere werden aktuell nicht unterstützt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTopLevelNodeIDs(Optional ByVal Level As Integer = 0) As List(Of Integer)
        Get
            Dim returnResult As New List(Of Integer)
            If Level = 0 Then

                returnResult = _topLevelNodeIDs.ToList

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
    ''' gibt den oder die Top-Level Node IDs für Skillgruppen zurück 
    ''' leere Liste , wenn keine Teams / Skillgruppen existieren. 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getTopLevelTeamIDs() As List(Of Integer)
        Get

            getTopLevelTeamIDs = _topLevelSkillParents.ToList

        End Get
    End Property

    ''' <summary>
    ''' bestimmt die relativen Psoitions-Indizes der einzelnen Rollen 
    ''' wird benötigt, um sie in Reports, Tabellen, MassEdit in der erwarteten Reihenfolge darzustellen
    ''' </summary>
    Private Sub setRelativePositionIndicesOfRoles()
        Dim topLevelNodes As List(Of Integer) = getTopLevelNodeIDs
        Dim posIX As Integer = 1

        If Not IsNothing(topLevelNodes) Then

            For Each topLevelNodeID As Integer In topLevelNodes

                Call setrelativeIndicesOFParentNode(topLevelNodeID, -1, posIX)

            Next

        End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="nodeID">ID des Top Knotens dessen Baum analysiert werden soll </param>
    ''' <param name="teamID">ID des Teams </param>
    ''' <param name="posIX">byref übergebener PositionsIndex</param>
    Private Sub setrelativeIndicesOFParentNode(ByVal nodeID As Integer, ByVal teamID As Integer, ByRef posIX As Integer)

        Dim roleNameID As String = bestimmeRoleNameID(nodeID, teamID)

        If Not _positionIndices.ContainsKey(roleNameID) Then
            _positionIndices.Add(roleNameID, posIX)
            posIX = posIX + 1

            Dim parentrole As clsRollenDefinition = getRoleDefByID(nodeID)
            If parentrole.isSkill Then
                teamID = parentrole.UID
            Else
                teamID = -1
            End If

            For ci As Integer = 1 To parentrole.getSubRoleCount
                Dim childrole As clsRollenDefinition = getRoleDefByID(parentrole.getSubRoleIDs.ElementAt(ci - 1).Key)
                If childrole.getSubRoleCount = 0 Then

                    roleNameID = bestimmeRoleNameID(childrole.UID, teamID)

                    If Not _positionIndices.ContainsKey(roleNameID) Then
                        _positionIndices.Add(roleNameID, posIX)
                        posIX = posIX + 1
                    Else
                        ' darf eigentlich nicht sein 
                        Call MsgBox("Error: Position in Role-Definition " & childrole.name & " (" & roleNameID & ")")
                    End If
                Else
                    Call setrelativeIndicesOFParentNode(childrole.UID, teamID, posIX)
                End If

            Next
        End If

    End Sub

    ''' <summary>
    ''' gibt die Positions-Indices zurück 
    ''' im Fehlerfall werden die Role-IDs als Sortier-Kriterium verwnedte 
    ''' </summary>
    ''' <param name="NameIDListe"></param>
    ''' <returns></returns>
    Public Function getPositionIndices(ByVal NameIDListe As String()) As SortedList(Of Integer, String)
        Dim tmpResult As New SortedList(Of Integer, String)
        Dim errorPos As Integer = 999999

        Dim errorOccurred As Boolean = False
        Try
            For Each tmpID As String In NameIDListe
                If _positionIndices.ContainsKey(tmpID) Then
                    Dim posIX As Integer = _positionIndices.Item(tmpID)
                    Do While tmpResult.ContainsKey(posIX)
                        posIX = posIX + 1
                    Loop
                    tmpResult.Add(posIX, tmpID)
                Else
                    errorPos = errorPos + 1
                    Do While tmpResult.ContainsKey(errorPos)
                        errorPos = errorPos + 1
                    Loop
                    tmpResult.Add(errorPos, tmpID)
                End If

            Next
        Catch ex As Exception
            errorOccurred = True

        End Try

        getPositionIndices = tmpResult

    End Function

    ''' <summary>
    ''' gibt die Position im Orga-Baum als Positions-Index zurück. 
    ''' dient dazu eine für den Anwender nachvollziehbare, weil am Organisations-Baum orientierte Reihenfolge herzustellen  
    ''' </summary>
    ''' <param name="nameID"></param>
    ''' <returns></returns>
    Public Function getPositionIndex(ByVal nameID As String) As Integer
        Dim posIX As Integer

        Try
            If _positionIndices.ContainsKey(nameID) Then
                posIX = _positionIndices.Item(nameID)
            Else
                posIX = -1
            End If

        Catch ex As Exception
            posIX = -1
        End Try

        getPositionIndex = posIX
    End Function

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
                        Dim tmpName As String = _allRollen.Item(roleID).name
                        If Not returnResult.Contains(tmpName) Then
                            returnResult.Add(tmpName, tmpName)
                        End If
                    End If
                Next

            ElseIf Level = 1 Then

                For Each roleID As Integer In _topLevelNodeIDs

                    Dim subroleList As SortedList(Of Integer, Double) = _allRollen.Item(roleID).getSubRoleIDs()

                    If subroleList.Count > 0 Then
                        For Each srKvP As KeyValuePair(Of Integer, Double) In subroleList
                            If _allRollen.ContainsKey(srKvP.Key) Then
                                Dim tmpName As String = _allRollen.Item(srKvP.Key).name
                                If Not returnResult.Contains(tmpName) Then
                                    returnResult.Add(tmpName, tmpName)
                                End If
                            End If
                        Next
                    Else
                        If _allRollen.ContainsKey(roleID) Then
                            Dim tmpName As String = _allRollen.Item(roleID).name
                            If Not returnResult.Contains(tmpName) Then
                                returnResult.Add(tmpName, tmpName)
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

                    Dim subRole As clsRollenDefinition = Me.getRoleDefByID(srKvP.Key)
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
    ''' returns a list of intern employyes which are employed during the given timeframe and do have a default capacity of > 0  
    ''' Returns Nothing, if  es keine aktiven Internen im Zeitraum gibt ..
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getActiveInterns(ByVal vonDate As Date, ByVal bisDate As Date) As Integer()
        Get
            Dim tmpResult() As Integer = Nothing
            Dim tmpList As New SortedList(Of Integer, Boolean)

            For r As Integer = 1 To _allRollen.Count
                Dim tmpRole As clsRollenDefinition = _allRollen.ElementAt(r - 1).Value
                If Not tmpRole.isCombinedRole Then
                    If Not tmpRole.isExternRole Then

                        If tmpRole.isActiveRole And tmpRole.defaultKapa > 0 Then

                            Try
                                tmpList.Add(tmpRole.UID, True)
                            Catch ex As Exception

                            End Try

                        End If

                    End If

                End If
            Next

            If tmpList.Count > 0 Then
                tmpResult = tmpList.Keys.ToArray
            End If

            getActiveInterns = tmpResult

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
    ''' calculates for all intern employees fullCost, consisting of defaultKapa per month , multiplied with generalCostFactor multiplied with dayRate
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <returns></returns>
    Public ReadOnly Property getFullCost(ByVal von As Integer, ByVal bis As Integer) As Double()
        Get
            Dim tmpResult(0) As Double
            If von > 0 And bis >= von Then
                ReDim tmpResult(bis - von)
                ' correction faktor: multiply default value with correctionfaktor to get the cash-flow relevant value per Month
                ' company have to pay full cost , including illness, holiday, general cost factor,  
                ' tk 26.6 das muss parametrisiert werden ... 
                Dim generalCostFactor As Double = 1.22


                For Each topLEvelID As Integer In _topLevelNodeIDs
                    Dim roleName As String = getRoleDefByID(topLEvelID).name
                    Dim listOfAllChilds As SortedList(Of Integer, Double) = getSubRoleIDsOf(roleName)

                    For Each kvp As KeyValuePair(Of Integer, Double) In listOfAllChilds
                        Dim curRole As clsRollenDefinition = getRoleDefByID(kvp.Key)
                        If Not IsNothing(curRole) Then
                            If Not (curRole.isExternRole Or curRole.isCombinedRole) Then

                                Dim columnOfEntryDate As Integer = getColumnOfDate(curRole.entryDate)
                                Dim columnOfExitDate As Integer = getColumnOfDate(curRole.exitDate)

                                ' dann und nur dann handelt es sich um eine interne Person, die im Zeitraum auch aktiv beschäftigt und nicht extern ist 
                                For i As Integer = von To bis
                                    If i >= columnOfEntryDate And i < columnOfExitDate Then
                                        ' dann und nur dann muss die Person bezahlt werden ... 
                                        tmpResult(i - von) = tmpResult(i - von) + curRole.defaultKapa * generalCostFactor * curRole.tagessatzIntern / 1000
                                    End If
                                Next

                            End If
                        End If
                    Next
                Next

            End If
            getFullCost = tmpResult
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

                Dim initialRole As clsRollenDefinition = Me.getRoledef(roleName)
                If Not IsNothing(initialRole) Then

                    For sr As Integer = 1 To tmpCollection.Count
                        Dim tmpRole As clsRollenDefinition = Me.getRoledef(CStr(tmpCollection.Item(sr)))
                        Dim subRoleIDs As SortedList(Of Integer, Double) = Me.getSubRoleIDsOf(tmpRole.name, PTcbr.all)

                        If Not subRoleIDs.ContainsKey(initialRole.UID) Then
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
    ''' gibt true zurück, wenn es Kinder gibt, die Teams sind
    ''' </summary>
    ''' <param name="skillName"></param>
    ''' <returns></returns>
    Public ReadOnly Property isParentOfSkills(ByVal skillName As String) As Boolean
        Get
            Dim teamFound As Boolean = False
            Dim ausschluss As Collection = getTopLevelNodeNames

            Dim childNameIds As SortedList(Of Integer, Double) = getSubRoleIDsOf(skillName)


            For Each kvp As KeyValuePair(Of Integer, Double) In childNameIds
                teamFound = _allRollen.Item(kvp.Key).isSkill
                If teamFound Then
                    Exit For
                End If
            Next kvp

            isParentOfSkills = teamFound
        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn die angegebene roleUID Eltern-Teil von allen Teams ist 
    ''' wird für den Aufbau / den Ausschluss des obersten Team-Knotens beim Portfolio Manager benötigt 
    ''' </summary>
    ''' <param name="roleUID"></param>
    ''' <returns></returns>
    Private ReadOnly Property isParentOfAllSkills(ByVal roleUID As Integer) As Boolean
        Get
            Dim allTeamIDs As SortedList(Of Integer, Double) = getAllSkillIDs
            Dim tmpResult As Boolean = False
            Dim firstTime As Boolean = True

            For Each kvp As KeyValuePair(Of Integer, Double) In allTeamIDs

                If firstTime Then
                    firstTime = False
                    tmpResult = True
                End If

                tmpResult = tmpResult And hasAnyChildParentRelationsship(kvp.Key, roleUID)
                If tmpResult = False Then
                    Exit For ' es reicht, wenn eines nicht dazu gehört ... 
                End If

            Next

            isParentOfAllSkills = tmpResult
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

    ''' <summary>
    ''' gibt zu dem übergebenen String, der RoleNames in der Form D-BOSV-KB1; D-BOSV-KB2; etc enthält 
    ''' die gültigen NameIDs in Form eines Id-Arrays zurück 
    ''' </summary>
    ''' <param name="aufzaehlung"></param>
    ''' <returns></returns>
    Public ReadOnly Property getIDArray(ByVal aufzaehlung As String) As Integer()
        Get
            Dim tmpResult() As Integer = Nothing
            Dim realAnzahl As Integer = 0

            If IsNothing(aufzaehlung) Then
                ' nichts tun
            Else
                If aufzaehlung.Length > 0 Then
                    Dim tmpStr() As String = aufzaehlung.Split(New Char() {CChar(";")})
                    For Each tmpName As String In tmpStr
                        tmpName = tmpName.Trim
                        If Me.containsNameOrID(tmpName) Then
                            realAnzahl = realAnzahl + 1
                        End If
                    Next

                    If realAnzahl > 0 Then
                        ReDim tmpResult(realAnzahl - 1)
                        Dim ix As Integer = 0
                        For Each tmpName As String In tmpStr
                            tmpName = tmpName.Trim
                            If Me.containsNameOrID(tmpName) Then
                                Dim teamID As Integer
                                tmpResult(ix) = Me.getRoleDefByIDKennung(tmpName, teamID).UID
                                ix = ix + 1
                            End If
                        Next
                    End If
                End If
            End If


            getIDArray = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt zu der angegebenen Menge von Namen bzw. nameIDs die zugehörigen roleIDs als Array zurück 
    ''' </summary>
    ''' <param name="nameCollection"></param>
    ''' <returns></returns>
    Public ReadOnly Property getIDArray(ByVal nameCollection As Collection) As Integer()
        Get
            Dim tmpResult() As Integer = Nothing

            Dim tmpCollection As New Collection

            For Each nameID As String In nameCollection
                Dim teamID As Integer = -1
                Dim roleID As Integer = parseRoleNameID(nameID, teamID)
                If roleID > 0 Then
                    tmpCollection.Add(roleID)
                End If
            Next

            If tmpCollection.Count > 0 Then
                ReDim tmpResult(tmpCollection.Count - 1)
                Dim ix As Integer = 0
                For Each roleID As Integer In tmpCollection
                    tmpResult(ix) = roleID
                    ix = ix + 1
                Next
            End If

            getIDArray = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt die Rolle zurück, die ein Eltern-/GroßElternteil der angegebenen Rolel ist 
    ''' </summary>
    ''' <param name="roleNameID"></param>
    ''' <param name="summaryRoleIDs"></param>
    ''' <returns></returns>
    Public Function chooseParentFromList(ByVal roleNameID As String, ByVal summaryRoleIDs() As Integer, ByVal includingVirtualChilds As Boolean) As String
        Dim tmpResult As String = ""

        Dim teamID As Integer = -1

        Dim roleID As Integer = Me.parseRoleNameID(roleNameID, teamID)
        roleNameID = Me.bestimmeRoleNameID(roleID, teamID)

        If summaryRoleIDs.Contains(roleID) Then
            tmpResult = Me.getRoleDefByID(roleID).name

        Else
            For Each summaryRoleID As Integer In summaryRoleIDs
                Dim chckRelation As Boolean = hasAnyChildParentRelationsship(roleNameID, summaryRoleID, includingVirtualChilds)
                If chckRelation = True Then
                    tmpResult = Me.getRoleDefByID(summaryRoleID).name
                    Exit For
                End If
            Next

        End If

        chooseParentFromList = tmpResult
    End Function

    ''' <summary>
    ''' überprüft, ob die angegebene roleNameID in der Form roleID;teamID bzw roleId Kind einer der angegebenen Sammelrollen ist
    ''' wenn summaryRoleIDS = Nothing und roleNAmeID tatsächlich existiert, dann true
    ''' </summary>
    ''' <param name="roleNameID"></param>
    ''' <param name="summaryRoleIDs"></param>
    ''' <returns></returns>
    Public Function hasAnyChildParentRelationsship(ByVal roleNameID As String, ByVal summaryRoleIDs() As Integer,
                                                   Optional includingVirtualChilds As Boolean = False) As Boolean

        Dim tmpResult As Boolean = False
        Dim teamID As Integer = -1

        ' tk 1.6.20 , wenn das mit Nothing aufgerufen wird, dann ist das true 
        If IsNothing(summaryRoleIDs) Then
            tmpResult = RoleDefinitions.containsNameOrID(roleNameID)
        Else
            Dim roleID As Integer = Me.parseRoleNameID(roleNameID, teamID)
            If summaryRoleIDs.Contains(roleID) Then
                tmpResult = True

            Else
                For Each summaryRoleID As Integer In summaryRoleIDs
                    tmpResult = hasAnyChildParentRelationsship(roleNameID, summaryRoleID, includingVirtualChilds = includingVirtualChilds)
                    If tmpResult = True Then
                        Exit For
                    End If
                Next

            End If
        End If


        hasAnyChildParentRelationsship = tmpResult
    End Function

    ''' <summary>
    ''' checks, whether two parentIDs do have common childs. Example: parent1=teamID, parent2=orga-unit     '''
    ''' </summary>
    ''' <param name="parentID1">for example Teamid</param>
    ''' <param name="parentID2">for example orga-ID</param>
    ''' <returns>empty list, when no common childs</returns>
    Public Function getCommonChildsOfParents(ByVal parentID1 As Integer, parentID2 As Integer) As List(Of Integer)
        Dim returnResult As New List(Of Integer)

        ' es dürfen keine virtualChilds berücksichtigt werden .. 
        Dim allChilds1 As SortedList(Of Integer, Double) = getSubRoleIDsOf(getRoleDefByID(parentID1).name)
        Dim allChilds2 As SortedList(Of Integer, Double) = getSubRoleIDsOf(getRoleDefByID(parentID2).name)
        Dim smallerList As List(Of Integer)
        Dim biggerList As List(Of Integer)

        If allChilds1.Count < allChilds1.Count Then
            smallerList = allChilds1.Keys.ToList
            biggerList = allChilds2.Keys.ToList
        Else
            smallerList = allChilds2.Keys.ToList
            biggerList = allChilds1.Keys.ToList
        End If

        ' jetzt wird verglichen ... 
        For Each commonID As Integer In smallerList
            If biggerList.Contains(commonID) Then
                ' commonID muss nicht auf Vorhandensein gepräft werden, da es aus einer sortierten Liste kommt 
                returnResult.Add(commonID)
            End If
        Next

        getCommonChildsOfParents = returnResult
    End Function

    ''' <summary>
    ''' determines whether or not roleNameID is child/child-of-child/.. of potential parent summaryRoleID 
    ''' when includingVirtualChilds = true: considers a team, when completely contained in summaryRoleID also as child 
    ''' completely contained in parentID means: all members of team are childs of given parentID 
    ''' </summary>
    ''' <param name="roleNameID"></param>
    ''' <param name="summaryRoleID"></param>
    ''' <returns></returns>
    Public Function hasAnyChildParentRelationsship(ByVal roleNameID As String, ByVal summaryRoleID As Integer,
                                                   Optional includingVirtualChilds As Boolean = False) As Boolean
        Dim tmpResult As Boolean = False
        Dim skillID As Integer = -1

        If Not IsNothing(roleNameID) And Not IsNothing(summaryRoleID) Then
            If roleNameID <> "" And summaryRoleID > 0 Then

                Dim roleID As Integer = Me.parseRoleNameID(roleNameID, skillID)
                roleNameID = Me.bestimmeRoleNameID(roleID, skillID)

                ' now determine whether summaryRoleID is Skill or Role 
                Dim summaryRole As clsRollenDefinition = getRoleDefByID(summaryRoleID)
                Dim curRole As clsRollenDefinition = getRoleDefByID(roleID)

                Dim childIDs As SortedList(Of Integer, Double) = Me.getSubRoleIDsOf(summaryRole.name)
                If summaryRole.isSkill And skillID > 0 Then
                    ' Summary Role ist Skill 
                    tmpResult = childIDs.ContainsKey(skillID)

                ElseIf curRole.isCombinedRole Then
                    ' Summary Role ist Orga-Unit, gesuchte Rolle ist SammelRolle
                    tmpResult = getCommonChildsOfParents(roleID, summaryRoleID).Count > 0
                Else
                    ' Sumary Rolle ist Orga-Unit, gesuchte Rolle ist Person
                    tmpResult = childIDs.ContainsKey(roleID)
                End If
            End If

        End If


        hasAnyChildParentRelationsship = tmpResult
    End Function

    ''' <summary>
    ''' determines whether or not roleID is child/child-of-child/.. of potential parent summaryRoleID 
    ''' no consideration of virtualchilds
    ''' </summary>
    ''' <param name="roleID"></param>
    ''' <param name="summaryRoleID"></param>
    ''' <returns></returns>
    Public Function hasAnyChildParentRelationsship(ByVal roleID As Integer, ByVal summaryRoleID As Integer) As Boolean

        Dim tmpResult As Boolean = False

        If roleID = summaryRoleID Then
            tmpResult = True
        Else
            Dim sRole As clsRollenDefinition = Me.getRoleDefByID(summaryRoleID)
            If Not IsNothing(sRole) Then
                Dim alleChildIDs As SortedList(Of Integer, Double) = Me.getSubRoleIDsOf(sRole.name, type:=PTcbr.all)
                If alleChildIDs.Count > 0 Then
                    tmpResult = alleChildIDs.ContainsKey(roleID)
                End If
            End If
        End If


        hasAnyChildParentRelationsship = tmpResult

    End Function

    ''' <summary>
    ''' True, if at least one of the roleNAmeIDs is child / child-child/ .. of summaryRoleID;
    ''' Input is sortedList of roleNameIDs roleUId;teamUid, summaryRoleID is ID of potential parentID    ''' 
    ''' no consideration of virtualChilds  
    ''' </summary>
    ''' <param name="roleIDs"></param>
    ''' <param name="summaryRoleID"></param>
    ''' <returns></returns>
    Public Function hasAnyChildParentRelationsship(ByVal roleIDs As SortedList(Of String, String), ByVal summaryRoleID As Integer) As Boolean
        Dim tmpResult As Boolean = False

        For Each kvp As KeyValuePair(Of String, String) In roleIDs
            Dim teamID As Integer = -1
            Dim sRole As clsRollenDefinition = Me.getRoleDefByIDKennung(kvp.Value, teamID)

            If hasAnyChildParentRelationsship(sRole.UID, summaryRoleID) Then
                tmpResult = True
                Exit For
            End If

        Next

        hasAnyChildParentRelationsship = tmpResult

    End Function

    ''' <summary>
    ''' True, if at least one of the roleNAmeIDs Is child / child-child/ .. of summaryRoleID;
    ''' Input is sortedList of roleNameIDs roleUId;teamUid, summaryRoleID is ID of potential parentID    ''' 
    ''' no consideration of virtualChilds  
    ''' </summary>
    ''' <param name="idArray"></param>
    ''' <param name="summaryRoleID"></param>
    ''' <returns></returns>
    Public Function hasAnyChildParentRelationsship(ByVal idArray As String(), ByVal summaryRoleID As Integer) As Boolean
        Dim tmpResult As Boolean = False

        For Each nameID As String In idArray
            Dim teamID As Integer = -1
            Dim sRole As clsRollenDefinition = Me.getRoleDefByIDKennung(nameID, teamID)

            If hasAnyChildParentRelationsship(sRole.UID, summaryRoleID) Then
                tmpResult = True
                Exit For
            End If

        Next

        hasAnyChildParentRelationsship = tmpResult
    End Function


    ''' <summary>
    ''' True if child with roleName is child of at least on of the IDs in collection
    ''' or if child-ID is in collection itself  
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <param name="roleCostCollection"></param>
    ''' <returns></returns>
    Public Function hasAnyChildParentRelationsship(ByVal roleName As String, ByVal roleCostCollection As Collection) As Boolean

        Dim tmpRole As clsRollenDefinition = Me.getRoledef(roleName)
        Dim isRole As Boolean = Me.containsName(roleName)

        Dim iscost As Boolean = False
        Dim found As Boolean = False
        Dim ix As Integer = 1
        Dim myIDs As SortedList(Of Integer, Double)

        If isRole Then

            ' ist es eine Gruppe ...
            If tmpRole.isSkill Then
                myIDs = Me.getSubRoleIDsOf(roleName)
            Else
                myIDs = New SortedList(Of Integer, Double)
                Dim myUID As Integer = Me.getRoledef(roleName).UID
                myIDs.Add(myUID, 1.0)
            End If

            If roleCostCollection.Contains(roleName) Then
                found = True
            Else
                Do While Not found And ix <= roleCostCollection.Count

                    Dim parentName As String = CStr(roleCostCollection.Item(ix))

                    If Me.containsName(parentName) Then
                        Dim myUID As Integer = Me.getRoledef(roleName).UID
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
        Dim isTeamMember As Boolean = (teamID > 0)

        If isTeamMember Then
            isTeamMember = _allRollen.ContainsKey(teamID)
        End If

        If _allRollen.ContainsKey(roleUID) Then
            Dim nodeName As String = roleUID.ToString

            ' stellt sicher dass in einer sortierten Liste mit roleNameIDs alle Rollen mit der gleichen roleUID beieinander stehen  
            If isTeamMember Then
                nodeName = roleUID.ToString & ";" & teamID.ToString
            Else
                nodeName = roleUID.ToString & ";"
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
    ''' <param name="skillName"></param>
    ''' <returns></returns>
    Public Function bestimmeRoleNameID(ByVal roleName As String, ByVal skillName As String) As String
        Dim tmpResult As String = ""
        Dim tmpRole As clsRollenDefinition = Me.getRoledef(roleName)

        Try
            If Not IsNothing(tmpRole) Then

                If skillName.Length > 0 Then
                    Dim tmpSkill As clsRollenDefinition = Me.getRoledef(skillName)
                    If Not IsNothing(tmpSkill) Then

                        If Me.getCommonChildsOfParents(tmpRole.UID, tmpSkill.UID).Count > 0 Then
                            tmpResult = Me.bestimmeRoleNameID(tmpRole.UID, tmpSkill.UID)
                        Else
                            Dim dummy As Integer = -1
                            tmpResult = Me.bestimmeRoleNameID(tmpRole.UID, dummy)
                        End If
                        ' tk 23.8.20 alte Variante, hat nur zugelassen, dass Personen Skill-Spezifikation haben können 
                        'If tmpSkill.getSubRoleIDs.ContainsKey(tmpRole.UID) Then
                        '    tmpResult = Me.bestimmeRoleNameID(tmpRole.UID, tmpSkill.UID)
                        'Else
                        '    Dim dummy As Integer = -1
                        '    tmpResult = Me.bestimmeRoleNameID(tmpRole.UID, dummy)
                        'End If
                    Else
                        tmpResult = Me.bestimmeRoleNameID(tmpRole.UID, -1)
                    End If
                Else
                    tmpResult = Me.bestimmeRoleNameID(tmpRole.UID, -1)
                End If

            End If
        Catch ex As Exception
            tmpResult = ""
        End Try

        bestimmeRoleNameID = tmpResult
    End Function



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
            Dim initialRole As clsRollenDefinition = Me.getRoledef(roleName)



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
                                    sammelRollenCollection.Add(kvp.Key, 1.0)

                                    Dim listofSubRoles As SortedList(Of Integer, Double) = roleDef.getSubRoleIDs

                                    If Not IsNothing(listofSubRoles) Then

                                        For Each srkvp As KeyValuePair(Of Integer, Double) In listofSubRoles

                                            ' 
                                            If Not realCollection.ContainsKey(srkvp.Key) And Not addToRealCollection.ContainsKey(srkvp.Key) Then
                                                addToRealCollection.Add(srkvp.Key, 1.0)
                                                ' tk 18.10.20
                                                'If Not initialRoleISSkill Then
                                                '    ' do it anyway in case initial role was Non-Skill, because then leafs are also non-skills
                                                '    addToRealCollection.Add(srkvp.Key, srkvp.Value)
                                                'Else
                                                '    ' because final leaf of skill is always role: make sure roles are not taken as childs of skills 
                                                '    ' except when askedd for realRoles
                                                '    Dim childRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(srkvp.Key)
                                                '    If childRole.isSkill Then
                                                '        addToRealCollection.Add(srkvp.Key, srkvp.Value)
                                                '    Else
                                                '        ' initial Role was skill, child-Role is no skill: take it when asked for realRoles
                                                '        If type = PTcbr.realRoles Then
                                                '            addToRealCollection.Add(srkvp.Key, srkvp.Value)
                                                '        Else
                                                '            ' in this case all other childs will be non-skills
                                                '            Exit For
                                                '        End If

                                                '    End If
                                                'End If
                                                ' this is not any more needed because there is no caoacity percentage given for skills any more
                                                ' capacity of a skill is always defined by the person having this skill 
                                                'ElseIf addToRealCollection.ContainsKey(srkvp.Key) Then
                                                '    ' addieren, aber Gesamt-Summe darf nie größer 1 sein
                                                '    Dim newValue As Double = addToRealCollection(srkvp.Key) + srkvp.Value
                                                '    If newValue > 1.0 Then
                                                '        newValue = 1.0
                                                '    End If
                                                '    addToRealCollection(srkvp.Key) = newValue
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
                            realCollection.Add(kvp.Key, 1.0)
                            ' tk 18.10 das wird nicht mehr benötigt: keine Angabe der % mehr, wurde bsiher benutzt um Kapa von Skill zu berechnen 
                            'Else
                            '    Dim newValue As Double = realCollection(kvp.Key) + kvp.Value
                            '    If newValue > 1.0 Then
                            '        newValue = 1.0
                            '    End If
                            '    realCollection(kvp.Key) = newValue
                        End If
                    Next

                    addToRealCollection.Clear()

                Loop

                ' jetzt müssen die realCollections ggf noch bereinigt werden: die Namen der Sammelrollen müssen raus

                If type = PTcbr.all Then
                    ' nichts tun - realCollections enthält schon alles - auch includingVirtualChilds ist nicht mehr nötig ... 

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
                        Dim tmpRole As clsRollenDefinition = Me.getRoleDefByIDKennung(exclName, teamID)

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
            ElseIf name <> "" Then
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

    ''' <summary>
    ''' gibt true zurück, wenn es sich bei dem NameID-String roleUID;teamID um eine valide Kombination handelt
    ''' Strongtest=true: gibt nur true zurück, wenn roleUID, TeamUID existieren und RoleUId auch Kind von TeamID ist 
    ''' strongTest=false: gibt true zurück, wenn RoleUID und TeamID existieren  
    ''' </summary>
    ''' <param name="nameID"></param>
    ''' <returns></returns>
    Public ReadOnly Property containsNameOrID(ByVal nameID As String, Optional ByVal strongTest As Boolean = True) As Boolean
        Get
            Dim tmpResult As Boolean = False
            Dim teamID As Integer = -1
            Dim roleUID As Integer = parseRoleNameID(nameID, teamID)

            If _allRollen.ContainsKey(roleUID) Then
                ' ist ein Team angegeben ? 
                If teamID <> -1 Then
                    If _allRollen.ContainsKey(teamID) Then
                        If Not strongTest Then
                            tmpResult = True
                        Else
                            ' ist die RoleUID auch Kind des Teams ? 
                            If getRoleDefByID(teamID).getSubRoleIDs.ContainsKey(roleUID) Then
                                tmpResult = True
                            End If
                        End If

                    End If
                Else
                    tmpResult = True
                End If
            Else
                tmpResult = False
            End If

            containsNameOrID = tmpResult

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
    ''' gibt den Bezeichner zurück
    ''' </summary>
    ''' <param name="nameID"></param>
    ''' <returns></returns>
    Public ReadOnly Property getBezeichner(ByVal nameID As String) As String
        Get
            Dim tmpResult As String = ""
            Dim teamID As Integer
            Dim roleID As Integer = parseRoleNameID(nameID, teamID)
            If teamID <> -1 Then
                tmpResult = _allRollen.Item(roleID).name & " (" & _allRollen.Item(teamID).name & ")"
            Else
                tmpResult = _allRollen.Item(roleID).name
            End If

            getBezeichner = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' returns a list of Role-Names containing 
    ''' </summary>
    ''' <param name="substr"></param>
    ''' <returns></returns>
    Public Function getRoleNamesContainingSubStr(ByVal substr As String, ByVal skillName As String) As List(Of String)
        Dim tmpResult As New SortedList(Of String, Boolean)

        If substr.Length > 0 Then
            For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In _allRollen
                If Not (kvp.Value.isSkill Or kvp.Value.isSkillParent) Then
                    If kvp.Value.name.Contains(substr) Then
                        If Not tmpResult.ContainsKey(kvp.Value.name) Then
                            If skillName = "" Then
                                tmpResult.Add(kvp.Value.name, True)
                            Else
                                Dim tmpSkill As clsRollenDefinition = getRoledef(skillName)
                                If Not IsNothing(tmpSkill) Then
                                    If tmpSkill.isSkill Then
                                        Dim commonList As List(Of Integer) = Me.getCommonChildsOfParents(tmpSkill.UID, kvp.Value.UID)
                                        If commonList.Count > 0 Then
                                            tmpResult.Add(kvp.Value.name, True)
                                        End If
                                    Else
                                        tmpResult.Add(kvp.Value.name, True)
                                    End If

                                Else
                                    tmpResult.Add(kvp.Value.name, True)
                                End If
                            End If

                        End If
                    End If
                End If
            Next
        End If

        getRoleNamesContainingSubStr = tmpResult.Keys.ToList
    End Function

    ''' <summary>
    ''' returns a list of Role-Names containing 
    ''' if roleName is given , only Skills are shown which belong to roleName
    ''' </summary>
    ''' <param name="substr"></param>
    ''' <returns></returns>
    Public Function getSkillNamesContainingSubStr(ByVal substr As String, ByVal roleName As String) As List(Of String)
        Dim tmpResult As New SortedList(Of String, Boolean)

        If substr.Length > 0 Then
            For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In _allRollen
                If kvp.Value.isSkill Or kvp.Value.isSkillParent Then
                    If kvp.Value.name.Contains(substr) Then
                        If Not tmpResult.ContainsKey(kvp.Value.name) Then
                            If roleName = "" Then
                                tmpResult.Add(kvp.Value.name, True)
                            Else
                                Dim tmpRole As clsRollenDefinition = getRoledef(roleName)
                                If Not IsNothing(tmpRole) Then
                                    If Not tmpRole.isSkill Then
                                        Dim commonList As List(Of Integer) = Me.getCommonChildsOfParents(kvp.Value.UID, tmpRole.UID)
                                        If commonList.Count > 0 Then
                                            tmpResult.Add(kvp.Value.name, True)
                                        End If
                                    Else
                                        tmpResult.Add(kvp.Value.name, True)
                                    End If
                                Else
                                    tmpResult.Add(kvp.Value.name, True)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If

        getSkillNamesContainingSubStr = tmpResult.Keys.ToList
    End Function

    ''' <summary>
    ''' bestimmt aus dem übergebenen SelectedRolesItem Angaben wie RoleUID, ggf. die zugehörige TeamID
    ''' </summary>
    ''' <param name="selRoleItem"></param>
    ''' <param name="teamID"></param>
    ''' <returns></returns>
    Public Function parseRoleNameID(ByVal selRoleItem As String, ByRef teamID As Integer) As Integer
        ' der Name wird bestimmt, je nachdem ob es sich um eine normale Orga-Einheit , ein Team oder ein Team-Member handelt 


        Dim roleID As Integer = -1

        ' Vorbesetzung von teamID 
        teamID = -1

        If Not IsNothing(selRoleItem) Then
            If selRoleItem <> "" Then
                Dim tmpStr() As String = selRoleItem.Split(New Char() {CChar(";")})

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
    ''' gibt die RollenDefinition zurück, die zu de rPersonal-Nummer employeeNr gehört
    ''' </summary>
    ''' <param name="employeeNr"></param>
    ''' <returns>Nothing or RoleDefinition</returns>
    Public Function getRoledefByEmployeeNr(ByVal employeeNr As String) As clsRollenDefinition

        Dim result As clsRollenDefinition = Nothing

        For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In _allRollen
            If kvp.Value.employeeNr = employeeNr Then
                result = kvp.Value
                Exit For
            End If
        Next

        getRoledefByEmployeeNr = result
    End Function
    ''' <summary>
    ''' gibt die Rolle zurück, die die gesuchte ID hat ...
    ''' _allRollen ist eine sortierte Liste ..
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
                Dim parentRole As clsRollenDefinition = Me.getRoleDefByID(parentUID)

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
        _positionIndices = New SortedList(Of String, Integer)

        ' wird erst in buildOrgaTeamChilds initialisiert und aufgebaut 
        _orgaSkillChilds = Nothing
        _topLevelSkillParents = Nothing

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

        ' zurücksetzen ... wenn der Portfolio Manager die Gruppen nicht angezeigt bekommen soll 
        If _topLevelNodeIDs.Count > 0 Then
            _topLevelNodeIDs = New List(Of Integer)
        End If

        While (i <= _allRollen.Count)

            ' Level 0 Knoten
            currentRole = _allRollen.ElementAt(i - 1).Value
            Dim parentRole As clsRollenDefinition = Me.getParentRoleOf(currentRole.UID)

            If IsNothing(parentRole) Then
                If Not _topLevelNodeIDs.Contains(currentRole.UID) Then

                    ' aufnehmen als Top Level Node ...
                    ' auch ein Portfolio Manager soll die Skillgruppen sehen können ... 
                    _topLevelNodeIDs.Add(currentRole.UID)

                    ' tk 15.10.20 das muss raus, weil auch ein Portfolio Manager die Teams sehen können soll 
                    'If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                    '    If Not isParentOfAllTeams(currentRole.UID) Then
                    '        _topLevelNodeIDs.Add(currentRole.UID)
                    '    End If
                    'Else
                    '    _topLevelNodeIDs.Add(currentRole.UID)
                    'End If

                End If
            End If

            i = i + 1

        End While

        '
        ' tk 25.7.19
        ' jetzt werden noch die relativen Indices aufgebaut ... 
        Try
            Call setRelativePositionIndicesOfRoles()
        Catch ex As Exception

        End Try

    End Sub


End Class
