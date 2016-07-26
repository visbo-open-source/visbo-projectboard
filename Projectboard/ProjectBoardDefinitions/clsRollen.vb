''' <summary>
''' Die Rollen müssen immer in der customization file in der ursprünglichen Reihenfolge aufgeführt sein; 
''' ein Name kann umbenannt werden , aber er darf auf keinen Fall an eine andere Psoiton verchoben werden 
''' neue Rolle müssen immer ans Ende gestellt werden - alte Rollen müssen immer mitgeschrieben werden ... 
''' </summary>
''' <remarks></remarks>
Public Class clsRollen


    Private _allRollen As SortedList(Of Integer, clsRollenDefinition)



    Public Sub Add(roledef As clsRollenDefinition)

        ' Änderung tk: umgestellt auf 
        If Not _allRollen.ContainsKey(roledef.UID) Then
            _allRollen.Add(roledef.UID, roledef)
        Else
            Throw New ArgumentException(roledef.UID.ToString & " existiert bereits")
        End If



    End Sub

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
                Dim subRoleList As Collection = Me.getSubRoleNamesOf(roleName:=sammelRolle, _
                                                                     type:=PTcbr.all)
                For Each subRole As String In subRoleList
                    If ((tmpCollection.Contains(CStr(subRole))) And (subRole <> sammelRolle)) Then
                        tmpCollection.Remove(CStr(subRole))
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

                For sr As Integer = 1 To tmpCollection.Count
                    Dim tmpRole As clsRollenDefinition = Me.getRoledef(CStr(tmpCollection.Item(sr)))
                    Dim subRoleNames As Collection = Me.getSubRoleNamesOf(tmpRole.name, PTcbr.all)

                    If Not subRoleNames.Contains(roleName) Then
                        removeList.Add(tmpRole.name, tmpRole.name)
                    End If
                Next

                For rm As Integer = 1 To removeList.Count
                    tmpCollection.Remove(CStr(removeList.Item(rm)))
                Next

            End If

            getSummaryRoles = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt in einer eindeutigen Liste die Namen aller vorkommenden SubRoles in einer Collection zurück, das heisst alle Platzhalter und die realen Rollen , oder nur die Platzhalter oder nur die realen Rollen  
    ''' es werden also alle Rollen-Namen zurückgegeben, Platzhalter und reale Rollen-Namen, oder nur eine Kategorie davon 
    ''' wenn die excludedNames angegeben sind, dann werden nur die Rollen aufgenommen, die nicht in den excluded Names drin sind. 
    ''' Das stellt sicher, dass im Falle einer Ressourcen Auswertung Rollen nicht dopplet gezählt werden, weil sie einmal als Sammerolle gewertet werden, einmal als explizit angegebene Rolle 
    ''' 
    ''' das funktioniert auch über mehrstufige Sammelrollen, also wenn Fig2 FIG22, FIG23, enthält, die wiederum Engineering enthalten, die wiederum Namen enthalten
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSubRoleNamesOf(ByVal roleName As String, _
                                               Optional ByVal type As Integer = PTcbr.all, _
                                               Optional ByVal excludedNames As Collection = Nothing) As Collection

        Get

            ' hier muss überprüft werden, ob die myCollection Sammelrollen enthält 
            ' wenn ja, werden die alle solange um die enthaltenen Sammelrollen ergänzt, bis keine Sammelrolle mehr in der Collection drin ist  
            ' die Sammelrollen werden am Schluss wieder aufgenommen, weil sie ja als Platzhalter Rollen ihre Bedarfs-Werte auch mit geben müssen 

            Dim sammelRollenCollection As New Collection
            Dim realCollection As New Collection
            Dim addToRealCollection As New Collection
            Dim noUntreatedCombinedRole As Boolean = False

            ' initial besetzen, um es in Gang zu setzen
            realCollection.Add(roleName, roleName)

            Do Until noUntreatedCombinedRole

                noUntreatedCombinedRole = True

                For Each tmpRole As String In realCollection

                    If RoleDefinitions.containsName(tmpRole) Then


                        Dim roleDef As clsRollenDefinition = Me.getRoledef(tmpRole)

                        If roleDef.isCombinedRole Then


                            If Not sammelRollenCollection.Contains(tmpRole) Then

                                noUntreatedCombinedRole = False
                                ' dann wurde sie nicht schon mal ersetzt  und die Kinder müssen aufgenommen werden  
                                sammelRollenCollection.Add(tmpRole, tmpRole)

                                Dim listofSubRoles As SortedList(Of Integer, String) = roleDef.getSubRoleIDs

                                If Not IsNothing(listofSubRoles) Then
                                    For Each kvp As KeyValuePair(Of Integer, String) In listofSubRoles

                                        Dim subRole As String
                                        If kvp.Key >= 1 And kvp.Key <= Me.Count Then
                                            subRole = Me.getRoledef(kvp.Key).name
                                            If Not realCollection.Contains(subRole) And Not addToRealCollection.Contains(subRole) Then
                                                addToRealCollection.Add(subRole, subRole)
                                            End If
                                        End If

                                    Next
                                End If

                            End If

                        End If
                    End If
                Next

                ' jetzt müssen die addToRealCollection Items übertragen werden 
                For Each tmpItem As String In addToRealCollection
                    If Not realCollection.Contains(tmpItem) Then
                        realCollection.Add(tmpItem, tmpItem)
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
                For Each combinedRole As String In sammelRollenCollection
                    If realCollection.Contains(combinedRole) Then
                        realCollection.Remove(combinedRole)
                    End If
                Next

            Else
                ' nichts tun - realCollection enthält alles  
            End If


            If Not IsNothing(excludedNames) Then
                ' jetzt müssen aus realCollection alle Namen raus, die in excludedNames drin sind ... 
                For Each exclName As String In excludedNames
                    If realCollection.Contains(exclName) And exclName <> roleName Then
                        realCollection.Remove(exclName)
                    End If
                Next
            End If

            getSubRoleNamesOf = realCollection


            ' '' ------ alt , Änderung tk am 10.616 
            ' ''Dim tmpCollection As New Collection
            ' ''Dim tmpRole As clsRollenDefinition = Me.getRoledef(roleName)
            ' ''If Not IsNothing(tmpRole) Then

            ' ''    Dim listOfSubRoles As SortedList(Of Integer, String) = tmpRole.getSubRoleIDs

            ' ''    If Not IsNothing(listOfSubRoles) Then
            ' ''        Dim anzSubroles As Integer = listOfSubRoles.Count

            ' ''        If anzSubroles > 0 Then
            ' ''            For i As Integer = 1 To anzSubroles
            ' ''                Dim subRoleName As String = listOfSubRoles.ElementAt(i - 1).Value
            ' ''                If subRoleName <> roleName And Not tmpCollection.Contains(subRoleName) Then
            ' ''                    tmpCollection.Add(subRoleName, subRoleName)
            ' ''                End If
            ' ''            Next
            ' ''        End If
            ' ''    Else
            ' ''        ' nichts tun
            ' ''    End If

            ' ''End If

            ' ''getSubRoleNamesOf = tmpCollection

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
                Dim ix As Integer = 0
                Do While ix <= _allRollen.Count - 1 And Not found
                    If _allRollen.ElementAt(ix).Value.name = name Then
                        found = True
                    Else
                        ix = ix + 1
                    End If
                Loop
            End If
            
            containsName = found
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

            Dim found As Boolean = False
            Dim ix As Integer = 0

            Do While ix <= _allRollen.Count - 1 And Not found
                If _allRollen.ElementAt(ix).Value.name = myitem Then
                    found = True
                    tmpValue = _allRollen.ElementAt(ix).Value
                Else
                    ix = ix + 1
                End If
            Loop

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

    Public Sub New()

        _allRollen = New SortedList(Of Integer, clsRollenDefinition)

    End Sub

End Class
