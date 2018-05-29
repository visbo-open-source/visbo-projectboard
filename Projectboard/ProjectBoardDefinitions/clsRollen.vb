''' <summary>
''' Die Rollen müssen immer in der customization file in der ursprünglichen Reihenfolge aufgeführt sein; 
''' ein Name kann umbenannt werden , aber er darf auf keinen Fall an eine andere Psoiton verschoben werden 
''' neue Rollen müssen immer ans Ende gestellt werden - alte Rollen müssen immer mitgeschrieben werden ... 
''' </summary>
''' <remarks></remarks>
Public Class clsRollen


    Private _allRollen As SortedList(Of Integer, clsRollenDefinition)
    Private _topLevelNodeIDs As List(Of Integer)

    Public Sub Add(roledef As clsRollenDefinition)

        ' Änderung tk: umgestellt auf 
        If Not _allRollen.ContainsKey(roledef.UID) Then
            _allRollen.Add(roledef.UID, roledef)
        Else
            Throw New ArgumentException(roledef.UID.ToString & " existiert bereits")
        End If

    End Sub

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
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTopLevelNodeIDs() As List(Of Integer)
        Get
            getTopLevelNodeIDs = _topLevelNodeIDs
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

                        Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(exclName)
                        If Not IsNothing(tmpRole) Then
                            If realCollection.ContainsKey(tmpRole.UID) And exclName <> roleName Then
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

    Public Sub New()

        _allRollen = New SortedList(Of Integer, clsRollenDefinition)
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
                ' aufnehmen als Top Level Node ...
                Me._topLevelNodeIDs.Add(currentRole.UID)
            End If

            i = i + 1

        End While

    End Sub


End Class
