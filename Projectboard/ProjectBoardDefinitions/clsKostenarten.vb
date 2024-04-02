Public Class clsKostenarten

    'sortiert nach UID
    Private _allKostenarten As SortedList(Of Integer, clsKostenartDefinition)

    Private _topLevelNodeIDs As List(Of Integer)


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
                    tmpName = _allKostenarten.Item(_topLevelNodeIDs.First).name
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
                For Each costID As Integer In _topLevelNodeIDs
                    Dim subCostList As SortedList(Of Integer, Double) = _allKostenarten.Item(costID).getSubCostIDs()

                    For Each srKvP As KeyValuePair(Of Integer, Double) In subCostList
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



    Public Sub Add(costdef As clsKostenartDefinition)

        If Not IsNothing(costdef) Then
            If Not _allKostenarten.ContainsKey(costdef.UID) Then
                _allKostenarten.Add(costdef.UID, costdef)
            Else
                Throw New ArgumentException(costdef.UID.ToString & " existiert bereits")
            End If
        Else
            Throw New ArgumentException("Kostenart darf nicht Nothing sein")
        End If


        ''Try
        ''    _allKostenarten.Add(Item:=costdef, Key:=costdef.name)
        ''Catch ex As Exception
        ''    Throw New ArgumentException(costdef.name & " existiert bereits")
        ''End Try


    End Sub

    Public ReadOnly Property liste() As SortedList(Of Integer, clsKostenartDefinition)
        Get
            liste = _allKostenarten
        End Get
    End Property

    ''Public Sub Remove(myitem As Object)

    ''    Try
    ''        _allKostenarten.Remove(myitem)
    ''    Catch ex As Exception
    ''        Throw New ArgumentException("Fehler bei Kostenart entfernen")
    ''    End Try


    ''End Sub


    ''' <summary>
    ''' liefert true zurück, wenn alle Kostendefinitionen der einen Liste identisch mit der anderen sind
    ''' </summary>
    ''' <param name="vglDefinitionen"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglDefinitionen As clsKostenarten)
        Get
            Dim stillIdentical = True

            If Me.Count = vglDefinitionen.Count Then
                Dim i As Integer = 0
                Do While i < _allKostenarten.Count And stillIdentical
                    stillIdentical = _allKostenarten.ElementAt(i).Value.isIdenticalTo(vglDefinitionen.getCostdef(i + 1))
                    i = i + 1
                Loop

            Else
                stillIdentical = False
            End If

            isIdenticalTo = stillIdentical
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Count = _allKostenarten.Count
        End Get
    End Property

    ''' <summary>
    ''' prüft, ob name in der Kostenarten Collection enthalten ist 
    ''' </summary>
    ''' <param name="name">typ string</param>
    ''' <value></value>
    ''' <returns>wahr, wenn enthalten; falsch sonst</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsName(name As String) As Boolean
        Get

            Dim found As Boolean = False
            If IsNothing(name) Then
                ' found bleibt auf false
            Else
                Dim ix As Integer = 0
                Do While ix <= _allKostenarten.Count - 1 And Not found
                    If _allKostenarten.ElementAt(ix).Value.name = name Then
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

            containsUid = _allKostenarten.ContainsKey(uid)

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
                For Each costID As Integer In _topLevelNodeIDs
                    If _allKostenarten.ContainsKey(costID) Then
                        Dim tmpName As String = _allKostenarten.Item(costID).name
                        If Not returnResult.Contains(tmpName) Then
                            returnResult.Add(tmpName, tmpName)
                        End If
                    End If
                Next

            ElseIf Level = 1 Then

                For Each costID As Integer In _topLevelNodeIDs

                    Dim subcostList As SortedList(Of Integer, Double) = _allKostenarten.Item(costID).getSubCostIDs()

                    If subcostList.Count > 0 Then
                        For Each srKvP As KeyValuePair(Of Integer, Double) In subcostList
                            If _allKostenarten.ContainsKey(srKvP.Key) Then
                                Dim tmpName As String = _allKostenarten.Item(srKvP.Key).name
                                If Not returnResult.Contains(tmpName) Then
                                    returnResult.Add(tmpName, tmpName)
                                End If
                            End If
                        Next
                    Else
                        If _allKostenarten.ContainsKey(costID) Then
                            Dim tmpName As String = _allKostenarten.Item(costID).name
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
    ''' gibt in einer eindeutigen Liste die Namen aller vorkommenden SubCostIDs in einer sortierten Liste integer, double zurück, das heisst alle Sammler und die realen Kostenarten , oder nur die Sammler oder nur die realen Kostenarten  
    ''' es werden also alle Cost-IDs zurückgegeben, Platzhalter und Basis Kostenarten, oder nur eine Kategorie davon 
    ''' wenn die excludedNames angegeben sind, dann werden nur die Kostenarten aufgenommen, die nicht in den excluded Names drin sind. 
    ''' Das stellt sicher, dass im Falle einer Kosten Auswertung Kosten nicht dopplet gezählt werden, weil sie einmal als Sammler gewertet werden, einmal als explizit angegebene Kostenart 
    ''' 
    ''' das funktioniert auch über mehrstufige Sammler
    ''' </summary>
    ''' <param name="costName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSubCostIDsOf(ByVal costName As String,
                                               Optional ByVal type As Integer = PTcbr.all,
                                               Optional ByVal excludedNames As Collection = Nothing) As SortedList(Of Integer, Double)

        Get

            ' hier muss überprüft werden, ob die myCollection Sammelrollen enthält 
            ' wenn ja, werden die alle solange um die enthaltenen Sammelrollen ergänzt, bis keine Sammelrolle mehr in der Collection drin ist  
            ' die Sammelrollen werden am Schluss wieder aufgenommen, weil sie ja als Platzhalter Rollen ihre Bedarfs-Werte auch mit geben müssen 

            Dim sammlerCollection As New SortedList(Of Integer, Double)
            Dim realCollection As New SortedList(Of Integer, Double)
            Dim addToRealCollection As New SortedList(Of Integer, Double)
            Dim noUntreatedCombinedCost As Boolean = False
            Dim initialCost As clsKostenartDefinition = getCostdef(costName)



            If Not IsNothing(initialCost) Then


                ' initial besetzen, um es in Gang zu setzen
                'realCollection.Add(roleName, roleName)
                realCollection.Add(initialCost.UID, 1.0)

                Do Until noUntreatedCombinedCost

                    noUntreatedCombinedCost = True

                    For Each kvp As KeyValuePair(Of Integer, Double) In realCollection

                        Dim costDef As clsKostenartDefinition = Me.getCostDefByID(kvp.Key)

                        If Not IsNothing(costDef) Then

                            If costDef.isCombinedCost Then

                                If Not sammlerCollection.ContainsKey(kvp.Key) Then

                                    noUntreatedCombinedCost = False
                                    ' dann wurde sie nicht schon mal ersetzt  und die Kinder müssen aufgenommen werden  
                                    sammlerCollection.Add(kvp.Key, 1.0)

                                    Dim listofSubCosts As SortedList(Of Integer, Double) = costDef.getSubCostIDs

                                    If Not IsNothing(listofSubCosts) Then

                                        For Each sckvp As KeyValuePair(Of Integer, Double) In listofSubCosts

                                            ' 
                                            If Not realCollection.ContainsKey(sckvp.Key) And Not addToRealCollection.ContainsKey(sckvp.Key) Then
                                                addToRealCollection.Add(sckvp.Key, 1.0)
                                            End If


                                        Next

                                    Else
                                        ' darf eigentlich nicht sein , aber ist im Fehlerfall notwenig, um Endlos schleife zu verhindern 
                                        noUntreatedCombinedCost = True
                                    End If

                                End If

                            End If
                        End If


                    Next

                    ' jetzt müssen die addToRealCollection Items übertragen werden 
                    For Each kvp As KeyValuePair(Of Integer, Double) In addToRealCollection
                        If Not realCollection.ContainsKey(kvp.Key) Then
                            realCollection.Add(kvp.Key, 1.0)
                        End If
                    Next

                    addToRealCollection.Clear()

                Loop

                ' jetzt müssen die realCollections ggf noch bereinigt werden: die Namen der Sammelrollen müssen raus

                If type = PTcbr.all Then
                    ' nichts tun - realCollections enthält schon alles - auch includingVirtualChilds ist nicht mehr nötig ... 

                ElseIf type = PTcbr.placeholders Then
                    realCollection = sammlerCollection


                ElseIf type = PTcbr.realRoles Then
                    For Each cRKvp As KeyValuePair(Of Integer, Double) In sammlerCollection
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

                        Dim tmpCost As clsKostenartDefinition = Me.getCostdef(exclName)

                        If Not IsNothing(tmpCost) Then
                            If realCollection.ContainsKey(tmpCost.UID) And tmpCost.name <> costName Then
                                realCollection.Remove(tmpCost.UID)
                            End If
                        End If

                    Next
                End If
            End If


            getSubCostIDsOf = realCollection


        End Get
    End Property


    ''' <summary>
    ''' returns a list of cost-Names having 'substr' in the Name
    ''' </summary>
    ''' <param name="substr"></param>
    ''' <returns></returns>
    Public Function getCostNamesContainingSubStr(ByVal substr As String) As List(Of String)
        Dim tmpResult As New SortedList(Of String, Boolean)

        If substr.Length > 0 Then
            For Each kvp As KeyValuePair(Of Integer, clsKostenartDefinition) In _allKostenarten

                If kvp.Value.name.Contains(substr) Then
                    If Not tmpResult.ContainsKey(kvp.Value.name) Then
                        tmpResult.Add(kvp.Value.name, True)
                    End If
                End If
            Next
        End If

        getCostNamesContainingSubStr = tmpResult.Keys.ToList
    End Function


    ''' <summary>
    ''' gibt eine Collection zurück, die nur die Kostenarten enthält , die Sammler Kosten sind
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSummaryCosts(Optional ByVal costName As String = Nothing) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim removeList As New Collection

            For r As Integer = 1 To _allKostenarten.Count
                Dim tmpCost As clsKostenartDefinition = _allKostenarten.ElementAt(r - 1).Value
                If tmpCost.isCombinedCost Then
                    If IsNothing(costName) Then
                        tmpCollection.Add(tmpCost.name, tmpCost.name)
                    ElseIf tmpCost.name <> costName Then
                        tmpCollection.Add(tmpCost.name, tmpCost.name)
                    End If
                End If
            Next

            If Not IsNothing(costName) Then

                Dim initialCost As clsKostenartDefinition = Me.getCostdef(costName)
                If Not IsNothing(initialCost) Then

                    For sr As Integer = 1 To tmpCollection.Count
                        Dim tmpCost As clsKostenartDefinition = getCostdef(CStr(tmpCollection.Item(sr)))
                        Dim subCostIDs As SortedList(Of Integer, Double) = getSubCostIDsOf(tmpCost.name, PTcbr.all)

                        If Not subCostIDs.ContainsKey(initialCost.UID) Then
                            removeList.Add(tmpCost.name, tmpCost.name)
                        End If
                    Next

                    For rm As Integer = 1 To removeList.Count
                        tmpCollection.Remove(CStr(removeList.Item(rm)))
                    Next

                End If


            End If

            getSummaryCosts = tmpCollection

        End Get
    End Property


    ''' <summary>
    ''' gibt zu der angegebenen Kostenart den Sammler zurück, die die Kostenart als direkte Sub-Cost enthält 
    ''' leerer String, wenn kein Sammler existiert, der die angegebene Kostenart enthält  
    ''' </summary>
    ''' <param name="costUID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getParentCostOf(ByVal costUID As Integer) As clsKostenartDefinition
        Get

            Dim sammlerCost As Collection = Me.getSummaryCosts
            Dim found As Boolean = False
            Dim i As Integer = 1
            Dim parentCost As clsKostenartDefinition = Nothing

            If _allKostenarten.ContainsKey(costUID) Then
                While Not found And i <= sammlerCost.Count

                    Dim tmpCost As clsKostenartDefinition = getCostdef(CStr(sammlerCost.Item(i)))
                    If Not IsNothing(tmpCost) Then
                        Dim subCostIDs As SortedList(Of Integer, Double) = tmpCost.getSubCostIDs
                        If subCostIDs.ContainsKey(costUID) Then
                            found = True
                            parentCost = tmpCost
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

            getParentCostOf = parentCost

        End Get
    End Property


    ''' <summary>
    ''' gibt die Kostenart mit ID = uid zurück; Nothing, wenn sie nicht existiert
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <returns></returns>
    Public ReadOnly Property getCostDefByID(ByVal uid As Integer) As clsKostenartDefinition
        Get
            If _allKostenarten.ContainsKey(uid) Then
                getCostDefByID = _allKostenarten.Item(uid)
            Else
                getCostDefByID = Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property getCostdef(ByVal myitem As String) As clsKostenartDefinition
        Get

            Dim tmpValue As clsKostenartDefinition = Nothing

            Dim found As Boolean = False
            Dim ix As Integer = 0

            Do While ix <= _allKostenarten.Count - 1 And Not found
                If _allKostenarten.ElementAt(ix).Value.name = myitem Then
                    found = True
                    tmpValue = _allKostenarten.ElementAt(ix).Value
                Else
                    ix = ix + 1
                End If
            Loop

            getCostdef = tmpValue

        End Get
    End Property

    Public ReadOnly Property getCostdef(ByVal myitem As Integer) As clsKostenartDefinition
        Get


            If myitem > 0 And myitem <= _allKostenarten.Count Then
                getCostdef = _allKostenarten.ElementAt(myitem - 1).Value
            Else
                getCostdef = Nothing
            End If


        End Get
    End Property

    ''' <summary>
    ''' baut die Hierarchie der Kostenarten auf; dabei muss hier nur der bzw. die Top Nodes aufgenommen werden 
    ''' in den clsCostNode Elementen sind bereits die Kinder verzeichnet 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub buildTopNodes()
        ' TopKnoten aufbauen
        Dim i As Integer = 1
        Dim currentCost As clsKostenartDefinition
        Dim hparent As New clsKostenartDefinition

        ' zurücksetzen ... wenn der Portfolio Manager die Gruppen nicht angezeigt bekommen soll 
        If _topLevelNodeIDs.Count > 0 Then
            _topLevelNodeIDs = New List(Of Integer)
        End If

        While (i <= _allKostenarten.Count)

            ' Level 0 Knoten
            currentCost = _allKostenarten.ElementAt(i - 1).Value
            Dim parentCost As clsKostenartDefinition = Me.getParentCostOf(currentCost.UID)

            If IsNothing(parentCost) Then
                If Not _topLevelNodeIDs.Contains(currentCost.UID) Then

                    ' aufnehmen als Top Level Node ...
                    ' auch ein Portfolio Manager soll die Skillgruppen sehen können ... 
                    _topLevelNodeIDs.Add(currentCost.UID)


                End If
            End If

            i = i + 1

        End While

        ' tk 27.3.24 wird hier nicht beötigt ... 
        ' tk 11.1.21
        ' das ist notwendig, um die Relative Positions richtig zu haben ...
        'Call buildOrgaSkillChilds()

        ''
        '' tk 25.7.19
        '' jetzt werden noch die relativen Indices aufgebaut ... 
        'Try
        '    Call setRelativePositionIndicesOfRoles()
        'Catch ex As Exception

        'End Try
        ' Ende tk 27.3.24

    End Sub


    Public Sub New()
        _allKostenarten = New SortedList(Of Integer, clsKostenartDefinition)
        _topLevelNodeIDs = New List(Of Integer)
    End Sub

End Class
