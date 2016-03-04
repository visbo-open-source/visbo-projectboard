Public Class clsHierarchy
    Private _allNodes As SortedList(Of String, clsHierarchyNode)

    ''' <summary>
    ''' gibt die eindeutigen Element-Namen oder Element-IDs der Kinder zurück , abhängig von Kennung werden 
    ''' Meilenstein-, Phasen- oder alle Kinder zurückgegeben  
    ''' provideIDs = true: ElemIDs der Kinder 
    ''' provideIDs = false: ElemNames der Kinder  
    ''' ACHTUNG: wenn beides zurückgegeben wird und nur die Element-Namen, kann es sein, dass 
    ''' </summary>
    ''' <param name="elemID">Element-ID , dessen Kinder gesucht werden </param>
    ''' <param name="lookingForMilestones">true: es werden die Namen der Meilenstein Kinder zurückgegeben 
    ''' false: es werden die NAmen der Phasen-Kinder zurückgegeben</param>
    ''' <returns>Collection mit den sortierten Namen der jeweiligen Kinder</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getChildNamesOf(ByVal elemID As String, ByVal lookingForMilestones As Boolean) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim currentNode As clsHierarchyNode = _allNodes.Item(elemID)
            Dim currentChildID As String
            Dim isMilestone As Boolean
            Dim tmpName As String

            If _allNodes.ContainsKey(elemID) Then
                currentNode = _allNodes.Item(elemID)
                If Not IsNothing(currentNode) Then

                    For i As Integer = 1 To currentNode.childCount

                        currentChildID = currentNode.getChild(i)
                        isMilestone = elemIDIstMeilenstein(currentChildID)


                        If lookingForMilestones Then
                            If isMilestone Then
                                tmpName = elemNameOfElemID(currentNode.getChild(i))
                                If Not tmpCollection.Contains(tmpName) Then
                                    tmpCollection.Add(tmpName, tmpName)
                                End If

                            End If
                        Else
                            If Not isMilestone Then
                                tmpName = elemNameOfElemID(currentNode.getChild(i))
                                If Not tmpCollection.Contains(tmpName) Then
                                    tmpCollection.Add(tmpName, tmpName)
                                End If
                            End If
                        End If
                    Next

                End If
            Else
                ' nichts tun, leere Collection zurück geben 
            End If

            getChildNamesOf = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste der IDs der Kinder des Elements zurück , die Phasen bzw. Meilensteine sind
    ''' je nachdem, wie lookingforMilestones gesetzt ist ; Liste ist in der Reihenfolge des Auftretens der Kinder 
    ''' </summary>
    ''' <param name="elemID">ID des aktuellen Elements, dessen Kinder gesucht werden sollen </param>
    ''' <param name="lookingForMilestones">0: Phasen gesucht ; 1: Meilensteine gesucht </param>
    ''' <value></value>
    ''' <returns>nach ID sortierte Collection</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getChildIDsOf(ByVal elemID As String, ByVal lookingForMilestones As Boolean) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim currentNode As clsHierarchyNode = _allNodes.Item(elemID)
            Dim currentChildID As String

            If _allNodes.ContainsKey(elemID) Then
                currentNode = _allNodes.Item(elemID)

                If Not IsNothing(currentNode) Then

                    For i As Integer = 1 To currentNode.childCount

                        currentChildID = currentNode.getChild(i)

                        If lookingForMilestones Then
                            If elemIDIstMeilenstein(currentChildID) Then

                                tmpCollection.Add(currentChildID)

                            End If
                        Else
                            If Not elemIDIstMeilenstein(currentChildID) Then

                                tmpCollection.Add(currentChildID)

                            End If
                        End If
                    Next

                End If
            Else
                ' nichts tun, leere Collection zurück geben 
            End If

            getChildIDsOf = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' fügt der Hierarchy einen Knoten hinzu
    ''' </summary>
    ''' <param name="elemNode"></param>
    ''' <remarks></remarks>
    ''' 
    Public Sub addNode(ByRef elemNode As clsHierarchyNode, ByVal elemKey As String)

        Dim parentNode As clsHierarchyNode
        Dim parentNodeKey As String = elemNode.parentNodeKey

        ' wenn der elemKey bereits existiert, so soll nichts gemacht werden 
        ' das ist dann der Fall, wenn ein Projekt durch Kopie aus einem anderen entsteht

        If _allNodes.ContainsKey(elemKey) Then
            ' nichts tun 
        Else

            ' jetzt wird der Parent-Node bestimmt , sofern er existiert 
            If parentNodeKey.Length > 0 Then
                If _allNodes.ContainsKey(parentNodeKey) Then
                    parentNode = _allNodes.Item(parentNodeKey)
                Else
                    Throw New Exception(parentNodeKey & " existiert nicht ")
                End If
            Else
                parentNode = Nothing
            End If

            ' jetzt den Parent Key eintragen 
            elemNode.parentNodeKey = parentNodeKey
            _allNodes.Add(elemKey, elemNode)

            ' jetzt das Child im Parent Node verankern 
            If Not IsNothing(parentNode) Then
                parentNode.addChild(elemKey)
            End If


        End If




    End Sub

    ''' <summary>
    ''' löscht das Element , das durch uniqueID gekennzeichnet ist, in der Hierarchie-Liste 
    ''' der Parent-Node und die Kind-Nodes werden entsprechend konsistent gehalten 
    ''' Exception, wenn rootphase gelöscht werden soll  
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <param name="deleteALLChilds">true: alle Kinder werden rekursiv gelöscht 
    ''' false: alle Kinder bekommen den "Großvater" als Vater</param>
    ''' <remarks></remarks>
    Public Sub removeNode(ByVal uniqueID As String, Optional ByVal deleteALLChilds As Boolean = True)

        If uniqueID = rootPhaseName Then
            Throw New ArgumentException(message:="Root Phase kann nicht gelöscht werden ", paramName:=uniqueID)
        End If



        If _allNodes.ContainsKey(uniqueID) Then
            Dim elemNode As clsHierarchyNode = _allNodes.Item(uniqueID)
            Dim parentNodeID As String = elemNode.parentNodeKey

            Dim parentNode As clsHierarchyNode = Me.parentNodeItem(uniqueID)

            ' Eltern-Element aktualisieren 
            If Not IsNothing(parentNode) Then
                ' Kind-Eintrag löschen 
                parentNode.removeChild(uniqueID)
            End If

            ' kind-Elemente löschen oder umhängen
            If deleteALLChilds Then
                ' jetzt alle Kind-Elemente löschen 
                For i As Integer = 1 To elemNode.childCount
                    Dim childID As String = elemNode.getChild(i)
                    Me.removeNode(childID, True)
                Next
            Else
                ' jetzt alle Kind-Elemente umhängen
                For i As Integer = 1 To elemNode.childCount
                    Dim childNode As clsHierarchyNode
                    If _allNodes.ContainsKey(elemNode.getChild(i)) Then
                        childNode = _allNodes.Item(elemNode.getChild(i))
                        childNode.parentNodeKey = parentNodeID
                    End If
                Next
            End If

            ' jetzt das eigentliche Element löschen 
            _allNodes.Remove(uniqueID)

        Else
            ' nichts tun, ist eh nicht mehr existent ...
        End If


    End Sub

    ''' <summary>
    ''' löscht das Element an Stelle Index ;
    ''' Index kann Werte von 0 bis count-1 annehmen 
    ''' Vorsicht : mit dieser Funktionwerden keinerlei Konsistenzprüfungen vorgenommen, 
    ''' was die Behandlung von Kind- und Vater-Elementen angeht  
    ''' Diese Sub darf daher nur aufgerufen werden, wo die Konsistenz durch die übergeordnete Methode bereits sichergestellt wird. 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <remarks></remarks>
    Public Sub removeAt(ByVal index As Integer)

        If index >= 0 And index <= _allNodes.Count - 1 Then
            If _allNodes.ElementAt(index).Key <> rootPhaseName Then
                _allNodes.RemoveAt(index)
            End If

        End If

    End Sub

    ''' <summary>
    ''' erhöht / erniedrigt in der Hierarchie-Liste die Phasen-Verweise (indexOfElem) um increment
    ''' das wird benötigt, wenn zuvor ein Element gelöscht bzw neu in der Phasen-Liste ergänzt wurde 
    ''' </summary>
    ''' <param name="indexInPhaseList"></param>
    ''' <param name="increment"></param>
    ''' <remarks></remarks>
    Public Sub updatePhasenVerweise(ByVal indexInPhaseList As Integer, ByVal increment As Integer)

        Dim lastPhase As Integer = Me.getIndexOf1stMilestone - 1

        If lastPhase < 0 Then
            lastPhase = Me.count
        End If

        For i As Integer = 1 To lastPhase
            If _allNodes.ElementAt(i - 1).Value.indexOfElem > indexInPhaseList Then
                _allNodes.ElementAt(i - 1).Value.indexOfElem = _allNodes.ElementAt(i - 1).Value.indexOfElem + increment
            End If
        Next


    End Sub

    ''' <summary>
    ''' erhöht / erniedrigt in der Hierarchie-Liste die Meilenstein-Verweise (indexofElem) um increment
    ''' das darf aber nur bei Meilensteinen getan werden, die auch zum angegebenen Vater gehören 
    ''' </summary>
    ''' <param name="indexInMeilensteinListe"></param>
    ''' <param name="parentID"></param>
    ''' <param name="increment"></param>
    ''' <remarks></remarks>
    Public Sub updateMeilensteinVerweise(ByVal indexInMeilensteinListe As Integer, ByVal parentID As String, ByVal increment As Integer)

        Dim firstMilestone As Integer = Me.getIndexOf1stMilestone
        If firstMilestone < 0 Then
            ' nichts tun, es gibt keine Meilensteine 
        End If

        For i As Integer = firstMilestone To _allNodes.Count
            ' nur Meilensteine behandeln, deren Vater-ID mit der übergebenen parentID identisch ist 
            If _allNodes.ElementAt(i - 1).Value.parentNodeKey = parentID Then
                If _allNodes.ElementAt(i - 1).Value.indexOfElem > indexInMeilensteinListe Then
                    _allNodes.ElementAt(i - 1).Value.indexOfElem = _allNodes.ElementAt(i - 1).Value.indexOfElem + increment
                End If
            End If

        Next


    End Sub

    ''' <summary>
    ''' kopiert den Hierarchie Knoten ohne Überprüfungen in die Hierarchie
    ''' dies wird benötigt, wenn 
    ''' eine Projektvorlage in ein Projekt kopiert wird
    ''' ein Projekt in ein Projekt kopiert wird 
    ''' 
    ''' </summary>
    ''' <param name="elemNode">ausgefüllter elemNode</param>
    ''' <param name="elemKey">unique Key</param>
    ''' <remarks></remarks>
    Public Sub copyNode(ByVal elemNode As clsHierarchyNode, ByVal elemKey As String)

        If _allNodes.ContainsKey(elemKey) Then
            ' nichts tun 
        Else
            _allNodes.Add(elemKey, elemNode)
        End If

    End Sub


    ''' <summary>
    ''' gibt die Gesamt-Anzahl Elemente = Anzahl Phasen plus Anzahl Meilensteine zurück  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property count As Integer
        Get
            count = Me._allNodes.Count
        End Get
    End Property

    ''' <summary>
    ''' gibt die ID zurück, die das angegebene Element mit Nummer index hat 
    ''' index darf Werte von 1 .. Anzahl Elemente annehmen 
    ''' bei unzulässigem Index wird der leere String zurückgegeben  
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getIDAtIndex(ByVal index As Integer) As String
        Get
            If index >= 1 And index <= Me._allNodes.Count Then
                getIDAtIndex = Me._allNodes.ElementAt(index - 1).Key
            Else
                getIDAtIndex = ""
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt den Index zurück, an dem der erste Meilenstein steht ; mögliche Werte sind 1 ... Anzahl Elemente in der Hierarchie Liste 
    ''' da die Liste sortiert ist und Meilensteine alle mit 1§ beginnen, ist das erste Auftreten zugleich 
    ''' eins nach der letzten Phase
    ''' wenn es keinen Meilenstein gibt, dann ist das Ergebnis -1 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getIndexOf1stMilestone() As Integer
        Get
            Dim left As Integer = 0, right As Integer = Me._allNodes.Count - 1
            Dim anzElems As Integer = Me._allNodes.Count
            Dim curptr As Integer, testL As Integer, testR As Integer
            Dim found As Boolean = False
            Dim firstMilestone As Integer = -1

            curptr = CInt(anzElems / 2)
            curptr = left + CInt((right - left + 1) / 2)

            If curptr - 1 >= 0 Then
                testL = curptr - 1
            Else
                testL = 0
            End If

            If curptr + 1 <= anzElems - 1 Then
                testR = curptr + 1
            Else
                testR = anzElems - 1
            End If

            Do While Not found And (right - left) >= 1

                If Me._allNodes.ElementAt(curptr).Key.Substring(0, 1) <> Me._allNodes.ElementAt(testR).Key.Substring(0, 1) Then
                    ' found 
                    found = True
                    firstMilestone = curptr + 1

                ElseIf Me._allNodes.ElementAt(curptr).Key.Substring(0, 1) <> Me._allNodes.ElementAt(testL).Key.Substring(0, 1) Then
                    ' found
                    found = True
                    firstMilestone = curptr

                ElseIf Me._allNodes.ElementAt(curptr).Key.Substring(0, 1) = "1" Then
                    ' suche links weiter
                    right = testL

                Else
                    ' suche rechts weiter 
                    left = testR

                End If

                If Not found Then
                    curptr = left + CInt((right - left + 1) / 2)

                    If curptr - 1 >= 0 Then
                        testL = curptr - 1
                    Else
                        testL = 0
                    End If

                    If curptr + 1 <= anzElems - 1 Then
                        testR = curptr + 1
                    Else
                        testR = anzElems - 1
                    End If
                End If

            Loop

            ' wenn nicht gefunden, dann ist firstMilestone = -1 , das heisst, es gibt keine Meilensteine
            getIndexOf1stMilestone = firstMilestone + 1


        End Get
    End Property

    ''' <summary>
    ''' findet für elemName einen innerhalb seiner Geschwister eindeutigen Namen; 
    ''' falls ein neuer Name erzeugt wird, wird ein einentsprechender Eintrag in Phase/milestone bzw. den entsprechenden MissingDefs gemacht
    ''' wenn notwendig wird elemName solange mit einer lfdNr inkrementiert, bis der Name innerhalb seiner Geschwistergruppe eindeutig ist
    ''' Zu Geschwistern zählen die Kinder des gleichen Vaters, die auch vom gleichen Typ (Phasen oder Meilensteine sind)
    ''' es kann also einen Meilenstein und eine Phase gleichen Namens auf der gleichen Hierarchie-Stufe geben 
    ''' </summary>
    ''' <param name="parentElemID">ElemID des Vater-Knotens</param>
    ''' <param name="elemName">Urspünglicher Name des Elements</param>
    ''' <param name="isMilestone">handelt es sich um einen Meilenstein </param>
    ''' <value></value>
    ''' <returns>gibt einen eindeutigen Geschwister/Typ Namen zurück</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property findUniqueGeschwisterName(ByVal parentElemID As String, ByVal elemName As String, ByVal isMilestone As Boolean) As String

        Get
            Dim kennung As Integer
            If isMilestone Then
                kennung = 1
            Else
                kennung = 0
            End If

            Dim geschwisterTypGruppe As Collection = Me.getChildNamesOf(parentElemID, isMilestone)
            ' geschwistergruppe rechnest du dadrin aus ..
            Dim lfdNr As Integer = 2

            Dim uniqueSiblingName As String = elemName

            ' wenn bereits enthalten: suche einen neuen, abgeleiteten Namen
            Do While geschwisterTypGruppe.Contains(uniqueSiblingName)
                uniqueSiblingName = elemName & " " & lfdNr.ToString
                lfdNr = lfdNr + 1
            Loop

            ' hier muss ggf das Element noch in Phase- bzw. Milestone-Definitions bzw. die entsprechenden Missing-Definitions aufgenommen werden 
            If uniqueSiblingName <> elemName Then
                Dim isMissing As Boolean = False

                If isMilestone Then
                    ' Meilenstein-Behandlung 
                    '
                    Dim newMSDef As New clsMeilensteinDefinition
                    Dim sisterDef As clsMeilensteinDefinition = MilestoneDefinitions.getMilestoneDef(elemName)

                    If IsNothing(sisterDef) Then
                        sisterDef = missingMilestoneDefinitions.getMilestoneDef(elemName)
                        isMissing = True
                    End If

                    If Not IsNothing(sisterDef) Then

                        With newMSDef
                            .name = uniqueSiblingName
                            .schwellWert = 0
                            .shortName = sisterDef.shortName
                            .darstellungsKlasse = sisterDef.darstellungsKlasse
                            '.farbe = sisterDef.farbe
                        End With

                        If Not isMissing Then
                            If Not MilestoneDefinitions.Contains(newMSDef.name) Then
                                MilestoneDefinitions.Add(newMSDef)
                            End If
                        Else
                            If Not missingMilestoneDefinitions.Contains(newMSDef.name) Then
                                missingMilestoneDefinitions.Add(newMSDef)
                            End If
                        End If


                    End If
                Else
                    ' Phasen-Behandlung 
                    '
                    Dim newPhDef As New clsPhasenDefinition
                    Dim sisterDef As clsPhasenDefinition = PhaseDefinitions.getPhaseDef(elemName)

                    If IsNothing(sisterDef) Then
                        sisterDef = missingPhaseDefinitions.getPhaseDef(elemName)
                        isMissing = True
                    End If

                    If Not IsNothing(sisterDef) Then

                        With newPhDef
                            .name = uniqueSiblingName
                            .schwellWert = 0
                            .shortName = sisterDef.shortName
                            .darstellungsKlasse = sisterDef.darstellungsKlasse
                            '.farbe = sisterDef.farbe
                        End With

                        If Not isMissing Then
                            If Not PhaseDefinitions.Contains(newPhDef.name) Then
                                PhaseDefinitions.Add(newPhDef)
                            End If
                        Else
                            If Not missingPhaseDefinitions.Contains(newPhDef.name) Then
                                missingPhaseDefinitions.Add(newPhDef)
                            End If
                        End If


                    End If
                End If
            End If
            ' Ende der Aufnahme in Phase/MilestoneDes; notwenig aufgrund neuen Sibling Namens 

            findUniqueGeschwisterName = uniqueSiblingName

        End Get
    End Property
    ''' <summary>
    ''' berechnet den Unique Namen ElemKey für den gegebenen elemName und Angabe, ob Meilenstein oder nicht 
    ''' findet auf jeden Fall einen Namen, der in der sortedList noch nicht enthalten ist  
    ''' </summary>
    ''' <param name="elemName"></param>
    ''' <param name="isMilestone"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property findUniqueElemKey(ByVal elemName As String, ByVal isMilestone As Boolean) As String
        Get
            Dim elemKey As String
            Dim lfdNr As Integer = 2


            elemKey = calcHryElemKey(elemName, isMilestone)

            If _allNodes.ContainsKey(elemKey) Then

                elemKey = calcHryElemKey(elemName, isMilestone, lfdNr)

                Do While _allNodes.ContainsKey(elemKey)
                    lfdNr = lfdNr + 1
                    elemKey = calcHryElemKey(elemName, isMilestone, lfdNr)
                Loop

            End If

            findUniqueElemKey = elemKey

        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn der angebenene Meilenstein in der angebenen Hierarchie schon mal existiert 
    ''' false sonst
    ''' </summary>
    ''' <param name="elemName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsMilestone(ByVal elemName As String, Optional ByVal breadcrumb As String = "") As Boolean
        Get

            Dim milestoneIndices(,) As Integer = Me.getMilestoneIndices(elemName, breadcrumb)

            If milestoneIndices(0, 0) > 0 And milestoneIndices(1, 0) > 0 Then
                containsMilestone = True
            Else
                containsMilestone = False
            End If


        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn die angegebene Phase in der angegebenen Hierarchie existiert  
    ''' </summary>
    ''' <param name="elemName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsPhase(ByVal elemName As String, Optional ByVal breadcrumb As String = "") As Boolean
        Get
            Dim phaseIndices() As Integer = Me.getPhaseIndices(elemName, breadcrumb)

            If phaseIndices(0) > 0 Then
                containsPhase = True
            Else
                containsPhase = False
            End If

        End Get
    End Property


    ''' <summary>
    ''' gibt an, ob dieser Schlüssel bereits in der Liste vorhanden ist
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsKey(ByVal uniqueID As String) As Boolean
        Get
            containsKey = _allNodes.ContainsKey(uniqueID)
        End Get
    End Property

    ''' <summary>
    ''' gibt den  Index in der Liste der Phasen bzw. Meilensteine zurück
    ''' wenn es sich um eine Phase handelt: welcher Index von 1 .. countPhases ist es 
    ''' wenn es sich um einen Meilenstein handelt: welcher Index von 1..countMilestones ist es 
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPMIndexOfID(ByVal uniqueID As String) As Integer
        Get
            If _allNodes.ContainsKey(uniqueID) Then
                getPMIndexOfID = _allNodes.Item(uniqueID).indexOfElem
            Else
                getPMIndexOfID = -1
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt den kürzesten eindeutigen Namen für das Element zurück, der sich finden lässt
    ''' optional kann die SwimlaneID mitgegeben werden - dann wird nur nach eindeutigen Namen innerhalb der swimlanes gesucht 
    ''' wenn das Element eh eindeutig ist im Projekt, dann wird nur der Elem-Name zurückgegeben 
    ''' </summary>
    ''' <param name="nameID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBestNameOfID(ByVal nameID As String, _
                                             ByVal ShowStdNames As Boolean, ByVal showAbbrev As Boolean, _
                                             Optional ByVal swimlaneID As String = rootPhaseName) As String
        Get
            Dim elemName As String = elemNameOfElemID(nameID)
            Dim isMilestone As Boolean
            Dim curBC As String = ""
            Dim oldBC As String = ""
            Dim anzElements As Integer
            Dim anzElementsBefore As Integer
            Dim level As Integer = 1
            Dim tmpName As String = elemName
            Dim rootreached As Boolean = False
            Dim description1 As String = "", description2 As String = elemName
            Dim phDef As clsPhasenDefinition
            Dim swlBC As String = ""



            isMilestone = elemIDIstMeilenstein(nameID)

            If swimlaneID = rootPhaseName Then
                swlBC = ""
            Else
                If istElemID(swimlaneID) Then
                    swlBC = calcHryFullname(elemNameOfElemID(swimlaneID), _
                                                  Me.getBreadCrumb(swimlaneID))
                End If
            End If

            Try
                If isMilestone Then

                    ' Änderung tk: es wird der eindeutige Namen unterhalb der swimlaneID gesucht  
                    'Dim milestoneIndices(,) As Integer = Me.getMilestoneIndices(elemName, "")
                    Dim milestoneIndices(,) As Integer = Me.getMilestoneIndices(elemName, swlBC)
                    anzElements = CInt(milestoneIndices.Length / 2)

                    If anzElements > 1 Then

                        anzElementsBefore = anzElements

                        Do Until anzElements = 1 Or rootreached
                            curBC = Me.getBreadCrumb(nameID, level)

                            If oldBC = curBC Then
                                rootreached = True
                            Else
                                oldBC = curBC
                            End If

                            If Not rootreached Then
                                milestoneIndices = Me.getMilestoneIndices(elemName, curBC)
                                anzElements = CInt(milestoneIndices.Length / 2)
                                If anzElements < anzElementsBefore Then
                                    anzElementsBefore = anzElements
                                    tmpName = calcHryFullname(elemName, curBC)
                                End If
                            End If

                            level = level + 1

                        Loop
                    Else
                        tmpName = elemName
                    End If


                Else
                    ' es handelt sich um eine Phase
                    'Dim phaseIndices() As Integer = Me.getPhaseIndices(elemName, "")
                    ' Änderung tk: es wird der eindeutige Namen unterhalb der swimlaneID gesucht  
                    Dim phaseIndices() As Integer = Me.getPhaseIndices(elemName, swlBC)
                    anzElements = phaseIndices.Length

                    If anzElements > 1 Then

                        anzElementsBefore = anzElements

                        Do Until anzElements = 1 Or rootreached
                            curBC = Me.getBreadCrumb(nameID, level)

                            If oldBC = curBC Then
                                rootreached = True
                            Else
                                oldBC = curBC
                            End If

                            If Not rootreached Then
                                phaseIndices = Me.getPhaseIndices(elemName, curBC)
                                anzElements = phaseIndices.Length
                                If anzElements < anzElementsBefore Then
                                    anzElementsBefore = anzElements
                                    tmpName = calcHryFullname(elemName, curBC)
                                End If
                            End If

                            level = level + 1

                        Loop
                    Else
                        tmpName = elemName
                    End If
                End If
            Catch ex As Exception

            End Try

            ' jetzt wird unterschieden, ob Abbrev gezeigt werden soll oder Standard Name ... 
            If ShowStdNames Then
                If showAbbrev Then

                    If awinSettings.showBestName And Not awinSettings.drawphases Then
                        ' den bestmöglichen, also den kürzesten Breadcrumb Namen, der (möglichst) eindeutig ist
                        ' anzeigen; aber nur, wenn im Ein-Zeile-Modus beschriftet wird, weil dann der Kontext fehlt ... 
                        Call splitHryFullnameTo2(tmpName, description2, description1)

                        Dim tmpStr() As String = description1.Split(New Char() {CChar("#")}, 20)

                        ' jetzt den Abbrev String zusammensetzen 
                        Dim newDesc1 As String = ""
                        For i As Integer = 1 To tmpStr.Length
                            Dim tmpPhName As String = tmpStr(i - 1)
                            phDef = PhaseDefinitions.getPhaseDef(tmpPhName)

                            If IsNothing(phDef) Then
                                If i = 1 And tmpPhName <> elemNameOfElemID(rootPhaseName) And tmpPhName <> "" Then
                                    ' den tmpPhName eintragen 
                                    newDesc1 = tmpPhName
                                ElseIf i > 1 Then
                                    newDesc1 = newDesc1 & tmpPhName
                                End If

                            Else
                                If i = 1 Then
                                    If phDef.shortName = "" Then
                                        newDesc1 = tmpPhName
                                    Else
                                        newDesc1 = phDef.shortName
                                    End If

                                Else
                                    If phDef.shortName = "" Then
                                        newDesc1 = newDesc1 & tmpPhName
                                    Else
                                        newDesc1 = newDesc1 & "-" & phDef.shortName
                                    End If

                                End If
                            End If

                        Next
                        description1 = newDesc1

                        If isMilestone Then

                            Dim msDef As clsMeilensteinDefinition
                            msDef = MilestoneDefinitions.getMilestoneDef(description2)
                            If IsNothing(msDef) Then
                                msDef = missingMilestoneDefinitions.getMilestoneDef(description2)
                            End If

                            If IsNothing(msDef) Then
                                ' nichts zu tun
                            Else
                                If IsNothing(msDef.shortName) Then
                                    'description2 = "-"
                                Else
                                    If msDef.shortName = "" Then
                                        'description2 = "-"
                                    Else
                                        description2 = msDef.shortName
                                    End If

                                End If
                            End If

                        Else

                            phDef = PhaseDefinitions.getPhaseDef(description2)
                            If IsNothing(phDef) Then
                                
                                phDef = missingPhaseDefinitions.getPhaseDef(description2)
                            End If

                            If IsNothing(phDef) Then
                                ' nichts zu tun
                            Else

                                If IsNothing(phDef.shortName) Then
                                    'description2 = "-"
                                Else
                                    If phDef.shortName = "" Then
                                        'description2 = "-"
                                    Else
                                        description2 = phDef.shortName
                                    End If

                                End If
                            End If

                        End If
                Else
                    description1 = ""

                        If isMilestone Then

                            Dim msDef As clsMeilensteinDefinition
                            msDef = MilestoneDefinitions.getMilestoneDef(description2)
                            If IsNothing(msDef) Then
                                msDef = missingMilestoneDefinitions.getMilestoneDef(description2)
                            End If

                            If IsNothing(msDef) Then
                                'description2 = "-"
                            Else

                                If IsNothing(msDef.shortName) Then
                                    'description2 = "-"
                                Else
                                    If msDef.shortName = "" Then
                                        'description2 = msDef.name
                                    Else
                                        description2 = msDef.shortName
                                    End If

                                End If
                            End If

                        Else

                            phDef = PhaseDefinitions.getPhaseDef(description2)
                            If IsNothing(phDef) Then

                                phDef = missingPhaseDefinitions.getPhaseDef(description2)
                            End If

                            If IsNothing(phDef) Then
                                'description2 = "-"
                            Else

                                If IsNothing(phDef.shortName) Then
                                    'description2 = "-"
                                Else
                                    If phDef.shortName = "" Then
                                        'description2 = phDef.name
                                    Else
                                        description2 = phDef.shortName
                                    End If

                                End If
                            End If

                        End If


                End If
            Else

                Call splitHryFullnameTo2(tmpName, description2, description1)
                Dim tmpStr() As String = description1.Split(New Char() {CChar("#")}, 20)

                ' jetzt den Std-Name zusammensetzen 
                Dim newDesc1 As String = ""
                For i As Integer = 1 To tmpStr.Length
                    Dim tmpPhName As String = tmpStr(i - 1)

                    If i = 1 Then
                        If tmpPhName = elemNameOfElemID(rootPhaseName) Then
                            ' nichts tun
                        Else
                            newDesc1 = tmpPhName
                        End If
                    ElseIf i > 1 Then
                        newDesc1 = newDesc1 & "-" & tmpPhName
                    End If

                Next
                description1 = newDesc1


            End If
            Else
            description2 = Me.nodeItem(nameID).origName
            End If

            Dim description As String = ""
            Try
                If description1 <> "" Then
                    description = description1 & "-" & description2
                Else
                    description = description2
                End If
            Catch ex As Exception

            End Try

            getBestNameOfID = description

        End Get
    End Property
    ''' <summary>
    ''' gibt den Index in der Hierarchie zurück, den das Element mit uniqueID hat
    ''' wenn uniqueID nicht existiert, dann wird als Wert 0 zurückgegeben  
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getIndexOfID(ByVal uniqueID As String) As Integer
        Get
            If _allNodes.ContainsKey(uniqueID) Then
                getIndexOfID = _allNodes.IndexOfKey(uniqueID) + 1
            Else
                getIndexOfID = 0
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt die ParentID zurück, die das Element mit ID uniqueID hat
    ''' wenn das Element gar nicht existiert wird "" zurückgegeben; ebenso, wenn es keinen PArent gibt 
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getParentIDOfID(ByVal uniqueID As String) As String
        Get
            If _allNodes.ContainsKey(uniqueID) Then
                getParentIDOfID = _allNodes.Item(uniqueID).parentNodeKey
            Else
                getParentIDOfID = ""
            End If
        End Get
    End Property


    ''' <summary>
    ''' gibt den Hierarchie Knoten zurück , der die angegebene uniqueID hat
    ''' VORSICHT: liefert Nothing zurück , wenn die uniqueID nicht existiert 
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property nodeItem(ByVal uniqueID As String) As clsHierarchyNode
        Get
            If _allNodes.ContainsKey(uniqueID) Then
                nodeItem = _allNodes.Item(uniqueID)
            Else
                nodeItem = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt den Hierarchie Knoten zurück des Parent-Elements von uniqueID zurück 
    ''' wenn der nicht existiert: Nothing 
    ''' also auch im Fall rootPhaseName 
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property parentNodeItem(ByVal uniqueID As String) As clsHierarchyNode
        Get
            If uniqueID = rootPhaseName Then
                parentNodeItem = Nothing
            Else
                Dim elemNode As clsHierarchyNode
                If _allNodes.ContainsKey(uniqueID) Then
                    elemNode = _allNodes.Item(uniqueID)
                    If _allNodes.ContainsKey(elemNode.parentNodeKey) Then
                        parentNodeItem = _allNodes.Item(elemNode.parentNodeKey)
                    Else
                        parentNodeItem = Nothing
                    End If
                Else
                    parentNodeItem = Nothing
                End If

            End If
        End Get
    End Property


    ''' <summary>
    ''' gibt den Hierarchie Knoten zurück, der in der sortierten Liste an Position Index steht 
    ''' Index läuft von 1 .. Anzahl Knoten 
    ''' bei ungültigem Index wird Nothing zurückgegeben 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property nodeItem(ByVal index As Integer) As clsHierarchyNode
        Get
            If index >= 1 And index <= _allNodes.Count Then
                nodeItem = _allNodes.ElementAt(index - 1).Value
            Else
                nodeItem = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt einen Array von Index Nummern aus der Hierarchie Liste zurück; 
    ''' gültige Werte sind 1 .. Anzahl Hierarchie Einträge
    ''' Wert = 0 bedeutet: Element existiert nicht 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseHryIndices(ByVal name As String, ByVal breadcrumb As String) As Integer()
        Get

            Dim phaseIndices() As Integer
            Dim first As Integer, last As Integer
            Dim elemID As String = calcHryElemKey(name, False)
            Dim i As Integer, k As Integer
            Dim anzahlNodes As Integer = _allNodes.Count

            ' da die Liste sortiert ist, reicht es den Index des ersten und den Index des letzten Elements zu bestimmen 

            ReDim phaseIndices(0)
            If _allNodes.ContainsKey(elemID) Then
                first = _allNodes.IndexOfKey(calcHryElemKey(name, False))

                i = first + 1
                Dim otherNamefound As Boolean = False

                Do While Not otherNamefound And i <= anzahlNodes - 1
                    If elemNameOfElemID(_allNodes.ElementAt(i).Key) <> name Then
                        otherNamefound = True
                    Else
                        i = i + 1
                    End If
                Loop

                last = i - 1

                If breadcrumb = "" Then
                    ReDim phaseIndices(last - first)
                    For i = 0 To last - first
                        phaseIndices(i) = first + i + 1
                    Next
                Else
                    Dim tmpIndices(last - first) As Integer
                    Dim identicalBC As Boolean = True
                    Dim bcLevels() As String
                    bcLevels = breadcrumb.Split((New Char() {CChar("#")}), 20)
                    Dim anzLevel As Integer = bcLevels.Length

                    ' herausfinden, welche Elemente auch den gleichen Breadcrumb haben ..
                    ' sie können den gleichen Element-Namen haben, aber evtl ganz unterschiedliche Breadcrumbs
                    k = 0
                    For i = 0 To last - first
                        Dim vglbreadCrumb As String = Me.getBreadCrumb(_allNodes.ElementAt(first + i).Key, anzLevel)
                        If vglbreadCrumb = breadcrumb Then
                            tmpIndices(k) = first + i + 1
                            k = k + 1
                        End If
                    Next

                    ' jetzt müssen die gefundenen Elemente umkopiert werden 
                    If k > 0 Then
                        ' es wurde mindestens ein Element gefunden 
                        k = k - 1
                        ReDim phaseIndices(k)
                        For ii As Integer = 0 To k
                            phaseIndices(ii) = tmpIndices(ii)
                        Next
                    End If

                End If

            End If

            getPhaseHryIndices = phaseIndices


        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste von Phasen-Indices zurück; jeder Index bezeichnet eine Phase, die den Elem-Namen trägt 
    ''' und die mit der angegeben Hierarchie übereinstimmt; Hierarchie kann  
    ''' wenn ein 1-elementiger Array zurückgegeben wird, dessen Wert = 0 ist, so existiert diese Phase nicht  
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseIndices(ByVal name As String, Optional ByVal breadcrumb As String = "") As Integer()
        Get
            Dim phaseIndices() As Integer
            Dim first As Integer, last As Integer
            Dim elemID As String = calcHryElemKey(name, False)
            Dim i As Integer, k As Integer
            Dim anzahlNodes As Integer = _allNodes.Count

            ' da die Liste sortiert ist, reicht es den Index des ersten und den Index des letzten Elements zu bestimmen 

            ReDim phaseIndices(0)
            If _allNodes.ContainsKey(elemID) Then
                first = _allNodes.IndexOfKey(calcHryElemKey(name, False))

                i = first + 1
                Dim otherNamefound As Boolean = False

                Do While Not otherNamefound And i <= anzahlNodes - 1
                    If elemNameOfElemID(_allNodes.ElementAt(i).Key) <> name Then
                        otherNamefound = True
                    Else
                        i = i + 1
                    End If
                Loop

                last = i - 1

                If breadcrumb = "" Then
                    ReDim phaseIndices(last - first)
                    For i = 0 To last - first
                        phaseIndices(i) = _allNodes.ElementAt(first + i).Value.indexOfElem
                    Next
                Else
                    Dim tmpIndices(last - first) As Integer
                    Dim identicalBC As Boolean = True
                    Dim bcLevels() As String
                    bcLevels = breadcrumb.Split((New Char() {CChar("#")}), 20)
                    Dim anzLevel As Integer = bcLevels.Length

                    ' herausfinden, welche Elemente auch den gleichen Breadcrumb haben ..
                    ' sie können den gleichen Element-Namen haben, aber evtl ganz unterschiedliche Breadcrumbs
                    k = 0
                    For i = 0 To last - first
                        Dim vglbreadCrumb As String = Me.getBreadCrumb(_allNodes.ElementAt(first + i).Key, anzLevel)
                        If vglbreadCrumb = breadcrumb Then
                            tmpIndices(k) = _allNodes.ElementAt(first + i).Value.indexOfElem
                            k = k + 1
                        End If
                    Next

                    ' jetzt müssen die gefundenen Elemente umkopiert werden 
                    If k > 0 Then
                        ' es wurde mindestens ein Element gefunden 
                        k = k - 1
                        ReDim phaseIndices(k)
                        For ii As Integer = 0 To k
                            phaseIndices(ii) = tmpIndices(ii)
                        Next
                    End If

                End If

            End If

            getPhaseIndices = phaseIndices

        End Get
    End Property

    ''' <summary>
    ''' gibt einen Array von Meilenstein Indices zurück; Index(0) bezeichnet den ersten Meilenstein in der Hierarchie, der den übergebenen NAmen und den
    ''' übergebenen Breadcrumb trägt; Index(x) kann Werte von 1...anZahl Elemente in Hierarchie Liste sein
    ''' Index(x) = 0 bedeutet, es gibt ihn nicht   
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneHryIndices(ByVal name As String, Optional ByVal breadcrumb As String = "") As Integer()
        Get
            Dim milestoneIndices() As Integer
            Dim first As Integer, last As Integer
            Dim elemID As String = calcHryElemKey(name, True)
            Dim i As Integer, k As Integer
            Dim anzahlNodes As Integer = _allNodes.Count

            ' da die Liste sortiert ist, reicht es den Index des ersten und den Index des letzten Elements zu bestimmen 

            ReDim milestoneIndices(0)
            If _allNodes.ContainsKey(elemID) Then
                first = _allNodes.IndexOfKey(elemID)
                i = first + 1

                Dim otherNamefound As Boolean = False
                Do While Not otherNamefound And i <= anzahlNodes - 1
                    If elemNameOfElemID(_allNodes.ElementAt(i).Key) <> name Then
                        otherNamefound = True
                    Else
                        i = i + 1
                    End If
                Loop

                last = i - 1

                If breadcrumb = "" Then
                    ReDim milestoneIndices(last - first)
                    For i = 0 To last - first
                        milestoneIndices(i) = first + i + 1 ' da 1 das Elementat(0) bezeichnet 
                    Next
                Else
                    Dim tmpIndices(last - first) As Integer
                    Dim identicalBC As Boolean = True
                    Dim bcLevels() As String
                    bcLevels = breadcrumb.Split((New Char() {CChar("#")}), 20)
                    Dim anzLevel As Integer = bcLevels.Length
                    k = 0

                    For i = 0 To last - first
                        Dim vglbreadCrumb As String = Me.getBreadCrumb(_allNodes.ElementAt(first + i).Key, anzLevel)
                        If vglbreadCrumb = breadcrumb Then
                            tmpIndices(k) = first + i + 1 ' da 1 das ElementAt(0) bezeichnet
                            k = k + 1
                        End If
                    Next

                    ' jetzt muss das gefundene Ergebnis umkopiert werden  
                    If k > 0 Then
                        ' es wurde mindestens ein Element gefunden 
                        k = k - 1
                        ReDim milestoneIndices(k)
                        For ii As Integer = 0 To k
                            milestoneIndices(ii) = tmpIndices(ii)
                        Next
                    End If

                End If

            End If

            getMilestoneHryIndices = milestoneIndices

        End Get
    End Property


    ''' <summary>
    ''' gibt eine Liste von Meilenstein-Indices zurück;  Index(0,) bezeichnet die Phasen-Nummer, Index(1,) bezeichnet die Meilenstein-Nummer im entsprechenden Projekt
    ''' es werden nur die Elemente zurückgegeben,  die mit der angegeben Hierarchie übereinstimmen; Breadcrumb kann leer sein, dann wird alles gesucht   
    ''' wenn ein 2-elementiger Array zurückgegeben wird, dessen Wert jeweils 0 ist, so existiert dieser Meilenstein nicht   
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneIndices(ByVal name As String, Optional ByVal breadcrumb As String = "") As Integer(,)
        Get
            Dim milestoneIndices(,) As Integer
            Dim first As Integer, last As Integer
            Dim elemID As String = calcHryElemKey(name, True)
            Dim i As Integer, k As Integer
            Dim phaseID As String
            Dim anzahlNodes As Integer = _allNodes.Count

            ' da die Liste sortiert ist, reicht es den Index des ersten und den Index des letzten Elements zu bestimmen 

            ReDim milestoneIndices(1, 0)
            If _allNodes.ContainsKey(elemID) Then
                first = _allNodes.IndexOfKey(elemID)
                i = first + 1

                Dim otherNamefound As Boolean = False
                Do While Not otherNamefound And i <= anzahlNodes - 1
                    If elemNameOfElemID(_allNodes.ElementAt(i).Key) <> name Then
                        otherNamefound = True
                    Else
                        i = i + 1
                    End If
                Loop

                last = i - 1

                If breadcrumb = "" Then
                    ReDim milestoneIndices(1, last - first)
                    For i = 0 To last - first
                        phaseID = _allNodes.ElementAt(first + i).Value.parentNodeKey
                        milestoneIndices(0, i) = Me.getPMIndexOfID(phaseID)
                        milestoneIndices(1, i) = _allNodes.ElementAt(first + i).Value.indexOfElem
                    Next
                Else
                    Dim tmpIndices(1, last - first) As Integer
                    Dim identicalBC As Boolean = True
                    Dim bcLevels() As String
                    bcLevels = breadcrumb.Split((New Char() {CChar("#")}), 20)
                    Dim anzLevel As Integer = bcLevels.Length
                    k = 0
                    For i = 0 To last - first
                        Dim vglbreadCrumb As String = Me.getBreadCrumb(_allNodes.ElementAt(first + i).Key, anzLevel)
                        If vglbreadCrumb = breadcrumb Then
                            phaseID = _allNodes.ElementAt(first + i).Value.parentNodeKey
                            tmpIndices(0, k) = Me.getPMIndexOfID(phaseID)
                            tmpIndices(1, k) = _allNodes.ElementAt(first + i).Value.indexOfElem
                            k = k + 1
                        End If
                    Next

                    If k > 0 Then
                        ' es wurde mindestens ein Element gefunden 
                        k = k - 1
                        ReDim milestoneIndices(1, k)
                        For ii As Integer = 0 To k
                            milestoneIndices(0, ii) = tmpIndices(0, ii)
                            milestoneIndices(1, ii) = tmpIndices(1, ii)
                        Next
                    End If

                End If

            End If

            getMilestoneIndices = milestoneIndices

        End Get
    End Property



    ''' <summary>
    ''' gibt für die angegebene elemID den Breadcrumb zurück
    ''' die einzelnen Ebenen werden mit # voneinander getrennt
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <param name="ebene">100: die gesamte Hierarchie
    ''' 1, 2, ..: soviele Stufen wie angegeben sind; 
    ''' wenn mehr Stufen angegeben sind als vorhanden, wird bei . (root) abgebrochen</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBreadCrumb(ByVal elemID As String, Optional ByVal ebene As Integer = 100) As String
        Get
            Dim tmpBreadCrumb = ""
            Dim tmpEbene As Integer = 1
            Dim rootReached As Boolean = False
            Dim currentElemID As String = elemID
            Dim parentID As String = ""
            Dim rootkey As String = calcHryElemKey(".", False)

            ' sicherstellen, dass ebene keinen blödsinningen Wert hat ;
            ' wenn der Wert < 0 ist , wird automatisch auf 0 gesetzt 
            If ebene < 0 Then
                ebene = 0
            End If

            Dim ok As Boolean = True

            If elemID <> rootkey And ebene > 0 Then

                Do While tmpEbene <= ebene And Not rootReached And ok

                    If Me._allNodes.ContainsKey(currentElemID) Then

                        parentID = Me._allNodes.Item(currentElemID).parentNodeKey
                        If parentID = "" Then
                            rootReached = True
                        End If
                        If Me._allNodes.ContainsKey(parentID) Then

                            If tmpEbene = 1 Then
                                tmpBreadCrumb = Me._allNodes.Item(parentID).elemName
                            Else
                                tmpBreadCrumb = Me._allNodes.Item(parentID).elemName & "#" & tmpBreadCrumb
                            End If
                            currentElemID = parentID

                        Else
                            ok = False
                            If tmpEbene = 1 Then
                                tmpBreadCrumb = "?"
                            ElseIf Not rootReached Then
                                tmpBreadCrumb = "?" & "#" & tmpBreadCrumb
                            End If
                        End If

                    Else
                        ok = False
                        If tmpEbene = 1 Then
                            tmpBreadCrumb = "?"
                        Else
                            tmpBreadCrumb = "?" & "#" & tmpBreadCrumb
                        End If
                    End If


                    tmpEbene = tmpEbene + 1


                Loop

            Else
                tmpBreadCrumb = ""
            End If

            getBreadCrumb = tmpBreadCrumb

        End Get
    End Property

    ''' <summary>
    ''' gibt den Hierarchie Level der elemID zurück 
    ''' 0: es handelt sich um den RootKnoten
    ''' x: Element ist auf der x-Ten Hierarchie Stufe 
    ''' -1: elemID oder eiens der Vater Elemente existieren nicht in der Hierarchie
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getIndentLevel(ByVal elemID As String) As Integer
        Get
            Dim tmpEbene As Integer = 0

            ' sicherstellen, ob elemID überhaupt existiert , wenn nein, dann wird "-1" zurückgegeben 
            If Me._allNodes.ContainsKey(elemID) Then

                Dim tmpBreadCrumb = ""
                Dim rootReached As Boolean = False
                Dim currentElemID As String = elemID
                Dim parentID As String = ""

                Dim ok As Boolean = True

                If elemID <> rootPhaseName Then

                    Do While Not rootReached And ok

                        parentID = Me._allNodes.Item(currentElemID).parentNodeKey

                        If currentElemID = rootPhaseName Then
                            rootReached = True

                        ElseIf Me._allNodes.ContainsKey(parentID) Then
                            tmpEbene = tmpEbene + 1
                            currentElemID = parentID

                        Else
                            ok = False
                            tmpEbene = -1
                        End If

                    Loop

                Else
                    tmpEbene = 0
                End If

            Else
                tmpEbene = -1
            End If

            getIndentLevel = tmpEbene
        End Get
    End Property
    Public Sub New()
        _allNodes = New SortedList(Of String, clsHierarchyNode)
    End Sub


End Class
