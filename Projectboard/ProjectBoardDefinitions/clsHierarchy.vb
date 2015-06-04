Public Class clsHierarchy
    Private _allNodes As SortedList(Of String, clsHierarchyNode)

    ''' <summary>
    ''' fügt der Hierarchy einen Knoten hinzu
    ''' </summary>
    ''' <param name="elemNode"></param>
    ''' <remarks></remarks>
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
    ''' gibt den eindeutigsten Namen für das element zurück, der sich finden lässt
    ''' entweder den das Element eindeutig machenden Breadcrumb Namen oder den Breadcrumb Namen, mit dem am wenigsten Mehrdeutigkeiten existieren
    ''' wenn das Element eh eindeutig ist im Projekt, dann wird nur der Elem-Name zurückgegeben 
    ''' </summary>
    ''' <param name="nameID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBestNameOfID(ByVal nameID As String) As String
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
            isMilestone = elemIDIstMeilenstein(nameID)

            Try
                If isMilestone Then

                    Dim milestoneIndices(,) As Integer = Me.getMilestoneIndices(elemName, "")
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


                        Loop
                    Else
                        tmpName = elemName
                    End If
                Else
                    ' es handelt sich um eine Phase
                    Dim phaseIndices() As Integer = Me.getPhaseIndices(elemName, "")
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


                        Loop
                    Else
                        tmpName = elemName
                    End If
                End If
            Catch ex As Exception

            End Try
            


            getBestNameOfID = tmpName

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
