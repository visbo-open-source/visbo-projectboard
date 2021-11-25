''' <summary>
''' Klasse für AlleProjekte
''' </summary>
''' <remarks></remarks>
Public Class clsProjekteAlle

    ' in dieser Klasse ist der Key zusammengesetzt aus ProjektName und VariantName mit calcProjektKey(hproj)
    Private _allProjects As SortedList(Of String, clsProjekt)

    Public Sub New()
        _allProjects = New SortedList(Of String, clsProjekt)
    End Sub

    ''' <summary>
    ''' erstellt eine Kopie der Liste 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property createCopy(Optional filteredBy As clsConstellation = Nothing) As clsProjekteAlle
        Get
            Dim tmpKopie As New clsProjekteAlle

            If IsNothing(filteredBy) Then
                For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                    If Not tmpKopie.Containskey(kvp.Key) Then
                        tmpKopie.Add(kvp.Value, updateCurrentConstellation:=False)
                    End If
                Next
            Else
                ' nur die übernehmen, die auch in der Constellation enthalten sind 
                For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                    If filteredBy.contains(kvp.Key, False) And Not tmpKopie.Containskey(kvp.Key) Then
                        tmpKopie.Add(kvp.Value, updateCurrentConstellation:=False)
                    End If
                Next
            End If

            createCopy = tmpKopie

        End Get
    End Property

    ''' <summary>
    ''' gibt zurück, ob die kdNr bereits in einem der Projekte von der KlassenInstanz clsProjekteAlle enthalten ist ...
    ''' </summary>
    ''' <param name="kdNr"></param>
    ''' <returns></returns>
    Public ReadOnly Property containsPNr(ByVal kdNr As String) As Boolean
        Get
            Dim tmpResult As Boolean = False

            If Not IsNothing(kdNr) Then

                If kdNr <> "" Then
                    For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                        If Not IsNothing(kvp.Value.kundenNummer) Then
                            If kvp.Value.kundenNummer = kdNr Then
                                tmpResult = True
                                Exit For
                            End If
                        End If

                    Next
                End If

            End If

            containsPNr = tmpResult
        End Get
    End Property
    ''' <summary>
    ''' gets the RoleNameIDs of existing skills  
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getRoleSkillIDs() As Collection
        Get

            Dim roleSkillIDs As New Collection

            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getSkillNameIds

                For Each tmpName As String In tmpCollection
                    If Not roleSkillIDs.Contains(tmpName) Then
                        roleSkillIDs.Add(tmpName, tmpName)
                    End If
                Next

            Next


            getRoleSkillIDs = roleSkillIDs
        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück wenn irgendein Summary Projekt in der Liste enthalten ist 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property containsAnySummaryProject() As Boolean

        Get
            Dim tmpResult As Boolean = False
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                If kvp.Value.projectType = ptPRPFType.portfolio Then
                    tmpResult = True
                    Exit For
                End If
            Next
            containsAnySummaryProject = tmpResult
        End Get

    End Property

    ''' <summary>
    ''' gibt true zurück, wenn es Konflikte gibt, das heisst wenn dieses Element in der sortedListInQuestion enthalten ist
    ''' kommt dann vor, wenn ein Summary Projekt  in AlleProjekte soll, das dort enthaltene Projekte umfasst   
    ''' </summary>
    ''' <param name="myElem"></param>
    ''' <param name="sortedListInQuestion"></param>
    ''' <returns></returns>
    Private Function elemHasConflictsWith(ByVal myElem As String, ByVal sortedListInQuestion As SortedList(Of String, Boolean)) As Boolean

        Dim tmpResult As Boolean = False
        Dim mySortedList As New SortedList(Of String, Boolean)

        ' nur das eine Element untersuchen 
        If _allProjects.ContainsKey(myElem) Then
            Dim myProject As clsProjekt = _allProjects.Item(myElem)

            If myProject.projectType = ptPRPFType.portfolio Then
                ' hier müssen die Projekte eingetragen werden, die in der entsprechenden Constellation verzeichnet sind ... 
                Try
                    ' das Element selber eintragen ...
                    If Not mySortedList.ContainsKey(myElem) Then
                        mySortedList.Add(myElem, True)
                    End If

                    Dim curConstellation As clsConstellation = projectConstellations.getConstellation(myElem)

                    If Not IsNothing(curConstellation) Then
                        Dim teilergebnisListe As SortedList(Of String, Boolean) = curConstellation.getBasicProjectIDs

                        For Each teKvP As KeyValuePair(Of String, Boolean) In teilergebnisListe
                            If mySortedList.ContainsKey(teKvP.Key) Then
                                ' nichts tun, ist schon drin 
                            Else
                                mySortedList.Add(teKvP.Key, teKvP.Value)
                            End If

                        Next
                    End If

                Catch ex As Exception

                End Try


            Else
                ' einfach nur den pvname eintragen 
                If Not mySortedList.ContainsKey(myElem) Then
                    mySortedList.Add(myElem, True)
                End If
            End If

        End If

        ' und jetzt kommt die Prüfung ..
        Dim checkList1 As SortedList(Of String, Boolean)
        Dim checklist2 As SortedList(Of String, Boolean)

        If mySortedList.Count < sortedListInQuestion.Count Then
            checkList1 = mySortedList
            checklist2 = sortedListInQuestion
        Else
            checkList1 = sortedListInQuestion
            checklist2 = mySortedList
        End If

        For Each checkKvP As KeyValuePair(Of String, Boolean) In checkList1
            If checklist2.ContainsKey(checkKvP.Key) Then
                tmpResult = True
                Exit For
            End If
        Next

        elemHasConflictsWith = tmpResult
    End Function


    ''' <summary>
    ''' gibt true zurück, wenn es in clsAlleProjekte-Instanz und Constellation gemeinsame Projekte gibt 
    ''' </summary>
    ''' <param name="pvName">der Name des neuen Objekts, Projekt oder Summary Projekt </param>
    ''' <param name="isConstellation">gibt an , ob es sich um ein Summary Projekt / Constellation handelt </param>
    ''' <returns></returns>
    Public Function hasAnyConflictsWith(ByVal pvName As String, ByVal isConstellation As Boolean) As Boolean

        Dim tmpResult As Boolean = False

        Dim sortedListSession As New SortedList(Of String, Boolean)
        Dim sortedListInQuestion As New SortedList(Of String, Boolean)


        ' alles untersuchen 
        ' Aufbau aller in AlleProjekte referenzierten PRojekte und Summary Projekte 
        For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

            If kvp.Value.projectType = ptPRPFType.portfolio Then
                ' hier müssen die Projekte eingetragen werden, die in der entsprechenden Constellation verzeichnet sind ... 
                Try
                    ' das Element selber eintragen ...
                    If Not sortedListSession.ContainsKey(kvp.Key) Then
                        sortedListSession.Add(kvp.Key, True)
                    End If

                    Dim curConstellation As clsConstellation = projectConstellations.getConstellation(kvp.Value.name)
                    Dim teilergebnisListe As SortedList(Of String, Boolean) = curConstellation.getBasicProjectIDs

                    For Each teKvP As KeyValuePair(Of String, Boolean) In teilergebnisListe
                        If sortedListSession.ContainsKey(teKvP.Key) Then
                            ' nichts tun, ist schon drin 
                        Else
                            sortedListSession.Add(teKvP.Key, teKvP.Value)
                        End If

                    Next
                Catch ex As Exception

                End Try


            Else
                ' einfach nur den pvname eintragen 
                If Not sortedListSession.ContainsKey(kvp.Key) Then
                    sortedListSession.Add(kvp.Key, True)
                End If
            End If
        Next


        ' Aufbau der inQuestion Sorted Liste 
        If isConstellation Then
            Dim tmpconstellation As clsConstellation = projectConstellations.getConstellation(pvName)
            Dim summaryName As String = calcProjektKey(pvName, "")

            If Not sortedListInQuestion.ContainsKey(summaryName) Then
                sortedListInQuestion.Add(summaryName, True)
            End If

            If Not IsNothing(tmpconstellation) Then
                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In tmpconstellation.Liste
                    ' tk 28.12. reasontoExclude wurde umbenannt / umgewidmet in projectTyp 
                    If kvp.Value.projectTyp = ptPRPFType.portfolio.ToString Then
                        Try
                            If sortedListInQuestion.ContainsKey(kvp.Key) Then
                                ' nichts tun, ist schon drin 
                            Else
                                sortedListInQuestion.Add(kvp.Key, kvp.Value.show)
                            End If
                            Dim teilErgebnisListe As SortedList(Of String, Boolean) = tmpconstellation.getBasicProjectIDs

                            For Each teKvP As KeyValuePair(Of String, Boolean) In teilErgebnisListe
                                If sortedListInQuestion.ContainsKey(teKvP.Key) Then
                                    ' nichts tun, ist schon drin 
                                Else
                                    sortedListInQuestion.Add(teKvP.Key, teKvP.Value)
                                End If

                            Next
                        Catch ex As Exception

                        End Try
                    Else
                        If Not sortedListInQuestion.ContainsKey(kvp.Key) Then
                            sortedListInQuestion.Add(kvp.Key, kvp.Value.show)
                        End If

                    End If
                Next
            End If


        Else
            sortedListInQuestion.Add(pvName, True)
        End If

        ' und jetzt kommt die Prüfung ..
        Dim checkList1 As SortedList(Of String, Boolean)
        Dim checklist2 As SortedList(Of String, Boolean)

        If sortedListSession.Count < sortedListInQuestion.Count Then
            checkList1 = sortedListSession
            checklist2 = sortedListInQuestion
        Else
            checkList1 = sortedListInQuestion
            checklist2 = sortedListSession
        End If

        For Each checkKvP As KeyValuePair(Of String, Boolean) In checkList1
            If checklist2.ContainsKey(checkKvP.Key) Then
                tmpResult = True
                Exit For
            End If
        Next

        hasAnyConflictsWith = tmpResult
    End Function

    ''' <summary>
    ''' fügt der Sorted List ein Projekt-Element mit Schlüssel key hinzu 
    ''' in jedem clsPRojekteAlle Aufruf soll updateCurrentConstellation by default immer auf False sein 
    ''' später soll die Aufrufleiste bereinigt werden ... 
    ''' checkOnConflicts wird benötigt, um zu entscheiden, ob ein Summary Projekt-Konflkikt vorliegt. 
    ''' Eigentlich sollen bei allen AlleProjekte.add Aufrufen der checkOn auf true gesetzt sein 
    ''' wenn der Schlüssel bereits existiert, wird eine Argument-Exception geworfen 
    ''' </summary>
    ''' <param name="project"></param>
    ''' <param name="updateCurrentConstellation">soll die currentConstellation aktualisiert werden; nur bei AlleProjekte</param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal project As clsProjekt,
                   Optional ByVal updateCurrentConstellation As Boolean = True,
                   Optional ByVal sortkey As Integer = -1,
                   Optional ByVal checkOnConflicts As Boolean = False)


        Dim keyReal As String = calcProjektKey(project.name, project.variantName)
        Dim pKey As String = calcProjektKey(project)

        If Not IsNothing(project) Then
            ' jetzt muss geprüft werden, ob es sich bei dem neuen um ein Union Projekt handelt ...
            If checkOnConflicts Then

                If project.projectType = ptPRPFType.portfolio Then

                    Dim myConstellation As clsConstellation = projectConstellations.getConstellation(project.name)

                    If Not IsNothing(myConstellation) Then
                        Dim deleteCollection As New Collection


                        Dim sortListInQuestion As SortedList(Of String, Boolean) = myConstellation.getBasicProjectIDs
                        If Not sortListInQuestion.ContainsKey(pKey) Then
                            sortListInQuestion.Add(pKey, True)
                        End If

                        For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                            If elemHasConflictsWith(kvp.Key, sortListInQuestion) Then
                                deleteCollection.Add(kvp.Key)
                            End If
                        Next

                        ' jetzt muss ggf die komplette deleteCollection durchgegangen werden 
                        For Each item As String In deleteCollection
                            Me.Remove(item, True)
                        Next

                    End If

                Else

                    If Me.hasAnyConflictsWith(pKey, False) Then
                        Throw New ArgumentException("Summary Projekt Konflikt: " & project.name)
                    End If

                End If
            End If


            ' existiert es bereits ? 
            ' wenn ja, dann löschen ...
            If _allProjects.ContainsKey(keyReal) Then
                _allProjects.Remove(keyReal)
            End If
            _allProjects.Add(keyReal, project)

            ' 21.3.17
            ' soll die currentConstellation upgedated werden ? 
            If updateCurrentConstellation Then
                Dim cItem As New clsConstellationItem
                With cItem
                    .projectName = project.name
                    .variantName = project.variantName

                    If sortkey >= 2 Then
                        .zeile = sortkey
                    End If

                    .projectTyp = CType(project.projectType, ptPRPFType).ToString
                End With
                currentSessionConstellation.add(cItem, sKey:=sortkey)
            End If
        End If


    End Sub


    ''' <summary>
    ''' macht einen Update, wenn das Element mit Schlüssel key bereits existiert 
    ''' macht einen Insert, wenn das Element mit Schlüssel key noch nicht existiert 
    ''' der key wird bestimmt aus project.name und .variantname
    ''' </summary>
    ''' <param name="project"></param>
    ''' <remarks>wenn project Nothing ist, dann bleibt die Liste unverändert </remarks>
    Public Sub upsert(ByVal project As clsProjekt)

        If Not IsNothing(project) Then
            Dim key As String = calcProjektKey(project.name, project.variantName)
            If _allProjects.ContainsKey(key) Then
                _allProjects.Remove(key)
            End If
            _allProjects.Add(key, project)
        End If

    End Sub

    ''' <summary>
    ''' gibt die Liste der pvNames in der Klassen-Instanz zurück 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getPvNameListe() As Collection
        Get
            Dim tmpResult As New Collection
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                If Not tmpResult.Contains(kvp.Key) Then
                    tmpResult.Add(kvp.Key, kvp.Key)
                End If
            Next

            getPvNameListe = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gets or sets the sortedlist of (string, clsprojekt)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property liste() As SortedList(Of String, clsProjekt)
        Get
            liste = _allProjects
        End Get

        Set(value As SortedList(Of String, clsProjekt))

            If Not IsNothing(value) Then
                _allProjects = value
            End If

        End Set

    End Property

    ''' <summary>
    ''' true, wenn die SortedList ein Element mit angegebenem Key enthält
    ''' false, sonst
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Containskey(ByVal key As String) As Boolean
        Get
            Containskey = _allProjects.ContainsKey(key)
        End Get
    End Property



    ''' <summary>
    ''' gibt die Anzahl Listenelemente der Sorted Liste zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count() As Integer
        Get
            Count = _allProjects.Count
        End Get
    End Property



    ''' <summary>
    ''' gibt das erste Element der Liste zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property First() As clsProjekt
        Get
            If _allProjects.Count > 0 Then
                First = _allProjects.First.Value
            Else
                First = Nothing
            End If
        End Get
    End Property


    ''' <summary>
    ''' gibt eine Liste der vorkommenden Meilenstein Namen in der Menge von Projekte zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneNames() As Collection

        Get

            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getMilestoneNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next

            getMilestoneNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste der vorkommenden Meilenstein Namen in der Menge von Projekte zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneCategoryNames() As Collection

        Get

            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getMilestoneCategoryNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next

            getMilestoneCategoryNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' gibt die Liste der vorkommenden Phasen-Namen in der Menge der Projekte an ...  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseNames() As Collection

        Get

            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getPhaseNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next


            getPhaseNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' gibt die Liste der vorkommenden Phasen-Namen in der Menge der Projekte an ...  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseCategoryNames() As Collection

        Get

            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getPhaseCategoryNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next


            getPhaseCategoryNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Rollen, die in der Menge von Projekten vorkommen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleNames() As Collection
        Get
            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getRoleNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next


            getRoleNames = tmpListe
        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Kostenarten, die in der Menge von Projekten vorkommen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostNames() As Collection
        Get
            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getCostNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next

            getCostNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Business Units, die in der Menge von Projekten vorkommen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBUNames() As Collection
        Get
            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpBU As String = kvp.Value.businessUnit
                If Not IsNothing(tmpBU) Then
                    If tmpBU.Trim.Length > 0 Then
                        If Not tmpListe.Contains(tmpBU) Then
                            tmpListe.Add(tmpBU, tmpBU)
                        End If
                    End If
                End If

            Next

            getBUNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Projektvorlagen, die in der Menge von Projekten vorkommen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTypNames() As Collection
        Get
            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpTyp As String = kvp.Value.VorlagenName
                If Not IsNothing(tmpTyp) Then
                    If tmpTyp.Trim.Length > 0 Then
                        If Not tmpListe.Contains(tmpTyp) Then
                            tmpListe.Add(tmpTyp, tmpTyp)
                        End If
                    End If
                End If

            Next

            getTypNames = tmpListe

        End Get
    End Property
    ''' <summary>
    ''' gibt die Namen der existierenden Varianten in einer Liste zurück 
    ''' die "leere" Variante wird als () zurückgegeben , alle anderen Varianten als (Variante-Name)
    ''' Voraussetzung: _allprojects ist eine sortierte Liste
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getVariantNames(ByVal pName As String, ByVal mitKlammer As Boolean) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim i As Integer = 0
            Dim found As Boolean = False
            Dim vName As String

            ' Positioniere i auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    found = True
                Else
                    i = i + 1
                End If
            End While

            ' Schreibe alle Varianten in die Ergebnis-Liste tmpCollection
            While i < _allProjects.Count And found

                If _allProjects.ElementAt(i).Value.name = pName Then

                    If mitKlammer Then
                        vName = "(" & _allProjects.ElementAt(i).Value.variantName & ")"
                    Else
                        vName = _allProjects.ElementAt(i).Value.variantName
                    End If

                    tmpCollection.Add(vName)
                    i = i + 1
                Else
                    found = False
                End If

            End While

            getVariantNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt das kleinste Start-Datum zurück, das alle Varianten des Projektes haben 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMinDate(ByVal pName As String) As Date
        Get
            Dim tmpDate As Date = StartofCalendar
            Dim i As Integer = 0
            Dim found As Boolean = False


            ' Positioniere i auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    tmpDate = _allProjects.ElementAt(i).Value.startDate
                    found = True
                    i = i + 1
                Else
                    i = i + 1
                End If
            End While

            ' ist ein Datum einer weiteren Variante kleiner ? 


            While i < _allProjects.Count And found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    If DateDiff(DateInterval.Day, tmpDate, _allProjects.ElementAt(i).Value.startDate) < 0 Then
                        tmpDate = _allProjects.ElementAt(i).Value.startDate
                    End If
                    i = i + 1
                Else
                    found = False
                End If

            End While

            getMinDate = tmpDate

        End Get
    End Property


    ''' <summary>
    ''' gibt das größte Ende-Datum zurück, das alle Varianten des Projekts haben
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMaxDate(ByVal pName As String) As Date
        Get

            Dim tmpDate As Date = StartofCalendar.AddMonths(240)
            Dim i As Integer = 0
            Dim found As Boolean = False


            ' Positioniere i auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    tmpDate = _allProjects.ElementAt(i).Value.endeDate
                    found = True
                    i = i + 1
                Else
                    i = i + 1
                End If
            End While

            ' ist ein Datum einer weiteren Variante größer ? 


            While i < _allProjects.Count And found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    If DateDiff(DateInterval.Day, tmpDate, _allProjects.ElementAt(i).Value.endeDate) > 0 Then
                        tmpDate = _allProjects.ElementAt(i).Value.endeDate
                    End If
                    i = i + 1
                Else
                    found = False
                End If

            End While

            getMaxDate = tmpDate


        End Get
    End Property

    ''' <summary>
    ''' gibt das Element zurück, das den pName, vName als Projekt- bzw. Varianten-NAme enthält
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(ByVal pName As String, ByVal vName As String) As clsProjekt
        Get
            Dim key As String = calcProjektKey(pName, vName)
            If _allProjects.ContainsKey(key) Then
                getProject = _allProjects(key)
            Else
                getProject = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt das Element zurück, das den angegebenen Schlüssel key enthält
    ''' </summary>
    ''' <param name="key">key = pName#vName</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(ByVal key As String) As clsProjekt
        Get

            If _allProjects.ContainsKey(key) Then
                getProject = _allProjects(key)
            Else
                getProject = Nothing
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt das Projekt zurück, das den angegebenen Schlüssel kdNr enthält
    ''' </summary>
    ''' <param name="kdNr">kdNr = kundenNummer</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProjectByKDNr(ByVal kdNr As String) As clsProjekt
        Get
            Dim tmpResult As clsProjekt = Nothing

            If Not IsNothing(kdNr) Then

                If kdNr <> "" Then
                    For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                        If Not IsNothing(kvp.Value.kundenNummer) Then
                            If kvp.Value.kundenNummer = kdNr Then
                                tmpResult = kvp.Value
                                Exit For
                            End If
                        End If

                    Next
                End If

            End If

            getProjectByKDNr = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt die entsprechende bezeichnete Variante zurück
    ''' VariantNummer = 0 => 1. Projekt-Vorkommen, meist mit Varianten-Namen "" 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="variantNummer"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(ByVal pName As String, ByVal variantNummer As Integer) As clsProjekt
        Get


            Dim i As Integer = 0
            Dim found As Boolean = False

            ' Positioniere position auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found

                If _allProjects.ElementAt(i).Value.name = pName Then
                    found = True
                Else
                    i = i + 1
                End If



            End While


            If found Then
                getProject = _allProjects.ElementAt(i + variantNummer).Value
            Else
                getProject = Nothing
            End If


        End Get
    End Property




    ''' <summary>
    ''' gibt die Anzahl Varianten für den übergebenen pName an 
    ''' Das Projekt mit variantName = "" zählt dabei nicht als Variante 
    ''' es gibt nur das Projekt mit Variante "": 0
    ''' es gibt nicht einmal das Projekt mit Namen pName: -1
    ''' Anzahl Varianten mit variantName ungleich "": sonst
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getVariantZahl(ByVal pName As String) As Integer
        Get
            Dim anzahl As Integer = 0
            Dim i As Integer = 0
            Dim found As Boolean = False

            ' Positioniere i auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    found = True
                    anzahl = anzahl + 1
                End If
                i = i + 1

            End While

            ' zähle alle weiteren Vorkommnisse
            While i < _allProjects.Count And found

                If _allProjects.ElementAt(i).Value.name = pName Then
                    anzahl = anzahl + 1
                Else
                    found = False
                End If

                i = i + 1
            End While

            getVariantZahl = anzahl - 1

        End Get
    End Property

    ''' <summary>
    ''' gibt die Liste der unterschiedlichen Projekt-Namen zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProjectNames() As Collection
        Get
            Dim tmpCollection As New Collection
            Dim pName As String

            For i As Integer = 0 To Me.Count - 1
                pName = _allProjects.ElementAt(i).Value.name
                If Not tmpCollection.Contains(pName) Then
                    tmpCollection.Add(pName, pName)
                End If
            Next

            getProjectNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' entfernt das Element mit Schlüssel "Key" aus der Sorted List
    ''' es wird - im Falle AlleProjekte (Aufruf-Schnittstelle beachten) auch die currentConstellation aktualisiert 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Sub Remove(ByVal key As String, Optional ByVal updateCurrentConstellation As Boolean = True)

        Try
            If updateCurrentConstellation Then
                currentSessionConstellation.remove(key)
            End If

            If _allProjects.ContainsKey(key) Then
                _allProjects.Remove(key)
            End If
        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' entfernt alle Projekt-Varianten mit ProjektNamen = pName
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <remarks></remarks>
    Public Sub RemoveAllVariantsOf(ByVal pName As String, Optional ByVal updateCurrentConstellation As Boolean = True)

        Dim i As Integer = 0
        Dim found As Boolean = False

        ' Positioniere i auf das erste Vorkommen von pName in der Liste 
        While i < _allProjects.Count And Not found
            If _allProjects.ElementAt(i).Value.name = pName Then
                found = True
            Else
                i = i + 1
            End If
        End While

        ' Lösche alle Varianten mit ProjektName = pName 
        While found

            If i < _allProjects.Count Then

                If _allProjects.ElementAt(i).Value.name = pName Then
                    ' jetzt die currentConstellation aktualisieren 
                    Dim key As String = _allProjects.ElementAt(i).Key
                    If updateCurrentConstellation Then
                        currentSessionConstellation.remove(key)
                    End If

                    _allProjects.RemoveAt(i)
                Else
                    found = False
                End If

            Else
                found = False
            End If

        End While

    End Sub

    ''' <summary>
    ''' setzt die Liste der Projekte zurück 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear(Optional ByVal updateCurrentConstellation As Boolean = True)

        _allProjects.Clear()
        ' die currentSessionConstellation neu aufsetzen
        ' dei bekommt damit den Namen last<dbUSerName>

        If updateCurrentConstellation Then
            currentSessionConstellation = New clsConstellation
        End If


    End Sub

End Class
