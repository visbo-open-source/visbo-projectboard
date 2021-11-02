
Imports xlNS = Microsoft.Office.Interop.Excel

Public Class clsProjekte
    ' in dieser Klasse ist der Key immer nur der ProjektName pname (andere Variante kann nicht in ShowProjekte enthalten sein)
    Private _allProjects As SortedList(Of String, clsProjekt)
    Private _allShapes As SortedList(Of String, String)
    Private _allCoord As SortedList(Of String, Double())

    ''' <summary>
    ''' trägt ein Projekt mit dem Schlüssel Projekt-NAme in die Liste ein 
    ''' trägt die Shape ID (shpUID) in die Shape Liste ein 
    ''' wenn der Projekt-Name bereits existiert, wird eine Exception geworfen 
    ''' </summary>
    ''' <param name="project"></param>
    ''' <remarks></remarks>
    Public Sub Add(project As clsProjekt, Optional ByVal updateCurrentConstellation As Boolean = True)

        Try
            If Not IsNothing(project) Then

                Dim pname As String = project.name
                Dim shpUID As String = project.shpUID

                _allProjects.Add(pname, project)

                If shpUID <> "" Then
                    _allShapes.Add(shpUID, pname)
                End If

                If updateCurrentConstellation Then
                    currentConstellationPvName = calcLastSessionScenarioName()

                    Dim key As String = calcProjektKey(project)
                    If currentSessionConstellation.contains(key, False) Then
                        Call currentSessionConstellation.setItemToShow(key, True)
                    End If

                End If

            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try


    End Sub

    ''' <summary>
    ''' liefert das kleinste auftretende actualDatauntil zurück 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property actualDataUntil() As Date
        Get
            Dim tmpResult As Date = Date.Now

            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                Dim currentUntilDate As Date = kvp.Value.actualDataUntil

                If currentUntilDate < tmpResult Then
                    ' nur dann machen, wenn das Projekt nicht gestoppt ist und nicht bereits beendet 
                    If DateDiff(DateInterval.Month, currentUntilDate, kvp.Value.endeDate) > 0 Then
                        tmpResult = kvp.Value.actualDataUntil
                    End If
                End If

            Next

            actualDataUntil = tmpResult
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
    ''' gibt true zurück, wenn es in ShowProjekte-Instanz und Constellation gemeinsame Projekte gibt 
    ''' anders als n AlleProjekte.hasAnyConflictsWith tritt hier bereits ein Konflikt auf, wenn der pName gleich ist; 
    ''' in ShowProjekte darf von jedem Projekt nur höchstens eine Variante sein. 
    ''' </summary>
    ''' <param name="pvName">der Name des neuen Objekts, Projekt oder Summary Projekt </param>
    ''' <param name="isConstellation">gibt an , ob es sich um ein Summary Projekt / Constellation handelt </param>
    ''' <returns></returns>
    Public Function hasAnyConflictsWith(ByVal pvName As String, ByVal isConstellation As Boolean) As Boolean

        Dim tmpResult As Boolean = False

        Dim sortedListSession As New SortedList(Of String, Boolean)
        Dim sortedListInQuestion As New SortedList(Of String, Boolean)


        ' alles untersuchen 
        ' Aufbau aller in ShowProjekte referenzierten PRojekte und Summary Projekte 
        For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

            If kvp.Value.projectType = ptPRPFType.portfolio Then
                ' hier müssen die Projekte mit ihrem pName eingetragen werden, die in der entsprechenden Constellation verzeichnet sind ... 
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
            ' anders als in AlleProjekte-Methonde nur den Namen verwenden ..
            'Dim summaryName As String = calcProjektKey(pvName, "")
            Dim summaryName As String = pvName

            ' tk 22.7.19 wenn loadPFV , dann muss der Varianten

            If Not sortedListInQuestion.ContainsKey(summaryName) Then
                sortedListInQuestion.Add(summaryName, True)
            End If

            If Not IsNothing(tmpconstellation) Then
                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In tmpconstellation.Liste
                    ' tk 28.12. reasontoExclude wurde umbenannt / umgewidmet in projectTyp 
                    If kvp.Value.projectTyp = ptPRPFType.portfolio.ToString Then
                        Try
                            If sortedListInQuestion.ContainsKey(kvp.Value.projectName) Then
                                ' nichts tun, ist schon drin 
                            Else
                                sortedListInQuestion.Add(kvp.Value.projectName, kvp.Value.show)
                            End If
                            Dim teilErgebnisListe As SortedList(Of String, Boolean) = tmpconstellation.getBasicProjectIDs

                            For Each teKvP As KeyValuePair(Of String, Boolean) In teilErgebnisListe
                                Dim pName As String = getPnameFromKey(teKvP.Key)
                                If sortedListInQuestion.ContainsKey(pName) Then
                                    ' nichts tun, ist schon drin 
                                Else
                                    sortedListInQuestion.Add(pName, teKvP.Value)
                                End If

                            Next
                        Catch ex As Exception

                        End Try
                    Else
                        Dim pName As String = kvp.Value.projectName
                        If Not sortedListInQuestion.ContainsKey(pName) Then
                            sortedListInQuestion.Add(pName, kvp.Value.show)
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
    ''' trägt die Zuordnung Shape/Projekt in die AllShape Liste ein 
    ''' Fehler, wenn pname gar nicht in der AllProjects Liste ist 
    ''' </summary>
    ''' <param name="pname">Name / Key des Projekts</param>
    ''' <param name="shpUID">Key des Shpelements</param>
    ''' <remarks></remarks>
    Public Sub AddShape(pname As String, shpUID As String)


        If _allProjects.ContainsKey(pname) Then
            Try
                If _allShapes.ContainsValue(pname) Then
                    Dim ix As Integer
                    ix = _allShapes.IndexOfValue(pname)
                    _allShapes.RemoveAt(ix)
                End If
                _allShapes.Add(shpUID, pname)

            Catch ex As Exception
                Throw New ArgumentException(ex.Message)
            End Try
        Else
            Throw New ArgumentException("Shape kann nicht einem nicht-existierenden Projekt hinzugefügt werden - ")
        End If



    End Sub

    ''' <summary>
    ''' gibt die Zeile zurück, in der das Projekt auf der Projekt-Tafel gezeichnet werden soll 
    ''' aktuell ist das die alphabetische Reihenfolge
    ''' das muss später noch angepasst werden ... 
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPTZeile(ByVal projectName As String) As Integer
        Get
            ' wurde am 21.3.17 ersetzt durch das unten folgende 
            ''If _allProjects.ContainsKey(projectName) Then
            ''    getPTZeile = _allProjects.IndexOfKey(projectName) + 2
            ''Else
            ''    getPTZeile = 0
            ''End If

            ' seit 21.3.17
            getPTZeile = currentSessionConstellation.getBoardZeile(projectName)

        End Get
    End Property

    ''' <summary>
    ''' nimmt das Projekt mit dem übergebenen Namen aus der Liste heraus  
    ''' wirft eine Exception, wenn Projekt nicht ind er Liste oder ShpUID nicht ind er zugehörigen Shape-Liste
    ''' aktualisiert die currentSessionConstellation
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <remarks></remarks>
    Public Sub Remove(projectname As String, Optional ByVal updateCurrentConstellation As Boolean = True)

        Try
            Dim vname As String = ""

            If _allProjects.ContainsKey(projectname) Then
                Dim SID As String = _allProjects.Item(projectname).shpUID
                vname = _allProjects.Item(projectname).variantName
                _allProjects.Remove(projectname)
                If SID <> "" Then
                    _allShapes.Remove(SID)
                End If
            End If

            If updateCurrentConstellation Then
                Dim key As String = calcProjektKey(projectname, vName)

                If currentSessionConstellation.contains(key, False) Then
                    Call currentSessionConstellation.setItemToShow(key, False)
                End If

            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try



    End Sub

    ''' <summary>
    ''' nimmt das Projekt mit der übergebenen Shape UID aus der Liste der Projekte und der Liste der Shapes heraus
    ''' wirft Exception, wenn Fehler 
    ''' </summary>
    ''' <param name="SID"></param>
    ''' <remarks></remarks>
    Public Sub RemoveS(SID As String)

        Try
            Dim pname As String = _allShapes.Item(SID)
            _allProjects.Remove(pname)
            _allShapes.Remove(SID)

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try



    End Sub

    ''' <summary>
    ''' setzt die Liste der Projekte und die Liste der Shapes zurück 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear(Optional ByVal updateCurrentConstellation As Boolean = True)

        _allProjects.Clear()
        _allShapes.Clear()

        ' jetzt die currentConstellation, alle Items auf noShow setzen 
        If updateCurrentConstellation Then
            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In currentSessionConstellation.Liste
                kvp.Value.show = False
            Next
        End If
        

    End Sub

    ''' <summary>
    ''' gibt an, ob die Liste den angegebenen Schlüssel enthält oder nicht 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property contains(ByVal key As String) As Boolean
        Get

            If IsNothing(key) Then
                contains = False
            Else
                If _allProjects.ContainsKey(key) Then
                    contains = True
                Else
                    contains = False
                End If
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt in der Ergebnis Collection alle Kind Namen von Phasen zurück 
    ''' </summary>
    ''' <param name="phaseName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhasesOfPhase(ByVal phaseName As String, _
                                                  Optional ByVal breadcrumb As String = "") As Collection
        Get

            Dim tmpCollection As Collection = New Collection
            Dim zwischenresult As Collection = New Collection
            Dim phaseIndices() As Integer
            Dim elemID As String
            Dim elemName As String
            Dim childID As String
            Dim curNode As clsHierarchyNode
            Dim cphase As clsPhase


            If Not IsNothing(phaseName) Then
                If phaseName.Trim.Length > 0 Then
                    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                        phaseIndices = kvp.Value.hierarchy.getPhaseIndices(phaseName, breadcrumb)
                        For px As Integer = 0 To phaseIndices.Length - 1

                            cphase = kvp.Value.getPhase(phaseIndices(px))
                            If Not IsNothing(cphase) Then

                                elemID = cphase.nameID
                                curNode = kvp.Value.hierarchy.nodeItem(elemID)

                                If Not IsNothing(curNode) Then
                                    For ix As Integer = 1 To curNode.childCount
                                        childID = curNode.getChild(ix)
                                        If Not elemIDIstMeilenstein(childID) Then
                                            elemName = elemNameOfElemID(childID)
                                            If Not tmpCollection.Contains(elemName) And elemName.Trim.Length > 0 Then
                                                tmpCollection.Add(elemName, elemName)
                                            End If
                                        End If
                                    Next
                                End If

                            End If

                        Next

                    Next
                End If
            End If

            getPhasesOfPhase = tmpCollection

        End Get
    End Property


    ''' <summary>
    ''' gibt die Sammlung von Meilensteinen zurück, die eine Phase in irgendeinem Projekt hat  
    ''' </summary>
    ''' <param name="phaseName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestonesOfPhase(ByVal phaseName As String, _
                                                        Optional ByVal breadcrumb As String = "") As Collection
        Get
            Dim tmpCollection As Collection = New Collection
            Dim zwischenresult As Collection = New Collection
            Dim phaseIndices() As Integer
            Dim elemID As String

            Dim cphase As clsPhase


            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                phaseIndices = kvp.Value.hierarchy.getPhaseIndices(phaseName, breadcrumb)
                For px As Integer = 0 To phaseIndices.Length - 1

                    cphase = kvp.Value.getPhase(phaseIndices(px))
                    If Not IsNothing(cphase) Then

                        For mx As Integer = 1 To cphase.countMilestones

                            Dim cMilestone As clsMeilenstein = cphase.getMilestone(mx)
                            If Not IsNothing(cMilestone) Then
                                If Not tmpCollection.Contains(cMilestone.name) Then
                                    tmpCollection.Add(cMilestone.name, cMilestone.name)
                                End If
                            End If

                        Next

                    End If

                    ' Übertragen der Ergebnisse in zwischen result
                    For i As Integer = 1 To tmpCollection.Count
                        Dim newItem As String = CStr(tmpCollection.Item(i))
                        If Not zwischenresult.Contains(newItem) Then
                            zwischenresult.Add(newItem, newItem)
                        End If

                    Next
                    tmpCollection.Clear()

                    ' jetzt müssen alle Kind-Phasen des Elements bearbeitet werden  
                    Dim anzahlChilds As Integer
                    Try
                        Dim childNode As clsHierarchyNode
                        childNode = kvp.Value.hierarchy.nodeItem(cphase.nameID)
                        anzahlChilds = childNode.childCount
                        For cx = 1 To anzahlChilds
                            elemID = childNode.getChild(cx)
                            If Not elemIDIstMeilenstein(elemID) Then
                                tmpCollection = getMilestonesOfPhase(elemID)
                            End If

                            ' Übertragen der Ergebnisse in zwischen result
                            For i As Integer = 1 To tmpCollection.Count
                                Dim newItem As String = CStr(tmpCollection.Item(i))
                                If Not zwischenresult.Contains(newItem) Then
                                    zwischenresult.Add(newItem, newItem)
                                End If

                            Next
                            tmpCollection.Clear()

                        Next
                    Catch ex As Exception

                    End Try

                Next

            Next

            getMilestonesOfPhase = zwischenresult

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

                'Dim tmpCollection As Collection = kvp.Value.getMilestoneNames
                Dim tmpCollection As Collection
                If awinSettings.considerCategories Then
                    tmpCollection = kvp.Value.getMilestoneCategoryNames
                Else
                    tmpCollection = kvp.Value.getMilestoneNames
                End If

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
    ''' gibt eine Liste der vorkommenden Meilenstein Klassen in der Menge von Projekte zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneCategoryNames() As Collection

        Get

            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                'Dim tmpCollection As Collection = kvp.Value.getMilestoneNames
                Dim tmpCollection As Collection
                tmpCollection = kvp.Value.getMilestoneCategoryNames

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
    ''' gibt die Liste der vorkommenden Phasen-KlassenNamen in der Menge der Projekte an ...  
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
    ''' liefert die Namen der Rollen, die in der Menge von Projekten vorkommen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleNames(Optional ByVal includingParentRoles As Boolean = False) As Collection
        Get
            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getRoleNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                        If includingParentRoles Then
                            Dim tmprole As clsRollenDefinition = RoleDefinitions.getRoledef(tmpName)
                            Dim parentRole As clsRollenDefinition = RoleDefinitions.getParentRoleOf(tmprole.UID)
                            Dim grandparentRole As clsRollenDefinition = Nothing
                            If Not IsNothing(parentRole) Then
                                If Not tmpListe.Contains(parentRole.name) Then
                                    tmpListe.Add(parentRole.name, parentRole.name)
                                    grandparentRole = RoleDefinitions.getParentRoleOf(parentRole.UID)
                                    Do While Not IsNothing(grandparentRole)
                                        If Not tmpListe.Contains(grandparentRole.name) Then
                                            tmpListe.Add(grandparentRole.name, grandparentRole.name)
                                            grandparentRole = RoleDefinitions.getParentRoleOf(grandparentRole.UID)
                                        Else
                                            grandparentRole = Nothing
                                        End If
                                    Loop
                                End If
                            End If
                        End If
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
    ''' gibt die nach Namen sortierte Liste von Projekten zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Liste() As SortedList(Of String, clsProjekt)
        Get
            Liste = _allProjects
        End Get
        'Set(value As SortedList(Of String, clsProjekt))
        '    AllProjects = value
        'End Set
    End Property

    ''' <summary>
    ''' gibt die Anzahl der Projekte in der Liste an 
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
    ''' gibt das Element an der Stelle mit Index zurück; das 1. Element hat den Index 1
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(index As Integer) As clsProjekt
        Get

            If index >= 1 And index <= _allProjects.Count Then
                getProject = _allProjects.ElementAt(index - 1).Value
            Else
                getProject = Nothing
            End If
            ' Änderung tk 6.12.15 ein Get sollte keine Exception werfen, nur Nothing zurückgeben 
            'Try
            '    getProject = AllProjects.ElementAt(index - 1).Value
            'Catch ex As Exception
            '    Throw New ArgumentException("Index nicht vorhanden:" & index.ToString)
            'End Try
        End Get
    End Property


    ''' <summary>
    ''' gibt das Shape Element zurück, das zum Projekt gehört
    ''' </summary>
    ''' <param name="pName">Name des Projektes 
    ''' (ist auch gleichzeitig der NAme des Shapes)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShape(ByVal pName As String) As xlNS.Shape
        Get
            Dim shapes As xlNS.Shapes
            Dim projectShape As xlNS.ShapeRange


            'With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), xlNS.Worksheet)
            With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), xlNS.Worksheet)
                shapes = .Shapes
                Try
                    projectShape = shapes.Range(pName)
                    getShape = projectShape.Item(1)
                Catch ex As Exception
                    getShape = Nothing
                End Try
            End With


        End Get
    End Property

    ''' <summary>
    ''' gibt die Namen der Projekte zurück, die "markiert" sind 
    ''' leere Collection, wenn es keine gibt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMarkedProjects() As Collection
        Get
            Dim tmpCollection As New Collection
            For Each kvp As KeyValuePair(Of String, clsProjekt) In Me.Liste
                If kvp.Value.marker = True Then
                    If Not tmpCollection.Contains(kvp.Key) Then
                        tmpCollection.Add(kvp.Key, kvp.Key)
                    End If
                End If
            Next

            getMarkedProjects = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' bestimmt die kleinste auftretende Spalten-Column über alle Projekte  
    ''' wenn eine liste angegeben ist, werden nur die in der Liste vorhandenen PRoekte betrachtet 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMinMonthColumn(Optional ByVal liste As Collection = Nothing) As Integer
        Get
            Dim tmpMin As Integer = 10000
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                If IsNothing(liste) Then
                    If kvp.Value.Start < tmpMin Then
                        tmpMin = kvp.Value.Start
                    End If
                Else
                    If liste.Contains(kvp.Key) Then
                        If kvp.Value.Start < tmpMin Then
                            tmpMin = kvp.Value.Start
                        End If
                    End If
                End If
                
            Next
            getMinMonthColumn = tmpMin
        End Get
    End Property

    ''' <summary>
    ''' bestimmt die größte auftretende Spalten-Column über alle Projekte  
    ''' wenn eine liste angegeben ist, werden nur die in der Liste vorhandenen PRoekte betrachtet 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMaxMonthColumn(Optional ByVal liste As Collection = Nothing) As Integer
        Get
            Dim tmpMax As Integer = 0
            Dim endeCol As Integer
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                If IsNothing(liste) Then
                    endeCol = getColumnOfDate(kvp.Value.endeDate)
                    If endeCol > tmpMax Then
                        tmpMax = endeCol
                    End If
                Else
                    If liste.Contains(kvp.Key) Then
                        endeCol = getColumnOfDate(kvp.Value.endeDate)
                        If endeCol > tmpMax Then
                            tmpMax = endeCol
                        End If
                    End If
                End If
            Next
            getMaxMonthColumn = tmpMax
        End Get
    End Property


    ''' <summary>
    ''' gibt das vollständige Projekt aus der Liste zurück, das den angegebenen Namen hat 
    ''' </summary>
    ''' <param name="itemName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(itemName As String,
                                        Optional ByVal tryOnceMore As Boolean = False) As clsProjekt

        Get


            If _allProjects.ContainsKey(itemName) Then
                getProject = _allProjects.Item(itemName)
            ElseIf tryOnceMore Then

                Dim pName As String = extractName(itemName, PTshty.projektN)
                If pName.Length > 0 Then
                    If _allProjects.ContainsKey(pName) Then
                        getProject = _allProjects.Item(pName)
                    Else
                        Throw New ArgumentException("ProjektName " & itemName & " nicht vorhanden")
                    End If
                Else
                    Throw New ArgumentException("ProjektName " & itemName & " nicht vorhanden")
                End If
            Else
                Throw New ArgumentException("ProjektName " & itemName & " nicht vorhanden")
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
    ''' gibt die maximale Zeile auf der Projekt-Tafel zurück, die von allen angezeigten Projekten beansprucht wird  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property maxZeile() As Integer

        Get
            Dim mx As Integer = 0

            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                If kvp.Value.tfZeile > mx Then
                    mx = kvp.Value.tfZeile
                End If
            Next
            maxZeile = mx
        End Get

    End Property

    ''' <summary>
    ''' gibt das Projekt zurück, das die angegebene shpID hat. 
    ''' </summary>
    ''' <param name="shpID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProjectS(shpID As String) As clsProjekt

        Get
            Dim pname As String
            Try

                pname = _allShapes.Item(shpID)
                getProjectS = _allProjects.Item(pname)

            Catch ex As Exception
                Throw New ArgumentException("projectname nicht vorhanden")
            End Try

        End Get

    End Property

    ''' <summary>
    ''' gibt die Shape-Liste zurück: ShpID, Projekt-Name  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property shpListe() As SortedList(Of String, String)
        Get
            shpListe = _allShapes
        End Get
    End Property

    ''' <summary>
    ''' gibt eine Collection von Projekt-Namen zurück, die im Zeitraum liegen und ausserdem dem 
    ''' Selektion Kriterium genügen; aktuell ist nur "keine Einschränkung" vorgesehen
    ''' -1 - keine Einschränkung 
    ''' 
    ''' </summary>
    ''' <param name="selectionType"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property withinTimeFrame(selectionType As Integer, von As Integer, bis As Integer) As Collection
        Get
            Dim tmpListe As New Collection
            ' selection type wird aktuell noch ignoriert .... 



            For Each kvp As KeyValuePair(Of String, clsProjekt) In Me._allProjects

                With kvp.Value

                    Select Case selectionType

                        Case PTpsel.alle
                            ' Aufteilung in if .. elseif gemacht, um Geschwindigkeit zu gewinnen 
                            If bis - von < 1 Then
                                tmpListe.Add(kvp.Key, kvp.Key)
                            ElseIf Not ((getColumnOfDate(.startDate) > bis) Or (getColumnOfDate(.endeDate) < von)) Then
                                ' kein TimeFrame oder liegt innerhalb des TimeFrame ... dann wird es übernommen 
                                tmpListe.Add(kvp.Key, kvp.Key)
                            End If

                        Case PTpsel.lfundab

                            If bis - von < 1 Then
                                If DateDiff(DateInterval.Day, .startDate, Date.Now) > 0 Then
                                    tmpListe.Add(kvp.Key, kvp.Key)
                                End If
                            ElseIf Not ((getColumnOfDate(.startDate) > bis) Or (getColumnOfDate(.endeDate) < von)) Then
                                ' Projekt liegt innerhalb des TimeFrames 
                                If DateDiff(DateInterval.Day, .startDate, Date.Now) > 0 Then
                                    tmpListe.Add(kvp.Key, kvp.Key)
                                End If
                            End If

                        Case Else

                            Call MsgBox("Selektion in clsProjekte.withinTimeFrame noch nicht implementiert ")

                    End Select


                End With

            Next

            withinTimeFrame = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste von Projektnamen in der showprojekte zurück, die einen der übergebenen SelItems1 bzw. SelItems enthalten 
    ''' 
    ''' </summary>
    ''' <param name="suchTyp">0: selItems1/selitems2 = Phasen/Meilensteine</param>
    ''' <param name="selItems1">Phasen , Rollen Kosten</param>
    ''' <param name="selItems2">Meilensteine</param>
    ''' <param name="von">Start des betrachteten Zeitraums</param>
    ''' <param name="bis">Ende des betrachteten Zeitraums</param>
    ''' <value></value>
    ''' <returns>Collection of projectnames</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property withinTimeFrame(ByVal suchTyp As Integer, ByVal selItems1 As Collection, ByVal selItems2 As Collection, ByVal von As Integer, ByVal bis As Integer) As SortedList(Of Double, String)
        Get
            Dim tmpListe As New SortedList(Of Double, String)
            Dim cphase As clsPhase
            Dim cMilestone As clsMeilenstein
            Dim projektstart As Integer
            Dim found As Boolean
            Dim key As Double
            ' selection type wird aktuell noch ignoriert .... 

            suchTyp = 0

            For Each kvp As KeyValuePair(Of String, clsProjekt) In Me._allProjects

                found = False
                With kvp.Value

                    projektstart = .Start + .StartOffset

                    If (projektstart > bis) Or (projektstart + .anzahlRasterElemente - 1 < von) Then
                        ' dann liegt das Projekt ausserhalb des Zeitraums und muss überhaupt nicht berücksichtig werden 

                    Else

                        For Each fullphaseName As String In selItems1

                            Dim breadcrumb As String = ""
                            Dim phaseName As String = ""
                            Dim type As Integer = -1
                            Dim pvName As String = ""
                            Call splitHryFullnameTo2(fullphaseName, phaseName, breadcrumb, type, pvName)

                            If type = -1 Or
                                (type = PTItemType.projekt And pvName = calcProjektKey(kvp.Value)) Or
                                (type = PTItemType.vorlage) Then

                                Dim phaseIndices() As Integer = kvp.Value.hierarchy.getPhaseIndices(phaseName, breadcrumb)

                                For px As Integer = 0 To phaseIndices.Length - 1
                                    cphase = kvp.Value.getPhase(phaseIndices(px))
                                    If Not IsNothing(cphase) Then
                                        If (projektstart + cphase.relStart - 1 > bis) Or (projektstart + cphase.relEnde - 1 < von) Then
                                            ' dann liegt die Phase ausserhalb des betrachteten Zeitraums und muss nicht berücksichtigt werden 
                                        Else
                                            found = True
                                            Exit For
                                        End If
                                    End If
                                Next

                            End If

                            If found Then
                                Exit For
                            End If

                        Next

                        ' wenn noch keine Phase gefunen wurde 

                        If Not found Then
                            For Each fullmilestoneName As String In selItems2

                                Dim breadcrumb As String = ""
                                Dim milestoneName As String = ""
                                Dim type As Integer = -1
                                Dim pvName As String = ""
                                Call splitHryFullnameTo2(fullmilestoneName, milestoneName, breadcrumb, type, pvName)

                                If type = -1 Or
                                    (type = PTItemType.projekt And pvName = calcProjektKey(kvp.Value)) Or
                                    (type = PTItemType.vorlage) Then

                                    Dim milestoneIndices(,) As Integer = kvp.Value.hierarchy.getMilestoneIndices(milestoneName, breadcrumb)
                                    ' in milestoneIndices sind jetzt die Phasen- und Meilenstein Index der Phasen bzw Meilenstein Liste

                                    For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1
                                        cMilestone = .getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx))
                                        If Not IsNothing(cMilestone) Then
                                            Dim msColumn As Integer = getColumnOfDate(cMilestone.getDate)
                                            If msColumn > bis Or msColumn < von Then
                                            Else
                                                found = True
                                                Exit For
                                            End If
                                        End If
                                    Next

                                End If

                                If found Then
                                    Exit For
                                End If
                            Next
                        End If


                    End If


                End With

                If found Then
                    key = kvp.Value.tfZeile + kvp.Value.anzahlRasterElemente / 10000
                    tmpListe.Add(key, kvp.Value.name)
                End If

            Next

            withinTimeFrame = tmpListe

        End Get
    End Property


    ''' <summary>
    ''' gibt einen zweidimensionalen Array zurück, der die Namen der Projekte enthält, die eines der angegebenen Elemente im jeweiligen Zeitraum enthalten 
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="prcTyp"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProjectsWithElemNameInMonth(ByVal myCollection As Collection, ByVal prcTyp As String) As String(,)
        Get
            Dim zeitraum As Integer = showRangeRight - showRangeLeft
            Dim maxAnzahl As Integer = ShowProjekte.Count - 1
            Dim curMonat As Integer = 0
            Dim hproj As clsProjekt
            Dim cMilestone As clsMeilenstein

            Dim roleValues(zeitraum) As Double
            Dim costValues(zeitraum) As Double
            Dim ergebnisListe(zeitraum, maxAnzahl) As String
            Dim curElemIX(zeitraum) As Integer
            Dim abbrev As String = "-"



            Dim elemName As String = ""
            Dim breadCrumb As String = ""

            If showRangeRight = 0 Or showRangeLeft = 0 Then
                ' nichts tun 
            Else
                For cix As Integer = 1 To myCollection.Count
                    Dim pvName As String = ""
                    Dim type As Integer = -1
                    If prcTyp = DiagrammTypen(1) Then
                        Dim tmpTeamID As Integer = -1
                        elemName = RoleDefinitions.getRoleDefByIDKennung(CStr(myCollection.Item(cix)), tmpTeamID).name
                    Else
                        Call splitHryFullnameTo2(CStr(myCollection.Item(cix)), elemName, breadCrumb, type, pvName)
                    End If


                    If prcTyp = DiagrammTypen(0) Then
                        ' Phasen
                        Dim hphase As clsPhase
                        Dim prAnfang As Integer
                        Dim prEnde As Integer
                        Dim phAnfang As Integer
                        Dim phEnde As Integer
                        Dim ixZeitraum As Integer
                        Dim anzLoops As Integer
                        Dim Dauer As Integer


                        For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                            If type = -1 Or
                                (type = PTItemType.projekt And pvName = calcProjektKey(kvp.Value)) Or
                                (type = PTItemType.vorlage) Then

                                hproj = kvp.Value
                                Dauer = hproj.anzahlRasterElemente
                                Dim tempArray(Dauer - 1) As Double

                                Dim phaseIndices() As Integer = hproj.hierarchy.getPhaseIndices(elemName, breadCrumb)

                                For px As Integer = 0 To phaseIndices.Length - 1

                                    If phaseIndices(px) > 0 And phaseIndices(px) <= hproj.CountPhases Then
                                        hphase = hproj.getPhase(phaseIndices(px))
                                    Else
                                        hphase = Nothing
                                    End If


                                    If Not hphase Is Nothing Then

                                        abbrev = PhaseDefinitions.getAbbrev(hphase.name)

                                        With hproj
                                            prAnfang = .Start + .StartOffset
                                            prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                                        End With


                                        If istBereichInTimezone(prAnfang, prEnde) Then
                                            'projektstart = hproj.Start

                                            With hphase
                                                phAnfang = prAnfang + .relStart - 1
                                                phEnde = prAnfang + .relEnde - 1
                                            End With

                                            'Dim ixKorrektur As Integer = hphase.relStart - 1
                                            Dim ix As Integer
                                            Call awinIntersectZeitraum(phAnfang, phEnde, ixZeitraum, ix, anzLoops)

                                            If anzLoops > 0 Then
                                                ' dann ist die Phase enthalten 

                                                Try

                                                    tempArray = hproj.getPhasenBedarf(elemName)
                                                    ' ix bezeichnet aktuell den start mit Nullpunkt Phasen-Start, das muss jetzt korrigiert werden 
                                                    ix = ix + hphase.relStart - 1

                                                Catch ex As Exception

                                                End Try

                                                For al As Integer = 1 To anzLoops
                                                    If ixZeitraum + al - 1 > zeitraum Then
                                                        ' Fehlerprotokoll schreiben ...  
                                                    Else
                                                        ' wenn mehr als ein Element angezeigt werden soll, soll die Abkürzung dazugeschrieben werden 
                                                        If myCollection.Count > 1 Then
                                                            ' nach dem Doppelpunkt solte immer der Wert stehen, nicht der Bezeichner
                                                            'ergebnisListe(ixZeitraum + al - 1, curElemIX(ixZeitraum + al - 1)) = hproj.getShapeText & ":" & abbrev
                                                            ergebnisListe(ixZeitraum + al - 1, curElemIX(ixZeitraum + al - 1)) = hproj.getShapeText _
                                                                                                                                    & ":X"
                                                        Else
                                                            If awinSettings.phasesProzentual Then
                                                                ergebnisListe(ixZeitraum + al - 1, curElemIX(ixZeitraum + al - 1)) = hproj.getShapeText _
                                                                                                                                & ":" & tempArray(ix + al - 1).ToString("0%")
                                                            Else
                                                                ergebnisListe(ixZeitraum + al - 1, curElemIX(ixZeitraum + al - 1)) = hproj.getShapeText _
                                                                                                                                    & ":X"
                                                            End If

                                                        End If

                                                        curElemIX(ixZeitraum + al - 1) = curElemIX(ixZeitraum + al - 1) + 1
                                                    End If

                                                Next

                                            End If


                                        End If
                                    End If
                                Next

                            End If


                        Next kvp

                    ElseIf prcTyp = DiagrammTypen(1) Then
                        ' Rollen

                        Dim Dauer As Integer
                        For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                            hproj = kvp.Value

                            Dauer = hproj.anzahlRasterElemente
                            Dim tempArray(Dauer - 1) As Double
                            Dim prAnfang As Integer, prEnde As Integer
                            Dim ixZeitraum As Integer, ixArray As Integer, anzLoops As Integer

                            With hproj
                                prAnfang = .Start + .StartOffset
                                prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                            End With

                            anzLoops = 0
                            Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ixArray, anzLoops)

                            If anzLoops > 0 Then

                                Try

                                    tempArray = hproj.getRessourcenBedarf(elemName, inclSubRoles:=True)

                                Catch ex As Exception

                                End Try
                            End If

                            If tempArray.Sum > 0 Then

                                For al As Integer = 1 To anzLoops
                                    If ixZeitraum + al - 1 > zeitraum Then
                                        ' Fehlerprotokoll schreiben ...  
                                    ElseIf tempArray(ixArray + al - 1) > 0 Then
                                        ergebnisListe(ixZeitraum + al - 1, curElemIX(ixZeitraum + al - 1)) = hproj.getShapeText & ":" & CInt(tempArray(ixArray + al - 1)).ToString
                                        curElemIX(ixZeitraum + al - 1) = curElemIX(ixZeitraum + al - 1) + 1
                                    End If

                                Next

                            End If




                        Next

                    ElseIf prcTyp = DiagrammTypen(2) Then
                        ' Kostenarten

                        Dim Dauer As Integer
                        For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                            hproj = kvp.Value

                            Dauer = hproj.anzahlRasterElemente
                            Dim tempArray(Dauer - 1) As Double
                            Dim prAnfang As Integer, prEnde As Integer
                            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer

                            With hproj
                                prAnfang = .Start + .StartOffset
                                prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                            End With

                            anzLoops = 0
                            Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                            If anzLoops > 0 Then
                                Try

                                    tempArray = hproj.getKostenBedarf(elemName)

                                Catch ex As Exception

                                End Try
                            End If

                            If tempArray.Sum > 0 Then
                                ' andernfalls kein Kostenbedarf 

                                For al As Integer = 1 To anzLoops
                                    If ixZeitraum + al - 1 > zeitraum Then
                                        ' Fehlerprotokoll schreiben ...  
                                    ElseIf tempArray(ix + al - 1) > 0 Then
                                        ergebnisListe(ixZeitraum + al - 1, curElemIX(ixZeitraum + al - 1)) = hproj.getShapeText & ":" & CInt(tempArray(ix + al - 1)).ToString
                                        curElemIX(ixZeitraum + al - 1) = curElemIX(ixZeitraum + al - 1) + 1
                                    End If

                                Next

                            End If




                        Next


                    ElseIf prcTyp = DiagrammTypen(5) Then
                        ' Meilensteine 

                        For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                            If type = -1 Or
                                (type = PTItemType.projekt And pvName = kvp.Value.name) Or
                                (type = PTItemType.vorlage And pvName = kvp.Value.VorlagenName) Then

                                hproj = kvp.Value

                                ' neuer Code
                                Dim milestoneIndices(,) As Integer = hproj.hierarchy.getMilestoneIndices(elemName, breadCrumb)

                                For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                                    cMilestone = hproj.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx))

                                    If Not IsNothing(cMilestone) Then

                                        abbrev = MilestoneDefinitions.getAbbrev(cMilestone.name)

                                        Dim ix As Integer
                                        ' bestimme den monatsbezogenen Index im Array 
                                        ix = getColumnOfDate(cMilestone.getDate) - showRangeLeft

                                        If ix >= 0 And ix <= zeitraum Then

                                            If myCollection.Count > 1 Then
                                                'ergebnisListe(ix, curElemIX(ix)) = hproj.getShapeText & ":" & abbrev
                                                ergebnisListe(ix, curElemIX(ix)) = hproj.getShapeText & "-" & abbrev & ":X"
                                            Else
                                                ergebnisListe(ix, curElemIX(ix)) = hproj.getShapeText & ":X"
                                            End If

                                            curElemIX(ix) = curElemIX(ix) + 1

                                        End If


                                    End If

                                Next

                            End If


                        Next kvp



                    End If


                Next
            End If


            getProjectsWithElemNameInMonth = ergebnisListe

        End Get
    End Property


    Public ReadOnly Property getCountMilestoneCategoriesInMonth(ByVal categoryName As String) As Double()
        Get
            Dim milestoneValues() As Double
            Dim zeitraum As Integer
            Dim anzProjekte As Integer

            'Dim cphase As clsPhase
            'Dim cresult As clsMeilenstein
            Dim cMilestone As clsMeilenstein
            Dim hproj As clsProjekt
            Dim ix As Integer
            Dim idFarbe As Integer

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim milestoneValues(zeitraum)

            anzProjekte = _allProjects.Count

            ' Schleife über alle Projekte 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                hproj = kvp.Value
                ' hole die IDs aller Meilensteine, die die besagte Category haben 
                Dim IDCollection As Collection = kvp.Value.getMilestoneIDsWithCat(categoryName)

                For Each elemID As String In IDCollection

                    cMilestone = kvp.Value.getMilestoneByID(elemID)
                    If Not IsNothing(cMilestone) Then
                        ' bestimme den monatsbezogenen Index im Array 
                        ix = getColumnOfDate(cMilestone.getDate) - showRangeLeft

                        If ix >= 0 And ix <= zeitraum Then

                            If cMilestone.bewertungsCount > 0 Then
                                idFarbe = cMilestone.getBewertung(1).colorIndex
                            Else
                                idFarbe = 0
                            End If

                            milestoneValues(ix) = milestoneValues(ix) + 1

                        End If
                    End If
                Next

            Next kvp

            getCountMilestoneCategoriesInMonth = milestoneValues

        End Get
    End Property

    ''' <summary>
    ''' gibt den Zusatz Text zurück: 21 P = 6+4+8+3
    ''' wobei die erste Zahl die Gesamtsumme der Projekte in dieser Phase darstellt, 
    ''' die zweite Zahl die Zahl der nicht bewerteten Projekte, 
    ''' die dritte Zahl die grünen, dann gelben sowie roten PRojekte 
    ''' Die einzelnen Ziffern sollen auch entspreched eingefärbt werden 
    ''' der phaseName Finished darf nicht vorkommen 
    ''' </summary>
    ''' <param name="phaseName"></param>
    ''' <returns></returns>
    Public Function bestimmeAddOnTxtPfContainer(ByVal phaseName As String, ByVal timestamp As Date) As String()

        Dim tmpresult(4) As String
        ' hier muss ggf noch untersucht werden, ob der Phase-Name bereits den Breadcrum enzhält; 
        ' dann muss der hier noch bestimmt werden 
        ' aktuell wird davon ausgegangen, dass die Phasen-Namen eindeutig sind und es deshalb reicht, nur den Phasen NAmen anzugeben 

        Dim tmpSum As Integer = 0
        For i As Integer = 0 To 3
            Dim tmpAnz As Integer = Me.getCountProjectsInPhaseWithColor(phaseName, i, timestamp)
            tmpresult(i + 1) = tmpAnz.ToString
            tmpSum = tmpSum + tmpAnz
        Next

        tmpresult(0) = tmpSum.ToString
        bestimmeAddOnTxtPfContainer = tmpresult
    End Function


    ''' <summary>
    ''' gibt für die Showprojekte die Anzahl der Projekte zurück, die sich aktuell in der angegebenen Phase befinden 
    ''' phName enthält ggf. einen Teil des Breadcrumbs 
    ''' </summary>
    ''' <param name="phName"></param>
    ''' <param name="colorCode"></param>
    ''' <param name="timestamp"></param>
    ''' <returns></returns>
    Public ReadOnly Property getCountProjectsInPhaseWithColor(ByVal phName As String,
                                                              ByVal colorCode As Integer,
                                                              ByVal timestamp As Date) As Integer
        Get
            Dim tmpResult As Integer = 0
            ' Schleife über alle Projekte 
            Dim endeKennwort As String = "$finished"

            If phName = endeKennwort Then
                For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                    If kvp.Value.endeDate < timestamp Then
                        If kvp.Value.ampelStatus = colorCode Then
                            tmpResult = tmpResult + 1
                        End If
                    End If

                Next
            Else
                For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                    Dim delta As Integer
                    If kvp.Value.isInPhase(phName, timestamp, delta) Then

                        If kvp.Value.ampelStatus = colorCode Then
                            tmpResult = tmpResult + 1
                        End If

                    End If

                Next
            End If

            getCountProjectsInPhaseWithColor = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt die Infos der Projekte aus Showprojekte zurück, die aktuell in der angegebenen Phase sind
    ''' </summary>
    ''' <param name="phName"></param>
    ''' <param name="breadCrumb"></param>
    ''' <returns></returns>
    Public ReadOnly Property getInfosOfProjectsInPhase(ByVal phName As String,
                                                       ByVal breadCrumb As String) As Collection
        Get

            Dim tmpresult As New Collection
            getInfosOfProjectsInPhase = tmpresult

        End Get
    End Property

    ''' <summary>
    ''' gibt einen Array zurück, der angibt wie oft der übergebene Milestone im jeweiligen Monat vorkommt 
    ''' showrangeleft und showrangeright spannen den Betrachtungszeitraum auf
    ''' es wird ein Array der Dimension (3,zeitraum) zurückgegeben: 
    ''' 0: nicht bewertet, 1: grün, 2:gelb, 3: rot
    ''' </summary>
    ''' <param name="milestoneName">Name des Meilensteins</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCountMilestonesInMonth(ByVal milestoneName As String, ByVal breadcrumb As String, ByVal type As Integer, ByVal pvName As String) As Double(,)
        Get

            Dim milestoneValues(,) As Double
            Dim zeitraum As Integer
            Dim anzProjekte As Integer

            'Dim cphase As clsPhase
            'Dim cresult As clsMeilenstein
            Dim cMilestone As clsMeilenstein
            Dim hproj As clsProjekt
            Dim ix As Integer
            Dim idFarbe As Integer

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim milestoneValues(3, zeitraum)

            anzProjekte = _allProjects.Count

            ' Schleife über alle Projekte 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                hproj = kvp.Value

                If type = -1 Or
                    (type = PTItemType.vorlage) Or
                    (type = PTItemType.projekt And pvName = calcProjektKey(hproj)) Then
                    ' Aktion machen

                    ' neuer Code
                    Dim milestoneIndices(,) As Integer = hproj.hierarchy.getMilestoneIndices(milestoneName, breadcrumb)

                    For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                        cMilestone = hproj.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx))

                        If Not IsNothing(cMilestone) Then

                            ' bestimme den monatsbezogenen Index im Array 
                            ix = getColumnOfDate(cMilestone.getDate) - showRangeLeft

                            If ix >= 0 And ix <= zeitraum Then

                                If cMilestone.bewertungsCount > 0 Then
                                    idFarbe = cMilestone.getBewertung(1).colorIndex
                                Else
                                    idFarbe = 0
                                End If

                                milestoneValues(idFarbe, ix) = milestoneValues(idFarbe, ix) + 1

                            End If


                        End If

                    Next


                End If

            Next kvp


            getCountMilestonesInMonth = milestoneValues


        End Get
    End Property

    ''' <summary>
    ''' gibt einen Array zurück, der angibt, wie oft die angegebene Phase vorkommt
    ''' showrangeleft und showrangeright spannen den Betrachtungszeitraum auf 
    ''' 
    ''' </summary>
    ''' <param name="phaseName">Name der Phase</param>
    ''' <param name="type">wurde per Vorlage oder projekt eingeschränkt ?</param>
    ''' <param name="pvName" >wie heisst die Vorlage oder das Projekt </param>
    ''' <value></value>
    ''' <returns>gibt einen Array der Länge (showrangeright-showrangeleft+1) zurück </returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCountPhasesInMonth(phaseName As String, ByVal breadcrumb As String,
                                                   ByVal type As Integer, pvName As String) As Double()

        Get
            Dim phasevalues() As Double

            'Dim anzPhasen As Integer
            Dim zeitraum As Integer
            'Dim projektstart As Integer
            Dim anzProjekte As Integer
            'Dim found As Boolean
            Dim i As Integer ', pr As Integer, ph As Integer
            Dim hphase As clsPhase
            Dim hproj As clsProjekt
            'Dim lookforIndex As Boolean
            'Dim phasenStart As Integer, phasenEnde As Integer
            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer, phAnfang As Integer, phEnde As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            'lookforIndex = IsNumeric(phaseId)
            zeitraum = showRangeRight - showRangeLeft
            ReDim phasevalues(zeitraum)

            anzProjekte = _allProjects.Count

            ' anzPhasen = AllPhases.Count

            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                hproj = kvp.Value

                If type = -1 Or
                    (type = PTItemType.vorlage) Or
                    (type = PTItemType.projekt And pvName = calcProjektKey(hproj)) Then
                    ' Aktion machen

                    Dim phaseIndices() As Integer = hproj.hierarchy.getPhaseIndices(phaseName, breadcrumb)

                    For px As Integer = 0 To phaseIndices.Length - 1

                        If phaseIndices(px) > 0 And phaseIndices(px) <= hproj.CountPhases Then
                            hphase = hproj.getPhase(phaseIndices(px))
                        Else
                            hphase = Nothing
                        End If


                        If Not hphase Is Nothing Then

                            With hproj
                                prAnfang = .Start + .StartOffset
                                prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                            End With


                            If istBereichInTimezone(prAnfang, prEnde) Then
                                'projektstart = hproj.Start

                                With hphase
                                    phAnfang = prAnfang + .relStart - 1
                                    phEnde = prAnfang + .relEnde - 1
                                End With

                                Dim ixKorrektur As Integer = hphase.relStart - 1

                                Call awinIntersectZeitraum(phAnfang, phEnde, ixZeitraum, ix, anzLoops)

                                If anzLoops > 0 Then

                                    'ReDim tempArray(phEnde - phAnfang)
                                    tempArray = hproj.getPhasenBedarf(phaseName)

                                    For i = 0 To anzLoops - 1
                                        ' das awinintersect ermittelt die Werte für Projekt-Anfang, Projekt-Ende 
                                        ' in temparray stehen dagegen , deswegen muss um .relstart-1 erhöht werden 
                                        phasevalues(ixZeitraum + i) = phasevalues(ixZeitraum + i) + tempArray(ix + i + ixKorrektur)
                                    Next i

                                End If


                            End If
                        End If
                    Next


                End If


            Next kvp


            getCountPhasesInMonth = phasevalues

        End Get

    End Property

    ''' <summary>
    ''' gibt einen Array zurück, der angibt, wie oft die angegebene Phasen-Klasse vorkommt
    ''' showrangeleft und showrangeright spannen den Betrachtungszeitraum auf 
    ''' 
    ''' </summary>
    ''' <param name="categoryName">Name der Phase</param>
    ''' <value></value>
    ''' <returns>gibt einen Array der Länge (showrangeright-showrangeleft+1) zurück </returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCountPhaseCategoriesInMonth(ByVal categoryName As String) As Double()

        Get
            Dim phasevalues() As Double

            'Dim anzPhasen As Integer
            Dim zeitraum As Integer
            'Dim projektstart As Integer
            Dim anzProjekte As Integer
            'Dim found As Boolean
            Dim i As Integer ', pr As Integer, ph As Integer
            Dim hphase As clsPhase
            Dim hproj As clsProjekt
            'Dim lookforIndex As Boolean
            'Dim phasenStart As Integer, phasenEnde As Integer
            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer, phAnfang As Integer, phEnde As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            'lookforIndex = IsNumeric(phaseId)
            zeitraum = showRangeRight - showRangeLeft
            ReDim phasevalues(zeitraum)

            anzProjekte = _allProjects.Count

            ' anzPhasen = AllPhases.Count

            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                hproj = kvp.Value

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                End With

                Dim IDCollection As Collection = kvp.Value.getMilestoneIDsWithCat(categoryName)

                For Each elemID As String In IDCollection

                    hphase = kvp.Value.getPhaseByID(elemID)

                    If Not IsNothing(hphase) Then


                        If istBereichInTimezone(prAnfang, prEnde) Then
                            'projektstart = hproj.Start

                            With hphase
                                phAnfang = prAnfang + .relStart - 1
                                phEnde = prAnfang + .relEnde - 1
                            End With

                            Dim ixKorrektur As Integer = hphase.relStart - 1

                            Call awinIntersectZeitraum(phAnfang, phEnde, ixZeitraum, ix, anzLoops)

                            If anzLoops > 0 Then

                                'ReDim tempArray(phEnde - phAnfang)
                                tempArray = hproj.getPhasenBedarf(hphase.name)

                                For i = 0 To anzLoops - 1
                                    ' das awinintersect ermittelt die Werte für Projekt-Anfang, Projekt-Ende 
                                    ' in temparray stehen dagegen , deswegen muss um .relstart-1 erhöht werden 
                                    phasevalues(ixZeitraum + i) = phasevalues(ixZeitraum + i) + tempArray(ix + i + ixKorrektur)
                                Next i

                            End If


                        End If

                    End If

                Next elemID


            Next kvp


            getCountPhaseCategoriesInMonth = phasevalues

        End Get

    End Property
    ''' <summary>
    ''' returns true if in any month there is a overutilization of more than three time overloadCriterion
    ''' 
    ''' </summary>
    ''' <param name="roleIDs"></param>
    ''' <param name="skillIDs"></param>
    ''' <param name="overloadCriterion">returns true, if any month is overloaded more than overloadCriterion and onlyStrictly=false </param>
    ''' <param name="onlyStrictly">false: single months overloads should be taken into account even when overall timeframe is not overloaded at all </param>
    ''' <param name="totalOverloadCriterion">returns true if total sum of roles is larger than totalOverloadCriterion * kapa </param>
    ''' <returns></returns>
    Public Function overLoadFound(ByVal roleIDs As List(Of String),
                                  ByVal skillIDs As List(Of String),
                                  ByVal onlyStrictly As Boolean,
                                  ByVal overloadCriterion As Double,
                                  ByVal totalOverloadCriterion As Double) As Boolean

        Dim overloaded As Boolean = False
        Dim monthlyCriterion As Double = 3 * overloadCriterion
        Dim curIDs As List(Of String) = Nothing

        For i As Integer = 1 To 2

            If i = 1 Then
                curIDs = roleIDs
            Else
                curIDs = skillIDs
            End If

            If Not IsNothing(curIDs) Then

                For Each roleIDstr As String In curIDs

                    Dim roleValues As Double() = getRoleValuesInMonth(roleIDstr, considerAllSubRoles:=True)
                    Dim myCollection As New Collection From {
                        roleIDstr
                     }
                    Dim kapaValues As Double() = getRoleKapasInMonth(myCollection)

                    If Not onlyStrictly Then
                        For ix As Integer = 0 To roleValues.Length - 1
                            If roleValues(ix) >= overloadCriterion * kapaValues(ix) Then
                                overloaded = True
                                Exit For
                            End If
                        Next
                    End If

                    If Not overloaded And (roleValues.Sum >= totalOverloadCriterion * kapaValues.Sum) Then
                        overloaded = True
                    End If

                    If overloaded Then
                        Exit For
                    End If

                Next

            End If

            If overloaded Then
                Exit For
            End If

        Next

        overLoadFound = overloaded
    End Function

    ''' <summary>
    ''' bestimmt für den betrachteten Zeitraum für die angegebene Rolle die benötigte Summe pro Monat; roleid wird als String oder Key(Integer) übergeben
    ''' </summary>
    ''' <param name="roleIDStr"></param>
    ''' <value>String für Rollenbezeichner oder Integer für den Key der Rolle</value>
    ''' <returns>Array, der die Werte der gefragten Rolle pro Monat des betrachteten Zeitraums enthält</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleValuesInMonth(ByVal roleIDStr As String,
                                                  Optional ByVal considerAllSubRoles As Boolean = False,
                                                  Optional ByVal type As PTcbr = PTcbr.all,
                                                  Optional ByVal considerAllNeedsOfRolesHavingTheseSkills As Boolean = False,
                                                  Optional ByVal excludedNames As Collection = Nothing) As Double()
        Get
            Dim roleValues() As Double
            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim anzProjekte As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt

            Dim tempArray() As Double
            Dim testArray() As Double
            Dim prAnfang As Integer, prEnde As Integer

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum
            Dim teamID As Integer
            Dim roleID As Integer = RoleDefinitions.parseRoleNameID(roleIDStr, teamID)
            Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(roleID, teamID)

            Dim currentRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(roleID)


            zeitraum = showRangeRight - showRangeLeft
            ReDim roleValues(zeitraum)


            anzProjekte = _allProjects.Count

            ' anzPhasen = AllPhases.Count

            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                hproj = kvp.Value

                Dauer = hproj.anzahlRasterElemente

                ReDim tempArray(Dauer - 1)
                ReDim testArray(Dauer - 1)

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                End With

                anzLoops = 0
                Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                If anzLoops > 0 Then

                    Try
                        tempArray = hproj.getRessourcenBedarf(roleNameID,
                                                              inclSubRoles:=considerAllSubRoles,
                                                              considerAllOtherNeedsOfRolesHavingTheseSkills:=considerAllNeedsOfRolesHavingTheseSkills)

                        For i = 0 To anzLoops - 1
                            roleValues(ixZeitraum + i) = roleValues(ixZeitraum + i) + tempArray(ix + i)
                        Next i

                    Catch ex As Exception

                    End Try

                End If

            Next kvp


            getRoleValuesInMonth = roleValues

        End Get


    End Property

    ''' <summary>
    ''' gibt für den aktuellen Zeitraum und die übergebene Collection mit Phasen-Namen die Schwellwerte an  
    ''' es muss berücksichtigt werden, dass in der myCollection jetzt ggf noch die Kennung steht, welche Vorlage bzw. welches PRojekt denn gemeint ist ...
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseSchwellWerteInMonth(myCollection As Collection) As Double()
        Get

            Dim schwellWerte() As Double

            Dim hkapa As Double
            Dim rname As String = ""
            Dim zeitraum As Integer
            Dim r As Integer, m As Integer
            Dim breadcrumb As String = ""
            Dim ok As Boolean = True


            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum
            zeitraum = showRangeRight - showRangeLeft
            ReDim schwellWerte(zeitraum)

            For r = 1 To myCollection.Count

                Dim pvName As String = ""
                Dim type As Integer = -1
                Call splitHryFullnameTo2(CStr(myCollection.Item(r)), rname, breadcrumb, type, pvName)

                If PhaseDefinitions.Contains(rname) And breadcrumb = "" And ok Then
                    hkapa = PhaseDefinitions.getPhaseDef(rname).schwellWert
                Else
                    hkapa = 0
                    ok = False
                End If

                ' nur wenn es sich um die uneingeschränkte Auswahl des Namens handelt bzw. wenn jedes Element aus der Liste einen Schwellwert hat ;
                ' soll der Schwellwert angezeigt werden 
                If ok Then
                    For m = 0 To zeitraum
                        ' Änderung 31.5 Holen der Schwellwerte einer Phase 
                        schwellWerte(m) = schwellWerte(m) + hkapa
                    Next m
                End If



            Next r

            getPhaseSchwellWerteInMonth = schwellWerte

        End Get
    End Property


    ''' <summary>
    ''' gibt die Meilenstein Kapa Werte zurück 
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneSchwellWerteInMonth(myCollection As Collection) As Double()
        Get

            Dim schwellWerte() As Double

            Dim hkapa As Double
            Dim msName As String = ""
            Dim zeitraum As Integer
            Dim r As Integer, m As Integer
            Dim breadcrumb As String = ""
            Dim ok As Boolean = True


            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum
            zeitraum = showRangeRight - showRangeLeft
            ReDim schwellWerte(zeitraum)

            For r = 1 To myCollection.Count

                'msName = CStr(myCollection.Item(r))
                Dim type As Integer = -1
                Dim pvName As String = ""
                Call splitHryFullnameTo2(CStr(myCollection.Item(r)), msName, breadcrumb, type, pvName)
                ' nur wenn es sich um die uneingeschränkte Auswahl des Namens handelt bzw. wenn jedes Element aus der Liste einen Schwellwert hat ;
                ' soll der Schwellwert angezeigt werden 
                If MilestoneDefinitions.Contains(msName) And breadcrumb = "" And ok Then
                    hkapa = MilestoneDefinitions.getMilestoneDef(msName).schwellWert
                Else
                    hkapa = 0
                    ok = False
                End If


                If ok Then
                    For m = 0 To zeitraum
                        ' Änderung 31.5 Holen der Schwellwerte einer Phase 
                        schwellWerte(m) = schwellWerte(m) + hkapa
                    Next m
                End If



            Next r

            getMilestoneSchwellWerteInMonth = schwellWerte

        End Get
    End Property

    ''' <summary>
    ''' gibt für die in myCollection übergebenen Rollen die Kapazitäten zurück 
    ''' wenn includingExterns = true, dann inkl der bereits beauftragten externen Ressourcen
    ''' die Aufschlüsselung ist den Ressource Rollen Dateien zu finden 
    ''' </summary>
    ''' <param name="myCollection">enthält die Liste der zu betrachtenden Rollen-Namen</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleKapasInMonth(ByVal myCollection As Collection,
                                                 Optional ByVal onlyIntern As Boolean = False) As Double()

        ' tk 12.10.20 es werden jetzt bei Skills / teams nicht mehr die anteilige Kapazität berücksichtigt, sondern immer die volle
        '  
        Get
            Dim kapaValues() As Double
            Dim tmpValues() As Double



            Dim zeitraum As Integer
            Dim m As Integer


            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            ' hier muss überprüft werden, welche Rollen denn Sammelrollen sind und deswegen ersetzt werden müssen durch ihre
            ' subroles ... 

            Dim realCollection As New SortedList(Of Integer, Double)

            For ix As Integer = 1 To myCollection.Count

                Dim teamID As Integer = -1
                Dim curRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(CStr(myCollection.Item(ix)), teamID)

                If Not IsNothing(curRole) Then

                    Dim roleName As String = curRole.name

                    If teamID > 0 Then
                        Dim subRoleList As List(Of Integer) = RoleDefinitions.getCommonChildsOfParents(curRole.UID, teamID)

                        For Each tmpID As Integer In subRoleList
                            Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(tmpID)
                            If Not IsNothing(tmpRole) Then
                                If Not tmpRole.isCombinedRole Then
                                    If Not realCollection.ContainsKey(tmpRole.UID) Then                                        '
                                        realCollection.Add(tmpRole.UID, 1.0)
                                    End If
                                End If
                            End If
                        Next

                    Else
                        If curRole.isCombinedRole Then
                            ' es handelt sich um eine Sammelrolle
                            ' Kapas sind nur in den realRoles , also den nicht Sammelrollen vorhanden ...
                            Dim subRoleListe As SortedList(Of Integer, Double) = RoleDefinitions.getSubRoleIDsOf(roleName:=roleName,
                                                                                            type:=PTcbr.realRoles,
                                                                                            excludedNames:=myCollection)

                            If subRoleListe.Count = 0 Then

                                If Not realCollection.ContainsKey(curRole.UID) Then
                                    ' es gibt keine Kinder 
                                    realCollection.Add(curRole.UID, 1.0)
                                End If

                            Else
                                ' jetzt müssen alle Elemente von tmpCollection aufgenommen werden, sofern sie nicht schon eh aufgenommen sind 

                                For Each srKvP As KeyValuePair(Of Integer, Double) In subRoleListe

                                    If Not realCollection.ContainsKey(srKvP.Key) Then
                                        realCollection.Add(srKvP.Key, 1.0)
                                    End If

                                Next


                            End If

                        Else

                            If Not realCollection.ContainsKey(curRole.UID) Then

                                realCollection.Add(curRole.UID, 1.0)

                            End If

                        End If

                    End If

                End If

            Next


            ' RealCollection enthält jetzt all die gesuchten Sub-Roles und ggf separat angegebenen Rollen inkl der Prozentsätze, wie die Kapa zu berechnen ist
            ' ist dann relevant, wenn es sich um eine Gruppe handelt, wo eine bestimmte Basis Rolle nur zu x% KApa beiträgt 

            zeitraum = showRangeRight - showRangeLeft
            ReDim kapaValues(zeitraum)
            ReDim tmpValues(zeitraum)

            For Each kvp As KeyValuePair(Of Integer, Double) In realCollection

                Dim curRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(kvp.Key)
                If Not IsNothing(curRole) Then

                    ' wenn onlyIntern gesucht wird, dann werden nur die Rollen betrachtet, die interne sind 
                    If Not onlyIntern Or Not curRole.isExternRole Then
                        ' in kvp.value steht jetzt der Prozentsatz, mit dem die Kapa der Rolle berücksichtig werden soll 
                        For i = showRangeLeft To showRangeRight

                            'tmpValues(i - showRangeLeft) = kvp.Value * curRole.kapazitaet(i)
                            tmpValues(i - showRangeLeft) = curRole.kapazitaet(i)

                            If tmpValues(i - showRangeLeft) < 0 Then
                                tmpValues(i - showRangeLeft) = 0
                            End If
                        Next


                        For m = 0 To zeitraum
                            ' Änderung 27.7 Holen der Kapa Werte , jetzt aufgeschlüsselt nach 
                            'kapaValues(m) = kapaValues(m) + hkapa
                            kapaValues(m) = kapaValues(m) + tmpValues(m)
                        Next m
                    End If

                End If

            Next

            getRoleKapasInMonth = kapaValues
        End Get

    End Property

    ''' <summary>
    ''' gibt zurück, wieviele rote, grüne, gelbe und graue Bewertungen im betrachteten Zeitraum vorhanden sind 
    ''' future gibt an, was betrachtet werden soll
    ''' -1: nur heute und Vergangenheit 
    ''' 0: Vergangenheit und Zukunft 
    ''' +1: Zukunft 
    ''' </summary>
    ''' <param name="colorIndex"></param>
    ''' <param name="future"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getColorsInMonth(ByVal colorIndex As Integer, ByVal future As Integer) As Integer()
        Get
            Dim colorsInMonth() As Integer

            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt

            Dim tempArray() As Integer
            Dim prAnfang As Integer, prEnde As Integer
            Dim heuteColumn As Integer = getColumnOfDate(Date.Now)
            Dim vglWert As Integer = heuteColumn - showRangeLeft



            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim colorsInMonth(zeitraum)

            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                hproj = kvp.Value

                Dauer = hproj.anzahlRasterElemente

                ReDim tempArray(Dauer - 1)

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                End With

                anzLoops = 0
                Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                If anzLoops > 0 Then

                    tempArray = hproj.getNrColorIndexes(colorIndex)

                    For i = 0 To anzLoops - 1
                        colorsInMonth(ixZeitraum + i) = colorsInMonth(ixZeitraum + i) + tempArray(ix + i)
                    Next i


                End If
                'hproj = Nothing
            Next kvp

            If future = 1 Then

                ' die Werte, die für die Vergangenheit stehen, werden gelöscht 
                For i = 0 To vglWert
                    colorsInMonth(i) = 0
                Next

            ElseIf future = -1 Then

                ' die Werte, die für die Zukunft stehen werden gelöscht 
                If vglWert >= -1 Then
                    For i = vglWert + 1 To zeitraum
                        colorsInMonth(i) = 0
                    Next
                End If

            End If


            getColorsInMonth = colorsInMonth



        End Get
    End Property

    ''' <summary>
    ''' gibt die Summe der Einzelbudgets der einzelnen angezeigten (Summary-)Projekte aus 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getBudgetOfShownProjects() As Double
        Get
            Dim tmpResult As Double = 0.0

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                tmpResult = tmpResult + kvp.Value.budgetWerte.Sum
            Next

            getBudgetOfShownProjects = tmpResult
        End Get
    End Property



    ''' <summary>
    ''' gibt über alle betrachteten Projekte die anteiligen Budget Werte zurück 
    ''' das Budget wird jetzt nicht mehr über die budgetvalues berechnet, sondern über costvalue * marge 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Mit diesem neuen Ansatz wird sichergestellt, dass nur soviel vom Gesamtbudget aufgebraucht wird wie tatsächlich aufgrund der 
    ''' angefallenen Kosten in dem Monat auch benötigt wird </remarks>
    Public ReadOnly Property getBudgetValuesInMonth() As Double()
        Get

            Dim projektBudget As Double
            Dim budgetValues() As Double

            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt

            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer


            Dim avgBudget As Double


            zeitraum = showRangeRight - showRangeLeft
            ReDim budgetValues(zeitraum)


            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                hproj = kvp.Value

                Dauer = hproj.anzahlRasterElemente
                projektBudget = hproj.Erloes
                avgBudget = projektBudget / hproj.anzahlRasterElemente


                'ReDim tempArray(Dauer - 1)
                tempArray = kvp.Value.budgetWerte

                If IsNothing(tempArray) Then
                    ReDim tempArray(Dauer - 1)
                    For i = 0 To Dauer - 1
                        tempArray(i) = avgBudget
                    Next
                Else
                    If tempArray.Sum = 0 Then
                        ReDim tempArray(Dauer - 1)
                        For i = 0 To Dauer - 1
                            tempArray(i) = avgBudget
                        Next
                    End If
                End If


                With hproj

                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset

                End With

                anzLoops = 0
                Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                If anzLoops > 0 Then


                    For i = 0 To anzLoops - 1
                        budgetValues(ixZeitraum + i) = budgetValues(ixZeitraum + i) + tempArray(ix + i)
                    Next i


                End If

            Next kvp

            getBudgetValuesInMonth = budgetValues

        End Get
    End Property

    ''' <summary>
    ''' aletr aNsatz , der noch auf die budgetWerte abhob ... 
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBudgetValuesInMonth_deprecated() As Double()
        Get

            Dim projektBudget As Double
            Dim budgetValues() As Double
            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt

            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer


            Dim avgBudget As Double


            zeitraum = showRangeRight - showRangeLeft
            ReDim budgetValues(zeitraum)


            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                hproj = kvp.Value

                Dauer = hproj.anzahlRasterElemente
                projektBudget = hproj.Erloes
                avgBudget = projektBudget / hproj.anzahlRasterElemente


                'ReDim tempArray(Dauer - 1)
                tempArray = kvp.Value.budgetWerte

                If IsNothing(tempArray) Then
                    ReDim tempArray(Dauer - 1)
                    For i = 0 To Dauer - 1
                        tempArray(i) = avgBudget
                    Next
                Else
                    If tempArray.Sum = 0 Then
                        ReDim tempArray(Dauer - 1)
                        For i = 0 To Dauer - 1
                            tempArray(i) = avgBudget
                        Next
                    End If
                End If


                With hproj

                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset

                End With

                anzLoops = 0
                Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                If anzLoops > 0 Then


                    For i = 0 To anzLoops - 1
                        budgetValues(ixZeitraum + i) = budgetValues(ixZeitraum + i) + tempArray(ix + i)
                    Next i


                End If

            Next kvp

            getBudgetValuesInMonth_deprecated = budgetValues

        End Get
    End Property

    Public ReadOnly Property getInvoices() As Double()
        Get
            Dim invoiceValues() As Double = Nothing
            Dim tempArray() As Double
            Dim zeitraum As Integer
            Dim prAnfang As Integer, prEnde As Integer

            Dim anzLoops As Integer = 0
            Dim ixZeitraum As Integer, ix As Integer


            Dim hproj As clsProjekt = Nothing
            zeitraum = showRangeRight - showRangeLeft
            ReDim invoiceValues(zeitraum)

            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                hproj = kvp.Value

                prAnfang = hproj.Start + hproj.StartOffset

                tempArray = hproj.getInvoicesPenalties

                If tempArray.Sum > 0 Then
                    prEnde = prAnfang + tempArray.Length - 1

                    Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                    If anzLoops > 0 Then

                        For i = 0 To anzLoops - 1
                            invoiceValues(ixZeitraum + i) = invoiceValues(ixZeitraum + i) + tempArray(ix + i)
                        Next i

                    End If
                End If


            Next

            getInvoices = invoiceValues

        End Get
    End Property

    ''' <summary>
    ''' calculates cashFlow values in months starting in column von and ending in column bis
    ''' if von > bis then change values accordingly 
    ''' if values are negativ: consider showRangeLeft and showrangeRight 
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <returns></returns>
    Public ReadOnly Property getCashFlow(Optional ByVal von As Integer = -1, Optional ByVal bis As Integer = -1) As Double()
        Get
            ' check validity

            If ((von < 0) Or (bis < 0)) Then
                If ((showRangeRight > 0) And (showRangeRight - showRangeLeft > 0)) Then
                    von = showRangeLeft
                    bis = showRangeRight
                Else
                    ' define next month a timeframe of 6 months 
                    von = getColumnOfDate(Date.Now) + 1
                    bis = von + 5
                End If
            Else
                ' take values of von and bis
            End If

            If bis < von Then
                Dim sav As Integer = von
                von = bis
                bis = sav
            End If

            Dim saveShowrangeLeft As Integer = showRangeLeft
            Dim saveShowrangeRight As Integer = showRangeRight

            ' now consider von and bis 
            showRangeLeft = von
            showRangeRight = bis

            Dim zeitraum As Integer = showRangeRight - showRangeLeft
            Dim kugCome As Double()
            Dim kugGo As Double()

            Dim shortTermQuota As Double = 0.67

            ReDim kugCome(zeitraum)
            ReDim kugGo(zeitraum)

            Dim cashFlowValues As Double() = Nothing
            ReDim cashFlowValues(zeitraum)
            'Dim cashFlowValues1 As Double() = Nothing
            'ReDim cashFlowValues1(zeitraum)

            Try
                Dim invoices() As Double = ShowProjekte.getInvoices

                ' den Vormonat mit betrachten 


                Dim orgaFullCost As Double() = RoleDefinitions.getFullCost(showRangeLeft, showRangeRight)
                Dim externCost As Double() = getCostGpValuesInMonth(PTrt.extern)
                Dim internCost As Double() = getCostGpValuesInMonth(PTrt.intern)
                Dim otherCost As Double() = getTotalCostValuesInMonth(False)

                For i As Integer = 0 To zeitraum
                    If internCost(i) > orgaFullCost(i) Then
                        ' das kann ja dann gar nicht geleistet werden - 
                        ' es muss von Externen gemacht werden 
                        Dim diff As Double = internCost(i) - orgaFullCost(i)
                        internCost(i) = orgaFullCost(i)
                        externCost(i) = externCost(i) + diff
                    End If
                Next

                If awinSettings.kurzarbeitActivated Then

                    If showRangeLeft > 1 Then
                        showRangeLeft = showRangeLeft - 1
                    End If
                    'Dim notUtilizedCapacity As Double() = ShowProjekte.getCostoValuesInMonth(provideKUGData:=True)
                    Dim notUtilizedCapacity As Double() = ShowProjekte.getNotUtilizedCapaValuesInMonth()
                    showRangeLeft = saveShowrangeLeft

                    ' jetzt muss die nicht ausgelastete Zeit abgezogen werden 
                    For i As Integer = 0 To zeitraum
                        kugCome(i) = notUtilizedCapacity(i) * shortTermQuota
                        kugGo(i) = notUtilizedCapacity(i + 1) * shortTermQuota
                    Next

                    For i As Integer = 0 To zeitraum
                        ' notUtilizedCapacity(i+1) adressiert den gleichen Monat wie invoices(i)
                        cashFlowValues(i) = invoices(i) + kugCome(i) - (kugGo(i) + externCost(i) + otherCost(i) + orgaFullCost(i) - notUtilizedCapacity(i + 1))
                        'cashFlowValues1(i) = invoices(i) + kugCome(i) - (externCost(i) + otherCost(i) + orgaFullCost(i) - notUtilizedCapacity(i + 1) * (1 - shortTermQuota))
                    Next
                    'If arraysAreDifferent(cashFlowValues, cashFlowValues1) Then
                    '    Call MsgBox("Stop , unterschiedliche Werte")
                    'End If

                Else
                    For i As Integer = 0 To zeitraum
                        cashFlowValues(i) = invoices(i) - (externCost(i) + otherCost(i) + orgaFullCost(i))
                    Next
                End If

            Catch ex As Exception
                Call MsgBox("Fehler")
            End Try



            getCashFlow = cashFlowValues
        End Get
    End Property

    ''' <summary>
    ''' gibt über alle betrachteten Projekte die Earned Values zurück; 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getEarnedValuesInMonth() As Double()

        Get
            Dim earnedValues() As Double
            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt
            'Dim lookforIndex As Boolean
            'Dim isPersCost As Boolean
            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer
            'Dim persCost As Boolean
            'Dim SRweight As Double ' nimmt die Gewichtung auf:= strategic Fit / Risiko
            Dim projektMarge As Double

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim earnedValues(zeitraum)


            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                hproj = kvp.Value

                Dauer = hproj.anzahlRasterElemente

                ReDim tempArray(Dauer - 1)

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                    projektMarge = .ProjectMarge

                End With

                anzLoops = 0
                Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                If anzLoops > 0 Then

                    tempArray = hproj.getGesamtKostenBedarf

                    For i = 0 To anzLoops - 1
                        earnedValues(ixZeitraum + i) = earnedValues(ixZeitraum + i) + tempArray(ix + i) * projektMarge
                    Next i


                End If
                'hproj = Nothing
            Next kvp

            getEarnedValuesInMonth = earnedValues

        End Get

    End Property
    '

    ''' <summary>
    ''' gibt für den betrachteten Zeitraum den Wert pro Monat an, um den der Earned Value 
    ''' aufgrund der Risiko Betrachtung und strategischen Einordnung rediziert werden sollte 
    ''' errechnet sich aus : strategicFit * WeightStrategicFit / risk * earned Value
    ''' der Wert für  strategicFit * WeightStrategicFit / risk kann dabei niemals größer als 1 werden 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getWeightedRiskValuesInMonth() As Double()

        Get
            Dim riskValues() As Double
            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt
            'Dim lookforIndex As Boolean
            'Dim isPersCost As Boolean
            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer
            'Dim persCost As Boolean
            'Dim SRweight As Double ' nimmt die Gewichtung auf:= strategic Fit / Risiko
            Dim riskweightedMarge As Double
            Dim heuteColumn As Integer

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim riskValues(zeitraum)

            heuteColumn = getColumnOfDate(Date.Today)

            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                hproj = kvp.Value

                Dauer = hproj.anzahlRasterElemente

                ReDim tempArray(Dauer - 1)

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset

                End With

                Dim heuteIndex As Integer = heuteColumn - prAnfang

                anzLoops = 0
                Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                If anzLoops > 0 Then

                    With hproj
                        tempArray = .getGesamtKostenBedarf
                        riskweightedMarge = .risikoKostenfaktor
                        If riskweightedMarge < 0 Then
                            riskweightedMarge = 0
                        End If

                    End With

                    If heuteColumn > showRangeRight Then
                        ' nichts mehr tun, es existieren keine Risiken mehr 
                    Else
                        ' die 
                        For i = 0 To anzLoops - 1
                            riskValues(ixZeitraum + i) = riskValues(ixZeitraum + i) + tempArray(ix + i) * riskweightedMarge
                        Next i
                    End If



                End If
                'hproj = Nothing
            Next kvp

            getWeightedRiskValuesInMonth = riskValues

        End Get

    End Property

    ''' <summary>
    ''' gibt die Gesamtkosten über alle Projekte im betrachteten Zeitraum zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTotalCostValuesInMonth(Optional ByVal includingPersonalCosts As Boolean = True) As Double()
        Get
            Dim costValues() As Double
            Dim zeitraum As Integer
            Dim tempArray() As Double

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)

            Dim anzCosts As Integer = CostDefinitions.Count
            ' die Persoanlkosten sind immer der letzte Eintrag in der Liste der Kostenarten ... 
            If Not includingPersonalCosts Then
                anzCosts = anzCosts - 1
            End If

            For k As Integer = 1 To anzCosts
                tempArray = Me.getCostValuesInMonth(k)
                For l As Integer = 0 To tempArray.Length - 1
                    costValues(l) = costValues(l) + tempArray(l)
                Next
            Next


            ' Änderung tk 19.5.16 
            ' alt und falsch: weil die Überstundenkosten nicht berücksichtigt werden ... 
            ''For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects
            ''    hproj = kvp.Value

            ''    Dauer = hproj.anzahlRasterElemente

            ''    ReDim tempArray(Dauer - 1)

            ''    With hproj
            ''        prAnfang = .Start + .StartOffset
            ''        prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
            ''    End With

            ''    anzLoops = 0
            ''    Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

            ''    If anzLoops > 0 Then

            ''        tempArray = hproj.getGesamtKostenBedarf

            ''        For i = 0 To anzLoops - 1
            ''            costValues(ixZeitraum + i) = costValues(ixZeitraum + i) + tempArray(ix + i)
            ''        Next i


            ''    End If
            ''    'hproj = Nothing
            ''Next kvp

            getTotalCostValuesInMonth = costValues

        End Get
    End Property

    ''' <summary>
    ''' gibt die Sonstigen Kosten zurück, also alle Kosten minus die Personalkosten
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getOtherCostValuesInMonth() As Double()
        Get
            Dim costValues() As Double
            Dim zeitraum As Integer
            Dim tempArray() As Double

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)

            Dim anzCosts As Integer = CostDefinitions.Count
            '
            ' die Personalkosten sind immer die letzte Kostenart ...
            For k As Integer = 1 To anzCosts - 1
                tempArray = Me.getCostValuesInMonth(k)
                For l As Integer = 0 To tempArray.Length - 1
                    costValues(l) = costValues(l) + tempArray(l)
                Next
            Next

            getOtherCostValuesInMonth = costValues

        End Get
    End Property

    '
    '
    '
    ''' <summary>
    ''' gibt die Gesamtkosten , Personalkosten und alle sonstigen Kosten im betrachteten Zeitraum zurück 
    ''' bei den Personalkosten sind die Überstundensätze bzw. externen Tagessätze im Normalfall nicht berücksichtigt  
    ''' </summary>
    ''' <param name="CostID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostValuesInMonth(CostID As Object) As Double()

        Get
            Dim costValues() As Double
            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt
            Dim lookforIndex As Boolean
            Dim isPersCost As Boolean
            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            lookforIndex = IsNumeric(CostID)

            If lookforIndex Then
                If CostID = CostDefinitions.Count Then
                    isPersCost = True
                End If
            Else
                If CostID = "Personalkosten" Then
                    isPersCost = True
                End If
            End If

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)


            If isPersCost Then
                costValues = Me.getCostGpValuesInMonth
            Else

                For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                    hproj = kvp.Value

                    Dauer = hproj.anzahlRasterElemente

                    ReDim tempArray(Dauer - 1)

                    With hproj
                        prAnfang = .Start + .StartOffset
                        prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                    End With

                    anzLoops = 0
                    Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                    If anzLoops > 0 Then

                        tempArray = hproj.getKostenBedarf(CostID)

                        For i = 0 To anzLoops - 1
                            costValues(ixZeitraum + i) = costValues(ixZeitraum + i) + tempArray(ix + i)
                        Next i


                    End If
                    'hproj = Nothing
                Next kvp

            End If




            getCostValuesInMonth = costValues

        End Get

    End Property

    ''' <summary>
    ''' gibt die Gesamtkosten , Personalkosten und alle sonstigen Kosten im betrachteten Zeitraum zurück 
    ''' bei den Personalkosten sind die Überstundensätze bzw. externen Tagessätze im Normalfall nicht berücksichtigt  
    ''' </summary>
    ''' <param name="CostID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostValuesInMonthNew(CostID As Object) As Double()

        Get
            Dim costValues() As Double
            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt
            Dim lookforIndex As Boolean
            Dim isPersCost As Boolean
            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            lookforIndex = IsNumeric(CostID)

            If lookforIndex Then
                If CostID = CostDefinitions.Count Then
                    isPersCost = True
                End If
            Else
                If CostID = "Personalkosten" Then
                    isPersCost = True
                End If
            End If

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)


            If isPersCost Then
                costValues = Me.getCostGpValuesInMonth
            Else

                For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                    hproj = kvp.Value

                    Dauer = hproj.anzahlRasterElemente

                    ReDim tempArray(Dauer - 1)

                    With hproj
                        prAnfang = .Start + .StartOffset
                        prEnde = .Start + .anzahlRasterElemente - 1 + .StartOffset
                    End With

                    anzLoops = 0
                    Call awinIntersectZeitraum(prAnfang, prEnde, ixZeitraum, ix, anzLoops)

                    If anzLoops > 0 Then

                        tempArray = hproj.getKostenBedarf(CostID)

                        For i = 0 To anzLoops - 1
                            costValues(ixZeitraum + i) = costValues(ixZeitraum + i) + tempArray(ix + i)
                        Next i


                    End If
                    'hproj = Nothing
                Next kvp

            End If




            getCostValuesInMonthNew = costValues

        End Get

    End Property

    ''' <summary>
    ''' gibt je nach Typ die Auslastungs-Werte im Zeitraum left, right zurück
    ''' </summary>
    ''' <param name="typus">0: Auslastung, 1: Überauslastung 2: Unterauslastung</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAuslastungsValues(typus As Integer) As Double()

        Get
            Dim roleValues() As Double
            Dim kapaValues() As Double
            Dim tmpValues() As Double
            Dim roleID As String
            Dim myCollection As New Collection
            Dim i As Integer, ix As Integer
            Dim zeitraum As Integer


            zeitraum = showRangeRight - showRangeLeft
            ReDim roleValues(zeitraum)
            ReDim kapaValues(zeitraum)
            ReDim tmpValues(zeitraum)


            ' hier wird die todo Liste bestimmt, die enthält nur Sammel-Rollen und ggf Rollen, die keiner der Sammelrollen angehören .. 
            Dim uniqueList As Collection = RoleDefinitions.getUniqueRoleList


            For i = 1 To uniqueList.Count

                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(CStr(uniqueList.Item(i)))

                Dim istSammelRolle As Boolean = tmpRole.isCombinedRole
                roleID = tmpRole.UID.ToString

                If istSammelRolle Then
                    ' nur Platzhalter Rollenbedarfe berücksichtigen 
                    roleValues = Me.getRoleValuesInMonth(roleIDStr:=roleID,
                                                         considerAllSubRoles:=True,
                                                         type:=PTcbr.all,
                                                         excludedNames:=Nothing)
                Else
                    roleValues = Me.getRoleValuesInMonth(roleID)

                End If

                myCollection.Add(roleID, roleID)
                kapaValues = Me.getRoleKapasInMonth(myCollection)
                myCollection.Clear()

                Select Case typus

                    Case 0
                        ' Auslastung

                        For ix = 0 To zeitraum

                            If roleValues(ix) > kapaValues(ix) Then
                                ' es werden die maximale Kapa dieser Rolle berücksichtigt 
                                tmpValues(ix) = tmpValues(ix) + kapaValues(ix)
                            Else
                                ' die internen Ressourcen reichen aus 
                                tmpValues(ix) = tmpValues(ix) + roleValues(ix)
                            End If

                        Next ix

                    Case 1
                        ' Überauslastung

                        For ix = 0 To zeitraum

                            If roleValues(ix) > kapaValues(ix) Then
                                ' es gibt Überauslastung  
                                tmpValues(ix) = tmpValues(ix) + roleValues(ix) - kapaValues(ix)
                            Else
                                ' es gibt keine Überauslastung 
                            End If

                        Next ix

                    Case 2
                        ' Unterauslastung
                        For ix = 0 To zeitraum

                            If roleValues(ix) < kapaValues(ix) Then
                                ' es gibt Unterauslastung  
                                tmpValues(ix) = tmpValues(ix) + kapaValues(ix) - roleValues(ix)
                            Else
                                ' es gibt keine Unterauslastung bzw.
                            End If

                        Next ix

                End Select

            Next i


            getAuslastungsValues = tmpValues


        End Get

    End Property

    ''' <summary>
    ''' gibt je nach Typ die Auslastungs-Werte für roleName im Zeitraum left, right zurück
    ''' </summary>
    ''' <param name="roleName">muss der Bezeichner einer Rolle sein</param>
    ''' <param name="typus">0: Auslastung, 1: Überauslastung 2: Unterauslastung</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAuslastungsValues(ByVal roleName As String, ByVal typus As Integer) As Double()

        Get
            Dim roleValues() As Double
            Dim kapaValues() As Double
            Dim tmpValues() As Double
            Dim myCollection As New Collection
            Dim ix As Integer
            Dim zeitraum As Integer

            Dim istSammelRolle As Boolean = RoleDefinitions.getRoledef(roleName).isCombinedRole

            zeitraum = showRangeRight - showRangeLeft
            ReDim roleValues(zeitraum)
            ReDim kapaValues(zeitraum)
            ReDim tmpValues(zeitraum)

            If RoleDefinitions.containsName(roleName) Then

                If istSammelRolle Then
                    myCollection.Add(roleName, roleName)
                    ' dann sollen alle Werte, aslo inkl der Subroles berücksichtigt werden  
                    roleValues = Me.getRoleValuesInMonth(roleName, considerAllSubRoles:=True)
                    kapaValues = Me.getRoleKapasInMonth(myCollection)

                Else

                    myCollection.Add(roleName, roleName)
                    roleValues = Me.getRoleValuesInMonth(roleName)
                    kapaValues = Me.getRoleKapasInMonth(myCollection)
                    myCollection.Clear()
                End If


                Select Case typus

                    Case 0
                        ' Auslastung

                        For ix = 0 To zeitraum
                            If roleValues(ix) > kapaValues(ix) Then
                                ' es werden die maximale Anzahl Leute dieser Rolle berücksichtigt 
                                tmpValues(ix) = tmpValues(ix) + kapaValues(ix)
                            Else
                                ' die internen Ressourcen reichen aus
                                tmpValues(ix) = tmpValues(ix) + roleValues(ix)
                            End If
                        Next ix

                    Case 1
                        ' Überauslastung

                        For ix = 0 To zeitraum
                            If roleValues(ix) > kapaValues(ix) Then
                                ' es gibt Überauslastung  
                                tmpValues(ix) = tmpValues(ix) + roleValues(ix) - kapaValues(ix)
                            Else
                                ' es gibt keine Überauslastung 

                            End If
                        Next ix

                    Case 2
                        ' Unterauslastung
                        For ix = 0 To zeitraum
                            If roleValues(ix) < kapaValues(ix) Then
                                ' es gibt Unterauslastung  
                                tmpValues(ix) = tmpValues(ix) + kapaValues(ix) - roleValues(ix)
                            Else
                                ' es gibt keine Unterauslastung 

                            End If
                        Next ix

                End Select

            End If



            getAuslastungsValues = tmpValues


        End Get

    End Property

    ''' <summary>
    ''' wird für den Massenedit benötigt 
    ''' gibt pro Projekt, Phase und Rolle bzw. Kostenart eine Zeile zurück, die die absoluten Bedarfs-Werte im betrachteten Monat enthält und ausserdem 
    ''' pro Rolle die Gesamt bzw. monatl. Auslastungswerte 
    ''' standardmäßig werden die prozentualen Auslastungswerte angezeigt; es können bei Setzung bon absoluteValues = true aich die noch freien Tage in dem Monat angezeigt werden 
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAuslastungsArray(ByVal von As Integer, ByVal bis As Integer,
                                                 ByVal percentValues As Boolean) As Double(,)
        Get
            Dim tmpArray(,) As Double
            Dim anzahlRollen As Integer = RoleDefinitions.Count
            ReDim tmpArray(anzahlRollen - 1, bis - von + 1)

            For r = 1 To RoleDefinitions.Count
                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(r)
                Dim roleUID As Integer = tmpRole.UID
                Dim roleName As String = tmpRole.name

                ' nur für Testzwecke, nachher wieder rausmachen ...
                'If r <> roleUID Then
                '    Call MsgBox("RoleID ist ungleich laufender Nummer !")
                'End If


                Dim roleValues() As Double
                Dim kapaValues() As Double
                Dim myCollection As New Collection
                Dim ix As Integer
                Dim zeitraum As Integer = bis - von

                Dim istSammelRolle As Boolean = tmpRole.isCombinedRole

                ReDim roleValues(zeitraum)
                ReDim kapaValues(zeitraum)

                myCollection.Add(roleName, roleName)
                kapaValues = Me.getRoleKapasInMonth(myCollection)
                myCollection.Clear()

                If istSammelRolle Then
                    ' alle Bedarfe berücksichtigen
                    roleValues = Me.getRoleValuesInMonth(roleIDStr:=roleUID.ToString, considerAllSubRoles:=True,
                                                         type:=PTcbr.all)
                Else
                    roleValues = Me.getRoleValuesInMonth(roleUID.ToString)
                End If

                ' jetzt wird der Array aufgebaut

                If Not percentValues Then

                    tmpArray(r - 1, 0) = kapaValues.Sum - roleValues.Sum
                    For ix = 1 To bis - von + 1
                        tmpArray(r - 1, ix) = kapaValues(ix - 1) - roleValues(ix - 1)
                    Next

                Else
                    If kapaValues.Sum > 0 Then
                        tmpArray(r - 1, 0) = roleValues.Sum / kapaValues.Sum
                    Else
                        tmpArray(r - 1, 0) = 999 ' Kennzeichen für unendlich 
                    End If

                    For ix = 1 To bis - von + 1
                        If kapaValues(ix - 1) > 0 Then
                            tmpArray(r - 1, ix) = roleValues(ix - 1) / kapaValues(ix - 1)
                        Else
                            tmpArray(r - 1, ix) = 999 ' Kennzeichen für unendlich ...
                        End If

                    Next
                End If


            Next

            getAuslastungsArray = tmpArray

        End Get
    End Property

    ''' <summary>
    ''' does an automatic allocation of people for all summary Roles / Skills  
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="variantName"></param>
    ''' <param name="errMsg"></param>
    Public Sub autoAllocate(ByVal pName As String, ByVal variantName As String,
                            ByVal allowOvertime As Boolean, ByRef errMsg As String)

        Dim hproj As clsProjekt = Nothing
        Dim placeHolderNeeds As New SortedList(Of String, Double())

        Dim projectScopeCandidates As SortedList(Of Double, Integer) = Nothing

        ' 1. create a variant, if variantName is provided 
        If variantName <> "" Then
            Dim baseProject As clsProjekt = getProject(pName)
            hproj = baseProject.createVariant(variantName, "created for auto-allocation")
        Else
            hproj = getProject(pName)
        End If


        ' get a list of summary roles used in hproj 
        Dim placeholderIDs As SortedList(Of String, Double) = hproj.getPlaceholderRoles



        ' now define freeAmount of capacity over the whole project scope ...
        Dim projectPhase As clsPhase = hproj.getPhase(1)


        ' now checkout the resource needs and available capacities
        For Each kvp As KeyValuePair(Of String, Double) In placeholderIDs

            Dim myCollection As New Collection From {
                kvp.Key
            }
            Dim bedarf() As Double = hproj.getRessourcenBedarf(kvp.Key, inclSubRoles:=False)

            If Not placeHolderNeeds.ContainsKey(kvp.Key) Then
                placeHolderNeeds.Add(kvp.Key, bedarf)
            Else
                ' kann eigentlich nicht sein ,,
                Call MsgBox("unexpected Error 3522 cP")
            End If

        Next

        ' now with resource Needs of placeHolders, possible candidates and defined priority people the Auto Allocation can be done ...
        ' for this just go over each phase 
        For p = 1 To hproj.CountPhases

            Dim cPhase As clsPhase = hproj.getPhase(p)
            Dim phasePlaceHolderNeeds As SortedList(Of String, Double) = cPhase.getRoleNameIDsAndValues(onlySummaryRoles:=True)

            ' who is already on the team ? 
            Dim peopleIDs As SortedList(Of String, Double) = hproj.getPeople()

            For Each phasePlaceHolder As KeyValuePair(Of String, Double) In phasePlaceHolderNeeds

                Dim myCurrentSkillID As Integer = -1
                Dim myCurrentRoleID As Integer = RoleDefinitions.parseRoleNameID(phasePlaceHolder.Key, myCurrentSkillID)

                Dim candidates As SortedList(Of Double, Integer) = cPhase.getCandidates(phasePlaceHolder.Key, 0.5, phasePlaceHolder.Value)
                projectScopeCandidates = projectPhase.getCandidates(phasePlaceHolder.Key, 2, phasePlaceHolder.Value)

                Dim bestCandidates As SortedList(Of Double, Integer) = calcBestCandidates(peopleIDs,
                                                                                          myCurrentSkillID,
                                                                                            candidates,
                                                                                            projectScopeCandidates,
                                                                                            phasePlaceHolder.Value)

                ' now best candidates do replace the placeHolder Role with the required Value , may Contain one or more items
                For Each substitution As KeyValuePair(Of Double, Integer) In bestCandidates
                    Dim newNameID As String = RoleDefinitions.bestimmeRoleNameID(substitution.Value, myCurrentSkillID)
                    Dim ok As Boolean = cPhase.substituteRole(phasePlaceHolder.Key, newNameID, allowOvertime, substitution.Key)
                    If Not ok Then
                        Dim msgTxt As String = phasePlaceHolder.Key & " -> " & newNameID
                        Call logger(errLevel:=ptErrLevel.logWarning, addOn:="Auto-Allocation: Substitution failed", strLog:=msgTxt)
                    End If
                Next

            Next

        Next

        ' ok, Done ! 

    End Sub



    ''' <summary>
    ''' returns an array which takes avaibale capacity into account
    ''' it redistibutes values such that availabe capacity does cover it
    ''' if there is an overutilization then the part of overutilization is distributed equally over the complete timeframe , if allowOvertime = true
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <param name="teamID"></param>
    ''' <param name="xValues">the desired distribution of values, need to be fitted to available capacity</param>
    ''' <param name="xStartDate"></param>
    ''' <param name="allowOvertime">distribute all, even it leads to overtime</param>
    ''' <param name="oldRoleValues">current values of the respective role </param>
    ''' <returns></returns>
    Public ReadOnly Property adjustToCapacity(ByVal uid As Integer, ByVal teamID As Integer, ByVal allowOvertime As Boolean, ByVal xValues As Double(), ByVal xStartDate As Date,
                                              ByVal oldRoleValues As Double()) As Double()
        Get

            Dim length As Integer = xValues.Length
            Dim result As Double()
            ReDim result(length - 1)

            Dim checkLength As Integer = oldRoleValues.Length

            Dim stillToDistribute As Double = 0

            Dim freeCapacity As Double()
            ReDim freeCapacity(length - 1)

            Dim stillFreeCapacity As Double()
            ReDim stillFreeCapacity(length - 1)

            Dim addOvertime As Double()
            ReDim addOvertime(length - 1)

            Dim von As Integer = getColumnOfDate(xStartDate)
            Dim bis As Integer = von + length - 1

            ' now remember global variables showRangeLeft and showRangeRight 
            Dim srlSav As Integer = showRangeLeft
            Dim srrSav As Integer = showRangeRight

            showRangeLeft = von
            showRangeRight = bis


            Dim myRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(uid)

            Dim myTeam As clsRollenDefinition = Nothing
            If teamID > 0 Then
                myTeam = RoleDefinitions.getRoleDefByID(teamID)
            End If

            freeCapacity = getFreeCapacityOfRole(uid, teamID, von, bis)
            ' now put the oldRoleValues on top of freeCapacity; if role had already some values this need to be added to the freeCapacity 
            ' because the oldValues will be substituted 
            If oldRoleValues.Sum > 0 Then
                For ix As Integer = 0 To length - 1
                    freeCapacity(ix) = freeCapacity(ix) + oldRoleValues(ix)
                Next
            End If



            For ix As Integer = 0 To length - 1
                If freeCapacity(ix) > xValues(ix) Then
                    'all ok, can be done 

                    If stillToDistribute > 0 Then
                        Dim freeAmount As Double = freeCapacity(ix) - xValues(ix)
                        If freeAmount >= stillToDistribute Then
                            result(ix) = xValues(ix) + stillToDistribute
                            stillToDistribute = 0
                        ElseIf freeAmount > 0 Then
                            result(ix) = xValues(ix) + freeAmount
                            stillToDistribute = stillToDistribute - freeAmount
                        End If
                    Else
                        result(ix) = xValues(ix)
                    End If

                    ' remember when there is still capacity left when it comes to distribute the amount of stillTodistribute
                    If freeCapacity(ix) > result(ix) Then
                        stillFreeCapacity(ix) = freeCapacity(ix) - result(ix)
                    End If


                ElseIf freeCapacity(ix) < xValues(ix) Then

                    If freeCapacity(ix) >= 0 Then
                        result(ix) = freeCapacity(ix)
                        stillToDistribute = stillToDistribute + xValues(ix) - freeCapacity(ix)
                    Else
                        stillToDistribute = stillToDistribute + xValues(ix)
                    End If


                Else
                    ' it is exact the same 
                    result(ix) = xValues(ix)
                End If
            Next

            If stillToDistribute > 0 Then

                ' first of all : use all stillFreeCapacity
                If stillFreeCapacity.Sum > 0 Then
                    For cf As Integer = 0 To length - 1
                        If stillFreeCapacity(cf) > 0 Then
                            If stillToDistribute > stillFreeCapacity(cf) Then
                                result(cf) = result(cf) + stillFreeCapacity(cf)
                                stillToDistribute = stillToDistribute - stillFreeCapacity(cf)
                            Else
                                result(cf) = result(cf) + stillToDistribute
                                stillToDistribute = 0
                            End If
                        End If
                        If stillToDistribute <= 0 Then
                            Exit For
                        End If
                    Next
                End If

                If allowOvertime Or myRole.isExternRole Then
                    Dim i As Integer = 0
                    Do While stillToDistribute >= 1
                        result(i) = result(i) + 1
                        stillToDistribute = stillToDistribute - 1
                        i = i + 1
                        If i > length - 1 Then
                            i = 0
                        End If
                    Loop

                    If stillToDistribute > 0 Then
                        result(i) = result(i) + stillToDistribute
                    End If
                End If


            End If

            ' now restore values from srlSav ..
            showRangeLeft = srlSav
            showRangeRight = srrSav


            adjustToCapacity = result

        End Get
    End Property

    ''' <summary>
    ''' returns the free capacity of roleID in given 
    ''' Attention: some of the values maybe negative!
    ''' </summary>
    ''' <param name="roleID"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getFreeCapacityOfRole(ByVal roleID As Integer, ByVal skillID As Integer, ByVal von As Integer, ByVal bis As Integer) As Double()
        Get
            Dim tmpArray() As Double
            ReDim tmpArray(bis - von)

            ' save showrangeLeft and showrangeRight 
            Dim srlSav As Integer = showRangeLeft
            Dim srrSav As Integer = showRangeRight

            showRangeLeft = von
            showRangeRight = bis

            Try
                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(roleID)

                Dim roleUID As Integer = tmpRole.UID
                Dim roleName As String = tmpRole.name

                Dim roleNameID = RoleDefinitions.bestimmeRoleNameID(roleID, skillID)


                Dim roleValues() As Double
                Dim allOtherValues() As Double
                Dim kapaValues() As Double
                Dim myCollection As New Collection
                Dim ix As Integer
                Dim zeitraum As Integer = bis - von

                Dim istSammelRolle As Boolean = tmpRole.isCombinedRole

                ReDim roleValues(zeitraum)
                ReDim kapaValues(zeitraum)

                'myCollection.Add(roleName, roleName)
                myCollection.Add(roleNameID, roleNameID)
                kapaValues = Me.getRoleKapasInMonth(myCollection)
                myCollection.Clear()

                If istSammelRolle Then
                    ' alle Bedarfe berücksichtigen
                    If skillID > 0 Then
                        ' only skill needs required, not all other activities 
                        roleValues = Me.getRoleValuesInMonth(roleIDStr:=roleNameID, considerAllSubRoles:=True, considerAllNeedsOfRolesHavingTheseSkills:=False,
                                                         type:=PTcbr.all)
                        allOtherValues = Me.getRoleValuesInMonth(roleIDStr:=roleNameID, considerAllSubRoles:=True, considerAllNeedsOfRolesHavingTheseSkills:=True,
                                                         type:=PTcbr.all)

                        If allOtherValues.Sum > 0 Then
                            For i As Integer = 0 To roleValues.Length - 1
                                roleValues(i) = roleValues(i) + allOtherValues(i)
                            Next
                        End If
                    Else
                        roleValues = Me.getRoleValuesInMonth(roleIDStr:=roleUID.ToString, considerAllSubRoles:=True,
                                                         type:=PTcbr.all)
                    End If


                Else
                    roleValues = Me.getRoleValuesInMonth(roleUID.ToString)
                End If

                ' jetzt wird der Array aufgebaut
                For ix = 0 To bis - von
                    tmpArray(ix) = kapaValues(ix) - roleValues(ix)
                Next



            Catch ex As Exception
                ReDim tmpArray(bis - von)
            End Try


            ' Restore showRangeLeft and showRangeRight 
            showRangeLeft = srlSav
            showRangeRight = srrSav

            getFreeCapacityOfRole = tmpArray

        End Get
    End Property

    ''' <summary>
    ''' gibt die durch Projekt-Arbeit verursachten Personalkosten zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostiValuesInMonth() As Double()

        Get
            Dim costValues() As Double
            Dim roleValues() As Double
            Dim kapaValues() As Double
            Dim alleRoleValues() As Double
            Dim alleSubRoleValues() As Double
            Dim alleKapaValues() As Double

            Dim roleName As String
            Dim roleUIDStr As String
            Dim myCollection As New Collection
            Dim i As Integer, ix As Integer
            Dim zeitraum As Integer
            Dim faktor As Double = 1

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Or awinSettings.kapaEinheit = "PD" Then
                faktor = 1
            Else
                faktor = 1
            End If


            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)
            ReDim roleValues(zeitraum)
            ReDim kapaValues(zeitraum)
            ReDim alleRoleValues(zeitraum)
            ReDim alleSubRoleValues(zeitraum)
            ReDim alleKapaValues(zeitraum)

            For i = 1 To RoleDefinitions.Count

                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(i)
                If Not IsNothing(tmpRole) Then

                    Dim istSammelRolle As Boolean = tmpRole.isCombinedRole
                    roleName = tmpRole.name
                    roleUIDStr = tmpRole.UID.ToString
                    roleValues = Me.getRoleValuesInMonth(roleUIDStr)


                    If istSammelRolle Then

                        ReDim kapaValues(zeitraum)

                        myCollection.Add(roleName, roleName)
                        alleKapaValues = Me.getRoleKapasInMonth(myCollection)
                        alleRoleValues = Me.getRoleValuesInMonth(roleIDStr:=roleUIDStr,
                                                                 considerAllSubRoles:=True,
                                                                 type:=PTcbr.all,
                                                                 excludedNames:=Nothing)
                        myCollection.Clear()

                    Else
                        myCollection.Add(roleName, roleName)
                        kapaValues = Me.getRoleKapasInMonth(myCollection)
                        myCollection.Clear()
                    End If

                    For ix = 0 To zeitraum

                        If istSammelRolle Then

                            If alleRoleValues(ix) > alleKapaValues(ix) Then
                                ' der Anteil FIG22 intern ist beschränkt auf alleKapas - SummealleSubroles ohne Sammelrolle 
                                ' 
                                alleSubRoleValues = Me.getRoleValuesInMonth(roleIDStr:=roleUIDStr,
                                                                 considerAllSubRoles:=True,
                                                                 type:=PTcbr.realRoles,
                                                                 excludedNames:=Nothing)
                                Dim diff As Double = alleKapaValues(ix) - alleSubRoleValues(ix)

                                If diff > 0 Then
                                    ' nur dann gibt es noch einen internen Anteil für den Platzhalter
                                    If roleValues(ix) <= diff Then
                                        ' der interne Teil ist maximal Diff oder eben roleValues ...
                                        diff = roleValues(ix)
                                    End If
                                End If

                                costValues(ix) = costValues(ix) +
                                                 diff * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000

                            Else
                                ' die internen Ressourcen reichen aus
                                costValues(ix) = costValues(ix) +
                                                 roleValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                            End If

                        Else

                            If roleValues(ix) > kapaValues(ix) Then
                                ' es werden die maximale Anzahl Leute dieser Rolle berücksichtigt 
                                costValues(ix) = costValues(ix) +
                                                 kapaValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                            Else
                                ' die internen Ressourcen reichen aus
                                costValues(ix) = costValues(ix) +
                                                 roleValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                            End If

                        End If

                    Next ix

                End If

            Next i


            getCostiValuesInMonth = costValues


        End Get

    End Property
    '
    ' property gibt die externen Kosten zurück, die durch die Hinzuziehung externer Ressourcen entstehen 
    '
    ''' <summary>
    ''' gibt die Kosten zurück, die für externe Kräfte ausgegeben werden , um die Projekte leisten zu können 
    ''' Ergebnis ist die Absolut Betrachtung, keine Delta Betrachtung 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCosteValuesInMonth(Optional ByVal isDeltaCalculation As Boolean = False) As Double()

        Get
            Dim costValues() As Double
            Dim roleValues() As Double
            Dim kapaValues() As Double
            Dim alleRoleValues() As Double
            Dim alleSubRoleValues() As Double
            Dim alleKapaValues() As Double

            Dim calculationValue As Double


            Dim includesOverloadCost As Boolean = True

            Dim roleName As String
            Dim roleUIDStr As String
            Dim myCollection As New Collection
            Dim i As Integer, ix As Integer
            Dim zeitraum As Integer
            Dim faktor As Double = 1

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Or awinSettings.kapaEinheit = "PD" Then
                faktor = 1
            Else
                faktor = 1
            End If

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)
            ReDim roleValues(zeitraum)
            ReDim kapaValues(zeitraum)
            ReDim alleRoleValues(zeitraum)
            ReDim alleSubRoleValues(zeitraum)
            ReDim alleKapaValues(zeitraum)

            For i = 1 To RoleDefinitions.Count

                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(i)

                If Not IsNothing(tmpRole) Then
                    Dim istSammelRolle As Boolean = tmpRole.isCombinedRole
                    roleName = tmpRole.name
                    roleUIDStr = tmpRole.UID.ToString


                    roleValues = Me.getRoleValuesInMonth(roleUIDStr)


                    ' mit welchem Wert wird gerechnet 

                    With tmpRole
                        If isDeltaCalculation Then
                            calculationValue = 0
                        Else
                            calculationValue = .tagessatzIntern

                        End If
                    End With

                    If istSammelRolle Then

                        ReDim kapaValues(zeitraum)

                        myCollection.Add(roleName, roleName)
                        alleKapaValues = Me.getRoleKapasInMonth(myCollection)
                        alleRoleValues = Me.getRoleValuesInMonth(roleIDStr:=roleUIDStr,
                                                                    considerAllSubRoles:=True,
                                                                    type:=PTcbr.all,
                                                                    excludedNames:=Nothing)
                        myCollection.Clear()

                    Else
                        myCollection.Add(roleName, roleName)
                        kapaValues = Me.getRoleKapasInMonth(myCollection)
                        myCollection.Clear()
                    End If


                    For ix = 0 To zeitraum

                        If istSammelRolle Then
                            Dim diff As Double
                            If alleRoleValues(ix) > alleKapaValues(ix) Then

                                alleSubRoleValues = Me.getRoleValuesInMonth(roleIDStr:=roleUIDStr,
                                                                 considerAllSubRoles:=True,
                                                                 type:=PTcbr.realRoles,
                                                                 excludedNames:=Nothing)

                                If alleSubRoleValues(ix) >= alleKapaValues(ix) Then
                                    ' der gesamte Platzhalter Anteil ist Overtime
                                    diff = roleValues(ix)
                                Else
                                    diff = alleRoleValues(ix) - alleKapaValues(ix)
                                End If

                                costValues(ix) = costValues(ix) +
                                                 diff * calculationValue * faktor / 1000


                            Else
                                ' die internen Ressourcen reichen aus  

                            End If

                        Else

                            If roleValues(ix) > kapaValues(ix) Then
                                ' Overtime Kosten fallen an
                                costValues(ix) = costValues(ix) +
                                                 (roleValues(ix) - kapaValues(ix)) * calculationValue * faktor / 1000
                            Else
                                ' die internen Ressourcen reichen aus

                            End If

                        End If

                    Next ix
                End If


            Next i


            getCosteValuesInMonth = costValues

        End Get

    End Property

    ' tk 11.6.2020 hier ist viel unnötig , wurde geändert siehe unten 
    '''' <summary>
    '''' gibt die Personalkosten im Zeitraum zurück, dabei werden die Überstundensätze entsprechend des optionalen Parameters berücksichtigt 
    '''' 
    '''' der optionale Parameter bestimmt, ob die Überlast-Situationen berücksichtigt werden sollen oder nicht ... 
    '''' </summary>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public ReadOnly Property getCostGpValuesInMonth(Optional ByVal includesOverloadCost As Boolean = False) As Double()

    '    Get
    '        Dim costValues() As Double
    '        Dim roleValues() As Double
    '        Dim kapaValues() As Double
    '        Dim alleRoleValues() As Double
    '        Dim alleKapaValues() As Double

    '        Dim roleName As String
    '        Dim roleUIDStr As String
    '        Dim myCollection As New Collection
    '        Dim i As Integer, ix As Integer
    '        Dim zeitraum As Integer
    '        Dim faktor As Double = 1

    '        If awinSettings.kapaEinheit = "PM" Then
    '            faktor = nrOfDaysMonth
    '        ElseIf awinSettings.kapaEinheit = "PW" Then
    '            faktor = 5
    '        ElseIf awinSettings.kapaEinheit = "PT" Or awinSettings.kapaEinheit = "PD" Then
    '            faktor = 1
    '        Else
    '            faktor = 1
    '        End If

    '        zeitraum = showRangeRight - showRangeLeft
    '        ReDim costValues(zeitraum)
    '        ReDim roleValues(zeitraum)
    '        ReDim kapaValues(zeitraum)
    '        ReDim alleRoleValues(zeitraum)
    '        ReDim alleKapaValues(zeitraum)


    '        For i = 1 To RoleDefinitions.Count

    '            Dim istSammelRolle As Boolean = RoleDefinitions.getRoledef(i).isCombinedRole
    '            roleName = RoleDefinitions.getRoledef(i).name
    '            roleUIDStr = RoleDefinitions.getRoledef(i).UID.ToString

    '            roleValues = Me.getRoleValuesInMonth(roleUIDStr)

    '            If istSammelRolle Then

    '                ReDim kapaValues(zeitraum)

    '                If includesOverloadCost Then
    '                    myCollection.Add(roleName, roleName)
    '                    alleKapaValues = Me.getRoleKapasInMonth(myCollection)
    '                    alleRoleValues = Me.getRoleValuesInMonth(roleIDStr:=roleUIDStr,
    '                                                             considerAllSubRoles:=True,
    '                                                             type:=PTcbr.all,
    '                                                             excludedNames:=Nothing)
    '                    myCollection.Clear()
    '                End If


    '            Else
    '                myCollection.Add(roleName, roleName)
    '                kapaValues = Me.getRoleKapasInMonth(myCollection)
    '                myCollection.Clear()
    '            End If




    '            For ix = 0 To zeitraum

    '                ' die internen Ressourcen reichen aus oder die Kosten durch Überlast sollen nicht berücksichtigt werden 
    '                costValues(ix) = costValues(ix) +
    '                                         roleValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000

    '            Next ix

    '        Next i


    '        getCostGpValuesInMonth = costValues

    '    End Get

    'End Property

    ''' <summary>
    ''' gibt die Personalkosten im Zeitraum zurück, 
    ''' der optionale Parameter bestimmt, ob alle, nur interne, nur externe Personalkosten übergeben werden 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostGpValuesInMonth(Optional ByVal scope As PTrt = PTrt.all) As Double()

        Get
            Dim costValues() As Double
            Dim roleValues() As Double


            Dim roleUIDStr As String

            Dim i As Integer, ix As Integer
            Dim zeitraum As Integer
            Dim faktor As Double = 1

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Or awinSettings.kapaEinheit = "PD" Then
                faktor = 1
            Else
                faktor = 1
            End If

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)
            ReDim roleValues(zeitraum)




            For i = 1 To RoleDefinitions.Count

                Dim curRole As clsRollenDefinition = RoleDefinitions.getRoledef(i)
                Dim dailyRate As Double = curRole.tagessatzIntern

                roleUIDStr = curRole.UID.ToString

                If scope = PTrt.all Or
                        (scope = PTrt.intern And Not curRole.isExternRole) Or
                        (scope = PTrt.extern And curRole.isExternRole) Then

                    roleValues = Me.getRoleValuesInMonth(roleUIDStr)


                    For ix = 0 To zeitraum

                        ' die internen Ressourcen reichen aus oder die Kosten durch Überlast sollen nicht berücksichtigt werden 
                        costValues(ix) = costValues(ix) +
                                                 roleValues(ix) * dailyRate * faktor / 1000

                    Next ix

                End If



            Next i


            getCostGpValuesInMonth = costValues

        End Get

    End Property

    ' '' ''' <summary>
    ' '' ''' gibt die Mehrkosten, die im Zeitraum durch den Einsatz von Externen verursacht werden , zurück 
    ' '' ''' der Wert repräsentiert dabei den Unterschied zu den Kosten, die durch den Einsatz von Internen anfallen würden
    ' '' ''' </summary>
    ' '' ''' <value></value>
    ' '' ''' <returns></returns>
    ' '' ''' <remarks></remarks>
    ' ''Public ReadOnly Property getadditionalECostinMonth() As Double()

    ' ''    Get
    ' ''        Dim costValues() As Double
    ' ''        Dim roleValues() As Double
    ' ''        Dim kapaValues() As Double
    ' ''        Dim alleRoleValues() As Double
    ' ''        Dim alleKapaValues() As Double

    ' ''        Dim roleName As String
    ' ''        Dim myCollection As New Collection
    ' ''        Dim i As Integer, ix As Integer
    ' ''        Dim zeitraum As Integer
    ' ''        Dim tagessatzDifferenz As Double
    ' ''        Dim faktor As Double = nrOfDaysMonth

    ' ''        If awinSettings.kapaEinheit = "PM" Then
    ' ''            faktor = nrOfDaysMonth
    ' ''        ElseIf awinSettings.kapaEinheit = "PW" Then
    ' ''            faktor = 5
    ' ''        ElseIf awinSettings.kapaEinheit = "PT" Then
    ' ''            faktor = 1
    ' ''        Else
    ' ''            faktor = 1
    ' ''        End If

    ' ''        zeitraum = showRangeRight - showRangeLeft
    ' ''        ReDim costValues(zeitraum)
    ' ''        ReDim roleValues(zeitraum)
    ' ''        ReDim kapaValues(zeitraum)
    ' ''        ReDim alleRoleValues(zeitraum)
    ' ''        ReDim alleKapaValues(zeitraum)

    ' ''        For i = 1 To RoleDefinitions.Count

    ' ''            Dim istSammelRolle As Boolean = RoleDefinitions.getRoledef(i).isCombinedRole
    ' ''            roleName = RoleDefinitions.getRoledef(i).name
    ' ''            roleValues = Me.getRoleValuesInMonth(roleName)

    ' ''            If istSammelRolle Then

    ' ''                ReDim kapaValues(zeitraum)

    ' ''                myCollection.Add(roleName, roleName)
    ' ''                alleKapaValues = Me.getRoleKapasInMonth(myCollection, False)
    ' ''                alleRoleValues = Me.getRoleValuesInMonth(roleName, True)
    ' ''                myCollection.Clear()

    ' ''            Else
    ' ''                myCollection.Add(roleName, roleName)
    ' ''                kapaValues = Me.getRoleKapasInMonth(myCollection, False)
    ' ''                myCollection.Clear()
    ' ''            End If


    ' ''            With RoleDefinitions.getRoledef(roleName)
    ' ''                tagessatzDifferenz = .tagessatzExtern - .tagessatzIntern
    ' ''            End With

    ' ''            For ix = 0 To zeitraum

    ' ''                If istSammelRolle Then


    ' ''                Else
    ' ''                    If roleValues(ix) > kapaValues(ix) Then
    ' ''                        ' externe Ressourcen müssen hinzugezogen werden
    ' ''                        costValues(ix) = costValues(ix) + _
    ' ''                                         (roleValues(ix) - kapaValues(ix)) * tagessatzDifferenz * faktor / 1000
    ' ''                    Else
    ' ''                        ' die internen Ressourcen reichen aus

    ' ''                    End If
    ' ''                End If

    ' ''            Next ix

    ' ''        Next i


    ' ''        getadditionalECostinMonth = costValues

    ' ''    End Get

    ' ''End Property

    ''' <summary>
    ''' gibt für den betrachteten Zeotraum die Ergebnisskennzahl zurück 
    ''' ergebniskennzahl = (zeitraumbudget - (zeitraumCost + zeitraumRisiko))-(Überlast-Kosten + Opp.-Kosten))
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getErgebniskennzahl() As Double
        Get

            Dim zeitraumBudget As Double
            Dim zeitraumCost As Double
            Dim zeitraumRisiko As Double
            Dim earnedValue As Double
            Dim additionalCostExt As Double
            Dim internwithoutProject As Double
            Dim ertragsWert As Double


            ' Ausrechnen amteiliges Budget, das im Zeitraum zur Verfügung steht und der im Zeitraum anfallenden Kosten  
            zeitraumBudget = System.Math.Round(ShowProjekte.getBudgetValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
            zeitraumCost = System.Math.Round(ShowProjekte.getTotalCostValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10

            ' das ist der Risiko Abschlag  
            zeitraumRisiko = System.Math.Round(ShowProjekte.getWeightedRiskValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10


            ' das ist der Earned Value 
            earnedValue = zeitraumBudget - (zeitraumCost + zeitraumRisiko)


            ' das sind die Zusatzkosten, die durch Externe bzw. Überstunden (wg Überauslastung) verursacht werden
            additionalCostExt = System.Math.Round(ShowProjekte.getCosteValuesInMonth(True).Sum / 10, mode:=MidpointRounding.ToEven) * 10

            ' das sind die durch Unterauslastung verursachten Kosten , also Personal-Kosten von Leuten, die in keinem Projekt sind
            internwithoutProject = System.Math.Round(ShowProjekte.getCostoValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10

            ' das ist der Ertrag 
            ertragsWert = earnedValue - (additionalCostExt + internwithoutProject)

            getErgebniskennzahl = ertragsWert

        End Get
    End Property


    ''' <summary>
    ''' gibt die Summe an "bad cost" an, das sind die durch Einsatz externer Kräfte verursachten zusätzlichen Kosten und die 
    ''' durch untätige Ressourcen verursachten Kosten der übergebenen Rolle(n im betrachteten Zeitraum 
    ''' wird für die Optimierung der Ressourcen Diagramm Verläufe zugrundegelegt
    ''' </summary>
    ''' <param name="roleCollection"></param>
    ''' <value></value>
    ''' <returns>einen Double Wert , der die Gesamt Summe an bad cost enthält</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getbadCostOfRole(ByVal roleCollection As Collection) As Double
        Get
            Dim roleValues() As Double
            Dim kapaValues() As Double
            Dim costValue As Double
            Dim myCollection As New Collection

            Dim zeitraum As Integer
            Dim tagessatzIntern As Double
            Dim roleName As String
            Dim i As Integer
            Dim faktor As Double = 1

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Or awinSettings.kapaEinheit = "PD" Then
                faktor = 1
            Else
                faktor = 1
            End If

            zeitraum = showRangeRight - showRangeLeft
            costValue = 0.0

            For i = 1 To roleCollection.Count
                ReDim roleValues(zeitraum)
                ReDim kapaValues(zeitraum)
                roleName = CStr(roleCollection.Item(i))

                ' Änderung tk: Berücksichtigung von SammelRollen 
                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)

                If Not IsNothing(tmpRole) Then
                    tagessatzIntern = tmpRole.tagessatzIntern

                    ' tk 30.11. es gibt keine externen Kostensätze mehr 
                    'If tagessatzExtern <> tagessatzIntern Then
                    '    diff = tagessatzExtern - tagessatzIntern
                    '    myCollection.Add(roleName, roleName)

                    '    If tmpRole.isCombinedRole Then
                    '        roleValues = Me.getRoleValuesInMonth(roleID:=roleName,
                    '                                             considerAllSubRoles:=True,
                    '                                             type:=PTcbr.all,
                    '                                             excludedNames:=roleCollection)

                    '    Else
                    '        roleValues = Me.getRoleValuesInMonth(roleID:=roleName)
                    '    End If

                    '    kapaValues = Me.getRoleKapasInMonth(myCollection)
                    '    myCollection.Clear()

                    '    For ix = 0 To zeitraum
                    '        If roleValues(ix) > kapaValues(ix) Then
                    '            ' Kosten der externen Ressourcen
                    '            costValue = costValue +
                    '                             (roleValues(ix) - kapaValues(ix)) * diff * faktor / 1000
                    '        ElseIf roleValues(ix) < kapaValues(ix) Then
                    '            ' Kosten der internen Ressourcen, die nicht in Projekten arbeiten  
                    '            costValue = costValue +
                    '                             (kapaValues(ix) - roleValues(ix)) * tagessatzIntern * faktor / 1000

                    '        End If
                    '    Next ix
                    'End If
                End If


            Next i

            getbadCostOfRole = costValue

        End Get
    End Property

    ''' <summary>
    ''' gibt für die übergebenen Phasen/Rollen/Kostenarten im betrachteten Zeitraum den Durchschnittswert an 
    ''' </summary>
    ''' <param name="myCollection">enthält die zu betrachtenden Phasen/Rollen/Kostenarten</param>
    ''' <param name="diagrammtyp">gibt an, worum es sich handelt: Phase, Rolle, Kostenart; 
    ''' andere Werte werden aktuell nicht unterstützt </param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAverage(ByVal myCollection As Collection, ByVal diagrammtyp As String) As Double
        Get

            Dim tmpValues() As Double
            Dim tmpValues2(,) As Double
            Dim tmpSum As Double
            Dim zwSum As Double

            Dim zeitraum As Integer
            Dim rcName As String
            Dim i As Integer
            Dim anzElements As Integer

            anzElements = myCollection.Count
            zeitraum = showRangeRight - showRangeLeft
            'ReDim tmpValues2(3, zeitraum)

            tmpSum = 0.0

            For i = 1 To myCollection.Count
                rcName = CStr(myCollection.Item(i))

                ReDim tmpValues(zeitraum)

                If diagrammtyp = DiagrammTypen(0) Then
                    tmpValues = Me.getCountPhasesInMonth(rcName, "", -1, "")
                ElseIf diagrammtyp = DiagrammTypen(1) Then
                    tmpValues = Me.getRoleValuesInMonth(rcName)
                ElseIf diagrammtyp = DiagrammTypen(2) Then
                    tmpValues = Me.getCostValuesInMonth(rcName)
                ElseIf diagrammtyp = DiagrammTypen(5) Then
                    tmpValues2 = Me.getCountMilestonesInMonth(rcName, "", -1, "")

                    For ix = 0 To zeitraum
                        zwSum = 0
                        For ik = 0 To 3
                            zwSum = zwSum + tmpValues2(ik, ix)
                        Next
                        tmpValues(ix) = zwSum
                    Next
                End If


                tmpSum = tmpSum + tmpValues.Sum

            Next i

            getAverage = tmpSum / (zeitraum + 1)

        End Get
    End Property

    ''' <summary>
    ''' gibt für den betrachteten Zeitraum showrangeleft und showrangeright die Abweichung vom Durchschnitt an  
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="avgValue"></param>
    ''' <param name="diagrammTyp"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getDeviationfromAverage(ByVal myCollection As Collection, ByVal avgValue As Double, ByVal diagrammTyp As String) As Double

        Get
            Dim rcValues() As Double, tmpValues() As Double
            Dim sumAboveAvg As Double, tmpSum As Double
            Dim ix As Integer
            Dim zeitraum As Integer
            Dim rcName As String
            Dim i As Integer
            Dim anzElements As Integer

            anzElements = myCollection.Count
            zeitraum = showRangeRight - showRangeLeft
            ReDim rcValues(zeitraum)
            tmpSum = 0.0

            For i = 1 To myCollection.Count
                rcName = CStr(myCollection.Item(i))

                ReDim tmpValues(zeitraum)

                If diagrammTyp = DiagrammTypen(0) Then
                    tmpValues = Me.getCountPhasesInMonth(rcName, "", -1, "")
                ElseIf diagrammTyp = DiagrammTypen(1) Then
                    tmpValues = Me.getRoleValuesInMonth(rcName)
                ElseIf diagrammTyp = DiagrammTypen(2) Then
                    tmpValues = Me.getCostValuesInMonth(rcName)
                End If


                For ix = 0 To zeitraum
                    rcValues(ix) = rcValues(ix) + tmpValues(ix)
                Next ix

            Next i

            sumAboveAvg = 0.0

            For ix = 0 To zeitraum

                sumAboveAvg = sumAboveAvg + (rcValues(ix) - avgValue) * (rcValues(ix) - avgValue)

            Next ix

            getDeviationfromAverage = sumAboveAvg


        End Get
    End Property

    ''' <summary>
    ''' returns in T€ rated values of non-utilized capacity in roles given in myCustomerRole.specfics
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getNotUtilizedCapaValuesInMonth() As Double()
        Get
            Dim costValues() As Double
            Dim curCostValues() As Double
            Dim zeitraum As Integer

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)

            Dim IDArray() As Integer = RoleDefinitions.getIDArray(myCustomUserRole.specifics)

            If IsNothing(IDArray) Then
                ReDim IDArray(1)
                IDArray(1) = -1 ' das bedeutet nachher: betrachte einfache alle Rollen in der Organisation als Summe
            End If

            For Each roleID As Integer In IDArray
                curCostValues = getCostoValuesInMonth(roleID, provideKUGData:=True)
                For ix As Integer = 0 To zeitraum
                    costValues(ix) = costValues(ix) + curCostValues(ix)
                Next
            Next

            getNotUtilizedCapaValuesInMonth = costValues

        End Get
    End Property

    ''' <summary>
    ''' returns costs which occur through underutilization 
    ''' when KUGData is required there is no summation over the timespan
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>im Falle Kurzarbeit wird wird das für jeden Monat betrachtet, es gibt keinen ausgleich ! also Mrz ist unterausgelastet, Apr ist überausgelastet</remarks>
    Public ReadOnly Property getCostoValuesInMonth(Optional ByVal topNodeID As Integer = -1,
                                                   Optional ByVal provideKUGData As Boolean = False,
                                                   Optional ByVal strictly As Boolean = False) As Double()

        Get
            Dim costValues() As Double
            Dim roleValues() As Double

            Dim kapaValues() As Double
            Dim roleName As String
            Dim myCollection As New Collection
            Dim i As Integer, ix As Integer
            Dim zeitraum As Integer
            Dim faktor As Double = 1

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Or awinSettings.kapaEinheit = "PD" Then
                faktor = 1
            Else
                faktor = 1
            End If

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)
            ReDim roleValues(zeitraum)

            ReDim kapaValues(zeitraum)



            For i = 1 To RoleDefinitions.Count

                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(i)
                Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpRole.UID, -1)
                Dim weitermachen As Boolean = True


                If topNodeID <> -1 And tmpRole.UID <> topNodeID Then
                    weitermachen = RoleDefinitions.hasAnyChildParentRelationsship(roleNameID:=roleNameID, summaryRoleID:=topNodeID)
                End If

                If weitermachen Then

                    If Not IsNothing(tmpRole) Then
                        Dim dailyRate As Double = tmpRole.tagessatzIntern

                        If Not tmpRole.isExternRole Then
                            Dim istSammelRolle As Boolean = tmpRole.isCombinedRole

                            roleName = tmpRole.name
                            roleValues = Me.getRoleValuesInMonth(tmpRole.UID.ToString)

                            If istSammelRolle Then
                                ReDim kapaValues(zeitraum)
                            Else
                                myCollection.Add(roleName, roleName)
                                kapaValues = Me.getRoleKapasInMonth(myCollection)
                                myCollection.Clear()
                            End If



                            For ix = 0 To zeitraum

                                If istSammelRolle Then
                                    ' das sind ja Platzhalter Bedarfe, die irgendwann von irgendeiner realen Ressource wahrgenommen werden
                                    ' deswegen müssen die von den anderen abgezogen werden ... weil sonst zuviel als "ohne Arbeit" ausgewiesen wird 
                                    costValues(ix) = costValues(ix) - roleValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                                Else

                                    ' wenn rolevalue > kapavalues, dann enstehen negativeZahlen - die müssen dann nachher verrechnet werden ...

                                    ' d.h wenn Achim Überstunden macht, dann werden die Überstunden mit der Unterauslastung von Annabell verrechnet - sofern beide topNodeID als Parent haben 
                                    If strictly Then
                                        ' Überstunden in einem Monat verringern nicht (!) die Unterauslastungen in einem anderen Monat
                                        If kapaValues(ix) - roleValues(ix) > 0 Then
                                            costValues(ix) = costValues(ix) +
                                                 (kapaValues(ix) - roleValues(ix)) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                                        End If
                                    Else
                                        costValues(ix) = costValues(ix) +
                                                 (kapaValues(ix) - roleValues(ix)) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                                    End If




                                End If


                            Next ix
                        End If

                    End If

                End If


            Next i

            ' falls die costValues negativ sind .. korrigieren auf Null 
            ' aber so dass die positiven durch die negativen erniedrigt werden ... 
            ' andernfalls werden Unterauslastungen ausgewiesen, wenn ein einiziger Monta eine Unterauslastung hat, alle anderen sind dramatisch überausgelastet ... 

            If provideKUGData Then
                ' nur die Elemente, die negativ sind, also Überauslastung haben, auf Null setzen 
                For ix = 0 To zeitraum

                    If costValues(ix) < 0 Then
                        costValues(ix) = 0
                    End If

                Next ix
            Else
                If costValues.Sum <= 0 Then
                    ReDim costValues(zeitraum)
                Else
                    ' if there are months with over and under-utilization then smoothen it out 
                    Dim hasAnyNegative As Boolean = False
                    For ix = 0 To zeitraum
                        If costValues(ix) < 0 Then
                            hasAnyNegative = True
                            Exit For
                        End If
                    Next

                    If hasAnyNegative Then
                        ' jetzt muss einfach die Gesamt-Summe darauf verteilt werden 
                        Dim checkSum(0) As Double
                        checkSum(0) = costValues.Sum
                        costValues = calcVerteilungAufMonate(StartofCalendar.AddMonths(showRangeLeft), StartofCalendar.AddMonths(showRangeRight), checkSum, 1.0)
                    End If
                End If
            End If





            getCostoValuesInMonth = costValues


        End Get

    End Property

    '' ''' <summary>
    '' ''' Konstruktor: intilaisert die sortierte Liste der Projekte und Shapes   
    '' ''' </summary>
    '' ''' <remarks></remarks>
    Public Sub New()

        _allProjects = New SortedList(Of String, clsProjekt)
        _allShapes = New SortedList(Of String, String)

    End Sub

End Class
