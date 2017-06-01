
Imports xlNS = Microsoft.Office.Interop.Excel

Public Class clsProjekte

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
            Dim pname As String = project.name
            Dim shpUID As String = project.shpUID

            If Not IsNothing(project) Then
                _allProjects.Add(pname, project)

                If shpUID <> "" Then
                    _allShapes.Add(shpUID, pname)
                End If

                If updateCurrentConstellation Then
                    currentConstellationName = calcLastSessionScenarioName()

                    Dim key As String = calcProjektKey(project)
                    If currentSessionConstellation.contains(key, False) Then
                        Call currentSessionConstellation.setItemToShow(key, True)
                    End If

                End If

                '' mit diesem Vorgang wird die Konstellation geändert , deshalb muss die currentConstellation zurückgesetzt werden 
                'If Not currentConstellationName.EndsWith("(*)") And currentConstellationName <> "Last" Then
                '    currentConstellationName = currentConstellationName & "(*)"
                'End If
            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try


    End Sub


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
            Dim SID As String = _allProjects.Item(projectname).shpUID
            Dim vName As String = _allProjects.Item(projectname).variantName
            _allProjects.Remove(projectname)
            If SID <> "" Then
                _allShapes.Remove(SID)
            End If

            If updateCurrentConstellation Then
                Dim key As String = calcProjektKey(projectname, vName)

                If currentSessionConstellation.contains(key, False) Then
                    Call currentSessionConstellation.setItemToShow(key, False)
                End If

            End If
            '' mit diesem Vorgang wird die Konstellation geändert , deshalb muss das zurückgesetzt werden 
            'If Not currentConstellationName.EndsWith("(*)") And currentConstellationName <> "Last" Then
            '    currentConstellationName = currentConstellationName & "(*)"
            'End If

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
            If _allProjects.ContainsKey(key) Then
                contains = True
            Else
                contains = False
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

    ' wurde ersetzt durch andere getPhaseNAmes
    '' ''' <summary>
    '' ''' gibt eine sortierte Liste der vorkommenden Phasen Namen in der Menge von Projekten zurück 
    '' ''' </summary>
    '' ''' <value></value>
    '' ''' <returns></returns>
    '' ''' <remarks></remarks>
    ''Public ReadOnly Property getPhaseNames() As Collection

    ''    Get

    ''        Dim tmpListe As New Collection
    ''        Dim cphase As clsPhase
    ''        Dim phaseName As String

    ''        For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

    ''            Try
    ''                ' beginnt bei 2, weil die 1.Phase immer die mit der Projektlänge identische Phase ist ...
    ''                For p = 2 To kvp.Value.CountPhases
    ''                    cphase = kvp.Value.getPhase(p)
    ''                    phaseName = cphase.name

    ''                    If tmpListe.Contains(phaseName) Then
    ''                        ' nichts tun 
    ''                    Else
    ''                        tmpListe.Add(phaseName, phaseName)
    ''                    End If


    ''                Next
    ''            Catch ex As Exception

    ''            End Try


    ''        Next

    ''        getPhaseNames = tmpListe

    ''    End Get
    ''End Property


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
            ' neu Ende

            ' alt : ohne Ausnutzung Hierarchy ...
            ''For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

            ''    Try
            ''        For p = 1 To kvp.Value.CountPhases

            ''            cphase = kvp.Value.getPhase(p)
            ''            For r = 1 To cphase.countMilestones

            ''                msName = cphase.getMilestone(r).name
            ''                If tmpListe.Contains(msName) Then
            ''                Else
            ''                    tmpListe.Add(msName, msName)
            ''                End If

            ''            Next

            ''        Next
            ''    Catch ex As Exception

            ''    End Try


            ''Next

            getMilestoneNames = tmpListe

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
    ''' bestimmt die kleinste auftretende Spalten-Column über alle Projekte  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMinMonthColumn() As Integer
        Get
            Dim tmpMin As Integer = 10000
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                If kvp.Value.Start < tmpMin Then
                    tmpMin = kvp.Value.Start
                End If
            Next
            getMinMonthColumn = tmpMin
        End Get
    End Property

    ''' <summary>
    ''' bestimmt die größte auftretende Spalten-Column über alle Projekte  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMaxMonthColumn() As Integer
        Get
            Dim tmpMax As Integer = 0
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects
                Dim endeCol As Integer = getColumnOfDate(kvp.Value.endeDate)
                If endeCol > tmpMax Then
                    tmpMax = endeCol
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
    Public ReadOnly Property getProject(itemName As String, _
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

                    If (.Start + .StartOffset > bis) Or (.Start + .StartOffset + .anzahlRasterElemente - 1 < von) Then
                        ' dann liegt das Projekt ausserhalb des Zeitraums und muss überhaupt nicht berücksichtig werden 
                    Else

                        Select Case selectionType

                            Case PTpsel.alle
                                tmpListe.Add(kvp.Key, kvp.Key)

                            Case PTpsel.laufend

                                If DateDiff(DateInterval.Day, .startDate, Date.Now) > 0 And _
                                    .Status <> ProjektStatus(3) And _
                                    .Status <> ProjektStatus(4) Then

                                    tmpListe.Add(kvp.Key, kvp.Key)

                                End If

                            Case PTpsel.lfundab

                                If DateDiff(DateInterval.Day, .startDate, Date.Now) > 0 Then

                                    tmpListe.Add(kvp.Key, kvp.Key)

                                End If

                            Case PTpsel.abgeschlossen

                                If DateDiff(DateInterval.Day, .startDate, Date.Now) > 0 And _
                                   (.Status = ProjektStatus(3) Or _
                                   .Status = ProjektStatus(4)) Then

                                    tmpListe.Add(kvp.Key, kvp.Key)

                                End If

                        End Select


                    End If
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
                                (type = PTProjektType.projekt And pvName = kvp.Value.name) Or _
                                (type = PTProjektType.vorlage And pvName = kvp.Value.VorlagenName) Then

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

                                If type = -1 Or _
                                    (type = PTProjektType.projekt And pvName = kvp.Value.name) Or _
                                    (type = PTProjektType.vorlage And pvName = kvp.Value.VorlagenName) Then

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
                    Call splitHryFullnameTo2(CStr(myCollection.Item(cix)), elemName, breadCrumb, type, pvName)

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

                            If type = -1 Or _
                                (type = PTProjektType.projekt And pvName = kvp.Value.name) Or _
                                (type = PTProjektType.vorlage And pvName = kvp.Value.VorlagenName) Then

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

                                    tempArray = hproj.getRessourcenBedarf(elemName)

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

                            If type = -1 Or _
                                (type = PTProjektType.projekt And pvName = kvp.Value.name) Or _
                                (type = PTProjektType.vorlage And pvName = kvp.Value.VorlagenName) Then

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

                If type = -1 Or _
                    (type = PTProjektType.vorlage And pvName = hproj.VorlagenName) Or _
                    (type = PTProjektType.projekt And pvName = hproj.name) Then
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
    Public ReadOnly Property getCountPhasesInMonth(phaseName As String, ByVal breadcrumb As String, _
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

                If type = -1 Or _
                    (type = PTProjektType.vorlage And pvName = hproj.VorlagenName) Or _
                    (type = PTProjektType.projekt And pvName = hproj.name) Then
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
    '
    '
    '
    ''' <summary>
    ''' bestimmt für den betrachteten Zeitraum für die angegebene Rolle die benötigte Summe pro Monat; roleid wird als String oder Key(Integer) übergeben
    ''' </summary>
    ''' <param name="roleID"></param>
    ''' <value>String für Rollenbezeichner oder Integer für den Key der Rolle</value>
    ''' <returns>Array, der die Werte der gefragten Rolle pro Monat des betrachteten Zeitraums enthält</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleValuesInMonth(ByVal roleID As Object, _
                                                  Optional ByVal considerAllSubRoles As Boolean = False, _
                                                  Optional ByVal type As Integer = PTcbr.all, _
                                                  Optional ByVal excludedNames As Collection = Nothing) As Double()

        Get
            Dim roleValues() As Double
            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim anzProjekte As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt
            Dim lookforIndex As Boolean
            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer
            Dim roleName As String

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            lookforIndex = IsNumeric(roleID)
            zeitraum = showRangeRight - showRangeLeft
            ReDim roleValues(zeitraum)

            If lookforIndex Then
                roleName = RoleDefinitions.getRoledef(CInt(roleID)).name
            Else
                roleName = CStr(roleID)
            End If

            Dim toDoCollection As New Collection
            ' wenn considerAllSubroles  = true , dann muss 

            If considerAllSubRoles Then
                toDoCollection = RoleDefinitions.getSubRoleNamesOf(roleName, type:=type, excludedNames:=excludedNames)
                ' Änderung tk: das Folgende darf nicht mehr drin sein, da ja das Kommando getSubRoleNamesOf jetzt alles erledigt 
                'If Not toDoCollection.Contains(roleName) Then
                '    toDoCollection.Add(roleName, roleName)
                'End If
            Else
                toDoCollection.Add(roleName, roleName)
            End If



            anzProjekte = _allProjects.Count

            ' anzPhasen = AllPhases.Count

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

                    Try

                        ' hier muss die Schleife für alle Items aus toDoCollection hin 
                        For k = 1 To toDoCollection.Count

                            Dim curRole As String = CStr(toDoCollection.Item(k))
                            tempArray = hproj.getRessourcenBedarf(curRole)

                            For i = 0 To anzLoops - 1
                                roleValues(ixZeitraum + i) = roleValues(ixZeitraum + i) + tempArray(ix + i)
                            Next i

                        Next k



                    Catch ex As Exception

                    End Try


                End If

            Next kvp



            getRoleValuesInMonth = roleValues

        End Get

    End Property

    ''' <summary>
    ''' bestimmt für den betrachteten Zeitraum für die angegebene Rolle die benötigte Summe pro Monat; roleid wird als String oder Key(Integer) übergeben
    ''' </summary>
    ''' <param name="roleID"></param>
    ''' <value>String für Rollenbezeichner oder Integer für den Key der Rolle</value>
    ''' <returns>Array, der die Werte der gefragten Rolle pro Monat des betrachteten Zeitraums enthält</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleValuesInMonthNew(ByVal roleID As Object, _
                                                  Optional ByVal considerAllSubRoles As Boolean = False, _
                                                  Optional ByVal type As Integer = PTcbr.all, _
                                                  Optional ByVal excludedNames As Collection = Nothing) As Double()

        Get
            Dim roleValues() As Double
            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim anzProjekte As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt
            Dim lookforIndex As Boolean
            Dim tempArray() As Double
            Dim testArray() As Double
            Dim prAnfang As Integer, prEnde As Integer
            Dim roleName As String

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            lookforIndex = IsNumeric(roleID)
            zeitraum = showRangeRight - showRangeLeft
            ReDim roleValues(zeitraum)

            If lookforIndex Then
                roleName = RoleDefinitions.getRoledef(CInt(roleID)).name
            Else
                roleName = CStr(roleID)
            End If

            Dim toDoCollection As New Collection
            ' wenn considerAllSubroles  = true , dann muss 

            If considerAllSubRoles Then
                toDoCollection = RoleDefinitions.getSubRoleNamesOf(roleName, type:=type, excludedNames:=excludedNames)
                ' Änderung tk: das Folgende darf nicht mehr drin sein, da ja das Kommando getSubRoleNamesOf jetzt alles erledigt 
                'If Not toDoCollection.Contains(roleName) Then
                '    toDoCollection.Add(roleName, roleName)
                'End If
            Else
                toDoCollection.Add(roleName, roleName)
            End If



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

                    Dim listOfRoles As Collection = hproj.rcLists.getRoleNames

                    Try

                        ' hier muss die Schleife für alle Items aus toDoCollection hin 
                        For k = 1 To toDoCollection.Count
                            Dim curRole As String = CStr(toDoCollection.Item(k))


                            If listOfRoles.Contains(curRole) Then
                                tempArray = hproj.getRessourcenBedarfNew(curRole)

                                For i = 0 To anzLoops - 1
                                    roleValues(ixZeitraum + i) = roleValues(ixZeitraum + i) + tempArray(ix + i)
                                Next i
                            End If


                        Next k



                    Catch ex As Exception

                    End Try


                End If

            Next kvp



            getRoleValuesInMonthNew = roleValues

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
    ''' <param name="includingExterns">gibt an, ob die Werte inkl. der Externen zurückgegeben werden soll</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleKapasInMonth(ByVal myCollection As Collection, _
                                                 ByVal includingExterns As Boolean) As Double()

        Get
            Dim kapaValues() As Double
            Dim tmpValues() As Double

            Dim hkapa As Double
            Dim rname As String
            Dim zeitraum As Integer
            Dim r As Integer, m As Integer


            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            ' hier muss überprüft werden, welche Rollen denn Sammelrollen sind und deswegen ersetzt werden müssen durch ihre
            ' subroles ... 

            Dim realCollection As New Collection
            Dim sammelRollenCollection As New Collection

            For ix As Integer = 1 To myCollection.Count
                Dim roleName As String = CStr(myCollection.Item(ix))
                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)

                If Not IsNothing(tmpRole) Then

                    If tmpRole.isCombinedRole Then
                        ' es handelt sich um eine Sammelrolle
                        ' Kapas sind nur in den realRoles , also den nicht Sammelrollen vorhanden ...
                        Dim tmpCollection As Collection = RoleDefinitions.getSubRoleNamesOf(roleName:=roleName, _
                                                                                            type:=PTcbr.realRoles, _
                                                                                            excludedNames:=myCollection)

                        If tmpCollection.Count = 0 Then

                            If Not realCollection.Contains(roleName) Then
                                realCollection.Add(roleName, roleName)
                            End If

                        Else
                            ' jetzt müssen alle Elemente von tmpCollection aufgenommen werden, sofern sie nicht schon eh aufgenommen sind 
                            ' die Sammelrolle wird nicht betrachtet ... 

                            If Not sammelRollenCollection.Contains(roleName) Then
                                sammelRollenCollection.Add(roleName, roleName)
                            End If

                            For k As Integer = 1 To tmpCollection.Count
                                roleName = CStr(tmpCollection.Item(k))

                                If Not realCollection.Contains(roleName) Then
                                    realCollection.Add(roleName, roleName)
                                End If

                            Next

                        End If

                    Else

                        If Not realCollection.Contains(roleName) Then
                            realCollection.Add(roleName, roleName)
                        End If

                    End If

                End If

            Next


            ' RealCollection enthält jetzt all die gesuchten Sub-Roles und ggf separat angegebenen Rollen  

            zeitraum = showRangeRight - showRangeLeft
            ReDim kapaValues(zeitraum)
            ReDim tmpValues(zeitraum)


            For r = 1 To realCollection.Count
                rname = CStr(realCollection.Item(r))
                hkapa = RoleDefinitions.getRoledef(rname).defaultKapa

                For i = showRangeLeft To showRangeRight
                    If includingExterns Then
                        tmpValues(i - showRangeLeft) = RoleDefinitions.getRoledef(rname).kapazitaet(i) + _
                                                        RoleDefinitions.getRoledef(rname).externeKapazitaet(i)
                    Else
                        tmpValues(i - showRangeLeft) = RoleDefinitions.getRoledef(rname).kapazitaet(i)
                    End If

                    If tmpValues(i - showRangeLeft) < 0 Then
                        tmpValues(i - showRangeLeft) = hkapa
                    End If
                Next


                For m = 0 To zeitraum
                    ' Änderung 27.7 Holen der Kapa Werte , jetzt aufgeschlüsselt nach 
                    'kapaValues(m) = kapaValues(m) + hkapa
                    kapaValues(m) = kapaValues(m) + tmpValues(m)
                Next m

                'hkapa = 0
            Next r

            ' falls es SammelRollen gibt, müssen deren externe Kapas noch berücksichtigt werden ... 

            If includingExterns And sammelRollenCollection.Count > 0 Then

                ReDim tmpValues(zeitraum)
                For r = 1 To sammelRollenCollection.Count

                    rname = CStr(sammelRollenCollection.Item(r))

                    For i = showRangeLeft To showRangeRight

                        tmpValues(i - showRangeLeft) = RoleDefinitions.getRoledef(rname).externeKapazitaet(i)
                        If tmpValues(i - showRangeLeft) < 0 Then
                            tmpValues(i - showRangeLeft) = 0
                        End If

                    Next


                    For m = 0 To zeitraum
                        ' Änderung 27.7 Holen der Kapa Werte , jetzt aufgeschlüsselt nach 
                        'kapaValues(m) = kapaValues(m) + hkapa
                        kapaValues(m) = kapaValues(m) + tmpValues(m)
                    Next m


                Next


            End If

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

                    'If projektMarge < 0 Then
                    '    ' jetzt wird das Gewicht als Quotient Risiko / strategic Fit betrachtet 
                    '    If .StrategicFit > 1 Then
                    '        SRweight = .Risiko / .StrategicFit
                    '    Else
                    '        SRweight = .Risiko
                    '    End If
                    '    If SRweight = 0 Then
                    '        SRweight = 1
                    '    End If
                    'Else
                    '    If .Risiko > 1 Then
                    '        SRweight = .StrategicFit / .Risiko
                    '    Else
                    '        SRweight = .StrategicFit
                    '    End If
                    'End If

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
    Public ReadOnly Property getTotalCostValuesInMonth() As Double()
        Get
            Dim costValues() As Double
            Dim zeitraum As Integer
            Dim tempArray() As Double

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)

            Dim anzCosts As Integer = CostDefinitions.Count

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

                        tempArray = hproj.getKostenBedarfNew(CostID)

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
            Dim roleName As String
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
                roleName = tmpRole.name

                If istSammelRolle Then
                    ' nur Platzhalter Rollenbedarfe berücksichtigen 
                    roleValues = Me.getRoleValuesInMonth(roleID:=roleName, _
                                                         considerAllSubRoles:=True, _
                                                         type:=PTcbr.all, _
                                                         excludedNames:=Nothing)
                Else
                    roleValues = Me.getRoleValuesInMonth(roleName)

                End If

                myCollection.Add(roleName, roleName)
                kapaValues = Me.getRoleKapasInMonth(myCollection, False)
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
                    ' nur Platzhalter Rollenbedarfe berücksichtigen 
                    roleValues = Me.getRoleValuesInMonth(roleName)
                    ReDim kapaValues(zeitraum)

                Else

                    myCollection.Add(roleName, roleName)
                    roleValues = Me.getRoleValuesInMonth(roleName)
                    kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                    myCollection.Clear()
                End If


                Select Case typus

                    Case 0
                        ' Auslastung

                        For ix = 0 To zeitraum
                            If roleValues(ix) > kapaValues(ix) And Not istSammelRolle Then
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
                            If roleValues(ix) > kapaValues(ix) And Not istSammelRolle Then
                                ' es gibt Überauslastung  
                                tmpValues(ix) = tmpValues(ix) + roleValues(ix) - kapaValues(ix)
                            Else
                                ' es gibt keine Überauslastung 

                            End If
                        Next ix

                    Case 2
                        ' Unterauslastung
                        For ix = 0 To zeitraum
                            If roleValues(ix) < kapaValues(ix) And Not istSammelRolle Then
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
    Public ReadOnly Property getAuslastungsArray(ByVal von As Integer, ByVal bis As Integer, _
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
                kapaValues = Me.getRoleKapasInMonth(myCollection, awinSettings.meAuslastungIsInclExt)
                myCollection.Clear()

                If istSammelRolle Then
                    ' alle Bedarfe berücksichtigen
                    roleValues = Me.getRoleValuesInMonth(roleID:=roleName, considerAllSubRoles:=True, _
                                                         type:=PTcbr.all)
                Else
                    roleValues = Me.getRoleValuesInMonth(roleName)
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
    ''' wird für MassenEdit benötigt
    ''' </summary>
    ''' <param name="roleID"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="percentValues"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAuslastungsArrayOfRole(ByVal roleID As Integer, ByVal von As Integer, ByVal bis As Integer, _
                                                       ByVal percentValues As Boolean) As Double()
        Get
            Dim tmpArray() As Double
            ReDim tmpArray(bis - von + 1)

            Try
                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(roleID)

                Dim roleUID As Integer = tmpRole.UID
                Dim roleName As String = tmpRole.name


                Dim roleValues() As Double
                Dim kapaValues() As Double
                Dim myCollection As New Collection
                Dim ix As Integer
                Dim zeitraum As Integer = bis - von

                Dim istSammelRolle As Boolean = tmpRole.isCombinedRole

                ReDim roleValues(zeitraum)
                ReDim kapaValues(zeitraum)

                myCollection.Add(roleName, roleName)
                kapaValues = Me.getRoleKapasInMonth(myCollection, awinSettings.meAuslastungIsInclExt)
                myCollection.Clear()

                If istSammelRolle Then
                    ' alle Bedarfe berücksichtigen
                    roleValues = Me.getRoleValuesInMonth(roleID:=roleName, considerAllSubRoles:=True, _
                                                         type:=PTcbr.all)
                Else
                    roleValues = Me.getRoleValuesInMonth(roleName)
                End If

                ' jetzt wird der Array aufgebaut

                If Not percentValues Then

                    tmpArray(0) = kapaValues.Sum - roleValues.Sum
                    For ix = 1 To bis - von + 1
                        tmpArray(ix) = kapaValues(ix - 1) - roleValues(ix - 1)
                    Next

                Else
                    If kapaValues.Sum > 0 Then
                        tmpArray(0) = roleValues.Sum / kapaValues.Sum
                    Else
                        tmpArray(0) = 999 ' Kennzeichen für unendlich 
                    End If

                    For ix = 1 To bis - von + 1
                        If kapaValues(ix - 1) > 0 Then
                            tmpArray(ix) = roleValues(ix - 1) / kapaValues(ix - 1)
                        Else
                            tmpArray(ix) = 999 ' Kennzeichen für unendlich ...
                        End If

                    Next
                End If

            Catch ex As Exception
                ReDim tmpArray(bis - von + 1)
            End Try

            getAuslastungsArrayOfRole = tmpArray

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
                    roleValues = Me.getRoleValuesInMonth(roleName)


                    If istSammelRolle Then

                        ReDim kapaValues(zeitraum)

                        myCollection.Add(roleName, roleName)
                        alleKapaValues = Me.getRoleKapasInMonth(myCollection, False)
                        alleRoleValues = Me.getRoleValuesInMonth(roleID:=roleName, _
                                                                 considerAllSubRoles:=True, _
                                                                 type:=PTcbr.all, _
                                                                 excludedNames:=Nothing)
                        myCollection.Clear()

                    Else
                        myCollection.Add(roleName, roleName)
                        kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                        myCollection.Clear()
                    End If

                    For ix = 0 To zeitraum

                        If istSammelRolle Then

                            If alleRoleValues(ix) > alleKapaValues(ix) Then
                                ' der Anteil FIG22 intern ist beschränkt auf alleKapas - SummealleSubroles ohne Sammelrolle 
                                ' 
                                alleSubRoleValues = Me.getRoleValuesInMonth(roleID:=roleName, _
                                                                 considerAllSubRoles:=True, _
                                                                 type:=PTcbr.realRoles, _
                                                                 excludedNames:=Nothing)
                                Dim diff As Double = alleKapaValues(ix) - alleSubRoleValues(ix)

                                If diff > 0 Then
                                    ' nur dann gibt es noch einen internen Anteil für den Platzhalter
                                    If roleValues(ix) <= diff Then
                                        ' der interne Teil ist maximal Diff oder eben roleValues ...
                                        diff = roleValues(ix)
                                    End If
                                End If

                                costValues(ix) = costValues(ix) + _
                                                 diff * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000

                            Else
                                ' die internen Ressourcen reichen aus
                                costValues(ix) = costValues(ix) + _
                                                 roleValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                            End If

                        Else

                            If roleValues(ix) > kapaValues(ix) Then
                                ' es werden die maximale Anzahl Leute dieser Rolle berücksichtigt 
                                costValues(ix) = costValues(ix) + _
                                                 kapaValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                            Else
                                ' die internen Ressourcen reichen aus
                                costValues(ix) = costValues(ix) + _
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


                    roleValues = Me.getRoleValuesInMonth(roleName)


                    ' mit welchem Wert wird gerechnet 

                    With tmpRole
                        If isDeltaCalculation Then
                            calculationValue = .tagessatzExtern - .tagessatzIntern
                        Else
                            calculationValue = .tagessatzExtern
                        End If
                    End With

                    If istSammelRolle Then

                        ReDim kapaValues(zeitraum)

                        myCollection.Add(roleName, roleName)
                        alleKapaValues = Me.getRoleKapasInMonth(myCollection, False)
                        alleRoleValues = Me.getRoleValuesInMonth(roleID:=roleName, _
                                                                    considerAllSubRoles:=True, _
                                                                    type:=PTcbr.all, _
                                                                    excludedNames:=Nothing)
                        myCollection.Clear()

                    Else
                        myCollection.Add(roleName, roleName)
                        kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                        myCollection.Clear()
                    End If


                    For ix = 0 To zeitraum

                        If istSammelRolle Then
                            Dim diff As Double
                            If alleRoleValues(ix) > alleKapaValues(ix) Then

                                alleSubRoleValues = Me.getRoleValuesInMonth(roleID:=roleName, _
                                                                 considerAllSubRoles:=True, _
                                                                 type:=PTcbr.realRoles, _
                                                                 excludedNames:=Nothing)

                                If alleSubRoleValues(ix) >= alleKapaValues(ix) Then
                                    ' der gesamte Platzhalter Anteil ist Overtime
                                    diff = roleValues(ix)
                                Else
                                    diff = alleRoleValues(ix) - alleKapaValues(ix)
                                End If

                                costValues(ix) = costValues(ix) + _
                                                 diff * calculationValue * faktor / 1000


                            Else
                                ' die internen Ressourcen reichen aus  

                            End If

                        Else

                            If roleValues(ix) > kapaValues(ix) Then
                                ' Overtime Kosten fallen an
                                costValues(ix) = costValues(ix) + _
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

    ''' <summary>
    ''' gibt die Personalkosten im Zeitraum zurück, dabei werden die Überstundensätze entsprechend des optionalen Parameters berücksichtigt 
    ''' 
    ''' der optionale Parameter bestimmt, ob die Überlast-Situationen berücksichtigt werden sollen oder nicht ... 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostGpValuesInMonth(Optional ByVal includesOverloadCost As Boolean = False) As Double()

        Get
            Dim costValues() As Double
            Dim roleValues() As Double
            Dim kapaValues() As Double
            Dim alleRoleValues() As Double
            Dim alleKapaValues() As Double

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
            ReDim alleRoleValues(zeitraum)
            ReDim alleKapaValues(zeitraum)


            For i = 1 To RoleDefinitions.Count

                Dim istSammelRolle As Boolean = RoleDefinitions.getRoledef(i).isCombinedRole
                roleName = RoleDefinitions.getRoledef(i).name
                roleValues = Me.getRoleValuesInMonth(roleName)

                If istSammelRolle Then

                    ReDim kapaValues(zeitraum)

                    If includesOverloadCost Then
                        myCollection.Add(roleName, roleName)
                        alleKapaValues = Me.getRoleKapasInMonth(myCollection, False)
                        alleRoleValues = Me.getRoleValuesInMonth(roleID:=roleName, _
                                                                 considerAllSubRoles:=True, _
                                                                 type:=PTcbr.all,
                                                                 excludedNames:=Nothing)
                        myCollection.Clear()
                    End If


                Else
                    myCollection.Add(roleName, roleName)
                    kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                    myCollection.Clear()
                End If




                For ix = 0 To zeitraum


                    If istSammelRolle Then

                        If alleRoleValues(ix) > alleKapaValues(ix) And includesOverloadCost Then
                            ' Overtime Kosten fallen an 
                            ' bei der Sammelrolle ist der Beitrag der Überlast auf die Höhe des Platzhalter-Wertes beschränkt 
                            Dim diff As Double = alleRoleValues(ix) - alleKapaValues(ix)
                            If diff > roleValues(ix) Then
                                diff = roleValues(ix)
                            End If
                            costValues(ix) = costValues(ix) + _
                                             alleKapaValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                            costValues(ix) = costValues(ix) + _
                                             diff * RoleDefinitions.getRoledef(roleName).tagessatzExtern * faktor / 1000
                        Else
                            ' die internen Ressourcen reichen aus oder die Kosten durch Überlast sollen nicht berücksichtigt werden 
                            costValues(ix) = costValues(ix) + _
                                             roleValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000

                        End If

                    Else

                        If roleValues(ix) > kapaValues(ix) And includesOverloadCost Then
                            ' externe Ressourcen müssen hinzugezogen werden
                            costValues(ix) = costValues(ix) + _
                                             kapaValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                            costValues(ix) = costValues(ix) + _
                                             (roleValues(ix) - kapaValues(ix)) * RoleDefinitions.getRoledef(roleName).tagessatzExtern * faktor / 1000
                        Else
                            ' die internen Ressourcen reichen aus oder die Kosten durch Überlast sollen nicht berücksichtigt werden 
                            costValues(ix) = costValues(ix) + _
                                             roleValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000

                        End If

                    End If

                Next ix

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
            Dim ix As Integer
            Dim zeitraum As Integer
            Dim tagessatzExtern As Double, tagessatzIntern As Double, diff As Double
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
                    tagessatzExtern = tmpRole.tagessatzExtern
                    tagessatzIntern = tmpRole.tagessatzIntern

                    If tagessatzExtern <> tagessatzIntern Then
                        diff = tagessatzExtern - tagessatzIntern
                        myCollection.Add(roleName, roleName)

                        If tmpRole.isCombinedRole Then
                            roleValues = Me.getRoleValuesInMonth(roleID:=roleName, _
                                                                 considerAllSubRoles:=True, _
                                                                 type:=PTcbr.all, _
                                                                 excludedNames:=roleCollection)

                        Else
                            roleValues = Me.getRoleValuesInMonth(roleID:=roleName)
                        End If

                        kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                        myCollection.Clear()

                        For ix = 0 To zeitraum
                            If roleValues(ix) > kapaValues(ix) Then
                                ' Kosten der externen Ressourcen
                                costValue = costValue + _
                                                 (roleValues(ix) - kapaValues(ix)) * diff * faktor / 1000
                            ElseIf roleValues(ix) < kapaValues(ix) Then
                                ' Kosten der internen Ressourcen, die nicht in Projekten arbeiten  
                                costValue = costValue + _
                                                 (kapaValues(ix) - roleValues(ix)) * tagessatzIntern * faktor / 1000

                            End If
                        Next ix
                    End If
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
    ''' gibt die Personalkosten zurück, die durch die internen Rollen entstehen, die in keinen Projekten gebunden sind 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostoValuesInMonth() As Double()

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

                If Not IsNothing(tmpRole) Then
                    Dim istSammelRolle As Boolean = tmpRole.isCombinedRole

                    roleName = tmpRole.name
                    roleValues = Me.getRoleValuesInMonth(roleName)

                    If istSammelRolle Then
                        ReDim kapaValues(zeitraum)
                    Else
                        myCollection.Add(roleName, roleName)
                        kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                        myCollection.Clear()
                    End If



                    For ix = 0 To zeitraum

                        If istSammelRolle Then
                            ' das sind ja Platzhalter Bedarfe, die irgendwann von irgendeiner realen Ressource wahrgenommen werden
                            ' deswegen müssen die von den anderen abgezogen werden ... weil sonst zuviel als "ohne Arbeit" ausgewiesen wird 
                            costValues(ix) = costValues(ix) - roleValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000

                        Else
                            If roleValues(ix) < kapaValues(ix) Then
                                ' interne Ressourcen kosten , können aber nicht verrechnet werden 
                                costValues(ix) = costValues(ix) + _
                                                 (kapaValues(ix) - roleValues(ix)) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                            Else
                                ' keine Opportunity Kosten 
                            End If
                        End If


                    Next ix
                End If


            Next i

            ' falls die costValues negativ sind .. korrigieren auf Null 
            For ix = 0 To zeitraum
                If costValues(ix) < 0 Then
                    costValues(ix) = 0
                End If
            Next


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
