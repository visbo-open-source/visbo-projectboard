
Imports xlNS = Microsoft.Office.Interop.Excel

Public Class clsProjekte

    Private AllProjects As SortedList(Of String, clsProjekt)
    Private AllShapes As SortedList(Of String, String)
    Private AllCoord As SortedList(Of String, Double())

    ''' <summary>
    ''' trägt ein Projekt mit dem Schlüssel Projekt-NAme in die Liste ein 
    ''' trägt die Shape ID (shpUID) in die Shape Liste ein 
    ''' wenn der Projekt-Name bereits existiert, wird eine Exception geworfen 
    ''' </summary>
    ''' <param name="project"></param>
    ''' <remarks></remarks>
    Public Sub Add(project As clsProjekt)

        Try
            Dim pname As String = project.name
            Dim shpUID As String = project.shpUID

            AllProjects.Add(pname, project)

            If shpUID <> "" Then
                AllShapes.Add(shpUID, pname)
            End If

            ' mit diesem Vorgang wird die Konstellation geändert , deshalb muss die currentConstellation zurückgesetzt werden 
            currentConstellation = ""

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


        If AllProjects.ContainsKey(pname) Then
            Try
                If AllShapes.ContainsValue(pname) Then
                    Dim ix As Integer
                    ix = AllShapes.IndexOfValue(pname)
                    AllShapes.RemoveAt(ix)
                End If
                AllShapes.Add(shpUID, pname)

            Catch ex As Exception
                Throw New ArgumentException(ex.Message)
            End Try
        Else
            Throw New ArgumentException("Shape kann nicht einem nicht-existierenden Projekt hinzugefügt werden - ")
        End If



    End Sub


    ''' <summary>
    ''' nimmt das Projekt mit dem übergebenen Namen aus der Liste heraus  
    ''' wirft eine Exception, wenn Projekt nicht ind er Liste oder ShpUID nicht ind er zugehörigen Shape-Liste
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <remarks></remarks>
    Public Sub Remove(projectname As String)

        Try
            Dim SID As String = AllProjects.Item(projectname).shpUID
            AllProjects.Remove(projectname)
            If SID <> "" Then
                AllShapes.Remove(SID)
            End If

            ' mit diesem Vorgang wird die Konstellation geändert , deshalb muss das zurückgesetzt werden 
            currentConstellation = ""

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
            Dim pname As String = AllShapes.Item(SID)
            AllProjects.Remove(pname)
            AllShapes.Remove(SID)

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try



    End Sub

    ''' <summary>
    ''' setzt die Liste der Projekte und die Liste der Shapes zurück 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()

        AllProjects.Clear()
        AllShapes.Clear()

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
            If AllProjects.ContainsKey(key) Then
                contains = True
            Else
                contains = False
            End If
        End Get
    End Property


    ''' <summary>
    ''' gibt eine sortierte Liste der vorkommenden Phasen Namen in der Menge von Projekten zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseNames() As Collection

        Get

            Dim tmpListe As New Collection
            Dim cphase As clsPhase

            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects

                Try
                    ' beginnt bei 2, weil die 1.Phase immer die mit der Projektlänge identische Phase ist ...
                    For p = 2 To kvp.Value.CountPhases

                        cphase = kvp.Value.getPhase(p)

                        If tmpListe.Contains(cphase.name) Then
                            ' nichts tun 
                        Else
                            tmpListe.Add(cphase.name, cphase.name)
                        End If


                    Next
                Catch ex As Exception

                End Try


            Next

            getPhaseNames = tmpListe

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
            Dim cphase As clsPhase

            Dim msName As String

            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects

                Try
                    For p = 1 To kvp.Value.CountPhases

                        cphase = kvp.Value.getPhase(p)
                        For r = 1 To cphase.CountResults

                            msName = cphase.getResult(r).name
                            If tmpListe.Contains(msName) Then
                            Else
                                tmpListe.Add(msName, msName)
                            End If

                        Next

                    Next
                Catch ex As Exception

                End Try


            Next

            getMilestoneNames = tmpListe

        End Get
    End Property



    ''' <summary>
    ''' gibt die nach Namen sortierte Liste von Projekten zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Liste() As SortedList(Of String, clsProjekt)
        Get
            Liste = AllProjects
        End Get
        Set(value As SortedList(Of String, clsProjekt))
            AllProjects = value
        End Set
    End Property

    ''' <summary>
    ''' gibt die Anzahl der Projekte in der Liste an 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count() As Integer

        Get
            Count = AllProjects.Count
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
            Try
                getProject = AllProjects.ElementAt(index - 1).Value
            Catch ex As Exception
                Throw New ArgumentException("Index nicht vorhanden:" & index.ToString)
            End Try
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

            With CType(appInstance.Worksheets(arrWsNames(3)), xlNS.Worksheet)
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
    ''' gibt das vollständige Projekt aus der Liste zurück, das den angegebenen Namen hat 
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(projectname As String) As clsProjekt

        Get
            Try
                getProject = AllProjects.Item(projectname)
            Catch ex As Exception
                Throw New ArgumentException("projectname nicht vorhanden")
            End Try

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

            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects
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

                pname = AllShapes.Item(shpID)
                getProjectS = AllProjects.Item(pname)

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
            shpListe = AllShapes
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



            For Each kvp As KeyValuePair(Of String, clsProjekt) In Me.AllProjects

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
            Dim cmileStone As clsMeilenstein
            Dim projektstart As Integer
            Dim found As Boolean
            Dim key As Double
            ' selection type wird aktuell noch ignoriert .... 

            suchTyp = 0

            For Each kvp As KeyValuePair(Of String, clsProjekt) In Me.AllProjects

                found = False
                With kvp.Value

                    projektstart = .Start + .StartOffset

                    If (projektstart > bis) Or (projektstart + .anzahlRasterElemente - 1 < von) Then
                        ' dann liegt das Projekt ausserhalb des Zeitraums und muss überhaupt nicht berücksichtig werden 

                    Else

                        For Each phaseName As String In selItems1

                            cphase = kvp.Value.getPhase(phaseName)
                            If Not IsNothing(cphase) Then
                                If (projektstart + cphase.relStart - 1 > bis) Or (projektstart + cphase.relEnde - 1 < von) Then
                                    ' dann liegt die Phase ausserhalb des betrachteten Zeitraums und muss nicht berücksichtigt werden 
                                Else
                                    found = True
                                    Exit For
                                End If
                            End If

                            For Each milestoneName As String In selItems2
                                cmileStone = cphase.getResult(milestoneName)
                                If Not IsNothing(cmileStone) Then
                                    Dim msColumn As Integer = getColumnOfDate(cmileStone.getDate)
                                    If msColumn > bis Or msColumn < von Then
                                    Else
                                        found = True
                                        Exit For
                                    End If
                                End If

                            Next

                            If found Then
                                Exit For
                            End If

                        Next

                        If Not found And selItems1.Count = 0 Then
                            For Each milestoneName As String In selItems2
                                For Each hphase As clsPhase In kvp.Value.AllPhases
                                    cmileStone = hphase.getResult(milestoneName)
                                    If Not IsNothing(cmileStone) Then
                                        Dim msColumn As Integer = getColumnOfDate(cmileStone.getDate)
                                        If msColumn > bis Or msColumn < von Then
                                        Else
                                            found = True
                                            Exit For
                                        End If
                                    End If
                                Next

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
    ''' bestimmt für den angegebenen Zeitraum die Projekte, die eine der angegeben Phasen oder Meilensteine im Zeitraum enthalten. 
    ''' bestimmt darüber hinaus das minimale bzw. maximale Datum , das die Phasen der Projekte aufspannen , die den Zeitraum "berühren"  
    ''' </summary>
    ''' <param name="selectedPhases">die Phasen, nach denen gesúcht wird </param>
    ''' <param name="selectedMilestones">die Meilensteine, nach denen gesucht wird</param>
    ''' <param name="von">linker Rand des Zeitraums</param>
    ''' <param name="bis">rechter Rand des zeitraums</param>
    ''' <param name="projektListe">Ergebnis enthält alle Projekt-Namen die eine der Phasen oder einen der Meilensteine im angegebenen Zeitraum enthalten </param>
    ''' <param name="minDate">das kleinste auftretende Start-Datum einer Phase</param>
    ''' <param name="maxDate">das größte auftretende Ende-Datum einer Phase </param>
    ''' <remarks></remarks>
    Public Sub bestimmeProjekteAndMinMaxDates(ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                              ByVal von As Integer, ByVal bis As Integer, _
                                                  ByRef projektListe As SortedList(Of Double, String), ByRef minDate As Date, ByRef maxDate As Date)

        Dim tmpMinimum As Date = StartofCalendar.AddMonths(von - 1)
        Dim tmpMaximum As Date = StartofCalendar.AddMonths(bis).AddDays(-1)
        Dim tmpDate As Date


        Dim hproj As clsProjekt
        Dim cphase As clsPhase
        Dim projektstart As Integer
        Dim found As Boolean
        Dim key As Double
        ' selection type wird aktuell noch ignoriert .... 

        ' in der ersten Welle werden die Projektnamen aufgesammelt, die eine der Phasen oder Meilensteine enthalten  
        For Each kvp As KeyValuePair(Of String, clsProjekt) In Me.AllProjects

            found = False

            With kvp.Value

                projektstart = .Start + .StartOffset


                If (projektstart > bis) Or (projektstart + .anzahlRasterElemente - 1 < von) Then
                    ' dann liegt das Projekt ausserhalb des Zeitraums und muss überhaupt nicht berücksichtig werden 

                Else

                    Dim ix As Integer = 1
                    Dim phaseName As String
                    While ix <= selectedPhases.Count And Not found

                        phaseName = CStr(selectedPhases.Item(ix))
                        cphase = kvp.Value.getPhase(phaseName)

                        If Not IsNothing(cphase) Then
                            If phaseWithinTimeFrame(projektstart, cphase.relStart, cphase.relEnde, von, bis) Then
                                found = True
                            Else
                                ix = ix + 1
                            End If
                        Else
                            ix = ix + 1
                        End If
                        


                    End While

                    ' das muss nur gemacht werden, wenn found <> true
                    If Not found Then
                        ix = 1
                        Dim milestoneName As String

                        While ix <= selectedMilestones.Count And Not found

                            milestoneName = CStr(selectedMilestones.Item(ix))
                            tmpDate = kvp.Value.getMilestoneDate(milestoneName)

                            If milestoneWithinTimeFrame(tmpDate, von, bis) Then
                                found = True
                            Else
                                ix = ix + 1
                            End If


                        End While

                    End If


                End If


            End With

            If found Then
                key = kvp.Value.tfZeile + kvp.Value.anzahlRasterElemente / 10000
                projektListe.Add(key, kvp.Value.name)
            End If

        Next

        ' jetzt muss die zweite Welle nachkommen .. bestimmen , welches die erweiterten Min / Max Werte sind, falls fullyContained bzw. showAllIfOne 
        ' hier jetzt für alle Projekte in projektliste für jedes Element aus selectedphases und selectedmilestones das Minimum / Maximum bestimmen


        For Each kvp As KeyValuePair(Of Double, String) In projektListe

            hproj = Me.getProject(kvp.Value)
            projektstart = hproj.Start + hproj.StartOffset

            ' Phasen checken 
            For Each phaseName As String In selectedPhases

                cphase = hproj.getPhase(phaseName)
                If Not IsNothing(cphase) Then
                    If awinSettings.mppShowAllIfOne Then

                        If DateDiff(DateInterval.Day, cphase.getStartDate, tmpMinimum) > 0 Then
                            tmpMinimum = cphase.getStartDate
                        End If

                        If DateDiff(DateInterval.Day, cphase.getEndDate, tmpMaximum) < 0 Then
                            tmpMaximum = cphase.getEndDate
                        End If


                    Else
                        If phaseWithinTimeFrame(projektstart, cphase.relStart, cphase.relEnde, von, bis) Then

                            If DateDiff(DateInterval.Day, cphase.getStartDate, tmpMinimum) > 0 Then
                                tmpMinimum = cphase.getStartDate
                            End If

                            If DateDiff(DateInterval.Day, cphase.getEndDate, tmpMaximum) < 0 Then
                                tmpMaximum = cphase.getEndDate
                            End If

                        End If
                    End If
                End If


            Next

            ' Meilensteine 
            ' das muss nur gemacht werden, wenn showAllIfOne=true 
            If awinSettings.mppShowAllIfOne Then
                For Each msName As String In selectedMilestones

                    tmpDate = hproj.getMilestoneDate(msName)

                    If DateDiff(DateInterval.Day, StartofCalendar, tmpDate) >= 0 Then
                        If DateDiff(DateInterval.Day, tmpDate, tmpMinimum) > 0 Then
                            tmpMinimum = tmpDate

                        End If

                        If DateDiff(DateInterval.Day, tmpDate, tmpMaximum) < 0 Then
                            tmpMaximum = tmpDate
                        End If
                    End If

                Next
            End If
            

        Next

        If Not awinSettings.mppfullyContained Then
            minDate = StartofCalendar.AddMonths(von - 1)
            maxDate = StartofCalendar.AddMonths(bis).AddDays(-1)
        Else
            minDate = tmpMinimum
            maxDate = tmpMaximum
        End If
        

    End Sub


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
    Public ReadOnly Property getCountMilestonesInMonth(milestoneName As String) As Double(,)
        Get

            Dim milestoneValues(,) As Double
            Dim zeitraum As Integer
            Dim anzProjekte As Integer

            Dim cphase As clsPhase
            Dim cresult As clsMeilenstein
            Dim hproj As clsProjekt
            Dim ix As Integer
            Dim idFarbe As Integer

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim milestoneValues(3, zeitraum)

            anzProjekte = AllProjects.Count

            ' Schleife über alle Projekte 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects

                hproj = kvp.Value

                ' alle Phasen durchgehen und nach dem Meilenstein-Namen suchen 
                Dim p As Integer
                For p = 1 To hproj.CountPhases


                    cphase = hproj.getPhase(p)
                    cresult = cphase.getResult(milestoneName)

                    If IsNothing(cresult) Then
                    Else

                        ' bestimme den monatsbezogenen Index im Array 
                        ix = getColumnOfDate(cresult.getDate) - showRangeLeft

                        If ix >= 0 And ix <= zeitraum Then

                            If cresult.bewertungsCount > 0 Then
                                idFarbe = cresult.getBewertung(1).colorIndex
                            Else
                                idFarbe = 0
                            End If

                            milestoneValues(idFarbe, ix) = milestoneValues(idFarbe, ix) + 1

                        End If


                    End If



                Next


            Next kvp


            getCountMilestonesInMonth = milestoneValues


        End Get
    End Property

    ''' <summary>
    ''' gibt einen Array zurück, der angibt, wie oft die angegebene Phase vorkommt
    ''' showrangeleft und showrangeright spannen den Betrachtungszeitraum auf 
    ''' </summary>
    ''' <param name="phaseName">Name der Phase</param>
    ''' <value></value>
    ''' <returns>gibt einen Array der Länge (showrangeright-showrangeleft+1) zurück </returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCountPhasesInMonth(phaseName As String) As Double()

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

            anzProjekte = AllProjects.Count

            ' anzPhasen = AllPhases.Count

            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects

                hproj = kvp.Value



                Try
                    hphase = hproj.getPhase(phaseName)
                Catch ex As Exception
                    hphase = Nothing
                End Try


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
    Public ReadOnly Property getRoleValuesInMonth(roleID As Object) As Double()

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

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            lookforIndex = IsNumeric(roleID)
            zeitraum = showRangeRight - showRangeLeft
            ReDim roleValues(zeitraum)

            anzProjekte = AllProjects.Count

            ' anzPhasen = AllPhases.Count

            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects

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

                        tempArray = hproj.getRessourcenBedarf(roleID)

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
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseSchwellWerteInMonth(myCollection As Collection) As Double()
        Get

            Dim schwellWerte() As Double

            Dim hkapa As Double
            Dim rname As String
            Dim zeitraum As Integer
            Dim r As Integer, m As Integer


            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum
            zeitraum = showRangeRight - showRangeLeft
            ReDim schwellWerte(zeitraum)

            For r = 1 To myCollection.Count

                rname = CStr(myCollection.Item(r))
                hkapa = PhaseDefinitions.getPhaseDef(rname).schwellWert

                For m = 0 To zeitraum
                    ' Änderung 31.5 Holen der Schwellwerte einer Phase 
                    schwellWerte(m) = schwellWerte(m) + hkapa
                Next m


            Next r

            getPhaseSchwellWerteInMonth = schwellWerte

        End Get
    End Property

    '
    '
    '
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


            zeitraum = showRangeRight - showRangeLeft
            ReDim kapaValues(zeitraum)
            ReDim tmpValues(zeitraum)


            For r = 1 To myCollection.Count
                rname = CStr(myCollection.Item(r))
                hkapa = RoleDefinitions.getRoledef(rname).Startkapa

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

            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects
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
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
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


            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects

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


            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects
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

            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects
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


                    For i = 0 To anzLoops - 1
                        ' Änderung 2.5.14 : es sollen nur die in der Zukunft liegenden Monate mit einem Risiko Aufschlag bedacht werden 
                        If ix + i >= heuteIndex Then
                            riskValues(ixZeitraum + i) = riskValues(ixZeitraum + i) + tempArray(ix + i) * riskweightedMarge
                        End If

                    Next i


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
            Dim Dauer As Integer
            Dim zeitraum As Integer
            Dim i As Integer
            Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
            Dim hproj As clsProjekt
            Dim tempArray() As Double
            Dim prAnfang As Integer, prEnde As Integer


            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum



            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)


            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects
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

                    tempArray = hproj.getGesamtKostenBedarf

                    For i = 0 To anzLoops - 1
                        costValues(ixZeitraum + i) = costValues(ixZeitraum + i) + tempArray(ix + i)
                    Next i


                End If
                'hproj = Nothing
            Next kvp

            getTotalCostValuesInMonth = costValues

        End Get
    End Property

    '
    '
    '
    ''' <summary>
    ''' gibt die Gesamtkosten , Personalkosten und alle sonstigen Kosten im betrachteten Zeitraum zurück 
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
            Dim persCost As Boolean

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            lookforIndex = IsNumeric(CostID)
            persCost = False

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


            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects
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

                    If isPersCost Then
                        tempArray = hproj.getAllPersonalKosten
                    Else
                        tempArray = hproj.getKostenBedarf(CostID)
                    End If

                    For i = 0 To anzLoops - 1
                        costValues(ixZeitraum + i) = costValues(ixZeitraum + i) + tempArray(ix + i)
                    Next i


                End If
                'hproj = Nothing
            Next kvp

            getCostValuesInMonth = costValues

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





            For i = 1 To RoleDefinitions.Count
                roleName = RoleDefinitions.getRoledef(i).name
                myCollection.Add(roleName)
                roleValues = Me.getRoleValuesInMonth(roleName)
                kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                myCollection.Clear()

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
    Public ReadOnly Property getAuslastungsValues(roleName As String, typus As Integer) As Double()

        Get
            Dim roleValues() As Double
            Dim kapaValues() As Double
            Dim tmpValues() As Double
            Dim myCollection As New Collection
            Dim ix As Integer
            Dim zeitraum As Integer


            zeitraum = showRangeRight - showRangeLeft
            ReDim roleValues(zeitraum)
            ReDim kapaValues(zeitraum)
            ReDim tmpValues(zeitraum)

            myCollection.Add(roleName)
            roleValues = Me.getRoleValuesInMonth(roleName)
            kapaValues = Me.getRoleKapasInMonth(myCollection, False)
            myCollection.Clear()

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

            getAuslastungsValues = tmpValues


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
            Dim roleName As String
            Dim myCollection As New Collection
            Dim i As Integer, ix As Integer
            Dim zeitraum As Integer
            Dim faktor As Double = nrOfDaysMonth

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Then
                faktor = 1
            Else
                faktor = 1
            End If


            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)
            ReDim roleValues(zeitraum)
            ReDim kapaValues(zeitraum)

            For i = 1 To RoleDefinitions.Count
                roleName = RoleDefinitions.getRoledef(i).name
                myCollection.Add(roleName)
                roleValues = Me.getRoleValuesInMonth(roleName)
                kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                myCollection.Clear()

                For ix = 0 To zeitraum
                    If roleValues(ix) > kapaValues(ix) Then
                        ' es werden die maximale Anzahl Leute dieser Rolle berücksichtigt 
                        costValues(ix) = costValues(ix) + _
                                         kapaValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                    Else
                        ' die internen Ressourcen reichen aus
                        costValues(ix) = costValues(ix) + _
                                         roleValues(ix) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                    End If
                Next ix

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
    Public ReadOnly Property getCosteValuesInMonth() As Double()

        Get
            Dim costValues() As Double
            Dim roleValues() As Double
            Dim kapaValues() As Double
            Dim roleName As String
            Dim myCollection As New Collection
            Dim i As Integer, ix As Integer
            Dim zeitraum As Integer
            Dim faktor As Double = nrOfDaysMonth

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Then
                faktor = 1
            Else
                faktor = 1
            End If

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)
            ReDim roleValues(zeitraum)
            ReDim kapaValues(zeitraum)

            For i = 1 To RoleDefinitions.Count
                roleName = RoleDefinitions.getRoledef(i).name
                myCollection.Add(roleName)
                roleValues = Me.getRoleValuesInMonth(roleName)
                kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                myCollection.Clear()

                For ix = 0 To zeitraum
                    If roleValues(ix) > kapaValues(ix) Then
                        ' externe Ressourcen müssen hinzugezogen werden
                        costValues(ix) = costValues(ix) + _
                                         (roleValues(ix) - kapaValues(ix)) * RoleDefinitions.getRoledef(roleName).tagessatzExtern * faktor / 1000
                    Else
                        ' die internen Ressourcen reichen aus

                    End If
                Next ix

            Next i


            getCosteValuesInMonth = costValues

        End Get

    End Property

    ''' <summary>
    ''' gibt die Mehrkosten, die im Zeitraum durch den Einsatz von Externen verursacht werden , zurück 
    ''' der Wert repräsentiert dabei den Unterschied zu den Kosten, die durch den Einsatz von Internen anfallen würden
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getadditionalECostinMonth() As Double()

        Get
            Dim costValues() As Double
            Dim roleValues() As Double
            Dim kapaValues() As Double
            Dim roleName As String
            Dim myCollection As New Collection
            Dim i As Integer, ix As Integer
            Dim zeitraum As Integer
            Dim tagessatzDifferenz As Double
            Dim faktor As Double = nrOfDaysMonth

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Then
                faktor = 1
            Else
                faktor = 1
            End If

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)
            ReDim roleValues(zeitraum)
            ReDim kapaValues(zeitraum)

            For i = 1 To RoleDefinitions.Count
                roleName = RoleDefinitions.getRoledef(i).name
                myCollection.Add(roleName)
                roleValues = Me.getRoleValuesInMonth(roleName)
                kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                myCollection.Clear()

                With RoleDefinitions.getRoledef(roleName)
                    tagessatzDifferenz = .tagessatzExtern - .tagessatzIntern
                End With

                For ix = 0 To zeitraum
                    If roleValues(ix) > kapaValues(ix) Then
                        ' externe Ressourcen müssen hinzugezogen werden
                        costValues(ix) = costValues(ix) + _
                                         (roleValues(ix) - kapaValues(ix)) * tagessatzDifferenz * faktor / 1000
                    Else
                        ' die internen Ressourcen reichen aus

                    End If
                Next ix

            Next i


            getadditionalECostinMonth = costValues

        End Get

    End Property

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


            ' das sind die Zusatzkosten, die durch Externe (wg Überauslastung) verursacht werden
            additionalCostExt = System.Math.Round(ShowProjekte.getadditionalECostinMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10

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
    Public ReadOnly Property getbadCostOfRole(roleCollection As Collection) As Double
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
            Dim faktor As Double = nrOfDaysMonth

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Then
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

                tagessatzExtern = RoleDefinitions.getRoledef(roleName).tagessatzExtern
                tagessatzIntern = RoleDefinitions.getRoledef(roleName).tagessatzIntern

                If tagessatzExtern <> tagessatzIntern Then
                    diff = tagessatzExtern - tagessatzIntern
                    myCollection.Add(roleName)
                    roleValues = Me.getRoleValuesInMonth(roleName)
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
            Dim tmpSum As Double

            Dim zeitraum As Integer
            Dim rcName As String
            Dim i As Integer
            Dim anzElements As Integer

            anzElements = myCollection.Count
            zeitraum = showRangeRight - showRangeLeft


            tmpSum = 0.0

            For i = 1 To myCollection.Count
                rcName = CStr(myCollection.Item(i))

                ReDim tmpValues(zeitraum)

                If diagrammtyp = DiagrammTypen(0) Then
                    tmpValues = Me.getCountPhasesInMonth(rcName)
                ElseIf diagrammtyp = DiagrammTypen(1) Then
                    tmpValues = Me.getRoleValuesInMonth(rcName)
                ElseIf diagrammtyp = DiagrammTypen(2) Then
                    tmpValues = Me.getCostValuesInMonth(rcName)
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
                    tmpValues = Me.getCountPhasesInMonth(rcName)
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
            Dim faktor As Double = nrOfDaysMonth

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Then
                faktor = 1
            Else
                faktor = 1
            End If

            zeitraum = showRangeRight - showRangeLeft
            ReDim costValues(zeitraum)
            ReDim roleValues(zeitraum)
            ReDim kapaValues(zeitraum)

            For i = 1 To RoleDefinitions.Count
                roleName = RoleDefinitions.getRoledef(i).name
                myCollection.Add(roleName)
                roleValues = Me.getRoleValuesInMonth(roleName)
                kapaValues = Me.getRoleKapasInMonth(myCollection, False)
                myCollection.Clear()

                For ix = 0 To zeitraum
                    If roleValues(ix) < kapaValues(ix) Then
                        ' interne Ressourcen kosten , können aber nicht verrechnet werden 
                        costValues(ix) = costValues(ix) + _
                                         (kapaValues(ix) - roleValues(ix)) * RoleDefinitions.getRoledef(roleName).tagessatzIntern * faktor / 1000
                    Else
                        ' die internen Ressourcen reichen aus

                    End If
                Next ix

            Next i


            getCostoValuesInMonth = costValues


        End Get

    End Property

    '' ''' <summary>
    '' ''' Konstruktor: intilaisert die sortierte Liste der Projekte und Shapes   
    '' ''' </summary>
    '' ''' <remarks></remarks>
    Public Sub New()

        AllProjects = New SortedList(Of String, clsProjekt)
        AllShapes = New SortedList(Of String, String)

    End Sub

End Class
