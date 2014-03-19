Public Class clsProjekte

    Private AllProjects As SortedList(Of String, clsProjekt)
    Private AllShapes As SortedList(Of String, String)


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

    Public Sub RemoveS(SID As String)

        Try
            Dim pname As String = AllShapes.Item(SID)
            AllProjects.Remove(pname)
            AllShapes.Remove(SID)

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try



    End Sub

    Public Sub Clear()

        AllProjects.Clear()
        AllShapes.Clear()

    End Sub


    ''' <summary>
    ''' gibt eine sortierte Liste der vorkommenden Phasen Namen in der Menge von Projekten zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseNames() As SortedList(Of String, String)

        Get

            Dim tmpListe As New SortedList(Of String, String)
            Dim cphase As clsPhase
            
            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects

                Try
                    ' beginnt bei 2, weil die 1.Phase immer die mit der Projektlänge identische Phase ist ...
                    For p = 2 To kvp.Value.CountPhases

                        cphase = kvp.Value.getPhase(p)

                        If tmpListe.ContainsKey(cphase.name) Then
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
    Public ReadOnly Property getMilestoneNames() As SortedList(Of String, String)

        Get

            Dim tmpListe As New SortedList(Of String, String)
            Dim cphase As clsPhase

            Dim msName As String

            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects

                Try
                    For p = 1 To kvp.Value.CountPhases

                        cphase = kvp.Value.getPhase(p)
                        For r = 1 To cphase.CountResults

                            msName = cphase.getResult(r).name
                            If tmpListe.ContainsKey(msName) Then
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


    Public Property Liste() As SortedList(Of String, clsProjekt)
        Get
            Liste = AllProjects
        End Get
        Set(value As SortedList(Of String, clsProjekt))
            AllProjects = value
        End Set
    End Property

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


    Public ReadOnly Property getProject(projectname As String) As clsProjekt

        Get
            Try
                getProject = AllProjects.Item(projectname)
            Catch ex As Exception
                Throw New ArgumentException("projectname nicht vorhanden")
            End Try

        End Get

    End Property

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

            For Each kvp In Me.AllProjects
                With kvp.Value
                    If (.Start + .StartOffset > bis) Or (.Start + .StartOffset + .Dauer - 1 < von) Then
                    Else
                        tmpListe.Add(kvp.Key, kvp.Key)
                    End If
                End With
            Next

            withinTimeFrame = tmpListe

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
    Public ReadOnly Property getCountMilestonesInMonth(milestoneName As String) As Double(,)
        Get

            Dim milestoneValues(,) As Double
            Dim zeitraum As Integer
            Dim anzProjekte As Integer

            Dim cphase As clsPhase
            Dim cresult As clsResult
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
                        prEnde = .Start + .Dauer - 1 + .StartOffset
                    End With




                    If istBereichInTimezone(prAnfang, prEnde) Then
                        'projektstart = hproj.Start

                        With hphase
                            phAnfang = prAnfang + .relStart - 1
                            phEnde = prAnfang + .relEnde - 1
                        End With

                        Call awinIntersectZeitraum(phAnfang, phEnde, ixZeitraum, ix, anzLoops)

                        If anzLoops > 0 Then

                            ReDim tempArray(phEnde - phAnfang)
                            tempArray = hproj.getPhasenBedarf(phaseName)

                            For i = 0 To anzLoops - 1
                                ' das awinintersect ermittelt die Werte für Projekt-Anfang, Projekt-Ende 
                                ' in temparray stehen dagegen 
                                phasevalues(ixZeitraum + i) = phasevalues(ixZeitraum + i) + tempArray(ix + i)
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

                Dauer = hproj.Dauer

                ReDim tempArray(Dauer - 1)

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .Dauer - 1 + .StartOffset
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
    '
    '
    '
    ''' <summary>
    ''' gibt für die in myCollection übergebenen Rollen die Kapazitäten zurück 
    ''' </summary>
    ''' <param name="myCollection">enthält die Liste der zu betrachtenden Rollen-Namen</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleKapasInMonth(myCollection As Collection) As Double()

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
                rname = myCollection.Item(r)
                hkapa = RoleDefinitions.getRoledef(rname).Startkapa

                For i = showRangeLeft To showRangeRight
                    tmpValues(i - showRangeLeft) = RoleDefinitions.getRoledef(rname).kapazitaet(i)
                    If tmpValues(i - showRangeLeft) <= 0 Then
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

                Dauer = hproj.Dauer

                ReDim tempArray(Dauer - 1)

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .Dauer - 1 + .StartOffset
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

                Dauer = hproj.Dauer

                ReDim tempArray(Dauer - 1)

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .Dauer - 1 + .StartOffset
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

            ' showRangeLeft As Integer, showRangeRight sind die beiden Markierungen für den betrachteten Zeitraum

            zeitraum = showRangeRight - showRangeLeft
            ReDim riskValues(zeitraum)


            For Each kvp As KeyValuePair(Of String, clsProjekt) In AllProjects
                hproj = kvp.Value

                Dauer = hproj.Dauer

                ReDim tempArray(Dauer - 1)

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .Dauer - 1 + .StartOffset

                End With

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
                        riskValues(ixZeitraum + i) = riskValues(ixZeitraum + i) + tempArray(ix + i) * riskweightedMarge
                    Next i


                End If
                'hproj = Nothing
            Next kvp

            getWeightedRiskValuesInMonth = riskValues

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

                Dauer = hproj.Dauer

                ReDim tempArray(Dauer - 1)

                With hproj
                    prAnfang = .Start + .StartOffset
                    prEnde = .Start + .Dauer - 1 + .StartOffset
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
                kapaValues = Me.getRoleKapasInMonth(myCollection)
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
            kapaValues = Me.getRoleKapasInMonth(myCollection)
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
                kapaValues = Me.getRoleKapasInMonth(myCollection)
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
                kapaValues = Me.getRoleKapasInMonth(myCollection)
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
                kapaValues = Me.getRoleKapasInMonth(myCollection)
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
                roleName = roleCollection.Item(i)

                tagessatzExtern = RoleDefinitions.getRoledef(roleName).tagessatzExtern
                tagessatzIntern = RoleDefinitions.getRoledef(roleName).tagessatzIntern

                If tagessatzExtern <> tagessatzIntern Then
                    diff = tagessatzExtern - tagessatzIntern
                    myCollection.Add(roleName)
                    roleValues = Me.getRoleValuesInMonth(roleName)
                    kapaValues = Me.getRoleKapasInMonth(myCollection)
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

    Public ReadOnly Property getAverage(myCollection As Collection, isRole As Boolean) As Double
        Get
            Dim rcValues(,) As Double, tmpValues() As Double
            Dim tmpSum As Double
            Dim ix As Integer
            Dim zeitraum As Integer
            Dim rcName As String
            Dim i As Integer
            Dim anzElements As Integer

            anzElements = myCollection.Count
            zeitraum = showRangeRight - showRangeLeft
            ReDim rcValues(anzElements - 1, zeitraum)

            tmpSum = 0.0

            For i = 1 To myCollection.Count
                rcName = myCollection.Item(i)

                ReDim tmpValues(zeitraum)
                If isRole Then
                    tmpValues = Me.getRoleValuesInMonth(rcName)
                Else
                    tmpValues = Me.getCostValuesInMonth(rcName)
                End If


                For ix = 0 To zeitraum
                    tmpSum = tmpSum + tmpValues(ix)
                    rcValues(i - 1, ix) = tmpValues(ix)
                Next ix
            Next i

            getAverage = tmpSum / (zeitraum + 1)

        End Get
    End Property

    Public ReadOnly Property getDeviationfromAverage(myCollection As Collection, avgValue As Double, isRole As Boolean) As Double

        Get
            Dim rcValues(,) As Double, tmpValues() As Double
            Dim sumAboveAvg As Double, tmpSum As Double
            Dim ix As Integer
            Dim zeitraum As Integer
            Dim rcName As String
            Dim i As Integer
            Dim anzElements As Integer

            anzElements = myCollection.Count
            zeitraum = showRangeRight - showRangeLeft
            ReDim rcValues(anzElements - 1, zeitraum)
            tmpSum = 0.0

            For i = 1 To myCollection.Count
                rcName = myCollection.Item(i)

                ReDim tmpValues(zeitraum)
                If isRole Then
                    tmpValues = Me.getRoleValuesInMonth(rcName)
                Else
                    tmpValues = Me.getCostValuesInMonth(rcName)
                End If


                For ix = 0 To zeitraum
                    tmpSum = tmpSum + tmpValues(ix)
                    rcValues(i - 1, ix) = tmpValues(ix)
                Next ix
            Next i

            sumAboveAvg = 0.0

            For ix = 0 To zeitraum
                tmpSum = 0.0
                For i = 1 To myCollection.Count
                    tmpSum = tmpSum + rcValues(i - 1, ix)
                Next i
                sumAboveAvg = sumAboveAvg + (tmpSum - avgValue) * (tmpSum - avgValue)
            Next ix

            getDeviationfromAverage = sumAboveAvg


        End Get
    End Property
    '
    ' property gibt die Personalkosten zurück, die durch die internen Rollen entstehen, die in keinen Projekten gebunden sind - ohne Projekte
    '
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
                kapaValues = Me.getRoleKapasInMonth(myCollection)
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

    Public Sub New()

        AllProjects = New SortedList(Of String, clsProjekt)
        AllShapes = New SortedList(Of String, String)

    End Sub

End Class
