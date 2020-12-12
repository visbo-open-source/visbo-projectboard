Imports Microsoft.Office.Interop.Excel
Public Class clsPhase

    ' earliestStart und latestStart sind absolute Werte im "koordinaten-System" des Projektes
    ' von daher ist es anders gelöst als in clsProjekt, wo earlieststart und latestStart relative Angaben sind 
    ' tk Änderung 26.10.17 , 
    ' - es wurde ein Attribut für percentDone aufgenommen
    ' - Phasen können jetzt auch Deliverables haben, damit muss eine Phase nicht mit einem Meilenstein abgeschlossen werden , um ein 
    '   oder mehrere Deliverables zu haben 
    Private _percentDone As Double

    ' Liste an Deliverables, die die Phase haben kann 
    Private _deliverables As List(Of String)

    Private _nameID As String
    Private _parentProject As clsProjekt
    Private _vorlagenParent As clsProjektvorlage

    Private _shortName As String
    Private _originalName As String
    Private _appearance As String
    Private _color As Integer

    ' die Dokumenten Url für den Meilenstein
    Private _docURL As String

    ' die Applikations-ID mit der die Dok-Url geöffnet werden kann / soll
    Private _docUrlAppID As String

    ' wer ist für die Phase, die Ergebnisse und Einhaltung der Ressourcen verantwortlich? 
    Private _verantwortlich As String
    ' wird benötigt, um bei Optimierungs-Läufen einen Tryout Wert zu haben ..
    Private _offset As Integer
    ' ist der eigentlich Offsetin Tagen vom Projekt-Start weg gerechnet
    Private _startOffsetinDays As Integer

    Private _earliestStart As Integer
    Private _latestStart As Integer

    Private _relStart As Integer
    Private _relEnde As Integer

    Private _dauerInDays As Integer


    Private _bewertungen As SortedList(Of String, clsBewertung)
    Private _allMilestones As List(Of clsMeilenstein)
    Private _allRoles As List(Of clsRolle)
    Private _allCosts As List(Of clsKostenart)

    ' tk ergänzt am 12,6,20
    Private _invoice As KeyValuePair(Of Double, Integer)
    Private _penalty As KeyValuePair(Of Date, Double)

    ''' <summary>
    ''' liest / schreibt den Betrag, der beim Erreichen dieses Meilensteins als Rechnung gestellt werden kann 
    ''' key: Summe in T€
    ''' Value: Terms of payment
    ''' Vorsicht: kann Nothing sein. 
    ''' </summary>
    ''' <returns></returns>
    Public Property invoice As KeyValuePair(Of Double, Integer)
        Get
            invoice = _invoice
        End Get
        Set(value As KeyValuePair(Of Double, Integer))
            If Not IsNothing(value) Then
                If value.Key >= 0 And value.Value >= 0 Then
                    _invoice = value
                End If
            Else
                _invoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
            End If
        End Set
    End Property

    Public Property penalty As KeyValuePair(Of Date, Double)
        Get
            penalty = _penalty
        End Get
        Set(value As KeyValuePair(Of Date, Double))
            If Not IsNothing(value) Then
                If value.Key = Date.MinValue Then
                    _penalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, value.Value)
                Else
                    _penalty = value
                End If

            Else
                _penalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, 0.0)
            End If
        End Set
    End Property

    ''' <summary>
    ''' löscht alle Rollen der Phase
    ''' </summary>
    Public Sub clearRoles()
        _allRoles.Clear()
    End Sub

    '''' <summary>
    '''' entfernt die Rolle mit Name rolename aus der Phase
    '''' wenn die nicht als Rollendefinition gar nicht existiert, gibt es eine Exception
    '''' andernfalls, wenn Rolle nur nicht in der Phase vorkommt, gibt es keine Meldung 
    '''' 
    '''' </summary>
    '''' <param name="roleName"></param>
    'Public Sub deleteRole(ByVal roleName As String)

    '    If RoleDefinitions.containsName(roleName) Then
    '        Dim ix As Integer = 0
    '        Dim found As Boolean = False

    '        While Not found And ix <= _allRoles.Count - 1
    '            If _allRoles.Item(ix).name = roleName Then
    '                found = True
    '            Else
    '                ix = ix + 1
    '            End If
    '        End While

    '        If found Then
    '            _allRoles.RemoveAt(ix)
    '        End If
    '    Else
    '        'Fehler ...
    '        Dim errmsg As String
    '        If awinSettings.englishLanguage Then
    '            errmsg = "role unknown: " & roleName
    '        Else
    '            errmsg = "unbekannte Rolle: " & roleName
    '        End If
    '        Throw New ArgumentException(errmsg)
    '    End If

    'End Sub

    '''' <summary>
    '''' entfernt die Kostenart mit Name costname aus der Phase
    '''' wenn die als Kostenartdefinition gar nicht existiert, gibt es eine Exception
    '''' andernfalls, wenn Kostenart nur nicht in der Phase vorkommt, gibt es keine Meldung 
    '''' </summary>
    '''' <param name="costname"></param>
    'Public Sub deleteCost(ByVal costname As String)
    '    If CostDefinitions.containsName(costname) Then
    '        Dim ix As Integer = 0
    '        Dim found As Boolean = False

    '        While Not found And ix <= _allCosts.Count - 1
    '            If _allCosts.Item(ix).name = costname Then
    '                found = True
    '            Else
    '                ix = ix + 1
    '            End If
    '        End While

    '        If found Then
    '            _allCosts.RemoveAt(ix)
    '        End If
    '    Else
    '        'Fehler ...
    '        Dim errmsg As String
    '        If awinSettings.englishLanguage Then
    '            errmsg = "role unknown: " & costname
    '        Else
    '            errmsg = "unbekannte Rolle: " & costname
    '        End If
    '        Throw New ArgumentException(errmsg)
    '    End If

    'End Sub

    ''' <summary>
    ''' löscht alle Kostenbedarfe der Phase
    ''' </summary>
    Public Sub clearCosts()
        _allCosts.Clear()
    End Sub

    ''' <summary>
    ''' es wird eine PercentDone Regelung eingeführt , mit der beurteilt werden kann, wie wit die Ergebnisse bereits sind  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property percentDone As Double
        Get
            percentDone = _percentDone
        End Get
        Set(value As Double)

            If Not IsNothing(value) Then
                If value >= 0 Then
                    If value <= 1.0 Then
                        _percentDone = value
                    Else
                        ' dann müssen die PErcentDone Werte erst noch normiert werden 
                        _percentDone = value / 100
                    End If

                Else
                    Throw New ArgumentException("percent Done Value must not be negativ ...")
                End If
            Else
                ' einfach nichts tun ... 
            End If


        End Set
    End Property

    ''' <summary>
    ''' liest schreibt den String, der eine Dokumenten URL darstellt, wo Dokumente abgelegt sind, die zum Meilenstein gehören 
    ''' </summary>
    ''' <returns></returns>
    Public Property DocURL() As String
        Get
            DocURL = _docURL
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _docURL = value
            Else
                _docURL = ""
            End If
        End Set
    End Property

    ''' <summary>
    ''' liest schreibt den String, der die ID der Appliaktion darstellt, mit der auf die Dokumenten Url zugegriffen werden kann 
    ''' </summary>
    ''' <returns></returns>
    Public Property DocUrlAppID() As String
        Get
            DocUrlAppID = _docUrlAppID
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _docUrlAppID = value
            Else
                _docUrlAppID = ""
            End If
        End Set
    End Property

    ''' <summary>
    ''' summiert die tValues ab dem Start-Element in die Phasen-Xvalues 
    ''' </summary>
    ''' <param name="tValues">der Werte Array</param>
    ''' <param name="start">1 ... dauer: soll ab dem ersten oder x. Element addiert werden </param>
    ''' <remarks></remarks>
    Public Sub addTaskEfforts(ByVal tValues() As Double, _
                              ByVal rcID As Integer, ByVal rcType As Integer, _
                              ByVal start As Integer)

        If tValues.Length + start - 1 > _relEnde - relStart + 1 Then
            Throw New ArgumentException("dimensions of values do not fit")
        Else
            If rcType = PThcc.persbedarf Then

                Dim rcName As String = RoleDefinitions.getRoleDefByID(rcID).name
                Dim role As clsRolle = Me.getRole(rcName)
                If Not IsNothing(role) Then
                    For i As Integer = 1 To tValues.Length
                        role.Xwerte(start - 1) = role.Xwerte(start - 1) + tValues(i - 1)
                    Next
                Else
                    Dim dimension As Integer = _relEnde - _relStart
                    role = New clsRolle(dimension)
                    With role
                        .uid = rcID
                        For i As Integer = 1 To tValues.Length
                            role.Xwerte(start - 1) = role.Xwerte(start - 1) + tValues(i - 1)
                        Next
                    End With
                    ' Rolle hinzufügen
                    With Me
                        .addRole(role)
                    End With
                End If

            ElseIf rcType = PThcc.othercost Then

                Dim rcName As String = CostDefinitions.getCostdef(rcID).name
                Dim ccost As clsKostenart = Me.getCost(rcName)
                If Not IsNothing(ccost) Then
                    For i As Integer = 1 To tValues.Length
                        ccost.Xwerte(start - 1) = ccost.Xwerte(start - 1) + tValues(i - 1)
                    Next
                Else
                    Dim dimension As Integer = _relEnde - _relStart
                    ccost = New clsKostenart(dimension)
                    With ccost
                        .KostenTyp = rcID
                        For i As Integer = 1 To tValues.Length
                            ccost.Xwerte(start - 1) = ccost.Xwerte(start - 1) + tValues(i - 1)
                        Next
                    End With
                    ' Rolle hinzufügen
                    With Me
                        .AddCost(ccost)
                    End With
                End If

            End If

        End If
    End Sub

    ''' <summary>
    ''' gibt zurück, ob die Phase identisch mit der übergebenen Phase ist  
    ''' </summary>
    ''' <param name="vPhase"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vPhase As clsPhase) As Boolean
        Get
            Dim stillOK As Boolean = False
            Dim ix As Integer = 0

            Try
                ' administratives ...
                ' tk 16.5 Namensgleicheit reicht hier eigentlich
                ' sonmst wird das bei ensureIDStability zur Ungleichheit führen 
                ' und eine Phase, die alle Attribute identisch hat , aber in der lfd Nmmer abweicht , ist als identisch zusenen 
                ' If Me.nameID = vPhase.nameID Then
                If Me.name = vPhase.name Then

                    If Me.dauerInDays = vPhase.dauerInDays And
                            Me.startOffsetinDays = vPhase.startOffsetinDays Then

                        If Me.countCosts = vPhase.countCosts And
                                Me.countRoles = vPhase.countRoles And
                                Me.countDeliverables = vPhase.countDeliverables And
                                Me.countMilestones = vPhase.countMilestones And
                                Me.DocURL = vPhase.DocURL And
                                Me.DocUrlAppID = vPhase.DocUrlAppID And
                                Me.percentDone = vPhase.percentDone Then
                            'ur: 20180110 Me.bewertungsCount = .bewertungsCount Then

                            If Me.ampelErlaeuterung = vPhase.ampelErlaeuterung And
                                    Me.ampelStatus = vPhase.ampelStatus Then

                                If Me.shortName = vPhase.shortName And
                                        Me.originalName = vPhase.originalName And
                                        Me.verantwortlich = vPhase.verantwortlich Then

                                    ' tk 25.11.19 appearance und individualColor müssen nicht gecheckt werden

                                    If Me.earliestStart = vPhase.earliestStart And
                                           Me.latestStart = vPhase.latestStart And
                                           Me.offset = vPhase.offset Then

                                        stillOK = True

                                    End If

                                    'If Me.appearance = vPhase.appearance And
                                    '        Me.individualColor = vPhase.individualColor And
                                    '        Me.earliestStart = vPhase.earliestStart And
                                    '        Me.latestStart = vPhase.latestStart And
                                    '        Me.offset = vPhase.offset Then

                                    '    stillOK = True

                                    'End If

                                End If

                            End If

                        End If

                    End If

                End If

                ' jetzt die Deliverables prüfen  
                If stillOK Then
                    Dim MeDelis As String = Me.getAllDeliverables("#")
                    Dim vglDelis As String = vPhase.getAllDeliverables("#")

                    If MeDelis = vglDelis Then
                        ' prüfen auf Bewertungen ... 
                        ix = 1
                        Do While stillOK And ix <= Me.bewertungsCount
                            Dim MeBewertung As clsBewertung = Me.getBewertung(ix)
                            Dim vglBewertung As clsBewertung = vPhase.getBewertung(ix)
                            If MeBewertung.isIdenticalTo(vglBewertung) Then
                                ix = ix + 1
                            Else
                                stillOK = False
                            End If
                        Loop
                    Else
                        stillOK = False
                    End If

                End If


                ' jetzt die Rollen, Kosten, Milestones und Bewertungen abfragen 
                If stillOK Then
                    ' sind die Rollen identisch 
                    ix = 1
                    Do While stillOK And ix <= Me.countRoles
                        Dim MeRole As clsRolle = Me.getRole(ix)
                        Dim vglRole As clsRolle = vPhase.getRole(ix)
                        If MeRole.isIdenticalTo(vglRole) Then
                            ix = ix + 1
                        Else
                            stillOK = False
                        End If
                    Loop

                    If stillOK Then
                        ' sind die Kostenarten identisch ?
                        ix = 1
                        Do While stillOK And ix <= Me.countCosts
                            Dim MeCost As clsKostenart = Me.getCost(ix)
                            Dim vglCost As clsKostenart = vPhase.getCost(ix)
                            If MeCost.isIdenticalTo(vglCost) Then
                                ix = ix + 1
                            Else
                                stillOK = False
                            End If
                        Loop

                        If stillOK Then
                            ' sind die Phasen Bewertungen identisch?
                            ix = 1
                            Do While stillOK And ix <= Me.bewertungsCount
                                Dim MeBewertung As clsBewertung = Me.getBewertung(ix)
                                Dim vglBewertung As clsBewertung = vPhase.getBewertung(ix)
                                If MeBewertung.isIdenticalTo(vglBewertung) Then
                                    ix = ix + 1
                                Else
                                    stillOK = False
                                End If
                            Loop

                            If stillOK Then
                                ' jetzt die Meilensteine, Bewertungen und Deliverables prüfen ... 
                                ix = 1
                                Do While stillOK And ix <= Me.countMilestones
                                    Dim MeMs As clsMeilenstein = Me.getMilestone(ix)
                                    Dim vglMs As clsMeilenstein = vPhase.getMilestone(ix)
                                    If MeMs.isIdenticalTo(vglMs) Then
                                        ix = ix + 1
                                    Else
                                        stillOK = False
                                    End If
                                Loop
                            End If

                        End If

                    End If

                End If

                ' jetzt die Invoices und Penalties abfragen 
                If stillOK Then
                    stillOK = Me.invoice.Key = vPhase.invoice.Key And
                        Me.invoice.Value = vPhase.invoice.Value And
                        Me.penalty.Key = vPhase.penalty.Key And
                        Me.penalty.Value = vPhase.penalty.Value
                End If



            Catch ex As Exception
                stillOK = False
            End Try

            isIdenticalTo = stillOK

        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl Deliverables für diese Phase zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property countDeliverables As Integer
        Get
            countDeliverables = _deliverables.Count
        End Get
    End Property

    ''' <summary>
    ''' gibt an , ob das Deliverable existiert ...
    ''' </summary>
    ''' <param name="item"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsDeliverable(ByVal item As String) As Boolean
        Get
            containsDeliverable = _deliverables.Contains(item)
        End Get
    End Property

    ''' <summary>
    ''' löscht alle Deliverables des Meilensteines 
    ''' </summary>
    Public Sub clearDeliverables()
        _deliverables.Clear()
    End Sub

    ''' <summary>
    ''' fügt das Deliverable Item der Liste hinzu; 
    ''' wenn das Item bereits in der Liste vorhanden ist, passiert nichts 
    ''' </summary>
    ''' <param name="item"></param>
    ''' <remarks></remarks>
    Public Sub addDeliverable(ByVal item As String)

        If Not _deliverables.Contains(item) Then
            _deliverables.Add(item)
        End If

    End Sub

    ''' <summary>
    ''' gibt das Element an der bezeichneten Stelle zurück
    ''' index kann Werte zwischen 1 .. count annehmen 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getDeliverable(ByVal index As Integer) As String
        Get
            Dim tmpValue As String = ""
            If index >= 1 And index <= _deliverables.Count Then
                tmpValue = _deliverables.Item(index - 1)
            End If
            getDeliverable = tmpValue
        End Get
    End Property

    ''' <summary>
    ''' gibt die Liste der Deliverables einer Phase als einen String zurück; 
    ''' die einzelnen Deliverables sind by default durch einen vblf getrennt
    ''' oder getrennt durch das übergebene trennzeichen  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAllDeliverables(ByVal trennzeichen As String) As String
        Get
            Dim tmpDeliverables As String = ""
            Dim tmp_i As Integer = 1

            For i As Integer = 1 To _deliverables.Count
                ' ur:07.02.2020 nur nicht leere Deliverables sind relevant
                If _deliverables.Item(i - 1) <> "" Then
                    If tmp_i = 1 Then
                        tmpDeliverables = _deliverables.Item(i - 1)
                        tmp_i = tmp_i + 1
                    Else
                        tmpDeliverables = tmpDeliverables & trennzeichen &
                            _deliverables.Item(i - 1)
                        tmp_i = tmp_i + 1
                    End If
                End If

            Next

            getAllDeliverables = tmpDeliverables

        End Get
    End Property



    ''' <summary>
    ''' liest/schreibt das Feld für verantwortlich
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property verantwortlich As String
        Get
            verantwortlich = _verantwortlich
        End Get
        Set(value As String)
            _verantwortlich = value
        End Set
    End Property

    Public Sub addBewertung(ByVal b As clsBewertung)
        Dim key As String

        If Not b.bewerterName Is Nothing Then
            key = b.bewerterName & "#" & b.datum.ToString("MMM yy")
        Else
            key = "#" & b.datum.ToString("MMM yy")
        End If

        Try
            If _bewertungen.ContainsKey(key) Then
                _bewertungen.Remove(key)
            End If
            _bewertungen.Add(key, b)
        Catch ex As Exception
            Throw New ArgumentException("Bewertung wurde bereits vergeben ..")
        End Try

    End Sub

    ''' <summary>
    ''' gibt Anzahl Bewertungen zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property bewertungsCount As Integer
        Get
            bewertungsCount = _bewertungen.Count

        End Get
    End Property

    ''' <summary>
    ''' löscht die Bewertungen der Phase
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearBewertungen()

        Try
            _bewertungen.Clear()
        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' gibt die Bewertungsliste zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property bewertungsListe() As SortedList(Of String, clsBewertung)

        Get
            bewertungsListe = _bewertungen
        End Get
    End Property

    ''' <summary>
    ''' gibt die Bewertung mit der angegebenen Nr zurück
    ''' Nr kann zwischen 1 und Count liegen  
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBewertung(ByVal index As Integer) As clsBewertung

        Get

            If index > _bewertungen.Count Then
                'getBewertung = Nothing
                getBewertung = New clsBewertung
            Else
                getBewertung = _bewertungen.ElementAt(index - 1).Value
            End If

        End Get

    End Property

    ''' <summary>
    ''' liest / setzt die individuelle appearance für diese Phase 
    ''' normalerweise wird die Appearance aber über die PhaseDefinitions oder missingPhaseDefinitions definiert 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property appearanceName As String
        Get
            appearanceName = _appearance
        End Get
        Set(value As String)
            If appearanceDefinitions.liste.ContainsKey(value) Then
                _appearance = value
            Else
                _appearance = awinSettings.defaultPhaseClass
            End If
        End Set
    End Property


    ''' <summary>
    ''' liest/schreibt den Ampel-Status, das ist die 1. Bewertung
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ampelStatus As Integer
        Get
            If Me.bewertungsCount >= 1 Then
                ampelStatus = Me.getBewertung(Me.bewertungsCount).colorIndex
            Else
                ampelStatus = 0
            End If
        End Get

        Set(value As Integer)
            If IsNothing(value) Then
                value = 0
            ElseIf value < 0 Or value > 3 Then
                value = 0
            End If

            If Me.bewertungsCount >= 1 Then
                Me.getBewertung(Me.bewertungsCount).colorIndex = value
            Else

                Dim tmpB As New clsBewertung
                With tmpB
                    .description = ""
                    .colorIndex = value
                End With

                Me.addBewertung(tmpB)

            End If
        End Set

    End Property

    ''' <summary>
    ''' liest/schreibt die Ampel-Erläuterung, das ist die 1. Bewertung
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ampelErlaeuterung As String
        Get
            If Me.bewertungsCount >= 1 Then
                ampelErlaeuterung = Me.getBewertung(Me.bewertungsCount).description
            Else
                ampelErlaeuterung = ""
            End If
        End Get
        Set(value As String)
            If IsNothing(value) Then
                value = ""
            End If

            If Me.bewertungsCount >= 1 Then
                Me.getBewertung(Me.bewertungsCount).description = value
            Else
                Dim tmpB As New clsBewertung
                With tmpB
                    .description = value
                    .colorIndex = 0
                End With

                Me.addBewertung(tmpB)

            End If

        End Set

    End Property

    ''' <summary>
    ''' prüft ob die Phase in ihren Werten Dauer in Monaten konsistent zu den Xwert-Dimensionen der Rollen und Kosten ist 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isConsistent As Boolean

        Get
            Dim tmpValue As Boolean = True
            Dim dimension As Integer
            Dim phaseStart As Date, phaseEnd As Date
            Dim r As Integer = 1, k As Integer = 1

            ' prüfen, ob die Gesamtlänge übereinstimmt  


            phaseEnd = Me.getEndDate
            phaseStart = Me.getStartDate

            dimension = getColumnOfDate(phaseEnd) - getColumnOfDate(phaseStart)

            While tmpValue And r <= Me.countRoles
                If dimension <> Me.getRole(r).Xwerte.Length - 1 Then
                    tmpValue = False
                End If
                r = r + 1
            End While

            While tmpValue And k <= Me.countCosts
                If dimension <> Me.getCost(k).Xwerte.Length - 1 Then
                    tmpValue = False
                End If
                k = k + 1
            End While


            isConsistent = tmpValue

        End Get

    End Property

    ''' <summary>
    ''' wird verwendet um Termine entweder per Drag and Drop zu verändern , unter Berücksichtigung der ActualData 
    ''' oder aber im MassEditTermine 
    ''' </summary>
    ''' <param name="newOffsetInTagen"></param>
    ''' <param name="newDauerInTagen"></param>
    ''' <param name="autoAdjustChilds"></param>
    ''' <returns></returns>
    Public Function adjustPhaseAndChilds(ByVal newOffsetInTagen As Long, ByVal newDauerInTagen As Long,
                                         ByVal autoAdjustChilds As Boolean) As clsPhase

        Dim tmpResult As clsPhase = Nothing

        Dim elemID As String = Me.nameID

        Dim hproj As clsProjekt = parentProject

        Dim deltaOffset As Long = newOffsetInTagen - Me.startOffsetinDays
        Dim deltaDauer As Long = newDauerInTagen - Me.dauerInDays

        ' Merken des Offsets Phase, die später Parent ihrer childs ist 
        Dim parentPhaseOldOffset As Long = Me.startOffsetinDays

        Dim faktor As Double = 1.0

        If Me.dauerInDays > 0 Then
            faktor = newDauerInTagen / Me.dauerInDays
        End If


        ' jetzt wird diese Phase entsprechend geändert ...
        Call Me.adjustStartandDauer(newOffsetInTagen, newDauerInTagen)

        If autoAdjustChilds Then

            ' jetzt die Kind-Phasen anpassen 
            For Each childPhaseNameID As String In hproj.hierarchy.getChildIDsOf(elemID, False)

                Dim childPhase As clsPhase = hproj.getPhaseByID(childPhaseNameID)
                Dim childPhaseDeltaOffset As Long = childPhase.startOffsetinDays - parentPhaseOldOffset

                'Dim newChildOffset As Long = CLng(faktor * childPhase.startOffsetinDays)
                Dim newChildOffset As Long = newOffsetInTagen + CLng(faktor * childPhaseDeltaOffset)
                Dim newChildDuration As Long = CLng(faktor * childPhase.dauerInDays)

                Dim newCalculationNecessary As Boolean = (childPhase.getStartDate.Date <> parentProject.startDate.AddDays(newChildOffset).Date) Or
                                                    (childPhase.getEndDate.Date <> parentProject.startDate.AddDays(newChildOffset + newChildDuration - 1).Date)

                ' jetzt prüfen, ob es actualdata gibt 
                If hproj.hasActualValues Then
                    If getColumnOfDate(childPhase.getStartDate) <= getColumnOfDate(hproj.actualDataUntil) Then
                        ' bisheriges Startdatum liegt vor ActualData-Date: es darf gar nicht verändert werden 
                        Dim diffOffset As Long = DateDiff(DateInterval.Day, childPhase.getStartDate.Date, parentProject.startDate.AddDays(newChildOffset).Date)

                        ' hier muss das aktuelle Projekt-Ende Datum ermittlet werden 
                        If diffOffset <> 0 Then
                            ' neu bestimmen der Notwendigkeit für Neuberechnung 
                            newCalculationNecessary = (childPhase.getStartDate.Date <> parentProject.startDate.AddDays(newChildOffset).Date) Or
                                                    (childPhase.getEndDate.Date <> parentProject.startDate.AddDays(newChildOffset + newChildDuration + diffOffset - 1).Date)
                        End If

                        'der Offset muss unverändert bleiben, da das Startdatum links vom ActualData Date liegt ..
                        newChildOffset = childPhase.startOffsetinDays


                        If getColumnOfDate(childPhase.getEndDate) <= getColumnOfDate(hproj.actualDataUntil) Then
                            ' bisheriges Endedatum liegt vor ActualData-Date: unverändert lassen ...
                            newChildDuration = childPhase.dauerInDays

                            ' in diesem Fall ist keine Neu-Berechnung notwednig bzw. es führt dann zu Fehlern ... 
                            ' weil Ende-Datum vor dem ActualDataUntil liegt 
                            newCalculationNecessary = False
                        Else
                            ' liegt das neue Ende-Datum vor ActualData Date? 
                            If hproj.startDate.AddDays(newChildOffset + newChildDuration - 1).Date <= hproj.actualDataUntil.Date Then
                                ' wird auf den letzten Tag des zum ActualDataUntil folgenden Monats gelegt
                                newChildDuration = DateDiff(DateInterval.Day, hproj.startDate.AddDays(newChildOffset).Date, getDateofColumn(getColumnOfDate(hproj.actualDataUntil) + 1, True).Date) + 1

                            Else
                                ' kann übernommen werden , newChildDuration ist ohnehin schon gesetzt 
                                ' hier muss jetzt die ChildDuration um den diffOffset korrigiert werden 
                                ' wenn das Startdatum nicht fet´stgehalten würde, dann wäre das Enddatum entsprechend weiter hinten buw. vorne - 
                                ' deshalb muss der Duration Wert jetzt korrigiert werden, um dem Rechnung zu tragen  
                                If diffOffset > 0 Then
                                    ' das Phasen Ende wird nach rechts verschoben 
                                    newChildDuration = newChildDuration + diffOffset

                                Else
                                    'newChildDuration = newChildDuration + diffOffset
                                    ' das Phasen Ende wird nach links verschoben , darf aber nicht weiter als bis zum Ende des Folge-Monats auf ActualDataUntil sein 
                                    If DateDiff(DateInterval.Day, getDateofColumn(getColumnOfDate(hproj.actualDataUntil), True).Date, hproj.startDate.AddDays(newChildOffset + newChildDuration - 1).Date) > 0 Then
                                        ' alles in Ordnung 

                                    Else
                                        newChildDuration = DateDiff(DateInterval.Day, hproj.startDate.AddDays(newChildOffset).Date, getDateofColumn(getColumnOfDate(hproj.actualDataUntil) + 1, True).Date) + 1
                                    End If

                                End If

                            End If

                            If newChildOffset + newChildDuration <= newOffsetInTagen + newDauerInTagen Then
                                ' alles in Ordnung 
                            Else
                                newChildDuration = newOffsetInTagen + newDauerInTagen - newChildOffset
                            End If

                            ' wurde durch den oberen Absatz ersetzt 
                            '' gilt für alle oberen Zweige ... 
                            'If DateDiff(DateInterval.Day, hproj.startDate.AddDays(newChildOffset + newChildDuration - 1).Date, hproj.endeDate.Date) > 0 Then
                            '    ' alles in Ordnung 
                            'Else
                            '    newChildDuration = DateDiff(DateInterval.Day, hproj.startDate.AddDays(newChildOffset).Date, hproj.endeDate.Date) + 1
                            'End If

                            newCalculationNecessary = (childPhase.getStartDate.Date <> hproj.startDate.AddDays(newChildOffset).Date) Or
                                                    (childPhase.getEndDate.Date <> hproj.startDate.AddDays(newChildOffset + newChildDuration - 1).Date)

                        End If

                    Else

                        ' hier muss aber noch überprüft werden, ob das neue (!) Startdatum vor dem hproj.actualdata liegt 
                        If getColumnOfDate(hproj.startDate.AddDays(newChildOffset).Date) <= getColumnOfDate(hproj.actualDataUntil) Then
                            ' das Startdatum der Phase  nach dem ActualData-Datum schieben  
                            newChildOffset = DateDiff(DateInterval.Day, hproj.startDate.Date, getDateofColumn(getColumnOfDate(hproj.actualDataUntil) + 1, False).Date)

                            If newChildOffset + newChildDuration <= newOffsetInTagen + newDauerInTagen Then
                                ' alles in Ordnung 
                            Else
                                newChildDuration = newOffsetInTagen + newDauerInTagen - newChildOffset
                            End If

                            ' wurde durch den oberen Absatz ersetzt 
                            'If DateDiff(DateInterval.Day, hproj.startDate.AddDays(newChildOffset + newChildDuration - 1).Date, hproj.endeDate.Date) > 0 Then
                            '    ' alles in Ordnung 
                            'Else
                            '    newChildDuration = DateDiff(DateInterval.Day, hproj.startDate.AddDays(newChildOffset).Date, hproj.endeDate.Date) + 1
                            'End If

                            newCalculationNecessary = (childPhase.getStartDate.Date <> hproj.startDate.AddDays(newChildOffset).Date) Or
                                                    (childPhase.getEndDate.Date <> hproj.startDate.AddDays(newChildOffset + newChildDuration - 1).Date)

                        Else
                            ' kann komplett übernommen werden 
                            ' das neue startdatum liegt rechts von hproj.ActualDataUntil ..
                        End If

                    End If



                End If


                If newChildDuration = 0 Then
                    newChildDuration = 1
                End If

                Try
                    If newCalculationNecessary Then
                        childPhase = childPhase.adjustPhaseAndChilds(newChildOffset, newChildDuration, autoAdjustChilds)
                    End If
                Catch ex As Exception

                End Try

            Next


            ' jetzt die Meilensteine der Phase  anpassen 
            For Each childMilestoneNameID As String In hproj.hierarchy.getChildIDsOf(elemID, True)

                Dim childMilestone As clsMeilenstein = hproj.getMilestoneByID(childMilestoneNameID)
                Dim newChildOffset As Long = CLng(childMilestone.offset * faktor)
                ' jetzt prüfen, ob es actualdata gibt 

                If hproj.hasActualValues Then
                    If getColumnOfDate(childMilestone.getDate) <= getColumnOfDate(hproj.actualDataUntil) Then
                        ' bisheriges Meilensteindatum liegt vor ActualData-Date: unverändert lassen ...
                        newChildOffset = childMilestone.offset

                    Else
                        ' liegt das neue Datum vor ActualData Date? 
                        If Me.getStartDate.AddDays(newChildOffset).Date <= hproj.actualDataUntil.Date Then
                            ' wird auf den ersten des zum ActualDataUntil folgenden Monats gelegt
                            newChildOffset = DateDiff(DateInterval.Day, Me.getStartDate, getDateofColumn(getColumnOfDate(hproj.actualDataUntil) + 1, False))
                        Else
                            ' kann übernommen werden , newChildOffset
                        End If
                    End If
                End If

                ' falls der Rundungsfehler zu einem zu späten Meilenstein führt ... 
                If newChildOffset > Me.dauerInDays - 1 Then
                    newChildOffset = Me.dauerInDays - 1
                End If

                ' falls der Rundungsfehler zu einem zu frühen Meilenstein führt ... 
                If newChildOffset < 0 Then
                    newChildOffset = 0
                End If

                childMilestone.setDate = Me.getStartDate.AddDays(newChildOffset)

            Next

        End If

        adjustPhaseAndChilds = Me

    End Function

    ''' <summary>
    ''' ähnlich wie changeStartAnd Dauer, nur mit Modifikationen, die für adjustPhaseAndChilds notwendig sind ... 
    ''' ändert die Daten der Phase, also Startdatum und Ende-Datum. 
    ''' Allerdings nur , wenn erlaubt. 
    ''' Nicht erlaubt: es gibt actualData, Starttermin liegt vor ActualData und soll verschoeben werden -> geht nicht 
    ''' Start- oder Ende-Termin soll vor ActualData verschoeben werden ... 
    ''' </summary>
    ''' <param name="startOffset"></param>
    ''' <param name="dauer"></param>
    Private Sub adjustStartandDauer(ByVal startOffset As Long, ByVal dauer As Long)
        Dim projektStartdate As Date
        Dim projektstartColumn As Integer
        Dim oldDauerinDays As Integer = Me._dauerInDays
        Dim faktor As Double
        Dim dimension As Integer

        Dim errMsg As String = ""

        ' hier muss unterschieden werden, ob Me.dauerIndays überhaupt schon was enthält, andernfalls muss keine Neuberechnung der Xwerte erfolgen
        ' die muss nur dann erfolgen wenn aus zwei enthaltenen Monaten plötzlich drei werden . Dann muss die Bedarfs-Summe eben entsprechend neu verteilt werden  
        Dim newCalculationNecessary As Boolean = (Me.nameID = rootPhaseName) Or
                                                    (((Me.getStartDate.Date <> parentProject.startDate.AddDays(startOffset).Date) Or
                                                    (Me.getEndDate.Date <> parentProject.startDate.AddDays(startOffset + dauer - 1).Date)) And
                                                    Me.dauerInDays > 0)

        ' damit wird bestimmt, ob die Verteilung auch dann neu berechnet werden soll, wenn die Dimension des alten und des neuen Arrays gleich ist.  
        Dim calcAnyhow As Boolean = True

        If Me.nameID <> rootPhaseName And Not IsNothing(parentProject) Then
            If System.Math.Abs(Me.getStartDate.Day - parentProject.startDate.AddDays(startOffset).Day) <= 1 And
            System.Math.Abs(dauer - Me.dauerInDays) <= 1 Then
                calcAnyhow = False
            End If
        End If


        If dauer < 0 Then
            If awinSettings.englishLanguage Then
                errMsg = "Dauer must not be negative!"
            Else
                errMsg = "Dauer kann nicht negativ sein!"
            End If

            Throw New ArgumentException(errMsg)

        ElseIf startOffset < 0 Then

            If awinSettings.englishLanguage Then
                errMsg = "Phase may not begin before project starts!"
            Else
                errMsg = "Phase kann nicht vor Projektstart beginnen"
            End If

            Throw New ArgumentException(errMsg)

        ElseIf Me.hasActualData And Me.dauerInDays > 0 Then
            ' wenn die Phase gerade aufgebaut wird, darf das kein Abbruch geben ..
            ' unzulässig Startdatum verändert sich , altes oder neues Startdatum liegt vor ActualDatauntil 
            If Me.startOffsetinDays <> startOffset Then
                If Me.getStartDate < parentProject.actualDataUntil Or parentProject.startDate.AddDays(startOffset) < parentProject.actualDataUntil Then
                    ' unzulässig 

                    If awinSettings.englishLanguage Then
                        errMsg = "Start-Date may not be changed because of existing actual data!"
                    Else
                        errMsg = "Start-Datum kann nicht verändert werden, da es bereits Ist-Daten gibt. "
                    End If

                    Throw New ArgumentException(errMsg)
                End If
            End If

            ' Überprüfung des Ende-Datums 
            If parentProject.startDate.AddDays(startOffset + dauer - 1).Date < parentProject.actualDataUntil.Date Then
                ' unzulässig 

                If awinSettings.englishLanguage Then
                    errMsg = "End-Date may not be before actual data - date!"
                Else
                    errMsg = "Ende-Datum kann nicht vor das Ist-Daten Datum gelegt werden ... "
                End If

                Throw New ArgumentException(errMsg)

            End If

        End If


        Try
            ' Änderung tk, 20.6.18 .startDate.Date um zu normieren ..
            projektStartdate = Me.parentProject.startDate.Date
            projektstartColumn = Me.parentProject.Start

            If dauer = 0 And _relEnde > 0 Then

                ' dann sind die Werte initial noch nicht gesetzt worden 
                _startOffsetinDays = CInt(DateDiff(DateInterval.Day, projektStartdate, projektStartdate.AddMonths(_relStart - 1)))
                _dauerInDays = calcDauerIndays(projektStartdate.AddDays(_startOffsetinDays), _relEnde - _relStart + 1, True)


            ElseIf dauer = 0 And _relEnde = 0 Then

                Throw New ArgumentException("Phase kann nicht Dauer = 0 haben ")

            Else
                '  
                If _dauerInDays > 0 And dauer > 0 And awinSettings.propAnpassRess = True Then
                    faktor = dauer / _dauerInDays
                Else
                    faktor = 1
                End If


                _startOffsetinDays = CInt(startOffset)
                _dauerInDays = CInt(dauer)



                Dim oldlaenge As Integer = _relEnde - _relStart + 1


                Dim phaseStartdate As Date = Me.getStartDate
                Dim phaseEndDate As Date = Me.getEndDate


                _relStart = getColumnOfDate(phaseStartdate) - projektstartColumn + 1
                _relEnde = getColumnOfDate(phaseEndDate) - projektstartColumn + 1

                ' jetzt muss geprüft werden, ob die Phase die Dauer des Projektes verlängert 
                ' dieser Aufruf korrigiert notfalls die intern gehaltene

                Try
                    If Not IsNothing(Me.parentProject.getPhase(1)) Then
                        If Me.nameID <> Me.parentProject.getPhase(1).nameID Then
                            ' wenn es nicht die erste Phase ist, die gerade behandelt wird, dann soll die erste Phase auf Konsistenz geprüft werden 
                            Me.parentProject.keepPhase1consistent(Me.startOffsetinDays + Me.dauerInDays)
                        End If
                    End If

                Catch ex As Exception
                    Dim b As Integer = 0
                End Try


                If newCalculationNecessary Then


                    Dim newvalues() As Double

                    dimension = _relEnde - _relStart
                    ReDim newvalues(dimension)

                    If Me.countRoles > 0 Or Me.countCosts > 0 Then

                        ' hier müssen jetzt die Xwerte neu gesetzt werden 
                        Call Me.calcNewXwerte(dimension, faktor, calcAnyhow:=calcAnyhow)

                    End If


                End If




            End If


        Catch ex As Exception
            ' bei einer Projektvorlage gibt es kein Datum - es sollen aber die Werte für Offset und Dauer übernommen werden

            If dauer = 0 And _relEnde > 0 Then


                ' dann sind die Werte initial noch nicht gesetzt worden 
                _startOffsetinDays = CInt(DateDiff(DateInterval.Day, StartofCalendar, StartofCalendar.AddMonths(_relStart - 1)))
                '_dauerInDays = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(_relStart - 1), _
                '                        StartofCalendar.AddMonths(_relEnde).AddDays(-1)) + 1
                _dauerInDays = calcDauerIndays(projektStartdate.AddDays(_startOffsetinDays), _relEnde - _relStart + 1, True)


            Else
                '  
                _startOffsetinDays = CInt(startOffset)
                _dauerInDays = CInt(dauer)

                _relStart = CInt(DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(startOffset)) + 1)
                _relEnde = CInt(DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(startOffset + _dauerInDays - 1)) + 1)


            End If

        End Try

    End Sub
    ''' <summary>
    ''' ändert die Daten der Phase, also Startdatum und Ende-Datum. 
    ''' Allerdings nur , wenn erlaubt. 
    ''' Nicht erlaubt: es gibt actualData, Starttermin liegt vor ActualData und soll verschoeben werden -> geht nicht 
    ''' Start- oder Ende-Termin soll vor ActualData verschoeben werden ... 
    ''' </summary>
    ''' <param name="startOffset"></param>
    ''' <param name="dauer"></param>
    Public Sub changeStartandDauer(ByVal startOffset As Long, ByVal dauer As Long)

        Dim projektStartdate As Date
        Dim projektstartColumn As Integer
        Dim oldDauerinDays As Integer = Me._dauerInDays
        Dim faktor As Double
        Dim dimension As Integer



        If dauer < 0 Then
            Throw New ArgumentException("Dauer kann nicht negativ sein")

        ElseIf startOffset < 0 Then
            Throw New ArgumentException("Phase kann nicht vor Projektstart beginnen")

        End If


        Try
            ' Änderung tk, 20.6.18 .startDate.Date um zu normieren ..
            projektStartdate = Me.parentProject.startDate.Date
            projektstartColumn = Me.parentProject.Start

            If dauer = 0 And _relEnde > 0 Then

                ' dann sind die Werte initial noch nicht gesetzt worden 
                _startOffsetinDays = CInt(DateDiff(DateInterval.Day, projektStartdate, projektStartdate.AddMonths(_relStart - 1)))
                _dauerInDays = calcDauerIndays(projektStartdate.AddDays(_startOffsetinDays), _relEnde - _relStart + 1, True)


            ElseIf dauer = 0 And _relEnde = 0 Then

                Throw New ArgumentException("Phase kann nicht Dauer = 0 haben ")

            Else
                '  
                If _dauerInDays > 0 And dauer > 0 Then
                    faktor = dauer / _dauerInDays
                Else
                    faktor = 1
                End If


                _startOffsetinDays = CInt(startOffset)
                _dauerInDays = CInt(dauer)



                Dim oldlaenge As Integer = _relEnde - _relStart + 1


                Dim phaseStartdate As Date = Me.getStartDate
                Dim phaseEndDate As Date = Me.getEndDate


                _relStart = getColumnOfDate(phaseStartdate) - projektstartColumn + 1
                _relEnde = getColumnOfDate(phaseEndDate) - projektstartColumn + 1

                ' jetzt muss geprüft werden, ob die Phase die Dauer des Projektes verlängert 
                ' dieser Aufruf korrigiert notfalls die intern gehaltene

                Try
                    If Not IsNothing(Me.parentProject.getPhase(1)) Then
                        If Me.nameID <> Me.parentProject.getPhase(1).nameID Then
                            ' wenn es nicht die erste Phase ist, die gerade behandelt wird, dann soll die erste Phase auf Konsistenz geprüft werden 
                            Me.parentProject.keepPhase1consistent(Me.startOffsetinDays + Me.dauerInDays)
                        End If
                    End If

                Catch ex As Exception
                    Dim b As Integer = 0
                End Try


                If awinSettings.autoCorrectBedarfe Then


                    Dim newvalues() As Double
                    Dim notYetDone As Boolean = True

                    dimension = _relEnde - _relStart
                    ReDim newvalues(dimension)

                    If Me.countRoles > 0 Then

                        ' hier müssen jetzt die Xwerte neu gesetzt werden 
                        Call Me.calcNewXwerte(dimension, faktor)
                        notYetDone = False

                    End If

                    If Me.countCosts > 0 And notYetDone Then

                        ' hier müssen jetzt die Xwerte neu gesetzt werden 
                        Call Me.calcNewXwerte(dimension, 1)

                    End If


                End If




            End If


        Catch ex As Exception
            ' bei einer Projektvorlage gibt es kein Datum - es sollen aber die Werte für Offset und Dauer übernommen werden

            If dauer = 0 And _relEnde > 0 Then


                ' dann sind die Werte initial noch nicht gesetzt worden 
                _startOffsetinDays = CInt(DateDiff(DateInterval.Day, StartofCalendar, StartofCalendar.AddMonths(_relStart - 1)))
                '_dauerInDays = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(_relStart - 1), _
                '                        StartofCalendar.AddMonths(_relEnde).AddDays(-1)) + 1
                _dauerInDays = calcDauerIndays(projektStartdate.AddDays(_startOffsetinDays), _relEnde - _relStart + 1, True)


            Else
                '  
                _startOffsetinDays = CInt(startOffset)
                _dauerInDays = CInt(dauer)

                _relStart = CInt(DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(startOffset)) + 1)
                _relEnde = CInt(DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(startOffset + _dauerInDays - 1)) + 1)


            End If

        End Try



    End Sub

    ''' <summary>
    ''' stellt sicher, daß die Länge der Phase 1 auch der Projektlänge entspricht 
    ''' </summary>
    ''' <param name="startOffset"></param>
    ''' <param name="dauer"></param>
    ''' <remarks></remarks>
    Public Sub changeStartandDauerPhase1(ByVal startOffset As Integer, ByVal dauer As Integer)

        Dim projektStartdate As Date
        Dim projektstartColumn As Integer
        Dim faktor As Double = 1.0
        Dim dimension As Integer

        ' hier muss unterschieden werden, ob Me.dauerIndays überhaupt schon was enthält, andernfalls muss keine Neuberechnung der Xwerte erfolgen
        ' die muss nur dann erfolgen wenn aus zwei enthaltenen Monaten plötzlich drei werden . Dann muss die Bedarfs-Summe eben entsprechend neu verteitl werden  
        Dim newCalculationNecessary As Boolean = (Me.nameID = rootPhaseName) Or ((startOffset <> Me.startOffsetinDays Or dauer <> Me.dauerInDays) And Me.dauerInDays > 0)

        If dauer < 0 Then
            Throw New ArgumentException("Dauer kann nicht negativ sein")

        ElseIf startOffset < 0 Then
            Throw New ArgumentException("Phase kann nicht vor Projektstart beginnen")

        End If


        Try
            ' Änderung tk 20.6.18 .startDate.Date um zu normieren 
            projektStartdate = Me.parentProject.startDate.Date
            projektstartColumn = Me.parentProject.Start

            If dauer = 0 And _relEnde > 0 Then

                ' dann sind die Werte initial noch nicht gesetzt worden 
                _startOffsetinDays = CInt(DateDiff(DateInterval.Day, projektStartdate, projektStartdate.AddMonths(_relStart - 1)))
                _dauerInDays = calcDauerIndays(projektStartdate.AddDays(_startOffsetinDays), _relEnde - _relStart + 1, True)


            ElseIf dauer = 0 And _relEnde = 0 Then

                Throw New ArgumentException("Phase kann nicht Dauer = 0 haben ")

            Else
                '  

                If _dauerInDays > 0 And dauer > 0 And awinSettings.propAnpassRess = True Then
                    faktor = dauer / _dauerInDays
                Else
                    faktor = 1
                End If

                _startOffsetinDays = startOffset
                _dauerInDays = dauer

                Dim oldlaenge As Integer = _relEnde - _relStart + 1


                Dim phaseStartdate As Date = Me.getStartDate
                Dim phaseEndDate As Date = Me.getEndDate


                _relStart = getColumnOfDate(phaseStartdate) - projektstartColumn + 1
                _relEnde = getColumnOfDate(phaseEndDate) - projektstartColumn + 1


                If newCalculationNecessary Then

                    Dim newvalues() As Double
                    'Dim notYetDone As Boolean = True

                    dimension = _relEnde - _relStart
                    ReDim newvalues(dimension)

                    If Me.countRoles > 0 Or Me.countCosts > 0 Then

                        ' hier müssen jetzt die Xwerte neu gesetzt werden 
                        Call Me.calcNewXwerte(dimension, faktor)
                        'notYetDone = False

                    End If

                    'If Me.countCosts > 0 And notYetDone Then

                    '    ' hier müssen jetzt die Xwerte neu gesetzt werden 
                    '    Call Me.calcNewXwerte(dimension, 1)

                    'End If

                End If




            End If


        Catch ex As Exception
            ' bei einer Projektvorlage gibt es kein Datum - es sollen aber die Werte für Offset und Dauer übernommen werden

            If dauer = 0 And _relEnde > 0 Then


                ' dann sind die Werte initial noch nicht gesetzt worden 
                _startOffsetinDays = CInt(DateDiff(DateInterval.Day, StartofCalendar, StartofCalendar.AddMonths(_relStart - 1)))
                '_dauerInDays = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(_relStart - 1), _
                '                        StartofCalendar.AddMonths(_relEnde).AddDays(-1)) + 1
                _dauerInDays = calcDauerIndays(projektStartdate.AddDays(_startOffsetinDays), _relEnde - _relStart + 1, True)


            Else
                '  
                _startOffsetinDays = startOffset
                _dauerInDays = dauer

                _relStart = CInt(DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(startOffset)) + 1)
                _relEnde = CInt(DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(startOffset + _dauerInDays - 1)) + 1)


            End If

        End Try


    End Sub

    Public ReadOnly Property dauerInDays As Integer

        Get
            dauerInDays = _dauerInDays
        End Get

    End Property




    Public ReadOnly Property startOffsetinDays As Integer

        Get
            startOffsetinDays = _startOffsetinDays
        End Get


    End Property


    Public Property offset As Integer
        Get
            offset = _offset
        End Get
        Set(value As Integer)
            If _earliestStart = -999 Or _latestStart = -999 Then
                _offset = value
            Else
                If value >= _earliestStart - _relStart And value <= _latestStart - _relStart Then
                    _offset = value
                Else
                    Throw New ApplicationException("Wert für Offset liegt ausserhalb der zugelassenen Grenzen")
                End If
            End If

        End Set
    End Property

    ''' <summary>
    ''' liest/schreibt den Original Name
    ''' gibt den Original Namen einer Phase zurück 
    ''' wenn der leer ist, dann wird der Phasen Name zurück gegeben 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property originalName As String
        Get

            If _originalName = "" Then
                originalName = Me.name
            Else
                originalName = _originalName
            End If

        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                If value.Trim.Length > 0 Then
                    _originalName = value
                End If
            End If
        End Set
    End Property

    ''' <summary>
    ''' gibt die Abkürzung der Phase zurück 
    ''' entweder als Abkürzung der phaseDefinitions, als Abkürzung der missingphaseDefinitions oder der leere String
    ''' Später: alternativeAbbrev
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property shortName As String
        Get
            Dim abbrev As String = ""
            Dim tmpName As String = Me.name


            If PhaseDefinitions.Contains(tmpName) Then
                abbrev = PhaseDefinitions.getAbbrev(tmpName)
            ElseIf missingPhaseDefinitions.Contains(tmpName) Then
                abbrev = missingPhaseDefinitions.getAbbrev(tmpName)
            Else
                abbrev = ""
            End If

            shortName = abbrev

        End Get

        Set(value As String)
            If Not IsNothing(value) Then
                If value.Trim.Length > 0 Then
                    _shortName = value
                End If
            End If
        End Set

    End Property

    ''' <summary>
    ''' gets the penalty value, Read-Only
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getPenaltyValue As Double
        Get
            getPenaltyValue = _penalty.Value
        End Get
    End Property
    Public ReadOnly Property getPenaltyDate As Date
        Get
            getPenaltyDate = _penalty.Key
        End Get
    End Property

    ''' <summary>
    ''' gets the amount of invoice, due at the end of the phase
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getPaymentValue As Double
        Get
            getPaymentValue = _invoice.Key
        End Get
    End Property
    ''' <summary>
    ''' gets the date of payment/cash arrival 
    ''' is termsofpayments days later than end of phase
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getPaymentDate As Date
        Get
            getPaymentDate = getEndDate.AddDays(_invoice.Value)
        End Get
    End Property


    ''' <summary>
    ''' liefert das StartDatum der Phase
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getStartDate As Date
        Get
            getStartDate = Me.parentProject.startDate.AddDays(_startOffsetinDays)
        End Get
    End Property

    ''' <summary>
    ''' liefert das Ende-Datum einer Phase
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getEndDate As Date

        Get
            If _dauerInDays > 0 Then
                getEndDate = Me.parentProject.startDate.AddDays(_startOffsetinDays + _dauerInDays - 1)
            Else
                'Throw New Exception("Dauer muss mindestens 1 Tag sein ...")
                getEndDate = Me.parentProject.startDate.AddDays(_startOffsetinDays)
            End If

        End Get

    End Property

    ''' <summary>
    ''' gibt die individuelle Farbe zurück, also die Einstellung, die verwendet wird 
    ''' wenn es sich nicht um einen categorized namen handelt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property individualColor As Integer
        Get
            individualColor = Me._color
        End Get
    End Property

    ''' <summary>
    ''' schreibt die individuelle Farbe, also die Farbe die verwendet wird, wenn es weder in PhaseDefinitions
    ''' noch in missingPhaseDefinitions einen Eintrag dazu gibt ...
    ''' gibt die Farbe einer Phase zurück; das ist die Farbe der Darstellungsklasse, wenn die Phase zur Liste der
    ''' bekannten Elemente gehört, sonst die AlternativeFare, die ggf beim auslesen z.b. aus MS Project ermittelt wird
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property farbe As Integer
        Get
            Try

                Dim phName As String = elemNameOfElemID(_nameID)

                If Not IsNothing(appearanceDefinitions.getPhaseAppearance(name, appearanceName)) Then
                    farbe = appearanceDefinitions.liste.Item(Me.appearanceName).FGcolor
                Else
                    farbe = _color
                End If

            Catch ex As Exception
                farbe = _color
            End Try

        End Get
        Set(value As Integer)
            If value >= RGB(0, 0, 0) And value <= RGB(255, 255, 255) Then
                _color = value
            End If
        End Set
    End Property


    ' ''' <summary>
    ' ''' setzt die Farbe einer Phase; macht  dann Sinn, wenn die Phase nicht zur 
    ' ''' Liste der bekannten/missing Phasen gehört 
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <remarks></remarks>
    'Public WriteOnly Property setFarbe As Long
    '    Set(value As Long)

    '        If value >= RGB(0, 0, 0) And value <= RGB(255, 255, 255) Then
    '            _alternativeColor = value
    '        Else
    '            ' unverändert lassen - wird ja auch im New initial gesetzt 
    '        End If

    '    End Set
    'End Property


    ''' <summary>
    ''' ist die Anzahl in Tagen, die die Phase vor ihrem aktuellen Startdatum beginnen kann
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property earliestStart As Integer
        Get
            earliestStart = _earliestStart
        End Get
        Set(value As Integer)
            If value <= 0 Then
                ' tk 17.11.15: hier muss noch eine Konsistenzprüfung rein ...
                _earliestStart = value

            ElseIf value = -999 Then ' die undefiniert Bedingung
                _earliestStart = value
            Else
                Throw New ApplicationException("Wert für Earliest Start kann nicht größer Null sein")
            End If

        End Set
    End Property

    ''' <summary>
    ''' ist die Anzahl in Tagen, die die Phase nach ihrem aktuellen Startdatum beginnen kann
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property latestStart As Integer
        Get
            latestStart = _latestStart
        End Get
        Set(value As Integer)
            If value >= 0 Then
                ' tk 17.11.15 hier muss noch eine Konsistenzprüfung rein ... 
                _latestStart = value

            ElseIf value = -999 Then ' die undefiniert Bedingung
                _latestStart = value
            Else
                Throw New ApplicationException("Wert für Latest Start kann nicht negativ sein")
            End If

        End Set
    End Property

    'Public Property minDauer As Integer
    '    Get
    '        minDauer = _minDauer
    '    End Get
    '    Set(value As Integer)
    '        If value >= 1 Then
    '            If _maxDauer <> -999 Then
    '                If value <= _maxDauer Then
    '                    _minDauer = value
    '                Else
    '                    Throw New ApplicationException("Mindest-Dauer kann nicht größer als Max Dauer sein")
    '                End If
    '            Else
    '                _minDauer = value
    '            End If
    '        Else
    '            Throw New ApplicationException("Mindest-Dauer kann nicht negativ oder Null sein")
    '        End If

    '    End Set
    'End Property

    'Public Property maxDauer As Integer
    '    Get
    '        maxDauer = _maxDauer
    '    End Get
    '    Set(value As Integer)
    '        If value >= 1 Then
    '            If _minDauer <> -999 Then
    '                If value >= _minDauer Then
    '                    _maxDauer = value
    '                Else
    '                    Throw New ApplicationException("Maximal-Dauer kann nicht kleiner als Min Dauer sein")
    '                End If
    '            Else
    '                _maxDauer = value
    '            End If
    '        Else
    '            Throw New ApplicationException("Maximal-Dauer kann nicht negativ oder Null sein")
    '        End If

    '    End Set
    'End Property


    Public ReadOnly Property relStart As Integer
        Get

            Dim isVorlage As Boolean
            Dim tmpValue As Integer
            'Dim checkValue As Integer = _relStart + _Offset

            Try
                isVorlage = (Me.parentProject.projectType = ptPRPFType.projectTemplate)

                ' tk 30.9.18 , durch obiges Statement ersetzt
                'If Me.parentProject Is Nothing Then
                '    isVorlage = True
                'Else
                '    isVorlage = False
                'End If
            Catch ex As Exception
                isVorlage = True
            End Try

            If isVorlage Then
                tmpValue = getColumnOfDate(StartofCalendar.AddDays(Me.startOffsetinDays))
            Else
                tmpValue = getColumnOfDate(Me.parentProject.startDate.AddDays(Me.startOffsetinDays)) - Me.parentProject.Start + 1
            End If

            'If checkValue <> tmpValue Then 
            '    Call MsgBox("oops in relStart")
            'End If

            ' kann später eliminiert werden - vorläufig bleibt das zur Sicherheit noch drin ... 
            _relStart = tmpValue

            ' Return Wert
            relStart = tmpValue




        End Get


    End Property



    Public ReadOnly Property relEnde As Integer
        Get

            Dim isVorlage As Boolean
            Dim tmpValue As Integer

            Try
                isVorlage = (Me.parentProject.projectType = ptPRPFType.projectTemplate)

                ' tk 30.9.18
                'If Me.parentProject Is Nothing Then
                '    isVorlage = True
                'Else
                '    isVorlage = False
                'End If
            Catch ex As Exception
                isVorlage = True
            End Try

            If isVorlage Then
                tmpValue = getColumnOfDate(StartofCalendar.AddDays(Me.startOffsetinDays))
            Else
                tmpValue = getColumnOfDate(Me.parentProject.startDate.AddDays(Me.startOffsetinDays + Me.dauerInDays - 1)) - Me.parentProject.Start + 1
            End If

            ' kann später eliminiert werden - vorläufig bleibt das zur Sicherheit noch drin ... 
            _relEnde = tmpValue

            ' Return Wert
            relEnde = tmpValue

        End Get

    End Property

    ''' <summary>
    ''' setzt bzw liest die NamensID einer Phase; die NamensID setzt sich zusammen aus 
    ''' dem Kennzeichen Phase/Meilenstein 0/1, dem eigentlichen Namen der Phase und der laufenden Nummer. 
    ''' Getrennt sind die Elemente durch das Zeichen § 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property nameID As String
        Get
            nameID = _nameID
        End Get
        Set(value As String)
            Dim tmpstr() As String
            tmpstr = value.Split(New Char() {CChar("§")}, 3)
            If Len(value) > 0 Then
                If value.StartsWith("0§") And tmpstr.Length >= 2 Then
                    _nameID = value
                Else
                    Throw New ApplicationException("unzulässige Namens-ID: " & value)
                End If

            Else
                Throw New ApplicationException("Name darf nicht leer sein ...")
            End If

        End Set
    End Property

    ''' <summary>
    ''' liest den Namensteil der NamensID 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property name As String
        Get
            name = elemNameOfElemID(_nameID)
        End Get
    End Property

    ''' <summary>
    ''' berechnet die Shape Koordinaten dieser Phase 
    ''' </summary>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Public Sub calculatePhaseShapeCoord(ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)

        Try

            Dim projektStartdate As Date = Me.parentProject.startDate
            Dim tfzeile As Integer = Me.parentProject.tfZeile
            Dim startpunkt As Integer = CInt(DateDiff(DateInterval.Day, StartofCalendar, projektStartdate))

            Dim faktor As Double = 0.4

            If startpunkt < 0 Then
                Throw New Exception("calculate Shape Coord für Phase: Projektstart liegt vor Start of Calendar ...")
            End If

            Dim phasenStart As Integer = startpunkt + Me.startOffsetinDays
            Dim phasenDauer As Integer = Me.dauerInDays



            If tfzeile > 1 And phasenStart >= 1 And phasenDauer > 0 Then


                'top = topOfMagicBoard + (tfzeile - 1) * boxHeight + 0.5 * (0.8 - 0.23) * boxHeight
                top = topOfMagicBoard + (tfzeile - 1) * boxHeight + 0.5 * (0.8 - faktor) * boxHeight
                left = (phasenStart / 365) * boxWidth * 12
                width = ((phasenDauer) / 365) * boxWidth * 12
                'height = 0.23 * boxHeight
                height = faktor * boxHeight

            Else
                Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.nameID)
            End If

        Catch ex As Exception
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.nameID)
        End Try


    End Sub

    'Public Sub calculateLineCoord(ByVal zeile As Integer, ByVal nummer As Integer, ByVal gesamtZahl As Integer, _
    '                              ByRef top1 As Double, ByRef left1 As Double, ByRef top2 As Double, ByRef left2 As Double, ByVal linienDicke As Double)

    '    Try

    '        Dim projektStartdate As Date = Me.Parent.startDate

    '        Dim korrPosition As Double = nummer / gesamtZahl
    '        Dim faktor As Double = linienDicke / boxHeight
    '        Dim startpunkt As Integer = CInt(DateDiff(DateInterval.Day, StartofCalendar, projektStartdate))

    '        If startpunkt < 0 Then
    '            Throw New Exception("calculate Line Coord: Projektstart liegt vor Start of Calendar ...")
    '        End If

    '        Dim phasenStart As Integer = startpunkt + Me.startOffsetinDays
    '        Dim phasenDauer As Integer = Me.dauerInDays

    '        ' absolute Setzung - dadurch wird verhindert, daß die Linien sehr schmal gezeichnet werden ... 
    '        ' es soll immer gleich groß gezeichnet werden - einfach überschreiben - das ist rvtl besser;
    '        ' das muss einfach noch herausgefunden werden 
    '        gesamtZahl = 1
    '        nummer = 1


    '        If gesamtZahl <= 0 Then
    '            Throw New ArgumentException("unzulässige Gesamtzahl" & gesamtZahl)
    '        End If

    '        ' korrigiere, aber breche nicht ab wenn die Nummer der Line größer als die Gesamtzahl ist ... 
    '        If nummer > gesamtZahl Then
    '            nummer = gesamtZahl
    '        End If

    '        ' ausrechnen des Korrekturfaktors

    '        korrPosition = nummer / (gesamtZahl + 1)


    '        If phasenStart >= 0 And phasenDauer > 0 Then

    '            ' das folgende ist mühsam ausprobiert - um die Linien in unterschiedicher Stärke in der Projekt Form zu platzieren - möglichst auch jeweils mittig
    '            If gesamtZahl <= 3 Then
    '                top1 = topOfMagicBoard + (zeile - 0.95) * boxHeight + korrPosition * boxHeight - linienDicke / 2
    '            Else
    '                top1 = topOfMagicBoard + (zeile - 1.06) * boxHeight + korrPosition * boxHeight - linienDicke / 2
    '            End If

    '            top2 = top1

    '            left1 = (phasenStart / 365) * boxWidth * 12
    '            left2 = ((phasenStart + phasenDauer) / 365) * boxWidth * 12

    '        Else
    '            Throw New ArgumentException("es kann keine Line berechnet werden für : " & Me.name)
    '        End If

    '    Catch ex As Exception
    '        Throw New ArgumentException("es kann keine Line berechnet werden für : " & Me.name)
    '    End Try


    ''' <summary>
    ''' gibt die Rollen Instanz der Rolle zurück, die den Namen roleName hat; wenn teamID = 0, dann egal in welchem Team
    ''' wenn teamID angegeben ist, dann nur die Rolle in der Eigenschaft als Team-MEmber
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRole(ByVal roleName As String, Optional ByVal teamID As Integer = -1) As clsRolle

        Get
            Dim returnValue As clsRolle = Nothing
            Dim ix As Integer = 0
            Dim found As Boolean = False

            If teamID = 0 Then
                ' teamID ist bei der suche nicht relevant
                While Not found And ix <= _allRoles.Count - 1
                    If _allRoles.Item(ix).name = roleName Then
                        found = True
                        returnValue = _allRoles.Item(ix)
                    Else
                        ix = ix + 1
                    End If
                End While
            Else
                While Not found And ix <= _allRoles.Count - 1
                    If _allRoles.Item(ix).name = roleName And _allRoles.Item(ix).teamID = teamID Then
                        found = True
                        returnValue = _allRoles.Item(ix)
                    Else
                        ix = ix + 1
                    End If
                End While
            End If

            getRole = returnValue

        End Get

    End Property


    ''' <summary>
    ''' liefert die Namen und Bedarfs-Summen aller Rollen, die in der Phase referenziert werden ...
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleNamesAndValues() As SortedList(Of String, Double)
        Get
            Dim zwResult As New SortedList(Of String, Double)

            For i As Integer = 1 To _allRoles.Count
                Dim tmpRole As clsRolle = _allRoles.Item(i - 1)

                If Not zwResult.ContainsKey(tmpRole.name) Then
                    zwResult.Add(tmpRole.name, tmpRole.summe)
                Else
                    zwResult.Item(tmpRole.name) = zwResult.Item(tmpRole.name) + tmpRole.summe
                End If
            Next

            getRoleNamesAndValues = zwResult

        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen und Bedarfs-Summen aller Rollen, die in der Phase referenziert werden ...
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostNamesAndValues() As SortedList(Of String, Double)
        Get
            Dim zwResult As New SortedList(Of String, Double)

            For i As Integer = 1 To _allCosts.Count
                Dim tmpCost As clsKostenart = _allCosts.Item(i - 1)

                If Not zwResult.ContainsKey(tmpCost.name) Then
                    zwResult.Add(tmpCost.name, tmpCost.summe)
                Else
                    zwResult.Item(tmpCost.name) = zwResult.Item(tmpCost.name) + tmpCost.summe
                End If
            Next

            getCostNamesAndValues = zwResult

        End Get
    End Property

    ''' <summary>
    ''' checks whether or not phase has roles with resourcen-needsand role  has already left company or is not yet part of the company 
    ''' </summary>
    ''' <returns></returns>
    Public Function hasRolesWithInvalidNeeds() As Collection
        Dim allInvalidNames As New Collection
        Try
            Dim startColumn As Integer = parentProject.Start + relStart - 1
            Dim endColumn As Integer = parentProject.Start + relEnde - 1

            For Each role As clsRolle In _allRoles
                If isRoleWithInvalidNeeds(role, startColumn, endColumn) Then
                    Dim roleName As String = role.name
                    If Not allInvalidNames.Contains(roleName) Then
                        allInvalidNames.Add(roleName, roleName)
                    End If
                End If
            Next
        Catch ex As Exception
            If awinSettings.visboDebug Then
                Call MsgBox("Érror-Code 9973276-0")
            End If
        End Try


        hasRolesWithInvalidNeeds = allInvalidNames
    End Function

    ''' <summary>
    ''' returns whether or not this role has resource needs where role ist not yet at the company or not any more. 
    ''' </summary>
    ''' <param name="tmprole"></param>
    ''' <returns></returns>
    Public Function isRoleWithInvalidNeeds(ByVal tmprole As clsRolle, ByVal startColumn As Integer, ByVal endColumn As Integer) As Boolean

        Dim tmpResult As Boolean = False
        Try
            Dim currentRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(tmprole.uid)
            ' nur bei Personen-Rollen oder Team-Roles relevant und zu prüfen  

            If Not currentRole.isCombinedRole Or currentRole.isTeam Then

                Dim weiterMachen As Boolean = True

                If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or
                        myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then

                    Dim teamID As Integer = -1
                    Dim myTopRoleID As Integer = RoleDefinitions.getRoleDefByIDKennung(myCustomUserRole.specifics, teamID).UID

                    If Not RoleDefinitions.hasAnyChildParentRelationsship(currentRole.UID, myTopRoleID) Then
                        weiterMachen = False
                    End If
                End If

                If weiterMachen Then
                    Dim startOfEmployee As Integer = getColumnOfDate(currentRole.entryDate)
                    Dim leaveOFEmployee As Integer = getColumnOfDate(currentRole.exitDate)

                    ' wann ist es kritisch 
                    If startOfEmployee > startColumn Or leaveOFEmployee <= endColumn Then
                        If startOfEmployee > endColumn Or leaveOFEmployee <= startColumn Then
                            ' nur dann ungültig, wenn es auch Werte > 0 gibt  
                            tmpResult = tmprole.Xwerte.Sum > 0

                        Else
                            ' hier ist gesichert, dass StartOfEmployee <= endColumn ist ..
                            For i As Integer = startColumn To startOfEmployee
                                If tmprole.Xwerte(i - startColumn) > 0 Then
                                    tmpResult = True
                                    Exit For
                                End If
                            Next

                            If Not tmpResult And leaveOFEmployee <= endColumn Then
                                For i As Integer = leaveOFEmployee To endColumn
                                    If tmprole.Xwerte(i - startColumn) > 0 Then
                                        tmpResult = True
                                        Exit For
                                    End If
                                Next
                            End If

                        End If
                    End If
                End If

            End If

        Catch ex As Exception
            If awinSettings.visboDebug Then
                Call MsgBox("Érror-Code 9973276-1")
            End If
        End Try



        isRoleWithInvalidNeeds = tmpResult
    End Function

    ''' <summary>
    ''' erstellt eine neue Rolle, weist der Rolle monatliche Ressourcenbedarfe zu, deren Summe dem Wert der Variable summe entspricht  
    ''' der RoleName muss in Roledefinitions existieren , sonst gibt es eine Fehlermeldung 
    ''' </summary>
    ''' <param name="roleNameID"></param>
    ''' <param name="summe"></param>
    ''' <param name="addToExisting"></param>
    Public Sub AddRole(ByVal roleNameID As String, ByVal summe As Double, ByVal addToExisting As Boolean)

        Dim rSum As Double()
        ReDim rSum(0)
        rSum(0) = summe

        Dim teamID As Integer = -1
        Dim roleID As Integer = RoleDefinitions.parseRoleNameID(roleNameID, teamID)

        Dim tmpRole As clsRolle = Me.getRoleByRoleNameID(roleNameID)
        Dim xWerte As Double() = Me.berechneBedarfeNew(Me.getStartDate, Me.getEndDate, rSum, 1.0)

        If IsNothing(tmpRole) Then
            ' die Rolle hat bisher noch nicht existiert ...
            Dim dimension As Integer = Me.relEnde - Me.relStart
            tmpRole = New clsRolle(dimension)

            With tmpRole
                .uid = roleID
                .teamID = teamID
                .Xwerte = xWerte
            End With

            ' jetzt muss die Rolle ergänzt werden 
            _allRoles.Add(tmpRole)

        Else
            ' die Rolle hat bereits existiert 
            If addToExisting Then
                If tmpRole.Xwerte.Length = xWerte.Length Then
                    ' hier dann aufsummieren 
                    Dim oldXwerte As Double() = tmpRole.Xwerte
                    For i As Integer = 0 To oldXwerte.Length - 1
                        xWerte(i) = xWerte(i) + oldXwerte(i)
                    Next

                Else
                    ' darf eigentlich nicht sein 
                    ' Test: 
                    'Call MsgBox("Fehler in Rollen-Zuordnung")
                    ' es wird dann einfach gar nichts gemacht 
                End If
            Else
                ' nichts weiter tun 
            End If

            tmpRole.Xwerte() = xWerte
        End If


        ' jetzt müssen die sortierten Listen im Projekt entsprechend aktualisiert werden 
        Try
            Me.parentProject.rcLists.addRP(tmpRole.uid, Me.nameID, teamID:=teamID)
        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' fügt der Phase die Rollen und Kosten hinzu, wie angegeben
    ''' </summary>
    ''' <param name="roleNames">die Namen der Rollen</param>
    ''' <param name="roleValues">die Werte der Rollen</param>
    ''' <param name="costNames">die Namen der Kostenarten</param>
    ''' <param name="costValues">die Werte der Kostenarten</param>
    ''' <param name="prozentSatz">wenn nur ein bestimmter Prozentsatz auf die Phase verteilt werden sollen; by Default 1</param>
    Public Sub addCostsAndRoles(ByVal roleNames() As String, ByVal roleValues() As Double,
                                ByVal costNames() As String, ByVal costValues() As Double,
                                Optional ByVal prozentSatz As Double = 1.0,
                                Optional ByVal roleNamesAreIds As Boolean = False,
                                Optional ByVal createCostsRolesAnyhow As Boolean = False)

        Dim anzRoles As Integer
        Dim anzCosts As Integer
        Dim teamID As Integer = -1
        Dim roleID As Integer = 0

        Dim tmpRCvalue As Double = 0.0
        Dim tmpRCnameID As String

        If IsNothing(roleNames) Then
            anzRoles = 0
        Else
            anzRoles = roleNames.Length
        End If

        If IsNothing(costNames) Then
            anzCosts = 0
        Else
            anzCosts = costNames.Length
        End If

        For r = 0 To anzRoles - 1
            tmpRCvalue = prozentSatz * roleValues(r)
            tmpRCnameID = RoleDefinitions.bestimmeRoleNameID(roleNames(r), "")
            If roleNamesAreIds Then
                ' dann ist es schon in der Form RoleId;TeamID bzw RoleID
                tmpRCnameID = roleNames(r)
            Else
                tmpRCnameID = RoleDefinitions.bestimmeRoleNameID(roleNames(r), "")
            End If

            If tmpRCvalue > 0 Then
                ' whenexisting sollte immer dazu addiert werden ... ! 
                'Me.addCostRole(tmpRCnameID, tmpRCvalue, True, False)
                Me.addCostRole(tmpRCnameID, tmpRCvalue, True, False)
            Else
                If createCostsRolesAnyhow Then
                    Me.addCostRole(tmpRCnameID, tmpRCvalue, True, False)
                End If
            End If

        Next

        For c = 0 To anzCosts - 1
            tmpRCvalue = prozentSatz * costValues(c)
            tmpRCnameID = costNames(c)
            If tmpRCvalue > 0 Then
                ' wenn existing sollte immer dazu addiert werden 
                'Me.addCostRole(tmpRCnameID, tmpRCvalue, False, False)
                Me.addCostRole(tmpRCnameID, tmpRCvalue, False, False)
            Else
                If createCostsRolesAnyhow Then
                    Me.addCostRole(tmpRCnameID, tmpRCvalue, True, False)
                End If
            End If
        Next

    End Sub

    ''' <summary>
    ''' fügt der aktuellen Phase eine Rolle bzw. Kostenart hinzu
    ''' wenn es eine Rolle ist, so ist sie in der form rcNameID roleUID;teamID bzw roleUID anzugeben
    ''' </summary>
    ''' <param name="rcNameID"></param>
    ''' <param name="summe"></param>
    ''' <param name="isrole"></param>
    ''' <param name="addWhenExisting"></param>
    Public Sub addCostRole(ByVal rcNameID As String, ByVal summe As Double,
                              ByVal isrole As Boolean,
                              ByVal addWhenExisting As Boolean)


        If isrole Then
            ' eine Rolle wird hinzugefügt 
            Call Me.AddRole(rcNameID, summe, addWhenExisting)

        Else
            ' eine Kostenart wird hinzugefügt
            Call Me.AddCost(rcNameID, summe, addWhenExisting)
        End If


    End Sub

    ''' <summary>
    ''' addRole fügt die Rollen Instanz hinzu, wenn sie nicht schon existiert
    ''' wenn sie schon existiert, dann werden die Werte zu den schon existierenden Werten addiert ...
    ''' </summary>
    ''' <param name="role"></param>
    ''' <remarks></remarks>
    Public Sub addRole(ByVal role As clsRolle)

        'sollte nach dem 8.7.16 aktiviert werden 
        'ebenso für addCost, mehrere Rollen/Kosten des gleichen NAmens sollen aufsummiert werden 
        Dim roleName As String = role.name
        Dim teamID As Integer = role.teamID

        Dim returnValue As clsRolle = Nothing
        Dim ix As Integer = 0
        Dim found As Boolean = False
        Dim oldXWerte() As Double
        Dim newXwerte() As Double

        While Not found And ix <= _allRoles.Count - 1
            If _allRoles.Item(ix).name = roleName And _allRoles.Item(ix).teamID = teamID Then
                found = True
            Else
                ix = ix + 1
            End If
        End While

        If found Then
            oldXWerte = _allRoles.Item(ix).Xwerte()
            newXwerte = role.Xwerte
            If oldXWerte.Length = newXwerte.Length Then
                ' hier dann aufsummieren 
                For i As Integer = 0 To oldXWerte.Length - 1
                    newXwerte(i) = newXwerte(i) + oldXWerte(i)
                Next

                _allRoles.Item(ix).Xwerte() = newXwerte

            Else
                ' darf eigentlich nicht sein 
                ' Test: 
                Call MsgBox("Fehler in Rollen-Zuordnung")
                ' es wird dann einfach gar nichts gemacht 
            End If
        Else
            _allRoles.Add(role)
        End If

        ' jetzt müssen die sortierten Listen im Projekt entsprechend aktualisiert werden 
        Try
            Me.parentProject.rcLists.addRP(role.uid, Me.nameID, teamID)
        Catch ex As Exception

        End Try


        ' '' Code vor dem 8.7.16
        ''If Not _allRoles.Contains(role) Then
        ''    _allRoles.Add(role)
        ''Else
        ''    'Call logfileSchreiben("Fehler: Rolle '" & role.name & "' ist bereits in der Phase '" & Me.name & "' enthalten", "", anzFehler)
        ''End If


    End Sub

    ''' <summary>
    ''' entfernt alle Rollen-Instanzen mit Rollen-Name aus der Phase
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <remarks></remarks>
    Public Sub removeRoleByName(ByVal roleName As String, Optional ByVal teamID As Integer = -1)

        Dim toDoList As New List(Of clsRolle)

        For i As Integer = 1 To _allRoles.Count
            Dim tmpRole As clsRolle = _allRoles.Item(i - 1)
            If tmpRole.name = roleName And tmpRole.teamID = teamID Then
                toDoList.Add(tmpRole)
            End If
        Next

        For Each tmpRole As clsRolle In toDoList
            _allRoles.Remove(tmpRole)
            ' Änderung tk 20.09.16
            ' jetzt müssen die sortierten Listen im Projekt entsprechend aktualisiert werden 
            Me.parentProject.rcLists.removeRP(tmpRole.uid, Me.nameID, teamID, False)
        Next


    End Sub

    ''' <summary>
    ''' entfernt alle Rollen-Instanzen mit RolleName-ID "roleuid;teamUid" aus der Phase
    ''' </summary>
    ''' <param name="roleNameID"></param>
    Public Sub removeRoleByNameID(ByVal roleNameID As String)

        Dim toDoList As New List(Of clsRolle)

        For i As Integer = 1 To _allRoles.Count
            Dim tmpRole As clsRolle = _allRoles.Item(i - 1)
            Dim IdStr As String = RoleDefinitions.bestimmeRoleNameID(tmpRole.uid, tmpRole.teamID)
            If IdStr = roleNameID Then
                toDoList.Add(tmpRole)
            End If
        Next

        For Each tmpRole As clsRolle In toDoList
            _allRoles.Remove(tmpRole)
            Me.parentProject.rcLists.removeRP(tmpRole.uid, Me.nameID, tmpRole.teamID, False)
        Next

    End Sub

    ''' <summary>
    ''' es wird überprüft, ob der Meilenstein-Name schon existiert 
    ''' wenn er bereits existiert, wird eine ArgumentException geworfen  
    ''' </summary>
    ''' <param name="milestone"></param>
    ''' <remarks></remarks>
    Public Sub addMilestone(ByVal milestone As clsMeilenstein,
                            Optional ByVal origName As String = "")


        Dim anzElements As Integer = _allMilestones.Count - 1
        Dim ix As Integer = 0
        Dim found As Boolean = False

        Dim elemName As String = elemNameOfElemID(milestone.nameID)

        ' wenn der Origname gesetzt werden soll ...
        If origName <> "" Then
            If milestone.originalName <> origName Then
                milestone.originalName = origName
            End If
        End If

        Do While ix <= anzElements And Not found
            If _allMilestones.Item(ix).nameID = milestone.nameID Then
                found = True
            Else
                ix = ix + 1
            End If
        Loop

        If found Then
            Throw New ArgumentException("Meilenstein existiert bereits in dieser Phase!" & milestone.nameID)
        Else
            _allMilestones.Add(milestone)
        End If

        ' jetzt muss der Meilenstein in die Projekt-Hierarchie aufgenommen werden , 
        ' aber nur, wenn die Phase bereits in der Projekt-Hierarchie vorhanden ist ... 
        Dim elemID As String = milestone.nameID
        Dim currentElementNode As New clsHierarchyNode
        Dim hproj As New clsProjekt, vproj As New clsProjektvorlage
        Dim parentIsVorlage As Boolean
        Dim milestoneIndex As Integer = _allMilestones.Count
        Dim phaseID As String = Me.nameID
        Dim ok As Boolean = False

        If Not istElemID(elemID) Then
            elemID = vproj.hierarchy.findUniqueElemKey(elemName, True)
        End If

        If IsNothing(Me.parentProject) Then
            parentIsVorlage = True
            vproj = Me.VorlagenParent
            If vproj.hierarchy.containsKey(phaseID) Then
                ' Phase ist bereits in der Projekt-Hierarchie eingetragen
                ok = True
            End If
        Else
            parentIsVorlage = False
            hproj = Me.parentProject
            If hproj.hierarchy.containsKey(phaseID) Then
                ' Phase ist bereits in der Projekt-Hierarchie eingetragen
                ok = True
            End If
        End If

        If ok Then

            With currentElementNode

                .elemName = elemName

                ' '' Änderung tk 29.5.16 : Origname ist nicht mehr Bestandteil von hierarchyNode ... 
                ''If origName = "" Then
                ''    .origName = .elemName
                ''Else
                ''    .origName = origName
                ''End If

                .indexOfElem = milestoneIndex
                .parentNodeKey = phaseID

            End With

            If parentIsVorlage Then
                vproj.hierarchy.addNode(currentElementNode, elemID)
            Else
                hproj.hierarchy.addNode(currentElementNode, elemID)
            End If


        End If


    End Sub

    ''' <summary>
    ''' löscht den Meilenstein an Position index; Index kann Werte 1 .. Anzahl Meilensteine haben 
    ''' wenn checkname ungleich "" ist , so wird der Meilenstein nur dann gelöscht, wenn die NameID mit checkname übereinstimmt  
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="checkID"></param>
    ''' <remarks></remarks>
    Public Sub removeMilestoneAt(ByVal index As Integer, Optional ByVal checkID As String = "")
        Dim ok As Boolean = True

        If index >= 0 And index <= _allMilestones.Count - 1 Then
            If checkID <> "" Then
                If _allMilestones.ElementAt(index).nameID = checkID Then
                    ok = True
                Else
                    ok = False
                End If
            End If
        Else
            ok = False
        End If


        If ok Then
            _allMilestones.RemoveAt(index)
        End If

    End Sub


    Public ReadOnly Property rollenListe() As List(Of clsRolle)

        Get
            rollenListe = _allRoles
        End Get

    End Property

    Public ReadOnly Property meilensteinListe() As List(Of clsMeilenstein)

        Get
            meilensteinListe = _allMilestones
        End Get

    End Property

    Public ReadOnly Property kostenListe() As List(Of clsKostenart)

        Get
            kostenListe = _allCosts
        End Get

    End Property


    Public ReadOnly Property countRoles() As Integer

        Get
            countRoles = _allRoles.Count
        End Get

    End Property

    Public ReadOnly Property countMilestones() As Integer

        Get
            countMilestones = _allMilestones.Count
        End Get

    End Property



    ''' <summary>
    ''' Property, die die aktuelle Phase in die newphase kopiert.
    ''' mapping = true: es werden keine Rollen, Kosten und Meilensteine übernommen
    '''                 auch die nameID wird nicht übernommen sondern hinterher neu berechnet
    ''' mapping = false: alles wird übernommen
    ''' </summary>
    ''' <param name="newphase"></param>
    ''' <param name="withoutNameID">default false: kopiert auch die NameID</param>
    ''' <param name="withoutMS">default false: kopiert inkl Meilensteine</param>
    ''' <param name="withoutRolesCosts">default false: kopiert inkl Rollen und Kosten</param>
    ''' <param name="withoutDeliverables">default false: kopiert inkl Deliverables</param>
    ''' <param name="withoutBewertungen">default false: kopiert also inkl Bewertungen</param>
    ''' <remarks></remarks>
    Public Sub copyTo(ByRef newphase As clsPhase,
                      Optional ByVal withoutNameID As Boolean = False,
                      Optional ByVal withoutMS As Boolean = False,
                      Optional ByVal withoutRolesCosts As Boolean = False,
                      Optional ByVal withoutDeliverables As Boolean = False,
                      Optional ByVal withoutBewertungen As Boolean = False)

        Dim r As Integer, k As Integer
        Dim newrole As clsRolle
        Dim newcost As clsKostenart
        Dim newresult As clsMeilenstein
        ' Dimension ist die Länge des Arrays , der kopiert werden soll; 
        ' mit der eingeführten Unschärfe ist nicht mehr gewährleistet, 
        ' daß relende-relstart die tatsächliche Dimension des Arrays wiedergibt 
        Dim dimension As Integer

        With newphase

            ' tk 25.11.19 , das Auskommentierte führte zu Fehlern ...
            ' insbesondere bei appearance und farbe 

            ' korrekt 25.11.19 
            .earliestStart = earliestStart
            .latestStart = latestStart
            .offset = offset

            ' eindeutiger Name muss bei Mapping neu zusammengesetzt werden
            ' wird also bei Mapping nicht übernommen
            If Not withoutNameID Then
                .nameID = nameID
            End If


            ' sonstigen Elemente übernehmen 
            .shortName = shortName
            .originalName = originalName


            .appearanceName = appearanceName
            .farbe = farbe
            .verantwortlich = verantwortlich
            .percentDone = percentDone

            ' tk 2.6.20
            .invoice = _invoice
            .penalty = _penalty

            ' tk 1.6.2020 das wird vor dem Übertragen der Rollen gemacht 
            ' bis 1.6 war das nach if Not WithoutRolesCosts ...
            .changeStartandDauer(Me._startOffsetinDays, Me._dauerInDays)

            ' Rollen und kosten werden bei Mapping nicht übernommen
            If Not withoutRolesCosts Then

                For r = 1 To Me.countRoles
                    'newrole = New clsRolle(relEnde - relStart)

                    dimension = Me.getRole(r).getDimension
                    newrole = New clsRolle(dimension)
                    Me.getRole(r).CopyTo(newrole)
                    .addRole(newrole)
                Next r


                For k = 1 To Me.countCosts
                    'newcost = New clsKostenart(relEnde - relStart)

                    dimension = Me.getCost(k).getDimension
                    newcost = New clsKostenart(dimension)
                    Me.getCost(k).CopyTo(newcost)
                    .AddCost(newcost)
                Next k

            End If


            ' Änderung 16.1.2014: zuerst die Rollen und Kosten übertragen, dann die relStart und RelEnde, dann die Results
            ' die evtl. enstehende Inkonsistenz zwischen Längen der Arrays der Rollen/Kostenarten und dem neuen relende/relstart wird in Kauf genommen 
            ' und nur korrigiert , wenn explizit gewünscht (Parameter awinsettings.autoCorrectBedarfe = true 

            '.changeStartandDauer(Me._startOffsetinDays, Me._dauerInDays)

            ' Meilensteine werden bei Mapping nicht übernommen
            If Not withoutMS Then

                For r = 1 To Me._allMilestones.Count
                    newresult = New clsMeilenstein(parent:=newphase)
                    Me.getMilestone(r).copyTo(newresult)

                    Try
                        .addMilestone(newresult)
                    Catch ex As Exception

                    End Try

                Next
            End If


            ' jetzt noch die evtl vorhandenen Bewertungen kopieren 
            If Not withoutBewertungen Then
                For b As Integer = 1 To Me._bewertungen.Count
                    Dim newb As New clsBewertung
                    Me.getBewertung(b).copyto(newb)
                    Try
                        .addBewertung(newb)
                    Catch ex As Exception

                    End Try

                Next
            End If


            If Not withoutDeliverables Then
                ' jetzt noch die Deliverables kopieren ... 
                For i = 1 To Me.countDeliverables
                    Dim deli As String = Me.getDeliverable(i)
                    .addDeliverable(deli)
                Next

            End If



        End With

    End Sub
    ''' <summary>
    ''' kopiert Phase mit ihren Ressourcen- und Kostenbedarfen in eine neue Phase
    ''' wenn newPhaseNameID ungleich "", dann wird als neue PhaseNameID verwendet; ist für den modularen Aufbau von Projekten wichtig
    ''' correctfactor streckt / staucht die Dauer des Projektes und passt, wenn awinsettings.propanpassRess = true auch die Ressoucen- und Kostenbedarfe proportional an 
    ''' wenn zielrenditeFaktor angegeben ist, dann wird die Länge gemäßss corrFactor angepasst, nicht aber die Ressourcenbedarfe. Die werden über den ZielrenditeFaktor berechnet
    ''' </summary>
    ''' <param name="newphase"></param>
    ''' <param name="corrFactor"></param>
    ''' <param name="newPhaseNameID"></param>
    ''' <param name="zielrenditeFaktor">wenn Nothing angegeben wird: die Ressourcen- und Kostenbedarfe werden über corrFactor und propanpassRess bestimmt
    ''' wenn ein Wert angegeben ist, dann werden die alten Ressourcen- und Kosten-Summen mit diesem Wert modifiziert; das sichert eine vorgegebene Rendite </param>
    ''' <remarks></remarks>
    Public Sub korrCopyTo(ByRef newphase As clsPhase, ByVal corrFactor As Double, ByVal newPhaseNameID As String,
                          Optional ByVal zielrenditeFaktor As Double = -99999.0)
        Dim r As Integer, k As Integer
        Dim newrole As clsRolle, oldrole As clsRolle
        Dim newcost As clsKostenart, oldcost As clsKostenart
        Dim newresult As clsMeilenstein
        ' Dimension ist die Länge des Arrays , der kopiert werden soll; 
        ' mit der eingeführten Unschärfe ist nicht mehr gewährleistet, 
        ' daß relende-relstart die tatsächliche Dimension des Arrays wiedergibt 
        Dim dimension As Integer
        Dim hname As String
        Dim newXwerte() As Double
        'Dim h1wert As Double
        'Dim h2wert As Double

        With newphase
            '.minDauer = Me._minDauer
            '.maxDauer = Me._maxDauer
            .earliestStart = Me._earliestStart
            .latestStart = Me._latestStart
            .offset = Me._offset

            If newPhaseNameID = "" Then
                .nameID = _nameID
            Else
                .nameID = newPhaseNameID
            End If

            ' ergänzt am 25.11.19 
            ' sonstigen Elemente übernehmen 
            .shortName = shortName
            .originalName = originalName


            .appearanceName = appearanceName
            .farbe = farbe
            .verantwortlich = verantwortlich
            .percentDone = percentDone
            ' Ende ergänzt am 25.11 

            .changeStartandDauer(CInt(Me._startOffsetinDays * corrFactor), CInt(Me._dauerInDays * corrFactor))

            For r = 1 To Me.countRoles
                Try

                    oldrole = Me.getRole(r)
                    dimension = newphase.relEnde - newphase.relStart
                    newrole = New clsRolle(dimension)
                    ReDim newXwerte(dimension)
                    hname = oldrole.name

                    If zielrenditeFaktor = -99999.0 Then
                        ' undefiniert, deswegen corrfactor nehmen 
                        If awinSettings.propAnpassRess Then
                            Call berechneBedarfe(newphase.getStartDate.Date, newphase.getEndDate.Date, oldrole.Xwerte, corrFactor, newXwerte)
                        Else
                            Call berechneBedarfe(newphase.getStartDate.Date, newphase.getEndDate.Date, oldrole.Xwerte, 1.0, newXwerte)
                        End If

                    Else
                        Call berechneBedarfe(newphase.getStartDate.Date, newphase.getEndDate.Date, oldrole.Xwerte, zielrenditeFaktor, newXwerte)
                    End If


                    With newrole
                        .uid = oldrole.uid
                        .teamID = oldrole.teamID
                        .Xwerte = newXwerte
                    End With
                    With newphase
                        .addRole(newrole)
                    End With
                Catch ex As Exception

                    Call MsgBox("Fehler in clsphase.korrcopyto")

                End Try

            Next r


            For k = 1 To Me.countCosts
                Try
                    oldcost = Me.getCost(k)
                    newcost = New clsKostenart(newphase.relEnde - newphase.relStart)

                    ReDim newXwerte(newphase.relEnde - newphase.relStart)
                    hname = oldcost.name

                    If zielrenditeFaktor = -99999.0 Then
                        ' undefiniert, deswegen corrfactor nehmen
                        If awinSettings.propAnpassRess Then
                            Call berechneBedarfe(newphase.getStartDate.Date, newphase.getEndDate.Date, oldcost.Xwerte, corrFactor, newXwerte)
                        Else
                            Call berechneBedarfe(newphase.getStartDate.Date, newphase.getEndDate.Date, oldcost.Xwerte, 1.0, newXwerte)
                        End If

                    Else
                        Call berechneBedarfe(newphase.getStartDate.Date, newphase.getEndDate.Date, oldcost.Xwerte, zielrenditeFaktor, newXwerte)
                    End If


                    With newcost
                        .KostenTyp = oldcost.KostenTyp
                        .Xwerte = newXwerte
                    End With
                    With newphase
                        .AddCost(newcost)
                    End With

                Catch ex As Exception

                    Call MsgBox("Fehler in clsphase.korrcopyto")

                End Try
            Next k


            ' Änderung 16.1.2014: zuerst die Rollen und Kosten übertragen, dann die relStart und RelEnde, dann die Results
            ' die evtl. enstehende Inkonsistenz zwischen Längen der Arrays der Rollen/Kostenarten und dem neuen relende/relstart wird in Kauf genommen 
            ' und nur korrigiert , wenn explizit gewünscht (Parameter awinsettings.autoCorrectBedarfe = true 

            ' alt .changeStartandDauer(Me._startOffsetinDays, Me._dauerInDays)

            For r = 1 To Me._allMilestones.Count
                newresult = New clsMeilenstein(parent:=newphase)
                If newPhaseNameID = "" Then
                    Me.getMilestone(r).copyTo(newresult)
                Else
                    Dim newMSNameID As String = newphase.parentProject.hierarchy.findUniqueElemKey(Me.getMilestone(r).name, True)
                    Me.getMilestone(r).copyTo(newresult, newMSNameID)
                End If

                ' korrigiert den Offset der Meilensteine 
                newresult.offset = CLng(System.Math.Round(CLng(Me.getMilestone(r).offset * corrFactor)))

                Try
                    .addMilestone(newresult)
                Catch ex As Exception

                End Try

            Next

            ' 16.12.19 Bewertungen auch übernehmen; in den Meilensteinen werden sie schon kängst übernommen ...


            For b As Integer = 1 To Me._bewertungen.Count
                Dim newb As New clsBewertung
                Me.getBewertung(b).copyto(newb)
                Try
                    .addBewertung(newb)
                Catch ex As Exception

                End Try

            Next

            ' Deliverables sollen immer übernommen werden ...
            ' jetzt noch die Deliverables kopieren ... 
            For i = 1 To Me.countDeliverables
                Dim deli As String = Me.getDeliverable(i)
                .addDeliverable(deli)
            Next


        End With

    End Sub

    ''' <summary>
    ''' passt die Offsets der Meilensteine an, wenn per Drag und Drop die entsprechende Phase 
    ''' gedehnt oder gestaucht wurde  
    ''' </summary>
    ''' <param name="faktor"></param>
    ''' <remarks></remarks>

    Public Sub adjustMilestones(ByVal faktor As Double)
        Dim newOffset As Integer
        For r = 1 To Me._allMilestones.Count

            ' korrigiert den Offset der Meilensteine 
            newOffset = CInt(System.Math.Round(CLng(Me.getMilestone(r).offset * faktor)))

            If newOffset < 0 Then
                newOffset = 0
            ElseIf newOffset > Me.dauerInDays Then
                newOffset = Me.dauerInDays - 1
            End If

            Me.getMilestone(r).offset = newOffset
        Next

    End Sub

    'Public Property Role(ByVal index As Integer) As clsRolle
    '    Get
    '        Role = _allRoles.Item(index - 1)
    '    End Get

    '    Set(value As clsRolle)
    '        _allRoles.Item(index - 1) = value
    '    End Set

    'End Property

    'Public Property Cost(ByVal index As Integer) As clsKostenart
    '    Get
    '        Cost = _allCosts.Item(index - 1)
    '    End Get

    '    Set(value As clsKostenart)
    '        _allCosts.Item(index - 1) = value
    '    End Set

    'End Property

    ''' <summary>
    ''' liefert die Rolle an Index-Stelle i; i darf Werte zwischen 1 und AnzahlRollen annehmen
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRole(ByVal index As Integer) As clsRolle

        Get
            If index > 0 And index <= _allRoles.Count Then
                getRole = _allRoles.Item(index - 1)
            Else
                getRole = Nothing
            End If

        End Get

    End Property

    ''' <summary>
    ''' liefert zu der angegebenen ID in Form von roleID;teamId die zugehörige Rolle, sofern sie in der Phase existiert
    ''' </summary>
    ''' <param name="roleNameID"></param>
    ''' <returns></returns>
    Public ReadOnly Property getRoleByRoleNameID(ByVal roleNameID As String) As clsRolle
        Get
            Dim tmpResult As clsRolle = Nothing
            Dim found As Boolean = False
            Dim ix As Integer = 0
            Dim teamID As Integer = -1
            Dim roleID As Integer = RoleDefinitions.parseRoleNameID(roleNameID, teamID)

            Do While Not found And ix <= _allRoles.Count - 1
                found = _allRoles.Item(ix).uid = roleID And _allRoles.Item(ix).teamID = teamID
                If found Then
                    tmpResult = _allRoles.Item(ix)
                Else
                    ix = ix + 1
                End If
            Loop

            getRoleByRoleNameID = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt eine Collection all der Meilensteine in der Phase zurück 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getMilestoneIDs() As Collection
        Get
            Dim tmpResult As New Collection
            For i As Integer = 1 To countMilestones
                Dim msID As String = getMilestone(i).nameID
                If Not tmpResult.Contains(msID) Then
                    tmpResult.Add(msID, msID)
                End If
            Next
            getMilestoneIDs = tmpResult
        End Get
    End Property


    ''' <summary>
    ''' gibt den ix-ten Meilenstein in der Phase zurück; ix muss zwischen 1 .. und count liegen
    ''' </summary>
    ''' <param name="index"></param>
    ''' <returns></returns>
    Public ReadOnly Property getMilestone(ByVal index As Integer) As clsMeilenstein

        Get
            If index < 1 Or index > _allMilestones.Count Then
                getMilestone = Nothing
            Else
                getMilestone = _allMilestones.Item(index - 1)
            End If

        End Get

    End Property

    ''' <summary>
    ''' gibt das Objekt Meilenstein mit der angegebenen ElemID zurück. 
    ''' beim Key kann es sich um eine ElemID handeln oder aber um einen Meilenstein-Namen, optional mit Nummer 
    ''' Wenn der Meilenstein nicht existiert, wird Nothing zurückgegeben 
    ''' </summary>
    ''' <param name="key">Name des Meilensteines</param>
    ''' <value></value>
    ''' <returns>Objekt vom Typ Result</returns>
    ''' <remarks>
    ''' Rückgabe von Nothing ist schneller als über Throw Exception zu arbeiten</remarks>
    Public ReadOnly Property getMilestone(ByVal key As String, Optional ByVal lfdNr As Integer = 1) As clsMeilenstein

        Get
            Dim tmpMilestone As clsMeilenstein = Nothing
            Dim found As Boolean = False
            Dim anzahl As Integer = 0
            Dim index As Integer
            Dim hryNode As clsHierarchyNode


            ' fedtlegen, worum es sich handelt: elemID oder Name

            If istElemID(key) Then

                hryNode = Me.parentProject.hierarchy.nodeItem(key)
                If Not IsNothing(hryNode) Then

                    ' prüfen, ob der Meilenstein überhaupt zu dieser Phase gehört 
                    If hryNode.parentNodeKey = Me.nameID Then
                        index = hryNode.indexOfElem
                        tmpMilestone = _allMilestones.Item(index - 1)
                    End If

                End If


            Else

                Dim r As Integer = 1
                While r <= Me.countMilestones And Not found

                    If elemNameOfElemID(_allMilestones.Item(r - 1).nameID) = key Then
                        anzahl = anzahl + 1
                        If anzahl >= lfdNr Then
                            found = True
                            tmpMilestone = _allMilestones.Item(r - 1)
                        End If
                    Else
                        r = r + 1
                    End If

                End While

            End If


            getMilestone = tmpMilestone


        End Get

    End Property

    ''' <summary>
    ''' gibt die laufende Nummer des Meilensteins in der Phase zurück
    ''' 0: wenn nicht gefunden
    ''' </summary>
    ''' <param name="msNameID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getlfdNr(ByVal msNameID As String) As Integer
        Get
            Dim r As Integer = 1
            Dim found As Boolean = False
            Dim tmpValue As Integer = 0

            While r <= Me.countMilestones And Not found
                If Me.getMilestone(r).nameID = msNameID Then
                    found = True
                    tmpValue = r
                Else
                    r = r + 1
                End If
            End While

            getlfdNr = tmpValue

        End Get
    End Property

    ''' <summary>
    ''' fügt die Kostenart Instanz der Liste der Kosten hinzu;
    ''' wenn sie schon existiert, werden die Xwerte aufsummiert  
    ''' </summary>
    ''' <param name="cost"></param>
    ''' <remarks></remarks>
    Public Sub AddCost(ByVal cost As clsKostenart)

        'sollte nach dem 8.7.16 aktiviert werden 
        'ebenso für addCost, mehrere Rollen/Kosten des gleichen NAmens sollen aufsummiert werden 
        Dim costName As String = cost.name

        Dim ix As Integer = 0
        Dim found As Boolean = False
        Dim oldXWerte() As Double
        Dim newXwerte() As Double

        While Not found And ix <= _allCosts.Count - 1
            If _allCosts.Item(ix).name = costName Then
                found = True
            Else
                ix = ix + 1
            End If
        End While

        If found Then
            oldXWerte = _allCosts.Item(ix).Xwerte()
            newXwerte = cost.Xwerte
            If oldXWerte.Length = newXwerte.Length Then
                ' hier dann aufsummieren 
                For i As Integer = 0 To oldXWerte.Length - 1
                    newXwerte(i) = newXwerte(i) + oldXWerte(i)
                Next

                _allCosts.Item(ix).Xwerte() = newXwerte

            Else
                ' darf eigentlich nicht sein 
                ' Test: 
                Call MsgBox("Fehler in Kosten-Zuordnung")
                ' es wird dann einfach gar nichts gemacht 
            End If
        Else
            _allCosts.Add(cost)
        End If

        ' jetzt müssen die sortierten Listen im Projekt entsprechend aktualisiert werden 
        Me.parentProject.rcLists.addCP(cost.KostenTyp, Me.nameID)


        ' vor dem 8.7.16
        ''If Not _allCosts.Contains(cost) Then
        ''    _allCosts.Add(cost)
        ''Else
        ''    Throw New Exception("Fehler: Kostenart '" & cost.name & "' ist bereits in der Phase '" & Me.name & "' enthalten")
        ''End If

    End Sub

    ''' <summary>
    ''' erstellt eine neue Kostenart, weist der Kostenart monatliche Bedarfe zu, deren Summe dem Wert der Variable summe entspricht  
    ''' </summary>
    ''' <param name="costName"></param>
    ''' <param name="summe"></param>
    ''' <param name="addToExisting"></param>
    Public Sub AddCost(ByVal costName As String, ByVal summe As Double, ByVal addToExisting As Boolean)

        Dim cSum As Double()
        ReDim cSum(0)
        cSum(0) = summe

        Dim tmpCost As clsKostenart = Me.getCost(costName)
        Dim xWerte As Double() = Me.berechneBedarfeNew(Me.getStartDate, Me.getEndDate, cSum, 1.0)

        If IsNothing(tmpCost) Then
            ' die Rolle hat bisher noch nicht existiert ...
            Dim dimension As Integer = Me.relEnde - Me.relStart
            tmpCost = New clsKostenart(dimension)

            With tmpCost
                .KostenTyp = CostDefinitions.getCostdef(costName).UID
                .Xwerte = xWerte
            End With

            ' jetzt muss die Rolle ergänzt werden 
            _allCosts.Add(tmpCost)

        Else
            ' die Rolle hat bereits existiert 
            If addToExisting Then
                If tmpCost.Xwerte.Length = xWerte.Length Then
                    ' hier dann aufsummieren 
                    Dim oldXwerte As Double() = tmpCost.Xwerte
                    For i As Integer = 0 To oldXwerte.Length - 1
                        xWerte(i) = xWerte(i) + oldXwerte(i)
                    Next

                Else
                    ' darf eigentlich nicht sein 
                    ' Test: 
                    'Call MsgBox("Fehler in Rollen-Zuordnung")
                    ' es wird dann einfach gar nichts gemacht 
                End If
            Else
                ' nichts weiter tun 
            End If

            tmpCost.Xwerte() = xWerte
        End If


        ' jetzt müssen die sortierten Listen im Projekt entsprechend aktualisiert werden 
        Try
            Me.parentProject.rcLists.addCP(tmpCost.KostenTyp, Me.nameID)
        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' entfernt alle Rollen-Instanzen mit Rollen-Name aus der Phase
    ''' </summary>
    ''' <param name="costName"></param>
    ''' <remarks></remarks>
    Public Sub removeCostByName(ByVal costName As String)

        Dim toDoList As New List(Of clsKostenart)

        For i As Integer = 1 To _allCosts.Count
            Dim tmpCost As clsKostenart = _allCosts.Item(i - 1)
            If tmpCost.name = costName Then
                toDoList.Add(tmpCost)
            End If
        Next

        For Each tmpCost As clsKostenart In toDoList
            _allCosts.Remove(tmpCost)
            ' jetzt müssen die sortierten Listen im Projekt entsprechend aktualisiert werden 
            Me.parentProject.rcLists.removeCP(tmpCost.KostenTyp, Me.nameID)
        Next


    End Sub



    ''' <summary>
    ''' gibt die Kostenart Instanz der Phase zurück, die den Namen costName hat 
    ''' </summary>
    ''' <param name="costName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCost(ByVal costName As String) As clsKostenart

        Get
            Dim returnValue As clsKostenart = Nothing
            Dim ix As Integer = 0
            Dim found As Boolean = False

            While Not found And ix <= _allCosts.Count - 1
                If _allCosts.Item(ix).name = costName Then
                    found = True
                    returnValue = _allCosts.Item(ix)
                Else
                    ix = ix + 1
                End If
            End While

            getCost = returnValue

        End Get

    End Property


    Public ReadOnly Property countCosts() As Integer

        Get
            countCosts = _allCosts.Count
        End Get

    End Property



    Public ReadOnly Property getCost(ByVal index As Integer) As clsKostenart

        Get
            If index > 0 And index <= _allCosts.Count Then
                getCost = _allCosts.Item(index - 1)
            Else
                getCost = Nothing
            End If

        End Get

    End Property

    ''' <summary>
    ''' gibt true zurück, wenn die Phase  actualValues enthält, d.h wenn die Phase cor oder im Monat startet, bis wohin actulaValues gehen 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property hasActualData As Boolean
        Get
            Dim tmpResult As Boolean = False
            If _parentProject.hasActualValues Then
                tmpResult = getStartDate < _parentProject.actualDataUntil
            End If
            hasActualData = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt treu zurück, wenn diese Phase noch Monate enthält , zu denen Forecast Planungen eingegeben werden können 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property hasForecastMonths As Boolean
        Get
            Dim tmpResult As Boolean = True
            If _parentProject.hasActualValues Then
                tmpResult = getColumnOfDate(getEndDate) > getColumnOfDate(_parentProject.actualDataUntil)
            End If
            hasForecastMonths = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt zur den Array an Ist-Werten der angegebenen Rolle / Kostenart zurück  
    ''' </summary>
    ''' <param name="rcNameID"></param>
    ''' <param name="isRole"></param>
    ''' <param name="outPutInEuro"></param>
    ''' <returns></returns>
    Public Function getActualRCValues(ByVal rcNameID As String, ByVal isRole As Boolean, ByVal outPutInEuro As Boolean) As Double()

        Dim tmpResult() As Double = Nothing

        Dim xWerte() As Double = Nothing
        Dim notYetDone As Boolean = True
        Dim tagessatz As Double = 800



        Dim pstart As Integer = getColumnOfDate(getStartDate)
        Dim pEnde As Integer = getColumnOfDate(getEndDate)
        Dim actualIX As Integer
        Dim arrayEnde As Integer

        If DateDiff(DateInterval.Month, StartofCalendar, parentProject.actualDataUntil) > 0 Then
            actualIX = getColumnOfDate(parentProject.actualDataUntil)
            arrayEnde = System.Math.Min(pEnde, actualIX)
        Else
            ' das ist das Abbruch-Kriterium, es gibt keine Ist-Daten
            arrayEnde = pstart - 1
        End If


        If pstart > arrayEnde Then
            ' es kann noch keine Ist-Daten geben 
            ReDim tmpResult(0)
            tmpResult(0) = 0

        ElseIf pstart <= arrayEnde Then
            ReDim tmpResult(arrayEnde - pstart)
            If isRole Then
                ' enthält diese Phase überhaupt diese Rolle ?
                Dim teamID As Integer = -1
                Dim roleID As Integer = RoleDefinitions.parseRoleNameID(rcNameID, teamID)

                Dim tmpRole As clsRolle = getRoleByRoleNameID(rcNameID)
                If Not IsNothing(tmpRole) Then
                    tagessatz = tmpRole.tagessatzIntern
                    xWerte = tmpRole.Xwerte
                Else
                    ReDim tmpResult(0)
                    tmpResult(0) = 0
                    notYetDone = False
                End If

            ElseIf rcNameID <> "" Then
                If CostDefinitions.containsName(rcNameID) Then
                    Dim costID As Integer = CostDefinitions.getCostdef(rcNameID).UID

                    Dim tmpCost As clsKostenart = getCost(rcNameID)
                    If Not IsNothing(tmpCost) Then
                        xWerte = tmpCost.Xwerte
                    Else
                        ReDim tmpResult(0)
                        tmpResult(0) = 0
                        notYetDone = False
                    End If

                Else
                    notYetDone = False
                End If


            Else
                notYetDone = False
            End If

            If notYetDone Then

                For i As Integer = 0 To arrayEnde - pstart
                    If isRole And outPutInEuro Then
                        ' mit Tagessatz multiplizieren 
                        tmpResult(i) = xWerte(i) * tagessatz
                    Else
                        tmpResult(i) = xWerte(i)
                    End If

                Next
            Else
                ReDim tmpResult(0)
                tmpResult(0) = 0
            End If

        End If



        getActualRCValues = tmpResult
    End Function

    ''' <summary>
    ''' liefert den Index zurück, bis zu dem ActualData in der Phase existiert 
    ''' -1 es existiert kein ActualData in der Phase 
    ''' 0 .LE. x .LE. dimension-1  die Monate xwerte(0), xwerte(1), ..xwerte(x) sind ActualData Monate  
    ''' relende-relstart .LE. x alles ist actual data    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getActualDataIndex As Integer
        Get
            Dim tmpResult As Integer = -1
            If hasActualData Then
                tmpResult = getColumnOfDate(_parentProject.actualDataUntil) - getColumnOfDate(getStartDate)
            End If

            getActualDataIndex = tmpResult
        End Get
    End Property

    Public ReadOnly Property parentProject() As clsProjekt
        Get
            parentProject = _parentProject
        End Get
    End Property

    Public ReadOnly Property VorlagenParent() As clsProjektvorlage
        Get
            VorlagenParent = _vorlagenParent
        End Get
    End Property

    Public Sub New(ByRef parent As clsProjekt)

        _nameID = ""
        _parentProject = parent
        _vorlagenParent = Nothing

        ' Vorbesetzen der Dokumenten-URL und App-ID , mit der die Dokumente bearbeitet werden können 
        _docURL = ""
        _docUrlAppID = ""

        _percentDone = 0.0
        _deliverables = New List(Of String)

        _bewertungen = New SortedList(Of String, clsBewertung)
        ' bei der Initialisierung wird nicht automatisch eine Bewertung angelegt ..
        ' tk 28.12.16 jede Phase bekommt eine leere Bewertung 
        'Dim tmpB As New clsBewertung
        'With tmpB
        '    .description = ""
        '    .colorIndex = 0
        'End With
        'Me.addBewertung(tmpB)

        _allRoles = New List(Of clsRolle)
        _allCosts = New List(Of clsKostenart)
        _allMilestones = New List(Of clsMeilenstein)

        _shortName = ""
        _originalName = ""
        _appearance = awinSettings.defaultPhaseClass

        Try
            _color = XlRgbColor.rgbDarkGrey
            ''If appearanceDefinitions.ContainsKey(_appearance) Then
            ''    If Not IsNothing(appearanceDefinitions.Item(_appearance).form) Then
            ''        _color = appearanceDefinitions.Item(_appearance).form.Fill.ForeColor.RGB
            ''    End If
            ''End If

        Catch ex As Exception

        End Try

        _verantwortlich = ""

        _offset = 0
        _earliestStart = -999
        _latestStart = -999

        _invoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
        _penalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, 0)



    End Sub

    Public Sub New(ByRef parent As clsProjektvorlage, ByVal isVorlage As Boolean)
        ' Variable isVorlage dient lediglich dazu, eine weitere Signatur für einen Konstruktor zu bekommen 
        ' dieser Konstruktor wird für parent = Vorlage benutzt 

        _nameID = ""
        _parentProject = Nothing
        _vorlagenParent = parent

        ' Vorbesetzen der Dokumenten-URL und App-ID , mit der die Dokumente bearbeitet werden können 
        _docURL = ""
        _docUrlAppID = ""

        _percentDone = 0.0
        _deliverables = New List(Of String)

        _bewertungen = New SortedList(Of String, clsBewertung)
        ' Änderung tk, bei der Initialisierung wird nicht automatisch eine Bewertung angelegt .. 
        ' tk 28.12.16 jede Phase bekommt eine leere Bewertung 
        'Dim tmpB As New clsBewertung
        'With tmpB
        '    .description = ""
        '    .colorIndex = 0
        'End With
        'Me.addBewertung(tmpB)

        _allRoles = New List(Of clsRolle)
        _allCosts = New List(Of clsKostenart)
        _allMilestones = New List(Of clsMeilenstein)

        _shortName = ""
        _originalName = ""
        _appearance = awinSettings.defaultPhaseClass

        Try
            _color = XlRgbColor.rgbDarkGrey
            ''If appearanceDefinitions.ContainsKey(_appearance) Then
            ''    If Not IsNothing(appearanceDefinitions.Item(_appearance).form) Then
            ''        _color = appearanceDefinitions.Item(_appearance).form.Fill.ForeColor.RGB
            ''    End If
            ''End If

        Catch ex As Exception

        End Try

        _verantwortlich = ""

        _offset = 0
        _earliestStart = -999
        _latestStart = -999

        _invoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
        _penalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, 0)



    End Sub

    ''' <summary>
    ''' synchronisiert bzw. berechnet die Xwerte der Rollen und Kosten
    ''' </summary>
    ''' <param name="calcAnyhow">wenn true: berechnet die Verteilung neu, auch wenn die Dimension des Arrays gleich bleibt</param>
    ''' <remarks></remarks>
    Public Sub calcNewXwerte(ByVal dimension As Integer, ByVal faktor As Double,
                             ByVal Optional calcAnyhow As Boolean = False)
        Dim newXwerte() As Double
        Dim oldXwerte() As Double
        Dim oldSum(0) As Double

        Dim r As Integer, k As Integer

        ' hier wird jetzt berücksichtigt, dass sich Werte aus den Ist-Daten nicht mehr verändern dürfen ..
        Dim actualIndex As Integer = getActualDataIndex


        If actualIndex < 0 Then
            ' alles wie bisher , ohne Istdaten
            For r = 1 To Me.countRoles
                oldXwerte = Me.getRole(r).Xwerte
                oldSum(0) = oldXwerte.Sum
                ReDim newXwerte(dimension)
                If calcAnyhow Then
                    Call berechneBedarfe(Me.getStartDate.Date, Me.getEndDate.Date, oldSum, faktor, newXwerte)
                Else
                    Call berechneBedarfe(Me.getStartDate.Date, Me.getEndDate.Date, oldXwerte, faktor, newXwerte)
                End If

                Me.getRole(r).Xwerte = newXwerte
            Next

            For k = 1 To Me.countCosts
                oldXwerte = Me.getCost(k).Xwerte
                oldSum(0) = oldXwerte.Sum
                ReDim newXwerte(dimension)
                If calcAnyhow Then
                    Call berechneBedarfe(Me.getStartDate.Date, Me.getEndDate.Date, oldSum, faktor, newXwerte)
                Else
                    Call berechneBedarfe(Me.getStartDate.Date, Me.getEndDate.Date, oldXwerte, faktor, newXwerte)
                End If

                Me.getCost(k).Xwerte = newXwerte
            Next

        Else
            ' jetzt müssen die Ist-Daten unverändert bleiben 
            Dim forecastDimension As Integer = dimension - actualIndex
            Dim firstForecastMonth As Date = getDateofColumn(getColumnOfDate(parentProject.actualDataUntil) + 1, False)

            For r = 1 To Me.countRoles

                oldXwerte = Me.getRole(r).Xwerte

                ReDim newXwerte(dimension)
                Dim newBedarf As Double = oldXwerte.Sum * faktor

                For ri As Integer = 0 To actualIndex
                    newXwerte(ri) = oldXwerte(ri)
                Next
                ' die bisher übertragenen Werte repräsentieren die Gesamt-summe an Ist-Daten 
                Dim istDataSum As Double = newXwerte.Sum

                Dim forecast(0) As Double
                forecast(0) = newBedarf - istDataSum

                ' darf nicht negativ werden  
                If forecast(0) < 0 Then
                    forecast(0) = 0
                End If


                Dim newForecastXWerte() As Double = calcVerteilungAufMonate(firstForecastMonth, Me.getEndDate, forecast, 1.0)


                ' jetzt die Forecast Werte übernehmen 
                For ri As Integer = actualIndex + 1 To dimension
                    newXwerte(ri) = newForecastXWerte(ri - (actualIndex + 1))
                Next

                Me.getRole(r).Xwerte = newXwerte

            Next

            For k = 1 To Me.countCosts
                oldXwerte = Me.getCost(k).Xwerte

                ReDim newXwerte(dimension)
                Dim newBedarf As Double = oldXwerte.Sum * faktor

                For ri As Integer = 0 To actualIndex
                    newXwerte(ri) = oldXwerte(ri)
                Next
                ' die bisher übertragenen Werte repräsentieren die Gesamt-summe an Ist-Daten 
                Dim istDataSum As Double = newXwerte.Sum

                Dim forecast(0) As Double
                forecast(0) = newBedarf - istDataSum

                ' darf nicht negativ werden  
                If forecast(0) < 0 Then
                    forecast(0) = 0
                End If

                Dim newForecastXWerte() As Double = calcVerteilungAufMonate(firstForecastMonth, Me.getEndDate, forecast, 1.0)

                ' jetzt die Forecast Werte übernehmen 
                For ri As Integer = actualIndex + 1 To dimension
                    newXwerte(ri) = newForecastXWerte(ri - (actualIndex + 1))
                Next

                Me.getCost(k).Xwerte = newXwerte

            Next
        End If



    End Sub


    ''' <summary>
    ''' berechnet die Bedarfe (Rollen,Kosten) der Phase gemäß Startdate und endedate, und corrFakt neu
    ''' neu: wird immer gemacht, nicht mehr in Abhängigkeit von propAnpassRess
    ''' </summary>
    ''' <param name="startdate"></param>
    ''' <param name="endedate"></param>
    ''' <param name="oldXwerte"></param>
    ''' <param name="corrFakt"></param>
    ''' <param name="newValues"></param>
    ''' <remarks></remarks>
    Public Sub berechneBedarfe(ByVal startdate As Date, ByVal endedate As Date, ByVal oldXwerte() As Double, _
                               ByVal corrFakt As Double, ByRef newValues() As Double)
        'Dim k As Integer
        'Dim newXwerte() As Double
        'Dim gesBedarf As Double
        'Dim Rest As Integer
        'Dim hDatum As Date
        'Dim anzDaysthisMonth As Double

        newValues = Me.berechneBedarfeNew(startdate, endedate, oldXwerte, corrFakt)

        ' Änderung tk 4.6.18 die berechneBedarfeNew verwendet 
        'Try
        '    ReDim newXwerte(newValues.Length - 1)

        '    If corrFakt <= 0 Then
        '        corrFakt = 1.0
        '    End If

        '    gesBedarf = oldXwerte.Sum
        '    gesBedarf = System.Math.Round(gesBedarf * corrFakt)


        '    If newValues.Length = oldXwerte.Length Then

        '        'Bedarfe-Verteilung bleibt wie gehabt, aber die corrfakt ist hier unberücksichtigt ..? 

        '        'If gesBedarf = oldXwerte.Sum Then
        '        If corrFakt = 1.0 Then
        '            newXwerte = oldXwerte
        '        Else
        '            For i = 0 To newValues.Length - 1
        '                newXwerte(i) = System.Math.Round(oldXwerte(i) * corrFakt)
        '            Next

        '            ' jetzt ggf die Reste verteilen 
        '            Rest = CInt(System.Math.Round(oldXwerte.Sum * corrFakt - newXwerte.Sum))

        '            k = newXwerte.Length - 1
        '            While Rest <> 0

        '                If Rest > 0 Then
        '                    newXwerte(k) = newXwerte(k) + 1
        '                    Rest = Rest - 1
        '                Else

        '                    If newXwerte(k) - 1 >= 0 Then
        '                        newXwerte(k) = newXwerte(k) - 1
        '                        Rest = Rest + 1
        '                    End If

        '                End If
        '                k = k - 1
        '                If k < 0 Then
        '                    k = newXwerte.Length - 1
        '                End If

        '            End While

        '        End If

        '    Else

        '        Dim tmpSum As Double = 0
        '        For k = 0 To newXwerte.Length - 1

        '            If k = 0 Then
        '                ' damit ist 00:00 des Startdates gemeint 
        '                hDatum = startdate

        '                anzDaysthisMonth = DateDiff(DateInterval.Day, hDatum, hDatum.AddDays(-1 * hDatum.Day + 1).AddMonths(1))

        '                'anzDaysthisMonth = DateDiff("d", hDatum, DateSerial(hDatum.Year, hDatum.Month + 1, hDatum.Day))
        '                'anzDaysthisMonth = anzDaysthisMonth - DateDiff("d", DateSerial(hDatum.Year, hDatum.Month, 1), hDatum) - 1

        '            ElseIf k = newXwerte.Length - 1 Then
        '                ' damit hDatum das End-Datum um 23.00 Uhr

        '                anzDaysthisMonth = endedate.Day
        '                'hDatum = endedate.AddHours(23)
        '                'anzDaysthisMonth = DateDiff("d", DateSerial(hDatum.Year, hDatum.Month, 1), hDatum)

        '            Else
        '                hDatum = startdate
        '                anzDaysthisMonth = DateDiff(DateInterval.Day, startdate.AddMonths(k), startdate.AddMonths(k + 1))
        '                'anzDaysthisMonth = DateDiff("d", DateSerial(hDatum.Year, hDatum.Month + k, hDatum.Day), DateSerial(hDatum.Year, hDatum.Month + k + 1, hDatum.Day))
        '            End If

        '            newXwerte(k) = System.Math.Round(anzDaysthisMonth / (Me.dauerInDays * corrFakt) * gesBedarf)
        '            tmpSum = tmpSum + anzDaysthisMonth
        '        Next k

        '        ' Kontrolle für Test ... aChck muss immer Null sein !
        '        'Dim aChck As Double = Me.dauerInDays - tmpSum


        '        ' Rest wird auf alle newXwerte verteilt

        '        Rest = CInt(gesBedarf - newXwerte.Sum)

        '        k = newXwerte.Length - 1
        '        While Rest <> 0
        '            If Rest > 0 Then
        '                newXwerte(k) = newXwerte(k) + 1
        '                Rest = Rest - 1
        '            Else
        '                If newXwerte(k) - 1 >= 0 Then
        '                    newXwerte(k) = newXwerte(k) - 1
        '                    Rest = Rest + 1
        '                End If
        '            End If
        '            k = k - 1
        '            If k < 0 Then
        '                k = newXwerte.Length - 1
        '            End If

        '        End While

        '    End If

        '    newValues = newXwerte

        'Catch ex As Exception

        '    Call MsgBox("Fehler in berechneBedarfe: " & vbLf & ex.Message)

        'End Try




    End Sub

    ''' <summary>
    ''' berechnet die Bedarfe (Rollen,Kosten) der Phase gemäß Startdate und endedate, und corrFakt neu
    ''' berücksichtigt die ActualDataUntil
    ''' ist jetzt als Function realisiert, die die Dimension aus Startdatum, Endedatum zieht 
    ''' wie die MEthode vorher ja auch ... 
    ''' </summary>
    ''' <param name="startdate"></param>
    ''' <param name="endedate"></param>
    ''' <param name="oldXwerte"></param>
    ''' <param name="corrFakt"></param>
    ''' <remarks></remarks>
    Public Function berechneBedarfeNew(ByVal startdate As Date, ByVal endedate As Date, ByVal oldXwerte() As Double, _
                               ByVal corrFakt As Double) As Double()

        ' wenn sich der bewährt: übernehmen ..
        'berechneBedarfeNew = calcVerteilungAufMonate(startdate, endedate, oldXwerte, corrFakt)

        ' tk 11.2.19
        ' alles folgende sollte, wenn sich Module1.calcVerteilungAufMonate(..) bewährt hat durch obigen Befehl ersetzt werden 
        Dim k As Integer
        Dim newXwerte() As Double
        Dim gesBedarf As Double
        Dim Rest As Integer
        Dim hDatum As Date
        Dim anzDaysthisMonth As Double
        Dim newLength As Integer = getColumnOfDate(endedate) - getColumnOfDate(startdate) + 1
        Dim gesBedarfReal As Double = 0.0

        ReDim newXwerte(newLength - 1)

        ' Vorbereitung für Summen Berechnung nur bei Forecast
        'Dim hasActualData As Boolean = Me.parentProject.actualDataUntil <> Date.MinValue
        'Dim actualDataColumn As Integer = -1

        'If hasActualData Then
        '    actualDataColumn = getColumnOfDate(Me.parentProject.actualDataUntil)
        'End If

        ' nur wenn überhaupt was zu verteilen ist, muss alles folgende gemacht werdne 
        ' andernfalls ist eh schon alles richtig 
        If oldXwerte.Sum > 0 Then

            Try

                gesBedarfReal = oldXwerte.Sum * corrFakt
                gesBedarf = System.Math.Round(gesBedarfReal)


                If newLength = oldXwerte.Length Then

                    'Bedarfe-Verteilung bleibt wie gehabt ... allerdings unter Berücksichtigung corrFakt


                    For i = 0 To newLength - 1
                        newXwerte(i) = oldXwerte(i) * corrFakt
                    Next

                    ' jetzt ggf die Reste verteilen 
                    Rest = CInt(gesBedarf - newXwerte.Sum)

                    k = newXwerte.Length - 1
                    While Rest <> 0

                        If Rest > 0 Then
                            newXwerte(k) = newXwerte(k) + 1
                            Rest = Rest - 1
                        Else

                            If newXwerte(k) - 1 >= 0 Then
                                newXwerte(k) = newXwerte(k) - 1
                                Rest = Rest + 1
                            End If

                        End If
                        k = k - 1
                        If k < 0 Then
                            k = newXwerte.Length - 1
                        End If

                    End While

                    ' letzter Test: wenn jetzt durch die Rundungen immer noch ein abs(Rest) von < 1 ist 
                    k = newXwerte.Length - 1
                    If newXwerte.Sum <> gesBedarfReal Then
                        Dim RestDbl As Double = gesBedarfReal - newXwerte.Sum
                        If Math.Abs(RestDbl) <= 1 And Math.Abs(RestDbl) >= 0 Then
                            ' alles ok 

                            ' positioniere auf ein k, dessen Wert größer ist als abs(restdbl) 
                            Do While newXwerte(k) < Math.Abs(RestDbl) And k > 0
                                k = k - 1
                            Loop
                            ' jetzt ist ein k erreicht 
                            newXwerte(k) = newXwerte(k) + RestDbl
                            If newXwerte(k) < 0 Then
                                newXwerte(k) = 0.0 ' darf eigentlich nie passieren ..
                            End If

                        Else
                            Dim a As Double = RestDbl ' kann / darf eigentlich nicht sein 
                        End If
                    End If


                Else

                    Dim tmpSum As Double = 0
                    For k = 0 To newXwerte.Length - 1

                        If k = 0 Then
                            ' damit ist 00:00 des Startdates gemeint 
                            hDatum = startdate

                            anzDaysthisMonth = DateDiff(DateInterval.Day, hDatum, hDatum.AddDays(-1 * hDatum.Day + 1).AddMonths(1))

                            'anzDaysthisMonth = DateDiff("d", hDatum, DateSerial(hDatum.Year, hDatum.Month + 1, hDatum.Day))
                            'anzDaysthisMonth = anzDaysthisMonth - DateDiff("d", DateSerial(hDatum.Year, hDatum.Month, 1), hDatum) - 1

                        ElseIf k = newXwerte.Length - 1 Then
                            ' damit hDatum das End-Datum um 23.00 Uhr

                            anzDaysthisMonth = endedate.Day
                            'hDatum = endedate.AddHours(23)
                            'anzDaysthisMonth = DateDiff("d", DateSerial(hDatum.Year, hDatum.Month, 1), hDatum)

                        Else
                            hDatum = startdate
                            anzDaysthisMonth = DateDiff(DateInterval.Day, startdate.AddMonths(k), startdate.AddMonths(k + 1))
                            'anzDaysthisMonth = DateDiff("d", DateSerial(hDatum.Year, hDatum.Month + k, hDatum.Day), DateSerial(hDatum.Year, hDatum.Month + k + 1, hDatum.Day))
                        End If

                        newXwerte(k) = System.Math.Round(anzDaysthisMonth / (Me.dauerInDays * corrFakt) * gesBedarf)
                        tmpSum = tmpSum + anzDaysthisMonth
                    Next k

                    ' Kontrolle für Test ... aChck muss immer Null sein !
                    'Dim aChck As Double = Me.dauerInDays - tmpSum


                    ' Rest wird auf alle newXwerte verteilt

                    Rest = CInt(gesBedarf - newXwerte.Sum)

                    k = newXwerte.Length - 1
                    While Rest <> 0
                        If Rest > 0 Then
                            newXwerte(k) = newXwerte(k) + 1
                            Rest = Rest - 1
                        Else
                            If newXwerte(k) - 1 >= 0 Then
                                newXwerte(k) = newXwerte(k) - 1
                                Rest = Rest + 1
                            End If
                        End If
                        k = k - 1
                        If k < 0 Then
                            k = newXwerte.Length - 1
                        End If

                    End While

                    ' letzter Test: wenn jetzt durch die Rundungen immer noch ein abs(Rest) von < 1 ist 
                    k = newXwerte.Length - 1
                    If newXwerte.Sum <> gesBedarfReal Then
                        Dim RestDbl As Double = gesBedarfReal - newXwerte.Sum
                        If Math.Abs(RestDbl) <= 1 And Math.Abs(RestDbl) >= 0 Then
                            ' alles ok 

                            ' positioniere auf ein k, dessen Wert größer ist als abs(restdbl) 
                            Do While newXwerte(k) < Math.Abs(RestDbl) And k > 0
                                k = k - 1
                            Loop
                            ' jetzt ist ein k erreicht 
                            newXwerte(k) = newXwerte(k) + RestDbl
                            If newXwerte(k) < 0 Then
                                newXwerte(k) = 0.0 ' darf eigentlich nie passieren ..
                            End If

                        Else
                            Dim a As Double = RestDbl ' kann / darf eigentlich nicht sein 
                        End If
                    End If

                End If



            Catch ex As Exception

            End Try

        Else
            ' alles auf Null setzen 
            For ix = 0 To newLength - 1
                newXwerte(ix) = 0
            Next
        End If

        berechneBedarfeNew = newXwerte

    End Function

End Class
